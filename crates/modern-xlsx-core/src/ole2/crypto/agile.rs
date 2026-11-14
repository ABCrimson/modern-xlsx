//! Agile Encryption (MS-OFFCRYPTO 2.3.4.10-2.3.4.15): the password check, the
//! EncryptedPackage cipher in 4096-byte AES-CBC segments, the `dataIntegrity`
//! HMAC (the spec scheme, then the scheme modern-xlsx 1.2.x wrote) and the
//! writer behind [`encrypt_file`](super::encrypt_file).
//!
//! The key derivation, the AES-CBC primitives, the hash and HMAC helpers and
//! the package size check are shared with Standard Encryption and live in the
//! parent module.

use core::hint::cold_path;

use aes::{Aes128Dec, Aes128Enc, Aes192Dec, Aes192Enc, Aes256Dec, Aes256Enc};
use cbc::cipher::{
    BlockCipherDecrypt, BlockCipherEncrypt, BlockModeDecrypt, BlockModeEncrypt, InnerIvInit,
    KeyInit,
    block_padding::{NoPadding, ZeroPadding},
    consts::U16,
};
use sha1::Sha1;
use sha2::{Digest, Sha256, Sha384, Sha512};
use zeroize::Zeroizing;

use super::{
    AES_BLOCK, BLOCK_KEY_ENCRYPTED_KEY, BLOCK_KEY_HMAC_KEY, BLOCK_KEY_HMAC_VALUE,
    BLOCK_KEY_VERIFIER_INPUT, BLOCK_KEY_VERIFIER_VALUE, SEGMENT_SIZE, WRITE_SALT_SIZE,
    aes_cbc_decrypt_no_pad, aes_cbc_encrypt_zero_pad, check_kdf_params, compute_hmac,
    declared_package_size, finalize_key_for, hash_bytes, hash_bytes_with_prefix, hmac_mismatch,
    iterated_hash_for, secure_random, unsupported_aes_key, verify_hmac_tag,
};
use crate::Result;
use crate::errors::ModernXlsxError;
use crate::ole2::base64::encode_base64;
use crate::ole2::cfb;
use crate::ole2::encryption_info::{AgileEncryptionInfo, HashAlg};

// ---------------------------------------------------------------------------
// Password verification
// ---------------------------------------------------------------------------

/// The IV of the password encryptor's fields (MS-OFFCRYPTO 2.3.4.12): its
/// `saltValue` padded with 0x36 up to the 16-byte AES block, or truncated to it.
fn salt_iv(salt: &[u8]) -> [u8; AES_BLOCK] {
    let mut iv = [0x36u8; AES_BLOCK];
    let n = salt.len().min(AES_BLOCK);
    iv[..n].copy_from_slice(&salt[..n]);
    iv
}

/// Verifies a password against Agile encryption info (MS-OFFCRYPTO 2.3.4.13).
/// Returns the decrypted data encryption key (the secret key), zeroized on
/// drop, if the password is correct.
///
/// The spin-count chain H_n runs once; the three password keys (verifier input,
/// verifier value, encrypted key) are each one finalization of it. That reuse is
/// valid because all three share the password key encryptor's salt, spin count
/// and hash algorithm. Each field is AES-CBC-decrypted under its key (sized by
/// the encryptor's `keyBits`) with IV = the encryptor's `saltValue` padded with
/// 0x36 or truncated to 16 bytes (MS-OFFCRYPTO 2.3.4.12) and no unpadding, then
/// truncated: the verifier input to `saltSize`, the verifier hash to the digest
/// length and the encrypted key to the `keyData` `keyBits / 8`. Excel's
/// block-aligned fields and the PKCS#7 fields modern-xlsx 1.2.x wrote both read
/// this way.
///
/// # Errors
/// `PasswordProtected` for a wrong password, a `hashSize` that is not the digest
/// length, an unsupported key size or a field shorter than its declared length,
/// and `Security` for a spin count above `MAX_SPIN_COUNT` (10,000,000). The
/// encryptor `keyBits` and the spin count are checked before any hashing, also
/// for a descriptor built by hand; the `keyData` `keyBits` is checked after the
/// password check, before the secret key is unlocked.
pub fn verify_password_agile(
    password: &str,
    info: &AgileEncryptionInfo,
) -> Result<Zeroizing<Vec<u8>>> {
    check_kdf_params(info.pw_spin_count, info.pw_key_bits)?;
    let alg = info.pw_hash_alg;
    let salt = info.pw_salt.as_slice();
    let iv = salt_iv(salt);
    // H_n, every derived key and every decrypted field are Zeroizing buffers:
    // wiped on drop.
    let h_n = iterated_hash_for(alg, password, salt, info.pw_spin_count);
    let unlock = |block_key: &[u8], field: &[u8]| -> Result<Zeroizing<Vec<u8>>> {
        let key = Zeroizing::new(finalize_key_for(alg, &h_n, block_key, info.pw_key_bits));
        Ok(Zeroizing::new(aes_cbc_decrypt_no_pad(&key, &iv, field)?))
    };

    // 1. Hash the verifier input, truncated to saltSize.
    let input = unlock(
        &BLOCK_KEY_VERIFIER_INPUT,
        &info.pw_encrypted_verifier_hash_input,
    )?;
    let computed = hash_bytes(
        alg,
        checked_prefix(&input, salt.len(), "encryptedVerifierHashInput")?,
    );

    // 2. Compare it in constant time with the verifier hash, truncated to the
    //    digest length. A hashSize other than the digest length (parse already
    //    rejects it; this guards hand-built descriptors) fails like a wrong
    //    password.
    let hash_len = alg.digest_len();
    let value = unlock(
        &BLOCK_KEY_VERIFIER_VALUE,
        &info.pw_encrypted_verifier_hash_value,
    )?;
    let expected = checked_prefix(&value, hash_len, "encryptedVerifierHashValue")?;
    if usize::try_from(info.pw_hash_size).ok() != Some(hash_len)
        || !constant_time_eq::constant_time_eq(&computed, expected)
    {
        cold_path();
        return Err(ModernXlsxError::PasswordProtected(
            "Incorrect password.".into(),
        ));
    }

    // 3. Unlock the secret key, truncated to the keyData keyBits / 8 (the
    //    password encryptor's keyBits sizes only the key-encrypting keys).
    if !matches!(info.key_bits, 128 | 192 | 256) {
        cold_path();
        return Err(ModernXlsxError::PasswordProtected(format!(
            "unsupported encryption: key size {} bits",
            info.key_bits
        )));
    }
    let key_blob = unlock(&BLOCK_KEY_ENCRYPTED_KEY, &info.pw_encrypted_key_value)?;
    let key_len = (info.key_bits / 8) as usize;
    Ok(Zeroizing::new(
        checked_prefix(&key_blob, key_len, "encryptedKeyValue")?.to_vec(),
    ))
}

// ---------------------------------------------------------------------------
// Segment-based decryption
// ---------------------------------------------------------------------------

/// Decrypts the EncryptedPackage stream.
///
/// Layout: `[8 bytes: original size LE64] [encrypted segments of 4096 bytes each]`
/// Each segment IV = `Hash(keyData salt + LE32(segment_index))[..16]`; `keyData`
/// `blockSize` must be 16. The payload after the size prefix is copied once
/// (one allocation of its own length) and decrypted in place, then truncated
/// to the original size.
///
/// # Errors
/// `PasswordProtected` for a stream shorter than its 8-byte size prefix, a bad
/// `blockSize`, key or ciphertext, and `Security` when the prefix declares more
/// bytes than the stream holds.
pub fn decrypt_package(
    data_key: &[u8],
    info: &AgileEncryptionInfo,
    encrypted_package: &[u8],
) -> Result<Vec<u8>> {
    let Some((size, payload)) = encrypted_package.split_first_chunk::<8>() else {
        cold_path();
        return Err(ModernXlsxError::PasswordProtected(
            "EncryptedPackage too short".into(),
        ));
    };
    let original_size = declared_package_size(*size, payload.len())?;

    let alg = info.key_hash_alg;
    require_aes_block_size(info.key_block_size as usize)?;

    let mut out = payload.to_vec();
    let salt = &info.key_salt;
    match data_key.len() {
        16 => decrypt_segments::<Aes128Dec>(data_key, salt, alg, &mut out)?,
        24 => decrypt_segments::<Aes192Dec>(data_key, salt, alg, &mut out)?,
        32 => decrypt_segments::<Aes256Dec>(data_key, salt, alg, &mut out)?,
        n => return Err(unsupported_aes_key(n)),
    }

    out.truncate(original_size);
    Ok(out)
}

/// Borrow the first `n` bytes of `s`, or return a recoverable error if `s` is
/// shorter than `n`.
///
/// The lengths passed here (`keyBits / 8`, `saltSize`, the digest length) come
/// from the attacker-controlled EncryptionInfo XML, while the buffers they
/// slice are hash digests or decrypted fields whose lengths the same file
/// chooses. A plain `&s[..n]` would therefore panic on a malformed or hostile
/// file (e.g. `keyBits=2048` slicing a 64-byte SHA-512 digest, or a 16-byte
/// `encryptedKeyValue` for a 256-bit key). This keeps the reader's "never panic
/// on bad input" contract by turning that into an `Err`. For well-formed files
/// `n <= s.len()` always holds.
fn checked_prefix<'a>(s: &'a [u8], n: usize, field: &str) -> Result<&'a [u8]> {
    s.get(..n).ok_or_else(|| {
        cold_path();
        ModernXlsxError::PasswordProtected(format!(
            "invalid {field} in EncryptionInfo: need {n} bytes but have {}",
            s.len()
        ))
    })
}

/// Rejects a `blockSize` other than the AES block (16 bytes).
///
/// The value comes from the attacker-controlled EncryptionInfo XML; the segment
/// cipher is AES-CBC, whose IV is exactly one 16-byte block.
fn require_aes_block_size(block_size: usize) -> Result<()> {
    if block_size == AES_BLOCK {
        Ok(())
    } else {
        cold_path();
        Err(ModernXlsxError::PasswordProtected(format!(
            "unsupported encryption: blockSize {block_size} (expected {AES_BLOCK})"
        )))
    }
}

/// Derives the IV for a given segment index: `H(salt || LE32(segment_index))[..16]`.
///
/// Errors when `block_size` is not 16 (see [`require_aes_block_size`]).
pub(super) fn derive_segment_iv(
    salt: &[u8],
    segment_index: u32,
    block_size: usize,
    alg: HashAlg,
) -> Result<[u8; AES_BLOCK]> {
    derive_iv(salt, &segment_index.to_le_bytes(), block_size, alg)
}

/// Derives an Agile IV from the `keyData` salt: `H(salt || suffix)[..blockSize]`,
/// where the suffix is LE32 of a package segment index or the block key of a
/// `dataIntegrity` field (MS-OFFCRYPTO 2.3.4.12, 2.3.4.14). Every supported
/// digest is at least 20 bytes long, so the 16-byte prefix always exists.
///
/// Errors when `block_size` is not 16 (see [`require_aes_block_size`]).
fn derive_iv(
    salt: &[u8],
    suffix: &[u8],
    block_size: usize,
    alg: HashAlg,
) -> Result<[u8; AES_BLOCK]> {
    require_aes_block_size(block_size)?;
    Ok(match alg {
        HashAlg::Sha1 => iv_prefix::<Sha1>(salt, suffix),
        HashAlg::Sha256 => iv_prefix::<Sha256>(salt, suffix),
        HashAlg::Sha384 => iv_prefix::<Sha384>(salt, suffix),
        HashAlg::Sha512 => iv_prefix::<Sha512>(salt, suffix),
    })
}

/// `D(salt || suffix)[..16]`.
fn iv_prefix<D: Digest>(salt: &[u8], suffix: &[u8]) -> [u8; AES_BLOCK] {
    let hash = D::new_with_prefix(salt).chain_update(suffix).finalize();
    let mut iv = [0u8; AES_BLOCK];
    iv.copy_from_slice(&hash[..AES_BLOCK]);
    iv
}

/// Decrypts every 4096-byte segment of `buf` in place with AES-CBC and no
/// padding. The key schedule is built once and cloned into each segment's CBC
/// state; segment `i` uses the IV of [`derive_segment_iv`]. A segment that is not
/// a whole number of AES blocks is an error.
fn decrypt_segments<C>(key: &[u8], salt: &[u8], alg: HashAlg, buf: &mut [u8]) -> Result<()>
where
    C: BlockCipherDecrypt<BlockSize = U16> + KeyInit + Clone,
{
    let cipher = C::new_from_slice(key)
        .map_err(|e| ModernXlsxError::PasswordProtected(format!("AES init error: {e}")))?;
    for (segment, index) in buf.chunks_mut(SEGMENT_SIZE).zip(0u32..) {
        let iv = derive_segment_iv(salt, index, AES_BLOCK, alg)?;
        cbc::Decryptor::<C>::inner_iv_init(cipher.clone(), (&iv).into())
            .decrypt_padded::<NoPadding>(segment)
            .map_err(|_| ModernXlsxError::PasswordProtected("AES decryption failed".into()))?;
    }
    Ok(())
}

/// Appends the AES-CBC encryption of every 4096-byte segment of `data` to `out`.
/// Each segment is copied once, zero-filled to a whole AES block and encrypted in
/// place (`ZeroPadding` adds nothing to an aligned segment). The key schedule is
/// built once and cloned into each segment's CBC state.
fn encrypt_segments<C>(
    key: &[u8],
    salt: &[u8],
    alg: HashAlg,
    data: &[u8],
    out: &mut Vec<u8>,
) -> Result<()>
where
    C: BlockCipherEncrypt<BlockSize = U16> + KeyInit + Clone,
{
    let cipher = C::new_from_slice(key)
        .map_err(|e| ModernXlsxError::PasswordProtected(format!("AES init error: {e}")))?;
    for (segment, index) in data.chunks(SEGMENT_SIZE).zip(0u32..) {
        let iv = derive_segment_iv(salt, index, AES_BLOCK, alg)?;
        let start = out.len();
        out.extend_from_slice(segment);
        out.resize(start + segment.len().next_multiple_of(AES_BLOCK), 0);
        cbc::Encryptor::<C>::inner_iv_init(cipher.clone(), (&iv).into())
            .encrypt_padded::<ZeroPadding>(&mut out[start..], segment.len())
            .map_err(|_| ModernXlsxError::PasswordProtected("AES encryption failed".into()))?;
    }
    Ok(())
}

// ---------------------------------------------------------------------------
// HMAC integrity verification
// ---------------------------------------------------------------------------

/// Verifies the `dataIntegrity` HMAC of the encrypted package (MS-OFFCRYPTO
/// 2.3.4.14) under the secret key `data_key`.
///
/// The spec scheme comes first: `encryptedHmacKey` and `encryptedHmacValue` are
/// AES-CBC under `data_key` with IV = H(`keyData` salt || block key)[..16],
/// decrypted without unpadding and truncated to the digest length; the HMAC of
/// the whole EncryptedPackage stream (size prefix included) must equal the
/// value in full, compared in constant time ([`Mac::verify_slice`]).
///
/// If that fails, the scheme modern-xlsx 1.2.x wrote is tried, so its
/// password-protected files stay readable: the same fields under the key
/// H(`data_key` || block key)[..`keyBits` / 8] with IV = the `keyData` salt,
/// decrypted without unpadding (dropping its PKCS#7 block) and truncated the
/// same way. Both schemes need the secret key, so the fallback lets no one
/// without the password forge a package.
///
/// # Errors
/// `PasswordProtected` when neither scheme verifies, with the spec scheme's
/// error: an HMAC mismatch, a `hashSize` other than the digest length, a bad
/// `blockSize`, or a field that is unaligned or shorter than the digest.
///
/// [`Mac::verify_slice`]: hmac::Mac::verify_slice
pub(crate) fn verify_hmac(
    data_key: &[u8],
    info: &AgileEncryptionInfo,
    encrypted_package: &[u8],
) -> Result<()> {
    verify_hmac_spec(data_key, info, encrypted_package).or_else(|spec_err| {
        verify_hmac_legacy(data_key, info, encrypted_package).map_err(|_| spec_err)
    })
}

/// The MS-OFFCRYPTO 2.3.4.14 `dataIntegrity` check: see [`verify_hmac`].
pub(super) fn verify_hmac_spec(
    data_key: &[u8],
    info: &AgileEncryptionInfo,
    encrypted_package: &[u8],
) -> Result<()> {
    let alg = info.key_hash_alg;
    let block_size = info.key_block_size as usize;
    let key_iv = derive_iv(&info.key_salt, &BLOCK_KEY_HMAC_KEY, block_size, alg)?;
    let value_iv = derive_iv(&info.key_salt, &BLOCK_KEY_HMAC_VALUE, block_size, alg)?;
    let hmac_key = Zeroizing::new(aes_cbc_decrypt_no_pad(
        data_key,
        &key_iv,
        &info.encrypted_hmac_key,
    )?);
    let hmac_value = aes_cbc_decrypt_no_pad(data_key, &value_iv, &info.encrypted_hmac_value)?;
    check_package_hmac(info, &hmac_key, &hmac_value, encrypted_package)
}

/// The modern-xlsx 1.2.x `dataIntegrity` scheme, read-only: see [`verify_hmac`].
pub(super) fn verify_hmac_legacy(
    data_key: &[u8],
    info: &AgileEncryptionInfo,
    encrypted_package: &[u8],
) -> Result<()> {
    let alg = info.key_hash_alg;
    let key_len = (info.key_bits / 8) as usize;
    let key_key = Zeroizing::new(hash_bytes_with_prefix(data_key, &BLOCK_KEY_HMAC_KEY, alg));
    let value_key = Zeroizing::new(hash_bytes_with_prefix(data_key, &BLOCK_KEY_HMAC_VALUE, alg));
    let hmac_key = Zeroizing::new(aes_cbc_decrypt_no_pad(
        checked_prefix(&key_key, key_len, "keyBits")?,
        &info.key_salt,
        &info.encrypted_hmac_key,
    )?);
    let hmac_value = aes_cbc_decrypt_no_pad(
        checked_prefix(&value_key, key_len, "keyBits")?,
        &info.key_salt,
        &info.encrypted_hmac_value,
    )?;
    check_package_hmac(info, &hmac_key, &hmac_value, encrypted_package)
}

/// HMACs the whole EncryptedPackage stream under `hmac_key` truncated to the
/// digest length, and compares the result with `hmac_value` truncated the same
/// way. A `hashSize` other than the digest length fails as a mismatch.
fn check_package_hmac(
    info: &AgileEncryptionInfo,
    hmac_key: &[u8],
    hmac_value: &[u8],
    encrypted_package: &[u8],
) -> Result<()> {
    let alg = info.key_hash_alg;
    let hash_len = alg.digest_len();
    if usize::try_from(info.key_hash_size).ok() != Some(hash_len) {
        cold_path();
        return Err(hmac_mismatch());
    }
    verify_hmac_tag(
        alg,
        checked_prefix(hmac_key, hash_len, "encryptedHmacKey")?,
        encrypted_package,
        checked_prefix(hmac_value, hash_len, "encryptedHmacValue")?,
    )
}

// ---------------------------------------------------------------------------
// Segment-based encryption
// ---------------------------------------------------------------------------

/// Encrypts a ZIP package using ECMA-376 Agile segment-based encryption.
///
/// Layout: `[8 bytes: original size LE64] [encrypted segments of 4096 bytes]`
/// Each segment IV = `Hash(salt + LE32(segment_index))[..16]`; `block_size` must
/// be 16. The final segment is zero-filled to a whole AES block.
pub(crate) fn encrypt_package(
    data_key: &[u8],
    salt: &[u8],
    zip_bytes: &[u8],
    block_size: usize,
    alg: HashAlg,
) -> Result<Vec<u8>> {
    require_aes_block_size(block_size)?;

    // Every full segment is block-aligned, so the ciphertext is the input
    // rounded up to a whole number of AES blocks.
    let mut result = Vec::with_capacity(8 + zip_bytes.len().next_multiple_of(AES_BLOCK));
    result.extend_from_slice(&(zip_bytes.len() as u64).to_le_bytes());

    match data_key.len() {
        16 => encrypt_segments::<Aes128Enc>(data_key, salt, alg, zip_bytes, &mut result)?,
        24 => encrypt_segments::<Aes192Enc>(data_key, salt, alg, zip_bytes, &mut result)?,
        32 => encrypt_segments::<Aes256Enc>(data_key, salt, alg, zip_bytes, &mut result)?,
        n => return Err(unsupported_aes_key(n)),
    }

    Ok(result)
}

/// Builds the `dataIntegrity` fields for `encrypted_package` (MS-OFFCRYPTO
/// 2.3.4.14): a random HMAC key of the digest length, the HMAC of the whole
/// stream under it, and both encrypted with AES-CBC under the secret key
/// `data_key` with IV = the first `block_size` bytes of H(`key_salt` || block
/// key), the last block zero-filled (a SHA-1 field becomes 32 bytes; the
/// others stay at the digest length).
///
/// Returns `(encrypted_hmac_key, encrypted_hmac_value)`.
///
/// # Errors
/// `PasswordProtected` for a `block_size` other than 16 or a `data_key` that is
/// not an AES-128/192/256 key, and for a CSPRNG failure.
pub(crate) fn compute_and_encrypt_hmac(
    data_key: &[u8],
    key_salt: &[u8],
    encrypted_package: &[u8],
    block_size: usize,
    hash_alg: HashAlg,
) -> Result<(Vec<u8>, Vec<u8>)> {
    let key_iv = derive_iv(key_salt, &BLOCK_KEY_HMAC_KEY, block_size, hash_alg)?;
    let value_iv = derive_iv(key_salt, &BLOCK_KEY_HMAC_VALUE, block_size, hash_alg)?;

    // The HMAC key is a Zeroizing buffer: wiped on drop.
    let hmac_key = Zeroizing::new(secure_random(hash_alg.digest_len())?);
    let hmac_value = compute_hmac(hash_alg, &hmac_key, encrypted_package)?;

    Ok((
        aes_cbc_encrypt_zero_pad(data_key, &key_iv, &hmac_key)?,
        aes_cbc_encrypt_zero_pad(data_key, &value_iv, &hmac_value)?,
    ))
}

// ---------------------------------------------------------------------------
// Full file encryption
// ---------------------------------------------------------------------------

/// Encrypts a ZIP archive with Agile Encryption (MS-OFFCRYPTO 2.3.4.10-2.3.4.15)
/// under `hash`, an AES key of `key_bits` and `spin_count` KDF rounds.
///
/// Produces a complete OLE2 compound document containing:
/// - `EncryptionInfo` stream (version 4.4 + Agile XML descriptor), whose
///   `keyData` and password `encryptedKey` both declare `saltSize` 16,
///   `blockSize` 16, `keyBits`, `hashSize` (the digest length) and
///   `hashAlgorithm` from these parameters;
/// - `EncryptedPackage` stream (the ZIP in 4096-byte AES-CBC segments).
///
/// The verifier input (16 random bytes), its hash and the secret key are
/// AES-CBC-encrypted under the three password keys with IV = the password salt,
/// and the HMAC fields under the secret key with the `dataIntegrity` IVs
/// ([`compute_and_encrypt_hmac`]); every field's last block is zero-filled, so
/// each is its plaintext length rounded up to 16: the verifier input 16, the
/// hashes 32 (SHA-1, SHA-256), 48 or 64, the secret key 16 or 32 (AES-192's 24
/// bytes become 32).
///
/// # Errors
/// `PasswordProtected` for a `key_bits` outside {128, 192, 256} and `Security`
/// for a `spin_count` above [`MAX_SPIN_COUNT`], both before any work; otherwise
/// the errors of the CSPRNG and the OLE2 writer.
///
/// [`MAX_SPIN_COUNT`]: crate::ole2::encryption_info::MAX_SPIN_COUNT
pub(crate) fn encrypt_file_with(
    zip_bytes: &[u8],
    password: &str,
    hash: HashAlg,
    key_bits: u32,
    spin_count: u32,
) -> Result<Vec<u8>> {
    check_kdf_params(spin_count, key_bits)?;
    let hash_size = hash.digest_len();

    // 1. Random salts and secret key (Zeroizing wipes them on drop).
    let key_salt = secure_random(WRITE_SALT_SIZE)?;
    let pw_salt = secure_random(WRITE_SALT_SIZE)?;
    let data_key = Zeroizing::new(secure_random((key_bits / 8) as usize)?);

    // 2. The segment-encrypted package and its dataIntegrity fields.
    let encrypted_package = encrypt_package(&data_key, &key_salt, zip_bytes, AES_BLOCK, hash)?;
    let (encrypted_hmac_key, encrypted_hmac_value) =
        compute_and_encrypt_hmac(&data_key, &key_salt, &encrypted_package, AES_BLOCK, hash)?;

    // 3. The password verifier and the encrypted secret key. H_n runs once: the
    //    three password keys share pw_salt, spin_count and the hash algorithm.
    let verifier_input = secure_random(WRITE_SALT_SIZE)?;
    let verifier_hash = hash_bytes(hash, &verifier_input);
    let h_n = iterated_hash_for(hash, password, &pw_salt, spin_count);
    let seal = |block_key: &[u8], plain: &[u8]| -> Result<Vec<u8>> {
        let key = Zeroizing::new(finalize_key_for(hash, &h_n, block_key, key_bits));
        aes_cbc_encrypt_zero_pad(&key, &pw_salt, plain)
    };
    let encrypted_verifier_input = seal(&BLOCK_KEY_VERIFIER_INPUT, &verifier_input)?;
    let encrypted_verifier_value = seal(&BLOCK_KEY_VERIFIER_VALUE, &verifier_hash)?;
    let encrypted_key_value = seal(&BLOCK_KEY_ENCRYPTED_KEY, &data_key)?;

    // 4. Build the EncryptionInfo XML.
    let enc_info_xml = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<encryption xmlns="http://schemas.microsoft.com/office/2006/encryption" xmlns:p="http://schemas.microsoft.com/office/2006/keyEncryptor/password">
  <keyData saltSize="{WRITE_SALT_SIZE}" blockSize="{AES_BLOCK}" keyBits="{key_bits}" hashSize="{hash_size}" cipherAlgorithm="AES" cipherChaining="ChainingModeCBC" hashAlgorithm="{hash_name}" saltValue="{key_salt_b64}"/>
  <dataIntegrity encryptedHmacKey="{hmac_key_b64}" encryptedHmacValue="{hmac_value_b64}"/>
  <keyEncryptors>
    <keyEncryptor uri="http://schemas.microsoft.com/office/2006/keyEncryptor/password">
      <p:encryptedKey spinCount="{spin_count}" saltSize="{WRITE_SALT_SIZE}" blockSize="{AES_BLOCK}" keyBits="{key_bits}" hashSize="{hash_size}" cipherAlgorithm="AES" cipherChaining="ChainingModeCBC" hashAlgorithm="{hash_name}" saltValue="{pw_salt_b64}" encryptedKeyValue="{enc_key_b64}" encryptedVerifierHashInput="{enc_verifier_input_b64}" encryptedVerifierHashValue="{enc_verifier_value_b64}"/>
    </keyEncryptor>
  </keyEncryptors>
</encryption>"#,
        hash_name = hash.as_str(),
        key_salt_b64 = encode_base64(&key_salt),
        hmac_key_b64 = encode_base64(&encrypted_hmac_key),
        hmac_value_b64 = encode_base64(&encrypted_hmac_value),
        pw_salt_b64 = encode_base64(&pw_salt),
        enc_key_b64 = encode_base64(&encrypted_key_value),
        enc_verifier_input_b64 = encode_base64(&encrypted_verifier_input),
        enc_verifier_value_b64 = encode_base64(&encrypted_verifier_value),
    );

    // 5. Build EncryptionInfo stream: version 4.4 header + XML
    let xml_bytes = enc_info_xml.as_bytes();
    let mut enc_info_stream = Vec::with_capacity(8 + xml_bytes.len());
    enc_info_stream.extend_from_slice(&4u16.to_le_bytes()); // major = 4
    enc_info_stream.extend_from_slice(&4u16.to_le_bytes()); // minor = 4
    enc_info_stream.extend_from_slice(&0x40u32.to_le_bytes()); // flags (Agile)
    enc_info_stream.extend_from_slice(xml_bytes);

    // 6. Wrap in OLE2 compound document
    cfb::writer::write_ole2(&[
        ("EncryptionInfo", &enc_info_stream),
        ("EncryptedPackage", &encrypted_package),
    ])
}
