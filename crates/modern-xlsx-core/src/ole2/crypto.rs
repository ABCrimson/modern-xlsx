//! ECMA-376 password encryption (MS-OFFCRYPTO 2.3.4): key derivation, password
//! verification, decryption, integrity checking and Agile encryption.
//!
//! - Standard Encryption (2.3.4.5-2.3.4.9, read only): the SHA-1 key of
//!   `standard_key` (H_final, then X1/X2 with 0x36/0x5C), and AES-ECB with no
//!   IV and no padding for both the verifier and the EncryptedPackage.
//! - Agile Encryption (2.3.4.10-2.3.4.15): the password key encryptor's
//!   verifier input, verifier hash and `encryptedKeyValue` are AES-CBC under
//!   IV = its `saltValue` padded with 0x36 or truncated to 16 bytes,
//!   block-aligned with no PKCS#7: the reader decrypts without unpadding and
//!   truncates to `saltSize`, the digest length and the `keyData`
//!   `keyBits / 8`; the writer zero-fills the last block. The package is
//!   AES-CBC in 4096-byte segments. The `dataIntegrity` HMAC key and value are
//!   AES-CBC under the secret key with IV = H(`keyData` salt || block key),
//!   zero-filled the same way; the reader also accepts the scheme modern-xlsx
//!   1.2.x wrote (see `verify_hmac`).
//!
//! SHA-1, SHA-256, SHA-384 and SHA-512 ([`HashAlg`]) serve key derivation,
//! segment IVs, hashing and HMAC; AES-128, AES-192 and AES-256 work throughout.
//! Derived and unlocked keys are returned as [`Zeroizing`] buffers, wiped on
//! drop. [`decrypt_file`], [`encrypt_file`], the two password checks and the
//! two package decryptors are `pub`; the primitives are crate-internal.
//!
//! Each scheme lives in a private submodule whose public and crate-internal
//! entry points are re-exported here: `agile` (password check, segment cipher,
//! `dataIntegrity` HMAC and the Agile writer) and `standard` (AES-ECB,
//! password check and package decryption). This module keeps what both schemes
//! share: the constants, the key derivation (Standard's `standard_key`
//! included), the AES-CBC primitives, the CSPRNG, the hash and HMAC helpers,
//! the package size check, and the [`decrypt_file`] and [`encrypt_file`] entry
//! points, with `decrypt_file_with_cfb`, the form of `decrypt_file` the read
//! path calls on the compound file it has already parsed to classify it.

use core::hint::cold_path;

use aes::{Aes128Dec, Aes128Enc, Aes192Dec, Aes192Enc, Aes256Dec, Aes256Enc};
use cbc::cipher::{
    BlockCipherDecrypt, BlockCipherEncrypt, BlockModeDecrypt, BlockModeEncrypt, KeyIvInit,
    block_padding::{NoPadding, Padding, ZeroPadding},
};
use digest::{DynDigest, Output};
use hmac::{Hmac, KeyInit, Mac};
use sha1::Sha1;
use sha2::{Digest, Sha256, Sha384, Sha512, digest::FixedOutputReset};
use zeroize::{Zeroize, Zeroizing};

use super::cfb::reader::Cfb;
use super::encryption_info::{HashAlg, MAX_SPIN_COUNT};
use crate::Result;
use crate::errors::ModernXlsxError;

mod agile;
mod standard;

pub use agile::{decrypt_package, verify_password_agile};
pub(crate) use agile::{encrypt_file_with, verify_hmac};
pub use standard::{decrypt_standard_package, verify_password_standard};

/// Segment size for ECMA-376 Agile Encryption.
const SEGMENT_SIZE: usize = 4096;

/// AES block size in bytes: the only `blockSize` the Agile segment cipher accepts.
const AES_BLOCK: usize = 16;

/// Fixed SHA-1 iteration count of the Standard Encryption KDF (MS-OFFCRYPTO 2.3.4.7).
const STANDARD_SPIN_COUNT: u32 = 50_000;

/// `saltSize` the writer uses for `keyData` and the password key encryptor,
/// and the length of its random verifier input.
const WRITE_SALT_SIZE: usize = 16;

/// Block key constants per MS-OFFCRYPTO 2.3.4.13 (verifier input/value, encrypted key) and 2.3.4.14 (HMAC key/value).
pub(crate) const BLOCK_KEY_VERIFIER_INPUT: [u8; 8] =
    [0xFE, 0xA7, 0xD2, 0x76, 0x3B, 0x4B, 0x9E, 0x79];
pub(crate) const BLOCK_KEY_VERIFIER_VALUE: [u8; 8] =
    [0xD7, 0xAA, 0x0F, 0x6D, 0x30, 0x61, 0x34, 0x4E];
pub(crate) const BLOCK_KEY_ENCRYPTED_KEY: [u8; 8] =
    [0x14, 0x6E, 0x0B, 0xE7, 0xAB, 0xAC, 0xD0, 0xD6];
pub(crate) const BLOCK_KEY_HMAC_KEY: [u8; 8] = [0x5F, 0xB2, 0xAD, 0x01, 0x0C, 0xB9, 0xE1, 0xF6];
pub(crate) const BLOCK_KEY_HMAC_VALUE: [u8; 8] = [0xA0, 0x67, 0x7F, 0x02, 0xB2, 0x2C, 0x84, 0x33];

// ---------------------------------------------------------------------------
// Key derivation
// ---------------------------------------------------------------------------

/// Derives a key from a password per MS-OFFCRYPTO 2.3.4.11 (Encryption Key Generation, Agile Encryption).
///
/// One full pass: the iterated hash H_n ([`iterated_hash_for`]) followed by the
/// block-key finalization ([`finalize_key_for`]). The reader and the writer need
/// three keys from one password, salt, spin count and algorithm, so they compute
/// H_n once and finalize it per block key; this single-pass form exists for the
/// tests that pin the KDF against it.
///
/// # Arguments
/// * `password` - User password (encoded to UTF-16LE internally)
/// * `salt` - Salt from EncryptionInfo
/// * `spin_count` - Iteration count (typically 100,000; at most [`MAX_SPIN_COUNT`])
/// * `key_bits` - Desired key length in bits (128, 192 or 256)
/// * `block_key` - Block key constant (determines which key is derived)
/// * `hash_alg` - Hash algorithm, parsed from `hashAlgorithm` by [`EncryptionInfo::parse`](super::encryption_info::EncryptionInfo::parse)
///
/// # Returns
/// Derived key bytes of length `key_bits / 8`, zeroized on drop.
///
/// # Errors
/// The parameter checks of [`check_kdf_params`] run before any allocation or
/// hashing: `PasswordProtected` for a key size outside {128, 192, 256} and
/// `Security` for a spin count above [`MAX_SPIN_COUNT`].
#[cfg(test)]
pub(crate) fn derive_key(
    password: &str,
    salt: &[u8],
    spin_count: u32,
    key_bits: u32,
    block_key: &[u8],
    hash_alg: HashAlg,
) -> Result<Zeroizing<Vec<u8>>> {
    check_kdf_params(spin_count, key_bits)?;
    let h_n = iterated_hash_for(hash_alg, password, salt, spin_count);
    Ok(Zeroizing::new(finalize_key_for(
        hash_alg, &h_n, block_key, key_bits,
    )))
}

/// Rejects password-KDF parameters no supported file carries, mirroring the
/// [`EncryptionInfo::parse`](super::encryption_info::EncryptionInfo::parse)
/// checks for descriptors built by hand: a `key_bits` outside {128, 192, 256}
/// is `PasswordProtected("unsupported encryption: key size N bits")` and a
/// `spin_count` above [`MAX_SPIN_COUNT`] is `Security`. Without them a hostile
/// `keyBits` sizes the key allocation (up to 512 MiB) and a hostile
/// `spinCount` pins the thread for billions of hash rounds.
fn check_kdf_params(spin_count: u32, key_bits: u32) -> Result<()> {
    if !matches!(key_bits, 128 | 192 | 256) {
        cold_path();
        return Err(ModernXlsxError::PasswordProtected(format!(
            "unsupported encryption: key size {key_bits} bits"
        )));
    }
    if spin_count > MAX_SPIN_COUNT {
        cold_path();
        return Err(ModernXlsxError::Security(format!(
            "spinCount {spin_count} exceeds the MS-OFFCRYPTO maximum of {MAX_SPIN_COUNT}"
        )));
    }
    Ok(())
}

/// H_n of the Agile KDF (steps 1-3) for the parsed hash algorithm.
///
/// H_n depends only on the password, `salt`, `spin_count` and `alg`, so it may be
/// reused across block keys only while all four are the same: the password
/// key encryptor's salt, never `keyData`'s salt, and never across algorithms.
pub(crate) fn iterated_hash_for(
    alg: HashAlg,
    password: &str,
    salt: &[u8],
    spin_count: u32,
) -> Zeroizing<Vec<u8>> {
    match alg {
        HashAlg::Sha1 => iterated_hash::<Sha1>(password, salt, spin_count),
        HashAlg::Sha256 => iterated_hash::<Sha256>(password, salt, spin_count),
        HashAlg::Sha384 => iterated_hash::<Sha384>(password, salt, spin_count),
        HashAlg::Sha512 => iterated_hash::<Sha512>(password, salt, spin_count),
    }
}

/// Step 4 of the Agile KDF for the parsed hash algorithm: see [`finalize_key`].
pub(crate) fn finalize_key_for(
    alg: HashAlg,
    h_n: &[u8],
    block_key: &[u8],
    key_bits: u32,
) -> Vec<u8> {
    match alg {
        HashAlg::Sha1 => finalize_key::<Sha1>(h_n, block_key, key_bits),
        HashAlg::Sha256 => finalize_key::<Sha256>(h_n, block_key, key_bits),
        HashAlg::Sha384 => finalize_key::<Sha384>(h_n, block_key, key_bits),
        HashAlg::Sha512 => finalize_key::<Sha512>(h_n, block_key, key_bits),
    }
}

/// Steps 1-3 of the Agile KDF over any fixed-output digest:
/// H0 = H(salt || UTF-16LE(password)), then H_{i+1} = H(LE32(i) || H_i) for
/// `i` in `0..spin_count`. Returns H_n, where n = `spin_count`.
///
/// The UTF-16LE password and the chain buffer are zeroized. The password buffer
/// is sized up front (UTF-16 never needs more than two bytes per UTF-8 byte), so
/// it never reallocates and leaves no unwiped copy behind. The chain runs in one
/// `Output<D>` buffer refilled by `finalize_into_reset`, with no per-round
/// allocation.
fn iterated_hash<D: Digest + Default + FixedOutputReset>(
    password: &str,
    salt: &[u8],
    spin_count: u32,
) -> Zeroizing<Vec<u8>> {
    let mut pw_bytes = Zeroizing::new(Vec::with_capacity(password.len() * 2));
    pw_bytes.extend(password.encode_utf16().flat_map(u16::to_le_bytes));

    let mut hasher = D::default();
    let mut h = Output::<D>::default();
    Digest::update(&mut hasher, salt);
    Digest::update(&mut hasher, pw_bytes.as_slice());
    Digest::finalize_into_reset(&mut hasher, &mut h);

    for i in 0..spin_count {
        Digest::update(&mut hasher, i.to_le_bytes());
        Digest::update(&mut hasher, &h);
        Digest::finalize_into_reset(&mut hasher, &mut h);
    }

    let h_n = Zeroizing::new(h.to_vec());
    h.as_mut_slice().zeroize();
    h_n
}

/// Step 4 of the Agile KDF: H_final = H(H_n || block_key), truncated to
/// `key_bits / 8` bytes or padded to that length with 0x36. `key_bits` sizes
/// the allocation, so a value read from a file passes [`check_kdf_params`] first.
fn finalize_key<D: Digest>(h_n: &[u8], block_key: &[u8], key_bits: u32) -> Vec<u8> {
    let mut derived = D::new_with_prefix(h_n).chain_update(block_key).finalize();

    let key_len = (key_bits / 8) as usize;
    let mut key = vec![0x36u8; key_len];
    let copy_len = key_len.min(derived.len());
    key[..copy_len].copy_from_slice(&derived[..copy_len]);
    derived.as_mut_slice().zeroize();

    key
}

/// Derives the Standard Encryption key (MS-OFFCRYPTO 2.3.4.7, SHA-1 and AES).
///
/// 1. H_n: the SHA-1 chain of [`iterated_hash`] over `salt` and the UTF-16LE
///    password, [`STANDARD_SPIN_COUNT`] (50,000) rounds.
/// 2. H_final = SHA1(H_n || LE32(0)), block number 0.
/// 3. X1 = SHA1(64 bytes of 0x36 with H_final XORed into the first 20), and
///    X2 = the same with 0x5C.
/// 4. The key is the first `key_bits / 8` bytes of X1 || X2.
///
/// H_final, the two 64-byte pad blocks and X1 || X2 are zeroized; the key is
/// returned as a [`Zeroizing`] buffer.
///
/// # Errors
/// `PasswordProtected("unsupported encryption: key size N bits")` for a
/// `key_bits` outside {128, 192, 256}, checked before any hashing.
pub(crate) fn standard_key(
    password: &str,
    salt: &[u8],
    key_bits: u32,
) -> Result<Zeroizing<Vec<u8>>> {
    check_kdf_params(STANDARD_SPIN_COUNT, key_bits)?;
    let h_n = iterated_hash::<Sha1>(password, salt, STANDARD_SPIN_COUNT);
    let mut h_final = Sha1::new_with_prefix(&*h_n)
        .chain_update(0u32.to_le_bytes())
        .finalize();

    let mut x = Zeroizing::new([0u8; 40]);
    let (halves, _) = x.as_chunks_mut::<20>();
    for (half, pad) in halves.iter_mut().zip([0x36u8, 0x5C]) {
        let mut block = Zeroizing::new([pad; 64]);
        block.iter_mut().zip(&h_final).for_each(|(b, h)| *b ^= h);
        Digest::finalize_into(Sha1::new_with_prefix(block.as_slice()), half.into());
    }
    h_final.as_mut_slice().zeroize();

    Ok(Zeroizing::new(x[..(key_bits / 8) as usize].to_vec()))
}

// ---------------------------------------------------------------------------
// AES-CBC primitives
// ---------------------------------------------------------------------------

/// Decrypts AES-CBC data without removing any padding: the Agile verifier,
/// key and HMAC fields carry no PKCS#7, so the caller truncates the plaintext
/// to the field's declared length. The ciphertext must be a whole number of
/// AES blocks.
///
/// Selects AES-128, AES-192 or AES-256 from the key length (16, 24 or 32 bytes).
pub(crate) fn aes_cbc_decrypt_no_pad(key: &[u8], iv: &[u8], data: &[u8]) -> Result<Vec<u8>> {
    require_block_aligned(data, "AES-CBC")?;
    match key.len() {
        16 => cbc_decrypt::<Aes128Dec>(key, iv, data),
        24 => cbc_decrypt::<Aes192Dec>(key, iv, data),
        32 => cbc_decrypt::<Aes256Dec>(key, iv, data),
        n => Err(unsupported_aes_key(n)),
    }
}

/// CBC-decrypts block-aligned `data` under the block cipher `C`, no unpadding.
fn cbc_decrypt<C>(key: &[u8], iv: &[u8], data: &[u8]) -> Result<Vec<u8>>
where
    C: BlockCipherDecrypt + KeyInit,
{
    cbc::Decryptor::<C>::new_from_slices(key, iv)
        .map_err(|e| ModernXlsxError::PasswordProtected(format!("AES init error: {e}")))?
        .decrypt_padded_vec::<NoPadding>(data)
        .map_err(|_| ModernXlsxError::PasswordProtected("AES decryption failed".into()))
}

/// Rejects `data` that is not a whole number of AES blocks. The no-padding
/// cipher paths need aligned input, and the `cipher` crate's
/// `encrypt_padded_vec` panics on anything else under `NoPadding`.
fn require_block_aligned(data: &[u8], mode: &str) -> Result<()> {
    if data.len().is_multiple_of(AES_BLOCK) {
        Ok(())
    } else {
        cold_path();
        Err(ModernXlsxError::PasswordProtected(format!(
            "{mode} data length must be a multiple of {AES_BLOCK}"
        )))
    }
}

/// Error for a key that is not an AES-128/192/256 key.
#[cold]
fn unsupported_aes_key(len: usize) -> ModernXlsxError {
    ModernXlsxError::PasswordProtected(format!(
        "Unsupported AES key length: {len} bytes (expected 16, 24 or 32)"
    ))
}

// ---------------------------------------------------------------------------
// AES-CBC encryption primitives
// ---------------------------------------------------------------------------

/// Encrypts data using AES-CBC, zero-filling the last block (`ZeroPadding`):
/// block-aligned input gains no extra block. This is the padding MS-OFFCRYPTO
/// 2.3.4.13-2.3.4.14 gives the Agile verifier, key and HMAC fields (a 20-byte
/// SHA-1 hash becomes 32 bytes, a 64-byte SHA-512 hash stays 64).
///
/// Selects AES-128, AES-192 or AES-256 from the key length (16, 24 or 32 bytes).
pub(crate) fn aes_cbc_encrypt_zero_pad(key: &[u8], iv: &[u8], data: &[u8]) -> Result<Vec<u8>> {
    match key.len() {
        16 => cbc_encrypt::<Aes128Enc, ZeroPadding>(key, iv, data),
        24 => cbc_encrypt::<Aes192Enc, ZeroPadding>(key, iv, data),
        32 => cbc_encrypt::<Aes256Enc, ZeroPadding>(key, iv, data),
        n => Err(unsupported_aes_key(n)),
    }
}

/// Encrypts data without adding padding (for test use); unaligned `data` is an
/// error, never a cipher panic.
///
/// Selects AES-128, AES-192 or AES-256 from the key length (16, 24 or 32 bytes).
#[cfg(test)]
fn aes_cbc_encrypt_no_pad(key: &[u8], iv: &[u8], data: &[u8]) -> Result<Vec<u8>> {
    require_block_aligned(data, "AES-CBC")?;
    match key.len() {
        16 => cbc_encrypt::<Aes128Enc, NoPadding>(key, iv, data),
        24 => cbc_encrypt::<Aes192Enc, NoPadding>(key, iv, data),
        32 => cbc_encrypt::<Aes256Enc, NoPadding>(key, iv, data),
        n => Err(unsupported_aes_key(n)),
    }
}

/// CBC-encrypts `data` under the block cipher `C` with padding `P`. With
/// `NoPadding` the caller checks alignment first ([`require_block_aligned`]).
fn cbc_encrypt<C, P>(key: &[u8], iv: &[u8], data: &[u8]) -> Result<Vec<u8>>
where
    C: BlockCipherEncrypt + KeyInit,
    P: Padding,
{
    Ok(cbc::Encryptor::<C>::new_from_slices(key, iv)
        .map_err(|e| ModernXlsxError::PasswordProtected(format!("AES init error: {e}")))?
        .encrypt_padded_vec::<P>(data))
}

// ---------------------------------------------------------------------------
// Secure random
// ---------------------------------------------------------------------------

/// Generates cryptographically secure random bytes via `getrandom`.
///
/// Works in both native and WASM targets (via the `wasm_js` feature).
pub(crate) fn secure_random(len: usize) -> Result<Vec<u8>> {
    let mut buf = vec![0u8; len];
    getrandom::fill(&mut buf)
        .map_err(|e| ModernXlsxError::PasswordProtected(format!("CSPRNG error: {e}")))?;
    Ok(buf)
}

// ---------------------------------------------------------------------------
// Hashing helpers
// ---------------------------------------------------------------------------

/// Dispatch a hash operation on `alg`. The closure receives a fresh `Digest`
/// instance (SHA-1, SHA-256, SHA-384 or SHA-512) and returns the result.
fn with_hash_alg<F>(alg: HashAlg, f: F) -> Vec<u8>
where
    F: Fn(&mut dyn DynDigest) -> Vec<u8>,
{
    match alg {
        HashAlg::Sha1 => f(&mut Sha1::new()),
        HashAlg::Sha256 => f(&mut Sha256::new()),
        HashAlg::Sha384 => f(&mut Sha384::new()),
        HashAlg::Sha512 => f(&mut Sha512::new()),
    }
}

/// Hash bytes with `alg`.
fn hash_bytes(alg: HashAlg, data: &[u8]) -> Vec<u8> {
    with_hash_alg(alg, |h| {
        h.update(data);
        h.finalize_reset().into_vec()
    })
}

/// Hash the concatenation of two byte slices: H(prefix || data).
///
/// The digest keeps its own allocation (`into_vec`, no copy), so a caller that
/// wraps a key-derivation result in [`Zeroizing`] wipes the only buffer.
fn hash_bytes_with_prefix(prefix: &[u8], data: &[u8], alg: HashAlg) -> Vec<u8> {
    with_hash_alg(alg, |h| {
        h.update(prefix);
        h.update(data);
        h.finalize_reset().into_vec()
    })
}

/// Compute the HMAC of `data` under `key` with `alg` (RFC 2104).
fn compute_hmac(alg: HashAlg, key: &[u8], data: &[u8]) -> Result<Vec<u8>> {
    match alg {
        HashAlg::Sha1 => hmac_with::<Hmac<Sha1>>(key, data),
        HashAlg::Sha256 => hmac_with::<Hmac<Sha256>>(key, data),
        HashAlg::Sha384 => hmac_with::<Hmac<Sha384>>(key, data),
        HashAlg::Sha512 => hmac_with::<Hmac<Sha512>>(key, data),
    }
}

/// One-shot HMAC over `data` for a concrete `Hmac<D>`.
fn hmac_with<M: Mac + KeyInit>(key: &[u8], data: &[u8]) -> Result<Vec<u8>> {
    let mut mac = <M as KeyInit>::new_from_slice(key)
        .map_err(|e| ModernXlsxError::PasswordProtected(format!("HMAC init: {e}")))?;
    mac.update(data);
    Ok(mac.finalize().into_bytes().to_vec())
}

/// Checks `tag` against the full HMAC of `data` under `key` with `alg`, in
/// constant time ([`Mac::verify_slice`]). A tag of any other length than the
/// digest fails.
fn verify_hmac_tag(alg: HashAlg, key: &[u8], data: &[u8], tag: &[u8]) -> Result<()> {
    match alg {
        HashAlg::Sha1 => hmac_verify_with::<Hmac<Sha1>>(key, data, tag),
        HashAlg::Sha256 => hmac_verify_with::<Hmac<Sha256>>(key, data, tag),
        HashAlg::Sha384 => hmac_verify_with::<Hmac<Sha384>>(key, data, tag),
        HashAlg::Sha512 => hmac_verify_with::<Hmac<Sha512>>(key, data, tag),
    }
}

/// One-shot full-length HMAC check over `data` for a concrete `Hmac<D>`.
fn hmac_verify_with<M: Mac + KeyInit>(key: &[u8], data: &[u8], tag: &[u8]) -> Result<()> {
    let mut mac = <M as KeyInit>::new_from_slice(key)
        .map_err(|e| ModernXlsxError::PasswordProtected(format!("HMAC init: {e}")))?;
    mac.update(data);
    mac.verify_slice(tag).map_err(|_| hmac_mismatch())
}

/// The integrity failure: a wrong HMAC, or a `hashSize` that is not the digest length.
#[cold]
fn hmac_mismatch() -> ModernXlsxError {
    ModernXlsxError::PasswordProtected(
        "HMAC verification failed — file may be corrupted or tampered with.".into(),
    )
}

// ---------------------------------------------------------------------------
// EncryptedPackage size prefix (Agile and Standard)
// ---------------------------------------------------------------------------

/// The EncryptedPackage `StreamSize` prefix as a length no larger than the
/// `payload_len` bytes that follow it.
///
/// The prefix is attacker-controlled (the HMAC covers it, but whoever builds a
/// file also picks its password), so a size above the payload is hostile input:
/// `Security`, never an allocation or a silent short read. The size converts
/// with `usize::try_from`, never `as`, so a 32-bit target cannot truncate
/// `0x1_0000_0010` to 16 and pass it.
fn declared_package_size(size: [u8; 8], payload_len: usize) -> Result<usize> {
    let declared = u64::from_le_bytes(size);
    usize::try_from(declared)
        .ok()
        .filter(|&n| n <= payload_len)
        .ok_or_else(|| {
            cold_path();
            ModernXlsxError::Security(format!(
                "EncryptedPackage declares {declared} bytes but holds {payload_len}"
            ))
        })
}

// ---------------------------------------------------------------------------
// Full file encryption pipeline
// ---------------------------------------------------------------------------

/// Decrypts an OLE2-wrapped encrypted XLSX file.
/// Returns the decrypted ZIP bytes ready for normal reading.
///
/// The data key has one owner, the `Zeroizing` buffer the password check
/// returns, so it is wiped on drop, including on every early return.
pub fn decrypt_file(data: &[u8], password: &str) -> Result<Vec<u8>> {
    decrypt_file_with_cfb(Cfb::parse(data)?, password)
}

/// Decrypts an already parsed compound file, as [`decrypt_file`] does: the
/// `EncryptionInfo` read and parse, the password check, then the
/// `EncryptedPackage` read. It takes the `Cfb` by value and drops it once the
/// package is read, so its FAT and directory are freed before the package is
/// verified and decrypted.
pub(crate) fn decrypt_file_with_cfb(cfb: Cfb<'_>, password: &str) -> Result<Vec<u8>> {
    let enc_info_bytes = cfb.read_stream("EncryptionInfo")?;
    let enc_info = super::encryption_info::EncryptionInfo::parse(&enc_info_bytes)?;

    match enc_info {
        super::encryption_info::EncryptionInfo::Agile(ref agile) => {
            let data_key = verify_password_agile(password, agile)?;
            let encrypted_package = cfb.read_stream("EncryptedPackage")?;
            drop(cfb);
            verify_hmac(&data_key, agile, &encrypted_package)?;
            decrypt_package(&data_key, agile, &encrypted_package)
        }
        super::encryption_info::EncryptionInfo::Standard(ref std_info) => {
            let data_key = verify_password_standard(password, std_info)?;
            let encrypted_package = cfb.read_stream("EncryptedPackage")?;
            drop(cfb);
            decrypt_standard_package(&data_key, std_info, &encrypted_package)
        }
    }
}

/// Encrypts a ZIP archive with Agile Encryption at Excel's defaults: AES-256-CBC,
/// SHA-512 and spin count 100,000 (`encrypt_file_with` holds the scheme).
pub fn encrypt_file(zip_bytes: &[u8], password: &str) -> Result<Vec<u8>> {
    encrypt_file_with(zip_bytes, password, HashAlg::Sha512, 256, 100_000)
}

#[cfg(test)]
mod tests;
