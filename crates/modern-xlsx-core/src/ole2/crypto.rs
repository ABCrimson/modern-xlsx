//! ECMA-376 Agile + Standard Encryption key derivation, AES decryption, and
//! AES encryption.
//!
//! Implements password-based key derivation, password verification, AES-CBC/ECB
//! decryption/encryption (segment-based for Agile, single-stream for Standard),
//! HMAC integrity verification, and full file encryption for OOXML encryption
//! (SHA-512/SHA-256/SHA-1, AES-128/AES-256).

use core::hint::cold_path;

use aes::{Aes128, Aes256};
use cbc::cipher::{BlockDecryptMut, BlockEncryptMut, KeyIvInit, block_padding::{NoPadding, Pkcs7}};
use hmac::{Hmac, Mac};
use sha1::Sha1;
use digest::DynDigest;
use sha2::{Digest, Sha256, Sha512, digest::FixedOutputReset};
use zeroize::{Zeroize, ZeroizeOnDrop};

use super::encryption_info::{AgileEncryptionInfo, StandardEncryptionInfo};
use crate::errors::ModernXlsxError;

type Result<T> = std::result::Result<T, ModernXlsxError>;

/// RAII wrapper that zeroizes key material on Drop, regardless of control flow.
///
/// Prevents key leaks from early returns (`?`), panics, or forgotten cleanup.
#[derive(Zeroize, ZeroizeOnDrop)]
pub(crate) struct SensitiveKey(pub(super) Vec<u8>);

impl SensitiveKey {
    pub(crate) fn new(data: Vec<u8>) -> Self {
        Self(data)
    }
}

impl std::ops::Deref for SensitiveKey {
    type Target = [u8];
    fn deref(&self) -> &[u8] {
        &self.0
    }
}

impl AsMut<[u8]> for SensitiveKey {
    fn as_mut(&mut self) -> &mut [u8] {
        &mut self.0
    }
}

type Aes128CbcDec = cbc::Decryptor<Aes128>;
type Aes256CbcDec = cbc::Decryptor<Aes256>;
type Aes128CbcEnc = cbc::Encryptor<Aes128>;
type Aes256CbcEnc = cbc::Encryptor<Aes256>;

/// Segment size for ECMA-376 Agile Encryption.
const SEGMENT_SIZE: usize = 4096;

/// Default iteration count for Standard Encryption (MS-OFFCRYPTO 2.3.6.2).
const STANDARD_SPIN_COUNT: u32 = 50_000;

/// Block key for Standard Encryption key derivation (MS-OFFCRYPTO 2.3.6.2).
const STANDARD_BLOCK_KEY: [u8; 8] = [0x14, 0x6E, 0x0B, 0xE7, 0xAB, 0xAC, 0xD0, 0xD6];

/// Block key constants per ECMA-376 SS2.3.6.2.
pub const BLOCK_KEY_VERIFIER_INPUT: [u8; 8] = [0xFE, 0xA7, 0xD2, 0x76, 0x3B, 0x4B, 0x9E, 0x79];
pub const BLOCK_KEY_VERIFIER_VALUE: [u8; 8] = [0xD7, 0xAA, 0x0F, 0x6D, 0x30, 0x61, 0x34, 0x4E];
pub const BLOCK_KEY_ENCRYPTED_KEY: [u8; 8] = [0x14, 0x6E, 0x0B, 0xE7, 0xAB, 0xAC, 0xD0, 0xD6];
pub const BLOCK_KEY_HMAC_KEY: [u8; 8] = [0x5F, 0xB2, 0xAD, 0x01, 0x0C, 0xB9, 0xE1, 0xF6];
pub const BLOCK_KEY_HMAC_VALUE: [u8; 8] = [0xA0, 0x67, 0x7F, 0x02, 0xB2, 0x2C, 0x84, 0x33];

// ---------------------------------------------------------------------------
// Key derivation
// ---------------------------------------------------------------------------

/// Derives a key from a password per ECMA-376 Agile Encryption SS2.3.6.2.
///
/// # Arguments
/// * `password` - User password (encoded to UTF-16LE internally)
/// * `salt` - Salt from EncryptionInfo
/// * `spin_count` - Iteration count (typically 100,000)
/// * `key_bits` - Desired key length in bits (128 or 256)
/// * `block_key` - Block key constant (determines which key is derived)
/// * `hash_alg` - Hash algorithm name ("SHA512" or "SHA256")
///
/// # Returns
/// Derived key bytes of length `key_bits / 8`.
pub fn derive_key(
    password: &str,
    salt: &[u8],
    spin_count: u32,
    key_bits: u32,
    block_key: &[u8],
    hash_alg: &str,
) -> Result<Vec<u8>> {
    match hash_alg {
        "SHA512" => derive_key_impl::<Sha512>(password, salt, spin_count, key_bits, block_key),
        "SHA256" => derive_key_impl::<Sha256>(password, salt, spin_count, key_bits, block_key),
        "SHA1" => derive_key_impl::<Sha1>(password, salt, spin_count, key_bits, block_key),
        _ => {
            cold_path();
            Err(ModernXlsxError::PasswordProtected(format!(
                "Unsupported hash algorithm: {hash_alg}"
            )))
        }
    }
}

/// Generic key derivation over any fixed-output hash digest (SHA-1, SHA-256, SHA-512).
fn derive_key_impl<D: Digest + Default + FixedOutputReset>(
    password: &str,
    salt: &[u8],
    spin_count: u32,
    key_bits: u32,
    block_key: &[u8],
) -> Result<Vec<u8>> {
    // Step 1: Encode password as UTF-16LE
    let pw_bytes: Vec<u8> = password
        .encode_utf16()
        .flat_map(|c| c.to_le_bytes())
        .collect();

    // Step 2: H0 = Hash(salt + password)
    let mut hasher = D::default();
    Digest::update(&mut hasher, salt);
    Digest::update(&mut hasher, &pw_bytes);
    let mut hash = hasher.finalize_reset();

    // Step 3: Iterate: H_i = Hash(LE32(i) + H_{i-1})
    for i in 0..spin_count {
        Digest::update(&mut hasher, i.to_le_bytes());
        Digest::update(&mut hasher, &hash);
        hash = hasher.finalize_reset();
    }

    // Step 4: H_final = Hash(H_n + blockKey)
    Digest::update(&mut hasher, &hash);
    Digest::update(&mut hasher, block_key);
    let derived = hasher.finalize();

    // Step 5: Truncate or pad to key_bits/8 (pad with 0x36 per spec)
    let key_len = (key_bits / 8) as usize;
    let mut key = vec![0x36u8; key_len];
    let copy_len = key_len.min(derived.len());
    key[..copy_len].copy_from_slice(&derived[..copy_len]);

    Ok(key)
}

// ---------------------------------------------------------------------------
// AES-CBC primitives
// ---------------------------------------------------------------------------

/// Decrypts data using AES-CBC with PKCS#7 padding removal.
///
/// Automatically selects AES-128 or AES-256 based on key length.
pub fn aes_cbc_decrypt(key: &[u8], iv: &[u8], data: &[u8]) -> Result<Vec<u8>> {
    let mut buf = data.to_vec();
    match key.len() {
        16 => {
            let decryptor = Aes128CbcDec::new_from_slices(key, iv)
                .map_err(|e| ModernXlsxError::PasswordProtected(format!("AES init error: {e}")))?;
            let plaintext = decryptor
                .decrypt_padded_mut::<Pkcs7>(&mut buf)
                .map_err(|_| {
                    ModernXlsxError::PasswordProtected(
                        "AES decryption failed (bad padding)".into(),
                    )
                })?;
            Ok(plaintext.to_vec())
        }
        32 => {
            let decryptor = Aes256CbcDec::new_from_slices(key, iv)
                .map_err(|e| ModernXlsxError::PasswordProtected(format!("AES init error: {e}")))?;
            let plaintext = decryptor
                .decrypt_padded_mut::<Pkcs7>(&mut buf)
                .map_err(|_| {
                    ModernXlsxError::PasswordProtected(
                        "AES decryption failed (bad padding)".into(),
                    )
                })?;
            Ok(plaintext.to_vec())
        }
        n => {
            cold_path();
            Err(ModernXlsxError::PasswordProtected(format!(
                "Unsupported AES key length: {n} bytes (expected 16 or 32)"
            )))
        }
    }
}

/// Decrypts data without removing padding (for raw segment blocks).
///
/// Automatically selects AES-128 or AES-256 based on key length.
pub fn aes_cbc_decrypt_no_pad(key: &[u8], iv: &[u8], data: &[u8]) -> Result<Vec<u8>> {
    let mut buf = data.to_vec();
    match key.len() {
        16 => {
            let decryptor = Aes128CbcDec::new_from_slices(key, iv)
                .map_err(|e| ModernXlsxError::PasswordProtected(format!("AES init error: {e}")))?;
            decryptor
                .decrypt_padded_mut::<NoPadding>(&mut buf)
                .map_err(|_| {
                    ModernXlsxError::PasswordProtected("AES decryption failed".into())
                })?;
            Ok(buf)
        }
        32 => {
            let decryptor = Aes256CbcDec::new_from_slices(key, iv)
                .map_err(|e| ModernXlsxError::PasswordProtected(format!("AES init error: {e}")))?;
            decryptor
                .decrypt_padded_mut::<NoPadding>(&mut buf)
                .map_err(|_| {
                    ModernXlsxError::PasswordProtected("AES decryption failed".into())
                })?;
            Ok(buf)
        }
        n => {
            cold_path();
            Err(ModernXlsxError::PasswordProtected(format!(
                "Unsupported AES key length: {n} bytes (expected 16 or 32)"
            )))
        }
    }
}

// ---------------------------------------------------------------------------
// AES-CBC encryption primitives
// ---------------------------------------------------------------------------

/// Encrypts data using AES-CBC with PKCS#7 padding.
///
/// Automatically selects AES-128 or AES-256 based on key length.
pub fn aes_cbc_encrypt(key: &[u8], iv: &[u8], data: &[u8]) -> Result<Vec<u8>> {
    match key.len() {
        16 => {
            let encryptor = Aes128CbcEnc::new_from_slices(key, iv)
                .map_err(|e| ModernXlsxError::PasswordProtected(format!("AES init error: {e}")))?;
            Ok(encryptor.encrypt_padded_vec_mut::<Pkcs7>(data))
        }
        32 => {
            let encryptor = Aes256CbcEnc::new_from_slices(key, iv)
                .map_err(|e| ModernXlsxError::PasswordProtected(format!("AES init error: {e}")))?;
            Ok(encryptor.encrypt_padded_vec_mut::<Pkcs7>(data))
        }
        n => {
            cold_path();
            Err(ModernXlsxError::PasswordProtected(format!(
                "Unsupported AES key length: {n} bytes (expected 16 or 32)"
            )))
        }
    }
}

/// Encrypts data without adding padding (data must be block-aligned).
///
/// Automatically selects AES-128 or AES-256 based on key length.
pub fn aes_cbc_encrypt_no_pad(key: &[u8], iv: &[u8], data: &[u8]) -> Result<Vec<u8>> {
    match key.len() {
        16 => {
            let encryptor = Aes128CbcEnc::new_from_slices(key, iv)
                .map_err(|e| ModernXlsxError::PasswordProtected(format!("AES init error: {e}")))?;
            Ok(encryptor.encrypt_padded_vec_mut::<NoPadding>(data))
        }
        32 => {
            let encryptor = Aes256CbcEnc::new_from_slices(key, iv)
                .map_err(|e| ModernXlsxError::PasswordProtected(format!("AES init error: {e}")))?;
            Ok(encryptor.encrypt_padded_vec_mut::<NoPadding>(data))
        }
        n => {
            cold_path();
            Err(ModernXlsxError::PasswordProtected(format!(
                "Unsupported AES key length: {n} bytes (expected 16 or 32)"
            )))
        }
    }
}

// ---------------------------------------------------------------------------
// Secure random
// ---------------------------------------------------------------------------

/// Generates cryptographically secure random bytes via `getrandom`.
///
/// Works in both native and WASM targets (via the `wasm_js` feature).
pub fn secure_random(len: usize) -> Result<Vec<u8>> {
    let mut buf = vec![0u8; len];
    getrandom::fill(&mut buf)
        .map_err(|e| ModernXlsxError::PasswordProtected(format!("CSPRNG error: {e}")))?;
    Ok(buf)
}

// ---------------------------------------------------------------------------
// Hashing helpers
// ---------------------------------------------------------------------------

/// Dispatch a hash operation by algorithm name. The closure receives a fresh
/// `Digest` instance (SHA-512, SHA-256, or SHA-1) and must return the result.
fn with_hash_alg<F>(alg: &str, f: F) -> Result<Vec<u8>>
where
    F: Fn(&mut dyn DynDigest) -> Vec<u8>,
{
    match alg {
        "SHA512" => Ok(f(&mut Sha512::new())),
        "SHA256" => Ok(f(&mut Sha256::new())),
        "SHA1" => Ok(f(&mut Sha1::new())),
        _ => {
            cold_path();
            Err(ModernXlsxError::PasswordProtected(format!(
                "Unsupported hash algorithm: {alg}"
            )))
        }
    }
}

/// Hash bytes with the named algorithm.
fn hash_bytes(alg: &str, data: &[u8]) -> Result<Vec<u8>> {
    with_hash_alg(alg, |h| {
        h.update(data);
        h.finalize_reset().to_vec()
    })
}

/// Hash the concatenation of two byte slices: H(prefix || data).
fn hash_bytes_with_prefix(prefix: &[u8], data: &[u8], alg: &str) -> Result<Vec<u8>> {
    with_hash_alg(alg, |h| {
        h.update(prefix);
        h.update(data);
        h.finalize_reset().to_vec()
    })
}

/// Compute HMAC of `data` using the named hash algorithm and `key`.
fn compute_hmac(alg: &str, key: &[u8], data: &[u8]) -> Result<Vec<u8>> {
    match alg {
        "SHA512" => {
            let mut mac = Hmac::<Sha512>::new_from_slice(key)
                .map_err(|e| ModernXlsxError::PasswordProtected(format!("HMAC init: {e}")))?;
            mac.update(data);
            Ok(mac.finalize().into_bytes().to_vec())
        }
        "SHA256" => {
            let mut mac = Hmac::<Sha256>::new_from_slice(key)
                .map_err(|e| ModernXlsxError::PasswordProtected(format!("HMAC init: {e}")))?;
            mac.update(data);
            Ok(mac.finalize().into_bytes().to_vec())
        }
        _ => {
            cold_path();
            Err(ModernXlsxError::PasswordProtected(format!(
                "Unsupported hash algorithm for HMAC: {alg}"
            )))
        }
    }
}

// ---------------------------------------------------------------------------
// Password verification
// ---------------------------------------------------------------------------

/// Verifies a password against Agile encryption info.
/// Returns the decrypted data encryption key if password is correct.
pub fn verify_password_agile(
    password: &str,
    info: &AgileEncryptionInfo,
) -> Result<Vec<u8>> {
    // 1. Derive verifier hash input key (SensitiveKey auto-zeroizes on Drop)
    let input_key = SensitiveKey::new(derive_key(
        password,
        &info.pw_salt,
        info.pw_spin_count,
        info.pw_key_bits,
        &BLOCK_KEY_VERIFIER_INPUT,
        &info.pw_hash_alg,
    )?);

    // 2. Decrypt verifier hash input
    let verifier_input = aes_cbc_decrypt(
        &input_key,
        &info.pw_salt,
        &info.pw_encrypted_verifier_hash_input,
    )?;

    // 3. Derive verifier hash value key
    let value_key = SensitiveKey::new(derive_key(
        password,
        &info.pw_salt,
        info.pw_spin_count,
        info.pw_key_bits,
        &BLOCK_KEY_VERIFIER_VALUE,
        &info.pw_hash_alg,
    )?);

    // 4. Decrypt verifier hash value
    let verifier_hash = aes_cbc_decrypt(
        &value_key,
        &info.pw_salt,
        &info.pw_encrypted_verifier_hash_value,
    )?;

    // 5. Hash the decrypted verifier input
    let computed = hash_bytes(&info.pw_hash_alg, &verifier_input)?;

    // 6. Constant-time comparison (truncate to hash_size)
    let hash_len = info.pw_hash_size as usize;
    if computed.len() < hash_len || verifier_hash.len() < hash_len {
        cold_path();
        return Err(ModernXlsxError::PasswordProtected(
            "Incorrect password.".into(),
        ));
    }
    if !constant_time_eq::constant_time_eq(&computed[..hash_len], &verifier_hash[..hash_len]) {
        cold_path();
        return Err(ModernXlsxError::PasswordProtected(
            "Incorrect password.".into(),
        ));
    }

    // 7. Derive data encryption key
    let enc_key_key = SensitiveKey::new(derive_key(
        password,
        &info.pw_salt,
        info.pw_spin_count,
        info.pw_key_bits,
        &BLOCK_KEY_ENCRYPTED_KEY,
        &info.pw_hash_alg,
    )?);

    // 8. Decrypt the actual data encryption key
    let data_key = aes_cbc_decrypt(&enc_key_key, &info.pw_salt, &info.pw_encrypted_key_value)?;

    let key_len = (info.pw_key_bits / 8) as usize;
    Ok(data_key[..key_len.min(data_key.len())].to_vec())
}

// ---------------------------------------------------------------------------
// Segment-based decryption
// ---------------------------------------------------------------------------

/// Decrypts the EncryptedPackage stream.
///
/// Layout: `[8 bytes: original size LE64] [encrypted segments of 4096 bytes each]`
/// Each segment IV = `Hash(salt + LE32(segment_index))` truncated to `block_size`.
pub fn decrypt_package(
    data_key: &[u8],
    info: &AgileEncryptionInfo,
    encrypted_package: &[u8],
) -> Result<Vec<u8>> {
    if encrypted_package.len() < 8 {
        cold_path();
        return Err(ModernXlsxError::PasswordProtected(
            "EncryptedPackage too short".into(),
        ));
    }

    let original_size =
        u64::from_le_bytes(encrypted_package[..8].try_into().unwrap_or_default()) as usize;
    let payload = &encrypted_package[8..];

    let block_size = info.key_block_size as usize;
    let mut result = Vec::with_capacity(original_size);

    for (segment_idx, chunk) in payload.chunks(SEGMENT_SIZE).enumerate() {
        let iv = derive_segment_iv(
            &info.key_salt,
            segment_idx as u32,
            block_size,
            &info.key_hash_alg,
        )?;
        let decrypted = aes_cbc_decrypt_no_pad(data_key, &iv, chunk)?;
        result.extend_from_slice(&decrypted);
    }

    result.truncate(original_size);
    Ok(result)
}

/// Derives the IV for a given segment index.
fn derive_segment_iv(
    salt: &[u8],
    segment_index: u32,
    block_size: usize,
    hash_alg: &str,
) -> Result<Vec<u8>> {
    let hash =
        hash_bytes_with_prefix(salt, &segment_index.to_le_bytes(), hash_alg)?;
    Ok(hash[..block_size].to_vec())
}

// ---------------------------------------------------------------------------
// HMAC integrity verification
// ---------------------------------------------------------------------------

/// Verifies HMAC integrity of the encrypted package.
pub fn verify_hmac(
    data_key: &[u8],
    info: &AgileEncryptionInfo,
    encrypted_package: &[u8],
) -> Result<()> {
    let key_len = (info.key_bits / 8) as usize;

    // 1. Derive HMAC key (SensitiveKey auto-zeroizes on Drop)
    let hmac_key_derivation = SensitiveKey::new(
        hash_bytes_with_prefix(data_key, &BLOCK_KEY_HMAC_KEY, &info.key_hash_alg)?,
    );
    let hmac_key = SensitiveKey::new(aes_cbc_decrypt(
        &hmac_key_derivation[..key_len],
        &info.key_salt,
        &info.encrypted_hmac_key,
    )?);

    // 2. Derive HMAC value key and decrypt expected HMAC
    let hmac_value_derivation = SensitiveKey::new(
        hash_bytes_with_prefix(data_key, &BLOCK_KEY_HMAC_VALUE, &info.key_hash_alg)?,
    );
    let expected_hmac = SensitiveKey::new(aes_cbc_decrypt(
        &hmac_value_derivation[..key_len],
        &info.key_salt,
        &info.encrypted_hmac_value,
    )?);

    // 3. Compute HMAC of encrypted package (entire encrypted stream including size prefix)
    let hash_len = info.key_hash_size as usize;
    let computed = compute_hmac(&info.key_hash_alg, &hmac_key[..hash_len], encrypted_package)?;

    // 4. Constant-time compare
    if computed.len() < hash_len || expected_hmac.len() < hash_len {
        cold_path();
        return Err(ModernXlsxError::PasswordProtected(
            "HMAC verification failed — truncated hash".into(),
        ));
    }
    if !constant_time_eq::constant_time_eq(&computed[..hash_len], &expected_hmac[..hash_len]) {
        cold_path();
        return Err(ModernXlsxError::PasswordProtected(
            "HMAC verification failed — file may be corrupted or tampered with.".into(),
        ));
    }

    Ok(())
}

// ---------------------------------------------------------------------------
// Segment-based encryption
// ---------------------------------------------------------------------------

/// Encrypts a ZIP package using ECMA-376 Agile segment-based encryption.
///
/// Layout: `[8 bytes: original size LE64] [encrypted segments of 4096 bytes]`
/// Each segment IV = `Hash(salt + LE32(segment_index))` truncated to `block_size`.
pub fn encrypt_package(
    data_key: &[u8],
    salt: &[u8],
    zip_bytes: &[u8],
    block_size: usize,
    hash_alg: &str,
) -> Result<Vec<u8>> {
    let original_size = zip_bytes.len() as u64;
    let num_segments = zip_bytes.len().div_ceil(SEGMENT_SIZE);

    // Pre-allocate: 8-byte header + segments (each padded to multiple of AES block)
    let mut result = Vec::with_capacity(8 + num_segments * SEGMENT_SIZE);
    result.extend_from_slice(&original_size.to_le_bytes());

    for (segment_idx, chunk) in zip_bytes.chunks(SEGMENT_SIZE).enumerate() {
        let iv = derive_segment_iv(salt, segment_idx as u32, block_size, hash_alg)?;

        // Pad chunk to AES block boundary (16 bytes)
        let mut padded = chunk.to_vec();
        let pad_len = (16 - (padded.len() % 16)) % 16;
        padded.resize(padded.len() + pad_len, 0);

        let encrypted = aes_cbc_encrypt_no_pad(data_key, &iv, &padded)?;
        result.extend_from_slice(&encrypted);
    }

    Ok(result)
}

/// Computes HMAC over the encrypted package and encrypts both the HMAC key and value.
///
/// Returns `(encrypted_hmac_key, encrypted_hmac_value)`.
pub fn compute_and_encrypt_hmac(
    data_key: &[u8],
    key_salt: &[u8],
    encrypted_package: &[u8],
    key_bits: u32,
    hash_alg: &str,
) -> Result<(Vec<u8>, Vec<u8>)> {
    let key_len = (key_bits / 8) as usize;
    let hash_len = match hash_alg {
        "SHA512" => 64,
        "SHA256" => 32,
        _ => {
            cold_path();
            return Err(ModernXlsxError::PasswordProtected(format!(
                "Unsupported hash algorithm for HMAC: {hash_alg}"
            )));
        }
    };

    // 1. Generate random HMAC key (SensitiveKey auto-zeroizes on Drop)
    let hmac_key = SensitiveKey::new(secure_random(hash_len)?);

    // 2. Compute HMAC of the encrypted package
    let hmac_value = compute_hmac(hash_alg, &hmac_key, encrypted_package)?;

    // 3. Derive encryption keys for HMAC key and value
    let hmac_key_enc_key = SensitiveKey::new(
        hash_bytes_with_prefix(data_key, &BLOCK_KEY_HMAC_KEY, hash_alg)?,
    );
    let hmac_value_enc_key = SensitiveKey::new(
        hash_bytes_with_prefix(data_key, &BLOCK_KEY_HMAC_VALUE, hash_alg)?,
    );

    // 4. Encrypt the HMAC key and value
    let encrypted_hmac_key =
        aes_cbc_encrypt(&hmac_key_enc_key[..key_len], key_salt, &hmac_key)?;
    let encrypted_hmac_value =
        aes_cbc_encrypt(&hmac_value_enc_key[..key_len], key_salt, &hmac_value)?;

    // All SensitiveKey fields auto-zeroize when dropped
    Ok((encrypted_hmac_key, encrypted_hmac_value))
}

// ---------------------------------------------------------------------------
// Full file encryption pipeline
// ---------------------------------------------------------------------------

/// Base64 encode helper (inline implementation — no external dep).
pub fn encode_base64(data: &[u8]) -> String {
    const CHARS: &[u8; 64] =
        b"ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/";

    let mut result = String::with_capacity(data.len().div_ceil(3) * 4);

    for chunk in data.chunks(3) {
        let b0 = chunk[0] as u32;
        let b1 = chunk.get(1).copied().unwrap_or(0) as u32;
        let b2 = chunk.get(2).copied().unwrap_or(0) as u32;
        let triple = (b0 << 16) | (b1 << 8) | b2;

        result.push(CHARS[((triple >> 18) & 0x3F) as usize] as char);
        result.push(CHARS[((triple >> 12) & 0x3F) as usize] as char);

        match chunk.len() {
            1 => {
                result.push('=');
                result.push('=');
            }
            2 => {
                result.push(CHARS[((triple >> 6) & 0x3F) as usize] as char);
                result.push('=');
            }
            _ => {
                result.push(CHARS[((triple >> 6) & 0x3F) as usize] as char);
                result.push(CHARS[(triple & 0x3F) as usize] as char);
            }
        }
    }

    result
}

/// Encrypts a ZIP archive with Agile Encryption (AES-256-CBC, SHA-512).
///
/// Produces a complete OLE2 compound document containing:
/// - `EncryptionInfo` stream (version 4.4 + Agile XML descriptor)
/// - `EncryptedPackage` stream (segment-encrypted ZIP data)
///
/// The output is compatible with our `decrypt_file` read path and with
/// Microsoft Excel's decryption.
pub fn encrypt_file(zip_bytes: &[u8], password: &str) -> Result<Vec<u8>> {
    // Agile Encryption parameters (AES-256-CBC, SHA-512)
    let key_bits: u32 = 256;
    let block_size: u32 = 16;
    let hash_size: u32 = 64;
    let hash_alg = "SHA512";
    let spin_count: u32 = 100_000;

    // 1. Generate random salts and data encryption key (SensitiveKey auto-zeroizes)
    let key_salt = secure_random(16)?;
    let pw_salt = secure_random(16)?;
    let data_key = SensitiveKey::new(secure_random((key_bits / 8) as usize)?);

    // 2. Encrypt the package (segment-based)
    let encrypted_package =
        encrypt_package(&data_key, &key_salt, zip_bytes, block_size as usize, hash_alg)?;

    // 3. Compute and encrypt HMAC
    let (encrypted_hmac_key, encrypted_hmac_value) =
        compute_and_encrypt_hmac(&data_key, &key_salt, &encrypted_package, key_bits, hash_alg)?;

    // 4. Generate password verifier
    let verifier_input = secure_random(16)?;
    let verifier_hash = hash_bytes(hash_alg, &verifier_input)?;

    // 5. Derive password-based keys and encrypt verifier + data key
    let input_key = SensitiveKey::new(derive_key(
        password, &pw_salt, spin_count, key_bits, &BLOCK_KEY_VERIFIER_INPUT, hash_alg,
    )?);
    let encrypted_verifier_input = aes_cbc_encrypt(&input_key, &pw_salt, &verifier_input)?;

    let value_key = SensitiveKey::new(derive_key(
        password, &pw_salt, spin_count, key_bits, &BLOCK_KEY_VERIFIER_VALUE, hash_alg,
    )?);
    let encrypted_verifier_value = aes_cbc_encrypt(&value_key, &pw_salt, &verifier_hash)?;

    let enc_key_key = SensitiveKey::new(derive_key(
        password, &pw_salt, spin_count, key_bits, &BLOCK_KEY_ENCRYPTED_KEY, hash_alg,
    )?);
    let encrypted_key_value = aes_cbc_encrypt(&enc_key_key, &pw_salt, &data_key)?;

    // All SensitiveKey fields (data_key, input_key, value_key, enc_key_key) auto-zeroize on drop

    // 7. Build EncryptionInfo XML
    let enc_info_xml = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<encryption xmlns="http://schemas.microsoft.com/office/2006/encryption" xmlns:p="http://schemas.microsoft.com/office/2006/keyEncryptor/password">
  <keyData saltSize="16" blockSize="{block_size}" keyBits="{key_bits}" hashSize="{hash_size}" cipherAlgorithm="AES" cipherChaining="ChainingModeCBC" hashAlgorithm="{hash_alg}" saltValue="{key_salt_b64}"/>
  <dataIntegrity encryptedHmacKey="{hmac_key_b64}" encryptedHmacValue="{hmac_value_b64}"/>
  <keyEncryptors>
    <keyEncryptor uri="http://schemas.microsoft.com/office/2006/keyEncryptor/password">
      <p:encryptedKey spinCount="{spin_count}" saltSize="16" blockSize="{block_size}" keyBits="{key_bits}" hashSize="{hash_size}" cipherAlgorithm="AES" cipherChaining="ChainingModeCBC" hashAlgorithm="{hash_alg}" saltValue="{pw_salt_b64}" encryptedKeyValue="{enc_key_b64}" encryptedVerifierHashInput="{enc_verifier_input_b64}" encryptedVerifierHashValue="{enc_verifier_value_b64}"/>
    </keyEncryptor>
  </keyEncryptors>
</encryption>"#,
        key_salt_b64 = encode_base64(&key_salt),
        hmac_key_b64 = encode_base64(&encrypted_hmac_key),
        hmac_value_b64 = encode_base64(&encrypted_hmac_value),
        pw_salt_b64 = encode_base64(&pw_salt),
        enc_key_b64 = encode_base64(&encrypted_key_value),
        enc_verifier_input_b64 = encode_base64(&encrypted_verifier_input),
        enc_verifier_value_b64 = encode_base64(&encrypted_verifier_value),
    );

    // 8. Build EncryptionInfo stream: version 4.4 header + XML
    let xml_bytes = enc_info_xml.as_bytes();
    let mut enc_info_stream = Vec::with_capacity(8 + xml_bytes.len());
    enc_info_stream.extend_from_slice(&4u16.to_le_bytes()); // major = 4
    enc_info_stream.extend_from_slice(&4u16.to_le_bytes()); // minor = 4
    enc_info_stream.extend_from_slice(&0x40u32.to_le_bytes()); // flags (Agile)
    enc_info_stream.extend_from_slice(xml_bytes);

    // 9. Wrap in OLE2 compound document
    super::writer::write_ole2(&[
        ("EncryptionInfo", &enc_info_stream),
        ("EncryptedPackage", &encrypted_package),
    ])
}

// ---------------------------------------------------------------------------
// AES-ECB primitives (Standard Encryption)
// ---------------------------------------------------------------------------

/// Decrypts data using AES-ECB with no padding removal.
///
/// Standard Encryption uses ECB mode for password verifier decryption.
/// Implemented using raw `BlockDecrypt` from the `aes` crate (block-by-block).
/// Automatically selects AES-128 or AES-256 based on key length.
fn aes_ecb_decrypt_no_pad(key: &[u8], data: &[u8]) -> Result<Vec<u8>> {
    use aes::cipher::{BlockDecrypt, KeyInit, generic_array::GenericArray};

    if !data.len().is_multiple_of(16) {
        cold_path();
        return Err(ModernXlsxError::PasswordProtected(
            "AES-ECB data length must be a multiple of 16".into(),
        ));
    }

    let mut buf = data.to_vec();

    match key.len() {
        16 => {
            let cipher = Aes128::new_from_slice(key)
                .map_err(|e| ModernXlsxError::PasswordProtected(format!("AES-ECB init: {e}")))?;
            for block in buf.chunks_exact_mut(16) {
                cipher.decrypt_block(GenericArray::from_mut_slice(block));
            }
        }
        32 => {
            let cipher = Aes256::new_from_slice(key)
                .map_err(|e| ModernXlsxError::PasswordProtected(format!("AES-ECB init: {e}")))?;
            for block in buf.chunks_exact_mut(16) {
                cipher.decrypt_block(GenericArray::from_mut_slice(block));
            }
        }
        n => {
            cold_path();
            return Err(ModernXlsxError::PasswordProtected(format!(
                "Unsupported AES key length for ECB: {n} bytes (expected 16 or 32)"
            )));
        }
    }

    Ok(buf)
}

/// Encrypts data using AES-ECB with no padding (for test use).
///
/// Block-by-block encryption using raw `BlockEncrypt`.
#[cfg(test)]
fn aes_ecb_encrypt_no_pad(key: &[u8], data: &[u8]) -> Result<Vec<u8>> {
    use aes::cipher::{BlockEncrypt, KeyInit, generic_array::GenericArray};

    if !data.len().is_multiple_of(16) {
        cold_path();
        return Err(ModernXlsxError::PasswordProtected(
            "AES-ECB data length must be a multiple of 16".into(),
        ));
    }

    let mut buf = data.to_vec();

    match key.len() {
        16 => {
            let cipher = Aes128::new_from_slice(key)
                .map_err(|e| ModernXlsxError::PasswordProtected(format!("AES-ECB init: {e}")))?;
            for block in buf.chunks_exact_mut(16) {
                cipher.encrypt_block(GenericArray::from_mut_slice(block));
            }
        }
        32 => {
            let cipher = Aes256::new_from_slice(key)
                .map_err(|e| ModernXlsxError::PasswordProtected(format!("AES-ECB init: {e}")))?;
            for block in buf.chunks_exact_mut(16) {
                cipher.encrypt_block(GenericArray::from_mut_slice(block));
            }
        }
        n => {
            cold_path();
            return Err(ModernXlsxError::PasswordProtected(format!(
                "Unsupported AES key length for ECB: {n} bytes (expected 16 or 32)"
            )));
        }
    }

    Ok(buf)
}

// ---------------------------------------------------------------------------
// Standard Encryption — password verification & decryption
// ---------------------------------------------------------------------------

/// Verifies a password against Standard encryption info (MS-OFFCRYPTO 2.3.6.2).
///
/// Returns the derived decryption key if the password is correct.
///
/// Key derivation (SHA-1 based):
/// 1. H0 = SHA1(salt + UTF-16LE(password))
/// 2. H_i = SHA1(LE32(i) + H_{i-1}) for `STANDARD_SPIN_COUNT` iterations
/// 3. H_final = SHA1(H_last + blockKey)
/// 4. Key = H_final[..keySize/8] (for AES-128: first 16 bytes of 20-byte SHA-1)
///
/// Verification:
/// - Decrypt `encrypted_verifier` (16 bytes) with AES-ECB
/// - Compute SHA-1(decrypted_verifier)
/// - Decrypt `encrypted_verifier_hash` (32 bytes) with AES-ECB
/// - Compare first `verifier_hash_size` bytes
pub fn verify_password_standard(
    password: &str,
    info: &StandardEncryptionInfo,
) -> Result<Vec<u8>> {
    // 1. Derive key using the generic derive_key with SHA1
    //    Wrapped in SensitiveKey so it's zeroized even on early-return error paths.
    //    On success, we move the inner Vec out before the wrapper drops.
    let mut derived = SensitiveKey::new(derive_key(
        password,
        &info.salt,
        STANDARD_SPIN_COUNT,
        info.key_size,
        &STANDARD_BLOCK_KEY,
        "SHA1",
    )?);

    // 2. Decrypt the encrypted verifier (16 bytes) with AES-ECB
    let decrypted_verifier = aes_ecb_decrypt_no_pad(&derived, &info.encrypted_verifier)?;

    // 3. Compute SHA-1 of the decrypted verifier
    let verifier_hash = hash_bytes("SHA1", &decrypted_verifier)?;

    // 4. Decrypt the encrypted verifier hash (32 bytes) with AES-ECB
    let decrypted_verifier_hash =
        aes_ecb_decrypt_no_pad(&derived, &info.encrypted_verifier_hash)?;

    // 5. Compare first verifier_hash_size bytes
    let hash_len = info.verifier_hash_size as usize;
    if verifier_hash.len() < hash_len || decrypted_verifier_hash.len() < hash_len {
        cold_path();
        return Err(ModernXlsxError::PasswordProtected(
            "Incorrect password.".into(),
        ));
    }
    if !constant_time_eq::constant_time_eq(
        &verifier_hash[..hash_len],
        &decrypted_verifier_hash[..hash_len],
    ) {
        cold_path();
        return Err(ModernXlsxError::PasswordProtected(
            "Incorrect password.".into(),
        ));
    }

    // 6. Return the derived key — take ownership, replace inner with empty (zeroized on drop)
    Ok(std::mem::take(&mut derived.0))
}

/// Decrypts the EncryptedPackage stream for Standard Encryption.
///
/// Standard encryption uses a single-stream AES-CBC decryption (NOT segmented
/// like Agile). Layout: `[8 bytes: original size LE64] [encrypted data]`.
/// IV = salt from the encryption info (not zeros).
pub fn decrypt_standard_package(
    data_key: &[u8],
    info: &StandardEncryptionInfo,
    encrypted_package: &[u8],
) -> Result<Vec<u8>> {
    if encrypted_package.len() < 8 {
        cold_path();
        return Err(ModernXlsxError::PasswordProtected(
            "EncryptedPackage too short".into(),
        ));
    }

    let original_size =
        u64::from_le_bytes(encrypted_package[..8].try_into().unwrap_or_default()) as usize;
    let payload = &encrypted_package[8..];

    if payload.is_empty() {
        cold_path();
        return Err(ModernXlsxError::PasswordProtected(
            "EncryptedPackage has no encrypted data".into(),
        ));
    }

    // Standard encryption uses AES-CBC with the salt as IV for package decryption
    let mut result = aes_cbc_decrypt_no_pad(data_key, &info.salt, payload)?;
    result.truncate(original_size);

    Ok(result)
}

// ---------------------------------------------------------------------------
// Tests
// ---------------------------------------------------------------------------

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_derive_key_known_vector() {
        // Use a simple known input and verify deterministic output
        let salt = [
            0x01, 0x02, 0x03, 0x04, 0x05, 0x06, 0x07, 0x08, 0x09, 0x0A, 0x0B, 0x0C, 0x0D, 0x0E,
            0x0F, 0x10,
        ];
        let key =
            derive_key("password", &salt, 1, 256, &BLOCK_KEY_ENCRYPTED_KEY, "SHA512").unwrap();
        assert_eq!(key.len(), 32); // 256 bits = 32 bytes

        // Verify deterministic (same input -> same output)
        let key2 =
            derive_key("password", &salt, 1, 256, &BLOCK_KEY_ENCRYPTED_KEY, "SHA512").unwrap();
        assert_eq!(key, key2);

        // Verify different block key gives different result
        let key3 = derive_key("password", &salt, 1, 256, &BLOCK_KEY_HMAC_KEY, "SHA512").unwrap();
        assert_ne!(key, key3);
    }

    #[test]
    fn test_derive_key_empty_password() {
        let salt = [0u8; 16];
        let key = derive_key("", &salt, 1, 256, &BLOCK_KEY_ENCRYPTED_KEY, "SHA512").unwrap();
        assert_eq!(key.len(), 32);

        // Empty password should still produce a valid key (not an error)
        // and should be different from non-empty password
        let key2 = derive_key("x", &salt, 1, 256, &BLOCK_KEY_ENCRYPTED_KEY, "SHA512").unwrap();
        assert_ne!(key, key2);
    }

    #[test]
    fn test_derive_key_unicode() {
        let salt = [0x42u8; 16];
        // Russian word for "password"
        let key = derive_key(
            "\u{043F}\u{0430}\u{0440}\u{043E}\u{043B}\u{044C}",
            &salt,
            1,
            256,
            &BLOCK_KEY_ENCRYPTED_KEY,
            "SHA512",
        )
        .unwrap();
        assert_eq!(key.len(), 32);

        // Different from ASCII password
        let key2 =
            derive_key("password", &salt, 1, 256, &BLOCK_KEY_ENCRYPTED_KEY, "SHA512").unwrap();
        assert_ne!(key, key2);

        // Verify deterministic
        let key3 = derive_key(
            "\u{043F}\u{0430}\u{0440}\u{043E}\u{043B}\u{044C}",
            &salt,
            1,
            256,
            &BLOCK_KEY_ENCRYPTED_KEY,
            "SHA512",
        )
        .unwrap();
        assert_eq!(key, key3);
    }

    #[test]
    fn test_derive_key_128_bit() {
        let salt = [0xAA; 16];
        let key = derive_key("test", &salt, 1, 128, &BLOCK_KEY_ENCRYPTED_KEY, "SHA256").unwrap();
        assert_eq!(key.len(), 16); // 128 bits = 16 bytes
    }

    #[test]
    fn test_derive_key_sha256() {
        let salt = [0xBB; 16];
        let key_256 =
            derive_key("test", &salt, 1, 128, &BLOCK_KEY_ENCRYPTED_KEY, "SHA256").unwrap();
        let key_512 =
            derive_key("test", &salt, 1, 128, &BLOCK_KEY_ENCRYPTED_KEY, "SHA512").unwrap();
        // SHA-256 and SHA-512 should produce different keys even with same inputs
        assert_ne!(key_256, key_512);
        assert_eq!(key_256.len(), 16);
        assert_eq!(key_512.len(), 16);
    }

    // --- AES-CBC tests ---

    #[test]
    fn test_aes_cbc_decrypt_known_vector() {
        // AES-256-CBC encrypt-then-decrypt roundtrip
        use cbc::cipher::{BlockEncryptMut, block_padding::Pkcs7 as Pkcs7Pad};
        type Aes256CbcEnc = cbc::Encryptor<Aes256>;

        let key = [0x60u8; 32];
        let iv = [0x00u8; 16];

        let plaintext = b"Hello AES-256-CBC Test!";
        let encryptor = Aes256CbcEnc::new_from_slices(&key, &iv).unwrap();
        let ciphertext = encryptor.encrypt_padded_vec_mut::<Pkcs7Pad>(plaintext);

        let decrypted = aes_cbc_decrypt(&key, &iv, &ciphertext).unwrap();
        assert_eq!(decrypted, plaintext);
    }

    #[test]
    fn test_aes_cbc_decrypt_no_pad_roundtrip() {
        use cbc::cipher::{BlockEncryptMut, block_padding::NoPadding as NoPad};
        type Aes256CbcEnc = cbc::Encryptor<Aes256>;

        let key = [0x42u8; 32];
        let iv = [0x01u8; 16];

        // Must be exact multiple of block size (16)
        let plaintext = [0xAA; 32]; // 2 blocks
        let encryptor = Aes256CbcEnc::new_from_slices(&key, &iv).unwrap();
        let ciphertext = encryptor.encrypt_padded_vec_mut::<NoPad>(&plaintext);

        let decrypted = aes_cbc_decrypt_no_pad(&key, &iv, &ciphertext).unwrap();
        assert_eq!(decrypted, plaintext);
    }

    #[test]
    fn test_decrypt_single_segment() {
        // Create a fake encrypted package: 8-byte size + 1 segment of encrypted data
        use cbc::cipher::{BlockEncryptMut, block_padding::NoPadding as NoPad};

        let key = [0x33u8; 32];
        let salt = [0x11u8; 16];
        let original = vec![0xBB; 100]; // 100 bytes original data

        // Encrypt one segment: pad original to 4096, encrypt with segment IV
        let mut padded = original.clone();
        padded.resize(4096, 0); // Pad to segment size

        // Derive segment 0 IV: SHA-512(salt + LE32(0)), truncated to 16
        let mut hasher = Sha512::new();
        Digest::update(&mut hasher, salt);
        Digest::update(&mut hasher, 0u32.to_le_bytes());
        let hash = hasher.finalize();
        let iv: Vec<u8> = hash[..16].to_vec();

        let encryptor = cbc::Encryptor::<Aes256>::new_from_slices(&key, &iv).unwrap();
        let ciphertext = encryptor.encrypt_padded_vec_mut::<NoPad>(&padded);

        // Build encrypted package: size(8) + ciphertext
        let mut package = Vec::new();
        package.extend_from_slice(&(100u64).to_le_bytes());
        package.extend_from_slice(&ciphertext);

        // Build minimal AgileEncryptionInfo
        let info = AgileEncryptionInfo {
            key_salt: salt.to_vec(),
            key_block_size: 16,
            key_bits: 256,
            key_hash_size: 64,
            key_cipher: "AES".into(),
            key_chaining: "ChainingModeCBC".into(),
            key_hash_alg: "SHA512".into(),
            encrypted_hmac_key: vec![],
            encrypted_hmac_value: vec![],
            pw_spin_count: 100000,
            pw_salt: vec![],
            pw_block_size: 16,
            pw_key_bits: 256,
            pw_hash_size: 64,
            pw_cipher: "AES".into(),
            pw_chaining: "ChainingModeCBC".into(),
            pw_hash_alg: "SHA512".into(),
            pw_encrypted_key_value: vec![],
            pw_encrypted_verifier_hash_input: vec![],
            pw_encrypted_verifier_hash_value: vec![],
        };

        let result = decrypt_package(&key, &info, &package).unwrap();
        assert_eq!(result.len(), 100);
        assert_eq!(result, original);
    }

    #[test]
    fn test_decrypt_multi_segment() {
        // 3 segments = 12288 bytes encrypted + last partial
        use cbc::cipher::{BlockEncryptMut, block_padding::NoPadding as NoPad};

        let key = [0x44u8; 32];
        let salt = [0x22u8; 16];
        let original = vec![0xCC; 10000]; // 10000 bytes = 2 full segments + 1 partial

        let mut package = Vec::new();
        package.extend_from_slice(&(10000u64).to_le_bytes());

        // Encrypt each segment
        let padded_size = if original.len().is_multiple_of(4096) {
            original.len()
        } else {
            ((original.len() / 4096) + 1) * 4096
        };
        let mut padded = original.clone();
        padded.resize(padded_size, 0);

        for (seg_idx, chunk) in padded.chunks(4096).enumerate() {
            let mut hasher = Sha512::new();
            Digest::update(&mut hasher, salt);
            Digest::update(&mut hasher, (seg_idx as u32).to_le_bytes());
            let hash = hasher.finalize();
            let iv: Vec<u8> = hash[..16].to_vec();

            let encryptor = cbc::Encryptor::<Aes256>::new_from_slices(&key, &iv).unwrap();
            let ct = encryptor.encrypt_padded_vec_mut::<NoPad>(chunk);
            package.extend_from_slice(&ct);
        }

        let info = AgileEncryptionInfo {
            key_salt: salt.to_vec(),
            key_block_size: 16,
            key_bits: 256,
            key_hash_size: 64,
            key_cipher: "AES".into(),
            key_chaining: "ChainingModeCBC".into(),
            key_hash_alg: "SHA512".into(),
            encrypted_hmac_key: vec![],
            encrypted_hmac_value: vec![],
            pw_spin_count: 100000,
            pw_salt: vec![],
            pw_block_size: 16,
            pw_key_bits: 256,
            pw_hash_size: 64,
            pw_cipher: "AES".into(),
            pw_chaining: "ChainingModeCBC".into(),
            pw_hash_alg: "SHA512".into(),
            pw_encrypted_key_value: vec![],
            pw_encrypted_verifier_hash_input: vec![],
            pw_encrypted_verifier_hash_value: vec![],
        };

        let result = decrypt_package(&key, &info, &package).unwrap();
        assert_eq!(result.len(), 10000);
        assert_eq!(result, original);
    }

    #[test]
    fn test_encrypted_package_too_short() {
        let key = [0u8; 32];
        let info = AgileEncryptionInfo {
            key_salt: vec![0; 16],
            key_block_size: 16,
            key_bits: 256,
            key_hash_size: 64,
            key_cipher: "AES".into(),
            key_chaining: "ChainingModeCBC".into(),
            key_hash_alg: "SHA512".into(),
            encrypted_hmac_key: vec![],
            encrypted_hmac_value: vec![],
            pw_spin_count: 100000,
            pw_salt: vec![],
            pw_block_size: 16,
            pw_key_bits: 256,
            pw_hash_size: 64,
            pw_cipher: "AES".into(),
            pw_chaining: "ChainingModeCBC".into(),
            pw_hash_alg: "SHA512".into(),
            pw_encrypted_key_value: vec![],
            pw_encrypted_verifier_hash_input: vec![],
            pw_encrypted_verifier_hash_value: vec![],
        };
        let result = decrypt_package(&key, &info, &[0u8; 4]); // too short
        assert!(result.is_err());
        assert!(result.unwrap_err().to_string().contains("too short"));
    }

    // --- Standard Encryption tests ---

    #[test]
    fn test_standard_derive_key_sha1() {
        // SHA-1 key derivation produces a 16-byte key (AES-128) and is deterministic
        let salt = [0x01u8; 16];
        let key = derive_key("password", &salt, 1, 128, &STANDARD_BLOCK_KEY, "SHA1").unwrap();
        assert_eq!(key.len(), 16); // 128 bits = 16 bytes

        // Deterministic
        let key2 = derive_key("password", &salt, 1, 128, &STANDARD_BLOCK_KEY, "SHA1").unwrap();
        assert_eq!(key, key2);

        // Different password => different key
        let key3 = derive_key("other", &salt, 1, 128, &STANDARD_BLOCK_KEY, "SHA1").unwrap();
        assert_ne!(key, key3);

        // SHA-1 key differs from SHA-256 key
        let key_256 =
            derive_key("password", &salt, 1, 128, &STANDARD_BLOCK_KEY, "SHA256").unwrap();
        assert_ne!(key, key_256);
    }

    #[test]
    fn test_standard_verify_password_roundtrip() {
        // Create a known StandardEncryptionInfo by encrypting a verifier with a derived key.
        let password = "test123";
        let salt = [0x42u8; 16];

        // Derive the key the same way verify_password_standard will
        let key =
            derive_key(password, &salt, STANDARD_SPIN_COUNT, 128, &STANDARD_BLOCK_KEY, "SHA1")
                .unwrap();

        // Create a verifier (16 bytes, block-aligned)
        let verifier = [0xDE; 16];

        // Compute SHA-1 hash of verifier
        let mut hasher = Sha1::new();
        Digest::update(&mut hasher, verifier);
        let verifier_hash = hasher.finalize();

        // Encrypt verifier with AES-128-ECB
        let encrypted_verifier = aes_ecb_encrypt_no_pad(&key, &verifier).unwrap();

        // Encrypt verifier hash (pad to 32 bytes for AES-ECB block alignment)
        let mut hash_padded = verifier_hash.to_vec();
        hash_padded.resize(32, 0); // pad to 32 bytes (2 AES blocks)
        let encrypted_verifier_hash = aes_ecb_encrypt_no_pad(&key, &hash_padded).unwrap();

        let info = StandardEncryptionInfo {
            alg_id: 0x6801,      // AES-128
            hash_alg_id: 0x8004, // SHA-1
            key_size: 128,
            provider: "Microsoft Enhanced RSA and AES Cryptographic Provider".into(),
            salt: salt.to_vec(),
            encrypted_verifier,
            verifier_hash_size: 20,
            encrypted_verifier_hash,
        };

        let result = verify_password_standard(password, &info);
        assert!(result.is_ok(), "verify_password_standard failed: {result:?}");
        let derived = result.unwrap();
        assert_eq!(derived.len(), 16);
        assert_eq!(derived, key);
    }

    #[test]
    fn test_standard_wrong_password() {
        // Same setup as roundtrip but verify with wrong password
        let password = "correct";
        let salt = [0x55u8; 16];

        let key =
            derive_key(password, &salt, STANDARD_SPIN_COUNT, 128, &STANDARD_BLOCK_KEY, "SHA1")
                .unwrap();

        let verifier = [0xAB; 16];
        let mut hasher = Sha1::new();
        Digest::update(&mut hasher, verifier);
        let verifier_hash = hasher.finalize();

        let encrypted_verifier = aes_ecb_encrypt_no_pad(&key, &verifier).unwrap();

        let mut hash_padded = verifier_hash.to_vec();
        hash_padded.resize(32, 0);
        let encrypted_verifier_hash = aes_ecb_encrypt_no_pad(&key, &hash_padded).unwrap();

        let info = StandardEncryptionInfo {
            alg_id: 0x6801,
            hash_alg_id: 0x8004,
            key_size: 128,
            provider: "Microsoft Enhanced RSA and AES Cryptographic Provider".into(),
            salt: salt.to_vec(),
            encrypted_verifier,
            verifier_hash_size: 20,
            encrypted_verifier_hash,
        };

        let result = verify_password_standard("wrong_password", &info);
        assert!(result.is_err());
        assert!(
            result.unwrap_err().to_string().contains("Incorrect password"),
            "Expected 'Incorrect password' error"
        );
    }

    #[test]
    fn test_standard_decrypt_package() {
        // Encrypt-then-decrypt roundtrip for Standard single-stream AES-CBC
        use cbc::cipher::{BlockEncryptMut, block_padding::NoPadding as NoPad};

        let key = [0x77u8; 16]; // AES-128
        let salt = [0x33u8; 16]; // Used as IV
        let original = vec![0xEE; 200];

        // Pad to block boundary for AES-CBC
        let mut padded = original.clone();
        let pad_len = (16 - (padded.len() % 16)) % 16;
        padded.resize(padded.len() + pad_len, 0);

        let encryptor = cbc::Encryptor::<Aes128>::new_from_slices(&key, &salt).unwrap();
        let ciphertext = encryptor.encrypt_padded_vec_mut::<NoPad>(&padded);

        // Build package: 8-byte size + ciphertext
        let mut package = Vec::new();
        package.extend_from_slice(&(200u64).to_le_bytes());
        package.extend_from_slice(&ciphertext);

        let info = StandardEncryptionInfo {
            alg_id: 0x6801,
            hash_alg_id: 0x8004,
            key_size: 128,
            provider: "AES".into(),
            salt: salt.to_vec(),
            encrypted_verifier: vec![0; 16],
            verifier_hash_size: 20,
            encrypted_verifier_hash: vec![0; 32],
        };

        let result = decrypt_standard_package(&key, &info, &package).unwrap();
        assert_eq!(result.len(), 200);
        assert_eq!(result, original);
    }

    #[test]
    fn test_corrupted_encrypted_package() {
        let key = [0u8; 16];
        let info = StandardEncryptionInfo {
            alg_id: 0x6801,
            hash_alg_id: 0x8004,
            key_size: 128,
            provider: "AES".into(),
            salt: vec![0; 16],
            encrypted_verifier: vec![0; 16],
            verifier_hash_size: 20,
            encrypted_verifier_hash: vec![0; 32],
        };

        // Truncated package (only 4 bytes — too short for the 8-byte size header)
        let result = decrypt_standard_package(&key, &info, &[0u8; 4]);
        assert!(result.is_err());
        assert!(result.unwrap_err().to_string().contains("too short"));

        // Package with size header but no encrypted data
        let mut empty_package = vec![0u8; 8];
        empty_package[..8].copy_from_slice(&(100u64).to_le_bytes());
        let result2 = decrypt_standard_package(&key, &info, &empty_package);
        assert!(result2.is_err());
        assert!(result2.unwrap_err().to_string().contains("no encrypted data"));
    }

    // --- Encryption write tests ---

    #[test]
    fn test_encrypt_decrypt_roundtrip() {
        // Create a minimal ZIP-like payload, encrypt it, then decrypt and verify match.
        let original_data = b"PK\x03\x04 This is test ZIP content for encryption roundtrip.";

        let encrypted_ole2 = encrypt_file(original_data, "test_password").unwrap();

        // The output should be a valid OLE2 file.
        assert!(
            encrypted_ole2.len() > 512,
            "OLE2 output should be larger than a header"
        );

        // Decrypt it using our read path.
        let enc_info_stream =
            crate::ole2::detect::read_stream(&encrypted_ole2, "EncryptionInfo").unwrap();
        let encrypted_package =
            crate::ole2::detect::read_stream(&encrypted_ole2, "EncryptedPackage").unwrap();

        let info = crate::ole2::encryption_info::EncryptionInfo::parse(&enc_info_stream).unwrap();
        let agile = match info {
            crate::ole2::encryption_info::EncryptionInfo::Agile(a) => a,
            other => {
                cold_path();
                unreachable!("Expected Agile encryption info, got {other:?}")
            }
        };

        // Verify password and get data key.
        let data_key = verify_password_agile("test_password", &agile).unwrap();

        // Verify HMAC integrity.
        verify_hmac(&data_key, &agile, &encrypted_package).unwrap();

        // Decrypt the package.
        let decrypted = decrypt_package(&data_key, &agile, &encrypted_package).unwrap();

        assert_eq!(decrypted, original_data);
    }

    #[test]
    fn test_encrypt_package_segment_sizes() {
        // Verify the encrypted package has correct structure:
        // 8-byte LE64 header + AES-block-aligned encrypted segments.
        let data_key = [0x55u8; 32];
        let salt = [0xAA; 16];
        let payload = vec![0xBB; 10000]; // 2 full segments + 1 partial

        let encrypted =
            encrypt_package(&data_key, &salt, &payload, 16, "SHA512").unwrap();

        // First 8 bytes = original size.
        let original_size = u64::from_le_bytes(encrypted[..8].try_into().unwrap_or_default());
        assert_eq!(original_size, 10000);

        // Total encrypted payload (after 8-byte header):
        // Segment 0: 4096 bytes -> already block-aligned -> 4096 encrypted
        // Segment 1: 4096 bytes -> already block-aligned -> 4096 encrypted
        // Segment 2: 1808 bytes -> padded to 1808 + (16 - 1808%16)%16 = 1808 + 0 = 1808
        //   1808 % 16 = 0, so no extra padding needed
        let encrypted_payload_len = encrypted.len() - 8;
        // Each segment is padded to 16-byte boundary before encryption
        let seg0 = 4096; // already aligned
        let seg1 = 4096; // already aligned
        let seg2_raw = 10000 - 4096 * 2; // 1808
        let seg2 = seg2_raw + (16 - seg2_raw % 16) % 16; // pad to 16
        assert_eq!(encrypted_payload_len, seg0 + seg1 + seg2);
    }

    #[test]
    fn test_secure_random_uniqueness() {
        let a = secure_random(32).unwrap();
        let b = secure_random(32).unwrap();
        assert_eq!(a.len(), 32);
        assert_eq!(b.len(), 32);
        // Two random outputs should be different (probability of collision is negligible).
        assert_ne!(a, b);
    }

    #[test]
    fn test_encode_base64_known_values() {
        // Known base64 values.
        assert_eq!(encode_base64(b""), "");
        assert_eq!(encode_base64(b"f"), "Zg==");
        assert_eq!(encode_base64(b"fo"), "Zm8=");
        assert_eq!(encode_base64(b"foo"), "Zm9v");
        assert_eq!(encode_base64(b"foobar"), "Zm9vYmFy");
        assert_eq!(encode_base64(b"Hello, World!"), "SGVsbG8sIFdvcmxkIQ==");

        // Binary data round trip: encode then verify format.
        let binary = vec![0u8, 1, 2, 255, 254, 128];
        let encoded = encode_base64(&binary);
        assert!(
            encoded
                .chars()
                .all(|c| c.is_ascii_alphanumeric() || c == '+' || c == '/' || c == '='),
            "Base64 output should contain only valid characters"
        );
    }
}
