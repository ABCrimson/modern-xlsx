//! ECMA-376 Agile + Standard Encryption key derivation and AES decryption.
//!
//! Implements password-based key derivation, password verification, AES-CBC/ECB
//! decryption (segment-based for Agile, single-stream for Standard), and HMAC
//! integrity verification for OOXML encryption (SHA-512/SHA-256/SHA-1,
//! AES-128/AES-256).

use aes::{Aes128, Aes256};
use cbc::cipher::{BlockDecryptMut, KeyIvInit, block_padding::{NoPadding, Pkcs7}};
use hmac::{Hmac, Mac};
use sha1::Sha1;
use sha2::{Digest, Sha256, Sha512, digest::FixedOutputReset};

use super::encryption_info::{AgileEncryptionInfo, StandardEncryptionInfo};
use crate::errors::ModernXlsxError;

type Result<T> = std::result::Result<T, ModernXlsxError>;

type Aes128CbcDec = cbc::Decryptor<Aes128>;
type Aes256CbcDec = cbc::Decryptor<Aes256>;

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
        _ => Err(ModernXlsxError::PasswordProtected(format!(
            "Unsupported hash algorithm: {hash_alg}"
        ))),
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
        n => Err(ModernXlsxError::PasswordProtected(format!(
            "Unsupported AES key length: {n} bytes (expected 16 or 32)"
        ))),
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
        n => Err(ModernXlsxError::PasswordProtected(format!(
            "Unsupported AES key length: {n} bytes (expected 16 or 32)"
        ))),
    }
}

// ---------------------------------------------------------------------------
// Hashing helpers
// ---------------------------------------------------------------------------

/// Hash bytes with the named algorithm.
fn hash_bytes(alg: &str, data: &[u8]) -> Result<Vec<u8>> {
    match alg {
        "SHA512" => {
            let mut h = Sha512::new();
            Digest::update(&mut h, data);
            Ok(h.finalize().to_vec())
        }
        "SHA256" => {
            let mut h = Sha256::new();
            Digest::update(&mut h, data);
            Ok(h.finalize().to_vec())
        }
        "SHA1" => {
            let mut h = Sha1::new();
            Digest::update(&mut h, data);
            Ok(h.finalize().to_vec())
        }
        _ => Err(ModernXlsxError::PasswordProtected(format!(
            "Unsupported hash algorithm: {alg}"
        ))),
    }
}

/// Hash the concatenation of two byte slices: H(prefix || data).
fn hash_bytes_with_prefix(prefix: &[u8], data: &[u8], alg: &str) -> Result<Vec<u8>> {
    match alg {
        "SHA512" => {
            let mut h = Sha512::new();
            Digest::update(&mut h, prefix);
            Digest::update(&mut h, data);
            Ok(h.finalize().to_vec())
        }
        "SHA256" => {
            let mut h = Sha256::new();
            Digest::update(&mut h, prefix);
            Digest::update(&mut h, data);
            Ok(h.finalize().to_vec())
        }
        "SHA1" => {
            let mut h = Sha1::new();
            Digest::update(&mut h, prefix);
            Digest::update(&mut h, data);
            Ok(h.finalize().to_vec())
        }
        _ => Err(ModernXlsxError::PasswordProtected(format!(
            "Unsupported hash algorithm: {alg}"
        ))),
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
    // 1. Derive verifier hash input key
    let input_key = derive_key(
        password,
        &info.pw_salt,
        info.pw_spin_count,
        info.pw_key_bits,
        &BLOCK_KEY_VERIFIER_INPUT,
        &info.pw_hash_alg,
    )?;

    // 2. Decrypt verifier hash input
    let verifier_input = aes_cbc_decrypt(
        &input_key,
        &info.pw_salt,
        &info.pw_encrypted_verifier_hash_input,
    )?;

    // 3. Derive verifier hash value key
    let value_key = derive_key(
        password,
        &info.pw_salt,
        info.pw_spin_count,
        info.pw_key_bits,
        &BLOCK_KEY_VERIFIER_VALUE,
        &info.pw_hash_alg,
    )?;

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
        return Err(ModernXlsxError::PasswordProtected(
            "Incorrect password.".into(),
        ));
    }
    if !constant_time_eq::constant_time_eq(&computed[..hash_len], &verifier_hash[..hash_len]) {
        return Err(ModernXlsxError::PasswordProtected(
            "Incorrect password.".into(),
        ));
    }

    // 7. Derive data encryption key
    let enc_key_key = derive_key(
        password,
        &info.pw_salt,
        info.pw_spin_count,
        info.pw_key_bits,
        &BLOCK_KEY_ENCRYPTED_KEY,
        &info.pw_hash_alg,
    )?;

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
        return Err(ModernXlsxError::PasswordProtected(
            "EncryptedPackage too short".into(),
        ));
    }

    let original_size =
        u64::from_le_bytes(encrypted_package[..8].try_into().unwrap()) as usize;
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

    // 1. Derive HMAC key: Hash(dataKey + blockKeyHmacKey), decrypt encrypted HMAC key
    let hmac_key_derivation =
        hash_bytes_with_prefix(data_key, &BLOCK_KEY_HMAC_KEY, &info.key_hash_alg)?;
    let hmac_key = aes_cbc_decrypt(
        &hmac_key_derivation[..key_len],
        &info.key_salt,
        &info.encrypted_hmac_key,
    )?;

    // 2. Derive HMAC value key and decrypt expected HMAC
    let hmac_value_derivation =
        hash_bytes_with_prefix(data_key, &BLOCK_KEY_HMAC_VALUE, &info.key_hash_alg)?;
    let expected_hmac = aes_cbc_decrypt(
        &hmac_value_derivation[..key_len],
        &info.key_salt,
        &info.encrypted_hmac_value,
    )?;

    // 3. Compute HMAC of encrypted package (entire encrypted stream including size prefix)
    let hash_len = info.key_hash_size as usize;
    let computed = match info.key_hash_alg.as_str() {
        "SHA512" => {
            let mut mac = Hmac::<Sha512>::new_from_slice(&hmac_key[..hash_len])
                .map_err(|e| {
                    ModernXlsxError::PasswordProtected(format!("HMAC init: {e}"))
                })?;
            mac.update(encrypted_package);
            mac.finalize().into_bytes().to_vec()
        }
        "SHA256" => {
            let mut mac = Hmac::<Sha256>::new_from_slice(&hmac_key[..hash_len])
                .map_err(|e| {
                    ModernXlsxError::PasswordProtected(format!("HMAC init: {e}"))
                })?;
            mac.update(encrypted_package);
            mac.finalize().into_bytes().to_vec()
        }
        _ => {
            return Err(ModernXlsxError::PasswordProtected(format!(
                "Unsupported hash algorithm for HMAC: {}",
                info.key_hash_alg
            )));
        }
    };

    // 4. Constant-time compare
    if computed.len() < hash_len || expected_hmac.len() < hash_len {
        return Err(ModernXlsxError::PasswordProtected(
            "HMAC verification failed — truncated hash".into(),
        ));
    }
    if !constant_time_eq::constant_time_eq(&computed[..hash_len], &expected_hmac[..hash_len]) {
        return Err(ModernXlsxError::PasswordProtected(
            "HMAC verification failed — file may be corrupted or tampered with.".into(),
        ));
    }

    Ok(())
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
    let derived_key = derive_key(
        password,
        &info.salt,
        STANDARD_SPIN_COUNT,
        info.key_size,
        &STANDARD_BLOCK_KEY,
        "SHA1",
    )?;

    // 2. Decrypt the encrypted verifier (16 bytes) with AES-ECB
    let decrypted_verifier = aes_ecb_decrypt_no_pad(&derived_key, &info.encrypted_verifier)?;

    // 3. Compute SHA-1 of the decrypted verifier
    let verifier_hash = hash_bytes("SHA1", &decrypted_verifier)?;

    // 4. Decrypt the encrypted verifier hash (32 bytes) with AES-ECB
    let decrypted_verifier_hash =
        aes_ecb_decrypt_no_pad(&derived_key, &info.encrypted_verifier_hash)?;

    // 5. Compare first verifier_hash_size bytes
    let hash_len = info.verifier_hash_size as usize;
    if verifier_hash.len() < hash_len || decrypted_verifier_hash.len() < hash_len {
        return Err(ModernXlsxError::PasswordProtected(
            "Incorrect password.".into(),
        ));
    }
    if !constant_time_eq::constant_time_eq(
        &verifier_hash[..hash_len],
        &decrypted_verifier_hash[..hash_len],
    ) {
        return Err(ModernXlsxError::PasswordProtected(
            "Incorrect password.".into(),
        ));
    }

    // 6. Return the derived key for package decryption
    Ok(derived_key)
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
        return Err(ModernXlsxError::PasswordProtected(
            "EncryptedPackage too short".into(),
        ));
    }

    let original_size =
        u64::from_le_bytes(encrypted_package[..8].try_into().unwrap()) as usize;
    let payload = &encrypted_package[8..];

    if payload.is_empty() {
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
        Digest::update(&mut hasher, &salt);
        Digest::update(&mut hasher, &0u32.to_le_bytes());
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
        let padded_size = if original.len() % 4096 == 0 {
            original.len()
        } else {
            ((original.len() / 4096) + 1) * 4096
        };
        let mut padded = original.clone();
        padded.resize(padded_size, 0);

        for (seg_idx, chunk) in padded.chunks(4096).enumerate() {
            let mut hasher = Sha512::new();
            Digest::update(&mut hasher, &salt);
            Digest::update(&mut hasher, &(seg_idx as u32).to_le_bytes());
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
        Digest::update(&mut hasher, &verifier);
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
        Digest::update(&mut hasher, &verifier);
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
}
