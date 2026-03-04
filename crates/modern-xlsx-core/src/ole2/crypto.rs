//! ECMA-376 SS2.3.6.2 Agile Encryption key derivation.
//!
//! Implements password-based key derivation for OOXML Agile Encryption using
//! SHA-512 or SHA-256. The derived key is used to decrypt the encrypted package
//! key, HMAC key, and verifier values.

use sha2::{Digest, Sha256, Sha512, digest::FixedOutputReset};

use crate::errors::ModernXlsxError;

type Result<T> = std::result::Result<T, ModernXlsxError>;

/// Block key constants per ECMA-376 SS2.3.6.2.
pub const BLOCK_KEY_VERIFIER_INPUT: [u8; 8] = [0xFE, 0xA7, 0xD2, 0x76, 0x3B, 0x4B, 0x9E, 0x79];
pub const BLOCK_KEY_VERIFIER_VALUE: [u8; 8] = [0xD7, 0xAA, 0x0F, 0x6D, 0x30, 0x61, 0x34, 0x4E];
pub const BLOCK_KEY_ENCRYPTED_KEY: [u8; 8] = [0x14, 0x6E, 0x0B, 0xE7, 0xAB, 0xAC, 0xD0, 0xD6];
pub const BLOCK_KEY_HMAC_KEY: [u8; 8] = [0x5F, 0xB2, 0xAD, 0x01, 0x0C, 0xB9, 0xE1, 0xF6];
pub const BLOCK_KEY_HMAC_VALUE: [u8; 8] = [0xA0, 0x67, 0x7F, 0x02, 0xB2, 0x2C, 0x84, 0x33];

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
        _ => Err(ModernXlsxError::PasswordProtected(format!(
            "Unsupported hash algorithm: {hash_alg}"
        ))),
    }
}

/// Generic key derivation over any SHA-2 digest.
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
}
