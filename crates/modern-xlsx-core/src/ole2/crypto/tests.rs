use super::*;
use aes::Aes256;

use super::agile::{
    compute_and_encrypt_hmac, derive_segment_iv, encrypt_package, verify_hmac_legacy,
    verify_hmac_spec,
};
use super::standard::{aes_ecb_decrypt_no_pad, aes_ecb_encrypt_no_pad};
use crate::ole2::base64::encode_base64;
use crate::ole2::encryption_info::{AgileEncryptionInfo, StandardEncryptionInfo};

/// Builds an `AgileEncryptionInfo` with SHA-512 (64-byte digest) defaults and
/// otherwise-empty fields, for panic-safety regression tests.
fn agile_info_sha512() -> AgileEncryptionInfo {
    AgileEncryptionInfo {
        key_salt: vec![0u8; 16],
        key_block_size: 16,
        key_bits: 256,
        key_hash_size: 64,
        key_hash_alg: HashAlg::Sha512,
        encrypted_hmac_key: vec![0u8; 64],
        encrypted_hmac_value: vec![0u8; 64],
        pw_spin_count: 1,
        pw_salt: vec![0u8; 16],
        pw_key_bits: 256,
        pw_hash_size: 64,
        pw_hash_alg: HashAlg::Sha512,
        pw_encrypted_key_value: vec![0u8; 32],
        pw_encrypted_verifier_hash_input: vec![0u8; 16],
        pw_encrypted_verifier_hash_value: vec![0u8; 64],
    }
}

/// An out-of-range `keyBits` (2048, so 256 bytes) asks the modern-xlsx 1.2.x
/// HMAC scheme in `verify_hmac` for more key bytes than the 64-byte SHA-512
/// digest holds. The checked prefix makes that an `Err`, never a panic.
#[test]
fn verify_hmac_rejects_out_of_range_key_bits_without_panicking() {
    let mut info = agile_info_sha512();
    info.key_bits = 2048; // 2048/8 = 256 > 64-byte digest
    let data_key = vec![0u8; 32];
    let package = vec![0u8; 64];
    let result = verify_hmac(&data_key, &info, &package);
    assert!(
        result.is_err(),
        "out-of-range keyBits must error, not panic"
    );
}

/// Companion: an out-of-range `hashSize` (9999) must error, not panic by
/// slicing the decrypted HMAC key past its end.
#[test]
fn verify_hmac_rejects_out_of_range_hash_size_without_panicking() {
    let mut info = agile_info_sha512();
    info.key_hash_size = 9999;
    let data_key = vec![0u8; 32];
    let package = vec![0u8; 64];
    let result = verify_hmac(&data_key, &info, &package);
    assert!(
        result.is_err(),
        "out-of-range hashSize must error, not panic"
    );
}

/// `derive_segment_iv` must not slice a digest by an unvalidated
/// `blockSize`: an over-large block size errors rather than panics.
#[test]
fn derive_segment_iv_rejects_oversized_block_size_without_panicking() {
    // SHA-512 digest is 64 bytes; ask for 256.
    let result = derive_segment_iv(&[0u8; 16], 0, 256, HashAlg::Sha512);
    assert!(result.is_err(), "oversized blockSize must error, not panic");
    // blockSize 16 yields the first 16 bytes of SHA-512(salt || LE32(0)).
    let ok = derive_segment_iv(&[0u8; 16], 0, 16, HashAlg::Sha512).unwrap();
    let full = Sha512::new_with_prefix([0u8; 16])
        .chain_update(0u32.to_le_bytes())
        .finalize();
    assert_eq!(ok[..], full[..16]);
}

#[test]
fn test_derive_key_known_vector() {
    // Use a simple known input and verify deterministic output
    let salt = [
        0x01, 0x02, 0x03, 0x04, 0x05, 0x06, 0x07, 0x08, 0x09, 0x0A, 0x0B, 0x0C, 0x0D, 0x0E, 0x0F,
        0x10,
    ];
    let key = derive_key(
        "password",
        &salt,
        1,
        256,
        &BLOCK_KEY_ENCRYPTED_KEY,
        HashAlg::Sha512,
    )
    .unwrap();
    assert_eq!(key.len(), 32); // 256 bits = 32 bytes

    // Verify deterministic (same input -> same output)
    let key2 = derive_key(
        "password",
        &salt,
        1,
        256,
        &BLOCK_KEY_ENCRYPTED_KEY,
        HashAlg::Sha512,
    )
    .unwrap();
    assert_eq!(key, key2);

    // Verify different block key gives different result
    let key3 = derive_key(
        "password",
        &salt,
        1,
        256,
        &BLOCK_KEY_HMAC_KEY,
        HashAlg::Sha512,
    )
    .unwrap();
    assert_ne!(key, key3);
}

#[test]
fn test_derive_key_empty_password() {
    let salt = [0u8; 16];
    let key = derive_key("", &salt, 1, 256, &BLOCK_KEY_ENCRYPTED_KEY, HashAlg::Sha512).unwrap();
    assert_eq!(key.len(), 32);

    // Empty password should still produce a valid key (not an error)
    // and should be different from non-empty password
    let key2 = derive_key(
        "x",
        &salt,
        1,
        256,
        &BLOCK_KEY_ENCRYPTED_KEY,
        HashAlg::Sha512,
    )
    .unwrap();
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
        HashAlg::Sha512,
    )
    .unwrap();
    assert_eq!(key.len(), 32);

    // Different from ASCII password
    let key2 = derive_key(
        "password",
        &salt,
        1,
        256,
        &BLOCK_KEY_ENCRYPTED_KEY,
        HashAlg::Sha512,
    )
    .unwrap();
    assert_ne!(key, key2);

    // Verify deterministic
    let key3 = derive_key(
        "\u{043F}\u{0430}\u{0440}\u{043E}\u{043B}\u{044C}",
        &salt,
        1,
        256,
        &BLOCK_KEY_ENCRYPTED_KEY,
        HashAlg::Sha512,
    )
    .unwrap();
    assert_eq!(key, key3);
}

#[test]
fn test_derive_key_128_bit() {
    let salt = [0xAA; 16];
    let key = derive_key(
        "test",
        &salt,
        1,
        128,
        &BLOCK_KEY_ENCRYPTED_KEY,
        HashAlg::Sha256,
    )
    .unwrap();
    assert_eq!(key.len(), 16); // 128 bits = 16 bytes
}

#[test]
fn test_derive_key_sha256() {
    let salt = [0xBB; 16];
    let key_256 = derive_key(
        "test",
        &salt,
        1,
        128,
        &BLOCK_KEY_ENCRYPTED_KEY,
        HashAlg::Sha256,
    )
    .unwrap();
    let key_512 = derive_key(
        "test",
        &salt,
        1,
        128,
        &BLOCK_KEY_ENCRYPTED_KEY,
        HashAlg::Sha512,
    )
    .unwrap();
    // SHA-256 and SHA-512 should produce different keys even with same inputs
    assert_ne!(key_256, key_512);
    assert_eq!(key_256.len(), 16);
    assert_eq!(key_512.len(), 16);
}

// --- AES-CBC tests ---

/// The zero-padding encryptor fills a ragged tail with zeros, adds no block
/// to aligned input, and the no-pad decryptor returns the zero-filled text.
#[test]
fn aes_cbc_zero_pad_roundtrip() {
    let key = [0x60u8; 32];
    let iv = [0x00u8; 16];

    let plaintext = b"Hello AES-256-CBC Test!"; // 23 bytes
    let ciphertext = aes_cbc_encrypt_zero_pad(&key, &iv, plaintext).unwrap();
    assert_eq!(ciphertext.len(), 32);
    let decrypted = aes_cbc_decrypt_no_pad(&key, &iv, &ciphertext).unwrap();
    assert_eq!(&decrypted[..23], plaintext);
    assert_eq!(decrypted[23..], [0u8; 9]);

    for len in [16, 20, 32, 48, 64] {
        let ct = aes_cbc_encrypt_zero_pad(&key, &iv, &vec![0xAB; len]).unwrap();
        assert_eq!(ct.len(), len.next_multiple_of(16), "{len} bytes");
    }
}

#[test]
fn test_aes_cbc_decrypt_no_pad_roundtrip() {
    use cbc::cipher::{BlockModeEncrypt, block_padding::NoPadding as NoPad};
    type Aes256CbcEnc = cbc::Encryptor<Aes256>;

    let key = [0x42u8; 32];
    let iv = [0x01u8; 16];

    // Must be exact multiple of block size (16)
    let plaintext = [0xAA; 32]; // 2 blocks
    let encryptor = Aes256CbcEnc::new_from_slices(&key, &iv).unwrap();
    let ciphertext = encryptor.encrypt_padded_vec::<NoPad>(&plaintext);

    let decrypted = aes_cbc_decrypt_no_pad(&key, &iv, &ciphertext).unwrap();
    assert_eq!(decrypted, plaintext);
}

#[test]
fn test_decrypt_single_segment() {
    // Create a fake encrypted package: 8-byte size + 1 segment of encrypted data
    use cbc::cipher::{BlockModeEncrypt, block_padding::NoPadding as NoPad};

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
    let ciphertext = encryptor.encrypt_padded_vec::<NoPad>(&padded);

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
        key_hash_alg: HashAlg::Sha512,
        encrypted_hmac_key: vec![],
        encrypted_hmac_value: vec![],
        pw_spin_count: 100000,
        pw_salt: vec![],
        pw_key_bits: 256,
        pw_hash_size: 64,
        pw_hash_alg: HashAlg::Sha512,
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
    // Three segments of 4096 + 4096 + 1808 plaintext bytes; the test
    // zero-fills the last one to a full segment before encrypting it.
    use cbc::cipher::{BlockModeEncrypt, block_padding::NoPadding as NoPad};

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
        let ct = encryptor.encrypt_padded_vec::<NoPad>(chunk);
        package.extend_from_slice(&ct);
    }

    let info = AgileEncryptionInfo {
        key_salt: salt.to_vec(),
        key_block_size: 16,
        key_bits: 256,
        key_hash_size: 64,
        key_hash_alg: HashAlg::Sha512,
        encrypted_hmac_key: vec![],
        encrypted_hmac_value: vec![],
        pw_spin_count: 100000,
        pw_salt: vec![],
        pw_key_bits: 256,
        pw_hash_size: 64,
        pw_hash_alg: HashAlg::Sha512,
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
        key_hash_alg: HashAlg::Sha512,
        encrypted_hmac_key: vec![],
        encrypted_hmac_value: vec![],
        pw_spin_count: 100000,
        pw_salt: vec![],
        pw_key_bits: 256,
        pw_hash_size: 64,
        pw_hash_alg: HashAlg::Sha512,
        pw_encrypted_key_value: vec![],
        pw_encrypted_verifier_hash_input: vec![],
        pw_encrypted_verifier_hash_value: vec![],
    };
    let result = decrypt_package(&key, &info, &[0u8; 4]); // too short
    assert!(result.is_err());
    assert!(result.unwrap_err().to_string().contains("too short"));
}

// --- Standard Encryption tests ---

/// MS-OFFCRYPTO 2.3.4.7 known answer, from msoffcrypto-tool
/// `ecma376_standard.py` (MIT): the key of the Word-authored Standard
/// fixture, whose verifier salt is e88266490c5bd1eebd2b4394e3f830ef.
#[test]
fn standard_key_known_answer() {
    let salt = [
        0xe8, 0x82, 0x66, 0x49, 0x0c, 0x5b, 0xd1, 0xee, 0xbd, 0x2b, 0x43, 0x94, 0xe3, 0xf8, 0x30,
        0xef,
    ];
    let key = standard_key("Password1234_", &salt, 128).unwrap();
    assert_eq!(to_hex(&key), "40b13a71f90b966e375408f2d181a1aa");

    // 192 and 256 bits extend the same X1 || X2 stream.
    let k192 = standard_key("Password1234_", &salt, 192).unwrap();
    let k256 = standard_key("Password1234_", &salt, 256).unwrap();
    assert_eq!((k192.len(), k256.len()), (24, 32));
    assert_eq!(k192[..16], key[..]);
    assert_eq!(k256[..24], k192[..]);
}

/// standard_key is deterministic, password- and salt-sensitive, and
/// rejects a key size outside {128, 192, 256} before hashing.
#[test]
fn standard_key_properties() {
    let salt = [0x01u8; 16];
    let key = standard_key("password", &salt, 128).unwrap();
    assert_eq!(key.len(), 16);
    assert_eq!(key, standard_key("password", &salt, 128).unwrap());
    assert_ne!(key, standard_key("other", &salt, 128).unwrap());
    assert_ne!(key, standard_key("password", &[0x02; 16], 128).unwrap());

    let err = standard_key("password", &salt, 512).unwrap_err();
    assert_eq!(err.code(), "PASSWORD_PROTECTED", "{err}");
    assert!(
        err.to_string()
            .contains("unsupported encryption: key size 512 bits"),
        "{err}"
    );
}

#[test]
fn test_standard_verify_password_roundtrip() {
    // Create a known StandardEncryptionInfo by encrypting a verifier with a derived key.
    let password = "test123";
    let salt = [0x42u8; 16];

    // Derive the key the same way verify_password_standard will
    let key = standard_key(password, &salt, 128).unwrap();

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
        alg_id: 0x660E,      // AES-128
        hash_alg_id: 0x8004, // SHA-1
        key_size: 128,
        salt: salt.to_vec(),
        encrypted_verifier,
        verifier_hash_size: 20,
        encrypted_verifier_hash,
    };

    let result = verify_password_standard(password, &info);
    assert!(
        result.is_ok(),
        "verify_password_standard failed: {result:?}"
    );
    let derived = result.unwrap();
    assert_eq!(derived.len(), 16);
    assert_eq!(derived, key);
}

#[test]
fn test_standard_wrong_password() {
    // Same setup as roundtrip but verify with wrong password
    let password = "correct";
    let salt = [0x55u8; 16];

    let key = standard_key(password, &salt, 128).unwrap();

    let verifier = [0xAB; 16];
    let mut hasher = Sha1::new();
    Digest::update(&mut hasher, verifier);
    let verifier_hash = hasher.finalize();

    let encrypted_verifier = aes_ecb_encrypt_no_pad(&key, &verifier).unwrap();

    let mut hash_padded = verifier_hash.to_vec();
    hash_padded.resize(32, 0);
    let encrypted_verifier_hash = aes_ecb_encrypt_no_pad(&key, &hash_padded).unwrap();

    let info = StandardEncryptionInfo {
        alg_id: 0x660E, // AES-128
        hash_alg_id: 0x8004,
        key_size: 128,
        salt: salt.to_vec(),
        encrypted_verifier,
        verifier_hash_size: 20,
        encrypted_verifier_hash,
    };

    let result = verify_password_standard("wrong_password", &info);
    assert!(result.is_err());
    assert!(
        result
            .unwrap_err()
            .to_string()
            .contains("Incorrect password"),
        "Expected 'Incorrect password' error"
    );
}

#[test]
fn test_standard_decrypt_package() {
    // Encrypt-then-decrypt roundtrip for the Standard single-stream AES-ECB
    let key = [0x77u8; 16]; // AES-128
    let salt = [0x33u8; 16]; // not an IV: ECB has none
    let original: Vec<u8> = (0..200u8).collect();

    // Zero-fill to a block boundary for AES-ECB
    let mut padded = original.clone();
    padded.resize(original.len().next_multiple_of(16), 0);
    let ciphertext = aes_ecb_encrypt_no_pad(&key, &padded).unwrap();

    // Build package: 8-byte size + ciphertext
    let mut package = Vec::new();
    package.extend_from_slice(&(200u64).to_le_bytes());
    package.extend_from_slice(&ciphertext);

    let info = StandardEncryptionInfo {
        alg_id: 0x660E, // AES-128
        hash_alg_id: 0x8004,
        key_size: 128,
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
        alg_id: 0x660E, // AES-128
        hash_alg_id: 0x8004,
        key_size: 128,
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
    assert!(
        result2
            .unwrap_err()
            .to_string()
            .contains("no encrypted data")
    );
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
        crate::ole2::cfb::reader::read_stream(&encrypted_ole2, "EncryptionInfo").unwrap();
    let encrypted_package =
        crate::ole2::cfb::reader::read_stream(&encrypted_ole2, "EncryptedPackage").unwrap();

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

// --- Integrity and zero-length hash guards ---

/// Splits an encrypted OLE2 file into its Agile info, the data key unlocked
/// by `password`, and the raw EncryptedPackage stream.
fn agile_parts(ole2: &[u8], password: &str) -> (AgileEncryptionInfo, Zeroizing<Vec<u8>>, Vec<u8>) {
    use crate::ole2::encryption_info::EncryptionInfo;
    let info_stream = crate::ole2::cfb::reader::read_stream(ole2, "EncryptionInfo").unwrap();
    let package = crate::ole2::cfb::reader::read_stream(ole2, "EncryptedPackage").unwrap();
    let EncryptionInfo::Agile(agile) = EncryptionInfo::parse(&info_stream).unwrap() else {
        cold_path();
        unreachable!("expected Agile encryption info");
    };
    let data_key = verify_password_agile(password, &agile).unwrap();
    (*agile, data_key, package)
}

fn hmac_error(result: Result<()>) -> String {
    let err = result.expect_err("HMAC verification must fail");
    assert!(
        matches!(err, ModernXlsxError::PasswordProtected(_)),
        "{err:?}"
    );
    err.to_string()
}

/// `hashSize = 0` fails the integrity check instead of comparing two empty
/// slices and passing vacuously.
#[test]
fn verify_hmac_rejects_zero_hash_size() {
    let ole2 = encrypt_file(b"PK\x03\x04 zero hashSize", "pw").unwrap();
    let (mut agile, data_key, package) = agile_parts(&ole2, "pw");
    verify_hmac(&data_key, &agile, &package).unwrap();

    agile.key_hash_size = 0;
    let msg = hmac_error(verify_hmac(&data_key, &agile, &package));
    assert!(
        msg.contains("HMAC verification failed — file may be corrupted or tampered with."),
        "{msg}"
    );

    let mut info = agile_info_sha512();
    info.key_hash_size = 0;
    assert!(verify_hmac(&[0u8; 32], &info, &[0u8; 64]).is_err());
}

/// A single flipped byte anywhere in the EncryptedPackage stream (size
/// prefix or a ciphertext segment) fails the integrity check.
#[test]
fn verify_hmac_rejects_tampered_package() {
    let payload = vec![0x5Au8; 5000]; // two segments
    let ole2 = encrypt_file(&payload, "pw").unwrap();
    let (agile, data_key, package) = agile_parts(&ole2, "pw");
    verify_hmac(&data_key, &agile, &package).unwrap();

    for index in [0, 8, package.len() - 1] {
        let mut tampered = package.clone();
        tampered[index] ^= 0x01;
        let msg = hmac_error(verify_hmac(&data_key, &agile, &tampered));
        assert!(
            msg.contains("HMAC verification failed — file may be corrupted or tampered with."),
            "byte {index}: {msg}"
        );
    }
}

/// A modern-xlsx 1.2.1-written Agile file passes the integrity check and
/// fails it once tampered with.
#[test]
fn verify_hmac_accepts_legacy_fixture_package() {
    let file = include_bytes!(concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/tests/fixtures/legacy/modern-xlsx-1.2.1-agile.xlsx"
    ));
    let (agile, data_key, package) = agile_parts(file, "Password1234_");
    verify_hmac(&data_key, &agile, &package).unwrap();
    let zip = decrypt_package(&data_key, &agile, &package).unwrap();
    assert!(zip.starts_with(b"PK\x03\x04"));

    let mut tampered = package;
    let last = tampered.len() - 1;
    tampered[last] ^= 0x80;
    hmac_error(verify_hmac(&data_key, &agile, &tampered));
}

/// `passwordKeyEncryptor hashSize = 0` must not turn the verifier compare
/// into an empty-slice match that accepts any password.
#[test]
fn verify_password_agile_rejects_zero_hash_size() {
    let ole2 = encrypt_file(b"PK\x03\x04 zero pw hashSize", "pw").unwrap();
    let (mut agile, _, _) = agile_parts(&ole2, "pw");
    agile.pw_hash_size = 0;
    let err = verify_password_agile("pw", &agile).expect_err("hashSize 0 must fail");
    assert!(
        matches!(&err, ModernXlsxError::PasswordProtected(m) if m == "Incorrect password."),
        "{err:?}"
    );
}

/// `verifierHashSize = 0` must not turn the Standard verifier compare into
/// an empty-slice match that accepts any password.
#[test]
fn verify_password_standard_rejects_zero_hash_size() {
    let password = "test123";
    let salt = [0x42u8; 16];
    let key = standard_key(password, &salt, 128).unwrap();
    let verifier = [0xDE; 16];
    let mut hash_padded = Sha1::digest(verifier).to_vec();
    hash_padded.resize(32, 0);
    let mut info = StandardEncryptionInfo {
        alg_id: 0x660E,      // AES-128
        hash_alg_id: 0x8004, // SHA-1
        key_size: 128,
        salt: salt.to_vec(),
        encrypted_verifier: aes_ecb_encrypt_no_pad(&key, &verifier).unwrap(),
        verifier_hash_size: 20,
        encrypted_verifier_hash: aes_ecb_encrypt_no_pad(&key, &hash_padded).unwrap(),
    };
    verify_password_standard(password, &info).unwrap();

    info.verifier_hash_size = 0;
    let err = verify_password_standard(password, &info).expect_err("hashSize 0 must fail");
    assert!(
        matches!(&err, ModernXlsxError::PasswordProtected(m) if m == "Incorrect password."),
        "{err:?}"
    );
}

#[test]
fn test_encrypt_package_segment_sizes() {
    // Verify the encrypted package has correct structure:
    // 8-byte LE64 header + AES-block-aligned encrypted segments.
    let data_key = [0x55u8; 32];
    let salt = [0xAA; 16];
    let payload = vec![0xBB; 10000]; // 2 full segments + 1 partial

    let encrypted = encrypt_package(&data_key, &salt, &payload, 16, HashAlg::Sha512).unwrap();

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

/// Lowercase hex of `bytes`.
fn to_hex(bytes: &[u8]) -> String {
    use std::fmt::Write as _;
    bytes
        .iter()
        .fold(String::with_capacity(bytes.len() * 2), |mut hex, b| {
            let _ = write!(hex, "{b:02x}");
            hex
        })
}

/// Lowercase hex of SHA-256(`bytes`), for pinning large outputs in a test.
fn sha256_hex(bytes: &[u8]) -> String {
    to_hex(&Sha256::digest(bytes))
}

/// The segment encryptor's output is pinned byte-for-byte: SHA-256 of
/// `encrypt_package` for a 4096-byte input (one full segment) and a 4097-byte
/// input (a full segment plus a 1-byte tail zero-filled to one AES block).
/// Any change to the bytes the segment encryptor writes fails this test.
#[test]
fn encrypt_package_bytes_pinned() {
    let data: Vec<u8> = (0..4097u32).map(|i| i as u8).collect();
    let full = encrypt_package(&[0x11; 32], &[0x22; 16], &data, 16, HashAlg::Sha512).unwrap();
    assert_eq!(full.len(), 8 + 4096 + 16);
    assert_eq!(
        sha256_hex(&full),
        "af3241075ab4b4d4c2376df6d925f2ac7c78736bc111be5f62d9883a03090663"
    );
    let prefix =
        encrypt_package(&[0x11; 32], &[0x22; 16], &data[..4096], 16, HashAlg::Sha512).unwrap();
    assert_eq!(prefix.len(), 8 + 4096);
    assert_eq!(
        sha256_hex(&prefix),
        "b9107779eb2dceb1f19c94cd4618003015f72907adb1524f8adfccf83ba5d13b"
    );
}

/// One H_n pass finalized per block key equals a full `derive_key` per block
/// key, for the three Agile password keys (SHA-512/256-bit at the Excel spin
/// count) and for SHA-1/128-bit at spin 1.
#[test]
fn finalize_matches_derive_key() {
    let salt = [0x5Au8; 16];
    let block_keys = [
        BLOCK_KEY_VERIFIER_INPUT,
        BLOCK_KEY_VERIFIER_VALUE,
        BLOCK_KEY_ENCRYPTED_KEY,
    ];
    for (alg, spin, key_bits) in [(HashAlg::Sha512, 100_000, 256), (HashAlg::Sha1, 1, 128)] {
        let h_n = iterated_hash_for(alg, "pw", &salt, spin);
        for bk in &block_keys {
            assert_eq!(
                finalize_key_for(alg, &h_n, bk, key_bits),
                *derive_key("pw", &salt, spin, key_bits, bk, alg).unwrap(),
                "{alg} spin {spin}"
            );
        }
    }
}

/// `derive_key` output is pinned for SHA-512/256-bit and SHA-1/128-bit at a
/// non-trivial spin count (1000), so any change to the KDF's output fails
/// this test.
#[test]
fn derive_key_bytes_pinned() {
    let salt: Vec<u8> = (1..=16).collect();
    let k512 = derive_key(
        "pw",
        &salt,
        1000,
        256,
        &BLOCK_KEY_VERIFIER_INPUT,
        HashAlg::Sha512,
    )
    .unwrap();
    assert_eq!(
        sha256_hex(&k512),
        "adef6ed75ff04f313dde4efc554053b9a8257e1da43d06592291b4cd423697a8"
    );
    let k1 = derive_key(
        "pw",
        &salt,
        1000,
        128,
        &BLOCK_KEY_ENCRYPTED_KEY,
        HashAlg::Sha1,
    )
    .unwrap();
    assert_eq!(
        sha256_hex(&k1),
        "756550804cf3191c801fb586150f6de895caffefd81537601905b7c84b68734a"
    );
}

/// A ciphertext that is not a whole number of AES blocks must be rejected,
/// whether it is the only segment or the tail after a full one, and a
/// `blockSize` other than 16 must be rejected before any decryption.
#[test]
fn decrypt_package_rejects_unaligned() {
    let key = [0x33u8; 32];
    let info = agile_info_sha512();

    let mut short = 1000u64.to_le_bytes().to_vec();
    short.extend_from_slice(&[0xAA; 15]);
    assert!(decrypt_package(&key, &info, &short).is_err());

    let mut tail = 5000u64.to_le_bytes().to_vec();
    tail.extend_from_slice(&[0xAA; 4096 + 15]);
    assert!(decrypt_package(&key, &info, &tail).is_err());

    let mut aligned = 32u64.to_le_bytes().to_vec();
    aligned.extend_from_slice(&[0xAA; 32]);
    assert_eq!(decrypt_package(&key, &info, &aligned).unwrap().len(), 32);

    for block_size in [8, 32, 256] {
        let mut bad = agile_info_sha512();
        bad.key_block_size = block_size;
        assert!(
            decrypt_package(&key, &bad, &aligned).is_err(),
            "blockSize {block_size} must be rejected"
        );
    }
}

/// RFC 2202 test case 1: HMAC-SHA-1, key 0x0b x 20, data "Hi There".
#[test]
fn compute_hmac_sha1_rfc2202_case1() {
    let mac = compute_hmac(HashAlg::Sha1, &[0x0b; 20], b"Hi There").unwrap();
    assert_eq!(to_hex(&mac), "b617318655057264e28bc0b6fb378c8ef146be00");
}

/// RFC 4231 test case 1: HMAC-SHA-384, key 0x0b x 20, data "Hi There".
#[test]
fn compute_hmac_sha384_rfc4231_case1() {
    let mac = compute_hmac(HashAlg::Sha384, &[0x0b; 20], b"Hi There").unwrap();
    assert_eq!(
        to_hex(&mac),
        "afd03944d84895626b0825f4ab46907f15f9dadbe4101ec682aa034c7cebc59cfaea9ea9076ede7f4af152e8b2fa9cb6"
    );
}

/// Every algorithm yields a `digest_len()`-byte HMAC.
#[test]
fn compute_hmac_len_matches_digest_len() {
    for alg in [
        HashAlg::Sha1,
        HashAlg::Sha256,
        HashAlg::Sha384,
        HashAlg::Sha512,
    ] {
        let mac = compute_hmac(alg, b"key", b"data").unwrap();
        assert_eq!(mac.len(), alg.digest_len(), "{alg}");
    }
}

/// A `blockSize` other than 16 is an Err, never a slice panic, and every
/// algorithm and AES key size yields fields of the digest length rounded up
/// to an AES block (SHA-1: 32 bytes) under the secret key itself.
#[test]
fn compute_and_encrypt_hmac_spec_lengths() {
    for block_size in [0, 8, 20, 32, 64] {
        assert!(
            compute_and_encrypt_hmac(&[0x11; 16], &[0x22; 16], b"pkg", block_size, HashAlg::Sha1)
                .is_err(),
            "blockSize {block_size}"
        );
    }
    for alg in [
        HashAlg::Sha1,
        HashAlg::Sha256,
        HashAlg::Sha384,
        HashAlg::Sha512,
    ] {
        for key_len in [16, 24, 32] {
            let (k, v) =
                compute_and_encrypt_hmac(&vec![0x11; key_len], &[0x22; 16], b"pkg", 16, alg)
                    .unwrap();
            let want = alg.digest_len().next_multiple_of(16);
            assert_eq!((k.len(), v.len()), (want, want), "{alg} {key_len}");
        }
    }
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

// --- Hostile sizes, KDF guards, AES-192 and zeroized keys ---

/// A Standard AES-128 descriptor whose verifier fields are unused.
fn standard_info_aes128() -> StandardEncryptionInfo {
    StandardEncryptionInfo {
        alg_id: 0x660E,
        hash_alg_id: 0x8004,
        key_size: 128,
        salt: vec![0; 16],
        encrypted_verifier: vec![0; 16],
        verifier_hash_size: 20,
        encrypted_verifier_hash: vec![0; 32],
    }
}

/// An EncryptedPackage stream: the LE64 `declared` size, then `payload`.
fn package_with_size(declared: u64, payload: &[u8]) -> Vec<u8> {
    let mut package = declared.to_le_bytes().to_vec();
    package.extend_from_slice(payload);
    package
}

fn assert_security(result: Result<Vec<u8>>, what: &str) {
    let err = result.expect_err(what);
    assert_eq!(err.code(), "SECURITY", "{what}: {err}");
}

/// A declared size above the bytes present is hostile input: SECURITY, with
/// no capacity-overflow panic, for both package layouts. The size converts
/// with `usize::try_from`, never `as`, so 0x1_0000_0010 (16 after a 32-bit
/// truncation) is rejected too.
#[test]
fn decrypt_package_rejects_declared_size_above_payload() {
    let agile = agile_info_sha512();
    let standard = standard_info_aes128();
    for declared in [u64::MAX, 17, 0x1_0000_0010] {
        let package = package_with_size(declared, &[0; 16]);
        assert_security(
            decrypt_package(&[0; 32], &agile, &package),
            &format!("Agile declared {declared:#x}"),
        );
        assert_security(
            decrypt_standard_package(&[0; 16], &standard, &package),
            &format!("Standard declared {declared:#x}"),
        );
    }
    let err = decrypt_package(&[0; 32], &agile, &package_with_size(17, &[0; 16])).unwrap_err();
    assert_eq!(
        err.to_string(),
        "security violation: EncryptedPackage declares 17 bytes but holds 16"
    );

    // A size equal to the payload still decrypts.
    let exact = package_with_size(16, &[0; 16]);
    assert_eq!(decrypt_package(&[0; 32], &agile, &exact).unwrap().len(), 16);
    assert_eq!(
        decrypt_standard_package(&[0; 16], &standard, &exact)
            .unwrap()
            .len(),
        16
    );
}

/// encrypt_package output still decrypts to the input with the same key, for
/// AES-128, AES-192 and AES-256 data keys.
#[test]
fn encrypt_package_roundtrips_through_decrypt_package() {
    let data: Vec<u8> = (0..9000u32).map(|i| (i * 7) as u8).collect();
    for key_len in [16, 24, 32] {
        let key = vec![0x5C; key_len];
        let mut info = agile_info_sha512();
        info.key_salt = vec![0x21; 16];
        let package = encrypt_package(&key, &info.key_salt, &data, 16, HashAlg::Sha512).unwrap();
        assert_eq!(
            decrypt_package(&key, &info, &package).unwrap(),
            data,
            "{key_len}-byte key"
        );
    }
}

/// derive_key rejects a key size outside {128, 192, 256} before allocating
/// the key, and a spinCount above the MS-OFFCRYPTO cap as SECURITY.
#[test]
fn derive_key_rejects_hostile_parameters() {
    let salt = [0u8; 16];
    for key_bits in [0, 64, 4096, 4_294_967_288] {
        let started = std::time::Instant::now();
        let err = derive_key(
            "pw",
            &salt,
            1,
            key_bits,
            &BLOCK_KEY_ENCRYPTED_KEY,
            HashAlg::Sha512,
        )
        .unwrap_err();
        assert_eq!(err.code(), "PASSWORD_PROTECTED", "{err}");
        assert!(
            err.to_string()
                .contains(&format!("unsupported encryption: key size {key_bits} bits")),
            "{err}"
        );
        assert!(started.elapsed() < std::time::Duration::from_secs(1));
    }
    let err = derive_key(
        "pw",
        &salt,
        MAX_SPIN_COUNT + 1,
        256,
        &BLOCK_KEY_ENCRYPTED_KEY,
        HashAlg::Sha512,
    )
    .unwrap_err();
    assert_eq!(err.code(), "SECURITY", "{err}");
    for key_bits in [128, 192, 256] {
        let key = derive_key(
            "pw",
            &salt,
            1,
            key_bits,
            &BLOCK_KEY_ENCRYPTED_KEY,
            HashAlg::Sha512,
        )
        .unwrap();
        assert_eq!(key.len() * 8, key_bits as usize);
    }
}

/// verify_password_agile applies the same guards to a hand-built descriptor.
#[test]
fn verify_password_agile_rejects_hostile_parameters() {
    let mut info = agile_info_sha512();
    info.pw_key_bits = 4_294_967_288;
    let err = verify_password_agile("pw", &info).unwrap_err();
    assert_eq!(err.code(), "PASSWORD_PROTECTED", "{err}");

    let mut info = agile_info_sha512();
    info.pw_spin_count = MAX_SPIN_COUNT + 1;
    let err = verify_password_agile("pw", &info).unwrap_err();
    assert_eq!(err.code(), "SECURITY", "{err}");
}

/// A no-pad CBC encrypt of a length that is not a whole number of AES blocks
/// is an Err for every key size, never a cipher panic; the decrypt side too.
#[test]
fn aes_cbc_no_pad_rejects_unaligned_input() {
    for key_len in [16, 24, 32] {
        let key = vec![0u8; key_len];
        for len in [1, 15, 17, 31] {
            assert!(
                aes_cbc_encrypt_no_pad(&key, &[0; 16], &vec![0; len]).is_err(),
                "encrypt {key_len}-byte key, {len} bytes"
            );
            assert!(
                aes_cbc_decrypt_no_pad(&key, &[0; 16], &vec![0; len]).is_err(),
                "decrypt {key_len}-byte key, {len} bytes"
            );
        }
        assert_eq!(
            aes_cbc_encrypt_no_pad(&key, &[0; 16], &[0; 32])
                .unwrap()
                .len(),
            32
        );
    }
}

/// AES-192 works in every CBC and ECB helper.
#[test]
fn aes192_roundtrips() {
    let key = [0x19u8; 24];
    let iv = [0x02u8; 16];
    let block_aligned: Vec<u8> = (0..48u8).collect();
    let ct = aes_cbc_encrypt_no_pad(&key, &iv, &block_aligned).unwrap();
    assert_ne!(ct, block_aligned);
    assert_eq!(
        aes_cbc_decrypt_no_pad(&key, &iv, &ct).unwrap(),
        block_aligned
    );

    let ragged = b"AES-192 with zero padding";
    let ct = aes_cbc_encrypt_zero_pad(&key, &iv, ragged).unwrap();
    assert_eq!(ct.len(), 32);
    assert_eq!(
        aes_cbc_decrypt_no_pad(&key, &iv, &ct).unwrap()[..ragged.len()],
        ragged[..]
    );

    let ct = aes_ecb_encrypt_no_pad(&key, &block_aligned).unwrap();
    assert_ne!(ct, block_aligned);
    assert_eq!(aes_ecb_decrypt_no_pad(&key, &ct).unwrap(), block_aligned);

    // AES-192-CBC equals the RustCrypto reference for the same key and IV.
    use cbc::cipher::{BlockModeEncrypt, block_padding::NoPadding as NoPad};
    let reference = cbc::Encryptor::<aes::Aes192>::new_from_slices(&key, &iv)
        .unwrap()
        .encrypt_padded_vec::<NoPad>(&block_aligned);
    assert_eq!(
        aes_cbc_encrypt_no_pad(&key, &iv, &block_aligned).unwrap(),
        reference
    );
}

/// Every other key length still names the three supported sizes.
#[test]
fn aes_helpers_name_supported_key_sizes() {
    let key = [0u8; 20];
    for err in [
        aes_cbc_decrypt_no_pad(&key, &[0; 16], &[0; 16]).unwrap_err(),
        aes_cbc_encrypt_zero_pad(&key, &[0; 16], &[0; 16]).unwrap_err(),
        aes_cbc_encrypt_no_pad(&key, &[0; 16], &[0; 16]).unwrap_err(),
        aes_ecb_decrypt_no_pad(&key, &[0; 16]).unwrap_err(),
        aes_ecb_encrypt_no_pad(&key, &[0; 16]).unwrap_err(),
    ] {
        assert!(err.to_string().contains("expected 16, 24 or 32"), "{err}");
    }
}

/// The derived and unlocked keys come back as `Zeroizing<Vec<u8>>`, wiped
/// on drop (a compile-time proof of the return types).
#[test]
fn derived_keys_are_zeroizing() -> Result<()> {
    let ole2 = encrypt_file(b"PK\x03\x04 zeroizing", "pw")?;
    let (agile, _, _) = agile_parts(&ole2, "pw");
    let k: Zeroizing<Vec<u8>> = verify_password_agile("pw", &agile)?;
    assert_eq!(k.len(), 32);
    let d: Zeroizing<Vec<u8>> = derive_key(
        "pw",
        &[0; 16],
        1,
        128,
        &BLOCK_KEY_ENCRYPTED_KEY,
        HashAlg::Sha1,
    )?;
    assert_eq!(d.len(), 16);

    let verifier = [0xDE; 16];
    let mut hash_padded = Sha1::digest(verifier).to_vec();
    hash_padded.resize(32, 0);
    let key: Zeroizing<Vec<u8>> = standard_key("pw", &[0x42; 16], 128)?;
    let info = StandardEncryptionInfo {
        salt: vec![0x42; 16],
        encrypted_verifier: aes_ecb_encrypt_no_pad(&key, &verifier)?,
        encrypted_verifier_hash: aes_ecb_encrypt_no_pad(&key, &hash_padded)?,
        ..standard_info_aes128()
    };
    let s: Zeroizing<Vec<u8>> = verify_password_standard("pw", &info)?;
    assert_eq!(s, key);
    Ok(())
}

// --- MS-OFFCRYPTO conformance against third-party files ---
//
// tests/fixtures/msoffcrypto holds files encrypted by Excel and Word, copied
// from msoffcrypto-tool v6.0.0 (MIT, see its README.md); their password is
// `Password1234_`. tests/fixtures/legacy holds a modern-xlsx 1.2.1 output.

const FIXTURE_PASSWORD: &str = "Password1234_";
const EXCEL_AGILE: &[u8] = include_bytes!(concat!(
    env!("CARGO_MANIFEST_DIR"),
    "/tests/fixtures/msoffcrypto/example_password.xlsx"
));
const EXCEL_AGILE_PLAIN: &[u8] = include_bytes!(concat!(
    env!("CARGO_MANIFEST_DIR"),
    "/tests/fixtures/msoffcrypto/example.xlsx"
));
const WORD_STANDARD: &[u8] = include_bytes!(concat!(
    env!("CARGO_MANIFEST_DIR"),
    "/tests/fixtures/msoffcrypto/ecma376standard_password.docx"
));
const WORD_STANDARD_PLAIN: &[u8] = include_bytes!(concat!(
    env!("CARGO_MANIFEST_DIR"),
    "/tests/fixtures/msoffcrypto/ecma376standard_password_plain.docx"
));
const LEGACY_AGILE: &[u8] = include_bytes!(concat!(
    env!("CARGO_MANIFEST_DIR"),
    "/tests/fixtures/legacy/modern-xlsx-1.2.1-agile.xlsx"
));

/// A Word-authored Standard (AES-128, SHA-1) file decrypts byte-exact:
/// spec key derivation (X1/X2) plus an AES-ECB package.
#[test]
fn decrypt_file_word_standard_fixture() {
    let zip = decrypt_file(WORD_STANDARD, FIXTURE_PASSWORD).unwrap();
    assert_eq!(zip, WORD_STANDARD_PLAIN);
}

/// An Excel-authored Agile (AES-256, SHA-512) file decrypts byte-exact:
/// unpadded verifier fields and the spec dataIntegrity HMAC. A wrong
/// password is PASSWORD_PROTECTED.
#[test]
fn decrypt_file_excel_agile_fixture() {
    let zip = decrypt_file(EXCEL_AGILE, FIXTURE_PASSWORD).unwrap();
    assert_eq!(zip, EXCEL_AGILE_PLAIN);
    let err = decrypt_file(EXCEL_AGILE, "Password1234").unwrap_err();
    assert_eq!(err.code(), "PASSWORD_PROTECTED", "{err}");
    assert!(err.to_string().contains("Incorrect password"), "{err}");
}

/// A file written by modern-xlsx 1.2.1 `encrypt_file(example.xlsx)` (PKCS#7
/// verifier fields, 1.2.x HMAC scheme) still decrypts through the fallback.
#[test]
fn decrypt_file_legacy_1_2_1_fixture() {
    let zip = decrypt_file(LEGACY_AGILE, FIXTURE_PASSWORD).unwrap();
    assert_eq!(zip, EXCEL_AGILE_PLAIN);
}

/// The writer emits the spec field lengths for AES-256/SHA-512 (no PKCS#7
/// block), and its output decrypts.
#[test]
fn encrypt_file_writes_spec_field_lengths() {
    let ole2 = encrypt_file(EXCEL_AGILE_PLAIN, "pw").unwrap();
    let (agile, _, _) = agile_parts(&ole2, "pw");
    assert_eq!(agile.pw_encrypted_key_value.len(), 32);
    assert_eq!(agile.pw_encrypted_verifier_hash_input.len(), 16);
    assert_eq!(agile.pw_encrypted_verifier_hash_value.len(), 64);
    assert_eq!(agile.encrypted_hmac_key.len(), 64);
    assert_eq!(agile.encrypted_hmac_value.len(), 64);
    assert_eq!(decrypt_file(&ole2, "pw").unwrap(), EXCEL_AGILE_PLAIN);
}

/// One flipped ciphertext byte (package offset 24, inside the first
/// segment) fails the integrity check under both HMAC schemes.
#[test]
fn verify_hmac_rejects_flip_at_offset_24() {
    let ole2 = encrypt_file(&[0x42; 3000], "pw").unwrap();
    let (agile, data_key, mut package) = agile_parts(&ole2, "pw");
    package[24] ^= 0x01;
    let err = verify_hmac(&data_key, &agile, &package).unwrap_err();
    assert!(err.to_string().contains("HMAC"), "{err}");
}

/// Our output verifies under the spec HMAC scheme alone and fails the 1.2.x
/// one; the 1.2.1 fixture is the reverse, and the Excel file passes the
/// spec scheme.
#[test]
fn hmac_schemes_split_between_spec_and_legacy_writers() {
    let ole2 = encrypt_file(b"PK\x03\x04 spec hmac", "pw").unwrap();
    let (agile, key, package) = agile_parts(&ole2, "pw");
    verify_hmac_spec(&key, &agile, &package).unwrap();
    assert!(verify_hmac_legacy(&key, &agile, &package).is_err());

    let (agile, key, package) = agile_parts(LEGACY_AGILE, FIXTURE_PASSWORD);
    assert!(verify_hmac_spec(&key, &agile, &package).is_err());
    verify_hmac_legacy(&key, &agile, &package).unwrap();

    let (agile, key, package) = agile_parts(EXCEL_AGILE, FIXTURE_PASSWORD);
    verify_hmac_spec(&key, &agile, &package).unwrap();
}

/// encrypt_file_with round-trips other hash and key sizes through
/// decrypt_file, and declares them in both descriptor elements. SHA-1
/// fields are its 20-byte digest zero-filled to 32 bytes.
#[test]
fn encrypt_file_with_roundtrips_hash_and_key_sizes() {
    let zip = EXCEL_AGILE_PLAIN;
    for (alg, key_bits) in [
        (HashAlg::Sha1, 128),
        (HashAlg::Sha384, 192),
        (HashAlg::Sha256, 256),
    ] {
        let ole2 = encrypt_file_with(zip, "pw", alg, key_bits, 1).unwrap();
        assert_eq!(decrypt_file(&ole2, "pw").unwrap(), zip, "{alg} {key_bits}");

        let (agile, key, _) = agile_parts(&ole2, "pw");
        assert_eq!(key.len() * 8, key_bits as usize);
        assert_eq!((agile.key_hash_alg, agile.pw_hash_alg), (alg, alg));
        assert_eq!((agile.key_bits, agile.pw_key_bits), (key_bits, key_bits));
        let digest = alg.digest_len();
        assert_eq!(agile.key_hash_size as usize, digest);
        assert_eq!(agile.pw_spin_count, 1);
        let padded = digest.next_multiple_of(16);
        assert_eq!(agile.pw_encrypted_verifier_hash_value.len(), padded);
        assert_eq!(agile.encrypted_hmac_key.len(), padded);
        assert_eq!(agile.encrypted_hmac_value.len(), padded);
        assert_eq!(agile.pw_encrypted_verifier_hash_input.len(), 16);
        assert_eq!(
            agile.pw_encrypted_key_value.len(),
            (key_bits as usize / 8).next_multiple_of(16)
        );
        assert!(decrypt_file(&ole2, "wrong").is_err());
    }
}

/// encrypt_file_with checks its parameters before any work.
#[test]
fn encrypt_file_with_rejects_bad_parameters() {
    let err = encrypt_file_with(b"PK", "pw", HashAlg::Sha512, 512, 1).unwrap_err();
    assert_eq!(err.code(), "PASSWORD_PROTECTED", "{err}");
    let err = encrypt_file_with(b"PK", "pw", HashAlg::Sha512, 256, MAX_SPIN_COUNT + 1).unwrap_err();
    assert_eq!(err.code(), "SECURITY", "{err}");
}

/// An EncryptedPackage whose size prefix is u64::MAX, with a
/// recomputed (valid) HMAC so the integrity check passes, is rejected on
/// the decrypt path as SECURITY, without a capacity-overflow panic.
#[test]
fn decrypt_file_rejects_forged_size_with_valid_hmac() {
    let ole2 = encrypt_file_with(b"PK\x03\x04 forged size", "pw", HashAlg::Sha512, 256, 1).unwrap();
    let info_stream = crate::ole2::cfb::reader::read_stream(&ole2, "EncryptionInfo").unwrap();
    let (agile, key, mut package) = agile_parts(&ole2, "pw");
    package[..8].copy_from_slice(&u64::MAX.to_le_bytes());
    let (hmac_key, hmac_value) =
        compute_and_encrypt_hmac(&key, &agile.key_salt, &package, 16, agile.key_hash_alg).unwrap();

    let xml = std::str::from_utf8(&info_stream[8..]).unwrap();
    let old_key = encode_base64(&agile.encrypted_hmac_key);
    let old_value = encode_base64(&agile.encrypted_hmac_value);
    assert!(xml.contains(&old_key) && xml.contains(&old_value));
    let xml = xml
        .replace(&old_key, &encode_base64(&hmac_key))
        .replace(&old_value, &encode_base64(&hmac_value));
    let mut forged_info = info_stream[..8].to_vec();
    forged_info.extend_from_slice(xml.as_bytes());
    let forged = crate::ole2::cfb::writer::write_ole2(&[
        ("EncryptionInfo", &forged_info),
        ("EncryptedPackage", &package),
    ])
    .unwrap();

    let (forged_agile, forged_key, forged_package) = agile_parts(&forged, "pw");
    verify_hmac(&forged_key, &forged_agile, &forged_package).unwrap();
    let err = decrypt_file(&forged, "pw").unwrap_err();
    assert_eq!(err.code(), "SECURITY", "{err}");
}

/// An 8 MiB package needs more than the 109 header FAT slots
/// (DIFAT) and still round-trips.
#[test]
fn encrypt_file_with_roundtrips_8_mib_through_difat() {
    let mut state = 0x9E37_79B9_7F4A_7C15u64;
    let zip: Vec<u8> = (0..8 << 20)
        .map(|_| {
            state ^= state << 13;
            state ^= state >> 7;
            state ^= state << 17;
            state as u8
        })
        .collect();
    let ole2 = encrypt_file_with(&zip, "pw", HashAlg::Sha512, 256, 1).unwrap();
    assert_ne!(
        u32::from_le_bytes(ole2[68..72].try_into().unwrap()),
        0xFFFF_FFFE,
        "the header must chain a DIFAT sector"
    );
    assert!(decrypt_file(&ole2, "pw").unwrap() == zip);
}

// --- Password-encryptor IV and secret-key length (MS-OFFCRYPTO 2.3.4.12, 2.3.4.13) ---

/// Re-seals the password verifier and the secret key of `info` under KEKs of
/// `kek_bits` and the 16-byte IV `iv`, for the password `pw`. The `info` salt,
/// hash and spin count are the ones used to derive H_n.
fn reseal_password_fields(
    info: &mut AgileEncryptionInfo,
    pw: &str,
    iv: &[u8],
    kek_bits: u32,
    data_key: &[u8],
) {
    let alg = info.pw_hash_alg;
    let h_n = iterated_hash_for(alg, pw, &info.pw_salt, info.pw_spin_count);
    let seal = |block_key: &[u8], plain: &[u8]| -> Vec<u8> {
        let kek = finalize_key_for(alg, &h_n, block_key, kek_bits);
        aes_cbc_encrypt_zero_pad(&kek, iv, plain).unwrap()
    };
    let verifier_input = vec![0x5Au8; info.pw_salt.len()];
    let verifier_hash = hash_bytes(alg, &verifier_input);
    info.pw_encrypted_verifier_hash_input = seal(&BLOCK_KEY_VERIFIER_INPUT, &verifier_input);
    info.pw_encrypted_verifier_hash_value = seal(&BLOCK_KEY_VERIFIER_VALUE, &verifier_hash);
    info.pw_encrypted_key_value = seal(&BLOCK_KEY_ENCRYPTED_KEY, data_key);
}

/// A `saltSize` other than 16 reaches the password fields' IV padded with 0x36
/// (shorter) or truncated (longer), as MS-OFFCRYPTO 2.3.4.12 says; the IV was
/// the raw salt, so every password failed with "AES init error".
#[test]
fn verify_password_agile_pads_or_truncates_the_salt_iv() {
    let zip = b"PK\x03\x04 salt iv";
    let ole2 = encrypt_file_with(zip, "pw", HashAlg::Sha512, 256, 1).unwrap();
    let (mut info, data_key, package) = agile_parts(&ole2, "pw");
    for salt_len in [8usize, 20] {
        info.pw_salt = (0..salt_len).map(|i| 0x10 + i as u8).collect();
        let mut iv = [0x36u8; 16];
        let n = salt_len.min(16);
        iv[..n].copy_from_slice(&info.pw_salt[..n]);
        reseal_password_fields(&mut info, "pw", &iv, 256, &data_key);
        let unlocked = verify_password_agile("pw", &info)
            .unwrap_or_else(|e| panic!("saltSize {salt_len}: {e}"));
        assert_eq!(*unlocked, *data_key, "saltSize {salt_len}");
        assert_eq!(decrypt_package(&unlocked, &info, &package).unwrap(), zip);
    }
}

/// The secret key is sized by `keyData` keyBits (2.3.4.13), not by the password
/// encryptor's keyBits, which only sizes the key-encrypting keys.
#[test]
fn verify_password_agile_sizes_the_secret_key_by_key_data_bits() {
    let zip = b"PK\x03\x04 key bits";
    let ole2 = encrypt_file_with(zip, "pw", HashAlg::Sha512, 256, 1).unwrap();
    let (mut info, data_key, package) = agile_parts(&ole2, "pw");
    assert_eq!(data_key.len(), 32);
    let iv = info.pw_salt.clone();
    reseal_password_fields(&mut info, "pw", &iv, 128, &data_key);
    info.pw_key_bits = 128;
    let unlocked = verify_password_agile("pw", &info).unwrap();
    assert_eq!(*unlocked, *data_key);
    verify_hmac(&unlocked, &info, &package).unwrap();
    assert_eq!(decrypt_package(&unlocked, &info, &package).unwrap(), zip);
}

/// A `keyData` keyBits outside {128, 192, 256} is an unsupported key size.
#[test]
fn verify_password_agile_rejects_unsupported_key_data_bits() {
    let ole2 = encrypt_file_with(b"PK\x03\x04", "pw", HashAlg::Sha512, 256, 1).unwrap();
    let (mut info, _, _) = agile_parts(&ole2, "pw");
    info.key_bits = 100;
    let err = verify_password_agile("pw", &info).unwrap_err();
    assert!(
        matches!(err, ModernXlsxError::PasswordProtected(_)),
        "{err:?}"
    );
    assert!(
        err.to_string()
            .contains("unsupported encryption: key size 100 bits"),
        "{err}"
    );
}
