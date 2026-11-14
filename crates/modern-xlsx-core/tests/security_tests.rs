//! Security hardening integration tests for OLE2 encryption.
//!
//! Validates that:
//! - Error messages never leak key material, passwords, or derived values
//! - Password verification path works correctly (wrong passwords rejected)
//! - Malformed/truncated EncryptionInfo bytes produce errors, not panics
//! - Truncated/empty EncryptedPackage data produces errors, not panics
#![cfg(feature = "encryption")]

use modern_xlsx_core::ole2::encryption_info::{EncryptionInfo, HashAlg};

// ---------------------------------------------------------------------------
// Test 1: Error messages must not contain key-like patterns
// ---------------------------------------------------------------------------

#[test]
fn test_no_key_in_error_messages() {
    // Patterns that would indicate key material leaking into error messages
    let key_patterns = [
        "key=",
        "Key=",
        "KEY=",
        "password=",
        "Password=",
        "derived=",
        "hmac_key=",
        "data_key=",
        "secret=",
    ];

    // Collect error messages from various failure modes
    let mut error_messages: Vec<String> = Vec::new();

    // 1. Wrong password on a real encrypted file (encrypt then try wrong password)
    let test_data = b"PK\x03\x04 test content for security tests";
    let encrypted = modern_xlsx_core::ole2::crypto::encrypt_file(test_data, "correct_pw").unwrap();
    let err = modern_xlsx_core::ole2::crypto::decrypt_file(&encrypted, "wrong_pw");
    assert!(err.is_err());
    error_messages.push(err.unwrap_err().to_string());

    // 2. Corrupted EncryptionInfo stream (1 byte)
    let err = EncryptionInfo::parse(&[0xFF]);
    assert!(err.is_err());
    error_messages.push(err.unwrap_err().to_string());

    // 3. Corrupted EncryptionInfo stream (4 bytes)
    let err = EncryptionInfo::parse(&[0x00, 0x00, 0x00, 0x00]);
    assert!(err.is_err());
    error_messages.push(err.unwrap_err().to_string());

    // 4. Unsupported version
    let mut bad_version = vec![0u8; 8];
    bad_version[0..2].copy_from_slice(&99u16.to_le_bytes());
    bad_version[2..4].copy_from_slice(&99u16.to_le_bytes());
    let err = EncryptionInfo::parse(&bad_version);
    assert!(err.is_err());
    error_messages.push(err.unwrap_err().to_string());

    // 5. Truncated standard encryption
    let mut truncated_std = Vec::new();
    truncated_std.extend_from_slice(&4u16.to_le_bytes()); // major
    truncated_std.extend_from_slice(&2u16.to_le_bytes()); // minor
    truncated_std.extend_from_slice(&0u32.to_le_bytes()); // flags
    truncated_std.extend_from_slice(&10u32.to_le_bytes()); // headerSize (too small for verifier)
    let err = EncryptionInfo::parse(&truncated_std);
    assert!(err.is_err());
    error_messages.push(err.unwrap_err().to_string());

    // Scan ALL error messages for key-like patterns
    for msg in &error_messages {
        for pattern in &key_patterns {
            assert!(
                !msg.contains(pattern),
                "Error message contains key-like pattern '{pattern}': {msg}"
            );
        }
        // Check for suspicious long hex sequences that might be key material.
        // A 16-byte key would be 32 hex chars; version numbers are short.
        let mut max_hex_run = 0usize;
        let mut current_run = 0usize;
        for c in msg.chars() {
            if c.is_ascii_hexdigit() {
                current_run += 1;
                max_hex_run = max_hex_run.max(current_run);
            } else {
                current_run = 0;
            }
        }
        assert!(
            max_hex_run < 32,
            "Error message contains suspicious hex sequence (length {max_hex_run}): {msg}"
        );
    }
}

// ---------------------------------------------------------------------------
// Test 2: Wrong passwords are rejected, the correct one is accepted
// ---------------------------------------------------------------------------

#[test]
fn test_wrong_password_rejected_correct_accepted() {
    // Wrong passwords are rejected with "Incorrect password" (the verifier
    // fields are decrypted without unpadding, so no AES padding error can
    // surface first) and the correct password succeeds. This test checks
    // behaviour only; it does not measure timing.

    let test_data = b"PK\x03\x04 verification path test data";
    let password = "secure_test_password_123!";

    // Encrypt with known password
    let encrypted = modern_xlsx_core::ole2::crypto::encrypt_file(test_data, password).unwrap();

    // Correct password should succeed
    let result = modern_xlsx_core::ole2::crypto::decrypt_file(&encrypted, password);
    assert!(
        result.is_ok(),
        "Correct password should decrypt successfully"
    );
    assert_eq!(result.unwrap(), test_data);

    // Wrong passwords should all fail with "Incorrect password".
    let wrong_passwords = [
        "",
        "wrong",
        "SECURE_TEST_PASSWORD_123!", // case-swapped
        "secure_test_password_123",  // missing trailing char
        "secure_test_password_124!", // one char different
        "\u{043F}\u{0430}\u{0440}\u{043E}\u{043B}\u{044C}", // Unicode
    ];

    for wrong_pw in &wrong_passwords {
        let err = modern_xlsx_core::ole2::crypto::decrypt_file(&encrypted, wrong_pw);
        assert!(err.is_err(), "Wrong password '{wrong_pw}' should fail");
        let err_msg = err.unwrap_err().to_string();
        assert!(
            err_msg.contains("Incorrect password"),
            "Expected \"Incorrect password\" for '{wrong_pw}', got: {err_msg}"
        );
    }
}

// ---------------------------------------------------------------------------
// Test 3: Malformed EncryptionInfo bytes must not panic
// ---------------------------------------------------------------------------

#[test]
fn test_malformed_encryption_info_no_panic() {
    let test_inputs: Vec<Vec<u8>> = vec![
        // 1 byte
        vec![0x00],
        // 4 bytes: a version with no flags field (the header is 8 bytes)
        vec![0x04, 0x00, 0x04, 0x00],
        // 7 bytes (just under minimum)
        vec![0x04, 0x00, 0x04, 0x00, 0x40, 0x00, 0x00],
        // 8 bytes with valid Standard header but no data after flags
        {
            let mut v = Vec::new();
            v.extend_from_slice(&4u16.to_le_bytes());
            v.extend_from_slice(&2u16.to_le_bytes());
            v.extend_from_slice(&0x24u32.to_le_bytes());
            v
        },
        // 100 random-ish bytes
        (0..100u8).collect(),
        // All 0xFF bytes (50 bytes)
        vec![0xFF; 50],
        // Valid version header + garbage XML (Agile)
        {
            let mut v = Vec::new();
            v.extend_from_slice(&4u16.to_le_bytes());
            v.extend_from_slice(&4u16.to_le_bytes());
            v.extend_from_slice(&0x40u32.to_le_bytes());
            v.extend_from_slice(b"<<<garbage>>not xml at all!!!");
            v
        },
        // Valid standard header + truncated header data
        {
            let mut v = Vec::new();
            v.extend_from_slice(&4u16.to_le_bytes());
            v.extend_from_slice(&2u16.to_le_bytes());
            v.extend_from_slice(&0x24u32.to_le_bytes());
            v.extend_from_slice(&100u32.to_le_bytes()); // headerSize = 100 but not enough data
            v
        },
        // Empty input
        vec![],
    ];

    for (i, input) in test_inputs.iter().enumerate() {
        let Err(err) = EncryptionInfo::parse(input) else {
            panic!("malformed input #{i} must be rejected")
        };
        assert_eq!(err.code(), "PASSWORD_PROTECTED", "#{i}: {err}");
        let msg = err.to_string();
        assert!(
            !msg.contains("key=") && !msg.contains("password="),
            "input #{i} leaks key material: {msg}"
        );
    }
}

// ---------------------------------------------------------------------------
// Test 4: Truncated EncryptedPackage data must not panic
// ---------------------------------------------------------------------------

#[test]
fn test_truncated_encrypted_package_no_panic() {
    // Create a real encrypted file to get valid EncryptionInfo
    let test_data = b"PK\x03\x04 truncation test data for security";
    let encrypted = modern_xlsx_core::ole2::crypto::encrypt_file(test_data, "test_pw").unwrap();

    // Read the real EncryptionInfo to get valid parsed info
    let enc_info_bytes =
        modern_xlsx_core::ole2::cfb::reader::read_stream(&encrypted, "EncryptionInfo").unwrap();
    let enc_info = EncryptionInfo::parse(&enc_info_bytes).unwrap();

    // Get the real data key
    let agile = match &enc_info {
        EncryptionInfo::Agile(a) => a,
        other => unreachable!("Expected Agile encryption, got {other:?}"),
    };
    let data_key = modern_xlsx_core::ole2::crypto::verify_password_agile("test_pw", agile).unwrap();

    // Test with various truncated EncryptedPackage payloads
    let truncated_packages: Vec<Vec<u8>> = vec![
        // 4 bytes (too short for size header)
        vec![0x00; 4],
        // 7 bytes (just under 8-byte size header)
        vec![0x00; 7],
        // 9 bytes: a size header declaring 100 bytes, then 1 byte of payload
        {
            let mut v = Vec::new();
            v.extend_from_slice(&100u64.to_le_bytes());
            v.push(0xFF);
            v
        },
        // A size header declaring 1000 bytes, then 15 bytes of payload
        {
            let mut v = Vec::new();
            v.extend_from_slice(&1000u64.to_le_bytes());
            v.extend_from_slice(&[0xAA; 15]);
            v
        },
        // Zero-length
        vec![],
    ];

    for (i, package) in truncated_packages.iter().enumerate() {
        // decrypt_package must not panic — it should return an error
        let result = modern_xlsx_core::ole2::crypto::decrypt_package(&data_key, agile, package);
        // Each input is shorter than 8 bytes or declares more bytes than it holds, so each errors
        assert!(
            result.is_err(),
            "Truncated package #{i} (len={}) should produce an error",
            package.len()
        );
        // Error message should not contain key material
        let err_msg = result.unwrap_err().to_string();
        assert!(
            !err_msg.contains("key="),
            "Error for truncated package #{i} should not leak key material: {err_msg}"
        );
    }
}

// ---------------------------------------------------------------------------
// Test 5: Empty EncryptedPackage must produce a clear error
// ---------------------------------------------------------------------------

#[test]
fn test_empty_encrypted_package_no_panic() {
    // Use a minimal AgileEncryptionInfo for testing
    let info = modern_xlsx_core::ole2::encryption_info::AgileEncryptionInfo {
        key_salt: vec![0; 16],
        key_block_size: 16,
        key_bits: 256,
        key_hash_size: 64,
        key_hash_alg: HashAlg::Sha512,
        encrypted_hmac_key: vec![],
        encrypted_hmac_value: vec![],
        pw_spin_count: 100_000,
        pw_salt: vec![],
        pw_key_bits: 256,
        pw_hash_size: 64,
        pw_hash_alg: HashAlg::Sha512,
        pw_encrypted_key_value: vec![],
        pw_encrypted_verifier_hash_input: vec![],
        pw_encrypted_verifier_hash_value: vec![],
    };

    let key = [0u8; 32];

    // Zero-length package
    let result = modern_xlsx_core::ole2::crypto::decrypt_package(&key, &info, &[]);
    assert!(result.is_err(), "Empty package should return an error");
    let err_msg = result.unwrap_err().to_string();
    assert!(
        err_msg.contains("too short"),
        "Empty package error should mention 'too short': {err_msg}"
    );

    // Also test with StandardEncryptionInfo
    let std_info = modern_xlsx_core::ole2::encryption_info::StandardEncryptionInfo {
        alg_id: 0x660E, // AES-128
        hash_alg_id: 0x8004,
        key_size: 128,
        salt: vec![0; 16],
        encrypted_verifier: vec![0; 16],
        verifier_hash_size: 20,
        encrypted_verifier_hash: vec![0; 32],
    };

    let std_key = [0u8; 16];

    // Zero-length package for standard encryption
    let result = modern_xlsx_core::ole2::crypto::decrypt_standard_package(&std_key, &std_info, &[]);
    assert!(
        result.is_err(),
        "Empty standard package should return an error"
    );
    let err_msg = result.unwrap_err().to_string();
    assert!(
        err_msg.contains("too short"),
        "Empty standard package error should mention 'too short': {err_msg}"
    );

    // 8-byte header with no encrypted data (standard)
    let mut header_only = Vec::new();
    header_only.extend_from_slice(&100u64.to_le_bytes());
    let result =
        modern_xlsx_core::ole2::crypto::decrypt_standard_package(&std_key, &std_info, &header_only);
    assert!(
        result.is_err(),
        "Header-only standard package should return an error"
    );
    let err_msg = result.unwrap_err().to_string();
    assert!(
        err_msg.contains("no encrypted data"),
        "Header-only standard package should mention 'no encrypted data': {err_msg}"
    );
}

// ---------------------------------------------------------------------------
// Test 6: Malformed OLE2 directory must not panic (fuzzer regression)
// ---------------------------------------------------------------------------

/// Regression for a fuzzer-found panic (`index out of bounds: len 128, index
/// 128`) in `ole2::cfb::reader::read_directory`: a directory entry whose name-length
/// field (offset 64-65) exceeds the 64-byte name field caused the UTF-16 decode
/// loop to index past the 128-byte entry. The reader must return an error, never
/// panic, on a compound file with no decryptable payload.
#[test]
fn read_xlsx_does_not_panic_on_malformed_ole2_directory() {
    const OLE2_MAGIC: [u8; 8] = [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1];
    let crash = include_bytes!("fixtures/fuzz_crash_ole2.bin");
    assert!(
        crash.starts_with(&OLE2_MAGIC),
        "fixture must stay an OLE2 file"
    );
    assert_eq!(
        modern_xlsx_core::reader::read_xlsx(crash)
            .unwrap_err()
            .code(),
        "UNRECOGNIZED_FORMAT"
    );

    // Also exercise the boundary directly: a minimal 512-byte buffer carrying only
    // the OLE2/CFB magic must be rejected, not read and not a panic.
    let mut magic_only = vec![0u8; 512];
    magic_only[..8].copy_from_slice(&OLE2_MAGIC);
    assert_eq!(
        modern_xlsx_core::reader::read_xlsx(&magic_only)
            .unwrap_err()
            .code(),
        "UNRECOGNIZED_FORMAT"
    );
}
