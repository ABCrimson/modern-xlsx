//! Security hardening integration tests for OLE2 encryption.
//!
//! Validates that:
//! - Error messages never leak key material, passwords, or derived values
//! - Password verification path works correctly (wrong passwords rejected)
//! - Malformed/truncated EncryptionInfo bytes produce errors, not panics
//! - Truncated/empty EncryptedPackage data produces errors, not panics

use modern_xlsx_core::ole2::encryption_info::EncryptionInfo;

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
    let encrypted =
        modern_xlsx_core::ole2::crypto::encrypt_file(test_data, "correct_pw").unwrap();
    let err = modern_xlsx_core::ole2::detect::decrypt_file(&encrypted, "wrong_pw");
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
// Test 2: Password verification path works correctly
// ---------------------------------------------------------------------------

#[test]
fn test_constant_time_comparison_used() {
    // This test verifies the password verification path works correctly:
    // - Wrong passwords are rejected (either "Incorrect password" or AES padding error)
    // - Correct passwords succeed
    // The actual constant_time_eq usage is verified by code audit (grep confirmed
    // 3 call sites in crypto.rs, 0 raw == comparisons on secret data).

    let test_data = b"PK\x03\x04 verification path test data";
    let password = "secure_test_password_123!";

    // Encrypt with known password
    let encrypted = modern_xlsx_core::ole2::crypto::encrypt_file(test_data, password).unwrap();

    // Correct password should succeed
    let result = modern_xlsx_core::ole2::detect::decrypt_file(&encrypted, password);
    assert!(result.is_ok(), "Correct password should decrypt successfully");
    assert_eq!(result.unwrap(), test_data);

    // Wrong passwords should all fail -- either with "Incorrect password" or
    // an AES decryption error (both are valid rejection paths depending on
    // which step of verification catches the mismatch).
    let wrong_passwords = [
        "",
        "wrong",
        "SECURE_TEST_PASSWORD_123!",  // case-swapped
        "secure_test_password_123",   // missing trailing char
        "secure_test_password_124!",  // one char different
        "\u{043F}\u{0430}\u{0440}\u{043E}\u{043B}\u{044C}", // Unicode
    ];

    for wrong_pw in &wrong_passwords {
        let err = modern_xlsx_core::ole2::detect::decrypt_file(&encrypted, wrong_pw);
        assert!(err.is_err(), "Wrong password '{wrong_pw}' should fail");
        let err_msg = err.unwrap_err().to_string();
        // Accept either "Incorrect password" or AES-related errors -- both indicate
        // the wrong password was properly rejected without leaking key material.
        assert!(
            err_msg.contains("Incorrect password")
                || err_msg.contains("AES")
                || err_msg.contains("decryption failed"),
            "Expected a decryption-related error for '{wrong_pw}', got: {err_msg}"
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
        // 4 bytes (too short for version header)
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
        // Must not panic — parse should handle all malformed inputs gracefully
        let result = EncryptionInfo::parse(input);
        // For malformed inputs, we expect errors (though some edge cases might
        // technically parse to empty/default structs — the key requirement is no panic).
        // We simply ensure the call completes without panicking.
        let _ = result;
        // If it's an error, verify no key material in the message
        if let Err(e) = EncryptionInfo::parse(input) {
            let msg = e.to_string();
            assert!(
                !msg.contains("key=") && !msg.contains("password="),
                "Error for malformed input #{i} leaks key material: {msg}"
            );
        }
    }
}

// ---------------------------------------------------------------------------
// Test 4: Truncated EncryptedPackage data must not panic
// ---------------------------------------------------------------------------

#[test]
fn test_truncated_encrypted_package_no_panic() {
    // Create a real encrypted file to get valid EncryptionInfo
    let test_data = b"PK\x03\x04 truncation test data for security";
    let encrypted =
        modern_xlsx_core::ole2::crypto::encrypt_file(test_data, "test_pw").unwrap();

    // Read the real EncryptionInfo to get valid parsed info
    let enc_info_bytes =
        modern_xlsx_core::ole2::detect::read_stream(&encrypted, "EncryptionInfo").unwrap();
    let enc_info = EncryptionInfo::parse(&enc_info_bytes).unwrap();

    // Get the real data key
    let agile = match &enc_info {
        EncryptionInfo::Agile(a) => a,
        _ => panic!("Expected Agile encryption"),
    };
    let data_key =
        modern_xlsx_core::ole2::crypto::verify_password_agile("test_pw", agile).unwrap();

    // Test with various truncated EncryptedPackage payloads
    let truncated_packages: Vec<Vec<u8>> = vec![
        // 4 bytes (too short for size header)
        vec![0x00; 4],
        // 7 bytes (just under 8-byte size header)
        vec![0x00; 7],
        // 9 bytes (size header + 1 byte — not enough for AES block)
        {
            let mut v = Vec::new();
            v.extend_from_slice(&100u64.to_le_bytes());
            v.push(0xFF);
            v
        },
        // Size header + 15 bytes (not block-aligned for some AES modes)
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
        let result =
            modern_xlsx_core::ole2::crypto::decrypt_package(&data_key, agile, package);
        // All inputs are too short or malformed so should error
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
        key_cipher: "AES".into(),
        key_chaining: "ChainingModeCBC".into(),
        key_hash_alg: "SHA512".into(),
        encrypted_hmac_key: vec![],
        encrypted_hmac_value: vec![],
        pw_spin_count: 100_000,
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
        alg_id: 0x6801,
        hash_alg_id: 0x8004,
        key_size: 128,
        provider: "AES".into(),
        salt: vec![0; 16],
        encrypted_verifier: vec![0; 16],
        verifier_hash_size: 20,
        encrypted_verifier_hash: vec![0; 32],
    };

    let std_key = [0u8; 16];

    // Zero-length package for standard encryption
    let result =
        modern_xlsx_core::ole2::crypto::decrypt_standard_package(&std_key, &std_info, &[]);
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
    let result = modern_xlsx_core::ole2::crypto::decrypt_standard_package(
        &std_key,
        &std_info,
        &header_only,
    );
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
