//! Encryption interop corpus: decrypts files written by other implementations
//! (msoffcrypto-tool's fixtures, MIT; modern-xlsx 1.2.1's own output) and, in an
//! ignored test, has msoffcrypto-tool decrypt what this crate encrypts.
#![cfg(feature = "encryption")]

use std::path::{Path, PathBuf};
use std::process::Command;

use modern_xlsx_core::ole2::crypto::{decrypt_file, encrypt_file};
use modern_xlsx_core::reader::read_xlsx;

const PASSWORD: &str = "Password1234_";

fn fixture(dir: &str, name: &str) -> Vec<u8> {
    let path = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("tests/fixtures")
        .join(dir)
        .join(name);
    std::fs::read(&path).unwrap_or_else(|e| panic!("read {}: {e}", path.display()))
}

/// Sheet names and every sheet's rows and cells, as JSON, for a value comparison.
fn sheets_json(xlsx: &[u8]) -> (Vec<String>, serde_json::Value) {
    let wb = read_xlsx(xlsx).expect("read_xlsx");
    let names = wb.sheets.iter().map(|s| s.name.clone()).collect();
    (
        names,
        serde_json::to_value(&wb.sheets).expect("serialize sheets"),
    )
}

#[test]
fn agile_fixture_decrypts_to_the_plain_xlsx() {
    let plain = fixture("msoffcrypto", "example.xlsx");
    let decrypted = decrypt_file(&fixture("msoffcrypto", "example_password.xlsx"), PASSWORD)
        .expect("decrypt Agile fixture");
    assert_eq!(decrypted, plain);
    let (names, cells) = sheets_json(&decrypted);
    let (plain_names, plain_cells) = sheets_json(&plain);
    assert!(!names.is_empty());
    assert_eq!(names, plain_names);
    assert_eq!(cells, plain_cells);
}

#[test]
fn standard_fixture_decrypts_to_the_plain_docx() {
    let decrypted = decrypt_file(
        &fixture("msoffcrypto", "ecma376standard_password.docx"),
        PASSWORD,
    )
    .expect("decrypt Standard fixture");
    assert_eq!(
        decrypted,
        fixture("msoffcrypto", "ecma376standard_password_plain.docx")
    );
}

#[test]
fn legacy_1_2_1_fixture_decrypts_to_example_xlsx() {
    let decrypted = decrypt_file(&fixture("legacy", "modern-xlsx-1.2.1-agile.xlsx"), PASSWORD)
        .expect("decrypt 1.2.1 fixture");
    assert_eq!(decrypted, fixture("msoffcrypto", "example.xlsx"));
}

#[test]
fn wrong_password_is_password_protected() {
    for (dir, name) in [
        ("msoffcrypto", "example_password.xlsx"),
        ("msoffcrypto", "ecma376standard_password.docx"),
        ("legacy", "modern-xlsx-1.2.1-agile.xlsx"),
    ] {
        let err = decrypt_file(&fixture(dir, name), "not the password").expect_err(name);
        assert_eq!(err.code(), "PASSWORD_PROTECTED", "{name}: {err}");
    }
}

/// Deterministic pseudo-random bytes (xorshift64*) of `len` bytes in total, ending
/// in an empty ZIP end-of-central-directory record: msoffcrypto-tool accepts a
/// decrypted payload only if it is a ZIP, and the size pushes the encrypted
/// package into the DIFAT range of the compound file.
fn pseudo_random(len: usize) -> Vec<u8> {
    const EMPTY_ZIP_EOCD: [u8; 22] = [
        0x50, 0x4B, 0x05, 0x06, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0,
    ];
    let len = len - EMPTY_ZIP_EOCD.len();
    let mut state: u64 = 0x9E37_79B9_7F4A_7C15;
    let mut out = Vec::with_capacity(len + 8);
    while out.len() < len {
        state ^= state >> 12;
        state ^= state << 25;
        state ^= state >> 27;
        out.extend_from_slice(&state.wrapping_mul(0x2545_F491_4F6C_DD1D).to_le_bytes());
    }
    out.truncate(len);
    out.extend_from_slice(&EMPTY_ZIP_EOCD);
    out
}

#[test]
#[ignore = "needs msoffcrypto-tool; the CI oracle job runs it with --ignored"]
fn msoffcrypto_oracle_decrypts_our_output() {
    let python = std::env::var("MX_MSOFFCRYPTO_PYTHON")
        .expect("MX_MSOFFCRYPTO_PYTHON must name a Python with msoffcrypto-tool==6.1.0 installed");
    let oracle = Path::new(env!("CARGO_MANIFEST_DIR")).join("tests/interop/oracle.py");
    let dir: PathBuf = std::env::temp_dir().join(format!("mx-oracle-{}", std::process::id()));
    std::fs::create_dir_all(&dir).expect("create temp dir");

    let cases: [(&str, Vec<u8>); 2] = [
        ("example", fixture("msoffcrypto", "example.xlsx")),
        ("random-8mib", pseudo_random(8 * 1024 * 1024)),
    ];
    for (label, plain) in &cases {
        let enc_path = dir.join(format!("{label}.enc"));
        let expected_path = dir.join(format!("{label}.expected"));
        std::fs::write(
            &enc_path,
            encrypt_file(plain, PASSWORD).expect("encrypt_file"),
        )
        .expect("write encrypted");
        std::fs::write(&expected_path, plain).expect("write expected");

        let output = Command::new(&python)
            .arg("-I")
            .arg(&oracle)
            .arg(&enc_path)
            .arg(PASSWORD)
            .arg(&expected_path)
            .output()
            .expect("spawn python");
        assert!(
            output.status.success(),
            "{label}: oracle exited with {}\nstderr: {}",
            output.status,
            String::from_utf8_lossy(&output.stderr)
        );
    }
    let _ = std::fs::remove_dir_all(&dir);
}
