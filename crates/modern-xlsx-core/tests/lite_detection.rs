//! Lite build (`--no-default-features`): format sniffing and the precise error
//! codes for OLE2 input. The CFB reader, the sniffing and the `EncryptionInfo`
//! parser are crypto-free, so lite names a legacy .xls or an encrypted file
//! instead of reporting `UNRECOGNIZED_FORMAT`; only decryption is feature-gated.
//!
//! Run with `cargo test -p modern-xlsx-core --no-default-features`.

#![cfg(not(feature = "encryption"))]

use modern_xlsx_core::ole2::cfb::writer::write_ole2;
use modern_xlsx_core::reader::{read_xlsx, read_xlsx_json_with_password};
use modern_xlsx_core::writer::write_xlsx;
use modern_xlsx_core::{SheetData, WorkbookData};

const ENCRYPTED: &[u8] = include_bytes!("fixtures/msoffcrypto/example_password.xlsx");

#[test]
fn legacy_xls_reports_legacy_format() {
    let ole2 = write_ole2(&[("Workbook", b"x")]).unwrap();
    let err = read_xlsx(&ole2).unwrap_err();
    assert_eq!(err.code(), "LEGACY_FORMAT", "{err}");
}

#[test]
fn encrypted_with_password_still_reports_password_protected() {
    let err = read_xlsx_json_with_password(ENCRYPTED, "Password1234_").unwrap_err();
    assert_eq!(err.code(), "PASSWORD_PROTECTED", "{err}");
    assert!(err.to_string().contains("lite"), "{err}");
}

#[test]
fn encrypted_without_password_reports_password_protected() {
    let err = read_xlsx(ENCRYPTED).unwrap_err();
    assert_eq!(err.code(), "PASSWORD_PROTECTED", "{err}");
    assert!(err.to_string().contains("lite"), "{err}");
}

#[test]
fn unknown_ole2_and_garbage_report_unrecognized_format() {
    let ole2 = write_ole2(&[("Foo", b"x")]).unwrap();
    assert_eq!(read_xlsx(&ole2).unwrap_err().code(), "UNRECOGNIZED_FORMAT");
    assert_eq!(
        read_xlsx(b"garbage").unwrap_err().code(),
        "UNRECOGNIZED_FORMAT"
    );
}

#[test]
fn plain_xlsx_reads() {
    let wb = WorkbookData {
        sheets: vec![SheetData {
            name: "S".into(),
            ..Default::default()
        }],
        ..Default::default()
    };
    let bytes = write_xlsx(&wb).unwrap();
    assert_eq!(read_xlsx(&bytes).unwrap().sheets.len(), 1);
}
