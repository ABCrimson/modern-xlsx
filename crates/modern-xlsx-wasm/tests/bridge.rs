//! wasm-bindgen bridge tests: they drive the exported functions through the real
//! wasm-bindgen glue under `wasm-pack test --node`. Native builds compile this file
//! to an empty test binary.
#![cfg(target_arch = "wasm32")]

use modern_xlsx_core::ooxml::worksheet::{Cell, CellType, Row, WorksheetXml};
use modern_xlsx_core::{SheetData, WorkbookData};
use modern_xlsx_wasm::*;
use wasm_bindgen::{JsCast, JsError, JsValue};
use wasm_bindgen_test::*;

/// Message of the `Error` object a `JsError` turns into on the JS side.
fn msg(e: JsError) -> String {
    JsValue::from(e)
        .dyn_into::<js_sys::Error>()
        .unwrap()
        .message()
        .into()
}

/// A one-sheet workbook whose `S!A1` holds the number 42, as wire JSON.
fn workbook_json() -> String {
    let workbook = WorkbookData {
        sheets: vec![SheetData {
            name: "S".into(),
            worksheet: WorksheetXml {
                rows: vec![Row {
                    index: 1,
                    cells: vec![Cell {
                        reference: "A1".into(),
                        cell_type: CellType::Number,
                        value: Some("42".into()),
                        ..Default::default()
                    }],
                    ..Default::default()
                }],
                ..Default::default()
            },
            ..Default::default()
        }],
        ..Default::default()
    };
    serde_json::to_string(&workbook).unwrap()
}

/// The `value` of `sheets[0]` cell A1 in a `read()` result.
fn a1_value(read_json: &str) -> String {
    let v: serde_json::Value = serde_json::from_str(read_json).unwrap();
    let cells = &v["sheets"][0]["worksheet"]["rows"][0]["cells"];
    let a1 = cells
        .as_array()
        .unwrap()
        .iter()
        .find(|c| c["reference"] == "A1")
        .expect("A1 present");
    a1["value"].as_str().unwrap().to_owned()
}

#[wasm_bindgen_test]
fn write_then_read_round_trips_a_cell() {
    let bytes = write(&workbook_json()).unwrap().to_vec();
    let json = read(&bytes, None).unwrap();
    assert_eq!(a1_value(&json), "42");
}

#[wasm_bindgen_test]
fn validate_returns_a_report_with_issues() {
    let report: serde_json::Value =
        serde_json::from_str(&validate(&workbook_json()).unwrap()).unwrap();
    assert!(report["issues"].is_array());
}

#[wasm_bindgen_test]
fn repair_returns_workbook_report_and_count() {
    let repaired: serde_json::Value =
        serde_json::from_str(&repair(&workbook_json()).unwrap()).unwrap();
    assert!(repaired["repairCount"].is_number());
    assert!(repaired["workbook"].is_object());
}

#[cfg(feature = "encryption")]
#[wasm_bindgen_test]
fn password_round_trip_and_wrong_password() {
    let bytes = write_with_password(&workbook_json(), "pw")
        .unwrap()
        .to_vec();
    let json = read_with_password(&bytes, "pw", None).unwrap();
    assert_eq!(a1_value(&json), "42");

    let err = msg(read_with_password(&bytes, "wrong", None).unwrap_err());
    assert!(err.starts_with("[PASSWORD_PROTECTED]"), "got: {err}");
}

#[wasm_bindgen_test]
fn streaming_writer_lifecycle() {
    let mut writer = StreamingWriter::new();
    writer.start_sheet("S").unwrap();
    writer
        .write_row(r#"[{"value":"1","cellType":"number"}]"#)
        .unwrap();
    let bytes = writer.finish().unwrap().to_vec();
    assert!(!bytes.is_empty());
    read(&bytes, None).unwrap();

    let err = msg(writer.finish().unwrap_err());
    assert_eq!(err, "[INVALID_ARGUMENT] StreamingWriter already finished");
}

#[wasm_bindgen_test]
fn coded_errors_carry_their_code() {
    let err = msg(read(b"not a zip", None).unwrap_err());
    assert!(err.starts_with("[UNRECOGNIZED_FORMAT]"), "got: {err}");

    let mut truncated = b"PK\x03\x04".to_vec();
    truncated.extend_from_slice(&[0; 60]);
    let err = msg(read(&truncated, None).unwrap_err());
    assert!(err.starts_with("[ZIP_READ]"), "got: {err}");
}

#[wasm_bindgen_test]
fn write_rejects_invalid_json() {
    // A serde error crosses the bridge as `ModernXlsxError::Json`, code "JSON".
    let err = msg(write("{not json").unwrap_err());
    assert!(err.starts_with("[JSON] "), "got: {err}");
}

#[wasm_bindgen_test]
fn version_is_the_wasm_crate_semver() {
    assert!(semver::Version::parse(&version()).is_ok());
    assert_eq!(version(), env!("CARGO_PKG_VERSION"));
}
