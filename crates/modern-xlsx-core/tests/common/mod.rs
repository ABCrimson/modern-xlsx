//! Shared helpers for the golden-file integration tests.
//!
//! Included as `mod common;` by `tests/package_tests.rs` and, through
//! `#[path]`, by `tests/golden/main.rs`. Every test binary compiles its own
//! copy and uses only a subset of these helpers.
//!
//! Goldens live in `tests/golden/<name>.json`. Run a golden test with
//! `UPDATE_GOLDEN=1` to (re)write the file instead of comparing; a missing
//! golden is a failure. Review every regenerated diff, and never set
//! `UPDATE_GOLDEN` in CI. New goldens serialise only the subtree under test
//! (see [`golden_subtree`]).

#![allow(
    dead_code,
    reason = "each integration-test binary includes this module and uses only a subset of the helpers"
)]

use modern_xlsx_core::ooxml::worksheet::{Cell, CellType, Row, WorksheetXml};
use modern_xlsx_core::reader::read_xlsx;
use modern_xlsx_core::writer::write_xlsx;
use modern_xlsx_core::{SheetData, WorkbookData};
use serde::Serialize;

/// Create a minimal `WorksheetXml` with only the specified rows.
pub fn minimal_worksheet(rows: Vec<Row>) -> WorksheetXml {
    WorksheetXml {
        rows,
        ..Default::default()
    }
}

/// Create a minimal `WorkbookData` wrapping a single worksheet.
pub fn minimal_workbook(name: &str, worksheet: WorksheetXml) -> WorkbookData {
    WorkbookData {
        sheets: vec![SheetData {
            name: name.into(),
            worksheet,
            ..Default::default()
        }],
        ..Default::default()
    }
}

/// Compare `json` with `tests/golden/<name>.json`.
///
/// With `UPDATE_GOLDEN=1` the golden is (re)written instead of compared.
/// A missing golden is a failure.
pub fn assert_golden(json: &str, name: &str) {
    let dir = std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("tests/golden");
    let path = dir.join(format!("{name}.json"));
    if std::env::var("UPDATE_GOLDEN").as_deref() == Ok("1") {
        std::fs::create_dir_all(&dir).expect("create tests/golden");
        std::fs::write(&path, json).expect("write golden file");
        return;
    }
    let expected = std::fs::read_to_string(&path).unwrap_or_else(|e| {
        panic!(
            "golden {} missing or unreadable ({e}); run UPDATE_GOLDEN=1 cargo test -p modern-xlsx-core --test golden and commit it",
            path.display()
        )
    });
    pretty_assertions::assert_eq!(
        json.trim(),
        expected.trim(),
        "Golden file mismatch for {name}"
    );
}

/// Write then read `wb`, returning the re-read workbook.
pub fn roundtrip(wb: &WorkbookData) -> WorkbookData {
    let bytes = write_xlsx(wb).expect("write_xlsx failed");
    read_xlsx(&bytes).expect("read_xlsx failed")
}

/// Write -> read -> compare the whole re-read workbook with the golden `name`.
///
/// Returns the re-read workbook so the caller can compare subtrees with
/// [`assert_same`].
pub fn roundtrip_golden(wb: &WorkbookData, name: &str) -> WorkbookData {
    let wb2 = roundtrip(wb);
    let json = serde_json::to_string_pretty(&wb2).expect("JSON serialization failed");
    assert_golden(&json, name);
    wb2
}

/// Assert two values serialise to the same JSON, field by field.
///
/// The reader normalises some fields (for example `shared_strings: None`
/// reads back as an empty table), so tests compare only the subtrees they set.
pub fn assert_same<T: Serialize>(a: &T, b: &T, what: &str) {
    pretty_assertions::assert_eq!(
        serde_json::to_value(a).expect("serialize left side"),
        serde_json::to_value(b).expect("serialize right side"),
        "{what} changed in round-trip"
    );
}

/// Compare only `value` (the subtree under test) with the golden `name`.
pub fn golden_subtree<T: Serialize>(value: &T, name: &str) {
    let json = serde_json::to_string_pretty(value).expect("JSON serialization failed");
    assert_golden(&json, name);
}

/// Build a one-cell row holding a number. Pass numeric strings only.
pub fn number_row(index: u32, cell_ref: &str, value: &str) -> Row {
    Row {
        index,
        cells: vec![Cell {
            reference: cell_ref.into(),
            cell_type: CellType::Number,
            value: Some(value.into()),
            ..Default::default()
        }],
        ..Default::default()
    }
}

/// Build a shared-string cell. The writer builds the SST from `value`.
pub fn text_cell(cell_ref: &str, text: &str) -> Cell {
    Cell {
        reference: cell_ref.into(),
        cell_type: CellType::SharedString,
        value: Some(text.into()),
        ..Default::default()
    }
}

/// Build a one-cell row holding text as a shared string.
pub fn text_row(index: u32, cell_ref: &str, text: &str) -> Row {
    Row {
        index,
        cells: vec![text_cell(cell_ref, text)],
        ..Default::default()
    }
}
