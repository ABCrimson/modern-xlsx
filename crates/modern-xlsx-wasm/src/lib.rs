use js_sys::Uint8Array;
use wasm_bindgen::prelude::*;
use web_sys::{Blob, BlobPropertyBag};

/// Read an XLSX file and return parsed workbook data as a JSON string.
///
/// Accepts a `Uint8Array` containing the raw `.xlsx` bytes.
/// Returns a JSON string representing the parsed workbook (sheets, styles, etc.).
/// The caller should use `JSON.parse()` on the JS side to deserialize.
///
/// Uses `read_xlsx_json` which streams row/cell data directly as JSON during
/// XML parsing, avoiding millions of intermediate struct/String allocations.
/// This is critical for WASM performance where `memory.grow` calls are expensive.
#[wasm_bindgen]
pub fn read(data: &[u8]) -> Result<String, JsError> {
    modern_xlsx_core::reader::read_xlsx_json(data)
        .map_err(|e| JsError::new(&e.to_string()))
}

/// Write XLSX file bytes from a JSON string describing the workbook.
///
/// Accepts a JSON string (from `JSON.stringify()` on the JS side).
/// Returns a `Uint8Array` containing the resulting `.xlsx` bytes.
///
/// Uses `serde_json::from_str()` for fast deserialization, matching the
/// JSON string approach used in `read()`.
#[wasm_bindgen]
pub fn write(json: &str) -> Result<Uint8Array, JsError> {
    let workbook: modern_xlsx_core::WorkbookData =
        serde_json::from_str(json)
            .map_err(|e| JsError::new(&e.to_string()))?;
    let bytes = modern_xlsx_core::writer::write_xlsx(&workbook)
        .map_err(|e| JsError::new(&e.to_string()))?;
    let arr = Uint8Array::new_with_length(bytes.len() as u32);
    arr.copy_from(&bytes);
    Ok(arr)
}

/// Write XLSX and return as a `Blob` for browser download.
///
/// Accepts a JSON string describing the workbook.
/// Returns a `Blob` with MIME type `application/vnd.openxmlformats-officedocument.spreadsheetml.sheet`.
#[wasm_bindgen(js_name = writeBlob)]
pub fn write_blob(json: &str) -> Result<Blob, JsError> {
    let arr = write(json)?;
    let parts = js_sys::Array::new();
    parts.push(&arr.buffer());
    let opts = BlobPropertyBag::new();
    opts.set_type("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    Blob::new_with_buffer_source_sequence_and_options(&parts, &opts)
        .map_err(|e| JsError::new(&format!("{e:?}")))
}

/// Validate a workbook and return a JSON report.
///
/// Accepts a JSON string describing the workbook (same format as `write`).
/// Returns a JSON string containing the `ValidationReport`.
#[wasm_bindgen]
pub fn validate(json: &str) -> Result<String, JsError> {
    let workbook: modern_xlsx_core::WorkbookData =
        serde_json::from_str(json)
            .map_err(|e| JsError::new(&e.to_string()))?;
    let report = modern_xlsx_core::validate::validate_workbook(&workbook);
    serde_json::to_string(&report)
        .map_err(|e| JsError::new(&e.to_string()))
}

/// Validate and auto-repair a workbook. Returns repaired workbook as JSON.
///
/// Accepts a JSON string describing the workbook.
/// Returns a JSON object with `{ workbook, report, repairCount }`.
#[wasm_bindgen]
pub fn repair(json: &str) -> Result<String, JsError> {
    let mut workbook: modern_xlsx_core::WorkbookData =
        serde_json::from_str(json)
            .map_err(|e| JsError::new(&e.to_string()))?;
    let repair_count = modern_xlsx_core::validate::repair_workbook(&mut workbook);
    let report = modern_xlsx_core::validate::validate_workbook(&workbook);

    // Build combined result as JSON.
    let result = serde_json::json!({
        "workbook": workbook,
        "report": report,
        "repairCount": repair_count,
    });
    serde_json::to_string(&result)
        .map_err(|e| JsError::new(&e.to_string()))
}

/// Get the library version.
#[wasm_bindgen]
pub fn version() -> String {
    env!("CARGO_PKG_VERSION").into()
}

#[cfg(test)]
mod tests {
    use super::*;
    use wasm_bindgen_test::*;

    #[wasm_bindgen_test]
    fn test_version() {
        assert!(!version().is_empty());
        assert!(version().split('.').count() == 3);
    }
}
