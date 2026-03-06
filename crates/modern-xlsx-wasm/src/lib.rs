use js_sys::Uint8Array;
use wasm_bindgen::prelude::*;
use web_sys::{Blob, BlobPropertyBag};

/// Parse a JSON workbook string, mapping serde errors to `JsError`.
#[inline]
fn parse_workbook(json: &str) -> Result<modern_xlsx_core::WorkbookData, JsError> {
    serde_json::from_str(json).map_err(|e| JsError::new(&format!("JSON parse error: {e}")))
}

/// Convert a `ModernXlsxError` to `JsError`, encoding the error code in the
/// message as `"[CODE] human message"` so the TypeScript layer can parse it.
#[inline]
fn to_js_err(e: modern_xlsx_core::ModernXlsxError) -> JsError {
    JsError::new(&e.to_coded_string())
}

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
    modern_xlsx_core::reader::read_xlsx_json(data).map_err(to_js_err)
}

/// Read an encrypted XLSX file with a password.
///
/// Accepts a `Uint8Array` containing the raw (possibly encrypted) `.xlsx` bytes
/// and a password string. Returns a JSON string representing the parsed workbook.
/// If the file is not encrypted, the password is ignored and reading proceeds normally.
#[cfg(feature = "encryption")]
#[wasm_bindgen(js_name = readWithPassword)]
pub fn read_with_password(data: &[u8], password: &str) -> Result<String, JsError> {
    modern_xlsx_core::reader::read_xlsx_json_with_password(data, password).map_err(to_js_err)
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
    let workbook = parse_workbook(json)?;
    let bytes = modern_xlsx_core::writer::write_xlsx(&workbook).map_err(to_js_err)?;
    let arr = Uint8Array::new_with_length(bytes.len() as u32);
    arr.copy_from(&bytes);
    Ok(arr)
}

/// Write a password-protected XLSX file using Agile Encryption (AES-256-CBC, SHA-512).
///
/// Accepts a JSON string describing the workbook and a password.
/// Returns a `Uint8Array` containing the encrypted OLE2 compound document.
#[cfg(feature = "encryption")]
#[wasm_bindgen(js_name = writeWithPassword)]
pub fn write_with_password(json: &str, password: &str) -> Result<Uint8Array, JsError> {
    let workbook = parse_workbook(json)?;
    let bytes =
        modern_xlsx_core::writer::write_xlsx_with_password(&workbook, password).map_err(to_js_err)?;
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
    let workbook = parse_workbook(json)?;
    let report = modern_xlsx_core::validate::validate_workbook(&workbook);
    serde_json::to_string(&report)
        .map_err(|e| JsError::new(&format!("[XML_PARSE] Failed to serialize validation report: {e}")))
}

/// Validate and auto-repair a workbook. Returns repaired workbook as JSON.
///
/// Accepts a JSON string describing the workbook.
/// Returns a JSON object with `{ workbook, report, repairCount }`.
#[wasm_bindgen]
pub fn repair(json: &str) -> Result<String, JsError> {
    let mut workbook = parse_workbook(json)?;
    let repair_count = modern_xlsx_core::validate::repair_workbook(&mut workbook);
    let report = modern_xlsx_core::validate::validate_workbook(&workbook);

    // Build combined result as JSON.
    let result = serde_json::json!({
        "workbook": workbook,
        "report": report,
        "repairCount": repair_count,
    });
    serde_json::to_string(&result)
        .map_err(|e| JsError::new(&format!("[XML_PARSE] Failed to serialize repair result: {e}")))
}

/// Get the library version.
#[wasm_bindgen]
pub fn version() -> String {
    env!("CARGO_PKG_VERSION").into()
}

// ---------------------------------------------------------------------------
// Streaming Writer
// ---------------------------------------------------------------------------

/// Streaming XLSX writer that writes rows directly to ZIP entries.
///
/// Unlike the standard `write()` function which requires the entire workbook
/// in memory as a JSON string, `StreamingWriter` writes worksheet rows
/// incrementally — peak memory is proportional to the number of unique
/// strings (the SST), not the total row count.
///
/// Usage from JavaScript:
/// ```js
/// const writer = new StreamingWriter();
/// writer.startSheet("Sheet1");
/// writer.writeRow(JSON.stringify([
///   { value: "Hello", cellType: "sharedString" },
///   { value: "42", cellType: "number" },
/// ]));
/// const xlsx = writer.finish(); // Uint8Array
/// ```
#[wasm_bindgen]
pub struct StreamingWriter {
    inner: Option<modern_xlsx_core::streaming_writer::StreamingWriterCore>,
}

impl Default for StreamingWriter {
    fn default() -> Self {
        Self::new()
    }
}

#[wasm_bindgen]
impl StreamingWriter {
    /// Create a new streaming XLSX writer.
    #[wasm_bindgen(constructor)]
    pub fn new() -> Self {
        Self {
            inner: Some(modern_xlsx_core::streaming_writer::StreamingWriterCore::new()),
        }
    }

    /// Set custom styles XML (the complete xl/styles.xml body).
    #[wasm_bindgen(js_name = setStylesXml)]
    pub fn set_styles_xml(&mut self, xml: &str) -> Result<(), JsError> {
        self.inner
            .as_mut()
            .ok_or_else(|| JsError::new("StreamingWriter already finished"))
            .map(|w| w.set_styles_xml(xml.to_string()))
    }

    /// Start a new worksheet with the given name.
    #[wasm_bindgen(js_name = startSheet)]
    pub fn start_sheet(&mut self, name: &str) -> Result<(), JsError> {
        self.inner
            .as_mut()
            .ok_or_else(|| JsError::new("StreamingWriter already finished"))?
            .start_sheet(name)
            .map_err(to_js_err)
    }

    /// Write a row of cells (passed as a JSON string array of StreamingCell objects).
    #[wasm_bindgen(js_name = writeRow)]
    pub fn write_row(&mut self, cells_json: &str) -> Result<(), JsError> {
        let cells: Vec<modern_xlsx_core::streaming_writer::StreamingCell> =
            serde_json::from_str(cells_json)
                .map_err(|e| JsError::new(&format!("JSON parse error: {e}")))?;
        self.inner
            .as_mut()
            .ok_or_else(|| JsError::new("StreamingWriter already finished"))?
            .write_row(&cells)
            .map_err(to_js_err)
    }

    /// Finish writing and return the complete XLSX as a Uint8Array.
    ///
    /// Consumes the writer — calling any method after `finish()` will error.
    pub fn finish(&mut self) -> Result<Uint8Array, JsError> {
        let core = self
            .inner
            .take()
            .ok_or_else(|| JsError::new("StreamingWriter already finished"))?;
        let bytes = core.finish().map_err(to_js_err)?;
        let arr = Uint8Array::new_with_length(bytes.len() as u32);
        arr.copy_from(&bytes);
        Ok(arr)
    }
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
