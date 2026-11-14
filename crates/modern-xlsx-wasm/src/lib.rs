//! wasm-bindgen bridge between the `modern-xlsx-core` engine and the
//! TypeScript package.
//!
//! Exports: `read`, `readWithPassword`, `write`, `writeWithPassword`,
//! `writeBlob`, `validate`, `repair`, `version` and the `StreamingWriter`
//! class. Workbook data crosses the boundary as JSON strings. Errors surface
//! as `JsError` whose message is prefixed with the error code, `[CODE] message`
//! (see `to_js_err`); a spent `StreamingWriter` reports `[INVALID_ARGUMENT]`
//! and a failed `Blob` construction in `writeBlob` reports `[IO_ERROR]`. The
//! `encryption` feature gates `readWithPassword` and `writeWithPassword`.

use js_sys::Uint8Array;
use wasm_bindgen::prelude::*;
use web_sys::{Blob, BlobPropertyBag};

use modern_xlsx_core::zip::reader::ZipSecurityLimits;

/// Parse a JSON workbook string; a serde error becomes `[JSON] ...`.
#[inline]
fn parse_workbook(json: &str) -> Result<modern_xlsx_core::WorkbookData, JsError> {
    serde_json::from_str(json).map_err(|e| to_js_err(e.into()))
}

/// Parse the optional `ReadOptions.limits` JSON into clamped [`ZipSecurityLimits`].
///
/// `None` means the defaults. A malformed or unknown-key object is an
/// `INVALID_ARGUMENT` error; values above a ceiling are clamped, not rejected.
fn parse_limits(limits: Option<&str>) -> Result<ZipSecurityLimits, JsError> {
    match limits {
        None => Ok(ZipSecurityLimits::default()),
        Some(json) => serde_json::from_str::<ZipSecurityLimits>(json)
            .map(ZipSecurityLimits::clamped)
            .map_err(|e| JsError::new(&format!("[INVALID_ARGUMENT] invalid read limits: {e}"))),
    }
}

/// Convert a `ModernXlsxError` to `JsError`, encoding the error code in the
/// message as `"[CODE] human message"` so the TypeScript layer can parse it.
#[inline]
fn to_js_err(e: modern_xlsx_core::ModernXlsxError) -> JsError {
    JsError::new(&e.to_coded_string())
}

/// Install a readable panic hook on module load — debug builds only.
///
/// In `--dev` WASM builds (`debug_assertions` on) Rust panics are routed to
/// `console.error` with the panic message + a JS stack trace, making failures
/// debuggable. The hook is compiled only with `debug_assertions` on `wasm32`.
#[cfg(all(target_arch = "wasm32", debug_assertions))]
#[wasm_bindgen(start)]
pub fn __install_debug_panic_hook() {
    console_error_panic_hook::set_once();
}

/// Read an XLSX file and return parsed workbook data as a JSON string.
///
/// Accepts a `Uint8Array` containing the raw `.xlsx` bytes.
/// Returns a JSON string representing the parsed workbook (sheets, styles, etc.).
/// The caller should use `JSON.parse()` on the JS side to deserialize.
///
/// Uses `read_xlsx_json_with_options` which streams row/cell data directly as JSON during
/// XML parsing, avoiding intermediate per-cell struct and `String` allocations.
///
/// `limits` is an optional JSON object (`ReadOptions.limits`) with the keys
/// `maxDecompressedSize`, `maxCompressionRatio`, `maxEntries` and `maxJsonLength`;
/// `None` keeps the defaults and values above a ceiling are clamped. A malformed
/// object fails with `[INVALID_ARGUMENT]`.
#[wasm_bindgen]
pub fn read(data: &[u8], limits: Option<String>) -> Result<String, JsError> {
    let limits = parse_limits(limits.as_deref())?;
    modern_xlsx_core::reader::read_xlsx_json_with_options(data, &limits).map_err(to_js_err)
}

/// Read an encrypted XLSX file with a password.
///
/// Accepts a `Uint8Array` containing the raw (possibly encrypted) `.xlsx` bytes
/// and a password string. Returns a JSON string representing the parsed workbook.
/// If the file is not encrypted, the password is ignored and reading proceeds normally.
/// `limits` is the same optional JSON object as in `read`.
#[cfg(feature = "encryption")]
#[wasm_bindgen(js_name = readWithPassword)]
pub fn read_with_password(
    data: &[u8],
    password: &str,
    limits: Option<String>,
) -> Result<String, JsError> {
    let limits = parse_limits(limits.as_deref())?;
    modern_xlsx_core::reader::read_xlsx_json_with_password_and_options(data, password, &limits)
        .map_err(to_js_err)
}

/// Write XLSX file bytes from a JSON string describing the workbook.
///
/// Accepts a JSON string (from `JSON.stringify()` on the JS side).
/// Returns a `Uint8Array` containing the resulting `.xlsx` bytes.
///
/// Deserializes the workbook with `serde_json::from_str()`, the same JSON
/// string approach as `read()`.
#[wasm_bindgen]
pub fn write(json: &str) -> Result<Uint8Array, JsError> {
    let workbook = parse_workbook(json)?;
    let bytes = modern_xlsx_core::writer::write_xlsx(&workbook).map_err(to_js_err)?;
    Ok(Uint8Array::new_from_slice(&bytes))
}

/// Write a password-protected XLSX file using Agile Encryption (AES-256-CBC, SHA-512).
///
/// Accepts a JSON string describing the workbook and a password.
/// Returns a `Uint8Array` containing the encrypted OLE2 compound document.
#[cfg(feature = "encryption")]
#[wasm_bindgen(js_name = writeWithPassword)]
pub fn write_with_password(json: &str, password: &str) -> Result<Uint8Array, JsError> {
    let workbook = parse_workbook(json)?;
    let bytes = modern_xlsx_core::writer::write_xlsx_with_password(&workbook, password)
        .map_err(to_js_err)?;
    Ok(Uint8Array::new_from_slice(&bytes))
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
    Blob::new_with_buffer_source_sequence_and_options(&parts, &opts).map_err(|e| {
        let cause = e
            .dyn_ref::<js_sys::Error>()
            .map_or_else(|| format!("{e:?}"), |err| String::from(err.message()));
        JsError::new(&format!("[IO_ERROR] Blob construction failed: {cause}"))
    })
}

/// Validate a workbook and return a JSON report.
///
/// Accepts a JSON string describing the workbook (same format as `write`).
/// Returns a JSON string containing the `ValidationReport`.
#[wasm_bindgen]
pub fn validate(json: &str) -> Result<String, JsError> {
    let workbook = parse_workbook(json)?;
    let report = modern_xlsx_core::validate::validate_workbook(&workbook);
    serde_json::to_string(&report).map_err(|e| to_js_err(e.into()))
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
    serde_json::to_string(&result).map_err(|e| to_js_err(e.into()))
}

/// Returns the Rust engine crate version (`CARGO_PKG_VERSION` of
/// modern-xlsx-wasm, set by the workspace version in Cargo.toml), which is
/// independent of the npm package version (`VERSION` in
/// packages/modern-xlsx/src/exports-shared.ts).
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
/// in memory as a JSON string, `StreamingWriter` renders each row into one
/// reused buffer and deflates it straight into an in-memory ZIP, so the
/// uncompressed worksheet XML is never held whole. Memory is the compressed
/// archive written so far, plus the shared string table, plus one row buffer.
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

/// Error for any call on a spent `StreamingWriter`.
fn finished() -> JsError {
    JsError::new("[INVALID_ARGUMENT] StreamingWriter already finished")
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
    pub fn set_styles_xml(&mut self, xml: String) -> Result<(), JsError> {
        self.inner
            .as_mut()
            .ok_or_else(finished)
            .map(|w| w.set_styles_xml(xml))
    }

    /// Start a new worksheet with the given name.
    #[wasm_bindgen(js_name = startSheet)]
    pub fn start_sheet(&mut self, name: &str) -> Result<(), JsError> {
        self.inner
            .as_mut()
            .ok_or_else(finished)?
            .start_sheet(name)
            .map_err(to_js_err)
    }

    /// Write a row of cells (passed as a JSON string array of StreamingCell objects).
    #[wasm_bindgen(js_name = writeRow)]
    pub fn write_row(&mut self, cells_json: &str) -> Result<(), JsError> {
        let cells: Vec<modern_xlsx_core::streaming_writer::StreamingCell> =
            serde_json::from_str(cells_json).map_err(|e| to_js_err(e.into()))?;
        self.inner
            .as_mut()
            .ok_or_else(finished)?
            .write_row(&cells)
            .map_err(to_js_err)
    }

    /// Finish writing and return the complete XLSX as a Uint8Array.
    ///
    /// Finalizes the archive and returns it. The writer is spent afterwards:
    /// any later call fails; `setStylesXml`, `startSheet`, `finish` and a
    /// `writeRow` with valid JSON fail with `[INVALID_ARGUMENT] StreamingWriter
    /// already finished`.
    pub fn finish(&mut self) -> Result<Uint8Array, JsError> {
        let core = self.inner.take().ok_or_else(finished)?;
        let bytes = core.finish().map_err(to_js_err)?;
        Ok(Uint8Array::new_from_slice(&bytes))
    }
}
