use wasm_bindgen::prelude::*;

/// Read an XLSX file and return parsed workbook data as a JS value.
///
/// Accepts a `Uint8Array` containing the raw `.xlsx` bytes.
/// Returns a JS object representing the parsed workbook (sheets, styles, etc.).
#[wasm_bindgen]
pub fn read(data: &[u8]) -> Result<JsValue, JsError> {
    let workbook = ironsheet_core::reader::read_xlsx(data)
        .map_err(|e| JsError::new(&e.to_string()))?;
    serde_wasm_bindgen::to_value(&workbook)
        .map_err(|e| JsError::new(&e.to_string()))
}

/// Write XLSX file bytes from a JS workbook object.
///
/// Accepts a JS object describing the workbook to write.
/// Returns a `Uint8Array` containing the resulting `.xlsx` bytes.
#[wasm_bindgen]
pub fn write(val: JsValue) -> Result<Vec<u8>, JsError> {
    let workbook: ironsheet_core::writer::WorkbookData =
        serde_wasm_bindgen::from_value(val)
            .map_err(|e| JsError::new(&e.to_string()))?;
    ironsheet_core::writer::write_xlsx(&workbook)
        .map_err(|e| JsError::new(&e.to_string()))
}

/// Get the library version.
#[wasm_bindgen]
pub fn version() -> String {
    env!("CARGO_PKG_VERSION").to_string()
}

#[cfg(test)]
mod tests {
    use super::*;
    use wasm_bindgen_test::*;

    #[wasm_bindgen_test]
    fn test_version() {
        assert_eq!(version(), "0.1.0");
    }
}
