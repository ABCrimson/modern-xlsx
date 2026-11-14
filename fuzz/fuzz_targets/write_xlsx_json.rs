#![no_main]
//! Fuzz the write path from arbitrary `WorkbookData` JSON, then read the output back.
//! Property: `write_xlsx` and `read_xlsx` must never panic, whatever workbook the JSON describes.
use libfuzzer_sys::fuzz_target;
use modern_xlsx_core::WorkbookData;
use modern_xlsx_core::reader::read_xlsx;
use modern_xlsx_core::writer::write_xlsx;

fuzz_target!(|data: &[u8]| {
    if let Ok(wb) = serde_json::from_slice::<WorkbookData>(data)
        && let Ok(bytes) = write_xlsx(&wb)
    {
        let _ = read_xlsx(&bytes);
    }
});
