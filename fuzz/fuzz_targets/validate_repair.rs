#![no_main]
//! Fuzz workbook validation and repair on arbitrary `WorkbookData` JSON.
//! Property: `validate_workbook` and `repair_workbook` must never panic.
use libfuzzer_sys::fuzz_target;
use modern_xlsx_core::WorkbookData;
use modern_xlsx_core::validate::{repair_workbook, validate_workbook};

fuzz_target!(|data: &[u8]| {
    if let Ok(mut wb) = serde_json::from_slice::<WorkbookData>(data) {
        let _ = validate_workbook(&wb);
        let _ = repair_workbook(&mut wb);
        let _ = validate_workbook(&wb);
    }
});
