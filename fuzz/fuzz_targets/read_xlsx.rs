#![no_main]
//! Fuzz the XLSX reader against arbitrary/malformed input.
//! Property: `read_xlsx` must never panic — it returns `Ok` or `Err`, never aborts.
use libfuzzer_sys::fuzz_target;

fuzz_target!(|data: &[u8]| {
    let _ = modern_xlsx_core::reader::read_xlsx(data);
});
