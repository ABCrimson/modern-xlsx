#![no_main]
//! Fuzz the JSON-bridge read path (ZIP + XML parse + JSON serialization).
//! Property: `read_xlsx_json` must never panic on arbitrary input.
use libfuzzer_sys::fuzz_target;

fuzz_target!(|data: &[u8]| {
    let _ = modern_xlsx_core::reader::read_xlsx_json(data);
});
