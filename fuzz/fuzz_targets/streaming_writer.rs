#![no_main]
//! Fuzz the streaming writer with arbitrary rows of `StreamingCell` JSON.
//! Property: `StreamingWriterCore` must never panic; a rejected row stops the feed and the sheet is still finished.
use libfuzzer_sys::fuzz_target;
use modern_xlsx_core::streaming_writer::{StreamingCell, StreamingWriterCore};

fuzz_target!(|data: &[u8]| {
    let Ok(rows) = serde_json::from_slice::<Vec<Vec<StreamingCell>>>(data) else {
        return;
    };
    let mut writer = StreamingWriterCore::new();
    let _ = writer.start_sheet("S");
    for row in &rows {
        if writer.write_row(row).is_err() {
            break;
        }
    }
    let _ = writer.finish();
});
