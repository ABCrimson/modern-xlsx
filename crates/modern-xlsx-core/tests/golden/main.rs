//! Golden-file integration tests, one test binary (`--test golden`).
//!
//! Each module builds workbooks, writes and re-reads them, and compares the
//! subtree under test with `tests/golden/<name>.json` through the shared
//! helpers in `tests/common/mod.rs`.

#[path = "../common/mod.rs"]
mod common;

mod cells;
mod charts;
mod parts;
mod pivot;
mod roundtrip;
mod sst;
mod workbook;
