//! Performance benchmark tests for large workbooks.
//!
//! These tests measure write and roundtrip performance for workbooks with
//! 10K and 100K rows.  They use `std::time::Instant` for wall-clock timing
//! and print results via `println!` so that `cargo test -- --nocapture`
//! shows them.  No correctness assertions are made on timing; the goal is
//! to have a quick regression check that can be run locally.

use std::time::Instant;

use modern_xlsx_core::dates::DateSystem;
use modern_xlsx_core::ooxml::styles::Styles;
use modern_xlsx_core::ooxml::worksheet::{Cell, CellType, Row, WorksheetXml};
use modern_xlsx_core::reader::read_xlsx;
use modern_xlsx_core::writer::write_xlsx;
use modern_xlsx_core::{SheetData, WorkbookData};

/// Convert a 1-based column number to an Excel column letter (A, B, …, Z, AA, …).
fn col_letter(col: u32) -> String {
    let mut result = String::new();
    let mut n = col;
    while n > 0 {
        n -= 1;
        result.insert(0, (b'A' + (n % 26) as u8) as char);
        n /= 26;
    }
    result
}

/// Build a workbook with `row_count` rows and 5 columns each.
///
/// Columns A, C, E are numbers; columns B, D are shared strings.
fn make_workbook(row_count: usize) -> WorkbookData {
    let mut rows = Vec::with_capacity(row_count);
    for r in 0..row_count {
        let row_num = (r + 1) as u32;
        let cells = vec![
            Cell {
                reference: format!("A{row_num}"),
                cell_type: CellType::Number,
                value: Some(row_num.to_string()),
                ..Default::default()
            },
            Cell {
                reference: format!("B{row_num}"),
                cell_type: CellType::SharedString,
                value: Some(format!("Name_{}", r % 500)),
                ..Default::default()
            },
            Cell {
                reference: format!("C{row_num}"),
                cell_type: CellType::Number,
                value: Some(format!("{:.2}", r as f64 * 1.5)),
                ..Default::default()
            },
            Cell {
                reference: format!("D{row_num}"),
                cell_type: CellType::SharedString,
                value: Some(format!("Category_{}", r % 20)),
                ..Default::default()
            },
            Cell {
                reference: format!("E{row_num}"),
                cell_type: CellType::Number,
                value: Some(format!("{}", r * 100)),
                ..Default::default()
            },
        ];
        rows.push(Row {
            index: row_num,
            cells,
            height: None,
            hidden: false,
            outline_level: None,
            collapsed: false,
        });
    }

    WorkbookData {
        sheets: vec![SheetData {
            name: "Data".into(),
            state: None,
            worksheet: WorksheetXml {
                dimension: Some(format!("A1:{}{row_count}", col_letter(5))),
                rows,
                merge_cells: Vec::new(),
                auto_filter: None,
                frozen_pane: None,
                split_pane: None,
                pane_selections: Vec::new(),
                sheet_view: None,
                columns: Vec::new(),
                data_validations: Vec::new(),
                conditional_formatting: Vec::new(),
                hyperlinks: Vec::new(),
                page_setup: None,
                sheet_protection: None,
                comments: Vec::new(),
                tab_color: None,
                tables: Vec::new(),
                header_footer: None,
                outline_properties: None,
                sparkline_groups: Vec::new(),
                charts: Vec::new(),
                pivot_tables: Vec::new(),
                threaded_comments: Vec::new(),
                slicers: Vec::new(),
                timelines: Vec::new(),
                preserved_extensions: Vec::new(),
            },
        }],
        date_system: DateSystem::Date1900,
        styles: Styles::default_styles(),
        defined_names: Vec::new(),
        shared_strings: None,
        doc_properties: None,
        theme_colors: None,
        calc_chain: Vec::new(),
        workbook_views: Vec::new(),
        protection: None,
        pivot_caches: Vec::new(),
        pivot_cache_records: Vec::new(),
        persons: Vec::new(),
        slicer_caches: Vec::new(),
        timeline_caches: Vec::new(),
        preserved_entries: std::collections::BTreeMap::new(),
    }
}

#[test]
fn bench_write_10k() {
    let wb = make_workbook(10_000);
    let start = Instant::now();
    let bytes = write_xlsx(&wb).unwrap();
    let elapsed = start.elapsed();
    println!(
        "Write 10K rows: {elapsed:?}, {:.1} KB",
        bytes.len() as f64 / 1024.0
    );
}

#[test]
fn bench_write_100k() {
    let wb = make_workbook(100_000);
    let start = Instant::now();
    let bytes = write_xlsx(&wb).unwrap();
    let elapsed = start.elapsed();
    println!(
        "Write 100K rows: {elapsed:?}, {:.1} KB",
        bytes.len() as f64 / 1024.0
    );
}

#[test]
fn bench_roundtrip_10k() {
    let wb = make_workbook(10_000);
    let bytes = write_xlsx(&wb).unwrap();
    let start = Instant::now();
    let wb2 = read_xlsx(&bytes).unwrap();
    let elapsed = start.elapsed();
    println!("Read 10K rows: {elapsed:?}");
    assert_eq!(wb2.sheets[0].worksheet.rows.len(), 10_000);
}

#[test]
fn bench_roundtrip_100k() {
    let wb = make_workbook(100_000);
    let bytes = write_xlsx(&wb).unwrap();
    let start = Instant::now();
    let wb2 = read_xlsx(&bytes).unwrap();
    let elapsed = start.elapsed();
    println!("Read 100K rows: {elapsed:?}");
    assert_eq!(wb2.sheets[0].worksheet.rows.len(), 100_000);
}
