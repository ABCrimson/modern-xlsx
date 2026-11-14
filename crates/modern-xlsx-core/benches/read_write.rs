//! Criterion benchmarks: `write_xlsx` / `read_xlsx` at 10K and 100K rows, and
//! `read_xlsx` + `serde_json::to_string` against the streaming `read_xlsx_json`
//! at 100K rows.
//!
//! Run with `cargo bench -p modern-xlsx-core --bench read_write`; add
//! `-- --test` to run each benchmark once without measuring.

use std::hint::black_box;

use criterion::{Criterion, criterion_group, criterion_main};
use modern_xlsx_core::ooxml::cell::col_to_letters;
use modern_xlsx_core::ooxml::worksheet::{Cell, CellType, Row, WorksheetXml};
use modern_xlsx_core::reader::{read_xlsx, read_xlsx_json};
use modern_xlsx_core::writer::write_xlsx;
use modern_xlsx_core::{SheetData, WorkbookData};

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
            ..Default::default()
        });
    }

    WorkbookData {
        sheets: vec![SheetData {
            name: "Data".into(),
            worksheet: WorksheetXml {
                dimension: Some(format!("A1:{}{row_count}", col_to_letters(4))),
                rows,
                ..Default::default()
            },
            ..Default::default()
        }],
        ..Default::default()
    }
}

const SIZES: [usize; 2] = [10_000, 100_000];

fn bench_write(c: &mut Criterion) {
    let mut group = c.benchmark_group("write_xlsx");
    for rows in SIZES {
        if rows >= 100_000 {
            group.sample_size(10);
        }
        let wb = make_workbook(rows);
        group.bench_function(format!("{rows}_rows"), |b| {
            b.iter(|| write_xlsx(black_box(&wb)).unwrap());
        });
    }
    group.finish();
}

fn bench_read(c: &mut Criterion) {
    let mut group = c.benchmark_group("read_xlsx");
    for rows in SIZES {
        if rows >= 100_000 {
            group.sample_size(10);
        }
        let bytes = write_xlsx(&make_workbook(rows)).unwrap();
        group.bench_function(format!("{rows}_rows"), |b| {
            b.iter(|| read_xlsx(black_box(&bytes)).unwrap());
        });
    }
    group.finish();
}

fn bench_read_json(c: &mut Criterion) {
    let bytes = write_xlsx(&make_workbook(100_000)).unwrap();
    let mut group = c.benchmark_group("read_json_100k_rows");
    group.sample_size(10);
    group.bench_function("read_xlsx_then_to_string", |b| {
        b.iter(|| {
            let wb = read_xlsx(black_box(&bytes)).unwrap();
            serde_json::to_string(&wb).unwrap()
        });
    });
    group.bench_function("read_xlsx_json", |b| {
        b.iter(|| read_xlsx_json(black_box(&bytes)).unwrap());
    });
    group.finish();
}

criterion_group!(benches, bench_write, bench_read, bench_read_json);
criterion_main!(benches);
