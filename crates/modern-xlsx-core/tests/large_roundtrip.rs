//! Round-trip test for a large (10K-row) workbook.
//!
//! No timing: it asserts the row count and two cell values after `write_xlsx` -> `read_xlsx`.
//! Timing lives in `benches/read_write.rs`.

use modern_xlsx_core::ooxml::cell::col_to_letters;
use modern_xlsx_core::ooxml::worksheet::{Cell, CellType, Row, WorksheetXml};
use modern_xlsx_core::reader::read_xlsx;
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

#[test]
fn roundtrip_10k_rows() {
    let bytes = write_xlsx(&make_workbook(10_000)).unwrap();
    let wb = read_xlsx(&bytes).unwrap();
    let rows = &wb.sheets[0].worksheet.rows;
    assert_eq!(rows.len(), 10_000);

    let b1 = &rows[0].cells[1];
    assert_eq!(b1.reference, "B1");
    assert_eq!(b1.cell_type, CellType::SharedString);
    assert_eq!(b1.value.as_deref(), Some("Name_0"));

    let a_last = &rows[9_999].cells[0];
    assert_eq!(a_last.reference, "A10000");
    assert_eq!(a_last.value.as_deref(), Some("10000"));
}
