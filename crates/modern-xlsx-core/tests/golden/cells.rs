//! Golden tests for cell content that the reader must return exactly:
//! formulas (cached values, shared and array masters, entity text), inline
//! strings (entities and edge whitespace) and error cells.
//!
//! Each test writes a workbook, reads it back, checks the rows against the
//! input, and compares only `sheets[0].worksheet.rows` with its golden.

use modern_xlsx_core::WorkbookData;
use modern_xlsx_core::ooxml::worksheet::{Cell, CellType, Row};

use crate::common::{assert_same, golden_subtree, minimal_workbook, minimal_worksheet, roundtrip};

/// Write and read `rows` as the only sheet, compare the rows read back with
/// the golden `name`, and return them.
fn roundtrip_rows(rows: Vec<Row>, name: &str) -> (WorkbookData, WorkbookData) {
    let wb = minimal_workbook("Cells", minimal_worksheet(rows));
    let wb2 = roundtrip(&wb);
    golden_subtree(&wb2.sheets[0].worksheet.rows, name);
    (wb, wb2)
}

fn cell(reference: &str, cell_type: CellType, value: Option<&str>) -> Cell {
    Cell {
        reference: reference.into(),
        cell_type,
        value: value.map(Into::into),
        ..Default::default()
    }
}

fn row(index: u32, cells: Vec<Cell>) -> Row {
    Row {
        index,
        cells,
        ..Default::default()
    }
}

#[test]
fn golden_cells_formulas() {
    let rows = vec![
        row(
            1,
            vec![
                cell("A1", CellType::Number, Some("1")),
                Cell {
                    formula: Some("A1*2".into()),
                    formula_type: Some("shared".into()),
                    formula_ref: Some("B1:B2".into()),
                    shared_index: Some(0),
                    ..cell("B1", CellType::Number, Some("2"))
                },
                Cell {
                    formula: Some(r#"IF(A1<1,"a & b",0)"#.into()),
                    ..cell("C1", CellType::FormulaStr, Some("a & b"))
                },
                Cell {
                    formula: Some("A1:A2*10".into()),
                    formula_type: Some("array".into()),
                    formula_ref: Some("D1:D2".into()),
                    ..cell("D1", CellType::Number, Some("10"))
                },
            ],
        ),
        row(
            2,
            vec![
                cell("A2", CellType::Number, Some("3")),
                Cell {
                    formula_type: Some("shared".into()),
                    shared_index: Some(0),
                    ..cell("B2", CellType::Number, Some("6"))
                },
            ],
        ),
    ];
    let (wb, wb2) = roundtrip_rows(rows, "cells_formulas");
    assert_same(
        &wb.sheets[0].worksheet.rows,
        &wb2.sheets[0].worksheet.rows,
        "formula rows",
    );
}

#[test]
fn golden_cells_inline_strings() {
    let inline = |reference: &str, text: &str| Cell {
        inline_string: Some(text.into()),
        ..cell(reference, CellType::InlineStr, Some(text))
    };
    let rows = vec![row(
        1,
        vec![
            inline("A1", "Tom & Jerry"),
            inline("B1", "  pad  "),
            inline("C1", "a < b > c"),
        ],
    )];
    let (wb, wb2) = roundtrip_rows(rows, "cells_inline_strings");
    assert_same(
        &wb.sheets[0].worksheet.rows,
        &wb2.sheets[0].worksheet.rows,
        "inline string rows",
    );
}

#[test]
fn golden_cells_errors() {
    let rows = vec![row(
        1,
        vec![
            cell("A1", CellType::Error, Some("#N/A")),
            cell("B1", CellType::Error, Some("#DIV/0!")),
            Cell {
                formula: Some("1/0".into()),
                ..cell("C1", CellType::Error, Some("#DIV/0!"))
            },
        ],
    )];
    let (wb, wb2) = roundtrip_rows(rows, "cells_errors");
    assert_same(
        &wb.sheets[0].worksheet.rows,
        &wb2.sheets[0].worksheet.rows,
        "error rows",
    );
}
