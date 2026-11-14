//! Golden tests for the table and comment parts. Each serialises only the
//! subtree under test (see `crate::common::golden_subtree`).

use modern_xlsx_core::ooxml::tables::{TableColumn, TableDefinition, TableStyleInfo};
use modern_xlsx_core::ooxml::threaded_comments::{MentionData, PersonData, ThreadedCommentData};
use modern_xlsx_core::ooxml::worksheet::{Cell, Row, WorksheetXml};

use crate::common::{golden_subtree, minimal_workbook, number_row, roundtrip, text_row};

/// A cell holding only a formula.
fn formula_cell(cell_ref: &str, formula: &str) -> Cell {
    Cell {
        reference: cell_ref.into(),
        formula: Some(formula.into()),
        ..Default::default()
    }
}

/// A table with a totals row: a label, a built-in function and a custom
/// function whose `totalsRowFormula` must survive write -> read.
#[test]
fn golden_table_totals_row() {
    let table = TableDefinition {
        id: 1,
        name: Some("Sales".into()),
        display_name: "Sales".into(),
        ref_range: "A1:C4".into(),
        header_row_count: 1,
        totals_row_count: 1,
        totals_row_shown: true,
        columns: vec![
            TableColumn {
                id: 1,
                name: "Region & Area".into(),
                totals_row_label: Some("Total".into()),
                ..Default::default()
            },
            TableColumn {
                id: 2,
                name: "Amt".into(),
                totals_row_function: Some("sum".into()),
                ..Default::default()
            },
            TableColumn {
                id: 3,
                name: "Double".into(),
                totals_row_function: Some("custom".into()),
                calculated_column_formula: Some("Sales[[#This Row],[Amt]]*2".into()),
                totals_row_formula: Some("SUBTOTAL(109,[Double])".into()),
                ..Default::default()
            },
        ],
        style_info: Some(TableStyleInfo {
            name: Some("TableStyleMedium2".into()),
            show_first_column: false,
            show_last_column: false,
            show_row_stripes: true,
            show_column_stripes: false,
        }),
        auto_filter_ref: Some("A1:C3".into()),
    };
    let rows: Vec<Row> = vec![
        Row {
            index: 1,
            cells: vec![
                crate::common::text_cell("A1", "Region & Area"),
                crate::common::text_cell("B1", "Amt"),
                crate::common::text_cell("C1", "Double"),
            ],
            ..Default::default()
        },
        text_row(2, "A2", "East"),
        number_row(3, "B3", "10"),
        // The totals row: Excel rejects a column with a totals formula whose
        // totals-row cell holds no formula.
        Row {
            index: 4,
            cells: vec![
                crate::common::text_cell("A4", "Total"),
                formula_cell("B4", "SUBTOTAL(109,Sales[Amt])"),
                formula_cell("C4", "SUBTOTAL(109,Sales[Double])"),
            ],
            ..Default::default()
        },
    ];
    let mut ws = WorksheetXml {
        rows,
        ..Default::default()
    };
    ws.tables = vec![table];
    let wb = minimal_workbook("Sheet1", ws);

    let wb2 = roundtrip(&wb);
    golden_subtree(&wb2.sheets[0].worksheet.tables, "table_totals_row");
}

/// A threaded comment with `done` and a mention, and its person with a
/// `userId`, survive write -> read.
#[test]
fn golden_threaded_comments() {
    let person = "{D0000000-0000-4000-8000-0000000000A1}";
    let mut ws = WorksheetXml {
        rows: vec![text_row(1, "A1", "Hello")],
        ..Default::default()
    };
    ws.threaded_comments = vec![
        ThreadedCommentData {
            id: "{A0000000-0000-4000-8000-000000000001}".into(),
            ref_cell: "A1".into(),
            person_id: person.into(),
            text: "Hi @Jo & team".into(),
            timestamp: "2024-01-15T10:30:00.00".into(),
            parent_id: None,
            done: Some(true),
            mentions: vec![MentionData {
                mention_person_id: person.into(),
                mention_id: "{E0000000-0000-4000-8000-0000000000B1}".into(),
                start_index: 3,
                length: 3,
            }],
        },
        ThreadedCommentData {
            id: "{B0000000-0000-4000-8000-000000000002}".into(),
            ref_cell: "A1".into(),
            person_id: person.into(),
            text: "A reply".into(),
            timestamp: "2024-01-15T11:00:00.00".into(),
            parent_id: Some("{A0000000-0000-4000-8000-000000000001}".into()),
            done: None,
            mentions: Vec::new(),
        },
    ];
    let mut wb = minimal_workbook("Sheet1", ws);
    wb.persons = vec![PersonData {
        id: person.into(),
        display_name: "Jo Doe".into(),
        user_id: Some("S::jo@example.com::guid".into()),
        provider_id: Some("AD".into()),
    }];

    let wb2 = roundtrip(&wb);
    golden_subtree(
        &wb2.sheets[0].worksheet.threaded_comments,
        "threaded_comments",
    );
    assert_eq!(wb2.persons, wb.persons);
}
