//! Golden round-trip tests: each case builds a workbook, writes and re-reads it through `tests/common`, and compares the result with `tests/golden/<name>.json`; `golden_inputs_are_valid` checks that every number cell in the committed goldens parses as a number.

use modern_xlsx_core::ooxml::comments::Comment;
use modern_xlsx_core::ooxml::styles::{
    Alignment, Border, BorderSide, CellXf, DxfStyle, Fill, Font, Styles,
};
use modern_xlsx_core::ooxml::worksheet::{
    Cell, CellType, ColumnInfo, ConditionalFormatting, ConditionalFormattingRule, DataValidation,
    FrozenPane, HeaderFooter, Hyperlink, PageSetup, Row, WorksheetXml,
};
use modern_xlsx_core::validate::validate_workbook;
use modern_xlsx_core::{SheetData, WorkbookData};

use crate::common::{
    assert_golden, assert_same, minimal_workbook, minimal_worksheet, number_row, roundtrip,
    roundtrip_golden, text_cell, text_row,
};

/// Every golden input must be a valid workbook: no errors and no warnings.
fn assert_valid(wb: &WorkbookData) {
    let report = validate_workbook(wb);
    assert_eq!(
        report.error_count, 0,
        "invalid golden input: {:?}",
        report.issues
    );
    assert_eq!(
        report.warning_count, 0,
        "invalid golden input: {:?}",
        report.issues
    );
}

/// The first sheet's worksheet.
fn sheet0(wb: &WorkbookData) -> &WorksheetXml {
    &wb.sheets[0].worksheet
}

// ---------------------------------------------------------------------------
// Test 1: Simple roundtrip (original)
// ---------------------------------------------------------------------------

#[test]
fn golden_simple_roundtrip() {
    let ws = minimal_worksheet(vec![Row {
        index: 1,
        cells: vec![
            Cell {
                reference: "A1".into(),
                cell_type: CellType::Number,
                value: Some("42".into()),
                ..Default::default()
            },
            Cell {
                reference: "B1".into(),
                cell_type: CellType::Boolean,
                value: Some("1".into()),
                ..Default::default()
            },
        ],
        ..Default::default()
    }]);
    let wb = minimal_workbook("Numbers", ws);
    assert_valid(&wb);
    let wb2 = roundtrip_golden(&wb, "simple_roundtrip");
    assert_same(&sheet0(&wb).rows, &sheet0(&wb2).rows, "rows");
}

// ---------------------------------------------------------------------------
// Test 2: Multi-sheet roundtrip (original)
// ---------------------------------------------------------------------------

#[test]
fn golden_multi_sheet_roundtrip() {
    let wb = WorkbookData {
        sheets: vec![
            SheetData {
                name: "Sheet1".into(),
                worksheet: minimal_worksheet(vec![number_row(1, "A1", "100")]),
                ..Default::default()
            },
            SheetData {
                name: "Sheet2".into(),
                worksheet: minimal_worksheet(vec![]),
                ..Default::default()
            },
        ],
        ..Default::default()
    };
    assert_valid(&wb);
    let wb2 = roundtrip_golden(&wb, "multi_sheet_roundtrip");
    let names = |w: &WorkbookData| w.sheets.iter().map(|s| s.name.clone()).collect::<Vec<_>>();
    assert_eq!(names(&wb), names(&wb2), "sheet names changed in round-trip");
    for (a, b) in wb.sheets.iter().zip(&wb2.sheets) {
        assert_same(&a.worksheet.rows, &b.worksheet.rows, "rows");
    }
}

// ---------------------------------------------------------------------------
// Test 3: Styled workbook — multiple cell style indices
// ---------------------------------------------------------------------------

#[test]
fn golden_styled_workbook() {
    let mut styles = Styles::default_styles();

    // Add a bold font (index 1)
    styles.fonts.push(Font {
        name: Some("Aptos".into()),
        size: Some(14.0),
        bold: true,
        italic: false,
        underline: false,
        strike: false,
        color: Some("FFFF0000".into()),
        ..Default::default()
    });

    // Add a fill with a solid yellow background (index 2)
    styles.fills.push(Fill {
        pattern_type: "solid".into(),
        fg_color: Some("FFFFFF00".into()),
        bg_color: None,
        gradient_fill: None,
        ..Default::default()
    });

    // Add a border with thin sides (index 1)
    styles.borders.push(Border {
        left: Some(BorderSide {
            style: "thin".into(),
            color: Some("FF000000".into()),
            color_spec: None,
        }),
        right: Some(BorderSide {
            style: "thin".into(),
            color: Some("FF000000".into()),
            color_spec: None,
        }),
        top: Some(BorderSide {
            style: "thin".into(),
            color: Some("FF000000".into()),
            color_spec: None,
        }),
        bottom: Some(BorderSide {
            style: "thin".into(),
            color: Some("FF000000".into()),
            color_spec: None,
        }),
        diagonal: None,
        diagonal_up: false,
        diagonal_down: false,
    });

    // CellXf index 1: bold font
    styles.cell_xfs.push(CellXf {
        font_id: 1,
        apply_font: true,
        ..Default::default()
    });

    // CellXf index 2: solid yellow fill
    styles.cell_xfs.push(CellXf {
        fill_id: 2,
        apply_fill: true,
        ..Default::default()
    });

    // CellXf index 3: thin borders + center alignment
    styles.cell_xfs.push(CellXf {
        border_id: 1,
        apply_border: true,
        apply_alignment: true,
        alignment: Some(Alignment {
            horizontal: Some("center".into()),
            vertical: Some("center".into()),
            wrap_text: true,
            ..Default::default()
        }),
        ..Default::default()
    });

    let ws = minimal_worksheet(vec![Row {
        index: 1,
        cells: vec![
            Cell {
                reference: "A1".into(),
                cell_type: CellType::Number,
                value: Some("1".into()),
                style_index: Some(1),
                ..Default::default()
            },
            Cell {
                reference: "B1".into(),
                cell_type: CellType::Number,
                value: Some("2".into()),
                style_index: Some(2),
                ..Default::default()
            },
            Cell {
                reference: "C1".into(),
                cell_type: CellType::Number,
                value: Some("3".into()),
                style_index: Some(3),
                ..Default::default()
            },
        ],
        ..Default::default()
    }]);

    let mut wb = minimal_workbook("Styled", ws);
    wb.styles = styles;
    assert_valid(&wb);
    let wb2 = roundtrip_golden(&wb, "styled_workbook");
    assert_same(&sheet0(&wb).rows, &sheet0(&wb2).rows, "rows");
    assert_same(&wb.styles.cell_xfs, &wb2.styles.cell_xfs, "styles.cell_xfs");
}

// ---------------------------------------------------------------------------
// Test 4: Frozen pane
// ---------------------------------------------------------------------------

#[test]
fn golden_frozen_pane() {
    let mut ws = minimal_worksheet(vec![
        text_row(1, "A1", "Header1"),
        text_row(2, "A2", "Data1"),
        text_row(3, "A3", "Data2"),
    ]);
    ws.frozen_pane = Some(FrozenPane { rows: 1, cols: 1 });
    let wb = minimal_workbook("FrozenPane", ws);
    assert_valid(&wb);
    let wb2 = roundtrip_golden(&wb, "frozen_pane");
    assert_same(&sheet0(&wb).rows, &sheet0(&wb2).rows, "rows");
    assert_same(
        &sheet0(&wb).frozen_pane,
        &sheet0(&wb2).frozen_pane,
        "frozen_pane",
    );
}

// ---------------------------------------------------------------------------
// Test 5: Data validation
// ---------------------------------------------------------------------------

#[test]
fn golden_data_validation() {
    let mut ws = minimal_worksheet(vec![number_row(1, "A1", "5")]);
    ws.data_validations = vec![
        // List validation
        DataValidation {
            sqref: "B1:B10".into(),
            validation_type: Some("list".into()),
            operator: None,
            formula1: Some("\"Yes,No,Maybe\"".into()),
            formula2: None,
            allow_blank: Some(true),
            show_error_message: Some(true),
            error_title: Some("Invalid".into()),
            error_message: Some("Please select from the list".into()),
            show_input_message: Some(true),
            prompt_title: Some("Choose".into()),
            prompt: Some("Select a value".into()),
            error_style: None,
            ime_mode: None,
            show_drop_down: None,
        },
        // Decimal range validation
        DataValidation {
            sqref: "C1:C10".into(),
            validation_type: Some("decimal".into()),
            operator: Some("between".into()),
            formula1: Some("0".into()),
            formula2: Some("100".into()),
            allow_blank: Some(false),
            show_error_message: Some(true),
            error_title: Some("Out of range".into()),
            error_message: Some("Value must be between 0 and 100".into()),
            show_input_message: None,
            prompt_title: None,
            prompt: None,
            error_style: None,
            ime_mode: None,
            show_drop_down: None,
        },
    ];
    let wb = minimal_workbook("DataValidation", ws);
    assert_valid(&wb);
    let wb2 = roundtrip_golden(&wb, "data_validation");
    assert_same(&sheet0(&wb).rows, &sheet0(&wb2).rows, "rows");
    assert_same(
        &sheet0(&wb).data_validations,
        &sheet0(&wb2).data_validations,
        "data_validations",
    );
}

// ---------------------------------------------------------------------------
// Test 6: Conditional formatting
// ---------------------------------------------------------------------------

#[test]
fn golden_conditional_formatting() {
    let mut styles = Styles::default_styles();
    // Add a DXF for the conditional formatting rule (dxf_id 0)
    styles.dxfs.push(DxfStyle {
        font: Some(Font {
            bold: true,
            color: Some("FF9C0006".into()),
            ..Default::default()
        }),
        fill: Some(Fill {
            pattern_type: "solid".into(),
            fg_color: Some("FFFFC7CE".into()),
            bg_color: None,
            gradient_fill: None,
            ..Default::default()
        }),
        border: None,
        num_fmt: None,
        ..Default::default()
    });

    let mut ws = minimal_worksheet(vec![
        number_row(1, "A1", "10"),
        number_row(2, "A2", "50"),
        number_row(3, "A3", "90"),
    ]);
    ws.conditional_formatting = vec![ConditionalFormatting {
        sqref: "A1:A3".into(),
        rules: vec![ConditionalFormattingRule {
            rule_type: "cellIs".into(),
            priority: 1,
            operator: Some("greaterThan".into()),
            formula: Some("50".into()),
            additional_formulas: Vec::new(),
            dxf_id: Some(0),
            stop_if_true: None,
            above_average: None,
            percent: None,
            bottom: None,
            text: None,
            time_period: None,
            rank: None,
            std_dev: None,
            equal_average: None,
            color_scale: None,
            data_bar: None,
            icon_set: None,
        }],
    }];

    let mut wb = minimal_workbook("ConditionalFormatting", ws);
    wb.styles = styles;
    assert_valid(&wb);
    let wb2 = roundtrip_golden(&wb, "conditional_formatting");
    assert_same(&sheet0(&wb).rows, &sheet0(&wb2).rows, "rows");
    assert_same(
        &sheet0(&wb).conditional_formatting,
        &sheet0(&wb2).conditional_formatting,
        "conditional_formatting",
    );
}

// ---------------------------------------------------------------------------
// Test 7: Hyperlinks
// ---------------------------------------------------------------------------

#[test]
fn golden_hyperlinks() {
    let mut ws = minimal_worksheet(vec![Row {
        index: 1,
        cells: vec![
            Cell {
                reference: "A1".into(),
                cell_type: CellType::SharedString,
                value: Some("Click here".into()),
                ..Default::default()
            },
            Cell {
                reference: "B1".into(),
                cell_type: CellType::SharedString,
                value: Some("Go to Sheet2".into()),
                ..Default::default()
            },
        ],
        ..Default::default()
    }]);
    ws.hyperlinks = vec![
        Hyperlink {
            cell_ref: "A1".into(),
            target: Some("https://example.com".into()),
            display: Some("Example Site".into()),
            tooltip: Some("Visit example.com".into()),
            ..Default::default()
        },
        Hyperlink {
            cell_ref: "B1".into(),
            location: Some("Sheet2!A1".into()),
            display: Some("Go to Sheet2".into()),
            ..Default::default()
        },
    ];
    let wb = minimal_workbook("Hyperlinks", ws);
    assert_valid(&wb);
    let mut wb2 = roundtrip(&wb);
    assert_same(&sheet0(&wb).rows, &sheet0(&wb2).rows, "rows");
    assert_same(
        &sheet0(&wb).hyperlinks,
        &sheet0(&wb2).hyperlinks,
        "hyperlinks",
    );
    // The sheet rels part is asserted as text; the golden keeps no byte arrays.
    let rels = String::from_utf8_lossy(
        wb2.preserved_entries
            .get("xl/worksheets/_rels/sheet1.xml.rels")
            .expect("sheet1 rels preserved"),
    )
    .into_owned();
    for needle in [
        "relationships/hyperlink",
        r#"Target="https://example.com""#,
        r#"TargetMode="External""#,
    ] {
        assert!(rels.contains(needle), "missing {needle} in {rels}");
    }
    wb2.preserved_entries.clear();
    assert_golden(
        &serde_json::to_string_pretty(&wb2).expect("JSON serialization failed"),
        "hyperlinks",
    );
}

// ---------------------------------------------------------------------------
// Test 8: Comments
// ---------------------------------------------------------------------------

#[test]
fn golden_comments() {
    let mut ws = minimal_worksheet(vec![number_row(1, "A1", "10"), number_row(2, "A2", "20")]);
    ws.comments = vec![
        Comment {
            cell_ref: "A1".into(),
            author: "Alice".into(),
            text: "This is the first value".into(),
        },
        Comment {
            cell_ref: "A2".into(),
            author: "Bob".into(),
            text: "This is the second value".into(),
        },
    ];
    let wb = minimal_workbook("Comments", ws);
    assert_valid(&wb);
    let mut wb2 = roundtrip(&wb);
    assert_same(&sheet0(&wb).rows, &sheet0(&wb2).rows, "rows");
    assert_same(&sheet0(&wb).comments, &sheet0(&wb2).comments, "comments");

    // The comments relationship lives in the preserved sheet rels part. Check
    // it here, then drop the opaque entries so the golden stays readable.
    let rels = wb2
        .preserved_entries
        .get("xl/worksheets/_rels/sheet1.xml.rels")
        .expect("sheet1 rels preserved");
    let rels = String::from_utf8_lossy(rels);
    assert!(
        rels.contains("relationships/comments"),
        "sheet1 rels must hold the comments relationship: {rels}"
    );
    wb2.preserved_entries.clear();
    let json = serde_json::to_string_pretty(&wb2).expect("JSON serialization failed");
    assert_golden(&json, "comments");
}

// ---------------------------------------------------------------------------
// Test 9: Merge cells
// ---------------------------------------------------------------------------

#[test]
fn golden_merge_cells() {
    let mut ws = minimal_worksheet(vec![
        Row {
            index: 1,
            cells: vec![Cell {
                reference: "A1".into(),
                cell_type: CellType::SharedString,
                value: Some("Merged Header".into()),
                ..Default::default()
            }],
            ..Default::default()
        },
        Row {
            index: 3,
            cells: vec![Cell {
                reference: "B3".into(),
                cell_type: CellType::Number,
                value: Some("99".into()),
                ..Default::default()
            }],
            ..Default::default()
        },
    ]);
    ws.merge_cells = vec!["A1:D1".into(), "B3:C5".into()];
    let wb = minimal_workbook("MergeCells", ws);
    assert_valid(&wb);
    let wb2 = roundtrip_golden(&wb, "merge_cells");
    assert_same(&sheet0(&wb).rows, &sheet0(&wb2).rows, "rows");
    assert_same(
        &sheet0(&wb).merge_cells,
        &sheet0(&wb2).merge_cells,
        "merge_cells",
    );
}

// ---------------------------------------------------------------------------
// Test 10: Column widths and hidden columns
// ---------------------------------------------------------------------------

#[test]
fn golden_column_widths() {
    let mut ws = minimal_worksheet(vec![Row {
        index: 1,
        cells: vec![text_cell("A1", "Wide"), text_cell("D1", "Visible")],
        ..Default::default()
    }]);
    ws.columns = vec![
        ColumnInfo {
            min: 1,
            max: 1,
            width: 25.0,
            hidden: false,
            custom_width: true,
            outline_level: None,
            collapsed: false,
            style_index: None,
            best_fit: false,
            phonetic: false,
        },
        ColumnInfo {
            min: 2,
            max: 2,
            width: 8.0,
            hidden: true,
            custom_width: true,
            outline_level: None,
            collapsed: false,
            style_index: None,
            best_fit: false,
            phonetic: false,
        },
        ColumnInfo {
            min: 3,
            max: 3,
            width: 15.5,
            hidden: false,
            custom_width: true,
            outline_level: Some(1),
            collapsed: false,
            style_index: None,
            best_fit: false,
            phonetic: false,
        },
        ColumnInfo {
            min: 4,
            max: 4,
            width: 12.0,
            hidden: false,
            custom_width: true,
            outline_level: None,
            collapsed: false,
            style_index: None,
            best_fit: false,
            phonetic: false,
        },
    ];
    let wb = minimal_workbook("ColumnWidths", ws);
    assert_valid(&wb);
    let wb2 = roundtrip_golden(&wb, "column_widths");
    assert_same(&sheet0(&wb).rows, &sheet0(&wb2).rows, "rows");
    assert_same(&sheet0(&wb).columns, &sheet0(&wb2).columns, "columns");
}

// ---------------------------------------------------------------------------
// Test 11: Page setup
// ---------------------------------------------------------------------------

#[test]
fn golden_page_setup() {
    let mut ws = minimal_worksheet(vec![text_row(1, "A1", "Print me")]);
    ws.page_setup = Some(PageSetup {
        paper_size: Some(9), // A4
        orientation: Some("landscape".into()),
        fit_to_width: Some(1),
        fit_to_height: Some(0),
        scale: Some(85),
        first_page_number: Some(1),
        horizontal_dpi: Some(300),
        vertical_dpi: Some(300),
        fit_to_page: None,
        margins: None,
    });
    let wb = minimal_workbook("PageSetup", ws);
    assert_valid(&wb);
    let wb2 = roundtrip_golden(&wb, "page_setup");
    assert_same(&sheet0(&wb).rows, &sheet0(&wb2).rows, "rows");
    assert_same(
        &sheet0(&wb).page_setup,
        &sheet0(&wb2).page_setup,
        "page_setup",
    );
}

// ---------------------------------------------------------------------------
// Test 12: Header and footer
// ---------------------------------------------------------------------------

#[test]
fn golden_header_footer() {
    let mut ws = minimal_worksheet(vec![text_row(1, "A1", "Report Data")]);
    ws.header_footer = Some(HeaderFooter {
        odd_header: Some("&LCompany Name&CMonthly Report&R&D".into()),
        odd_footer: Some("&LConfidential&CPage &P of &N&R&F".into()),
        even_header: None,
        even_footer: None,
        first_header: Some("&CFirst Page Header".into()),
        first_footer: None,
        different_odd_even: false,
        different_first: true,
        scale_with_doc: true,
        align_with_margins: true,
    });
    let wb = minimal_workbook("HeaderFooter", ws);
    assert_valid(&wb);
    let wb2 = roundtrip_golden(&wb, "header_footer");
    assert_same(&sheet0(&wb).rows, &sheet0(&wb2).rows, "rows");
    assert_same(
        &sheet0(&wb).header_footer,
        &sheet0(&wb2).header_footer,
        "header_footer",
    );
}

// ---------------------------------------------------------------------------
// Golden inputs must be valid
// ---------------------------------------------------------------------------

/// A number cell in a committed golden must hold a value that parses as f64.
#[test]
fn golden_inputs_are_valid() {
    fn walk(v: &serde_json::Value, file: &str) {
        match v {
            serde_json::Value::Object(map) => {
                if map.get("cellType").and_then(|t| t.as_str()) == Some("number")
                    && let Some(value) = map.get("value").and_then(|x| x.as_str())
                {
                    assert!(
                        value.parse::<f64>().is_ok(),
                        "{file}: number cell {:?} holds non-numeric value {value:?}",
                        map.get("reference")
                    );
                }
                map.values().for_each(|x| walk(x, file));
            }
            serde_json::Value::Array(items) => items.iter().for_each(|x| walk(x, file)),
            _ => {}
        }
    }

    let dir = std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("tests/golden");
    let mut seen = 0;
    for entry in std::fs::read_dir(&dir).expect("read tests/golden") {
        let path = entry.expect("dir entry").path();
        if path.extension().is_some_and(|e| e == "json") {
            let text = std::fs::read_to_string(&path).expect("read golden");
            let value: serde_json::Value = serde_json::from_str(&text).expect("golden is JSON");
            walk(&value, &path.display().to_string());
            seen += 1;
        }
    }
    assert!(seen >= 12, "expected at least 12 goldens, found {seen}");
}
