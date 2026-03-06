use modern_xlsx_core::dates::DateSystem;
use modern_xlsx_core::ooxml::comments::Comment;
use modern_xlsx_core::ooxml::styles::{
    Alignment, Border, BorderSide, CellXf, DxfStyle, Fill, Font, Styles,
};
use modern_xlsx_core::ooxml::worksheet::{
    Cell, CellType, ColumnInfo, ConditionalFormatting, ConditionalFormattingRule, DataValidation,
    FrozenPane, HeaderFooter, Hyperlink, PageSetup, Row, WorksheetXml,
};
use modern_xlsx_core::reader::read_xlsx;
use modern_xlsx_core::writer::write_xlsx;
use modern_xlsx_core::{SheetData, WorkbookData};
use pretty_assertions::assert_eq;

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

/// Create a minimal WorksheetXml with only the specified rows.
fn minimal_worksheet(rows: Vec<Row>) -> WorksheetXml {
    WorksheetXml {
        rows,
        merge_cells: vec![],
        auto_filter: None,
        frozen_pane: None,
        split_pane: None,
        pane_selections: vec![],
        sheet_view: None,
        columns: vec![],
        dimension: None,
        data_validations: vec![],
        conditional_formatting: vec![],
        hyperlinks: vec![],
        page_setup: None,
        sheet_protection: None,
        comments: vec![],
        tab_color: None,
        header_footer: None,
        outline_properties: None,
        tables: vec![],
        sparkline_groups: vec![],
        charts: vec![],
        pivot_tables: vec![],
        threaded_comments: vec![],
        preserved_extensions: vec![],
    }
}

/// Create a minimal WorkbookData wrapping a single worksheet.
fn minimal_workbook(name: &str, worksheet: WorksheetXml) -> WorkbookData {
    WorkbookData {
        sheets: vec![SheetData {
            name: name.into(),
            state: None,
            worksheet,
        }],
        date_system: DateSystem::Date1900,
        styles: Styles::default_styles(),
        defined_names: vec![],
        shared_strings: None,
        doc_properties: None,
        theme_colors: None,
        calc_chain: vec![],
        workbook_views: vec![],
        protection: None,
        pivot_caches: vec![],
        pivot_cache_records: vec![],
        persons: vec![],
        preserved_entries: std::collections::BTreeMap::new(),
    }
}

/// Compare JSON against golden file (create if missing).
fn assert_golden(json: &str, name: &str) {
    let golden_path = format!(
        "{}/tests/golden/{}.json",
        env!("CARGO_MANIFEST_DIR"),
        name
    );
    if std::path::Path::new(&golden_path).exists() {
        let expected = std::fs::read_to_string(&golden_path).expect("read golden file");
        assert_eq!(
            json.trim(),
            expected.trim(),
            "Golden file mismatch for {}",
            name
        );
    } else {
        std::fs::create_dir_all(format!("{}/tests/golden", env!("CARGO_MANIFEST_DIR"))).ok();
        std::fs::write(&golden_path, json).expect("write golden file");
        println!("Created golden file: {golden_path}");
    }
}

/// Build a simple row with a single cell.
fn simple_row(index: u32, cell_ref: &str, value: &str) -> Row {
    Row {
        index,
        cells: vec![Cell {
            reference: cell_ref.into(),
            cell_type: CellType::Number,
            value: Some(value.into()),
            ..Default::default()
        }],
        height: None,
        hidden: false,
        outline_level: None,
        collapsed: false,
    }
}

/// Write -> Read -> JSON roundtrip, then compare against golden file.
fn roundtrip_golden(wb: &WorkbookData, name: &str) {
    let xlsx_bytes = write_xlsx(wb).expect("write_xlsx failed");
    let wb2 = read_xlsx(&xlsx_bytes).expect("read_xlsx failed");
    let json = serde_json::to_string_pretty(&wb2).expect("JSON serialization failed");
    assert_golden(&json, name);
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
        height: None,
        hidden: false,
        outline_level: None,
        collapsed: false,
    }]);
    let wb = minimal_workbook("Numbers", ws);
    roundtrip_golden(&wb, "simple_roundtrip");
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
                state: None,
                worksheet: minimal_worksheet(vec![simple_row(1, "A1", "100")]),
            },
            SheetData {
                name: "Sheet2".into(),
                state: None,
                worksheet: minimal_worksheet(vec![]),
            },
        ],
        date_system: DateSystem::Date1900,
        styles: Styles::default_styles(),
        defined_names: vec![],
        shared_strings: None,
        doc_properties: None,
        theme_colors: None,
        calc_chain: vec![],
        workbook_views: vec![],
        protection: None,
        pivot_caches: vec![],
        pivot_cache_records: vec![],
        persons: vec![],
        preserved_entries: std::collections::BTreeMap::new(),
    };
    roundtrip_golden(&wb, "multi_sheet_roundtrip");
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
    });

    // Add a border with thin sides (index 1)
    styles.borders.push(Border {
        left: Some(BorderSide {
            style: "thin".into(),
            color: Some("FF000000".into()),
        }),
        right: Some(BorderSide {
            style: "thin".into(),
            color: Some("FF000000".into()),
        }),
        top: Some(BorderSide {
            style: "thin".into(),
            color: Some("FF000000".into()),
        }),
        bottom: Some(BorderSide {
            style: "thin".into(),
            color: Some("FF000000".into()),
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
        height: None,
        hidden: false,
        outline_level: None,
        collapsed: false,
    }]);

    let mut wb = minimal_workbook("Styled", ws);
    wb.styles = styles;
    roundtrip_golden(&wb, "styled_workbook");
}

// ---------------------------------------------------------------------------
// Test 4: Frozen pane
// ---------------------------------------------------------------------------

#[test]
fn golden_frozen_pane() {
    let mut ws = minimal_worksheet(vec![
        simple_row(1, "A1", "Header1"),
        simple_row(2, "A2", "Data1"),
        simple_row(3, "A3", "Data2"),
    ]);
    ws.frozen_pane = Some(FrozenPane { rows: 1, cols: 1 });
    let wb = minimal_workbook("FrozenPane", ws);
    roundtrip_golden(&wb, "frozen_pane");
}

// ---------------------------------------------------------------------------
// Test 5: Data validation
// ---------------------------------------------------------------------------

#[test]
fn golden_data_validation() {
    let mut ws = minimal_worksheet(vec![simple_row(1, "A1", "5")]);
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
        },
    ];
    let wb = minimal_workbook("DataValidation", ws);
    roundtrip_golden(&wb, "data_validation");
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
        }),
        border: None,
        num_fmt: None,
    });

    let mut ws = minimal_worksheet(vec![
        simple_row(1, "A1", "10"),
        simple_row(2, "A2", "50"),
        simple_row(3, "A3", "90"),
    ]);
    ws.conditional_formatting = vec![ConditionalFormatting {
        sqref: "A1:A3".into(),
        rules: vec![ConditionalFormattingRule {
            rule_type: "cellIs".into(),
            priority: 1,
            operator: Some("greaterThan".into()),
            formula: Some("50".into()),
            dxf_id: Some(0),
            color_scale: None,
            data_bar: None,
            icon_set: None,
        }],
    }];

    let mut wb = minimal_workbook("ConditionalFormatting", ws);
    wb.styles = styles;
    roundtrip_golden(&wb, "conditional_formatting");
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
        height: None,
        hidden: false,
        outline_level: None,
        collapsed: false,
    }]);
    ws.hyperlinks = vec![
        Hyperlink {
            cell_ref: "A1".into(),
            location: Some("https://example.com".into()),
            display: Some("Example Site".into()),
            tooltip: Some("Visit example.com".into()),
        },
        Hyperlink {
            cell_ref: "B1".into(),
            location: Some("Sheet2!A1".into()),
            display: Some("Go to Sheet2".into()),
            tooltip: None,
        },
    ];
    let wb = minimal_workbook("Hyperlinks", ws);
    roundtrip_golden(&wb, "hyperlinks");
}

// ---------------------------------------------------------------------------
// Test 8: Comments
// ---------------------------------------------------------------------------

#[test]
fn golden_comments() {
    let mut ws = minimal_worksheet(vec![
        simple_row(1, "A1", "10"),
        simple_row(2, "A2", "20"),
    ]);
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
    roundtrip_golden(&wb, "comments");
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
            height: None,
            hidden: false,
            outline_level: None,
            collapsed: false,
        },
        Row {
            index: 3,
            cells: vec![Cell {
                reference: "B3".into(),
                cell_type: CellType::Number,
                value: Some("99".into()),
                ..Default::default()
            }],
            height: None,
            hidden: false,
            outline_level: None,
            collapsed: false,
        },
    ]);
    ws.merge_cells = vec!["A1:D1".into(), "B3:C5".into()];
    let wb = minimal_workbook("MergeCells", ws);
    roundtrip_golden(&wb, "merge_cells");
}

// ---------------------------------------------------------------------------
// Test 10: Column widths and hidden columns
// ---------------------------------------------------------------------------

#[test]
fn golden_column_widths() {
    let mut ws = minimal_worksheet(vec![simple_row(1, "A1", "Wide"), simple_row(1, "D1", "Visible")]);
    ws.columns = vec![
        ColumnInfo {
            min: 1,
            max: 1,
            width: 25.0,
            hidden: false,
            custom_width: true,
            outline_level: None,
            collapsed: false,
        },
        ColumnInfo {
            min: 2,
            max: 2,
            width: 8.0,
            hidden: true,
            custom_width: true,
            outline_level: None,
            collapsed: false,
        },
        ColumnInfo {
            min: 3,
            max: 3,
            width: 15.5,
            hidden: false,
            custom_width: true,
            outline_level: Some(1),
            collapsed: false,
        },
        ColumnInfo {
            min: 4,
            max: 4,
            width: 12.0,
            hidden: false,
            custom_width: true,
            outline_level: None,
            collapsed: false,
        },
    ];
    let wb = minimal_workbook("ColumnWidths", ws);
    roundtrip_golden(&wb, "column_widths");
}

// ---------------------------------------------------------------------------
// Test 11: Page setup
// ---------------------------------------------------------------------------

#[test]
fn golden_page_setup() {
    let mut ws = minimal_worksheet(vec![simple_row(1, "A1", "Print me")]);
    ws.page_setup = Some(PageSetup {
        paper_size: Some(9),          // A4
        orientation: Some("landscape".into()),
        fit_to_width: Some(1),
        fit_to_height: Some(0),
        scale: Some(85),
        first_page_number: Some(1),
        horizontal_dpi: Some(300),
        vertical_dpi: Some(300),
    });
    let wb = minimal_workbook("PageSetup", ws);
    roundtrip_golden(&wb, "page_setup");
}

// ---------------------------------------------------------------------------
// Test 12: Header and footer
// ---------------------------------------------------------------------------

#[test]
fn golden_header_footer() {
    let mut ws = minimal_worksheet(vec![simple_row(1, "A1", "Report Data")]);
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
    roundtrip_golden(&wb, "header_footer");
}
