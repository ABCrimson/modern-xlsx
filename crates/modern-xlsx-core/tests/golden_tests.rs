use modern_xlsx_core::dates::DateSystem;
use modern_xlsx_core::ooxml::styles::Styles;
use modern_xlsx_core::ooxml::worksheet::{Cell, CellType, Row, WorksheetXml};
use modern_xlsx_core::reader::read_xlsx;
use modern_xlsx_core::writer::write_xlsx;
use modern_xlsx_core::{SheetData, WorkbookData};
use pretty_assertions::assert_eq;

fn build_simple_workbook() -> WorkbookData {
    WorkbookData {
        sheets: vec![SheetData {
            name: "Numbers".into(),
            worksheet: WorksheetXml {
                rows: vec![Row {
                    index: 1,
                    cells: vec![
                        Cell {
                            reference: "A1".into(),
                            cell_type: CellType::Number,
                            value: Some("42".into()),
                            formula: None,
                            formula_type: None,
                            formula_ref: None,
                            shared_index: None,
                            inline_string: None,
                            dynamic_array: None,
                            style_index: None,
                        },
                        Cell {
                            reference: "B1".into(),
                            cell_type: CellType::Boolean,
                            value: Some("1".into()),
                            formula: None,
                            formula_type: None,
                            formula_ref: None,
                            shared_index: None,
                            inline_string: None,
                            dynamic_array: None,
                            style_index: None,
                        },
                    ],
                    height: None,
                    hidden: false,
                    outline_level: None,
                    collapsed: false,
                }],
                merge_cells: vec![],
                auto_filter: None,
                frozen_pane: None,
            split_pane: None,
            pane_selections: vec![],
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
            },
        }],
        date_system: DateSystem::Date1900,
        styles: Styles::default_styles(),
        defined_names: vec![],
        shared_strings: None,
        doc_properties: None,
        theme_colors: None,
        calc_chain: vec![],
        workbook_views: vec![],
        preserved_entries: std::collections::BTreeMap::new(),
    }
}

#[test]
fn golden_simple_roundtrip() {
    let wb = build_simple_workbook();

    // Write to XLSX bytes
    let xlsx_bytes = write_xlsx(&wb).expect("write_xlsx failed");

    // Read it back
    let wb2 = read_xlsx(&xlsx_bytes).expect("read_xlsx failed");

    // Serialize to JSON for comparison
    let json = serde_json::to_string_pretty(&wb2).expect("JSON serialization failed");

    // For the first run, we create the golden file.
    // On subsequent runs, we compare against the snapshot.
    let golden_path = concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/tests/golden/simple_roundtrip.json"
    );

    if std::path::Path::new(golden_path).exists() {
        let expected = std::fs::read_to_string(golden_path).expect("read golden file");
        assert_eq!(
            json.trim(),
            expected.trim(),
            "Golden file mismatch for simple_roundtrip"
        );
    } else {
        // Create the golden file directory and write it
        std::fs::create_dir_all(concat!(env!("CARGO_MANIFEST_DIR"), "/tests/golden")).ok();
        std::fs::write(golden_path, &json).expect("write golden file");
        println!("Created golden file: {golden_path}");
    }
}

#[test]
fn golden_multi_sheet_roundtrip() {
    let wb = WorkbookData {
        sheets: vec![
            SheetData {
                name: "Sheet1".into(),
                worksheet: WorksheetXml {
                    rows: vec![Row {
                        index: 1,
                        cells: vec![Cell {
                            reference: "A1".into(),
                            cell_type: CellType::Number,
                            value: Some("100".into()),
                            formula: None,
                            formula_type: None,
                            formula_ref: None,
                            shared_index: None,
                            inline_string: None,
                            dynamic_array: None,
                            style_index: None,
                        }],
                        height: None,
                        hidden: false,
                        outline_level: None,
                        collapsed: false,
                    }],
                    merge_cells: vec![],
                    auto_filter: None,
                    frozen_pane: None,
            split_pane: None,
            pane_selections: vec![],
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
                },
            },
            SheetData {
                name: "Sheet2".into(),
                worksheet: WorksheetXml {
                    rows: vec![],
                    merge_cells: vec![],
                    auto_filter: None,
                    frozen_pane: None,
            split_pane: None,
            pane_selections: vec![],
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
                },
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
        preserved_entries: std::collections::BTreeMap::new(),
    };

    let xlsx_bytes = write_xlsx(&wb).expect("write_xlsx failed");
    let wb2 = read_xlsx(&xlsx_bytes).expect("read_xlsx failed");
    let json = serde_json::to_string_pretty(&wb2).expect("JSON serialization failed");

    let golden_path = concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/tests/golden/multi_sheet_roundtrip.json"
    );

    if std::path::Path::new(golden_path).exists() {
        let expected = std::fs::read_to_string(golden_path).expect("read golden file");
        assert_eq!(
            json.trim(),
            expected.trim(),
            "Golden file mismatch for multi_sheet_roundtrip"
        );
    } else {
        std::fs::create_dir_all(concat!(env!("CARGO_MANIFEST_DIR"), "/tests/golden")).ok();
        std::fs::write(golden_path, &json).expect("write golden file");
        println!("Created golden file: {golden_path}");
    }
}
