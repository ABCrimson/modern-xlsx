use super::*;
use crate::SheetData;
use crate::dates::DateSystem;
use crate::ooxml::shared_strings::SharedStringTable;
use crate::ooxml::workbook::WorkbookXml as WbXml;
use crate::ooxml::worksheet::{Cell, CellType, Row, WorksheetXml};
use crate::zip::reader::{ZipSecurityLimits, read_zip_entries};
use pretty_assertions::assert_eq;
use std::collections::BTreeMap;

/// Helper: build a minimal single-sheet workbook with the given rows.
fn minimal_workbook(name: &str, rows: Vec<Row>) -> WorkbookData {
    WorkbookData {
        sheets: vec![SheetData {
            name: name.to_string(),
            worksheet: WorksheetXml {
                dimension: Some("A1".to_string()),
                rows,
                ..Default::default()
            },
            ..Default::default()
        }],
        ..Default::default()
    }
}

#[test]
fn test_write_empty_workbook() {
    let wb = WorkbookData {
        sheets: vec![SheetData {
            name: "Sheet1".to_string(),
            worksheet: WorksheetXml::default(),
            ..Default::default()
        }],
        ..Default::default()
    };

    let bytes = write_xlsx(&wb).expect("write_xlsx should succeed");
    assert!(!bytes.is_empty(), "output should not be empty");

    // Verify it is a valid ZIP with expected entries.
    let limits = ZipSecurityLimits::default();
    let entries = read_zip_entries(&bytes, &limits).expect("should be a valid ZIP");

    assert!(entries.contains_key("[Content_Types].xml"));
    assert!(entries.contains_key("_rels/.rels"));
    assert!(entries.contains_key("xl/workbook.xml"));
    assert!(entries.contains_key("xl/_rels/workbook.xml.rels"));
    assert!(entries.contains_key("xl/sharedStrings.xml"));
    assert!(entries.contains_key("xl/styles.xml"));
    assert!(entries.contains_key("xl/worksheets/sheet1.xml"));
}

#[test]
fn test_write_zero_sheets_errors() {
    let wb = WorkbookData::default();
    assert_eq!(write_xlsx(&wb).unwrap_err().code(), "INVALID_WORKBOOK");
}

#[test]
fn test_zip_contains_expected_entries_for_two_sheets() {
    let wb = WorkbookData {
        sheets: vec![
            SheetData {
                name: "First".to_string(),
                worksheet: WorksheetXml::default(),
                ..Default::default()
            },
            SheetData {
                name: "Second".to_string(),
                worksheet: WorksheetXml::default(),
                ..Default::default()
            },
        ],
        ..Default::default()
    };

    let bytes = write_xlsx(&wb).expect("write_xlsx should succeed");
    let limits = ZipSecurityLimits::default();
    let entries = read_zip_entries(&bytes, &limits).unwrap();

    assert!(entries.contains_key("xl/worksheets/sheet1.xml"));
    assert!(entries.contains_key("xl/worksheets/sheet2.xml"));
    assert!(!entries.contains_key("xl/worksheets/sheet3.xml"));
}

#[test]
fn test_write_with_string_cells() {
    let rows = vec![Row {
        index: 1,
        cells: vec![
            Cell {
                reference: "A1".to_string(),
                cell_type: CellType::SharedString,
                style_index: None,
                value: Some("Hello".to_string()),
                ..Default::default()
            },
            Cell {
                reference: "B1".to_string(),
                cell_type: CellType::SharedString,
                style_index: None,
                value: Some("World".to_string()),
                ..Default::default()
            },
        ],
        ..Default::default()
    }];

    let wb = minimal_workbook("Sheet1", rows);
    let bytes = write_xlsx(&wb).expect("write_xlsx should succeed");

    let limits = ZipSecurityLimits::default();
    let entries = read_zip_entries(&bytes, &limits).unwrap();

    // Verify the shared string table contains both strings.
    let sst_xml = std::str::from_utf8(entries.get("xl/sharedStrings.xml").unwrap()).unwrap();
    let sst = SharedStringTable::parse(sst_xml.as_bytes()).unwrap();
    assert_eq!(sst.len(), 2);
    assert_eq!(sst.get(0), Some("Hello"));
    assert_eq!(sst.get(1), Some("World"));

    // Verify the worksheet uses SST indices, not raw strings.
    let ws_xml = std::str::from_utf8(entries.get("xl/worksheets/sheet1.xml").unwrap()).unwrap();
    let ws = WorksheetXml::parse(ws_xml.as_bytes()).unwrap();
    assert_eq!(ws.rows.len(), 1);
    assert_eq!(ws.rows[0].cells.len(), 2);
    assert_eq!(ws.rows[0].cells[0].cell_type, CellType::SharedString);
    assert_eq!(ws.rows[0].cells[0].value.as_deref(), Some("0"));
    assert_eq!(ws.rows[0].cells[1].cell_type, CellType::SharedString);
    assert_eq!(ws.rows[0].cells[1].value.as_deref(), Some("1"));
}

#[test]
fn test_write_with_number_cells() {
    let rows = vec![Row {
        index: 1,
        cells: vec![
            Cell {
                reference: "A1".to_string(),
                cell_type: CellType::Number,
                style_index: None,
                value: Some("42".to_string()),
                ..Default::default()
            },
            Cell {
                reference: "B1".to_string(),
                cell_type: CellType::Number,
                style_index: None,
                value: Some("3.14".to_string()),
                ..Default::default()
            },
        ],
        ..Default::default()
    }];

    let wb = minimal_workbook("Sheet1", rows);
    let bytes = write_xlsx(&wb).expect("write_xlsx should succeed");

    let limits = ZipSecurityLimits::default();
    let entries = read_zip_entries(&bytes, &limits).unwrap();

    let ws_xml = std::str::from_utf8(entries.get("xl/worksheets/sheet1.xml").unwrap()).unwrap();
    let ws = WorksheetXml::parse(ws_xml.as_bytes()).unwrap();
    assert_eq!(ws.rows[0].cells[0].cell_type, CellType::Number);
    assert_eq!(ws.rows[0].cells[0].value.as_deref(), Some("42"));
    assert_eq!(ws.rows[0].cells[1].value.as_deref(), Some("3.14"));
}

#[test]
fn test_write_with_boolean_cells() {
    let rows = vec![Row {
        index: 1,
        cells: vec![
            Cell {
                reference: "A1".to_string(),
                cell_type: CellType::Boolean,
                style_index: None,
                value: Some("1".to_string()),
                ..Default::default()
            },
            Cell {
                reference: "B1".to_string(),
                cell_type: CellType::Boolean,
                style_index: None,
                value: Some("0".to_string()),
                ..Default::default()
            },
        ],
        ..Default::default()
    }];

    let wb = minimal_workbook("Sheet1", rows);
    let bytes = write_xlsx(&wb).expect("write_xlsx should succeed");

    let limits = ZipSecurityLimits::default();
    let entries = read_zip_entries(&bytes, &limits).unwrap();

    let ws_xml = std::str::from_utf8(entries.get("xl/worksheets/sheet1.xml").unwrap()).unwrap();
    let ws = WorksheetXml::parse(ws_xml.as_bytes()).unwrap();
    assert_eq!(ws.rows[0].cells[0].cell_type, CellType::Boolean);
    assert_eq!(ws.rows[0].cells[0].value.as_deref(), Some("1"));
    assert_eq!(ws.rows[0].cells[1].cell_type, CellType::Boolean);
    assert_eq!(ws.rows[0].cells[1].value.as_deref(), Some("0"));
}

#[test]
fn test_write_with_mixed_cell_types() {
    let rows = vec![
        Row {
            index: 1,
            cells: vec![
                Cell {
                    reference: "A1".to_string(),
                    cell_type: CellType::SharedString,
                    style_index: None,
                    value: Some("Name".to_string()),
                    ..Default::default()
                },
                Cell {
                    reference: "B1".to_string(),
                    cell_type: CellType::SharedString,
                    style_index: None,
                    value: Some("Value".to_string()),
                    ..Default::default()
                },
                Cell {
                    reference: "C1".to_string(),
                    cell_type: CellType::SharedString,
                    style_index: None,
                    value: Some("Active".to_string()),
                    ..Default::default()
                },
            ],
            ..Default::default()
        },
        Row {
            index: 2,
            cells: vec![
                Cell {
                    reference: "A2".to_string(),
                    cell_type: CellType::SharedString,
                    style_index: None,
                    value: Some("Pi".to_string()),
                    ..Default::default()
                },
                Cell {
                    reference: "B2".to_string(),
                    cell_type: CellType::Number,
                    style_index: None,
                    value: Some("3.14159".to_string()),
                    ..Default::default()
                },
                Cell {
                    reference: "C2".to_string(),
                    cell_type: CellType::Boolean,
                    style_index: None,
                    value: Some("1".to_string()),
                    ..Default::default()
                },
            ],
            ..Default::default()
        },
    ];

    let wb = minimal_workbook("MixedTypes", rows);
    let bytes = write_xlsx(&wb).expect("write_xlsx should succeed");

    let limits = ZipSecurityLimits::default();
    let entries = read_zip_entries(&bytes, &limits).unwrap();

    // Verify SST has 4 unique strings: "Name", "Value", "Active", "Pi".
    let sst_xml = std::str::from_utf8(entries.get("xl/sharedStrings.xml").unwrap()).unwrap();
    let sst = SharedStringTable::parse(sst_xml.as_bytes()).unwrap();
    assert_eq!(sst.len(), 4);

    // Verify workbook XML round-trips.
    let wb_xml = std::str::from_utf8(entries.get("xl/workbook.xml").unwrap()).unwrap();
    let wb_parsed = WbXml::parse(wb_xml.as_bytes()).unwrap();
    assert_eq!(wb_parsed.sheets.len(), 1);
    assert_eq!(wb_parsed.sheets[0].name, "MixedTypes");

    // Verify the worksheet.
    let ws_xml = std::str::from_utf8(entries.get("xl/worksheets/sheet1.xml").unwrap()).unwrap();
    let ws = WorksheetXml::parse(ws_xml.as_bytes()).unwrap();
    assert_eq!(ws.rows.len(), 2);

    // Row 1: all shared strings (SST indices).
    assert_eq!(ws.rows[0].cells[0].cell_type, CellType::SharedString);
    assert_eq!(ws.rows[0].cells[0].value.as_deref(), Some("0")); // "Name"
    assert_eq!(ws.rows[0].cells[1].value.as_deref(), Some("1")); // "Value"
    assert_eq!(ws.rows[0].cells[2].value.as_deref(), Some("2")); // "Active"

    // Row 2: string, number, boolean.
    assert_eq!(ws.rows[1].cells[0].cell_type, CellType::SharedString);
    assert_eq!(ws.rows[1].cells[0].value.as_deref(), Some("3")); // "Pi"
    assert_eq!(ws.rows[1].cells[1].cell_type, CellType::Number);
    assert_eq!(ws.rows[1].cells[1].value.as_deref(), Some("3.14159"));
    assert_eq!(ws.rows[1].cells[2].cell_type, CellType::Boolean);
    assert_eq!(ws.rows[1].cells[2].value.as_deref(), Some("1"));
}

#[test]
fn test_duplicate_strings_are_deduplicated() {
    let rows = vec![
        Row {
            index: 1,
            cells: vec![Cell {
                reference: "A1".to_string(),
                cell_type: CellType::SharedString,
                style_index: None,
                value: Some("Repeat".to_string()),
                ..Default::default()
            }],
            ..Default::default()
        },
        Row {
            index: 2,
            cells: vec![Cell {
                reference: "A2".to_string(),
                cell_type: CellType::SharedString,
                style_index: None,
                value: Some("Repeat".to_string()),
                ..Default::default()
            }],
            ..Default::default()
        },
        Row {
            index: 3,
            cells: vec![Cell {
                reference: "A3".to_string(),
                cell_type: CellType::SharedString,
                style_index: None,
                value: Some("Other".to_string()),
                ..Default::default()
            }],
            ..Default::default()
        },
    ];

    let wb = minimal_workbook("Sheet1", rows);
    let bytes = write_xlsx(&wb).expect("write_xlsx should succeed");

    let limits = ZipSecurityLimits::default();
    let entries = read_zip_entries(&bytes, &limits).unwrap();

    // Only 2 unique strings despite 3 cells.
    let sst_xml = std::str::from_utf8(entries.get("xl/sharedStrings.xml").unwrap()).unwrap();
    let sst = SharedStringTable::parse(sst_xml.as_bytes()).unwrap();
    assert_eq!(sst.len(), 2);
    assert_eq!(sst.get(0), Some("Repeat"));
    assert_eq!(sst.get(1), Some("Other"));

    // Both "Repeat" cells should map to SST index 0.
    let ws_xml = std::str::from_utf8(entries.get("xl/worksheets/sheet1.xml").unwrap()).unwrap();
    let ws = WorksheetXml::parse(ws_xml.as_bytes()).unwrap();
    assert_eq!(ws.rows[0].cells[0].value.as_deref(), Some("0"));
    assert_eq!(ws.rows[1].cells[0].value.as_deref(), Some("0"));
    assert_eq!(ws.rows[2].cells[0].value.as_deref(), Some("1"));
}

#[test]
fn rich_text_sst_bytes_pinned() {
    use crate::ooxml::shared_strings::RichTextRun;

    fn run(text: &str, bold: Option<bool>) -> RichTextRun {
        RichTextRun {
            text: text.to_string(),
            bold,
            ..Default::default()
        }
    }
    fn cell(reference: &str, val: &str, runs: Option<Vec<RichTextRun>>) -> Cell {
        Cell {
            reference: reference.to_string(),
            cell_type: CellType::SharedString,
            value: Some(val.to_string()),
            rich_text: runs,
            ..Default::default()
        }
    }
    fn sheet(name: &str, cells: Vec<Cell>) -> SheetData {
        SheetData {
            name: name.to_string(),
            worksheet: WorksheetXml {
                dimension: Some("A1".to_string()),
                rows: vec![Row {
                    index: 1,
                    cells,
                    ..Default::default()
                }],
                ..Default::default()
            },
            ..Default::default()
        }
    }

    let rich = || Some(vec![run("Bold", Some(true)), run(" tail", None)]);
    let wb = WorkbookData {
        sheets: vec![
            sheet(
                "One",
                vec![
                    cell("A1", "plain", None),
                    cell("B1", "Bold tail", rich()),
                    cell("C1", "Bold tail", None),
                ],
            ),
            sheet(
                "Two",
                vec![
                    cell("A1", "Bold tail", rich()),
                    cell("B1", "other", None),
                    cell("C1", "plain", None),
                ],
            ),
        ],
        ..Default::default()
    };
    let bytes = write_xlsx(&wb).expect("write_xlsx should succeed");
    let entries = read_zip_entries(&bytes, &ZipSecurityLimits::default()).unwrap();
    let sst = std::str::from_utf8(entries.get("xl/sharedStrings.xml").unwrap()).unwrap();
    assert_eq!(
        sst,
        concat!(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#,
            r#"<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="6" uniqueCount="4">"#,
            r#"<si><t>plain</t></si>"#,
            r#"<si><r><rPr><b/></rPr><t>Bold</t></r><r><t xml:space="preserve"> tail</t></r></si>"#,
            r#"<si><t>Bold tail</t></si>"#,
            r#"<si><t>other</t></si></sst>"#
        )
    );
}

/// A plain and a rich shared string with the same text, and two rich ones
/// with different runs, each keep their own `<si>` through write and read.
#[test]
fn same_text_plain_and_rich_cells_keep_their_own_si() {
    use crate::ooxml::shared_strings::RichTextRun;

    let run = |bold, italic| RichTextRun {
        text: "Total".to_string(),
        bold,
        italic,
        ..Default::default()
    };
    let cell = |reference: &str, rich_text: Option<Vec<RichTextRun>>| Cell {
        reference: reference.to_string(),
        cell_type: CellType::SharedString,
        value: Some("Total".to_string()),
        rich_text,
        ..Default::default()
    };
    let wb = WorkbookData {
        sheets: vec![SheetData {
            name: "S".to_string(),
            worksheet: WorksheetXml {
                rows: vec![Row {
                    index: 1,
                    cells: vec![
                        cell("A1", None),
                        cell("B1", Some(vec![run(Some(true), None)])),
                        cell("C1", Some(vec![run(None, Some(true))])),
                        cell("D1", None),
                    ],
                    ..Default::default()
                }],
                ..Default::default()
            },
            ..Default::default()
        }],
        ..Default::default()
    };
    let bytes = write_xlsx(&wb).expect("write_xlsx should succeed");
    let back = crate::reader::read_xlsx(&bytes).unwrap();
    let cells = &back.sheets[0].worksheet.rows[0].cells;
    let values: Vec<_> = cells.iter().map(|c| c.value.as_deref()).collect();
    assert_eq!(values, [Some("Total"); 4]);
    assert!(cells[0].rich_text.is_none(), "{:?}", cells[0].rich_text);
    assert!(cells[3].rich_text.is_none(), "{:?}", cells[3].rich_text);
    let b1 = cells[1].rich_text.as_deref().expect("B1 rich");
    assert_eq!((b1[0].bold, b1[0].italic), (Some(true), None));
    let c1 = cells[2].rich_text.as_deref().expect("C1 rich");
    assert_eq!((c1[0].bold, c1[0].italic), (None, Some(true)));
    assert_eq!(back.shared_strings.as_ref().unwrap().strings.len(), 3);
}

#[test]
fn test_workbook_xml_contains_sheet_names() {
    let wb = WorkbookData {
        sheets: vec![
            SheetData {
                name: "Sales".to_string(),
                worksheet: WorksheetXml::default(),
                ..Default::default()
            },
            SheetData {
                name: "Inventory".to_string(),
                worksheet: WorksheetXml::default(),
                ..Default::default()
            },
        ],
        ..Default::default()
    };

    let bytes = write_xlsx(&wb).expect("write_xlsx should succeed");
    let limits = ZipSecurityLimits::default();
    let entries = read_zip_entries(&bytes, &limits).unwrap();

    let wb_xml = std::str::from_utf8(entries.get("xl/workbook.xml").unwrap()).unwrap();
    let wb_parsed = WbXml::parse(wb_xml.as_bytes()).unwrap();
    assert_eq!(wb_parsed.sheets.len(), 2);
    assert_eq!(wb_parsed.sheets[0].name, "Sales");
    assert_eq!(wb_parsed.sheets[0].sheet_id, 1);
    assert_eq!(wb_parsed.sheets[0].r_id, "rId1");
    assert_eq!(wb_parsed.sheets[1].name, "Inventory");
    assert_eq!(wb_parsed.sheets[1].sheet_id, 2);
    assert_eq!(wb_parsed.sheets[1].r_id, "rId2");
}

#[test]
fn test_to_xml_with_sst_inline_remap() {
    let ws = WorksheetXml {
        dimension: Some("A1:B1".to_string()),
        rows: vec![Row {
            index: 1,
            cells: vec![
                Cell {
                    reference: "A1".to_string(),
                    cell_type: CellType::SharedString,
                    style_index: None,
                    value: Some("Alpha".to_string()),
                    ..Default::default()
                },
                Cell {
                    reference: "B1".to_string(),
                    cell_type: CellType::Number,
                    style_index: Some(1),
                    value: Some("99".to_string()),
                    ..Default::default()
                },
            ],
            ..Default::default()
        }],
        ..Default::default()
    };

    let mut sst = SharedStringTableBuilder::new();
    // Insert some dummy strings first to push "Alpha" to index 7.
    for i in 0..7 {
        sst.insert(&format!("dummy{i}"));
    }
    sst.insert("Alpha"); // index 7

    // Generate XML with inline SST remapping (no clone needed).
    let xml_bytes = ws
        .to_xml_with_sst(Some(&sst), &[], None, &[])
        .expect("to_xml_with_sst should succeed");
    let xml_str = std::str::from_utf8(&xml_bytes).unwrap();

    // The SharedString cell's <v> should contain the SST index "7", not "Alpha".
    assert!(
        xml_str.contains("<v>7</v>"),
        "SST cell should have index 7, got: {xml_str}"
    );
    // The Number cell should be unchanged.
    assert!(
        xml_str.contains("<v>99</v>"),
        "Number cell should remain 99"
    );
}

#[test]
fn test_content_types_has_correct_overrides() {
    let wb = minimal_workbook("Sheet1", Vec::new());
    let bytes = write_xlsx(&wb).expect("write_xlsx should succeed");

    let limits = ZipSecurityLimits::default();
    let entries = read_zip_entries(&bytes, &limits).unwrap();

    let ct_xml = std::str::from_utf8(entries.get("[Content_Types].xml").unwrap()).unwrap();
    let ct = ContentTypes::parse(ct_xml.as_bytes()).unwrap();

    assert!(ct.overrides.contains_key("/xl/workbook.xml"));
    assert!(ct.overrides.contains_key("/xl/sharedStrings.xml"));
    assert!(ct.overrides.contains_key("/xl/styles.xml"));
    assert!(ct.overrides.contains_key("/xl/worksheets/sheet1.xml"));
}

#[test]
fn test_relationships_are_correct() {
    let wb = WorkbookData {
        sheets: vec![
            SheetData {
                name: "S1".to_string(),
                worksheet: WorksheetXml::default(),
                ..Default::default()
            },
            SheetData {
                name: "S2".to_string(),
                worksheet: WorksheetXml::default(),
                ..Default::default()
            },
        ],
        ..Default::default()
    };

    let bytes = write_xlsx(&wb).expect("write_xlsx should succeed");
    let limits = ZipSecurityLimits::default();
    let entries = read_zip_entries(&bytes, &limits).unwrap();

    // Root rels should point to xl/workbook.xml.
    let root_rels_xml = std::str::from_utf8(entries.get("_rels/.rels").unwrap()).unwrap();
    let root_rels = Relationships::parse(root_rels_xml.as_bytes()).unwrap();
    assert_eq!(root_rels.len(), 1);
    assert_eq!(root_rels.iter().next().unwrap().target, "xl/workbook.xml");

    // Workbook rels: 2 worksheets + sharedStrings + styles + theme = 5.
    let wb_rels_xml =
        std::str::from_utf8(entries.get("xl/_rels/workbook.xml.rels").unwrap()).unwrap();
    let wb_rels = Relationships::parse(wb_rels_xml.as_bytes()).unwrap();
    assert_eq!(wb_rels.len(), 5);
    assert_eq!(
        wb_rels.get_by_id("rId1").unwrap().target,
        "worksheets/sheet1.xml"
    );
    assert_eq!(
        wb_rels.get_by_id("rId2").unwrap().target,
        "worksheets/sheet2.xml"
    );
    assert_eq!(
        wb_rels.get_by_id("rId3").unwrap().target,
        "sharedStrings.xml"
    );
    assert_eq!(wb_rels.get_by_id("rId4").unwrap().target, "styles.xml");
    assert_eq!(
        wb_rels.get_by_id("rId5").unwrap().target,
        "theme/theme1.xml"
    );
}

#[test]
fn test_write_then_read_roundtrip() {
    // Write a workbook with mixed cell types, then read it back via the
    // reader module and verify the data survived the round-trip.
    let rows = vec![Row {
        index: 1,
        cells: vec![
            Cell {
                reference: "A1".to_string(),
                cell_type: CellType::SharedString,
                style_index: None,
                value: Some("Header".to_string()),
                ..Default::default()
            },
            Cell {
                reference: "B1".to_string(),
                cell_type: CellType::Number,
                style_index: None,
                value: Some("100".to_string()),
                ..Default::default()
            },
            Cell {
                reference: "C1".to_string(),
                cell_type: CellType::Boolean,
                style_index: None,
                value: Some("1".to_string()),
                ..Default::default()
            },
        ],
        ..Default::default()
    }];

    let wb = minimal_workbook("RoundTrip", rows);
    let xlsx_bytes = write_xlsx(&wb).expect("write_xlsx should succeed");

    // Read back via the reader module.
    let read_wb =
        crate::reader::read_xlsx(&xlsx_bytes).expect("read_xlsx should succeed on writer output");

    assert_eq!(read_wb.sheets.len(), 1);
    assert_eq!(read_wb.sheets[0].name, "RoundTrip");
    assert_eq!(read_wb.date_system, DateSystem::Date1900);

    // Shared strings table should have 1 entry: "Header".
    let sst = read_wb
        .shared_strings
        .as_ref()
        .expect("shared_strings should be Some");
    assert_eq!(sst.len(), 1);
    assert_eq!(sst.get(0), Some("Header"));

    // Worksheet should have 1 row with 3 cells.
    let ws = &read_wb.sheets[0].worksheet;
    assert_eq!(ws.rows.len(), 1);
    assert_eq!(ws.rows[0].cells.len(), 3);

    // A1: SharedString resolved from SST index "0" -> "Header".
    assert_eq!(ws.rows[0].cells[0].cell_type, CellType::SharedString);
    assert_eq!(ws.rows[0].cells[0].value.as_deref(), Some("Header"));

    // B1: Number "100".
    assert_eq!(ws.rows[0].cells[1].cell_type, CellType::Number);
    assert_eq!(ws.rows[0].cells[1].value.as_deref(), Some("100"));

    // C1: Boolean "1".
    assert_eq!(ws.rows[0].cells[2].cell_type, CellType::Boolean);
    assert_eq!(ws.rows[0].cells[2].value.as_deref(), Some("1"));
}

#[test]
fn test_formula_roundtrip() {
    // Write a workbook with formula cells, read it back, and verify
    // formulas survive the full XLSX round-trip.
    let rows = vec![
        Row {
            index: 1,
            cells: vec![Cell {
                reference: "A1".to_string(),
                cell_type: CellType::Number,
                style_index: None,
                value: Some("10".to_string()),
                ..Default::default()
            }],
            ..Default::default()
        },
        Row {
            index: 2,
            cells: vec![Cell {
                reference: "A2".to_string(),
                cell_type: CellType::Number,
                style_index: None,
                value: Some("20".to_string()),
                ..Default::default()
            }],
            ..Default::default()
        },
        Row {
            index: 3,
            cells: vec![Cell {
                reference: "A3".to_string(),
                cell_type: CellType::Number,
                value: Some("30".to_string()),
                formula: Some("SUM(A1:A2)".to_string()),
                ..Default::default()
            }],
            ..Default::default()
        },
    ];

    let wb = minimal_workbook("Formulas", rows);
    let xlsx_bytes = write_xlsx(&wb).expect("write_xlsx should succeed");

    // Read back via the reader module.
    let read_wb =
        crate::reader::read_xlsx(&xlsx_bytes).expect("read_xlsx should succeed on writer output");

    assert_eq!(read_wb.sheets.len(), 1);
    assert_eq!(read_wb.sheets[0].name, "Formulas");

    let ws = &read_wb.sheets[0].worksheet;
    assert_eq!(ws.rows.len(), 3);

    // A1: plain number, no formula.
    assert_eq!(ws.rows[0].cells[0].reference, "A1");
    assert_eq!(ws.rows[0].cells[0].cell_type, CellType::Number);
    assert_eq!(ws.rows[0].cells[0].value.as_deref(), Some("10"));
    assert_eq!(ws.rows[0].cells[0].formula, None);

    // A2: plain number, no formula.
    assert_eq!(ws.rows[1].cells[0].reference, "A2");
    assert_eq!(ws.rows[1].cells[0].cell_type, CellType::Number);
    assert_eq!(ws.rows[1].cells[0].value.as_deref(), Some("20"));
    assert_eq!(ws.rows[1].cells[0].formula, None);

    // A3: formula cell with cached value.
    assert_eq!(ws.rows[2].cells[0].reference, "A3");
    assert_eq!(ws.rows[2].cells[0].formula, Some("SUM(A1:A2)".to_string()));
    assert_eq!(ws.rows[2].cells[0].value.as_deref(), Some("30"));
}

#[test]
fn test_formula_str_roundtrip() {
    // Test formula with t="str" type (formula returning a string result).
    let rows = vec![Row {
        index: 1,
        cells: vec![Cell {
            reference: "B1".to_string(),
            cell_type: CellType::FormulaStr,
            value: Some("100".to_string()),
            formula: Some("A1*2".to_string()),
            ..Default::default()
        }],
        ..Default::default()
    }];

    let wb = minimal_workbook("FormulaStr", rows);
    let xlsx_bytes = write_xlsx(&wb).expect("write_xlsx should succeed");

    let read_wb =
        crate::reader::read_xlsx(&xlsx_bytes).expect("read_xlsx should succeed on writer output");

    let ws = &read_wb.sheets[0].worksheet;
    assert_eq!(ws.rows.len(), 1);
    assert_eq!(ws.rows[0].cells[0].reference, "B1");
    assert_eq!(ws.rows[0].cells[0].cell_type, CellType::FormulaStr);
    assert_eq!(ws.rows[0].cells[0].formula, Some("A1*2".to_string()));
    assert_eq!(ws.rows[0].cells[0].value.as_deref(), Some("100"));
}

#[test]
fn test_write_workbook_with_tables() {
    use crate::ooxml::content_types::ContentTypes;
    use crate::ooxml::tables::{TableColumn, TableDefinition, TableStyleInfo};

    let rows = vec![
        Row {
            index: 1,
            cells: vec![
                Cell {
                    reference: "A1".to_string(),
                    cell_type: CellType::SharedString,
                    style_index: None,
                    value: Some("Name".to_string()),
                    ..Default::default()
                },
                Cell {
                    reference: "B1".to_string(),
                    cell_type: CellType::SharedString,
                    style_index: None,
                    value: Some("Age".to_string()),
                    ..Default::default()
                },
            ],
            ..Default::default()
        },
        Row {
            index: 2,
            cells: vec![
                Cell {
                    reference: "A2".to_string(),
                    cell_type: CellType::SharedString,
                    style_index: None,
                    value: Some("Alice".to_string()),
                    ..Default::default()
                },
                Cell {
                    reference: "B2".to_string(),
                    cell_type: CellType::Number,
                    style_index: None,
                    value: Some("30".to_string()),
                    ..Default::default()
                },
            ],
            ..Default::default()
        },
    ];

    let table = TableDefinition {
        id: 1,
        name: Some("Table1".into()),
        display_name: "Table1".into(),
        ref_range: "A1:B2".into(),
        header_row_count: 1,
        totals_row_count: 0,
        totals_row_shown: false,
        columns: vec![
            TableColumn {
                id: 1,
                name: "Name".into(),
                ..Default::default()
            },
            TableColumn {
                id: 2,
                name: "Age".into(),
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
        auto_filter_ref: Some("A1:B2".into()),
    };

    let wb = WorkbookData {
        sheets: vec![SheetData {
            name: "Sheet1".to_string(),
            worksheet: WorksheetXml {
                dimension: Some("A1:B2".to_string()),
                rows,
                tables: vec![table],
                ..Default::default()
            },
            ..Default::default()
        }],
        ..Default::default()
    };

    let bytes = write_xlsx(&wb).expect("write_xlsx should succeed");

    // Verify it is a valid ZIP with table entries.
    let limits = ZipSecurityLimits::default();
    let zip_entries = read_zip_entries(&bytes, &limits).expect("should be a valid ZIP");

    // Table XML file should exist.
    assert!(
        zip_entries.contains_key("xl/tables/table1.xml"),
        "should contain xl/tables/table1.xml"
    );

    // Worksheet .rels should reference the table.
    let rels_xml = std::str::from_utf8(
        zip_entries
            .get("xl/worksheets/_rels/sheet1.xml.rels")
            .expect("should have worksheet rels"),
    )
    .unwrap();
    assert!(
        rels_xml.contains("tables/table1.xml"),
        "rels should reference table: {rels_xml}"
    );

    // Content types should include the table override.
    let ct_xml = std::str::from_utf8(zip_entries.get("[Content_Types].xml").unwrap()).unwrap();
    let ct = ContentTypes::parse(ct_xml.as_bytes()).unwrap();
    assert!(
        ct.overrides.contains_key("/xl/tables/table1.xml"),
        "content types should have table override"
    );

    // Worksheet XML should contain <tableParts>.
    let ws_xml = std::str::from_utf8(zip_entries.get("xl/worksheets/sheet1.xml").unwrap()).unwrap();
    assert!(
        ws_xml.contains("<tableParts"),
        "worksheet should contain tableParts: {ws_xml}"
    );
    assert!(
        ws_xml.contains("r:id=\"rId1\""),
        "tablePart should reference rId1: {ws_xml}"
    );

    // Table XML should parse back correctly.
    let table_data = zip_entries.get("xl/tables/table1.xml").unwrap();
    let parsed_table = TableDefinition::parse(table_data).unwrap();
    assert_eq!(parsed_table.display_name, "Table1");
    assert_eq!(parsed_table.columns.len(), 2);
}

#[test]
fn test_write_workbook_with_tables_across_sheets() {
    use crate::ooxml::tables::{TableColumn, TableDefinition, TableStyleInfo};

    let make_table = |id: u32, name: &str| TableDefinition {
        id,
        name: Some(name.into()),
        display_name: name.into(),
        ref_range: "A1:B2".into(),
        header_row_count: 1,
        totals_row_count: 0,
        totals_row_shown: true,
        columns: vec![TableColumn {
            id: 1,
            name: "Col1".into(),
            ..Default::default()
        }],
        style_info: Some(TableStyleInfo {
            name: Some("TableStyleLight1".into()),
            show_first_column: false,
            show_last_column: false,
            show_row_stripes: true,
            show_column_stripes: false,
        }),
        auto_filter_ref: None,
    };

    let make_sheet = |name: &str, tables: Vec<TableDefinition>| SheetData {
        name: name.to_string(),
        worksheet: WorksheetXml {
            tables,
            ..Default::default()
        },
        ..Default::default()
    };

    let wb = WorkbookData {
        sheets: vec![
            make_sheet("Sheet1", vec![make_table(1, "T1"), make_table(2, "T2")]),
            make_sheet("Sheet2", vec![make_table(3, "T3")]),
        ],
        ..Default::default()
    };

    let bytes = write_xlsx(&wb).expect("write_xlsx should succeed");
    let limits = ZipSecurityLimits::default();
    let zip_entries = read_zip_entries(&bytes, &limits).unwrap();

    // Global table numbering: table1, table2 on Sheet1; table3 on Sheet2.
    assert!(zip_entries.contains_key("xl/tables/table1.xml"));
    assert!(zip_entries.contains_key("xl/tables/table2.xml"));
    assert!(zip_entries.contains_key("xl/tables/table3.xml"));

    // Sheet1 rels should have 2 table references.
    let rels1 = std::str::from_utf8(
        zip_entries
            .get("xl/worksheets/_rels/sheet1.xml.rels")
            .unwrap(),
    )
    .unwrap();
    assert!(rels1.contains("tables/table1.xml"));
    assert!(rels1.contains("tables/table2.xml"));

    // Sheet2 rels should have 1 table reference.
    let rels2 = std::str::from_utf8(
        zip_entries
            .get("xl/worksheets/_rels/sheet2.xml.rels")
            .unwrap(),
    )
    .unwrap();
    assert!(rels2.contains("tables/table3.xml"));

    // Sheet1 worksheet should have tableParts count="2".
    let ws1 = std::str::from_utf8(zip_entries.get("xl/worksheets/sheet1.xml").unwrap()).unwrap();
    assert!(
        ws1.contains("count=\"2\""),
        "sheet1 tableParts count should be 2: {ws1}"
    );

    // Sheet2 worksheet should have tableParts count="1".
    let ws2 = std::str::from_utf8(zip_entries.get("xl/worksheets/sheet2.xml").unwrap()).unwrap();
    assert!(
        ws2.contains("count=\"1\""),
        "sheet2 tableParts count should be 1: {ws2}"
    );
}

#[test]
fn test_sheet_state_hidden_roundtrip() {
    let wb = WorkbookData {
        sheets: vec![
            SheetData {
                name: "Visible".into(),
                worksheet: WorksheetXml::default(),
                ..Default::default()
            },
            SheetData {
                name: "Hidden".into(),
                state: Some("hidden".into()),
                ..Default::default()
            },
            SheetData {
                name: "VeryHidden".into(),
                state: Some("veryHidden".into()),
                ..Default::default()
            },
        ],
        ..Default::default()
    };

    let bytes = write_xlsx(&wb).expect("write_xlsx should succeed");
    let wb2 = crate::reader::read_xlsx(&bytes).expect("read_xlsx should succeed");

    assert_eq!(wb2.sheets.len(), 3);

    // Visible sheet: state should be None (omitted).
    assert!(
        wb2.sheets[0].state.is_none(),
        "visible sheet state should be None"
    );
    assert_eq!(wb2.sheets[0].name, "Visible");

    // Hidden sheet: state should be Some("hidden").
    assert_eq!(wb2.sheets[1].state.as_deref(), Some("hidden"));
    assert_eq!(wb2.sheets[1].name, "Hidden");

    // VeryHidden sheet: state should be Some("veryHidden").
    assert_eq!(wb2.sheets[2].state.as_deref(), Some("veryHidden"));
    assert_eq!(wb2.sheets[2].name, "VeryHidden");
}

#[test]
fn test_sheet_state_hidden_json_roundtrip() {
    let wb = WorkbookData {
        sheets: vec![
            SheetData {
                name: "Visible".into(),
                worksheet: WorksheetXml::default(),
                ..Default::default()
            },
            SheetData {
                name: "Hidden".into(),
                state: Some("hidden".into()),
                ..Default::default()
            },
        ],
        ..Default::default()
    };

    let bytes = write_xlsx(&wb).expect("write_xlsx should succeed");

    // Use the streaming JSON reader path.
    let json = crate::reader::read_xlsx_json(&bytes).expect("read_xlsx_json should succeed");
    let wb2: WorkbookData = serde_json::from_str(&json).expect("JSON parse should succeed");

    assert_eq!(wb2.sheets.len(), 2);
    assert!(wb2.sheets[0].state.is_none());
    assert_eq!(wb2.sheets[1].state.as_deref(), Some("hidden"));
}

#[test]
fn test_external_link_content_type() {
    use crate::ooxml::content_types::{CT_EXTERNAL_LINK, ContentTypes};

    let mut wb = minimal_workbook("Sheet1", Vec::new());
    wb.preserved_entries.insert(
        "xl/externalLinks/externalLink1.xml".to_string(),
        b"<externalLink/>".to_vec(),
    );

    let bytes = write_xlsx(&wb).expect("write_xlsx should succeed");

    // Read the ZIP and inspect [Content_Types].xml for the external link override.
    let limits = ZipSecurityLimits::default();
    let entries = read_zip_entries(&bytes, &limits).expect("should be a valid ZIP");
    let ct_bytes = entries
        .get("[Content_Types].xml")
        .expect("missing [Content_Types].xml");
    let ct = ContentTypes::parse(ct_bytes).expect("parse content types");

    assert_eq!(
        ct.overrides.get("/xl/externalLinks/externalLink1.xml"),
        Some(&CT_EXTERNAL_LINK.to_string()),
        "external link content type override should be present"
    );

    // Verify roundtrip: the preserved entry survives write -> read.
    let wb2 = crate::reader::read_xlsx(&bytes).expect("read_xlsx should succeed");
    assert!(
        wb2.preserved_entries
            .contains_key("xl/externalLinks/externalLink1.xml")
    );
}

#[test]
fn test_custom_xml_content_type() {
    use crate::ooxml::content_types::{CT_CUSTOM_XML_PROPS, CT_XML, ContentTypes};

    let mut wb = minimal_workbook("Sheet1", Vec::new());
    wb.preserved_entries
        .insert("customXml/item1.xml".to_string(), b"<root/>".to_vec());
    wb.preserved_entries.insert(
        "customXml/itemProps1.xml".to_string(),
        b"<ds:datastoreItem/>".to_vec(),
    );

    let bytes = write_xlsx(&wb).expect("write_xlsx should succeed");

    // Read the ZIP and inspect [Content_Types].xml for both custom XML overrides.
    let limits = ZipSecurityLimits::default();
    let entries = read_zip_entries(&bytes, &limits).expect("should be a valid ZIP");
    let ct_bytes = entries
        .get("[Content_Types].xml")
        .expect("missing [Content_Types].xml");
    let ct = ContentTypes::parse(ct_bytes).expect("parse content types");

    assert_eq!(
        ct.overrides.get("/customXml/item1.xml"),
        Some(&CT_XML.to_string()),
        "custom XML item content type override should be present"
    );
    assert_eq!(
        ct.overrides.get("/customXml/itemProps1.xml"),
        Some(&CT_CUSTOM_XML_PROPS.to_string()),
        "custom XML properties content type override should be present"
    );

    // Verify roundtrip: both preserved entries survive write -> read.
    let wb2 = crate::reader::read_xlsx(&bytes).expect("read_xlsx should succeed");
    assert!(wb2.preserved_entries.contains_key("customXml/item1.xml"));
    assert!(
        wb2.preserved_entries
            .contains_key("customXml/itemProps1.xml")
    );
}

/// Helper: build a single-sheet workbook with chart(s).
fn make_workbook_with_charts(charts: Vec<crate::ooxml::charts::WorksheetChart>) -> WorkbookData {
    WorkbookData {
        sheets: vec![SheetData {
            name: "Sheet1".to_string(),
            worksheet: WorksheetXml {
                dimension: Some("A1:B5".to_string()),
                rows: vec![Row {
                    index: 1,
                    cells: vec![
                        Cell {
                            reference: "A1".to_string(),
                            cell_type: CellType::SharedString,
                            value: Some("Category".to_string()),
                            ..Default::default()
                        },
                        Cell {
                            reference: "B1".to_string(),
                            cell_type: CellType::SharedString,
                            value: Some("Value".to_string()),
                            ..Default::default()
                        },
                    ],
                    ..Default::default()
                }],
                charts,
                ..Default::default()
            },
            ..Default::default()
        }],
        ..Default::default()
    }
}

#[test]
fn test_write_workbook_with_charts() {
    use crate::ooxml::charts::*;

    let chart = WorksheetChart {
        chart: ChartData {
            title: Some(ChartTitle {
                text: "Sales".into(),
                overlay: false,
                font_size: None,
                bold: None,
                color: None,
            }),
            series: vec![ChartSeries {
                name: Some("Revenue".into()),
                cat_ref: Some("Sheet1!$A$2:$A$5".into()),
                val_ref: "Sheet1!$B$2:$B$5".into(),
                ..Default::default()
            }],
            cat_axis: Some(ChartAxis {
                cross_ax: 1,
                ..Default::default()
            }),
            val_axis: Some(ChartAxis {
                id: 1,
                major_gridlines: true,
                ..Default::default()
            }),
            grouping: Some(ChartGrouping::Clustered),
            ..Default::default()
        },
        anchor: ChartAnchor {
            from_col: 0,
            from_row: 0,
            from_col_off: 0,
            from_row_off: 0,
            to_col: 8,
            to_row: 15,
            to_col_off: 0,
            to_row_off: 0,
            ext_cx: None,
            ext_cy: None,
            ..Default::default()
        },
    };

    let wb = make_workbook_with_charts(vec![chart]);
    let bytes = write_xlsx(&wb).unwrap();

    // Parse ZIP and verify contents.
    let limits = ZipSecurityLimits::default();
    let zip_entries = read_zip_entries(&bytes, &limits).unwrap();

    // Should contain chart XML.
    assert!(
        zip_entries.contains_key("xl/charts/chart1.xml"),
        "should contain chart XML"
    );

    // Should contain drawing XML.
    assert!(
        zip_entries.contains_key("xl/drawings/drawing1.xml"),
        "should contain drawing XML"
    );

    // Should contain drawing rels.
    assert!(
        zip_entries.contains_key("xl/drawings/_rels/drawing1.xml.rels"),
        "should contain drawing rels"
    );

    // Check content types.
    let ct_xml = std::str::from_utf8(zip_entries.get("[Content_Types].xml").unwrap()).unwrap();
    assert!(
        ct_xml.contains("drawingml.chart+xml"),
        "content types should include chart type"
    );
    assert!(
        ct_xml.contains("drawing+xml"),
        "content types should include drawing type"
    );

    // Check worksheet has drawing reference.
    let ws_xml = std::str::from_utf8(zip_entries.get("xl/worksheets/sheet1.xml").unwrap()).unwrap();
    assert!(
        ws_xml.contains("<drawing"),
        "worksheet should reference drawing"
    );

    // Check drawing XML has anchor.
    let drawing_xml =
        std::str::from_utf8(zip_entries.get("xl/drawings/drawing1.xml").unwrap()).unwrap();
    assert!(
        drawing_xml.contains("<xdr:twoCellAnchor>"),
        "drawing should have anchor"
    );
    assert!(
        drawing_xml.contains("<xdr:col>8</xdr:col>"),
        "drawing anchor should reference col 8"
    );

    // Check chart XML.
    let chart_xml = std::str::from_utf8(zip_entries.get("xl/charts/chart1.xml").unwrap()).unwrap();
    assert!(
        chart_xml.contains("<c:barChart>"),
        "chart should be bar chart"
    );

    // Check drawing rels references the chart.
    let drawing_rels = std::str::from_utf8(
        zip_entries
            .get("xl/drawings/_rels/drawing1.xml.rels")
            .unwrap(),
    )
    .unwrap();
    assert!(
        drawing_rels.contains("chart1.xml"),
        "drawing rels should reference chart1.xml"
    );

    // Check worksheet rels references the drawing.
    let ws_rels = std::str::from_utf8(
        zip_entries
            .get("xl/worksheets/_rels/sheet1.xml.rels")
            .unwrap(),
    )
    .unwrap();
    assert!(
        ws_rels.contains("drawing1.xml"),
        "worksheet rels should reference drawing1.xml"
    );
}

#[test]
fn test_multiple_charts_on_one_sheet() {
    use crate::ooxml::charts::*;

    let make_chart = |chart_type, from_col, to_col| WorksheetChart {
        chart: ChartData {
            chart_type,
            series: vec![ChartSeries {
                val_ref: "Sheet1!$A$1:$A$5".into(),
                ..Default::default()
            }],
            ..Default::default()
        },
        anchor: ChartAnchor {
            from_col,
            from_row: 0,
            from_col_off: 0,
            from_row_off: 0,
            to_col,
            to_row: 10,
            to_col_off: 0,
            to_row_off: 0,
            ext_cx: None,
            ext_cy: None,
            ..Default::default()
        },
    };

    let wb = make_workbook_with_charts(vec![
        make_chart(ChartType::Pie, 0, 5),
        make_chart(ChartType::Line, 6, 12),
    ]);
    let bytes = write_xlsx(&wb).unwrap();

    let limits = ZipSecurityLimits::default();
    let zip_entries = read_zip_entries(&bytes, &limits).unwrap();

    // Should have two chart XMLs.
    assert!(zip_entries.contains_key("xl/charts/chart1.xml"));
    assert!(zip_entries.contains_key("xl/charts/chart2.xml"));

    // Single drawing XML with two anchors.
    let drawing_xml =
        std::str::from_utf8(zip_entries.get("xl/drawings/drawing1.xml").unwrap()).unwrap();
    let anchor_count = drawing_xml.matches("<xdr:twoCellAnchor>").count();
    assert_eq!(anchor_count, 2, "should have two chart anchors in drawing");

    // Drawing rels should reference both charts.
    let drawing_rels = std::str::from_utf8(
        zip_entries
            .get("xl/drawings/_rels/drawing1.xml.rels")
            .unwrap(),
    )
    .unwrap();
    assert!(drawing_rels.contains("chart1.xml"));
    assert!(drawing_rels.contains("chart2.xml"));

    // Verify chart types.
    let chart1 = std::str::from_utf8(zip_entries.get("xl/charts/chart1.xml").unwrap()).unwrap();
    assert!(chart1.contains("<c:pieChart>"), "first chart should be pie");
    let chart2 = std::str::from_utf8(zip_entries.get("xl/charts/chart2.xml").unwrap()).unwrap();
    assert!(
        chart2.contains("<c:lineChart>"),
        "second chart should be line"
    );
}

// -----------------------------------------------------------------------
// Drawing merge / collision tests
// -----------------------------------------------------------------------

#[test]
fn test_extract_anchors_from_drawing() {
    let drawing = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <xdr:twoCellAnchor editAs="oneCell">
    <xdr:from><xdr:col>0</xdr:col></xdr:from>
    <xdr:to><xdr:col>3</xdr:col></xdr:to>
    <xdr:pic>
      <xdr:blipFill><a:blip r:embed="rId1"/></xdr:blipFill>
    </xdr:pic>
    <xdr:clientData/>
  </xdr:twoCellAnchor>
  <xdr:twoCellAnchor editAs="oneCell">
    <xdr:from><xdr:col>4</xdr:col></xdr:from>
    <xdr:to><xdr:col>7</xdr:col></xdr:to>
    <xdr:pic>
      <xdr:blipFill><a:blip r:embed="rId2"/></xdr:blipFill>
    </xdr:pic>
    <xdr:clientData/>
  </xdr:twoCellAnchor>
</xdr:wsDr>"#;

    let kept = kept_drawing_children(drawing, "xl/drawings/drawing1.xml").unwrap();
    let anchors = kept.children;
    assert_eq!(anchors.len(), 2, "should extract 2 anchor blocks");
    assert_eq!(kept.namespaces.len(), 3);
    assert!(
        anchors[0].contains("rId1"),
        "first anchor should reference rId1"
    );
    assert!(
        anchors[1].contains("rId2"),
        "second anchor should reference rId2"
    );
    assert!(
        anchors[0].contains("<xdr:pic>"),
        "first anchor should contain pic"
    );
    assert!(
        anchors[1].contains("<xdr:pic>"),
        "second anchor should contain pic"
    );
}

#[test]
fn test_extract_anchors_empty_drawing() {
    let drawing = r#"<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"/>"#;
    let kept = kept_drawing_children(drawing, "xl/drawings/drawing1.xml").unwrap();
    assert!(
        kept.children.is_empty(),
        "empty drawing should have no anchors"
    );
}

#[test]
fn malformed_preserved_drawing_names_its_path() {
    let err = kept_drawing_children("<xdr:wsDr><a></b></xdr:wsDr>", "xl/drawings/drawing7.xml")
        .err()
        .unwrap();
    assert!(matches!(err, ModernXlsxError::XmlParse(_)), "{err:?}");
    assert!(
        err.to_string().contains("xl/drawings/drawing7.xml"),
        "{err}"
    );
}

#[test]
fn rel_id_remap_is_one_pass() {
    let remap: HashMap<&str, &str> = [("rId1", "rId2"), ("rId2", "rId3")].into();
    let mut out = String::new();
    push_remapped_rel_ids(
        &mut out,
        r#"<a:blip r:embed="rId1" r:link='rId2'/><a:hlinkClick r:id="rId2"/><xdr:pic xr:id="rId1" r:embed="rId9">rId1</xdr:pic>"#,
        &remap,
    );
    assert_eq!(
        out,
        r#"<a:blip r:embed="rId2" r:link='rId3'/><a:hlinkClick r:id="rId3"/><xdr:pic xr:id="rId1" r:embed="rId9">rId1</xdr:pic>"#
    );
}

#[test]
fn allocation_skips_used_names_case_insensitively() {
    let mut names = PartNames::new(["xl/charts/chart1.xml", "xl/comments2.xml"]);
    assert_eq!(
        names.alloc("xl/charts/Chart", ".xml"),
        "xl/charts/Chart2.xml"
    );
    assert_eq!(
        names.alloc("xl/charts/chart", ".xml"),
        "xl/charts/chart3.xml"
    );
    assert_eq!(
        names.alloc_preferring("xl/comments", 3, ".xml"),
        "xl/comments3.xml"
    );
    assert_eq!(
        names.alloc_preferring("xl/comments", 2, ".xml"),
        "xl/comments1.xml"
    );
}

#[test]
fn test_chart_and_image_drawing_merge() {
    use crate::ooxml::charts::*;

    // Build a workbook with a chart AND image drawing XML in preserved entries.
    let mut wb = make_workbook_with_charts(vec![WorksheetChart {
        chart: ChartData {
            title: Some(ChartTitle {
                text: "Sales".into(),
                overlay: false,
                font_size: None,
                bold: None,
                color: None,
            }),
            series: vec![ChartSeries {
                val_ref: "Sheet1!$A$1:$A$5".into(),
                ..Default::default()
            }],
            ..Default::default()
        },
        anchor: ChartAnchor {
            from_col: 4,
            from_row: 0,
            from_col_off: 0,
            from_row_off: 0,
            to_col: 12,
            to_row: 15,
            to_col_off: 0,
            to_row_off: 0,
            ext_cx: None,
            ext_cy: None,
            ..Default::default()
        },
    }]);

    // Simulate addImage: put image drawing XML in preserved entries.
    let image_drawing = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <xdr:twoCellAnchor editAs="oneCell">
    <xdr:from>
      <xdr:col>0</xdr:col>
      <xdr:colOff>0</xdr:colOff>
      <xdr:row>5</xdr:row>
      <xdr:rowOff>0</xdr:rowOff>
    </xdr:from>
    <xdr:to>
      <xdr:col>3</xdr:col>
      <xdr:colOff>0</xdr:colOff>
      <xdr:row>12</xdr:row>
      <xdr:rowOff>0</xdr:rowOff>
    </xdr:to>
    <xdr:pic>
      <xdr:nvPicPr>
        <xdr:cNvPr id="2" name="Barcode 1"/>
        <xdr:cNvPicPr>
          <a:picLocks noChangeAspect="1"/>
        </xdr:cNvPicPr>
      </xdr:nvPicPr>
      <xdr:blipFill>
        <a:blip r:embed="rId1"/>
        <a:stretch>
          <a:fillRect/>
        </a:stretch>
      </xdr:blipFill>
      <xdr:spPr>
        <a:xfrm>
          <a:off x="0" y="0"/>
          <a:ext cx="0" cy="0"/>
        </a:xfrm>
        <a:prstGeom prst="rect">
          <a:avLst/>
        </a:prstGeom>
      </xdr:spPr>
    </xdr:pic>
    <xdr:clientData/>
  </xdr:twoCellAnchor>
</xdr:wsDr>"#;

    let image_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>
</Relationships>"#;

    // Simulate the sheet rels that addImage creates.
    let sheet_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdDrawing1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/></Relationships>"#;

    // Fake PNG bytes.
    let fake_png = vec![0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A];

    wb.preserved_entries.insert(
        "xl/drawings/drawing1.xml".to_string(),
        image_drawing.as_bytes().to_vec(),
    );
    wb.preserved_entries.insert(
        "xl/drawings/_rels/drawing1.xml.rels".to_string(),
        image_rels.as_bytes().to_vec(),
    );
    wb.preserved_entries.insert(
        "xl/worksheets/_rels/sheet1.xml.rels".to_string(),
        sheet_rels.as_bytes().to_vec(),
    );
    wb.preserved_entries
        .insert("xl/media/image1.png".to_string(), fake_png.clone());

    // Write the workbook.
    let bytes = write_xlsx(&wb).unwrap();

    let limits = ZipSecurityLimits::default();
    let zip_entries = read_zip_entries(&bytes, &limits).unwrap();

    // The merged drawing XML should contain BOTH chart and image anchors.
    let drawing_xml =
        std::str::from_utf8(zip_entries.get("xl/drawings/drawing1.xml").unwrap()).unwrap();

    // Chart anchor (graphicFrame).
    assert!(
        drawing_xml.contains("<xdr:graphicFrame"),
        "merged drawing should contain chart graphicFrame:\n{drawing_xml}"
    );
    // Image anchor (pic).
    assert!(
        drawing_xml.contains("<xdr:pic>"),
        "merged drawing should contain image pic:\n{drawing_xml}"
    );

    // Should have exactly 2 twoCellAnchors (1 chart + 1 image).
    let anchor_count = drawing_xml.matches("<xdr:twoCellAnchor").count();
    assert_eq!(
        anchor_count, 2,
        "should have 2 anchors (1 chart + 1 image):\n{drawing_xml}"
    );

    // The drawing rels should reference both the chart AND the image.
    let drawing_rels = std::str::from_utf8(
        zip_entries
            .get("xl/drawings/_rels/drawing1.xml.rels")
            .unwrap(),
    )
    .unwrap();
    assert!(
        drawing_rels.contains("chart1.xml"),
        "drawing rels should reference chart:\n{drawing_rels}"
    );
    assert!(
        drawing_rels.contains("image1.png"),
        "drawing rels should reference image:\n{drawing_rels}"
    );

    // The image rId should be remapped (chart uses rId1, so image should be rId2).
    assert!(
        drawing_xml.contains("r:embed=\"rId2\""),
        "image rId should be remapped to rId2:\n{drawing_xml}"
    );

    // Image media bytes should be present.
    assert_eq!(
        zip_entries.get("xl/media/image1.png").unwrap(),
        &fake_png,
        "image media should survive"
    );

    // Worksheet .rels should NOT have duplicate drawing relationships.
    let ws_rels = std::str::from_utf8(
        zip_entries
            .get("xl/worksheets/_rels/sheet1.xml.rels")
            .unwrap(),
    )
    .unwrap();
    let drawing_rel_count = ws_rels.matches("relationships/drawing").count();
    assert_eq!(
        drawing_rel_count, 1,
        "worksheet rels should have exactly 1 drawing relationship:\n{ws_rels}"
    );
}

#[test]
fn test_chart_only_no_regression() {
    // Ensure chart-only workbooks still work correctly after the merge changes.
    use crate::ooxml::charts::*;

    let wb = make_workbook_with_charts(vec![WorksheetChart {
        chart: ChartData {
            chart_type: ChartType::Line,
            title: Some(ChartTitle {
                text: "Trend".into(),
                overlay: false,
                font_size: None,
                bold: None,
                color: None,
            }),
            series: vec![ChartSeries {
                val_ref: "Sheet1!$A$1:$A$5".into(),
                ..Default::default()
            }],
            ..Default::default()
        },
        anchor: ChartAnchor {
            from_col: 0,
            from_row: 0,
            from_col_off: 0,
            from_row_off: 0,
            to_col: 8,
            to_row: 10,
            to_col_off: 0,
            to_row_off: 0,
            ext_cx: None,
            ext_cy: None,
            ..Default::default()
        },
    }]);

    let bytes = write_xlsx(&wb).unwrap();

    let limits = ZipSecurityLimits::default();
    let zip_entries = read_zip_entries(&bytes, &limits).unwrap();

    let drawing_xml =
        std::str::from_utf8(zip_entries.get("xl/drawings/drawing1.xml").unwrap()).unwrap();
    assert!(
        drawing_xml.contains("<xdr:graphicFrame"),
        "chart-only drawing should have graphicFrame"
    );
    assert!(
        !drawing_xml.contains("<xdr:pic>"),
        "chart-only drawing should NOT have pic"
    );

    let anchor_count = drawing_xml.matches("<xdr:twoCellAnchor").count();
    assert_eq!(anchor_count, 1, "chart-only should have 1 anchor");
}

/// A chart-free sheet whose only drawing lives in preserved parts (the shape
/// `addImage` produces) must still reference it with `<drawing r:id>`.
fn image_only_workbook(with_drawing_part: bool) -> WorkbookData {
    let mut wb = minimal_workbook("Sheet1", Vec::new());
    wb.preserved_entries.insert(
        "xl/worksheets/_rels/sheet1.xml.rels".to_string(),
        br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/></Relationships>"#
            .to_vec(),
    );
    if with_drawing_part {
        wb.preserved_entries.insert(
            "xl/drawings/drawing1.xml".to_string(),
            br#"<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"/>"#
                .to_vec(),
        );
    }
    wb
}

#[test]
fn test_image_only_sheet_emits_drawing_element() {
    let wb = image_only_workbook(true);
    let bytes = write_xlsx(&wb).expect("write_xlsx should succeed");
    let zip_entries = read_zip_entries(&bytes, &ZipSecurityLimits::default()).unwrap();
    let sheet_xml =
        std::str::from_utf8(zip_entries.get("xl/worksheets/sheet1.xml").unwrap()).unwrap();
    assert_eq!(
        sheet_xml.matches("<drawing ").count(),
        1,
        "exactly one <drawing>:
{sheet_xml}"
    );
    assert!(
        sheet_xml.contains(r#"<drawing r:id="rId1"/>"#),
        "sheet should reference the preserved drawing:
{sheet_xml}"
    );
    assert!(zip_entries.contains_key("xl/drawings/drawing1.xml"));
    let rels = std::str::from_utf8(
        zip_entries
            .get("xl/worksheets/_rels/sheet1.xml.rels")
            .unwrap(),
    )
    .unwrap();
    assert!(rels.contains("drawing1.xml"));
}

#[test]
fn test_image_only_sheet_with_missing_drawing_part_emits_nothing() {
    let wb = image_only_workbook(false);
    let bytes = write_xlsx(&wb).expect("write_xlsx should succeed");
    let zip_entries = read_zip_entries(&bytes, &ZipSecurityLimits::default()).unwrap();
    let sheet_xml =
        std::str::from_utf8(zip_entries.get("xl/worksheets/sheet1.xml").unwrap()).unwrap();
    assert!(
        !sheet_xml.contains("<drawing"),
        "a dangling drawing rId makes Excel repair the file:
{sheet_xml}"
    );
}

// -----------------------------------------------------------------------
// Cell metadata: cm on <c> and the xl/metadata.xml part
// -----------------------------------------------------------------------

/// A sheet whose A1 holds the dynamic array `=SEQUENCE(3)` spilling to
/// A1:A3, marked by `cell_metadata` and/or the deprecated `dynamic_array`.
fn sequence_workbook(cell_metadata: Option<u32>, dynamic_array: Option<bool>) -> WorkbookData {
    minimal_workbook(
        "Sheet1",
        vec![Row {
            index: 1,
            cells: vec![Cell {
                reference: "A1".into(),
                value: Some("1".into()),
                formula: Some("SEQUENCE(3)".into()),
                formula_type: Some("array".into()),
                formula_ref: Some("A1:A3".into()),
                cell_metadata,
                dynamic_array,
                ..Default::default()
            }],
            ..Default::default()
        }],
    )
}

/// The `sheetMetadata` relationships of the workbook part in `entries`.
fn metadata_rels(entries: &HashMap<String, Vec<u8>>) -> Vec<String> {
    let rels = Relationships::parse(&entries["xl/_rels/workbook.xml.rels"]).unwrap();
    rels.iter()
        .filter(|r| r.rel_type == REL_SHEET_METADATA)
        .map(|r| r.target.to_string())
        .collect()
}

#[test]
fn dynamic_array_cell_gets_cm_on_c_and_a_wired_metadata_part() {
    for (cm, dynamic) in [(Some(1), Some(true)), (None, Some(true)), (Some(1), None)] {
        let bytes = write_xlsx(&sequence_workbook(cm, dynamic)).unwrap();
        let entries = read_zip_entries(&bytes, &ZipSecurityLimits::default()).unwrap();

        let sheet = std::str::from_utf8(&entries["xl/worksheets/sheet1.xml"]).unwrap();
        assert!(
            sheet.contains(r#"<c r="A1" cm="1"><f t="array" ref="A1:A3">"#),
            "{sheet}"
        );
        assert_eq!(sheet.matches("cm=").count(), 1, "cm only on <c>: {sheet}");

        assert_eq!(
            entries.get(METADATA_PATH).map(Vec::as_slice),
            Some(DYNAMIC_ARRAY_METADATA_XML.as_bytes())
        );
        assert_eq!(metadata_rels(&entries), ["metadata.xml"]);
        let ct = ContentTypes::parse(&entries["[Content_Types].xml"]).unwrap();
        assert_eq!(
            ct.overrides.get("/xl/metadata.xml").map(String::as_str),
            Some(CT_SHEET_METADATA)
        );

        // Read back: cm survives, and a second write keeps one part and one rel.
        let wb2 = crate::reader::read_xlsx(&bytes).unwrap();
        let a1 = &wb2.sheets[0].worksheet.rows[0].cells[0];
        assert_eq!((a1.cell_metadata, a1.dynamic_array), (Some(1), Some(true)));
        let again = write_xlsx(&wb2).unwrap();
        let entries2 = read_zip_entries(&again, &ZipSecurityLimits::default()).unwrap();
        assert_eq!(
            entries2.get(METADATA_PATH).map(Vec::as_slice),
            Some(DYNAMIC_ARRAY_METADATA_XML.as_bytes())
        );
        assert_eq!(metadata_rels(&entries2), ["metadata.xml"]);
    }
}

#[test]
fn workbook_without_cell_metadata_has_no_metadata_part() {
    let bytes = write_xlsx(&sequence_workbook(None, None)).unwrap();
    let entries = read_zip_entries(&bytes, &ZipSecurityLimits::default()).unwrap();
    assert!(!entries.contains_key(METADATA_PATH));
    assert!(metadata_rels(&entries).is_empty());
    let sheet = std::str::from_utf8(&entries["xl/worksheets/sheet1.xml"]).unwrap();
    assert!(!sheet.contains("cm="), "{sheet}");
}

#[test]
fn preserved_metadata_part_keeps_its_bytes_and_gains_rel_and_override() {
    let own = b"<metadata xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"/>";
    // The preserved part wins even for a cm index the generated part lacks.
    let mut wb = sequence_workbook(Some(2), Some(true));
    wb.preserved_entries
        .insert(METADATA_PATH.into(), own.to_vec());
    let bytes = write_xlsx(&wb).unwrap();
    let entries = read_zip_entries(&bytes, &ZipSecurityLimits::default()).unwrap();
    assert_eq!(entries[METADATA_PATH], own);
    assert_eq!(metadata_rels(&entries), ["metadata.xml"]);
    let ct = ContentTypes::parse(&entries["[Content_Types].xml"]).unwrap();
    assert_eq!(
        ct.overrides.get("/xl/metadata.xml").map(String::as_str),
        Some(CT_SHEET_METADATA)
    );
    let sheet = std::str::from_utf8(&entries["xl/worksheets/sheet1.xml"]).unwrap();
    assert!(sheet.contains(r#"<c r="A1" cm="2">"#), "{sheet}");

    // Without cells carrying cm the preserved part is still wired.
    let mut plain = sequence_workbook(None, None);
    plain
        .preserved_entries
        .insert(METADATA_PATH.into(), own.to_vec());
    let entries =
        read_zip_entries(&write_xlsx(&plain).unwrap(), &ZipSecurityLimits::default()).unwrap();
    assert_eq!(metadata_rels(&entries), ["metadata.xml"]);
}

#[test]
fn cell_metadata_index_without_a_metadata_part_is_rejected() {
    let err = write_xlsx(&sequence_workbook(Some(2), None)).unwrap_err();
    assert_eq!(err.code(), "INVALID_CELL_VALUE", "{err}");
}

// -----------------------------------------------------------------------
// calcChain entries: only those on a formula cell are written
// -----------------------------------------------------------------------

/// Sheet "S": A1 holds a formula (the master of shared group 0, over
/// A1:C1), B1 a plain value, C1 a shared-formula follower of group 0 with a
/// value; `chain` lists `(cell, sheet_id)` pairs.
fn calc_chain_workbook(chain: &[(&str, u32)]) -> WorkbookData {
    let cell = |reference: &str| Cell {
        reference: reference.into(),
        cell_type: CellType::Number,
        value: Some("1".into()),
        ..Default::default()
    };
    let mut wb = minimal_workbook(
        "S",
        vec![Row {
            index: 1,
            cells: vec![
                Cell {
                    formula: Some("1+0".into()),
                    formula_type: Some("shared".into()),
                    formula_ref: Some("A1:C1".into()),
                    shared_index: Some(0),
                    ..cell("A1")
                },
                cell("B1"),
                Cell {
                    formula_type: Some("shared".into()),
                    shared_index: Some(0),
                    ..cell("C1")
                },
            ],
            ..Default::default()
        }],
    );
    wb.calc_chain = chain
        .iter()
        .map(|&(cell_ref, sheet_id)| CalcChainEntry {
            cell_ref: cell_ref.into(),
            sheet_id,
            ..Default::default()
        })
        .collect();
    wb
}

fn has_calc_chain_wiring(entries: &HashMap<String, Vec<u8>>) -> (bool, bool) {
    (
        part_text(entries, "xl/_rels/workbook.xml.rels").contains(REL_CALC_CHAIN),
        part_text(entries, "[Content_Types].xml").contains(CT_CALC_CHAIN),
    )
}

#[test]
fn calc_chain_keeps_only_formula_cells() {
    let wb = calc_chain_workbook(&[("A1", 1), ("B1", 1)]);
    let entries = zip_map(&write_xlsx(&wb).unwrap());
    let chain = calc_chain::parse(&entries["xl/calcChain.xml"]).unwrap();
    let refs: Vec<&str> = chain.iter().map(|e| e.cell_ref.as_str()).collect();
    assert_eq!(refs, ["A1"]);
    assert_eq!(has_calc_chain_wiring(&entries), (true, true));
}

#[test]
fn calc_chain_keeps_shared_followers_and_drops_unknown_sheets() {
    let wb = calc_chain_workbook(&[("C1", 1), ("A1", 0), ("A1", 2), ("D1", 1)]);
    let entries = zip_map(&write_xlsx(&wb).unwrap());
    let chain = calc_chain::parse(&entries["xl/calcChain.xml"]).unwrap();
    let got: Vec<(&str, u32)> = chain
        .iter()
        .map(|e| (e.cell_ref.as_str(), e.sheet_id))
        .collect();
    assert_eq!(got, [("C1", 1)]);
}

/// The calcChain keeps exactly the cells the worksheet writer gives an
/// `<f>`: a follower without a value whose master is in the sheet is kept;
/// a follower of a group without a master, and a cleared master (shared keys,
/// empty text) with its followers, are dropped.
#[test]
fn calc_chain_follows_the_shared_formula_masters_of_the_sheet() {
    let follower = |reference: &str, si: u32, value: Option<&str>| Cell {
        reference: reference.into(),
        cell_type: CellType::Number,
        value: value.map(Into::into),
        formula_type: Some("shared".into()),
        shared_index: Some(si),
        ..Default::default()
    };
    let mut wb = calc_chain_workbook(&[]);
    wb.sheets[0].worksheet.rows.push(Row {
        index: 2,
        cells: vec![
            follower("C2", 0, None),
            follower("D2", 7, Some("2")),
            follower("E2", 7, None),
            Cell {
                formula: Some(String::new()),
                formula_ref: Some("F2:F3".into()),
                ..follower("F2", 1, Some("3"))
            },
            follower("F3", 1, Some("4")),
        ],
        ..Default::default()
    });
    wb.calc_chain = ["A1", "C1", "C2", "D2", "E2", "F2", "F3"]
        .into_iter()
        .map(|cell_ref| CalcChainEntry {
            cell_ref: cell_ref.into(),
            sheet_id: 1,
            ..Default::default()
        })
        .collect();
    let entries = zip_map(&write_xlsx(&wb).unwrap());
    let chain = calc_chain::parse(&entries["xl/calcChain.xml"]).unwrap();
    let refs: Vec<&str> = chain.iter().map(|e| e.cell_ref.as_str()).collect();
    assert_eq!(refs, ["A1", "C1", "C2"]);
    let sheet = part_text(&entries, "xl/worksheets/sheet1.xml");
    for want in [
        r#"<c r="C2"><f t="shared" si="0"/></c>"#,
        r#"<c r="D2"><v>2</v></c>"#,
        r#"<c r="E2"/>"#,
        r#"<c r="F2"><v>3</v></c>"#,
        r#"<c r="F3"><v>4</v></c>"#,
    ] {
        assert!(sheet.contains(want), "missing {want} in {sheet}");
    }
}

/// A master needs text and a ref: the calcChain keeps a shared cell with
/// text whose group has no such master (written as a plain formula) and a
/// child with text whose master is in the sheet; it drops a textless child
/// of a group without a master and a cell whose only text is blank.
#[test]
fn calc_chain_follows_the_master_definition_text_and_ref() {
    let shared =
        |reference: &str, si: u32, formula: Option<&str>, formula_ref: Option<&str>| Cell {
            reference: reference.into(),
            cell_type: CellType::Number,
            value: Some("2".into()),
            formula: formula.map(Into::into),
            formula_type: Some("shared".into()),
            formula_ref: formula_ref.map(Into::into),
            shared_index: Some(si),
            ..Default::default()
        };
    let mut wb = calc_chain_workbook(&[]);
    wb.sheets[0].worksheet.rows.push(Row {
        index: 2,
        cells: vec![
            shared("G2", 5, Some("1+1"), None),
            shared("H2", 5, None, None),
            shared("I2", 6, Some(" "), Some("I2:I3")),
            shared("J2", 0, Some("A2+1"), None),
        ],
        ..Default::default()
    });
    wb.calc_chain = ["A1", "C1", "G2", "H2", "I2", "J2"]
        .into_iter()
        .map(|cell_ref| CalcChainEntry {
            cell_ref: cell_ref.into(),
            sheet_id: 1,
            ..Default::default()
        })
        .collect();
    let entries = zip_map(&write_xlsx(&wb).unwrap());
    let chain = calc_chain::parse(&entries["xl/calcChain.xml"]).unwrap();
    let refs: Vec<&str> = chain.iter().map(|e| e.cell_ref.as_str()).collect();
    assert_eq!(refs, ["A1", "C1", "G2", "J2"]);
    let sheet = part_text(&entries, "xl/worksheets/sheet1.xml");
    for want in [
        r#"<c r="G2"><f>1+1</f><v>2</v></c>"#,
        r#"<c r="H2"><v>2</v></c>"#,
        r#"<c r="I2"><v>2</v></c>"#,
        r#"<c r="J2"><f t="shared" si="0">A2+1</f><v>2</v></c>"#,
    ] {
        assert!(sheet.contains(want), "missing {want} in {sheet}");
    }
}

#[test]
fn calc_chain_without_formula_cells_writes_no_part() {
    let wb = calc_chain_workbook(&[("B1", 1)]);
    let entries = zip_map(&write_xlsx(&wb).unwrap());
    assert!(!entries.contains_key("xl/calcChain.xml"));
    assert_eq!(has_calc_chain_wiring(&entries), (false, false));
}

// -----------------------------------------------------------------------
// External hyperlinks: r:id + TargetMode="External" relationships
// -----------------------------------------------------------------------

const SHEET1_RELS: &str = "xl/worksheets/_rels/sheet1.xml.rels";

/// An Excel-shaped package: B2 links through `r:id` alone to an external
/// URL with a query string, A1 carries a comment, and the sheet rels also
/// hold an external OLE link the model does not read.
fn excel_shaped_link_package() -> Vec<u8> {
    let part = |name: &str, xml: &str| ZipEntry {
        name: name.into(),
        data: xml.as_bytes().to_vec(),
    };
    write_zip(&[
        part(
            "[Content_Types].xml",
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/comments1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml"/></Types>"#,
        ),
        part(
            "_rels/.rels",
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>"#,
        ),
        part(
            "xl/workbook.xml",
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>"#,
        ),
        part(
            "xl/_rels/workbook.xml.rels",
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>"#,
        ),
        part(
            "xl/worksheets/sheet1.xml",
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetData><row r="2"><c r="B2" t="inlineStr"><is><t>site</t></is></c></row></sheetData><hyperlinks><hyperlink ref="B2" r:id="rId1" display="site"/></hyperlinks></worksheet>"#,
        ),
        part(
            SHEET1_RELS,
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://e.com/?a=1&amp;b=2" TargetMode="External"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="../comments1.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject" Target="file:///C:/data/x.bin" TargetMode="External"/></Relationships>"#,
        ),
        part(
            "xl/comments1.xml",
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><authors><author>A</author></authors><commentList><comment ref="A1" authorId="0"><text><r><t>note</t></r></text></comment></commentList></comments>"#,
        ),
    ])
    .unwrap()
}

/// The parsed sheet1 rels of `bytes`, and its raw text.
fn sheet1_rels(bytes: &[u8]) -> (Relationships, String) {
    let entries = read_zip_entries(bytes, &ZipSecurityLimits::default()).unwrap();
    let raw = entries[SHEET1_RELS].clone();
    (
        Relationships::parse(&raw).unwrap(),
        String::from_utf8(raw).unwrap(),
    )
}

/// Every internal sheet1 rel targets a part present in `bytes`.
fn assert_no_dangling_sheet1_rels(bytes: &[u8]) {
    let entries = read_zip_entries(bytes, &ZipSecurityLimits::default()).unwrap();
    let (rels, _) = sheet1_rels(bytes);
    for rel in rels
        .iter()
        .filter(|r| r.target_mode == crate::ooxml::relationships::TargetMode::Internal)
    {
        let path = resolve_target(SHEET_PART, &rel.target);
        assert!(entries.contains_key(&path), "{} -> {path} missing", rel.id);
    }
}

#[test]
fn excel_external_hyperlink_survives_read_write_read() {
    let wb = crate::reader::read_xlsx(&excel_shaped_link_package()).unwrap();
    let link = &wb.sheets[0].worksheet.hyperlinks[0];
    assert_eq!(link.target.as_deref(), Some("https://e.com/?a=1&b=2"));
    assert_eq!(link.location, None);
    assert_eq!(wb.sheets[0].worksheet.comments.len(), 1);

    let bytes = write_xlsx(&wb).unwrap();
    let (rels, raw) = sheet1_rels(&bytes);
    let links: Vec<_> = rels.find_by_type(REL_HYPERLINK).collect();
    assert_eq!(links.len(), 1, "{raw}");
    assert_eq!(links[0].target, "https://e.com/?a=1&b=2");
    assert_eq!(
        links[0].target_mode,
        crate::ooxml::relationships::TargetMode::External
    );
    assert_eq!(raw.matches("&amp;").count(), 1, "{raw}");
    assert!(!raw.contains("&amp;amp;"), "{raw}");
    // The OLE link the model never read keeps TargetMode="External".
    let ole = rels.get_by_id("rId3").unwrap();
    assert_eq!(ole.target, "file:///C:/data/x.bin");
    assert_eq!(
        ole.target_mode,
        crate::ooxml::relationships::TargetMode::External
    );
    assert_eq!(raw.matches(r#"TargetMode="External""#).count(), 2, "{raw}");
    assert_no_dangling_sheet1_rels(&bytes);

    let entries = read_zip_entries(&bytes, &ZipSecurityLimits::default()).unwrap();
    let sheet = std::str::from_utf8(&entries["xl/worksheets/sheet1.xml"]).unwrap();
    let tag = format!(r#"<hyperlink ref="B2" r:id="{}""#, links[0].id);
    assert!(sheet.contains(&tag), "{sheet}");
    assert!(!sheet.contains("location="), "{sheet}");

    let wb2 = crate::reader::read_xlsx(&bytes).unwrap();
    let link2 = &wb2.sheets[0].worksheet.hyperlinks[0];
    assert_eq!(link2.target.as_deref(), Some("https://e.com/?a=1&b=2"));
    assert_eq!(link2.location, None);
    assert_eq!(wb2.sheets[0].worksheet.comments.len(), 1);

    // A second cycle neither duplicates nor orphans hyperlink rels.
    let bytes2 = write_xlsx(&wb2).unwrap();
    let (rels2, raw2) = sheet1_rels(&bytes2);
    assert_eq!(rels2.find_by_type(REL_HYPERLINK).count(), 1, "{raw2}");
    assert_eq!(rels2.len(), rels.len(), "{raw2}");
}

#[test]
fn malformed_preserved_sheet_rels_do_not_fail_the_write() {
    // A raw `&` in Target: the reader keeps the malformed part verbatim as a
    // preserved entry, and the writer must not re-parse it into a fatal
    // error.
    const RAW_AMP_RELS: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="http://a/?x=1&y=2" TargetMode="External"/></Relationships>"#;
    let mut parts =
        read_zip_entries(&excel_shaped_link_package(), &ZipSecurityLimits::default()).unwrap();
    parts.insert(SHEET1_RELS.into(), RAW_AMP_RELS.as_bytes().to_vec());
    let package = write_zip(
        &parts
            .into_iter()
            .map(|(name, data)| ZipEntry { name, data })
            .collect::<Vec<_>>(),
    )
    .unwrap();

    let wb = crate::reader::read_xlsx(&package).unwrap();
    let bytes = write_xlsx(&wb).unwrap();
    // No generated rels on this sheet: the verbatim copy is written as read.
    let entries = read_zip_entries(&bytes, &ZipSecurityLimits::default()).unwrap();
    assert_eq!(entries[SHEET1_RELS], RAW_AMP_RELS.as_bytes());
}

#[test]
fn json_read_path_resolves_external_hyperlinks() {
    let json = crate::reader::read_xlsx_json(&excel_shaped_link_package()).unwrap();
    let wb: WorkbookData = serde_json::from_str(&json).unwrap();
    let link = &wb.sheets[0].worksheet.hyperlinks[0];
    assert_eq!(link.target.as_deref(), Some("https://e.com/?a=1&b=2"));
    assert_eq!(link.location, None);
}

#[test]
fn removed_external_hyperlinks_leave_no_rel_behind() {
    let mut wb = crate::reader::read_xlsx(&excel_shaped_link_package()).unwrap();
    wb.sheets[0].worksheet.hyperlinks.clear();
    let bytes = write_xlsx(&wb).unwrap();
    let (rels, raw) = sheet1_rels(&bytes);
    assert_eq!(rels.find_by_type(REL_HYPERLINK).count(), 0, "{raw}");
    assert!(rels.get_by_id("rId3").is_some(), "{raw}");
    assert_no_dangling_sheet1_rels(&bytes);

    // A sheet whose only relationship was a hyperlink loses its rels part.
    let mut only_link = wb.clone();
    only_link.sheets[0].worksheet.comments.clear();
    only_link.preserved_entries.insert(
        SHEET1_RELS.into(),
        br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://e.com/" TargetMode="External"/></Relationships>"#.to_vec(),
    );
    let entries = read_zip_entries(
        &write_xlsx(&only_link).unwrap(),
        &ZipSecurityLimits::default(),
    )
    .unwrap();
    assert!(!entries.contains_key(SHEET1_RELS));
}

#[test]
fn a_1_2_x_url_location_is_written_as_an_external_relationship() {
    let ws = WorksheetXml {
        hyperlinks: vec![
            crate::ooxml::worksheet::Hyperlink {
                cell_ref: "A1".into(),
                location: Some("https://x.y/".into()),
                ..Default::default()
            },
            crate::ooxml::worksheet::Hyperlink {
                cell_ref: "A2".into(),
                location: Some("Data!A1".into()),
                ..Default::default()
            },
        ],
        ..Default::default()
    };
    let wb = WorkbookData {
        sheets: vec![SheetData {
            name: "Data".into(),
            worksheet: ws,
            ..Default::default()
        }],
        ..Default::default()
    };
    let bytes = write_xlsx(&wb).unwrap();
    let (rels, raw) = sheet1_rels(&bytes);
    assert_eq!(rels.len(), 1, "{raw}");
    assert!(
        raw.contains(r#"Target="https://x.y/" TargetMode="External""#),
        "{raw}"
    );

    let back = crate::reader::read_xlsx(&bytes).unwrap();
    let links = &back.sheets[0].worksheet.hyperlinks;
    assert_eq!(links[0].target.as_deref(), Some("https://x.y/"));
    assert_eq!(links[0].location, None);
    assert_eq!(links[1].target, None);
    assert_eq!(links[1].location.as_deref(), Some("Data!A1"));
}

// -----------------------------------------------------------------------
// Drawing and media integrity: part allocation, one-pass rId remap,
// anchor order, content types
// -----------------------------------------------------------------------

const DRAWING_NS: &str = r#"xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships""#;

/// A bar chart over `Sheet1!$A$1:$A$5` anchored at column `col`.
fn bar_chart(col: u32) -> crate::ooxml::charts::WorksheetChart {
    use crate::ooxml::charts::{ChartAnchor, ChartData, ChartSeries, WorksheetChart};
    WorksheetChart {
        chart: ChartData {
            series: vec![ChartSeries {
                val_ref: "Sheet1!$A$1:$A$5".into(),
                ..Default::default()
            }],
            ..Default::default()
        },
        anchor: ChartAnchor {
            from_col: col,
            to_col: col + 5,
            to_row: 10,
            ..Default::default()
        },
    }
}

/// One `<xdr:{tag}>` picture anchor named `name` embedding `rid`.
fn pic_anchor(tag: &str, name: &str, rid: &str) -> String {
    format!(
        r#"<xdr:{tag}><xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:ext cx="1" cy="1"/><xdr:pic><xdr:nvPicPr><xdr:cNvPr id="9" name="{name}"/><xdr:cNvPicPr/></xdr:nvPicPr><xdr:blipFill><a:blip r:embed="{rid}"/></xdr:blipFill><xdr:spPr/></xdr:pic><xdr:clientData/></xdr:{tag}>"#
    )
}

fn drawing_part(children: &str) -> Vec<u8> {
    format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><xdr:wsDr {DRAWING_NS}>{children}</xdr:wsDr>"#
    )
    .into_bytes()
}

/// A drawing .rels holding `rId1..` image relationships, in order.
fn image_rels(targets: &[&str]) -> Vec<u8> {
    use std::fmt::Write as _;
    let mut rels = String::new();
    for (i, target) in targets.iter().enumerate() {
        write!(
            rels,
            r#"<Relationship Id="rId{}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="{target}"/>"#,
            i + 1
        )
        .unwrap();
    }
    let xml = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">{rels}</Relationships>"#
    );
    xml.into_bytes()
}

/// A worksheet .rels whose only relationship is a drawing at `target`.
fn sheet_drawing_rels(target: &str) -> Vec<u8> {
    format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdDrawing1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="{target}"/></Relationships>"#
    )
    .into_bytes()
}

/// Adds a preserved drawing (`drawing1.xml`, its .rels and media) to
/// sheet `sheet_num`, the shape `addImage` leaves in `preserved_entries`.
fn preserve_drawing(wb: &mut WorkbookData, sheet_num: usize, children: &str, media: &[&str]) {
    let targets: Vec<String> = media.iter().map(|m| format!("../media/{m}")).collect();
    let targets: Vec<&str> = targets.iter().map(String::as_str).collect();
    wb.preserved_entries
        .insert("xl/drawings/drawing1.xml".into(), drawing_part(children));
    wb.preserved_entries.insert(
        "xl/drawings/_rels/drawing1.xml.rels".into(),
        image_rels(&targets),
    );
    wb.preserved_entries.insert(
        format!("xl/worksheets/_rels/sheet{sheet_num}.xml.rels"),
        sheet_drawing_rels("../drawings/drawing1.xml"),
    );
    for (i, m) in media.iter().enumerate() {
        wb.preserved_entries
            .insert(format!("xl/media/{m}"), vec![u8::try_from(i).unwrap()]);
    }
}

fn zip_map(bytes: &[u8]) -> HashMap<String, Vec<u8>> {
    read_zip_entries(bytes, &ZipSecurityLimits::default())
        .unwrap()
        .into_iter()
        .collect()
}

fn part_text(entries: &HashMap<String, Vec<u8>>, path: &str) -> String {
    String::from_utf8(
        entries
            .get(path)
            .unwrap_or_else(|| panic!("missing {path}"))
            .clone(),
    )
    .unwrap()
}

/// The values of every `r:embed` attribute, in document order.
fn embed_ids(drawing_xml: &str) -> Vec<&str> {
    drawing_xml
        .split("r:embed=\"")
        .skip(1)
        .map(|rest| &rest[..rest.find('"').unwrap()])
        .collect()
}

/// Asserts every internal target of the .rels of `source` is a ZIP entry,
/// and returns the parsed relationships.
fn assert_rel_targets_exist(entries: &HashMap<String, Vec<u8>>, source: &str) -> Relationships {
    let rels = Relationships::parse(entries.get(&rels_path(source)).unwrap()).unwrap();
    for rel in &rels {
        let path = resolve_target(source, &rel.target);
        assert!(
            entries.contains_key(&path),
            "{source}: {} targets missing {path}",
            rel.id
        );
    }
    rels
}

/// The drawing part sheet `sheet_num` references.
fn sheet_drawing_path(entries: &HashMap<String, Vec<u8>>, sheet_num: usize) -> String {
    let sheet = format!("xl/worksheets/sheet{sheet_num}.xml");
    let rels = assert_rel_targets_exist(entries, &sheet);
    let rel = rels
        .find_by_type(REL_DRAWING)
        .next()
        .expect("sheet has a drawing rel");
    resolve_target(&sheet, &rel.target)
}

#[test]
fn chart_plus_two_images_keep_distinct_embed_ids() {
    let mut wb = make_workbook_with_charts(vec![bar_chart(4)]);
    let children =
        pic_anchor("twoCellAnchor", "Img1", "rId1") + &pic_anchor("twoCellAnchor", "Img2", "rId2");
    preserve_drawing(&mut wb, 1, &children, &["image1.png", "image2.png"]);

    let entries = zip_map(&write_xlsx(&wb).unwrap());
    let drawing = part_text(&entries, "xl/drawings/drawing1.xml");
    assert_eq!(embed_ids(&drawing), ["rId2", "rId3"], "{drawing}");

    let rels = assert_rel_targets_exist(&entries, "xl/drawings/drawing1.xml");
    assert_eq!(rels.get_by_id("rId1").unwrap().rel_type, REL_CHART);
    assert_eq!(
        rels.get_by_id("rId2").unwrap().target,
        "../media/image1.png"
    );
    assert_eq!(
        rels.get_by_id("rId3").unwrap().target,
        "../media/image2.png"
    );
    assert_rel_targets_exist(&entries, "xl/worksheets/sheet1.xml");
}

/// A 3-D column chart part, which the chart model keeps verbatim.
const BAR_3D_CHART: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><c:chart><c:autoTitleDeleted val="1"/><c:view3D><c:rotX val="15"/><c:rotY val="20"/><c:rAngAx val="1"/></c:view3D><c:plotArea><c:layout/><c:bar3DChart><c:barDir val="col"/><c:grouping val="clustered"/><c:varyColors val="0"/><c:ser><c:idx val="0"/><c:order val="0"/><c:val><c:numRef><c:f>Sheet1!$B$2:$B$5</c:f></c:numRef></c:val></c:ser><c:gapWidth val="150"/><c:shape val="box"/><c:axId val="1"/><c:axId val="2"/></c:bar3DChart><c:catAx><c:axId val="1"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:crossAx val="2"/></c:catAx><c:valAx><c:axId val="2"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/><c:crossAx val="1"/></c:valAx></c:plotArea><c:plotVisOnly val="1"/></c:chart></c:chartSpace>"#;

#[test]
fn raw_chart_and_its_rels_are_written_verbatim_and_read_back() {
    let parsed = crate::ooxml::charts::ChartData::parse(BAR_3D_CHART.as_bytes()).unwrap();
    assert_eq!(parsed.raw_xml.as_deref(), Some(BAR_3D_CHART));
    let rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/package" Target="../embeddings/Book1.xlsx"/></Relationships>"#;
    let mut chart = bar_chart(4);
    chart.chart = crate::ooxml::charts::ChartData {
        raw_rels: Some(rels.into()),
        ..parsed
    };
    let mut wb = make_workbook_with_charts(vec![chart]);
    // A preserved chart part makes the raw chart move to chart2.xml.
    wb.preserved_entries
        .insert("xl/charts/chart1.xml".into(), b"<c:chartSpace/>".to_vec());
    wb.preserved_entries
        .insert("xl/embeddings/Book1.xlsx".into(), vec![0x50, 0x4b]);

    let bytes = write_xlsx(&wb).unwrap();
    let entries = zip_map(&bytes);
    assert_eq!(entries["xl/charts/chart2.xml"], BAR_3D_CHART.as_bytes());
    assert_eq!(entries["xl/charts/_rels/chart2.xml.rels"], rels.as_bytes());
    assert_rel_targets_exist(&entries, "xl/charts/chart2.xml");

    // Read back: the chart and its .rels are model data again, not preserved
    // entries, and a second write is byte-identical for both.
    let wb2 = crate::reader::read_xlsx(&bytes).unwrap();
    let chart2 = &wb2.sheets[0].worksheet.charts[0].chart;
    assert_eq!(chart2.raw_xml.as_deref(), Some(BAR_3D_CHART));
    assert_eq!(chart2.raw_rels.as_deref(), Some(rels));
    assert!(
        !wb2.preserved_entries
            .contains_key("xl/charts/_rels/chart2.xml.rels")
    );
    let entries2 = zip_map(&write_xlsx(&wb2).unwrap());
    let drawing = sheet_drawing_path(&entries2, 1);
    let drawing_rels = assert_rel_targets_exist(&entries2, &drawing);
    let target = drawing_rels
        .find_by_type(REL_CHART)
        .next()
        .unwrap()
        .target
        .clone();
    let chart_path = resolve_target(&drawing, &target);
    assert_eq!(entries2[&chart_path], BAR_3D_CHART.as_bytes());
    assert_eq!(entries2[&rels_path(&chart_path)], rels.as_bytes());
}

/// A chart .rels with a chartStyle and a chartColorStyle relationship to
/// `style{n}.xml` and `colors{n}.xml`.
fn chart_style_rels(n: usize) -> String {
    format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId2" Type="http://schemas.microsoft.com/office/2011/relationships/chartColorStyle" Target="colors{n}.xml"/><Relationship Id="rId1" Type="http://schemas.microsoft.com/office/2011/relationships/chartStyle" Target="style{n}.xml"/></Relationships>"#
    )
}

#[test]
fn modelled_chart_rels_follow_the_chart_when_an_earlier_chart_is_removed() {
    let mut wb = make_workbook_with_charts(vec![bar_chart(0), bar_chart(8)]);
    for n in 1..=2 {
        wb.preserved_entries.insert(
            format!("xl/charts/_rels/chart{n}.xml.rels"),
            chart_style_rels(n).into_bytes(),
        );
        wb.preserved_entries.insert(
            format!("xl/charts/style{n}.xml"),
            b"<cs:chartStyle/>".to_vec(),
        );
        wb.preserved_entries.insert(
            format!("xl/charts/colors{n}.xml"),
            b"<cs:colorStyle/>".to_vec(),
        );
    }
    let bytes = write_xlsx(&wb).unwrap();
    let entries = zip_map(&bytes);
    assert_eq!(
        entries["xl/charts/_rels/chart2.xml.rels"],
        chart_style_rels(2).as_bytes()
    );

    let mut wb2 = crate::reader::read_xlsx(&bytes).unwrap();
    wb2.sheets[0].worksheet.charts.remove(0);
    let entries2 = zip_map(&write_xlsx(&wb2).unwrap());
    assert_eq!(
        part_text(&entries2, "xl/charts/_rels/chart1.xml.rels"),
        chart_style_rels(2)
    );
    assert!(!entries2.contains_key("xl/charts/_rels/chart2.xml.rels"));
    assert!(!entries2.contains_key("xl/charts/chart2.xml"));
    assert_rel_targets_exist(&entries2, "xl/charts/chart1.xml");
    // The style parts the written chart names are declared by Override.
    let ct = ContentTypes::parse(&entries2["[Content_Types].xml"]).unwrap();
    assert_eq!(ct.overrides["/xl/charts/style2.xml"], CT_CHART_STYLE);
    assert_eq!(ct.overrides["/xl/charts/colors2.xml"], CT_CHART_COLOR_STYLE);

    // Both charts were modelled; each carried its own .rels.
    let wb3 = crate::reader::read_xlsx(&bytes).unwrap();
    let charts = &wb3.sheets[0].worksheet.charts;
    assert!(charts.iter().all(|wsc| wsc.chart.raw_xml.is_none()));
    for (wsc, n) in charts.iter().zip(1..) {
        assert_eq!(
            wsc.chart.raw_rels.as_deref(),
            Some(chart_style_rels(n).as_str())
        );
    }
    assert!(
        !wb3.preserved_entries
            .keys()
            .any(|path| path.starts_with("xl/charts/_rels/"))
    );
}

const CHART_NS: &str = r#"xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships""#;

/// A line chart the model represents in full, with `tail` appended to its
/// `c:chartSpace` after `c:chart`.
fn line_chart_xml(tail: &str) -> String {
    format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace {CHART_NS}><c:chart><c:autoTitleDeleted val="1"/><c:plotArea><c:layout/><c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/><c:ser><c:idx val="0"/><c:order val="0"/><c:val><c:numRef><c:f>Sheet1!$B$2:$B$5</c:f></c:numRef></c:val></c:ser><c:marker val="1"/><c:axId val="1"/><c:axId val="2"/></c:lineChart><c:catAx><c:axId val="1"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:crossAx val="2"/></c:catAx><c:valAx><c:axId val="2"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/><c:crossAx val="1"/></c:valAx></c:plotArea><c:plotVisOnly val="1"/></c:chart>{tail}</c:chartSpace>"#
    )
}

/// A chart .rels holding the single relationship `rel_id` of `rel_type`.
fn single_rel(rel_id: &str, rel_type: &str, target: &str) -> String {
    format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="{rel_id}" Type="{rel_type}" Target="{target}"/></Relationships>"#
    )
}

/// A `c:userShapes` part holding one text box over the chart's top left.
const USER_SHAPES_PART: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:userShapes xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><cdr:relSizeAnchor xmlns:cdr="http://schemas.openxmlformats.org/drawingml/2006/chartDrawing"><cdr:from><cdr:x>0.1</cdr:x><cdr:y>0.1</cdr:y></cdr:from><cdr:to><cdr:x>0.4</cdr:x><cdr:y>0.25</cdr:y></cdr:to><cdr:sp macro="" textlink=""><cdr:nvSpPr><cdr:cNvPr id="2" name="TextBox 1"/><cdr:cNvSpPr txBox="1"/></cdr:nvSpPr><cdr:spPr><a:xfrm xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></a:xfrm><a:prstGeom xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" prst="rect"><a:avLst/></a:prstGeom></cdr:spPr><cdr:txBody><a:bodyPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" wrap="none" rtlCol="0"/><a:p xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:r><a:rPr lang="en-US" sz="1100"/><a:t>Note</a:t></a:r></a:p></cdr:txBody></cdr:sp></cdr:relSizeAnchor></c:userShapes>"#;

/// Writes a package whose sheet holds `chart_xml` with `chart_rels` and the
/// extra preserved parts, reads it, writes it again, and returns the second
/// read's workbook, the rewritten entries and the chart's path in them.
fn chart_read_write(
    chart_xml: &str,
    chart_rels: &str,
    parts: &[(&str, &[u8])],
) -> (WorkbookData, HashMap<String, Vec<u8>>, String) {
    let mut chart = bar_chart(0);
    chart.chart = crate::ooxml::charts::ChartData {
        raw_xml: Some(chart_xml.into()),
        raw_rels: Some(chart_rels.into()),
        ..Default::default()
    };
    let mut wb = make_workbook_with_charts(vec![chart]);
    for &(path, data) in parts {
        wb.preserved_entries.insert(path.into(), data.to_vec());
    }
    let bytes = write_xlsx(&wb).unwrap();
    assert_eq!(
        part_text(&zip_map(&bytes), "xl/charts/chart1.xml"),
        chart_xml
    );

    let wb2 = crate::reader::read_xlsx(&bytes).unwrap();
    let entries2 = zip_map(&write_xlsx(&wb2).unwrap());
    let drawing = sheet_drawing_path(&entries2, 1);
    let drawing_rels = assert_rel_targets_exist(&entries2, &drawing);
    let target = &drawing_rels.find_by_type(REL_CHART).next().unwrap().target;
    let chart_path = resolve_target(&drawing, target);
    (wb2, entries2, chart_path)
}

#[test]
fn chart_user_shapes_keep_their_reference_and_content_type() {
    let chart_xml = line_chart_xml(r#"<c:userShapes r:id="rId2"/>"#);
    let rels = single_rel("rId2", REL_CHART_USER_SHAPES, "../drawings/drawing2.xml");
    let (wb2, entries2, chart_path) = chart_read_write(
        &chart_xml,
        &rels,
        &[("xl/drawings/drawing2.xml", USER_SHAPES_PART.as_bytes())],
    );

    let chart = &wb2.sheets[0].worksheet.charts[0].chart;
    assert_eq!(chart.raw_xml.as_deref(), Some(chart_xml.as_str()));
    assert_eq!(chart.raw_rels.as_deref(), Some(rels.as_str()));
    assert!(part_text(&entries2, &chart_path).contains(r#"<c:userShapes r:id="rId2"/>"#));
    let chart_rels = assert_rel_targets_exist(&entries2, &chart_path);
    let shapes_rel = chart_rels.get_by_id("rId2").unwrap();
    assert_eq!(shapes_rel.rel_type, REL_CHART_USER_SHAPES);
    assert_eq!(
        resolve_target(&chart_path, &shapes_rel.target),
        "xl/drawings/drawing2.xml"
    );
    assert_eq!(
        entries2["xl/drawings/drawing2.xml"],
        USER_SHAPES_PART.as_bytes()
    );

    let drawing = sheet_drawing_path(&entries2, 1);
    assert_ne!(drawing, "xl/drawings/drawing2.xml");
    let ct = ContentTypes::parse(&entries2["[Content_Types].xml"]).unwrap();
    assert_eq!(ct.overrides["/xl/drawings/drawing2.xml"], CT_CHART_SHAPES);
    assert_eq!(ct.overrides[&format!("/{drawing}")], CT_DRAWING);
}

#[test]
fn preserved_chart_rels_give_their_user_shapes_the_chartshapes_type() {
    // A chart part the reader left preserved (no sheet drawing reaches it):
    // its .rels still types the userShapes part it names.
    let mut wb = make_workbook_with_charts(Vec::new());
    wb.preserved_entries.insert(
        "xl/charts/chart7.xml".into(),
        line_chart_xml(r#"<c:userShapes r:id="rId1"/>"#).into_bytes(),
    );
    wb.preserved_entries.insert(
        "xl/charts/_rels/chart7.xml.rels".into(),
        single_rel("rId1", REL_CHART_USER_SHAPES, "/xl/drawings/drawing5.xml").into_bytes(),
    );
    wb.preserved_entries.insert(
        "xl/drawings/drawing5.xml".into(),
        USER_SHAPES_PART.as_bytes().to_vec(),
    );
    let entries = zip_map(&write_xlsx(&wb).unwrap());
    let ct = ContentTypes::parse(&entries["[Content_Types].xml"]).unwrap();
    assert_eq!(ct.overrides["/xl/drawings/drawing5.xml"], CT_CHART_SHAPES);
    assert_eq!(ct.overrides["/xl/charts/chart7.xml"], CT_CHART);
}

#[test]
fn chart_external_data_keeps_the_chart_raw() {
    let embed_rel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/package";
    for tail in [
        r#"<c:externalData r:id="rId1"><c:autoUpdate val="0"/></c:externalData>"#,
        r#"<c:externalData r:id="rId1"/>"#,
        r#"<c:userShapes r:id="rId1"/>"#,
    ] {
        let chart_xml = line_chart_xml(tail);
        let parsed = crate::ooxml::charts::ChartData::parse(chart_xml.as_bytes()).unwrap();
        assert_eq!(
            parsed.raw_xml.as_deref(),
            Some(chart_xml.as_str()),
            "{tail}"
        );
        // The representable parts are still modelled.
        assert_eq!(
            parsed.chart_type,
            crate::ooxml::charts::ChartType::Line,
            "{tail}"
        );
    }

    let chart_xml =
        line_chart_xml(r#"<c:externalData r:id="rId1"><c:autoUpdate val="0"/></c:externalData>"#);
    let rels = single_rel("rId1", embed_rel, "../embeddings/Book1.xlsx");
    let (wb2, entries2, chart_path) = chart_read_write(
        &chart_xml,
        &rels,
        &[("xl/embeddings/Book1.xlsx", &[0x50, 0x4b])],
    );
    let chart = &wb2.sheets[0].worksheet.charts[0].chart;
    assert_eq!(chart.raw_xml.as_deref(), Some(chart_xml.as_str()));
    assert_eq!(entries2[&chart_path], chart_xml.as_bytes());
    assert_eq!(entries2[&rels_path(&chart_path)], rels.as_bytes());
    assert_rel_targets_exist(&entries2, &chart_path);
}

#[test]
fn unchanged_modelled_chart_without_rels_writes_identical_xml() {
    let bytes = write_xlsx(&make_workbook_with_charts(vec![bar_chart(0)])).unwrap();
    let entries = zip_map(&bytes);
    let wb2 = crate::reader::read_xlsx(&bytes).unwrap();
    let chart = &wb2.sheets[0].worksheet.charts[0].chart;
    assert!(chart.raw_xml.is_none() && chart.raw_rels.is_none());

    let entries2 = zip_map(&write_xlsx(&wb2).unwrap());
    assert_eq!(
        entries2["xl/charts/chart1.xml"],
        entries["xl/charts/chart1.xml"]
    );
    assert!(
        !entries2
            .keys()
            .any(|path| path.starts_with("xl/charts/_rels/"))
    );
    let ct = ContentTypes::parse(&entries2["[Content_Types].xml"]).unwrap();
    assert_eq!(ct.overrides["/xl/drawings/drawing1.xml"], CT_DRAWING);
}

#[test]
fn preserved_chart_part_name_is_not_reused_by_a_new_chart() {
    let mut wb = make_workbook_with_charts(vec![bar_chart(4)]);
    wb.preserved_entries
        .insert("xl/charts/chart1.xml".into(), b"<c:chartSpace/>".to_vec());

    let bytes = write_xlsx(&wb).expect("a preserved chart1.xml must not collide");
    let entries = zip_map(&bytes);
    assert_eq!(entries["xl/charts/chart1.xml"], b"<c:chartSpace/>");
    assert!(part_text(&entries, "xl/charts/chart2.xml").contains("<c:barChart>"));

    let drawing = sheet_drawing_path(&entries, 1);
    let rels = assert_rel_targets_exist(&entries, &drawing);
    let chart_rel = rels.find_by_type(REL_CHART).next().unwrap();
    assert_eq!(chart_rel.target, "../charts/chart2.xml");
    let ct = ContentTypes::parse(&entries["[Content_Types].xml"]).unwrap();
    assert_eq!(ct.overrides["/xl/charts/chart2.xml"], CT_CHART);
}

#[test]
fn sheet2_preserved_drawing_never_gains_sheet1_chart() {
    let mut wb = make_workbook_with_charts(vec![bar_chart(4)]);
    wb.sheets.push(SheetData {
        name: "Sheet2".into(),
        ..Default::default()
    });
    preserve_drawing(
        &mut wb,
        2,
        &pic_anchor("twoCellAnchor", "Sheet2Img", "rId1"),
        &["image1.png"],
    );
    let preserved_drawing = wb.preserved_entries["xl/drawings/drawing1.xml"].clone();

    let entries = zip_map(&write_xlsx(&wb).unwrap());
    assert_eq!(sheet_drawing_path(&entries, 2), "xl/drawings/drawing1.xml");
    assert_eq!(entries["xl/drawings/drawing1.xml"], preserved_drawing);

    let sheet1_drawing = sheet_drawing_path(&entries, 1);
    assert_ne!(sheet1_drawing, "xl/drawings/drawing1.xml");
    let chart_drawing = part_text(&entries, &sheet1_drawing);
    assert!(
        chart_drawing.contains("<xdr:graphicFrame"),
        "{chart_drawing}"
    );
    assert!(!chart_drawing.contains("Sheet2Img"), "{chart_drawing}");
    assert_rel_targets_exist(&entries, &sheet1_drawing);
    assert_rel_targets_exist(&entries, "xl/drawings/drawing1.xml");
}

#[test]
fn merged_drawing_keeps_preserved_anchor_order() {
    let mut wb = make_workbook_with_charts(vec![bar_chart(4)]);
    let children = pic_anchor("oneCellAnchor", "BackImg", "rId1")
        + &pic_anchor("twoCellAnchor", "FrontImg", "rId1");
    preserve_drawing(&mut wb, 1, &children, &["image1.png"]);

    let entries = zip_map(&write_xlsx(&wb).unwrap());
    let drawing = part_text(&entries, "xl/drawings/drawing1.xml");
    let back = drawing.find("BackImg").unwrap();
    let front = drawing.find("FrontImg").unwrap();
    assert!(back < front, "z-order reversed:\n{drawing}");
    assert_eq!(embed_ids(&drawing), ["rId2", "rId2"], "{drawing}");
}

/// The sheet-1 drawing of `bytes` holds the chart frame only: no preserved
/// anchor was merged into it.
fn assert_chart_only_drawing(bytes: &[u8]) {
    let entries = zip_map(bytes);
    let drawing = part_text(&entries, &sheet_drawing_path(&entries, 1));
    assert!(drawing.contains("<xdr:graphicFrame"), "{drawing}");
    assert!(!drawing.contains("<xdr:pic>"), "{drawing}");
}

#[test]
fn malformed_preserved_drawing_rels_is_not_fatal() {
    let mut wb = make_workbook_with_charts(vec![bar_chart(4)]);
    preserve_drawing(
        &mut wb,
        1,
        &pic_anchor("twoCellAnchor", "Img", "rId1"),
        &["image1.png"],
    );
    wb.preserved_entries.insert(
        "xl/drawings/_rels/drawing1.xml.rels".into(),
        b"<Relationships><a></b>".to_vec(),
    );

    let bytes = write_xlsx(&wb).expect("a malformed drawing .rels is a warning, not an error");
    assert_chart_only_drawing(&bytes);
    let entries = zip_map(&bytes);
    let drawing = sheet_drawing_path(&entries, 1);
    let rels = assert_rel_targets_exist(&entries, &drawing);
    assert_eq!(rels.len(), 1);
    assert_eq!(rels.find_by_type(REL_CHART).count(), 1);
}

#[test]
fn non_utf8_preserved_drawing_is_not_merged() {
    let mut wb = make_workbook_with_charts(vec![bar_chart(4)]);
    preserve_drawing(
        &mut wb,
        1,
        &pic_anchor("twoCellAnchor", "Img", "rId1"),
        &["image1.png"],
    );
    let utf8 = String::from_utf8(wb.preserved_entries["xl/drawings/drawing1.xml"].clone())
        .unwrap()
        .replace("UTF-8", "UTF-16");
    let mut utf16 = vec![0xFF, 0xFE];
    for unit in utf8.encode_utf16() {
        utf16.extend_from_slice(&unit.to_le_bytes());
    }
    wb.preserved_entries
        .insert("xl/drawings/drawing1.xml".into(), utf16);

    assert_chart_only_drawing(&write_xlsx(&wb).expect("a damaged preserved drawing is not fatal"));
}

#[test]
fn preserved_chart_frames_are_replaced_and_root_namespaces_copied() {
    let mut wb = make_workbook_with_charts(vec![bar_chart(4)]);
    let old_chart = concat!(
        r#"<xdr:twoCellAnchor><xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>"#,
        r#"<xdr:to><xdr:col>3</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>3</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>"#,
        r#"<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="3" name="OldChart"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>"#,
        r#"<xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">"#,
        r#"<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="rId2"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor>"#,
    );
    let alternate = r#"<mc:AlternateContent><mc:Choice Requires="a14"><xdr:twoCellAnchor><xdr:sp><xdr:nvSpPr><xdr:cNvPr id="4" name="Slicer"/></xdr:nvSpPr></xdr:sp><xdr:clientData/></xdr:twoCellAnchor></mc:Choice><mc:Fallback/></mc:AlternateContent>"#;
    let children = format!(
        "{}{old_chart}<!-- comment --><![CDATA[x]]>{alternate}",
        pic_anchor("twoCellAnchor", "Img", "rId1")
    );
    let drawing = String::from_utf8(drawing_part(&children))
        .unwrap()
        .replace(
            "<xdr:wsDr ",
            r#"<xdr:wsDr xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main" "#,
        );
    wb.preserved_entries
        .insert("xl/drawings/drawing1.xml".into(), drawing.into_bytes());
    wb.preserved_entries.insert(
        "xl/drawings/_rels/drawing1.xml.rels".into(),
        String::from_utf8(image_rels(&["../media/image1.png"]))
            .unwrap()
            .replace(
                "</Relationships>",
                r#"<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart9.xml"/></Relationships>"#,
            )
            .into_bytes(),
    );
    wb.preserved_entries.insert(
        "xl/worksheets/_rels/sheet1.xml.rels".into(),
        sheet_drawing_rels("../drawings/drawing1.xml"),
    );
    wb.preserved_entries
        .insert("xl/media/image1.png".into(), vec![1]);

    let entries = zip_map(&write_xlsx(&wb).unwrap());
    let out = part_text(&entries, "xl/drawings/drawing1.xml");
    assert_eq!(out.matches("<xdr:graphicFrame").count(), 1, "{out}");
    assert!(!out.contains("OldChart"), "{out}");
    assert_eq!(out.matches("<xdr:pic>").count(), 1, "{out}");
    assert_eq!(out.matches("<mc:AlternateContent>").count(), 1, "{out}");
    assert!(out.contains(alternate), "kept verbatim:\n{out}");
    let root = &out[out.find("<xdr:wsDr").unwrap()..];
    let root = &root[..root.find('>').unwrap()];
    for decl in [
        r#"xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006""#,
        r#"xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main""#,
    ] {
        assert_eq!(root.matches(decl).count(), 1, "{root}");
    }
    assert_eq!(root.matches("xmlns:xdr=").count(), 1, "{root}");

    let rels = assert_rel_targets_exist(&entries, "xl/drawings/drawing1.xml");
    assert_eq!(rels.find_by_type(REL_CHART).count(), 1);
    assert!(rels.iter().all(|rel| rel.target != "../charts/chart9.xml"));
}

#[test]
fn rel_id_remap_covers_every_r_attribute() {
    let remap: HashMap<&str, &str> = [("rId1", "rId2"), ("rId2", "rId3")].into();
    let mut out = String::new();
    push_remapped_rel_ids(
        &mut out,
        r#"<dgm:relIds r:dm = "rId1" r:lo='rId2' r:qs="rId7"/><r:x r:pict="rId1"/>"#,
        &remap,
    );
    assert_eq!(
        out,
        r#"<dgm:relIds r:dm = "rId2" r:lo='rId3' r:qs="rId7"/><r:x r:pict="rId2"/>"#
    );
}

/// SpreadsheetDrawing, the namespace of every `xdr:` element.
const SPREADSHEET_DRAWING_NS: &str =
    "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";

#[test]
fn openpyxl_default_namespace_drawing_stays_valid_after_merge() {
    use quick_xml::NsReader;
    use quick_xml::name::{Namespace, ResolveResult};

    let mut wb = make_workbook_with_charts(vec![bar_chart(4)]);
    preserve_drawing(&mut wb, 1, "", &["image1.png"]);
    // The shape openpyxl 3.1 writes: SpreadsheetDrawing is the default
    // namespace and the anchors are unprefixed.
    let openpyxl = format!(
        r#"<?xml version='1.0' encoding='UTF-8'?><wsDr xmlns="{SPREADSHEET_DRAWING_NS}"><oneCellAnchor><from><col>0</col><colOff>0</colOff><row>0</row><rowOff>0</rowOff></from><ext cx="1" cy="1"/><pic><nvPicPr><cNvPr id="1" name="OpenpyxlImg"/><cNvPicPr/></nvPicPr><blipFill><a:blip xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="rId1"/></blipFill><spPr/></pic><clientData/></oneCellAnchor></wsDr>"#
    );
    wb.preserved_entries
        .insert("xl/drawings/drawing1.xml".into(), openpyxl.into_bytes());

    let entries = zip_map(&write_xlsx(&wb).unwrap());
    let out = part_text(&entries, "xl/drawings/drawing1.xml");
    let root = &out[out.find("<xdr:wsDr").unwrap()..];
    let root = &root[..root.find('>').unwrap()];
    assert_eq!(
        root.matches(&format!(r#" xmlns="{SPREADSHEET_DRAWING_NS}""#))
            .count(),
        1,
        "{root}"
    );
    assert!(out.contains("OpenpyxlImg"), "{out}");
    assert_eq!(embed_ids(&out), ["rId2"], "{out}");

    let mut reader = NsReader::from_str(&out);
    let mut anchors = 0;
    loop {
        let (res, event) = reader.read_resolved_event().unwrap();
        match event {
            Event::Start(e) | Event::Empty(e) => {
                assert!(
                    !matches!(res, ResolveResult::Unknown(_)),
                    "unbound prefix {res:?}:\n{out}"
                );
                if e.local_name().as_ref() == "oneCellAnchor" {
                    anchors += 1;
                    assert_eq!(
                        res,
                        ResolveResult::Bound(Namespace(SPREADSHEET_DRAWING_NS)),
                        "{out}"
                    );
                }
            }
            Event::Eof => break,
            _ => {}
        }
    }
    assert_eq!(anchors, 1, "{out}");
}

#[test]
fn smartart_rel_ids_follow_their_renumbered_rels() {
    use std::fmt::Write as _;
    const REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    let mut wb = make_workbook_with_charts(vec![bar_chart(4)]);
    let smartart = concat!(
        r#"<xdr:twoCellAnchor><xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>"#,
        r#"<xdr:to><xdr:col>3</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>3</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>"#,
        r#"<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="2" name="Diagram 1"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>"#,
        r#"<xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/diagram">"#,
        r#"<dgm:relIds xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" r:dm="rId1" r:lo="rId2" r:qs="rId3" r:cs="rId4"/>"#,
        r#"</a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor>"#,
    );
    preserve_drawing(&mut wb, 1, smartart, &[]);
    let kinds = [
        ("dm", "diagramData", "data1"),
        ("lo", "diagramLayout", "layout1"),
        ("qs", "diagramQuickStyle", "quickStyle1"),
        ("cs", "diagramColors", "colors1"),
    ];
    let mut rels = String::new();
    for (i, (_, kind, file)) in kinds.iter().enumerate() {
        write!(
            rels,
            r#"<Relationship Id="rId{}" Type="{REL}/{kind}" Target="../diagrams/{file}.xml"/>"#,
            i + 1
        )
        .unwrap();
        wb.preserved_entries
            .insert(format!("xl/diagrams/{file}.xml"), b"<x/>".to_vec());
    }
    wb.preserved_entries.insert(
        "xl/drawings/_rels/drawing1.xml.rels".into(),
        format!(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">{rels}</Relationships>"#
        )
        .into_bytes(),
    );

    let entries = zip_map(&write_xlsx(&wb).unwrap());
    let out = part_text(&entries, "xl/drawings/drawing1.xml");
    let rels = assert_rel_targets_exist(&entries, "xl/drawings/drawing1.xml");
    assert_eq!(rels.get_by_id("rId1").unwrap().rel_type, REL_CHART);
    for (i, (attr, kind, file)) in kinds.iter().enumerate() {
        let id = format!("rId{}", i + 2);
        assert!(out.contains(&format!(r#"r:{attr}="{id}""#)), "{out}");
        let rel = rels.get_by_id(&id).unwrap();
        assert_eq!(rel.rel_type, format!("{REL}/{kind}"));
        assert_eq!(rel.target, format!("../diagrams/{file}.xml"));
    }
}

#[test]
fn truncated_preserved_drawing_is_not_merged() {
    let mut wb = make_workbook_with_charts(vec![bar_chart(4)]);
    preserve_drawing(
        &mut wb,
        1,
        &pic_anchor("twoCellAnchor", "Img", "rId1"),
        &["image1.png"],
    );
    let whole =
        String::from_utf8(wb.preserved_entries["xl/drawings/drawing1.xml"].clone()).unwrap();
    let truncated = whole
        .strip_suffix("</xdr:twoCellAnchor></xdr:wsDr>")
        .unwrap()
        .to_owned();
    wb.preserved_entries
        .insert("xl/drawings/drawing1.xml".into(), truncated.into_bytes());

    assert_chart_only_drawing(&write_xlsx(&wb).expect("a damaged preserved drawing is not fatal"));
}

#[test]
fn preserved_media_content_types_use_lowercased_extensions() {
    let mut wb = minimal_workbook("Sheet1", Vec::new());
    for path in [
        "xl/media/image1.jpg",
        "xl/media/image2.PNG",
        "xl/drawings/vmlDrawing1.vml",
        "xl/printerSettings/printerSettings1.bin",
    ] {
        wb.preserved_entries.insert(path.into(), vec![0]);
    }

    let entries = zip_map(&write_xlsx(&wb).unwrap());
    let ct = ContentTypes::parse(&entries["[Content_Types].xml"]).unwrap();
    assert_eq!(
        ct.defaults.keys().map(String::as_str).collect::<Vec<_>>(),
        ["jpg", "png", "rels", "vml", "xml"]
    );
    assert_eq!(ct.defaults["jpg"], "image/jpeg");
    assert_eq!(ct.defaults["png"], "image/png");
    assert_eq!(ct.defaults["vml"], CT_VML_DRAWING);
    assert_eq!(
        ct.overrides["/xl/printerSettings/printerSettings1.bin"],
        CT_PRINTER_SETTINGS
    );
    assert_eq!(
        ct.overrides
            .values()
            .filter(|v| *v == CT_PRINTER_SETTINGS)
            .count(),
        1
    );
}

// -----------------------------------------------------------------------
// Workbook part graph: pivotCaches, externalReferences, theme1.xml and
// repointed sheet rels
// -----------------------------------------------------------------------

/// Deserializes a test fixture value from JSON.
fn from_json<T: serde::de::DeserializeOwned>(v: serde_json::Value) -> T {
    serde_json::from_value(v).expect("fixture JSON should match the model")
}

/// The parsed `xl/_rels/workbook.xml.rels` of a written package.
fn workbook_rels(entries: &HashMap<String, Vec<u8>>) -> Relationships {
    Relationships::parse(&entries["xl/_rels/workbook.xml.rels"]).unwrap()
}

/// One sheet with one pivot table over `table_cache_id`, a `<calcPr>`, and
/// one pivot cache (with records) per entry of `cache_ids`.
fn pivot_workbook(cache_ids: &[Option<u32>], table_cache_id: u32) -> WorkbookData {
    use crate::ooxml::pivot_table::PivotCacheDefinitionData;
    use serde_json::json;

    let mut wb = minimal_workbook("Sheet1", Vec::new());
    wb.calc_pr = Some(crate::ooxml::workbook::CalcPr {
        calc_id: Some(191_029),
        ..Default::default()
    });
    wb.sheets[0].worksheet.pivot_tables = from_json(json!([
        { "name": "PT1", "location": { "ref": "D1:E3" }, "cacheId": table_cache_id },
    ]));
    wb.pivot_caches = cache_ids
        .iter()
        .map(|&cache_id| PivotCacheDefinitionData {
            cache_id,
            ..from_json(json!({
                "source": { "ref": "A1:B5", "sheet": "Sheet1" },
                "fields": [{ "name": "A" }],
            }))
        })
        .collect();
    wb.pivot_cache_records = cache_ids
        .iter()
        .map(|_| from_json(json!({ "records": [[{ "type": "number", "v": 1.0 }]] })))
        .collect();
    wb
}

#[test]
fn pivot_cache_is_listed_after_calc_pr_and_targets_its_definition() {
    let bytes = write_xlsx(&pivot_workbook(&[Some(7)], 7)).unwrap();
    let entries = zip_map(&bytes);
    let wb_xml = part_text(&entries, "xl/workbook.xml");
    let caches = WbXml::parse(wb_xml.as_bytes()).unwrap().pivot_caches;
    assert_eq!(caches.len(), 1, "{wb_xml}");
    assert_eq!(caches[0].cache_id, 7);
    assert!(
        wb_xml.contains(&format!(
            r#"<pivotCaches><pivotCache cacheId="7" r:id="{}"/></pivotCaches>"#,
            caches[0].r_id
        )),
        "{wb_xml}"
    );
    assert!(
        wb_xml.find("<calcPr").unwrap() < wb_xml.find("<pivotCaches>").unwrap(),
        "{wb_xml}"
    );
    let rel = workbook_rels(&entries)
        .get_by_id(&caches[0].r_id)
        .cloned()
        .unwrap();
    assert_eq!(rel.rel_type, REL_PIVOT_CACHE_DEF);
    assert_eq!(rel.target, "pivotCache/pivotCacheDefinition1.xml");

    let wb2 = crate::reader::read_xlsx(&bytes).unwrap();
    assert_eq!(wb2.pivot_caches[0].cache_id, Some(7));
    assert_eq!(wb2.sheets[0].worksheet.pivot_tables[0].cache_id, 7);
    let json: serde_json::Value =
        serde_json::from_str(&crate::reader::read_xlsx_json(&bytes).unwrap()).unwrap();
    assert_eq!(json["pivotCaches"][0]["cacheId"], 7);

    // The read result writes the same cacheId again.
    let again = zip_map(&write_xlsx(&wb2).unwrap());
    assert!(part_text(&again, "xl/workbook.xml").contains(r#"<pivotCache cacheId="7" "#));
}

#[test]
fn pivot_cache_without_cache_id_is_listed_by_index() {
    let entries = zip_map(&write_xlsx(&pivot_workbook(&[None, None], 1)).unwrap());
    let wb_xml = part_text(&entries, "xl/workbook.xml");
    let caches = WbXml::parse(wb_xml.as_bytes()).unwrap().pivot_caches;
    assert_eq!(
        caches.iter().map(|c| c.cache_id).collect::<Vec<_>>(),
        [0, 1],
        "{wb_xml}"
    );
    let rels = workbook_rels(&entries);
    let targets: Vec<&str> = caches
        .iter()
        .map(|c| rels.get_by_id(&c.r_id).unwrap().target.as_ref())
        .collect();
    assert_eq!(
        targets,
        [
            "pivotCache/pivotCacheDefinition1.xml",
            "pivotCache/pivotCacheDefinition2.xml"
        ]
    );
}

#[test]
fn duplicate_or_missing_pivot_cache_ids_are_invalid_workbooks() {
    for (cache_ids, table_cache_id) in [
        (&[Some(3), Some(3)][..], 3),
        // A cache without an id takes its index, here 1.
        (&[Some(1), None][..], 1),
        (&[None][..], 4),
        (&[][..], 0),
    ] {
        let err = write_xlsx(&pivot_workbook(cache_ids, table_cache_id)).unwrap_err();
        assert_eq!(
            err.code(),
            "INVALID_WORKBOOK",
            "{cache_ids:?} / {table_cache_id}: {err}"
        );
    }
}

/// 60,000 pivot caches with ids 0 to 59,999, and 60,000 tables on one sheet
/// naming them in reverse order. Finding each table's cache scanned the ids
/// (tables x caches, 1.8 billion comparisons in the pre-write count alone:
/// 3.4 s in a debug build, 0.34 s in a release one); it is now one lookup in
/// the index by id that `pivot_cache_ids` builds once per write (26 ms and
/// 4 ms for both steps).
#[test]
fn pivot_tables_find_their_cache_by_one_lookup() {
    use crate::ooxml::pivot_table::{PivotCacheDefinitionData, PivotTableData};
    const N: u32 = 60_000;
    let base = pivot_workbook(&[Some(0)], 0);
    let cache = &base.pivot_caches[0];
    let table = &base.sheets[0].worksheet.pivot_tables[0];
    let mut wb = minimal_workbook("Sheet1", Vec::new());
    wb.pivot_caches = (0..N)
        .map(|id| PivotCacheDefinitionData {
            cache_id: Some(id),
            ..cache.clone()
        })
        .collect();
    wb.sheets[0].worksheet.pivot_tables = (0..N)
        .rev()
        .map(|cache_id| PivotTableData {
            cache_id,
            ..table.clone()
        })
        .collect();

    let started = std::time::Instant::now();
    let (ids, index) = pivot_cache_ids(&wb).unwrap();
    check_pivot_additions(&wb, &index).unwrap();
    let elapsed = started.elapsed();
    eprintln!("{N} tables over {N} caches: {elapsed:?}");
    assert!(
        elapsed < std::time::Duration::from_secs(1),
        "took {elapsed:?}"
    );
    assert_eq!(ids, (0..N).collect::<Vec<_>>());
    for table in &wb.sheets[0].worksheet.pivot_tables {
        let found = pivot_cache(&wb, &index, table.cache_id).expect("its cache");
        assert_eq!(found.cache_id, Some(table.cache_id));
    }
    // A cache without an id is found under its index.
    let wb = pivot_workbook(&[Some(9), None, Some(4)], 1);
    let (ids, index) = pivot_cache_ids(&wb).unwrap();
    assert_eq!(ids, [9, 1, 4]);
    assert_eq!(pivot_cache(&wb, &index, 1), Some(&wb.pivot_caches[1]));
    assert_eq!(pivot_cache(&wb, &index, 1).unwrap().cache_id, None);
    assert!(pivot_cache(&wb, &index, 2).is_none());
}

/// Caches listed as cacheId 5 (a 2-item field), 2 (a 3-item field) and 8
/// (a 1-item field); one sheet whose tables use 2, 5 and 8, the first two
/// built in code (a row field without items), the third as Excel saves one.
fn authored_pivot_workbook() -> WorkbookData {
    use crate::ooxml::pivot_table::PivotCacheDefinitionData;
    use serde_json::json;

    let cache = |cache_id: u32, shared: usize| PivotCacheDefinitionData {
        cache_id: Some(cache_id),
        ..from_json(json!({
            "source": { "ref": "A1:B5", "sheet": "Sheet1" },
            "fields": [
                {
                    "name": "Region",
                    "sharedItems": (0..shared)
                        .map(|i| json!({ "type": "string", "v": format!("R{i}") }))
                        .collect::<Vec<_>>(),
                },
                { "name": "Amount" },
            ],
        }))
    };
    let table = |name: &str, cache_id: u32, items: serde_json::Value| {
        json!({
            "name": name,
            "cacheId": cache_id,
            "location": { "ref": "D1:E4" },
            "pivotFields": [{ "axis": "axisRow", "items": items }, { "dataField": true }],
            "rowFields": [{ "x": 0 }],
            "dataFields": [{ "fld": 1, "subtotal": "sum" }],
        })
    };
    let mut wb = minimal_workbook("Sheet1", Vec::new());
    wb.pivot_caches = vec![cache(5, 2), cache(2, 3), cache(8, 1)];
    wb.sheets[0].worksheet.pivot_tables = from_json(json!([
        table("Two", 2, json!([])),
        table("Five", 5, json!([])),
        table("Eight", 8, json!([{ "x": 0 }, { "t": "default" }])),
    ]));
    wb
}

#[test]
fn authored_pivot_fields_get_the_items_of_the_cache_their_cache_id_names() {
    let wb = authored_pivot_workbook();
    let before = serde_json::to_string(&wb).unwrap();
    let entries = zip_map(&write_xlsx(&wb).unwrap());
    assert_eq!(
        serde_json::to_string(&wb).unwrap(),
        before,
        "the model is unchanged"
    );

    use std::fmt::Write as _;

    // Tables are written in sheet order; caches in pivotCaches order.
    let item_x = |n: usize| {
        (0..n).fold(String::new(), |mut out, i| {
            write!(out, r#"<item x="{i}"/>"#).unwrap();
            out
        })
    };
    for (table, cache_def, field) in [
        (
            1,
            2,
            format!(
                r#"<pivotField axis="axisRow"><items count="4">{}<item t="default"/></items></pivotField>"#,
                item_x(3)
            ),
        ),
        (
            2,
            1,
            format!(
                r#"<pivotField axis="axisRow"><items count="3">{}<item t="default"/></items></pivotField>"#,
                item_x(2)
            ),
        ),
        (
            3,
            3,
            r#"<pivotField axis="axisRow"><items count="2"><item x="0"/><item t="default"/></items></pivotField>"#
                .to_owned(),
        ),
    ] {
        let path = format!("xl/pivotTables/pivotTable{table}.xml");
        let xml = part_text(&entries, &path);
        assert!(xml.contains(&field), "{path}: {xml}");
        let rels = Relationships::parse(&entries[&rels_path(&path)]).unwrap();
        assert_eq!(
            rels.get_by_id("rId1").unwrap().target,
            format!("../pivotCache/pivotCacheDefinition{cache_def}.xml"),
            "{path}"
        );
    }
    // A cache that a table built in code uses refreshes on load, so Excel
    // lays the table out and shows its values; the Excel-made table's cache
    // does not.
    for (cache_def, refresh) in [(1, true), (2, true), (3, false)] {
        let path = format!("xl/pivotCache/pivotCacheDefinition{cache_def}.xml");
        let xml = part_text(&entries, &path);
        assert_eq!(
            xml.contains(r#" refreshOnLoad="1""#),
            refresh,
            "{path}: {xml}"
        );
    }
}

#[test]
fn excel_pivot_parts_are_written_as_before() {
    use crate::ooxml::pivot_table::tests::{
        EXCEL_TWO_AXES_CACHE, EXCEL_TWO_AXES_CACHE_WRITTEN, EXCEL_TWO_AXES_TABLE,
        EXCEL_TWO_AXES_TABLE_WRITTEN,
    };
    use crate::ooxml::pivot_table::{
        CacheValue, PivotCacheDefinitionData, PivotCacheRecordsData, PivotTableData,
    };

    let mut wb = minimal_workbook("Sheet1", Vec::new());
    wb.sheets[0].worksheet.pivot_tables =
        vec![PivotTableData::parse(EXCEL_TWO_AXES_TABLE.as_bytes()).unwrap()];
    wb.pivot_caches = vec![PivotCacheDefinitionData {
        cache_id: Some(9),
        ..PivotCacheDefinitionData::parse(EXCEL_TWO_AXES_CACHE.as_bytes()).unwrap()
    }];
    wb.pivot_cache_records = vec![PivotCacheRecordsData {
        records: vec![vec![
            CacheValue::Index { v: 0 },
            CacheValue::Index { v: 0 },
            CacheValue::Index { v: 0 },
            CacheValue::Number { v: 10.0 },
            CacheValue::Number { v: 1.0 },
        ]],
    }];
    let entries = zip_map(&write_xlsx(&wb).unwrap());
    assert_eq!(
        part_text(&entries, "xl/pivotTables/pivotTable1.xml"),
        EXCEL_TWO_AXES_TABLE_WRITTEN
    );
    assert_eq!(
        part_text(&entries, "xl/pivotCache/pivotCacheDefinition1.xml"),
        EXCEL_TWO_AXES_CACHE_WRITTEN
    );
}

#[test]
fn a_field_of_a_million_shared_items_is_completed() {
    use crate::ooxml::pivot_table::CacheValue;

    // Under the per-field cap (Excel's 1,048,576 unique items) and the
    // workbook ceiling, a large field is written in full.
    let mut wb = authored_pivot_workbook();
    wb.sheets[0].worksheet.pivot_tables.truncate(1);
    wb.pivot_caches[1].fields[0].shared_items = vec![CacheValue::Missing; 1_000_000];
    // A million `<m/>` compress past the reader's default ratio limit.
    let limits = ZipSecurityLimits {
        max_compression_ratio: f64::INFINITY,
        ..ZipSecurityLimits::default()
    };
    let entries: HashMap<String, Vec<u8>> = read_zip_entries(&write_xlsx(&wb).unwrap(), &limits)
        .unwrap()
        .into_iter()
        .collect();
    let xml = part_text(&entries, "xl/pivotTables/pivotTable1.xml");
    assert!(
        xml.contains(r#"<items count="1000001"><item x="0"/><item x="1"/>"#),
        "{}",
        &xml[..xml.len().min(2000)]
    );
    assert!(xml.contains(r#"<item x="999999"/><item t="default"/></items>"#));
}

/// Preserved external links 1, 2, 3 and 10, with only `listed` in
/// `external_references`.
fn external_link_workbook(listed: &[&str]) -> WorkbookData {
    let mut wb = minimal_workbook("Sheet1", Vec::new());
    for n in [1, 2, 3, 10] {
        wb.preserved_entries.insert(
            format!("xl/externalLinks/externalLink{n}.xml"),
            format!("<externalLink><!-- {n} --></externalLink>").into_bytes(),
        );
    }
    wb.external_references = listed.iter().map(|&p| p.to_owned()).collect();
    wb
}

/// The `<externalReference>` targets of a written package, in order.
fn external_reference_targets(entries: &HashMap<String, Vec<u8>>) -> Vec<String> {
    let wb_xml = part_text(entries, "xl/workbook.xml");
    let rels = workbook_rels(entries);
    WbXml::parse(wb_xml.as_bytes())
        .unwrap()
        .external_references
        .iter()
        .map(|r_id| {
            let rel = rels.get_by_id(r_id).unwrap();
            assert_eq!(rel.rel_type, REL_EXTERNAL_LINK);
            rel.target.to_string()
        })
        .collect()
}

#[test]
fn external_references_keep_their_order_and_sort_unlisted_links_numerically() {
    let wb = external_link_workbook(&[
        "xl/externalLinks/externalLink2.xml",
        "xl/externalLinks/externalLink1.xml",
    ]);
    let bytes = write_xlsx(&wb).unwrap();
    let entries = zip_map(&bytes);
    let expected = [
        "externalLinks/externalLink2.xml",
        "externalLinks/externalLink1.xml",
        "externalLinks/externalLink3.xml",
        "externalLinks/externalLink10.xml",
    ];
    assert_eq!(external_reference_targets(&entries), expected);
    let wb_xml = part_text(&entries, "xl/workbook.xml");
    let refs = wb_xml.find("<externalReferences>").unwrap();
    assert!(wb_xml.find("</sheets>").unwrap() < refs, "{wb_xml}");

    // The reader records that order, so a second round trip keeps it.
    let wb2 = crate::reader::read_xlsx(&bytes).unwrap();
    assert_eq!(wb2.external_references, expected.map(|t| format!("xl/{t}")));
    let json: serde_json::Value =
        serde_json::from_str(&crate::reader::read_xlsx_json(&bytes).unwrap()).unwrap();
    assert_eq!(
        json["externalReferences"],
        serde_json::json!(expected.map(|t| format!("xl/{t}")))
    );
    let again = zip_map(&write_xlsx(&wb2).unwrap());
    assert_eq!(external_reference_targets(&again), expected);
}

#[test]
fn external_references_before_defined_names() {
    let mut wb = external_link_workbook(&[]);
    wb.defined_names = vec![crate::ooxml::workbook::DefinedName {
        name: "Ext".into(),
        value: "[1]Sheet1!$A$1".into(),
        ..Default::default()
    }];
    let entries = zip_map(&write_xlsx(&wb).unwrap());
    let wb_xml = part_text(&entries, "xl/workbook.xml");
    assert!(
        wb_xml.find("</externalReferences>").unwrap() < wb_xml.find("<definedNames>").unwrap(),
        "{wb_xml}"
    );
    assert_eq!(external_reference_targets(&entries).len(), 4);
}

/// The single theme relationship of a written package, and the Override.
fn assert_theme_wired(entries: &HashMap<String, Vec<u8>>) {
    let rels = workbook_rels(entries);
    let themes: Vec<&Relationship> = rels.find_by_type(REL_THEME).collect();
    assert_eq!(themes.len(), 1, "{rels:?}");
    assert_eq!(themes[0].target, "theme/theme1.xml");
    let ct = ContentTypes::parse(&entries["[Content_Types].xml"]).unwrap();
    assert_eq!(
        ct.overrides.get("/xl/theme/theme1.xml").map(String::as_str),
        Some(CT_THEME)
    );
    assert!(entries.contains_key("xl/theme/theme1.xml"));
}

#[test]
fn new_workbook_writes_the_office_theme_with_rel_and_override() {
    let bytes = write_xlsx(&minimal_workbook("Sheet1", Vec::new())).unwrap();
    let entries = zip_map(&bytes);
    assert_theme_wired(&entries);
    let theme = part_text(&entries, "xl/theme/theme1.xml");
    assert!(
        theme.contains(r#"<a:minorFont><a:latin typeface="Calibri"/>"#),
        "{theme}"
    );
    assert!(
        theme.contains(r#"<a:accent1><a:srgbClr val="4472C4"/></a:accent1>"#),
        "{theme}"
    );
    let mut read = crate::reader::read_xlsx(&bytes).unwrap();
    assert_eq!(read.theme_colors.as_ref().unwrap().accent1, "4472C4");
    // A generated theme is consumed, so its colours stay editable.
    assert!(!read.preserved_entries.contains_key(THEME_PATH));
    read.theme_colors.as_mut().unwrap().accent1 = "ABCDEF".into();
    let again = crate::reader::read_xlsx(&write_xlsx(&read).unwrap()).unwrap();
    assert_eq!(again.theme_colors.unwrap().accent1, "ABCDEF");
}

#[test]
fn generated_theme_uses_the_workbook_theme_colors() {
    let mut wb = minimal_workbook("Sheet1", Vec::new());
    wb.theme_colors = Some(ThemeColors {
        accent1: "112233".into(),
        ..Default::default()
    });
    let bytes = write_xlsx(&wb).unwrap();
    assert_theme_wired(&zip_map(&bytes));
    let colors = crate::reader::read_xlsx(&bytes)
        .unwrap()
        .theme_colors
        .unwrap();
    assert_eq!(colors.accent1, "112233");
}

#[test]
fn theme_color_that_is_not_hex_is_an_invalid_workbook() {
    let mut wb = minimal_workbook("Sheet1", Vec::new());
    wb.theme_colors = Some(ThemeColors {
        hlink: r#"00"/><x"#.into(),
        ..Default::default()
    });
    let err = write_xlsx(&wb).unwrap_err();
    assert_eq!(err.code(), "INVALID_WORKBOOK", "{err}");
    assert!(err.to_string().contains("hlink"), "{err}");
}

#[test]
fn preserved_theme_keeps_its_bytes_and_gains_rel_and_override() {
    let mut wb = minimal_workbook("Sheet1", Vec::new());
    let custom = b"<a:theme xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" name=\"Custom\"/>".to_vec();
    wb.preserved_entries
        .insert("xl/theme/theme1.xml".into(), custom.clone());
    let entries = zip_map(&write_xlsx(&wb).unwrap());
    assert_theme_wired(&entries);
    assert_eq!(entries["xl/theme/theme1.xml"], custom);
}

/// Writes `wb`, renames the part `from` to `to` the way another producer
/// would name it (patching sheet 2's rels and the content types), and
/// returns the package and the Id of the sheet 2 relationship of type
/// `rel_type`.
fn with_foreign_part_name(
    wb: &WorkbookData,
    from: &str,
    to: &str,
    rel_type: &str,
) -> (Vec<u8>, String) {
    use crate::zip::writer::{ZipEntry, write_zip};

    let from_file = from.rsplit_once('/').unwrap().1;
    let to_file = to.rsplit_once('/').unwrap().1;
    let entries = zip_map(&write_xlsx(wb).unwrap());
    assert!(entries.contains_key(from), "first write lacks {from}");
    let mut out: Vec<ZipEntry> = entries
        .into_iter()
        .map(|(name, data)| match name.as_str() {
            n if n == from => ZipEntry {
                name: to.to_owned(),
                data,
            },
            "xl/worksheets/_rels/sheet2.xml.rels" | "[Content_Types].xml" => ZipEntry {
                data: String::from_utf8(data)
                    .unwrap()
                    .replace(from_file, to_file)
                    .into_bytes(),
                name,
            },
            _ => ZipEntry { name, data },
        })
        .collect();
    out.sort_by(|a, b| a.name.cmp(&b.name));
    let foreign = write_zip(&out).unwrap();
    let rels =
        Relationships::parse(&zip_map(&foreign)["xl/worksheets/_rels/sheet2.xml.rels"]).unwrap();
    let rel = rels.find_by_type(rel_type).next().unwrap();
    assert!(rel.target.ends_with(to_file), "{rel:?}");
    (foreign, rel.id.clone())
}

/// Two sheets; `fill` adds the feature under test to sheet 2.
fn two_sheets(fill: impl FnOnce(&mut WorkbookData)) -> WorkbookData {
    let mut wb = minimal_workbook("S1", Vec::new());
    wb.sheets.push(SheetData {
        name: "S2".into(),
        worksheet: WorksheetXml::default(),
        ..Default::default()
    });
    fill(&mut wb);
    wb
}

/// Sheet 2's rel of `rel_type` names a part this writer would name
/// otherwise. After read -> write the rel keeps its Id, targets the
/// written part, and the data reads back.
fn assert_foreign_part_repointed(
    wb: &WorkbookData,
    canonical: &str,
    foreign: &str,
    rel_type: &str,
    count: fn(&WorkbookData) -> usize,
) {
    let (package, r_id) = with_foreign_part_name(wb, canonical, foreign, rel_type);
    let read = crate::reader::read_xlsx(&package).unwrap();
    assert_eq!(count(&read), 1, "{foreign}: the reader follows the rel");

    let out = write_xlsx(&read).unwrap();
    let entries = zip_map(&out);
    let rels = assert_rel_targets_exist(&entries, "xl/worksheets/sheet2.xml");
    let typed: Vec<&Relationship> = rels.find_by_type(rel_type).collect();
    assert_eq!(typed.len(), 1, "{rels:?}");
    assert_eq!(typed[0].id, r_id, "{foreign}: the Id is kept");
    assert_eq!(
        count(&crate::reader::read_xlsx(&out).unwrap()),
        1,
        "{foreign}: the data survives the round trip"
    );
}

#[test]
fn foreign_comments_part_name_is_repointed() {
    let wb = two_sheets(|wb| {
        wb.sheets[1].worksheet.comments = from_json(serde_json::json!([
            { "cellRef": "A1", "author": "me", "text": "hello" },
        ]));
    });
    assert_foreign_part_repointed(
        &wb,
        "xl/comments2.xml",
        "xl/comments1.xml",
        REL_COMMENTS,
        |wb| wb.sheets[1].worksheet.comments.len(),
    );
}

#[test]
fn foreign_threaded_comments_part_name_is_repointed() {
    let wb = two_sheets(|wb| {
        wb.sheets[1].worksheet.threaded_comments = from_json(serde_json::json!([{
            "id": "{00000000-0000-0000-0000-000000000001}",
            "refCell": "A1",
            "personId": "{00000000-0000-0000-0000-0000000000AA}",
            "text": "hello",
            "timestamp": "2026-11-11T00:00:00.00",
        }]));
    });
    assert_foreign_part_repointed(
        &wb,
        "xl/threadedComments/threadedComment2.xml",
        "xl/threadedComments/threadedComment9.xml",
        REL_THREADED_COMMENTS,
        |wb| wb.sheets[1].worksheet.threaded_comments.len(),
    );
}

#[test]
fn foreign_slicer_part_name_is_repointed() {
    let wb = two_sheets(|wb| {
        wb.sheets[1].worksheet.slicers = from_json(serde_json::json!([
            { "name": "Slicer_A", "cacheName": "Slicer_A" },
        ]));
    });
    assert_foreign_part_repointed(
        &wb,
        "xl/slicers/slicer1.xml",
        "xl/slicers/slicer7.xml",
        REL_SLICER,
        |wb| wb.sheets[1].worksheet.slicers.len(),
    );
}

#[test]
fn foreign_timeline_part_name_is_repointed() {
    let wb = two_sheets(|wb| {
        wb.sheets[1].worksheet.timelines = from_json(serde_json::json!([
            { "name": "Timeline_A", "cacheName": "NativeTimeline_A" },
        ]));
    });
    assert_foreign_part_repointed(
        &wb,
        "xl/timelines/timeline1.xml",
        "xl/timelines/timeline7.xml",
        REL_TIMELINE,
        |wb| wb.sheets[1].worksheet.timelines.len(),
    );
}

#[test]
fn cleared_comments_leave_no_comments_rel() {
    let wb = two_sheets(|wb| {
        wb.sheets[1].worksheet.comments = from_json(serde_json::json!([
            { "cellRef": "A1", "author": "me", "text": "hello" },
        ]));
    });
    let (package, _) =
        with_foreign_part_name(&wb, "xl/comments2.xml", "xl/comments1.xml", REL_COMMENTS);
    let mut read = crate::reader::read_xlsx(&package).unwrap();
    read.sheets[1].worksheet.comments.clear();

    let entries = zip_map(&write_xlsx(&read).unwrap());
    if let Some(data) = entries.get("xl/worksheets/_rels/sheet2.xml.rels") {
        let rels = Relationships::parse(data).unwrap();
        assert!(rels.find_by_type(REL_COMMENTS).next().is_none(), "{rels:?}");
    }
    assert!(!part_text(&entries, "[Content_Types].xml").contains("comments"));
}

#[test]
fn extra_rels_of_a_regenerated_type_are_dropped() {
    let mut wb = minimal_workbook("Sheet1", Vec::new());
    wb.sheets[0].worksheet.comments = from_json(serde_json::json!([
        { "cellRef": "A1", "author": "me", "text": "hello" },
    ]));
    let mut rels = Relationships::new();
    rels.add("rId3", REL_COMMENTS, "../foo.xml");
    rels.add("rId5", REL_COMMENTS, "../bar.xml");
    wb.preserved_entries.insert(
        "xl/worksheets/_rels/sheet1.xml.rels".into(),
        rels.to_xml().unwrap(),
    );

    let entries = zip_map(&write_xlsx(&wb).unwrap());
    let rels = assert_rel_targets_exist(&entries, "xl/worksheets/sheet1.xml");
    let comments: Vec<&Relationship> = rels.find_by_type(REL_COMMENTS).collect();
    assert_eq!(comments.len(), 1, "{rels:?}");
    assert_eq!(comments[0].id, "rId3");
    assert_eq!(comments[0].target, "../comments1.xml");
}

// -----------------------------------------------------------------------
// Package graph fidelity (fresh audit AW-01): worksheet .rels keyed by
// position, pivot table .rels, source content types, custom XML rels,
// drawing targets, the persons part
// -----------------------------------------------------------------------

/// Zips `entries` in name order.
fn rezip(entries: HashMap<String, Vec<u8>>) -> Vec<u8> {
    let mut out: Vec<ZipEntry> = entries
        .into_iter()
        .map(|(name, data)| ZipEntry { name, data })
        .collect();
    out.sort_by(|a, b| a.name.cmp(&b.name));
    write_zip(&out).unwrap()
}

/// Sheet A (with a picture) and sheet B, with A's worksheet part renamed to
/// `xl/worksheets/sheet2.xml` and B's to `sheet1.xml`: the layout POI leaves
/// after `removeSheetAt`.
fn swapped_sheet_parts_package() -> Vec<u8> {
    let wb = two_sheets(|wb| {
        wb.sheets[0].name = "A".into();
        wb.sheets[1].name = "B".into();
        preserve_drawing(
            wb,
            1,
            &pic_anchor("twoCellAnchor", "Img1", "rId1"),
            &["image1.png"],
        );
    });
    let entries = zip_map(&write_xlsx(&wb).unwrap());
    let swapped = entries
        .into_iter()
        .map(|(name, data)| match name.as_str() {
            "xl/worksheets/sheet1.xml" => ("xl/worksheets/sheet2.xml".to_owned(), data),
            "xl/worksheets/sheet2.xml" => ("xl/worksheets/sheet1.xml".to_owned(), data),
            "xl/worksheets/_rels/sheet1.xml.rels" => {
                ("xl/worksheets/_rels/sheet2.xml.rels".to_owned(), data)
            }
            "xl/_rels/workbook.xml.rels" => {
                let xml = String::from_utf8(data)
                    .unwrap()
                    .replace("worksheets/sheet1.xml", "worksheets/sheetA.xml")
                    .replace("worksheets/sheet2.xml", "worksheets/sheet1.xml")
                    .replace("worksheets/sheetA.xml", "worksheets/sheet2.xml");
                (name, xml.into_bytes())
            }
            _ => (name, data),
        })
        .collect();
    rezip(swapped)
}

#[test]
fn worksheet_rels_follow_the_sheet_position_not_the_source_part_name() {
    let package = swapped_sheet_parts_package();
    let source = zip_map(&package);
    assert!(part_text(&source, "xl/worksheets/sheet2.xml").contains("<drawing "));
    assert!(!source.contains_key("xl/worksheets/_rels/sheet1.xml.rels"));

    // Both readers key A's .rels by A's position (1).
    let read = crate::reader::read_xlsx(&package).unwrap();
    assert_eq!(read.sheets[0].name, "A");
    assert!(
        read.preserved_entries
            .contains_key("xl/worksheets/_rels/sheet1.xml.rels")
    );
    assert!(
        !read
            .preserved_entries
            .contains_key("xl/worksheets/_rels/sheet2.xml.rels")
    );
    let json: serde_json::Value =
        serde_json::from_str(&crate::reader::read_xlsx_json(&package).unwrap()).unwrap();
    let preserved = json["preservedEntries"].as_object().unwrap();
    assert!(preserved.contains_key("xl/worksheets/_rels/sheet1.xml.rels"));
    assert!(!preserved.contains_key("xl/worksheets/_rels/sheet2.xml.rels"));

    // The picture stays on A after write -> read.
    let out = write_xlsx(&read).unwrap();
    let entries = zip_map(&out);
    assert!(
        part_text(&entries, "xl/worksheets/sheet1.xml").contains("<drawing r:id="),
        "A keeps its drawing"
    );
    assert!(
        !part_text(&entries, "xl/worksheets/sheet2.xml").contains("<drawing"),
        "B gains no drawing"
    );
    let drawing = sheet_drawing_path(&entries, 1);
    assert!(part_text(&entries, &drawing).contains(r#"name="Img1""#));
    let again = crate::reader::read_xlsx(&out).unwrap();
    assert_eq!(again.sheets[0].name, "A");
    assert!(
        again
            .preserved_entries
            .contains_key("xl/worksheets/_rels/sheet1.xml.rels")
    );
}

#[test]
fn pivot_table_rels_target_its_cache_definition() {
    for (cache_ids, table_cache_id, target) in [
        (&[Some(7)][..], 7, "../pivotCache/pivotCacheDefinition1.xml"),
        (
            &[Some(3), Some(7)][..],
            7,
            "../pivotCache/pivotCacheDefinition2.xml",
        ),
    ] {
        let bytes = write_xlsx(&pivot_workbook(cache_ids, table_cache_id)).unwrap();
        let entries = zip_map(&bytes);
        let rels =
            Relationships::parse(&entries["xl/pivotTables/_rels/pivotTable1.xml.rels"]).unwrap();
        assert_eq!(rels.len(), 1, "{rels:?}");
        let rel = rels.iter().next().unwrap();
        assert_eq!(rel.id, "rId1");
        assert_eq!(rel.rel_type, REL_PIVOT_CACHE_DEF);
        assert_eq!(rel.target, target);
        assert_rel_targets_exist(&entries, "xl/pivotTables/pivotTable1.xml");

        // The reader consumes the .rels with the table; a second write has
        // exactly one.
        let read = crate::reader::read_xlsx(&bytes).unwrap();
        assert!(
            !read
                .preserved_entries
                .keys()
                .any(|k| k.starts_with("xl/pivotTables/")),
            "{:?}",
            read.preserved_entries.keys().collect::<Vec<_>>()
        );
        let again = zip_map(&write_xlsx(&read).unwrap());
        let table_rels: Vec<&String> = again
            .keys()
            .filter(|k| k.starts_with("xl/pivotTables/_rels/"))
            .collect();
        assert_eq!(table_rels, ["xl/pivotTables/_rels/pivotTable1.xml.rels"]);
        let rels =
            Relationships::parse(&again["xl/pivotTables/_rels/pivotTable1.xml.rels"]).unwrap();
        assert_eq!(rels.len(), 1, "{rels:?}");
        assert_eq!(rels.iter().next().unwrap().target, target);
    }
}

const CT_CHART_EX_LITERAL: &str = "application/vnd.ms-office.chartex+xml";
const CT_XLSX_LITERAL: &str = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
const CT_OLE_OBJECT_LITERAL: &str = "application/vnd.openxmlformats-officedocument.oleObject";
const CT_CTRL_PROPS_LITERAL: &str = "application/vnd.ms-excel.controlproperties+xml";

#[test]
fn preserved_embeddings_chart_ex_and_hd_photo_parts_get_content_types() {
    let mut wb = minimal_workbook("Sheet1", Vec::new());
    for path in [
        "xl/charts/chartEx1.xml",
        "xl/charts/chart1.xml",
        "xl/embeddings/Book1.xlsx",
        "xl/embeddings/Doc1.docx",
        "xl/embeddings/Deck1.pptx",
        "xl/embeddings/oleObject1.bin",
        "xl/media/hdphoto1.wdp",
    ] {
        wb.preserved_entries.insert(path.into(), vec![0]);
    }
    let entries = zip_map(&write_xlsx(&wb).unwrap());
    let ct = ContentTypes::parse(&entries["[Content_Types].xml"]).unwrap();
    assert_eq!(ct.overrides["/xl/charts/chartEx1.xml"], CT_CHART_EX_LITERAL);
    assert_eq!(ct.overrides["/xl/charts/chart1.xml"], CT_CHART);
    assert_eq!(ct.defaults["xlsx"], CT_XLSX_LITERAL);
    assert_eq!(
        ct.defaults["docx"],
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    );
    assert_eq!(
        ct.defaults["pptx"],
        "application/vnd.openxmlformats-officedocument.presentationml.presentation"
    );
    assert_eq!(ct.defaults["wdp"], "image/vnd.ms-photo");
    assert_eq!(
        ct.overrides["/xl/embeddings/oleObject1.bin"],
        CT_OLE_OBJECT_LITERAL
    );
    assert!(!ct.defaults.contains_key("bin"), "{:?}", ct.defaults);
}

#[test]
fn source_content_types_of_preserved_parts_survive_read_write() {
    let mut wb = minimal_workbook("Sheet1", Vec::new());
    for path in [
        "xl/charts/chartEx1.xml",
        "xl/embeddings/Book1.xlsx",
        "xl/embeddings/oleObject1.bin",
        "xl/ctrlProps/ctrlProp1.xml",
    ] {
        wb.preserved_entries.insert(path.into(), b"<x/>".to_vec());
    }
    let mut entries = zip_map(&write_xlsx(&wb).unwrap());
    // The source package declares each part's type as Excel does.
    let mut source_ct = ContentTypes::parse(&entries["[Content_Types].xml"]).unwrap();
    source_ct.add_override("/xl/charts/chartEx1.xml", CT_CHART_EX_LITERAL);
    source_ct.add_default("xlsx", CT_XLSX_LITERAL);
    source_ct.add_override("/xl/embeddings/oleObject1.bin", CT_OLE_OBJECT_LITERAL);
    source_ct.add_override("/xl/ctrlProps/ctrlProp1.xml", CT_CTRL_PROPS_LITERAL);
    entries.insert("[Content_Types].xml".into(), source_ct.to_xml().unwrap());
    let package = rezip(entries);

    let json: serde_json::Value =
        serde_json::from_str(&crate::reader::read_xlsx_json(&package).unwrap()).unwrap();
    let recorded = &json["preservedContentTypes"];
    assert_eq!(
        recorded["xl/ctrlProps/ctrlProp1.xml"],
        CT_CTRL_PROPS_LITERAL
    );
    assert_eq!(recorded["xl/embeddings/Book1.xlsx"], CT_XLSX_LITERAL);
    assert_eq!(recorded["xl/charts/chartEx1.xml"], CT_CHART_EX_LITERAL);

    for read in [
        crate::reader::read_xlsx(&package).unwrap(),
        from_json::<WorkbookData>(json.clone()),
    ] {
        let out = zip_map(&write_xlsx(&read).unwrap());
        let ct = ContentTypes::parse(&out["[Content_Types].xml"]).unwrap();
        assert_eq!(ct.overrides["/xl/charts/chartEx1.xml"], CT_CHART_EX_LITERAL);
        assert_eq!(ct.defaults["xlsx"], CT_XLSX_LITERAL);
        assert!(
            !ct.overrides.contains_key("/xl/embeddings/Book1.xlsx"),
            "a type the Default gives needs no Override"
        );
        assert_eq!(
            ct.overrides["/xl/embeddings/oleObject1.bin"],
            CT_OLE_OBJECT_LITERAL
        );
        assert_eq!(
            ct.overrides["/xl/ctrlProps/ctrlProp1.xml"],
            CT_CTRL_PROPS_LITERAL
        );
    }
}

const REL_CUSTOM_XML_LITERAL: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml";
const REL_CUSTOM_PROPERTIES_LITERAL: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties";

/// The targets of the relationships of `rel_type` in the .rels at `path`.
fn rel_targets(entries: &HashMap<String, Vec<u8>>, path: &str, rel_type: &str) -> Vec<String> {
    Relationships::parse(&entries[path])
        .unwrap()
        .find_by_type(rel_type)
        .map(|rel| rel.target.to_string())
        .collect()
}

#[test]
fn custom_xml_items_and_custom_properties_keep_their_relationships() {
    let mut wb = minimal_workbook("Sheet1", Vec::new());
    for (path, data) in [
        ("customXml/item1.xml", &b"<root/>"[..]),
        ("customXml/item2.xml", b"<root2/>"),
        ("customXml/itemProps1.xml", b"<ds:datastoreItem/>"),
        ("customXml/itemProps2.xml", b"<ds:datastoreItem/>"),
        ("customXml/_rels/item1.xml.rels", b"<Relationships/>"),
        ("docProps/custom.xml", b"<Properties/>"),
    ] {
        wb.preserved_entries.insert(path.into(), data.to_vec());
    }
    let first = write_xlsx(&wb).unwrap();
    let second = write_xlsx(&crate::reader::read_xlsx(&first).unwrap()).unwrap();
    for bytes in [first, second] {
        let entries = zip_map(&bytes);
        assert_eq!(
            rel_targets(
                &entries,
                "xl/_rels/workbook.xml.rels",
                REL_CUSTOM_XML_LITERAL
            ),
            ["../customXml/item1.xml", "../customXml/item2.xml"]
        );
        assert_eq!(
            rel_targets(&entries, "_rels/.rels", REL_CUSTOM_PROPERTIES_LITERAL),
            ["docProps/custom.xml"]
        );
        let ct = ContentTypes::parse(&entries["[Content_Types].xml"]).unwrap();
        assert_eq!(
            ct.overrides["/docProps/custom.xml"],
            "application/vnd.openxmlformats-officedocument.custom-properties+xml"
        );
        assert_rel_targets_exist(&entries, "xl/workbook.xml");
        let root = Relationships::parse(&entries["_rels/.rels"]).unwrap();
        for rel in &root {
            assert!(entries.contains_key(rel.target.as_ref()), "{rel:?}");
        }
    }
}

const CT_CUSTOM_PROPERTIES_LITERAL: &str =
    "application/vnd.openxmlformats-officedocument.custom-properties+xml";

/// A 1.2.x package typed `docProps/custom.xml` only through the `xml` Default
/// (it wrote no Override for it): the source type the reader records equals
/// that Default, and the custom-properties Override the writer sets for the
/// package relationship must survive.
#[test]
fn custom_properties_override_survives_a_source_typed_by_the_xml_default() {
    let mut wb = minimal_workbook("Sheet1", Vec::new());
    wb.preserved_entries
        .insert("docProps/custom.xml".into(), b"<Properties/>".to_vec());
    let mut entries = zip_map(&write_xlsx(&wb).unwrap());
    let mut source_ct = ContentTypes::parse(&entries["[Content_Types].xml"]).unwrap();
    assert!(source_ct.overrides.remove("/docProps/custom.xml").is_some());
    entries.insert("[Content_Types].xml".into(), source_ct.to_xml().unwrap());
    let package = rezip(entries);

    let read = crate::reader::read_xlsx(&package).unwrap();
    assert_eq!(
        read.preserved_content_types["docProps/custom.xml"],
        "application/xml"
    );
    let out = zip_map(&write_xlsx(&read).unwrap());
    let ct = ContentTypes::parse(&out["[Content_Types].xml"]).unwrap();
    assert_eq!(
        ct.overrides["/docProps/custom.xml"],
        CT_CUSTOM_PROPERTIES_LITERAL
    );
    assert_eq!(
        rel_targets(&out, "_rels/.rels", REL_CUSTOM_PROPERTIES_LITERAL),
        ["docProps/custom.xml"]
    );
}

/// A preserved drawing recorded as `application/xml` (a source Override whose
/// PartName differs only in case misses, so the `xml` Default is recorded)
/// keeps the drawing Override the writer's own rule sets.
#[test]
fn preserved_drawing_recorded_as_the_xml_default_keeps_the_drawing_type() {
    let mut wb = minimal_workbook("Sheet1", Vec::new());
    wb.preserved_entries
        .insert("xl/drawings/drawing1.xml".into(), b"<xdr:wsDr/>".to_vec());
    wb.preserved_content_types
        .insert("xl/drawings/drawing1.xml".into(), "application/xml".into());
    let out = zip_map(&write_xlsx(&wb).unwrap());
    let ct = ContentTypes::parse(&out["[Content_Types].xml"]).unwrap();
    assert_eq!(ct.defaults["xml"], "application/xml");
    assert_eq!(ct.overrides["/xl/drawings/drawing1.xml"], CT_DRAWING);
}

#[test]
fn pivot_cache_definition_names_its_records_by_r_id() {
    let entries = zip_map(&write_xlsx(&pivot_workbook(&[Some(7)], 7)).unwrap());
    let definition = part_text(&entries, "xl/pivotCache/pivotCacheDefinition1.xml");
    assert!(
        definition.contains(
            r#"xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId1""#
        ),
        "{definition}"
    );
    let cache_rels = assert_rel_targets_exist(&entries, "xl/pivotCache/pivotCacheDefinition1.xml");
    assert_eq!(
        cache_rels.get_by_id("rId1").unwrap().rel_type,
        REL_PIVOT_CACHE_REC
    );

    let mut wb = pivot_workbook(&[Some(7)], 7);
    wb.pivot_cache_records.clear();
    let entries = zip_map(&write_xlsx(&wb).unwrap());
    let definition = part_text(&entries, "xl/pivotCache/pivotCacheDefinition1.xml");
    assert!(!definition.contains("r:id="), "{definition}");
    assert!(!entries.contains_key("xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels"));
}

#[test]
fn chart_drawing_never_takes_a_non_drawing_target() {
    let theme = b"<a:theme xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" name=\"Custom\"/>".to_vec();
    let mut wb = make_workbook_with_charts(vec![bar_chart(4)]);
    wb.preserved_entries
        .insert(THEME_PATH.into(), theme.clone());
    wb.preserved_entries.insert(
        "xl/worksheets/_rels/sheet1.xml.rels".into(),
        sheet_drawing_rels("../theme/theme1.xml"),
    );
    let entries = zip_map(&write_xlsx(&wb).unwrap());
    assert_eq!(entries[THEME_PATH], theme);
    assert_theme_wired(&entries);
    assert_eq!(sheet_drawing_path(&entries, 1), "xl/drawings/drawing1.xml");
    let ct = ContentTypes::parse(&entries["[Content_Types].xml"]).unwrap();
    assert_eq!(ct.overrides["/xl/drawings/drawing1.xml"], CT_DRAWING);

    // A target naming a part the writer generates no longer collides with it.
    let mut wb = make_workbook_with_charts(vec![bar_chart(4)]);
    wb.preserved_entries.insert(
        "xl/worksheets/_rels/sheet1.xml.rels".into(),
        sheet_drawing_rels("../styles.xml"),
    );
    let entries = zip_map(&write_xlsx(&wb).unwrap());
    assert_eq!(sheet_drawing_path(&entries, 1), "xl/drawings/drawing1.xml");
}

#[test]
fn image_only_drawing_rel_to_a_generated_non_drawing_part_is_not_referenced() {
    let mut wb = minimal_workbook("Sheet1", Vec::new());
    wb.preserved_entries.insert(
        "xl/worksheets/_rels/sheet1.xml.rels".into(),
        sheet_drawing_rels("../styles.xml"),
    );
    let entries = zip_map(&write_xlsx(&wb).unwrap());
    let sheet = part_text(&entries, "xl/worksheets/sheet1.xml");
    assert!(!sheet.contains("<drawing"), "{sheet}");
}

#[test]
fn malformed_preserved_persons_part_is_superseded_by_the_generated_one() {
    let mut wb = minimal_workbook("Sheet1", Vec::new());
    wb.preserved_entries.insert(
        "xl/persons/person.xml".into(),
        b"<personList><a></b>".to_vec(),
    );
    wb.persons = from_json(serde_json::json!([
        { "id": "{00000000-0000-0000-0000-0000000000AA}", "displayName": "Ann" },
    ]));
    let bytes = write_xlsx(&wb).unwrap();
    let names: Vec<String> = read_zip_entries(&bytes, &ZipSecurityLimits::default())
        .unwrap()
        .into_keys()
        .filter(|name| name == "xl/persons/person.xml")
        .collect();
    assert_eq!(names.len(), 1);
    let persons = part_text(&zip_map(&bytes), "xl/persons/person.xml");
    assert!(persons.contains(r#"displayName="Ann""#), "{persons}");
    assert_eq!(crate::reader::read_xlsx(&bytes).unwrap().persons.len(), 1);
}

// ---------------------------------------------------------------------------
// One sheet holding a table, a note, a chart and more, as Excel saves it
// ---------------------------------------------------------------------------

/// The features of [`mixed_sheet`]'s one sheet.
#[derive(Debug, Clone, Copy, Default)]
struct Mix {
    table: bool,
    note: bool,
    chart: bool,
    picture: bool,
    threaded: bool,
    header_picture: bool,
}

/// `table`, `note` and `chart` alone.
const fn mix(table: bool, note: bool, chart: bool) -> Mix {
    Mix {
        table,
        note,
        chart,
        picture: false,
        threaded: false,
        header_picture: false,
    }
}

/// Every subset of {table, note, chart}, then all three with a picture, a
/// threaded comment or both, and a header picture beside a note and a chart
/// and alone.
const MIXES: [Mix; 12] = [
    mix(true, false, false),
    mix(false, true, false),
    mix(false, false, true),
    mix(true, true, false),
    mix(true, false, true),
    mix(false, true, true),
    mix(true, true, true),
    Mix {
        picture: true,
        ..mix(true, true, true)
    },
    Mix {
        threaded: true,
        ..mix(true, true, true)
    },
    Mix {
        picture: true,
        threaded: true,
        ..mix(true, true, true)
    },
    Mix {
        header_picture: true,
        ..mix(false, true, true)
    },
    Mix {
        header_picture: true,
        ..mix(false, false, false)
    },
];

/// The VML drawing Microsoft Excel 16.0 saved on 2026-11-12 for one note on
/// D1, on one line.
const NOTE_VML: &str = r##"<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel"><o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="1"/></o:shapelayout><v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe"><v:stroke joinstyle="miter"/><v:path gradientshapeok="t" o:connecttype="rect"/></v:shapetype><v:shape id="_x0000_s1025" type="#_x0000_t202" style="position:absolute;margin-left:199.5pt;margin-top:1pt;width:104pt;height:61.5pt;z-index:1;visibility:hidden" fillcolor="infoBackground [80]" strokecolor="none [81]" o:insetmode="auto"><v:fill color2="infoBackground [80]"/><v:shadow color="none [81]" obscured="t"/><v:path o:connecttype="none"/><v:textbox style="mso-direction-alt:auto"><div style="text-align:left"></div></v:textbox><x:ClientData ObjectType="Note"><x:MoveWithCells/><x:SizeWithCells/><x:Anchor>4, 10, 0, 1, 6, 20, 4, 6</x:Anchor><x:AutoFill>False</x:AutoFill><x:Row>0</x:Row><x:Column>3</x:Column></x:ClientData></v:shape></xml>"##;

/// The VML drawing Excel saved the same day for a centre-header picture
/// (`&G`), its picture shape named `CH`, on one line and without the shape
/// type's formulas.
const HEADER_PICTURE_VML: &str = r##"<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel"><o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="1"/></o:shapelayout><v:shapetype id="_x0000_t75" coordsize="21600,21600" o:spt="75" o:preferrelative="t" path="m@4@5l@4@11@9@11@9@5xe" filled="f" stroked="f"><v:stroke joinstyle="miter"/><v:path o:extrusionok="f" gradientshapeok="t" o:connecttype="rect"/><o:lock v:ext="edit" aspectratio="t"/></v:shapetype><v:shape id="CH" o:spid="_x0000_s1025" type="#_x0000_t75" style="position:absolute;margin-left:0;margin-top:0;width:30pt;height:22.5pt;z-index:1"><v:imagedata o:relid="rId1" o:title="img"/><o:lock v:ext="edit" rotation="t"/></v:shape></xml>"##;

/// A .rels part holding `(Id, type, target)` relationships.
fn rels_part(rels: &[(&str, &str, &str)]) -> Vec<u8> {
    use std::fmt::Write as _;
    let mut xml = String::from(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">"#,
    );
    for (id, rel_type, target) in rels {
        write!(
            xml,
            r#"<Relationship Id="{id}" Type="{rel_type}" Target="{target}"/>"#
        )
        .unwrap();
    }
    xml.push_str("</Relationships>");
    xml.into_bytes()
}

/// A text cell (written through the shared string table).
fn text_cell(reference: String, text: &str) -> Cell {
    Cell {
        reference,
        cell_type: CellType::SharedString,
        value: Some(text.into()),
        ..Default::default()
    }
}

/// A one-sheet workbook with `mix`'s features over the data A1:B5, in the
/// state the reader leaves an Excel-saved file in: the model holds the table,
/// notes, chart and threaded comment, and `preserved_entries` the parts it
/// does not model (VML drawings, a picture's drawing) with the sheet .rels
/// naming them under Excel's Ids.
fn mixed_sheet(mix: Mix) -> WorkbookData {
    use serde_json::json;

    let mut rows = vec![Row {
        index: 1,
        cells: vec![
            text_cell("A1".into(), "Name"),
            text_cell("B1".into(), "Value"),
        ],
        ..Default::default()
    }];
    for (r, name) in (2..).zip(["alpha", "beta", "gamma", "delta"]) {
        rows.push(Row {
            index: r,
            cells: vec![
                text_cell(format!("A{r}"), name),
                Cell {
                    reference: format!("B{r}"),
                    cell_type: CellType::Number,
                    value: Some(((r - 1) * 10).to_string()),
                    ..Default::default()
                },
            ],
            ..Default::default()
        });
    }
    let mut wb = minimal_workbook("Sheet1", rows);
    let ws = &mut wb.sheets[0].worksheet;
    ws.dimension = Some("A1:B5".into());
    let mut sheet_rels: Vec<(&str, &str, &str)> = Vec::new();

    if mix.table {
        ws.tables = from_json(json!([{
            "id": 1, "displayName": "Table1", "ref": "A1:B5",
            "columns": [{ "id": 1, "name": "Name" }, { "id": 2, "name": "Value" }],
        }]));
    }
    if mix.note || mix.threaded {
        let mut notes = Vec::new();
        if mix.note {
            notes.push(json!({ "cellRef": "D1", "author": "Ann", "text": "note" }));
        }
        if mix.threaded {
            // Excel gives each threaded comment a note of its own too.
            notes.push(json!({ "cellRef": "E1", "author": "tc={0}", "text": "thread" }));
            ws.threaded_comments = from_json(json!([{
                "id": "{00000000-0000-0000-0000-000000000001}", "refCell": "E1",
                "personId": "{00000000-0000-0000-0000-0000000000AA}",
                "text": "thread", "timestamp": "2026-11-12T00:00:00.00",
            }]));
            wb.persons = from_json(json!([
                { "id": "{00000000-0000-0000-0000-0000000000AA}", "displayName": "Ann" },
            ]));
        }
        ws.comments = from_json(serde_json::Value::Array(notes));
        sheet_rels.push(("rId2", REL_VML_DRAWING, "../drawings/vmlDrawing1.vml"));
        wb.preserved_entries.insert(
            "xl/drawings/vmlDrawing1.vml".into(),
            NOTE_VML.as_bytes().to_vec(),
        );
    }
    if mix.chart {
        ws.charts = vec![bar_chart(4)];
    }
    if mix.picture {
        sheet_rels.push(("rId1", REL_DRAWING, "../drawings/drawing1.xml"));
        wb.preserved_entries.insert(
            "xl/drawings/drawing1.xml".into(),
            drawing_part(&pic_anchor("oneCellAnchor", "Picture 1", "rId1")),
        );
        wb.preserved_entries.insert(
            "xl/drawings/_rels/drawing1.xml.rels".into(),
            image_rels(&["../media/image1.png"]),
        );
        wb.preserved_entries
            .insert("xl/media/image1.png".into(), b"PNG".to_vec());
    }
    if mix.header_picture {
        ws.header_footer = from_json(json!({ "oddHeader": "&C&G" }));
        sheet_rels.push(("rId4", REL_VML_DRAWING, "../drawings/vmlDrawing2.vml"));
        wb.preserved_entries.insert(
            "xl/drawings/vmlDrawing2.vml".into(),
            HEADER_PICTURE_VML.as_bytes().to_vec(),
        );
        wb.preserved_entries.insert(
            "xl/drawings/_rels/vmlDrawing2.vml.rels".into(),
            image_rels(&["../media/image2.png"]),
        );
        wb.preserved_entries
            .insert("xl/media/image2.png".into(), b"PNG".to_vec());
    }
    if !sheet_rels.is_empty() {
        wb.preserved_entries
            .insert(SHEET1_RELS.into(), rels_part(&sheet_rels));
    }
    wb
}

/// The relationship type each `<worksheet>` element naming a part by `r:id`
/// must point at.
fn sheet_element_rel_type(element: &str) -> &'static str {
    match element {
        "drawing" => REL_DRAWING,
        "legacyDrawing" | "legacyDrawingHF" => REL_VML_DRAWING,
        "tablePart" => REL_TABLE,
        "hyperlink" => REL_HYPERLINK,
        other => panic!("unexpected r:id on <{other}>"),
    }
}

/// Asserts the part references of sheet1 in `entries` hold together: each
/// `r:id` in the sheet XML names a relationship of the type its element
/// needs; the sheet .rels has no Id twice, and its internal relationships
/// name parts present in the package, no part twice; each VML drawing it
/// names is named by `<legacyDrawing>` (one with note shapes) or
/// `<legacyDrawingHF>` (one without); and those elements come in
/// CT_Worksheet order.
fn assert_sheet_part_refs(entries: &HashMap<String, Vec<u8>>, label: &str) {
    let sheet = part_text(entries, "xl/worksheets/sheet1.xml");
    let Some(raw_rels) = entries.get(SHEET1_RELS) else {
        assert!(
            !sheet.contains("r:id="),
            "{label}: r:id without rels: {sheet}"
        );
        return;
    };
    let rels = Relationships::parse(raw_rels).unwrap();
    let raw_rels = String::from_utf8_lossy(raw_rels);

    let ids: Vec<&str> = raw_rels
        .split(" Id=\"")
        .skip(1)
        .map(|rest| &rest[..rest.find('"').unwrap()])
        .collect();
    let unique: HashSet<&str> = ids.iter().copied().collect();
    assert_eq!(unique.len(), ids.len(), "{label}: an Id twice: {raw_rels}");

    let mut targets = HashSet::new();
    for rel in rels
        .iter()
        .filter(|rel| rel.target_mode == crate::ooxml::relationships::TargetMode::Internal)
    {
        let path = resolve_target(SHEET_PART, &rel.target);
        assert!(
            entries.contains_key(&path),
            "{label}: {} -> {path} missing",
            rel.id
        );
        assert!(
            targets.insert(path.clone()),
            "{label}: {path} named twice: {raw_rels}"
        );
    }

    let mut named: HashMap<&str, &str> = HashMap::new();
    for (at, _) in sheet.match_indices(" r:id=\"") {
        let element = sheet[..at].rsplit('<').next().unwrap();
        let element = element.split_whitespace().next().unwrap();
        let rest = &sheet[at + " r:id=\"".len()..];
        let id = &rest[..rest.find('"').unwrap()];
        let rel = rels
            .get_by_id(id)
            .unwrap_or_else(|| panic!("{label}: <{element} r:id=\"{id}\"> names no relationship"));
        assert_eq!(
            rel.rel_type,
            sheet_element_rel_type(element),
            "{label}: <{element}>"
        );
        named.insert(id, element);
    }

    for rel in rels.find_by_type(REL_VML_DRAWING) {
        let vml = part_text(entries, &resolve_target(SHEET_PART, &rel.target));
        let expected = if vml.contains("ClientData") {
            "legacyDrawing"
        } else {
            "legacyDrawingHF"
        };
        assert_eq!(
            named.get(rel.id.as_str()),
            Some(&expected),
            "{label}: VML drawing {} not named by <{expected}>: {sheet}",
            rel.id
        );
    }

    let order: Vec<usize> = [
        "<drawing ",
        "<legacyDrawing ",
        "<legacyDrawingHF ",
        "<tableParts",
    ]
    .iter()
    .filter_map(|tag| sheet.find(tag))
    .collect();
    assert!(
        order.is_sorted(),
        "{label}: out of CT_Worksheet order: {sheet}"
    );
}

#[test]
fn sheet_part_references_resolve_for_every_feature_mix() {
    for mix in MIXES {
        let first = write_xlsx(&mixed_sheet(mix)).unwrap();
        let second = write_xlsx(&crate::reader::read_xlsx(&first).unwrap()).unwrap();
        let third = write_xlsx(&crate::reader::read_xlsx(&second).unwrap()).unwrap();
        for (pass, bytes) in [("first", &first), ("second", &second), ("third", &third)] {
            assert_sheet_part_refs(&zip_map(bytes), &format!("{mix:?}, {pass} write"));
        }
        let wb = crate::reader::read_xlsx(&third).unwrap();
        let ws = &wb.sheets[0].worksheet;
        assert_eq!(ws.tables.len(), usize::from(mix.table), "{mix:?}");
        assert_eq!(
            ws.comments.len(),
            usize::from(mix.note) + usize::from(mix.threaded),
            "{mix:?}"
        );
        assert_eq!(ws.charts.len(), usize::from(mix.chart), "{mix:?}");
        assert_eq!(
            ws.threaded_comments.len(),
            usize::from(mix.threaded),
            "{mix:?}"
        );
    }
}

#[test]
fn removed_tables_and_pivot_tables_leave_no_relationship_behind() {
    // Read with a table beside a note, then written without the table: no
    // relationship may still name the table part the reader consumed.
    let first = write_xlsx(&mixed_sheet(mix(true, true, false))).unwrap();
    let mut wb = crate::reader::read_xlsx(&first).unwrap();
    wb.sheets[0].worksheet.tables.clear();
    let written = write_xlsx(&wb).unwrap();
    assert_sheet_part_refs(&zip_map(&written), "table removed");
    let (rels, raw) = sheet1_rels(&written);
    assert_eq!(rels.find_by_type(REL_TABLE).count(), 0, "{raw}");

    // The same for a pivot table beside a note, after two reads.
    let mut wb = mixed_sheet(mix(false, true, false));
    wb.sheets[0].worksheet.pivot_tables = from_json(serde_json::json!([
        { "name": "PT1", "location": { "ref": "D3:E5" }, "cacheId": 0 },
    ]));
    wb.pivot_caches = from_json(serde_json::json!([{
        "source": { "ref": "A1:B5", "sheet": "Sheet1" },
        "fields": [{ "name": "Name" }, { "name": "Value" }],
    }]));
    let once = write_xlsx(&crate::reader::read_xlsx(&write_xlsx(&wb).unwrap()).unwrap()).unwrap();
    let (rels, raw) = sheet1_rels(&once);
    assert_eq!(rels.find_by_type(REL_PIVOT_TABLE).count(), 1, "{raw}");
    let mut wb = crate::reader::read_xlsx(&once).unwrap();
    wb.sheets[0].worksheet.pivot_tables.clear();
    let written = write_xlsx(&wb).unwrap();
    assert_sheet_part_refs(&zip_map(&written), "pivot table removed");
    let (rels, raw) = sheet1_rels(&written);
    assert_eq!(rels.find_by_type(REL_PIVOT_TABLE).count(), 0, "{raw}");
}

// -----------------------------------------------------------------------
// Note drawings: the writer generates a VML drawing for the notes of
// a sheet that has no preserved one with cell-anchored shapes, and merges a
// preserved one whose notes no longer match the sheet's comments.
// -----------------------------------------------------------------------

/// The check box form control Microsoft Excel 16.0 saved on 2026-11-12 (its
/// only shape), on one line.
const FORM_CONTROL_VML: &str = r##"<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel"><o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="1"/></o:shapelayout><v:shapetype id="_x0000_t201" coordsize="21600,21600" o:spt="201" path="m,l,21600r21600,l21600,xe"><v:stroke joinstyle="miter"/><v:path shadowok="f" o:extrusionok="f" strokeok="f" fillok="f" o:connecttype="rect"/><o:lock v:ext="edit" shapetype="t"/></v:shapetype><v:shape id="_x0000_s1025" type="#_x0000_t201" style='position:absolute;margin-left:100pt;margin-top:30pt;width:80pt;height:18pt;z-index:1;mso-wrap-style:tight' filled="f" fillcolor="windowText [64]" stroked="f" strokecolor="window [65]" strokeweight="3e-5mm" o:insetmode="auto"><v:fill color2="window [65]"/><v:path shadowok="t" strokeok="t" fillok="t"/><o:lock v:ext="edit" rotation="t"/><v:textbox style='mso-direction-alt:auto' o:singleclick="f"><div style='text-align:left'><font face="Segoe UI" size="160" color="auto">Check Box 1</font></div></v:textbox><x:ClientData ObjectType="Checkbox"><x:SizeWithCells/><x:Anchor>2, 8, 2, 2, 3, 72, 3, 9</x:Anchor><x:AutoFill>False</x:AutoFill><x:AutoLine>False</x:AutoLine><x:TextVAlign>Center</x:TextVAlign><x:NoThreeD/></x:ClientData></v:shape></xml>"##;

/// The note shape Excel saved for C5 on 2026-11-12 (`_x0000_s1026`, after a
/// note on A1), on one line.
const C5_NOTE_SHAPE: &str = r##"<v:shape id="_x0000_s1026" type="#_x0000_t202" style='position:absolute;margin-left:151.5pt;margin-top:53pt;width:104pt;height:61.5pt;z-index:2;visibility:hidden' fillcolor="infoBackground [80]" strokecolor="none [81]" o:insetmode="auto"><v:fill color2="infoBackground [80]"/><v:shadow color="none [81]" obscured="t"/><v:path o:connecttype="none"/><v:textbox style='mso-direction-alt:auto'><div style='text-align:left'></div></v:textbox><x:ClientData ObjectType="Note"><x:MoveWithCells/><x:SizeWithCells/><x:Anchor>3, 10, 3, 12, 5, 20, 7, 17</x:Anchor><x:AutoFill>False</x:AutoFill><x:Row>4</x:Row><x:Column>2</x:Column></x:ClientData></v:shape>"##;

/// [`NOTE_VML`] (a note on D1) with Excel's note on C5 after it.
fn d1_c5_note_vml() -> String {
    NOTE_VML.replace(
        "</v:shape></xml>",
        &format!("</v:shape>{C5_NOTE_SHAPE}</xml>"),
    )
}

/// Sets sheet `sheet`'s comments to one note per cell of `cells`.
fn set_notes(wb: &mut WorkbookData, sheet: usize, cells: &[&str]) {
    wb.sheets[sheet].worksheet.comments = cells
        .iter()
        .map(|cell| comments::Comment {
            cell_ref: (*cell).to_owned(),
            author: "Ann".into(),
            text: format!("note {cell}"),
        })
        .collect();
}

/// A one-sheet workbook with a note on each of `cells` and a bar chart, as
/// the API builds it: no preserved part.
fn api_notes(cells: &[&str]) -> WorkbookData {
    let mut wb = minimal_workbook("Sheet1", Vec::new());
    set_notes(&mut wb, 0, cells);
    wb.sheets[0].worksheet.charts = vec![bar_chart(4)];
    wb
}

/// The path and text of the VML drawing sheet `sheet_num` names by
/// `<legacyDrawing>`, checking that its relationship is a vmlDrawing one.
fn legacy_vml(entries: &HashMap<String, Vec<u8>>, sheet_num: usize) -> (String, String) {
    let sheet = part_text(entries, &format!("xl/worksheets/sheet{sheet_num}.xml"));
    let open = "<legacyDrawing r:id=\"";
    let at = sheet
        .find(open)
        .unwrap_or_else(|| panic!("sheet{sheet_num} has no <legacyDrawing>: {sheet}"))
        + open.len();
    let id = &sheet[at..at + sheet[at..].find('"').unwrap()];
    let rels =
        Relationships::parse(&entries[&format!("xl/worksheets/_rels/sheet{sheet_num}.xml.rels")])
            .unwrap();
    let rel = rels.get_by_id(id).unwrap();
    assert_eq!(rel.rel_type, REL_VML_DRAWING);
    let path = resolve_target(SHEET_PART, &rel.target);
    let text = part_text(entries, &path);
    (path, text)
}

/// `(Row, Column, Anchor)` of each Note shape of a VML drawing, in order.
fn note_shapes(vml: &str) -> Vec<(u32, u32, String)> {
    vml.split(r#"<x:ClientData ObjectType="Note">"#)
        .skip(1)
        .map(|data| {
            let text = |tag: &str| {
                let open = format!("<x:{tag}>");
                let from = data.find(&open).unwrap() + open.len();
                data[from..from + data[from..].find('<').unwrap()]
                    .trim()
                    .to_owned()
            };
            (
                text("Row").parse().unwrap(),
                text("Column").parse().unwrap(),
                text("Anchor"),
            )
        })
        .collect()
}

/// Every `_x0000_sNNNN` shape id of a VML drawing, in order.
fn shape_ids(vml: &str) -> Vec<&str> {
    vml.match_indices("id=\"_x0000_s")
        .map(|(at, _)| {
            let from = at + "id=\"".len();
            &vml[from..from + vml[from..].find('"').unwrap()]
        })
        .collect()
}

#[test]
fn api_notes_get_a_generated_note_drawing() {
    let written = write_xlsx(&api_notes(&["A1", "C5"])).unwrap();
    let entries = zip_map(&written);
    assert_sheet_part_refs(&entries, "API notes + chart");
    let (path, vml) = legacy_vml(&entries, 1);
    assert_eq!(path, "xl/drawings/vmlDrawing1.vml");
    assert!(
        vml.starts_with(concat!(
            r#"<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel">"#,
            r#"<o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="1"/></o:shapelayout><v:shapetype id="_x0000_t202""#,
        )),
        "{vml}"
    );
    assert_eq!(vml.matches("visibility:hidden").count(), 2, "{vml}");
    assert_eq!(shape_ids(&vml), ["_x0000_s1025", "_x0000_s1026"]);
    let ct = ContentTypes::parse(&entries["[Content_Types].xml"]).unwrap();
    assert_eq!(ct.defaults["vml"], CT_VML_DRAWING);
    // <drawing> before <legacyDrawing> (CT_Worksheet order, checked above too).
    let sheet = part_text(&entries, "xl/worksheets/sheet1.xml");
    assert!(sheet.find("<drawing ").unwrap() < sheet.find("<legacyDrawing ").unwrap());

    // The notes read back, and a second write keeps the drawing byte for byte.
    let read = crate::reader::read_xlsx(&written).unwrap();
    assert_eq!(read.sheets[0].worksheet.comments.len(), 2);
    let again = zip_map(&write_xlsx(&read).unwrap());
    assert_sheet_part_refs(&again, "API notes, second write");
    assert_eq!(legacy_vml(&again, 1), (path, vml));
}

#[test]
fn generated_note_shapes_follow_excels_anchor_rule() {
    // Each anchor is the one Microsoft Excel 16.0 saved for a note on that
    // cell on 2026-11-12 (default column widths and row heights).
    let cases: [(&str, u32, u32, &str); 9] = [
        ("A1", 0, 0, "1, 10, 0, 1, 3, 20, 4, 6"),
        ("B2", 1, 1, "2, 10, 0, 12, 4, 20, 4, 17"),
        ("C5", 4, 2, "3, 10, 3, 12, 5, 20, 7, 17"),
        ("AB200", 199, 27, "28, 10, 198, 12, 30, 20, 202, 17"),
        ("XFA13", 12, 16380, "16381, 10, 11, 12, 16383, 20, 15, 17"),
        ("XFC1", 0, 16382, "16379, 43, 0, 1, 16381, 54, 4, 6"),
        ("XFD1", 0, 16383, "16380, 43, 0, 1, 16382, 54, 4, 6"),
        (
            "A1048576",
            1_048_575,
            0,
            "1, 10, 1048571, 0, 3, 20, 1048575, 5",
        ),
        (
            "XFD1048576",
            1_048_575,
            16383,
            "16380, 43, 1048571, 0, 16382, 54, 1048575, 5",
        ),
    ];
    let cells: Vec<&str> = cases.iter().map(|case| case.0).collect();
    let entries = zip_map(&write_xlsx(&api_notes(&cells)).unwrap());
    let (_, vml) = legacy_vml(&entries, 1);
    let shapes = note_shapes(&vml);
    assert_eq!(shapes.len(), cases.len(), "{vml}");
    for ((cell, row, col, anchor), shape) in cases.iter().zip(&shapes) {
        assert_eq!(shape, &(*row, *col, (*anchor).to_owned()), "{cell}");
    }
    let ids = shape_ids(&vml);
    assert_eq!(ids.len(), cases.len());
    assert_eq!(
        ids.iter().collect::<HashSet<_>>().len(),
        ids.len(),
        "{ids:?}"
    );
}

#[test]
fn two_sheets_with_notes_get_two_note_drawings() {
    let mut wb = api_notes(&["A1"]);
    wb.sheets.push(SheetData {
        name: "Sheet2".into(),
        ..Default::default()
    });
    set_notes(&mut wb, 1, &["B2", "B2"]);
    let entries = zip_map(&write_xlsx(&wb).unwrap());
    let (first, vml1) = legacy_vml(&entries, 1);
    let (second, vml2) = legacy_vml(&entries, 2);
    assert_eq!(
        (first.as_str(), second.as_str()),
        ("xl/drawings/vmlDrawing1.vml", "xl/drawings/vmlDrawing2.vml")
    );
    // Excel numbers each sheet's shapes from its own idmap block of 1024; a
    // cell with two comments gets one note shape.
    assert!(
        vml2.contains(r#"<o:idmap v:ext="edit" data="2"/>"#),
        "{vml2}"
    );
    assert_eq!(shape_ids(&vml1), ["_x0000_s1025"]);
    assert_eq!(shape_ids(&vml2), ["_x0000_s2049"]);
}

/// Sheet `sheet`'s comments: one note without author or text on each of
/// `count` cells, A1 down column A.
fn set_bare_notes(wb: &mut WorkbookData, sheet: usize, count: usize) {
    wb.sheets[sheet].worksheet.comments = (1..=count)
        .map(|row| comments::Comment {
            cell_ref: format!("A{row}"),
            author: String::new(),
            text: String::new(),
        })
        .collect();
}

/// A thread root on `cell` with id `{0000000N-...}`, without text.
fn bare_thread_root(n: usize, cell: &str) -> threaded_comments::ThreadedCommentData {
    threaded_comments::ThreadedCommentData {
        id: format!("{{{n:08X}-0000-0000-0000-000000000000}}"),
        ref_cell: cell.to_owned(),
        person_id: "{00000000-0000-0000-0000-000000000001}".into(),
        text: String::new(),
        timestamp: String::new(),
        parent_id: None,
        done: None,
        mentions: Vec::new(),
    }
}

/// Asserts `result` is the `SizeLimit` of the generated-note ceiling with
/// `total` reached.
fn assert_note_ceiling(result: Result<Vec<u8>>, total: usize, label: &str) {
    let e = result.expect_err(label);
    assert_eq!(e.code(), "SIZE_LIMIT", "{label}: {e}");
    let msg = e.to_string();
    assert!(
        msg.contains(&format!("reach {total} in one workbook, more than 131072")),
        "{label}: {msg}"
    );
}

#[test]
fn exactly_the_ceiling_of_generated_notes_writes_and_one_more_fails() {
    use crate::ooxml::vml::MAX_GENERATED_NOTE_SHAPES;
    assert_eq!(MAX_GENERATED_NOTE_SHAPES, 131_072);
    // Legacy notes on A1..A131071 and one thread root on B1: its placeholder
    // and the note shape it gets count once, so the write generates exactly
    // the ceiling.
    let mut wb = minimal_workbook("Sheet1", Vec::new());
    set_bare_notes(&mut wb, 0, MAX_GENERATED_NOTE_SHAPES - 1);
    wb.sheets[0].worksheet.threaded_comments = vec![bare_thread_root(1, "B1")];
    let entries = zip_map(&write_xlsx(&wb).expect("exactly the ceiling writes"));
    let (_, vml) = legacy_vml(&entries, 1);
    assert_eq!(
        vml.matches("<v:shape ").count(),
        MAX_GENERATED_NOTE_SHAPES,
        "one note shape per generated note"
    );
    let comments_xml = part_text(&entries, "xl/comments1.xml");
    assert_eq!(
        comments_xml.matches("<comment ").count(),
        MAX_GENERATED_NOTE_SHAPES
    );
    drop((entries, vml, comments_xml));

    // One more thread root, or one more legacy note, passes it.
    let mut one_more_root = wb.clone();
    one_more_root.sheets[0]
        .worksheet
        .threaded_comments
        .push(bare_thread_root(2, "B2"));
    assert_note_ceiling(
        write_xlsx(&one_more_root),
        MAX_GENERATED_NOTE_SHAPES + 1,
        "one more thread root",
    );
    let mut one_more_note = wb;
    one_more_note.sheets[0]
        .worksheet
        .comments
        .push(comments::Comment {
            cell_ref: "C1".into(),
            author: String::new(),
            text: String::new(),
        });
    assert_note_ceiling(
        write_xlsx(&one_more_note),
        MAX_GENERATED_NOTE_SHAPES + 1,
        "one more legacy note",
    );
}

#[test]
fn a_preserved_note_drawing_matching_its_notes_writes_past_the_ceiling() {
    use crate::ooxml::vml::{MAX_GENERATED_NOTE_SHAPES, note_cells, note_drawing};
    // Past the ceiling, with the drawing note_drawing makes for those notes
    // (what a write of them saves) as the sheet's preserved note drawing:
    // the writer generates nothing, so nothing is counted.
    let mut wb = minimal_workbook("Sheet1", Vec::new());
    set_bare_notes(&mut wb, 0, MAX_GENERATED_NOTE_SHAPES + 1);
    let cells = note_cells(&wb.sheets[0].worksheet.comments, 0, &mut 0).unwrap();
    let (vml, _) = note_drawing(&cells, 1);
    let path = "xl/drawings/vmlDrawing1.vml";
    let mut rels = Relationships::new();
    rels.add("rId1", REL_VML_DRAWING, "../drawings/vmlDrawing1.vml");
    wb.preserved_entries.insert(
        "xl/worksheets/_rels/sheet1.xml.rels".into(),
        rels.to_xml().unwrap(),
    );
    wb.preserved_entries.insert(path.into(), vml);
    let entries = zip_map(&write_xlsx(&wb).expect("a matching drawing counts nothing"));
    let (written, _) = legacy_vml(&entries, 1);
    assert_eq!(written, path);
    assert!(
        entries[path] == wb.preserved_entries[path],
        "the drawing is written as it is"
    );
}

#[test]
fn generated_notes_are_counted_per_write_across_sheets() {
    use crate::ooxml::vml::MAX_GENERATED_NOTE_SHAPES;
    // Each sheet alone is within the ceiling; together they pass it, and the
    // write fails at the second sheet, before its drawing is built.
    let mut wb = minimal_workbook("Sheet1", Vec::new());
    wb.sheets.push(SheetData {
        name: "Sheet2".into(),
        ..Default::default()
    });
    set_bare_notes(&mut wb, 0, 1);
    set_bare_notes(&mut wb, 1, MAX_GENERATED_NOTE_SHAPES);
    assert_note_ceiling(
        write_xlsx(&wb),
        MAX_GENERATED_NOTE_SHAPES + 1,
        "legacy notes on two sheets",
    );
    // The same with thread roots on the second sheet.
    wb.sheets[1].worksheet.comments.clear();
    wb.sheets[1].worksheet.threaded_comments = (1..=MAX_GENERATED_NOTE_SHAPES)
        .map(|n| bare_thread_root(n, &format!("A{n}")))
        .collect();
    assert_note_ceiling(
        write_xlsx(&wb),
        MAX_GENERATED_NOTE_SHAPES + 1,
        "a legacy note and thread roots on two sheets",
    );
}

#[test]
fn a_preserved_note_drawing_matching_the_notes_is_kept_byte_for_byte() {
    for (vml, cells) in [
        (NOTE_VML.to_owned(), &["D1"][..]),
        (d1_c5_note_vml(), &["C5", "D1"][..]),
    ] {
        let mut wb = mixed_sheet(mix(false, true, true));
        wb.preserved_entries.insert(
            "xl/drawings/vmlDrawing1.vml".into(),
            vml.clone().into_bytes(),
        );
        set_notes(&mut wb, 0, cells);
        let entries = zip_map(&write_xlsx(&wb).unwrap());
        assert_sheet_part_refs(&entries, "matching preserved notes");
        assert_eq!(
            legacy_vml(&entries, 1),
            ("xl/drawings/vmlDrawing1.vml".to_owned(), vml)
        );
        assert!(!entries.contains_key("xl/drawings/vmlDrawing2.vml"));
    }
}

/// Writes `wb` with sheet 1's notes set to `cells` and returns its note
/// drawing, after checking the package's sheet references and that the
/// notes read back.
fn noted_drawing(mut wb: WorkbookData, cells: &[&str], label: &str) -> (String, String) {
    set_notes(&mut wb, 0, cells);
    let written = write_xlsx(&wb).unwrap();
    let entries = zip_map(&written);
    assert_sheet_part_refs(&entries, label);
    let read = crate::reader::read_xlsx(&written).unwrap();
    assert_eq!(
        read.sheets[0].worksheet.comments.len(),
        cells.len(),
        "{label}"
    );
    legacy_vml(&entries, 1)
}

/// The `<v:shape>` element of `vml` whose id is `id`.
fn shape_by_id<'a>(vml: &'a str, id: &str) -> &'a str {
    let from = vml.find(&format!(r#"<v:shape id="{id}""#)).unwrap();
    &vml[from..from + vml[from..].find("</v:shape>").unwrap() + "</v:shape>".len()]
}

#[test]
fn a_preserved_note_drawing_is_merged_with_the_notes_it_lacks_or_no_longer_has() {
    let d1_shape = shape_by_id(NOTE_VML, "_x0000_s1025");
    let excel_two = d1_c5_note_vml();
    let with = |vml: &str| {
        let mut wb = mixed_sheet(mix(false, true, true));
        wb.preserved_entries.insert(
            "xl/drawings/vmlDrawing1.vml".into(),
            vml.as_bytes().to_vec(),
        );
        wb
    };

    // Added: the D1 shape is kept as it is and F10 gets a shape numbered
    // above the highest id in the file.
    let (path, vml) = noted_drawing(with(NOTE_VML), &["D1", "F10"], "note added");
    assert_eq!(path, "xl/drawings/vmlDrawing1.vml");
    assert!(
        vml.starts_with(&NOTE_VML[..NOTE_VML.find("<v:shape ").unwrap()]),
        "{vml}"
    );
    assert!(vml.contains(d1_shape), "{vml}");
    assert_eq!(shape_ids(&vml), ["_x0000_s1025", "_x0000_s1026"]);
    assert_eq!(
        note_shapes(&vml)[1],
        (9, 5, "6, 10, 8, 12, 8, 20, 12, 17".to_owned())
    );

    // Removed: C5's shape goes and D1's stays byte for byte.
    let (_, vml) = noted_drawing(with(&excel_two), &["D1"], "note removed");
    assert_eq!(vml, NOTE_VML);

    // Moved (rows inserted above C5): the stale shape goes and C8 gets a new
    // one above the highest id in the file.
    let (_, vml) = noted_drawing(with(&excel_two), &["D1", "C8"], "note moved");
    assert!(vml.contains(d1_shape), "{vml}");
    assert!(!vml.contains(C5_NOTE_SHAPE), "{vml}");
    assert_eq!(shape_ids(&vml), ["_x0000_s1025", "_x0000_s1027"]);
    assert_eq!(note_shapes(&vml)[1].0, 7);

    // A drawing holding only a form control gains the note shape type and
    // the note; the check box is kept as it is, and no second part appears.
    let (path, vml) = noted_drawing(with(FORM_CONTROL_VML), &["A1"], "form control + note");
    assert_eq!(path, "xl/drawings/vmlDrawing1.vml");
    assert!(
        vml.contains(shape_by_id(FORM_CONTROL_VML, "_x0000_s1025")),
        "{vml}"
    );
    assert!(vml.contains(r#"<v:shapetype id="_x0000_t202""#), "{vml}");
    assert_eq!(shape_ids(&vml), ["_x0000_s1025", "_x0000_s1026"]);
    assert_eq!(
        note_shapes(&vml),
        [(0, 0, "1, 10, 0, 1, 3, 20, 4, 6".to_owned())]
    );
}

#[test]
fn an_unparseable_preserved_note_drawing_is_kept_and_named() {
    let broken = r#"<xml xmlns:x="urn:schemas-microsoft-com:office:excel"><v:shape><x:ClientData ObjectType="Note"><x:Row>0<br></x:Row></x:ClientData></v:shape></xml>"#;
    let mut wb = mixed_sheet(mix(false, true, true));
    wb.preserved_entries.insert(
        "xl/drawings/vmlDrawing1.vml".into(),
        broken.as_bytes().to_vec(),
    );
    let (path, vml) = noted_drawing(wb, &["A1", "B7"], "unparseable VML");
    assert_eq!(
        (path.as_str(), vml.as_str()),
        ("xl/drawings/vmlDrawing1.vml", broken)
    );
}

#[test]
fn a_note_drawing_with_a_second_root_is_kept_as_it_is() {
    // A crafted drawing whose root is followed by a second one: D1's note
    // would be dropped and F10's added, but the drawing is written as it is.
    let crafted = format!("{NOTE_VML}<xml></xml>");
    let mut wb = mixed_sheet(mix(false, true, true));
    wb.preserved_entries.insert(
        "xl/drawings/vmlDrawing1.vml".into(),
        crafted.clone().into_bytes(),
    );
    let (path, vml) = noted_drawing(wb, &["F10"], "second root");
    assert_eq!(
        (path.as_str(), vml.as_str()),
        ("xl/drawings/vmlDrawing1.vml", crafted.as_str())
    );
}

#[test]
fn a_header_picture_drawing_is_no_note_drawing() {
    // Header picture VML only: the notes get a drawing of their own, named
    // by <legacyDrawing>, beside the <legacyDrawingHF> one.
    let mut wb = mixed_sheet(Mix {
        header_picture: true,
        ..mix(false, false, true)
    });
    set_notes(&mut wb, 0, &["A1"]);
    let entries = zip_map(&write_xlsx(&wb).unwrap());
    assert_sheet_part_refs(&entries, "header picture + API note");
    let (path, vml) = legacy_vml(&entries, 1);
    assert_eq!(path, "xl/drawings/vmlDrawing1.vml");
    assert_eq!(note_shapes(&vml).len(), 1);
    assert_eq!(
        part_text(&entries, "xl/drawings/vmlDrawing2.vml"),
        HEADER_PICTURE_VML
    );
}

#[test]
fn a_merged_note_drawing_named_xml_keeps_the_vml_drawing_type() {
    // The note drawing at an .xml name with a vmlDrawing Override, as the
    // reader records it: merged with an added note, it keeps that type
    // instead of falling to the xml Default.
    let path = "xl/drawings/vmlDrawing1.xml";
    let mut wb = mixed_sheet(mix(false, true, true));
    let vml = wb
        .preserved_entries
        .remove("xl/drawings/vmlDrawing1.vml")
        .unwrap();
    wb.preserved_entries.insert(path.into(), vml);
    wb.preserved_entries.insert(
        SHEET1_RELS.into(),
        rels_part(&[("rId2", REL_VML_DRAWING, "../drawings/vmlDrawing1.xml")]),
    );
    wb.preserved_content_types
        .insert(path.into(), CT_VML_DRAWING.into());
    set_notes(&mut wb, 0, &["D1", "F10"]);
    let entries = zip_map(&write_xlsx(&wb).unwrap());
    assert_sheet_part_refs(&entries, "note added to an .xml-named drawing");
    let (named, vml) = legacy_vml(&entries, 1);
    assert_eq!(named, path);
    assert_eq!(shape_ids(&vml), ["_x0000_s1025", "_x0000_s1026"]);
    let ct = ContentTypes::parse(&entries["[Content_Types].xml"]).unwrap();
    assert_eq!(ct.overrides[&format!("/{path}")], CT_VML_DRAWING);
}

#[test]
fn a_vml_drawing_shared_by_two_sheets_stays_named_by_both_after_a_merge() {
    // Two sheets name one form-control drawing; a note added to sheet 1 is
    // merged into it, and sheet 2 still names it by <legacyDrawing>.
    let shared = "xl/drawings/vmlDrawing1.vml";
    let rels = rels_part(&[("rId1", REL_VML_DRAWING, "../drawings/vmlDrawing1.vml")]);
    let mut wb = minimal_workbook("Sheet1", Vec::new());
    wb.sheets.push(SheetData {
        name: "Sheet2".into(),
        ..Default::default()
    });
    wb.preserved_entries
        .insert(shared.into(), FORM_CONTROL_VML.as_bytes().to_vec());
    wb.preserved_entries
        .insert(SHEET1_RELS.into(), rels.clone());
    wb.preserved_entries
        .insert("xl/worksheets/_rels/sheet2.xml.rels".into(), rels);
    set_notes(&mut wb, 0, &["F10"]);
    let entries = zip_map(&write_xlsx(&wb).unwrap());
    assert_sheet_part_refs(&entries, "note added to a shared drawing");
    let (first, vml) = legacy_vml(&entries, 1);
    assert_eq!(first, shared);
    assert_eq!(note_shapes(&vml).len(), 1, "{vml}");
    assert_eq!(legacy_vml(&entries, 2), (first, vml.clone()));

    // With notes on sheet 2 as well, the drawing is merged once (a second
    // entry at its path would fail the write) and both sheets name it.
    set_notes(&mut wb, 1, &["B2"]);
    let entries = zip_map(&write_xlsx(&wb).unwrap());
    assert_eq!(legacy_vml(&entries, 1), (shared.to_owned(), vml.clone()));
    assert_eq!(legacy_vml(&entries, 2), (shared.to_owned(), vml));
}

// ---------------------------------------------------------------------------
// vmlDrawing relationships whose target is no VML drawing
// ---------------------------------------------------------------------------

/// A custom XML data part that holds the bytes `ClientData`, like the
/// security review's sample H0.
const CLIENT_DATA_XML: &str = "<root>ClientData</root>";

/// The second sheet's preserved .rels: one relationship whose Id is
/// `ClientData`, like the security review's sample G0.
const SHEET2_RELS: &str = "xl/worksheets/_rels/sheet2.xml.rels";

/// Asserts that sheet 1 names one generated note drawing with a note on A1
/// by `<legacyDrawing>` and nothing by `<legacyDrawingHF>`.
fn assert_generated_note_drawing(entries: &HashMap<String, Vec<u8>>, label: &str) {
    let (path, vml) = legacy_vml(entries, 1);
    assert_eq!(path, "xl/drawings/vmlDrawing1.vml", "{label}");
    assert_eq!(
        note_shapes(&vml),
        [(0, 0, "1, 10, 0, 1, 3, 20, 4, 6".to_owned())],
        "{label}"
    );
    let sheet = part_text(entries, "xl/worksheets/sheet1.xml");
    assert!(!sheet.contains("<legacyDrawingHF"), "{label}: {sheet}");
}

#[test]
fn a_vml_drawing_relationship_to_a_part_that_is_no_vml_drawing_is_not_merged() {
    // A vmlDrawing relationship of sheet 1 targets (i) the sheet's own .rels,
    // (ii) sheet 2's .rels, (iii) a custom XML part outside xl/, each holding
    // the bytes ClientData, and sheet 1 has a note on A1. The part is neither
    // merged into nor named: sheet 1 gets a generated note drawing,
    // xl/drawings/vmlDrawing1.vml, which <legacyDrawing> names, as if it had
    // no preserved one. Sheet 2's .rels and the custom XML part keep their
    // bytes; sheet 1's own .rels, which the writer rebuilds for the note,
    // keeps the relationship as it was.
    let custom = "customXml/item1.xml";
    let cases = [
        ("own .rels", "_rels/sheet1.xml.rels"),
        ("another sheet's .rels", "_rels/sheet2.xml.rels"),
        ("a part outside xl/", "../../customXml/item1.xml"),
    ];
    for (label, target) in cases {
        let mut wb = minimal_workbook("Sheet1", Vec::new());
        wb.sheets.push(SheetData {
            name: "Sheet2".into(),
            ..Default::default()
        });
        let sheet2_rels = rels_part(&[("ClientData", "http://example.com/x", "../x.bin")]);
        wb.preserved_entries.insert(
            SHEET1_RELS.into(),
            rels_part(&[("ClientData", REL_VML_DRAWING, target)]),
        );
        wb.preserved_entries
            .insert(SHEET2_RELS.into(), sheet2_rels.clone());
        wb.preserved_entries
            .insert(custom.into(), CLIENT_DATA_XML.as_bytes().to_vec());
        set_notes(&mut wb, 0, &["A1"]);
        let written = write_xlsx(&wb).unwrap_or_else(|e| panic!("{label}: {e}"));
        let entries = zip_map(&written);
        assert_generated_note_drawing(&entries, label);
        assert!(
            entries[SHEET2_RELS] == sheet2_rels,
            "{label}: sheet 2 .rels changed"
        );
        assert_eq!(part_text(&entries, custom), CLIENT_DATA_XML, "{label}");
        let rels = Relationships::parse(&entries[SHEET1_RELS]).unwrap();
        let kept = rels.get_by_id("ClientData").unwrap();
        assert_eq!(
            (&*kept.rel_type, &*kept.target),
            (REL_VML_DRAWING, target),
            "{label}"
        );
    }
}

#[test]
fn hostile_vml_drawing_targets_read_and_written_back_stay_untouched() {
    // The security review's samples F0, G0, H0 and Habs0, built here: a
    // vmlDrawing relationship to the sheet's own .rels (Id ClientData), to
    // sheet 2's .rels (which holds Id ClientData), and to a custom XML part
    // holding ClientData by a relative target that climbs past the package
    // root and by an absolute one. Read and written back, with no edit and
    // with a note added on A1, the targeted part keeps its bytes and no
    // <legacyDrawing> or <legacyDrawingHF> names it.
    let custom = "customXml/item1.xml";
    let cases = [
        ("F0", "ClientData", "_rels/sheet1.xml.rels", SHEET1_RELS),
        ("G0", "rId1", "_rels/sheet2.xml.rels", SHEET2_RELS),
        ("H0", "rId1", "../../../../customXml/item1.xml", custom),
        ("Habs0", "rId1", "/customXml/item1.xml", custom),
    ];
    for (label, id, target, part) in cases {
        let mut wb = minimal_workbook("Sheet1", Vec::new());
        wb.sheets.push(SheetData {
            name: "Sheet2".into(),
            ..Default::default()
        });
        let mut entries = zip_map(&write_xlsx(&wb).unwrap());
        entries.insert(
            SHEET1_RELS.into(),
            rels_part(&[(id, REL_VML_DRAWING, target)]),
        );
        entries.insert(
            SHEET2_RELS.into(),
            rels_part(&[("ClientData", "http://example.com/x", "../x.bin")]),
        );
        entries.insert(custom.into(), CLIENT_DATA_XML.as_bytes().to_vec());
        let package = rezip(entries.clone());

        let mut read = crate::reader::read_xlsx(&package).unwrap();
        let again = zip_map(&write_xlsx(&read).unwrap_or_else(|e| panic!("{label}: {e}")));
        let sheet = part_text(&again, "xl/worksheets/sheet1.xml");
        assert!(!sheet.contains("<legacyDrawing"), "{label}: {sheet}");
        assert!(again[part] == entries[part], "{label}: {part} changed");

        set_notes(&mut read, 0, &["A1"]);
        let noted = zip_map(&write_xlsx(&read).unwrap_or_else(|e| panic!("{label} + note: {e}")));
        assert_generated_note_drawing(&noted, label);
        if part != SHEET1_RELS {
            assert!(
                noted[part] == entries[part],
                "{label} + note: {part} changed"
            );
        }
    }
}

// ---------------------------------------------------------------------------
// Threaded comments: the placeholder note Excel needs beside each thread
// ---------------------------------------------------------------------------

/// The notice Microsoft Excel 16.0 wrote into each placeholder note on
/// 2026-11-12, between `[Threaded comment]` and `Comment:`.
const EXCEL_THREAD_NOTICE: &str = "Your version of Excel allows you to read this threaded comment; however, any edits to it will get removed if the file is opened in a newer version of Excel. Learn more: https://go.microsoft.com/fwlink/?linkid=870924";

/// The placeholder text Excel wrote for a thread: the notice, the root's
/// text after `Comment:` and each reply's after `Reply:`, every heading on
/// its own line and every text indented by four spaces (as read: Excel's
/// CRLF line ends read as LF).
fn excel_placeholder_text(root: &str, replies: &[&str]) -> String {
    let mut text = format!("[Threaded comment]\n\n{EXCEL_THREAD_NOTICE}\n\nComment:\n    {root}");
    for reply in replies {
        text.push_str("\nReply:\n    ");
        text.push_str(reply);
    }
    text
}

/// A threaded comment by person `{…AA}` (a reply when `parent` is given).
fn thread_comment(
    id: &str,
    cell: &str,
    text: &str,
    parent: Option<&str>,
) -> threaded_comments::ThreadedCommentData {
    threaded_comments::ThreadedCommentData {
        id: id.into(),
        ref_cell: cell.into(),
        person_id: "{00000000-0000-0000-0000-0000000000AA}".into(),
        text: text.into(),
        timestamp: "2026-11-12T00:00:00.00".into(),
        parent_id: parent.map(Into::into),
        done: None,
        mentions: Vec::new(),
    }
}

/// A legacy comment.
fn note(cell: &str, author: &str, text: &str) -> comments::Comment {
    comments::Comment {
        cell_ref: cell.into(),
        author: author.into(),
        text: text.into(),
    }
}

/// `(cell, author, text)` of each comment of sheet `sheet_num`'s comments
/// part, in order, checking that the sheet .rels names the part.
fn sheet_comments(
    entries: &HashMap<String, Vec<u8>>,
    sheet_num: usize,
) -> Vec<(String, String, String)> {
    let rels =
        Relationships::parse(&entries[&format!("xl/worksheets/_rels/sheet{sheet_num}.xml.rels")])
            .unwrap();
    let rel = rels.find_by_type(REL_COMMENTS).next().unwrap();
    let path = resolve_target(SHEET_PART, &rel.target);
    let ct = part_text(entries, "[Content_Types].xml");
    assert!(ct.contains(&format!(r#"PartName="/{path}""#)), "{ct}");
    comments::parse_comments(&entries[&path])
        .unwrap()
        .into_iter()
        .map(|c| (c.cell_ref, c.author, c.text))
        .collect()
}

const ROOT_ID: &str = "{5FA6E5A7-6BD7-4105-A288-53EB03A03DCF}";
const REPLY1_ID: &str = "{87D66092-362F-4C38-9C8D-1611ABA57577}";
const REPLY2_ID: &str = "{87D66092-362F-4C39-9C8D-1611ABA57577}";

/// One sheet with `threads` and `notes`, and the thread author in `persons`.
fn threaded_workbook(
    threads: Vec<threaded_comments::ThreadedCommentData>,
    notes: Vec<comments::Comment>,
) -> WorkbookData {
    let mut wb = minimal_workbook("Sheet1", Vec::new());
    wb.sheets[0].worksheet.threaded_comments = threads;
    wb.sheets[0].worksheet.comments = notes;
    wb.persons = from_json(serde_json::json!([
        { "id": "{00000000-0000-0000-0000-0000000000AA}", "displayName": "Ann" },
    ]));
    wb
}

#[test]
fn each_thread_gets_the_placeholder_note_excel_needs() {
    // Excel shows a thread only when its cell has a legacy comment, the
    // placeholder older versions display, whose author is `tc={root id}`
    // (Excel 16 hides the thread without it). The writer adds one
    // for each thread root whose cell has no comment; a cell that has one
    // keeps it as it is (Excel then shows the note, not the thread).
    type Case = (
        &'static str,
        Vec<threaded_comments::ThreadedCommentData>,
        Vec<comments::Comment>,
        Vec<(&'static str, String, String)>,
    );
    let cases: Vec<Case> = vec![
        (
            "one thread, no note",
            vec![thread_comment(ROOT_ID, "B2", "Please check", None)],
            Vec::new(),
            vec![(
                "B2",
                format!("tc={ROOT_ID}"),
                excel_placeholder_text("Please check", &[]),
            )],
        ),
        (
            "a thread with two replies",
            vec![
                thread_comment(ROOT_ID, "D4", "Root on D4", None),
                thread_comment(REPLY1_ID, "D4", "First reply", Some(ROOT_ID)),
                thread_comment(REPLY2_ID, "D4", "Second reply", Some(ROOT_ID)),
            ],
            Vec::new(),
            vec![(
                "D4",
                format!("tc={ROOT_ID}"),
                excel_placeholder_text("Root on D4", &["First reply", "Second reply"]),
            )],
        ),
        (
            "multi-line texts",
            vec![
                thread_comment(ROOT_ID, "F6", "Line one\nLine two", None),
                thread_comment(REPLY1_ID, "F6", "Reply one\nReply two", Some(ROOT_ID)),
            ],
            Vec::new(),
            vec![(
                "F6",
                format!("tc={ROOT_ID}"),
                excel_placeholder_text("Line one\nLine two", &["Reply one\nReply two"]),
            )],
        ),
        (
            "a note already on the thread's cell",
            vec![thread_comment(ROOT_ID, "B2", "Please check", None)],
            vec![note("B2", "Ann", "user note")],
            vec![("B2", "Ann".into(), "user note".into())],
        ),
        (
            "a note on another cell",
            vec![thread_comment(ROOT_ID, "B2", "Please check", None)],
            vec![note("A1", "Ann", "user note")],
            vec![
                ("A1", "Ann".into(), "user note".into()),
                (
                    "B2",
                    format!("tc={ROOT_ID}"),
                    excel_placeholder_text("Please check", &[]),
                ),
            ],
        ),
    ];
    for (label, threads, notes, expected) in cases {
        let wb = threaded_workbook(threads, notes);
        let before = serde_json::to_value(&wb).unwrap();
        let written = write_xlsx(&wb).unwrap_or_else(|e| panic!("{label}: {e}"));
        assert_eq!(
            serde_json::to_value(&wb).unwrap(),
            before,
            "{label}: the model changed"
        );
        let entries = zip_map(&written);
        assert_sheet_part_refs(&entries, label);
        let expected: Vec<(String, String, String)> = expected
            .into_iter()
            .map(|(cell, author, text)| (cell.to_owned(), author, text))
            .collect();
        assert_eq!(sheet_comments(&entries, 1), expected, "{label}");

        // One hidden note shape per commented cell, named by <legacyDrawing>.
        let (_, vml) = legacy_vml(&entries, 1);
        assert_eq!(note_shapes(&vml).len(), expected.len(), "{label}: {vml}");
        assert_eq!(
            vml.matches("visibility:hidden").count(),
            expected.len(),
            "{label}"
        );

        // Read back, the placeholder is an ordinary comment, and a second
        // write adds nothing and keeps the comments and the drawing.
        let read = crate::reader::read_xlsx(&written).unwrap();
        assert_eq!(
            read.sheets[0].worksheet.comments.len(),
            expected.len(),
            "{label}"
        );
        let again = zip_map(&write_xlsx(&read).unwrap());
        assert_eq!(sheet_comments(&again, 1), expected, "{label}: second write");
        assert_eq!(legacy_vml(&again, 1).1, vml, "{label}: second write");
    }
}

#[test]
fn only_thread_roots_get_a_placeholder_and_one_per_cell() {
    // A reply whose parent is not on the sheet and a second root on the
    // same cell get none; replies follow their root whatever the order.
    let other = "{6E58D533-3A4A-48B4-87EF-54644A101AD4}";
    let wb = threaded_workbook(
        vec![
            thread_comment(REPLY1_ID, "D4", "First reply", Some(ROOT_ID)),
            thread_comment(ROOT_ID, "D4", "Root on D4", None),
            thread_comment(other, "D4", "Second root", None),
            thread_comment(
                REPLY2_ID,
                "C3",
                "Orphan",
                Some("{00000000-0000-0000-0000-000000000009}"),
            ),
        ],
        Vec::new(),
    );
    let entries = zip_map(&write_xlsx(&wb).unwrap());
    assert_eq!(
        sheet_comments(&entries, 1),
        [(
            "D4".to_owned(),
            format!("tc={ROOT_ID}"),
            excel_placeholder_text("Root on D4", &["First reply"]),
        )]
    );
}

#[test]
fn two_sheets_with_threads_get_their_own_placeholders_and_note_drawings() {
    let second = "{7FA93661-956B-479C-BBC4-110295297E1B}";
    let mut wb = threaded_workbook(
        vec![thread_comment(ROOT_ID, "B2", "Please check", None)],
        Vec::new(),
    );
    wb.sheets.push(SheetData {
        name: "Sheet2".into(),
        worksheet: WorksheetXml {
            threaded_comments: vec![thread_comment(second, "C3", "Sheet2 thread", None)],
            comments: vec![note("E5", "Ann", "An ordinary note")],
            ..Default::default()
        },
        ..Default::default()
    });
    let before = serde_json::to_value(&wb).unwrap();
    let entries = zip_map(&write_xlsx(&wb).unwrap());
    assert_eq!(serde_json::to_value(&wb).unwrap(), before);
    assert_eq!(
        sheet_comments(&entries, 1),
        [(
            "B2".to_owned(),
            format!("tc={ROOT_ID}"),
            excel_placeholder_text("Please check", &[]),
        )]
    );
    assert_eq!(
        sheet_comments(&entries, 2),
        [
            (
                "E5".to_owned(),
                "Ann".to_owned(),
                "An ordinary note".to_owned()
            ),
            (
                "C3".to_owned(),
                format!("tc={second}"),
                excel_placeholder_text("Sheet2 thread", &[]),
            ),
        ]
    );
    let (first_vml, vml1) = legacy_vml(&entries, 1);
    let (second_vml, vml2) = legacy_vml(&entries, 2);
    assert_ne!(first_vml, second_vml);
    assert_eq!(
        note_shapes(&vml1),
        [(1, 1, "2, 10, 0, 12, 4, 20, 4, 17".to_owned())]
    );
    assert_eq!(note_shapes(&vml2).len(), 2, "{vml2}");
}

/// Sheet 1 of a workbook Microsoft Excel 16.0 saved on 2026-11-12 with
/// threads on B2, D4 (two replies) and F6 (multi-line, one reply): its
/// threaded comments part as Excel wrote it.
const EXCEL_THREADS_XML: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<ThreadedComments xmlns="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments" xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><threadedComment ref="B2" dT="2026-11-12T15:33:14.84" personId="{C71F189D-7DBF-40FC-91D4-FE2936EFAE3F}" id="{5FA6E5A7-6BD7-4105-A288-53EB03A03DCF}"><text>Please check</text></threadedComment><threadedComment ref="D4" dT="2026-11-12T15:33:14.85" personId="{C71F189D-7DBF-40FC-91D4-FE2936EFAE3F}" id="{CDE37F6B-39FC-41E4-A6C1-0816C3C6187E}"><text>Root on D4</text></threadedComment><threadedComment ref="D4" dT="2026-11-12T15:33:14.85" personId="{C71F189D-7DBF-40FC-91D4-FE2936EFAE3F}" id="{87D66092-362F-4C38-9C8D-1611ABA57577}" parentId="{CDE37F6B-39FC-41E4-A6C1-0816C3C6187E}"><text>First reply</text></threadedComment><threadedComment ref="D4" dT="2026-11-12T15:33:14.85" personId="{C71F189D-7DBF-40FC-91D4-FE2936EFAE3F}" id="{87D66092-362F-4C39-9C8D-1611ABA57577}" parentId="{CDE37F6B-39FC-41E4-A6C1-0816C3C6187E}"><text>Second reply</text></threadedComment><threadedComment ref="F6" dT="2026-11-12T15:33:14.85" personId="{C71F189D-7DBF-40FC-91D4-FE2936EFAE3F}" id="{BFCEF060-A399-48BE-BFBB-B59039308924}"><text>Line one
Line two</text></threadedComment><threadedComment ref="F6" dT="2026-11-12T15:33:14.85" personId="{C71F189D-7DBF-40FC-91D4-FE2936EFAE3F}" id="{6E58D533-3A4A-48B4-87EF-54644A101AD4}" parentId="{BFCEF060-A399-48BE-BFBB-B59039308924}"><text>Reply one
Reply two</text></threadedComment></ThreadedComments>"#;

/// The persons part of the same file, with the signed-in account's display
/// name and user id replaced by neutral values.
const EXCEL_PERSONS_XML: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<personList xmlns="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments" xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><person displayName="Ann" id="{C71F189D-7DBF-40FC-91D4-FE2936EFAE3F}" userId="0123456789abcdef" providerId="Windows Live"/></personList>"#;

/// The note drawing of the same file, with LF line ends (Excel wrote CRLF;
/// [`excel_threads_package`] puts them back).
const EXCEL_THREADS_VML: &str = r##"<xml xmlns:v="urn:schemas-microsoft-com:vml"
 xmlns:o="urn:schemas-microsoft-com:office:office"
 xmlns:x="urn:schemas-microsoft-com:office:excel">
 <o:shapelayout v:ext="edit">
  <o:idmap v:ext="edit" data="1"/>
 </o:shapelayout><v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202"
  path="m,l,21600r21600,l21600,xe">
  <v:stroke joinstyle="miter"/>
  <v:path gradientshapeok="t" o:connecttype="rect"/>
 </v:shapetype><v:shape id="_x0000_s1025" type="#_x0000_t202" style='position:absolute;
  margin-left:103.5pt;margin-top:9.5pt;width:104pt;height:61.5pt;z-index:1;
  visibility:hidden' fillcolor="infoBackground [80]" strokecolor="none [81]"
  o:insetmode="auto">
  <v:fill color2="infoBackground [80]"/>
  <v:shadow color="none [81]" obscured="t"/>
  <v:path o:connecttype="none"/>
  <v:textbox style='mso-direction-alt:auto'>
   <div style='text-align:left'></div>
  </v:textbox>
  <x:ClientData ObjectType="Note">
   <x:MoveWithCells/>
   <x:SizeWithCells/>
   <x:Anchor>
    2, 10, 0, 12, 4, 20, 4, 17</x:Anchor>
   <x:AutoFill>False</x:AutoFill>
   <x:Row>1</x:Row>
   <x:Column>1</x:Column>
  </x:ClientData>
 </v:shape><v:shape id="_x0000_s1026" type="#_x0000_t202" style='position:absolute;
  margin-left:199.5pt;margin-top:38.5pt;width:104pt;height:61.5pt;z-index:2;
  visibility:hidden' fillcolor="infoBackground [80]" strokecolor="none [81]"
  o:insetmode="auto">
  <v:fill color2="infoBackground [80]"/>
  <v:shadow color="none [81]" obscured="t"/>
  <v:path o:connecttype="none"/>
  <v:textbox style='mso-direction-alt:auto'>
   <div style='text-align:left'></div>
  </v:textbox>
  <x:ClientData ObjectType="Note">
   <x:MoveWithCells/>
   <x:SizeWithCells/>
   <x:Anchor>
    4, 10, 2, 12, 6, 20, 6, 17</x:Anchor>
   <x:AutoFill>False</x:AutoFill>
   <x:Row>3</x:Row>
   <x:Column>3</x:Column>
  </x:ClientData>
 </v:shape><v:shape id="_x0000_s1027" type="#_x0000_t202" style='position:absolute;
  margin-left:295.5pt;margin-top:67.5pt;width:104pt;height:61.5pt;z-index:3;
  visibility:hidden' fillcolor="infoBackground [80]" strokecolor="none [81]"
  o:insetmode="auto">
  <v:fill color2="infoBackground [80]"/>
  <v:shadow color="none [81]" obscured="t"/>
  <v:path o:connecttype="none"/>
  <v:textbox style='mso-direction-alt:auto'>
   <div style='text-align:left'></div>
  </v:textbox>
  <x:ClientData ObjectType="Note">
   <x:MoveWithCells/>
   <x:SizeWithCells/>
   <x:Anchor>
    6, 10, 4, 12, 8, 20, 8, 17</x:Anchor>
   <x:AutoFill>False</x:AutoFill>
   <x:Row>5</x:Row>
   <x:Column>5</x:Column>
  </x:ClientData>
 </v:shape></xml>"##;

/// The comments part of the same file: one placeholder per thread, with
/// Excel's CRLF line ends inside the text.
fn excel_threads_comments_xml() -> String {
    let notice = format!("[Threaded comment]\r\n\r\n{EXCEL_THREAD_NOTICE}\r\n\r\nComment:\r\n    ");
    format!(
        concat!(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#,
            "\r\n",
            r#"<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="xr" xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision">"#,
            "<authors><author>tc={{5FA6E5A7-6BD7-4105-A288-53EB03A03DCF}}</author><author>tc={{CDE37F6B-39FC-41E4-A6C1-0816C3C6187E}}</author><author>tc={{BFCEF060-A399-48BE-BFBB-B59039308924}}</author></authors>",
            r#"<commentList><comment ref="B2" authorId="0" shapeId="0" xr:uid="{{5FA6E5A7-6BD7-4105-A288-53EB03A03DCF}}"><text><t>{n}Please check</t></text></comment>"#,
            r#"<comment ref="D4" authorId="1" shapeId="0" xr:uid="{{CDE37F6B-39FC-41E4-A6C1-0816C3C6187E}}"><text><t>{n}Root on D4"#,
            "\r\nReply:\r\n    First reply\r\nReply:\r\n    Second reply</t></text></comment>",
            r#"<comment ref="F6" authorId="2" shapeId="0" xr:uid="{{BFCEF060-A399-48BE-BFBB-B59039308924}}"><text><t>{n}Line one"#,
            "\r\nLine two\r\nReply:\r\n    Reply one\r\nReply two</t></text></comment></commentList></comments>",
        ),
        n = notice
    )
}

/// A package whose threads, persons, comments and note drawing are the
/// parts Excel saved ([`EXCEL_THREADS_XML`] and the others), under the
/// names and relationships this writer gives them.
fn excel_threads_package() -> Vec<u8> {
    let mut wb = minimal_workbook("Sheet1", Vec::new());
    let ws = &mut wb.sheets[0].worksheet;
    ws.threaded_comments =
        threaded_comments::parse_threaded_comments(EXCEL_THREADS_XML.as_bytes()).unwrap();
    ws.comments = comments::parse_comments(excel_threads_comments_xml().as_bytes()).unwrap();
    wb.persons = threaded_comments::parse_persons(EXCEL_PERSONS_XML.as_bytes()).unwrap();
    let mut entries = zip_map(&write_xlsx(&wb).unwrap());
    for (path, data) in [
        (
            "xl/threadedComments/threadedComment1.xml",
            EXCEL_THREADS_XML.to_owned(),
        ),
        ("xl/persons/person.xml", EXCEL_PERSONS_XML.to_owned()),
        ("xl/comments1.xml", excel_threads_comments_xml()),
        (
            "xl/drawings/vmlDrawing1.vml",
            EXCEL_THREADS_VML.replace('\n', "\r\n"),
        ),
    ] {
        assert!(entries.contains_key(path), "the writer names {path}");
        entries.insert(path.to_owned(), data.into_bytes());
    }
    rezip(entries)
}

#[test]
fn an_excel_saved_thread_file_read_and_written_back_keeps_its_threads_and_placeholders() {
    let package = excel_threads_package();
    let read = crate::reader::read_xlsx(&package).unwrap();
    let ws = &read.sheets[0].worksheet;
    assert_eq!(ws.threaded_comments.len(), 6);
    assert_eq!(ws.comments.len(), 3);

    let before = serde_json::to_value(&read).unwrap();
    let written = write_xlsx(&read).unwrap();
    assert_eq!(
        serde_json::to_value(&read).unwrap(),
        before,
        "the model changed"
    );
    let entries = zip_map(&written);
    assert_sheet_part_refs(&entries, "Excel threads");

    // Ids, persons and placeholders as Excel wrote them: nothing added.
    let again = crate::reader::read_xlsx(&written).unwrap();
    assert_eq!(
        again.sheets[0].worksheet.threaded_comments,
        ws.threaded_comments
    );
    assert_eq!(again.persons, read.persons);
    assert_eq!(
        sheet_comments(&entries, 1),
        [
            ("B2", ROOT_ID, excel_placeholder_text("Please check", &[])),
            (
                "D4",
                "{CDE37F6B-39FC-41E4-A6C1-0816C3C6187E}",
                excel_placeholder_text("Root on D4", &["First reply", "Second reply"]),
            ),
            (
                "F6",
                "{BFCEF060-A399-48BE-BFBB-B59039308924}",
                excel_placeholder_text("Line one\nLine two", &["Reply one\nReply two"]),
            ),
        ]
        .map(|(cell, id, text)| (cell.to_owned(), format!("tc={id}"), text))
    );
    // Excel's note drawing matches the placeholders and keeps its bytes.
    assert_eq!(
        legacy_vml(&entries, 1).1,
        EXCEL_THREADS_VML.replace('\n', "\r\n")
    );
    let persons = part_text(&entries, "xl/persons/person.xml");
    assert!(
        persons.contains(r#"id="{C71F189D-7DBF-40FC-91D4-FE2936EFAE3F}""#)
            && persons.contains(r#"userId="0123456789abcdef""#)
            && persons.contains(r#"providerId="Windows Live""#),
        "{persons}"
    );
}

// ---------------------------------------------------------------------------
// Part-name allocation in linear time
// ---------------------------------------------------------------------------

/// The allocator as it was before the per-prefix cursor: probes `n` from 1 on
/// every call. Kept as the reference the cursor allocator must agree with.
fn probing_alloc(used: &mut HashSet<String>, prefix: &str, ext: &str) -> String {
    let mut n = 1;
    loop {
        let name = format!("{prefix}{n}{ext}");
        if used.insert(name.to_ascii_lowercase()) {
            return name;
        }
        n += 1;
    }
}

/// [`probing_alloc`] with a preferred number, as `alloc_preferring` was.
fn probing_alloc_preferring(
    used: &mut HashSet<String>,
    prefix: &str,
    preferred: usize,
    ext: &str,
) -> String {
    let name = format!("{prefix}{preferred}{ext}");
    if used.insert(name.to_ascii_lowercase()) {
        name
    } else {
        probing_alloc(used, prefix, ext)
    }
}

#[test]
fn cursor_allocation_returns_the_names_of_the_probing_loop() {
    const PREFIXES: [(&str, &str); 5] = [
        ("xl/tables/table", ".xml"),
        ("xl/charts/chart", ".xml"),
        ("xl/charts/Chart", ".xml"),
        ("xl/drawings/drawing", ".xml"),
        ("xl/drawings/vmlDrawing", ".vml"),
    ];
    // xorshift64: a fixed pseudo-random sequence of operations.
    let mut x: u64 = 0x2545_F491_4F6C_DD1D;
    let mut next = move |bound: u64| {
        x ^= x << 13;
        x ^= x >> 7;
        x ^= x << 17;
        usize::try_from(x % bound).unwrap()
    };
    for round in 0..20 {
        // Preserved names, some in another case, some far above the rest.
        let preserved: Vec<String> = (0..next(40))
            .map(|_| {
                let (prefix, ext) = PREFIXES[next(5)];
                let name = format!("{prefix}{}{ext}", 1 + next(60));
                if next(2) == 0 {
                    name.to_ascii_uppercase()
                } else {
                    name
                }
            })
            .collect();
        let mut reference: HashSet<String> =
            preserved.iter().map(|p| p.to_ascii_lowercase()).collect();
        let mut names = PartNames::new(preserved.iter());
        for step in 0..400 {
            let (prefix, ext) = PREFIXES[next(5)];
            let (got, want) = match next(4) {
                0 | 1 => (
                    names.alloc(prefix, ext),
                    probing_alloc(&mut reference, prefix, ext),
                ),
                2 => {
                    let preferred = 1 + next(120);
                    (
                        names.alloc_preferring(prefix, preferred, ext),
                        probing_alloc_preferring(&mut reference, prefix, preferred, ext),
                    )
                }
                _ => {
                    // A name taken outside alloc (a drawing kept at its own path).
                    let name = format!("{prefix}{}{ext}", 1 + next(150));
                    names.take(&name);
                    reference.insert(name.to_ascii_lowercase());
                    continue;
                }
            };
            assert_eq!(got, want, "round {round}, step {step}");
        }
    }
}

#[test]
fn thirty_thousand_tables_write_in_linear_time() {
    use crate::ooxml::tables::{TableColumn, TableDefinition};
    use std::time::{Duration, Instant};
    const TABLES: usize = 30_000;
    const SHEETS: usize = 3;
    // Measured in a debug build on the dev host: 5.3 s with the cursor, 345 s
    // with the old probe-from-1 allocator (about 450 million probes).
    const LIMIT: Duration = Duration::from_secs(20);

    let table = |n: usize| TableDefinition {
        id: u32::try_from(n).unwrap(),
        name: Some(format!("T{n}")),
        display_name: format!("T{n}"),
        ref_range: "A1:A2".into(),
        header_row_count: 1,
        totals_row_count: 0,
        totals_row_shown: false,
        columns: vec![TableColumn {
            id: 1,
            name: "C".into(),
            ..Default::default()
        }],
        style_info: None,
        auto_filter_ref: None,
    };
    let per_sheet = TABLES / SHEETS;
    let wb = WorkbookData {
        sheets: (0..SHEETS)
            .map(|s| SheetData {
                name: format!("S{}", s + 1),
                worksheet: WorksheetXml {
                    tables: (1..=per_sheet).map(|t| table(s * per_sheet + t)).collect(),
                    ..Default::default()
                },
                ..Default::default()
            })
            .collect(),
        ..Default::default()
    };
    let start = Instant::now();
    let bytes = write_xlsx(&wb).expect("write");
    let elapsed = start.elapsed();
    assert!(
        elapsed < LIMIT,
        "{TABLES} tables took {elapsed:?}, limit {LIMIT:?}"
    );

    let limits = ZipSecurityLimits::default();
    let names: HashSet<String> = read_zip_entries(&bytes, &limits)
        .unwrap()
        .into_keys()
        .collect();
    assert!((1..=TABLES).all(|n| names.contains(&format!("xl/tables/table{n}.xml"))));
    assert!(!names.contains(&format!("xl/tables/table{}.xml", TABLES + 1)));
}

/// A one-cell sheet named `name` whose A1 holds the number `value`.
fn number_sheet(name: &str, value: &str) -> SheetData {
    let mut wb = minimal_workbook(
        name,
        vec![Row {
            index: 1,
            cells: vec![Cell {
                reference: "A1".into(),
                cell_type: CellType::Number,
                value: Some(value.into()),
                ..Default::default()
            }],
            ..Default::default()
        }],
    );
    wb.sheets.remove(0)
}

#[test]
fn an_unreferenced_preserved_part_at_a_generated_name_gives_way() {
    // Sheets S1 (sheet1.xml) and S2 (sheet2.xml); a mangled <sheet> row
    // leaves sheet1.xml unreferenced, so the read keeps it as a preserved
    // entry while S2 becomes the first sheet, which the writer names
    // sheet1.xml. The write failed with ZIP_WRITE ("Duplicate filename").
    let wb = WorkbookData {
        sheets: vec![number_sheet("S1", "1"), number_sheet("S2", "2")],
        ..Default::default()
    };
    let mut parts: Vec<ZipEntry> = zip_map(&write_xlsx(&wb).unwrap())
        .into_iter()
        .map(|(name, data)| ZipEntry { name, data })
        .collect();
    parts.sort_by(|a, b| a.name.cmp(&b.name));
    let workbook = parts
        .iter_mut()
        .find(|p| p.name == "xl/workbook.xml")
        .unwrap();
    let xml = String::from_utf8(std::mem::take(&mut workbook.data)).unwrap();
    assert!(xml.contains(r#"<sheet name="S1""#), "{xml}");
    workbook.data = xml
        .replace(r#"<sheet name="S1""#, r#"<sxeet name="S1""#)
        .into_bytes();
    let read = crate::reader::read_xlsx(&write_zip(&parts).unwrap()).unwrap();
    assert_eq!(read.sheets.len(), 1);
    assert_eq!(read.sheets[0].name, "S2");
    let stale = read.preserved_entries["xl/worksheets/sheet1.xml"].clone();

    let written = write_xlsx(&read).expect("the generated part wins");
    let entries = zip_map(&written);
    assert!(
        entries["xl/worksheets/sheet1.xml"] != stale,
        "sheet1.xml is S2's generated part"
    );
    let again = crate::reader::read_xlsx(&written).unwrap();
    assert_eq!(again.sheets.len(), 1);
    assert_eq!(again.sheets[0].name, "S2");
    let a1 = &again.sheets[0].worksheet.rows[0].cells[0];
    assert_eq!(
        (a1.reference.as_str(), a1.value.as_deref()),
        ("A1", Some("2"))
    );
    // The dropped entry adds no content type of its own.
    let ct = ContentTypes::parse(&entries["[Content_Types].xml"]).unwrap();
    assert_eq!(
        ct.overrides["/xl/worksheets/sheet1.xml"],
        "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"
    );
}

/// The consolidation source Excel 16 saves for a pivot table over a range in
/// another workbook: its `<rangeSet r:id>` names the cache definition's
/// `externalLinkPath` relationship (`ID` stands for that id).
const EXTERNAL_CONSOLIDATION_SOURCE: &str = r#"<cacheSource type="consolidation"><consolidation autoPage="0"><rangeSets count="1"><rangeSet ref="A1:C3" sheet="Data" r:id="ID"/></rangeSets></consolidation></cacheSource>"#;

/// A one-sheet workbook with one pivot cache over
/// [`EXTERNAL_CONSOLIDATION_SOURCE`] (its `r:id` set to `source_id`) that
/// holds `relationships`, with records when `records` is set.
fn external_pivot_workbook(
    source_id: &str,
    relationships: serde_json::Value,
    records: bool,
) -> WorkbookData {
    use serde_json::json;

    let mut wb = minimal_workbook("Sheet1", Vec::new());
    wb.pivot_caches = vec![from_json(json!({
        "source": {
            "ref": "",
            "sheet": "",
            "rawXml": EXTERNAL_CONSOLIDATION_SOURCE.replace("ID", source_id),
        },
        "fields": [{ "name": "Row" }],
        "relationships": relationships,
    }))];
    if records {
        wb.pivot_cache_records = vec![from_json(
            json!({ "records": [[{ "type": "number", "v": 1.0 }]] }),
        )];
    }
    wb
}

/// An `externalLinkPath` relationship as the JSON bridge carries it.
fn external_link_json(id: &str) -> serde_json::Value {
    serde_json::json!({
        "id": id,
        "type": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath",
        "target": "xcons-source.xlsx",
        "targetMode": "External",
    })
}

#[test]
fn a_pivot_cache_writes_its_kept_relationships_with_their_ids() {
    use crate::ooxml::relationships::TargetMode;

    const CACHE: &str = "xl/pivotCache/pivotCacheDefinition1.xml";
    const CACHE_RELS: &str = "xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels";
    // (source r:id, kept relationship ids, records, records id)
    let cases: [(&str, &[&str], bool, Option<&str>); 4] = [
        // Excel's shape: the records keep rId1.
        ("rId2", &["rId2"], true, Some("rId1")),
        // A kept relationship holds rId1: the records take the next free id.
        ("rId1", &["rId1"], true, Some("rId2")),
        // No records: the .rels still carries the kept relationship.
        ("rId2", &["rId2"], false, None),
        // rId1 and rId2 are held.
        ("rId2", &["rId2", "rId1"], true, Some("rId3")),
    ];
    for (source_id, kept, records, records_id) in cases {
        let rels_json = kept.iter().map(|id| external_link_json(id)).collect();
        let wb = external_pivot_workbook(source_id, serde_json::Value::Array(rels_json), records);
        let once = write_xlsx(&wb).unwrap();
        let parts = zip_map(&once);
        let rels = Relationships::parse(&parts[CACHE_RELS]).unwrap();
        assert_eq!(rels.len(), kept.len() + usize::from(records), "{kept:?}");
        for id in kept {
            let rel = rels.get_by_id(id).unwrap();
            assert_eq!(
                (rel.rel_type.as_ref(), rel.target.as_ref(), rel.target_mode),
                (
                    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath",
                    "xcons-source.xlsx",
                    TargetMode::External
                )
            );
        }
        let def = String::from_utf8(parts[CACHE].clone()).unwrap();
        let root = &def[..def.find("<cacheSource").unwrap()];
        match records_id {
            Some(id) => {
                assert_eq!(rels.get_by_id(id).unwrap().rel_type, REL_PIVOT_CACHE_REC);
                assert!(root.contains(&format!(r#" r:id="{id}""#)), "{root}");
            }
            None => assert!(!root.contains(" r:id="), "{root}"),
        }
        let source = EXTERNAL_CONSOLIDATION_SOURCE.replace("ID", source_id);
        assert!(def.contains(&format!("{source}<cacheFields")), "{def}");

        // Read back as written, and a second generation writes the same parts.
        let read = crate::reader::read_xlsx(&once).unwrap();
        let cache = PivotCacheDefinitionData {
            cache_id: None,
            ..read.pivot_caches[0].clone()
        };
        assert_eq!(cache, wb.pivot_caches[0], "{kept:?}");
        let again = zip_map(&write_xlsx(&read).unwrap());
        assert_eq!(again[CACHE], parts[CACHE]);
        assert_eq!(again[CACHE_RELS], parts[CACHE_RELS]);
    }
}

#[test]
fn a_pivot_cache_without_other_relationships_is_written_as_before() {
    const CACHE_RELS: &str = "xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels";
    let wb = external_pivot_workbook("rId2", serde_json::json!([]), true);
    let parts = zip_map(&write_xlsx(&wb).unwrap());
    assert_eq!(
        String::from_utf8(parts[CACHE_RELS].clone()).unwrap(),
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheRecords" Target="pivotCacheRecords1.xml"/></Relationships>"#
    );
    // The source's r:id names no written relationship: it is dropped, so it
    // cannot dangle (Excel refuses such a file).
    let def = String::from_utf8(parts["xl/pivotCache/pivotCacheDefinition1.xml"].clone()).unwrap();
    assert!(
        def.contains(r#"<rangeSet ref="A1:C3" sheet="Data"/>"#),
        "{def}"
    );
    // Without records or kept relationships there is no .rels at all.
    let bare = external_pivot_workbook("rId2", serde_json::json!([]), false);
    assert!(!zip_map(&write_xlsx(&bare).unwrap()).contains_key(CACHE_RELS));
}

/// A written one-sheet package (A1 = 7 on sheet `Data`), its parts edited by
/// `edit`, then zipped again.
fn edited_package(edit: impl FnOnce(&mut BTreeMap<String, Vec<u8>>)) -> Vec<u8> {
    let wb = WorkbookData {
        sheets: vec![number_sheet("Data", "7")],
        ..Default::default()
    };
    let mut parts: BTreeMap<String, Vec<u8>> =
        zip_map(&write_xlsx(&wb).unwrap()).into_iter().collect();
    edit(&mut parts);
    let entries: Vec<ZipEntry> = parts
        .into_iter()
        .map(|(name, data)| ZipEntry { name, data })
        .collect();
    write_zip(&entries).unwrap()
}

/// Replaces `from` with `to` in the text part `name`, which must hold it.
fn replace_in(parts: &mut BTreeMap<String, Vec<u8>>, name: &str, from: &str, to: &str) {
    let text = String::from_utf8(parts[name].clone()).unwrap();
    assert!(text.contains(from), "{name}: {text}");
    parts.insert(name.to_owned(), text.replace(from, to).into_bytes());
}

/// Moves part `from` to `to`, renaming its workbook relationship target
/// (`rel_from` to `rel_to`) and its Override.
fn move_part(
    parts: &mut BTreeMap<String, Vec<u8>>,
    (from, to): (&str, &str),
    (rel_from, rel_to): (&str, &str),
) {
    let data = parts.remove(from).unwrap();
    parts.insert(to.to_owned(), data);
    replace_in(parts, "xl/_rels/workbook.xml.rels", rel_from, rel_to);
    replace_in(
        parts,
        "[Content_Types].xml",
        &format!("\"/{from}\""),
        &format!("\"/{to}\""),
    );
}

/// Adds the part `name` holding `data`, with an Override of `content_type`.
fn add_part(parts: &mut BTreeMap<String, Vec<u8>>, name: &str, data: Vec<u8>, content_type: &str) {
    parts.insert(name.to_owned(), data);
    replace_in(
        parts,
        "[Content_Types].xml",
        "</Types>",
        &format!(r#"<Override PartName="/{name}" ContentType="{content_type}"/></Types>"#),
    );
}

/// Reads `pkg`, which keeps `entry` as a preserved entry whose name differs
/// only in letter case from the `generated` part, writes it, and checks that
/// the written package names no part twice ignoring case, has no Override
/// for the dropped spelling and reads back A1 = 7.
fn assert_case_variant_gives_way(
    pkg: &[u8],
    entry: &str,
    generated: &str,
) -> HashMap<String, Vec<u8>> {
    let mut read = crate::reader::read_xlsx(pkg).unwrap();
    // The reader takes a case variant of a part it looks up by a fixed name
    // (xl/Styles.xml, xl/SharedStrings.xml) as that part and keeps no copy;
    // it is then supplied as a JSON caller would, so the give-way below is
    // still exercised.
    let read_as_fixed_part =
        ["xl/styles.xml", "xl/sharedstrings.xml"].contains(&entry.to_ascii_lowercase().as_str());
    assert_eq!(
        read.preserved_entries.contains_key(entry),
        !read_as_fixed_part,
        "{entry} is kept unless the reader reads it as its fixed-name part"
    );
    if read_as_fixed_part {
        read.preserved_entries
            .insert(entry.to_owned(), zip_map(pkg)[entry].clone());
    }
    let written = write_xlsx(&read).expect("the generated part wins");
    let entries = zip_map(&written);
    assert!(!entries.contains_key(entry), "{entry} is dropped");
    assert!(entries.contains_key(generated), "{generated} is written");
    let mut seen = HashSet::new();
    for name in entries.keys() {
        assert!(seen.insert(name.to_ascii_lowercase()), "{name} twice");
    }
    let ct = ContentTypes::parse(&entries["[Content_Types].xml"]).unwrap();
    assert!(
        !ct.overrides.contains_key(&format!("/{entry}")),
        "{:?}",
        ct.overrides
    );
    let mut seen = HashSet::new();
    for part in ct.overrides.keys() {
        assert!(
            seen.insert(part.to_ascii_lowercase()),
            "Override {part} twice"
        );
    }
    let again = crate::reader::read_xlsx(&written).unwrap();
    let a1 = &again.sheets[0].worksheet.rows[0].cells[0];
    assert_eq!(
        (a1.reference.as_str(), a1.value.as_deref()),
        ("A1", Some("7"))
    );
    entries
}

#[test]
fn a_preserved_entry_gives_way_to_a_generated_name_in_any_letter_case() {
    const WS: &str = "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml";
    const STYLES: &str = "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml";
    // P1: the sheet at xl/worksheets/data.xml, an unreferenced worksheet part
    // at xl/worksheets/Sheet1.xml beside the generated sheet1.xml.
    let p1 = edited_package(|parts| {
        let sheet = parts["xl/worksheets/sheet1.xml"].clone();
        move_part(
            parts,
            ("xl/worksheets/sheet1.xml", "xl/worksheets/data.xml"),
            ("worksheets/sheet1.xml", "worksheets/data.xml"),
        );
        add_part(parts, "xl/worksheets/Sheet1.xml", sheet, WS);
    });
    assert_case_variant_gives_way(&p1, "xl/worksheets/Sheet1.xml", "xl/worksheets/sheet1.xml");
    // P2: the styles at xl/st.xml, an unreferenced copy at xl/Styles.xml.
    let p2 = edited_package(|parts| {
        let styles = parts["xl/styles.xml"].clone();
        move_part(
            parts,
            ("xl/styles.xml", "xl/st.xml"),
            (r#"Target="styles.xml""#, r#"Target="st.xml""#),
        );
        add_part(parts, "xl/Styles.xml", styles, STYLES);
    });
    assert_case_variant_gives_way(&p2, "xl/Styles.xml", "xl/styles.xml");
    // P3: the shared strings at xl/sst.xml; at xl/SharedStrings.xml a custom
    // XML part a kept sheet relationship names. The entry still gives way:
    // the relationship now resolves (OPC names ignore case) to the generated
    // part, the one part that answers to that name.
    let p3 = edited_package(|parts| {
        move_part(
            parts,
            ("xl/sharedStrings.xml", "xl/sst.xml"),
            (r#"Target="sharedStrings.xml""#, r#"Target="sst.xml""#),
        );
        add_part(
            parts,
            "xl/SharedStrings.xml",
            br#"<?xml version="1.0" encoding="UTF-8"?><note>custom data</note>"#.to_vec(),
            "application/xml",
        );
        parts.insert(
            "xl/worksheets/_rels/sheet1.xml.rels".to_owned(),
            br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml" Target="../SharedStrings.xml"/></Relationships>"#.to_vec(),
        );
    });
    let entries =
        assert_case_variant_gives_way(&p3, "xl/SharedStrings.xml", "xl/sharedStrings.xml");
    assert!(
        entries["xl/sharedStrings.xml"]
            .windows(4)
            .any(|w| w == b"<sst")
    );
}
