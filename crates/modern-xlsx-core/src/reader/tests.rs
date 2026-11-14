use super::*;
use crate::dates::DateSystem;
use crate::ooxml::SPREADSHEET_NS;
use crate::ooxml::{
    content_types::ContentTypes,
    relationships::Relationships,
    shared_strings::SharedStringTableBuilder,
    styles::Styles,
    workbook::{SheetInfo, SheetState, WorkbookXml},
    worksheet::{Cell, CellType, Row, WorksheetXml},
};
use crate::zip::writer::{ZipEntry, write_zip};
use pretty_assertions::assert_eq;

/// Build a minimal valid XLSX archive as raw bytes.
///
/// The workbook contains one sheet named "Sheet1" with two cells:
///   A1 = "Hello" (shared string)
///   B1 = 42 (number)
fn build_minimal_xlsx() -> Vec<u8> {
    // Content types
    let ct = ContentTypes::for_basic_workbook(1);
    let ct_xml = ct.to_xml().unwrap();

    // Root rels
    let root_rels = Relationships::root_rels();
    let root_rels_xml = root_rels.to_xml().unwrap();

    // Workbook rels
    let wb_rels = Relationships::workbook_rels(1);
    let wb_rels_xml = wb_rels.to_xml().unwrap();

    // Workbook
    let wb = WorkbookXml {
        sheets: vec![SheetInfo {
            name: "Sheet1".to_string(),
            sheet_id: 1,
            r_id: "rId1".to_string(),
            state: SheetState::Visible,
        }],
        date_system: DateSystem::Date1900,
        defined_names: Vec::new(),
        workbook_views: Vec::new(),
        protection: None,
        ..Default::default()
    };
    let wb_xml = wb.to_xml().unwrap();

    // Shared strings: index 0 = "Hello"
    let mut sst_builder = SharedStringTableBuilder::new();
    sst_builder.insert("Hello");
    let sst_xml = sst_builder.to_xml().unwrap();

    // Styles
    let styles = Styles::default_styles();
    let styles_xml = styles.to_xml().unwrap();

    // Worksheet with A1="Hello" (shared string index 0), B1=42 (number)
    let ws = WorksheetXml {
        dimension: Some("A1:B1".to_string()),
        rows: vec![Row {
            index: 1,
            cells: vec![
                Cell {
                    reference: "A1".to_string(),
                    cell_type: CellType::SharedString,
                    style_index: None,
                    value: Some("0".to_string()),
                    ..Default::default()
                },
                Cell {
                    reference: "B1".to_string(),
                    cell_type: CellType::Number,
                    style_index: None,
                    value: Some("42".to_string()),
                    ..Default::default()
                },
            ],
            ..Default::default()
        }],
        ..Default::default()
    };
    let ws_xml = ws.to_xml().unwrap();

    let entries = vec![
        ZipEntry {
            name: "[Content_Types].xml".to_string(),
            data: ct_xml,
        },
        ZipEntry {
            name: "_rels/.rels".to_string(),
            data: root_rels_xml,
        },
        ZipEntry {
            name: "xl/workbook.xml".to_string(),
            data: wb_xml,
        },
        ZipEntry {
            name: "xl/_rels/workbook.xml.rels".to_string(),
            data: wb_rels_xml,
        },
        ZipEntry {
            name: "xl/sharedStrings.xml".to_string(),
            data: sst_xml,
        },
        ZipEntry {
            name: "xl/styles.xml".to_string(),
            data: styles_xml,
        },
        ZipEntry {
            name: "xl/worksheets/sheet1.xml".to_string(),
            data: ws_xml,
        },
    ];

    write_zip(&entries).unwrap()
}

/// Build an XLSX archive without a sharedStrings.xml part.
/// The worksheet uses only numeric cells (no shared strings needed).
fn build_xlsx_without_sst() -> Vec<u8> {
    // Content types — manually construct without the SST override
    let mut ct = ContentTypes::for_basic_workbook(1);
    ct.overrides.remove("/xl/sharedStrings.xml");
    let ct_xml = ct.to_xml().unwrap();

    // Root rels
    let root_rels = Relationships::root_rels();
    let root_rels_xml = root_rels.to_xml().unwrap();

    // Workbook rels — only worksheet + styles, no sharedStrings
    let mut wb_rels = Relationships::new();
    wb_rels.add(
        "rId1",
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet",
        "worksheets/sheet1.xml",
    );
    wb_rels.add(
        "rId2",
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles",
        "styles.xml",
    );
    let wb_rels_xml = wb_rels.to_xml().unwrap();

    // Workbook
    let wb = WorkbookXml {
        sheets: vec![SheetInfo {
            name: "Numbers".to_string(),
            sheet_id: 1,
            r_id: "rId1".to_string(),
            state: SheetState::Visible,
        }],
        date_system: DateSystem::Date1900,
        defined_names: Vec::new(),
        workbook_views: Vec::new(),
        protection: None,
        ..Default::default()
    };
    let wb_xml = wb.to_xml().unwrap();

    // Styles
    let styles = Styles::default_styles();
    let styles_xml = styles.to_xml().unwrap();

    // Worksheet — numeric only
    let ws = WorksheetXml {
        dimension: Some("A1".to_string()),
        rows: vec![Row {
            index: 1,
            cells: vec![Cell {
                reference: "A1".to_string(),
                cell_type: CellType::Number,
                style_index: None,
                value: Some("99".to_string()),
                ..Default::default()
            }],
            ..Default::default()
        }],
        ..Default::default()
    };
    let ws_xml = ws.to_xml().unwrap();

    let entries = vec![
        ZipEntry {
            name: "[Content_Types].xml".to_string(),
            data: ct_xml,
        },
        ZipEntry {
            name: "_rels/.rels".to_string(),
            data: root_rels_xml,
        },
        ZipEntry {
            name: "xl/workbook.xml".to_string(),
            data: wb_xml,
        },
        ZipEntry {
            name: "xl/_rels/workbook.xml.rels".to_string(),
            data: wb_rels_xml,
        },
        ZipEntry {
            name: "xl/styles.xml".to_string(),
            data: styles_xml,
        },
        ZipEntry {
            name: "xl/worksheets/sheet1.xml".to_string(),
            data: ws_xml,
        },
    ];

    write_zip(&entries).unwrap()
}

/// Build an XLSX archive that uses the 1904 date system.
fn build_xlsx_1904() -> Vec<u8> {
    // Content types
    let ct = ContentTypes::for_basic_workbook(1);
    let ct_xml = ct.to_xml().unwrap();

    // Root rels
    let root_rels = Relationships::root_rels();
    let root_rels_xml = root_rels.to_xml().unwrap();

    // Workbook rels
    let wb_rels = Relationships::workbook_rels(1);
    let wb_rels_xml = wb_rels.to_xml().unwrap();

    // Workbook with date1904
    let wb = WorkbookXml {
        sheets: vec![SheetInfo {
            name: "Sheet1".to_string(),
            sheet_id: 1,
            r_id: "rId1".to_string(),
            state: SheetState::Visible,
        }],
        date_system: DateSystem::Date1904,
        defined_names: Vec::new(),
        workbook_views: Vec::new(),
        protection: None,
        ..Default::default()
    };
    let wb_xml = wb.to_xml().unwrap();

    // Shared strings
    let sst_builder = SharedStringTableBuilder::new();
    let sst_xml = sst_builder.to_xml().unwrap();

    // Styles
    let styles = Styles::default_styles();
    let styles_xml = styles.to_xml().unwrap();

    // Empty worksheet
    let ws = WorksheetXml {
        ..Default::default()
    };
    let ws_xml = ws.to_xml().unwrap();

    let entries = vec![
        ZipEntry {
            name: "[Content_Types].xml".to_string(),
            data: ct_xml,
        },
        ZipEntry {
            name: "_rels/.rels".to_string(),
            data: root_rels_xml,
        },
        ZipEntry {
            name: "xl/workbook.xml".to_string(),
            data: wb_xml,
        },
        ZipEntry {
            name: "xl/_rels/workbook.xml.rels".to_string(),
            data: wb_rels_xml,
        },
        ZipEntry {
            name: "xl/sharedStrings.xml".to_string(),
            data: sst_xml,
        },
        ZipEntry {
            name: "xl/styles.xml".to_string(),
            data: styles_xml,
        },
        ZipEntry {
            name: "xl/worksheets/sheet1.xml".to_string(),
            data: ws_xml,
        },
    ];

    write_zip(&entries).unwrap()
}

#[test]
fn test_read_minimal_workbook() {
    let xlsx_bytes = build_minimal_xlsx();
    let wb = read_xlsx(&xlsx_bytes).expect("read_xlsx should succeed");

    // One sheet named "Sheet1".
    assert_eq!(wb.sheets.len(), 1);
    assert_eq!(wb.sheets[0].name, "Sheet1");

    // Date system is 1900 (default).
    assert_eq!(wb.date_system, DateSystem::Date1900);

    // Shared string table has one entry: "Hello".
    let sst = wb
        .shared_strings
        .as_ref()
        .expect("shared_strings should be Some");
    assert_eq!(sst.len(), 1);
    assert_eq!(sst.get(0), Some("Hello"));

    // Worksheet has one row with two cells.
    let ws = &wb.sheets[0].worksheet;
    assert_eq!(ws.rows.len(), 1);
    assert_eq!(ws.rows[0].cells.len(), 2);

    // A1 is a shared string reference (value resolved to "Hello").
    let a1 = &ws.rows[0].cells[0];
    assert_eq!(a1.reference, "A1");
    assert_eq!(a1.cell_type, CellType::SharedString);
    assert_eq!(a1.value.as_deref(), Some("Hello"));

    // B1 is a number (value = "42").
    let b1 = &ws.rows[0].cells[1];
    assert_eq!(b1.reference, "B1");
    assert_eq!(b1.cell_type, CellType::Number);
    assert_eq!(b1.value.as_deref(), Some("42"));
}

#[test]
fn test_read_without_shared_strings() {
    let xlsx_bytes = build_xlsx_without_sst();
    let wb = read_xlsx(&xlsx_bytes).expect("read_xlsx should succeed without SST");

    // Sheet should still load.
    assert_eq!(wb.sheets.len(), 1);
    assert_eq!(wb.sheets[0].name, "Numbers");

    // Shared strings table should be present but empty (not an error).
    let sst = wb
        .shared_strings
        .as_ref()
        .expect("shared_strings should be Some");
    assert!(sst.is_empty());

    // The numeric cell should parse correctly.
    let ws = &wb.sheets[0].worksheet;
    assert_eq!(ws.rows.len(), 1);
    let a1 = &ws.rows[0].cells[0];
    assert_eq!(a1.cell_type, CellType::Number);
    assert_eq!(a1.value.as_deref(), Some("99"));
}

#[test]
fn test_read_1904_date_system() {
    let xlsx_bytes = build_xlsx_1904();
    let wb = read_xlsx(&xlsx_bytes).expect("read_xlsx should succeed for 1904 workbook");

    assert_eq!(wb.date_system, DateSystem::Date1904);
    assert_eq!(wb.sheets.len(), 1);
    assert_eq!(wb.sheets[0].name, "Sheet1");
}

#[test]
fn test_read_missing_workbook_xml() {
    // Build a ZIP with no xl/workbook.xml.
    let entries = vec![ZipEntry {
        name: "[Content_Types].xml".to_string(),
        data: b"<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"/>"
            .to_vec(),
    }];
    let zip_bytes = write_zip(&entries).unwrap();

    let result = read_xlsx(&zip_bytes);
    assert!(result.is_err());
    let err = result.unwrap_err();
    assert!(
        matches!(err, ModernXlsxError::MissingPart(ref msg) if msg.contains("workbook.xml")),
        "expected MissingPart error, got: {err}"
    );
}

#[test]
fn test_preserved_entries_roundtrip() {
    // 1. Build a minimal XLSX with extra ZIP entries simulating drawings/media.
    let ct = ContentTypes::for_basic_workbook(1);
    let ct_xml = ct.to_xml().unwrap();
    let root_rels = Relationships::root_rels();
    let root_rels_xml = root_rels.to_xml().unwrap();
    let wb_rels = Relationships::workbook_rels(1);
    let wb_rels_xml = wb_rels.to_xml().unwrap();
    let wb = WorkbookXml {
        sheets: vec![SheetInfo {
            name: "Sheet1".to_string(),
            sheet_id: 1,
            r_id: "rId1".to_string(),
            state: SheetState::Visible,
        }],
        date_system: DateSystem::Date1900,
        defined_names: Vec::new(),
        workbook_views: Vec::new(),
        protection: None,
        ..Default::default()
    };
    let wb_xml = wb.to_xml().unwrap();
    let sst_builder = SharedStringTableBuilder::new();
    let sst_xml = sst_builder.to_xml().unwrap();
    let styles = Styles::default_styles();
    let styles_xml = styles.to_xml().unwrap();
    let ws = WorksheetXml {
        ..Default::default()
    };
    let ws_xml = ws.to_xml().unwrap();

    // Fake drawing and image data
    let drawing_data = b"<xdr:wsDr xmlns:xdr=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\"/>".to_vec();
    let image_data = vec![0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A]; // PNG header
    let chart_data =
        b"<c:chartSpace xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\"/>"
            .to_vec();
    let sheet_rels_data =
        b"<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"/>"
            .to_vec();

    let entries = vec![
        ZipEntry {
            name: "[Content_Types].xml".to_string(),
            data: ct_xml,
        },
        ZipEntry {
            name: "_rels/.rels".to_string(),
            data: root_rels_xml,
        },
        ZipEntry {
            name: "xl/workbook.xml".to_string(),
            data: wb_xml,
        },
        ZipEntry {
            name: "xl/_rels/workbook.xml.rels".to_string(),
            data: wb_rels_xml,
        },
        ZipEntry {
            name: "xl/sharedStrings.xml".to_string(),
            data: sst_xml,
        },
        ZipEntry {
            name: "xl/styles.xml".to_string(),
            data: styles_xml,
        },
        ZipEntry {
            name: "xl/worksheets/sheet1.xml".to_string(),
            data: ws_xml,
        },
        // Extra entries to preserve
        ZipEntry {
            name: "xl/drawings/drawing1.xml".to_string(),
            data: drawing_data.clone(),
        },
        ZipEntry {
            name: "xl/media/image1.png".to_string(),
            data: image_data.clone(),
        },
        ZipEntry {
            name: "xl/charts/chart1.xml".to_string(),
            data: chart_data.clone(),
        },
        ZipEntry {
            name: "xl/worksheets/_rels/sheet1.xml.rels".to_string(),
            data: sheet_rels_data.clone(),
        },
    ];

    let zip_bytes = write_zip(&entries).unwrap();

    // 2. Read the XLSX — preserved entries should be captured.
    let wb1 = read_xlsx(&zip_bytes).expect("read_xlsx should succeed");
    assert_eq!(wb1.preserved_entries.len(), 4);
    assert_eq!(
        wb1.preserved_entries
            .get("xl/drawings/drawing1.xml")
            .unwrap(),
        &drawing_data
    );
    assert_eq!(
        wb1.preserved_entries.get("xl/media/image1.png").unwrap(),
        &image_data
    );
    assert_eq!(
        wb1.preserved_entries.get("xl/charts/chart1.xml").unwrap(),
        &chart_data
    );
    assert_eq!(
        wb1.preserved_entries
            .get("xl/worksheets/_rels/sheet1.xml.rels")
            .unwrap(),
        &sheet_rels_data
    );

    // 3. Write it back and read again — the entries should survive the roundtrip.
    let xlsx_bytes2 = crate::writer::write_xlsx(&wb1).expect("write_xlsx should succeed");
    let wb2 = read_xlsx(&xlsx_bytes2).expect("read_xlsx should succeed on second pass");

    // The writer generated the theme, so the reader recognises it and
    // consumes it instead of preserving it.
    assert_eq!(wb2.preserved_entries.len(), 4);
    assert!(!wb2.preserved_entries.contains_key("xl/theme/theme1.xml"));
    assert!(wb2.theme_colors.is_some());
    assert_eq!(
        wb2.preserved_entries
            .get("xl/drawings/drawing1.xml")
            .unwrap(),
        &drawing_data
    );
    assert_eq!(
        wb2.preserved_entries.get("xl/media/image1.png").unwrap(),
        &image_data
    );
    assert_eq!(
        wb2.preserved_entries.get("xl/charts/chart1.xml").unwrap(),
        &chart_data
    );
    assert_eq!(
        wb2.preserved_entries
            .get("xl/worksheets/_rels/sheet1.xml.rels")
            .unwrap(),
        &sheet_rels_data
    );
}

#[test]
fn test_no_preserved_entries_for_standard_xlsx() {
    // A standard XLSX with no extra entries should have an empty preserved_entries map.
    let xlsx_bytes = build_minimal_xlsx();
    let wb = read_xlsx(&xlsx_bytes).expect("read_xlsx should succeed");
    assert!(
        wb.preserved_entries.is_empty(),
        "standard XLSX should have no preserved entries"
    );
}

/// Build an XLSX with multiple sheets to exercise parallel parsing.
fn build_multi_sheet_xlsx(sheet_count: usize) -> Vec<u8> {
    let ct = ContentTypes::for_basic_workbook(sheet_count);
    let ct_xml = ct.to_xml().unwrap();

    let root_rels = Relationships::root_rels();
    let root_rels_xml = root_rels.to_xml().unwrap();

    let wb_rels = Relationships::workbook_rels(sheet_count);
    let wb_rels_xml = wb_rels.to_xml().unwrap();

    let sheets_info: Vec<SheetInfo> = (1..=sheet_count)
        .map(|i| SheetInfo {
            name: format!("Sheet{i}"),
            sheet_id: i as u32,
            r_id: format!("rId{i}"),
            state: SheetState::Visible,
        })
        .collect();

    let wb = WorkbookXml {
        sheets: sheets_info,
        date_system: DateSystem::Date1900,
        defined_names: Vec::new(),
        workbook_views: Vec::new(),
        protection: None,
        ..Default::default()
    };
    let wb_xml = wb.to_xml().unwrap();

    let sst_builder = SharedStringTableBuilder::new();
    let sst_xml = sst_builder.to_xml().unwrap();

    let styles = Styles::default_styles();
    let styles_xml = styles.to_xml().unwrap();

    let mut entries = vec![
        ZipEntry {
            name: "[Content_Types].xml".to_string(),
            data: ct_xml,
        },
        ZipEntry {
            name: "_rels/.rels".to_string(),
            data: root_rels_xml,
        },
        ZipEntry {
            name: "xl/workbook.xml".to_string(),
            data: wb_xml,
        },
        ZipEntry {
            name: "xl/_rels/workbook.xml.rels".to_string(),
            data: wb_rels_xml,
        },
        ZipEntry {
            name: "xl/sharedStrings.xml".to_string(),
            data: sst_xml,
        },
        ZipEntry {
            name: "xl/styles.xml".to_string(),
            data: styles_xml,
        },
    ];

    // Create a unique worksheet for each sheet with a distinct cell value.
    for i in 1..=sheet_count {
        let ws = WorksheetXml {
            dimension: Some("A1".to_string()),
            rows: vec![Row {
                index: 1,
                cells: vec![Cell {
                    reference: "A1".to_string(),
                    cell_type: CellType::Number,
                    style_index: None,
                    value: Some(format!("{}", i * 100)),
                    ..Default::default()
                }],
                ..Default::default()
            }],
            ..Default::default()
        };
        let ws_xml = ws.to_xml().unwrap();
        entries.push(ZipEntry {
            name: format!("xl/worksheets/sheet{i}.xml"),
            data: ws_xml,
        });
    }

    write_zip(&entries).unwrap()
}

#[test]
fn test_multi_sheet_parsing() {
    // This test exercises the parse_sheets function (parallel or sequential
    // depending on the feature flag). With 4 sheets, the parallel path has
    // enough work to actually distribute across threads.
    let xlsx_bytes = build_multi_sheet_xlsx(4);
    let wb = read_xlsx(&xlsx_bytes).expect("read_xlsx should succeed");

    assert_eq!(wb.sheets.len(), 4);

    for (i, sheet) in wb.sheets.iter().enumerate() {
        let expected_name = format!("Sheet{}", i + 1);
        assert_eq!(sheet.name, expected_name);

        // Each sheet has one row with A1 = (i+1)*100.
        assert_eq!(sheet.worksheet.rows.len(), 1);
        let cell = &sheet.worksheet.rows[0].cells[0];
        assert_eq!(cell.reference, "A1");
        assert_eq!(cell.cell_type, CellType::Number);
        assert_eq!(cell.value.as_deref(), Some(&*format!("{}", (i + 1) * 100)));
    }
}

#[test]
fn test_comments_full_roundtrip() {
    use crate::ooxml::comments::Comment;

    // 1. Build a workbook with comments via the writer.
    let wb1 = WorkbookData {
        sheets: vec![SheetData {
            name: "Sheet1".to_string(),
            worksheet: WorksheetXml {
                dimension: Some("A1:B1".to_string()),
                rows: vec![Row {
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
                            value: Some("99".to_string()),
                            ..Default::default()
                        },
                    ],
                    ..Default::default()
                }],
                comments: vec![
                    Comment {
                        cell_ref: "A1".to_string(),
                        author: "Alice".to_string(),
                        text: "This is the answer".to_string(),
                    },
                    Comment {
                        cell_ref: "B1".to_string(),
                        author: "Bob".to_string(),
                        text: "Almost 100".to_string(),
                    },
                ],
                ..Default::default()
            },
            ..Default::default()
        }],
        ..Default::default()
    };

    // Write to XLSX bytes.
    let xlsx1 = crate::writer::write_xlsx(&wb1).expect("write_xlsx should succeed");

    // Read it back.
    let wb2 = read_xlsx(&xlsx1).expect("read_xlsx should succeed");
    assert_eq!(wb2.sheets.len(), 1);
    assert_eq!(wb2.sheets[0].worksheet.comments.len(), 2);

    let c0 = &wb2.sheets[0].worksheet.comments[0];
    assert_eq!(c0.cell_ref, "A1");
    assert_eq!(c0.author, "Alice");
    assert_eq!(c0.text, "This is the answer");

    let c1 = &wb2.sheets[0].worksheet.comments[1];
    assert_eq!(c1.cell_ref, "B1");
    assert_eq!(c1.author, "Bob");
    assert_eq!(c1.text, "Almost 100");

    // Write again and read again (second roundtrip).
    let xlsx2 = crate::writer::write_xlsx(&wb2).expect("second write_xlsx should succeed");
    let wb3 = read_xlsx(&xlsx2).expect("second read_xlsx should succeed");

    assert_eq!(wb3.sheets[0].worksheet.comments.len(), 2);
    assert_eq!(wb3.sheets[0].worksheet.comments[0].cell_ref, "A1");
    assert_eq!(wb3.sheets[0].worksheet.comments[0].author, "Alice");
    assert_eq!(
        wb3.sheets[0].worksheet.comments[0].text,
        "This is the answer"
    );
    assert_eq!(wb3.sheets[0].worksheet.comments[1].cell_ref, "B1");
    assert_eq!(wb3.sheets[0].worksheet.comments[1].author, "Bob");
    assert_eq!(wb3.sheets[0].worksheet.comments[1].text, "Almost 100");
}

#[test]
fn test_comments_empty_no_crash() {
    // Workbook with no comments should parse and roundtrip fine.
    let xlsx_bytes = build_minimal_xlsx();
    let wb = read_xlsx(&xlsx_bytes).expect("read_xlsx should succeed");

    // No comments should be present.
    assert!(wb.sheets[0].worksheet.comments.is_empty());

    // Roundtrip should work.
    let xlsx2 = crate::writer::write_xlsx(&wb).expect("write_xlsx should succeed");
    let wb2 = read_xlsx(&xlsx2).expect("read_xlsx should succeed");
    assert!(wb2.sheets[0].worksheet.comments.is_empty());
}

#[test]
fn test_comments_multi_sheet_roundtrip() {
    use crate::ooxml::comments::Comment;

    // Build a workbook with 2 sheets, only the second has comments.
    let wb1 = WorkbookData {
        sheets: vec![
            SheetData {
                name: "NoComments".to_string(),
                worksheet: WorksheetXml {
                    ..Default::default()
                },
                ..Default::default()
            },
            SheetData {
                name: "WithComments".to_string(),
                worksheet: WorksheetXml {
                    comments: vec![Comment {
                        cell_ref: "C3".to_string(),
                        author: "Charlie".to_string(),
                        text: "Note on C3".to_string(),
                    }],
                    ..Default::default()
                },
                ..Default::default()
            },
        ],
        ..Default::default()
    };

    let xlsx1 = crate::writer::write_xlsx(&wb1).expect("write_xlsx should succeed");
    let wb2 = read_xlsx(&xlsx1).expect("read_xlsx should succeed");

    // First sheet has no comments.
    assert!(wb2.sheets[0].worksheet.comments.is_empty());

    // Second sheet has one comment.
    assert_eq!(wb2.sheets[1].worksheet.comments.len(), 1);
    assert_eq!(wb2.sheets[1].worksheet.comments[0].cell_ref, "C3");
    assert_eq!(wb2.sheets[1].worksheet.comments[0].author, "Charlie");
    assert_eq!(wb2.sheets[1].worksheet.comments[0].text, "Note on C3");
}

/// Rebuilds `xlsx` with `part` set to `data` (added when missing).
fn with_part(xlsx: &[u8], part: &str, data: &[u8]) -> Vec<u8> {
    let mut entries: Vec<ZipEntry> = read_zip_entries(xlsx, &ZipSecurityLimits::default())
        .expect("fixture should be readable")
        .into_iter()
        .filter(|(name, _)| name != part)
        .map(|(name, data)| ZipEntry { name, data })
        .collect();
    entries.sort_by(|a, b| a.name.cmp(&b.name));
    entries.push(ZipEntry {
        name: part.to_string(),
        data: data.to_vec(),
    });
    write_zip(&entries).expect("fixture should be writable")
}

#[test]
fn test_empty_custom_part_is_preserved_through_roundtrip() {
    let xlsx = with_part(&build_minimal_xlsx(), "customXml/item1.xml", b"");

    let wb = read_xlsx(&xlsx).expect("read should succeed");
    assert_eq!(
        wb.preserved_entries.get("customXml/item1.xml"),
        Some(&Vec::new())
    );

    let written = crate::writer::write_xlsx(&wb).expect("write should succeed");
    let wb2 = read_xlsx(&written).expect("re-read should succeed");
    assert_eq!(
        wb2.preserved_entries.get("customXml/item1.xml"),
        Some(&Vec::new())
    );
}

#[test]
fn test_empty_styles_part_reads() {
    let xlsx = with_part(&build_minimal_xlsx(), "xl/styles.xml", b"");
    read_xlsx(&xlsx).expect("an empty styles part must still read");
}

#[test]
fn test_empty_workbook_part_is_missing_part() {
    let xlsx = with_part(&build_minimal_xlsx(), "xl/workbook.xml", b"");
    let err = read_xlsx(&xlsx).expect_err("an empty workbook part must be rejected");
    assert_eq!(err.code(), "MISSING_PART");
    assert!(err.to_string().contains("is empty"), "got: {err}");
}

/// Deserializes a test fixture value from JSON.
fn from_json<T: serde::de::DeserializeOwned>(v: serde_json::Value) -> T {
    serde_json::from_value(v).expect("fixture JSON should match the model")
}

/// A `write_xlsx` package whose one sheet holds a chart, two pivot tables over
/// two pivot caches (each with records) and a slicer with its cache.
fn build_optional_parts_xlsx() -> Vec<u8> {
    use crate::ooxml::charts::{ChartData, ChartSeries};
    use serde_json::json;

    let chart = ChartData {
        series: vec![ChartSeries {
            val_ref: "Sheet1!$B$2:$B$5".into(),
            ..Default::default()
        }],
        ..Default::default()
    };
    let wb = WorkbookData {
        sheets: vec![SheetData {
            name: "Sheet1".into(),
            worksheet: WorksheetXml {
                charts: from_json(json!([{
                    "chart": chart,
                    "anchor": { "fromCol": 0, "fromRow": 0, "toCol": 5, "toRow": 10 },
                }])),
                pivot_tables: from_json(json!([
                    { "name": "PT1", "location": { "ref": "D1:E3" }, "cacheId": 0 },
                    { "name": "PT2", "location": { "ref": "G1:H3" }, "cacheId": 1 },
                ])),
                slicers: from_json(json!([{ "name": "Slicer_A", "cacheName": "Slicer_A" }])),
                ..Default::default()
            },
            ..Default::default()
        }],
        pivot_caches: from_json(json!([
            { "source": { "ref": "A1:B5", "sheet": "Sheet1" }, "fields": [{ "name": "A" }] },
            { "source": { "ref": "A1:B5", "sheet": "Sheet1" }, "fields": [{ "name": "B" }] },
        ])),
        pivot_cache_records: from_json(json!([
            { "records": [[{ "type": "number", "v": 1.0 }]] },
            { "records": [[{ "type": "number", "v": 2.0 }]] },
        ])),
        slicer_caches: from_json(json!([{ "name": "Slicer_A", "sourceName": "A" }])),
        ..Default::default()
    };
    crate::writer::write_xlsx(&wb).expect("fixture should be writable")
}

/// The ZIP entries of `xlsx` whose path starts with one of `prefixes`.
fn entries_under(xlsx: &[u8], prefixes: &[&str]) -> BTreeMap<String, Vec<u8>> {
    read_zip_entries(xlsx, &ZipSecurityLimits::default())
        .expect("fixture should be readable")
        .into_iter()
        .filter(|(path, _)| prefixes.iter().any(|p| path.starts_with(p)))
        .collect()
}

#[test]
fn test_optional_parts_fixture_reads_into_the_model() {
    let wb = read_xlsx(&build_optional_parts_xlsx()).expect("fixture should read");
    let ws = &wb.sheets[0].worksheet;
    assert_eq!(ws.charts.len(), 1);
    assert_eq!(ws.pivot_tables.len(), 2);
    assert_eq!(ws.slicers.len(), 1);
    assert_eq!(wb.pivot_caches.len(), 2);
    assert_eq!(wb.pivot_cache_records.len(), 2);
    assert_eq!(wb.slicer_caches.len(), 1);
    assert!(
        !wb.preserved_entries
            .keys()
            .any(|p| p.starts_with("xl/charts/")
                || p.starts_with("xl/drawings/")
                || p.starts_with("xl/pivot")
                || p.starts_with("xl/slicer")),
        "a well-formed optional part must be consumed: {:?}",
        wb.preserved_entries.keys().collect::<Vec<_>>()
    );
}

#[test]
fn test_malformed_chart_is_preserved_not_fatal() {
    let xlsx = with_part(
        &build_optional_parts_xlsx(),
        "xl/charts/chart1.xml",
        b"<c:chartSpace><",
    );

    let wb = read_xlsx(&xlsx).expect("a malformed chart must not fail the read");
    assert!(wb.sheets[0].worksheet.charts.is_empty());
    for part in [
        "xl/charts/chart1.xml",
        "xl/drawings/drawing1.xml",
        "xl/drawings/_rels/drawing1.xml.rels",
    ] {
        assert!(
            wb.preserved_entries.contains_key(part),
            "{part} should be preserved: {:?}",
            wb.preserved_entries.keys().collect::<Vec<_>>()
        );
    }
    assert_eq!(
        wb.preserved_entries["xl/charts/chart1.xml"],
        b"<c:chartSpace><".to_vec()
    );
    // The other optional parts still enter the model.
    assert_eq!(wb.sheets[0].worksheet.pivot_tables.len(), 2);
    assert_eq!(wb.sheets[0].worksheet.slicers.len(), 1);

    let json: serde_json::Value =
        serde_json::from_str(&read_xlsx_json(&xlsx).expect("the JSON read must succeed too"))
            .unwrap();
    assert!(json["sheets"][0]["worksheet"].get("charts").is_none());
    assert!(json["preservedEntries"]["xl/charts/chart1.xml"].is_array());
    assert!(json["preservedEntries"]["xl/drawings/drawing1.xml"].is_array());
}

#[test]
fn test_malformed_pivot_cache_keeps_the_whole_pivot_family_out() {
    let original = build_optional_parts_xlsx();
    let xlsx = with_part(
        &original,
        "xl/pivotCache/pivotCacheDefinition1.xml",
        b"<pivotCacheDefinition><a></b>",
    );
    let pivot_parts = entries_under(&xlsx, &["xl/pivotCache/", "xl/pivotTables/"]);
    assert!(pivot_parts.contains_key("xl/pivotCache/pivotCacheDefinition2.xml"));
    assert!(pivot_parts.contains_key("xl/pivotCache/pivotCacheRecords2.xml"));
    assert!(pivot_parts.contains_key("xl/pivotTables/pivotTable2.xml"));

    let wb = read_xlsx(&xlsx).expect("a malformed pivot cache must not fail the read");
    assert!(wb.pivot_caches.is_empty());
    assert!(wb.pivot_cache_records.is_empty());
    assert!(wb.sheets[0].worksheet.pivot_tables.is_empty());
    for (path, data) in &pivot_parts {
        assert_eq!(
            wb.preserved_entries.get(path),
            Some(data),
            "{path} should be preserved verbatim"
        );
    }
    // The other optional parts still enter the model.
    assert_eq!(wb.sheets[0].worksheet.charts.len(), 1);
    assert_eq!(wb.slicer_caches.len(), 1);

    let json: serde_json::Value =
        serde_json::from_str(&read_xlsx_json(&xlsx).expect("the JSON read must succeed too"))
            .unwrap();
    assert!(json.get("pivotCaches").is_none());
    assert!(json.get("pivotCacheRecords").is_none());
    assert!(json["sheets"][0]["worksheet"].get("pivotTables").is_none());
    for path in pivot_parts.keys() {
        assert!(json["preservedEntries"][path].is_array(), "{path}");
    }

    let written = crate::writer::write_xlsx(&wb).expect("the read result must be writable");
    let written_parts = entries_under(&written, &["xl/pivotCache/", "xl/pivotTables/"]);
    assert_eq!(written_parts, pivot_parts);
}

#[test]
fn test_malformed_slicer_is_preserved_not_fatal() {
    let bad = b"<slicers><a></b></slicers>";
    let xlsx = with_part(&build_optional_parts_xlsx(), "xl/slicers/slicer1.xml", bad);

    let wb = read_xlsx(&xlsx).expect("a malformed slicer part must not fail the read");
    assert!(wb.sheets[0].worksheet.slicers.is_empty());
    assert_eq!(
        wb.preserved_entries.get("xl/slicers/slicer1.xml"),
        Some(&bad.to_vec())
    );
    assert_eq!(wb.slicer_caches.len(), 1);

    let json: serde_json::Value =
        serde_json::from_str(&read_xlsx_json(&xlsx).expect("the JSON read must succeed too"))
            .unwrap();
    assert!(json["sheets"][0]["worksheet"].get("slicers").is_none());
    assert!(json["preservedEntries"]["xl/slicers/slicer1.xml"].is_array());
}

#[test]
fn test_optional_part_policy() {
    let mut log = LogDigest::default();
    let ok: Result<u8> = Ok(7);
    assert_eq!(optional_part(&mut log, "p", ok).unwrap(), Some(7));
    let malformed: Result<u8> = Err(ModernXlsxError::XmlParse("bad".into()));
    assert_eq!(optional_part(&mut log, "p", malformed).unwrap(), None);
    let hostile: Result<u8> = Err(ModernXlsxError::Security("bomb".into()));
    assert_eq!(
        optional_part(&mut log, "p", hostile).unwrap_err().code(),
        "SECURITY",
        "a security error must stay fatal"
    );
    // Only the malformed part is logged, once, with its path and error.
    assert_eq!(log.entries.len(), 1);
    assert_eq!(log.entries[0].count, 1);
    assert_eq!(log.entries[0].first, r#""p: XML parse error: bad""#);
}

/// A malformed `[Content_Types].xml` is never preserved (the writer generates
/// one), so the read warns that it was dropped, not that it was kept.
#[test]
fn malformed_content_types_is_logged_as_dropped_not_preserved() {
    let xlsx = with_part(
        &build_optional_parts_xlsx(),
        "[Content_Types].xml",
        b"<Types><a></b></Types>",
    );
    let ctx = parse_common(&xlsx, &ZipSecurityLimits::default(), None).unwrap();
    assert!(ctx.content_types.is_none());
    let dropped: Vec<_> = ctx
        .log
        .entries
        .iter()
        .filter(|e| e.kind == "malformed content types part dropped")
        .collect();
    assert_eq!(dropped.len(), 1);
    assert_eq!(dropped[0].count, 1);
    assert_eq!(digest_count(&ctx, "malformed part preserved verbatim"), 0);

    let wb = read_xlsx(&xlsx).unwrap();
    assert!(wb.preserved_content_types.is_empty());
    let json: serde_json::Value = serde_json::from_str(&read_xlsx_json(&xlsx).unwrap()).unwrap();
    assert!(json.get("preservedContentTypes").is_none());
}

#[test]
fn test_log_digest_counts_per_kind_and_keeps_the_first_value() {
    let mut log = LogDigest::default();
    log.note(Level::Warn, "a", "first");
    log.note(Level::Debug, "b", "other");
    log.note(Level::Warn, "a", "second");
    let summary: Vec<_> = log
        .entries
        .iter()
        .map(|e| (e.level, e.kind, e.count, e.first.as_str()))
        .collect();
    assert_eq!(
        summary,
        [
            (Level::Warn, "a", 2, r#""first""#),
            (Level::Debug, "b", 1, r#""other""#)
        ]
    );
}

#[test]
fn test_log_text_escapes_and_truncates() {
    assert_eq!(log_text("a\nb\u{1b}c\"d"), r#""a\nb\u{1b}c\"d""#);
    assert_eq!(log_text("line\u{2028}sep"), r#""line\u{2028}sep""#);

    let long = "x".repeat(1000);
    let logged = log_text(&long);
    assert_eq!(
        logged,
        format!("\"{}\" (truncated, 1000 bytes)", "x".repeat(LOG_TEXT_LIMIT))
    );

    // An escape that would cross the limit is dropped whole, never cut.
    let edge = format!("{}\u{1b}", "x".repeat(LOG_TEXT_LIMIT - 2));
    let logged = log_text(&edge);
    assert!(logged.starts_with(&format!("\"{}\"", "x".repeat(LOG_TEXT_LIMIT - 2))));
    assert!(logged.ends_with(" (truncated, 159 bytes)"), "{logged}");
}

#[test]
fn test_json_capacity_hint_is_bounded() {
    assert_eq!(json_capacity_hint(0), 4096);
    assert_eq!(json_capacity_hint(1_000), 15_000);
    assert_eq!(json_capacity_hint(usize::MAX), JSON_PREALLOC_CAP);
}

#[test]
fn test_json_length_limit_gives_size_limit() {
    let xlsx = build_minimal_xlsx();
    let tight = ZipSecurityLimits {
        max_json_length: 64,
        ..Default::default()
    };
    let err = read_xlsx_json_with_options(&xlsx, &tight).expect_err("limit must trip");
    assert_eq!(err.code(), "SIZE_LIMIT");
    assert!(err.to_string().contains("64-byte limit"), "got: {err}");
    assert!(read_xlsx_json_with_options(&xlsx, &ZipSecurityLimits::default()).is_ok());
}

#[test]
fn test_json_length_limit_on_password_path() {
    let xlsx = build_minimal_xlsx();
    let tight = ZipSecurityLimits {
        max_json_length: 64,
        ..Default::default()
    };
    let err =
        read_xlsx_json_with_password_and_options(&xlsx, "pw", &tight).expect_err("limit must trip");
    assert_eq!(err.code(), "SIZE_LIMIT");
    assert_eq!(
        read_xlsx_json_with_password(&xlsx, "pw").unwrap(),
        read_xlsx_json(&xlsx).unwrap()
    );
}

/// [`build_minimal_xlsx`] with its shared strings and worksheet replaced
/// by the given XML.
fn build_xlsx_with_parts(sst_xml: &str, sheet_xml: &str) -> Vec<u8> {
    let mut parts = read_zip_entries(&build_minimal_xlsx(), &ZipSecurityLimits::default()).unwrap();
    parts.insert("xl/sharedStrings.xml".into(), sst_xml.as_bytes().to_vec());
    parts.insert(
        "xl/worksheets/sheet1.xml".into(),
        sheet_xml.as_bytes().to_vec(),
    );
    let mut entries: Vec<ZipEntry> = parts
        .into_iter()
        .map(|(name, data)| ZipEntry { name, data })
        .collect();
    entries.sort_by(|a, b| a.name.cmp(&b.name));
    write_zip(&entries).unwrap()
}

#[test]
fn rich_runs_attach_by_sst_index_not_by_text() {
    let sst = r#"<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="2" uniqueCount="2"><si><t>A</t></si><si><r><rPr><b/></rPr><t>A</t></r></si></sst>"#;
    let sheet = r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>1</v></c></row></sheetData></worksheet>"#;
    let xlsx = build_xlsx_with_parts(sst, sheet);

    let wb = read_xlsx(&xlsx).unwrap();
    let from_json: WorkbookData = serde_json::from_str(&read_xlsx_json(&xlsx).unwrap()).unwrap();
    for model in [&wb, &from_json] {
        let cells = &model.sheets[0].worksheet.rows[0].cells;
        assert_eq!(cells[0].value.as_deref(), Some("A"));
        assert!(
            cells[0].rich_text.is_none(),
            "plain entry got the rich entry's runs"
        );
        assert_eq!(cells[1].value.as_deref(), Some("A"));
        let runs = cells[1]
            .rich_text
            .as_ref()
            .expect("rich entry lost its runs");
        assert_eq!(runs.len(), 1);
        assert_eq!(runs[0].text, "A");
        assert_eq!(runs[0].bold, Some(true));
    }
}

#[test]
fn json_with_shared_and_array_formulas_deserializes() {
    let sheet = r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1"><f t="shared" ref="A1:A2" si="0">B1+1</f><v>2</v></c><c r="B1"><v>1</v></c><c r="C1"><f t="array" ref="C1:C2">B1:B2*2</f><v>2</v></c></row><row r="2"><c r="A2"><f t="shared" si="0"/><v>3</v></c><c r="B2"><v>2</v></c></row></sheetData></worksheet>"#;
    let sst = r#"<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="0" uniqueCount="0"/>"#;
    let xlsx = build_xlsx_with_parts(sst, sheet);

    let wb: WorkbookData = serde_json::from_str(&read_xlsx_json(&xlsx).unwrap())
        .expect("read_xlsx_json output must deserialize into WorkbookData");
    let row1 = &wb.sheets[0].worksheet.rows[0].cells;
    assert_eq!(row1[0].formula_type.as_deref(), Some("shared"));
    assert_eq!(row1[0].formula_ref.as_deref(), Some("A1:A2"));
    assert_eq!(row1[0].shared_index, Some(0));
    assert_eq!(row1[2].formula_type.as_deref(), Some("array"));
    assert_eq!(row1[2].formula_ref.as_deref(), Some("C1:C2"));
    let a2 = &wb.sheets[0].worksheet.rows[1].cells[0];
    assert_eq!(a2.formula_type.as_deref(), Some("shared"));
    assert_eq!(a2.shared_index, Some(0));
}

// -----------------------------------------------------------------------
// Preserved worksheet <ext>, rich inline strings, shared formulas and
// t="d" serials
// -----------------------------------------------------------------------

const EMPTY_SST: &str = r#"<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="0" uniqueCount="0"/>"#;

/// A worksheet with the `x14` and `xm` prefixes bound on its root, a rich
/// inline string in A1 ("Hello " then bold "World"), a shared formula in
/// A2 and an ISO 8601 `t="d"` date in A3.
const BUG10_FIXTURE_SHEET: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main"><sheetData><row r="1"><c r="A1" t="inlineStr"><is><r><t xml:space="preserve">Hello </t></r><r><rPr><b/></rPr><t>World</t></r></is></c></row><row r="2"><c r="A2"><f t="shared" ref="A2:A3" si="0">B2*2</f><v>2</v></c></row><row r="3"><c r="A3" t="d"><v>2026-10-30T00:00:00</v></c></row></sheetData></worksheet>"#;

/// The x14 conditional-formatting `<ext>` of [`bug10_full_sheet`] as it
/// is kept: the `x14` and `xm` prefixes bound on the `<worksheet>` root
/// re-declared on `<ext>`.
const X14_CF_EXT: &str = r#"<ext xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main" uri="{78C0D931-6437-407d-A8EE-F0AAD7539E65}"><x14:conditionalFormattings><x14:conditionalFormatting><x14:cfRule type="dataBar" id="{1E5A0F2C-3B4D-4E6F-8A9B-0C1D2E3F4A5B}"><x14:dataBar minLength="0" maxLength="100"><x14:cfvo type="autoMin"/><x14:cfvo type="autoMax"/></x14:dataBar></x14:cfRule><xm:sqref>A1:A3</xm:sqref></x14:conditionalFormatting></x14:conditionalFormattings></ext>"#;

/// [`BUG10_FIXTURE_SHEET`] plus an `<rPh>` run, a theme-coloured data bar
/// and an x14 `<ext>` that uses the prefixes bound on the root.
fn bug10_full_sheet() -> String {
    BUG10_FIXTURE_SHEET
        .replace(
            "<t>World</t></r></is>",
            r#"<t>World</t></r><rPh sb="0" eb="1"><t>FURI</t></rPh></is>"#,
        )
        .replace(
            "</sheetData></worksheet>",
            concat!(
                r#"</sheetData><conditionalFormatting sqref="A1:A3"><cfRule type="dataBar" priority="1"><dataBar><cfvo type="min"/><cfvo type="max"/><color theme="5" tint="-0.25"/></dataBar></cfRule></conditionalFormatting>"#,
                r#"<extLst><ext uri="{78C0D931-6437-407d-A8EE-F0AAD7539E65}"><x14:conditionalFormattings><x14:conditionalFormatting><x14:cfRule type="dataBar" id="{1E5A0F2C-3B4D-4E6F-8A9B-0C1D2E3F4A5B}"><x14:dataBar minLength="0" maxLength="100"><x14:cfvo type="autoMin"/><x14:cfvo type="autoMax"/></x14:dataBar></x14:cfRule><xm:sqref>A1:A3</xm:sqref></x14:conditionalFormatting></x14:conditionalFormattings></ext></extLst></worksheet>"#,
            ),
        )
}

#[test]
fn bug10_fixture_json_deserializes_into_workbook_data() {
    for sheet in [BUG10_FIXTURE_SHEET.to_owned(), bug10_full_sheet()] {
        let xlsx = build_xlsx_with_parts(EMPTY_SST, &sheet);
        let json = read_xlsx_json(&xlsx).unwrap();
        let wb: WorkbookData = serde_json::from_str(&json)
            .unwrap_or_else(|e| panic!("read_xlsx_json output must deserialize: {e}\n{json}"));
        let ws = &wb.sheets[0].worksheet;
        assert_eq!(ws.rows[0].cells[0].value.as_deref(), Some("Hello World"));
        assert_eq!(ws.rows[1].cells[0].formula_type.as_deref(), Some("shared"));
        // 2026-10-30 is serial 46325 in the 1900 date system.
        let a3 = &ws.rows[2].cells[0];
        assert_eq!(
            (a3.cell_type, a3.value.as_deref()),
            (CellType::Number, Some("46325"))
        );
    }
}

#[test]
fn bug10_full_sheet_round_trips_through_write_xlsx() {
    let xlsx = build_xlsx_with_parts(EMPTY_SST, &bug10_full_sheet());
    let from_json: WorkbookData = serde_json::from_str(&read_xlsx_json(&xlsx).unwrap()).unwrap();
    let wb = read_xlsx(&xlsx).unwrap();
    for model in [&wb, &from_json] {
        let ws = &model.sheets[0].worksheet;
        assert_eq!(ws.preserved_extensions, [X14_CF_EXT]);
        let db = ws.conditional_formatting[0].rules[0]
            .data_bar
            .as_ref()
            .unwrap();
        let spec = db.color_spec.as_ref().unwrap();
        assert_eq!((spec.theme, spec.tint), (Some(5), Some(-0.25)));
    }

    // read_xlsx -> write_xlsx -> read_xlsx keeps the ext, once, inside <extLst>.
    let written = crate::writer::write_xlsx(&wb).unwrap();
    let parts = read_zip_entries(&written, &ZipSecurityLimits::default()).unwrap();
    let sheet = std::str::from_utf8(&parts["xl/worksheets/sheet1.xml"]).unwrap();
    assert_eq!(sheet.matches(X14_CF_EXT).count(), 1, "{sheet}");
    let ext_at = sheet.find(X14_CF_EXT).unwrap();
    assert!(sheet.find("<extLst>").unwrap() < ext_at, "{sheet}");
    assert!(ext_at < sheet.find("</extLst>").unwrap(), "{sheet}");
    for absent in [r#"rgb="""#, "<f cm", "dtr1", "2026-10-30"] {
        assert!(!sheet.contains(absent), "{absent} in {sheet}");
    }
    assert!(sheet.contains(r#"<c r="A3"><v>46325</v></c>"#), "{sheet}");
    assert!(
        sheet.contains(r#"<color theme="5" tint="-0.25"/>"#),
        "{sheet}"
    );

    let again = read_xlsx(&written).unwrap();
    assert_eq!(again.sheets[0].worksheet.preserved_extensions, [X14_CF_EXT]);
    let json = read_xlsx_json(&written).unwrap();
    let from_json: WorkbookData = serde_json::from_str(&json).unwrap();
    assert_eq!(
        from_json.sheets[0].worksheet.preserved_extensions,
        [X14_CF_EXT]
    );
}

#[test]
fn iso_date_cells_use_the_workbook_date_system() {
    let sheet = r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1" t="d"><v>2024-01-15T12:00:00</v></c></row></sheetData></worksheet>"#;
    let mut parts = read_zip_entries(
        &build_xlsx_with_parts(EMPTY_SST, sheet),
        &ZipSecurityLimits::default(),
    )
    .unwrap();
    let mut wb_xml = WorkbookXml::parse(&parts["xl/workbook.xml"]).unwrap();
    wb_xml.date_system = DateSystem::Date1904;
    parts.insert("xl/workbook.xml".into(), wb_xml.to_xml().unwrap());
    let mut entries: Vec<ZipEntry> = parts
        .into_iter()
        .map(|(name, data)| ZipEntry { name, data })
        .collect();
    entries.sort_by(|a, b| a.name.cmp(&b.name));
    let xlsx = write_zip(&entries).unwrap();

    let wb = read_xlsx(&xlsx).unwrap();
    let from_json: WorkbookData = serde_json::from_str(&read_xlsx_json(&xlsx).unwrap()).unwrap();
    for model in [&wb, &from_json] {
        let a1 = &model.sheets[0].worksheet.rows[0].cells[0];
        // 45306.5 in the 1900 system, less the 1462 days between epochs.
        assert_eq!(a1.value.as_deref(), Some("43844.5"));
        assert_eq!(a1.cell_type, CellType::Number);
    }
}

// -----------------------------------------------------------------------
// Drawings: only a chart-only drawing is consumed
// -----------------------------------------------------------------------

/// One sheet whose drawing holds a picture, stored in `preserved_entries`
/// the way `addImage` stores it, plus `charts` in the model.
fn picture_workbook(charts: Vec<charts::WorksheetChart>) -> WorkbookData {
    let mut wb = WorkbookData {
        sheets: vec![SheetData {
            name: "Sheet1".into(),
            worksheet: WorksheetXml {
                charts,
                ..Default::default()
            },
            ..Default::default()
        }],
        ..Default::default()
    };
    let drawing = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><xdr:twoCellAnchor><xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>2</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>2</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:pic><xdr:nvPicPr><xdr:cNvPr id="9" name="Img"/><xdr:cNvPicPr/></xdr:nvPicPr><xdr:blipFill><a:blip r:embed="rId1"/></xdr:blipFill><xdr:spPr/></xdr:pic><xdr:clientData/></xdr:twoCellAnchor></xdr:wsDr>"#;
    let drawing_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/></Relationships>"#;
    let sheet_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdDrawing1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/></Relationships>"#;
    for (path, data) in [
        ("xl/drawings/drawing1.xml", drawing.as_bytes()),
        (
            "xl/drawings/_rels/drawing1.xml.rels",
            drawing_rels.as_bytes(),
        ),
        ("xl/worksheets/_rels/sheet1.xml.rels", sheet_rels.as_bytes()),
        ("xl/media/image1.png", b"\x89PNG".as_slice()),
    ] {
        wb.preserved_entries.insert(path.into(), data.to_vec());
    }
    wb
}

fn line_chart() -> charts::WorksheetChart {
    charts::WorksheetChart {
        chart: charts::ChartData {
            chart_type: charts::ChartType::Line,
            series: vec![charts::ChartSeries {
                val_ref: "Sheet1!$A$1:$A$3".into(),
                ..Default::default()
            }],
            ..Default::default()
        },
        anchor: charts::ChartAnchor {
            from_col: 4,
            to_col: 9,
            to_row: 10,
            ..Default::default()
        },
    }
}

/// Asserts every internal target of `<source>`'s .rels (when it has one) is
/// an entry of `entries`.
fn assert_targets_exist(entries: &BTreeMap<String, Vec<u8>>, source: &str) {
    let Some(rels) = entries.get(&rels_path(source)) else {
        return;
    };
    for rel in &Relationships::parse(rels).unwrap() {
        let target = resolve_target(source, &rel.target);
        assert!(
            entries.contains_key(&target),
            "{source}: {} targets missing {target}",
            rel.id
        );
    }
}

fn written_entries(bytes: &[u8]) -> BTreeMap<String, Vec<u8>> {
    read_zip_entries(bytes, &ZipSecurityLimits::default())
        .unwrap()
        .into_iter()
        .collect()
}

#[test]
fn mixed_picture_and_chart_drawing_survives_two_writes() {
    let first = crate::writer::write_xlsx(&picture_workbook(vec![line_chart()])).unwrap();

    let wb = read_xlsx(&first).unwrap();
    let from_json: WorkbookData = serde_json::from_str(&read_xlsx_json(&first).unwrap()).unwrap();
    for model in [&wb, &from_json] {
        assert_eq!(model.sheets[0].worksheet.charts.len(), 1);
        for kept in [
            "xl/drawings/drawing1.xml",
            "xl/drawings/_rels/drawing1.xml.rels",
        ] {
            assert!(model.preserved_entries.contains_key(kept), "{kept}");
        }
        assert!(
            !model
                .preserved_entries
                .keys()
                .any(|p| p.starts_with("xl/charts/")),
            "the chart part enters the model"
        );
    }

    let second = crate::writer::write_xlsx(&wb).unwrap();
    let third = crate::writer::write_xlsx(&read_xlsx(&second).unwrap()).unwrap();
    for bytes in [&second, &third] {
        let entries = written_entries(bytes);
        assert_targets_exist(&entries, "xl/worksheets/sheet1.xml");
        assert_targets_exist(&entries, "xl/drawings/drawing1.xml");
        let drawing = std::str::from_utf8(&entries["xl/drawings/drawing1.xml"]).unwrap();
        assert_eq!(drawing.matches("<xdr:pic>").count(), 1, "{drawing}");
        assert_eq!(drawing.matches("<xdr:graphicFrame").count(), 1, "{drawing}");
    }
}

#[test]
fn image_only_drawing_survives_two_writes() {
    let first = crate::writer::write_xlsx(&picture_workbook(Vec::new())).unwrap();

    let wb = read_xlsx(&first).unwrap();
    assert!(wb.sheets[0].worksheet.charts.is_empty());
    assert!(
        wb.preserved_entries
            .contains_key("xl/drawings/drawing1.xml")
    );

    let entries = written_entries(&crate::writer::write_xlsx(&wb).unwrap());
    assert_targets_exist(&entries, "xl/worksheets/sheet1.xml");
    assert_targets_exist(&entries, "xl/drawings/drawing1.xml");
    let drawing = std::str::from_utf8(&entries["xl/drawings/drawing1.xml"]).unwrap();
    assert!(drawing.contains("<xdr:pic>"), "{drawing}");
}

#[test]
fn chart_only_drawing_is_consumed() {
    let first = crate::writer::write_xlsx(&WorkbookData {
        sheets: vec![SheetData {
            name: "Sheet1".into(),
            worksheet: WorksheetXml {
                charts: vec![line_chart()],
                ..Default::default()
            },
            ..Default::default()
        }],
        ..Default::default()
    })
    .unwrap();
    let wb = read_xlsx(&first).unwrap();
    assert_eq!(wb.sheets[0].worksheet.charts.len(), 1);
    assert!(
        !wb.preserved_entries
            .keys()
            .any(|p| p.starts_with("xl/drawings/") || p.starts_with("xl/charts/")),
        "{:?}",
        wb.preserved_entries.keys().collect::<Vec<_>>()
    );
}

/// A workbook carrying defined names, `calcPr`, `fileVersion`, `workbookPr`
/// and the 1904 date system.
fn workbook_props_fixture() -> WorkbookData {
    use crate::ooxml::workbook::{CalcPr, DefinedName, FileVersion, WorkbookPr};
    // R&D!A1, Two!A2 and Two!A3 hold formulas, so calcChain entries naming
    // them survive a write.
    let sheet = |name: &str, refs: &[&str]| SheetData {
        name: name.into(),
        worksheet: WorksheetXml {
            rows: refs
                .iter()
                .map(|&reference| Row {
                    index: reference[1..].parse().unwrap(),
                    cells: vec![Cell {
                        reference: reference.into(),
                        formula: Some("1+0".into()),
                        ..Default::default()
                    }],
                    ..Default::default()
                })
                .collect(),
            ..Default::default()
        },
        ..Default::default()
    };
    WorkbookData {
        sheets: vec![sheet("R&D", &["A1"]), sheet("Two", &["A2", "A3"])],
        date_system: DateSystem::Date1904,
        defined_names: vec![
            DefinedName {
                name: "F".into(),
                value: r#"IF(A1<1,"a & b",0)"#.into(),
                ..Default::default()
            },
            DefinedName {
                name: "_xlnm._FilterDatabase".into(),
                value: "'R&D'!$A$1:$C$9".into(),
                sheet_id: Some(0),
                hidden: true,
                ..Default::default()
            },
        ],
        calc_pr: Some(CalcPr {
            calc_id: Some(191_029),
            full_calc_on_load: Some(true),
            iterate_delta: Some(0.001),
            ..Default::default()
        }),
        file_version: Some(FileVersion {
            app_name: Some("xl".into()),
            code_name: Some("{ABC}".into()),
            ..Default::default()
        }),
        workbook_pr: Some(WorkbookPr {
            code_name: Some("ThisWorkbook".into()),
            default_theme_version: Some(166_925),
            ..Default::default()
        }),
        ..Default::default()
    }
}

#[test]
fn workbook_props_survive_struct_and_json_reads() {
    let wb = workbook_props_fixture();
    let xlsx = crate::writer::write_xlsx(&wb).unwrap();
    let back = read_xlsx(&xlsx).unwrap();
    assert_eq!(back.defined_names, wb.defined_names);
    assert_eq!(back.calc_pr, wb.calc_pr);
    assert_eq!(back.file_version, wb.file_version);
    assert_eq!(back.workbook_pr, wb.workbook_pr);
    assert_eq!(back.date_system, DateSystem::Date1904);
    assert_eq!(back.sheets[0].name, "R&D");

    // The hand-spliced JSON read agrees with serde's output field by field.
    let json: serde_json::Value = serde_json::from_str(&read_xlsx_json(&xlsx).unwrap()).unwrap();
    let expected = serde_json::to_value(&back).unwrap();
    for key in [
        "calcPr",
        "fileVersion",
        "workbookPr",
        "definedNames",
        "dateSystem",
    ] {
        assert_eq!(json[key], expected[key], "{key}");
    }

    // JSON read -> write -> read: nothing is lost, and the sheet name is
    // stable over repeated trips.
    let from_json: WorkbookData = serde_json::from_value(json).unwrap();
    let again = read_xlsx(&crate::writer::write_xlsx(&from_json).unwrap()).unwrap();
    assert_eq!(again.calc_pr, wb.calc_pr);
    assert_eq!(again.workbook_pr, wb.workbook_pr);
    assert_eq!(again.defined_names, wb.defined_names);
    assert_eq!(again.sheets[0].name, "R&D");
}

/// Rewrite `xlsx` with `sheetId="2"` renamed to 3 and a calcChain added.
fn with_calc_chain_over_sheet_ids_1_and_3(xlsx: &[u8], chain: &str) -> Vec<u8> {
    let mut parts = read_zip_entries(xlsx, &ZipSecurityLimits::default()).unwrap();
    let wb = String::from_utf8(parts["xl/workbook.xml"].clone()).unwrap();
    assert!(wb.contains(r#"sheetId="2""#));
    parts.insert(
        "xl/workbook.xml".into(),
        wb.replace(r#"sheetId="2""#, r#"sheetId="3""#).into_bytes(),
    );
    parts.insert("xl/calcChain.xml".into(), chain.as_bytes().to_vec());
    let entries: Vec<ZipEntry> = parts
        .into_iter()
        .map(|(name, data)| ZipEntry { name, data })
        .collect();
    write_zip(&entries).unwrap()
}

#[test]
fn calc_chain_sheet_id_is_mapped_to_sheet_position() {
    let base = crate::writer::write_xlsx(&workbook_props_fixture()).unwrap();
    // i=3 is the second sheet; i=9 names no sheet and is dropped; the
    // entry without i inherits 9 and is dropped too.
    let chain = r#"<calcChain xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><c r="A1" i="1"/><c r="A2" i="3" l="1"/><c r="A3"/><c r="A4" i="9"/><c r="A5"/></calcChain>"#;
    let xlsx = with_calc_chain_over_sheet_ids_1_and_3(&base, chain);

    let wb = read_xlsx(&xlsx).unwrap();
    let got: Vec<(&str, u32, bool)> = wb
        .calc_chain
        .iter()
        .map(|e| (e.cell_ref.as_str(), e.sheet_id, e.l))
        .collect();
    assert_eq!(got, [("A1", 1, false), ("A2", 2, true), ("A3", 2, false)]);

    let json: WorkbookData = serde_json::from_str(&read_xlsx_json(&xlsx).unwrap()).unwrap();
    assert_eq!(json.calc_chain, wb.calc_chain);

    // The writer renumbers sheetId by position, so the positions it
    // writes into calcChain resolve on the way back in.
    let again = read_xlsx(&crate::writer::write_xlsx(&wb).unwrap()).unwrap();
    assert_eq!(again.calc_chain, wb.calc_chain);
}

#[test]
fn bad_calc_chain_does_not_fail_the_read() {
    let base = crate::writer::write_xlsx(&workbook_props_fixture()).unwrap();
    let chain = r#"<calcChain xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><c r="A1" i="-1"/></calcChain>"#;
    let xlsx = with_calc_chain_over_sheet_ids_1_and_3(&base, chain);

    let wb = read_xlsx(&xlsx).unwrap();
    assert!(wb.calc_chain.is_empty());
    let json: WorkbookData = serde_json::from_str(&read_xlsx_json(&xlsx).unwrap()).unwrap();
    assert!(json.calc_chain.is_empty());
}

#[test]
fn malformed_app_xml_keeps_the_core_properties() {
    let wb = WorkbookData {
        sheets: vec![SheetData {
            name: "Sheet1".into(),
            ..Default::default()
        }],
        doc_properties: Some(doc_props::DocProperties {
            title: Some("Quarterly".into()),
            creator: Some("Ann".into()),
            ..Default::default()
        }),
        ..Default::default()
    };
    let base = crate::writer::write_xlsx(&wb).unwrap();
    let xlsx = with_part(
        &base,
        "docProps/app.xml",
        b"<Properties><a></b></Properties>",
    );

    let back = read_xlsx(&xlsx).expect("a malformed app.xml must not fail the read");
    let props = back
        .doc_properties
        .expect("the core properties must survive a malformed app.xml");
    assert_eq!(props.title.as_deref(), Some("Quarterly"));
    assert_eq!(props.creator.as_deref(), Some("Ann"));
    assert_eq!(props.application, None);

    let json: serde_json::Value = serde_json::from_str(&read_xlsx_json(&xlsx).unwrap()).unwrap();
    assert_eq!(json["docProperties"]["title"], "Quarterly");
    assert_eq!(json["docProperties"]["creator"], "Ann");
    assert!(json["docProperties"].get("application").is_none());
}

#[test]
fn malformed_theme_is_preserved_not_fatal() {
    let bad = b"<a:theme><a></b></a:theme>";
    let xlsx = with_part(&build_minimal_xlsx(), THEME_PATH, bad);

    let wb = read_xlsx(&xlsx).expect("a malformed theme must not fail the read");
    assert!(wb.theme_colors.is_none());
    assert_eq!(wb.preserved_entries.get(THEME_PATH), Some(&bad.to_vec()));

    let json: serde_json::Value =
        serde_json::from_str(&read_xlsx_json(&xlsx).expect("the JSON read must succeed too"))
            .unwrap();
    assert!(json.get("themeColors").is_none());
    assert!(json["preservedEntries"][THEME_PATH].is_array());

    // The writer keeps the preserved theme instead of generating one.
    let again = read_xlsx(&crate::writer::write_xlsx(&wb).unwrap()).unwrap();
    assert!(again.theme_colors.is_none());
    assert_eq!(again.preserved_entries.get(THEME_PATH), Some(&bad.to_vec()));
}

#[test]
fn two_slicer_parts_on_one_sheet_both_enter_the_model() {
    const SLICER_NS: &str = "http://schemas.microsoft.com/office/spreadsheetml/2010/slicer";
    let base = build_optional_parts_xlsx();
    let rels_path = "xl/worksheets/_rels/sheet1.xml.rels";
    let rels_xml =
        String::from_utf8(entries_under(&base, &[rels_path])[rels_path].clone()).unwrap();
    let slicer1_target = Relationships::parse(rels_xml.as_bytes())
        .unwrap()
        .find_by_type(REL_SLICER)
        .map(|rel| rel.target.clone())
        .next()
        .expect("the fixture sheet has a slicer relationship");
    assert!(slicer1_target.ends_with("slicer1.xml"), "{slicer1_target}");
    let slicer2_target = slicer1_target.replace("slicer1.xml", "slicer2.xml");
    let rels_xml = rels_xml.replace(
        "</Relationships>",
        &format!(
            r#"<Relationship Id="rIdSlicer2" Type="{REL_SLICER}" Target="{slicer2_target}"/></Relationships>"#
        ),
    );

    let slicer = |name: &str| {
        format!(r#"<slicers xmlns="{SLICER_NS}"><slicer name="{name}" cache="{name}"/></slicers>"#)
    };
    let xlsx = with_part(&base, rels_path, rels_xml.as_bytes());
    let xlsx = with_part(
        &xlsx,
        "xl/slicers/slicer1.xml",
        slicer("Slicer_A").as_bytes(),
    );
    let xlsx = with_part(
        &xlsx,
        "xl/slicers/slicer2.xml",
        slicer("Slicer_B").as_bytes(),
    );

    let wb = read_xlsx(&xlsx).unwrap();
    let mut names: Vec<&str> = wb.sheets[0]
        .worksheet
        .slicers
        .iter()
        .map(|s| s.name.as_str())
        .collect();
    names.sort_unstable();
    assert_eq!(names, ["Slicer_A", "Slicer_B"]);
    assert!(
        !wb.preserved_entries
            .keys()
            .any(|p| p.starts_with("xl/slicers/")),
        "both slicer parts are in the model: {:?}",
        wb.preserved_entries.keys().collect::<Vec<_>>()
    );

    let json: serde_json::Value = serde_json::from_str(&read_xlsx_json(&xlsx).unwrap()).unwrap();
    assert_eq!(
        json["sheets"][0]["worksheet"]["slicers"]
            .as_array()
            .map(Vec::len),
        Some(2)
    );
}

#[test]
fn large_preserved_entry_trips_the_json_limit_before_serializing() {
    // Incompressible bytes (xorshift), so the ZIP ratio guard stays quiet.
    let mut state = 0x9E37_79B9_7F4A_7C15_u64;
    let blob: Vec<u8> = (0..1 << 20)
        .map(|_| {
            state ^= state << 13;
            state ^= state >> 7;
            state ^= state << 17;
            state.to_le_bytes()[0]
        })
        .collect();
    let xlsx = with_part(&build_minimal_xlsx(), "customXml/blob.bin", &blob);

    let tight = ZipSecurityLimits {
        max_json_length: 4096,
        ..Default::default()
    };
    let err = read_xlsx_json_with_options(&xlsx, &tight).expect_err("limit must trip");
    assert_eq!(err.code(), "SIZE_LIMIT");
    assert!(
        err.to_string().contains("would reach at least"),
        "got: {err}"
    );
    assert!(err.to_string().contains("4096-byte limit"), "got: {err}");

    assert!(read_xlsx_json_with_options(&xlsx, &ZipSecurityLimits::default()).is_ok());
}

#[test]
fn relationship_targets_resolve_to_zip_paths() {
    let sheet = "xl/worksheets/sheet1.xml";
    assert_eq!(
        resolve_target(sheet, "/xl/comments1.xml"),
        "xl/comments1.xml"
    );
    assert_eq!(
        resolve_target(sheet, "../comments1.xml"),
        "xl/comments1.xml"
    );
    // One leading '/' marks a package-absolute part name; a '//' network-path
    // reference is not a part name and no longer resolves to one.
    assert_eq!(
        resolve_target(sheet, "//xl/comments1.xml"),
        "/xl/comments1.xml"
    );
    assert_eq!(workbook_part_path("/xl/a.xml"), "xl/a.xml");
    assert_eq!(workbook_part_path("a.xml"), "xl/a.xml");
}

// ---------------------------------------------------------------------------
// Many relationships to one part: the part is parsed once
// ---------------------------------------------------------------------------

/// How many relationships the amplification tests point at one part.
const REPEATS: usize = 10_000;

/// A `.rels` part with one relationship per `(type, target)` pair, in order.
fn rels_xml(rels: &[(&str, &str)]) -> String {
    let mut xml = String::from(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">"#,
    );
    for (n, (rel_type, target)) in rels.iter().enumerate() {
        let _ = write!(
            xml,
            r#"<Relationship Id="rId{}" Type="{rel_type}" Target="{target}"/>"#,
            n + 1
        );
    }
    xml.push_str("</Relationships>");
    xml
}

/// Rebuilds `xlsx` with the first relationship of `rel_type` in `rels_path`
/// named `copies` more times, under fresh Ids and with the same target.
fn with_repeated_rel(xlsx: &[u8], rels_path: &str, rel_type: &str, copies: usize) -> Vec<u8> {
    let rels = String::from_utf8(entries_under(xlsx, &[rels_path])[rels_path].clone()).unwrap();
    let target = Relationships::parse(rels.as_bytes())
        .unwrap()
        .find_by_type(rel_type)
        .next()
        .unwrap_or_else(|| panic!("{rels_path} has no {rel_type} relationship"))
        .target
        .to_string();
    let mut repeats = String::new();
    for n in 0..copies {
        let _ = write!(
            repeats,
            r#"<Relationship Id="rIdRepeat{n}" Type="{rel_type}" Target="{target}"/>"#
        );
    }
    let rels = rels.replace("</Relationships>", &format!("{repeats}</Relationships>"));
    with_part(xlsx, rels_path, rels.as_bytes())
}

/// The count of the log digest's `kind` events (0 when there are none).
fn digest_count(ctx: &ReaderContext, kind: &str) -> usize {
    ctx.log
        .entries
        .iter()
        .find(|e| e.kind == kind)
        .map_or(0, |e| e.count)
}

/// One per-sheet part type for the many-relationships tests.
struct PerSheetCase {
    rel_type: &'static str,
    /// The part's ZIP path, and the target the sheet's .rels names it by.
    part: &'static str,
    target: &'static str,
    /// The part's XML, which holds `items` model items.
    xml: &'static str,
    items: usize,
    json_key: &'static str,
    /// The log digest kind a repeated relationship is counted under.
    repeated: &'static str,
    /// The log digest kind a part an earlier sheet read is counted under.
    shared: &'static str,
    model_len: fn(&WorksheetXml) -> usize,
    resolve: fn(&mut ReaderContext) -> Result<Vec<usize>>,
}

/// The lengths of each sheet's list in a resolver's result.
fn lens<T>(per_sheet: Vec<Vec<T>>) -> Vec<usize> {
    per_sheet.iter().map(Vec::len).collect()
}

const PER_SHEET_CASES: [PerSheetCase; 5] = [
    PerSheetCase {
        rel_type: REL_COMMENTS,
        part: "xl/comments1.xml",
        target: "../comments1.xml",
        xml: r#"<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><authors><author>Ann</author></authors><commentList><comment ref="A1" authorId="0"><text><t>one</t></text></comment><comment ref="B1" authorId="0"><text><t>two</t></text></comment></commentList></comments>"#,
        items: 2,
        json_key: "comments",
        repeated: "comments part named again, skipped",
        shared: "comments part shared with an earlier sheet, skipped",
        model_len: |ws| ws.comments.len(),
        resolve: |ctx| resolve_comments(ctx, &mut ExpansionBudget::new(usize::MAX)).map(lens),
    },
    PerSheetCase {
        rel_type: REL_THREADED_COMMENTS,
        part: "xl/threadedComments/threadedComment1.xml",
        target: "../threadedComments/threadedComment1.xml",
        xml: r#"<ThreadedComments xmlns="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments"><threadedComment ref="A1" dT="2026-01-01T00:00:00.00" personId="{P}" id="{1}"><text>one</text></threadedComment><threadedComment ref="B1" dT="2026-01-01T00:00:00.00" personId="{P}" id="{2}"><text>two</text></threadedComment></ThreadedComments>"#,
        items: 2,
        json_key: "threadedComments",
        repeated: "threaded comments part named again, skipped",
        shared: "threaded comments part shared with an earlier sheet, skipped",
        model_len: |ws| ws.threaded_comments.len(),
        resolve: |ctx| resolve_threaded_comments(ctx).map(lens),
    },
    PerSheetCase {
        rel_type: REL_SLICER,
        part: "xl/slicers/slicer1.xml",
        target: "../slicers/slicer1.xml",
        xml: r#"<slicers xmlns="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"><slicer name="S1" cache="S1"/><slicer name="S2" cache="S2"/></slicers>"#,
        items: 2,
        json_key: "slicers",
        repeated: "slicer part named again, skipped",
        shared: "slicer part shared with an earlier sheet, skipped",
        model_len: |ws| ws.slicers.len(),
        resolve: |ctx| resolve_slicers(ctx).map(lens),
    },
    PerSheetCase {
        rel_type: REL_TIMELINE,
        part: "xl/timelines/timeline1.xml",
        target: "../timelines/timeline1.xml",
        xml: r#"<timelines xmlns="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main"><timeline name="T1" cache="T1"/><timeline name="T2" cache="T2"/></timelines>"#,
        items: 2,
        json_key: "timelines",
        repeated: "timeline part named again, skipped",
        shared: "timeline part shared with an earlier sheet, skipped",
        model_len: |ws| ws.timelines.len(),
        resolve: |ctx| resolve_timelines(ctx).map(lens),
    },
    PerSheetCase {
        rel_type: REL_TABLE,
        part: "xl/tables/table1.xml",
        target: "../tables/table1.xml",
        xml: r#"<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="T" displayName="T" ref="A1:B2"><tableColumns count="2"><tableColumn id="1" name="a"/><tableColumn id="2" name="b"/></tableColumns></table>"#,
        items: 1,
        json_key: "tables",
        repeated: "table part named again, skipped",
        shared: "table part shared with an earlier sheet, skipped",
        model_len: |ws| ws.tables.len(),
        resolve: |ctx| resolve_tables(ctx).map(lens),
    },
];

#[test]
fn one_part_named_by_many_relationships_is_read_once() {
    for case in &PER_SHEET_CASES {
        let label = case.json_key;
        let rels: Vec<(&str, &str)> = vec![(case.rel_type, case.target); REPEATS];
        let xlsx = with_part(&build_minimal_xlsx(), case.part, case.xml.as_bytes());
        let xlsx = with_part(
            &xlsx,
            "xl/worksheets/_rels/sheet1.xml.rels",
            rels_xml(&rels).as_bytes(),
        );

        // The resolver parses the part once and counts every repeat under
        // one aggregated warning.
        let mut ctx = parse_common(&xlsx, &ZipSecurityLimits::default(), None).unwrap();
        assert_eq!((case.resolve)(&mut ctx).unwrap(), [case.items], "{label}");
        let repeats = ctx
            .log
            .entries
            .iter()
            .filter(|e| e.kind == case.repeated)
            .map(|e| (e.level, e.count, e.first.clone()))
            .collect::<Vec<_>>();
        assert_eq!(
            repeats,
            [(Level::Warn, REPEATS - 1, format!("{:?}", case.part))],
            "{label}"
        );
        drop(ctx);

        let wb = read_xlsx(&xlsx).unwrap();
        assert_eq!(
            (case.model_len)(&wb.sheets[0].worksheet),
            case.items,
            "{label}"
        );
        assert!(!wb.preserved_entries.contains_key(case.part), "{label}");
        let json: serde_json::Value =
            serde_json::from_str(&read_xlsx_json(&xlsx).unwrap()).unwrap();
        assert_eq!(
            json["sheets"][0]["worksheet"][case.json_key]
                .as_array()
                .map(Vec::len),
            Some(case.items),
            "{label}"
        );
    }
}

#[test]
fn two_distinct_comments_parts_on_one_sheet_both_arrive() {
    let second = PER_SHEET_CASES[0]
        .xml
        .replace("one", "three")
        .replace("two", "four");
    let xlsx = with_part(
        &build_minimal_xlsx(),
        "xl/comments1.xml",
        PER_SHEET_CASES[0].xml.as_bytes(),
    );
    let xlsx = with_part(&xlsx, "xl/comments2.xml", second.as_bytes());
    let xlsx = with_part(
        &xlsx,
        "xl/worksheets/_rels/sheet1.xml.rels",
        rels_xml(&[
            (REL_COMMENTS, "../comments1.xml"),
            (REL_COMMENTS, "../comments2.xml"),
            (REL_COMMENTS, "../comments1.xml"),
            (REL_COMMENTS, "/xl/comments2.xml"),
        ])
        .as_bytes(),
    );

    let wb = read_xlsx(&xlsx).unwrap();
    let texts: Vec<&str> = wb.sheets[0]
        .worksheet
        .comments
        .iter()
        .map(|c| c.text.as_str())
        .collect();
    assert_eq!(texts, ["one", "two", "three", "four"]);
    assert!(!wb.preserved_entries.contains_key("xl/comments1.xml"));
    assert!(!wb.preserved_entries.contains_key("xl/comments2.xml"));
}

#[test]
fn a_comments_part_two_sheets_share_is_read_for_the_first_only() {
    let xlsx = with_part(
        &build_multi_sheet_xlsx(3),
        "xl/comments1.xml",
        PER_SHEET_CASES[0].xml.as_bytes(),
    );
    let rels = rels_xml(&[(REL_COMMENTS, "../comments1.xml")]);
    let xlsx = with_part(
        &xlsx,
        "xl/worksheets/_rels/sheet1.xml.rels",
        rels.as_bytes(),
    );
    let xlsx = with_part(
        &xlsx,
        "xl/worksheets/_rels/sheet3.xml.rels",
        rels.as_bytes(),
    );

    let mut ctx = parse_common(&xlsx, &ZipSecurityLimits::default(), None).unwrap();
    assert_eq!(
        lens(resolve_comments(&mut ctx, &mut ExpansionBudget::new(usize::MAX)).unwrap()),
        [2, 0, 0]
    );
    // One parse; the third sheet's relationship is skipped, not a repeat.
    assert_eq!(digest_count(&ctx, "comments part found"), 1);
    assert_eq!(digest_count(&ctx, PER_SHEET_CASES[0].repeated), 0);
    assert_eq!(digest_count(&ctx, PER_SHEET_CASES[0].shared), 1);
    drop(ctx);

    let wb = read_xlsx(&xlsx).unwrap();
    for (i, expected) in [(0, 2), (1, 0), (2, 0)] {
        assert_eq!(wb.sheets[i].worksheet.comments.len(), expected, "sheet {i}");
    }
}

#[test]
fn two_sheets_with_their_own_comments_parts_each_keep_their_own() {
    let comments2 = PER_SHEET_CASES[0]
        .xml
        .replace("one", "three")
        .replace("two", "four");
    let xlsx = with_part(
        &build_multi_sheet_xlsx(2),
        "xl/comments1.xml",
        PER_SHEET_CASES[0].xml.as_bytes(),
    );
    let xlsx = with_part(&xlsx, "xl/comments2.xml", comments2.as_bytes());
    let xlsx = with_part(
        &xlsx,
        "xl/worksheets/_rels/sheet1.xml.rels",
        rels_xml(&[(REL_COMMENTS, "../comments1.xml")]).as_bytes(),
    );
    let xlsx = with_part(
        &xlsx,
        "xl/worksheets/_rels/sheet2.xml.rels",
        rels_xml(&[(REL_COMMENTS, "../comments2.xml")]).as_bytes(),
    );

    let mut ctx = parse_common(&xlsx, &ZipSecurityLimits::default(), None).unwrap();
    assert_eq!(
        lens(resolve_comments(&mut ctx, &mut ExpansionBudget::new(usize::MAX)).unwrap()),
        [2, 2]
    );
    // A second owner of a part of its own is not sharing anything.
    assert_eq!(digest_count(&ctx, PER_SHEET_CASES[0].shared), 0);
    drop(ctx);

    let wb = read_xlsx(&xlsx).unwrap();
    let texts: Vec<Vec<&str>> = wb
        .sheets
        .iter()
        .map(|s| {
            s.worksheet
                .comments
                .iter()
                .map(|c| c.text.as_str())
                .collect()
        })
        .collect();
    assert_eq!(texts, [["one", "two"], ["three", "four"]]);
    assert_eq!(
        per_sheet_json(&read_xlsx_json(&xlsx).unwrap(), "comments"),
        [2, 2]
    );
}

/// One relationship of the optional-parts fixture named many more times.
struct FixtureRepeat {
    rels_path: &'static str,
    rel_type: &'static str,
    repeated: &'static str,
}

#[test]
fn repeated_drawing_pivot_and_cache_relationships_are_read_once() {
    let cases = [
        FixtureRepeat {
            rels_path: "xl/worksheets/_rels/sheet1.xml.rels",
            rel_type: REL_DRAWING,
            repeated: "drawing part named again, skipped",
        },
        FixtureRepeat {
            rels_path: "xl/worksheets/_rels/sheet1.xml.rels",
            rel_type: REL_PIVOT_TABLE,
            repeated: "pivot table part named again, skipped",
        },
        FixtureRepeat {
            rels_path: "xl/_rels/workbook.xml.rels",
            rel_type: REL_PIVOT_CACHE_DEF,
            repeated: "pivot cache definition part named again, skipped",
        },
        FixtureRepeat {
            rels_path: "xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels",
            rel_type: REL_PIVOT_CACHE_REC,
            repeated: "pivot cache records part named again, skipped",
        },
        FixtureRepeat {
            rels_path: "xl/_rels/workbook.xml.rels",
            rel_type: REL_SLICER_CACHE,
            repeated: "slicer cache part named again, skipped",
        },
    ];
    let base = build_optional_parts_xlsx();
    let honest = read_xlsx(&base).unwrap();
    for case in cases {
        let label = case.repeated;
        let xlsx = with_repeated_rel(&base, case.rels_path, case.rel_type, REPEATS - 1);

        let mut ctx = parse_common(&xlsx, &ZipSecurityLimits::default(), None).unwrap();
        let charts = resolve_charts(&mut ctx).unwrap();
        let pivots = resolve_pivots(&mut ctx).unwrap();
        let slicer_caches = resolve_slicer_caches(&mut ctx).unwrap();
        assert_eq!(lens(charts), [1], "{label}");
        assert_eq!(lens(pivots.sheet_tables), [2], "{label}");
        assert_eq!(pivots.caches.len(), 2, "{label}");
        assert_eq!(pivots.records.len(), 2, "{label}");
        assert_eq!(slicer_caches.len(), 1, "{label}");
        assert_eq!(digest_count(&ctx, case.repeated), REPEATS - 1, "{label}");
        drop(ctx);

        // The read is the honest one apart from the grown .rels part.
        let wb = read_xlsx(&xlsx).unwrap();
        let ws = &wb.sheets[0].worksheet;
        let honest_ws = &honest.sheets[0].worksheet;
        assert_eq!(
            serde_json::to_value(&ws.charts).unwrap(),
            serde_json::to_value(&honest_ws.charts).unwrap(),
            "{label}"
        );
        assert_eq!(ws.pivot_tables, honest_ws.pivot_tables, "{label}");
        assert_eq!(wb.pivot_caches, honest.pivot_caches, "{label}");
        assert_eq!(
            wb.pivot_cache_records, honest.pivot_cache_records,
            "{label}"
        );
        assert_eq!(
            wb.slicer_caches.len(),
            honest.slicer_caches.len(),
            "{label}"
        );
    }
}

#[test]
fn one_chart_named_by_many_anchors_is_read_once() {
    let base = build_optional_parts_xlsx();
    let drawing_path = "xl/drawings/drawing1.xml";
    let drawing =
        String::from_utf8(entries_under(&base, &[drawing_path])[drawing_path].clone()).unwrap();
    let start = drawing
        .find("<xdr:twoCellAnchor")
        .expect("the fixture drawing has a two-cell anchor");
    let end = drawing.rfind("</xdr:twoCellAnchor>").unwrap() + "</xdr:twoCellAnchor>".len();
    let anchor = &drawing[start..end];
    let many = format!(
        "{}{}{}",
        &drawing[..start],
        anchor.repeat(REPEATS),
        &drawing[end..]
    );
    let xlsx = with_part(&base, drawing_path, many.as_bytes());

    // The repeated anchor compresses past the default ratio guard.
    let limits = ZipSecurityLimits {
        max_compression_ratio: f64::INFINITY,
        ..Default::default()
    };
    let mut ctx = parse_common(&xlsx, &limits, None).unwrap();
    assert_eq!(lens(resolve_charts(&mut ctx).unwrap()), [1]);
    assert_eq!(digest_count(&ctx, "chart part found"), 1);
    assert_eq!(
        digest_count(&ctx, "chart part named again, skipped"),
        REPEATS - 1
    );
}

#[test]
fn pivot_cache_ids_follow_the_relationship_id_over_distinct_definitions() {
    let base = build_optional_parts_xlsx();
    let (wb_path, rels_path) = ("xl/workbook.xml", "xl/_rels/workbook.xml.rels");
    let def2 = "xl/pivotCache/pivotCacheDefinition2.xml";
    let parts = entries_under(&base, &[wb_path, rels_path, def2]);
    let rels = String::from_utf8(parts[rels_path].clone()).unwrap();
    let ids: Vec<String> = Relationships::parse(rels.as_bytes())
        .unwrap()
        .find_by_type(REL_PIVOT_CACHE_DEF)
        .map(|rel| rel.id.clone())
        .collect();
    assert_eq!(ids.len(), 2, "the fixture has two cache definitions");

    // A third definition under a relationship of its own. The first
    // definition's Id has two `<pivotCache>` entries (the first wins: 7, not
    // 8), the second has none and the third maps to 3.
    let rels = rels.replace(
        "</Relationships>",
        &format!(
            r#"<Relationship Id="rIdThird" Type="{REL_PIVOT_CACHE_DEF}" Target="pivotCache/pivotCacheDefinition3.xml"/></Relationships>"#
        ),
    );
    let workbook = String::from_utf8(parts[wb_path].clone()).unwrap();
    let (head, rest) = workbook.split_once("<pivotCaches>").unwrap();
    let (_, tail) = rest.split_once("</pivotCaches>").unwrap();
    let workbook = format!(
        r#"{head}<pivotCaches><pivotCache cacheId="7" r:id="{first}"/><pivotCache cacheId="8" r:id="{first}"/><pivotCache cacheId="3" r:id="rIdThird"/></pivotCaches>{tail}"#,
        first = ids[0]
    );
    let xlsx = with_part(&base, rels_path, rels.as_bytes());
    let xlsx = with_part(&xlsx, wb_path, workbook.as_bytes());
    let xlsx = with_part(
        &xlsx,
        "xl/pivotCache/pivotCacheDefinition3.xml",
        &parts[def2],
    );

    let ids: Vec<Option<u32>> = read_xlsx(&xlsx)
        .unwrap()
        .pivot_caches
        .iter()
        .map(|cache| cache.cache_id)
        .collect();
    assert_eq!(ids, [Some(7), None, Some(3)]);
}

#[test]
fn pivot_table_on_a_repeated_definitions_cache_id_keeps_the_family_out() {
    // Both `<pivotCache>` entries (cacheIds 0 and 1) now name the first
    // definition part, which is read once, under cacheId 0; the table bound
    // to cacheId 1 would name a cache the model lacks.
    let base = build_optional_parts_xlsx();
    let rels_path = "xl/_rels/workbook.xml.rels";
    let rels = String::from_utf8(entries_under(&base, &[rels_path])[rels_path].clone()).unwrap();
    assert!(rels.contains("pivotCache/pivotCacheDefinition2.xml"));
    let rels = rels.replace(
        "pivotCache/pivotCacheDefinition2.xml",
        "pivotCache/pivotCacheDefinition1.xml",
    );
    let xlsx = with_part(&base, rels_path, rels.as_bytes());

    let wb = read_xlsx(&xlsx).unwrap();
    assert!(wb.pivot_caches.is_empty());
    assert!(wb.sheets[0].worksheet.pivot_tables.is_empty());
    assert!(
        wb.preserved_entries
            .contains_key("xl/pivotCache/pivotCacheDefinition1.xml")
    );
    crate::writer::write_xlsx(&wb).expect("the kept-out family writes as preserved parts");
}

#[test]
fn one_timeline_cache_named_by_many_relationships_is_read_once() {
    let cache = r#"<timelineCacheDefinition xmlns="http://schemas.microsoft.com/office/spreadsheetml/2014/timeline" name="NativeTimeline_Date" sourceName="Date"/>"#;
    let rels_path = "xl/_rels/workbook.xml.rels";
    let base = with_part(
        &build_minimal_xlsx(),
        "xl/timelineCaches/timelineCache1.xml",
        cache.as_bytes(),
    );
    let rels = String::from_utf8(entries_under(&base, &[rels_path])[rels_path].clone()).unwrap();
    let rels = rels.replace(
        "</Relationships>",
        &format!(
            r#"<Relationship Id="rIdTimelineCache" Type="{REL_TIMELINE_CACHE}" Target="timelineCaches/timelineCache1.xml"/></Relationships>"#
        ),
    );
    let xlsx = with_part(&base, rels_path, rels.as_bytes());
    let xlsx = with_repeated_rel(&xlsx, rels_path, REL_TIMELINE_CACHE, REPEATS - 1);

    let mut ctx = parse_common(&xlsx, &ZipSecurityLimits::default(), None).unwrap();
    assert_eq!(resolve_timeline_caches(&mut ctx).unwrap().len(), 1);
    assert_eq!(
        digest_count(&ctx, "timeline cache part named again, skipped"),
        REPEATS - 1
    );
}

// ---------------------------------------------------------------------------
// One worksheet part per sheet, one owner per part
// ---------------------------------------------------------------------------

/// How many sheets the sharing tests point at one part.
const SHARING_SHEETS: usize = 200;

/// The relationship type of a worksheet part.
const REL_WORKSHEET_TYPE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet";

/// The `(level, count, first)` lines of the log digest's `kind` events.
fn digest_lines(ctx: &ReaderContext, kind: &str) -> Vec<(Level, usize, String)> {
    ctx.log
        .entries
        .iter()
        .filter(|e| e.kind == kind)
        .map(|e| (e.level, e.count, e.first.clone()))
        .collect()
}

/// One aggregated warning, counted `count` times, whose first value is `path`.
fn warned(path: &str, count: usize) -> [(Level, usize, String); 1] {
    [(Level::Warn, count, format!("{path:?}"))]
}

#[test]
fn sheet_rows_on_one_worksheet_part_fail_the_read() {
    let base = build_multi_sheet_xlsx(2);
    let (wb_path, rels_path) = ("xl/workbook.xml", "xl/_rels/workbook.xml.rels");
    let parts = entries_under(&base, &[wb_path, rels_path]);
    let workbook = String::from_utf8(parts[wb_path].clone()).unwrap();
    let rels = String::from_utf8(parts[rels_path].clone()).unwrap();
    assert!(workbook.contains(r#"name="Sheet2""#) && workbook.contains(r#"r:id="rId2""#));
    assert!(rels.contains(r#"Target="worksheets/sheet2.xml""#));
    // The second sheet's name holds a newline, which the message escapes.
    let workbook = workbook.replace(r#"name="Sheet2""#, r#"name="Two&#10;Lines""#);
    let cases = [
        (
            "one r:id on two rows",
            workbook.replace(r#"r:id="rId2""#, r#"r:id="rId1""#),
            rels.clone(),
        ),
        (
            "two r:ids with one target",
            workbook.clone(),
            rels.replace("worksheets/sheet2.xml", "worksheets/sheet1.xml"),
        ),
        (
            "a relative and an absolute target",
            workbook.clone(),
            rels.replace("worksheets/sheet2.xml", "/xl/worksheets/sheet1.xml"),
        ),
    ];

    for (label, workbook, rels) in cases {
        let xlsx = with_part(&base, wb_path, workbook.as_bytes());
        let xlsx = with_part(&xlsx, rels_path, rels.as_bytes());
        // Rejected while the sheet rows are resolved, before any sheet .rels
        // or worksheet is parsed, on both read paths.
        let results = [
            parse_common(&xlsx, &ZipSecurityLimits::default(), None).map(drop),
            read_xlsx(&xlsx).map(drop),
            read_xlsx_json(&xlsx).map(drop),
        ];
        for result in results {
            let err = result.expect_err(label);
            assert_eq!(err.code(), "SECURITY", "{label}: {err}");
            let msg = err.to_string();
            for quoted in [
                r#""xl/worksheets/sheet1.xml""#,
                r#""Sheet1""#,
                r#""Two\nLines""#,
            ] {
                assert!(msg.contains(quoted), "{label}: {msg}");
            }
        }
    }

    // Two rows on two parts still read, on both paths.
    assert_eq!(read_xlsx(&base).unwrap().sheets.len(), 2);
    assert_eq!(
        per_sheet_json(&read_xlsx_json(&base).unwrap(), "rows"),
        [1, 1]
    );
}

/// A package of `sheets` sheets, each on a worksheet part of its own, that
/// also holds `parts`; sheet `n` (from 1) gets the .rels `sheet_rels(n)`.
fn sheets_with_rels(
    sheets: usize,
    parts: impl IntoIterator<Item = (String, Vec<u8>)>,
    sheet_rels: impl Fn(usize) -> String,
) -> Vec<u8> {
    let mut entries: Vec<ZipEntry> = read_zip_entries(
        &build_multi_sheet_xlsx(sheets),
        &ZipSecurityLimits::default(),
    )
    .expect("fixture should be readable")
    .into_iter()
    .chain(parts)
    .chain((1..=sheets).map(|n| {
        (
            format!("xl/worksheets/_rels/sheet{n}.xml.rels"),
            sheet_rels(n).into_bytes(),
        )
    }))
    .map(|(name, data)| ZipEntry { name, data })
    .collect();
    entries.sort_by(|a, b| a.name.cmp(&b.name));
    write_zip(&entries).expect("fixture should be writable")
}

/// Each sheet's count of the items `model_len` counts.
fn per_sheet(wb: &WorkbookData, model_len: fn(&WorksheetXml) -> usize) -> Vec<usize> {
    wb.sheets.iter().map(|s| model_len(&s.worksheet)).collect()
}

/// Each sheet's count of `key` items in the JSON of a read.
fn per_sheet_json(json: &str, key: &str) -> Vec<usize> {
    let json: serde_json::Value = serde_json::from_str(json).unwrap();
    json["sheets"]
        .as_array()
        .unwrap()
        .iter()
        .map(|s| s["worksheet"][key].as_array().map_or(0, Vec::len))
        .collect()
}

/// `first` for the first of `sheets` sheets and 0 for every other.
fn first_only(first: usize, sheets: usize) -> Vec<usize> {
    let mut counts = vec![0; sheets];
    counts[0] = first;
    counts
}

#[test]
fn a_part_many_sheets_share_is_read_for_the_first_only() {
    for case in &PER_SHEET_CASES {
        let label = case.json_key;
        let rels = rels_xml(&[(case.rel_type, case.target)]);
        let xlsx = sheets_with_rels(
            SHARING_SHEETS,
            [(case.part.to_owned(), case.xml.as_bytes().to_vec())],
            |_| rels.clone(),
        );
        let expected = first_only(case.items, SHARING_SHEETS);

        let mut ctx = parse_common(&xlsx, &ZipSecurityLimits::default(), None).unwrap();
        assert_eq!((case.resolve)(&mut ctx).unwrap(), expected, "{label}");
        assert_eq!(
            digest_lines(&ctx, case.shared),
            warned(case.part, SHARING_SHEETS - 1),
            "{label}"
        );
        assert_eq!(digest_count(&ctx, case.repeated), 0, "{label}");
        drop(ctx);

        let wb = read_xlsx(&xlsx).unwrap();
        assert_eq!(per_sheet(&wb, case.model_len), expected, "{label}");
        assert!(!wb.preserved_entries.contains_key(case.part), "{label}");
        assert_eq!(
            per_sheet_json(&read_xlsx_json(&xlsx).unwrap(), case.json_key),
            expected,
            "{label}"
        );
    }
}

/// A sheet-owned part type of the optional-parts fixture, for the sharing
/// tests.
struct FixtureShare {
    /// The fixture entries the package holds.
    prefixes: &'static [&'static str],
    rel_type: &'static str,
    /// The part's ZIP path, and the target each sheet's .rels names it by.
    part: &'static str,
    target: &'static str,
    /// The log digest kinds a part found and a shared part are counted under.
    found: &'static str,
    shared: &'static str,
    model_len: fn(&WorksheetXml) -> usize,
    resolve: fn(&mut ReaderContext) -> Result<Vec<usize>>,
}

const FIXTURE_SHARES: [FixtureShare; 2] = [
    FixtureShare {
        prefixes: &["xl/drawings/", "xl/charts/"],
        rel_type: REL_DRAWING,
        part: "xl/drawings/drawing1.xml",
        target: "../drawings/drawing1.xml",
        found: "drawing part found",
        shared: "drawing part shared with an earlier sheet, skipped",
        model_len: |ws| ws.charts.len(),
        resolve: |ctx| resolve_charts(ctx).map(lens),
    },
    FixtureShare {
        prefixes: &["xl/pivotTables/pivotTable1.xml"],
        rel_type: REL_PIVOT_TABLE,
        part: "xl/pivotTables/pivotTable1.xml",
        target: "../pivotTables/pivotTable1.xml",
        found: "pivot table part found",
        shared: "pivot table part shared with an earlier sheet, skipped",
        model_len: |ws| ws.pivot_tables.len(),
        resolve: |ctx| resolve_pivots(ctx).map(|family| lens(family.sheet_tables)),
    },
];

#[test]
fn a_drawing_or_pivot_table_many_sheets_share_is_read_for_the_first_only() {
    let fixture = build_optional_parts_xlsx();
    for case in &FIXTURE_SHARES {
        let label = case.shared;
        let rels = rels_xml(&[(case.rel_type, case.target)]);
        let xlsx = sheets_with_rels(
            SHARING_SHEETS,
            entries_under(&fixture, case.prefixes),
            |_| rels.clone(),
        );
        let expected = first_only(1, SHARING_SHEETS);

        let mut ctx = parse_common(&xlsx, &ZipSecurityLimits::default(), None).unwrap();
        assert_eq!((case.resolve)(&mut ctx).unwrap(), expected, "{label}");
        assert_eq!(digest_count(&ctx, case.found), 1, "{label}");
        assert_eq!(
            digest_lines(&ctx, case.shared),
            warned(case.part, SHARING_SHEETS - 1),
            "{label}"
        );
        drop(ctx);

        let wb = read_xlsx(&xlsx).unwrap();
        assert_eq!(per_sheet(&wb, case.model_len), expected, "{label}");
    }
}

/// The optional-parts fixture with its second pivot cache definition naming
/// the first one's records part.
fn definitions_sharing_one_records_part() -> Vec<u8> {
    let base = build_optional_parts_xlsx();
    let rels_path = "xl/pivotCache/_rels/pivotCacheDefinition2.xml.rels";
    let rels = String::from_utf8(entries_under(&base, &[rels_path])[rels_path].clone()).unwrap();
    assert!(rels.contains("pivotCacheRecords2.xml"));
    let rels = rels.replace("pivotCacheRecords2.xml", "pivotCacheRecords1.xml");
    with_part(&base, rels_path, rels.as_bytes())
}

/// The log digest kind a records part an earlier definition read is counted
/// under.
const SHARED_RECORDS: &str =
    "pivot cache records part shared with an earlier cache definition, skipped";

#[test]
fn a_records_part_two_cache_definitions_share_is_read_for_the_first_only() {
    let xlsx = definitions_sharing_one_records_part();

    let mut ctx = parse_common(&xlsx, &ZipSecurityLimits::default(), None).unwrap();
    let family = resolve_pivots(&mut ctx).unwrap();
    assert_eq!(family.caches.len(), 2);
    assert_eq!(family.records.len(), 1);
    assert_eq!(lens(family.sheet_tables), [2]);
    assert_eq!(digest_count(&ctx, "pivot cache records part found"), 1);
    assert_eq!(
        digest_lines(&ctx, SHARED_RECORDS),
        warned("xl/pivotCache/pivotCacheRecords1.xml", 1)
    );
    drop(ctx);

    let wb = read_xlsx(&xlsx).unwrap();
    assert_eq!(wb.pivot_cache_records.len(), 1);
    assert!(
        !wb.preserved_entries
            .contains_key("xl/pivotCache/pivotCacheRecords1.xml")
    );
}

#[test]
fn records_read_after_a_shared_records_part_keep_the_pivot_family_out() {
    // A third definition, a copy of the second with its own records part,
    // comes after the second, which now names the first one's records: its
    // records would pair with the second cache.
    let base = definitions_sharing_one_records_part();
    let (wb_rels_path, def2, def2_rels) = (
        "xl/_rels/workbook.xml.rels",
        "xl/pivotCache/pivotCacheDefinition2.xml",
        "xl/pivotCache/_rels/pivotCacheDefinition2.xml.rels",
    );
    let parts = entries_under(&base, &[wb_rels_path, def2, def2_rels]);
    let wb_rels = String::from_utf8(parts[wb_rels_path].clone())
        .unwrap()
        .replace(
            "</Relationships>",
            &format!(
                r#"<Relationship Id="rIdThird" Type="{REL_PIVOT_CACHE_DEF}" Target="pivotCache/pivotCacheDefinition3.xml"/></Relationships>"#
            ),
        );
    let def3_rels = String::from_utf8(parts[def2_rels].clone())
        .unwrap()
        .replace("pivotCacheRecords1.xml", "pivotCacheRecords2.xml");
    let xlsx = with_part(&base, wb_rels_path, wb_rels.as_bytes());
    let xlsx = with_part(
        &xlsx,
        "xl/pivotCache/pivotCacheDefinition3.xml",
        &parts[def2],
    );
    let xlsx = with_part(
        &xlsx,
        "xl/pivotCache/_rels/pivotCacheDefinition3.xml.rels",
        def3_rels.as_bytes(),
    );

    let mut ctx = parse_common(&xlsx, &ZipSecurityLimits::default(), None).unwrap();
    let family = resolve_pivots(&mut ctx).unwrap();
    assert!(family.caches.is_empty());
    assert!(family.records.is_empty());
    assert_eq!(
        digest_lines(&ctx, SHARED_RECORDS),
        warned("xl/pivotCache/pivotCacheRecords1.xml", 1)
    );
    drop(ctx);

    let wb = read_xlsx(&xlsx).unwrap();
    assert!(wb.pivot_caches.is_empty());
    assert!(
        wb.preserved_entries
            .contains_key("xl/pivotCache/pivotCacheRecords2.xml")
    );
}

/// An empty relationships part.
const NO_RELS: &[u8] = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>"#;

/// Rebuilds `xlsx` without the entries named in `parts`.
fn without_parts(xlsx: &[u8], parts: &[&str]) -> Vec<u8> {
    let entries: Vec<ZipEntry> = read_zip_entries(xlsx, &ZipSecurityLimits::default())
        .expect("fixture should be readable")
        .into_iter()
        .filter(|(name, _)| !parts.contains(&name.as_str()))
        .map(|(name, data)| ZipEntry { name, data })
        .collect();
    write_zip(&entries).expect("fixture should be writable")
}

/// The optional-parts fixture with pivot cache definition `n` (1 or 2) saved
/// without records, the shape `PivotTable.SaveData = False` gives in Excel:
/// its .rels names no records part and its records part is gone.
fn cache_saved_without_records(n: usize) -> Vec<u8> {
    let base = build_optional_parts_xlsx();
    let records = format!("xl/pivotCache/pivotCacheRecords{n}.xml");
    let rels = format!("xl/pivotCache/_rels/pivotCacheDefinition{n}.xml.rels");
    assert!(entries_under(&base, &[&records]).contains_key(&records));
    with_part(&without_parts(&base, &[&records]), &rels, NO_RELS)
}

/// The one-record records part the optional-parts fixture gives a cache.
fn one_record(v: f64) -> PivotCacheRecordsData {
    from_json(serde_json::json!({ "records": [[{ "type": "number", "v": v }]] }))
}

/// Each pivot cache, named by its first field, with the records the model
/// pairs with it (`pivot_cache_records[i]` belongs to `pivot_caches[i]`),
/// sorted by name.
fn records_by_cache(wb: &WorkbookData) -> Vec<(String, Option<PivotCacheRecordsData>)> {
    let mut pairs: Vec<(String, Option<PivotCacheRecordsData>)> = wb
        .pivot_caches
        .iter()
        .enumerate()
        .map(|(i, cache)| {
            (
                cache.fields[0].name.clone(),
                wb.pivot_cache_records.get(i).cloned(),
            )
        })
        .collect();
    pairs.sort_by(|a, b| a.0.cmp(&b.0));
    pairs
}

/// The struct and JSON reads of `xlsx`.
fn both_reads(xlsx: &[u8]) -> [WorkbookData; 2] {
    [
        read_xlsx(xlsx).unwrap(),
        serde_json::from_str(&read_xlsx_json(xlsx).unwrap()).unwrap(),
    ]
}

#[test]
fn records_pair_with_the_cache_definition_whose_rels_name_them() {
    // The first definition has no records and the second has: the records
    // belong to the second cache, the first has none.
    for wb in both_reads(&cache_saved_without_records(1)) {
        assert_eq!(
            records_by_cache(&wb),
            [("A".into(), None), ("B".into(), Some(one_record(2.0)))]
        );
    }
    // The mirror case: the first has records, the second has none.
    for wb in both_reads(&cache_saved_without_records(2)) {
        assert_eq!(
            records_by_cache(&wb),
            [("A".into(), Some(one_record(1.0))), ("B".into(), None)]
        );
    }
}

#[test]
fn records_stay_on_their_cache_through_a_write() {
    for (n, expected) in [
        (1, [("A".into(), None), ("B".into(), Some(one_record(2.0)))]),
        (2, [("A".into(), Some(one_record(1.0))), ("B".into(), None)]),
    ] {
        let wb = read_xlsx(&cache_saved_without_records(n)).unwrap();
        let written = crate::writer::write_xlsx(&wb).unwrap();
        for wb in both_reads(&written) {
            assert_eq!(
                records_by_cache(&wb),
                expected,
                "cache {n} saved without records"
            );
            // Pivot tables keep their cache: PT1 names cacheId 0 (field A),
            // PT2 cacheId 1 (field B).
            for (table, field) in wb.sheets[0].worksheet.pivot_tables.iter().zip(["A", "B"]) {
                let cache = wb
                    .pivot_caches
                    .iter()
                    .find(|cache| cache.cache_id == Some(table.cache_id))
                    .expect("each pivot table names a cache");
                assert_eq!(cache.fields[0].name, field, "{}", table.name);
            }
        }
    }
}

#[test]
fn a_cache_definition_naming_two_records_parts_keeps_the_pivot_family_out() {
    // The first definition names both records parts; the second names none.
    let base = build_optional_parts_xlsx();
    let def1_rels = "xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels";
    let rels = String::from_utf8(entries_under(&base, &[def1_rels])[def1_rels].clone())
        .unwrap()
        .replace(
            "</Relationships>",
            &format!(
                r#"<Relationship Id="rId2" Type="{REL_PIVOT_CACHE_REC}" Target="pivotCacheRecords2.xml"/></Relationships>"#
            ),
        );
    let xlsx = with_part(&base, def1_rels, rels.as_bytes());
    let xlsx = with_part(
        &xlsx,
        "xl/pivotCache/_rels/pivotCacheDefinition2.xml.rels",
        NO_RELS,
    );

    for wb in both_reads(&xlsx) {
        assert!(wb.pivot_caches.is_empty());
        assert!(wb.pivot_cache_records.is_empty());
        for kept in [
            "xl/pivotCache/pivotCacheRecords1.xml",
            "xl/pivotCache/pivotCacheRecords2.xml",
        ] {
            assert!(wb.preserved_entries.contains_key(kept), "{kept}");
        }
    }
}

/// Three sheets, each with a drawing of its own (a copy of the optional-parts
/// fixture's drawing and its .rels), all three drawings naming the fixture's
/// one chart part.
fn three_drawings_naming_one_chart() -> Vec<u8> {
    let parts = entries_under(
        &build_optional_parts_xlsx(),
        &["xl/drawings/", "xl/charts/"],
    );
    let drawing = parts["xl/drawings/drawing1.xml"].clone();
    let drawing_rels = parts["xl/drawings/_rels/drawing1.xml.rels"].clone();
    assert!(String::from_utf8_lossy(&drawing_rels).contains(r#"Target="../charts/chart1.xml""#));
    let copies = (2..=3).flat_map(|n| {
        [
            (format!("xl/drawings/drawing{n}.xml"), drawing.clone()),
            (
                format!("xl/drawings/_rels/drawing{n}.xml.rels"),
                drawing_rels.clone(),
            ),
        ]
    });
    sheets_with_rels(3, parts.into_iter().chain(copies), |n| {
        rels_xml(&[(REL_DRAWING, format!("../drawings/drawing{n}.xml").as_str())])
    })
}

#[test]
fn a_chart_part_three_drawings_share_is_read_for_the_first_only() {
    let xlsx = three_drawings_naming_one_chart();

    let mut ctx = parse_common(&xlsx, &ZipSecurityLimits::default(), None).unwrap();
    assert_eq!(lens(resolve_charts(&mut ctx).unwrap()), [1, 0, 0]);
    assert_eq!(digest_count(&ctx, "drawing part found"), 3);
    assert_eq!(digest_count(&ctx, "chart part found"), 1);
    assert_eq!(
        digest_lines(&ctx, "chart part shared with an earlier drawing, skipped"),
        warned("xl/charts/chart1.xml", 2)
    );
    assert_eq!(digest_count(&ctx, "chart part named again, skipped"), 0);
    drop(ctx);

    let wb = read_xlsx(&xlsx).unwrap();
    assert_eq!(per_sheet(&wb, |ws| ws.charts.len()), [1, 0, 0]);
    assert_eq!(
        per_sheet_json(&read_xlsx_json(&xlsx).unwrap(), "charts"),
        [1, 0, 0]
    );
}

/// The optional-parts fixture with a comment, a threaded comment (with its
/// person), a table and a timeline (with its cache) on its sheet too, so that
/// it holds one part of each sheet-owned type the writer names itself.
fn build_shareable_parts_xlsx() -> Vec<u8> {
    use serde_json::json;

    let mut wb = read_xlsx(&build_optional_parts_xlsx()).unwrap();
    wb.persons = from_json(json!([{ "id": "{P}", "displayName": "Ann" }]));
    wb.timeline_caches = from_json(json!([{ "name": "Timeline_A", "sourceName": "A" }]));
    let ws = &mut wb.sheets[0].worksheet;
    ws.comments = vec![comments::Comment {
        cell_ref: "A1".into(),
        author: "Ann".into(),
        text: "one".into(),
    }];
    ws.threaded_comments = from_json(json!([{
        "id": "{1}",
        "refCell": "B1",
        "personId": "{P}",
        "text": "one",
        "timestamp": "2026-01-01T00:00:00.00",
    }]));
    ws.timelines = from_json(json!([{ "name": "Timeline_A", "cacheName": "Timeline_A" }]));
    let header = |reference: &str, text: &str| Cell {
        reference: reference.into(),
        cell_type: CellType::SharedString,
        value: Some(text.into()),
        ..Default::default()
    };
    ws.rows = vec![Row {
        index: 1,
        cells: vec![header("J1", "a"), header("K1", "b")],
        ..Default::default()
    }];
    ws.tables = vec![from_json(serde_json::json!({
        "id": 1,
        "name": "T",
        "displayName": "T",
        "ref": "J1:K2",
        "columns": [{ "id": 1, "name": "a" }, { "id": 2, "name": "b" }],
    }))];
    crate::writer::write_xlsx(&wb).expect("fixture should be writable")
}

/// `xlsx`, a one-sheet package, with a second sheet on a worksheet part of its
/// own whose .rels is `rels`.
fn with_second_sheet(xlsx: &[u8], rels: &str) -> Vec<u8> {
    let (wb_path, wb_rels_path) = ("xl/workbook.xml", "xl/_rels/workbook.xml.rels");
    let parts = entries_under(xlsx, &[wb_path, wb_rels_path]);
    let workbook = String::from_utf8(parts[wb_path].clone()).unwrap();
    let (head, tail) = workbook.split_once("</sheets>").unwrap();
    let workbook =
        format!(r#"{head}<sheet name="Sheet2" sheetId="2" r:id="rIdSheet2"/></sheets>{tail}"#);
    let wb_rels = String::from_utf8(parts[wb_rels_path].clone())
        .unwrap()
        .replace(
            "</Relationships>",
            &format!(
                r#"<Relationship Id="rIdSheet2" Type="{REL_WORKSHEET_TYPE}" Target="worksheets/sheet2.xml"/></Relationships>"#
            ),
        );
    let xlsx = with_part(xlsx, wb_path, workbook.as_bytes());
    let xlsx = with_part(&xlsx, wb_rels_path, wb_rels.as_bytes());
    let xlsx = with_part(
        &xlsx,
        "xl/worksheets/sheet2.xml",
        br#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>"#,
    );
    with_part(
        &xlsx,
        "xl/worksheets/_rels/sheet2.xml.rels",
        rels.as_bytes(),
    )
}

/// The accepted sharing cases, by label: a second sheet naming each of the
/// first sheet's comments, table, drawing, pivot table and slicer parts in
/// turn, three drawings naming one chart part, and two pivot cache
/// definitions naming one records part.
fn shared_part_packages() -> Vec<(String, Vec<u8>)> {
    let base = build_shareable_parts_xlsx();
    let rels_path = "xl/worksheets/_rels/sheet1.xml.rels";
    let sheet1_rels = Relationships::parse(&entries_under(&base, &[rels_path])[rels_path]).unwrap();
    let mut packages: Vec<(String, Vec<u8>)> = [
        REL_COMMENTS,
        REL_THREADED_COMMENTS,
        REL_TABLE,
        REL_DRAWING,
        REL_PIVOT_TABLE,
        REL_SLICER,
        REL_TIMELINE,
    ]
    .into_iter()
    .map(|rel_type| {
        let target = sheet1_rels
            .find_by_type(rel_type)
            .next()
            .unwrap_or_else(|| panic!("sheet 1 has no {rel_type} relationship"))
            .target
            .to_string();
        let xlsx = with_second_sheet(&base, &rels_xml(&[(rel_type, target.as_str())]));
        (format!("a second sheet naming {target}"), xlsx)
    })
    .collect();
    packages.push((
        "three drawings naming one chart part".into(),
        three_drawings_naming_one_chart(),
    ));
    packages.push((
        "two cache definitions naming one records part".into(),
        definitions_sharing_one_records_part(),
    ));
    packages
}

/// Each sheet's count of comments, tables, charts, pivot tables, slicers,
/// threaded comments and timelines.
fn part_counts(wb: &WorkbookData) -> Vec<[usize; 7]> {
    wb.sheets
        .iter()
        .map(|sheet| {
            let ws = &sheet.worksheet;
            [
                ws.comments.len(),
                ws.tables.len(),
                ws.charts.len(),
                ws.pivot_tables.len(),
                ws.slicers.len(),
                ws.threaded_comments.len(),
                ws.timelines.len(),
            ]
        })
        .collect()
}

#[test]
fn accepted_shared_part_packages_write_and_read_back() {
    for (label, xlsx) in shared_part_packages() {
        let wb = read_xlsx(&xlsx).unwrap_or_else(|e| panic!("{label}: {e}"));
        // Only the first owner holds the shared part's items.
        let counts = part_counts(&wb);
        assert!(
            counts[1..].iter().all(|c| *c == [0; 7]),
            "{label}: {counts:?}"
        );

        let written = crate::writer::write_xlsx(&wb).unwrap_or_else(|e| panic!("{label}: {e}"));
        // A later sheet never points at the drawing the first one's charts
        // were written to (Excel rejects two sheets on one drawing).
        for (path, xml) in entries_under(&written, &["xl/worksheets/sheet"]) {
            let xml = String::from_utf8(xml).unwrap();
            assert!(
                path == "xl/worksheets/sheet1.xml" || !xml.contains("<drawing "),
                "{label}: {path}: {xml}"
            );
        }
        let back = read_xlsx(&written).unwrap_or_else(|e| panic!("{label}: {e}"));
        assert_eq!(part_counts(&back), counts, "{label}");
        assert_eq!(back.pivot_caches.len(), wb.pivot_caches.len(), "{label}");
        assert_eq!(back.slicer_caches.len(), wb.slicer_caches.len(), "{label}");
    }
}

// ---------------------------------------------------------------------------
// Shared text budget (ExpansionBudget)
// ---------------------------------------------------------------------------

#[test]
fn optional_part_keeps_a_size_limit_fatal() {
    let mut log = LogDigest::default();
    let over: Result<u8> = Err(ModernXlsxError::SizeLimit("budget".into()));
    assert_eq!(
        optional_part(&mut log, "p", over).unwrap_err().code(),
        "SIZE_LIMIT",
        "a part that passes the shared text budget must not be dropped as malformed"
    );
    assert!(log.entries.is_empty());
}

/// Limits of `max_json_length` bytes, the other limits at their defaults.
fn json_limit(max_json_length: usize) -> ZipSecurityLimits {
    ZipSecurityLimits {
        max_json_length,
        ..Default::default()
    }
}

/// A comments part on sheet 1 with one author of 2,000 bytes on 25 comments
/// (50,000 bytes of copies), ending in `tail`, and 20 cells naming one
/// shared string of 1,000 bytes (20,000 more).
fn author_comments_xlsx(tail: &str) -> Vec<u8> {
    let author = "a".repeat(2_000);
    let mut part = format!(
        r#"<comments xmlns="{SPREADSHEET_NS}"><authors><author>{author}</author></authors><commentList>"#
    );
    for i in 1..=25 {
        let _ = write!(
            part,
            r#"<comment ref="A{i}" authorId="0"><text><t>n</t></text></comment>"#
        );
    }
    part.push_str(tail);
    let rels = format!(
        r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="{REL_COMMENTS}" Target="../comments1.xml"/></Relationships>"#
    );
    let mut rows = String::new();
    for r in 1..=20 {
        let _ = write!(rows, r#"<row r="{r}"><c r="A{r}" t="s"><v>0</v></c></row>"#);
    }
    let xlsx = build_xlsx_with_parts(
        &format!(
            r#"<sst xmlns="{SPREADSHEET_NS}" count="1" uniqueCount="1"><si><t>{}</t></si></sst>"#,
            "s".repeat(1_000)
        ),
        &format!(
            r#"<worksheet xmlns="{SPREADSHEET_NS}"><sheetData>{rows}</sheetData></worksheet>"#
        ),
    );
    let xlsx = with_part(&xlsx, "xl/comments1.xml", part.as_bytes());
    with_part(
        &xlsx,
        "xl/worksheets/_rels/sheet1.xml.rels",
        rels.as_bytes(),
    )
}

/// Each comment's copy of its author counts; a part that does not parse
/// gives its count back, since its comments never reach the result: 50,000
/// bytes of authors plus 20,000 of cells pass a 60,000-byte budget, the
/// 20,000 alone do not.
#[test]
fn comment_authors_count_and_a_malformed_part_gives_its_count_back() {
    let limits = json_limit(60_000);
    let good = author_comments_xlsx("</commentList></comments>");
    let err = read_xlsx_json_with_options(&good, &limits).unwrap_err();
    assert_eq!(err.code(), "SIZE_LIMIT");
    assert!(err.to_string().contains("shared text"), "{err}");
    assert!(err.to_string().contains("60000-byte limit"), "{err}");
    assert_eq!(
        read_xlsx_with_options(&good, &limits).unwrap_err().code(),
        "SIZE_LIMIT"
    );
    assert_eq!(
        read_xlsx(&good).unwrap().sheets[0].worksheet.comments.len(),
        25
    );

    let malformed = author_comments_xlsx("</commentList></wrong>");
    let json = read_xlsx_json_with_options(&malformed, &limits).unwrap();
    assert!(json.contains("\"xl/comments1.xml\""), "kept verbatim");
    let wb = read_xlsx_with_options(&malformed, &limits).unwrap();
    assert!(wb.sheets[0].worksheet.comments.is_empty());
    assert!(wb.preserved_entries.contains_key("xl/comments1.xml"));
}

/// A row cut off by the end of its part is dropped, so its cells give back
/// what they counted: 40,000 bytes in sheet 1's open row plus 30,000 in
/// sheet 2 fit a 60,000-byte budget.
#[test]
fn a_dropped_row_gives_its_count_back() {
    let entry = "s".repeat(1_000);
    let cells = |n: u32| -> String {
        let mut out = String::new();
        for c in 0..n {
            let col = crate::ooxml::cell::col_to_letters(c);
            let _ = write!(out, r#"<c r="{col}1" t="s"><v>0</v></c>"#);
        }
        out
    };
    let sst = format!(
        r#"<sst xmlns="{SPREADSHEET_NS}" count="1" uniqueCount="1"><si><t>{entry}</t></si></sst>"#
    );
    let open_row = format!(
        r#"<worksheet xmlns="{SPREADSHEET_NS}"><sheetData><row r="1">{}"#,
        cells(40)
    );
    let full_row = format!(
        r#"<worksheet xmlns="{SPREADSHEET_NS}"><sheetData><row r="1">{}</row></sheetData></worksheet>"#,
        cells(30)
    );
    let xlsx = with_part(
        &build_multi_sheet_xlsx(2),
        "xl/sharedStrings.xml",
        sst.as_bytes(),
    );
    let xlsx = with_part(&xlsx, "xl/worksheets/sheet1.xml", open_row.as_bytes());
    let xlsx = with_part(&xlsx, "xl/worksheets/sheet2.xml", full_row.as_bytes());

    let limits = json_limit(60_000);
    let wb = read_xlsx_with_options(&xlsx, &limits).unwrap();
    assert!(wb.sheets[0].worksheet.rows.is_empty());
    assert_eq!(wb.sheets[1].worksheet.rows[0].cells.len(), 30);
    let json = read_xlsx_json_with_options(&xlsx, &limits).unwrap();
    assert_eq!(json, read_xlsx_json(&xlsx).unwrap());

    // Both rows complete: 70,000 bytes of copies pass the budget.
    let closed = format!("{open_row}</row></sheetData></worksheet>");
    let xlsx = with_part(&xlsx, "xl/worksheets/sheet1.xml", closed.as_bytes());
    let err = read_xlsx_with_options(&xlsx, &limits).unwrap_err();
    assert!(err.to_string().contains("shared text"), "{err}");
    assert_eq!(
        read_xlsx_json_with_options(&xlsx, &limits)
            .unwrap_err()
            .code(),
        "SIZE_LIMIT"
    );
}

/// The JSON path stops at the first row that takes the JSON past the
/// limit, not at the end of the sheet.
#[test]
fn the_json_limit_is_checked_after_each_row() {
    let mut rows = String::new();
    for r in 1..=2_000 {
        let _ = write!(rows, r#"<row r="{r}"><c r="A{r}"><v>{r}</v></c></row>"#);
    }
    let sheet =
        format!(r#"<worksheet xmlns="{SPREADSHEET_NS}"><sheetData>{rows}</sheetData></worksheet>"#);
    let xlsx = with_part(
        &build_minimal_xlsx(),
        "xl/worksheets/sheet1.xml",
        sheet.as_bytes(),
    );
    let full = read_xlsx_json(&xlsx).unwrap().len();
    let err = read_xlsx_json_with_options(&xlsx, &json_limit(full / 2)).unwrap_err();
    let msg = err.to_string();
    let reached: usize = msg
        .strip_prefix("size limit exceeded: workbook JSON reached ")
        .and_then(|m| m.split(' ').next())
        .and_then(|n| n.parse().ok())
        .unwrap_or_else(|| panic!("{msg}"));
    // One row of this sheet is under 128 bytes of JSON.
    assert!(reached <= full / 2 + 128, "{msg}");
}

/// The `.rels` Excel 16 saves beside the cache definition of a pivot table
/// whose source is a range in another workbook: the `externalLinkPath` the
/// source's `r:id` names, then the records.
const EXCEL_EXTERNAL_SOURCE_CACHE_RELS: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath" Target="e-source.xlsx" TargetMode="External"/><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheRecords" Target="pivotCacheRecords1.xml"/></Relationships>"#;

#[test]
fn a_pivot_cache_keeps_its_other_relationships_through_a_read_and_write() {
    use serde_json::json;

    const CACHE: &str = "xl/pivotCache/pivotCacheDefinition1.xml";
    const CACHE_RELS: &str = "xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels";
    let base = build_optional_parts_xlsx();
    let def = String::from_utf8(entries_under(&base, &[CACHE])[CACHE].clone()).unwrap();
    let def = def.replace(
        r#"<worksheetSource ref="A1:B5" sheet="Sheet1"/>"#,
        r#"<worksheetSource ref="A1:B5" sheet="Sheet1" r:id="rId2"/>"#,
    );
    assert!(def.contains(r#"r:id="rId2""#), "{def}");
    let xlsx = with_part(
        &with_part(&base, CACHE, def.as_bytes()),
        CACHE_RELS,
        EXCEL_EXTERNAL_SOURCE_CACHE_RELS.as_bytes(),
    );

    let wb = read_xlsx(&xlsx).unwrap();
    assert_eq!(wb.pivot_caches.len(), 2);
    let cache = &wb.pivot_caches[0];
    assert_eq!(cache.source.r_id.as_deref(), Some("rId2"));
    // Every relationship but the records one, as read.
    let expected = json!([{
        "id": "rId2",
        "type": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath",
        "target": "e-source.xlsx",
        "targetMode": "External",
    }]);
    assert_eq!(
        serde_json::to_value(&cache.relationships).unwrap(),
        expected
    );
    assert!(wb.pivot_caches[1].relationships.is_empty());
    assert_eq!(wb.pivot_cache_records.len(), 2);
    assert!(!wb.preserved_entries.contains_key(CACHE_RELS));
    // The JSON read carries the same, and nothing for the other cache.
    let json: serde_json::Value = serde_json::from_str(&read_xlsx_json(&xlsx).unwrap()).unwrap();
    assert_eq!(json["pivotCaches"][0]["relationships"], expected);
    assert_eq!(json["pivotCaches"][0]["source"]["rId"], "rId2");
    assert!(json["pivotCaches"][1].get("relationships").is_none());
    assert!(json["pivotCaches"][1]["source"].get("rId").is_none());

    // Written back with their ids; the records keep rId1.
    let once = crate::writer::write_xlsx(&wb).unwrap();
    let parts = entries_under(&once, &["xl/pivotCache/"]);
    let rels = Relationships::parse(&parts[CACHE_RELS]).unwrap();
    assert_eq!(rels.len(), 2);
    let link = rels.get_by_id("rId2").unwrap();
    assert_eq!(
        (link.target.as_ref(), link.target_mode),
        (
            "e-source.xlsx",
            crate::ooxml::relationships::TargetMode::External
        )
    );
    assert_eq!(
        rels.get_by_id("rId1").unwrap().target,
        "pivotCacheRecords1.xml"
    );
    let written = String::from_utf8(parts[CACHE].clone()).unwrap();
    assert!(written.contains(r#" r:id="rId1""#), "{written}");
    assert!(
        written.contains(r#"<worksheetSource ref="A1:B5" sheet="Sheet1" r:id="rId2"/>"#),
        "{written}"
    );

    // A second generation writes the same pivot cache parts.
    let twice = crate::writer::write_xlsx(&read_xlsx(&once).unwrap()).unwrap();
    assert_eq!(entries_under(&twice, &["xl/pivotCache/"]), parts);
}
