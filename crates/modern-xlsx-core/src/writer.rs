//! Full XLSX write orchestrator.
//!
//! Assembles all XML parts from a [`WorkbookData`] struct and zips them into a
//! complete `.xlsx` file.

use std::collections::HashSet;

use log::{debug, trace};

use crate::ooxml::{
    calc_chain,
    comments,
    content_types::{ContentTypes, CT_COMMENTS},
    relationships::{Relationships, REL_COMMENTS},
    shared_strings::SharedStringTableBuilder,
    workbook::{SheetInfo, SheetState, WorkbookXml},
    worksheet::CellType,
};
use crate::zip::writer::{write_zip, ZipEntry};
use crate::{ModernXlsxError, Result, WorkbookData};

// ---------------------------------------------------------------------------
// Public API
// ---------------------------------------------------------------------------

/// Write a complete XLSX file from a [`WorkbookData`] struct.
///
/// For cells whose [`CellType`] is [`CellType::SharedString`], the `value`
/// field must contain the **actual string text** (not an SST index). This
/// function builds the shared string table, maps each string to an index,
/// and writes the correct index into the worksheet XML.
///
/// Returns the raw bytes of the resulting ZIP (`.xlsx`) file.
pub fn write_xlsx(workbook: &WorkbookData) -> Result<Vec<u8>> {
    if workbook.sheets.is_empty() {
        return Err(ModernXlsxError::InvalidCellValue(
            "workbook must contain at least one sheet".into(),
        ));
    }

    let sheet_count = workbook.sheets.len();
    debug!("collecting shared strings from {sheet_count} sheets");

    // 1. Build the SharedStringTable from all string cells.
    let mut sst_builder = SharedStringTableBuilder::new();

    for sheet in &workbook.sheets {
        for row in &sheet.worksheet.rows {
            for cell in &row.cells {
                if cell.cell_type == CellType::SharedString {
                    match cell.value {
                        Some(ref val) => {
                            sst_builder.insert(val);
                        }
                        None => {
                            return Err(ModernXlsxError::InvalidCellValue(format!(
                                "SharedString cell {} has no value",
                                cell.reference
                            )));
                        }
                    }
                }
            }
        }
    }

    debug!("SST built with {} unique strings", sst_builder.len());

    // 2. Generate XML parts.
    let mut content_types = ContentTypes::for_basic_workbook(sheet_count);
    let root_rels = Relationships::root_rels();
    let wb_rels = Relationships::workbook_rels(sheet_count);

    let wb_xml = WorkbookXml {
        sheets: workbook
            .sheets
            .iter()
            .enumerate()
            .map(|(i, s)| SheetInfo {
                name: s.name.clone(),
                sheet_id: (i + 1) as u32,
                r_id: format!("rId{}", i + 1),
                state: SheetState::Visible,
            })
            .collect(),
        date_system: workbook.date_system,
        defined_names: workbook.defined_names.clone(),
        workbook_views: workbook.workbook_views.clone(),
    };

    // 3. Assemble ZIP entries.
    // NOTE: [Content_Types].xml is added last so we can append comment
    // content-type overrides discovered while writing worksheets.
    let mut entries = Vec::with_capacity(6 + sheet_count);

    entries.push(ZipEntry {
        name: "_rels/.rels".to_string(),
        data: root_rels.to_xml()?,
    });
    entries.push(ZipEntry {
        name: "xl/workbook.xml".to_string(),
        data: wb_xml.to_xml()?,
    });
    entries.push(ZipEntry {
        name: "xl/_rels/workbook.xml.rels".to_string(),
        data: wb_rels.to_xml()?,
    });
    entries.push(ZipEntry {
        name: "xl/sharedStrings.xml".to_string(),
        data: sst_builder.to_xml()?,
    });
    entries.push(ZipEntry {
        name: "xl/styles.xml".to_string(),
        data: workbook.styles.to_xml()?,
    });

    // 4a. Write document properties if present.
    if let Some(ref props) = workbook.doc_properties {
        if props.has_core() {
            entries.push(ZipEntry {
                name: "docProps/core.xml".to_string(),
                data: props.to_core_xml()?,
            });
        }
        if props.has_app() {
            entries.push(ZipEntry {
                name: "docProps/app.xml".to_string(),
                data: props.to_app_xml()?,
            });
        }
    }

    // 4b. Write calculation chain if non-empty.
    if !workbook.calc_chain.is_empty() {
        entries.push(ZipEntry {
            name: "xl/calcChain.xml".to_string(),
            data: calc_chain::to_xml(&workbook.calc_chain)?,
        });
    }

    // 5. Generate each worksheet, remapping SST indices inline during XML
    // serialization (avoids cloning the entire worksheet).
    // Also write comments XML and worksheet .rels files for sheets with comments.
    for (i, sheet) in workbook.sheets.iter().enumerate() {
        let sheet_num = i + 1;
        let ws_xml = sheet.worksheet.to_xml_with_sst(Some(&sst_builder))?;
        entries.push(ZipEntry {
            name: format!("xl/worksheets/sheet{sheet_num}.xml"),
            data: ws_xml,
        });

        // Write comments if the worksheet has any.
        if !sheet.worksheet.comments.is_empty() {
            debug!(
                "writing {} comments for sheet {}",
                sheet.worksheet.comments.len(),
                sheet_num
            );
            let comments_xml = comments::write_comments(&sheet.worksheet.comments)?;
            let comments_path = format!("xl/comments{sheet_num}.xml");

            // Add content type override before moving the path into the entry.
            content_types.add_override(format!("/{comments_path}"), CT_COMMENTS);

            entries.push(ZipEntry {
                name: comments_path,
                data: comments_xml,
            });

            // Create (or merge into) the worksheet .rels file.
            // The comments target is relative to the worksheet directory:
            //   ../comments1.xml (from xl/worksheets/ up to xl/)
            let rels_path = format!("xl/worksheets/_rels/sheet{sheet_num}.xml.rels");

            // Check if a preserved worksheet .rels already exists.
            let mut ws_rels = if let Some(existing) = workbook.preserved_entries.get(&rels_path) {
                Relationships::parse(existing)?
            } else {
                Relationships::new()
            };

            // Only add the comments relationship if not already present.
            let already_has_comments = ws_rels
                .find_by_type(REL_COMMENTS)
                .next()
                .is_some();
            if !already_has_comments {
                // Pick an rId that does not collide with existing ones.
                // Use max existing numeric ID + 1 (not len+1) to avoid
                // collisions when IDs are non-sequential.
                let max_existing_id = ws_rels
                    .relationships
                    .iter()
                    .filter_map(|r| r.id.strip_prefix("rId").and_then(|n| n.parse::<usize>().ok()))
                    .max()
                    .unwrap_or(0);
                let next_id = max_existing_id + 1;
                ws_rels.add(
                    format!("rId{next_id}"),
                    REL_COMMENTS,
                    format!("../comments{sheet_num}.xml"),
                );
            }

            entries.push(ZipEntry {
                name: rels_path,
                data: ws_rels.to_xml()?,
            });
        }
    }

    // 6. Append preserved entries (drawings, media, charts, etc.)
    // Skip worksheet .rels files that we already generated for comments,
    // to avoid duplicate entries in the ZIP.
    let generated_rels: HashSet<String> = workbook
        .sheets
        .iter()
        .enumerate()
        .filter(|(_, s)| !s.worksheet.comments.is_empty())
        .map(|(i, _)| format!("xl/worksheets/_rels/sheet{}.xml.rels", i + 1))
        .collect();
    for (path, data) in &workbook.preserved_entries {
        if generated_rels.contains(path) {
            // Already written with comments relationship merged in.
            continue;
        }
        trace!("writing preserved ZIP entry: {}", path);
        entries.push(ZipEntry {
            name: path.clone(),
            data: data.clone(),
        });
    }

    // 6b. Auto-detect content types from preserved entries.
    // Images need extension defaults; drawings need part overrides.
    for path in workbook.preserved_entries.keys() {
        if let Some(ext) = path.rsplit('.').next() {
            match ext {
                "png" => { content_types.add_default("png", "image/png"); }
                "jpeg" | "jpg" => { content_types.add_default("jpeg", "image/jpeg"); }
                "gif" => { content_types.add_default("gif", "image/gif"); }
                "emf" => { content_types.add_default("emf", "image/x-emf"); }
                "wmf" => { content_types.add_default("wmf", "image/x-wmf"); }
                _ => {}
            }
        }
        if path.starts_with("xl/drawings/drawing") && path.ends_with(".xml") {
            content_types.add_override(
                format!("/{path}"),
                "application/vnd.openxmlformats-officedocument.drawing+xml",
            );
        }
        if path.starts_with("xl/charts/chart") && path.ends_with(".xml") {
            content_types.add_override(
                format!("/{path}"),
                "application/vnd.openxmlformats-officedocument.drawingml.chart+xml",
            );
        }
    }

    // 7. Write [Content_Types].xml (must be first entry in the ZIP archive).
    entries.push(ZipEntry {
        name: "[Content_Types].xml".to_string(),
        data: content_types.to_xml()?,
    });
    entries.rotate_right(1);

    for entry in &entries {
        trace!("writing ZIP entry: {}", entry.name);
    }

    write_zip(&entries)
}

// ---------------------------------------------------------------------------
// Tests
// ---------------------------------------------------------------------------

#[cfg(test)]
mod tests {
    use super::*;
    use pretty_assertions::assert_eq;
    use crate::SheetData;
    use crate::dates::DateSystem;
    use crate::ooxml::shared_strings::SharedStringTable;
    use crate::ooxml::styles::Styles;
    use crate::ooxml::workbook::WorkbookXml as WbXml;
    use crate::ooxml::worksheet::{Cell, CellType, Row, WorksheetXml};
    use crate::zip::reader::{read_zip_entries, ZipSecurityLimits};

    /// Helper: build a minimal single-sheet workbook with the given rows.
    fn minimal_workbook(name: &str, rows: Vec<Row>) -> WorkbookData {
        WorkbookData {
            sheets: vec![SheetData {
                name: name.to_string(),
                worksheet: WorksheetXml {
                    dimension: Some("A1".to_string()),
                    rows,
                    merge_cells: Vec::new(),
                    auto_filter: None,
                    frozen_pane: None,
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
            preserved_entries: std::collections::BTreeMap::new(),
        }
    }

    #[test]
    fn test_write_empty_workbook() {
        let wb = WorkbookData {
            sheets: vec![SheetData {
                name: "Sheet1".to_string(),
                worksheet: WorksheetXml {
                    dimension: None,
                    rows: Vec::new(),
                    merge_cells: Vec::new(),
                    auto_filter: None,
                    frozen_pane: None,
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
            preserved_entries: std::collections::BTreeMap::new(),
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
        let wb = WorkbookData {
            sheets: Vec::new(),
            date_system: DateSystem::Date1900,
            styles: Styles::default_styles(),
            defined_names: Vec::new(),
            shared_strings: None,
            doc_properties: None,
            theme_colors: None,
            calc_chain: Vec::new(),
            workbook_views: Vec::new(),
            preserved_entries: std::collections::BTreeMap::new(),
        };
        let result = write_xlsx(&wb);
        assert!(result.is_err());
    }

    #[test]
    fn test_zip_contains_expected_entries_for_two_sheets() {
        let wb = WorkbookData {
            sheets: vec![
                SheetData {
                    name: "First".to_string(),
                    worksheet: WorksheetXml {
                        dimension: None,
                        rows: Vec::new(),
                        merge_cells: Vec::new(),
                        auto_filter: None,
                        frozen_pane: None,
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
                    },
                },
                SheetData {
                    name: "Second".to_string(),
                    worksheet: WorksheetXml {
                        dimension: None,
                        rows: Vec::new(),
                        merge_cells: Vec::new(),
                        auto_filter: None,
                        frozen_pane: None,
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
                    },
                },
            ],
            date_system: DateSystem::Date1900,
            styles: Styles::default_styles(),
            defined_names: Vec::new(),
            shared_strings: None,
            doc_properties: None,
            theme_colors: None,
            calc_chain: Vec::new(),
            workbook_views: Vec::new(),
            preserved_entries: std::collections::BTreeMap::new(),
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
                    formula: None,
                    formula_type: None,
                    formula_ref: None,
                    shared_index: None,
                    inline_string: None,
                    dynamic_array: None,
                },
                Cell {
                    reference: "B1".to_string(),
                    cell_type: CellType::SharedString,
                    style_index: None,
                    value: Some("World".to_string()),
                    formula: None,
                    formula_type: None,
                    formula_ref: None,
                    shared_index: None,
                    inline_string: None,
                    dynamic_array: None,
                },
            ],
            height: None,
            hidden: false,
            outline_level: None,
            collapsed: false,
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
                    formula: None,
                    formula_type: None,
                    formula_ref: None,
                    shared_index: None,
                    inline_string: None,
                    dynamic_array: None,
                },
                Cell {
                    reference: "B1".to_string(),
                    cell_type: CellType::Number,
                    style_index: None,
                    value: Some("3.14".to_string()),
                    formula: None,
                    formula_type: None,
                    formula_ref: None,
                    shared_index: None,
                    inline_string: None,
                    dynamic_array: None,
                },
            ],
            height: None,
            hidden: false,
            outline_level: None,
            collapsed: false,
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
                    formula: None,
                    formula_type: None,
                    formula_ref: None,
                    shared_index: None,
                    inline_string: None,
                    dynamic_array: None,
                },
                Cell {
                    reference: "B1".to_string(),
                    cell_type: CellType::Boolean,
                    style_index: None,
                    value: Some("0".to_string()),
                    formula: None,
                    formula_type: None,
                    formula_ref: None,
                    shared_index: None,
                    inline_string: None,
                    dynamic_array: None,
                },
            ],
            height: None,
            hidden: false,
            outline_level: None,
            collapsed: false,
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
                        formula: None,
                        formula_type: None,
                        formula_ref: None,
                        shared_index: None,
                        inline_string: None,
                        dynamic_array: None,
                    },
                    Cell {
                        reference: "B1".to_string(),
                        cell_type: CellType::SharedString,
                        style_index: None,
                        value: Some("Value".to_string()),
                        formula: None,
                        formula_type: None,
                        formula_ref: None,
                        shared_index: None,
                        inline_string: None,
                        dynamic_array: None,
                    },
                    Cell {
                        reference: "C1".to_string(),
                        cell_type: CellType::SharedString,
                        style_index: None,
                        value: Some("Active".to_string()),
                        formula: None,
                        formula_type: None,
                        formula_ref: None,
                        shared_index: None,
                        inline_string: None,
                        dynamic_array: None,
                    },
                ],
                height: None,
                hidden: false,
                outline_level: None,
                collapsed: false,
            },
            Row {
                index: 2,
                cells: vec![
                    Cell {
                        reference: "A2".to_string(),
                        cell_type: CellType::SharedString,
                        style_index: None,
                        value: Some("Pi".to_string()),
                        formula: None,
                        formula_type: None,
                        formula_ref: None,
                        shared_index: None,
                        inline_string: None,
                        dynamic_array: None,
                    },
                    Cell {
                        reference: "B2".to_string(),
                        cell_type: CellType::Number,
                        style_index: None,
                        value: Some("3.14159".to_string()),
                        formula: None,
                        formula_type: None,
                        formula_ref: None,
                        shared_index: None,
                        inline_string: None,
                        dynamic_array: None,
                    },
                    Cell {
                        reference: "C2".to_string(),
                        cell_type: CellType::Boolean,
                        style_index: None,
                        value: Some("1".to_string()),
                        formula: None,
                        formula_type: None,
                        formula_ref: None,
                        shared_index: None,
                        inline_string: None,
                        dynamic_array: None,
                    },
                ],
                height: None,
                hidden: false,
                outline_level: None,
                collapsed: false,
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
                    formula: None,
                    formula_type: None,
                    formula_ref: None,
                    shared_index: None,
                    inline_string: None,
                    dynamic_array: None,
                }],
                height: None,
                hidden: false,
                outline_level: None,
                collapsed: false,
            },
            Row {
                index: 2,
                cells: vec![Cell {
                    reference: "A2".to_string(),
                    cell_type: CellType::SharedString,
                    style_index: None,
                    value: Some("Repeat".to_string()),
                    formula: None,
                    formula_type: None,
                    formula_ref: None,
                    shared_index: None,
                    inline_string: None,
                    dynamic_array: None,
                }],
                height: None,
                hidden: false,
                outline_level: None,
                collapsed: false,
            },
            Row {
                index: 3,
                cells: vec![Cell {
                    reference: "A3".to_string(),
                    cell_type: CellType::SharedString,
                    style_index: None,
                    value: Some("Other".to_string()),
                    formula: None,
                    formula_type: None,
                    formula_ref: None,
                    shared_index: None,
                    inline_string: None,
                    dynamic_array: None,
                }],
                height: None,
                hidden: false,
                outline_level: None,
                collapsed: false,
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
    fn test_workbook_xml_contains_sheet_names() {
        let wb = WorkbookData {
            sheets: vec![
                SheetData {
                    name: "Sales".to_string(),
                    worksheet: WorksheetXml {
                        dimension: None,
                        rows: Vec::new(),
                        merge_cells: Vec::new(),
                        auto_filter: None,
                        frozen_pane: None,
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
                    },
                },
                SheetData {
                    name: "Inventory".to_string(),
                    worksheet: WorksheetXml {
                        dimension: None,
                        rows: Vec::new(),
                        merge_cells: Vec::new(),
                        auto_filter: None,
                        frozen_pane: None,
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
                    },
                },
            ],
            date_system: DateSystem::Date1900,
            styles: Styles::default_styles(),
            defined_names: Vec::new(),
            shared_strings: None,
            doc_properties: None,
            theme_colors: None,
            calc_chain: Vec::new(),
            workbook_views: Vec::new(),
            preserved_entries: std::collections::BTreeMap::new(),
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
                        formula: None,
                        formula_type: None,
                        formula_ref: None,
                        shared_index: None,
                        inline_string: None,
                        dynamic_array: None,
                    },
                    Cell {
                        reference: "B1".to_string(),
                        cell_type: CellType::Number,
                        style_index: Some(1),
                        value: Some("99".to_string()),
                        formula: None,
                        formula_type: None,
                        formula_ref: None,
                        shared_index: None,
                        inline_string: None,
                        dynamic_array: None,
                    },
                ],
                height: None,
                hidden: false,
                outline_level: None,
                collapsed: false,
            }],
            merge_cells: Vec::new(),
            auto_filter: None,
            frozen_pane: None,
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
        };

        let mut sst = SharedStringTableBuilder::new();
        // Insert some dummy strings first to push "Alpha" to index 7.
        for i in 0..7 {
            sst.insert(&format!("dummy{i}"));
        }
        sst.insert("Alpha"); // index 7

        // Generate XML with inline SST remapping (no clone needed).
        let xml_bytes = ws.to_xml_with_sst(Some(&sst)).expect("to_xml_with_sst should succeed");
        let xml_str = std::str::from_utf8(&xml_bytes).unwrap();

        // The SharedString cell's <v> should contain the SST index "7", not "Alpha".
        assert!(xml_str.contains("<v>7</v>"), "SST cell should have index 7, got: {}", xml_str);
        // The Number cell should be unchanged.
        assert!(xml_str.contains("<v>99</v>"), "Number cell should remain 99");
    }

    #[test]
    fn test_content_types_has_correct_overrides() {
        let wb = minimal_workbook("Sheet1", Vec::new());
        let bytes = write_xlsx(&wb).expect("write_xlsx should succeed");

        let limits = ZipSecurityLimits::default();
        let entries = read_zip_entries(&bytes, &limits).unwrap();

        let ct_xml =
            std::str::from_utf8(entries.get("[Content_Types].xml").unwrap()).unwrap();
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
                    worksheet: WorksheetXml {
                        dimension: None,
                        rows: Vec::new(),
                        merge_cells: Vec::new(),
                        auto_filter: None,
                        frozen_pane: None,
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
                    },
                },
                SheetData {
                    name: "S2".to_string(),
                    worksheet: WorksheetXml {
                        dimension: None,
                        rows: Vec::new(),
                        merge_cells: Vec::new(),
                        auto_filter: None,
                        frozen_pane: None,
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
                    },
                },
            ],
            date_system: DateSystem::Date1900,
            styles: Styles::default_styles(),
            defined_names: Vec::new(),
            shared_strings: None,
            doc_properties: None,
            theme_colors: None,
            calc_chain: Vec::new(),
            workbook_views: Vec::new(),
            preserved_entries: std::collections::BTreeMap::new(),
        };

        let bytes = write_xlsx(&wb).expect("write_xlsx should succeed");
        let limits = ZipSecurityLimits::default();
        let entries = read_zip_entries(&bytes, &limits).unwrap();

        // Root rels should point to xl/workbook.xml.
        let root_rels_xml =
            std::str::from_utf8(entries.get("_rels/.rels").unwrap()).unwrap();
        let root_rels = Relationships::parse(root_rels_xml.as_bytes()).unwrap();
        assert_eq!(root_rels.relationships.len(), 1);
        assert_eq!(root_rels.relationships[0].target, "xl/workbook.xml");

        // Workbook rels should have 2 worksheets + sharedStrings + styles = 4.
        let wb_rels_xml =
            std::str::from_utf8(entries.get("xl/_rels/workbook.xml.rels").unwrap()).unwrap();
        let wb_rels = Relationships::parse(wb_rels_xml.as_bytes()).unwrap();
        assert_eq!(wb_rels.relationships.len(), 4);
        assert_eq!(wb_rels.get_by_id("rId1").unwrap().target, "worksheets/sheet1.xml");
        assert_eq!(wb_rels.get_by_id("rId2").unwrap().target, "worksheets/sheet2.xml");
        assert_eq!(wb_rels.get_by_id("rId3").unwrap().target, "sharedStrings.xml");
        assert_eq!(wb_rels.get_by_id("rId4").unwrap().target, "styles.xml");
    }

    #[test]
    fn test_write_then_read_roundtrip() {
        // Write a workbook with mixed cell types, then read it back via the
        // reader module and verify the data survived the round-trip.
        let rows = vec![
            Row {
                index: 1,
                cells: vec![
                    Cell {
                        reference: "A1".to_string(),
                        cell_type: CellType::SharedString,
                        style_index: None,
                        value: Some("Header".to_string()),
                        formula: None,
                        formula_type: None,
                        formula_ref: None,
                        shared_index: None,
                        inline_string: None,
                        dynamic_array: None,
                    },
                    Cell {
                        reference: "B1".to_string(),
                        cell_type: CellType::Number,
                        style_index: None,
                        value: Some("100".to_string()),
                        formula: None,
                        formula_type: None,
                        formula_ref: None,
                        shared_index: None,
                        inline_string: None,
                        dynamic_array: None,
                    },
                    Cell {
                        reference: "C1".to_string(),
                        cell_type: CellType::Boolean,
                        style_index: None,
                        value: Some("1".to_string()),
                        formula: None,
                        formula_type: None,
                        formula_ref: None,
                        shared_index: None,
                        inline_string: None,
                        dynamic_array: None,
                    },
                ],
                height: None,
                hidden: false,
                outline_level: None,
                collapsed: false,
            },
        ];

        let wb = minimal_workbook("RoundTrip", rows);
        let xlsx_bytes = write_xlsx(&wb).expect("write_xlsx should succeed");

        // Read back via the reader module.
        let read_wb = crate::reader::read_xlsx(&xlsx_bytes)
            .expect("read_xlsx should succeed on writer output");

        assert_eq!(read_wb.sheets.len(), 1);
        assert_eq!(read_wb.sheets[0].name, "RoundTrip");
        assert_eq!(read_wb.date_system, DateSystem::Date1900);

        // Shared strings table should have 1 entry: "Header".
        let sst = read_wb.shared_strings.as_ref().expect("shared_strings should be Some");
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
                    formula: None,
                    formula_type: None,
                    formula_ref: None,
                    shared_index: None,
                    inline_string: None,
                    dynamic_array: None,
                }],
                height: None,
                hidden: false,
                outline_level: None,
                collapsed: false,
            },
            Row {
                index: 2,
                cells: vec![Cell {
                    reference: "A2".to_string(),
                    cell_type: CellType::Number,
                    style_index: None,
                    value: Some("20".to_string()),
                    formula: None,
                    formula_type: None,
                    formula_ref: None,
                    shared_index: None,
                    inline_string: None,
                    dynamic_array: None,
                }],
                height: None,
                hidden: false,
                outline_level: None,
                collapsed: false,
            },
            Row {
                index: 3,
                cells: vec![Cell {
                    reference: "A3".to_string(),
                    cell_type: CellType::Number,
                    style_index: None,
                    value: Some("30".to_string()),
                    formula: Some("SUM(A1:A2)".to_string()),
                    formula_type: None,
                    formula_ref: None,
                    shared_index: None,
                    inline_string: None,
                    dynamic_array: None,
                }],
                height: None,
                hidden: false,
                outline_level: None,
                collapsed: false,
            },
        ];

        let wb = minimal_workbook("Formulas", rows);
        let xlsx_bytes = write_xlsx(&wb).expect("write_xlsx should succeed");

        // Read back via the reader module.
        let read_wb = crate::reader::read_xlsx(&xlsx_bytes)
            .expect("read_xlsx should succeed on writer output");

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
        assert_eq!(
            ws.rows[2].cells[0].formula,
            Some("SUM(A1:A2)".to_string())
        );
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
                style_index: None,
                value: Some("100".to_string()),
                formula: Some("A1*2".to_string()),
                formula_type: None,
                formula_ref: None,
                shared_index: None,
                inline_string: None,
                dynamic_array: None,
            }],
            height: None,
            hidden: false,
            outline_level: None,
            collapsed: false,
        }];

        let wb = minimal_workbook("FormulaStr", rows);
        let xlsx_bytes = write_xlsx(&wb).expect("write_xlsx should succeed");

        let read_wb = crate::reader::read_xlsx(&xlsx_bytes)
            .expect("read_xlsx should succeed on writer output");

        let ws = &read_wb.sheets[0].worksheet;
        assert_eq!(ws.rows.len(), 1);
        assert_eq!(ws.rows[0].cells[0].reference, "B1");
        assert_eq!(ws.rows[0].cells[0].cell_type, CellType::FormulaStr);
        assert_eq!(
            ws.rows[0].cells[0].formula,
            Some("A1*2".to_string())
        );
        assert_eq!(ws.rows[0].cells[0].value.as_deref(), Some("100"));
    }
}
