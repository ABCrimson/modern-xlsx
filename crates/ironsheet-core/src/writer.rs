//! Full XLSX write orchestrator.
//!
//! Assembles all XML parts from a [`WorkbookData`] struct and zips them into a
//! complete `.xlsx` file.

use serde::Deserialize;

use crate::dates::DateSystem;
use crate::ooxml::{
    content_types::ContentTypes,
    relationships::Relationships,
    shared_strings::SharedStringTableBuilder,
    styles::Styles,
    workbook::{SheetInfo, SheetState, WorkbookXml},
    worksheet::{CellType, WorksheetXml},
};
use crate::zip::writer::{write_zip, ZipEntry};
use crate::{IronsheetError, Result};

// ---------------------------------------------------------------------------
// Data types
// ---------------------------------------------------------------------------

/// Top-level representation of a workbook suitable for writing.
///
/// If `crate::reader` is available in the future, this type may be replaced by
/// or aliased to the reader's version. For now it is defined locally so the
/// write orchestrator can be developed independently.
#[derive(Debug, Clone, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct WorkbookData {
    /// The sheets in this workbook.
    pub sheets: Vec<SheetData>,
    /// The date system used (1900 or 1904).
    pub date_system: DateSystem,
    /// The styles object for this workbook.
    pub styles: Styles,
}

/// A single sheet inside a workbook.
#[derive(Debug, Clone, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct SheetData {
    /// The user-visible sheet name.
    pub name: String,
    /// The parsed worksheet content (rows, cells, merges, etc.).
    pub worksheet: WorksheetXml,
}

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
        return Err(IronsheetError::InvalidCellValue(
            "workbook must contain at least one sheet".into(),
        ));
    }

    let sheet_count = workbook.sheets.len();

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
                            return Err(IronsheetError::InvalidCellValue(format!(
                                "SharedString cell {} has no value",
                                cell.reference
                            )));
                        }
                    }
                }
            }
        }
    }

    // 2. Generate XML parts.
    let content_types = ContentTypes::for_basic_workbook(sheet_count);
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
        defined_names: Vec::new(),
    };

    // 3. Assemble ZIP entries.
    let mut entries = Vec::with_capacity(6 + sheet_count);

    entries.push(ZipEntry {
        name: "[Content_Types].xml".to_string(),
        data: content_types.to_xml()?.into_bytes(),
    });
    entries.push(ZipEntry {
        name: "_rels/.rels".to_string(),
        data: root_rels.to_xml()?.into_bytes(),
    });
    entries.push(ZipEntry {
        name: "xl/workbook.xml".to_string(),
        data: wb_xml.to_xml()?.into_bytes(),
    });
    entries.push(ZipEntry {
        name: "xl/_rels/workbook.xml.rels".to_string(),
        data: wb_rels.to_xml()?.into_bytes(),
    });
    entries.push(ZipEntry {
        name: "xl/sharedStrings.xml".to_string(),
        data: sst_builder.to_xml()?.into_bytes(),
    });
    entries.push(ZipEntry {
        name: "xl/styles.xml".to_string(),
        data: workbook.styles.to_xml()?.into_bytes(),
    });

    // 4. Generate each worksheet, replacing string values with SST indices.
    for (i, sheet) in workbook.sheets.iter().enumerate() {
        let mut ws_clone = sheet.worksheet.clone();
        remap_sst_indices(&mut ws_clone, &sst_builder);
        let ws_xml = ws_clone.to_xml()?;
        entries.push(ZipEntry {
            name: format!("xl/worksheets/sheet{}.xml", i + 1),
            data: ws_xml.into_bytes(),
        });
    }

    write_zip(&entries)
}

// ---------------------------------------------------------------------------
// Internal helpers
// ---------------------------------------------------------------------------

/// Mutate a worksheet in place, replacing each `SharedString` cell's value
/// (a raw string) with the corresponding SST index from the builder.
fn remap_sst_indices(ws: &mut WorksheetXml, sst: &SharedStringTableBuilder) {
    for row in &mut ws.rows {
        for cell in &mut row.cells {
            if cell.cell_type == CellType::SharedString {
                if let Some(ref val) = cell.value {
                    let idx = sst.get_index(val).expect("BUG: SharedString value not found in SST");
                    cell.value = Some(idx.to_string());
                }
            }
        }
    }
}

// ---------------------------------------------------------------------------
// Tests
// ---------------------------------------------------------------------------

#[cfg(test)]
mod tests {
    use super::*;
    use crate::ooxml::shared_strings::SharedStringTable;
    use crate::ooxml::styles::Styles;
    use crate::ooxml::workbook::WorkbookXml as WbXml;
    use crate::ooxml::worksheet::{Cell, Row, WorksheetXml};
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
                },
            }],
            date_system: DateSystem::Date1900,
            styles: Styles::default_styles(),
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
                },
            }],
            date_system: DateSystem::Date1900,
            styles: Styles::default_styles(),
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
                    },
                },
            ],
            date_system: DateSystem::Date1900,
            styles: Styles::default_styles(),
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
                },
                Cell {
                    reference: "B1".to_string(),
                    cell_type: CellType::SharedString,
                    style_index: None,
                    value: Some("World".to_string()),
                    formula: None,
                },
            ],
            height: None,
            hidden: false,
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
                },
                Cell {
                    reference: "B1".to_string(),
                    cell_type: CellType::Number,
                    style_index: None,
                    value: Some("3.14".to_string()),
                    formula: None,
                },
            ],
            height: None,
            hidden: false,
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
                },
                Cell {
                    reference: "B1".to_string(),
                    cell_type: CellType::Boolean,
                    style_index: None,
                    value: Some("0".to_string()),
                    formula: None,
                },
            ],
            height: None,
            hidden: false,
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
                    },
                    Cell {
                        reference: "B1".to_string(),
                        cell_type: CellType::SharedString,
                        style_index: None,
                        value: Some("Value".to_string()),
                        formula: None,
                    },
                    Cell {
                        reference: "C1".to_string(),
                        cell_type: CellType::SharedString,
                        style_index: None,
                        value: Some("Active".to_string()),
                        formula: None,
                    },
                ],
                height: None,
                hidden: false,
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
                    },
                    Cell {
                        reference: "B2".to_string(),
                        cell_type: CellType::Number,
                        style_index: None,
                        value: Some("3.14159".to_string()),
                        formula: None,
                    },
                    Cell {
                        reference: "C2".to_string(),
                        cell_type: CellType::Boolean,
                        style_index: None,
                        value: Some("1".to_string()),
                        formula: None,
                    },
                ],
                height: None,
                hidden: false,
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
                }],
                height: None,
                hidden: false,
            },
            Row {
                index: 2,
                cells: vec![Cell {
                    reference: "A2".to_string(),
                    cell_type: CellType::SharedString,
                    style_index: None,
                    value: Some("Repeat".to_string()),
                    formula: None,
                }],
                height: None,
                hidden: false,
            },
            Row {
                index: 3,
                cells: vec![Cell {
                    reference: "A3".to_string(),
                    cell_type: CellType::SharedString,
                    style_index: None,
                    value: Some("Other".to_string()),
                    formula: None,
                }],
                height: None,
                hidden: false,
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
                    },
                },
            ],
            date_system: DateSystem::Date1900,
            styles: Styles::default_styles(),
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
    fn test_remap_sst_indices_helper() {
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
                    },
                    Cell {
                        reference: "B1".to_string(),
                        cell_type: CellType::Number,
                        style_index: Some(1),
                        value: Some("99".to_string()),
                        formula: None,
                    },
                ],
                height: None,
                hidden: false,
            }],
            merge_cells: Vec::new(),
            auto_filter: None,
            frozen_pane: None,
            columns: Vec::new(),
        };

        let mut sst = SharedStringTableBuilder::new();
        // Insert some dummy strings first to push "Alpha" to index 7.
        for i in 0..7 {
            sst.insert(&format!("dummy{i}"));
        }
        sst.insert("Alpha"); // index 7

        let mut ws_clone = ws.clone();
        remap_sst_indices(&mut ws_clone, &sst);

        // The SharedString cell should now have SST index as its value.
        assert_eq!(ws_clone.rows[0].cells[0].value.as_deref(), Some("7"));
        assert_eq!(ws_clone.rows[0].cells[0].cell_type, CellType::SharedString);

        // The Number cell should be unchanged.
        assert_eq!(ws_clone.rows[0].cells[1].value.as_deref(), Some("99"));
        assert_eq!(ws_clone.rows[0].cells[1].cell_type, CellType::Number);
        assert_eq!(ws_clone.rows[0].cells[1].style_index, Some(1));
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
                    },
                },
            ],
            date_system: DateSystem::Date1900,
            styles: Styles::default_styles(),
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
                    },
                    Cell {
                        reference: "B1".to_string(),
                        cell_type: CellType::Number,
                        style_index: None,
                        value: Some("100".to_string()),
                        formula: None,
                    },
                    Cell {
                        reference: "C1".to_string(),
                        cell_type: CellType::Boolean,
                        style_index: None,
                        value: Some("1".to_string()),
                        formula: None,
                    },
                ],
                height: None,
                hidden: false,
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
        assert_eq!(read_wb.shared_strings.len(), 1);
        assert_eq!(read_wb.shared_strings.get(0), Some("Header"));

        // Worksheet should have 1 row with 3 cells.
        let ws = &read_wb.sheets[0].worksheet;
        assert_eq!(ws.rows.len(), 1);
        assert_eq!(ws.rows[0].cells.len(), 3);

        // A1: SharedString with SST index "0" -> "Header".
        assert_eq!(ws.rows[0].cells[0].cell_type, CellType::SharedString);
        assert_eq!(ws.rows[0].cells[0].value.as_deref(), Some("0"));

        // B1: Number "100".
        assert_eq!(ws.rows[0].cells[1].cell_type, CellType::Number);
        assert_eq!(ws.rows[0].cells[1].value.as_deref(), Some("100"));

        // C1: Boolean "1".
        assert_eq!(ws.rows[0].cells[2].cell_type, CellType::Boolean);
        assert_eq!(ws.rows[0].cells[2].value.as_deref(), Some("1"));
    }
}
