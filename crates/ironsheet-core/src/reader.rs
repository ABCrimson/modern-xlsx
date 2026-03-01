//! Full XLSX read orchestrator.
//!
//! Decompresses a `.xlsx` ZIP archive, parses all OPC / SpreadsheetML parts,
//! and assembles them into a [`WorkbookData`] struct.

use serde::Serialize;

use crate::dates::DateSystem;
use crate::ooxml::{
    relationships::Relationships,
    shared_strings::SharedStringTable,
    styles::Styles,
    workbook::WorkbookXml,
    worksheet::WorksheetXml,
};
use crate::zip::reader::{read_zip_entries, ZipSecurityLimits};
use crate::{IronsheetError, Result};

/// Complete parsed workbook data.
#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct WorkbookData {
    pub sheets: Vec<SheetData>,
    pub date_system: DateSystem,
    pub styles: Styles,
    pub shared_strings: SharedStringTable,
}

/// A single parsed worksheet with its tab name.
#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct SheetData {
    pub name: String,
    pub worksheet: WorksheetXml,
}

/// Read an XLSX file from bytes using default security limits.
pub fn read_xlsx(data: &[u8]) -> Result<WorkbookData> {
    read_xlsx_with_options(data, &ZipSecurityLimits::default())
}

/// Read an XLSX file from bytes with custom ZIP security limits.
pub fn read_xlsx_with_options(data: &[u8], limits: &ZipSecurityLimits) -> Result<WorkbookData> {
    let entries = read_zip_entries(data, limits)?;

    // Parse workbook (required part).
    let workbook_data = entries
        .get("xl/workbook.xml")
        .ok_or_else(|| IronsheetError::MissingPart("xl/workbook.xml".into()))?;
    let workbook_xml = WorkbookXml::parse(workbook_data)?;

    // Parse shared strings (optional — some workbooks have no strings).
    let sst = entries
        .get("xl/sharedStrings.xml")
        .map(|d| SharedStringTable::parse(d))
        .transpose()?
        .unwrap_or_else(SharedStringTable::empty);

    // Parse styles (optional — use defaults when absent).
    let styles = entries
        .get("xl/styles.xml")
        .map(|d| Styles::parse(d))
        .transpose()?
        .unwrap_or_else(Styles::default_styles);

    // Parse workbook relationships (optional — needed to resolve sheet targets).
    let wb_rels = entries
        .get("xl/_rels/workbook.xml.rels")
        .map(|d| Relationships::parse(d))
        .transpose()?
        .unwrap_or_else(Relationships::new);

    // Parse each worksheet referenced in the workbook.
    let mut sheets = Vec::new();
    for sheet_info in &workbook_xml.sheets {
        let rel = wb_rels.get_by_id(&sheet_info.r_id).ok_or_else(|| {
            IronsheetError::MissingPart(format!(
                "relationship {} for sheet '{}'",
                sheet_info.r_id, sheet_info.name
            ))
        })?;

        // Normalize target path: handle both relative ("worksheets/sheet1.xml")
        // and absolute ("/xl/worksheets/sheet1.xml") targets.
        let sheet_path = if rel.target.starts_with('/') {
            rel.target.trim_start_matches('/').to_string()
        } else {
            format!("xl/{}", rel.target)
        };

        let sheet_data = entries.get(&sheet_path).ok_or_else(|| {
            IronsheetError::MissingPart(format!("{} for sheet '{}'", sheet_path, sheet_info.name))
        })?;

        let worksheet = WorksheetXml::parse(sheet_data)?;
        sheets.push(SheetData {
            name: sheet_info.name.clone(),
            worksheet,
        });
    }

    Ok(WorkbookData {
        sheets,
        date_system: workbook_xml.date_system,
        styles,
        shared_strings: sst,
    })
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::ooxml::{
        content_types::ContentTypes,
        relationships::Relationships,
        shared_strings::SharedStringTableBuilder,
        styles::Styles,
        workbook::{SheetInfo, SheetState, WorkbookXml},
        worksheet::{Cell, CellType, Row, WorksheetXml},
    };
    use crate::zip::writer::{write_zip, ZipEntry};

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
                        formula: None,
                    },
                    Cell {
                        reference: "B1".to_string(),
                        cell_type: CellType::Number,
                        style_index: None,
                        value: Some("42".to_string()),
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
        let ws_xml = ws.to_xml().unwrap();

        let entries = vec![
            ZipEntry {
                name: "[Content_Types].xml".to_string(),
                data: ct_xml.into_bytes(),
            },
            ZipEntry {
                name: "_rels/.rels".to_string(),
                data: root_rels_xml.into_bytes(),
            },
            ZipEntry {
                name: "xl/workbook.xml".to_string(),
                data: wb_xml.into_bytes(),
            },
            ZipEntry {
                name: "xl/_rels/workbook.xml.rels".to_string(),
                data: wb_rels_xml.into_bytes(),
            },
            ZipEntry {
                name: "xl/sharedStrings.xml".to_string(),
                data: sst_xml.into_bytes(),
            },
            ZipEntry {
                name: "xl/styles.xml".to_string(),
                data: styles_xml.into_bytes(),
            },
            ZipEntry {
                name: "xl/worksheets/sheet1.xml".to_string(),
                data: ws_xml.into_bytes(),
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
                    formula: None,
                }],
                height: None,
                hidden: false,
            }],
            merge_cells: Vec::new(),
            auto_filter: None,
            frozen_pane: None,
            columns: Vec::new(),
        };
        let ws_xml = ws.to_xml().unwrap();

        let entries = vec![
            ZipEntry {
                name: "[Content_Types].xml".to_string(),
                data: ct_xml.into_bytes(),
            },
            ZipEntry {
                name: "_rels/.rels".to_string(),
                data: root_rels_xml.into_bytes(),
            },
            ZipEntry {
                name: "xl/workbook.xml".to_string(),
                data: wb_xml.into_bytes(),
            },
            ZipEntry {
                name: "xl/_rels/workbook.xml.rels".to_string(),
                data: wb_rels_xml.into_bytes(),
            },
            ZipEntry {
                name: "xl/styles.xml".to_string(),
                data: styles_xml.into_bytes(),
            },
            ZipEntry {
                name: "xl/worksheets/sheet1.xml".to_string(),
                data: ws_xml.into_bytes(),
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
            dimension: None,
            rows: Vec::new(),
            merge_cells: Vec::new(),
            auto_filter: None,
            frozen_pane: None,
            columns: Vec::new(),
        };
        let ws_xml = ws.to_xml().unwrap();

        let entries = vec![
            ZipEntry {
                name: "[Content_Types].xml".to_string(),
                data: ct_xml.into_bytes(),
            },
            ZipEntry {
                name: "_rels/.rels".to_string(),
                data: root_rels_xml.into_bytes(),
            },
            ZipEntry {
                name: "xl/workbook.xml".to_string(),
                data: wb_xml.into_bytes(),
            },
            ZipEntry {
                name: "xl/_rels/workbook.xml.rels".to_string(),
                data: wb_rels_xml.into_bytes(),
            },
            ZipEntry {
                name: "xl/sharedStrings.xml".to_string(),
                data: sst_xml.into_bytes(),
            },
            ZipEntry {
                name: "xl/styles.xml".to_string(),
                data: styles_xml.into_bytes(),
            },
            ZipEntry {
                name: "xl/worksheets/sheet1.xml".to_string(),
                data: ws_xml.into_bytes(),
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
        assert_eq!(wb.shared_strings.len(), 1);
        assert_eq!(wb.shared_strings.get(0), Some("Hello"));

        // Worksheet has one row with two cells.
        let ws = &wb.sheets[0].worksheet;
        assert_eq!(ws.rows.len(), 1);
        assert_eq!(ws.rows[0].cells.len(), 2);

        // A1 is a shared string reference (value = "0").
        let a1 = &ws.rows[0].cells[0];
        assert_eq!(a1.reference, "A1");
        assert_eq!(a1.cell_type, CellType::SharedString);
        assert_eq!(a1.value.as_deref(), Some("0"));

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

        // Shared strings table should be empty (not an error).
        assert!(wb.shared_strings.is_empty());

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
            data: b"<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"/>".to_vec(),
        }];
        let zip_bytes = write_zip(&entries).unwrap();

        let result = read_xlsx(&zip_bytes);
        assert!(result.is_err());
        let err = result.unwrap_err();
        assert!(
            matches!(err, IronsheetError::MissingPart(ref msg) if msg.contains("workbook.xml")),
            "expected MissingPart error, got: {err}"
        );
    }
}
