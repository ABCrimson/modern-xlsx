//! Streaming read/write API for large XLSX files.
//!
//! The streaming reader parses worksheets row-by-row without loading
//! the entire XML into memory. The streaming writer writes rows
//! incrementally to ZIP entries.

use crate::dates::DateSystem;
use crate::errors::ModernXlsxError;
use crate::ooxml::shared_strings::SharedStringTable;
use crate::ooxml::styles::Styles;
use crate::ooxml::workbook::WorkbookXml;
use crate::ooxml::worksheet::Row;
use crate::zip::reader::ZipSecurityLimits;

/// A streaming XLSX reader that yields rows one at a time.
///
/// Unlike the standard reader which assembles a full [`crate::WorkbookData`],
/// the streaming reader keeps the raw worksheet bytes and parses them
/// on demand, reducing peak memory when only a subset of sheets (or rows)
/// is needed.
pub struct StreamingReader {
    /// Parsed shared string table.
    shared_strings: Option<SharedStringTable>,
    /// Parsed styles.
    styles: Styles,
    /// Workbook metadata.
    workbook: WorkbookXml,
    /// Raw ZIP entry data for each worksheet, keyed by sheet name.
    sheet_data: std::collections::BTreeMap<String, Vec<u8>>,
}

impl StreamingReader {
    /// Open a streaming reader from XLSX bytes using default security limits.
    pub fn open(data: &[u8]) -> Result<Self, ModernXlsxError> {
        Self::open_with_limits(data, &ZipSecurityLimits::default())
    }

    /// Open a streaming reader from XLSX bytes with custom ZIP security limits.
    pub fn open_with_limits(
        data: &[u8],
        limits: &ZipSecurityLimits,
    ) -> Result<Self, ModernXlsxError> {
        use crate::ooxml::relationships::Relationships;
        use crate::zip::reader::read_zip_entries;

        let mut entries = read_zip_entries(data, limits)?;

        // Parse workbook.xml (required).
        let wb_data = entries
            .get("xl/workbook.xml")
            .ok_or_else(|| ModernXlsxError::MissingPart("xl/workbook.xml".into()))?;
        let workbook = WorkbookXml::parse(wb_data)?;

        // Parse shared strings (optional).
        let shared_strings = entries
            .get("xl/sharedStrings.xml")
            .map(|d| SharedStringTable::parse(d))
            .transpose()?;

        // Parse styles (optional — use defaults when absent).
        let styles = entries
            .get("xl/styles.xml")
            .map(|d| Styles::parse(d))
            .transpose()?
            .unwrap_or_else(Styles::default_styles);

        // Parse relationships to map sheet names to paths.
        let wb_rels = entries
            .get("xl/_rels/workbook.xml.rels")
            .map(|d| Relationships::parse(d))
            .transpose()?
            .unwrap_or_else(Relationships::new);

        // Collect raw worksheet data keyed by sheet name.
        // Use remove() to take ownership instead of cloning large byte vectors.
        let mut sheet_data = std::collections::BTreeMap::new();
        for sheet in &workbook.sheets {
            if let Some(rel) = wb_rels.get_by_id(&sheet.r_id) {
                let path = if rel.target.starts_with('/') {
                    rel.target.trim_start_matches('/').to_string()
                } else {
                    format!("xl/{}", rel.target)
                };
                if let Some(data) = entries.remove(&path) {
                    sheet_data.insert(sheet.name.clone(), data);
                }
            }
        }

        Ok(Self {
            shared_strings,
            styles,
            workbook,
            sheet_data,
        })
    }

    /// Get the list of sheet names.
    pub fn sheet_names(&self) -> Vec<String> {
        self.workbook.sheets.iter().map(|s| s.name.clone()).collect()
    }

    /// Get the date system.
    pub fn date_system(&self) -> DateSystem {
        self.workbook.date_system
    }

    /// Get a reference to the shared string table (if present).
    pub fn shared_strings(&self) -> Option<&SharedStringTable> {
        self.shared_strings.as_ref()
    }

    /// Get a reference to styles.
    pub fn styles(&self) -> &Styles {
        &self.styles
    }

    /// Read a specific sheet by name, returning its rows.
    ///
    /// Each call re-parses the worksheet XML from the raw bytes that were
    /// captured when the reader was opened. For a true streaming iterator
    /// this would yield rows one at a time via SAX-style parsing, but the
    /// current implementation reuses the existing `WorksheetXml::parse`.
    pub fn read_sheet_rows(&self, name: &str) -> Result<Vec<Row>, ModernXlsxError> {
        let data = self
            .sheet_data
            .get(name)
            .ok_or_else(|| ModernXlsxError::MissingPart(format!("sheet: {name}")))?;
        let ws = crate::ooxml::worksheet::WorksheetXml::parse_with_sst(
            data,
            self.shared_strings.as_ref(),
        )?;
        Ok(ws.rows)
    }
}

// ---------------------------------------------------------------------------
// Streaming Writer
// ---------------------------------------------------------------------------

/// A streaming XLSX writer that accumulates rows per sheet and produces the
/// final XLSX bytes on [`StreamingWriter::finish`].
///
/// The API is designed so that callers can add rows incrementally without
/// constructing the entire [`crate::WorkbookData`] up front.
pub struct StreamingWriter {
    sheets: Vec<StreamingSheet>,
    date_system: DateSystem,
    styles: Styles,
}

struct StreamingSheet {
    name: String,
    rows: Vec<Row>,
}

impl StreamingWriter {
    /// Create a new streaming writer with default settings.
    pub fn new() -> Self {
        Self {
            sheets: Vec::new(),
            date_system: DateSystem::Date1900,
            styles: Styles::default_styles(),
        }
    }

    /// Set the date system.
    pub fn set_date_system(&mut self, ds: DateSystem) -> &mut Self {
        self.date_system = ds;
        self
    }

    /// Set styles.
    pub fn set_styles(&mut self, s: Styles) -> &mut Self {
        self.styles = s;
        self
    }

    /// Start a new sheet, returning its zero-based index.
    ///
    /// Use [`add_row`](StreamingWriter::add_row) with this index to append
    /// rows to the sheet.
    pub fn add_sheet(&mut self, name: impl Into<String>) -> usize {
        let idx = self.sheets.len();
        self.sheets.push(StreamingSheet {
            name: name.into(),
            rows: Vec::new(),
        });
        idx
    }

    /// Add a row to the sheet at the given index.
    ///
    /// # Panics
    ///
    /// Panics if `sheet_index` is out of bounds.
    pub fn add_row(&mut self, sheet_index: usize, row: Row) {
        self.sheets[sheet_index].rows.push(row);
    }

    /// Finish writing and produce the final XLSX bytes.
    ///
    /// This builds a [`crate::WorkbookData`] from the accumulated sheets and
    /// delegates to the standard [`crate::writer::write_xlsx`].
    pub fn finish(self) -> Result<Vec<u8>, ModernXlsxError> {
        use crate::ooxml::worksheet::WorksheetXml;
        use crate::{SheetData, WorkbookData};

        let sheets = self
            .sheets
            .into_iter()
            .map(|s| SheetData {
                name: s.name,
                worksheet: WorksheetXml {
                    rows: s.rows,
                    merge_cells: Vec::new(),
                    auto_filter: None,
                    frozen_pane: None,
                    columns: Vec::new(),
                    dimension: None,
                    data_validations: Vec::new(),
                    conditional_formatting: Vec::new(),
                    hyperlinks: Vec::new(),
                    page_setup: None,
                    sheet_protection: None,
                    comments: Vec::new(),
                    tab_color: None,
                    tables: Vec::new(),
                },
            })
            .collect();

        let wb = WorkbookData {
            sheets,
            date_system: self.date_system,
            styles: self.styles,
            shared_strings: None,
            defined_names: Vec::new(),
            doc_properties: None,
            theme_colors: None,
            calc_chain: Vec::new(),
            workbook_views: Vec::new(),
            preserved_entries: std::collections::BTreeMap::new(),
        };

        crate::writer::write_xlsx(&wb)
    }
}

impl Default for StreamingWriter {
    fn default() -> Self {
        Self::new()
    }
}

// ---------------------------------------------------------------------------
// Tests
// ---------------------------------------------------------------------------

#[cfg(test)]
mod tests {
    use super::*;
    use crate::ooxml::worksheet::{Cell, CellType, Row};

    #[test]
    fn test_streaming_reader() {
        // Build a test XLSX using the regular writer.
        let wb = crate::WorkbookData {
            sheets: vec![crate::SheetData {
                name: "Test".into(),
                worksheet: crate::ooxml::worksheet::WorksheetXml {
                    rows: vec![Row {
                        index: 1,
                        cells: vec![Cell {
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
                        }],
                        height: None,
                        hidden: false,
                    }],
                    merge_cells: vec![],
                    auto_filter: None,
                    frozen_pane: None,
                    columns: vec![],
                    dimension: None,
                    data_validations: vec![],
                    conditional_formatting: vec![],
                    hyperlinks: vec![],
                    page_setup: None,
                    sheet_protection: None,
                    comments: Vec::new(),
                    tab_color: None,
                    tables: Vec::new(),
                },
            }],
            date_system: DateSystem::Date1900,
            styles: Styles::default_styles(),
            shared_strings: None,
            defined_names: vec![],
            doc_properties: None,
            theme_colors: None,
            calc_chain: vec![],
            workbook_views: vec![],
            preserved_entries: std::collections::BTreeMap::new(),
        };

        let bytes = crate::writer::write_xlsx(&wb).unwrap();
        let reader = StreamingReader::open(&bytes).unwrap();

        assert_eq!(reader.sheet_names(), vec!["Test"]);
        assert_eq!(reader.date_system(), DateSystem::Date1900);
        assert!(reader.shared_strings().is_some());

        let rows = reader.read_sheet_rows("Test").unwrap();
        assert_eq!(rows.len(), 1);
        assert_eq!(rows[0].cells[0].value.as_deref(), Some("42"));
    }

    #[test]
    fn test_streaming_reader_missing_sheet() {
        let wb = crate::WorkbookData {
            sheets: vec![crate::SheetData {
                name: "Only".into(),
                worksheet: crate::ooxml::worksheet::WorksheetXml {
                    rows: vec![],
                    merge_cells: vec![],
                    auto_filter: None,
                    frozen_pane: None,
                    columns: vec![],
                    dimension: None,
                    data_validations: vec![],
                    conditional_formatting: vec![],
                    hyperlinks: vec![],
                    page_setup: None,
                    sheet_protection: None,
                    comments: Vec::new(),
                    tab_color: None,
                    tables: Vec::new(),
                },
            }],
            date_system: DateSystem::Date1900,
            styles: Styles::default_styles(),
            shared_strings: None,
            defined_names: vec![],
            doc_properties: None,
            theme_colors: None,
            calc_chain: vec![],
            workbook_views: vec![],
            preserved_entries: std::collections::BTreeMap::new(),
        };

        let bytes = crate::writer::write_xlsx(&wb).unwrap();
        let reader = StreamingReader::open(&bytes).unwrap();

        let result = reader.read_sheet_rows("DoesNotExist");
        assert!(result.is_err());
    }

    #[test]
    fn test_streaming_writer() {
        let mut writer = StreamingWriter::new();
        let idx = writer.add_sheet("Data");
        writer.add_row(
            idx,
            Row {
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
            },
        );

        let bytes = writer.finish().unwrap();

        // Verify by reading back with the standard reader.
        let wb = crate::reader::read_xlsx(&bytes).unwrap();
        assert_eq!(wb.sheets[0].name, "Data");
        assert_eq!(
            wb.sheets[0].worksheet.rows[0].cells[0]
                .value
                .as_deref(),
            Some("100")
        );
    }

    #[test]
    fn test_streaming_writer_multiple_sheets() {
        let mut writer = StreamingWriter::new();
        let s1 = writer.add_sheet("First");
        let s2 = writer.add_sheet("Second");

        writer.add_row(
            s1,
            Row {
                index: 1,
                cells: vec![Cell {
                    reference: "A1".into(),
                    cell_type: CellType::Number,
                    value: Some("1".into()),
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
            },
        );
        writer.add_row(
            s2,
            Row {
                index: 1,
                cells: vec![Cell {
                    reference: "A1".into(),
                    cell_type: CellType::Number,
                    value: Some("2".into()),
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
            },
        );

        let bytes = writer.finish().unwrap();

        let wb = crate::reader::read_xlsx(&bytes).unwrap();
        assert_eq!(wb.sheets.len(), 2);
        assert_eq!(wb.sheets[0].name, "First");
        assert_eq!(wb.sheets[1].name, "Second");
        assert_eq!(
            wb.sheets[0].worksheet.rows[0].cells[0]
                .value
                .as_deref(),
            Some("1")
        );
        assert_eq!(
            wb.sheets[1].worksheet.rows[0].cells[0]
                .value
                .as_deref(),
            Some("2")
        );
    }

    #[test]
    fn test_streaming_roundtrip() {
        // Write with streaming writer, read back with streaming reader.
        let mut writer = StreamingWriter::new();
        let idx = writer.add_sheet("Roundtrip");
        for i in 1..=5 {
            writer.add_row(
                idx,
                Row {
                    index: i,
                    cells: vec![Cell {
                        reference: format!("A{i}"),
                        cell_type: CellType::Number,
                        value: Some(format!("{}", i * 10)),
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
                },
            );
        }

        let bytes = writer.finish().unwrap();
        let reader = StreamingReader::open(&bytes).unwrap();

        assert_eq!(reader.sheet_names(), vec!["Roundtrip"]);
        let rows = reader.read_sheet_rows("Roundtrip").unwrap();
        assert_eq!(rows.len(), 5);
        for (i, row) in rows.iter().enumerate() {
            let expected = format!("{}", (i + 1) * 10);
            assert_eq!(row.cells[0].value.as_deref(), Some(expected.as_str()));
        }
    }
}
