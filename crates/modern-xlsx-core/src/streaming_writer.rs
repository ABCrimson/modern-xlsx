//! True streaming XLSX writer that writes rows directly to ZIP entries.
//!
//! Unlike the existing [`crate::streaming::StreamingWriter`] which accumulates
//! all rows in memory before delegating to the standard writer, this module
//! writes worksheet XML incrementally into a [`ZipWriter`] — so peak memory is
//! proportional to the number of *unique strings* (the SST), not the total row
//! count.
//!
//! # Usage
//!
//! ```rust,no_run
//! use modern_xlsx_core::streaming_writer::{StreamingWriterCore, StreamingCell};
//! use modern_xlsx_core::ooxml::worksheet::CellType;
//!
//! let mut w = StreamingWriterCore::new();
//! w.start_sheet("Sheet1").unwrap();
//! w.write_row(&[
//!     StreamingCell { value: Some("Hello".into()), cell_type: Some(CellType::SharedString), style: None },
//!     StreamingCell { value: Some("42".into()), cell_type: Some(CellType::Number), style: None },
//! ]).unwrap();
//! let xlsx_bytes = w.finish().unwrap();
//! ```

use core::hint::cold_path;
use std::collections::HashMap;
use std::io::{Cursor, Write};

use zip::write::SimpleFileOptions;
use zip::{CompressionMethod, ZipWriter};

use crate::ooxml::worksheet::CellType;
use crate::{ModernXlsxError, Result};

// ---------------------------------------------------------------------------
// Data types
// ---------------------------------------------------------------------------

/// A cell value for the streaming writer.
#[derive(Debug, Clone, serde::Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct StreamingCell {
    /// Cell value as a string.
    #[serde(default)]
    pub value: Option<String>,
    /// Cell type (defaults to [`CellType::Number`]).
    #[serde(default)]
    pub cell_type: Option<CellType>,
    /// Style index (0-based, indexes into the styles `cellXfs` array).
    #[serde(default)]
    pub style: Option<u32>,
}

// ---------------------------------------------------------------------------
// Streaming XLSX writer
// ---------------------------------------------------------------------------

/// Streaming XLSX writer that writes rows incrementally to ZIP entries.
///
/// Worksheet XML is written directly into the ZIP stream as rows are added,
/// so the writer never holds the full worksheet in memory. The shared string
/// table (SST) is accumulated and written at the end in [`finish()`].
pub struct StreamingWriterCore {
    zip: ZipWriter<Cursor<Vec<u8>>>,
    /// Shared string table: string -> index.
    sst: HashMap<String, u32>,
    /// Insertion-ordered SST values.
    sst_order: Vec<String>,
    /// Number of sheets created so far.
    sheet_count: u32,
    /// Sheet names (in order).
    sheet_names: Vec<String>,
    /// Whether a sheet is currently open for writing.
    current_sheet_open: bool,
    /// Current 1-based row number within the open sheet.
    current_row: u32,
    /// Optional custom styles XML (entire `xl/styles.xml` content).
    styles_xml: Option<String>,
}

impl StreamingWriterCore {
    /// Create a new streaming writer with a 64 KiB initial buffer.
    pub fn new() -> Self {
        let buf = Vec::with_capacity(64 * 1024);
        let zip = ZipWriter::new(Cursor::new(buf));
        Self {
            zip,
            sst: HashMap::new(),
            sst_order: Vec::new(),
            sheet_count: 0,
            sheet_names: Vec::new(),
            current_sheet_open: false,
            current_row: 0,
            styles_xml: None,
        }
    }

    /// Set custom styles XML (the complete `xl/styles.xml` body).
    ///
    /// When set, this XML is written verbatim instead of the minimal default.
    pub fn set_styles_xml(&mut self, xml: String) {
        self.styles_xml = Some(xml);
    }

    /// Start a new worksheet with the given name.
    ///
    /// If another sheet is currently open, it is closed first.
    pub fn start_sheet(&mut self, name: &str) -> Result<()> {
        if self.current_sheet_open {
            self.end_sheet()?;
        }
        self.sheet_count += 1;
        self.sheet_names.push(name.to_string());
        self.current_row = 0;

        let path = format!("xl/worksheets/sheet{}.xml", self.sheet_count);
        let opts = zip_options();
        self.zip
            .start_file(path, opts)
            .map_err(|e| ModernXlsxError::ZipWrite(e.to_string()))?;

        // Worksheet XML preamble.
        write!(
            self.zip,
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\
             <worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" \
             xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">\
             <sheetData>"
        )
        .map_err(|e| ModernXlsxError::ZipWrite(e.to_string()))?;

        self.current_sheet_open = true;
        Ok(())
    }

    /// Write a row of cells to the currently open sheet.
    ///
    /// Cells are written left-to-right starting at column A. Empty cells
    /// (where `value` is `None`) are skipped.
    ///
    /// # Errors
    ///
    /// Returns an error if no sheet is currently open.
    pub fn write_row(&mut self, cells: &[StreamingCell]) -> Result<()> {
        if !self.current_sheet_open {
            cold_path();
            return Err(ModernXlsxError::XmlWrite(
                "No sheet is open — call start_sheet() first".into(),
            ));
        }
        self.current_row += 1;
        let row_num = self.current_row;

        let mut ibuf = itoa::Buffer::new();
        write!(self.zip, "<row r=\"{}\">", ibuf.format(row_num))
            .map_err(|e| ModernXlsxError::ZipWrite(e.to_string()))?;

        for (col_idx, cell) in cells.iter().enumerate() {
            let Some(ref val) = cell.value else {
                continue;
            };

            let col_letter = col_index_to_letter(col_idx as u32);
            // cell_ref = e.g. "A1", "B2", ...
            let cell_ref_row = ibuf.format(row_num);
            // We need to build "A1" etc. using the col letter + row number.
            // Since ibuf is reused, copy the row string first.
            let row_str: &str = cell_ref_row;

            let ct = cell.cell_type.unwrap_or(CellType::Number);
            let style_attr = match cell.style {
                Some(s) => {
                    let mut buf = String::with_capacity(8);
                    buf.push_str(" s=\"");
                    buf.push_str(itoa::Buffer::new().format(s));
                    buf.push('"');
                    buf
                }
                None => String::new(),
            };

            match ct {
                CellType::SharedString => {
                    let idx = self.intern_string(val);
                    write!(
                        self.zip,
                        "<c r=\"{col_letter}{row_str}\" t=\"s\"{style_attr}><v>{}</v></c>",
                        itoa::Buffer::new().format(idx)
                    )
                    .map_err(|e| ModernXlsxError::ZipWrite(e.to_string()))?;
                }
                CellType::Number => {
                    write!(
                        self.zip,
                        "<c r=\"{col_letter}{row_str}\"{style_attr}><v>{val}</v></c>"
                    )
                    .map_err(|e| ModernXlsxError::ZipWrite(e.to_string()))?;
                }
                CellType::Boolean => {
                    let bv = if val == "true" || val == "1" {
                        "1"
                    } else {
                        "0"
                    };
                    write!(
                        self.zip,
                        "<c r=\"{col_letter}{row_str}\" t=\"b\"{style_attr}><v>{bv}</v></c>"
                    )
                    .map_err(|e| ModernXlsxError::ZipWrite(e.to_string()))?;
                }
                // All remaining types: inline string, formula result, error, stub.
                CellType::InlineStr
                | CellType::FormulaStr
                | CellType::Error
                | CellType::Stub => {
                    write!(
                        self.zip,
                        "<c r=\"{col_letter}{row_str}\" t=\"inlineStr\"{style_attr}>\
                         <is><t>{}</t></is></c>",
                        xml_escape(val)
                    )
                    .map_err(|e| ModernXlsxError::ZipWrite(e.to_string()))?;
                }
            }
        }

        write!(self.zip, "</row>")
            .map_err(|e| ModernXlsxError::ZipWrite(e.to_string()))?;
        Ok(())
    }

    /// Close the currently open sheet (writes `</sheetData></worksheet>`).
    fn end_sheet(&mut self) -> Result<()> {
        write!(self.zip, "</sheetData></worksheet>")
            .map_err(|e| ModernXlsxError::ZipWrite(e.to_string()))?;
        self.current_sheet_open = false;
        Ok(())
    }

    /// Intern a string into the SST and return its zero-based index.
    #[inline]
    fn intern_string(&mut self, s: &str) -> u32 {
        if let Some(&idx) = self.sst.get(s) {
            return idx;
        }
        let idx = self.sst_order.len() as u32;
        self.sst.insert(s.to_string(), idx);
        self.sst_order.push(s.to_string());
        idx
    }

    /// Finish writing: close any open sheet, write metadata parts (content
    /// types, relationships, workbook, SST, styles), and return the complete
    /// XLSX bytes.
    pub fn finish(mut self) -> Result<Vec<u8>> {
        if self.current_sheet_open {
            self.end_sheet()?;
        }

        if self.sheet_count == 0 {
            cold_path();
            return Err(ModernXlsxError::InvalidCellValue(
                "workbook must contain at least one sheet".into(),
            ));
        }

        let opts = zip_options();
        let sc = self.sheet_count as usize;

        // --- [Content_Types].xml ---
        self.zip
            .start_file("[Content_Types].xml", opts)
            .map_err(|e| ModernXlsxError::ZipWrite(e.to_string()))?;
        write!(
            self.zip,
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\
             <Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">\
             <Default Extension=\"rels\" \
             ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>\
             <Default Extension=\"xml\" ContentType=\"application/xml\"/>\
             <Override PartName=\"/xl/workbook.xml\" \
             ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>\
             <Override PartName=\"/xl/sharedStrings.xml\" \
             ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml\"/>\
             <Override PartName=\"/xl/styles.xml\" \
             ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\"/>"
        )
        .map_err(|e| ModernXlsxError::ZipWrite(e.to_string()))?;
        for i in 1..=sc {
            write!(
                self.zip,
                "<Override PartName=\"/xl/worksheets/sheet{i}.xml\" \
                 ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>"
            )
            .map_err(|e| ModernXlsxError::ZipWrite(e.to_string()))?;
        }
        write!(self.zip, "</Types>")
            .map_err(|e| ModernXlsxError::ZipWrite(e.to_string()))?;

        // --- _rels/.rels ---
        self.zip
            .start_file("_rels/.rels", opts)
            .map_err(|e| ModernXlsxError::ZipWrite(e.to_string()))?;
        write!(
            self.zip,
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\
             <Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\
             <Relationship Id=\"rId1\" \
             Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" \
             Target=\"xl/workbook.xml\"/>\
             </Relationships>"
        )
        .map_err(|e| ModernXlsxError::ZipWrite(e.to_string()))?;

        // --- xl/workbook.xml ---
        self.zip
            .start_file("xl/workbook.xml", opts)
            .map_err(|e| ModernXlsxError::ZipWrite(e.to_string()))?;
        write!(
            self.zip,
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\
             <workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" \
             xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">\
             <sheets>"
        )
        .map_err(|e| ModernXlsxError::ZipWrite(e.to_string()))?;
        for (i, name) in self.sheet_names.iter().enumerate() {
            let sheet_id = i + 1;
            write!(
                self.zip,
                "<sheet name=\"{}\" sheetId=\"{sheet_id}\" r:id=\"rId{sheet_id}\"/>",
                xml_escape(name)
            )
            .map_err(|e| ModernXlsxError::ZipWrite(e.to_string()))?;
        }
        write!(self.zip, "</sheets></workbook>")
            .map_err(|e| ModernXlsxError::ZipWrite(e.to_string()))?;

        // --- xl/_rels/workbook.xml.rels ---
        self.zip
            .start_file("xl/_rels/workbook.xml.rels", opts)
            .map_err(|e| ModernXlsxError::ZipWrite(e.to_string()))?;
        write!(
            self.zip,
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\
             <Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
        )
        .map_err(|e| ModernXlsxError::ZipWrite(e.to_string()))?;
        for i in 1..=sc {
            write!(
                self.zip,
                "<Relationship Id=\"rId{i}\" \
                 Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" \
                 Target=\"worksheets/sheet{i}.xml\"/>"
            )
            .map_err(|e| ModernXlsxError::ZipWrite(e.to_string()))?;
        }
        let ss_id = sc + 1;
        let st_id = sc + 2;
        write!(
            self.zip,
            "<Relationship Id=\"rId{ss_id}\" \
             Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings\" \
             Target=\"sharedStrings.xml\"/>\
             <Relationship Id=\"rId{st_id}\" \
             Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" \
             Target=\"styles.xml\"/>\
             </Relationships>"
        )
        .map_err(|e| ModernXlsxError::ZipWrite(e.to_string()))?;

        // --- xl/sharedStrings.xml ---
        self.zip
            .start_file("xl/sharedStrings.xml", opts)
            .map_err(|e| ModernXlsxError::ZipWrite(e.to_string()))?;
        let sst_count = self.sst_order.len();
        write!(
            self.zip,
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\
             <sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" \
             count=\"{sst_count}\" uniqueCount=\"{sst_count}\">"
        )
        .map_err(|e| ModernXlsxError::ZipWrite(e.to_string()))?;
        for s in &self.sst_order {
            let needs_preserve = s.starts_with(' ')
                || s.ends_with(' ')
                || s.starts_with('\t')
                || s.ends_with('\t')
                || s.starts_with('\n')
                || s.ends_with('\n');
            if needs_preserve {
                write!(
                    self.zip,
                    "<si><t xml:space=\"preserve\">{}</t></si>",
                    xml_escape(s)
                )
                .map_err(|e| ModernXlsxError::ZipWrite(e.to_string()))?;
            } else {
                write!(self.zip, "<si><t>{}</t></si>", xml_escape(s))
                    .map_err(|e| ModernXlsxError::ZipWrite(e.to_string()))?;
            }
        }
        write!(self.zip, "</sst>")
            .map_err(|e| ModernXlsxError::ZipWrite(e.to_string()))?;

        // --- xl/styles.xml ---
        self.zip
            .start_file("xl/styles.xml", opts)
            .map_err(|e| ModernXlsxError::ZipWrite(e.to_string()))?;
        if let Some(ref custom) = self.styles_xml {
            write!(self.zip, "{custom}")
                .map_err(|e| ModernXlsxError::ZipWrite(e.to_string()))?;
        } else {
            write!(self.zip, "{MINIMAL_STYLES_XML}")
                .map_err(|e| ModernXlsxError::ZipWrite(e.to_string()))?;
        }

        // Finalize the ZIP archive.
        let cursor = self
            .zip
            .finish()
            .map_err(|e| ModernXlsxError::ZipFinalize(e.to_string()))?;
        Ok(cursor.into_inner())
    }
}

impl Default for StreamingWriterCore {
    fn default() -> Self {
        Self::new()
    }
}

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

/// ZIP file options used for all entries.
#[inline]
fn zip_options() -> SimpleFileOptions {
    SimpleFileOptions::default()
        .compression_method(CompressionMethod::Deflated)
        .compression_level(Some(6))
}

/// Convert a 0-based column index to an Excel-style letter (A, B, ..., Z, AA, AB, ...).
fn col_index_to_letter(mut idx: u32) -> String {
    let mut result = Vec::with_capacity(3);
    loop {
        result.push(b'A' + (idx % 26) as u8);
        if idx < 26 {
            break;
        }
        idx = idx / 26 - 1;
    }
    result.reverse();
    // SAFETY: all bytes are valid ASCII A-Z.
    unsafe { String::from_utf8_unchecked(result) }
}

/// XML-escape a string for use in XML text content or attribute values.
#[inline]
fn xml_escape(s: &str) -> String {
    let mut out = String::with_capacity(s.len());
    for ch in s.chars() {
        match ch {
            '&' => out.push_str("&amp;"),
            '<' => out.push_str("&lt;"),
            '>' => out.push_str("&gt;"),
            '"' => out.push_str("&quot;"),
            '\'' => out.push_str("&apos;"),
            _ => out.push(ch),
        }
    }
    out
}

/// Minimal `xl/styles.xml` that defines one default font, two fills (none + gray125),
/// one border, and one cell format — the absolute minimum Excel requires.
const MINIMAL_STYLES_XML: &str = "\
<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\
<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">\
<fonts count=\"1\"><font><sz val=\"11\"/><name val=\"Aptos\"/></font></fonts>\
<fills count=\"2\">\
<fill><patternFill patternType=\"none\"/></fill>\
<fill><patternFill patternType=\"gray125\"/></fill>\
</fills>\
<borders count=\"1\"><border><left/><right/><top/><bottom/><diagonal/></border></borders>\
<cellStyleXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/></cellStyleXfs>\
<cellXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/></cellXfs>\
</styleSheet>";

// ---------------------------------------------------------------------------
// Tests
// ---------------------------------------------------------------------------

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn col_letter_single() {
        assert_eq!(col_index_to_letter(0), "A");
        assert_eq!(col_index_to_letter(1), "B");
        assert_eq!(col_index_to_letter(25), "Z");
    }

    #[test]
    fn col_letter_double() {
        assert_eq!(col_index_to_letter(26), "AA");
        assert_eq!(col_index_to_letter(27), "AB");
        assert_eq!(col_index_to_letter(51), "AZ");
        assert_eq!(col_index_to_letter(52), "BA");
        assert_eq!(col_index_to_letter(701), "ZZ");
    }

    #[test]
    fn col_letter_triple() {
        assert_eq!(col_index_to_letter(702), "AAA");
    }

    #[test]
    fn xml_escape_basic() {
        assert_eq!(xml_escape("hello"), "hello");
        assert_eq!(xml_escape("a&b"), "a&amp;b");
        assert_eq!(xml_escape("<tag>"), "&lt;tag&gt;");
        assert_eq!(xml_escape("\"quoted\""), "&quot;quoted&quot;");
        assert_eq!(xml_escape("it's"), "it&apos;s");
    }

    #[test]
    fn streaming_write_single_sheet_numbers() {
        let mut w = StreamingWriterCore::new();
        w.start_sheet("Data").unwrap();
        for i in 1..=100 {
            w.write_row(&[
                StreamingCell {
                    value: Some(i.to_string()),
                    cell_type: Some(CellType::Number),
                    style: None,
                },
                StreamingCell {
                    value: Some((i * 10).to_string()),
                    cell_type: Some(CellType::Number),
                    style: None,
                },
            ])
            .unwrap();
        }
        let bytes = w.finish().unwrap();

        // Verify by reading back with the standard reader.
        let wb = crate::reader::read_xlsx(&bytes).unwrap();
        assert_eq!(wb.sheets.len(), 1);
        assert_eq!(wb.sheets[0].name, "Data");
        assert_eq!(wb.sheets[0].worksheet.rows.len(), 100);
        assert_eq!(
            wb.sheets[0].worksheet.rows[0].cells[0]
                .value
                .as_deref(),
            Some("1")
        );
        assert_eq!(
            wb.sheets[0].worksheet.rows[0].cells[1]
                .value
                .as_deref(),
            Some("10")
        );
        assert_eq!(
            wb.sheets[0].worksheet.rows[99].cells[0]
                .value
                .as_deref(),
            Some("100")
        );
    }

    #[test]
    fn streaming_write_shared_strings() {
        let mut w = StreamingWriterCore::new();
        w.start_sheet("SST").unwrap();
        w.write_row(&[
            StreamingCell {
                value: Some("Hello".into()),
                cell_type: Some(CellType::SharedString),
                style: None,
            },
            StreamingCell {
                value: Some("World".into()),
                cell_type: Some(CellType::SharedString),
                style: None,
            },
        ])
        .unwrap();
        // Write a duplicate to verify SST deduplication.
        w.write_row(&[StreamingCell {
            value: Some("Hello".into()),
            cell_type: Some(CellType::SharedString),
            style: None,
        }])
        .unwrap();
        let bytes = w.finish().unwrap();

        let wb = crate::reader::read_xlsx(&bytes).unwrap();
        let sst = wb.shared_strings.as_ref().unwrap();
        // Only 2 unique strings.
        assert_eq!(sst.strings.len(), 2);
        assert_eq!(sst.strings[0], "Hello");
        assert_eq!(sst.strings[1], "World");

        // Cell values should be resolved.
        assert_eq!(
            wb.sheets[0].worksheet.rows[0].cells[0]
                .value
                .as_deref(),
            Some("Hello")
        );
        assert_eq!(
            wb.sheets[0].worksheet.rows[0].cells[1]
                .value
                .as_deref(),
            Some("World")
        );
        assert_eq!(
            wb.sheets[0].worksheet.rows[1].cells[0]
                .value
                .as_deref(),
            Some("Hello")
        );
    }

    #[test]
    fn streaming_write_multiple_sheets() {
        let mut w = StreamingWriterCore::new();
        w.start_sheet("First").unwrap();
        w.write_row(&[StreamingCell {
            value: Some("1".into()),
            cell_type: Some(CellType::Number),
            style: None,
        }])
        .unwrap();

        w.start_sheet("Second").unwrap();
        w.write_row(&[StreamingCell {
            value: Some("2".into()),
            cell_type: Some(CellType::Number),
            style: None,
        }])
        .unwrap();

        let bytes = w.finish().unwrap();

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
    fn streaming_write_booleans() {
        let mut w = StreamingWriterCore::new();
        w.start_sheet("Bools").unwrap();
        w.write_row(&[
            StreamingCell {
                value: Some("true".into()),
                cell_type: Some(CellType::Boolean),
                style: None,
            },
            StreamingCell {
                value: Some("false".into()),
                cell_type: Some(CellType::Boolean),
                style: None,
            },
            StreamingCell {
                value: Some("1".into()),
                cell_type: Some(CellType::Boolean),
                style: None,
            },
        ])
        .unwrap();
        let bytes = w.finish().unwrap();

        let wb = crate::reader::read_xlsx(&bytes).unwrap();
        let cells = &wb.sheets[0].worksheet.rows[0].cells;
        assert_eq!(cells[0].value.as_deref(), Some("1"));
        assert_eq!(cells[1].value.as_deref(), Some("0"));
        assert_eq!(cells[2].value.as_deref(), Some("1"));
    }

    #[test]
    fn streaming_write_inline_strings() {
        let mut w = StreamingWriterCore::new();
        w.start_sheet("Inline").unwrap();
        w.write_row(&[
            StreamingCell {
                value: Some("plain text value".into()),
                cell_type: Some(CellType::InlineStr),
                style: None,
            },
            StreamingCell {
                value: Some("with special: A&B".into()),
                cell_type: Some(CellType::InlineStr),
                style: None,
            },
        ])
        .unwrap();
        let bytes = w.finish().unwrap();

        let wb = crate::reader::read_xlsx(&bytes).unwrap();
        assert_eq!(
            wb.sheets[0].worksheet.rows[0].cells[0]
                .value
                .as_deref(),
            Some("plain text value")
        );
        assert_eq!(
            wb.sheets[0].worksheet.rows[0].cells[1]
                .value
                .as_deref(),
            Some("with special: A&B")
        );
    }

    #[test]
    fn streaming_write_empty_cells_skipped() {
        let mut w = StreamingWriterCore::new();
        w.start_sheet("Gaps").unwrap();
        w.write_row(&[
            StreamingCell {
                value: Some("A".into()),
                cell_type: Some(CellType::SharedString),
                style: None,
            },
            StreamingCell {
                value: None,
                cell_type: None,
                style: None,
            },
            StreamingCell {
                value: Some("C".into()),
                cell_type: Some(CellType::SharedString),
                style: None,
            },
        ])
        .unwrap();
        let bytes = w.finish().unwrap();

        let wb = crate::reader::read_xlsx(&bytes).unwrap();
        let cells = &wb.sheets[0].worksheet.rows[0].cells;
        // Only 2 cells (A1 and C1) should be present; B1 is skipped.
        assert_eq!(cells.len(), 2);
        assert_eq!(cells[0].reference, "A1");
        assert_eq!(cells[0].value.as_deref(), Some("A"));
        assert_eq!(cells[1].reference, "C1");
        assert_eq!(cells[1].value.as_deref(), Some("C"));
    }

    #[test]
    fn streaming_write_no_sheet_errors() {
        let w = StreamingWriterCore::new();
        // Finishing without any sheets should error.
        let result = w.finish();
        assert!(result.is_err());
    }

    #[test]
    fn streaming_write_row_without_sheet_errors() {
        let mut w = StreamingWriterCore::new();
        let result = w.write_row(&[StreamingCell {
            value: Some("x".into()),
            cell_type: None,
            style: None,
        }]);
        assert!(result.is_err());
    }

    #[test]
    fn streaming_write_with_style_index() {
        let mut w = StreamingWriterCore::new();
        w.start_sheet("Styled").unwrap();
        w.write_row(&[StreamingCell {
            value: Some("42".into()),
            cell_type: Some(CellType::Number),
            style: Some(0),
        }])
        .unwrap();
        let bytes = w.finish().unwrap();

        let wb = crate::reader::read_xlsx(&bytes).unwrap();
        assert_eq!(
            wb.sheets[0].worksheet.rows[0].cells[0]
                .value
                .as_deref(),
            Some("42")
        );
        assert_eq!(wb.sheets[0].worksheet.rows[0].cells[0].style_index, Some(0));
    }

    #[test]
    fn streaming_write_1k_rows() {
        let mut w = StreamingWriterCore::new();
        w.start_sheet("BigData").unwrap();
        for i in 1..=1000 {
            w.write_row(&[
                StreamingCell {
                    value: Some(i.to_string()),
                    cell_type: Some(CellType::Number),
                    style: None,
                },
                StreamingCell {
                    value: Some(format!("row_{i}")),
                    cell_type: Some(CellType::SharedString),
                    style: None,
                },
            ])
            .unwrap();
        }
        let bytes = w.finish().unwrap();

        let wb = crate::reader::read_xlsx(&bytes).unwrap();
        assert_eq!(wb.sheets[0].worksheet.rows.len(), 1000);
        assert_eq!(
            wb.sheets[0].worksheet.rows[999].cells[0]
                .value
                .as_deref(),
            Some("1000")
        );
        assert_eq!(
            wb.sheets[0].worksheet.rows[999].cells[1]
                .value
                .as_deref(),
            Some("row_1000")
        );
    }

    #[test]
    fn streaming_write_xml_space_preserve() {
        let mut w = StreamingWriterCore::new();
        w.start_sheet("Spaces").unwrap();
        w.write_row(&[StreamingCell {
            value: Some(" leading".into()),
            cell_type: Some(CellType::SharedString),
            style: None,
        }])
        .unwrap();
        let bytes = w.finish().unwrap();

        let wb = crate::reader::read_xlsx(&bytes).unwrap();
        assert_eq!(
            wb.sheets[0].worksheet.rows[0].cells[0]
                .value
                .as_deref(),
            Some(" leading")
        );
    }

    #[test]
    fn streaming_write_auto_close_previous_sheet() {
        // Calling start_sheet() when a sheet is already open should auto-close it.
        let mut w = StreamingWriterCore::new();
        w.start_sheet("Sheet1").unwrap();
        w.write_row(&[StreamingCell {
            value: Some("a".into()),
            cell_type: Some(CellType::SharedString),
            style: None,
        }])
        .unwrap();
        // Directly start another sheet without explicit close.
        w.start_sheet("Sheet2").unwrap();
        w.write_row(&[StreamingCell {
            value: Some("b".into()),
            cell_type: Some(CellType::SharedString),
            style: None,
        }])
        .unwrap();
        let bytes = w.finish().unwrap();

        let wb = crate::reader::read_xlsx(&bytes).unwrap();
        assert_eq!(wb.sheets.len(), 2);
        assert_eq!(
            wb.sheets[0].worksheet.rows[0].cells[0]
                .value
                .as_deref(),
            Some("a")
        );
        assert_eq!(
            wb.sheets[1].worksheet.rows[0].cells[0]
                .value
                .as_deref(),
            Some("b")
        );
    }
}
