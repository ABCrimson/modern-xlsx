//! Incremental XLSX writer that deflates rows into ZIP entries as they arrive.
//!
//! Each row is rendered into one reused buffer and written straight into a
//! [`ZipWriter`] over an in-memory `Vec<u8>`, so the uncompressed worksheet XML
//! is never held whole. Memory is the compressed archive written so far, plus
//! the shared string table (SST), plus one row buffer; the archive stays in
//! memory until [`StreamingWriterCore::finish`] returns it.
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
use std::borrow::Cow;
use std::collections::HashMap;
use std::io::{Cursor, Write};

use zip::ZipWriter;

use crate::ooxml::cell::{MAX_COL, MAX_ROW, col_to_letters_buf};
use crate::ooxml::content_types::ContentTypes;
use crate::ooxml::shared_strings::SharedStringTableBuilder;
use crate::ooxml::worksheet::CellType;
use crate::ooxml::xml::{REL_NS, SPREADSHEET_NS, needs_preserve};
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
/// so the writer never holds the full uncompressed worksheet XML: each row is
/// rendered into one reused byte buffer and written with a single `write_all`.
/// The compressed archive itself is an in-memory `Vec<u8>` that grows with
/// every row until [`StreamingWriterCore::finish`] returns it.
/// The shared string table (SST) is accumulated in a
/// [`SharedStringTableBuilder`] (each unique string stored once) and streamed
/// into its ZIP entry at the end in [`StreamingWriterCore::finish`].
pub struct StreamingWriterCore {
    zip: ZipWriter<Cursor<Vec<u8>>>,
    /// The shared string table; every `intern_string` call, hits included,
    /// counts as one reference (the SST `count`).
    sst: SharedStringTableBuilder<'static>,
    /// Number of sheets created so far.
    sheet_count: u32,
    /// Sheet names (in order).
    sheet_names: Vec<String>,
    /// The lower case of each name in `sheet_names`, mapped to its position
    /// there, so a duplicate check is one lookup.
    sheet_keys: HashMap<String, usize>,
    /// Whether a sheet is currently open for writing.
    current_sheet_open: bool,
    /// Current 1-based row number within the open sheet.
    current_row: u32,
    /// Optional custom styles XML (entire `xl/styles.xml` content).
    styles_xml: Option<String>,
    /// Reusable render buffer for one row's XML bytes (cleared at the start of
    /// every `write_row` call, then handed to the ZIP stream in one `write_all`).
    row_buf: Vec<u8>,
}

impl StreamingWriterCore {
    /// Create a new streaming writer with a 64 KiB initial buffer.
    pub fn new() -> Self {
        let buf = Vec::with_capacity(64 * 1024);
        let zip = ZipWriter::new(Cursor::new(buf));
        Self {
            zip,
            sst: SharedStringTableBuilder::new(),
            sheet_count: 0,
            sheet_names: Vec::new(),
            sheet_keys: HashMap::new(),
            current_sheet_open: false,
            current_row: 0,
            styles_xml: None,
            row_buf: Vec::with_capacity(256),
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
    ///
    /// # Errors
    ///
    /// Returns [`ModernXlsxError::InvalidCellValue`] when the name is empty,
    /// longer than 31 characters, contains `\ / * ? : [ ]` or a control
    /// character, starts or ends with an apostrophe, is the reserved name
    /// `History` (any case), or duplicates an earlier sheet name
    /// (case-insensitively: the same `str::to_lowercase`, found with one hash
    /// lookup, so starting N sheets takes time linear in N). The check runs
    /// before anything is written, so a currently open sheet stays open after
    /// an error, and a refused name is not recorded.
    pub fn start_sheet(&mut self, name: &str) -> Result<()> {
        if let Some(why) = crate::validate::sheet_name_problem(name) {
            cold_path();
            return Err(ModernXlsxError::InvalidCellValue(format!(
                "Invalid sheet name '{name}': {why}"
            )));
        }
        let lower = name.to_lowercase();
        if let Some(&at) = self.sheet_keys.get(&lower) {
            cold_path();
            let existing = &self.sheet_names[at];
            return Err(ModernXlsxError::InvalidCellValue(format!(
                "Invalid sheet name '{name}': duplicate of '{existing}'"
            )));
        }
        if self.current_sheet_open {
            self.end_sheet()?;
        }
        self.sheet_count += 1;
        self.sheet_keys.insert(lower, self.sheet_names.len());
        self.sheet_names.push(name.to_string());
        self.current_row = 0;

        let path = format!("xl/worksheets/sheet{}.xml", self.sheet_count);
        let opts = crate::zip::writer::deflate_options();
        self.zip.start_file(&path, opts).map_err(|e| {
            ModernXlsxError::ZipWrite(format!("Failed to start worksheet ZIP entry '{path}': {e}"))
        })?;

        // Worksheet XML preamble, with the sheet view Excel writes for a
        // sheet: without any <sheetView>, Excel on a 150% display (144 DPI)
        // showed setRowHeight(1, 40) as 26.7 points; with this view it shows 40.
        write!(
            self.zip,
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\
             <worksheet xmlns=\"{SPREADSHEET_NS}\" \
             xmlns:r=\"{REL_NS}\">\
             <sheetViews><sheetView workbookViewId=\"0\"/></sheetViews>\
             <sheetData>"
        )
        .map_err(zip_err)?;

        self.current_sheet_open = true;
        Ok(())
    }

    /// Write a row of cells to the currently open sheet.
    ///
    /// Cells are written left-to-right starting at column A. Cells whose
    /// `value` is `None` are skipped, except a [`CellType::Stub`] with a
    /// `style`, which is written as an empty styled cell. A `cell_type` of
    /// `None` means [`CellType::Number`]. Per type:
    ///
    /// - `Number`: `value` must parse as a finite `f64`; it is written verbatim.
    /// - `Boolean`: `true`/`false` (any case), `1` or `0`; written as `1` or `0`.
    /// - `SharedString`: interned into the shared string table (`t="s"`).
    /// - `InlineStr`: written as an inline string (`t="inlineStr"`).
    /// - `Error`: `value` must start with `#`; written as `t="e"`.
    /// - `Stub`: written as an empty cell, ignoring `value`.
    /// - `FormulaStr`: not supported.
    ///
    /// The whole row is validated before anything is rendered or written, so an
    /// error leaves the stream untouched and the caller may keep writing rows.
    /// A valid row is rendered as bytes into one reused row buffer and handed
    /// to the ZIP stream with a single `write_all`.
    ///
    /// # Errors
    ///
    /// - [`ModernXlsxError::XmlWrite`] if no sheet is currently open.
    /// - [`ModernXlsxError::InvalidCellRef`] if the sheet already holds
    ///   1,048,576 rows or a cell would sit past column XFD (16,384 cells).
    /// - [`ModernXlsxError::InvalidCellValue`] for a non-finite or non-numeric
    ///   `Number`, an unrecognised `Boolean`, an `Error` value without a
    ///   leading `#`, or a `FormulaStr` cell.
    /// - [`ModernXlsxError::ZipWrite`] if the ZIP entry cannot be written.
    pub fn write_row(&mut self, cells: &[StreamingCell]) -> Result<()> {
        if !self.current_sheet_open {
            cold_path();
            return Err(ModernXlsxError::XmlWrite(
                "Failed to write row: no sheet is currently open — call start_sheet() before write_row()".into(),
            ));
        }
        if self.current_row > MAX_ROW {
            cold_path();
            return Err(ModernXlsxError::InvalidCellRef(format!(
                "row limit {} reached",
                MAX_ROW + 1
            )));
        }
        let row_num = self.current_row + 1;

        // Pass 1: validate the whole row. No I/O, no state change.
        for (col_idx, cell) in cells.iter().enumerate() {
            let ty = cell.cell_type.unwrap_or(CellType::Number);
            if cell.value.is_none() && !(ty == CellType::Stub && cell.style.is_some()) {
                continue;
            }
            if col_idx > MAX_COL as usize {
                cold_path();
                return Err(ModernXlsxError::InvalidCellRef(format!(
                    "column limit XFD exceeded at row {row_num} (a row holds at most {} cells)",
                    MAX_COL + 1
                )));
            }
            let val = cell.value.as_deref().unwrap_or("");
            let invalid = |msg: String| {
                cold_path();
                let mut letters = [0u8; 7];
                let len = col_to_letters_buf(col_idx as u32, &mut letters);
                ModernXlsxError::InvalidCellValue(format!(
                    "{} cell {}{row_num}: {msg}",
                    cell_type_label(ty),
                    String::from_utf8_lossy(&letters[..len])
                ))
            };
            match ty {
                CellType::Number => {
                    if !val.parse::<f64>().is_ok_and(f64::is_finite) {
                        return Err(invalid(format!("'{val}' is not a finite number")));
                    }
                }
                CellType::Boolean => {
                    if parse_bool(val).is_none() {
                        return Err(invalid(format!(
                            "'{val}' is not a boolean (use true, false, 1 or 0)"
                        )));
                    }
                }
                CellType::Error => {
                    if !val.starts_with('#') {
                        return Err(invalid(format!(
                            "'{val}' is not an error value (must start with '#')"
                        )));
                    }
                }
                CellType::FormulaStr => {
                    return Err(ModernXlsxError::InvalidCellValue(
                        "formulaStr is not supported by the streaming writer; use inlineStr or sharedString".into(),
                    ));
                }
                CellType::SharedString | CellType::InlineStr | CellType::Stub => {}
            }
        }

        // Pass 2: render into the reusable buffer.
        let mut buf = std::mem::take(&mut self.row_buf);
        buf.clear();
        let mut ibuf = itoa::Buffer::new();
        let row_str = ibuf.format(row_num).as_bytes();
        buf.extend_from_slice(b"<row r=\"");
        buf.extend_from_slice(row_str);
        buf.extend_from_slice(b"\">");

        for (col_idx, cell) in cells.iter().enumerate() {
            let ty = cell.cell_type.unwrap_or(CellType::Number);
            let val = match (cell.value.as_deref(), ty, cell.style) {
                (Some(v), _, _) => v,
                (None, CellType::Stub, Some(_)) => "",
                (None, _, _) => continue,
            };

            // Pass 1 guarantees col_idx <= MAX_COL for every cell rendered here.
            let mut letters = [0u8; 7];
            let letters_len = col_to_letters_buf(col_idx as u32, &mut letters);
            buf.extend_from_slice(b"<c r=\"");
            buf.extend_from_slice(&letters[..letters_len]);
            buf.extend_from_slice(row_str);
            buf.push(b'"');
            let ty_attr: &[u8] = match ty {
                CellType::SharedString => b" t=\"s\"",
                CellType::Number | CellType::Stub => b"",
                CellType::Boolean => b" t=\"b\"",
                CellType::Error => b" t=\"e\"",
                CellType::InlineStr | CellType::FormulaStr => b" t=\"inlineStr\"",
            };
            buf.extend_from_slice(ty_attr);
            if let Some(s) = cell.style {
                buf.extend_from_slice(b" s=\"");
                buf.extend_from_slice(itoa::Buffer::new().format(s).as_bytes());
                buf.push(b'"');
            }
            match ty {
                CellType::SharedString => {
                    let idx = self.intern_string(val);
                    buf.extend_from_slice(b"><v>");
                    buf.extend_from_slice(itoa::Buffer::new().format(idx).as_bytes());
                    buf.extend_from_slice(b"</v></c>");
                }
                CellType::Number => {
                    buf.extend_from_slice(b"><v>");
                    buf.extend_from_slice(val.as_bytes());
                    buf.extend_from_slice(b"</v></c>");
                }
                CellType::Boolean => {
                    buf.extend_from_slice(b"><v>");
                    buf.push(if parse_bool(val) == Some(true) {
                        b'1'
                    } else {
                        b'0'
                    });
                    buf.extend_from_slice(b"</v></c>");
                }
                CellType::Error => {
                    buf.extend_from_slice(b"><v>");
                    xml_escape_into(&mut buf, val);
                    buf.extend_from_slice(b"</v></c>");
                }
                CellType::Stub => buf.extend_from_slice(b"/>"),
                CellType::InlineStr | CellType::FormulaStr => {
                    buf.extend_from_slice(if needs_preserve(val) {
                        b"><is><t xml:space=\"preserve\">"
                    } else {
                        b"><is><t>"
                    });
                    xml_escape_into(&mut buf, val);
                    buf.extend_from_slice(b"</t></is></c>");
                }
            }
        }
        buf.extend_from_slice(b"</row>");

        let written = self.zip.write_all(&buf);
        self.row_buf = buf;
        written.map_err(zip_err)?;
        self.current_row = row_num;
        Ok(())
    }

    /// Close the currently open sheet (writes `</sheetData></worksheet>`).
    fn end_sheet(&mut self) -> Result<()> {
        write!(self.zip, "</sheetData></worksheet>").map_err(zip_err)?;
        self.current_sheet_open = false;
        Ok(())
    }

    /// Intern a string into the SST and return its zero-based index. Every
    /// call counts as one reference (the SST `count`).
    #[inline]
    fn intern_string(&mut self, s: &str) -> u32 {
        self.sst.insert(s) as u32
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
            return Err(ModernXlsxError::InvalidWorkbook(
                "Finish failed: workbook must contain at least one sheet — call start_sheet() before finish()".into(),
            ));
        }

        let opts = crate::zip::writer::deflate_options();
        let sc = self.sheet_count as usize;

        // --- [Content_Types].xml ---
        self.zip
            .start_file("[Content_Types].xml", opts)
            .map_err(zip_err)?;
        self.zip
            .write_all(&ContentTypes::for_basic_workbook(sc).to_xml()?)
            .map_err(zip_err)?;

        // --- _rels/.rels ---
        self.zip.start_file("_rels/.rels", opts).map_err(zip_err)?;
        write!(
            self.zip,
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\
             <Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\
             <Relationship Id=\"rId1\" \
             Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" \
             Target=\"xl/workbook.xml\"/>\
             </Relationships>"
        )
        .map_err(zip_err)?;

        // --- xl/workbook.xml ---
        self.zip
            .start_file("xl/workbook.xml", opts)
            .map_err(zip_err)?;
        write!(
            self.zip,
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\
             <workbook xmlns=\"{SPREADSHEET_NS}\" \
             xmlns:r=\"{REL_NS}\">\
             <sheets>"
        )
        .map_err(zip_err)?;
        for (i, name) in self.sheet_names.iter().enumerate() {
            let sheet_id = i + 1;
            write!(
                self.zip,
                "<sheet name=\"{}\" sheetId=\"{sheet_id}\" r:id=\"rId{sheet_id}\"/>",
                xml_escape(name)
            )
            .map_err(zip_err)?;
        }
        write!(self.zip, "</sheets></workbook>").map_err(zip_err)?;

        // --- xl/_rels/workbook.xml.rels ---
        self.zip
            .start_file("xl/_rels/workbook.xml.rels", opts)
            .map_err(zip_err)?;
        write!(
            self.zip,
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\
             <Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
        )
        .map_err(zip_err)?;
        for i in 1..=sc {
            write!(
                self.zip,
                "<Relationship Id=\"rId{i}\" \
                 Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" \
                 Target=\"worksheets/sheet{i}.xml\"/>"
            )
            .map_err(zip_err)?;
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
        .map_err(zip_err)?;

        // --- xl/sharedStrings.xml ---
        self.zip
            .start_file("xl/sharedStrings.xml", opts)
            .map_err(zip_err)?;
        // Streamed into the entry event by event, never buffered whole.
        self.sst.write_to(&mut self.zip)?;

        // --- xl/styles.xml ---
        self.zip
            .start_file("xl/styles.xml", opts)
            .map_err(zip_err)?;
        if let Some(ref custom) = self.styles_xml {
            self.zip.write_all(custom.as_bytes()).map_err(zip_err)?;
        } else {
            self.zip
                .write_all(MINIMAL_STYLES_XML.as_bytes())
                .map_err(zip_err)?;
        }

        // Finalize the ZIP archive.
        let cursor = self.zip.finish().map_err(|e| {
            ModernXlsxError::ZipFinalize(format!(
                "Failed to finalize streaming XLSX archive ({} sheets): {e}",
                self.sheet_count
            ))
        })?;
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

/// Parse a streaming boolean: `true`/`false` in any case, or `1`/`0`.
fn parse_bool(val: &str) -> Option<bool> {
    if val.eq_ignore_ascii_case("true") || val == "1" {
        Some(true)
    } else if val.eq_ignore_ascii_case("false") || val == "0" {
        Some(false)
    } else {
        None
    }
}

/// Human label of a cell type for error messages.
fn cell_type_label(ty: CellType) -> &'static str {
    match ty {
        CellType::Number => "Number",
        CellType::Boolean => "Boolean",
        CellType::Error => "Error",
        CellType::SharedString => "SharedString",
        CellType::InlineStr => "InlineStr",
        CellType::FormulaStr => "FormulaStr",
        CellType::Stub => "Stub",
    }
}

/// Maps any ZIP or I/O failure to [`ModernXlsxError::ZipWrite`].
#[cold]
fn zip_err(e: impl std::fmt::Display) -> ModernXlsxError {
    ModernXlsxError::ZipWrite(e.to_string())
}

/// XML-escape `s` (text content or attribute value) onto the end of `out`.
///
/// A CR is written as `&#13;` so it survives the XML 1.0 line-end
/// normalization a reader applies to literal CRs. Characters illegal in XML 1.0
/// (C0 control chars other than tab/LF/CR, plus the U+FFFE/U+FFFF
/// noncharacters) are dropped so a corrupt document is never emitted.
#[inline]
fn xml_escape_into(out: &mut Vec<u8>, s: &str) {
    for ch in s.chars() {
        match ch {
            '&' => out.extend_from_slice(b"&amp;"),
            '<' => out.extend_from_slice(b"&lt;"),
            '>' => out.extend_from_slice(b"&gt;"),
            '"' => out.extend_from_slice(b"&quot;"),
            '\'' => out.extend_from_slice(b"&apos;"),
            '\r' => out.extend_from_slice(b"&#13;"),
            c if crate::ooxml::is_illegal_xml_char(c) => {}
            c => out.extend_from_slice(c.encode_utf8(&mut [0u8; 4]).as_bytes()),
        }
    }
}

/// XML-escape a string; borrows `s` when nothing needs escaping or dropping,
/// otherwise allocates via [`xml_escape_into`].
#[inline]
fn xml_escape(s: &str) -> Cow<'_, str> {
    let clean = !s.chars().any(|c| {
        matches!(c, '&' | '<' | '>' | '"' | '\'' | '\r') || crate::ooxml::is_illegal_xml_char(c)
    });
    if clean {
        return Cow::Borrowed(s);
    }
    let mut out = Vec::with_capacity(s.len());
    xml_escape_into(&mut out, s);
    // Only whole UTF-8 chars and ASCII entities were appended.
    Cow::Owned(String::from_utf8(out).expect("xml_escape_into emits valid UTF-8"))
}

/// Minimal `xl/styles.xml` that defines one default font, two fills (none + gray125),
/// one border, and one cell format.
const MINIMAL_STYLES_XML: &str = "\
<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\
<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">\
<fonts count=\"1\"><font><sz val=\"11\"/><name val=\"Calibri\"/></font></fonts>\
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
    fn xml_escape_basic() {
        assert_eq!(xml_escape("hello"), "hello");
        assert_eq!(xml_escape("a&b"), "a&amp;b");
        assert_eq!(xml_escape("<tag>"), "&lt;tag&gt;");
        assert_eq!(xml_escape("\"quoted\""), "&quot;quoted&quot;");
        assert_eq!(xml_escape("it's"), "it&apos;s");
        assert_eq!(xml_escape("a\r\nb"), "a&#13;\nb");
    }

    /// A streamed CR is written as `&#13;`, so the reader's XML 1.0 line-end
    /// normalization keeps it.
    #[test]
    fn streaming_write_inline_string_keeps_cr() {
        let mut w = StreamingWriterCore::new();
        w.start_sheet("CR").unwrap();
        w.write_row(&[StreamingCell {
            value: Some("a\r\nb".into()),
            cell_type: Some(CellType::InlineStr),
            style: None,
        }])
        .unwrap();
        let bytes = w.finish().unwrap();

        let wb = crate::reader::read_xlsx(&bytes).unwrap();
        assert_eq!(
            wb.sheets[0].worksheet.rows[0].cells[0].value.as_deref(),
            Some("a\r\nb")
        );
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
            wb.sheets[0].worksheet.rows[0].cells[0].value.as_deref(),
            Some("1")
        );
        assert_eq!(
            wb.sheets[0].worksheet.rows[0].cells[1].value.as_deref(),
            Some("10")
        );
        assert_eq!(
            wb.sheets[0].worksheet.rows[99].cells[0].value.as_deref(),
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
            wb.sheets[0].worksheet.rows[0].cells[0].value.as_deref(),
            Some("Hello")
        );
        assert_eq!(
            wb.sheets[0].worksheet.rows[0].cells[1].value.as_deref(),
            Some("World")
        );
        assert_eq!(
            wb.sheets[0].worksheet.rows[1].cells[0].value.as_deref(),
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
            wb.sheets[0].worksheet.rows[0].cells[0].value.as_deref(),
            Some("1")
        );
        assert_eq!(
            wb.sheets[1].worksheet.rows[0].cells[0].value.as_deref(),
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
            wb.sheets[0].worksheet.rows[0].cells[0].value.as_deref(),
            Some("plain text value")
        );
        assert_eq!(
            wb.sheets[0].worksheet.rows[0].cells[1].value.as_deref(),
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
        assert_eq!(w.finish().unwrap_err().code(), "INVALID_WORKBOOK");
    }

    #[test]
    fn streaming_write_row_without_sheet_errors() {
        let mut w = StreamingWriterCore::new();
        let result = w.write_row(&[StreamingCell {
            value: Some("x".into()),
            cell_type: None,
            style: None,
        }]);
        assert_eq!(result.unwrap_err().code(), "XML_WRITE");
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
            wb.sheets[0].worksheet.rows[0].cells[0].value.as_deref(),
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
            wb.sheets[0].worksheet.rows[999].cells[0].value.as_deref(),
            Some("1000")
        );
        assert_eq!(
            wb.sheets[0].worksheet.rows[999].cells[1].value.as_deref(),
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
            wb.sheets[0].worksheet.rows[0].cells[0].value.as_deref(),
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
            wb.sheets[0].worksheet.rows[0].cells[0].value.as_deref(),
            Some("a")
        );
        assert_eq!(
            wb.sheets[1].worksheet.rows[0].cells[0].value.as_deref(),
            Some("b")
        );
    }

    // -- validate-then-render, finite numbers, t="e", stubs, limits --

    fn cell(value: Option<&str>, ct: Option<CellType>, style: Option<u32>) -> StreamingCell {
        StreamingCell {
            value: value.map(str::to_string),
            cell_type: ct,
            style,
        }
    }

    fn num(v: &str) -> StreamingCell {
        cell(Some(v), Some(CellType::Number), None)
    }

    fn assert_code(r: Result<()>, code: &str) {
        match r {
            Err(e) => assert_eq!(e.code(), code, "{e}"),
            Ok(()) => panic!("expected Err({code}), got Ok"),
        }
    }

    #[test]
    fn rs07_number_injection_rejected_and_stream_intact() {
        let mut w = StreamingWriterCore::new();
        w.start_sheet("S").unwrap();
        assert_code(
            w.write_row(&[num("1</v></c><c r=\"Z9\"><v>2")]),
            "INVALID_CELL_VALUE",
        );
        w.write_row(&[num("7")]).unwrap();
        let bytes = w.finish().unwrap();
        let wb = crate::reader::read_xlsx(&bytes).unwrap();
        let rows = &wb.sheets[0].worksheet.rows;
        assert_eq!(rows.len(), 1);
        assert_eq!(rows[0].cells[0].reference, "A1");
        assert_eq!(rows[0].cells[0].value.as_deref(), Some("7"));
    }

    #[test]
    fn rs07_non_finite_and_non_numeric_numbers_rejected() {
        for bad in ["NaN", "inf", "Infinity", "abc", "", "1e999"] {
            let mut w = StreamingWriterCore::new();
            w.start_sheet("S").unwrap();
            assert_code(w.write_row(&[num(bad)]), "INVALID_CELL_VALUE");
        }
    }

    #[test]
    fn rs07_number_written_verbatim() {
        let mut w = StreamingWriterCore::new();
        w.start_sheet("S").unwrap();
        w.write_row(&[num("1.5e3"), num("1")]).unwrap();
        let xml = part(&w.finish().unwrap(), "xl/worksheets/sheet1.xml");
        assert!(xml.contains("<v>1.5e3</v>"), "{xml}");
        assert!(xml.contains("<c r=\"B1\"><v>1</v></c>"), "{xml}");
    }

    #[test]
    fn rs07_error_cells() {
        let mut w = StreamingWriterCore::new();
        w.start_sheet("S").unwrap();
        w.write_row(&[cell(Some("#N/A"), Some(CellType::Error), None)])
            .unwrap();
        assert_code(
            w.write_row(&[cell(Some("N/A"), Some(CellType::Error), None)]),
            "INVALID_CELL_VALUE",
        );
        let bytes = w.finish().unwrap();
        assert!(
            part(&bytes, "xl/worksheets/sheet1.xml")
                .contains("<c r=\"A1\" t=\"e\"><v>#N/A</v></c>")
        );
        let wb = crate::reader::read_xlsx(&bytes).unwrap();
        let c = &wb.sheets[0].worksheet.rows[0].cells[0];
        assert_eq!(c.cell_type, CellType::Error);
        assert_eq!(c.value.as_deref(), Some("#N/A"));
    }

    #[test]
    fn rs07_stub_and_formula_str() {
        let mut w = StreamingWriterCore::new();
        w.start_sheet("S").unwrap();
        w.write_row(&[cell(None, Some(CellType::Stub), Some(0))])
            .unwrap();
        assert_code(
            w.write_row(&[cell(Some("=1+1"), Some(CellType::FormulaStr), None)]),
            "INVALID_CELL_VALUE",
        );
        let xml = part(&w.finish().unwrap(), "xl/worksheets/sheet1.xml");
        assert!(xml.contains("<c r=\"A1\" s=\"0\"/>"), "{xml}");
    }

    #[test]
    fn rs07_bad_sheet_names_rejected() {
        let long = "x".repeat(32);
        for bad in ["a[b", "", long.as_str(), "a/b", "'x", "x'"] {
            let mut w = StreamingWriterCore::new();
            assert_code(w.start_sheet(bad), "INVALID_CELL_VALUE");
        }
    }

    #[test]
    fn rs07_duplicate_sheet_name_keeps_open_sheet() {
        let mut w = StreamingWriterCore::new();
        w.start_sheet("S").unwrap();
        assert_code(w.start_sheet("s"), "INVALID_CELL_VALUE");
        w.write_row(&[num("5")]).unwrap();
        let wb = crate::reader::read_xlsx(&w.finish().unwrap()).unwrap();
        assert_eq!(wb.sheets.len(), 1);
        assert_eq!(wb.sheets[0].name, "S");
        assert_eq!(wb.sheets[0].worksheet.rows.len(), 1);
    }

    #[test]
    fn duplicate_sheet_names_are_found_by_lower_case_key() {
        let mut w = StreamingWriterCore::new();
        // A name refused for another reason is not recorded.
        assert_code(w.start_sheet("Ünïcödé/1"), "INVALID_CELL_VALUE");
        w.start_sheet("Ünïcödé 1").unwrap();
        w.start_sheet("Second").unwrap();
        for dup in ["üNÏCÖDÉ 1", "Ünïcödé 1", "SECOND"] {
            let err = w.start_sheet(dup).unwrap_err();
            let first = if dup == "SECOND" {
                "Second"
            } else {
                "Ünïcödé 1"
            };
            assert_eq!(
                err.to_string(),
                format!("invalid cell value: Invalid sheet name '{dup}': duplicate of '{first}'")
            );
        }
        w.start_sheet("Ünïcödé 1 ok").unwrap();
        w.write_row(&[num("5")]).unwrap();
        let wb = crate::reader::read_xlsx(&w.finish().unwrap()).unwrap();
        let names: Vec<&str> = wb.sheets.iter().map(|s| s.name.as_str()).collect();
        assert_eq!(names, ["Ünïcödé 1", "Second", "Ünïcödé 1 ok"]);
    }

    #[test]
    fn rs07_column_limit() {
        let mut w = StreamingWriterCore::new();
        w.start_sheet("S").unwrap();
        let too_wide = vec![num("1"); 16_385];
        assert_code(w.write_row(&too_wide), "INVALID_CELL_REF");
        let max_wide = vec![num("1"); 16_384];
        w.write_row(&max_wide).unwrap();
        let xml = part(&w.finish().unwrap(), "xl/worksheets/sheet1.xml");
        assert!(xml.contains("<row r=\"1\">"), "{xml}");
        assert!(xml.contains("r=\"XFD1\""), "{xml}");
        assert!(!xml.contains("XFE"), "{xml}");
    }

    #[test]
    fn rs07_boolean_case_insensitive() {
        let b = |v: &str| cell(Some(v), Some(CellType::Boolean), None);
        let mut w = StreamingWriterCore::new();
        w.start_sheet("S").unwrap();
        w.write_row(&[b("TRUE"), b("True"), b("1"), b("FALSE"), b("0")])
            .unwrap();
        assert_code(w.write_row(&[b("yes")]), "INVALID_CELL_VALUE");
        let xml = part(&w.finish().unwrap(), "xl/worksheets/sheet1.xml");
        assert_eq!(xml.matches("t=\"b\"><v>1</v>").count(), 3, "{xml}");
        assert_eq!(xml.matches("t=\"b\"><v>0</v>").count(), 2, "{xml}");
    }

    fn part(bytes: &[u8], name: &str) -> String {
        use std::io::Read;
        let mut archive = zip::ZipArchive::new(Cursor::new(bytes)).unwrap();
        let mut xml = String::new();
        archive
            .by_name(name)
            .unwrap()
            .read_to_string(&mut xml)
            .unwrap();
        xml
    }

    /// Inline strings with edge whitespace carry `xml:space="preserve"`.
    #[test]
    fn inline_str_edge_whitespace_preserved() {
        let mut w = StreamingWriterCore::new();
        w.start_sheet("S").unwrap();
        w.write_row(&[
            cell(Some("  pad "), Some(CellType::InlineStr), None),
            cell(Some("plain"), Some(CellType::InlineStr), None),
        ])
        .unwrap();
        let bytes = w.finish().unwrap();
        let xml = part(&bytes, "xl/worksheets/sheet1.xml");
        assert!(xml.contains(
            "<c r=\"A1\" t=\"inlineStr\"><is><t xml:space=\"preserve\">  pad </t></is></c>"
        ));
        assert!(xml.contains("<c r=\"B1\" t=\"inlineStr\"><is><t>plain</t></is></c>"));
        let wb = crate::reader::read_xlsx(&bytes).unwrap();
        assert_eq!(
            wb.sheets[0].worksheet.rows[0].cells[0].value.as_deref(),
            Some("  pad ")
        );
    }

    /// Pins the exact bytes the row-buffer renderer emits for a mixed row.
    #[test]
    fn mixed_row_bytes_pinned() {
        let mut w = StreamingWriterCore::new();
        w.start_sheet("S").unwrap();
        w.write_row(&[
            cell(Some("hi"), Some(CellType::SharedString), Some(1)),
            num("1.5"),
            cell(Some("1"), Some(CellType::Boolean), None),
            cell(Some("a<b&c"), Some(CellType::InlineStr), None),
            cell(None, Some(CellType::Number), None),
            num("2"),
        ])
        .unwrap();
        w.write_row(&[cell(Some("hi"), Some(CellType::SharedString), None)])
            .unwrap();
        let bytes = w.finish().unwrap();
        assert_eq!(
            part(&bytes, "xl/worksheets/sheet1.xml"),
            concat!(
                "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>",
                "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" ",
                "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">",
                "<sheetViews><sheetView workbookViewId=\"0\"/></sheetViews>",
                "<sheetData>",
                "<row r=\"1\">",
                "<c r=\"A1\" t=\"s\" s=\"1\"><v>0</v></c>",
                "<c r=\"B1\"><v>1.5</v></c>",
                "<c r=\"C1\" t=\"b\"><v>1</v></c>",
                "<c r=\"D1\" t=\"inlineStr\"><is><t>a&lt;b&amp;c</t></is></c>",
                "<c r=\"F1\"><v>2</v></c>",
                "</row>",
                "<row r=\"2\"><c r=\"A2\" t=\"s\"><v>0</v></c></row>",
                "</sheetData></worksheet>"
            )
        );
        assert_eq!(
            part(&bytes, "xl/sharedStrings.xml"),
            concat!(
                "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>",
                "<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" ",
                "count=\"2\" uniqueCount=\"1\"><si><t>hi</t></si></sst>"
            )
        );
    }

    /// Every streamed sheet carries the view Excel writes for a sheet: without
    /// one, Excel on a 150% display (144 DPI) showed setRowHeight(1, 40) as
    /// 26.7 points; with it, 40.
    #[test]
    fn every_sheet_has_excels_default_view_before_its_rows() {
        let mut w = StreamingWriterCore::new();
        w.start_sheet("A").unwrap();
        w.write_row(&[num("1")]).unwrap();
        w.start_sheet("B").unwrap();
        let bytes = w.finish().unwrap();
        for name in ["xl/worksheets/sheet1.xml", "xl/worksheets/sheet2.xml"] {
            let xml = part(&bytes, name);
            assert!(
                xml.contains(
                    "<sheetViews><sheetView workbookViewId=\"0\"/></sheetViews><sheetData>"
                ),
                "{name}: {xml}"
            );
            assert_eq!(xml.matches("<sheetViews>").count(), 1, "{name}: {xml}");
        }
        let wb = crate::reader::read_xlsx(&bytes).unwrap();
        assert!(wb.sheets.iter().all(|s| s.worksheet.sheet_view.is_none()));
    }

    #[test]
    fn rejected_wide_row_writes_nothing_and_later_row_works() {
        let mut w = StreamingWriterCore::new();
        w.start_sheet("S").unwrap();
        let mut wide = vec![cell(None, None, None); 16_385];
        wide[16_384] = num("1");
        assert_code(w.write_row(&wide), "INVALID_CELL_REF");
        w.write_row(&[num("7")]).unwrap();
        let xml = part(&w.finish().unwrap(), "xl/worksheets/sheet1.xml");
        assert!(
            xml.contains("<sheetData><row r=\"1\"><c r=\"A1\"><v>7</v></c></row></sheetData>"),
            "{xml}"
        );
    }

    #[test]
    fn sst_dedup_stores_once() {
        let mut w = StreamingWriterCore::new();
        w.start_sheet("S").unwrap();
        for _ in 0..3 {
            w.write_row(&[cell(Some("hi"), Some(CellType::SharedString), None)])
                .unwrap();
        }
        let bytes = w.finish().unwrap();
        assert_eq!(
            part(&bytes, "xl/sharedStrings.xml").matches("<si>").count(),
            1
        );
    }

    #[test]
    fn manifest_is_the_basic_workbook_manifest() {
        use crate::ooxml::content_types::ContentTypes;
        let mut w = StreamingWriterCore::new();
        for name in ["A", "B", "C"] {
            w.start_sheet(name).unwrap();
            w.write_row(&[num("1")]).unwrap();
        }
        let xml = part(&w.finish().unwrap(), "[Content_Types].xml");
        let streamed = ContentTypes::parse(xml.as_bytes()).unwrap();
        let basic = ContentTypes::for_basic_workbook(3);
        assert_eq!(streamed.defaults, basic.defaults);
        assert_eq!(streamed.overrides, basic.overrides);
    }

    #[test]
    fn sst_space_behind_a_dropped_control_char_is_preserved() {
        let mut w = StreamingWriterCore::new();
        w.start_sheet("S").unwrap();
        w.write_row(&[cell(
            Some("\u{1} total"),
            Some(CellType::SharedString),
            None,
        )])
        .unwrap();
        let sst = part(&w.finish().unwrap(), "xl/sharedStrings.xml");
        assert!(
            sst.contains("<si><t xml:space=\"preserve\"> total</t></si>"),
            "{sst}"
        );
    }

    #[test]
    fn sst_count_is_the_reference_count() {
        let mut w = StreamingWriterCore::new();
        w.start_sheet("S").unwrap();
        w.write_row(&[
            cell(Some("x"), Some(CellType::SharedString), None),
            cell(Some("x"), Some(CellType::SharedString), None),
            cell(Some("x"), Some(CellType::SharedString), None),
        ])
        .unwrap();
        let bytes = w.finish().unwrap();
        let sst = part(&bytes, "xl/sharedStrings.xml");
        assert!(sst.contains("count=\"3\" uniqueCount=\"1\""), "{sst}");
    }

    #[test]
    fn rs07_row_limit() {
        let mut w = StreamingWriterCore::new();
        w.start_sheet("S").unwrap();
        w.current_row = 1_048_576;
        assert_code(w.write_row(&[num("1")]), "INVALID_CELL_REF");
    }
}

#[cfg(test)]
mod prop_tests {
    use super::*;
    use proptest::prelude::*;

    /// Strings guaranteed free of characters illegal in XML 1.0, so that
    /// `xml_escape` performs a pure entity-escape (drops nothing) and therefore
    /// round-trips through `unescape`.
    fn xml_legal_string() -> impl Strategy<Value = String> {
        ".*".prop_filter("input must contain no XML-illegal chars", |s| {
            !s.chars().any(crate::ooxml::is_illegal_xml_char)
        })
    }

    proptest! {
        /// `xml_escape` never emits any character illegal in XML 1.0 — the C0
        /// control chars (except tab/LF/CR) and the U+FFFE/U+FFFF noncharacters.
        #[test]
        fn xml_escape_never_emits_xml_illegal_chars(s in ".*") {
            let escaped = xml_escape(&s);
            prop_assert!(
                !escaped.chars().any(crate::ooxml::is_illegal_xml_char),
                "escaped output retained an XML-1.0-illegal character"
            );
        }

        /// For input already free of illegal chars, escape -> unescape round-trips.
        #[test]
        fn xml_escape_roundtrip(s in xml_legal_string()) {
            let escaped = xml_escape(&s);
            let unescaped = quick_xml::escape::unescape(&escaped).unwrap();
            prop_assert_eq!(unescaped.as_ref(), s.as_str());
        }
    }

    /// The U+FFFE and U+FFFF noncharacters are explicitly dropped — they are
    /// legal Unicode scalars (so a parser would accept the bytes) yet illegal in
    /// an XML 1.0 document, and the `.*` strategy above is vanishingly unlikely
    /// to generate them.
    #[test]
    fn xml_escape_drops_noncharacters() {
        assert_eq!(xml_escape("a\u{FFFE}b\u{FFFF}c"), "abc");
    }
}
