use std::collections::HashMap;

use quick_xml::events::{BytesDecl, BytesEnd, BytesStart, BytesText, Event};
use quick_xml::{Reader, Writer};
use serde::{Deserialize, Serialize};

use crate::{ModernXlsxError, Result};

/// A single run of formatted text within a rich text shared string.
///
/// Each run carries its text content plus optional formatting properties that
/// were found in the `<rPr>` element (bold, italic, font name, size, color).
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct RichTextRun {
    pub text: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub bold: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub italic: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub font_name: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub font_size: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub color: Option<String>,
}

/// A read-only Shared String Table parsed from `sharedStrings.xml`.
///
/// In OOXML, cells that contain string values store an index into this table
/// rather than the string itself.
///
/// The `strings` vector always contains the plain-text value for every entry
/// (for rich text entries this is the concatenation of all run texts).
///
/// The `rich_runs` vector is parallel to `strings`: if entry *i* was a rich
/// text `<si>` with `<r>` elements, `rich_runs[i]` is `Some(runs)`;
/// otherwise it is `None`.
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct SharedStringTable {
    pub strings: Vec<String>,
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub rich_runs: Vec<Option<Vec<RichTextRun>>>,
}

impl SharedStringTable {
    /// Parse a Shared String Table from the raw XML bytes of `sharedStrings.xml`.
    pub fn parse(data: &[u8]) -> Result<Self> {
        let mut reader = Reader::from_reader(data);
        reader.config_mut().trim_text(false);

        let mut buf = Vec::new();
        let mut strings = Vec::new();
        let mut rich_runs: Vec<Option<Vec<RichTextRun>>> = Vec::new();

        // State tracking
        let mut in_si = false;
        let mut in_t = false;
        let mut in_r = false;
        let mut in_rpr = false;
        let mut current_text = String::new();

        // Rich text accumulation for the current <si>
        let mut runs: Vec<RichTextRun> = Vec::new();
        let mut run_text = String::new();
        let mut run_bold: Option<bool> = None;
        let mut run_italic: Option<bool> = None;
        let mut run_font_name: Option<String> = None;
        let mut run_font_size: Option<f64> = None;
        let mut run_color: Option<String> = None;

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(e)) => {
                    let local = e.local_name();
                    match local.as_ref() {
                        b"si" => {
                            in_si = true;
                            current_text.clear();
                            runs.clear();
                        }
                        b"r" if in_si => {
                            in_r = true;
                            run_text.clear();
                            run_bold = None;
                            run_italic = None;
                            run_font_name = None;
                            run_font_size = None;
                            run_color = None;
                        }
                        b"rPr" if in_r => {
                            in_rpr = true;
                        }
                        b"t" if in_si => {
                            in_t = true;
                        }
                        b"rFont" if in_rpr => {
                            for attr in e.attributes().flatten() {
                                if attr.key.as_ref() == b"val" {
                                    run_font_name = Some(
                                        String::from_utf8_lossy(&attr.value)
                                            .into_owned(),
                                    );
                                }
                            }
                        }
                        b"sz" if in_rpr => {
                            for attr in e.attributes().flatten() {
                                if attr.key.as_ref() == b"val"
                                    && let Ok(s) =
                                        std::str::from_utf8(&attr.value)
                                    {
                                        run_font_size = s.parse::<f64>().ok();
                                    }
                            }
                        }
                        b"color" if in_rpr => {
                            for attr in e.attributes().flatten() {
                                if attr.key.as_ref() == b"rgb" {
                                    run_color = Some(
                                        String::from_utf8_lossy(&attr.value)
                                            .into_owned(),
                                    );
                                }
                            }
                        }
                        _ => {}
                    }
                }
                Ok(Event::Empty(e)) => {
                    let local = e.local_name();
                    match local.as_ref() {
                        b"b" if in_rpr => {
                            run_bold = Some(true);
                        }
                        b"i" if in_rpr => {
                            run_italic = Some(true);
                        }
                        b"rFont" if in_rpr => {
                            for attr in e.attributes().flatten() {
                                if attr.key.as_ref() == b"val" {
                                    run_font_name = Some(
                                        String::from_utf8_lossy(&attr.value)
                                            .into_owned(),
                                    );
                                }
                            }
                        }
                        b"sz" if in_rpr => {
                            for attr in e.attributes().flatten() {
                                if attr.key.as_ref() == b"val"
                                    && let Ok(s) =
                                        std::str::from_utf8(&attr.value)
                                    {
                                        run_font_size = s.parse::<f64>().ok();
                                    }
                            }
                        }
                        b"color" if in_rpr => {
                            for attr in e.attributes().flatten() {
                                if attr.key.as_ref() == b"rgb" {
                                    run_color = Some(
                                        String::from_utf8_lossy(&attr.value)
                                            .into_owned(),
                                    );
                                }
                            }
                        }
                        _ => {}
                    }
                }
                Ok(Event::Text(e)) if in_t => {
                    let text =
                        std::str::from_utf8(e.as_ref()).unwrap_or_default();
                    current_text.push_str(text);
                    if in_r {
                        run_text.push_str(text);
                    }
                }
                Ok(Event::End(e)) => {
                    let local = e.local_name();
                    match local.as_ref() {
                        b"t" => {
                            in_t = false;
                        }
                        b"rPr" => {
                            in_rpr = false;
                        }
                        b"r" => {
                            runs.push(RichTextRun {
                                text: std::mem::take(&mut run_text),
                                bold: run_bold,
                                italic: run_italic,
                                font_name: run_font_name.take(),
                                font_size: run_font_size.take(),
                                color: run_color.take(),
                            });
                            in_r = false;
                        }
                        b"si" => {
                            strings
                                .push(std::mem::take(&mut current_text));
                            if runs.is_empty() {
                                rich_runs.push(None);
                            } else {
                                rich_runs
                                    .push(Some(std::mem::take(&mut runs)));
                            }
                            in_si = false;
                        }
                        _ => {}
                    }
                }
                Ok(Event::Eof) => break,
                Err(err) => {
                    return Err(ModernXlsxError::XmlParse(err.to_string()))
                }
                _ => {}
            }
            buf.clear();
        }

        // If no rich text was found at all, keep rich_runs empty to avoid
        // serializing an all-None vector.
        if rich_runs.iter().all(|r| r.is_none()) {
            rich_runs.clear();
        }

        Ok(SharedStringTable {
            strings,
            rich_runs,
        })
    }

    /// Return an empty Shared String Table.
    ///
    /// Used when the workbook does not contain a `sharedStrings.xml` part.
    pub fn empty() -> Self {
        SharedStringTable {
            strings: Vec::new(),
            rich_runs: Vec::new(),
        }
    }

    /// Return the rich text runs for the entry at `index`, if it is a rich
    /// text entry.
    pub fn get_rich_runs(&self, index: usize) -> Option<&[RichTextRun]> {
        self.rich_runs
            .get(index)
            .and_then(|opt| opt.as_deref())
    }

    /// Return the number of strings in the table.
    pub fn len(&self) -> usize {
        self.strings.len()
    }

    /// Return `true` if the table contains no strings.
    pub fn is_empty(&self) -> bool {
        self.strings.is_empty()
    }

    /// Look up a string by its zero-based index.
    pub fn get(&self, index: usize) -> Option<&str> {
        self.strings.get(index).map(|s| s.as_str())
    }

    /// Serialize the shared string table to XML bytes.
    ///
    /// Rich text entries are written using `<r>` / `<rPr>` elements; plain
    /// text entries use a simple `<si><t>…</t></si>` form.
    pub fn to_xml(&self) -> Result<Vec<u8>> {
        let mut buf: Vec<u8> =
            Vec::with_capacity(256 + self.strings.len() * 64);
        let mut writer = Writer::new(&mut buf);

        let map_write_err =
            |e: std::io::Error| ModernXlsxError::XmlWrite(e.to_string());

        // XML declaration.
        writer
            .write_event(Event::Decl(BytesDecl::new(
                "1.0",
                Some("UTF-8"),
                Some("yes"),
            )))
            .map_err(map_write_err)?;

        // <sst xmlns="..." count="N" uniqueCount="N">
        let mut ibuf = itoa::Buffer::new();
        let count_str = ibuf.format(self.strings.len());
        let mut sst_start = BytesStart::new("sst");
        sst_start.push_attribute((
            "xmlns",
            "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
        ));
        sst_start.push_attribute(("count", count_str));
        sst_start.push_attribute(("uniqueCount", count_str));
        writer
            .write_event(Event::Start(sst_start))
            .map_err(map_write_err)?;

        for (i, s) in self.strings.iter().enumerate() {
            let runs = self
                .rich_runs
                .get(i)
                .and_then(|opt| opt.as_ref());

            writer
                .write_event(Event::Start(BytesStart::new("si")))
                .map_err(map_write_err)?;

            if let Some(runs) = runs {
                // Rich text entry — write each <r>.
                for run in runs {
                    writer
                        .write_event(Event::Start(BytesStart::new("r")))
                        .map_err(map_write_err)?;

                    // Write <rPr> only if there is formatting.
                    if run.bold.is_some()
                        || run.italic.is_some()
                        || run.font_name.is_some()
                        || run.font_size.is_some()
                        || run.color.is_some()
                    {
                        writer
                            .write_event(Event::Start(BytesStart::new(
                                "rPr",
                            )))
                            .map_err(map_write_err)?;

                        if run.bold == Some(true) {
                            writer
                                .write_event(Event::Empty(
                                    BytesStart::new("b"),
                                ))
                                .map_err(map_write_err)?;
                        }
                        if run.italic == Some(true) {
                            writer
                                .write_event(Event::Empty(
                                    BytesStart::new("i"),
                                ))
                                .map_err(map_write_err)?;
                        }
                        if let Some(size) = run.font_size {
                            let mut sz = BytesStart::new("sz");
                            let size_str = size.to_string();
                            sz.push_attribute((
                                "val",
                                size_str.as_str(),
                            ));
                            writer
                                .write_event(Event::Empty(sz))
                                .map_err(map_write_err)?;
                        }
                        if let Some(ref color) = run.color {
                            let mut c = BytesStart::new("color");
                            c.push_attribute(("rgb", color.as_str()));
                            writer
                                .write_event(Event::Empty(c))
                                .map_err(map_write_err)?;
                        }
                        if let Some(ref name) = run.font_name {
                            let mut f = BytesStart::new("rFont");
                            f.push_attribute(("val", name.as_str()));
                            writer
                                .write_event(Event::Empty(f))
                                .map_err(map_write_err)?;
                        }

                        writer
                            .write_event(Event::End(BytesEnd::new("rPr")))
                            .map_err(map_write_err)?;
                    }

                    // <t>text</t>
                    writer
                        .write_event(Event::Start(BytesStart::new("t")))
                        .map_err(map_write_err)?;
                    writer
                        .write_event(Event::Text(BytesText::new(
                            &run.text,
                        )))
                        .map_err(map_write_err)?;
                    writer
                        .write_event(Event::End(BytesEnd::new("t")))
                        .map_err(map_write_err)?;

                    writer
                        .write_event(Event::End(BytesEnd::new("r")))
                        .map_err(map_write_err)?;
                }
            } else {
                // Plain text entry.
                writer
                    .write_event(Event::Start(BytesStart::new("t")))
                    .map_err(map_write_err)?;
                writer
                    .write_event(Event::Text(BytesText::new(s)))
                    .map_err(map_write_err)?;
                writer
                    .write_event(Event::End(BytesEnd::new("t")))
                    .map_err(map_write_err)?;
            }

            writer
                .write_event(Event::End(BytesEnd::new("si")))
                .map_err(map_write_err)?;
        }

        // </sst>
        writer
            .write_event(Event::End(BytesEnd::new("sst")))
            .map_err(map_write_err)?;

        Ok(buf)
    }
}

/// A builder for constructing a Shared String Table when writing an XLSX file.
///
/// Strings are deduplicated: inserting the same string twice returns the same
/// index both times.
#[derive(Debug, Clone)]
pub struct SharedStringTableBuilder {
    strings: Vec<String>,
    index: HashMap<String, usize>,
}

impl SharedStringTableBuilder {
    /// Create a new, empty builder.
    pub fn new() -> Self {
        SharedStringTableBuilder {
            strings: Vec::new(),
            index: HashMap::new(),
        }
    }

    /// Return the number of unique strings inserted so far.
    pub fn len(&self) -> usize {
        self.strings.len()
    }

    /// Return `true` if no strings have been inserted.
    pub fn is_empty(&self) -> bool {
        self.strings.is_empty()
    }

    /// Look up the index of a previously inserted string.
    ///
    /// Returns `None` if the string has not been inserted.
    pub fn get_index(&self, s: &str) -> Option<usize> {
        self.index.get(s).copied()
    }

    /// Insert a string and return its zero-based index.
    ///
    /// If the string has been inserted before, the existing index is returned
    /// without adding a duplicate entry.
    pub fn insert(&mut self, s: &str) -> usize {
        if let Some(&idx) = self.index.get(s) {
            return idx;
        }
        let idx = self.strings.len();
        let owned = s.to_owned();
        self.index.insert(owned.clone(), idx);
        self.strings.push(owned);
        idx
    }

    /// Serialize the shared string table to XML bytes suitable for
    /// inclusion in an XLSX archive as `xl/sharedStrings.xml`.
    pub fn to_xml(&self) -> Result<Vec<u8>> {
        let mut buf: Vec<u8> = Vec::with_capacity(256 + self.strings.len() * 64);
        let mut writer = Writer::new(&mut buf);

        let map_write_err =
            |e: std::io::Error| ModernXlsxError::XmlWrite(e.to_string());

        // XML declaration.
        writer
            .write_event(Event::Decl(BytesDecl::new(
                "1.0",
                Some("UTF-8"),
                Some("yes"),
            )))
            .map_err(map_write_err)?;

        // <sst xmlns="..." count="N" uniqueCount="N">
        let mut ibuf = itoa::Buffer::new();
        let count_str = ibuf.format(self.strings.len());
        let mut sst_start = BytesStart::new("sst");
        sst_start.push_attribute((
            "xmlns",
            "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
        ));
        sst_start.push_attribute(("count", count_str));
        sst_start.push_attribute(("uniqueCount", count_str));
        writer
            .write_event(Event::Start(sst_start))
            .map_err(map_write_err)?;

        // Each string: <si><t>text</t></si>
        for s in &self.strings {
            writer
                .write_event(Event::Start(BytesStart::new("si")))
                .map_err(map_write_err)?;
            writer
                .write_event(Event::Start(BytesStart::new("t")))
                .map_err(map_write_err)?;
            writer
                .write_event(Event::Text(BytesText::new(s)))
                .map_err(map_write_err)?;
            writer
                .write_event(Event::End(BytesEnd::new("t")))
                .map_err(map_write_err)?;
            writer
                .write_event(Event::End(BytesEnd::new("si")))
                .map_err(map_write_err)?;
        }

        // </sst>
        writer
            .write_event(Event::End(BytesEnd::new("sst")))
            .map_err(map_write_err)?;

        Ok(buf)
    }
}

impl Default for SharedStringTableBuilder {
    fn default() -> Self {
        Self::new()
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use pretty_assertions::assert_eq;

    #[test]
    fn test_parse_sst() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="3" uniqueCount="3">
  <si><t>Hello</t></si>
  <si><t>World</t></si>
  <si><t>Rust</t></si>
</sst>"#;

        let sst = SharedStringTable::parse(xml.as_bytes()).unwrap();
        assert_eq!(sst.len(), 3);
        assert!(!sst.is_empty());
        assert_eq!(sst.get(0), Some("Hello"));
        assert_eq!(sst.get(1), Some("World"));
        assert_eq!(sst.get(2), Some("Rust"));
    }

    #[test]
    fn test_sst_builder() {
        let mut builder = SharedStringTableBuilder::new();
        let i0 = builder.insert("Hello");
        let i1 = builder.insert("World");
        let i2 = builder.insert("Hello"); // duplicate

        assert_eq!(i0, 0);
        assert_eq!(i1, 1);
        assert_eq!(i2, 0); // dedup: same index as first "Hello"
        assert_eq!(builder.len(), 2);
    }

    #[test]
    fn test_sst_write_roundtrip() {
        let mut builder = SharedStringTableBuilder::new();
        builder.insert("Alpha");
        builder.insert("Beta");
        builder.insert("Gamma");
        builder.insert("Alpha"); // duplicate — should not add a new entry

        assert_eq!(builder.len(), 3);

        let xml = builder.to_xml().unwrap();
        let sst = SharedStringTable::parse(&xml).unwrap();

        assert_eq!(sst.len(), 3);
        assert_eq!(sst.get(0), Some("Alpha"));
        assert_eq!(sst.get(1), Some("Beta"));
        assert_eq!(sst.get(2), Some("Gamma"));
    }

    #[test]
    fn test_sst_out_of_bounds() {
        let sst = SharedStringTable::empty();
        assert_eq!(sst.get(999), None);

        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1">
  <si><t>Only</t></si>
</sst>"#;
        let sst = SharedStringTable::parse(xml.as_bytes()).unwrap();
        assert_eq!(sst.get(0), Some("Only"));
        assert_eq!(sst.get(999), None);
    }

    #[test]
    fn test_sst_empty() {
        let builder = SharedStringTableBuilder::new();
        assert!(builder.is_empty());

        let xml = builder.to_xml().unwrap();
        let sst = SharedStringTable::parse(&xml).unwrap();

        assert!(sst.is_empty());
        assert_eq!(sst.len(), 0);
    }

    #[test]
    fn test_parse_rich_text() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="2" uniqueCount="2">
  <si><t>Plain text</t></si>
  <si>
    <r>
      <rPr><b/><sz val="14"/><color rgb="FF0000"/><rFont val="Arial"/></rPr>
      <t>Bold red text</t>
    </r>
    <r>
      <t> normal text</t>
    </r>
  </si>
</sst>"#;

        let sst = SharedStringTable::parse(xml.as_bytes()).unwrap();
        assert_eq!(sst.len(), 2);

        // First entry is plain text.
        assert_eq!(sst.get(0), Some("Plain text"));
        assert!(sst.get_rich_runs(0).is_none());

        // Second entry is rich text — plain text is the concatenation.
        assert_eq!(sst.get(1), Some("Bold red text normal text"));

        let runs = sst.get_rich_runs(1).expect("should have rich runs");
        assert_eq!(runs.len(), 2);

        // First run: bold, font size 14, red, Arial.
        assert_eq!(runs[0].text, "Bold red text");
        assert_eq!(runs[0].bold, Some(true));
        assert_eq!(runs[0].italic, None);
        assert_eq!(runs[0].font_size, Some(14.0));
        assert_eq!(runs[0].color.as_deref(), Some("FF0000"));
        assert_eq!(runs[0].font_name.as_deref(), Some("Arial"));

        // Second run: no formatting.
        assert_eq!(runs[1].text, " normal text");
        assert_eq!(runs[1].bold, None);
        assert_eq!(runs[1].italic, None);
        assert_eq!(runs[1].font_size, None);
        assert_eq!(runs[1].color, None);
        assert_eq!(runs[1].font_name, None);
    }

    #[test]
    fn test_parse_rich_text_italic() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1">
  <si>
    <r>
      <rPr><i/></rPr>
      <t>italic</t>
    </r>
  </si>
</sst>"#;

        let sst = SharedStringTable::parse(xml.as_bytes()).unwrap();
        assert_eq!(sst.len(), 1);
        assert_eq!(sst.get(0), Some("italic"));

        let runs = sst.get_rich_runs(0).expect("should have rich runs");
        assert_eq!(runs.len(), 1);
        assert_eq!(runs[0].italic, Some(true));
        assert_eq!(runs[0].bold, None);
    }

    #[test]
    fn test_rich_text_roundtrip_via_sst_to_xml() {
        // Build an SST with both plain and rich text entries.
        let sst = SharedStringTable {
            strings: vec![
                "Plain".to_string(),
                "Bold italic".to_string(),
            ],
            rich_runs: vec![
                None,
                Some(vec![
                    RichTextRun {
                        text: "Bold".to_string(),
                        bold: Some(true),
                        italic: None,
                        font_name: Some("Calibri".to_string()),
                        font_size: Some(12.0),
                        color: Some("FF0000".to_string()),
                    },
                    RichTextRun {
                        text: " italic".to_string(),
                        bold: None,
                        italic: Some(true),
                        font_name: None,
                        font_size: None,
                        color: None,
                    },
                ]),
            ],
        };

        // Write to XML and re-parse.
        let xml = sst.to_xml().unwrap();
        let sst2 = SharedStringTable::parse(&xml).unwrap();

        assert_eq!(sst2.len(), 2);

        // Plain entry preserved.
        assert_eq!(sst2.get(0), Some("Plain"));
        assert!(sst2.get_rich_runs(0).is_none());

        // Rich entry preserved.
        assert_eq!(sst2.get(1), Some("Bold italic"));
        let runs = sst2.get_rich_runs(1).expect("should have rich runs");
        assert_eq!(runs.len(), 2);
        assert_eq!(runs[0].text, "Bold");
        assert_eq!(runs[0].bold, Some(true));
        assert_eq!(runs[0].font_name.as_deref(), Some("Calibri"));
        assert_eq!(runs[0].font_size, Some(12.0));
        assert_eq!(runs[0].color.as_deref(), Some("FF0000"));
        assert_eq!(runs[1].text, " italic");
        assert_eq!(runs[1].italic, Some(true));
    }

    #[test]
    fn test_plain_only_sst_has_empty_rich_runs() {
        // When all entries are plain, rich_runs should be empty (not
        // a vector of Nones) so that serialization stays compact.
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="2" uniqueCount="2">
  <si><t>A</t></si>
  <si><t>B</t></si>
</sst>"#;

        let sst = SharedStringTable::parse(xml.as_bytes()).unwrap();
        assert!(sst.rich_runs.is_empty(), "rich_runs should be empty for plain-only SST");
    }

    #[test]
    fn test_rich_text_serde_json_roundtrip() {
        let sst = SharedStringTable {
            strings: vec!["hello world".to_string()],
            rich_runs: vec![Some(vec![
                RichTextRun {
                    text: "hello".to_string(),
                    bold: Some(true),
                    italic: None,
                    font_name: None,
                    font_size: None,
                    color: None,
                },
                RichTextRun {
                    text: " world".to_string(),
                    bold: None,
                    italic: None,
                    font_name: None,
                    font_size: None,
                    color: None,
                },
            ])],
        };

        let json = serde_json::to_string(&sst).unwrap();
        let sst2: SharedStringTable = serde_json::from_str(&json).unwrap();

        assert_eq!(sst2.strings, sst.strings);
        assert_eq!(sst2.rich_runs.len(), 1);
        let runs = sst2.rich_runs[0].as_ref().unwrap();
        assert_eq!(runs.len(), 2);
        assert_eq!(runs[0].text, "hello");
        assert_eq!(runs[0].bold, Some(true));
        assert_eq!(runs[1].text, " world");
    }
}
