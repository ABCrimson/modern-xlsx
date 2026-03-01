use std::collections::HashMap;

use quick_xml::events::{BytesDecl, BytesEnd, BytesStart, BytesText, Event};
use quick_xml::{Reader, Writer};

use crate::{IronsheetError, Result};

/// A read-only Shared String Table parsed from `sharedStrings.xml`.
///
/// In OOXML, cells that contain string values store an index into this table
/// rather than the string itself.
#[derive(Debug, Clone)]
pub struct SharedStringTable {
    strings: Vec<String>,
}

impl SharedStringTable {
    /// Parse a Shared String Table from the raw XML bytes of `sharedStrings.xml`.
    pub fn parse(data: &[u8]) -> Result<Self> {
        let mut reader = Reader::from_reader(data);
        reader.config_mut().trim_text(false);

        let mut buf = Vec::new();
        let mut strings = Vec::new();

        let mut in_si = false;
        let mut in_t = false;
        let mut current_text = String::new();

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(e)) => {
                    let local = e.local_name();
                    match local.as_ref() {
                        b"si" => {
                            in_si = true;
                            current_text.clear();
                        }
                        b"t" if in_si => {
                            in_t = true;
                        }
                        _ => {}
                    }
                }
                Ok(Event::Text(e)) if in_t => {
                    current_text.push_str(
                        &String::from_utf8_lossy(e.as_ref()),
                    );
                }
                Ok(Event::End(e)) => {
                    let local = e.local_name();
                    match local.as_ref() {
                        b"t" => {
                            in_t = false;
                        }
                        b"si" => {
                            strings.push(std::mem::take(&mut current_text));
                            in_si = false;
                        }
                        _ => {}
                    }
                }
                Ok(Event::Eof) => break,
                Err(err) => return Err(IronsheetError::XmlParse(err.to_string())),
                _ => {}
            }
            buf.clear();
        }

        Ok(SharedStringTable { strings })
    }

    /// Return an empty Shared String Table.
    ///
    /// Used when the workbook does not contain a `sharedStrings.xml` part.
    pub fn empty() -> Self {
        SharedStringTable {
            strings: Vec::new(),
        }
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
        self.strings.push(s.to_owned());
        self.index.insert(s.to_owned(), idx);
        idx
    }

    /// Serialize the shared string table to an XML string suitable for
    /// inclusion in an XLSX archive as `xl/sharedStrings.xml`.
    pub fn to_xml(&self) -> Result<String> {
        let mut buf: Vec<u8> = Vec::new();
        let mut writer = Writer::new(&mut buf);

        let map_write_err =
            |e: std::io::Error| IronsheetError::XmlWrite(e.to_string());

        // XML declaration.
        writer
            .write_event(Event::Decl(BytesDecl::new(
                "1.0",
                Some("UTF-8"),
                Some("yes"),
            )))
            .map_err(map_write_err)?;

        // <sst xmlns="..." count="N" uniqueCount="N">
        let count_str = self.strings.len().to_string();
        let mut sst_start = BytesStart::new("sst");
        sst_start.push_attribute((
            "xmlns",
            "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
        ));
        sst_start.push_attribute(("count", count_str.as_str()));
        sst_start.push_attribute(("uniqueCount", count_str.as_str()));
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

        String::from_utf8(buf)
            .map_err(|e| IronsheetError::XmlWrite(format!("invalid UTF-8 in output: {e}")))
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
        let sst = SharedStringTable::parse(xml.as_bytes()).unwrap();

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
        let sst = SharedStringTable::parse(xml.as_bytes()).unwrap();

        assert!(sst.is_empty());
        assert_eq!(sst.len(), 0);
    }
}
