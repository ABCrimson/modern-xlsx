//! CalcChain parser and writer.
//!
//! Handles `xl/calcChain.xml` which records the calculation order for formula
//! cells. Each entry maps a cell reference to a sheet ID.

use quick_xml::events::{BytesDecl, BytesEnd, BytesStart, Event};
use quick_xml::{Reader, Writer};

use serde::{Deserialize, Serialize};

use crate::{ModernXlsxError, Result};

use super::SPREADSHEET_NS;

/// A single entry in the calculation chain.
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct CalcChainEntry {
    /// The cell reference (e.g. `"A1"`, `"B5"`).
    pub cell_ref: String,
    /// The sheet ID this cell belongs to.
    pub sheet_id: u32,
}

/// Parse `xl/calcChain.xml` from raw XML bytes.
///
/// The expected structure is:
/// ```xml
/// <calcChain xmlns="...">
///   <c r="A1" i="1"/>
///   <c r="B2" i="1"/>
///   <c r="A1" i="2"/>
/// </calcChain>
/// ```
pub fn parse(data: &[u8]) -> Result<Vec<CalcChainEntry>> {
    let mut reader = Reader::from_reader(data);
    reader.config_mut().trim_text(true);

    let mut buf = Vec::new();
    let mut entries = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Empty(ref e)) | Ok(Event::Start(ref e)) => {
                if e.local_name().as_ref() == b"c" {
                    let mut cell_ref = String::new();
                    let mut sheet_id: u32 = 0;

                    for attr in e.attributes().flatten() {
                        match attr.key.local_name().as_ref() {
                            b"r" => {
                                cell_ref = std::str::from_utf8(&attr.value)
                                    .unwrap_or_default()
                                    .to_owned();
                            }
                            b"i" => {
                                let val =
                                    std::str::from_utf8(&attr.value).unwrap_or_default();
                                sheet_id = val.parse::<u32>().unwrap_or(0);
                            }
                            _ => {}
                        }
                    }

                    if !cell_ref.is_empty() {
                        entries.push(CalcChainEntry { cell_ref, sheet_id });
                    }
                }
            }
            Ok(Event::Eof) => break,
            Err(err) => {
                return Err(ModernXlsxError::XmlParse(format!(
                    "error parsing calcChain.xml: {err}"
                )));
            }
            _ => {}
        }
        buf.clear();
    }

    Ok(entries)
}

/// Serialize a slice of [`CalcChainEntry`] to `xl/calcChain.xml` bytes.
pub fn to_xml(entries: &[CalcChainEntry]) -> Result<Vec<u8>> {
    let mut buf: Vec<u8> = Vec::with_capacity(256 + entries.len() * 32);
    let mut writer = Writer::new(&mut buf);
    let mut ibuf = itoa::Buffer::new();

    let map_err = |e: std::io::Error| ModernXlsxError::XmlWrite(e.to_string());

    // XML declaration.
    writer
        .write_event(Event::Decl(BytesDecl::new("1.0", Some("UTF-8"), Some("yes"))))
        .map_err(map_err)?;

    // <calcChain xmlns="...">
    let mut root = BytesStart::new("calcChain");
    root.push_attribute(("xmlns", SPREADSHEET_NS));
    writer.write_event(Event::Start(root)).map_err(map_err)?;

    for entry in entries {
        let mut elem = BytesStart::new("c");
        elem.push_attribute(("r", entry.cell_ref.as_str()));
        elem.push_attribute(("i", ibuf.format(entry.sheet_id)));
        writer.write_event(Event::Empty(elem)).map_err(map_err)?;
    }

    // </calcChain>
    writer
        .write_event(Event::End(BytesEnd::new("calcChain")))
        .map_err(map_err)?;

    Ok(buf)
}

#[cfg(test)]
mod tests {
    use super::*;
    use pretty_assertions::assert_eq;

    #[test]
    fn test_parse_calc_chain() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<calcChain xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <c r="A1" i="1"/>
  <c r="B2" i="1"/>
  <c r="C3" i="2"/>
</calcChain>"#;

        let entries = parse(xml.as_bytes()).unwrap();
        assert_eq!(entries.len(), 3);

        assert_eq!(entries[0].cell_ref, "A1");
        assert_eq!(entries[0].sheet_id, 1);

        assert_eq!(entries[1].cell_ref, "B2");
        assert_eq!(entries[1].sheet_id, 1);

        assert_eq!(entries[2].cell_ref, "C3");
        assert_eq!(entries[2].sheet_id, 2);
    }

    #[test]
    fn test_parse_empty_calc_chain() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<calcChain xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
</calcChain>"#;

        let entries = parse(xml.as_bytes()).unwrap();
        assert!(entries.is_empty());
    }

    #[test]
    fn test_calc_chain_roundtrip() {
        let entries = vec![
            CalcChainEntry {
                cell_ref: "A1".to_string(),
                sheet_id: 1,
            },
            CalcChainEntry {
                cell_ref: "D5".to_string(),
                sheet_id: 2,
            },
            CalcChainEntry {
                cell_ref: "Z100".to_string(),
                sheet_id: 3,
            },
        ];

        let xml = to_xml(&entries).unwrap();
        let parsed = parse(&xml).unwrap();

        assert_eq!(parsed.len(), 3);
        assert_eq!(parsed[0].cell_ref, "A1");
        assert_eq!(parsed[0].sheet_id, 1);
        assert_eq!(parsed[1].cell_ref, "D5");
        assert_eq!(parsed[1].sheet_id, 2);
        assert_eq!(parsed[2].cell_ref, "Z100");
        assert_eq!(parsed[2].sheet_id, 3);
    }

    #[test]
    fn test_write_empty_calc_chain() {
        let xml = to_xml(&[]).unwrap();
        let parsed = parse(&xml).unwrap();
        assert!(parsed.is_empty());
    }
}
