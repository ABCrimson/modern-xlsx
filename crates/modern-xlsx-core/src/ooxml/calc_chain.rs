//! CalcChain parser and writer.
//!
//! Handles `xl/calcChain.xml` which records the calculation order for formula
//! cells. Each entry maps a cell reference to a sheet.

use core::hint::cold_path;

use quick_xml::events::{BytesEnd, BytesStart, Event};
use quick_xml::{Reader, Writer};

use serde::{Deserialize, Serialize};

use crate::ooxml::xml::{SPREADSHEET_NS, attr_str, parse_attr, xml_attr, xsd_bool};
use crate::ooxml::{write_xml_decl, xml_write_err};
use crate::{ModernXlsxError, Result};

/// A single entry in the calculation chain (`CT_CalcCell`, ECMA-376 §18.6.1).
#[derive(Debug, Clone, Default, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct CalcChainEntry {
    /// The cell reference (e.g. `"A1"`, `"B5"`).
    pub cell_ref: String,
    /// The sheet this cell belongs to: its position among the workbook's
    /// sheets, 1-based. [`parse`] returns the `i` attribute as written (the
    /// source `sheetId`); the reader maps it to the position, because the
    /// writer numbers `sheetId` 1..n by position.
    pub sheet_id: u32,
    /// `s`: the cell is on a child chain.
    #[serde(default, skip_serializing_if = "crate::ooxml::is_false")]
    pub s: bool,
    /// `l`: a new dependency level starts at this cell.
    #[serde(default, skip_serializing_if = "crate::ooxml::is_false")]
    pub l: bool,
    /// `t`: the cell holds a new thread.
    #[serde(default, skip_serializing_if = "crate::ooxml::is_false")]
    pub t: bool,
    /// `a`: the cell is part of an array formula.
    #[serde(default, skip_serializing_if = "crate::ooxml::is_false")]
    pub a: bool,
}

/// Parse `xl/calcChain.xml` from raw XML bytes.
///
/// The expected structure is:
/// ```xml
/// <calcChain xmlns="...">
///   <c r="A1" i="1"/>
///   <c r="B2"/>
///   <c r="A1" i="2" l="1"/>
/// </calcChain>
/// ```
/// An entry without `i` is on the same sheet as the entry before it
/// (ECMA-376 §18.6.1); the first entry's default is `0`. `sheet_id` holds
/// the `i` value as written.
///
/// # Errors
///
/// `XmlParse` for malformed XML or a bad attribute value.
pub fn parse(data: &[u8]) -> Result<Vec<CalcChainEntry>> {
    let mut reader = Reader::from_reader(data);
    reader.config_mut().trim_text(true);

    let mut buf = Vec::with_capacity(256);
    let mut entries = Vec::new();
    // `i` of the previous entry, inherited when the attribute is omitted.
    let mut last_i: u32 = 0;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Empty(ref e) | Event::Start(ref e)) if e.local_name().as_ref() == "c" => {
                let mut entry = CalcChainEntry {
                    sheet_id: last_i,
                    ..Default::default()
                };

                for attr in e.attributes() {
                    let attr = attr.map_err(|err| {
                        cold_path();
                        ModernXlsxError::XmlParse(format!("<c> attributes: {err}"))
                    })?;
                    match attr.key.local_name().as_ref() {
                        "r" => entry.cell_ref = attr_str(&attr)?.into_owned(),
                        "i" => entry.sheet_id = parse_attr("c", &attr)?,
                        "s" => entry.s = xsd_bool(&attr_str(&attr)?).unwrap_or(false),
                        "l" => entry.l = xsd_bool(&attr_str(&attr)?).unwrap_or(false),
                        "t" => entry.t = xsd_bool(&attr_str(&attr)?).unwrap_or(false),
                        "a" => entry.a = xsd_bool(&attr_str(&attr)?).unwrap_or(false),
                        _ => {}
                    }
                }

                last_i = entry.sheet_id;
                if !entry.cell_ref.is_empty() {
                    entries.push(entry);
                }
            }
            Ok(Event::Eof) => break,
            Err(err) => {
                cold_path();
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
///
/// `i` is written on every entry; `s`, `l`, `t` and `a` only when set.
///
/// # Errors
///
/// `XmlWrite` when serialization fails.
pub fn to_xml(entries: &[CalcChainEntry]) -> Result<Vec<u8>> {
    let mut buf: Vec<u8> = Vec::with_capacity(256 + entries.len() * 32);
    let mut writer = Writer::new(&mut buf);
    let mut ibuf = itoa::Buffer::new();

    // XML declaration.
    write_xml_decl(&mut writer)?;

    // <calcChain xmlns="...">
    let mut root = BytesStart::new("calcChain");
    root.push_attribute(("xmlns", SPREADSHEET_NS));
    writer
        .write_event(Event::Start(root))
        .map_err(xml_write_err)?;

    for entry in entries {
        let mut elem = BytesStart::new("c");
        elem.push_attribute(xml_attr("r", &entry.cell_ref));
        elem.push_attribute(("i", ibuf.format(entry.sheet_id)));
        for (key, set) in [
            ("s", entry.s),
            ("l", entry.l),
            ("t", entry.t),
            ("a", entry.a),
        ] {
            if set {
                elem.push_attribute((key, "1"));
            }
        }
        writer
            .write_event(Event::Empty(elem))
            .map_err(xml_write_err)?;
    }

    // </calcChain>
    writer
        .write_event(Event::End(BytesEnd::new("calcChain")))
        .map_err(xml_write_err)?;

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
                ..Default::default()
            },
            CalcChainEntry {
                cell_ref: "D5".to_string(),
                sheet_id: 2,
                ..Default::default()
            },
            CalcChainEntry {
                cell_ref: "Z100".to_string(),
                sheet_id: 3,
                ..Default::default()
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

    #[test]
    fn missing_i_inherits_previous_and_flags_roundtrip() {
        let xml = r#"<calcChain xmlns="x"><c r="A1" i="2"/><c r="A2"/><c r="A3" i="1" l="1"/><c r="A4" s="true" t="1" a="1"/></calcChain>"#;
        let entries = parse(xml.as_bytes()).unwrap();
        let ids: Vec<u32> = entries.iter().map(|e| e.sheet_id).collect();
        assert_eq!(ids, [2, 2, 1, 1]);
        assert!(entries[2].l && !entries[2].s);
        assert!(entries[3].s && entries[3].t && entries[3].a && !entries[3].l);

        let out = String::from_utf8(to_xml(&entries).unwrap()).unwrap();
        assert!(out.contains(r#"<c r="A3" i="1" l="1"/>"#), "{out}");
        assert!(out.contains(r#"<c r="A2" i="2"/>"#), "{out}");
        assert_eq!(parse(out.as_bytes()).unwrap(), entries);
    }

    #[test]
    fn mismatched_end_tag_is_xml_parse_error() {
        assert!(matches!(
            parse(br#"<root xmlns="x"><a></b></root>"#),
            Err(ModernXlsxError::XmlParse(_))
        ));
    }

    #[test]
    fn malformed_i_is_error() {
        assert!(matches!(
            parse(br#"<calcChain xmlns="x"><c r="A1" i="x"/></calcChain>"#),
            Err(ModernXlsxError::XmlParse(_))
        ));
    }
}
