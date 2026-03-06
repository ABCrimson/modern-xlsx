use core::hint::cold_path;
use std::collections::BTreeMap;

use quick_xml::events::{BytesDecl, BytesStart, Event};
use quick_xml::{Reader, Writer};

use crate::{ModernXlsxError, Result};

// Content type constants.
const CT_RELATIONSHIPS: &str =
    "application/vnd.openxmlformats-package.relationships+xml";
pub(crate) const CT_XML: &str = "application/xml";
const CT_WORKBOOK: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml";
const CT_WORKSHEET: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml";
const CT_SHARED_STRINGS: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml";
const CT_STYLES: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml";
pub(crate) const CT_COMMENTS: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml";
pub const CT_TABLE: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml";
pub(crate) const CT_CHART: &str =
    "application/vnd.openxmlformats-officedocument.drawingml.chart+xml";
pub(crate) const CT_DRAWING: &str =
    "application/vnd.openxmlformats-officedocument.drawing+xml";
pub(crate) const CT_EXTERNAL_LINK: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml";
pub(crate) const CT_CUSTOM_XML_PROPS: &str =
    "application/vnd.openxmlformats-officedocument.customXmlProperties+xml";

const TYPES_NS: &str = "http://schemas.openxmlformats.org/package/2006/content-types";

/// Represents `[Content_Types].xml` -- the OPC package manifest.
///
/// Contains default content types keyed by file extension and
/// override content types keyed by part name.
#[derive(Debug, Clone)]
pub struct ContentTypes {
    /// Extension -> content type (e.g. "rels" -> relationships type).
    pub defaults: BTreeMap<String, String>,
    /// Part name -> content type (e.g. "/xl/workbook.xml" -> workbook type).
    pub overrides: BTreeMap<String, String>,
}

impl ContentTypes {
    /// Create an empty `ContentTypes`.
    pub fn new() -> Self {
        Self {
            defaults: BTreeMap::new(),
            overrides: BTreeMap::new(),
        }
    }

    /// Add a default content type mapping for a file extension.
    pub fn add_default(&mut self, extension: impl Into<String>, content_type: impl Into<String>) {
        self.defaults.insert(extension.into(), content_type.into());
    }

    /// Add an override content type mapping for a specific part name.
    pub fn add_override(&mut self, part_name: impl Into<String>, content_type: impl Into<String>) {
        self.overrides.insert(part_name.into(), content_type.into());
    }

    /// Create content types for a basic workbook with the given number of sheets.
    ///
    /// Includes default mappings for `rels` and `xml` extensions, plus overrides
    /// for the workbook, shared strings, styles, and each worksheet.
    pub fn for_basic_workbook(sheet_count: usize) -> Self {
        let mut ct = Self::new();
        ct.add_default("rels", CT_RELATIONSHIPS);
        ct.add_default("xml", CT_XML);

        ct.add_override("/xl/workbook.xml", CT_WORKBOOK);
        ct.add_override("/xl/sharedStrings.xml", CT_SHARED_STRINGS);
        ct.add_override("/xl/styles.xml", CT_STYLES);

        for i in 1..=sheet_count {
            ct.add_override(format!("/xl/worksheets/sheet{i}.xml"), CT_WORKSHEET);
        }

        ct
    }

    /// Parse `[Content_Types].xml` from raw XML bytes.
    pub fn parse(data: &[u8]) -> Result<Self> {
        let mut reader = Reader::from_reader(data);
        let mut buf = Vec::with_capacity(512);
        let mut ct = Self::new();

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Empty(ref e)) | Ok(Event::Start(ref e)) => {
                    match e.local_name().as_ref() {
                        b"Default" => {
                            let (mut ext, mut ctype) = (String::new(), String::new());
                            for attr in e.attributes().flatten() {
                                match attr.key.local_name().as_ref() {
                                    b"Extension" => {
                                        ext = std::str::from_utf8(&attr.value).unwrap_or_default().to_owned();
                                    }
                                    b"ContentType" => {
                                        ctype =
                                            std::str::from_utf8(&attr.value).unwrap_or_default().to_owned();
                                    }
                                    _ => {}
                                }
                            }
                            if !ext.is_empty() && !ctype.is_empty() {
                                ct.defaults.insert(ext, ctype);
                            }
                        }
                        b"Override" => {
                            let (mut part, mut ctype) = (String::new(), String::new());
                            for attr in e.attributes().flatten() {
                                match attr.key.local_name().as_ref() {
                                    b"PartName" => {
                                        part =
                                            std::str::from_utf8(&attr.value).unwrap_or_default().to_owned();
                                    }
                                    b"ContentType" => {
                                        ctype =
                                            std::str::from_utf8(&attr.value).unwrap_or_default().to_owned();
                                    }
                                    _ => {}
                                }
                            }
                            if !part.is_empty() && !ctype.is_empty() {
                                ct.overrides.insert(part, ctype);
                            }
                        }
                        _ => {}
                    }
                }
                Ok(Event::Eof) => break,
                Ok(_) => {}
                Err(e) => {
                    cold_path();
                    return Err(ModernXlsxError::XmlParse(format!(
                        "error parsing [Content_Types].xml: {e}"
                    )));
                }
            }
            buf.clear();
        }

        Ok(ct)
    }

    /// Serialize to XML bytes.
    pub fn to_xml(&self) -> Result<Vec<u8>> {
        let mut buf: Vec<u8> = Vec::with_capacity(1024);
        let mut writer = Writer::new(&mut buf);

        // XML declaration.
        writer
            .write_event(Event::Decl(BytesDecl::new("1.0", Some("UTF-8"), Some("yes"))))
            .map_err(|e| ModernXlsxError::XmlWrite(format!("writing XML decl: {e}")))?;

        // <Types xmlns="...">
        let mut types_start = BytesStart::new("Types");
        types_start.push_attribute(("xmlns", TYPES_NS));
        writer
            .write_event(Event::Start(types_start))
            .map_err(|e| ModernXlsxError::XmlWrite(format!("writing Types start: {e}")))?;

        // <Default> elements (BTreeMap iterates in sorted order).
        for (ext, ctype) in &self.defaults {
            let mut elem = BytesStart::new("Default");
            elem.push_attribute(("Extension", ext.as_str()));
            elem.push_attribute(("ContentType", ctype.as_str()));
            writer
                .write_event(Event::Empty(elem))
                .map_err(|e| ModernXlsxError::XmlWrite(format!("writing Default: {e}")))?;
        }

        // <Override> elements (BTreeMap iterates in sorted order).
        for (part, ctype) in &self.overrides {
            let mut elem = BytesStart::new("Override");
            elem.push_attribute(("PartName", part.as_str()));
            elem.push_attribute(("ContentType", ctype.as_str()));
            writer
                .write_event(Event::Empty(elem))
                .map_err(|e| ModernXlsxError::XmlWrite(format!("writing Override: {e}")))?;
        }

        // </Types>
        writer
            .write_event(Event::End(quick_xml::events::BytesEnd::new("Types")))
            .map_err(|e| ModernXlsxError::XmlWrite(format!("writing Types end: {e}")))?;

        Ok(buf)
    }
}

impl Default for ContentTypes {
    fn default() -> Self {
        Self::new()
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use pretty_assertions::assert_eq;

    #[test]
    fn test_parse_content_types() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>"#;

        let ct = ContentTypes::parse(xml.as_bytes()).unwrap();

        assert_eq!(ct.defaults.len(), 2);
        assert_eq!(ct.defaults["rels"], CT_RELATIONSHIPS);
        assert_eq!(ct.defaults["xml"], CT_XML);

        assert_eq!(ct.overrides.len(), 2);
        assert_eq!(ct.overrides["/xl/workbook.xml"], CT_WORKBOOK);
        assert_eq!(ct.overrides["/xl/worksheets/sheet1.xml"], CT_WORKSHEET);
    }

    #[test]
    fn test_write_content_types() {
        let mut ct = ContentTypes::new();
        ct.add_default("rels", CT_RELATIONSHIPS);
        ct.add_default("xml", CT_XML);
        ct.add_override("/xl/workbook.xml", CT_WORKBOOK);
        ct.add_override("/xl/worksheets/sheet1.xml", CT_WORKSHEET);

        let xml = ct.to_xml().unwrap();

        // Re-parse and verify roundtrip.
        let ct2 = ContentTypes::parse(&xml).unwrap();

        assert_eq!(ct2.defaults.len(), 2);
        assert_eq!(ct2.defaults["rels"], CT_RELATIONSHIPS);
        assert_eq!(ct2.defaults["xml"], CT_XML);

        assert_eq!(ct2.overrides.len(), 2);
        assert_eq!(ct2.overrides["/xl/workbook.xml"], CT_WORKBOOK);
        assert_eq!(ct2.overrides["/xl/worksheets/sheet1.xml"], CT_WORKSHEET);
    }

    #[test]
    fn test_content_types_for_basic_workbook() {
        let ct = ContentTypes::for_basic_workbook(1);

        // 2 defaults: rels, xml
        assert_eq!(ct.defaults.len(), 2);
        assert_eq!(ct.defaults["rels"], CT_RELATIONSHIPS);
        assert_eq!(ct.defaults["xml"], CT_XML);

        // 4 overrides: workbook, sharedStrings, styles, sheet1
        assert_eq!(ct.overrides.len(), 4);
        assert_eq!(ct.overrides["/xl/workbook.xml"], CT_WORKBOOK);
        assert_eq!(ct.overrides["/xl/sharedStrings.xml"], CT_SHARED_STRINGS);
        assert_eq!(ct.overrides["/xl/styles.xml"], CT_STYLES);
        assert_eq!(ct.overrides["/xl/worksheets/sheet1.xml"], CT_WORKSHEET);
    }
}
