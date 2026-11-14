//! `[Content_Types].xml`, the Open Packaging Conventions manifest that maps a
//! file extension (`Default`) or a part name (`Override`) to a content type,
//! with the `CT_*` content type constants and [`ContentTypes`] to parse and
//! write it.

use core::hint::cold_path;
use std::collections::BTreeMap;

use quick_xml::events::{BytesStart, Event};
use quick_xml::{Reader, Writer};

use crate::ooxml::write_xml_decl;
use crate::ooxml::xml::attr_str;
use crate::{ModernXlsxError, Result};

// Content type constants.
const CT_RELATIONSHIPS: &str = "application/vnd.openxmlformats-package.relationships+xml";
pub(crate) const CT_XML: &str = "application/xml";
const CT_WORKBOOK: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml";
const CT_WORKSHEET: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml";
const CT_SHARED_STRINGS: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml";
const CT_STYLES: &str = "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml";
pub(crate) const CT_COMMENTS: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml";
pub(crate) const CT_TABLE: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml";
pub(crate) const CT_CHART: &str =
    "application/vnd.openxmlformats-officedocument.drawingml.chart+xml";
pub(crate) const CT_DRAWING: &str = "application/vnd.openxmlformats-officedocument.drawing+xml";
/// A chart's `c:userShapes` part (the ECMA-376 Part 1 Chart Drawing Part),
/// named like a sheet drawing (`xl/drawings/drawing{n}.xml`) but typed apart.
pub(crate) const CT_CHART_SHAPES: &str =
    "application/vnd.openxmlformats-officedocument.drawingml.chartshapes+xml";
/// A chart's style part (`xl/charts/style{n}.xml`), as Microsoft Excel 16.0
/// declared it in a workbook saved on 2026-11-11.
pub(crate) const CT_CHART_STYLE: &str = "application/vnd.ms-office.chartstyle+xml";
/// A chart's colour-style part (`xl/charts/colors{n}.xml`), as Microsoft Excel
/// 16.0 declared it in a workbook saved on 2026-11-11.
pub(crate) const CT_CHART_COLOR_STYLE: &str = "application/vnd.ms-office.chartcolorstyle+xml";
pub(crate) const CT_EXTERNAL_LINK: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml";
pub(crate) const CT_PIVOT_TABLE: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml";
pub(crate) const CT_PIVOT_CACHE_DEF: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml";
pub(crate) const CT_PIVOT_CACHE_REC: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml";
pub(crate) const CT_CUSTOM_XML_PROPS: &str =
    "application/vnd.openxmlformats-officedocument.customXmlProperties+xml";
pub(crate) const CT_THREADED_COMMENTS: &str = "application/vnd.ms-excel.threadedcomments+xml";
pub(crate) const CT_PERSONS: &str = "application/vnd.ms-excel.person+xml";
pub(crate) const CT_SLICER: &str = "application/vnd.ms-excel.slicer+xml";
pub(crate) const CT_SLICER_CACHE: &str = "application/vnd.ms-excel.slicerCache+xml";
pub(crate) const CT_TIMELINE: &str = "application/vnd.ms-excel.timeline+xml";
pub(crate) const CT_TIMELINE_CACHE: &str = "application/vnd.ms-excel.timelineCache+xml";
/// `xl/metadata.xml` (ECMA-376 Part 1 §12.3.10 Metadata Part); the content
/// type Microsoft Excel 16.0 (build 20527) wrote for a `=SEQUENCE(3)` workbook
/// saved on 2026-11-11.
pub(crate) const CT_SHEET_METADATA: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml";
/// `xl/printerSettings/printerSettings{n}.bin`, the binary printer-settings
/// part a sheet's page setup refers to; registered by Override, never by a
/// `bin` Default.
pub(crate) const CT_PRINTER_SETTINGS: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.printerSettings";
/// Legacy VML drawings (`.vml`: comment boxes, form controls), registered as
/// the `vml` Default.
pub(crate) const CT_VML_DRAWING: &str = "application/vnd.openxmlformats-officedocument.vmlDrawing";
/// `xl/theme/theme1.xml`: the writer keeps the source workbook's theme part
/// when it has one and generates the Office theme otherwise.
pub(crate) const CT_THEME: &str = "application/vnd.openxmlformats-officedocument.theme+xml";
/// `xl/calcChain.xml`.
pub(crate) const CT_CALC_CHAIN: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml";
/// `docProps/core.xml`.
pub(crate) const CT_CORE_PROPERTIES: &str =
    "application/vnd.openxmlformats-package.core-properties+xml";
/// `docProps/app.xml`.
pub(crate) const CT_EXTENDED_PROPERTIES: &str =
    "application/vnd.openxmlformats-officedocument.extended-properties+xml";
/// `docProps/custom.xml` (ECMA-376 Part 1 §15.2.12.2 Custom File Properties
/// Part).
pub(crate) const CT_CUSTOM_PROPERTIES: &str =
    "application/vnd.openxmlformats-officedocument.custom-properties+xml";
/// `xl/charts/chartEx{n}.xml`, an Office 2016 chart (waterfall, sunburst,
/// treemap, ...), typed apart from a DrawingML chart.
pub(crate) const CT_CHART_EX: &str = "application/vnd.ms-office.chartex+xml";
/// `xl/embeddings/oleObject{n}.bin`, an embedded OLE object (the ECMA-376
/// Part 1 Embedded Object Part), registered by Override, never by a `bin`
/// Default.
pub(crate) const CT_OLE_OBJECT: &str = "application/vnd.openxmlformats-officedocument.oleObject";
/// An embedded SpreadsheetML package (`.xlsx`, the ECMA-376 Part 1 Embedded
/// Package Part), registered as the `xlsx` Default.
pub(crate) const CT_XLSX_PACKAGE: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
/// An embedded WordprocessingML package (`.docx`), registered as the `docx`
/// Default.
pub(crate) const CT_DOCX_PACKAGE: &str =
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
/// An embedded PresentationML package (`.pptx`), registered as the `pptx`
/// Default.
pub(crate) const CT_PPTX_PACKAGE: &str =
    "application/vnd.openxmlformats-officedocument.presentationml.presentation";
/// An HD Photo / JPEG XR image (`.wdp`), registered as the `wdp` Default.
pub(crate) const CT_HD_PHOTO: &str = "image/vnd.ms-photo";

const TYPES_NS: &str = "http://schemas.openxmlformats.org/package/2006/content-types";

/// Represents `[Content_Types].xml` -- the OPC package manifest.
///
/// Contains default content types keyed by file extension and
/// override content types keyed by part name.
#[derive(Debug, Clone, Default)]
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

    /// Parse `[Content_Types].xml` from raw XML bytes. Attribute values are
    /// decoded, so `to_xml` escapes them exactly once.
    ///
    /// # Errors
    ///
    /// `XmlParse` for malformed XML or an attribute value with a bad entity or
    /// character reference.
    pub fn parse(data: &[u8]) -> Result<Self> {
        let mut reader = Reader::from_reader(data);
        let mut buf = Vec::with_capacity(512);
        let mut ct = Self::new();

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Empty(ref e) | Event::Start(ref e)) => match e.local_name().as_ref() {
                    "Default" => {
                        let (mut ext, mut ctype) = (String::new(), String::new());
                        for attr in e.attributes() {
                            let attr = attr.map_err(|err| {
                                cold_path();
                                ModernXlsxError::XmlParse(format!(
                                    "error parsing [Content_Types].xml: <Default> attributes: {err}"
                                ))
                            })?;
                            match attr.key.local_name().as_ref() {
                                "Extension" => {
                                    ext = attr_str(&attr)?.into_owned();
                                }
                                "ContentType" => {
                                    ctype = attr_str(&attr)?.into_owned();
                                }
                                _ => {}
                            }
                        }
                        if !ext.is_empty() && !ctype.is_empty() {
                            ct.defaults.insert(ext, ctype);
                        }
                    }
                    "Override" => {
                        let (mut part, mut ctype) = (String::new(), String::new());
                        for attr in e.attributes() {
                            let attr = attr.map_err(|err| {
                                cold_path();
                                ModernXlsxError::XmlParse(format!(
                                    "error parsing [Content_Types].xml: <Override> attributes: {err}"
                                ))
                            })?;
                            match attr.key.local_name().as_ref() {
                                "PartName" => {
                                    part = attr_str(&attr)?.into_owned();
                                }
                                "ContentType" => {
                                    ctype = attr_str(&attr)?.into_owned();
                                }
                                _ => {}
                            }
                        }
                        if !part.is_empty() && !ctype.is_empty() {
                            ct.overrides.insert(part, ctype);
                        }
                    }
                    _ => {}
                },
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

        write_xml_decl(&mut writer)?;

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

    #[test]
    fn test_parse_mismatched_end_tag_is_xml_parse() {
        assert!(matches!(
            ContentTypes::parse(b"<Types><a></b></Types>"),
            Err(ModernXlsxError::XmlParse(_))
        ));
    }

    #[test]
    fn parse_decodes_escaped_part_name_and_to_xml_escapes_it_once() {
        const CT_CTRL_PROPS: &str = "application/vnd.ms-excel.controlproperties+xml";
        let xml = format!(
            r#"<Types><Override PartName="/xl/ctrlProps/a&amp;b.xml" ContentType="{CT_CTRL_PROPS}"/></Types>"#
        );
        let ct = ContentTypes::parse(xml.as_bytes()).unwrap();
        assert_eq!(ct.overrides.len(), 1);
        assert_eq!(ct.overrides["/xl/ctrlProps/a&b.xml"], CT_CTRL_PROPS);

        let out = String::from_utf8(ct.to_xml().unwrap()).unwrap();
        assert_eq!(
            out.matches(r#"PartName="/xl/ctrlProps/a&amp;b.xml""#)
                .count(),
            1
        );
        assert!(!out.contains("&amp;amp;"), "{out}");
    }

    #[test]
    fn duplicate_attribute_is_xml_parse_error() {
        for xml in [
            r#"<Types><Override PartName="/a" PartName="/b" ContentType="t"/></Types>"#,
            r#"<Types><Default Extension="a" Extension="b" ContentType="t"/></Types>"#,
        ] {
            let err = ContentTypes::parse(xml.as_bytes()).unwrap_err();
            assert_eq!(err.code(), "XML_PARSE", "{xml}: {err}");
        }
    }
}
