//! Document properties parser and writer.
//!
//! Handles `docProps/core.xml` (Dublin Core metadata) and `docProps/app.xml`
//! (extended application properties).

use core::hint::cold_path;

use quick_xml::events::{BytesEnd, BytesStart, BytesText, Event};
use quick_xml::{Reader, Writer};

use serde::{Deserialize, Serialize};

use crate::ooxml::xml::{push_general_ref, sanitize_xml_text};
use crate::{ModernXlsxError, Result};

use crate::ooxml::{write_xml_decl, xml_write_err};

// Namespace URIs used in core.xml.
const CP_NS: &str = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties";
const DC_NS: &str = "http://purl.org/dc/elements/1.1/";
const DCTERMS_NS: &str = "http://purl.org/dc/terms/";
const XSI_NS: &str = "http://www.w3.org/2001/XMLSchema-instance";

// Namespace URI for app.xml.
const APP_NS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties";
const VT_NS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes";

/// Combined document properties from `docProps/core.xml` and `docProps/app.xml`.
///
/// Only the properties listed here are read and written; every value is kept
/// as the raw element text, including surrounding white space. Other app
/// properties are not round-tripped. In particular `TitlesOfParts` and
/// `HeadingPairs` are dropped on purpose: they list sheet names and would go
/// stale as soon as a sheet is renamed, added or removed.
#[derive(Debug, Clone, Default, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct DocProperties {
    // Core properties (docProps/core.xml)
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub title: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub subject: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub creator: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub keywords: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub description: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub last_modified_by: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub created: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub modified: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub category: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub content_status: Option<String>,
    /// `cp:lastPrinted`, an ISO 8601 date-time.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub last_printed: Option<String>,
    /// `dc:language`, for example `en-US`.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub language: Option<String>,
    /// `dc:identifier`.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub identifier: Option<String>,
    /// `cp:version`.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub version: Option<String>,
    /// `cp:revision`.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub revision: Option<String>,
    // App properties (docProps/app.xml)
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub application: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub company: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub manager: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub app_version: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub hyperlink_base: Option<String>,
    /// `TotalTime`: minutes spent editing.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub total_time: Option<String>,
    /// `DocSecurity`: the document security level, a bit field.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub doc_security: Option<String>,
}

/// Collect the text of the property elements named in `wanted`, calling
/// `store(name, value)` at the element's end tag.
///
/// The text is kept raw (`trim_text(false)`): entities, character references
/// and CDATA are resolved and white space is preserved, so a value of only
/// white space is `Some`; only XML 1.0 line-end normalization applies (a
/// literal CRLF or CR reads as LF, `&#13;` stays a CR). An element with no
/// text gives `None`.
fn parse_properties(
    data: &[u8],
    part: &str,
    wanted: &[&'static str],
    mut store: impl FnMut(&'static str, Option<String>),
) -> Result<()> {
    let mut reader = Reader::from_reader(data);
    reader.config_mut().trim_text(false);

    let mut buf = Vec::with_capacity(512);

    // The property element being read, and its text so far.
    let mut current_element: Option<&'static str> = None;
    let mut text_buf = String::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) => {
                let local = e.local_name();
                if let Some(name) = wanted.iter().copied().find(|w| *w == local.as_ref()) {
                    current_element = Some(name);
                    text_buf.clear();
                }
            }
            Ok(Event::Text(ref e)) if current_element.is_some() => {
                text_buf.push_str(&e.xml10_content());
            }
            Ok(Event::CData(ref e)) if current_element.is_some() => {
                text_buf.push_str(&e.xml10_content());
            }
            Ok(Event::GeneralRef(ref e)) if current_element.is_some() => {
                push_general_ref(&mut text_buf, e)?;
            }
            Ok(Event::End(ref e)) => {
                if let Some(elem) = current_element
                    && e.local_name().as_ref() == elem
                {
                    let val = if text_buf.is_empty() {
                        None
                    } else {
                        Some(std::mem::take(&mut text_buf))
                    };
                    store(elem, val);
                    current_element = None;
                }
            }
            Ok(Event::Eof) => break,
            Err(err) => {
                cold_path();
                return Err(ModernXlsxError::XmlParse(format!(
                    "error parsing {part}: {err}"
                )));
            }
            _ => {}
        }
        buf.clear();
    }

    Ok(())
}

/// Parse `docProps/core.xml` into a [`DocProperties`].
///
/// The XML uses Dublin Core namespaces (dc:, dcterms:, cp:). We match on
/// local_name() to ignore namespace prefixes.
///
/// # Errors
///
/// `XmlParse` for malformed XML or a bad entity.
pub fn parse_core(data: &[u8]) -> Result<DocProperties> {
    let mut props = DocProperties::default();
    parse_properties(
        data,
        "docProps/core.xml",
        &[
            "title",
            "creator",
            "subject",
            "keywords",
            "description",
            "lastModifiedBy",
            "created",
            "modified",
            "category",
            "contentStatus",
            "revision",
            "lastPrinted",
            "language",
            "identifier",
            "version",
        ],
        |name, val| match name {
            "title" => props.title = val,
            "subject" => props.subject = val,
            "creator" => props.creator = val,
            "keywords" => props.keywords = val,
            "description" => props.description = val,
            "lastModifiedBy" => props.last_modified_by = val,
            "created" => props.created = val,
            "modified" => props.modified = val,
            "category" => props.category = val,
            "contentStatus" => props.content_status = val,
            "revision" => props.revision = val,
            "lastPrinted" => props.last_printed = val,
            "language" => props.language = val,
            "identifier" => props.identifier = val,
            "version" => props.version = val,
            _ => {}
        },
    )?;
    Ok(props)
}

/// Parse `docProps/app.xml` and fill in the app fields (application, company,
/// manager, ...) on an existing [`DocProperties`].
///
/// # Errors
///
/// `XmlParse` for malformed XML or a bad entity.
pub fn parse_app(props: &mut DocProperties, data: &[u8]) -> Result<()> {
    parse_properties(
        data,
        "docProps/app.xml",
        &[
            "Application",
            "Company",
            "Manager",
            "AppVersion",
            "HyperlinkBase",
            "TotalTime",
            "DocSecurity",
        ],
        |name, val| match name {
            "Application" => props.application = val,
            "Company" => props.company = val,
            "Manager" => props.manager = val,
            "AppVersion" => props.app_version = val,
            "HyperlinkBase" => props.hyperlink_base = val,
            "TotalTime" => props.total_time = val,
            "DocSecurity" => props.doc_security = val,
            _ => {}
        },
    )
}

/// Write `<tag>value</tag>` for a set property, the value with XML-illegal
/// characters removed; nothing for `None`. With `w3cdtf` the start tag carries
/// `xsi:type="dcterms:W3CDTF"` (`dcterms:created`, `dcterms:modified`).
fn write_prop<W: std::io::Write>(
    w: &mut Writer<W>,
    tag: &str,
    w3cdtf: bool,
    value: Option<&str>,
) -> Result<()> {
    let Some(v) = value else {
        return Ok(());
    };
    let mut el = w.create_element(tag);
    if w3cdtf {
        el = el.with_attribute(("xsi:type", "dcterms:W3CDTF"));
    }
    el.write_text_content(BytesText::new(&sanitize_xml_text(v)))
        .map_err(xml_write_err)?;
    Ok(())
}

impl DocProperties {
    /// Serialize to `docProps/core.xml` bytes.
    pub fn to_core_xml(&self) -> Result<Vec<u8>> {
        let mut buf: Vec<u8> = Vec::with_capacity(1024);
        let mut writer = Writer::new(&mut buf);

        // XML declaration.
        write_xml_decl(&mut writer)?;

        // <cp:coreProperties>
        let mut root = BytesStart::new("cp:coreProperties");
        root.push_attribute(("xmlns:cp", CP_NS));
        root.push_attribute(("xmlns:dc", DC_NS));
        root.push_attribute(("xmlns:dcterms", DCTERMS_NS));
        root.push_attribute(("xmlns:xsi", XSI_NS));
        writer
            .write_event(Event::Start(root))
            .map_err(xml_write_err)?;

        write_prop(&mut writer, "dc:title", false, self.title.as_deref())?;
        write_prop(&mut writer, "dc:subject", false, self.subject.as_deref())?;
        write_prop(&mut writer, "dc:creator", false, self.creator.as_deref())?;
        write_prop(&mut writer, "cp:keywords", false, self.keywords.as_deref())?;
        write_prop(
            &mut writer,
            "dc:description",
            false,
            self.description.as_deref(),
        )?;
        write_prop(
            &mut writer,
            "cp:lastModifiedBy",
            false,
            self.last_modified_by.as_deref(),
        )?;
        write_prop(&mut writer, "cp:category", false, self.category.as_deref())?;
        write_prop(
            &mut writer,
            "cp:contentStatus",
            false,
            self.content_status.as_deref(),
        )?;
        write_prop(&mut writer, "cp:revision", false, self.revision.as_deref())?;
        write_prop(
            &mut writer,
            "cp:lastPrinted",
            false,
            self.last_printed.as_deref(),
        )?;
        write_prop(&mut writer, "dc:language", false, self.language.as_deref())?;
        write_prop(
            &mut writer,
            "dc:identifier",
            false,
            self.identifier.as_deref(),
        )?;
        write_prop(&mut writer, "cp:version", false, self.version.as_deref())?;
        // dcterms:created and dcterms:modified carry xsi:type="dcterms:W3CDTF".
        write_prop(
            &mut writer,
            "dcterms:created",
            true,
            self.created.as_deref(),
        )?;
        write_prop(
            &mut writer,
            "dcterms:modified",
            true,
            self.modified.as_deref(),
        )?;

        // </cp:coreProperties>
        writer
            .write_event(Event::End(BytesEnd::new("cp:coreProperties")))
            .map_err(xml_write_err)?;

        Ok(buf)
    }

    /// Serialize to `docProps/app.xml` bytes.
    pub fn to_app_xml(&self) -> Result<Vec<u8>> {
        let mut buf: Vec<u8> = Vec::with_capacity(512);
        let mut writer = Writer::new(&mut buf);

        // XML declaration.
        write_xml_decl(&mut writer)?;

        // <Properties>
        let mut root = BytesStart::new("Properties");
        root.push_attribute(("xmlns", APP_NS));
        root.push_attribute(("xmlns:vt", VT_NS));
        writer
            .write_event(Event::Start(root))
            .map_err(xml_write_err)?;

        write_prop(
            &mut writer,
            "Application",
            false,
            self.application.as_deref(),
        )?;
        write_prop(&mut writer, "Company", false, self.company.as_deref())?;
        write_prop(&mut writer, "Manager", false, self.manager.as_deref())?;
        write_prop(
            &mut writer,
            "AppVersion",
            false,
            self.app_version.as_deref(),
        )?;
        write_prop(
            &mut writer,
            "HyperlinkBase",
            false,
            self.hyperlink_base.as_deref(),
        )?;
        write_prop(&mut writer, "TotalTime", false, self.total_time.as_deref())?;
        write_prop(
            &mut writer,
            "DocSecurity",
            false,
            self.doc_security.as_deref(),
        )?;

        // </Properties>
        writer
            .write_event(Event::End(BytesEnd::new("Properties")))
            .map_err(xml_write_err)?;

        Ok(buf)
    }

    /// Returns `true` if any core property (docProps/core.xml) is set.
    pub fn has_core(&self) -> bool {
        self.title.is_some()
            || self.subject.is_some()
            || self.creator.is_some()
            || self.keywords.is_some()
            || self.description.is_some()
            || self.last_modified_by.is_some()
            || self.created.is_some()
            || self.modified.is_some()
            || self.category.is_some()
            || self.content_status.is_some()
            || self.revision.is_some()
            || self.last_printed.is_some()
            || self.language.is_some()
            || self.identifier.is_some()
            || self.version.is_some()
    }

    /// Returns `true` if any app property (docProps/app.xml) is set.
    pub fn has_app(&self) -> bool {
        self.application.is_some()
            || self.company.is_some()
            || self.manager.is_some()
            || self.app_version.is_some()
            || self.hyperlink_base.is_some()
            || self.total_time.is_some()
            || self.doc_security.is_some()
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use pretty_assertions::assert_eq;

    #[test]
    fn test_parse_core_xml() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
                   xmlns:dc="http://purl.org/dc/elements/1.1/"
                   xmlns:dcterms="http://purl.org/dc/terms/"
                   xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>Test Workbook</dc:title>
  <dc:subject>Unit Tests</dc:subject>
  <dc:creator>Test Author</dc:creator>
  <cp:keywords>test;xlsx;modern-xlsx</cp:keywords>
  <dc:description>A test workbook</dc:description>
  <cp:lastModifiedBy>Another User</cp:lastModifiedBy>
  <dcterms:created xsi:type="dcterms:W3CDTF">2024-01-15T10:30:00Z</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">2024-06-20T14:00:00Z</dcterms:modified>
  <cp:category>Reports</cp:category>
  <cp:contentStatus>Final</cp:contentStatus>
</cp:coreProperties>"#;

        let props = parse_core(xml.as_bytes()).unwrap();
        assert_eq!(props.title.as_deref(), Some("Test Workbook"));
        assert_eq!(props.subject.as_deref(), Some("Unit Tests"));
        assert_eq!(props.creator.as_deref(), Some("Test Author"));
        assert_eq!(props.keywords.as_deref(), Some("test;xlsx;modern-xlsx"));
        assert_eq!(props.description.as_deref(), Some("A test workbook"));
        assert_eq!(props.last_modified_by.as_deref(), Some("Another User"));
        assert_eq!(props.created.as_deref(), Some("2024-01-15T10:30:00Z"));
        assert_eq!(props.modified.as_deref(), Some("2024-06-20T14:00:00Z"));
        assert_eq!(props.category.as_deref(), Some("Reports"));
        assert_eq!(props.content_status.as_deref(), Some("Final"));
    }

    #[test]
    fn test_parse_core_xml_partial() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
                   xmlns:dc="http://purl.org/dc/elements/1.1/">
  <dc:creator>Minimal Author</dc:creator>
</cp:coreProperties>"#;

        let props = parse_core(xml.as_bytes()).unwrap();
        assert_eq!(props.creator.as_deref(), Some("Minimal Author"));
        assert_eq!(props.title, None);
        assert_eq!(props.subject, None);
        assert_eq!(props.created, None);
    }

    #[test]
    fn test_parse_app_xml() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
            xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Application>Microsoft Excel</Application>
  <Company>Acme Corp</Company>
  <Manager>Jane Doe</Manager>
</Properties>"#;

        let mut props = DocProperties::default();
        parse_app(&mut props, xml.as_bytes()).unwrap();
        assert_eq!(props.application.as_deref(), Some("Microsoft Excel"));
        assert_eq!(props.company.as_deref(), Some("Acme Corp"));
        assert_eq!(props.manager.as_deref(), Some("Jane Doe"));
    }

    #[test]
    fn test_core_xml_roundtrip() {
        let props = DocProperties {
            title: Some("My Title".to_string()),
            subject: Some("My Subject".to_string()),
            creator: Some("Author Name".to_string()),
            keywords: Some("key1;key2".to_string()),
            description: Some("A description".to_string()),
            last_modified_by: Some("Editor".to_string()),
            created: Some("2024-01-01T00:00:00Z".to_string()),
            modified: Some("2024-06-01T12:00:00Z".to_string()),
            category: Some("Finance".to_string()),
            content_status: Some("Draft".to_string()),
            application: None,
            company: None,
            manager: None,
            app_version: None,
            hyperlink_base: None,
            revision: None,
            ..Default::default()
        };

        let xml = props.to_core_xml().unwrap();
        let parsed = parse_core(&xml).unwrap();

        assert_eq!(parsed.title, props.title);
        assert_eq!(parsed.subject, props.subject);
        assert_eq!(parsed.creator, props.creator);
        assert_eq!(parsed.keywords, props.keywords);
        assert_eq!(parsed.description, props.description);
        assert_eq!(parsed.last_modified_by, props.last_modified_by);
        assert_eq!(parsed.created, props.created);
        assert_eq!(parsed.modified, props.modified);
        assert_eq!(parsed.category, props.category);
        assert_eq!(parsed.content_status, props.content_status);
        assert_eq!(parsed.revision, props.revision);
    }

    #[test]
    fn core_xml_sanitizes_illegal_chars_in_dates_and_text() {
        // `created`/`modified` are free-form `Option<String>` with no W3CDTF
        // validation, so an illegal char in them (or in any text property) must
        // be stripped — otherwise the serialized core.xml part is corrupt.
        let props = DocProperties {
            title: Some("Re\u{0001}port".to_string()),
            created: Some("2024-01-01T00:00:00Z\u{FFFE}".to_string()),
            modified: Some("2024-06-01\u{0008}T12:00:00Z".to_string()),
            ..Default::default()
        };

        let xml = props.to_core_xml().unwrap();
        // No XML-1.0-illegal character may survive into the serialized part.
        let text = std::str::from_utf8(&xml).unwrap();
        assert!(
            !text.chars().any(crate::ooxml::is_illegal_xml_char),
            "serialized core.xml retained an XML-1.0-illegal character"
        );

        // The document still parses, with the illegal chars removed.
        let parsed = parse_core(&xml).unwrap();
        assert_eq!(parsed.title.as_deref(), Some("Report"));
        assert_eq!(parsed.created.as_deref(), Some("2024-01-01T00:00:00Z"));
        assert_eq!(parsed.modified.as_deref(), Some("2024-06-01T12:00:00Z"));
    }

    #[test]
    fn test_app_xml_roundtrip() {
        let props = DocProperties {
            application: Some("modern-xlsx".to_string()),
            company: Some("TestCo".to_string()),
            manager: Some("Boss".to_string()),
            ..Default::default()
        };

        let xml = props.to_app_xml().unwrap();
        let mut parsed = DocProperties::default();
        parse_app(&mut parsed, &xml).unwrap();

        assert_eq!(parsed.application, props.application);
        assert_eq!(parsed.company, props.company);
        assert_eq!(parsed.manager, props.manager);
    }

    #[test]
    fn test_empty_props() {
        let props = DocProperties::default();
        assert!(!props.has_core() && !props.has_app());
        assert!(!props.has_core());
        assert!(!props.has_app());

        let with_creator = DocProperties {
            creator: Some("X".to_string()),
            ..Default::default()
        };
        assert!(with_creator.has_core() || with_creator.has_app());
        assert!(with_creator.has_core());
        assert!(!with_creator.has_app());
    }

    #[test]
    fn test_revision_roundtrip() {
        let props = DocProperties {
            revision: Some("3".to_string()),
            ..Default::default()
        };
        assert!(props.has_core());

        let xml = props.to_core_xml().unwrap();
        let parsed = parse_core(&xml).unwrap();
        assert_eq!(parsed.revision.as_deref(), Some("3"));
    }

    #[test]
    fn test_app_version_roundtrip() {
        let props = DocProperties {
            app_version: Some("16.0300".to_string()),
            ..Default::default()
        };
        assert!(props.has_app());

        let xml = props.to_app_xml().unwrap();
        let mut parsed = DocProperties::default();
        parse_app(&mut parsed, &xml).unwrap();
        assert_eq!(parsed.app_version.as_deref(), Some("16.0300"));
    }

    #[test]
    fn test_hyperlink_base_roundtrip() {
        let props = DocProperties {
            hyperlink_base: Some("https://example.com/".to_string()),
            ..Default::default()
        };
        assert!(props.has_app());

        let xml = props.to_app_xml().unwrap();
        let mut parsed = DocProperties::default();
        parse_app(&mut parsed, &xml).unwrap();
        assert_eq!(
            parsed.hyperlink_base.as_deref(),
            Some("https://example.com/")
        );
    }

    #[test]
    fn test_parse_core_with_revision() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
                   xmlns:dc="http://purl.org/dc/elements/1.1/">
  <dc:creator>Author</dc:creator>
  <cp:revision>5</cp:revision>
</cp:coreProperties>"#;

        let props = parse_core(xml.as_bytes()).unwrap();
        assert_eq!(props.creator.as_deref(), Some("Author"));
        assert_eq!(props.revision.as_deref(), Some("5"));
    }

    #[test]
    fn test_parse_app_with_version_and_hyperlink_base() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
            xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Application>Microsoft Excel</Application>
  <AppVersion>16.0300</AppVersion>
  <HyperlinkBase>https://example.com/</HyperlinkBase>
</Properties>"#;

        let mut props = DocProperties::default();
        parse_app(&mut props, xml.as_bytes()).unwrap();
        assert_eq!(props.application.as_deref(), Some("Microsoft Excel"));
        assert_eq!(props.app_version.as_deref(), Some("16.0300"));
        assert_eq!(
            props.hyperlink_base.as_deref(),
            Some("https://example.com/")
        );
    }

    #[test]
    fn app_and_core_text_with_markup_roundtrips() {
        let props = DocProperties {
            company: Some("R&D <x>".to_owned()),
            title: Some("AT&T report".to_owned()),
            ..Default::default()
        };
        let mut app = DocProperties::default();
        parse_app(&mut app, &props.to_app_xml().unwrap()).unwrap();
        assert_eq!(app.company.as_deref(), Some("R&D <x>"));
        let core = parse_core(&props.to_core_xml().unwrap()).unwrap();
        assert_eq!(core.title.as_deref(), Some("AT&T report"));
    }

    #[test]
    fn cdata_and_whitespace_text_is_kept_raw() {
        let xml = r#"<cp:coreProperties xmlns:cp="c" xmlns:dc="d"><dc:title><![CDATA[a <b>]]></dc:title><dc:subject> </dc:subject><dc:creator>  Al  </dc:creator></cp:coreProperties>"#;
        let p = parse_core(xml.as_bytes()).unwrap();
        assert_eq!(p.title.as_deref(), Some("a <b>"));
        assert_eq!(p.subject.as_deref(), Some(" "));
        assert_eq!(p.creator.as_deref(), Some("  Al  "));
    }

    #[test]
    fn extra_core_and_app_fields_survive_parse_to_xml() {
        let core_xml = r#"<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:language>en-US</dc:language><dc:identifier>id-1</dc:identifier><cp:lastPrinted>2024-01-02T03:04:05Z</cp:lastPrinted><cp:version>1.2</cp:version></cp:coreProperties>"#;
        let app_xml = r#"<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><TotalTime>5</TotalTime><DocSecurity>0</DocSecurity></Properties>"#;
        let mut props = parse_core(core_xml.as_bytes()).unwrap();
        parse_app(&mut props, app_xml.as_bytes()).unwrap();
        assert!(props.has_core() && props.has_app() && (props.has_core() || props.has_app()));

        let core = String::from_utf8(props.to_core_xml().unwrap()).unwrap();
        assert!(core.contains("<dc:language>en-US</dc:language>"), "{core}");
        assert!(core.contains("<cp:lastPrinted>2024-01-02T03:04:05Z</cp:lastPrinted>"));
        assert!(core.contains("<dc:identifier>id-1</dc:identifier>"));
        assert!(core.contains("<cp:version>1.2</cp:version>"));
        let app = String::from_utf8(props.to_app_xml().unwrap()).unwrap();
        assert!(app.contains("<TotalTime>5</TotalTime>"), "{app}");
        assert!(app.contains("<DocSecurity>0</DocSecurity>"), "{app}");

        let only_app = DocProperties {
            doc_security: Some("0".to_owned()),
            ..Default::default()
        };
        assert!(only_app.has_app() && !only_app.has_core());
    }

    #[test]
    fn mismatched_end_tag_is_xml_parse_error() {
        let bad = br#"<root xmlns="x"><a></b></root>"#;
        assert!(matches!(parse_core(bad), Err(ModernXlsxError::XmlParse(_))));
        assert!(matches!(
            parse_app(&mut DocProperties::default(), bad),
            Err(ModernXlsxError::XmlParse(_))
        ));
    }

    #[test]
    fn literal_line_ends_are_normalized_and_char_refs_kept() {
        let xml = "<cp:coreProperties xmlns:cp=\"c\" xmlns:dc=\"d\"><dc:title>a\r\nb\rc</dc:title>\
                   <dc:subject>a&#13;&#10;b</dc:subject><dc:description><![CDATA[x\r\ny]]></dc:description>\
                   </cp:coreProperties>";
        let p = parse_core(xml.as_bytes()).unwrap();
        assert_eq!(p.title.as_deref(), Some("a\nb\nc"));
        assert_eq!(p.subject.as_deref(), Some("a\r\nb"));
        assert_eq!(p.description.as_deref(), Some("x\ny"));
    }
}
