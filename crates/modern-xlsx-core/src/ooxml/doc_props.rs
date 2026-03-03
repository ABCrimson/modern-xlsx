//! Document properties parser and writer.
//!
//! Handles `docProps/core.xml` (Dublin Core metadata) and `docProps/app.xml`
//! (extended application properties).

use quick_xml::events::{BytesDecl, BytesEnd, BytesStart, BytesText, Event};
use quick_xml::{Reader, Writer};

use serde::{Deserialize, Serialize};

use crate::{ModernXlsxError, Result};

// Namespace URIs used in core.xml.
const CP_NS: &str = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties";
const DC_NS: &str = "http://purl.org/dc/elements/1.1/";
const DCTERMS_NS: &str = "http://purl.org/dc/terms/";
const XSI_NS: &str = "http://www.w3.org/2001/XMLSchema-instance";

// Namespace URI for app.xml.
const APP_NS: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties";
const VT_NS: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes";

/// Combined document properties from `docProps/core.xml` and `docProps/app.xml`.
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
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub revision: Option<String>,
}

/// Parse `docProps/core.xml` into a [`DocProperties`].
///
/// The XML uses Dublin Core namespaces (dc:, dcterms:, cp:). We match on
/// local_name() to ignore namespace prefixes.
pub fn parse_core(data: &[u8]) -> Result<DocProperties> {
    let mut reader = Reader::from_reader(data);
    reader.config_mut().trim_text(true);

    let mut buf = Vec::with_capacity(512);
    let mut props = DocProperties::default();

    // Track which element we are inside to collect text.
    let mut current_element: Option<String> = None;
    let mut text_buf = String::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) => {
                let local = e.local_name();
                match local.as_ref() {
                    b"title" | b"creator" | b"subject" | b"keywords" | b"description"
                    | b"lastModifiedBy" | b"created" | b"modified" | b"category"
                    | b"contentStatus" | b"revision" => {
                        current_element =
                            Some(std::str::from_utf8(local.as_ref())
                                .unwrap_or_default()
                                .to_owned());
                        text_buf.clear();
                    }
                    _ => {}
                }
            }
            Ok(Event::Text(ref e)) => {
                if current_element.is_some() {
                    text_buf.push_str(std::str::from_utf8(e.as_ref()).unwrap_or_default());
                }
            }
            Ok(Event::End(_)) => {
                if let Some(ref elem) = current_element {
                    let val = if text_buf.is_empty() {
                        None
                    } else {
                        Some(std::mem::take(&mut text_buf))
                    };
                    match elem.as_str() {
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
                        _ => {}
                    }
                    current_element = None;
                }
            }
            Ok(Event::Eof) => break,
            Err(err) => {
                return Err(ModernXlsxError::XmlParse(format!(
                    "error parsing docProps/core.xml: {err}"
                )));
            }
            _ => {}
        }
        buf.clear();
    }

    Ok(props)
}

/// Parse `docProps/app.xml` and fill in the application/company/manager fields
/// on an existing [`DocProperties`].
pub fn parse_app(props: &mut DocProperties, data: &[u8]) -> Result<()> {
    let mut reader = Reader::from_reader(data);
    reader.config_mut().trim_text(true);

    let mut buf = Vec::with_capacity(512);
    let mut current_element: Option<String> = None;
    let mut text_buf = String::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) => {
                let local = e.local_name();
                match local.as_ref() {
                    b"Application" | b"Company" | b"Manager" | b"AppVersion" | b"HyperlinkBase" => {
                        current_element =
                            Some(std::str::from_utf8(local.as_ref())
                                .unwrap_or_default()
                                .to_owned());
                        text_buf.clear();
                    }
                    _ => {}
                }
            }
            Ok(Event::Text(ref e)) => {
                if current_element.is_some() {
                    text_buf.push_str(std::str::from_utf8(e.as_ref()).unwrap_or_default());
                }
            }
            Ok(Event::End(_)) => {
                if let Some(ref elem) = current_element {
                    let val = if text_buf.is_empty() {
                        None
                    } else {
                        Some(std::mem::take(&mut text_buf))
                    };
                    match elem.as_str() {
                        "Application" => props.application = val,
                        "Company" => props.company = val,
                        "Manager" => props.manager = val,
                        "AppVersion" => props.app_version = val,
                        "HyperlinkBase" => props.hyperlink_base = val,
                        _ => {}
                    }
                    current_element = None;
                }
            }
            Ok(Event::Eof) => break,
            Err(err) => {
                return Err(ModernXlsxError::XmlParse(format!(
                    "error parsing docProps/app.xml: {err}"
                )));
            }
            _ => {}
        }
        buf.clear();
    }

    Ok(())
}

impl DocProperties {
    /// Serialize to `docProps/core.xml` bytes.
    pub fn to_core_xml(&self) -> Result<Vec<u8>> {
        let mut buf: Vec<u8> = Vec::with_capacity(1024);
        let mut writer = Writer::new(&mut buf);

        let map_err = |e: std::io::Error| ModernXlsxError::XmlWrite(e.to_string());

        // XML declaration.
        writer
            .write_event(Event::Decl(BytesDecl::new("1.0", Some("UTF-8"), Some("yes"))))
            .map_err(map_err)?;

        // <cp:coreProperties>
        let mut root = BytesStart::new("cp:coreProperties");
        root.push_attribute(("xmlns:cp", CP_NS));
        root.push_attribute(("xmlns:dc", DC_NS));
        root.push_attribute(("xmlns:dcterms", DCTERMS_NS));
        root.push_attribute(("xmlns:xsi", XSI_NS));
        writer.write_event(Event::Start(root)).map_err(map_err)?;

        // Helper closure for writing simple text elements.
        let write_elem =
            |writer: &mut Writer<&mut Vec<u8>>, tag: &str, value: &Option<String>| -> Result<()> {
                if let Some(ref v) = *value {
                    writer
                        .write_event(Event::Start(BytesStart::new(tag)))
                        .map_err(map_err)?;
                    writer
                        .write_event(Event::Text(BytesText::new(v)))
                        .map_err(map_err)?;
                    writer
                        .write_event(Event::End(BytesEnd::new(tag)))
                        .map_err(map_err)?;
                }
                Ok(())
            };

        write_elem(&mut writer, "dc:title", &self.title)?;
        write_elem(&mut writer, "dc:subject", &self.subject)?;
        write_elem(&mut writer, "dc:creator", &self.creator)?;
        write_elem(&mut writer, "cp:keywords", &self.keywords)?;
        write_elem(&mut writer, "dc:description", &self.description)?;
        write_elem(&mut writer, "cp:lastModifiedBy", &self.last_modified_by)?;
        write_elem(&mut writer, "cp:category", &self.category)?;
        write_elem(&mut writer, "cp:contentStatus", &self.content_status)?;
        write_elem(&mut writer, "cp:revision", &self.revision)?;

        // dcterms:created and dcterms:modified have xsi:type="dcterms:W3CDTF"
        if let Some(ref v) = self.created {
            let mut elem = BytesStart::new("dcterms:created");
            elem.push_attribute(("xsi:type", "dcterms:W3CDTF"));
            writer.write_event(Event::Start(elem)).map_err(map_err)?;
            writer
                .write_event(Event::Text(BytesText::new(v)))
                .map_err(map_err)?;
            writer
                .write_event(Event::End(BytesEnd::new("dcterms:created")))
                .map_err(map_err)?;
        }

        if let Some(ref v) = self.modified {
            let mut elem = BytesStart::new("dcterms:modified");
            elem.push_attribute(("xsi:type", "dcterms:W3CDTF"));
            writer.write_event(Event::Start(elem)).map_err(map_err)?;
            writer
                .write_event(Event::Text(BytesText::new(v)))
                .map_err(map_err)?;
            writer
                .write_event(Event::End(BytesEnd::new("dcterms:modified")))
                .map_err(map_err)?;
        }

        // </cp:coreProperties>
        writer
            .write_event(Event::End(BytesEnd::new("cp:coreProperties")))
            .map_err(map_err)?;

        Ok(buf)
    }

    /// Serialize to `docProps/app.xml` bytes.
    pub fn to_app_xml(&self) -> Result<Vec<u8>> {
        let mut buf: Vec<u8> = Vec::with_capacity(512);
        let mut writer = Writer::new(&mut buf);

        let map_err = |e: std::io::Error| ModernXlsxError::XmlWrite(e.to_string());

        // XML declaration.
        writer
            .write_event(Event::Decl(BytesDecl::new("1.0", Some("UTF-8"), Some("yes"))))
            .map_err(map_err)?;

        // <Properties>
        let mut root = BytesStart::new("Properties");
        root.push_attribute(("xmlns", APP_NS));
        root.push_attribute(("xmlns:vt", VT_NS));
        writer.write_event(Event::Start(root)).map_err(map_err)?;

        let write_elem =
            |writer: &mut Writer<&mut Vec<u8>>, tag: &str, value: &Option<String>| -> Result<()> {
                if let Some(ref v) = *value {
                    writer
                        .write_event(Event::Start(BytesStart::new(tag)))
                        .map_err(map_err)?;
                    writer
                        .write_event(Event::Text(BytesText::new(v)))
                        .map_err(map_err)?;
                    writer
                        .write_event(Event::End(BytesEnd::new(tag)))
                        .map_err(map_err)?;
                }
                Ok(())
            };

        write_elem(&mut writer, "Application", &self.application)?;
        write_elem(&mut writer, "Company", &self.company)?;
        write_elem(&mut writer, "Manager", &self.manager)?;
        write_elem(&mut writer, "AppVersion", &self.app_version)?;
        write_elem(&mut writer, "HyperlinkBase", &self.hyperlink_base)?;

        // </Properties>
        writer
            .write_event(Event::End(BytesEnd::new("Properties")))
            .map_err(map_err)?;

        Ok(buf)
    }

    /// Returns `true` if all fields are `None` (i.e., no metadata set).
    pub fn is_empty(&self) -> bool {
        self.title.is_none()
            && self.subject.is_none()
            && self.creator.is_none()
            && self.keywords.is_none()
            && self.description.is_none()
            && self.last_modified_by.is_none()
            && self.created.is_none()
            && self.modified.is_none()
            && self.category.is_none()
            && self.content_status.is_none()
            && self.application.is_none()
            && self.company.is_none()
            && self.manager.is_none()
            && self.app_version.is_none()
            && self.hyperlink_base.is_none()
            && self.revision.is_none()
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
    }

    /// Returns `true` if any app property (docProps/app.xml) is set.
    pub fn has_app(&self) -> bool {
        self.application.is_some()
            || self.company.is_some()
            || self.manager.is_some()
            || self.app_version.is_some()
            || self.hyperlink_base.is_some()
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
        assert_eq!(
            props.created.as_deref(),
            Some("2024-01-15T10:30:00Z")
        );
        assert_eq!(
            props.modified.as_deref(),
            Some("2024-06-20T14:00:00Z")
        );
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
        assert!(props.is_empty());
        assert!(!props.has_core());
        assert!(!props.has_app());

        let with_creator = DocProperties {
            creator: Some("X".to_string()),
            ..Default::default()
        };
        assert!(!with_creator.is_empty());
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
        assert_eq!(parsed.hyperlink_base.as_deref(), Some("https://example.com/"));
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
        assert_eq!(props.hyperlink_base.as_deref(), Some("https://example.com/"));
    }
}
