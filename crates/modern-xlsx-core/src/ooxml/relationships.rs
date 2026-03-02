use std::collections::HashMap;

use quick_xml::events::{BytesDecl, BytesEnd, BytesStart, Event};
use quick_xml::{Reader, Writer};

use crate::{ModernXlsxError, Result};

// Relationship type URI constants.
const REL_OFFICE_DOCUMENT: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
const REL_WORKSHEET: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet";
const REL_SHARED_STRINGS: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings";
const REL_STYLES: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles";
pub(crate) const REL_COMMENTS: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments";

const RELS_NS: &str = "http://schemas.openxmlformats.org/package/2006/relationships";

/// A single OPC relationship entry.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Relationship {
    /// The relationship identifier (e.g. "rId1").
    pub id: String,
    /// The relationship type URI.
    pub rel_type: String,
    /// The target part path, relative to the source.
    pub target: String,
}

/// A collection of OPC relationships, as found in `.rels` files.
#[derive(Debug, Clone)]
pub struct Relationships {
    /// The ordered list of relationships.
    pub relationships: Vec<Relationship>,
    /// Index from relationship ID to position in the `relationships` vector.
    id_index: HashMap<String, usize>,
}

impl Relationships {
    /// Create an empty `Relationships` collection.
    pub fn new() -> Self {
        Self {
            relationships: Vec::new(),
            id_index: HashMap::new(),
        }
    }

    /// Add a relationship and index it by ID.
    pub fn add(
        &mut self,
        id: impl Into<String>,
        rel_type: impl Into<String>,
        target: impl Into<String>,
    ) {
        let id = id.into();
        let idx = self.relationships.len();
        self.id_index.insert(id.clone(), idx);
        self.relationships.push(Relationship {
            id,
            rel_type: rel_type.into(),
            target: target.into(),
        });
    }

    /// Look up a relationship by its ID.
    pub fn get_by_id(&self, id: &str) -> Option<&Relationship> {
        self.id_index.get(id).map(|&idx| &self.relationships[idx])
    }

    /// Find all relationships matching a given type URI.
    pub fn find_by_type<'a>(&'a self, rel_type: &'a str) -> impl Iterator<Item = &'a Relationship> {
        self.relationships.iter().filter(move |r| r.rel_type == rel_type)
    }

    /// Create the root `_rels/.rels` relationships for a basic workbook.
    ///
    /// Contains a single relationship pointing to `xl/workbook.xml`.
    pub fn root_rels() -> Self {
        let mut rels = Self::new();
        rels.add("rId1", REL_OFFICE_DOCUMENT, "xl/workbook.xml");
        rels
    }

    /// Create the workbook-level relationships (`xl/_rels/workbook.xml.rels`).
    ///
    /// For `sheet_count` sheets, produces:
    /// - rId1..rIdN for worksheets (`worksheets/sheet1.xml`, ...)
    /// - rId(N+1) for sharedStrings (`sharedStrings.xml`)
    /// - rId(N+2) for styles (`styles.xml`)
    pub fn workbook_rels(sheet_count: usize) -> Self {
        let mut rels = Self::new();
        for i in 1..=sheet_count {
            rels.add(
                format!("rId{i}"),
                REL_WORKSHEET,
                format!("worksheets/sheet{i}.xml"),
            );
        }
        let ss_id = sheet_count + 1;
        rels.add(format!("rId{ss_id}"), REL_SHARED_STRINGS, "sharedStrings.xml");
        let st_id = sheet_count + 2;
        rels.add(format!("rId{st_id}"), REL_STYLES, "styles.xml");
        rels
    }

    /// Parse a `.rels` XML file from raw bytes.
    pub fn parse(data: &[u8]) -> Result<Self> {
        let mut reader = Reader::from_reader(data);
        let mut buf = Vec::with_capacity(512);
        let mut rels = Self::new();

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Empty(ref e)) | Ok(Event::Start(ref e)) => {
                    if e.local_name().as_ref() == b"Relationship" {
                        let (mut id, mut rel_type, mut target) =
                            (String::new(), String::new(), String::new());
                        for attr in e.attributes().flatten() {
                            match attr.key.local_name().as_ref() {
                                b"Id" => {
                                    id = std::str::from_utf8(&attr.value).unwrap_or_default().to_owned();
                                }
                                b"Type" => {
                                    rel_type =
                                        std::str::from_utf8(&attr.value).unwrap_or_default().to_owned();
                                }
                                b"Target" => {
                                    target =
                                        std::str::from_utf8(&attr.value).unwrap_or_default().to_owned();
                                }
                                _ => {}
                            }
                        }
                        if !id.is_empty() && !rel_type.is_empty() && !target.is_empty() {
                            rels.add(id, rel_type, target);
                        }
                    }
                }
                Ok(Event::Eof) => break,
                Ok(_) => {}
                Err(e) => {
                    return Err(ModernXlsxError::XmlParse(format!(
                        "error parsing .rels XML: {e}"
                    )));
                }
            }
            buf.clear();
        }

        Ok(rels)
    }

    /// Serialize to XML bytes.
    pub fn to_xml(&self) -> Result<Vec<u8>> {
        let mut buf: Vec<u8> = Vec::with_capacity(512);
        let mut writer = Writer::new(&mut buf);

        // XML declaration.
        writer
            .write_event(Event::Decl(BytesDecl::new("1.0", Some("UTF-8"), Some("yes"))))
            .map_err(|e| ModernXlsxError::XmlWrite(format!("writing XML decl: {e}")))?;

        // <Relationships xmlns="...">
        let mut root = BytesStart::new("Relationships");
        root.push_attribute(("xmlns", RELS_NS));
        writer
            .write_event(Event::Start(root))
            .map_err(|e| ModernXlsxError::XmlWrite(format!("writing Relationships start: {e}")))?;

        // Each <Relationship ... /> (self-closing).
        for rel in &self.relationships {
            let mut elem = BytesStart::new("Relationship");
            elem.push_attribute(("Id", rel.id.as_str()));
            elem.push_attribute(("Type", rel.rel_type.as_str()));
            elem.push_attribute(("Target", rel.target.as_str()));
            writer
                .write_event(Event::Empty(elem))
                .map_err(|e| ModernXlsxError::XmlWrite(format!("writing Relationship: {e}")))?;
        }

        // </Relationships>
        writer
            .write_event(Event::End(BytesEnd::new("Relationships")))
            .map_err(|e| {
                ModernXlsxError::XmlWrite(format!("writing Relationships end: {e}"))
            })?;

        Ok(buf)
    }
}

impl Default for Relationships {
    fn default() -> Self {
        Self::new()
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use pretty_assertions::assert_eq;

    #[test]
    fn test_parse_relationships() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>"#;

        let rels = Relationships::parse(xml.as_bytes()).unwrap();
        assert_eq!(rels.relationships.len(), 3);

        let r1 = rels.get_by_id("rId1").unwrap();
        assert_eq!(r1.rel_type, REL_WORKSHEET);
        assert_eq!(r1.target, "worksheets/sheet1.xml");

        let r2 = rels.get_by_id("rId2").unwrap();
        assert_eq!(r2.rel_type, REL_SHARED_STRINGS);
        assert_eq!(r2.target, "sharedStrings.xml");

        let r3 = rels.get_by_id("rId3").unwrap();
        assert_eq!(r3.rel_type, REL_STYLES);
        assert_eq!(r3.target, "styles.xml");
    }

    #[test]
    fn test_find_by_type() {
        let mut rels = Relationships::new();
        rels.add("rId1", REL_WORKSHEET, "worksheets/sheet1.xml");
        rels.add("rId2", REL_WORKSHEET, "worksheets/sheet2.xml");
        rels.add("rId3", REL_SHARED_STRINGS, "sharedStrings.xml");

        let worksheets: Vec<_> = rels.find_by_type(REL_WORKSHEET).collect();
        assert_eq!(worksheets.len(), 2);
        assert_eq!(worksheets[0].target, "worksheets/sheet1.xml");
        assert_eq!(worksheets[1].target, "worksheets/sheet2.xml");

        let shared: Vec<_> = rels.find_by_type(REL_SHARED_STRINGS).collect();
        assert_eq!(shared.len(), 1);
        assert_eq!(shared[0].target, "sharedStrings.xml");

        let styles: Vec<_> = rels.find_by_type(REL_STYLES).collect();
        assert!(styles.is_empty());
    }

    #[test]
    fn test_write_relationships() {
        let mut rels = Relationships::new();
        rels.add("rId1", REL_WORKSHEET, "worksheets/sheet1.xml");
        rels.add("rId2", REL_SHARED_STRINGS, "sharedStrings.xml");

        let xml = rels.to_xml().unwrap();

        // Re-parse and verify roundtrip.
        let rels2 = Relationships::parse(&xml).unwrap();
        assert_eq!(rels2.relationships.len(), 2);

        let r1 = rels2.get_by_id("rId1").unwrap();
        assert_eq!(r1.rel_type, REL_WORKSHEET);
        assert_eq!(r1.target, "worksheets/sheet1.xml");

        let r2 = rels2.get_by_id("rId2").unwrap();
        assert_eq!(r2.rel_type, REL_SHARED_STRINGS);
        assert_eq!(r2.target, "sharedStrings.xml");
    }

    #[test]
    fn test_root_rels_for_basic_workbook() {
        let rels = Relationships::root_rels();
        assert_eq!(rels.relationships.len(), 1);

        let r1 = rels.get_by_id("rId1").unwrap();
        assert_eq!(r1.rel_type, REL_OFFICE_DOCUMENT);
        assert_eq!(r1.target, "xl/workbook.xml");
    }

    #[test]
    fn test_workbook_rels() {
        let rels = Relationships::workbook_rels(2);
        assert_eq!(rels.relationships.len(), 4);

        // rId1, rId2 = worksheets
        let r1 = rels.get_by_id("rId1").unwrap();
        assert_eq!(r1.rel_type, REL_WORKSHEET);
        assert_eq!(r1.target, "worksheets/sheet1.xml");

        let r2 = rels.get_by_id("rId2").unwrap();
        assert_eq!(r2.rel_type, REL_WORKSHEET);
        assert_eq!(r2.target, "worksheets/sheet2.xml");

        // rId3 = sharedStrings
        let r3 = rels.get_by_id("rId3").unwrap();
        assert_eq!(r3.rel_type, REL_SHARED_STRINGS);
        assert_eq!(r3.target, "sharedStrings.xml");

        // rId4 = styles
        let r4 = rels.get_by_id("rId4").unwrap();
        assert_eq!(r4.rel_type, REL_STYLES);
        assert_eq!(r4.target, "styles.xml");
    }
}
