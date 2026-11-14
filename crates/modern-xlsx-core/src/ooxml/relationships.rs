//! Open Packaging Conventions relationships: the `.rels` parts (`_rels/.rels`,
//! `xl/_rels/workbook.xml.rels`, ...), the `REL_*` relationship type URIs and
//! [`Relationships`] to parse, edit and write one.

use core::hint::cold_path;
use std::borrow::Cow;
use std::collections::HashMap;

use quick_xml::events::{BytesEnd, BytesStart, Event};
use quick_xml::{Reader, Writer};

use crate::ooxml::write_xml_decl;
use crate::ooxml::xml::{attr_str, xml_attr};
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
pub(crate) const REL_TABLE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/table";
pub(crate) const REL_CHART: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart";
/// Chart → its `c:userShapes` part (the ECMA-376 Part 1 Chart Drawing
/// Part), the text boxes and shapes drawn over the chart.
pub(crate) const REL_CHART_USER_SHAPES: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartUserShapes";
/// Chart → its style part (`style{n}.xml`); the URI Microsoft Excel 16.0 wrote
/// in a workbook saved on 2026-11-11.
pub(crate) const REL_CHART_STYLE: &str =
    "http://schemas.microsoft.com/office/2011/relationships/chartStyle";
/// Chart → its colour-style part (`colors{n}.xml`); the URI Microsoft Excel
/// 16.0 wrote in a workbook saved on 2026-11-11.
pub(crate) const REL_CHART_COLOR_STYLE: &str =
    "http://schemas.microsoft.com/office/2011/relationships/chartColorStyle";
pub(crate) const REL_DRAWING: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing";
/// Worksheet → a VML drawing, which the sheet names by `<legacyDrawing>`
/// (note and form-control shapes) or `<legacyDrawingHF>` (header and footer
/// pictures).
pub(crate) const REL_VML_DRAWING: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing";
/// Worksheet → hyperlink target (ECMA-376 Part 1 §18.3.1.47): an external URL,
/// `mailto:` address or file, always with `TargetMode="External"`.
pub(crate) const REL_HYPERLINK: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";
pub(crate) const REL_PIVOT_TABLE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable";
pub(crate) const REL_PIVOT_CACHE_DEF: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition";
pub(crate) const REL_PIVOT_CACHE_REC: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheRecords";
pub(crate) const REL_THREADED_COMMENTS: &str =
    "http://schemas.microsoft.com/office/2017/10/relationships/threadedComment";
pub(crate) const REL_PERSONS: &str =
    "http://schemas.microsoft.com/office/2017/10/relationships/person";
pub(crate) const REL_SLICER: &str = "http://schemas.microsoft.com/office/2007/relationships/slicer";
pub(crate) const REL_SLICER_CACHE: &str =
    "http://schemas.microsoft.com/office/2007/relationships/slicerCache";
pub(crate) const REL_TIMELINE: &str =
    "http://schemas.microsoft.com/office/2011/relationships/timeline";
pub(crate) const REL_TIMELINE_CACHE: &str =
    "http://schemas.microsoft.com/office/2011/relationships/timelineCache";
/// Workbook → `xl/metadata.xml` (ECMA-376 Part 1 §12.3.10 Metadata Part); the
/// URI Microsoft Excel 16.0 (build 20527) wrote for a `=SEQUENCE(3)` workbook
/// saved on 2026-11-11.
pub(crate) const REL_SHEET_METADATA: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sheetMetadata";
/// Workbook → `xl/theme/theme1.xml` (ECMA-376 Part 1 §14.2.7 Theme Part).
pub(crate) const REL_THEME: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme";
/// Workbook → `xl/externalLinks/externalLink{n}.xml` (ECMA-376 Part 1
/// §12.3.9 External Workbook References Part).
pub(crate) const REL_EXTERNAL_LINK: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink";
/// Workbook → `xl/calcChain.xml` (ECMA-376 Part 1 §12.3.1 Calculation Chain
/// Part).
pub(crate) const REL_CALC_CHAIN: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain";
/// Package → `docProps/core.xml` (ECMA-376 Part 1 §15.2.12.1 Core File
/// Properties Part).
pub(crate) const REL_CORE_PROPERTIES: &str =
    "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties";
/// Package → `docProps/app.xml` (ECMA-376 Part 1 §15.2.12.3 Extended File
/// Properties Part).
pub(crate) const REL_EXTENDED_PROPERTIES: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties";
/// Package → `docProps/custom.xml` (ECMA-376 Part 1 §15.2.12.2 Custom File
/// Properties Part).
pub(crate) const REL_CUSTOM_PROPERTIES: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties";
/// Workbook → `customXml/item{n}.xml` (the ECMA-376 Part 1 Custom XML Data
/// Storage Part).
pub(crate) const REL_CUSTOM_XML: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml";

/// Every relationship type constant of this module. [`Relationships::parse`]
/// maps a matching `Type` to the constant, so known types borrow instead of
/// allocating.
const KNOWN_REL_TYPES: [&str; 29] = [
    REL_OFFICE_DOCUMENT,
    REL_WORKSHEET,
    REL_SHARED_STRINGS,
    REL_STYLES,
    REL_COMMENTS,
    REL_TABLE,
    REL_CHART,
    REL_CHART_USER_SHAPES,
    REL_CHART_STYLE,
    REL_CHART_COLOR_STYLE,
    REL_DRAWING,
    REL_HYPERLINK,
    REL_PIVOT_TABLE,
    REL_PIVOT_CACHE_DEF,
    REL_PIVOT_CACHE_REC,
    REL_THREADED_COMMENTS,
    REL_PERSONS,
    REL_SLICER,
    REL_SLICER_CACHE,
    REL_TIMELINE,
    REL_TIMELINE_CACHE,
    REL_SHEET_METADATA,
    REL_THEME,
    REL_EXTERNAL_LINK,
    REL_CALC_CHAIN,
    REL_CORE_PROPERTIES,
    REL_EXTENDED_PROPERTIES,
    REL_CUSTOM_PROPERTIES,
    REL_CUSTOM_XML,
];

const RELS_NS: &str = "http://schemas.openxmlformats.org/package/2006/relationships";

/// Format an rId string (e.g. "rId3") without `format!()` heap overhead.
#[inline]
pub(crate) fn make_rid(num: impl itoa::Integer) -> String {
    let mut s = String::with_capacity(12);
    s.push_str("rId");
    s.push_str(itoa::Buffer::new().format(num));
    s
}

/// The ZIP entry name of `target`, a relationship target in the .rels of the
/// part at `part_path` (a ZIP entry name too).
///
/// A target with a leading `/` is package-absolute: that `/` is dropped.
/// Any other target is relative to the folder of `part_path`: `.` is skipped
/// and `..` climbs one folder, stopping at the package root, as RFC 3986
/// remove_dot_segments does. Empty segments are skipped too, a leniency: a
/// valid OPC part name has none.
pub(crate) fn resolve_target(part_path: &str, target: &str) -> String {
    if let Some(absolute) = target.strip_prefix('/') {
        return absolute.to_owned();
    }
    let mut parts: Vec<&str> = part_path.split('/').collect();
    parts.pop();
    for segment in target.split('/') {
        match segment {
            "" | "." => {}
            ".." => {
                parts.pop();
            }
            other => parts.push(other),
        }
    }
    parts.join("/")
}

/// The ZIP entry name of the .rels part of the part at `part_path`:
/// `xl/drawings/drawing1.xml` gives `xl/drawings/_rels/drawing1.xml.rels`, and
/// a part in the package root `_rels/<name>.rels`.
pub(crate) fn rels_path(part_path: &str) -> String {
    match part_path.rsplit_once('/') {
        Some((dir, file)) => format!("{dir}/_rels/{file}.rels"),
        None => format!("_rels/{part_path}.rels"),
    }
}

/// The `Type` constant equal to `raw`, borrowed, or `raw` itself, owned.
#[inline]
fn known_rel_type(raw: Cow<'_, str>) -> Cow<'static, str> {
    match KNOWN_REL_TYPES.iter().find(|&&known| known == raw) {
        Some(&known) => Cow::Borrowed(known),
        None => Cow::Owned(raw.into_owned()),
    }
}

/// How a relationship's `Target` resolves (OPC, ECMA-376 Part 2,
/// `ST_TargetMode`).
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub enum TargetMode {
    /// A part name inside the package, relative to the source part (the
    /// default when `TargetMode` is absent).
    #[default]
    Internal,
    /// A resource outside the package: a URL, `mailto:` address or file path.
    External,
}

/// A single OPC relationship entry.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Relationship {
    /// The relationship identifier (e.g. "rId1").
    pub id: String,
    /// The relationship type URI.
    pub rel_type: Cow<'static, str>,
    /// The target: a part path relative to the source when `target_mode` is
    /// `Internal`, else an external URI.
    pub target: Cow<'static, str>,
    /// Whether `target` names a package part or an external resource.
    pub target_mode: TargetMode,
}

/// The number `n` of an `rId<n>` identifier; `None` for any other ID or one
/// that does not fit a `u64`.
#[inline]
fn r_id_number(id: &str) -> Option<u64> {
    id.strip_prefix("rId")?.parse::<u64>().ok()
}

/// A collection of OPC relationships, as found in `.rels` files.
#[derive(Debug, Clone)]
pub struct Relationships {
    /// The ordered list of relationships.
    entries: Vec<Relationship>,
    /// Index from relationship ID to position in `entries`.
    id_index: HashMap<String, usize>,
    /// One past the highest `rId<n>` number held (1 when there is none),
    /// saturating at `u64::MAX`; kept current by `push` and `retain` so
    /// [`Self::next_r_id`] is O(1); the saturated case reads `low_free`.
    next_id: u64,
    /// The lowest `n` for which `rId<n>` is not held (1 when `rId1` is free).
    /// Every `rId1..rId<low_free - 1>` is held. It only advances on `push`
    /// (each number is probed once) and is recomputed by `retain`.
    low_free: u64,
}

impl Default for Relationships {
    fn default() -> Self {
        Self::new()
    }
}

impl Relationships {
    /// Create an empty `Relationships` collection.
    pub fn new() -> Self {
        Self {
            entries: Vec::new(),
            id_index: HashMap::new(),
            next_id: 1,
            low_free: 1,
        }
    }

    /// Add an internal relationship and index it by ID.
    pub fn add(
        &mut self,
        id: impl Into<String>,
        rel_type: impl Into<Cow<'static, str>>,
        target: impl Into<Cow<'static, str>>,
    ) {
        self.push(Relationship {
            id: id.into(),
            rel_type: rel_type.into(),
            target: target.into(),
            target_mode: TargetMode::Internal,
        });
    }

    /// Add `rel` as it is and index it by ID.
    pub(crate) fn push(&mut self, rel: Relationship) {
        if let Some(n) = r_id_number(&rel.id) {
            self.next_id = self.next_id.max(n.saturating_add(1));
            if n == self.low_free {
                self.id_index.insert(rel.id.clone(), self.entries.len());
                self.low_free = n.saturating_add(1);
                self.advance_low_free();
                self.entries.push(rel);
                return;
            }
        }
        self.id_index.insert(rel.id.clone(), self.entries.len());
        self.entries.push(rel);
    }

    /// Move `low_free` past every consecutively held `rId<n>`; each number is
    /// probed at most once over the life of the collection (until `retain`).
    fn advance_low_free(&mut self) {
        while self.low_free < u64::MAX && self.id_index.contains_key(&make_rid(self.low_free)) {
            self.low_free += 1;
        }
    }

    /// Add an external relationship (`TargetMode="External"`) under the next
    /// free rId (see [`Self::next_r_id`]) and return that rId.
    pub fn add_external(
        &mut self,
        rel_type: impl Into<Cow<'static, str>>,
        target: impl Into<Cow<'static, str>>,
    ) -> String {
        let id = make_rid(self.next_r_id());
        self.push(Relationship {
            id: id.clone(),
            rel_type: rel_type.into(),
            target: target.into(),
            target_mode: TargetMode::External,
        });
        id
    }

    /// Keep only the relationships for which `f` returns `true`, in order,
    /// and rebuild the ID index and the next free rId number.
    pub fn retain(&mut self, f: impl FnMut(&Relationship) -> bool) {
        self.entries.retain(f);
        self.id_index = self
            .entries
            .iter()
            .enumerate()
            .map(|(idx, rel)| (rel.id.clone(), idx))
            .collect();
        self.next_id = self
            .entries
            .iter()
            .filter_map(|r| r_id_number(&r.id))
            .max()
            .map_or(1, |n| n.saturating_add(1));
        self.low_free = 1;
        self.advance_low_free();
    }

    /// Look up a relationship by its ID.
    pub fn get_by_id(&self, id: &str) -> Option<&Relationship> {
        self.id_index.get(id).map(|&idx| &self.entries[idx])
    }

    /// The relationships in document order.
    pub fn iter(&self) -> std::slice::Iter<'_, Relationship> {
        self.entries.iter()
    }

    /// The number of relationships.
    pub fn len(&self) -> usize {
        self.entries.len()
    }

    /// Whether the collection holds no relationship.
    pub fn is_empty(&self) -> bool {
        self.entries.is_empty()
    }

    /// Find all relationships matching a given type URI.
    pub fn find_by_type<'a>(&'a self, rel_type: &'a str) -> impl Iterator<Item = &'a Relationship> {
        self.entries.iter().filter(move |r| r.rel_type == rel_type)
    }

    /// The next available rId number: one past the highest `rId<n>` held, 1
    /// when there is none. O(1); IDs not of the `rId<n>` form are ignored.
    /// When `rId18446744073709551615` is itself held, no number lies past the
    /// highest, so the lowest free number is returned instead.
    ///
    /// Invariant: the number returned is never held, so an ID already present
    /// (parsed or added) is never reissued; a collection that never holds
    /// `rId18446744073709551615` issues exactly one past its highest number.
    /// Cost: the saturated case reads the maintained `low_free` cursor, which
    /// only advances, so N allocations probe each low number at most once
    /// (amortized O(1) per allocation, never a rescan from 1).
    pub fn next_r_id(&self) -> u64 {
        if self.next_id == u64::MAX && self.id_index.contains_key(&make_rid(u64::MAX)) {
            return self.low_free;
        }
        self.next_id
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
                make_rid(i),
                REL_WORKSHEET,
                format!("worksheets/sheet{i}.xml"),
            );
        }
        rels.add(
            make_rid(sheet_count + 1),
            REL_SHARED_STRINGS,
            "sharedStrings.xml",
        );
        rels.add(make_rid(sheet_count + 2), REL_STYLES, "styles.xml");
        rels
    }

    /// Parse a `.rels` XML file from raw bytes.
    ///
    /// Attribute values are decoded (entities and character references
    /// resolved), a known `Type` borrows its constant, and
    /// `TargetMode="External"` gives [`TargetMode::External`]. A
    /// `<Relationship>` missing `Id`, `Type` or `Target` is skipped.
    ///
    /// # Errors
    ///
    /// `XmlParse` for malformed XML, a malformed attribute list or an
    /// attribute value that does not decode.
    pub fn parse(data: &[u8]) -> Result<Self> {
        let mut reader = Reader::from_reader(data);
        let mut buf = Vec::with_capacity(512);
        let mut rels = Self::new();

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Empty(ref e) | Event::Start(ref e))
                    if e.local_name().as_ref() == "Relationship" =>
                {
                    let mut id = String::new();
                    let mut rel_type: Cow<'static, str> = Cow::Borrowed("");
                    let mut target = String::new();
                    let mut target_mode = TargetMode::Internal;
                    for attr in e.attributes() {
                        let attr = attr.map_err(|err| {
                            cold_path();
                            ModernXlsxError::XmlParse(format!("<Relationship> attributes: {err}"))
                        })?;
                        match attr.key.local_name().as_ref() {
                            "Id" => id = attr_str(&attr)?.into_owned(),
                            "Type" => rel_type = known_rel_type(attr_str(&attr)?),
                            "Target" => target = attr_str(&attr)?.into_owned(),
                            "TargetMode" => {
                                target_mode = match attr_str(&attr)?.trim_ascii() {
                                    "External" => TargetMode::External,
                                    _ => TargetMode::Internal,
                                };
                            }
                            _ => {}
                        }
                    }
                    if !id.is_empty() && !rel_type.is_empty() && !target.is_empty() {
                        rels.push(Relationship {
                            id,
                            rel_type,
                            target: Cow::Owned(target),
                            target_mode,
                        });
                    }
                }
                Ok(Event::Eof) => break,
                Ok(_) => {}
                Err(e) => {
                    cold_path();
                    return Err(ModernXlsxError::XmlParse(format!(
                        "error parsing .rels XML: {e}"
                    )));
                }
            }
            buf.clear();
        }

        Ok(rels)
    }

    /// Serialize to XML bytes. An external relationship gets
    /// `TargetMode="External"`; an internal one omits the attribute.
    pub fn to_xml(&self) -> Result<Vec<u8>> {
        let mut buf: Vec<u8> = Vec::with_capacity(512);
        let mut writer = Writer::new(&mut buf);

        write_xml_decl(&mut writer)?;

        // <Relationships xmlns="...">
        let mut root = BytesStart::new("Relationships");
        root.push_attribute(("xmlns", RELS_NS));
        writer
            .write_event(Event::Start(root))
            .map_err(|e| ModernXlsxError::XmlWrite(format!("writing Relationships start: {e}")))?;

        // Each <Relationship ... /> (self-closing).
        for rel in &self.entries {
            let mut elem = BytesStart::new("Relationship");
            elem.push_attribute(xml_attr("Id", &rel.id));
            elem.push_attribute(xml_attr("Type", &rel.rel_type));
            elem.push_attribute(xml_attr("Target", &rel.target));
            if rel.target_mode == TargetMode::External {
                elem.push_attribute(("TargetMode", "External"));
            }
            writer
                .write_event(Event::Empty(elem))
                .map_err(|e| ModernXlsxError::XmlWrite(format!("writing Relationship: {e}")))?;
        }

        // </Relationships>
        writer
            .write_event(Event::End(BytesEnd::new("Relationships")))
            .map_err(|e| ModernXlsxError::XmlWrite(format!("writing Relationships end: {e}")))?;

        Ok(buf)
    }
}

impl<'a> IntoIterator for &'a Relationships {
    type Item = &'a Relationship;
    type IntoIter = std::slice::Iter<'a, Relationship>;

    fn into_iter(self) -> Self::IntoIter {
        self.entries.iter()
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use pretty_assertions::assert_eq;

    #[test]
    fn resolve_target_walks_relative_and_absolute_targets() {
        let sheet = "xl/worksheets/sheet1.xml";
        assert_eq!(
            resolve_target(sheet, "../../drawings/drawing1.xml"),
            "drawings/drawing1.xml"
        );
        assert_eq!(resolve_target(sheet, "/xl/media/a.png"), "xl/media/a.png");
        assert_eq!(
            resolve_target("xl/worksheets/sheet1.xml", "../drawings/./d.xml"),
            "xl/drawings/d.xml"
        );
        assert_eq!(
            resolve_target(sheet, "../comments1.xml"),
            "xl/comments1.xml"
        );
        assert_eq!(
            resolve_target(sheet, "comments1.xml"),
            "xl/worksheets/comments1.xml"
        );
        assert_eq!(
            resolve_target("xl/worksheets/sub/s.xml", "../../comments1.xml"),
            "xl/comments1.xml"
        );
        // One leading '/' marks a package-absolute part name; a '//'
        // network-path reference is not a part name.
        assert_eq!(
            resolve_target(sheet, "//xl/comments1.xml"),
            "/xl/comments1.xml"
        );
        // A part in the package root; '..' stops at the root.
        assert_eq!(resolve_target("drawing1.xml", "a/../../b.xml"), "b.xml");
    }

    #[test]
    fn rels_path_names_the_rels_part() {
        assert_eq!(
            rels_path("xl/drawings/drawing1.xml"),
            "xl/drawings/_rels/drawing1.xml.rels"
        );
        assert_eq!(rels_path("drawing1.xml"), "_rels/drawing1.xml.rels");
    }

    #[test]
    fn test_parse_relationships() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>"#;

        let rels = Relationships::parse(xml.as_bytes()).unwrap();
        assert_eq!(rels.entries.len(), 3);

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
        assert_eq!(rels2.entries.len(), 2);

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
        assert_eq!(rels.entries.len(), 1);

        let r1 = rels.get_by_id("rId1").unwrap();
        assert_eq!(r1.rel_type, REL_OFFICE_DOCUMENT);
        assert_eq!(r1.target, "xl/workbook.xml");
    }

    #[test]
    fn test_workbook_rels() {
        let rels = Relationships::workbook_rels(2);
        assert_eq!(rels.entries.len(), 4);

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

    const EXTERNAL_RELS: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://e.com/?a=1&amp;b=2" TargetMode="External"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="../comments1.xml"/></Relationships>"#;

    #[test]
    fn test_parse_external_target_mode_and_decoded_target() {
        let rels = Relationships::parse(EXTERNAL_RELS.as_bytes()).unwrap();
        let r1 = rels.get_by_id("rId1").unwrap();
        assert_eq!(r1.rel_type, REL_HYPERLINK);
        assert_eq!(r1.target, "https://e.com/?a=1&b=2");
        assert_eq!(r1.target_mode, TargetMode::External);
        let r2 = rels.get_by_id("rId2").unwrap();
        assert_eq!(r2.target_mode, TargetMode::Internal);
    }

    #[test]
    fn test_known_rel_types_are_borrowed_on_parse() {
        let rels = Relationships::parse(EXTERNAL_RELS.as_bytes()).unwrap();
        for rel in &rels {
            assert!(
                matches!(rel.rel_type, Cow::Borrowed(_)),
                "{} should map to its constant",
                rel.rel_type
            );
        }
    }

    #[test]
    fn test_to_xml_writes_target_mode_and_escapes_once() {
        let rels = Relationships::parse(EXTERNAL_RELS.as_bytes()).unwrap();
        let xml = String::from_utf8(rels.to_xml().unwrap()).unwrap();
        assert!(xml.contains(r#"TargetMode="External""#), "{xml}");
        assert_eq!(xml.matches("&amp;").count(), 1, "{xml}");
        assert!(
            xml.contains(r#"Target="https://e.com/?a=1&amp;b=2""#),
            "{xml}"
        );
        // The internal relationship carries no TargetMode (Internal is the default).
        assert_eq!(xml.matches("TargetMode").count(), 1, "{xml}");
    }

    #[test]
    fn test_next_r_id_skips_non_numeric_ids() {
        let mut rels = Relationships::new();
        rels.add("rId2", REL_WORKSHEET, "a.xml");
        rels.add("rId7", REL_WORKSHEET, "b.xml");
        rels.add("foo", REL_WORKSHEET, "c.xml");
        assert_eq!(rels.next_r_id(), 8);
        assert_eq!(make_rid(rels.next_r_id()), "rId8");
    }

    #[test]
    fn test_add_external_ids_follow_the_highest_r_id() {
        let mut rels = Relationships::new();
        rels.add("rId2", REL_WORKSHEET, "a.xml");
        rels.add("rId7", REL_WORKSHEET, "b.xml");
        rels.add("foo", REL_WORKSHEET, "c.xml");
        let ids: Vec<String> = (0..3)
            .map(|i| rels.add_external(REL_HYPERLINK, format!("https://e.com/{i}")))
            .collect();
        assert_eq!(ids, ["rId8", "rId9", "rId10"]);
        assert_eq!(rels.next_r_id(), 11);
        // retain recomputes the next free number from what is left.
        rels.retain(|r| r.id != "rId9" && r.id != "rId10");
        assert_eq!(rels.next_r_id(), 9);
        rels.retain(|r| r.id == "foo");
        assert_eq!(rels.next_r_id(), 1);
        assert_eq!(Relationships::default().next_r_id(), 1);
    }

    #[test]
    fn test_every_rel_type_constant_is_borrowed_on_parse() {
        // Named one by one, not through KNOWN_REL_TYPES, so a constant
        // missing from that table fails here.
        let named = [
            REL_OFFICE_DOCUMENT,
            REL_WORKSHEET,
            REL_SHARED_STRINGS,
            REL_STYLES,
            REL_COMMENTS,
            REL_TABLE,
            REL_CHART,
            REL_CHART_USER_SHAPES,
            REL_CHART_STYLE,
            REL_CHART_COLOR_STYLE,
            REL_DRAWING,
            REL_HYPERLINK,
            REL_PIVOT_TABLE,
            REL_PIVOT_CACHE_DEF,
            REL_PIVOT_CACHE_REC,
            REL_THREADED_COMMENTS,
            REL_PERSONS,
            REL_SLICER,
            REL_SLICER_CACHE,
            REL_TIMELINE,
            REL_TIMELINE_CACHE,
            REL_SHEET_METADATA,
            REL_THEME,
            REL_EXTERNAL_LINK,
            REL_CALC_CHAIN,
            REL_CORE_PROPERTIES,
            REL_EXTENDED_PROPERTIES,
            REL_CUSTOM_PROPERTIES,
            REL_CUSTOM_XML,
        ];
        assert_eq!(named.len(), KNOWN_REL_TYPES.len());
        for known in named {
            let xml = format!(
                r#"<Relationships xmlns="{RELS_NS}"><Relationship Id="rId1" Type="{known}" Target="x.xml"/></Relationships>"#
            );
            let rels = Relationships::parse(xml.as_bytes()).unwrap();
            let rel = rels.get_by_id("rId1").unwrap();
            assert!(matches!(rel.rel_type, Cow::Borrowed(_)), "{known}");
            assert_eq!(rel.rel_type, known);
        }
    }

    #[test]
    fn test_ordinary_parsed_part_id_sequence_is_pinned() {
        let xml = format!(
            r#"<Relationships xmlns="{RELS_NS}"><Relationship Id="rId1" Type="{REL_WORKSHEET}" Target="a.xml"/><Relationship Id="rId2" Type="{REL_WORKSHEET}" Target="b.xml"/><Relationship Id="rId3" Type="{REL_WORKSHEET}" Target="c.xml"/></Relationships>"#
        );
        let mut rels = Relationships::parse(xml.as_bytes()).unwrap();
        let ids: Vec<String> = (0..3)
            .map(|_| rels.add_external(REL_HYPERLINK, "https://a.example/"))
            .collect();
        assert_eq!(ids, ["rId4", "rId5", "rId6"]);
    }

    #[test]
    fn test_saturated_allocation_is_amortized_constant_time() {
        use std::collections::HashSet;
        use std::fmt::Write as _;
        let mut body = format!(
            r#"<Relationship Id="rId{}" Type="{REL_WORKSHEET}" Target="m.xml"/>"#,
            u64::MAX
        );
        for n in 1..=50_000u32 {
            write!(
                body,
                r#"<Relationship Id="rId{n}" Type="{REL_WORKSHEET}" Target="s.xml"/>"#
            )
            .unwrap();
        }
        let xml = format!(r#"<Relationships xmlns="{RELS_NS}">{body}</Relationships>"#);
        let mut rels = Relationships::parse(xml.as_bytes()).unwrap();
        let parsed: HashSet<String> = rels.iter().map(|r| r.id.clone()).collect();
        // Generous bound: the rescanning allocator needs ~3.75e9 probes (minutes).
        let start = std::time::Instant::now();
        let mut issued = HashSet::new();
        for _ in 0..50_000 {
            let id = rels.add_external(REL_HYPERLINK, "https://a.example/");
            assert!(!parsed.contains(&id), "{id} reissued");
            assert!(issued.insert(id), "duplicate id");
        }
        assert!(
            start.elapsed() < std::time::Duration::from_secs(10),
            "{:?}",
            start.elapsed()
        );
        assert_eq!(rels.len(), 100_001);
        assert!(issued.contains("rId50001"));
    }

    #[test]
    fn test_next_r_id_does_not_wrap_past_u32_max() {
        let mut rels = Relationships::new();
        rels.add("rId4294967294", REL_COMMENTS, "x");
        assert_eq!(
            rels.add_external(REL_HYPERLINK, "https://a.example/"),
            "rId4294967295"
        );
        assert_eq!(
            rels.add_external(REL_HYPERLINK, "https://b.example/"),
            "rId4294967296"
        );
        rels.retain(|_| true);
        assert_eq!(rels.next_r_id(), 4_294_967_297);
    }

    #[test]
    fn test_add_external_never_duplicates_at_u64_max() {
        let mut rels = Relationships::new();
        rels.add("rId18446744073709551614", REL_COMMENTS, "x");
        let a = rels.add_external(REL_HYPERLINK, "https://a.example/");
        let b = rels.add_external(REL_HYPERLINK, "https://b.example/");
        assert_ne!(a, b);
        assert_eq!(rels.len(), 3);
    }

    #[test]
    fn test_add_external_allocates_next_id() {
        let mut rels = Relationships::new();
        rels.add("rId3", REL_COMMENTS, "../comments1.xml");
        let id = rels.add_external(REL_HYPERLINK, "mailto:hi@example.com");
        assert_eq!(id, "rId4");
        let rel = rels.get_by_id("rId4").unwrap();
        assert_eq!(rel.target, "mailto:hi@example.com");
        assert_eq!(rel.target_mode, TargetMode::External);
        assert_eq!(rels.len(), 2);
        assert!(!rels.is_empty());
    }

    #[test]
    fn test_retain_removes_and_reindexes() {
        let mut rels = Relationships::new();
        rels.add("rId1", REL_HYPERLINK, "https://a.example/");
        rels.add("rId2", REL_COMMENTS, "../comments1.xml");
        rels.add("rId3", REL_HYPERLINK, "https://b.example/");
        rels.retain(|r| r.rel_type != REL_HYPERLINK);
        assert_eq!(rels.len(), 1);
        assert!(rels.get_by_id("rId1").is_none());
        assert!(rels.get_by_id("rId3").is_none());
        assert_eq!(rels.get_by_id("rId2").unwrap().target, "../comments1.xml");
    }

    #[test]
    fn test_parse_mismatched_end_tag_is_xml_parse_error() {
        let res = Relationships::parse(b"<Relationships><a></b></Relationships>");
        assert!(matches!(res, Err(ModernXlsxError::XmlParse(_))), "{res:?}");
    }

    #[test]
    fn test_parse_bad_attribute_entity_is_error() {
        let xml = br#"<Relationships><Relationship Id="rId1" Type="t" Target="a&bogus;b"/></Relationships>"#;
        let res = Relationships::parse(xml);
        assert!(matches!(res, Err(ModernXlsxError::XmlParse(_))), "{res:?}");
    }
}
