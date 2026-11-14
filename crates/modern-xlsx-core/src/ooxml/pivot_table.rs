//! Pivot tables, pivot cache definitions and pivot cache records
//! (`pivotTable*.xml`, `pivotCacheDefinition*.xml`, `pivotCacheRecords*.xml`).
//!
//! Every attribute is read through `xml::attr_str` (entities and character
//! references decoded) and written through `quick-xml`'s attribute escaping,
//! so a name such as `Q&A` survives a round trip. A numeric attribute that
//! does not parse is an error, never a silent default.
//!
//! The model is partial. The `showAll`, `sortType` and `numFmtId` attributes
//! of `<pivotField>`, the `rowItems` and `colItems` layout caches and
//! `pivotTableStyleInfo` are not modelled, nor the `rowPageCount` and
//! `colPageCount` of `<location>`: the parser ignores them and the writer
//! does not emit them, so they are absent after a round trip. A cache source
//! keeps the `ref`, `name`, `sheet` and `r:id` of its `<worksheetSource>` as
//! read, and a source that is not a worksheet keeps its whole `<cacheSource>`
//! element ([`CacheSource`]). The relationships of a cache definition's
//! `.rels` other than the one to its records (the `externalLinkPath` an
//! `r:id` in the source names) are kept on the definition
//! ([`PivotCacheDefinitionData::relationships`]).
//!
//! The writer does not compute a table's layout (its `rowItems`, `colItems`
//! and cell values). For a table built in code, whose row, column or page
//! fields have no items, it writes the items Excel needs to refresh the
//! table ([`PivotTableData::to_xml_with_cache`]) and `refreshOnLoad="1"` on
//! the cache the table uses, so Excel lays the table out when it opens the
//! file.

use core::hint::cold_path;
use std::borrow::Cow;
use std::collections::HashSet;
use std::fmt::Write as _;

use log::warn;
use quick_xml::events::attributes::Attribute;
use quick_xml::events::{BytesEnd, BytesStart, Event};
use quick_xml::name::ResolveResult;
use quick_xml::{NsReader, Reader, Writer};
use serde::{Deserialize, Serialize};

use super::SPREADSHEET_NS;
use super::relationships::{
    REL_PIVOT_CACHE_REC, Relationship, Relationships, TargetMode, make_rid,
};
use super::xml::{
    REL_NS, RootNamespaces, XML_S, XsdBuf, attr_str, attrs, fmt_xsd_double, parse_attr,
    push_general_ref, root_namespaces, self_contained_fragment, xml_attr, xsd_bool,
};
use crate::{ModernXlsxError, Result};

use crate::ooxml::{write_xml_decl, xml_write_err};

/// Axis position of a pivot field.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum PivotAxis {
    AxisRow,
    AxisCol,
    AxisPage,
    AxisValues,
}

impl PivotAxis {
    #[inline]
    pub fn from_xml(s: &str) -> Option<Self> {
        match s {
            "axisRow" => Some(Self::AxisRow),
            "axisCol" => Some(Self::AxisCol),
            "axisPage" => Some(Self::AxisPage),
            "axisValues" => Some(Self::AxisValues),
            _ => None,
        }
    }

    #[inline]
    pub fn xml_val(self) -> &'static str {
        match self {
            Self::AxisRow => "axisRow",
            Self::AxisCol => "axisCol",
            Self::AxisPage => "axisPage",
            Self::AxisValues => "axisValues",
        }
    }
}

/// Subtotal aggregation function for pivot data fields.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum SubtotalFunction {
    #[default]
    Sum,
    Count,
    Average,
    Max,
    Min,
    Product,
    CountNums,
    StdDev,
    #[serde(rename = "stdDevP")]
    StdDevP,
    Var,
    #[serde(rename = "varP")]
    VarP,
}

impl SubtotalFunction {
    #[inline]
    pub fn from_xml(s: &str) -> Option<Self> {
        match s {
            "sum" => Some(Self::Sum),
            "count" => Some(Self::Count),
            "average" => Some(Self::Average),
            "max" => Some(Self::Max),
            "min" => Some(Self::Min),
            "product" => Some(Self::Product),
            "countNums" => Some(Self::CountNums),
            "stdDev" => Some(Self::StdDev),
            "stdDevP" => Some(Self::StdDevP),
            "var" => Some(Self::Var),
            "varP" => Some(Self::VarP),
            _ => None,
        }
    }

    #[inline]
    pub fn xml_val(self) -> &'static str {
        match self {
            Self::Sum => "sum",
            Self::Count => "count",
            Self::Average => "average",
            Self::Max => "max",
            Self::Min => "min",
            Self::Product => "product",
            Self::CountNums => "countNums",
            Self::StdDev => "stdDev",
            Self::StdDevP => "stdDevP",
            Self::Var => "var",
            Self::VarP => "varP",
        }
    }

    /// The `<pivotField>` attribute that switches this function on as a
    /// subtotal (`sumSubtotal`, `countASubtotal`, ...). Note the spec's own
    /// naming: `countASubtotal` is Excel's Count (non-empty cells) and
    /// `countSubtotal` is Count Numbers.
    #[inline]
    fn subtotal_attr(self) -> &'static str {
        match self {
            Self::Sum => "sumSubtotal",
            Self::Count => "countASubtotal",
            Self::Average => "avgSubtotal",
            Self::Max => "maxSubtotal",
            Self::Min => "minSubtotal",
            Self::Product => "productSubtotal",
            Self::CountNums => "countSubtotal",
            Self::StdDev => "stdDevSubtotal",
            Self::StdDevP => "stdDevPSubtotal",
            Self::Var => "varSubtotal",
            Self::VarP => "varPSubtotal",
        }
    }

    /// Every function, in the order Excel writes a field's subtotal items.
    const ALL: [Self; 11] = [
        Self::Sum,
        Self::Count,
        Self::Average,
        Self::Max,
        Self::Min,
        Self::Product,
        Self::CountNums,
        Self::StdDev,
        Self::StdDevP,
        Self::Var,
        Self::VarP,
    ];

    /// The `t` of the `<item>` that shows this function as a field's
    /// subtotal, with the naming of [`Self::subtotal_attr`]: `countA` is
    /// Count and `count` is Count Numbers.
    #[inline]
    fn item_type(self) -> &'static str {
        match self {
            Self::Sum => "sum",
            Self::Count => "countA",
            Self::Average => "avg",
            Self::Max => "max",
            Self::Min => "min",
            Self::Product => "product",
            Self::CountNums => "count",
            Self::StdDev => "stdDev",
            Self::StdDevP => "stdDevP",
            Self::Var => "var",
            Self::VarP => "varP",
        }
    }

    /// Inverse of [`Self::subtotal_attr`].
    #[inline]
    fn from_subtotal_attr(s: &str) -> Option<Self> {
        match s {
            "sumSubtotal" => Some(Self::Sum),
            "countASubtotal" => Some(Self::Count),
            "avgSubtotal" => Some(Self::Average),
            "maxSubtotal" => Some(Self::Max),
            "minSubtotal" => Some(Self::Min),
            "productSubtotal" => Some(Self::Product),
            "countSubtotal" => Some(Self::CountNums),
            "stdDevSubtotal" => Some(Self::StdDev),
            "stdDevPSubtotal" => Some(Self::StdDevP),
            "varSubtotal" => Some(Self::Var),
            "varPSubtotal" => Some(Self::VarP),
            _ => None,
        }
    }
}

/// Top-level pivot table definition.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct PivotTableData {
    pub name: String,
    /// `dataCaption`; `None` is written as `Values`.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub data_caption: Option<String>,
    pub location: PivotLocation,
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub pivot_fields: Vec<PivotFieldData>,
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub row_fields: Vec<PivotFieldRef>,
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub col_fields: Vec<PivotFieldRef>,
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub data_fields: Vec<PivotDataFieldData>,
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub page_fields: Vec<PivotPageFieldData>,
    /// The cache this table summarises: a
    /// [`PivotCacheDefinitionData::cache_id`], or the cache's index in
    /// `pivotCaches` when that is `None`.
    pub cache_id: u32,
}

/// The location reference for a pivot table within a worksheet.
///
/// An unset offset is written with the value Excel gives a table of this
/// shape: `firstHeaderRow` 1; `firstDataRow` 2 with column fields, else 1;
/// `firstDataCol` 1 with row fields, else 0.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct PivotLocation {
    #[serde(rename = "ref")]
    pub ref_range: String,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub first_header_row: Option<u32>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub first_data_row: Option<u32>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub first_data_col: Option<u32>,
}

/// A single field definition within the pivot table.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct PivotFieldData {
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub axis: Option<PivotAxis>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub name: Option<String>,
    /// The field's `<item>`s. Empty on a row, column or page field means a
    /// field built in code: the writer then writes the items Excel needs
    /// (see [`PivotTableData::to_xml_with_cache`]).
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub items: Vec<PivotItem>,
    /// The functions whose `*Subtotal="1"` attribute is set.
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub subtotals: Vec<SubtotalFunction>,
    /// `compact`; the schema default is `true`, and it is always serialised.
    #[serde(default = "crate::ooxml::default_true")]
    pub compact: bool,
    /// `outline`; the schema default is `true`, and it is always serialised.
    #[serde(default = "crate::ooxml::default_true")]
    pub outline: bool,
    /// `dataField="1"`: the field also feeds the value area.
    #[serde(default, skip_serializing_if = "crate::ooxml::is_false")]
    pub data_field: bool,
    /// `defaultSubtotal`; `None` leaves the attribute out (schema default `true`).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub default_subtotal: Option<bool>,
}

/// An item within a pivot field.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct PivotItem {
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub t: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub x: Option<u32>,
}

/// A reference to a field by index, used in row/col field lists; `-2` is the
/// values field, which holds the data fields.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct PivotFieldRef {
    pub x: i32,
}

/// A data field (value area) definition in the pivot table.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct PivotDataFieldData {
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub name: Option<String>,
    pub fld: u32,
    #[serde(default)]
    pub subtotal: SubtotalFunction,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub num_fmt_id: Option<u32>,
}

/// A page field (filter area) definition in the pivot table.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct PivotPageFieldData {
    pub fld: i32,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub item: Option<u32>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub name: Option<String>,
}

/// Pivot cache definition from pivotCacheDefinition{n}.xml.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct PivotCacheDefinitionData {
    /// The `cacheId` workbook.xml `<pivotCaches>` lists this cache under, which
    /// pivot tables name in their own `cacheId`. `None` (a cache built in code,
    /// or a file without that list) writes the cache's 0-based index in
    /// `pivotCaches`.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cache_id: Option<u32>,
    pub source: CacheSource,
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub fields: Vec<CacheFieldData>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub record_count: Option<u32>,
    /// The relationships of the definition's `.rels` other than the one to
    /// its records, such as the `externalLinkPath` Excel writes for a source
    /// in another workbook, which the source names by `r:id`. Opaque: written
    /// back with their ids; the records relationship takes an id none of them
    /// uses (`rId1` when it is free). Skipped with a warning: one of the
    /// `pivotCacheRecords` type, one whose id is empty, does not start with an
    /// ASCII letter or `_`, or holds a character other than an ASCII letter,
    /// digit, `_`, `-` or `.` (a source `r:id` naming such an id is dropped),
    /// and a second one with the same id (the first is kept).
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub relationships: Vec<CacheRelationship>,
}

/// A relationship of a pivot cache definition other than the one to its
/// records, kept as read (OPC `<Relationship>`).
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct CacheRelationship {
    /// `Id`, which the definition's source names by `r:id`.
    pub id: String,
    /// `Type`, the relationship type URI.
    #[serde(rename = "type")]
    pub rel_type: String,
    /// `Target`: a file path or URI when external, else a part name relative
    /// to the definition.
    pub target: String,
    /// `TargetMode`: `External`, or absent for a part inside the package. A
    /// read gives `External`; the writer reads the value ignoring ASCII
    /// letter case and writes `External`, and any other value as absent.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub target_mode: Option<String>,
}

impl CacheRelationship {
    /// The relationship `rel` of a definition's `.rels`, as kept.
    pub(crate) fn from_relationship(rel: &Relationship) -> Self {
        Self {
            id: rel.id.clone(),
            rel_type: rel.rel_type.clone().into_owned(),
            target: rel.target.clone().into_owned(),
            target_mode: (rel.target_mode == TargetMode::External).then(|| "External".to_owned()),
        }
    }

    /// Why the writer does not write this relationship, or `None` when it
    /// does: one of the `pivotCacheRecords` type (the writer writes the
    /// records relationship itself), or one whose id is empty, does not start
    /// with an ASCII letter or `_`, or holds a character other than an ASCII
    /// letter, digit, `_`, `-` or `.`. A read never gives the first (the
    /// reader keeps the records relationship out of the list); an id outside
    /// the rule can come from a file another tool wrote, and is skipped the
    /// same way.
    pub(crate) fn skip_reason(&self) -> Option<&'static str> {
        if self.rel_type == REL_PIVOT_CACHE_REC {
            Some("is of the pivotCacheRecords type, which the writer writes itself")
        } else if self.id.is_empty()
            || !matches!(self.id.as_bytes()[0], b'A'..=b'Z' | b'a'..=b'z' | b'_')
            || !self
                .id
                .bytes()
                .all(|b| b.is_ascii_alphanumeric() || matches!(b, b'_' | b'-' | b'.'))
        {
            Some(
                "has an id that is empty, does not start with an ASCII letter or '_', or holds a character other than an ASCII letter, digit, '_', '-' or '.'",
            )
        } else {
            None
        }
    }
}

/// The source of a pivot cache: the `<worksheetSource>` (ECMA-376 Part 1,
/// 18.10.1.96) of a `<cacheSource type="worksheet">`, or, for any other
/// source type, the whole `<cacheSource>` element in [`Self::raw_xml`]. The
/// writer emits each `<worksheetSource>` attribute only when it has a value.
///
/// An `r:id` in the source (Excel writes one for a source in another
/// workbook) names one of the definition's
/// [`PivotCacheDefinitionData::relationships`]. The writer writes it only
/// when that relationship is kept: Excel refuses a file whose `r:id` names
/// no relationship, so a dangling one is dropped with a warning.
#[derive(Debug, Clone, Default, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct CacheSource {
    /// `ref`: the source range; empty when the file has none.
    #[serde(rename = "ref")]
    pub ref_range: String,
    /// `sheet`: the sheet of `ref` (or of a sheet-scoped `name`); empty when
    /// the file has none.
    pub sheet: String,
    /// `name`: a table or defined name. Excel writes it, with no `ref` or
    /// `sheet`, for a pivot table made from a table or a named range.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub name: Option<String>,
    /// `r:id`: the relationship naming the workbook that holds the source,
    /// for a source in another workbook.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub r_id: Option<String>,
    /// A source that is not a worksheet (`type` `consolidation`, `external`
    /// or `scenario`): the whole `<cacheSource>` element as read, opaque and
    /// written as read in place of the fields above, except that an `r:id`
    /// naming no kept relationship is dropped (see the type docs).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub raw_xml: Option<String>,
}

/// A field within the pivot cache (column metadata + shared items).
#[derive(Debug, Clone, Default, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct CacheFieldData {
    pub name: String,
    /// `numFmtId` of the cache field.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub num_fmt_id: Option<u32>,
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub shared_items: Vec<CacheValue>,
    /// The attributes of `<sharedItems>` (type flags and value ranges). The
    /// writer emits the element when this is `Some` or there are items.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub shared_items_attrs: Option<SharedItemsAttrs>,
}

/// The attributes of a cache field's `<sharedItems>` element, each present
/// only when the file had it. `count` is derived from the items on write.
#[derive(Debug, Clone, Default, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct SharedItemsAttrs {
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub contains_semi_mixed_types: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub contains_non_date: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub contains_date: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub contains_string: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub contains_blank: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub contains_mixed_types: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub contains_number: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub contains_integer: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub min_value: Option<f64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub max_value: Option<f64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub min_date: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub max_date: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub long_text: Option<bool>,
}

/// A cached value in a pivot cache record or shared items list.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", tag = "type")]
pub enum CacheValue {
    Number {
        v: f64,
    },
    String {
        v: String,
    },
    Boolean {
        v: bool,
    },
    DateTime {
        v: String,
    },
    Missing,
    Error {
        v: String,
    },
    /// Shared item index reference (used in records — `<x v="0"/>`).
    Index {
        v: u32,
    },
}

/// Pivot cache records from pivotCacheRecords{n}.xml.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct PivotCacheRecordsData {
    pub records: Vec<Vec<CacheValue>>,
}

// ---------------------------------------------------------------------------
// Shared helpers
// ---------------------------------------------------------------------------

/// Push `(key, value)` onto `e`, the value with illegal XML characters
/// removed; quick-xml escapes it exactly once.
#[inline]
fn push_text(e: &mut BytesStart<'_>, key: &str, value: &str) {
    e.push_attribute(xml_attr(key, value));
}

/// Push an `xsd:boolean` attribute as `1` / `0`.
#[inline]
fn push_flag(e: &mut BytesStart<'_>, key: &str, value: bool) {
    e.push_attribute((key, if value { "1" } else { "0" }));
}

/// The `v` attribute of `e`, if any.
fn find_v<'a>(e: &'a BytesStart<'_>) -> Result<Option<Attribute<'a>>> {
    for attr in attrs(e) {
        let attr = attr?;
        if attr.key.local_name().as_ref() == "v" {
            return Ok(Some(attr));
        }
    }
    Ok(None)
}

// ---------------------------------------------------------------------------
// Pivot cache definition — parser & writer
// ---------------------------------------------------------------------------

impl PivotCacheDefinitionData {
    /// Parse a `pivotCacheDefinition*.xml` file from raw XML bytes.
    ///
    /// # Errors
    ///
    /// `XmlParse` for malformed XML, a bad entity, or a numeric attribute that
    /// does not parse.
    pub fn parse(data: &[u8]) -> Result<Self> {
        let mut reader = Reader::from_reader(data);
        reader.config_mut().trim_text(true);
        // The offset in `data` of what the reader has not consumed. A verbatim
        // span is cut at these offsets, never at buffer_position(), which does
        // not count a UTF-8 BOM quick-xml skips: no prefix can shift it.
        let offset = |reader: &Reader<&[u8]>| data.len() - reader.get_ref().len();

        let mut buf = Vec::with_capacity(512);

        let mut st = CacheDefParser::new();

        loop {
            let ev_start = offset(&reader);
            match reader.read_event_into(&mut buf) {
                // A source that is not a worksheet is kept whole (CacheSource::raw_xml).
                Ok(Event::Start(ref e)) if is_raw_source(e)? => {
                    reader
                        .read_to_end_into(e.name(), &mut Vec::new())
                        .map_err(|err| {
                            cold_path();
                            ModernXlsxError::XmlParse(format!(
                                "error parsing pivot cache definition XML: {err}"
                            ))
                        })?;
                    st.keep_raw_source(data, ev_start, offset(&reader))?;
                }
                Ok(Event::Empty(ref e)) if is_raw_source(e)? => {
                    st.keep_raw_source(data, ev_start, offset(&reader))?;
                }
                Ok(Event::Start(ref e)) => {
                    st.handle_open(e, true)?;
                }
                Ok(Event::Empty(ref e)) => {
                    st.handle_open(e, false)?;
                }
                Ok(Event::End(ref e)) => st.handle_close(e.local_name().as_ref()),
                Ok(Event::Eof) => break,
                Ok(_) => {}
                Err(e) => {
                    cold_path();
                    return Err(ModernXlsxError::XmlParse(format!(
                        "error parsing pivot cache definition XML: {e}"
                    )));
                }
            }
            buf.clear();
        }
        st.flush_raw_source();

        // The cacheId lives in workbook.xml; the reader sets it.
        Ok(Self {
            cache_id: None,
            source: st.source,
            fields: st.fields,
            record_count: st.record_count,
            relationships: Vec::new(),
        })
    }

    /// Serialize this pivot cache definition to valid OOXML XML bytes, with no
    /// `r:id` (see [`Self::to_xml_with_records`]).
    ///
    /// # Errors
    ///
    /// `XmlWrite` when the writer fails or a number is not finite.
    pub fn to_xml(&self) -> Result<Vec<u8>> {
        self.to_xml_with_records(None)
    }

    /// The definition's [`Self::relationships`] as a `.rels` collection, each
    /// under its own id, external when its `target_mode` is `External` in
    /// any ASCII letter case. Skipped with a warning: a repeated id (a
    /// `.rels` holds each id once) and a relationship the writer does not
    /// keep ([`CacheRelationship::skip_reason`]).
    pub(crate) fn kept_relationships(&self) -> Relationships {
        let mut rels = Relationships::new();
        for rel in &self.relationships {
            if let Some(reason) = rel.skip_reason() {
                warn!(
                    "pivot cache relationship {:?} {reason}; skipped",
                    rel.id.chars().take(64).collect::<String>()
                );
                continue;
            }
            if rels.get_by_id(&rel.id).is_some() {
                warn!(
                    "pivot cache relationship {} repeated; the first is kept",
                    rel.id
                );
                continue;
            }
            rels.push(Relationship {
                id: rel.id.clone(),
                rel_type: rel.rel_type.clone().into(),
                target: rel.target.clone().into(),
                target_mode: if rel
                    .target_mode
                    .as_deref()
                    .is_some_and(|mode| mode.eq_ignore_ascii_case("External"))
                {
                    TargetMode::External
                } else {
                    TargetMode::Internal
                },
            });
        }
        rels
    }

    /// The id of the relationship to the records part: `rId1` when no kept
    /// relationship uses it, else the next id none of them uses.
    pub(crate) fn records_r_id(&self) -> String {
        let rels = self.kept_relationships();
        if rels.get_by_id("rId1").is_none() {
            "rId1".to_owned()
        } else {
            make_rid(rels.next_r_id())
        }
    }

    /// Serialize this pivot cache definition to valid OOXML XML bytes. With
    /// `records_r_id`, the root carries it as `r:id` (ECMA-376 Part 1
    /// `pivotCacheDefinition@r:id`), the definition's relationship to its
    /// pivot cache records part.
    ///
    /// # Errors
    ///
    /// `XmlWrite` when the writer fails or a number is not finite.
    pub fn to_xml_with_records(&self, records_r_id: Option<&str>) -> Result<Vec<u8>> {
        self.write_xml(records_r_id, false)
    }

    /// [`Self::to_xml_with_records`], with `refreshOnLoad="1"` on the root
    /// when `refresh_on_load` is set: Excel then refreshes the cache when it
    /// opens the file, which lays out its pivot tables and shows their
    /// values. The writer sets it for a cache that a pivot table built in
    /// code uses ([`PivotTableData::has_axis_field_without_items`]); the
    /// attribute is not modelled, so a read drops it.
    ///
    /// # Errors
    ///
    /// `XmlWrite` when the writer fails or a number is not finite.
    pub(crate) fn write_xml(
        &self,
        records_r_id: Option<&str>,
        refresh_on_load: bool,
    ) -> Result<Vec<u8>> {
        let mut buf: Vec<u8> = Vec::with_capacity(512 + self.fields.len() * 96);
        let mut writer = Writer::new(&mut buf);
        let mut ibuf = itoa::Buffer::new();
        let mut xbuf = XsdBuf::new();

        // XML declaration.
        write_xml_decl(&mut writer)?;

        // <pivotCacheDefinition xmlns="..." xmlns:r="..." r:id="..." refreshOnLoad="1" recordCount="...">
        let mut root = BytesStart::new("pivotCacheDefinition");
        root.push_attribute(("xmlns", SPREADSHEET_NS));
        root.push_attribute(("xmlns:r", REL_NS));
        if let Some(id) = records_r_id {
            root.push_attribute(("r:id", id));
        }
        if refresh_on_load {
            push_flag(&mut root, "refreshOnLoad", true);
        }
        if let Some(rc) = self.record_count {
            root.push_attribute(("recordCount", ibuf.format(rc)));
        }
        writer
            .write_event(Event::Start(root))
            .map_err(xml_write_err)?;

        let src = &self.source;
        if let Some(raw) = &src.raw_xml {
            let kept: HashSet<&str> = self
                .relationships
                .iter()
                .filter(|r| r.skip_reason().is_none())
                .map(|r| r.id.as_str())
                .collect();
            writer
                .get_mut()
                .extend_from_slice(drop_dangling_r_ids(raw, &kept).as_bytes());
        } else {
            // <cacheSource type="worksheet"><worksheetSource .../></cacheSource>,
            // with each attribute only when it has a value.
            let mut cs = BytesStart::new("cacheSource");
            cs.push_attribute(("type", "worksheet"));
            writer
                .write_event(Event::Start(cs))
                .map_err(xml_write_err)?;
            let mut ws = BytesStart::new("worksheetSource");
            let name = src.name.as_deref().unwrap_or_default();
            for (key, value) in [
                ("ref", src.ref_range.as_str()),
                ("name", name),
                ("sheet", src.sheet.as_str()),
            ] {
                if !value.is_empty() {
                    push_text(&mut ws, key, value);
                }
            }
            if let Some(id) = &src.r_id {
                if self
                    .relationships
                    .iter()
                    .any(|r| r.id == *id && r.skip_reason().is_none())
                {
                    push_text(&mut ws, "r:id", id);
                } else {
                    warn!("pivot cache source r:id {id} names no kept relationship; dropped");
                }
            }
            writer
                .write_event(Event::Empty(ws))
                .map_err(xml_write_err)?;
            writer
                .write_event(Event::End(BytesEnd::new("cacheSource")))
                .map_err(xml_write_err)?;
        }

        // <cacheFields count="N">, required (minOccurs=1) even with no field.
        if self.fields.is_empty() {
            let mut cf = BytesStart::new("cacheFields");
            cf.push_attribute(("count", "0"));
            writer
                .write_event(Event::Empty(cf))
                .map_err(xml_write_err)?;
        } else {
            let mut cf = BytesStart::new("cacheFields");
            cf.push_attribute(("count", ibuf.format(self.fields.len())));
            writer
                .write_event(Event::Start(cf))
                .map_err(xml_write_err)?;

            for field in &self.fields {
                let mut fe = BytesStart::new("cacheField");
                push_text(&mut fe, "name", &field.name);
                if let Some(id) = field.num_fmt_id {
                    fe.push_attribute(("numFmtId", ibuf.format(id)));
                }

                if field.shared_items.is_empty() && field.shared_items_attrs.is_none() {
                    writer
                        .write_event(Event::Empty(fe))
                        .map_err(xml_write_err)?;
                    continue;
                }
                writer
                    .write_event(Event::Start(fe))
                    .map_err(xml_write_err)?;

                // <sharedItems ...> (empty when there are no items)
                let mut si = BytesStart::new("sharedItems");
                if let Some(attrs) = &field.shared_items_attrs {
                    push_shared_items_attrs(&mut si, attrs, &mut xbuf)?;
                }
                if !field.shared_items.is_empty() {
                    si.push_attribute(("count", ibuf.format(field.shared_items.len())));
                }
                if field.shared_items.is_empty() {
                    writer
                        .write_event(Event::Empty(si))
                        .map_err(xml_write_err)?;
                } else {
                    writer
                        .write_event(Event::Start(si))
                        .map_err(xml_write_err)?;
                    for item in &field.shared_items {
                        write_cache_value(&mut writer, item, &mut ibuf, &mut xbuf)?;
                    }
                    writer
                        .write_event(Event::End(BytesEnd::new("sharedItems")))
                        .map_err(xml_write_err)?;
                }

                // </cacheField>
                writer
                    .write_event(Event::End(BytesEnd::new("cacheField")))
                    .map_err(xml_write_err)?;
            }

            // </cacheFields>
            writer
                .write_event(Event::End(BytesEnd::new("cacheFields")))
                .map_err(xml_write_err)?;
        }

        // </pivotCacheDefinition>
        writer
            .write_event(Event::End(BytesEnd::new("pivotCacheDefinition")))
            .map_err(xml_write_err)?;

        Ok(buf)
    }
}

/// Push every `Some` attribute of `attrs` onto a `<sharedItems>` element.
fn push_shared_items_attrs(
    e: &mut BytesStart<'_>,
    attrs: &SharedItemsAttrs,
    xbuf: &mut XsdBuf,
) -> Result<()> {
    let flags = [
        ("containsSemiMixedTypes", attrs.contains_semi_mixed_types),
        ("containsNonDate", attrs.contains_non_date),
        ("containsDate", attrs.contains_date),
        ("containsString", attrs.contains_string),
        ("containsBlank", attrs.contains_blank),
        ("containsMixedTypes", attrs.contains_mixed_types),
        ("containsNumber", attrs.contains_number),
        ("containsInteger", attrs.contains_integer),
    ];
    for (key, value) in flags {
        if let Some(v) = value {
            push_flag(e, key, v);
        }
    }
    if let Some(v) = attrs.min_value {
        e.push_attribute((
            "minValue",
            fmt_xsd_double(v, xbuf, "pivot sharedItems minValue")?,
        ));
    }
    if let Some(v) = attrs.max_value {
        e.push_attribute((
            "maxValue",
            fmt_xsd_double(v, xbuf, "pivot sharedItems maxValue")?,
        ));
    }
    if let Some(v) = &attrs.min_date {
        push_text(e, "minDate", v);
    }
    if let Some(v) = &attrs.max_date {
        push_text(e, "maxDate", v);
    }
    if let Some(v) = attrs.long_text {
        push_flag(e, "longText", v);
    }
    Ok(())
}

/// Read the attributes of a `<sharedItems>` element.
fn parse_shared_items_attrs(e: &BytesStart<'_>) -> Result<SharedItemsAttrs> {
    let mut out = SharedItemsAttrs::default();
    for attr in attrs(e) {
        let attr = attr?;
        let key = attr.key.local_name();
        match key.as_ref() {
            "containsSemiMixedTypes" => {
                out.contains_semi_mixed_types = xsd_bool(&attr_str(&attr)?);
            }
            "containsNonDate" => out.contains_non_date = xsd_bool(&attr_str(&attr)?),
            "containsDate" => out.contains_date = xsd_bool(&attr_str(&attr)?),
            "containsString" => out.contains_string = xsd_bool(&attr_str(&attr)?),
            "containsBlank" => out.contains_blank = xsd_bool(&attr_str(&attr)?),
            "containsMixedTypes" => out.contains_mixed_types = xsd_bool(&attr_str(&attr)?),
            "containsNumber" => out.contains_number = xsd_bool(&attr_str(&attr)?),
            "containsInteger" => out.contains_integer = xsd_bool(&attr_str(&attr)?),
            "minValue" => out.min_value = Some(parse_attr("sharedItems", &attr)?),
            "maxValue" => out.max_value = Some(parse_attr("sharedItems", &attr)?),
            "minDate" => out.min_date = Some(attr_str(&attr)?.into_owned()),
            "maxDate" => out.max_date = Some(attr_str(&attr)?.into_owned()),
            "longText" => out.long_text = xsd_bool(&attr_str(&attr)?),
            _ => {}
        }
    }
    Ok(out)
}

/// `raw` (a kept `<cacheSource>`) without each attribute in the
/// relationships namespace (`r:id`) whose value is not in `kept`, so no
/// written `r:id` names a missing relationship, which makes Excel refuse the
/// file; each one dropped is logged as a warning. The span is read in the
/// scope the written root gives it (`r` bound to the relationships
/// namespace). A span that does not parse is returned as it is.
fn drop_dangling_r_ids<'a>(raw: &'a str, kept: &HashSet<&str>) -> Cow<'a, str> {
    let mut reader = NsReader::from_str(raw);
    let scope = BytesStart::from_content(format!("w xmlns:r=\"{REL_NS}\""), 1);
    if reader.resolver_mut().push(&scope).is_err() {
        return Cow::Borrowed(raw);
    }
    let mut out = String::new();
    // The end of the part of `raw` already copied to `out`.
    let mut copied = 0;
    loop {
        let start = raw.len() - reader.get_ref().len();
        let (e, empty) = match reader.read_event() {
            Ok(Event::Start(e)) => (e, false),
            Ok(Event::Empty(e)) => (e, true),
            Ok(Event::Eof) => break,
            Ok(_) => continue,
            Err(_) => return Cow::Borrowed(raw),
        };
        let mut tag = String::new();
        let mut dropped = false;
        for attr in e.attributes() {
            let Ok(attr) = attr else {
                return Cow::Borrowed(raw);
            };
            let (ns, _) = reader.resolver().resolve_attribute(attr.key);
            if matches!(ns, ResolveResult::Bound(ns) if ns.as_ref() == REL_NS) {
                let Ok(id) = attr_str(&attr) else {
                    return Cow::Borrowed(raw);
                };
                if !kept.contains(id.as_ref()) {
                    warn!("pivot cache source r:id {id} names no kept relationship; dropped");
                    dropped = true;
                    continue;
                }
            }
            // The value as read (escaped); only a `"`, legal inside `'...'`,
            // needs escaping inside the double quotes written here. Writing
            // into a String cannot fail.
            let _ = write!(
                tag,
                " {}=\"{}\"",
                attr.key.as_ref(),
                attr.value.replace('"', "&quot;")
            );
        }
        if dropped {
            out.push_str(&raw[copied..start]);
            out.push('<');
            out.push_str(e.name().as_ref());
            out.push_str(&tag);
            out.push_str(if empty { "/>" } else { ">" });
            copied = raw.len() - reader.get_ref().len();
        }
    }
    if copied == 0 {
        return Cow::Borrowed(raw);
    }
    out.push_str(&raw[copied..]);
    Cow::Owned(out)
}

/// Check every reference in `span`, a `<cacheSource>` element kept as read:
/// in an attribute value through [`attr_str`], as any attribute the reader
/// decodes, and in text through [`push_general_ref`], refusing besides a
/// named entity other than the five predefined ones. The span is written
/// back as read, without the DTD that may declare such an entity, so
/// keeping one would write a part that is not well-formed.
///
/// # Errors
///
/// `XmlParse` for an undeclared entity or a malformed character reference.
fn check_kept_refs(span: &str) -> Result<()> {
    let mut reader = Reader::from_str(span);
    // The main reader has already matched the span's tags.
    reader.config_mut().check_end_names = false;
    let mut scratch = String::new();
    loop {
        match reader.read_event() {
            Ok(Event::Start(e) | Event::Empty(e)) => {
                // Only the values are checked here: the attribute list as
                // such was not checked before either.
                for attr in e.attributes().with_checks(false).flatten() {
                    attr_str(&attr)?;
                }
            }
            Ok(Event::GeneralRef(r)) => {
                push_general_ref(&mut scratch, &r)?;
                if !r.is_char_ref() && quick_xml::escape::resolve_xml_entity(&r).is_none() {
                    cold_path();
                    return Err(ModernXlsxError::XmlParse(format!(
                        "pivot cacheSource: undeclared entity &{};",
                        &*r
                    )));
                }
                scratch.clear();
            }
            Ok(Event::Eof) => return Ok(()),
            Ok(_) => {}
            Err(e) => {
                cold_path();
                return Err(ModernXlsxError::XmlParse(format!(
                    "error parsing pivot cache definition XML: {e}"
                )));
            }
        }
    }
}

/// Whether `e` is a `<cacheSource>` whose `type` is not `worksheet` (the
/// schema default when absent), which the model keeps whole.
fn is_raw_source(e: &BytesStart<'_>) -> Result<bool> {
    if e.local_name().as_ref() != "cacheSource" {
        return Ok(false);
    }
    for attr in attrs(e) {
        let attr = attr?;
        if attr.key.local_name().as_ref() == "type" {
            return Ok(attr_str(&attr)? != "worksheet");
        }
    }
    Ok(false)
}

/// Mutable SAX state for [`PivotCacheDefinitionData::parse`] over the part
/// `'d`.
struct CacheDefParser<'d> {
    source: CacheSource,
    fields: Vec<CacheFieldData>,
    record_count: Option<u32>,
    /// The namespace declarations of the root except `r` bound to the
    /// relationships namespace, which a kept `raw_xml` source re-declares
    /// (`xml::self_contained_fragment`); indexed once per root.
    root_ns: RootNamespaces,
    /// The last `<cacheSource>` kept whole under the current root, checked
    /// but not yet made self-contained: a part may repeat it, so it is made
    /// self-contained once ([`Self::flush_raw_source`]), not per occurrence,
    /// which would copy the root's declarations it uses once per occurrence.
    pending_raw: Option<&'d str>,
    in_cache_source: bool,
    in_cache_fields: bool,
    current_cache_field: Option<CacheFieldData>,
    in_shared_items: bool,
}

impl<'d> CacheDefParser<'d> {
    fn new() -> Self {
        Self {
            source: CacheSource {
                ref_range: String::new(),
                sheet: String::new(),
                name: None,
                r_id: None,
                raw_xml: None,
            },
            fields: Vec::new(),
            record_count: None,
            root_ns: RootNamespaces::default(),
            pending_raw: None,
            in_cache_source: false,
            in_cache_fields: false,
            current_cache_field: None,
            in_shared_items: false,
        }
    }

    /// Check the `<cacheSource>` element at `data[start..end]` and hold it as
    /// [`Self::pending_raw`], replacing a source held before: the last one
    /// wins. [`Self::flush_raw_source`] keeps it as [`CacheSource::raw_xml`],
    /// with the root's prefixes it uses declared, except `r` bound to the
    /// relationships namespace: the written root always declares that
    /// binding, so declaring it in the span too would write it twice and
    /// make a second read differ from the first.
    ///
    /// # Errors
    ///
    /// `XmlParse` when the span is out of range, not UTF-8, not one element
    /// (it must start with `<` and end with `>`), or holds a reference that
    /// is not a well-formed character reference or one of the five
    /// predefined entities ([`check_kept_refs`]).
    fn keep_raw_source(&mut self, data: &'d [u8], start: usize, end: usize) -> Result<()> {
        let span = data.get(start..end).ok_or_else(|| {
            cold_path();
            ModernXlsxError::XmlParse(format!(
                "pivot cacheSource at bytes {start}..{end} is out of range"
            ))
        })?;
        let text = std::str::from_utf8(span)
            .map_err(|e| {
                cold_path();
                ModernXlsxError::XmlParse(format!("pivot cacheSource is not UTF-8: {e}"))
            })?
            .trim_start_matches(XML_S);
        if !(text.starts_with('<') && text.ends_with('>')) {
            cold_path();
            return Err(ModernXlsxError::XmlParse(format!(
                "pivot cacheSource at bytes {start}..{end} is not one element"
            )));
        }
        check_kept_refs(text)?;
        self.pending_raw = Some(text);
        Ok(())
    }

    /// Keep the pending `<cacheSource>` span ([`Self::pending_raw`]) as
    /// [`CacheSource::raw_xml`], made self-contained against
    /// [`Self::root_ns`]: once per root, when the next root replaces it and
    /// at the end of the part.
    fn flush_raw_source(&mut self) {
        if let Some(text) = self.pending_raw.take() {
            self.source.raw_xml = Some(self_contained_fragment(text, &self.root_ns));
        }
    }

    /// Handle an opening or self-closing element during SAX parsing.
    fn handle_open(&mut self, e: &BytesStart<'_>, is_start: bool) -> Result<()> {
        let local = e.local_name();
        match local.as_ref() {
            "pivotCacheDefinition" => {
                // The pending span belongs to the root being replaced.
                self.flush_raw_source();
                self.root_ns = root_namespaces(e)?
                    .into_iter()
                    .filter(|(prefix, uri)| !(prefix == "r" && uri == REL_NS))
                    .collect();
                for attr in attrs(e) {
                    let attr = attr?;
                    if attr.key.local_name().as_ref() == "recordCount" {
                        self.record_count = Some(parse_attr("pivotCacheDefinition", &attr)?);
                    }
                }
            }
            "cacheSource" if is_start => {
                self.in_cache_source = true;
            }
            "worksheetSource" if self.in_cache_source => {
                for attr in attrs(e) {
                    let attr = attr?;
                    match attr.key.local_name().as_ref() {
                        "ref" => {
                            self.source.ref_range = attr_str(&attr)?.into_owned();
                        }
                        "sheet" => {
                            self.source.sheet = attr_str(&attr)?.into_owned();
                        }
                        "name" => self.source.name = Some(attr_str(&attr)?.into_owned()),
                        "id" if attr.key.prefix().is_some() => {
                            self.source.r_id = Some(attr_str(&attr)?.into_owned());
                        }
                        _ => {}
                    }
                }
            }
            "cacheFields" if is_start => {
                self.in_cache_fields = true;
            }
            "cacheField" if self.in_cache_fields => {
                let mut field = CacheFieldData::default();
                for attr in attrs(e) {
                    let attr = attr?;
                    match attr.key.local_name().as_ref() {
                        "name" => field.name = attr_str(&attr)?.into_owned(),
                        "numFmtId" => field.num_fmt_id = Some(parse_attr("cacheField", &attr)?),
                        _ => {}
                    }
                }
                if is_start {
                    self.current_cache_field = Some(field);
                } else {
                    // Self-closing <cacheField .../> — push immediately.
                    self.fields.push(field);
                }
            }
            "sharedItems" if let Some(field) = self.current_cache_field.as_mut() => {
                field.shared_items_attrs = Some(parse_shared_items_attrs(e)?);
                self.in_shared_items = is_start;
            }
            "s" | "n" | "b" | "d" | "m" | "e"
                if self.in_shared_items
                    && let Some(field) = self.current_cache_field.as_mut() =>
            {
                field
                    .shared_items
                    .push(parse_cache_value(local.as_ref(), e)?);
            }
            _ => {}
        }
        Ok(())
    }

    /// Handle a closing element during SAX parsing.
    #[inline]
    fn handle_close(&mut self, local: &str) {
        match local {
            "cacheSource" => self.in_cache_source = false,
            "cacheFields" => self.in_cache_fields = false,
            "cacheField" => {
                if let Some(field) = self.current_cache_field.take() {
                    self.fields.push(field);
                }
                self.in_shared_items = false;
            }
            "sharedItems" => {
                self.in_shared_items = false;
                // A `<sharedItems>` holding items and no attribute needs no
                // attribute record; the writer emits it for the items alone.
                if let Some(field) = self.current_cache_field.as_mut()
                    && !field.shared_items.is_empty()
                    && field.shared_items_attrs == Some(SharedItemsAttrs::default())
                {
                    field.shared_items_attrs = None;
                }
            }
            _ => {}
        }
    }
}

// ---------------------------------------------------------------------------
// Pivot cache records — parser & writer
// ---------------------------------------------------------------------------

impl PivotCacheRecordsData {
    /// Parse a `pivotCacheRecords*.xml` file from raw XML bytes.
    ///
    /// # Errors
    ///
    /// `XmlParse` for malformed XML, a bad entity, or a numeric value that
    /// does not parse.
    pub fn parse(data: &[u8]) -> Result<Self> {
        let mut reader = Reader::from_reader(data);
        reader.config_mut().trim_text(true);

        let mut buf = Vec::with_capacity(512);
        let mut skip_buf = Vec::new();
        let mut records: Vec<Vec<CacheValue>> = Vec::new();
        let mut in_record = false;
        let mut current_record: Vec<CacheValue> = Vec::new();

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) => {
                    let local = e.local_name();
                    match local.as_ref() {
                        "r" => {
                            in_record = true;
                            current_record.clear();
                        }
                        // A value in Start/End form: read it, then skip its
                        // member-property children (`tpls`, `x`).
                        "x" | "n" | "s" | "b" | "d" | "m" | "e" if in_record => {
                            current_record.push(parse_cache_value(local.as_ref(), e)?);
                            reader
                                .read_to_end_into(e.name(), &mut skip_buf)
                                .map_err(|err| {
                                    ModernXlsxError::XmlParse(format!(
                                        "error parsing pivot cache records XML: {err}"
                                    ))
                                })?;
                            skip_buf.clear();
                        }
                        _ => {}
                    }
                }
                // An empty record `<r/>` still holds a row.
                Ok(Event::Empty(ref e)) if !in_record && e.local_name().as_ref() == "r" => {
                    records.push(Vec::new());
                }
                Ok(Event::Empty(ref e)) if in_record => {
                    let local = e.local_name();
                    match local.as_ref() {
                        "x" | "n" | "s" | "b" | "d" | "m" | "e" => {
                            current_record.push(parse_cache_value(local.as_ref(), e)?);
                        }
                        _ => {}
                    }
                }
                Ok(Event::End(ref e)) => {
                    if e.local_name().as_ref() == "r" {
                        records.push(std::mem::take(&mut current_record));
                        in_record = false;
                    }
                }
                Ok(Event::Eof) => break,
                Ok(_) => {}
                Err(e) => {
                    cold_path();
                    return Err(ModernXlsxError::XmlParse(format!(
                        "error parsing pivot cache records XML: {e}"
                    )));
                }
            }
            buf.clear();
        }

        Ok(Self { records })
    }

    /// Serialize these pivot cache records to valid OOXML XML bytes.
    ///
    /// # Errors
    ///
    /// `XmlWrite` when the writer fails or a number is not finite.
    pub fn to_xml(&self) -> Result<Vec<u8>> {
        let mut buf: Vec<u8> = Vec::with_capacity(512 + self.records.len() * 64);
        let mut writer = Writer::new(&mut buf);
        let mut ibuf = itoa::Buffer::new();
        let mut xbuf = XsdBuf::new();

        // XML declaration.
        write_xml_decl(&mut writer)?;

        // <pivotCacheRecords xmlns="..." count="N">
        let mut root = BytesStart::new("pivotCacheRecords");
        root.push_attribute(("xmlns", SPREADSHEET_NS));
        root.push_attribute(("xmlns:r", REL_NS));
        root.push_attribute(("count", ibuf.format(self.records.len())));
        writer
            .write_event(Event::Start(root))
            .map_err(xml_write_err)?;

        for record in &self.records {
            // <r>
            writer
                .write_event(Event::Start(BytesStart::new("r")))
                .map_err(xml_write_err)?;

            for value in record {
                write_cache_value(&mut writer, value, &mut ibuf, &mut xbuf)?;
            }

            // </r>
            writer
                .write_event(Event::End(BytesEnd::new("r")))
                .map_err(xml_write_err)?;
        }

        // </pivotCacheRecords>
        writer
            .write_event(Event::End(BytesEnd::new("pivotCacheRecords")))
            .map_err(xml_write_err)?;

        Ok(buf)
    }
}

// ---------------------------------------------------------------------------
// Shared helpers for parsing / writing CacheValue elements
// ---------------------------------------------------------------------------

/// Parse a single cache value element (`<s>`, `<n>`, `<b>`, `<d>`, `<m>`, `<e>`, `<x>`).
///
/// A missing `v` reads as empty text, `0`, `false` or index `0`. A numeric `v`
/// (`<n>`, `<x>`) that does not parse, or any `v` holding a bad entity, is an
/// error; a `<b>` whose `v` is not an xsd:boolean reads as `false`.
fn parse_cache_value(tag: &str, e: &BytesStart<'_>) -> Result<CacheValue> {
    let v = find_v(e)?;
    let text = |v: Option<Attribute<'_>>| -> Result<String> {
        v.map_or_else(|| Ok(String::new()), |a| attr_str(&a).map(Cow::into_owned))
    };
    Ok(match tag {
        "s" => CacheValue::String { v: text(v)? },
        "n" => CacheValue::Number {
            v: v.map_or(Ok(0.0), |a| parse_attr("n", &a))?,
        },
        "b" => CacheValue::Boolean {
            v: match v {
                Some(a) => xsd_bool(&attr_str(&a)?).unwrap_or(false),
                None => false,
            },
        },
        "d" => CacheValue::DateTime { v: text(v)? },
        "e" => CacheValue::Error { v: text(v)? },
        "x" => CacheValue::Index {
            v: v.map_or(Ok(0), |a| parse_attr("x", &a))?,
        },
        _ => CacheValue::Missing,
    })
}

/// Write a single `CacheValue` as an XML element.
fn write_cache_value(
    writer: &mut Writer<&mut Vec<u8>>,
    value: &CacheValue,
    ibuf: &mut itoa::Buffer,
    xbuf: &mut XsdBuf,
) -> Result<()> {
    let mut elem;
    match value {
        CacheValue::String { v } => {
            elem = BytesStart::new("s");
            push_text(&mut elem, "v", v);
        }
        CacheValue::Number { v } => {
            elem = BytesStart::new("n");
            elem.push_attribute(("v", fmt_xsd_double(*v, xbuf, "pivot cache <n>")?));
        }
        CacheValue::Boolean { v } => {
            elem = BytesStart::new("b");
            push_flag(&mut elem, "v", *v);
        }
        CacheValue::DateTime { v } => {
            elem = BytesStart::new("d");
            push_text(&mut elem, "v", v);
        }
        CacheValue::Missing => elem = BytesStart::new("m"),
        CacheValue::Error { v } => {
            elem = BytesStart::new("e");
            push_text(&mut elem, "v", v);
        }
        CacheValue::Index { v } => {
            elem = BytesStart::new("x");
            elem.push_attribute(("v", ibuf.format(*v)));
        }
    }
    writer
        .write_event(Event::Empty(elem))
        .map_err(xml_write_err)
}

/// The `x` of the values field, the pseudo-field that holds a table's data
/// fields, in `<rowFields>` or `<colFields>`.
const VALUES_FIELD: i32 = -2;

/// The most `<item x>` the writer completes one axis field with: Excel's
/// limit of unique items per pivot field. A field over it fails the write
/// (`SizeLimit`).
pub const MAX_FIELD_ITEMS: usize = 1_048_576;

/// The most items and pivot fields the writer adds to the pivot tables of
/// one workbook (completed `<item x>`s and appended `<pivotField/>`s; what
/// the model gives is not counted). Four fields at [`MAX_FIELD_ITEMS`]: each
/// added element is at most 19 bytes (`<item x="1048575"/>`), so the tables
/// gain at most about 80 MB of XML, which a wasm32 write (4 GiB of memory)
/// holds with the rest of the package. Without it one cache of 100,000
/// shared items named by 2,000 tables, a package of 1.7 MB, wrote 200
/// million items (500 MB, a 5.8 GB heap). Past it the write fails
/// (`SizeLimit`).
pub const MAX_ADDED_ITEMS: usize = 4 * MAX_FIELD_ITEMS;

/// Write `<tag count="N"><field x=".."/>...</tag>` for a row or column field
/// list, with the values field ([`VALUES_FIELD`]) appended when
/// `add_values` is set; nothing when there is no field to write.
fn write_field_refs(
    writer: &mut Writer<&mut Vec<u8>>,
    tag: &str,
    refs: &[PivotFieldRef],
    add_values: bool,
    ibuf: &mut itoa::Buffer,
) -> Result<()> {
    let count = refs.len() + usize::from(add_values);
    if count == 0 {
        return Ok(());
    }
    let mut list = BytesStart::new(tag);
    list.push_attribute(("count", ibuf.format(count)));
    writer
        .write_event(Event::Start(list))
        .map_err(xml_write_err)?;
    let values = add_values.then_some(VALUES_FIELD);
    for x in refs.iter().map(|f| f.x).chain(values) {
        let mut fe = BytesStart::new("field");
        fe.push_attribute(("x", ibuf.format(x)));
        writer
            .write_event(Event::Empty(fe))
            .map_err(xml_write_err)?;
    }
    writer
        .write_event(Event::End(BytesEnd::new(tag)))
        .map_err(xml_write_err)
}

/// Whether the writer completes `field`'s items: it sits on the row, column
/// or page axis and the model gives it none. Excel saves every such field
/// with items, so this marks a field built in code.
fn needs_items(field: &PivotFieldData) -> bool {
    field.items.is_empty()
        && matches!(
            field.axis,
            Some(PivotAxis::AxisRow | PivotAxis::AxisCol | PivotAxis::AxisPage)
        )
}

/// The number of `<item x>` the writer completes the field at `index` with,
/// one per shared item of the cache field at that index (0 without one), or
/// `None` when it writes the field as modelled (see [`needs_items`]).
fn completed_items(
    field: &PivotFieldData,
    index: usize,
    cache: Option<&PivotCacheDefinitionData>,
) -> Option<usize> {
    needs_items(field).then(|| {
        cache
            .and_then(|cache| cache.fields.get(index))
            .map_or(0, |cf| cf.shared_items.len())
    })
}

/// The `t` of each subtotal item of `field`, as Excel writes them after its
/// value items: one per custom subtotal function in Excel's order, else
/// `default` unless `defaultSubtotal` is off.
fn subtotal_items(field: &PivotFieldData) -> impl Iterator<Item = &'static str> + '_ {
    let default =
        (field.subtotals.is_empty() && field.default_subtotal != Some(false)).then_some("default");
    SubtotalFunction::ALL
        .into_iter()
        .filter(|func| field.subtotals.contains(func))
        .map(SubtotalFunction::item_type)
        .chain(default)
}

impl PivotTableData {
    /// Parse a `pivotTable*.xml` file from raw XML bytes.
    ///
    /// # Errors
    ///
    /// `XmlParse` for malformed XML, a bad entity, or a numeric attribute that
    /// does not parse.
    pub fn parse(data: &[u8]) -> Result<Self> {
        let mut reader = Reader::from_reader(data);
        reader.config_mut().trim_text(true);

        let mut buf = Vec::with_capacity(512);

        let mut st = PivotTableParser::new();

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) => {
                    st.handle_open_tag(e, true)?;
                }
                Ok(Event::Empty(ref e)) => {
                    st.handle_open_tag(e, false)?;
                }
                Ok(Event::End(ref e)) => st.handle_close(e.local_name().as_ref()),
                Ok(Event::Eof) => break,
                Ok(_) => {}
                Err(e) => {
                    cold_path();
                    return Err(ModernXlsxError::XmlParse(format!(
                        "error parsing pivot table XML: {e}"
                    )));
                }
            }
            buf.clear();
        }

        Ok(PivotTableData {
            name: st.name,
            data_caption: st.data_caption,
            location: st.location,
            pivot_fields: st.pivot_fields,
            row_fields: st.row_fields,
            col_fields: st.col_fields,
            data_fields: st.data_fields,
            page_fields: st.page_fields,
            cache_id: st.cache_id,
        })
    }

    /// Serialize this pivot table definition to valid OOXML XML bytes, with
    /// no cache to complete items from (see [`Self::to_xml_with_cache`]).
    ///
    /// # Errors
    ///
    /// `SizeLimit` as for [`Self::to_xml_with_cache`]; `XmlWrite` when the
    /// writer fails.
    pub fn to_xml(&self) -> Result<Vec<u8>> {
        self.to_xml_with_cache(None, &mut 0)
    }

    /// The bare `<pivotField/>`s the writer appends so that `<pivotFields>`
    /// has one field per field of `cache` (Excel refuses a table with fewer).
    fn missing_fields(&self, cache: Option<&PivotCacheDefinitionData>) -> usize {
        cache.map_or(0, |cache| {
            cache.fields.len().saturating_sub(self.pivot_fields.len())
        })
    }

    /// Adds to `added` what the writer adds to this table from `cache`: the
    /// `<item x>`s it completes axis fields with (see
    /// [`Self::to_xml_with_cache`]) and the `<pivotField/>`s it appends. Items
    /// and fields the model gives are written as given and not counted.
    ///
    /// # Errors
    ///
    /// `SizeLimit` when a field would be completed with more than
    /// [`MAX_FIELD_ITEMS`] items, or `added` would pass [`MAX_ADDED_ITEMS`].
    pub fn count_added_items(
        &self,
        cache: Option<&PivotCacheDefinitionData>,
        added: &mut usize,
    ) -> Result<()> {
        let completed = self
            .pivot_fields
            .iter()
            .enumerate()
            .filter_map(|(index, field)| Some((index, completed_items(field, index, cache)?)));
        for (index, items) in completed {
            if items > MAX_FIELD_ITEMS {
                cold_path();
                return Err(ModernXlsxError::SizeLimit(format!(
                    "Write failed: pivot table '{}' field {index} would be completed with {items} items from its cache field, more than {MAX_FIELD_ITEMS} (Excel's limit of unique items per pivot field) — give the field explicit items, which are written as given",
                    self.name
                )));
            }
            *added = added.saturating_add(items);
        }
        *added = added.saturating_add(self.missing_fields(cache));
        if *added > MAX_ADDED_ITEMS {
            cold_path();
            return Err(ModernXlsxError::SizeLimit(format!(
                "Write failed: pivot table '{}' brings the items and pivot fields the writer adds from pivot caches past {MAX_ADDED_ITEMS} in one workbook — give the row, column and page fields explicit items and the table one pivot field per cache field, which are written as given and not counted",
                self.name
            )));
        }
        Ok(())
    }

    /// Whether a field on the row, column or page axis has no items. Excel
    /// saves every such field with items, so this marks a table built in
    /// code, whose layout Excel has not computed: the writer completes its
    /// items and has Excel refresh its cache on open.
    #[must_use]
    pub fn has_axis_field_without_items(&self) -> bool {
        self.pivot_fields.iter().any(needs_items)
    }

    /// Serialize this pivot table definition to valid OOXML XML bytes.
    /// `cache` is the definition this table's `cacheId` names.
    ///
    /// `dataCaption` and the three location offsets are always written, from
    /// their defaults when unset, and `<pageFields>` precedes `<dataFields>`
    /// as the schema orders them.
    ///
    /// A field on the row, column or page axis that the model leaves without
    /// items is written with the items Excel needs to refresh the table (an
    /// axis field without them crashes Excel 16 on Refresh): one `<item x>`
    /// per shared item of the cache field at the same index (none without a
    /// cache), then its subtotal items (`default`, or one per custom subtotal
    /// function); no `<items>` when that makes none. With two or more data
    /// fields and no values field (`x="-2"`) among the row and column fields,
    /// the values field is appended to `<colFields>`, where Excel puts it:
    /// Excel refuses the table without it. A field with items, and every
    /// table Excel saves, is written as modelled. When `cache` has more fields
    /// than the model has pivot fields, a bare `<pivotField/>` is appended for
    /// each missing one: Excel refuses a table with fewer pivot fields than
    /// its cache has fields. The model is not changed.
    ///
    /// `added` is the running total of what the writer has added to the
    /// workbook's pivot tables so far; this table's additions are counted
    /// ([`Self::count_added_items`]) before anything is written.
    ///
    /// # Errors
    ///
    /// `SizeLimit` when a field would be completed with more than
    /// [`MAX_FIELD_ITEMS`] items or `added` would pass [`MAX_ADDED_ITEMS`];
    /// `XmlWrite` when the writer fails.
    pub fn to_xml_with_cache(
        &self,
        cache: Option<&PivotCacheDefinitionData>,
        added: &mut usize,
    ) -> Result<Vec<u8>> {
        self.count_added_items(cache, added)?;
        let missing_fields = self.missing_fields(cache);
        let field_count = self.pivot_fields.len() + missing_fields;
        let mut buf: Vec<u8> = Vec::with_capacity(512 + self.pivot_fields.len() * 128);
        let mut writer = Writer::new(&mut buf);
        let mut ibuf = itoa::Buffer::new();

        // XML declaration.
        write_xml_decl(&mut writer)?;

        // <pivotTableDefinition xmlns="..." name="..." cacheId="..." dataCaption="...">
        let mut root = BytesStart::new("pivotTableDefinition");
        root.push_attribute(("xmlns", SPREADSHEET_NS));
        push_text(&mut root, "name", &self.name);
        root.push_attribute(("cacheId", ibuf.format(self.cache_id)));
        push_text(
            &mut root,
            "dataCaption",
            self.data_caption.as_deref().unwrap_or("Values"),
        );
        writer
            .write_event(Event::Start(root))
            .map_err(xml_write_err)?;

        // <location ref="..." firstHeaderRow="..." firstDataRow="..." firstDataCol="..."/>
        let mut loc = BytesStart::new("location");
        push_text(&mut loc, "ref", &self.location.ref_range);
        let header_row = self.location.first_header_row.unwrap_or(1);
        let data_row = self
            .location
            .first_data_row
            .unwrap_or(if self.col_fields.is_empty() { 1 } else { 2 });
        let data_col = self
            .location
            .first_data_col
            .unwrap_or(u32::from(!self.row_fields.is_empty()));
        loc.push_attribute(("firstHeaderRow", ibuf.format(header_row)));
        loc.push_attribute(("firstDataRow", ibuf.format(data_row)));
        loc.push_attribute(("firstDataCol", ibuf.format(data_col)));
        writer
            .write_event(Event::Empty(loc))
            .map_err(xml_write_err)?;

        // <pivotFields count="N">, N at least the cache's field count.
        if field_count > 0 {
            let mut pf_elem = BytesStart::new("pivotFields");
            pf_elem.push_attribute(("count", ibuf.format(field_count)));
            writer
                .write_event(Event::Start(pf_elem))
                .map_err(xml_write_err)?;

            for (index, field) in self.pivot_fields.iter().enumerate() {
                let mut fe = BytesStart::new("pivotField");
                if let Some(axis) = field.axis {
                    fe.push_attribute(("axis", axis.xml_val()));
                }
                if field.data_field {
                    push_flag(&mut fe, "dataField", true);
                }
                if let Some(ref name) = field.name {
                    push_text(&mut fe, "name", name);
                }
                if !field.compact {
                    push_flag(&mut fe, "compact", false);
                }
                if !field.outline {
                    push_flag(&mut fe, "outline", false);
                }
                if let Some(ds) = field.default_subtotal {
                    push_flag(&mut fe, "defaultSubtotal", ds);
                }
                for (i, func) in field.subtotals.iter().enumerate() {
                    // A repeated function must not repeat the attribute.
                    if !field.subtotals[..i].contains(func) {
                        push_flag(&mut fe, func.subtotal_attr(), true);
                    }
                }

                // An axis field without items gets one <item x> per shared
                // item of its cache field, then its subtotal items.
                let completed = completed_items(field, index, cache);
                let count = completed.map_or(field.items.len(), |shared| {
                    shared + subtotal_items(field).count()
                });

                if count == 0 {
                    // Self-closing <pivotField .../>
                    writer
                        .write_event(Event::Empty(fe))
                        .map_err(xml_write_err)?;
                } else {
                    // <pivotField ...>
                    writer
                        .write_event(Event::Start(fe))
                        .map_err(xml_write_err)?;

                    // <items count="N">
                    let mut items_elem = BytesStart::new("items");
                    items_elem.push_attribute(("count", ibuf.format(count)));
                    writer
                        .write_event(Event::Start(items_elem))
                        .map_err(xml_write_err)?;

                    if let Some(shared) = completed {
                        for x in 0..shared {
                            let mut ie = BytesStart::new("item");
                            ie.push_attribute(("x", ibuf.format(x)));
                            writer
                                .write_event(Event::Empty(ie))
                                .map_err(xml_write_err)?;
                        }
                        for t in subtotal_items(field) {
                            let mut ie = BytesStart::new("item");
                            ie.push_attribute(("t", t));
                            writer
                                .write_event(Event::Empty(ie))
                                .map_err(xml_write_err)?;
                        }
                    }
                    for item in &field.items {
                        let mut ie = BytesStart::new("item");
                        if let Some(ref t) = item.t {
                            push_text(&mut ie, "t", t);
                        }
                        if let Some(x) = item.x {
                            ie.push_attribute(("x", ibuf.format(x)));
                        }
                        writer
                            .write_event(Event::Empty(ie))
                            .map_err(xml_write_err)?;
                    }

                    // </items>
                    writer
                        .write_event(Event::End(BytesEnd::new("items")))
                        .map_err(xml_write_err)?;

                    // </pivotField>
                    writer
                        .write_event(Event::End(BytesEnd::new("pivotField")))
                        .map_err(xml_write_err)?;
                }
            }
            // A bare field for each cache field the model has none for.
            for _ in 0..missing_fields {
                writer
                    .write_event(Event::Empty(BytesStart::new("pivotField")))
                    .map_err(xml_write_err)?;
            }

            // </pivotFields>
            writer
                .write_event(Event::End(BytesEnd::new("pivotFields")))
                .map_err(xml_write_err)?;
        }

        let add_values = self.data_fields.len() > 1
            && !self
                .row_fields
                .iter()
                .chain(&self.col_fields)
                .any(|f| f.x == VALUES_FIELD);
        write_field_refs(&mut writer, "rowFields", &self.row_fields, false, &mut ibuf)?;
        write_field_refs(
            &mut writer,
            "colFields",
            &self.col_fields,
            add_values,
            &mut ibuf,
        )?;

        // <pageFields count="N">
        if !self.page_fields.is_empty() {
            let mut pf_elem = BytesStart::new("pageFields");
            pf_elem.push_attribute(("count", ibuf.format(self.page_fields.len())));
            writer
                .write_event(Event::Start(pf_elem))
                .map_err(xml_write_err)?;
            for pf in &self.page_fields {
                let mut pe = BytesStart::new("pageField");
                pe.push_attribute(("fld", ibuf.format(pf.fld)));
                if let Some(item) = pf.item {
                    pe.push_attribute(("item", ibuf.format(item)));
                }
                if let Some(ref name) = pf.name {
                    push_text(&mut pe, "name", name);
                }
                writer
                    .write_event(Event::Empty(pe))
                    .map_err(xml_write_err)?;
            }
            writer
                .write_event(Event::End(BytesEnd::new("pageFields")))
                .map_err(xml_write_err)?;
        }

        // <dataFields count="N">
        if !self.data_fields.is_empty() {
            let mut df_elem = BytesStart::new("dataFields");
            df_elem.push_attribute(("count", ibuf.format(self.data_fields.len())));
            writer
                .write_event(Event::Start(df_elem))
                .map_err(xml_write_err)?;
            for df in &self.data_fields {
                let mut de = BytesStart::new("dataField");
                if let Some(ref name) = df.name {
                    push_text(&mut de, "name", name);
                }
                de.push_attribute(("fld", ibuf.format(df.fld)));
                de.push_attribute(("subtotal", df.subtotal.xml_val()));
                if let Some(nfid) = df.num_fmt_id {
                    de.push_attribute(("numFmtId", ibuf.format(nfid)));
                }
                writer
                    .write_event(Event::Empty(de))
                    .map_err(xml_write_err)?;
            }
            writer
                .write_event(Event::End(BytesEnd::new("dataFields")))
                .map_err(xml_write_err)?;
        }

        // </pivotTableDefinition>
        writer
            .write_event(Event::End(BytesEnd::new("pivotTableDefinition")))
            .map_err(xml_write_err)?;

        Ok(buf)
    }
}

/// Mutable SAX state for [`PivotTableData::parse`].
struct PivotTableParser {
    name: String,
    data_caption: Option<String>,
    cache_id: u32,
    location: PivotLocation,
    pivot_fields: Vec<PivotFieldData>,
    row_fields: Vec<PivotFieldRef>,
    col_fields: Vec<PivotFieldRef>,
    data_fields: Vec<PivotDataFieldData>,
    page_fields: Vec<PivotPageFieldData>,
    in_pivot_fields: bool,
    in_row_fields: bool,
    in_col_fields: bool,
    in_data_fields: bool,
    in_page_fields: bool,
    in_items: bool,
    current_field: Option<PivotFieldData>,
}

impl PivotTableParser {
    fn new() -> Self {
        Self {
            name: String::new(),
            data_caption: None,
            cache_id: 0,
            location: PivotLocation {
                ref_range: String::new(),
                first_header_row: None,
                first_data_row: None,
                first_data_col: None,
            },
            pivot_fields: Vec::new(),
            row_fields: Vec::new(),
            col_fields: Vec::new(),
            data_fields: Vec::new(),
            page_fields: Vec::new(),
            in_pivot_fields: false,
            in_row_fields: false,
            in_col_fields: false,
            in_data_fields: false,
            in_page_fields: false,
            in_items: false,
            current_field: None,
        }
    }

    /// Process a closing element.
    #[inline]
    fn handle_close(&mut self, local: &str) {
        match local {
            "pivotFields" => {
                self.in_pivot_fields = false;
            }
            "pivotField" => {
                if let Some(field) = self.current_field.take() {
                    self.pivot_fields.push(field);
                }
                self.in_items = false;
            }
            "items" => {
                self.in_items = false;
            }
            "rowFields" => {
                self.in_row_fields = false;
            }
            "colFields" => {
                self.in_col_fields = false;
            }
            "dataFields" => {
                self.in_data_fields = false;
            }
            "pageFields" => {
                self.in_page_fields = false;
            }
            _ => {}
        }
    }

    /// Process an opening (`Event::Start`) or self-closing (`Event::Empty`)
    /// element. `is_start` is `true` for `Start`, `false` for `Empty`.
    fn handle_open_tag(&mut self, e: &BytesStart<'_>, is_start: bool) -> Result<()> {
        let local = e.local_name();
        match local.as_ref() {
            "pivotTableDefinition" => {
                for attr in attrs(e) {
                    let attr = attr?;
                    match attr.key.local_name().as_ref() {
                        "name" => {
                            self.name = attr_str(&attr)?.into_owned();
                        }
                        "dataCaption" => {
                            self.data_caption = Some(attr_str(&attr)?.into_owned());
                        }
                        "cacheId" => {
                            self.cache_id = parse_attr("pivotTableDefinition", &attr)?;
                        }
                        _ => {}
                    }
                }
            }
            "location" => {
                for attr in attrs(e) {
                    let attr = attr?;
                    match attr.key.local_name().as_ref() {
                        "ref" => {
                            self.location.ref_range = attr_str(&attr)?.into_owned();
                        }
                        "firstHeaderRow" => {
                            self.location.first_header_row = Some(parse_attr("location", &attr)?);
                        }
                        "firstDataRow" => {
                            self.location.first_data_row = Some(parse_attr("location", &attr)?);
                        }
                        "firstDataCol" => {
                            self.location.first_data_col = Some(parse_attr("location", &attr)?);
                        }
                        _ => {}
                    }
                }
            }
            "pivotFields" if !self.in_pivot_fields && is_start => {
                self.in_pivot_fields = true;
            }
            "pivotField" if self.in_pivot_fields => {
                // OOXML defaults: compact=true, outline=true.
                let mut field = PivotFieldData {
                    axis: None,
                    name: None,
                    items: Vec::new(),
                    subtotals: Vec::new(),
                    compact: true,
                    outline: true,
                    data_field: false,
                    default_subtotal: None,
                };
                for attr in attrs(e) {
                    let attr = attr?;
                    let key = attr.key.local_name();
                    match key.as_ref() {
                        "axis" => {
                            field.axis = PivotAxis::from_xml(&attr_str(&attr)?);
                        }
                        "name" => {
                            field.name = Some(attr_str(&attr)?.into_owned());
                        }
                        "compact" => {
                            field.compact = xsd_bool(&attr_str(&attr)?).unwrap_or(true);
                        }
                        "outline" => {
                            field.outline = xsd_bool(&attr_str(&attr)?).unwrap_or(true);
                        }
                        "dataField" => {
                            field.data_field = xsd_bool(&attr_str(&attr)?).unwrap_or(false);
                        }
                        "defaultSubtotal" => {
                            field.default_subtotal = xsd_bool(&attr_str(&attr)?);
                        }
                        other => {
                            if let Some(func) = SubtotalFunction::from_subtotal_attr(other)
                                && xsd_bool(&attr_str(&attr)?) == Some(true)
                            {
                                field.subtotals.push(func);
                            }
                        }
                    }
                }
                if is_start {
                    self.current_field = Some(field);
                } else {
                    // Self-closing <pivotField .../> — push immediately.
                    self.pivot_fields.push(field);
                }
            }
            "items" if self.current_field.is_some() && is_start => {
                self.in_items = true;
            }
            "item"
                if self.in_items
                    && let Some(field) = self.current_field.as_mut() =>
            {
                let mut item = PivotItem { t: None, x: None };
                for attr in attrs(e) {
                    let attr = attr?;
                    match attr.key.local_name().as_ref() {
                        "t" => {
                            item.t = Some(attr_str(&attr)?.into_owned());
                        }
                        "x" => {
                            item.x = Some(parse_attr("item", &attr)?);
                        }
                        _ => {}
                    }
                }
                field.items.push(item);
            }
            "rowFields" if is_start => {
                self.in_row_fields = true;
            }
            "colFields" if is_start => {
                self.in_col_fields = true;
            }
            "field" if self.in_row_fields || self.in_col_fields => {
                let mut x: i32 = 0;
                for attr in attrs(e) {
                    let attr = attr?;
                    if attr.key.local_name().as_ref() == "x" {
                        x = parse_attr("field", &attr)?;
                    }
                }
                let field_ref = PivotFieldRef { x };
                if self.in_row_fields {
                    self.row_fields.push(field_ref);
                } else {
                    self.col_fields.push(field_ref);
                }
            }
            "dataFields" if is_start => {
                self.in_data_fields = true;
            }
            "dataField" if self.in_data_fields => {
                let mut df = PivotDataFieldData {
                    name: None,
                    fld: 0,
                    subtotal: SubtotalFunction::default(),
                    num_fmt_id: None,
                };
                for attr in attrs(e) {
                    let attr = attr?;
                    match attr.key.local_name().as_ref() {
                        "name" => {
                            df.name = Some(attr_str(&attr)?.into_owned());
                        }
                        "fld" => {
                            df.fld = parse_attr("dataField", &attr)?;
                        }
                        "subtotal" => {
                            df.subtotal =
                                SubtotalFunction::from_xml(&attr_str(&attr)?).unwrap_or_default();
                        }
                        "numFmtId" => {
                            df.num_fmt_id = Some(parse_attr("dataField", &attr)?);
                        }
                        _ => {}
                    }
                }
                self.data_fields.push(df);
            }
            "pageFields" if is_start => {
                self.in_page_fields = true;
            }
            "pageField" if self.in_page_fields => {
                let mut pf = PivotPageFieldData {
                    fld: 0,
                    item: None,
                    name: None,
                };
                for attr in attrs(e) {
                    let attr = attr?;
                    match attr.key.local_name().as_ref() {
                        "fld" => {
                            pf.fld = parse_attr("pageField", &attr)?;
                        }
                        "item" => {
                            pf.item = Some(parse_attr("pageField", &attr)?);
                        }
                        "name" => {
                            pf.name = Some(attr_str(&attr)?.into_owned());
                        }
                        _ => {}
                    }
                }
                self.page_fields.push(pf);
            }
            _ => {}
        }
        Ok(())
    }
}

/// The rules the writer applies to [`PivotCacheDefinitionData::relationships`]
/// and to a kept `<cacheSource>` span.
#[cfg(test)]
mod kept_part_tests {
    use super::*;

    const EXTERNAL_LINK_PATH: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath";

    /// The `.rels` of the cache definition Excel 16 saved on 2026-11-12 for
    /// a pivot table over `[xext-source.xlsx]Data!R1C1:R3C3` (through COM,
    /// `PivotCaches().Create(1, ...)`; `Create(3, ...)`, a consolidation
    /// range, gave the same part), after its XML declaration.
    const EXCEL_EXTERNAL_CACHE_RELS: &str = r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath" Target="xext-source.xlsx" TargetMode="External"/><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheRecords" Target="pivotCacheRecords1.xml"/></Relationships>"#;

    fn rel(id: &str, rel_type: &str, target_mode: Option<&str>) -> CacheRelationship {
        CacheRelationship {
            id: id.to_owned(),
            rel_type: rel_type.to_owned(),
            target: "xext-source.xlsx".to_owned(),
            target_mode: target_mode.map(str::to_owned),
        }
    }

    /// A definition over a consolidation source whose `<rangeSet>` names
    /// `source_id`, with `relationships`.
    fn def(source_id: &str, relationships: Vec<CacheRelationship>) -> PivotCacheDefinitionData {
        PivotCacheDefinitionData {
            cache_id: None,
            source: CacheSource {
                raw_xml: Some(format!(
                    r#"<cacheSource type="consolidation"><consolidation autoPage="0"><rangeSets count="1"><rangeSet ref="A1:C3" sheet="Data" r:id="{source_id}"/></rangeSets></consolidation></cacheSource>"#
                )),
                ..Default::default()
            },
            fields: Vec::new(),
            record_count: None,
            relationships,
        }
    }

    fn written(def: &PivotCacheDefinitionData) -> String {
        String::from_utf8(def.to_xml().unwrap()).unwrap()
    }

    #[test]
    fn a_kept_target_mode_is_external_in_any_ascii_case() {
        for (mode, external) in [
            (Some("External"), true),
            (Some("external"), true),
            (Some("EXTERNAL"), true),
            (Some("eXtErNaL"), true),
            (None, false),
            (Some("Internal"), false),
            (Some("Externa"), false),
            (Some("External "), false),
            (Some("\u{212A}xternal"), false),
        ] {
            let def = def("rId2", vec![rel("rId2", EXTERNAL_LINK_PATH, mode)]);
            let rels = def.kept_relationships();
            assert_eq!(
                rels.get_by_id("rId2").unwrap().target_mode == TargetMode::External,
                external,
                "{mode:?}"
            );
            // The attribute is always written as `External`.
            let xml = String::from_utf8(rels.to_xml().unwrap()).unwrap();
            assert_eq!(
                xml.contains(r#" TargetMode="External"/>"#),
                external,
                "{mode:?}: {xml}"
            );
            assert!(written(&def).contains(r#"r:id="rId2""#), "{mode:?}");
        }
    }

    #[test]
    fn a_kept_relationship_of_the_records_type_is_skipped() {
        let def = def(
            "rId2",
            vec![
                rel("rId2", REL_PIVOT_CACHE_REC, None),
                rel("rId3", EXTERNAL_LINK_PATH, Some("External")),
            ],
        );
        let rels = def.kept_relationships();
        assert_eq!(rels.len(), 1);
        assert!(rels.get_by_id("rId3").is_some());
        // The records relationship takes rId1, and the source no longer
        // names the skipped one.
        assert_eq!(def.records_r_id(), "rId1");
        let out = written(&def);
        assert!(
            out.contains(r#"<rangeSet ref="A1:C3" sheet="Data"/>"#),
            "{out}"
        );
    }

    #[test]
    fn a_kept_relationship_with_an_id_outside_the_id_rule_is_skipped() {
        for bad in [
            "",
            "r Id2",
            "rId\u{e9}",
            "a:b",
            "rId2\t",
            "<x>",
            "a&b",
            "9",
            "-x",
            ".x",
        ] {
            // A kept span holds the id as an attribute value, so only ids
            // that need no escaping there go in one.
            for raw in [true, false] {
                if raw && bad.contains(['<', '&']) {
                    continue;
                }
                let mut def = def(bad, vec![rel(bad, EXTERNAL_LINK_PATH, Some("External"))]);
                if !raw {
                    def.source = CacheSource {
                        ref_range: "A1:C3".to_owned(),
                        sheet: "Data".to_owned(),
                        r_id: Some(bad.to_owned()),
                        ..Default::default()
                    };
                }
                assert!(def.kept_relationships().is_empty(), "{bad:?}");
                assert_eq!(def.records_r_id(), "rId1");
                // No written r:id names the skipped relationship.
                let out = written(&def);
                assert!(!out.contains(" r:id="), "{bad:?} raw={raw}: {out}");
            }
        }
        for good in ["rId2", "x_y-z.1", "R", "_x"] {
            let def = def(good, vec![rel(good, EXTERNAL_LINK_PATH, Some("External"))]);
            assert_eq!(def.kept_relationships().len(), 1, "{good:?}");
            assert!(
                written(&def).contains(&format!(r#"r:id="{good}""#)),
                "{good:?}"
            );
        }
    }

    #[test]
    fn a_kept_source_with_an_undeclared_entity_fails_to_read() {
        const MAIN: &str = r#"xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships""#;
        // The cases: an entity the part's DTD declares,
        // used in an attribute and in text of the kept span (the written
        // part carries no DTD), and one no DTD declares.
        for xml in [
            format!(
                r#"<!DOCTYPE p [<!ENTITY e "x">]><pivotCacheDefinition {MAIN}><cacheSource type="consolidation"><consolidation><rangeSets count="1"><rangeSet ref="A1" sheet="&e;" r:id="rId2"/></rangeSets></consolidation></cacheSource><cacheFields count="0"/></pivotCacheDefinition>"#
            ),
            format!(
                r#"<!DOCTYPE p [<!ENTITY e "x">]><pivotCacheDefinition {MAIN}><cacheSource type="consolidation"><consolidation>&e;</consolidation></cacheSource><cacheFields count="0"/></pivotCacheDefinition>"#
            ),
            format!(
                r#"<pivotCacheDefinition {MAIN}><cacheSource type="external" connectionId="1"><x a="&nbsp;"/></cacheSource><cacheFields count="0"/></pivotCacheDefinition>"#
            ),
            format!(
                r#"<pivotCacheDefinition {MAIN}><cacheSource type="external">&#0;</cacheSource><cacheFields count="0"/></pivotCacheDefinition>"#
            ),
        ] {
            let err = PivotCacheDefinitionData::parse(xml.as_bytes()).unwrap_err();
            assert!(
                matches!(err, ModernXlsxError::XmlParse(_)),
                "{xml}: {err:?}"
            );
        }
        // The five predefined entities and character references read and
        // are written as read.
        let span = r#"<cacheSource type="external" connectionId="1"><x a="Q&amp;A&#10;&lt;&gt;&quot;&apos;" b='&#x22;'>t&amp;u&#10;&lt;&gt;&quot;&apos;&#xA0;</x></cacheSource>"#;
        let def = PivotCacheDefinitionData::parse(
            format!(r#"<pivotCacheDefinition {MAIN}>{span}<cacheFields count="0"/></pivotCacheDefinition>"#)
                .as_bytes(),
        )
        .unwrap();
        assert_eq!(def.source.raw_xml.as_deref(), Some(span));
        let out = written(&def);
        assert!(out.contains(&format!("{span}<cacheFields")), "{out}");
        assert_eq!(
            PivotCacheDefinitionData::parse(out.as_bytes()).unwrap(),
            def
        );
    }

    #[test]
    fn an_excel_saved_external_relationship_list_is_written_back_unchanged() {
        let excel = Relationships::parse(EXCEL_EXTERNAL_CACHE_RELS.as_bytes()).unwrap();
        // As the reader keeps them: every relationship but the records one.
        let kept: Vec<CacheRelationship> = excel
            .iter()
            .filter(|rel| rel.rel_type != REL_PIVOT_CACHE_REC)
            .map(CacheRelationship::from_relationship)
            .collect();
        assert_eq!(
            kept,
            vec![rel("rId2", EXTERNAL_LINK_PATH, Some("External"))]
        );
        let def = def("rId2", kept);
        // As the writer builds the .rels.
        let mut rels = def.kept_relationships();
        rels.add(
            def.records_r_id(),
            REL_PIVOT_CACHE_REC,
            "pivotCacheRecords1.xml",
        );
        let xml = String::from_utf8(rels.to_xml().unwrap()).unwrap();
        assert_eq!(
            &xml[xml.find("<Relationships").unwrap()..],
            EXCEL_EXTERNAL_CACHE_RELS
        );
        assert!(written(&def).contains(r#"sheet="Data" r:id="rId2"/>"#));
    }
}

// `pub(crate)` so the writer's tests share the Excel-made pivot fixtures.
#[cfg(test)]
pub(crate) mod tests;
