use core::hint::cold_path;

use quick_xml::events::{BytesDecl, BytesEnd, BytesStart, Event};
use quick_xml::{Reader, Writer};
use serde::{Deserialize, Serialize};

use super::SPREADSHEET_NS;
use crate::{ModernXlsxError, Result};

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
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum SubtotalFunction {
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

impl Default for SubtotalFunction {
    fn default() -> Self {
        Self::Sum
    }
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
}

/// Top-level pivot table definition.
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct PivotTableData {
    pub name: String,
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
    pub cache_id: u32,
}

/// The location reference for a pivot table within a worksheet.
#[derive(Debug, Clone, Serialize, Deserialize)]
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
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct PivotFieldData {
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub axis: Option<PivotAxis>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub name: Option<String>,
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub items: Vec<PivotItem>,
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub subtotals: Vec<SubtotalFunction>,
    #[serde(default, skip_serializing_if = "crate::ooxml::is_false")]
    pub compact: bool,
    #[serde(default, skip_serializing_if = "crate::ooxml::is_false")]
    pub outline: bool,
}

/// An item within a pivot field.
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct PivotItem {
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub t: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub x: Option<u32>,
}

/// A reference to a field by index, used in row/col field lists.
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct PivotFieldRef {
    pub x: i32,
}

/// A data field (value area) definition in the pivot table.
#[derive(Debug, Clone, Serialize, Deserialize)]
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
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct PivotPageFieldData {
    pub fld: i32,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub item: Option<u32>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub name: Option<String>,
}

impl PivotTableData {
    /// Parse a `pivotTable*.xml` file from raw XML bytes.
    pub fn parse(data: &[u8]) -> Result<Self> {
        let mut reader = Reader::from_reader(data);
        reader.config_mut().trim_text(true);

        let mut buf = Vec::with_capacity(512);

        // Result fields.
        let mut name = String::new();
        let mut data_caption: Option<String> = None;
        let mut cache_id: u32 = 0;
        let mut location = PivotLocation {
            ref_range: String::new(),
            first_header_row: None,
            first_data_row: None,
            first_data_col: None,
        };
        let mut pivot_fields: Vec<PivotFieldData> = Vec::new();
        let mut row_fields: Vec<PivotFieldRef> = Vec::new();
        let mut col_fields: Vec<PivotFieldRef> = Vec::new();
        let mut data_fields: Vec<PivotDataFieldData> = Vec::new();
        let mut page_fields: Vec<PivotPageFieldData> = Vec::new();

        // State flags.
        let mut in_pivot_fields = false;
        let mut in_row_fields = false;
        let mut in_col_fields = false;
        let mut in_data_fields = false;
        let mut in_page_fields = false;
        let mut in_items = false;
        let mut current_field: Option<PivotFieldData> = None;

        loop {
            let is_start;
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) => {
                    is_start = true;
                    Self::handle_open_tag(
                        e,
                        is_start,
                        &mut name,
                        &mut data_caption,
                        &mut cache_id,
                        &mut location,
                        &mut pivot_fields,
                        &mut row_fields,
                        &mut col_fields,
                        &mut data_fields,
                        &mut page_fields,
                        &mut in_pivot_fields,
                        &mut in_row_fields,
                        &mut in_col_fields,
                        &mut in_data_fields,
                        &mut in_page_fields,
                        &mut in_items,
                        &mut current_field,
                    );
                }
                Ok(Event::Empty(ref e)) => {
                    is_start = false;
                    Self::handle_open_tag(
                        e,
                        is_start,
                        &mut name,
                        &mut data_caption,
                        &mut cache_id,
                        &mut location,
                        &mut pivot_fields,
                        &mut row_fields,
                        &mut col_fields,
                        &mut data_fields,
                        &mut page_fields,
                        &mut in_pivot_fields,
                        &mut in_row_fields,
                        &mut in_col_fields,
                        &mut in_data_fields,
                        &mut in_page_fields,
                        &mut in_items,
                        &mut current_field,
                    );
                }
                Ok(Event::End(ref e)) => {
                    let local = e.local_name();
                    match local.as_ref() {
                        b"pivotFields" => {
                            in_pivot_fields = false;
                        }
                        b"pivotField" => {
                            if let Some(field) = current_field.take() {
                                pivot_fields.push(field);
                            }
                            in_items = false;
                        }
                        b"items" => {
                            in_items = false;
                        }
                        b"rowFields" => {
                            in_row_fields = false;
                        }
                        b"colFields" => {
                            in_col_fields = false;
                        }
                        b"dataFields" => {
                            in_data_fields = false;
                        }
                        b"pageFields" => {
                            in_page_fields = false;
                        }
                        _ => {}
                    }
                }
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
            name,
            data_caption,
            location,
            pivot_fields,
            row_fields,
            col_fields,
            data_fields,
            page_fields,
            cache_id,
        })
    }

    /// Serialize this pivot table definition to valid OOXML XML bytes.
    pub fn to_xml(&self) -> Result<Vec<u8>> {
        let mut buf: Vec<u8> = Vec::with_capacity(512);
        let mut writer = Writer::new(&mut buf);
        let mut ibuf = itoa::Buffer::new();

        let map_err = |e: std::io::Error| ModernXlsxError::XmlWrite(e.to_string());

        // XML declaration.
        writer
            .write_event(Event::Decl(BytesDecl::new("1.0", Some("UTF-8"), Some("yes"))))
            .map_err(map_err)?;

        // <pivotTableDefinition xmlns="..." name="..." cacheId="..." dataCaption="...">
        let mut root = BytesStart::new("pivotTableDefinition");
        root.push_attribute(("xmlns", SPREADSHEET_NS));
        root.push_attribute(("name", self.name.as_str()));
        root.push_attribute(("cacheId", ibuf.format(self.cache_id)));
        if let Some(ref dc) = self.data_caption {
            root.push_attribute(("dataCaption", dc.as_str()));
        }
        writer.write_event(Event::Start(root)).map_err(map_err)?;

        // <location ref="..." firstHeaderRow="..." firstDataRow="..." firstDataCol="..."/>
        let mut loc = BytesStart::new("location");
        loc.push_attribute(("ref", self.location.ref_range.as_str()));
        if let Some(v) = self.location.first_header_row {
            loc.push_attribute(("firstHeaderRow", ibuf.format(v)));
        }
        if let Some(v) = self.location.first_data_row {
            loc.push_attribute(("firstDataRow", ibuf.format(v)));
        }
        if let Some(v) = self.location.first_data_col {
            loc.push_attribute(("firstDataCol", ibuf.format(v)));
        }
        writer.write_event(Event::Empty(loc)).map_err(map_err)?;

        // <pivotFields count="N">
        if !self.pivot_fields.is_empty() {
            let mut pf_elem = BytesStart::new("pivotFields");
            pf_elem.push_attribute(("count", ibuf.format(self.pivot_fields.len())));
            writer
                .write_event(Event::Start(pf_elem))
                .map_err(map_err)?;

            for field in &self.pivot_fields {
                let mut fe = BytesStart::new("pivotField");
                if let Some(axis) = field.axis {
                    fe.push_attribute(("axis", axis.xml_val()));
                }
                if let Some(ref name) = field.name {
                    fe.push_attribute(("name", name.as_str()));
                }
                if !field.compact {
                    fe.push_attribute(("compact", "0"));
                }
                if !field.outline {
                    fe.push_attribute(("outline", "0"));
                }

                if field.items.is_empty() {
                    // Self-closing <pivotField .../>
                    writer.write_event(Event::Empty(fe)).map_err(map_err)?;
                } else {
                    // <pivotField ...>
                    writer.write_event(Event::Start(fe)).map_err(map_err)?;

                    // <items count="N">
                    let mut items_elem = BytesStart::new("items");
                    items_elem.push_attribute(("count", ibuf.format(field.items.len())));
                    writer
                        .write_event(Event::Start(items_elem))
                        .map_err(map_err)?;

                    for item in &field.items {
                        let mut ie = BytesStart::new("item");
                        if let Some(ref t) = item.t {
                            ie.push_attribute(("t", t.as_str()));
                        }
                        if let Some(x) = item.x {
                            ie.push_attribute(("x", ibuf.format(x)));
                        }
                        writer.write_event(Event::Empty(ie)).map_err(map_err)?;
                    }

                    // </items>
                    writer
                        .write_event(Event::End(BytesEnd::new("items")))
                        .map_err(map_err)?;

                    // </pivotField>
                    writer
                        .write_event(Event::End(BytesEnd::new("pivotField")))
                        .map_err(map_err)?;
                }
            }

            // </pivotFields>
            writer
                .write_event(Event::End(BytesEnd::new("pivotFields")))
                .map_err(map_err)?;
        }

        // <rowFields count="N">
        if !self.row_fields.is_empty() {
            let mut rf_elem = BytesStart::new("rowFields");
            rf_elem.push_attribute(("count", ibuf.format(self.row_fields.len())));
            writer
                .write_event(Event::Start(rf_elem))
                .map_err(map_err)?;
            for f in &self.row_fields {
                let mut fe = BytesStart::new("field");
                fe.push_attribute(("x", ibuf.format(f.x)));
                writer.write_event(Event::Empty(fe)).map_err(map_err)?;
            }
            writer
                .write_event(Event::End(BytesEnd::new("rowFields")))
                .map_err(map_err)?;
        }

        // <colFields count="N">
        if !self.col_fields.is_empty() {
            let mut cf_elem = BytesStart::new("colFields");
            cf_elem.push_attribute(("count", ibuf.format(self.col_fields.len())));
            writer
                .write_event(Event::Start(cf_elem))
                .map_err(map_err)?;
            for f in &self.col_fields {
                let mut fe = BytesStart::new("field");
                fe.push_attribute(("x", ibuf.format(f.x)));
                writer.write_event(Event::Empty(fe)).map_err(map_err)?;
            }
            writer
                .write_event(Event::End(BytesEnd::new("colFields")))
                .map_err(map_err)?;
        }

        // <dataFields count="N">
        if !self.data_fields.is_empty() {
            let mut df_elem = BytesStart::new("dataFields");
            df_elem.push_attribute(("count", ibuf.format(self.data_fields.len())));
            writer
                .write_event(Event::Start(df_elem))
                .map_err(map_err)?;
            for df in &self.data_fields {
                let mut de = BytesStart::new("dataField");
                if let Some(ref name) = df.name {
                    de.push_attribute(("name", name.as_str()));
                }
                de.push_attribute(("fld", ibuf.format(df.fld)));
                de.push_attribute(("subtotal", df.subtotal.xml_val()));
                if let Some(nfid) = df.num_fmt_id {
                    de.push_attribute(("numFmtId", ibuf.format(nfid)));
                }
                writer.write_event(Event::Empty(de)).map_err(map_err)?;
            }
            writer
                .write_event(Event::End(BytesEnd::new("dataFields")))
                .map_err(map_err)?;
        }

        // <pageFields count="N">
        if !self.page_fields.is_empty() {
            let mut pf_elem = BytesStart::new("pageFields");
            pf_elem.push_attribute(("count", ibuf.format(self.page_fields.len())));
            writer
                .write_event(Event::Start(pf_elem))
                .map_err(map_err)?;
            for pf in &self.page_fields {
                let mut pe = BytesStart::new("pageField");
                pe.push_attribute(("fld", ibuf.format(pf.fld)));
                if let Some(item) = pf.item {
                    pe.push_attribute(("item", ibuf.format(item)));
                }
                if let Some(ref name) = pf.name {
                    pe.push_attribute(("name", name.as_str()));
                }
                writer.write_event(Event::Empty(pe)).map_err(map_err)?;
            }
            writer
                .write_event(Event::End(BytesEnd::new("pageFields")))
                .map_err(map_err)?;
        }

        // </pivotTableDefinition>
        writer
            .write_event(Event::End(BytesEnd::new("pivotTableDefinition")))
            .map_err(map_err)?;

        Ok(buf)
    }

    /// Process an opening (`Event::Start`) or self-closing (`Event::Empty`)
    /// element. `is_start` is `true` for `Start`, `false` for `Empty`.
    #[allow(clippy::too_many_arguments)]
    fn handle_open_tag(
        e: &quick_xml::events::BytesStart<'_>,
        is_start: bool,
        name: &mut String,
        data_caption: &mut Option<String>,
        cache_id: &mut u32,
        location: &mut PivotLocation,
        pivot_fields: &mut Vec<PivotFieldData>,
        row_fields: &mut Vec<PivotFieldRef>,
        col_fields: &mut Vec<PivotFieldRef>,
        data_fields: &mut Vec<PivotDataFieldData>,
        page_fields: &mut Vec<PivotPageFieldData>,
        in_pivot_fields: &mut bool,
        in_row_fields: &mut bool,
        in_col_fields: &mut bool,
        in_data_fields: &mut bool,
        in_page_fields: &mut bool,
        in_items: &mut bool,
        current_field: &mut Option<PivotFieldData>,
    ) {
        let local = e.local_name();
        match local.as_ref() {
            b"pivotTableDefinition" => {
                for attr in e.attributes().flatten() {
                    match attr.key.local_name().as_ref() {
                        b"name" => {
                            *name =
                                String::from_utf8_lossy(&attr.value).into_owned();
                        }
                        b"dataCaption" => {
                            *data_caption = Some(
                                String::from_utf8_lossy(&attr.value).into_owned(),
                            );
                        }
                        b"cacheId" => {
                            *cache_id = std::str::from_utf8(&attr.value)
                                .ok()
                                .and_then(|s| s.parse().ok())
                                .unwrap_or(0);
                        }
                        _ => {}
                    }
                }
            }
            b"location" => {
                for attr in e.attributes().flatten() {
                    match attr.key.local_name().as_ref() {
                        b"ref" => {
                            location.ref_range =
                                String::from_utf8_lossy(&attr.value).into_owned();
                        }
                        b"firstHeaderRow" => {
                            location.first_header_row =
                                std::str::from_utf8(&attr.value)
                                    .ok()
                                    .and_then(|s| s.parse().ok());
                        }
                        b"firstDataRow" => {
                            location.first_data_row =
                                std::str::from_utf8(&attr.value)
                                    .ok()
                                    .and_then(|s| s.parse().ok());
                        }
                        b"firstDataCol" => {
                            location.first_data_col =
                                std::str::from_utf8(&attr.value)
                                    .ok()
                                    .and_then(|s| s.parse().ok());
                        }
                        _ => {}
                    }
                }
            }
            b"pivotFields" if !*in_pivot_fields => {
                if is_start {
                    *in_pivot_fields = true;
                }
            }
            b"pivotField" if *in_pivot_fields => {
                // OOXML defaults: compact=true, outline=true.
                let mut field = PivotFieldData {
                    axis: None,
                    name: None,
                    items: Vec::new(),
                    subtotals: Vec::new(),
                    compact: true,
                    outline: true,
                };
                for attr in e.attributes().flatten() {
                    match attr.key.local_name().as_ref() {
                        b"axis" => {
                            field.axis = std::str::from_utf8(&attr.value)
                                .ok()
                                .and_then(PivotAxis::from_xml);
                        }
                        b"name" => {
                            field.name = Some(
                                String::from_utf8_lossy(&attr.value).into_owned(),
                            );
                        }
                        b"compact" => {
                            field.compact = attr.value.as_ref() != b"0";
                        }
                        b"outline" => {
                            field.outline = attr.value.as_ref() != b"0";
                        }
                        b"dataField" => {
                            // Presence noted but not stored directly;
                            // the dataFields section captures this.
                        }
                        _ => {}
                    }
                }
                if is_start {
                    *current_field = Some(field);
                } else {
                    // Self-closing <pivotField .../> — push immediately.
                    pivot_fields.push(field);
                }
            }
            b"items" if current_field.is_some() => {
                if is_start {
                    *in_items = true;
                }
            }
            b"item" if *in_items => {
                if let Some(ref mut field) = *current_field {
                    let mut item = PivotItem { t: None, x: None };
                    for attr in e.attributes().flatten() {
                        match attr.key.local_name().as_ref() {
                            b"t" => {
                                item.t = Some(
                                    String::from_utf8_lossy(&attr.value)
                                        .into_owned(),
                                );
                            }
                            b"x" => {
                                item.x = std::str::from_utf8(&attr.value)
                                    .ok()
                                    .and_then(|s| s.parse().ok());
                            }
                            _ => {}
                        }
                    }
                    field.items.push(item);
                }
            }
            b"rowFields" => {
                if is_start {
                    *in_row_fields = true;
                }
            }
            b"colFields" => {
                if is_start {
                    *in_col_fields = true;
                }
            }
            b"field" if *in_row_fields || *in_col_fields => {
                let mut x: i32 = 0;
                for attr in e.attributes().flatten() {
                    if attr.key.local_name().as_ref() == b"x" {
                        x = std::str::from_utf8(&attr.value)
                            .ok()
                            .and_then(|s| s.parse().ok())
                            .unwrap_or(0);
                    }
                }
                let field_ref = PivotFieldRef { x };
                if *in_row_fields {
                    row_fields.push(field_ref);
                } else {
                    col_fields.push(field_ref);
                }
            }
            b"dataFields" => {
                if is_start {
                    *in_data_fields = true;
                }
            }
            b"dataField" if *in_data_fields => {
                let mut df = PivotDataFieldData {
                    name: None,
                    fld: 0,
                    subtotal: SubtotalFunction::default(),
                    num_fmt_id: None,
                };
                for attr in e.attributes().flatten() {
                    match attr.key.local_name().as_ref() {
                        b"name" => {
                            df.name = Some(
                                String::from_utf8_lossy(&attr.value).into_owned(),
                            );
                        }
                        b"fld" => {
                            df.fld = std::str::from_utf8(&attr.value)
                                .ok()
                                .and_then(|s| s.parse().ok())
                                .unwrap_or(0);
                        }
                        b"subtotal" => {
                            df.subtotal = std::str::from_utf8(&attr.value)
                                .ok()
                                .and_then(SubtotalFunction::from_xml)
                                .unwrap_or_default();
                        }
                        b"numFmtId" => {
                            df.num_fmt_id = std::str::from_utf8(&attr.value)
                                .ok()
                                .and_then(|s| s.parse().ok());
                        }
                        _ => {}
                    }
                }
                data_fields.push(df);
            }
            b"pageFields" => {
                if is_start {
                    *in_page_fields = true;
                }
            }
            b"pageField" if *in_page_fields => {
                let mut pf = PivotPageFieldData {
                    fld: 0,
                    item: None,
                    name: None,
                };
                for attr in e.attributes().flatten() {
                    match attr.key.local_name().as_ref() {
                        b"fld" => {
                            pf.fld = std::str::from_utf8(&attr.value)
                                .ok()
                                .and_then(|s| s.parse().ok())
                                .unwrap_or(0);
                        }
                        b"item" => {
                            pf.item = std::str::from_utf8(&attr.value)
                                .ok()
                                .and_then(|s| s.parse().ok());
                        }
                        b"name" => {
                            pf.name = Some(
                                String::from_utf8_lossy(&attr.value).into_owned(),
                            );
                        }
                        _ => {}
                    }
                }
                page_fields.push(pf);
            }
            _ => {}
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn sample_pivot_table_xml() -> &'static [u8] {
        br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  name="PivotTable1" cacheId="0" dataCaption="Values">
  <location ref="A3:C11" firstHeaderRow="1" firstDataRow="2" firstDataCol="1"/>
  <pivotFields count="3">
    <pivotField axis="axisRow" compact="0" outline="0">
      <items count="2">
        <item x="0"/>
        <item t="default"/>
      </items>
    </pivotField>
    <pivotField axis="axisCol">
      <items count="2">
        <item x="0"/>
        <item t="default"/>
      </items>
    </pivotField>
    <pivotField dataField="1"/>
  </pivotFields>
  <rowFields count="1"><field x="0"/></rowFields>
  <colFields count="1"><field x="1"/></colFields>
  <dataFields count="1">
    <dataField name="Sum of Amount" fld="2" subtotal="sum"/>
  </dataFields>
  <pageFields count="1">
    <pageField fld="0" item="0" name="Filter"/>
  </pageFields>
</pivotTableDefinition>"#
    }

    #[test]
    fn pivot_table_roundtrip() {
        let pt = PivotTableData::parse(sample_pivot_table_xml()).unwrap();
        let xml = pt.to_xml().unwrap();
        let pt2 = PivotTableData::parse(&xml).unwrap();
        assert_eq!(pt2.name, pt.name);
        assert_eq!(pt2.cache_id, pt.cache_id);
        assert_eq!(pt2.data_caption, pt.data_caption);
        assert_eq!(pt2.location.ref_range, pt.location.ref_range);
        assert_eq!(pt2.location.first_header_row, pt.location.first_header_row);
        assert_eq!(pt2.location.first_data_row, pt.location.first_data_row);
        assert_eq!(pt2.location.first_data_col, pt.location.first_data_col);
        assert_eq!(pt2.pivot_fields.len(), pt.pivot_fields.len());
        for (a, b) in pt2.pivot_fields.iter().zip(&pt.pivot_fields) {
            assert_eq!(a.axis, b.axis);
            assert_eq!(a.compact, b.compact);
            assert_eq!(a.outline, b.outline);
            assert_eq!(a.items.len(), b.items.len());
        }
        assert_eq!(pt2.row_fields.len(), pt.row_fields.len());
        assert_eq!(pt2.col_fields.len(), pt.col_fields.len());
        assert_eq!(pt2.data_fields.len(), pt.data_fields.len());
        assert_eq!(pt2.data_fields[0].subtotal, pt.data_fields[0].subtotal);
    }

    #[test]
    fn parse_pivot_table_definition() {
        let pt = PivotTableData::parse(sample_pivot_table_xml()).unwrap();
        assert_eq!(pt.name, "PivotTable1");
        assert_eq!(pt.cache_id, 0);
        assert_eq!(pt.data_caption.as_deref(), Some("Values"));
        assert_eq!(pt.location.ref_range, "A3:C11");
        assert_eq!(pt.location.first_header_row, Some(1));
        assert_eq!(pt.location.first_data_row, Some(2));
        assert_eq!(pt.location.first_data_col, Some(1));
        assert_eq!(pt.pivot_fields.len(), 3);
        assert_eq!(pt.pivot_fields[0].axis, Some(PivotAxis::AxisRow));
        assert_eq!(pt.pivot_fields[0].compact, false);
        assert_eq!(pt.pivot_fields[0].outline, false);
        assert_eq!(pt.pivot_fields[0].items.len(), 2);
        assert_eq!(pt.pivot_fields[0].items[0].x, Some(0));
        assert_eq!(pt.pivot_fields[0].items[1].t.as_deref(), Some("default"));
        assert_eq!(pt.pivot_fields[1].axis, Some(PivotAxis::AxisCol));
        assert_eq!(pt.pivot_fields[2].axis, None); // dataField only, no axis
        assert_eq!(pt.row_fields.len(), 1);
        assert_eq!(pt.row_fields[0].x, 0);
        assert_eq!(pt.col_fields.len(), 1);
        assert_eq!(pt.col_fields[0].x, 1);
        assert_eq!(pt.data_fields.len(), 1);
        assert_eq!(pt.data_fields[0].name.as_deref(), Some("Sum of Amount"));
        assert_eq!(pt.data_fields[0].fld, 2);
        assert_eq!(pt.data_fields[0].subtotal, SubtotalFunction::Sum);
    }
}
