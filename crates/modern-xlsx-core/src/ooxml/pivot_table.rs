use serde::{Deserialize, Serialize};

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
