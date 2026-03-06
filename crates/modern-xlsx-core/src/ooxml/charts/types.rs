//! Chart type definitions — structs, enums, and small helper methods.

use serde::{Deserialize, Serialize};

/// The type of chart.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum ChartType {
    Bar,
    Column,
    Line,
    Pie,
    Doughnut,
    Scatter,
    Area,
    Radar,
    Bubble,
    Stock,
}

/// Bar/column grouping mode.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum ChartGrouping {
    Clustered,
    Stacked,
    PercentStacked,
    Standard,
}

/// Scatter chart style.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum ScatterStyle {
    LineMarker,
    Line,
    Marker,
    Smooth,
    SmoothMarker,
}

/// Radar chart style.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum RadarStyle {
    Standard,
    Marker,
    Filled,
}

/// A complete chart definition.
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ChartData {
    pub chart_type: ChartType,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub title: Option<ChartTitle>,
    pub series: Vec<ChartSeries>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cat_axis: Option<ChartAxis>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub val_axis: Option<ChartAxis>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub legend: Option<ChartLegend>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub data_labels: Option<DataLabels>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub grouping: Option<ChartGrouping>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub scatter_style: Option<ScatterStyle>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub radar_style: Option<RadarStyle>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub hole_size: Option<u32>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub bar_dir_horizontal: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub style_id: Option<u32>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub plot_area_layout: Option<ManualLayout>,
    /// Secondary chart for combo charts.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub secondary_chart: Option<Box<ChartData>>,
    /// Secondary value axis (for combo charts).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub secondary_val_axis: Option<ChartAxis>,
    /// Show data table below chart.
    #[serde(default, skip_serializing_if = "crate::ooxml::is_false")]
    pub show_data_table: bool,
    /// 3D rotation settings.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub view_3d: Option<View3D>,
}

/// A data series within a chart.
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ChartSeries {
    pub idx: u32,
    pub order: u32,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub name: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cat_ref: Option<String>,
    pub val_ref: String,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub x_val_ref: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub bubble_size_ref: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub fill_color: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub line_color: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub line_width: Option<u32>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub marker: Option<MarkerStyle>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub smooth: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub explosion: Option<u32>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub data_labels: Option<DataLabels>,
    /// Trendline.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub trendline: Option<Trendline>,
    /// Error bars.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub error_bars: Option<ErrorBars>,
}

/// Marker style for line/scatter series.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum MarkerStyle {
    Circle,
    Square,
    Diamond,
    Triangle,
    Star,
    X,
    Plus,
    Dash,
    Dot,
    None,
}

/// Chart title.
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ChartTitle {
    pub text: String,
    #[serde(default)]
    pub overlay: bool,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub font_size: Option<u32>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub bold: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub color: Option<String>,
}

/// Chart axis definition.
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ChartAxis {
    pub id: u32,
    pub cross_ax: u32,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub title: Option<ChartTitle>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub num_fmt: Option<String>,
    #[serde(default)]
    pub source_linked: bool,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub min: Option<f64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub max: Option<f64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub major_unit: Option<f64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub minor_unit: Option<f64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub log_base: Option<f64>,
    #[serde(default)]
    pub reversed: bool,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tick_lbl_pos: Option<TickLabelPosition>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub major_tick_mark: Option<TickMark>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub minor_tick_mark: Option<TickMark>,
    #[serde(default)]
    pub major_gridlines: bool,
    #[serde(default)]
    pub minor_gridlines: bool,
    #[serde(default)]
    pub delete: bool,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub position: Option<AxisPosition>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub crosses_at: Option<f64>,
    /// Font size for axis tick labels in hundredths of a point (e.g., 1400 = 14pt).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub font_size: Option<u32>,
}

/// Tick label position.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum TickLabelPosition {
    High,
    Low,
    NextTo,
    None,
}

/// Tick mark style.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum TickMark {
    Cross,
    In,
    Out,
    None,
}

/// Axis position.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum AxisPosition {
    Bottom,
    Top,
    Left,
    Right,
}

/// Legend configuration.
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ChartLegend {
    pub position: LegendPosition,
    #[serde(default)]
    pub overlay: bool,
}

/// Legend position.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum LegendPosition {
    Top,
    Bottom,
    Left,
    Right,
    TopRight,
}

/// Data labels configuration.
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct DataLabels {
    #[serde(default)]
    pub show_val: bool,
    #[serde(default)]
    pub show_cat_name: bool,
    #[serde(default)]
    pub show_ser_name: bool,
    #[serde(default)]
    pub show_percent: bool,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub num_fmt: Option<String>,
    #[serde(default)]
    pub show_leader_lines: bool,
}

/// Manual layout positioning (fractional 0.0-1.0).
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ManualLayout {
    pub x: f64,
    pub y: f64,
    pub w: f64,
    pub h: f64,
}

/// Trendline type for regression analysis.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum TrendlineType {
    Linear,
    Exponential,
    Logarithmic,
    Polynomial,
    Power,
    MovingAverage,
}

/// Trendline configuration for a chart series.
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct Trendline {
    pub trend_type: TrendlineType,
    /// Polynomial order (2-6).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub order: Option<u32>,
    /// Moving average period.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub period: Option<u32>,
    /// Forecast forward periods.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub forward: Option<f64>,
    /// Forecast backward periods.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub backward: Option<f64>,
    /// Display equation on chart.
    #[serde(default)]
    pub display_eq: bool,
    /// Display R-squared value on chart.
    #[serde(default)]
    pub display_r_sqr: bool,
}

/// Error bar type.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum ErrorBarType {
    FixedVal,
    Percentage,
    StdDev,
    StdErr,
    Custom,
}

/// Error bar direction.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum ErrorBarDirection {
    #[default]
    Both,
    Plus,
    Minus,
}

/// Error bars configuration for a chart series.
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ErrorBars {
    pub err_type: ErrorBarType,
    /// Direction: both, plus, or minus.
    #[serde(default)]
    pub direction: ErrorBarDirection,
    /// Value (for FixedVal or Percentage).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub value: Option<f64>,
}

/// 3D rotation settings.
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct View3D {
    /// X-axis rotation (-90 to 90).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub rot_x: Option<i32>,
    /// Y-axis rotation (0 to 360).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub rot_y: Option<i32>,
    /// Perspective (0 to 240).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub perspective: Option<u32>,
    /// Right-angle axes.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub r_ang_ax: Option<bool>,
}

/// Drawing anchor for positioning a chart on a worksheet.
///
/// For `twoCellAnchor`, `from_*` and `to_*` define the bounding box.
/// For `oneCellAnchor`, only `from_*` is meaningful and `ext_cx`/`ext_cy`
/// specify the width/height in EMUs. The `to_*` fields are set to 0.
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ChartAnchor {
    pub from_col: u32,
    pub from_row: u32,
    #[serde(default)]
    pub from_col_off: u64,
    #[serde(default)]
    pub from_row_off: u64,
    pub to_col: u32,
    pub to_row: u32,
    #[serde(default)]
    pub to_col_off: u64,
    #[serde(default)]
    pub to_row_off: u64,
    /// Width in EMUs (for oneCellAnchor). When `Some`, the anchor is a oneCellAnchor.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ext_cx: Option<u64>,
    /// Height in EMUs (for oneCellAnchor). When `Some`, the anchor is a oneCellAnchor.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ext_cy: Option<u64>,
}

// ---------------------------------------------------------------------------
// XML helper methods on enums
// ---------------------------------------------------------------------------

impl ChartGrouping {
    #[inline]
    pub(super) fn xml_val(self) -> &'static str {
        match self {
            Self::Clustered => "clustered",
            Self::Stacked => "stacked",
            Self::PercentStacked => "percentStacked",
            Self::Standard => "standard",
        }
    }

    #[inline]
    pub(super) fn from_xml(s: &str) -> Option<Self> {
        match s {
            "clustered" => Some(Self::Clustered),
            "stacked" => Some(Self::Stacked),
            "percentStacked" => Some(Self::PercentStacked),
            "standard" => Some(Self::Standard),
            _ => None,
        }
    }
}

impl ScatterStyle {
    #[inline]
    pub(super) fn xml_val(self) -> &'static str {
        match self {
            Self::LineMarker => "lineMarker",
            Self::Line => "line",
            Self::Marker => "marker",
            Self::Smooth => "smooth",
            Self::SmoothMarker => "smoothMarker",
        }
    }

    #[inline]
    pub(super) fn from_xml(s: &str) -> Option<Self> {
        match s {
            "lineMarker" => Some(Self::LineMarker),
            "line" => Some(Self::Line),
            "marker" => Some(Self::Marker),
            "smooth" => Some(Self::Smooth),
            "smoothMarker" => Some(Self::SmoothMarker),
            _ => None,
        }
    }
}

impl RadarStyle {
    #[inline]
    pub(super) fn xml_val(self) -> &'static str {
        match self {
            Self::Standard => "standard",
            Self::Marker => "marker",
            Self::Filled => "filled",
        }
    }

    #[inline]
    pub(super) fn from_xml(s: &str) -> Option<Self> {
        match s {
            "standard" => Some(Self::Standard),
            "marker" => Some(Self::Marker),
            "filled" => Some(Self::Filled),
            _ => None,
        }
    }
}

impl MarkerStyle {
    #[inline]
    pub(super) fn xml_val(self) -> &'static str {
        match self {
            Self::Circle => "circle",
            Self::Square => "square",
            Self::Diamond => "diamond",
            Self::Triangle => "triangle",
            Self::Star => "star",
            Self::X => "x",
            Self::Plus => "plus",
            Self::Dash => "dash",
            Self::Dot => "dot",
            Self::None => "none",
        }
    }

    #[inline]
    pub(super) fn from_xml(s: &str) -> Option<Self> {
        match s {
            "circle" => Some(Self::Circle),
            "square" => Some(Self::Square),
            "diamond" => Some(Self::Diamond),
            "triangle" => Some(Self::Triangle),
            "star" => Some(Self::Star),
            "x" => Some(Self::X),
            "plus" => Some(Self::Plus),
            "dash" => Some(Self::Dash),
            "dot" => Some(Self::Dot),
            "none" => Some(Self::None),
            _ => None,
        }
    }
}

impl LegendPosition {
    #[inline]
    pub(super) fn xml_val(self) -> &'static str {
        match self {
            Self::Top => "t",
            Self::Bottom => "b",
            Self::Left => "l",
            Self::Right => "r",
            Self::TopRight => "tr",
        }
    }

    #[inline]
    pub(super) fn from_xml(s: &str) -> Option<Self> {
        match s {
            "t" => Some(Self::Top),
            "b" => Some(Self::Bottom),
            "l" => Some(Self::Left),
            "r" => Some(Self::Right),
            "tr" => Some(Self::TopRight),
            _ => None,
        }
    }
}

impl AxisPosition {
    #[inline]
    pub(super) fn xml_val(self) -> &'static str {
        match self {
            Self::Bottom => "b",
            Self::Top => "t",
            Self::Left => "l",
            Self::Right => "r",
        }
    }

    #[inline]
    pub(super) fn from_xml(s: &str) -> Option<Self> {
        match s {
            "b" => Some(Self::Bottom),
            "t" => Some(Self::Top),
            "l" => Some(Self::Left),
            "r" => Some(Self::Right),
            _ => None,
        }
    }
}

impl TickLabelPosition {
    #[inline]
    pub(super) fn xml_val(self) -> &'static str {
        match self {
            Self::High => "high",
            Self::Low => "low",
            Self::NextTo => "nextTo",
            Self::None => "none",
        }
    }

    #[inline]
    pub(super) fn from_xml(s: &str) -> Option<Self> {
        match s {
            "high" => Some(Self::High),
            "low" => Some(Self::Low),
            "nextTo" => Some(Self::NextTo),
            "none" => Some(Self::None),
            _ => None,
        }
    }
}

impl TickMark {
    #[inline]
    pub(super) fn xml_val(self) -> &'static str {
        match self {
            Self::Cross => "cross",
            Self::In => "in",
            Self::Out => "out",
            Self::None => "none",
        }
    }

    #[inline]
    pub(super) fn from_xml(s: &str) -> Option<Self> {
        match s {
            "cross" => Some(Self::Cross),
            "in" => Some(Self::In),
            "out" => Some(Self::Out),
            "none" => Some(Self::None),
            _ => None,
        }
    }
}

impl TrendlineType {
    #[inline]
    pub(super) fn xml_val(self) -> &'static str {
        match self {
            Self::Linear => "linear",
            Self::Exponential => "exp",
            Self::Logarithmic => "log",
            Self::Polynomial => "poly",
            Self::Power => "power",
            Self::MovingAverage => "movingAvg",
        }
    }

    #[inline]
    pub(super) fn from_xml(s: &str) -> Option<Self> {
        match s {
            "linear" => Some(Self::Linear),
            "exp" => Some(Self::Exponential),
            "log" => Some(Self::Logarithmic),
            "poly" => Some(Self::Polynomial),
            "power" => Some(Self::Power),
            "movingAvg" => Some(Self::MovingAverage),
            _ => None,
        }
    }
}

impl ErrorBarType {
    #[inline]
    pub(super) fn xml_val(self) -> &'static str {
        match self {
            Self::FixedVal => "fixedVal",
            Self::Percentage => "percentage",
            Self::StdDev => "stdDev",
            Self::StdErr => "stdErr",
            Self::Custom => "cust",
        }
    }

    #[inline]
    pub(super) fn from_xml(s: &str) -> Option<Self> {
        match s {
            "fixedVal" => Some(Self::FixedVal),
            "percentage" => Some(Self::Percentage),
            "stdDev" => Some(Self::StdDev),
            "stdErr" => Some(Self::StdErr),
            "cust" => Some(Self::Custom),
            _ => None,
        }
    }
}

impl ErrorBarDirection {
    #[inline]
    pub(super) fn xml_val(self) -> &'static str {
        match self {
            Self::Both => "both",
            Self::Plus => "plus",
            Self::Minus => "minus",
        }
    }

    #[inline]
    pub(super) fn from_xml(s: &str) -> Option<Self> {
        match s {
            "both" => Some(Self::Both),
            "plus" => Some(Self::Plus),
            "minus" => Some(Self::Minus),
            _ => None,
        }
    }
}
