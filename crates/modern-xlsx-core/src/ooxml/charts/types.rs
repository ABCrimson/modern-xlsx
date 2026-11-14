//! Chart type definitions — structs, enums, and small helper methods.

use serde::{Deserialize, Serialize};

/// The type of chart.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum ChartType {
    /// Also the fallback the chart parser uses when no chart-type element is recognised.
    #[default]
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

/// The `c:grouping` value of a chart-type element.
///
/// Bar and column charts use `ST_BarGrouping` (`clustered`, `stacked`,
/// `percentStacked`, `standard`); line and area charts use `ST_Grouping`
/// (`standard`, `stacked`, `percentStacked`), so `Clustered` is a bar/column
/// value only. The writer emits `c:grouping` for bar, column, line and area
/// charts; the parser reads it from the primary and secondary (combo) chart-type elements it models.
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
#[derive(Debug, Clone, Default, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ChartData {
    pub chart_type: ChartType,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub title: Option<ChartTitle>,
    pub series: Vec<ChartSeries>,
    /// Primary category axis (`c:catAx`). For a scatter or bubble chart, whose two axes are
    /// both value axes, a set `cat_axis` is the X axis and is written as `c:valAx`, with
    /// `val_axis` as the Y axis.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cat_axis: Option<ChartAxis>,
    /// Primary value axis (`c:valAx`). For a scatter or bubble chart read from a file
    /// (`cat_axis` unset) this is the X axis, the valAx named first by the chart's `c:axId`s.
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
    /// Secondary chart for combo charts: one chart-type element written after the primary
    /// one, on `secondary_cat_axis` and `secondary_val_axis`. An unset slot shares the
    /// primary chart's axis in that slot (`cat_axis` / `val_axis`; for a scatter or bubble
    /// primary read from a file, `val_axis` / `secondary_val_axis`), else the default axis
    /// [`ChartData::to_xml`] writes, also under a pie or doughnut primary. Its own axis fields
    /// are not written, and a `secondary_chart` of its own makes [`ChartData::to_xml`] fail.
    /// Its series share one [`ChartSeries::order`] numbering with the primary chart's.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub secondary_chart: Option<Box<ChartData>>,
    /// Secondary value axis of a combo chart. For a scatter or bubble chart read from a
    /// file with two `c:valAx` (`cat_axis` unset) this is the Y axis, the valAx named
    /// second by the chart's `c:axId`s.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub secondary_val_axis: Option<ChartAxis>,
    /// Secondary category axis of a combo chart, written as a second `c:catAx` (Excel's
    /// secondary-axis combos carry one, usually with `delete` set).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub secondary_cat_axis: Option<ChartAxis>,
    /// Show data table below chart.
    #[serde(default, skip_serializing_if = "crate::ooxml::is_false")]
    pub show_data_table: bool,
    /// 3D rotation settings.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub view_3d: Option<View3D>,
    /// `c:autoTitleDeleted` flag. When `None` the writer emits `1` for a chart without a
    /// title and `0` for a chart with one, so Excel does not add an automatic title.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub auto_title_deleted: Option<bool>,
    /// The chart part exactly as read, kept when the file holds one of these, which the model
    /// cannot represent: a 3-D, surface or bar/pie-of-pie chart type, a third chart-type element,
    /// an axis the model has no slot for (a date or series axis, or a category or value
    /// axis beyond the modelled ones), or a `c:userShapes` or `c:externalData` reference
    /// (an `r:id` the model cannot hold). Other unmodelled elements are dropped without it.
    /// While it is set the writer emits these bytes
    /// unchanged and ignores every other field, so edits to the model have no effect;
    /// remove `raw_xml` to write the chart from the model instead, and with it the
    /// `raw_rels` relationships only these bytes referenced.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub raw_xml: Option<String>,
    /// The chart part's `.rels` as read, for every chart that has one, modelled or
    /// kept in `raw_xml`. The writer places it beside the chart at whatever path the
    /// chart is written to, so the parts it references (styles, colours, user shapes,
    /// embedded data) stay linked to their own chart.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub raw_rels: Option<String>,
}

/// A data series within a chart.
#[derive(Debug, Clone, Default, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ChartSeries {
    /// `c:idx`, written and read as given. Excel opens a chart whose series share it.
    pub idx: u32,
    /// `c:order`, the series' place in the plotting order. It must be unique across the
    /// series of the primary and the secondary chart: Excel refuses a file where two share
    /// it, so such a chart fails the write with `InvalidWorkbook` and `validate_workbook`
    /// reports it. Written and read as given, never renumbered.
    pub order: u32,
    /// Series name as a cell-reference formula (`c:tx/c:strRef/c:f`, for example
    /// `Sheet1!$B$1`), never a literal. Written in preference to `name_literal`.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub name: Option<String>,
    /// Series name as literal text (`c:tx/c:v`), written when `name` is unset.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub name_literal: Option<String>,
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
#[derive(Debug, Clone, Default, PartialEq, Serialize, Deserialize)]
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
#[derive(Debug, Clone, Default, PartialEq, Serialize, Deserialize)]
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
    /// `c:crossBetween` of a value axis (`c:valAx` only): whether the value axis
    /// crosses the category axis between categories or at their midpoint.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cross_between: Option<CrossBetween>,
}

/// `c:crossBetween` (ST_CrossBetween).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum CrossBetween {
    Between,
    MidCat,
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
#[derive(Debug, Clone, Default, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ChartLegend {
    pub position: LegendPosition,
    #[serde(default)]
    pub overlay: bool,
}

/// Legend position.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum LegendPosition {
    Top,
    Bottom,
    Left,
    /// Schema default (`r`), also the fallback the chart parser uses.
    #[default]
    Right,
    TopRight,
}

/// Data labels configuration.
#[derive(Debug, Clone, Default, PartialEq, Serialize, Deserialize)]
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
    /// `c:showLegendKey`; `None` when the element is absent.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub show_legend_key: Option<bool>,
    /// `c:showBubbleSize`; `None` when the element is absent.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub show_bubble_size: Option<bool>,
}

/// Manual layout positioning (fractional 0.0-1.0).
#[derive(Debug, Clone, Default, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ManualLayout {
    pub x: f64,
    pub y: f64,
    pub w: f64,
    pub h: f64,
}

/// Trendline type for regression analysis.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum TrendlineType {
    /// Schema default.
    #[default]
    Linear,
    Exponential,
    Logarithmic,
    Polynomial,
    Power,
    MovingAverage,
}

/// Trendline configuration for a chart series.
#[derive(Debug, Clone, Default, PartialEq, Serialize, Deserialize)]
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
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum ErrorBarType {
    /// Schema default (`fixedVal`).
    #[default]
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

/// Axis an error bar applies to (`c:errDir`, ST_ErrDir).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum ErrorBarAxis {
    X,
    Y,
}

/// Error bars configuration for a chart series.
#[derive(Debug, Clone, Default, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ErrorBars {
    /// `c:errDir`: the axis the bars follow; `None` when the element is absent.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub err_dir: Option<ErrorBarAxis>,
    pub err_type: ErrorBarType,
    /// Direction: both, plus, or minus.
    #[serde(default)]
    pub direction: ErrorBarDirection,
    /// Value (for FixedVal or Percentage).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub value: Option<f64>,
    /// Range of the positive custom error amounts (`c:plus`), used with
    /// [`ErrorBarType::Custom`].
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub plus_ref: Option<String>,
    /// Range of the negative custom error amounts (`c:minus`), used with
    /// [`ErrorBarType::Custom`].
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub minus_ref: Option<String>,
}

/// 3D rotation settings.
#[derive(Debug, Clone, Default, PartialEq, Serialize, Deserialize)]
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
/// For `absoluteAnchor`, `pos_x`/`pos_y` give the position and `ext_cx`/`ext_cy` the
/// size, all in EMUs; the cell fields are 0.
#[derive(Debug, Clone, Default, PartialEq, Serialize, Deserialize)]
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
    /// Width in EMUs, for a oneCellAnchor or absoluteAnchor. With `ext_cy` also set the writer emits a oneCellAnchor, or an absoluteAnchor when `pos_x`/`pos_y` are set too.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ext_cx: Option<u64>,
    /// Height in EMUs. See `ext_cx`.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ext_cy: Option<u64>,
    /// Absolute X position in EMUs. With `pos_y` (and `ext_cx`/`ext_cy`) the writer emits an
    /// `xdr:absoluteAnchor`.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub pos_x: Option<i64>,
    /// Absolute Y position in EMUs. See `pos_x`.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub pos_y: Option<i64>,
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

impl ErrorBarAxis {
    #[inline]
    pub(super) fn xml_val(self) -> &'static str {
        match self {
            Self::X => "x",
            Self::Y => "y",
        }
    }

    #[inline]
    pub(super) fn from_xml(s: &str) -> Option<Self> {
        match s {
            "x" => Some(Self::X),
            "y" => Some(Self::Y),
            _ => None,
        }
    }
}

impl CrossBetween {
    #[inline]
    pub(super) fn xml_val(self) -> &'static str {
        match self {
            Self::Between => "between",
            Self::MidCat => "midCat",
        }
    }

    #[inline]
    pub(super) fn from_xml(s: &str) -> Option<Self> {
        match s {
            "between" => Some(Self::Between),
            "midCat" => Some(Self::MidCat),
            _ => None,
        }
    }
}
