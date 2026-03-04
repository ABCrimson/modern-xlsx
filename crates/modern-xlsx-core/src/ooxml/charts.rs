//! Chart definitions — `xl/charts/chart{n}.xml`.

use quick_xml::events::{BytesDecl, BytesEnd, BytesStart, BytesText, Event};
use quick_xml::Writer;
use serde::{Deserialize, Serialize};

use crate::{ModernXlsxError, Result};

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

/// Drawing anchor for positioning a chart on a worksheet.
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
}

/// A chart embedded in a worksheet, with its anchor position.
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct WorksheetChart {
    pub chart: ChartData,
    pub anchor: ChartAnchor,
}

// ---------------------------------------------------------------------------
// XML Writer helpers
// ---------------------------------------------------------------------------

/// Map `std::io::Error` to `ModernXlsxError::XmlWrite`.
fn map_err(e: std::io::Error) -> ModernXlsxError {
    ModernXlsxError::XmlWrite(e.to_string())
}

impl ChartGrouping {
    fn xml_val(self) -> &'static str {
        match self {
            Self::Clustered => "clustered",
            Self::Stacked => "stacked",
            Self::PercentStacked => "percentStacked",
            Self::Standard => "standard",
        }
    }
}

impl ScatterStyle {
    fn xml_val(self) -> &'static str {
        match self {
            Self::LineMarker => "lineMarker",
            Self::Line => "line",
            Self::Marker => "marker",
            Self::Smooth => "smooth",
            Self::SmoothMarker => "smoothMarker",
        }
    }
}

impl RadarStyle {
    fn xml_val(self) -> &'static str {
        match self {
            Self::Standard => "standard",
            Self::Marker => "marker",
            Self::Filled => "filled",
        }
    }
}

impl MarkerStyle {
    fn xml_val(self) -> &'static str {
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
}

impl LegendPosition {
    fn xml_val(self) -> &'static str {
        match self {
            Self::Top => "t",
            Self::Bottom => "b",
            Self::Left => "l",
            Self::Right => "r",
            Self::TopRight => "tr",
        }
    }
}

impl AxisPosition {
    fn xml_val(self) -> &'static str {
        match self {
            Self::Bottom => "b",
            Self::Top => "t",
            Self::Left => "l",
            Self::Right => "r",
        }
    }
}

impl TickLabelPosition {
    fn xml_val(self) -> &'static str {
        match self {
            Self::High => "high",
            Self::Low => "low",
            Self::NextTo => "nextTo",
            Self::None => "none",
        }
    }
}

impl TickMark {
    fn xml_val(self) -> &'static str {
        match self {
            Self::Cross => "cross",
            Self::In => "in",
            Self::Out => "out",
            Self::None => "none",
        }
    }
}

impl ChartData {
    /// Serialize this chart to valid `xl/charts/chart{n}.xml` bytes.
    pub fn to_xml(&self) -> Result<Vec<u8>> {
        let mut buf: Vec<u8> = Vec::with_capacity(1024);
        let mut writer = Writer::new(&mut buf);

        // XML declaration
        writer
            .write_event(Event::Decl(BytesDecl::new("1.0", Some("UTF-8"), Some("yes"))))
            .map_err(map_err)?;

        // <c:chartSpace>
        let mut cs = BytesStart::new("c:chartSpace");
        cs.push_attribute((
            "xmlns:c",
            "http://schemas.openxmlformats.org/drawingml/2006/chart",
        ));
        cs.push_attribute((
            "xmlns:a",
            "http://schemas.openxmlformats.org/drawingml/2006/main",
        ));
        cs.push_attribute((
            "xmlns:r",
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
        ));
        writer.write_event(Event::Start(cs)).map_err(map_err)?;

        // <c:chart>
        writer
            .write_event(Event::Start(BytesStart::new("c:chart")))
            .map_err(map_err)?;

        // <c:title> (chart-level)
        if let Some(ref title) = self.title {
            Self::write_title(&mut writer, title)?;
        }

        // <c:plotArea>
        writer
            .write_event(Event::Start(BytesStart::new("c:plotArea")))
            .map_err(map_err)?;

        // <c:layout> or <c:layout/>
        self.write_layout(&mut writer)?;

        // Chart-type-specific element
        self.write_chart_type_element(&mut writer)?;

        // Axes
        if let Some(ref axis) = self.cat_axis {
            Self::write_axis(&mut writer, "c:catAx", axis)?;
        }
        if let Some(ref axis) = self.val_axis {
            Self::write_axis(&mut writer, "c:valAx", axis)?;
        }

        // </c:plotArea>
        writer
            .write_event(Event::End(BytesEnd::new("c:plotArea")))
            .map_err(map_err)?;

        // <c:legend>
        if let Some(ref legend) = self.legend {
            Self::write_legend(&mut writer, legend)?;
        }

        // </c:chart>
        writer
            .write_event(Event::End(BytesEnd::new("c:chart")))
            .map_err(map_err)?;

        // <c:style>
        if let Some(style_id) = self.style_id {
            let mut ibuf = itoa::Buffer::new();
            let mut style = BytesStart::new("c:style");
            style.push_attribute(("val", ibuf.format(style_id)));
            writer.write_event(Event::Empty(style)).map_err(map_err)?;
        }

        // <c:printSettings>
        Self::write_print_settings(&mut writer)?;

        // </c:chartSpace>
        writer
            .write_event(Event::End(BytesEnd::new("c:chartSpace")))
            .map_err(map_err)?;

        Ok(buf)
    }

    // -----------------------------------------------------------------------
    // Layout
    // -----------------------------------------------------------------------

    fn write_layout(&self, writer: &mut Writer<&mut Vec<u8>>) -> Result<()> {
        if let Some(ref layout) = self.plot_area_layout {
            writer
                .write_event(Event::Start(BytesStart::new("c:layout")))
                .map_err(map_err)?;

            writer
                .write_event(Event::Start(BytesStart::new("c:manualLayout")))
                .map_err(map_err)?;

            Self::write_f64_element(writer, "c:x", layout.x)?;
            Self::write_f64_element(writer, "c:y", layout.y)?;
            Self::write_f64_element(writer, "c:w", layout.w)?;
            Self::write_f64_element(writer, "c:h", layout.h)?;

            writer
                .write_event(Event::End(BytesEnd::new("c:manualLayout")))
                .map_err(map_err)?;
            writer
                .write_event(Event::End(BytesEnd::new("c:layout")))
                .map_err(map_err)?;
        } else {
            writer
                .write_event(Event::Empty(BytesStart::new("c:layout")))
                .map_err(map_err)?;
        }
        Ok(())
    }

    // -----------------------------------------------------------------------
    // Chart-type-specific element
    // -----------------------------------------------------------------------

    fn write_chart_type_element(&self, writer: &mut Writer<&mut Vec<u8>>) -> Result<()> {
        let tag = self.chart_type_xml_tag();

        writer
            .write_event(Event::Start(BytesStart::new(tag)))
            .map_err(map_err)?;

        let mut ibuf = itoa::Buffer::new();

        // Bar direction (bar/column)
        if matches!(self.chart_type, ChartType::Bar | ChartType::Column) {
            let dir = if self.chart_type == ChartType::Bar
                || self.bar_dir_horizontal == Some(true)
            {
                "bar"
            } else {
                "col"
            };
            let mut bd = BytesStart::new("c:barDir");
            bd.push_attribute(("val", dir));
            writer.write_event(Event::Empty(bd)).map_err(map_err)?;
        }

        // Grouping
        if let Some(grouping) = self.grouping
            && matches!(
                self.chart_type,
                ChartType::Bar
                    | ChartType::Column
                    | ChartType::Line
                    | ChartType::Area
            )
        {
            let mut g = BytesStart::new("c:grouping");
            g.push_attribute(("val", grouping.xml_val()));
            writer.write_event(Event::Empty(g)).map_err(map_err)?;
        }

        // Scatter style
        if let Some(style) = self.scatter_style
            && self.chart_type == ChartType::Scatter
        {
            let mut ss = BytesStart::new("c:scatterStyle");
            ss.push_attribute(("val", style.xml_val()));
            writer.write_event(Event::Empty(ss)).map_err(map_err)?;
        }

        // Radar style
        if let Some(style) = self.radar_style
            && self.chart_type == ChartType::Radar
        {
            let mut rs = BytesStart::new("c:radarStyle");
            rs.push_attribute(("val", style.xml_val()));
            writer.write_event(Event::Empty(rs)).map_err(map_err)?;
        }

        // Series
        let uses_xy =
            matches!(self.chart_type, ChartType::Scatter | ChartType::Bubble);

        for ser in &self.series {
            self.write_series(writer, ser, uses_xy, &mut ibuf)?;
        }

        // Chart-level data labels
        if let Some(ref dl) = self.data_labels {
            Self::write_data_labels(writer, dl)?;
        }

        // Hole size for doughnut
        if let Some(hole_size) = self.hole_size
            && self.chart_type == ChartType::Doughnut
        {
            let mut hs = BytesStart::new("c:holeSize");
            hs.push_attribute(("val", ibuf.format(hole_size)));
            writer.write_event(Event::Empty(hs)).map_err(map_err)?;
        }

        // Axis IDs for chart types that have axes
        if self.has_axes() {
            if let Some(ref cat_ax) = self.cat_axis {
                let mut id = BytesStart::new("c:axId");
                id.push_attribute(("val", ibuf.format(cat_ax.id)));
                writer.write_event(Event::Empty(id)).map_err(map_err)?;
            }
            if let Some(ref val_ax) = self.val_axis {
                let mut id = BytesStart::new("c:axId");
                id.push_attribute(("val", ibuf.format(val_ax.id)));
                writer.write_event(Event::Empty(id)).map_err(map_err)?;
            }
        }

        writer
            .write_event(Event::End(BytesEnd::new(tag)))
            .map_err(map_err)?;

        Ok(())
    }

    fn chart_type_xml_tag(&self) -> &'static str {
        match self.chart_type {
            ChartType::Bar | ChartType::Column => "c:barChart",
            ChartType::Line => "c:lineChart",
            ChartType::Pie => "c:pieChart",
            ChartType::Doughnut => "c:doughnutChart",
            ChartType::Scatter => "c:scatterChart",
            ChartType::Area => "c:areaChart",
            ChartType::Radar => "c:radarChart",
            ChartType::Bubble => "c:bubbleChart",
            ChartType::Stock => "c:stockChart",
        }
    }

    fn has_axes(&self) -> bool {
        !matches!(self.chart_type, ChartType::Pie | ChartType::Doughnut)
    }

    // -----------------------------------------------------------------------
    // Series
    // -----------------------------------------------------------------------

    fn write_series(
        &self,
        writer: &mut Writer<&mut Vec<u8>>,
        ser: &ChartSeries,
        uses_xy: bool,
        ibuf: &mut itoa::Buffer,
    ) -> Result<()> {
        writer
            .write_event(Event::Start(BytesStart::new("c:ser")))
            .map_err(map_err)?;

        // <c:idx val="0"/>
        let mut idx = BytesStart::new("c:idx");
        idx.push_attribute(("val", ibuf.format(ser.idx)));
        writer.write_event(Event::Empty(idx)).map_err(map_err)?;

        // <c:order val="0"/>
        let mut ord = BytesStart::new("c:order");
        ord.push_attribute(("val", ibuf.format(ser.order)));
        writer.write_event(Event::Empty(ord)).map_err(map_err)?;

        // <c:tx>
        if let Some(ref name) = ser.name {
            writer
                .write_event(Event::Start(BytesStart::new("c:tx")))
                .map_err(map_err)?;
            writer
                .write_event(Event::Start(BytesStart::new("c:strRef")))
                .map_err(map_err)?;
            writer
                .write_event(Event::Start(BytesStart::new("c:f")))
                .map_err(map_err)?;
            writer
                .write_event(Event::Text(BytesText::new(name)))
                .map_err(map_err)?;
            writer
                .write_event(Event::End(BytesEnd::new("c:f")))
                .map_err(map_err)?;
            writer
                .write_event(Event::End(BytesEnd::new("c:strRef")))
                .map_err(map_err)?;
            writer
                .write_event(Event::End(BytesEnd::new("c:tx")))
                .map_err(map_err)?;
        }

        // <c:spPr>
        Self::write_sp_pr(writer, ser)?;

        // <c:marker>
        if let Some(marker) = ser.marker {
            Self::write_marker(writer, marker)?;
        }

        // Explosion (pie/doughnut)
        if let Some(explosion) = ser.explosion {
            let mut exp = BytesStart::new("c:explosion");
            exp.push_attribute(("val", ibuf.format(explosion)));
            writer.write_event(Event::Empty(exp)).map_err(map_err)?;
        }

        if uses_xy {
            // Scatter/Bubble: xVal + yVal
            if let Some(ref x_ref) = ser.x_val_ref {
                Self::write_ref_element(writer, "c:xVal", "c:numRef", x_ref)?;
            }
            Self::write_ref_element(writer, "c:yVal", "c:numRef", &ser.val_ref)?;

            // Bubble size
            if let Some(ref bub_ref) = ser.bubble_size_ref {
                Self::write_ref_element(writer, "c:bubbleSize", "c:numRef", bub_ref)?;
            }
        } else {
            // Standard: cat + val
            if let Some(ref cat_ref) = ser.cat_ref {
                Self::write_ref_element(writer, "c:cat", "c:strRef", cat_ref)?;
            }
            Self::write_ref_element(writer, "c:val", "c:numRef", &ser.val_ref)?;
        }

        // <c:smooth>
        if let Some(smooth) = ser.smooth {
            let mut sm = BytesStart::new("c:smooth");
            sm.push_attribute(("val", if smooth { "1" } else { "0" }));
            writer.write_event(Event::Empty(sm)).map_err(map_err)?;
        }

        // Series-level data labels
        if let Some(ref dl) = ser.data_labels {
            Self::write_data_labels(writer, dl)?;
        }

        writer
            .write_event(Event::End(BytesEnd::new("c:ser")))
            .map_err(map_err)?;

        Ok(())
    }

    // -----------------------------------------------------------------------
    // Shape properties (fill + line)
    // -----------------------------------------------------------------------

    fn write_sp_pr(
        writer: &mut Writer<&mut Vec<u8>>,
        ser: &ChartSeries,
    ) -> Result<()> {
        if ser.fill_color.is_none() && ser.line_color.is_none() {
            return Ok(());
        }

        writer
            .write_event(Event::Start(BytesStart::new("c:spPr")))
            .map_err(map_err)?;

        // Fill
        if let Some(ref fill) = ser.fill_color {
            writer
                .write_event(Event::Start(BytesStart::new("a:solidFill")))
                .map_err(map_err)?;
            let mut clr = BytesStart::new("a:srgbClr");
            clr.push_attribute(("val", fill.as_str()));
            writer.write_event(Event::Empty(clr)).map_err(map_err)?;
            writer
                .write_event(Event::End(BytesEnd::new("a:solidFill")))
                .map_err(map_err)?;
        }

        // Line
        if let Some(ref line_color) = ser.line_color {
            let mut ln = BytesStart::new("a:ln");
            if let Some(w) = ser.line_width {
                let mut ibuf = itoa::Buffer::new();
                ln.push_attribute(("w", ibuf.format(w)));
            }
            writer.write_event(Event::Start(ln)).map_err(map_err)?;
            writer
                .write_event(Event::Start(BytesStart::new("a:solidFill")))
                .map_err(map_err)?;
            let mut clr = BytesStart::new("a:srgbClr");
            clr.push_attribute(("val", line_color.as_str()));
            writer.write_event(Event::Empty(clr)).map_err(map_err)?;
            writer
                .write_event(Event::End(BytesEnd::new("a:solidFill")))
                .map_err(map_err)?;
            writer
                .write_event(Event::End(BytesEnd::new("a:ln")))
                .map_err(map_err)?;
        }

        writer
            .write_event(Event::End(BytesEnd::new("c:spPr")))
            .map_err(map_err)?;

        Ok(())
    }

    // -----------------------------------------------------------------------
    // Marker
    // -----------------------------------------------------------------------

    fn write_marker(
        writer: &mut Writer<&mut Vec<u8>>,
        marker: MarkerStyle,
    ) -> Result<()> {
        writer
            .write_event(Event::Start(BytesStart::new("c:marker")))
            .map_err(map_err)?;

        let mut sym = BytesStart::new("c:symbol");
        sym.push_attribute(("val", marker.xml_val()));
        writer.write_event(Event::Empty(sym)).map_err(map_err)?;

        let mut sz = BytesStart::new("c:size");
        sz.push_attribute(("val", "5"));
        writer.write_event(Event::Empty(sz)).map_err(map_err)?;

        writer
            .write_event(Event::End(BytesEnd::new("c:marker")))
            .map_err(map_err)?;

        Ok(())
    }

    // -----------------------------------------------------------------------
    // Reference elements (cat/val/xVal/yVal/bubbleSize)
    // -----------------------------------------------------------------------

    fn write_ref_element(
        writer: &mut Writer<&mut Vec<u8>>,
        outer_tag: &str,
        ref_tag: &str,
        formula: &str,
    ) -> Result<()> {
        writer
            .write_event(Event::Start(BytesStart::new(outer_tag)))
            .map_err(map_err)?;
        writer
            .write_event(Event::Start(BytesStart::new(ref_tag)))
            .map_err(map_err)?;
        writer
            .write_event(Event::Start(BytesStart::new("c:f")))
            .map_err(map_err)?;
        writer
            .write_event(Event::Text(BytesText::new(formula)))
            .map_err(map_err)?;
        writer
            .write_event(Event::End(BytesEnd::new("c:f")))
            .map_err(map_err)?;
        writer
            .write_event(Event::End(BytesEnd::new(ref_tag)))
            .map_err(map_err)?;
        writer
            .write_event(Event::End(BytesEnd::new(outer_tag)))
            .map_err(map_err)?;

        Ok(())
    }

    // -----------------------------------------------------------------------
    // Title
    // -----------------------------------------------------------------------

    fn write_title(
        writer: &mut Writer<&mut Vec<u8>>,
        title: &ChartTitle,
    ) -> Result<()> {
        let mut ibuf = itoa::Buffer::new();

        writer
            .write_event(Event::Start(BytesStart::new("c:title")))
            .map_err(map_err)?;
        writer
            .write_event(Event::Start(BytesStart::new("c:tx")))
            .map_err(map_err)?;
        writer
            .write_event(Event::Start(BytesStart::new("c:rich")))
            .map_err(map_err)?;

        // <a:bodyPr/>
        writer
            .write_event(Event::Empty(BytesStart::new("a:bodyPr")))
            .map_err(map_err)?;
        // <a:lstStyle/>
        writer
            .write_event(Event::Empty(BytesStart::new("a:lstStyle")))
            .map_err(map_err)?;

        // <a:p>
        writer
            .write_event(Event::Start(BytesStart::new("a:p")))
            .map_err(map_err)?;

        let font_size = title.font_size.unwrap_or(1400);
        let bold = title.bold.unwrap_or(true);

        // <a:pPr>
        writer
            .write_event(Event::Start(BytesStart::new("a:pPr")))
            .map_err(map_err)?;

        // <a:defRPr>
        let mut def_rpr = BytesStart::new("a:defRPr");
        def_rpr.push_attribute(("sz", ibuf.format(font_size)));
        if bold {
            def_rpr.push_attribute(("b", "1"));
        }
        if title.color.is_some() {
            writer
                .write_event(Event::Start(def_rpr))
                .map_err(map_err)?;
            if let Some(ref color) = title.color {
                Self::write_solid_fill(writer, color)?;
            }
            writer
                .write_event(Event::End(BytesEnd::new("a:defRPr")))
                .map_err(map_err)?;
        } else {
            writer
                .write_event(Event::Empty(def_rpr))
                .map_err(map_err)?;
        }

        // </a:pPr>
        writer
            .write_event(Event::End(BytesEnd::new("a:pPr")))
            .map_err(map_err)?;

        // <a:r>
        writer
            .write_event(Event::Start(BytesStart::new("a:r")))
            .map_err(map_err)?;

        // <a:rPr>
        let mut rpr = BytesStart::new("a:rPr");
        rpr.push_attribute(("lang", "en-US"));
        rpr.push_attribute(("sz", ibuf.format(font_size)));
        if bold {
            rpr.push_attribute(("b", "1"));
        }
        if title.color.is_some() {
            writer.write_event(Event::Start(rpr)).map_err(map_err)?;
            if let Some(ref color) = title.color {
                Self::write_solid_fill(writer, color)?;
            }
            writer
                .write_event(Event::End(BytesEnd::new("a:rPr")))
                .map_err(map_err)?;
        } else {
            writer.write_event(Event::Empty(rpr)).map_err(map_err)?;
        }

        // <a:t>text</a:t>
        writer
            .write_event(Event::Start(BytesStart::new("a:t")))
            .map_err(map_err)?;
        writer
            .write_event(Event::Text(BytesText::new(&title.text)))
            .map_err(map_err)?;
        writer
            .write_event(Event::End(BytesEnd::new("a:t")))
            .map_err(map_err)?;

        // </a:r>
        writer
            .write_event(Event::End(BytesEnd::new("a:r")))
            .map_err(map_err)?;

        // </a:p>
        writer
            .write_event(Event::End(BytesEnd::new("a:p")))
            .map_err(map_err)?;

        // </c:rich>
        writer
            .write_event(Event::End(BytesEnd::new("c:rich")))
            .map_err(map_err)?;
        // </c:tx>
        writer
            .write_event(Event::End(BytesEnd::new("c:tx")))
            .map_err(map_err)?;

        // <c:overlay val="0"/>
        let mut overlay = BytesStart::new("c:overlay");
        overlay.push_attribute(("val", if title.overlay { "1" } else { "0" }));
        writer
            .write_event(Event::Empty(overlay))
            .map_err(map_err)?;

        // </c:title>
        writer
            .write_event(Event::End(BytesEnd::new("c:title")))
            .map_err(map_err)?;

        Ok(())
    }

    fn write_solid_fill(
        writer: &mut Writer<&mut Vec<u8>>,
        color: &str,
    ) -> Result<()> {
        writer
            .write_event(Event::Start(BytesStart::new("a:solidFill")))
            .map_err(map_err)?;
        let mut clr = BytesStart::new("a:srgbClr");
        clr.push_attribute(("val", color));
        writer.write_event(Event::Empty(clr)).map_err(map_err)?;
        writer
            .write_event(Event::End(BytesEnd::new("a:solidFill")))
            .map_err(map_err)?;
        Ok(())
    }

    // -----------------------------------------------------------------------
    // Data labels
    // -----------------------------------------------------------------------

    fn write_data_labels(
        writer: &mut Writer<&mut Vec<u8>>,
        dl: &DataLabels,
    ) -> Result<()> {
        writer
            .write_event(Event::Start(BytesStart::new("c:dLbls")))
            .map_err(map_err)?;

        Self::write_bool_element(writer, "c:showVal", dl.show_val)?;
        Self::write_bool_element(writer, "c:showCatName", dl.show_cat_name)?;
        Self::write_bool_element(writer, "c:showSerName", dl.show_ser_name)?;
        Self::write_bool_element(writer, "c:showPercent", dl.show_percent)?;
        Self::write_bool_element(writer, "c:showLeaderLines", dl.show_leader_lines)?;

        if let Some(ref fmt) = dl.num_fmt {
            let mut nf = BytesStart::new("c:numFmt");
            nf.push_attribute(("formatCode", fmt.as_str()));
            nf.push_attribute(("sourceLinked", "0"));
            writer.write_event(Event::Empty(nf)).map_err(map_err)?;
        }

        writer
            .write_event(Event::End(BytesEnd::new("c:dLbls")))
            .map_err(map_err)?;

        Ok(())
    }

    // -----------------------------------------------------------------------
    // Axis
    // -----------------------------------------------------------------------

    fn write_axis(
        writer: &mut Writer<&mut Vec<u8>>,
        tag: &str,
        axis: &ChartAxis,
    ) -> Result<()> {
        let mut ibuf = itoa::Buffer::new();

        writer
            .write_event(Event::Start(BytesStart::new(tag)))
            .map_err(map_err)?;

        // <c:axId>
        let mut ax_id = BytesStart::new("c:axId");
        ax_id.push_attribute(("val", ibuf.format(axis.id)));
        writer
            .write_event(Event::Empty(ax_id))
            .map_err(map_err)?;

        // <c:scaling>
        writer
            .write_event(Event::Start(BytesStart::new("c:scaling")))
            .map_err(map_err)?;

        let mut orient = BytesStart::new("c:orientation");
        orient.push_attribute(("val", if axis.reversed { "maxMin" } else { "minMax" }));
        writer
            .write_event(Event::Empty(orient))
            .map_err(map_err)?;

        if let Some(min) = axis.min {
            Self::write_f64_element(writer, "c:min", min)?;
        }
        if let Some(max) = axis.max {
            Self::write_f64_element(writer, "c:max", max)?;
        }
        if let Some(log_base) = axis.log_base {
            Self::write_f64_element(writer, "c:logBase", log_base)?;
        }

        writer
            .write_event(Event::End(BytesEnd::new("c:scaling")))
            .map_err(map_err)?;

        // <c:delete>
        let mut del = BytesStart::new("c:delete");
        del.push_attribute(("val", if axis.delete { "1" } else { "0" }));
        writer.write_event(Event::Empty(del)).map_err(map_err)?;

        // <c:axPos>
        let default_pos = if tag == "c:catAx" { "b" } else { "l" };
        let pos = axis
            .position
            .map(|p| p.xml_val())
            .unwrap_or(default_pos);
        let mut ax_pos = BytesStart::new("c:axPos");
        ax_pos.push_attribute(("val", pos));
        writer
            .write_event(Event::Empty(ax_pos))
            .map_err(map_err)?;

        // Axis title
        if let Some(ref title) = axis.title {
            Self::write_title(writer, title)?;
        }

        // Gridlines
        if axis.major_gridlines {
            writer
                .write_event(Event::Empty(BytesStart::new("c:majorGridlines")))
                .map_err(map_err)?;
        }
        if axis.minor_gridlines {
            writer
                .write_event(Event::Empty(BytesStart::new("c:minorGridlines")))
                .map_err(map_err)?;
        }

        // Number format
        {
            let fmt_code = axis.num_fmt.as_deref().unwrap_or("General");
            let src_linked = if axis.source_linked { "1" } else { "0" };
            let mut nf = BytesStart::new("c:numFmt");
            nf.push_attribute(("formatCode", fmt_code));
            nf.push_attribute(("sourceLinked", src_linked));
            writer.write_event(Event::Empty(nf)).map_err(map_err)?;
        }

        // Tick marks
        if let Some(tm) = axis.major_tick_mark {
            let mut el = BytesStart::new("c:majorTickMark");
            el.push_attribute(("val", tm.xml_val()));
            writer.write_event(Event::Empty(el)).map_err(map_err)?;
        }
        if let Some(tm) = axis.minor_tick_mark {
            let mut el = BytesStart::new("c:minorTickMark");
            el.push_attribute(("val", tm.xml_val()));
            writer.write_event(Event::Empty(el)).map_err(map_err)?;
        }

        // Tick label position
        if let Some(tlp) = axis.tick_lbl_pos {
            let mut el = BytesStart::new("c:tickLblPos");
            el.push_attribute(("val", tlp.xml_val()));
            writer.write_event(Event::Empty(el)).map_err(map_err)?;
        }

        // <c:crossAx>
        let mut cross = BytesStart::new("c:crossAx");
        cross.push_attribute(("val", ibuf.format(axis.cross_ax)));
        writer
            .write_event(Event::Empty(cross))
            .map_err(map_err)?;

        // Crosses at
        if let Some(crosses_at) = axis.crosses_at {
            Self::write_f64_element(writer, "c:crossesAt", crosses_at)?;
        }

        // Major/minor unit
        if let Some(major) = axis.major_unit {
            Self::write_f64_element(writer, "c:majorUnit", major)?;
        }
        if let Some(minor) = axis.minor_unit {
            Self::write_f64_element(writer, "c:minorUnit", minor)?;
        }

        writer
            .write_event(Event::End(BytesEnd::new(tag)))
            .map_err(map_err)?;

        Ok(())
    }

    // -----------------------------------------------------------------------
    // Legend
    // -----------------------------------------------------------------------

    fn write_legend(
        writer: &mut Writer<&mut Vec<u8>>,
        legend: &ChartLegend,
    ) -> Result<()> {
        writer
            .write_event(Event::Start(BytesStart::new("c:legend")))
            .map_err(map_err)?;

        let mut pos = BytesStart::new("c:legendPos");
        pos.push_attribute(("val", legend.position.xml_val()));
        writer.write_event(Event::Empty(pos)).map_err(map_err)?;

        let mut overlay = BytesStart::new("c:overlay");
        overlay.push_attribute(("val", if legend.overlay { "1" } else { "0" }));
        writer
            .write_event(Event::Empty(overlay))
            .map_err(map_err)?;

        writer
            .write_event(Event::End(BytesEnd::new("c:legend")))
            .map_err(map_err)?;

        Ok(())
    }

    // -----------------------------------------------------------------------
    // Print settings (always present)
    // -----------------------------------------------------------------------

    fn write_print_settings(writer: &mut Writer<&mut Vec<u8>>) -> Result<()> {
        writer
            .write_event(Event::Start(BytesStart::new("c:printSettings")))
            .map_err(map_err)?;

        // <c:headerFooter/>
        writer
            .write_event(Event::Empty(BytesStart::new("c:headerFooter")))
            .map_err(map_err)?;

        // <c:pageMargins>
        let mut pm = BytesStart::new("c:pageMargins");
        pm.push_attribute(("b", "0.75"));
        pm.push_attribute(("l", "0.7"));
        pm.push_attribute(("r", "0.7"));
        pm.push_attribute(("t", "0.75"));
        pm.push_attribute(("header", "0.3"));
        pm.push_attribute(("footer", "0.3"));
        writer.write_event(Event::Empty(pm)).map_err(map_err)?;

        // <c:pageSetup/>
        writer
            .write_event(Event::Empty(BytesStart::new("c:pageSetup")))
            .map_err(map_err)?;

        writer
            .write_event(Event::End(BytesEnd::new("c:printSettings")))
            .map_err(map_err)?;

        Ok(())
    }

    // -----------------------------------------------------------------------
    // Utility helpers
    // -----------------------------------------------------------------------

    fn write_bool_element(
        writer: &mut Writer<&mut Vec<u8>>,
        tag: &str,
        value: bool,
    ) -> Result<()> {
        let mut el = BytesStart::new(tag);
        el.push_attribute(("val", if value { "1" } else { "0" }));
        writer.write_event(Event::Empty(el)).map_err(map_err)?;
        Ok(())
    }

    fn write_f64_element(
        writer: &mut Writer<&mut Vec<u8>>,
        tag: &str,
        value: f64,
    ) -> Result<()> {
        let mut el = BytesStart::new(tag);
        let formatted = value.to_string();
        el.push_attribute(("val", formatted.as_str()));
        writer.write_event(Event::Empty(el)).map_err(map_err)?;
        Ok(())
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_chart_data_serde_roundtrip() {
        let chart = ChartData {
            chart_type: ChartType::Bar,
            title: Some(ChartTitle { text: "Sales".into(), overlay: false, font_size: None, bold: None, color: None }),
            series: vec![ChartSeries {
                idx: 0, order: 0,
                name: Some("Revenue".into()),
                cat_ref: Some("Sheet1!$A$2:$A$5".into()),
                val_ref: "Sheet1!$B$2:$B$5".into(),
                x_val_ref: None, bubble_size_ref: None,
                fill_color: Some("4472C4".into()),
                line_color: None, line_width: None,
                marker: None, smooth: None, explosion: None,
                data_labels: None,
            }],
            cat_axis: Some(ChartAxis { id: 0, cross_ax: 1, title: None, num_fmt: None, source_linked: false, min: None, max: None, major_unit: None, minor_unit: None, log_base: None, reversed: false, tick_lbl_pos: None, major_tick_mark: None, minor_tick_mark: None, major_gridlines: false, minor_gridlines: false, delete: false, position: None, crosses_at: None, font_size: None }),
            val_axis: Some(ChartAxis { id: 1, cross_ax: 0, title: None, num_fmt: None, source_linked: false, min: None, max: None, major_unit: None, minor_unit: None, log_base: None, reversed: false, tick_lbl_pos: None, major_tick_mark: None, minor_tick_mark: None, major_gridlines: true, minor_gridlines: false, delete: false, position: None, crosses_at: None, font_size: None }),
            legend: Some(ChartLegend { position: LegendPosition::Bottom, overlay: false }),
            data_labels: None,
            grouping: Some(ChartGrouping::Clustered),
            scatter_style: None, radar_style: None, hole_size: None,
            bar_dir_horizontal: Some(false),
            style_id: Some(2),
            plot_area_layout: None,
        };
        let json = serde_json::to_string(&chart).unwrap();
        let roundtrip: ChartData = serde_json::from_str(&json).unwrap();
        assert_eq!(roundtrip.chart_type, ChartType::Bar);
        assert_eq!(roundtrip.series.len(), 1);
        assert_eq!(roundtrip.series[0].val_ref, "Sheet1!$B$2:$B$5");
    }

    #[test]
    fn test_chart_type_variants() {
        for (variant, expected) in [
            (ChartType::Bar, "\"bar\""),
            (ChartType::Line, "\"line\""),
            (ChartType::Pie, "\"pie\""),
            (ChartType::Scatter, "\"scatter\""),
            (ChartType::Area, "\"area\""),
            (ChartType::Column, "\"column\""),
            (ChartType::Doughnut, "\"doughnut\""),
            (ChartType::Radar, "\"radar\""),
            (ChartType::Bubble, "\"bubble\""),
            (ChartType::Stock, "\"stock\""),
        ] {
            assert_eq!(serde_json::to_string(&variant).unwrap(), expected);
        }
    }

    #[test]
    fn test_chart_anchor_serde() {
        let anchor = ChartAnchor {
            from_col: 0, from_row: 0, from_col_off: 0, from_row_off: 0,
            to_col: 8, to_row: 15, to_col_off: 0, to_row_off: 0,
        };
        let json = serde_json::to_string(&anchor).unwrap();
        let rt: ChartAnchor = serde_json::from_str(&json).unwrap();
        assert_eq!(rt.to_col, 8);
        assert_eq!(rt.to_row, 15);
    }

    #[test]
    fn test_worksheet_chart_serde() {
        let wc = WorksheetChart {
            chart: ChartData {
                chart_type: ChartType::Pie,
                title: None, series: vec![], cat_axis: None, val_axis: None,
                legend: None, data_labels: None, grouping: None,
                scatter_style: None, radar_style: None, hole_size: None,
                bar_dir_horizontal: None, style_id: None, plot_area_layout: None,
            },
            anchor: ChartAnchor {
                from_col: 5, from_row: 0, from_col_off: 0, from_row_off: 0,
                to_col: 12, to_row: 18, to_col_off: 0, to_row_off: 0,
            },
        };
        let json = serde_json::to_string(&wc).unwrap();
        assert!(json.contains("\"chartType\":\"pie\""));
    }

    #[test]
    fn test_skip_serializing_optional_fields() {
        let chart = ChartData {
            chart_type: ChartType::Line,
            title: None, series: vec![], cat_axis: None, val_axis: None,
            legend: None, data_labels: None, grouping: None,
            scatter_style: None, radar_style: None, hole_size: None,
            bar_dir_horizontal: None, style_id: None, plot_area_layout: None,
        };
        let json = serde_json::to_string(&chart).unwrap();
        assert!(!json.contains("title"));
        assert!(!json.contains("catAxis"));
        assert!(!json.contains("legend"));
        assert!(!json.contains("scatterStyle"));
    }

    #[test]
    fn test_grouping_serde() {
        assert_eq!(serde_json::to_string(&ChartGrouping::Clustered).unwrap(), "\"clustered\"");
        assert_eq!(serde_json::to_string(&ChartGrouping::Stacked).unwrap(), "\"stacked\"");
        assert_eq!(serde_json::to_string(&ChartGrouping::PercentStacked).unwrap(), "\"percentStacked\"");
        assert_eq!(serde_json::to_string(&ChartGrouping::Standard).unwrap(), "\"standard\"");
    }

    #[test]
    fn test_marker_style_serde() {
        for (variant, expected) in [
            (MarkerStyle::Circle, "\"circle\""),
            (MarkerStyle::Square, "\"square\""),
            (MarkerStyle::Diamond, "\"diamond\""),
            (MarkerStyle::None, "\"none\""),
        ] {
            assert_eq!(serde_json::to_string(&variant).unwrap(), expected);
        }
    }

    #[test]
    fn test_data_labels_defaults() {
        let json = r#"{"showVal":true}"#;
        let dl: DataLabels = serde_json::from_str(json).unwrap();
        assert!(dl.show_val);
        assert!(!dl.show_cat_name);
        assert!(!dl.show_percent);
    }

    #[test]
    fn test_axis_with_scale_options() {
        let axis = ChartAxis {
            id: 1, cross_ax: 0,
            title: Some(ChartTitle { text: "Revenue ($)".into(), overlay: false, font_size: Some(1200), bold: Some(true), color: None }),
            num_fmt: Some("#,##0".into()),
            source_linked: false,
            min: Some(0.0), max: Some(100.0),
            major_unit: Some(20.0), minor_unit: Some(5.0),
            log_base: None, reversed: false,
            tick_lbl_pos: Some(TickLabelPosition::NextTo),
            major_tick_mark: Some(TickMark::Out),
            minor_tick_mark: Some(TickMark::None),
            major_gridlines: true, minor_gridlines: false,
            delete: false,
            position: Some(AxisPosition::Left),
            crosses_at: Some(0.0),
            font_size: Some(1000),
        };
        let json = serde_json::to_string(&axis).unwrap();
        let rt: ChartAxis = serde_json::from_str(&json).unwrap();
        assert_eq!(rt.min, Some(0.0));
        assert_eq!(rt.max, Some(100.0));
        assert_eq!(rt.position, Some(AxisPosition::Left));
    }

    #[test]
    fn test_manual_layout_serde() {
        let layout = ManualLayout { x: 0.1, y: 0.15, w: 0.8, h: 0.7 };
        let json = serde_json::to_string(&layout).unwrap();
        let rt: ManualLayout = serde_json::from_str(&json).unwrap();
        assert!((rt.x - 0.1).abs() < f64::EPSILON);
        assert!((rt.h - 0.7).abs() < f64::EPSILON);
    }

    // =======================================================================
    // XML Writer tests
    // =======================================================================

    #[test]
    fn test_bar_chart_to_xml() {
        let chart = ChartData {
            chart_type: ChartType::Bar,
            title: Some(ChartTitle { text: "Sales".into(), overlay: false, font_size: None, bold: None, color: None }),
            series: vec![ChartSeries {
                idx: 0, order: 0,
                name: Some("Revenue".into()),
                cat_ref: Some("Sheet1!$A$2:$A$5".into()),
                val_ref: "Sheet1!$B$2:$B$5".into(),
                x_val_ref: None, bubble_size_ref: None,
                fill_color: Some("4472C4".into()),
                line_color: None, line_width: None,
                marker: None, smooth: None, explosion: None,
                data_labels: None,
            }],
            cat_axis: Some(ChartAxis { id: 0, cross_ax: 1, title: None, num_fmt: None, source_linked: false, min: None, max: None, major_unit: None, minor_unit: None, log_base: None, reversed: false, tick_lbl_pos: None, major_tick_mark: None, minor_tick_mark: None, major_gridlines: false, minor_gridlines: false, delete: false, position: None, crosses_at: None, font_size: None }),
            val_axis: Some(ChartAxis { id: 1, cross_ax: 0, title: None, num_fmt: None, source_linked: false, min: None, max: None, major_unit: None, minor_unit: None, log_base: None, reversed: false, tick_lbl_pos: None, major_tick_mark: None, minor_tick_mark: None, major_gridlines: true, minor_gridlines: false, delete: false, position: None, crosses_at: None, font_size: None }),
            legend: Some(ChartLegend { position: LegendPosition::Bottom, overlay: false }),
            data_labels: None,
            grouping: Some(ChartGrouping::Clustered),
            scatter_style: None, radar_style: None, hole_size: None,
            bar_dir_horizontal: Some(true),
            style_id: None,
            plot_area_layout: None,
        };
        let xml = chart.to_xml().unwrap();
        let xml_str = std::str::from_utf8(&xml).unwrap();
        assert!(xml_str.contains("<c:barChart>"));
        assert!(xml_str.contains("<c:barDir val=\"bar\"/>"));
        assert!(xml_str.contains("<c:grouping val=\"clustered\"/>"));
        assert!(xml_str.contains("<c:f>Sheet1!$B$2:$B$5</c:f>"));
        assert!(xml_str.contains("<c:catAx>"));
        assert!(xml_str.contains("<c:valAx>"));
        assert!(xml_str.contains("<c:legend>"));
        assert!(xml_str.contains("<c:legendPos val=\"b\"/>"));
    }

    #[test]
    fn test_column_chart_to_xml() {
        let chart = ChartData {
            chart_type: ChartType::Column,
            title: None,
            series: vec![ChartSeries {
                idx: 0, order: 0,
                name: None,
                cat_ref: None,
                val_ref: "Sheet1!$B$2:$B$5".into(),
                x_val_ref: None, bubble_size_ref: None,
                fill_color: None, line_color: None, line_width: None,
                marker: None, smooth: None, explosion: None,
                data_labels: None,
            }],
            cat_axis: Some(ChartAxis { id: 0, cross_ax: 1, title: None, num_fmt: None, source_linked: false, min: None, max: None, major_unit: None, minor_unit: None, log_base: None, reversed: false, tick_lbl_pos: None, major_tick_mark: None, minor_tick_mark: None, major_gridlines: false, minor_gridlines: false, delete: false, position: None, crosses_at: None, font_size: None }),
            val_axis: Some(ChartAxis { id: 1, cross_ax: 0, title: None, num_fmt: None, source_linked: false, min: None, max: None, major_unit: None, minor_unit: None, log_base: None, reversed: false, tick_lbl_pos: None, major_tick_mark: None, minor_tick_mark: None, major_gridlines: true, minor_gridlines: false, delete: false, position: None, crosses_at: None, font_size: None }),
            legend: None,
            data_labels: None,
            grouping: Some(ChartGrouping::Clustered),
            scatter_style: None, radar_style: None, hole_size: None,
            bar_dir_horizontal: None,
            style_id: None,
            plot_area_layout: None,
        };
        let xml = chart.to_xml().unwrap();
        let xml_str = std::str::from_utf8(&xml).unwrap();
        assert!(xml_str.contains("<c:barChart>"));
        assert!(xml_str.contains("<c:barDir val=\"col\"/>"));
    }

    #[test]
    fn test_line_chart_with_markers_to_xml() {
        let chart = ChartData {
            chart_type: ChartType::Line,
            title: None,
            series: vec![ChartSeries {
                idx: 0, order: 0, name: None, cat_ref: None,
                val_ref: "Sheet1!$B$2:$B$5".into(),
                x_val_ref: None, bubble_size_ref: None,
                fill_color: None, line_color: Some("FF0000".into()), line_width: Some(25400),
                marker: Some(MarkerStyle::Circle), smooth: Some(true), explosion: None,
                data_labels: None,
            }],
            cat_axis: Some(ChartAxis { id: 0, cross_ax: 1, title: None, num_fmt: None, source_linked: false, min: None, max: None, major_unit: None, minor_unit: None, log_base: None, reversed: false, tick_lbl_pos: None, major_tick_mark: None, minor_tick_mark: None, major_gridlines: false, minor_gridlines: false, delete: false, position: None, crosses_at: None, font_size: None }),
            val_axis: Some(ChartAxis { id: 1, cross_ax: 0, title: None, num_fmt: None, source_linked: false, min: None, max: None, major_unit: None, minor_unit: None, log_base: None, reversed: false, tick_lbl_pos: None, major_tick_mark: None, minor_tick_mark: None, major_gridlines: true, minor_gridlines: false, delete: false, position: None, crosses_at: None, font_size: None }),
            legend: None, data_labels: None,
            grouping: Some(ChartGrouping::Standard),
            scatter_style: None, radar_style: None, hole_size: None,
            bar_dir_horizontal: None, style_id: None, plot_area_layout: None,
        };
        let xml = chart.to_xml().unwrap();
        let xml_str = std::str::from_utf8(&xml).unwrap();
        assert!(xml_str.contains("<c:lineChart>"));
        assert!(xml_str.contains("<c:marker>"));
        assert!(xml_str.contains("<c:symbol val=\"circle\"/>"));
        assert!(xml_str.contains("<c:smooth val=\"1\"/>"));
    }

    #[test]
    fn test_pie_chart_to_xml() {
        let chart = ChartData {
            chart_type: ChartType::Pie,
            title: Some(ChartTitle { text: "Distribution".into(), overlay: false, font_size: Some(1400), bold: Some(true), color: None }),
            series: vec![ChartSeries {
                idx: 0, order: 0,
                name: Some("Shares".into()),
                cat_ref: Some("Sheet1!$A$2:$A$5".into()),
                val_ref: "Sheet1!$B$2:$B$5".into(),
                x_val_ref: None, bubble_size_ref: None,
                fill_color: None, line_color: None, line_width: None,
                marker: None, smooth: None, explosion: Some(25),
                data_labels: None,
            }],
            cat_axis: None, val_axis: None,
            legend: Some(ChartLegend { position: LegendPosition::Right, overlay: false }),
            data_labels: Some(DataLabels { show_val: true, show_cat_name: false, show_ser_name: false, show_percent: true, num_fmt: None, show_leader_lines: true }),
            grouping: None, scatter_style: None, radar_style: None, hole_size: None,
            bar_dir_horizontal: None, style_id: None, plot_area_layout: None,
        };
        let xml = chart.to_xml().unwrap();
        let xml_str = std::str::from_utf8(&xml).unwrap();
        assert!(xml_str.contains("<c:pieChart>"));
        assert!(!xml_str.contains("<c:catAx>")); // no axes for pie
        assert!(!xml_str.contains("<c:valAx>"));
        assert!(xml_str.contains("<c:explosion val=\"25\"/>"));
        assert!(xml_str.contains("<c:showVal val=\"1\"/>"));
        assert!(xml_str.contains("<c:showPercent val=\"1\"/>"));
        assert!(xml_str.contains("<c:title>"));
        assert!(xml_str.contains("Distribution"));
    }

    #[test]
    fn test_scatter_chart_to_xml() {
        let chart = ChartData {
            chart_type: ChartType::Scatter,
            title: None,
            series: vec![ChartSeries {
                idx: 0, order: 0, name: None,
                cat_ref: None,
                val_ref: "Sheet1!$B$2:$B$5".into(),
                x_val_ref: Some("Sheet1!$A$2:$A$5".into()),
                bubble_size_ref: None,
                fill_color: None, line_color: None, line_width: None,
                marker: Some(MarkerStyle::Diamond), smooth: None, explosion: None,
                data_labels: None,
            }],
            cat_axis: None,
            val_axis: Some(ChartAxis { id: 0, cross_ax: 1, title: None, num_fmt: None, source_linked: false, min: None, max: None, major_unit: None, minor_unit: None, log_base: None, reversed: false, tick_lbl_pos: None, major_tick_mark: None, minor_tick_mark: None, major_gridlines: true, minor_gridlines: false, delete: false, position: None, crosses_at: None, font_size: None }),
            legend: None, data_labels: None,
            grouping: None,
            scatter_style: Some(ScatterStyle::LineMarker),
            radar_style: None, hole_size: None,
            bar_dir_horizontal: None, style_id: None, plot_area_layout: None,
        };
        let xml = chart.to_xml().unwrap();
        let xml_str = std::str::from_utf8(&xml).unwrap();
        assert!(xml_str.contains("<c:scatterChart>"));
        assert!(xml_str.contains("<c:scatterStyle val=\"lineMarker\"/>"));
        assert!(xml_str.contains("<c:xVal>"));
        assert!(xml_str.contains("<c:yVal>"));
        assert!(!xml_str.contains("<c:cat>"));
    }

    #[test]
    fn test_doughnut_chart_to_xml() {
        let chart = ChartData {
            chart_type: ChartType::Doughnut,
            title: None,
            series: vec![ChartSeries {
                idx: 0, order: 0, name: None,
                cat_ref: Some("Sheet1!$A$2:$A$5".into()),
                val_ref: "Sheet1!$B$2:$B$5".into(),
                x_val_ref: None, bubble_size_ref: None,
                fill_color: None, line_color: None, line_width: None,
                marker: None, smooth: None, explosion: None,
                data_labels: None,
            }],
            cat_axis: None, val_axis: None,
            legend: None, data_labels: None,
            grouping: None, scatter_style: None, radar_style: None,
            hole_size: Some(50),
            bar_dir_horizontal: None, style_id: None, plot_area_layout: None,
        };
        let xml = chart.to_xml().unwrap();
        let xml_str = std::str::from_utf8(&xml).unwrap();
        assert!(xml_str.contains("<c:doughnutChart>"));
        assert!(xml_str.contains("<c:holeSize val=\"50\"/>"));
    }

    #[test]
    fn test_area_chart_to_xml() {
        let chart = ChartData {
            chart_type: ChartType::Area,
            title: None,
            series: vec![ChartSeries {
                idx: 0, order: 0, name: None, cat_ref: None,
                val_ref: "Sheet1!$B$2:$B$5".into(),
                x_val_ref: None, bubble_size_ref: None,
                fill_color: Some("4472C4".into()), line_color: None, line_width: None,
                marker: None, smooth: None, explosion: None,
                data_labels: None,
            }],
            cat_axis: Some(ChartAxis { id: 0, cross_ax: 1, title: None, num_fmt: None, source_linked: false, min: None, max: None, major_unit: None, minor_unit: None, log_base: None, reversed: false, tick_lbl_pos: None, major_tick_mark: None, minor_tick_mark: None, major_gridlines: false, minor_gridlines: false, delete: false, position: None, crosses_at: None, font_size: None }),
            val_axis: Some(ChartAxis { id: 1, cross_ax: 0, title: None, num_fmt: None, source_linked: false, min: None, max: None, major_unit: None, minor_unit: None, log_base: None, reversed: false, tick_lbl_pos: None, major_tick_mark: None, minor_tick_mark: None, major_gridlines: true, minor_gridlines: false, delete: false, position: None, crosses_at: None, font_size: None }),
            legend: None, data_labels: None,
            grouping: Some(ChartGrouping::Stacked),
            scatter_style: None, radar_style: None, hole_size: None,
            bar_dir_horizontal: None, style_id: None, plot_area_layout: None,
        };
        let xml = chart.to_xml().unwrap();
        let xml_str = std::str::from_utf8(&xml).unwrap();
        assert!(xml_str.contains("<c:areaChart>"));
        assert!(xml_str.contains("<c:grouping val=\"stacked\"/>"));
    }

    #[test]
    fn test_radar_chart_to_xml() {
        let chart = ChartData {
            chart_type: ChartType::Radar,
            title: None,
            series: vec![ChartSeries {
                idx: 0, order: 0, name: None, cat_ref: None,
                val_ref: "Sheet1!$B$2:$B$5".into(),
                x_val_ref: None, bubble_size_ref: None,
                fill_color: None, line_color: None, line_width: None,
                marker: None, smooth: None, explosion: None,
                data_labels: None,
            }],
            cat_axis: Some(ChartAxis { id: 0, cross_ax: 1, title: None, num_fmt: None, source_linked: false, min: None, max: None, major_unit: None, minor_unit: None, log_base: None, reversed: false, tick_lbl_pos: None, major_tick_mark: None, minor_tick_mark: None, major_gridlines: false, minor_gridlines: false, delete: false, position: None, crosses_at: None, font_size: None }),
            val_axis: Some(ChartAxis { id: 1, cross_ax: 0, title: None, num_fmt: None, source_linked: false, min: None, max: None, major_unit: None, minor_unit: None, log_base: None, reversed: false, tick_lbl_pos: None, major_tick_mark: None, minor_tick_mark: None, major_gridlines: true, minor_gridlines: false, delete: false, position: None, crosses_at: None, font_size: None }),
            legend: None, data_labels: None,
            grouping: None, scatter_style: None,
            radar_style: Some(RadarStyle::Filled),
            hole_size: None,
            bar_dir_horizontal: None, style_id: None, plot_area_layout: None,
        };
        let xml = chart.to_xml().unwrap();
        let xml_str = std::str::from_utf8(&xml).unwrap();
        assert!(xml_str.contains("<c:radarChart>"));
        assert!(xml_str.contains("<c:radarStyle val=\"filled\"/>"));
    }

    #[test]
    fn test_chart_with_style_id() {
        let chart = ChartData {
            chart_type: ChartType::Line,
            title: None, series: vec![], cat_axis: None, val_axis: None,
            legend: None, data_labels: None,
            grouping: None, scatter_style: None, radar_style: None, hole_size: None,
            bar_dir_horizontal: None,
            style_id: Some(26),
            plot_area_layout: None,
        };
        let xml = chart.to_xml().unwrap();
        let xml_str = std::str::from_utf8(&xml).unwrap();
        assert!(xml_str.contains("<c:style val=\"26\"/>"));
    }

    #[test]
    fn test_chart_with_axis_title_and_scale() {
        let chart = ChartData {
            chart_type: ChartType::Line,
            title: None,
            series: vec![],
            cat_axis: Some(ChartAxis {
                id: 0, cross_ax: 1,
                title: Some(ChartTitle { text: "Month".into(), overlay: false, font_size: None, bold: None, color: None }),
                num_fmt: None, source_linked: false,
                min: None, max: None, major_unit: None, minor_unit: None,
                log_base: None, reversed: false,
                tick_lbl_pos: Some(TickLabelPosition::Low),
                major_tick_mark: Some(TickMark::Cross),
                minor_tick_mark: None,
                major_gridlines: true, minor_gridlines: false,
                delete: false, position: Some(AxisPosition::Bottom), crosses_at: None, font_size: None,
            }),
            val_axis: Some(ChartAxis {
                id: 1, cross_ax: 0,
                title: Some(ChartTitle { text: "Sales ($)".into(), overlay: false, font_size: Some(1200), bold: Some(true), color: Some("333333".into()) }),
                num_fmt: Some("#,##0".into()), source_linked: false,
                min: Some(0.0), max: Some(1000.0),
                major_unit: Some(200.0), minor_unit: None,
                log_base: None, reversed: false,
                tick_lbl_pos: None, major_tick_mark: None, minor_tick_mark: None,
                major_gridlines: true, minor_gridlines: false,
                delete: false, position: Some(AxisPosition::Left), crosses_at: None, font_size: None,
            }),
            legend: None, data_labels: None, grouping: None,
            scatter_style: None, radar_style: None, hole_size: None,
            bar_dir_horizontal: None, style_id: None, plot_area_layout: None,
        };
        let xml = chart.to_xml().unwrap();
        let xml_str = std::str::from_utf8(&xml).unwrap();
        assert!(xml_str.contains("Month"));
        assert!(xml_str.contains("Sales ($)"));
        assert!(xml_str.contains("<c:majorGridlines/>"));
    }

    #[test]
    fn test_chart_xml_has_proper_namespaces() {
        let chart = ChartData {
            chart_type: ChartType::Pie,
            title: None, series: vec![], cat_axis: None, val_axis: None,
            legend: None, data_labels: None, grouping: None,
            scatter_style: None, radar_style: None, hole_size: None,
            bar_dir_horizontal: None, style_id: None, plot_area_layout: None,
        };
        let xml = chart.to_xml().unwrap();
        let xml_str = std::str::from_utf8(&xml).unwrap();
        assert!(xml_str.contains("xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\""));
        assert!(xml_str.contains("xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\""));
        assert!(xml_str.contains("xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\""));
        assert!(xml_str.contains("<c:printSettings>"));
    }
}
