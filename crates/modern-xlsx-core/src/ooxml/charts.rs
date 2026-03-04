//! Chart definitions — `xl/charts/chart{n}.xml`.

use quick_xml::events::{BytesDecl, BytesEnd, BytesStart, BytesText, Event};
use quick_xml::{Reader, Writer};
use serde::{Deserialize, Serialize};

use super::push_entity;
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
// Drawing XML generation
// ---------------------------------------------------------------------------

impl ChartAnchor {
    /// Generate the complete `xl/drawings/drawing{n}.xml` for a worksheet's charts.
    ///
    /// Each chart gets a `<xdr:twoCellAnchor>` referencing its chart via the
    /// corresponding relationship ID from `chart_r_ids`.
    pub fn generate_drawing_xml(
        charts: &[WorksheetChart],
        chart_r_ids: &[String],
    ) -> Result<Vec<u8>> {
        let mut buf: Vec<u8> = Vec::with_capacity(512 + charts.len() * 512);
        let mut writer = Writer::new(&mut buf);
        let mut ibuf = itoa::Buffer::new();

        let map_err = |e: std::io::Error| ModernXlsxError::XmlWrite(e.to_string());

        // XML declaration.
        writer
            .write_event(Event::Decl(BytesDecl::new("1.0", Some("UTF-8"), Some("yes"))))
            .map_err(map_err)?;

        // <xdr:wsDr>
        let mut root = BytesStart::new("xdr:wsDr");
        root.push_attribute(("xmlns:xdr", "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"));
        root.push_attribute(("xmlns:a", "http://schemas.openxmlformats.org/drawingml/2006/main"));
        root.push_attribute(("xmlns:r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships"));
        writer.write_event(Event::Start(root)).map_err(map_err)?;

        for (i, wsc) in charts.iter().enumerate() {
            let anchor = &wsc.anchor;
            let r_id = &chart_r_ids[i];
            let cnv_id = i as u32 + 2; // id starts at 2

            // <xdr:twoCellAnchor>
            writer
                .write_event(Event::Start(BytesStart::new("xdr:twoCellAnchor")))
                .map_err(map_err)?;

            // <xdr:from>
            writer
                .write_event(Event::Start(BytesStart::new("xdr:from")))
                .map_err(map_err)?;
            Self::write_cell_pos(&mut writer, anchor.from_col, anchor.from_col_off, anchor.from_row, anchor.from_row_off, &mut ibuf)?;
            writer
                .write_event(Event::End(BytesEnd::new("xdr:from")))
                .map_err(map_err)?;

            // <xdr:to>
            writer
                .write_event(Event::Start(BytesStart::new("xdr:to")))
                .map_err(map_err)?;
            Self::write_cell_pos(&mut writer, anchor.to_col, anchor.to_col_off, anchor.to_row, anchor.to_row_off, &mut ibuf)?;
            writer
                .write_event(Event::End(BytesEnd::new("xdr:to")))
                .map_err(map_err)?;

            // <xdr:graphicFrame macro="">
            let mut gf = BytesStart::new("xdr:graphicFrame");
            gf.push_attribute(("macro", ""));
            writer.write_event(Event::Start(gf)).map_err(map_err)?;

            // <xdr:nvGraphicFramePr>
            writer
                .write_event(Event::Start(BytesStart::new("xdr:nvGraphicFramePr")))
                .map_err(map_err)?;
            let mut cnv_pr = BytesStart::new("xdr:cNvPr");
            cnv_pr.push_attribute(("id", ibuf.format(cnv_id)));
            let chart_name = format!("Chart {}", i + 1);
            cnv_pr.push_attribute(("name", chart_name.as_str()));
            writer.write_event(Event::Empty(cnv_pr)).map_err(map_err)?;
            writer
                .write_event(Event::Empty(BytesStart::new("xdr:cNvGraphicFramePr")))
                .map_err(map_err)?;
            writer
                .write_event(Event::End(BytesEnd::new("xdr:nvGraphicFramePr")))
                .map_err(map_err)?;

            // <xdr:xfrm>
            writer
                .write_event(Event::Start(BytesStart::new("xdr:xfrm")))
                .map_err(map_err)?;
            let mut off = BytesStart::new("a:off");
            off.push_attribute(("x", "0"));
            off.push_attribute(("y", "0"));
            writer.write_event(Event::Empty(off)).map_err(map_err)?;
            let mut ext = BytesStart::new("a:ext");
            ext.push_attribute(("cx", "0"));
            ext.push_attribute(("cy", "0"));
            writer.write_event(Event::Empty(ext)).map_err(map_err)?;
            writer
                .write_event(Event::End(BytesEnd::new("xdr:xfrm")))
                .map_err(map_err)?;

            // <a:graphic>
            writer
                .write_event(Event::Start(BytesStart::new("a:graphic")))
                .map_err(map_err)?;
            let mut gd = BytesStart::new("a:graphicData");
            gd.push_attribute(("uri", "http://schemas.openxmlformats.org/drawingml/2006/chart"));
            writer.write_event(Event::Start(gd)).map_err(map_err)?;
            let mut chart_ref = BytesStart::new("c:chart");
            chart_ref.push_attribute(("xmlns:c", "http://schemas.openxmlformats.org/drawingml/2006/chart"));
            chart_ref.push_attribute(("r:id", r_id.as_str()));
            writer.write_event(Event::Empty(chart_ref)).map_err(map_err)?;
            writer
                .write_event(Event::End(BytesEnd::new("a:graphicData")))
                .map_err(map_err)?;
            writer
                .write_event(Event::End(BytesEnd::new("a:graphic")))
                .map_err(map_err)?;

            // </xdr:graphicFrame>
            writer
                .write_event(Event::End(BytesEnd::new("xdr:graphicFrame")))
                .map_err(map_err)?;

            // <xdr:clientData/>
            writer
                .write_event(Event::Empty(BytesStart::new("xdr:clientData")))
                .map_err(map_err)?;

            // </xdr:twoCellAnchor>
            writer
                .write_event(Event::End(BytesEnd::new("xdr:twoCellAnchor")))
                .map_err(map_err)?;
        }

        // </xdr:wsDr>
        writer
            .write_event(Event::End(BytesEnd::new("xdr:wsDr")))
            .map_err(map_err)?;

        Ok(buf)
    }

    /// Write the `<xdr:col>`, `<xdr:colOff>`, `<xdr:row>`, `<xdr:rowOff>` children.
    fn write_cell_pos(
        writer: &mut Writer<&mut Vec<u8>>,
        col: u32,
        col_off: u64,
        row: u32,
        row_off: u64,
        ibuf: &mut itoa::Buffer,
    ) -> Result<()> {
        let map_err = |e: std::io::Error| ModernXlsxError::XmlWrite(e.to_string());

        writer
            .write_event(Event::Start(BytesStart::new("xdr:col")))
            .map_err(map_err)?;
        writer
            .write_event(Event::Text(BytesText::new(ibuf.format(col))))
            .map_err(map_err)?;
        writer
            .write_event(Event::End(BytesEnd::new("xdr:col")))
            .map_err(map_err)?;

        writer
            .write_event(Event::Start(BytesStart::new("xdr:colOff")))
            .map_err(map_err)?;
        writer
            .write_event(Event::Text(BytesText::new(ibuf.format(col_off))))
            .map_err(map_err)?;
        writer
            .write_event(Event::End(BytesEnd::new("xdr:colOff")))
            .map_err(map_err)?;

        writer
            .write_event(Event::Start(BytesStart::new("xdr:row")))
            .map_err(map_err)?;
        writer
            .write_event(Event::Text(BytesText::new(ibuf.format(row))))
            .map_err(map_err)?;
        writer
            .write_event(Event::End(BytesEnd::new("xdr:row")))
            .map_err(map_err)?;

        writer
            .write_event(Event::Start(BytesStart::new("xdr:rowOff")))
            .map_err(map_err)?;
        writer
            .write_event(Event::Text(BytesText::new(ibuf.format(row_off))))
            .map_err(map_err)?;
        writer
            .write_event(Event::End(BytesEnd::new("xdr:rowOff")))
            .map_err(map_err)?;

        Ok(())
    }
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

// ---------------------------------------------------------------------------
// XML Parser — chart XML → ChartData
// ---------------------------------------------------------------------------

/// Helper to extract `val` attribute from a `BytesStart`.
fn attr_val(e: &BytesStart<'_>, key: &[u8]) -> Option<String> {
    for attr in e.attributes().flatten() {
        if attr.key.as_ref() == key {
            return Some(
                std::str::from_utf8(&attr.value)
                    .unwrap_or_default()
                    .to_owned(),
            );
        }
    }
    None
}

/// Parse the `val` attribute as `&str`.
fn attr_val_str(e: &BytesStart<'_>) -> String {
    attr_val(e, b"val").unwrap_or_default()
}

impl ChartGrouping {
    fn from_xml(s: &str) -> Option<Self> {
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
    fn from_xml(s: &str) -> Option<Self> {
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
    fn from_xml(s: &str) -> Option<Self> {
        match s {
            "standard" => Some(Self::Standard),
            "marker" => Some(Self::Marker),
            "filled" => Some(Self::Filled),
            _ => None,
        }
    }
}

impl MarkerStyle {
    fn from_xml(s: &str) -> Option<Self> {
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
            _ => Option::None,
        }
    }
}

impl LegendPosition {
    fn from_xml(s: &str) -> Option<Self> {
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
    fn from_xml(s: &str) -> Option<Self> {
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
    fn from_xml(s: &str) -> Option<Self> {
        match s {
            "high" => Some(Self::High),
            "low" => Some(Self::Low),
            "nextTo" => Some(Self::NextTo),
            "none" => Some(Self::None),
            _ => Option::None,
        }
    }
}

impl TickMark {
    fn from_xml(s: &str) -> Option<Self> {
        match s {
            "cross" => Some(Self::Cross),
            "in" => Some(Self::In),
            "out" => Some(Self::Out),
            "none" => Some(Self::None),
            _ => Option::None,
        }
    }
}

/// Parser state machine context.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[allow(dead_code)]
enum ParseCtx {
    /// Top-level or unknown context.
    Root,
    /// Inside `<c:chart>`.
    Chart,
    /// Inside a chart-type element (`<c:barChart>`, etc.).
    ChartTypeElem,
    /// Inside `<c:ser>`.
    Series,
    /// Inside `<c:tx>` within a series.
    SerTx,
    /// Inside `<c:cat>` or `<c:val>` or `<c:xVal>` or `<c:yVal>` or `<c:bubbleSize>`.
    SerRef(SerRefKind),
    /// Inside `<c:spPr>` within a series.
    SerSpPr,
    /// Inside `<a:ln>` within `<c:spPr>`.
    SerSpPrLn,
    /// Inside `<a:solidFill>` within `<c:spPr>` (fill).
    SerSpPrFill,
    /// Inside `<a:solidFill>` within `<a:ln>` (line color).
    SerSpPrLnFill,
    /// Inside `<c:marker>` within a series.
    SerMarker,
    /// Inside `<c:dLbls>` (chart-level or series-level).
    DataLabels,
    /// Inside `<c:catAx>`.
    CatAxis,
    /// Inside `<c:valAx>`.
    ValAxis,
    /// Inside `<c:scaling>` within an axis.
    AxisScaling(AxisKind),
    /// Inside `<c:legend>`.
    Legend,
    /// Inside `<c:title>` — we track title context.
    Title(TitleOwner),
    /// Inside `<c:plotArea>`.
    PlotArea,
    /// Inside `<c:layout>` → `<c:manualLayout>`.
    ManualLayoutCtx,
    /// Inside `<c:dLbls>` within a series.
    SerDataLabels,
    /// Inside the title text run (`<a:r>` → `<a:t>`).
    TitleRun(TitleOwner),
    /// Inside `<a:rPr>` within a title run (to find solidFill for color).
    TitleRunPr(TitleOwner),
    /// Inside `<a:solidFill>` within `<a:rPr>` (title color).
    TitleRunPrFill(TitleOwner),
    /// Inside `<a:pPr>` within a title paragraph.
    TitlePPr(TitleOwner),
    /// Inside `<a:defRPr>` within `<a:pPr>`.
    TitleDefRPr(TitleOwner),
    /// Inside `<a:solidFill>` within `<a:defRPr>` (title color via pPr).
    TitleDefRPrFill(TitleOwner),
    /// Inside axis title.
    AxisTitle(AxisKind),
    /// Inside axis title text run.
    AxisTitleRun(AxisKind),
    /// Inside `<c:numFmt>` within an axis.
    AxisNumFmt(AxisKind),
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum TitleOwner {
    Chart,
    CatAxis,
    ValAxis,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum AxisKind {
    Cat,
    Val,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum SerRefKind {
    Cat,
    Val,
    XVal,
    YVal,
    BubbleSize,
}

impl ChartData {
    /// Parse a chart XML (`xl/charts/chart{n}.xml`) into a `ChartData`.
    pub fn parse(data: &[u8]) -> Result<Self> {
        let mut reader = Reader::from_reader(data);
        reader.config_mut().trim_text(true);
        let mut buf = Vec::with_capacity(512);

        // Result fields.
        let mut chart_type: Option<ChartType> = None;
        let mut title: Option<ChartTitle> = None;
        let mut series: Vec<ChartSeries> = Vec::new();
        let mut cat_axis: Option<ChartAxis> = None;
        let mut val_axis: Option<ChartAxis> = None;
        let mut legend: Option<ChartLegend> = None;
        let mut data_labels: Option<DataLabels> = None;
        let mut grouping: Option<ChartGrouping> = None;
        let mut scatter_style: Option<ScatterStyle> = None;
        let mut radar_style: Option<RadarStyle> = None;
        let mut hole_size: Option<u32> = None;
        let mut bar_dir_horizontal: Option<bool> = None;
        let mut style_id: Option<u32> = None;
        let mut plot_area_layout: Option<ManualLayout> = None;

        // Current series being built.
        let mut cur_ser: Option<ChartSeries> = None;
        // Current axis being built.
        let mut cur_cat_axis = AxisBuilder::new();
        let mut cur_val_axis = AxisBuilder::new();
        // Current data labels.
        let mut cur_dlbls = DataLabelsBuilder::new();
        let mut cur_ser_dlbls = DataLabelsBuilder::new();
        // Current legend.
        let mut cur_legend = LegendBuilder::new();
        // Current title.
        let mut cur_title = TitleBuilder::new();
        let mut cur_cat_axis_title = TitleBuilder::new();
        let mut cur_val_axis_title = TitleBuilder::new();
        // Manual layout.
        let mut layout_x: Option<f64> = None;
        let mut layout_y: Option<f64> = None;
        let mut layout_w: Option<f64> = None;
        let mut layout_h: Option<f64> = None;

        // Text capture buffer.
        let mut text_buf = String::new();
        let mut capturing_text = false;
        let mut text_target = TextTarget::None;

        // Context stack (simple — we push on Start, pop on End).
        let mut ctx_stack: Vec<ParseCtx> = vec![ParseCtx::Root];

        fn current_ctx(stack: &[ParseCtx]) -> ParseCtx {
            stack.last().copied().unwrap_or(ParseCtx::Root)
        }

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) => {
                    let local = e.local_name();
                    let local = local.as_ref();
                    let ctx = current_ctx(&ctx_stack);

                    match (ctx, local) {
                        (ParseCtx::Root, b"chart") => {
                            ctx_stack.push(ParseCtx::Chart);
                        }
                        (ParseCtx::Chart, b"plotArea") => {
                            ctx_stack.push(ParseCtx::PlotArea);
                        }
                        (ParseCtx::Chart, b"title") => {
                            cur_title = TitleBuilder::new();
                            ctx_stack.push(ParseCtx::Title(TitleOwner::Chart));
                        }
                        (ParseCtx::Chart, b"legend") => {
                            cur_legend = LegendBuilder::new();
                            ctx_stack.push(ParseCtx::Legend);
                        }
                        // Chart type elements inside plotArea.
                        (ParseCtx::PlotArea, b"barChart") => {
                            // Default to Column; barDir will refine.
                            chart_type = Some(ChartType::Column);
                            ctx_stack.push(ParseCtx::ChartTypeElem);
                        }
                        (ParseCtx::PlotArea, b"lineChart") => {
                            chart_type = Some(ChartType::Line);
                            ctx_stack.push(ParseCtx::ChartTypeElem);
                        }
                        (ParseCtx::PlotArea, b"pieChart") => {
                            chart_type = Some(ChartType::Pie);
                            ctx_stack.push(ParseCtx::ChartTypeElem);
                        }
                        (ParseCtx::PlotArea, b"doughnutChart") => {
                            chart_type = Some(ChartType::Doughnut);
                            ctx_stack.push(ParseCtx::ChartTypeElem);
                        }
                        (ParseCtx::PlotArea, b"scatterChart") => {
                            chart_type = Some(ChartType::Scatter);
                            ctx_stack.push(ParseCtx::ChartTypeElem);
                        }
                        (ParseCtx::PlotArea, b"areaChart") => {
                            chart_type = Some(ChartType::Area);
                            ctx_stack.push(ParseCtx::ChartTypeElem);
                        }
                        (ParseCtx::PlotArea, b"radarChart") => {
                            chart_type = Some(ChartType::Radar);
                            ctx_stack.push(ParseCtx::ChartTypeElem);
                        }
                        (ParseCtx::PlotArea, b"bubbleChart") => {
                            chart_type = Some(ChartType::Bubble);
                            ctx_stack.push(ParseCtx::ChartTypeElem);
                        }
                        (ParseCtx::PlotArea, b"stockChart") => {
                            chart_type = Some(ChartType::Stock);
                            ctx_stack.push(ParseCtx::ChartTypeElem);
                        }
                        (ParseCtx::PlotArea, b"layout") => {
                            ctx_stack.push(ParseCtx::ManualLayoutCtx);
                        }
                        (ParseCtx::ManualLayoutCtx, b"manualLayout") => {
                            // stay in ManualLayoutCtx
                            ctx_stack.push(ParseCtx::ManualLayoutCtx);
                        }
                        // Axes inside plotArea.
                        (ParseCtx::PlotArea, b"catAx") => {
                            cur_cat_axis = AxisBuilder::new();
                            ctx_stack.push(ParseCtx::CatAxis);
                        }
                        (ParseCtx::PlotArea, b"valAx") => {
                            cur_val_axis = AxisBuilder::new();
                            ctx_stack.push(ParseCtx::ValAxis);
                        }
                        // Series inside chart type element.
                        (ParseCtx::ChartTypeElem, b"ser") => {
                            cur_ser = Some(ChartSeries {
                                idx: 0,
                                order: 0,
                                name: None,
                                cat_ref: None,
                                val_ref: String::new(),
                                x_val_ref: None,
                                bubble_size_ref: None,
                                fill_color: None,
                                line_color: None,
                                line_width: None,
                                marker: None,
                                smooth: None,
                                explosion: None,
                                data_labels: None,
                            });
                            ctx_stack.push(ParseCtx::Series);
                        }
                        (ParseCtx::ChartTypeElem, b"dLbls") => {
                            cur_dlbls = DataLabelsBuilder::new();
                            ctx_stack.push(ParseCtx::DataLabels);
                        }
                        // Series children.
                        (ParseCtx::Series, b"tx") => {
                            ctx_stack.push(ParseCtx::SerTx);
                        }
                        (ParseCtx::Series, b"cat") => {
                            ctx_stack.push(ParseCtx::SerRef(SerRefKind::Cat));
                        }
                        (ParseCtx::Series, b"val") => {
                            ctx_stack.push(ParseCtx::SerRef(SerRefKind::Val));
                        }
                        (ParseCtx::Series, b"xVal") => {
                            ctx_stack.push(ParseCtx::SerRef(SerRefKind::XVal));
                        }
                        (ParseCtx::Series, b"yVal") => {
                            ctx_stack.push(ParseCtx::SerRef(SerRefKind::YVal));
                        }
                        (ParseCtx::Series, b"bubbleSize") => {
                            ctx_stack.push(ParseCtx::SerRef(SerRefKind::BubbleSize));
                        }
                        (ParseCtx::Series, b"spPr") => {
                            ctx_stack.push(ParseCtx::SerSpPr);
                        }
                        (ParseCtx::Series, b"marker") => {
                            ctx_stack.push(ParseCtx::SerMarker);
                        }
                        (ParseCtx::Series, b"dLbls") => {
                            cur_ser_dlbls = DataLabelsBuilder::new();
                            ctx_stack.push(ParseCtx::SerDataLabels);
                        }
                        // spPr children.
                        (ParseCtx::SerSpPr, b"solidFill") => {
                            ctx_stack.push(ParseCtx::SerSpPrFill);
                        }
                        (ParseCtx::SerSpPr, b"ln") => {
                            // Parse line width from `w` attribute.
                            if let Some(ref mut ser) = cur_ser
                                && let Some(w) = attr_val(e, b"w")
                            {
                                ser.line_width = w.parse().ok();
                            }
                            ctx_stack.push(ParseCtx::SerSpPrLn);
                        }
                        (ParseCtx::SerSpPrLn, b"solidFill") => {
                            ctx_stack.push(ParseCtx::SerSpPrLnFill);
                        }
                        // Series text formula capture.
                        (ParseCtx::SerTx, b"f") | (ParseCtx::SerTx, b"strRef") => {
                            if local == b"f" {
                                capturing_text = true;
                                text_target = TextTarget::SerName;
                                text_buf.clear();
                            }
                            ctx_stack.push(ctx); // stay in SerTx
                        }
                        // Ref formula capture.
                        (ParseCtx::SerRef(kind), b"f") => {
                            capturing_text = true;
                            text_target = match kind {
                                SerRefKind::Cat => TextTarget::CatRef,
                                SerRefKind::Val | SerRefKind::YVal => TextTarget::ValRef,
                                SerRefKind::XVal => TextTarget::XValRef,
                                SerRefKind::BubbleSize => TextTarget::BubbleRef,
                            };
                            text_buf.clear();
                            ctx_stack.push(ctx); // stay in same ref context
                        }
                        // Ref sub-elements (strRef, numRef) — pass through.
                        (ParseCtx::SerRef(_), _) => {
                            ctx_stack.push(ctx);
                        }
                        // Axis children.
                        (ParseCtx::CatAxis, b"scaling") => {
                            ctx_stack.push(ParseCtx::AxisScaling(AxisKind::Cat));
                        }
                        (ParseCtx::ValAxis, b"scaling") => {
                            ctx_stack.push(ParseCtx::AxisScaling(AxisKind::Val));
                        }
                        (ParseCtx::CatAxis, b"title") => {
                            cur_cat_axis_title = TitleBuilder::new();
                            ctx_stack.push(ParseCtx::Title(TitleOwner::CatAxis));
                        }
                        (ParseCtx::ValAxis, b"title") => {
                            cur_val_axis_title = TitleBuilder::new();
                            ctx_stack.push(ParseCtx::Title(TitleOwner::ValAxis));
                        }
                        // Title text structure.
                        (ParseCtx::Title(owner), b"r") => {
                            ctx_stack.push(ParseCtx::TitleRun(owner));
                        }
                        (ParseCtx::Title(owner), b"pPr") => {
                            ctx_stack.push(ParseCtx::TitlePPr(owner));
                        }
                        (ParseCtx::TitlePPr(owner), b"defRPr") => {
                            // Parse font_size and bold from defRPr attributes.
                            let title_b = match owner {
                                TitleOwner::Chart => &mut cur_title,
                                TitleOwner::CatAxis => &mut cur_cat_axis_title,
                                TitleOwner::ValAxis => &mut cur_val_axis_title,
                            };
                            if let Some(sz) = attr_val(e, b"sz") {
                                title_b.font_size = sz.parse().ok();
                            }
                            if let Some(b) = attr_val(e, b"b") {
                                title_b.bold = Some(b == "1" || b == "true");
                            }
                            ctx_stack.push(ParseCtx::TitleDefRPr(owner));
                        }
                        (ParseCtx::TitleDefRPr(owner), b"solidFill") => {
                            ctx_stack.push(ParseCtx::TitleDefRPrFill(owner));
                        }
                        (ParseCtx::TitleRun(owner), b"rPr") => {
                            // Parse font_size and bold from rPr attributes.
                            let title_b = match owner {
                                TitleOwner::Chart => &mut cur_title,
                                TitleOwner::CatAxis => &mut cur_cat_axis_title,
                                TitleOwner::ValAxis => &mut cur_val_axis_title,
                            };
                            if let Some(sz) = attr_val(e, b"sz") {
                                title_b.font_size = sz.parse().ok();
                            }
                            if let Some(b) = attr_val(e, b"b") {
                                title_b.bold = Some(b == "1" || b == "true");
                            }
                            ctx_stack.push(ParseCtx::TitleRunPr(owner));
                        }
                        (ParseCtx::TitleRunPr(owner), b"solidFill") => {
                            ctx_stack.push(ParseCtx::TitleRunPrFill(owner));
                        }
                        (ParseCtx::TitleRun(owner), b"t") => {
                            capturing_text = true;
                            text_target = match owner {
                                TitleOwner::Chart => TextTarget::TitleText,
                                TitleOwner::CatAxis => TextTarget::CatAxisTitleText,
                                TitleOwner::ValAxis => TextTarget::ValAxisTitleText,
                            };
                            text_buf.clear();
                            ctx_stack.push(ParseCtx::TitleRun(owner));
                        }
                        // Generic: push same context to track depth.
                        _ => {
                            ctx_stack.push(ctx);
                        }
                    }
                }
                Ok(Event::Empty(ref e)) => {
                    let local = e.local_name();
                    let local = local.as_ref();
                    let ctx = current_ctx(&ctx_stack);

                    match (ctx, local) {
                        // Chart-type element attributes.
                        (ParseCtx::ChartTypeElem, b"barDir") => {
                            let val = attr_val_str(e);
                            if val == "bar" {
                                chart_type = Some(ChartType::Bar);
                                bar_dir_horizontal = Some(true);
                            } else {
                                chart_type = Some(ChartType::Column);
                                bar_dir_horizontal = Some(false);
                            }
                        }
                        (ParseCtx::ChartTypeElem, b"grouping") => {
                            grouping = ChartGrouping::from_xml(&attr_val_str(e));
                        }
                        (ParseCtx::ChartTypeElem, b"scatterStyle") => {
                            scatter_style = ScatterStyle::from_xml(&attr_val_str(e));
                        }
                        (ParseCtx::ChartTypeElem, b"radarStyle") => {
                            radar_style = RadarStyle::from_xml(&attr_val_str(e));
                        }
                        (ParseCtx::ChartTypeElem, b"holeSize") => {
                            hole_size = attr_val_str(e).parse().ok();
                        }
                        (ParseCtx::ChartTypeElem, b"axId") => {
                            // Axis IDs inside chart type elem — ignored (parsed from axis elements).
                        }
                        // Series attributes.
                        (ParseCtx::Series, b"idx") => {
                            if let Some(ref mut ser) = cur_ser {
                                ser.idx = attr_val_str(e).parse().unwrap_or(0);
                            }
                        }
                        (ParseCtx::Series, b"order") => {
                            if let Some(ref mut ser) = cur_ser {
                                ser.order = attr_val_str(e).parse().unwrap_or(0);
                            }
                        }
                        (ParseCtx::Series, b"smooth") => {
                            if let Some(ref mut ser) = cur_ser {
                                let val = attr_val_str(e);
                                ser.smooth = Some(val == "1" || val == "true");
                            }
                        }
                        (ParseCtx::Series, b"explosion") => {
                            if let Some(ref mut ser) = cur_ser {
                                ser.explosion = attr_val_str(e).parse().ok();
                            }
                        }
                        // Fill color in spPr.
                        (ParseCtx::SerSpPrFill, b"srgbClr") => {
                            if let Some(ref mut ser) = cur_ser {
                                ser.fill_color = attr_val(e, b"val");
                            }
                        }
                        // Line color in spPr > ln.
                        (ParseCtx::SerSpPrLnFill, b"srgbClr") => {
                            if let Some(ref mut ser) = cur_ser {
                                ser.line_color = attr_val(e, b"val");
                            }
                        }
                        // Line width from ln element.
                        (ParseCtx::SerSpPrLn, b"ln") => {
                            // This shouldn't happen (ln is usually Start), but handle anyway.
                        }
                        // Marker symbol.
                        (ParseCtx::SerMarker, b"symbol") => {
                            if let Some(ref mut ser) = cur_ser {
                                ser.marker = MarkerStyle::from_xml(&attr_val_str(e));
                            }
                        }
                        // Data labels (chart-level).
                        (ParseCtx::DataLabels, b"showVal") => {
                            cur_dlbls.show_val = attr_val_str(e) == "1";
                        }
                        (ParseCtx::DataLabels, b"showCatName") => {
                            cur_dlbls.show_cat_name = attr_val_str(e) == "1";
                        }
                        (ParseCtx::DataLabels, b"showSerName") => {
                            cur_dlbls.show_ser_name = attr_val_str(e) == "1";
                        }
                        (ParseCtx::DataLabels, b"showPercent") => {
                            cur_dlbls.show_percent = attr_val_str(e) == "1";
                        }
                        (ParseCtx::DataLabels, b"showLeaderLines") => {
                            cur_dlbls.show_leader_lines = attr_val_str(e) == "1";
                        }
                        (ParseCtx::DataLabels, b"numFmt") => {
                            cur_dlbls.num_fmt = attr_val(e, b"formatCode");
                        }
                        // Data labels (series-level).
                        (ParseCtx::SerDataLabels, b"showVal") => {
                            cur_ser_dlbls.show_val = attr_val_str(e) == "1";
                        }
                        (ParseCtx::SerDataLabels, b"showCatName") => {
                            cur_ser_dlbls.show_cat_name = attr_val_str(e) == "1";
                        }
                        (ParseCtx::SerDataLabels, b"showSerName") => {
                            cur_ser_dlbls.show_ser_name = attr_val_str(e) == "1";
                        }
                        (ParseCtx::SerDataLabels, b"showPercent") => {
                            cur_ser_dlbls.show_percent = attr_val_str(e) == "1";
                        }
                        (ParseCtx::SerDataLabels, b"showLeaderLines") => {
                            cur_ser_dlbls.show_leader_lines = attr_val_str(e) == "1";
                        }
                        (ParseCtx::SerDataLabels, b"numFmt") => {
                            cur_ser_dlbls.num_fmt = attr_val(e, b"formatCode");
                        }
                        // Axis attributes.
                        (ParseCtx::CatAxis, b"axId") => {
                            cur_cat_axis.id = attr_val_str(e).parse().ok();
                        }
                        (ParseCtx::CatAxis, b"delete") => {
                            cur_cat_axis.delete = attr_val_str(e) == "1";
                        }
                        (ParseCtx::CatAxis, b"axPos") => {
                            cur_cat_axis.position = AxisPosition::from_xml(&attr_val_str(e));
                        }
                        (ParseCtx::CatAxis, b"majorGridlines") => {
                            cur_cat_axis.major_gridlines = true;
                        }
                        (ParseCtx::CatAxis, b"minorGridlines") => {
                            cur_cat_axis.minor_gridlines = true;
                        }
                        (ParseCtx::CatAxis, b"numFmt") => {
                            cur_cat_axis.num_fmt = attr_val(e, b"formatCode");
                            if let Some(sl) = attr_val(e, b"sourceLinked") {
                                cur_cat_axis.source_linked = sl == "1";
                            }
                        }
                        (ParseCtx::CatAxis, b"majorTickMark") => {
                            cur_cat_axis.major_tick_mark = TickMark::from_xml(&attr_val_str(e));
                        }
                        (ParseCtx::CatAxis, b"minorTickMark") => {
                            cur_cat_axis.minor_tick_mark = TickMark::from_xml(&attr_val_str(e));
                        }
                        (ParseCtx::CatAxis, b"tickLblPos") => {
                            cur_cat_axis.tick_lbl_pos = TickLabelPosition::from_xml(&attr_val_str(e));
                        }
                        (ParseCtx::CatAxis, b"crossAx") => {
                            cur_cat_axis.cross_ax = attr_val_str(e).parse().ok();
                        }
                        (ParseCtx::CatAxis, b"crossesAt") => {
                            cur_cat_axis.crosses_at = attr_val_str(e).parse().ok();
                        }
                        (ParseCtx::CatAxis, b"crosses") => {
                            // autoZero → None (default).
                        }
                        (ParseCtx::CatAxis, b"majorUnit") => {
                            cur_cat_axis.major_unit = attr_val_str(e).parse().ok();
                        }
                        (ParseCtx::CatAxis, b"minorUnit") => {
                            cur_cat_axis.minor_unit = attr_val_str(e).parse().ok();
                        }
                        // ValAxis.
                        (ParseCtx::ValAxis, b"axId") => {
                            cur_val_axis.id = attr_val_str(e).parse().ok();
                        }
                        (ParseCtx::ValAxis, b"delete") => {
                            cur_val_axis.delete = attr_val_str(e) == "1";
                        }
                        (ParseCtx::ValAxis, b"axPos") => {
                            cur_val_axis.position = AxisPosition::from_xml(&attr_val_str(e));
                        }
                        (ParseCtx::ValAxis, b"majorGridlines") => {
                            cur_val_axis.major_gridlines = true;
                        }
                        (ParseCtx::ValAxis, b"minorGridlines") => {
                            cur_val_axis.minor_gridlines = true;
                        }
                        (ParseCtx::ValAxis, b"numFmt") => {
                            cur_val_axis.num_fmt = attr_val(e, b"formatCode");
                            if let Some(sl) = attr_val(e, b"sourceLinked") {
                                cur_val_axis.source_linked = sl == "1";
                            }
                        }
                        (ParseCtx::ValAxis, b"majorTickMark") => {
                            cur_val_axis.major_tick_mark = TickMark::from_xml(&attr_val_str(e));
                        }
                        (ParseCtx::ValAxis, b"minorTickMark") => {
                            cur_val_axis.minor_tick_mark = TickMark::from_xml(&attr_val_str(e));
                        }
                        (ParseCtx::ValAxis, b"tickLblPos") => {
                            cur_val_axis.tick_lbl_pos = TickLabelPosition::from_xml(&attr_val_str(e));
                        }
                        (ParseCtx::ValAxis, b"crossAx") => {
                            cur_val_axis.cross_ax = attr_val_str(e).parse().ok();
                        }
                        (ParseCtx::ValAxis, b"crossesAt") => {
                            cur_val_axis.crosses_at = attr_val_str(e).parse().ok();
                        }
                        (ParseCtx::ValAxis, b"crosses") => {
                            // autoZero → None.
                        }
                        (ParseCtx::ValAxis, b"majorUnit") => {
                            cur_val_axis.major_unit = attr_val_str(e).parse().ok();
                        }
                        (ParseCtx::ValAxis, b"minorUnit") => {
                            cur_val_axis.minor_unit = attr_val_str(e).parse().ok();
                        }
                        // Axis scaling.
                        (ParseCtx::AxisScaling(AxisKind::Cat), b"orientation") => {
                            cur_cat_axis.reversed = attr_val_str(e) == "maxMin";
                        }
                        (ParseCtx::AxisScaling(AxisKind::Cat), b"min") => {
                            cur_cat_axis.min = attr_val_str(e).parse().ok();
                        }
                        (ParseCtx::AxisScaling(AxisKind::Cat), b"max") => {
                            cur_cat_axis.max = attr_val_str(e).parse().ok();
                        }
                        (ParseCtx::AxisScaling(AxisKind::Cat), b"logBase") => {
                            cur_cat_axis.log_base = attr_val_str(e).parse().ok();
                        }
                        (ParseCtx::AxisScaling(AxisKind::Val), b"orientation") => {
                            cur_val_axis.reversed = attr_val_str(e) == "maxMin";
                        }
                        (ParseCtx::AxisScaling(AxisKind::Val), b"min") => {
                            cur_val_axis.min = attr_val_str(e).parse().ok();
                        }
                        (ParseCtx::AxisScaling(AxisKind::Val), b"max") => {
                            cur_val_axis.max = attr_val_str(e).parse().ok();
                        }
                        (ParseCtx::AxisScaling(AxisKind::Val), b"logBase") => {
                            cur_val_axis.log_base = attr_val_str(e).parse().ok();
                        }
                        // Legend.
                        (ParseCtx::Legend, b"legendPos") => {
                            cur_legend.position = LegendPosition::from_xml(&attr_val_str(e));
                        }
                        (ParseCtx::Legend, b"overlay") => {
                            cur_legend.overlay = attr_val_str(e) == "1";
                        }
                        // Title overlay.
                        (ParseCtx::Title(owner), b"overlay") => {
                            let val = attr_val_str(e) == "1";
                            match owner {
                                TitleOwner::Chart => cur_title.overlay = val,
                                TitleOwner::CatAxis => cur_cat_axis_title.overlay = val,
                                TitleOwner::ValAxis => cur_val_axis_title.overlay = val,
                            }
                        }
                        // Title defRPr (Empty variant).
                        (ParseCtx::TitlePPr(owner), b"defRPr") => {
                            let title_b = match owner {
                                TitleOwner::Chart => &mut cur_title,
                                TitleOwner::CatAxis => &mut cur_cat_axis_title,
                                TitleOwner::ValAxis => &mut cur_val_axis_title,
                            };
                            if let Some(sz) = attr_val(e, b"sz") {
                                title_b.font_size = sz.parse().ok();
                            }
                            if let Some(b) = attr_val(e, b"b") {
                                title_b.bold = Some(b == "1" || b == "true");
                            }
                        }
                        // Title rPr (Empty variant).
                        (ParseCtx::TitleRun(owner), b"rPr") => {
                            let title_b = match owner {
                                TitleOwner::Chart => &mut cur_title,
                                TitleOwner::CatAxis => &mut cur_cat_axis_title,
                                TitleOwner::ValAxis => &mut cur_val_axis_title,
                            };
                            if let Some(sz) = attr_val(e, b"sz") {
                                title_b.font_size = sz.parse().ok();
                            }
                            if let Some(b) = attr_val(e, b"b") {
                                title_b.bold = Some(b == "1" || b == "true");
                            }
                        }
                        // Color in title defRPr fill.
                        (ParseCtx::TitleDefRPrFill(owner), b"srgbClr") => {
                            let title_b = match owner {
                                TitleOwner::Chart => &mut cur_title,
                                TitleOwner::CatAxis => &mut cur_cat_axis_title,
                                TitleOwner::ValAxis => &mut cur_val_axis_title,
                            };
                            title_b.color = attr_val(e, b"val");
                        }
                        // Color in title rPr fill.
                        (ParseCtx::TitleRunPrFill(owner), b"srgbClr") => {
                            let title_b = match owner {
                                TitleOwner::Chart => &mut cur_title,
                                TitleOwner::CatAxis => &mut cur_cat_axis_title,
                                TitleOwner::ValAxis => &mut cur_val_axis_title,
                            };
                            title_b.color = attr_val(e, b"val");
                        }
                        // Manual layout values.
                        (ParseCtx::ManualLayoutCtx, b"x") => {
                            layout_x = attr_val_str(e).parse().ok();
                        }
                        (ParseCtx::ManualLayoutCtx, b"y") => {
                            layout_y = attr_val_str(e).parse().ok();
                        }
                        (ParseCtx::ManualLayoutCtx, b"w") => {
                            layout_w = attr_val_str(e).parse().ok();
                        }
                        (ParseCtx::ManualLayoutCtx, b"h") => {
                            layout_h = attr_val_str(e).parse().ok();
                        }
                        // Style ID (on chartSpace level).
                        (_, b"style") => {
                            style_id = attr_val_str(e).parse().ok();
                        }
                        _ => {}
                    }
                }
                Ok(Event::End(ref e)) => {
                    let local = e.local_name();
                    let local = local.as_ref();
                    let ctx = current_ctx(&ctx_stack);

                    match (ctx, local) {
                        (ParseCtx::Series, b"ser") => {
                            if let Some(ser) = cur_ser.take() {
                                series.push(ser);
                            }
                            ctx_stack.pop();
                        }
                        (ParseCtx::DataLabels, b"dLbls") => {
                            data_labels = Some(cur_dlbls.build());
                            ctx_stack.pop();
                        }
                        (ParseCtx::SerDataLabels, b"dLbls") => {
                            if let Some(ref mut ser) = cur_ser {
                                ser.data_labels = Some(cur_ser_dlbls.build());
                            }
                            ctx_stack.pop();
                        }
                        (ParseCtx::CatAxis, b"catAx") => {
                            let builder = std::mem::replace(&mut cur_cat_axis, AxisBuilder::new());
                            cat_axis = Some(builder.build_with_title(
                                cur_cat_axis_title.build_option(),
                            ));
                            ctx_stack.pop();
                        }
                        (ParseCtx::ValAxis, b"valAx") => {
                            let builder = std::mem::replace(&mut cur_val_axis, AxisBuilder::new());
                            val_axis = Some(builder.build_with_title(
                                cur_val_axis_title.build_option(),
                            ));
                            ctx_stack.pop();
                        }
                        (ParseCtx::Legend, b"legend") => {
                            legend = Some(cur_legend.build());
                            ctx_stack.pop();
                        }
                        (ParseCtx::Title(TitleOwner::Chart), b"title") => {
                            title = cur_title.build_option();
                            ctx_stack.pop();
                        }
                        (ParseCtx::Title(TitleOwner::CatAxis), b"title") => {
                            // Title is stored in axis builder on close.
                            ctx_stack.pop();
                        }
                        (ParseCtx::Title(TitleOwner::ValAxis), b"title") => {
                            ctx_stack.pop();
                        }
                        // Text capture end.
                        (_, b"f") if capturing_text && matches!(text_target, TextTarget::SerName | TextTarget::CatRef | TextTarget::ValRef | TextTarget::XValRef | TextTarget::BubbleRef) => {
                            if let Some(ref mut ser) = cur_ser {
                                let val = std::mem::take(&mut text_buf);
                                match text_target {
                                    TextTarget::SerName => ser.name = Some(val),
                                    TextTarget::CatRef => ser.cat_ref = Some(val),
                                    TextTarget::ValRef => ser.val_ref = val,
                                    TextTarget::XValRef => ser.x_val_ref = Some(val),
                                    TextTarget::BubbleRef => ser.bubble_size_ref = Some(val),
                                    _ => {}
                                }
                            }
                            capturing_text = false;
                            text_target = TextTarget::None;
                            ctx_stack.pop();
                        }
                        (_, b"t") if capturing_text && matches!(text_target, TextTarget::TitleText | TextTarget::CatAxisTitleText | TextTarget::ValAxisTitleText) => {
                            let val = std::mem::take(&mut text_buf);
                            match text_target {
                                TextTarget::TitleText => cur_title.text = val,
                                TextTarget::CatAxisTitleText => cur_cat_axis_title.text = val,
                                TextTarget::ValAxisTitleText => cur_val_axis_title.text = val,
                                _ => {}
                            }
                            capturing_text = false;
                            text_target = TextTarget::None;
                            ctx_stack.pop();
                        }
                        // Manual layout end.
                        (ParseCtx::ManualLayoutCtx, b"layout") => {
                            if let (Some(x), Some(y), Some(w), Some(h)) =
                                (layout_x, layout_y, layout_w, layout_h)
                            {
                                plot_area_layout = Some(ManualLayout { x, y, w, h });
                            }
                            ctx_stack.pop();
                        }
                        _ => {
                            // Pop context for any End that matches the depth.
                            ctx_stack.pop();
                        }
                    }
                }
                Ok(Event::Text(ref e)) => {
                    if capturing_text {
                        text_buf.push_str(
                            std::str::from_utf8(e.as_ref()).unwrap_or_default(),
                        );
                    }
                }
                Ok(Event::GeneralRef(ref e)) => {
                    if capturing_text {
                        push_entity(&mut text_buf, e.as_ref());
                    }
                }
                Ok(Event::Eof) => break,
                Err(err) => {
                    return Err(ModernXlsxError::XmlParse(format!(
                        "chart parse error: {err}"
                    )));
                }
                _ => {}
            }
            buf.clear();
        }

        Ok(ChartData {
            chart_type: chart_type.unwrap_or(ChartType::Bar),
            title,
            series,
            cat_axis,
            val_axis,
            legend,
            data_labels,
            grouping,
            scatter_style,
            radar_style,
            hole_size,
            bar_dir_horizontal,
            style_id,
            plot_area_layout,
        })
    }
}

/// What text we are currently capturing.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum TextTarget {
    None,
    SerName,
    CatRef,
    ValRef,
    XValRef,
    BubbleRef,
    TitleText,
    CatAxisTitleText,
    ValAxisTitleText,
}

/// Builder for accumulating axis data during parsing.
struct AxisBuilder {
    id: Option<u32>,
    cross_ax: Option<u32>,
    num_fmt: Option<String>,
    source_linked: bool,
    min: Option<f64>,
    max: Option<f64>,
    major_unit: Option<f64>,
    minor_unit: Option<f64>,
    log_base: Option<f64>,
    reversed: bool,
    tick_lbl_pos: Option<TickLabelPosition>,
    major_tick_mark: Option<TickMark>,
    minor_tick_mark: Option<TickMark>,
    major_gridlines: bool,
    minor_gridlines: bool,
    delete: bool,
    position: Option<AxisPosition>,
    crosses_at: Option<f64>,
}

impl AxisBuilder {
    fn new() -> Self {
        Self {
            id: None,
            cross_ax: None,
            num_fmt: None,
            source_linked: false,
            min: None,
            max: None,
            major_unit: None,
            minor_unit: None,
            log_base: None,
            reversed: false,
            tick_lbl_pos: None,
            major_tick_mark: None,
            minor_tick_mark: None,
            major_gridlines: false,
            minor_gridlines: false,
            delete: false,
            position: None,
            crosses_at: None,
        }
    }

    fn build_with_title(self, title: Option<ChartTitle>) -> ChartAxis {
        ChartAxis {
            id: self.id.unwrap_or(0),
            cross_ax: self.cross_ax.unwrap_or(0),
            title,
            num_fmt: self.num_fmt,
            source_linked: self.source_linked,
            min: self.min,
            max: self.max,
            major_unit: self.major_unit,
            minor_unit: self.minor_unit,
            log_base: self.log_base,
            reversed: self.reversed,
            tick_lbl_pos: self.tick_lbl_pos,
            major_tick_mark: self.major_tick_mark,
            minor_tick_mark: self.minor_tick_mark,
            major_gridlines: self.major_gridlines,
            minor_gridlines: self.minor_gridlines,
            delete: self.delete,
            position: self.position,
            crosses_at: self.crosses_at,
            font_size: None,
        }
    }
}

/// Builder for accumulating data labels during parsing.
struct DataLabelsBuilder {
    show_val: bool,
    show_cat_name: bool,
    show_ser_name: bool,
    show_percent: bool,
    num_fmt: Option<String>,
    show_leader_lines: bool,
}

impl DataLabelsBuilder {
    fn new() -> Self {
        Self {
            show_val: false,
            show_cat_name: false,
            show_ser_name: false,
            show_percent: false,
            num_fmt: None,
            show_leader_lines: false,
        }
    }

    fn build(&self) -> DataLabels {
        DataLabels {
            show_val: self.show_val,
            show_cat_name: self.show_cat_name,
            show_ser_name: self.show_ser_name,
            show_percent: self.show_percent,
            num_fmt: self.num_fmt.clone(),
            show_leader_lines: self.show_leader_lines,
        }
    }
}

/// Builder for accumulating legend during parsing.
struct LegendBuilder {
    position: Option<LegendPosition>,
    overlay: bool,
}

impl LegendBuilder {
    fn new() -> Self {
        Self {
            position: None,
            overlay: false,
        }
    }

    fn build(&self) -> ChartLegend {
        ChartLegend {
            position: self.position.unwrap_or(LegendPosition::Right),
            overlay: self.overlay,
        }
    }
}

/// Builder for accumulating title text during parsing.
struct TitleBuilder {
    text: String,
    overlay: bool,
    font_size: Option<u32>,
    bold: Option<bool>,
    color: Option<String>,
}

impl TitleBuilder {
    fn new() -> Self {
        Self {
            text: String::new(),
            overlay: false,
            font_size: None,
            bold: None,
            color: None,
        }
    }

    fn build_option(&self) -> Option<ChartTitle> {
        if self.text.is_empty() {
            return None;
        }
        Some(ChartTitle {
            text: self.text.clone(),
            overlay: self.overlay,
            font_size: self.font_size,
            bold: self.bold,
            color: self.color.clone(),
        })
    }
}

// ---------------------------------------------------------------------------
// Drawing XML Parser — parse drawing anchors for chart references
// ---------------------------------------------------------------------------

/// Parse `xl/drawings/drawing{n}.xml` to extract chart anchors and chart rIds.
///
/// Returns a vec of `(ChartAnchor, rId)` pairs for each `<xdr:twoCellAnchor>`
/// that contains a chart reference.
pub fn parse_drawing_anchors(data: &[u8]) -> Result<Vec<(ChartAnchor, String)>> {
    let mut reader = Reader::from_reader(data);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::with_capacity(512);

    let mut result: Vec<(ChartAnchor, String)> = Vec::new();

    // State tracking.
    let mut in_anchor = false;
    let mut in_from = false;
    let mut in_to = false;
    let mut from_col: u32 = 0;
    let mut from_col_off: u64 = 0;
    let mut from_row: u32 = 0;
    let mut from_row_off: u64 = 0;
    let mut to_col: u32 = 0;
    let mut to_col_off: u64 = 0;
    let mut to_row: u32 = 0;
    let mut to_row_off: u64 = 0;
    let mut chart_r_id: Option<String> = None;

    // Text capture for position elements.
    let mut text_buf = String::new();
    #[derive(Clone, Copy, PartialEq, Eq)]
    enum DrawingTextTarget {
        None,
        FromCol,
        FromColOff,
        FromRow,
        FromRowOff,
        ToCol,
        ToColOff,
        ToRow,
        ToRowOff,
    }
    let mut text_target = DrawingTextTarget::None;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) => {
                let local = e.local_name();
                let local = local.as_ref();
                match local {
                    b"twoCellAnchor" => {
                        in_anchor = true;
                        from_col = 0;
                        from_col_off = 0;
                        from_row = 0;
                        from_row_off = 0;
                        to_col = 0;
                        to_col_off = 0;
                        to_row = 0;
                        to_row_off = 0;
                        chart_r_id = None;
                    }
                    b"from" if in_anchor => {
                        in_from = true;
                    }
                    b"to" if in_anchor => {
                        in_to = true;
                    }
                    b"col" if in_anchor && (in_from || in_to) => {
                        text_buf.clear();
                        text_target = if in_from {
                            DrawingTextTarget::FromCol
                        } else {
                            DrawingTextTarget::ToCol
                        };
                    }
                    b"colOff" if in_anchor && (in_from || in_to) => {
                        text_buf.clear();
                        text_target = if in_from {
                            DrawingTextTarget::FromColOff
                        } else {
                            DrawingTextTarget::ToColOff
                        };
                    }
                    b"row" if in_anchor && (in_from || in_to) => {
                        text_buf.clear();
                        text_target = if in_from {
                            DrawingTextTarget::FromRow
                        } else {
                            DrawingTextTarget::ToRow
                        };
                    }
                    b"rowOff" if in_anchor && (in_from || in_to) => {
                        text_buf.clear();
                        text_target = if in_from {
                            DrawingTextTarget::FromRowOff
                        } else {
                            DrawingTextTarget::ToRowOff
                        };
                    }
                    _ => {}
                }
            }
            Ok(Event::Empty(ref e)) => {
                let local = e.local_name();
                let local = local.as_ref();
                // <c:chart r:id="rId1"/> — may also appear as just `chart`.
                if local == b"chart" && in_anchor {
                    // Extract r:id attribute.
                    for attr in e.attributes().flatten() {
                        let key = attr.key.as_ref();
                        // Match both `r:id` and bare `id` with namespace prefix.
                        if key == b"r:id" || key.ends_with(b":id") {
                            chart_r_id = Some(
                                std::str::from_utf8(&attr.value)
                                    .unwrap_or_default()
                                    .to_owned(),
                            );
                            break;
                        }
                    }
                }
            }
            Ok(Event::End(ref e)) => {
                let local = e.local_name();
                let local = local.as_ref();
                match local {
                    b"twoCellAnchor" => {
                        if let Some(r_id) = chart_r_id.take() {
                            result.push((
                                ChartAnchor {
                                    from_col,
                                    from_row,
                                    from_col_off,
                                    from_row_off,
                                    to_col,
                                    to_row,
                                    to_col_off,
                                    to_row_off,
                                },
                                r_id,
                            ));
                        }
                        in_anchor = false;
                        in_from = false;
                        in_to = false;
                    }
                    b"from" => {
                        in_from = false;
                    }
                    b"to" => {
                        in_to = false;
                    }
                    b"col" | b"colOff" | b"row" | b"rowOff" => {
                        let val_str = std::mem::take(&mut text_buf);
                        match text_target {
                            DrawingTextTarget::FromCol => from_col = val_str.parse().unwrap_or(0),
                            DrawingTextTarget::FromColOff => from_col_off = val_str.parse().unwrap_or(0),
                            DrawingTextTarget::FromRow => from_row = val_str.parse().unwrap_or(0),
                            DrawingTextTarget::FromRowOff => from_row_off = val_str.parse().unwrap_or(0),
                            DrawingTextTarget::ToCol => to_col = val_str.parse().unwrap_or(0),
                            DrawingTextTarget::ToColOff => to_col_off = val_str.parse().unwrap_or(0),
                            DrawingTextTarget::ToRow => to_row = val_str.parse().unwrap_or(0),
                            DrawingTextTarget::ToRowOff => to_row_off = val_str.parse().unwrap_or(0),
                            DrawingTextTarget::None => {}
                        }
                        text_target = DrawingTextTarget::None;
                    }
                    _ => {}
                }
            }
            Ok(Event::Text(ref e)) => {
                if text_target != DrawingTextTarget::None {
                    text_buf.push_str(
                        std::str::from_utf8(e.as_ref()).unwrap_or_default(),
                    );
                }
            }
            Ok(Event::Eof) => break,
            Err(err) => {
                return Err(ModernXlsxError::XmlParse(format!(
                    "drawing parse error: {err}"
                )));
            }
            _ => {}
        }
        buf.clear();
    }

    Ok(result)
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

    // =======================================================================
    // Drawing XML tests
    // =======================================================================

    #[test]
    fn test_generate_drawing_xml_single_chart() {
        let charts = vec![WorksheetChart {
            chart: ChartData {
                chart_type: ChartType::Bar,
                title: None, series: vec![], cat_axis: None, val_axis: None,
                legend: None, data_labels: None, grouping: None,
                scatter_style: None, radar_style: None, hole_size: None,
                bar_dir_horizontal: None, style_id: None, plot_area_layout: None,
            },
            anchor: ChartAnchor {
                from_col: 0, from_row: 0, from_col_off: 0, from_row_off: 0,
                to_col: 8, to_row: 15, to_col_off: 0, to_row_off: 0,
            },
        }];
        let r_ids = vec!["rId1".to_string()];
        let xml = ChartAnchor::generate_drawing_xml(&charts, &r_ids).unwrap();
        let xml_str = std::str::from_utf8(&xml).unwrap();

        // Check namespaces.
        assert!(xml_str.contains("xmlns:xdr=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\""));
        assert!(xml_str.contains("xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\""));
        assert!(xml_str.contains("xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\""));

        // Check anchor structure.
        assert!(xml_str.contains("<xdr:twoCellAnchor>"));
        assert!(xml_str.contains("<xdr:from>"));
        assert!(xml_str.contains("<xdr:to>"));
        assert!(xml_str.contains("<xdr:col>0</xdr:col>"));
        assert!(xml_str.contains("<xdr:col>8</xdr:col>"));
        assert!(xml_str.contains("<xdr:row>15</xdr:row>"));

        // Check graphic frame.
        assert!(xml_str.contains(r#"<xdr:cNvPr id="2" name="Chart 1"/>"#));
        assert!(xml_str.contains(r#"r:id="rId1""#));
        assert!(xml_str.contains("<xdr:clientData/>"));
    }

    #[test]
    fn test_generate_drawing_xml_multiple_charts() {
        let charts = vec![
            WorksheetChart {
                chart: ChartData {
                    chart_type: ChartType::Bar,
                    title: None, series: vec![], cat_axis: None, val_axis: None,
                    legend: None, data_labels: None, grouping: None,
                    scatter_style: None, radar_style: None, hole_size: None,
                    bar_dir_horizontal: None, style_id: None, plot_area_layout: None,
                },
                anchor: ChartAnchor {
                    from_col: 0, from_row: 0, from_col_off: 0, from_row_off: 0,
                    to_col: 8, to_row: 15, to_col_off: 0, to_row_off: 0,
                },
            },
            WorksheetChart {
                chart: ChartData {
                    chart_type: ChartType::Line,
                    title: None, series: vec![], cat_axis: None, val_axis: None,
                    legend: None, data_labels: None, grouping: None,
                    scatter_style: None, radar_style: None, hole_size: None,
                    bar_dir_horizontal: None, style_id: None, plot_area_layout: None,
                },
                anchor: ChartAnchor {
                    from_col: 10, from_row: 0, from_col_off: 100, from_row_off: 200,
                    to_col: 18, to_row: 20, to_col_off: 300, to_row_off: 400,
                },
            },
        ];
        let r_ids = vec!["rId1".to_string(), "rId2".to_string()];
        let xml = ChartAnchor::generate_drawing_xml(&charts, &r_ids).unwrap();
        let xml_str = std::str::from_utf8(&xml).unwrap();

        // Two anchors.
        assert_eq!(xml_str.matches("<xdr:twoCellAnchor>").count(), 2);

        // Unique cNvPr ids.
        assert!(xml_str.contains(r#"id="2" name="Chart 1""#));
        assert!(xml_str.contains(r#"id="3" name="Chart 2""#));

        // Both rIds.
        assert!(xml_str.contains(r#"r:id="rId1""#));
        assert!(xml_str.contains(r#"r:id="rId2""#));

        // Second anchor offsets.
        assert!(xml_str.contains("<xdr:col>10</xdr:col>"));
        assert!(xml_str.contains("<xdr:colOff>100</xdr:colOff>"));
        assert!(xml_str.contains("<xdr:rowOff>200</xdr:rowOff>"));
    }

    // =======================================================================
    // XML Parse roundtrip tests
    // =======================================================================

    #[test]
    fn test_chart_parse_roundtrip_bar() {
        let original = ChartData {
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
            style_id: Some(2),
            plot_area_layout: None,
        };
        let xml = original.to_xml().unwrap();
        let parsed = ChartData::parse(&xml).unwrap();
        assert_eq!(parsed.chart_type, ChartType::Bar);
        assert_eq!(parsed.series.len(), 1);
        assert_eq!(parsed.series[0].val_ref, "Sheet1!$B$2:$B$5");
        assert_eq!(parsed.series[0].name.as_deref(), Some("Revenue"));
        assert_eq!(parsed.series[0].fill_color.as_deref(), Some("4472C4"));
        assert_eq!(parsed.grouping, Some(ChartGrouping::Clustered));
        assert_eq!(parsed.bar_dir_horizontal, Some(true));
        assert!(parsed.cat_axis.is_some());
        assert!(parsed.val_axis.is_some());
        assert_eq!(parsed.legend.as_ref().unwrap().position, LegendPosition::Bottom);
        assert_eq!(parsed.style_id, Some(2));
    }

    #[test]
    fn test_pie_chart_parse_roundtrip() {
        let original = ChartData {
            chart_type: ChartType::Pie,
            title: Some(ChartTitle { text: "Distribution".into(), overlay: false, font_size: Some(1400), bold: Some(true), color: None }),
            series: vec![ChartSeries {
                idx: 0, order: 0, name: None,
                cat_ref: Some("Sheet1!$A$2:$A$5".into()),
                val_ref: "Sheet1!$B$2:$B$5".into(),
                x_val_ref: None, bubble_size_ref: None,
                fill_color: None, line_color: None, line_width: None,
                marker: None, smooth: None,
                explosion: Some(25),
                data_labels: None,
            }],
            cat_axis: None, val_axis: None,
            legend: None,
            data_labels: Some(DataLabels { show_val: true, show_cat_name: false, show_ser_name: false, show_percent: true, num_fmt: None, show_leader_lines: true }),
            grouping: None, scatter_style: None, radar_style: None, hole_size: None,
            bar_dir_horizontal: None, style_id: None, plot_area_layout: None,
        };
        let xml = original.to_xml().unwrap();
        let parsed = ChartData::parse(&xml).unwrap();
        assert_eq!(parsed.chart_type, ChartType::Pie);
        assert!(parsed.cat_axis.is_none());
        assert!(parsed.val_axis.is_none());
        assert_eq!(parsed.series[0].explosion, Some(25));
        assert!(parsed.data_labels.as_ref().unwrap().show_val);
        assert!(parsed.data_labels.as_ref().unwrap().show_percent);
    }

    #[test]
    fn test_scatter_chart_parse_roundtrip() {
        let original = ChartData {
            chart_type: ChartType::Scatter,
            title: None,
            series: vec![ChartSeries {
                idx: 0, order: 0, name: None, cat_ref: None,
                val_ref: "Sheet1!$B$2:$B$5".into(),
                x_val_ref: Some("Sheet1!$A$2:$A$5".into()),
                bubble_size_ref: None,
                fill_color: None, line_color: None, line_width: None,
                marker: Some(MarkerStyle::Diamond), smooth: None, explosion: None,
                data_labels: None,
            }],
            cat_axis: None,
            val_axis: Some(ChartAxis { id: 0, cross_ax: 1, title: None, num_fmt: None, source_linked: false, min: None, max: None, major_unit: None, minor_unit: None, log_base: None, reversed: false, tick_lbl_pos: None, major_tick_mark: None, minor_tick_mark: None, major_gridlines: true, minor_gridlines: false, delete: false, position: None, crosses_at: None, font_size: None }),
            legend: None, data_labels: None, grouping: None,
            scatter_style: Some(ScatterStyle::LineMarker),
            radar_style: None, hole_size: None,
            bar_dir_horizontal: None, style_id: None, plot_area_layout: None,
        };
        let xml = original.to_xml().unwrap();
        let parsed = ChartData::parse(&xml).unwrap();
        assert_eq!(parsed.chart_type, ChartType::Scatter);
        assert_eq!(parsed.scatter_style, Some(ScatterStyle::LineMarker));
        assert_eq!(parsed.series[0].x_val_ref.as_deref(), Some("Sheet1!$A$2:$A$5"));
        assert_eq!(parsed.series[0].marker, Some(MarkerStyle::Diamond));
    }

    #[test]
    fn test_drawing_anchors_parse_roundtrip() {
        let charts = vec![
            WorksheetChart {
                chart: ChartData {
                    chart_type: ChartType::Line,
                    title: None, series: vec![], cat_axis: None, val_axis: None,
                    legend: None, data_labels: None, grouping: None,
                    scatter_style: None, radar_style: None, hole_size: None,
                    bar_dir_horizontal: None, style_id: None, plot_area_layout: None,
                },
                anchor: ChartAnchor {
                    from_col: 2, from_row: 5, from_col_off: 100, from_row_off: 200,
                    to_col: 12, to_row: 25, to_col_off: 300, to_row_off: 400,
                },
            },
        ];
        let r_ids = vec!["rId1".into()];
        let drawing_xml = ChartAnchor::generate_drawing_xml(&charts, &r_ids).unwrap();
        let anchors = parse_drawing_anchors(&drawing_xml).unwrap();
        assert_eq!(anchors.len(), 1);
        assert_eq!(anchors[0].0.from_col, 2);
        assert_eq!(anchors[0].0.from_row, 5);
        assert_eq!(anchors[0].0.from_col_off, 100);
        assert_eq!(anchors[0].0.from_row_off, 200);
        assert_eq!(anchors[0].0.to_col, 12);
        assert_eq!(anchors[0].0.to_row, 25);
        assert_eq!(anchors[0].0.to_col_off, 300);
        assert_eq!(anchors[0].0.to_row_off, 400);
        assert_eq!(anchors[0].1, "rId1");
    }

    #[test]
    fn test_column_chart_parse_roundtrip() {
        let original = ChartData {
            chart_type: ChartType::Column,
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
            grouping: Some(ChartGrouping::Clustered),
            scatter_style: None, radar_style: None, hole_size: None,
            bar_dir_horizontal: None,
            style_id: None, plot_area_layout: None,
        };
        let xml = original.to_xml().unwrap();
        let parsed = ChartData::parse(&xml).unwrap();
        assert_eq!(parsed.chart_type, ChartType::Column);
        assert_eq!(parsed.bar_dir_horizontal, Some(false));
    }

    #[test]
    fn test_doughnut_chart_parse_roundtrip() {
        let original = ChartData {
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
        let xml = original.to_xml().unwrap();
        let parsed = ChartData::parse(&xml).unwrap();
        assert_eq!(parsed.chart_type, ChartType::Doughnut);
        assert_eq!(parsed.hole_size, Some(50));
    }

    #[test]
    fn test_line_chart_with_markers_parse_roundtrip() {
        let original = ChartData {
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
        let xml = original.to_xml().unwrap();
        let parsed = ChartData::parse(&xml).unwrap();
        assert_eq!(parsed.chart_type, ChartType::Line);
        assert_eq!(parsed.series[0].line_color.as_deref(), Some("FF0000"));
        assert_eq!(parsed.series[0].line_width, Some(25400));
        assert_eq!(parsed.series[0].marker, Some(MarkerStyle::Circle));
        assert_eq!(parsed.series[0].smooth, Some(true));
    }

    #[test]
    fn test_chart_with_axis_title_parse_roundtrip() {
        let original = ChartData {
            chart_type: ChartType::Line,
            title: Some(ChartTitle { text: "My Chart".into(), overlay: false, font_size: Some(1400), bold: Some(true), color: Some("333333".into()) }),
            series: vec![],
            cat_axis: Some(ChartAxis {
                id: 0, cross_ax: 1,
                title: Some(ChartTitle { text: "Month".into(), overlay: false, font_size: None, bold: None, color: None }),
                num_fmt: None, source_linked: false,
                min: None, max: None, major_unit: None, minor_unit: None,
                log_base: None, reversed: false,
                tick_lbl_pos: None, major_tick_mark: None, minor_tick_mark: None,
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
        let xml = original.to_xml().unwrap();
        let parsed = ChartData::parse(&xml).unwrap();
        // Chart title.
        assert_eq!(parsed.title.as_ref().unwrap().text, "My Chart");
        assert_eq!(parsed.title.as_ref().unwrap().color.as_deref(), Some("333333"));
        // Cat axis title.
        let cat_ax = parsed.cat_axis.as_ref().unwrap();
        assert_eq!(cat_ax.title.as_ref().unwrap().text, "Month");
        assert_eq!(cat_ax.position, Some(AxisPosition::Bottom));
        assert!(cat_ax.major_gridlines);
        // Val axis title + scale.
        let val_ax = parsed.val_axis.as_ref().unwrap();
        assert_eq!(val_ax.title.as_ref().unwrap().text, "Sales ($)");
        assert_eq!(val_ax.title.as_ref().unwrap().color.as_deref(), Some("333333"));
        assert_eq!(val_ax.min, Some(0.0));
        assert_eq!(val_ax.max, Some(1000.0));
        assert_eq!(val_ax.major_unit, Some(200.0));
    }

    #[test]
    fn test_radar_chart_parse_roundtrip() {
        let original = ChartData {
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
        let xml = original.to_xml().unwrap();
        let parsed = ChartData::parse(&xml).unwrap();
        assert_eq!(parsed.chart_type, ChartType::Radar);
        assert_eq!(parsed.radar_style, Some(RadarStyle::Filled));
    }

    #[test]
    fn test_drawing_anchors_multiple_parse_roundtrip() {
        let charts = vec![
            WorksheetChart {
                chart: ChartData {
                    chart_type: ChartType::Bar,
                    title: None, series: vec![], cat_axis: None, val_axis: None,
                    legend: None, data_labels: None, grouping: None,
                    scatter_style: None, radar_style: None, hole_size: None,
                    bar_dir_horizontal: None, style_id: None, plot_area_layout: None,
                },
                anchor: ChartAnchor {
                    from_col: 0, from_row: 0, from_col_off: 0, from_row_off: 0,
                    to_col: 8, to_row: 15, to_col_off: 0, to_row_off: 0,
                },
            },
            WorksheetChart {
                chart: ChartData {
                    chart_type: ChartType::Line,
                    title: None, series: vec![], cat_axis: None, val_axis: None,
                    legend: None, data_labels: None, grouping: None,
                    scatter_style: None, radar_style: None, hole_size: None,
                    bar_dir_horizontal: None, style_id: None, plot_area_layout: None,
                },
                anchor: ChartAnchor {
                    from_col: 10, from_row: 0, from_col_off: 100, from_row_off: 200,
                    to_col: 18, to_row: 20, to_col_off: 300, to_row_off: 400,
                },
            },
        ];
        let r_ids = vec!["rId1".into(), "rId2".into()];
        let drawing_xml = ChartAnchor::generate_drawing_xml(&charts, &r_ids).unwrap();
        let anchors = parse_drawing_anchors(&drawing_xml).unwrap();
        assert_eq!(anchors.len(), 2);
        assert_eq!(anchors[0].1, "rId1");
        assert_eq!(anchors[1].1, "rId2");
        assert_eq!(anchors[1].0.from_col, 10);
        assert_eq!(anchors[1].0.to_row, 20);
    }
}
