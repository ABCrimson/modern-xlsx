//! Chart XML parser — `ChartData::parse` and drawing anchor parser.

use core::hint::cold_path;

use quick_xml::events::{BytesStart, Event};
use quick_xml::Reader;

use super::types::*;
use crate::ooxml::push_entity;
use crate::{ModernXlsxError, Result};

// ---------------------------------------------------------------------------
// XML Parser — chart XML -> ChartData
// ---------------------------------------------------------------------------

/// Helper to extract `val` attribute from a `BytesStart`.
#[inline]
fn attr_val(e: &BytesStart<'_>, key: &[u8]) -> Option<String> {
    e.attributes()
        .flatten()
        .find(|attr| attr.key.as_ref() == key)
        .map(|attr| {
            std::str::from_utf8(&attr.value)
                .unwrap_or_default()
                .to_owned()
        })
}

/// Parse the `val` attribute as `&str`.
#[inline]
fn attr_val_str(e: &BytesStart<'_>) -> String {
    attr_val(e, b"val").unwrap_or_default()
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
    /// Inside `<c:layout>` -> `<c:manualLayout>`.
    ManualLayoutCtx,
    /// Inside `<c:dLbls>` within a series.
    SerDataLabels,
    /// Inside the title text run (`<a:r>` -> `<a:t>`).
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
    /// Inside `<c:trendline>` within a series.
    SerTrendline,
    /// Inside `<c:errBars>` within a series.
    SerErrBars,
    /// Inside `<c:txPr>` within an axis (tick label font properties).
    AxisTxPr(AxisKind),
    /// Inside `<c:view3D>` within `<c:chart>`.
    View3D,
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
#[derive(Default)]
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
    font_size: Option<u32>,
}

impl AxisBuilder {
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
            font_size: self.font_size,
        }
    }
}

/// Builder for accumulating data labels during parsing.
#[derive(Default)]
struct DataLabelsBuilder {
    show_val: bool,
    show_cat_name: bool,
    show_ser_name: bool,
    show_percent: bool,
    num_fmt: Option<String>,
    show_leader_lines: bool,
}

impl DataLabelsBuilder {
    fn build_and_reset(&mut self) -> DataLabels {
        DataLabels {
            show_val: self.show_val,
            show_cat_name: self.show_cat_name,
            show_ser_name: self.show_ser_name,
            show_percent: self.show_percent,
            num_fmt: self.num_fmt.take(),
            show_leader_lines: self.show_leader_lines,
        }
    }
}

/// Builder for accumulating legend during parsing.
#[derive(Default)]
struct LegendBuilder {
    position: Option<LegendPosition>,
    overlay: bool,
}

impl LegendBuilder {
    fn build(&self) -> ChartLegend {
        ChartLegend {
            position: self.position.unwrap_or(LegendPosition::Right),
            overlay: self.overlay,
        }
    }
}

/// Builder for accumulating title text during parsing.
#[derive(Default)]
struct TitleBuilder {
    text: String,
    overlay: bool,
    font_size: Option<u32>,
    bold: Option<bool>,
    color: Option<String>,
}

impl TitleBuilder {
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

/// Builder for accumulating trendline data during parsing.
#[derive(Default)]
struct TrendlineBuilder {
    trend_type: Option<TrendlineType>,
    order: Option<u32>,
    period: Option<u32>,
    forward: Option<f64>,
    backward: Option<f64>,
    display_eq: bool,
    display_r_sqr: bool,
}

impl TrendlineBuilder {
    fn build(self) -> Option<Trendline> {
        let trend_type = self.trend_type?;
        Some(Trendline {
            trend_type,
            order: self.order,
            period: self.period,
            forward: self.forward,
            backward: self.backward,
            display_eq: self.display_eq,
            display_r_sqr: self.display_r_sqr,
        })
    }
}

/// Builder for accumulating error bars data during parsing.
#[derive(Default)]
struct ErrorBarsBuilder {
    err_type: Option<ErrorBarType>,
    direction: Option<ErrorBarDirection>,
    value: Option<f64>,
}

impl ErrorBarsBuilder {
    fn build(self) -> Option<ErrorBars> {
        let err_type = self.err_type?;
        Some(ErrorBars {
            err_type,
            direction: self.direction.unwrap_or_default(),
            value: self.value,
        })
    }
}

/// Builder for View3D during parsing.
#[derive(Default)]
struct View3DBuilder {
    rot_x: Option<i32>,
    rot_y: Option<i32>,
    perspective: Option<u32>,
    r_ang_ax: Option<bool>,
}

impl View3DBuilder {
    fn build(&self) -> Option<View3D> {
        if self.rot_x.is_none()
            && self.rot_y.is_none()
            && self.perspective.is_none()
            && self.r_ang_ax.is_none()
        {
            return None;
        }
        Some(View3D {
            rot_x: self.rot_x,
            rot_y: self.rot_y,
            perspective: self.perspective,
            r_ang_ax: self.r_ang_ax,
        })
    }
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

        // Combo-chart tracking: secondary chart type + series.
        let mut secondary_chart_type: Option<ChartType> = None;
        let mut secondary_series: Vec<ChartSeries> = Vec::new();
        let mut secondary_grouping: Option<ChartGrouping> = None;
        let mut secondary_scatter_style: Option<ScatterStyle> = None;
        let mut secondary_radar_style: Option<RadarStyle> = None;
        let mut secondary_hole_size: Option<u32> = None;
        let mut secondary_bar_dir_horizontal: Option<bool> = None;
        let mut secondary_data_labels: Option<DataLabels> = None;
        let mut is_secondary_chart = false;
        let mut val_axis_count: u32 = 0;
        let mut secondary_val_axis: Option<ChartAxis> = None;

        // Current series being built.
        let mut cur_ser: Option<ChartSeries> = None;
        // Current axis being built.
        let mut cur_cat_axis = AxisBuilder::default();
        let mut cur_val_axis = AxisBuilder::default();
        // Current data labels.
        let mut cur_dlbls = DataLabelsBuilder::default();
        let mut cur_ser_dlbls = DataLabelsBuilder::default();
        // Current legend.
        let mut cur_legend = LegendBuilder::default();
        // Current title.
        let mut cur_title = TitleBuilder::default();
        let mut cur_cat_axis_title = TitleBuilder::default();
        let mut cur_val_axis_title = TitleBuilder::default();
        // Manual layout.
        let mut layout_x: Option<f64> = None;
        let mut layout_y: Option<f64> = None;
        let mut layout_w: Option<f64> = None;
        let mut layout_h: Option<f64> = None;
        // Trendline builder (within series).
        let mut cur_trendline: Option<TrendlineBuilder> = None;
        // Error bars builder (within series).
        let mut cur_err_bars: Option<ErrorBarsBuilder> = None;
        // View3D.
        let mut view_3d: Option<View3D> = None;
        let mut cur_view_3d = View3DBuilder::default();
        // Data table.
        let mut show_data_table = false;

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
                            cur_title = TitleBuilder::default();
                            ctx_stack.push(ParseCtx::Title(TitleOwner::Chart));
                        }
                        (ParseCtx::Chart, b"legend") => {
                            cur_legend = LegendBuilder::default();
                            ctx_stack.push(ParseCtx::Legend);
                        }
                        (ParseCtx::Chart, b"view3D") => {
                            cur_view_3d = View3DBuilder::default();
                            ctx_stack.push(ParseCtx::View3D);
                        }
                        // Chart type elements inside plotArea.
                        // When a second chart-type element appears, it's the
                        // secondary chart in a combo-chart layout.
                        (ParseCtx::PlotArea, b"barChart")
                        | (ParseCtx::PlotArea, b"lineChart")
                        | (ParseCtx::PlotArea, b"pieChart")
                        | (ParseCtx::PlotArea, b"doughnutChart")
                        | (ParseCtx::PlotArea, b"scatterChart")
                        | (ParseCtx::PlotArea, b"areaChart")
                        | (ParseCtx::PlotArea, b"radarChart")
                        | (ParseCtx::PlotArea, b"bubbleChart")
                        | (ParseCtx::PlotArea, b"stockChart") => {
                            let ct = match local {
                                b"barChart" => ChartType::Column, // barDir refines
                                b"lineChart" => ChartType::Line,
                                b"pieChart" => ChartType::Pie,
                                b"doughnutChart" => ChartType::Doughnut,
                                b"scatterChart" => ChartType::Scatter,
                                b"areaChart" => ChartType::Area,
                                b"radarChart" => ChartType::Radar,
                                b"bubbleChart" => ChartType::Bubble,
                                b"stockChart" => ChartType::Stock,
                                _ => {
                                    cold_path();
                                    unreachable!()
                                }
                            };
                            if chart_type.is_none() {
                                // Primary chart.
                                chart_type = Some(ct);
                                is_secondary_chart = false;
                            } else {
                                // Secondary chart (combo).
                                secondary_chart_type = Some(ct);
                                is_secondary_chart = true;
                            }
                            ctx_stack.push(ParseCtx::ChartTypeElem);
                        }
                        (ParseCtx::PlotArea, b"layout") => {
                            ctx_stack.push(ParseCtx::ManualLayoutCtx);
                        }
                        (ParseCtx::PlotArea, b"dTable") => {
                            show_data_table = true;
                            ctx_stack.push(ParseCtx::PlotArea); // track depth
                        }
                        (ParseCtx::ManualLayoutCtx, b"manualLayout") => {
                            // stay in ManualLayoutCtx
                            ctx_stack.push(ParseCtx::ManualLayoutCtx);
                        }
                        // Axes inside plotArea.
                        (ParseCtx::PlotArea, b"catAx") => {
                            cur_cat_axis = AxisBuilder::default();
                            ctx_stack.push(ParseCtx::CatAxis);
                        }
                        (ParseCtx::PlotArea, b"valAx") => {
                            cur_val_axis = AxisBuilder::default();
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
                                trendline: None,
                                error_bars: None,
                            });
                            ctx_stack.push(ParseCtx::Series);
                        }
                        (ParseCtx::ChartTypeElem, b"dLbls") => {
                            cur_dlbls = DataLabelsBuilder::default();
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
                            cur_ser_dlbls = DataLabelsBuilder::default();
                            ctx_stack.push(ParseCtx::SerDataLabels);
                        }
                        (ParseCtx::Series, b"trendline") => {
                            cur_trendline = Some(TrendlineBuilder::default());
                            ctx_stack.push(ParseCtx::SerTrendline);
                        }
                        (ParseCtx::Series, b"errBars") => {
                            cur_err_bars = Some(ErrorBarsBuilder::default());
                            ctx_stack.push(ParseCtx::SerErrBars);
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
                            cur_cat_axis_title = TitleBuilder::default();
                            ctx_stack.push(ParseCtx::Title(TitleOwner::CatAxis));
                        }
                        (ParseCtx::ValAxis, b"title") => {
                            cur_val_axis_title = TitleBuilder::default();
                            ctx_stack.push(ParseCtx::Title(TitleOwner::ValAxis));
                        }
                        // Axis tick label text properties.
                        (ParseCtx::CatAxis, b"txPr") => {
                            ctx_stack.push(ParseCtx::AxisTxPr(AxisKind::Cat));
                        }
                        (ParseCtx::ValAxis, b"txPr") => {
                            ctx_stack.push(ParseCtx::AxisTxPr(AxisKind::Val));
                        }
                        (ParseCtx::AxisTxPr(_), _) => {
                            // Descend into txPr children (a:bodyPr, a:lstStyle, a:p, a:pPr, etc.)
                            ctx_stack.push(ctx);
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
                        // Chart-type element attributes — route to
                        // primary or secondary accumulator.
                        (ParseCtx::ChartTypeElem, b"barDir") => {
                            let val = attr_val_str(e);
                            let (ct, bh) = if val == "bar" {
                                (ChartType::Bar, Some(true))
                            } else {
                                (ChartType::Column, Some(false))
                            };
                            if is_secondary_chart {
                                secondary_chart_type = Some(ct);
                                secondary_bar_dir_horizontal = bh;
                            } else {
                                chart_type = Some(ct);
                                bar_dir_horizontal = bh;
                            }
                        }
                        (ParseCtx::ChartTypeElem, b"grouping") => {
                            let g = ChartGrouping::from_xml(&attr_val_str(e));
                            if is_secondary_chart {
                                secondary_grouping = g;
                            } else {
                                grouping = g;
                            }
                        }
                        (ParseCtx::ChartTypeElem, b"scatterStyle") => {
                            let s = ScatterStyle::from_xml(&attr_val_str(e));
                            if is_secondary_chart {
                                secondary_scatter_style = s;
                            } else {
                                scatter_style = s;
                            }
                        }
                        (ParseCtx::ChartTypeElem, b"radarStyle") => {
                            let r = RadarStyle::from_xml(&attr_val_str(e));
                            if is_secondary_chart {
                                secondary_radar_style = r;
                            } else {
                                radar_style = r;
                            }
                        }
                        (ParseCtx::ChartTypeElem, b"holeSize") => {
                            let h = attr_val_str(e).parse().ok();
                            if is_secondary_chart {
                                secondary_hole_size = h;
                            } else {
                                hole_size = h;
                            }
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
                            // autoZero -> None (default).
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
                            // autoZero -> None.
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
                        // Axis txPr — defRPr with sz attribute.
                        (ParseCtx::AxisTxPr(AxisKind::Cat), b"defRPr") => {
                            if let Some(sz) = attr_val(e, b"sz") {
                                cur_cat_axis.font_size = sz.parse().ok();
                            }
                        }
                        (ParseCtx::AxisTxPr(AxisKind::Val), b"defRPr") => {
                            if let Some(sz) = attr_val(e, b"sz") {
                                cur_val_axis.font_size = sz.parse().ok();
                            }
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
                        // Trendline attributes.
                        (ParseCtx::SerTrendline, b"trendlineType") => {
                            if let Some(ref mut tb) = cur_trendline {
                                tb.trend_type = TrendlineType::from_xml(&attr_val_str(e));
                            }
                        }
                        (ParseCtx::SerTrendline, b"order") => {
                            if let Some(ref mut tb) = cur_trendline {
                                tb.order = attr_val_str(e).parse().ok();
                            }
                        }
                        (ParseCtx::SerTrendline, b"period") => {
                            if let Some(ref mut tb) = cur_trendline {
                                tb.period = attr_val_str(e).parse().ok();
                            }
                        }
                        (ParseCtx::SerTrendline, b"forward") => {
                            if let Some(ref mut tb) = cur_trendline {
                                tb.forward = attr_val_str(e).parse().ok();
                            }
                        }
                        (ParseCtx::SerTrendline, b"backward") => {
                            if let Some(ref mut tb) = cur_trendline {
                                tb.backward = attr_val_str(e).parse().ok();
                            }
                        }
                        (ParseCtx::SerTrendline, b"dispEq") => {
                            if let Some(ref mut tb) = cur_trendline {
                                tb.display_eq = attr_val_str(e) == "1";
                            }
                        }
                        (ParseCtx::SerTrendline, b"dispRSqr") => {
                            if let Some(ref mut tb) = cur_trendline {
                                tb.display_r_sqr = attr_val_str(e) == "1";
                            }
                        }
                        // Error bars attributes.
                        (ParseCtx::SerErrBars, b"errBarType") => {
                            if let Some(ref mut eb) = cur_err_bars {
                                eb.direction = ErrorBarDirection::from_xml(&attr_val_str(e));
                            }
                        }
                        (ParseCtx::SerErrBars, b"errValType") => {
                            if let Some(ref mut eb) = cur_err_bars {
                                eb.err_type = ErrorBarType::from_xml(&attr_val_str(e));
                            }
                        }
                        (ParseCtx::SerErrBars, b"val") => {
                            if let Some(ref mut eb) = cur_err_bars {
                                eb.value = attr_val_str(e).parse().ok();
                            }
                        }
                        // View3D attributes.
                        (ParseCtx::View3D, b"rotX") => {
                            cur_view_3d.rot_x = attr_val_str(e).parse().ok();
                        }
                        (ParseCtx::View3D, b"rotY") => {
                            cur_view_3d.rot_y = attr_val_str(e).parse().ok();
                        }
                        (ParseCtx::View3D, b"perspective") => {
                            cur_view_3d.perspective = attr_val_str(e).parse().ok();
                        }
                        (ParseCtx::View3D, b"rAngAx") => {
                            let val = attr_val_str(e);
                            cur_view_3d.r_ang_ax = Some(val == "1" || val == "true");
                        }
                        // Data table (showKeys within dTable).
                        (ParseCtx::PlotArea, b"dTable") | (_, b"dTable") => {
                            // Self-closing <c:dTable/> or containing <c:showKeys/>.
                            show_data_table = true;
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
                                if is_secondary_chart {
                                    secondary_series.push(ser);
                                } else {
                                    series.push(ser);
                                }
                            }
                            ctx_stack.pop();
                        }
                        (ParseCtx::DataLabels, b"dLbls") => {
                            let dl = Some(cur_dlbls.build_and_reset());
                            if is_secondary_chart {
                                secondary_data_labels = dl;
                            } else {
                                data_labels = dl;
                            }
                            ctx_stack.pop();
                        }
                        (ParseCtx::SerDataLabels, b"dLbls") => {
                            if let Some(ref mut ser) = cur_ser {
                                ser.data_labels = Some(cur_ser_dlbls.build_and_reset());
                            }
                            ctx_stack.pop();
                        }
                        (ParseCtx::SerTrendline, b"trendline") => {
                            if let Some(ref mut ser) = cur_ser
                                && let Some(tb) = cur_trendline.take()
                            {
                                ser.trendline = tb.build();
                            }
                            ctx_stack.pop();
                        }
                        (ParseCtx::SerErrBars, b"errBars") => {
                            if let Some(ref mut ser) = cur_ser
                                && let Some(eb) = cur_err_bars.take()
                            {
                                ser.error_bars = eb.build();
                            }
                            ctx_stack.pop();
                        }
                        (ParseCtx::View3D, b"view3D") => {
                            view_3d = cur_view_3d.build();
                            ctx_stack.pop();
                        }
                        (ParseCtx::AxisTxPr(_), b"txPr") => {
                            ctx_stack.pop();
                        }
                        (ParseCtx::AxisTxPr(_), _) => {
                            // Pop child contexts within txPr (a:p, a:pPr, etc.)
                            ctx_stack.pop();
                        }
                        (ParseCtx::CatAxis, b"catAx") => {
                            let builder = std::mem::take(&mut cur_cat_axis);
                            cat_axis = Some(builder.build_with_title(
                                cur_cat_axis_title.build_option(),
                            ));
                            ctx_stack.pop();
                        }
                        (ParseCtx::ValAxis, b"valAx") => {
                            let builder = std::mem::take(&mut cur_val_axis);
                            let axis = builder.build_with_title(
                                cur_val_axis_title.build_option(),
                            );
                            if val_axis_count == 0 {
                                val_axis = Some(axis);
                            } else {
                                secondary_val_axis = Some(axis);
                            }
                            val_axis_count += 1;
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
                Ok(Event::Text(ref e)) if capturing_text => {
                    text_buf.push_str(
                        std::str::from_utf8(e.as_ref()).unwrap_or_default(),
                    );
                }
                Ok(Event::GeneralRef(ref e)) if capturing_text => {
                    push_entity(&mut text_buf, e.as_ref());
                }
                Ok(Event::Eof) => break,
                Err(err) => {
                    cold_path();
                    return Err(ModernXlsxError::XmlParse(format!(
                        "chart parse error: {err}"
                    )));
                }
                _ => {}
            }
            buf.clear();
        }

        // Build secondary chart if a combo chart was detected.
        let secondary_chart = secondary_chart_type.map(|ct| {
            Box::new(ChartData {
                chart_type: ct,
                title: None,
                series: secondary_series,
                cat_axis: None,
                val_axis: None,
                legend: None,
                data_labels: secondary_data_labels,
                grouping: secondary_grouping,
                scatter_style: secondary_scatter_style,
                radar_style: secondary_radar_style,
                hole_size: secondary_hole_size,
                bar_dir_horizontal: secondary_bar_dir_horizontal,
                style_id: None,
                plot_area_layout: None,
                secondary_chart: None,
                secondary_val_axis: None,
                show_data_table: false,
                view_3d: None,
            })
        });

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
            secondary_chart,
            secondary_val_axis,
            show_data_table,
            view_3d,
        })
    }
}

// ---------------------------------------------------------------------------
// Drawing XML Parser — parse drawing anchors for chart references
// ---------------------------------------------------------------------------

/// Parse `xl/drawings/drawing{n}.xml` to extract chart anchors and chart rIds.
///
/// Returns a vec of `(ChartAnchor, rId)` pairs for each `<xdr:twoCellAnchor>`
/// or `<xdr:oneCellAnchor>` that contains a chart reference.
pub fn parse_drawing_anchors(data: &[u8]) -> Result<Vec<(ChartAnchor, String)>> {
    let mut reader = Reader::from_reader(data);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::with_capacity(512);

    let mut result: Vec<(ChartAnchor, String)> = Vec::new();

    // State tracking.
    let mut in_anchor = false;
    let mut is_one_cell_anchor = false;
    let mut in_from = false;
    let mut in_to = false;
    let mut in_graphic_frame = false;
    let mut from_col: u32 = 0;
    let mut from_col_off: u64 = 0;
    let mut from_row: u32 = 0;
    let mut from_row_off: u64 = 0;
    let mut to_col: u32 = 0;
    let mut to_col_off: u64 = 0;
    let mut to_row: u32 = 0;
    let mut to_row_off: u64 = 0;
    let mut ext_cx: Option<u64> = None;
    let mut ext_cy: Option<u64> = None;
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

    /// Reset all anchor state for a new anchor element.
    macro_rules! reset_anchor_state {
        () => {
            in_anchor = true;
            from_col = 0;
            from_col_off = 0;
            from_row = 0;
            from_row_off = 0;
            to_col = 0;
            to_col_off = 0;
            to_row = 0;
            to_row_off = 0;
            ext_cx = None;
            ext_cy = None;
            chart_r_id = None;
        };
    }

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) => {
                let local = e.local_name();
                let local = local.as_ref();
                match local {
                    b"twoCellAnchor" => {
                        reset_anchor_state!();
                        is_one_cell_anchor = false;
                    }
                    b"oneCellAnchor" => {
                        reset_anchor_state!();
                        is_one_cell_anchor = true;
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
                    b"graphicFrame" if in_anchor => {
                        in_graphic_frame = true;
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
                    if let Some(attr) = e.attributes().flatten().find(|a| {
                        let key = a.key.as_ref();
                        key == b"r:id" || key.ends_with(b":id")
                    }) {
                        chart_r_id = Some(
                            std::str::from_utf8(&attr.value)
                                .unwrap_or_default()
                                .to_owned(),
                        );
                    }
                }
                // <xdr:ext cx="..." cy="..."/> — oneCellAnchor dimensions.
                // Only match the direct-child ext, not <a:ext> inside graphicFrame.
                if local == b"ext" && in_anchor && is_one_cell_anchor && !in_from && !in_to && !in_graphic_frame {
                    for attr in e.attributes().flatten() {
                        let key = attr.key.as_ref();
                        let val_str = std::str::from_utf8(&attr.value).unwrap_or_default();
                        if key == b"cx" {
                            ext_cx = Some(val_str.parse().unwrap_or(0));
                        } else if key == b"cy" {
                            ext_cy = Some(val_str.parse().unwrap_or(0));
                        }
                    }
                }
            }
            Ok(Event::End(ref e)) => {
                let local = e.local_name();
                let local = local.as_ref();
                match local {
                    b"twoCellAnchor" | b"oneCellAnchor" => {
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
                                    ext_cx: if is_one_cell_anchor { ext_cx } else { None },
                                    ext_cy: if is_one_cell_anchor { ext_cy } else { None },
                                },
                                r_id,
                            ));
                        }
                        in_anchor = false;
                        is_one_cell_anchor = false;
                        in_from = false;
                        in_to = false;
                    }
                    b"from" => {
                        in_from = false;
                    }
                    b"to" => {
                        in_to = false;
                    }
                    b"graphicFrame" => {
                        in_graphic_frame = false;
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
            Ok(Event::Text(ref e)) if text_target != DrawingTextTarget::None => {
                text_buf.push_str(
                    std::str::from_utf8(e.as_ref()).unwrap_or_default(),
                );
            }
            Ok(Event::Eof) => break,
            Err(err) => {
                cold_path();
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
