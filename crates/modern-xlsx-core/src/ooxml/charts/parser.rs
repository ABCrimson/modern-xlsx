//! Chart XML parser — [`ChartData::parse`] for `xl/charts/chart{n}.xml` and
//! [`parse_drawing_anchors`] for `xl/drawings/drawing{n}.xml`.

use core::hint::cold_path;
use std::borrow::Cow;
use std::str::FromStr;

use log::warn;
use quick_xml::Reader;
use quick_xml::events::attributes::Attribute;
use quick_xml::events::{BytesStart, Event};

use super::types::{
    AxisPosition, ChartAnchor, ChartAxis, ChartData, ChartGrouping, ChartLegend, ChartSeries,
    ChartTitle, ChartType, CrossBetween, DataLabels, ErrorBarAxis, ErrorBarDirection, ErrorBarType,
    ErrorBars, LegendPosition, ManualLayout, MarkerStyle, RadarStyle, ScatterStyle,
    TickLabelPosition, TickMark, Trendline, TrendlineType, View3D,
};
use crate::ooxml::shared_strings::ExpansionBudget;
use crate::ooxml::xml::{
    XML_S, attr_str, attrs, ct_boolean, parse_attr, push_general_ref, xsd_bool,
};
use crate::{ModernXlsxError, Result};

// ---------------------------------------------------------------------------
// XML Parser — chart XML -> ChartData
// ---------------------------------------------------------------------------

/// `CT_PlotArea` chart types the model does not represent. A chart holding one is
/// kept verbatim in [`ChartData::raw_xml`].
const UNMODELLED_CHART_TYPES: [&str; 7] = [
    "bar3DChart",
    "line3DChart",
    "pie3DChart",
    "area3DChart",
    "surfaceChart",
    "surface3DChart",
    "ofPieChart",
];

/// `XmlParse` for a reader error inside a chart part.
#[cold]
fn chart_xml_err(err: impl core::fmt::Display) -> ModernXlsxError {
    ModernXlsxError::XmlParse(format!("chart parse error: {err}"))
}

/// Attribute `key` of `e`, or `None` when absent.
///
/// # Errors
///
/// `XmlParse` when the attribute list is malformed.
fn find_attr<'a>(e: &'a BytesStart<'_>, key: &str) -> Result<Option<Attribute<'a>>> {
    for attr in attrs(e) {
        let attr = attr?;
        if attr.key.as_ref() == key {
            return Ok(Some(attr));
        }
    }
    Ok(None)
}

/// The decoded value (entities and character references resolved) of attribute
/// `key` of `e`, or `None` when absent.
///
/// # Errors
///
/// `XmlParse` for a malformed attribute list or a bad entity in the value.
fn attr_val(e: &BytesStart<'_>, key: &str) -> Result<Option<String>> {
    find_attr(e, key)?
        .map(|a| attr_str(&a).map(Cow::into_owned))
        .transpose()
}

/// The decoded `val` attribute of `e`, or `""` when absent.
///
/// # Errors
///
/// See [`attr_val`].
fn attr_val_str(e: &BytesStart<'_>) -> Result<String> {
    Ok(attr_val(e, "val")?.unwrap_or_default())
}

/// Numeric attribute `key` of `e`, or `None` when absent.
///
/// # Errors
///
/// `XmlParse` naming the element when the value is present but not a `T`.
fn num_attr<T: FromStr>(e: &BytesStart<'_>, key: &str) -> Result<Option<T>> {
    find_attr(e, key)?
        .map(|a| parse_attr(e.name().as_ref(), &a))
        .transpose()
}

/// Numeric `val` attribute of `e`, or `None` when absent.
///
/// # Errors
///
/// See [`num_attr`].
fn num_val<T: FromStr>(e: &BytesStart<'_>) -> Result<Option<T>> {
    num_attr(e, "val")
}

/// `c:holeSize@val`: `ST_HoleSize` is a percentage 1-90, written with or without a
/// trailing `%`.
///
/// # Errors
///
/// `XmlParse` when the value is present but not a whole percentage.
fn hole_size_val(e: &BytesStart<'_>) -> Result<Option<u32>> {
    let Some(raw) = attr_val(e, "val")? else {
        return Ok(None);
    };
    let trimmed = raw.trim_matches(XML_S);
    let digits = trimmed.strip_suffix('%').unwrap_or(trimmed);
    digits.parse().map(Some).map_err(|_| {
        cold_path();
        ModernXlsxError::XmlParse(format!("{} @val: invalid value '{raw}'", e.name().as_ref()))
    })
}

/// Parse the text content `raw` of element `elem` as `T`, after collapsing the
/// surrounding XML white space.
///
/// # Errors
///
/// `XmlParse("<elem>: invalid value '<raw>'")` when it does not parse.
fn parse_text<T: FromStr>(elem: &str, raw: &str) -> Result<T> {
    raw.trim_matches(XML_S).parse().map_err(|_| {
        cold_path();
        ModernXlsxError::XmlParse(format!("{elem}: invalid value '{raw}'"))
    })
}

/// Remove and return the axis whose id comes first in `ids`, in `ids` order.
fn take_axis(axes: &mut Vec<ChartAxis>, ids: &[u32]) -> Option<ChartAxis> {
    ids.iter()
        .find_map(|id| axes.iter().position(|a| a.id == *id))
        .map(|at| axes.remove(at))
}

/// Remove and return the first axis left in `axes`, in document order.
fn take_first(axes: &mut Vec<ChartAxis>) -> Option<ChartAxis> {
    (!axes.is_empty()).then(|| axes.remove(0))
}

/// Parser state machine context.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum ParseCtx {
    /// Top-level or unknown context.
    Root,
    /// Inside `<c:chart>`.
    Chart,
    /// Inside a chart-type element (`<c:barChart>`, etc.).
    ChartTypeElem,
    /// Inside `<c:ser>`.
    Series,
    /// Inside `<c:tx>` within a series: a direct `<c:v>` child is the literal name.
    SerTx,
    /// Inside `<c:strRef>` within a series `<c:tx>`: its `<c:f>` is the name formula,
    /// and the values of its `<c:strCache>` are never read as the name.
    SerTxStrRef,
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
    /// Inside the chart-level `<c:dLbls>` of a chart-type element.
    DataLabels,
    /// Inside a per-point `<c:dLbl>` of a `<c:dLbls>`; its children are not modelled
    /// and never reach the enclosing data labels.
    DLbl,
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
    /// Inside `<c:trendline>` within a series.
    SerTrendline,
    /// Inside `<c:errBars>` within a series.
    SerErrBars,
    /// Inside `<c:plus>` or `<c:minus>` within `<c:errBars>`.
    ErrBarsRef(ErrRefKind),
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

impl TitleOwner {
    /// Index of this owner's builder in the parser's title array.
    #[inline]
    fn index(self) -> usize {
        match self {
            Self::Chart => 0,
            Self::CatAxis => 1,
            Self::ValAxis => 2,
        }
    }
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

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum ErrRefKind {
    Plus,
    Minus,
}

/// What text we are currently capturing.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum TextTarget {
    None,
    SerName,
    SerNameLiteral,
    CatRef,
    ValRef,
    XValRef,
    BubbleRef,
    ErrRef(ErrRefKind),
    Title(TitleOwner),
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
    cross_between: Option<CrossBetween>,
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
            cross_between: self.cross_between,
        }
    }

    /// Read one empty child element `local` of `<c:catAx>` / `<c:valAx>`; children
    /// the model does not hold are ignored.
    fn read_child(&mut self, local: &str, e: &BytesStart<'_>) -> Result<()> {
        match local {
            "axId" => self.id = num_val(e)?,
            "delete" => self.delete = ct_boolean(e)?,
            "axPos" => self.position = AxisPosition::from_xml(&attr_val_str(e)?),
            "majorGridlines" => self.major_gridlines = true,
            "minorGridlines" => self.minor_gridlines = true,
            "numFmt" => {
                self.num_fmt = attr_val(e, "formatCode")?;
                if let Some(sl) = attr_val(e, "sourceLinked")? {
                    self.source_linked = xsd_bool(&sl).unwrap_or(false);
                }
            }
            "majorTickMark" => self.major_tick_mark = TickMark::from_xml(&attr_val_str(e)?),
            "minorTickMark" => self.minor_tick_mark = TickMark::from_xml(&attr_val_str(e)?),
            "tickLblPos" => self.tick_lbl_pos = TickLabelPosition::from_xml(&attr_val_str(e)?),
            "crossAx" => self.cross_ax = num_val(e)?,
            "crossesAt" => self.crosses_at = num_val(e)?,
            "majorUnit" => self.major_unit = num_val(e)?,
            "minorUnit" => self.minor_unit = num_val(e)?,
            // `c:crosses` (autoZero, min, max) is not modelled.
            _ => {}
        }
        Ok(())
    }

    /// Read one child element `local` of the axis `<c:scaling>`.
    fn read_scaling_child(&mut self, local: &str, e: &BytesStart<'_>) -> Result<()> {
        match local {
            "orientation" => self.reversed = attr_val_str(e)? == "maxMin",
            "min" => self.min = num_val(e)?,
            "max" => self.max = num_val(e)?,
            "logBase" => self.log_base = num_val(e)?,
            _ => {}
        }
        Ok(())
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
    show_legend_key: Option<bool>,
    show_bubble_size: Option<bool>,
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
            show_legend_key: self.show_legend_key,
            show_bubble_size: self.show_bubble_size,
        }
    }

    /// Read one empty child element `local` of `<c:dLbls>`; children the model does
    /// not hold are ignored.
    fn read_child(&mut self, local: &str, e: &BytesStart<'_>) -> Result<()> {
        match local {
            "showVal" => self.show_val = ct_boolean(e)?,
            "showCatName" => self.show_cat_name = ct_boolean(e)?,
            "showSerName" => self.show_ser_name = ct_boolean(e)?,
            "showPercent" => self.show_percent = ct_boolean(e)?,
            "showLeaderLines" => self.show_leader_lines = ct_boolean(e)?,
            "showLegendKey" => self.show_legend_key = Some(ct_boolean(e)?),
            "showBubbleSize" => self.show_bubble_size = Some(ct_boolean(e)?),
            "numFmt" => self.num_fmt = attr_val(e, "formatCode")?,
            _ => {}
        }
        Ok(())
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

    /// Read `sz` and `b` from a title's `<a:defRPr>` or `<a:rPr>`.
    fn read_run_props(&mut self, e: &BytesStart<'_>) -> Result<()> {
        if let Some(sz) = num_attr(e, "sz")? {
            self.font_size = Some(sz);
        }
        if let Some(b) = attr_val(e, "b")? {
            self.bold = xsd_bool(&b);
        }
        Ok(())
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
    err_dir: Option<ErrorBarAxis>,
    err_type: Option<ErrorBarType>,
    direction: Option<ErrorBarDirection>,
    value: Option<f64>,
    plus_ref: Option<String>,
    minus_ref: Option<String>,
}

impl ErrorBarsBuilder {
    fn build(self) -> Option<ErrorBars> {
        let err_type = self.err_type?;
        Some(ErrorBars {
            err_dir: self.err_dir,
            err_type,
            direction: self.direction.unwrap_or_default(),
            value: self.value,
            plus_ref: self.plus_ref,
            minus_ref: self.minus_ref,
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

/// The capacity of each string in `strings` that is set.
fn capacity_of<'a>(strings: impl IntoIterator<Item = Option<&'a String>>) -> usize {
    strings.into_iter().flatten().map(String::capacity).sum()
}

/// The memory a kept series takes: its [`ChartSeries`] plus the capacity of
/// its strings (its references, name, colours, data-label number format and
/// error-bar references).
fn series_memory(ser: &ChartSeries) -> usize {
    size_of::<ChartSeries>()
        + ser.val_ref.capacity()
        + capacity_of([
            ser.name.as_ref(),
            ser.name_literal.as_ref(),
            ser.cat_ref.as_ref(),
            ser.x_val_ref.as_ref(),
            ser.bubble_size_ref.as_ref(),
            ser.fill_color.as_ref(),
            ser.line_color.as_ref(),
            ser.data_labels.as_ref().and_then(|d| d.num_fmt.as_ref()),
            ser.error_bars.as_ref().and_then(|e| e.plus_ref.as_ref()),
            ser.error_bars.as_ref().and_then(|e| e.minus_ref.as_ref()),
        ])
}

/// The memory a kept axis takes: its [`ChartAxis`] plus the capacity of its
/// strings (its number format and its title's text and colour).
fn axis_memory(axis: &ChartAxis) -> usize {
    size_of::<ChartAxis>()
        + capacity_of([
            axis.num_fmt.as_ref(),
            axis.title.as_ref().map(|t| &t.text),
            axis.title.as_ref().and_then(|t| t.color.as_ref()),
        ])
}

impl ChartData {
    /// Parse a chart XML (`xl/charts/chart{n}.xml`) into a `ChartData`.
    ///
    /// Axes are matched to the chart-type elements by their `c:axId`s. A part the
    /// model cannot represent (see [`ChartData::raw_xml`]) still parses: the
    /// representable parts are modelled, `raw_xml` holds the input, and a warning
    /// is logged.
    ///
    /// # Errors
    ///
    /// `XmlParse` for malformed XML, a malformed attribute list, an undecodable
    /// entity or a numeric attribute that does not parse.
    pub fn parse(data: &[u8]) -> Result<Self> {
        Self::parse_within(data, &mut ExpansionBudget::new(usize::MAX))
    }

    /// [`Self::parse`], counting in `budget` (with [`ExpansionBudget::store`])
    /// each series and axis the part keeps before it is kept: a series at the
    /// size of its [`ChartSeries`] plus the capacity of its strings, an axis
    /// at the size of its [`ChartAxis`] plus the capacity of its strings.
    ///
    /// # Errors
    ///
    /// As [`Self::parse`], and `SizeLimit` once the count passes
    /// `budget.store_limit`.
    pub(crate) fn parse_within(data: &[u8], budget: &mut ExpansionBudget) -> Result<Self> {
        let mut reader = Reader::from_reader(data);
        // Text is read untrimmed so entity-split title and name text keeps its
        // spaces; white space between elements reaches no capture.
        reader.config_mut().trim_text(false);
        let mut buf = Vec::with_capacity(512);
        let mut skip_buf = Vec::new();

        // Result fields.
        let mut chart_type: Option<ChartType> = None;
        let mut title: Option<ChartTitle> = None;
        let mut series: Vec<ChartSeries> = Vec::new();
        let mut legend: Option<ChartLegend> = None;
        let mut data_labels: Option<DataLabels> = None;
        let mut grouping: Option<ChartGrouping> = None;
        let mut scatter_style: Option<ScatterStyle> = None;
        let mut radar_style: Option<RadarStyle> = None;
        let mut hole_size: Option<u32> = None;
        let mut bar_dir_horizontal: Option<bool> = None;
        let mut style_id: Option<u32> = None;
        let mut plot_area_layout: Option<ManualLayout> = None;
        let mut auto_title_deleted: Option<bool> = None;

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
        // Modelled chart-type elements seen so far (at most 2).
        let mut chart_type_elems = 0u8;

        // Axes in document order, and the ids each chart-type element names.
        let mut cat_axes: Vec<ChartAxis> = Vec::new();
        let mut val_axes: Vec<ChartAxis> = Vec::new();
        let mut primary_ax_ids: Vec<u32> = Vec::new();
        let mut secondary_ax_ids: Vec<u32> = Vec::new();

        // What the model cannot represent; non-empty keeps the part verbatim.
        let mut unmodelled: Vec<String> = Vec::new();

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
        // Current titles, indexed by `TitleOwner::index`: chart, catAx, valAx.
        let mut titles: [TitleBuilder; 3] = Default::default();
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
                        // Markup compatibility (ECMA-376 Part 3): an `mc:Choice` needs
                        // a namespace this reader does not implement (c14 and later),
                        // so skip it and read the `mc:Fallback` sibling instead.
                        (_, "Choice") => {
                            reader
                                .read_to_end_into(e.name(), &mut skip_buf)
                                .map_err(chart_xml_err)?;
                        }
                        (ParseCtx::Root, "chart") => {
                            ctx_stack.push(ParseCtx::Chart);
                        }
                        (ParseCtx::Chart, "plotArea") => {
                            ctx_stack.push(ParseCtx::PlotArea);
                        }
                        (ParseCtx::Chart, "title") => {
                            titles[TitleOwner::Chart.index()] = TitleBuilder::default();
                            ctx_stack.push(ParseCtx::Title(TitleOwner::Chart));
                        }
                        (ParseCtx::Chart, "legend") => {
                            cur_legend = LegendBuilder::default();
                            ctx_stack.push(ParseCtx::Legend);
                        }
                        (ParseCtx::Chart, "view3D") => {
                            cur_view_3d = View3DBuilder::default();
                            ctx_stack.push(ParseCtx::View3D);
                        }
                        // A chartSpace-level reference to another part (text boxes and
                        // shapes over the chart, or the embedded data workbook): its
                        // `r:id` has no slot in the model, so keep the part verbatim.
                        (ParseCtx::Root, "userShapes" | "externalData") => {
                            unmodelled.push(format!("a {local}"));
                            reader
                                .read_to_end_into(e.name(), &mut skip_buf)
                                .map_err(chart_xml_err)?;
                        }
                        // A chart type the model does not represent: skip it, never
                        // touching the modelled charts, and keep the part verbatim.
                        (ParseCtx::PlotArea, tag) if UNMODELLED_CHART_TYPES.contains(&tag) => {
                            unmodelled.push(format!("a {tag}"));
                            reader
                                .read_to_end_into(e.name(), &mut skip_buf)
                                .map_err(chart_xml_err)?;
                        }
                        // A date or series axis has no slot in the model: skip it and
                        // keep the part verbatim, so the chart keeps its two axIds.
                        (ParseCtx::PlotArea, "dateAx" | "serAx") => {
                            unmodelled.push(format!("a {local}"));
                            reader
                                .read_to_end_into(e.name(), &mut skip_buf)
                                .map_err(chart_xml_err)?;
                        }
                        // Chart type elements inside plotArea. The second one is the
                        // secondary chart of a combo; a third is not modelled.
                        (
                            ParseCtx::PlotArea,
                            "barChart" | "lineChart" | "pieChart" | "doughnutChart"
                            | "scatterChart" | "areaChart" | "radarChart" | "bubbleChart"
                            | "stockChart",
                        ) => {
                            if chart_type_elems == 2 {
                                unmodelled.push(format!("a third chart type ({local})"));
                                reader
                                    .read_to_end_into(e.name(), &mut skip_buf)
                                    .map_err(chart_xml_err)?;
                            } else {
                                let ct = match local {
                                    "barChart" => ChartType::Column, // barDir refines
                                    "lineChart" => ChartType::Line,
                                    "pieChart" => ChartType::Pie,
                                    "doughnutChart" => ChartType::Doughnut,
                                    "scatterChart" => ChartType::Scatter,
                                    "areaChart" => ChartType::Area,
                                    "radarChart" => ChartType::Radar,
                                    "bubbleChart" => ChartType::Bubble,
                                    "stockChart" => ChartType::Stock,
                                    _ => {
                                        cold_path();
                                        unreachable!()
                                    }
                                };
                                chart_type_elems += 1;
                                is_secondary_chart = chart_type_elems == 2;
                                if is_secondary_chart {
                                    secondary_chart_type = Some(ct);
                                } else {
                                    chart_type = Some(ct);
                                }
                                ctx_stack.push(ParseCtx::ChartTypeElem);
                            }
                        }
                        (ParseCtx::PlotArea, "layout") => {
                            ctx_stack.push(ParseCtx::ManualLayoutCtx);
                        }
                        (ParseCtx::PlotArea, "dTable") => {
                            show_data_table = true;
                            ctx_stack.push(ParseCtx::PlotArea); // track depth
                        }
                        (ParseCtx::ManualLayoutCtx, "manualLayout") => {
                            // stay in ManualLayoutCtx
                            ctx_stack.push(ParseCtx::ManualLayoutCtx);
                        }
                        // Axes inside plotArea.
                        (ParseCtx::PlotArea, "catAx") => {
                            cur_cat_axis = AxisBuilder::default();
                            titles[TitleOwner::CatAxis.index()] = TitleBuilder::default();
                            ctx_stack.push(ParseCtx::CatAxis);
                        }
                        (ParseCtx::PlotArea, "valAx") => {
                            cur_val_axis = AxisBuilder::default();
                            titles[TitleOwner::ValAxis.index()] = TitleBuilder::default();
                            ctx_stack.push(ParseCtx::ValAxis);
                        }
                        // Series inside chart type element.
                        (ParseCtx::ChartTypeElem, "ser") => {
                            cur_ser = Some(ChartSeries {
                                idx: 0,
                                order: 0,
                                name: None,
                                name_literal: None,
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
                        (ParseCtx::ChartTypeElem, "dLbls") => {
                            cur_dlbls = DataLabelsBuilder::default();
                            ctx_stack.push(ParseCtx::DataLabels);
                        }
                        // A per-point label inside either dLbls.
                        (ParseCtx::DataLabels | ParseCtx::SerDataLabels, "dLbl") => {
                            ctx_stack.push(ParseCtx::DLbl);
                        }
                        // Series children.
                        (ParseCtx::Series, "tx") => {
                            ctx_stack.push(ParseCtx::SerTx);
                        }
                        (ParseCtx::Series, "cat") => {
                            ctx_stack.push(ParseCtx::SerRef(SerRefKind::Cat));
                        }
                        (ParseCtx::Series, "val") => {
                            ctx_stack.push(ParseCtx::SerRef(SerRefKind::Val));
                        }
                        (ParseCtx::Series, "xVal") => {
                            ctx_stack.push(ParseCtx::SerRef(SerRefKind::XVal));
                        }
                        (ParseCtx::Series, "yVal") => {
                            ctx_stack.push(ParseCtx::SerRef(SerRefKind::YVal));
                        }
                        (ParseCtx::Series, "bubbleSize") => {
                            ctx_stack.push(ParseCtx::SerRef(SerRefKind::BubbleSize));
                        }
                        (ParseCtx::Series, "spPr") => {
                            ctx_stack.push(ParseCtx::SerSpPr);
                        }
                        (ParseCtx::Series, "marker") => {
                            ctx_stack.push(ParseCtx::SerMarker);
                        }
                        // A per-point override (CT_DPt: idx, explosion, spPr,
                        // marker, ...) is not modelled; skip it so its children
                        // never overwrite the series' own idx, fill or marker.
                        (ParseCtx::Series, "dPt") => {
                            reader
                                .read_to_end_into(e.name(), &mut skip_buf)
                                .map_err(chart_xml_err)?;
                        }
                        (ParseCtx::Series, "dLbls") => {
                            cur_ser_dlbls = DataLabelsBuilder::default();
                            ctx_stack.push(ParseCtx::SerDataLabels);
                        }
                        (ParseCtx::Series, "trendline") => {
                            cur_trendline = Some(TrendlineBuilder::default());
                            ctx_stack.push(ParseCtx::SerTrendline);
                        }
                        (ParseCtx::Series, "errBars") => {
                            cur_err_bars = Some(ErrorBarsBuilder::default());
                            ctx_stack.push(ParseCtx::SerErrBars);
                        }
                        // Custom error amounts.
                        (ParseCtx::SerErrBars, "plus") => {
                            ctx_stack.push(ParseCtx::ErrBarsRef(ErrRefKind::Plus));
                        }
                        (ParseCtx::SerErrBars, "minus") => {
                            ctx_stack.push(ParseCtx::ErrBarsRef(ErrRefKind::Minus));
                        }
                        (ParseCtx::ErrBarsRef(kind), "f") => {
                            capturing_text = true;
                            text_target = TextTarget::ErrRef(kind);
                            text_buf.clear();
                            ctx_stack.push(ctx);
                        }
                        // spPr children.
                        (ParseCtx::SerSpPr, "solidFill") => {
                            ctx_stack.push(ParseCtx::SerSpPrFill);
                        }
                        (ParseCtx::SerSpPr, "ln") => {
                            // Parse line width from `w` attribute.
                            if let Some(ref mut ser) = cur_ser
                                && let Some(w) = num_attr(e, "w")?
                            {
                                ser.line_width = Some(w);
                            }
                            ctx_stack.push(ParseCtx::SerSpPrLn);
                        }
                        (ParseCtx::SerSpPrLn, "solidFill") => {
                            ctx_stack.push(ParseCtx::SerSpPrLnFill);
                        }
                        // An `<a:srgbClr val>` with colour modifiers (`a:alpha`,
                        // `a:lumMod`, ...): the colour is read, the modifiers are
                        // not modelled, and the generic End pops the context.
                        (ParseCtx::SerSpPrFill, "srgbClr") => {
                            if let Some(ref mut ser) = cur_ser {
                                ser.fill_color = attr_val(e, "val")?;
                            }
                            ctx_stack.push(ctx);
                        }
                        (ParseCtx::SerSpPrLnFill, "srgbClr") => {
                            if let Some(ref mut ser) = cur_ser {
                                ser.line_color = attr_val(e, "val")?;
                            }
                            ctx_stack.push(ctx);
                        }
                        // Series name: the formula of `c:strRef`, or a literal `c:v`.
                        (ParseCtx::SerTx, "strRef") => {
                            ctx_stack.push(ParseCtx::SerTxStrRef);
                        }
                        (ParseCtx::SerTxStrRef, "f") => {
                            capturing_text = true;
                            text_target = TextTarget::SerName;
                            text_buf.clear();
                            ctx_stack.push(ctx);
                        }
                        (ParseCtx::SerTx, "v") => {
                            capturing_text = true;
                            text_target = TextTarget::SerNameLiteral;
                            text_buf.clear();
                            ctx_stack.push(ctx);
                        }
                        // Ref formula capture.
                        (ParseCtx::SerRef(kind), "f") => {
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
                        (ParseCtx::CatAxis, "scaling") => {
                            ctx_stack.push(ParseCtx::AxisScaling(AxisKind::Cat));
                        }
                        (ParseCtx::ValAxis, "scaling") => {
                            ctx_stack.push(ParseCtx::AxisScaling(AxisKind::Val));
                        }
                        (ParseCtx::CatAxis, "title") => {
                            ctx_stack.push(ParseCtx::Title(TitleOwner::CatAxis));
                        }
                        (ParseCtx::ValAxis, "title") => {
                            ctx_stack.push(ParseCtx::Title(TitleOwner::ValAxis));
                        }
                        // Gridlines with children (Excel writes `<c:spPr>` inside):
                        // set the flag and track depth for the generic End.
                        (ParseCtx::CatAxis, "majorGridlines" | "minorGridlines") => {
                            cur_cat_axis.read_child(local, e)?;
                            ctx_stack.push(ctx);
                        }
                        (ParseCtx::ValAxis, "majorGridlines" | "minorGridlines") => {
                            cur_val_axis.read_child(local, e)?;
                            ctx_stack.push(ctx);
                        }
                        // Axis tick label text properties.
                        (ParseCtx::CatAxis, "txPr") => {
                            ctx_stack.push(ParseCtx::AxisTxPr(AxisKind::Cat));
                        }
                        (ParseCtx::ValAxis, "txPr") => {
                            ctx_stack.push(ParseCtx::AxisTxPr(AxisKind::Val));
                        }
                        // `<a:defRPr sz>` with children (fill, latin font).
                        (ParseCtx::AxisTxPr(kind), "defRPr") => {
                            let axis = match kind {
                                AxisKind::Cat => &mut cur_cat_axis,
                                AxisKind::Val => &mut cur_val_axis,
                            };
                            if let Some(sz) = num_attr(e, "sz")? {
                                axis.font_size = Some(sz);
                            }
                            ctx_stack.push(ctx);
                        }
                        (ParseCtx::AxisTxPr(_), _) => {
                            // Descend into txPr children (a:bodyPr, a:lstStyle, a:p, a:pPr, etc.)
                            ctx_stack.push(ctx);
                        }
                        // Title text structure.
                        (ParseCtx::Title(owner), "r") => {
                            ctx_stack.push(ParseCtx::TitleRun(owner));
                        }
                        (ParseCtx::Title(owner), "pPr") => {
                            ctx_stack.push(ParseCtx::TitlePPr(owner));
                        }
                        (ParseCtx::TitlePPr(owner), "defRPr") => {
                            titles[owner.index()].read_run_props(e)?;
                            ctx_stack.push(ParseCtx::TitleDefRPr(owner));
                        }
                        (ParseCtx::TitleDefRPr(owner), "solidFill") => {
                            ctx_stack.push(ParseCtx::TitleDefRPrFill(owner));
                        }
                        (ParseCtx::TitleRun(owner), "rPr") => {
                            titles[owner.index()].read_run_props(e)?;
                            ctx_stack.push(ParseCtx::TitleRunPr(owner));
                        }
                        (ParseCtx::TitleRunPr(owner), "solidFill") => {
                            ctx_stack.push(ParseCtx::TitleRunPrFill(owner));
                        }
                        // A title colour with modifiers, as for the series above.
                        (
                            ParseCtx::TitleDefRPrFill(owner) | ParseCtx::TitleRunPrFill(owner),
                            "srgbClr",
                        ) => {
                            titles[owner.index()].color = attr_val(e, "val")?;
                            ctx_stack.push(ctx);
                        }
                        (ParseCtx::TitleRun(owner), "t") => {
                            capturing_text = true;
                            text_target = TextTarget::Title(owner);
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
                        (ParseCtx::Chart, "autoTitleDeleted") => {
                            auto_title_deleted = Some(ct_boolean(e)?);
                        }
                        // Chart-type element attributes — route to
                        // primary or secondary accumulator.
                        (ParseCtx::ChartTypeElem, "barDir") => {
                            let val = attr_val_str(e)?;
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
                        (ParseCtx::ChartTypeElem, "grouping") => {
                            let g = ChartGrouping::from_xml(&attr_val_str(e)?);
                            if is_secondary_chart {
                                secondary_grouping = g;
                            } else {
                                grouping = g;
                            }
                        }
                        (ParseCtx::ChartTypeElem, "scatterStyle") => {
                            let s = ScatterStyle::from_xml(&attr_val_str(e)?);
                            if is_secondary_chart {
                                secondary_scatter_style = s;
                            } else {
                                scatter_style = s;
                            }
                        }
                        (ParseCtx::ChartTypeElem, "radarStyle") => {
                            let r = RadarStyle::from_xml(&attr_val_str(e)?);
                            if is_secondary_chart {
                                secondary_radar_style = r;
                            } else {
                                radar_style = r;
                            }
                        }
                        (ParseCtx::ChartTypeElem, "holeSize") => {
                            let h = hole_size_val(e)?;
                            if is_secondary_chart {
                                secondary_hole_size = h;
                            } else {
                                hole_size = h;
                            }
                        }
                        // The axes this chart-type element is drawn on. Every
                        // CT_*Chart caps `c:axId` at maxOccurs 3 (3-D charts name a
                        // series axis), so ids past the third are ignored and
                        // `take_axis` stays linear in the axes of a crafted part.
                        (ParseCtx::ChartTypeElem, "axId") => {
                            if let Some(id) = num_val(e)? {
                                let ids = if is_secondary_chart {
                                    &mut secondary_ax_ids
                                } else {
                                    &mut primary_ax_ids
                                };
                                if ids.len() < 3 {
                                    ids.push(id);
                                }
                            }
                        }
                        // Series attributes.
                        (ParseCtx::Series, "idx") => {
                            if let Some(ref mut ser) = cur_ser {
                                ser.idx = num_val(e)?.unwrap_or(0);
                            }
                        }
                        (ParseCtx::Series, "order") => {
                            if let Some(ref mut ser) = cur_ser {
                                ser.order = num_val(e)?.unwrap_or(0);
                            }
                        }
                        (ParseCtx::Series, "smooth") => {
                            if let Some(ref mut ser) = cur_ser {
                                ser.smooth = Some(ct_boolean(e)?);
                            }
                        }
                        (ParseCtx::Series, "explosion") => {
                            if let Some(ref mut ser) = cur_ser {
                                ser.explosion = num_val(e)?;
                            }
                        }
                        // An empty literal series name.
                        (ParseCtx::SerTx, "v") => {
                            if let Some(ref mut ser) = cur_ser {
                                ser.name_literal = Some(String::new());
                            }
                        }
                        // Fill color in spPr.
                        (ParseCtx::SerSpPrFill, "srgbClr") => {
                            if let Some(ref mut ser) = cur_ser {
                                ser.fill_color = attr_val(e, "val")?;
                            }
                        }
                        // Line color in spPr > ln.
                        (ParseCtx::SerSpPrLnFill, "srgbClr") => {
                            if let Some(ref mut ser) = cur_ser {
                                ser.line_color = attr_val(e, "val")?;
                            }
                        }
                        // A childless `<a:ln w>` in spPr.
                        (ParseCtx::SerSpPr, "ln") => {
                            if let Some(ref mut ser) = cur_ser
                                && let Some(w) = num_attr(e, "w")?
                            {
                                ser.line_width = Some(w);
                            }
                        }
                        // Marker symbol.
                        (ParseCtx::SerMarker, "symbol") => {
                            if let Some(ref mut ser) = cur_ser {
                                ser.marker = MarkerStyle::from_xml(&attr_val_str(e)?);
                            }
                        }
                        // Data labels (chart-level, then series-level).
                        (ParseCtx::DataLabels, _) => cur_dlbls.read_child(local, e)?,
                        (ParseCtx::SerDataLabels, _) => cur_ser_dlbls.read_child(local, e)?,
                        // Axis children.
                        (ParseCtx::CatAxis, _) => cur_cat_axis.read_child(local, e)?,
                        // `c:crossBetween` exists on CT_ValAx only.
                        (ParseCtx::ValAxis, "crossBetween") => {
                            cur_val_axis.cross_between = CrossBetween::from_xml(&attr_val_str(e)?);
                        }
                        (ParseCtx::ValAxis, _) => cur_val_axis.read_child(local, e)?,
                        (ParseCtx::AxisScaling(AxisKind::Cat), _) => {
                            cur_cat_axis.read_scaling_child(local, e)?;
                        }
                        (ParseCtx::AxisScaling(AxisKind::Val), _) => {
                            cur_val_axis.read_scaling_child(local, e)?;
                        }
                        // Axis txPr — defRPr with sz attribute.
                        (ParseCtx::AxisTxPr(kind), "defRPr") => {
                            let axis = match kind {
                                AxisKind::Cat => &mut cur_cat_axis,
                                AxisKind::Val => &mut cur_val_axis,
                            };
                            if let Some(sz) = num_attr(e, "sz")? {
                                axis.font_size = Some(sz);
                            }
                        }
                        // Legend.
                        (ParseCtx::Legend, "legendPos") => {
                            cur_legend.position = LegendPosition::from_xml(&attr_val_str(e)?);
                        }
                        (ParseCtx::Legend, "overlay") => {
                            cur_legend.overlay = ct_boolean(e)?;
                        }
                        // Title overlay.
                        (ParseCtx::Title(owner), "overlay") => {
                            titles[owner.index()].overlay = ct_boolean(e)?;
                        }
                        // Title defRPr / rPr (Empty variants).
                        (ParseCtx::TitlePPr(owner), "defRPr")
                        | (ParseCtx::TitleRun(owner), "rPr") => {
                            titles[owner.index()].read_run_props(e)?;
                        }
                        // Color in the title defRPr or rPr fill.
                        (
                            ParseCtx::TitleDefRPrFill(owner) | ParseCtx::TitleRunPrFill(owner),
                            "srgbClr",
                        ) => {
                            titles[owner.index()].color = attr_val(e, "val")?;
                        }
                        // Manual layout values.
                        (ParseCtx::ManualLayoutCtx, "x") => {
                            layout_x = num_val(e)?;
                        }
                        (ParseCtx::ManualLayoutCtx, "y") => {
                            layout_y = num_val(e)?;
                        }
                        (ParseCtx::ManualLayoutCtx, "w") => {
                            layout_w = num_val(e)?;
                        }
                        (ParseCtx::ManualLayoutCtx, "h") => {
                            layout_h = num_val(e)?;
                        }
                        // Trendline attributes.
                        (ParseCtx::SerTrendline, _) => {
                            if let Some(ref mut tb) = cur_trendline {
                                match local {
                                    "trendlineType" => {
                                        tb.trend_type = TrendlineType::from_xml(&attr_val_str(e)?);
                                    }
                                    "order" => tb.order = num_val(e)?,
                                    "period" => tb.period = num_val(e)?,
                                    "forward" => tb.forward = num_val(e)?,
                                    "backward" => tb.backward = num_val(e)?,
                                    "dispEq" => tb.display_eq = ct_boolean(e)?,
                                    "dispRSqr" => tb.display_r_sqr = ct_boolean(e)?,
                                    _ => {}
                                }
                            }
                        }
                        // Error bars attributes.
                        (ParseCtx::SerErrBars, _) => {
                            if let Some(ref mut eb) = cur_err_bars {
                                match local {
                                    "errDir" => {
                                        eb.err_dir = ErrorBarAxis::from_xml(&attr_val_str(e)?);
                                    }
                                    "errBarType" => {
                                        eb.direction =
                                            ErrorBarDirection::from_xml(&attr_val_str(e)?);
                                    }
                                    "errValType" => {
                                        eb.err_type = ErrorBarType::from_xml(&attr_val_str(e)?);
                                    }
                                    "val" => eb.value = num_val(e)?,
                                    _ => {}
                                }
                            }
                        }
                        // View3D attributes.
                        (ParseCtx::View3D, "rotX") => {
                            cur_view_3d.rot_x = num_val(e)?;
                        }
                        (ParseCtx::View3D, "rotY") => {
                            cur_view_3d.rot_y = num_val(e)?;
                        }
                        (ParseCtx::View3D, "perspective") => {
                            cur_view_3d.perspective = num_val(e)?;
                        }
                        (ParseCtx::View3D, "rAngAx") => {
                            cur_view_3d.r_ang_ax = Some(ct_boolean(e)?);
                        }
                        // Data table (showKeys within dTable).
                        (_, "dTable") => {
                            // Self-closing <c:dTable/> or containing <c:showKeys/>.
                            show_data_table = true;
                        }
                        // `c:style` is a child of `c:chartSpace` only; a `c14:style`
                        // sits in an `mc:Choice`, which is skipped above.
                        (ParseCtx::Root, "style") => {
                            style_id = num_val(e)?;
                        }
                        (ParseCtx::Root, "userShapes" | "externalData") => {
                            unmodelled.push(format!("a {local}"));
                        }
                        _ => {}
                    }
                }
                Ok(Event::End(ref e)) => {
                    let local = e.local_name();
                    let local = local.as_ref();
                    let ctx = current_ctx(&ctx_stack);

                    match (ctx, local) {
                        (ParseCtx::Series, "ser") => {
                            if let Some(ser) = cur_ser.take() {
                                budget.store(series_memory(&ser))?;
                                if is_secondary_chart {
                                    secondary_series.push(ser);
                                } else {
                                    series.push(ser);
                                }
                            }
                            ctx_stack.pop();
                        }
                        (ParseCtx::DataLabels, "dLbls") => {
                            let dl = Some(cur_dlbls.build_and_reset());
                            if is_secondary_chart {
                                secondary_data_labels = dl;
                            } else {
                                data_labels = dl;
                            }
                            ctx_stack.pop();
                        }
                        (ParseCtx::SerDataLabels, "dLbls") => {
                            if let Some(ref mut ser) = cur_ser {
                                ser.data_labels = Some(cur_ser_dlbls.build_and_reset());
                            }
                            ctx_stack.pop();
                        }
                        (ParseCtx::SerTrendline, "trendline") => {
                            if let Some(ref mut ser) = cur_ser
                                && let Some(tb) = cur_trendline.take()
                            {
                                ser.trendline = tb.build();
                            }
                            ctx_stack.pop();
                        }
                        (ParseCtx::SerErrBars, "errBars") => {
                            if let Some(ref mut ser) = cur_ser
                                && let Some(eb) = cur_err_bars.take()
                            {
                                ser.error_bars = eb.build();
                            }
                            ctx_stack.pop();
                        }
                        (ParseCtx::View3D, "view3D") => {
                            view_3d = cur_view_3d.build();
                            ctx_stack.pop();
                        }
                        (ParseCtx::CatAxis, "catAx") => {
                            let builder = std::mem::take(&mut cur_cat_axis);
                            let axis = builder.build_with_title(
                                titles[TitleOwner::CatAxis.index()].build_option(),
                            );
                            budget.store(axis_memory(&axis))?;
                            cat_axes.push(axis);
                            ctx_stack.pop();
                        }
                        (ParseCtx::ValAxis, "valAx") => {
                            let builder = std::mem::take(&mut cur_val_axis);
                            let axis = builder.build_with_title(
                                titles[TitleOwner::ValAxis.index()].build_option(),
                            );
                            budget.store(axis_memory(&axis))?;
                            val_axes.push(axis);
                            ctx_stack.pop();
                        }
                        (ParseCtx::Legend, "legend") => {
                            legend = Some(cur_legend.build());
                            ctx_stack.pop();
                        }
                        (ParseCtx::Title(TitleOwner::Chart), "title") => {
                            title = titles[TitleOwner::Chart.index()].build_option();
                            ctx_stack.pop();
                        }
                        // Formula capture end.
                        (_, "f")
                            if capturing_text
                                && matches!(
                                    text_target,
                                    TextTarget::SerName
                                        | TextTarget::CatRef
                                        | TextTarget::ValRef
                                        | TextTarget::XValRef
                                        | TextTarget::BubbleRef
                                        | TextTarget::ErrRef(_)
                                ) =>
                        {
                            let val = std::mem::take(&mut text_buf);
                            if let TextTarget::ErrRef(kind) = text_target {
                                if let Some(ref mut eb) = cur_err_bars {
                                    match kind {
                                        ErrRefKind::Plus => eb.plus_ref = Some(val),
                                        ErrRefKind::Minus => eb.minus_ref = Some(val),
                                    }
                                }
                            } else if let Some(ref mut ser) = cur_ser {
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
                        // Literal series name end.
                        (_, "v") if capturing_text && text_target == TextTarget::SerNameLiteral => {
                            if let Some(ref mut ser) = cur_ser {
                                ser.name_literal = Some(std::mem::take(&mut text_buf));
                            }
                            capturing_text = false;
                            text_target = TextTarget::None;
                            ctx_stack.pop();
                        }
                        // Title text end: runs append, so a title split over several
                        // runs keeps all of its text.
                        (_, "t")
                            if capturing_text && matches!(text_target, TextTarget::Title(_)) =>
                        {
                            if let TextTarget::Title(owner) = text_target {
                                titles[owner.index()].text.push_str(&text_buf);
                            }
                            text_buf.clear();
                            capturing_text = false;
                            text_target = TextTarget::None;
                            ctx_stack.pop();
                        }
                        // Manual layout end.
                        (ParseCtx::ManualLayoutCtx, "layout") => {
                            if let (Some(x), Some(y), Some(w), Some(h)) =
                                (layout_x, layout_y, layout_w, layout_h)
                            {
                                plot_area_layout = Some(ManualLayout { x, y, w, h });
                            }
                            ctx_stack.pop();
                        }
                        _ => {
                            // Pop context for any End that matches the depth (axis
                            // titles are read from `titles` when their axis closes).
                            ctx_stack.pop();
                        }
                    }
                }
                Ok(Event::Text(ref e)) if capturing_text => {
                    text_buf.push_str(&e.xml10_content());
                }
                Ok(Event::CData(ref e)) if capturing_text => {
                    text_buf.push_str(&e.xml10_content());
                }
                Ok(Event::GeneralRef(ref e)) if capturing_text => {
                    push_general_ref(&mut text_buf, e)?;
                }
                Ok(Event::Eof) => break,
                Err(err) => {
                    cold_path();
                    return Err(chart_xml_err(err));
                }
                _ => {}
            }
            buf.clear();
        }

        // Resolve the axes by the ids each chart-type element names: document order
        // says nothing (Excel writes a secondary-axis combo as catAx, valAx, valAx,
        // catAx). Scatter and bubble charts have two value axes, X then Y.
        let primary_type = chart_type.unwrap_or(ChartType::Bar);
        let xy = matches!(primary_type, ChartType::Scatter | ChartType::Bubble);
        let mut cat_axis = take_axis(&mut cat_axes, &primary_ax_ids);
        let mut val_axis = take_axis(&mut val_axes, &primary_ax_ids);
        // A scatter X axis read as a catAx (older writers) leaves val_axis the Y axis.
        let mut secondary_val_axis = if xy && cat_axis.is_none() {
            take_axis(&mut val_axes, &primary_ax_ids)
        } else {
            None
        };
        let mut secondary_cat_axis = None;
        if secondary_chart_type.is_some() {
            secondary_cat_axis = take_axis(&mut cat_axes, &secondary_ax_ids);
            if secondary_val_axis.is_none() {
                secondary_val_axis = take_axis(&mut val_axes, &secondary_ax_ids);
            }
        }
        // Chart-type elements that name no (or unknown) axes: document order.
        if cat_axis.is_none() {
            cat_axis = take_first(&mut cat_axes);
        }
        if val_axis.is_none() {
            val_axis = take_first(&mut val_axes);
        }
        if secondary_val_axis.is_none() {
            secondary_val_axis = take_first(&mut val_axes);
        }
        if secondary_cat_axis.is_none() {
            secondary_cat_axis = take_first(&mut cat_axes);
        }
        if !cat_axes.is_empty() || !val_axes.is_empty() {
            unmodelled.push(format!(
                "{} axes beyond the modelled ones",
                cat_axes.len() + val_axes.len()
            ));
        }

        let raw_xml = if unmodelled.is_empty() {
            None
        } else {
            let raw = std::str::from_utf8(data).map_err(|err| {
                cold_path();
                ModernXlsxError::XmlParse(format!("chart part is not UTF-8: {err}"))
            })?;
            warn!(
                "chart holds {}, which the model does not represent; the part is kept verbatim (rawXml)",
                unmodelled.join(" and ")
            );
            Some(raw.to_owned())
        };

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
                secondary_cat_axis: None,
                show_data_table: false,
                view_3d: None,
                auto_title_deleted: None,
                raw_xml: None,
                raw_rels: None,
            })
        });

        Ok(ChartData {
            chart_type: primary_type,
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
            secondary_cat_axis,
            show_data_table,
            view_3d,
            auto_title_deleted,
            raw_xml,
            raw_rels: None,
        })
    }
}

// ---------------------------------------------------------------------------
// Drawing XML Parser — parse drawing anchors for chart references
// ---------------------------------------------------------------------------

/// The chart anchors of a drawing part and how many anchors it holds in all.
#[derive(Debug, Clone, Default, PartialEq)]
pub struct DrawingAnchors {
    /// `(ChartAnchor, rId)` for each top-level `<xdr:twoCellAnchor>`,
    /// `<xdr:oneCellAnchor>` or `<xdr:absoluteAnchor>` whose direct-child
    /// `<xdr:graphicFrame>` holds a `<c:chart r:id>`, in document order. A chart
    /// inside an `<xdr:grpSp>` is not listed, so its drawing stays verbatim.
    pub charts: Vec<(ChartAnchor, String)>,
    /// Every direct child of `<xdr:wsDr>`: anchors of any kind (pictures,
    /// shapes, charts) and `<mc:AlternateContent>` blocks. Equal to
    /// `charts.len()` only for a drawing that holds charts and nothing else.
    pub total: usize,
}

/// The value of the first prefixed `*:id` attribute of a `<c:chart>` element.
/// Excel binds that prefix to the relationships namespace; the binding is not
/// checked.
///
/// # Errors
///
/// `XmlParse` for a malformed attribute list or a bad entity in the value.
fn chart_rel_id(e: &BytesStart<'_>) -> Result<Option<String>> {
    for attr in attrs(e) {
        let attr = attr?;
        if attr.key.as_ref().ends_with(":id") {
            return Ok(Some(attr_str(&attr)?.into_owned()));
        }
    }
    Ok(None)
}

/// The kind of top-level drawing anchor being read.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum AnchorKind {
    TwoCell,
    OneCell,
    Absolute,
}

/// Parse `xl/drawings/drawing{n}.xml` to extract chart anchors and chart rIds.
///
/// Only direct children of `<xdr:wsDr>` are anchors: an anchor nested in
/// `<mc:AlternateContent>` counts towards `DrawingAnchors::total` through
/// its wrapper and is never read as a chart. Only a `<xdr:graphicFrame>`
/// directly under the anchor counts as a chart: a chart inside `<xdr:grpSp>`
/// is not read, so the anchor is not chart-only and the drawing stays verbatim.
///
/// # Errors
///
/// `XmlParse` for malformed XML, a malformed attribute list, or a position,
/// offset or size that is not a number of the model's type.
pub fn parse_drawing_anchors(data: &[u8]) -> Result<DrawingAnchors> {
    let mut reader = Reader::from_reader(data);
    // Untrimmed, like the chart parser; numbers are trimmed where they are parsed.
    reader.config_mut().trim_text(false);
    let mut buf = Vec::with_capacity(512);

    let mut result: Vec<(ChartAnchor, String)> = Vec::new();
    let mut total = 0usize;
    // Open elements: 1 inside `<xdr:wsDr>`, 2 inside one of its children.
    let mut depth = 0usize;

    // State tracking.
    let mut in_anchor = false;
    let mut kind = AnchorKind::TwoCell;
    let mut in_from = false;
    let mut in_to = false;
    let mut in_graphic_frame = false;
    let mut anchor = ChartAnchor::default();
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
                let parent_depth = depth;
                depth += 1;
                if parent_depth == 1 {
                    total += 1;
                }
                let local = e.local_name();
                let local = local.as_ref();
                match local {
                    "twoCellAnchor" | "oneCellAnchor" | "absoluteAnchor" if parent_depth == 1 => {
                        in_anchor = true;
                        kind = match local {
                            "twoCellAnchor" => AnchorKind::TwoCell,
                            "oneCellAnchor" => AnchorKind::OneCell,
                            _ => AnchorKind::Absolute,
                        };
                        anchor = ChartAnchor::default();
                        chart_r_id = None;
                    }
                    "chart" if in_graphic_frame => {
                        if let Some(r_id) = chart_rel_id(e)? {
                            chart_r_id = Some(r_id);
                        }
                    }
                    "from" if in_anchor => {
                        in_from = true;
                    }
                    "to" if in_anchor => {
                        in_to = true;
                    }
                    "col" | "colOff" | "row" | "rowOff" if in_anchor && (in_from || in_to) => {
                        text_buf.clear();
                        text_target = match (local, in_from) {
                            ("col", true) => DrawingTextTarget::FromCol,
                            ("colOff", true) => DrawingTextTarget::FromColOff,
                            ("row", true) => DrawingTextTarget::FromRow,
                            ("rowOff", true) => DrawingTextTarget::FromRowOff,
                            ("col", false) => DrawingTextTarget::ToCol,
                            ("colOff", false) => DrawingTextTarget::ToColOff,
                            ("row", false) => DrawingTextTarget::ToRow,
                            _ => DrawingTextTarget::ToRowOff,
                        };
                    }
                    // Only a direct anchor child (parent_depth 2): a chart in an
                    // `<xdr:grpSp>` is not read, so the drawing stays verbatim.
                    "graphicFrame" if in_anchor && parent_depth == 2 => {
                        in_graphic_frame = true;
                    }
                    _ => {}
                }
            }
            Ok(Event::Empty(ref e)) => {
                if depth == 1 {
                    total += 1;
                }
                let local = e.local_name();
                let local = local.as_ref();
                // <c:chart r:id="rId1"/> inside the anchor's graphicFrame — may
                // also appear unprefixed as `chart`.
                if local == "chart"
                    && in_graphic_frame
                    && let Some(r_id) = chart_rel_id(e)?
                {
                    chart_r_id = Some(r_id);
                }
                // Direct children of the anchor (depth 2), never the <a:off>/<a:ext>
                // of the graphicFrame's transform.
                if in_anchor && depth == 2 {
                    match (local, kind) {
                        // <xdr:ext cx="..." cy="..."/>: oneCellAnchor and absoluteAnchor size.
                        ("ext", AnchorKind::OneCell | AnchorKind::Absolute) => {
                            anchor.ext_cx = num_attr(e, "cx")?;
                            anchor.ext_cy = num_attr(e, "cy")?;
                        }
                        // <xdr:pos x="..." y="..."/>: absoluteAnchor position.
                        ("pos", AnchorKind::Absolute) => {
                            anchor.pos_x = num_attr(e, "x")?;
                            anchor.pos_y = num_attr(e, "y")?;
                        }
                        _ => {}
                    }
                }
            }
            Ok(Event::End(ref e)) => {
                depth = depth.saturating_sub(1);
                let local = e.local_name();
                let local = local.as_ref();
                match local {
                    "twoCellAnchor" | "oneCellAnchor" | "absoluteAnchor" if depth == 1 => {
                        if let Some(r_id) = chart_r_id.take() {
                            result.push((std::mem::take(&mut anchor), r_id));
                        }
                        in_anchor = false;
                        in_from = false;
                        in_to = false;
                    }
                    "from" => {
                        in_from = false;
                    }
                    "to" => {
                        in_to = false;
                    }
                    "graphicFrame" => {
                        in_graphic_frame = false;
                    }
                    "col" | "colOff" | "row" | "rowOff"
                        if text_target != DrawingTextTarget::None =>
                    {
                        let elem = e.name();
                        let elem = elem.as_ref();
                        let raw = std::mem::take(&mut text_buf);
                        match text_target {
                            DrawingTextTarget::FromCol => anchor.from_col = parse_text(elem, &raw)?,
                            DrawingTextTarget::FromColOff => {
                                anchor.from_col_off = parse_text(elem, &raw)?;
                            }
                            DrawingTextTarget::FromRow => anchor.from_row = parse_text(elem, &raw)?,
                            DrawingTextTarget::FromRowOff => {
                                anchor.from_row_off = parse_text(elem, &raw)?;
                            }
                            DrawingTextTarget::ToCol => anchor.to_col = parse_text(elem, &raw)?,
                            DrawingTextTarget::ToColOff => {
                                anchor.to_col_off = parse_text(elem, &raw)?;
                            }
                            DrawingTextTarget::ToRow => anchor.to_row = parse_text(elem, &raw)?,
                            DrawingTextTarget::ToRowOff => {
                                anchor.to_row_off = parse_text(elem, &raw)?;
                            }
                            DrawingTextTarget::None => {}
                        }
                        text_target = DrawingTextTarget::None;
                    }
                    _ => {}
                }
            }
            Ok(Event::Text(ref e)) if text_target != DrawingTextTarget::None => {
                text_buf.push_str(&e.xml10_content());
            }
            Ok(Event::CData(ref e)) if text_target != DrawingTextTarget::None => {
                text_buf.push_str(&e.xml10_content());
            }
            Ok(Event::GeneralRef(ref e)) if text_target != DrawingTextTarget::None => {
                push_general_ref(&mut text_buf, e)?;
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

    Ok(DrawingAnchors {
        charts: result,
        total,
    })
}

#[cfg(test)]
mod tests {
    use super::*;
    use pretty_assertions::assert_eq;

    const NS: &str = r#"xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006""#;

    fn chart_anchor(r_id: &str) -> String {
        format!(
            r#"<xdr:twoCellAnchor><xdr:from><xdr:col>1</xdr:col></xdr:from><xdr:to><xdr:col>4</xdr:col></xdr:to><xdr:graphicFrame macro=""><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="{r_id}"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor>"#
        )
    }

    const PIC_ANCHOR: &str = r#"<xdr:oneCellAnchor><xdr:from><xdr:col>0</xdr:col></xdr:from><xdr:ext cx="1" cy="1"/><xdr:pic><xdr:blipFill><a:blip r:embed="rId9"/></xdr:blipFill></xdr:pic><xdr:clientData/></xdr:oneCellAnchor>"#;

    fn parse(children: &str) -> DrawingAnchors {
        parse_drawing_anchors(format!("<xdr:wsDr {NS}>{children}</xdr:wsDr>").as_bytes()).unwrap()
    }

    #[test]
    fn chart_only_drawing_has_as_many_anchors_as_charts() {
        let parsed = parse(&(chart_anchor("rId1") + &chart_anchor("rId2")));
        assert_eq!(parsed.total, 2);
        let ids: Vec<&str> = parsed.charts.iter().map(|(_, id)| id.as_str()).collect();
        assert_eq!(ids, ["rId1", "rId2"]);
        assert_eq!(parsed.charts[0].0.to_col, 4);
    }

    #[test]
    fn pictures_and_alternate_content_count_but_are_not_charts() {
        let wrapped = format!(
            "<mc:AlternateContent><mc:Choice Requires=\"c14\">{}</mc:Choice><mc:Fallback>{}</mc:Fallback></mc:AlternateContent>",
            chart_anchor("rId7"),
            chart_anchor("rId8")
        );
        let parsed = parse(&format!(
            "{PIC_ANCHOR}{}{wrapped}<xdr:absoluteAnchor/>",
            chart_anchor("rId1")
        ));
        assert_eq!(parsed.total, 4);
        assert_eq!(parsed.charts.len(), 1);
        assert_eq!(parsed.charts[0].1, "rId1");
    }

    #[test]
    fn empty_drawing_has_no_anchor() {
        let parsed = parse_drawing_anchors(format!("<xdr:wsDr {NS}/>").as_bytes()).unwrap();
        assert_eq!(parsed, DrawingAnchors::default());
    }

    #[test]
    fn cdata_anchor_cells_are_read() {
        let anchor = chart_anchor("rId1").replace(
            "<xdr:to><xdr:col>4</xdr:col>",
            "<xdr:to><xdr:col><![CDATA[4]]></xdr:col>",
        );
        let parsed = parse(&anchor);
        assert_eq!(parsed.charts[0].0.to_col, 4);
    }
}
