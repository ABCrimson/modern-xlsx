//! Chart definitions — `xl/charts/chart{n}.xml`.

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
}
