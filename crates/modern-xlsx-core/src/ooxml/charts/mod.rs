//! Chart definitions — `xl/charts/chart{n}.xml`.

mod parser;
mod types;
mod writer;

pub use parser::parse_drawing_anchors;
pub use types::*;

use serde::{Deserialize, Serialize};

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
                trendline: None,
                error_bars: None,
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
            secondary_chart: None, secondary_val_axis: None, show_data_table: false, view_3d: None,
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
            ext_cx: None, ext_cy: None,
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
                bar_dir_horizontal: None, style_id: None, plot_area_layout: None, secondary_chart: None, secondary_val_axis: None, show_data_table: false, view_3d: None,
            },
            anchor: ChartAnchor {
                from_col: 5, from_row: 0, from_col_off: 0, from_row_off: 0,
                to_col: 12, to_row: 18, to_col_off: 0, to_row_off: 0,
                ext_cx: None, ext_cy: None,
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
            bar_dir_horizontal: None, style_id: None, plot_area_layout: None, secondary_chart: None, secondary_val_axis: None, show_data_table: false, view_3d: None,
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
            font_size: None,
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

    #[test]
    fn test_trendline_serde() {
        let t = Trendline {
            trend_type: TrendlineType::Polynomial,
            order: Some(3),
            period: None,
            forward: Some(2.0),
            backward: None,
            display_eq: true,
            display_r_sqr: false,
        };
        let json = serde_json::to_string(&t).unwrap();
        let rt: Trendline = serde_json::from_str(&json).unwrap();
        assert_eq!(rt.trend_type, TrendlineType::Polynomial);
        assert_eq!(rt.order, Some(3));
        assert!(rt.display_eq);
    }

    #[test]
    fn test_error_bars_serde() {
        let eb = ErrorBars {
            err_type: ErrorBarType::Percentage,
            direction: ErrorBarDirection::Both,
            value: Some(5.0),
        };
        let json = serde_json::to_string(&eb).unwrap();
        let rt: ErrorBars = serde_json::from_str(&json).unwrap();
        assert_eq!(rt.err_type, ErrorBarType::Percentage);
        assert_eq!(rt.value, Some(5.0));
    }

    #[test]
    fn test_view3d_serde() {
        let v = View3D {
            rot_x: Some(15),
            rot_y: Some(20),
            perspective: Some(30),
            r_ang_ax: Some(true),
        };
        let json = serde_json::to_string(&v).unwrap();
        let rt: View3D = serde_json::from_str(&json).unwrap();
        assert_eq!(rt.rot_x, Some(15));
        assert!(rt.r_ang_ax.unwrap());
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
                trendline: None,
                error_bars: None,
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
            secondary_chart: None, secondary_val_axis: None, show_data_table: false, view_3d: None,
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
                trendline: None,
                error_bars: None,
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
            secondary_chart: None, secondary_val_axis: None, show_data_table: false, view_3d: None,
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
                trendline: None,
                error_bars: None,
            }],
            cat_axis: Some(ChartAxis { id: 0, cross_ax: 1, title: None, num_fmt: None, source_linked: false, min: None, max: None, major_unit: None, minor_unit: None, log_base: None, reversed: false, tick_lbl_pos: None, major_tick_mark: None, minor_tick_mark: None, major_gridlines: false, minor_gridlines: false, delete: false, position: None, crosses_at: None, font_size: None }),
            val_axis: Some(ChartAxis { id: 1, cross_ax: 0, title: None, num_fmt: None, source_linked: false, min: None, max: None, major_unit: None, minor_unit: None, log_base: None, reversed: false, tick_lbl_pos: None, major_tick_mark: None, minor_tick_mark: None, major_gridlines: true, minor_gridlines: false, delete: false, position: None, crosses_at: None, font_size: None }),
            legend: None, data_labels: None,
            grouping: Some(ChartGrouping::Standard),
            scatter_style: None, radar_style: None, hole_size: None,
            bar_dir_horizontal: None, style_id: None, plot_area_layout: None, secondary_chart: None, secondary_val_axis: None, show_data_table: false, view_3d: None,
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
                trendline: None,
                error_bars: None,
            }],
            cat_axis: None, val_axis: None,
            legend: Some(ChartLegend { position: LegendPosition::Right, overlay: false }),
            data_labels: Some(DataLabels { show_val: true, show_cat_name: false, show_ser_name: false, show_percent: true, num_fmt: None, show_leader_lines: true }),
            grouping: None, scatter_style: None, radar_style: None, hole_size: None,
            bar_dir_horizontal: None, style_id: None, plot_area_layout: None, secondary_chart: None, secondary_val_axis: None, show_data_table: false, view_3d: None,
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
                trendline: None,
                error_bars: None,
            }],
            cat_axis: None,
            val_axis: Some(ChartAxis { id: 0, cross_ax: 1, title: None, num_fmt: None, source_linked: false, min: None, max: None, major_unit: None, minor_unit: None, log_base: None, reversed: false, tick_lbl_pos: None, major_tick_mark: None, minor_tick_mark: None, major_gridlines: true, minor_gridlines: false, delete: false, position: None, crosses_at: None, font_size: None }),
            legend: None, data_labels: None,
            grouping: None,
            scatter_style: Some(ScatterStyle::LineMarker),
            radar_style: None, hole_size: None,
            bar_dir_horizontal: None, style_id: None, plot_area_layout: None, secondary_chart: None, secondary_val_axis: None, show_data_table: false, view_3d: None,
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
                trendline: None,
                error_bars: None,
            }],
            cat_axis: None, val_axis: None,
            legend: None, data_labels: None,
            grouping: None, scatter_style: None, radar_style: None,
            hole_size: Some(50),
            bar_dir_horizontal: None, style_id: None, plot_area_layout: None, secondary_chart: None, secondary_val_axis: None, show_data_table: false, view_3d: None,
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
                trendline: None,
                error_bars: None,
            }],
            cat_axis: Some(ChartAxis { id: 0, cross_ax: 1, title: None, num_fmt: None, source_linked: false, min: None, max: None, major_unit: None, minor_unit: None, log_base: None, reversed: false, tick_lbl_pos: None, major_tick_mark: None, minor_tick_mark: None, major_gridlines: false, minor_gridlines: false, delete: false, position: None, crosses_at: None, font_size: None }),
            val_axis: Some(ChartAxis { id: 1, cross_ax: 0, title: None, num_fmt: None, source_linked: false, min: None, max: None, major_unit: None, minor_unit: None, log_base: None, reversed: false, tick_lbl_pos: None, major_tick_mark: None, minor_tick_mark: None, major_gridlines: true, minor_gridlines: false, delete: false, position: None, crosses_at: None, font_size: None }),
            legend: None, data_labels: None,
            grouping: Some(ChartGrouping::Stacked),
            scatter_style: None, radar_style: None, hole_size: None,
            bar_dir_horizontal: None, style_id: None, plot_area_layout: None, secondary_chart: None, secondary_val_axis: None, show_data_table: false, view_3d: None,
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
                trendline: None,
                error_bars: None,
            }],
            cat_axis: Some(ChartAxis { id: 0, cross_ax: 1, title: None, num_fmt: None, source_linked: false, min: None, max: None, major_unit: None, minor_unit: None, log_base: None, reversed: false, tick_lbl_pos: None, major_tick_mark: None, minor_tick_mark: None, major_gridlines: false, minor_gridlines: false, delete: false, position: None, crosses_at: None, font_size: None }),
            val_axis: Some(ChartAxis { id: 1, cross_ax: 0, title: None, num_fmt: None, source_linked: false, min: None, max: None, major_unit: None, minor_unit: None, log_base: None, reversed: false, tick_lbl_pos: None, major_tick_mark: None, minor_tick_mark: None, major_gridlines: true, minor_gridlines: false, delete: false, position: None, crosses_at: None, font_size: None }),
            legend: None, data_labels: None,
            grouping: None, scatter_style: None,
            radar_style: Some(RadarStyle::Filled),
            hole_size: None,
            bar_dir_horizontal: None, style_id: None, plot_area_layout: None, secondary_chart: None, secondary_val_axis: None, show_data_table: false, view_3d: None,
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
            secondary_chart: None, secondary_val_axis: None, show_data_table: false, view_3d: None,
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
            bar_dir_horizontal: None, style_id: None, plot_area_layout: None, secondary_chart: None, secondary_val_axis: None, show_data_table: false, view_3d: None,
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
            bar_dir_horizontal: None, style_id: None, plot_area_layout: None, secondary_chart: None, secondary_val_axis: None, show_data_table: false, view_3d: None,
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
                bar_dir_horizontal: None, style_id: None, plot_area_layout: None, secondary_chart: None, secondary_val_axis: None, show_data_table: false, view_3d: None,
            },
            anchor: ChartAnchor {
                from_col: 0, from_row: 0, from_col_off: 0, from_row_off: 0,
                to_col: 8, to_row: 15, to_col_off: 0, to_row_off: 0,
                ext_cx: None, ext_cy: None,
            },
        }];
        let r_ids = vec!["rId1".to_string()];
        let xml = ChartAnchor::generate_drawing_xml(&charts, &r_ids).unwrap();
        let xml_str = std::str::from_utf8(&xml).unwrap();

        assert!(xml_str.contains("xmlns:xdr=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\""));
        assert!(xml_str.contains("xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\""));
        assert!(xml_str.contains("xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\""));
        assert!(xml_str.contains("<xdr:twoCellAnchor>"));
        assert!(xml_str.contains("<xdr:from>"));
        assert!(xml_str.contains("<xdr:to>"));
        assert!(xml_str.contains("<xdr:col>0</xdr:col>"));
        assert!(xml_str.contains("<xdr:col>8</xdr:col>"));
        assert!(xml_str.contains("<xdr:row>15</xdr:row>"));
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
                    bar_dir_horizontal: None, style_id: None, plot_area_layout: None, secondary_chart: None, secondary_val_axis: None, show_data_table: false, view_3d: None,
                },
                anchor: ChartAnchor {
                    from_col: 0, from_row: 0, from_col_off: 0, from_row_off: 0,
                    to_col: 8, to_row: 15, to_col_off: 0, to_row_off: 0,
                    ext_cx: None, ext_cy: None,
                },
            },
            WorksheetChart {
                chart: ChartData {
                    chart_type: ChartType::Line,
                    title: None, series: vec![], cat_axis: None, val_axis: None,
                    legend: None, data_labels: None, grouping: None,
                    scatter_style: None, radar_style: None, hole_size: None,
                    bar_dir_horizontal: None, style_id: None, plot_area_layout: None, secondary_chart: None, secondary_val_axis: None, show_data_table: false, view_3d: None,
                },
                anchor: ChartAnchor {
                    from_col: 10, from_row: 0, from_col_off: 100, from_row_off: 200,
                    to_col: 18, to_row: 20, to_col_off: 300, to_row_off: 400,
                    ext_cx: None, ext_cy: None,
                },
            },
        ];
        let r_ids = vec!["rId1".to_string(), "rId2".to_string()];
        let xml = ChartAnchor::generate_drawing_xml(&charts, &r_ids).unwrap();
        let xml_str = std::str::from_utf8(&xml).unwrap();

        assert_eq!(xml_str.matches("<xdr:twoCellAnchor>").count(), 2);
        assert!(xml_str.contains(r#"id="2" name="Chart 1""#));
        assert!(xml_str.contains(r#"id="3" name="Chart 2""#));
        assert!(xml_str.contains(r#"r:id="rId1""#));
        assert!(xml_str.contains(r#"r:id="rId2""#));
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
                data_labels: None, trendline: None, error_bars: None,
            }],
            cat_axis: Some(ChartAxis { id: 0, cross_ax: 1, title: None, num_fmt: None, source_linked: false, min: None, max: None, major_unit: None, minor_unit: None, log_base: None, reversed: false, tick_lbl_pos: None, major_tick_mark: None, minor_tick_mark: None, major_gridlines: false, minor_gridlines: false, delete: false, position: None, crosses_at: None, font_size: None }),
            val_axis: Some(ChartAxis { id: 1, cross_ax: 0, title: None, num_fmt: None, source_linked: false, min: None, max: None, major_unit: None, minor_unit: None, log_base: None, reversed: false, tick_lbl_pos: None, major_tick_mark: None, minor_tick_mark: None, major_gridlines: true, minor_gridlines: false, delete: false, position: None, crosses_at: None, font_size: None }),
            legend: Some(ChartLegend { position: LegendPosition::Bottom, overlay: false }),
            data_labels: None, grouping: Some(ChartGrouping::Clustered),
            scatter_style: None, radar_style: None, hole_size: None,
            bar_dir_horizontal: Some(true), style_id: Some(2),
            plot_area_layout: None,
            secondary_chart: None, secondary_val_axis: None, show_data_table: false, view_3d: None,
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
                marker: None, smooth: None, explosion: Some(25),
                data_labels: None, trendline: None, error_bars: None,
            }],
            cat_axis: None, val_axis: None, legend: None,
            data_labels: Some(DataLabels { show_val: true, show_cat_name: false, show_ser_name: false, show_percent: true, num_fmt: None, show_leader_lines: true }),
            grouping: None, scatter_style: None, radar_style: None, hole_size: None,
            bar_dir_horizontal: None, style_id: None, plot_area_layout: None, secondary_chart: None, secondary_val_axis: None, show_data_table: false, view_3d: None,
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
            chart_type: ChartType::Scatter, title: None,
            series: vec![ChartSeries {
                idx: 0, order: 0, name: None, cat_ref: None,
                val_ref: "Sheet1!$B$2:$B$5".into(),
                x_val_ref: Some("Sheet1!$A$2:$A$5".into()),
                bubble_size_ref: None,
                fill_color: None, line_color: None, line_width: None,
                marker: Some(MarkerStyle::Diamond), smooth: None, explosion: None,
                data_labels: None, trendline: None, error_bars: None,
            }],
            cat_axis: None,
            val_axis: Some(ChartAxis { id: 0, cross_ax: 1, title: None, num_fmt: None, source_linked: false, min: None, max: None, major_unit: None, minor_unit: None, log_base: None, reversed: false, tick_lbl_pos: None, major_tick_mark: None, minor_tick_mark: None, major_gridlines: true, minor_gridlines: false, delete: false, position: None, crosses_at: None, font_size: None }),
            legend: None, data_labels: None, grouping: None,
            scatter_style: Some(ScatterStyle::LineMarker),
            radar_style: None, hole_size: None,
            bar_dir_horizontal: None, style_id: None, plot_area_layout: None, secondary_chart: None, secondary_val_axis: None, show_data_table: false, view_3d: None,
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
        let charts = vec![WorksheetChart {
            chart: ChartData {
                chart_type: ChartType::Line, title: None, series: vec![], cat_axis: None, val_axis: None,
                legend: None, data_labels: None, grouping: None,
                scatter_style: None, radar_style: None, hole_size: None,
                bar_dir_horizontal: None, style_id: None, plot_area_layout: None, secondary_chart: None, secondary_val_axis: None, show_data_table: false, view_3d: None,
            },
            anchor: ChartAnchor {
                from_col: 2, from_row: 5, from_col_off: 100, from_row_off: 200,
                to_col: 12, to_row: 25, to_col_off: 300, to_row_off: 400,
                ext_cx: None, ext_cy: None,
            },
        }];
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
            chart_type: ChartType::Column, title: None,
            series: vec![ChartSeries { idx: 0, order: 0, name: None, cat_ref: None, val_ref: "Sheet1!$B$2:$B$5".into(), x_val_ref: None, bubble_size_ref: None, fill_color: None, line_color: None, line_width: None, marker: None, smooth: None, explosion: None, data_labels: None, trendline: None, error_bars: None }],
            cat_axis: Some(ChartAxis { id: 0, cross_ax: 1, title: None, num_fmt: None, source_linked: false, min: None, max: None, major_unit: None, minor_unit: None, log_base: None, reversed: false, tick_lbl_pos: None, major_tick_mark: None, minor_tick_mark: None, major_gridlines: false, minor_gridlines: false, delete: false, position: None, crosses_at: None, font_size: None }),
            val_axis: Some(ChartAxis { id: 1, cross_ax: 0, title: None, num_fmt: None, source_linked: false, min: None, max: None, major_unit: None, minor_unit: None, log_base: None, reversed: false, tick_lbl_pos: None, major_tick_mark: None, minor_tick_mark: None, major_gridlines: true, minor_gridlines: false, delete: false, position: None, crosses_at: None, font_size: None }),
            legend: None, data_labels: None, grouping: Some(ChartGrouping::Clustered),
            scatter_style: None, radar_style: None, hole_size: None,
            bar_dir_horizontal: None, style_id: None, plot_area_layout: None,
            secondary_chart: None, secondary_val_axis: None, show_data_table: false, view_3d: None,
        };
        let xml = original.to_xml().unwrap();
        let parsed = ChartData::parse(&xml).unwrap();
        assert_eq!(parsed.chart_type, ChartType::Column);
        assert_eq!(parsed.bar_dir_horizontal, Some(false));
    }

    #[test]
    fn test_doughnut_chart_parse_roundtrip() {
        let original = ChartData {
            chart_type: ChartType::Doughnut, title: None,
            series: vec![ChartSeries { idx: 0, order: 0, name: None, cat_ref: Some("Sheet1!$A$2:$A$5".into()), val_ref: "Sheet1!$B$2:$B$5".into(), x_val_ref: None, bubble_size_ref: None, fill_color: None, line_color: None, line_width: None, marker: None, smooth: None, explosion: None, data_labels: None, trendline: None, error_bars: None }],
            cat_axis: None, val_axis: None, legend: None, data_labels: None,
            grouping: None, scatter_style: None, radar_style: None, hole_size: Some(50),
            bar_dir_horizontal: None, style_id: None, plot_area_layout: None, secondary_chart: None, secondary_val_axis: None, show_data_table: false, view_3d: None,
        };
        let xml = original.to_xml().unwrap();
        let parsed = ChartData::parse(&xml).unwrap();
        assert_eq!(parsed.chart_type, ChartType::Doughnut);
        assert_eq!(parsed.hole_size, Some(50));
    }

    #[test]
    fn test_line_chart_with_markers_parse_roundtrip() {
        let original = ChartData {
            chart_type: ChartType::Line, title: None,
            series: vec![ChartSeries { idx: 0, order: 0, name: None, cat_ref: None, val_ref: "Sheet1!$B$2:$B$5".into(), x_val_ref: None, bubble_size_ref: None, fill_color: None, line_color: Some("FF0000".into()), line_width: Some(25400), marker: Some(MarkerStyle::Circle), smooth: Some(true), explosion: None, data_labels: None, trendline: None, error_bars: None }],
            cat_axis: Some(ChartAxis { id: 0, cross_ax: 1, title: None, num_fmt: None, source_linked: false, min: None, max: None, major_unit: None, minor_unit: None, log_base: None, reversed: false, tick_lbl_pos: None, major_tick_mark: None, minor_tick_mark: None, major_gridlines: false, minor_gridlines: false, delete: false, position: None, crosses_at: None, font_size: None }),
            val_axis: Some(ChartAxis { id: 1, cross_ax: 0, title: None, num_fmt: None, source_linked: false, min: None, max: None, major_unit: None, minor_unit: None, log_base: None, reversed: false, tick_lbl_pos: None, major_tick_mark: None, minor_tick_mark: None, major_gridlines: true, minor_gridlines: false, delete: false, position: None, crosses_at: None, font_size: None }),
            legend: None, data_labels: None, grouping: Some(ChartGrouping::Standard),
            scatter_style: None, radar_style: None, hole_size: None,
            bar_dir_horizontal: None, style_id: None, plot_area_layout: None, secondary_chart: None, secondary_val_axis: None, show_data_table: false, view_3d: None,
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
            cat_axis: Some(ChartAxis { id: 0, cross_ax: 1, title: Some(ChartTitle { text: "Month".into(), overlay: false, font_size: None, bold: None, color: None }), num_fmt: None, source_linked: false, min: None, max: None, major_unit: None, minor_unit: None, log_base: None, reversed: false, tick_lbl_pos: None, major_tick_mark: None, minor_tick_mark: None, major_gridlines: true, minor_gridlines: false, delete: false, position: Some(AxisPosition::Bottom), crosses_at: None, font_size: None }),
            val_axis: Some(ChartAxis { id: 1, cross_ax: 0, title: Some(ChartTitle { text: "Sales ($)".into(), overlay: false, font_size: Some(1200), bold: Some(true), color: Some("333333".into()) }), num_fmt: Some("#,##0".into()), source_linked: false, min: Some(0.0), max: Some(1000.0), major_unit: Some(200.0), minor_unit: None, log_base: None, reversed: false, tick_lbl_pos: None, major_tick_mark: None, minor_tick_mark: None, major_gridlines: true, minor_gridlines: false, delete: false, position: Some(AxisPosition::Left), crosses_at: None, font_size: None }),
            legend: None, data_labels: None, grouping: None,
            scatter_style: None, radar_style: None, hole_size: None,
            bar_dir_horizontal: None, style_id: None, plot_area_layout: None, secondary_chart: None, secondary_val_axis: None, show_data_table: false, view_3d: None,
        };
        let xml = original.to_xml().unwrap();
        let parsed = ChartData::parse(&xml).unwrap();
        assert_eq!(parsed.title.as_ref().unwrap().text, "My Chart");
        assert_eq!(parsed.title.as_ref().unwrap().color.as_deref(), Some("333333"));
        let cat_ax = parsed.cat_axis.as_ref().unwrap();
        assert_eq!(cat_ax.title.as_ref().unwrap().text, "Month");
        assert_eq!(cat_ax.position, Some(AxisPosition::Bottom));
        assert!(cat_ax.major_gridlines);
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
            chart_type: ChartType::Radar, title: None,
            series: vec![ChartSeries { idx: 0, order: 0, name: None, cat_ref: None, val_ref: "Sheet1!$B$2:$B$5".into(), x_val_ref: None, bubble_size_ref: None, fill_color: None, line_color: None, line_width: None, marker: None, smooth: None, explosion: None, data_labels: None, trendline: None, error_bars: None }],
            cat_axis: Some(ChartAxis { id: 0, cross_ax: 1, title: None, num_fmt: None, source_linked: false, min: None, max: None, major_unit: None, minor_unit: None, log_base: None, reversed: false, tick_lbl_pos: None, major_tick_mark: None, minor_tick_mark: None, major_gridlines: false, minor_gridlines: false, delete: false, position: None, crosses_at: None, font_size: None }),
            val_axis: Some(ChartAxis { id: 1, cross_ax: 0, title: None, num_fmt: None, source_linked: false, min: None, max: None, major_unit: None, minor_unit: None, log_base: None, reversed: false, tick_lbl_pos: None, major_tick_mark: None, minor_tick_mark: None, major_gridlines: true, minor_gridlines: false, delete: false, position: None, crosses_at: None, font_size: None }),
            legend: None, data_labels: None, grouping: None, scatter_style: None,
            radar_style: Some(RadarStyle::Filled), hole_size: None,
            bar_dir_horizontal: None, style_id: None, plot_area_layout: None, secondary_chart: None, secondary_val_axis: None, show_data_table: false, view_3d: None,
        };
        let xml = original.to_xml().unwrap();
        let parsed = ChartData::parse(&xml).unwrap();
        assert_eq!(parsed.chart_type, ChartType::Radar);
        assert_eq!(parsed.radar_style, Some(RadarStyle::Filled));
    }

    #[test]
    fn test_drawing_anchors_multiple_parse_roundtrip() {
        let charts = vec![
            WorksheetChart { chart: ChartData { chart_type: ChartType::Bar, title: None, series: vec![], cat_axis: None, val_axis: None, legend: None, data_labels: None, grouping: None, scatter_style: None, radar_style: None, hole_size: None, bar_dir_horizontal: None, style_id: None, plot_area_layout: None, secondary_chart: None, secondary_val_axis: None, show_data_table: false, view_3d: None }, anchor: ChartAnchor { from_col: 0, from_row: 0, from_col_off: 0, from_row_off: 0, to_col: 8, to_row: 15, to_col_off: 0, to_row_off: 0, ext_cx: None, ext_cy: None } },
            WorksheetChart { chart: ChartData { chart_type: ChartType::Line, title: None, series: vec![], cat_axis: None, val_axis: None, legend: None, data_labels: None, grouping: None, scatter_style: None, radar_style: None, hole_size: None, bar_dir_horizontal: None, style_id: None, plot_area_layout: None, secondary_chart: None, secondary_val_axis: None, show_data_table: false, view_3d: None }, anchor: ChartAnchor { from_col: 10, from_row: 0, from_col_off: 100, from_row_off: 200, to_col: 18, to_row: 20, to_col_off: 300, to_row_off: 400, ext_cx: None, ext_cy: None } },
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

    #[test]
    fn test_parse_one_cell_anchor_drawing() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <xdr:oneCellAnchor>
    <xdr:from>
      <xdr:col>3</xdr:col>
      <xdr:colOff>100</xdr:colOff>
      <xdr:row>5</xdr:row>
      <xdr:rowOff>200</xdr:rowOff>
    </xdr:from>
    <xdr:ext cx="5400000" cy="3240000"/>
    <xdr:graphicFrame macro="">
      <xdr:nvGraphicFramePr>
        <xdr:cNvPr id="2" name="Chart 1"/>
        <xdr:cNvGraphicFramePr/>
      </xdr:nvGraphicFramePr>
      <xdr:xfrm>
        <a:off x="0" y="0"/>
        <a:ext cx="0" cy="0"/>
      </xdr:xfrm>
      <a:graphic>
        <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
          <c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="rId1"/>
        </a:graphicData>
      </a:graphic>
    </xdr:graphicFrame>
    <xdr:clientData/>
  </xdr:oneCellAnchor>
</xdr:wsDr>"#;

        let anchors = parse_drawing_anchors(xml).unwrap();
        assert_eq!(anchors.len(), 1, "should parse oneCellAnchor");
        let (anchor, r_id) = &anchors[0];
        assert_eq!(r_id, "rId1");
        assert_eq!(anchor.from_col, 3);
        assert_eq!(anchor.from_row, 5);
        assert_eq!(anchor.from_col_off, 100);
        assert_eq!(anchor.from_row_off, 200);
        assert_eq!(anchor.ext_cx, Some(5_400_000));
        assert_eq!(anchor.ext_cy, Some(3_240_000));
        assert_eq!(anchor.to_col, 0);
        assert_eq!(anchor.to_row, 0);
    }

    #[test]
    fn test_one_cell_anchor_generate_and_parse_roundtrip() {
        let charts = vec![WorksheetChart {
            chart: ChartData { chart_type: ChartType::Bar, title: None, series: vec![], cat_axis: None, val_axis: None, legend: None, data_labels: None, grouping: None, scatter_style: None, radar_style: None, hole_size: None, bar_dir_horizontal: None, style_id: None, plot_area_layout: None, secondary_chart: None, secondary_val_axis: None, show_data_table: false, view_3d: None },
            anchor: ChartAnchor { from_col: 2, from_row: 3, from_col_off: 50, from_row_off: 75, to_col: 0, to_row: 0, to_col_off: 0, to_row_off: 0, ext_cx: Some(5_400_000), ext_cy: Some(3_240_000) },
        }];
        let r_ids = vec!["rId1".into()];
        let drawing_xml = ChartAnchor::generate_drawing_xml(&charts, &r_ids).unwrap();
        let xml_str = std::str::from_utf8(&drawing_xml).unwrap();

        assert!(xml_str.contains("<xdr:oneCellAnchor>"), "should write oneCellAnchor tag");
        assert!(!xml_str.contains("<xdr:twoCellAnchor>"), "should not write twoCellAnchor");
        assert!(xml_str.contains("cx=\"5400000\""), "should write ext cx");
        assert!(xml_str.contains("cy=\"3240000\""), "should write ext cy");
        assert!(!xml_str.contains("<xdr:to>"), "should not write <xdr:to>");

        let anchors = parse_drawing_anchors(&drawing_xml).unwrap();
        assert_eq!(anchors.len(), 1);
        let (anchor, r_id) = &anchors[0];
        assert_eq!(r_id, "rId1");
        assert_eq!(anchor.from_col, 2);
        assert_eq!(anchor.from_row, 3);
        assert_eq!(anchor.from_col_off, 50);
        assert_eq!(anchor.from_row_off, 75);
        assert_eq!(anchor.ext_cx, Some(5_400_000));
        assert_eq!(anchor.ext_cy, Some(3_240_000));
    }

    #[test]
    fn test_mixed_one_cell_and_two_cell_anchors() {
        let charts = vec![
            WorksheetChart { chart: ChartData { chart_type: ChartType::Bar, title: None, series: vec![], cat_axis: None, val_axis: None, legend: None, data_labels: None, grouping: None, scatter_style: None, radar_style: None, hole_size: None, bar_dir_horizontal: None, style_id: None, plot_area_layout: None, secondary_chart: None, secondary_val_axis: None, show_data_table: false, view_3d: None }, anchor: ChartAnchor { from_col: 0, from_row: 0, from_col_off: 0, from_row_off: 0, to_col: 8, to_row: 15, to_col_off: 0, to_row_off: 0, ext_cx: None, ext_cy: None } },
            WorksheetChart { chart: ChartData { chart_type: ChartType::Line, title: None, series: vec![], cat_axis: None, val_axis: None, legend: None, data_labels: None, grouping: None, scatter_style: None, radar_style: None, hole_size: None, bar_dir_horizontal: None, style_id: None, plot_area_layout: None, secondary_chart: None, secondary_val_axis: None, show_data_table: false, view_3d: None }, anchor: ChartAnchor { from_col: 10, from_row: 0, from_col_off: 0, from_row_off: 0, to_col: 0, to_row: 0, to_col_off: 0, to_row_off: 0, ext_cx: Some(7_200_000), ext_cy: Some(4_320_000) } },
        ];
        let r_ids = vec!["rId1".into(), "rId2".into()];
        let drawing_xml = ChartAnchor::generate_drawing_xml(&charts, &r_ids).unwrap();
        let xml_str = std::str::from_utf8(&drawing_xml).unwrap();

        assert_eq!(xml_str.matches("<xdr:twoCellAnchor>").count(), 1);
        assert_eq!(xml_str.matches("<xdr:oneCellAnchor>").count(), 1);

        let anchors = parse_drawing_anchors(&drawing_xml).unwrap();
        assert_eq!(anchors.len(), 2);
        assert_eq!(anchors[0].0.ext_cx, None);
        assert_eq!(anchors[0].0.ext_cy, None);
        assert_eq!(anchors[0].0.to_col, 8);
        assert_eq!(anchors[0].1, "rId1");
        assert_eq!(anchors[1].0.ext_cx, Some(7_200_000));
        assert_eq!(anchors[1].0.ext_cy, Some(4_320_000));
        assert_eq!(anchors[1].0.from_col, 10);
        assert_eq!(anchors[1].1, "rId2");
    }

    #[test]
    fn test_one_cell_anchor_serde() {
        let anchor = ChartAnchor {
            from_col: 3, from_row: 5, from_col_off: 0, from_row_off: 0,
            to_col: 0, to_row: 0, to_col_off: 0, to_row_off: 0,
            ext_cx: Some(5_400_000), ext_cy: Some(3_240_000),
        };
        let json = serde_json::to_string(&anchor).unwrap();
        assert!(json.contains("\"extCx\":5400000"));
        assert!(json.contains("\"extCy\":3240000"));
        let rt: ChartAnchor = serde_json::from_str(&json).unwrap();
        assert_eq!(rt.ext_cx, Some(5_400_000));
        assert_eq!(rt.ext_cy, Some(3_240_000));
        assert_eq!(rt.from_col, 3);

        let anchor2 = ChartAnchor {
            from_col: 0, from_row: 0, from_col_off: 0, from_row_off: 0,
            to_col: 8, to_row: 15, to_col_off: 0, to_row_off: 0,
            ext_cx: None, ext_cy: None,
        };
        let json2 = serde_json::to_string(&anchor2).unwrap();
        assert!(!json2.contains("extCx"));
        assert!(!json2.contains("extCy"));
    }

    #[test]
    fn combo_chart_parse_bar_plus_line() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
              xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <c:chart>
    <c:plotArea>
      <c:layout/>
      <c:barChart>
        <c:barDir val="col"/>
        <c:grouping val="clustered"/>
        <c:ser>
          <c:idx val="0"/>
          <c:order val="0"/>
          <c:tx><c:strRef><c:f>Sheet1!$B$1</c:f></c:strRef></c:tx>
          <c:val><c:numRef><c:f>Sheet1!$B$2:$B$5</c:f></c:numRef></c:val>
        </c:ser>
        <c:axId val="111"/>
        <c:axId val="222"/>
      </c:barChart>
      <c:lineChart>
        <c:grouping val="standard"/>
        <c:ser>
          <c:idx val="1"/>
          <c:order val="1"/>
          <c:tx><c:strRef><c:f>Sheet1!$C$1</c:f></c:strRef></c:tx>
          <c:val><c:numRef><c:f>Sheet1!$C$2:$C$5</c:f></c:numRef></c:val>
        </c:ser>
        <c:axId val="111"/>
        <c:axId val="333"/>
      </c:lineChart>
      <c:catAx>
        <c:axId val="111"/>
        <c:scaling><c:orientation val="minMax"/></c:scaling>
        <c:crossAx val="222"/>
      </c:catAx>
      <c:valAx>
        <c:axId val="222"/>
        <c:scaling><c:orientation val="minMax"/></c:scaling>
        <c:crossAx val="111"/>
      </c:valAx>
      <c:valAx>
        <c:axId val="333"/>
        <c:scaling><c:orientation val="minMax"/></c:scaling>
        <c:crossAx val="111"/>
        <c:axPos val="r"/>
      </c:valAx>
    </c:plotArea>
  </c:chart>
</c:chartSpace>"#;

        let chart = ChartData::parse(xml.as_bytes()).unwrap();
        assert_eq!(chart.chart_type, ChartType::Column);
        assert_eq!(chart.grouping, Some(ChartGrouping::Clustered));
        assert_eq!(chart.series.len(), 1);
        assert_eq!(chart.series[0].name.as_deref(), Some("Sheet1!$B$1"));
        assert_eq!(chart.series[0].val_ref, "Sheet1!$B$2:$B$5");

        let sec = chart.secondary_chart.as_ref().expect("secondary_chart missing");
        assert_eq!(sec.chart_type, ChartType::Line);
        assert_eq!(sec.grouping, Some(ChartGrouping::Standard));
        assert_eq!(sec.series.len(), 1);
        assert_eq!(sec.series[0].name.as_deref(), Some("Sheet1!$C$1"));
        assert_eq!(sec.series[0].val_ref, "Sheet1!$C$2:$C$5");
        assert_eq!(sec.series[0].idx, 1);

        let va = chart.val_axis.as_ref().expect("val_axis missing");
        assert_eq!(va.id, 222);
        assert_eq!(va.cross_ax, 111);

        let sva = chart.secondary_val_axis.as_ref().expect("secondary_val_axis missing");
        assert_eq!(sva.id, 333);
        assert_eq!(sva.cross_ax, 111);
        assert_eq!(sva.position, Some(AxisPosition::Right));
    }

    #[test]
    fn combo_chart_roundtrip() {
        let primary = ChartData {
            chart_type: ChartType::Column,
            title: Some(ChartTitle { text: "Combo".into(), overlay: false, font_size: None, bold: None, color: None }),
            series: vec![
                ChartSeries { idx: 0, order: 0, name: Some("Sales".into()), cat_ref: Some("Sheet1!$A$2:$A$5".into()),
                    val_ref: "Sheet1!$B$2:$B$5".into(), x_val_ref: None, bubble_size_ref: None,
                    fill_color: None, line_color: None, line_width: None, marker: None,
                    smooth: None, explosion: None, data_labels: None, trendline: None, error_bars: None },
            ],
            cat_axis: Some(ChartAxis { id: 100, cross_ax: 200, title: None, num_fmt: None, source_linked: false, min: None, max: None, major_unit: None, minor_unit: None, log_base: None, reversed: false, tick_lbl_pos: None, major_tick_mark: None, minor_tick_mark: None, major_gridlines: false, minor_gridlines: false, delete: false, position: None, crosses_at: None, font_size: None }),
            val_axis: Some(ChartAxis { id: 200, cross_ax: 100, title: None, num_fmt: None, source_linked: false, min: None, max: None, major_unit: None, minor_unit: None, log_base: None, reversed: false, tick_lbl_pos: None, major_tick_mark: None, minor_tick_mark: None, major_gridlines: false, minor_gridlines: false, delete: false, position: None, crosses_at: None, font_size: None }),
            legend: None, data_labels: None,
            grouping: Some(ChartGrouping::Clustered),
            scatter_style: None, radar_style: None, hole_size: None,
            bar_dir_horizontal: Some(false), style_id: None, plot_area_layout: None,
            secondary_chart: Some(Box::new(ChartData {
                chart_type: ChartType::Line, title: None,
                series: vec![
                    ChartSeries { idx: 1, order: 1, name: Some("Trend".into()), cat_ref: Some("Sheet1!$A$2:$A$5".into()),
                        val_ref: "Sheet1!$C$2:$C$5".into(), x_val_ref: None, bubble_size_ref: None,
                        fill_color: None, line_color: None, line_width: None, marker: None,
                        smooth: None, explosion: None, data_labels: None, trendline: None, error_bars: None },
                ],
                cat_axis: None, val_axis: None, legend: None, data_labels: None,
                grouping: Some(ChartGrouping::Standard),
                scatter_style: None, radar_style: None, hole_size: None,
                bar_dir_horizontal: None, style_id: None, plot_area_layout: None,
                secondary_chart: None, secondary_val_axis: None, show_data_table: false, view_3d: None,
            })),
            secondary_val_axis: Some(ChartAxis { id: 300, cross_ax: 100, title: None, num_fmt: None, source_linked: false, min: None, max: None, major_unit: None, minor_unit: None, log_base: None, reversed: false, tick_lbl_pos: None, major_tick_mark: None, minor_tick_mark: None, major_gridlines: false, minor_gridlines: false, delete: false, position: Some(AxisPosition::Right), crosses_at: None, font_size: None }),
            show_data_table: false, view_3d: None,
        };

        let xml_bytes = primary.to_xml().unwrap();
        let parsed = ChartData::parse(&xml_bytes).unwrap();

        assert_eq!(parsed.chart_type, ChartType::Column);
        assert_eq!(parsed.grouping, Some(ChartGrouping::Clustered));
        assert_eq!(parsed.series.len(), 1);
        assert_eq!(parsed.series[0].name.as_deref(), Some("Sales"));

        let sec = parsed.secondary_chart.as_ref().expect("secondary roundtrip");
        assert_eq!(sec.chart_type, ChartType::Line);
        assert_eq!(sec.grouping, Some(ChartGrouping::Standard));
        assert_eq!(sec.series.len(), 1);
        assert_eq!(sec.series[0].name.as_deref(), Some("Trend"));
        assert_eq!(sec.series[0].val_ref, "Sheet1!$C$2:$C$5");

        assert_eq!(parsed.val_axis.as_ref().unwrap().id, 200);
        let sva = parsed.secondary_val_axis.as_ref().expect("secondary_val_axis roundtrip");
        assert_eq!(sva.id, 300);
        assert_eq!(sva.cross_ax, 100);
        assert_eq!(sva.position, Some(AxisPosition::Right));
    }
}
