use super::*;

#[test]
fn test_chart_data_serde_roundtrip() {
    let chart = ChartData {
        title: Some(ChartTitle {
            text: "Sales".into(),
            ..Default::default()
        }),
        series: vec![ChartSeries {
            name: Some("Revenue".into()),
            cat_ref: Some("Sheet1!$A$2:$A$5".into()),
            val_ref: "Sheet1!$B$2:$B$5".into(),
            fill_color: Some("4472C4".into()),
            ..Default::default()
        }],
        cat_axis: Some(ChartAxis {
            cross_ax: 1,
            ..Default::default()
        }),
        val_axis: Some(ChartAxis {
            id: 1,
            major_gridlines: true,
            ..Default::default()
        }),
        legend: Some(ChartLegend {
            position: LegendPosition::Bottom,
            ..Default::default()
        }),
        grouping: Some(ChartGrouping::Clustered),
        bar_dir_horizontal: Some(false),
        style_id: Some(2),
        ..Default::default()
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
        to_col: 8,
        to_row: 15,
        ..Default::default()
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
            ..Default::default()
        },
        anchor: ChartAnchor {
            from_col: 5,
            to_col: 12,
            to_row: 18,
            ..Default::default()
        },
    };
    let json = serde_json::to_string(&wc).unwrap();
    assert!(json.contains("\"chartType\":\"pie\""));
}

#[test]
fn test_skip_serializing_optional_fields() {
    let chart = ChartData {
        chart_type: ChartType::Line,
        ..Default::default()
    };
    let json = serde_json::to_string(&chart).unwrap();
    assert!(!json.contains("title"));
    assert!(!json.contains("catAxis"));
    assert!(!json.contains("legend"));
    assert!(!json.contains("scatterStyle"));
}

#[test]
fn test_grouping_serde() {
    assert_eq!(
        serde_json::to_string(&ChartGrouping::Clustered).unwrap(),
        "\"clustered\""
    );
    assert_eq!(
        serde_json::to_string(&ChartGrouping::Stacked).unwrap(),
        "\"stacked\""
    );
    assert_eq!(
        serde_json::to_string(&ChartGrouping::PercentStacked).unwrap(),
        "\"percentStacked\""
    );
    assert_eq!(
        serde_json::to_string(&ChartGrouping::Standard).unwrap(),
        "\"standard\""
    );
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
        id: 1,
        title: Some(ChartTitle {
            text: "Revenue ($)".into(),
            font_size: Some(1200),
            bold: Some(true),
            ..Default::default()
        }),
        num_fmt: Some("#,##0".into()),
        min: Some(0.0),
        max: Some(100.0),
        major_unit: Some(20.0),
        minor_unit: Some(5.0),
        tick_lbl_pos: Some(TickLabelPosition::NextTo),
        major_tick_mark: Some(TickMark::Out),
        minor_tick_mark: Some(TickMark::None),
        major_gridlines: true,
        position: Some(AxisPosition::Left),
        crosses_at: Some(0.0),
        ..Default::default()
    };
    let json = serde_json::to_string(&axis).unwrap();
    let rt: ChartAxis = serde_json::from_str(&json).unwrap();
    assert_eq!(rt.min, Some(0.0));
    assert_eq!(rt.max, Some(100.0));
    assert_eq!(rt.position, Some(AxisPosition::Left));
}

#[test]
fn test_manual_layout_serde() {
    let layout = ManualLayout {
        x: 0.1,
        y: 0.15,
        w: 0.8,
        h: 0.7,
    };
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
        forward: Some(2.0),
        display_eq: true,
        ..Default::default()
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
        ..Default::default()
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
        title: Some(ChartTitle {
            text: "Sales".into(),
            ..Default::default()
        }),
        series: vec![ChartSeries {
            name: Some("Revenue".into()),
            cat_ref: Some("Sheet1!$A$2:$A$5".into()),
            val_ref: "Sheet1!$B$2:$B$5".into(),
            fill_color: Some("4472C4".into()),
            ..Default::default()
        }],
        cat_axis: Some(ChartAxis {
            cross_ax: 1,
            ..Default::default()
        }),
        val_axis: Some(ChartAxis {
            id: 1,
            major_gridlines: true,
            ..Default::default()
        }),
        legend: Some(ChartLegend {
            position: LegendPosition::Bottom,
            ..Default::default()
        }),
        grouping: Some(ChartGrouping::Clustered),
        bar_dir_horizontal: Some(true),
        ..Default::default()
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
        series: vec![ChartSeries {
            val_ref: "Sheet1!$B$2:$B$5".into(),
            ..Default::default()
        }],
        cat_axis: Some(ChartAxis {
            cross_ax: 1,
            ..Default::default()
        }),
        val_axis: Some(ChartAxis {
            id: 1,
            major_gridlines: true,
            ..Default::default()
        }),
        grouping: Some(ChartGrouping::Clustered),
        ..Default::default()
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
        series: vec![ChartSeries {
            val_ref: "Sheet1!$B$2:$B$5".into(),
            line_color: Some("FF0000".into()),
            line_width: Some(25400),
            marker: Some(MarkerStyle::Circle),
            smooth: Some(true),
            ..Default::default()
        }],
        cat_axis: Some(ChartAxis {
            cross_ax: 1,
            ..Default::default()
        }),
        val_axis: Some(ChartAxis {
            id: 1,
            major_gridlines: true,
            ..Default::default()
        }),
        grouping: Some(ChartGrouping::Standard),
        ..Default::default()
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
        title: Some(ChartTitle {
            text: "Distribution".into(),
            font_size: Some(1400),
            bold: Some(true),
            ..Default::default()
        }),
        series: vec![ChartSeries {
            name: Some("Shares".into()),
            cat_ref: Some("Sheet1!$A$2:$A$5".into()),
            val_ref: "Sheet1!$B$2:$B$5".into(),
            explosion: Some(25),
            ..Default::default()
        }],
        legend: Some(ChartLegend {
            ..Default::default()
        }),
        data_labels: Some(DataLabels {
            show_val: true,
            show_percent: true,
            show_leader_lines: true,
            ..Default::default()
        }),
        ..Default::default()
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
        series: vec![ChartSeries {
            val_ref: "Sheet1!$B$2:$B$5".into(),
            x_val_ref: Some("Sheet1!$A$2:$A$5".into()),
            marker: Some(MarkerStyle::Diamond),
            ..Default::default()
        }],
        val_axis: Some(ChartAxis {
            cross_ax: 1,
            major_gridlines: true,
            ..Default::default()
        }),
        scatter_style: Some(ScatterStyle::LineMarker),
        ..Default::default()
    };
    let xml = chart.to_xml().unwrap();
    let xml_str = std::str::from_utf8(&xml).unwrap();
    assert!(xml_str.contains("<c:scatterChart>"));
    assert!(xml_str.contains("<c:scatterStyle val=\"lineMarker\"/>"));
    assert!(xml_str.contains("<c:xVal>"));
    assert!(xml_str.contains("<c:yVal>"));
    assert!(!xml_str.contains("<c:cat>"));
    // CT_ScatterChart requires two c:axId, each naming an axis of the plot area.
    let element = element_text(xml_str, "c:scatterChart");
    assert_eq!(element.matches("<c:axId ").count(), 2, "{element}");
    assert_eq!(xml_str.matches("<c:valAx>").count(), 2, "{xml_str}");
}

/// The text of the first `<tag>...</tag>` element in `xml`, tags included.
fn element_text<'a>(xml: &'a str, tag: &str) -> &'a str {
    let start = xml
        .find(&format!("<{tag}>"))
        .unwrap_or_else(|| panic!("no <{tag}> in {xml}"));
    let close = format!("</{tag}>");
    let end = xml[start..]
        .find(&close)
        .unwrap_or_else(|| panic!("no </{tag}> in {xml}"));
    &xml[start..start + end + close.len()]
}

#[test]
fn scatter_chart_without_style_writes_line_marker() {
    let chart = ChartData {
        chart_type: ChartType::Scatter,
        series: vec![ChartSeries {
            val_ref: "Sheet1!$B$2:$B$5".into(),
            x_val_ref: Some("Sheet1!$A$2:$A$5".into()),
            ..Default::default()
        }],
        scatter_style: None,
        ..Default::default()
    };
    let xml = chart.to_xml().unwrap();
    let xml_str = std::str::from_utf8(&xml).unwrap();
    assert!(
        xml_str.contains("<c:scatterChart><c:scatterStyle val=\"lineMarker\"/>"),
        "{xml_str}"
    );
}

#[test]
fn radar_chart_without_style_writes_marker() {
    let chart = ChartData {
        chart_type: ChartType::Radar,
        series: vec![ChartSeries {
            val_ref: "Sheet1!$B$2:$B$5".into(),
            ..Default::default()
        }],
        radar_style: None,
        ..Default::default()
    };
    let xml = chart.to_xml().unwrap();
    let xml_str = std::str::from_utf8(&xml).unwrap();
    assert!(
        xml_str.contains("<c:radarChart><c:radarStyle val=\"marker\"/>"),
        "{xml_str}"
    );
}

#[test]
fn line_chart_without_axes_gets_two_crossing_axes() {
    let chart = ChartData {
        chart_type: ChartType::Line,
        series: vec![ChartSeries {
            cat_ref: Some("Sheet1!$A$2:$A$5".into()),
            val_ref: "Sheet1!$B$2:$B$5".into(),
            ..Default::default()
        }],
        cat_axis: None,
        val_axis: None,
        ..Default::default()
    };
    let xml = chart.to_xml().unwrap();
    let xml_str = std::str::from_utf8(&xml).unwrap();
    // CT_LineChart requires two c:axId; the smallest unused ids are 0 and 1.
    let line = element_text(xml_str, "c:lineChart");
    assert!(
        line.ends_with("<c:axId val=\"0\"/><c:axId val=\"1\"/></c:lineChart>"),
        "{line}"
    );
    let cat = element_text(xml_str, "c:catAx");
    assert!(cat.starts_with("<c:catAx><c:axId val=\"0\"/>"), "{cat}");
    assert!(cat.contains("<c:axPos val=\"b\"/>"), "{cat}");
    assert!(cat.contains("<c:crossAx val=\"1\"/>"), "{cat}");
    let val = element_text(xml_str, "c:valAx");
    assert!(val.starts_with("<c:valAx><c:axId val=\"1\"/>"), "{val}");
    assert!(val.contains("<c:axPos val=\"l\"/>"), "{val}");
    assert!(val.contains("<c:crossAx val=\"0\"/>"), "{val}");
    assert_eq!(xml_str.matches("<c:catAx>").count(), 1, "{xml_str}");
    assert_eq!(xml_str.matches("<c:valAx>").count(), 1, "{xml_str}");
}

#[test]
fn drawing_anchors_is_nameable_from_the_charts_module() {
    let empty: DrawingAnchors = parse_drawing_anchors(
        b"<xdr:wsDr xmlns:xdr=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\"/>",
    )
    .unwrap();
    assert_eq!(empty, DrawingAnchors::default());
}

#[test]
fn test_doughnut_chart_to_xml() {
    let chart = ChartData {
        chart_type: ChartType::Doughnut,
        series: vec![ChartSeries {
            cat_ref: Some("Sheet1!$A$2:$A$5".into()),
            val_ref: "Sheet1!$B$2:$B$5".into(),
            ..Default::default()
        }],
        hole_size: Some(50),
        ..Default::default()
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
        series: vec![ChartSeries {
            val_ref: "Sheet1!$B$2:$B$5".into(),
            fill_color: Some("4472C4".into()),
            ..Default::default()
        }],
        cat_axis: Some(ChartAxis {
            cross_ax: 1,
            ..Default::default()
        }),
        val_axis: Some(ChartAxis {
            id: 1,
            major_gridlines: true,
            ..Default::default()
        }),
        grouping: Some(ChartGrouping::Stacked),
        ..Default::default()
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
        series: vec![ChartSeries {
            val_ref: "Sheet1!$B$2:$B$5".into(),
            ..Default::default()
        }],
        cat_axis: Some(ChartAxis {
            cross_ax: 1,
            ..Default::default()
        }),
        val_axis: Some(ChartAxis {
            id: 1,
            major_gridlines: true,
            ..Default::default()
        }),
        radar_style: Some(RadarStyle::Filled),
        ..Default::default()
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
        style_id: Some(26),
        ..Default::default()
    };
    let xml = chart.to_xml().unwrap();
    let xml_str = std::str::from_utf8(&xml).unwrap();
    assert!(xml_str.contains("<c:style val=\"26\"/>"));
}

#[test]
fn test_chart_with_axis_title_and_scale() {
    let chart = ChartData {
        chart_type: ChartType::Line,
        cat_axis: Some(ChartAxis {
            cross_ax: 1,
            title: Some(ChartTitle {
                text: "Month".into(),
                ..Default::default()
            }),
            tick_lbl_pos: Some(TickLabelPosition::Low),
            major_tick_mark: Some(TickMark::Cross),
            major_gridlines: true,
            position: Some(AxisPosition::Bottom),
            ..Default::default()
        }),
        val_axis: Some(ChartAxis {
            id: 1,
            title: Some(ChartTitle {
                text: "Sales ($)".into(),
                font_size: Some(1200),
                bold: Some(true),
                color: Some("333333".into()),
                ..Default::default()
            }),
            num_fmt: Some("#,##0".into()),
            min: Some(0.0),
            max: Some(1000.0),
            major_unit: Some(200.0),
            major_gridlines: true,
            position: Some(AxisPosition::Left),
            ..Default::default()
        }),
        ..Default::default()
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
        ..Default::default()
    };
    let xml = chart.to_xml().unwrap();
    let xml_str = std::str::from_utf8(&xml).unwrap();
    assert!(xml_str.contains("xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\""));
    assert!(xml_str.contains("xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\""));
    assert!(xml_str.contains(
        "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\""
    ));
    assert!(xml_str.contains("<c:printSettings>"));
}

// =======================================================================
// Drawing XML tests
// =======================================================================

#[test]
fn test_generate_drawing_xml_single_chart() {
    let charts = vec![WorksheetChart {
        chart: ChartData {
            ..Default::default()
        },
        anchor: ChartAnchor {
            to_col: 8,
            to_row: 15,
            ..Default::default()
        },
    }];
    let r_ids = vec!["rId1".to_string()];
    let xml = ChartAnchor::generate_drawing_xml(&charts, &r_ids).unwrap();
    let xml_str = std::str::from_utf8(&xml).unwrap();

    assert!(xml_str.contains(
        "xmlns:xdr=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\""
    ));
    assert!(xml_str.contains("xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\""));
    assert!(xml_str.contains(
        "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\""
    ));
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
                ..Default::default()
            },
            anchor: ChartAnchor {
                to_col: 8,
                to_row: 15,
                ..Default::default()
            },
        },
        WorksheetChart {
            chart: ChartData {
                chart_type: ChartType::Line,
                ..Default::default()
            },
            anchor: ChartAnchor {
                from_col: 10,
                from_col_off: 100,
                from_row_off: 200,
                to_col: 18,
                to_row: 20,
                to_col_off: 300,
                to_row_off: 400,
                ..Default::default()
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
        title: Some(ChartTitle {
            text: "Sales".into(),
            ..Default::default()
        }),
        series: vec![ChartSeries {
            name_literal: Some("Revenue".into()),
            cat_ref: Some("Sheet1!$A$2:$A$5".into()),
            val_ref: "Sheet1!$B$2:$B$5".into(),
            fill_color: Some("4472C4".into()),
            ..Default::default()
        }],
        cat_axis: Some(ChartAxis {
            cross_ax: 1,
            ..Default::default()
        }),
        val_axis: Some(ChartAxis {
            id: 1,
            major_gridlines: true,
            ..Default::default()
        }),
        legend: Some(ChartLegend {
            position: LegendPosition::Bottom,
            ..Default::default()
        }),
        grouping: Some(ChartGrouping::Clustered),
        bar_dir_horizontal: Some(true),
        style_id: Some(2),
        ..Default::default()
    };
    let xml = original.to_xml().unwrap();
    let parsed = ChartData::parse(&xml).unwrap();
    assert_eq!(parsed.chart_type, ChartType::Bar);
    assert_eq!(parsed.series.len(), 1);
    assert_eq!(parsed.series[0].val_ref, "Sheet1!$B$2:$B$5");
    // A display name is a literal (`c:tx/c:v`), not a formula.
    assert_eq!(parsed.series[0].name, None);
    assert_eq!(parsed.series[0].name_literal.as_deref(), Some("Revenue"));
    assert_eq!(parsed.series[0].fill_color.as_deref(), Some("4472C4"));
    assert_eq!(parsed.grouping, Some(ChartGrouping::Clustered));
    assert_eq!(parsed.bar_dir_horizontal, Some(true));
    assert!(parsed.cat_axis.is_some());
    assert!(parsed.val_axis.is_some());
    assert_eq!(
        parsed.legend.as_ref().unwrap().position,
        LegendPosition::Bottom
    );
    assert_eq!(parsed.style_id, Some(2));
}

#[test]
fn test_pie_chart_parse_roundtrip() {
    let original = ChartData {
        chart_type: ChartType::Pie,
        title: Some(ChartTitle {
            text: "Distribution".into(),
            font_size: Some(1400),
            bold: Some(true),
            ..Default::default()
        }),
        series: vec![ChartSeries {
            cat_ref: Some("Sheet1!$A$2:$A$5".into()),
            val_ref: "Sheet1!$B$2:$B$5".into(),
            explosion: Some(25),
            ..Default::default()
        }],
        data_labels: Some(DataLabels {
            show_val: true,
            show_percent: true,
            show_leader_lines: true,
            ..Default::default()
        }),
        ..Default::default()
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
        series: vec![ChartSeries {
            val_ref: "Sheet1!$B$2:$B$5".into(),
            x_val_ref: Some("Sheet1!$A$2:$A$5".into()),
            marker: Some(MarkerStyle::Diamond),
            ..Default::default()
        }],
        val_axis: Some(ChartAxis {
            cross_ax: 1,
            major_gridlines: true,
            ..Default::default()
        }),
        scatter_style: Some(ScatterStyle::LineMarker),
        ..Default::default()
    };
    let xml = original.to_xml().unwrap();
    let parsed = ChartData::parse(&xml).unwrap();
    assert_eq!(parsed.chart_type, ChartType::Scatter);
    assert_eq!(parsed.scatter_style, Some(ScatterStyle::LineMarker));
    assert_eq!(
        parsed.series[0].x_val_ref.as_deref(),
        Some("Sheet1!$A$2:$A$5")
    );
    assert_eq!(parsed.series[0].marker, Some(MarkerStyle::Diamond));
}

#[test]
fn test_drawing_anchors_parse_roundtrip() {
    let charts = vec![WorksheetChart {
        chart: ChartData {
            chart_type: ChartType::Line,
            ..Default::default()
        },
        anchor: ChartAnchor {
            from_col: 2,
            from_row: 5,
            from_col_off: 100,
            from_row_off: 200,
            to_col: 12,
            to_row: 25,
            to_col_off: 300,
            to_row_off: 400,
            ..Default::default()
        },
    }];
    let r_ids = vec!["rId1".into()];
    let drawing_xml = ChartAnchor::generate_drawing_xml(&charts, &r_ids).unwrap();
    let anchors = parse_drawing_anchors(&drawing_xml).unwrap().charts;
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
        series: vec![ChartSeries {
            val_ref: "Sheet1!$B$2:$B$5".into(),
            ..Default::default()
        }],
        cat_axis: Some(ChartAxis {
            cross_ax: 1,
            ..Default::default()
        }),
        val_axis: Some(ChartAxis {
            id: 1,
            major_gridlines: true,
            ..Default::default()
        }),
        grouping: Some(ChartGrouping::Clustered),
        ..Default::default()
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
        series: vec![ChartSeries {
            cat_ref: Some("Sheet1!$A$2:$A$5".into()),
            val_ref: "Sheet1!$B$2:$B$5".into(),
            ..Default::default()
        }],
        hole_size: Some(50),
        ..Default::default()
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
        series: vec![ChartSeries {
            val_ref: "Sheet1!$B$2:$B$5".into(),
            line_color: Some("FF0000".into()),
            line_width: Some(25400),
            marker: Some(MarkerStyle::Circle),
            smooth: Some(true),
            ..Default::default()
        }],
        cat_axis: Some(ChartAxis {
            cross_ax: 1,
            ..Default::default()
        }),
        val_axis: Some(ChartAxis {
            id: 1,
            major_gridlines: true,
            ..Default::default()
        }),
        grouping: Some(ChartGrouping::Standard),
        ..Default::default()
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
        title: Some(ChartTitle {
            text: "My Chart".into(),
            font_size: Some(1400),
            bold: Some(true),
            color: Some("333333".into()),
            ..Default::default()
        }),
        cat_axis: Some(ChartAxis {
            cross_ax: 1,
            title: Some(ChartTitle {
                text: "Month".into(),
                ..Default::default()
            }),
            major_gridlines: true,
            position: Some(AxisPosition::Bottom),
            ..Default::default()
        }),
        val_axis: Some(ChartAxis {
            id: 1,
            title: Some(ChartTitle {
                text: "Sales ($)".into(),
                font_size: Some(1200),
                bold: Some(true),
                color: Some("333333".into()),
                ..Default::default()
            }),
            num_fmt: Some("#,##0".into()),
            min: Some(0.0),
            max: Some(1000.0),
            major_unit: Some(200.0),
            major_gridlines: true,
            position: Some(AxisPosition::Left),
            ..Default::default()
        }),
        ..Default::default()
    };
    let xml = original.to_xml().unwrap();
    let parsed = ChartData::parse(&xml).unwrap();
    assert_eq!(parsed.title.as_ref().unwrap().text, "My Chart");
    assert_eq!(
        parsed.title.as_ref().unwrap().color.as_deref(),
        Some("333333")
    );
    let cat_ax = parsed.cat_axis.as_ref().unwrap();
    assert_eq!(cat_ax.title.as_ref().unwrap().text, "Month");
    assert_eq!(cat_ax.position, Some(AxisPosition::Bottom));
    assert!(cat_ax.major_gridlines);
    let val_ax = parsed.val_axis.as_ref().unwrap();
    assert_eq!(val_ax.title.as_ref().unwrap().text, "Sales ($)");
    assert_eq!(
        val_ax.title.as_ref().unwrap().color.as_deref(),
        Some("333333")
    );
    assert_eq!(val_ax.min, Some(0.0));
    assert_eq!(val_ax.max, Some(1000.0));
    assert_eq!(val_ax.major_unit, Some(200.0));
}

#[test]
fn test_radar_chart_parse_roundtrip() {
    let original = ChartData {
        chart_type: ChartType::Radar,
        series: vec![ChartSeries {
            val_ref: "Sheet1!$B$2:$B$5".into(),
            ..Default::default()
        }],
        cat_axis: Some(ChartAxis {
            cross_ax: 1,
            ..Default::default()
        }),
        val_axis: Some(ChartAxis {
            id: 1,
            major_gridlines: true,
            ..Default::default()
        }),
        radar_style: Some(RadarStyle::Filled),
        ..Default::default()
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
                ..Default::default()
            },
            anchor: ChartAnchor {
                to_col: 8,
                to_row: 15,
                ..Default::default()
            },
        },
        WorksheetChart {
            chart: ChartData {
                chart_type: ChartType::Line,
                ..Default::default()
            },
            anchor: ChartAnchor {
                from_col: 10,
                from_col_off: 100,
                from_row_off: 200,
                to_col: 18,
                to_row: 20,
                to_col_off: 300,
                to_row_off: 400,
                ..Default::default()
            },
        },
    ];
    let r_ids = vec!["rId1".into(), "rId2".into()];
    let drawing_xml = ChartAnchor::generate_drawing_xml(&charts, &r_ids).unwrap();
    let anchors = parse_drawing_anchors(&drawing_xml).unwrap().charts;
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

    let anchors = parse_drawing_anchors(xml).unwrap().charts;
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
        chart: ChartData {
            ..Default::default()
        },
        anchor: ChartAnchor {
            from_col: 2,
            from_row: 3,
            from_col_off: 50,
            from_row_off: 75,
            ext_cx: Some(5_400_000),
            ext_cy: Some(3_240_000),
            ..Default::default()
        },
    }];
    let r_ids = vec!["rId1".into()];
    let drawing_xml = ChartAnchor::generate_drawing_xml(&charts, &r_ids).unwrap();
    let xml_str = std::str::from_utf8(&drawing_xml).unwrap();

    assert!(
        xml_str.contains("<xdr:oneCellAnchor>"),
        "should write oneCellAnchor tag"
    );
    assert!(
        !xml_str.contains("<xdr:twoCellAnchor>"),
        "should not write twoCellAnchor"
    );
    assert!(xml_str.contains("cx=\"5400000\""), "should write ext cx");
    assert!(xml_str.contains("cy=\"3240000\""), "should write ext cy");
    assert!(!xml_str.contains("<xdr:to>"), "should not write <xdr:to>");

    let anchors = parse_drawing_anchors(&drawing_xml).unwrap().charts;
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
        WorksheetChart {
            chart: ChartData {
                ..Default::default()
            },
            anchor: ChartAnchor {
                to_col: 8,
                to_row: 15,
                ..Default::default()
            },
        },
        WorksheetChart {
            chart: ChartData {
                chart_type: ChartType::Line,
                ..Default::default()
            },
            anchor: ChartAnchor {
                from_col: 10,
                ext_cx: Some(7_200_000),
                ext_cy: Some(4_320_000),
                ..Default::default()
            },
        },
    ];
    let r_ids = vec!["rId1".into(), "rId2".into()];
    let drawing_xml = ChartAnchor::generate_drawing_xml(&charts, &r_ids).unwrap();
    let xml_str = std::str::from_utf8(&drawing_xml).unwrap();

    assert_eq!(xml_str.matches("<xdr:twoCellAnchor>").count(), 1);
    assert_eq!(xml_str.matches("<xdr:oneCellAnchor>").count(), 1);

    let anchors = parse_drawing_anchors(&drawing_xml).unwrap().charts;
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
        from_col: 3,
        from_row: 5,
        ext_cx: Some(5_400_000),
        ext_cy: Some(3_240_000),
        ..Default::default()
    };
    let json = serde_json::to_string(&anchor).unwrap();
    assert!(json.contains("\"extCx\":5400000"));
    assert!(json.contains("\"extCy\":3240000"));
    let rt: ChartAnchor = serde_json::from_str(&json).unwrap();
    assert_eq!(rt.ext_cx, Some(5_400_000));
    assert_eq!(rt.ext_cy, Some(3_240_000));
    assert_eq!(rt.from_col, 3);

    let anchor2 = ChartAnchor {
        to_col: 8,
        to_row: 15,
        ..Default::default()
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

    let sec = chart
        .secondary_chart
        .as_ref()
        .expect("secondary_chart missing");
    assert_eq!(sec.chart_type, ChartType::Line);
    assert_eq!(sec.grouping, Some(ChartGrouping::Standard));
    assert_eq!(sec.series.len(), 1);
    assert_eq!(sec.series[0].name.as_deref(), Some("Sheet1!$C$1"));
    assert_eq!(sec.series[0].val_ref, "Sheet1!$C$2:$C$5");
    assert_eq!(sec.series[0].idx, 1);

    let va = chart.val_axis.as_ref().expect("val_axis missing");
    assert_eq!(va.id, 222);
    assert_eq!(va.cross_ax, 111);

    let sva = chart
        .secondary_val_axis
        .as_ref()
        .expect("secondary_val_axis missing");
    assert_eq!(sva.id, 333);
    assert_eq!(sva.cross_ax, 111);
    assert_eq!(sva.position, Some(AxisPosition::Right));
}

#[test]
fn combo_chart_roundtrip() {
    let primary = ChartData {
        chart_type: ChartType::Column,
        title: Some(ChartTitle {
            text: "Combo".into(),
            ..Default::default()
        }),
        series: vec![ChartSeries {
            name: Some("Sales".into()),
            cat_ref: Some("Sheet1!$A$2:$A$5".into()),
            val_ref: "Sheet1!$B$2:$B$5".into(),
            ..Default::default()
        }],
        cat_axis: Some(ChartAxis {
            id: 100,
            cross_ax: 200,
            ..Default::default()
        }),
        val_axis: Some(ChartAxis {
            id: 200,
            cross_ax: 100,
            ..Default::default()
        }),
        grouping: Some(ChartGrouping::Clustered),
        bar_dir_horizontal: Some(false),
        secondary_chart: Some(Box::new(ChartData {
            chart_type: ChartType::Line,
            series: vec![ChartSeries {
                idx: 1,
                order: 1,
                name: Some("Trend".into()),
                cat_ref: Some("Sheet1!$A$2:$A$5".into()),
                val_ref: "Sheet1!$C$2:$C$5".into(),
                ..Default::default()
            }],
            grouping: Some(ChartGrouping::Standard),
            ..Default::default()
        })),
        secondary_val_axis: Some(ChartAxis {
            id: 300,
            cross_ax: 100,
            position: Some(AxisPosition::Right),
            ..Default::default()
        }),
        ..Default::default()
    };

    let xml_bytes = primary.to_xml().unwrap();
    let parsed = ChartData::parse(&xml_bytes).unwrap();

    assert_eq!(parsed.chart_type, ChartType::Column);
    assert_eq!(parsed.grouping, Some(ChartGrouping::Clustered));
    assert_eq!(parsed.series.len(), 1);
    assert_eq!(parsed.series[0].name.as_deref(), Some("Sales"));

    let sec = parsed
        .secondary_chart
        .as_ref()
        .expect("secondary roundtrip");
    assert_eq!(sec.chart_type, ChartType::Line);
    assert_eq!(sec.grouping, Some(ChartGrouping::Standard));
    assert_eq!(sec.series.len(), 1);
    assert_eq!(sec.series[0].name.as_deref(), Some("Trend"));
    assert_eq!(sec.series[0].val_ref, "Sheet1!$C$2:$C$5");

    assert_eq!(parsed.val_axis.as_ref().unwrap().id, 200);
    let sva = parsed
        .secondary_val_axis
        .as_ref()
        .expect("secondary_val_axis roundtrip");
    assert_eq!(sva.id, 300);
    assert_eq!(sva.cross_ax, 100);
    assert_eq!(sva.position, Some(AxisPosition::Right));
}

// =======================================================================
// Schema-order tests (ECMA-376 dml-chart.xsd)
// =======================================================================

fn chart_xml(chart: &ChartData) -> String {
    String::from_utf8(chart.to_xml().unwrap()).unwrap()
}

/// Assert that every needle occurs and that the first occurrences are strictly increasing.
#[track_caller]
fn assert_in_order(xml: &str, needles: &[&str]) {
    let mut prev = 0;
    let mut prev_needle = "";
    for needle in needles {
        let at = xml
            .find(needle)
            .unwrap_or_else(|| panic!("{needle} missing from {xml}"));
        assert!(
            prev_needle.is_empty() || at > prev,
            "{needle} (at {at}) must come after {prev_needle} (at {prev}) in {xml}"
        );
        prev = at;
        prev_needle = needle;
    }
}

fn base_series() -> ChartSeries {
    ChartSeries {
        name: Some("Sheet1!$B$1".into()),
        cat_ref: Some("Sheet1!$A$2:$A$5".into()),
        val_ref: "Sheet1!$B$2:$B$5".into(),
        ..Default::default()
    }
}

#[test]
fn series_children_follow_schema_order() {
    let chart = ChartData {
        chart_type: ChartType::Line,
        series: vec![ChartSeries {
            fill_color: Some("4472C4".into()),
            marker: Some(MarkerStyle::Circle),
            smooth: Some(true),
            data_labels: Some(DataLabels {
                show_val: true,
                ..Default::default()
            }),
            trendline: Some(Trendline::default()),
            error_bars: Some(ErrorBars::default()),
            ..base_series()
        }],
        ..Default::default()
    };
    let xml = chart_xml(&chart);
    assert_in_order(
        &xml,
        &[
            "<c:idx",
            "<c:order",
            "<c:tx>",
            "<c:spPr>",
            "<c:marker>",
            "<c:dLbls>",
            "<c:trendline>",
            "<c:errBars>",
            "<c:cat>",
            "<c:val>",
            "<c:smooth",
        ],
    );
}

#[test]
fn scatter_series_children_follow_schema_order() {
    let chart = ChartData {
        chart_type: ChartType::Scatter,
        series: vec![ChartSeries {
            x_val_ref: Some("Sheet1!$A$2:$A$5".into()),
            marker: Some(MarkerStyle::Square),
            smooth: Some(false),
            trendline: Some(Trendline::default()),
            ..base_series()
        }],
        ..Default::default()
    };
    let xml = chart_xml(&chart);
    assert_in_order(
        &xml,
        &[
            "<c:marker>",
            "<c:trendline>",
            "<c:xVal>",
            "<c:yVal>",
            "<c:smooth",
        ],
    );
}

#[test]
fn axis_scaling_follows_schema_order() {
    let chart = ChartData {
        val_axis: Some(ChartAxis {
            id: 2,
            cross_ax: 1,
            min: Some(0.0),
            max: Some(10.0),
            log_base: Some(10.0),
            ..Default::default()
        }),
        series: vec![base_series()],
        ..Default::default()
    };
    let xml = chart_xml(&chart);
    assert!(
        xml.contains(
            "<c:scaling><c:logBase val=\"10\"/><c:orientation val=\"minMax\"/><c:max val=\"10\"/><c:min val=\"0\"/></c:scaling>"
        ),
        "{xml}"
    );
}

#[test]
fn axis_gridlines_precede_title() {
    let chart = ChartData {
        cat_axis: Some(ChartAxis {
            id: 1,
            cross_ax: 2,
            major_gridlines: true,
            minor_gridlines: true,
            title: Some(ChartTitle {
                text: "Month".into(),
                ..Default::default()
            }),
            ..Default::default()
        }),
        series: vec![base_series()],
        ..Default::default()
    };
    let xml = chart_xml(&chart);
    assert_in_order(
        &xml,
        &[
            "<c:axPos",
            "<c:majorGridlines",
            "<c:minorGridlines",
            "<c:title>",
            "<c:numFmt",
        ],
    );
}

#[test]
fn val_axis_tail_follows_schema_order() {
    let chart = ChartData {
        val_axis: Some(ChartAxis {
            id: 2,
            cross_ax: 1,
            font_size: Some(1000),
            crosses_at: Some(5.0),
            cross_between: Some(CrossBetween::Between),
            major_unit: Some(2.0),
            minor_unit: Some(1.0),
            tick_lbl_pos: Some(TickLabelPosition::Low),
            ..Default::default()
        }),
        series: vec![base_series()],
        ..Default::default()
    };
    let xml = chart_xml(&chart);
    assert_in_order(
        &xml,
        &[
            "<c:tickLblPos",
            "<c:txPr>",
            "<c:crossAx",
            "<c:crossesAt",
            "<c:crossBetween",
            "<c:majorUnit",
            "<c:minorUnit",
        ],
    );
}

#[test]
fn cat_axis_never_writes_major_or_minor_unit() {
    let chart = ChartData {
        cat_axis: Some(ChartAxis {
            id: 1,
            cross_ax: 2,
            major_unit: Some(1.0),
            minor_unit: Some(1.0),
            ..Default::default()
        }),
        series: vec![base_series()],
        ..Default::default()
    };
    let xml = chart_xml(&chart);
    assert!(!xml.contains("<c:majorUnit"), "{xml}");
    assert!(!xml.contains("<c:minorUnit"), "{xml}");
}

#[test]
fn data_labels_num_fmt_comes_first() {
    let chart = ChartData {
        series: vec![ChartSeries {
            data_labels: Some(DataLabels {
                show_val: true,
                show_leader_lines: true,
                num_fmt: Some("0.0%".into()),
                ..Default::default()
            }),
            ..base_series()
        }],
        ..Default::default()
    };
    let xml = chart_xml(&chart);
    assert_in_order(
        &xml,
        &[
            "<c:dLbls><c:numFmt",
            "<c:showVal",
            "<c:showCatName",
            "<c:showSerName",
            "<c:showPercent",
            "<c:showLeaderLines",
        ],
    );
}

#[test]
fn trendline_disp_r_sqr_precedes_disp_eq() {
    let chart = ChartData {
        chart_type: ChartType::Line,
        series: vec![ChartSeries {
            trendline: Some(Trendline {
                trend_type: TrendlineType::Polynomial,
                order: Some(2),
                display_eq: true,
                display_r_sqr: true,
                ..Default::default()
            }),
            ..base_series()
        }],
        ..Default::default()
    };
    let xml = chart_xml(&chart);
    assert_in_order(
        &xml,
        &[
            "<c:trendlineType",
            "<c:order val=\"2\"",
            "<c:dispRSqr",
            "<c:dispEq",
        ],
    );
}

#[test]
fn chart_title_precedes_auto_title_deleted_and_view_3d() {
    let chart = ChartData {
        title: Some(ChartTitle {
            text: "T".into(),
            ..Default::default()
        }),
        view_3d: Some(View3D {
            rot_x: Some(15),
            ..Default::default()
        }),
        series: vec![base_series()],
        ..Default::default()
    };
    let xml = chart_xml(&chart);
    assert_in_order(
        &xml,
        &[
            "<c:title>",
            "<c:autoTitleDeleted val=\"0\"/>",
            "<c:view3D>",
            "<c:plotArea>",
        ],
    );
}

#[test]
fn auto_title_deleted_defaults_from_title_presence() {
    let untitled = ChartData {
        series: vec![base_series()],
        ..Default::default()
    };
    assert!(chart_xml(&untitled).contains("<c:autoTitleDeleted val=\"1\"/>"));

    let forced = ChartData {
        auto_title_deleted: Some(false),
        ..untitled.clone()
    };
    assert!(chart_xml(&forced).contains("<c:autoTitleDeleted val=\"0\"/>"));

    let json = serde_json::to_string(&untitled).unwrap();
    assert!(!json.contains("autoTitleDeleted"));
    let rt: ChartData = serde_json::from_str("{\"chartType\":\"bar\",\"series\":[]}").unwrap();
    assert_eq!(rt.auto_title_deleted, None);
}

#[test]
fn stacked_bars_write_overlap_100_before_axis_ids() {
    for (chart_type, grouping) in [
        (ChartType::Column, ChartGrouping::Stacked),
        (ChartType::Bar, ChartGrouping::PercentStacked),
    ] {
        let chart = ChartData {
            chart_type,
            grouping: Some(grouping),
            series: vec![base_series()],
            data_labels: Some(DataLabels::default()),
            cat_axis: Some(ChartAxis {
                id: 1,
                cross_ax: 2,
                ..Default::default()
            }),
            val_axis: Some(ChartAxis {
                id: 2,
                cross_ax: 1,
                ..Default::default()
            }),
            ..Default::default()
        };
        let xml = chart_xml(&chart);
        assert_in_order(
            &xml,
            &["</c:ser><c:dLbls>", "<c:overlap val=\"100\"/>", "<c:axId"],
        );
    }

    let clustered = ChartData {
        chart_type: ChartType::Column,
        grouping: Some(ChartGrouping::Clustered),
        series: vec![base_series()],
        ..Default::default()
    };
    assert!(!chart_xml(&clustered).contains("<c:overlap"));
}

#[test]
fn series_children_are_gated_by_chart_type() {
    let loaded = ChartSeries {
        marker: Some(MarkerStyle::Circle),
        explosion: Some(10),
        smooth: Some(true),
        trendline: Some(Trendline::default()),
        error_bars: Some(ErrorBars::default()),
        ..base_series()
    };
    let render = |chart_type| {
        chart_xml(&ChartData {
            chart_type,
            series: vec![loaded.clone()],
            ..Default::default()
        })
    };

    let bar = render(ChartType::Bar);
    for tag in ["<c:marker", "<c:explosion", "<c:smooth"] {
        assert!(!bar.contains(tag), "bar wrote {tag}: {bar}");
    }
    assert!(bar.contains("<c:trendline>") && bar.contains("<c:errBars>"));

    let pie = render(ChartType::Pie);
    assert!(pie.contains("<c:explosion val=\"10\"/>"));
    for tag in ["<c:marker", "<c:smooth", "<c:trendline", "<c:errBars"] {
        assert!(!pie.contains(tag), "pie wrote {tag}: {pie}");
    }
    let doughnut = render(ChartType::Doughnut);
    assert!(doughnut.contains("<c:explosion") && !doughnut.contains("<c:trendline"));

    let radar = render(ChartType::Radar);
    assert!(radar.contains("<c:marker>"));
    for tag in ["<c:explosion", "<c:smooth", "<c:trendline", "<c:errBars"] {
        assert!(!radar.contains(tag), "radar wrote {tag}: {radar}");
    }

    for chart_type in [ChartType::Line, ChartType::Stock, ChartType::Scatter] {
        let xml = render(chart_type);
        assert!(xml.contains("<c:marker>") && xml.contains("<c:smooth"));
        assert!(!xml.contains("<c:explosion"));
    }

    let area = render(ChartType::Area);
    for tag in ["<c:marker", "<c:explosion", "<c:smooth"] {
        assert!(!area.contains(tag), "area wrote {tag}: {area}");
    }
}

#[test]
fn generate_drawing_xml_rejects_mismatched_r_ids() {
    let wsc = |c| WorksheetChart {
        chart: ChartData::default(),
        anchor: ChartAnchor {
            to_col: c,
            ..Default::default()
        },
    };
    let err =
        ChartAnchor::generate_drawing_xml(&[wsc(1), wsc(2)], &["rId1".to_string()]).unwrap_err();
    assert!(
        matches!(err, crate::ModernXlsxError::XmlWrite(_)),
        "{err:?}"
    );
    assert!(
        ChartAnchor::generate_drawing_xml(&[wsc(1)], &["rId1".to_string(), "rId2".to_string()])
            .is_err()
    );
}

#[test]
fn title_without_font_overrides_writes_no_size_or_bold() {
    let chart = ChartData {
        title: Some(ChartTitle {
            text: "Plain".into(),
            ..Default::default()
        }),
        series: vec![base_series()],
        ..Default::default()
    };
    let xml = chart_xml(&chart);
    assert!(!xml.contains("sz="), "{xml}");
    // The page margins carry a legitimate `b=` attribute, so scope the check to the title.
    let title_xml = xml.split("</c:title>").next().unwrap();
    assert!(!title_xml.contains(" b="), "{xml}");

    let styled = ChartData {
        title: Some(ChartTitle {
            text: "Styled".into(),
            font_size: Some(1800),
            bold: Some(true),
            ..Default::default()
        }),
        ..chart
    };
    let xml = chart_xml(&styled);
    assert!(
        xml.contains("sz=\"1800\"") && xml.contains("b=\"1\""),
        "{xml}"
    );
}

#[test]
fn custom_error_bars_write_plus_and_minus_refs() {
    let chart = ChartData {
        chart_type: ChartType::Line,
        series: vec![ChartSeries {
            error_bars: Some(ErrorBars {
                err_type: ErrorBarType::Custom,
                direction: ErrorBarDirection::Both,
                plus_ref: Some("Sheet1!$C$2:$C$5".into()),
                minus_ref: Some("Sheet1!$D$2:$D$5".into()),
                value: Some(1.0),
                err_dir: None,
            }),
            ..base_series()
        }],
        ..Default::default()
    };
    let xml = chart_xml(&chart);
    assert_in_order(
        &xml,
        &[
            "<c:errBarType",
            "<c:errValType val=\"cust\"/>",
            "<c:plus><c:numRef><c:f>Sheet1!$C$2:$C$5</c:f></c:numRef></c:plus>",
            "<c:minus><c:numRef><c:f>Sheet1!$D$2:$D$5</c:f></c:numRef></c:minus>",
            "<c:val val=\"1\"/>",
            "</c:errBars>",
        ],
    );
}

#[test]
fn absolute_anchor_is_written_when_pos_is_set() {
    let charts = vec![WorksheetChart {
        chart: ChartData::default(),
        anchor: ChartAnchor {
            pos_x: Some(952_500),
            pos_y: Some(-1_000),
            ext_cx: Some(5_400_000),
            ext_cy: Some(3_240_000),
            ..Default::default()
        },
    }];
    let xml = ChartAnchor::generate_drawing_xml(&charts, &["rId1".to_string()]).unwrap();
    let xml = String::from_utf8(xml).unwrap();
    assert_in_order(
        &xml,
        &[
            "<xdr:absoluteAnchor>",
            "<xdr:pos x=\"952500\" y=\"-1000\"/>",
            "<xdr:ext cx=\"5400000\" cy=\"3240000\"/>",
            "<xdr:graphicFrame",
            "</xdr:absoluteAnchor>",
        ],
    );
    assert!(!xml.contains("xdr:from") && !xml.contains("oneCellAnchor"));
}

#[test]
fn user_strings_in_attributes_drop_illegal_characters() {
    let chart = ChartData {
        val_axis: Some(ChartAxis {
            id: 2,
            cross_ax: 1,
            num_fmt: Some("0\u{1}.0\"x\"".into()),
            ..Default::default()
        }),
        series: vec![ChartSeries {
            fill_color: Some("44\u{1}72C4".into()),
            ..base_series()
        }],
        ..Default::default()
    };
    let xml = chart_xml(&chart);
    assert!(xml.contains("val=\"4472C4\""), "{xml}");
    assert!(!xml.contains('\u{1}'), "{xml}");
    assert!(xml.contains("formatCode=\"0.0&quot;x&quot;\""), "{xml}");
    // Fails closed on non-finite doubles: the element is skipped, not written as NaN.
    let nan = ChartData {
        val_axis: Some(ChartAxis {
            id: 2,
            cross_ax: 1,
            min: Some(f64::NAN),
            max: Some(1e16),
            ..Default::default()
        }),
        series: vec![base_series()],
        ..Default::default()
    };
    let xml = chart_xml(&nan);
    assert!(!xml.contains("NaN") && !xml.contains("<c:min"), "{xml}");
    assert!(xml.contains("<c:max val=\"1e+16\"/>"), "{xml}");
}

#[test]
fn chart_xml_ct_order() {
    let full_axis = |id, cross_ax| ChartAxis {
        id,
        cross_ax,
        title: Some(ChartTitle {
            text: "Axis".into(),
            ..Default::default()
        }),
        num_fmt: Some("0".into()),
        min: Some(0.0),
        max: Some(9.0),
        log_base: Some(10.0),
        major_tick_mark: Some(TickMark::Out),
        minor_tick_mark: Some(TickMark::None),
        tick_lbl_pos: Some(TickLabelPosition::NextTo),
        major_gridlines: true,
        minor_gridlines: true,
        crosses_at: Some(1.0),
        cross_between: Some(CrossBetween::Between),
        major_unit: Some(2.0),
        minor_unit: Some(1.0),
        font_size: Some(900),
        ..Default::default()
    };
    let chart = ChartData {
        chart_type: ChartType::Column,
        grouping: Some(ChartGrouping::Stacked),
        title: Some(ChartTitle {
            text: "All".into(),
            ..Default::default()
        }),
        view_3d: Some(View3D {
            rot_x: Some(10),
            r_ang_ax: Some(true),
            perspective: Some(30),
            ..Default::default()
        }),
        plot_area_layout: Some(ManualLayout {
            x: 0.1,
            y: 0.1,
            w: 0.8,
            h: 0.8,
        }),
        series: vec![ChartSeries {
            fill_color: Some("112233".into()),
            data_labels: Some(DataLabels {
                show_val: true,
                num_fmt: Some("0".into()),
                ..Default::default()
            }),
            trendline: Some(Trendline::default()),
            error_bars: Some(ErrorBars::default()),
            ..base_series()
        }],
        data_labels: Some(DataLabels::default()),
        cat_axis: Some(full_axis(1, 2)),
        val_axis: Some(full_axis(2, 1)),
        show_data_table: true,
        legend: Some(ChartLegend::default()),
        ..Default::default()
    };
    let xml = chart_xml(&chart);
    // CT_Chart, CT_PlotArea, CT_BarChart, CT_BarSer.
    assert_in_order(
        &xml,
        &[
            "<c:chart>",
            "<c:title>",
            "<c:autoTitleDeleted",
            "<c:view3D><c:rotX",
            "<c:rAngAx",
            "<c:perspective",
            "<c:plotArea>",
            "<c:layout>",
            "<c:barChart>",
            "<c:barDir",
            "<c:grouping",
            "<c:ser><c:idx",
            "<c:order val=\"0\"",
            "<c:tx><c:strRef>",
            "<c:spPr><a:solidFill>",
            "<c:dLbls>",
            "<c:trendline>",
            "<c:errBars>",
            "<c:cat>",
            "<c:val>",
            "</c:ser><c:dLbls>",
            "<c:overlap",
            "<c:axId",
            "</c:barChart>",
            "<c:catAx>",
            "<c:valAx>",
            "<c:dTable>",
            "</c:plotArea>",
            "<c:legend>",
            "</c:chart>",
            "<c:printSettings>",
        ],
    );
    // EG_AxShared then the axis-type tail, inside the c:valAx element.
    let val_ax = &xml[xml.find("<c:valAx>").unwrap()..];
    assert_in_order(
        val_ax,
        &[
            "<c:axId",
            "<c:scaling>",
            "<c:logBase",
            "<c:orientation",
            "<c:max",
            "<c:min",
            "<c:delete",
            "<c:axPos",
            "<c:majorGridlines",
            "<c:minorGridlines",
            "<c:title>",
            "<c:numFmt",
            "<c:majorTickMark",
            "<c:minorTickMark",
            "<c:tickLblPos",
            "<c:txPr>",
            "<c:crossAx",
            "<c:crossesAt",
            "<c:crossBetween",
            "<c:majorUnit",
            "<c:minorUnit",
        ],
    );
}

// =======================================================================
// Parser fidelity: axis mapping, literals, entities, verbatim charts
// =======================================================================

/// A chart part whose `c:chartSpace` holds `body`.
fn chart_space(body: &str) -> String {
    format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:c14="http://schemas.microsoft.com/office/drawing/2007/8/2/chart">{body}</c:chartSpace>"#
    )
}

/// A chart part whose plot area holds `plot_area`.
fn plot_area_chart(plot_area: &str) -> String {
    chart_space(&format!(
        "<c:chart><c:plotArea><c:layout/>{plot_area}</c:plotArea></c:chart>"
    ))
}

/// An Excel-shaped axis: `extra` goes after `c:axPos`, `tail` after `c:crossAx`.
fn excel_axis(tag: &str, id: u32, cross: u32, delete: u8, pos: &str, extra: &str) -> String {
    format!(
        r#"<c:{tag}><c:axId val="{id}"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="{delete}"/><c:axPos val="{pos}"/>{extra}<c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="out"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:crossAx val="{cross}"/><c:crosses val="autoZero"/></c:{tag}>"#
    )
}

fn ser_xml(idx: u32) -> String {
    format!(
        r#"<c:ser><c:idx val="{idx}"/><c:order val="{idx}"/><c:val><c:numRef><c:f>Sheet1!$B${idx}:$B$9</c:f></c:numRef></c:val></c:ser>"#
    )
}

/// The first `<tag>...</tag>` element of `xml`.
fn element<'a>(xml: &'a str, tag: &str) -> &'a str {
    let start = xml
        .find(&format!("<{tag}>"))
        .unwrap_or_else(|| panic!("<{tag}> missing from {xml}"));
    let close = format!("</{tag}>");
    let end = start + xml[start..].find(&close).unwrap() + close.len();
    &xml[start..end]
}

/// The `val` of every `<c:axId>` in `xml`, in order.
fn ax_ids(xml: &str) -> Vec<u32> {
    xml.match_indices("<c:axId val=\"")
        .map(|(at, m)| {
            let rest = &xml[at + m.len()..];
            rest[..rest.find('"').unwrap()].parse().unwrap()
        })
        .collect()
}

/// The id of every `c:catAx` / `c:valAx` in `xml` (an axis's first child is its axId).
fn axis_ids(xml: &str, tag: &str) -> Vec<u32> {
    xml.match_indices(&format!("<{tag}>"))
        .map(|(at, _)| ax_ids(&xml[at..])[0])
        .collect()
}

/// Excel's secondary-axis combo: catAx 10 / valAx 20 for the bars, a deleted
/// catAx 30 and a right-hand valAx 40 for the line.
fn excel_combo_xml() -> String {
    plot_area_chart(&format!(
        r#"<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:varyColors val="0"/>{}<c:gapWidth val="150"/><c:axId val="10"/><c:axId val="20"/></c:barChart><c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>{}<c:marker val="1"/><c:axId val="30"/><c:axId val="40"/></c:lineChart>{}{}{}{}"#,
        ser_xml(0),
        ser_xml(1),
        excel_axis("catAx", 10, 20, 0, "b", ""),
        excel_axis("valAx", 20, 10, 0, "l", ""),
        excel_axis("valAx", 40, 30, 0, "r", ""),
        excel_axis("catAx", 30, 40, 1, "b", ""),
    ))
}

#[test]
fn excel_combo_maps_axes_by_ax_id() {
    let chart = ChartData::parse(excel_combo_xml().as_bytes()).unwrap();
    let cat = chart.cat_axis.as_ref().unwrap();
    assert_eq!((cat.id, cat.delete), (10, false));
    assert_eq!(chart.val_axis.as_ref().unwrap().id, 20);
    let sec_cat = chart.secondary_cat_axis.as_ref().unwrap();
    assert_eq!((sec_cat.id, sec_cat.delete), (30, true));
    assert_eq!(chart.secondary_val_axis.as_ref().unwrap().id, 40);
    assert_eq!(chart.raw_xml, None);

    let xml = chart_xml(&chart);
    let plot_area = element(&xml, "c:plotArea");
    let mut axes = axis_ids(plot_area, "c:catAx");
    assert_eq!(axes, [10, 30], "both catAx are written: {xml}");
    axes.extend(axis_ids(plot_area, "c:valAx"));
    for tag in ["c:barChart", "c:lineChart"] {
        let ids = ax_ids(element(&xml, tag));
        assert_eq!(ids.len(), 2, "{tag} needs exactly 2 axId: {xml}");
        for id in ids {
            assert!(axes.contains(&id), "{tag} axId {id} names no axis: {xml}");
        }
    }
    assert_eq!(ax_ids(element(&xml, "c:barChart")), [10, 20]);
    assert_eq!(ax_ids(element(&xml, "c:lineChart")), [30, 40]);
}

#[test]
fn combo_without_secondary_axes_shares_the_primary_ones() {
    let chart = ChartData {
        chart_type: ChartType::Column,
        series: vec![base_series()],
        cat_axis: Some(ChartAxis {
            id: 1,
            cross_ax: 2,
            ..Default::default()
        }),
        val_axis: Some(ChartAxis {
            id: 2,
            cross_ax: 1,
            ..Default::default()
        }),
        secondary_chart: Some(Box::new(ChartData {
            chart_type: ChartType::Line,
            series: vec![base_series()],
            ..Default::default()
        })),
        ..Default::default()
    };
    let xml = chart_xml(&chart);
    assert_eq!(ax_ids(element(&xml, "c:lineChart")), [1, 2], "{xml}");
}

#[test]
fn excel_scatter_keeps_both_value_axes() {
    let xml = plot_area_chart(&format!(
        r#"<c:scatterChart><c:scatterStyle val="lineMarker"/><c:varyColors val="0"/><c:ser><c:idx val="0"/><c:order val="0"/><c:xVal><c:numRef><c:f>Sheet1!$A$2:$A$5</c:f></c:numRef></c:xVal><c:yVal><c:numRef><c:f>Sheet1!$B$2:$B$5</c:f></c:numRef></c:yVal><c:smooth val="0"/></c:ser><c:axId val="1"/><c:axId val="2"/></c:scatterChart>{}{}"#,
        excel_axis("valAx", 1, 2, 0, "b", ""),
        excel_axis("valAx", 2, 1, 0, "l", ""),
    ));
    let chart = ChartData::parse(xml.as_bytes()).unwrap();
    assert_eq!(chart.cat_axis, None);
    assert_eq!(chart.val_axis.as_ref().unwrap().id, 1, "X axis");
    assert_eq!(chart.secondary_val_axis.as_ref().unwrap().id, 2, "Y axis");

    let out = chart_xml(&chart);
    assert_eq!(ax_ids(element(&out, "c:scatterChart")), [1, 2], "{out}");
    assert_eq!(out.matches("<c:valAx>").count(), 2, "{out}");
    assert!(!out.contains("<c:catAx"), "{out}");
    let x_axis = element(&out, "c:valAx");
    assert!(x_axis.contains("<c:axPos val=\"b\"/>"), "{x_axis}");
}

#[test]
fn scatter_cat_axis_is_written_as_the_x_value_axis() {
    let chart = ChartData {
        chart_type: ChartType::Scatter,
        series: vec![ChartSeries {
            x_val_ref: Some("Sheet1!$A$2:$A$5".into()),
            val_ref: "Sheet1!$B$2:$B$5".into(),
            ..Default::default()
        }],
        cat_axis: Some(ChartAxis {
            id: 0,
            cross_ax: 1,
            ..Default::default()
        }),
        val_axis: Some(ChartAxis {
            id: 1,
            cross_ax: 0,
            ..Default::default()
        }),
        ..Default::default()
    };
    let xml = chart_xml(&chart);
    assert_eq!(ax_ids(element(&xml, "c:scatterChart")), [0, 1], "{xml}");
    assert_eq!(axis_ids(&xml, "c:valAx"), [0, 1], "{xml}");
    assert!(!xml.contains("<c:catAx"), "{xml}");
    assert!(
        element(&xml, "c:valAx").contains("<c:axPos val=\"b\"/>"),
        "{xml}"
    );
}

#[test]
fn start_form_gridlines_are_read() {
    let grid = "<c:majorGridlines><c:spPr><a:ln/></c:spPr></c:majorGridlines><c:minorGridlines><c:spPr><a:ln w=\"9525\"/></c:spPr></c:minorGridlines>";
    let xml = plot_area_chart(&format!(
        r#"<c:barChart><c:barDir val="col"/>{}<c:axId val="7"/><c:axId val="8"/></c:barChart>{}{}"#,
        ser_xml(0),
        excel_axis("catAx", 7, 8, 0, "b", grid),
        excel_axis("valAx", 8, 7, 0, "l", grid),
    ));
    let chart = ChartData::parse(xml.as_bytes()).unwrap();
    for axis in [chart.cat_axis.unwrap(), chart.val_axis.unwrap()] {
        assert!(axis.major_gridlines && axis.minor_gridlines, "{axis:?}");
        assert_eq!(axis.cross_ax, if axis.id == 7 { 8 } else { 7 }, "{axis:?}");
        assert_eq!(axis.tick_lbl_pos, Some(TickLabelPosition::NextTo));
    }
}

#[test]
fn literal_series_name_is_written_as_c_v_and_read_back() {
    let chart = ChartData {
        series: vec![ChartSeries {
            name_literal: Some("Revenue".into()),
            val_ref: "Sheet1!$B$2:$B$5".into(),
            ..Default::default()
        }],
        ..Default::default()
    };
    let xml = chart_xml(&chart);
    assert!(xml.contains("<c:tx><c:v>Revenue</c:v></c:tx>"), "{xml}");
    assert!(!xml.contains("<c:strRef><c:f>Revenue"), "{xml}");
    let parsed = ChartData::parse(xml.as_bytes()).unwrap();
    assert_eq!(parsed.series[0].name_literal.as_deref(), Some("Revenue"));
    assert_eq!(parsed.series[0].name, None);

    // A formula name wins over a literal one.
    let both = ChartData {
        series: vec![ChartSeries {
            name: Some("Sheet1!$B$1".into()),
            ..chart.series[0].clone()
        }],
        ..Default::default()
    };
    let xml = chart_xml(&both);
    assert!(
        xml.contains("<c:tx><c:strRef><c:f>Sheet1!$B$1</c:f></c:strRef></c:tx>")
            && !xml.contains("<c:v>"),
        "{xml}"
    );
}

#[test]
fn str_ref_cache_is_not_a_literal_name() {
    let xml = plot_area_chart(
        r#"<c:lineChart><c:ser><c:idx val="0"/><c:order val="0"/><c:tx><c:strRef><c:f>Sheet1!$B$1</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>Revenue</c:v></c:pt></c:strCache></c:strRef></c:tx><c:cat><c:strRef><c:f>Sheet1!$A$2:$A$3</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>Q1</c:v></c:pt></c:strCache></c:strRef></c:cat><c:val><c:numRef><c:f>Sheet1!$B$2:$B$3</c:f></c:numRef></c:val></c:ser></c:lineChart>"#,
    );
    let ser = &ChartData::parse(xml.as_bytes()).unwrap().series[0];
    assert_eq!(ser.name.as_deref(), Some("Sheet1!$B$1"));
    assert_eq!(ser.name_literal, None);
    assert_eq!(ser.cat_ref.as_deref(), Some("Sheet1!$A$2:$A$3"));
}

#[test]
fn num_fmt_entities_are_decoded_once_and_escaped_once() {
    let fmt = r#"formatCode="&quot;$&quot;#,##0 &amp; x""#;
    let xml = plot_area_chart(&format!(
        r#"<c:barChart><c:barDir val="col"/><c:ser><c:idx val="0"/><c:order val="0"/><c:dLbls><c:numFmt {fmt} sourceLinked="0"/><c:showVal val="1"/></c:dLbls><c:val><c:numRef><c:f>Sheet1!$B$2:$B$5</c:f></c:numRef></c:val></c:ser><c:axId val="1"/><c:axId val="2"/></c:barChart>{}<c:valAx><c:axId val="2"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/><c:numFmt {fmt} sourceLinked="0"/><c:crossAx val="1"/></c:valAx>"#,
        excel_axis("catAx", 1, 2, 0, "b", ""),
    ));
    let chart = ChartData::parse(xml.as_bytes()).unwrap();
    let decoded = "\"$\"#,##0 & x";
    assert_eq!(
        chart.val_axis.as_ref().unwrap().num_fmt.as_deref(),
        Some(decoded)
    );
    assert_eq!(
        chart.series[0]
            .data_labels
            .as_ref()
            .unwrap()
            .num_fmt
            .as_deref(),
        Some(decoded)
    );
    let out = chart_xml(&chart);
    assert_eq!(out.matches(fmt).count(), 2, "{out}");
    assert!(
        !out.contains("&amp;quot;") && !out.contains("&amp;amp;"),
        "{out}"
    );
}

#[test]
fn ct_boolean_without_val_is_true() {
    let xml = plot_area_chart(&format!(
        r#"<c:lineChart><c:ser><c:idx val="0"/><c:order val="0"/><c:val><c:numRef><c:f>S!$B$2:$B$5</c:f></c:numRef></c:val><c:smooth/></c:ser><c:dLbls><c:showVal/><c:showPercent val="false"/></c:dLbls><c:axId val="1"/><c:axId val="2"/></c:lineChart><c:catAx><c:axId val="1"/><c:scaling/><c:delete/><c:axPos val="b"/><c:crossAx val="2"/></c:catAx>{}"#,
        excel_axis("valAx", 2, 1, 0, "l", ""),
    ));
    let chart = ChartData::parse(xml.as_bytes()).unwrap();
    assert!(chart.cat_axis.as_ref().unwrap().delete);
    assert!(!chart.val_axis.as_ref().unwrap().delete);
    assert_eq!(chart.series[0].smooth, Some(true));
    let dl = chart.data_labels.as_ref().unwrap();
    assert!(dl.show_val && !dl.show_percent, "{dl:?}");
}

#[test]
fn title_text_keeps_entities_and_spaces() {
    let rich = |text: &str| {
        format!(
            r#"<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr/></a:pPr><a:r><a:rPr lang="en-US"/><a:t>{text}</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title>"#
        )
    };
    let xml = chart_space(&format!(
        r#"<c:chart>{}<c:autoTitleDeleted val="0"/><c:plotArea><c:layout/><c:barChart><c:barDir val="col"/>{}<c:axId val="1"/><c:axId val="2"/></c:barChart><c:catAx><c:axId val="1"/><c:scaling/><c:delete val="0"/><c:axPos val="b"/>{}<c:crossAx val="2"/></c:catAx>{}</c:plotArea></c:chart>"#,
        rich("R&amp;D Sales"),
        ser_xml(0),
        rich(" R &amp; D &#x41; "),
        excel_axis("valAx", 2, 1, 0, "l", ""),
    ));
    let chart = ChartData::parse(xml.as_bytes()).unwrap();
    assert_eq!(chart.title.as_ref().unwrap().text, "R&D Sales");
    assert_eq!(
        chart
            .cat_axis
            .as_ref()
            .unwrap()
            .title
            .as_ref()
            .unwrap()
            .text,
        " R & D A "
    );
}

#[test]
fn auto_title_deleted_is_read_and_kept() {
    let xml = chart_space(&format!(
        r#"<c:chart><c:autoTitleDeleted val="0"/><c:plotArea><c:layout/><c:lineChart>{}</c:lineChart></c:plotArea></c:chart>"#,
        ser_xml(0)
    ));
    let chart = ChartData::parse(xml.as_bytes()).unwrap();
    assert_eq!(chart.auto_title_deleted, Some(false));
    assert!(
        chart_xml(&chart).contains("<c:autoTitleDeleted val=\"0\"/>"),
        "an untitled chart keeps Excel's automatic title"
    );
}

#[test]
fn c14_style_in_mc_choice_is_skipped() {
    let alt = |fallback: &str| {
        chart_space(&format!(
            r#"<mc:AlternateContent><mc:Choice Requires="c14"><c14:style val="102"/></mc:Choice>{fallback}</mc:AlternateContent><c:chart><c:plotArea><c:layout/><c:lineChart>{}</c:lineChart></c:plotArea></c:chart>"#,
            ser_xml(0)
        ))
    };
    let only_c14 = ChartData::parse(alt("").as_bytes()).unwrap();
    assert_eq!(only_c14.style_id, None);
    assert_eq!(only_c14.series.len(), 1);
    let with_fallback =
        ChartData::parse(alt(r#"<mc:Fallback><c:style val="2"/></mc:Fallback>"#).as_bytes())
            .unwrap();
    assert_eq!(with_fallback.style_id, Some(2));
}

#[test]
fn per_point_d_lbl_does_not_leak_into_series_d_lbls() {
    let xml = plot_area_chart(
        r#"<c:barChart><c:barDir val="col"/><c:ser><c:idx val="0"/><c:order val="0"/><c:dLbls><c:dLbl><c:idx val="2"/><c:numFmt formatCode="0.00" sourceLinked="0"/><c:showVal val="1"/><c:showCatName val="1"/></c:dLbl><c:showVal val="0"/><c:showCatName val="0"/></c:dLbls><c:val><c:numRef><c:f>S!$B$2:$B$5</c:f></c:numRef></c:val></c:ser><c:dLbls><c:dLbl><c:idx val="0"/><c:showPercent val="1"/></c:dLbl><c:showPercent val="0"/></c:dLbls></c:barChart>"#,
    );
    let chart = ChartData::parse(xml.as_bytes()).unwrap();
    let dl = chart.series[0].data_labels.as_ref().unwrap();
    assert!(!dl.show_val && !dl.show_cat_name, "{dl:?}");
    assert_eq!(dl.num_fmt, None);
    assert!(!chart.data_labels.as_ref().unwrap().show_percent);
}

#[test]
fn absolute_anchor_chart_is_read() {
    let xml = r#"<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><xdr:absoluteAnchor><xdr:pos x="12700" y="25400"/><xdr:ext cx="5000000" cy="3000000"/><xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="2" name="Chart 1"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="rId1"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:absoluteAnchor></xdr:wsDr>"#;
    let anchors = parse_drawing_anchors(xml.as_bytes()).unwrap();
    assert_eq!((anchors.total, anchors.charts.len()), (1, 1));
    let (anchor, r_id) = &anchors.charts[0];
    assert_eq!(r_id, "rId1");
    assert_eq!(
        *anchor,
        ChartAnchor {
            pos_x: Some(12_700),
            pos_y: Some(25_400),
            ext_cx: Some(5_000_000),
            ext_cy: Some(3_000_000),
            ..Default::default()
        }
    );
    // The writer emits the same anchor kind again.
    let written = ChartAnchor::generate_drawing_xml(
        &[WorksheetChart {
            chart: ChartData::default(),
            anchor: anchor.clone(),
        }],
        &["rId1".into()],
    )
    .unwrap();
    let reparsed = parse_drawing_anchors(&written).unwrap();
    assert_eq!(reparsed.charts[0].0, *anchor);
}

#[test]
fn malformed_numeric_attribute_is_an_error() {
    let xml = plot_area_chart(
        r#"<c:lineChart><c:ser><c:idx val="x"/><c:order val="0"/></c:ser></c:lineChart>"#,
    );
    let err = ChartData::parse(xml.as_bytes()).unwrap_err().to_string();
    assert!(err.contains("invalid value 'x'"), "{err}");
    // XML white space around a number is allowed.
    let xml = plot_area_chart(
        r#"<c:lineChart><c:ser><c:idx val=" 3 "/><c:order val="0"/></c:ser></c:lineChart>"#,
    );
    assert_eq!(ChartData::parse(xml.as_bytes()).unwrap().series[0].idx, 3);
}

#[test]
fn custom_error_bar_ranges_are_read() {
    let xml = plot_area_chart(
        r#"<c:lineChart><c:ser><c:idx val="0"/><c:order val="0"/><c:errBars><c:errBarType val="both"/><c:errValType val="cust"/><c:plus><c:numRef><c:f>Sheet1!$C$2:$C$5</c:f><c:numCache><c:ptCount val="1"/><c:pt idx="0"><c:v>1</c:v></c:pt></c:numCache></c:numRef></c:plus><c:minus><c:numRef><c:f>Sheet1!$D$2:$D$5</c:f></c:numRef></c:minus></c:errBars><c:val><c:numRef><c:f>Sheet1!$B$2:$B$5</c:f></c:numRef></c:val></c:ser></c:lineChart>"#,
    );
    let ser = &ChartData::parse(xml.as_bytes()).unwrap().series[0];
    let eb = ser.error_bars.as_ref().unwrap();
    assert_eq!(eb.err_type, ErrorBarType::Custom);
    assert_eq!(eb.plus_ref.as_deref(), Some("Sheet1!$C$2:$C$5"));
    assert_eq!(eb.minus_ref.as_deref(), Some("Sheet1!$D$2:$D$5"));
    assert_eq!(ser.val_ref, "Sheet1!$B$2:$B$5");
}

/// A 3-D bar chart, which the model does not represent.
fn bar_3d_xml() -> String {
    chart_space(
        r#"<c:chart><c:autoTitleDeleted val="1"/><c:view3D><c:rotX val="15"/><c:rotY val="20"/></c:view3D><c:plotArea><c:layout/><c:bar3DChart><c:barDir val="col"/><c:grouping val="clustered"/><c:varyColors val="0"/><c:ser><c:idx val="0"/><c:order val="0"/><c:tx><c:strRef><c:f>Sheet1!$B$1</c:f></c:strRef></c:tx><c:val><c:numRef><c:f>Sheet1!$B$2:$B$5</c:f></c:numRef></c:val></c:ser><c:gapWidth val="150"/><c:shape val="box"/><c:axId val="1"/><c:axId val="2"/><c:axId val="0"/></c:bar3DChart><c:catAx><c:axId val="1"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:crossAx val="2"/></c:catAx><c:valAx><c:axId val="2"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/><c:crossAx val="1"/></c:valAx></c:plotArea></c:chart>"#,
    )
}

#[test]
fn unsupported_chart_types_are_kept_verbatim() {
    let xml = bar_3d_xml();
    let chart = ChartData::parse(xml.as_bytes()).unwrap();
    assert_eq!(chart.raw_xml.as_deref(), Some(xml.as_str()));
    assert_eq!(
        chart.to_xml().unwrap(),
        xml.as_bytes(),
        "written byte-identical"
    );

    for tag in [
        "line3DChart",
        "pie3DChart",
        "area3DChart",
        "surfaceChart",
        "surface3DChart",
        "ofPieChart",
    ] {
        let xml = plot_area_chart(&format!("<c:{tag}>{}</c:{tag}>", ser_xml(0)));
        let chart = ChartData::parse(xml.as_bytes()).unwrap();
        assert_eq!(chart.raw_xml.as_deref(), Some(xml.as_str()), "{tag}");
        assert!(chart.series.is_empty(), "{tag} series are not modelled");
    }
}

#[test]
fn triple_combo_keeps_two_types_and_the_part_verbatim() {
    let xml = plot_area_chart(&format!(
        r#"<c:barChart><c:barDir val="col"/>{}<c:axId val="1"/><c:axId val="2"/></c:barChart><c:lineChart><c:grouping val="standard"/>{}<c:axId val="1"/><c:axId val="2"/></c:lineChart><c:areaChart><c:grouping val="stacked"/>{}<c:axId val="1"/><c:axId val="2"/></c:areaChart>{}{}"#,
        ser_xml(0),
        ser_xml(1),
        ser_xml(2),
        excel_axis("catAx", 1, 2, 0, "b", ""),
        excel_axis("valAx", 2, 1, 0, "l", ""),
    ));
    let chart = ChartData::parse(xml.as_bytes()).unwrap();
    assert_eq!(chart.chart_type, ChartType::Column);
    let sec = chart.secondary_chart.as_ref().unwrap();
    assert_eq!(sec.chart_type, ChartType::Line);
    let idx: Vec<u32> = sec.series.iter().map(|s| s.idx).collect();
    assert_eq!(idx, [1]);
    assert_eq!(chart.raw_xml.as_deref(), Some(xml.as_str()));
    assert_eq!(chart.to_xml().unwrap(), xml.as_bytes());
}

#[test]
fn date_axis_keeps_the_part_verbatim() {
    // Excel's line chart over dates: a dateAx X axis the model has no slot for.
    let xml = plot_area_chart(&format!(
        r#"<c:lineChart><c:grouping val="standard"/>{}<c:axId val="1"/><c:axId val="2"/></c:lineChart>{}{}"#,
        ser_xml(0),
        excel_axis("dateAx", 1, 2, 0, "b", ""),
        excel_axis("valAx", 2, 1, 0, "l", ""),
    ));
    let chart = ChartData::parse(xml.as_bytes()).unwrap();
    assert_eq!(chart.raw_xml.as_deref(), Some(xml.as_str()));
    assert_eq!(
        chart.to_xml().unwrap(),
        xml.as_bytes(),
        "written byte-identical"
    );
}

#[test]
fn nested_secondary_chart_is_an_error() {
    let chart = ChartData {
        secondary_chart: Some(Box::new(ChartData {
            chart_type: ChartType::Line,
            secondary_chart: Some(Box::new(ChartData {
                chart_type: ChartType::Area,
                ..Default::default()
            })),
            ..Default::default()
        })),
        ..Default::default()
    };
    let err = chart.to_xml().unwrap_err();
    assert!(
        matches!(err, crate::ModernXlsxError::XmlWrite(ref m) if m == "nested secondary charts are not supported; use at most one secondaryChart"),
        "{err:?}"
    );
}

#[test]
fn mismatched_chart_xml_is_an_error() {
    assert!(ChartData::parse(b"<c:chartSpace><c:chart></c:chartSpace>").is_err());
}

/// Writing then parsing `chart` gives `chart` back, field for field.
#[track_caller]
fn assert_model_roundtrip(chart: &ChartData) {
    let xml = chart.to_xml().unwrap();
    let parsed = ChartData::parse(&xml).unwrap();
    assert_eq!(&parsed, chart, "{}", String::from_utf8_lossy(&xml));
}

/// An axis in the form the parser returns (numFmt and axPos are always written).
fn written_axis(id: u32, cross_ax: u32, position: AxisPosition) -> ChartAxis {
    ChartAxis {
        id,
        cross_ax,
        num_fmt: Some("General".into()),
        position: Some(position),
        ..Default::default()
    }
}

#[test]
fn full_feature_chart_roundtrips_field_for_field() {
    let axis_title = |text: &str| ChartTitle {
        text: text.into(),
        ..Default::default()
    };
    let chart = ChartData {
        chart_type: ChartType::Line,
        title: Some(ChartTitle {
            text: "Full".into(),
            font_size: Some(1200),
            bold: Some(false),
            color: Some("1F3864".into()),
            ..Default::default()
        }),
        auto_title_deleted: Some(false),
        view_3d: Some(View3D {
            rot_x: Some(15),
            rot_y: Some(20),
            perspective: Some(30),
            r_ang_ax: Some(true),
        }),
        plot_area_layout: Some(ManualLayout {
            x: 0.125,
            y: 0.25,
            w: 0.5,
            h: 0.75,
        }),
        show_data_table: true,
        grouping: Some(ChartGrouping::Standard),
        style_id: Some(2),
        series: vec![ChartSeries {
            name: Some("Sheet1!$B$1".into()),
            cat_ref: Some("Sheet1!$A$2:$A$5".into()),
            val_ref: "Sheet1!$B$2:$B$5".into(),
            fill_color: Some("4472C4".into()),
            line_color: Some("ED7D31".into()),
            line_width: Some(12_700),
            marker: Some(MarkerStyle::Circle),
            smooth: Some(true),
            data_labels: Some(DataLabels {
                show_val: true,
                show_cat_name: true,
                show_ser_name: true,
                show_percent: true,
                num_fmt: Some("0.0".into()),
                show_leader_lines: true,
                show_legend_key: Some(true),
                show_bubble_size: Some(false),
            }),
            trendline: Some(Trendline {
                trend_type: TrendlineType::Polynomial,
                order: Some(3),
                period: Some(2),
                forward: Some(1.5),
                backward: Some(0.5),
                display_eq: true,
                display_r_sqr: true,
            }),
            error_bars: Some(ErrorBars {
                err_type: ErrorBarType::Custom,
                direction: ErrorBarDirection::Plus,
                value: Some(2.5),
                plus_ref: Some("Sheet1!$C$2:$C$5".into()),
                minus_ref: Some("Sheet1!$D$2:$D$5".into()),
                err_dir: Some(ErrorBarAxis::X),
            }),
            ..Default::default()
        }],
        data_labels: Some(DataLabels {
            show_val: true,
            ..Default::default()
        }),
        cat_axis: Some(ChartAxis {
            title: Some(axis_title("Quarter")),
            font_size: Some(900),
            major_tick_mark: Some(TickMark::Out),
            tick_lbl_pos: Some(TickLabelPosition::NextTo),
            ..written_axis(1, 2, AxisPosition::Bottom)
        }),
        val_axis: Some(ChartAxis {
            title: Some(axis_title("Amount")),
            font_size: Some(900),
            num_fmt: Some("#,##0".into()),
            source_linked: true,
            min: Some(0.0),
            max: Some(100.0),
            log_base: Some(10.0),
            reversed: true,
            major_gridlines: true,
            minor_gridlines: true,
            crosses_at: Some(5.0),
            major_unit: Some(20.0),
            minor_unit: Some(5.0),
            cross_between: Some(CrossBetween::MidCat),
            ..written_axis(2, 1, AxisPosition::Left)
        }),
        legend: Some(ChartLegend {
            position: LegendPosition::Bottom,
            overlay: true,
        }),
        ..Default::default()
    };
    assert_model_roundtrip(&chart);

    for direction in [
        ErrorBarDirection::Both,
        ErrorBarDirection::Plus,
        ErrorBarDirection::Minus,
    ] {
        let mut each = chart.clone();
        each.series[0].error_bars = Some(ErrorBars {
            err_type: ErrorBarType::FixedVal,
            direction,
            value: Some(1.0),
            ..Default::default()
        });
        assert_model_roundtrip(&each);
    }
}

#[test]
fn bubble_and_stock_charts_roundtrip_field_for_field() {
    let bubble = ChartData {
        chart_type: ChartType::Bubble,
        auto_title_deleted: Some(true),
        series: vec![ChartSeries {
            name_literal: Some("Size & weight".into()),
            x_val_ref: Some("Sheet1!$A$2:$A$5".into()),
            val_ref: "Sheet1!$B$2:$B$5".into(),
            bubble_size_ref: Some("Sheet1!$C$2:$C$5".into()),
            ..Default::default()
        }],
        val_axis: Some(written_axis(10, 20, AxisPosition::Bottom)),
        secondary_val_axis: Some(written_axis(20, 10, AxisPosition::Left)),
        ..Default::default()
    };
    assert_model_roundtrip(&bubble);

    let stock = ChartData {
        chart_type: ChartType::Stock,
        auto_title_deleted: Some(true),
        series: ["B", "C", "D"]
            .iter()
            .zip(0u32..)
            .map(|(col, i)| ChartSeries {
                idx: i,
                order: i,
                cat_ref: Some("Sheet1!$A$2:$A$5".into()),
                val_ref: format!("Sheet1!${col}$2:${col}$5"),
                ..Default::default()
            })
            .collect(),
        cat_axis: Some(written_axis(1, 2, AxisPosition::Bottom)),
        val_axis: Some(ChartAxis {
            font_size: Some(900),
            ..written_axis(2, 1, AxisPosition::Left)
        }),
        ..Default::default()
    };
    assert_model_roundtrip(&stock);
}

/// A line chart whose series carries an X error bar and data labels with a
/// legend key and bubble size, and whose valAx sets `c:crossBetween`.
fn fidelity_xml() -> String {
    let ser = r#"<c:ser><c:idx val="0"/><c:order val="0"/><c:dLbls><c:dLblPos val="r"/><c:showLegendKey val="1"/><c:showVal val="1"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="0"/><c:showBubbleSize/></c:dLbls><c:errBars><c:errDir val="x"/><c:errBarType val="both"/><c:errValType val="fixedVal"/><c:val val="2"/></c:errBars><c:val><c:numRef><c:f>Sheet1!$B$1:$B$9</c:f></c:numRef></c:val></c:ser>"#;
    plot_area_chart(&format!(
        r#"<c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>{ser}<c:axId val="1"/><c:axId val="2"/></c:lineChart>{}{}"#,
        excel_axis("catAx", 1, 2, 0, "b", ""),
        r#"<c:valAx><c:axId val="2"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/><c:numFmt formatCode="General" sourceLinked="1"/><c:crossAx val="1"/><c:crosses val="autoZero"/><c:crossBetween val="midCat"/><c:majorUnit val="2"/></c:valAx>"#,
    ))
}

#[test]
fn err_dir_legend_key_bubble_size_and_cross_between_survive_roundtrip() {
    let chart = ChartData::parse(fidelity_xml().as_bytes()).unwrap();
    let ser = &chart.series[0];
    assert_eq!(
        ser.error_bars.as_ref().unwrap().err_dir,
        Some(ErrorBarAxis::X)
    );
    let dl = ser.data_labels.as_ref().unwrap();
    assert_eq!(dl.show_legend_key, Some(true));
    assert_eq!(dl.show_bubble_size, Some(true));
    assert_eq!(
        chart.val_axis.as_ref().unwrap().cross_between,
        Some(CrossBetween::MidCat)
    );
    assert_eq!(chart.cat_axis.as_ref().unwrap().cross_between, None);

    let xml = String::from_utf8(chart.to_xml().unwrap()).unwrap();
    let eb = element(&xml, "c:errBars");
    assert!(
        eb.starts_with(r#"<c:errBars><c:errDir val="x"/><c:errBarType"#),
        "{eb}"
    );
    let dl = element(&xml, "c:dLbls");
    let order = [
        r#"<c:showLegendKey val="1"/>"#,
        r#"<c:showVal val="1"/>"#,
        r#"<c:showPercent val="0"/>"#,
        r#"<c:showBubbleSize val="1"/>"#,
        r"<c:showLeaderLines",
    ];
    let at: Vec<usize> = order
        .iter()
        .map(|t| {
            dl.find(t)
                .unwrap_or_else(|| panic!("{t} missing from {dl}"))
        })
        .collect();
    assert!(at.windows(2).all(|w| w[0] < w[1]), "{dl}");
    let val_ax = element(&xml[xml.find("<c:valAx>").unwrap()..], "c:valAx");
    assert!(
        val_ax.contains(r#"<c:crossAx val="1"/><c:crossBetween val="midCat"/><c:majorUnit"#),
        "{val_ax}"
    );
    assert_eq!(xml.matches("<c:crossBetween").count(), 1, "{xml}");

    // The writer always emits c:autoTitleDeleted.
    let mut chart = chart;
    chart.auto_title_deleted = Some(true);
    assert_model_roundtrip(&chart);
}

#[test]
fn new_fidelity_fields_serde_roundtrip_and_skip_when_unset() {
    let chart = ChartData::parse(fidelity_xml().as_bytes()).unwrap();
    let json = serde_json::to_string(&chart).unwrap();
    for key in [
        r#""errDir":"x""#,
        r#""showLegendKey":true"#,
        r#""showBubbleSize":true"#,
        r#""crossBetween":"midCat""#,
    ] {
        assert!(json.contains(key), "{key} missing from {json}");
    }
    assert_eq!(serde_json::from_str::<ChartData>(&json).unwrap(), chart);

    let plain = serde_json::to_string(&ChartData {
        series: vec![ChartSeries {
            data_labels: Some(DataLabels::default()),
            error_bars: Some(ErrorBars::default()),
            ..Default::default()
        }],
        val_axis: Some(ChartAxis::default()),
        ..Default::default()
    })
    .unwrap();
    for key in ["errDir", "showLegendKey", "showBubbleSize", "crossBetween"] {
        assert!(!plain.contains(key), "{key} serialised in {plain}");
    }
}

#[test]
fn excel_combo_roundtrips_field_for_field() {
    let mut chart = ChartData::parse(excel_combo_xml().as_bytes()).unwrap();
    // The fixture has no c:autoTitleDeleted; the writer always emits one.
    chart.auto_title_deleted = Some(true);
    assert_model_roundtrip(&chart);
}

// =======================================================================
// Parser hardening: axis-id cap, per-point overrides, grouped charts,
// percent hole sizes, colour modifiers
// =======================================================================

#[test]
fn many_ax_ids_and_axes_parse_in_bounded_time() {
    let ids = r#"<c:axId val="999999"/>"#.repeat(50_000);
    let axes = excel_axis("valAx", 0, 1, 0, "l", "").repeat(50_000);
    let xml = plot_area_chart(&format!(
        "<c:lineChart>{}{ids}</c:lineChart>{axes}",
        ser_xml(0)
    ));
    let start = std::time::Instant::now();
    let chart = ChartData::parse(xml.as_bytes()).unwrap();
    let elapsed = start.elapsed();
    assert_eq!(
        chart.raw_xml.as_deref(),
        Some(xml.as_str()),
        "the extra axes are unmodelled"
    );
    assert!(
        elapsed < std::time::Duration::from_secs(5),
        "parse took {elapsed:?}"
    );
}

#[test]
fn data_point_overrides_do_not_reach_the_series() {
    let xml = plot_area_chart(
        r#"<c:pieChart><c:varyColors val="1"/><c:ser><c:idx val="0"/><c:order val="0"/><c:spPr><a:solidFill><a:srgbClr val="4472C4"/></a:solidFill></c:spPr><c:dPt><c:idx val="3"/><c:bubble3D val="0"/><c:explosion val="20"/><c:spPr><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill><a:ln w="19050"/></c:spPr></c:dPt><c:val><c:numRef><c:f>Sheet1!$B$2:$B$5</c:f></c:numRef></c:val></c:ser><c:firstSliceAng val="0"/></c:pieChart>"#,
    );
    let chart = ChartData::parse(xml.as_bytes()).unwrap();
    let ser = &chart.series[0];
    assert_eq!(ser.idx, 0);
    assert_eq!(ser.fill_color.as_deref(), Some("4472C4"));
    assert_eq!(ser.explosion, None);
    assert_eq!(ser.line_width, None);
    assert_eq!(ser.val_ref, "Sheet1!$B$2:$B$5");
    assert_eq!(chart.raw_xml, None, "a dPt does not keep the part verbatim");
}

/// A drawing whose only anchor groups a chart with a rectangle.
const GROUPED_CHART_DRAWING: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><xdr:twoCellAnchor><xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>8</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>16</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:grpSp><xdr:nvGrpSpPr><xdr:cNvPr id="4" name="Group 3"/><xdr:cNvGrpSpPr/></xdr:nvGrpSpPr><xdr:grpSpPr><a:xfrm><a:off x="609600" y="190500"/><a:ext cx="4267200" cy="2857500"/><a:chOff x="609600" y="190500"/><a:chExt cx="4267200" cy="2857500"/></a:xfrm></xdr:grpSpPr><xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="2" name="Chart 1"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x="609600" y="190500"/><a:ext cx="4267200" cy="2286000"/></xdr:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="rId1"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:sp macro="" textlink=""><xdr:nvSpPr><xdr:cNvPr id="3" name="Rectangle 2"/><xdr:cNvSpPr/></xdr:nvSpPr><xdr:spPr><a:xfrm><a:off x="609600" y="2476500"/><a:ext cx="4267200" cy="571500"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:sp></xdr:grpSp><xdr:clientData/></xdr:twoCellAnchor></xdr:wsDr>"#;

#[test]
fn grouped_chart_is_not_a_chart_anchor() {
    let anchors = parse_drawing_anchors(GROUPED_CHART_DRAWING.as_bytes()).unwrap();
    assert!(anchors.charts.is_empty(), "{:?}", anchors.charts);
    assert_eq!(anchors.total, 1);
}

/// One sheet whose drawing (stored the way the reader preserves it) groups a
/// line chart with a rectangle.
fn grouped_chart_workbook() -> crate::WorkbookData {
    let chart = ChartData {
        chart_type: ChartType::Line,
        series: vec![ChartSeries {
            val_ref: "Sheet1!$A$1:$A$3".into(),
            ..Default::default()
        }],
        // CT_LineChart requires two c:axId, so the part opens in Excel.
        cat_axis: Some(written_axis(1, 2, AxisPosition::Bottom)),
        val_axis: Some(written_axis(2, 1, AxisPosition::Left)),
        ..Default::default()
    }
    .to_xml()
    .unwrap();
    let drawing_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/></Relationships>"#;
    let sheet_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/></Relationships>"#;
    let mut wb = crate::WorkbookData {
        sheets: vec![crate::SheetData {
            name: "Sheet1".into(),
            ..Default::default()
        }],
        ..Default::default()
    };
    for (path, data) in [
        ("xl/drawings/drawing1.xml", GROUPED_CHART_DRAWING.as_bytes()),
        (
            "xl/drawings/_rels/drawing1.xml.rels",
            drawing_rels.as_bytes(),
        ),
        ("xl/charts/chart1.xml", chart.as_slice()),
        ("xl/worksheets/_rels/sheet1.xml.rels", sheet_rels.as_bytes()),
    ] {
        wb.preserved_entries.insert(path.into(), data.to_vec());
    }
    wb
}

#[test]
fn grouped_chart_drawing_survives_a_read_write_round_trip() {
    use crate::zip::reader::{ZipSecurityLimits, read_zip_entries};

    let first = crate::writer::write_xlsx(&grouped_chart_workbook()).unwrap();
    let wb = crate::reader::read_xlsx(&first).unwrap();
    assert!(wb.sheets[0].worksheet.charts.is_empty());
    assert_eq!(
        wb.preserved_entries
            .get("xl/drawings/drawing1.xml")
            .map(Vec::as_slice),
        Some(GROUPED_CHART_DRAWING.as_bytes()),
        "the grouped drawing is kept verbatim"
    );
    let second = crate::writer::write_xlsx(&wb).unwrap();
    let entries = read_zip_entries(&second, &ZipSecurityLimits::default()).unwrap();
    assert_eq!(
        entries.get("xl/drawings/drawing1.xml").map(Vec::as_slice),
        Some(GROUPED_CHART_DRAWING.as_bytes()),
        "the group's shapes survive the write"
    );
}

#[test]
fn hole_size_accepts_percent_and_white_space() {
    let hole = |val: &str| {
        ChartData::parse(
            plot_area_chart(&format!(
                r#"<c:doughnutChart><c:varyColors val="1"/>{}<c:firstSliceAng val="0"/><c:holeSize val="{val}"/></c:doughnutChart>"#,
                ser_xml(0)
            ))
            .as_bytes(),
        )
    };
    assert_eq!(hole("75%").unwrap().hole_size, Some(75));
    assert_eq!(hole(" 60 ").unwrap().hole_size, Some(60));
    let bad = hole("x%");
    assert!(
        matches!(bad, Err(crate::ModernXlsxError::XmlParse(_))),
        "{bad:?}"
    );
}

#[test]
fn colours_with_modifiers_are_read() {
    let tinted =
        r#"<a:solidFill><a:srgbClr val="4472C4"><a:alpha val="50000"/></a:srgbClr></a:solidFill>"#;
    let cat_title = format!(
        r#"<c:title><c:tx><c:rich><a:bodyPr/><a:p><a:pPr><a:defRPr b="1">{tinted}</a:defRPr></a:pPr><a:r><a:t>Month</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title>"#
    );
    let xml = chart_space(&format!(
        r#"<c:chart><c:title><c:tx><c:rich><a:bodyPr/><a:p><a:r><a:rPr lang="en-US" sz="1400">{tinted}</a:rPr><a:t>Sales</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title><c:plotArea><c:layout/><c:lineChart><c:grouping val="standard"/><c:ser><c:idx val="0"/><c:order val="0"/><c:spPr>{tinted}<a:ln w="28575">{tinted}</a:ln></c:spPr><c:val><c:numRef><c:f>Sheet1!$B$2:$B$5</c:f></c:numRef></c:val></c:ser><c:axId val="1"/><c:axId val="2"/></c:lineChart>{}{}</c:plotArea></c:chart>"#,
        excel_axis("catAx", 1, 2, 0, "b", &cat_title),
        excel_axis("valAx", 2, 1, 0, "l", ""),
    ));
    let chart = ChartData::parse(xml.as_bytes()).unwrap();
    let ser = &chart.series[0];
    assert_eq!(ser.fill_color.as_deref(), Some("4472C4"), "series fill");
    assert_eq!(ser.line_color.as_deref(), Some("4472C4"), "series line");
    assert_eq!(ser.line_width, Some(28_575));
    let title = chart.title.as_ref().unwrap();
    assert_eq!(title.text, "Sales");
    assert_eq!(title.color.as_deref(), Some("4472C4"), "title rPr fill");
    let axis_title = chart.cat_axis.as_ref().unwrap().title.as_ref().unwrap();
    assert_eq!(axis_title.text, "Month");
    assert_eq!(
        axis_title.color.as_deref(),
        Some("4472C4"),
        "axis title defRPr fill"
    );
}

/// A bar chart with one series whose `c:tx` and `c:val` children are `tx` and `val`.
fn one_series_chart(tx: &str, val: &str) -> ChartData {
    let xml = format!(
        r#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:chart><c:plotArea><c:barChart><c:barDir val="col"/><c:ser><c:idx val="0"/><c:order val="0"/><c:tx>{tx}</c:tx><c:val><c:numRef>{val}</c:numRef></c:val></c:ser><c:axId val="1"/><c:axId val="2"/></c:barChart></c:plotArea></c:chart></c:chartSpace>"#
    );
    ChartData::parse(xml.as_bytes()).unwrap()
}

#[test]
fn chart_text_line_ends_are_normalized_and_char_refs_kept() {
    let chart = one_series_chart("<c:v>Line1\r\nLine2</c:v>", "<c:f>Sheet1!$B$2</c:f>");
    assert_eq!(
        chart.series[0].name_literal.as_deref(),
        Some("Line1\nLine2")
    );
    let chart = one_series_chart("<c:v>a&#13;&#10;b</c:v>", "<c:f>Sheet1!$B$2</c:f>");
    assert_eq!(chart.series[0].name_literal.as_deref(), Some("a\r\nb"));
}

#[test]
fn chart_cdata_formulas_are_read() {
    let chart = one_series_chart(
        "<c:strRef><c:f><![CDATA[Sheet1!$B$1]]></c:f></c:strRef>",
        "<c:f><![CDATA[Sheet1!$A$1]]></c:f>",
    );
    assert_eq!(chart.series[0].val_ref, "Sheet1!$A$1");
    assert_eq!(chart.series[0].name.as_deref(), Some("Sheet1!$B$1"));
}

/// Every chart type whose chart-type element requires two `c:axId`.
const AXIS_BEARING: [ChartType; 8] = [
    ChartType::Bar,
    ChartType::Column,
    ChartType::Line,
    ChartType::Scatter,
    ChartType::Area,
    ChartType::Radar,
    ChartType::Bubble,
    ChartType::Stock,
];

/// The chart-type element tag of `chart_type`.
fn chart_type_tag(chart_type: ChartType) -> &'static str {
    match chart_type {
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

/// A one-series chart of `chart_type` with no axis set.
fn chart_without_axes(chart_type: ChartType) -> ChartData {
    let xy = matches!(chart_type, ChartType::Scatter | ChartType::Bubble);
    ChartData {
        chart_type,
        series: vec![ChartSeries {
            cat_ref: (!xy).then(|| "Sheet1!$A$2:$A$5".into()),
            x_val_ref: xy.then(|| "Sheet1!$A$2:$A$5".into()),
            bubble_size_ref: (chart_type == ChartType::Bubble).then(|| "Sheet1!$C$2:$C$5".into()),
            val_ref: "Sheet1!$B$2:$B$5".into(),
            ..Default::default()
        }],
        ..Default::default()
    }
}

/// The `c:crossAx` of the axis whose id is `id` in `xml`.
fn cross_ax_of(xml: &str, id: u32) -> u32 {
    for tag in ["c:catAx", "c:valAx"] {
        for (at, _) in xml.match_indices(&format!("<{tag}>")) {
            let axis = element(&xml[at..], tag);
            if ax_ids(axis)[0] == id {
                let rest = &axis[axis.find("<c:crossAx val=\"").unwrap() + 16..];
                return rest[..rest.find('"').unwrap()].parse().unwrap();
            }
        }
    }
    panic!("no axis {id} in {xml}");
}

#[test]
fn every_axis_bearing_chart_without_axes_gets_two_crossing_default_axes() {
    for chart_type in AXIS_BEARING {
        let xml = chart_xml(&chart_without_axes(chart_type));
        let tag = chart_type_tag(chart_type);
        let ids = ax_ids(element(&xml, tag));
        // The ids the TypeScript ChartBuilder gives its default axes.
        assert_eq!(ids, [0, 1], "{tag} needs exactly 2 axId: {xml}");
        let plot_area = element(&xml, "c:plotArea");
        let (cat, val) = (
            axis_ids(plot_area, "c:catAx"),
            axis_ids(plot_area, "c:valAx"),
        );
        if matches!(chart_type, ChartType::Scatter | ChartType::Bubble) {
            assert_eq!((cat, val), (vec![], vec![0, 1]), "{xml}");
        } else {
            assert_eq!((cat, val), (vec![0], vec![1]), "{xml}");
        }
        assert_eq!(cross_ax_of(plot_area, 0), 1, "{xml}");
        assert_eq!(cross_ax_of(plot_area, 1), 0, "{xml}");
        // Bottom first axis, left second axis, as the builder's defaults are written.
        assert_in_order(
            plot_area,
            &[
                "<c:axId val=\"0\"/><c:scaling>",
                "<c:axPos val=\"b\"/>",
                "<c:axId val=\"1\"/><c:scaling>",
                "<c:axPos val=\"l\"/>",
            ],
        );
    }
}

#[test]
fn pie_and_doughnut_without_axes_stay_axis_free() {
    for chart_type in [ChartType::Pie, ChartType::Doughnut] {
        let xml = chart_xml(&chart_without_axes(chart_type));
        assert!(!xml.contains("<c:axId"), "{xml}");
        assert!(
            !xml.contains("<c:catAx") && !xml.contains("<c:valAx"),
            "{xml}"
        );
    }
}

#[test]
fn combo_without_axes_draws_both_elements_on_the_default_pair() {
    for secondary_type in AXIS_BEARING {
        let chart = ChartData {
            secondary_chart: Some(Box::new(chart_without_axes(secondary_type))),
            ..chart_without_axes(ChartType::Column)
        };
        let xml = chart_xml(&chart);
        let plot_area = element(&xml, "c:plotArea");
        let mut axes = axis_ids(plot_area, "c:catAx");
        axes.extend(axis_ids(plot_area, "c:valAx"));
        axes.sort_unstable();
        assert_eq!(axes, [0, 1], "one default pair: {xml}");
        // The secondary element comes after the primary one; with a Column
        // secondary both are c:barChart.
        let second = &plot_area[plot_area.find("</c:barChart>").unwrap()..];
        let secondary = element(second, chart_type_tag(secondary_type));
        assert_eq!(ax_ids(secondary), [0, 1], "{xml}");
        assert_eq!(ax_ids(element(&xml, "c:barChart")), [0, 1], "{xml}");
    }
}

#[test]
fn combo_with_only_secondary_axes_draws_its_primary_element_on_defaults() {
    let chart = ChartData {
        secondary_chart: Some(Box::new(chart_without_axes(ChartType::Line))),
        secondary_cat_axis: Some(ChartAxis {
            id: 0,
            cross_ax: 1,
            delete: true,
            ..Default::default()
        }),
        secondary_val_axis: Some(ChartAxis {
            id: 1,
            cross_ax: 0,
            position: Some(AxisPosition::Right),
            ..Default::default()
        }),
        ..chart_without_axes(ChartType::Column)
    };
    let xml = chart_xml(&chart);
    // The defaults take the smallest ids the secondary axes leave free.
    assert_eq!(ax_ids(element(&xml, "c:barChart")), [2, 3], "{xml}");
    assert_eq!(ax_ids(element(&xml, "c:lineChart")), [0, 1], "{xml}");
    let plot_area = element(&xml, "c:plotArea");
    assert_eq!(axis_ids(plot_area, "c:catAx"), [0, 2], "{xml}");
    assert_eq!(axis_ids(plot_area, "c:valAx"), [1, 3], "{xml}");
}

#[test]
fn axis_bearing_secondary_of_a_pie_gets_a_default_pair() {
    let chart = ChartData {
        secondary_chart: Some(Box::new(chart_without_axes(ChartType::Line))),
        ..chart_without_axes(ChartType::Pie)
    };
    let xml = chart_xml(&chart);
    assert!(!element(&xml, "c:pieChart").contains("<c:axId"), "{xml}");
    assert_eq!(ax_ids(element(&xml, "c:lineChart")), [0, 1], "{xml}");
    let plot_area = element(&xml, "c:plotArea");
    assert_eq!(axis_ids(plot_area, "c:catAx"), [0], "{xml}");
    assert_eq!(axis_ids(plot_area, "c:valAx"), [1], "{xml}");
}

#[test]
fn scatter_secondary_of_a_doughnut_gets_two_default_value_axes() {
    let chart = ChartData {
        secondary_chart: Some(Box::new(chart_without_axes(ChartType::Scatter))),
        secondary_val_axis: Some(written_axis(4, 0, AxisPosition::Left)),
        ..chart_without_axes(ChartType::Doughnut)
    };
    let xml = chart_xml(&chart);
    assert!(
        !element(&xml, "c:doughnutChart").contains("<c:axId"),
        "{xml}"
    );
    // The unset X slot takes the smallest free id and crosses the set Y axis.
    assert_eq!(ax_ids(element(&xml, "c:scatterChart")), [0, 4], "{xml}");
    let plot_area = element(&xml, "c:plotArea");
    assert_eq!(axis_ids(plot_area, "c:catAx"), Vec::<u32>::new(), "{xml}");
    assert_eq!(axis_ids(plot_area, "c:valAx"), [4, 0], "{xml}");
    assert_eq!(cross_ax_of(plot_area, 0), 4, "{xml}");
}

#[test]
fn scatter_combo_with_only_an_x_axis_shares_the_default_y_axis() {
    let chart = ChartData {
        val_axis: Some(written_axis(5, 6, AxisPosition::Bottom)),
        secondary_chart: Some(Box::new(chart_without_axes(ChartType::Scatter))),
        ..chart_without_axes(ChartType::Scatter)
    };
    let xml = chart_xml(&chart);
    let plot_area = element(&xml, "c:plotArea");
    let primary = element(plot_area, "c:scatterChart");
    assert_eq!(ax_ids(primary), [5, 0], "{xml}");
    let secondary = element(
        &plot_area[plot_area.find("</c:scatterChart>").unwrap()..],
        "c:scatterChart",
    );
    assert_eq!(ax_ids(secondary), [5, 0], "{xml}");
    assert_eq!(axis_ids(plot_area, "c:valAx"), [5, 0], "{xml}");
}

/// Charts that set every axis they are drawn on, each with its part as written
/// before default axes were added for unset slots.
fn charts_with_explicit_axes() -> Vec<(ChartData, &'static str)> {
    let series = || {
        vec![ChartSeries {
            val_ref: "Sheet1!$B$2:$B$5".into(),
            ..Default::default()
        }]
    };
    let axis = |id, cross_ax| ChartAxis {
        id,
        cross_ax,
        ..Default::default()
    };
    vec![
        (
            ChartData {
                chart_type: ChartType::Line,
                series: series(),
                cat_axis: Some(axis(7, 8)),
                val_axis: Some(axis(8, 7)),
                ..Default::default()
            },
            GOLDEN_LINE,
        ),
        (
            ChartData {
                chart_type: ChartType::Column,
                series: series(),
                cat_axis: Some(axis(1, 2)),
                val_axis: Some(axis(2, 1)),
                secondary_chart: Some(Box::new(ChartData {
                    chart_type: ChartType::Line,
                    series: series(),
                    ..Default::default()
                })),
                secondary_cat_axis: Some(ChartAxis {
                    delete: true,
                    ..axis(3, 4)
                }),
                secondary_val_axis: Some(ChartAxis {
                    position: Some(AxisPosition::Right),
                    ..axis(4, 3)
                }),
                ..Default::default()
            },
            GOLDEN_COMBO,
        ),
        (
            ChartData {
                chart_type: ChartType::Column,
                series: series(),
                cat_axis: Some(axis(1, 2)),
                val_axis: Some(axis(2, 1)),
                secondary_chart: Some(Box::new(ChartData {
                    chart_type: ChartType::Line,
                    series: series(),
                    ..Default::default()
                })),
                ..Default::default()
            },
            GOLDEN_COMBO_SHARED,
        ),
        (
            ChartData {
                chart_type: ChartType::Scatter,
                series: series(),
                val_axis: Some(axis(1, 2)),
                secondary_val_axis: Some(axis(2, 1)),
                ..Default::default()
            },
            GOLDEN_SCATTER,
        ),
        (
            ChartData {
                chart_type: ChartType::Pie,
                series: series(),
                ..Default::default()
            },
            GOLDEN_PIE,
        ),
    ]
}

const GOLDEN_LINE: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><c:chart><c:autoTitleDeleted val="1"/><c:plotArea><c:layout/><c:lineChart><c:ser><c:idx val="0"/><c:order val="0"/><c:val><c:numRef><c:f>Sheet1!$B$2:$B$5</c:f></c:numRef></c:val></c:ser><c:axId val="7"/><c:axId val="8"/></c:lineChart><c:catAx><c:axId val="7"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:numFmt formatCode="General" sourceLinked="0"/><c:crossAx val="8"/></c:catAx><c:valAx><c:axId val="8"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/><c:numFmt formatCode="General" sourceLinked="0"/><c:crossAx val="7"/></c:valAx></c:plotArea></c:chart><c:printSettings><c:headerFooter/><c:pageMargins b="0.75" l="0.7" r="0.7" t="0.75" header="0.3" footer="0.3"/><c:pageSetup/></c:printSettings></c:chartSpace>"#;
const GOLDEN_COMBO: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><c:chart><c:autoTitleDeleted val="1"/><c:plotArea><c:layout/><c:barChart><c:barDir val="col"/><c:ser><c:idx val="0"/><c:order val="0"/><c:val><c:numRef><c:f>Sheet1!$B$2:$B$5</c:f></c:numRef></c:val></c:ser><c:axId val="1"/><c:axId val="2"/></c:barChart><c:lineChart><c:ser><c:idx val="0"/><c:order val="0"/><c:val><c:numRef><c:f>Sheet1!$B$2:$B$5</c:f></c:numRef></c:val></c:ser><c:axId val="3"/><c:axId val="4"/></c:lineChart><c:catAx><c:axId val="1"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:numFmt formatCode="General" sourceLinked="0"/><c:crossAx val="2"/></c:catAx><c:valAx><c:axId val="2"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/><c:numFmt formatCode="General" sourceLinked="0"/><c:crossAx val="1"/></c:valAx><c:valAx><c:axId val="4"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="r"/><c:numFmt formatCode="General" sourceLinked="0"/><c:crossAx val="3"/></c:valAx><c:catAx><c:axId val="3"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="1"/><c:axPos val="b"/><c:numFmt formatCode="General" sourceLinked="0"/><c:crossAx val="4"/></c:catAx></c:plotArea></c:chart><c:printSettings><c:headerFooter/><c:pageMargins b="0.75" l="0.7" r="0.7" t="0.75" header="0.3" footer="0.3"/><c:pageSetup/></c:printSettings></c:chartSpace>"#;
const GOLDEN_COMBO_SHARED: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><c:chart><c:autoTitleDeleted val="1"/><c:plotArea><c:layout/><c:barChart><c:barDir val="col"/><c:ser><c:idx val="0"/><c:order val="0"/><c:val><c:numRef><c:f>Sheet1!$B$2:$B$5</c:f></c:numRef></c:val></c:ser><c:axId val="1"/><c:axId val="2"/></c:barChart><c:lineChart><c:ser><c:idx val="0"/><c:order val="0"/><c:val><c:numRef><c:f>Sheet1!$B$2:$B$5</c:f></c:numRef></c:val></c:ser><c:axId val="1"/><c:axId val="2"/></c:lineChart><c:catAx><c:axId val="1"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:numFmt formatCode="General" sourceLinked="0"/><c:crossAx val="2"/></c:catAx><c:valAx><c:axId val="2"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/><c:numFmt formatCode="General" sourceLinked="0"/><c:crossAx val="1"/></c:valAx></c:plotArea></c:chart><c:printSettings><c:headerFooter/><c:pageMargins b="0.75" l="0.7" r="0.7" t="0.75" header="0.3" footer="0.3"/><c:pageSetup/></c:printSettings></c:chartSpace>"#;
const GOLDEN_SCATTER: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><c:chart><c:autoTitleDeleted val="1"/><c:plotArea><c:layout/><c:scatterChart><c:scatterStyle val="lineMarker"/><c:ser><c:idx val="0"/><c:order val="0"/><c:yVal><c:numRef><c:f>Sheet1!$B$2:$B$5</c:f></c:numRef></c:yVal></c:ser><c:axId val="1"/><c:axId val="2"/></c:scatterChart><c:valAx><c:axId val="1"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:numFmt formatCode="General" sourceLinked="0"/><c:crossAx val="2"/></c:valAx><c:valAx><c:axId val="2"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/><c:numFmt formatCode="General" sourceLinked="0"/><c:crossAx val="1"/></c:valAx></c:plotArea></c:chart><c:printSettings><c:headerFooter/><c:pageMargins b="0.75" l="0.7" r="0.7" t="0.75" header="0.3" footer="0.3"/><c:pageSetup/></c:printSettings></c:chartSpace>"#;
const GOLDEN_PIE: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><c:chart><c:autoTitleDeleted val="1"/><c:plotArea><c:layout/><c:pieChart><c:ser><c:idx val="0"/><c:order val="0"/><c:val><c:numRef><c:f>Sheet1!$B$2:$B$5</c:f></c:numRef></c:val></c:ser></c:pieChart></c:plotArea></c:chart><c:printSettings><c:headerFooter/><c:pageMargins b="0.75" l="0.7" r="0.7" t="0.75" header="0.3" footer="0.3"/><c:pageSetup/></c:printSettings></c:chartSpace>"#;

#[test]
fn charts_with_explicit_axes_are_written_as_before() {
    for (chart, golden) in charts_with_explicit_axes() {
        assert_eq!(chart_xml(&chart), golden);
    }
}
