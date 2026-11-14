//! Golden tests for chart parts. Each serialises only the subtree under test
//! (see `crate::common::golden_subtree`).

use modern_xlsx_core::ooxml::charts::{
    AxisPosition, ChartAnchor, ChartAxis, ChartData, ChartGrouping, ChartLegend, ChartSeries,
    ChartTitle, ChartType, LegendPosition, WorksheetChart,
};
use modern_xlsx_core::ooxml::worksheet::{Row, WorksheetXml};

use crate::common::{golden_subtree, minimal_workbook, number_row, roundtrip, text_row};

/// A clustered column chart with a chart title and both axis titles survives
/// write -> read.
#[test]
fn golden_chart_bar_axis_titles() {
    let title = |text: &str| ChartTitle {
        text: text.into(),
        ..Default::default()
    };
    let chart = ChartData {
        chart_type: ChartType::Column,
        title: Some(ChartTitle {
            font_size: Some(1400),
            bold: Some(true),
            ..title("Sales & Returns")
        }),
        series: vec![ChartSeries {
            name: Some("Sheet1!$B$1".into()),
            cat_ref: Some("Sheet1!$A$2:$A$3".into()),
            val_ref: "Sheet1!$B$2:$B$3".into(),
            fill_color: Some("4472C4".into()),
            ..Default::default()
        }],
        cat_axis: Some(ChartAxis {
            id: 1,
            cross_ax: 2,
            title: Some(title("Quarter")),
            position: Some(AxisPosition::Bottom),
            ..Default::default()
        }),
        val_axis: Some(ChartAxis {
            id: 2,
            cross_ax: 1,
            title: Some(title("Amount (USD)")),
            num_fmt: Some("#,##0".into()),
            major_gridlines: true,
            position: Some(AxisPosition::Left),
            ..Default::default()
        }),
        legend: Some(ChartLegend {
            position: LegendPosition::Bottom,
            overlay: false,
        }),
        grouping: Some(ChartGrouping::Clustered),
        bar_dir_horizontal: Some(false),
        ..Default::default()
    };
    let rows: Vec<Row> = vec![
        text_row(1, "B1", "Sales"),
        text_row(2, "A2", "Q1"),
        number_row(3, "B3", "20"),
    ];
    let ws = WorksheetXml {
        rows,
        charts: vec![WorksheetChart {
            chart,
            anchor: ChartAnchor {
                from_col: 3,
                from_row: 1,
                to_col: 10,
                to_row: 16,
                ..Default::default()
            },
        }],
        ..Default::default()
    };
    let wb = minimal_workbook("Sheet1", ws);

    let wb2 = roundtrip(&wb);
    golden_subtree(&wb2.sheets[0].worksheet.charts, "chart_bar_axis_titles");
}
