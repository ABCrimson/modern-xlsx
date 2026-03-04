# 0.8.x — Charts & Visualizations Implementation Plan

> **For Claude:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task.

**Goal:** Add full chart creation, reading, and roundtrip support — enabling modern-xlsx to programmatically generate Excel charts (bar, line, pie, scatter, area, and more) with titles, legends, data labels, trendlines, and combo charts.

**Architecture:** Chart data is modeled as Rust structs (`ChartData`, `ChartSeries`, `ChartAxis`, etc.) with serde JSON serialization, crossing the WASM boundary like all other data. Chart XML follows ECMA-376 DrawingML Chart (`c:` namespace) and Spreadsheet Drawing (`xdr:` namespace). The Rust writer generates `xl/charts/chart{n}.xml` and `xl/drawings/drawing{n}.xml` with proper content types and relationships. Charts transition from opaque `preserved_entries` passthrough to structured typed data, following the same pattern used for tables (`ooxml/tables.rs`). TypeScript provides a fluent `ChartBuilder` API. Reader replaces chart blobs with parsed `ChartData` structs while preserving unknown chart extensions.

**Tech Stack:** Rust (quick-xml SAX, serde), TypeScript (Vitest), OOXML DrawingML Chart spec (ECMA-376 Part 1 §21.2)

---

## Version Map

| Version | Feature | Scope |
|---------|---------|-------|
| 0.8.0 | Chart Data Model | Rust types + serde + TS exports |
| 0.8.1 | Chart Axis, Legend & Formatting Types | Rust types + serde + TS exports |
| 0.8.2 | Chart XML Writer (Bar, Column, Line) | Rust writer + tests |
| 0.8.3 | Chart XML Writer (Pie, Scatter, Area) | Rust writer + tests |
| 0.8.4 | Chart Drawing Anchors & Packaging | Rust writer + content types + rels |
| 0.8.5 | Chart Titles, Labels & Legend Writer | Rust writer + tests |
| 0.8.6 | TypeScript Chart Creation API | TS ChartBuilder + tests |
| 0.8.7 | Chart XML Reader | Rust parser + TS getter + tests |
| 0.8.8 | Chart Roundtrip & Style Presets | Roundtrip tests + style system |
| 0.8.9 | Advanced: Combo, Trendlines, Error Bars | Rust + TS + tests |

---

## 0.8.0 — Chart Data Model

### Task 1: Define core chart Rust types

**Files:**
- Create: `crates/modern-xlsx-core/src/ooxml/charts.rs`
- Modify: `crates/modern-xlsx-core/src/ooxml/mod.rs` — add `pub mod charts;`
- Modify: `crates/modern-xlsx-core/src/lib.rs` — add chart imports to `WorksheetXml`

**Step 1: Create `charts.rs` with core types**

```rust
// crates/modern-xlsx-core/src/ooxml/charts.rs
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
    /// Chart type.
    pub chart_type: ChartType,
    /// Chart title (optional).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub title: Option<ChartTitle>,
    /// Data series.
    pub series: Vec<ChartSeries>,
    /// Category axis (axis id 0). None for pie/doughnut.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cat_axis: Option<ChartAxis>,
    /// Value axis (axis id 1). None for pie/doughnut.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub val_axis: Option<ChartAxis>,
    /// Legend settings.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub legend: Option<ChartLegend>,
    /// Data labels (series-level default).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub data_labels: Option<DataLabels>,
    /// Grouping (bar/column/area/line).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub grouping: Option<ChartGrouping>,
    /// Scatter style (scatter only).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub scatter_style: Option<ScatterStyle>,
    /// Radar style (radar only).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub radar_style: Option<RadarStyle>,
    /// Doughnut hole size percentage (0-90, doughnut only).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub hole_size: Option<u32>,
    /// Bar direction: true = horizontal bars, false = vertical columns.
    /// Only relevant for Bar/Column type.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub bar_dir_horizontal: Option<bool>,
    /// Chart style ID (1-48, matches Excel's built-in chart styles).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub style_id: Option<u32>,
    /// Plot area manual layout (fractional 0.0-1.0).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub plot_area_layout: Option<ManualLayout>,
}

/// A data series within a chart.
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ChartSeries {
    /// Series index (0-based).
    pub idx: u32,
    /// Series order in the chart.
    pub order: u32,
    /// Series name (plain text or cell reference like "Sheet1!$B$1").
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub name: Option<String>,
    /// Category data reference (e.g. "Sheet1!$A$2:$A$5").
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cat_ref: Option<String>,
    /// Value data reference (e.g. "Sheet1!$B$2:$B$5").
    pub val_ref: String,
    /// X-value reference for scatter/bubble charts.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub x_val_ref: Option<String>,
    /// Bubble size reference for bubble charts.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub bubble_size_ref: Option<String>,
    /// Fill color (hex RGB, e.g. "4472C4").
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub fill_color: Option<String>,
    /// Line/border color (hex RGB).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub line_color: Option<String>,
    /// Line width in EMUs (12700 = 1pt).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub line_width: Option<u32>,
    /// Marker style (for line/scatter charts).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub marker: Option<MarkerStyle>,
    /// Whether to use smooth lines (line/scatter).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub smooth: Option<bool>,
    /// Pie/doughnut explosion percentage (0-100).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub explosion: Option<u32>,
    /// Series-level data labels override.
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
    /// Title text.
    pub text: String,
    /// Whether the title overlays the chart area.
    #[serde(default)]
    pub overlay: bool,
    /// Font size in hundredths of a point (e.g. 1400 = 14pt).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub font_size: Option<u32>,
    /// Font bold.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub bold: Option<bool>,
    /// Font color (hex RGB).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub color: Option<String>,
}

/// Chart axis definition.
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ChartAxis {
    /// Axis ID (must be unique within the chart, typically 0 and 1).
    pub id: u32,
    /// ID of the crossing axis.
    pub cross_ax: u32,
    /// Axis title.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub title: Option<ChartTitle>,
    /// Number format code (e.g. "General", "#,##0.00").
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub num_fmt: Option<String>,
    /// Whether number format is linked to source data.
    #[serde(default)]
    pub source_linked: bool,
    /// Minimum scale value.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub min: Option<f64>,
    /// Maximum scale value.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub max: Option<f64>,
    /// Major unit (e.g. 10 for gridlines every 10 units).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub major_unit: Option<f64>,
    /// Minor unit.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub minor_unit: Option<f64>,
    /// Logarithmic base (e.g. 10 for log scale).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub log_base: Option<f64>,
    /// Axis orientation: false = minMax (normal), true = maxMin (reversed).
    #[serde(default)]
    pub reversed: bool,
    /// Tick label position.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tick_lbl_pos: Option<TickLabelPosition>,
    /// Major tick mark style.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub major_tick_mark: Option<TickMark>,
    /// Minor tick mark style.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub minor_tick_mark: Option<TickMark>,
    /// Show major gridlines.
    #[serde(default)]
    pub major_gridlines: bool,
    /// Show minor gridlines.
    #[serde(default)]
    pub minor_gridlines: bool,
    /// Delete axis (hide it but keep scaling).
    #[serde(default)]
    pub delete: bool,
    /// Axis position: bottom, top, left, right.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub position: Option<AxisPosition>,
    /// Crossing point value (where this axis crosses the other).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub crosses_at: Option<f64>,
    /// Font size for tick labels.
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
    /// Legend position.
    pub position: LegendPosition,
    /// Whether the legend overlays the plot area.
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
    /// Show the numeric value.
    #[serde(default)]
    pub show_val: bool,
    /// Show the category name.
    #[serde(default)]
    pub show_cat_name: bool,
    /// Show the series name.
    #[serde(default)]
    pub show_ser_name: bool,
    /// Show percentage (pie/doughnut).
    #[serde(default)]
    pub show_percent: bool,
    /// Number format for label values.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub num_fmt: Option<String>,
    /// Show leader lines (pie/doughnut).
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
    /// Starting column (0-based).
    pub from_col: u32,
    /// Starting row (0-based).
    pub from_row: u32,
    /// Column offset in EMUs.
    #[serde(default)]
    pub from_col_off: u64,
    /// Row offset in EMUs.
    #[serde(default)]
    pub from_row_off: u64,
    /// Ending column (0-based).
    pub to_col: u32,
    /// Ending row (0-based).
    pub to_row: u32,
    /// Column offset in EMUs.
    #[serde(default)]
    pub to_col_off: u64,
    /// Row offset in EMUs.
    #[serde(default)]
    pub to_row_off: u64,
}

/// A chart embedded in a worksheet, with its anchor position.
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct WorksheetChart {
    /// The chart definition.
    pub chart: ChartData,
    /// The anchor position on the worksheet.
    pub anchor: ChartAnchor,
}
```

**Step 2: Register module and add charts to WorksheetXml**

In `crates/modern-xlsx-core/src/ooxml/mod.rs`, add:
```rust
pub mod charts;
```

In `crates/modern-xlsx-core/src/ooxml/worksheet.rs`, add to `WorksheetXml`:
```rust
use super::charts::WorksheetChart;

// In the WorksheetXml struct:
/// Charts embedded in this worksheet.
#[serde(default, skip_serializing_if = "Vec::is_empty")]
pub charts: Vec<WorksheetChart>,
```

**Step 3: Write unit tests**

```rust
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
        // Optional fields with None should be omitted
        assert!(!json.contains("title"));
        assert!(!json.contains("catAxis"));
        assert!(!json.contains("legend"));
        assert!(!json.contains("scatterStyle"));
    }
}
```

**Step 4: Run tests**

Run: `cargo test -p modern-xlsx-core -- charts`
Expected: 5 tests pass

**Step 5: Add TypeScript type exports**

Modify `packages/modern-xlsx/src/types.ts` — add all chart-related interfaces mirroring the Rust types:

```typescript
// Chart types
export type ChartType = 'bar' | 'column' | 'line' | 'pie' | 'doughnut' | 'scatter' | 'area' | 'radar' | 'bubble' | 'stock';
export type ChartGrouping = 'clustered' | 'stacked' | 'percentStacked' | 'standard';
export type ScatterStyle = 'lineMarker' | 'line' | 'marker' | 'smooth' | 'smoothMarker';
export type RadarStyle = 'standard' | 'marker' | 'filled';
export type MarkerStyle = 'circle' | 'square' | 'diamond' | 'triangle' | 'star' | 'x' | 'plus' | 'dash' | 'dot' | 'none';
export type TickLabelPosition = 'high' | 'low' | 'nextTo' | 'none';
export type TickMark = 'cross' | 'in' | 'out' | 'none';
export type AxisPosition = 'bottom' | 'top' | 'left' | 'right';
export type LegendPosition = 'top' | 'bottom' | 'left' | 'right' | 'topRight';

export interface ChartTitleData { text: string; overlay?: boolean; fontSize?: number; bold?: boolean; color?: string; }
export interface ChartAxisData { id: number; crossAx: number; title?: ChartTitleData | null; numFmt?: string | null; sourceLinked?: boolean; min?: number | null; max?: number | null; majorUnit?: number | null; minorUnit?: number | null; logBase?: number | null; reversed?: boolean; tickLblPos?: TickLabelPosition | null; majorTickMark?: TickMark | null; minorTickMark?: TickMark | null; majorGridlines?: boolean; minorGridlines?: boolean; delete?: boolean; position?: AxisPosition | null; crossesAt?: number | null; fontSize?: number | null; }
export interface ChartLegendData { position: LegendPosition; overlay?: boolean; }
export interface DataLabelsData { showVal?: boolean; showCatName?: boolean; showSerName?: boolean; showPercent?: boolean; numFmt?: string | null; showLeaderLines?: boolean; }
export interface ChartSeriesData { idx: number; order: number; name?: string | null; catRef?: string | null; valRef: string; xValRef?: string | null; bubbleSizeRef?: string | null; fillColor?: string | null; lineColor?: string | null; lineWidth?: number | null; marker?: MarkerStyle | null; smooth?: boolean | null; explosion?: number | null; dataLabels?: DataLabelsData | null; }
export interface ManualLayoutData { x: number; y: number; w: number; h: number; }
export interface ChartData { chartType: ChartType; title?: ChartTitleData | null; series: ChartSeriesData[]; catAxis?: ChartAxisData | null; valAxis?: ChartAxisData | null; legend?: ChartLegendData | null; dataLabels?: DataLabelsData | null; grouping?: ChartGrouping | null; scatterStyle?: ScatterStyle | null; radarStyle?: RadarStyle | null; holeSize?: number | null; barDirHorizontal?: boolean | null; styleId?: number | null; plotAreaLayout?: ManualLayoutData | null; }
export interface ChartAnchorData { fromCol: number; fromRow: number; fromColOff?: number; fromRowOff?: number; toCol: number; toRow: number; toColOff?: number; toRowOff?: number; }
export interface WorksheetChartData { chart: ChartData; anchor: ChartAnchorData; }
```

Add `charts` field to `WorksheetData`:
```typescript
// In WorksheetData interface:
charts?: WorksheetChartData[];
```

Export all chart types from `packages/modern-xlsx/src/index.ts`.

**Step 6: Commit**

```bash
git add crates/modern-xlsx-core/src/ooxml/charts.rs crates/modern-xlsx-core/src/ooxml/mod.rs crates/modern-xlsx-core/src/ooxml/worksheet.rs packages/modern-xlsx/src/types.ts packages/modern-xlsx/src/index.ts
git commit -m "feat(charts): add chart data model types (v0.8.0)"
```

---

## 0.8.2 — DrawingML Chart Writer (Bar, Column, Line)

### Task 2: Write chart XML generation for bar, column, and line charts

**Files:**
- Modify: `crates/modern-xlsx-core/src/ooxml/charts.rs` — add `to_xml()` method
- Test: embedded `#[cfg(test)]` module

The `to_xml()` method generates complete `xl/charts/chart{n}.xml`. Structure:

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
              xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
              xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <c:chart>
    <c:title>...</c:title>
    <c:plotArea>
      <c:layout/>
      <c:barChart>  <!-- or c:lineChart -->
        <c:barDir val="col"/>
        <c:grouping val="clustered"/>
        <c:ser>
          <c:idx val="0"/>
          <c:order val="0"/>
          <c:tx><c:strRef><c:f>Sheet1!$B$1</c:f></c:strRef></c:tx>
          <c:cat><c:strRef><c:f>Sheet1!$A$2:$A$5</c:f></c:strRef></c:cat>
          <c:val><c:numRef><c:f>Sheet1!$B$2:$B$5</c:f></c:numRef></c:val>
        </c:ser>
        <c:axId val="0"/>
        <c:axId val="1"/>
      </c:barChart>
      <c:catAx>
        <c:axId val="0"/>
        <c:scaling><c:orientation val="minMax"/></c:scaling>
        <c:delete val="0"/>
        <c:axPos val="b"/>
        <c:crossAx val="1"/>
      </c:catAx>
      <c:valAx>
        <c:axId val="1"/>
        <c:scaling><c:orientation val="minMax"/></c:scaling>
        <c:delete val="0"/>
        <c:axPos val="l"/>
        <c:crossAx val="0"/>
        <c:numFmt formatCode="General" sourceLinked="1"/>
        <c:majorGridlines/>
      </c:valAx>
    </c:plotArea>
    <c:legend>
      <c:legendPos val="b"/>
      <c:overlay val="0"/>
    </c:legend>
  </c:chart>
</c:chartSpace>
```

Implement `ChartData::to_xml(&self) -> Result<Vec<u8>>` using quick-xml Writer.

Tests: bar chart → valid XML, column chart → valid XML, line chart with markers → valid XML. Parse back the generated XML to verify element structure.

---

## 0.8.3 — DrawingML Chart Writer (Pie, Scatter, Area)

### Task 3: Chart XML for pie, doughnut, scatter, area, radar

Same approach as Task 2 but for:
- `<c:pieChart>` — no axes, supports explosion
- `<c:doughnutChart>` — holeSize element
- `<c:scatterChart>` — uses `<c:xVal>` and `<c:yVal>` instead of cat/val
- `<c:areaChart>` — uses grouping (standard/stacked/percentStacked)
- `<c:radarChart>` — uses radarStyle

---

## 0.8.4 — Chart Drawing Anchors & Packaging

### Task 4: Generate drawing XML, relationships, and content types for charts

**Files:**
- Modify: `crates/modern-xlsx-core/src/ooxml/charts.rs` — add `generate_drawing_xml()` and `generate_drawing_rels()`
- Modify: `crates/modern-xlsx-core/src/writer.rs` — integrate chart writing into ZIP output
- Modify: `crates/modern-xlsx-core/src/ooxml/content_types.rs` — add chart content type constant

The writer needs to:
1. For each sheet with `charts`, generate `xl/charts/chart{globalId}.xml`
2. Generate `xl/drawings/drawing{sheetNum}.xml` with `<xdr:twoCellAnchor>` referencing chart
3. Generate `xl/drawings/_rels/drawing{sheetNum}.xml.rels`
4. Update `xl/worksheets/_rels/sheet{sheetNum}.xml.rels` with drawing relationship
5. Add content type overrides for chart and drawing parts

Drawing XML for a chart anchor:
```xml
<xdr:wsDr xmlns:xdr="...spreadsheetDrawing" xmlns:a="...main" xmlns:r="...relationships">
  <xdr:twoCellAnchor>
    <xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
    <xdr:to><xdr:col>8</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>15</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
    <xdr:graphicFrame macro="">
      <xdr:nvGraphicFramePr>
        <xdr:cNvPr id="2" name="Chart 1"/>
        <xdr:cNvGraphicFramePr/>
      </xdr:nvGraphicFramePr>
      <xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>
      <a:graphic>
        <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
          <c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="rId1"/>
        </a:graphicData>
      </a:graphic>
    </xdr:graphicFrame>
    <xdr:clientData/>
  </xdr:twoCellAnchor>
</xdr:wsDr>
```

---

## 0.8.5 — Chart Titles, Labels & Legend Writer

### Task 5: Write title, axis title, data labels, and legend XML

Extend the `to_xml()` method to write:
- `<c:title>` with rich text (`<c:tx><c:rich>`)
- Axis titles (same structure as chart title, nested in `<c:catAx>` / `<c:valAx>`)
- `<c:dLbls>` with `<c:showVal>`, `<c:showCatName>`, `<c:showPercent>`, `<c:numFmt>`
- `<c:legend>` with position and overlay

---

## 0.8.6 — TypeScript Chart Creation API

### Task 6: Add addChart() to Worksheet and ChartBuilder fluent API

**Files:**
- Create: `packages/modern-xlsx/src/chart-builder.ts`
- Modify: `packages/modern-xlsx/src/workbook.ts` — add `addChart()` to Worksheet
- Modify: `packages/modern-xlsx/src/index.ts` — export ChartBuilder
- Create: `packages/modern-xlsx/__tests__/chart-builder.test.ts`

```typescript
// packages/modern-xlsx/src/chart-builder.ts
export class ChartBuilder {
  private data: Partial<ChartData> = {};
  private anchorData: ChartAnchorData = { fromCol: 0, fromRow: 0, toCol: 8, toRow: 15 };

  constructor(type: ChartType) { this.data.chartType = type; }

  title(text: string, opts?: { bold?: boolean; fontSize?: number; color?: string }): this { ... }
  addSeries(opts: { name?: string; catRef?: string; valRef: string; fillColor?: string; ... }): this { ... }
  catAxis(opts?: Partial<ChartAxisData>): this { ... }
  valAxis(opts?: Partial<ChartAxisData>): this { ... }
  legend(position?: LegendPosition, overlay?: boolean): this { ... }
  dataLabels(opts: Partial<DataLabelsData>): this { ... }
  grouping(g: ChartGrouping): this { ... }
  anchor(from: { col: number; row: number }, to: { col: number; row: number }): this { ... }
  style(id: number): this { ... }
  build(): WorksheetChartData { ... }
}
```

Worksheet API:
```typescript
class Worksheet {
  addChart(type: ChartType, options: (builder: ChartBuilder) => void): void {
    const builder = new ChartBuilder(type);
    options(builder);
    this.data.worksheet.charts.push(builder.build());
  }

  get charts(): readonly WorksheetChartData[] {
    return this.data.worksheet.charts ?? [];
  }
}
```

Tests: create bar chart → verify JSON structure, create pie chart → verify no axes, create scatter → verify xValRef, create chart with full options → verify all fields.

---

## 0.8.7 — Chart XML Reader

### Task 7: Parse chart XML into ChartData structs

**Files:**
- Modify: `crates/modern-xlsx-core/src/ooxml/charts.rs` — add `parse()` method
- Modify: `crates/modern-xlsx-core/src/reader.rs` — add `resolve_charts()` function
- Create: `packages/modern-xlsx/__tests__/chart-roundtrip.test.ts`

The reader needs to:
1. Find chart relationships in `xl/worksheets/_rels/sheet{n}.xml.rels` → drawing
2. Parse `xl/drawings/drawing{n}.xml` for `<xdr:graphicFrame>` with chart refs
3. Extract chart rId → resolve to `xl/charts/chart{n}.xml`
4. Parse chart XML SAX-style into ChartData
5. Extract anchor positioning from `<xdr:twoCellAnchor>`
6. Store as `Vec<WorksheetChart>` on the `WorksheetXml`

Move parsed charts OUT of `preserved_entries` (same pattern as tables). Unknown chart elements are preserved as raw XML in a `preserved_xml: Option<String>` field on ChartData.

---

## 0.8.8 — Chart Roundtrip & Style Presets

### Task 8: Full roundtrip fidelity and chart style presets

- Create chart → write → read → compare all fields
- Read Excel-generated chart → write → read → compare
- Chart style presets: map style_id (1-48) to predefined color palettes
- Apply palette during XML writing (fill colors, line styles from theme)

---

## 0.8.9 — Advanced: Combo Charts, Trendlines, Error Bars

### Task 9: Advanced chart features

**Files:**
- Modify: `crates/modern-xlsx-core/src/ooxml/charts.rs` — add trendline/error bar types

Add to ChartSeries:
```rust
/// Trendline (linear, exponential, logarithmic, polynomial, moving average).
#[serde(default, skip_serializing_if = "Option::is_none")]
pub trendline: Option<Trendline>,
/// Error bars.
#[serde(default, skip_serializing_if = "Option::is_none")]
pub error_bars: Option<ErrorBars>,
```

Add to ChartData:
```rust
/// Secondary chart for combo charts (e.g. bar + line).
#[serde(default, skip_serializing_if = "Option::is_none")]
pub secondary_chart: Option<Box<ChartData>>,
/// Secondary value axis (axis id 2, for combo charts).
#[serde(default, skip_serializing_if = "Option::is_none")]
pub secondary_val_axis: Option<ChartAxis>,
/// Data table below chart.
#[serde(default)]
pub show_data_table: bool,
/// 3D rotation settings.
#[serde(default, skip_serializing_if = "Option::is_none")]
pub view_3d: Option<View3D>,
```

New types:
```rust
pub struct Trendline {
    pub trend_type: TrendlineType,
    pub order: Option<u32>,    // polynomial order
    pub period: Option<u32>,   // moving average period
    pub forward: Option<f64>,  // forecast forward
    pub backward: Option<f64>, // forecast backward
    pub display_eq: bool,
    pub display_r_sqr: bool,
}

pub enum TrendlineType { Linear, Exponential, Logarithmic, Polynomial, Power, MovingAverage }

pub struct ErrorBars {
    pub err_type: ErrorBarType,
    pub value: Option<f64>,
}

pub enum ErrorBarType { FixedVal, Percentage, StdDev, StdErr, Custom }

pub struct View3D {
    pub rot_x: Option<i32>,   // -90 to 90
    pub rot_y: Option<i32>,   // 0 to 360
    pub perspective: Option<u32>, // 0 to 240
    pub r_ang_ax: Option<bool>,
}
```

---

## Dependencies

```
0.8.0 (Data Model) ──→ 0.8.2 (Bar/Line Writer) ──→ 0.8.4 (Drawing Anchors)
       │                0.8.3 (Pie/Scatter Writer) ─┘         │
       │                                                       ├──→ 0.8.5 (Titles/Labels)
       └──→ 0.8.1 (Axis Types) ──────────────────────────────┘         │
                                                                         ├──→ 0.8.6 (TS API)
                                                                         ├──→ 0.8.7 (Reader)
                                                                         └──→ 0.8.8 (Roundtrip) ──→ 0.8.9 (Advanced)
```

## Audit Requirement

After all tasks are complete, perform the standard comprehensive audit:
1. Every Rust and TypeScript file touched — modernization, performance, correctness
2. All new code uses `Vec::with_capacity`, zero-alloc patterns, saturating arithmetic
3. All serde attributes follow `camelCase` convention
4. All XML generation uses namespace-prefixed elements
5. Content types and relationships are spec-compliant
6. TypeScript types mirror Rust types 1:1
