use super::charts::WorksheetChart;
use super::shared_strings::RichTextRun;
use serde::{Deserialize, Serialize};

mod json;
mod parser;
mod writer;

pub use json::json_escape_to_pub;
// Inherent methods (parse, parse_with_sst, parse_to_json, to_xml, to_xml_with_sst)
// are automatically available on WorksheetXml — no re-export needed.

// `parse_col_element` is used by parser.rs directly via `super::parse_col_element`.
// Re-export it so the path resolves.
use writer::parse_col_element;

/// Header/footer configuration from `<headerFooter>`.
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct HeaderFooter {
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub odd_header: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub odd_footer: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub even_header: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub even_footer: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub first_header: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub first_footer: Option<String>,
    #[serde(default, skip_serializing_if = "crate::ooxml::is_false")]
    pub different_odd_even: bool,
    #[serde(default, skip_serializing_if = "crate::ooxml::is_false")]
    pub different_first: bool,
    #[serde(default = "crate::ooxml::default_true", skip_serializing_if = "crate::ooxml::is_true")]
    pub scale_with_doc: bool,
    #[serde(default = "crate::ooxml::default_true", skip_serializing_if = "crate::ooxml::is_true")]
    pub align_with_margins: bool,
}

impl Default for HeaderFooter {
    fn default() -> Self {
        Self {
            odd_header: None,
            odd_footer: None,
            even_header: None,
            even_footer: None,
            first_header: None,
            first_footer: None,
            different_odd_even: false,
            different_first: false,
            scale_with_doc: true,
            align_with_margins: true,
        }
    }
}

/// Outline (grouping) properties from `<sheetPr><outlinePr>`.
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct OutlineProperties {
    #[serde(default = "crate::ooxml::default_true")]
    pub summary_below: bool,
    #[serde(default = "crate::ooxml::default_true")]
    pub summary_right: bool,
}

/// Page break data for row or column breaks.
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct PageBreak {
    pub id: u32,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub min: Option<u32>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub max: Option<u32>,
    #[serde(default, skip_serializing_if = "crate::ooxml::is_false")]
    pub man: bool,
}

/// Row and column page breaks.
#[derive(Debug, Clone, Default, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct PageBreaks {
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub row_breaks: Vec<PageBreak>,
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub col_breaks: Vec<PageBreak>,
}

/// Parsed representation of a worksheet XML file (`xl/worksheets/sheet*.xml`).
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct WorksheetXml {
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub dimension: Option<String>,
    pub rows: Vec<Row>,
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub merge_cells: Vec<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub auto_filter: Option<AutoFilter>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub frozen_pane: Option<FrozenPane>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub split_pane: Option<SplitPane>,
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub pane_selections: Vec<PaneSelection>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub sheet_view: Option<SheetViewData>,
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub columns: Vec<ColumnInfo>,
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub data_validations: Vec<DataValidation>,
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub conditional_formatting: Vec<ConditionalFormatting>,
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub hyperlinks: Vec<Hyperlink>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub page_setup: Option<PageSetup>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub sheet_protection: Option<SheetProtection>,
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub comments: Vec<super::comments::Comment>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tab_color: Option<String>,
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub tables: Vec<super::tables::TableDefinition>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub header_footer: Option<HeaderFooter>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub page_breaks: Option<PageBreaks>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub outline_properties: Option<OutlineProperties>,
    /// Sparkline groups (from x14 extension in extLst).
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub sparkline_groups: Vec<SparklineGroup>,
    /// Charts embedded in this worksheet.
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub charts: Vec<WorksheetChart>,
    /// Pivot tables attached to this worksheet.
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub pivot_tables: Vec<super::pivot_table::PivotTableData>,
    /// Threaded comments (Microsoft 365 modern comments) attached to this worksheet.
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub threaded_comments: Vec<super::threaded_comments::ThreadedCommentData>,
    /// Slicers attached to this worksheet.
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub slicers: Vec<super::slicers::SlicerData>,
    /// Timelines attached to this worksheet.
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub timelines: Vec<super::timelines::TimelineData>,
    /// Non-sparkline extension XML preserved as raw strings.
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub preserved_extensions: Vec<String>,
}

/// A single row in the worksheet.
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct Row {
    /// 1-based row index (as in the XML `r` attribute).
    pub index: u32,
    pub cells: Vec<Cell>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub height: Option<f64>,
    #[serde(default, skip_serializing_if = "crate::ooxml::is_false")]
    pub hidden: bool,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub outline_level: Option<u8>,
    #[serde(default, skip_serializing_if = "crate::ooxml::is_false")]
    pub collapsed: bool,
}

/// A single cell in a row.
#[derive(Debug, Clone, Default, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct Cell {
    /// Cell reference string, e.g. "A1".
    pub reference: String,
    pub cell_type: CellType,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub style_index: Option<u32>,
    /// Raw `<v>` content (the value element text).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub value: Option<String>,
    /// Raw `<f>` content (the formula element text).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub formula: Option<String>,
    /// Formula type: "array", "shared", or "dataTable", from `t` attribute on `<f>`.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub formula_type: Option<String>,
    /// Range for array/shared formulas, from `ref` attribute on `<f>`.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub formula_ref: Option<String>,
    /// Shared formula index, from `si` attribute on `<f>`.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub shared_index: Option<u32>,
    /// Inline string value from `<is><t>...</t></is>`.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub inline_string: Option<String>,
    /// Whether this is a dynamic array formula (CSE/SPILL), from `cm` attribute.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub dynamic_array: Option<bool>,
    /// Data table 1st input cell (`r1` attribute on `<f>`).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub formula_r1: Option<String>,
    /// Data table 2nd input cell (`r2` attribute on `<f>`).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub formula_r2: Option<String>,
    /// 2D data table flag (`dt2D` attribute on `<f>`).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub formula_dt2d: Option<bool>,
    /// Data table row input deleted (`dtr1` attribute on `<f>`).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub formula_dtr1: Option<bool>,
    /// Data table column input deleted (`dtr2` attribute on `<f>`).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub formula_dtr2: Option<bool>,
    /// Rich text runs for this cell (when the cell contains mixed-format text).
    /// Populated from the shared string table during reading, or set directly
    /// by the API when constructing rich text cells for writing.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub rich_text: Option<Vec<RichTextRun>>,
}

/// The type of a cell, determined by the `t` attribute on the `<c>` element.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum CellType {
    /// `t="s"` — value is a shared string table index.
    SharedString,
    /// `t="n"` or omitted — numeric value.
    #[default]
    Number,
    /// `t="b"` — boolean (0 or 1).
    Boolean,
    /// `t="e"` — error value.
    Error,
    /// `t="str"` — formula string result.
    FormulaStr,
    /// `t="inlineStr"` — inline string (value stored in `<is><t>...</t></is>`).
    InlineStr,
    /// Explicitly empty cell (equivalent to SheetJS type "z").
    Stub,
}

/// Frozen pane configuration.
#[derive(Debug, Clone, Serialize, Deserialize, PartialEq)]
#[serde(rename_all = "camelCase")]
pub struct FrozenPane {
    /// Number of frozen rows (ySplit).
    pub rows: u32,
    /// Number of frozen columns (xSplit).
    pub cols: u32,
}

/// Split pane configuration — divides the sheet view into 2 or 4 scrollable regions.
/// `xSplit` / `ySplit` values are in **twips** (1/20th of a point) for split mode.
#[derive(Debug, Clone, Serialize, Deserialize, PartialEq)]
#[serde(rename_all = "camelCase")]
pub struct SplitPane {
    /// Horizontal split position in twips (ySplit). `None` or `0.0` means no horizontal split.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub horizontal: Option<f64>,
    /// Vertical split position in twips (xSplit). `None` or `0.0` means no vertical split.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub vertical: Option<f64>,
    /// Cell reference for the top-left cell in the bottom-right pane.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub top_left_cell: Option<String>,
    /// The active pane: `"topLeft"`, `"topRight"`, `"bottomLeft"`, `"bottomRight"`.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub active_pane: Option<String>,
}

/// Per-pane selection state within a `<sheetView>`.
/// Each visible pane can have its own active cell and selection range.
#[derive(Debug, Clone, Serialize, Deserialize, PartialEq)]
#[serde(rename_all = "camelCase")]
pub struct PaneSelection {
    /// Which pane this selection belongs to.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub pane: Option<String>,
    /// The active (focused) cell reference, e.g. `"A1"`.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub active_cell: Option<String>,
    /// The selected range, e.g. `"A1:C5"`.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub sqref: Option<String>,
}

/// Sheet view configuration from `<sheetView>` attributes (ECMA-376 §18.3.1.87).
#[derive(Debug, Clone, Serialize, Deserialize, PartialEq)]
#[serde(rename_all = "camelCase")]
pub struct SheetViewData {
    /// Whether grid lines are visible (default: true).
    #[serde(default = "crate::ooxml::default_true", skip_serializing_if = "crate::ooxml::is_true")]
    pub show_grid_lines: bool,
    /// Whether row and column headers are visible (default: true).
    #[serde(default = "crate::ooxml::default_true", skip_serializing_if = "crate::ooxml::is_true")]
    pub show_row_col_headers: bool,
    /// Whether zero values are displayed (default: true).
    #[serde(default = "crate::ooxml::default_true", skip_serializing_if = "crate::ooxml::is_true")]
    pub show_zeros: bool,
    /// Right-to-left display mode (default: false).
    #[serde(default, skip_serializing_if = "crate::ooxml::is_false")]
    pub right_to_left: bool,
    /// Whether this sheet tab is selected (default: false).
    #[serde(default, skip_serializing_if = "crate::ooxml::is_false")]
    pub tab_selected: bool,
    /// Whether the ruler is shown in Page Layout view (default: true).
    #[serde(default = "crate::ooxml::default_true", skip_serializing_if = "crate::ooxml::is_true")]
    pub show_ruler: bool,
    /// Whether outline (grouping) symbols are shown (default: true).
    #[serde(default = "crate::ooxml::default_true", skip_serializing_if = "crate::ooxml::is_true")]
    pub show_outline_symbols: bool,
    /// Whether white space around the page is shown in Page Layout view (default: true).
    #[serde(default = "crate::ooxml::default_true", skip_serializing_if = "crate::ooxml::is_true")]
    pub show_white_space: bool,
    /// Whether the default grid color is used (default: true).
    #[serde(default = "crate::ooxml::default_true", skip_serializing_if = "crate::ooxml::is_true")]
    pub default_grid_color: bool,
    /// Zoom percentage (10-400).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub zoom_scale: Option<u16>,
    /// Zoom percentage for Normal view.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub zoom_scale_normal: Option<u16>,
    /// Zoom percentage for Page Layout view.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub zoom_scale_page_layout_view: Option<u16>,
    /// Zoom percentage for Page Break Preview.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub zoom_scale_sheet_layout_view: Option<u16>,
    /// Theme color ID for the grid color.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub color_id: Option<u32>,
    /// View mode: `"normal"`, `"pageBreakPreview"`, or `"pageLayout"`.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub view: Option<String>,
}

impl Default for SheetViewData {
    fn default() -> Self {
        Self {
            show_grid_lines: true,
            show_row_col_headers: true,
            show_zeros: true,
            right_to_left: false,
            tab_selected: false,
            show_ruler: true,
            show_outline_symbols: true,
            show_white_space: true,
            default_grid_color: true,
            zoom_scale: None,
            zoom_scale_normal: None,
            zoom_scale_page_layout_view: None,
            zoom_scale_sheet_layout_view: None,
            color_id: None,
            view: None,
        }
    }
}

/// A group of sparklines sharing the same type, style, and options (ECMA-376 x14 extension).
#[derive(Debug, Clone, Serialize, Deserialize, PartialEq)]
#[serde(rename_all = "camelCase")]
pub struct SparklineGroup {
    /// Sparkline type: "line" (default), "column", or "stacked" (win/loss).
    #[serde(default = "default_sparkline_type", skip_serializing_if = "is_default_sparkline_type")]
    pub sparkline_type: String,
    /// Individual sparklines in this group.
    pub sparklines: Vec<Sparkline>,
    /// Series color (RGB hex, e.g. "FF376092").
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub color_series: Option<String>,
    /// Negative value color.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub color_negative: Option<String>,
    /// Axis color.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub color_axis: Option<String>,
    /// Marker color.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub color_markers: Option<String>,
    /// First point color.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub color_first: Option<String>,
    /// Last point color.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub color_last: Option<String>,
    /// High point color.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub color_high: Option<String>,
    /// Low point color.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub color_low: Option<String>,
    /// Line weight in points.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub line_weight: Option<f64>,
    /// Show markers on data points.
    #[serde(default, skip_serializing_if = "std::ops::Not::not")]
    pub markers: bool,
    /// Highlight high point.
    #[serde(default, skip_serializing_if = "std::ops::Not::not")]
    pub high: bool,
    /// Highlight low point.
    #[serde(default, skip_serializing_if = "std::ops::Not::not")]
    pub low: bool,
    /// Highlight first point.
    #[serde(default, skip_serializing_if = "std::ops::Not::not")]
    pub first: bool,
    /// Highlight last point.
    #[serde(default, skip_serializing_if = "std::ops::Not::not")]
    pub last: bool,
    /// Highlight negative values.
    #[serde(default, skip_serializing_if = "std::ops::Not::not")]
    pub negative: bool,
    /// Show the horizontal axis.
    #[serde(default, skip_serializing_if = "std::ops::Not::not")]
    pub display_x_axis: bool,
    /// How to display empty cells: "gap", "zero", or "span".
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub display_empty_cells_as: Option<String>,
    /// Manual minimum value for axis scaling.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub manual_min: Option<f64>,
    /// Manual maximum value for axis scaling.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub manual_max: Option<f64>,
    /// Right-to-left display.
    #[serde(default, skip_serializing_if = "std::ops::Not::not")]
    pub right_to_left: bool,
}

/// A single sparkline: data range -> display cell.
#[derive(Debug, Clone, Serialize, Deserialize, PartialEq)]
#[serde(rename_all = "camelCase")]
pub struct Sparkline {
    /// Data range formula (e.g. "Sheet1!A1:A10").
    pub formula: String,
    /// Cell where the sparkline renders (e.g. "B1").
    pub sqref: String,
}

impl Default for SparklineGroup {
    fn default() -> Self {
        Self {
            sparkline_type: default_sparkline_type(),
            sparklines: Vec::new(),
            color_series: None,
            color_negative: None,
            color_axis: None,
            color_markers: None,
            color_first: None,
            color_last: None,
            color_high: None,
            color_low: None,
            line_weight: None,
            markers: false,
            high: false,
            low: false,
            first: false,
            last: false,
            negative: false,
            display_x_axis: false,
            display_empty_cells_as: None,
            manual_min: None,
            manual_max: None,
            right_to_left: false,
        }
    }
}

#[inline]
fn default_sparkline_type() -> String {
    "line".to_string()
}

#[inline]
fn is_default_sparkline_type(s: &str) -> bool {
    s == "line"
}

/// Column formatting information from the `<cols>` section.
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ColumnInfo {
    pub min: u32,
    pub max: u32,
    pub width: f64,
    pub hidden: bool,
    pub custom_width: bool,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub outline_level: Option<u8>,
    #[serde(default, skip_serializing_if = "crate::ooxml::is_false")]
    pub collapsed: bool,
}

/// A data validation rule applied to a range of cells.
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct DataValidation {
    pub sqref: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub validation_type: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub operator: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub formula1: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub formula2: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub allow_blank: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub show_error_message: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub error_title: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub error_message: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub show_input_message: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub prompt_title: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub prompt: Option<String>,
}

/// A `<conditionalFormatting>` block in the worksheet.
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ConditionalFormatting {
    pub sqref: String,
    pub rules: Vec<ConditionalFormattingRule>,
}

/// A single `<cfRule>` inside a `<conditionalFormatting>` block.
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ConditionalFormattingRule {
    pub rule_type: String,
    pub priority: u32,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub operator: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub formula: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub dxf_id: Option<u32>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub color_scale: Option<ColorScale>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub data_bar: Option<DataBar>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub icon_set: Option<IconSet>,
}

/// A hyperlink in the worksheet.
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct Hyperlink {
    pub cell_ref: String,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub location: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub display: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tooltip: Option<String>,
}

/// Page setup configuration.
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct PageSetup {
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub paper_size: Option<u32>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub orientation: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub fit_to_width: Option<u32>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub fit_to_height: Option<u32>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub scale: Option<u32>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub first_page_number: Option<u32>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub horizontal_dpi: Option<u32>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub vertical_dpi: Option<u32>,
}

/// Sheet protection settings.
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct SheetProtection {
    pub sheet: bool,
    pub objects: bool,
    pub scenarios: bool,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub password: Option<String>,
    pub format_cells: bool,
    pub format_columns: bool,
    pub format_rows: bool,
    pub insert_columns: bool,
    pub insert_rows: bool,
    pub delete_columns: bool,
    pub delete_rows: bool,
    pub sort: bool,
    pub auto_filter: bool,
}

/// Auto filter with optional filter columns.
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct AutoFilter {
    pub range: String,
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub filter_columns: Vec<FilterColumn>,
}

/// A filter column within an auto filter.
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct FilterColumn {
    pub col_id: u32,
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub filters: Vec<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub custom_filters: Option<CustomFilters>,
}

/// Custom filters within a filter column (e.g., greater-than, contains).
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct CustomFilters {
    /// When `true`, all conditions must match (AND); when `false`, any match suffices (OR).
    #[serde(default, skip_serializing_if = "crate::ooxml::is_false")]
    pub and_op: bool,
    pub filters: Vec<CustomFilter>,
}

/// A single custom filter condition.
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct CustomFilter {
    /// Operator: "lessThan", "lessThanOrEqual", "equal", "notEqual",
    /// "greaterThanOrEqual", "greaterThan".  Absent means "equal".
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub operator: Option<String>,
    pub val: String,
}

/// A conditional value object used in color scales, data bars, and icon sets.
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct Cfvo {
    pub cfvo_type: String,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub val: Option<String>,
}

/// A color scale conditional formatting rule.
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ColorScale {
    pub cfvos: Vec<Cfvo>,
    pub colors: Vec<String>,
}

/// A data bar conditional formatting rule.
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct DataBar {
    pub cfvos: Vec<Cfvo>,
    pub color: String,
}

/// An icon set conditional formatting rule.
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct IconSet {
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub icon_set_type: Option<String>,
    pub cfvos: Vec<Cfvo>,
}

// ---------------------------------------------------------------------------
// Parser state machine
// ---------------------------------------------------------------------------

/// Internal parsing state.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) enum ParseState {
    Root,
    SheetPr,
    SheetViews,
    SheetView,
    Cols,
    SheetData,
    InRow,
    InCell,
    InCellValue,
    InCellFormula,
    InInlineStr,
    InInlineStrT,
    MergeCells,
    InDataValidations,
    InDataValidation,
    InDVFormula1,
    InDVFormula2,
    InConditionalFormatting,
    InCfRule,
    InCfRuleFormula,
    InHyperlinks,
    InAutoFilter,
    InFilterColumn,
    InFilters,
    InCustomFilters,
    InColorScale,
    InDataBar,
    InIconSet,
    HeaderFooter,
    HeaderFooterChild(u8), // 0=oddHeader, 1=oddFooter, 2=evenHeader, 3=evenFooter, 4=firstHeader, 5=firstFooter
    InRowBreaks,
    InColBreaks,
    ExtLst,
    ExtSparklines,
    SparklineGroups,
    SparklineGroup,
    Sparklines,
    SparklineItem,
    SparklineFormula,
    SparklineSqref,
    ExtOther(usize), // depth counter for skipping non-sparkline extensions
}

#[cfg(test)]
mod tests {
    use super::*;
    use super::writer::col_index_to_letter;
    use pretty_assertions::assert_eq;

    fn sample_worksheet_xml() -> &'static str {
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="A1:D2"/>
  <sheetViews>
    <sheetView workbookViewId="0">
      <pane ySplit="1" xSplit="0" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>
    </sheetView>
  </sheetViews>
  <cols>
    <col min="1" max="1" width="15" customWidth="1" hidden="0"/>
    <col min="2" max="2" width="20.5" customWidth="1"/>
  </cols>
  <sheetData>
    <row r="1" spans="1:4" ht="18">
      <c r="A1" t="s" s="3"><v>0</v></c>
      <c r="B1"><v>42.5</v></c>
      <c r="C1" t="b"><v>1</v></c>
      <c r="D1"><f>SUM(B1:B1)</f><v>42.5</v></c>
    </row>
    <row r="2" hidden="1">
      <c r="A2" t="inlineStr"><is><t>Hello World</t></is></c>
      <c r="B2" t="e"><v>#DIV/0!</v></c>
    </row>
  </sheetData>
  <mergeCells count="1">
    <mergeCell ref="A1:C1"/>
  </mergeCells>
  <autoFilter ref="A1:D1"/>
</worksheet>"#
    }

    #[test]
    fn test_parse_dimension() {
        let ws = WorksheetXml::parse(sample_worksheet_xml().as_bytes()).unwrap();
        assert_eq!(ws.dimension, Some("A1:D2".to_string()));
    }

    #[test]
    fn test_parse_frozen_pane() {
        let ws = WorksheetXml::parse(sample_worksheet_xml().as_bytes()).unwrap();
        let pane = ws.frozen_pane.as_ref().unwrap();
        assert_eq!(pane.rows, 1);
        assert_eq!(pane.cols, 0);
    }

    #[test]
    fn test_parse_horizontal_split_pane() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetViews>
    <sheetView workbookViewId="0">
      <pane ySplit="2400" topLeftCell="A5" activePane="bottomLeft" state="split"/>
      <selection pane="bottomLeft" activeCell="A5" sqref="A5"/>
    </sheetView>
  </sheetViews>
  <sheetData/>
</worksheet>"#;
        let ws = WorksheetXml::parse(xml.as_bytes()).unwrap();
        assert!(ws.frozen_pane.is_none(), "split pane must not set frozen_pane");
        let sp = ws.split_pane.as_ref().expect("split_pane should be Some");
        assert!((sp.horizontal.unwrap() - 2400.0).abs() < f64::EPSILON);
        assert!(sp.vertical.is_none());
        assert_eq!(sp.top_left_cell.as_deref(), Some("A5"));
        assert_eq!(sp.active_pane.as_deref(), Some("bottomLeft"));
        assert_eq!(ws.pane_selections.len(), 1);
        assert_eq!(ws.pane_selections[0].pane.as_deref(), Some("bottomLeft"));
        assert_eq!(ws.pane_selections[0].active_cell.as_deref(), Some("A5"));
        assert_eq!(ws.pane_selections[0].sqref.as_deref(), Some("A5"));
    }

    #[test]
    fn test_parse_vertical_split_pane() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetViews>
    <sheetView workbookViewId="0">
      <pane xSplit="3000" topLeftCell="D1" activePane="topRight" state="split"/>
      <selection pane="topRight" activeCell="D1" sqref="D1"/>
    </sheetView>
  </sheetViews>
  <sheetData/>
</worksheet>"#;
        let ws = WorksheetXml::parse(xml.as_bytes()).unwrap();
        assert!(ws.frozen_pane.is_none());
        let sp = ws.split_pane.as_ref().unwrap();
        assert!(sp.horizontal.is_none());
        assert!((sp.vertical.unwrap() - 3000.0).abs() < f64::EPSILON);
        assert_eq!(sp.top_left_cell.as_deref(), Some("D1"));
        assert_eq!(sp.active_pane.as_deref(), Some("topRight"));
    }

    #[test]
    fn test_parse_four_way_split_pane() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetViews>
    <sheetView workbookViewId="0">
      <pane xSplit="3000" ySplit="2400" topLeftCell="D5" activePane="bottomRight" state="split"/>
      <selection pane="topRight" activeCell="D1" sqref="D1"/>
      <selection pane="bottomLeft" activeCell="A5" sqref="A5"/>
      <selection pane="bottomRight" activeCell="D5" sqref="D5"/>
    </sheetView>
  </sheetViews>
  <sheetData/>
</worksheet>"#;
        let ws = WorksheetXml::parse(xml.as_bytes()).unwrap();
        assert!(ws.frozen_pane.is_none());
        let sp = ws.split_pane.as_ref().unwrap();
        assert!((sp.vertical.unwrap() - 3000.0).abs() < f64::EPSILON);
        assert!((sp.horizontal.unwrap() - 2400.0).abs() < f64::EPSILON);
        assert_eq!(sp.top_left_cell.as_deref(), Some("D5"));
        assert_eq!(sp.active_pane.as_deref(), Some("bottomRight"));
        assert_eq!(ws.pane_selections.len(), 3);
    }

    #[test]
    fn test_frozen_and_split_mutually_exclusive() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetViews>
    <sheetView workbookViewId="0">
      <pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>
    </sheetView>
  </sheetViews>
  <sheetData/>
</worksheet>"#;
        let ws = WorksheetXml::parse(xml.as_bytes()).unwrap();
        assert!(ws.frozen_pane.is_some());
        assert!(ws.split_pane.is_none());
    }

    #[test]
    fn test_parse_split_no_state_attr() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetViews>
    <sheetView workbookViewId="0">
      <pane xSplit="1500" ySplit="900" topLeftCell="B2"/>
    </sheetView>
  </sheetViews>
  <sheetData/>
</worksheet>"#;
        let ws = WorksheetXml::parse(xml.as_bytes()).unwrap();
        assert!(ws.frozen_pane.is_none());
        let sp = ws.split_pane.as_ref().unwrap();
        assert!((sp.vertical.unwrap() - 1500.0).abs() < f64::EPSILON);
        assert!((sp.horizontal.unwrap() - 900.0).abs() < f64::EPSILON);
    }

    #[test]
    fn test_parse_columns() {
        let ws = WorksheetXml::parse(sample_worksheet_xml().as_bytes()).unwrap();
        assert_eq!(ws.columns.len(), 2);

        assert_eq!(ws.columns[0].min, 1);
        assert_eq!(ws.columns[0].max, 1);
        assert_eq!(ws.columns[0].width, 15.0);
        assert!(ws.columns[0].custom_width);
        assert!(!ws.columns[0].hidden);

        assert_eq!(ws.columns[1].min, 2);
        assert_eq!(ws.columns[1].max, 2);
        assert_eq!(ws.columns[1].width, 20.5);
        assert!(ws.columns[1].custom_width);
    }

    #[test]
    fn test_parse_rows_and_cells() {
        let ws = WorksheetXml::parse(sample_worksheet_xml().as_bytes()).unwrap();
        assert_eq!(ws.rows.len(), 2);

        // Row 1.
        let row1 = &ws.rows[0];
        assert_eq!(row1.index, 1);
        assert_eq!(row1.height, Some(18.0));
        assert!(!row1.hidden);
        assert_eq!(row1.cells.len(), 4);

        // A1: shared string.
        assert_eq!(row1.cells[0].reference, "A1");
        assert_eq!(row1.cells[0].cell_type, CellType::SharedString);
        assert_eq!(row1.cells[0].style_index, Some(3));
        assert_eq!(row1.cells[0].value, Some("0".to_string()));
        assert_eq!(row1.cells[0].formula, None);

        // B1: number.
        assert_eq!(row1.cells[1].reference, "B1");
        assert_eq!(row1.cells[1].cell_type, CellType::Number);
        assert_eq!(row1.cells[1].style_index, None);
        assert_eq!(row1.cells[1].value, Some("42.5".to_string()));

        // C1: boolean.
        assert_eq!(row1.cells[2].reference, "C1");
        assert_eq!(row1.cells[2].cell_type, CellType::Boolean);
        assert_eq!(row1.cells[2].value, Some("1".to_string()));

        // D1: formula.
        assert_eq!(row1.cells[3].reference, "D1");
        assert_eq!(row1.cells[3].cell_type, CellType::Number);
        assert_eq!(row1.cells[3].formula, Some("SUM(B1:B1)".to_string()));
        assert_eq!(row1.cells[3].value, Some("42.5".to_string()));

        // Row 2.
        let row2 = &ws.rows[1];
        assert_eq!(row2.index, 2);
        assert!(row2.hidden);
        assert_eq!(row2.cells.len(), 2);

        // A2: inline string.
        assert_eq!(row2.cells[0].reference, "A2");
        assert_eq!(row2.cells[0].cell_type, CellType::InlineStr);
        assert_eq!(row2.cells[0].value, Some("Hello World".to_string()));

        // B2: error.
        assert_eq!(row2.cells[1].reference, "B2");
        assert_eq!(row2.cells[1].cell_type, CellType::Error);
        assert_eq!(row2.cells[1].value, Some("#DIV/0!".to_string()));
    }

    #[test]
    fn test_parse_merge_cells() {
        let ws = WorksheetXml::parse(sample_worksheet_xml().as_bytes()).unwrap();
        assert_eq!(ws.merge_cells, vec!["A1:C1"]);
    }

    #[test]
    fn test_parse_auto_filter() {
        let ws = WorksheetXml::parse(sample_worksheet_xml().as_bytes()).unwrap();
        let af = ws.auto_filter.as_ref().unwrap();
        assert_eq!(af.range, "A1:D1");
        assert!(af.filter_columns.is_empty());
    }

    #[test]
    fn test_write_full_roundtrip() {
        let ws = WorksheetXml {
            dimension: None,
            rows: vec![
                Row {
                    index: 1,
                    cells: vec![
                        Cell {
                            reference: "A1".to_string(),
                            cell_type: CellType::SharedString,
                            style_index: Some(0),
                            value: Some("0".to_string()),
                            ..Default::default()
                        },
                        Cell {
                            reference: "B1".to_string(),
                            cell_type: CellType::Number,
                            style_index: None,
                            value: Some("42.5".to_string()),
                            ..Default::default()
                        },
                        Cell {
                            reference: "C1".to_string(),
                            cell_type: CellType::Number,
                            value: Some("42.5".to_string()),
                            formula: Some("SUM(A1:B1)".to_string()),
                            ..Default::default()
                        },
                    ],
                    height: Some(18.0),
                    hidden: false,
                    outline_level: None,
                    collapsed: false,
                },
                Row {
                    index: 2,
                    cells: vec![Cell {
                        reference: "A2".to_string(),
                        cell_type: CellType::Boolean,
                        style_index: None,
                        value: Some("1".to_string()),
                        ..Default::default()
                    }],
                    height: None,
                    hidden: true,
                    outline_level: None,
                    collapsed: false,
                },
            ],
            merge_cells: vec!["A1:C1".to_string()],
            auto_filter: Some(AutoFilter { range: "A1:C1".to_string(), filter_columns: vec![] }),
            frozen_pane: Some(FrozenPane { rows: 1, cols: 0 }),
            split_pane: None,
            pane_selections: vec![],
            sheet_view: None,
            columns: vec![ColumnInfo {
                min: 1,
                max: 1,
                width: 15.0,
                hidden: false,
                custom_width: true,
                outline_level: None,
                collapsed: false,
            }],
            data_validations: vec![],
            conditional_formatting: vec![],
            hyperlinks: vec![],
            page_setup: None,
            sheet_protection: None,
            comments: Vec::new(),
            tab_color: None,
            tables: vec![],
            header_footer: None,
            page_breaks: None,
            outline_properties: None,
            sparkline_groups: vec![],
            charts: vec![],
            pivot_tables: vec![],
            threaded_comments: vec![],
            slicers: Vec::new(),
            timelines: Vec::new(),
            preserved_extensions: vec![],
        };

        let xml = ws.to_xml().unwrap();
        let ws2 = WorksheetXml::parse(&xml).unwrap();

        // Verify frozen pane.
        let pane = ws2.frozen_pane.as_ref().unwrap();
        assert_eq!(pane.rows, 1);
        assert_eq!(pane.cols, 0);

        // Verify columns.
        assert_eq!(ws2.columns.len(), 1);
        assert_eq!(ws2.columns[0].min, 1);
        assert_eq!(ws2.columns[0].max, 1);
        assert_eq!(ws2.columns[0].width, 15.0);
        assert!(ws2.columns[0].custom_width);

        // Verify rows.
        assert_eq!(ws2.rows.len(), 2);
        assert_eq!(ws2.rows[0].index, 1);
        assert_eq!(ws2.rows[0].height, Some(18.0));
        assert!(!ws2.rows[0].hidden);
        assert_eq!(ws2.rows[0].cells.len(), 3);

        // Cell A1.
        assert_eq!(ws2.rows[0].cells[0].reference, "A1");
        assert_eq!(ws2.rows[0].cells[0].cell_type, CellType::SharedString);
        assert_eq!(ws2.rows[0].cells[0].style_index, Some(0));
        assert_eq!(ws2.rows[0].cells[0].value, Some("0".to_string()));

        // Cell B1.
        assert_eq!(ws2.rows[0].cells[1].reference, "B1");
        assert_eq!(ws2.rows[0].cells[1].cell_type, CellType::Number);
        assert_eq!(ws2.rows[0].cells[1].value, Some("42.5".to_string()));

        // Cell C1 (formula).
        assert_eq!(ws2.rows[0].cells[2].reference, "C1");
        assert_eq!(ws2.rows[0].cells[2].formula, Some("SUM(A1:B1)".to_string()));
        assert_eq!(ws2.rows[0].cells[2].value, Some("42.5".to_string()));

        // Row 2.
        assert_eq!(ws2.rows[1].index, 2);
        assert!(ws2.rows[1].hidden);
        assert_eq!(ws2.rows[1].cells[0].cell_type, CellType::Boolean);

        // Merge cells.
        assert_eq!(ws2.merge_cells, vec!["A1:C1"]);

        // Auto filter.
        assert_eq!(ws2.auto_filter.as_ref().unwrap().range, "A1:C1");
    }

    #[test]
    fn test_write_minimal_worksheet() {
        let ws = WorksheetXml {
            dimension: None,
            rows: vec![Row {
                index: 1,
                cells: vec![Cell {
                    reference: "A1".to_string(),
                    cell_type: CellType::Number,
                    style_index: None,
                    value: Some("100".to_string()),
                    ..Default::default()
                }],
                height: None,
                hidden: false,
                outline_level: None,
                collapsed: false,
            }],
            merge_cells: Vec::new(),
            auto_filter: None,
            frozen_pane: None,
            split_pane: None,
            pane_selections: vec![],
            sheet_view: None,
            columns: Vec::new(),
            data_validations: vec![],
            conditional_formatting: vec![],
            hyperlinks: vec![],
            page_setup: None,
            sheet_protection: None,
            comments: Vec::new(),
            tab_color: None,
            tables: vec![],
            header_footer: None,
            page_breaks: None,
            outline_properties: None,
            sparkline_groups: vec![],
            charts: vec![],
            pivot_tables: vec![],
            threaded_comments: vec![],
            slicers: Vec::new(),
            timelines: Vec::new(),
            preserved_extensions: vec![],
        };

        let xml = ws.to_xml().unwrap();
        let xml_str = String::from_utf8(xml.clone()).unwrap();

        // Verify no optional sections appear.
        assert!(!xml_str.contains("sheetViews"));
        assert!(!xml_str.contains("<cols"));
        assert!(!xml_str.contains("mergeCells"));
        assert!(!xml_str.contains("autoFilter"));

        // But sheetData is present.
        assert!(xml_str.contains("sheetData"));

        // Re-parse.
        let ws2 = WorksheetXml::parse(&xml).unwrap();
        assert_eq!(ws2.rows.len(), 1);
        assert_eq!(ws2.rows[0].cells[0].value, Some("100".to_string()));
        assert!(ws2.frozen_pane.is_none());
        assert!(ws2.columns.is_empty());
        assert!(ws2.merge_cells.is_empty());
        assert!(ws2.auto_filter.is_none());
    }

    #[test]
    fn test_col_index_to_letter() {
        assert_eq!(col_index_to_letter(1), "A");
        assert_eq!(col_index_to_letter(26), "Z");
        assert_eq!(col_index_to_letter(27), "AA");
        assert_eq!(col_index_to_letter(28), "AB");
        assert_eq!(col_index_to_letter(702), "ZZ");
    }

    #[test]
    fn test_parse_formula_str_type() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="str"><f>IF(TRUE,"yes","no")</f><v>yes</v></c>
    </row>
  </sheetData>
</worksheet>"#;

        let ws = WorksheetXml::parse(xml.as_bytes()).unwrap();
        assert_eq!(ws.rows[0].cells[0].cell_type, CellType::FormulaStr);
        assert_eq!(
            ws.rows[0].cells[0].formula,
            Some("IF(TRUE,\"yes\",\"no\")".to_string())
        );
        assert_eq!(ws.rows[0].cells[0].value, Some("yes".to_string()));
    }

    #[test]
    fn test_write_does_not_emit_t_for_number() {
        let ws = WorksheetXml {
            dimension: None,
            rows: vec![Row {
                index: 1,
                cells: vec![Cell {
                    reference: "A1".to_string(),
                    cell_type: CellType::Number,
                    style_index: None,
                    value: Some("42".to_string()),
                    ..Default::default()
                }],
                height: None,
                hidden: false,
                outline_level: None,
                collapsed: false,
            }],
            merge_cells: Vec::new(),
            auto_filter: None,
            frozen_pane: None,
            split_pane: None,
            pane_selections: vec![],
            sheet_view: None,
            columns: Vec::new(),
            data_validations: vec![],
            conditional_formatting: vec![],
            hyperlinks: vec![],
            page_setup: None,
            sheet_protection: None,
            comments: Vec::new(),
            tab_color: None,
            tables: vec![],
            header_footer: None,
            page_breaks: None,
            outline_properties: None,
            sparkline_groups: vec![],
            charts: vec![],
            pivot_tables: vec![],
            threaded_comments: vec![],
            slicers: Vec::new(),
            timelines: Vec::new(),
            preserved_extensions: vec![],
        };

        let xml = ws.to_xml().unwrap();
        let xml_str = String::from_utf8(xml).unwrap();
        // The <c> element for a Number cell should NOT have a t="..." attribute.
        assert!(xml_str.contains(r#"<c r="A1"><v>42</v></c>"#));
    }

    #[test]
    fn test_write_emits_t_for_shared_string() {
        let ws = WorksheetXml {
            dimension: None,
            rows: vec![Row {
                index: 1,
                cells: vec![Cell {
                    reference: "A1".to_string(),
                    cell_type: CellType::SharedString,
                    style_index: Some(1),
                    value: Some("5".to_string()),
                    ..Default::default()
                }],
                height: None,
                hidden: false,
                outline_level: None,
                collapsed: false,
            }],
            merge_cells: Vec::new(),
            auto_filter: None,
            frozen_pane: None,
            split_pane: None,
            pane_selections: vec![],
            sheet_view: None,
            columns: Vec::new(),
            data_validations: vec![],
            conditional_formatting: vec![],
            hyperlinks: vec![],
            page_setup: None,
            sheet_protection: None,
            comments: Vec::new(),
            tab_color: None,
            tables: vec![],
            header_footer: None,
            page_breaks: None,
            outline_properties: None,
            sparkline_groups: vec![],
            charts: vec![],
            pivot_tables: vec![],
            threaded_comments: vec![],
            slicers: Vec::new(),
            timelines: Vec::new(),
            preserved_extensions: vec![],
        };

        let xml = ws.to_xml().unwrap();
        let xml_str = String::from_utf8(xml).unwrap();
        assert!(xml_str.contains(r#"t="s""#));
        assert!(xml_str.contains(r#"s="1""#));
    }

    #[test]
    fn test_empty_cell_no_value() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" s="2"/>
    </row>
  </sheetData>
</worksheet>"#;

        let ws = WorksheetXml::parse(xml.as_bytes()).unwrap();
        assert_eq!(ws.rows[0].cells[0].reference, "A1");
        assert_eq!(ws.rows[0].cells[0].style_index, Some(2));
        assert_eq!(ws.rows[0].cells[0].value, None);
        assert_eq!(ws.rows[0].cells[0].formula, None);
    }

    #[test]
    fn test_parse_data_validations() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData/>
  <dataValidations count="2">
    <dataValidation type="list" allowBlank="1" showErrorMessage="1" sqref="A1:A10">
      <formula1>"Yes,No,Maybe"</formula1>
    </dataValidation>
    <dataValidation type="whole" operator="between" sqref="B1:B10" errorTitle="Invalid" error="Must be 1-100">
      <formula1>1</formula1>
      <formula2>100</formula2>
    </dataValidation>
  </dataValidations>
</worksheet>"#;

        let ws = WorksheetXml::parse(xml.as_bytes()).unwrap();
        assert_eq!(ws.data_validations.len(), 2);

        let dv0 = &ws.data_validations[0];
        assert_eq!(dv0.sqref, "A1:A10");
        assert_eq!(dv0.validation_type, Some("list".to_string()));
        assert_eq!(dv0.operator, None);
        assert_eq!(dv0.allow_blank, Some(true));
        assert_eq!(dv0.show_error_message, Some(true));
        assert_eq!(dv0.formula1, Some("\"Yes,No,Maybe\"".to_string()));
        assert_eq!(dv0.formula2, None);

        let dv1 = &ws.data_validations[1];
        assert_eq!(dv1.sqref, "B1:B10");
        assert_eq!(dv1.validation_type, Some("whole".to_string()));
        assert_eq!(dv1.operator, Some("between".to_string()));
        assert_eq!(dv1.error_title, Some("Invalid".to_string()));
        assert_eq!(dv1.error_message, Some("Must be 1-100".to_string()));
        assert_eq!(dv1.formula1, Some("1".to_string()));
        assert_eq!(dv1.formula2, Some("100".to_string()));
    }

    #[test]
    fn test_data_validations_roundtrip() {
        let ws = WorksheetXml {
            dimension: None,
            rows: vec![],
            merge_cells: vec![],
            auto_filter: None,
            frozen_pane: None,
            split_pane: None,
            pane_selections: vec![],
            sheet_view: None,
            columns: vec![],
            data_validations: vec![
                DataValidation {
                    sqref: "A1:A10".to_string(),
                    validation_type: Some("list".to_string()),
                    operator: None,
                    formula1: Some("\"Yes,No,Maybe\"".to_string()),
                    formula2: None,
                    allow_blank: Some(true),
                    show_error_message: Some(true),
                    error_title: None,
                    error_message: None,
                    show_input_message: None,
                    prompt_title: None,
                    prompt: None,
                },
                DataValidation {
                    sqref: "B1:B10".to_string(),
                    validation_type: Some("whole".to_string()),
                    operator: Some("between".to_string()),
                    formula1: Some("1".to_string()),
                    formula2: Some("100".to_string()),
                    allow_blank: None,
                    show_error_message: None,
                    error_title: Some("Invalid".to_string()),
                    error_message: Some("Must be 1-100".to_string()),
                    show_input_message: None,
                    prompt_title: None,
                    prompt: None,
                },
            ],
            conditional_formatting: vec![],
            hyperlinks: vec![],
            page_setup: None,
            sheet_protection: None,
            comments: Vec::new(),
            tab_color: None,
            tables: vec![],
            header_footer: None,
            page_breaks: None,
            outline_properties: None,
            sparkline_groups: vec![],
            charts: vec![],
            pivot_tables: vec![],
            threaded_comments: vec![],
            slicers: Vec::new(),
            timelines: Vec::new(),
            preserved_extensions: vec![],
        };

        let xml = ws.to_xml().unwrap();
        let xml_str = String::from_utf8(xml.clone()).unwrap();

        // Verify XML contains data validations.
        assert!(xml_str.contains("dataValidations"));
        assert!(xml_str.contains(r#"count="2""#));
        assert!(xml_str.contains(r#"type="list""#));
        assert!(xml_str.contains(r#"sqref="A1:A10""#));

        // Re-parse and verify.
        let ws2 = WorksheetXml::parse(&xml).unwrap();
        assert_eq!(ws2.data_validations.len(), 2);

        let dv0 = &ws2.data_validations[0];
        assert_eq!(dv0.sqref, "A1:A10");
        assert_eq!(dv0.validation_type, Some("list".to_string()));
        assert_eq!(dv0.allow_blank, Some(true));
        assert_eq!(dv0.show_error_message, Some(true));
        assert_eq!(dv0.formula1, Some("\"Yes,No,Maybe\"".to_string()));
        assert_eq!(dv0.formula2, None);

        let dv1 = &ws2.data_validations[1];
        assert_eq!(dv1.sqref, "B1:B10");
        assert_eq!(dv1.validation_type, Some("whole".to_string()));
        assert_eq!(dv1.operator, Some("between".to_string()));
        assert_eq!(dv1.formula1, Some("1".to_string()));
        assert_eq!(dv1.formula2, Some("100".to_string()));
        assert_eq!(dv1.error_title, Some("Invalid".to_string()));
        assert_eq!(dv1.error_message, Some("Must be 1-100".to_string()));
    }

    #[test]
    fn test_parse_conditional_formatting() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1"><v>200</v></c>
    </row>
  </sheetData>
  <conditionalFormatting sqref="A1:B10">
    <cfRule type="cellIs" dxfId="0" priority="1" operator="greaterThan">
      <formula>100</formula>
    </cfRule>
  </conditionalFormatting>
  <conditionalFormatting sqref="C1:C10">
    <cfRule type="colorScale" priority="2">
      <colorScale>
        <cfvo type="min"/>
        <cfvo type="max"/>
        <color rgb="FFF8696B"/>
        <color rgb="FF63BE7B"/>
      </colorScale>
    </cfRule>
  </conditionalFormatting>
</worksheet>"#;

        let ws = WorksheetXml::parse(xml.as_bytes()).unwrap();
        assert_eq!(ws.conditional_formatting.len(), 2);

        // First block.
        let cf0 = &ws.conditional_formatting[0];
        assert_eq!(cf0.sqref, "A1:B10");
        assert_eq!(cf0.rules.len(), 1);
        assert_eq!(cf0.rules[0].rule_type, "cellIs");
        assert_eq!(cf0.rules[0].priority, 1);
        assert_eq!(cf0.rules[0].operator, Some("greaterThan".to_string()));
        assert_eq!(cf0.rules[0].dxf_id, Some(0));
        assert_eq!(cf0.rules[0].formula, Some("100".to_string()));

        // Second block (colorScale — no formula at cfRule level).
        let cf1 = &ws.conditional_formatting[1];
        assert_eq!(cf1.sqref, "C1:C10");
        assert_eq!(cf1.rules.len(), 1);
        assert_eq!(cf1.rules[0].rule_type, "colorScale");
        assert_eq!(cf1.rules[0].priority, 2);
        assert_eq!(cf1.rules[0].operator, None);
        assert_eq!(cf1.rules[0].dxf_id, None);
        assert_eq!(cf1.rules[0].formula, None);
    }

    #[test]
    fn test_conditional_formatting_roundtrip() {
        let ws = WorksheetXml {
            dimension: None,
            rows: vec![Row {
                index: 1,
                cells: vec![Cell {
                    reference: "A1".to_string(),
                    cell_type: CellType::Number,
                    style_index: None,
                    value: Some("200".to_string()),
                    ..Default::default()
                }],
                height: None,
                hidden: false,
                outline_level: None,
                collapsed: false,
            }],
            merge_cells: Vec::new(),
            auto_filter: None,
            frozen_pane: None,
            split_pane: None,
            pane_selections: vec![],
            sheet_view: None,
            columns: Vec::new(),
            data_validations: Vec::new(),
            conditional_formatting: vec![
                ConditionalFormatting {
                    sqref: "A1:B10".to_string(),
                    rules: vec![ConditionalFormattingRule {
                        rule_type: "cellIs".to_string(),
                        priority: 1,
                        operator: Some("greaterThan".to_string()),
                        formula: Some("100".to_string()),
                        dxf_id: Some(0),
                        color_scale: None,
                        data_bar: None,
                        icon_set: None,
                    }],
                },
                ConditionalFormatting {
                    sqref: "C1:C10".to_string(),
                    rules: vec![ConditionalFormattingRule {
                        rule_type: "colorScale".to_string(),
                        priority: 2,
                        operator: None,
                        formula: None,
                        dxf_id: None,
                        color_scale: None,
                        data_bar: None,
                        icon_set: None,
                    }],
                },
            ],
            hyperlinks: vec![],
            page_setup: None,
            sheet_protection: None,
            comments: Vec::new(),
            tab_color: None,
            tables: vec![],
            header_footer: None,
            page_breaks: None,
            outline_properties: None,
            sparkline_groups: vec![],
            charts: vec![],
            pivot_tables: vec![],
            threaded_comments: vec![],
            slicers: Vec::new(),
            timelines: Vec::new(),
            preserved_extensions: vec![],
        };

        // Write to XML.
        let xml = ws.to_xml().unwrap();
        let xml_str = String::from_utf8(xml.clone()).unwrap();

        // Verify the XML contains conditional formatting elements.
        assert!(xml_str.contains("conditionalFormatting"));
        assert!(xml_str.contains(r#"sqref="A1:B10"#));
        assert!(xml_str.contains(r#"type="cellIs""#));
        assert!(xml_str.contains(r#"operator="greaterThan""#));
        assert!(xml_str.contains("<formula>100</formula>"));

        // Parse back.
        let ws2 = WorksheetXml::parse(&xml).unwrap();
        assert_eq!(ws2.conditional_formatting.len(), 2);

        // Verify first block.
        let cf0 = &ws2.conditional_formatting[0];
        assert_eq!(cf0.sqref, "A1:B10");
        assert_eq!(cf0.rules.len(), 1);
        assert_eq!(cf0.rules[0].rule_type, "cellIs");
        assert_eq!(cf0.rules[0].priority, 1);
        assert_eq!(cf0.rules[0].operator, Some("greaterThan".to_string()));
        assert_eq!(cf0.rules[0].dxf_id, Some(0));
        assert_eq!(cf0.rules[0].formula, Some("100".to_string()));

        // Verify second block.
        let cf1 = &ws2.conditional_formatting[1];
        assert_eq!(cf1.sqref, "C1:C10");
        assert_eq!(cf1.rules.len(), 1);
        assert_eq!(cf1.rules[0].rule_type, "colorScale");
        assert_eq!(cf1.rules[0].priority, 2);
        assert_eq!(cf1.rules[0].operator, None);
        assert_eq!(cf1.rules[0].dxf_id, None);
        assert_eq!(cf1.rules[0].formula, None);
    }

    #[test]
    fn test_parse_hyperlinks() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData/>
  <hyperlinks>
    <hyperlink ref="A1" location="Sheet2!A1" display="Go to Sheet2" tooltip="Click to navigate"/>
    <hyperlink ref="B1" location="'My Sheet'!C3"/>
  </hyperlinks>
</worksheet>"#;

        let ws = WorksheetXml::parse(xml.as_bytes()).unwrap();
        assert_eq!(ws.hyperlinks.len(), 2);

        assert_eq!(ws.hyperlinks[0].cell_ref, "A1");
        assert_eq!(ws.hyperlinks[0].location, Some("Sheet2!A1".to_string()));
        assert_eq!(ws.hyperlinks[0].display, Some("Go to Sheet2".to_string()));
        assert_eq!(ws.hyperlinks[0].tooltip, Some("Click to navigate".to_string()));

        assert_eq!(ws.hyperlinks[1].cell_ref, "B1");
        assert_eq!(ws.hyperlinks[1].location, Some("'My Sheet'!C3".to_string()));
        assert_eq!(ws.hyperlinks[1].display, None);
        assert_eq!(ws.hyperlinks[1].tooltip, None);
    }

    #[test]
    fn test_hyperlinks_roundtrip() {
        let ws = WorksheetXml {
            dimension: None,
            rows: vec![],
            merge_cells: vec![],
            auto_filter: None,
            frozen_pane: None,
            split_pane: None,
            pane_selections: vec![],
            sheet_view: None,
            columns: vec![],
            data_validations: vec![],
            conditional_formatting: vec![],
            hyperlinks: vec![
                Hyperlink {
                    cell_ref: "A1".to_string(),
                    location: Some("Sheet2!A1".to_string()),
                    display: Some("Go to Sheet2".to_string()),
                    tooltip: Some("Click here".to_string()),
                },
                Hyperlink {
                    cell_ref: "B1".to_string(),
                    location: Some("Sheet3!B2".to_string()),
                    display: None,
                    tooltip: None,
                },
            ],
            page_setup: None,
            sheet_protection: None,
            comments: Vec::new(),
            tab_color: None,
            tables: vec![],
            header_footer: None,
            page_breaks: None,
            outline_properties: None,
            sparkline_groups: vec![],
            charts: vec![],
            pivot_tables: vec![],
            threaded_comments: vec![],
            slicers: Vec::new(),
            timelines: Vec::new(),
            preserved_extensions: vec![],
        };

        let xml = ws.to_xml().unwrap();
        let ws2 = WorksheetXml::parse(&xml).unwrap();
        assert_eq!(ws2.hyperlinks.len(), 2);
        assert_eq!(ws2.hyperlinks[0].cell_ref, "A1");
        assert_eq!(ws2.hyperlinks[0].location, Some("Sheet2!A1".to_string()));
        assert_eq!(ws2.hyperlinks[0].display, Some("Go to Sheet2".to_string()));
        assert_eq!(ws2.hyperlinks[0].tooltip, Some("Click here".to_string()));
        assert_eq!(ws2.hyperlinks[1].cell_ref, "B1");
        assert_eq!(ws2.hyperlinks[1].location, Some("Sheet3!B2".to_string()));
    }

    #[test]
    fn test_parse_page_setup() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData/>
  <pageSetup paperSize="9" orientation="landscape" scale="75" horizontalDpi="300" verticalDpi="300" fitToWidth="1" fitToHeight="0"/>
</worksheet>"#;

        let ws = WorksheetXml::parse(xml.as_bytes()).unwrap();
        let ps = ws.page_setup.as_ref().unwrap();
        assert_eq!(ps.paper_size, Some(9));
        assert_eq!(ps.orientation, Some("landscape".to_string()));
        assert_eq!(ps.scale, Some(75));
        assert_eq!(ps.horizontal_dpi, Some(300));
        assert_eq!(ps.vertical_dpi, Some(300));
        assert_eq!(ps.fit_to_width, Some(1));
        assert_eq!(ps.fit_to_height, Some(0));
    }

    #[test]
    fn test_page_setup_roundtrip() {
        let ws = WorksheetXml {
            dimension: None,
            rows: vec![],
            merge_cells: vec![],
            auto_filter: None,
            frozen_pane: None,
            split_pane: None,
            pane_selections: vec![],
            sheet_view: None,
            columns: vec![],
            data_validations: vec![],
            conditional_formatting: vec![],
            hyperlinks: vec![],
            page_setup: Some(PageSetup {
                paper_size: Some(1),
                orientation: Some("portrait".to_string()),
                fit_to_width: None,
                fit_to_height: None,
                scale: Some(100),
                first_page_number: None,
                horizontal_dpi: Some(600),
                vertical_dpi: Some(600),
            }),
            sheet_protection: None,
            comments: Vec::new(),
            tab_color: None,
            tables: vec![],
            header_footer: None,
            page_breaks: None,
            outline_properties: None,
            sparkline_groups: vec![],
            charts: vec![],
            pivot_tables: vec![],
            threaded_comments: vec![],
            slicers: Vec::new(),
            timelines: Vec::new(),
            preserved_extensions: vec![],
        };

        let xml = ws.to_xml().unwrap();
        let ws2 = WorksheetXml::parse(&xml).unwrap();
        let ps = ws2.page_setup.as_ref().unwrap();
        assert_eq!(ps.paper_size, Some(1));
        assert_eq!(ps.orientation, Some("portrait".to_string()));
        assert_eq!(ps.scale, Some(100));
        assert_eq!(ps.horizontal_dpi, Some(600));
        assert_eq!(ps.vertical_dpi, Some(600));
    }

    #[test]
    fn test_parse_sheet_protection() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData/>
  <sheetProtection sheet="1" objects="1" scenarios="1" password="ABCD" formatCells="1" insertRows="1" deleteRows="1" sort="1" autoFilter="1"/>
</worksheet>"#;

        let ws = WorksheetXml::parse(xml.as_bytes()).unwrap();
        let sp = ws.sheet_protection.as_ref().unwrap();
        assert!(sp.sheet);
        assert!(sp.objects);
        assert!(sp.scenarios);
        assert_eq!(sp.password, Some("ABCD".to_string()));
        assert!(sp.format_cells);
        assert!(!sp.format_columns);
        assert!(!sp.format_rows);
        assert!(!sp.insert_columns);
        assert!(sp.insert_rows);
        assert!(!sp.delete_columns);
        assert!(sp.delete_rows);
        assert!(sp.sort);
        assert!(sp.auto_filter);
    }

    #[test]
    fn test_sheet_protection_roundtrip() {
        let ws = WorksheetXml {
            dimension: None,
            rows: vec![],
            merge_cells: vec![],
            auto_filter: None,
            frozen_pane: None,
            split_pane: None,
            pane_selections: vec![],
            sheet_view: None,
            columns: vec![],
            data_validations: vec![],
            conditional_formatting: vec![],
            hyperlinks: vec![],
            page_setup: None,
            sheet_protection: Some(SheetProtection {
                sheet: true,
                objects: true,
                scenarios: false,
                password: Some("EFGH".to_string()),
                format_cells: false,
                format_columns: false,
                format_rows: false,
                insert_columns: false,
                insert_rows: true,
                delete_columns: false,
                delete_rows: true,
                sort: true,
                auto_filter: true,
            }),
            comments: Vec::new(),
            tab_color: None,
            tables: vec![],
            header_footer: None,
            page_breaks: None,
            outline_properties: None,
            sparkline_groups: vec![],
            charts: vec![],
            pivot_tables: vec![],
            threaded_comments: vec![],
            slicers: Vec::new(),
            timelines: Vec::new(),
            preserved_extensions: vec![],
        };

        let xml = ws.to_xml().unwrap();
        let ws2 = WorksheetXml::parse(&xml).unwrap();
        let sp = ws2.sheet_protection.as_ref().unwrap();
        assert!(sp.sheet);
        assert!(sp.objects);
        assert!(!sp.scenarios);
        assert_eq!(sp.password, Some("EFGH".to_string()));
        assert!(sp.insert_rows);
        assert!(sp.delete_rows);
        assert!(sp.sort);
        assert!(sp.auto_filter);
    }

    #[test]
    fn test_parse_array_formula() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1"><f t="array" ref="A1:A3">SUM(B1:B3*C1:C3)</f><v>100</v></c>
    </row>
  </sheetData>
</worksheet>"#;

        let ws = WorksheetXml::parse(xml.as_bytes()).unwrap();
        let cell = &ws.rows[0].cells[0];
        assert_eq!(cell.formula, Some("SUM(B1:B3*C1:C3)".to_string()));
        assert_eq!(cell.formula_type, Some("array".to_string()));
        assert_eq!(cell.formula_ref, Some("A1:A3".to_string()));
        assert_eq!(cell.shared_index, None);
        assert_eq!(cell.value, Some("100".to_string()));
    }

    #[test]
    fn test_parse_shared_formula() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1"><f t="shared" ref="A1:A3" si="0">B1+C1</f><v>10</v></c>
    </row>
    <row r="2">
      <c r="A2"><f t="shared" si="0"/><v>20</v></c>
    </row>
  </sheetData>
</worksheet>"#;

        let ws = WorksheetXml::parse(xml.as_bytes()).unwrap();
        let cell0 = &ws.rows[0].cells[0];
        assert_eq!(cell0.formula, Some("B1+C1".to_string()));
        assert_eq!(cell0.formula_type, Some("shared".to_string()));
        assert_eq!(cell0.formula_ref, Some("A1:A3".to_string()));
        assert_eq!(cell0.shared_index, Some(0));

        let cell1 = &ws.rows[1].cells[0];
        assert_eq!(cell1.formula, None);
        assert_eq!(cell1.formula_type, Some("shared".to_string()));
        assert_eq!(cell1.shared_index, Some(0));
    }

    #[test]
    fn test_array_formula_roundtrip() {
        let ws = WorksheetXml {
            dimension: None,
            rows: vec![Row {
                index: 1,
                cells: vec![Cell {
                    reference: "A1".to_string(),
                    cell_type: CellType::Number,
                    value: Some("100".to_string()),
                    formula: Some("SUM(B1:B3*C1:C3)".to_string()),
                    formula_type: Some("array".to_string()),
                    formula_ref: Some("A1:A3".to_string()),
                    ..Default::default()
                }],
                height: None,
                hidden: false,
                outline_level: None,
                collapsed: false,
            }],
            merge_cells: vec![],
            auto_filter: None,
            frozen_pane: None,
            split_pane: None,
            pane_selections: vec![],
            sheet_view: None,
            columns: vec![],
            data_validations: vec![],
            conditional_formatting: vec![],
            hyperlinks: vec![],
            page_setup: None,
            sheet_protection: None,
            comments: Vec::new(),
            tab_color: None,
            tables: vec![],
            header_footer: None,
            page_breaks: None,
            outline_properties: None,
            sparkline_groups: vec![],
            charts: vec![],
            pivot_tables: vec![],
            threaded_comments: vec![],
            slicers: Vec::new(),
            timelines: Vec::new(),
            preserved_extensions: vec![],
        };

        let xml = ws.to_xml().unwrap();
        let xml_str = String::from_utf8(xml.clone()).unwrap();
        assert!(xml_str.contains(r#"t="array""#));
        assert!(xml_str.contains(r#"ref="A1:A3""#));

        let ws2 = WorksheetXml::parse(&xml).unwrap();
        let cell = &ws2.rows[0].cells[0];
        assert_eq!(cell.formula_type, Some("array".to_string()));
        assert_eq!(cell.formula_ref, Some("A1:A3".to_string()));
        assert_eq!(cell.formula, Some("SUM(B1:B3*C1:C3)".to_string()));
    }

    #[test]
    fn test_parse_inline_string() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="inlineStr"><is><t>Hello World</t></is></c>
    </row>
  </sheetData>
</worksheet>"#;

        let ws = WorksheetXml::parse(xml.as_bytes()).unwrap();
        let cell = &ws.rows[0].cells[0];
        assert_eq!(cell.cell_type, CellType::InlineStr);
        assert_eq!(cell.value, Some("Hello World".to_string()));
        assert_eq!(cell.inline_string, Some("Hello World".to_string()));
    }

    #[test]
    fn test_parse_data_validation_prompts() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData/>
  <dataValidations count="1">
    <dataValidation type="list" allowBlank="1" showInputMessage="1" promptTitle="Choose" prompt="Pick a value" sqref="A1:A10">
      <formula1>"Yes,No"</formula1>
    </dataValidation>
  </dataValidations>
</worksheet>"#;

        let ws = WorksheetXml::parse(xml.as_bytes()).unwrap();
        let dv = &ws.data_validations[0];
        assert_eq!(dv.show_input_message, Some(true));
        assert_eq!(dv.prompt_title, Some("Choose".to_string()));
        assert_eq!(dv.prompt, Some("Pick a value".to_string()));
    }

    #[test]
    fn test_data_validation_prompts_roundtrip() {
        let ws = WorksheetXml {
            dimension: None,
            rows: vec![],
            merge_cells: vec![],
            auto_filter: None,
            frozen_pane: None,
            split_pane: None,
            pane_selections: vec![],
            sheet_view: None,
            columns: vec![],
            data_validations: vec![DataValidation {
                sqref: "A1:A10".to_string(),
                validation_type: Some("list".to_string()),
                operator: None,
                formula1: Some("\"Yes,No\"".to_string()),
                formula2: None,
                allow_blank: Some(true),
                show_error_message: None,
                error_title: None,
                error_message: None,
                show_input_message: Some(true),
                prompt_title: Some("Choose".to_string()),
                prompt: Some("Pick a value".to_string()),
            }],
            conditional_formatting: vec![],
            hyperlinks: vec![],
            page_setup: None,
            sheet_protection: None,
            comments: Vec::new(),
            tab_color: None,
            tables: vec![],
            header_footer: None,
            page_breaks: None,
            outline_properties: None,
            sparkline_groups: vec![],
            charts: vec![],
            pivot_tables: vec![],
            threaded_comments: vec![],
            slicers: Vec::new(),
            timelines: Vec::new(),
            preserved_extensions: vec![],
        };

        let xml = ws.to_xml().unwrap();
        let ws2 = WorksheetXml::parse(&xml).unwrap();
        let dv = &ws2.data_validations[0];
        assert_eq!(dv.show_input_message, Some(true));
        assert_eq!(dv.prompt_title, Some("Choose".to_string()));
        assert_eq!(dv.prompt, Some("Pick a value".to_string()));
    }

    #[test]
    fn test_parse_auto_filter_with_columns() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData/>
  <autoFilter ref="A1:D10">
    <filterColumn colId="0">
      <filters>
        <filter val="Apple"/>
        <filter val="Banana"/>
      </filters>
    </filterColumn>
    <filterColumn colId="2">
      <filters>
        <filter val="Red"/>
      </filters>
    </filterColumn>
  </autoFilter>
</worksheet>"#;

        let ws = WorksheetXml::parse(xml.as_bytes()).unwrap();
        let af = ws.auto_filter.as_ref().unwrap();
        assert_eq!(af.range, "A1:D10");
        assert_eq!(af.filter_columns.len(), 2);
        assert_eq!(af.filter_columns[0].col_id, 0);
        assert_eq!(af.filter_columns[0].filters, vec!["Apple", "Banana"]);
        assert_eq!(af.filter_columns[1].col_id, 2);
        assert_eq!(af.filter_columns[1].filters, vec!["Red"]);
    }

    #[test]
    fn test_auto_filter_with_columns_roundtrip() {
        let ws = WorksheetXml {
            dimension: None,
            rows: vec![],
            merge_cells: vec![],
            auto_filter: Some(AutoFilter {
                range: "A1:D10".to_string(),
                filter_columns: vec![
                    FilterColumn {
                        col_id: 0,
                        filters: vec!["Apple".to_string(), "Banana".to_string()],
                        custom_filters: None,
                    },
                    FilterColumn {
                        col_id: 2,
                        filters: vec!["Red".to_string()],
                        custom_filters: None,
                    },
                ],
            }),
            frozen_pane: None,
            split_pane: None,
            pane_selections: vec![],
            sheet_view: None,
            columns: vec![],
            data_validations: vec![],
            conditional_formatting: vec![],
            hyperlinks: vec![],
            page_setup: None,
            sheet_protection: None,
            comments: Vec::new(),
            tab_color: None,
            tables: vec![],
            header_footer: None,
            page_breaks: None,
            outline_properties: None,
            sparkline_groups: vec![],
            charts: vec![],
            pivot_tables: vec![],
            threaded_comments: vec![],
            slicers: Vec::new(),
            timelines: Vec::new(),
            preserved_extensions: vec![],
        };

        let xml = ws.to_xml().unwrap();
        let ws2 = WorksheetXml::parse(&xml).unwrap();
        let af = ws2.auto_filter.as_ref().unwrap();
        assert_eq!(af.range, "A1:D10");
        assert_eq!(af.filter_columns.len(), 2);
        assert_eq!(af.filter_columns[0].col_id, 0);
        assert_eq!(af.filter_columns[0].filters, vec!["Apple", "Banana"]);
        assert_eq!(af.filter_columns[1].col_id, 2);
        assert_eq!(af.filter_columns[1].filters, vec!["Red"]);
    }

    #[test]
    fn test_parse_auto_filter_custom_filters() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData/>
  <autoFilter ref="A1:D10">
    <filterColumn colId="1">
      <customFilters and="1">
        <customFilter operator="greaterThanOrEqual" val="10"/>
        <customFilter operator="lessThan" val="100"/>
      </customFilters>
    </filterColumn>
  </autoFilter>
</worksheet>"#;

        let ws = WorksheetXml::parse(xml.as_bytes()).unwrap();
        let af = ws.auto_filter.as_ref().unwrap();
        assert_eq!(af.range, "A1:D10");
        assert_eq!(af.filter_columns.len(), 1);
        let fc = &af.filter_columns[0];
        assert_eq!(fc.col_id, 1);
        assert!(fc.filters.is_empty());
        let cf = fc.custom_filters.as_ref().unwrap();
        assert!(cf.and_op);
        assert_eq!(cf.filters.len(), 2);
        assert_eq!(cf.filters[0].operator.as_deref(), Some("greaterThanOrEqual"));
        assert_eq!(cf.filters[0].val, "10");
        assert_eq!(cf.filters[1].operator.as_deref(), Some("lessThan"));
        assert_eq!(cf.filters[1].val, "100");
    }

    #[test]
    fn test_auto_filter_custom_filters_roundtrip() {
        let ws = WorksheetXml {
            dimension: None,
            rows: vec![],
            merge_cells: vec![],
            auto_filter: Some(AutoFilter {
                range: "B2:E20".to_string(),
                filter_columns: vec![
                    FilterColumn {
                        col_id: 0,
                        filters: vec!["Yes".to_string()],
                        custom_filters: None,
                    },
                    FilterColumn {
                        col_id: 3,
                        filters: vec![],
                        custom_filters: Some(CustomFilters {
                            and_op: false,
                            filters: vec![
                                CustomFilter { operator: Some("greaterThan".to_string()), val: "50".to_string() },
                            ],
                        }),
                    },
                    FilterColumn {
                        col_id: 1,
                        filters: vec![],
                        custom_filters: Some(CustomFilters {
                            and_op: true,
                            filters: vec![
                                CustomFilter { operator: Some("greaterThanOrEqual".to_string()), val: "5".to_string() },
                                CustomFilter { operator: Some("lessThanOrEqual".to_string()), val: "99".to_string() },
                            ],
                        }),
                    },
                ],
            }),
            frozen_pane: None,
            split_pane: None,
            pane_selections: vec![],
            sheet_view: None,
            columns: vec![],
            data_validations: vec![],
            conditional_formatting: vec![],
            hyperlinks: vec![],
            page_setup: None,
            sheet_protection: None,
            comments: Vec::new(),
            tab_color: None,
            tables: vec![],
            header_footer: None,
            page_breaks: None,
            outline_properties: None,
            sparkline_groups: vec![],
            charts: vec![],
            pivot_tables: vec![],
            threaded_comments: vec![],
            slicers: Vec::new(),
            timelines: Vec::new(),
            preserved_extensions: vec![],
        };

        let xml = ws.to_xml().unwrap();
        let ws2 = WorksheetXml::parse(&xml).unwrap();
        let af = ws2.auto_filter.as_ref().unwrap();
        assert_eq!(af.range, "B2:E20");
        assert_eq!(af.filter_columns.len(), 3);

        // First: regular filters.
        assert_eq!(af.filter_columns[0].col_id, 0);
        assert_eq!(af.filter_columns[0].filters, vec!["Yes"]);
        assert!(af.filter_columns[0].custom_filters.is_none());

        // Second: single custom filter (OR).
        let cf1 = af.filter_columns[1].custom_filters.as_ref().unwrap();
        assert!(!cf1.and_op);
        assert_eq!(cf1.filters.len(), 1);
        assert_eq!(cf1.filters[0].operator.as_deref(), Some("greaterThan"));
        assert_eq!(cf1.filters[0].val, "50");

        // Third: two custom filters (AND).
        let cf2 = af.filter_columns[2].custom_filters.as_ref().unwrap();
        assert!(cf2.and_op);
        assert_eq!(cf2.filters.len(), 2);
        assert_eq!(cf2.filters[0].operator.as_deref(), Some("greaterThanOrEqual"));
        assert_eq!(cf2.filters[0].val, "5");
        assert_eq!(cf2.filters[1].operator.as_deref(), Some("lessThanOrEqual"));
        assert_eq!(cf2.filters[1].val, "99");
    }

    #[test]
    fn test_auto_filter_custom_filter_no_operator() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData/>
  <autoFilter ref="A1:B5">
    <filterColumn colId="0">
      <customFilters>
        <customFilter val="*test*"/>
      </customFilters>
    </filterColumn>
  </autoFilter>
</worksheet>"#;

        let ws = WorksheetXml::parse(xml.as_bytes()).unwrap();
        let af = ws.auto_filter.as_ref().unwrap();
        let fc = &af.filter_columns[0];
        let cf = fc.custom_filters.as_ref().unwrap();
        assert!(!cf.and_op);
        assert_eq!(cf.filters.len(), 1);
        assert!(cf.filters[0].operator.is_none());
        assert_eq!(cf.filters[0].val, "*test*");

        // Roundtrip.
        let xml_out = ws.to_xml().unwrap();
        let ws2 = WorksheetXml::parse(&xml_out).unwrap();
        let cf2 = ws2.auto_filter.as_ref().unwrap().filter_columns[0]
            .custom_filters
            .as_ref()
            .unwrap();
        assert!(cf2.filters[0].operator.is_none());
        assert_eq!(cf2.filters[0].val, "*test*");
    }

    #[test]
    fn test_parse_color_scale() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData/>
  <conditionalFormatting sqref="A1:A10">
    <cfRule type="colorScale" priority="1">
      <colorScale>
        <cfvo type="min"/>
        <cfvo type="max"/>
        <color rgb="FFF8696B"/>
        <color rgb="FF63BE7B"/>
      </colorScale>
    </cfRule>
  </conditionalFormatting>
</worksheet>"#;

        let ws = WorksheetXml::parse(xml.as_bytes()).unwrap();
        let rule = &ws.conditional_formatting[0].rules[0];
        assert_eq!(rule.rule_type, "colorScale");
        let cs = rule.color_scale.as_ref().unwrap();
        assert_eq!(cs.cfvos.len(), 2);
        assert_eq!(cs.cfvos[0].cfvo_type, "min");
        assert_eq!(cs.cfvos[0].val, None);
        assert_eq!(cs.cfvos[1].cfvo_type, "max");
        assert_eq!(cs.colors, vec!["FFF8696B", "FF63BE7B"]);
    }

    #[test]
    fn test_color_scale_roundtrip() {
        let ws = WorksheetXml {
            dimension: None,
            rows: vec![],
            merge_cells: vec![],
            auto_filter: None,
            frozen_pane: None,
            split_pane: None,
            pane_selections: vec![],
            sheet_view: None,
            columns: vec![],
            data_validations: vec![],
            conditional_formatting: vec![ConditionalFormatting {
                sqref: "A1:A10".to_string(),
                rules: vec![ConditionalFormattingRule {
                    rule_type: "colorScale".to_string(),
                    priority: 1,
                    operator: None,
                    formula: None,
                    dxf_id: None,
                    color_scale: Some(ColorScale {
                        cfvos: vec![
                            Cfvo { cfvo_type: "min".to_string(), val: None },
                            Cfvo { cfvo_type: "max".to_string(), val: None },
                        ],
                        colors: vec!["FFF8696B".to_string(), "FF63BE7B".to_string()],
                    }),
                    data_bar: None,
                    icon_set: None,
                }],
            }],
            hyperlinks: vec![],
            page_setup: None,
            sheet_protection: None,
            comments: Vec::new(),
            tab_color: None,
            tables: vec![],
            header_footer: None,
            page_breaks: None,
            outline_properties: None,
            sparkline_groups: vec![],
            charts: vec![],
            pivot_tables: vec![],
            threaded_comments: vec![],
            slicers: Vec::new(),
            timelines: Vec::new(),
            preserved_extensions: vec![],
        };

        let xml = ws.to_xml().unwrap();
        let ws2 = WorksheetXml::parse(&xml).unwrap();
        let cs = ws2.conditional_formatting[0].rules[0].color_scale.as_ref().unwrap();
        assert_eq!(cs.cfvos.len(), 2);
        assert_eq!(cs.cfvos[0].cfvo_type, "min");
        assert_eq!(cs.cfvos[1].cfvo_type, "max");
        assert_eq!(cs.colors, vec!["FFF8696B", "FF63BE7B"]);
    }

    #[test]
    fn test_parse_data_bar() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData/>
  <conditionalFormatting sqref="B1:B10">
    <cfRule type="dataBar" priority="1">
      <dataBar>
        <cfvo type="min"/>
        <cfvo type="max"/>
        <color rgb="FF638EC6"/>
      </dataBar>
    </cfRule>
  </conditionalFormatting>
</worksheet>"#;

        let ws = WorksheetXml::parse(xml.as_bytes()).unwrap();
        let rule = &ws.conditional_formatting[0].rules[0];
        let db = rule.data_bar.as_ref().unwrap();
        assert_eq!(db.cfvos.len(), 2);
        assert_eq!(db.color, "FF638EC6");
    }

    #[test]
    fn test_parse_icon_set() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData/>
  <conditionalFormatting sqref="C1:C10">
    <cfRule type="iconSet" priority="1">
      <iconSet iconSet="3Arrows">
        <cfvo type="percent" val="0"/>
        <cfvo type="percent" val="33"/>
        <cfvo type="percent" val="67"/>
      </iconSet>
    </cfRule>
  </conditionalFormatting>
</worksheet>"#;

        let ws = WorksheetXml::parse(xml.as_bytes()).unwrap();
        let rule = &ws.conditional_formatting[0].rules[0];
        let is = rule.icon_set.as_ref().unwrap();
        assert_eq!(is.icon_set_type, Some("3Arrows".to_string()));
        assert_eq!(is.cfvos.len(), 3);
        assert_eq!(is.cfvos[0].cfvo_type, "percent");
        assert_eq!(is.cfvos[0].val, Some("0".to_string()));
        assert_eq!(is.cfvos[1].val, Some("33".to_string()));
        assert_eq!(is.cfvos[2].val, Some("67".to_string()));
    }

    #[test]
    fn test_icon_set_roundtrip() {
        let ws = WorksheetXml {
            dimension: None,
            rows: vec![],
            merge_cells: vec![],
            auto_filter: None,
            frozen_pane: None,
            split_pane: None,
            pane_selections: vec![],
            sheet_view: None,
            columns: vec![],
            data_validations: vec![],
            conditional_formatting: vec![ConditionalFormatting {
                sqref: "C1:C10".to_string(),
                rules: vec![ConditionalFormattingRule {
                    rule_type: "iconSet".to_string(),
                    priority: 1,
                    operator: None,
                    formula: None,
                    dxf_id: None,
                    color_scale: None,
                    data_bar: None,
                    icon_set: Some(IconSet {
                        icon_set_type: Some("3TrafficLights1".to_string()),
                        cfvos: vec![
                            Cfvo { cfvo_type: "percent".to_string(), val: Some("0".to_string()) },
                            Cfvo { cfvo_type: "percent".to_string(), val: Some("33".to_string()) },
                            Cfvo { cfvo_type: "percent".to_string(), val: Some("67".to_string()) },
                        ],
                    }),
                }],
            }],
            hyperlinks: vec![],
            page_setup: None,
            sheet_protection: None,
            comments: Vec::new(),
            tab_color: None,
            tables: vec![],
            header_footer: None,
            page_breaks: None,
            outline_properties: None,
            sparkline_groups: vec![],
            charts: vec![],
            pivot_tables: vec![],
            threaded_comments: vec![],
            slicers: Vec::new(),
            timelines: Vec::new(),
            preserved_extensions: vec![],
        };

        let xml = ws.to_xml().unwrap();
        let ws2 = WorksheetXml::parse(&xml).unwrap();
        let is = ws2.conditional_formatting[0].rules[0].icon_set.as_ref().unwrap();
        assert_eq!(is.icon_set_type, Some("3TrafficLights1".to_string()));
        assert_eq!(is.cfvos.len(), 3);
    }

    #[test]
    fn test_parse_header_footer() {
        let xml = r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData/>
  <headerFooter differentFirst="1">
    <oddHeader>&amp;CConfidential Report</oddHeader>
    <oddFooter>&amp;LPage &amp;P of &amp;N&amp;R&amp;D</oddFooter>
    <firstHeader>&amp;C&amp;BTitle Page</firstHeader>
  </headerFooter>
</worksheet>"#;
        let ws = WorksheetXml::parse(xml.as_bytes()).unwrap();
        let hf = ws.header_footer.as_ref().unwrap();
        assert_eq!(hf.odd_header.as_deref(), Some("&CConfidential Report"));
        assert_eq!(hf.odd_footer.as_deref(), Some("&LPage &P of &N&R&D"));
        assert_eq!(hf.first_header.as_deref(), Some("&C&BTitle Page"));
        assert!(hf.different_first);
        assert!(!hf.different_odd_even);
    }

    #[test]
    fn test_header_footer_roundtrip() {
        let ws = WorksheetXml {
            dimension: None,
            rows: vec![],
            merge_cells: vec![],
            auto_filter: None,
            frozen_pane: None,
            split_pane: None,
            pane_selections: vec![],
            sheet_view: None,
            columns: vec![],
            data_validations: vec![],
            conditional_formatting: vec![],
            hyperlinks: vec![],
            page_setup: None,
            sheet_protection: None,
            comments: vec![],
            tab_color: None,
            tables: vec![],
            header_footer: Some(HeaderFooter {
                odd_header: Some("&CPage &P of &N".into()),
                odd_footer: Some("&L&D&R&F".into()),
                ..Default::default()
            }),
            page_breaks: None,
            outline_properties: None,
            sparkline_groups: vec![],
            charts: vec![],
            pivot_tables: vec![],
            threaded_comments: vec![],
            slicers: Vec::new(),
            timelines: Vec::new(),
            preserved_extensions: vec![],
        };
        let xml = ws.to_xml_with_sst(None, &[], None).unwrap();
        let ws2 = WorksheetXml::parse(&xml).unwrap();
        let hf = ws2.header_footer.as_ref().unwrap();
        assert_eq!(hf.odd_header.as_deref(), Some("&CPage &P of &N"));
        assert_eq!(hf.odd_footer.as_deref(), Some("&L&D&R&F"));
    }

    #[test]
    fn test_row_outline_level_roundtrip() {
        let ws = WorksheetXml {
            dimension: None,
            rows: vec![
                Row { index: 1, cells: vec![], height: None, hidden: false, outline_level: None, collapsed: false },
                Row { index: 2, cells: vec![], height: None, hidden: false, outline_level: Some(1), collapsed: false },
                Row { index: 3, cells: vec![], height: None, hidden: false, outline_level: Some(1), collapsed: false },
                Row { index: 4, cells: vec![], height: None, hidden: false, outline_level: None, collapsed: false },
            ],
            merge_cells: vec![],
            auto_filter: None,
            frozen_pane: None,
            split_pane: None,
            pane_selections: vec![],
            sheet_view: None,
            columns: vec![],
            data_validations: vec![],
            conditional_formatting: vec![],
            hyperlinks: vec![],
            page_setup: None,
            sheet_protection: None,
            comments: vec![],
            tab_color: None,
            tables: vec![],
            header_footer: None,
            page_breaks: None,
            outline_properties: None,
            sparkline_groups: vec![],
            charts: vec![],
            pivot_tables: vec![],
            threaded_comments: vec![],
            slicers: Vec::new(),
            timelines: Vec::new(),
            preserved_extensions: vec![],
        };
        let xml = ws.to_xml_with_sst(None, &[], None).unwrap();
        let ws2 = WorksheetXml::parse(&xml).unwrap();
        assert_eq!(ws2.rows[1].outline_level, Some(1));
        assert_eq!(ws2.rows[2].outline_level, Some(1));
        assert!(ws2.rows[0].outline_level.is_none());
    }

    #[test]
    fn test_collapsed_row_group_roundtrip() {
        let xml = r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="2" outlineLevel="1" hidden="1"><c r="A2"><v>100</v></c></row>
    <row r="3" outlineLevel="1" hidden="1"><c r="A3"><v>200</v></c></row>
    <row r="4" collapsed="1"><c r="A4"><v>300</v></c></row>
  </sheetData>
</worksheet>"#;
        let ws = WorksheetXml::parse(xml.as_bytes()).unwrap();
        assert_eq!(ws.rows[0].outline_level, Some(1));
        assert!(ws.rows[0].hidden);
        assert!(ws.rows[2].collapsed);

        let xml2 = ws.to_xml_with_sst(None, &[], None).unwrap();
        let ws2 = WorksheetXml::parse(&xml2).unwrap();
        assert_eq!(ws2.rows[0].outline_level, Some(1));
        assert!(ws2.rows[0].hidden);
        assert!(ws2.rows[2].collapsed);
    }

    #[test]
    fn test_column_outline_level_roundtrip() {
        let ws = WorksheetXml {
            dimension: None,
            rows: vec![],
            merge_cells: vec![],
            auto_filter: None,
            frozen_pane: None,
            split_pane: None,
            pane_selections: vec![],
            sheet_view: None,
            columns: vec![
                ColumnInfo { min: 1, max: 1, width: 10.0, hidden: false, custom_width: true, outline_level: None, collapsed: false },
                ColumnInfo { min: 2, max: 3, width: 10.0, hidden: false, custom_width: true, outline_level: Some(1), collapsed: false },
                ColumnInfo { min: 4, max: 4, width: 12.0, hidden: false, custom_width: true, outline_level: None, collapsed: false },
            ],
            data_validations: vec![],
            conditional_formatting: vec![],
            hyperlinks: vec![],
            page_setup: None,
            sheet_protection: None,
            comments: vec![],
            tab_color: None,
            tables: vec![],
            header_footer: None,
            page_breaks: None,
            outline_properties: None,
            sparkline_groups: vec![],
            charts: vec![],
            pivot_tables: vec![],
            threaded_comments: vec![],
            slicers: Vec::new(),
            timelines: Vec::new(),
            preserved_extensions: vec![],
        };
        let xml = ws.to_xml_with_sst(None, &[], None).unwrap();
        let ws2 = WorksheetXml::parse(&xml).unwrap();
        assert!(ws2.columns[0].outline_level.is_none());
        assert_eq!(ws2.columns[1].outline_level, Some(1));
    }

    #[test]
    fn test_outline_properties_roundtrip() {
        let ws = WorksheetXml {
            dimension: None,
            rows: vec![],
            merge_cells: vec![],
            auto_filter: None,
            frozen_pane: None,
            split_pane: None,
            pane_selections: vec![],
            sheet_view: None,
            columns: vec![],
            data_validations: vec![],
            conditional_formatting: vec![],
            hyperlinks: vec![],
            page_setup: None,
            sheet_protection: None,
            comments: vec![],
            tab_color: None,
            tables: vec![],
            header_footer: None,
            page_breaks: None,
            outline_properties: Some(OutlineProperties {
                summary_below: false,
                summary_right: true,
            }),
            sparkline_groups: vec![],
            charts: vec![],
            pivot_tables: vec![],
            threaded_comments: vec![],
            slicers: Vec::new(),
            timelines: Vec::new(),
            preserved_extensions: vec![],
        };
        let xml = ws.to_xml_with_sst(None, &[], None).unwrap();
        let ws2 = WorksheetXml::parse(&xml).unwrap();
        let op = ws2.outline_properties.as_ref().unwrap();
        assert!(!op.summary_below);
        assert!(op.summary_right);
    }

    #[test]
    fn test_roundtrip_horizontal_split() {
        let ws = WorksheetXml {
            dimension: None,
            rows: vec![],
            merge_cells: vec![],
            auto_filter: None,
            frozen_pane: None,
            split_pane: Some(SplitPane {
                horizontal: Some(2400.0),
                vertical: None,
                top_left_cell: Some("A5".to_owned()),
                active_pane: Some("bottomLeft".to_owned()),
            }),
            pane_selections: vec![PaneSelection {
                pane: Some("bottomLeft".to_owned()),
                active_cell: Some("A5".to_owned()),
                sqref: Some("A5".to_owned()),
            }],
            sheet_view: None,
            columns: vec![],
            data_validations: vec![],
            conditional_formatting: vec![],
            hyperlinks: vec![],
            page_setup: None,
            sheet_protection: None,
            comments: vec![],
            tab_color: None,
            tables: vec![],
            header_footer: None,
            page_breaks: None,
            outline_properties: None,
            sparkline_groups: vec![],
            charts: vec![],
            pivot_tables: vec![],
            threaded_comments: vec![],
            slicers: Vec::new(),
            timelines: Vec::new(),
            preserved_extensions: vec![],
        };
        let xml = ws.to_xml().unwrap();
        let ws2 = WorksheetXml::parse(&xml).unwrap();
        assert!(ws2.frozen_pane.is_none());
        let sp = ws2.split_pane.as_ref().unwrap();
        assert!((sp.horizontal.unwrap() - 2400.0).abs() < f64::EPSILON);
        assert!(sp.vertical.is_none());
        assert_eq!(sp.top_left_cell.as_deref(), Some("A5"));
        assert_eq!(ws2.pane_selections.len(), 1);
    }

    #[test]
    fn test_roundtrip_four_way_split() {
        let ws = WorksheetXml {
            dimension: None,
            rows: vec![],
            merge_cells: vec![],
            auto_filter: None,
            frozen_pane: None,
            split_pane: Some(SplitPane {
                horizontal: Some(2400.0),
                vertical: Some(3000.0),
                top_left_cell: Some("D5".to_owned()),
                active_pane: Some("bottomRight".to_owned()),
            }),
            pane_selections: vec![
                PaneSelection { pane: Some("topRight".to_owned()), active_cell: Some("D1".to_owned()), sqref: Some("D1".to_owned()) },
                PaneSelection { pane: Some("bottomLeft".to_owned()), active_cell: Some("A5".to_owned()), sqref: Some("A5".to_owned()) },
                PaneSelection { pane: Some("bottomRight".to_owned()), active_cell: Some("D5".to_owned()), sqref: Some("D5".to_owned()) },
            ],
            sheet_view: None,
            columns: vec![],
            data_validations: vec![],
            conditional_formatting: vec![],
            hyperlinks: vec![],
            page_setup: None,
            sheet_protection: None,
            comments: vec![],
            tab_color: None,
            tables: vec![],
            header_footer: None,
            page_breaks: None,
            outline_properties: None,
            sparkline_groups: vec![],
            charts: vec![],
            pivot_tables: vec![],
            threaded_comments: vec![],
            slicers: Vec::new(),
            timelines: Vec::new(),
            preserved_extensions: vec![],
        };
        let xml = ws.to_xml().unwrap();
        let ws2 = WorksheetXml::parse(&xml).unwrap();
        let sp = ws2.split_pane.as_ref().unwrap();
        assert!((sp.vertical.unwrap() - 3000.0).abs() < f64::EPSILON);
        assert!((sp.horizontal.unwrap() - 2400.0).abs() < f64::EPSILON);
        assert_eq!(ws2.pane_selections.len(), 3);
    }

    #[test]
    fn test_split_pane_overrides_frozen_on_write() {
        let ws = WorksheetXml {
            dimension: None,
            rows: vec![],
            merge_cells: vec![],
            auto_filter: None,
            frozen_pane: Some(FrozenPane { rows: 1, cols: 0 }),
            split_pane: Some(SplitPane {
                horizontal: Some(2400.0),
                vertical: None,
                top_left_cell: Some("A5".to_owned()),
                active_pane: Some("bottomLeft".to_owned()),
            }),
            pane_selections: vec![],
            sheet_view: None,
            columns: vec![],
            data_validations: vec![],
            conditional_formatting: vec![],
            hyperlinks: vec![],
            page_setup: None,
            sheet_protection: None,
            comments: vec![],
            tab_color: None,
            tables: vec![],
            header_footer: None,
            page_breaks: None,
            outline_properties: None,
            sparkline_groups: vec![],
            charts: vec![],
            pivot_tables: vec![],
            threaded_comments: vec![],
            slicers: Vec::new(),
            timelines: Vec::new(),
            preserved_extensions: vec![],
        };
        let xml = ws.to_xml().unwrap();
        let xml_str = std::str::from_utf8(&xml).unwrap();
        assert!(xml_str.contains(r#"state="split""#));
        assert!(!xml_str.contains(r#"state="frozen""#));
    }

    #[test]
    fn test_parse_sheet_view_hide_gridlines_zoom() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetViews>
    <sheetView showGridLines="0" zoomScale="150" workbookViewId="0"/>
  </sheetViews>
  <sheetData/>
</worksheet>"#;
        let ws = WorksheetXml::parse(xml.as_bytes()).unwrap();
        let sv = ws.sheet_view.as_ref().unwrap();
        assert!(!sv.show_grid_lines);
        assert_eq!(sv.zoom_scale, Some(150));
    }

    #[test]
    fn test_parse_sheet_view_rtl() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetViews>
    <sheetView rightToLeft="1" showRowColHeaders="0" workbookViewId="0"/>
  </sheetViews>
  <sheetData/>
</worksheet>"#;
        let ws = WorksheetXml::parse(xml.as_bytes()).unwrap();
        let sv = ws.sheet_view.as_ref().unwrap();
        assert!(sv.right_to_left);
        assert!(!sv.show_row_col_headers);
    }

    #[test]
    fn test_parse_sheet_view_page_break_preview() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetViews>
    <sheetView view="pageBreakPreview" zoomScaleNormal="100" zoomScale="60" workbookViewId="0"/>
  </sheetViews>
  <sheetData/>
</worksheet>"#;
        let ws = WorksheetXml::parse(xml.as_bytes()).unwrap();
        let sv = ws.sheet_view.as_ref().unwrap();
        assert_eq!(sv.view.as_deref(), Some("pageBreakPreview"));
        assert_eq!(sv.zoom_scale, Some(60));
        assert_eq!(sv.zoom_scale_normal, Some(100));
    }

    #[test]
    fn test_parse_sheet_view_page_layout() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetViews>
    <sheetView view="pageLayout" showRuler="0" showWhiteSpace="0" zoomScalePageLayoutView="75" workbookViewId="0"/>
  </sheetViews>
  <sheetData/>
</worksheet>"#;
        let ws = WorksheetXml::parse(xml.as_bytes()).unwrap();
        let sv = ws.sheet_view.as_ref().unwrap();
        assert_eq!(sv.view.as_deref(), Some("pageLayout"));
        assert!(!sv.show_ruler);
        assert!(!sv.show_white_space);
        assert_eq!(sv.zoom_scale_page_layout_view, Some(75));
    }

    #[test]
    fn test_parse_sheet_view_defaults_only() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetViews>
    <sheetView workbookViewId="0"/>
  </sheetViews>
  <sheetData/>
</worksheet>"#;
        let ws = WorksheetXml::parse(xml.as_bytes()).unwrap();
        // All defaults -> sheet_view should be None (no non-default attributes found)
        assert!(ws.sheet_view.is_none());
    }

    #[test]
    fn test_parse_sheet_view_with_frozen_pane() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetViews>
    <sheetView showGridLines="0" zoomScale="120" workbookViewId="0">
      <pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>
    </sheetView>
  </sheetViews>
  <sheetData/>
</worksheet>"#;
        let ws = WorksheetXml::parse(xml.as_bytes()).unwrap();
        let sv = ws.sheet_view.as_ref().unwrap();
        assert!(!sv.show_grid_lines);
        assert_eq!(sv.zoom_scale, Some(120));
        assert!(ws.frozen_pane.is_some());
    }

    /// Helper to construct a default WorksheetXml for testing.
    fn default_test_ws() -> WorksheetXml {
        WorksheetXml {
            dimension: None,
            rows: vec![],
            merge_cells: vec![],
            auto_filter: None,
            frozen_pane: None,
            split_pane: None,
            pane_selections: vec![],
            sheet_view: None,
            columns: vec![],
            data_validations: vec![],
            conditional_formatting: vec![],
            hyperlinks: vec![],
            page_setup: None,
            sheet_protection: None,
            comments: vec![],
            tab_color: None,
            tables: vec![],
            header_footer: None,
            page_breaks: None,
            outline_properties: None,
            sparkline_groups: vec![],
            charts: vec![],
            pivot_tables: vec![],
            threaded_comments: vec![],
            slicers: Vec::new(),
            timelines: Vec::new(),
            preserved_extensions: vec![],
        }
    }

    #[test]
    fn test_roundtrip_sheet_view_hide_gridlines_zoom() {
        let mut ws = default_test_ws();
        ws.sheet_view = Some(SheetViewData {
            show_grid_lines: false,
            zoom_scale: Some(150),
            ..SheetViewData::default()
        });
        let xml = ws.to_xml().unwrap();
        let ws2 = WorksheetXml::parse(&xml).unwrap();
        let sv = ws2.sheet_view.as_ref().unwrap();
        assert!(!sv.show_grid_lines);
        assert_eq!(sv.zoom_scale, Some(150));
    }

    #[test]
    fn test_roundtrip_sheet_view_rtl() {
        let mut ws = default_test_ws();
        ws.sheet_view = Some(SheetViewData {
            right_to_left: true,
            show_row_col_headers: false,
            ..SheetViewData::default()
        });
        let xml = ws.to_xml().unwrap();
        let ws2 = WorksheetXml::parse(&xml).unwrap();
        let sv = ws2.sheet_view.as_ref().unwrap();
        assert!(sv.right_to_left);
        assert!(!sv.show_row_col_headers);
    }

    #[test]
    fn test_roundtrip_sheet_view_page_break_preview() {
        let mut ws = default_test_ws();
        ws.sheet_view = Some(SheetViewData {
            view: Some("pageBreakPreview".to_owned()),
            zoom_scale: Some(60),
            zoom_scale_normal: Some(100),
            ..SheetViewData::default()
        });
        let xml = ws.to_xml().unwrap();
        let ws2 = WorksheetXml::parse(&xml).unwrap();
        let sv = ws2.sheet_view.as_ref().unwrap();
        assert_eq!(sv.view.as_deref(), Some("pageBreakPreview"));
        assert_eq!(sv.zoom_scale, Some(60));
        assert_eq!(sv.zoom_scale_normal, Some(100));
    }

    #[test]
    fn test_roundtrip_sheet_view_with_frozen_pane() {
        let mut ws = default_test_ws();
        ws.sheet_view = Some(SheetViewData {
            show_grid_lines: false,
            zoom_scale: Some(120),
            ..SheetViewData::default()
        });
        ws.frozen_pane = Some(FrozenPane { rows: 1, cols: 0 });
        let xml = ws.to_xml().unwrap();
        let ws2 = WorksheetXml::parse(&xml).unwrap();
        assert!(ws2.sheet_view.is_some());
        assert!(ws2.frozen_pane.is_some());
        assert!(!ws2.sheet_view.unwrap().show_grid_lines);
    }

    #[test]
    fn test_sheet_view_defaults_minimal_xml() {
        let mut ws = default_test_ws();
        ws.sheet_view = Some(SheetViewData::default());
        let xml = ws.to_xml().unwrap();
        let xml_str = std::str::from_utf8(&xml).unwrap();
        // All defaults → no non-default attributes should appear
        assert!(!xml_str.contains("showGridLines"));
        assert!(!xml_str.contains("zoomScale"));
        assert!(!xml_str.contains("rightToLeft"));
        // But sheetViews should still be written
        assert!(xml_str.contains("sheetViews"));
        assert!(xml_str.contains("sheetView"));
    }

    // ---- Sparkline parser tests ----

    #[test]
    fn test_parse_line_sparkline() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData/>
  <extLst>
    <ext uri="{05C60535-1F16-4fd2-B633-F4F36011B0BD}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">
      <x14:sparklineGroups xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">
        <x14:sparklineGroup>
          <x14:sparklines>
            <x14:sparkline>
              <xm:f>Sheet1!A1:A10</xm:f>
              <xm:sqref>B1</xm:sqref>
            </x14:sparkline>
            <x14:sparkline>
              <xm:f>Sheet1!A1:A10</xm:f>
              <xm:sqref>B2</xm:sqref>
            </x14:sparkline>
          </x14:sparklines>
        </x14:sparklineGroup>
      </x14:sparklineGroups>
    </ext>
  </extLst>
</worksheet>"#;
        let ws = WorksheetXml::parse(xml).unwrap();
        assert_eq!(ws.sparkline_groups.len(), 1);
        assert_eq!(ws.sparkline_groups[0].sparkline_type, "line");
        assert_eq!(ws.sparkline_groups[0].sparklines.len(), 2);
        assert_eq!(ws.sparkline_groups[0].sparklines[0].formula, "Sheet1!A1:A10");
        assert_eq!(ws.sparkline_groups[0].sparklines[0].sqref, "B1");
    }

    #[test]
    fn test_parse_column_sparkline_with_colors() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData/>
  <extLst>
    <ext uri="{05C60535-1F16-4fd2-B633-F4F36011B0BD}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">
      <x14:sparklineGroups xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">
        <x14:sparklineGroup type="column" markers="1" high="1" low="1">
          <x14:colorSeries rgb="FF376092"/>
          <x14:colorNegative rgb="FFC00000"/>
          <x14:sparklines>
            <x14:sparkline>
              <xm:f>Sheet1!A1:A5</xm:f>
              <xm:sqref>C1</xm:sqref>
            </x14:sparkline>
          </x14:sparklines>
        </x14:sparklineGroup>
      </x14:sparklineGroups>
    </ext>
  </extLst>
</worksheet>"#;
        let ws = WorksheetXml::parse(xml).unwrap();
        assert_eq!(ws.sparkline_groups[0].sparkline_type, "column");
        assert!(ws.sparkline_groups[0].markers);
        assert!(ws.sparkline_groups[0].high);
        assert!(ws.sparkline_groups[0].low);
        assert_eq!(ws.sparkline_groups[0].color_series.as_deref(), Some("FF376092"));
        assert_eq!(ws.sparkline_groups[0].color_negative.as_deref(), Some("FFC00000"));
    }

    #[test]
    fn test_parse_sparkline_with_options() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData/>
  <extLst>
    <ext uri="{05C60535-1F16-4fd2-B633-F4F36011B0BD}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">
      <x14:sparklineGroups xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">
        <x14:sparklineGroup type="stacked" displayEmptyCellsAs="gap" lineWeight="1.25" displayXAxis="1" negative="1" first="1" last="1" manualMin="0" manualMax="100">
          <x14:sparklines>
            <x14:sparkline>
              <xm:f>Sheet1!D1:D10</xm:f>
              <xm:sqref>E1</xm:sqref>
            </x14:sparkline>
          </x14:sparklines>
        </x14:sparklineGroup>
      </x14:sparklineGroups>
    </ext>
  </extLst>
</worksheet>"#;
        let ws = WorksheetXml::parse(xml).unwrap();
        let g = &ws.sparkline_groups[0];
        assert_eq!(g.sparkline_type, "stacked");
        assert_eq!(g.display_empty_cells_as.as_deref(), Some("gap"));
        assert_eq!(g.line_weight, Some(1.25));
        assert!(g.display_x_axis);
        assert!(g.negative);
        assert!(g.first);
        assert!(g.last);
        assert_eq!(g.manual_min, Some(0.0));
        assert_eq!(g.manual_max, Some(100.0));
    }

    #[test]
    fn test_parse_no_sparklines() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData/>
</worksheet>"#;
        let ws = WorksheetXml::parse(xml).unwrap();
        assert!(ws.sparkline_groups.is_empty());
    }

    // ---- Sparkline writer roundtrip tests ----

    #[test]
    fn test_sparkline_line_roundtrip() {
        let mut ws = default_test_ws();
        ws.sparkline_groups = vec![SparklineGroup {
            sparkline_type: "line".into(),
            sparklines: vec![
                Sparkline { formula: "Sheet1!A1:A10".into(), sqref: "B1".into() },
                Sparkline { formula: "Sheet1!A1:A10".into(), sqref: "B2".into() },
            ],
            ..Default::default()
        }];
        let xml = ws.to_xml_with_sst(None, &[], None).unwrap();
        let ws2 = WorksheetXml::parse(&xml).unwrap();
        assert_eq!(ws2.sparkline_groups.len(), 1);
        assert_eq!(ws2.sparkline_groups[0].sparklines.len(), 2);
        assert_eq!(ws2.sparkline_groups[0].sparklines[0].formula, "Sheet1!A1:A10");
        assert_eq!(ws2.sparkline_groups[0].sparklines[0].sqref, "B1");
        assert_eq!(ws2.sparkline_groups[0].sparklines[1].sqref, "B2");
        assert_eq!(ws2.sparkline_groups[0].sparkline_type, "line");
    }

    #[test]
    fn test_sparkline_column_with_colors_roundtrip() {
        let mut ws = default_test_ws();
        ws.sparkline_groups = vec![SparklineGroup {
            sparkline_type: "column".into(),
            sparklines: vec![
                Sparkline { formula: "Sheet1!C1:C5".into(), sqref: "D1".into() },
            ],
            color_series: Some("FF376092".into()),
            color_negative: Some("FFC00000".into()),
            markers: true,
            high: true,
            low: true,
            ..Default::default()
        }];
        let xml = ws.to_xml_with_sst(None, &[], None).unwrap();
        let ws2 = WorksheetXml::parse(&xml).unwrap();
        assert_eq!(ws2.sparkline_groups[0].sparkline_type, "column");
        assert_eq!(ws2.sparkline_groups[0].color_series.as_deref(), Some("FF376092"));
        assert_eq!(ws2.sparkline_groups[0].color_negative.as_deref(), Some("FFC00000"));
        assert!(ws2.sparkline_groups[0].markers);
        assert!(ws2.sparkline_groups[0].high);
        assert!(ws2.sparkline_groups[0].low);
    }

    #[test]
    fn test_multiple_sparkline_groups_roundtrip() {
        let mut ws = default_test_ws();
        ws.sparkline_groups = vec![
            SparklineGroup {
                sparkline_type: "line".into(),
                sparklines: vec![
                    Sparkline { formula: "Sheet1!A1:A10".into(), sqref: "B1".into() },
                ],
                ..Default::default()
            },
            SparklineGroup {
                sparkline_type: "column".into(),
                sparklines: vec![
                    Sparkline { formula: "Sheet1!C1:C10".into(), sqref: "D1".into() },
                ],
                negative: true,
                display_empty_cells_as: Some("gap".into()),
                ..Default::default()
            },
        ];
        let xml = ws.to_xml_with_sst(None, &[], None).unwrap();
        let ws2 = WorksheetXml::parse(&xml).unwrap();
        assert_eq!(ws2.sparkline_groups.len(), 2);
        assert_eq!(ws2.sparkline_groups[0].sparkline_type, "line");
        assert_eq!(ws2.sparkline_groups[1].sparkline_type, "column");
        assert!(ws2.sparkline_groups[1].negative);
        assert_eq!(
            ws2.sparkline_groups[1].display_empty_cells_as.as_deref(),
            Some("gap")
        );
    }

    #[test]
    fn test_sparkline_all_options_roundtrip() {
        let mut ws = default_test_ws();
        ws.sparkline_groups = vec![SparklineGroup {
            sparkline_type: "stacked".into(),
            sparklines: vec![
                Sparkline { formula: "Sheet1!D1:D10".into(), sqref: "E1".into() },
            ],
            color_series: Some("FF376092".into()),
            color_negative: Some("FFC00000".into()),
            color_axis: Some("FF000000".into()),
            color_markers: Some("FF0070C0".into()),
            color_first: Some("FF00B050".into()),
            color_last: Some("FFFFC000".into()),
            color_high: Some("FF00B0F0".into()),
            color_low: Some("FFFF0000".into()),
            line_weight: Some(1.25),
            markers: true,
            high: true,
            low: true,
            first: true,
            last: true,
            negative: true,
            display_x_axis: true,
            display_empty_cells_as: Some("gap".into()),
            manual_min: Some(0.0),
            manual_max: Some(100.0),
            right_to_left: true,
        }];
        let xml = ws.to_xml_with_sst(None, &[], None).unwrap();
        let ws2 = WorksheetXml::parse(&xml).unwrap();
        let g = &ws2.sparkline_groups[0];
        assert_eq!(g.sparkline_type, "stacked");
        assert_eq!(g.color_series.as_deref(), Some("FF376092"));
        assert_eq!(g.color_negative.as_deref(), Some("FFC00000"));
        assert_eq!(g.color_axis.as_deref(), Some("FF000000"));
        assert_eq!(g.color_markers.as_deref(), Some("FF0070C0"));
        assert_eq!(g.color_first.as_deref(), Some("FF00B050"));
        assert_eq!(g.color_last.as_deref(), Some("FFFFC000"));
        assert_eq!(g.color_high.as_deref(), Some("FF00B0F0"));
        assert_eq!(g.color_low.as_deref(), Some("FFFF0000"));
        assert_eq!(g.line_weight, Some(1.25));
        assert!(g.markers);
        assert!(g.high);
        assert!(g.low);
        assert!(g.first);
        assert!(g.last);
        assert!(g.negative);
        assert!(g.display_x_axis);
        assert_eq!(g.display_empty_cells_as.as_deref(), Some("gap"));
        assert_eq!(g.manual_min, Some(0.0));
        assert_eq!(g.manual_max, Some(100.0));
        assert!(g.right_to_left);
    }

    #[test]
    fn test_parse_data_table_formula() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="2">
      <c r="B2">
        <f t="dataTable" r1="B1" r2="A2"/>
        <v>42</v>
      </c>
    </row>
  </sheetData>
</worksheet>"#;
        let ws = WorksheetXml::parse(xml).unwrap();
        let cell = &ws.rows[0].cells[0];
        assert_eq!(cell.formula_type.as_deref(), Some("dataTable"));
        assert_eq!(cell.formula_r1.as_deref(), Some("B1"));
        assert_eq!(cell.formula_r2.as_deref(), Some("A2"));
    }

    #[test]
    fn test_parse_2d_data_table() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="2">
      <c r="B2">
        <f t="dataTable" r1="B1" r2="A2" dt2D="1" dtr1="1" dtr2="1"/>
        <v>99</v>
      </c>
    </row>
  </sheetData>
</worksheet>"#;
        let ws = WorksheetXml::parse(xml).unwrap();
        let cell = &ws.rows[0].cells[0];
        assert_eq!(cell.formula_dt2d, Some(true));
        assert_eq!(cell.formula_dtr1, Some(true));
        assert_eq!(cell.formula_dtr2, Some(true));
    }

    #[test]
    fn test_data_table_formula_roundtrip() {
        let mut ws = default_test_ws();
        ws.rows.push(Row {
            index: 2,
            cells: vec![Cell {
                reference: "B2".into(),
                cell_type: CellType::Number,
                value: Some("42".into()),
                formula: Some(String::new()),
                formula_type: Some("dataTable".into()),
                formula_r1: Some("B1".into()),
                formula_r2: Some("A2".into()),
                formula_dt2d: Some(true),
                ..Default::default()
            }],
            height: None,
            hidden: false,
            outline_level: None,
            collapsed: false,
        });
        let xml = ws.to_xml_with_sst(None, &[], None).unwrap();
        let ws2 = WorksheetXml::parse(&xml).unwrap();
        let cell = &ws2.rows[0].cells[0];
        assert_eq!(cell.formula_type.as_deref(), Some("dataTable"));
        assert_eq!(cell.formula_r1.as_deref(), Some("B1"));
        assert_eq!(cell.formula_r2.as_deref(), Some("A2"));
        assert_eq!(cell.formula_dt2d, Some(true));
    }

    #[test]
    fn test_normal_formula_no_data_table_attrs() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1">
        <f>SUM(B1:B10)</f>
        <v>55</v>
      </c>
    </row>
  </sheetData>
</worksheet>"#;
        let ws = WorksheetXml::parse(xml).unwrap();
        let cell = &ws.rows[0].cells[0];
        assert!(cell.formula_r1.is_none());
        assert!(cell.formula_r2.is_none());
        assert!(cell.formula_dt2d.is_none());
    }

    // ---- Page breaks tests ----

    #[test]
    fn test_parse_row_breaks() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData/>
  <rowBreaks count="2" manualBreakCount="2">
    <brk id="10" max="16383" man="1"/>
    <brk id="20" max="16383" man="1"/>
  </rowBreaks>
</worksheet>"#;
        let ws = WorksheetXml::parse(xml).unwrap();
        let pb = ws.page_breaks.as_ref().unwrap();
        assert_eq!(pb.row_breaks.len(), 2);
        assert_eq!(pb.row_breaks[0].id, 10);
        assert_eq!(pb.row_breaks[0].max, Some(16383));
        assert!(pb.row_breaks[0].man);
        assert_eq!(pb.row_breaks[1].id, 20);
        assert!(pb.col_breaks.is_empty());
    }

    #[test]
    fn test_parse_col_breaks() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData/>
  <colBreaks count="1" manualBreakCount="1">
    <brk id="5" max="1048575" man="1"/>
  </colBreaks>
</worksheet>"#;
        let ws = WorksheetXml::parse(xml).unwrap();
        let pb = ws.page_breaks.as_ref().unwrap();
        assert!(pb.row_breaks.is_empty());
        assert_eq!(pb.col_breaks.len(), 1);
        assert_eq!(pb.col_breaks[0].id, 5);
        assert_eq!(pb.col_breaks[0].max, Some(1048575));
        assert!(pb.col_breaks[0].man);
    }

    #[test]
    fn test_parse_both_row_and_col_breaks() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData/>
  <rowBreaks count="1" manualBreakCount="1">
    <brk id="10" max="16383" man="1"/>
  </rowBreaks>
  <colBreaks count="1" manualBreakCount="1">
    <brk id="3" max="1048575" man="1"/>
  </colBreaks>
</worksheet>"#;
        let ws = WorksheetXml::parse(xml).unwrap();
        let pb = ws.page_breaks.as_ref().unwrap();
        assert_eq!(pb.row_breaks.len(), 1);
        assert_eq!(pb.col_breaks.len(), 1);
        assert_eq!(pb.row_breaks[0].id, 10);
        assert_eq!(pb.col_breaks[0].id, 3);
    }

    #[test]
    fn test_roundtrip_page_breaks() {
        let mut ws = default_test_ws();
        ws.page_breaks = Some(PageBreaks {
            row_breaks: vec![
                PageBreak { id: 10, min: None, max: Some(16383), man: true },
                PageBreak { id: 25, min: Some(0), max: Some(16383), man: true },
            ],
            col_breaks: vec![
                PageBreak { id: 5, min: None, max: Some(1048575), man: true },
            ],
        });

        let xml = ws.to_xml().unwrap();
        let xml_str = std::str::from_utf8(&xml).unwrap();
        assert!(xml_str.contains("<rowBreaks"));
        assert!(xml_str.contains("<colBreaks"));
        assert!(xml_str.contains("manualBreakCount=\"2\""));

        let ws2 = WorksheetXml::parse(&xml).unwrap();
        let pb = ws2.page_breaks.as_ref().unwrap();
        assert_eq!(pb.row_breaks.len(), 2);
        assert_eq!(pb.row_breaks[0].id, 10);
        assert_eq!(pb.row_breaks[0].max, Some(16383));
        assert!(pb.row_breaks[0].man);
        assert_eq!(pb.row_breaks[1].id, 25);
        assert_eq!(pb.row_breaks[1].min, Some(0));
        assert_eq!(pb.col_breaks.len(), 1);
        assert_eq!(pb.col_breaks[0].id, 5);
    }

    #[test]
    fn test_no_page_breaks_by_default() {
        let ws = default_test_ws();
        let xml = ws.to_xml().unwrap();
        let xml_str = std::str::from_utf8(&xml).unwrap();
        assert!(!xml_str.contains("rowBreaks"));
        assert!(!xml_str.contains("colBreaks"));
    }
}
