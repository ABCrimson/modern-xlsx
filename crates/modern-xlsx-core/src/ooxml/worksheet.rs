use quick_xml::events::{BytesDecl, BytesEnd, BytesStart, BytesText, Event};
use quick_xml::{Reader, Writer};

use super::push_entity;
use serde::{Deserialize, Serialize};

use crate::{ModernXlsxError, Result};

use super::SPREADSHEET_NS;

fn is_false(v: &bool) -> bool {
    !v
}

fn default_true_hf() -> bool {
    true
}

fn is_true(v: &bool) -> bool {
    *v
}

/// Public wrapper for JSON escaping, used by `reader::read_xlsx_json`.
pub fn json_escape_to_pub(out: &mut String, s: &str) {
    json_escape_to(out, s);
}

/// Write a JSON-escaped string to the output buffer.
/// Handles `"`, `\`, and control characters (0x00-0x1F).
fn json_escape_to(out: &mut String, s: &str) {
    for c in s.chars() {
        match c {
            '"' => out.push_str("\\\""),
            '\\' => out.push_str("\\\\"),
            '\n' => out.push_str("\\n"),
            '\r' => out.push_str("\\r"),
            '\t' => out.push_str("\\t"),
            c if (c as u32) < 0x20 => {
                // Control characters → \uXXXX
                out.push_str("\\u00");
                let hi = (c as u8) >> 4;
                let lo = (c as u8) & 0x0F;
                out.push(char::from(if hi < 10 { b'0' + hi } else { b'a' + hi - 10 }));
                out.push(char::from(if lo < 10 { b'0' + lo } else { b'a' + lo - 10 }));
            }
            c => out.push(c),
        }
    }
}

/// Convert a `CellType` to its camelCase JSON string (matching serde rename).
fn cell_type_json_str(ct: CellType) -> &'static str {
    match ct {
        CellType::SharedString => "sharedString",
        CellType::Number => "number",
        CellType::Boolean => "boolean",
        CellType::Error => "error",
        CellType::FormulaStr => "formulaStr",
        CellType::InlineStr => "inlineStr",
        CellType::Stub => "stub",
    }
}

/// Write an f64 to the JSON output buffer matching serde_json's formatting.
fn write_f64_json(out: &mut String, v: f64) {
    use std::fmt::Write;
    if !v.is_finite() {
        out.push_str("null");
        return;
    }
    // serde_json formats floats without trailing zeros for integers.
    // Writing to String is infallible — the fmt::Write impl never returns Err.
    if v == v.floor() && v.abs() < 1e15 {
        let _ = write!(out, "{}.0", v as i64);
    } else {
        let _ = write!(out, "{v}");
    }
}

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
    #[serde(default, skip_serializing_if = "is_false")]
    pub different_odd_even: bool,
    #[serde(default, skip_serializing_if = "is_false")]
    pub different_first: bool,
    #[serde(default = "default_true_hf", skip_serializing_if = "is_true")]
    pub scale_with_doc: bool,
    #[serde(default = "default_true_hf", skip_serializing_if = "is_true")]
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
    #[serde(default = "default_true_hf")]
    pub summary_below: bool,
    #[serde(default = "default_true_hf")]
    pub summary_right: bool,
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
    pub outline_properties: Option<OutlineProperties>,
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
    #[serde(default, skip_serializing_if = "is_false")]
    pub hidden: bool,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub outline_level: Option<u8>,
    #[serde(default, skip_serializing_if = "is_false")]
    pub collapsed: bool,
}

/// A single cell in a row.
#[derive(Debug, Clone, Serialize, Deserialize)]
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
    /// Formula type: "array" or "shared", from `t` attribute on `<f>`.
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
}

/// The type of a cell, determined by the `t` attribute on the `<c>` element.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum CellType {
    /// `t="s"` — value is a shared string table index.
    SharedString,
    /// `t="n"` or omitted — numeric value.
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
    #[serde(default = "default_true_hf", skip_serializing_if = "is_true")]
    pub show_grid_lines: bool,
    /// Whether row and column headers are visible (default: true).
    #[serde(default = "default_true_hf", skip_serializing_if = "is_true")]
    pub show_row_col_headers: bool,
    /// Whether zero values are displayed (default: true).
    #[serde(default = "default_true_hf", skip_serializing_if = "is_true")]
    pub show_zeros: bool,
    /// Right-to-left display mode (default: false).
    #[serde(default, skip_serializing_if = "is_false")]
    pub right_to_left: bool,
    /// Whether this sheet tab is selected (default: false).
    #[serde(default, skip_serializing_if = "is_false")]
    pub tab_selected: bool,
    /// Whether the ruler is shown in Page Layout view (default: true).
    #[serde(default = "default_true_hf", skip_serializing_if = "is_true")]
    pub show_ruler: bool,
    /// Whether outline (grouping) symbols are shown (default: true).
    #[serde(default = "default_true_hf", skip_serializing_if = "is_true")]
    pub show_outline_symbols: bool,
    /// Whether white space around the page is shown in Page Layout view (default: true).
    #[serde(default = "default_true_hf", skip_serializing_if = "is_true")]
    pub show_white_space: bool,
    /// Whether the default grid color is used (default: true).
    #[serde(default = "default_true_hf", skip_serializing_if = "is_true")]
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
    #[serde(default, skip_serializing_if = "is_false")]
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
    pub filters: Vec<String>,
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
enum ParseState {
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
    InColorScale,
    InDataBar,
    InIconSet,
    HeaderFooter,
    HeaderFooterChild(u8), // 0=oddHeader, 1=oddFooter, 2=evenHeader, 3=evenFooter, 4=firstHeader, 5=firstFooter
}

impl WorksheetXml {
    /// Parse a worksheet XML file from raw bytes.
    pub fn parse(data: &[u8]) -> Result<Self> {
        Self::parse_with_sst(data, None)
    }

    /// Parse a worksheet XML file, optionally resolving shared string indices
    /// inline during parsing. This avoids a costly post-parse pass and
    /// eliminates intermediate index-string allocations for SharedString cells.
    pub fn parse_with_sst(data: &[u8], sst: Option<&super::shared_strings::SharedStringTable>) -> Result<Self> {
        let mut reader = Reader::from_reader(data);
        reader.config_mut().trim_text(true);

        // Pre-allocate buffers proportional to input size for fewer reallocations.
        let estimated_rows = (data.len() / 200).max(64); // ~200 bytes per row heuristic
        let mut buf = Vec::with_capacity(512);

        let mut dimension: Option<String> = None;
        let mut rows: Vec<Row> = Vec::with_capacity(estimated_rows);
        let mut merge_cells: Vec<String> = Vec::new();
        let mut auto_filter: Option<AutoFilter> = None;
        let mut frozen_pane: Option<FrozenPane> = None;
        let mut split_pane = None::<SplitPane>;
        let mut pane_selections = Vec::<PaneSelection>::new();
        let mut columns: Vec<ColumnInfo> = Vec::new();
        let mut data_validations: Vec<DataValidation> = Vec::new();
        let mut conditional_formatting: Vec<ConditionalFormatting> = Vec::new();
        let mut hyperlinks: Vec<Hyperlink> = Vec::new();
        let mut page_setup: Option<PageSetup> = None;
        let mut sheet_protection: Option<SheetProtection> = None;
        let mut tab_color: Option<String> = None;
        let mut header_footer: Option<HeaderFooter> = None;
        let mut hf_text_buf = String::new();
        let mut outline_properties: Option<OutlineProperties> = None;

        let mut state = ParseState::Root;

        // Current row being built.
        let mut cur_row_index: u32 = 0;
        let mut cur_row_height: Option<f64> = None;
        let mut cur_row_hidden = false;
        let mut cur_row_outline_level: Option<u8> = None;
        let mut cur_row_collapsed = false;
        let mut cur_row_cells: Vec<Cell> = Vec::with_capacity(16);

        // Current cell being built.
        let mut cur_cell_ref = String::new();
        let mut cur_cell_type = CellType::Number;
        let mut cur_cell_style: Option<u32> = None;
        let mut cur_cell_value: Option<String> = None;
        let mut cur_cell_formula: Option<String> = None;
        let mut cur_cell_formula_type: Option<String> = None;
        let mut cur_cell_formula_ref: Option<String> = None;
        let mut cur_cell_shared_index: Option<u32> = None;
        let mut cur_cell_inline_string: Option<String> = None;
        let mut cur_cell_dynamic_array: Option<bool> = None;

        // Buffers for text content.
        let mut text_buf = String::with_capacity(256);

        // Current data validation being built.
        let mut cur_dv: Option<DataValidation> = None;

        // Current conditional formatting being built.
        let mut cur_cf_sqref = String::new();
        let mut cur_cf_rules: Vec<ConditionalFormattingRule> = Vec::new();
        let mut cur_cf_rule: Option<ConditionalFormattingRule> = None;

        // Current auto filter being built.
        let mut cur_af_range = String::new();
        let mut cur_af_columns: Vec<FilterColumn> = Vec::new();
        let mut cur_filter_col_id: u32 = 0;
        let mut cur_filter_vals: Vec<String> = Vec::new();

        // Conditional formatting sub-element state.
        let mut cur_cfvos: Vec<Cfvo> = Vec::new();
        let mut cur_cf_colors: Vec<String> = Vec::new();
        let mut cur_cf_bar_color = String::new();
        let mut cur_icon_set_type: Option<String> = None;

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) => {
                    let local = e.local_name();
                    match (state, local.as_ref()) {
                        (ParseState::Root, b"sheetPr") => {
                            state = ParseState::SheetPr;
                        }
                        (ParseState::Root, b"sheetViews") => {
                            state = ParseState::SheetViews;
                        }
                        (ParseState::SheetViews, b"sheetView") => {
                            state = ParseState::SheetView;
                        }
                        (ParseState::SheetView, b"selection") => {
                            // Handle <selection> as Start element (non-self-closing).
                            let mut sel_pane = None::<String>;
                            let mut sel_active = None::<String>;
                            let mut sel_sqref = None::<String>;
                            for attr in e.attributes().flatten() {
                                let ln = attr.key.local_name();
                                match ln.as_ref() {
                                    b"pane" => {
                                        sel_pane = Some(std::str::from_utf8(&attr.value).unwrap_or_default().to_owned());
                                    }
                                    b"activeCell" => {
                                        sel_active = Some(std::str::from_utf8(&attr.value).unwrap_or_default().to_owned());
                                    }
                                    b"sqref" => {
                                        sel_sqref = Some(std::str::from_utf8(&attr.value).unwrap_or_default().to_owned());
                                    }
                                    _ => {}
                                }
                            }
                            pane_selections.push(PaneSelection {
                                pane: sel_pane,
                                active_cell: sel_active,
                                sqref: sel_sqref,
                            });
                        }
                        (ParseState::Root, b"cols") => {
                            state = ParseState::Cols;
                        }
                        (ParseState::Root, b"sheetData") => {
                            state = ParseState::SheetData;
                        }
                        (ParseState::SheetData, b"row") => {
                            state = ParseState::InRow;
                            cur_row_index = 0;
                            cur_row_height = None;
                            cur_row_hidden = false;
                            cur_row_outline_level = None;
                            cur_row_collapsed = false;
                            cur_row_cells.clear();

                            for attr in e.attributes().flatten() {
                                let ln = attr.key.local_name();
                                match ln.as_ref() {
                                    b"r" => {
                                        let val = std::str::from_utf8(&attr.value).unwrap_or_default();
                                        cur_row_index = val.parse::<u32>().unwrap_or(0);
                                    }
                                    b"ht" => {
                                        let val = std::str::from_utf8(&attr.value).unwrap_or_default();
                                        cur_row_height = val.parse::<f64>().ok();
                                    }
                                    b"hidden" => {
                                        let val = std::str::from_utf8(&attr.value).unwrap_or_default();
                                        cur_row_hidden = val == "1" || val.eq_ignore_ascii_case("true");
                                    }
                                    b"outlineLevel" => {
                                        cur_row_outline_level = std::str::from_utf8(&attr.value)
                                            .ok()
                                            .and_then(|v| v.parse::<u8>().ok())
                                            .filter(|&v| v > 0);
                                    }
                                    b"collapsed" => {
                                        cur_row_collapsed = std::str::from_utf8(&attr.value).unwrap_or("0") == "1";
                                    }
                                    _ => {}
                                }
                            }
                        }
                        (ParseState::InRow, b"c") => {
                            state = ParseState::InCell;
                            cur_cell_ref.clear();
                            cur_cell_type = CellType::Number;
                            cur_cell_style = None;
                            cur_cell_value = None;
                            cur_cell_formula = None;
                            cur_cell_formula_type = None;
                            cur_cell_formula_ref = None;
                            cur_cell_shared_index = None;
                            cur_cell_inline_string = None;
                            cur_cell_dynamic_array = None;

                            for attr in e.attributes().flatten() {
                                let ln = attr.key.local_name();
                                match ln.as_ref() {
                                    b"r" => {
                                        cur_cell_ref =
                                            std::str::from_utf8(&attr.value).unwrap_or_default().to_owned();
                                    }
                                    b"t" => {
                                        let val = std::str::from_utf8(&attr.value).unwrap_or_default();
                                        cur_cell_type = match val {
                                            "s" => CellType::SharedString,
                                            "b" => CellType::Boolean,
                                            "e" => CellType::Error,
                                            "str" => CellType::FormulaStr,
                                            "inlineStr" => CellType::InlineStr,
                                            _ => CellType::Number,
                                        };
                                    }
                                    b"s" => {
                                        let val = std::str::from_utf8(&attr.value).unwrap_or_default();
                                        cur_cell_style = val.parse::<u32>().ok();
                                    }
                                    _ => {}
                                }
                            }
                        }
                        (ParseState::InCell, b"v") => {
                            state = ParseState::InCellValue;
                            text_buf.clear();
                        }
                        (ParseState::InCell, b"f") => {
                            state = ParseState::InCellFormula;
                            text_buf.clear();
                            for attr in e.attributes().flatten() {
                                let ln = attr.key.local_name();
                                let val = std::str::from_utf8(&attr.value).unwrap_or_default();
                                match ln.as_ref() {
                                    b"t" => cur_cell_formula_type = Some(val.to_owned()),
                                    b"ref" => cur_cell_formula_ref = Some(val.to_owned()),
                                    b"si" => cur_cell_shared_index = val.parse::<u32>().ok(),
                                    b"cm" => {
                                        if val == "1" {
                                            cur_cell_dynamic_array = Some(true);
                                        }
                                    }
                                    _ => {}
                                }
                            }
                        }
                        (ParseState::InCell, b"is") => {
                            state = ParseState::InInlineStr;
                        }
                        (ParseState::InInlineStr, b"t") => {
                            state = ParseState::InInlineStrT;
                            text_buf.clear();
                        }
                        (ParseState::Root, b"mergeCells") => {
                            state = ParseState::MergeCells;
                        }
                        (ParseState::Root, b"dataValidations") => {
                            state = ParseState::InDataValidations;
                        }
                        (ParseState::InDataValidations, b"dataValidation") => {
                            state = ParseState::InDataValidation;
                            let mut dv = DataValidation {
                                sqref: String::new(),
                                validation_type: None,
                                operator: None,
                                formula1: None,
                                formula2: None,
                                allow_blank: None,
                                show_error_message: None,
                                error_title: None,
                                error_message: None,
                                show_input_message: None,
                                prompt_title: None,
                                prompt: None,
                            };
                            for attr in e.attributes().flatten() {
                                let ln = attr.key.local_name();
                                let val = std::str::from_utf8(&attr.value).unwrap_or_default();
                                match ln.as_ref() {
                                    b"sqref" => dv.sqref = val.to_owned(),
                                    b"type" => dv.validation_type = Some(val.to_owned()),
                                    b"operator" => dv.operator = Some(val.to_owned()),
                                    b"allowBlank" => {
                                        dv.allow_blank = Some(val == "1" || val.eq_ignore_ascii_case("true"));
                                    }
                                    b"showErrorMessage" => {
                                        dv.show_error_message = Some(val == "1" || val.eq_ignore_ascii_case("true"));
                                    }
                                    b"errorTitle" => dv.error_title = Some(val.to_owned()),
                                    b"error" => dv.error_message = Some(val.to_owned()),
                                    b"showInputMessage" => {
                                        dv.show_input_message = Some(val == "1" || val.eq_ignore_ascii_case("true"));
                                    }
                                    b"promptTitle" => dv.prompt_title = Some(val.to_owned()),
                                    b"prompt" => dv.prompt = Some(val.to_owned()),
                                    _ => {}
                                }
                            }
                            cur_dv = Some(dv);
                        }
                        (ParseState::InDataValidation, b"formula1") => {
                            state = ParseState::InDVFormula1;
                            text_buf.clear();
                        }
                        (ParseState::InDataValidation, b"formula2") => {
                            state = ParseState::InDVFormula2;
                            text_buf.clear();
                        }
                        (ParseState::Root, b"conditionalFormatting") => {
                            state = ParseState::InConditionalFormatting;
                            cur_cf_sqref.clear();
                            cur_cf_rules.clear();
                            for attr in e.attributes().flatten() {
                                if attr.key.local_name().as_ref() == b"sqref" {
                                    cur_cf_sqref = std::str::from_utf8(&attr.value)
                                        .unwrap_or_default()
                                        .to_owned();
                                }
                            }
                        }
                        (ParseState::InConditionalFormatting, b"cfRule") => {
                            state = ParseState::InCfRule;
                            let mut rule = ConditionalFormattingRule {
                                rule_type: String::new(),
                                priority: 0,
                                operator: None,
                                formula: None,
                                dxf_id: None,
                                color_scale: None,
                                data_bar: None,
                                icon_set: None,
                            };
                            for attr in e.attributes().flatten() {
                                let ln = attr.key.local_name();
                                let val =
                                    std::str::from_utf8(&attr.value).unwrap_or_default();
                                match ln.as_ref() {
                                    b"type" => rule.rule_type = val.to_owned(),
                                    b"priority" => {
                                        rule.priority = val.parse::<u32>().unwrap_or(0);
                                    }
                                    b"operator" => rule.operator = Some(val.to_owned()),
                                    b"dxfId" => {
                                        rule.dxf_id = val.parse::<u32>().ok();
                                    }
                                    _ => {}
                                }
                            }
                            cur_cf_rule = Some(rule);
                        }
                        (ParseState::InCfRule, b"formula") => {
                            state = ParseState::InCfRuleFormula;
                            text_buf.clear();
                        }
                        (ParseState::InCfRule, b"colorScale") => {
                            state = ParseState::InColorScale;
                            cur_cfvos.clear();
                            cur_cf_colors.clear();
                        }
                        (ParseState::InCfRule, b"dataBar") => {
                            state = ParseState::InDataBar;
                            cur_cfvos.clear();
                            cur_cf_bar_color.clear();
                        }
                        (ParseState::InCfRule, b"iconSet") => {
                            state = ParseState::InIconSet;
                            cur_cfvos.clear();
                            cur_icon_set_type = None;
                            for attr in e.attributes().flatten() {
                                if attr.key.local_name().as_ref() == b"iconSet" {
                                    cur_icon_set_type = Some(
                                        std::str::from_utf8(&attr.value)
                                            .unwrap_or_default()
                                            .to_owned(),
                                    );
                                }
                            }
                        }
                        (ParseState::Root, b"hyperlinks") => {
                            state = ParseState::InHyperlinks;
                        }
                        (ParseState::Root, b"autoFilter") => {
                            state = ParseState::InAutoFilter;
                            cur_af_range.clear();
                            cur_af_columns.clear();
                            for attr in e.attributes().flatten() {
                                if attr.key.local_name().as_ref() == b"ref" {
                                    cur_af_range = std::str::from_utf8(&attr.value)
                                        .unwrap_or_default()
                                        .to_owned();
                                }
                            }
                        }
                        (ParseState::InAutoFilter, b"filterColumn") => {
                            state = ParseState::InFilterColumn;
                            cur_filter_col_id = 0;
                            cur_filter_vals.clear();
                            for attr in e.attributes().flatten() {
                                if attr.key.local_name().as_ref() == b"colId" {
                                    cur_filter_col_id = std::str::from_utf8(&attr.value)
                                        .unwrap_or_default()
                                        .parse::<u32>()
                                        .unwrap_or(0);
                                }
                            }
                        }
                        (ParseState::InFilterColumn, b"filters") => {
                            state = ParseState::InFilters;
                        }
                        (ParseState::Root, b"headerFooter") => {
                            let mut hf = HeaderFooter::default();
                            for attr in e.attributes().flatten() {
                                match attr.key.as_ref() {
                                    b"differentOddEven" => hf.different_odd_even = std::str::from_utf8(&attr.value).unwrap_or("0") == "1",
                                    b"differentFirst" => hf.different_first = std::str::from_utf8(&attr.value).unwrap_or("0") == "1",
                                    b"scaleWithDoc" => hf.scale_with_doc = std::str::from_utf8(&attr.value).unwrap_or("1") != "0",
                                    b"alignWithMargins" => hf.align_with_margins = std::str::from_utf8(&attr.value).unwrap_or("1") != "0",
                                    _ => {}
                                }
                            }
                            header_footer = Some(hf);
                            state = ParseState::HeaderFooter;
                        }
                        (ParseState::HeaderFooter, b"oddHeader") => { hf_text_buf.clear(); reader.config_mut().trim_text(false); state = ParseState::HeaderFooterChild(0); }
                        (ParseState::HeaderFooter, b"oddFooter") => { hf_text_buf.clear(); reader.config_mut().trim_text(false); state = ParseState::HeaderFooterChild(1); }
                        (ParseState::HeaderFooter, b"evenHeader") => { hf_text_buf.clear(); reader.config_mut().trim_text(false); state = ParseState::HeaderFooterChild(2); }
                        (ParseState::HeaderFooter, b"evenFooter") => { hf_text_buf.clear(); reader.config_mut().trim_text(false); state = ParseState::HeaderFooterChild(3); }
                        (ParseState::HeaderFooter, b"firstHeader") => { hf_text_buf.clear(); reader.config_mut().trim_text(false); state = ParseState::HeaderFooterChild(4); }
                        (ParseState::HeaderFooter, b"firstFooter") => { hf_text_buf.clear(); reader.config_mut().trim_text(false); state = ParseState::HeaderFooterChild(5); }
                        _ => {}
                    }
                }
                Ok(Event::Empty(ref e)) => {
                    let local = e.local_name();
                    match (state, local.as_ref()) {
                        (ParseState::Root, b"dimension") => {
                            for attr in e.attributes().flatten() {
                                if attr.key.local_name().as_ref() == b"ref" {
                                    dimension = Some(
                                        std::str::from_utf8(&attr.value).unwrap_or_default().to_owned(),
                                    );
                                }
                            }
                        }
                        (ParseState::SheetPr, b"tabColor") => {
                            for attr in e.attributes().flatten() {
                                if attr.key.local_name().as_ref() == b"rgb" {
                                    tab_color = Some(
                                        std::str::from_utf8(&attr.value).unwrap_or_default().to_owned(),
                                    );
                                }
                            }
                        }
                        (ParseState::SheetPr, b"outlinePr") => {
                            let mut op = OutlineProperties { summary_below: true, summary_right: true };
                            for attr in e.attributes().flatten() {
                                match attr.key.as_ref() {
                                    b"summaryBelow" => op.summary_below = std::str::from_utf8(&attr.value).unwrap_or("1") != "0",
                                    b"summaryRight" => op.summary_right = std::str::from_utf8(&attr.value).unwrap_or("1") != "0",
                                    _ => {}
                                }
                            }
                            outline_properties = Some(op);
                        }
                        (ParseState::Root, b"autoFilter") => {
                            let mut af_range = String::new();
                            for attr in e.attributes().flatten() {
                                if attr.key.local_name().as_ref() == b"ref" {
                                    af_range = std::str::from_utf8(&attr.value)
                                        .unwrap_or_default()
                                        .to_owned();
                                }
                            }
                            auto_filter = Some(AutoFilter {
                                range: af_range,
                                filter_columns: Vec::new(),
                            });
                        }
                        (ParseState::SheetView, b"pane") => {
                            let mut y_split_raw = String::new();
                            let mut x_split_raw = String::new();
                            let mut state_val = String::new();
                            let mut top_left_cell_val = None::<String>;
                            let mut active_pane_val = None::<String>;

                            for attr in e.attributes().flatten() {
                                let ln = attr.key.local_name();
                                match ln.as_ref() {
                                    b"ySplit" => {
                                        y_split_raw = std::str::from_utf8(&attr.value).unwrap_or_default().to_owned();
                                    }
                                    b"xSplit" => {
                                        x_split_raw = std::str::from_utf8(&attr.value).unwrap_or_default().to_owned();
                                    }
                                    b"state" => {
                                        state_val = std::str::from_utf8(&attr.value).unwrap_or_default().to_owned();
                                    }
                                    b"topLeftCell" => {
                                        top_left_cell_val = Some(std::str::from_utf8(&attr.value).unwrap_or_default().to_owned());
                                    }
                                    b"activePane" => {
                                        active_pane_val = Some(std::str::from_utf8(&attr.value).unwrap_or_default().to_owned());
                                    }
                                    _ => {}
                                }
                            }

                            if state_val == "frozen" {
                                let y: u32 = y_split_raw.parse().unwrap_or(0);
                                let x: u32 = x_split_raw.parse().unwrap_or(0);
                                if y > 0 || x > 0 {
                                    frozen_pane = Some(FrozenPane { rows: y, cols: x });
                                }
                            } else {
                                let y: f64 = y_split_raw.parse().unwrap_or(0.0);
                                let x: f64 = x_split_raw.parse().unwrap_or(0.0);
                                if y > 0.0 || x > 0.0 {
                                    split_pane = Some(SplitPane {
                                        horizontal: if y > 0.0 { Some(y) } else { None },
                                        vertical: if x > 0.0 { Some(x) } else { None },
                                        top_left_cell: top_left_cell_val,
                                        active_pane: active_pane_val,
                                    });
                                }
                            }
                        }
                        (ParseState::SheetView, b"selection") => {
                            let mut sel_pane = None::<String>;
                            let mut sel_active = None::<String>;
                            let mut sel_sqref = None::<String>;
                            for attr in e.attributes().flatten() {
                                let ln = attr.key.local_name();
                                match ln.as_ref() {
                                    b"pane" => {
                                        sel_pane = Some(std::str::from_utf8(&attr.value).unwrap_or_default().to_owned());
                                    }
                                    b"activeCell" => {
                                        sel_active = Some(std::str::from_utf8(&attr.value).unwrap_or_default().to_owned());
                                    }
                                    b"sqref" => {
                                        sel_sqref = Some(std::str::from_utf8(&attr.value).unwrap_or_default().to_owned());
                                    }
                                    _ => {}
                                }
                            }
                            pane_selections.push(PaneSelection {
                                pane: sel_pane,
                                active_cell: sel_active,
                                sqref: sel_sqref,
                            });
                        }
                        (ParseState::Cols, b"col") => {
                            columns.push(parse_col_element(e));
                        }
                        (ParseState::MergeCells, b"mergeCell") => {
                            for attr in e.attributes().flatten() {
                                if attr.key.local_name().as_ref() == b"ref" {
                                    merge_cells.push(
                                        std::str::from_utf8(&attr.value).unwrap_or_default().to_owned(),
                                    );
                                }
                            }
                        }
                        (ParseState::InRow, b"c") => {
                            // Self-closing <c ... /> — cell with no child elements.
                            let mut cell_ref = String::new();
                            let mut cell_type = CellType::Number;
                            let mut cell_style: Option<u32> = None;

                            for attr in e.attributes().flatten() {
                                let ln = attr.key.local_name();
                                match ln.as_ref() {
                                    b"r" => {
                                        cell_ref =
                                            std::str::from_utf8(&attr.value).unwrap_or_default().to_owned();
                                    }
                                    b"t" => {
                                        let val = std::str::from_utf8(&attr.value).unwrap_or_default();
                                        cell_type = match val {
                                            "s" => CellType::SharedString,
                                            "b" => CellType::Boolean,
                                            "e" => CellType::Error,
                                            "str" => CellType::FormulaStr,
                                            "inlineStr" => CellType::InlineStr,
                                            _ => CellType::Number,
                                        };
                                    }
                                    b"s" => {
                                        let val = std::str::from_utf8(&attr.value).unwrap_or_default();
                                        cell_style = val.parse::<u32>().ok();
                                    }
                                    _ => {}
                                }
                            }
                            cur_row_cells.push(Cell {
                                reference: cell_ref,
                                cell_type,
                                style_index: cell_style,
                                value: None,
                                formula: None,
                                formula_type: None,
                                formula_ref: None,
                                shared_index: None,
                                inline_string: None,
                                dynamic_array: None,
                            });
                        }
                        (ParseState::SheetData, b"row") => {
                            // Self-closing <row ... /> — empty row.
                            let mut row_index: u32 = 0;
                            let mut row_height: Option<f64> = None;
                            let mut row_hidden = false;
                            let mut row_ol: Option<u8> = None;
                            let mut row_coll = false;

                            for attr in e.attributes().flatten() {
                                let ln = attr.key.local_name();
                                match ln.as_ref() {
                                    b"r" => {
                                        let val = std::str::from_utf8(&attr.value).unwrap_or_default();
                                        row_index = val.parse::<u32>().unwrap_or(0);
                                    }
                                    b"ht" => {
                                        let val = std::str::from_utf8(&attr.value).unwrap_or_default();
                                        row_height = val.parse::<f64>().ok();
                                    }
                                    b"hidden" => {
                                        let val = std::str::from_utf8(&attr.value).unwrap_or_default();
                                        row_hidden = val == "1" || val.eq_ignore_ascii_case("true");
                                    }
                                    b"outlineLevel" => {
                                        row_ol = std::str::from_utf8(&attr.value)
                                            .ok()
                                            .and_then(|v| v.parse::<u8>().ok())
                                            .filter(|&v| v > 0);
                                    }
                                    b"collapsed" => {
                                        row_coll = std::str::from_utf8(&attr.value).unwrap_or("0") == "1";
                                    }
                                    _ => {}
                                }
                            }
                            rows.push(Row {
                                index: row_index,
                                cells: Vec::new(),
                                height: row_height,
                                hidden: row_hidden,
                                outline_level: row_ol,
                                collapsed: row_coll,
                            });
                        }
                        (ParseState::InCell, b"v") => {
                            // Empty <v/> — no value.
                        }
                        (ParseState::InCell, b"f") => {
                            // Empty <f/> — parse attributes even on self-closing.
                            for attr in e.attributes().flatten() {
                                let ln = attr.key.local_name();
                                let val = std::str::from_utf8(&attr.value).unwrap_or_default();
                                match ln.as_ref() {
                                    b"t" => cur_cell_formula_type = Some(val.to_owned()),
                                    b"ref" => cur_cell_formula_ref = Some(val.to_owned()),
                                    b"si" => cur_cell_shared_index = val.parse::<u32>().ok(),
                                    b"cm" => {
                                        if val == "1" {
                                            cur_cell_dynamic_array = Some(true);
                                        }
                                    }
                                    _ => {}
                                }
                            }
                        }
                        (ParseState::InDataValidations, b"dataValidation") => {
                            // Self-closing <dataValidation ... /> with no formula children.
                            let mut dv = DataValidation {
                                sqref: String::new(),
                                validation_type: None,
                                operator: None,
                                formula1: None,
                                formula2: None,
                                allow_blank: None,
                                show_error_message: None,
                                error_title: None,
                                error_message: None,
                                show_input_message: None,
                                prompt_title: None,
                                prompt: None,
                            };
                            for attr in e.attributes().flatten() {
                                let ln = attr.key.local_name();
                                let val = std::str::from_utf8(&attr.value).unwrap_or_default();
                                match ln.as_ref() {
                                    b"sqref" => dv.sqref = val.to_owned(),
                                    b"type" => dv.validation_type = Some(val.to_owned()),
                                    b"operator" => dv.operator = Some(val.to_owned()),
                                    b"allowBlank" => {
                                        dv.allow_blank = Some(val == "1" || val.eq_ignore_ascii_case("true"));
                                    }
                                    b"showErrorMessage" => {
                                        dv.show_error_message = Some(val == "1" || val.eq_ignore_ascii_case("true"));
                                    }
                                    b"errorTitle" => dv.error_title = Some(val.to_owned()),
                                    b"error" => dv.error_message = Some(val.to_owned()),
                                    b"showInputMessage" => {
                                        dv.show_input_message = Some(val == "1" || val.eq_ignore_ascii_case("true"));
                                    }
                                    b"promptTitle" => dv.prompt_title = Some(val.to_owned()),
                                    b"prompt" => dv.prompt = Some(val.to_owned()),
                                    _ => {}
                                }
                            }
                            data_validations.push(dv);
                        }
                        (ParseState::InConditionalFormatting, b"cfRule") => {
                            // Self-closing <cfRule ... /> with no formula child.
                            let mut rule = ConditionalFormattingRule {
                                rule_type: String::new(),
                                priority: 0,
                                operator: None,
                                formula: None,
                                dxf_id: None,
                                color_scale: None,
                                data_bar: None,
                                icon_set: None,
                            };
                            for attr in e.attributes().flatten() {
                                let ln = attr.key.local_name();
                                let val =
                                    std::str::from_utf8(&attr.value).unwrap_or_default();
                                match ln.as_ref() {
                                    b"type" => rule.rule_type = val.to_owned(),
                                    b"priority" => {
                                        rule.priority = val.parse::<u32>().unwrap_or(0);
                                    }
                                    b"operator" => rule.operator = Some(val.to_owned()),
                                    b"dxfId" => {
                                        rule.dxf_id = val.parse::<u32>().ok();
                                    }
                                    _ => {}
                                }
                            }
                            cur_cf_rules.push(rule);
                        }
                        // Hyperlink (self-closing).
                        (ParseState::InHyperlinks, b"hyperlink") => {
                            let mut hl = Hyperlink {
                                cell_ref: String::new(),
                                location: None,
                                display: None,
                                tooltip: None,
                            };
                            for attr in e.attributes().flatten() {
                                let ln = attr.key.local_name();
                                let val = std::str::from_utf8(&attr.value).unwrap_or_default();
                                match ln.as_ref() {
                                    b"ref" => hl.cell_ref = val.to_owned(),
                                    b"location" => hl.location = Some(val.to_owned()),
                                    b"display" => hl.display = Some(val.to_owned()),
                                    b"tooltip" => hl.tooltip = Some(val.to_owned()),
                                    _ => {}
                                }
                            }
                            hyperlinks.push(hl);
                        }
                        // pageSetup (always self-closing).
                        (ParseState::Root, b"pageSetup") => {
                            let mut ps = PageSetup {
                                paper_size: None,
                                orientation: None,
                                fit_to_width: None,
                                fit_to_height: None,
                                scale: None,
                                first_page_number: None,
                                horizontal_dpi: None,
                                vertical_dpi: None,
                            };
                            for attr in e.attributes().flatten() {
                                let ln = attr.key.local_name();
                                let val = std::str::from_utf8(&attr.value).unwrap_or_default();
                                match ln.as_ref() {
                                    b"paperSize" => ps.paper_size = val.parse::<u32>().ok(),
                                    b"orientation" => ps.orientation = Some(val.to_owned()),
                                    b"fitToWidth" => ps.fit_to_width = val.parse::<u32>().ok(),
                                    b"fitToHeight" => ps.fit_to_height = val.parse::<u32>().ok(),
                                    b"scale" => ps.scale = val.parse::<u32>().ok(),
                                    b"firstPageNumber" => ps.first_page_number = val.parse::<u32>().ok(),
                                    b"horizontalDpi" => ps.horizontal_dpi = val.parse::<u32>().ok(),
                                    b"verticalDpi" => ps.vertical_dpi = val.parse::<u32>().ok(),
                                    _ => {}
                                }
                            }
                            page_setup = Some(ps);
                        }
                        // sheetProtection (always self-closing).
                        (ParseState::Root, b"sheetProtection") => {
                            let mut sp = SheetProtection {
                                sheet: false,
                                objects: false,
                                scenarios: false,
                                password: None,
                                format_cells: false,
                                format_columns: false,
                                format_rows: false,
                                insert_columns: false,
                                insert_rows: false,
                                delete_columns: false,
                                delete_rows: false,
                                sort: false,
                                auto_filter: false,
                            };
                            for attr in e.attributes().flatten() {
                                let ln = attr.key.local_name();
                                let val = std::str::from_utf8(&attr.value).unwrap_or_default();
                                let as_bool = val == "1" || val.eq_ignore_ascii_case("true");
                                match ln.as_ref() {
                                    b"sheet" => sp.sheet = as_bool,
                                    b"objects" => sp.objects = as_bool,
                                    b"scenarios" => sp.scenarios = as_bool,
                                    b"password" => sp.password = Some(val.to_owned()),
                                    b"formatCells" => sp.format_cells = as_bool,
                                    b"formatColumns" => sp.format_columns = as_bool,
                                    b"formatRows" => sp.format_rows = as_bool,
                                    b"insertColumns" => sp.insert_columns = as_bool,
                                    b"insertRows" => sp.insert_rows = as_bool,
                                    b"deleteColumns" => sp.delete_columns = as_bool,
                                    b"deleteRows" => sp.delete_rows = as_bool,
                                    b"sort" => sp.sort = as_bool,
                                    b"autoFilter" => sp.auto_filter = as_bool,
                                    _ => {}
                                }
                            }
                            sheet_protection = Some(sp);
                        }
                        // cfvo in colorScale/dataBar/iconSet.
                        (ParseState::InColorScale | ParseState::InDataBar | ParseState::InIconSet, b"cfvo") => {
                            let mut cfvo_type = String::new();
                            let mut cfvo_val: Option<String> = None;
                            for attr in e.attributes().flatten() {
                                let ln = attr.key.local_name();
                                let v = std::str::from_utf8(&attr.value).unwrap_or_default();
                                match ln.as_ref() {
                                    b"type" => cfvo_type = v.to_owned(),
                                    b"val" => cfvo_val = Some(v.to_owned()),
                                    _ => {}
                                }
                            }
                            cur_cfvos.push(Cfvo { cfvo_type, val: cfvo_val });
                        }
                        // color in colorScale.
                        (ParseState::InColorScale, b"color") => {
                            for attr in e.attributes().flatten() {
                                if attr.key.local_name().as_ref() == b"rgb" {
                                    cur_cf_colors.push(
                                        std::str::from_utf8(&attr.value)
                                            .unwrap_or_default()
                                            .to_owned(),
                                    );
                                }
                            }
                        }
                        // color in dataBar.
                        (ParseState::InDataBar, b"color") => {
                            for attr in e.attributes().flatten() {
                                if attr.key.local_name().as_ref() == b"rgb" {
                                    cur_cf_bar_color = std::str::from_utf8(&attr.value)
                                        .unwrap_or_default()
                                        .to_owned();
                                }
                            }
                        }
                        // filter values.
                        (ParseState::InFilters, b"filter") => {
                            for attr in e.attributes().flatten() {
                                if attr.key.local_name().as_ref() == b"val" {
                                    cur_filter_vals.push(
                                        std::str::from_utf8(&attr.value)
                                            .unwrap_or_default()
                                            .to_owned(),
                                    );
                                }
                            }
                        }
                        _ => {}
                    }
                }
                Ok(Event::Text(ref e)) => {
                    match state {
                        ParseState::InCellValue
                        | ParseState::InCellFormula
                        | ParseState::InInlineStrT
                        | ParseState::InDVFormula1
                        | ParseState::InDVFormula2
                        | ParseState::InCfRuleFormula => {
                            text_buf.push_str(std::str::from_utf8(e.as_ref()).unwrap_or_default());
                        }
                        ParseState::HeaderFooterChild(_) => {
                            hf_text_buf.push_str(std::str::from_utf8(e.as_ref()).unwrap_or_default());
                        }
                        _ => {}
                    }
                }
                Ok(Event::GeneralRef(ref e)) => {
                    match state {
                        ParseState::InCellValue
                        | ParseState::InCellFormula
                        | ParseState::InInlineStrT
                        | ParseState::InDVFormula1
                        | ParseState::InDVFormula2
                        | ParseState::InCfRuleFormula => {
                            push_entity(&mut text_buf, e.as_ref());
                        }
                        ParseState::HeaderFooterChild(_) => {
                            push_entity(&mut hf_text_buf, e.as_ref());
                        }
                        _ => {}
                    }
                }
                Ok(Event::End(ref e)) => {
                    let local = e.local_name();
                    match (state, local.as_ref()) {
                        (ParseState::InCellValue, b"v") => {
                            // If this is a SharedString cell and we have the SST,
                            // resolve the index directly instead of storing the raw index.
                            // This avoids an extra pass AND eliminates the intermediate
                            // index-string allocation.
                            if cur_cell_type == CellType::SharedString
                                && let Some(sst_ref) = sst
                                    && let Ok(idx) = text_buf.parse::<usize>()
                                        && let Some(resolved) = sst_ref.strings.get(idx) {
                                            cur_cell_value = Some(resolved.clone());
                                            text_buf.clear();
                                            state = ParseState::InCell;
                                            continue;
                                        }
                            cur_cell_value = Some(std::mem::take(&mut text_buf));
                            state = ParseState::InCell;
                        }
                        (ParseState::InCellFormula, b"f") => {
                            cur_cell_formula = Some(std::mem::take(&mut text_buf));
                            state = ParseState::InCell;
                        }
                        (ParseState::InInlineStrT, b"t") => {
                            let text = std::mem::take(&mut text_buf);
                            cur_cell_value = Some(text.clone());
                            cur_cell_inline_string = Some(text);
                            state = ParseState::InInlineStr;
                        }
                        (ParseState::InInlineStr, b"is") => {
                            state = ParseState::InCell;
                        }
                        (ParseState::InCell, b"c") => {
                            cur_row_cells.push(Cell {
                                reference: std::mem::take(&mut cur_cell_ref),
                                cell_type: cur_cell_type,
                                style_index: cur_cell_style.take(),
                                value: cur_cell_value.take(),
                                formula: cur_cell_formula.take(),
                                formula_type: cur_cell_formula_type.take(),
                                formula_ref: cur_cell_formula_ref.take(),
                                shared_index: cur_cell_shared_index.take(),
                                inline_string: cur_cell_inline_string.take(),
                                dynamic_array: cur_cell_dynamic_array.take(),
                            });
                            state = ParseState::InRow;
                        }
                        (ParseState::InRow, b"row") => {
                            rows.push(Row {
                                index: cur_row_index,
                                cells: std::mem::take(&mut cur_row_cells),
                                height: cur_row_height.take(),
                                hidden: cur_row_hidden,
                                outline_level: cur_row_outline_level.take(),
                                collapsed: cur_row_collapsed,
                            });
                            state = ParseState::SheetData;
                        }
                        (ParseState::SheetData, b"sheetData") => {
                            state = ParseState::Root;
                        }
                        (ParseState::SheetPr, b"sheetPr") => {
                            state = ParseState::Root;
                        }
                        (ParseState::SheetView, b"sheetView") => {
                            state = ParseState::SheetViews;
                        }
                        (ParseState::SheetViews, b"sheetViews") => {
                            state = ParseState::Root;
                        }
                        (ParseState::Cols, b"cols") => {
                            state = ParseState::Root;
                        }
                        (ParseState::MergeCells, b"mergeCells") => {
                            state = ParseState::Root;
                        }
                        (ParseState::InDVFormula1, b"formula1") => {
                            if let Some(ref mut dv) = cur_dv {
                                dv.formula1 = Some(std::mem::take(&mut text_buf));
                            }
                            state = ParseState::InDataValidation;
                        }
                        (ParseState::InDVFormula2, b"formula2") => {
                            if let Some(ref mut dv) = cur_dv {
                                dv.formula2 = Some(std::mem::take(&mut text_buf));
                            }
                            state = ParseState::InDataValidation;
                        }
                        (ParseState::InDataValidation, b"dataValidation") => {
                            if let Some(dv) = cur_dv.take() {
                                data_validations.push(dv);
                            }
                            state = ParseState::InDataValidations;
                        }
                        (ParseState::InDataValidations, b"dataValidations") => {
                            state = ParseState::Root;
                        }
                        (ParseState::InCfRuleFormula, b"formula") => {
                            if let Some(ref mut rule) = cur_cf_rule {
                                rule.formula = Some(std::mem::take(&mut text_buf));
                            }
                            state = ParseState::InCfRule;
                        }
                        (ParseState::InCfRule, b"cfRule") => {
                            if let Some(rule) = cur_cf_rule.take() {
                                cur_cf_rules.push(rule);
                            }
                            state = ParseState::InConditionalFormatting;
                        }
                        (ParseState::InConditionalFormatting, b"conditionalFormatting") => {
                            conditional_formatting.push(ConditionalFormatting {
                                sqref: std::mem::take(&mut cur_cf_sqref),
                                rules: std::mem::take(&mut cur_cf_rules),
                            });
                            state = ParseState::Root;
                        }
                        (ParseState::InHyperlinks, b"hyperlinks") => {
                            state = ParseState::Root;
                        }
                        (ParseState::InAutoFilter, b"autoFilter") => {
                            auto_filter = Some(AutoFilter {
                                range: std::mem::take(&mut cur_af_range),
                                filter_columns: std::mem::take(&mut cur_af_columns),
                            });
                            state = ParseState::Root;
                        }
                        (ParseState::InFilters, b"filters") => {
                            state = ParseState::InFilterColumn;
                        }
                        (ParseState::InFilterColumn, b"filterColumn") => {
                            cur_af_columns.push(FilterColumn {
                                col_id: cur_filter_col_id,
                                filters: std::mem::take(&mut cur_filter_vals),
                            });
                            state = ParseState::InAutoFilter;
                        }
                        (ParseState::InColorScale, b"colorScale") => {
                            if let Some(ref mut rule) = cur_cf_rule {
                                rule.color_scale = Some(ColorScale {
                                    cfvos: std::mem::take(&mut cur_cfvos),
                                    colors: std::mem::take(&mut cur_cf_colors),
                                });
                            }
                            state = ParseState::InCfRule;
                        }
                        (ParseState::InDataBar, b"dataBar") => {
                            if let Some(ref mut rule) = cur_cf_rule {
                                rule.data_bar = Some(DataBar {
                                    cfvos: std::mem::take(&mut cur_cfvos),
                                    color: std::mem::take(&mut cur_cf_bar_color),
                                });
                            }
                            state = ParseState::InCfRule;
                        }
                        (ParseState::InIconSet, b"iconSet") => {
                            if let Some(ref mut rule) = cur_cf_rule {
                                rule.icon_set = Some(IconSet {
                                    icon_set_type: cur_icon_set_type.take(),
                                    cfvos: std::mem::take(&mut cur_cfvos),
                                });
                            }
                            state = ParseState::InCfRule;
                        }
                        (ParseState::HeaderFooterChild(idx), _) => {
                            reader.config_mut().trim_text(true);
                            if let Some(ref mut hf) = header_footer {
                                let text = if hf_text_buf.is_empty() { None } else { Some(std::mem::take(&mut hf_text_buf)) };
                                match idx {
                                    0 => hf.odd_header = text,
                                    1 => hf.odd_footer = text,
                                    2 => hf.even_header = text,
                                    3 => hf.even_footer = text,
                                    4 => hf.first_header = text,
                                    5 => hf.first_footer = text,
                                    _ => {}
                                }
                            }
                            state = ParseState::HeaderFooter;
                        }
                        (ParseState::HeaderFooter, b"headerFooter") => {
                            state = ParseState::Root;
                        }
                        _ => {}
                    }
                }
                Ok(Event::Eof) => break,
                Err(err) => {
                    return Err(ModernXlsxError::XmlParse(format!(
                        "error parsing worksheet XML: {err}"
                    )));
                }
                _ => {}
            }
            buf.clear();
        }

        Ok(WorksheetXml {
            dimension,
            rows,
            merge_cells,
            auto_filter,
            frozen_pane,
            split_pane,
            pane_selections,
            sheet_view: None,
            columns,
            data_validations,
            conditional_formatting,
            hyperlinks,
            page_setup,
            sheet_protection,
            comments: Vec::new(),
            tab_color,
            tables: Vec::new(),
            header_footer,
            outline_properties,
        })
    }

    /// Parse worksheet XML and write the result directly as a JSON object to `out`.
    ///
    /// This is a WASM-optimized alternative to `parse_with_sst()` + `serde_json::to_string()`.
    /// Instead of creating millions of intermediate `Cell`/`Row`/`String` objects,
    /// row and cell data is written directly to the JSON output buffer during parsing.
    /// Metadata (merge cells, columns, etc.) is still collected in small structs and
    /// serialized at the end via `serde_json` since they are negligible in size.
    ///
    /// For a 100K-row workbook, this eliminates ~1M `Cell` struct allocations,
    /// ~100K `Vec<Cell>` allocations, and ~2M `String` allocations, which is the
    /// primary cause of poor WASM performance due to frequent `memory.grow` calls.
    pub fn parse_to_json(
        data: &[u8],
        sst: Option<&super::shared_strings::SharedStringTable>,
        comments: &[super::comments::Comment],
        tables: &[super::tables::TableDefinition],
        out: &mut String,
    ) -> Result<()> {
        let mut reader = Reader::from_reader(data);
        reader.config_mut().trim_text(true);

        let mut buf = Vec::with_capacity(512);
        let mut text_buf = String::with_capacity(256);
        let mut itoa_buf = itoa::Buffer::new();

        // Metadata (small, collected as structs, serialized at the end).
        let mut dimension: Option<String> = None;
        let mut merge_cells: Vec<String> = Vec::new();
        let mut auto_filter: Option<AutoFilter> = None;
        let mut frozen_pane: Option<FrozenPane> = None;
        let mut split_pane = None::<SplitPane>;
        let mut pane_selections = Vec::<PaneSelection>::new();
        let mut columns: Vec<ColumnInfo> = Vec::new();
        let mut data_validations: Vec<DataValidation> = Vec::new();
        let mut conditional_formatting: Vec<ConditionalFormatting> = Vec::new();
        let mut hyperlinks: Vec<Hyperlink> = Vec::new();
        let mut page_setup: Option<PageSetup> = None;
        let mut sheet_protection: Option<SheetProtection> = None;
        let mut tab_color: Option<String> = None;
        let mut header_footer: Option<HeaderFooter> = None;
        let mut hf_text_buf = String::new();
        let mut outline_properties: Option<OutlineProperties> = None;

        // Row/cell streaming state (reused each row/cell, no accumulation).
        let mut state = ParseState::Root;
        let mut first_row = true;
        let mut first_cell = true;
        let mut cur_row_height: Option<f64> = None;
        let mut cur_row_hidden = false;
        let mut cur_row_outline_level: Option<u8> = None;
        let mut cur_row_collapsed = false;

        // Cell attribute storage (reused each cell via clear, keeps allocation).
        let mut cur_cell_ref = String::with_capacity(10);
        let mut cur_cell_type = CellType::Number;
        #[allow(unused_assignments)]
        let mut cur_cell_style: Option<u32> = None;
        let mut cur_cell_formula_type: Option<String> = None;
        let mut cur_cell_formula_ref: Option<String> = None;
        let mut cur_cell_shared_index: Option<u32> = None;
        let mut cur_cell_dynamic_array: Option<bool> = None;

        // Metadata builder state (same as parse_with_sst).
        let mut cur_dv: Option<DataValidation> = None;
        let mut cur_cf_sqref = String::new();
        let mut cur_cf_rules: Vec<ConditionalFormattingRule> = Vec::new();
        let mut cur_cf_rule: Option<ConditionalFormattingRule> = None;
        let mut cur_af_range = String::new();
        let mut cur_af_columns: Vec<FilterColumn> = Vec::new();
        let mut cur_filter_col_id: u32 = 0;
        let mut cur_filter_vals: Vec<String> = Vec::new();
        let mut cur_cfvos: Vec<Cfvo> = Vec::new();
        let mut cur_cf_colors: Vec<String> = Vec::new();
        let mut cur_cf_bar_color = String::new();
        let mut cur_icon_set_type: Option<String> = None;

        // Start the worksheet JSON object with rows array first.
        out.push_str("{\"rows\":[");

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) => {
                    let local = e.local_name();
                    match (state, local.as_ref()) {
                        // ---- Sheet properties (metadata) ----
                        (ParseState::Root, b"sheetPr") => state = ParseState::SheetPr,

                        // ---- Sheet views / pane (metadata) ----
                        (ParseState::Root, b"sheetViews") => state = ParseState::SheetViews,
                        (ParseState::SheetViews, b"sheetView") => state = ParseState::SheetView,
                        (ParseState::SheetView, b"selection") => {
                            // Handle <selection> as Start element (non-self-closing).
                            let mut sel_pane = None::<String>;
                            let mut sel_active = None::<String>;
                            let mut sel_sqref = None::<String>;
                            for attr in e.attributes().flatten() {
                                let ln = attr.key.local_name();
                                match ln.as_ref() {
                                    b"pane" => {
                                        sel_pane = Some(std::str::from_utf8(&attr.value).unwrap_or_default().to_owned());
                                    }
                                    b"activeCell" => {
                                        sel_active = Some(std::str::from_utf8(&attr.value).unwrap_or_default().to_owned());
                                    }
                                    b"sqref" => {
                                        sel_sqref = Some(std::str::from_utf8(&attr.value).unwrap_or_default().to_owned());
                                    }
                                    _ => {}
                                }
                            }
                            pane_selections.push(PaneSelection {
                                pane: sel_pane,
                                active_cell: sel_active,
                                sqref: sel_sqref,
                            });
                        }
                        (ParseState::Root, b"cols") => state = ParseState::Cols,

                        // ---- Sheet data (rows/cells → streamed to JSON) ----
                        (ParseState::Root, b"sheetData") => state = ParseState::SheetData,

                        (ParseState::SheetData, b"row") => {
                            state = ParseState::InRow;
                            cur_row_height = None;
                            cur_row_hidden = false;
                            cur_row_outline_level = None;
                            cur_row_collapsed = false;

                            let mut row_index: u32 = 0;
                            for attr in e.attributes().flatten() {
                                let ln = attr.key.local_name();
                                match ln.as_ref() {
                                    b"r" => {
                                        let val = std::str::from_utf8(&attr.value).unwrap_or_default();
                                        row_index = val.parse::<u32>().unwrap_or(0);
                                    }
                                    b"ht" => {
                                        let val = std::str::from_utf8(&attr.value).unwrap_or_default();
                                        cur_row_height = val.parse::<f64>().ok();
                                    }
                                    b"hidden" => {
                                        let val = std::str::from_utf8(&attr.value).unwrap_or_default();
                                        cur_row_hidden = val == "1" || val.eq_ignore_ascii_case("true");
                                    }
                                    b"outlineLevel" => {
                                        cur_row_outline_level = std::str::from_utf8(&attr.value)
                                            .ok()
                                            .and_then(|v| v.parse::<u8>().ok())
                                            .filter(|&v| v > 0);
                                    }
                                    b"collapsed" => {
                                        cur_row_collapsed = std::str::from_utf8(&attr.value).unwrap_or("0") == "1";
                                    }
                                    _ => {}
                                }
                            }

                            // Write row JSON start.
                            if !first_row { out.push(','); }
                            first_row = false;
                            out.push_str("{\"index\":");
                            out.push_str(itoa_buf.format(row_index));
                            out.push_str(",\"cells\":[");
                            first_cell = true;
                        }

                        (ParseState::InRow, b"c") => {
                            state = ParseState::InCell;
                            cur_cell_ref.clear();
                            cur_cell_type = CellType::Number;
                            cur_cell_style = None;
                            cur_cell_formula_type = None;
                            cur_cell_formula_ref = None;
                            cur_cell_shared_index = None;
                            cur_cell_dynamic_array = None;

                            for attr in e.attributes().flatten() {
                                let ln = attr.key.local_name();
                                match ln.as_ref() {
                                    b"r" => {
                                        cur_cell_ref.push_str(
                                            std::str::from_utf8(&attr.value).unwrap_or_default(),
                                        );
                                    }
                                    b"t" => {
                                        let val = std::str::from_utf8(&attr.value).unwrap_or_default();
                                        cur_cell_type = match val {
                                            "s" => CellType::SharedString,
                                            "b" => CellType::Boolean,
                                            "e" => CellType::Error,
                                            "str" => CellType::FormulaStr,
                                            "inlineStr" => CellType::InlineStr,
                                            _ => CellType::Number,
                                        };
                                    }
                                    b"s" => {
                                        let val = std::str::from_utf8(&attr.value).unwrap_or_default();
                                        cur_cell_style = val.parse::<u32>().ok();
                                    }
                                    _ => {}
                                }
                            }

                            // Write cell JSON header (reference + cellType + optional styleIndex).
                            if !first_cell { out.push(','); }
                            first_cell = false;
                            out.push_str("{\"reference\":\"");
                            out.push_str(&cur_cell_ref);
                            out.push_str("\",\"cellType\":\"");
                            out.push_str(cell_type_json_str(cur_cell_type));
                            out.push('"');
                            if let Some(si) = cur_cell_style {
                                out.push_str(",\"styleIndex\":");
                                out.push_str(itoa_buf.format(si));
                            }
                        }

                        (ParseState::InCell, b"v") => {
                            state = ParseState::InCellValue;
                            text_buf.clear();
                        }
                        (ParseState::InCell, b"f") => {
                            state = ParseState::InCellFormula;
                            text_buf.clear();
                            for attr in e.attributes().flatten() {
                                let ln = attr.key.local_name();
                                let val = std::str::from_utf8(&attr.value).unwrap_or_default();
                                match ln.as_ref() {
                                    b"t" => cur_cell_formula_type = Some(val.to_owned()),
                                    b"ref" => cur_cell_formula_ref = Some(val.to_owned()),
                                    b"si" => cur_cell_shared_index = val.parse::<u32>().ok(),
                                    b"cm" => {
                                        if val == "1" {
                                            cur_cell_dynamic_array = Some(true);
                                        }
                                    }
                                    _ => {}
                                }
                            }
                        }
                        (ParseState::InCell, b"is") => state = ParseState::InInlineStr,
                        (ParseState::InInlineStr, b"t") => {
                            state = ParseState::InInlineStrT;
                            text_buf.clear();
                        }

                        // ---- Merge cells (metadata) ----
                        (ParseState::Root, b"mergeCells") => state = ParseState::MergeCells,

                        // ---- Data validations (metadata) ----
                        (ParseState::Root, b"dataValidations") => state = ParseState::InDataValidations,
                        (ParseState::InDataValidations, b"dataValidation") => {
                            state = ParseState::InDataValidation;
                            let mut dv = DataValidation {
                                sqref: String::new(),
                                validation_type: None,
                                operator: None,
                                formula1: None,
                                formula2: None,
                                allow_blank: None,
                                show_error_message: None,
                                error_title: None,
                                error_message: None,
                                show_input_message: None,
                                prompt_title: None,
                                prompt: None,
                            };
                            for attr in e.attributes().flatten() {
                                let ln = attr.key.local_name();
                                let val = std::str::from_utf8(&attr.value).unwrap_or_default();
                                match ln.as_ref() {
                                    b"sqref" => dv.sqref = val.to_owned(),
                                    b"type" => dv.validation_type = Some(val.to_owned()),
                                    b"operator" => dv.operator = Some(val.to_owned()),
                                    b"allowBlank" => {
                                        dv.allow_blank = Some(val == "1" || val.eq_ignore_ascii_case("true"));
                                    }
                                    b"showErrorMessage" => {
                                        dv.show_error_message = Some(val == "1" || val.eq_ignore_ascii_case("true"));
                                    }
                                    b"errorTitle" => dv.error_title = Some(val.to_owned()),
                                    b"error" => dv.error_message = Some(val.to_owned()),
                                    b"showInputMessage" => {
                                        dv.show_input_message = Some(val == "1" || val.eq_ignore_ascii_case("true"));
                                    }
                                    b"promptTitle" => dv.prompt_title = Some(val.to_owned()),
                                    b"prompt" => dv.prompt = Some(val.to_owned()),
                                    _ => {}
                                }
                            }
                            cur_dv = Some(dv);
                        }
                        (ParseState::InDataValidation, b"formula1") => {
                            state = ParseState::InDVFormula1;
                            text_buf.clear();
                        }
                        (ParseState::InDataValidation, b"formula2") => {
                            state = ParseState::InDVFormula2;
                            text_buf.clear();
                        }

                        // ---- Conditional formatting (metadata) ----
                        (ParseState::Root, b"conditionalFormatting") => {
                            state = ParseState::InConditionalFormatting;
                            cur_cf_sqref.clear();
                            cur_cf_rules.clear();
                            for attr in e.attributes().flatten() {
                                if attr.key.local_name().as_ref() == b"sqref" {
                                    cur_cf_sqref = std::str::from_utf8(&attr.value)
                                        .unwrap_or_default()
                                        .to_owned();
                                }
                            }
                        }
                        (ParseState::InConditionalFormatting, b"cfRule") => {
                            state = ParseState::InCfRule;
                            let mut rule = ConditionalFormattingRule {
                                rule_type: String::new(),
                                priority: 0,
                                operator: None,
                                formula: None,
                                dxf_id: None,
                                color_scale: None,
                                data_bar: None,
                                icon_set: None,
                            };
                            for attr in e.attributes().flatten() {
                                let ln = attr.key.local_name();
                                let val = std::str::from_utf8(&attr.value).unwrap_or_default();
                                match ln.as_ref() {
                                    b"type" => rule.rule_type = val.to_owned(),
                                    b"priority" => rule.priority = val.parse::<u32>().unwrap_or(0),
                                    b"operator" => rule.operator = Some(val.to_owned()),
                                    b"dxfId" => rule.dxf_id = val.parse::<u32>().ok(),
                                    _ => {}
                                }
                            }
                            cur_cf_rule = Some(rule);
                        }
                        (ParseState::InCfRule, b"formula") => {
                            state = ParseState::InCfRuleFormula;
                            text_buf.clear();
                        }
                        (ParseState::InCfRule, b"colorScale") => {
                            state = ParseState::InColorScale;
                            cur_cfvos.clear();
                            cur_cf_colors.clear();
                        }
                        (ParseState::InCfRule, b"dataBar") => {
                            state = ParseState::InDataBar;
                            cur_cfvos.clear();
                            cur_cf_bar_color.clear();
                        }
                        (ParseState::InCfRule, b"iconSet") => {
                            state = ParseState::InIconSet;
                            cur_cfvos.clear();
                            cur_icon_set_type = None;
                            for attr in e.attributes().flatten() {
                                if attr.key.local_name().as_ref() == b"iconSet" {
                                    cur_icon_set_type = Some(
                                        std::str::from_utf8(&attr.value)
                                            .unwrap_or_default()
                                            .to_owned(),
                                    );
                                }
                            }
                        }

                        // ---- Hyperlinks (metadata) ----
                        (ParseState::Root, b"hyperlinks") => state = ParseState::InHyperlinks,

                        // ---- Auto filter (metadata) ----
                        (ParseState::Root, b"autoFilter") => {
                            state = ParseState::InAutoFilter;
                            cur_af_range.clear();
                            cur_af_columns.clear();
                            for attr in e.attributes().flatten() {
                                if attr.key.local_name().as_ref() == b"ref" {
                                    cur_af_range = std::str::from_utf8(&attr.value)
                                        .unwrap_or_default()
                                        .to_owned();
                                }
                            }
                        }
                        (ParseState::InAutoFilter, b"filterColumn") => {
                            state = ParseState::InFilterColumn;
                            cur_filter_col_id = 0;
                            cur_filter_vals.clear();
                            for attr in e.attributes().flatten() {
                                if attr.key.local_name().as_ref() == b"colId" {
                                    cur_filter_col_id = std::str::from_utf8(&attr.value)
                                        .unwrap_or_default()
                                        .parse::<u32>()
                                        .unwrap_or(0);
                                }
                            }
                        }
                        (ParseState::InFilterColumn, b"filters") => state = ParseState::InFilters,

                        (ParseState::Root, b"headerFooter") => {
                            let mut hf = HeaderFooter::default();
                            for attr in e.attributes().flatten() {
                                match attr.key.as_ref() {
                                    b"differentOddEven" => hf.different_odd_even = std::str::from_utf8(&attr.value).unwrap_or("0") == "1",
                                    b"differentFirst" => hf.different_first = std::str::from_utf8(&attr.value).unwrap_or("0") == "1",
                                    b"scaleWithDoc" => hf.scale_with_doc = std::str::from_utf8(&attr.value).unwrap_or("1") != "0",
                                    b"alignWithMargins" => hf.align_with_margins = std::str::from_utf8(&attr.value).unwrap_or("1") != "0",
                                    _ => {}
                                }
                            }
                            header_footer = Some(hf);
                            state = ParseState::HeaderFooter;
                        }
                        (ParseState::HeaderFooter, b"oddHeader") => { hf_text_buf.clear(); reader.config_mut().trim_text(false); state = ParseState::HeaderFooterChild(0); }
                        (ParseState::HeaderFooter, b"oddFooter") => { hf_text_buf.clear(); reader.config_mut().trim_text(false); state = ParseState::HeaderFooterChild(1); }
                        (ParseState::HeaderFooter, b"evenHeader") => { hf_text_buf.clear(); reader.config_mut().trim_text(false); state = ParseState::HeaderFooterChild(2); }
                        (ParseState::HeaderFooter, b"evenFooter") => { hf_text_buf.clear(); reader.config_mut().trim_text(false); state = ParseState::HeaderFooterChild(3); }
                        (ParseState::HeaderFooter, b"firstHeader") => { hf_text_buf.clear(); reader.config_mut().trim_text(false); state = ParseState::HeaderFooterChild(4); }
                        (ParseState::HeaderFooter, b"firstFooter") => { hf_text_buf.clear(); reader.config_mut().trim_text(false); state = ParseState::HeaderFooterChild(5); }

                        _ => {}
                    }
                }

                Ok(Event::Empty(ref e)) => {
                    let local = e.local_name();
                    match (state, local.as_ref()) {
                        (ParseState::Root, b"dimension") => {
                            for attr in e.attributes().flatten() {
                                if attr.key.local_name().as_ref() == b"ref" {
                                    dimension = Some(
                                        std::str::from_utf8(&attr.value).unwrap_or_default().to_owned(),
                                    );
                                }
                            }
                        }
                        (ParseState::SheetPr, b"tabColor") => {
                            for attr in e.attributes().flatten() {
                                if attr.key.local_name().as_ref() == b"rgb" {
                                    tab_color = Some(
                                        std::str::from_utf8(&attr.value).unwrap_or_default().to_owned(),
                                    );
                                }
                            }
                        }
                        (ParseState::SheetPr, b"outlinePr") => {
                            let mut op = OutlineProperties { summary_below: true, summary_right: true };
                            for attr in e.attributes().flatten() {
                                match attr.key.as_ref() {
                                    b"summaryBelow" => op.summary_below = std::str::from_utf8(&attr.value).unwrap_or("1") != "0",
                                    b"summaryRight" => op.summary_right = std::str::from_utf8(&attr.value).unwrap_or("1") != "0",
                                    _ => {}
                                }
                            }
                            outline_properties = Some(op);
                        }
                        (ParseState::Root, b"autoFilter") => {
                            let mut af_range = String::new();
                            for attr in e.attributes().flatten() {
                                if attr.key.local_name().as_ref() == b"ref" {
                                    af_range = std::str::from_utf8(&attr.value)
                                        .unwrap_or_default()
                                        .to_owned();
                                }
                            }
                            auto_filter = Some(AutoFilter {
                                range: af_range,
                                filter_columns: Vec::new(),
                            });
                        }
                        (ParseState::SheetView, b"pane") => {
                            let mut y_split_raw = String::new();
                            let mut x_split_raw = String::new();
                            let mut state_val = String::new();
                            let mut top_left_cell_val = None::<String>;
                            let mut active_pane_val = None::<String>;

                            for attr in e.attributes().flatten() {
                                let ln = attr.key.local_name();
                                match ln.as_ref() {
                                    b"ySplit" => {
                                        y_split_raw = std::str::from_utf8(&attr.value).unwrap_or_default().to_owned();
                                    }
                                    b"xSplit" => {
                                        x_split_raw = std::str::from_utf8(&attr.value).unwrap_or_default().to_owned();
                                    }
                                    b"state" => {
                                        state_val = std::str::from_utf8(&attr.value).unwrap_or_default().to_owned();
                                    }
                                    b"topLeftCell" => {
                                        top_left_cell_val = Some(std::str::from_utf8(&attr.value).unwrap_or_default().to_owned());
                                    }
                                    b"activePane" => {
                                        active_pane_val = Some(std::str::from_utf8(&attr.value).unwrap_or_default().to_owned());
                                    }
                                    _ => {}
                                }
                            }

                            if state_val == "frozen" {
                                let y: u32 = y_split_raw.parse().unwrap_or(0);
                                let x: u32 = x_split_raw.parse().unwrap_or(0);
                                if y > 0 || x > 0 {
                                    frozen_pane = Some(FrozenPane { rows: y, cols: x });
                                }
                            } else {
                                let y: f64 = y_split_raw.parse().unwrap_or(0.0);
                                let x: f64 = x_split_raw.parse().unwrap_or(0.0);
                                if y > 0.0 || x > 0.0 {
                                    split_pane = Some(SplitPane {
                                        horizontal: if y > 0.0 { Some(y) } else { None },
                                        vertical: if x > 0.0 { Some(x) } else { None },
                                        top_left_cell: top_left_cell_val,
                                        active_pane: active_pane_val,
                                    });
                                }
                            }
                        }
                        (ParseState::SheetView, b"selection") => {
                            let mut sel_pane = None::<String>;
                            let mut sel_active = None::<String>;
                            let mut sel_sqref = None::<String>;
                            for attr in e.attributes().flatten() {
                                let ln = attr.key.local_name();
                                match ln.as_ref() {
                                    b"pane" => {
                                        sel_pane = Some(std::str::from_utf8(&attr.value).unwrap_or_default().to_owned());
                                    }
                                    b"activeCell" => {
                                        sel_active = Some(std::str::from_utf8(&attr.value).unwrap_or_default().to_owned());
                                    }
                                    b"sqref" => {
                                        sel_sqref = Some(std::str::from_utf8(&attr.value).unwrap_or_default().to_owned());
                                    }
                                    _ => {}
                                }
                            }
                            pane_selections.push(PaneSelection {
                                pane: sel_pane,
                                active_cell: sel_active,
                                sqref: sel_sqref,
                            });
                        }
                        (ParseState::Cols, b"col") => columns.push(parse_col_element(e)),
                        (ParseState::MergeCells, b"mergeCell") => {
                            for attr in e.attributes().flatten() {
                                if attr.key.local_name().as_ref() == b"ref" {
                                    merge_cells.push(
                                        std::str::from_utf8(&attr.value).unwrap_or_default().to_owned(),
                                    );
                                }
                            }
                        }
                        // Self-closing <c ... /> — cell with no children.
                        (ParseState::InRow, b"c") => {
                            let mut cell_ref_buf = String::new();
                            let mut cell_type = CellType::Number;
                            let mut cell_style: Option<u32> = None;
                            for attr in e.attributes().flatten() {
                                let ln = attr.key.local_name();
                                match ln.as_ref() {
                                    b"r" => {
                                        cell_ref_buf.push_str(
                                            std::str::from_utf8(&attr.value).unwrap_or_default(),
                                        );
                                    }
                                    b"t" => {
                                        let val = std::str::from_utf8(&attr.value).unwrap_or_default();
                                        cell_type = match val {
                                            "s" => CellType::SharedString,
                                            "b" => CellType::Boolean,
                                            "e" => CellType::Error,
                                            "str" => CellType::FormulaStr,
                                            "inlineStr" => CellType::InlineStr,
                                            _ => CellType::Number,
                                        };
                                    }
                                    b"s" => {
                                        let val = std::str::from_utf8(&attr.value).unwrap_or_default();
                                        cell_style = val.parse::<u32>().ok();
                                    }
                                    _ => {}
                                }
                            }
                            if !first_cell { out.push(','); }
                            first_cell = false;
                            out.push_str("{\"reference\":\"");
                            out.push_str(&cell_ref_buf);
                            out.push_str("\",\"cellType\":\"");
                            out.push_str(cell_type_json_str(cell_type));
                            out.push('"');
                            if let Some(si) = cell_style {
                                out.push_str(",\"styleIndex\":");
                                out.push_str(itoa_buf.format(si));
                            }
                            out.push('}');
                        }
                        // Self-closing <row ... /> — empty row.
                        (ParseState::SheetData, b"row") => {
                            let mut row_index: u32 = 0;
                            let mut row_height: Option<f64> = None;
                            let mut row_hidden = false;
                            let mut row_ol: Option<u8> = None;
                            let mut row_coll = false;
                            for attr in e.attributes().flatten() {
                                let ln = attr.key.local_name();
                                match ln.as_ref() {
                                    b"r" => {
                                        let val = std::str::from_utf8(&attr.value).unwrap_or_default();
                                        row_index = val.parse::<u32>().unwrap_or(0);
                                    }
                                    b"ht" => {
                                        let val = std::str::from_utf8(&attr.value).unwrap_or_default();
                                        row_height = val.parse::<f64>().ok();
                                    }
                                    b"hidden" => {
                                        let val = std::str::from_utf8(&attr.value).unwrap_or_default();
                                        row_hidden = val == "1" || val.eq_ignore_ascii_case("true");
                                    }
                                    b"outlineLevel" => {
                                        row_ol = std::str::from_utf8(&attr.value)
                                            .ok()
                                            .and_then(|v| v.parse::<u8>().ok())
                                            .filter(|&v| v > 0);
                                    }
                                    b"collapsed" => {
                                        row_coll = std::str::from_utf8(&attr.value).unwrap_or("0") == "1";
                                    }
                                    _ => {}
                                }
                            }
                            if !first_row { out.push(','); }
                            first_row = false;
                            out.push_str("{\"index\":");
                            out.push_str(itoa_buf.format(row_index));
                            out.push_str(",\"cells\":[]");
                            if let Some(h) = row_height {
                                out.push_str(",\"height\":");
                                write_f64_json(out, h);
                            }
                            if row_hidden {
                                out.push_str(",\"hidden\":true");
                            }
                            if let Some(level) = row_ol {
                                out.push_str(",\"outlineLevel\":");
                                out.push_str(itoa_buf.format(level));
                            }
                            if row_coll {
                                out.push_str(",\"collapsed\":true");
                            }
                            out.push('}');
                        }
                        (ParseState::InCell, b"v") => { /* Empty <v/> — no value */ }
                        (ParseState::InCell, b"f") => {
                            for attr in e.attributes().flatten() {
                                let ln = attr.key.local_name();
                                let val = std::str::from_utf8(&attr.value).unwrap_or_default();
                                match ln.as_ref() {
                                    b"t" => cur_cell_formula_type = Some(val.to_owned()),
                                    b"ref" => cur_cell_formula_ref = Some(val.to_owned()),
                                    b"si" => cur_cell_shared_index = val.parse::<u32>().ok(),
                                    b"cm" => {
                                        if val == "1" {
                                            cur_cell_dynamic_array = Some(true);
                                        }
                                    }
                                    _ => {}
                                }
                            }
                        }
                        (ParseState::InDataValidations, b"dataValidation") => {
                            let mut dv = DataValidation {
                                sqref: String::new(),
                                validation_type: None,
                                operator: None,
                                formula1: None,
                                formula2: None,
                                allow_blank: None,
                                show_error_message: None,
                                error_title: None,
                                error_message: None,
                                show_input_message: None,
                                prompt_title: None,
                                prompt: None,
                            };
                            for attr in e.attributes().flatten() {
                                let ln = attr.key.local_name();
                                let val = std::str::from_utf8(&attr.value).unwrap_or_default();
                                match ln.as_ref() {
                                    b"sqref" => dv.sqref = val.to_owned(),
                                    b"type" => dv.validation_type = Some(val.to_owned()),
                                    b"operator" => dv.operator = Some(val.to_owned()),
                                    b"allowBlank" => {
                                        dv.allow_blank = Some(val == "1" || val.eq_ignore_ascii_case("true"));
                                    }
                                    b"showErrorMessage" => {
                                        dv.show_error_message = Some(val == "1" || val.eq_ignore_ascii_case("true"));
                                    }
                                    b"errorTitle" => dv.error_title = Some(val.to_owned()),
                                    b"error" => dv.error_message = Some(val.to_owned()),
                                    b"showInputMessage" => {
                                        dv.show_input_message = Some(val == "1" || val.eq_ignore_ascii_case("true"));
                                    }
                                    b"promptTitle" => dv.prompt_title = Some(val.to_owned()),
                                    b"prompt" => dv.prompt = Some(val.to_owned()),
                                    _ => {}
                                }
                            }
                            data_validations.push(dv);
                        }
                        (ParseState::InConditionalFormatting, b"cfRule") => {
                            let mut rule = ConditionalFormattingRule {
                                rule_type: String::new(),
                                priority: 0,
                                operator: None,
                                formula: None,
                                dxf_id: None,
                                color_scale: None,
                                data_bar: None,
                                icon_set: None,
                            };
                            for attr in e.attributes().flatten() {
                                let ln = attr.key.local_name();
                                let val = std::str::from_utf8(&attr.value).unwrap_or_default();
                                match ln.as_ref() {
                                    b"type" => rule.rule_type = val.to_owned(),
                                    b"priority" => rule.priority = val.parse::<u32>().unwrap_or(0),
                                    b"operator" => rule.operator = Some(val.to_owned()),
                                    b"dxfId" => rule.dxf_id = val.parse::<u32>().ok(),
                                    _ => {}
                                }
                            }
                            cur_cf_rules.push(rule);
                        }
                        (ParseState::InHyperlinks, b"hyperlink") => {
                            let mut hl = Hyperlink {
                                cell_ref: String::new(),
                                location: None,
                                display: None,
                                tooltip: None,
                            };
                            for attr in e.attributes().flatten() {
                                let ln = attr.key.local_name();
                                let val = std::str::from_utf8(&attr.value).unwrap_or_default();
                                match ln.as_ref() {
                                    b"ref" => hl.cell_ref = val.to_owned(),
                                    b"location" => hl.location = Some(val.to_owned()),
                                    b"display" => hl.display = Some(val.to_owned()),
                                    b"tooltip" => hl.tooltip = Some(val.to_owned()),
                                    _ => {}
                                }
                            }
                            hyperlinks.push(hl);
                        }
                        (ParseState::Root, b"pageSetup") => {
                            let mut ps = PageSetup {
                                paper_size: None,
                                orientation: None,
                                fit_to_width: None,
                                fit_to_height: None,
                                scale: None,
                                first_page_number: None,
                                horizontal_dpi: None,
                                vertical_dpi: None,
                            };
                            for attr in e.attributes().flatten() {
                                let ln = attr.key.local_name();
                                let val = std::str::from_utf8(&attr.value).unwrap_or_default();
                                match ln.as_ref() {
                                    b"paperSize" => ps.paper_size = val.parse::<u32>().ok(),
                                    b"orientation" => ps.orientation = Some(val.to_owned()),
                                    b"fitToWidth" => ps.fit_to_width = val.parse::<u32>().ok(),
                                    b"fitToHeight" => ps.fit_to_height = val.parse::<u32>().ok(),
                                    b"scale" => ps.scale = val.parse::<u32>().ok(),
                                    b"firstPageNumber" => ps.first_page_number = val.parse::<u32>().ok(),
                                    b"horizontalDpi" => ps.horizontal_dpi = val.parse::<u32>().ok(),
                                    b"verticalDpi" => ps.vertical_dpi = val.parse::<u32>().ok(),
                                    _ => {}
                                }
                            }
                            page_setup = Some(ps);
                        }
                        (ParseState::Root, b"sheetProtection") => {
                            let mut sp = SheetProtection {
                                sheet: false,
                                objects: false,
                                scenarios: false,
                                password: None,
                                format_cells: false,
                                format_columns: false,
                                format_rows: false,
                                insert_columns: false,
                                insert_rows: false,
                                delete_columns: false,
                                delete_rows: false,
                                sort: false,
                                auto_filter: false,
                            };
                            for attr in e.attributes().flatten() {
                                let ln = attr.key.local_name();
                                let val = std::str::from_utf8(&attr.value).unwrap_or_default();
                                let as_bool = val == "1" || val.eq_ignore_ascii_case("true");
                                match ln.as_ref() {
                                    b"sheet" => sp.sheet = as_bool,
                                    b"objects" => sp.objects = as_bool,
                                    b"scenarios" => sp.scenarios = as_bool,
                                    b"password" => sp.password = Some(val.to_owned()),
                                    b"formatCells" => sp.format_cells = as_bool,
                                    b"formatColumns" => sp.format_columns = as_bool,
                                    b"formatRows" => sp.format_rows = as_bool,
                                    b"insertColumns" => sp.insert_columns = as_bool,
                                    b"insertRows" => sp.insert_rows = as_bool,
                                    b"deleteColumns" => sp.delete_columns = as_bool,
                                    b"deleteRows" => sp.delete_rows = as_bool,
                                    b"sort" => sp.sort = as_bool,
                                    b"autoFilter" => sp.auto_filter = as_bool,
                                    _ => {}
                                }
                            }
                            sheet_protection = Some(sp);
                        }
                        (ParseState::InColorScale | ParseState::InDataBar | ParseState::InIconSet, b"cfvo") => {
                            let mut cfvo_type = String::new();
                            let mut cfvo_val: Option<String> = None;
                            for attr in e.attributes().flatten() {
                                let ln = attr.key.local_name();
                                let v = std::str::from_utf8(&attr.value).unwrap_or_default();
                                match ln.as_ref() {
                                    b"type" => cfvo_type = v.to_owned(),
                                    b"val" => cfvo_val = Some(v.to_owned()),
                                    _ => {}
                                }
                            }
                            cur_cfvos.push(Cfvo { cfvo_type, val: cfvo_val });
                        }
                        (ParseState::InColorScale, b"color") => {
                            for attr in e.attributes().flatten() {
                                if attr.key.local_name().as_ref() == b"rgb" {
                                    cur_cf_colors.push(
                                        std::str::from_utf8(&attr.value).unwrap_or_default().to_owned(),
                                    );
                                }
                            }
                        }
                        (ParseState::InDataBar, b"color") => {
                            for attr in e.attributes().flatten() {
                                if attr.key.local_name().as_ref() == b"rgb" {
                                    cur_cf_bar_color = std::str::from_utf8(&attr.value)
                                        .unwrap_or_default()
                                        .to_owned();
                                }
                            }
                        }
                        (ParseState::InFilters, b"filter") => {
                            for attr in e.attributes().flatten() {
                                if attr.key.local_name().as_ref() == b"val" {
                                    cur_filter_vals.push(
                                        std::str::from_utf8(&attr.value).unwrap_or_default().to_owned(),
                                    );
                                }
                            }
                        }
                        _ => {}
                    }
                }

                Ok(Event::Text(ref e)) => {
                    match state {
                        ParseState::InCellValue
                        | ParseState::InCellFormula
                        | ParseState::InInlineStrT
                        | ParseState::InDVFormula1
                        | ParseState::InDVFormula2
                        | ParseState::InCfRuleFormula => {
                            text_buf.push_str(std::str::from_utf8(e.as_ref()).unwrap_or_default());
                        }
                        ParseState::HeaderFooterChild(_) => {
                            hf_text_buf.push_str(std::str::from_utf8(e.as_ref()).unwrap_or_default());
                        }
                        _ => {}
                    }
                }
                Ok(Event::GeneralRef(ref e)) => {
                    match state {
                        ParseState::InCellValue
                        | ParseState::InCellFormula
                        | ParseState::InInlineStrT
                        | ParseState::InDVFormula1
                        | ParseState::InDVFormula2
                        | ParseState::InCfRuleFormula => {
                            push_entity(&mut text_buf, e.as_ref());
                        }
                        ParseState::HeaderFooterChild(_) => {
                            push_entity(&mut hf_text_buf, e.as_ref());
                        }
                        _ => {}
                    }
                }

                Ok(Event::End(ref e)) => {
                    let local = e.local_name();
                    match (state, local.as_ref()) {
                        // ---- Cell value end → write JSON directly ----
                        (ParseState::InCellValue, b"v") => {
                            if cur_cell_type == CellType::SharedString
                                && let Some(sst_ref) = sst
                                    && let Ok(idx) = text_buf.parse::<usize>()
                                        && let Some(resolved) = sst_ref.strings.get(idx) {
                                            out.push_str(",\"value\":\"");
                                            json_escape_to(out, resolved);
                                            out.push('"');
                                            text_buf.clear();
                                            state = ParseState::InCell;
                                            continue;
                                        }
                            out.push_str(",\"value\":\"");
                            json_escape_to(out, &text_buf);
                            out.push('"');
                            text_buf.clear();
                            state = ParseState::InCell;
                        }

                        // ---- Formula end → write JSON directly ----
                        (ParseState::InCellFormula, b"f") => {
                            out.push_str(",\"formula\":\"");
                            json_escape_to(out, &text_buf);
                            out.push('"');
                            text_buf.clear();
                            if let Some(ref ft) = cur_cell_formula_type {
                                out.push_str(",\"formulaType\":\"");
                                json_escape_to(out, ft);
                                out.push('"');
                            }
                            if let Some(ref fr) = cur_cell_formula_ref {
                                out.push_str(",\"formulaRef\":\"");
                                json_escape_to(out, fr);
                                out.push('"');
                            }
                            if let Some(si) = cur_cell_shared_index {
                                out.push_str(",\"sharedIndex\":");
                                out.push_str(itoa_buf.format(si));
                            }
                            if cur_cell_dynamic_array == Some(true) {
                                out.push_str(",\"dynamicArray\":true");
                            }
                            state = ParseState::InCell;
                        }

                        // ---- Inline string end → write both value and inlineString ----
                        (ParseState::InInlineStrT, b"t") => {
                            out.push_str(",\"value\":\"");
                            json_escape_to(out, &text_buf);
                            out.push_str("\",\"inlineString\":\"");
                            json_escape_to(out, &text_buf);
                            out.push('"');
                            text_buf.clear();
                            state = ParseState::InInlineStr;
                        }
                        (ParseState::InInlineStr, b"is") => state = ParseState::InCell,

                        // ---- Cell end → close cell JSON object ----
                        (ParseState::InCell, b"c") => {
                            // Write formula attributes even if no formula text was present
                            // (happens with self-closing <f/> parsed in Empty handler).
                            if cur_cell_formula_type.is_some()
                                || cur_cell_formula_ref.is_some()
                                || cur_cell_shared_index.is_some()
                            {
                                if let Some(ref ft) = cur_cell_formula_type.take() {
                                    out.push_str(",\"formulaType\":\"");
                                    json_escape_to(out, ft);
                                    out.push('"');
                                }
                                if let Some(ref fr) = cur_cell_formula_ref.take() {
                                    out.push_str(",\"formulaRef\":\"");
                                    json_escape_to(out, fr);
                                    out.push('"');
                                }
                                if let Some(si) = cur_cell_shared_index.take() {
                                    out.push_str(",\"sharedIndex\":");
                                    out.push_str(itoa_buf.format(si));
                                }
                            }
                            if cur_cell_dynamic_array.take() == Some(true) {
                                out.push_str(",\"dynamicArray\":true");
                            }
                            out.push('}');
                            state = ParseState::InRow;
                        }

                        // ---- Row end → close row JSON object ----
                        (ParseState::InRow, b"row") => {
                            out.push(']'); // close cells array
                            if let Some(h) = cur_row_height {
                                out.push_str(",\"height\":");
                                write_f64_json(out, h);
                            }
                            if cur_row_hidden {
                                out.push_str(",\"hidden\":true");
                            }
                            if let Some(level) = cur_row_outline_level {
                                out.push_str(",\"outlineLevel\":");
                                out.push_str(itoa_buf.format(level));
                            }
                            if cur_row_collapsed {
                                out.push_str(",\"collapsed\":true");
                            }
                            out.push('}');
                            state = ParseState::SheetData;
                        }

                        (ParseState::SheetData, b"sheetData") => state = ParseState::Root,
                        (ParseState::SheetPr, b"sheetPr") => state = ParseState::Root,
                        (ParseState::SheetView, b"sheetView") => state = ParseState::SheetViews,
                        (ParseState::SheetViews, b"sheetViews") => state = ParseState::Root,
                        (ParseState::Cols, b"cols") => state = ParseState::Root,
                        (ParseState::MergeCells, b"mergeCells") => state = ParseState::Root,

                        // ---- Data validation metadata end ----
                        (ParseState::InDVFormula1, b"formula1") => {
                            if let Some(ref mut dv) = cur_dv {
                                dv.formula1 = Some(std::mem::take(&mut text_buf));
                            }
                            state = ParseState::InDataValidation;
                        }
                        (ParseState::InDVFormula2, b"formula2") => {
                            if let Some(ref mut dv) = cur_dv {
                                dv.formula2 = Some(std::mem::take(&mut text_buf));
                            }
                            state = ParseState::InDataValidation;
                        }
                        (ParseState::InDataValidation, b"dataValidation") => {
                            if let Some(dv) = cur_dv.take() {
                                data_validations.push(dv);
                            }
                            state = ParseState::InDataValidations;
                        }
                        (ParseState::InDataValidations, b"dataValidations") => state = ParseState::Root,

                        // ---- Conditional formatting metadata end ----
                        (ParseState::InCfRuleFormula, b"formula") => {
                            if let Some(ref mut rule) = cur_cf_rule {
                                rule.formula = Some(std::mem::take(&mut text_buf));
                            }
                            state = ParseState::InCfRule;
                        }
                        (ParseState::InCfRule, b"cfRule") => {
                            if let Some(rule) = cur_cf_rule.take() {
                                cur_cf_rules.push(rule);
                            }
                            state = ParseState::InConditionalFormatting;
                        }
                        (ParseState::InConditionalFormatting, b"conditionalFormatting") => {
                            conditional_formatting.push(ConditionalFormatting {
                                sqref: std::mem::take(&mut cur_cf_sqref),
                                rules: std::mem::take(&mut cur_cf_rules),
                            });
                            state = ParseState::Root;
                        }
                        (ParseState::InHyperlinks, b"hyperlinks") => state = ParseState::Root,
                        (ParseState::InAutoFilter, b"autoFilter") => {
                            auto_filter = Some(AutoFilter {
                                range: std::mem::take(&mut cur_af_range),
                                filter_columns: std::mem::take(&mut cur_af_columns),
                            });
                            state = ParseState::Root;
                        }
                        (ParseState::InFilters, b"filters") => state = ParseState::InFilterColumn,
                        (ParseState::InFilterColumn, b"filterColumn") => {
                            cur_af_columns.push(FilterColumn {
                                col_id: cur_filter_col_id,
                                filters: std::mem::take(&mut cur_filter_vals),
                            });
                            state = ParseState::InAutoFilter;
                        }
                        (ParseState::InColorScale, b"colorScale") => {
                            if let Some(ref mut rule) = cur_cf_rule {
                                rule.color_scale = Some(ColorScale {
                                    cfvos: std::mem::take(&mut cur_cfvos),
                                    colors: std::mem::take(&mut cur_cf_colors),
                                });
                            }
                            state = ParseState::InCfRule;
                        }
                        (ParseState::InDataBar, b"dataBar") => {
                            if let Some(ref mut rule) = cur_cf_rule {
                                rule.data_bar = Some(DataBar {
                                    cfvos: std::mem::take(&mut cur_cfvos),
                                    color: std::mem::take(&mut cur_cf_bar_color),
                                });
                            }
                            state = ParseState::InCfRule;
                        }
                        (ParseState::InIconSet, b"iconSet") => {
                            if let Some(ref mut rule) = cur_cf_rule {
                                rule.icon_set = Some(IconSet {
                                    icon_set_type: cur_icon_set_type.take(),
                                    cfvos: std::mem::take(&mut cur_cfvos),
                                });
                            }
                            state = ParseState::InCfRule;
                        }
                        (ParseState::HeaderFooterChild(idx), _) => {
                            reader.config_mut().trim_text(true);
                            if let Some(ref mut hf) = header_footer {
                                let text = if hf_text_buf.is_empty() { None } else { Some(std::mem::take(&mut hf_text_buf)) };
                                match idx {
                                    0 => hf.odd_header = text,
                                    1 => hf.odd_footer = text,
                                    2 => hf.even_header = text,
                                    3 => hf.even_footer = text,
                                    4 => hf.first_header = text,
                                    5 => hf.first_footer = text,
                                    _ => {}
                                }
                            }
                            state = ParseState::HeaderFooter;
                        }
                        (ParseState::HeaderFooter, b"headerFooter") => {
                            state = ParseState::Root;
                        }

                        _ => {}
                    }
                }

                Ok(Event::Eof) => break,
                Err(err) => {
                    return Err(ModernXlsxError::XmlParse(format!(
                        "error parsing worksheet XML: {err}"
                    )));
                }
                _ => {}
            }
            buf.clear();
        }

        // Close the rows array.
        out.push(']');

        // Write metadata fields using serde_json (these are tiny).
        if let Some(ref d) = dimension {
            out.push_str(",\"dimension\":\"");
            json_escape_to(out, d);
            out.push('"');
        }
        if !merge_cells.is_empty() {
            out.push_str(",\"mergeCells\":");
            out.push_str(&serde_json::to_string(&merge_cells)
                .map_err(|e| ModernXlsxError::XmlParse(e.to_string()))?);
        }
        if let Some(ref af) = auto_filter {
            out.push_str(",\"autoFilter\":");
            out.push_str(&serde_json::to_string(af)
                .map_err(|e| ModernXlsxError::XmlParse(e.to_string()))?);
        }
        if let Some(ref fp) = frozen_pane {
            out.push_str(",\"frozenPane\":");
            out.push_str(&serde_json::to_string(fp)
                .map_err(|e| ModernXlsxError::XmlParse(e.to_string()))?);
        }
        if let Some(ref sp) = split_pane {
            out.push_str(",\"splitPane\":");
            out.push_str(&serde_json::to_string(sp)
                .map_err(|e| ModernXlsxError::XmlParse(e.to_string()))?);
        }
        if !pane_selections.is_empty() {
            out.push_str(",\"paneSelections\":");
            out.push_str(&serde_json::to_string(&pane_selections)
                .map_err(|e| ModernXlsxError::XmlParse(e.to_string()))?);
        }
        if !columns.is_empty() {
            out.push_str(",\"columns\":");
            out.push_str(&serde_json::to_string(&columns)
                .map_err(|e| ModernXlsxError::XmlParse(e.to_string()))?);
        }
        if !data_validations.is_empty() {
            out.push_str(",\"dataValidations\":");
            out.push_str(&serde_json::to_string(&data_validations)
                .map_err(|e| ModernXlsxError::XmlParse(e.to_string()))?);
        }
        if !conditional_formatting.is_empty() {
            out.push_str(",\"conditionalFormatting\":");
            out.push_str(&serde_json::to_string(&conditional_formatting)
                .map_err(|e| ModernXlsxError::XmlParse(e.to_string()))?);
        }
        if !hyperlinks.is_empty() {
            out.push_str(",\"hyperlinks\":");
            out.push_str(&serde_json::to_string(&hyperlinks)
                .map_err(|e| ModernXlsxError::XmlParse(e.to_string()))?);
        }
        if let Some(ref ps) = page_setup {
            out.push_str(",\"pageSetup\":");
            out.push_str(&serde_json::to_string(ps)
                .map_err(|e| ModernXlsxError::XmlParse(e.to_string()))?);
        }
        if let Some(ref sp) = sheet_protection {
            out.push_str(",\"sheetProtection\":");
            out.push_str(&serde_json::to_string(sp)
                .map_err(|e| ModernXlsxError::XmlParse(e.to_string()))?);
        }
        if !comments.is_empty() {
            out.push_str(",\"comments\":");
            out.push_str(&serde_json::to_string(comments)
                .map_err(|e| ModernXlsxError::XmlParse(e.to_string()))?);
        }
        if !tables.is_empty() {
            out.push_str(",\"tables\":");
            out.push_str(&serde_json::to_string(tables)
                .map_err(|e| ModernXlsxError::XmlParse(e.to_string()))?);
        }
        if let Some(ref tc) = tab_color {
            out.push_str(",\"tabColor\":\"");
            json_escape_to(out, tc);
            out.push('"');
        }
        if let Some(ref hf) = header_footer {
            out.push_str(",\"headerFooter\":");
            out.push_str(&serde_json::to_string(hf)
                .map_err(|e| ModernXlsxError::XmlParse(e.to_string()))?);
        }
        if let Some(ref op) = outline_properties {
            out.push_str(",\"outlineProperties\":");
            out.push_str(&serde_json::to_string(op)
                .map_err(|e| ModernXlsxError::XmlParse(e.to_string()))?);
        }

        // Close the worksheet JSON object.
        out.push('}');

        Ok(())
    }

    /// Serialize this worksheet to valid XML bytes.
    pub fn to_xml(&self) -> Result<Vec<u8>> {
        self.to_xml_with_sst(None, &[])
    }

    /// Serialize to worksheet XML bytes, optionally remapping SharedString
    /// cell values to SST indices on-the-fly (avoiding a full clone of the
    /// worksheet).
    ///
    /// `table_r_ids` are the relationship IDs for `<tableParts>` elements;
    /// pass an empty slice when no tables are attached to this sheet.
    pub fn to_xml_with_sst(
        &self,
        sst: Option<&super::shared_strings::SharedStringTableBuilder>,
        table_r_ids: &[String],
    ) -> Result<Vec<u8>> {
        let mut buf: Vec<u8> = Vec::with_capacity(1024 + self.rows.len() * 128);
        let mut writer = Writer::new(&mut buf);
        let mut ibuf = itoa::Buffer::new();

        let map_err = |e: std::io::Error| ModernXlsxError::XmlWrite(e.to_string());

        // XML declaration.
        writer
            .write_event(Event::Decl(BytesDecl::new("1.0", Some("UTF-8"), Some("yes"))))
            .map_err(map_err)?;

        // <worksheet xmlns="..." xmlns:r="...">
        let mut ws = BytesStart::new("worksheet");
        ws.push_attribute(("xmlns", SPREADSHEET_NS));
        ws.push_attribute(("xmlns:r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships"));
        writer.write_event(Event::Start(ws)).map_err(map_err)?;

        // <sheetPr> — write if tab_color or outline_properties.
        if self.tab_color.is_some() || self.outline_properties.is_some() {
            writer
                .write_event(Event::Start(BytesStart::new("sheetPr")))
                .map_err(map_err)?;
            if let Some(ref color) = self.tab_color {
                let mut tc = BytesStart::new("tabColor");
                tc.push_attribute(("rgb", color.as_str()));
                writer.write_event(Event::Empty(tc)).map_err(map_err)?;
            }
            if let Some(ref op) = self.outline_properties {
                let mut elem = BytesStart::new("outlinePr");
                if !op.summary_below {
                    elem.push_attribute(("summaryBelow", "0"));
                }
                if !op.summary_right {
                    elem.push_attribute(("summaryRight", "0"));
                }
                writer.write_event(Event::Empty(elem)).map_err(map_err)?;
            }
            writer
                .write_event(Event::End(BytesEnd::new("sheetPr")))
                .map_err(map_err)?;
        }

        // <sheetViews> — if frozen_pane or split_pane is present.
        // Split pane takes priority over frozen pane (mutually exclusive).
        let has_pane = self.split_pane.is_some() || self.frozen_pane.is_some();
        if has_pane {
            writer
                .write_event(Event::Start(BytesStart::new("sheetViews")))
                .map_err(map_err)?;

            let mut sv = BytesStart::new("sheetView");
            sv.push_attribute(("workbookViewId", "0"));
            writer.write_event(Event::Start(sv)).map_err(map_err)?;

            if let Some(ref sp) = self.split_pane {
                // --- Split pane ---
                let mut pane_elem = BytesStart::new("pane");
                if let Some(x) = sp.vertical {
                    let s = format_f64(x);
                    pane_elem.push_attribute(("xSplit", s.as_str()));
                }
                if let Some(y) = sp.horizontal {
                    let s = format_f64(y);
                    pane_elem.push_attribute(("ySplit", s.as_str()));
                }
                if let Some(ref tlc) = sp.top_left_cell {
                    pane_elem.push_attribute(("topLeftCell", tlc.as_str()));
                }
                if let Some(ref ap) = sp.active_pane {
                    pane_elem.push_attribute(("activePane", ap.as_str()));
                }
                pane_elem.push_attribute(("state", "split"));
                writer.write_event(Event::Empty(pane_elem)).map_err(map_err)?;
            } else if let Some(ref pane) = self.frozen_pane {
                // --- Frozen pane (existing logic preserved exactly) ---
                let mut pane_elem = BytesStart::new("pane");
                if pane.cols > 0 {
                    pane_elem.push_attribute(("xSplit", ibuf.format(pane.cols)));
                }
                if pane.rows > 0 {
                    pane_elem.push_attribute(("ySplit", ibuf.format(pane.rows)));
                }
                let mut top_left = col_index_to_letter(pane.cols + 1);
                top_left.push_str(ibuf.format(pane.rows + 1));
                pane_elem.push_attribute(("topLeftCell", top_left.as_str()));
                let active_pane = match (pane.rows > 0, pane.cols > 0) {
                    (true, true) => "bottomRight",
                    (true, false) => "bottomLeft",
                    (false, true) => "topRight",
                    (false, false) => "bottomLeft",
                };
                pane_elem.push_attribute(("activePane", active_pane));
                pane_elem.push_attribute(("state", "frozen"));
                writer.write_event(Event::Empty(pane_elem)).map_err(map_err)?;
            }

            // Write <selection> elements for each pane selection.
            for sel in &self.pane_selections {
                let mut sel_elem = BytesStart::new("selection");
                if let Some(ref p) = sel.pane {
                    sel_elem.push_attribute(("pane", p.as_str()));
                }
                if let Some(ref ac) = sel.active_cell {
                    sel_elem.push_attribute(("activeCell", ac.as_str()));
                }
                if let Some(ref sq) = sel.sqref {
                    sel_elem.push_attribute(("sqref", sq.as_str()));
                }
                writer.write_event(Event::Empty(sel_elem)).map_err(map_err)?;
            }

            writer.write_event(Event::End(BytesEnd::new("sheetView"))).map_err(map_err)?;
            writer.write_event(Event::End(BytesEnd::new("sheetViews"))).map_err(map_err)?;
        }

        // <cols> — only if non-empty.
        if !self.columns.is_empty() {
            writer
                .write_event(Event::Start(BytesStart::new("cols")))
                .map_err(map_err)?;

            for col in &self.columns {
                let mut elem = BytesStart::new("col");
                elem.push_attribute(("min", ibuf.format(col.min)));
                elem.push_attribute(("max", ibuf.format(col.max)));
                let width_s = format_f64(col.width);
                elem.push_attribute(("width", width_s.as_str()));
                if col.custom_width {
                    elem.push_attribute(("customWidth", "1"));
                }
                if col.hidden {
                    elem.push_attribute(("hidden", "1"));
                }
                if let Some(level) = col.outline_level
                    && level > 0
                {
                    elem.push_attribute(("outlineLevel", ibuf.format(level)));
                }
                if col.collapsed {
                    elem.push_attribute(("collapsed", "1"));
                }
                writer.write_event(Event::Empty(elem)).map_err(map_err)?;
            }

            writer
                .write_event(Event::End(BytesEnd::new("cols")))
                .map_err(map_err)?;
        }

        // <sheetData>
        writer
            .write_event(Event::Start(BytesStart::new("sheetData")))
            .map_err(map_err)?;

        for row in &self.rows {
            let mut row_elem = BytesStart::new("row");
            row_elem.push_attribute(("r", ibuf.format(row.index)));
            if let Some(ht) = row.height {
                let ht_s = format_f64(ht);
                row_elem.push_attribute(("ht", ht_s.as_str()));
            }
            if row.hidden {
                row_elem.push_attribute(("hidden", "1"));
            }
            if let Some(level) = row.outline_level
                && level > 0
            {
                row_elem.push_attribute(("outlineLevel", ibuf.format(level)));
            }
            if row.collapsed {
                row_elem.push_attribute(("collapsed", "1"));
            }

            if row.cells.is_empty() {
                writer
                    .write_event(Event::Empty(row_elem))
                    .map_err(map_err)?;
            } else {
                writer
                    .write_event(Event::Start(row_elem))
                    .map_err(map_err)?;

                for cell in &row.cells {
                    let mut c_elem = BytesStart::new("c");
                    c_elem.push_attribute(("r", cell.reference.as_str()));

                    // Only write t attribute if not Number (the default).
                    // Stub cells are not written to XML.
                    match cell.cell_type {
                        CellType::Number | CellType::Stub => {}
                        CellType::SharedString => c_elem.push_attribute(("t", "s")),
                        CellType::Boolean => c_elem.push_attribute(("t", "b")),
                        CellType::Error => c_elem.push_attribute(("t", "e")),
                        CellType::FormulaStr => c_elem.push_attribute(("t", "str")),
                        CellType::InlineStr => c_elem.push_attribute(("t", "inlineStr")),
                    }

                    // Only write s attribute if style_index is present.
                    if let Some(si) = cell.style_index {
                        c_elem.push_attribute(("s", ibuf.format(si)));
                    }

                    let has_inline = cell.cell_type == CellType::InlineStr && cell.inline_string.is_some();
                    let has_content = cell.formula.is_some() || cell.value.is_some() || has_inline;

                    if has_content {
                        writer
                            .write_event(Event::Start(c_elem))
                            .map_err(map_err)?;

                        // <f>...</f>
                        if let Some(ref formula) = cell.formula {
                            let mut f_elem = BytesStart::new("f");
                            if let Some(ref ft) = cell.formula_type {
                                f_elem.push_attribute(("t", ft.as_str()));
                            }
                            if let Some(ref fr) = cell.formula_ref {
                                f_elem.push_attribute(("ref", fr.as_str()));
                            }
                            if let Some(si) = cell.shared_index {
                                f_elem.push_attribute(("si", ibuf.format(si)));
                            }
                            if cell.dynamic_array == Some(true) {
                                f_elem.push_attribute(("cm", "1"));
                            }
                            writer
                                .write_event(Event::Start(f_elem))
                                .map_err(map_err)?;
                            writer
                                .write_event(Event::Text(BytesText::new(formula)))
                                .map_err(map_err)?;
                            writer
                                .write_event(Event::End(BytesEnd::new("f")))
                                .map_err(map_err)?;
                        } else if cell.formula_type.is_some() || cell.shared_index.is_some() {
                            // Self-closing <f> with attributes but no text (shared formula reference).
                            let mut f_elem = BytesStart::new("f");
                            if let Some(ref ft) = cell.formula_type {
                                f_elem.push_attribute(("t", ft.as_str()));
                            }
                            if let Some(ref fr) = cell.formula_ref {
                                f_elem.push_attribute(("ref", fr.as_str()));
                            }
                            if let Some(si) = cell.shared_index {
                                f_elem.push_attribute(("si", ibuf.format(si)));
                            }
                            if cell.dynamic_array == Some(true) {
                                f_elem.push_attribute(("cm", "1"));
                            }
                            writer
                                .write_event(Event::Empty(f_elem))
                                .map_err(map_err)?;
                        }

                        // <is><t>...</t></is> for inline strings.
                        if has_inline {
                            writer
                                .write_event(Event::Start(BytesStart::new("is")))
                                .map_err(map_err)?;
                            writer
                                .write_event(Event::Start(BytesStart::new("t")))
                                .map_err(map_err)?;
                            writer
                                .write_event(Event::Text(BytesText::new(
                                    cell.inline_string.as_ref().unwrap(),
                                )))
                                .map_err(map_err)?;
                            writer
                                .write_event(Event::End(BytesEnd::new("t")))
                                .map_err(map_err)?;
                            writer
                                .write_event(Event::End(BytesEnd::new("is")))
                                .map_err(map_err)?;
                        }

                        // <v>...</v>
                        if let Some(ref value) = cell.value {
                            writer
                                .write_event(Event::Start(BytesStart::new("v")))
                                .map_err(map_err)?;
                            // If an SST builder is provided and this is a SharedString cell,
                            // write the SST index instead of the raw string value.
                            if cell.cell_type == CellType::SharedString {
                                if let Some(sst_builder) = sst {
                                    let idx = sst_builder.get_index(value).ok_or_else(|| {
                                        ModernXlsxError::InvalidCellValue(format!(
                                            "SharedString cell has unmapped value: {}",
                                            value
                                        ))
                                    })?;
                                    writer
                                        .write_event(Event::Text(BytesText::new(ibuf.format(idx))))
                                        .map_err(map_err)?;
                                } else {
                                    writer
                                        .write_event(Event::Text(BytesText::new(value)))
                                        .map_err(map_err)?;
                                }
                            } else {
                                writer
                                    .write_event(Event::Text(BytesText::new(value)))
                                    .map_err(map_err)?;
                            }
                            writer
                                .write_event(Event::End(BytesEnd::new("v")))
                                .map_err(map_err)?;
                        }

                        writer
                            .write_event(Event::End(BytesEnd::new("c")))
                            .map_err(map_err)?;
                    } else {
                        writer
                            .write_event(Event::Empty(c_elem))
                            .map_err(map_err)?;
                    }
                }

                writer
                    .write_event(Event::End(BytesEnd::new("row")))
                    .map_err(map_err)?;
            }
        }

        // </sheetData>
        writer
            .write_event(Event::End(BytesEnd::new("sheetData")))
            .map_err(map_err)?;

        // <mergeCells> — only if non-empty.
        if !self.merge_cells.is_empty() {
            let mut mc = BytesStart::new("mergeCells");
            mc.push_attribute(("count", ibuf.format(self.merge_cells.len())));
            writer.write_event(Event::Start(mc)).map_err(map_err)?;

            for ref_str in &self.merge_cells {
                let mut elem = BytesStart::new("mergeCell");
                elem.push_attribute(("ref", ref_str.as_str()));
                writer.write_event(Event::Empty(elem)).map_err(map_err)?;
            }

            writer
                .write_event(Event::End(BytesEnd::new("mergeCells")))
                .map_err(map_err)?;
        }

        // <autoFilter> — only if present.
        if let Some(ref af) = self.auto_filter {
            let mut elem = BytesStart::new("autoFilter");
            elem.push_attribute(("ref", af.range.as_str()));
            if af.filter_columns.is_empty() {
                writer.write_event(Event::Empty(elem)).map_err(map_err)?;
            } else {
                writer.write_event(Event::Start(elem)).map_err(map_err)?;
                for fc in &af.filter_columns {
                    let mut fc_elem = BytesStart::new("filterColumn");
                    fc_elem.push_attribute(("colId", ibuf.format(fc.col_id)));
                    writer.write_event(Event::Start(fc_elem)).map_err(map_err)?;
                    writer.write_event(Event::Start(BytesStart::new("filters"))).map_err(map_err)?;
                    for fv in &fc.filters {
                        let mut f_elem = BytesStart::new("filter");
                        f_elem.push_attribute(("val", fv.as_str()));
                        writer.write_event(Event::Empty(f_elem)).map_err(map_err)?;
                    }
                    writer.write_event(Event::End(BytesEnd::new("filters"))).map_err(map_err)?;
                    writer.write_event(Event::End(BytesEnd::new("filterColumn"))).map_err(map_err)?;
                }
                writer.write_event(Event::End(BytesEnd::new("autoFilter"))).map_err(map_err)?;
            }
        }

        // <dataValidations> — only if non-empty.
        if !self.data_validations.is_empty() {
            let mut dvs = BytesStart::new("dataValidations");
            dvs.push_attribute(("count", ibuf.format(self.data_validations.len())));
            writer.write_event(Event::Start(dvs)).map_err(map_err)?;

            for dv in &self.data_validations {
                let mut elem = BytesStart::new("dataValidation");
                if let Some(ref t) = dv.validation_type {
                    elem.push_attribute(("type", t.as_str()));
                }
                if let Some(ref op) = dv.operator {
                    elem.push_attribute(("operator", op.as_str()));
                }
                if let Some(ab) = dv.allow_blank {
                    elem.push_attribute(("allowBlank", if ab { "1" } else { "0" }));
                }
                if let Some(sem) = dv.show_error_message {
                    elem.push_attribute(("showErrorMessage", if sem { "1" } else { "0" }));
                }
                if let Some(ref et) = dv.error_title {
                    elem.push_attribute(("errorTitle", et.as_str()));
                }
                if let Some(ref em) = dv.error_message {
                    elem.push_attribute(("error", em.as_str()));
                }
                if let Some(sim) = dv.show_input_message {
                    elem.push_attribute(("showInputMessage", if sim { "1" } else { "0" }));
                }
                if let Some(ref pt) = dv.prompt_title {
                    elem.push_attribute(("promptTitle", pt.as_str()));
                }
                if let Some(ref p) = dv.prompt {
                    elem.push_attribute(("prompt", p.as_str()));
                }
                elem.push_attribute(("sqref", dv.sqref.as_str()));

                let has_formulas = dv.formula1.is_some() || dv.formula2.is_some();
                if has_formulas {
                    writer
                        .write_event(Event::Start(elem))
                        .map_err(map_err)?;

                    if let Some(ref f1) = dv.formula1 {
                        writer
                            .write_event(Event::Start(BytesStart::new("formula1")))
                            .map_err(map_err)?;
                        writer
                            .write_event(Event::Text(BytesText::from_escaped(f1.as_str())))
                            .map_err(map_err)?;
                        writer
                            .write_event(Event::End(BytesEnd::new("formula1")))
                            .map_err(map_err)?;
                    }
                    if let Some(ref f2) = dv.formula2 {
                        writer
                            .write_event(Event::Start(BytesStart::new("formula2")))
                            .map_err(map_err)?;
                        writer
                            .write_event(Event::Text(BytesText::from_escaped(f2.as_str())))
                            .map_err(map_err)?;
                        writer
                            .write_event(Event::End(BytesEnd::new("formula2")))
                            .map_err(map_err)?;
                    }

                    writer
                        .write_event(Event::End(BytesEnd::new("dataValidation")))
                        .map_err(map_err)?;
                } else {
                    writer
                        .write_event(Event::Empty(elem))
                        .map_err(map_err)?;
                }
            }

            writer
                .write_event(Event::End(BytesEnd::new("dataValidations")))
                .map_err(map_err)?;
        }

        // <conditionalFormatting> — only if non-empty.
        for cf in &self.conditional_formatting {
            let mut cf_elem = BytesStart::new("conditionalFormatting");
            cf_elem.push_attribute(("sqref", cf.sqref.as_str()));
            writer
                .write_event(Event::Start(cf_elem))
                .map_err(map_err)?;

            for rule in &cf.rules {
                let mut rule_elem = BytesStart::new("cfRule");
                rule_elem.push_attribute(("type", rule.rule_type.as_str()));
                if let Some(dxf_id) = rule.dxf_id {
                    rule_elem.push_attribute(("dxfId", ibuf.format(dxf_id)));
                }
                rule_elem.push_attribute(("priority", ibuf.format(rule.priority)));
                if let Some(ref op) = rule.operator {
                    rule_elem.push_attribute(("operator", op.as_str()));
                }

                let has_children = rule.formula.is_some()
                    || rule.color_scale.is_some()
                    || rule.data_bar.is_some()
                    || rule.icon_set.is_some();

                if has_children {
                    writer
                        .write_event(Event::Start(rule_elem))
                        .map_err(map_err)?;

                    if let Some(ref formula) = rule.formula {
                        writer
                            .write_event(Event::Start(BytesStart::new("formula")))
                            .map_err(map_err)?;
                        writer
                            .write_event(Event::Text(BytesText::new(formula)))
                            .map_err(map_err)?;
                        writer
                            .write_event(Event::End(BytesEnd::new("formula")))
                            .map_err(map_err)?;
                    }

                    if let Some(ref cs) = rule.color_scale {
                        writer.write_event(Event::Start(BytesStart::new("colorScale"))).map_err(map_err)?;
                        for cfvo in &cs.cfvos {
                            let mut cfvo_elem = BytesStart::new("cfvo");
                            cfvo_elem.push_attribute(("type", cfvo.cfvo_type.as_str()));
                            if let Some(ref v) = cfvo.val {
                                cfvo_elem.push_attribute(("val", v.as_str()));
                            }
                            writer.write_event(Event::Empty(cfvo_elem)).map_err(map_err)?;
                        }
                        for color in &cs.colors {
                            let mut color_elem = BytesStart::new("color");
                            color_elem.push_attribute(("rgb", color.as_str()));
                            writer.write_event(Event::Empty(color_elem)).map_err(map_err)?;
                        }
                        writer.write_event(Event::End(BytesEnd::new("colorScale"))).map_err(map_err)?;
                    }

                    if let Some(ref db) = rule.data_bar {
                        writer.write_event(Event::Start(BytesStart::new("dataBar"))).map_err(map_err)?;
                        for cfvo in &db.cfvos {
                            let mut cfvo_elem = BytesStart::new("cfvo");
                            cfvo_elem.push_attribute(("type", cfvo.cfvo_type.as_str()));
                            if let Some(ref v) = cfvo.val {
                                cfvo_elem.push_attribute(("val", v.as_str()));
                            }
                            writer.write_event(Event::Empty(cfvo_elem)).map_err(map_err)?;
                        }
                        let mut color_elem = BytesStart::new("color");
                        color_elem.push_attribute(("rgb", db.color.as_str()));
                        writer.write_event(Event::Empty(color_elem)).map_err(map_err)?;
                        writer.write_event(Event::End(BytesEnd::new("dataBar"))).map_err(map_err)?;
                    }

                    if let Some(ref is) = rule.icon_set {
                        let mut is_elem = BytesStart::new("iconSet");
                        if let Some(ref ist) = is.icon_set_type {
                            is_elem.push_attribute(("iconSet", ist.as_str()));
                        }
                        writer.write_event(Event::Start(is_elem)).map_err(map_err)?;
                        for cfvo in &is.cfvos {
                            let mut cfvo_elem = BytesStart::new("cfvo");
                            cfvo_elem.push_attribute(("type", cfvo.cfvo_type.as_str()));
                            if let Some(ref v) = cfvo.val {
                                cfvo_elem.push_attribute(("val", v.as_str()));
                            }
                            writer.write_event(Event::Empty(cfvo_elem)).map_err(map_err)?;
                        }
                        writer.write_event(Event::End(BytesEnd::new("iconSet"))).map_err(map_err)?;
                    }

                    writer
                        .write_event(Event::End(BytesEnd::new("cfRule")))
                        .map_err(map_err)?;
                } else {
                    writer
                        .write_event(Event::Empty(rule_elem))
                        .map_err(map_err)?;
                }
            }

            writer
                .write_event(Event::End(BytesEnd::new("conditionalFormatting")))
                .map_err(map_err)?;
        }

        // <hyperlinks> — only if non-empty.
        if !self.hyperlinks.is_empty() {
            writer
                .write_event(Event::Start(BytesStart::new("hyperlinks")))
                .map_err(map_err)?;
            for hl in &self.hyperlinks {
                let mut elem = BytesStart::new("hyperlink");
                elem.push_attribute(("ref", hl.cell_ref.as_str()));
                if let Some(ref loc) = hl.location {
                    elem.push_attribute(("location", loc.as_str()));
                }
                if let Some(ref disp) = hl.display {
                    elem.push_attribute(("display", disp.as_str()));
                }
                if let Some(ref tt) = hl.tooltip {
                    elem.push_attribute(("tooltip", tt.as_str()));
                }
                writer.write_event(Event::Empty(elem)).map_err(map_err)?;
            }
            writer
                .write_event(Event::End(BytesEnd::new("hyperlinks")))
                .map_err(map_err)?;
        }

        // <pageSetup> — only if present.
        if let Some(ref ps) = self.page_setup {
            let mut elem = BytesStart::new("pageSetup");
            if let Some(paper) = ps.paper_size {
                elem.push_attribute(("paperSize", ibuf.format(paper)));
            }
            if let Some(ref orient) = ps.orientation {
                elem.push_attribute(("orientation", orient.as_str()));
            }
            if let Some(ftw) = ps.fit_to_width {
                elem.push_attribute(("fitToWidth", ibuf.format(ftw)));
            }
            if let Some(fth) = ps.fit_to_height {
                elem.push_attribute(("fitToHeight", ibuf.format(fth)));
            }
            if let Some(sc) = ps.scale {
                elem.push_attribute(("scale", ibuf.format(sc)));
            }
            if let Some(fpn) = ps.first_page_number {
                elem.push_attribute(("firstPageNumber", ibuf.format(fpn)));
            }
            if let Some(hdpi) = ps.horizontal_dpi {
                elem.push_attribute(("horizontalDpi", ibuf.format(hdpi)));
            }
            if let Some(vdpi) = ps.vertical_dpi {
                elem.push_attribute(("verticalDpi", ibuf.format(vdpi)));
            }
            writer.write_event(Event::Empty(elem)).map_err(map_err)?;
        }

        // <sheetProtection> — only if present.
        if let Some(ref sp) = self.sheet_protection {
            let mut elem = BytesStart::new("sheetProtection");
            if sp.sheet {
                elem.push_attribute(("sheet", "1"));
            }
            if sp.objects {
                elem.push_attribute(("objects", "1"));
            }
            if sp.scenarios {
                elem.push_attribute(("scenarios", "1"));
            }
            if let Some(ref pw) = sp.password {
                elem.push_attribute(("password", pw.as_str()));
            }
            if sp.format_cells {
                elem.push_attribute(("formatCells", "1"));
            }
            if sp.format_columns {
                elem.push_attribute(("formatColumns", "1"));
            }
            if sp.format_rows {
                elem.push_attribute(("formatRows", "1"));
            }
            if sp.insert_columns {
                elem.push_attribute(("insertColumns", "1"));
            }
            if sp.insert_rows {
                elem.push_attribute(("insertRows", "1"));
            }
            if sp.delete_columns {
                elem.push_attribute(("deleteColumns", "1"));
            }
            if sp.delete_rows {
                elem.push_attribute(("deleteRows", "1"));
            }
            if sp.sort {
                elem.push_attribute(("sort", "1"));
            }
            if sp.auto_filter {
                elem.push_attribute(("autoFilter", "1"));
            }
            writer.write_event(Event::Empty(elem)).map_err(map_err)?;
        }

        // <headerFooter> — only if present.
        if let Some(ref hf) = self.header_footer {
            let mut elem = BytesStart::new("headerFooter");
            if hf.different_odd_even {
                elem.push_attribute(("differentOddEven", "1"));
            }
            if hf.different_first {
                elem.push_attribute(("differentFirst", "1"));
            }
            if !hf.scale_with_doc {
                elem.push_attribute(("scaleWithDoc", "0"));
            }
            if !hf.align_with_margins {
                elem.push_attribute(("alignWithMargins", "0"));
            }
            writer.write_event(Event::Start(elem)).map_err(map_err)?;

            for (tag, val) in [
                ("oddHeader", &hf.odd_header),
                ("oddFooter", &hf.odd_footer),
                ("evenHeader", &hf.even_header),
                ("evenFooter", &hf.even_footer),
                ("firstHeader", &hf.first_header),
                ("firstFooter", &hf.first_footer),
            ] {
                if let Some(text) = val {
                    writer.write_event(Event::Start(BytesStart::new(tag))).map_err(map_err)?;
                    writer.write_event(Event::Text(BytesText::new(text))).map_err(map_err)?;
                    writer.write_event(Event::End(BytesEnd::new(tag))).map_err(map_err)?;
                }
            }

            writer.write_event(Event::End(BytesEnd::new("headerFooter"))).map_err(map_err)?;
        }

        // <tableParts> — only if table rIds are provided.
        if !table_r_ids.is_empty() {
            let mut tp = BytesStart::new("tableParts");
            tp.push_attribute(("count", ibuf.format(table_r_ids.len())));
            writer.write_event(Event::Start(tp)).map_err(map_err)?;
            for rid in table_r_ids {
                let mut part = BytesStart::new("tablePart");
                part.push_attribute(("r:id", rid.as_str()));
                writer.write_event(Event::Empty(part)).map_err(map_err)?;
            }
            writer
                .write_event(Event::End(BytesEnd::new("tableParts")))
                .map_err(map_err)?;
        }

        // </worksheet>
        writer
            .write_event(Event::End(BytesEnd::new("worksheet")))
            .map_err(map_err)?;

        Ok(buf)
    }
}

/// Parse a `<col>` element into a `ColumnInfo`.
fn parse_col_element(e: &BytesStart<'_>) -> ColumnInfo {
    let mut min: u32 = 1;
    let mut max: u32 = 1;
    let mut width: f64 = 8.43;
    let mut hidden = false;
    let mut custom_width = false;
    let mut outline_level: Option<u8> = None;
    let mut collapsed = false;

    for attr in e.attributes().flatten() {
        let ln = attr.key.local_name();
        match ln.as_ref() {
            b"min" => {
                let val = std::str::from_utf8(&attr.value).unwrap_or_default();
                min = val.parse::<u32>().unwrap_or(1);
            }
            b"max" => {
                let val = std::str::from_utf8(&attr.value).unwrap_or_default();
                max = val.parse::<u32>().unwrap_or(1);
            }
            b"width" => {
                let val = std::str::from_utf8(&attr.value).unwrap_or_default();
                width = val.parse::<f64>().unwrap_or(8.43);
            }
            b"hidden" => {
                let val = std::str::from_utf8(&attr.value).unwrap_or_default();
                hidden = val == "1" || val.eq_ignore_ascii_case("true");
            }
            b"customWidth" => {
                let val = std::str::from_utf8(&attr.value).unwrap_or_default();
                custom_width = val == "1" || val.eq_ignore_ascii_case("true");
            }
            b"outlineLevel" => {
                outline_level = std::str::from_utf8(&attr.value)
                    .ok()
                    .and_then(|v| v.parse::<u8>().ok())
                    .filter(|&v| v > 0);
            }
            b"collapsed" => {
                collapsed = std::str::from_utf8(&attr.value).unwrap_or("0") == "1";
            }
            _ => {}
        }
    }

    ColumnInfo {
        min,
        max,
        width,
        hidden,
        custom_width,
        outline_level,
        collapsed,
    }
}

/// Convert a 1-based column index to a letter string (1 -> "A", 26 -> "Z", 27 -> "AA").
fn col_index_to_letter(col: u32) -> String {
    crate::ooxml::cell::col_to_letters(col.saturating_sub(1))
}

/// Format an f64 to a string, removing trailing zeros after the decimal point
/// but always keeping at least one decimal place if the number is not an integer.
fn format_f64(val: f64) -> String {
    if val == val.floor() {
        // Integer value — format without decimal places.
        format!("{}", val as i64)
    } else {
        format!("{val}")
    }
}

#[cfg(test)]
mod tests {
    use super::*;
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
                            formula: None,
                            formula_type: None,
                            formula_ref: None,
                            shared_index: None,
                            inline_string: None,
                            dynamic_array: None,
                        },
                        Cell {
                            reference: "B1".to_string(),
                            cell_type: CellType::Number,
                            style_index: None,
                            value: Some("42.5".to_string()),
                            formula: None,
                            formula_type: None,
                            formula_ref: None,
                            shared_index: None,
                            inline_string: None,
                            dynamic_array: None,
                        },
                        Cell {
                            reference: "C1".to_string(),
                            cell_type: CellType::Number,
                            style_index: None,
                            value: Some("42.5".to_string()),
                            formula: Some("SUM(A1:B1)".to_string()),
                            formula_type: None,
                            formula_ref: None,
                            shared_index: None,
                            inline_string: None,
                            dynamic_array: None,
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
                        formula: None,
                        formula_type: None,
                        formula_ref: None,
                        shared_index: None,
                        inline_string: None,
                        dynamic_array: None,
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
            outline_properties: None,
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
                    formula: None,
                    formula_type: None,
                    formula_ref: None,
                    shared_index: None,
                    inline_string: None,
                    dynamic_array: None,
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
            outline_properties: None,
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
                    formula: None,
                    formula_type: None,
                    formula_ref: None,
                    shared_index: None,
                    inline_string: None,
                    dynamic_array: None,
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
            outline_properties: None,
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
                    formula: None,
                    formula_type: None,
                    formula_ref: None,
                    shared_index: None,
                    inline_string: None,
                    dynamic_array: None,
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
            outline_properties: None,
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
            outline_properties: None,
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
                    formula: None,
                    formula_type: None,
                    formula_ref: None,
                    shared_index: None,
                    inline_string: None,
                    dynamic_array: None,
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
            outline_properties: None,
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
            outline_properties: None,
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
            outline_properties: None,
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
            outline_properties: None,
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
                    style_index: None,
                    value: Some("100".to_string()),
                    formula: Some("SUM(B1:B3*C1:C3)".to_string()),
                    formula_type: Some("array".to_string()),
                    formula_ref: Some("A1:A3".to_string()),
                    shared_index: None,
                    inline_string: None,
                    dynamic_array: None,
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
            outline_properties: None,
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
            outline_properties: None,
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
                    },
                    FilterColumn {
                        col_id: 2,
                        filters: vec!["Red".to_string()],
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
            outline_properties: None,
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
            outline_properties: None,
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
            outline_properties: None,
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
            outline_properties: None,
        };
        let xml = ws.to_xml_with_sst(None, &[]).unwrap();
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
            outline_properties: None,
        };
        let xml = ws.to_xml_with_sst(None, &[]).unwrap();
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

        let xml2 = ws.to_xml_with_sst(None, &[]).unwrap();
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
            outline_properties: None,
        };
        let xml = ws.to_xml_with_sst(None, &[]).unwrap();
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
            outline_properties: Some(OutlineProperties {
                summary_below: false,
                summary_right: true,
            }),
        };
        let xml = ws.to_xml_with_sst(None, &[]).unwrap();
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
            outline_properties: None,
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
            outline_properties: None,
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
            outline_properties: None,
        };
        let xml = ws.to_xml().unwrap();
        let xml_str = std::str::from_utf8(&xml).unwrap();
        assert!(xml_str.contains(r#"state="split""#));
        assert!(!xml_str.contains(r#"state="frozen""#));
    }
}
