//! The worksheet part, `xl/worksheets/sheet{n}.xml` (ECMA-376 Part 1 §18.3):
//! [`WorksheetXml`] and its child types.
//!
//! One SAX driver (`driver`) reads the part and feeds its rows and cells to
//! one of two sinks, with the SAX helpers in `parser`: for
//! [`WorksheetXml::parse`] and the crate-internal `parse_with_sst` and
//! `parse_with_sst_within`, the struct sink (`sink_struct`) builds the typed
//! model; for the crate-internal `parse_to_json_within`, the JSON sink
//! (`sink_json`) writes the sheet as JSON for `reader::read_xlsx_json` and its
//! variants.
//! [`WorksheetXml::to_xml`], [`WorksheetXml::to_xml_with_sst`] and
//! [`WorksheetXml::to_xml_with_legacy_drawings`] write the part; given the
//! workbook's SST builder, the latter two write each shared-string cell's
//! index as they write the cell, without cloning the worksheet, and the last
//! also names the sheet's VML drawings (`<legacyDrawing>`,
//! `<legacyDrawingHF>`). `SheetContext` carries the workbook's date system
//! and the sheet part's relationships into the driver, for both read paths.

use super::charts::WorksheetChart;
use super::color::Color;
use super::relationships::Relationships;
use super::shared_strings::RichTextRun;
use crate::dates::DateSystem;
use serde::{Deserialize, Serialize};

mod driver;
mod json;
mod parser;
mod sink_json;
mod sink_struct;
mod writer;

/// Workbook-level facts a worksheet part is parsed against.
#[derive(Debug, Clone, Copy, Default)]
pub(crate) struct SheetContext<'a> {
    /// The workbook's date system: `t="d"` cells become serials in it.
    pub date_system: DateSystem,
    /// The sheet part's relationships, for resolving `r:id` references (an
    /// external hyperlink's `<hyperlink r:id>` becomes `Hyperlink::target`).
    pub rels: Option<&'a Relationships>,
}

/// The `<ext uri>` of the x14 `<sparklineGroups>` extension: Excel matches
/// an extension by this exact URI, so the writer uses it.
const SPARKLINE_EXT_URI: &str = "{05C60535-1F16-4fd2-B633-F4F36F0B64E0}";

/// The sparkline `<ext uri>` that 1.2.x wrote (Excel does not recognise it),
/// still read as sparklines.
const LEGACY_SPARKLINE_EXT_URI: &str = "{05C60535-1F16-4fd2-B633-F4F36011B0BD}";

pub(crate) use json::json_escape_to;
// `WorksheetXml`'s inherent methods need no re-export: `parse` (pub) and
// `parse_with_sst`, `parse_with_sst_within` and `parse_to_json_within` (all
// pub(crate)) live in `parser`;
// `to_xml`, `to_xml_with_sst` and `to_xml_with_legacy_drawings` (all pub)
// live in `writer`.

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
    #[serde(
        default = "crate::ooxml::default_true",
        skip_serializing_if = "crate::ooxml::is_true"
    )]
    pub scale_with_doc: bool,
    #[serde(
        default = "crate::ooxml::default_true",
        skip_serializing_if = "crate::ooxml::is_true"
    )]
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
#[derive(Debug, Clone, Default, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct WorksheetXml {
    /// The `<dimension ref>` read from the source file. The writer ignores it
    /// and recomputes the used range from `rows`.
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
    /// The `rgb` of `<sheetPr><tabColor>`; see `tab_color_spec` for the other
    /// colour kinds.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tab_color: Option<String>,
    /// The full `<tabColor>` when it has a theme, indexed, tint or auto
    /// attribute. On write it is used while its `rgb` equals `tab_color` (an
    /// empty string counts as absent); otherwise `tab_color` is written as
    /// plain `rgb`.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tab_color_spec: Option<Color>,
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
    /// Every `<extLst><ext>` other than the sparkline one, kept verbatim for
    /// round-trip. A namespace prefix the source bound outside the `<ext>`
    /// (on `<worksheet>`) is re-declared on the `<ext>` start tag, so each
    /// entry is well-formed on its own. The writer emits them inside
    /// `<extLst>` after the sparkline extension.
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub preserved_extensions: Vec<String>,
}

/// A single row in the worksheet.
#[derive(Debug, Clone, Default, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct Row {
    /// 1-based row index (as in the XML `r` attribute).
    pub index: u32,
    pub cells: Vec<Cell>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub height: Option<f64>,
    /// `customHeight`: `height` was set by hand (ECMA-376 Part 1, §18.3.1.73).
    /// Excel keeps `height` only when this is `true`; without it Excel sizes the
    /// row to its content. The writer emits `customHeight="1"` after `hidden`
    /// only when this is `true` and a height is present.
    #[serde(default, skip_serializing_if = "crate::ooxml::is_false")]
    pub custom_height: bool,
    #[serde(default, skip_serializing_if = "crate::ooxml::is_false")]
    pub hidden: bool,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub outline_level: Option<u8>,
    #[serde(default, skip_serializing_if = "crate::ooxml::is_false")]
    pub collapsed: bool,
    /// `<row s>`: the `cellXfs` index Excel applies to the row's empty cells,
    /// read only when `customFormat` is true (ECMA-376 Part 1, §18.3.1.73),
    /// and written back as `s` with `customFormat="1"`. Named like
    /// [`Cell::style_index`].
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub style_index: Option<u32>,
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
    /// The cell value: the `<v>` text, entities decoded and whitespace kept.
    /// On read a shared-string cell holds its SST entry's text (when the SST
    /// is supplied), an inline-string cell holds the inline text, and a
    /// `t="d"` cell holds its date serial (a value that is not a date stays
    /// text, in a `FormulaStr` cell).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub value: Option<String>,
    /// Raw `<f>` content (the formula element text).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub formula: Option<String>,
    /// Formula type: "normal", "array", "shared" or "dataTable", from the `t`
    /// attribute on `<f>`.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub formula_type: Option<String>,
    /// Range for array/shared formulas, from `ref` attribute on `<f>`.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub formula_ref: Option<String>,
    /// Shared formula index, from `si` attribute on `<f>`.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub shared_index: Option<u32>,
    /// Inline string value from `<is>`: its `<t>` text, or the concatenated
    /// `<t>` text of its `<r>` runs; phonetic `<rPh>` text is not included.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub inline_string: Option<String>,
    /// `<c cm>`: the 1-based index of the cell's record in the workbook's
    /// `cellMetadata` (`xl/metadata.xml`). Excel gives a dynamic-array formula
    /// `cm="1"`, pointing at its `XLDAPR` record.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cell_metadata: Option<u32>,
    /// Whether this is a dynamic array formula: set on read whenever `<c>`
    /// carries `cm` (or a 1.2.x file has `cm="1"` on `<f>`); on write it yields
    /// `cm="1"` on `<c>` when `cell_metadata` is unset.
    ///
    /// Deprecated, removed in 2.0; use `cellMetadata`.
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
    /// One-dimensional data table is a row (`dtr` attribute on `<f>`): `r1`
    /// is a row input cell.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub formula_dtr: Option<bool>,
    /// Data table input cell 1 was deleted (`del1` attribute on `<f>`).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub formula_del1: Option<bool>,
    /// Data table input cell 2 was deleted (`del2` attribute on `<f>`).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub formula_del2: Option<bool>,
    /// Mirrors `formula_del1` on read (and the 1.2.x `dtr1` attribute); on
    /// write it yields `del1="1"` when `formula_del1` is unset.
    ///
    /// Deprecated, removed in 2.0; use `formulaDel1`.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub formula_dtr1: Option<bool>,
    /// Mirrors `formula_del2` on read (and the 1.2.x `dtr2` attribute); on
    /// write it yields `del2="1"` when `formula_del2` is unset.
    ///
    /// Deprecated, removed in 2.0; use `formulaDel2`.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub formula_dtr2: Option<bool>,
    /// Rich text runs for this cell (when the cell contains mixed-format text).
    /// On read, a shared-string cell gets the runs of the SST entry its index
    /// names (never those of another entry with the same text); inline-string
    /// runs are not modelled. Set directly by the API to write rich text.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub rich_text: Option<Vec<RichTextRun>>,
}

/// The type of a cell, determined by the `t` attribute on the `<c>` element.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum CellType {
    /// `t="s"` — a shared string. In the file `<v>` is the SST index; in the
    /// model `value` is the entry's text when the parser has the SST (else
    /// the index), and `to_xml_with_sst`, given the workbook's SST builder,
    /// writes the entry's index (without a builder, and from `to_xml`,
    /// `value` is written as it is).
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
    /// An explicitly empty cell (SheetJS type "z"), set through the API; the
    /// reader never produces it, since no `t` value maps to it. Written as
    /// `<c>` with no `t` attribute (a bare stub is `<c r="A1"/>`; `s`, `cm`
    /// and any content are written as for other cells) and left out of the
    /// `<dimension>` range.
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
    #[serde(
        default = "crate::ooxml::default_true",
        skip_serializing_if = "crate::ooxml::is_true"
    )]
    pub show_grid_lines: bool,
    /// Whether row and column headers are visible (default: true).
    #[serde(
        default = "crate::ooxml::default_true",
        skip_serializing_if = "crate::ooxml::is_true"
    )]
    pub show_row_col_headers: bool,
    /// Whether zero values are displayed (default: true).
    #[serde(
        default = "crate::ooxml::default_true",
        skip_serializing_if = "crate::ooxml::is_true"
    )]
    pub show_zeros: bool,
    /// Right-to-left display mode (default: false).
    #[serde(default, skip_serializing_if = "crate::ooxml::is_false")]
    pub right_to_left: bool,
    /// Whether this sheet tab is selected (default: false).
    #[serde(default, skip_serializing_if = "crate::ooxml::is_false")]
    pub tab_selected: bool,
    /// Whether the ruler is shown in Page Layout view (default: true).
    #[serde(
        default = "crate::ooxml::default_true",
        skip_serializing_if = "crate::ooxml::is_true"
    )]
    pub show_ruler: bool,
    /// Whether outline (grouping) symbols are shown (default: true).
    #[serde(
        default = "crate::ooxml::default_true",
        skip_serializing_if = "crate::ooxml::is_true"
    )]
    pub show_outline_symbols: bool,
    /// Whether white space around the page is shown in Page Layout view (default: true).
    #[serde(
        default = "crate::ooxml::default_true",
        skip_serializing_if = "crate::ooxml::is_true"
    )]
    pub show_white_space: bool,
    /// Whether the default grid color is used (default: true).
    #[serde(
        default = "crate::ooxml::default_true",
        skip_serializing_if = "crate::ooxml::is_true"
    )]
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
    /// Indexed colour (ICV) of the gridlines and row/column headings.
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
    #[serde(
        default = "default_sparkline_type",
        skip_serializing_if = "is_default_sparkline_type"
    )]
    pub sparkline_type: String,
    /// Individual sparklines in this group.
    pub sparklines: Vec<Sparkline>,
    /// Series color (RGB hex, e.g. "FF376092"): the `rgb` of
    /// `<x14:colorSeries>`; see `color_series_spec` for the other colour kinds.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub color_series: Option<String>,
    /// The full `<x14:colorSeries>` when it has a theme, indexed, tint or auto
    /// attribute. On write it is used while its `rgb` equals `color_series`
    /// (an empty string counts as absent); otherwise `color_series` is written
    /// as plain `rgb`. The other seven `*_spec` fields work the same way.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub color_series_spec: Option<Color>,
    /// Negative value color.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub color_negative: Option<String>,
    /// The full `<x14:colorNegative>`; see `color_series_spec`.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub color_negative_spec: Option<Color>,
    /// Axis color.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub color_axis: Option<String>,
    /// The full `<x14:colorAxis>`; see `color_series_spec`.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub color_axis_spec: Option<Color>,
    /// Marker color.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub color_markers: Option<String>,
    /// The full `<x14:colorMarkers>`; see `color_series_spec`.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub color_markers_spec: Option<Color>,
    /// First point color.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub color_first: Option<String>,
    /// The full `<x14:colorFirst>`; see `color_series_spec`.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub color_first_spec: Option<Color>,
    /// Last point color.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub color_last: Option<String>,
    /// The full `<x14:colorLast>`; see `color_series_spec`.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub color_last_spec: Option<Color>,
    /// High point color.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub color_high: Option<String>,
    /// The full `<x14:colorHigh>`; see `color_series_spec`.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub color_high_spec: Option<Color>,
    /// Low point color.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub color_low: Option<String>,
    /// The full `<x14:colorLow>`; see `color_series_spec`.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub color_low_spec: Option<Color>,
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
    /// Manual minimum value for axis scaling; it takes effect only with
    /// `min_axis_type` `"custom"`, which the writer emits when
    /// `min_axis_type` is unset.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub manual_min: Option<f64>,
    /// Manual maximum value for axis scaling; it takes effect only with
    /// `max_axis_type` `"custom"`, which the writer emits when
    /// `max_axis_type` is unset.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub manual_max: Option<f64>,
    /// Right-to-left display.
    #[serde(default, skip_serializing_if = "std::ops::Not::not")]
    pub right_to_left: bool,
    /// `minAxisType`: how the vertical axis minimum is chosen, `"individual"`
    /// (per sparkline, the default), `"group"` (shared by the group) or
    /// `"custom"` (`manual_min`).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub min_axis_type: Option<String>,
    /// `maxAxisType`: as `min_axis_type`, for the maximum (`manual_max`).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub max_axis_type: Option<String>,
    /// `dateAxis`: the horizontal axis is a date axis over `date_range`.
    #[serde(default, skip_serializing_if = "std::ops::Not::not")]
    pub date_axis: bool,
    /// `displayHidden`: data in hidden rows and columns is plotted.
    #[serde(default, skip_serializing_if = "std::ops::Not::not")]
    pub display_hidden: bool,
    /// The group's `<xm:f>`: the range holding the dates of a date axis
    /// (`date_axis`), written before `<x14:sparklines>`.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub date_range: Option<String>,
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
            color_series_spec: None,
            color_negative: None,
            color_negative_spec: None,
            color_axis: None,
            color_axis_spec: None,
            color_markers: None,
            color_markers_spec: None,
            color_first: None,
            color_first_spec: None,
            color_last: None,
            color_last_spec: None,
            color_high: None,
            color_high_spec: None,
            color_low: None,
            color_low_spec: None,
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
            min_axis_type: None,
            max_axis_type: None,
            date_axis: false,
            display_hidden: false,
            date_range: None,
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
    /// `<col style>`: the `cellXfs` index Excel applies to the column's empty
    /// cells and shows as the column's format; named like [`Cell::style_index`].
    /// An explicit `style="0"` reads as `Some(0)` and is written back.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub style_index: Option<u32>,
    /// `<col bestFit>`: the column is "best fit", so Excel widens it to show
    /// numbers typed into its cells (its width was never set by hand).
    #[serde(default, skip_serializing_if = "crate::ooxml::is_false")]
    pub best_fit: bool,
    /// `<col phonetic>`: Excel shows the phonetic (furigana) guide for the
    /// column's cells.
    #[serde(default, skip_serializing_if = "crate::ooxml::is_false")]
    pub phonetic: bool,
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
    /// `errorStyle`: what invalid input does, `"stop"` (rejected), `"warning"`
    /// or `"information"` (both allow it). Absent means `"stop"`.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub error_style: Option<String>,
    /// `imeMode`: the input method editor mode while a cell is selected,
    /// such as `"off"`, `"on"` or `"hiragana"`. Absent means `"noControl"`.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ime_mode: Option<String>,
    /// `showDropDown` of a `list` validation. Despite its name, `true` HIDES
    /// the in-cell drop-down arrow. Absent means false (the arrow is shown).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub show_drop_down: Option<bool>,
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
    /// The first `<formula>` child.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub formula: Option<String>,
    /// The second and third `<formula>` children (ECMA-376 Part 1
    /// §18.3.1.10 allows three; a `cellIs` `between` or `notBetween` rule's
    /// upper bound is the second), written after `formula` in order; a rule
    /// with more than three formulas in all fails the write with
    /// `InvalidWorkbook` and is reported by `validate_workbook`. An empty
    /// `<formula/>` reads as `""` in its position. Empty when the rule has one
    /// formula or none.
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub additional_formulas: Vec<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub dxf_id: Option<u32>,
    /// `stopIfTrue`: no rule of lower priority is applied once this one is
    /// true. Absent means false.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub stop_if_true: Option<bool>,
    /// `aboveAverage` of an `aboveAverage` rule: false highlights the cells
    /// below the average. Absent means true.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub above_average: Option<bool>,
    /// `percent` of a `top10` rule: `rank` is a percentage, not a count.
    /// Absent means false.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub percent: Option<bool>,
    /// `bottom` of a `top10` rule: the bottom `rank` values, not the top.
    /// Absent means false.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub bottom: Option<bool>,
    /// `text` of a `containsText`, `notContainsText`, `beginsWith` or
    /// `endsWith` rule: the text the rule looks for.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub text: Option<String>,
    /// `timePeriod` of a `timePeriod` rule, such as `"lastWeek"` or `"today"`.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub time_period: Option<String>,
    /// `rank` of a `top10` rule: how many values (or what percentage, with
    /// `percent`) are highlighted.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub rank: Option<u32>,
    /// `stdDev` of an `aboveAverage` rule: how many standard deviations from
    /// the average a value must lie.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub std_dev: Option<i32>,
    /// `equalAverage` of an `aboveAverage` rule: values equal to the average
    /// also match. Absent means false.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub equal_average: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub color_scale: Option<ColorScale>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub data_bar: Option<DataBar>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub icon_set: Option<IconSet>,
}

/// The most `<formula>` children a `<cfRule>` holds (ECMA-376 Part 1,
/// §18.3.1.10).
pub(crate) const MAX_CF_RULE_FORMULAS: usize = 3;

impl ConditionalFormattingRule {
    /// How many `<formula>` children the writer emits: `formula`, then each
    /// of `additional_formulas`.
    pub(crate) fn formula_count(&self) -> usize {
        usize::from(self.formula.is_some()) + self.additional_formulas.len()
    }

    /// Record the next `<formula>` read, in document order (an empty one as
    /// `""`): the first is `formula`, the second and third
    /// `additional_formulas`; past [`MAX_CF_RULE_FORMULAS`] it is dropped.
    pub(super) fn push_formula(&mut self, text: String) {
        if self.formula.is_none() {
            self.formula = Some(text);
        } else if self.formula_count() < MAX_CF_RULE_FORMULAS {
            self.additional_formulas.push(text);
        }
    }
}

/// Why a `<cfRule>` with `count` formulas (more than
/// [`MAX_CF_RULE_FORMULAS`]) cannot be written, naming the sheet when it is
/// known.
pub(crate) fn cf_formula_count_message(
    sheet: Option<&str>,
    sqref: &str,
    priority: u32,
    count: usize,
) -> String {
    let whose = sheet.map_or_else(String::new, |name| format!(" of sheet '{name}'"));
    format!(
        "the conditional formatting rule with priority {priority} on '{sqref}'{whose} has {count} formulas; a cfRule holds at most {MAX_CF_RULE_FORMULAS} (formula plus two additionalFormulas)"
    )
}

/// A hyperlink in the worksheet (`<hyperlink>`, ECMA-376 Part 1 §18.3.1.47).
#[derive(Debug, Clone, Default, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct Hyperlink {
    pub cell_ref: String,
    /// An external target (URL, `mailto:` address or file), stored as a sheet
    /// relationship with `TargetMode="External"` and referenced by `r:id`.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub target: Option<String>,
    /// A location inside the workbook (`Sheet2!A1`, a defined name) or, with
    /// `target`, a fragment within that target.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub location: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub display: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tooltip: Option<String>,
}

/// Whether `s` starts with a URI scheme, `^[A-Za-z][A-Za-z0-9+.-]*:` (RFC 3986
/// §3.1): `https:`, `mailto:`, `tel:`, `file:`, and a drive path such as
/// `C:\a.xlsx`. A sheet-qualified reference (`Sheet2!A1`) never matches; an
/// unqualified A1 range (`A1:B2`, `A1:$B$2`, `B:D`) fits the scheme grammar
/// and is excluded by [`is_unqualified_a1_range`]. TS `addHyperlink` applies
/// the same rule.
pub(crate) fn has_uri_scheme(s: &str) -> bool {
    let Some((scheme, _)) = s.split_once(':') else {
        return false;
    };
    let mut chars = scheme.chars();
    chars.next().is_some_and(|c| c.is_ascii_alphabetic())
        && chars.all(|c| c.is_ascii_alphanumeric() || matches!(c, '+' | '.' | '-'))
        && !is_unqualified_a1_range(s)
}

/// Whether `s` fully matches `^[A-Za-z]{1,3}[0-9]+:[$]?[A-Za-z]{1,3}[$]?[0-9]+$`
/// (a cell range such as `A1:$B$2`) or `^[A-Za-z]{1,3}:[$]?[A-Za-z]{1,3}$` (a
/// column range such as `B:D`).
fn is_unqualified_a1_range(s: &str) -> bool {
    /// The rest of `t` after 1-3 leading ASCII letters, if it has 1-3.
    fn after_column(t: &str) -> Option<&str> {
        let n = t.bytes().take_while(u8::is_ascii_alphabetic).count();
        (1..=3).contains(&n).then(|| &t[n..])
    }
    fn is_row(t: &str) -> bool {
        !t.is_empty() && t.bytes().all(|b| b.is_ascii_digit())
    }
    let Some((first, last)) = s.split_once(':') else {
        return false;
    };
    let (Some(first_row), Some(last_row)) = (
        after_column(first),
        after_column(last.strip_prefix('$').unwrap_or(last)),
    ) else {
        return false;
    };
    if first_row.is_empty() {
        last_row.is_empty()
    } else {
        is_row(first_row) && is_row(last_row.strip_prefix('$').unwrap_or(last_row))
    }
}

impl Hyperlink {
    /// The external target the writer stores as a sheet relationship:
    /// `target`, or else a `location` that starts with a URI scheme (1.2.x
    /// wrote external links into `location`).
    pub(crate) fn external_target(&self) -> Option<&str> {
        self.target.as_deref().or_else(|| {
            self.location
                .as_deref()
                .filter(|location| has_uri_scheme(location))
        })
    }

    /// The `location` attribute to write: none when `location` itself is the
    /// external target (see [`Self::external_target`]).
    pub(crate) fn written_location(&self) -> Option<&str> {
        let location = self.location.as_deref()?;
        (self.target.is_some() || !has_uri_scheme(location)).then_some(location)
    }
}

/// Page margins in inches (`<pageMargins>`). The writer fills any missing
/// side with Excel's Normal preset (left/right 0.7, top/bottom 0.75,
/// header/footer 0.3).
#[derive(Debug, Clone, Default, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct PageMargins {
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub top: Option<f64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub bottom: Option<f64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub left: Option<f64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub right: Option<f64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub header: Option<f64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub footer: Option<f64>,
}

/// Page setup configuration: `<pageSetup>` attributes, the sheet's
/// `<pageMargins>` (wire path `pageSetup.margins`) and
/// `<sheetPr><pageSetUpPr fitToPage>` (`fit_to_page`).
#[derive(Debug, Clone, Default, Serialize, Deserialize)]
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
    /// `<sheetPr><pageSetUpPr fitToPage>`. Excel applies `fit_to_width` and
    /// `fit_to_height` only when this is `true`; the writer emits the element
    /// for `Some(true)` only.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub fit_to_page: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub margins: Option<PageMargins>,
}

/// Sheet protection settings: `<sheetProtection>` (`CT_SheetProtection`,
/// ECMA-376 Part 1 §18.3.1.85).
///
/// Every flag keeps its wire meaning: `true` locks the action while the sheet
/// is protected, `false` allows it. [`Default`] is the schema default of every
/// attribute (`sheet` false, so the sheet is not protected). The reader starts
/// from those defaults and the writer emits an attribute exactly when it differs
/// from its default.
///
/// The password is stored as the file stores it, either the legacy 16-bit hash
/// (`password`) or the iterated hash (`algorithm_name`, `hash_value`,
/// `salt_value`, `spin_count`), and written back verbatim. `plain_password` is
/// the write-only way to set a new one; [`Debug`] prints it as
/// `Some("<redacted>")`.
#[derive(Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct SheetProtection {
    /// Protects the sheet's contents (schema default false).
    pub sheet: bool,
    /// Locks objects (schema default false).
    pub objects: bool,
    /// Locks scenarios (schema default false).
    pub scenarios: bool,
    /// Legacy 16-bit password hash, 1 to 4 hex digits (`ST_UnsignedShortHex`),
    /// not the password itself. Written verbatim unless `plain_password` is set;
    /// any other value fails the write with `InvalidWorkbook` and is reported by
    /// `validate_workbook`. It is never hashed for the caller: a password such
    /// as `cafe` is also a valid hash, so no rule can tell the two apart.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub password: Option<String>,
    /// Locks formatting cells (schema default true).
    pub format_cells: bool,
    /// Locks formatting columns (schema default true).
    pub format_columns: bool,
    /// Locks formatting rows (schema default true).
    pub format_rows: bool,
    /// Locks inserting columns (schema default true).
    pub insert_columns: bool,
    /// Locks inserting rows (schema default true).
    pub insert_rows: bool,
    /// Locks deleting columns (schema default true).
    pub delete_columns: bool,
    /// Locks deleting rows (schema default true).
    pub delete_rows: bool,
    /// Locks sorting (schema default true).
    pub sort: bool,
    /// Locks AutoFilter use (schema default true).
    pub auto_filter: bool,
    /// Locks inserting hyperlinks (schema default true).
    #[serde(
        default = "crate::ooxml::default_true",
        skip_serializing_if = "crate::ooxml::is_true"
    )]
    pub insert_hyperlinks: bool,
    /// Locks PivotTable use (schema default true).
    #[serde(
        default = "crate::ooxml::default_true",
        skip_serializing_if = "crate::ooxml::is_true"
    )]
    pub pivot_tables: bool,
    /// Blocks selecting locked cells (schema default false).
    #[serde(default, skip_serializing_if = "crate::ooxml::is_false")]
    pub select_locked_cells: bool,
    /// Blocks selecting unlocked cells (schema default false).
    #[serde(default, skip_serializing_if = "crate::ooxml::is_false")]
    pub select_unlocked_cells: bool,
    /// Hash algorithm of `hash_value`, such as `SHA-512`.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub algorithm_name: Option<String>,
    /// Base64 iterated password hash (ECMA-376 Part 1 §18.2.29 scheme).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub hash_value: Option<String>,
    /// Base64 salt of `hash_value`.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub salt_value: Option<String>,
    /// Hash iterations of `hash_value`.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub spin_count: Option<u32>,
    /// Plain-text password to protect the sheet with. Write-only: never read
    /// from a file, never serialized and redacted by `Debug`. When set, the
    /// writer replaces `password` and the stored hash with a fresh one:
    /// SHA-512 with a random 16-byte salt and 100,000 iterations when the
    /// `encryption` feature is enabled, otherwise the legacy 16-bit `password`
    /// hash (ASCII passwords only).
    #[serde(default, skip_serializing)]
    pub plain_password: Option<String>,
}

impl Default for SheetProtection {
    /// The schema default of every `CT_SheetProtection` attribute.
    fn default() -> Self {
        Self {
            sheet: false,
            objects: false,
            scenarios: false,
            password: None,
            format_cells: true,
            format_columns: true,
            format_rows: true,
            insert_columns: true,
            insert_rows: true,
            delete_columns: true,
            delete_rows: true,
            sort: true,
            auto_filter: true,
            insert_hyperlinks: true,
            pivot_tables: true,
            select_locked_cells: false,
            select_unlocked_cells: false,
            algorithm_name: None,
            hash_value: None,
            salt_value: None,
            spin_count: None,
            plain_password: None,
        }
    }
}

impl std::fmt::Debug for SheetProtection {
    /// Every field, with `plain_password` shown as `Some("<redacted>")` or
    /// `None` so that a printed worksheet or workbook never holds a password.
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        // Destructured so that a new field cannot be left out of the output.
        let Self {
            sheet,
            objects,
            scenarios,
            password,
            format_cells,
            format_columns,
            format_rows,
            insert_columns,
            insert_rows,
            delete_columns,
            delete_rows,
            sort,
            auto_filter,
            insert_hyperlinks,
            pivot_tables,
            select_locked_cells,
            select_unlocked_cells,
            algorithm_name,
            hash_value,
            salt_value,
            spin_count,
            plain_password,
        } = self;
        f.debug_struct("SheetProtection")
            .field("sheet", sheet)
            .field("objects", objects)
            .field("scenarios", scenarios)
            .field("password", password)
            .field("format_cells", format_cells)
            .field("format_columns", format_columns)
            .field("format_rows", format_rows)
            .field("insert_columns", insert_columns)
            .field("insert_rows", insert_rows)
            .field("delete_columns", delete_columns)
            .field("delete_rows", delete_rows)
            .field("sort", sort)
            .field("auto_filter", auto_filter)
            .field("insert_hyperlinks", insert_hyperlinks)
            .field("pivot_tables", pivot_tables)
            .field("select_locked_cells", select_locked_cells)
            .field("select_unlocked_cells", select_unlocked_cells)
            .field("algorithm_name", algorithm_name)
            .field("hash_value", hash_value)
            .field("salt_value", salt_value)
            .field("spin_count", spin_count)
            .field(
                "plain_password",
                &plain_password.as_ref().map(|_| "<redacted>"),
            )
            .finish()
    }
}

/// Whether `value` is a legacy password hash as `ST_UnsignedShortHex` allows
/// it: 1 to 4 hex digits, either case.
fn is_legacy_password_hash(value: &str) -> bool {
    (1..=4).contains(&value.len()) && value.bytes().all(|b| b.is_ascii_hexdigit())
}

/// Why a `password` that is not a legacy hash cannot be written, naming the
/// sheet when it is known. The value is never quoted: it is likely a password.
pub(crate) fn non_hash_password_message(sheet: Option<&str>) -> String {
    let whose = sheet.map_or_else(
        || "the sheetProtection password".to_owned(),
        |name| format!("the sheetProtection password of sheet '{name}'"),
    );
    format!(
        "{whose} is not a legacy password hash (1 to 4 hex digits): `password` holds the hash a file stores, not the password — set plainPassword to protect the sheet with a password"
    )
}

impl SheetProtection {
    /// Whether the writer would emit `password` verbatim although it is not a
    /// legacy hash (1 to 4 hex digits). A set `plain_password` replaces it, so
    /// then it is never written.
    pub(crate) fn has_non_hash_password(&self) -> bool {
        self.plain_password.is_none()
            && self
                .password
                .as_deref()
                .is_some_and(|p| !is_legacy_password_hash(p))
    }

    /// The boolean attributes in `CT_SheetProtection` schema order, each with
    /// its current value and its schema default.
    pub(crate) fn flags(&self) -> [(&'static str, bool, bool); 16] {
        [
            ("sheet", self.sheet, false),
            ("objects", self.objects, false),
            ("scenarios", self.scenarios, false),
            ("formatCells", self.format_cells, true),
            ("formatColumns", self.format_columns, true),
            ("formatRows", self.format_rows, true),
            ("insertColumns", self.insert_columns, true),
            ("insertRows", self.insert_rows, true),
            ("insertHyperlinks", self.insert_hyperlinks, true),
            ("deleteColumns", self.delete_columns, true),
            ("deleteRows", self.delete_rows, true),
            ("selectLockedCells", self.select_locked_cells, false),
            ("sort", self.sort, true),
            ("autoFilter", self.auto_filter, true),
            ("pivotTables", self.pivot_tables, true),
            ("selectUnlockedCells", self.select_unlocked_cells, false),
        ]
    }

    /// A copy whose password attributes are a fresh hash of `plain_password`
    /// (see that field), or `None` when it is unset.
    pub(crate) fn with_hashed_password(&self) -> crate::Result<Option<Self>> {
        let Some(plain) = self.plain_password.as_deref() else {
            return Ok(None);
        };
        let mut hashed = Self {
            password: None,
            algorithm_name: None,
            hash_value: None,
            salt_value: None,
            spin_count: None,
            plain_password: None,
            ..self.clone()
        };
        #[cfg(feature = "encryption")]
        {
            let salt = crate::ole2::crypto::secure_random(PASSWORD_SALT_LEN)?;
            hashed.hash_value = Some(hash_password_sha512(plain, &salt, PASSWORD_SPIN_COUNT));
            hashed.salt_value = Some(crate::ole2::base64::encode_base64(&salt));
            hashed.algorithm_name = Some("SHA-512".to_owned());
            hashed.spin_count = Some(PASSWORD_SPIN_COUNT);
        }
        #[cfg(not(feature = "encryption"))]
        {
            hashed.password = Some(hash_password_legacy(plain)?);
        }
        Ok(Some(hashed))
    }
}

/// Salt length of a sheet password hash, as Excel writes it.
#[cfg(feature = "encryption")]
const PASSWORD_SALT_LEN: usize = 16;
/// Iterations of a sheet password hash, as Excel writes it.
#[cfg(feature = "encryption")]
const PASSWORD_SPIN_COUNT: u32 = 100_000;

/// The base64 SHA-512 sheet password hash of ECMA-376 Part 1 §18.2.29:
/// H0 = SHA-512(salt || UTF-16LE(password)), then
/// H_i = SHA-512(H_{i-1} || LE32(i - 1)) for `i` in `1..=spin_count`.
/// The iterator follows the previous hash, unlike the MS-OFFCRYPTO Agile KDF.
#[cfg(feature = "encryption")]
pub fn hash_password_sha512(password: &str, salt: &[u8], spin_count: u32) -> String {
    use sha2::digest::Output;
    use sha2::{Digest, Sha512};
    use zeroize::Zeroizing;

    let mut pw_bytes = Zeroizing::new(Vec::with_capacity(password.len() * 2));
    pw_bytes.extend(password.encode_utf16().flat_map(u16::to_le_bytes));

    let mut hasher = Sha512::new();
    let mut h = Output::<Sha512>::default();
    Digest::update(&mut hasher, salt);
    Digest::update(&mut hasher, pw_bytes.as_slice());
    Digest::finalize_into_reset(&mut hasher, &mut h);
    for i in 0..spin_count {
        Digest::update(&mut hasher, h.as_slice());
        Digest::update(&mut hasher, i.to_le_bytes());
        Digest::finalize_into_reset(&mut hasher, &mut h);
    }
    crate::ole2::base64::encode_base64(&h)
}

/// The legacy 16-bit sheet password hash as 4 uppercase hex digits
/// (`ST_UnsignedShortHex`, the `password` attribute).
///
/// # Errors
/// `InvalidWorkbook` for a password with a non-ASCII character: Excel hashes
/// the password's code-page bytes, which only ASCII pins down.
pub fn hash_password_legacy(password: &str) -> crate::Result<String> {
    if !password.is_ascii() {
        return Err(crate::ModernXlsxError::InvalidWorkbook(
            "sheet protection: the legacy password hash takes ASCII passwords only".to_owned(),
        ));
    }
    let mut hash: u16 = 0;
    for &b in password.as_bytes().iter().rev() {
        hash = ((hash >> 14) & 1) | ((hash << 1) & 0x7FFF);
        hash ^= u16::from(b);
    }
    hash = ((hash >> 14) & 1) | ((hash << 1) & 0x7FFF);
    hash ^= password.len() as u16;
    hash ^= 0xCE4B;
    Ok(format!("{hash:04X}"))
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
    /// One entry per `<color>` stop: its `rgb`, or "" when the stop has none.
    pub colors: Vec<String>,
    /// The full `<color>` of every stop, in stop order, filled only when some
    /// stop has a theme, indexed, tint or auto attribute. On write stop `i`
    /// uses `color_specs[i]` while its `rgb` equals `colors[i]` (an empty
    /// string counts as absent); otherwise `colors[i]` as plain `rgb`. A stop
    /// with neither is written with the data-bar default `FF638EC6`, so the
    /// stop count is kept.
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub color_specs: Vec<Color>,
}

/// A data bar conditional formatting rule.
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct DataBar {
    pub cfvos: Vec<Cfvo>,
    /// The `rgb` of the bar `<color>`, or "" when it has none.
    pub color: String,
    /// The full bar `<color>` when it has a theme, indexed, tint or auto
    /// attribute; same write rule as [`ColorScale::color_specs`]. A bar with
    /// neither is written with Excel's default bar colour `FF638EC6`.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub color_spec: Option<Color>,
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

/// Which `<headerFooter>` child element is being read.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) enum HfPart {
    /// `<oddHeader>`
    OddHeader,
    /// `<oddFooter>`
    OddFooter,
    /// `<evenHeader>`
    EvenHeader,
    /// `<evenFooter>`
    EvenFooter,
    /// `<firstHeader>`
    FirstHeader,
    /// `<firstFooter>`
    FirstFooter,
}

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
    InInlineRph, // <rPh> phonetic run inside <is>: its text is not cell text
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
    HeaderFooterChild(HfPart),
    InRowBreaks,
    InColBreaks,
    ExtLst,
    ExtSparklines,
    SparklineGroups,
    SparklineGroup,
    SparklineDateRange, // the group's own <xm:f>, a date axis's range
    Sparklines,
    SparklineItem,
    SparklineFormula,
    SparklineSqref,
    ExtOther(usize), // element depth inside a non-sparkline <ext> being captured verbatim
}

#[cfg(test)]
mod tests;
