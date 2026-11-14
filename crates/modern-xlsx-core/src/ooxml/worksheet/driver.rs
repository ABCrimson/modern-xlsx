//! The worksheet SAX driver: one quick-xml state machine over a worksheet
//! part. [`drive`] handles every element outside `<sheetData>` itself,
//! collecting it into a [`SheetMeta`], and hands the rows and cells to a
//! [`RowSink`]: `begin_row` at each `<row>`, `begin_cell` at each `<c>`,
//! `cell_part` as each child of the cell ends, `end_cell` at `</c>` (a
//! self-closing `<c/>` calls `begin_cell` and `end_cell`), `end_row` at
//! `</row>` (a self-closing `<row/>` calls both), and `abort_row` when the
//! input ends inside a `<row>`. `sink_struct::StructSink` builds the typed
//! rows for `WorksheetXml::parse_with_sst`; `sink_json::JsonSink` writes them
//! as JSON for `WorksheetXml::parse_to_json_within`.
//!
//! What the driver holds for one cell is its attributes and the text of the
//! child being read, whatever the number of children. Each element it keeps
//! in a list of the sheet (a data validation, a conditional-format rule, a
//! column, a hyperlink, ...) is counted, before it is kept, at the size of its
//! struct plus the heap bytes of its strings in the read's
//! [`ExpansionBudget::store`], so a part of millions of small elements ends
//! in `SizeLimit` instead of holding a hundred times its XML.

use core::hint::cold_path;

use quick_xml::Reader;
use quick_xml::events::{BytesStart, Event};

use super::parser::{
    CellAttrs, FormulaState, RowAttrs, apply_sparkline_color, assign_hf_field, captured_ext,
    find_attr, find_attr_str, is_sparkline_ext, merge_page_setup_attrs, parse_brk_attrs,
    parse_cell_attrs_into, parse_cf_rule_attrs, parse_cfvo_attrs, parse_col_element,
    parse_data_validation_attrs, parse_formula_attrs, parse_header_footer_attrs,
    parse_hyperlink_attrs, parse_outline_pr_attrs, parse_page_margins_attrs, parse_pane_attrs,
    parse_row_attrs, parse_selection_attrs, parse_sheet_protection_attrs, parse_sheet_view_attrs,
    parse_sparkline_group_attrs, stop_color_specs, take_exact, text_sink,
};
use super::{
    AutoFilter, Cfvo, ColorScale, ColumnInfo, ConditionalFormatting, ConditionalFormattingRule,
    CustomFilter, CustomFilters, DataBar, DataValidation, FilterColumn, FrozenPane, HeaderFooter,
    HfPart, Hyperlink, IconSet, OutlineProperties, PageBreak, PageBreaks, PageSetup, PaneSelection,
    ParseState, SheetContext, SheetProtection, SheetViewData, Sparkline, SparklineGroup, SplitPane,
};
use crate::ooxml::cell::{MAX_COL, col_to_letters_buf, letters_to_col};
use crate::ooxml::color::{Color, split_color};
use crate::ooxml::shared_strings::ExpansionBudget;
use crate::ooxml::xml::{
    REL_NS, RootNamespaces, inherited_declarations, parse_attr, push_general_ref, root_namespaces,
    with_declarations,
};
use crate::{ModernXlsxError, Result};

/// Where the rows and cells of `<sheetData>` go. Static dispatch only:
/// [`drive`] is generic over the sink.
pub(super) trait RowSink {
    /// A `<row>` starts, with its attributes.
    fn begin_row(&mut self, attrs: &RowAttrs);
    /// A cell of the current row starts, with its `<c>` attributes: the
    /// `cell_type` of its `t` attribute (`t="d"` reads as a number, with
    /// `iso_date` set) and its `cm` index as `cell_metadata`. The driver
    /// passes the same attributes to each [`Self::cell_part`] and to the
    /// [`Self::end_cell`] of the cell.
    fn begin_cell(&mut self, attrs: &CellAttrs);
    /// A child of the open cell ended. The children come one by one, in
    /// document order, so the sink decides the result of every child order
    /// (a `<v>` before an `<f>`, a repeated `<v>`, an `<is>` beside a `<v>`).
    /// The text the sink copies into the cell from a table (an SST entry) is
    /// counted in `budget` first.
    ///
    /// # Errors
    ///
    /// Whatever the sink's own conversion of the child returns, `SizeLimit`
    /// among them; the driver stops at the first error.
    fn cell_part(
        &mut self,
        attrs: &CellAttrs,
        part: CellPart<'_>,
        budget: &mut ExpansionBudget,
    ) -> Result<()>;
    /// The open cell ends.
    ///
    /// # Errors
    ///
    /// As [`Self::cell_part`].
    fn end_cell(&mut self, attrs: &CellAttrs, budget: &mut ExpansionBudget) -> Result<()>;
    /// The current row ends.
    ///
    /// # Errors
    ///
    /// Whatever the sink's own limit returns; the driver stops at it.
    fn end_row(&mut self) -> Result<()>;
    /// The part ended (end of input, no error) inside a `<row>`, possibly
    /// inside one of its cells: drop that row, keeping only the rows that
    /// ended. Called at most once, last, and never after `end_row` for the
    /// same row.
    fn abort_row(&mut self);
}

/// One child of a `<c>`, handed to [`RowSink::cell_part`] as it ends. Its
/// text lives in a buffer that the driver reuses for every child.
pub(super) enum CellPart<'a> {
    /// `<v>`: its text. A self-closing `<v/>` is no part.
    Value(&'a str),
    /// `<f>` or a self-closing `<f/>` (`text` is `None`): the attributes that
    /// one element set, parsed by `parse_formula_attrs` from defaults (the
    /// sink may take them), and its text.
    Formula {
        attrs: &'a mut FormulaState,
        text: Option<&'a str>,
    },
    /// `<is>` holding at least one `<t>`: the concatenated text of its `<t>`
    /// runs (also those inside `<r>`, never `<rPh>`). An `<is>` without a
    /// `<t>` is no part.
    InlineStr(&'a str),
}

/// Everything in a worksheet part outside `<sheetData>`, as [`drive`]
/// collected it. The fields are the `WorksheetXml` fields of the same name.
pub(super) struct SheetMeta {
    pub(super) dimension: Option<String>,
    pub(super) merge_cells: Vec<String>,
    pub(super) auto_filter: Option<AutoFilter>,
    pub(super) frozen_pane: Option<FrozenPane>,
    pub(super) split_pane: Option<SplitPane>,
    pub(super) pane_selections: Vec<PaneSelection>,
    pub(super) sheet_view: Option<SheetViewData>,
    pub(super) columns: Vec<ColumnInfo>,
    pub(super) data_validations: Vec<DataValidation>,
    pub(super) conditional_formatting: Vec<ConditionalFormatting>,
    pub(super) hyperlinks: Vec<Hyperlink>,
    pub(super) page_setup: Option<PageSetup>,
    pub(super) sheet_protection: Option<SheetProtection>,
    pub(super) tab_color: Option<String>,
    pub(super) tab_color_spec: Option<Color>,
    pub(super) header_footer: Option<HeaderFooter>,
    pub(super) page_breaks: Option<PageBreaks>,
    pub(super) outline_properties: Option<OutlineProperties>,
    pub(super) sparkline_groups: Vec<SparklineGroup>,
    pub(super) preserved_extensions: Vec<String>,
}

/// Run the worksheet state machine over `data`, sending rows and cells to
/// `row_sink` and returning the rest of the part. `ctx.rels` resolves an external
/// hyperlink's `r:id`. A `<row>` or `<c>` without `r` takes the next
/// position (`place_row`, `place_cell`), so every row reaches the sink with
/// its index and every cell with its reference.
///
/// `budget` counts the text copied from a table: hyperlink targets and the
/// namespace declarations of the root and of the `<extLst>` re-declared on
/// each kept `<ext>` here (`kept_ext`), SST entries in the sink. A row dropped
/// by `abort_row` never reaches the result, so what its cells counted is given
/// back. `budget` also stores the memory of every element kept in a list of
/// [`SheetMeta`] (and of the namespace declarations of the root and of each
/// `<extLst>`, with their indexes), before it is kept; nested lists (a rule's colour
/// scale values, a filter column's values) are moved into their element at
/// their exact length.
///
/// # Errors
///
/// `XmlParse` for ill-formed XML, a malformed attribute or a cell without `r`
/// past column XFD; `SizeLimit` when the copies or the stored elements pass
/// the budget; and any error the sink returns.
pub(super) fn drive<S: RowSink>(
    data: &[u8],
    ctx: &SheetContext<'_>,
    row_sink: &mut S,
    budget: &mut ExpansionBudget,
) -> Result<SheetMeta> {
    let mut reader = Reader::from_reader(data);
    // Text is never trimmed: quick-xml splits text at every entity, so
    // trimming would eat the spaces around `&amp;` and xml:space="preserve"
    // edges. Whitespace between elements only reaches structural states,
    // whose Text arms ignore it.
    reader.config_mut().trim_text(false);

    // Set at the document element's `<worksheet>` start tag: from then on, a
    // child of the root that no arm reads (a nested `<worksheet>` among them)
    // is skipped with its whole subtree.
    let mut in_worksheet = false;

    let mut dimension: Option<String> = None;
    let mut merge_cells: Vec<String> = Vec::new();
    let mut auto_filter: Option<AutoFilter> = None;
    let mut frozen_pane: Option<FrozenPane> = None;
    let mut split_pane = None::<SplitPane>;
    let mut pane_selections = Vec::<PaneSelection>::new();
    let mut sheet_view = None::<SheetViewData>;
    // Only the first <sheetView> counts (see CT_SheetViews).
    let mut sheet_views_seen: u32 = 0;
    let mut in_first_view = false;
    let mut columns: Vec<ColumnInfo> = Vec::new();
    let mut data_validations: Vec<DataValidation> = Vec::new();
    let mut conditional_formatting: Vec<ConditionalFormatting> = Vec::new();
    let mut hyperlinks: Vec<Hyperlink> = Vec::new();
    let mut page_setup: Option<PageSetup> = None;
    let mut sheet_protection: Option<SheetProtection> = None;
    let mut tab_color: Option<String> = None;
    let mut tab_color_spec: Option<Color> = None;
    let mut header_footer: Option<HeaderFooter> = None;
    let mut hf_text_buf = String::new();
    let mut page_breaks: Option<PageBreaks> = None;
    let mut outline_properties: Option<OutlineProperties> = None;

    // Sparkline/extLst parsing state. A non-sparkline <ext> is kept as the
    // byte span from its start tag (`ext_start`) to its end tag, with the
    // namespace declarations it uses from the open `<extLst>` (`ext_lst_ns`,
    // when that declares any; replaced at the next `<extLst>` start tag and
    // read only inside the element) and the root (`root_ns`) re-declared on
    // it. The group `<xm:f>` (a date axis's range) collects into
    // `current_sparkline_dates`.
    let mut sparkline_groups: Vec<SparklineGroup> = Vec::new();
    let mut preserved_extensions: Vec<String> = Vec::new();
    let mut current_sparkline_group: Option<SparklineGroup> = None;
    let mut current_sparklines: Vec<Sparkline> = Vec::new();
    let mut current_sparkline_formula = String::new();
    let mut current_sparkline_sqref = String::new();
    let mut current_sparkline_dates = String::new();
    let mut root_ns = RootNamespaces::default();
    let mut ext_lst_ns: Option<RootNamespaces> = None;
    // The prefixes `root_ns` binds to the relationships namespace, in the
    // sorted order of its index and each once, so each hyperlink's `r:id`
    // test is one binary search however many prefixes the root declares,
    // with no hasher for the prefixes a file chooses to make collide.
    let mut rel_prefixes: Vec<String> = Vec::new();
    let mut ext_start: u64 = 0;

    let mut state = ParseState::Root;
    // `budget.spent` when the open row began, restored if the row is dropped.
    let mut row_spent = budget.spent;

    // The attributes of the cell being read, reused by every `<c>`:
    // `parse_cell_attrs_into` resets `ca`.
    let mut ca = CellAttrs::default();
    // The position an `r`-less `<row>` or `<c>` takes: the index of the
    // current (last begun) row, and the 0-based column after the row's
    // previous cell. See `place_row` and `place_cell`.
    let mut prev_row: u32 = 0;
    let mut next_col: u32 = 0;
    // The attributes of the open `<f>` (or of a self-closing `<f/>`), until
    // its end hands them to the sink.
    let mut cur_f = FormulaState::default();

    // Text content. The text of a cell's open `<v>`, `<f>` or `<is>`
    // accumulates here from a cleared buffer; under `<is>` only the `<t>`
    // runs collect text, and `inline_has_t` records that a `<t>` was seen.
    // Data-validation and conditional-formatting formulas use it from a
    // cleared buffer too.
    let mut text_buf = String::with_capacity(256);
    let mut inline_has_t = false;
    // Where the open `<sparklineGroup>` starts: its stored size counts the
    // bytes of its span, which hold every string it keeps.
    let mut group_start: u64 = 0;

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
    let mut cur_custom_filters: Option<CustomFilters> = None;
    let mut cur_custom_filter_and: bool = false;
    let mut cur_custom_filter_items: Vec<CustomFilter> = Vec::new();

    // Conditional formatting sub-element state.
    let mut cur_cfvos: Vec<Cfvo> = Vec::new();
    let mut cur_cf_colors: Vec<String> = Vec::new();
    let mut cur_cf_color_specs: Vec<Color> = Vec::new();
    let mut cur_cf_bar_color = String::new();
    let mut cur_cf_bar_spec: Option<Color> = None;
    let mut cur_icon_set_type: Option<String> = None;

    'events: loop {
        // Where the next event starts: the opening `<` of an <ext> span.
        let ev_start = reader.buffer_position();
        let event = reader.read_event();
        // Leaf elements, which have no children this parser reads, are read
        // the same whether self-closing or written as a start/end pair (the
        // XML infoset does not tell the two apart); the end tag of the pair
        // matches no End arm. Handled before the Start and Empty arms, and so
        // before the Start arm's skip of unread root children.
        if let Ok(Event::Start(ref e) | Event::Empty(ref e)) = event {
            let local = e.local_name();
            'leaf: {
                match (state, local.as_ref()) {
                    (ParseState::Root, "dimension") => {
                        dimension = find_attr_str(e, "ref")?;
                    }
                    (ParseState::Root, "pageSetup") => {
                        merge_page_setup_attrs(e, page_setup.get_or_insert_default())?;
                    }
                    (ParseState::Root, "pageMargins") => {
                        page_setup.get_or_insert_default().margins =
                            Some(parse_page_margins_attrs(e)?);
                    }
                    (ParseState::Root, "sheetProtection") => {
                        sheet_protection = Some(parse_sheet_protection_attrs(e)?);
                    }
                    (ParseState::SheetPr, "tabColor") => {
                        (tab_color, tab_color_spec) = split_color(Color::from_attrs(e)?);
                    }
                    (ParseState::SheetPr, "outlinePr") => {
                        outline_properties = Some(parse_outline_pr_attrs(e)?);
                    }
                    (ParseState::SheetPr, "pageSetUpPr") => {
                        if let Some(b) = find_attr(e, "fitToPage")?.and_then(|a| a.as_bool()) {
                            page_setup.get_or_insert_default().fit_to_page = Some(b);
                        }
                    }
                    (ParseState::SheetView, "pane") if in_first_view => {
                        let (fp, sp) = parse_pane_attrs(e)?;
                        if fp.is_some() {
                            frozen_pane = fp;
                        }
                        if sp.is_some() {
                            split_pane = sp;
                        }
                    }
                    (ParseState::Cols, "col") => {
                        budget.store(size_of::<ColumnInfo>())?;
                        columns.push(parse_col_element(e)?);
                    }
                    (ParseState::MergeCells, "mergeCell") => {
                        if let Some(val) = find_attr_str(e, "ref")? {
                            budget.store(size_of::<String>() + val.capacity())?;
                            merge_cells.push(val);
                        }
                    }
                    (ParseState::InHyperlinks, "hyperlink") => {
                        let link = parse_hyperlink_attrs(e, ctx, &rel_prefixes, budget)?;
                        budget.store(hyperlink_bytes(&link))?;
                        hyperlinks.push(link);
                    }
                    // A page break.
                    (ParseState::InRowBreaks, "brk") => {
                        if let Some(ref mut pb) = page_breaks {
                            budget.store(size_of::<PageBreak>())?;
                            pb.row_breaks.push(parse_brk_attrs(e)?);
                        }
                    }
                    (ParseState::InColBreaks, "brk") => {
                        if let Some(ref mut pb) = page_breaks {
                            budget.store(size_of::<PageBreak>())?;
                            pb.col_breaks.push(parse_brk_attrs(e)?);
                        }
                    }
                    // cfvo in colorScale/dataBar/iconSet.
                    (
                        ParseState::InColorScale | ParseState::InDataBar | ParseState::InIconSet,
                        "cfvo",
                    ) => {
                        let cfvo = parse_cfvo_attrs(e)?;
                        budget.store(
                            size_of::<Cfvo>()
                                + cfvo.cfvo_type.capacity()
                                + cfvo.val.as_ref().map_or(0, String::capacity),
                        )?;
                        cur_cfvos.push(cfvo);
                    }
                    // color in colorScale: one entry per stop, rgb or not.
                    (ParseState::InColorScale, "color") => {
                        let c = Color::from_attrs(e)?.unwrap_or_default();
                        let rgb = c.rgb.clone().unwrap_or_default();
                        budget.store(
                            size_of::<String>()
                                + rgb.capacity()
                                + size_of::<Color>()
                                + c.rgb.as_ref().map_or(0, String::capacity),
                        )?;
                        cur_cf_colors.push(rgb);
                        cur_cf_color_specs.push(c);
                    }
                    // color in dataBar.
                    (ParseState::InDataBar, "color") => {
                        let (rgb, spec) = split_color(Color::from_attrs(e)?);
                        cur_cf_bar_color = rgb.unwrap_or_default();
                        cur_cf_bar_spec = spec;
                    }
                    // filter values.
                    (ParseState::InFilters, "filter") => {
                        if let Some(val) = find_attr_str(e, "val")? {
                            budget.store(size_of::<String>() + val.capacity())?;
                            cur_filter_vals.push(val);
                        }
                    }
                    // custom filter values.
                    (ParseState::InCustomFilters, "customFilter") => {
                        let val = find_attr_str(e, "val")?.unwrap_or_default();
                        let operator = find_attr_str(e, "operator")?;
                        budget.store(
                            size_of::<CustomFilter>()
                                + val.capacity()
                                + operator.as_ref().map_or(0, String::capacity),
                        )?;
                        cur_custom_filter_items.push(CustomFilter { operator, val });
                    }
                    // The eight sparkline group colour elements.
                    (
                        ParseState::SparklineGroup,
                        "colorSeries" | "colorNegative" | "colorAxis" | "colorMarkers"
                        | "colorFirst" | "colorLast" | "colorHigh" | "colorLow",
                    ) => {
                        if let Some(ref mut group) = current_sparkline_group {
                            apply_sparkline_color(group, local.as_ref(), e)?;
                        }
                    }
                    _ => break 'leaf,
                }
                continue 'events;
            }
        }
        match event {
            Ok(Event::Start(ref e)) => {
                let local = e.local_name();
                match (state, local.as_ref()) {
                    // Only the document element records the root bindings.
                    (ParseState::Root, "worksheet") if !in_worksheet => {
                        in_worksheet = true;
                        root_ns = RootNamespaces::new(root_namespaces(e)?);
                        budget.store(root_ns.stored_bytes())?;
                        rel_prefixes = root_ns
                            .prefixes_bound_to(REL_NS)
                            .map(str::to_owned)
                            .collect();
                        budget.store(
                            rel_prefixes.capacity() * size_of::<String>()
                                + rel_prefixes.iter().map(String::capacity).sum::<usize>(),
                        )?;
                    }
                    (ParseState::Root, "sheetPr") => {
                        state = ParseState::SheetPr;
                    }
                    (ParseState::Root, "sheetViews") => {
                        state = ParseState::SheetViews;
                    }
                    (ParseState::SheetViews, "sheetView") => {
                        state = ParseState::SheetView;
                        in_first_view = sheet_views_seen == 0;
                        sheet_views_seen += 1;
                        // Only the first view counts; later views are
                        // skipped but still tracked so End events balance.
                        if in_first_view {
                            let (sv, has_non_default) = parse_sheet_view_attrs(e)?;
                            if has_non_default {
                                sheet_view = Some(sv);
                            }
                        }
                    }
                    (ParseState::SheetView, "selection") if in_first_view => {
                        let sel = parse_selection_attrs(e)?;
                        budget.store(selection_bytes(&sel))?;
                        pane_selections.push(sel);
                    }
                    (ParseState::Root, "cols") => {
                        state = ParseState::Cols;
                    }
                    (ParseState::Root, "sheetData") => {
                        state = ParseState::SheetData;
                    }
                    (ParseState::SheetData, "row") => {
                        state = ParseState::InRow;
                        let mut ra = parse_row_attrs(e)?;
                        place_row(&mut ra, &mut prev_row, &mut next_col)?;
                        row_spent = budget.spent;
                        row_sink.begin_row(&ra);
                    }
                    (ParseState::InRow, "c") => {
                        state = ParseState::InCell;
                        parse_cell_attrs_into(e, &mut ca)?;
                        place_cell(&mut ca, prev_row, &mut next_col)?;
                        row_sink.begin_cell(&ca);
                    }
                    (ParseState::InCell, "v") => {
                        state = ParseState::InCellValue;
                        text_buf.clear();
                    }
                    (ParseState::InCell, "f") => {
                        state = ParseState::InCellFormula;
                        text_buf.clear();
                        cur_f = FormulaState::default();
                        parse_formula_attrs(e, &mut cur_f)?;
                    }
                    (ParseState::InCell, "is") => {
                        state = ParseState::InInlineStr;
                        text_buf.clear();
                        inline_has_t = false;
                    }
                    (ParseState::InInlineStr, "t") => {
                        state = ParseState::InInlineStrT;
                        inline_has_t = true;
                    }
                    (ParseState::InInlineStr, "rPh") => {
                        state = ParseState::InInlineRph;
                    }
                    (ParseState::Root, "mergeCells") => {
                        state = ParseState::MergeCells;
                    }
                    (ParseState::Root, "dataValidations") => {
                        state = ParseState::InDataValidations;
                    }
                    (ParseState::InDataValidations, "dataValidation") => {
                        state = ParseState::InDataValidation;
                        cur_dv = Some(parse_data_validation_attrs(e)?);
                    }
                    (ParseState::InDataValidation, "formula1") => {
                        state = ParseState::InDVFormula1;
                        text_buf.clear();
                    }
                    (ParseState::InDataValidation, "formula2") => {
                        state = ParseState::InDVFormula2;
                        text_buf.clear();
                    }
                    (ParseState::Root, "conditionalFormatting") => {
                        state = ParseState::InConditionalFormatting;
                        cur_cf_sqref.clear();
                        cur_cf_rules.clear();
                        if let Some(val) = find_attr_str(e, "sqref")? {
                            cur_cf_sqref = val;
                        }
                    }
                    (ParseState::InConditionalFormatting, "cfRule") => {
                        state = ParseState::InCfRule;
                        cur_cf_rule = Some(parse_cf_rule_attrs(e)?);
                    }
                    (ParseState::InCfRule, "formula") => {
                        state = ParseState::InCfRuleFormula;
                        text_buf.clear();
                    }
                    (ParseState::InCfRule, "colorScale") => {
                        state = ParseState::InColorScale;
                        cur_cfvos.clear();
                        cur_cf_colors.clear();
                        cur_cf_color_specs.clear();
                    }
                    (ParseState::InCfRule, "dataBar") => {
                        state = ParseState::InDataBar;
                        cur_cfvos.clear();
                        cur_cf_bar_color.clear();
                        cur_cf_bar_spec = None;
                    }
                    (ParseState::InCfRule, "iconSet") => {
                        state = ParseState::InIconSet;
                        cur_cfvos.clear();
                        cur_icon_set_type = find_attr_str(e, "iconSet")?;
                    }
                    (ParseState::Root, "hyperlinks") => {
                        state = ParseState::InHyperlinks;
                    }
                    (ParseState::Root, "autoFilter") => {
                        state = ParseState::InAutoFilter;
                        cur_af_range.clear();
                        cur_af_columns.clear();
                        if let Some(val) = find_attr_str(e, "ref")? {
                            cur_af_range = val;
                        }
                    }
                    (ParseState::InAutoFilter, "filterColumn") => {
                        state = ParseState::InFilterColumn;
                        cur_filter_col_id = find_attr(e, "colId")?
                            .map(|a| parse_attr("filterColumn", &a))
                            .transpose()?
                            .unwrap_or(0);
                        cur_filter_vals.clear();
                        cur_custom_filters = None;
                    }
                    (ParseState::InFilterColumn, "filters") => {
                        state = ParseState::InFilters;
                    }
                    (ParseState::InFilterColumn, "customFilters") => {
                        state = ParseState::InCustomFilters;
                        cur_custom_filter_and =
                            find_attr(e, "and")?.and_then(|a| a.as_bool()) == Some(true);
                        cur_custom_filter_items.clear();
                    }
                    (ParseState::Root, "headerFooter") => {
                        header_footer = Some(parse_header_footer_attrs(e)?);
                        state = ParseState::HeaderFooter;
                    }
                    (ParseState::HeaderFooter, "oddHeader") => {
                        hf_text_buf.clear();
                        state = ParseState::HeaderFooterChild(HfPart::OddHeader);
                    }
                    (ParseState::HeaderFooter, "oddFooter") => {
                        hf_text_buf.clear();
                        state = ParseState::HeaderFooterChild(HfPart::OddFooter);
                    }
                    (ParseState::HeaderFooter, "evenHeader") => {
                        hf_text_buf.clear();
                        state = ParseState::HeaderFooterChild(HfPart::EvenHeader);
                    }
                    (ParseState::HeaderFooter, "evenFooter") => {
                        hf_text_buf.clear();
                        state = ParseState::HeaderFooterChild(HfPart::EvenFooter);
                    }
                    (ParseState::HeaderFooter, "firstHeader") => {
                        hf_text_buf.clear();
                        state = ParseState::HeaderFooterChild(HfPart::FirstHeader);
                    }
                    (ParseState::HeaderFooter, "firstFooter") => {
                        hf_text_buf.clear();
                        state = ParseState::HeaderFooterChild(HfPart::FirstFooter);
                    }

                    // ---- rowBreaks / colBreaks ----
                    (ParseState::Root, "rowBreaks") => {
                        page_breaks.get_or_insert_default();
                        state = ParseState::InRowBreaks;
                    }
                    (ParseState::Root, "colBreaks") => {
                        page_breaks.get_or_insert_default();
                        state = ParseState::InColBreaks;
                    }

                    // ---- extLst / sparklines ----
                    (ParseState::Root, "extLst") => {
                        ext_lst_ns = ext_lst_namespaces(e, budget)?;
                        state = ParseState::ExtLst;
                    }
                    (ParseState::ExtLst, "ext") => {
                        if is_sparkline_ext(e)? {
                            state = ParseState::ExtSparklines;
                        } else {
                            ext_start = ev_start;
                            state = ParseState::ExtOther(1);
                        }
                    }
                    (ParseState::ExtSparklines, "sparklineGroups") => {
                        state = ParseState::SparklineGroups;
                    }
                    (ParseState::SparklineGroups, "sparklineGroup") => {
                        current_sparkline_group = Some(parse_sparkline_group_attrs(e)?);
                        group_start = ev_start;
                        state = ParseState::SparklineGroup;
                    }
                    (ParseState::SparklineGroup, "sparklines") => {
                        state = ParseState::Sparklines;
                    }
                    (ParseState::SparklineGroup, "f") => {
                        current_sparkline_dates.clear();
                        state = ParseState::SparklineDateRange;
                    }
                    (ParseState::Sparklines, "sparkline") => {
                        current_sparkline_formula.clear();
                        current_sparkline_sqref.clear();
                        state = ParseState::SparklineItem;
                    }
                    (ParseState::SparklineItem, "f") => {
                        state = ParseState::SparklineFormula;
                    }
                    (ParseState::SparklineItem, "sqref") => {
                        state = ParseState::SparklineSqref;
                    }
                    (ParseState::ExtOther(depth), _) => state = ParseState::ExtOther(depth + 1),
                    // A child of the root that no arm above reads, such as
                    // <customSheetViews> or <protectedRanges>: skip it whole,
                    // so that its children (a custom view's <autoFilter>,
                    // <rowBreaks>, <pageSetup>, <extLst>) never reach the
                    // Root arms and the sheet's own model. Input that ends
                    // inside the skipped child ends the read like Eof, keeping
                    // what was read before it. The buffered read_to_end_into
                    // is the copy the chart, pivot and styles parsers already
                    // compile; the slice-only read_to_end would be a second one.
                    (ParseState::Root, _) if in_worksheet => {
                        match reader.read_to_end_into(e.name(), &mut Vec::new()) {
                            Ok(_) => {}
                            Err(quick_xml::Error::IllFormed(
                                quick_xml::errors::IllFormedError::MissingEndTag(_),
                            )) => break 'events,
                            Err(err) => {
                                cold_path();
                                return Err(ModernXlsxError::XmlParse(format!(
                                    "error parsing worksheet XML: {err}"
                                )));
                            }
                        }
                    }
                    _ => {}
                }
            }
            Ok(Event::Empty(ref e)) => {
                let local = e.local_name();
                match (state, local.as_ref()) {
                    (ParseState::Root, "autoFilter") => {
                        auto_filter = Some(AutoFilter {
                            range: find_attr_str(e, "ref")?.unwrap_or_default(),
                            filter_columns: Vec::new(),
                        });
                    }
                    (ParseState::SheetView, "selection") if in_first_view => {
                        let sel = parse_selection_attrs(e)?;
                        budget.store(selection_bytes(&sel))?;
                        pane_selections.push(sel);
                    }
                    (ParseState::SheetViews, "sheetView") if sheet_views_seen == 0 => {
                        sheet_views_seen += 1;
                        let (sv, has_non_default) = parse_sheet_view_attrs(e)?;
                        if has_non_default {
                            sheet_view = Some(sv);
                        }
                    }
                    // A self-closing non-sparkline <ext/> is kept as is.
                    (ParseState::ExtLst, "ext") => {
                        if !is_sparkline_ext(e)? {
                            preserved_extensions.push(kept_ext(
                                data,
                                ev_start,
                                reader.buffer_position(),
                                ext_lst_ns.as_ref(),
                                &root_ns,
                                budget,
                            )?);
                        }
                    }
                    (ParseState::InRow, "c") => {
                        // Self-closing <c ... /> — cell with no child elements.
                        parse_cell_attrs_into(e, &mut ca)?;
                        place_cell(&mut ca, prev_row, &mut next_col)?;
                        row_sink.begin_cell(&ca);
                        row_sink.end_cell(&ca, budget)?;
                    }
                    (ParseState::SheetData, "row") => {
                        // Self-closing <row ... /> — empty row.
                        let mut ra = parse_row_attrs(e)?;
                        place_row(&mut ra, &mut prev_row, &mut next_col)?;
                        row_sink.begin_row(&ra);
                        row_sink.end_row()?;
                    }
                    (ParseState::InCell, "v") => {
                        // Empty <v/> — no value.
                    }
                    (ParseState::InCell, "f") => {
                        // Empty <f/> — parse attributes even on self-closing.
                        cur_f = FormulaState::default();
                        parse_formula_attrs(e, &mut cur_f)?;
                        row_sink.cell_part(
                            &ca,
                            CellPart::Formula {
                                attrs: &mut cur_f,
                                text: None,
                            },
                            budget,
                        )?;
                    }
                    (ParseState::InDataValidations, "dataValidation") => {
                        // Self-closing <dataValidation ... /> with no formula children.
                        let dv = parse_data_validation_attrs(e)?;
                        budget.store(data_validation_bytes(&dv))?;
                        data_validations.push(dv);
                    }
                    (ParseState::InConditionalFormatting, "cfRule") => {
                        // Self-closing <cfRule ... /> with no formula child.
                        let rule = parse_cf_rule_attrs(e)?;
                        budget.store(cf_rule_bytes(&rule))?;
                        cur_cf_rules.push(rule);
                    }
                    // An empty <formula/> keeps its position, so the operands
                    // after it do not shift.
                    (ParseState::InCfRule, "formula") => {
                        if let Some(ref mut rule) = cur_cf_rule {
                            rule.push_formula(String::new());
                        }
                    }
                    _ => {}
                }
            }
            Ok(Event::Text(ref e)) => {
                if let Some(sink) = text_sink(
                    state,
                    &mut text_buf,
                    &mut hf_text_buf,
                    &mut current_sparkline_formula,
                    &mut current_sparkline_sqref,
                    &mut current_sparkline_dates,
                ) {
                    sink.push_str(&e.xml10_content());
                }
            }
            Ok(Event::CData(ref e)) => {
                if let Some(sink) = text_sink(
                    state,
                    &mut text_buf,
                    &mut hf_text_buf,
                    &mut current_sparkline_formula,
                    &mut current_sparkline_sqref,
                    &mut current_sparkline_dates,
                ) {
                    sink.push_str(&e.xml10_content());
                }
            }
            Ok(Event::GeneralRef(ref e)) => {
                if let Some(sink) = text_sink(
                    state,
                    &mut text_buf,
                    &mut hf_text_buf,
                    &mut current_sparkline_formula,
                    &mut current_sparkline_sqref,
                    &mut current_sparkline_dates,
                ) {
                    push_general_ref(sink, e)?;
                }
            }
            Ok(Event::End(ref e)) => {
                let local = e.local_name();
                match (state, local.as_ref()) {
                    (ParseState::InCellValue, "v") => {
                        row_sink.cell_part(&ca, CellPart::Value(&text_buf), budget)?;
                        state = ParseState::InCell;
                    }
                    (ParseState::InCellFormula, "f") => {
                        row_sink.cell_part(
                            &ca,
                            CellPart::Formula {
                                attrs: &mut cur_f,
                                text: Some(&text_buf),
                            },
                            budget,
                        )?;
                        state = ParseState::InCell;
                    }
                    (ParseState::InInlineStrT, "t") => {
                        state = ParseState::InInlineStr;
                    }
                    (ParseState::InInlineRph, "rPh") => {
                        state = ParseState::InInlineStr;
                    }
                    (ParseState::InInlineStr, "is") => {
                        if inline_has_t {
                            row_sink.cell_part(&ca, CellPart::InlineStr(&text_buf), budget)?;
                        }
                        state = ParseState::InCell;
                    }
                    (ParseState::InCell, "c") => {
                        row_sink.end_cell(&ca, budget)?;
                        state = ParseState::InRow;
                    }
                    (ParseState::InRow, "row") => {
                        row_sink.end_row()?;
                        state = ParseState::SheetData;
                    }
                    (ParseState::SheetData, "sheetData") => {
                        state = ParseState::Root;
                    }
                    (ParseState::SheetPr, "sheetPr") => {
                        state = ParseState::Root;
                    }
                    (ParseState::SheetView, "sheetView") => {
                        state = ParseState::SheetViews;
                    }
                    (ParseState::SheetViews, "sheetViews") => {
                        state = ParseState::Root;
                    }
                    (ParseState::Cols, "cols") => {
                        state = ParseState::Root;
                    }
                    (ParseState::MergeCells, "mergeCells") => {
                        state = ParseState::Root;
                    }
                    (ParseState::InDVFormula1, "formula1") => {
                        if let Some(ref mut dv) = cur_dv {
                            dv.formula1 = Some(std::mem::take(&mut text_buf));
                        }
                        state = ParseState::InDataValidation;
                    }
                    (ParseState::InDVFormula2, "formula2") => {
                        if let Some(ref mut dv) = cur_dv {
                            dv.formula2 = Some(std::mem::take(&mut text_buf));
                        }
                        state = ParseState::InDataValidation;
                    }
                    (ParseState::InDataValidation, "dataValidation") => {
                        if let Some(dv) = cur_dv.take() {
                            budget.store(data_validation_bytes(&dv))?;
                            data_validations.push(dv);
                        }
                        state = ParseState::InDataValidations;
                    }
                    (ParseState::InDataValidations, "dataValidations") => {
                        state = ParseState::Root;
                    }
                    (ParseState::InCfRuleFormula, "formula") => {
                        if let Some(ref mut rule) = cur_cf_rule {
                            rule.push_formula(std::mem::take(&mut text_buf));
                        }
                        state = ParseState::InCfRule;
                    }
                    (ParseState::InCfRule, "cfRule") => {
                        if let Some(rule) = cur_cf_rule.take() {
                            budget.store(cf_rule_bytes(&rule))?;
                            cur_cf_rules.push(rule);
                        }
                        state = ParseState::InConditionalFormatting;
                    }
                    (ParseState::InConditionalFormatting, "conditionalFormatting") => {
                        let sqref = std::mem::take(&mut cur_cf_sqref);
                        budget.store(size_of::<ConditionalFormatting>() + sqref.capacity())?;
                        conditional_formatting.push(ConditionalFormatting {
                            sqref,
                            rules: take_exact(&mut cur_cf_rules),
                        });
                        state = ParseState::Root;
                    }
                    (ParseState::InHyperlinks, "hyperlinks") => {
                        state = ParseState::Root;
                    }
                    (ParseState::InAutoFilter, "autoFilter") => {
                        auto_filter = Some(AutoFilter {
                            range: std::mem::take(&mut cur_af_range),
                            filter_columns: take_exact(&mut cur_af_columns),
                        });
                        state = ParseState::Root;
                    }
                    (ParseState::InFilters, "filters") => {
                        state = ParseState::InFilterColumn;
                    }
                    (ParseState::InCustomFilters, "customFilters") => {
                        cur_custom_filters = Some(CustomFilters {
                            and_op: cur_custom_filter_and,
                            filters: take_exact(&mut cur_custom_filter_items),
                        });
                        state = ParseState::InFilterColumn;
                    }
                    (ParseState::InFilterColumn, "filterColumn") => {
                        budget.store(size_of::<FilterColumn>())?;
                        cur_af_columns.push(FilterColumn {
                            col_id: cur_filter_col_id,
                            filters: take_exact(&mut cur_filter_vals),
                            custom_filters: cur_custom_filters.take(),
                        });
                        state = ParseState::InAutoFilter;
                    }
                    (ParseState::InColorScale, "colorScale") => {
                        if let Some(ref mut rule) = cur_cf_rule {
                            rule.color_scale = Some(ColorScale {
                                cfvos: take_exact(&mut cur_cfvos),
                                colors: take_exact(&mut cur_cf_colors),
                                color_specs: stop_color_specs(&mut cur_cf_color_specs),
                            });
                        }
                        state = ParseState::InCfRule;
                    }
                    (ParseState::InDataBar, "dataBar") => {
                        if let Some(ref mut rule) = cur_cf_rule {
                            rule.data_bar = Some(DataBar {
                                cfvos: take_exact(&mut cur_cfvos),
                                color: std::mem::take(&mut cur_cf_bar_color),
                                color_spec: cur_cf_bar_spec.take(),
                            });
                        }
                        state = ParseState::InCfRule;
                    }
                    (ParseState::InIconSet, "iconSet") => {
                        if let Some(ref mut rule) = cur_cf_rule {
                            rule.icon_set = Some(IconSet {
                                icon_set_type: cur_icon_set_type.take(),
                                cfvos: take_exact(&mut cur_cfvos),
                            });
                        }
                        state = ParseState::InCfRule;
                    }
                    (ParseState::HeaderFooterChild(part), _) => {
                        if let Some(ref mut hf) = header_footer {
                            let text = if hf_text_buf.is_empty() {
                                None
                            } else {
                                Some(std::mem::take(&mut hf_text_buf))
                            };
                            assign_hf_field(hf, part, text);
                        }
                        state = ParseState::HeaderFooter;
                    }
                    (ParseState::HeaderFooter, "headerFooter") => {
                        state = ParseState::Root;
                    }
                    (ParseState::InRowBreaks, "rowBreaks") => {
                        state = ParseState::Root;
                    }
                    (ParseState::InColBreaks, "colBreaks") => {
                        state = ParseState::Root;
                    }

                    // ---- extLst / sparklines end ----
                    (ParseState::SparklineFormula, "f") => {
                        state = ParseState::SparklineItem;
                    }
                    (ParseState::SparklineDateRange, "f") => {
                        if let Some(ref mut group) = current_sparkline_group {
                            group.date_range = Some(std::mem::take(&mut current_sparkline_dates));
                        }
                        state = ParseState::SparklineGroup;
                    }
                    (ParseState::SparklineSqref, "sqref") => {
                        state = ParseState::SparklineItem;
                    }
                    (ParseState::SparklineItem, "sparkline") => {
                        let formula = std::mem::take(&mut current_sparkline_formula);
                        let sqref = std::mem::take(&mut current_sparkline_sqref);
                        budget.store(
                            size_of::<Sparkline>() + formula.capacity() + sqref.capacity(),
                        )?;
                        current_sparklines.push(Sparkline { formula, sqref });
                        state = ParseState::Sparklines;
                    }
                    (ParseState::Sparklines, "sparklines") => {
                        state = ParseState::SparklineGroup;
                    }
                    (ParseState::SparklineGroup, "sparklineGroup") => {
                        if let Some(mut group) = current_sparkline_group.take() {
                            // Every string the group keeps is copied from its
                            // own span (its attributes, colour elements and
                            // date range), so the span bounds their bytes.
                            let span = reader.buffer_position().saturating_sub(group_start);
                            budget.store(
                                size_of::<SparklineGroup>()
                                    .saturating_add(usize::try_from(span).unwrap_or(usize::MAX)),
                            )?;
                            group.sparklines = take_exact(&mut current_sparklines);
                            sparkline_groups.push(group);
                        }
                        state = ParseState::SparklineGroups;
                    }
                    (ParseState::SparklineGroups, "sparklineGroups") => {
                        state = ParseState::ExtSparklines;
                    }
                    (ParseState::ExtSparklines, "ext") => {
                        state = ParseState::ExtLst;
                    }
                    (ParseState::ExtLst, "extLst") => {
                        state = ParseState::Root;
                    }
                    (ParseState::ExtOther(depth), _) => {
                        if depth <= 1 {
                            preserved_extensions.push(kept_ext(
                                data,
                                ext_start,
                                reader.buffer_position(),
                                ext_lst_ns.as_ref(),
                                &root_ns,
                                budget,
                            )?);
                            state = ParseState::ExtLst;
                        } else {
                            state = ParseState::ExtOther(depth - 1);
                        }
                    }
                    _ => {}
                }
            }
            Ok(Event::Eof) => break,
            Err(err) => {
                cold_path();
                return Err(ModernXlsxError::XmlParse(format!(
                    "error parsing worksheet XML: {err}"
                )));
            }
            _ => {}
        }
    }

    // quick-xml reports end of input as Eof even with elements still open: a
    // row that never ended is dropped, with any cell it holds, and the text
    // its cells counted is given back.
    if matches!(
        state,
        ParseState::InRow
            | ParseState::InCell
            | ParseState::InCellValue
            | ParseState::InCellFormula
            | ParseState::InInlineStr
            | ParseState::InInlineStrT
            | ParseState::InInlineRph
    ) {
        row_sink.abort_row();
        budget.spent = row_spent;
    }

    Ok(SheetMeta {
        dimension,
        merge_cells,
        auto_filter,
        frozen_pane,
        split_pane,
        pane_selections,
        sheet_view,
        columns,
        data_validations,
        conditional_formatting,
        hyperlinks,
        page_setup,
        sheet_protection,
        tab_color,
        tab_color_spec,
        header_footer,
        page_breaks,
        outline_properties,
        sparkline_groups,
        preserved_extensions,
    })
}

/// The [`captured_ext`] of `data[start..end]` with the namespace declarations
/// it is missing inserted ([`inherited_declarations`]): a prefix bound on the
/// `<extLst>` (`ext_lst`, when it declares any) takes that binding, any other
/// the `<worksheet>` root's (`root_ns`). Counted in `budget` before the string
/// is built, so a refused `<ext>` builds no kept string: the inserted
/// declarations as copied text (one declaration is copied onto every `<ext>`
/// that uses its prefix, as a shared string is copied into every cell naming
/// it), and the element's own bytes and the string's header as a list
/// element, so each byte of the kept string is counted once.
///
/// # Errors
///
/// As [`captured_ext`]; `SizeLimit` when the count passes the budget.
fn kept_ext(
    data: &[u8],
    start: u64,
    end: u64,
    ext_lst: Option<&RootNamespaces>,
    root_ns: &RootNamespaces,
    budget: &mut ExpansionBudget,
) -> Result<String> {
    let text = captured_ext(data, start, end)?;
    let both;
    let scopes: &[&RootNamespaces] = match ext_lst {
        Some(inner) => {
            both = [inner, root_ns];
            &both
        }
        None => std::slice::from_ref(&root_ns),
    };
    let inserted = inherited_declarations(text, scopes);
    budget.spend(
        inserted
            .as_ref()
            .map_or(0, |(_, decls)| decls.iter().map(|d| d.len()).sum()),
    )?;
    budget.store(size_of::<String>() + text.len())?;
    Ok(match inserted {
        Some((insert_at, decls)) => with_declarations(text, insert_at, &decls),
        None => text.to_owned(),
    })
}

/// The namespace declarations of an `<extLst>` start tag, indexed and stored
/// in `budget` when the element starts (nothing gives the count back at
/// `</extLst>`, so it is kept for the rest of the read; the index is replaced
/// at the next `<extLst>`); `None` when it declares none (as in every sheet Excel
/// writes) or when its attribute list is malformed, which leaves its `<ext>`
/// children to the root's declarations alone.
///
/// # Errors
///
/// `SizeLimit` when the count passes the budget.
fn ext_lst_namespaces(
    e: &BytesStart<'_>,
    budget: &mut ExpansionBudget,
) -> Result<Option<RootNamespaces>> {
    let Some(list) = root_namespaces(e).ok().filter(|l| !l.is_empty()) else {
        return Ok(None);
    };
    let ns = RootNamespaces::new(list);
    budget.store(ns.stored_bytes())?;
    Ok(Some(ns))
}

/// The heap bytes of the strings in `strings`.
fn strings_heap<'s>(strings: impl IntoIterator<Item = &'s Option<String>>) -> usize {
    strings.into_iter().flatten().map(String::capacity).sum()
}

/// The memory a kept [`DataValidation`] takes: its struct and its strings.
/// The destructuring names every field, so a new string cannot be left out.
fn data_validation_bytes(dv: &DataValidation) -> usize {
    let DataValidation {
        sqref,
        validation_type,
        operator,
        error_style,
        ime_mode,
        show_drop_down: _,
        formula1,
        formula2,
        allow_blank: _,
        show_error_message: _,
        error_title,
        error_message,
        show_input_message: _,
        prompt_title,
        prompt,
    } = dv;
    size_of::<DataValidation>()
        + sqref.capacity()
        + strings_heap([
            validation_type,
            operator,
            error_style,
            ime_mode,
            formula1,
            formula2,
            error_title,
            error_message,
            prompt_title,
            prompt,
        ])
}

/// The memory a kept [`ConditionalFormattingRule`] takes: its struct, its
/// strings and its extra formulas, with its data bar's colour and its icon
/// set's type. The values of its colour scale, data bar and icon set and the
/// colours of its colour scale are stored as they are read.
fn cf_rule_bytes(rule: &ConditionalFormattingRule) -> usize {
    let ConditionalFormattingRule {
        rule_type,
        priority: _,
        operator,
        formula,
        additional_formulas,
        dxf_id: _,
        stop_if_true: _,
        above_average: _,
        percent: _,
        bottom: _,
        text,
        time_period,
        rank: _,
        std_dev: _,
        equal_average: _,
        color_scale: _,
        data_bar,
        icon_set,
    } = rule;
    size_of::<ConditionalFormattingRule>()
        + rule_type.capacity()
        + strings_heap([operator, formula, text, time_period])
        + additional_formulas.capacity() * size_of::<String>()
        + additional_formulas
            .iter()
            .map(String::capacity)
            .sum::<usize>()
        + data_bar.as_ref().map_or(0, |b| {
            b.color.capacity() + b.color_spec.as_ref().map_or(0, |c| strings_heap([&c.rgb]))
        })
        + icon_set
            .as_ref()
            .map_or(0, |s| strings_heap([&s.icon_set_type]))
}

/// The memory a kept [`Hyperlink`] takes beside its `target`: its struct
/// and its other strings. The target is a copy of a relationship's target,
/// already counted as one when the hyperlink was read.
fn hyperlink_bytes(link: &Hyperlink) -> usize {
    let Hyperlink {
        cell_ref,
        target: _,
        location,
        display,
        tooltip,
    } = link;
    size_of::<Hyperlink>() + cell_ref.capacity() + strings_heap([location, display, tooltip])
}

/// The memory a kept [`PaneSelection`] takes: its struct and its strings.
fn selection_bytes(sel: &PaneSelection) -> usize {
    let PaneSelection {
        pane,
        active_cell,
        sqref,
    } = sel;
    size_of::<PaneSelection>() + strings_heap([pane, active_cell, sqref])
}

/// Number a `<row>` without `r` after the previous row (ECMA-376 Part 1,
/// §18.3.1.73 makes `r` optional), then make it the row the following cells
/// sit in: `prev_row` becomes its index and `next_col` restarts at column A.
///
/// # Errors
///
/// `XmlParse` when the inferred number would overflow `u32`.
fn place_row(ra: &mut RowAttrs, prev_row: &mut u32, next_col: &mut u32) -> Result<()> {
    if !ra.has_r {
        ra.index = prev_row.checked_add(1).ok_or_else(|| {
            cold_path();
            ModernXlsxError::XmlParse(format!(
                "<row> without r after row {prev_row}: row number overflows"
            ))
        })?;
    }
    *prev_row = ra.index;
    *next_col = 0;
    Ok(())
}

/// Give a `<c>` without `r` the next position in its row (ECMA-376 Part 1,
/// §18.3.1.4 makes `r` optional): the column after the previous cell, in
/// `row`. A present `r` is kept verbatim and sets `next_col` to the column
/// after its letters; one whose letters are no column still takes a slot.
///
/// # Errors
///
/// `XmlParse` when an `r`-less cell would fall past column XFD.
fn place_cell(ca: &mut CellAttrs, row: u32, next_col: &mut u32) -> Result<()> {
    if ca.reference.is_empty() {
        if *next_col > MAX_COL {
            cold_path();
            return Err(ModernXlsxError::XmlParse(format!(
                "<c> without r in row {row} falls past column XFD"
            )));
        }
        let mut letters = [0u8; 7];
        let len = col_to_letters_buf(*next_col, &mut letters);
        ca.reference
            .extend(letters[..len].iter().copied().map(char::from));
        ca.reference.push_str(itoa::Buffer::new().format(row));
        *next_col += 1;
    } else {
        let letters = ca
            .reference
            .bytes()
            .take_while(u8::is_ascii_alphabetic)
            .count();
        *next_col = letters_to_col(&ca.reference[..letters])
            .map_or_else(|_| next_col.saturating_add(1), |col| col + 1);
    }
    Ok(())
}
