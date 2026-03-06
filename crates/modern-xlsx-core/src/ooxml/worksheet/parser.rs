use core::hint::cold_path;

use quick_xml::events::Event;
use quick_xml::Reader;

use super::json::{cell_type_json_str, json_escape_to, write_f64_json};
use super::{
    AutoFilter, Cell, CellType, Cfvo, ColorScale, ColumnInfo, ConditionalFormatting,
    ConditionalFormattingRule, DataBar, DataValidation, FilterColumn, FrozenPane, HeaderFooter,
    Hyperlink, IconSet, OutlineProperties, PageSetup, PaneSelection, ParseState, Row,
    SheetProtection, SheetViewData, Sparkline, SparklineGroup, SplitPane, WorksheetXml,
    parse_col_element,
};
use crate::ooxml::push_entity;
use crate::{ModernXlsxError, Result};

impl WorksheetXml {
    /// Parse a worksheet XML file from raw bytes.
    pub fn parse(data: &[u8]) -> Result<Self> {
        Self::parse_with_sst(data, None)
    }

    /// Parse a worksheet XML file, optionally resolving shared string indices
    /// inline during parsing. This avoids a costly post-parse pass and
    /// eliminates intermediate index-string allocations for SharedString cells.
    pub fn parse_with_sst(data: &[u8], sst: Option<&crate::ooxml::shared_strings::SharedStringTable>) -> Result<Self> {
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
        let mut sheet_view = None::<SheetViewData>;
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

        // Sparkline/extLst parsing state.
        let mut sparkline_groups: Vec<SparklineGroup> = Vec::new();
        let mut preserved_extensions: Vec<String> = Vec::new();
        let mut current_sparkline_group: Option<SparklineGroup> = None;
        let mut current_sparklines: Vec<Sparkline> = Vec::new();
        let mut current_sparkline_formula = String::new();
        let mut current_sparkline_sqref = String::new();
        let mut ext_other_buf = String::new();

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
        let mut cur_cell_formula_r1: Option<String> = None;
        let mut cur_cell_formula_r2: Option<String> = None;
        let mut cur_cell_formula_dt2d: Option<bool> = None;
        let mut cur_cell_formula_dtr1: Option<bool> = None;
        let mut cur_cell_formula_dtr2: Option<bool> = None;

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
                            let mut sv = SheetViewData::default();
                            let mut has_non_default = false;
                            for attr in e.attributes().flatten() {
                                let val = std::str::from_utf8(&attr.value).unwrap_or_default();
                                match attr.key.local_name().as_ref() {
                                    b"showGridLines" if val == "0" => { sv.show_grid_lines = false; has_non_default = true; }
                                    b"showRowColHeaders" if val == "0" => { sv.show_row_col_headers = false; has_non_default = true; }
                                    b"showZeros" if val == "0" => { sv.show_zeros = false; has_non_default = true; }
                                    b"rightToLeft" if val == "1" => { sv.right_to_left = true; has_non_default = true; }
                                    b"tabSelected" if val == "1" => { sv.tab_selected = true; has_non_default = true; }
                                    b"showRuler" if val == "0" => { sv.show_ruler = false; has_non_default = true; }
                                    b"showOutlineSymbols" if val == "0" => { sv.show_outline_symbols = false; has_non_default = true; }
                                    b"showWhiteSpace" if val == "0" => { sv.show_white_space = false; has_non_default = true; }
                                    b"defaultGridColor" if val == "0" => { sv.default_grid_color = false; has_non_default = true; }
                                    b"zoomScale" => { sv.zoom_scale = val.parse().ok(); has_non_default = true; }
                                    b"zoomScaleNormal" => { sv.zoom_scale_normal = val.parse().ok(); has_non_default = true; }
                                    b"zoomScalePageLayoutView" => { sv.zoom_scale_page_layout_view = val.parse().ok(); has_non_default = true; }
                                    b"zoomScaleSheetLayoutView" => { sv.zoom_scale_sheet_layout_view = val.parse().ok(); has_non_default = true; }
                                    b"colorId" => { sv.color_id = val.parse().ok(); has_non_default = true; }
                                    b"view" if val != "normal" => { sv.view = Some(val.to_owned()); has_non_default = true; }
                                    _ => {}
                                }
                            }
                            if has_non_default {
                                sheet_view = Some(sv);
                            }
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
                            cur_cell_formula_r1 = None;
                            cur_cell_formula_r2 = None;
                            cur_cell_formula_dt2d = None;
                            cur_cell_formula_dtr1 = None;
                            cur_cell_formula_dtr2 = None;

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
                                    b"cm" if val == "1" => {
                                        cur_cell_dynamic_array = Some(true);
                                    }
                                    b"r1" => cur_cell_formula_r1 = Some(val.to_owned()),
                                    b"r2" => cur_cell_formula_r2 = Some(val.to_owned()),
                                    b"dt2D"
                                        if val == "1" || val.eq_ignore_ascii_case("true") =>
                                    {
                                        cur_cell_formula_dt2d = Some(true);
                                    }
                                    b"dtr1"
                                        if val == "1" || val.eq_ignore_ascii_case("true") =>
                                    {
                                        cur_cell_formula_dtr1 = Some(true);
                                    }
                                    b"dtr2"
                                        if val == "1" || val.eq_ignore_ascii_case("true") =>
                                    {
                                        cur_cell_formula_dtr2 = Some(true);
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
                            if let Some(attr) = e.attributes().flatten()
                                .find(|a| a.key.local_name().as_ref() == b"sqref")
                            {
                                cur_cf_sqref = std::str::from_utf8(&attr.value)
                                    .unwrap_or_default()
                                    .to_owned();
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
                            if let Some(attr) = e.attributes().flatten()
                                .find(|a| a.key.local_name().as_ref() == b"iconSet")
                            {
                                cur_icon_set_type = Some(
                                    std::str::from_utf8(&attr.value)
                                        .unwrap_or_default()
                                        .to_owned(),
                                );
                            }
                        }
                        (ParseState::Root, b"hyperlinks") => {
                            state = ParseState::InHyperlinks;
                        }
                        (ParseState::Root, b"autoFilter") => {
                            state = ParseState::InAutoFilter;
                            cur_af_range.clear();
                            cur_af_columns.clear();
                            if let Some(attr) = e.attributes().flatten()
                                .find(|a| a.key.local_name().as_ref() == b"ref")
                            {
                                cur_af_range = std::str::from_utf8(&attr.value)
                                    .unwrap_or_default()
                                    .to_owned();
                            }
                        }
                        (ParseState::InAutoFilter, b"filterColumn") => {
                            state = ParseState::InFilterColumn;
                            cur_filter_col_id = 0;
                            cur_filter_vals.clear();
                            if let Some(attr) = e.attributes().flatten()
                                .find(|a| a.key.local_name().as_ref() == b"colId")
                            {
                                cur_filter_col_id = std::str::from_utf8(&attr.value)
                                    .unwrap_or_default()
                                    .parse::<u32>()
                                    .unwrap_or(0);
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

                        // ---- extLst / sparklines ----
                        (ParseState::Root, b"extLst") => {
                            state = ParseState::ExtLst;
                        }
                        (ParseState::ExtLst, b"ext") => {
                            let is_sparkline_ext = e.attributes().flatten()
                                .find(|a| a.key.local_name().as_ref() == b"uri")
                                .is_some_and(|attr| {
                                    let uri = std::str::from_utf8(&attr.value).unwrap_or_default();
                                    uri.contains("05C60535")
                                });
                            if is_sparkline_ext {
                                state = ParseState::ExtSparklines;
                            } else {
                                ext_other_buf.clear();
                                state = ParseState::ExtOther(1);
                            }
                        }
                        (ParseState::ExtSparklines, b"sparklineGroups") => {
                            state = ParseState::SparklineGroups;
                        }
                        (ParseState::SparklineGroups, b"sparklineGroup") => {
                            let mut group = SparklineGroup::default();
                            for attr in e.attributes().flatten() {
                                let val = std::str::from_utf8(&attr.value).unwrap_or_default();
                                match attr.key.local_name().as_ref() {
                                    b"type" => group.sparkline_type = val.to_owned(),
                                    b"displayEmptyCellsAs" => group.display_empty_cells_as = Some(val.to_owned()),
                                    b"markers" => group.markers = val == "1" || val.eq_ignore_ascii_case("true"),
                                    b"high" => group.high = val == "1" || val.eq_ignore_ascii_case("true"),
                                    b"low" => group.low = val == "1" || val.eq_ignore_ascii_case("true"),
                                    b"first" => group.first = val == "1" || val.eq_ignore_ascii_case("true"),
                                    b"last" => group.last = val == "1" || val.eq_ignore_ascii_case("true"),
                                    b"negative" => group.negative = val == "1" || val.eq_ignore_ascii_case("true"),
                                    b"displayXAxis" => group.display_x_axis = val == "1" || val.eq_ignore_ascii_case("true"),
                                    b"lineWeight" => group.line_weight = val.parse::<f64>().ok(),
                                    b"manualMin" => group.manual_min = val.parse::<f64>().ok(),
                                    b"manualMax" => group.manual_max = val.parse::<f64>().ok(),
                                    b"rightToLeft" => group.right_to_left = val == "1" || val.eq_ignore_ascii_case("true"),
                                    _ => {}
                                }
                            }
                            current_sparkline_group = Some(group);
                            state = ParseState::SparklineGroup;
                        }
                        (ParseState::SparklineGroup, b"sparklines") => {
                            state = ParseState::Sparklines;
                        }
                        (ParseState::Sparklines, b"sparkline") => {
                            current_sparkline_formula.clear();
                            current_sparkline_sqref.clear();
                            state = ParseState::SparklineItem;
                        }
                        (ParseState::SparklineItem, b"f") => {
                            state = ParseState::SparklineFormula;
                        }
                        (ParseState::SparklineItem, b"sqref") => {
                            state = ParseState::SparklineSqref;
                        }
                        (ParseState::ExtOther(_), _) => {
                            if let ParseState::ExtOther(ref mut depth) = state {
                                *depth += 1;
                            }
                        }
                        _ => {}
                    }
                }
                Ok(Event::Empty(ref e)) => {
                    let local = e.local_name();
                    match (state, local.as_ref()) {
                        (ParseState::Root, b"dimension") => {
                            if let Some(attr) = e.attributes().flatten()
                                .find(|a| a.key.local_name().as_ref() == b"ref")
                            {
                                dimension = Some(
                                    std::str::from_utf8(&attr.value).unwrap_or_default().to_owned(),
                                );
                            }
                        }
                        (ParseState::SheetPr, b"tabColor") => {
                            if let Some(attr) = e.attributes().flatten()
                                .find(|a| a.key.local_name().as_ref() == b"rgb")
                            {
                                tab_color = Some(
                                    std::str::from_utf8(&attr.value).unwrap_or_default().to_owned(),
                                );
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
                            if let Some(attr) = e.attributes().flatten()
                                .find(|a| a.key.local_name().as_ref() == b"ref")
                            {
                                af_range = std::str::from_utf8(&attr.value)
                                    .unwrap_or_default()
                                    .to_owned();
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
                        (ParseState::SheetViews, b"sheetView") => {
                            // Self-closing <sheetView ... /> — parse attributes.
                            let mut sv = SheetViewData::default();
                            let mut has_non_default = false;
                            for attr in e.attributes().flatten() {
                                let val = std::str::from_utf8(&attr.value).unwrap_or_default();
                                match attr.key.local_name().as_ref() {
                                    b"showGridLines" if val == "0" => { sv.show_grid_lines = false; has_non_default = true; }
                                    b"showRowColHeaders" if val == "0" => { sv.show_row_col_headers = false; has_non_default = true; }
                                    b"showZeros" if val == "0" => { sv.show_zeros = false; has_non_default = true; }
                                    b"rightToLeft" if val == "1" => { sv.right_to_left = true; has_non_default = true; }
                                    b"tabSelected" if val == "1" => { sv.tab_selected = true; has_non_default = true; }
                                    b"showRuler" if val == "0" => { sv.show_ruler = false; has_non_default = true; }
                                    b"showOutlineSymbols" if val == "0" => { sv.show_outline_symbols = false; has_non_default = true; }
                                    b"showWhiteSpace" if val == "0" => { sv.show_white_space = false; has_non_default = true; }
                                    b"defaultGridColor" if val == "0" => { sv.default_grid_color = false; has_non_default = true; }
                                    b"zoomScale" => { sv.zoom_scale = val.parse().ok(); has_non_default = true; }
                                    b"zoomScaleNormal" => { sv.zoom_scale_normal = val.parse().ok(); has_non_default = true; }
                                    b"zoomScalePageLayoutView" => { sv.zoom_scale_page_layout_view = val.parse().ok(); has_non_default = true; }
                                    b"zoomScaleSheetLayoutView" => { sv.zoom_scale_sheet_layout_view = val.parse().ok(); has_non_default = true; }
                                    b"colorId" => { sv.color_id = val.parse().ok(); has_non_default = true; }
                                    b"view" if val != "normal" => { sv.view = Some(val.to_owned()); has_non_default = true; }
                                    _ => {}
                                }
                            }
                            if has_non_default {
                                sheet_view = Some(sv);
                            }
                        }
                        (ParseState::Cols, b"col") => {
                            columns.push(parse_col_element(e));
                        }
                        (ParseState::MergeCells, b"mergeCell") => {
                            if let Some(attr) = e.attributes().flatten()
                                .find(|a| a.key.local_name().as_ref() == b"ref")
                            {
                                merge_cells.push(
                                    std::str::from_utf8(&attr.value).unwrap_or_default().to_owned(),
                                );
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
                                ..Default::default()
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
                                    b"cm" if val == "1" => {
                                        cur_cell_dynamic_array = Some(true);
                                    }
                                    b"r1" => cur_cell_formula_r1 = Some(val.to_owned()),
                                    b"r2" => cur_cell_formula_r2 = Some(val.to_owned()),
                                    b"dt2D"
                                        if val == "1" || val.eq_ignore_ascii_case("true") =>
                                    {
                                        cur_cell_formula_dt2d = Some(true);
                                    }
                                    b"dtr1"
                                        if val == "1" || val.eq_ignore_ascii_case("true") =>
                                    {
                                        cur_cell_formula_dtr1 = Some(true);
                                    }
                                    b"dtr2"
                                        if val == "1" || val.eq_ignore_ascii_case("true") =>
                                    {
                                        cur_cell_formula_dtr2 = Some(true);
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
                            if let Some(attr) = e.attributes().flatten()
                                .find(|a| a.key.local_name().as_ref() == b"rgb")
                            {
                                cur_cf_colors.push(
                                    std::str::from_utf8(&attr.value)
                                        .unwrap_or_default()
                                        .to_owned(),
                                );
                            }
                        }
                        // color in dataBar.
                        (ParseState::InDataBar, b"color") => {
                            if let Some(attr) = e.attributes().flatten()
                                .find(|a| a.key.local_name().as_ref() == b"rgb")
                            {
                                cur_cf_bar_color = std::str::from_utf8(&attr.value)
                                    .unwrap_or_default()
                                    .to_owned();
                            }
                        }
                        // filter values.
                        (ParseState::InFilters, b"filter") => {
                            if let Some(attr) = e.attributes().flatten()
                                .find(|a| a.key.local_name().as_ref() == b"val")
                            {
                                cur_filter_vals.push(
                                    std::str::from_utf8(&attr.value)
                                        .unwrap_or_default()
                                        .to_owned(),
                                );
                            }
                        }
                        // Sparkline color elements (self-closing).
                        (ParseState::SparklineGroup, _) => {
                            if let Some(ref mut group) = current_sparkline_group
                                && let Some(attr) = e.attributes().flatten()
                                    .find(|a| a.key.local_name().as_ref() == b"rgb")
                            {
                                let rgb = std::str::from_utf8(&attr.value).unwrap_or_default().to_owned();
                                match local.as_ref() {
                                    b"colorSeries" => group.color_series = Some(rgb),
                                    b"colorNegative" => group.color_negative = Some(rgb),
                                    b"colorAxis" => group.color_axis = Some(rgb),
                                    b"colorMarkers" => group.color_markers = Some(rgb),
                                    b"colorFirst" => group.color_first = Some(rgb),
                                    b"colorLast" => group.color_last = Some(rgb),
                                    b"colorHigh" => group.color_high = Some(rgb),
                                    b"colorLow" => group.color_low = Some(rgb),
                                    _ => {}
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
                        ParseState::SparklineFormula => {
                            current_sparkline_formula.push_str(std::str::from_utf8(e.as_ref()).unwrap_or_default());
                        }
                        ParseState::SparklineSqref => {
                            current_sparkline_sqref.push_str(std::str::from_utf8(e.as_ref()).unwrap_or_default());
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
                        ParseState::SparklineFormula => {
                            push_entity(&mut current_sparkline_formula, e.as_ref());
                        }
                        ParseState::SparklineSqref => {
                            push_entity(&mut current_sparkline_sqref, e.as_ref());
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
                            cur_cell_inline_string = Some(text.clone());
                            cur_cell_value = Some(text);
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
                                formula_r1: cur_cell_formula_r1.take(),
                                formula_r2: cur_cell_formula_r2.take(),
                                formula_dt2d: cur_cell_formula_dt2d.take(),
                                formula_dtr1: cur_cell_formula_dtr1.take(),
                                formula_dtr2: cur_cell_formula_dtr2.take(),
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

                        // ---- extLst / sparklines end ----
                        (ParseState::SparklineFormula, b"f") => {
                            state = ParseState::SparklineItem;
                        }
                        (ParseState::SparklineSqref, b"sqref") => {
                            state = ParseState::SparklineItem;
                        }
                        (ParseState::SparklineItem, b"sparkline") => {
                            current_sparklines.push(Sparkline {
                                formula: std::mem::take(&mut current_sparkline_formula),
                                sqref: std::mem::take(&mut current_sparkline_sqref),
                            });
                            state = ParseState::Sparklines;
                        }
                        (ParseState::Sparklines, b"sparklines") => {
                            state = ParseState::SparklineGroup;
                        }
                        (ParseState::SparklineGroup, b"sparklineGroup") => {
                            if let Some(mut group) = current_sparkline_group.take() {
                                group.sparklines = std::mem::take(&mut current_sparklines);
                                sparkline_groups.push(group);
                            }
                            state = ParseState::SparklineGroups;
                        }
                        (ParseState::SparklineGroups, b"sparklineGroups") => {
                            state = ParseState::ExtSparklines;
                        }
                        (ParseState::ExtSparklines, b"ext") => {
                            state = ParseState::ExtLst;
                        }
                        (ParseState::ExtLst, b"extLst") => {
                            state = ParseState::Root;
                        }
                        (ParseState::ExtOther(depth), _) => {
                            if depth <= 1 {
                                let ext_xml = std::mem::take(&mut ext_other_buf);
                                if !ext_xml.is_empty() {
                                    preserved_extensions.push(ext_xml);
                                }
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
            sheet_view,
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
            sparkline_groups,
            charts: Vec::new(),
            pivot_tables: Vec::new(),
            threaded_comments: Vec::new(),
            slicers: Vec::new(),
            timelines: Vec::new(),
            preserved_extensions,
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
        sst: Option<&crate::ooxml::shared_strings::SharedStringTable>,
        comments: &[crate::ooxml::comments::Comment],
        tables: &[crate::ooxml::tables::TableDefinition],
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
        let mut sheet_view = None::<SheetViewData>;
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

        // Sparkline/extLst parsing state.
        let mut sparkline_groups: Vec<SparklineGroup> = Vec::new();
        let mut current_sparkline_group: Option<SparklineGroup> = None;
        let mut current_sparklines: Vec<Sparkline> = Vec::new();
        let mut current_sparkline_formula = String::new();
        let mut current_sparkline_sqref = String::new();

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
        let mut cur_cell_formula_r1: Option<String> = None;
        let mut cur_cell_formula_r2: Option<String> = None;
        let mut cur_cell_formula_dt2d: Option<bool> = None;
        let mut cur_cell_formula_dtr1: Option<bool> = None;
        let mut cur_cell_formula_dtr2: Option<bool> = None;

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
                        (ParseState::SheetViews, b"sheetView") => {
                            state = ParseState::SheetView;
                            let mut sv = SheetViewData::default();
                            let mut has_non_default = false;
                            for attr in e.attributes().flatten() {
                                let val = std::str::from_utf8(&attr.value).unwrap_or_default();
                                match attr.key.local_name().as_ref() {
                                    b"showGridLines" if val == "0" => { sv.show_grid_lines = false; has_non_default = true; }
                                    b"showRowColHeaders" if val == "0" => { sv.show_row_col_headers = false; has_non_default = true; }
                                    b"showZeros" if val == "0" => { sv.show_zeros = false; has_non_default = true; }
                                    b"rightToLeft" if val == "1" => { sv.right_to_left = true; has_non_default = true; }
                                    b"tabSelected" if val == "1" => { sv.tab_selected = true; has_non_default = true; }
                                    b"showRuler" if val == "0" => { sv.show_ruler = false; has_non_default = true; }
                                    b"showOutlineSymbols" if val == "0" => { sv.show_outline_symbols = false; has_non_default = true; }
                                    b"showWhiteSpace" if val == "0" => { sv.show_white_space = false; has_non_default = true; }
                                    b"defaultGridColor" if val == "0" => { sv.default_grid_color = false; has_non_default = true; }
                                    b"zoomScale" => { sv.zoom_scale = val.parse().ok(); has_non_default = true; }
                                    b"zoomScaleNormal" => { sv.zoom_scale_normal = val.parse().ok(); has_non_default = true; }
                                    b"zoomScalePageLayoutView" => { sv.zoom_scale_page_layout_view = val.parse().ok(); has_non_default = true; }
                                    b"zoomScaleSheetLayoutView" => { sv.zoom_scale_sheet_layout_view = val.parse().ok(); has_non_default = true; }
                                    b"colorId" => { sv.color_id = val.parse().ok(); has_non_default = true; }
                                    b"view" if val != "normal" => { sv.view = Some(val.to_owned()); has_non_default = true; }
                                    _ => {}
                                }
                            }
                            if has_non_default {
                                sheet_view = Some(sv);
                            }
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
                        (ParseState::Root, b"cols") => state = ParseState::Cols,

                        // ---- Sheet data (rows/cells -> streamed to JSON) ----
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
                            cur_cell_formula_r1 = None;
                            cur_cell_formula_r2 = None;
                            cur_cell_formula_dt2d = None;
                            cur_cell_formula_dtr1 = None;
                            cur_cell_formula_dtr2 = None;

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
                            json_escape_to(out, &cur_cell_ref);
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
                                    b"cm" if val == "1" => {
                                        cur_cell_dynamic_array = Some(true);
                                    }
                                    b"r1" => cur_cell_formula_r1 = Some(val.to_owned()),
                                    b"r2" => cur_cell_formula_r2 = Some(val.to_owned()),
                                    b"dt2D"
                                        if val == "1" || val.eq_ignore_ascii_case("true") =>
                                    {
                                        cur_cell_formula_dt2d = Some(true);
                                    }
                                    b"dtr1"
                                        if val == "1" || val.eq_ignore_ascii_case("true") =>
                                    {
                                        cur_cell_formula_dtr1 = Some(true);
                                    }
                                    b"dtr2"
                                        if val == "1" || val.eq_ignore_ascii_case("true") =>
                                    {
                                        cur_cell_formula_dtr2 = Some(true);
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
                            if let Some(attr) = e.attributes().flatten()
                                .find(|a| a.key.local_name().as_ref() == b"sqref")
                            {
                                cur_cf_sqref = std::str::from_utf8(&attr.value)
                                    .unwrap_or_default()
                                    .to_owned();
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
                            if let Some(attr) = e.attributes().flatten()
                                .find(|a| a.key.local_name().as_ref() == b"iconSet")
                            {
                                cur_icon_set_type = Some(
                                    std::str::from_utf8(&attr.value)
                                        .unwrap_or_default()
                                        .to_owned(),
                                );
                            }
                        }

                        // ---- Hyperlinks (metadata) ----
                        (ParseState::Root, b"hyperlinks") => state = ParseState::InHyperlinks,

                        // ---- Auto filter (metadata) ----
                        (ParseState::Root, b"autoFilter") => {
                            state = ParseState::InAutoFilter;
                            cur_af_range.clear();
                            cur_af_columns.clear();
                            if let Some(attr) = e.attributes().flatten()
                                .find(|a| a.key.local_name().as_ref() == b"ref")
                            {
                                cur_af_range = std::str::from_utf8(&attr.value)
                                    .unwrap_or_default()
                                    .to_owned();
                            }
                        }
                        (ParseState::InAutoFilter, b"filterColumn") => {
                            state = ParseState::InFilterColumn;
                            cur_filter_col_id = 0;
                            cur_filter_vals.clear();
                            if let Some(attr) = e.attributes().flatten()
                                .find(|a| a.key.local_name().as_ref() == b"colId")
                            {
                                cur_filter_col_id = std::str::from_utf8(&attr.value)
                                    .unwrap_or_default()
                                    .parse::<u32>()
                                    .unwrap_or(0);
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

                        // ---- extLst / sparklines ----
                        (ParseState::Root, b"extLst") => {
                            state = ParseState::ExtLst;
                        }
                        (ParseState::ExtLst, b"ext") => {
                            let is_sparkline_ext = e.attributes().flatten()
                                .find(|a| a.key.local_name().as_ref() == b"uri")
                                .is_some_and(|attr| {
                                    let uri = std::str::from_utf8(&attr.value).unwrap_or_default();
                                    uri.contains("05C60535")
                                });
                            if is_sparkline_ext {
                                state = ParseState::ExtSparklines;
                            } else {
                                state = ParseState::ExtOther(1);
                            }
                        }
                        (ParseState::ExtSparklines, b"sparklineGroups") => {
                            state = ParseState::SparklineGroups;
                        }
                        (ParseState::SparklineGroups, b"sparklineGroup") => {
                            let mut group = SparklineGroup::default();
                            for attr in e.attributes().flatten() {
                                let val = std::str::from_utf8(&attr.value).unwrap_or_default();
                                match attr.key.local_name().as_ref() {
                                    b"type" => group.sparkline_type = val.to_owned(),
                                    b"displayEmptyCellsAs" => group.display_empty_cells_as = Some(val.to_owned()),
                                    b"markers" => group.markers = val == "1" || val.eq_ignore_ascii_case("true"),
                                    b"high" => group.high = val == "1" || val.eq_ignore_ascii_case("true"),
                                    b"low" => group.low = val == "1" || val.eq_ignore_ascii_case("true"),
                                    b"first" => group.first = val == "1" || val.eq_ignore_ascii_case("true"),
                                    b"last" => group.last = val == "1" || val.eq_ignore_ascii_case("true"),
                                    b"negative" => group.negative = val == "1" || val.eq_ignore_ascii_case("true"),
                                    b"displayXAxis" => group.display_x_axis = val == "1" || val.eq_ignore_ascii_case("true"),
                                    b"lineWeight" => group.line_weight = val.parse::<f64>().ok(),
                                    b"manualMin" => group.manual_min = val.parse::<f64>().ok(),
                                    b"manualMax" => group.manual_max = val.parse::<f64>().ok(),
                                    b"rightToLeft" => group.right_to_left = val == "1" || val.eq_ignore_ascii_case("true"),
                                    _ => {}
                                }
                            }
                            current_sparkline_group = Some(group);
                            state = ParseState::SparklineGroup;
                        }
                        (ParseState::SparklineGroup, b"sparklines") => {
                            state = ParseState::Sparklines;
                        }
                        (ParseState::Sparklines, b"sparkline") => {
                            current_sparkline_formula.clear();
                            current_sparkline_sqref.clear();
                            state = ParseState::SparklineItem;
                        }
                        (ParseState::SparklineItem, b"f") => {
                            state = ParseState::SparklineFormula;
                        }
                        (ParseState::SparklineItem, b"sqref") => {
                            state = ParseState::SparklineSqref;
                        }
                        (ParseState::ExtOther(_), _) => {
                            if let ParseState::ExtOther(ref mut depth) = state {
                                *depth += 1;
                            }
                        }
                        _ => {}
                    }
                }

                Ok(Event::Empty(ref e)) => {
                    let local = e.local_name();
                    match (state, local.as_ref()) {
                        (ParseState::Root, b"dimension") => {
                            if let Some(attr) = e.attributes().flatten()
                                .find(|a| a.key.local_name().as_ref() == b"ref")
                            {
                                dimension = Some(
                                    std::str::from_utf8(&attr.value).unwrap_or_default().to_owned(),
                                );
                            }
                        }
                        (ParseState::SheetPr, b"tabColor") => {
                            if let Some(attr) = e.attributes().flatten()
                                .find(|a| a.key.local_name().as_ref() == b"rgb")
                            {
                                tab_color = Some(
                                    std::str::from_utf8(&attr.value).unwrap_or_default().to_owned(),
                                );
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
                            if let Some(attr) = e.attributes().flatten()
                                .find(|a| a.key.local_name().as_ref() == b"ref")
                            {
                                af_range = std::str::from_utf8(&attr.value)
                                    .unwrap_or_default()
                                    .to_owned();
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
                        (ParseState::SheetViews, b"sheetView") => {
                            // Self-closing <sheetView ... /> — parse attributes.
                            let mut sv = SheetViewData::default();
                            let mut has_non_default = false;
                            for attr in e.attributes().flatten() {
                                let val = std::str::from_utf8(&attr.value).unwrap_or_default();
                                match attr.key.local_name().as_ref() {
                                    b"showGridLines" if val == "0" => { sv.show_grid_lines = false; has_non_default = true; }
                                    b"showRowColHeaders" if val == "0" => { sv.show_row_col_headers = false; has_non_default = true; }
                                    b"showZeros" if val == "0" => { sv.show_zeros = false; has_non_default = true; }
                                    b"rightToLeft" if val == "1" => { sv.right_to_left = true; has_non_default = true; }
                                    b"tabSelected" if val == "1" => { sv.tab_selected = true; has_non_default = true; }
                                    b"showRuler" if val == "0" => { sv.show_ruler = false; has_non_default = true; }
                                    b"showOutlineSymbols" if val == "0" => { sv.show_outline_symbols = false; has_non_default = true; }
                                    b"showWhiteSpace" if val == "0" => { sv.show_white_space = false; has_non_default = true; }
                                    b"defaultGridColor" if val == "0" => { sv.default_grid_color = false; has_non_default = true; }
                                    b"zoomScale" => { sv.zoom_scale = val.parse().ok(); has_non_default = true; }
                                    b"zoomScaleNormal" => { sv.zoom_scale_normal = val.parse().ok(); has_non_default = true; }
                                    b"zoomScalePageLayoutView" => { sv.zoom_scale_page_layout_view = val.parse().ok(); has_non_default = true; }
                                    b"zoomScaleSheetLayoutView" => { sv.zoom_scale_sheet_layout_view = val.parse().ok(); has_non_default = true; }
                                    b"colorId" => { sv.color_id = val.parse().ok(); has_non_default = true; }
                                    b"view" if val != "normal" => { sv.view = Some(val.to_owned()); has_non_default = true; }
                                    _ => {}
                                }
                            }
                            if has_non_default {
                                sheet_view = Some(sv);
                            }
                        }
                        (ParseState::Cols, b"col") => columns.push(parse_col_element(e)),
                        (ParseState::MergeCells, b"mergeCell") => {
                            if let Some(attr) = e.attributes().flatten()
                                .find(|a| a.key.local_name().as_ref() == b"ref")
                            {
                                merge_cells.push(
                                    std::str::from_utf8(&attr.value).unwrap_or_default().to_owned(),
                                );
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
                            json_escape_to(out, &cell_ref_buf);
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
                                    b"cm" if val == "1" => {
                                        cur_cell_dynamic_array = Some(true);
                                    }
                                    b"r1" => cur_cell_formula_r1 = Some(val.to_owned()),
                                    b"r2" => cur_cell_formula_r2 = Some(val.to_owned()),
                                    b"dt2D"
                                        if val == "1" || val.eq_ignore_ascii_case("true") =>
                                    {
                                        cur_cell_formula_dt2d = Some(true);
                                    }
                                    b"dtr1"
                                        if val == "1" || val.eq_ignore_ascii_case("true") =>
                                    {
                                        cur_cell_formula_dtr1 = Some(true);
                                    }
                                    b"dtr2"
                                        if val == "1" || val.eq_ignore_ascii_case("true") =>
                                    {
                                        cur_cell_formula_dtr2 = Some(true);
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
                            if let Some(attr) = e.attributes().flatten()
                                .find(|a| a.key.local_name().as_ref() == b"rgb")
                            {
                                cur_cf_colors.push(
                                    std::str::from_utf8(&attr.value).unwrap_or_default().to_owned(),
                                );
                            }
                        }
                        (ParseState::InDataBar, b"color") => {
                            if let Some(attr) = e.attributes().flatten()
                                .find(|a| a.key.local_name().as_ref() == b"rgb")
                            {
                                cur_cf_bar_color = std::str::from_utf8(&attr.value)
                                    .unwrap_or_default()
                                    .to_owned();
                            }
                        }
                        (ParseState::InFilters, b"filter") => {
                            if let Some(attr) = e.attributes().flatten()
                                .find(|a| a.key.local_name().as_ref() == b"val")
                            {
                                cur_filter_vals.push(
                                    std::str::from_utf8(&attr.value).unwrap_or_default().to_owned(),
                                );
                            }
                        }
                        // Sparkline color elements (self-closing).
                        (ParseState::SparklineGroup, _) => {
                            if let Some(ref mut group) = current_sparkline_group
                                && let Some(attr) = e.attributes().flatten()
                                    .find(|a| a.key.local_name().as_ref() == b"rgb")
                            {
                                let rgb = std::str::from_utf8(&attr.value).unwrap_or_default().to_owned();
                                match local.as_ref() {
                                    b"colorSeries" => group.color_series = Some(rgb),
                                    b"colorNegative" => group.color_negative = Some(rgb),
                                    b"colorAxis" => group.color_axis = Some(rgb),
                                    b"colorMarkers" => group.color_markers = Some(rgb),
                                    b"colorFirst" => group.color_first = Some(rgb),
                                    b"colorLast" => group.color_last = Some(rgb),
                                    b"colorHigh" => group.color_high = Some(rgb),
                                    b"colorLow" => group.color_low = Some(rgb),
                                    _ => {}
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
                        ParseState::SparklineFormula => {
                            current_sparkline_formula.push_str(std::str::from_utf8(e.as_ref()).unwrap_or_default());
                        }
                        ParseState::SparklineSqref => {
                            current_sparkline_sqref.push_str(std::str::from_utf8(e.as_ref()).unwrap_or_default());
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
                        ParseState::SparklineFormula => {
                            push_entity(&mut current_sparkline_formula, e.as_ref());
                        }
                        ParseState::SparklineSqref => {
                            push_entity(&mut current_sparkline_sqref, e.as_ref());
                        }
                        _ => {}
                    }
                }

                Ok(Event::End(ref e)) => {
                    let local = e.local_name();
                    match (state, local.as_ref()) {
                        // ---- Cell value end -> write JSON directly ----
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

                        // ---- Formula end -> write JSON directly ----
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
                            if let Some(ref r1) = cur_cell_formula_r1 {
                                out.push_str(",\"formulaR1\":\"");
                                json_escape_to(out, r1);
                                out.push('"');
                            }
                            if let Some(ref r2) = cur_cell_formula_r2 {
                                out.push_str(",\"formulaR2\":\"");
                                json_escape_to(out, r2);
                                out.push('"');
                            }
                            if cur_cell_formula_dt2d == Some(true) {
                                out.push_str(",\"formulaDt2d\":true");
                            }
                            if cur_cell_formula_dtr1 == Some(true) {
                                out.push_str(",\"formulaDtr1\":true");
                            }
                            if cur_cell_formula_dtr2 == Some(true) {
                                out.push_str(",\"formulaDtr2\":true");
                            }
                            state = ParseState::InCell;
                        }

                        // ---- Inline string end -> write both value and inlineString ----
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

                        // ---- Cell end -> close cell JSON object ----
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
                            if let Some(ref r1) = cur_cell_formula_r1.take() {
                                out.push_str(",\"formulaR1\":\"");
                                json_escape_to(out, r1);
                                out.push('"');
                            }
                            if let Some(ref r2) = cur_cell_formula_r2.take() {
                                out.push_str(",\"formulaR2\":\"");
                                json_escape_to(out, r2);
                                out.push('"');
                            }
                            if cur_cell_formula_dt2d.take() == Some(true) {
                                out.push_str(",\"formulaDt2d\":true");
                            }
                            if cur_cell_formula_dtr1.take() == Some(true) {
                                out.push_str(",\"formulaDtr1\":true");
                            }
                            if cur_cell_formula_dtr2.take() == Some(true) {
                                out.push_str(",\"formulaDtr2\":true");
                            }
                            out.push('}');
                            state = ParseState::InRow;
                        }

                        // ---- Row end -> close row JSON object ----
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

                        // ---- extLst / sparklines end ----
                        (ParseState::SparklineFormula, b"f") => {
                            state = ParseState::SparklineItem;
                        }
                        (ParseState::SparklineSqref, b"sqref") => {
                            state = ParseState::SparklineItem;
                        }
                        (ParseState::SparklineItem, b"sparkline") => {
                            current_sparklines.push(Sparkline {
                                formula: std::mem::take(&mut current_sparkline_formula),
                                sqref: std::mem::take(&mut current_sparkline_sqref),
                            });
                            state = ParseState::Sparklines;
                        }
                        (ParseState::Sparklines, b"sparklines") => {
                            state = ParseState::SparklineGroup;
                        }
                        (ParseState::SparklineGroup, b"sparklineGroup") => {
                            if let Some(mut group) = current_sparkline_group.take() {
                                group.sparklines = std::mem::take(&mut current_sparklines);
                                sparkline_groups.push(group);
                            }
                            state = ParseState::SparklineGroups;
                        }
                        (ParseState::SparklineGroups, b"sparklineGroups") => {
                            state = ParseState::ExtSparklines;
                        }
                        (ParseState::ExtSparklines, b"ext") => {
                            state = ParseState::ExtLst;
                        }
                        (ParseState::ExtLst, b"extLst") => {
                            state = ParseState::Root;
                        }
                        (ParseState::ExtOther(depth), _) => {
                            if depth <= 1 {
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
            out.push_str(&serde_json::to_string(&merge_cells)?);
        }
        if let Some(ref af) = auto_filter {
            out.push_str(",\"autoFilter\":");
            out.push_str(&serde_json::to_string(af)?);
        }
        if let Some(ref fp) = frozen_pane {
            out.push_str(",\"frozenPane\":");
            out.push_str(&serde_json::to_string(fp)?);
        }
        if let Some(ref sp) = split_pane {
            out.push_str(",\"splitPane\":");
            out.push_str(&serde_json::to_string(sp)?);
        }
        if !pane_selections.is_empty() {
            out.push_str(",\"paneSelections\":");
            out.push_str(&serde_json::to_string(&pane_selections)?);
        }
        if let Some(ref sv) = sheet_view {
            out.push_str(",\"sheetView\":");
            out.push_str(&serde_json::to_string(sv)?);
        }
        if !columns.is_empty() {
            out.push_str(",\"columns\":");
            out.push_str(&serde_json::to_string(&columns)?);
        }
        if !data_validations.is_empty() {
            out.push_str(",\"dataValidations\":");
            out.push_str(&serde_json::to_string(&data_validations)?);
        }
        if !conditional_formatting.is_empty() {
            out.push_str(",\"conditionalFormatting\":");
            out.push_str(&serde_json::to_string(&conditional_formatting)?);
        }
        if !hyperlinks.is_empty() {
            out.push_str(",\"hyperlinks\":");
            out.push_str(&serde_json::to_string(&hyperlinks)?);
        }
        if let Some(ref ps) = page_setup {
            out.push_str(",\"pageSetup\":");
            out.push_str(&serde_json::to_string(ps)?);
        }
        if let Some(ref sp) = sheet_protection {
            out.push_str(",\"sheetProtection\":");
            out.push_str(&serde_json::to_string(sp)?);
        }
        if !comments.is_empty() {
            out.push_str(",\"comments\":");
            out.push_str(&serde_json::to_string(comments)?);
        }
        if !tables.is_empty() {
            out.push_str(",\"tables\":");
            out.push_str(&serde_json::to_string(tables)?);
        }
        if let Some(ref tc) = tab_color {
            out.push_str(",\"tabColor\":\"");
            json_escape_to(out, tc);
            out.push('"');
        }
        if let Some(ref hf) = header_footer {
            out.push_str(",\"headerFooter\":");
            out.push_str(&serde_json::to_string(hf)?);
        }
        if let Some(ref op) = outline_properties {
            out.push_str(",\"outlineProperties\":");
            out.push_str(&serde_json::to_string(op)?);
        }
        if !sparkline_groups.is_empty() {
            out.push_str(",\"sparklineGroups\":");
            out.push_str(&serde_json::to_string(&sparkline_groups)?);
        }

        // Close the worksheet JSON object.
        out.push('}');

        Ok(())
    }
}
