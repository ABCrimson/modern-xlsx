//! The SAX helpers for both worksheet read paths, and their entry points.
//! Both run the shared driver (`driver::drive`): `WorksheetXml::parse`,
//! `parse_with_sst` and `parse_with_sst_within` with the struct sink
//! (`sink_struct::StructSink`), which builds the typed model, and
//! `parse_to_json_within` with the JSON sink (`sink_json::JsonSink`), which
//! writes rows and cells straight to a JSON buffer. The `_within` variants
//! count the text they copy from the SST and the sheet's relationships in a
//! read's `ExpansionBudget`. The driver and the sinks share the helpers below.

use core::hint::cold_path;

use quick_xml::events::BytesStart;
use quick_xml::events::attributes::Attribute;

use super::driver::{SheetMeta, drive};
use super::sink_json::JsonSink;
use super::sink_struct::StructSink;
use super::{
    CellType, Cfvo, ColumnInfo, ConditionalFormattingRule, DataValidation, FrozenPane,
    HeaderFooter, HfPart, Hyperlink, OutlineProperties, PageBreak, PageMargins, PageSetup,
    PaneSelection, ParseState, SheetContext, SheetProtection, SheetViewData, SparklineGroup,
    SplitPane, WorksheetXml,
};
use crate::dates::{DateSystem, iso8601_to_serial};
use crate::ooxml::color::{Color, split_color};
use crate::ooxml::relationships::TargetMode;
use crate::ooxml::shared_strings::{
    ExpansionBudget, RichTextRun, SharedStringTable, runs_json_floor,
};
use crate::ooxml::xml::{XsdBuf, attr_str, attrs, fmt_xsd_double, parse_attr};
use crate::{ModernXlsxError, Result};

// ---------------------------------------------------------------------------
// Helper functions shared by the driver and the sinks
// ---------------------------------------------------------------------------
//
// String attributes are decoded with `attr_str` (entities and character
// references resolved once), numeric attributes with `parse_attr` (a present
// but malformed value is `XmlParse`; an absent one keeps its default), and
// xsd:boolean attributes with `Attribute::as_bool` (`1`/`true`/`0`/`false`;
// any other value keeps the schema default). Attribute lists are walked with
// `attrs` (a malformed list is `XmlParse`).

/// An SST entry as [`resolve_sst`] returns it: its plain text and, for a rich
/// `<si>`, its runs.
pub(super) type SstEntry<'s> = (&'s String, Option<&'s Vec<RichTextRun>>);

/// Resolve a shared-string cell's `<v>` text to its SST entry: the plain text
/// and, for a rich `<si>`, its runs. Only the index is trimmed (xsd:int
/// collapses white space); `None` when there is no SST or no such entry.
///
/// The caller copies what this returns into the cell, so the entry's bytes
/// (and, for a rich entry, [`runs_json_floor`] of its runs) are counted in
/// `budget` first.
///
/// # Errors
///
/// `SizeLimit` when the count passes the budget.
#[inline]
pub(super) fn resolve_sst<'s>(
    sst: Option<&'s SharedStringTable>,
    v: &str,
    budget: &mut ExpansionBudget,
) -> Result<Option<SstEntry<'s>>> {
    let Some(sst) = sst else { return Ok(None) };
    let Ok(idx) = v.trim_ascii().parse::<usize>() else {
        return Ok(None);
    };
    let Some(text) = sst.strings.get(idx) else {
        return Ok(None);
    };
    budget.spend(text.len())?;
    let runs = sst.rich_runs.get(idx).and_then(Option::as_ref);
    if let Some(runs) = runs {
        budget.spend(runs_json_floor(runs))?;
    }
    Ok(Some((text, runs)))
}

/// Parse the `t` attribute value into a `CellType`. `n` and anything unknown
/// are numbers, and so is `d` (an ISO 8601 date) until its `<v>` is converted
/// by [`iso_date_cell`].
#[inline]
fn parse_cell_type_attr(val: &str) -> CellType {
    match val {
        "s" => CellType::SharedString,
        "b" => CellType::Boolean,
        "e" => CellType::Error,
        "str" => CellType::FormulaStr,
        "inlineStr" => CellType::InlineStr,
        _ => CellType::Number,
    }
}

/// Return value for row attribute parsing.
#[derive(Clone, Copy, Default)]
pub(super) struct RowAttrs {
    /// `r`, the 1-based row number: 0 here when `r` is absent, until the
    /// driver numbers the row after the previous one.
    pub(super) index: u32,
    /// Whether `r` was present. Internal to the driver: `RowData` keeps only
    /// the index.
    pub(super) has_r: bool,
    pub(super) height: Option<f64>,
    pub(super) custom_height: bool,
    pub(super) hidden: bool,
    pub(super) outline_level: Option<u8>,
    pub(super) collapsed: bool,
    /// `s`, kept only when `customFormat` is true.
    pub(super) style_index: Option<u32>,
}

/// Parse `<row>` attributes into structured data. The `s` style index counts
/// only when `customFormat` is true (ECMA-376 Part 1, §18.3.1.73).
#[inline]
pub(super) fn parse_row_attrs(e: &BytesStart<'_>) -> Result<RowAttrs> {
    let mut ra = RowAttrs::default();
    let mut style = None;
    let mut custom_format = false;
    for attr in attrs(e) {
        let attr = attr?;
        match attr.key.local_name().as_ref() {
            "r" => {
                ra.index = parse_attr("row", &attr)?;
                ra.has_r = true;
            }
            "s" => style = Some(parse_attr::<u32>("row", &attr)?),
            "customFormat" => {
                if let Some(b) = attr.as_bool() {
                    custom_format = b;
                }
            }
            "ht" => ra.height = Some(parse_attr("row", &attr)?),
            "customHeight" => {
                if let Some(b) = attr.as_bool() {
                    ra.custom_height = b;
                }
            }
            "hidden" => {
                if let Some(b) = attr.as_bool() {
                    ra.hidden = b;
                }
            }
            "outlineLevel" => {
                ra.outline_level = Some(parse_attr::<u8>("row", &attr)?).filter(|&v| v > 0);
            }
            "collapsed" => {
                if let Some(b) = attr.as_bool() {
                    ra.collapsed = b;
                }
            }
            _ => {}
        }
    }
    ra.style_index = style.filter(|_| custom_format);
    Ok(ra)
}

/// Parse a `<col>` element into a [`ColumnInfo`]. Absent attributes keep the
/// defaults (`min`/`max` 1, `width` 8.43, flags false, no outline, no style);
/// a present but malformed number is `XmlParse`. An explicit `style="0"` is
/// kept as `Some(0)`, like a cell's `s="0"`.
pub(super) fn parse_col_element(e: &BytesStart<'_>) -> Result<ColumnInfo> {
    let mut col = ColumnInfo {
        min: 1,
        max: 1,
        width: 8.43,
        hidden: false,
        custom_width: false,
        outline_level: None,
        collapsed: false,
        style_index: None,
        best_fit: false,
        phonetic: false,
    };
    for attr in attrs(e) {
        let attr = attr?;
        match attr.key.local_name().as_ref() {
            "min" => col.min = parse_attr("col", &attr)?,
            "max" => col.max = parse_attr("col", &attr)?,
            "width" => col.width = parse_attr("col", &attr)?,
            "style" => col.style_index = Some(parse_attr::<u32>("col", &attr)?),
            "hidden" => {
                if let Some(b) = attr.as_bool() {
                    col.hidden = b;
                }
            }
            "bestFit" => {
                if let Some(b) = attr.as_bool() {
                    col.best_fit = b;
                }
            }
            "customWidth" => {
                if let Some(b) = attr.as_bool() {
                    col.custom_width = b;
                }
            }
            "phonetic" => {
                if let Some(b) = attr.as_bool() {
                    col.phonetic = b;
                }
            }
            "outlineLevel" => {
                col.outline_level = Some(parse_attr::<u8>("col", &attr)?).filter(|&v| v > 0);
            }
            "collapsed" => {
                if let Some(b) = attr.as_bool() {
                    col.collapsed = b;
                }
            }
            _ => {}
        }
    }
    Ok(col)
}

/// Cell attributes filled in place by [`parse_cell_attrs_into`].
#[derive(Default)]
pub(super) struct CellAttrs {
    /// `r`: empty here when the attribute is absent, until the driver fills
    /// in the next position in the row.
    pub(super) reference: String,
    pub(super) cell_type: CellType,
    /// `t="d"`: the `<v>` holds an ISO 8601 date.
    pub(super) iso_date: bool,
    pub(super) style_index: Option<u32>,
    /// `cm`: the cell-metadata index.
    pub(super) cell_metadata: Option<u32>,
}

/// Parse `<c>` attributes (reference, type, style index, cell metadata) into
/// `ca`. Every field is reset first, so one `CellAttrs` serves every cell and
/// its reference buffer is reused. Attributes match by local name and quick-xml
/// rejects only an identical QName, so `r` and a prefixed `x:r` can both
/// appear: the last one wins, as for every other attribute.
#[inline]
pub(super) fn parse_cell_attrs_into(e: &BytesStart<'_>, ca: &mut CellAttrs) -> Result<()> {
    ca.reference.clear();
    ca.cell_type = CellType::default();
    ca.iso_date = false;
    ca.style_index = None;
    ca.cell_metadata = None;
    for attr in attrs(e) {
        let attr = attr?;
        match attr.key.local_name().as_ref() {
            "r" => {
                ca.reference.clear();
                ca.reference.push_str(&attr_str(&attr)?);
            }
            "t" => {
                let t = attr_str(&attr)?;
                ca.iso_date = t == "d";
                ca.cell_type = parse_cell_type_attr(&t);
            }
            "s" => ca.style_index = Some(parse_attr("c", &attr)?),
            "cm" => ca.cell_metadata = Some(parse_attr("c", &attr)?),
            _ => {}
        }
    }
    Ok(())
}

/// The type and value of a `t="d"` cell from its `<v>` text: a number cell
/// holding the serial in `date_system`, or, when the text is not an ISO 8601
/// value that date system holds, a `t="str"` cell keeping the text.
pub(super) fn iso_date_cell(
    text: String,
    date_system: DateSystem,
    xsd: &mut XsdBuf,
) -> Result<(CellType, String)> {
    Ok(match iso8601_to_serial(&text, date_system) {
        Some(serial) => (
            CellType::Number,
            fmt_xsd_double(serial, xsd, "t=\"d\" cell serial")?.to_owned(),
        ),
        None => (CellType::FormulaStr, text),
    })
}

/// The elements of `list`, moved out and shrunk to their exact length,
/// leaving `list` empty without a buffer. A child list moved as it grew would
/// keep the spare capacity of its growth (at least four slots), so a flood of
/// one-child elements would hold four times what it stores.
pub(super) fn take_exact<T>(list: &mut Vec<T>) -> Vec<T> {
    let mut out = std::mem::take(list);
    out.shrink_to_fit();
    out
}

/// The colour-scale stop colours to keep beside the rgb strings: every stop's
/// full colour when any stop holds more than an `rgb`, else none, moved out
/// by [`take_exact`].
pub(super) fn stop_color_specs(stops: &mut Vec<Color>) -> Vec<Color> {
    if stops.iter().any(Color::has_non_rgb) {
        take_exact(stops)
    } else {
        stops.clear();
        Vec::new()
    }
}

/// The verbatim `<ext>` element at `data[start..end]`, without the white space
/// before it, borrowed from `data`; the caller inserts the namespace
/// declarations it inherits (`xml::inherited_declarations`).
///
/// # Errors
///
/// `XmlParse` when the span is not UTF-8.
pub(super) fn captured_ext(data: &[u8], start: u64, end: u64) -> Result<&str> {
    let span = data.get(start as usize..end as usize).unwrap_or_default();
    let text = std::str::from_utf8(span).map_err(|e| {
        cold_path();
        ModernXlsxError::XmlParse(format!("worksheet <ext> is not UTF-8: {e}"))
    })?;
    Ok(text.trim_start())
}

/// The formula attributes of a cell's `<f>`, plus the `<c cm>` cell-metadata
/// index, which travels with them because both set `dynamic_array`.
#[derive(Default)]
pub(super) struct FormulaState {
    pub(super) formula_type: Option<String>,
    pub(super) formula_ref: Option<String>,
    pub(super) shared_index: Option<u32>,
    pub(super) cell_metadata: Option<u32>,
    pub(super) dynamic_array: Option<bool>,
    pub(super) formula_r1: Option<String>,
    pub(super) formula_r2: Option<String>,
    pub(super) formula_dt2d: Option<bool>,
    pub(super) formula_dtr: Option<bool>,
    pub(super) formula_del1: Option<bool>,
    pub(super) formula_del2: Option<bool>,
    pub(super) formula_dtr1: Option<bool>,
    pub(super) formula_dtr2: Option<bool>,
}

impl FormulaState {
    /// Record the `<c cm>` index; any cell metadata marks a dynamic array.
    pub(super) fn set_cell_metadata(&mut self, cm: Option<u32>) {
        if cm.is_some() {
            self.cell_metadata = cm;
            self.dynamic_array = Some(true);
        }
    }

    /// Take every field that `from` holds into `self`, replacing the field
    /// there. `parse_formula_attrs` only ever sets a field, so merging the
    /// attributes of one `<f>` (parsed from defaults) into `self` leaves
    /// `self` as parsing that `<f>` straight into it would.
    pub(super) fn merge_from(&mut self, from: &mut Self) {
        fn take_set<T>(into: &mut Option<T>, from: &mut Option<T>) {
            if from.is_some() {
                *into = from.take();
            }
        }
        // Exhaustive, so a new field cannot be left out.
        let Self {
            formula_type,
            formula_ref,
            shared_index,
            cell_metadata,
            dynamic_array,
            formula_r1,
            formula_r2,
            formula_dt2d,
            formula_dtr,
            formula_del1,
            formula_del2,
            formula_dtr1,
            formula_dtr2,
        } = from;
        take_set(&mut self.formula_type, formula_type);
        take_set(&mut self.formula_ref, formula_ref);
        take_set(&mut self.shared_index, shared_index);
        take_set(&mut self.cell_metadata, cell_metadata);
        take_set(&mut self.dynamic_array, dynamic_array);
        take_set(&mut self.formula_r1, formula_r1);
        take_set(&mut self.formula_r2, formula_r2);
        take_set(&mut self.formula_dt2d, formula_dt2d);
        take_set(&mut self.formula_dtr, formula_dtr);
        take_set(&mut self.formula_del1, formula_del1);
        take_set(&mut self.formula_del2, formula_del2);
        take_set(&mut self.formula_dtr1, formula_dtr1);
        take_set(&mut self.formula_dtr2, formula_dtr2);
    }
}

/// Parse `<f>` (formula) element attributes into the given state fields.
///
/// The `CT_CellFormula` flags `dt2D`, `dtr`, `del1` and `del2` are kept when
/// true; `del1`/`del2` also set the 1.x `formula_dtr1`/`formula_dtr2`. The
/// 1.2.x writer put `cm`, `dtr1` and `dtr2` on `<f>`, outside the schema;
/// they are still read: `cm="1"` into `dynamic_array`, `dtr1` and `dtr2`
/// into the 1.x fields.
#[inline]
pub(super) fn parse_formula_attrs(e: &BytesStart<'_>, fs: &mut FormulaState) -> Result<()> {
    for attr in attrs(e) {
        let attr = attr?;
        let on = || attr.as_bool() == Some(true);
        match attr.key.local_name().as_ref() {
            "t" => fs.formula_type = Some(attr_str(&attr)?.into_owned()),
            "ref" => fs.formula_ref = Some(attr_str(&attr)?.into_owned()),
            "si" => fs.shared_index = Some(parse_attr("f", &attr)?),
            "r1" => fs.formula_r1 = Some(attr_str(&attr)?.into_owned()),
            "r2" => fs.formula_r2 = Some(attr_str(&attr)?.into_owned()),
            "dt2D" if on() => fs.formula_dt2d = Some(true),
            "dtr" if on() => fs.formula_dtr = Some(true),
            "del1" if on() => (fs.formula_del1, fs.formula_dtr1) = (Some(true), Some(true)),
            "del2" if on() => (fs.formula_del2, fs.formula_dtr2) = (Some(true), Some(true)),
            "cm" => {
                if parse_attr::<u32>("f", &attr)? == 1 {
                    fs.dynamic_array = Some(true);
                }
            }
            "dtr1" if on() => fs.formula_dtr1 = Some(true),
            "dtr2" if on() => fs.formula_dtr2 = Some(true),
            _ => {}
        }
    }
    Ok(())
}

/// Parse `<sheetView>` attributes. Returns `(SheetViewData, has_non_default)`,
/// where `has_non_default` is true when any value differs from its schema
/// default.
pub(super) fn parse_sheet_view_attrs(e: &BytesStart<'_>) -> Result<(SheetViewData, bool)> {
    let mut sv = SheetViewData::default();
    for attr in attrs(e) {
        let attr = attr?;
        let flag = match attr.key.local_name().as_ref() {
            "showGridLines" => &mut sv.show_grid_lines,
            "showRowColHeaders" => &mut sv.show_row_col_headers,
            "showZeros" => &mut sv.show_zeros,
            "rightToLeft" => &mut sv.right_to_left,
            "tabSelected" => &mut sv.tab_selected,
            "showRuler" => &mut sv.show_ruler,
            "showOutlineSymbols" => &mut sv.show_outline_symbols,
            "showWhiteSpace" => &mut sv.show_white_space,
            "defaultGridColor" => &mut sv.default_grid_color,
            "zoomScale" => {
                sv.zoom_scale = Some(parse_attr("sheetView", &attr)?);
                continue;
            }
            "zoomScaleNormal" => {
                sv.zoom_scale_normal = Some(parse_attr("sheetView", &attr)?);
                continue;
            }
            "zoomScalePageLayoutView" => {
                sv.zoom_scale_page_layout_view = Some(parse_attr("sheetView", &attr)?);
                continue;
            }
            "zoomScaleSheetLayoutView" => {
                sv.zoom_scale_sheet_layout_view = Some(parse_attr("sheetView", &attr)?);
                continue;
            }
            "colorId" => {
                sv.color_id = Some(parse_attr("sheetView", &attr)?);
                continue;
            }
            "view" => {
                let view = attr_str(&attr)?;
                if view != "normal" {
                    sv.view = Some(view.into_owned());
                }
                continue;
            }
            _ => continue,
        };
        if let Some(b) = attr.as_bool() {
            *flag = b;
        }
    }
    let has_non_default = sv != SheetViewData::default();
    Ok((sv, has_non_default))
}

/// Parse `<selection>` attributes into a `PaneSelection`.
#[inline]
pub(super) fn parse_selection_attrs(e: &BytesStart<'_>) -> Result<PaneSelection> {
    let mut sel = PaneSelection {
        pane: None,
        active_cell: None,
        sqref: None,
    };
    for attr in attrs(e) {
        let attr = attr?;
        let slot = match attr.key.local_name().as_ref() {
            "pane" => &mut sel.pane,
            "activeCell" => &mut sel.active_cell,
            "sqref" => &mut sel.sqref,
            _ => continue,
        };
        *slot = Some(attr_str(&attr)?.into_owned());
    }
    Ok(sel)
}

/// Parse `<pane>` attributes into frozen pane or split pane.
pub(super) fn parse_pane_attrs(
    e: &BytesStart<'_>,
) -> Result<(Option<FrozenPane>, Option<SplitPane>)> {
    let mut y_split = None::<Attribute<'_>>;
    let mut x_split = None::<Attribute<'_>>;
    let mut frozen = false;
    let mut top_left_cell_val = None::<String>;
    let mut active_pane_val = None::<String>;

    for attr in attrs(e) {
        let attr = attr?;
        match attr.key.local_name().as_ref() {
            "ySplit" => y_split = Some(attr),
            "xSplit" => x_split = Some(attr),
            "state" => frozen = attr_str(&attr)? == "frozen",
            "topLeftCell" => top_left_cell_val = Some(attr_str(&attr)?.into_owned()),
            "activePane" => active_pane_val = Some(attr_str(&attr)?.into_owned()),
            _ => {}
        }
    }

    if frozen {
        let y: u32 = y_split
            .map(|a| parse_attr("pane", &a))
            .transpose()?
            .unwrap_or(0);
        let x: u32 = x_split
            .map(|a| parse_attr("pane", &a))
            .transpose()?
            .unwrap_or(0);
        if y > 0 || x > 0 {
            return Ok((Some(FrozenPane { rows: y, cols: x }), None));
        }
    } else {
        let y: f64 = y_split
            .map(|a| parse_attr("pane", &a))
            .transpose()?
            .unwrap_or(0.0);
        let x: f64 = x_split
            .map(|a| parse_attr("pane", &a))
            .transpose()?
            .unwrap_or(0.0);
        if y > 0.0 || x > 0.0 {
            return Ok((
                None,
                Some(SplitPane {
                    horizontal: (y > 0.0).then_some(y),
                    vertical: (x > 0.0).then_some(x),
                    top_left_cell: top_left_cell_val,
                    active_pane: active_pane_val,
                }),
            ));
        }
    }
    Ok((None, None))
}

/// Parse `<dataValidation>` attributes into a `DataValidation`.
pub(super) fn parse_data_validation_attrs(e: &BytesStart<'_>) -> Result<DataValidation> {
    let mut dv = DataValidation {
        sqref: String::new(),
        validation_type: None,
        operator: None,
        error_style: None,
        ime_mode: None,
        show_drop_down: None,
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
    for attr in attrs(e) {
        let attr = attr?;
        match attr.key.local_name().as_ref() {
            "sqref" => dv.sqref = attr_str(&attr)?.into_owned(),
            "type" => dv.validation_type = Some(attr_str(&attr)?.into_owned()),
            "operator" => dv.operator = Some(attr_str(&attr)?.into_owned()),
            "errorStyle" => dv.error_style = Some(attr_str(&attr)?.into_owned()),
            "imeMode" => dv.ime_mode = Some(attr_str(&attr)?.into_owned()),
            "showDropDown" => dv.show_drop_down = attr.as_bool(),
            "allowBlank" => dv.allow_blank = attr.as_bool(),
            "showErrorMessage" => dv.show_error_message = attr.as_bool(),
            "errorTitle" => dv.error_title = Some(attr_str(&attr)?.into_owned()),
            "error" => dv.error_message = Some(attr_str(&attr)?.into_owned()),
            "showInputMessage" => dv.show_input_message = attr.as_bool(),
            "promptTitle" => dv.prompt_title = Some(attr_str(&attr)?.into_owned()),
            "prompt" => dv.prompt = Some(attr_str(&attr)?.into_owned()),
            _ => {}
        }
    }
    Ok(dv)
}

/// Parse `<cfRule>` attributes into a `ConditionalFormattingRule`.
#[inline]
pub(super) fn parse_cf_rule_attrs(e: &BytesStart<'_>) -> Result<ConditionalFormattingRule> {
    let mut rule = ConditionalFormattingRule {
        rule_type: String::new(),
        priority: 0,
        operator: None,
        formula: None,
        additional_formulas: Vec::new(),
        dxf_id: None,
        stop_if_true: None,
        above_average: None,
        percent: None,
        bottom: None,
        text: None,
        time_period: None,
        rank: None,
        std_dev: None,
        equal_average: None,
        color_scale: None,
        data_bar: None,
        icon_set: None,
    };
    for attr in attrs(e) {
        let attr = attr?;
        match attr.key.local_name().as_ref() {
            "type" => rule.rule_type = attr_str(&attr)?.into_owned(),
            "priority" => rule.priority = parse_attr("cfRule", &attr)?,
            "operator" => rule.operator = Some(attr_str(&attr)?.into_owned()),
            "dxfId" => rule.dxf_id = Some(parse_attr("cfRule", &attr)?),
            "stopIfTrue" => rule.stop_if_true = attr.as_bool(),
            "aboveAverage" => rule.above_average = attr.as_bool(),
            "percent" => rule.percent = attr.as_bool(),
            "bottom" => rule.bottom = attr.as_bool(),
            "equalAverage" => rule.equal_average = attr.as_bool(),
            "rank" => rule.rank = Some(parse_attr("cfRule", &attr)?),
            "stdDev" => rule.std_dev = Some(parse_attr("cfRule", &attr)?),
            "text" => rule.text = Some(attr_str(&attr)?.into_owned()),
            "timePeriod" => rule.time_period = Some(attr_str(&attr)?.into_owned()),
            _ => {}
        }
    }
    Ok(rule)
}

/// Parse `<hyperlink>` attributes into a `Hyperlink`.
///
/// `r:id` (an `id` whose prefix is in `rel_prefixes`, the prefixes the
/// `<worksheet>` declarations bind to the relationships namespace;
/// CT_Hyperlink has no unprefixed one) becomes `target` when it names an
/// external relationship in `ctx.rels`; any other `r:id`, and an `id` under
/// any other prefix, is ignored. `rel_prefixes` is sorted and deduplicated,
/// so the prefix lookup is one binary search: a part that declares many
/// prefixes costs a search logarithmic in them per hyperlink, and no hasher
/// lets prefixes the file chooses collide.
///
/// One relationship may be named by any number of hyperlinks, each getting a
/// copy of its target, so the target the hyperlink keeps is counted in
/// `budget`.
///
/// # Errors
///
/// `XmlParse` for a malformed attribute; `SizeLimit` when the count passes
/// the budget.
#[inline]
pub(super) fn parse_hyperlink_attrs(
    e: &BytesStart<'_>,
    ctx: &SheetContext<'_>,
    rel_prefixes: &[String],
    budget: &mut ExpansionBudget,
) -> Result<Hyperlink> {
    let mut hl = Hyperlink {
        cell_ref: String::new(),
        target: None,
        location: None,
        display: None,
        tooltip: None,
    };
    for attr in attrs(e) {
        let attr = attr?;
        match attr.key.local_name().as_ref() {
            "ref" => hl.cell_ref = attr_str(&attr)?.into_owned(),
            "id" if attr.key.prefix().is_some_and(|p| {
                rel_prefixes
                    .binary_search_by(|q| q.as_str().cmp(p.as_ref()))
                    .is_ok()
            }) =>
            {
                let r_id = attr_str(&attr)?;
                hl.target = ctx
                    .rels
                    .and_then(|rels| rels.get_by_id(&r_id))
                    .filter(|rel| rel.target_mode == TargetMode::External)
                    .map(|rel| rel.target.clone().into_owned());
            }
            "location" => hl.location = Some(attr_str(&attr)?.into_owned()),
            "display" => hl.display = Some(attr_str(&attr)?.into_owned()),
            "tooltip" => hl.tooltip = Some(attr_str(&attr)?.into_owned()),
            _ => {}
        }
    }
    // Counted once, after the loop: a second `r:id` replaces the first copy.
    if let Some(target) = &hl.target {
        budget.spend(target.len())?;
    }
    Ok(hl)
}

/// Merge `<pageSetup>` attributes into `ps`.
///
/// `<pageMargins>` precedes `<pageSetup>` in the file and lands in the same
/// `PageSetup`, so the attributes are merged, never replaced.
#[inline]
pub(super) fn merge_page_setup_attrs(e: &BytesStart<'_>, ps: &mut PageSetup) -> Result<()> {
    for attr in attrs(e) {
        let attr = attr?;
        let slot = match attr.key.local_name().as_ref() {
            "paperSize" => &mut ps.paper_size,
            "orientation" => {
                ps.orientation = Some(attr_str(&attr)?.into_owned());
                continue;
            }
            "fitToWidth" => &mut ps.fit_to_width,
            "fitToHeight" => &mut ps.fit_to_height,
            "scale" => &mut ps.scale,
            "firstPageNumber" => &mut ps.first_page_number,
            "horizontalDpi" => &mut ps.horizontal_dpi,
            "verticalDpi" => &mut ps.vertical_dpi,
            _ => continue,
        };
        *slot = Some(parse_attr("pageSetup", &attr)?);
    }
    Ok(())
}

/// Parse `<pageMargins>` attributes into a `PageMargins`.
#[inline]
pub(super) fn parse_page_margins_attrs(e: &BytesStart<'_>) -> Result<PageMargins> {
    let mut pm = PageMargins::default();
    for attr in attrs(e) {
        let attr = attr?;
        let slot = match attr.key.local_name().as_ref() {
            "top" => &mut pm.top,
            "bottom" => &mut pm.bottom,
            "left" => &mut pm.left,
            "right" => &mut pm.right,
            "header" => &mut pm.header,
            "footer" => &mut pm.footer,
            _ => continue,
        };
        *slot = Some(parse_attr("pageMargins", &attr)?);
    }
    Ok(pm)
}

/// Parse `<sheetProtection>` attributes into a `SheetProtection`, starting from
/// the schema default of every attribute (`SheetProtection::default`).
pub(super) fn parse_sheet_protection_attrs(e: &BytesStart<'_>) -> Result<SheetProtection> {
    let mut sp = SheetProtection::default();
    for attr in attrs(e) {
        let attr = attr?;
        let flag = match attr.key.local_name().as_ref() {
            "password" => {
                sp.password = Some(attr_str(&attr)?.into_owned());
                continue;
            }
            "algorithmName" => {
                sp.algorithm_name = Some(attr_str(&attr)?.into_owned());
                continue;
            }
            "hashValue" => {
                sp.hash_value = Some(attr_str(&attr)?.into_owned());
                continue;
            }
            "saltValue" => {
                sp.salt_value = Some(attr_str(&attr)?.into_owned());
                continue;
            }
            "spinCount" => {
                sp.spin_count = Some(parse_attr("sheetProtection", &attr)?);
                continue;
            }
            "sheet" => &mut sp.sheet,
            "objects" => &mut sp.objects,
            "scenarios" => &mut sp.scenarios,
            "formatCells" => &mut sp.format_cells,
            "formatColumns" => &mut sp.format_columns,
            "formatRows" => &mut sp.format_rows,
            "insertColumns" => &mut sp.insert_columns,
            "insertRows" => &mut sp.insert_rows,
            "insertHyperlinks" => &mut sp.insert_hyperlinks,
            "deleteColumns" => &mut sp.delete_columns,
            "deleteRows" => &mut sp.delete_rows,
            "selectLockedCells" => &mut sp.select_locked_cells,
            "sort" => &mut sp.sort,
            "autoFilter" => &mut sp.auto_filter,
            "pivotTables" => &mut sp.pivot_tables,
            "selectUnlockedCells" => &mut sp.select_unlocked_cells,
            _ => continue,
        };
        if let Some(b) = attr.as_bool() {
            *flag = b;
        }
    }
    Ok(sp)
}

/// Parse `<cfvo>` attributes into a `Cfvo`.
#[inline]
pub(super) fn parse_cfvo_attrs(e: &BytesStart<'_>) -> Result<Cfvo> {
    let mut cfvo = Cfvo {
        cfvo_type: String::new(),
        val: None,
    };
    for attr in attrs(e) {
        let attr = attr?;
        match attr.key.local_name().as_ref() {
            "type" => cfvo.cfvo_type = attr_str(&attr)?.into_owned(),
            "val" => cfvo.val = Some(attr_str(&attr)?.into_owned()),
            _ => {}
        }
    }
    Ok(cfvo)
}

/// Parse `<headerFooter>` attributes into a `HeaderFooter`.
#[inline]
pub(super) fn parse_header_footer_attrs(e: &BytesStart<'_>) -> Result<HeaderFooter> {
    let mut hf = HeaderFooter::default();
    for attr in attrs(e) {
        let attr = attr?;
        let flag = match attr.key.as_ref() {
            "differentOddEven" => &mut hf.different_odd_even,
            "differentFirst" => &mut hf.different_first,
            "scaleWithDoc" => &mut hf.scale_with_doc,
            "alignWithMargins" => &mut hf.align_with_margins,
            _ => continue,
        };
        if let Some(b) = attr.as_bool() {
            *flag = b;
        }
    }
    Ok(hf)
}

/// Parse `<outlinePr>` attributes into an `OutlineProperties`.
#[inline]
pub(super) fn parse_outline_pr_attrs(e: &BytesStart<'_>) -> Result<OutlineProperties> {
    let mut op = OutlineProperties {
        summary_below: true,
        summary_right: true,
    };
    for attr in attrs(e) {
        let attr = attr?;
        let flag = match attr.key.as_ref() {
            "summaryBelow" => &mut op.summary_below,
            "summaryRight" => &mut op.summary_right,
            _ => continue,
        };
        if let Some(b) = attr.as_bool() {
            *flag = b;
        }
    }
    Ok(op)
}

/// Parse `<brk>` (break) attributes into a `PageBreak`.
#[inline]
pub(super) fn parse_brk_attrs(e: &BytesStart<'_>) -> Result<PageBreak> {
    let mut brk = PageBreak {
        id: 0,
        min: None,
        max: None,
        man: false,
    };
    for attr in attrs(e) {
        let attr = attr?;
        match attr.key.local_name().as_ref() {
            "id" => brk.id = parse_attr("brk", &attr)?,
            "min" => brk.min = Some(parse_attr("brk", &attr)?),
            "max" => brk.max = Some(parse_attr("brk", &attr)?),
            "man" => {
                if let Some(b) = attr.as_bool() {
                    brk.man = b;
                }
            }
            _ => {}
        }
    }
    Ok(brk)
}

/// Parse `<sparklineGroup>` attributes into a `SparklineGroup`.
pub(super) fn parse_sparkline_group_attrs(e: &BytesStart<'_>) -> Result<SparklineGroup> {
    let mut group = SparklineGroup::default();
    for attr in attrs(e) {
        let attr = attr?;
        let flag = match attr.key.local_name().as_ref() {
            "type" => {
                group.sparkline_type = attr_str(&attr)?.into_owned();
                continue;
            }
            "displayEmptyCellsAs" => {
                group.display_empty_cells_as = Some(attr_str(&attr)?.into_owned());
                continue;
            }
            "lineWeight" => {
                group.line_weight = Some(parse_attr("sparklineGroup", &attr)?);
                continue;
            }
            "manualMin" => {
                group.manual_min = Some(parse_attr("sparklineGroup", &attr)?);
                continue;
            }
            "manualMax" => {
                group.manual_max = Some(parse_attr("sparklineGroup", &attr)?);
                continue;
            }
            "minAxisType" => {
                group.min_axis_type = Some(attr_str(&attr)?.into_owned());
                continue;
            }
            "maxAxisType" => {
                group.max_axis_type = Some(attr_str(&attr)?.into_owned());
                continue;
            }
            "dateAxis" => &mut group.date_axis,
            "displayHidden" => &mut group.display_hidden,
            "markers" => &mut group.markers,
            "high" => &mut group.high,
            "low" => &mut group.low,
            "first" => &mut group.first,
            "last" => &mut group.last,
            "negative" => &mut group.negative,
            "displayXAxis" => &mut group.display_x_axis,
            "rightToLeft" => &mut group.right_to_left,
            _ => continue,
        };
        if let Some(b) = attr.as_bool() {
            *flag = b;
        }
    }
    Ok(group)
}

/// The first attribute of `e` whose local name is `name`.
#[inline]
pub(super) fn find_attr<'a>(e: &'a BytesStart<'_>, name: &str) -> Result<Option<Attribute<'a>>> {
    for attr in attrs(e) {
        let attr = attr?;
        if attr.key.local_name().as_ref() == name {
            return Ok(Some(attr));
        }
    }
    Ok(None)
}

/// The decoded value of the attribute of `e` whose local name is `name`.
#[inline]
pub(super) fn find_attr_str(e: &BytesStart<'_>, name: &str) -> Result<Option<String>> {
    find_attr(e, name)?
        .map(|a| attr_str(&a).map(std::borrow::Cow::into_owned))
        .transpose()
}

/// Whether `<ext>` element `e` holds sparkline groups: its `uri` is the x14
/// sparklineGroups URI or the one 1.2.x wrote, compared ignoring ASCII case
/// (GUID hex digits are case-insensitive). No `uri` is `Ok(false)`.
pub(super) fn is_sparkline_ext(e: &BytesStart<'_>) -> Result<bool> {
    let Some(a) = find_attr(e, "uri")? else {
        return Ok(false);
    };
    let uri = attr_str(&a)?;
    Ok([super::SPARKLINE_EXT_URI, super::LEGACY_SPARKLINE_EXT_URI]
        .iter()
        .any(|known| uri.eq_ignore_ascii_case(known)))
}

/// Parse a sparkline colour element (self-closing or a start/end pair) from a
/// `BytesStart` into a group: its `rgb` into the legacy string field and, when it has a theme,
/// indexed, tint or auto attribute, the full colour into the `*_spec` field.
#[inline]
pub(super) fn apply_sparkline_color(
    group: &mut SparklineGroup,
    local: &str,
    e: &BytesStart<'_>,
) -> Result<()> {
    let (rgb_slot, spec_slot) = match local {
        "colorSeries" => (&mut group.color_series, &mut group.color_series_spec),
        "colorNegative" => (&mut group.color_negative, &mut group.color_negative_spec),
        "colorAxis" => (&mut group.color_axis, &mut group.color_axis_spec),
        "colorMarkers" => (&mut group.color_markers, &mut group.color_markers_spec),
        "colorFirst" => (&mut group.color_first, &mut group.color_first_spec),
        "colorLast" => (&mut group.color_last, &mut group.color_last_spec),
        "colorHigh" => (&mut group.color_high, &mut group.color_high_spec),
        "colorLow" => (&mut group.color_low, &mut group.color_low_spec),
        _ => return Ok(()),
    };
    (*rgb_slot, *spec_slot) = split_color(Color::from_attrs(e)?);
    Ok(())
}

/// The buffer that character data (text, CDATA and entity references) in
/// `state` accumulates into, or `None` for a structural state, whose text is
/// only inter-element whitespace and is dropped. Under `<is>` only `<t>` runs
/// collect text; `<rPh>` phonetic text and the whitespace between runs do not.
#[inline]
pub(super) fn text_sink<'b>(
    state: ParseState,
    text: &'b mut String,
    header_footer: &'b mut String,
    sparkline_formula: &'b mut String,
    sparkline_sqref: &'b mut String,
    sparkline_dates: &'b mut String,
) -> Option<&'b mut String> {
    match state {
        ParseState::InCellValue
        | ParseState::InCellFormula
        | ParseState::InInlineStrT
        | ParseState::InDVFormula1
        | ParseState::InDVFormula2
        | ParseState::InCfRuleFormula => Some(text),
        ParseState::HeaderFooterChild(_) => Some(header_footer),
        ParseState::SparklineFormula => Some(sparkline_formula),
        ParseState::SparklineSqref => Some(sparkline_sqref),
        ParseState::SparklineDateRange => Some(sparkline_dates),
        _ => None,
    }
}

/// Assign a header/footer text field by child part.
#[inline]
pub(super) fn assign_hf_field(hf: &mut HeaderFooter, part: HfPart, text: Option<String>) {
    match part {
        HfPart::OddHeader => hf.odd_header = text,
        HfPart::OddFooter => hf.odd_footer = text,
        HfPart::EvenHeader => hf.even_header = text,
        HfPart::EvenFooter => hf.even_footer = text,
        HfPart::FirstHeader => hf.first_header = text,
        HfPart::FirstFooter => hf.first_footer = text,
    }
}

// ---------------------------------------------------------------------------
// Main parser implementations
// ---------------------------------------------------------------------------

impl WorksheetXml {
    /// Parse a worksheet XML file from raw bytes.
    /// `t="d"` cells are read in the 1900 date system.
    pub fn parse(data: &[u8]) -> Result<Self> {
        Self::parse_with_sst(data, None, &SheetContext::default())
    }

    /// Parse a worksheet XML file, optionally resolving shared string indices
    /// inline during parsing: a SharedString cell gets its SST entry's text as
    /// `value` and, for a rich entry, that entry's runs as `rich_text`, so
    /// there is no second pass over the cells and no index string is kept
    /// for a SharedString cell. A `t="d"` cell becomes
    /// a number cell holding its serial in `ctx.date_system`. The text copied
    /// from the SST is not counted ([`Self::parse_with_sst_within`] counts it).
    pub(crate) fn parse_with_sst(
        data: &[u8],
        sst: Option<&SharedStringTable>,
        ctx: &SheetContext<'_>,
    ) -> Result<Self> {
        Self::parse_with_sst_within(data, sst, ctx, &mut ExpansionBudget::new(usize::MAX))
    }

    /// [`Self::parse_with_sst`], counting in `budget` the text it copies into
    /// cells (SST entries) and hyperlinks (relationship targets), and the
    /// worksheet namespace declarations re-declared on each kept `<ext>`.
    ///
    /// # Errors
    ///
    /// As [`Self::parse_with_sst`], and `SizeLimit` when the count passes the
    /// budget.
    pub(crate) fn parse_with_sst_within(
        data: &[u8],
        sst: Option<&SharedStringTable>,
        ctx: &SheetContext<'_>,
        budget: &mut ExpansionBudget,
    ) -> Result<Self> {
        // Pre-allocate buffers proportional to input size for fewer reallocations.
        let estimated_rows = (data.len() / 200).clamp(64, 1 << 16); // one row per 200 bytes, clamped
        let mut sink = StructSink::new(sst, ctx.date_system, estimated_rows);
        let SheetMeta {
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
        } = drive(data, ctx, &mut sink, budget)?;
        let rows = sink.into_rows();

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
            tab_color_spec,
            tables: Vec::new(),
            header_footer,
            page_breaks,
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
    /// The alternative to `parse_with_sst()` followed by `serde_json::to_string()`
    /// that `reader::read_xlsx_json` uses: the driver runs with a `JsonSink`,
    /// which writes each row and cell to `out` as it is read, with no `Row` or
    /// `Cell` built. The metadata (merge cells, columns, hyperlinks, ...) is
    /// collected into its structs by the driver and written after the rows,
    /// mostly through `serde_json`. The
    /// `read_json_100k_rows` group in `benches/read_write.rs` times the two
    /// read paths (`read_xlsx` plus `serde_json::to_string` against
    /// `read_xlsx_json`) on a 100,000-row workbook.
    ///
    /// `budget` counts the text copied into cells (SST entries) and
    /// hyperlinks (relationship targets), and the worksheet namespace
    /// declarations re-declared on each kept `<ext>`, and its limit is the
    /// `max_json_length` that `out` is held to after each shared string a
    /// `<v>` names and each cell, row, comment, hyperlink and kept extension.
    ///
    /// # Errors
    ///
    /// As [`Self::parse_with_sst`]; `SizeLimit` when the count passes the
    /// budget or `out` grows past its limit; `Json` when serde_json fails.
    pub(crate) fn parse_to_json_within(
        data: &[u8],
        sst: Option<&SharedStringTable>,
        comments: &[crate::ooxml::comments::Comment],
        tables: &[crate::ooxml::tables::TableDefinition],
        ctx: &SheetContext<'_>,
        out: &mut String,
        budget: &mut ExpansionBudget,
    ) -> Result<()> {
        let mut sink = JsonSink::new(out, sst, ctx.date_system, budget.limit);
        let meta = drive(data, ctx, &mut sink, budget)?;
        sink.finish(&meta, comments, tables)
    }

    /// [`Self::parse_to_json_within`] without a limit, for the tests that
    /// compare the two read paths.
    #[cfg(test)]
    pub(crate) fn parse_to_json(
        data: &[u8],
        sst: Option<&SharedStringTable>,
        comments: &[crate::ooxml::comments::Comment],
        tables: &[crate::ooxml::tables::TableDefinition],
        ctx: &SheetContext<'_>,
        out: &mut String,
    ) -> Result<()> {
        Self::parse_to_json_within(
            data,
            sst,
            comments,
            tables,
            ctx,
            out,
            &mut ExpansionBudget::new(usize::MAX),
        )
    }
}

#[cfg(test)]
mod tests {
    use super::WorksheetXml;
    use pretty_assertions::assert_eq;

    #[test]
    fn formula_fields_reset_between_cells() {
        let xml = br#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1"><f t="shared" ref="A1:A2" si="0">1+1</f><v>2</v></c><c r="B1"><v>3</v></c></row></sheetData></worksheet>"#;
        let ws = WorksheetXml::parse(xml).unwrap();
        let cells = &ws.rows[0].cells;
        assert_eq!(cells[0].formula_type.as_deref(), Some("shared"));
        assert_eq!(cells[0].formula_ref.as_deref(), Some("A1:A2"));
        assert_eq!(cells[0].shared_index, Some(0));
        assert_eq!(cells[1].formula_type, None);
        assert_eq!(cells[1].formula_ref, None);
        assert_eq!(cells[1].shared_index, None);
    }

    /// Parse `<cols>` through both parsers; the JSON path is checked against
    /// the model path's columns.
    fn parse_cols(col: &str) -> crate::Result<Vec<crate::ooxml::worksheet::ColumnInfo>> {
        let xml = format!(
            r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><cols>{col}</cols><sheetData/></worksheet>"#
        );
        let model = WorksheetXml::parse_with_sst(
            xml.as_bytes(),
            None,
            &crate::ooxml::worksheet::SheetContext::default(),
        );
        let mut json = String::new();
        let streamed = WorksheetXml::parse_to_json(
            xml.as_bytes(),
            None,
            &[],
            &[],
            &crate::ooxml::worksheet::SheetContext::default(),
            &mut json,
        );
        assert_eq!(model.is_ok(), streamed.is_ok(), "{col}");
        match (model, streamed) {
            (Ok(ws), Ok(())) => {
                let v: serde_json::Value = serde_json::from_str(&json).unwrap();
                assert_eq!(
                    v["columns"],
                    serde_json::to_value(&ws.columns).unwrap(),
                    "{col}"
                );
                Ok(ws.columns)
            }
            (Err(e), Err(s)) => {
                assert_eq!(e.code(), s.code(), "{col}");
                Err(e)
            }
            _ => unreachable!("both paths agree, asserted above"),
        }
    }

    #[test]
    fn col_collapsed_true_reads_true() {
        let cols = parse_cols(r#"<col min="1" max="1" collapsed="true"/>"#).unwrap();
        assert!(cols[0].collapsed);
    }

    #[test]
    fn col_flags_are_case_sensitive_and_keep_defaults() {
        let cols =
            parse_cols(r#"<col min="1" max="1" hidden="TRUE" customWidth="TRUE"/>"#).unwrap();
        assert!(!cols[0].hidden);
        assert!(!cols[0].custom_width);
        let cols =
            parse_cols(r#"<col hidden="1" customWidth="true" width="12.5" outlineLevel="2"/>"#)
                .unwrap();
        assert_eq!((cols[0].min, cols[0].max), (1, 1));
        assert!(cols[0].hidden && cols[0].custom_width);
        assert_eq!(cols[0].width, 12.5);
        assert_eq!(cols[0].outline_level, Some(2));
    }

    #[test]
    fn col_malformed_number_is_xml_parse_error() {
        for col in [
            r#"<col min="x" max="1"/>"#,
            r#"<col min="1" max="1" width="wide"/>"#,
            r#"<col min="1" max="1" outlineLevel="300"/>"#,
        ] {
            assert_eq!(parse_cols(col).unwrap_err().code(), "XML_PARSE", "{col}");
        }
    }

    /// Parse one `<c>` through both parsers and return its reference from the
    /// model path and from the JSON path.
    fn cell_reference_both_paths(cell: &str) -> (String, serde_json::Value) {
        let xml = format!(
            r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1">{cell}</row></sheetData></worksheet>"#
        );
        let ws = WorksheetXml::parse(xml.as_bytes()).unwrap();
        let mut json = String::new();
        WorksheetXml::parse_to_json(
            xml.as_bytes(),
            None,
            &[],
            &[],
            &crate::ooxml::worksheet::SheetContext::default(),
            &mut json,
        )
        .unwrap();
        let v: serde_json::Value = serde_json::from_str(&json).unwrap();
        (
            ws.rows[0].cells[0].reference.clone(),
            v["rows"][0]["cells"][0]["reference"].clone(),
        )
    }

    // quick-xml rejects only an identical QName, so `r` and `x:r` both reach
    // the local-name match in `parse_cell_attrs_into`: the last one wins.

    #[test]
    fn cell_prefixed_duplicate_r_last_wins() {
        let (model, json) =
            cell_reference_both_paths(r#"<c r="A1" xmlns:x="urn:x" x:r="B2"><v>1</v></c>"#);
        assert_eq!(model, "B2");
        assert_eq!(json, "B2");
    }

    #[test]
    fn self_closing_cell_prefixed_duplicate_r_last_wins() {
        let (model, json) = cell_reference_both_paths(r#"<c r="A1" xmlns:x="urn:x" x:r="B2"/>"#);
        assert_eq!(model, "B2");
        assert_eq!(json, "B2");
    }

    #[test]
    fn col_duplicate_attribute_is_error() {
        let err = parse_cols(r#"<col min="1" min="2" max="1"/>"#).unwrap_err();
        assert_eq!(err.code(), "XML_PARSE");
    }

    /// A hyperlink's target counts once, as kept: a second `r:id` (under
    /// another prefix bound to the relationships namespace) replaces the
    /// first copy.
    #[test]
    fn hyperlink_target_counts_once_per_hyperlink() {
        use crate::ooxml::relationships::{REL_HYPERLINK, Relationships};
        use crate::ooxml::shared_strings::ExpansionBudget;
        use crate::ooxml::worksheet::SheetContext;

        let target = "https://example.com/abcdefghij";
        let mut rels = Relationships::new();
        let id = rels.add_external(REL_HYPERLINK, target);
        let r = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        let xml = format!(
            r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="{r}" xmlns:q="{r}"><sheetData/><hyperlinks><hyperlink ref="A1" r:id="{id}" q:id="{id}"/></hyperlinks></worksheet>"#
        );
        let ctx = SheetContext {
            rels: Some(&rels),
            ..SheetContext::default()
        };
        for within in [true, false] {
            let limit = if within {
                target.len()
            } else {
                target.len() - 1
            };
            // Only the copies are under test: the stored list elements have
            // a limit of their own.
            let copies_only = ExpansionBudget {
                store_limit: usize::MAX,
                ..ExpansionBudget::new(limit)
            };
            let mut budget = copies_only;
            let model =
                WorksheetXml::parse_with_sst_within(xml.as_bytes(), None, &ctx, &mut budget);
            let mut budget = copies_only;
            let mut json = String::new();
            let streamed = WorksheetXml::parse_to_json_within(
                xml.as_bytes(),
                None,
                &[],
                &[],
                &ctx,
                &mut json,
                &mut budget,
            );
            if within {
                assert_eq!(model.unwrap().hyperlinks[0].target.as_deref(), Some(target));
                // The count fits; the sheet JSON, checked after each
                // hyperlink it writes, is longer than this limit.
                let err = streamed.unwrap_err();
                assert!(err.to_string().contains("workbook JSON reached"), "{err}");
                assert_eq!(budget.spent, target.len());
            } else {
                assert_eq!(model.unwrap_err().code(), "SIZE_LIMIT");
                assert_eq!(streamed.unwrap_err().code(), "SIZE_LIMIT");
            }
        }
    }

    /// Every prefix the root binds to the relationships namespace names a
    /// relationship, whatever the order of the declarations (the lookup
    /// searches the prefixes in sorted order); another prefix, another letter
    /// case and a colon-leading `:id` name none, also when the root binds the
    /// empty prefix to the relationships namespace.
    #[test]
    fn hyperlink_r_id_resolves_under_every_relationship_prefix() {
        use crate::ooxml::relationships::{REL_HYPERLINK, Relationships};
        use crate::ooxml::shared_strings::ExpansionBudget;
        use crate::ooxml::worksheet::SheetContext;

        let target = "https://example.com/x";
        let mut rels = Relationships::new();
        let id = rels.add_external(REL_HYPERLINK, target);
        let r = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        let xml = format!(
            r#"<worksheet xmlns:="{r}" xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="{r}" xmlns:q="{r}" xmlns:Z="{r}" xmlns:b="urn:other" xmlns:a="{r}"><sheetData/><hyperlinks><hyperlink ref="A1" r:id="{id}"/><hyperlink ref="A2" q:id="{id}"/><hyperlink ref="A3" Z:id="{id}"/><hyperlink ref="A4" a:id="{id}"/><hyperlink ref="A5" b:id="{id}"/><hyperlink ref="A6" :id="{id}"/><hyperlink ref="A7" R:id="{id}"/></hyperlinks></worksheet>"#
        );
        let ctx = SheetContext {
            rels: Some(&rels),
            ..SheetContext::default()
        };
        let mut budget = ExpansionBudget {
            store_limit: usize::MAX,
            ..ExpansionBudget::new(usize::MAX)
        };
        let ws = WorksheetXml::parse_with_sst_within(xml.as_bytes(), None, &ctx, &mut budget)
            .expect("sheet reads");
        let targets: Vec<Option<&str>> =
            ws.hyperlinks.iter().map(|h| h.target.as_deref()).collect();
        assert_eq!(
            targets,
            [
                Some(target),
                Some(target),
                Some(target),
                Some(target),
                None,
                None,
                None
            ]
        );
    }

    /// A rich entry counts its text and the floor of its runs on both paths.
    #[test]
    fn rich_shared_string_counts_text_and_runs() {
        use crate::ooxml::shared_strings::{
            ExpansionBudget, RichTextRun, SharedStringTable, runs_json_floor,
        };
        use crate::ooxml::worksheet::SheetContext;

        let runs = vec![RichTextRun {
            text: "ab".into(),
            font_name: Some("Calibri".into()),
            ..RichTextRun::default()
        }];
        let sst = SharedStringTable {
            strings: vec!["ab".into()],
            rich_runs: vec![Some(runs.clone())],
        };
        let xml = br#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>0</v></c></row></sheetData></worksheet>"#;
        let per_cell = 2 + runs_json_floor(&runs);
        let ctx = SheetContext::default();
        let mut budget = ExpansionBudget::new(usize::MAX);
        let ws = WorksheetXml::parse_with_sst_within(xml, Some(&sst), &ctx, &mut budget).unwrap();
        assert_eq!(ws.rows[0].cells[1].rich_text.as_deref(), Some(&runs[..]));
        assert_eq!(budget.spent, 2 * per_cell);
        let mut budget = ExpansionBudget::new(usize::MAX);
        let mut json = String::new();
        WorksheetXml::parse_to_json_within(xml, Some(&sst), &[], &[], &ctx, &mut json, &mut budget)
            .unwrap();
        assert_eq!(budget.spent, 2 * per_cell);
        assert!(budget.spent <= json.len());
    }
}
