//! The worksheet XML writer: `WorksheetXml::to_xml`, `to_xml_with_sst` and
//! `to_xml_with_legacy_drawings` write the children of `<worksheet>` in
//! `CT_Worksheet` order and recompute `<dimension>` from the rows;
//! `to_xml_with_legacy_drawings` takes the rIds of the sheet's tables,
//! drawing, VML drawings (`<legacyDrawing>`, `<legacyDrawingHF>`) and
//! external hyperlinks from the caller, and `to_xml_with_sst` is it without
//! the VML drawings.

use core::hint::cold_path;
use std::collections::HashSet;

use quick_xml::Writer;
use quick_xml::events::{BytesEnd, BytesStart, BytesText, Event};

use super::{
    Cell, CellType, Cfvo, MAX_CF_RULE_FORMULAS, PageBreak, WorksheetXml, cf_formula_count_message,
    non_hash_password_message,
};
use crate::ooxml::cell::{CellRef, col_to_letters};
use crate::ooxml::cell::{MAX_COL, MAX_ROW};
use crate::ooxml::color::{ColorOut, resolve_color, write_color};
use crate::ooxml::xml::{REL_NS, SPREADSHEET_NS, XsdBuf, fmt_xsd_double, needs_preserve, xml_attr};
use crate::ooxml::{write_xml_decl, xml_write_err};
use crate::{ModernXlsxError, Result};

/// Excel's default data-bar colour, written for a bar (or colour-scale stop)
/// that carries no colour of its own.
const DEFAULT_BAR_RGB: &str = "FF638EC6";

/// Write one `<cfvo type val?/>` per entry of `cfvos`.
fn write_cfvos<W: std::io::Write>(writer: &mut Writer<W>, cfvos: &[Cfvo]) -> Result<()> {
    for cfvo in cfvos {
        let mut cfvo_elem = BytesStart::new("cfvo");
        cfvo_elem.push_attribute(xml_attr("type", cfvo.cfvo_type.as_str()));
        if let Some(ref v) = cfvo.val {
            cfvo_elem.push_attribute(xml_attr("val", v.as_str()));
        }
        writer
            .write_event(Event::Empty(cfvo_elem))
            .map_err(xml_write_err)?;
    }
    Ok(())
}

/// Write `<tag count manualBreakCount>`, one `<brk id min? max? man="1"?/>`
/// per entry of `breaks`, and `</tag>` (`<rowBreaks>` or `<colBreaks>`).
fn write_breaks<W: std::io::Write>(
    writer: &mut Writer<W>,
    tag: &str,
    breaks: &[PageBreak],
    ibuf: &mut itoa::Buffer,
) -> Result<()> {
    let mut start = BytesStart::new(tag);
    start.push_attribute(("count", ibuf.format(breaks.len())));
    let manual_count = breaks.iter().filter(|b| b.man).count();
    start.push_attribute(("manualBreakCount", ibuf.format(manual_count)));
    writer
        .write_event(Event::Start(start))
        .map_err(xml_write_err)?;
    for brk in breaks {
        let mut elem = BytesStart::new("brk");
        elem.push_attribute(("id", ibuf.format(brk.id)));
        if let Some(min) = brk.min {
            elem.push_attribute(("min", ibuf.format(min)));
        }
        if let Some(max) = brk.max {
            elem.push_attribute(("max", ibuf.format(max)));
        }
        if brk.man {
            elem.push_attribute(("man", "1"));
        }
        writer
            .write_event(Event::Empty(elem))
            .map_err(xml_write_err)?;
    }
    writer
        .write_event(Event::End(BytesEnd::new(tag)))
        .map_err(xml_write_err)
}

/// Push the `CT_CellFormula` attributes of `cell` onto `<f>`. The data-table
/// delete flags come from `formula_del1`/`formula_del2`, else from the 1.x
/// `formula_dtr1`/`formula_dtr2`; `cm` belongs on `<c>` and is not written
/// here.
fn push_formula_attrs(f_elem: &mut BytesStart<'_>, cell: &Cell, ibuf: &mut itoa::Buffer) {
    if let Some(ref ft) = cell.formula_type {
        f_elem.push_attribute(xml_attr("t", ft.as_str()));
    }
    if let Some(ref fr) = cell.formula_ref {
        f_elem.push_attribute(xml_attr("ref", fr.as_str()));
    }
    if let Some(si) = cell.shared_index {
        f_elem.push_attribute(("si", ibuf.format(si)));
    }
    if let Some(ref r1) = cell.formula_r1 {
        f_elem.push_attribute(xml_attr("r1", r1.as_str()));
    }
    if let Some(ref r2) = cell.formula_r2 {
        f_elem.push_attribute(xml_attr("r2", r2.as_str()));
    }
    if cell.formula_dt2d == Some(true) {
        f_elem.push_attribute(("dt2D", "1"));
    }
    if cell.formula_dtr == Some(true) {
        f_elem.push_attribute(("dtr", "1"));
    }
    if cell.formula_del1.or(cell.formula_dtr1) == Some(true) {
        f_elem.push_attribute(("del1", "1"));
    }
    if cell.formula_del2.or(cell.formula_dtr2) == Some(true) {
        f_elem.push_attribute(("del2", "1"));
    }
}

impl Cell {
    /// Whether the cell carries shared-formula keys: `formula_type` `"shared"`
    /// or a `shared_index`.
    fn has_shared_keys(&self) -> bool {
        self.formula_type.as_deref() == Some("shared") || self.shared_index.is_some()
    }

    /// The formula text the writer puts in `<f>` ([`formula_text`]) when it is
    /// non-empty after trimming; `None` for no text, blank text or a bare `=`.
    fn usable_formula_text(&self) -> Option<&str> {
        self.formula
            .as_deref()
            .map(|f| formula_text(f, self.formula_type.as_deref()))
            .filter(|text| !text.trim().is_empty())
    }

    /// The group whose master this cell is: `formula_type` `"shared"`, a
    /// `shared_index`, a non-empty `formula_ref` and usable formula text
    /// ([`Cell::usable_formula_text`]). Excel refuses a shared master without
    /// text or without a ref. The one definition of a master, used by
    /// [`WorksheetXml::shared_formula_masters`] and [`Cell::formula_markup`].
    fn shared_master_of(&self) -> Option<u32> {
        let is_master = self.formula_type.as_deref() == Some("shared")
            && self.formula_ref.as_deref().is_some_and(|r| !r.is_empty())
            && self.usable_formula_text().is_some();
        self.shared_index.filter(|_| is_master)
    }

    /// How the worksheet writer writes this cell's formula; `shared_masters`
    /// is [`WorksheetXml::shared_formula_masters`] of its sheet. A cell with
    /// `formula_type` `"shared"` and `shared_index` `N`:
    /// - when the sheet has the master of `N` ([`Cell::shared_master_of`]),
    ///   the master itself or any other cell of `N`: [`FormulaMarkup::Stored`]
    ///   (`<f t="shared" si="N"/>` for a child without text,
    ///   `<f t="shared" si="N">text</f>` for one with text);
    /// - without the master of `N` in the sheet, with usable text:
    ///   [`FormulaMarkup::Plain`], an ordinary `<f>text</f>`;
    /// - without the master of `N` and without usable text:
    ///   [`FormulaMarkup::Omit`], since Excel refuses a child whose group has
    ///   no master.
    ///
    /// Any other cell with shared keys (`formula_type` `"shared"` without a
    /// `shared_index`, or a `shared_index` with another type) gets
    /// [`FormulaMarkup::Stored`] when it has non-empty text and no `<f>`
    /// otherwise. A cell without shared keys gets [`FormulaMarkup::Stored`]
    /// when it has formula text (empty included) or a `formula_type` (an array
    /// or data-table formula).
    fn formula_markup(&self, shared_masters: &HashSet<u32>) -> FormulaMarkup<'_> {
        if let (Some("shared"), Some(si)) = (self.formula_type.as_deref(), self.shared_index) {
            // The master of `si` is in `shared_masters` too.
            if shared_masters.contains(&si) {
                return FormulaMarkup::Stored;
            }
            return self
                .usable_formula_text()
                .map_or(FormulaMarkup::Omit, FormulaMarkup::Plain);
        }
        let stored = if self.has_shared_keys() {
            self.formula.as_deref().is_some_and(|f| !f.is_empty())
        } else {
            self.formula.is_some() || self.formula_type.is_some()
        };
        if stored {
            FormulaMarkup::Stored
        } else {
            FormulaMarkup::Omit
        }
    }

    /// Whether the worksheet writer gives this cell an `<f>`: its
    /// [`Cell::formula_markup`] is not [`FormulaMarkup::Omit`]. The calcChain
    /// writer keeps exactly the cells this accepts.
    pub(crate) fn writes_formula(&self, shared_masters: &HashSet<u32>) -> bool {
        !matches!(self.formula_markup(shared_masters), FormulaMarkup::Omit)
    }
}

/// How a cell's formula is written; see [`Cell::formula_markup`].
enum FormulaMarkup<'a> {
    /// No `<f>`: the cell is written with its value alone, or empty.
    Omit,
    /// `<f>` with the cell's own attributes and text; self-closing when the
    /// cell has no formula text, or empty text and shared keys.
    Stored,
    /// `<f>text</f>` without attributes: a shared cell whose group has no
    /// master in the sheet, written as an ordinary formula.
    Plain(&'a str),
}

/// Normalise cell formula text for the `<f>` element.
///
/// ECMA-376 Part 1 §18.3.1.40 (`f`): the element content is the formula
/// without a leading `=` and without the `{...}` the Excel UI shows around an
/// array formula. For an array formula (`t="array"`) whose trimmed text is
/// `{=...}` only the inner text is returned; otherwise one leading `=` is
/// removed. A `{` not followed by `=` (an array constant such as `{1,2,3}`)
/// is never touched. Borrows from the input, no allocation.
fn formula_text<'a>(formula: &'a str, formula_type: Option<&str>) -> &'a str {
    if formula_type == Some("array")
        && let Some(inner) = formula
            .trim()
            .strip_prefix("{=")
            .and_then(|rest| rest.strip_suffix('}'))
    {
        return inner;
    }
    formula.strip_prefix('=').unwrap_or(formula)
}

impl WorksheetXml {
    /// The `si` of every shared-formula master in the sheet
    /// ([`Cell::shared_master_of`]: `formula_type` `"shared"`, a
    /// `shared_index`, a non-empty `formula_ref` and usable formula text).
    /// One pass over the cells; see [`Cell::writes_formula`].
    pub(crate) fn shared_formula_masters(&self) -> HashSet<u32> {
        self.rows
            .iter()
            .flat_map(|row| &row.cells)
            .filter_map(Cell::shared_master_of)
            .collect()
    }

    /// The used range for `<dimension ref>`: the bounding box of the rows
    /// that hold a cell other than a `Stub`, `"A1"` when there is none, and
    /// a single reference when the box is one cell.
    fn used_range(&self) -> String {
        let (mut min_row, mut max_row) = (u32::MAX, 0);
        let (mut min_col, mut max_col) = (u32::MAX, 0);
        for row in &self.rows {
            let mut has_cell = false;
            for cell in row.cells.iter().filter(|c| c.cell_type != CellType::Stub) {
                if let Ok(r) = CellRef::parse(&cell.reference) {
                    min_col = min_col.min(r.col);
                    max_col = max_col.max(r.col);
                    has_cell = true;
                }
            }
            if has_cell {
                min_row = min_row.min(row.index);
                max_row = max_row.max(row.index);
            }
        }
        if min_row == u32::MAX {
            return "A1".to_owned();
        }
        let mut ibuf = itoa::Buffer::new();
        let mut out = col_to_letters(min_col);
        out.push_str(ibuf.format(min_row));
        if (min_col, min_row) != (max_col, max_row) {
            out.push(':');
            out.push_str(&col_to_letters(max_col));
            out.push_str(ibuf.format(max_row));
        }
        out
    }

    /// Serialize this worksheet to valid XML bytes. With no relationship part
    /// to hold them, external hyperlink targets are not written.
    pub fn to_xml(&self) -> Result<Vec<u8>> {
        self.to_xml_with_sst(None, &[], None, &[])
    }

    /// Serialize to worksheet XML bytes, optionally remapping SharedString
    /// cell values to SST indices on-the-fly (avoiding a full clone of the
    /// worksheet).
    ///
    /// `table_r_ids` are the relationship IDs for `<tableParts>` elements;
    /// pass an empty slice when no tables are attached to this sheet.
    ///
    /// `drawing_r_id` is the relationship ID for the `<drawing>` element
    /// referencing the drawing XML that contains chart anchors.
    ///
    /// `hyperlink_r_ids` runs parallel to `hyperlinks`: entry `i` is the rId
    /// of hyperlink `i`'s external relationship, written as `r:id`. A missing
    /// or `None` entry writes no `r:id`.
    pub fn to_xml_with_sst(
        &self,
        sst: Option<&crate::ooxml::shared_strings::SharedStringTableBuilder<'_>>,
        table_r_ids: &[String],
        drawing_r_id: Option<&str>,
        hyperlink_r_ids: &[Option<String>],
    ) -> Result<Vec<u8>> {
        self.to_xml_with_legacy_drawings(
            sst,
            table_r_ids,
            drawing_r_id,
            None,
            None,
            hyperlink_r_ids,
        )
    }

    /// [`Self::to_xml_with_sst`] for a sheet that also names VML drawings:
    /// `legacy_drawing_r_id` is written as `<legacyDrawing>` (the note and
    /// form-control shapes) and `legacy_drawing_hf_r_id` as
    /// `<legacyDrawingHF>` (the header and footer pictures), both after
    /// `<drawing>`.
    pub fn to_xml_with_legacy_drawings(
        &self,
        sst: Option<&crate::ooxml::shared_strings::SharedStringTableBuilder<'_>>,
        table_r_ids: &[String],
        drawing_r_id: Option<&str>,
        legacy_drawing_r_id: Option<&str>,
        legacy_drawing_hf_r_id: Option<&str>,
        hyperlink_r_ids: &[Option<String>],
    ) -> Result<Vec<u8>> {
        let mut buf: Vec<u8> = Vec::with_capacity(1024 + self.rows.len() * 128);
        let mut writer = Writer::new(&mut buf);
        let mut ibuf = itoa::Buffer::new();
        let mut xsd = XsdBuf::new();

        write_xml_decl(&mut writer)?;

        // <worksheet xmlns="..." xmlns:r="...">
        let mut ws = BytesStart::new("worksheet");
        ws.push_attribute(("xmlns", SPREADSHEET_NS));
        ws.push_attribute(("xmlns:r", REL_NS));
        writer
            .write_event(Event::Start(ws))
            .map_err(xml_write_err)?;

        // <sheetPr> — write if there is a tab colour, outline_properties or
        // fit-to-page.
        let tab_color = resolve_color(self.tab_color.as_deref(), self.tab_color_spec.as_ref())
            .filter(|c| !c.is_default_spec());
        let fit_to_page = self
            .page_setup
            .as_ref()
            .is_some_and(|ps| ps.fit_to_page == Some(true));
        if tab_color.is_some() || self.outline_properties.is_some() || fit_to_page {
            writer
                .write_event(Event::Start(BytesStart::new("sheetPr")))
                .map_err(xml_write_err)?;
            if let Some(color) = tab_color {
                write_color(&mut writer, "tabColor", color, &mut xsd)?;
            }
            if let Some(ref op) = self.outline_properties {
                let mut elem = BytesStart::new("outlinePr");
                if !op.summary_below {
                    elem.push_attribute(("summaryBelow", "0"));
                }
                if !op.summary_right {
                    elem.push_attribute(("summaryRight", "0"));
                }
                writer
                    .write_event(Event::Empty(elem))
                    .map_err(xml_write_err)?;
            }
            if fit_to_page {
                let mut elem = BytesStart::new("pageSetUpPr");
                elem.push_attribute(("fitToPage", "1"));
                writer
                    .write_event(Event::Empty(elem))
                    .map_err(xml_write_err)?;
            }
            writer
                .write_event(Event::End(BytesEnd::new("sheetPr")))
                .map_err(xml_write_err)?;
        }

        // <dimension> — recomputed from the rows; the stored `dimension` is
        // only what the source file said.
        let mut dim = BytesStart::new("dimension");
        dim.push_attribute(("ref", self.used_range().as_str()));
        writer
            .write_event(Event::Empty(dim))
            .map_err(xml_write_err)?;

        // <sheetViews> — always written. A sheet with no sheet_view, pane or
        // selection gets the view Excel writes for one,
        // `<sheetView workbookViewId="0"/>`: without any <sheetView>, Excel on
        // a 150% display (144 DPI) showed setRowHeight(1, 40) as 26.7 points;
        // with this view it shows 40. The reader takes that bare view back as
        // no sheet_view, so a view-less model round-trips.
        writer
            .write_event(Event::Start(BytesStart::new("sheetViews")))
            .map_err(xml_write_err)?;

        let mut sv_elem = BytesStart::new("sheetView");

        // Write non-default sheetView attributes.
        if let Some(ref sv) = self.sheet_view {
            if !sv.show_grid_lines {
                sv_elem.push_attribute(("showGridLines", "0"));
            }
            if !sv.show_row_col_headers {
                sv_elem.push_attribute(("showRowColHeaders", "0"));
            }
            if !sv.show_zeros {
                sv_elem.push_attribute(("showZeros", "0"));
            }
            if sv.right_to_left {
                sv_elem.push_attribute(("rightToLeft", "1"));
            }
            if sv.tab_selected {
                sv_elem.push_attribute(("tabSelected", "1"));
            }
            if !sv.show_ruler {
                sv_elem.push_attribute(("showRuler", "0"));
            }
            if !sv.show_outline_symbols {
                sv_elem.push_attribute(("showOutlineSymbols", "0"));
            }
            if !sv.show_white_space {
                sv_elem.push_attribute(("showWhiteSpace", "0"));
            }
            if !sv.default_grid_color {
                sv_elem.push_attribute(("defaultGridColor", "0"));
            }
            if let Some(z) = sv.zoom_scale {
                sv_elem.push_attribute(("zoomScale", ibuf.format(z)));
            }
            if let Some(z) = sv.zoom_scale_normal {
                sv_elem.push_attribute(("zoomScaleNormal", ibuf.format(z)));
            }
            if let Some(z) = sv.zoom_scale_page_layout_view {
                sv_elem.push_attribute(("zoomScalePageLayoutView", ibuf.format(z)));
            }
            if let Some(z) = sv.zoom_scale_sheet_layout_view {
                sv_elem.push_attribute(("zoomScaleSheetLayoutView", ibuf.format(z)));
            }
            if let Some(c) = sv.color_id {
                sv_elem.push_attribute(("colorId", ibuf.format(c)));
            }
            if let Some(ref v) = sv.view {
                sv_elem.push_attribute(xml_attr("view", v.as_str()));
            }
        }

        sv_elem.push_attribute(("workbookViewId", "0"));

        // Determine if sheetView has child elements (pane, selection).
        let has_children = self.split_pane.is_some()
            || self.frozen_pane.is_some()
            || !self.pane_selections.is_empty();

        if has_children {
            writer
                .write_event(Event::Start(sv_elem))
                .map_err(xml_write_err)?;

            // Write pane element (split takes priority over frozen).
            if let Some(ref sp) = self.split_pane {
                let mut pane_elem = BytesStart::new("pane");
                if let Some(x) = sp.vertical {
                    pane_elem
                        .push_attribute(("xSplit", fmt_xsd_double(x, &mut xsd, "pane xSplit")?));
                }
                if let Some(y) = sp.horizontal {
                    pane_elem
                        .push_attribute(("ySplit", fmt_xsd_double(y, &mut xsd, "pane ySplit")?));
                }
                if let Some(ref tlc) = sp.top_left_cell {
                    pane_elem.push_attribute(xml_attr("topLeftCell", tlc.as_str()));
                }
                if let Some(ref ap) = sp.active_pane {
                    pane_elem.push_attribute(xml_attr("activePane", ap.as_str()));
                }
                pane_elem.push_attribute(("state", "split"));
                writer
                    .write_event(Event::Empty(pane_elem))
                    .map_err(xml_write_err)?;
            } else if let Some(ref pane) = self.frozen_pane {
                if pane.cols > MAX_COL || pane.rows > MAX_ROW {
                    return Err(ModernXlsxError::InvalidCellRef(format!(
                        "frozen pane of {} columns and {} rows puts topLeftCell past XFD1048576",
                        pane.cols, pane.rows
                    )));
                }
                let mut pane_elem = BytesStart::new("pane");
                if pane.cols > 0 {
                    pane_elem.push_attribute(("xSplit", ibuf.format(pane.cols)));
                }
                if pane.rows > 0 {
                    pane_elem.push_attribute(("ySplit", ibuf.format(pane.rows)));
                }
                let mut top_left = col_to_letters(pane.cols);
                top_left.push_str(ibuf.format(pane.rows.saturating_add(1)));
                pane_elem.push_attribute(xml_attr("topLeftCell", top_left.as_str()));
                let active_pane = match (pane.rows > 0, pane.cols > 0) {
                    (true, true) => "bottomRight",
                    (true, false) => "bottomLeft",
                    (false, true) => "topRight",
                    (false, false) => "bottomLeft",
                };
                pane_elem.push_attribute(("activePane", active_pane));
                pane_elem.push_attribute(("state", "frozen"));
                writer
                    .write_event(Event::Empty(pane_elem))
                    .map_err(xml_write_err)?;
            }

            // Write <selection> elements.
            for sel in &self.pane_selections {
                let mut sel_elem = BytesStart::new("selection");
                if let Some(ref p) = sel.pane {
                    sel_elem.push_attribute(xml_attr("pane", p.as_str()));
                }
                if let Some(ref ac) = sel.active_cell {
                    sel_elem.push_attribute(xml_attr("activeCell", ac.as_str()));
                }
                if let Some(ref sq) = sel.sqref {
                    sel_elem.push_attribute(xml_attr("sqref", sq.as_str()));
                }
                writer
                    .write_event(Event::Empty(sel_elem))
                    .map_err(xml_write_err)?;
            }

            writer
                .write_event(Event::End(BytesEnd::new("sheetView")))
                .map_err(xml_write_err)?;
        } else {
            // No children — write self-closing <sheetView ... />
            writer
                .write_event(Event::Empty(sv_elem))
                .map_err(xml_write_err)?;
        }

        writer
            .write_event(Event::End(BytesEnd::new("sheetViews")))
            .map_err(xml_write_err)?;

        // <cols> — only if non-empty. Attributes follow CT_Col order (min, max,
        // width, style, ..., outlineLevel, collapsed) except that customWidth
        // comes before hidden; bestFit follows hidden, phonetic follows
        // bestFit, and false flags and an absent style are left out.
        if !self.columns.is_empty() {
            writer
                .write_event(Event::Start(BytesStart::new("cols")))
                .map_err(xml_write_err)?;

            for col in &self.columns {
                let mut elem = BytesStart::new("col");
                elem.push_attribute(("min", ibuf.format(col.min)));
                elem.push_attribute(("max", ibuf.format(col.max)));
                elem.push_attribute(("width", fmt_xsd_double(col.width, &mut xsd, "col width")?));
                if let Some(style) = col.style_index {
                    elem.push_attribute(("style", ibuf.format(style)));
                }
                if col.custom_width {
                    elem.push_attribute(("customWidth", "1"));
                }
                if col.hidden {
                    elem.push_attribute(("hidden", "1"));
                }
                if col.best_fit {
                    elem.push_attribute(("bestFit", "1"));
                }
                if col.phonetic {
                    elem.push_attribute(("phonetic", "1"));
                }
                if let Some(level) = col.outline_level
                    && level > 0
                {
                    elem.push_attribute(("outlineLevel", ibuf.format(level)));
                }
                if col.collapsed {
                    elem.push_attribute(("collapsed", "1"));
                }
                writer
                    .write_event(Event::Empty(elem))
                    .map_err(xml_write_err)?;
            }

            writer
                .write_event(Event::End(BytesEnd::new("cols")))
                .map_err(xml_write_err)?;
        }

        // <sheetData>
        writer
            .write_event(Event::Start(BytesStart::new("sheetData")))
            .map_err(xml_write_err)?;

        let shared_masters = self.shared_formula_masters();
        for row in &self.rows {
            let mut row_elem = BytesStart::new("row");
            row_elem.push_attribute(("r", ibuf.format(row.index)));
            // `s` applies only with `customFormat` (ECMA-376 Part 1, §18.3.1.73).
            if let Some(s) = row.style_index {
                row_elem.push_attribute(("s", ibuf.format(s)));
                row_elem.push_attribute(("customFormat", "1"));
            }
            if let Some(ht) = row.height {
                row_elem.push_attribute(("ht", fmt_xsd_double(ht, &mut xsd, "row ht")?));
            }
            if row.hidden {
                row_elem.push_attribute(("hidden", "1"));
            }
            // Excel keeps `ht` only with `customHeight`; no height, nothing to keep.
            if row.custom_height && row.height.is_some() {
                row_elem.push_attribute(("customHeight", "1"));
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
                    .map_err(xml_write_err)?;
            } else {
                writer
                    .write_event(Event::Start(row_elem))
                    .map_err(xml_write_err)?;

                for cell in &row.cells {
                    let mut c_elem = BytesStart::new("c");
                    c_elem.push_attribute(xml_attr("r", cell.reference.as_str()));

                    // Only write t attribute if not Number (the default); a Stub
                    // cell writes no t either.
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

                    // cm: the cell-metadata index; a 1.x dynamic array means
                    // index 1 (the XLDAPR record of the generated metadata part).
                    if let Some(cm) = cell.cell_metadata {
                        c_elem.push_attribute(("cm", ibuf.format(cm)));
                    } else if cell.dynamic_array == Some(true) {
                        c_elem.push_attribute(("cm", "1"));
                    }

                    // An inline string keeps its text in <is><t>, never in <v>;
                    // `inline_string` wins over `value` (the reader fills both).
                    let is_inline = cell.cell_type == CellType::InlineStr;
                    let inline_text = if is_inline {
                        cell.inline_string.as_deref().or(cell.value.as_deref())
                    } else {
                        None
                    };
                    let plain_value = if is_inline {
                        None
                    } else {
                        cell.value.as_deref()
                    };
                    // An <f> is content: a shared-formula child (t="shared" si, no
                    // text) of a master in this sheet keeps its self-closing <f>
                    // without a cached value.
                    let formula_markup = cell.formula_markup(&shared_masters);
                    let writes_formula = !matches!(formula_markup, FormulaMarkup::Omit);
                    let has_content =
                        writes_formula || inline_text.is_some() || plain_value.is_some();

                    if has_content {
                        writer
                            .write_event(Event::Start(c_elem))
                            .map_err(xml_write_err)?;

                        // <f>...</f> with the cell's attributes, or self-closing
                        // with attributes and no text (a shared-formula child,
                        // or an array or data-table cell without text); a shared
                        // cell whose group has no master gets a plain <f>text</f>.
                        let (f_elem, text) = match formula_markup {
                            FormulaMarkup::Omit => (None, None),
                            FormulaMarkup::Stored => {
                                let mut f_elem = BytesStart::new("f");
                                push_formula_attrs(&mut f_elem, cell, &mut ibuf);
                                let text = cell
                                    .formula
                                    .as_deref()
                                    .filter(|f| !f.is_empty() || !cell.has_shared_keys())
                                    .map(|f| formula_text(f, cell.formula_type.as_deref()));
                                (Some(f_elem), text)
                            }
                            FormulaMarkup::Plain(text) => (Some(BytesStart::new("f")), Some(text)),
                        };
                        match (f_elem, text) {
                            (Some(f_elem), Some(text)) => {
                                writer
                                    .write_event(Event::Start(f_elem))
                                    .map_err(xml_write_err)?;
                                writer
                                    .write_event(Event::Text(BytesText::new(
                                        crate::ooxml::sanitize_xml_text(text).as_ref(),
                                    )))
                                    .map_err(xml_write_err)?;
                                writer
                                    .write_event(Event::End(BytesEnd::new("f")))
                                    .map_err(xml_write_err)?;
                            }
                            (Some(f_elem), None) => {
                                writer
                                    .write_event(Event::Empty(f_elem))
                                    .map_err(xml_write_err)?;
                            }
                            (None, _) => {}
                        }

                        // <is><t>...</t></is> for inline strings. CT_Cell order is
                        // f, v, is: an inline cell has no <v> at all.
                        if let Some(text) = inline_text {
                            let text = crate::ooxml::sanitize_xml_text(text);
                            let mut t_elem = BytesStart::new("t");
                            if needs_preserve(&text) {
                                t_elem.push_attribute(("xml:space", "preserve"));
                            }
                            writer
                                .write_event(Event::Start(BytesStart::new("is")))
                                .map_err(xml_write_err)?;
                            writer
                                .write_event(Event::Start(t_elem))
                                .map_err(xml_write_err)?;
                            writer
                                .write_event(Event::Text(BytesText::new(text.as_ref())))
                                .map_err(xml_write_err)?;
                            writer
                                .write_event(Event::End(BytesEnd::new("t")))
                                .map_err(xml_write_err)?;
                            writer
                                .write_event(Event::End(BytesEnd::new("is")))
                                .map_err(xml_write_err)?;
                        }

                        // <v>...</v>
                        if let Some(value) = plain_value {
                            writer
                                .write_event(Event::Start(BytesStart::new("v")))
                                .map_err(xml_write_err)?;
                            // If an SST builder is provided and this is a SharedString cell,
                            // write the SST index instead of the raw string value: the
                            // plain entry, or the rich entry with this cell's runs.
                            if cell.cell_type == CellType::SharedString {
                                if let Some(sst_builder) = sst {
                                    let runs = cell.rich_text.as_deref();
                                    let idx =
                                        sst_builder.index_of(value, runs).ok_or_else(|| {
                                            ModernXlsxError::InvalidCellValue(format!(
                                                "SharedString cell has unmapped value: {value}"
                                            ))
                                        })?;
                                    writer
                                        .write_event(Event::Text(BytesText::new(ibuf.format(idx))))
                                        .map_err(xml_write_err)?;
                                } else {
                                    writer
                                        .write_event(Event::Text(BytesText::new(
                                            crate::ooxml::sanitize_xml_text(value).as_ref(),
                                        )))
                                        .map_err(xml_write_err)?;
                                }
                            } else {
                                writer
                                    .write_event(Event::Text(BytesText::new(
                                        crate::ooxml::sanitize_xml_text(value).as_ref(),
                                    )))
                                    .map_err(xml_write_err)?;
                            }
                            writer
                                .write_event(Event::End(BytesEnd::new("v")))
                                .map_err(xml_write_err)?;
                        }

                        writer
                            .write_event(Event::End(BytesEnd::new("c")))
                            .map_err(xml_write_err)?;
                    } else {
                        writer
                            .write_event(Event::Empty(c_elem))
                            .map_err(xml_write_err)?;
                    }
                }

                writer
                    .write_event(Event::End(BytesEnd::new("row")))
                    .map_err(xml_write_err)?;
            }
        }

        // </sheetData>
        writer
            .write_event(Event::End(BytesEnd::new("sheetData")))
            .map_err(xml_write_err)?;

        // <sheetProtection> — only if present. Attributes follow CT_SheetProtection
        // order; a flag is written ("1" or "0") only when it differs from its
        // schema default, and a `plain_password` is hashed here. A `password`
        // that is not a legacy hash is an error rather than a sheet that looks
        // protected: Excel treats such a value as no password at all.
        if let Some(ref stored) = self.sheet_protection {
            if stored.has_non_hash_password() {
                cold_path();
                return Err(ModernXlsxError::InvalidWorkbook(format!(
                    "Write failed: {}",
                    non_hash_password_message(None)
                )));
            }
            let hashed = stored.with_hashed_password()?;
            let sp = hashed.as_ref().unwrap_or(stored);
            let mut elem = BytesStart::new("sheetProtection");
            for (name, value) in [
                ("password", &sp.password),
                ("algorithmName", &sp.algorithm_name),
                ("hashValue", &sp.hash_value),
                ("saltValue", &sp.salt_value),
            ] {
                if let Some(v) = value {
                    elem.push_attribute(xml_attr(name, v.as_str()));
                }
            }
            if let Some(n) = sp.spin_count {
                elem.push_attribute(("spinCount", ibuf.format(n)));
            }
            for (name, value, default) in sp.flags() {
                if value != default {
                    elem.push_attribute((name, if value { "1" } else { "0" }));
                }
            }
            writer
                .write_event(Event::Empty(elem))
                .map_err(xml_write_err)?;
        }

        // <autoFilter> — only if present.
        if let Some(ref af) = self.auto_filter {
            let mut elem = BytesStart::new("autoFilter");
            elem.push_attribute(xml_attr("ref", af.range.as_str()));
            if af.filter_columns.is_empty() {
                writer
                    .write_event(Event::Empty(elem))
                    .map_err(xml_write_err)?;
            } else {
                writer
                    .write_event(Event::Start(elem))
                    .map_err(xml_write_err)?;
                for fc in &af.filter_columns {
                    let mut fc_elem = BytesStart::new("filterColumn");
                    fc_elem.push_attribute(("colId", ibuf.format(fc.col_id)));
                    writer
                        .write_event(Event::Start(fc_elem))
                        .map_err(xml_write_err)?;

                    if let Some(ref cf) = fc.custom_filters {
                        // <customFilters>
                        let mut cf_elem = BytesStart::new("customFilters");
                        if cf.and_op {
                            cf_elem.push_attribute(("and", "1"));
                        }
                        writer
                            .write_event(Event::Start(cf_elem))
                            .map_err(xml_write_err)?;
                        for item in &cf.filters {
                            let mut item_elem = BytesStart::new("customFilter");
                            if let Some(ref op) = item.operator {
                                item_elem.push_attribute(xml_attr("operator", op.as_str()));
                            }
                            item_elem.push_attribute(xml_attr("val", item.val.as_str()));
                            writer
                                .write_event(Event::Empty(item_elem))
                                .map_err(xml_write_err)?;
                        }
                        writer
                            .write_event(Event::End(BytesEnd::new("customFilters")))
                            .map_err(xml_write_err)?;
                    } else {
                        // <filters>
                        writer
                            .write_event(Event::Start(BytesStart::new("filters")))
                            .map_err(xml_write_err)?;
                        for fv in &fc.filters {
                            let mut f_elem = BytesStart::new("filter");
                            f_elem.push_attribute(xml_attr("val", fv.as_str()));
                            writer
                                .write_event(Event::Empty(f_elem))
                                .map_err(xml_write_err)?;
                        }
                        writer
                            .write_event(Event::End(BytesEnd::new("filters")))
                            .map_err(xml_write_err)?;
                    }

                    writer
                        .write_event(Event::End(BytesEnd::new("filterColumn")))
                        .map_err(xml_write_err)?;
                }
                writer
                    .write_event(Event::End(BytesEnd::new("autoFilter")))
                    .map_err(xml_write_err)?;
            }
        }

        // <mergeCells> — only if non-empty.
        if !self.merge_cells.is_empty() {
            let mut mc = BytesStart::new("mergeCells");
            mc.push_attribute(("count", ibuf.format(self.merge_cells.len())));
            writer
                .write_event(Event::Start(mc))
                .map_err(xml_write_err)?;

            for ref_str in &self.merge_cells {
                let mut elem = BytesStart::new("mergeCell");
                elem.push_attribute(xml_attr("ref", ref_str.as_str()));
                writer
                    .write_event(Event::Empty(elem))
                    .map_err(xml_write_err)?;
            }

            writer
                .write_event(Event::End(BytesEnd::new("mergeCells")))
                .map_err(xml_write_err)?;
        }

        // <conditionalFormatting> — only if non-empty.
        for cf in &self.conditional_formatting {
            let mut cf_elem = BytesStart::new("conditionalFormatting");
            cf_elem.push_attribute(xml_attr("sqref", cf.sqref.as_str()));
            writer
                .write_event(Event::Start(cf_elem))
                .map_err(xml_write_err)?;

            for rule in &cf.rules {
                // CT_CfRule holds at most three <formula>: more is an error,
                // never a truncated rule.
                let count = rule.formula_count();
                if count > MAX_CF_RULE_FORMULAS {
                    cold_path();
                    return Err(ModernXlsxError::InvalidWorkbook(format!(
                        "Write failed: {}",
                        cf_formula_count_message(None, &cf.sqref, rule.priority, count)
                    )));
                }
                let mut rule_elem = BytesStart::new("cfRule");
                rule_elem.push_attribute(xml_attr("type", rule.rule_type.as_str()));
                if let Some(dxf_id) = rule.dxf_id {
                    rule_elem.push_attribute(("dxfId", ibuf.format(dxf_id)));
                }
                rule_elem.push_attribute(("priority", ibuf.format(rule.priority)));
                // The rest in CT_CfRule order: stopIfTrue, aboveAverage,
                // percent, bottom, operator, text, timePeriod, rank, stdDev,
                // equalAverage.
                for (key, flag) in [
                    ("stopIfTrue", rule.stop_if_true),
                    ("aboveAverage", rule.above_average),
                    ("percent", rule.percent),
                    ("bottom", rule.bottom),
                ] {
                    if let Some(b) = flag {
                        rule_elem.push_attribute((key, if b { "1" } else { "0" }));
                    }
                }
                if let Some(ref op) = rule.operator {
                    rule_elem.push_attribute(xml_attr("operator", op.as_str()));
                }
                if let Some(ref text) = rule.text {
                    rule_elem.push_attribute(xml_attr("text", text.as_str()));
                }
                if let Some(ref period) = rule.time_period {
                    rule_elem.push_attribute(xml_attr("timePeriod", period.as_str()));
                }
                if let Some(rank) = rule.rank {
                    rule_elem.push_attribute(("rank", ibuf.format(rank)));
                }
                if let Some(std_dev) = rule.std_dev {
                    rule_elem.push_attribute(("stdDev", ibuf.format(std_dev)));
                }
                if let Some(b) = rule.equal_average {
                    rule_elem.push_attribute(("equalAverage", if b { "1" } else { "0" }));
                }

                let has_children = rule.formula.is_some()
                    || !rule.additional_formulas.is_empty()
                    || rule.color_scale.is_some()
                    || rule.data_bar.is_some()
                    || rule.icon_set.is_some();

                if has_children {
                    writer
                        .write_event(Event::Start(rule_elem))
                        .map_err(xml_write_err)?;

                    // `formula`, then each of `additional_formulas`, in order.
                    for formula in rule.formula.iter().chain(&rule.additional_formulas) {
                        writer
                            .write_event(Event::Start(BytesStart::new("formula")))
                            .map_err(xml_write_err)?;
                        writer
                            .write_event(Event::Text(BytesText::new(
                                crate::ooxml::sanitize_xml_text(formula).as_ref(),
                            )))
                            .map_err(xml_write_err)?;
                        writer
                            .write_event(Event::End(BytesEnd::new("formula")))
                            .map_err(xml_write_err)?;
                    }

                    if let Some(ref cs) = rule.color_scale {
                        writer
                            .write_event(Event::Start(BytesStart::new("colorScale")))
                            .map_err(xml_write_err)?;
                        write_cfvos(&mut writer, &cs.cfvos)?;
                        // One <color> per stop; a stop with no colour (or an
                        // all-default one) gets Excel's default bar colour.
                        for i in 0..cs.colors.len().max(cs.color_specs.len()) {
                            let color = resolve_color(
                                cs.colors.get(i).map(String::as_str),
                                cs.color_specs.get(i),
                            )
                            .filter(|c| !c.is_default_spec())
                            .unwrap_or(ColorOut::Rgb(DEFAULT_BAR_RGB));
                            write_color(&mut writer, "color", color, &mut xsd)?;
                        }
                        writer
                            .write_event(Event::End(BytesEnd::new("colorScale")))
                            .map_err(xml_write_err)?;
                    }

                    if let Some(ref db) = rule.data_bar {
                        writer
                            .write_event(Event::Start(BytesStart::new("dataBar")))
                            .map_err(xml_write_err)?;
                        write_cfvos(&mut writer, &db.cfvos)?;
                        let color = resolve_color(Some(&db.color), db.color_spec.as_ref())
                            .filter(|c| !c.is_default_spec())
                            .unwrap_or(ColorOut::Rgb(DEFAULT_BAR_RGB));
                        write_color(&mut writer, "color", color, &mut xsd)?;
                        writer
                            .write_event(Event::End(BytesEnd::new("dataBar")))
                            .map_err(xml_write_err)?;
                    }

                    if let Some(ref is) = rule.icon_set {
                        let mut is_elem = BytesStart::new("iconSet");
                        if let Some(ref ist) = is.icon_set_type {
                            is_elem.push_attribute(xml_attr("iconSet", ist.as_str()));
                        }
                        writer
                            .write_event(Event::Start(is_elem))
                            .map_err(xml_write_err)?;
                        write_cfvos(&mut writer, &is.cfvos)?;
                        writer
                            .write_event(Event::End(BytesEnd::new("iconSet")))
                            .map_err(xml_write_err)?;
                    }

                    writer
                        .write_event(Event::End(BytesEnd::new("cfRule")))
                        .map_err(xml_write_err)?;
                } else {
                    writer
                        .write_event(Event::Empty(rule_elem))
                        .map_err(xml_write_err)?;
                }
            }

            writer
                .write_event(Event::End(BytesEnd::new("conditionalFormatting")))
                .map_err(xml_write_err)?;
        }

        // <dataValidations> — only if non-empty.
        if !self.data_validations.is_empty() {
            let mut dvs = BytesStart::new("dataValidations");
            dvs.push_attribute(("count", ibuf.format(self.data_validations.len())));
            writer
                .write_event(Event::Start(dvs))
                .map_err(xml_write_err)?;

            for dv in &self.data_validations {
                let mut elem = BytesStart::new("dataValidation");
                if let Some(ref t) = dv.validation_type {
                    elem.push_attribute(xml_attr("type", t.as_str()));
                }
                if let Some(ref op) = dv.operator {
                    elem.push_attribute(xml_attr("operator", op.as_str()));
                }
                if let Some(ref style) = dv.error_style {
                    elem.push_attribute(xml_attr("errorStyle", style.as_str()));
                }
                if let Some(ref ime) = dv.ime_mode {
                    elem.push_attribute(xml_attr("imeMode", ime.as_str()));
                }
                if let Some(ab) = dv.allow_blank {
                    elem.push_attribute(("allowBlank", if ab { "1" } else { "0" }));
                }
                if let Some(sdd) = dv.show_drop_down {
                    elem.push_attribute(("showDropDown", if sdd { "1" } else { "0" }));
                }
                if let Some(sem) = dv.show_error_message {
                    elem.push_attribute(("showErrorMessage", if sem { "1" } else { "0" }));
                }
                if let Some(ref et) = dv.error_title {
                    elem.push_attribute(xml_attr("errorTitle", et.as_str()));
                }
                if let Some(ref em) = dv.error_message {
                    elem.push_attribute(xml_attr("error", em.as_str()));
                }
                if let Some(sim) = dv.show_input_message {
                    elem.push_attribute(("showInputMessage", if sim { "1" } else { "0" }));
                }
                if let Some(ref pt) = dv.prompt_title {
                    elem.push_attribute(xml_attr("promptTitle", pt.as_str()));
                }
                if let Some(ref p) = dv.prompt {
                    elem.push_attribute(xml_attr("prompt", p.as_str()));
                }
                elem.push_attribute(xml_attr("sqref", dv.sqref.as_str()));

                let has_formulas = dv.formula1.is_some() || dv.formula2.is_some();
                if has_formulas {
                    writer
                        .write_event(Event::Start(elem))
                        .map_err(xml_write_err)?;

                    if let Some(ref f1) = dv.formula1 {
                        writer
                            .write_event(Event::Start(BytesStart::new("formula1")))
                            .map_err(xml_write_err)?;
                        writer
                            .write_event(Event::Text(BytesText::new(
                                crate::ooxml::sanitize_xml_text(f1).as_ref(),
                            )))
                            .map_err(xml_write_err)?;
                        writer
                            .write_event(Event::End(BytesEnd::new("formula1")))
                            .map_err(xml_write_err)?;
                    }
                    if let Some(ref f2) = dv.formula2 {
                        writer
                            .write_event(Event::Start(BytesStart::new("formula2")))
                            .map_err(xml_write_err)?;
                        writer
                            .write_event(Event::Text(BytesText::new(
                                crate::ooxml::sanitize_xml_text(f2).as_ref(),
                            )))
                            .map_err(xml_write_err)?;
                        writer
                            .write_event(Event::End(BytesEnd::new("formula2")))
                            .map_err(xml_write_err)?;
                    }

                    writer
                        .write_event(Event::End(BytesEnd::new("dataValidation")))
                        .map_err(xml_write_err)?;
                } else {
                    writer
                        .write_event(Event::Empty(elem))
                        .map_err(xml_write_err)?;
                }
            }

            writer
                .write_event(Event::End(BytesEnd::new("dataValidations")))
                .map_err(xml_write_err)?;
        }

        // <hyperlinks> — only if non-empty. External targets live in the sheet
        // rels (TargetMode="External") as `r:id`; `location` is an in-workbook ref or fragment.
        if !self.hyperlinks.is_empty() {
            writer
                .write_event(Event::Start(BytesStart::new("hyperlinks")))
                .map_err(xml_write_err)?;
            for (i, hl) in self.hyperlinks.iter().enumerate() {
                let mut elem = BytesStart::new("hyperlink");
                elem.push_attribute(xml_attr("ref", hl.cell_ref.as_str()));
                if let Some(Some(r_id)) = hyperlink_r_ids.get(i) {
                    elem.push_attribute(xml_attr("r:id", r_id));
                }
                if let Some(loc) = hl.written_location() {
                    elem.push_attribute(xml_attr("location", loc));
                }
                if let Some(ref disp) = hl.display {
                    elem.push_attribute(xml_attr("display", disp.as_str()));
                }
                if let Some(ref tt) = hl.tooltip {
                    elem.push_attribute(xml_attr("tooltip", tt.as_str()));
                }
                writer
                    .write_event(Event::Empty(elem))
                    .map_err(xml_write_err)?;
            }
            writer
                .write_event(Event::End(BytesEnd::new("hyperlinks")))
                .map_err(xml_write_err)?;
        }

        // <pageMargins> — all six attributes are required by CT_PageMargins;
        // a side the model leaves unset gets Excel's Normal preset value.
        if let Some(pm) = self.page_setup.as_ref().and_then(|ps| ps.margins.as_ref()) {
            let mut elem = BytesStart::new("pageMargins");
            for (name, value, default) in [
                ("left", pm.left, 0.7),
                ("right", pm.right, 0.7),
                ("top", pm.top, 0.75),
                ("bottom", pm.bottom, 0.75),
                ("header", pm.header, 0.3),
                ("footer", pm.footer, 0.3),
            ] {
                elem.push_attribute((
                    name,
                    fmt_xsd_double(value.unwrap_or(default), &mut xsd, "pageMargins")?,
                ));
            }
            writer
                .write_event(Event::Empty(elem))
                .map_err(xml_write_err)?;
        }

        // <pageSetup> — only when a non-margin attribute is set.
        if let Some(ref ps) = self.page_setup
            && (ps.paper_size.is_some()
                || ps.orientation.is_some()
                || ps.fit_to_width.is_some()
                || ps.fit_to_height.is_some()
                || ps.scale.is_some()
                || ps.first_page_number.is_some()
                || ps.horizontal_dpi.is_some()
                || ps.vertical_dpi.is_some())
        {
            let mut elem = BytesStart::new("pageSetup");
            if let Some(paper) = ps.paper_size {
                elem.push_attribute(("paperSize", ibuf.format(paper)));
            }
            if let Some(ref orient) = ps.orientation {
                elem.push_attribute(xml_attr("orientation", orient.as_str()));
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
            writer
                .write_event(Event::Empty(elem))
                .map_err(xml_write_err)?;
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
            writer
                .write_event(Event::Start(elem))
                .map_err(xml_write_err)?;

            for (tag, val) in [
                ("oddHeader", &hf.odd_header),
                ("oddFooter", &hf.odd_footer),
                ("evenHeader", &hf.even_header),
                ("evenFooter", &hf.even_footer),
                ("firstHeader", &hf.first_header),
                ("firstFooter", &hf.first_footer),
            ] {
                if let Some(text) = val {
                    writer
                        .write_event(Event::Start(BytesStart::new(tag)))
                        .map_err(xml_write_err)?;
                    writer
                        .write_event(Event::Text(BytesText::new(
                            crate::ooxml::sanitize_xml_text(text).as_ref(),
                        )))
                        .map_err(xml_write_err)?;
                    writer
                        .write_event(Event::End(BytesEnd::new(tag)))
                        .map_err(xml_write_err)?;
                }
            }

            writer
                .write_event(Event::End(BytesEnd::new("headerFooter")))
                .map_err(xml_write_err)?;
        }

        // <rowBreaks> / <colBreaks> — only if page breaks are present.
        if let Some(ref pb) = self.page_breaks {
            if !pb.row_breaks.is_empty() {
                write_breaks(&mut writer, "rowBreaks", &pb.row_breaks, &mut ibuf)?;
            }
            if !pb.col_breaks.is_empty() {
                write_breaks(&mut writer, "colBreaks", &pb.col_breaks, &mut ibuf)?;
            }
        }

        // <drawing>, <legacyDrawing> and <legacyDrawingHF r:id="..."/>, in the
        // order CT_Worksheet gives them, all before <tableParts>.
        for (tag, r_id) in [
            ("drawing", drawing_r_id),
            ("legacyDrawing", legacy_drawing_r_id),
            ("legacyDrawingHF", legacy_drawing_hf_r_id),
        ] {
            if let Some(r_id) = r_id {
                let mut elem = BytesStart::new(tag);
                elem.push_attribute(("r:id", r_id));
                writer
                    .write_event(Event::Empty(elem))
                    .map_err(xml_write_err)?;
            }
        }

        // <tableParts> — only if table rIds are provided.
        if !table_r_ids.is_empty() {
            let mut tp = BytesStart::new("tableParts");
            tp.push_attribute(("count", ibuf.format(table_r_ids.len())));
            writer
                .write_event(Event::Start(tp))
                .map_err(xml_write_err)?;
            for rid in table_r_ids {
                let mut part = BytesStart::new("tablePart");
                part.push_attribute(xml_attr("r:id", rid.as_str()));
                writer
                    .write_event(Event::Empty(part))
                    .map_err(xml_write_err)?;
            }
            writer
                .write_event(Event::End(BytesEnd::new("tableParts")))
                .map_err(xml_write_err)?;
        }

        // <extLst> — sparkline groups and/or preserved extensions.
        if !self.sparkline_groups.is_empty() || !self.preserved_extensions.is_empty() {
            writer
                .write_event(Event::Start(BytesStart::new("extLst")))
                .map_err(xml_write_err)?;

            // Write sparkline extension.
            if !self.sparkline_groups.is_empty() {
                let mut ext = BytesStart::new("ext");
                ext.push_attribute(("uri", super::SPARKLINE_EXT_URI));
                ext.push_attribute((
                    "xmlns:x14",
                    "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main",
                ));
                writer
                    .write_event(Event::Start(ext))
                    .map_err(xml_write_err)?;

                let mut groups_elem = BytesStart::new("x14:sparklineGroups");
                groups_elem.push_attribute((
                    "xmlns:xm",
                    "http://schemas.microsoft.com/office/excel/2006/main",
                ));
                writer
                    .write_event(Event::Start(groups_elem))
                    .map_err(xml_write_err)?;

                for group in &self.sparkline_groups {
                    let mut group_elem = BytesStart::new("x14:sparklineGroup");

                    // Type attribute (omit if "line" which is default).
                    if group.sparkline_type != "line" {
                        group_elem.push_attribute(xml_attr("type", group.sparkline_type.as_str()));
                    }
                    if group.date_axis {
                        group_elem.push_attribute(("dateAxis", "1"));
                    }

                    // Boolean attributes (only write if true, as "1").
                    if group.markers {
                        group_elem.push_attribute(("markers", "1"));
                    }
                    if group.high {
                        group_elem.push_attribute(("high", "1"));
                    }
                    if group.low {
                        group_elem.push_attribute(("low", "1"));
                    }
                    if group.first {
                        group_elem.push_attribute(("first", "1"));
                    }
                    if group.last {
                        group_elem.push_attribute(("last", "1"));
                    }
                    if group.negative {
                        group_elem.push_attribute(("negative", "1"));
                    }
                    if group.display_x_axis {
                        group_elem.push_attribute(("displayXAxis", "1"));
                    }
                    if group.display_hidden {
                        group_elem.push_attribute(("displayHidden", "1"));
                    }
                    if group.right_to_left {
                        group_elem.push_attribute(("rightToLeft", "1"));
                    }

                    // Optional string/numeric attributes.
                    if let Some(ref d) = group.display_empty_cells_as {
                        group_elem.push_attribute(xml_attr("displayEmptyCellsAs", d.as_str()));
                    }
                    if let Some(w) = group.line_weight {
                        group_elem.push_attribute((
                            "lineWeight",
                            fmt_xsd_double(w, &mut xsd, "sparklineGroup lineWeight")?,
                        ));
                    }
                    if let Some(v) = group.manual_min {
                        group_elem.push_attribute((
                            "manualMin",
                            fmt_xsd_double(v, &mut xsd, "sparklineGroup manualMin")?,
                        ));
                    }
                    if let Some(v) = group.manual_max {
                        group_elem.push_attribute((
                            "manualMax",
                            fmt_xsd_double(v, &mut xsd, "sparklineGroup manualMax")?,
                        ));
                    }
                    // A manual bound applies only on a "custom" axis, so a
                    // bound without an axis type implies one (as XlsxWriter writes).
                    for (key, explicit, manual) in [
                        ("minAxisType", &group.min_axis_type, group.manual_min),
                        ("maxAxisType", &group.max_axis_type, group.manual_max),
                    ] {
                        let custom = manual.map(|_| "custom");
                        if let Some(t) = explicit.as_deref().or(custom) {
                            group_elem.push_attribute(xml_attr(key, t));
                        }
                    }

                    writer
                        .write_event(Event::Start(group_elem))
                        .map_err(xml_write_err)?;

                    // Colour elements (self-closing), each from its full colour
                    // while that matches the rgb string, else as plain rgb.
                    for (tag, rgb, spec) in [
                        (
                            "x14:colorSeries",
                            &group.color_series,
                            &group.color_series_spec,
                        ),
                        (
                            "x14:colorNegative",
                            &group.color_negative,
                            &group.color_negative_spec,
                        ),
                        ("x14:colorAxis", &group.color_axis, &group.color_axis_spec),
                        (
                            "x14:colorMarkers",
                            &group.color_markers,
                            &group.color_markers_spec,
                        ),
                        (
                            "x14:colorFirst",
                            &group.color_first,
                            &group.color_first_spec,
                        ),
                        ("x14:colorLast", &group.color_last, &group.color_last_spec),
                        ("x14:colorHigh", &group.color_high, &group.color_high_spec),
                        ("x14:colorLow", &group.color_low, &group.color_low_spec),
                    ] {
                        if let Some(color) = resolve_color(rgb.as_deref(), spec.as_ref())
                            .filter(|c| !c.is_default_spec())
                        {
                            write_color(&mut writer, tag, color, &mut xsd)?;
                        }
                    }

                    // The date axis range, between the colours and the sparklines.
                    if let Some(ref dates) = group.date_range {
                        writer
                            .write_event(Event::Start(BytesStart::new("xm:f")))
                            .map_err(xml_write_err)?;
                        writer
                            .write_event(Event::Text(BytesText::new(
                                crate::ooxml::sanitize_xml_text(dates).as_ref(),
                            )))
                            .map_err(xml_write_err)?;
                        writer
                            .write_event(Event::End(BytesEnd::new("xm:f")))
                            .map_err(xml_write_err)?;
                    }

                    // Write sparklines.
                    writer
                        .write_event(Event::Start(BytesStart::new("x14:sparklines")))
                        .map_err(xml_write_err)?;
                    for sparkline in &group.sparklines {
                        writer
                            .write_event(Event::Start(BytesStart::new("x14:sparkline")))
                            .map_err(xml_write_err)?;

                        writer
                            .write_event(Event::Start(BytesStart::new("xm:f")))
                            .map_err(xml_write_err)?;
                        writer
                            .write_event(Event::Text(BytesText::new(
                                crate::ooxml::sanitize_xml_text(&sparkline.formula).as_ref(),
                            )))
                            .map_err(xml_write_err)?;
                        writer
                            .write_event(Event::End(BytesEnd::new("xm:f")))
                            .map_err(xml_write_err)?;

                        writer
                            .write_event(Event::Start(BytesStart::new("xm:sqref")))
                            .map_err(xml_write_err)?;
                        writer
                            .write_event(Event::Text(BytesText::new(
                                crate::ooxml::sanitize_xml_text(&sparkline.sqref).as_ref(),
                            )))
                            .map_err(xml_write_err)?;
                        writer
                            .write_event(Event::End(BytesEnd::new("xm:sqref")))
                            .map_err(xml_write_err)?;

                        writer
                            .write_event(Event::End(BytesEnd::new("x14:sparkline")))
                            .map_err(xml_write_err)?;
                    }
                    writer
                        .write_event(Event::End(BytesEnd::new("x14:sparklines")))
                        .map_err(xml_write_err)?;

                    writer
                        .write_event(Event::End(BytesEnd::new("x14:sparklineGroup")))
                        .map_err(xml_write_err)?;
                }

                writer
                    .write_event(Event::End(BytesEnd::new("x14:sparklineGroups")))
                    .map_err(xml_write_err)?;
                writer
                    .write_event(Event::End(BytesEnd::new("ext")))
                    .map_err(xml_write_err)?;
            }

            // Write preserved non-sparkline extensions.
            for ext_xml in &self.preserved_extensions {
                writer.get_mut().extend_from_slice(ext_xml.as_bytes());
            }

            writer
                .write_event(Event::End(BytesEnd::new("extLst")))
                .map_err(xml_write_err)?;
        }

        // </worksheet>
        writer
            .write_event(Event::End(BytesEnd::new("worksheet")))
            .map_err(xml_write_err)?;

        Ok(buf)
    }
}

#[cfg(test)]
mod frozen_pane_tests {
    use super::super::WorksheetXml;

    fn sheet_with_pane(rows: u32, cols: u32) -> WorksheetXml {
        serde_json::from_str(&format!(
            r#"{{"rows":[],"frozenPane":{{"rows":{rows},"cols":{cols}}}}}"#
        ))
        .unwrap()
    }

    #[test]
    fn frozen_pane_past_xfd_is_an_error() {
        let err = sheet_with_pane(0, 20_000).to_xml().unwrap_err();
        assert_eq!(err.code(), "INVALID_CELL_REF");
        let err = sheet_with_pane(1_048_576, 0).to_xml().unwrap_err();
        assert_eq!(err.code(), "INVALID_CELL_REF");
    }

    #[test]
    fn frozen_pane_at_the_edge_is_written() {
        let xml = String::from_utf8(sheet_with_pane(0, 16_383).to_xml().unwrap()).unwrap();
        assert!(xml.contains("topLeftCell=\"XFD1\""), "{xml}");
    }
}

#[cfg(test)]
mod xsd_double_attr_tests {
    use super::super::WorksheetXml;

    fn sheet() -> WorksheetXml {
        serde_json::from_str(
            r#"{"rows":[],"columns":[{"min":1,"max":1,"width":1e20,"hidden":false,"customWidth":true},{"min":2,"max":2,"width":15.0,"hidden":false,"customWidth":true}],"splitPane":{"horizontal":-0.0,"vertical":8.43}}"#,
        )
        .unwrap()
    }

    #[test]
    fn widths_and_splits_round_trip_as_xsd_double() {
        let xml = String::from_utf8(sheet().to_xml().unwrap()).unwrap();
        // 1e20 is written as an xsd:double, not saturated to i64::MAX (9223372036854775807).
        assert!(xml.contains(r#"width="1e+20""#), "{xml}");
        assert!(xml.contains(r#"width="15""#), "{xml}");
        assert!(xml.contains(r#"xSplit="8.43""#), "{xml}");
        assert!(xml.contains(r#"ySplit="-0""#), "{xml}");
    }

    #[test]
    fn non_finite_width_is_a_write_error() {
        let mut ws = sheet();
        ws.columns[0].width = f64::INFINITY;
        let err = ws.to_xml().unwrap_err();
        assert_eq!(err.code(), "XML_WRITE");
        assert!(err.to_string().contains("col width"), "{err}");
    }
}

#[cfg(test)]
mod schema_tests {
    use super::super::{Cell, CellType, Row, WorksheetXml};

    /// A worksheet with eleven optional children, out of `CT_Worksheet` order.
    const FULL: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>
<tableParts count="0"/>
<headerFooter><oddHeader>h</oddHeader></headerFooter>
<pageSetup orientation="landscape"/>
<pageMargins left="1" right="1" top="1" bottom="1" header="0.5" footer="0.5"/>
<rowBreaks count="1" manualBreakCount="1"><brk id="3" man="1"/></rowBreaks>
<hyperlinks><hyperlink ref="A1" location="Sheet2!A1"/></hyperlinks>
<dataValidations count="1"><dataValidation type="list" sqref="A1"><formula1>"a,b"</formula1></dataValidation></dataValidations>
<conditionalFormatting sqref="A1"><cfRule type="expression" priority="1"><formula>A1&gt;1</formula></cfRule></conditionalFormatting>
<mergeCells count="1"><mergeCell ref="A1:B1"/></mergeCells>
<autoFilter ref="A1:B2"/>
<sheetProtection sheet="1"/>
</worksheet>"#;

    fn text(bytes: Vec<u8>) -> String {
        String::from_utf8(bytes).unwrap()
    }

    fn cell(reference: &str) -> Cell {
        Cell {
            reference: reference.into(),
            value: Some("1".into()),
            ..Default::default()
        }
    }

    fn sheet_with(cells: &[(u32, &str)]) -> WorksheetXml {
        WorksheetXml {
            rows: cells
                .iter()
                .map(|&(index, r)| Row {
                    index,
                    cells: vec![cell(r)],
                    ..Default::default()
                })
                .collect(),
            ..Default::default()
        }
    }

    #[test]
    fn children_follow_ct_worksheet_order() {
        let mut ws = WorksheetXml::parse(FULL.as_bytes()).unwrap();
        ws.preserved_extensions
            .push(r#"<ext uri="{00000000-0000-0000-0000-000000000000}"><foo/></ext>"#.to_owned());
        let xml = text(
            ws.to_xml_with_sst(None, &["rId1".to_owned()], Some("rId2"), &[])
                .unwrap(),
        );
        let order = [
            "<dimension ",
            "<sheetData>",
            "<sheetProtection ",
            "<autoFilter ",
            "<mergeCells ",
            "<conditionalFormatting ",
            "<dataValidations ",
            "<hyperlinks>",
            "<pageMargins ",
            "<pageSetup ",
            "<headerFooter",
            "<rowBreaks ",
            "<drawing ",
            "<tableParts ",
            "<extLst>",
        ];
        let mut last = 0;
        for tag in order {
            let at = xml
                .find(tag)
                .unwrap_or_else(|| panic!("{tag} missing in {xml}"));
            assert!(at > last || last == 0, "{tag} out of order in {xml}");
            last = at;
        }
    }

    #[test]
    fn dv_formulas_are_escaped_and_round_trip() {
        let xml = r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/><dataValidations count="1"><dataValidation type="custom" sqref="A1"><formula1>"R&amp;D,Ops"</formula1><formula2>A1&lt;10</formula2></dataValidation></dataValidations></worksheet>"#;
        let ws = WorksheetXml::parse(xml.as_bytes()).unwrap();
        let out = text(ws.to_xml().unwrap());
        assert!(out.contains("R&amp;D"), "{out}");
        assert!(out.contains("A1&lt;10"), "{out}");
        let back = WorksheetXml::parse(out.as_bytes()).unwrap();
        let dv = &back.data_validations[0];
        assert_eq!(dv.formula1.as_deref(), Some("\"R&D,Ops\""));
        assert_eq!(dv.formula2.as_deref(), Some("A1<10"));
    }

    #[test]
    fn inline_string_writes_only_is_with_preserve() {
        let ws = WorksheetXml {
            rows: vec![Row {
                index: 1,
                cells: vec![Cell {
                    reference: "A1".into(),
                    cell_type: CellType::InlineStr,
                    inline_string: Some("  pad  ".into()),
                    value: Some("  pad  ".into()),
                    ..Default::default()
                }],
                ..Default::default()
            }],
            ..Default::default()
        };
        let out = text(ws.to_xml().unwrap());
        assert!(
            out.contains(
                r#"<c r="A1" t="inlineStr"><is><t xml:space="preserve">  pad  </t></is></c>"#
            ),
            "{out}"
        );
        let back = WorksheetXml::parse(out.as_bytes()).unwrap();
        let c = &back.rows[0].cells[0];
        assert_eq!(c.value.as_deref(), Some("  pad  "));
        assert_eq!(c.inline_string.as_deref(), Some("  pad  "));
    }

    #[test]
    fn inline_string_falls_back_to_value_and_formula_stays_first() {
        let ws = WorksheetXml {
            rows: vec![Row {
                index: 1,
                cells: vec![Cell {
                    reference: "A1".into(),
                    cell_type: CellType::InlineStr,
                    value: Some("x".into()),
                    formula: Some("B1".into()),
                    ..Default::default()
                }],
                ..Default::default()
            }],
            ..Default::default()
        };
        let out = text(ws.to_xml().unwrap());
        assert!(
            out.contains(r#"<c r="A1" t="inlineStr"><f>B1</f><is><t>x</t></is></c>"#),
            "{out}"
        );
    }

    #[test]
    fn attribute_control_characters_are_removed() {
        let xml = r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/><hyperlinks><hyperlink ref="A1" location="S!A1"/></hyperlinks></worksheet>"#;
        let mut ws = WorksheetXml::parse(xml.as_bytes()).unwrap();
        ws.hyperlinks[0].display = Some("a\u{1}b".into());
        ws.hyperlinks[0].tooltip = Some("t\u{8}t".into());
        let bytes = ws.to_xml().unwrap();
        assert!(
            bytes
                .iter()
                .all(|b| !matches!(b, 0x00..=0x08 | 0x0B | 0x0C | 0x0E..=0x1F)),
            "{}",
            text(bytes.clone())
        );
        let back = WorksheetXml::parse(&bytes).unwrap();
        assert_eq!(back.hyperlinks[0].display.as_deref(), Some("ab"));
        assert_eq!(back.hyperlinks[0].tooltip.as_deref(), Some("tt"));
    }

    #[test]
    fn dimension_is_computed_from_the_rows() {
        let xml = text(sheet_with(&[(1, "A1"), (5, "C5")]).to_xml().unwrap());
        let dim = xml.find(r#"<dimension ref="A1:C5"/>"#).expect(&xml);
        assert!(dim < xml.find("<sheetData>").unwrap());

        let mut ws = sheet_with(&[(1, "A1"), (5, "C5")]);
        ws.sheet_view = Some(Default::default());
        ws.dimension = Some("Z9:Z99".into());
        let xml = text(ws.to_xml().unwrap());
        assert!(xml.contains(r#"<dimension ref="A1:C5"/>"#), "{xml}");
        assert!(xml.find("<dimension ") < xml.find("<sheetViews>"), "{xml}");

        let xml = text(WorksheetXml::default().to_xml().unwrap());
        assert!(xml.contains(r#"<dimension ref="A1"/>"#), "{xml}");
        let xml = text(sheet_with(&[(3, "B3")]).to_xml().unwrap());
        assert!(xml.contains(r#"<dimension ref="B3"/>"#), "{xml}");
    }

    #[test]
    fn stub_cells_write_bare_c_elements() {
        let ws = WorksheetXml {
            rows: vec![Row {
                index: 1,
                cells: vec![
                    Cell {
                        reference: "A1".into(),
                        cell_type: CellType::Stub,
                        ..Default::default()
                    },
                    Cell {
                        reference: "B1".into(),
                        cell_type: CellType::Stub,
                        style_index: Some(3),
                        ..Default::default()
                    },
                ],
                ..Default::default()
            }],
            ..Default::default()
        };
        let xml = text(ws.to_xml().unwrap());
        assert!(xml.contains(r#"<c r="A1"/><c r="B1" s="3"/>"#), "{xml}");
        assert!(xml.contains(r#"<dimension ref="A1"/>"#), "{xml}");
    }

    #[test]
    fn row_capacity_is_capped() {
        let mut data = b"<worksheet><sheetData/></worksheet>".to_vec();
        data.resize(20_000_000, b' ');
        let ws = WorksheetXml::parse(&data).unwrap();
        assert!(ws.rows.capacity() <= 1 << 16, "{}", ws.rows.capacity());
    }

    #[test]
    fn page_margins_round_trip_before_page_setup() {
        let xml = r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/><pageMargins left="1" right="1" top="1.5" bottom="1.5" header="0.5" footer="0.5"/><pageSetup orientation="landscape"/></worksheet>"#;
        let ws = WorksheetXml::parse(xml.as_bytes()).unwrap();
        let ps = ws.page_setup.as_ref().unwrap();
        assert_eq!(ps.orientation.as_deref(), Some("landscape"));
        let m = ps.margins.as_ref().unwrap();
        assert_eq!((m.top, m.footer), (Some(1.5), Some(0.5)));
        let out = text(ws.to_xml().unwrap());
        assert!(
            out.contains(
                r#"<pageMargins left="1" right="1" top="1.5" bottom="1.5" header="0.5" footer="0.5"/><pageSetup orientation="landscape"/>"#
            ),
            "{out}"
        );
        let back = WorksheetXml::parse(out.as_bytes()).unwrap();
        assert_eq!(back.page_setup.unwrap().margins, ps.margins);
    }

    #[test]
    fn margins_alone_write_no_page_setup_and_fill_defaults() {
        let ws: WorksheetXml =
            serde_json::from_str(r#"{"rows":[],"pageSetup":{"margins":{"top":2}}}"#).unwrap();
        let out = text(ws.to_xml().unwrap());
        assert!(
            out.contains(
                r#"<pageMargins left="0.7" right="0.7" top="2" bottom="0.75" header="0.3" footer="0.3"/>"#
            ),
            "{out}"
        );
        assert!(!out.contains("<pageSetup"), "{out}");
    }
}
