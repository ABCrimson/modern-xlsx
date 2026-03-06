use quick_xml::events::{BytesDecl, BytesEnd, BytesStart, BytesText, Event};
use quick_xml::Writer;

use super::{
    CellType, WorksheetXml,
};
use crate::ooxml::SPREADSHEET_NS;
use crate::{ModernXlsxError, Result};

impl WorksheetXml {
    /// Serialize this worksheet to valid XML bytes.
    pub fn to_xml(&self) -> Result<Vec<u8>> {
        self.to_xml_with_sst(None, &[], None)
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
    pub fn to_xml_with_sst(
        &self,
        sst: Option<&crate::ooxml::shared_strings::SharedStringTableBuilder>,
        table_r_ids: &[String],
        drawing_r_id: Option<&str>,
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

        // <sheetViews> — if sheet_view, frozen_pane, or split_pane is present.
        let has_sheet_views = self.sheet_view.is_some()
            || self.split_pane.is_some()
            || self.frozen_pane.is_some();
        if has_sheet_views {
            writer
                .write_event(Event::Start(BytesStart::new("sheetViews")))
                .map_err(map_err)?;

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
                    sv_elem.push_attribute(("view", v.as_str()));
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
                    .map_err(map_err)?;

                // Write pane element (split takes priority over frozen).
                if let Some(ref sp) = self.split_pane {
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
                    writer
                        .write_event(Event::Empty(pane_elem))
                        .map_err(map_err)?;
                } else if let Some(ref pane) = self.frozen_pane {
                    let mut pane_elem = BytesStart::new("pane");
                    if pane.cols > 0 {
                        pane_elem.push_attribute(("xSplit", ibuf.format(pane.cols)));
                    }
                    if pane.rows > 0 {
                        pane_elem.push_attribute(("ySplit", ibuf.format(pane.rows)));
                    }
                    let mut top_left = col_index_to_letter(pane.cols.saturating_add(1));
                    top_left.push_str(ibuf.format(pane.rows.saturating_add(1)));
                    pane_elem.push_attribute(("topLeftCell", top_left.as_str()));
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
                        .map_err(map_err)?;
                }

                // Write <selection> elements.
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
                    writer
                        .write_event(Event::Empty(sel_elem))
                        .map_err(map_err)?;
                }

                writer
                    .write_event(Event::End(BytesEnd::new("sheetView")))
                    .map_err(map_err)?;
            } else {
                // No children — write self-closing <sheetView ... />
                writer
                    .write_event(Event::Empty(sv_elem))
                    .map_err(map_err)?;
            }

            writer
                .write_event(Event::End(BytesEnd::new("sheetViews")))
                .map_err(map_err)?;
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
                            if let Some(ref r1) = cell.formula_r1 {
                                f_elem.push_attribute(("r1", r1.as_str()));
                            }
                            if let Some(ref r2) = cell.formula_r2 {
                                f_elem.push_attribute(("r2", r2.as_str()));
                            }
                            if cell.formula_dt2d == Some(true) {
                                f_elem.push_attribute(("dt2D", "1"));
                            }
                            if cell.formula_dtr1 == Some(true) {
                                f_elem.push_attribute(("dtr1", "1"));
                            }
                            if cell.formula_dtr2 == Some(true) {
                                f_elem.push_attribute(("dtr2", "1"));
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
                            if let Some(ref r1) = cell.formula_r1 {
                                f_elem.push_attribute(("r1", r1.as_str()));
                            }
                            if let Some(ref r2) = cell.formula_r2 {
                                f_elem.push_attribute(("r2", r2.as_str()));
                            }
                            if cell.formula_dt2d == Some(true) {
                                f_elem.push_attribute(("dt2D", "1"));
                            }
                            if cell.formula_dtr1 == Some(true) {
                                f_elem.push_attribute(("dtr1", "1"));
                            }
                            if cell.formula_dtr2 == Some(true) {
                                f_elem.push_attribute(("dtr2", "1"));
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
                                    // inline_string presence guaranteed by has_inline guard above
                                    cell.inline_string.as_deref().unwrap_or_default(),
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

                    if let Some(ref cf) = fc.custom_filters {
                        // <customFilters>
                        let mut cf_elem = BytesStart::new("customFilters");
                        if cf.and_op {
                            cf_elem.push_attribute(("and", "1"));
                        }
                        writer.write_event(Event::Start(cf_elem)).map_err(map_err)?;
                        for item in &cf.filters {
                            let mut item_elem = BytesStart::new("customFilter");
                            if let Some(ref op) = item.operator {
                                item_elem.push_attribute(("operator", op.as_str()));
                            }
                            item_elem.push_attribute(("val", item.val.as_str()));
                            writer.write_event(Event::Empty(item_elem)).map_err(map_err)?;
                        }
                        writer
                            .write_event(Event::End(BytesEnd::new("customFilters")))
                            .map_err(map_err)?;
                    } else {
                        // <filters>
                        writer
                            .write_event(Event::Start(BytesStart::new("filters")))
                            .map_err(map_err)?;
                        for fv in &fc.filters {
                            let mut f_elem = BytesStart::new("filter");
                            f_elem.push_attribute(("val", fv.as_str()));
                            writer.write_event(Event::Empty(f_elem)).map_err(map_err)?;
                        }
                        writer
                            .write_event(Event::End(BytesEnd::new("filters")))
                            .map_err(map_err)?;
                    }

                    writer
                        .write_event(Event::End(BytesEnd::new("filterColumn")))
                        .map_err(map_err)?;
                }
                writer
                    .write_event(Event::End(BytesEnd::new("autoFilter")))
                    .map_err(map_err)?;
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

        // <rowBreaks> / <colBreaks> — only if page breaks are present.
        if let Some(ref pb) = self.page_breaks {
            if !pb.row_breaks.is_empty() {
                let mut rb = BytesStart::new("rowBreaks");
                rb.push_attribute(("count", ibuf.format(pb.row_breaks.len())));
                let manual_count = pb.row_breaks.iter().filter(|b| b.man).count();
                rb.push_attribute(("manualBreakCount", ibuf.format(manual_count)));
                writer.write_event(Event::Start(rb)).map_err(map_err)?;
                for brk in &pb.row_breaks {
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
                    writer.write_event(Event::Empty(elem)).map_err(map_err)?;
                }
                writer.write_event(Event::End(BytesEnd::new("rowBreaks"))).map_err(map_err)?;
            }
            if !pb.col_breaks.is_empty() {
                let mut cb = BytesStart::new("colBreaks");
                cb.push_attribute(("count", ibuf.format(pb.col_breaks.len())));
                let manual_count = pb.col_breaks.iter().filter(|b| b.man).count();
                cb.push_attribute(("manualBreakCount", ibuf.format(manual_count)));
                writer.write_event(Event::Start(cb)).map_err(map_err)?;
                for brk in &pb.col_breaks {
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
                    writer.write_event(Event::Empty(elem)).map_err(map_err)?;
                }
                writer.write_event(Event::End(BytesEnd::new("colBreaks"))).map_err(map_err)?;
            }
        }

        // <drawing r:id="..."/> — ECMA-376 schema requires <drawing> before <tableParts>.
        if let Some(rid) = drawing_r_id {
            let mut drawing_elem = BytesStart::new("drawing");
            drawing_elem.push_attribute(("r:id", rid));
            writer.write_event(Event::Empty(drawing_elem)).map_err(map_err)?;
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

        // <extLst> — sparkline groups and/or preserved extensions.
        if !self.sparkline_groups.is_empty() || !self.preserved_extensions.is_empty() {
            writer
                .write_event(Event::Start(BytesStart::new("extLst")))
                .map_err(map_err)?;

            // Write sparkline extension.
            if !self.sparkline_groups.is_empty() {
                let mut ext = BytesStart::new("ext");
                ext.push_attribute(("uri", "{05C60535-1F16-4fd2-B633-F4F36011B0BD}"));
                ext.push_attribute((
                    "xmlns:x14",
                    "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main",
                ));
                writer.write_event(Event::Start(ext)).map_err(map_err)?;

                let mut groups_elem = BytesStart::new("x14:sparklineGroups");
                groups_elem.push_attribute((
                    "xmlns:xm",
                    "http://schemas.microsoft.com/office/excel/2006/main",
                ));
                writer
                    .write_event(Event::Start(groups_elem))
                    .map_err(map_err)?;

                for group in &self.sparkline_groups {
                    let mut group_elem = BytesStart::new("x14:sparklineGroup");

                    // Type attribute (omit if "line" which is default).
                    if group.sparkline_type != "line" {
                        group_elem
                            .push_attribute(("type", group.sparkline_type.as_str()));
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
                    if group.right_to_left {
                        group_elem.push_attribute(("rightToLeft", "1"));
                    }

                    // Optional string/numeric attributes.
                    if let Some(ref d) = group.display_empty_cells_as {
                        group_elem.push_attribute(("displayEmptyCellsAs", d.as_str()));
                    }
                    if let Some(w) = group.line_weight {
                        let s = format!("{w}");
                        group_elem.push_attribute(("lineWeight", s.as_str()));
                    }
                    if let Some(v) = group.manual_min {
                        let s = format!("{v}");
                        group_elem.push_attribute(("manualMin", s.as_str()));
                    }
                    if let Some(v) = group.manual_max {
                        let s = format!("{v}");
                        group_elem.push_attribute(("manualMax", s.as_str()));
                    }

                    writer
                        .write_event(Event::Start(group_elem))
                        .map_err(map_err)?;

                    // Write color elements (self-closing).
                    macro_rules! write_sparkline_color {
                        ($name:expr, $field:expr) => {
                            if let Some(ref c) = $field {
                                let mut e = BytesStart::new($name);
                                e.push_attribute(("rgb", c.as_str()));
                                writer.write_event(Event::Empty(e)).map_err(map_err)?;
                            }
                        };
                    }
                    write_sparkline_color!("x14:colorSeries", group.color_series);
                    write_sparkline_color!("x14:colorNegative", group.color_negative);
                    write_sparkline_color!("x14:colorAxis", group.color_axis);
                    write_sparkline_color!("x14:colorMarkers", group.color_markers);
                    write_sparkline_color!("x14:colorFirst", group.color_first);
                    write_sparkline_color!("x14:colorLast", group.color_last);
                    write_sparkline_color!("x14:colorHigh", group.color_high);
                    write_sparkline_color!("x14:colorLow", group.color_low);

                    // Write sparklines.
                    writer
                        .write_event(Event::Start(BytesStart::new("x14:sparklines")))
                        .map_err(map_err)?;
                    for sparkline in &group.sparklines {
                        writer
                            .write_event(Event::Start(BytesStart::new("x14:sparkline")))
                            .map_err(map_err)?;

                        writer
                            .write_event(Event::Start(BytesStart::new("xm:f")))
                            .map_err(map_err)?;
                        writer
                            .write_event(Event::Text(BytesText::new(&sparkline.formula)))
                            .map_err(map_err)?;
                        writer
                            .write_event(Event::End(BytesEnd::new("xm:f")))
                            .map_err(map_err)?;

                        writer
                            .write_event(Event::Start(BytesStart::new("xm:sqref")))
                            .map_err(map_err)?;
                        writer
                            .write_event(Event::Text(BytesText::new(&sparkline.sqref)))
                            .map_err(map_err)?;
                        writer
                            .write_event(Event::End(BytesEnd::new("xm:sqref")))
                            .map_err(map_err)?;

                        writer
                            .write_event(Event::End(BytesEnd::new("x14:sparkline")))
                            .map_err(map_err)?;
                    }
                    writer
                        .write_event(Event::End(BytesEnd::new("x14:sparklines")))
                        .map_err(map_err)?;

                    writer
                        .write_event(Event::End(BytesEnd::new("x14:sparklineGroup")))
                        .map_err(map_err)?;
                }

                writer
                    .write_event(Event::End(BytesEnd::new("x14:sparklineGroups")))
                    .map_err(map_err)?;
                writer
                    .write_event(Event::End(BytesEnd::new("ext")))
                    .map_err(map_err)?;
            }

            // Write preserved non-sparkline extensions.
            for ext_xml in &self.preserved_extensions {
                writer.get_mut().extend_from_slice(ext_xml.as_bytes());
            }

            writer
                .write_event(Event::End(BytesEnd::new("extLst")))
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
#[inline]
pub(super) fn parse_col_element(e: &quick_xml::events::BytesStart<'_>) -> super::ColumnInfo {
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

    super::ColumnInfo {
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
#[inline]
pub(super) fn col_index_to_letter(col: u32) -> String {
    crate::ooxml::cell::col_to_letters(col.saturating_sub(1))
}

/// Format an f64 to a string, removing trailing zeros after the decimal point
/// but always keeping at least one decimal place if the number is not an integer.
#[inline]
pub(super) fn format_f64(val: f64) -> String {
    if val == val.floor() {
        // Integer value — use itoa to avoid format! overhead.
        itoa::Buffer::new().format(val as i64).to_owned()
    } else {
        ryu::Buffer::new().format(val).to_owned()
    }
}
