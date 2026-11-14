//! The [`RowSink`] behind `WorksheetXml::parse_to_json_within`: [`JsonSink`] writes
//! the rows and cells that the driver reads as JSON text straight into the
//! output buffer, then the rest of the sheet and the closing `}`.

use std::ops::Range;

use super::driver::{CellPart, RowSink, SheetMeta};
use super::json::{cell_type_json_str, json_escape_to, write_f64_json};
use super::parser::{CellAttrs, FormulaState, RowAttrs, iso_date_cell, resolve_sst};
use super::{AutoFilter, CellType, CustomFilters, FilterColumn, PageBreaks};
use crate::Result;
use crate::dates::DateSystem;
use crate::ooxml::comments::Comment;
use crate::ooxml::shared_strings::{ExpansionBudget, SharedStringTable};
use crate::ooxml::tables::TableDefinition;
use crate::ooxml::xml::XsdBuf;
use crate::reader::{check_json_len, clamp_json_growth};

/// Writes a worksheet as one JSON object into `out`: `{"rows":[` from
/// [`JsonSink::new`], each row, cell and child of a cell as the driver reads
/// it, and the remaining keys and the closing `}` from [`JsonSink::finish`].
///
/// A cell's keys follow its children in document order: `reference`,
/// `cellType` and `styleIndex` first; then for each `<v>` its `value` (a
/// SharedString cell's SST entry text and a rich entry's `richText`, a
/// `t="d"` cell's serial), for each `<f>` with text its `formula` and the
/// formula fields not yet written, for each `<is>` its `value` and
/// `inlineString`; last the formula fields not yet written (those of a
/// self-closing `<f/>`, or the `cm` index of a cell without a formula).
/// Repeated children repeat their keys, except `richText`: a `<v>` cuts the
/// runs an earlier `<v>` of the cell wrote back out of `out`, so a cell
/// keeps the runs of its last `<v>` alone, as on the model path (an `<is>`
/// after that `<v>` keeps them, as there).
///
/// A row whose part ends before its `</row>` is cut back out of `out` by
/// [`RowSink::abort_row`], so the JSON stays valid and holds only finished
/// rows. The workbook JSON in `out` is checked against `max_json_length`
/// after each child of a cell, each cell, each row, and each element of the
/// sheet's lists ([`JsonSink::finish`] names them), so a sheet fails as soon
/// as the JSON passes the limit rather than when the sheet is done. A copy
/// of a shared string, comment author or hyperlink target is counted in the
/// read's `ExpansionBudget` at its byte length but written escaped (2 bytes
/// for a backslash or quote, 6 for a control character), so the budget
/// alone does not bound the JSON: these checks hold it to the limit plus one
/// child's keys (a `<v>`'s entry), comment, hyperlink, extension or other
/// list element. Near the end of `out`'s capacity they also keep it from doubling
/// past the limit ([`clamp_json_growth`]).
pub(super) struct JsonSink<'a> {
    out: &'a mut String,
    sst: Option<&'a SharedStringTable>,
    date_system: DateSystem,
    /// The most bytes the workbook JSON in `out` may hold.
    max_json_length: usize,
    /// The length of `out` past which the check after a cell or a shared
    /// string's `<v>` runs [`near_limit`]: `max_json_length`, or less when
    /// `out` is near the end of its capacity. One comparison per check.
    slow_at: usize,
    xsd: XsdBuf,
    itoa: itoa::Buffer,
    first_row: bool,
    first_cell: bool,
    /// Where the open row starts in `out`, before its separating `,`.
    row_start: usize,
    /// `first_row` as it was before the open row began.
    row_was_first: bool,
    /// The attributes of the open `<row>`, written at its end.
    row: RowAttrs,
    /// The formula fields of the current cell not yet written.
    formula: FormulaState,
    /// Where the current cell's `cellType` text starts in `out`.
    cell_type_at: usize,
    /// The length of the current cell's `cellType` text at `cell_type_at`.
    cell_type_len: usize,
    /// Where the `,"richText":[...]` text of the current cell's last `<v>`
    /// lies in `out`, when that `<v>` named a rich entry.
    rich_text_at: Option<Range<usize>>,
}

impl<'a> JsonSink<'a> {
    /// A sink writing into `out`, which holds at most `max_json_length`
    /// bytes after each child of a cell, each cell and row, and each element
    /// of the sheet's lists, resolving shared strings in `sst` and `t="d"`
    /// cells in `date_system`. Writes `{"rows":[`.
    pub(super) fn new(
        out: &'a mut String,
        sst: Option<&'a SharedStringTable>,
        date_system: DateSystem,
        max_json_length: usize,
    ) -> Self {
        out.push_str("{\"rows\":[");
        let slow_at = slow_at(out, max_json_length);
        Self {
            out,
            sst,
            date_system,
            max_json_length,
            slow_at,
            xsd: XsdBuf::new(),
            itoa: itoa::Buffer::new(),
            first_row: true,
            first_cell: true,
            row_start: 0,
            row_was_first: true,
            row: RowAttrs::default(),
            formula: FormulaState::default(),
            cell_type_at: 0,
            cell_type_len: 0,
            rich_text_at: None,
        }
    }

    /// Close the rows array and write the rest of the sheet, each key only
    /// when it holds something, in the order below (`comments` and `tables`
    /// before `tabColor`); then the closing `}`. `dimension` and `tabColor`
    /// go through `json_escape_to`; `autoFilter` and `pageBreaks` are written
    /// field by field, the same text as serde_json gives; the others go
    /// through serde_json.
    ///
    /// Every list of the sheet (`mergeCells`, `paneSelections`, `columns`,
    /// `dataValidations`, `conditionalFormatting`, `hyperlinks`,
    /// `comments`, `tables`, `sparklineGroups`, `preservedExtensions`, the
    /// auto filter's columns with the values and custom filters of each, and
    /// the row and column page breaks) is written one element at a time,
    /// each followed by a check of `out` against `max_json_length`, so no
    /// list is serialized whole beside `out` and an element holding a copy of
    /// text stored once is caught at once. One element of a list (a table
    /// with its columns, a conditional format with its rules) is serialized
    /// whole. Each other key holds one value of fixed fields (numbers, flags
    /// and strings), and the caller checks `out` again when the sheet is
    /// done.
    ///
    /// # Errors
    ///
    /// `Json` when serde_json fails to serialize a value; `SizeLimit` when
    /// `out` grows past `max_json_length` after a list element.
    pub(super) fn finish(
        self,
        meta: &SheetMeta,
        comments: &[Comment],
        tables: &[TableDefinition],
    ) -> Result<()> {
        let max = self.max_json_length;
        let out = self.out;
        out.push(']');
        if let Some(ref d) = meta.dimension {
            out.push_str(",\"dimension\":\"");
            json_escape_to(out, d);
            out.push('"');
        }
        if !meta.merge_cells.is_empty() {
            out.push_str(",\"mergeCells\":");
            push_json_array_checked(out, &meta.merge_cells, max)?;
        }
        if let Some(ref af) = meta.auto_filter {
            out.push_str(",\"autoFilter\":");
            push_auto_filter_checked(out, af, max)?;
        }
        if let Some(ref fp) = meta.frozen_pane {
            out.push_str(",\"frozenPane\":");
            out.push_str(&serde_json::to_string(fp)?);
        }
        if let Some(ref sp) = meta.split_pane {
            out.push_str(",\"splitPane\":");
            out.push_str(&serde_json::to_string(sp)?);
        }
        if !meta.pane_selections.is_empty() {
            out.push_str(",\"paneSelections\":");
            push_json_array_checked(out, &meta.pane_selections, max)?;
        }
        if let Some(ref sv) = meta.sheet_view {
            out.push_str(",\"sheetView\":");
            out.push_str(&serde_json::to_string(sv)?);
        }
        if !meta.columns.is_empty() {
            out.push_str(",\"columns\":");
            push_json_array_checked(out, &meta.columns, max)?;
        }
        if !meta.data_validations.is_empty() {
            out.push_str(",\"dataValidations\":");
            push_json_array_checked(out, &meta.data_validations, max)?;
        }
        if !meta.conditional_formatting.is_empty() {
            out.push_str(",\"conditionalFormatting\":");
            push_json_array_checked(out, &meta.conditional_formatting, max)?;
        }
        if !meta.hyperlinks.is_empty() {
            out.push_str(",\"hyperlinks\":");
            push_json_array_checked(out, &meta.hyperlinks, max)?;
        }
        if let Some(ref ps) = meta.page_setup {
            out.push_str(",\"pageSetup\":");
            out.push_str(&serde_json::to_string(ps)?);
        }
        if let Some(ref sp) = meta.sheet_protection {
            out.push_str(",\"sheetProtection\":");
            out.push_str(&serde_json::to_string(sp)?);
        }
        if !comments.is_empty() {
            out.push_str(",\"comments\":");
            push_json_array_checked(out, comments, max)?;
        }
        if !tables.is_empty() {
            out.push_str(",\"tables\":");
            push_json_array_checked(out, tables, max)?;
        }
        if let Some(ref tc) = meta.tab_color {
            out.push_str(",\"tabColor\":\"");
            json_escape_to(out, tc);
            out.push('"');
        }
        if let Some(ref spec) = meta.tab_color_spec {
            out.push_str(",\"tabColorSpec\":");
            out.push_str(&serde_json::to_string(spec)?);
        }
        if let Some(ref hf) = meta.header_footer {
            out.push_str(",\"headerFooter\":");
            out.push_str(&serde_json::to_string(hf)?);
        }
        if let Some(ref pb) = meta.page_breaks {
            out.push_str(",\"pageBreaks\":");
            push_page_breaks_checked(out, pb, max)?;
        }
        if let Some(ref op) = meta.outline_properties {
            out.push_str(",\"outlineProperties\":");
            out.push_str(&serde_json::to_string(op)?);
        }
        if !meta.sparkline_groups.is_empty() {
            out.push_str(",\"sparklineGroups\":");
            push_json_array_checked(out, &meta.sparkline_groups, max)?;
        }
        if !meta.preserved_extensions.is_empty() {
            out.push_str(",\"preservedExtensions\":");
            push_json_array_checked(out, &meta.preserved_extensions, max)?;
        }
        out.push('}');
        Ok(())
    }
}

impl RowSink for JsonSink<'_> {
    fn begin_row(&mut self, attrs: &RowAttrs) {
        self.row = *attrs;
        self.row_start = self.out.len();
        self.row_was_first = self.first_row;
        let out = &mut *self.out;
        if !self.first_row {
            out.push(',');
        }
        self.first_row = false;
        out.push_str("{\"index\":");
        out.push_str(self.itoa.format(attrs.index));
        out.push_str(",\"cells\":[");
        self.first_cell = true;
    }

    /// Write the cell's attributes: `reference`, `cellType` and `styleIndex`.
    fn begin_cell(&mut self, attrs: &CellAttrs) {
        let out = &mut *self.out;
        if !self.first_cell {
            out.push(',');
        }
        self.first_cell = false;
        out.push_str("{\"reference\":\"");
        json_escape_to(out, &attrs.reference);
        out.push_str("\",\"cellType\":\"");
        // A `t="d"` cell's `cellType` is written as "number" here and
        // rewritten in place by each of its `<v>`.
        let cell_type = cell_type_json_str(attrs.cell_type);
        self.cell_type_at = out.len();
        self.cell_type_len = cell_type.len();
        out.push_str(cell_type);
        out.push('"');
        if let Some(si) = attrs.style_index {
            out.push_str(",\"styleIndex\":");
            out.push_str(self.itoa.format(si));
        }
        self.formula = FormulaState::default();
        self.formula.set_cell_metadata(attrs.cell_metadata);
        self.rich_text_at = None;
    }

    /// Write the keys of one child, then hold the JSON to the limit.
    fn cell_part(
        &mut self,
        attrs: &CellAttrs,
        part: CellPart<'_>,
        budget: &mut ExpansionBudget,
    ) -> Result<()> {
        let out = &mut *self.out;
        match part {
            CellPart::Value(v) => {
                // This `<v>` replaces the value of an earlier one and with it
                // that one's runs: the model path takes `richText` from the
                // last `<v>` alone.
                if let Some(earlier) = self.rich_text_at.take() {
                    out.replace_range(earlier, "");
                }
                out.push_str(",\"value\":\"");
                // A SharedString cell writes its SST entry's text, then the
                // runs of a rich entry, both resolved by index and
                // counted in `budget` before they are written.
                if attrs.cell_type == CellType::SharedString
                    && let Some((entry, runs)) = resolve_sst(self.sst, v, budget)?
                {
                    json_escape_to(out, entry);
                    out.push('"');
                    if let Some(runs) = runs {
                        let start = out.len();
                        out.push_str(",\"richText\":");
                        out.push_str(&serde_json::to_string(runs)?);
                        self.rich_text_at = Some(start..out.len());
                    }
                } else if attrs.iso_date {
                    // Each `<v>` sets the type, as on the model path: a date
                    // gives "number", other text "formulaStr", whatever an
                    // earlier `<v>` wrote.
                    let (cell_type, value) =
                        iso_date_cell(v.to_owned(), self.date_system, &mut self.xsd)?;
                    let cell_type = cell_type_json_str(cell_type);
                    out.replace_range(
                        self.cell_type_at..self.cell_type_at + self.cell_type_len,
                        cell_type,
                    );
                    self.cell_type_len = cell_type.len();
                    json_escape_to(out, &value);
                    out.push('"');
                } else {
                    json_escape_to(out, v);
                    out.push('"');
                }
            }
            CellPart::Formula { attrs: fa, text } => {
                self.formula.merge_from(fa);
                if let Some(text) = text {
                    out.push_str(",\"formula\":\"");
                    json_escape_to(out, text);
                    out.push('"');
                    write_formula_json_fields(out, &mut self.itoa, &mut self.formula);
                }
            }
            CellPart::InlineStr(t) => {
                out.push_str(",\"value\":\"");
                json_escape_to(out, t);
                out.push_str("\",\"inlineString\":\"");
                json_escape_to(out, t);
                out.push('"');
            }
        }
        // Every child writes its keys again (a repeated `<v>` its entry
        // again), so a cell of many children is held to the limit after each.
        if out.len() > self.slow_at {
            near_limit(out, self.max_json_length, &mut self.slow_at)?;
        }
        Ok(())
    }

    /// Write the formula fields not yet written and close the cell.
    fn end_cell(&mut self, _attrs: &CellAttrs, _budget: &mut ExpansionBudget) -> Result<()> {
        write_formula_json_fields(self.out, &mut self.itoa, &mut self.formula);
        self.out.push('}');
        // The cell's copies were counted at their byte length but written
        // escaped (up to 6x), so the JSON is held to the limit here, not only
        // at the row's end.
        if self.out.len() > self.slow_at {
            near_limit(self.out, self.max_json_length, &mut self.slow_at)?;
        }
        Ok(())
    }

    fn end_row(&mut self) -> Result<()> {
        let row = &self.row;
        let out = &mut *self.out;
        out.push(']');
        if let Some(h) = row.height {
            out.push_str(",\"height\":");
            write_f64_json(out, h);
        }
        if row.custom_height {
            out.push_str(",\"customHeight\":true");
        }
        if row.hidden {
            out.push_str(",\"hidden\":true");
        }
        if let Some(level) = row.outline_level {
            out.push_str(",\"outlineLevel\":");
            out.push_str(self.itoa.format(level));
        }
        if row.collapsed {
            out.push_str(",\"collapsed\":true");
        }
        if let Some(s) = row.style_index {
            out.push_str(",\"styleIndex\":");
            out.push_str(self.itoa.format(s));
        }
        out.push('}');
        check_json_len(out, self.max_json_length)
    }

    /// Cut the open row, with its separating `,`, back out of `out`.
    fn abort_row(&mut self) {
        self.out.truncate(self.row_start);
        self.first_row = self.row_was_first;
        self.formula = FormulaState::default();
    }
}

/// The length of `out` past which [`near_limit`] runs: `max`, or seven
/// eighths of `out`'s capacity when that is less, so the check after a cell
/// or a shared string's `<v>` can keep `out` from doubling past the limit
/// before it grows.
fn slow_at(out: &String, max: usize) -> usize {
    let cap = out.capacity();
    max.min(cap - cap / 8)
}

/// The check after a cell or a shared string's `<v>` once `out` passes
/// `slow_at`: fails past `max`, otherwise keeps `out` from doubling past it
/// and moves `slow_at`.
///
/// # Errors
///
/// `SizeLimit` when `out` is longer than `max`.
fn near_limit(out: &mut String, max: usize, slow_at_len: &mut usize) -> Result<()> {
    check_json_len(out, max)?;
    clamp_json_growth(out, max);
    *slow_at_len = slow_at(out, max);
    Ok(())
}

/// Write `items` into `out` as a JSON array, the same text as
/// `serde_json::to_string(items)`, checking `out` against `max` after each
/// element.
///
/// # Errors
///
/// `Json` when serde_json fails; `SizeLimit` when `out` grows past `max`.
fn push_json_array_checked<T: serde::Serialize>(
    out: &mut String,
    items: &[T],
    max: usize,
) -> Result<()> {
    out.push('[');
    for (i, item) in items.iter().enumerate() {
        if i > 0 {
            out.push(',');
        }
        out.push_str(&serde_json::to_string(item)?);
        check_json_len(out, max)?;
        clamp_json_growth(out, max);
    }
    out.push(']');
    Ok(())
}

/// Write `"key":` into the JSON object being written in `out`, after a `,`
/// unless it is the object's first key (`first`, cleared here).
fn push_key(out: &mut String, first: &mut bool, key: &str) {
    if !std::mem::take(first) {
        out.push(',');
    }
    out.push('"');
    out.push_str(key);
    out.push_str("\":");
}

/// Write `pb` into `out`, the same text as `serde_json::to_string(pb)`
/// (field order and the skip of an empty list included), with its row and
/// column breaks written by [`push_json_array_checked`].
///
/// # Errors
///
/// `Json` when serde_json fails; `SizeLimit` when `out` grows past `max`.
fn push_page_breaks_checked(out: &mut String, pb: &PageBreaks, max: usize) -> Result<()> {
    let PageBreaks {
        row_breaks,
        col_breaks,
    } = pb;
    let mut first = true;
    out.push('{');
    if !row_breaks.is_empty() {
        push_key(out, &mut first, "rowBreaks");
        push_json_array_checked(out, row_breaks, max)?;
    }
    if !col_breaks.is_empty() {
        push_key(out, &mut first, "colBreaks");
        push_json_array_checked(out, col_breaks, max)?;
    }
    out.push('}');
    Ok(())
}

/// Write `af` into `out`, the same text as `serde_json::to_string(af)`
/// (field order and skip rules included), checking `out` against `max`
/// after each filter column and, inside one, after each filter value and
/// each custom filter ([`push_json_array_checked`]). The range is one
/// string, written whole.
///
/// # Errors
///
/// `Json` when serde_json fails; `SizeLimit` when `out` grows past `max`.
fn push_auto_filter_checked(out: &mut String, af: &AutoFilter, max: usize) -> Result<()> {
    let AutoFilter {
        range,
        filter_columns,
    } = af;
    out.push_str("{\"range\":");
    out.push_str(&serde_json::to_string(range)?);
    if !filter_columns.is_empty() {
        out.push_str(",\"filterColumns\":[");
        for (i, column) in filter_columns.iter().enumerate() {
            if i > 0 {
                out.push(',');
            }
            push_filter_column_checked(out, column, max)?;
            check_json_len(out, max)?;
            clamp_json_growth(out, max);
        }
        out.push(']');
    }
    out.push('}');
    Ok(())
}

/// Write `column` into `out`, the same text as
/// `serde_json::to_string(column)`, with its filter values and its custom
/// filters written by [`push_json_array_checked`].
///
/// # Errors
///
/// `Json` when serde_json fails; `SizeLimit` when `out` grows past `max`.
fn push_filter_column_checked(out: &mut String, column: &FilterColumn, max: usize) -> Result<()> {
    let FilterColumn {
        col_id,
        filters,
        custom_filters,
    } = column;
    out.push_str("{\"colId\":");
    out.push_str(itoa::Buffer::new().format(*col_id));
    if !filters.is_empty() {
        out.push_str(",\"filters\":");
        push_json_array_checked(out, filters, max)?;
    }
    if let Some(custom) = custom_filters {
        let CustomFilters { and_op, filters } = custom;
        out.push_str(",\"customFilters\":{");
        if *and_op {
            out.push_str("\"andOp\":true,");
        }
        out.push_str("\"filters\":");
        push_json_array_checked(out, filters, max)?;
        out.push('}');
    }
    out.push('}');
    Ok(())
}

/// Write the formula fields held in `fs` (and the `<c cm>` index it carries)
/// as JSON into `out`, taking each one so that a later call writes only what
/// has been set since: [`JsonSink`] calls this after each `<f>` with text and
/// again at the cell's end (for a self-closing `<f/>`, or a `cm` cell without
/// a formula), and a field is written again only when a later `<f>` sets it.
#[inline]
fn write_formula_json_fields(out: &mut String, itoa_buf: &mut itoa::Buffer, fs: &mut FormulaState) {
    if let Some(cm) = fs.cell_metadata.take() {
        out.push_str(",\"cellMetadata\":");
        out.push_str(itoa_buf.format(cm));
    }
    if let Some(ft) = fs.formula_type.take() {
        out.push_str(",\"formulaType\":\"");
        json_escape_to(out, &ft);
        out.push('"');
    }
    if let Some(fr) = fs.formula_ref.take() {
        out.push_str(",\"formulaRef\":\"");
        json_escape_to(out, &fr);
        out.push('"');
    }
    if let Some(si) = fs.shared_index.take() {
        out.push_str(",\"sharedIndex\":");
        out.push_str(itoa_buf.format(si));
    }
    if fs.dynamic_array.take() == Some(true) {
        out.push_str(",\"dynamicArray\":true");
    }
    if let Some(r1) = fs.formula_r1.take() {
        out.push_str(",\"formulaR1\":\"");
        json_escape_to(out, &r1);
        out.push('"');
    }
    if let Some(r2) = fs.formula_r2.take() {
        out.push_str(",\"formulaR2\":\"");
        json_escape_to(out, &r2);
        out.push('"');
    }
    if fs.formula_dt2d.take() == Some(true) {
        out.push_str(",\"formulaDt2d\":true");
    }
    if fs.formula_dtr.take() == Some(true) {
        out.push_str(",\"formulaDtr\":true");
    }
    if fs.formula_del1.take() == Some(true) {
        out.push_str(",\"formulaDel1\":true");
    }
    if fs.formula_del2.take() == Some(true) {
        out.push_str(",\"formulaDel2\":true");
    }
    if fs.formula_dtr1.take() == Some(true) {
        out.push_str(",\"formulaDtr1\":true");
    }
    if fs.formula_dtr2.take() == Some(true) {
        out.push_str(",\"formulaDtr2\":true");
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::ModernXlsxError;
    use crate::ooxml::tables::{TableColumn, TableStyleInfo};
    use crate::ooxml::worksheet::{CustomFilter, PageBreak, SheetContext, WorksheetXml};

    /// Strings with every kind of JSON escape, `\b` and `\f` among them
    /// (which serde_json spells by their short forms), and none.
    const TEXTS: [&str; 3] = ["A1:D10", "a\"b\\c\u{8}\u{c}\n\t\u{1}\u{1f}é😀", ""];

    fn breaks() -> Vec<Vec<PageBreak>> {
        let b = |id, min, max, man| PageBreak { id, min, max, man };
        vec![
            Vec::new(),
            vec![b(1, None, None, false)],
            vec![
                b(2, Some(0), None, false),
                b(3, None, Some(16_383), false),
                b(4, Some(1), Some(1_048_575), true),
                b(5, None, None, true),
            ],
        ]
    }

    fn filter_columns() -> Vec<FilterColumn> {
        let mut customs = vec![None];
        for and_op in [false, true] {
            for filters in [
                Vec::new(),
                vec![CustomFilter {
                    operator: None,
                    val: TEXTS[1].into(),
                }],
                vec![
                    CustomFilter {
                        operator: Some("greaterThan".into()),
                        val: "1".into(),
                    },
                    CustomFilter {
                        operator: Some(TEXTS[1].into()),
                        val: String::new(),
                    },
                ],
            ] {
                customs.push(Some(CustomFilters { and_op, filters }));
            }
        }
        let mut columns = Vec::new();
        for (i, custom_filters) in customs.into_iter().enumerate() {
            for filters in [Vec::new(), TEXTS.map(String::from).to_vec()] {
                columns.push(FilterColumn {
                    col_id: u32::try_from(i).unwrap() * 1_000_000,
                    filters,
                    custom_filters: custom_filters.clone(),
                });
            }
        }
        columns
    }

    fn checked(write: impl FnOnce(&mut String, usize) -> Result<()>) -> String {
        let mut out = String::new();
        write(&mut out, usize::MAX).unwrap();
        out
    }

    #[test]
    fn page_breaks_match_serde_json() {
        for row_breaks in breaks() {
            for col_breaks in breaks() {
                let pb = PageBreaks {
                    row_breaks: row_breaks.clone(),
                    col_breaks,
                };
                let ours = checked(|out, max| push_page_breaks_checked(out, &pb, max));
                assert_eq!(ours, serde_json::to_string(&pb).unwrap());
            }
        }
    }

    #[test]
    fn auto_filter_matches_serde_json() {
        let columns = filter_columns();
        let mut column_lists = vec![Vec::new(), columns.clone()];
        column_lists.extend(columns.into_iter().map(|c| vec![c]));
        for range in TEXTS {
            for filter_columns in &column_lists {
                let af = AutoFilter {
                    range: range.into(),
                    filter_columns: filter_columns.clone(),
                };
                let ours = checked(|out, max| push_auto_filter_checked(out, &af, max));
                assert_eq!(ours, serde_json::to_string(&af).unwrap());
            }
        }
    }

    #[test]
    fn tables_match_serde_json() {
        let bare = TableDefinition {
            id: 1,
            name: None,
            display_name: "T".into(),
            ref_range: "A1:B2".into(),
            header_row_count: 0,
            totals_row_count: 0,
            totals_row_shown: false,
            columns: Vec::new(),
            style_info: None,
            auto_filter_ref: None,
        };
        let full = TableDefinition {
            id: 2,
            name: Some(TEXTS[1].into()),
            display_name: TEXTS[1].into(),
            ref_range: "A1:C9".into(),
            header_row_count: 1,
            totals_row_count: 1,
            totals_row_shown: true,
            columns: vec![
                TableColumn {
                    id: 1,
                    name: TEXTS[1].into(),
                    ..Default::default()
                },
                TableColumn {
                    id: 2,
                    name: "B".into(),
                    totals_row_function: Some("custom".into()),
                    totals_row_label: Some("L".into()),
                    calculated_column_formula: Some("[@A]*2".into()),
                    totals_row_formula: Some("SUM(T[B])".into()),
                    header_row_dxf_id: Some(0),
                    data_dxf_id: Some(1),
                    totals_row_dxf_id: Some(2),
                },
            ],
            style_info: Some(TableStyleInfo {
                name: Some("TableStyleMedium2".into()),
                show_first_column: true,
                show_last_column: true,
                show_row_stripes: false,
                show_column_stripes: true,
            }),
            auto_filter_ref: Some("A1:C8".into()),
        };
        for tables in [vec![bare.clone()], vec![full.clone()], vec![bare, full]] {
            let ours = checked(|out, max| push_json_array_checked(out, &tables, max));
            assert_eq!(ours, serde_json::to_string(&tables).unwrap());
        }
    }

    /// Each list of the page breaks and the auto filter fails as soon as
    /// `out` passes the limit, one element past it at most.
    #[test]
    fn page_breaks_and_auto_filter_fail_at_the_limit() {
        const MAX: usize = 1_000;
        let many_breaks = vec![
            PageBreak {
                id: 1,
                min: None,
                max: None,
                man: true,
            };
            10_000
        ];
        let column = |filters: Vec<String>, custom: Vec<CustomFilter>| FilterColumn {
            col_id: 0,
            filters,
            custom_filters: (!custom.is_empty()).then_some(CustomFilters {
                and_op: false,
                filters: custom,
            }),
        };
        let custom = CustomFilter {
            operator: None,
            val: "v".into(),
        };
        let row_breaks = PageBreaks {
            row_breaks: many_breaks.clone(),
            col_breaks: Vec::new(),
        };
        let col_breaks = PageBreaks {
            row_breaks: Vec::new(),
            col_breaks: many_breaks,
        };
        let af = |filter_columns| AutoFilter {
            range: "A1".into(),
            filter_columns,
        };
        let columns = af(vec![column(Vec::new(), Vec::new()); 10_000]);
        let values = af(vec![column(vec!["v".into(); 10_000], Vec::new())]);
        let customs = af(vec![column(Vec::new(), vec![custom; 10_000])]);
        type Write<'a> = &'a dyn Fn(&mut String) -> Result<()>;
        let cases: [(&str, Write<'_>); 5] = [
            ("row breaks", &|out| {
                push_page_breaks_checked(out, &row_breaks, MAX)
            }),
            ("column breaks", &|out| {
                push_page_breaks_checked(out, &col_breaks, MAX)
            }),
            ("filter columns", &|out| {
                push_auto_filter_checked(out, &columns, MAX)
            }),
            ("filter values", &|out| {
                push_auto_filter_checked(out, &values, MAX)
            }),
            ("custom filters", &|out| {
                push_auto_filter_checked(out, &customs, MAX)
            }),
        ];
        for (label, write) in cases {
            let mut out = String::new();
            let err = write(&mut out).unwrap_err();
            assert!(
                matches!(err, ModernXlsxError::SizeLimit(_)),
                "{label}: {err}"
            );
            assert!(out.len() <= MAX + 64, "{label}: out grew to {}", out.len());
        }
    }

    /// A sheet whose row breaks, column breaks or filter values are counted
    /// in the budget under its limit but whose JSON passes it fails in
    /// `WorksheetXml::parse_to_json_within`, one element past the limit at
    /// most: only the checks of [`JsonSink::finish`] catch these lists, as
    /// the read checks `out` again only once the sheet is written.
    #[test]
    fn sheet_lists_fail_at_the_limit() {
        const MAX: usize = 100_000;
        // 3,000 breaks are counted at `size_of::<PageBreak>()` each (72,000
        // bytes on a 64-bit target) and written as 38 bytes each with the
        // separating comma.
        let breaks = "<brk id=\"1048575\" max=\"16383\" man=\"1\"/>".repeat(3_000);
        // 1,800 values of 28 backslashes are counted at
        // `size_of::<String>()` plus 28 bytes each (93,600 on a 64-bit
        // target) and written as 59 bytes each with the separating comma,
        // every backslash escaped.
        let values = format!("<filter val=\"{}\"/>", "\\".repeat(28)).repeat(1_800);
        let cases = [
            ("row breaks", format!("<rowBreaks>{breaks}</rowBreaks>")),
            ("column breaks", format!("<colBreaks>{breaks}</colBreaks>")),
            (
                "filter values",
                format!(
                    "<autoFilter ref=\"A1:A2\"><filterColumn colId=\"0\"><filters>{values}</filters></filterColumn></autoFilter>"
                ),
            ),
        ];
        for (label, body) in cases {
            let xml = format!(
                "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><sheetData/>{body}</worksheet>"
            );
            let mut out = String::new();
            let mut budget = ExpansionBudget::new(MAX);
            let err = WorksheetXml::parse_to_json_within(
                xml.as_bytes(),
                None,
                &[],
                &[],
                &SheetContext::default(),
                &mut out,
                &mut budget,
            )
            .unwrap_err();
            assert!(
                matches!(err, ModernXlsxError::SizeLimit(ref m) if m.starts_with("workbook JSON reached")),
                "{label}: {err}"
            );
            assert!(
                budget.stored <= MAX,
                "{label}: counted {} bytes",
                budget.stored
            );
            assert!(out.len() <= MAX + 64, "{label}: out grew to {}", out.len());
        }
    }
}
