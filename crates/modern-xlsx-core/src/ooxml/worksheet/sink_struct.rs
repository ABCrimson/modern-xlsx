//! The [`RowSink`] behind `WorksheetXml::parse_with_sst`: [`StructSink`]
//! builds the typed [`Row`]s and [`Cell`]s from what the driver reads.

use super::driver::{CellPart, RowSink};
use super::parser::{CellAttrs, FormulaState, RowAttrs, iso_date_cell, resolve_sst};
use super::{Cell, CellType, Row};
use crate::Result;
use crate::dates::DateSystem;
use crate::ooxml::shared_strings::{ExpansionBudget, SharedStringTable};
use crate::ooxml::xml::XsdBuf;

/// Builds the rows of a worksheet. A SharedString cell gets the SST entry its
/// last `<v>` names: the text as `value` and, for a rich entry, the runs as
/// `rich_text` (no index string is kept; an earlier `<v>` is neither resolved
/// nor counted in the read's `ExpansionBudget`); a `t="d"` cell becomes a number cell holding
/// its serial in `date_system`, or a `str` cell when its text is not a date.
/// A cell is built as its children arrive, so it holds one value, formula
/// and inline string at a time, however many children it has.
pub(super) struct StructSink<'s> {
    sst: Option<&'s SharedStringTable>,
    date_system: DateSystem,
    xsd: XsdBuf,
    rows: Vec<Row>,
    /// The attributes of the open `<row>`.
    row: RowAttrs,
    /// The cells of the open `<row>`.
    cells: Vec<Cell>,
    /// The open cell.
    open: OpenCell,
}

/// What the children of the open cell set so far.
#[derive(Default)]
struct OpenCell {
    cell_type: CellType,
    value: Option<String>,
    formula: Option<String>,
    inline_string: Option<String>,
    /// The formula fields merged from every `<f>`, and the `cm` index.
    f: FormulaState,
    /// A SharedString cell's last `<v>` text, resolved at the cell's end.
    sst_index: String,
    /// Whether the cell had such a `<v>`.
    has_sst_index: bool,
    /// Whether that `<v>` gives the value (no `<is>` came after it).
    value_is_sst: bool,
}

impl<'s> StructSink<'s> {
    /// A sink resolving shared strings in `sst` and `t="d"` cells in
    /// `date_system`, with room for `row_capacity` rows.
    pub(super) fn new(
        sst: Option<&'s SharedStringTable>,
        date_system: DateSystem,
        row_capacity: usize,
    ) -> Self {
        Self {
            sst,
            date_system,
            xsd: XsdBuf::new(),
            rows: Vec::with_capacity(row_capacity),
            row: RowAttrs::default(),
            cells: Vec::with_capacity(16),
            open: OpenCell::default(),
        }
    }

    /// The finished rows. A `<row>` that never ended is not among them.
    pub(super) fn into_rows(self) -> Vec<Row> {
        self.rows
    }
}

impl RowSink for StructSink<'_> {
    fn begin_row(&mut self, attrs: &RowAttrs) {
        self.row = *attrs;
        self.cells.clear();
    }

    fn begin_cell(&mut self, attrs: &CellAttrs) {
        let open = &mut self.open;
        open.cell_type = attrs.cell_type;
        open.value = None;
        open.formula = None;
        open.inline_string = None;
        open.f = FormulaState::default();
        open.f.set_cell_metadata(attrs.cell_metadata);
        open.sst_index.clear();
        open.has_sst_index = false;
        open.value_is_sst = false;
    }

    fn cell_part(
        &mut self,
        attrs: &CellAttrs,
        part: CellPart<'_>,
        _budget: &mut ExpansionBudget,
    ) -> Result<()> {
        let open = &mut self.open;
        // Parts apply in document order: a `<v>` or an `<is>` replaces the
        // value an earlier one set, while an `<f>` merges its attributes into
        // the earlier ones and replaces the formula text only when it has
        // text (a self-closing `<f/>` keeps it).
        match part {
            CellPart::Value(v) => {
                if open.cell_type == CellType::SharedString {
                    // Only the last `<v>` of a SharedString cell gives its
                    // value, so its index is resolved at the cell's end: an
                    // earlier one would be counted in `budget` and dropped.
                    open.sst_index.clear();
                    open.sst_index.push_str(v);
                    open.has_sst_index = true;
                    open.value_is_sst = true;
                } else if attrs.iso_date {
                    let (iso_type, iso_value) =
                        iso_date_cell(v.to_owned(), self.date_system, &mut self.xsd)?;
                    open.cell_type = iso_type;
                    open.value = Some(iso_value);
                } else if let Some(value) = &mut open.value {
                    // A repeated `<v>` reuses the value's buffer.
                    value.clear();
                    value.push_str(v);
                } else {
                    open.value = Some(v.to_owned());
                }
            }
            CellPart::Formula { attrs: fa, text } => {
                open.f.merge_from(fa);
                if let Some(text) = text {
                    open.formula = Some(text.to_owned());
                }
            }
            CellPart::InlineStr(t) => {
                open.inline_string = Some(t.to_owned());
                open.value = Some(t.to_owned());
                open.value_is_sst = false;
            }
        }
        Ok(())
    }

    fn end_cell(&mut self, attrs: &CellAttrs, budget: &mut ExpansionBudget) -> Result<()> {
        let open = &mut self.open;
        let mut rich_text = None;
        if open.has_sst_index {
            // The last `<v>` resolves to the SST entry its index names: the
            // text as `value` (unless an `<is>` came after it) and the runs
            // of a rich entry; `resolve_sst` counts both in `budget` before
            // they are cloned. Without an SST or such an entry, the index
            // text is the value.
            match resolve_sst(self.sst, &open.sst_index, budget)? {
                Some((entry, runs)) => {
                    if open.value_is_sst {
                        open.value = Some(entry.clone());
                    }
                    rich_text = runs.cloned();
                }
                None => {
                    if open.value_is_sst {
                        open.value = Some(open.sst_index.clone());
                    }
                }
            }
        }
        let f = std::mem::take(&mut open.f);
        self.cells.push(Cell {
            reference: attrs.reference.clone(),
            cell_type: open.cell_type,
            style_index: attrs.style_index,
            value: open.value.take(),
            formula: open.formula.take(),
            formula_type: f.formula_type,
            formula_ref: f.formula_ref,
            shared_index: f.shared_index,
            inline_string: open.inline_string.take(),
            cell_metadata: f.cell_metadata,
            dynamic_array: f.dynamic_array,
            formula_r1: f.formula_r1,
            formula_r2: f.formula_r2,
            formula_dt2d: f.formula_dt2d,
            formula_dtr: f.formula_dtr,
            formula_del1: f.formula_del1,
            formula_del2: f.formula_del2,
            formula_dtr1: f.formula_dtr1,
            formula_dtr2: f.formula_dtr2,
            rich_text,
        });
        Ok(())
    }

    fn end_row(&mut self) -> Result<()> {
        let row = &self.row;
        self.rows.push(Row {
            index: row.index,
            cells: std::mem::take(&mut self.cells),
            height: row.height,
            custom_height: row.custom_height,
            hidden: row.hidden,
            outline_level: row.outline_level,
            collapsed: row.collapsed,
            style_index: row.style_index,
        });
        Ok(())
    }

    /// Drop the cells of the open row; it never reaches `rows`.
    fn abort_row(&mut self) {
        self.cells.clear();
    }
}
