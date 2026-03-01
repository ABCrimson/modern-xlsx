use quick_xml::events::{BytesDecl, BytesEnd, BytesStart, BytesText, Event};
use quick_xml::{Reader, Writer};
use serde::{Deserialize, Serialize};

use crate::{IronsheetError, Result};

const SPREADSHEET_NS: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

/// Parsed representation of a worksheet XML file (`xl/worksheets/sheet*.xml`).
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct WorksheetXml {
    pub dimension: Option<String>,
    pub rows: Vec<Row>,
    pub merge_cells: Vec<String>,
    pub auto_filter: Option<String>,
    pub frozen_pane: Option<FrozenPane>,
    pub columns: Vec<ColumnInfo>,
}

/// A single row in the worksheet.
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct Row {
    /// 1-based row index (as in the XML `r` attribute).
    pub index: u32,
    pub cells: Vec<Cell>,
    pub height: Option<f64>,
    pub hidden: bool,
}

/// A single cell in a row.
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct Cell {
    /// Cell reference string, e.g. "A1".
    pub reference: String,
    pub cell_type: CellType,
    pub style_index: Option<u32>,
    /// Raw `<v>` content (the value element text).
    pub value: Option<String>,
    /// Raw `<f>` content (the formula element text).
    pub formula: Option<String>,
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
}

/// Frozen pane configuration.
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct FrozenPane {
    /// Number of frozen rows (ySplit).
    pub rows: u32,
    /// Number of frozen columns (xSplit).
    pub cols: u32,
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
}

// ---------------------------------------------------------------------------
// Parser state machine
// ---------------------------------------------------------------------------

/// Internal parsing state.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum ParseState {
    Root,
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
}

impl WorksheetXml {
    /// Parse a worksheet XML file from raw bytes.
    pub fn parse(data: &[u8]) -> Result<Self> {
        let mut reader = Reader::from_reader(data);
        reader.config_mut().trim_text(true);

        let mut buf = Vec::new();

        let mut dimension: Option<String> = None;
        let mut rows: Vec<Row> = Vec::new();
        let mut merge_cells: Vec<String> = Vec::new();
        let mut auto_filter: Option<String> = None;
        let mut frozen_pane: Option<FrozenPane> = None;
        let mut columns: Vec<ColumnInfo> = Vec::new();

        let mut state = ParseState::Root;

        // Current row being built.
        let mut cur_row_index: u32 = 0;
        let mut cur_row_height: Option<f64> = None;
        let mut cur_row_hidden = false;
        let mut cur_row_cells: Vec<Cell> = Vec::new();

        // Current cell being built.
        let mut cur_cell_ref = String::new();
        let mut cur_cell_type = CellType::Number;
        let mut cur_cell_style: Option<u32> = None;
        let mut cur_cell_value: Option<String> = None;
        let mut cur_cell_formula: Option<String> = None;

        // Buffers for text content.
        let mut text_buf = String::new();

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) => {
                    let local = e.local_name();
                    match (state, local.as_ref()) {
                        (ParseState::Root, b"sheetViews") => {
                            state = ParseState::SheetViews;
                        }
                        (ParseState::SheetViews, b"sheetView") => {
                            state = ParseState::SheetView;
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
                            cur_row_cells.clear();

                            for attr in e.attributes().flatten() {
                                let ln = attr.key.local_name();
                                match ln.as_ref() {
                                    b"r" => {
                                        let val = String::from_utf8_lossy(&attr.value);
                                        cur_row_index = val.parse::<u32>().unwrap_or(0);
                                    }
                                    b"ht" => {
                                        let val = String::from_utf8_lossy(&attr.value);
                                        cur_row_height = val.parse::<f64>().ok();
                                    }
                                    b"hidden" => {
                                        let val = String::from_utf8_lossy(&attr.value);
                                        cur_row_hidden = val == "1" || val.eq_ignore_ascii_case("true");
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

                            for attr in e.attributes().flatten() {
                                let ln = attr.key.local_name();
                                match ln.as_ref() {
                                    b"r" => {
                                        cur_cell_ref =
                                            String::from_utf8_lossy(&attr.value).into_owned();
                                    }
                                    b"t" => {
                                        let val = String::from_utf8_lossy(&attr.value);
                                        cur_cell_type = match val.as_ref() {
                                            "s" => CellType::SharedString,
                                            "b" => CellType::Boolean,
                                            "e" => CellType::Error,
                                            "str" => CellType::FormulaStr,
                                            "inlineStr" => CellType::InlineStr,
                                            _ => CellType::Number,
                                        };
                                    }
                                    b"s" => {
                                        let val = String::from_utf8_lossy(&attr.value);
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
                                        String::from_utf8_lossy(&attr.value).into_owned(),
                                    );
                                }
                            }
                        }
                        (ParseState::Root, b"autoFilter") => {
                            for attr in e.attributes().flatten() {
                                if attr.key.local_name().as_ref() == b"ref" {
                                    auto_filter = Some(
                                        String::from_utf8_lossy(&attr.value).into_owned(),
                                    );
                                }
                            }
                        }
                        (ParseState::SheetView, b"pane") => {
                            let mut y_split: u32 = 0;
                            let mut x_split: u32 = 0;
                            let mut is_frozen = false;

                            for attr in e.attributes().flatten() {
                                let ln = attr.key.local_name();
                                match ln.as_ref() {
                                    b"ySplit" => {
                                        let val = String::from_utf8_lossy(&attr.value);
                                        y_split = val.parse::<u32>().unwrap_or(0);
                                    }
                                    b"xSplit" => {
                                        let val = String::from_utf8_lossy(&attr.value);
                                        x_split = val.parse::<u32>().unwrap_or(0);
                                    }
                                    b"state" => {
                                        let val = String::from_utf8_lossy(&attr.value);
                                        is_frozen = val == "frozen";
                                    }
                                    _ => {}
                                }
                            }

                            if is_frozen && (y_split > 0 || x_split > 0) {
                                frozen_pane = Some(FrozenPane {
                                    rows: y_split,
                                    cols: x_split,
                                });
                            }
                        }
                        (ParseState::Cols, b"col") => {
                            columns.push(parse_col_element(e));
                        }
                        (ParseState::MergeCells, b"mergeCell") => {
                            for attr in e.attributes().flatten() {
                                if attr.key.local_name().as_ref() == b"ref" {
                                    merge_cells.push(
                                        String::from_utf8_lossy(&attr.value).into_owned(),
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
                                            String::from_utf8_lossy(&attr.value).into_owned();
                                    }
                                    b"t" => {
                                        let val = String::from_utf8_lossy(&attr.value);
                                        cell_type = match val.as_ref() {
                                            "s" => CellType::SharedString,
                                            "b" => CellType::Boolean,
                                            "e" => CellType::Error,
                                            "str" => CellType::FormulaStr,
                                            "inlineStr" => CellType::InlineStr,
                                            _ => CellType::Number,
                                        };
                                    }
                                    b"s" => {
                                        let val = String::from_utf8_lossy(&attr.value);
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
                            });
                        }
                        (ParseState::SheetData, b"row") => {
                            // Self-closing <row ... /> — empty row.
                            let mut row_index: u32 = 0;
                            let mut row_height: Option<f64> = None;
                            let mut row_hidden = false;

                            for attr in e.attributes().flatten() {
                                let ln = attr.key.local_name();
                                match ln.as_ref() {
                                    b"r" => {
                                        let val = String::from_utf8_lossy(&attr.value);
                                        row_index = val.parse::<u32>().unwrap_or(0);
                                    }
                                    b"ht" => {
                                        let val = String::from_utf8_lossy(&attr.value);
                                        row_height = val.parse::<f64>().ok();
                                    }
                                    b"hidden" => {
                                        let val = String::from_utf8_lossy(&attr.value);
                                        row_hidden = val == "1" || val.eq_ignore_ascii_case("true");
                                    }
                                    _ => {}
                                }
                            }
                            rows.push(Row {
                                index: row_index,
                                cells: Vec::new(),
                                height: row_height,
                                hidden: row_hidden,
                            });
                        }
                        (ParseState::InCell, b"v") => {
                            // Empty <v/> — no value.
                        }
                        (ParseState::InCell, b"f") => {
                            // Empty <f/> — no formula content.
                        }
                        _ => {}
                    }
                }
                Ok(Event::Text(ref e)) => {
                    match state {
                        ParseState::InCellValue
                        | ParseState::InCellFormula
                        | ParseState::InInlineStrT => {
                            text_buf.push_str(&String::from_utf8_lossy(e.as_ref()));
                        }
                        _ => {}
                    }
                }
                Ok(Event::End(ref e)) => {
                    let local = e.local_name();
                    match (state, local.as_ref()) {
                        (ParseState::InCellValue, b"v") => {
                            cur_cell_value = Some(std::mem::take(&mut text_buf));
                            state = ParseState::InCell;
                        }
                        (ParseState::InCellFormula, b"f") => {
                            cur_cell_formula = Some(std::mem::take(&mut text_buf));
                            state = ParseState::InCell;
                        }
                        (ParseState::InInlineStrT, b"t") => {
                            cur_cell_value = Some(std::mem::take(&mut text_buf));
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
                            });
                            state = ParseState::InRow;
                        }
                        (ParseState::InRow, b"row") => {
                            rows.push(Row {
                                index: cur_row_index,
                                cells: std::mem::take(&mut cur_row_cells),
                                height: cur_row_height.take(),
                                hidden: cur_row_hidden,
                            });
                            state = ParseState::SheetData;
                        }
                        (ParseState::SheetData, b"sheetData") => {
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
                        _ => {}
                    }
                }
                Ok(Event::Eof) => break,
                Err(err) => {
                    return Err(IronsheetError::XmlParse(format!(
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
            columns,
        })
    }

    /// Serialize this worksheet to a valid XML string.
    pub fn to_xml(&self) -> Result<String> {
        let mut buf: Vec<u8> = Vec::new();
        let mut writer = Writer::new(&mut buf);

        let map_err = |e: std::io::Error| IronsheetError::XmlWrite(e.to_string());

        // XML declaration.
        writer
            .write_event(Event::Decl(BytesDecl::new("1.0", Some("UTF-8"), Some("yes"))))
            .map_err(map_err)?;

        // <worksheet xmlns="...">
        let mut ws = BytesStart::new("worksheet");
        ws.push_attribute(("xmlns", SPREADSHEET_NS));
        writer.write_event(Event::Start(ws)).map_err(map_err)?;

        // <sheetViews> — only if frozen_pane is present.
        if let Some(ref pane) = self.frozen_pane {
            writer
                .write_event(Event::Start(BytesStart::new("sheetViews")))
                .map_err(map_err)?;

            let mut sv = BytesStart::new("sheetView");
            sv.push_attribute(("workbookViewId", "0"));
            writer.write_event(Event::Start(sv)).map_err(map_err)?;

            let mut pane_elem = BytesStart::new("pane");
            if pane.cols > 0 {
                let xs = pane.cols.to_string();
                pane_elem.push_attribute(("xSplit", xs.as_str()));
            }
            if pane.rows > 0 {
                let ys = pane.rows.to_string();
                pane_elem.push_attribute(("ySplit", ys.as_str()));
            }
            // Compute topLeftCell.
            let top_left = format!(
                "{}{}",
                col_index_to_letter(pane.cols + 1),
                pane.rows + 1
            );
            pane_elem.push_attribute(("topLeftCell", top_left.as_str()));
            pane_elem.push_attribute(("activePane", "bottomLeft"));
            pane_elem.push_attribute(("state", "frozen"));
            writer
                .write_event(Event::Empty(pane_elem))
                .map_err(map_err)?;

            writer
                .write_event(Event::End(BytesEnd::new("sheetView")))
                .map_err(map_err)?;
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
                let min_s = col.min.to_string();
                let max_s = col.max.to_string();
                let width_s = format_f64(col.width);
                elem.push_attribute(("min", min_s.as_str()));
                elem.push_attribute(("max", max_s.as_str()));
                elem.push_attribute(("width", width_s.as_str()));
                if col.custom_width {
                    elem.push_attribute(("customWidth", "1"));
                }
                if col.hidden {
                    elem.push_attribute(("hidden", "1"));
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
            let r_str = row.index.to_string();
            row_elem.push_attribute(("r", r_str.as_str()));
            if let Some(ht) = row.height {
                let ht_s = format_f64(ht);
                row_elem.push_attribute(("ht", ht_s.as_str()));
            }
            if row.hidden {
                row_elem.push_attribute(("hidden", "1"));
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
                    match cell.cell_type {
                        CellType::Number => {}
                        CellType::SharedString => c_elem.push_attribute(("t", "s")),
                        CellType::Boolean => c_elem.push_attribute(("t", "b")),
                        CellType::Error => c_elem.push_attribute(("t", "e")),
                        CellType::FormulaStr => c_elem.push_attribute(("t", "str")),
                        CellType::InlineStr => c_elem.push_attribute(("t", "inlineStr")),
                    }

                    // Only write s attribute if style_index is present.
                    if let Some(si) = cell.style_index {
                        let si_s = si.to_string();
                        c_elem.push_attribute(("s", si_s.as_str()));
                    }

                    let has_content = cell.formula.is_some() || cell.value.is_some();

                    if has_content {
                        writer
                            .write_event(Event::Start(c_elem))
                            .map_err(map_err)?;

                        // <f>...</f>
                        if let Some(ref formula) = cell.formula {
                            writer
                                .write_event(Event::Start(BytesStart::new("f")))
                                .map_err(map_err)?;
                            writer
                                .write_event(Event::Text(BytesText::new(formula)))
                                .map_err(map_err)?;
                            writer
                                .write_event(Event::End(BytesEnd::new("f")))
                                .map_err(map_err)?;
                        }

                        // <v>...</v>
                        if let Some(ref value) = cell.value {
                            writer
                                .write_event(Event::Start(BytesStart::new("v")))
                                .map_err(map_err)?;
                            writer
                                .write_event(Event::Text(BytesText::new(value)))
                                .map_err(map_err)?;
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
            let count_s = self.merge_cells.len().to_string();
            mc.push_attribute(("count", count_s.as_str()));
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
            elem.push_attribute(("ref", af.as_str()));
            writer.write_event(Event::Empty(elem)).map_err(map_err)?;
        }

        // </worksheet>
        writer
            .write_event(Event::End(BytesEnd::new("worksheet")))
            .map_err(map_err)?;

        String::from_utf8(buf)
            .map_err(|e| IronsheetError::XmlWrite(format!("invalid UTF-8 in output: {e}")))
    }
}

/// Parse a `<col>` element into a `ColumnInfo`.
fn parse_col_element(e: &BytesStart<'_>) -> ColumnInfo {
    let mut min: u32 = 1;
    let mut max: u32 = 1;
    let mut width: f64 = 8.43;
    let mut hidden = false;
    let mut custom_width = false;

    for attr in e.attributes().flatten() {
        let ln = attr.key.local_name();
        match ln.as_ref() {
            b"min" => {
                let val = String::from_utf8_lossy(&attr.value);
                min = val.parse::<u32>().unwrap_or(1);
            }
            b"max" => {
                let val = String::from_utf8_lossy(&attr.value);
                max = val.parse::<u32>().unwrap_or(1);
            }
            b"width" => {
                let val = String::from_utf8_lossy(&attr.value);
                width = val.parse::<f64>().unwrap_or(8.43);
            }
            b"hidden" => {
                let val = String::from_utf8_lossy(&attr.value);
                hidden = val == "1" || val.eq_ignore_ascii_case("true");
            }
            b"customWidth" => {
                let val = String::from_utf8_lossy(&attr.value);
                custom_width = val == "1" || val.eq_ignore_ascii_case("true");
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
    }
}

/// Convert a 1-based column index to a letter string (1 -> "A", 26 -> "Z", 27 -> "AA").
fn col_index_to_letter(col: u32) -> String {
    let mut result = String::new();
    let mut c = col;
    while c > 0 {
        c -= 1;
        result.insert(0, (b'A' + (c % 26) as u8) as char);
        c /= 26;
    }
    result
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
        assert_eq!(ws.auto_filter, Some("A1:D1".to_string()));
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
                        },
                        Cell {
                            reference: "B1".to_string(),
                            cell_type: CellType::Number,
                            style_index: None,
                            value: Some("42.5".to_string()),
                            formula: None,
                        },
                        Cell {
                            reference: "C1".to_string(),
                            cell_type: CellType::Number,
                            style_index: None,
                            value: Some("42.5".to_string()),
                            formula: Some("SUM(A1:B1)".to_string()),
                        },
                    ],
                    height: Some(18.0),
                    hidden: false,
                },
                Row {
                    index: 2,
                    cells: vec![Cell {
                        reference: "A2".to_string(),
                        cell_type: CellType::Boolean,
                        style_index: None,
                        value: Some("1".to_string()),
                        formula: None,
                    }],
                    height: None,
                    hidden: true,
                },
            ],
            merge_cells: vec!["A1:C1".to_string()],
            auto_filter: Some("A1:C1".to_string()),
            frozen_pane: Some(FrozenPane { rows: 1, cols: 0 }),
            columns: vec![ColumnInfo {
                min: 1,
                max: 1,
                width: 15.0,
                hidden: false,
                custom_width: true,
            }],
        };

        let xml = ws.to_xml().unwrap();
        let ws2 = WorksheetXml::parse(xml.as_bytes()).unwrap();

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
        assert_eq!(ws2.auto_filter, Some("A1:C1".to_string()));
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
                }],
                height: None,
                hidden: false,
            }],
            merge_cells: Vec::new(),
            auto_filter: None,
            frozen_pane: None,
            columns: Vec::new(),
        };

        let xml = ws.to_xml().unwrap();

        // Verify no optional sections appear.
        assert!(!xml.contains("sheetViews"));
        assert!(!xml.contains("<cols"));
        assert!(!xml.contains("mergeCells"));
        assert!(!xml.contains("autoFilter"));

        // But sheetData is present.
        assert!(xml.contains("sheetData"));

        // Re-parse.
        let ws2 = WorksheetXml::parse(xml.as_bytes()).unwrap();
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
                }],
                height: None,
                hidden: false,
            }],
            merge_cells: Vec::new(),
            auto_filter: None,
            frozen_pane: None,
            columns: Vec::new(),
        };

        let xml = ws.to_xml().unwrap();
        // The <c> element for a Number cell should NOT have a t="..." attribute.
        assert!(xml.contains(r#"<c r="A1"><v>42</v></c>"#));
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
                }],
                height: None,
                hidden: false,
            }],
            merge_cells: Vec::new(),
            auto_filter: None,
            frozen_pane: None,
            columns: Vec::new(),
        };

        let xml = ws.to_xml().unwrap();
        assert!(xml.contains(r#"t="s""#));
        assert!(xml.contains(r#"s="1""#));
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
}
