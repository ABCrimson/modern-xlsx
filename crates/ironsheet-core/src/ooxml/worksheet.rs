use quick_xml::events::{BytesStart, Event};
use quick_xml::Reader;

use crate::{IronsheetError, Result};

/// Parsed representation of a worksheet XML file (`xl/worksheets/sheet*.xml`).
#[derive(Debug, Clone)]
pub struct WorksheetXml {
    pub dimension: Option<String>,
    pub rows: Vec<Row>,
    pub merge_cells: Vec<String>,
    pub auto_filter: Option<String>,
    pub frozen_pane: Option<FrozenPane>,
    pub columns: Vec<ColumnInfo>,
}

/// A single row in the worksheet.
#[derive(Debug, Clone)]
pub struct Row {
    /// 1-based row index (as in the XML `r` attribute).
    pub index: u32,
    pub cells: Vec<Cell>,
    pub height: Option<f64>,
    pub hidden: bool,
}

/// A single cell in a row.
#[derive(Debug, Clone)]
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
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
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
#[derive(Debug, Clone)]
pub struct FrozenPane {
    /// Number of frozen rows (ySplit).
    pub rows: u32,
    /// Number of frozen columns (xSplit).
    pub cols: u32,
}

/// Column formatting information from the `<cols>` section.
#[derive(Debug, Clone)]
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
