use super::*;
use pretty_assertions::assert_eq;
use std::assert_matches;

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
fn test_parse_horizontal_split_pane() {
    let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetViews>
    <sheetView workbookViewId="0">
      <pane ySplit="2400" topLeftCell="A5" activePane="bottomLeft" state="split"/>
      <selection pane="bottomLeft" activeCell="A5" sqref="A5"/>
    </sheetView>
  </sheetViews>
  <sheetData/>
</worksheet>"#;
    let ws = WorksheetXml::parse(xml.as_bytes()).unwrap();
    assert!(
        ws.frozen_pane.is_none(),
        "split pane must not set frozen_pane"
    );
    let sp = ws.split_pane.as_ref().expect("split_pane should be Some");
    assert!((sp.horizontal.unwrap() - 2400.0).abs() < f64::EPSILON);
    assert!(sp.vertical.is_none());
    assert_eq!(sp.top_left_cell.as_deref(), Some("A5"));
    assert_eq!(sp.active_pane.as_deref(), Some("bottomLeft"));
    assert_eq!(ws.pane_selections.len(), 1);
    assert_eq!(ws.pane_selections[0].pane.as_deref(), Some("bottomLeft"));
    assert_eq!(ws.pane_selections[0].active_cell.as_deref(), Some("A5"));
    assert_eq!(ws.pane_selections[0].sqref.as_deref(), Some("A5"));
}

#[test]
fn test_parse_vertical_split_pane() {
    let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetViews>
    <sheetView workbookViewId="0">
      <pane xSplit="3000" topLeftCell="D1" activePane="topRight" state="split"/>
      <selection pane="topRight" activeCell="D1" sqref="D1"/>
    </sheetView>
  </sheetViews>
  <sheetData/>
</worksheet>"#;
    let ws = WorksheetXml::parse(xml.as_bytes()).unwrap();
    assert!(ws.frozen_pane.is_none());
    let sp = ws.split_pane.as_ref().unwrap();
    assert!(sp.horizontal.is_none());
    assert!((sp.vertical.unwrap() - 3000.0).abs() < f64::EPSILON);
    assert_eq!(sp.top_left_cell.as_deref(), Some("D1"));
    assert_eq!(sp.active_pane.as_deref(), Some("topRight"));
}

#[test]
fn test_parse_four_way_split_pane() {
    let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetViews>
    <sheetView workbookViewId="0">
      <pane xSplit="3000" ySplit="2400" topLeftCell="D5" activePane="bottomRight" state="split"/>
      <selection pane="topRight" activeCell="D1" sqref="D1"/>
      <selection pane="bottomLeft" activeCell="A5" sqref="A5"/>
      <selection pane="bottomRight" activeCell="D5" sqref="D5"/>
    </sheetView>
  </sheetViews>
  <sheetData/>
</worksheet>"#;
    let ws = WorksheetXml::parse(xml.as_bytes()).unwrap();
    assert!(ws.frozen_pane.is_none());
    let sp = ws.split_pane.as_ref().unwrap();
    assert!((sp.vertical.unwrap() - 3000.0).abs() < f64::EPSILON);
    assert!((sp.horizontal.unwrap() - 2400.0).abs() < f64::EPSILON);
    assert_eq!(sp.top_left_cell.as_deref(), Some("D5"));
    assert_eq!(sp.active_pane.as_deref(), Some("bottomRight"));
    assert_eq!(ws.pane_selections.len(), 3);
}

#[test]
fn test_frozen_and_split_mutually_exclusive() {
    let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetViews>
    <sheetView workbookViewId="0">
      <pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>
    </sheetView>
  </sheetViews>
  <sheetData/>
</worksheet>"#;
    let ws = WorksheetXml::parse(xml.as_bytes()).unwrap();
    assert!(ws.frozen_pane.is_some());
    assert!(ws.split_pane.is_none());
}

#[test]
fn test_parse_split_no_state_attr() {
    let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetViews>
    <sheetView workbookViewId="0">
      <pane xSplit="1500" ySplit="900" topLeftCell="B2"/>
    </sheetView>
  </sheetViews>
  <sheetData/>
</worksheet>"#;
    let ws = WorksheetXml::parse(xml.as_bytes()).unwrap();
    assert!(ws.frozen_pane.is_none());
    let sp = ws.split_pane.as_ref().unwrap();
    assert!((sp.vertical.unwrap() - 1500.0).abs() < f64::EPSILON);
    assert!((sp.horizontal.unwrap() - 900.0).abs() < f64::EPSILON);
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
    let af = ws.auto_filter.as_ref().unwrap();
    assert_eq!(af.range, "A1:D1");
    assert!(af.filter_columns.is_empty());
}

#[test]
fn test_write_full_roundtrip() {
    let ws = WorksheetXml {
        rows: vec![
            Row {
                index: 1,
                cells: vec![
                    Cell {
                        reference: "A1".to_string(),
                        cell_type: CellType::SharedString,
                        style_index: Some(0),
                        value: Some("0".to_string()),
                        ..Default::default()
                    },
                    Cell {
                        reference: "B1".to_string(),
                        cell_type: CellType::Number,
                        style_index: None,
                        value: Some("42.5".to_string()),
                        ..Default::default()
                    },
                    Cell {
                        reference: "C1".to_string(),
                        cell_type: CellType::Number,
                        value: Some("42.5".to_string()),
                        formula: Some("SUM(A1:B1)".to_string()),
                        ..Default::default()
                    },
                ],
                height: Some(18.0),
                ..Default::default()
            },
            Row {
                index: 2,
                cells: vec![Cell {
                    reference: "A2".to_string(),
                    cell_type: CellType::Boolean,
                    style_index: None,
                    value: Some("1".to_string()),
                    ..Default::default()
                }],
                hidden: true,
                ..Default::default()
            },
        ],
        merge_cells: vec!["A1:C1".to_string()],
        auto_filter: Some(AutoFilter {
            range: "A1:C1".to_string(),
            filter_columns: vec![],
        }),
        frozen_pane: Some(FrozenPane { rows: 1, cols: 0 }),
        columns: vec![ColumnInfo {
            min: 1,
            max: 1,
            width: 15.0,
            hidden: false,
            custom_width: true,
            outline_level: None,
            collapsed: false,
            style_index: None,
            best_fit: false,
            phonetic: false,
        }],
        ..Default::default()
    };

    let xml = ws.to_xml().unwrap();
    let ws2 = WorksheetXml::parse(&xml).unwrap();

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
    assert_eq!(ws2.auto_filter.as_ref().unwrap().range, "A1:C1");
}

#[test]
fn test_write_minimal_worksheet() {
    let ws = WorksheetXml {
        rows: vec![Row {
            index: 1,
            cells: vec![Cell {
                reference: "A1".to_string(),
                cell_type: CellType::Number,
                style_index: None,
                value: Some("100".to_string()),
                ..Default::default()
            }],
            ..Default::default()
        }],
        ..Default::default()
    };

    let xml = ws.to_xml().unwrap();
    let xml_str = String::from_utf8(xml.clone()).unwrap();

    // Verify no optional sections appear; the sheet view is Excel's default.
    assert!(xml_str.contains(DEFAULT_SHEET_VIEWS));
    assert!(!xml_str.contains("<cols"));
    assert!(!xml_str.contains("mergeCells"));
    assert!(!xml_str.contains("autoFilter"));

    // But sheetData is present.
    assert!(xml_str.contains("sheetData"));

    // Re-parse.
    let ws2 = WorksheetXml::parse(&xml).unwrap();
    assert_eq!(ws2.rows.len(), 1);
    assert_eq!(ws2.rows[0].cells[0].value, Some("100".to_string()));
    assert!(ws2.frozen_pane.is_none());
    assert!(ws2.columns.is_empty());
    assert!(ws2.merge_cells.is_empty());
    assert!(ws2.auto_filter.is_none());
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
        rows: vec![Row {
            index: 1,
            cells: vec![Cell {
                reference: "A1".to_string(),
                cell_type: CellType::Number,
                style_index: None,
                value: Some("42".to_string()),
                ..Default::default()
            }],
            ..Default::default()
        }],
        ..Default::default()
    };

    let xml = ws.to_xml().unwrap();
    let xml_str = String::from_utf8(xml).unwrap();
    // The <c> element for a Number cell should NOT have a t="..." attribute.
    assert!(xml_str.contains(r#"<c r="A1"><v>42</v></c>"#));
}

#[test]
fn test_write_emits_t_for_shared_string() {
    let ws = WorksheetXml {
        rows: vec![Row {
            index: 1,
            cells: vec![Cell {
                reference: "A1".to_string(),
                cell_type: CellType::SharedString,
                style_index: Some(1),
                value: Some("5".to_string()),
                ..Default::default()
            }],
            ..Default::default()
        }],
        ..Default::default()
    };

    let xml = ws.to_xml().unwrap();
    let xml_str = String::from_utf8(xml).unwrap();
    assert!(xml_str.contains(r#"t="s""#));
    assert!(xml_str.contains(r#"s="1""#));
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

#[test]
fn test_parse_data_validations() {
    let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData/>
  <dataValidations count="2">
    <dataValidation type="list" allowBlank="1" showErrorMessage="1" sqref="A1:A10">
      <formula1>"Yes,No,Maybe"</formula1>
    </dataValidation>
    <dataValidation type="whole" operator="between" sqref="B1:B10" errorTitle="Invalid" error="Must be 1-100">
      <formula1>1</formula1>
      <formula2>100</formula2>
    </dataValidation>
  </dataValidations>
</worksheet>"#;

    let ws = WorksheetXml::parse(xml.as_bytes()).unwrap();
    assert_eq!(ws.data_validations.len(), 2);

    let dv0 = &ws.data_validations[0];
    assert_eq!(dv0.sqref, "A1:A10");
    assert_eq!(dv0.validation_type, Some("list".to_string()));
    assert_eq!(dv0.operator, None);
    assert_eq!(dv0.allow_blank, Some(true));
    assert_eq!(dv0.show_error_message, Some(true));
    assert_eq!(dv0.formula1, Some("\"Yes,No,Maybe\"".to_string()));
    assert_eq!(dv0.formula2, None);

    let dv1 = &ws.data_validations[1];
    assert_eq!(dv1.sqref, "B1:B10");
    assert_eq!(dv1.validation_type, Some("whole".to_string()));
    assert_eq!(dv1.operator, Some("between".to_string()));
    assert_eq!(dv1.error_title, Some("Invalid".to_string()));
    assert_eq!(dv1.error_message, Some("Must be 1-100".to_string()));
    assert_eq!(dv1.formula1, Some("1".to_string()));
    assert_eq!(dv1.formula2, Some("100".to_string()));
}

#[test]
fn test_data_validations_roundtrip() {
    let ws = WorksheetXml {
        data_validations: vec![
            DataValidation {
                sqref: "A1:A10".to_string(),
                validation_type: Some("list".to_string()),
                operator: None,
                formula1: Some("\"Yes,No,Maybe\"".to_string()),
                formula2: None,
                allow_blank: Some(true),
                show_error_message: Some(true),
                error_title: None,
                error_message: None,
                show_input_message: None,
                prompt_title: None,
                prompt: None,
                error_style: None,
                ime_mode: None,
                show_drop_down: None,
            },
            DataValidation {
                sqref: "B1:B10".to_string(),
                validation_type: Some("whole".to_string()),
                operator: Some("between".to_string()),
                formula1: Some("1".to_string()),
                formula2: Some("100".to_string()),
                allow_blank: None,
                show_error_message: None,
                error_title: Some("Invalid".to_string()),
                error_message: Some("Must be 1-100".to_string()),
                show_input_message: None,
                prompt_title: None,
                prompt: None,
                error_style: None,
                ime_mode: None,
                show_drop_down: None,
            },
        ],
        ..Default::default()
    };

    let xml = ws.to_xml().unwrap();
    let xml_str = String::from_utf8(xml.clone()).unwrap();

    // Verify XML contains data validations.
    assert!(xml_str.contains("dataValidations"));
    assert!(xml_str.contains(r#"count="2""#));
    assert!(xml_str.contains(r#"type="list""#));
    assert!(xml_str.contains(r#"sqref="A1:A10""#));

    // Re-parse and verify.
    let ws2 = WorksheetXml::parse(&xml).unwrap();
    assert_eq!(ws2.data_validations.len(), 2);

    let dv0 = &ws2.data_validations[0];
    assert_eq!(dv0.sqref, "A1:A10");
    assert_eq!(dv0.validation_type, Some("list".to_string()));
    assert_eq!(dv0.allow_blank, Some(true));
    assert_eq!(dv0.show_error_message, Some(true));
    assert_eq!(dv0.formula1, Some("\"Yes,No,Maybe\"".to_string()));
    assert_eq!(dv0.formula2, None);

    let dv1 = &ws2.data_validations[1];
    assert_eq!(dv1.sqref, "B1:B10");
    assert_eq!(dv1.validation_type, Some("whole".to_string()));
    assert_eq!(dv1.operator, Some("between".to_string()));
    assert_eq!(dv1.formula1, Some("1".to_string()));
    assert_eq!(dv1.formula2, Some("100".to_string()));
    assert_eq!(dv1.error_title, Some("Invalid".to_string()));
    assert_eq!(dv1.error_message, Some("Must be 1-100".to_string()));
}

#[test]
fn test_parse_conditional_formatting() {
    let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1"><v>200</v></c>
    </row>
  </sheetData>
  <conditionalFormatting sqref="A1:B10">
    <cfRule type="cellIs" dxfId="0" priority="1" operator="greaterThan">
      <formula>100</formula>
    </cfRule>
  </conditionalFormatting>
  <conditionalFormatting sqref="C1:C10">
    <cfRule type="colorScale" priority="2">
      <colorScale>
        <cfvo type="min"/>
        <cfvo type="max"/>
        <color rgb="FFF8696B"/>
        <color rgb="FF63BE7B"/>
      </colorScale>
    </cfRule>
  </conditionalFormatting>
</worksheet>"#;

    let ws = WorksheetXml::parse(xml.as_bytes()).unwrap();
    assert_eq!(ws.conditional_formatting.len(), 2);

    // First block.
    let cf0 = &ws.conditional_formatting[0];
    assert_eq!(cf0.sqref, "A1:B10");
    assert_eq!(cf0.rules.len(), 1);
    assert_eq!(cf0.rules[0].rule_type, "cellIs");
    assert_eq!(cf0.rules[0].priority, 1);
    assert_eq!(cf0.rules[0].operator, Some("greaterThan".to_string()));
    assert_eq!(cf0.rules[0].dxf_id, Some(0));
    assert_eq!(cf0.rules[0].formula, Some("100".to_string()));

    // Second block (colorScale — no formula at cfRule level).
    let cf1 = &ws.conditional_formatting[1];
    assert_eq!(cf1.sqref, "C1:C10");
    assert_eq!(cf1.rules.len(), 1);
    assert_eq!(cf1.rules[0].rule_type, "colorScale");
    assert_eq!(cf1.rules[0].priority, 2);
    assert_eq!(cf1.rules[0].operator, None);
    assert_eq!(cf1.rules[0].dxf_id, None);
    assert_eq!(cf1.rules[0].formula, None);
}

#[test]
fn test_conditional_formatting_roundtrip() {
    let ws = WorksheetXml {
        rows: vec![Row {
            index: 1,
            cells: vec![Cell {
                reference: "A1".to_string(),
                cell_type: CellType::Number,
                style_index: None,
                value: Some("200".to_string()),
                ..Default::default()
            }],
            ..Default::default()
        }],
        conditional_formatting: vec![
            ConditionalFormatting {
                sqref: "A1:B10".to_string(),
                rules: vec![ConditionalFormattingRule {
                    rule_type: "cellIs".to_string(),
                    priority: 1,
                    operator: Some("greaterThan".to_string()),
                    formula: Some("100".to_string()),
                    additional_formulas: Vec::new(),
                    dxf_id: Some(0),
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
                }],
            },
            ConditionalFormatting {
                sqref: "C1:C10".to_string(),
                rules: vec![ConditionalFormattingRule {
                    rule_type: "colorScale".to_string(),
                    priority: 2,
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
                }],
            },
        ],
        ..Default::default()
    };

    // Write to XML.
    let xml = ws.to_xml().unwrap();
    let xml_str = String::from_utf8(xml.clone()).unwrap();

    // Verify the XML contains conditional formatting elements.
    assert!(xml_str.contains("conditionalFormatting"));
    assert!(xml_str.contains(r#"sqref="A1:B10"#));
    assert!(xml_str.contains(r#"type="cellIs""#));
    assert!(xml_str.contains(r#"operator="greaterThan""#));
    assert!(xml_str.contains("<formula>100</formula>"));

    // Parse back.
    let ws2 = WorksheetXml::parse(&xml).unwrap();
    assert_eq!(ws2.conditional_formatting.len(), 2);

    // Verify first block.
    let cf0 = &ws2.conditional_formatting[0];
    assert_eq!(cf0.sqref, "A1:B10");
    assert_eq!(cf0.rules.len(), 1);
    assert_eq!(cf0.rules[0].rule_type, "cellIs");
    assert_eq!(cf0.rules[0].priority, 1);
    assert_eq!(cf0.rules[0].operator, Some("greaterThan".to_string()));
    assert_eq!(cf0.rules[0].dxf_id, Some(0));
    assert_eq!(cf0.rules[0].formula, Some("100".to_string()));

    // Verify second block.
    let cf1 = &ws2.conditional_formatting[1];
    assert_eq!(cf1.sqref, "C1:C10");
    assert_eq!(cf1.rules.len(), 1);
    assert_eq!(cf1.rules[0].rule_type, "colorScale");
    assert_eq!(cf1.rules[0].priority, 2);
    assert_eq!(cf1.rules[0].operator, None);
    assert_eq!(cf1.rules[0].dxf_id, None);
    assert_eq!(cf1.rules[0].formula, None);
}

#[test]
fn test_parse_hyperlinks() {
    let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData/>
  <hyperlinks>
    <hyperlink ref="A1" location="Sheet2!A1" display="Go to Sheet2" tooltip="Click to navigate"/>
    <hyperlink ref="B1" location="'My Sheet'!C3"/>
  </hyperlinks>
</worksheet>"#;

    let ws = WorksheetXml::parse(xml.as_bytes()).unwrap();
    assert_eq!(ws.hyperlinks.len(), 2);

    assert_eq!(ws.hyperlinks[0].cell_ref, "A1");
    assert_eq!(ws.hyperlinks[0].location, Some("Sheet2!A1".to_string()));
    assert_eq!(ws.hyperlinks[0].display, Some("Go to Sheet2".to_string()));
    assert_eq!(
        ws.hyperlinks[0].tooltip,
        Some("Click to navigate".to_string())
    );

    assert_eq!(ws.hyperlinks[1].cell_ref, "B1");
    assert_eq!(ws.hyperlinks[1].location, Some("'My Sheet'!C3".to_string()));
    assert_eq!(ws.hyperlinks[1].display, None);
    assert_eq!(ws.hyperlinks[1].tooltip, None);
}

#[test]
fn test_hyperlinks_roundtrip() {
    let ws = WorksheetXml {
        hyperlinks: vec![
            Hyperlink {
                cell_ref: "A1".to_string(),
                location: Some("Sheet2!A1".to_string()),
                display: Some("Go to Sheet2".to_string()),
                tooltip: Some("Click here".to_string()),
                ..Default::default()
            },
            Hyperlink {
                cell_ref: "B1".to_string(),
                location: Some("Sheet3!B2".to_string()),
                ..Default::default()
            },
        ],
        ..Default::default()
    };

    let xml = ws.to_xml().unwrap();
    let ws2 = WorksheetXml::parse(&xml).unwrap();
    assert_eq!(ws2.hyperlinks.len(), 2);
    assert_eq!(ws2.hyperlinks[0].cell_ref, "A1");
    assert_eq!(ws2.hyperlinks[0].location, Some("Sheet2!A1".to_string()));
    assert_eq!(ws2.hyperlinks[0].display, Some("Go to Sheet2".to_string()));
    assert_eq!(ws2.hyperlinks[0].tooltip, Some("Click here".to_string()));
    assert_eq!(ws2.hyperlinks[1].cell_ref, "B1");
    assert_eq!(ws2.hyperlinks[1].location, Some("Sheet3!B2".to_string()));
}

#[test]
fn test_parse_page_setup() {
    let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData/>
  <pageSetup paperSize="9" orientation="landscape" scale="75" horizontalDpi="300" verticalDpi="300" fitToWidth="1" fitToHeight="0"/>
</worksheet>"#;

    let ws = WorksheetXml::parse(xml.as_bytes()).unwrap();
    let ps = ws.page_setup.as_ref().unwrap();
    assert_eq!(ps.paper_size, Some(9));
    assert_eq!(ps.orientation, Some("landscape".to_string()));
    assert_eq!(ps.scale, Some(75));
    assert_eq!(ps.horizontal_dpi, Some(300));
    assert_eq!(ps.vertical_dpi, Some(300));
    assert_eq!(ps.fit_to_width, Some(1));
    assert_eq!(ps.fit_to_height, Some(0));
}

#[test]
fn test_page_setup_roundtrip() {
    let ws = WorksheetXml {
        page_setup: Some(PageSetup {
            paper_size: Some(1),
            orientation: Some("portrait".to_string()),
            fit_to_width: None,
            fit_to_height: None,
            scale: Some(100),
            first_page_number: None,
            horizontal_dpi: Some(600),
            vertical_dpi: Some(600),
            fit_to_page: None,
            margins: None,
        }),
        ..Default::default()
    };

    let xml = ws.to_xml().unwrap();
    let ws2 = WorksheetXml::parse(&xml).unwrap();
    let ps = ws2.page_setup.as_ref().unwrap();
    assert_eq!(ps.paper_size, Some(1));
    assert_eq!(ps.orientation, Some("portrait".to_string()));
    assert_eq!(ps.scale, Some(100));
    assert_eq!(ps.horizontal_dpi, Some(600));
    assert_eq!(ps.vertical_dpi, Some(600));
}

#[test]
fn test_parse_sheet_protection() {
    let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData/>
  <sheetProtection sheet="1" objects="1" scenarios="1" password="ABCD" formatCells="1" insertRows="1" deleteRows="1" sort="1" autoFilter="1"/>
</worksheet>"#;

    let ws = WorksheetXml::parse(xml.as_bytes()).unwrap();
    let sp = ws.sheet_protection.as_ref().unwrap();
    assert!(sp.sheet);
    assert!(sp.objects);
    assert!(sp.scenarios);
    assert_eq!(sp.password, Some("ABCD".to_string()));
    assert!(sp.format_cells);
    // Absent flags take the CT_SheetProtection default, true (locked).
    assert!(sp.format_columns);
    assert!(sp.format_rows);
    assert!(sp.insert_columns);
    assert!(sp.insert_rows);
    assert!(sp.delete_columns);
    assert!(sp.delete_rows);
    assert!(sp.sort);
    assert!(sp.auto_filter);
}

#[test]
fn test_sheet_protection_roundtrip() {
    let ws = WorksheetXml {
        sheet_protection: Some(SheetProtection {
            sheet: true,
            objects: true,
            scenarios: false,
            password: Some("CBEB".to_string()),
            format_cells: false,
            format_columns: false,
            format_rows: false,
            insert_columns: false,
            insert_rows: true,
            delete_columns: false,
            delete_rows: true,
            sort: true,
            auto_filter: true,
            ..Default::default()
        }),
        ..Default::default()
    };

    let xml = ws.to_xml().unwrap();
    let ws2 = WorksheetXml::parse(&xml).unwrap();
    let sp = ws2.sheet_protection.as_ref().unwrap();
    assert!(sp.sheet);
    assert!(sp.objects);
    assert!(!sp.scenarios);
    assert_eq!(sp.password, Some("CBEB".to_string()));
    assert!(sp.insert_rows);
    assert!(sp.delete_rows);
    assert!(sp.sort);
    assert!(sp.auto_filter);
    assert!(!sp.format_cells && !sp.format_columns && !sp.format_rows);
    assert!(!sp.insert_columns && !sp.delete_columns);
}

#[test]
fn test_parse_array_formula() {
    let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1"><f t="array" ref="A1:A3">SUM(B1:B3*C1:C3)</f><v>100</v></c>
    </row>
  </sheetData>
</worksheet>"#;

    let ws = WorksheetXml::parse(xml.as_bytes()).unwrap();
    let cell = &ws.rows[0].cells[0];
    assert_eq!(cell.formula, Some("SUM(B1:B3*C1:C3)".to_string()));
    assert_eq!(cell.formula_type, Some("array".to_string()));
    assert_eq!(cell.formula_ref, Some("A1:A3".to_string()));
    assert_eq!(cell.shared_index, None);
    assert_eq!(cell.value, Some("100".to_string()));
}

#[test]
fn test_parse_shared_formula() {
    let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1"><f t="shared" ref="A1:A3" si="0">B1+C1</f><v>10</v></c>
    </row>
    <row r="2">
      <c r="A2"><f t="shared" si="0"/><v>20</v></c>
    </row>
  </sheetData>
</worksheet>"#;

    let ws = WorksheetXml::parse(xml.as_bytes()).unwrap();
    let cell0 = &ws.rows[0].cells[0];
    assert_eq!(cell0.formula, Some("B1+C1".to_string()));
    assert_eq!(cell0.formula_type, Some("shared".to_string()));
    assert_eq!(cell0.formula_ref, Some("A1:A3".to_string()));
    assert_eq!(cell0.shared_index, Some(0));

    let cell1 = &ws.rows[1].cells[0];
    assert_eq!(cell1.formula, None);
    assert_eq!(cell1.formula_type, Some("shared".to_string()));
    assert_eq!(cell1.shared_index, Some(0));
}

#[test]
fn test_array_formula_roundtrip() {
    let ws = WorksheetXml {
        rows: vec![Row {
            index: 1,
            cells: vec![Cell {
                reference: "A1".to_string(),
                cell_type: CellType::Number,
                value: Some("100".to_string()),
                formula: Some("SUM(B1:B3*C1:C3)".to_string()),
                formula_type: Some("array".to_string()),
                formula_ref: Some("A1:A3".to_string()),
                ..Default::default()
            }],
            ..Default::default()
        }],
        ..Default::default()
    };

    let xml = ws.to_xml().unwrap();
    let xml_str = String::from_utf8(xml.clone()).unwrap();
    assert!(xml_str.contains(r#"t="array""#));
    assert!(xml_str.contains(r#"ref="A1:A3""#));

    let ws2 = WorksheetXml::parse(&xml).unwrap();
    let cell = &ws2.rows[0].cells[0];
    assert_eq!(cell.formula_type, Some("array".to_string()));
    assert_eq!(cell.formula_ref, Some("A1:A3".to_string()));
    assert_eq!(cell.formula, Some("SUM(B1:B3*C1:C3)".to_string()));
}

/// Serialize a one-cell sheet and return the XML text.
fn formula_cell_xml(formula: &str, formula_type: Option<&str>) -> String {
    let ws = WorksheetXml {
        rows: vec![Row {
            index: 1,
            cells: vec![Cell {
                reference: "A1".to_string(),
                formula: Some(formula.to_string()),
                formula_type: formula_type.map(str::to_string),
                ..Default::default()
            }],
            ..Default::default()
        }],
        ..Default::default()
    };
    String::from_utf8(ws.to_xml().unwrap()).unwrap()
}

#[test]
fn test_formula_text_written_without_ui_braces_or_equals() {
    // Array formula entered with the Excel UI braces: only the inner text is written.
    let xml = formula_cell_xml("{=SUM(B1:B3*C1:C3)}", Some("array"));
    assert!(xml.contains(">SUM(B1:B3*C1:C3)</f>"), "{xml}");
    assert!(!xml.contains("{="), "{xml}");

    // An array constant is not a UI-brace wrapper and stays verbatim.
    let xml = formula_cell_xml("{1,2,3}", None);
    assert!(xml.contains(">{1,2,3}</f>"), "{xml}");
    let xml = formula_cell_xml("{1,2,3}", Some("array"));
    assert!(xml.contains(">{1,2,3}</f>"), "{xml}");

    // A plain formula loses exactly one leading '='.
    let xml = formula_cell_xml("=A1+1", None);
    assert!(xml.contains(">A1+1</f>"), "{xml}");
    let xml = formula_cell_xml("==A1", None);
    assert!(xml.contains(">=A1</f>"), "{xml}");
}

#[test]
fn test_parse_inline_string() {
    let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="inlineStr"><is><t>Hello World</t></is></c>
    </row>
  </sheetData>
</worksheet>"#;

    let ws = WorksheetXml::parse(xml.as_bytes()).unwrap();
    let cell = &ws.rows[0].cells[0];
    assert_eq!(cell.cell_type, CellType::InlineStr);
    assert_eq!(cell.value, Some("Hello World".to_string()));
    assert_eq!(cell.inline_string, Some("Hello World".to_string()));
}

#[test]
fn test_parse_data_validation_prompts() {
    let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData/>
  <dataValidations count="1">
    <dataValidation type="list" allowBlank="1" showInputMessage="1" promptTitle="Choose" prompt="Pick a value" sqref="A1:A10">
      <formula1>"Yes,No"</formula1>
    </dataValidation>
  </dataValidations>
</worksheet>"#;

    let ws = WorksheetXml::parse(xml.as_bytes()).unwrap();
    let dv = &ws.data_validations[0];
    assert_eq!(dv.show_input_message, Some(true));
    assert_eq!(dv.prompt_title, Some("Choose".to_string()));
    assert_eq!(dv.prompt, Some("Pick a value".to_string()));
}

#[test]
fn test_data_validation_prompts_roundtrip() {
    let ws = WorksheetXml {
        data_validations: vec![DataValidation {
            sqref: "A1:A10".to_string(),
            validation_type: Some("list".to_string()),
            operator: None,
            formula1: Some("\"Yes,No\"".to_string()),
            formula2: None,
            allow_blank: Some(true),
            show_error_message: None,
            error_title: None,
            error_message: None,
            show_input_message: Some(true),
            prompt_title: Some("Choose".to_string()),
            prompt: Some("Pick a value".to_string()),
            error_style: None,
            ime_mode: None,
            show_drop_down: None,
        }],
        ..Default::default()
    };

    let xml = ws.to_xml().unwrap();
    let ws2 = WorksheetXml::parse(&xml).unwrap();
    let dv = &ws2.data_validations[0];
    assert_eq!(dv.show_input_message, Some(true));
    assert_eq!(dv.prompt_title, Some("Choose".to_string()));
    assert_eq!(dv.prompt, Some("Pick a value".to_string()));
}

#[test]
fn test_parse_auto_filter_with_columns() {
    let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData/>
  <autoFilter ref="A1:D10">
    <filterColumn colId="0">
      <filters>
        <filter val="Apple"/>
        <filter val="Banana"/>
      </filters>
    </filterColumn>
    <filterColumn colId="2">
      <filters>
        <filter val="Red"/>
      </filters>
    </filterColumn>
  </autoFilter>
</worksheet>"#;

    let ws = WorksheetXml::parse(xml.as_bytes()).unwrap();
    let af = ws.auto_filter.as_ref().unwrap();
    assert_eq!(af.range, "A1:D10");
    assert_eq!(af.filter_columns.len(), 2);
    assert_eq!(af.filter_columns[0].col_id, 0);
    assert_eq!(af.filter_columns[0].filters, vec!["Apple", "Banana"]);
    assert_eq!(af.filter_columns[1].col_id, 2);
    assert_eq!(af.filter_columns[1].filters, vec!["Red"]);
}

#[test]
fn test_auto_filter_with_columns_roundtrip() {
    let ws = WorksheetXml {
        auto_filter: Some(AutoFilter {
            range: "A1:D10".to_string(),
            filter_columns: vec![
                FilterColumn {
                    col_id: 0,
                    filters: vec!["Apple".to_string(), "Banana".to_string()],
                    custom_filters: None,
                },
                FilterColumn {
                    col_id: 2,
                    filters: vec!["Red".to_string()],
                    custom_filters: None,
                },
            ],
        }),
        ..Default::default()
    };

    let xml = ws.to_xml().unwrap();
    let ws2 = WorksheetXml::parse(&xml).unwrap();
    let af = ws2.auto_filter.as_ref().unwrap();
    assert_eq!(af.range, "A1:D10");
    assert_eq!(af.filter_columns.len(), 2);
    assert_eq!(af.filter_columns[0].col_id, 0);
    assert_eq!(af.filter_columns[0].filters, vec!["Apple", "Banana"]);
    assert_eq!(af.filter_columns[1].col_id, 2);
    assert_eq!(af.filter_columns[1].filters, vec!["Red"]);
}

#[test]
fn test_parse_auto_filter_custom_filters() {
    let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData/>
  <autoFilter ref="A1:D10">
    <filterColumn colId="1">
      <customFilters and="1">
        <customFilter operator="greaterThanOrEqual" val="10"/>
        <customFilter operator="lessThan" val="100"/>
      </customFilters>
    </filterColumn>
  </autoFilter>
</worksheet>"#;

    let ws = WorksheetXml::parse(xml.as_bytes()).unwrap();
    let af = ws.auto_filter.as_ref().unwrap();
    assert_eq!(af.range, "A1:D10");
    assert_eq!(af.filter_columns.len(), 1);
    let fc = &af.filter_columns[0];
    assert_eq!(fc.col_id, 1);
    assert!(fc.filters.is_empty());
    let cf = fc.custom_filters.as_ref().unwrap();
    assert!(cf.and_op);
    assert_eq!(cf.filters.len(), 2);
    assert_eq!(
        cf.filters[0].operator.as_deref(),
        Some("greaterThanOrEqual")
    );
    assert_eq!(cf.filters[0].val, "10");
    assert_eq!(cf.filters[1].operator.as_deref(), Some("lessThan"));
    assert_eq!(cf.filters[1].val, "100");
}

#[test]
fn test_auto_filter_custom_filters_roundtrip() {
    let ws = WorksheetXml {
        auto_filter: Some(AutoFilter {
            range: "B2:E20".to_string(),
            filter_columns: vec![
                FilterColumn {
                    col_id: 0,
                    filters: vec!["Yes".to_string()],
                    custom_filters: None,
                },
                FilterColumn {
                    col_id: 3,
                    filters: vec![],
                    custom_filters: Some(CustomFilters {
                        and_op: false,
                        filters: vec![CustomFilter {
                            operator: Some("greaterThan".to_string()),
                            val: "50".to_string(),
                        }],
                    }),
                },
                FilterColumn {
                    col_id: 1,
                    filters: vec![],
                    custom_filters: Some(CustomFilters {
                        and_op: true,
                        filters: vec![
                            CustomFilter {
                                operator: Some("greaterThanOrEqual".to_string()),
                                val: "5".to_string(),
                            },
                            CustomFilter {
                                operator: Some("lessThanOrEqual".to_string()),
                                val: "99".to_string(),
                            },
                        ],
                    }),
                },
            ],
        }),
        ..Default::default()
    };

    let xml = ws.to_xml().unwrap();
    let ws2 = WorksheetXml::parse(&xml).unwrap();
    let af = ws2.auto_filter.as_ref().unwrap();
    assert_eq!(af.range, "B2:E20");
    assert_eq!(af.filter_columns.len(), 3);

    // First: regular filters.
    assert_eq!(af.filter_columns[0].col_id, 0);
    assert_eq!(af.filter_columns[0].filters, vec!["Yes"]);
    assert!(af.filter_columns[0].custom_filters.is_none());

    // Second: single custom filter (OR).
    let cf1 = af.filter_columns[1].custom_filters.as_ref().unwrap();
    assert!(!cf1.and_op);
    assert_eq!(cf1.filters.len(), 1);
    assert_eq!(cf1.filters[0].operator.as_deref(), Some("greaterThan"));
    assert_eq!(cf1.filters[0].val, "50");

    // Third: two custom filters (AND).
    let cf2 = af.filter_columns[2].custom_filters.as_ref().unwrap();
    assert!(cf2.and_op);
    assert_eq!(cf2.filters.len(), 2);
    assert_eq!(
        cf2.filters[0].operator.as_deref(),
        Some("greaterThanOrEqual")
    );
    assert_eq!(cf2.filters[0].val, "5");
    assert_eq!(cf2.filters[1].operator.as_deref(), Some("lessThanOrEqual"));
    assert_eq!(cf2.filters[1].val, "99");
}

#[test]
fn test_auto_filter_custom_filter_no_operator() {
    let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData/>
  <autoFilter ref="A1:B5">
    <filterColumn colId="0">
      <customFilters>
        <customFilter val="*test*"/>
      </customFilters>
    </filterColumn>
  </autoFilter>
</worksheet>"#;

    let ws = WorksheetXml::parse(xml.as_bytes()).unwrap();
    let af = ws.auto_filter.as_ref().unwrap();
    let fc = &af.filter_columns[0];
    let cf = fc.custom_filters.as_ref().unwrap();
    assert!(!cf.and_op);
    assert_eq!(cf.filters.len(), 1);
    assert!(cf.filters[0].operator.is_none());
    assert_eq!(cf.filters[0].val, "*test*");

    // Roundtrip.
    let xml_out = ws.to_xml().unwrap();
    let ws2 = WorksheetXml::parse(&xml_out).unwrap();
    let cf2 = ws2.auto_filter.as_ref().unwrap().filter_columns[0]
        .custom_filters
        .as_ref()
        .unwrap();
    assert!(cf2.filters[0].operator.is_none());
    assert_eq!(cf2.filters[0].val, "*test*");
}

#[test]
fn test_parse_color_scale() {
    let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData/>
  <conditionalFormatting sqref="A1:A10">
    <cfRule type="colorScale" priority="1">
      <colorScale>
        <cfvo type="min"/>
        <cfvo type="max"/>
        <color rgb="FFF8696B"/>
        <color rgb="FF63BE7B"/>
      </colorScale>
    </cfRule>
  </conditionalFormatting>
</worksheet>"#;

    let ws = WorksheetXml::parse(xml.as_bytes()).unwrap();
    let rule = &ws.conditional_formatting[0].rules[0];
    assert_eq!(rule.rule_type, "colorScale");
    let cs = rule.color_scale.as_ref().unwrap();
    assert_eq!(cs.cfvos.len(), 2);
    assert_eq!(cs.cfvos[0].cfvo_type, "min");
    assert_eq!(cs.cfvos[0].val, None);
    assert_eq!(cs.cfvos[1].cfvo_type, "max");
    assert_eq!(cs.colors, vec!["FFF8696B", "FF63BE7B"]);
}

#[test]
fn test_color_scale_roundtrip() {
    let ws = WorksheetXml {
        conditional_formatting: vec![ConditionalFormatting {
            sqref: "A1:A10".to_string(),
            rules: vec![ConditionalFormattingRule {
                rule_type: "colorScale".to_string(),
                priority: 1,
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
                color_scale: Some(ColorScale {
                    cfvos: vec![
                        Cfvo {
                            cfvo_type: "min".to_string(),
                            val: None,
                        },
                        Cfvo {
                            cfvo_type: "max".to_string(),
                            val: None,
                        },
                    ],
                    colors: vec!["FFF8696B".to_string(), "FF63BE7B".to_string()],
                    color_specs: Vec::new(),
                }),
                data_bar: None,
                icon_set: None,
            }],
        }],
        ..Default::default()
    };

    let xml = ws.to_xml().unwrap();
    let ws2 = WorksheetXml::parse(&xml).unwrap();
    let cs = ws2.conditional_formatting[0].rules[0]
        .color_scale
        .as_ref()
        .unwrap();
    assert_eq!(cs.cfvos.len(), 2);
    assert_eq!(cs.cfvos[0].cfvo_type, "min");
    assert_eq!(cs.cfvos[1].cfvo_type, "max");
    assert_eq!(cs.colors, vec!["FFF8696B", "FF63BE7B"]);
}

#[test]
fn test_parse_data_bar() {
    let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData/>
  <conditionalFormatting sqref="B1:B10">
    <cfRule type="dataBar" priority="1">
      <dataBar>
        <cfvo type="min"/>
        <cfvo type="max"/>
        <color rgb="FF638EC6"/>
      </dataBar>
    </cfRule>
  </conditionalFormatting>
</worksheet>"#;

    let ws = WorksheetXml::parse(xml.as_bytes()).unwrap();
    let rule = &ws.conditional_formatting[0].rules[0];
    let db = rule.data_bar.as_ref().unwrap();
    assert_eq!(db.cfvos.len(), 2);
    assert_eq!(db.color, "FF638EC6");
}

#[test]
fn test_parse_icon_set() {
    let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData/>
  <conditionalFormatting sqref="C1:C10">
    <cfRule type="iconSet" priority="1">
      <iconSet iconSet="3Arrows">
        <cfvo type="percent" val="0"/>
        <cfvo type="percent" val="33"/>
        <cfvo type="percent" val="67"/>
      </iconSet>
    </cfRule>
  </conditionalFormatting>
</worksheet>"#;

    let ws = WorksheetXml::parse(xml.as_bytes()).unwrap();
    let rule = &ws.conditional_formatting[0].rules[0];
    let is = rule.icon_set.as_ref().unwrap();
    assert_eq!(is.icon_set_type, Some("3Arrows".to_string()));
    assert_eq!(is.cfvos.len(), 3);
    assert_eq!(is.cfvos[0].cfvo_type, "percent");
    assert_eq!(is.cfvos[0].val, Some("0".to_string()));
    assert_eq!(is.cfvos[1].val, Some("33".to_string()));
    assert_eq!(is.cfvos[2].val, Some("67".to_string()));
}

#[test]
fn test_icon_set_roundtrip() {
    let ws = WorksheetXml {
        conditional_formatting: vec![ConditionalFormatting {
            sqref: "C1:C10".to_string(),
            rules: vec![ConditionalFormattingRule {
                rule_type: "iconSet".to_string(),
                priority: 1,
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
                icon_set: Some(IconSet {
                    icon_set_type: Some("3TrafficLights1".to_string()),
                    cfvos: vec![
                        Cfvo {
                            cfvo_type: "percent".to_string(),
                            val: Some("0".to_string()),
                        },
                        Cfvo {
                            cfvo_type: "percent".to_string(),
                            val: Some("33".to_string()),
                        },
                        Cfvo {
                            cfvo_type: "percent".to_string(),
                            val: Some("67".to_string()),
                        },
                    ],
                }),
            }],
        }],
        ..Default::default()
    };

    let xml = ws.to_xml().unwrap();
    let ws2 = WorksheetXml::parse(&xml).unwrap();
    let is = ws2.conditional_formatting[0].rules[0]
        .icon_set
        .as_ref()
        .unwrap();
    assert_eq!(is.icon_set_type, Some("3TrafficLights1".to_string()));
    assert_eq!(is.cfvos.len(), 3);
}

#[test]
fn test_parse_header_footer() {
    let xml = r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData/>
  <headerFooter differentFirst="1">
    <oddHeader>&amp;CConfidential Report</oddHeader>
    <oddFooter>&amp;LPage &amp;P of &amp;N&amp;R&amp;D</oddFooter>
    <firstHeader>&amp;C&amp;BTitle Page</firstHeader>
  </headerFooter>
</worksheet>"#;
    let ws = WorksheetXml::parse(xml.as_bytes()).unwrap();
    let hf = ws.header_footer.as_ref().unwrap();
    assert_eq!(hf.odd_header.as_deref(), Some("&CConfidential Report"));
    assert_eq!(hf.odd_footer.as_deref(), Some("&LPage &P of &N&R&D"));
    assert_eq!(hf.first_header.as_deref(), Some("&C&BTitle Page"));
    assert!(hf.different_first);
    assert!(!hf.different_odd_even);
}

#[test]
fn header_footer_all_six_parts_parse() {
    let xml = sheet(
        r#"<sheetData/><headerFooter differentOddEven="1" differentFirst="1"><oddHeader>&amp;Codd head</oddHeader><oddFooter>&amp;Codd foot</oddFooter><evenHeader>&amp;Ceven head</evenHeader><evenFooter>&amp;Ceven foot</evenFooter><firstHeader>&amp;Cfirst head</firstHeader><firstFooter>&amp;Cfirst foot</firstFooter></headerFooter>"#,
    );
    let ws = WorksheetXml::parse(xml.as_bytes()).unwrap();
    let hf = ws.header_footer.as_ref().unwrap();
    assert_eq!(hf.odd_header.as_deref(), Some("&Codd head"));
    assert_eq!(hf.odd_footer.as_deref(), Some("&Codd foot"));
    assert_eq!(hf.even_header.as_deref(), Some("&Ceven head"));
    assert_eq!(hf.even_footer.as_deref(), Some("&Ceven foot"));
    assert_eq!(hf.first_header.as_deref(), Some("&Cfirst head"));
    assert_eq!(hf.first_footer.as_deref(), Some("&Cfirst foot"));
    assert!(hf.different_odd_even);
    assert!(hf.different_first);

    let json: serde_json::Value = serde_json::from_str(&parse_json(&xml).unwrap()).unwrap();
    let hf = &json["headerFooter"];
    for (key, text) in [
        ("oddHeader", "&Codd head"),
        ("oddFooter", "&Codd foot"),
        ("evenHeader", "&Ceven head"),
        ("evenFooter", "&Ceven foot"),
        ("firstHeader", "&Cfirst head"),
        ("firstFooter", "&Cfirst foot"),
    ] {
        assert_eq!(hf[key], text, "{key}");
    }
    assert_eq!(hf["differentOddEven"], true);
    assert_eq!(hf["differentFirst"], true);
}

#[test]
fn test_header_footer_roundtrip() {
    let ws = WorksheetXml {
        header_footer: Some(HeaderFooter {
            odd_header: Some("&CPage &P of &N".into()),
            odd_footer: Some("&L&D&R&F".into()),
            ..Default::default()
        }),
        ..Default::default()
    };
    let xml = ws.to_xml_with_sst(None, &[], None, &[]).unwrap();
    let ws2 = WorksheetXml::parse(&xml).unwrap();
    let hf = ws2.header_footer.as_ref().unwrap();
    assert_eq!(hf.odd_header.as_deref(), Some("&CPage &P of &N"));
    assert_eq!(hf.odd_footer.as_deref(), Some("&L&D&R&F"));
}

#[test]
fn test_row_outline_level_roundtrip() {
    let ws = WorksheetXml {
        rows: vec![
            Row {
                index: 1,
                ..Default::default()
            },
            Row {
                index: 2,
                outline_level: Some(1),
                ..Default::default()
            },
            Row {
                index: 3,
                outline_level: Some(1),
                ..Default::default()
            },
            Row {
                index: 4,
                ..Default::default()
            },
        ],
        ..Default::default()
    };
    let xml = ws.to_xml_with_sst(None, &[], None, &[]).unwrap();
    let ws2 = WorksheetXml::parse(&xml).unwrap();
    assert_eq!(ws2.rows[1].outline_level, Some(1));
    assert_eq!(ws2.rows[2].outline_level, Some(1));
    assert!(ws2.rows[0].outline_level.is_none());
}

#[test]
fn test_collapsed_row_group_roundtrip() {
    let xml = r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="2" outlineLevel="1" hidden="1"><c r="A2"><v>100</v></c></row>
    <row r="3" outlineLevel="1" hidden="1"><c r="A3"><v>200</v></c></row>
    <row r="4" collapsed="1"><c r="A4"><v>300</v></c></row>
  </sheetData>
</worksheet>"#;
    let ws = WorksheetXml::parse(xml.as_bytes()).unwrap();
    assert_eq!(ws.rows[0].outline_level, Some(1));
    assert!(ws.rows[0].hidden);
    assert!(ws.rows[2].collapsed);

    let xml2 = ws.to_xml_with_sst(None, &[], None, &[]).unwrap();
    let ws2 = WorksheetXml::parse(&xml2).unwrap();
    assert_eq!(ws2.rows[0].outline_level, Some(1));
    assert!(ws2.rows[0].hidden);
    assert!(ws2.rows[2].collapsed);
}

#[test]
fn row_custom_height_roundtrips_with_both_attributes() {
    let xml = sheet(
        r#"<sheetData><row r="1" ht="40" customHeight="1"><c r="A1"><v>1</v></c></row></sheetData>"#,
    );
    let (ws, from_json, _) = parse_both(&xml);
    for model in [&ws, &from_json] {
        assert_eq!(model.rows[0].height, Some(40.0));
        assert!(model.rows[0].custom_height);
    }
    let out = String::from_utf8(ws.to_xml_with_sst(None, &[], None, &[]).unwrap()).unwrap();
    assert!(
        out.contains(r#"<row r="1" ht="40" customHeight="1">"#),
        "{out}"
    );
}

#[test]
fn row_auto_fit_height_is_written_back_without_custom_height() {
    let xml = sheet(r#"<sheetData><row r="1" ht="40"><c r="A1"><v>1</v></c></row></sheetData>"#);
    let (ws, from_json, json) = parse_both(&xml);
    for model in [&ws, &from_json] {
        assert_eq!(model.rows[0].height, Some(40.0));
        assert!(!model.rows[0].custom_height);
    }
    assert!(!json.contains("customHeight"), "{json}");
    let out = String::from_utf8(ws.to_xml_with_sst(None, &[], None, &[]).unwrap()).unwrap();
    assert!(out.contains(r#"<row r="1" ht="40">"#), "{out}");
    assert!(!out.contains("customHeight"), "{out}");
}

#[test]
fn row_custom_height_follows_xsd_boolean_and_orders_after_hidden() {
    let xml = sheet(
        r#"<sheetData><row r="1" ht="20" customHeight="true" hidden="1" outlineLevel="1"/><row r="2" ht="20" customHeight="bogus"/><row r="3" customHeight="1"/></sheetData>"#,
    );
    let (ws, from_json, _) = parse_both(&xml);
    for model in [&ws, &from_json] {
        assert!(model.rows[0].custom_height);
        assert!(
            !model.rows[1].custom_height,
            "bogus keeps the default false"
        );
        assert!(model.rows[2].custom_height);
    }
    let out = String::from_utf8(ws.to_xml_with_sst(None, &[], None, &[]).unwrap()).unwrap();
    assert!(
        out.contains(r#"<row r="1" ht="20" hidden="1" customHeight="1" outlineLevel="1"/>"#),
        "{out}"
    );
    assert!(
        out.contains(r#"<row r="3"/>"#),
        "a flag without a height writes nothing: {out}"
    );
}

#[test]
fn test_column_outline_level_roundtrip() {
    let ws = WorksheetXml {
        columns: vec![
            ColumnInfo {
                min: 1,
                max: 1,
                width: 10.0,
                hidden: false,
                custom_width: true,
                outline_level: None,
                collapsed: false,
                style_index: None,
                best_fit: false,
                phonetic: false,
            },
            ColumnInfo {
                min: 2,
                max: 3,
                width: 10.0,
                hidden: false,
                custom_width: true,
                outline_level: Some(1),
                collapsed: false,
                style_index: None,
                best_fit: false,
                phonetic: false,
            },
            ColumnInfo {
                min: 4,
                max: 4,
                width: 12.0,
                hidden: false,
                custom_width: true,
                outline_level: None,
                collapsed: false,
                style_index: None,
                best_fit: false,
                phonetic: false,
            },
        ],
        ..Default::default()
    };
    let xml = ws.to_xml_with_sst(None, &[], None, &[]).unwrap();
    let ws2 = WorksheetXml::parse(&xml).unwrap();
    assert!(ws2.columns[0].outline_level.is_none());
    assert_eq!(ws2.columns[1].outline_level, Some(1));
}

#[test]
fn test_outline_properties_roundtrip() {
    let ws = WorksheetXml {
        outline_properties: Some(OutlineProperties {
            summary_below: false,
            summary_right: true,
        }),
        ..Default::default()
    };
    let xml = ws.to_xml_with_sst(None, &[], None, &[]).unwrap();
    let ws2 = WorksheetXml::parse(&xml).unwrap();
    let op = ws2.outline_properties.as_ref().unwrap();
    assert!(!op.summary_below);
    assert!(op.summary_right);
}

#[test]
fn test_roundtrip_horizontal_split() {
    let ws = WorksheetXml {
        split_pane: Some(SplitPane {
            horizontal: Some(2400.0),
            vertical: None,
            top_left_cell: Some("A5".to_owned()),
            active_pane: Some("bottomLeft".to_owned()),
        }),
        pane_selections: vec![PaneSelection {
            pane: Some("bottomLeft".to_owned()),
            active_cell: Some("A5".to_owned()),
            sqref: Some("A5".to_owned()),
        }],
        ..Default::default()
    };
    let xml = ws.to_xml().unwrap();
    let ws2 = WorksheetXml::parse(&xml).unwrap();
    assert!(ws2.frozen_pane.is_none());
    let sp = ws2.split_pane.as_ref().unwrap();
    assert!((sp.horizontal.unwrap() - 2400.0).abs() < f64::EPSILON);
    assert!(sp.vertical.is_none());
    assert_eq!(sp.top_left_cell.as_deref(), Some("A5"));
    assert_eq!(ws2.pane_selections.len(), 1);
}

#[test]
fn test_roundtrip_four_way_split() {
    let ws = WorksheetXml {
        split_pane: Some(SplitPane {
            horizontal: Some(2400.0),
            vertical: Some(3000.0),
            top_left_cell: Some("D5".to_owned()),
            active_pane: Some("bottomRight".to_owned()),
        }),
        pane_selections: vec![
            PaneSelection {
                pane: Some("topRight".to_owned()),
                active_cell: Some("D1".to_owned()),
                sqref: Some("D1".to_owned()),
            },
            PaneSelection {
                pane: Some("bottomLeft".to_owned()),
                active_cell: Some("A5".to_owned()),
                sqref: Some("A5".to_owned()),
            },
            PaneSelection {
                pane: Some("bottomRight".to_owned()),
                active_cell: Some("D5".to_owned()),
                sqref: Some("D5".to_owned()),
            },
        ],
        ..Default::default()
    };
    let xml = ws.to_xml().unwrap();
    let ws2 = WorksheetXml::parse(&xml).unwrap();
    let sp = ws2.split_pane.as_ref().unwrap();
    assert!((sp.vertical.unwrap() - 3000.0).abs() < f64::EPSILON);
    assert!((sp.horizontal.unwrap() - 2400.0).abs() < f64::EPSILON);
    assert_eq!(ws2.pane_selections.len(), 3);
}

#[test]
fn test_split_pane_overrides_frozen_on_write() {
    let ws = WorksheetXml {
        frozen_pane: Some(FrozenPane { rows: 1, cols: 0 }),
        split_pane: Some(SplitPane {
            horizontal: Some(2400.0),
            vertical: None,
            top_left_cell: Some("A5".to_owned()),
            active_pane: Some("bottomLeft".to_owned()),
        }),
        ..Default::default()
    };
    let xml = ws.to_xml().unwrap();
    let xml_str = std::str::from_utf8(&xml).unwrap();
    assert!(xml_str.contains(r#"state="split""#));
    assert!(!xml_str.contains(r#"state="frozen""#));
}

#[test]
fn test_parse_sheet_view_hide_gridlines_zoom() {
    let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetViews>
    <sheetView showGridLines="0" zoomScale="150" workbookViewId="0"/>
  </sheetViews>
  <sheetData/>
</worksheet>"#;
    let ws = WorksheetXml::parse(xml.as_bytes()).unwrap();
    let sv = ws.sheet_view.as_ref().unwrap();
    assert!(!sv.show_grid_lines);
    assert_eq!(sv.zoom_scale, Some(150));
}

#[test]
fn test_parse_sheet_view_rtl() {
    let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetViews>
    <sheetView rightToLeft="1" showRowColHeaders="0" workbookViewId="0"/>
  </sheetViews>
  <sheetData/>
</worksheet>"#;
    let ws = WorksheetXml::parse(xml.as_bytes()).unwrap();
    let sv = ws.sheet_view.as_ref().unwrap();
    assert!(sv.right_to_left);
    assert!(!sv.show_row_col_headers);
}

#[test]
fn test_parse_sheet_view_page_break_preview() {
    let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetViews>
    <sheetView view="pageBreakPreview" zoomScaleNormal="100" zoomScale="60" workbookViewId="0"/>
  </sheetViews>
  <sheetData/>
</worksheet>"#;
    let ws = WorksheetXml::parse(xml.as_bytes()).unwrap();
    let sv = ws.sheet_view.as_ref().unwrap();
    assert_eq!(sv.view.as_deref(), Some("pageBreakPreview"));
    assert_eq!(sv.zoom_scale, Some(60));
    assert_eq!(sv.zoom_scale_normal, Some(100));
}

#[test]
fn test_parse_sheet_view_page_layout() {
    let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetViews>
    <sheetView view="pageLayout" showRuler="0" showWhiteSpace="0" zoomScalePageLayoutView="75" workbookViewId="0"/>
  </sheetViews>
  <sheetData/>
</worksheet>"#;
    let ws = WorksheetXml::parse(xml.as_bytes()).unwrap();
    let sv = ws.sheet_view.as_ref().unwrap();
    assert_eq!(sv.view.as_deref(), Some("pageLayout"));
    assert!(!sv.show_ruler);
    assert!(!sv.show_white_space);
    assert_eq!(sv.zoom_scale_page_layout_view, Some(75));
}

#[test]
fn test_parse_sheet_view_defaults_only() {
    let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetViews>
    <sheetView workbookViewId="0"/>
  </sheetViews>
  <sheetData/>
</worksheet>"#;
    let ws = WorksheetXml::parse(xml.as_bytes()).unwrap();
    // All defaults -> sheet_view should be None (no non-default attributes found)
    assert!(ws.sheet_view.is_none());
}

#[test]
fn test_parse_sheet_view_with_frozen_pane() {
    let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetViews>
    <sheetView showGridLines="0" zoomScale="120" workbookViewId="0">
      <pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>
    </sheetView>
  </sheetViews>
  <sheetData/>
</worksheet>"#;
    let ws = WorksheetXml::parse(xml.as_bytes()).unwrap();
    let sv = ws.sheet_view.as_ref().unwrap();
    assert!(!sv.show_grid_lines);
    assert_eq!(sv.zoom_scale, Some(120));
    assert!(ws.frozen_pane.is_some());
}

/// Helper to construct a default WorksheetXml for testing.
fn default_test_ws() -> WorksheetXml {
    WorksheetXml::default()
}

#[test]
fn test_roundtrip_sheet_view_hide_gridlines_zoom() {
    let mut ws = default_test_ws();
    ws.sheet_view = Some(SheetViewData {
        show_grid_lines: false,
        zoom_scale: Some(150),
        ..SheetViewData::default()
    });
    let xml = ws.to_xml().unwrap();
    let ws2 = WorksheetXml::parse(&xml).unwrap();
    let sv = ws2.sheet_view.as_ref().unwrap();
    assert!(!sv.show_grid_lines);
    assert_eq!(sv.zoom_scale, Some(150));
}

#[test]
fn test_roundtrip_sheet_view_rtl() {
    let mut ws = default_test_ws();
    ws.sheet_view = Some(SheetViewData {
        right_to_left: true,
        show_row_col_headers: false,
        ..SheetViewData::default()
    });
    let xml = ws.to_xml().unwrap();
    let ws2 = WorksheetXml::parse(&xml).unwrap();
    let sv = ws2.sheet_view.as_ref().unwrap();
    assert!(sv.right_to_left);
    assert!(!sv.show_row_col_headers);
}

#[test]
fn test_roundtrip_sheet_view_page_break_preview() {
    let mut ws = default_test_ws();
    ws.sheet_view = Some(SheetViewData {
        view: Some("pageBreakPreview".to_owned()),
        zoom_scale: Some(60),
        zoom_scale_normal: Some(100),
        ..SheetViewData::default()
    });
    let xml = ws.to_xml().unwrap();
    let ws2 = WorksheetXml::parse(&xml).unwrap();
    let sv = ws2.sheet_view.as_ref().unwrap();
    assert_eq!(sv.view.as_deref(), Some("pageBreakPreview"));
    assert_eq!(sv.zoom_scale, Some(60));
    assert_eq!(sv.zoom_scale_normal, Some(100));
}

#[test]
fn test_roundtrip_sheet_view_with_frozen_pane() {
    let mut ws = default_test_ws();
    ws.sheet_view = Some(SheetViewData {
        show_grid_lines: false,
        zoom_scale: Some(120),
        ..SheetViewData::default()
    });
    ws.frozen_pane = Some(FrozenPane { rows: 1, cols: 0 });
    let xml = ws.to_xml().unwrap();
    let ws2 = WorksheetXml::parse(&xml).unwrap();
    assert!(ws2.sheet_view.is_some());
    assert!(ws2.frozen_pane.is_some());
    assert!(!ws2.sheet_view.unwrap().show_grid_lines);
}

#[test]
fn test_sheet_view_defaults_minimal_xml() {
    let mut ws = default_test_ws();
    ws.sheet_view = Some(SheetViewData::default());
    let xml = ws.to_xml().unwrap();
    let xml_str = std::str::from_utf8(&xml).unwrap();
    // All defaults → no non-default attributes should appear
    assert!(!xml_str.contains("showGridLines"));
    assert!(!xml_str.contains("zoomScale"));
    assert!(!xml_str.contains("rightToLeft"));
    // But sheetViews should still be written
    assert!(xml_str.contains("sheetViews"));
    assert!(xml_str.contains("sheetView"));
}

// ---- Sparkline parser tests ----

#[test]
fn test_parse_line_sparkline() {
    let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData/>
  <extLst>
    <ext uri="{05C60535-1F16-4fd2-B633-F4F36F0B64E0}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">
      <x14:sparklineGroups xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">
        <x14:sparklineGroup>
          <x14:sparklines>
            <x14:sparkline>
              <xm:f>Sheet1!A1:A10</xm:f>
              <xm:sqref>B1</xm:sqref>
            </x14:sparkline>
            <x14:sparkline>
              <xm:f>Sheet1!A1:A10</xm:f>
              <xm:sqref>B2</xm:sqref>
            </x14:sparkline>
          </x14:sparklines>
        </x14:sparklineGroup>
      </x14:sparklineGroups>
    </ext>
  </extLst>
</worksheet>"#;
    let ws = WorksheetXml::parse(xml).unwrap();
    assert_eq!(ws.sparkline_groups.len(), 1);
    assert_eq!(ws.sparkline_groups[0].sparkline_type, "line");
    assert_eq!(ws.sparkline_groups[0].sparklines.len(), 2);
    assert_eq!(
        ws.sparkline_groups[0].sparklines[0].formula,
        "Sheet1!A1:A10"
    );
    assert_eq!(ws.sparkline_groups[0].sparklines[0].sqref, "B1");
}

#[test]
fn test_parse_column_sparkline_with_colors() {
    let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData/>
  <extLst>
    <ext uri="{05C60535-1F16-4fd2-B633-F4F36F0B64E0}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">
      <x14:sparklineGroups xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">
        <x14:sparklineGroup type="column" markers="1" high="1" low="1">
          <x14:colorSeries rgb="FF376092"/>
          <x14:colorNegative rgb="FFC00000"/>
          <x14:sparklines>
            <x14:sparkline>
              <xm:f>Sheet1!A1:A5</xm:f>
              <xm:sqref>C1</xm:sqref>
            </x14:sparkline>
          </x14:sparklines>
        </x14:sparklineGroup>
      </x14:sparklineGroups>
    </ext>
  </extLst>
</worksheet>"#;
    let ws = WorksheetXml::parse(xml).unwrap();
    assert_eq!(ws.sparkline_groups[0].sparkline_type, "column");
    assert!(ws.sparkline_groups[0].markers);
    assert!(ws.sparkline_groups[0].high);
    assert!(ws.sparkline_groups[0].low);
    assert_eq!(
        ws.sparkline_groups[0].color_series.as_deref(),
        Some("FF376092")
    );
    assert_eq!(
        ws.sparkline_groups[0].color_negative.as_deref(),
        Some("FFC00000")
    );
}

#[test]
fn test_parse_sparkline_with_options() {
    let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData/>
  <extLst>
    <ext uri="{05C60535-1F16-4fd2-B633-F4F36F0B64E0}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">
      <x14:sparklineGroups xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">
        <x14:sparklineGroup type="stacked" displayEmptyCellsAs="gap" lineWeight="1.25" displayXAxis="1" negative="1" first="1" last="1" manualMin="0" manualMax="100">
          <x14:sparklines>
            <x14:sparkline>
              <xm:f>Sheet1!D1:D10</xm:f>
              <xm:sqref>E1</xm:sqref>
            </x14:sparkline>
          </x14:sparklines>
        </x14:sparklineGroup>
      </x14:sparklineGroups>
    </ext>
  </extLst>
</worksheet>"#;
    let ws = WorksheetXml::parse(xml).unwrap();
    let g = &ws.sparkline_groups[0];
    assert_eq!(g.sparkline_type, "stacked");
    assert_eq!(g.display_empty_cells_as.as_deref(), Some("gap"));
    assert_eq!(g.line_weight, Some(1.25));
    assert!(g.display_x_axis);
    assert!(g.negative);
    assert!(g.first);
    assert!(g.last);
    assert_eq!(g.manual_min, Some(0.0));
    assert_eq!(g.manual_max, Some(100.0));
}

#[test]
fn test_parse_no_sparklines() {
    let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData/>
</worksheet>"#;
    let ws = WorksheetXml::parse(xml).unwrap();
    assert!(ws.sparkline_groups.is_empty());
}

// ---- Sparkline writer roundtrip tests ----

#[test]
fn test_sparkline_line_roundtrip() {
    let mut ws = default_test_ws();
    ws.sparkline_groups = vec![SparklineGroup {
        sparkline_type: "line".into(),
        sparklines: vec![
            Sparkline {
                formula: "Sheet1!A1:A10".into(),
                sqref: "B1".into(),
            },
            Sparkline {
                formula: "Sheet1!A1:A10".into(),
                sqref: "B2".into(),
            },
        ],
        ..Default::default()
    }];
    let xml = ws.to_xml_with_sst(None, &[], None, &[]).unwrap();
    assert!(
        std::str::from_utf8(&xml)
            .unwrap()
            .contains(r#"<ext uri="{05C60535-1F16-4fd2-B633-F4F36F0B64E0}""#)
    );
    let ws2 = WorksheetXml::parse(&xml).unwrap();
    assert_eq!(ws2.sparkline_groups.len(), 1);
    assert_eq!(ws2.sparkline_groups[0].sparklines.len(), 2);
    assert_eq!(
        ws2.sparkline_groups[0].sparklines[0].formula,
        "Sheet1!A1:A10"
    );
    assert_eq!(ws2.sparkline_groups[0].sparklines[0].sqref, "B1");
    assert_eq!(ws2.sparkline_groups[0].sparklines[1].sqref, "B2");
    assert_eq!(ws2.sparkline_groups[0].sparkline_type, "line");
}

#[test]
fn test_sparkline_column_with_colors_roundtrip() {
    let mut ws = default_test_ws();
    ws.sparkline_groups = vec![SparklineGroup {
        sparkline_type: "column".into(),
        sparklines: vec![Sparkline {
            formula: "Sheet1!C1:C5".into(),
            sqref: "D1".into(),
        }],
        color_series: Some("FF376092".into()),
        color_negative: Some("FFC00000".into()),
        markers: true,
        high: true,
        low: true,
        ..Default::default()
    }];
    let xml = ws.to_xml_with_sst(None, &[], None, &[]).unwrap();
    let ws2 = WorksheetXml::parse(&xml).unwrap();
    assert_eq!(ws2.sparkline_groups[0].sparkline_type, "column");
    assert_eq!(
        ws2.sparkline_groups[0].color_series.as_deref(),
        Some("FF376092")
    );
    assert_eq!(
        ws2.sparkline_groups[0].color_negative.as_deref(),
        Some("FFC00000")
    );
    assert!(ws2.sparkline_groups[0].markers);
    assert!(ws2.sparkline_groups[0].high);
    assert!(ws2.sparkline_groups[0].low);
}

#[test]
fn test_multiple_sparkline_groups_roundtrip() {
    let mut ws = default_test_ws();
    ws.sparkline_groups = vec![
        SparklineGroup {
            sparkline_type: "line".into(),
            sparklines: vec![Sparkline {
                formula: "Sheet1!A1:A10".into(),
                sqref: "B1".into(),
            }],
            ..Default::default()
        },
        SparklineGroup {
            sparkline_type: "column".into(),
            sparklines: vec![Sparkline {
                formula: "Sheet1!C1:C10".into(),
                sqref: "D1".into(),
            }],
            negative: true,
            display_empty_cells_as: Some("gap".into()),
            ..Default::default()
        },
    ];
    let xml = ws.to_xml_with_sst(None, &[], None, &[]).unwrap();
    let ws2 = WorksheetXml::parse(&xml).unwrap();
    assert_eq!(ws2.sparkline_groups.len(), 2);
    assert_eq!(ws2.sparkline_groups[0].sparkline_type, "line");
    assert_eq!(ws2.sparkline_groups[1].sparkline_type, "column");
    assert!(ws2.sparkline_groups[1].negative);
    assert_eq!(
        ws2.sparkline_groups[1].display_empty_cells_as.as_deref(),
        Some("gap")
    );
}

#[test]
fn test_sparkline_all_options_roundtrip() {
    let mut ws = default_test_ws();
    ws.sparkline_groups = vec![SparklineGroup {
        sparkline_type: "stacked".into(),
        sparklines: vec![Sparkline {
            formula: "Sheet1!D1:D10".into(),
            sqref: "E1".into(),
        }],
        color_series: Some("FF376092".into()),
        color_negative: Some("FFC00000".into()),
        color_axis: Some("FF000000".into()),
        color_markers: Some("FF0070C0".into()),
        color_first: Some("FF00B050".into()),
        color_last: Some("FFFFC000".into()),
        color_high: Some("FF00B0F0".into()),
        color_low: Some("FFFF0000".into()),
        line_weight: Some(1.25),
        markers: true,
        high: true,
        low: true,
        first: true,
        last: true,
        negative: true,
        display_x_axis: true,
        display_empty_cells_as: Some("gap".into()),
        manual_min: Some(0.0),
        manual_max: Some(100.0),
        right_to_left: true,
        ..Default::default()
    }];
    let xml = ws.to_xml_with_sst(None, &[], None, &[]).unwrap();
    let ws2 = WorksheetXml::parse(&xml).unwrap();
    let g = &ws2.sparkline_groups[0];
    assert_eq!(g.sparkline_type, "stacked");
    assert_eq!(g.color_series.as_deref(), Some("FF376092"));
    assert_eq!(g.color_negative.as_deref(), Some("FFC00000"));
    assert_eq!(g.color_axis.as_deref(), Some("FF000000"));
    assert_eq!(g.color_markers.as_deref(), Some("FF0070C0"));
    assert_eq!(g.color_first.as_deref(), Some("FF00B050"));
    assert_eq!(g.color_last.as_deref(), Some("FFFFC000"));
    assert_eq!(g.color_high.as_deref(), Some("FF00B0F0"));
    assert_eq!(g.color_low.as_deref(), Some("FFFF0000"));
    assert_eq!(g.line_weight, Some(1.25));
    assert!(g.markers);
    assert!(g.high);
    assert!(g.low);
    assert!(g.first);
    assert!(g.last);
    assert!(g.negative);
    assert!(g.display_x_axis);
    assert_eq!(g.display_empty_cells_as.as_deref(), Some("gap"));
    assert_eq!(g.manual_min, Some(0.0));
    assert_eq!(g.manual_max, Some(100.0));
    assert!(g.right_to_left);
}

#[test]
fn test_parse_data_table_formula() {
    let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="2">
      <c r="B2">
        <f t="dataTable" r1="B1" r2="A2"/>
        <v>42</v>
      </c>
    </row>
  </sheetData>
</worksheet>"#;
    let ws = WorksheetXml::parse(xml).unwrap();
    let cell = &ws.rows[0].cells[0];
    assert_eq!(cell.formula_type.as_deref(), Some("dataTable"));
    assert_eq!(cell.formula_r1.as_deref(), Some("B1"));
    assert_eq!(cell.formula_r2.as_deref(), Some("A2"));
}

#[test]
fn test_parse_2d_data_table() {
    let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="2">
      <c r="B2">
        <f t="dataTable" r1="B1" r2="A2" dt2D="1" dtr1="1" dtr2="1"/>
        <v>99</v>
      </c>
    </row>
  </sheetData>
</worksheet>"#;
    let ws = WorksheetXml::parse(xml).unwrap();
    let cell = &ws.rows[0].cells[0];
    assert_eq!(cell.formula_dt2d, Some(true));
    assert_eq!(cell.formula_dtr1, Some(true));
    assert_eq!(cell.formula_dtr2, Some(true));
}

#[test]
fn test_data_table_formula_roundtrip() {
    let mut ws = default_test_ws();
    ws.rows.push(Row {
        index: 2,
        cells: vec![Cell {
            reference: "B2".into(),
            cell_type: CellType::Number,
            value: Some("42".into()),
            formula: Some(String::new()),
            formula_type: Some("dataTable".into()),
            formula_r1: Some("B1".into()),
            formula_r2: Some("A2".into()),
            formula_dt2d: Some(true),
            ..Default::default()
        }],
        ..Default::default()
    });
    let xml = ws.to_xml_with_sst(None, &[], None, &[]).unwrap();
    let ws2 = WorksheetXml::parse(&xml).unwrap();
    let cell = &ws2.rows[0].cells[0];
    assert_eq!(cell.formula_type.as_deref(), Some("dataTable"));
    assert_eq!(cell.formula_r1.as_deref(), Some("B1"));
    assert_eq!(cell.formula_r2.as_deref(), Some("A2"));
    assert_eq!(cell.formula_dt2d, Some(true));
}

#[test]
fn test_normal_formula_no_data_table_attrs() {
    let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1">
        <f>SUM(B1:B10)</f>
        <v>55</v>
      </c>
    </row>
  </sheetData>
</worksheet>"#;
    let ws = WorksheetXml::parse(xml).unwrap();
    let cell = &ws.rows[0].cells[0];
    assert!(cell.formula_r1.is_none());
    assert!(cell.formula_r2.is_none());
    assert!(cell.formula_dt2d.is_none());
}

// ---- Page breaks tests ----

#[test]
fn test_parse_row_breaks() {
    let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData/>
  <rowBreaks count="2" manualBreakCount="2">
    <brk id="10" max="16383" man="1"/>
    <brk id="20" max="16383" man="1"/>
  </rowBreaks>
</worksheet>"#;
    let ws = WorksheetXml::parse(xml).unwrap();
    let pb = ws.page_breaks.as_ref().unwrap();
    assert_eq!(pb.row_breaks.len(), 2);
    assert_eq!(pb.row_breaks[0].id, 10);
    assert_eq!(pb.row_breaks[0].max, Some(16383));
    assert!(pb.row_breaks[0].man);
    assert_eq!(pb.row_breaks[1].id, 20);
    assert!(pb.col_breaks.is_empty());
}

#[test]
fn test_parse_col_breaks() {
    let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData/>
  <colBreaks count="1" manualBreakCount="1">
    <brk id="5" max="1048575" man="1"/>
  </colBreaks>
</worksheet>"#;
    let ws = WorksheetXml::parse(xml).unwrap();
    let pb = ws.page_breaks.as_ref().unwrap();
    assert!(pb.row_breaks.is_empty());
    assert_eq!(pb.col_breaks.len(), 1);
    assert_eq!(pb.col_breaks[0].id, 5);
    assert_eq!(pb.col_breaks[0].max, Some(1048575));
    assert!(pb.col_breaks[0].man);
}

#[test]
fn test_parse_both_row_and_col_breaks() {
    let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData/>
  <rowBreaks count="1" manualBreakCount="1">
    <brk id="10" max="16383" man="1"/>
  </rowBreaks>
  <colBreaks count="1" manualBreakCount="1">
    <brk id="3" max="1048575" man="1"/>
  </colBreaks>
</worksheet>"#;
    let ws = WorksheetXml::parse(xml).unwrap();
    let pb = ws.page_breaks.as_ref().unwrap();
    assert_eq!(pb.row_breaks.len(), 1);
    assert_eq!(pb.col_breaks.len(), 1);
    assert_eq!(pb.row_breaks[0].id, 10);
    assert_eq!(pb.col_breaks[0].id, 3);
}

#[test]
fn test_roundtrip_page_breaks() {
    let mut ws = default_test_ws();
    ws.page_breaks = Some(PageBreaks {
        row_breaks: vec![
            PageBreak {
                id: 10,
                min: None,
                max: Some(16383),
                man: true,
            },
            PageBreak {
                id: 25,
                min: Some(0),
                max: Some(16383),
                man: true,
            },
        ],
        col_breaks: vec![PageBreak {
            id: 5,
            min: None,
            max: Some(1048575),
            man: true,
        }],
    });

    let xml = ws.to_xml().unwrap();
    let xml_str = std::str::from_utf8(&xml).unwrap();
    assert!(xml_str.contains("<rowBreaks"));
    assert!(xml_str.contains("<colBreaks"));
    assert!(xml_str.contains("manualBreakCount=\"2\""));

    let ws2 = WorksheetXml::parse(&xml).unwrap();
    let pb = ws2.page_breaks.as_ref().unwrap();
    assert_eq!(pb.row_breaks.len(), 2);
    assert_eq!(pb.row_breaks[0].id, 10);
    assert_eq!(pb.row_breaks[0].max, Some(16383));
    assert!(pb.row_breaks[0].man);
    assert_eq!(pb.row_breaks[1].id, 25);
    assert_eq!(pb.row_breaks[1].min, Some(0));
    assert_eq!(pb.col_breaks.len(), 1);
    assert_eq!(pb.col_breaks[0].id, 5);
}

#[test]
fn test_no_page_breaks_by_default() {
    let ws = default_test_ws();
    let xml = ws.to_xml().unwrap();
    let xml_str = std::str::from_utf8(&xml).unwrap();
    assert!(!xml_str.contains("rowBreaks"));
    assert!(!xml_str.contains("colBreaks"));
}

// -----------------------------------------------------------------------
// Text, attribute and SST-index fidelity, alike on both parse paths
// -----------------------------------------------------------------------

/// Wrap `body` in a SpreadsheetML `<worksheet>` root.
fn sheet(body: &str) -> String {
    format!(
        r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">{body}</worksheet>"#
    )
}

/// The JSON `parse_to_json` writes for `xml` (no SST, comments or tables).
fn parse_json(xml: &str) -> crate::Result<String> {
    parse_json_in(xml, &SheetContext::default())
}

/// `parse_json` against `ctx`.
fn parse_json_in(xml: &str, ctx: &SheetContext<'_>) -> crate::Result<String> {
    let mut out = String::new();
    WorksheetXml::parse_to_json(xml.as_bytes(), None, &[], &[], ctx, &mut out)?;
    Ok(out)
}

/// Parse `xml` on the struct path and on the JSON path (deserialised with
/// serde, which rejects duplicate keys); returns both models and the raw JSON.
fn parse_both(xml: &str) -> (WorksheetXml, WorksheetXml, String) {
    parse_both_in(xml, &SheetContext::default())
}

/// `parse_both` against `ctx`.
fn parse_both_in(xml: &str, ctx: &SheetContext<'_>) -> (WorksheetXml, WorksheetXml, String) {
    let ws = WorksheetXml::parse_with_sst(xml.as_bytes(), None, ctx).unwrap();
    let json = parse_json_in(xml, ctx).unwrap();
    let from_json: WorksheetXml = serde_json::from_str(&json)
        .unwrap_or_else(|e| panic!("parse_to_json output does not deserialize: {e}\n{json}"));
    (ws, from_json, json)
}

#[test]
fn only_the_first_sheet_view_counts() {
    let xml = sheet(
        r#"<sheetViews><sheetView showGridLines="0" workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/><selection activeCell="A1" sqref="A1"/></sheetView><sheetView tabSelected="1" workbookViewId="1"><pane xSplit="3" state="frozen"/><selection activeCell="B2" sqref="B2"/></sheetView><sheetView zoomScale="50" workbookViewId="2"/></sheetViews><sheetData/>"#,
    );
    let (ws, from_json, _) = parse_both(&xml);
    for model in [&ws, &from_json] {
        assert_eq!(model.pane_selections.len(), 1);
        assert_eq!(model.pane_selections[0].active_cell.as_deref(), Some("A1"));
        let sv = model.sheet_view.as_ref().unwrap();
        assert!(!sv.show_grid_lines && !sv.tab_selected && sv.zoom_scale.is_none());
        assert_eq!(
            model.frozen_pane.as_ref().map(|p| (p.rows, p.cols)),
            Some((1, 0))
        );
    }
}

#[test]
fn hyperlink_attribute_entities_decode_once_and_escape_once() {
    let xml = sheet(
        r#"<sheetData/><hyperlinks><hyperlink ref="A1" location="Sheet1!A1" display="A &amp; B" tooltip="x&#10;y"/></hyperlinks>"#,
    );
    let (ws, from_json, _) = parse_both(&xml);
    for model in [&ws, &from_json] {
        let hl = &model.hyperlinks[0];
        assert_eq!(hl.display.as_deref(), Some("A & B"));
        assert_eq!(hl.tooltip.as_deref(), Some("x\ny"));
    }
    let out = String::from_utf8(ws.to_xml().unwrap()).unwrap();
    assert_eq!(out.matches(r#"display="A &amp; B""#).count(), 1, "{out}");
    assert_eq!(out.matches(r#"tooltip="x&#10;y""#).count(), 1, "{out}");
    assert!(!out.contains("&amp;amp;"), "{out}");
}

#[test]
fn inline_string_entities_keep_adjacent_spaces() {
    let xml = sheet(
        r#"<sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>Tom &amp; Jerry</t></is></c></row></sheetData>"#,
    );
    let (ws, from_json, _) = parse_both(&xml);
    for model in [&ws, &from_json] {
        let cell = &model.rows[0].cells[0];
        assert_eq!(cell.value.as_deref(), Some("Tom & Jerry"));
        assert_eq!(cell.inline_string.as_deref(), Some("Tom & Jerry"));
    }
}

#[test]
fn formula_text_entities_keep_adjacent_spaces() {
    let xml = sheet(
        r#"<sheetData><row r="1"><c r="B1" t="str"><f>IF(A1&lt;1,"a &amp; b",0)</f><v>a &amp; b</v></c></row></sheetData>"#,
    );
    let (ws, from_json, _) = parse_both(&xml);
    for model in [&ws, &from_json] {
        let cell = &model.rows[0].cells[0];
        assert_eq!(cell.formula.as_deref(), Some(r#"IF(A1<1,"a & b",0)"#));
        assert_eq!(cell.value.as_deref(), Some("a & b"));
    }
}

#[test]
fn preserved_leading_whitespace_survives() {
    let xml = sheet(
        r#"<sheetData><row r="1"><c r="A1" t="inlineStr"><is><t xml:space="preserve">  lead</t></is></c></row></sheetData>"#,
    );
    let (ws, from_json, _) = parse_both(&xml);
    for model in [&ws, &from_json] {
        assert_eq!(model.rows[0].cells[0].value.as_deref(), Some("  lead"));
    }
}

#[test]
fn cdata_cell_text_is_kept() {
    let xml = sheet(
        r#"<sheetData><row r="1"><c r="A1" t="str"><f><![CDATA[A2&"x"]]></f><v>1</v></c></row></sheetData>"#,
    );
    let (ws, from_json, _) = parse_both(&xml);
    for model in [&ws, &from_json] {
        assert_eq!(model.rows[0].cells[0].formula.as_deref(), Some(r#"A2&"x""#));
    }
}

#[test]
fn sparkline_formula_after_header_footer_keeps_entity_spaces() {
    let xml = sheet(
        r#"<sheetData/><headerFooter><oddHeader>x</oddHeader></headerFooter><extLst><ext uri="{05C60535-1F16-4fd2-B633-F4F36F0B64E0}"><x14:sparklineGroups><x14:sparklineGroup><x14:sparklines><x14:sparkline><xm:f>'A &amp; B'!A1:B1</xm:f><xm:sqref>C1</xm:sqref></x14:sparkline></x14:sparklines></x14:sparklineGroup></x14:sparklineGroups></ext></extLst>"#,
    );
    let (ws, from_json, _) = parse_both(&xml);
    for model in [&ws, &from_json] {
        assert_eq!(
            model.header_footer.as_ref().unwrap().odd_header.as_deref(),
            Some("x")
        );
        let spark = &model.sparkline_groups[0].sparklines[0];
        assert_eq!(spark.formula, "'A & B'!A1:B1");
        assert_eq!(spark.sqref, "C1");
    }
}

#[test]
fn rich_inline_string_concatenates_runs_and_skips_phonetic_text() {
    let xml = sheet(
        r#"<sheetData><row r="1"><c r="A1" t="inlineStr"><is><r><t xml:space="preserve">Hello </t></r><r><rPr><b/></rPr><t>World</t></r><rPh sb="0" eb="1"><t>x</t></rPh><phoneticPr fontId="1"/></is></c></row></sheetData>"#,
    );
    let (ws, from_json, json) = parse_both(&xml);
    for model in [&ws, &from_json] {
        let cell = &model.rows[0].cells[0];
        assert_eq!(cell.value.as_deref(), Some("Hello World"));
        assert_eq!(cell.inline_string.as_deref(), Some("Hello World"));
        assert!(cell.rich_text.is_none());
    }
    assert_eq!(json.matches(r#""value":"#).count(), 1, "{json}");
    assert_eq!(json.matches(r#""inlineString":"#).count(), 1, "{json}");
}

#[test]
fn xsd_booleans_follow_the_lexical_space_and_keep_defaults() {
    let xml = sheet(
        r#"<sheetPr><outlinePr summaryBelow="false" summaryRight="bogus"/></sheetPr><sheetViews><sheetView showGridLines="false" showZeros="bogus" workbookViewId="0"/></sheetViews><sheetData><row r="1" collapsed="true" hidden="bogus"/></sheetData><headerFooter scaleWithDoc="false" differentFirst="true" alignWithMargins="bogus"><oddHeader>x</oddHeader></headerFooter>"#,
    );
    let (ws, from_json, _) = parse_both(&xml);
    for model in [&ws, &from_json] {
        let op = model.outline_properties.as_ref().unwrap();
        assert!(!op.summary_below);
        assert!(op.summary_right, "bogus keeps the default true");
        let sv = model.sheet_view.as_ref().unwrap();
        assert!(!sv.show_grid_lines);
        assert!(sv.show_zeros, "bogus keeps the default true");
        assert!(model.rows[0].collapsed);
        assert!(!model.rows[0].hidden, "bogus keeps the default false");
        let hf = model.header_footer.as_ref().unwrap();
        assert!(!hf.scale_with_doc);
        assert!(hf.different_first);
        assert!(hf.align_with_margins, "bogus keeps the default true");
    }
}

#[test]
fn xsd_boolean_uppercase_is_not_a_boolean() {
    let xml = sheet(
        r#"<sheetData><row r="1" collapsed="TRUE"/></sheetData><sheetProtection sheet="TRUE" objects="true"/>"#,
    );
    let (ws, from_json, _) = parse_both(&xml);
    for model in [&ws, &from_json] {
        assert!(!model.rows[0].collapsed);
        let sp = model.sheet_protection.as_ref().unwrap();
        assert!(!sp.sheet);
        assert!(sp.objects);
    }
}

#[test]
fn default_sheet_view_values_do_not_create_a_sheet_view() {
    let xml = sheet(
        r#"<sheetViews><sheetView showGridLines="true" tabSelected="0" view="normal" workbookViewId="0"/></sheetViews><sheetData/>"#,
    );
    let (ws, from_json, _) = parse_both(&xml);
    assert!(ws.sheet_view.is_none());
    assert!(from_json.sheet_view.is_none());
}

#[test]
fn malformed_numeric_attribute_is_xml_parse() {
    for body in [
        r#"<sheetData><row r="x"><c r="A1"><v>1</v></c></row></sheetData>"#,
        r#"<sheetData><row r="x"/></sheetData>"#,
        r#"<sheetData><row r="1"><c r="A1" s="bold"><v>1</v></c></row></sheetData>"#,
        r#"<sheetViews><sheetView zoomScale="big" workbookViewId="0"/></sheetViews><sheetData/>"#,
    ] {
        let xml = sheet(body);
        assert_matches!(
            WorksheetXml::parse(xml.as_bytes()),
            Err(crate::ModernXlsxError::XmlParse(_)),
            "struct path accepted {body}"
        );
        assert_matches!(
            parse_json(&xml),
            Err(crate::ModernXlsxError::XmlParse(_)),
            "JSON path accepted {body}"
        );
    }
}

#[test]
fn padded_numeric_attribute_still_parses() {
    let xml = sheet(r#"<sheetData><row r=" 7 "><c r="A7" s=" 2 "><v>1</v></c></row></sheetData>"#);
    let (ws, from_json, _) = parse_both(&xml);
    for model in [&ws, &from_json] {
        assert_eq!(model.rows[0].index, 7);
        assert_eq!(model.rows[0].cells[0].style_index, Some(2));
    }
}

#[test]
fn bad_attribute_entity_is_xml_parse() {
    let xml = sheet(
        r#"<sheetData/><hyperlinks><hyperlink ref="A1" display="A &bogus; B"/></hyperlinks>"#,
    );
    assert_matches!(
        WorksheetXml::parse(xml.as_bytes()),
        Err(crate::ModernXlsxError::XmlParse(_))
    );
    assert_matches!(parse_json(&xml), Err(crate::ModernXlsxError::XmlParse(_)));
}

#[test]
fn shared_and_array_master_formulas_emit_each_key_once() {
    let xml = sheet(
        r#"<sheetData><row r="1"><c r="A1"><f t="shared" ref="A1:A2" si="0">B1+1</f><v>2</v></c><c r="C1"><f t="array" ref="C1:C3">SEQUENCE(3)</f><v>1</v></c></row><row r="2"><c r="A2"><f t="shared" si="0"/><v>3</v></c></row></sheetData>"#,
    );
    let (_, from_json, json) = parse_both(&xml);
    assert_eq!(json.matches(r#""formulaType":"#).count(), 3, "{json}");
    assert_eq!(json.matches(r#""formulaRef":"#).count(), 2, "{json}");
    let a1 = &from_json.rows[0].cells[0];
    assert_eq!(a1.formula_type.as_deref(), Some("shared"));
    assert_eq!(a1.formula_ref.as_deref(), Some("A1:A2"));
    assert_eq!(a1.shared_index, Some(0));
    let a2 = &from_json.rows[1].cells[0];
    assert_eq!(a2.formula_type.as_deref(), Some("shared"));
    assert_eq!(a2.shared_index, Some(0));
    assert_eq!(a2.formula, None);
}

#[test]
fn rich_runs_resolve_by_sst_index_on_both_paths() {
    use crate::ooxml::shared_strings::{RichTextRun, SharedStringTable};
    let bold = RichTextRun {
        text: "A".into(),
        bold: Some(true),
        ..RichTextRun::default()
    };
    let sst = SharedStringTable {
        strings: vec!["A".into(), "A".into()],
        rich_runs: vec![None, Some(vec![bold])],
    };
    // The second index is padded: only the SST index is trimmed.
    let xml = sheet(
        r#"<sheetData><row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v> 1 </v></c></row></sheetData>"#,
    );
    let ctx = SheetContext::default();
    let ws = WorksheetXml::parse_with_sst(xml.as_bytes(), Some(&sst), &ctx).unwrap();
    let mut json = String::new();
    WorksheetXml::parse_to_json(xml.as_bytes(), Some(&sst), &[], &[], &ctx, &mut json).unwrap();
    let from_json: WorksheetXml = serde_json::from_str(&json).unwrap();
    for model in [&ws, &from_json] {
        let cells = &model.rows[0].cells;
        assert_eq!(cells[0].value.as_deref(), Some("A"));
        assert!(cells[0].rich_text.is_none(), "plain entry got rich runs");
        assert_eq!(cells[1].value.as_deref(), Some("A"));
        let runs = cells[1]
            .rich_text
            .as_ref()
            .expect("rich entry lost its runs");
        assert_eq!(runs.len(), 1);
        assert_eq!(runs[0].bold, Some(true));
    }
}

// -----------------------------------------------------------------------
// Preserved <ext>, t="d" serials, cm/dtr/del1/del2, colour specs
// -----------------------------------------------------------------------

const X14_NS: &str = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main";
const XM_NS: &str = "http://schemas.microsoft.com/office/excel/2006/main";

/// An x14 conditional-formatting extension whose `x14:` and `xm:` prefixes
/// are bound on the `<worksheet>` root (see `sheet`).
const X14_CF_EXT: &str = r#"<ext uri="{78C0D931-6437-407d-A8EE-F0AAD7539E65}"><x14:conditionalFormattings><x14:conditionalFormatting><x14:cfRule type="dataBar" id="{1E5A0F2C-3B4D-4E6F-8A9B-0C1D2E3F4A5B}"><x14:dataBar minLength="0" maxLength="100"><x14:cfvo type="autoMin"/><x14:cfvo type="autoMax"/></x14:dataBar></x14:cfRule><xm:sqref>A1:A3</xm:sqref></x14:conditionalFormatting></x14:conditionalFormattings></ext>"#;

/// `X14_CF_EXT` as it is kept: the root-bound prefixes re-declared on `<ext>`.
fn x14_cf_ext_kept() -> String {
    X14_CF_EXT.replacen(
        "<ext ",
        &format!(r#"<ext xmlns:x14="{X14_NS}" xmlns:xm="{XM_NS}" "#),
        1,
    )
}

/// Assert that every element and attribute prefix in `xml` is bound where
/// it is used (Namespaces in XML 1.0, §5).
fn assert_namespace_well_formed(xml: &[u8]) {
    use quick_xml::NsReader;
    use quick_xml::events::Event;
    use quick_xml::name::ResolveResult;
    let mut r = NsReader::from_reader(xml);
    let mut buf = Vec::new();
    loop {
        let (res, ev) = r
            .read_resolved_event_into(&mut buf)
            .expect("well-formed XML");
        assert!(
            !matches!(res, ResolveResult::Unknown(_)),
            "unbound element prefix: {res:?}"
        );
        match ev {
            Event::Start(e) | Event::Empty(e) => {
                for a in e.attributes() {
                    let a = a.expect("well-formed attribute");
                    if a.key.as_namespace_binding().is_none() {
                        let (res, _) = r.resolver().resolve_attribute(a.key);
                        assert!(
                            !matches!(res, ResolveResult::Unknown(_)),
                            "unbound attribute prefix: {res:?}"
                        );
                    }
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }
}

/// A one-cell sheet whose A1 is `<c r="A1" t="d"><v>{v}</v></c>`.
fn iso_date_sheet(v: &str) -> String {
    sheet(&format!(
        r#"<sheetData><row r="1"><c r="A1" t="d"><v>{v}</v></c></row></sheetData>"#
    ))
}

fn written(ws: &WorksheetXml) -> String {
    String::from_utf8(ws.to_xml().unwrap()).unwrap()
}

#[test]
fn non_sparkline_extensions_are_kept_verbatim_with_their_namespaces() {
    let spark = r#"<ext uri="{05C60535-1F16-4fd2-B633-F4F36F0B64E0}"><x14:sparklineGroups><x14:sparklineGroup><x14:sparklines><x14:sparkline><xm:f>Sheet1!A1:B1</xm:f><xm:sqref>C1</xm:sqref></x14:sparkline></x14:sparklines></x14:sparklineGroup></x14:sparklineGroups></ext>"#;
    let bare = r#"<ext uri="{00000000-0000-0000-0000-00000000000A}"/>"#;
    let xml = sheet(&format!(
        "<sheetData/><extLst>{spark}\n  {X14_CF_EXT}{bare}</extLst>"
    ));
    let kept = vec![x14_cf_ext_kept(), bare.to_owned()];
    let (ws, from_json, json) = parse_both(&xml);
    for model in [&ws, &from_json] {
        assert_eq!(model.preserved_extensions, kept);
        assert_eq!(model.sparkline_groups.len(), 1);
    }
    assert_eq!(
        json.matches(r#""preservedExtensions":"#).count(),
        1,
        "{json}"
    );

    let out = ws.to_xml().unwrap();
    let text = String::from_utf8(out.clone()).unwrap();
    assert_eq!(
        text.matches("<x14:conditionalFormattings>").count(),
        1,
        "{text}"
    );
    let open = text.find("<extLst>").unwrap();
    let close = text.find("</extLst>").unwrap();
    let at = text.find(&kept[0]).expect("kept <ext> written verbatim");
    assert!(open < at && at < close, "{text}");
    assert_namespace_well_formed(&out);
    // A kept <ext> is already self-contained: reading it back changes nothing.
    assert_eq!(
        WorksheetXml::parse(&out).unwrap().preserved_extensions,
        kept
    );
}

#[test]
fn iso_date_cells_read_as_serials_in_the_workbook_date_system() {
    let xml = iso_date_sheet("2024-01-15T12:00:00");
    for (date_system, serial) in [
        (DateSystem::Date1900, "45306.5"),
        (DateSystem::Date1904, "43844.5"),
    ] {
        let ctx = SheetContext {
            date_system,
            rels: None,
        };
        let (ws, from_json, _) = parse_both_in(&xml, &ctx);
        for model in [&ws, &from_json] {
            let cell = &model.rows[0].cells[0];
            assert_eq!(cell.cell_type, CellType::Number, "{date_system:?}");
            assert_eq!(cell.value.as_deref(), Some(serial), "{date_system:?}");
        }
    }
    let out = written(&WorksheetXml::parse(xml.as_bytes()).unwrap());
    assert!(out.contains(r#"<c r="A1"><v>45306.5</v></c>"#), "{out}");
}

#[test]
fn iso_date_lexical_forms() {
    for (v, serial) in [
        ("2024-01-15", "45306"),
        ("2024-01-15T12:00:00Z", "45306.5"),
        ("2024-01-15T06:00:00.000+05:30", "45306.25"),
        (" 2024-01-15T18:00:00-08:00 ", "45306.75"),
        ("T18:00:00", "0.75"),
        ("1900-02-28", "59"),
        ("1900-03-01", "61"),
    ] {
        let (ws, from_json, _) = parse_both(&iso_date_sheet(v));
        for model in [&ws, &from_json] {
            assert_eq!(model.rows[0].cells[0].value.as_deref(), Some(serial), "{v}");
        }
    }
}

#[test]
fn unparseable_iso_date_cell_is_kept_as_a_string_cell() {
    for v in [
        "garbage",
        "2024-13-01",
        "2024-01-15T24:00:01",
        "10000-01-01",
    ] {
        let (ws, from_json, json) = parse_both(&iso_date_sheet(v));
        for model in [&ws, &from_json] {
            let cell = &model.rows[0].cells[0];
            assert_eq!(cell.cell_type, CellType::FormulaStr, "{v}");
            assert_eq!(cell.value.as_deref(), Some(v), "{v}");
        }
        assert_eq!(json.matches(r#""cellType":"#).count(), 1, "{json}");
        let out = written(&ws);
        assert!(
            out.contains(&format!(r#"<c r="A1" t="str"><v>{v}</v></c>"#)),
            "{out}"
        );
    }
}

#[test]
fn cell_metadata_is_read_from_c_and_written_there_only() {
    let xml = sheet(
        r#"<sheetData><row r="1"><c r="A1" cm="1"><f t="array" ref="A1:A3">SEQUENCE(3)</f><v>1</v></c><c r="B1" cm="1"/></row></sheetData>"#,
    );
    let (ws, from_json, json) = parse_both(&xml);
    for model in [&ws, &from_json] {
        for cell in &model.rows[0].cells {
            assert_eq!(cell.cell_metadata, Some(1), "{}", cell.reference);
            assert_eq!(cell.dynamic_array, Some(true), "{}", cell.reference);
        }
    }
    assert_eq!(json.matches(r#""cellMetadata":1"#).count(), 2, "{json}");
    assert_eq!(json.matches(r#""dynamicArray":true"#).count(), 2, "{json}");
    let out = written(&ws);
    assert!(
        out.contains(r#"<c r="A1" cm="1"><f t="array" ref="A1:A3">SEQUENCE(3)</f><v>1</v></c>"#),
        "{out}"
    );
    assert!(out.contains(r#"<c r="B1" cm="1"/>"#), "{out}");
}

#[test]
fn legacy_formula_cm_dtr1_dtr2_are_read_and_rewritten_per_schema() {
    let xml = sheet(
        r#"<sheetData><row r="2"><c r="B2"><f t="array" ref="B2:B3" cm="1">SORT(A1:A2)</f><v>1</v></c><c r="C2"><f t="dataTable" ref="C2:D3" dt2D="1" dtr1="1" dtr2="1" r1="A1" r2="B1"/><v>9</v></c></row></sheetData>"#,
    );
    let (ws, from_json, _) = parse_both(&xml);
    for model in [&ws, &from_json] {
        let [b2, c2] = &model.rows[0].cells[..] else {
            panic!("two cells expected");
        };
        assert_eq!(b2.dynamic_array, Some(true));
        assert_eq!(b2.cell_metadata, None);
        assert_eq!((c2.formula_dtr1, c2.formula_dtr2), (Some(true), Some(true)));
        assert_eq!((c2.formula_del1, c2.formula_del2), (None, None));
    }
    let out = written(&ws);
    assert!(
        out.contains(r#"<c r="B2" cm="1"><f t="array" ref="B2:B3">SORT(A1:A2)</f>"#),
        "{out}"
    );
    assert!(out.contains(r#"dt2D="1" del1="1" del2="1"/>"#), "{out}");
    assert!(!out.contains("dtr1") && !out.contains("dtr2"), "{out}");
}

#[test]
fn data_table_dtr_del1_del2_round_trip() {
    let xml = sheet(
        r#"<sheetData><row r="2"><c r="B2"><f t="dataTable" ref="B2:C3" dtr="1" del1="1" del2="1" r1="B1" r2="C1"/><v>1</v></c></row></sheetData>"#,
    );
    let (ws, from_json, _) = parse_both(&xml);
    for model in [&ws, &from_json] {
        let c = &model.rows[0].cells[0];
        assert_eq!(
            (c.formula_dtr, c.formula_del1, c.formula_del2),
            (Some(true), Some(true), Some(true))
        );
        // The 1.x names keep being emitted until 2.0.
        assert_eq!((c.formula_dtr1, c.formula_dtr2), (Some(true), Some(true)));
    }
    let out = written(&ws);
    assert!(
        out.contains(r#"<f t="dataTable" ref="B2:C3" r1="B1" r2="C1" dtr="1" del1="1" del2="1"/>"#),
        "{out}"
    );
    assert!(!out.contains("dtr1") && !out.contains("dtr2"), "{out}");
}

#[test]
fn json_with_only_the_1x_names_still_writes_schema_attributes() {
    let ws: WorksheetXml = serde_json::from_str(
        r#"{"rows":[{"index":2,"cells":[{"reference":"B2","cellType":"number","value":"1","formulaType":"dataTable","formulaRef":"B2:B3","formulaR1":"A1","formulaDtr1":true},{"reference":"C2","cellType":"number","value":"1","formula":"SORT(A1:A2)","formulaType":"array","formulaRef":"C2:C3","dynamicArray":true}]}]}"#,
    )
    .unwrap();
    let out = written(&ws);
    assert!(out.contains(r#"r1="A1" del1="1"/>"#), "{out}");
    assert!(
        out.contains(r#"<c r="C2" cm="1"><f t="array" ref="C2:C3">"#),
        "{out}"
    );
    for gone in ["dtr1", "del2", "dtr="] {
        assert!(!out.contains(gone), "{gone} in {out}");
    }
}

#[test]
fn theme_colours_of_data_bar_color_scale_and_tab_round_trip() {
    let xml = sheet(
        r#"<sheetPr><tabColor theme="4"/></sheetPr><sheetData/><conditionalFormatting sqref="A1:A3"><cfRule type="dataBar" priority="1"><dataBar><cfvo type="min"/><cfvo type="max"/><color theme="5" tint="-0.25"/></dataBar></cfRule><cfRule type="colorScale" priority="2"><colorScale><cfvo type="min"/><cfvo type="percentile" val="50"/><cfvo type="max"/><color rgb="FFF8696B"/><color theme="4"/><color rgb="FF63BE7B"/></colorScale></cfRule></conditionalFormatting>"#,
    );
    let (ws, from_json, _) = parse_both(&xml);
    for model in [&ws, &from_json] {
        assert_eq!(model.tab_color, None);
        assert_eq!(
            model.tab_color_spec,
            Some(Color {
                theme: Some(4),
                ..Color::default()
            })
        );
        let rules = &model.conditional_formatting[0].rules;
        let db = rules[0].data_bar.as_ref().unwrap();
        assert_eq!(db.color, "");
        assert_eq!(
            db.color_spec,
            Some(Color {
                theme: Some(5),
                tint: Some(-0.25),
                ..Color::default()
            })
        );
        let cs = rules[1].color_scale.as_ref().unwrap();
        assert_eq!(cs.colors, ["FFF8696B", "", "FF63BE7B"]);
        assert_eq!(cs.color_specs.len(), 3, "one spec per stop");
        assert_eq!(cs.color_specs[0].rgb.as_deref(), Some("FFF8696B"));
        assert_eq!(cs.color_specs[1].theme, Some(4));
    }
    let out = written(&ws);
    for want in [
        r#"<tabColor theme="4"/>"#,
        r#"<color theme="5" tint="-0.25"/></dataBar>"#,
        r#"<color rgb="FFF8696B"/><color theme="4"/><color rgb="FF63BE7B"/></colorScale>"#,
    ] {
        assert!(out.contains(want), "{want} missing from {out}");
    }
    assert!(!out.contains(r#"rgb="""#), "{out}");
}

#[test]
fn colours_without_rgb_or_spec_never_write_an_empty_rgb() {
    let ws: WorksheetXml = serde_json::from_str(
        r#"{"rows":[],"tabColor":"","conditionalFormatting":[{"sqref":"A1:A2","rules":[{"ruleType":"dataBar","priority":1,"dataBar":{"cfvos":[{"cfvoType":"min"},{"cfvoType":"max"}],"color":""}}]}]}"#,
    )
    .unwrap();
    let out = written(&ws);
    assert!(
        out.contains(r#"<color rgb="FF638EC6"/></dataBar>"#),
        "{out}"
    );
    assert!(!out.contains("tabColor"), "{out}");
    assert!(!out.contains(r#"rgb="""#), "{out}");
}

#[test]
fn a_changed_legacy_colour_wins_over_a_stale_spec() {
    let xml = sheet(r#"<sheetPr><tabColor theme="4"/></sheetPr><sheetData/>"#);
    let mut ws = WorksheetXml::parse(xml.as_bytes()).unwrap();
    ws.tab_color = Some("ff0000".into());
    let out = written(&ws);
    assert!(out.contains(r#"<tabColor rgb="FFFF0000"/>"#), "{out}");
}

// -----------------------------------------------------------------------
// External hyperlinks: read through r:id, written as r:id plus a
// TargetMode="External" relationship
// -----------------------------------------------------------------------

/// Sheet rels holding an external hyperlink (rId1), an internal one (rId2)
/// and a comments part (rId3).
fn link_rels() -> Relationships {
    let mut rels = Relationships::new();
    rels.add_external(
        crate::ooxml::relationships::REL_HYPERLINK,
        "https://e.com/?a=1&b=2",
    );
    rels.add(
        "rId2",
        crate::ooxml::relationships::REL_HYPERLINK,
        "other.xml",
    );
    rels.add(
        "rId3",
        crate::ooxml::relationships::REL_COMMENTS,
        "../comments1.xml",
    );
    rels
}

#[test]
fn hyperlink_r_id_resolves_to_its_external_target_on_both_paths() {
    let xml = format!(
        r#"<worksheet xmlns="{}" xmlns:r="{}"><sheetData/><hyperlinks><hyperlink ref="B2" r:id="rId1" display="Ex"/><hyperlink ref="B3" r:id="rId2"/><hyperlink ref="B4" r:id="rId3"/><hyperlink ref="B5" r:id="rId9"/><hyperlink ref="B6" location="Data!A1"/><hyperlink ref="B7" r:id="rId1" location="frag"/></hyperlinks></worksheet>"#,
        crate::ooxml::xml::SPREADSHEET_NS,
        crate::ooxml::xml::REL_NS,
    );
    let rels = link_rels();
    let ctx = SheetContext {
        rels: Some(&rels),
        ..SheetContext::default()
    };
    let (ws, from_json, _) = parse_both_in(&xml, &ctx);
    for model in [&ws, &from_json] {
        let links = &model.hyperlinks;
        assert_eq!(links[0].target.as_deref(), Some("https://e.com/?a=1&b=2"));
        assert_eq!(links[0].location, None);
        assert_eq!(links[0].display.as_deref(), Some("Ex"));
        // Internal, non-hyperlink and dangling rIds are ignored.
        for hl in &links[1..=3] {
            assert_eq!(hl.target, None, "{}", hl.cell_ref);
        }
        assert_eq!(links[4].location.as_deref(), Some("Data!A1"));
        assert_eq!(links[4].target, None);
        assert_eq!(links[5].target.as_deref(), Some("https://e.com/?a=1&b=2"));
        assert_eq!(links[5].location.as_deref(), Some("frag"));
    }
    // Without rels nothing resolves.
    let (ws, _, _) = parse_both(&xml);
    assert!(ws.hyperlinks.iter().all(|hl| hl.target.is_none()));
}

#[test]
fn external_hyperlinks_write_r_id_and_no_location() {
    let ws = WorksheetXml {
        hyperlinks: vec![
            Hyperlink {
                cell_ref: "A1".into(),
                target: Some("https://e.com/".into()),
                ..Default::default()
            },
            Hyperlink {
                cell_ref: "A2".into(),
                location: Some("Data!A1".into()),
                ..Default::default()
            },
            // 1.2.x shape: the URL sits in `location`.
            Hyperlink {
                cell_ref: "A3".into(),
                location: Some("mailto:hi@example.com".into()),
                ..Default::default()
            },
            Hyperlink {
                cell_ref: "A4".into(),
                target: Some("https://e.com/".into()),
                location: Some("frag".into()),
                ..Default::default()
            },
        ],
        ..Default::default()
    };
    let r_ids = [
        Some("rId1".to_owned()),
        None,
        Some("rId2".to_owned()),
        Some("rId3".to_owned()),
    ];
    let out = String::from_utf8(ws.to_xml_with_sst(None, &[], None, &r_ids).unwrap()).unwrap();
    assert!(
        out.contains(r#"<hyperlink ref="A1" r:id="rId1"/>"#),
        "{out}"
    );
    assert!(
        out.contains(r#"<hyperlink ref="A2" location="Data!A1"/>"#),
        "{out}"
    );
    assert!(
        out.contains(r#"<hyperlink ref="A3" r:id="rId2"/>"#),
        "{out}"
    );
    assert!(
        out.contains(r#"<hyperlink ref="A4" r:id="rId3" location="frag"/>"#),
        "{out}"
    );
    assert!(!out.contains("https:") && !out.contains("mailto:"), "{out}");
}

#[test]
fn uri_scheme_rule() {
    for s in [
        "https://e.com",
        "mailto:a@b.c",
        "file:///x",
        "C:\\a.xlsx",
        "tel:911",
        "a+b.c-d:x",
        // Near-misses of the A1 range shapes stay schemes.
        "ABCD1:B2",
        "A1:B",
        "B:D5",
        "A1:B2$",
    ] {
        assert!(has_uri_scheme(s), "{s}");
    }
    for s in [
        "A1:B2",
        "A1:$B$2",
        "B:D",
        "XFD1:$XFD$1048576",
        "a:$c",
        "Sheet2!A1",
        "'My Sheet'!C3",
        "#Sheet2!A1",
        "other.xlsx",
        ":x",
        "1a:b",
        "a b:c",
        "",
    ] {
        assert!(!has_uri_scheme(s), "{s}");
    }
}

#[test]
fn same_sheet_range_locations_stay_locations() {
    for range in ["A1:B2", "A1:$B$2", "B:D"] {
        let hl = Hyperlink {
            cell_ref: "C1".into(),
            location: Some(range.into()),
            ..Default::default()
        };
        assert_eq!(hl.external_target(), None, "{range}");
        assert_eq!(hl.written_location(), Some(range), "{range}");
    }
}

#[test]
fn hyperlink_id_in_a_foreign_namespace_is_not_r_id() {
    let xml = format!(
        r#"<worksheet xmlns="{}" xmlns:r="{}" xmlns:x="urn:example:other"><sheetData/><hyperlinks><hyperlink ref="B2" x:id="rId1"/><hyperlink ref="B3" r:id="rId1"/></hyperlinks></worksheet>"#,
        crate::ooxml::xml::SPREADSHEET_NS,
        crate::ooxml::xml::REL_NS,
    );
    let rels = link_rels();
    let ctx = SheetContext {
        rels: Some(&rels),
        ..SheetContext::default()
    };
    let (ws, from_json, _) = parse_both_in(&xml, &ctx);
    for model in [&ws, &from_json] {
        assert_eq!(model.hyperlinks[0].target, None);
        assert_eq!(
            model.hyperlinks[1].target.as_deref(),
            Some("https://e.com/?a=1&b=2")
        );
    }
}

#[test]
fn test_fit_to_page_parses_page_set_up_pr() {
    let ws = WorksheetXml::parse(
        br#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetPr><pageSetUpPr fitToPage="1"/></sheetPr><sheetData/><pageSetup fitToWidth="1" fitToHeight="0"/></worksheet>"#,
    )
    .unwrap();
    let ps = ws.page_setup.as_ref().unwrap();
    assert_eq!(ps.fit_to_page, Some(true));
    assert_eq!(ps.fit_to_width, Some(1));
    assert_eq!(ps.fit_to_height, Some(0));
}

#[test]
fn test_page_set_up_pr_without_fit_to_page_leaves_page_setup_none() {
    let ws = WorksheetXml::parse(
        br#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetPr><pageSetUpPr autoPageBreaks="0"/></sheetPr><sheetData/></worksheet>"#,
    )
    .unwrap();
    assert!(ws.page_setup.is_none());
}

#[test]
fn test_fit_to_page_writes_page_set_up_pr_after_outline_pr() {
    let ws = WorksheetXml {
        outline_properties: Some(OutlineProperties {
            summary_below: false,
            summary_right: true,
        }),
        page_setup: Some(PageSetup {
            fit_to_width: Some(1),
            fit_to_height: Some(0),
            fit_to_page: Some(true),
            ..Default::default()
        }),
        ..Default::default()
    };
    let xml = String::from_utf8(ws.to_xml().unwrap()).unwrap();
    assert!(
        xml.contains(
            r#"<sheetPr><outlinePr summaryBelow="0"/><pageSetUpPr fitToPage="1"/></sheetPr>"#
        ),
        "{xml}"
    );
    let back = WorksheetXml::parse(xml.as_bytes()).unwrap();
    assert_eq!(back.page_setup.unwrap().fit_to_page, Some(true));
}

#[test]
fn test_fit_to_page_none_writes_no_sheet_pr() {
    let ws = WorksheetXml {
        page_setup: Some(PageSetup {
            fit_to_width: Some(1),
            fit_to_page: None,
            ..Default::default()
        }),
        ..Default::default()
    };
    let xml = String::from_utf8(ws.to_xml().unwrap()).unwrap();
    assert!(
        !xml.contains("sheetPr") && !xml.contains("pageSetUpPr"),
        "{xml}"
    );
}

/// One sparkline group in an `<ext>` with the given `uri`.
fn sparkline_ext(uri: &str) -> String {
    format!(
        r#"<ext uri="{uri}"><x14:sparklineGroups><x14:sparklineGroup><x14:sparklines><x14:sparkline><xm:f>Sheet1!A1:B1</xm:f><xm:sqref>C1</xm:sqref></x14:sparkline></x14:sparklines></x14:sparklineGroup></x14:sparklineGroups></ext>"#
    )
}

#[test]
fn sparklines_are_written_under_the_uri_excel_recognises() {
    // Excel's x14 sparklineGroups extension URI (MS-XLSX 2.2.4.12 Worksheet, XlsxWriter,
    // rust_xlsxwriter); 1.2.1 wrote ...F4F36011B0BD, which Excel skips.
    const EXCEL_URI: &str = "{05C60535-1F16-4fd2-B633-F4F36F0B64E0}";
    const LEGACY_URI: &str = "{05C60535-1F16-4fd2-B633-F4F36011B0BD}";
    for uri in [EXCEL_URI, LEGACY_URI, &EXCEL_URI.to_ascii_lowercase()] {
        let xml = sheet(&format!(
            "<sheetData/><extLst>{}</extLst>",
            sparkline_ext(uri)
        ));
        let (ws, from_json, _) = parse_both(&xml);
        for model in [&ws, &from_json] {
            assert_eq!(model.sparkline_groups.len(), 1, "{uri}");
            assert!(model.preserved_extensions.is_empty(), "{uri}");
        }
        let out = String::from_utf8(ws.to_xml().unwrap()).unwrap();
        assert_eq!(
            out.matches(&format!(r#"uri="{EXCEL_URI}""#)).count(),
            1,
            "{out}"
        );
        assert!(!out.contains(LEGACY_URI), "{out}");
        let back = WorksheetXml::parse(out.as_bytes()).unwrap();
        assert_eq!(back.sparkline_groups, ws.sparkline_groups);
    }
}

#[test]
fn an_ext_that_only_shares_the_sparkline_uri_prefix_is_kept_verbatim() {
    let other = sparkline_ext("{05C60535-0000-0000-0000-000000000000}");
    let xml = sheet(&format!("<sheetData/><extLst>{other}</extLst>"));
    let (ws, from_json, _) = parse_both(&xml);
    for model in [&ws, &from_json] {
        assert!(model.sparkline_groups.is_empty());
        assert_eq!(model.preserved_extensions.len(), 1);
    }
}

#[test]
fn selection_on_a_default_sheet_view_round_trips() {
    let xml = sheet(
        r#"<sheetViews><sheetView workbookViewId="0"><selection activeCell="C5" sqref="C5"/></sheetView></sheetViews><sheetData/>"#,
    );
    let (ws, from_json, _) = parse_both(&xml);
    for model in [&ws, &from_json] {
        assert!(model.sheet_view.is_none());
        assert_eq!(model.pane_selections.len(), 1);
    }
    let out = String::from_utf8(ws.to_xml().unwrap()).unwrap();
    assert!(
        out.contains(
            r#"<sheetViews><sheetView workbookViewId="0"><selection activeCell="C5" sqref="C5"/></sheetView></sheetViews>"#
        ),
        "{out}"
    );
    let back = WorksheetXml::parse(out.as_bytes()).unwrap();
    assert_eq!(back.pane_selections, ws.pane_selections);

    // A sheet with no view data writes Excel's default view.
    let out = String::from_utf8(WorksheetXml::default().to_xml().unwrap()).unwrap();
    assert!(out.contains(DEFAULT_SHEET_VIEWS), "{out}");
}

#[test]
fn sparkline_group_doubles_are_written_as_xsd_doubles() {
    let ws: WorksheetXml = serde_json::from_str(
        r#"{"rows":[],"sparklineGroups":[{"sparklines":[{"formula":"Sheet1!A1:B1","sqref":"C1"}],"lineWeight":0.75,"manualMin":-2.0}]}"#,
    )
    .unwrap();
    let out = String::from_utf8(ws.to_xml().unwrap()).unwrap();
    assert!(out.contains(r#"lineWeight="0.75""#), "{out}");
    assert!(out.contains(r#"manualMin="-2""#), "{out}");

    // "INF" parses as an f64, but no xsd:double is written for it.
    let xml = sheet(&format!(
        "<sheetData/><extLst>{}</extLst>",
        sparkline_ext("{05C60535-1F16-4fd2-B633-F4F36F0B64E0}").replace(
            "<x14:sparklineGroup>",
            r#"<x14:sparklineGroup manualMax="INF">"#
        )
    ));
    let ws = WorksheetXml::parse(xml.as_bytes()).unwrap();
    assert_eq!(ws.sparkline_groups[0].manual_max, Some(f64::INFINITY));
    let err = ws.to_xml().unwrap_err();
    assert_eq!(err.code(), "XML_WRITE");
    assert!(
        err.to_string().contains("sparklineGroup manualMax"),
        "{err}"
    );
}

/// `parse_both` on input that ends inside `<sheetData>`: the JSON is valid,
/// both read paths agree, and only the finished rows are kept.
fn parse_truncated(xml: &str) -> (WorksheetXml, String) {
    let (ws, from_json, json) = parse_both(xml);
    serde_json::from_str::<serde_json::Value>(&json)
        .unwrap_or_else(|e| panic!("invalid JSON for {xml}: {e}\n{json}"));
    assert_eq!(
        serde_json::to_value(&ws).unwrap(),
        serde_json::to_value(&from_json).unwrap(),
        "{xml}"
    );
    (ws, json)
}

#[test]
fn input_ending_inside_a_cell_drops_the_open_row() {
    let (ws, json) = parse_truncated(r#"<worksheet><sheetData><row r="1"><c r="A1"><v>1</v>"#);
    assert!(ws.rows.is_empty());
    assert_eq!(json, r#"{"rows":[]}"#);
}

#[test]
fn input_ending_inside_a_row_drops_the_open_row() {
    let (ws, json) = parse_truncated(r#"<worksheet><sheetData><row r="1">"#);
    assert!(ws.rows.is_empty());
    assert_eq!(json, r#"{"rows":[]}"#);
}

#[test]
fn input_ending_inside_a_later_row_keeps_the_finished_rows() {
    let complete =
        r#"<worksheet><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData></worksheet>"#;
    let (ws, json) = parse_truncated(
        r#"<worksheet><sheetData><row r="1"><c r="A1"><v>1</v></c></row><row r="2"><c r="A2" t="str"><f>A1</f><v>2</v>"#,
    );
    assert_eq!(ws.rows.len(), 1);
    assert_eq!(ws.rows[0].index, 1);
    assert_eq!(ws.rows[0].cells[0].value.as_deref(), Some("1"));
    // The JSON equals that of the complete part holding only row 1.
    assert_eq!(json, parse_json(complete).unwrap());
}

#[test]
fn input_ending_inside_a_skipped_root_child_keeps_the_rows() {
    let (ws, json) = parse_truncated(
        r#"<worksheet><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData><ignoredErrors>"#,
    );
    assert_eq!(ws.rows.len(), 1);
    assert_eq!(
        json,
        parse_json(
            r#"<worksheet><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData></worksheet>"#
        )
        .unwrap()
    );
}

// -----------------------------------------------------------------------
// Hostile namespace counts: prefix lookups are hashed, not scanned
// -----------------------------------------------------------------------

#[test]
fn many_rel_ns_prefixes_and_hyperlinks_parse_in_linear_time() {
    const N: usize = 50_000;
    use std::fmt::Write as _;
    let rel_ns = crate::ooxml::xml::REL_NS;
    let mut decls = String::new();
    let mut links = String::new();
    for i in 0..N {
        write!(decls, r#" xmlns:p{i}="{rel_ns}""#).unwrap();
        write!(
            links,
            r#"<hyperlink ref="A{}" p{}:id="rId1"/>"#,
            i + 1,
            N - 1
        )
        .unwrap();
    }
    let xml = format!(
        r#"<worksheet xmlns="{}"{decls}><sheetData/><hyperlinks>{links}</hyperlinks></worksheet>"#,
        crate::ooxml::xml::SPREADSHEET_NS,
    );
    let rels = link_rels();
    let ctx = SheetContext {
        rels: Some(&rels),
        ..SheetContext::default()
    };
    let start = std::time::Instant::now();
    let (ws, from_json, _) = parse_both_in(&xml, &ctx);
    let elapsed = start.elapsed();
    for model in [&ws, &from_json] {
        assert_eq!(model.hyperlinks.len(), N);
        assert!(
            model
                .hyperlinks
                .iter()
                .all(|hl| hl.target.as_deref() == Some("https://e.com/?a=1&b=2"))
        );
    }
    assert!(
        elapsed < std::time::Duration::from_secs(5),
        "{N} declarations x {N} hyperlinks took {elapsed:?}"
    );
}

// -----------------------------------------------------------------------
// Attributes that decide what a rule, row, validation or sparkline means
// -----------------------------------------------------------------------

#[test]
fn cf_rule_meaning_attributes_round_trip_on_both_paths() {
    let xml = sheet(
        r#"<sheetData/><conditionalFormatting sqref="A1:A9"><cfRule type="aboveAverage" dxfId="0" priority="1" aboveAverage="0" stdDev="1"/><cfRule type="top10" dxfId="1" priority="2" bottom="1" percent="1" rank="5"/><cfRule type="containsText" dxfId="2" priority="3" operator="containsText" text="abc"><formula>NOT(ISERROR(SEARCH("abc",A1)))</formula></cfRule><cfRule type="timePeriod" dxfId="3" priority="4" timePeriod="lastWeek"><formula>A1&lt;TODAY()</formula></cfRule><cfRule type="expression" dxfId="4" priority="5" stopIfTrue="1"><formula>A1&gt;1</formula></cfRule><cfRule type="aboveAverage" dxfId="5" priority="6" equalAverage="1"/><cfRule type="cellIs" dxfId="6" priority="7" operator="equal"><formula>1</formula></cfRule></conditionalFormatting>"#,
    );
    let (ws, from_json, _) = parse_both(&xml);
    for model in [&ws, &from_json] {
        let r = &model.conditional_formatting[0].rules;
        assert_eq!(r[0].above_average, Some(false));
        assert_eq!(r[0].std_dev, Some(1));
        assert_eq!(
            (r[1].bottom, r[1].percent, r[1].rank),
            (Some(true), Some(true), Some(5))
        );
        assert_eq!(r[2].text.as_deref(), Some("abc"));
        assert_eq!(r[3].time_period.as_deref(), Some("lastWeek"));
        assert_eq!(r[4].stop_if_true, Some(true));
        assert_eq!(r[5].equal_average, Some(true));
        let plain = &r[6];
        assert_eq!(
            (
                plain.stop_if_true,
                plain.above_average,
                plain.percent,
                plain.bottom
            ),
            (None, None, None, None)
        );
        assert_eq!(
            (plain.equal_average, plain.rank, plain.std_dev),
            (None, None, None)
        );
        assert_eq!((&plain.text, &plain.time_period), (&None, &None));
    }
    let out = String::from_utf8(ws.to_xml().unwrap()).unwrap();
    for tag in [
        r#"<cfRule type="aboveAverage" dxfId="0" priority="1" aboveAverage="0" stdDev="1"/>"#,
        r#"<cfRule type="top10" dxfId="1" priority="2" percent="1" bottom="1" rank="5"/>"#,
        r#"<cfRule type="containsText" dxfId="2" priority="3" operator="containsText" text="abc">"#,
        r#"<cfRule type="timePeriod" dxfId="3" priority="4" timePeriod="lastWeek">"#,
        r#"<cfRule type="expression" dxfId="4" priority="5" stopIfTrue="1">"#,
        r#"<cfRule type="aboveAverage" dxfId="5" priority="6" equalAverage="1"/>"#,
        // A rule without them is written as before.
        r#"<cfRule type="cellIs" dxfId="6" priority="7" operator="equal"><formula>1</formula></cfRule>"#,
    ] {
        assert!(out.contains(tag), "{tag} missing from {out}");
    }
    let back = WorksheetXml::parse(out.as_bytes()).unwrap();
    assert_eq!(
        serde_json::to_value(&back.conditional_formatting).unwrap(),
        serde_json::to_value(&ws.conditional_formatting).unwrap()
    );
}

#[test]
fn row_style_is_read_only_with_custom_format_and_written_back() {
    let xml = sheet(
        r#"<sheetData><row r="3" s="5" customFormat="1"/><row r="4" s="5"/><row r="5" s="0" customFormat="true"><c r="A5"><v>1</v></c></row><row r="6" s="2" customFormat="0"/></sheetData>"#,
    );
    let (ws, from_json, json) = parse_both(&xml);
    for model in [&ws, &from_json] {
        let styles: Vec<_> = model.rows.iter().map(|r| r.style_index).collect();
        assert_eq!(styles, [Some(5), None, Some(0), None]);
    }
    assert!(
        json.contains(r#"{"index":3,"cells":[],"styleIndex":5}"#),
        "{json}"
    );
    let out = String::from_utf8(ws.to_xml().unwrap()).unwrap();
    for tag in [
        r#"<row r="3" s="5" customFormat="1"/>"#,
        r#"<row r="4"/>"#,
        r#"<row r="5" s="0" customFormat="1"><c r="A5">"#,
        r#"<row r="6"/>"#,
    ] {
        assert!(out.contains(tag), "{tag} missing from {out}");
    }
}

#[test]
fn data_validation_error_style_drop_down_and_ime_mode_round_trip() {
    let xml = sheet(
        r#"<sheetData/><dataValidations count="3"><dataValidation type="whole" errorStyle="warning" imeMode="off" operator="greaterThan" showErrorMessage="1" sqref="A1"><formula1>1</formula1></dataValidation><dataValidation type="list" allowBlank="1" showDropDown="1" sqref="B1"><formula1>"a,b"</formula1></dataValidation><dataValidation type="list" sqref="C1"><formula1>"c"</formula1></dataValidation></dataValidations>"#,
    );
    let (ws, from_json, _) = parse_both(&xml);
    for model in [&ws, &from_json] {
        let dv = &model.data_validations;
        assert_eq!(dv[0].error_style.as_deref(), Some("warning"));
        assert_eq!(dv[0].ime_mode.as_deref(), Some("off"));
        assert_eq!(dv[0].show_drop_down, None);
        assert_eq!(dv[1].show_drop_down, Some(true));
        assert_eq!((&dv[2].error_style, &dv[2].ime_mode), (&None, &None));
        assert_eq!(dv[2].show_drop_down, None);
    }
    let out = String::from_utf8(ws.to_xml().unwrap()).unwrap();
    for tag in [
        r#"<dataValidation type="whole" operator="greaterThan" errorStyle="warning" imeMode="off" showErrorMessage="1" sqref="A1">"#,
        r#"<dataValidation type="list" allowBlank="1" showDropDown="1" sqref="B1">"#,
        // A validation without them is written as before.
        r#"<dataValidation type="list" sqref="C1"><formula1>&quot;c&quot;</formula1></dataValidation>"#,
    ] {
        assert!(out.contains(tag), "{tag} missing from {out}");
    }
    let back = WorksheetXml::parse(out.as_bytes()).unwrap();
    assert_eq!(
        serde_json::to_value(&back.data_validations).unwrap(),
        serde_json::to_value(&ws.data_validations).unwrap()
    );
}

/// A sparkline group with custom axes, a date axis over `Sheet1!A20:J20`,
/// hidden-cell plotting and theme colours, then a group on a shared axis.
const SPARKLINE_AXES_EXT: &str = r#"<ext uri="{05C60535-1F16-4fd2-B633-F4F36F0B64E0}"><x14:sparklineGroups><x14:sparklineGroup manualMax="10" manualMin="0" dateAxis="1" displayHidden="1" minAxisType="custom" maxAxisType="custom"><x14:colorSeries theme="4" tint="-0.499984740745262"/><x14:colorNegative rgb="FFC00000"/><x14:colorMarkers theme="5"/><xm:f>Sheet1!A20:J20</xm:f><x14:sparklines><x14:sparkline><xm:f>Sheet1!A1:J1</xm:f><xm:sqref>K1</xm:sqref></x14:sparkline></x14:sparklines></x14:sparklineGroup><x14:sparklineGroup type="column" minAxisType="group" maxAxisType="group"><x14:sparklines><x14:sparkline><xm:f>Sheet1!A2:J2</xm:f><xm:sqref>K2</xm:sqref></x14:sparkline></x14:sparklines></x14:sparklineGroup></x14:sparklineGroups></ext>"#;

#[test]
fn sparkline_axis_types_date_axis_and_theme_colours_round_trip() {
    let xml = sheet(&format!(
        "<sheetData/><extLst>{SPARKLINE_AXES_EXT}</extLst>"
    ));
    let (ws, from_json, _) = parse_both(&xml);
    for model in [&ws, &from_json] {
        let g = &model.sparkline_groups[0];
        assert_eq!(g.min_axis_type.as_deref(), Some("custom"));
        assert_eq!(g.max_axis_type.as_deref(), Some("custom"));
        assert!(g.date_axis);
        assert!(g.display_hidden);
        assert_eq!(g.date_range.as_deref(), Some("Sheet1!A20:J20"));
        assert_eq!(g.sparklines[0].formula, "Sheet1!A1:J1");
        assert_eq!(g.color_series, None);
        assert_eq!(
            g.color_series_spec,
            Some(Color {
                theme: Some(4),
                tint: Some(-0.499984740745262),
                ..Color::default()
            })
        );
        assert_eq!(g.color_negative.as_deref(), Some("FFC00000"));
        assert_eq!(g.color_negative_spec, None);
        assert_eq!(g.color_markers_spec.as_ref().and_then(|c| c.theme), Some(5));
        let shared = &model.sparkline_groups[1];
        assert_eq!(shared.min_axis_type.as_deref(), Some("group"));
        assert_eq!(shared.max_axis_type.as_deref(), Some("group"));
        assert!(!shared.date_axis && !shared.display_hidden);
        assert_eq!(shared.date_range, None);
    }
    let out = String::from_utf8(ws.to_xml().unwrap()).unwrap();
    for part in [
        r#"dateAxis="1""#,
        r#"displayHidden="1""#,
        r#"minAxisType="custom""#,
        r#"maxAxisType="custom""#,
        r#"minAxisType="group""#,
        r#"<x14:colorSeries theme="4" tint="-0.499984740745262"/><x14:colorNegative rgb="FFC00000"/><x14:colorMarkers theme="5"/><xm:f>Sheet1!A20:J20</xm:f><x14:sparklines>"#,
    ] {
        assert!(out.contains(part), "{part} missing from {out}");
    }
    let back = WorksheetXml::parse(out.as_bytes()).unwrap();
    assert_eq!(back.sparkline_groups, ws.sparkline_groups);
}

#[test]
fn sparkline_manual_bounds_without_an_axis_type_are_written_as_custom() {
    let group = |manual_min, manual_max| SparklineGroup {
        sparklines: vec![Sparkline {
            formula: "Sheet1!A1:B1".into(),
            sqref: "C1".into(),
        }],
        manual_min,
        manual_max,
        ..Default::default()
    };
    let written = |g: SparklineGroup| {
        let ws = WorksheetXml {
            sparkline_groups: vec![g],
            ..Default::default()
        };
        String::from_utf8(ws.to_xml().unwrap()).unwrap()
    };
    let out = written(group(Some(-1.0), Some(5.0)));
    assert!(
        out.contains(r#"manualMin="-1" manualMax="5" minAxisType="custom" maxAxisType="custom""#),
        "{out}"
    );
    let out = written(group(None, Some(5.0)));
    assert!(out.contains(r#"maxAxisType="custom""#), "{out}");
    assert!(!out.contains("minAxisType"), "{out}");
    // No bounds, no axis types: written as before.
    let out = written(group(None, None));
    assert!(
        out.contains("<x14:sparklineGroup><x14:sparklines>"),
        "{out}"
    );
    // An explicit axis type wins over the implied one.
    let mut g = group(Some(0.0), None);
    g.min_axis_type = Some("group".into());
    let out = written(g);
    assert!(out.contains(r#"minAxisType="group""#), "{out}");
    assert!(!out.contains(r#"minAxisType="custom""#), "{out}");
}

#[test]
fn literal_line_ends_are_normalized_and_char_refs_kept() {
    let xml = sheet(
        "<sheetData><row r=\"1\">\
         <c r=\"A1\" t=\"inlineStr\"><is><t>a\r\nb\rc</t></is></c>\
         <c r=\"B1\" t=\"inlineStr\"><is><t>a&#13;&#10;b</t></is></c>\
         <c r=\"C1\" t=\"str\"><f>A1&amp;\"\r\n\"</f><v>x</v></c>\
         </row></sheetData>",
    );
    let (ws, from_json, _) = parse_both(&xml);
    for model in [&ws, &from_json] {
        let cells = &model.rows[0].cells;
        assert_eq!(cells[0].value.as_deref(), Some("a\nb\nc"));
        assert_eq!(cells[1].value.as_deref(), Some("a\r\nb"));
        assert_eq!(cells[2].formula.as_deref(), Some("A1&\"\n\""));
    }
}

#[test]
fn custom_sheet_view_children_stay_out_of_the_sheet_model() {
    let xml = sheet(
        r#"<sheetData/><autoFilter ref="A1:D50"/><customSheetViews><customSheetView guid="{8E5C1D2A-3B4F-4C6D-9E7F-0A1B2C3D4E5F}"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/><rowBreaks count="1" manualBreakCount="1"><brk id="9" max="16383" man="1"/></rowBreaks><pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/><pageSetup orientation="landscape"/><headerFooter><oddHeader>V</oddHeader></headerFooter><autoFilter ref="A1:B2"/><extLst><ext uri="{11111111-2222-3333-4444-555555555555}"><x14:view/></ext></extLst></customSheetView></customSheetViews><rowBreaks count="1" manualBreakCount="1"><brk id="5" max="16383" man="1"/></rowBreaks>"#,
    );
    let (ws, from_json, _) = parse_both(&xml);
    for model in [&ws, &from_json] {
        assert_eq!(
            model.auto_filter.as_ref().map(|af| af.range.as_str()),
            Some("A1:D50")
        );
        let breaks = model.page_breaks.as_ref().unwrap();
        assert_eq!(
            breaks.row_breaks.iter().map(|b| b.id).collect::<Vec<_>>(),
            [5]
        );
        assert!(breaks.col_breaks.is_empty());
        assert!(model.frozen_pane.is_none());
        assert!(model.page_setup.is_none());
        assert!(model.header_footer.is_none());
        assert!(model.preserved_extensions.is_empty());
    }
    let out = String::from_utf8(ws.to_xml().unwrap()).unwrap();
    assert_eq!(out.matches("<brk ").count(), 1, "{out}");
    assert!(out.contains(r#"<autoFilter ref="A1:D50"/>"#), "{out}");
}

#[test]
fn start_end_leaf_elements_read_like_self_closing_ones() {
    let self_closing = sheet(
        r#"<sheetPr><pageSetUpPr fitToPage="1"/></sheetPr><cols><col min="1" max="2" width="12" customWidth="1"/></cols><sheetData/><sheetProtection sheet="1" password="CBEB"/><mergeCells count="1"><mergeCell ref="A1:B2"/></mergeCells><hyperlinks><hyperlink ref="C3" location="Sheet1!A1" display="go"/></hyperlinks><pageSetup orientation="landscape" paperSize="9"/>"#,
    );
    let start_end = sheet(
        r#"<sheetPr><pageSetUpPr fitToPage="1"></pageSetUpPr></sheetPr><cols><col min="1" max="2" width="12" customWidth="1"></col></cols><sheetData/><sheetProtection sheet="1" password="CBEB"></sheetProtection><mergeCells count="1"><mergeCell ref="A1:B2"></mergeCell></mergeCells><hyperlinks><hyperlink ref="C3" location="Sheet1!A1" display="go"></hyperlink></hyperlinks><pageSetup orientation="landscape" paperSize="9"></pageSetup>"#,
    );
    let (ws_a, json_a, _) = parse_both(&self_closing);
    let (ws_b, json_b, _) = parse_both(&start_end);
    // The self-closing forms are read; the start/end forms must match them.
    assert!(ws_a.sheet_protection.is_some());
    assert!(ws_a.page_setup.is_some());
    assert_eq!(ws_a.page_setup.as_ref().unwrap().fit_to_page, Some(true));
    assert_eq!(ws_a.columns.len(), 1);
    assert_eq!(ws_a.merge_cells, ["A1:B2"]);
    assert_eq!(ws_a.hyperlinks.len(), 1);
    let model = |ws: &WorksheetXml| serde_json::to_string(ws).unwrap();
    assert_eq!(model(&ws_b), model(&ws_a));
    assert_eq!(model(&json_b), model(&json_a));
    assert_eq!(model(&json_a), model(&ws_a));
}

#[test]
fn cf_rule_keeps_every_formula_in_order() {
    let xml = sheet(
        r#"<sheetData/><conditionalFormatting sqref="A1:A9"><cfRule type="cellIs" operator="between" priority="1" dxfId="0"><formula>1</formula><formula>10</formula></cfRule><cfRule type="cellIs" operator="equal" priority="2" dxfId="1"><formula>5</formula></cfRule><cfRule type="expression" priority="3" dxfId="2"><formula>A</formula><formula>B</formula><formula>C</formula><formula>D</formula></cfRule></conditionalFormatting>"#,
    );
    let (ws, from_json, json) = parse_both(&xml);
    for model in [&ws, &from_json] {
        let rules = &model.conditional_formatting[0].rules;
        assert_eq!(rules[0].formula.as_deref(), Some("1"));
        assert_eq!(rules[0].additional_formulas, ["10"]);
        assert_eq!(rules[1].formula.as_deref(), Some("5"));
        assert!(rules[1].additional_formulas.is_empty());
        // CT_CfRule holds at most three <formula>; a fourth is dropped.
        assert_eq!(rules[2].formula.as_deref(), Some("A"));
        assert_eq!(rules[2].additional_formulas, ["B", "C"]);
    }
    // A one-formula rule has no additionalFormulas key.
    assert_eq!(json.matches("\"additionalFormulas\"").count(), 2, "{json}");
    let out = String::from_utf8(ws.to_xml().unwrap()).unwrap();
    assert!(
        out.contains(r#"operator="between"><formula>1</formula><formula>10</formula></cfRule>"#),
        "{out}"
    );
    assert!(
        out.contains(r"<formula>A</formula><formula>B</formula><formula>C</formula></cfRule>"),
        "{out}"
    );
    assert!(
        out.contains(r#"operator="equal"><formula>5</formula></cfRule>"#),
        "{out}"
    );
}

/// The `(reference, value)` pairs of one row's cells.
type CellRefs = Vec<(String, Option<String>)>;

/// The `(row index, cells)` pairs of every row in `model`.
fn row_refs(model: &WorksheetXml) -> Vec<(u32, CellRefs)> {
    model
        .rows
        .iter()
        .map(|r| {
            let cells = r
                .cells
                .iter()
                .map(|c| (c.reference.clone(), c.value.clone()))
                .collect();
            (r.index, cells)
        })
        .collect()
}

#[test]
fn rows_and_cells_without_r_take_the_next_position() {
    // ECMA-376 Part 1 §18.3.1.4 and §18.3.1.73: `c@r` and `row@r` are
    // optional; an absent one continues from the previous row or cell.
    let xml = sheet(
        r#"<sheetData><row><c><v>1</v></c><c><v>2</v></c></row><row><c r="C2"><v>3</v></c><c><v>4</v></c></row></sheetData>"#,
    );
    let (ws, from_json, _) = parse_both(&xml);
    let some = |s: &str| Some(s.to_owned());
    let expected = vec![
        (
            1,
            vec![("A1".to_owned(), some("1")), ("B1".to_owned(), some("2"))],
        ),
        (
            2,
            vec![("C2".to_owned(), some("3")), ("D2".to_owned(), some("4"))],
        ),
    ];
    assert_eq!(row_refs(&ws), expected);
    assert_eq!(row_refs(&from_json), expected);
}

#[test]
fn self_closing_rows_and_cells_without_r_take_the_next_position() {
    let xml = sheet(
        r#"<sheetData><row r="3"><c/><c r="E3"/><c/></row><row/><row><c t="b"><v>1</v></c></row></sheetData>"#,
    );
    let (ws, from_json, _) = parse_both(&xml);
    let expected = vec![
        (
            3,
            vec![
                ("A3".to_owned(), None),
                ("E3".to_owned(), None),
                ("F3".to_owned(), None),
            ],
        ),
        (4, vec![]),
        (5, vec![("A5".to_owned(), Some("1".to_owned()))]),
    ];
    assert_eq!(row_refs(&ws), expected);
    assert_eq!(row_refs(&from_json), expected);
}

#[test]
fn rows_and_cells_with_r_read_verbatim() {
    let xml = sheet(
        r#"<sheetData><row r="7"><c r="B7"><v>1</v></c><c r="A7"><v>2</v></c></row><row r="2"><c r="Z2"><v>3</v></c></row></sheetData>"#,
    );
    let (ws, from_json, _) = parse_both(&xml);
    let some = |s: &str| Some(s.to_owned());
    let expected = vec![
        (
            7,
            vec![("B7".to_owned(), some("1")), ("A7".to_owned(), some("2"))],
        ),
        (2, vec![("Z2".to_owned(), some("3"))]),
    ];
    assert_eq!(row_refs(&ws), expected);
    assert_eq!(row_refs(&from_json), expected);
}

#[test]
fn a_cell_without_r_past_xfd_is_an_xml_parse_error() {
    for cell in ["<c/>", "<c><v>1</v></c>"] {
        let xml = sheet(&format!(
            r#"<sheetData><row r="1"><c r="XFD1"/>{cell}</row></sheetData>"#
        ));
        assert_matches!(
            WorksheetXml::parse_with_sst(xml.as_bytes(), None, &SheetContext::default()),
            Err(crate::ModernXlsxError::XmlParse(_)),
            "{cell}"
        );
        assert_matches!(
            parse_json(&xml),
            Err(crate::ModernXlsxError::XmlParse(_)),
            "{cell}"
        );
    }
}

// -----------------------------------------------------------------------
// cfRule formula count and empty formulas, nested <worksheet>
// -----------------------------------------------------------------------

/// A one-sheet workbook whose sheet `name` holds `worksheet`.
fn workbook_of(name: &str, worksheet: WorksheetXml) -> crate::WorkbookData {
    crate::WorkbookData {
        sheets: vec![crate::SheetData {
            name: name.to_owned(),
            state: None,
            worksheet,
        }],
        ..Default::default()
    }
}

#[test]
fn a_cf_rule_with_more_than_three_formulas_fails_the_write() {
    let xml = sheet(
        r#"<sheetData/><conditionalFormatting sqref="A1:A9"><cfRule type="expression" dxfId="0" priority="1"><formula>A1</formula><formula>B1</formula><formula>C1</formula></cfRule></conditionalFormatting>"#,
    );
    let mut ws = WorksheetXml::parse(xml.as_bytes()).unwrap();
    // Three formulas are the schema's limit and are written in order.
    assert!(
        written(&ws)
            .contains("<formula>A1</formula><formula>B1</formula><formula>C1</formula></cfRule>"),
        "{}",
        written(&ws)
    );
    let wb = workbook_of("Rules", ws.clone());
    crate::writer::write_xlsx(&wb).unwrap();
    assert!(
        crate::validate::validate_workbook(&wb)
            .issues
            .iter()
            .all(|i| i.category != crate::validate::IssueCategory::ConditionalFormatting)
    );

    ws.conditional_formatting[0].rules[0]
        .additional_formulas
        .push("D1".to_owned());
    let err = ws.to_xml().unwrap_err();
    assert_eq!(err.code(), "INVALID_WORKBOOK", "{err}");
    let wb = workbook_of("Rules", ws);
    let err = crate::writer::write_xlsx(&wb).unwrap_err();
    assert_eq!(err.code(), "INVALID_WORKBOOK", "{err}");
    assert!(err.to_string().contains("sheet 'Rules'"), "{err}");
    assert!(err.to_string().contains("4 formulas"), "{err}");
    let report = crate::validate::validate_workbook(&wb);
    let issue = report
        .issues
        .iter()
        .find(|i| i.category == crate::validate::IssueCategory::ConditionalFormatting)
        .expect("a conditional formatting issue");
    assert_eq!(issue.severity, crate::validate::Severity::Error);
    assert_eq!(issue.location, "Rules!A1:A9");
    assert!(issue.message.contains("4 formulas"), "{}", issue.message);
    assert!(!report.is_valid);
}

#[test]
fn an_empty_cf_formula_keeps_its_position_on_both_paths() {
    let xml = sheet(
        r#"<sheetData/><conditionalFormatting sqref="A1:A9"><cfRule type="cellIs" operator="between" priority="1" dxfId="0"><formula/><formula>5</formula></cfRule><cfRule type="expression" priority="2" dxfId="1"><formula>1</formula><formula/><formula>3</formula></cfRule></conditionalFormatting>"#,
    );
    let (ws, from_json, _) = parse_both(&xml);
    for model in [&ws, &from_json] {
        let rules = &model.conditional_formatting[0].rules;
        assert_eq!(rules[0].formula.as_deref(), Some(""));
        assert_eq!(rules[0].additional_formulas, ["5"]);
        assert_eq!(rules[1].formula.as_deref(), Some("1"));
        assert_eq!(rules[1].additional_formulas, ["", "3"]);
    }
    let back = WorksheetXml::parse(&ws.to_xml().unwrap()).unwrap();
    let rules = &back.conditional_formatting[0].rules;
    assert_eq!(rules[0].formula.as_deref(), Some(""));
    assert_eq!(rules[0].additional_formulas, ["5"]);
    assert_eq!(rules[1].additional_formulas, ["", "3"]);
}

#[test]
fn a_nested_worksheet_element_does_not_replace_the_root_namespaces() {
    let xml = sheet(&format!(
        r#"<sheetData/><worksheet xmlns:x14="urn:not-x14" xmlns:xm="urn:not-xm"><extLst><ext uri="{{00000000-0000-0000-0000-00000000000B}}"/></extLst></worksheet><extLst>{X14_CF_EXT}</extLst>"#
    ));
    let (ws, from_json, _) = parse_both(&xml);
    for model in [&ws, &from_json] {
        // The kept <ext> re-declares the document element's bindings; the
        // nested element is skipped whole, like any unknown child.
        assert_eq!(model.preserved_extensions, [x14_cf_ext_kept()]);
    }
}

// -----------------------------------------------------------------------
// The default sheet view (Excel on a 150% display (144 DPI) showed
// setRowHeight(1, 40) as 26.7 points without any <sheetView>, 40 with one)
// -----------------------------------------------------------------------

/// The `<sheetViews>` Excel writes for a sheet with no view of its own.
const DEFAULT_SHEET_VIEWS: &str = r#"<sheetViews><sheetView workbookViewId="0"/></sheetViews>"#;

#[test]
fn a_sheet_without_a_view_writes_excels_default_view_in_schema_position() {
    let ws = WorksheetXml {
        tab_color: Some("FFFF0000".into()),
        columns: vec![ColumnInfo {
            min: 1,
            max: 1,
            width: 30.0,
            hidden: false,
            custom_width: true,
            outline_level: None,
            collapsed: false,
            style_index: None,
            best_fit: false,
            phonetic: false,
        }],
        rows: vec![Row {
            index: 1,
            height: Some(40.0),
            custom_height: true,
            ..Default::default()
        }],
        ..Default::default()
    };
    let before = serde_json::to_string(&ws).unwrap();
    let xml = String::from_utf8(ws.to_xml().unwrap()).unwrap();
    assert_eq!(
        serde_json::to_string(&ws).unwrap(),
        before,
        "the model is not changed"
    );
    let at = xml.find(DEFAULT_SHEET_VIEWS).expect(&xml);
    assert_eq!(xml.matches("<sheetViews>").count(), 1, "{xml}");
    // CT_Worksheet: sheetPr, dimension, sheetViews, sheetFormatPr, cols, sheetData.
    assert!(xml.find("</sheetPr>").unwrap() < at, "{xml}");
    assert!(xml.find("<dimension ").unwrap() < at, "{xml}");
    assert!(at < xml.find("<cols>").unwrap(), "{xml}");
    assert!(at < xml.find("<sheetData>").unwrap(), "{xml}");

    let back = WorksheetXml::parse(xml.as_bytes()).unwrap();
    assert_eq!(
        back.sheet_view, None,
        "the bare default view reads as no view"
    );
    assert_eq!(back.frozen_pane, None);
    assert!(back.pane_selections.is_empty());

    let xml = String::from_utf8(WorksheetXml::default().to_xml().unwrap()).unwrap();
    assert!(
        xml.contains(&format!(
            r#"<dimension ref="A1"/>{DEFAULT_SHEET_VIEWS}<sheetData>"#
        )),
        "{xml}"
    );
}

#[test]
fn a_sheet_with_a_view_is_written_as_before() {
    let mut ws = WorksheetXml {
        rows: vec![Row {
            index: 1,
            height: Some(40.0),
            custom_height: true,
            ..Default::default()
        }],
        frozen_pane: Some(FrozenPane { rows: 1, cols: 0 }),
        ..Default::default()
    };
    let head = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><dimension ref="A1"/>"#;
    let rows = r#"<sheetData><row r="1" ht="40" customHeight="1"/></sheetData></worksheet>"#;
    assert_eq!(
        String::from_utf8(ws.to_xml().unwrap()).unwrap(),
        format!(
            r#"{head}<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>{rows}"#
        )
    );
    ws.frozen_pane = None;
    ws.sheet_view = Some(SheetViewData {
        show_grid_lines: false,
        tab_selected: true,
        zoom_scale: Some(120),
        ..SheetViewData::default()
    });
    ws.pane_selections = vec![PaneSelection {
        pane: None,
        active_cell: Some("B2".into()),
        sqref: Some("B2".into()),
    }];
    assert_eq!(
        String::from_utf8(ws.to_xml().unwrap()).unwrap(),
        format!(
            r#"{head}<sheetViews><sheetView showGridLines="0" tabSelected="1" zoomScale="120" workbookViewId="0"><selection activeCell="B2" sqref="B2"/></sheetView></sheetViews>{rows}"#
        )
    );
    ws.pane_selections.clear();
    ws.sheet_view = Some(SheetViewData::default());
    assert_eq!(
        String::from_utf8(ws.to_xml().unwrap()).unwrap(),
        format!("{head}{DEFAULT_SHEET_VIEWS}{rows}")
    );
}

#[test]
fn an_excel_sheet_view_with_content_is_kept_on_read() {
    // Excel's own view of the active sheet: tabSelected and a selection.
    for (body, check) in [
        (
            r#"<sheetViews><sheetView tabSelected="1" workbookViewId="0"><selection activeCell="C3" sqref="C3"/></sheetView></sheetViews><sheetData/>"#,
            "tabSelected",
        ),
        (
            r#"<sheetViews><sheetView workbookViewId="0"><selection activeCell="C3" sqref="C3"/></sheetView></sheetViews><sheetData/>"#,
            "selection",
        ),
        (
            r#"<sheetViews><sheetView zoomScale="80" workbookViewId="0"/></sheetViews><sheetData/>"#,
            "zoomScale",
        ),
        (
            r#"<sheetViews><sheetView showGridLines="0" workbookViewId="0"/></sheetViews><sheetData/>"#,
            "showGridLines",
        ),
        (
            r#"<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews><sheetData/>"#,
            "pane",
        ),
    ] {
        let (ws, from_json, _) = parse_both(&sheet(body));
        for model in [&ws, &from_json] {
            let kept = model.sheet_view.is_some()
                || model.frozen_pane.is_some()
                || !model.pane_selections.is_empty();
            assert!(kept, "{check}: {model:?}");
        }
        let xml = String::from_utf8(ws.to_xml().unwrap()).unwrap();
        let back = WorksheetXml::parse(xml.as_bytes()).unwrap();
        assert_eq!(back.sheet_view, ws.sheet_view, "{check}: {xml}");
        assert_eq!(back.frozen_pane, ws.frozen_pane, "{check}: {xml}");
        assert_eq!(back.pane_selections, ws.pane_selections, "{check}: {xml}");
    }
}

/// A shared-formula child without a cached value keeps `<f t="shared" si>`,
/// and the other cells of the group are written as they are.
#[test]
fn shared_formula_child_without_value_keeps_its_formula() {
    let child = |reference: &str, value: Option<&str>| Cell {
        reference: reference.to_string(),
        cell_type: CellType::Number,
        value: value.map(str::to_string),
        formula_type: Some("shared".to_string()),
        shared_index: Some(0),
        ..Default::default()
    };
    let ws = WorksheetXml {
        rows: vec![
            Row {
                index: 1,
                cells: vec![Cell {
                    reference: "D1".to_string(),
                    cell_type: CellType::Number,
                    value: Some("11".to_string()),
                    formula: Some("A1+10".to_string()),
                    formula_type: Some("shared".to_string()),
                    formula_ref: Some("D1:D4".to_string()),
                    shared_index: Some(0),
                    ..Default::default()
                }],
                ..Default::default()
            },
            Row {
                index: 2,
                cells: vec![child("D2", None)],
                ..Default::default()
            },
            Row {
                index: 3,
                cells: vec![child("D3", Some("13"))],
                ..Default::default()
            },
            Row {
                index: 4,
                cells: vec![child("D4", None)],
                ..Default::default()
            },
            Row {
                index: 5,
                cells: vec![
                    // An array master without a value, a normal formula
                    // without a value, and truly empty cells: unchanged.
                    Cell {
                        reference: "E5".to_string(),
                        formula: Some("SEQUENCE(2)".to_string()),
                        formula_type: Some("array".to_string()),
                        formula_ref: Some("E5:E6".to_string()),
                        ..Default::default()
                    },
                    Cell {
                        reference: "F5".to_string(),
                        formula: Some("A1+1".to_string()),
                        ..Default::default()
                    },
                    Cell {
                        reference: "G5".to_string(),
                        ..Default::default()
                    },
                    Cell {
                        reference: "H5".to_string(),
                        style_index: Some(3),
                        ..Default::default()
                    },
                ],
                ..Default::default()
            },
        ],
        ..Default::default()
    };
    let xml = String::from_utf8(ws.to_xml().unwrap()).unwrap();
    for expected in [
        r#"<c r="D1"><f t="shared" ref="D1:D4" si="0">A1+10</f><v>11</v></c>"#,
        r#"<c r="D2"><f t="shared" si="0"/></c>"#,
        r#"<c r="D3"><f t="shared" si="0"/><v>13</v></c>"#,
        r#"<c r="D4"><f t="shared" si="0"/></c>"#,
        r#"<c r="E5"><f t="array" ref="E5:E6">SEQUENCE(2)</f></c>"#,
        r#"<c r="F5"><f>A1+1</f></c>"#,
        r#"<c r="G5"/>"#,
        r#"<c r="H5" s="3"/>"#,
    ] {
        assert!(xml.contains(expected), "missing {expected} in {xml}");
    }
    let back = WorksheetXml::parse(xml.as_bytes()).unwrap();
    for row in &back.rows[1..=3] {
        let c = &row.cells[0];
        assert_eq!(c.formula_type.as_deref(), Some("shared"));
        assert_eq!(c.shared_index, Some(0));
    }
}

/// One cell per row, row `n` holding `cells[n - 1]`; written with `to_xml`.
fn shared_group_xml(cells: Vec<Cell>) -> String {
    let ws = WorksheetXml {
        rows: (1..)
            .zip(cells)
            .map(|(index, cell)| Row {
                index,
                cells: vec![cell],
                ..Default::default()
            })
            .collect(),
        ..Default::default()
    };
    String::from_utf8(ws.to_xml().unwrap()).unwrap()
}

/// A cell of shared group `si` in column D, with formula text `formula`
/// (`ref` `D1:D4` when it has text or is `D1`) and cached value `value`.
fn shared_cell(reference: &str, si: u32, formula: Option<&str>, value: Option<&str>) -> Cell {
    Cell {
        reference: reference.to_string(),
        cell_type: CellType::Number,
        value: value.map(str::to_string),
        formula: formula.map(str::to_string),
        formula_type: Some("shared".to_string()),
        formula_ref: (reference == "D1").then(|| "D1:D4".to_string()),
        shared_index: Some(si),
        ..Default::default()
    }
}

fn assert_cells(xml: &str, expected: &[&str]) {
    for want in expected {
        assert!(xml.contains(want), "missing {want} in {xml}");
    }
}

/// A child whose `si` names no master in the sheet gets no `<f>`: it is
/// written with its value alone, or empty. A child of a group that has its
/// master in the same sheet keeps `<f t="shared" si>`.
#[test]
fn shared_formula_child_without_a_master_is_written_without_f() {
    let xml = shared_group_xml(vec![
        shared_cell("D1", 0, Some("A1+10"), Some("11")),
        shared_cell("D2", 0, None, None),
        shared_cell("D3", 7, None, None),
        shared_cell("D4", 7, None, Some("14")),
    ]);
    assert_cells(
        &xml,
        &[
            r#"<c r="D1"><f t="shared" ref="D1:D4" si="0">A1+10</f><v>11</v></c>"#,
            r#"<c r="D2"><f t="shared" si="0"/></c>"#,
            r#"<c r="D3"/>"#,
            r#"<c r="D4"><v>14</v></c>"#,
        ],
    );
}

/// A master whose formula text is absent or empty but which keeps its
/// shared keys is written without `<f>`, and so are its children: with
/// their cached values as plain values, or empty.
#[test]
fn cleared_shared_master_and_its_children_are_written_without_f() {
    for cleared in [None, Some("")] {
        let xml = shared_group_xml(vec![
            shared_cell("D1", 0, cleared, Some("11")),
            shared_cell("D2", 0, None, None),
            shared_cell("D3", 0, None, Some("13")),
            shared_cell("D4", 0, Some(""), Some("14")),
        ]);
        assert_cells(
            &xml,
            &[
                r#"<c r="D1"><v>11</v></c>"#,
                r#"<c r="D2"/>"#,
                r#"<c r="D3"><v>13</v></c>"#,
                r#"<c r="D4"><v>14</v></c>"#,
            ],
        );
        assert!(!xml.contains("<f"), "{cleared:?}: {xml}");
        let back = WorksheetXml::parse(xml.as_bytes()).unwrap();
        let cells: Vec<&Cell> = back.rows.iter().flat_map(|r| &r.cells).collect();
        assert!(
            cells.iter().all(|c| c.formula.is_none()
                && c.formula_type.is_none()
                && c.shared_index.is_none()),
            "{cleared:?}: {cells:?}"
        );
        let values: Vec<Option<&str>> = cells.iter().map(|c| c.value.as_deref()).collect();
        assert_eq!(values, [Some("11"), None, Some("13"), Some("14")]);
    }
}

/// Formula cells outside a shared group are written as before: an array
/// master with empty text, a data-table cell (attributes, no text) with and
/// without a value, and a formula with text but no keys.
#[test]
fn array_and_data_table_formulas_are_written_as_before() {
    let data_table = |reference: &str, value: Option<&str>| Cell {
        reference: reference.to_string(),
        value: value.map(str::to_string),
        formula_type: Some("dataTable".to_string()),
        formula_ref: Some("B2:C3".to_string()),
        formula_r1: Some("A1".to_string()),
        formula_r2: Some("A2".to_string()),
        formula_dt2d: Some(true),
        ..Default::default()
    };
    let xml = shared_group_xml(vec![
        Cell {
            reference: "E1".to_string(),
            formula: Some(String::new()),
            formula_type: Some("array".to_string()),
            formula_ref: Some("E1:E2".to_string()),
            ..Default::default()
        },
        data_table("B2", Some("5")),
        data_table("B3", None),
        Cell {
            reference: "F4".to_string(),
            formula: Some(String::new()),
            ..Default::default()
        },
    ]);
    assert_cells(
        &xml,
        &[
            r#"<c r="E1"><f t="array" ref="E1:E2"></f></c>"#,
            r#"<c r="B2"><f t="dataTable" ref="B2:C3" r1="A1" r2="A2" dt2D="1"/><v>5</v></c>"#,
            r#"<c r="B3"><f t="dataTable" ref="B2:C3" r1="A1" r2="A2" dt2D="1"/></c>"#,
            r#"<c r="F4"><f></f></c>"#,
        ],
    );
}

/// `xml` read and written again: the second generation of a write.
fn second_generation(xml: &str) -> String {
    let back = WorksheetXml::parse(xml.as_bytes()).unwrap();
    String::from_utf8(back.to_xml().unwrap()).unwrap()
}

/// A cell of shared group `si` with formula text `formula`, ref `formula_ref`
/// and cached value `value`.
fn shared_cell_ref(
    reference: &str,
    si: u32,
    formula: Option<&str>,
    formula_ref: Option<&str>,
    value: Option<&str>,
) -> Cell {
    Cell {
        formula_ref: formula_ref.map(str::to_string),
        ..shared_cell(reference, si, formula, value)
    }
}

/// A shared group whose master D1 was cleared (by
/// `cell.formula = null`, which drops its keys, or in a raw model that keeps
/// them): D2 is a child that carries its own text and D3, D4 are children
/// without text. No cell is the master of group 0, so D2 is written as a
/// plain formula and D3, D4 with their values alone; the second generation
/// is the same.
#[test]
fn shared_group_without_master_writes_a_text_child_as_a_plain_formula() {
    let cleared_master = [
        Cell {
            reference: "D1".to_string(),
            cell_type: CellType::Number,
            value: Some("2".to_string()),
            ..Default::default()
        },
        shared_cell_ref("D1", 0, None, Some("D1:D4"), Some("2")),
    ];
    for master in cleared_master {
        let xml = shared_group_xml(vec![
            master,
            shared_cell_ref("D2", 0, Some("B2*2"), None, Some("4")),
            shared_cell_ref("D3", 0, None, None, Some("6")),
            shared_cell_ref("D4", 0, None, None, Some("8")),
        ]);
        assert_cells(
            &xml,
            &[
                r#"<c r="D1"><v>2</v></c>"#,
                r#"<c r="D2"><f>B2*2</f><v>4</v></c>"#,
                r#"<c r="D3"><v>6</v></c>"#,
                r#"<c r="D4"><v>8</v></c>"#,
            ],
        );
        assert!(!xml.contains("t=\"shared\""), "{xml}");
        assert_eq!(second_generation(&xml), xml);
    }
}

/// A shared cell with a ref whose text is blank, or only `=` (which the
/// writer strips), is no master: it and the children of its group are
/// written with their values alone; a child that carries text gets a plain
/// formula.
#[test]
fn shared_master_with_blank_or_bare_equals_text_is_no_master() {
    for text in [" ", "=", "= "] {
        let xml = shared_group_xml(vec![
            shared_cell_ref("D1", 0, Some(text), Some("D1:D4"), Some("11")),
            shared_cell_ref("D2", 0, None, None, Some("12")),
            shared_cell_ref("D3", 0, Some("=A3+10"), None, Some("13")),
            shared_cell_ref("D4", 0, None, None, None),
        ]);
        assert_cells(
            &xml,
            &[
                r#"<c r="D1"><v>11</v></c>"#,
                r#"<c r="D2"><v>12</v></c>"#,
                r#"<c r="D3"><f>A3+10</f><v>13</v></c>"#,
                r#"<c r="D4"/>"#,
            ],
        );
        assert!(!xml.contains("t=\"shared\""), "{text:?}: {xml}");
        assert_eq!(second_generation(&xml), xml, "{text:?}");
    }
}

/// A shared cell with text but no ref is no master: without another master
/// of its group it is written as a plain formula and its textless children
/// with their values alone.
#[test]
fn shared_master_without_ref_is_written_as_a_plain_formula() {
    for formula_ref in [None, Some("")] {
        let xml = shared_group_xml(vec![
            shared_cell_ref("D1", 0, Some("A1+10"), formula_ref, Some("11")),
            shared_cell_ref("D2", 0, None, None, Some("12")),
            shared_cell_ref("D3", 0, None, None, None),
        ]);
        assert_cells(
            &xml,
            &[
                r#"<c r="D1"><f>A1+10</f><v>11</v></c>"#,
                r#"<c r="D2"><v>12</v></c>"#,
                r#"<c r="D3"/>"#,
            ],
        );
        assert!(!xml.contains("t=\"shared\""), "{formula_ref:?}: {xml}");
        assert_eq!(second_generation(&xml), xml, "{formula_ref:?}");
    }
}

/// A child that carries text in a group whose master is in the sheet is
/// written as read, `<f t="shared" si>text</f>`, byte for byte through two
/// generations; the master and the textless children likewise.
#[test]
fn shared_text_child_with_its_master_round_trips_byte_for_byte() {
    let source = concat!(
        r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>"#,
        r#"<row r="1"><c r="D1"><f t="shared" ref="D1:D4" si="0">A1+10</f><v>11</v></c></row>"#,
        r#"<row r="2"><c r="D2"><f t="shared" si="0">A2+10</f><v>12</v></c></row>"#,
        r#"<row r="3"><c r="D3"><f t="shared" si="0"/></c></row>"#,
        r#"<row r="4"><c r="D4"><f t="shared" si="0"/><v>14</v></c></row>"#,
        r#"</sheetData></worksheet>"#,
    );
    let first = second_generation(source);
    assert_cells(
        &first,
        &[
            r#"<c r="D1"><f t="shared" ref="D1:D4" si="0">A1+10</f><v>11</v></c>"#,
            r#"<c r="D2"><f t="shared" si="0">A2+10</f><v>12</v></c>"#,
            r#"<c r="D3"><f t="shared" si="0"/></c>"#,
            r#"<c r="D4"><f t="shared" si="0"/><v>14</v></c>"#,
        ],
    );
    assert_eq!(second_generation(&first), first);
}
