//! Excel Table (ListObject) definitions — `xl/tables/table{n}.xml`.

use quick_xml::events::{BytesDecl, BytesEnd, BytesStart, BytesText, Event};
use quick_xml::{Reader, Writer};
use serde::{Deserialize, Serialize};

use super::push_entity;
use super::SPREADSHEET_NS;
use crate::{ModernXlsxError, Result};

fn default_header_row_count() -> u32 {
    1
}

fn default_true() -> bool {
    true
}

fn is_false(v: &bool) -> bool {
    !v
}

/// An Excel Table (ListObject) definition from `xl/tables/table{n}.xml`.
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct TableDefinition {
    /// Unique table ID across the workbook.
    pub id: u32,
    /// Internal name (used by VBA).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub name: Option<String>,
    /// Display name (must be unique, used in structured references).
    pub display_name: String,
    /// Cell range including header row, e.g. `"A1:D10"`.
    #[serde(rename = "ref")]
    pub ref_range: String,
    /// Number of header rows (default 1, set 0 for no header).
    #[serde(default = "default_header_row_count")]
    pub header_row_count: u32,
    /// Number of totals rows (default 0).
    #[serde(default)]
    pub totals_row_count: u32,
    /// Whether the totals row is shown.
    #[serde(default = "default_true")]
    pub totals_row_shown: bool,
    /// Table columns.
    pub columns: Vec<TableColumn>,
    /// Style information.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub style_info: Option<TableStyleInfo>,
    /// Table-scoped auto-filter range.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub auto_filter_ref: Option<String>,
}

/// A column within a table definition.
#[derive(Debug, Clone, Serialize, Deserialize, Default)]
#[serde(rename_all = "camelCase")]
pub struct TableColumn {
    /// Column ID (unique within the table).
    pub id: u32,
    /// Column name (matches header cell text).
    pub name: String,
    /// Totals row aggregate function (sum, min, max, average, count, countNums, stdDev, var, custom, none).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub totals_row_function: Option<String>,
    /// Totals row label text (when function is "none").
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub totals_row_label: Option<String>,
    /// Calculated column formula (structured references).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub calculated_column_formula: Option<String>,
    /// DXF ID for header row styling.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub header_row_dxf_id: Option<u32>,
    /// DXF ID for data area styling.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub data_dxf_id: Option<u32>,
    /// DXF ID for totals row styling.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub totals_row_dxf_id: Option<u32>,
}

/// Table style info from `<tableStyleInfo>`.
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct TableStyleInfo {
    /// Built-in style name (e.g. `"TableStyleMedium2"`).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub name: Option<String>,
    #[serde(default, skip_serializing_if = "is_false")]
    pub show_first_column: bool,
    #[serde(default, skip_serializing_if = "is_false")]
    pub show_last_column: bool,
    #[serde(default = "default_true")]
    pub show_row_stripes: bool,
    #[serde(default, skip_serializing_if = "is_false")]
    pub show_column_stripes: bool,
}

impl TableDefinition {
    /// Serialize this table definition to valid `xl/tables/table{n}.xml` bytes.
    pub fn to_xml(&self) -> Result<Vec<u8>> {
        let mut buf: Vec<u8> = Vec::with_capacity(512);
        let mut writer = Writer::new(&mut buf);
        let mut ibuf = itoa::Buffer::new();

        let map_err = |e: std::io::Error| ModernXlsxError::XmlWrite(e.to_string());

        // XML declaration.
        writer
            .write_event(Event::Decl(BytesDecl::new("1.0", Some("UTF-8"), Some("yes"))))
            .map_err(map_err)?;

        // <table xmlns="..." id="..." ...>
        let mut table_elem = BytesStart::new("table");
        table_elem.push_attribute(("xmlns", SPREADSHEET_NS));
        table_elem.push_attribute(("id", ibuf.format(self.id)));
        if let Some(ref name) = self.name {
            table_elem.push_attribute(("name", name.as_str()));
        }
        table_elem.push_attribute(("displayName", self.display_name.as_str()));
        table_elem.push_attribute(("ref", self.ref_range.as_str()));
        if self.header_row_count != 1 {
            table_elem.push_attribute(("headerRowCount", ibuf.format(self.header_row_count)));
        }
        if self.totals_row_count > 0 {
            table_elem.push_attribute(("totalsRowCount", ibuf.format(self.totals_row_count)));
        }
        if !self.totals_row_shown {
            table_elem.push_attribute(("totalsRowShown", "0"));
        }
        writer
            .write_event(Event::Start(table_elem))
            .map_err(map_err)?;

        // <autoFilter ref="..."/>
        if let Some(ref af_ref) = self.auto_filter_ref {
            let mut af = BytesStart::new("autoFilter");
            af.push_attribute(("ref", af_ref.as_str()));
            writer.write_event(Event::Empty(af)).map_err(map_err)?;
        }

        // <tableColumns count="N">
        let mut tc_elem = BytesStart::new("tableColumns");
        tc_elem.push_attribute(("count", ibuf.format(self.columns.len())));
        writer
            .write_event(Event::Start(tc_elem))
            .map_err(map_err)?;

        for col in &self.columns {
            let mut col_elem = BytesStart::new("tableColumn");
            col_elem.push_attribute(("id", ibuf.format(col.id)));
            col_elem.push_attribute(("name", col.name.as_str()));
            if let Some(ref func) = col.totals_row_function {
                col_elem.push_attribute(("totalsRowFunction", func.as_str()));
            }
            if let Some(ref label) = col.totals_row_label {
                col_elem.push_attribute(("totalsRowLabel", label.as_str()));
            }
            if let Some(dxf_id) = col.header_row_dxf_id {
                col_elem.push_attribute(("headerRowDxfId", ibuf.format(dxf_id)));
            }
            if let Some(dxf_id) = col.data_dxf_id {
                col_elem.push_attribute(("dataDxfId", ibuf.format(dxf_id)));
            }
            if let Some(dxf_id) = col.totals_row_dxf_id {
                col_elem.push_attribute(("totalsRowDxfId", ibuf.format(dxf_id)));
            }

            if let Some(ref formula) = col.calculated_column_formula {
                writer
                    .write_event(Event::Start(col_elem))
                    .map_err(map_err)?;

                writer
                    .write_event(Event::Start(BytesStart::new("calculatedColumnFormula")))
                    .map_err(map_err)?;
                writer
                    .write_event(Event::Text(BytesText::new(formula)))
                    .map_err(map_err)?;
                writer
                    .write_event(Event::End(BytesEnd::new("calculatedColumnFormula")))
                    .map_err(map_err)?;

                writer
                    .write_event(Event::End(BytesEnd::new("tableColumn")))
                    .map_err(map_err)?;
            } else {
                writer
                    .write_event(Event::Empty(col_elem))
                    .map_err(map_err)?;
            }
        }

        // </tableColumns>
        writer
            .write_event(Event::End(BytesEnd::new("tableColumns")))
            .map_err(map_err)?;

        // <tableStyleInfo .../>
        if let Some(ref si) = self.style_info {
            let mut si_elem = BytesStart::new("tableStyleInfo");
            if let Some(ref name) = si.name {
                si_elem.push_attribute(("name", name.as_str()));
            }
            si_elem.push_attribute(("showFirstColumn", if si.show_first_column { "1" } else { "0" }));
            si_elem.push_attribute(("showLastColumn", if si.show_last_column { "1" } else { "0" }));
            si_elem.push_attribute(("showRowStripes", if si.show_row_stripes { "1" } else { "0" }));
            si_elem.push_attribute(("showColumnStripes", if si.show_column_stripes { "1" } else { "0" }));
            writer
                .write_event(Event::Empty(si_elem))
                .map_err(map_err)?;
        }

        // </table>
        writer
            .write_event(Event::End(BytesEnd::new("table")))
            .map_err(map_err)?;

        Ok(buf)
    }

    /// Parse a table definition from `xl/tables/table{n}.xml` bytes.
    pub fn parse(data: &[u8]) -> Result<Self> {
        let mut reader = Reader::from_reader(data);
        reader.config_mut().trim_text(true);
        let mut buf = Vec::with_capacity(512);

        let mut id: u32 = 0;
        let mut name: Option<String> = None;
        let mut display_name = String::new();
        let mut ref_range = String::new();
        let mut header_row_count: u32 = 1;
        let mut totals_row_count: u32 = 0;
        let mut totals_row_shown: bool = true;
        let mut columns: Vec<TableColumn> = Vec::new();
        let mut style_info: Option<TableStyleInfo> = None;
        let mut auto_filter_ref: Option<String> = None;

        // Current column being parsed.
        let mut current_col: Option<TableColumn> = None;
        let mut in_calc_formula = false;
        let mut formula_buf = String::new();
        let mut in_totals_formula = false;
        let mut totals_formula_buf = String::new();

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) => {
                    match e.local_name().as_ref() {
                        b"table" => {
                            Self::parse_table_attrs(
                                e,
                                &mut id,
                                &mut name,
                                &mut display_name,
                                &mut ref_range,
                                &mut header_row_count,
                                &mut totals_row_count,
                                &mut totals_row_shown,
                            );
                        }
                        b"autoFilter" => {
                            for attr in e.attributes().flatten() {
                                if attr.key.as_ref() == b"ref" {
                                    auto_filter_ref = Some(
                                        std::str::from_utf8(&attr.value)
                                            .unwrap_or_default()
                                            .to_owned(),
                                    );
                                }
                            }
                        }
                        b"tableColumn" => {
                            current_col = Some(Self::parse_column_attrs(e));
                        }
                        b"calculatedColumnFormula" => {
                            in_calc_formula = true;
                            formula_buf.clear();
                        }
                        b"totalsRowFormula" => {
                            in_totals_formula = true;
                            totals_formula_buf.clear();
                        }
                        b"tableStyleInfo" => {
                            style_info = Some(Self::parse_style_info(e));
                        }
                        _ => {}
                    }
                }
                Ok(Event::Empty(ref e)) => {
                    match e.local_name().as_ref() {
                        b"table" => {
                            Self::parse_table_attrs(
                                e,
                                &mut id,
                                &mut name,
                                &mut display_name,
                                &mut ref_range,
                                &mut header_row_count,
                                &mut totals_row_count,
                                &mut totals_row_shown,
                            );
                        }
                        b"autoFilter" => {
                            for attr in e.attributes().flatten() {
                                if attr.key.as_ref() == b"ref" {
                                    auto_filter_ref = Some(
                                        std::str::from_utf8(&attr.value)
                                            .unwrap_or_default()
                                            .to_owned(),
                                    );
                                }
                            }
                        }
                        b"tableColumn" => {
                            columns.push(Self::parse_column_attrs(e));
                        }
                        b"tableStyleInfo" => {
                            style_info = Some(Self::parse_style_info(e));
                        }
                        _ => {}
                    }
                }
                Ok(Event::End(ref e)) => {
                    match e.local_name().as_ref() {
                        b"tableColumn" => {
                            if let Some(mut col) = current_col.take() {
                                if !formula_buf.is_empty() {
                                    col.calculated_column_formula =
                                        Some(std::mem::take(&mut formula_buf));
                                }
                                columns.push(col);
                            }
                        }
                        b"calculatedColumnFormula" => {
                            in_calc_formula = false;
                        }
                        b"totalsRowFormula" => {
                            in_totals_formula = false;
                        }
                        _ => {}
                    }
                }
                Ok(Event::Text(ref e)) => {
                    if in_calc_formula {
                        formula_buf
                            .push_str(std::str::from_utf8(e.as_ref()).unwrap_or_default());
                    } else if in_totals_formula {
                        totals_formula_buf
                            .push_str(std::str::from_utf8(e.as_ref()).unwrap_or_default());
                    }
                }
                Ok(Event::GeneralRef(ref e)) => {
                    if in_calc_formula {
                        push_entity(&mut formula_buf, e.as_ref());
                    } else if in_totals_formula {
                        push_entity(&mut totals_formula_buf, e.as_ref());
                    }
                }
                Ok(Event::Eof) => break,
                Err(err) => {
                    return Err(crate::ModernXlsxError::XmlParse(format!(
                        "table parse error: {err}"
                    )));
                }
                _ => {}
            }
            buf.clear();
        }

        Ok(TableDefinition {
            id,
            name,
            display_name,
            ref_range,
            header_row_count,
            totals_row_count,
            totals_row_shown,
            columns,
            style_info,
            auto_filter_ref,
        })
    }

    #[allow(clippy::too_many_arguments)]
    fn parse_table_attrs(
        e: &BytesStart<'_>,
        id: &mut u32,
        name: &mut Option<String>,
        display_name: &mut String,
        ref_range: &mut String,
        header_row_count: &mut u32,
        totals_row_count: &mut u32,
        totals_row_shown: &mut bool,
    ) {
        for attr in e.attributes().flatten() {
            match attr.key.as_ref() {
                b"id" => {
                    *id = std::str::from_utf8(&attr.value)
                        .unwrap_or("0")
                        .parse()
                        .unwrap_or(0);
                }
                b"name" => {
                    *name = Some(
                        std::str::from_utf8(&attr.value)
                            .unwrap_or_default()
                            .to_owned(),
                    );
                }
                b"displayName" => {
                    *display_name = std::str::from_utf8(&attr.value)
                        .unwrap_or_default()
                        .to_owned();
                }
                b"ref" => {
                    *ref_range = std::str::from_utf8(&attr.value)
                        .unwrap_or_default()
                        .to_owned();
                }
                b"headerRowCount" => {
                    *header_row_count = std::str::from_utf8(&attr.value)
                        .unwrap_or("1")
                        .parse()
                        .unwrap_or(1);
                }
                b"totalsRowCount" => {
                    *totals_row_count = std::str::from_utf8(&attr.value)
                        .unwrap_or("0")
                        .parse()
                        .unwrap_or(0);
                }
                b"totalsRowShown" => {
                    *totals_row_shown =
                        std::str::from_utf8(&attr.value).unwrap_or("1") != "0";
                }
                _ => {}
            }
        }
    }

    fn parse_column_attrs(e: &BytesStart<'_>) -> TableColumn {
        let mut col = TableColumn::default();
        for attr in e.attributes().flatten() {
            match attr.key.as_ref() {
                b"id" => {
                    col.id = std::str::from_utf8(&attr.value)
                        .unwrap_or("0")
                        .parse()
                        .unwrap_or(0);
                }
                b"name" => {
                    col.name = std::str::from_utf8(&attr.value)
                        .unwrap_or_default()
                        .to_owned();
                }
                b"totalsRowFunction" => {
                    col.totals_row_function = Some(
                        std::str::from_utf8(&attr.value)
                            .unwrap_or_default()
                            .to_owned(),
                    );
                }
                b"totalsRowLabel" => {
                    col.totals_row_label = Some(
                        std::str::from_utf8(&attr.value)
                            .unwrap_or_default()
                            .to_owned(),
                    );
                }
                b"headerRowDxfId" => {
                    col.header_row_dxf_id = std::str::from_utf8(&attr.value)
                        .ok()
                        .and_then(|v| v.parse().ok());
                }
                b"dataDxfId" => {
                    col.data_dxf_id = std::str::from_utf8(&attr.value)
                        .ok()
                        .and_then(|v| v.parse().ok());
                }
                b"totalsRowDxfId" => {
                    col.totals_row_dxf_id = std::str::from_utf8(&attr.value)
                        .ok()
                        .and_then(|v| v.parse().ok());
                }
                _ => {}
            }
        }
        col
    }

    fn parse_style_info(e: &BytesStart<'_>) -> TableStyleInfo {
        let mut si = TableStyleInfo {
            name: None,
            show_first_column: false,
            show_last_column: false,
            show_row_stripes: true,
            show_column_stripes: false,
        };
        for attr in e.attributes().flatten() {
            match attr.key.as_ref() {
                b"name" => {
                    si.name = Some(
                        std::str::from_utf8(&attr.value)
                            .unwrap_or_default()
                            .to_owned(),
                    );
                }
                b"showFirstColumn" => {
                    si.show_first_column =
                        std::str::from_utf8(&attr.value).unwrap_or("0") == "1";
                }
                b"showLastColumn" => {
                    si.show_last_column =
                        std::str::from_utf8(&attr.value).unwrap_or("0") == "1";
                }
                b"showRowStripes" => {
                    si.show_row_stripes =
                        std::str::from_utf8(&attr.value).unwrap_or("1") != "0";
                }
                b"showColumnStripes" => {
                    si.show_column_stripes =
                        std::str::from_utf8(&attr.value).unwrap_or("0") == "1";
                }
                _ => {}
            }
        }
        si
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_table_definition_serde_roundtrip() {
        let table = TableDefinition {
            id: 1,
            name: Some("Table1".into()),
            display_name: "Table1".into(),
            ref_range: "A1:C4".into(),
            header_row_count: 1,
            totals_row_count: 0,
            totals_row_shown: false,
            columns: vec![
                TableColumn {
                    id: 1,
                    name: "Name".into(),
                    ..Default::default()
                },
                TableColumn {
                    id: 2,
                    name: "Age".into(),
                    ..Default::default()
                },
                TableColumn {
                    id: 3,
                    name: "City".into(),
                    ..Default::default()
                },
            ],
            style_info: Some(TableStyleInfo {
                name: Some("TableStyleMedium2".into()),
                show_first_column: false,
                show_last_column: false,
                show_row_stripes: true,
                show_column_stripes: false,
            }),
            auto_filter_ref: Some("A1:C4".into()),
        };
        let json = serde_json::to_string(&table).unwrap();
        let parsed: TableDefinition = serde_json::from_str(&json).unwrap();
        assert_eq!(parsed.id, 1);
        assert_eq!(parsed.display_name, "Table1");
        assert_eq!(parsed.columns.len(), 3);
        assert_eq!(parsed.columns[0].name, "Name");
        assert!(parsed.style_info.is_some());
    }

    #[test]
    fn test_table_column_with_formula() {
        let col = TableColumn {
            id: 1,
            name: "Total".into(),
            totals_row_function: Some("sum".into()),
            calculated_column_formula: Some("Sales[@Qty]*Sales[@Price]".into()),
            ..Default::default()
        };
        let json = serde_json::to_string(&col).unwrap();
        let parsed: TableColumn = serde_json::from_str(&json).unwrap();
        assert_eq!(parsed.totals_row_function.as_deref(), Some("sum"));
        assert_eq!(
            parsed.calculated_column_formula.as_deref(),
            Some("Sales[@Qty]*Sales[@Price]")
        );
    }

    #[test]
    fn test_table_style_info_defaults() {
        let json = r#"{"showRowStripes":true}"#;
        let si: TableStyleInfo = serde_json::from_str(json).unwrap();
        assert!(si.show_row_stripes);
        assert!(!si.show_first_column);
        assert!(!si.show_last_column);
        assert!(!si.show_column_stripes);
        assert!(si.name.is_none());
    }

    #[test]
    fn test_parse_basic_table() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
       id="1" name="Table1" displayName="Table1"
       ref="A1:C4" totalsRowShown="0">
  <autoFilter ref="A1:C4"/>
  <tableColumns count="3">
    <tableColumn id="1" name="Name"/>
    <tableColumn id="2" name="Age"/>
    <tableColumn id="3" name="City"/>
  </tableColumns>
  <tableStyleInfo name="TableStyleMedium2"
    showFirstColumn="0" showLastColumn="0"
    showRowStripes="1" showColumnStripes="0"/>
</table>"#;
        let table = TableDefinition::parse(xml).unwrap();
        assert_eq!(table.id, 1);
        assert_eq!(table.name.as_deref(), Some("Table1"));
        assert_eq!(table.display_name, "Table1");
        assert_eq!(table.ref_range, "A1:C4");
        assert!(!table.totals_row_shown);
        assert_eq!(table.columns.len(), 3);
        assert_eq!(table.columns[0].name, "Name");
        assert_eq!(table.columns[1].name, "Age");
        assert_eq!(table.columns[2].name, "City");
        assert_eq!(table.auto_filter_ref.as_deref(), Some("A1:C4"));
        let si = table.style_info.as_ref().unwrap();
        assert_eq!(si.name.as_deref(), Some("TableStyleMedium2"));
        assert!(si.show_row_stripes);
        assert!(!si.show_column_stripes);
        assert!(!si.show_first_column);
    }

    #[test]
    fn test_parse_table_with_totals_and_formulas() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
       id="2" displayName="Sales" ref="A1:D6" totalsRowCount="1">
  <autoFilter ref="A1:D5"/>
  <tableColumns count="4">
    <tableColumn id="1" name="Product"/>
    <tableColumn id="2" name="Qty" totalsRowFunction="sum"/>
    <tableColumn id="3" name="Price" totalsRowFunction="average"/>
    <tableColumn id="4" name="Total" totalsRowFunction="sum">
      <calculatedColumnFormula>Sales[@Qty]*Sales[@Price]</calculatedColumnFormula>
    </tableColumn>
  </tableColumns>
  <tableStyleInfo name="TableStyleMedium9" showRowStripes="1"/>
</table>"#;
        let table = TableDefinition::parse(xml).unwrap();
        assert_eq!(table.id, 2);
        assert_eq!(table.display_name, "Sales");
        assert_eq!(table.totals_row_count, 1);
        assert_eq!(table.columns.len(), 4);
        assert_eq!(table.columns[1].totals_row_function.as_deref(), Some("sum"));
        assert_eq!(table.columns[2].totals_row_function.as_deref(), Some("average"));
        assert_eq!(
            table.columns[3].calculated_column_formula.as_deref(),
            Some("Sales[@Qty]*Sales[@Price]")
        );
    }

    #[test]
    fn test_table_xml_roundtrip_basic() {
        let table = TableDefinition {
            id: 1,
            name: Some("Table1".into()),
            display_name: "Table1".into(),
            ref_range: "A1:C4".into(),
            header_row_count: 1,
            totals_row_count: 0,
            totals_row_shown: false,
            columns: vec![
                TableColumn {
                    id: 1,
                    name: "Name".into(),
                    ..Default::default()
                },
                TableColumn {
                    id: 2,
                    name: "Age".into(),
                    ..Default::default()
                },
                TableColumn {
                    id: 3,
                    name: "City".into(),
                    ..Default::default()
                },
            ],
            style_info: Some(TableStyleInfo {
                name: Some("TableStyleMedium2".into()),
                show_first_column: false,
                show_last_column: false,
                show_row_stripes: true,
                show_column_stripes: false,
            }),
            auto_filter_ref: Some("A1:C4".into()),
        };
        let xml = table.to_xml().unwrap();
        let parsed = TableDefinition::parse(&xml).unwrap();
        assert_eq!(parsed.id, table.id);
        assert_eq!(parsed.name.as_deref(), Some("Table1"));
        assert_eq!(parsed.display_name, table.display_name);
        assert_eq!(parsed.ref_range, table.ref_range);
        assert_eq!(parsed.header_row_count, 1);
        assert_eq!(parsed.totals_row_count, 0);
        assert!(!parsed.totals_row_shown);
        assert_eq!(parsed.columns.len(), 3);
        assert_eq!(parsed.columns[0].name, "Name");
        assert_eq!(parsed.columns[1].name, "Age");
        assert_eq!(parsed.columns[2].name, "City");
        assert_eq!(parsed.auto_filter_ref.as_deref(), Some("A1:C4"));
        let si = parsed.style_info.as_ref().unwrap();
        assert_eq!(si.name.as_deref(), Some("TableStyleMedium2"));
        assert!(!si.show_first_column);
        assert!(!si.show_last_column);
        assert!(si.show_row_stripes);
        assert!(!si.show_column_stripes);
    }

    #[test]
    fn test_table_xml_roundtrip_with_formulas_and_totals() {
        let table = TableDefinition {
            id: 2,
            name: None,
            display_name: "Sales".into(),
            ref_range: "A1:D6".into(),
            header_row_count: 1,
            totals_row_count: 1,
            totals_row_shown: true,
            columns: vec![
                TableColumn {
                    id: 1,
                    name: "Product".into(),
                    totals_row_label: Some("Total".into()),
                    ..Default::default()
                },
                TableColumn {
                    id: 2,
                    name: "Qty".into(),
                    totals_row_function: Some("sum".into()),
                    ..Default::default()
                },
                TableColumn {
                    id: 3,
                    name: "Price".into(),
                    totals_row_function: Some("average".into()),
                    ..Default::default()
                },
                TableColumn {
                    id: 4,
                    name: "Total".into(),
                    totals_row_function: Some("sum".into()),
                    totals_row_label: None,
                    calculated_column_formula: Some("Sales[@Qty]*Sales[@Price]".into()),
                    header_row_dxf_id: Some(0),
                    data_dxf_id: Some(1),
                    totals_row_dxf_id: Some(2),
                },
            ],
            style_info: Some(TableStyleInfo {
                name: Some("TableStyleMedium9".into()),
                show_first_column: true,
                show_last_column: false,
                show_row_stripes: true,
                show_column_stripes: true,
            }),
            auto_filter_ref: Some("A1:D5".into()),
        };
        let xml = table.to_xml().unwrap();
        let parsed = TableDefinition::parse(&xml).unwrap();
        assert_eq!(parsed.id, 2);
        assert!(parsed.name.is_none());
        assert_eq!(parsed.display_name, "Sales");
        assert_eq!(parsed.totals_row_count, 1);
        assert!(parsed.totals_row_shown);
        assert_eq!(parsed.columns.len(), 4);
        assert_eq!(parsed.columns[0].totals_row_label.as_deref(), Some("Total"));
        assert_eq!(parsed.columns[1].totals_row_function.as_deref(), Some("sum"));
        assert_eq!(parsed.columns[2].totals_row_function.as_deref(), Some("average"));
        assert_eq!(
            parsed.columns[3].calculated_column_formula.as_deref(),
            Some("Sales[@Qty]*Sales[@Price]")
        );
        assert_eq!(parsed.columns[3].header_row_dxf_id, Some(0));
        assert_eq!(parsed.columns[3].data_dxf_id, Some(1));
        assert_eq!(parsed.columns[3].totals_row_dxf_id, Some(2));
        assert_eq!(parsed.auto_filter_ref.as_deref(), Some("A1:D5"));
        let si = parsed.style_info.as_ref().unwrap();
        assert!(si.show_first_column);
        assert!(si.show_column_stripes);
    }

    #[test]
    fn test_table_xml_no_style_no_autofilter() {
        let table = TableDefinition {
            id: 5,
            name: None,
            display_name: "Plain".into(),
            ref_range: "B2:D10".into(),
            header_row_count: 0,
            totals_row_count: 0,
            totals_row_shown: true,
            columns: vec![TableColumn {
                id: 1,
                name: "Col1".into(),
                ..Default::default()
            }],
            style_info: None,
            auto_filter_ref: None,
        };
        let xml = table.to_xml().unwrap();
        let xml_str = std::str::from_utf8(&xml).unwrap();
        // headerRowCount="0" should be written since != 1
        assert!(xml_str.contains("headerRowCount=\"0\""));
        // No totalsRowCount since 0 is default
        assert!(!xml_str.contains("totalsRowCount"));
        // No totalsRowShown since true is default
        assert!(!xml_str.contains("totalsRowShown"));
        // No autoFilter or tableStyleInfo
        assert!(!xml_str.contains("autoFilter"));
        assert!(!xml_str.contains("tableStyleInfo"));

        let parsed = TableDefinition::parse(&xml).unwrap();
        assert_eq!(parsed.id, 5);
        assert_eq!(parsed.header_row_count, 0);
        assert!(parsed.style_info.is_none());
        assert!(parsed.auto_filter_ref.is_none());
    }
}
