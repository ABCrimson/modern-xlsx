//! Excel Table (ListObject) definitions — `xl/tables/table{n}.xml`.
//!
//! Modelled: the `<table>` attributes `id`, `name`, `displayName`, `ref`,
//! `headerRowCount`, `totalsRowCount` and `totalsRowShown`; the `<autoFilter>`
//! `ref`; each `<tableColumn>` with its totals row function and label, DXF ids
//! and `calculatedColumnFormula` / `totalsRowFormula`; and `<tableStyleInfo>`.
//! Every other attribute and child element, such as `<sortState>` and
//! `<extLst>`, is ignored on read and absent after a write.

use core::hint::cold_path;

use quick_xml::events::{BytesEnd, BytesStart, BytesText, Event};
use quick_xml::{Reader, Writer};
use serde::{Deserialize, Serialize};

use super::SPREADSHEET_NS;
use super::xml::{attr_str, attrs, parse_attr, push_general_ref, xml_attr, xsd_bool};
use crate::{ModernXlsxError, Result};

use crate::ooxml::{write_xml_decl, xml_write_err};

#[inline]
fn default_header_row_count() -> u32 {
    1
}

/// An Excel Table (ListObject) definition from `xl/tables/table{n}.xml`.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
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
    #[serde(default = "crate::ooxml::default_true")]
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
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize, Default)]
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
    /// Totals row formula, used when `totals_row_function` is `"custom"`.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub totals_row_formula: Option<String>,
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
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct TableStyleInfo {
    /// Built-in style name (e.g. `"TableStyleMedium2"`).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub name: Option<String>,
    #[serde(default, skip_serializing_if = "crate::ooxml::is_false")]
    pub show_first_column: bool,
    #[serde(default, skip_serializing_if = "crate::ooxml::is_false")]
    pub show_last_column: bool,
    #[serde(default = "crate::ooxml::default_true")]
    pub show_row_stripes: bool,
    #[serde(default, skip_serializing_if = "crate::ooxml::is_false")]
    pub show_column_stripes: bool,
}

/// Attributes of the `<table>` root element, as read by `parse_table_attrs`.
struct TableAttrs {
    id: u32,
    name: Option<String>,
    display_name: String,
    ref_range: String,
    header_row_count: u32,
    totals_row_count: u32,
    totals_row_shown: bool,
}

impl Default for TableAttrs {
    /// The values an absent attribute leaves behind: one header row, no
    /// totals row, `totalsRowShown` on.
    fn default() -> Self {
        Self {
            id: 0,
            name: None,
            display_name: String::new(),
            ref_range: String::new(),
            header_row_count: 1,
            totals_row_count: 0,
            totals_row_shown: true,
        }
    }
}

/// Everything `TableDefinition::parse` accumulates while it walks the part.
#[derive(Default)]
struct ParseState {
    attrs: TableAttrs,
    columns: Vec<TableColumn>,
    style_info: Option<TableStyleInfo>,
    auto_filter_ref: Option<String>,
    /// The `<tableColumn>` whose children are being read.
    current_col: Option<TableColumn>,
    in_calc_formula: bool,
    formula_buf: String,
    in_totals_formula: bool,
    totals_formula_buf: String,
}

impl ParseState {
    /// Handle the opening tag of `e`, whether it arrived as `Start` or
    /// `Empty`.
    fn open(&mut self, e: &BytesStart<'_>) -> Result<()> {
        match e.local_name().as_ref() {
            "table" => self.attrs = TableDefinition::parse_table_attrs(e)?,
            "autoFilter" => {
                for attr in attrs(e) {
                    let attr = attr?;
                    if attr.key.as_ref() == "ref" {
                        self.auto_filter_ref = Some(attr_str(&attr)?.into_owned());
                    }
                }
            }
            "tableColumn" => self.current_col = Some(TableDefinition::parse_column_attrs(e)?),
            "calculatedColumnFormula" => {
                self.in_calc_formula = true;
                self.formula_buf.clear();
            }
            "totalsRowFormula" => {
                self.in_totals_formula = true;
                self.totals_formula_buf.clear();
            }
            "tableStyleInfo" => self.style_info = Some(TableDefinition::parse_style_info(e)?),
            _ => {}
        }
        Ok(())
    }

    /// Handle the end of the element `local`: for a `<tableColumn>` attach
    /// the collected formulas and keep the column.
    fn close(&mut self, local: &str) {
        match local {
            "tableColumn" => {
                if let Some(mut col) = self.current_col.take() {
                    if !self.formula_buf.is_empty() {
                        col.calculated_column_formula = Some(std::mem::take(&mut self.formula_buf));
                    }
                    if !self.totals_formula_buf.is_empty() {
                        col.totals_row_formula = Some(std::mem::take(&mut self.totals_formula_buf));
                    }
                    self.columns.push(col);
                }
            }
            "calculatedColumnFormula" => self.in_calc_formula = false,
            "totalsRowFormula" => self.in_totals_formula = false,
            _ => {}
        }
    }
}

impl TableDefinition {
    /// Serialize this table definition to valid `xl/tables/table{n}.xml` bytes.
    pub fn to_xml(&self) -> Result<Vec<u8>> {
        let mut buf: Vec<u8> = Vec::with_capacity(512 + self.columns.len() * 80);
        let mut writer = Writer::new(&mut buf);
        let mut ibuf = itoa::Buffer::new();

        // XML declaration.
        write_xml_decl(&mut writer)?;

        // <table xmlns="..." id="..." ...>
        let mut table_elem = BytesStart::new("table");
        table_elem.push_attribute(("xmlns", SPREADSHEET_NS));
        table_elem.push_attribute(("id", ibuf.format(self.id)));
        if let Some(ref name) = self.name {
            table_elem.push_attribute(xml_attr("name", name));
        }
        table_elem.push_attribute(xml_attr("displayName", &self.display_name));
        table_elem.push_attribute(xml_attr("ref", &self.ref_range));
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
            .map_err(xml_write_err)?;

        // <autoFilter ref="..."/>
        if let Some(ref af_ref) = self.auto_filter_ref {
            let mut af = BytesStart::new("autoFilter");
            af.push_attribute(xml_attr("ref", af_ref));
            writer
                .write_event(Event::Empty(af))
                .map_err(xml_write_err)?;
        }

        // <tableColumns count="N">
        let mut tc_elem = BytesStart::new("tableColumns");
        tc_elem.push_attribute(("count", ibuf.format(self.columns.len())));
        writer
            .write_event(Event::Start(tc_elem))
            .map_err(xml_write_err)?;

        for col in &self.columns {
            let mut col_elem = BytesStart::new("tableColumn");
            col_elem.push_attribute(("id", ibuf.format(col.id)));
            col_elem.push_attribute(xml_attr("name", &col.name));
            if let Some(ref func) = col.totals_row_function {
                col_elem.push_attribute(xml_attr("totalsRowFunction", func));
            }
            if let Some(ref label) = col.totals_row_label {
                col_elem.push_attribute(xml_attr("totalsRowLabel", label));
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

            if col.calculated_column_formula.is_some() || col.totals_row_formula.is_some() {
                writer
                    .write_event(Event::Start(col_elem))
                    .map_err(xml_write_err)?;

                // CT_TableColumn orders calculatedColumnFormula before
                // totalsRowFormula.
                for (tag, formula) in [
                    ("calculatedColumnFormula", &col.calculated_column_formula),
                    ("totalsRowFormula", &col.totals_row_formula),
                ] {
                    let Some(formula) = formula else { continue };
                    writer
                        .write_event(Event::Start(BytesStart::new(tag)))
                        .map_err(xml_write_err)?;
                    writer
                        .write_event(Event::Text(BytesText::new(
                            crate::ooxml::sanitize_xml_text(formula).as_ref(),
                        )))
                        .map_err(xml_write_err)?;
                    writer
                        .write_event(Event::End(BytesEnd::new(tag)))
                        .map_err(xml_write_err)?;
                }

                writer
                    .write_event(Event::End(BytesEnd::new("tableColumn")))
                    .map_err(xml_write_err)?;
            } else {
                writer
                    .write_event(Event::Empty(col_elem))
                    .map_err(xml_write_err)?;
            }
        }

        // </tableColumns>
        writer
            .write_event(Event::End(BytesEnd::new("tableColumns")))
            .map_err(xml_write_err)?;

        // <tableStyleInfo .../>
        if let Some(ref si) = self.style_info {
            let mut si_elem = BytesStart::new("tableStyleInfo");
            if let Some(ref name) = si.name {
                si_elem.push_attribute(xml_attr("name", name));
            }
            si_elem.push_attribute((
                "showFirstColumn",
                if si.show_first_column { "1" } else { "0" },
            ));
            si_elem.push_attribute((
                "showLastColumn",
                if si.show_last_column { "1" } else { "0" },
            ));
            si_elem.push_attribute((
                "showRowStripes",
                if si.show_row_stripes { "1" } else { "0" },
            ));
            si_elem.push_attribute((
                "showColumnStripes",
                if si.show_column_stripes { "1" } else { "0" },
            ));
            writer
                .write_event(Event::Empty(si_elem))
                .map_err(xml_write_err)?;
        }

        // </table>
        writer
            .write_event(Event::End(BytesEnd::new("table")))
            .map_err(xml_write_err)?;

        Ok(buf)
    }

    /// Parse a table definition from `xl/tables/table{n}.xml` bytes.
    ///
    /// # Errors
    ///
    /// `XmlParse` for malformed XML, an attribute holding a bad entity, or a
    /// numeric attribute that is not a valid number.
    pub fn parse(data: &[u8]) -> Result<Self> {
        let mut reader = Reader::from_reader(data);
        // Text is kept verbatim apart from XML 1.0 line-end normalization
        // (CRLF and CR read as LF): spaces around an entity inside a formula
        // are content. Whitespace between elements is ignored by the arms below.
        reader.config_mut().trim_text(false);
        let mut buf = Vec::with_capacity(512);
        let mut st = ParseState::default();

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) => st.open(e)?,
                Ok(Event::Empty(ref e)) => {
                    st.open(e)?;
                    // A self-closing element has no end tag: close it now.
                    st.close(e.local_name().as_ref());
                }
                Ok(Event::End(ref e)) => st.close(e.local_name().as_ref()),
                Ok(Event::Text(ref e)) => {
                    if st.in_calc_formula {
                        st.formula_buf.push_str(&e.xml10_content());
                    } else if st.in_totals_formula {
                        st.totals_formula_buf.push_str(&e.xml10_content());
                    }
                }
                Ok(Event::CData(ref e)) => {
                    if st.in_calc_formula {
                        st.formula_buf.push_str(&e.xml10_content());
                    } else if st.in_totals_formula {
                        st.totals_formula_buf.push_str(&e.xml10_content());
                    }
                }
                Ok(Event::GeneralRef(ref e)) => {
                    if st.in_calc_formula {
                        push_general_ref(&mut st.formula_buf, e)?;
                    } else if st.in_totals_formula {
                        push_general_ref(&mut st.totals_formula_buf, e)?;
                    }
                }
                Ok(Event::Eof) => break,
                Err(err) => {
                    cold_path();
                    return Err(ModernXlsxError::XmlParse(format!(
                        "table parse error: {err}"
                    )));
                }
                _ => {}
            }
            buf.clear();
        }

        let TableAttrs {
            id,
            name,
            display_name,
            ref_range,
            header_row_count,
            totals_row_count,
            totals_row_shown,
        } = st.attrs;

        Ok(TableDefinition {
            id,
            name,
            display_name,
            ref_range,
            header_row_count,
            totals_row_count,
            totals_row_shown,
            columns: st.columns,
            style_info: st.style_info,
            auto_filter_ref: st.auto_filter_ref,
        })
    }

    #[inline]
    fn parse_table_attrs(e: &BytesStart<'_>) -> Result<TableAttrs> {
        let mut out = TableAttrs::default();
        for attr in attrs(e) {
            let attr = attr?;
            match attr.key.as_ref() {
                "id" => out.id = parse_attr("table", &attr)?,
                "name" => out.name = Some(attr_str(&attr)?.into_owned()),
                "displayName" => out.display_name = attr_str(&attr)?.into_owned(),
                "ref" => out.ref_range = attr_str(&attr)?.into_owned(),
                "headerRowCount" => out.header_row_count = parse_attr("table", &attr)?,
                "totalsRowCount" => out.totals_row_count = parse_attr("table", &attr)?,
                "totalsRowShown" => {
                    out.totals_row_shown = xsd_bool(&attr_str(&attr)?).unwrap_or(true);
                }
                _ => {}
            }
        }
        Ok(out)
    }

    #[inline]
    fn parse_column_attrs(e: &BytesStart<'_>) -> Result<TableColumn> {
        let mut col = TableColumn::default();
        for attr in attrs(e) {
            let attr = attr?;
            match attr.key.as_ref() {
                "id" => col.id = parse_attr("tableColumn", &attr)?,
                "name" => col.name = attr_str(&attr)?.into_owned(),
                "totalsRowFunction" => {
                    col.totals_row_function = Some(attr_str(&attr)?.into_owned());
                }
                "totalsRowLabel" => col.totals_row_label = Some(attr_str(&attr)?.into_owned()),
                "headerRowDxfId" => {
                    col.header_row_dxf_id = Some(parse_attr("tableColumn", &attr)?);
                }
                "dataDxfId" => col.data_dxf_id = Some(parse_attr("tableColumn", &attr)?),
                "totalsRowDxfId" => {
                    col.totals_row_dxf_id = Some(parse_attr("tableColumn", &attr)?);
                }
                _ => {}
            }
        }
        Ok(col)
    }

    #[inline]
    fn parse_style_info(e: &BytesStart<'_>) -> Result<TableStyleInfo> {
        let mut si = TableStyleInfo {
            name: None,
            show_first_column: false,
            show_last_column: false,
            show_row_stripes: true,
            show_column_stripes: false,
        };
        for attr in attrs(e) {
            let attr = attr?;
            match attr.key.as_ref() {
                "name" => si.name = Some(attr_str(&attr)?.into_owned()),
                "showFirstColumn" => {
                    si.show_first_column = xsd_bool(&attr_str(&attr)?).unwrap_or(false);
                }
                "showLastColumn" => {
                    si.show_last_column = xsd_bool(&attr_str(&attr)?).unwrap_or(false);
                }
                "showRowStripes" => {
                    si.show_row_stripes = xsd_bool(&attr_str(&attr)?).unwrap_or(true);
                }
                "showColumnStripes" => {
                    si.show_column_stripes = xsd_bool(&attr_str(&attr)?).unwrap_or(false);
                }
                _ => {}
            }
        }
        Ok(si)
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn table_attrs_defaults_preserved() {
        let xml = br#"<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="3" displayName="T" ref="A1:B2"/>"#;
        let t = TableDefinition::parse(xml).unwrap();
        assert_eq!(t.id, 3);
        assert_eq!(t.name, None);
        assert_eq!(t.display_name, "T");
        assert_eq!(t.ref_range, "A1:B2");
        assert_eq!(t.header_row_count, 1);
        assert_eq!(t.totals_row_count, 0);
        assert!(t.totals_row_shown);
    }

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
        assert_eq!(
            table.columns[2].totals_row_function.as_deref(),
            Some("average")
        );
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
                    totals_row_formula: None,
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
        assert_eq!(
            parsed.columns[1].totals_row_function.as_deref(),
            Some("sum")
        );
        assert_eq!(
            parsed.columns[2].totals_row_function.as_deref(),
            Some("average")
        );
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

    const NS: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

    #[test]
    fn column_name_with_entities_parses_and_writes_once_escaped() {
        let xml = format!(
            r#"<table xmlns="{NS}" id="1" displayName="T" ref="A1:A2"><tableColumns count="1"><tableColumn id="1" name="Q&amp;A &quot;x&quot;"/></tableColumns></table>"#
        );
        let t = TableDefinition::parse(xml.as_bytes()).unwrap();
        assert_eq!(t.columns[0].name, "Q&A \"x\"");
        let out = String::from_utf8(t.to_xml().unwrap()).unwrap();
        assert!(out.contains(r#"name="Q&amp;A &quot;x&quot;""#), "{out}");
        assert!(!out.contains("&amp;amp;"), "{out}");
    }

    #[test]
    fn table_and_attribute_names_with_entities_do_not_double_escape() {
        let t = TableDefinition {
            id: 1,
            name: Some("Q&A".into()),
            display_name: "Q&A".into(),
            ref_range: "A1:B2".into(),
            header_row_count: 1,
            totals_row_count: 0,
            totals_row_shown: true,
            columns: vec![
                TableColumn {
                    id: 1,
                    name: "A & B".into(),
                    ..Default::default()
                },
                TableColumn {
                    id: 2,
                    name: "\"x\"".into(),
                    ..Default::default()
                },
            ],
            style_info: Some(TableStyleInfo {
                name: Some("Style & \"Co\"".into()),
                show_first_column: false,
                show_last_column: false,
                show_row_stripes: true,
                show_column_stripes: false,
            }),
            auto_filter_ref: Some("A1:B2".into()),
        };
        let xml = t.to_xml().unwrap();
        let s = std::str::from_utf8(&xml).unwrap();
        assert!(s.contains(r#"displayName="Q&amp;A""#), "{s}");
        assert!(s.contains(r#"name="A &amp; B""#), "{s}");
        assert!(s.contains(r#"name="&quot;x&quot;""#), "{s}");
        assert!(!s.contains("&amp;amp;"), "{s}");
        assert_eq!(TableDefinition::parse(&xml).unwrap(), t);
    }

    #[test]
    fn bad_entity_in_attribute_is_xml_parse_error() {
        let xml = format!(r#"<table xmlns="{NS}" id="1" displayName="A &bad; B" ref="A1"/>"#);
        assert!(matches!(
            TableDefinition::parse(xml.as_bytes()),
            Err(ModernXlsxError::XmlParse(_))
        ));
    }

    #[test]
    fn invalid_numeric_attribute_is_xml_parse_error() {
        let xml = format!(r#"<table xmlns="{NS}" id="x" displayName="T" ref="A1"/>"#);
        assert!(matches!(
            TableDefinition::parse(xml.as_bytes()),
            Err(ModernXlsxError::XmlParse(_))
        ));
    }

    #[test]
    fn custom_totals_row_formula_round_trips_after_calculated_formula() {
        let xml = format!(
            r#"<table xmlns="{NS}" id="1" displayName="T" ref="A1:B3" totalsRowCount="1"><tableColumns count="1"><tableColumn id="1" name="Amt" totalsRowFunction="custom"><calculatedColumnFormula>T[[#This Row],[Amt]]</calculatedColumnFormula><totalsRowFormula>SUBTOTAL(109,[Amt])*2</totalsRowFormula></tableColumn></tableColumns></table>"#
        );
        let t = TableDefinition::parse(xml.as_bytes()).unwrap();
        assert_eq!(
            t.columns[0].totals_row_formula.as_deref(),
            Some("SUBTOTAL(109,[Amt])*2")
        );
        let out = t.to_xml().unwrap();
        let s = std::str::from_utf8(&out).unwrap();
        let calc = s.find("<calculatedColumnFormula>").unwrap();
        let totals = s.find("<totalsRowFormula>SUBTOTAL(109,[Amt])*2</totalsRowFormula>");
        assert!(totals.unwrap() > calc, "{s}");
        assert_eq!(TableDefinition::parse(&out).unwrap(), t);
    }

    #[test]
    fn totals_row_formula_alone_is_written_with_start_and_end() {
        let t = TableDefinition {
            id: 1,
            name: None,
            display_name: "T".into(),
            ref_range: "A1:A3".into(),
            header_row_count: 1,
            totals_row_count: 1,
            totals_row_shown: true,
            columns: vec![TableColumn {
                id: 1,
                name: "A".into(),
                totals_row_function: Some("custom".into()),
                totals_row_formula: Some("1+1".into()),
                ..Default::default()
            }],
            style_info: None,
            auto_filter_ref: None,
        };
        let xml = t.to_xml().unwrap();
        assert!(
            std::str::from_utf8(&xml)
                .unwrap()
                .contains("<totalsRowFormula>1+1</totalsRowFormula></tableColumn>")
        );
        assert_eq!(TableDefinition::parse(&xml).unwrap(), t);
    }

    #[test]
    fn formula_keeps_spaces_around_entities() {
        let xml = format!(
            r#"<table xmlns="{NS}" id="1" displayName="T" ref="A1:A2"><tableColumns count="1"><tableColumn id="1" name="A"><calculatedColumnFormula>A1 &amp; " x " &amp; B1</calculatedColumnFormula></tableColumn></tableColumns></table>"#
        );
        let t = TableDefinition::parse(xml.as_bytes()).unwrap();
        assert_eq!(
            t.columns[0].calculated_column_formula.as_deref(),
            Some("A1 & \" x \" & B1")
        );
    }

    #[test]
    fn self_closing_table_parses() {
        let xml = format!(r#"<table xmlns="{NS}" id="4" displayName="T" ref="A1:B2"/>"#);
        let t = TableDefinition::parse(xml.as_bytes()).unwrap();
        assert_eq!(t.id, 4);
        assert!(t.columns.is_empty());
    }

    #[test]
    fn self_closing_autofilter_and_style_info_parse() {
        let xml = format!(
            r#"<table xmlns="{NS}" id="1" displayName="T" ref="A1:B2"><autoFilter ref="A1:B2"/><tableColumns count="0"/><tableStyleInfo name="S" showRowStripes="false"/></table>"#
        );
        let t = TableDefinition::parse(xml.as_bytes()).unwrap();
        assert_eq!(t.auto_filter_ref.as_deref(), Some("A1:B2"));
        assert!(!t.style_info.unwrap().show_row_stripes);
    }

    #[test]
    fn mismatched_end_tag_is_xml_parse_error() {
        assert!(matches!(
            TableDefinition::parse(b"<root><a></b></root>"),
            Err(ModernXlsxError::XmlParse(_))
        ));
    }

    fn one_column(column_children: &str) -> TableColumn {
        let xml = format!(
            r#"<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" displayName="T" ref="A1:A3" totalsRowCount="1"><tableColumns count="1"><tableColumn id="1" name="A">{column_children}</tableColumn></tableColumns></table>"#
        );
        TableDefinition::parse(xml.as_bytes())
            .unwrap()
            .columns
            .remove(0)
    }

    #[test]
    fn formula_line_ends_are_normalized_and_char_refs_kept() {
        let col = one_column(
            "<calculatedColumnFormula>[@A]\r\n+1</calculatedColumnFormula>\
             <totalsRowFormula>SUM(\r[A])&#13;&#10;</totalsRowFormula>",
        );
        assert_eq!(col.calculated_column_formula.as_deref(), Some("[@A]\n+1"));
        assert_eq!(col.totals_row_formula.as_deref(), Some("SUM(\n[A])\r\n"));
    }

    #[test]
    fn cdata_formulas_are_kept() {
        let col = one_column(
            "<calculatedColumnFormula><![CDATA[[@A]&\"<x>\"]]></calculatedColumnFormula>\
             <totalsRowFormula><![CDATA[SUBTOTAL(109,[A])]]></totalsRowFormula>",
        );
        assert_eq!(
            col.calculated_column_formula.as_deref(),
            Some("[@A]&\"<x>\"")
        );
        assert_eq!(col.totals_row_formula.as_deref(), Some("SUBTOTAL(109,[A])"));
    }
}
