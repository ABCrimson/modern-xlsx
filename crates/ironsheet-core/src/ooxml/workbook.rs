use quick_xml::events::{BytesDecl, BytesEnd, BytesStart, BytesText, Event};
use quick_xml::{Reader, Writer};

use crate::dates::DateSystem;
use crate::{IronsheetError, Result};

const SPREADSHEET_NS: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const RELATIONSHIPS_NS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

/// Parsed representation of `xl/workbook.xml`.
#[derive(Debug, Clone)]
pub struct WorkbookXml {
    pub sheets: Vec<SheetInfo>,
    pub date_system: DateSystem,
    pub defined_names: Vec<DefinedName>,
}

/// Metadata for a single worksheet as declared in the workbook.
#[derive(Debug, Clone)]
pub struct SheetInfo {
    pub name: String,
    pub sheet_id: u32,
    pub r_id: String,
    pub state: SheetState,
}

/// Visibility state of a worksheet.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum SheetState {
    Visible,
    Hidden,
    VeryHidden,
}

/// A defined name (named range or formula alias).
#[derive(Debug, Clone)]
pub struct DefinedName {
    pub name: String,
    pub value: String,
    pub sheet_id: Option<u32>,
}

impl WorkbookXml {
    /// Parse `xl/workbook.xml` from raw XML bytes.
    pub fn parse(data: &[u8]) -> Result<Self> {
        let mut reader = Reader::from_reader(data);
        reader.config_mut().trim_text(true);

        let mut buf = Vec::new();
        let mut sheets = Vec::new();
        let mut date_system = DateSystem::Date1900;
        let mut defined_names = Vec::new();

        // State for collecting definedName text content.
        let mut in_defined_name = false;
        let mut current_dn_name = String::new();
        let mut current_dn_sheet_id: Option<u32> = None;
        let mut current_dn_value = String::new();

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Empty(ref e)) => {
                    match e.local_name().as_ref() {
                        b"workbookPr" => {
                            for attr in e.attributes().flatten() {
                                if attr.key.local_name().as_ref() == b"date1904" {
                                    let val = String::from_utf8_lossy(&attr.value);
                                    if val == "1" || val.eq_ignore_ascii_case("true") {
                                        date_system = DateSystem::Date1904;
                                    }
                                }
                            }
                        }
                        b"sheet" => {
                            sheets.push(parse_sheet_element(e)?);
                        }
                        _ => {}
                    }
                }
                Ok(Event::Start(ref e)) => {
                    match e.local_name().as_ref() {
                        b"workbookPr" => {
                            for attr in e.attributes().flatten() {
                                if attr.key.local_name().as_ref() == b"date1904" {
                                    let val = String::from_utf8_lossy(&attr.value);
                                    if val == "1" || val.eq_ignore_ascii_case("true") {
                                        date_system = DateSystem::Date1904;
                                    }
                                }
                            }
                        }
                        b"sheet" => {
                            sheets.push(parse_sheet_element(e)?);
                        }
                        b"definedName" => {
                            in_defined_name = true;
                            current_dn_name.clear();
                            current_dn_sheet_id = None;
                            current_dn_value.clear();
                            for attr in e.attributes().flatten() {
                                match attr.key.local_name().as_ref() {
                                    b"name" => {
                                        current_dn_name =
                                            String::from_utf8_lossy(&attr.value).into_owned();
                                    }
                                    b"localSheetId" => {
                                        let val = String::from_utf8_lossy(&attr.value);
                                        current_dn_sheet_id = val.parse::<u32>().ok();
                                    }
                                    _ => {}
                                }
                            }
                        }
                        _ => {}
                    }
                }
                Ok(Event::Text(ref e)) if in_defined_name => {
                    current_dn_value
                        .push_str(&String::from_utf8_lossy(e.as_ref()));
                }
                Ok(Event::End(ref e)) => {
                    if e.local_name().as_ref() == b"definedName" && in_defined_name {
                        defined_names.push(DefinedName {
                            name: std::mem::take(&mut current_dn_name),
                            value: std::mem::take(&mut current_dn_value),
                            sheet_id: current_dn_sheet_id.take(),
                        });
                        in_defined_name = false;
                    }
                }
                Ok(Event::Eof) => break,
                Err(err) => {
                    return Err(IronsheetError::XmlParse(format!(
                        "error parsing workbook.xml: {err}"
                    )));
                }
                _ => {}
            }
            buf.clear();
        }

        Ok(WorkbookXml {
            sheets,
            date_system,
            defined_names,
        })
    }

    /// Serialize to a valid `xl/workbook.xml` string.
    pub fn to_xml(&self) -> Result<String> {
        let mut buf: Vec<u8> = Vec::new();
        let mut writer = Writer::new(&mut buf);

        let map_err = |e: std::io::Error| IronsheetError::XmlWrite(e.to_string());

        // XML declaration.
        writer
            .write_event(Event::Decl(BytesDecl::new("1.0", Some("UTF-8"), Some("yes"))))
            .map_err(map_err)?;

        // <workbook xmlns="..." xmlns:r="...">
        let mut wb = BytesStart::new("workbook");
        wb.push_attribute(("xmlns", SPREADSHEET_NS));
        wb.push_attribute(("xmlns:r", RELATIONSHIPS_NS));
        writer.write_event(Event::Start(wb)).map_err(map_err)?;

        // <workbookPr /> — only if date1904.
        if self.date_system == DateSystem::Date1904 {
            let mut pr = BytesStart::new("workbookPr");
            pr.push_attribute(("date1904", "1"));
            writer.write_event(Event::Empty(pr)).map_err(map_err)?;
        }

        // <sheets>
        writer
            .write_event(Event::Start(BytesStart::new("sheets")))
            .map_err(map_err)?;

        for sheet in &self.sheets {
            let mut elem = BytesStart::new("sheet");
            elem.push_attribute(("name", sheet.name.as_str()));
            let id_str = sheet.sheet_id.to_string();
            elem.push_attribute(("sheetId", id_str.as_str()));
            elem.push_attribute(("r:id", sheet.r_id.as_str()));
            match sheet.state {
                SheetState::Visible => {} // default, omit
                SheetState::Hidden => {
                    elem.push_attribute(("state", "hidden"));
                }
                SheetState::VeryHidden => {
                    elem.push_attribute(("state", "veryHidden"));
                }
            }
            writer.write_event(Event::Empty(elem)).map_err(map_err)?;
        }

        // </sheets>
        writer
            .write_event(Event::End(BytesEnd::new("sheets")))
            .map_err(map_err)?;

        // <definedNames> (only if non-empty)
        if !self.defined_names.is_empty() {
            writer
                .write_event(Event::Start(BytesStart::new("definedNames")))
                .map_err(map_err)?;

            for dn in &self.defined_names {
                let mut elem = BytesStart::new("definedName");
                elem.push_attribute(("name", dn.name.as_str()));
                if let Some(sid) = dn.sheet_id {
                    let sid_str = sid.to_string();
                    elem.push_attribute(("localSheetId", sid_str.as_str()));
                }
                writer
                    .write_event(Event::Start(elem))
                    .map_err(map_err)?;
                writer
                    .write_event(Event::Text(BytesText::new(&dn.value)))
                    .map_err(map_err)?;
                writer
                    .write_event(Event::End(BytesEnd::new("definedName")))
                    .map_err(map_err)?;
            }

            writer
                .write_event(Event::End(BytesEnd::new("definedNames")))
                .map_err(map_err)?;
        }

        // </workbook>
        writer
            .write_event(Event::End(BytesEnd::new("workbook")))
            .map_err(map_err)?;

        String::from_utf8(buf)
            .map_err(|e| IronsheetError::XmlWrite(format!("invalid UTF-8 in output: {e}")))
    }
}

/// Parse a `<sheet>` element (either Empty or Start) into a `SheetInfo`.
fn parse_sheet_element(e: &BytesStart<'_>) -> Result<SheetInfo> {
    let mut name = String::new();
    let mut sheet_id: u32 = 0;
    let mut r_id = String::new();
    let mut state = SheetState::Visible;

    for attr in e.attributes().flatten() {
        let key = attr.key.as_ref();
        let local_name = attr.key.local_name();
        let local = local_name.as_ref();
        match local {
            b"name" => {
                name = String::from_utf8_lossy(&attr.value).into_owned();
            }
            b"sheetId" => {
                let val = String::from_utf8_lossy(&attr.value);
                sheet_id = val.parse::<u32>().unwrap_or(0);
            }
            b"state" => {
                let val = String::from_utf8_lossy(&attr.value);
                state = match val.as_ref() {
                    "hidden" => SheetState::Hidden,
                    "veryHidden" => SheetState::VeryHidden,
                    _ => SheetState::Visible,
                };
            }
            _ => {
                // Check for r:id — the full key might be "r:id" or just "id"
                // depending on namespace handling.
                if key == b"r:id" || local == b"id" {
                    r_id = String::from_utf8_lossy(&attr.value).into_owned();
                }
            }
        }
    }

    Ok(SheetInfo {
        name,
        sheet_id,
        r_id,
        state,
    })
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_parse_workbook_two_sheets() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
    <sheet name="Data" sheetId="2" r:id="rId2" state="hidden"/>
  </sheets>
</workbook>"#;

        let wb = WorkbookXml::parse(xml.as_bytes()).unwrap();
        assert_eq!(wb.sheets.len(), 2);

        assert_eq!(wb.sheets[0].name, "Sheet1");
        assert_eq!(wb.sheets[0].sheet_id, 1);
        assert_eq!(wb.sheets[0].r_id, "rId1");
        assert_eq!(wb.sheets[0].state, SheetState::Visible);

        assert_eq!(wb.sheets[1].name, "Data");
        assert_eq!(wb.sheets[1].sheet_id, 2);
        assert_eq!(wb.sheets[1].r_id, "rId2");
        assert_eq!(wb.sheets[1].state, SheetState::Hidden);
    }

    #[test]
    fn test_detect_date1904() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <workbookPr date1904="1"/>
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>"#;

        let wb = WorkbookXml::parse(xml.as_bytes()).unwrap();
        assert_eq!(wb.date_system, DateSystem::Date1904);
    }

    #[test]
    fn test_default_date_system_is_1900() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>"#;

        let wb = WorkbookXml::parse(xml.as_bytes()).unwrap();
        assert_eq!(wb.date_system, DateSystem::Date1900);
    }

    #[test]
    fn test_parse_defined_names() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
  <definedNames>
    <definedName name="MyRange" localSheetId="0">Sheet1!$A$1:$B$10</definedName>
    <definedName name="GlobalName">Sheet1!$C$1</definedName>
  </definedNames>
</workbook>"#;

        let wb = WorkbookXml::parse(xml.as_bytes()).unwrap();
        assert_eq!(wb.defined_names.len(), 2);

        assert_eq!(wb.defined_names[0].name, "MyRange");
        assert_eq!(wb.defined_names[0].value, "Sheet1!$A$1:$B$10");
        assert_eq!(wb.defined_names[0].sheet_id, Some(0));

        assert_eq!(wb.defined_names[1].name, "GlobalName");
        assert_eq!(wb.defined_names[1].value, "Sheet1!$C$1");
        assert_eq!(wb.defined_names[1].sheet_id, None);
    }

    #[test]
    fn test_parse_very_hidden_sheet() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Config" sheetId="1" r:id="rId1" state="veryHidden"/>
  </sheets>
</workbook>"#;

        let wb = WorkbookXml::parse(xml.as_bytes()).unwrap();
        assert_eq!(wb.sheets[0].state, SheetState::VeryHidden);
    }

    #[test]
    fn test_write_reparse_roundtrip() {
        let wb = WorkbookXml {
            sheets: vec![
                SheetInfo {
                    name: "Sheet1".to_string(),
                    sheet_id: 1,
                    r_id: "rId1".to_string(),
                    state: SheetState::Visible,
                },
                SheetInfo {
                    name: "Hidden".to_string(),
                    sheet_id: 2,
                    r_id: "rId2".to_string(),
                    state: SheetState::Hidden,
                },
            ],
            date_system: DateSystem::Date1904,
            defined_names: vec![DefinedName {
                name: "TestRange".to_string(),
                value: "Sheet1!$A$1:$Z$100".to_string(),
                sheet_id: Some(0),
            }],
        };

        let xml = wb.to_xml().unwrap();
        let wb2 = WorkbookXml::parse(xml.as_bytes()).unwrap();

        assert_eq!(wb2.sheets.len(), 2);
        assert_eq!(wb2.sheets[0].name, "Sheet1");
        assert_eq!(wb2.sheets[0].r_id, "rId1");
        assert_eq!(wb2.sheets[0].state, SheetState::Visible);
        assert_eq!(wb2.sheets[1].name, "Hidden");
        assert_eq!(wb2.sheets[1].r_id, "rId2");
        assert_eq!(wb2.sheets[1].state, SheetState::Hidden);

        assert_eq!(wb2.date_system, DateSystem::Date1904);

        assert_eq!(wb2.defined_names.len(), 1);
        assert_eq!(wb2.defined_names[0].name, "TestRange");
        assert_eq!(wb2.defined_names[0].value, "Sheet1!$A$1:$Z$100");
        assert_eq!(wb2.defined_names[0].sheet_id, Some(0));
    }

    #[test]
    fn test_default_workbook_single_sheet_1900() {
        let wb = WorkbookXml {
            sheets: vec![SheetInfo {
                name: "Sheet1".to_string(),
                sheet_id: 1,
                r_id: "rId1".to_string(),
                state: SheetState::Visible,
            }],
            date_system: DateSystem::Date1900,
            defined_names: Vec::new(),
        };

        let xml = wb.to_xml().unwrap();
        let wb2 = WorkbookXml::parse(xml.as_bytes()).unwrap();

        assert_eq!(wb2.sheets.len(), 1);
        assert_eq!(wb2.sheets[0].name, "Sheet1");
        assert_eq!(wb2.date_system, DateSystem::Date1900);
        assert!(wb2.defined_names.is_empty());

        // Verify date1904 attribute is NOT in the output for 1900 system.
        assert!(!xml.contains("date1904"));
    }
}
