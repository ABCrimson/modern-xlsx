use core::hint::cold_path;

use quick_xml::events::{BytesDecl, BytesEnd, BytesStart, BytesText, Event};
use quick_xml::{Reader, Writer};

use serde::{Deserialize, Serialize};

use crate::dates::DateSystem;
use crate::{ModernXlsxError, Result};

use super::SPREADSHEET_NS;
const RELATIONSHIPS_NS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

/// Workbook-level protection from `<workbookProtection>` (ECMA-376 §18.2.29).
#[derive(Debug, Clone, Serialize, Deserialize, PartialEq)]
#[serde(rename_all = "camelCase")]
pub struct WorkbookProtection {
    /// Prevent structural changes (add/delete/rename/move/copy sheets).
    #[serde(default, skip_serializing_if = "crate::ooxml::is_false")]
    pub lock_structure: bool,
    /// Prevent window position/size changes.
    #[serde(default, skip_serializing_if = "crate::ooxml::is_false")]
    pub lock_windows: bool,
    /// Prevent revision log changes.
    #[serde(default, skip_serializing_if = "crate::ooxml::is_false")]
    pub lock_revision: bool,
    /// Algorithm name for workbook password hash (e.g., "SHA-512").
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub workbook_algorithm_name: Option<String>,
    /// Base64-encoded hash value for workbook password.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub workbook_hash_value: Option<String>,
    /// Base64-encoded salt for workbook password hash.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub workbook_salt_value: Option<String>,
    /// Spin count for workbook password hash iteration.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub workbook_spin_count: Option<u32>,
    /// Algorithm name for revision password hash.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub revisions_algorithm_name: Option<String>,
    /// Base64-encoded hash value for revision password.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub revisions_hash_value: Option<String>,
    /// Base64-encoded salt for revision password hash.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub revisions_salt_value: Option<String>,
    /// Spin count for revision password hash iteration.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub revisions_spin_count: Option<u32>,
    /// Legacy 16-bit password hash (hex string).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub workbook_password: Option<String>,
    /// Legacy 16-bit revision password hash (hex string).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub revisions_password: Option<String>,
}

/// Parsed representation of `xl/workbook.xml`.
#[derive(Debug, Clone)]
pub struct WorkbookXml {
    pub sheets: Vec<SheetInfo>,
    pub date_system: DateSystem,
    pub defined_names: Vec<DefinedName>,
    pub workbook_views: Vec<WorkbookView>,
    pub protection: Option<WorkbookProtection>,
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
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct DefinedName {
    pub name: String,
    pub value: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub sheet_id: Option<u32>,
}

/// A workbook view as declared in `<bookViews><workbookView .../>`.
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct WorkbookView {
    pub active_tab: u32,
    pub first_sheet: u32,
    pub show_horizontal_scroll: bool,
    pub show_vertical_scroll: bool,
    pub show_sheet_tabs: bool,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub window_width: Option<u32>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub window_height: Option<u32>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tab_ratio: Option<u32>,
}

impl Default for WorkbookView {
    fn default() -> Self {
        Self {
            active_tab: 0,
            first_sheet: 0,
            show_horizontal_scroll: true,
            show_vertical_scroll: true,
            show_sheet_tabs: true,
            window_width: None,
            window_height: None,
            tab_ratio: None,
        }
    }
}

impl WorkbookXml {
    /// Parse `xl/workbook.xml` from raw XML bytes.
    pub fn parse(data: &[u8]) -> Result<Self> {
        let mut reader = Reader::from_reader(data);
        reader.config_mut().trim_text(true);

        let mut buf = Vec::with_capacity(512);
        let mut sheets = Vec::new();
        let mut date_system = DateSystem::Date1900;
        let mut defined_names = Vec::new();
        let mut workbook_views = Vec::new();
        let mut protection = None::<WorkbookProtection>;

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
                            if let Some(attr) = e.attributes().flatten()
                                .find(|a| a.key.local_name().as_ref() == b"date1904")
                            {
                                let val = std::str::from_utf8(&attr.value).unwrap_or_default();
                                if val == "1" || val.eq_ignore_ascii_case("true") {
                                    date_system = DateSystem::Date1904;
                                }
                            }
                        }
                        b"sheet" => {
                            sheets.push(parse_sheet_element(e)?);
                        }
                        b"workbookView" => {
                            workbook_views.push(parse_workbook_view_element(e));
                        }
                        b"workbookProtection" => {
                            protection = Some(parse_workbook_protection_element(e));
                        }
                        _ => {}
                    }
                }
                Ok(Event::Start(ref e)) => {
                    match e.local_name().as_ref() {
                        b"workbookPr" => {
                            if let Some(attr) = e.attributes().flatten()
                                .find(|a| a.key.local_name().as_ref() == b"date1904")
                            {
                                let val = std::str::from_utf8(&attr.value).unwrap_or_default();
                                if val == "1" || val.eq_ignore_ascii_case("true") {
                                    date_system = DateSystem::Date1904;
                                }
                            }
                        }
                        b"sheet" => {
                            sheets.push(parse_sheet_element(e)?);
                        }
                        b"workbookView" => {
                            workbook_views.push(parse_workbook_view_element(e));
                        }
                        b"workbookProtection" => {
                            protection = Some(parse_workbook_protection_element(e));
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
                                            std::str::from_utf8(&attr.value).unwrap_or_default().to_owned();
                                    }
                                    b"localSheetId" => {
                                        let val = std::str::from_utf8(&attr.value).unwrap_or_default();
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
                        .push_str(std::str::from_utf8(e.as_ref()).unwrap_or_default());
                }
                Ok(Event::End(ref e))
                    if e.local_name().as_ref() == b"definedName" && in_defined_name =>
                {
                    defined_names.push(DefinedName {
                        name: std::mem::take(&mut current_dn_name),
                        value: std::mem::take(&mut current_dn_value),
                        sheet_id: current_dn_sheet_id.take(),
                    });
                    in_defined_name = false;
                }
                Ok(Event::Eof) => break,
                Err(err) => {
                    cold_path();
                    return Err(ModernXlsxError::XmlParse(format!(
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
            workbook_views,
            protection,
        })
    }

    /// Serialize to valid `xl/workbook.xml` bytes.
    pub fn to_xml(&self) -> Result<Vec<u8>> {
        let mut buf: Vec<u8> = Vec::with_capacity(512);
        let mut writer = Writer::new(&mut buf);
        let mut ibuf = itoa::Buffer::new();

        let map_err = |e: std::io::Error| ModernXlsxError::XmlWrite(e.to_string());

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

        // <workbookProtection>
        if let Some(ref prot) = self.protection {
            let mut elem = BytesStart::new("workbookProtection");
            if prot.lock_structure {
                elem.push_attribute(("lockStructure", "1"));
            }
            if prot.lock_windows {
                elem.push_attribute(("lockWindows", "1"));
            }
            if prot.lock_revision {
                elem.push_attribute(("lockRevision", "1"));
            }
            if let Some(ref v) = prot.workbook_algorithm_name {
                elem.push_attribute(("workbookAlgorithmName", v.as_str()));
            }
            if let Some(ref v) = prot.workbook_hash_value {
                elem.push_attribute(("workbookHashValue", v.as_str()));
            }
            if let Some(ref v) = prot.workbook_salt_value {
                elem.push_attribute(("workbookSaltValue", v.as_str()));
            }
            if let Some(sc) = prot.workbook_spin_count {
                elem.push_attribute(("workbookSpinCount", ibuf.format(sc)));
            }
            if let Some(ref v) = prot.revisions_algorithm_name {
                elem.push_attribute(("revisionsAlgorithmName", v.as_str()));
            }
            if let Some(ref v) = prot.revisions_hash_value {
                elem.push_attribute(("revisionsHashValue", v.as_str()));
            }
            if let Some(ref v) = prot.revisions_salt_value {
                elem.push_attribute(("revisionsSaltValue", v.as_str()));
            }
            if let Some(sc) = prot.revisions_spin_count {
                elem.push_attribute(("revisionsSpinCount", ibuf.format(sc)));
            }
            if let Some(ref v) = prot.workbook_password {
                elem.push_attribute(("workbookPassword", v.as_str()));
            }
            if let Some(ref v) = prot.revisions_password {
                elem.push_attribute(("revisionsPassword", v.as_str()));
            }
            writer.write_event(Event::Empty(elem)).map_err(map_err)?;
        }

        // <bookViews> (only if non-empty)
        if !self.workbook_views.is_empty() {
            writer
                .write_event(Event::Start(BytesStart::new("bookViews")))
                .map_err(map_err)?;

            for view in &self.workbook_views {
                let mut elem = BytesStart::new("workbookView");
                elem.push_attribute(("activeTab", ibuf.format(view.active_tab)));
                elem.push_attribute(("firstSheet", ibuf.format(view.first_sheet)));
                elem.push_attribute((
                    "showHorizontalScroll",
                    if view.show_horizontal_scroll { "1" } else { "0" },
                ));
                elem.push_attribute((
                    "showVerticalScroll",
                    if view.show_vertical_scroll { "1" } else { "0" },
                ));
                elem.push_attribute((
                    "showSheetTabs",
                    if view.show_sheet_tabs { "1" } else { "0" },
                ));
                if let Some(w) = view.window_width {
                    elem.push_attribute(("windowWidth", ibuf.format(w)));
                }
                if let Some(h) = view.window_height {
                    elem.push_attribute(("windowHeight", ibuf.format(h)));
                }
                if let Some(ratio) = view.tab_ratio {
                    elem.push_attribute(("tabRatio", ibuf.format(ratio)));
                }
                writer.write_event(Event::Empty(elem)).map_err(map_err)?;
            }

            writer
                .write_event(Event::End(BytesEnd::new("bookViews")))
                .map_err(map_err)?;
        }

        // <sheets>
        writer
            .write_event(Event::Start(BytesStart::new("sheets")))
            .map_err(map_err)?;

        for sheet in &self.sheets {
            let mut elem = BytesStart::new("sheet");
            elem.push_attribute(("name", sheet.name.as_str()));
            elem.push_attribute(("sheetId", ibuf.format(sheet.sheet_id)));
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
                    elem.push_attribute(("localSheetId", ibuf.format(sid)));
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

        Ok(buf)
    }
}

/// Parse a `<workbookProtection>` element (either Empty or Start) into a `WorkbookProtection`.
fn parse_workbook_protection_element(e: &BytesStart<'_>) -> WorkbookProtection {
    let mut prot = WorkbookProtection {
        lock_structure: false,
        lock_windows: false,
        lock_revision: false,
        workbook_algorithm_name: None,
        workbook_hash_value: None,
        workbook_salt_value: None,
        workbook_spin_count: None,
        revisions_algorithm_name: None,
        revisions_hash_value: None,
        revisions_salt_value: None,
        revisions_spin_count: None,
        workbook_password: None,
        revisions_password: None,
    };
    for attr in e.attributes().flatten() {
        let val = std::str::from_utf8(&attr.value).unwrap_or_default();
        match attr.key.local_name().as_ref() {
            b"lockStructure" => prot.lock_structure = val == "1",
            b"lockWindows" => prot.lock_windows = val == "1",
            b"lockRevision" => prot.lock_revision = val == "1",
            b"workbookAlgorithmName" => prot.workbook_algorithm_name = Some(val.to_owned()),
            b"workbookHashValue" => prot.workbook_hash_value = Some(val.to_owned()),
            b"workbookSaltValue" => prot.workbook_salt_value = Some(val.to_owned()),
            b"workbookSpinCount" => prot.workbook_spin_count = val.parse().ok(),
            b"revisionsAlgorithmName" => prot.revisions_algorithm_name = Some(val.to_owned()),
            b"revisionsHashValue" => prot.revisions_hash_value = Some(val.to_owned()),
            b"revisionsSaltValue" => prot.revisions_salt_value = Some(val.to_owned()),
            b"revisionsSpinCount" => prot.revisions_spin_count = val.parse().ok(),
            b"workbookPassword" => prot.workbook_password = Some(val.to_owned()),
            b"revisionsPassword" => prot.revisions_password = Some(val.to_owned()),
            _ => {}
        }
    }
    prot
}

/// Parse a `<workbookView>` element (either Empty or Start) into a `WorkbookView`.
fn parse_workbook_view_element(e: &BytesStart<'_>) -> WorkbookView {
    let mut view = WorkbookView::default();

    for attr in e.attributes().flatten() {
        let local = attr.key.local_name();
        let val = std::str::from_utf8(&attr.value).unwrap_or_default();
        match local.as_ref() {
            b"activeTab" => {
                view.active_tab = val.parse::<u32>().unwrap_or(0);
            }
            b"firstSheet" => {
                view.first_sheet = val.parse::<u32>().unwrap_or(0);
            }
            b"showHorizontalScroll" => {
                view.show_horizontal_scroll = val != "0" && !val.eq_ignore_ascii_case("false");
            }
            b"showVerticalScroll" => {
                view.show_vertical_scroll = val != "0" && !val.eq_ignore_ascii_case("false");
            }
            b"showSheetTabs" => {
                view.show_sheet_tabs = val != "0" && !val.eq_ignore_ascii_case("false");
            }
            b"windowWidth" => {
                view.window_width = val.parse::<u32>().ok();
            }
            b"windowHeight" => {
                view.window_height = val.parse::<u32>().ok();
            }
            b"tabRatio" => {
                view.tab_ratio = val.parse::<u32>().ok();
            }
            _ => {}
        }
    }

    view
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
                name = std::str::from_utf8(&attr.value).unwrap_or_default().to_owned();
            }
            b"sheetId" => {
                let val = std::str::from_utf8(&attr.value).unwrap_or_default();
                sheet_id = val.parse::<u32>().unwrap_or(0);
            }
            b"state" => {
                let val = std::str::from_utf8(&attr.value).unwrap_or_default();
                state = match val {
                    "hidden" => SheetState::Hidden,
                    "veryHidden" => SheetState::VeryHidden,
                    _ => SheetState::Visible,
                };
            }
            _ => {
                // Check for r:id — the full key might be "r:id" or just "id"
                // depending on namespace handling.
                if key == b"r:id" || local == b"id" {
                    r_id = std::str::from_utf8(&attr.value).unwrap_or_default().to_owned();
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
            workbook_views: Vec::new(),
            protection: None,
        };

        let xml = wb.to_xml().unwrap();
        let wb2 = WorkbookXml::parse(&xml).unwrap();

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
    fn test_parse_workbook_views() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <bookViews>
    <workbookView activeTab="2" firstSheet="1" showHorizontalScroll="1" showVerticalScroll="0" showSheetTabs="1" windowWidth="28800" windowHeight="12300" tabRatio="600"/>
  </bookViews>
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>"#;

        let wb = WorkbookXml::parse(xml.as_bytes()).unwrap();
        assert_eq!(wb.workbook_views.len(), 1);
        let view = &wb.workbook_views[0];
        assert_eq!(view.active_tab, 2);
        assert_eq!(view.first_sheet, 1);
        assert!(view.show_horizontal_scroll);
        assert!(!view.show_vertical_scroll);
        assert!(view.show_sheet_tabs);
        assert_eq!(view.window_width, Some(28800));
        assert_eq!(view.window_height, Some(12300));
        assert_eq!(view.tab_ratio, Some(600));
    }

    #[test]
    fn test_workbook_view_roundtrip() {
        let wb = WorkbookXml {
            sheets: vec![SheetInfo {
                name: "Sheet1".to_string(),
                sheet_id: 1,
                r_id: "rId1".to_string(),
                state: SheetState::Visible,
            }],
            date_system: DateSystem::Date1900,
            defined_names: Vec::new(),
            workbook_views: vec![WorkbookView {
                active_tab: 1,
                first_sheet: 0,
                show_horizontal_scroll: true,
                show_vertical_scroll: false,
                show_sheet_tabs: true,
                window_width: Some(19200),
                window_height: Some(10800),
                tab_ratio: Some(500),
            }],
            protection: None,
        };

        let xml = wb.to_xml().unwrap();
        let wb2 = WorkbookXml::parse(&xml).unwrap();

        assert_eq!(wb2.workbook_views.len(), 1);
        let view = &wb2.workbook_views[0];
        assert_eq!(view.active_tab, 1);
        assert_eq!(view.first_sheet, 0);
        assert!(view.show_horizontal_scroll);
        assert!(!view.show_vertical_scroll);
        assert!(view.show_sheet_tabs);
        assert_eq!(view.window_width, Some(19200));
        assert_eq!(view.window_height, Some(10800));
        assert_eq!(view.tab_ratio, Some(500));
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
            workbook_views: Vec::new(),
            protection: None,
        };

        let xml = wb.to_xml().unwrap();
        let wb2 = WorkbookXml::parse(&xml).unwrap();

        assert_eq!(wb2.sheets.len(), 1);
        assert_eq!(wb2.sheets[0].name, "Sheet1");
        assert_eq!(wb2.date_system, DateSystem::Date1900);
        assert!(wb2.defined_names.is_empty());

        // Verify date1904 attribute is NOT in the output for 1900 system.
        let xml_str = String::from_utf8(xml).unwrap();
        assert!(!xml_str.contains("date1904"));
    }

    #[test]
    fn test_parse_workbook_protection_lock_structure() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <workbookProtection lockStructure="1"/>
  <sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>
</workbook>"#;
        let wb = WorkbookXml::parse(xml.as_bytes()).unwrap();
        let prot = wb.protection.as_ref().unwrap();
        assert!(prot.lock_structure);
        assert!(!prot.lock_windows);
        assert!(!prot.lock_revision);
    }

    #[test]
    fn test_parse_workbook_protection_with_hash() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <workbookProtection lockStructure="1" lockWindows="1" workbookAlgorithmName="SHA-512" workbookHashValue="abc123" workbookSaltValue="def456" workbookSpinCount="100000"/>
  <sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>
</workbook>"#;
        let wb = WorkbookXml::parse(xml.as_bytes()).unwrap();
        let prot = wb.protection.as_ref().unwrap();
        assert!(prot.lock_structure);
        assert!(prot.lock_windows);
        assert_eq!(prot.workbook_algorithm_name.as_deref(), Some("SHA-512"));
        assert_eq!(prot.workbook_hash_value.as_deref(), Some("abc123"));
        assert_eq!(prot.workbook_salt_value.as_deref(), Some("def456"));
        assert_eq!(prot.workbook_spin_count, Some(100000));
    }

    #[test]
    fn test_parse_no_workbook_protection() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>
</workbook>"#;
        let wb = WorkbookXml::parse(xml.as_bytes()).unwrap();
        assert!(wb.protection.is_none());
    }

    #[test]
    fn test_roundtrip_workbook_protection() {
        let wb = WorkbookXml {
            sheets: vec![SheetInfo {
                name: "Sheet1".to_owned(),
                sheet_id: 1,
                r_id: "rId1".to_owned(),
                state: SheetState::Visible,
            }],
            date_system: DateSystem::Date1900,
            defined_names: vec![],
            workbook_views: vec![],
            protection: Some(WorkbookProtection {
                lock_structure: true,
                lock_windows: true,
                lock_revision: false,
                workbook_algorithm_name: Some("SHA-512".to_owned()),
                workbook_hash_value: Some("abc123".to_owned()),
                workbook_salt_value: Some("def456".to_owned()),
                workbook_spin_count: Some(100000),
                revisions_algorithm_name: None,
                revisions_hash_value: None,
                revisions_salt_value: None,
                revisions_spin_count: None,
                workbook_password: None,
                revisions_password: None,
            }),
        };
        let xml = wb.to_xml().unwrap();
        let wb2 = WorkbookXml::parse(&xml).unwrap();
        let prot = wb2.protection.as_ref().unwrap();
        assert!(prot.lock_structure);
        assert!(prot.lock_windows);
        assert_eq!(prot.workbook_algorithm_name.as_deref(), Some("SHA-512"));
        assert_eq!(prot.workbook_hash_value.as_deref(), Some("abc123"));
        assert_eq!(prot.workbook_spin_count, Some(100000));
    }

    #[test]
    fn test_roundtrip_clear_protection() {
        let wb = WorkbookXml {
            sheets: vec![SheetInfo {
                name: "Sheet1".to_owned(),
                sheet_id: 1,
                r_id: "rId1".to_owned(),
                state: SheetState::Visible,
            }],
            date_system: DateSystem::Date1900,
            defined_names: vec![],
            workbook_views: vec![],
            protection: None,
        };
        let xml = wb.to_xml().unwrap();
        let xml_str = std::str::from_utf8(&xml).unwrap();
        assert!(!xml_str.contains("workbookProtection"));
    }
}
