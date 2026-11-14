//! `xl/workbook.xml`, the workbook part (ECMA-376 §18.2): sheet list, defined
//! names, workbook views, protection, calculation settings, file version,
//! workbook properties, external references and pivot cache references.
//!
//! [`WorkbookXml::parse`] reads the elements it models and ignores the rest;
//! [`WorkbookXml::to_xml`] writes only those, so unmodelled elements such as
//! `<extLst>` are absent after a round trip.

use core::hint::cold_path;

use quick_xml::events::attributes::Attribute;
use quick_xml::events::{BytesEnd, BytesStart, BytesText, Event};
use quick_xml::{Reader, Writer};

use serde::{Deserialize, Serialize};

use crate::dates::DateSystem;
use crate::ooxml::shared_strings::ExpansionBudget;
use crate::ooxml::xml::{
    REL_NS, SPREADSHEET_NS, XsdBuf, attr_str, attrs, fmt_xsd_double, parse_attr, push_general_ref,
    xml_attr, xsd_bool,
};
use crate::ooxml::{write_xml_decl, xml_write_err};
use crate::{ModernXlsxError, Result};

/// Workbook-level protection from `<workbookProtection>` (ECMA-376 §18.2.29).
#[derive(Debug, Clone, Default, Serialize, Deserialize, PartialEq)]
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

/// Calculation settings from `<calcPr>` (ECMA-376 §18.2.2, `CT_CalcPr`).
///
/// Every field is `None` when the attribute was absent, so a round trip never
/// invents a value; an absent attribute means its schema default.
#[derive(Debug, Clone, Default, Serialize, Deserialize, PartialEq)]
#[serde(rename_all = "camelCase")]
pub struct CalcPr {
    /// Calculation engine version that last calculated the workbook.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub calc_id: Option<u32>,
    /// `auto`, `autoNoTable` or `manual`.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub calc_mode: Option<String>,
    /// Recalculate every formula when the file opens.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub full_calc_on_load: Option<bool>,
    /// Reference style, `A1` or `R1C1`.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ref_mode: Option<String>,
    /// Allow iterative calculation of circular references.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub iterate: Option<bool>,
    /// Maximum number of iterations.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub iterate_count: Option<u32>,
    /// Maximum change between iterations.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub iterate_delta: Option<f64>,
    /// Calculate with the precision the cell displays instead of full precision.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub full_precision: Option<bool>,
    /// The calculation finished before the file was saved.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub calc_completed: Option<bool>,
    /// Recalculate before saving.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub calc_on_save: Option<bool>,
    /// Allow multi-threaded calculation.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub concurrent_calc: Option<bool>,
    /// Number of threads for manual multi-threaded calculation.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub concurrent_manual_count: Option<u32>,
    /// Force a full calculation on the next recalculation.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub force_full_calc: Option<bool>,
}

/// Application that last saved the workbook, from `<fileVersion>`
/// (ECMA-376 §18.2.13, `CT_FileVersion`).
#[derive(Debug, Clone, Default, Serialize, Deserialize, PartialEq)]
#[serde(rename_all = "camelCase")]
pub struct FileVersion {
    /// Application name, usually `xl`.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub app_name: Option<String>,
    /// Version of the application that last saved the file.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub last_edited: Option<String>,
    /// Oldest version that can open the file without loss.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub lowest_edited: Option<String>,
    /// Build number of the application.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub rup_build: Option<String>,
    /// VBA code name of the workbook.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub code_name: Option<String>,
}

/// Workbook properties from `<workbookPr>` (ECMA-376 §18.2.28,
/// `CT_WorkbookPr`) other than `date1904`, which lives in
/// [`WorkbookXml::date_system`].
#[derive(Debug, Clone, Default, Serialize, Deserialize, PartialEq)]
#[serde(rename_all = "camelCase")]
pub struct WorkbookPr {
    /// How objects are shown: `all`, `placeholders` or `none`.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub show_objects: Option<String>,
    /// Show the border of unselected tables.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub show_border_unselected_tables: Option<bool>,
    /// Remove personal information on save.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub filter_privacy: Option<bool>,
    /// Prompted solutions were offered.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub prompted_solutions: Option<bool>,
    /// Show ink annotations.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub show_ink_annotation: Option<bool>,
    /// Create a backup file on save.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub backup_file: Option<bool>,
    /// Save external link values.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub save_external_link_values: Option<bool>,
    /// How links to other workbooks update: `userSet`, `never` or `always`.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub update_links: Option<String>,
    /// VBA code name of the workbook.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub code_name: Option<String>,
    /// Hide the pivot field list.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub hide_pivot_field_list: Option<bool>,
    /// Show the pivot chart filter.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub show_pivot_chart_filter: Option<bool>,
    /// Allow query refresh.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub allow_refresh_query: Option<bool>,
    /// Publish items to the server.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub publish_items: Option<bool>,
    /// Run the compatibility checker on save.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub check_compatibility: Option<bool>,
    /// Compress pictures automatically.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub auto_compress_pictures: Option<bool>,
    /// Refresh all connections on open.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub refresh_all_connections: Option<bool>,
    /// Default theme version.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub default_theme_version: Option<u32>,
}

/// Parsed representation of `xl/workbook.xml`.
#[derive(Debug, Clone, Default)]
pub struct WorkbookXml {
    pub sheets: Vec<SheetInfo>,
    pub date_system: DateSystem,
    pub defined_names: Vec<DefinedName>,
    pub workbook_views: Vec<WorkbookView>,
    pub protection: Option<WorkbookProtection>,
    /// `<calcPr>`, written right after `</definedNames>`.
    pub calc_pr: Option<CalcPr>,
    /// `<fileVersion>`, written first.
    pub file_version: Option<FileVersion>,
    /// `<workbookPr>` attributes other than `date1904`.
    pub workbook_pr: Option<WorkbookPr>,
    /// `<externalReferences>`: the workbook relationship Id of each external
    /// link, in order (formulas refer to link `n` as `[n]`); written after
    /// `</sheets>`.
    pub external_references: Vec<String>,
    /// `<pivotCaches>`, written after `<calcPr>`. Only the `CT_Workbook` list
    /// is read: a `pivotCache` inside `<extLst>` (the x14 list) is ignored.
    /// The x14 `slicerCaches` and x15 `timelineCacheRefs` lists are neither
    /// read nor written, and the workbook `<extLst>` is not kept, so it is
    /// absent after a round trip.
    pub pivot_caches: Vec<PivotCacheRef>,
}

/// One `<pivotCache>` of `<pivotCaches>` (`CT_PivotCache`).
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PivotCacheRef {
    /// `cacheId`, the id pivot tables name in their own `cacheId`.
    pub cache_id: u32,
    /// `r:id`, the workbook relationship to the pivot cache definition.
    pub r_id: String,
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

/// A defined name (named range or formula alias), `CT_DefinedName`
/// (ECMA-376 §18.2.5). The element text is `value`; every other attribute is
/// optional and absent means its schema default (`false` for the flags).
#[derive(Debug, Clone, Default, Serialize, Deserialize, PartialEq)]
#[serde(rename_all = "camelCase")]
pub struct DefinedName {
    pub name: String,
    pub value: String,
    /// `localSheetId`: the 0-based sheet position the name is scoped to.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub sheet_id: Option<u32>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub comment: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub custom_menu: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub description: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub help: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub status_bar: Option<String>,
    /// Hidden from the Name Manager (for example `_xlnm._FilterDatabase`).
    #[serde(default, skip_serializing_if = "crate::ooxml::is_false")]
    pub hidden: bool,
    /// The name is a function (macro or user-defined).
    #[serde(default, skip_serializing_if = "crate::ooxml::is_false")]
    pub function: bool,
    /// The name is a Visual Basic procedure.
    #[serde(default, skip_serializing_if = "crate::ooxml::is_false")]
    pub vb_procedure: bool,
    /// The name is an XLM macro.
    #[serde(default, skip_serializing_if = "crate::ooxml::is_false")]
    pub xlm: bool,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub function_group_id: Option<u32>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub shortcut_key: Option<String>,
    #[serde(default, skip_serializing_if = "crate::ooxml::is_false")]
    pub publish_to_server: bool,
    #[serde(default, skip_serializing_if = "crate::ooxml::is_false")]
    pub workbook_parameter: bool,
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

/// An optional xsd:boolean attribute: `None` when the value is not a valid
/// boolean, so the schema default applies.
fn opt_bool(a: &Attribute<'_>) -> Result<Option<bool>> {
    Ok(xsd_bool(&attr_str(a)?))
}

/// An optional string attribute, entity-decoded.
fn opt_string(a: &Attribute<'_>) -> Result<Option<String>> {
    Ok(Some(attr_str(a)?.into_owned()))
}

/// An optional boolean attribute written as `1` or `0`.
fn push_bool(elem: &mut BytesStart<'_>, key: &'static str, v: Option<bool>) {
    if let Some(b) = v {
        elem.push_attribute((key, if b { "1" } else { "0" }));
    }
}

/// An optional string attribute, with characters illegal in XML removed.
fn push_str(elem: &mut BytesStart<'_>, key: &str, v: Option<&String>) {
    if let Some(v) = v {
        elem.push_attribute(xml_attr(key, v));
    }
}

/// An optional `u32` attribute.
fn push_u32(elem: &mut BytesStart<'_>, ibuf: &mut itoa::Buffer, key: &'static str, v: Option<u32>) {
    if let Some(n) = v {
        elem.push_attribute((key, ibuf.format(n)));
    }
}

/// The memory a kept defined name takes: its [`DefinedName`] plus the
/// capacity of its strings.
fn defined_name_memory(dn: &DefinedName) -> usize {
    size_of::<DefinedName>()
        + dn.name.capacity()
        + dn.value.capacity()
        + [
            &dn.comment,
            &dn.custom_menu,
            &dn.description,
            &dn.help,
            &dn.status_bar,
            &dn.shortcut_key,
        ]
        .into_iter()
        .flatten()
        .map(String::capacity)
        .sum::<usize>()
}

impl WorkbookXml {
    /// Parse `xl/workbook.xml` from raw XML bytes.
    ///
    /// # Errors
    ///
    /// `XmlParse` for malformed XML, a bad attribute value or a bad entity.
    pub fn parse(data: &[u8]) -> Result<Self> {
        Self::parse_within(data, &mut ExpansionBudget::new(usize::MAX))
    }

    /// [`Self::parse`], counting in `budget` (with [`ExpansionBudget::store`])
    /// each defined name before it is kept, at the size of its
    /// [`DefinedName`] plus the capacity of its strings.
    ///
    /// # Errors
    ///
    /// As [`Self::parse`], and `SizeLimit` once the count passes
    /// `budget.store_limit`.
    pub(crate) fn parse_within(data: &[u8], budget: &mut ExpansionBudget) -> Result<Self> {
        let mut reader = Reader::from_reader(data);
        // definedName text is a formula: its surrounding white space is part
        // of the value, and no other element's text is read.
        reader.config_mut().trim_text(false);

        let mut buf = Vec::with_capacity(512);
        let mut wb = WorkbookXml::default();

        // State for collecting definedName text content.
        let mut current_dn: Option<DefinedName> = None;
        // `<extLst>` nesting depth: its x14/x15 lists reuse the `pivotCache`
        // name and are not the `CT_Workbook` lists.
        let mut ext_lst_depth: u32 = 0;

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == "extLst" => {
                    ext_lst_depth += 1;
                }
                Ok(Event::End(ref e)) if e.local_name().as_ref() == "extLst" => {
                    ext_lst_depth = ext_lst_depth.saturating_sub(1);
                }
                Ok(Event::Empty(ref e) | Event::Start(ref e))
                    if ext_lst_depth == 0 && e.local_name().as_ref() == "pivotCache" =>
                {
                    if let Some(cache) = parse_pivot_cache_element(e)? {
                        wb.pivot_caches.push(cache);
                    }
                }
                Ok(Event::Empty(ref e) | Event::Start(ref e))
                    if ext_lst_depth == 0 && e.local_name().as_ref() == "externalReference" =>
                {
                    if let Some(r_id) = r_id_attr(e)? {
                        wb.external_references.push(r_id);
                    }
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == "definedName" => {
                    current_dn = Some(parse_defined_name_element(e)?);
                }
                Ok(Event::Empty(ref e)) if e.local_name().as_ref() == "definedName" => {
                    let dn = parse_defined_name_element(e)?;
                    budget.store(defined_name_memory(&dn))?;
                    wb.defined_names.push(dn);
                }
                Ok(Event::Empty(ref e) | Event::Start(ref e)) => match e.local_name().as_ref() {
                    "workbookPr" => parse_workbook_pr_element(e, &mut wb)?,
                    "fileVersion" => wb.file_version = Some(parse_file_version_element(e)?),
                    "calcPr" => wb.calc_pr = Some(parse_calc_pr_element(e)?),
                    "sheet" => wb.sheets.push(parse_sheet_element(e)?),
                    "workbookView" => wb.workbook_views.push(parse_workbook_view_element(e)?),
                    "workbookProtection" => {
                        wb.protection = Some(parse_workbook_protection_element(e)?);
                    }
                    _ => {}
                },
                Ok(Event::Text(ref e)) if current_dn.is_some() => {
                    if let Some(dn) = current_dn.as_mut() {
                        dn.value.push_str(&e.xml10_content());
                    }
                }
                Ok(Event::CData(ref e)) if current_dn.is_some() => {
                    if let Some(dn) = current_dn.as_mut() {
                        dn.value.push_str(&e.xml10_content());
                    }
                }
                Ok(Event::GeneralRef(ref e)) if current_dn.is_some() => {
                    if let Some(dn) = current_dn.as_mut() {
                        push_general_ref(&mut dn.value, e)?;
                    }
                }
                Ok(Event::End(ref e)) if e.local_name().as_ref() == "definedName" => {
                    if let Some(dn) = current_dn.take() {
                        budget.store(defined_name_memory(&dn))?;
                        wb.defined_names.push(dn);
                    }
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

        Ok(wb)
    }

    /// Serialize to valid `xl/workbook.xml` bytes.
    ///
    /// Children follow the `CT_Workbook` sequence: `fileVersion`,
    /// `workbookPr`, `workbookProtection`, `bookViews`, `sheets`,
    /// `externalReferences`, `definedNames`, `calcPr`, `pivotCaches`.
    ///
    /// # Errors
    ///
    /// `XmlWrite` when serialization fails.
    pub fn to_xml(&self) -> Result<Vec<u8>> {
        let mut buf: Vec<u8> = Vec::with_capacity(512);
        let mut writer = Writer::new(&mut buf);
        let mut ibuf = itoa::Buffer::new();
        let mut xbuf = XsdBuf::new();

        write_xml_decl(&mut writer)?;

        // <workbook xmlns="..." xmlns:r="...">
        let mut wb = BytesStart::new("workbook");
        wb.push_attribute(("xmlns", SPREADSHEET_NS));
        wb.push_attribute(("xmlns:r", REL_NS));
        writer
            .write_event(Event::Start(wb))
            .map_err(xml_write_err)?;

        // <fileVersion />
        if let Some(fv) = &self.file_version {
            let mut elem = BytesStart::new("fileVersion");
            push_str(&mut elem, "appName", fv.app_name.as_ref());
            push_str(&mut elem, "lastEdited", fv.last_edited.as_ref());
            push_str(&mut elem, "lowestEdited", fv.lowest_edited.as_ref());
            push_str(&mut elem, "rupBuild", fv.rup_build.as_ref());
            push_str(&mut elem, "codeName", fv.code_name.as_ref());
            writer
                .write_event(Event::Empty(elem))
                .map_err(xml_write_err)?;
        }

        // <workbookPr /> — when date1904 or any other property is set.
        let date1904 = self.date_system == DateSystem::Date1904;
        let pr = self
            .workbook_pr
            .as_ref()
            .filter(|p| **p != WorkbookPr::default());
        if date1904 || pr.is_some() {
            let mut elem = BytesStart::new("workbookPr");
            if date1904 {
                elem.push_attribute(("date1904", "1"));
            }
            if let Some(p) = pr {
                push_str(&mut elem, "showObjects", p.show_objects.as_ref());
                push_bool(
                    &mut elem,
                    "showBorderUnselectedTables",
                    p.show_border_unselected_tables,
                );
                push_bool(&mut elem, "filterPrivacy", p.filter_privacy);
                push_bool(&mut elem, "promptedSolutions", p.prompted_solutions);
                push_bool(&mut elem, "showInkAnnotation", p.show_ink_annotation);
                push_bool(&mut elem, "backupFile", p.backup_file);
                push_bool(
                    &mut elem,
                    "saveExternalLinkValues",
                    p.save_external_link_values,
                );
                push_str(&mut elem, "updateLinks", p.update_links.as_ref());
                push_str(&mut elem, "codeName", p.code_name.as_ref());
                push_bool(&mut elem, "hidePivotFieldList", p.hide_pivot_field_list);
                push_bool(&mut elem, "showPivotChartFilter", p.show_pivot_chart_filter);
                push_bool(&mut elem, "allowRefreshQuery", p.allow_refresh_query);
                push_bool(&mut elem, "publishItems", p.publish_items);
                push_bool(&mut elem, "checkCompatibility", p.check_compatibility);
                push_bool(&mut elem, "autoCompressPictures", p.auto_compress_pictures);
                push_bool(
                    &mut elem,
                    "refreshAllConnections",
                    p.refresh_all_connections,
                );
                push_u32(
                    &mut elem,
                    &mut ibuf,
                    "defaultThemeVersion",
                    p.default_theme_version,
                );
            }
            writer
                .write_event(Event::Empty(elem))
                .map_err(xml_write_err)?;
        }

        // <workbookProtection>
        if let Some(prot) = &self.protection {
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
            push_str(
                &mut elem,
                "workbookAlgorithmName",
                prot.workbook_algorithm_name.as_ref(),
            );
            push_str(
                &mut elem,
                "workbookHashValue",
                prot.workbook_hash_value.as_ref(),
            );
            push_str(
                &mut elem,
                "workbookSaltValue",
                prot.workbook_salt_value.as_ref(),
            );
            push_u32(
                &mut elem,
                &mut ibuf,
                "workbookSpinCount",
                prot.workbook_spin_count,
            );
            push_str(
                &mut elem,
                "revisionsAlgorithmName",
                prot.revisions_algorithm_name.as_ref(),
            );
            push_str(
                &mut elem,
                "revisionsHashValue",
                prot.revisions_hash_value.as_ref(),
            );
            push_str(
                &mut elem,
                "revisionsSaltValue",
                prot.revisions_salt_value.as_ref(),
            );
            push_u32(
                &mut elem,
                &mut ibuf,
                "revisionsSpinCount",
                prot.revisions_spin_count,
            );
            push_str(
                &mut elem,
                "workbookPassword",
                prot.workbook_password.as_ref(),
            );
            push_str(
                &mut elem,
                "revisionsPassword",
                prot.revisions_password.as_ref(),
            );
            writer
                .write_event(Event::Empty(elem))
                .map_err(xml_write_err)?;
        }

        // <bookViews> (only if non-empty)
        if !self.workbook_views.is_empty() {
            writer
                .write_event(Event::Start(BytesStart::new("bookViews")))
                .map_err(xml_write_err)?;

            for view in &self.workbook_views {
                let mut elem = BytesStart::new("workbookView");
                elem.push_attribute(("activeTab", ibuf.format(view.active_tab)));
                elem.push_attribute(("firstSheet", ibuf.format(view.first_sheet)));
                elem.push_attribute((
                    "showHorizontalScroll",
                    if view.show_horizontal_scroll {
                        "1"
                    } else {
                        "0"
                    },
                ));
                elem.push_attribute((
                    "showVerticalScroll",
                    if view.show_vertical_scroll { "1" } else { "0" },
                ));
                elem.push_attribute((
                    "showSheetTabs",
                    if view.show_sheet_tabs { "1" } else { "0" },
                ));
                push_u32(&mut elem, &mut ibuf, "windowWidth", view.window_width);
                push_u32(&mut elem, &mut ibuf, "windowHeight", view.window_height);
                push_u32(&mut elem, &mut ibuf, "tabRatio", view.tab_ratio);
                writer
                    .write_event(Event::Empty(elem))
                    .map_err(xml_write_err)?;
            }

            writer
                .write_event(Event::End(BytesEnd::new("bookViews")))
                .map_err(xml_write_err)?;
        }

        // <sheets>
        writer
            .write_event(Event::Start(BytesStart::new("sheets")))
            .map_err(xml_write_err)?;

        for sheet in &self.sheets {
            let mut elem = BytesStart::new("sheet");
            elem.push_attribute(xml_attr("name", &sheet.name));
            elem.push_attribute(("sheetId", ibuf.format(sheet.sheet_id)));
            elem.push_attribute(xml_attr("r:id", &sheet.r_id));
            match sheet.state {
                SheetState::Visible => {} // default, omit
                SheetState::Hidden => {
                    elem.push_attribute(("state", "hidden"));
                }
                SheetState::VeryHidden => {
                    elem.push_attribute(("state", "veryHidden"));
                }
            }
            writer
                .write_event(Event::Empty(elem))
                .map_err(xml_write_err)?;
        }

        // </sheets>
        writer
            .write_event(Event::End(BytesEnd::new("sheets")))
            .map_err(xml_write_err)?;

        // <externalReferences> (only if non-empty)
        if !self.external_references.is_empty() {
            writer
                .write_event(Event::Start(BytesStart::new("externalReferences")))
                .map_err(xml_write_err)?;
            for r_id in &self.external_references {
                let mut elem = BytesStart::new("externalReference");
                elem.push_attribute(xml_attr("r:id", r_id));
                writer
                    .write_event(Event::Empty(elem))
                    .map_err(xml_write_err)?;
            }
            writer
                .write_event(Event::End(BytesEnd::new("externalReferences")))
                .map_err(xml_write_err)?;
        }

        // <definedNames> (only if non-empty)
        if !self.defined_names.is_empty() {
            writer
                .write_event(Event::Start(BytesStart::new("definedNames")))
                .map_err(xml_write_err)?;

            for dn in &self.defined_names {
                // CT_DefinedName attribute order.
                let mut elem = BytesStart::new("definedName");
                elem.push_attribute(xml_attr("name", &dn.name));
                push_str(&mut elem, "comment", dn.comment.as_ref());
                push_str(&mut elem, "customMenu", dn.custom_menu.as_ref());
                push_str(&mut elem, "description", dn.description.as_ref());
                push_str(&mut elem, "help", dn.help.as_ref());
                push_str(&mut elem, "statusBar", dn.status_bar.as_ref());
                push_u32(&mut elem, &mut ibuf, "localSheetId", dn.sheet_id);
                for (key, set) in [
                    ("hidden", dn.hidden),
                    ("function", dn.function),
                    ("vbProcedure", dn.vb_procedure),
                    ("xlm", dn.xlm),
                ] {
                    if set {
                        elem.push_attribute((key, "1"));
                    }
                }
                push_u32(
                    &mut elem,
                    &mut ibuf,
                    "functionGroupId",
                    dn.function_group_id,
                );
                push_str(&mut elem, "shortcutKey", dn.shortcut_key.as_ref());
                for (key, set) in [
                    ("publishToServer", dn.publish_to_server),
                    ("workbookParameter", dn.workbook_parameter),
                ] {
                    if set {
                        elem.push_attribute((key, "1"));
                    }
                }
                writer
                    .write_event(Event::Start(elem))
                    .map_err(xml_write_err)?;
                writer
                    .write_event(Event::Text(BytesText::new(
                        crate::ooxml::sanitize_xml_text(&dn.value).as_ref(),
                    )))
                    .map_err(xml_write_err)?;
                writer
                    .write_event(Event::End(BytesEnd::new("definedName")))
                    .map_err(xml_write_err)?;
            }

            writer
                .write_event(Event::End(BytesEnd::new("definedNames")))
                .map_err(xml_write_err)?;
        }

        // <calcPr /> — right after </definedNames>.
        if let Some(cp) = &self.calc_pr {
            let mut elem = BytesStart::new("calcPr");
            push_u32(&mut elem, &mut ibuf, "calcId", cp.calc_id);
            push_str(&mut elem, "calcMode", cp.calc_mode.as_ref());
            push_bool(&mut elem, "fullCalcOnLoad", cp.full_calc_on_load);
            push_str(&mut elem, "refMode", cp.ref_mode.as_ref());
            push_bool(&mut elem, "iterate", cp.iterate);
            push_u32(&mut elem, &mut ibuf, "iterateCount", cp.iterate_count);
            if let Some(d) = cp.iterate_delta {
                let text = fmt_xsd_double(d, &mut xbuf, "calcPr @iterateDelta")?;
                elem.push_attribute(("iterateDelta", text));
            }
            push_bool(&mut elem, "fullPrecision", cp.full_precision);
            push_bool(&mut elem, "calcCompleted", cp.calc_completed);
            push_bool(&mut elem, "calcOnSave", cp.calc_on_save);
            push_bool(&mut elem, "concurrentCalc", cp.concurrent_calc);
            push_u32(
                &mut elem,
                &mut ibuf,
                "concurrentManualCount",
                cp.concurrent_manual_count,
            );
            push_bool(&mut elem, "forceFullCalc", cp.force_full_calc);
            writer
                .write_event(Event::Empty(elem))
                .map_err(xml_write_err)?;
        }

        // <pivotCaches> (only if non-empty) — after calcPr (and the unmodelled
        // oleSize and customWorkbookViews).
        if !self.pivot_caches.is_empty() {
            writer
                .write_event(Event::Start(BytesStart::new("pivotCaches")))
                .map_err(xml_write_err)?;
            for cache in &self.pivot_caches {
                let mut elem = BytesStart::new("pivotCache");
                elem.push_attribute(("cacheId", ibuf.format(cache.cache_id)));
                elem.push_attribute(xml_attr("r:id", &cache.r_id));
                writer
                    .write_event(Event::Empty(elem))
                    .map_err(xml_write_err)?;
            }
            writer
                .write_event(Event::End(BytesEnd::new("pivotCaches")))
                .map_err(xml_write_err)?;
        }
        // No <extLst> is written: the x14 slicerCaches and x15
        // timelineCacheRefs lists are not modelled.

        // </workbook>
        writer
            .write_event(Event::End(BytesEnd::new("workbook")))
            .map_err(xml_write_err)?;

        Ok(buf)
    }
}

/// Parse a `<definedName>` start tag (the text is collected by the caller).
fn parse_defined_name_element(e: &BytesStart<'_>) -> Result<DefinedName> {
    let mut dn = DefinedName::default();
    for attr in attrs(e) {
        let attr = attr?;
        match attr.key.local_name().as_ref() {
            "name" => dn.name = attr_str(&attr)?.into_owned(),
            "comment" => dn.comment = opt_string(&attr)?,
            "customMenu" => dn.custom_menu = opt_string(&attr)?,
            "description" => dn.description = opt_string(&attr)?,
            "help" => dn.help = opt_string(&attr)?,
            "statusBar" => dn.status_bar = opt_string(&attr)?,
            "localSheetId" => dn.sheet_id = Some(parse_attr("definedName", &attr)?),
            "hidden" => dn.hidden = opt_bool(&attr)?.unwrap_or(false),
            "function" => dn.function = opt_bool(&attr)?.unwrap_or(false),
            "vbProcedure" => dn.vb_procedure = opt_bool(&attr)?.unwrap_or(false),
            "xlm" => dn.xlm = opt_bool(&attr)?.unwrap_or(false),
            "functionGroupId" => dn.function_group_id = Some(parse_attr("definedName", &attr)?),
            "shortcutKey" => dn.shortcut_key = opt_string(&attr)?,
            "publishToServer" => dn.publish_to_server = opt_bool(&attr)?.unwrap_or(false),
            "workbookParameter" => dn.workbook_parameter = opt_bool(&attr)?.unwrap_or(false),
            _ => {}
        }
    }
    Ok(dn)
}

/// Read `<workbookPr>`: `date1904` sets the date system, the rest is kept in
/// [`WorkbookXml::workbook_pr`].
fn parse_workbook_pr_element(e: &BytesStart<'_>, wb: &mut WorkbookXml) -> Result<()> {
    let mut pr = WorkbookPr::default();
    for attr in attrs(e) {
        let attr = attr?;
        match attr.key.local_name().as_ref() {
            "date1904" => {
                if opt_bool(&attr)?.unwrap_or(false) {
                    wb.date_system = DateSystem::Date1904;
                }
            }
            "showObjects" => pr.show_objects = opt_string(&attr)?,
            "showBorderUnselectedTables" => pr.show_border_unselected_tables = opt_bool(&attr)?,
            "filterPrivacy" => pr.filter_privacy = opt_bool(&attr)?,
            "promptedSolutions" => pr.prompted_solutions = opt_bool(&attr)?,
            "showInkAnnotation" => pr.show_ink_annotation = opt_bool(&attr)?,
            "backupFile" => pr.backup_file = opt_bool(&attr)?,
            "saveExternalLinkValues" => pr.save_external_link_values = opt_bool(&attr)?,
            "updateLinks" => pr.update_links = opt_string(&attr)?,
            "codeName" => pr.code_name = opt_string(&attr)?,
            "hidePivotFieldList" => pr.hide_pivot_field_list = opt_bool(&attr)?,
            "showPivotChartFilter" => pr.show_pivot_chart_filter = opt_bool(&attr)?,
            "allowRefreshQuery" => pr.allow_refresh_query = opt_bool(&attr)?,
            "publishItems" => pr.publish_items = opt_bool(&attr)?,
            "checkCompatibility" => pr.check_compatibility = opt_bool(&attr)?,
            "autoCompressPictures" => pr.auto_compress_pictures = opt_bool(&attr)?,
            "refreshAllConnections" => pr.refresh_all_connections = opt_bool(&attr)?,
            "defaultThemeVersion" => {
                pr.default_theme_version = Some(parse_attr("workbookPr", &attr)?);
            }
            _ => {}
        }
    }
    if pr != WorkbookPr::default() {
        wb.workbook_pr = Some(pr);
    }
    Ok(())
}

/// Parse a `<fileVersion>` element.
fn parse_file_version_element(e: &BytesStart<'_>) -> Result<FileVersion> {
    let mut fv = FileVersion::default();
    for attr in attrs(e) {
        let attr = attr?;
        match attr.key.local_name().as_ref() {
            "appName" => fv.app_name = opt_string(&attr)?,
            "lastEdited" => fv.last_edited = opt_string(&attr)?,
            "lowestEdited" => fv.lowest_edited = opt_string(&attr)?,
            "rupBuild" => fv.rup_build = opt_string(&attr)?,
            "codeName" => fv.code_name = opt_string(&attr)?,
            _ => {}
        }
    }
    Ok(fv)
}

/// Parse a `<calcPr>` element.
fn parse_calc_pr_element(e: &BytesStart<'_>) -> Result<CalcPr> {
    let mut cp = CalcPr::default();
    for attr in attrs(e) {
        let attr = attr?;
        match attr.key.local_name().as_ref() {
            "calcId" => cp.calc_id = Some(parse_attr("calcPr", &attr)?),
            "calcMode" => cp.calc_mode = opt_string(&attr)?,
            "fullCalcOnLoad" => cp.full_calc_on_load = opt_bool(&attr)?,
            "refMode" => cp.ref_mode = opt_string(&attr)?,
            "iterate" => cp.iterate = opt_bool(&attr)?,
            "iterateCount" => cp.iterate_count = Some(parse_attr("calcPr", &attr)?),
            "iterateDelta" => cp.iterate_delta = Some(parse_attr("calcPr", &attr)?),
            "fullPrecision" => cp.full_precision = opt_bool(&attr)?,
            "calcCompleted" => cp.calc_completed = opt_bool(&attr)?,
            "calcOnSave" => cp.calc_on_save = opt_bool(&attr)?,
            "concurrentCalc" => cp.concurrent_calc = opt_bool(&attr)?,
            "concurrentManualCount" => {
                cp.concurrent_manual_count = Some(parse_attr("calcPr", &attr)?);
            }
            "forceFullCalc" => cp.force_full_calc = opt_bool(&attr)?,
            _ => {}
        }
    }
    Ok(cp)
}

/// Parse a `<workbookProtection>` element (either Empty or Start) into a `WorkbookProtection`.
fn parse_workbook_protection_element(e: &BytesStart<'_>) -> Result<WorkbookProtection> {
    let mut prot = WorkbookProtection::default();
    for attr in attrs(e) {
        let attr = attr?;
        match attr.key.local_name().as_ref() {
            "lockStructure" => prot.lock_structure = opt_bool(&attr)?.unwrap_or(false),
            "lockWindows" => prot.lock_windows = opt_bool(&attr)?.unwrap_or(false),
            "lockRevision" => prot.lock_revision = opt_bool(&attr)?.unwrap_or(false),
            "workbookAlgorithmName" => prot.workbook_algorithm_name = opt_string(&attr)?,
            "workbookHashValue" => prot.workbook_hash_value = opt_string(&attr)?,
            "workbookSaltValue" => prot.workbook_salt_value = opt_string(&attr)?,
            "workbookSpinCount" => {
                prot.workbook_spin_count = Some(parse_attr("workbookProtection", &attr)?);
            }
            "revisionsAlgorithmName" => prot.revisions_algorithm_name = opt_string(&attr)?,
            "revisionsHashValue" => prot.revisions_hash_value = opt_string(&attr)?,
            "revisionsSaltValue" => prot.revisions_salt_value = opt_string(&attr)?,
            "revisionsSpinCount" => {
                prot.revisions_spin_count = Some(parse_attr("workbookProtection", &attr)?);
            }
            "workbookPassword" => prot.workbook_password = opt_string(&attr)?,
            "revisionsPassword" => prot.revisions_password = opt_string(&attr)?,
            _ => {}
        }
    }
    Ok(prot)
}

/// Parse a `<workbookView>` element (either Empty or Start) into a `WorkbookView`.
fn parse_workbook_view_element(e: &BytesStart<'_>) -> Result<WorkbookView> {
    let mut view = WorkbookView::default();

    for attr in attrs(e) {
        let attr = attr?;
        match attr.key.local_name().as_ref() {
            "activeTab" => view.active_tab = parse_attr("workbookView", &attr)?,
            "firstSheet" => view.first_sheet = parse_attr("workbookView", &attr)?,
            "showHorizontalScroll" => {
                view.show_horizontal_scroll = opt_bool(&attr)?.unwrap_or(true);
            }
            "showVerticalScroll" => view.show_vertical_scroll = opt_bool(&attr)?.unwrap_or(true),
            "showSheetTabs" => view.show_sheet_tabs = opt_bool(&attr)?.unwrap_or(true),
            "windowWidth" => view.window_width = Some(parse_attr("workbookView", &attr)?),
            "windowHeight" => view.window_height = Some(parse_attr("workbookView", &attr)?),
            "tabRatio" => view.tab_ratio = Some(parse_attr("workbookView", &attr)?),
            _ => {}
        }
    }

    Ok(view)
}

/// The relationship id of `e`: its `r:id` or, when the producer binds the
/// namespace to another prefix, any attribute with local name `id`.
fn r_id_attr(e: &BytesStart<'_>) -> Result<Option<String>> {
    for attr in attrs(e) {
        let attr = attr?;
        if attr.key.as_ref() == "r:id" || attr.key.local_name().as_ref() == "id" {
            return Ok(Some(attr_str(&attr)?.into_owned()));
        }
    }
    Ok(None)
}

/// Parse a `<pivotCache>` element; `None` when `cacheId` or `r:id` is missing.
fn parse_pivot_cache_element(e: &BytesStart<'_>) -> Result<Option<PivotCacheRef>> {
    let mut cache_id = None;
    for attr in attrs(e) {
        let attr = attr?;
        if attr.key.local_name().as_ref() == "cacheId" {
            cache_id = Some(parse_attr("pivotCache", &attr)?);
        }
    }
    Ok(cache_id
        .zip(r_id_attr(e)?)
        .map(|(cache_id, r_id)| PivotCacheRef { cache_id, r_id }))
}

/// Parse a `<sheet>` element (either Empty or Start) into a `SheetInfo`.
fn parse_sheet_element(e: &BytesStart<'_>) -> Result<SheetInfo> {
    let mut name = String::new();
    let mut sheet_id: u32 = 0;
    let mut r_id = String::new();
    let mut state = SheetState::Visible;

    for attr in attrs(e) {
        let attr = attr?;
        let key = attr.key.as_ref();
        let local_name = attr.key.local_name();
        let local = local_name.as_ref();
        match local {
            "name" => name = attr_str(&attr)?.into_owned(),
            "sheetId" => sheet_id = parse_attr("sheet", &attr)?,
            "state" => {
                state = match attr_str(&attr)?.as_ref() {
                    "hidden" => SheetState::Hidden,
                    "veryHidden" => SheetState::VeryHidden,
                    _ => SheetState::Visible,
                };
            }
            // The relationship id is `r:id` or, when the producer binds the
            // namespace to another prefix, any attribute with local name `id`.
            _ if key == "r:id" || local == "id" => r_id = attr_str(&attr)?.into_owned(),
            _ => {}
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
                ..Default::default()
            }],
            workbook_views: Vec::new(),
            protection: None,
            ..Default::default()
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
            ..Default::default()
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
            ..Default::default()
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
            ..Default::default()
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
            ..Default::default()
        };
        let xml = wb.to_xml().unwrap();
        let xml_str = std::str::from_utf8(&xml).unwrap();
        assert!(!xml_str.contains("workbookProtection"));
    }

    const WB_OPEN: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">"#;

    fn one_sheet() -> Vec<SheetInfo> {
        vec![SheetInfo {
            name: "Sheet1".to_owned(),
            sheet_id: 1,
            r_id: "rId1".to_owned(),
            state: SheetState::Visible,
        }]
    }

    #[test]
    fn defined_name_formula_with_entities_roundtrips() {
        let wb = WorkbookXml {
            sheets: one_sheet(),
            defined_names: vec![DefinedName {
                name: "F".to_owned(),
                value: r#"IF(A1<1,"a & b",0)"#.to_owned(),
                ..Default::default()
            }],
            ..Default::default()
        };
        let xml = wb.to_xml().unwrap();
        let wb2 = WorkbookXml::parse(&xml).unwrap();
        assert_eq!(wb2.defined_names[0].value, r#"IF(A1<1,"a & b",0)"#);
        // Writing what was parsed is byte-stable.
        assert_eq!(xml, wb2.to_xml().unwrap());
    }

    #[test]
    fn defined_name_parses_entities_and_cdata() {
        let xml = format!(
            r#"{WB_OPEN}<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets><definedNames>
<definedName name="X"><![CDATA[Sheet1!$A$1]]></definedName>
<definedName name="Y">IF(A1&lt;1,&quot;a &amp; b&quot;,&#x41;)</definedName>
<definedName name="Z"> Sheet1!$B$2 </definedName>
</definedNames></workbook>"#
        );
        let wb = WorkbookXml::parse(xml.as_bytes()).unwrap();
        assert_eq!(wb.defined_names[0].value, "Sheet1!$A$1");
        assert_eq!(wb.defined_names[1].value, r#"IF(A1<1,"a & b",A)"#);
        assert_eq!(wb.defined_names[2].value, " Sheet1!$B$2 ");
    }

    #[test]
    fn defined_name_hidden_and_attrs_roundtrip() {
        let xml = format!(
            r#"{WB_OPEN}<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets><definedNames>
<definedName name="_xlnm._FilterDatabase" localSheetId="0" hidden="1">Sheet1!$A$1:$C$9</definedName>
<definedName name="M" comment="c &amp; d" customMenu="cm" description="de" help="h" statusBar="sb" function="true" vbProcedure="1" xlm="1" functionGroupId="3" shortcutKey="k" publishToServer="1" workbookParameter="1">Sheet1!$A$1</definedName>
</definedNames></workbook>"#
        );
        let wb = WorkbookXml::parse(xml.as_bytes()).unwrap();
        assert!(wb.defined_names[0].hidden);
        assert_eq!(wb.defined_names[0].sheet_id, Some(0));
        let m = &wb.defined_names[1];
        assert_eq!(m.comment.as_deref(), Some("c & d"));
        assert!(m.function && m.vb_procedure && m.xlm && m.publish_to_server);
        assert!(m.workbook_parameter && !m.hidden);
        assert_eq!(m.function_group_id, Some(3));

        let out = String::from_utf8(wb.to_xml().unwrap()).unwrap();
        assert!(out.contains(r#"hidden="1""#), "{out}");
        let wb2 = WorkbookXml::parse(out.as_bytes()).unwrap();
        assert_eq!(wb2.defined_names, wb.defined_names);
    }

    #[test]
    fn sheet_name_with_ampersand_survives_two_roundtrips() {
        let mut wb = WorkbookXml {
            sheets: one_sheet(),
            ..Default::default()
        };
        wb.sheets[0].name = "R&D".to_owned();
        let once = WorkbookXml::parse(&wb.to_xml().unwrap()).unwrap();
        let twice = WorkbookXml::parse(&once.to_xml().unwrap()).unwrap();
        assert_eq!(once.sheets[0].name, "R&D");
        assert_eq!(twice.sheets[0].name, "R&D");
    }

    #[test]
    fn calc_pr_file_version_workbook_pr_roundtrip_in_order() {
        let xml = format!(
            r#"{WB_OPEN}<fileVersion appName="xl" lastEdited="7" lowestEdited="7" rupBuild="22228" codeName="{{ABC}}"/><workbookPr date1904="1" codeName="ThisWorkbook" defaultThemeVersion="166925" filterPrivacy="true"/><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets><definedNames><definedName name="N">Sheet1!$A$1</definedName></definedNames><calcPr calcId="191029" calcMode="manual" fullCalcOnLoad="1" iterate="1" iterateCount="50" iterateDelta="0.01" fullPrecision="0"/></workbook>"#
        );
        let wb = WorkbookXml::parse(xml.as_bytes()).unwrap();
        assert_eq!(wb.date_system, DateSystem::Date1904);
        let pr = wb.workbook_pr.as_ref().unwrap();
        assert_eq!(pr.code_name.as_deref(), Some("ThisWorkbook"));
        assert_eq!(pr.default_theme_version, Some(166_925));
        assert_eq!(pr.filter_privacy, Some(true));
        let cp = wb.calc_pr.as_ref().unwrap();
        assert_eq!(cp.calc_mode.as_deref(), Some("manual"));
        assert_eq!(cp.full_calc_on_load, Some(true));
        assert_eq!(cp.full_precision, Some(false));
        assert_eq!(cp.iterate_delta, Some(0.01));
        assert_eq!(
            wb.file_version.as_ref().unwrap().rup_build.as_deref(),
            Some("22228")
        );

        let out = String::from_utf8(wb.to_xml().unwrap()).unwrap();
        let fv = out.find("<fileVersion").unwrap();
        let wpr = out.find("<workbookPr").unwrap();
        let sheets = out.find("<sheets>").unwrap();
        let dn_end = out.find("</definedNames>").unwrap();
        let calc = out.find("<calcPr").unwrap();
        assert!(
            fv < wpr && wpr < sheets && sheets < dn_end && dn_end < calc,
            "{out}"
        );
        assert!(out.contains(r#"fullCalcOnLoad="1""#));
        let wb2 = WorkbookXml::parse(out.as_bytes()).unwrap();
        assert_eq!(wb2.calc_pr, wb.calc_pr);
        assert_eq!(wb2.workbook_pr, wb.workbook_pr);
        assert_eq!(wb2.file_version, wb.file_version);
        assert_eq!(wb2.date_system, DateSystem::Date1904);
    }

    #[test]
    fn workbook_pr_without_date1904_is_written_only_when_set() {
        let plain = WorkbookXml {
            sheets: one_sheet(),
            ..Default::default()
        };
        assert!(
            !String::from_utf8(plain.to_xml().unwrap())
                .unwrap()
                .contains("workbookPr")
        );
        let coded = WorkbookXml {
            workbook_pr: Some(WorkbookPr {
                code_name: Some("ThisWorkbook".to_owned()),
                ..Default::default()
            }),
            ..plain
        };
        let out = String::from_utf8(coded.to_xml().unwrap()).unwrap();
        assert!(
            out.contains(r#"<workbookPr codeName="ThisWorkbook"/>"#),
            "{out}"
        );
        assert!(!out.contains("date1904"));
    }

    #[test]
    fn workbook_protection_booleans_use_xsd_lexical_forms() {
        let xml = format!(
            r#"{WB_OPEN}<workbookProtection lockStructure="true" lockWindows="0" lockRevision="false"/><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>"#
        );
        let prot = WorkbookXml::parse(xml.as_bytes())
            .unwrap()
            .protection
            .unwrap();
        assert!(prot.lock_structure);
        assert!(!prot.lock_windows);
        assert!(!prot.lock_revision);
    }

    #[test]
    fn workbook_protection_revisions_roundtrip() {
        let wb = WorkbookXml {
            sheets: one_sheet(),
            protection: Some(WorkbookProtection {
                lock_revision: true,
                revisions_algorithm_name: Some("SHA-512".to_owned()),
                revisions_hash_value: Some("h".to_owned()),
                revisions_salt_value: Some("s".to_owned()),
                revisions_spin_count: Some(100_000),
                revisions_password: Some("CC1B".to_owned()),
                workbook_password: Some("DEAD".to_owned()),
                ..Default::default()
            }),
            ..Default::default()
        };
        let wb2 = WorkbookXml::parse(&wb.to_xml().unwrap()).unwrap();
        assert_eq!(wb2.protection, wb.protection);
    }

    #[test]
    fn malformed_attribute_values_are_errors() {
        let xml = format!(
            r#"{WB_OPEN}<sheets><sheet name="S" sheetId="x" r:id="rId1"/></sheets></workbook>"#
        );
        assert!(matches!(
            WorkbookXml::parse(xml.as_bytes()),
            Err(ModernXlsxError::XmlParse(_))
        ));
    }

    #[test]
    fn pivot_caches_and_external_references_roundtrip_in_ct_workbook_order() {
        let xml = format!(
            r#"{WB_OPEN}<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets><externalReferences><externalReference r:id="rId5"/><externalReference r:id="rId4"/></externalReferences><definedNames><definedName name="N">[1]Sheet1!$A$1</definedName></definedNames><calcPr calcId="191029"/><pivotCaches><pivotCache cacheId="7" r:id="rId6"/><pivotCache cacheId="2" r:id="rId7"/></pivotCaches><extLst><ext uri="{{876F7934-8845-4945-9796-88D515C7AA90}}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"><x14:pivotCaches><pivotCache cacheId="9" r:id="rId8"/></x14:pivotCaches></ext></extLst></workbook>"#
        );
        let wb = WorkbookXml::parse(xml.as_bytes()).unwrap();
        assert_eq!(wb.external_references, ["rId5", "rId4"]);
        // The x14 extLst list is not the CT_Workbook <pivotCaches>.
        assert_eq!(
            wb.pivot_caches,
            [
                PivotCacheRef {
                    cache_id: 7,
                    r_id: "rId6".to_owned(),
                },
                PivotCacheRef {
                    cache_id: 2,
                    r_id: "rId7".to_owned(),
                },
            ]
        );

        let out = String::from_utf8(wb.to_xml().unwrap()).unwrap();
        assert!(
            out.contains(r#"<externalReferences><externalReference r:id="rId5"/><externalReference r:id="rId4"/></externalReferences>"#),
            "{out}"
        );
        assert!(
            out.contains(r#"<pivotCaches><pivotCache cacheId="7" r:id="rId6"/><pivotCache cacheId="2" r:id="rId7"/></pivotCaches>"#),
            "{out}"
        );
        let sheets_end = out.find("</sheets>").unwrap();
        let ext_refs = out.find("<externalReferences>").unwrap();
        let names = out.find("<definedNames>").unwrap();
        let calc = out.find("<calcPr").unwrap();
        let caches = out.find("<pivotCaches>").unwrap();
        assert!(
            sheets_end < ext_refs && ext_refs < names && names < calc && calc < caches,
            "{out}"
        );
        let wb2 = WorkbookXml::parse(out.as_bytes()).unwrap();
        assert_eq!(wb2.pivot_caches, wb.pivot_caches);
        assert_eq!(wb2.external_references, wb.external_references);
    }

    #[test]
    fn empty_pivot_caches_and_external_references_are_not_written() {
        let out = String::from_utf8(
            WorkbookXml {
                sheets: one_sheet(),
                ..Default::default()
            }
            .to_xml()
            .unwrap(),
        )
        .unwrap();
        assert!(!out.contains("pivotCaches"), "{out}");
        assert!(!out.contains("externalReferences"), "{out}");
    }

    #[test]
    fn mismatched_end_tag_is_xml_parse_error() {
        assert!(matches!(
            WorkbookXml::parse(br#"<root xmlns="x"><a></b></root>"#),
            Err(ModernXlsxError::XmlParse(_))
        ));
    }

    #[test]
    fn defined_name_line_ends_are_normalized_and_char_refs_kept() {
        let xml = format!(
            "{WB_OPEN}<sheets><sheet name=\"Sheet1\" sheetId=\"1\" r:id=\"rId1\"/></sheets><definedNames>\
             <definedName name=\"X\">A1&amp;\"\r\n\"</definedName>\
             <definedName name=\"Y\">A1&#13;&#10;</definedName>\
             <definedName name=\"Z\"><![CDATA[B1\r\n]]></definedName>\
             </definedNames></workbook>"
        );
        let wb = WorkbookXml::parse(xml.as_bytes()).unwrap();
        assert_eq!(wb.defined_names[0].value, "A1&\"\n\"");
        assert_eq!(wb.defined_names[1].value, "A1\r\n");
        assert_eq!(wb.defined_names[2].value, "B1\n");
    }
}
