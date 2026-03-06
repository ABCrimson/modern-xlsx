//! Full XLSX read orchestrator.
//!
//! Decompresses a `.xlsx` ZIP archive, parses all OPC / SpreadsheetML parts,
//! and assembles them into a [`WorkbookData`] struct.

use core::hint::cold_path;
use std::collections::{BTreeMap, HashSet};

use log::{debug, warn};

#[cfg(feature = "parallel")]
use rayon::prelude::*;

use crate::ooxml::{
    calc_chain,
    charts,
    comments,
    doc_props,
    pivot_table::{PivotCacheDefinitionData, PivotCacheRecordsData, PivotTableData},
    relationships::{
        Relationships, REL_COMMENTS, REL_DRAWING, REL_PIVOT_CACHE_DEF, REL_PIVOT_CACHE_REC,
        REL_PIVOT_TABLE, REL_SLICER, REL_SLICER_CACHE, REL_TABLE, REL_THREADED_COMMENTS,
        REL_TIMELINE, REL_TIMELINE_CACHE,
    },
    shared_strings::SharedStringTable,
    slicers::{self, SlicerCacheData, SlicerData},
    styles::Styles,
    tables::TableDefinition,
    theme,
    threaded_comments::{self, PersonData, ThreadedCommentData},
    timelines::{self, TimelineCacheData, TimelineData},
    workbook::{SheetState, WorkbookXml},
    worksheet::WorksheetXml,
};
use crate::ole2::detect::{classify_ole2, detect_format, FileFormat, Ole2Kind};
use crate::zip::reader::{read_zip_entries, ZipSecurityLimits};
use crate::{ModernXlsxError, Result, SheetData, WorkbookData};

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

/// Convert a [`SheetState`] to the JSON-bridge string representation.
///
/// `Visible` maps to `None` (omitted via `skip_serializing_if`),
/// while `Hidden` and `VeryHidden` map to their camelCase strings.
fn sheet_state_to_str(state: SheetState) -> Option<String> {
    match state {
        SheetState::Visible => None,
        SheetState::Hidden => Some("hidden".into()),
        SheetState::VeryHidden => Some("veryHidden".into()),
    }
}

// ---------------------------------------------------------------------------
// Shared reader context
// ---------------------------------------------------------------------------

/// Well-known ZIP entry paths that are parsed (not preserved verbatim).
const KNOWN_STATIC_PATHS: &[&str] = &[
    "[Content_Types].xml",
    "_rels/.rels",
    "xl/workbook.xml",
    "xl/sharedStrings.xml",
    "xl/styles.xml",
    "xl/_rels/workbook.xml.rels",
    "docProps/core.xml",
    "docProps/app.xml",
    "xl/calcChain.xml",
];

/// Common parsed data shared between the struct-based and JSON-streaming readers.
struct ReaderContext {
    entries: std::collections::HashMap<String, Vec<u8>>,
    workbook_xml: WorkbookXml,
    sst: SharedStringTable,
    styles: Styles,
    doc_properties: Option<crate::ooxml::doc_props::DocProperties>,
    theme_colors: Option<crate::ooxml::theme::ThemeColors>,
    calc_chain: Vec<crate::ooxml::calc_chain::CalcChainEntry>,
    /// Workbook-level relationships (needed for pivot cache resolution).
    wb_rels: Relationships,
    /// (sheet_name, zip_path, visibility_state) triples in workbook order.
    sheet_targets: Vec<(String, String, SheetState)>,
    /// Paths of parsed entries (used to compute preserved_entries).
    known_dynamic: HashSet<String>,
}

/// Parse all shared XLSX metadata from the ZIP entries.
///
/// When `password` is `Some`, encrypted OLE2 files are decrypted before parsing.
/// When `password` is `None`, encrypted files produce a descriptive error.
fn parse_common(data: &[u8], limits: &ZipSecurityLimits, password: Option<&str>) -> Result<ReaderContext> {
    use crate::ole2::detect::{ERR_LEGACY_XLS, ERR_NOT_XLSX, ERR_OLE2_UNKNOWN};

    // Format detection with optional decryption.
    let decrypted: Option<Vec<u8>>;
    let actual_data: &[u8] = match detect_format(data) {
        FileFormat::Zip => data,
        FileFormat::Ole2 => match classify_ole2(data)? {
            Ole2Kind::EncryptedXlsx => {
                if let Some(pw) = password {
                    decrypted = Some(crate::ole2::detect::decrypt_file(data, pw)?);
                    decrypted.as_deref().unwrap()
                } else {
                    cold_path();
                    return Err(crate::ole2::encryption_info::build_encrypted_error(data));
                }
            }
            Ole2Kind::LegacyXls => {
                cold_path();
                return Err(ModernXlsxError::LegacyFormat(ERR_LEGACY_XLS.into()));
            }
            Ole2Kind::Unknown => {
                cold_path();
                return Err(ModernXlsxError::UnrecognizedFormat(
                    ERR_OLE2_UNKNOWN.into(),
                ));
            }
        },
        FileFormat::Unknown => {
            cold_path();
            return Err(ModernXlsxError::UnrecognizedFormat(ERR_NOT_XLSX.into()));
        }
    };

    let entries = read_zip_entries(actual_data, limits)?;

    // Parse workbook (required part).
    debug!("parsing workbook.xml");
    let workbook_data = entries
        .get("xl/workbook.xml")
        .ok_or_else(|| ModernXlsxError::MissingPart("xl/workbook.xml".into()))?;
    let workbook_xml = WorkbookXml::parse(workbook_data)?;

    // Parse shared strings (optional — some workbooks have no strings).
    let sst = entries
        .get("xl/sharedStrings.xml")
        .map(|d| SharedStringTable::parse(d))
        .transpose()?
        .unwrap_or_else(|| {
            warn!("shared strings table not found");
            SharedStringTable::empty()
        });

    // Parse styles (optional — use defaults when absent).
    let styles = entries
        .get("xl/styles.xml")
        .map(|d| Styles::parse(d))
        .transpose()?
        .unwrap_or_else(Styles::default_styles);

    // Parse document properties (optional).
    let mut doc_properties = entries
        .get("docProps/core.xml")
        .map(|d| doc_props::parse_core(d))
        .transpose()?;
    if let Some(app_data) = entries.get("docProps/app.xml") {
        let props = doc_properties.get_or_insert_with(Default::default);
        doc_props::parse_app(props, app_data)?;
    }

    // Parse theme colors (optional — xl/theme/theme1.xml).
    let theme_colors = entries
        .get("xl/theme/theme1.xml")
        .map(|d| theme::parse(d))
        .transpose()?;

    // Parse calculation chain (optional — xl/calcChain.xml).
    let calc_chain = entries
        .get("xl/calcChain.xml")
        .map(|d| calc_chain::parse(d))
        .transpose()?
        .unwrap_or_default();

    // Parse workbook relationships (optional — needed to resolve sheet targets).
    let wb_rels = entries
        .get("xl/_rels/workbook.xml.rels")
        .map(|d| Relationships::parse(d))
        .transpose()?
        .unwrap_or_else(Relationships::new);

    // Resolve sheet paths from workbook relationships.
    // NOTE: xl/theme/theme1.xml is NOT added to known_paths — we still
    // preserve the full theme XML verbatim so it survives roundtrip.
    let mut known_dynamic: HashSet<String> = HashSet::new();
    let mut sheet_targets: Vec<(String, String, SheetState)> = Vec::new();

    for sheet_info in &workbook_xml.sheets {
        let rel = wb_rels.get_by_id(&sheet_info.r_id).ok_or_else(|| {
            warn!("could not resolve sheet target for rId: {}", sheet_info.r_id);
            ModernXlsxError::MissingPart(format!(
                "relationship {} for sheet '{}'",
                sheet_info.r_id, sheet_info.name
            ))
        })?;

        // Normalize target path: handle both relative ("worksheets/sheet1.xml")
        // and absolute ("/xl/worksheets/sheet1.xml") targets.
        let sheet_path = if rel.target.starts_with('/') {
            rel.target.trim_start_matches('/').to_string()
        } else {
            format!("xl/{}", rel.target)
        };

        known_dynamic.insert(sheet_path.clone());
        sheet_targets.push((sheet_info.name.clone(), sheet_path, sheet_info.state));
    }

    Ok(ReaderContext {
        entries,
        workbook_xml,
        sst,
        styles,
        doc_properties,
        theme_colors,
        calc_chain,
        wb_rels,
        sheet_targets,
        known_dynamic,
    })
}

/// Resolve comments for each sheet from worksheet .rels files.
fn resolve_comments(
    ctx: &mut ReaderContext,
) -> Result<Vec<Vec<comments::Comment>>> {
    let mut sheet_comments = vec![Vec::new(); ctx.sheet_targets.len()];

    for (i, (_name, sheet_path, _state)) in ctx.sheet_targets.iter().enumerate() {
        let rels_path = derive_rels_path(sheet_path);

        if let Some(rels_data) = ctx.entries.get(&rels_path) {
            let ws_rels = Relationships::parse(rels_data)?;

            for rel in ws_rels.find_by_type(REL_COMMENTS) {
                let comments_path = resolve_rel_target(sheet_path, &rel.target);

                if let Some(comments_data) = ctx.entries.get(&comments_path) {
                    debug!("parsing comments from: {}", comments_path);
                    sheet_comments[i] = comments::parse_comments(comments_data)?;
                    ctx.known_dynamic.insert(comments_path);
                } else {
                    warn!("comments file not found: {}", comments_path);
                }
            }
        }
    }

    Ok(sheet_comments)
}

/// Resolve table definitions for each sheet from worksheet .rels files.
fn resolve_tables(
    ctx: &mut ReaderContext,
) -> Result<Vec<Vec<TableDefinition>>> {
    let mut sheet_tables = vec![Vec::new(); ctx.sheet_targets.len()];

    for (i, (_name, sheet_path, _state)) in ctx.sheet_targets.iter().enumerate() {
        let rels_path = derive_rels_path(sheet_path);

        if let Some(rels_data) = ctx.entries.get(&rels_path) {
            let ws_rels = Relationships::parse(rels_data)?;

            for rel in ws_rels.find_by_type(REL_TABLE) {
                let table_path = resolve_rel_target(sheet_path, &rel.target);

                if let Some(table_data) = ctx.entries.get(&table_path) {
                    debug!("parsing table from: {}", table_path);
                    sheet_tables[i].push(TableDefinition::parse(table_data)?);
                    ctx.known_dynamic.insert(table_path);
                } else {
                    warn!("table file not found: {}", table_path);
                }
            }
        }
    }

    Ok(sheet_tables)
}

/// Resolve chart definitions for each sheet.
///
/// Charts require a two-step lookup: sheet .rels → drawing rel → drawing.xml
/// (which contains chart anchors) → drawing .rels → chart XML.
fn resolve_charts(
    ctx: &mut ReaderContext,
) -> Result<Vec<Vec<charts::WorksheetChart>>> {
    let mut sheet_charts = vec![Vec::new(); ctx.sheet_targets.len()];

    for (i, (_name, sheet_path, _state)) in ctx.sheet_targets.iter().enumerate() {
        let rels_path = derive_rels_path(sheet_path);

        if let Some(rels_data) = ctx.entries.get(&rels_path) {
            let ws_rels = Relationships::parse(rels_data)?;

            // Find drawing relationships.
            for rel in ws_rels.find_by_type(REL_DRAWING) {
                let drawing_path = resolve_rel_target(sheet_path, &rel.target);

                if let Some(drawing_data) = ctx.entries.get(&drawing_path) {
                    debug!("parsing drawing from: {}", drawing_path);
                    let anchors = charts::parse_drawing_anchors(drawing_data)?;

                    // Resolve each chart rId via drawing .rels.
                    let drawing_rels_path = derive_rels_path(&drawing_path);
                    let drawing_rels =
                        if let Some(dr) = ctx.entries.get(&drawing_rels_path) {
                            Relationships::parse(dr)?
                        } else {
                            continue;
                        };

                    for (anchor, chart_r_id) in &anchors {
                        if let Some(chart_rel) = drawing_rels
                            .relationships
                            .iter()
                            .find(|r| r.id == *chart_r_id)
                        {
                            let chart_path =
                                resolve_rel_target(&drawing_path, &chart_rel.target);

                            if let Some(chart_data) = ctx.entries.get(&chart_path) {
                                debug!("parsing chart from: {}", chart_path);
                                let chart = charts::ChartData::parse(chart_data)?;
                                sheet_charts[i].push(charts::WorksheetChart {
                                    chart,
                                    anchor: anchor.clone(),
                                });
                                ctx.known_dynamic.insert(chart_path);
                            } else {
                                warn!("chart file not found: {}", chart_path);
                            }
                        }
                    }

                    ctx.known_dynamic.insert(drawing_path);
                    ctx.known_dynamic.insert(drawing_rels_path);
                } else {
                    warn!("drawing file not found: {}", drawing_path);
                }
            }
        }
    }

    Ok(sheet_charts)
}

/// Resolve pivot table definitions for each sheet from worksheet .rels files.
fn resolve_pivot_tables(
    ctx: &mut ReaderContext,
) -> Result<Vec<Vec<PivotTableData>>> {
    let mut sheet_pivots = vec![Vec::new(); ctx.sheet_targets.len()];

    for (i, (_name, sheet_path, _state)) in ctx.sheet_targets.iter().enumerate() {
        let rels_path = derive_rels_path(sheet_path);

        if let Some(rels_data) = ctx.entries.get(&rels_path) {
            let ws_rels = Relationships::parse(rels_data)?;

            for rel in ws_rels.find_by_type(REL_PIVOT_TABLE) {
                let pivot_path = resolve_rel_target(sheet_path, &rel.target);

                if let Some(pivot_data) = ctx.entries.get(&pivot_path) {
                    debug!("parsing pivot table from: {}", pivot_path);
                    sheet_pivots[i].push(PivotTableData::parse(pivot_data)?);
                    ctx.known_dynamic.insert(pivot_path);
                } else {
                    warn!("pivot table file not found: {}", pivot_path);
                }
            }
        }
    }

    Ok(sheet_pivots)
}

/// Resolve pivot cache definitions and records from workbook .rels.
fn resolve_pivot_caches(
    ctx: &mut ReaderContext,
) -> Result<(Vec<PivotCacheDefinitionData>, Vec<PivotCacheRecordsData>)> {
    let mut cache_defs: Vec<PivotCacheDefinitionData> = Vec::new();
    let mut cache_recs: Vec<PivotCacheRecordsData> = Vec::new();

    // Collect targets first to avoid borrow conflict with ctx.entries.
    let cache_targets: Vec<String> = ctx
        .wb_rels
        .find_by_type(REL_PIVOT_CACHE_DEF)
        .map(|rel| {
            if rel.target.starts_with('/') {
                rel.target.trim_start_matches('/').to_string()
            } else {
                format!("xl/{}", rel.target)
            }
        })
        .collect();

    for cache_path in cache_targets {
        if let Some(cache_data) = ctx.entries.get(&cache_path) {
            debug!("parsing pivot cache definition from: {}", cache_path);
            cache_defs.push(PivotCacheDefinitionData::parse(cache_data)?);
            ctx.known_dynamic.insert(cache_path.clone());

            // Look for corresponding records via cache definition .rels.
            let cache_rels_path = derive_rels_path(&cache_path);
            if let Some(cr_data) = ctx.entries.get(&cache_rels_path) {
                let cache_rels = Relationships::parse(cr_data)?;

                for rel in cache_rels.find_by_type(REL_PIVOT_CACHE_REC) {
                    let records_path = resolve_rel_target(&cache_path, &rel.target);

                    if let Some(rec_data) = ctx.entries.get(&records_path) {
                        debug!("parsing pivot cache records from: {}", records_path);
                        cache_recs.push(PivotCacheRecordsData::parse(rec_data)?);
                        ctx.known_dynamic.insert(records_path);
                    } else {
                        warn!("pivot cache records file not found: {}", records_path);
                    }
                }

                ctx.known_dynamic.insert(cache_rels_path);
            }
        } else {
            warn!("pivot cache definition file not found: {}", cache_path);
        }
    }

    Ok((cache_defs, cache_recs))
}

/// Resolve threaded comments for each sheet from worksheet .rels files.
fn resolve_threaded_comments(
    ctx: &mut ReaderContext,
) -> Result<Vec<Vec<ThreadedCommentData>>> {
    let mut sheet_tc = vec![Vec::new(); ctx.sheet_targets.len()];

    for (i, (_name, sheet_path, _state)) in ctx.sheet_targets.iter().enumerate() {
        let rels_path = derive_rels_path(sheet_path);

        if let Some(rels_data) = ctx.entries.get(&rels_path) {
            let ws_rels = Relationships::parse(rels_data)?;

            for rel in ws_rels.find_by_type(REL_THREADED_COMMENTS) {
                let tc_path = resolve_rel_target(sheet_path, &rel.target);

                if let Some(tc_data) = ctx.entries.get(&tc_path) {
                    debug!("parsing threaded comments from: {}", tc_path);
                    sheet_tc[i] = threaded_comments::parse_threaded_comments(tc_data)?;
                    ctx.known_dynamic.insert(tc_path);
                } else {
                    warn!("threaded comments file not found: {}", tc_path);
                }
            }
        }
    }

    Ok(sheet_tc)
}

/// Resolve persons list from the well-known path `xl/persons/person.xml`.
fn resolve_persons(ctx: &mut ReaderContext) -> Result<Vec<PersonData>> {
    let persons_path = "xl/persons/person.xml";

    if let Some(persons_data) = ctx.entries.get(persons_path) {
        debug!("parsing persons from: {}", persons_path);
        let persons = threaded_comments::parse_persons(persons_data)?;
        ctx.known_dynamic.insert(persons_path.to_string());
        Ok(persons)
    } else {
        Ok(Vec::new())
    }
}

/// Resolve slicers per worksheet via REL_SLICER in each sheet's .rels.
fn resolve_slicers(ctx: &mut ReaderContext) -> Result<Vec<Vec<SlicerData>>> {
    let mut sheet_slicers = vec![Vec::new(); ctx.sheet_targets.len()];

    for (i, (_name, sheet_path, _state)) in ctx.sheet_targets.iter().enumerate() {
        let rels_path = derive_rels_path(sheet_path);

        if let Some(rels_data) = ctx.entries.get(&rels_path) {
            let ws_rels = Relationships::parse(rels_data)?;

            for rel in ws_rels.find_by_type(REL_SLICER) {
                let slicer_path = resolve_rel_target(sheet_path, &rel.target);

                if let Some(slicer_data) = ctx.entries.get(&slicer_path) {
                    debug!("parsing slicers from: {}", slicer_path);
                    sheet_slicers[i] = slicers::parse_slicers(slicer_data)?;
                    ctx.known_dynamic.insert(slicer_path);
                } else {
                    warn!("slicer file not found: {}", slicer_path);
                }
            }
        }
    }

    Ok(sheet_slicers)
}

/// Resolve slicer cache definitions from workbook .rels.
fn resolve_slicer_caches(ctx: &mut ReaderContext) -> Result<Vec<SlicerCacheData>> {
    let mut caches: Vec<SlicerCacheData> = Vec::new();

    let cache_targets: Vec<String> = ctx
        .wb_rels
        .find_by_type(REL_SLICER_CACHE)
        .map(|rel| {
            if rel.target.starts_with('/') {
                rel.target.trim_start_matches('/').to_string()
            } else {
                format!("xl/{}", rel.target)
            }
        })
        .collect();

    for cache_path in cache_targets {
        if let Some(cache_data) = ctx.entries.get(&cache_path) {
            debug!("parsing slicer cache from: {}", cache_path);
            caches.push(SlicerCacheData::parse(cache_data)?);
            ctx.known_dynamic.insert(cache_path);
        } else {
            warn!("slicer cache file not found: {}", cache_path);
        }
    }

    Ok(caches)
}

/// Resolve timelines per worksheet via REL_TIMELINE in each sheet's .rels.
fn resolve_timelines(ctx: &mut ReaderContext) -> Result<Vec<Vec<TimelineData>>> {
    let mut sheet_timelines = vec![Vec::new(); ctx.sheet_targets.len()];

    for (i, (_name, sheet_path, _state)) in ctx.sheet_targets.iter().enumerate() {
        let rels_path = derive_rels_path(sheet_path);

        if let Some(rels_data) = ctx.entries.get(&rels_path) {
            let ws_rels = Relationships::parse(rels_data)?;

            for rel in ws_rels.find_by_type(REL_TIMELINE) {
                let tl_path = resolve_rel_target(sheet_path, &rel.target);

                if let Some(tl_data) = ctx.entries.get(&tl_path) {
                    debug!("parsing timelines from: {}", tl_path);
                    sheet_timelines[i] = timelines::parse_timelines(tl_data)?;
                    ctx.known_dynamic.insert(tl_path);
                } else {
                    warn!("timeline file not found: {}", tl_path);
                }
            }
        }
    }

    Ok(sheet_timelines)
}

/// Resolve timeline cache definitions from workbook .rels.
fn resolve_timeline_caches(ctx: &mut ReaderContext) -> Result<Vec<TimelineCacheData>> {
    let mut caches: Vec<TimelineCacheData> = Vec::new();

    let cache_targets: Vec<String> = ctx
        .wb_rels
        .find_by_type(REL_TIMELINE_CACHE)
        .map(|rel| {
            if rel.target.starts_with('/') {
                rel.target.trim_start_matches('/').to_string()
            } else {
                format!("xl/{}", rel.target)
            }
        })
        .collect();

    for cache_path in cache_targets {
        if let Some(cache_data) = ctx.entries.get(&cache_path) {
            debug!("parsing timeline cache from: {}", cache_path);
            caches.push(TimelineCacheData::parse(cache_data)?);
            ctx.known_dynamic.insert(cache_path);
        } else {
            warn!("timeline cache file not found: {}", cache_path);
        }
    }

    Ok(caches)
}

/// Collect all ZIP entries that were not parsed into preserved_entries.
///
/// Takes ownership of entries via `drain()` to avoid cloning large byte
/// vectors (e.g. embedded images, charts).
fn collect_preserved(ctx: &mut ReaderContext) -> BTreeMap<String, Vec<u8>> {
    let mut preserved = BTreeMap::new();
    for (path, data) in ctx.entries.drain() {
        if !KNOWN_STATIC_PATHS.contains(&path.as_str()) && !ctx.known_dynamic.contains(&path) {
            debug!("preserving unknown ZIP entry: {}", path);
            preserved.insert(path, data);
        }
    }
    preserved
}

/// Derive the .rels path for a worksheet.
/// e.g. "xl/worksheets/sheet1.xml" -> "xl/worksheets/_rels/sheet1.xml.rels"
fn derive_rels_path(sheet_path: &str) -> String {
    if let Some(slash_pos) = sheet_path.rfind('/') {
        let dir = &sheet_path[..slash_pos];
        let file = &sheet_path[slash_pos + 1..];
        format!("{dir}/_rels/{file}.rels")
    } else {
        format!("_rels/{sheet_path}.rels")
    }
}

/// Resolve a relationship target path relative to a worksheet path.
fn resolve_rel_target(sheet_path: &str, target: &str) -> String {
    if target.starts_with('/') {
        target.trim_start_matches('/').to_string()
    } else if let Some(slash_pos) = sheet_path.rfind('/') {
        let dir = &sheet_path[..slash_pos];
        resolve_relative_path(dir, target)
    } else {
        target.to_string()
    }
}

// ---------------------------------------------------------------------------
// Public API — struct-based reader
// ---------------------------------------------------------------------------

/// Read an XLSX file from bytes using default security limits.
pub fn read_xlsx(data: &[u8]) -> Result<WorkbookData> {
    read_xlsx_with_options(data, &ZipSecurityLimits::default())
}

/// Read an XLSX file from bytes with custom ZIP security limits.
pub fn read_xlsx_with_options(data: &[u8], limits: &ZipSecurityLimits) -> Result<WorkbookData> {
    let mut ctx = parse_common(data, limits, None)?;
    let sheet_comments = resolve_comments(&mut ctx)?;
    let sheet_tables = resolve_tables(&mut ctx)?;
    let sheet_charts = resolve_charts(&mut ctx)?;
    let sheet_pivots = resolve_pivot_tables(&mut ctx)?;
    let (pivot_caches, pivot_cache_records) = resolve_pivot_caches(&mut ctx)?;
    let sheet_threaded = resolve_threaded_comments(&mut ctx)?;
    let persons = resolve_persons(&mut ctx)?;
    let sheet_slicers = resolve_slicers(&mut ctx)?;
    let slicer_caches = resolve_slicer_caches(&mut ctx)?;
    let sheet_timelines = resolve_timelines(&mut ctx)?;
    let timeline_caches = resolve_timeline_caches(&mut ctx)?;

    // Parse each worksheet XML.
    // When the "parallel" feature is enabled, sheets are parsed concurrently.
    let mut sheets = parse_sheets(&ctx.entries, &ctx.sheet_targets, &ctx.sst)?;

    // Attach comments to their respective worksheets.
    for (sheet, comments) in sheets.iter_mut().zip(sheet_comments) {
        if !comments.is_empty() {
            sheet.worksheet.comments = comments;
        }
    }

    // Attach tables to their respective worksheets.
    for (sheet, tables) in sheets.iter_mut().zip(sheet_tables) {
        if !tables.is_empty() {
            sheet.worksheet.tables = tables;
        }
    }

    // Attach charts to their respective worksheets.
    for (sheet, charts_vec) in sheets.iter_mut().zip(sheet_charts) {
        if !charts_vec.is_empty() {
            sheet.worksheet.charts = charts_vec;
        }
    }

    // Attach pivot tables to their respective worksheets.
    for (sheet, pivots) in sheets.iter_mut().zip(sheet_pivots) {
        if !pivots.is_empty() {
            sheet.worksheet.pivot_tables = pivots;
        }
    }

    // Attach threaded comments to their respective worksheets.
    for (sheet, tc) in sheets.iter_mut().zip(sheet_threaded) {
        if !tc.is_empty() {
            sheet.worksheet.threaded_comments = tc;
        }
    }

    // Attach slicers to their respective worksheets.
    for (sheet, sl) in sheets.iter_mut().zip(sheet_slicers) {
        if !sl.is_empty() {
            sheet.worksheet.slicers = sl;
        }
    }

    // Attach timelines to their respective worksheets.
    for (sheet, tl) in sheets.iter_mut().zip(sheet_timelines) {
        if !tl.is_empty() {
            sheet.worksheet.timelines = tl;
        }
    }

    let preserved_entries = collect_preserved(&mut ctx);

    Ok(WorkbookData {
        sheets,
        date_system: ctx.workbook_xml.date_system,
        styles: ctx.styles,
        defined_names: ctx.workbook_xml.defined_names,
        shared_strings: Some(ctx.sst),
        doc_properties: ctx.doc_properties,
        theme_colors: ctx.theme_colors,
        calc_chain: ctx.calc_chain,
        workbook_views: ctx.workbook_xml.workbook_views,
        protection: ctx.workbook_xml.protection,
        pivot_caches,
        pivot_cache_records,
        persons,
        slicer_caches,
        timeline_caches,
        preserved_entries,
    })
}

// ---------------------------------------------------------------------------
// Public API — streaming JSON reader
// ---------------------------------------------------------------------------

/// Read an XLSX file and return the result directly as a JSON string.
///
/// This is a WASM-optimized path that avoids creating millions of intermediate
/// `Cell`/`Row`/`String` objects. Worksheet row/cell data is written directly
/// as JSON during XML parsing, while small metadata (styles, defined names,
/// etc.) is serialized normally via `serde_json`.
///
/// The JSON output is compatible with `WorkbookData`'s serde format, so the
/// TypeScript side can use `JSON.parse()` to get the same structure.
pub fn read_xlsx_json(data: &[u8]) -> Result<String> {
    read_xlsx_json_with_options(data, &ZipSecurityLimits::default())
}

/// Read an XLSX file as JSON with custom ZIP security limits.
pub fn read_xlsx_json_with_options(data: &[u8], limits: &ZipSecurityLimits) -> Result<String> {
    build_json_from_context(parse_common(data, limits, None)?, data.len())
}

/// Read an XLSX file (possibly encrypted) from bytes with a password.
///
/// If the file is a plain ZIP (not encrypted), the password is ignored
/// and the file is read normally. If the file is OLE2-encrypted and the
/// password is empty, a descriptive error is returned.
pub fn read_xlsx_json_with_password(data: &[u8], password: &str) -> Result<String> {
    let pw = if password.is_empty() { None } else { Some(password) };
    build_json_from_context(parse_common(data, &ZipSecurityLimits::default(), pw)?, data.len())
}

/// Shared JSON builder used by both `read_xlsx_json_with_options` and
/// `read_xlsx_json_with_password`.
fn build_json_from_context(mut ctx: ReaderContext, data_len: usize) -> Result<String> {
    let sheet_comments = resolve_comments(&mut ctx)?;
    let sheet_tables = resolve_tables(&mut ctx)?;
    let sheet_charts = resolve_charts(&mut ctx)?;
    let sheet_pivots = resolve_pivot_tables(&mut ctx)?;
    let (pivot_caches, pivot_cache_records) = resolve_pivot_caches(&mut ctx)?;
    let sheet_threaded = resolve_threaded_comments(&mut ctx)?;
    let persons = resolve_persons(&mut ctx)?;
    let sheet_slicers = resolve_slicers(&mut ctx)?;
    let slicer_caches = resolve_slicer_caches(&mut ctx)?;
    let sheet_timelines = resolve_timelines(&mut ctx)?;
    let timeline_caches = resolve_timeline_caches(&mut ctx)?;

    // --- Build JSON output ---
    // Estimate ~80 bytes per cell × 10 cells/row × number of rows.
    // For a 5MB XLSX, expand factor ~15x is reasonable.
    let estimated_size = (data_len * 15).max(4096);
    let mut out = String::with_capacity(estimated_size);

    let serde_err = |e: serde_json::Error| ModernXlsxError::XmlParse(e.to_string());

    out.push_str("{\"sheets\":[");

    for (i, (name, path, state)) in ctx.sheet_targets.iter().enumerate() {
        if i > 0 { out.push(','); }
        out.push_str("{\"name\":\"");
        crate::ooxml::worksheet::json_escape_to_pub(&mut out, name);
        out.push('"');

        // Emit state only for non-visible sheets (skip_serializing_if equivalent).
        match state {
            SheetState::Hidden => out.push_str(",\"state\":\"hidden\""),
            SheetState::VeryHidden => out.push_str(",\"state\":\"veryHidden\""),
            SheetState::Visible => {}
        }

        out.push_str(",\"worksheet\":");

        let ws_data = ctx.entries.get(path).ok_or_else(|| {
            ModernXlsxError::MissingPart(format!("{} for sheet '{}'", path, name))
        })?;
        WorksheetXml::parse_to_json(ws_data, Some(&ctx.sst), &sheet_comments[i], &sheet_tables[i], &mut out)?;

        // Inject charts into the worksheet JSON (before the closing '}').
        if !sheet_charts[i].is_empty() {
            // Pop the closing '}' written by parse_to_json.
            debug_assert_eq!(out.as_bytes().last(), Some(&b'}'));
            out.pop();
            out.push_str(",\"charts\":");
            out.push_str(
                &serde_json::to_string(&sheet_charts[i]).map_err(serde_err)?,
            );
            out.push('}');
        }

        // Inject pivot tables into the worksheet JSON (before the closing '}').
        if !sheet_pivots[i].is_empty() {
            debug_assert_eq!(out.as_bytes().last(), Some(&b'}'));
            out.pop();
            out.push_str(",\"pivotTables\":");
            out.push_str(
                &serde_json::to_string(&sheet_pivots[i]).map_err(serde_err)?,
            );
            out.push('}');
        }

        // Inject threaded comments into the worksheet JSON (before the closing '}').
        if !sheet_threaded[i].is_empty() {
            debug_assert_eq!(out.as_bytes().last(), Some(&b'}'));
            out.pop();
            out.push_str(",\"threadedComments\":");
            out.push_str(
                &serde_json::to_string(&sheet_threaded[i]).map_err(serde_err)?,
            );
            out.push('}');
        }

        // Inject slicers into the worksheet JSON (before the closing '}').
        if !sheet_slicers[i].is_empty() {
            debug_assert_eq!(out.as_bytes().last(), Some(&b'}'));
            out.pop();
            out.push_str(",\"slicers\":");
            out.push_str(
                &serde_json::to_string(&sheet_slicers[i]).map_err(serde_err)?,
            );
            out.push('}');
        }

        // Inject timelines into the worksheet JSON (before the closing '}').
        if !sheet_timelines[i].is_empty() {
            debug_assert_eq!(out.as_bytes().last(), Some(&b'}'));
            out.pop();
            out.push_str(",\"timelines\":");
            out.push_str(
                &serde_json::to_string(&sheet_timelines[i]).map_err(serde_err)?,
            );
            out.push('}');
        }

        out.push('}');
    }

    out.push(']');

    // Drain entries AFTER sheets are parsed to avoid cloning preserved data.
    let preserved_entries = collect_preserved(&mut ctx);

    // dateSystem
    out.push_str(",\"dateSystem\":");
    out.push_str(&serde_json::to_string(&ctx.workbook_xml.date_system).map_err(serde_err)?);

    // styles
    out.push_str(",\"styles\":");
    out.push_str(&serde_json::to_string(&ctx.styles).map_err(serde_err)?);

    // definedNames
    if !ctx.workbook_xml.defined_names.is_empty() {
        out.push_str(",\"definedNames\":");
        out.push_str(&serde_json::to_string(&ctx.workbook_xml.defined_names).map_err(serde_err)?);
    }

    // sharedStrings
    out.push_str(",\"sharedStrings\":");
    out.push_str(&serde_json::to_string(&ctx.sst).map_err(serde_err)?);

    // docProperties
    if let Some(ref dp) = ctx.doc_properties {
        out.push_str(",\"docProperties\":");
        out.push_str(&serde_json::to_string(dp).map_err(serde_err)?);
    }

    // themeColors
    if let Some(ref tc) = ctx.theme_colors {
        out.push_str(",\"themeColors\":");
        out.push_str(&serde_json::to_string(tc).map_err(serde_err)?);
    }

    // calcChain
    if !ctx.calc_chain.is_empty() {
        out.push_str(",\"calcChain\":");
        out.push_str(&serde_json::to_string(&ctx.calc_chain).map_err(serde_err)?);
    }

    // workbookViews
    if !ctx.workbook_xml.workbook_views.is_empty() {
        out.push_str(",\"workbookViews\":");
        out.push_str(&serde_json::to_string(&ctx.workbook_xml.workbook_views).map_err(serde_err)?);
    }

    // protection
    if let Some(ref prot) = ctx.workbook_xml.protection {
        out.push_str(",\"protection\":");
        out.push_str(&serde_json::to_string(prot).map_err(serde_err)?);
    }

    // pivotCaches
    if !pivot_caches.is_empty() {
        out.push_str(",\"pivotCaches\":");
        out.push_str(&serde_json::to_string(&pivot_caches).map_err(serde_err)?);
    }

    // pivotCacheRecords
    if !pivot_cache_records.is_empty() {
        out.push_str(",\"pivotCacheRecords\":");
        out.push_str(&serde_json::to_string(&pivot_cache_records).map_err(serde_err)?);
    }

    // persons
    if !persons.is_empty() {
        out.push_str(",\"persons\":");
        out.push_str(&serde_json::to_string(&persons).map_err(serde_err)?);
    }

    // slicerCaches
    if !slicer_caches.is_empty() {
        out.push_str(",\"slicerCaches\":");
        out.push_str(&serde_json::to_string(&slicer_caches).map_err(serde_err)?);
    }

    // timelineCaches
    if !timeline_caches.is_empty() {
        out.push_str(",\"timelineCaches\":");
        out.push_str(&serde_json::to_string(&timeline_caches).map_err(serde_err)?);
    }

    // preservedEntries
    if !preserved_entries.is_empty() {
        out.push_str(",\"preservedEntries\":");
        out.push_str(&serde_json::to_string(&preserved_entries).map_err(serde_err)?);
    }

    out.push('}');

    Ok(out)
}

/// Parse worksheet XML for each (name, path) pair.
///
/// When compiled with the `parallel` feature, parsing runs on rayon's
/// thread-pool via `par_iter()`. The `entries` map is `&HashMap` which is
/// `Sync`, so it can be shared across threads safely.
///
/// The SST is passed through to resolve shared string indices inline
/// during parsing, which is significantly faster than a post-parse pass.
#[cfg(feature = "parallel")]
fn parse_sheets(
    entries: &std::collections::HashMap<String, Vec<u8>>,
    sheet_targets: &[(String, String, SheetState)],
    sst: &SharedStringTable,
) -> Result<Vec<SheetData>> {
    sheet_targets
        .par_iter()
        .map(|(name, path, state)| {
            debug!("parsing worksheet (parallel): {}", name);
            let sheet_data = entries.get(path).ok_or_else(|| {
                ModernXlsxError::MissingPart(format!("{} for sheet '{}'", path, name))
            })?;
            let worksheet = WorksheetXml::parse_with_sst(sheet_data, Some(sst))?;
            Ok(SheetData {
                name: name.clone(),
                state: sheet_state_to_str(*state),
                worksheet,
            })
        })
        .collect()
}

/// Parse worksheet XML for each (name, path, state) triple sequentially.
///
/// The SST is passed through to resolve shared string indices inline
/// during parsing, which is significantly faster than a post-parse pass.
#[cfg(not(feature = "parallel"))]
fn parse_sheets(
    entries: &std::collections::HashMap<String, Vec<u8>>,
    sheet_targets: &[(String, String, SheetState)],
    sst: &SharedStringTable,
) -> Result<Vec<SheetData>> {
    sheet_targets
        .iter()
        .map(|(name, path, state)| {
            debug!("parsing worksheet: {}", name);
            let sheet_data = entries.get(path).ok_or_else(|| {
                ModernXlsxError::MissingPart(format!("{} for sheet '{}'", path, name))
            })?;
            let worksheet = WorksheetXml::parse_with_sst(sheet_data, Some(sst))?;
            Ok(SheetData {
                name: name.clone(),
                state: sheet_state_to_str(*state),
                worksheet,
            })
        })
        .collect()
}

/// Resolve a relative path (which may contain `..` segments) against a base
/// directory. For example, resolving `"../comments1.xml"` against
/// `"xl/worksheets"` yields `"xl/comments1.xml"`.
fn resolve_relative_path(base_dir: &str, relative: &str) -> String {
    let mut parts: Vec<&str> = base_dir.split('/').collect();
    for segment in relative.split('/') {
        match segment {
            ".." => {
                if parts.len() > 1 {
                    parts.pop();
                }
            }
            "." | "" => {}
            other => {
                parts.push(other);
            }
        }
    }
    parts.join("/")
}

#[cfg(test)]
mod tests {
    use super::*;
    use pretty_assertions::assert_eq;
    use crate::dates::DateSystem;
    use crate::ooxml::{
        content_types::ContentTypes,
        relationships::Relationships,
        shared_strings::SharedStringTableBuilder,
        styles::Styles,
        workbook::{SheetInfo, SheetState, WorkbookXml},
        worksheet::{Cell, CellType, Row, WorksheetXml},
    };
    use crate::zip::writer::{write_zip, ZipEntry};

    /// Build a minimal valid XLSX archive as raw bytes.
    ///
    /// The workbook contains one sheet named "Sheet1" with two cells:
    ///   A1 = "Hello" (shared string)
    ///   B1 = 42 (number)
    fn build_minimal_xlsx() -> Vec<u8> {
        // Content types
        let ct = ContentTypes::for_basic_workbook(1);
        let ct_xml = ct.to_xml().unwrap();

        // Root rels
        let root_rels = Relationships::root_rels();
        let root_rels_xml = root_rels.to_xml().unwrap();

        // Workbook rels
        let wb_rels = Relationships::workbook_rels(1);
        let wb_rels_xml = wb_rels.to_xml().unwrap();

        // Workbook
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
        let wb_xml = wb.to_xml().unwrap();

        // Shared strings: index 0 = "Hello"
        let mut sst_builder = SharedStringTableBuilder::new();
        sst_builder.insert("Hello");
        let sst_xml = sst_builder.to_xml().unwrap();

        // Styles
        let styles = Styles::default_styles();
        let styles_xml = styles.to_xml().unwrap();

        // Worksheet with A1="Hello" (shared string index 0), B1=42 (number)
        let ws = WorksheetXml {
            dimension: Some("A1:B1".to_string()),
            rows: vec![Row {
                index: 1,
                cells: vec![
                    Cell {
                        reference: "A1".to_string(),
                        cell_type: CellType::SharedString,
                        style_index: None,
                        value: Some("0".to_string()),
                        ..Default::default()
                    },
                    Cell {
                        reference: "B1".to_string(),
                        cell_type: CellType::Number,
                        style_index: None,
                        value: Some("42".to_string()),
                        ..Default::default()
                    },
                ],
                height: None,
                hidden: false,
                outline_level: None,
                collapsed: false,
            }],
            merge_cells: Vec::new(),
            auto_filter: None,
            frozen_pane: None,
            split_pane: None,
            pane_selections: vec![],
            sheet_view: None,
            columns: Vec::new(),
            data_validations: Vec::new(),
            conditional_formatting: Vec::new(),
            hyperlinks: Vec::new(),
            page_setup: None,
            sheet_protection: None,
            comments: Vec::new(),
            tab_color: None,
            tables: Vec::new(),
            header_footer: None,
            outline_properties: None,
            sparkline_groups: Vec::new(),
            charts: Vec::new(),
            pivot_tables: Vec::new(),
            threaded_comments: Vec::new(),
            slicers: Vec::new(),
            timelines: Vec::new(),
            preserved_extensions: Vec::new(),
        };
        let ws_xml = ws.to_xml().unwrap();

        let entries = vec![
            ZipEntry {
                name: "[Content_Types].xml".to_string(),
                data: ct_xml,
            },
            ZipEntry {
                name: "_rels/.rels".to_string(),
                data: root_rels_xml,
            },
            ZipEntry {
                name: "xl/workbook.xml".to_string(),
                data: wb_xml,
            },
            ZipEntry {
                name: "xl/_rels/workbook.xml.rels".to_string(),
                data: wb_rels_xml,
            },
            ZipEntry {
                name: "xl/sharedStrings.xml".to_string(),
                data: sst_xml,
            },
            ZipEntry {
                name: "xl/styles.xml".to_string(),
                data: styles_xml,
            },
            ZipEntry {
                name: "xl/worksheets/sheet1.xml".to_string(),
                data: ws_xml,
            },
        ];

        write_zip(&entries).unwrap()
    }

    /// Build an XLSX archive without a sharedStrings.xml part.
    /// The worksheet uses only numeric cells (no shared strings needed).
    fn build_xlsx_without_sst() -> Vec<u8> {
        // Content types — manually construct without the SST override
        let mut ct = ContentTypes::for_basic_workbook(1);
        ct.overrides.remove("/xl/sharedStrings.xml");
        let ct_xml = ct.to_xml().unwrap();

        // Root rels
        let root_rels = Relationships::root_rels();
        let root_rels_xml = root_rels.to_xml().unwrap();

        // Workbook rels — only worksheet + styles, no sharedStrings
        let mut wb_rels = Relationships::new();
        wb_rels.add(
            "rId1",
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet",
            "worksheets/sheet1.xml",
        );
        wb_rels.add(
            "rId2",
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles",
            "styles.xml",
        );
        let wb_rels_xml = wb_rels.to_xml().unwrap();

        // Workbook
        let wb = WorkbookXml {
            sheets: vec![SheetInfo {
                name: "Numbers".to_string(),
                sheet_id: 1,
                r_id: "rId1".to_string(),
                state: SheetState::Visible,
            }],
            date_system: DateSystem::Date1900,
            defined_names: Vec::new(),
            workbook_views: Vec::new(),
            protection: None,
        };
        let wb_xml = wb.to_xml().unwrap();

        // Styles
        let styles = Styles::default_styles();
        let styles_xml = styles.to_xml().unwrap();

        // Worksheet — numeric only
        let ws = WorksheetXml {
            dimension: Some("A1".to_string()),
            rows: vec![Row {
                index: 1,
                cells: vec![Cell {
                    reference: "A1".to_string(),
                    cell_type: CellType::Number,
                    style_index: None,
                    value: Some("99".to_string()),
                    ..Default::default()
                }],
                height: None,
                hidden: false,
                outline_level: None,
                collapsed: false,
            }],
            merge_cells: Vec::new(),
            auto_filter: None,
            frozen_pane: None,
            split_pane: None,
            pane_selections: vec![],
            sheet_view: None,
            columns: Vec::new(),
            data_validations: Vec::new(),
            conditional_formatting: Vec::new(),
            hyperlinks: Vec::new(),
            page_setup: None,
            sheet_protection: None,
            comments: Vec::new(),
            tab_color: None,
            tables: Vec::new(),
            header_footer: None,
            outline_properties: None,
            sparkline_groups: Vec::new(),
            charts: Vec::new(),
            pivot_tables: Vec::new(),
            threaded_comments: Vec::new(),
            slicers: Vec::new(),
            timelines: Vec::new(),
            preserved_extensions: Vec::new(),
        };
        let ws_xml = ws.to_xml().unwrap();

        let entries = vec![
            ZipEntry {
                name: "[Content_Types].xml".to_string(),
                data: ct_xml,
            },
            ZipEntry {
                name: "_rels/.rels".to_string(),
                data: root_rels_xml,
            },
            ZipEntry {
                name: "xl/workbook.xml".to_string(),
                data: wb_xml,
            },
            ZipEntry {
                name: "xl/_rels/workbook.xml.rels".to_string(),
                data: wb_rels_xml,
            },
            ZipEntry {
                name: "xl/styles.xml".to_string(),
                data: styles_xml,
            },
            ZipEntry {
                name: "xl/worksheets/sheet1.xml".to_string(),
                data: ws_xml,
            },
        ];

        write_zip(&entries).unwrap()
    }

    /// Build an XLSX archive that uses the 1904 date system.
    fn build_xlsx_1904() -> Vec<u8> {
        // Content types
        let ct = ContentTypes::for_basic_workbook(1);
        let ct_xml = ct.to_xml().unwrap();

        // Root rels
        let root_rels = Relationships::root_rels();
        let root_rels_xml = root_rels.to_xml().unwrap();

        // Workbook rels
        let wb_rels = Relationships::workbook_rels(1);
        let wb_rels_xml = wb_rels.to_xml().unwrap();

        // Workbook with date1904
        let wb = WorkbookXml {
            sheets: vec![SheetInfo {
                name: "Sheet1".to_string(),
                sheet_id: 1,
                r_id: "rId1".to_string(),
                state: SheetState::Visible,
            }],
            date_system: DateSystem::Date1904,
            defined_names: Vec::new(),
            workbook_views: Vec::new(),
            protection: None,
        };
        let wb_xml = wb.to_xml().unwrap();

        // Shared strings
        let sst_builder = SharedStringTableBuilder::new();
        let sst_xml = sst_builder.to_xml().unwrap();

        // Styles
        let styles = Styles::default_styles();
        let styles_xml = styles.to_xml().unwrap();

        // Empty worksheet
        let ws = WorksheetXml {
            dimension: None,
            rows: Vec::new(),
            merge_cells: Vec::new(),
            auto_filter: None,
            frozen_pane: None,
            split_pane: None,
            pane_selections: vec![],
            sheet_view: None,
            columns: Vec::new(),
            data_validations: Vec::new(),
            conditional_formatting: Vec::new(),
            hyperlinks: Vec::new(),
            page_setup: None,
            sheet_protection: None,
            comments: Vec::new(),
            tab_color: None,
            tables: Vec::new(),
            header_footer: None,
            outline_properties: None,
            sparkline_groups: Vec::new(),
            charts: Vec::new(),
            pivot_tables: Vec::new(),
            threaded_comments: Vec::new(),
            slicers: Vec::new(),
            timelines: Vec::new(),
            preserved_extensions: Vec::new(),
        };
        let ws_xml = ws.to_xml().unwrap();

        let entries = vec![
            ZipEntry {
                name: "[Content_Types].xml".to_string(),
                data: ct_xml,
            },
            ZipEntry {
                name: "_rels/.rels".to_string(),
                data: root_rels_xml,
            },
            ZipEntry {
                name: "xl/workbook.xml".to_string(),
                data: wb_xml,
            },
            ZipEntry {
                name: "xl/_rels/workbook.xml.rels".to_string(),
                data: wb_rels_xml,
            },
            ZipEntry {
                name: "xl/sharedStrings.xml".to_string(),
                data: sst_xml,
            },
            ZipEntry {
                name: "xl/styles.xml".to_string(),
                data: styles_xml,
            },
            ZipEntry {
                name: "xl/worksheets/sheet1.xml".to_string(),
                data: ws_xml,
            },
        ];

        write_zip(&entries).unwrap()
    }

    #[test]
    fn test_read_minimal_workbook() {
        let xlsx_bytes = build_minimal_xlsx();
        let wb = read_xlsx(&xlsx_bytes).expect("read_xlsx should succeed");

        // One sheet named "Sheet1".
        assert_eq!(wb.sheets.len(), 1);
        assert_eq!(wb.sheets[0].name, "Sheet1");

        // Date system is 1900 (default).
        assert_eq!(wb.date_system, DateSystem::Date1900);

        // Shared string table has one entry: "Hello".
        let sst = wb.shared_strings.as_ref().expect("shared_strings should be Some");
        assert_eq!(sst.len(), 1);
        assert_eq!(sst.get(0), Some("Hello"));

        // Worksheet has one row with two cells.
        let ws = &wb.sheets[0].worksheet;
        assert_eq!(ws.rows.len(), 1);
        assert_eq!(ws.rows[0].cells.len(), 2);

        // A1 is a shared string reference (value resolved to "Hello").
        let a1 = &ws.rows[0].cells[0];
        assert_eq!(a1.reference, "A1");
        assert_eq!(a1.cell_type, CellType::SharedString);
        assert_eq!(a1.value.as_deref(), Some("Hello"));

        // B1 is a number (value = "42").
        let b1 = &ws.rows[0].cells[1];
        assert_eq!(b1.reference, "B1");
        assert_eq!(b1.cell_type, CellType::Number);
        assert_eq!(b1.value.as_deref(), Some("42"));
    }

    #[test]
    fn test_read_without_shared_strings() {
        let xlsx_bytes = build_xlsx_without_sst();
        let wb = read_xlsx(&xlsx_bytes).expect("read_xlsx should succeed without SST");

        // Sheet should still load.
        assert_eq!(wb.sheets.len(), 1);
        assert_eq!(wb.sheets[0].name, "Numbers");

        // Shared strings table should be present but empty (not an error).
        let sst = wb.shared_strings.as_ref().expect("shared_strings should be Some");
        assert!(sst.is_empty());

        // The numeric cell should parse correctly.
        let ws = &wb.sheets[0].worksheet;
        assert_eq!(ws.rows.len(), 1);
        let a1 = &ws.rows[0].cells[0];
        assert_eq!(a1.cell_type, CellType::Number);
        assert_eq!(a1.value.as_deref(), Some("99"));
    }

    #[test]
    fn test_read_1904_date_system() {
        let xlsx_bytes = build_xlsx_1904();
        let wb = read_xlsx(&xlsx_bytes).expect("read_xlsx should succeed for 1904 workbook");

        assert_eq!(wb.date_system, DateSystem::Date1904);
        assert_eq!(wb.sheets.len(), 1);
        assert_eq!(wb.sheets[0].name, "Sheet1");
    }

    #[test]
    fn test_read_missing_workbook_xml() {
        // Build a ZIP with no xl/workbook.xml.
        let entries = vec![ZipEntry {
            name: "[Content_Types].xml".to_string(),
            data: b"<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"/>".to_vec(),
        }];
        let zip_bytes = write_zip(&entries).unwrap();

        let result = read_xlsx(&zip_bytes);
        assert!(result.is_err());
        let err = result.unwrap_err();
        assert!(
            matches!(err, ModernXlsxError::MissingPart(ref msg) if msg.contains("workbook.xml")),
            "expected MissingPart error, got: {err}"
        );
    }

    #[test]
    fn test_preserved_entries_roundtrip() {
        // 1. Build a minimal XLSX with extra ZIP entries simulating drawings/media.
        let ct = ContentTypes::for_basic_workbook(1);
        let ct_xml = ct.to_xml().unwrap();
        let root_rels = Relationships::root_rels();
        let root_rels_xml = root_rels.to_xml().unwrap();
        let wb_rels = Relationships::workbook_rels(1);
        let wb_rels_xml = wb_rels.to_xml().unwrap();
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
        let wb_xml = wb.to_xml().unwrap();
        let sst_builder = SharedStringTableBuilder::new();
        let sst_xml = sst_builder.to_xml().unwrap();
        let styles = Styles::default_styles();
        let styles_xml = styles.to_xml().unwrap();
        let ws = WorksheetXml {
            dimension: None,
            rows: Vec::new(),
            merge_cells: Vec::new(),
            auto_filter: None,
            frozen_pane: None,
            split_pane: None,
            pane_selections: vec![],
            sheet_view: None,
            columns: Vec::new(),
            data_validations: Vec::new(),
            conditional_formatting: Vec::new(),
            hyperlinks: Vec::new(),
            page_setup: None,
            sheet_protection: None,
            comments: Vec::new(),
            tab_color: None,
            tables: Vec::new(),
            header_footer: None,
            outline_properties: None,
            sparkline_groups: Vec::new(),
            charts: Vec::new(),
            pivot_tables: Vec::new(),
            threaded_comments: Vec::new(),
            slicers: Vec::new(),
            timelines: Vec::new(),
            preserved_extensions: Vec::new(),
        };
        let ws_xml = ws.to_xml().unwrap();

        // Fake drawing and image data
        let drawing_data = b"<xdr:wsDr xmlns:xdr=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\"/>".to_vec();
        let image_data = vec![0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A]; // PNG header
        let chart_data = b"<c:chartSpace xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\"/>".to_vec();
        let sheet_rels_data = b"<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"/>".to_vec();

        let entries = vec![
            ZipEntry { name: "[Content_Types].xml".to_string(), data: ct_xml },
            ZipEntry { name: "_rels/.rels".to_string(), data: root_rels_xml },
            ZipEntry { name: "xl/workbook.xml".to_string(), data: wb_xml },
            ZipEntry { name: "xl/_rels/workbook.xml.rels".to_string(), data: wb_rels_xml },
            ZipEntry { name: "xl/sharedStrings.xml".to_string(), data: sst_xml },
            ZipEntry { name: "xl/styles.xml".to_string(), data: styles_xml },
            ZipEntry { name: "xl/worksheets/sheet1.xml".to_string(), data: ws_xml },
            // Extra entries to preserve
            ZipEntry { name: "xl/drawings/drawing1.xml".to_string(), data: drawing_data.clone() },
            ZipEntry { name: "xl/media/image1.png".to_string(), data: image_data.clone() },
            ZipEntry { name: "xl/charts/chart1.xml".to_string(), data: chart_data.clone() },
            ZipEntry { name: "xl/worksheets/_rels/sheet1.xml.rels".to_string(), data: sheet_rels_data.clone() },
        ];

        let zip_bytes = write_zip(&entries).unwrap();

        // 2. Read the XLSX — preserved entries should be captured.
        let wb1 = read_xlsx(&zip_bytes).expect("read_xlsx should succeed");
        assert_eq!(wb1.preserved_entries.len(), 4);
        assert_eq!(wb1.preserved_entries.get("xl/drawings/drawing1.xml").unwrap(), &drawing_data);
        assert_eq!(wb1.preserved_entries.get("xl/media/image1.png").unwrap(), &image_data);
        assert_eq!(wb1.preserved_entries.get("xl/charts/chart1.xml").unwrap(), &chart_data);
        assert_eq!(wb1.preserved_entries.get("xl/worksheets/_rels/sheet1.xml.rels").unwrap(), &sheet_rels_data);

        // 3. Write it back and read again — the entries should survive the roundtrip.
        let xlsx_bytes2 = crate::writer::write_xlsx(&wb1).expect("write_xlsx should succeed");
        let wb2 = read_xlsx(&xlsx_bytes2).expect("read_xlsx should succeed on second pass");

        assert_eq!(wb2.preserved_entries.len(), 4);
        assert_eq!(wb2.preserved_entries.get("xl/drawings/drawing1.xml").unwrap(), &drawing_data);
        assert_eq!(wb2.preserved_entries.get("xl/media/image1.png").unwrap(), &image_data);
        assert_eq!(wb2.preserved_entries.get("xl/charts/chart1.xml").unwrap(), &chart_data);
        assert_eq!(wb2.preserved_entries.get("xl/worksheets/_rels/sheet1.xml.rels").unwrap(), &sheet_rels_data);
    }

    #[test]
    fn test_no_preserved_entries_for_standard_xlsx() {
        // A standard XLSX with no extra entries should have an empty preserved_entries map.
        let xlsx_bytes = build_minimal_xlsx();
        let wb = read_xlsx(&xlsx_bytes).expect("read_xlsx should succeed");
        assert!(wb.preserved_entries.is_empty(), "standard XLSX should have no preserved entries");
    }

    /// Build an XLSX with multiple sheets to exercise parallel parsing.
    fn build_multi_sheet_xlsx(sheet_count: usize) -> Vec<u8> {
        let ct = ContentTypes::for_basic_workbook(sheet_count);
        let ct_xml = ct.to_xml().unwrap();

        let root_rels = Relationships::root_rels();
        let root_rels_xml = root_rels.to_xml().unwrap();

        let wb_rels = Relationships::workbook_rels(sheet_count);
        let wb_rels_xml = wb_rels.to_xml().unwrap();

        let sheets_info: Vec<SheetInfo> = (1..=sheet_count)
            .map(|i| SheetInfo {
                name: format!("Sheet{}", i),
                sheet_id: i as u32,
                r_id: format!("rId{}", i),
                state: SheetState::Visible,
            })
            .collect();

        let wb = WorkbookXml {
            sheets: sheets_info,
            date_system: DateSystem::Date1900,
            defined_names: Vec::new(),
            workbook_views: Vec::new(),
            protection: None,
        };
        let wb_xml = wb.to_xml().unwrap();

        let sst_builder = SharedStringTableBuilder::new();
        let sst_xml = sst_builder.to_xml().unwrap();

        let styles = Styles::default_styles();
        let styles_xml = styles.to_xml().unwrap();

        let mut entries = vec![
            ZipEntry { name: "[Content_Types].xml".to_string(), data: ct_xml },
            ZipEntry { name: "_rels/.rels".to_string(), data: root_rels_xml },
            ZipEntry { name: "xl/workbook.xml".to_string(), data: wb_xml },
            ZipEntry { name: "xl/_rels/workbook.xml.rels".to_string(), data: wb_rels_xml },
            ZipEntry { name: "xl/sharedStrings.xml".to_string(), data: sst_xml },
            ZipEntry { name: "xl/styles.xml".to_string(), data: styles_xml },
        ];

        // Create a unique worksheet for each sheet with a distinct cell value.
        for i in 1..=sheet_count {
            let ws = WorksheetXml {
                dimension: Some("A1".to_string()),
                rows: vec![Row {
                    index: 1,
                    cells: vec![Cell {
                        reference: "A1".to_string(),
                        cell_type: CellType::Number,
                        style_index: None,
                        value: Some(format!("{}", i * 100)),
                        ..Default::default()
                    }],
                    height: None,
                    hidden: false,
                    outline_level: None,
                    collapsed: false,
                }],
                merge_cells: Vec::new(),
                auto_filter: None,
                frozen_pane: None,
            split_pane: None,
            pane_selections: vec![],
            sheet_view: None,
                columns: Vec::new(),
                data_validations: Vec::new(),
                conditional_formatting: Vec::new(),
                hyperlinks: Vec::new(),
                page_setup: None,
                sheet_protection: None,
                comments: Vec::new(),
                tab_color: None,
                tables: Vec::new(),
                header_footer: None,
                outline_properties: None,
                sparkline_groups: Vec::new(),
                charts: Vec::new(),
                pivot_tables: Vec::new(),
                threaded_comments: Vec::new(),
                slicers: Vec::new(),
                timelines: Vec::new(),
                preserved_extensions: Vec::new(),
            };
            let ws_xml = ws.to_xml().unwrap();
            entries.push(ZipEntry {
                name: format!("xl/worksheets/sheet{}.xml", i),
                data: ws_xml,
            });
        }

        write_zip(&entries).unwrap()
    }

    #[test]
    fn test_multi_sheet_parsing() {
        // This test exercises the parse_sheets function (parallel or sequential
        // depending on the feature flag). With 4 sheets, the parallel path has
        // enough work to actually distribute across threads.
        let xlsx_bytes = build_multi_sheet_xlsx(4);
        let wb = read_xlsx(&xlsx_bytes).expect("read_xlsx should succeed");

        assert_eq!(wb.sheets.len(), 4);

        for (i, sheet) in wb.sheets.iter().enumerate() {
            let expected_name = format!("Sheet{}", i + 1);
            assert_eq!(sheet.name, expected_name);

            // Each sheet has one row with A1 = (i+1)*100.
            assert_eq!(sheet.worksheet.rows.len(), 1);
            let cell = &sheet.worksheet.rows[0].cells[0];
            assert_eq!(cell.reference, "A1");
            assert_eq!(cell.cell_type, CellType::Number);
            assert_eq!(cell.value.as_deref(), Some(&*format!("{}", (i + 1) * 100)));
        }
    }

    #[test]
    fn test_comments_full_roundtrip() {
        use crate::ooxml::comments::Comment;

        // 1. Build a workbook with comments via the writer.
        let wb1 = WorkbookData {
            sheets: vec![SheetData {
                name: "Sheet1".to_string(),
                state: None,
                worksheet: WorksheetXml {
                    dimension: Some("A1:B1".to_string()),
                    rows: vec![Row {
                        index: 1,
                        cells: vec![
                            Cell {
                                reference: "A1".to_string(),
                                cell_type: CellType::Number,
                                style_index: None,
                                value: Some("42".to_string()),
                                ..Default::default()
                            },
                            Cell {
                                reference: "B1".to_string(),
                                cell_type: CellType::Number,
                                style_index: None,
                                value: Some("99".to_string()),
                                ..Default::default()
                            },
                        ],
                        height: None,
                        hidden: false,
                        outline_level: None,
                        collapsed: false,
                    }],
                    merge_cells: Vec::new(),
                    auto_filter: None,
                    frozen_pane: None,
            split_pane: None,
            pane_selections: vec![],
            sheet_view: None,
                    columns: Vec::new(),
                    data_validations: Vec::new(),
                    conditional_formatting: Vec::new(),
                    hyperlinks: Vec::new(),
                    page_setup: None,
                    sheet_protection: None,
                    comments: vec![
                        Comment {
                            cell_ref: "A1".to_string(),
                            author: "Alice".to_string(),
                            text: "This is the answer".to_string(),
                        },
                        Comment {
                            cell_ref: "B1".to_string(),
                            author: "Bob".to_string(),
                            text: "Almost 100".to_string(),
                        },
                    ],
                    tab_color: None,
                    tables: Vec::new(),
                    header_footer: None,
                    outline_properties: None,
                    sparkline_groups: Vec::new(),
                    charts: Vec::new(),
                    pivot_tables: Vec::new(),
                    threaded_comments: Vec::new(),
                    slicers: Vec::new(),
                    timelines: Vec::new(),
                    preserved_extensions: Vec::new(),
                },
            }],
            date_system: DateSystem::Date1900,
            styles: Styles::default_styles(),
            defined_names: Vec::new(),
            shared_strings: None,
            doc_properties: None,
            theme_colors: None,
            calc_chain: Vec::new(),
            workbook_views: Vec::new(),
            protection: None,
            pivot_caches: Vec::new(),
            pivot_cache_records: Vec::new(),
            persons: Vec::new(),
            slicer_caches: Vec::new(),
            timeline_caches: Vec::new(),
            preserved_entries: std::collections::BTreeMap::new(),
        };

        // Write to XLSX bytes.
        let xlsx1 = crate::writer::write_xlsx(&wb1).expect("write_xlsx should succeed");

        // Read it back.
        let wb2 = read_xlsx(&xlsx1).expect("read_xlsx should succeed");
        assert_eq!(wb2.sheets.len(), 1);
        assert_eq!(wb2.sheets[0].worksheet.comments.len(), 2);

        let c0 = &wb2.sheets[0].worksheet.comments[0];
        assert_eq!(c0.cell_ref, "A1");
        assert_eq!(c0.author, "Alice");
        assert_eq!(c0.text, "This is the answer");

        let c1 = &wb2.sheets[0].worksheet.comments[1];
        assert_eq!(c1.cell_ref, "B1");
        assert_eq!(c1.author, "Bob");
        assert_eq!(c1.text, "Almost 100");

        // Write again and read again (second roundtrip).
        let xlsx2 = crate::writer::write_xlsx(&wb2).expect("second write_xlsx should succeed");
        let wb3 = read_xlsx(&xlsx2).expect("second read_xlsx should succeed");

        assert_eq!(wb3.sheets[0].worksheet.comments.len(), 2);
        assert_eq!(wb3.sheets[0].worksheet.comments[0].cell_ref, "A1");
        assert_eq!(wb3.sheets[0].worksheet.comments[0].author, "Alice");
        assert_eq!(wb3.sheets[0].worksheet.comments[0].text, "This is the answer");
        assert_eq!(wb3.sheets[0].worksheet.comments[1].cell_ref, "B1");
        assert_eq!(wb3.sheets[0].worksheet.comments[1].author, "Bob");
        assert_eq!(wb3.sheets[0].worksheet.comments[1].text, "Almost 100");
    }

    #[test]
    fn test_comments_empty_no_crash() {
        // Workbook with no comments should parse and roundtrip fine.
        let xlsx_bytes = build_minimal_xlsx();
        let wb = read_xlsx(&xlsx_bytes).expect("read_xlsx should succeed");

        // No comments should be present.
        assert!(wb.sheets[0].worksheet.comments.is_empty());

        // Roundtrip should work.
        let xlsx2 = crate::writer::write_xlsx(&wb).expect("write_xlsx should succeed");
        let wb2 = read_xlsx(&xlsx2).expect("read_xlsx should succeed");
        assert!(wb2.sheets[0].worksheet.comments.is_empty());
    }

    #[test]
    fn test_comments_multi_sheet_roundtrip() {
        use crate::ooxml::comments::Comment;

        // Build a workbook with 2 sheets, only the second has comments.
        let wb1 = WorkbookData {
            sheets: vec![
                SheetData {
                    name: "NoComments".to_string(),
                    state: None,
                    worksheet: WorksheetXml {
                        dimension: None,
                        rows: Vec::new(),
                        merge_cells: Vec::new(),
                        auto_filter: None,
                        frozen_pane: None,
            split_pane: None,
            pane_selections: vec![],
            sheet_view: None,
                        columns: Vec::new(),
                        data_validations: Vec::new(),
                        conditional_formatting: Vec::new(),
                        hyperlinks: Vec::new(),
                        page_setup: None,
                        sheet_protection: None,
                        comments: Vec::new(),
                        tab_color: None,
                        tables: Vec::new(),
                        header_footer: None,
                        outline_properties: None,
                        sparkline_groups: Vec::new(),
                        charts: Vec::new(),
                        pivot_tables: Vec::new(),
                        threaded_comments: Vec::new(),
                        slicers: Vec::new(),
                        timelines: Vec::new(),
                        preserved_extensions: Vec::new(),
                    },
                },
                SheetData {
                    name: "WithComments".to_string(),
                    state: None,
                    worksheet: WorksheetXml {
                        dimension: None,
                        rows: Vec::new(),
                        merge_cells: Vec::new(),
                        auto_filter: None,
                        frozen_pane: None,
            split_pane: None,
            pane_selections: vec![],
            sheet_view: None,
                        columns: Vec::new(),
                        data_validations: Vec::new(),
                        conditional_formatting: Vec::new(),
                        hyperlinks: Vec::new(),
                        page_setup: None,
                        sheet_protection: None,
                        comments: vec![
                            Comment {
                                cell_ref: "C3".to_string(),
                                author: "Charlie".to_string(),
                                text: "Note on C3".to_string(),
                            },
                        ],
                        tab_color: None,
                        tables: Vec::new(),
                        header_footer: None,
                        outline_properties: None,
                        sparkline_groups: Vec::new(),
                        charts: Vec::new(),
                        pivot_tables: Vec::new(),
                        threaded_comments: Vec::new(),
                        slicers: Vec::new(),
                        timelines: Vec::new(),
                        preserved_extensions: Vec::new(),
                    },
                },
            ],
            date_system: DateSystem::Date1900,
            styles: Styles::default_styles(),
            defined_names: Vec::new(),
            shared_strings: None,
            doc_properties: None,
            theme_colors: None,
            calc_chain: Vec::new(),
            workbook_views: Vec::new(),
            protection: None,
            pivot_caches: Vec::new(),
            pivot_cache_records: Vec::new(),
            persons: Vec::new(),
            slicer_caches: Vec::new(),
            timeline_caches: Vec::new(),
            preserved_entries: std::collections::BTreeMap::new(),
        };

        let xlsx1 = crate::writer::write_xlsx(&wb1).expect("write_xlsx should succeed");
        let wb2 = read_xlsx(&xlsx1).expect("read_xlsx should succeed");

        // First sheet has no comments.
        assert!(wb2.sheets[0].worksheet.comments.is_empty());

        // Second sheet has one comment.
        assert_eq!(wb2.sheets[1].worksheet.comments.len(), 1);
        assert_eq!(wb2.sheets[1].worksheet.comments[0].cell_ref, "C3");
        assert_eq!(wb2.sheets[1].worksheet.comments[0].author, "Charlie");
        assert_eq!(wb2.sheets[1].worksheet.comments[0].text, "Note on C3");
    }

    #[test]
    fn test_resolve_relative_path() {
        assert_eq!(
            resolve_relative_path("xl/worksheets", "../comments1.xml"),
            "xl/comments1.xml"
        );
        assert_eq!(
            resolve_relative_path("xl/worksheets", "comments1.xml"),
            "xl/worksheets/comments1.xml"
        );
        assert_eq!(
            resolve_relative_path("xl/worksheets/sub", "../../comments1.xml"),
            "xl/comments1.xml"
        );
    }
}
