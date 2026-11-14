//! Full XLSX write orchestrator.
//!
//! Assembles all XML parts from a [`WorkbookData`] struct and zips them into a
//! complete `.xlsx` file.

use core::hint::cold_path;
use std::borrow::Cow;
use std::collections::{HashMap, HashSet};

use log::{debug, trace, warn};
use quick_xml::Reader;
use quick_xml::events::{BytesStart, Event};

use crate::ooxml::{
    calc_chain::{self, CalcChainEntry},
    charts::ChartAnchor,
    comments,
    content_types::{
        CT_CALC_CHAIN, CT_CHART, CT_CHART_COLOR_STYLE, CT_CHART_EX, CT_CHART_SHAPES,
        CT_CHART_STYLE, CT_COMMENTS, CT_CORE_PROPERTIES, CT_CUSTOM_PROPERTIES, CT_CUSTOM_XML_PROPS,
        CT_DOCX_PACKAGE, CT_DRAWING, CT_EXTENDED_PROPERTIES, CT_EXTERNAL_LINK, CT_HD_PHOTO,
        CT_OLE_OBJECT, CT_PERSONS, CT_PIVOT_CACHE_DEF, CT_PIVOT_CACHE_REC, CT_PIVOT_TABLE,
        CT_PPTX_PACKAGE, CT_PRINTER_SETTINGS, CT_SHEET_METADATA, CT_SLICER, CT_SLICER_CACHE,
        CT_TABLE, CT_THEME, CT_THREADED_COMMENTS, CT_TIMELINE, CT_TIMELINE_CACHE, CT_VML_DRAWING,
        CT_XLSX_PACKAGE, CT_XML, ContentTypes,
    },
    pivot_table::PivotCacheDefinitionData,
    relationships::{
        REL_CALC_CHAIN, REL_CHART, REL_CHART_COLOR_STYLE, REL_CHART_STYLE, REL_CHART_USER_SHAPES,
        REL_COMMENTS, REL_CORE_PROPERTIES, REL_CUSTOM_PROPERTIES, REL_CUSTOM_XML, REL_DRAWING,
        REL_EXTENDED_PROPERTIES, REL_EXTERNAL_LINK, REL_HYPERLINK, REL_PERSONS,
        REL_PIVOT_CACHE_DEF, REL_PIVOT_CACHE_REC, REL_PIVOT_TABLE, REL_SHEET_METADATA, REL_SLICER,
        REL_SLICER_CACHE, REL_TABLE, REL_THEME, REL_THREADED_COMMENTS, REL_TIMELINE,
        REL_TIMELINE_CACHE, REL_VML_DRAWING, Relationship, Relationships, make_rid, rels_path,
        resolve_target,
    },
    shared_strings::SharedStringTableBuilder,
    slicers,
    theme::ThemeColors,
    threaded_comments, timelines, vml,
    workbook::{PivotCacheRef, SheetInfo, SheetState, WorkbookXml},
    worksheet::{Cell, CellType},
    xml::root_namespaces,
};
use crate::zip::writer::{ZipEntry, write_zip};
use crate::{ModernXlsxError, Result, WorkbookData};

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

/// A worksheet part's ZIP path: the folder worksheet relationship targets are
/// resolved against (the file name does not matter).
const SHEET_PART: &str = "xl/worksheets/sheet.xml";

/// The theme part's ZIP path, preserved or generated.
const THEME_PATH: &str = "xl/theme/theme1.xml";

/// The content types part's ZIP path; always generated.
const CONTENT_TYPES_PATH: &str = "[Content_Types].xml";

/// The ZIP path prefix of external link parts (`externalLink{n}.xml`).
const EXTERNAL_LINK_PREFIX: &str = "xl/externalLinks/externalLink";

/// Whether `path` names an external link part.
fn is_external_link(path: &str) -> bool {
    path.starts_with(EXTERNAL_LINK_PREFIX) && path.ends_with(".xml")
}

/// The custom file properties part's ZIP path.
const CUSTOM_PROPERTIES_PATH: &str = "docProps/custom.xml";

/// Whether `path` names a custom XML data part (`customXml/item{n}.xml`):
/// not its properties part (`itemProps{n}.xml`) nor a `.rels`.
fn is_custom_xml_item(path: &str) -> bool {
    path.starts_with("customXml/item")
        && path.ends_with(".xml")
        && !path.contains("Props")
        && !path.contains("/_rels/")
}

/// Whether `path` (ASCII case ignored) can be a sheet drawing part: an
/// `.xml` part under `xl/drawings/` that is not a `.rels`.
fn is_sheet_drawing_part(path: &str) -> bool {
    let path = path.to_ascii_lowercase();
    path.starts_with("xl/drawings/") && path.ends_with(".xml") && !path.contains("/_rels/")
}

/// The `cacheId` each pivot cache is written under, in `pivot_caches` order:
/// its own `cache_id`, else its index; and the index of the cache each id
/// names, so a table finds its cache in one lookup ([`pivot_cache`]).
///
/// # Errors
///
/// `InvalidWorkbook` when two caches share an id, or a pivot table's
/// `cacheId` names no cache.
fn pivot_cache_ids(workbook: &WorkbookData) -> Result<(Vec<u32>, HashMap<u32, usize>)> {
    let ids: Vec<u32> = workbook
        .pivot_caches
        .iter()
        .enumerate()
        .map(|(i, cache)| {
            cache
                .cache_id
                .unwrap_or_else(|| u32::try_from(i).unwrap_or(u32::MAX))
        })
        .collect();
    let mut index = HashMap::with_capacity(ids.len());
    for (i, &id) in ids.iter().enumerate() {
        if index.insert(id, i).is_some() {
            cold_path();
            return Err(ModernXlsxError::InvalidWorkbook(format!(
                "Write failed: two pivot caches have cacheId {id} — give each cache in pivotCaches its own cacheId (an unset cacheId is the cache's index)"
            )));
        }
    }
    for sheet in &workbook.sheets {
        for table in &sheet.worksheet.pivot_tables {
            if !index.contains_key(&table.cache_id) {
                cold_path();
                return Err(ModernXlsxError::InvalidWorkbook(format!(
                    "Write failed: pivot table '{}' in sheet '{}' has cacheId {}, which no pivot cache has — add its cache to pivotCaches (its cacheId, or its index when cacheId is unset)",
                    table.name, sheet.name, table.cache_id
                )));
            }
        }
    }
    Ok((ids, index))
}

/// The pivot cache written under `cache_id`, given the index by id from
/// [`pivot_cache_ids`].
fn pivot_cache<'w>(
    workbook: &'w WorkbookData,
    index: &HashMap<u32, usize>,
    cache_id: u32,
) -> Option<&'w PivotCacheDefinitionData> {
    index
        .get(&cache_id)
        .and_then(|&i| workbook.pivot_caches.get(i))
}

/// Counts what the writer adds to every pivot table from its cache before
/// any part is written, so a workbook past the limits fails at once rather
/// than after writing the tables before the one that crosses them.
///
/// # Errors
///
/// `SizeLimit` as for `PivotTableData::count_added_items`.
fn check_pivot_additions(workbook: &WorkbookData, index: &HashMap<u32, usize>) -> Result<()> {
    let mut added = 0;
    for sheet in &workbook.sheets {
        for table in &sheet.worksheet.pivot_tables {
            table.count_added_items(pivot_cache(workbook, index, table.cache_id), &mut added)?;
        }
    }
    Ok(())
}

/// Checks that every colour of a theme to be generated is 6 hex digits, the
/// only form `generate_theme_xml` writes into `srgbClr`/`sysClr` values.
///
/// # Errors
///
/// `InvalidWorkbook` naming the first bad slot.
fn check_theme_colors(colors: &ThemeColors) -> Result<()> {
    let slots = [
        ("dk1", &colors.dk1),
        ("lt1", &colors.lt1),
        ("dk2", &colors.dk2),
        ("lt2", &colors.lt2),
        ("accent1", &colors.accent1),
        ("accent2", &colors.accent2),
        ("accent3", &colors.accent3),
        ("accent4", &colors.accent4),
        ("accent5", &colors.accent5),
        ("accent6", &colors.accent6),
        ("hlink", &colors.hlink),
        ("folHlink", &colors.fol_hlink),
    ];
    for (slot, value) in slots {
        if value.len() != 6 || !value.bytes().all(|b| b.is_ascii_hexdigit()) {
            cold_path();
            return Err(ModernXlsxError::InvalidWorkbook(format!(
                "Write failed: theme colour {slot} is '{value}', not 6 hex digits (RRGGBB) — fix themeColors or keep the workbook's own xl/theme/theme1.xml"
            )));
        }
    }
    Ok(())
}

/// The external link parts to reference from workbook.xml, in order: the
/// preserved ones `external_references` lists, then the other preserved
/// ones by their number.
fn external_link_order(workbook: &WorkbookData) -> Vec<&str> {
    let mut order: Vec<&str> = Vec::new();
    // The paths in `order`: both loops below look up every candidate once, and
    // both counts come from the file.
    let mut seen: HashSet<&str> = HashSet::new();
    for path in &workbook.external_references {
        if !workbook.preserved_entries.contains_key(path) {
            warn!("external reference {path} is not a preserved entry; skipped");
        } else if seen.insert(path.as_str()) {
            order.push(path);
        }
    }
    let mut rest: Vec<&str> = workbook
        .preserved_entries
        .keys()
        .map(String::as_str)
        .filter(|path| is_external_link(path) && !seen.contains(path))
        .collect();
    rest.sort_by_key(|path| {
        let number = path[EXTERNAL_LINK_PREFIX.len()..path.len() - ".xml".len()]
            .parse::<u64>()
            .unwrap_or(u64::MAX);
        (number, *path)
    });
    order.extend(rest);
    order
}

/// Points the sheet's relationship of `rel_type` at `target`: the first one
/// keeps its Id (a preserved worksheet `extLst` may refer to it) and any
/// further one is removed; with none, one is added under the next free rId.
fn upsert_typed_rel(ws_rels: &mut Relationships, rel_type: &'static str, target: String) {
    let existing = ws_rels
        .find_by_type(rel_type)
        .next()
        .map(|rel| rel.id.clone());
    ws_rels.retain(|rel| rel.rel_type != rel_type);
    let r_id = existing.unwrap_or_else(|| make_rid(ws_rels.next_r_id()));
    ws_rels.add(r_id, rel_type, target);
}

/// Drops the sheet's relationships of `rel_type` whose part the reader
/// consumed: the writer writes the model's parts of that type, if any, under
/// relationships of its own. One whose target is still a preserved entry (a
/// part the reader kept verbatim) stays.
fn drop_unwritten_typed_rels(
    ws_rels: &mut Relationships,
    rel_type: &str,
    preserved: &std::collections::BTreeMap<String, Vec<u8>>,
) {
    ws_rels.retain(|rel| {
        rel.rel_type != rel_type || preserved.contains_key(&resolve_target(SHEET_PART, &rel.target))
    });
}

/// What a VML drawing holding note or form-control shapes has in each shape
/// (`<x:ClientData ObjectType="Note">` and the like); a VML drawing of header
/// and footer pictures has none.
const VML_CLIENT_DATA: &[u8] = b"ClientData";

/// Whether `path` ends in `.vml`, ASCII case ignored: the name the vml
/// Default of `[Content_Types].xml` covers.
fn has_vml_extension(path: &str) -> bool {
    path.rsplit_once('.')
        .is_some_and(|(_, ext)| ext.eq_ignore_ascii_case("vml"))
}

/// Whether the preserved part at `path` is a VML drawing: its name ends in
/// `.vml` ([`has_vml_extension`]), or the source package declared it with
/// the VML drawing content type (`preserved_content_types`; a `.rels` is
/// always a relationships part). Only such a part is named by
/// `<legacyDrawing>` or `<legacyDrawingHF>` and merged with a sheet's notes,
/// whatever a vmlDrawing relationship targets: a `.rels` part, or any part
/// neither named `.vml` nor typed as a VML drawing, is never rewritten as a
/// drawing, wherever it sits.
fn is_vml_drawing(path: &str, workbook: &WorkbookData) -> bool {
    has_vml_extension(path)
        || (!path.ends_with(".rels")
            && workbook
                .preserved_content_types
                .get(path)
                .is_some_and(|content_type| content_type == CT_VML_DRAWING))
}

/// The `r:id`s a sheet names its preserved VML drawings by: `<legacyDrawing>`
/// for the first one with cell-anchored shapes ([`VML_CLIENT_DATA`]: notes or
/// form controls) and `<legacyDrawingHF>` for the first one without; the
/// search stops once both are found. A relationship counts when its target
/// is a VML drawing ([`is_vml_drawing`]) that is a preserved entry no
/// generated part supersedes, or one an earlier sheet merged its notes into
/// (`merged`: the drawing stays at its path, so every sheet that shares it
/// keeps naming it).
///
/// `cell_anchored` holds, by path, whether a drawing's preserved bytes hold
/// [`VML_CLIENT_DATA`]: each drawing is searched once per write, however
/// many relationships and sheets name it.
///
/// For a sheet with comments, the writer then merges the notes into the
/// `<legacyDrawing>` one, or generates a note drawing when there is none
/// (see [`crate::ooxml::vml`]).
fn legacy_drawing_r_ids<'r>(
    rels: &'r Relationships,
    workbook: &WorkbookData,
    generated: &HashSet<String>,
    merged: &HashSet<String>,
    cell_anchored: &mut HashMap<String, bool>,
) -> (Option<&'r str>, Option<&'r str>) {
    let (mut legacy, mut header_footer) = (None, None);
    for rel in rels.find_by_type(REL_VML_DRAWING) {
        if legacy.is_some() && header_footer.is_some() {
            break;
        }
        let target = resolve_target(SHEET_PART, &rel.target);
        if !is_vml_drawing(&target, workbook) {
            continue;
        }
        let Some(vml) = workbook
            .preserved_entries
            .get(&target)
            .filter(|_| !generated.contains(&target) || merged.contains(&target))
        else {
            continue;
        };
        let slot = if *cell_anchored.entry(target).or_insert_with(|| {
            vml.windows(VML_CLIENT_DATA.len())
                .any(|window| window == VML_CLIENT_DATA)
        }) {
            &mut legacy
        } else {
            &mut header_footer
        };
        slot.get_or_insert(rel.id.as_str());
    }
    (legacy, header_footer)
}

/// The metadata part's ZIP path.
const METADATA_PATH: &str = "xl/metadata.xml";

/// The metadata part generated for workbooks whose cells carry `cm="1"`: one
/// XLDAPR (dynamic array) record, byte for byte the part Microsoft Excel 16.0
/// (build 20527) saved for a workbook holding `=SEQUENCE(3)` in A1 on
/// 2026-11-11.
const DYNAMIC_ARRAY_METADATA_XML: &str = concat!(
    r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#,
    "\r\n",
    r#"<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:xda="http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray">"#,
    r#"<metadataTypes count="1"><metadataType name="XLDAPR" minSupportedVersion="120000" copy="1" pasteAll="1" pasteValues="1" merge="1" splitFirst="1" rowColShift="1" clearFormats="1" clearComments="1" assign="1" coerce="1" cellMeta="1"/></metadataTypes>"#,
    r#"<futureMetadata name="XLDAPR" count="1"><bk><extLst><ext uri="{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}"><xda:dynamicArrayProperties fDynamic="1" fCollapsed="0"/></ext></extLst></bk></futureMetadata>"#,
    r#"<cellMetadata count="1"><bk><rc t="1" v="0"/></bk></cellMetadata>"#,
    "</metadata>",
);

/// Whether a cell is written with `cm` on `<c>` (see the worksheet writer).
fn writes_cell_metadata(cell: &Cell) -> bool {
    cell.cell_metadata.is_some() || cell.dynamic_array == Some(true)
}

/// The metadata part to add as a generated entry, if any: none when the
/// workbook keeps its own `xl/metadata.xml` or no cell is written with `cm`,
/// else the one-record dynamic-array part.
///
/// # Errors
///
/// `InvalidCellValue` when the part is generated and a cell's
/// `cell_metadata` is not 1, the only record the generated part holds.
fn generated_metadata_part(workbook: &WorkbookData) -> Result<Option<&'static str>> {
    if workbook.preserved_entries.contains_key(METADATA_PATH) {
        return Ok(None);
    }
    let mut any = false;
    for sheet in &workbook.sheets {
        for cell in sheet.worksheet.rows.iter().flat_map(|r| &r.cells) {
            if let Some(cm) = cell.cell_metadata.filter(|&cm| cm != 1) {
                cold_path();
                return Err(ModernXlsxError::InvalidCellValue(format!(
                    "Write failed: cell '{}' in sheet '{}' has cellMetadata {cm}, but the workbook has no xl/metadata.xml to hold it — only 1 (a dynamic array) can be generated",
                    cell.reference, sheet.name
                )));
            }
            any |= writes_cell_metadata(cell);
        }
    }
    Ok(any.then_some(DYNAMIC_ARRAY_METADATA_XML))
}

/// The calcChain entries to write: those whose `sheet_id` is a sheet position
/// (`1..=sheets.len()`) and whose cell gets an `<f>` ([`Cell::writes_formula`]
/// against the shared-formula masters of its sheet, the predicate the
/// worksheet writer uses).
/// Excel does not open a package whose linked calcChain names a cell without
/// a formula, so the rest are dropped. Borrowed when every entry is kept.
fn written_calc_chain(workbook: &WorkbookData) -> Cow<'_, [CalcChainEntry]> {
    let sheets = &workbook.sheets;
    // Formula cell references per sheet position, built on first use.
    let mut formula_cells: Vec<Option<HashSet<&str>>> = vec![None; sheets.len()];
    let mut keep = |entry: &CalcChainEntry| {
        let Some(pos) = usize::try_from(entry.sheet_id)
            .ok()
            .and_then(|id| id.checked_sub(1))
            .filter(|&pos| pos < sheets.len())
        else {
            return false;
        };
        formula_cells[pos]
            .get_or_insert_with(|| {
                let worksheet = &sheets[pos].worksheet;
                let shared_masters = worksheet.shared_formula_masters();
                worksheet
                    .rows
                    .iter()
                    .flat_map(|row| &row.cells)
                    .filter(|cell| cell.writes_formula(&shared_masters))
                    .map(|cell| cell.reference.as_str())
                    .collect()
            })
            .contains(entry.cell_ref.as_str())
    };
    let chain = workbook.calc_chain.as_slice();
    let Some(first) = chain.iter().position(|entry| !keep(entry)) else {
        return Cow::Borrowed(chain);
    };
    let mut kept = chain[..first].to_vec();
    kept.extend(
        chain[first + 1..]
            .iter()
            .filter(|entry| keep(entry))
            .cloned(),
    );
    debug!(
        "calcChain: {} of {} entries name no formula cell; dropped",
        chain.len() - kept.len(),
        chain.len()
    );
    Cow::Owned(kept)
}

// ---------------------------------------------------------------------------
// Public API
// ---------------------------------------------------------------------------

/// Write a complete XLSX file from a [`WorkbookData`] struct.
///
/// For cells whose [`CellType`] is [`CellType::SharedString`], the `value`
/// field must contain the **actual string text** (not an SST index). This
/// function builds the shared string table, maps each string to an index,
/// and writes the correct index into the worksheet XML.
///
/// Returns the raw bytes of the resulting ZIP (`.xlsx`) file.
pub fn write_xlsx(workbook: &WorkbookData) -> Result<Vec<u8>> {
    if workbook.sheets.is_empty() {
        cold_path();
        return Err(ModernXlsxError::InvalidWorkbook(
            "Write failed: workbook must contain at least one sheet — add a sheet before calling write()".into(),
        ));
    }

    let sheet_count = workbook.sheets.len();
    let (pivot_cache_ids, pivot_cache_index) = pivot_cache_ids(workbook)?;
    check_pivot_additions(workbook, &pivot_cache_index)?;
    crate::validate::check_writable(workbook)?;
    debug!("collecting shared strings from {sheet_count} sheets");

    // 1. Build the SharedStringTable from all string cells.
    let mut sst_builder = SharedStringTableBuilder::new();

    for sheet in &workbook.sheets {
        for row in &sheet.worksheet.rows {
            for cell in &row.cells {
                if cell.cell_type == CellType::SharedString {
                    let Some(ref val) = cell.value else {
                        cold_path();
                        return Err(ModernXlsxError::InvalidCellValue(format!(
                            "SharedString cell '{}' in sheet '{}' has no value — set cell.value to the string text",
                            cell.reference, sheet.name
                        )));
                    };
                    if let Some(ref runs) = cell.rich_text {
                        sst_builder.insert_rich(val, runs.as_slice());
                    } else {
                        sst_builder.insert(val);
                    }
                }
            }
        }
    }

    debug!("SST built with {} unique strings", sst_builder.len());

    // 2. The content types and the package and workbook relationships. The
    // workbook relationships grow with every workbook-level part below;
    // workbook.xml and its .rels are written in step 5h, once every
    // relationship Id they list exists.
    let mut content_types = ContentTypes::for_basic_workbook(sheet_count);
    let mut root_rels = Relationships::root_rels();
    let mut wb_rels = Relationships::workbook_rels(sheet_count);

    // 3. Assemble ZIP entries, starting with the shared strings and styles.
    // [Content_Types].xml is pushed last (step 7), once every part below has
    // added its Override or Default.
    let preserved_count = workbook.preserved_entries.len();
    let mut entries = Vec::with_capacity(11 + sheet_count * 3 + preserved_count);

    entries.push(ZipEntry {
        name: "xl/sharedStrings.xml".to_string(),
        data: sst_builder.to_xml()?,
    });
    entries.push(ZipEntry {
        name: "xl/styles.xml".to_string(),
        data: workbook.styles.to_xml()?,
    });

    // 3b. Theme: a preserved xl/theme/theme1.xml keeps its bytes (step 6);
    // otherwise the writer emits a generated Office theme
    // (`generate_theme_xml`) in the workbook's theme colours, or the default
    // ones when it has none. Either way it gets its workbook relationship and
    // Override.
    if !workbook.preserved_entries.contains_key(THEME_PATH) {
        if let Some(colors) = &workbook.theme_colors {
            check_theme_colors(colors)?;
        }
        entries.push(ZipEntry {
            name: THEME_PATH.to_owned(),
            data: crate::ooxml::theme::generate_theme_xml(workbook.theme_colors.as_ref()),
        });
    }
    wb_rels.add(
        make_rid(wb_rels.next_r_id()),
        REL_THEME,
        target_from_workbook(THEME_PATH).to_owned(),
    );
    content_types.add_override(format!("/{THEME_PATH}"), CT_THEME);

    // 4a. Write document properties if present, each with its package
    // relationship and Override.
    if let Some(ref props) = workbook.doc_properties {
        if props.has_core() {
            entries.push(ZipEntry {
                name: "docProps/core.xml".to_string(),
                data: props.to_core_xml()?,
            });
            root_rels.add(
                make_rid(root_rels.next_r_id()),
                REL_CORE_PROPERTIES,
                "docProps/core.xml",
            );
            content_types.add_override("/docProps/core.xml", CT_CORE_PROPERTIES);
        }
        if props.has_app() {
            entries.push(ZipEntry {
                name: "docProps/app.xml".to_string(),
                data: props.to_app_xml()?,
            });
            root_rels.add(
                make_rid(root_rels.next_r_id()),
                REL_EXTENDED_PROPERTIES,
                "docProps/app.xml",
            );
            content_types.add_override("/docProps/app.xml", CT_EXTENDED_PROPERTIES);
        }
    }
    // A preserved docProps/custom.xml (step 6) gets its package relationship
    // and Override.
    if workbook
        .preserved_entries
        .contains_key(CUSTOM_PROPERTIES_PATH)
    {
        root_rels.add(
            make_rid(root_rels.next_r_id()),
            REL_CUSTOM_PROPERTIES,
            CUSTOM_PROPERTIES_PATH,
        );
        content_types.add_override(format!("/{CUSTOM_PROPERTIES_PATH}"), CT_CUSTOM_PROPERTIES);
    }
    entries.push(ZipEntry {
        name: "_rels/.rels".to_string(),
        data: root_rels.to_xml()?,
    });

    // 4b. Write the calculation chain, keeping only entries on a formula cell
    // (see `written_calc_chain`); with none left there is no part, rel or
    // Override.
    let chain = written_calc_chain(workbook);
    if !chain.is_empty() {
        entries.push(ZipEntry {
            name: "xl/calcChain.xml".to_string(),
            data: calc_chain::to_xml(&chain)?,
        });
        wb_rels.add(
            make_rid(wb_rels.next_r_id()),
            REL_CALC_CHAIN,
            "calcChain.xml",
        );
        content_types.add_override("/xl/calcChain.xml", CT_CALC_CHAIN);
    }

    // 5. Write each worksheet, looking up its shared-string indices in the SST
    // builder while its XML is serialized (the worksheet is not cloned), with
    // its tables, charts and drawing, comments, threaded comments, pivot
    // tables, slicers, timelines and .rels.
    // Generated part names are allocated around the preserved ones (see
    // `PartNames`).
    let mut names = PartNames::new(workbook.preserved_entries.keys());
    // Generated part paths that supersede a preserved entry at the same path
    // (step 6 skips those entries).
    let mut generated_paths: HashSet<String> = HashSet::with_capacity(sheet_count * 2);
    // The preserved VML drawings a sheet merged its notes into: written once,
    // at their own paths (also in `generated_paths`).
    let mut merged_vml: HashSet<String> = HashSet::new();
    // Whether each preserved VML drawing holds cell-anchored shapes, by path
    // (see `legacy_drawing_r_ids`).
    let mut vml_cell_anchored: HashMap<String, bool> = HashMap::new();
    // The preserved VML drawings handed to the note merge: each one at most
    // once per write, for the first sheet with comments that names it.
    let mut examined_vml: HashSet<String> = HashSet::new();
    // The parts the chart .rels name (user shapes, chart styles, colour
    // styles), with their content types (declared in step 6b).
    let mut chart_part_types: HashMap<String, &'static str> = HashMap::new();
    // Each written pivot table's path and cacheId; its .rels to the cache
    // definition is written in step 5b, once the caches have their paths.
    let mut pivot_table_paths: Vec<(String, u32)> = Vec::new();
    // The cacheIds of the caches a pivot table built in code uses: step 5b
    // writes them with `refreshOnLoad`.
    let mut refresh_on_load: HashSet<u32> = HashSet::new();
    // The items and pivot fields added to the pivot tables written so far
    // (`check_pivot_additions` has checked the total).
    let mut pivot_added: usize = 0;
    // The first `o:idmap` block no generated note drawing has taken yet.
    let mut next_vml_block: u32 = 1;
    // The note shapes and thread placeholders generated so far, each counted
    // before it is built (`vml::MAX_GENERATED_NOTE_SHAPES`).
    let mut generated_notes: usize = 0;

    for (i, sheet) in workbook.sheets.iter().enumerate() {
        let sheet_num = i + 1;
        // The legacy comments the sheet's comments part holds: its own, then
        // the `tc={root id}` placeholder Excel needs to show each thread
        // whose cell has none (`threaded_comments::placeholder_comments`).
        // They exist only in the written file: the model is not changed, and
        // the sheet's comments are copied only when a placeholder is added.
        let placeholders = threaded_comments::placeholder_comments(
            &sheet.worksheet.threaded_comments,
            &sheet.worksheet.comments,
            &mut generated_notes,
        )?;
        let placeholder_count = placeholders.len();
        let sheet_comments: Cow<'_, [comments::Comment]> = if placeholders.is_empty() {
            Cow::Borrowed(&sheet.worksheet.comments)
        } else {
            let mut all = sheet.worksheet.comments.clone();
            all.extend(placeholders);
            Cow::Owned(all)
        };
        let has_comments = !sheet_comments.is_empty();
        let has_tables = !sheet.worksheet.tables.is_empty();
        let has_charts = !sheet.worksheet.charts.is_empty();
        let has_pivot_tables = !sheet.worksheet.pivot_tables.is_empty();
        let has_threaded_comments = !sheet.worksheet.threaded_comments.is_empty();
        let has_slicers = !sheet.worksheet.slicers.is_empty();
        let has_timelines = !sheet.worksheet.timelines.is_empty();
        let has_external_links = sheet
            .worksheet
            .hyperlinks
            .iter()
            .any(|hl| hl.external_target().is_some());

        // Build (or merge) the worksheet .rels when we need to add relationships.
        let sheet_rels_path = format!("xl/worksheets/_rels/sheet{sheet_num}.xml.rels");
        // Parsed once: moved into `ws_rels` when the sheet needs generated rels;
        // otherwise kept here as the only place a preserved drawing relationship shows.
        // A malformed preserved .rels is not merged (a warning is logged): without
        // generated rels its verbatim copy is written; otherwise the generated rels
        // supersede it.
        let preserved_rels = match workbook
            .preserved_entries
            .get(&sheet_rels_path)
            .map(|bytes| Relationships::parse(bytes))
        {
            Some(Ok(rels)) => Some(rels),
            Some(Err(e)) => {
                cold_path();
                warn!("{sheet_rels_path}: {e}; preserved rels not merged");
                None
            }
            None => None,
        };
        // Hyperlink rels are regenerated from the model, so preserved ones must
        // go even when no hyperlink is left.
        let has_preserved_link_rels = preserved_rels
            .as_ref()
            .is_some_and(|rels| rels.find_by_type(REL_HYPERLINK).next().is_some());
        // Comments, threaded comments, slicers and timelines are written as one
        // part each, under this writer's own names: a preserved rel of such a
        // type is repointed at the written part or, when the sheet has none of
        // it now, dropped unless it still names a preserved part.
        let typed_parts = [
            (REL_COMMENTS, has_comments),
            (REL_THREADED_COMMENTS, has_threaded_comments),
            (REL_SLICER, has_slicers),
            (REL_TIMELINE, has_timelines),
        ];
        // Tables and pivot tables are written under new relationships, one per
        // part: a preserved rel of such a type is dropped unless it still names
        // a preserved part, whether the sheet has any of them now or not.
        let regenerated_parts = [(REL_TABLE, has_tables), (REL_PIVOT_TABLE, has_pivot_tables)];
        let has_stale_typed_rels = preserved_rels.as_ref().is_some_and(|rels| {
            typed_parts
                .iter()
                .chain(&regenerated_parts)
                .any(|&(rel_type, has)| !has && rels.find_by_type(rel_type).next().is_some())
        });
        let needs_rels = has_comments
            || has_tables
            || has_charts
            || has_pivot_tables
            || has_threaded_comments
            || has_slicers
            || has_timelines
            || has_external_links
            || has_preserved_link_rels
            || has_stale_typed_rels;
        let (mut ws_rels, preserved_rels) = if needs_rels {
            (preserved_rels.unwrap_or_default(), None)
        } else {
            (Relationships::new(), preserved_rels)
        };

        for (rel_type, has) in typed_parts {
            if !has {
                drop_unwritten_typed_rels(&mut ws_rels, rel_type, &workbook.preserved_entries);
            }
        }
        for (rel_type, _) in regenerated_parts {
            drop_unwritten_typed_rels(&mut ws_rels, rel_type, &workbook.preserved_entries);
        }

        // --- External hyperlinks: one TargetMode="External" rel each ---
        ws_rels.retain(|rel| rel.rel_type != REL_HYPERLINK);
        let hyperlink_r_ids: Vec<Option<String>> = if has_external_links {
            sheet
                .worksheet
                .hyperlinks
                .iter()
                .map(|hl| {
                    hl.external_target()
                        .map(|target| ws_rels.add_external(REL_HYPERLINK, target.to_owned()))
                })
                .collect()
        } else {
            Vec::new()
        };

        // --- Tables ---
        let mut table_r_ids: Vec<String> = Vec::with_capacity(sheet.worksheet.tables.len());
        if has_tables {
            debug!(
                "writing {} tables for sheet {}",
                sheet.worksheet.tables.len(),
                sheet_num
            );
            for table in &sheet.worksheet.tables {
                let table_xml = table.to_xml()?;
                let table_path = names.alloc("xl/tables/table", ".xml");

                content_types.add_override(format!("/{table_path}"), CT_TABLE);

                let rid_num = ws_rels.next_r_id();
                let rid = make_rid(rid_num);
                ws_rels.add(rid.clone(), REL_TABLE, target_from_xl_folder(&table_path));
                table_r_ids.push(rid);

                generated_paths.insert(table_path.clone());
                entries.push(ZipEntry {
                    name: table_path,
                    data: table_xml,
                });
            }
        }

        // --- Charts ---
        let mut drawing_r_id_str: Option<String> = None;
        if has_charts {
            debug!(
                "writing {} charts for sheet {}",
                sheet.worksheet.charts.len(),
                sheet_num
            );

            // The sheet's own drawing relationship (from its preserved .rels) says
            // where its drawing lives: that part (an `addImage` drawing, or the
            // drawing a read consumed) is regenerated in place under the same rId.
            // A target that is not a sheet drawing part (`is_sheet_drawing_part`),
            // or that another sheet's drawing already took, is moved to a free
            // name under the same rId; a sheet without one gets a free name and a
            // new rel.
            let existing = ws_rels
                .find_by_type(REL_DRAWING)
                .next()
                .map(|rel| (rel.id.clone(), resolve_target(SHEET_PART, &rel.target)));
            let (drawing_r_id, drawing_path) = match existing {
                Some((r_id, path))
                    if is_sheet_drawing_part(&path) && !generated_paths.contains(&path) =>
                {
                    names.take(&path);
                    (r_id, path)
                }
                Some((r_id, _)) => {
                    let path = names.alloc("xl/drawings/drawing", ".xml");
                    ws_rels.retain(|rel| rel.id != r_id);
                    ws_rels.add(r_id.clone(), REL_DRAWING, target_from_xl_folder(&path));
                    (r_id, path)
                }
                None => {
                    let path = names.alloc_preferring("xl/drawings/drawing", sheet_num, ".xml");
                    let r_id = make_rid(ws_rels.next_r_id());
                    ws_rels.add(r_id.clone(), REL_DRAWING, target_from_xl_folder(&path));
                    (r_id, path)
                }
            };
            let drawing_rels_path = rels_path(&drawing_path);

            // Create drawing .rels for this sheet's charts.
            let mut drawing_rels = Relationships::new();
            let mut chart_r_ids: Vec<String> = Vec::with_capacity(sheet.worksheet.charts.len());

            for (ci, wsc) in sheet.worksheet.charts.iter().enumerate() {
                // A chart with `raw_xml` comes back verbatim from `to_xml`.
                let chart_xml = wsc.chart.to_xml()?;
                let chart_path = names.alloc("xl/charts/chart", ".xml");

                content_types.add_override(format!("/{chart_path}"), CT_CHART);

                let rid = make_rid(ci + 1);
                drawing_rels.add(rid.clone(), REL_CHART, target_from_xl_folder(&chart_path));
                chart_r_ids.push(rid);

                // Its .rels goes beside it at the new path; the targets are relative
                // to xl/charts/, so they still resolve.
                if let Some(ref raw_rels) = wsc.chart.raw_rels {
                    collect_chart_part_types(
                        &chart_path,
                        raw_rels.as_bytes(),
                        &mut chart_part_types,
                    );
                    let chart_rels_path = rels_path(&chart_path);
                    generated_paths.insert(chart_rels_path.clone());
                    entries.push(ZipEntry {
                        name: chart_rels_path,
                        data: raw_rels.as_bytes().to_vec(),
                    });
                }

                generated_paths.insert(chart_path.clone());
                entries.push(ZipEntry {
                    name: chart_path,
                    data: chart_xml,
                });
            }

            // Generate drawing XML, then keep the pictures, shapes and other
            // anchors of the drawing it replaces (`addImage` stores pictures as a
            // preserved drawing at this path).
            let mut drawing_xml =
                ChartAnchor::generate_drawing_xml(&sheet.worksheet.charts, &chart_r_ids)?;
            if let Some(preserved) = workbook.preserved_entries.get(&drawing_path) {
                drawing_xml = merge_preserved_drawing(
                    drawing_xml,
                    &drawing_path,
                    preserved,
                    workbook
                        .preserved_entries
                        .get(&drawing_rels_path)
                        .map(Vec::as_slice),
                    &mut drawing_rels,
                )?;
            }

            content_types.add_override(format!("/{drawing_path}"), CT_DRAWING);

            entries.push(ZipEntry {
                name: drawing_path.clone(),
                data: drawing_xml,
            });
            entries.push(ZipEntry {
                name: drawing_rels_path.clone(),
                data: drawing_rels.to_xml()?,
            });

            // The preserved copies of both are superseded.
            generated_paths.insert(drawing_path);
            generated_paths.insert(drawing_rels_path);
            drawing_r_id_str = Some(drawing_r_id);
        }

        // --- Preserved-only drawing (image-only sheets) ---
        // `addImage` stores the drawing, its rels and the sheet rels as preserved
        // entries, so a chart-free sheet reaches this point with no drawing rId even
        // though its rels already point at one. Reference it from the worksheet XML,
        // but only when the target is a preserved entry, looked up by key: a
        // dangling rId makes Excel repair the file. A part generated at that
        // path supersedes the entry and belongs to another owner (for a
        // drawing, the earlier sheet the reader gave a shared drawing to), and
        // Excel rejects two sheets on one drawing, so it is not referenced.
        if drawing_r_id_str.is_none() {
            let drawing_rel = ws_rels.find_by_type(REL_DRAWING).next().or_else(|| {
                preserved_rels
                    .as_ref()
                    .and_then(|rels| rels.find_by_type(REL_DRAWING).next())
            });
            if let Some(rel) = drawing_rel {
                let target = resolve_target(SHEET_PART, &rel.target);
                if workbook.preserved_entries.contains_key(&target)
                    && !generated_paths.contains(&target)
                {
                    drawing_r_id_str = Some(rel.id.clone());
                }
            }
        }

        // --- VML drawings ---
        // The sheet's preserved VML drawings (notes and form controls, header
        // and footer pictures) and their rels are kept, and the sheet XML
        // names them: Excel rejects a sheet with notes and a <drawing> but no
        // <legacyDrawing>. A sheet with comments (thread placeholders
        // included) needs a note shape for each commented cell in the
        // <legacyDrawing> one: a preserved drawing is
        // merged with the comments (kept byte for byte when its notes match
        // them, or when it cannot be merged), and without one a note drawing
        // is generated. Only a VML drawing is named or merged
        // (`is_vml_drawing`). A drawing shared by several sheets goes to the
        // merge once, for the first sheet with comments that names it by
        // <legacyDrawing> (`examined_vml`); each sheet that shares it names
        // it, merged or as it is. The shapes added for legacy comments count
        // toward `vml::MAX_GENERATED_NOTE_SHAPES` before the drawing is built
        // (the shapes of thread placeholders were counted with them).
        let (legacy_drawing_r_id, legacy_drawing_hf_r_id) = legacy_drawing_r_ids(
            preserved_rels.as_ref().unwrap_or(&ws_rels),
            workbook,
            &generated_paths,
            &merged_vml,
            &mut vml_cell_anchored,
        );
        let legacy_drawing_hf_r_id = legacy_drawing_hf_r_id.map(str::to_owned);
        let mut legacy_drawing_r_id = legacy_drawing_r_id.map(str::to_owned);
        if has_comments {
            // With comments the sheet's rels are `ws_rels` (`needs_rels`).
            let preserved_vml = legacy_drawing_r_id
                .as_deref()
                .and_then(|r_id| ws_rels.get_by_id(r_id))
                .map(|rel| resolve_target(SHEET_PART, &rel.target));
            // The placeholders come last, each on a cell of its own that no
            // legacy comment names: the last `placeholder_count` cells. A
            // generated drawing has a shape for each legacy comment's cell,
            // counted as the cells are found; a merge counts its own.
            let legacy_count = sheet_comments.len() - placeholder_count;
            let cells = vml::note_cells(
                &sheet_comments,
                if preserved_vml.is_none() {
                    legacy_count
                } else {
                    0
                },
                &mut generated_notes,
            )?;
            if let Some(path) = preserved_vml {
                // A drawing an earlier sheet examined is named as it is,
                // merged or not.
                let merged = if examined_vml.insert(path.clone())
                    && let Some(vml) = workbook.preserved_entries.get(&path)
                {
                    vml::merge_note_drawing(
                        vml,
                        &path,
                        &cells,
                        placeholder_count,
                        &mut generated_notes,
                    )?
                } else {
                    None
                };
                if let Some(merged) = merged {
                    // A generated part loses its preserved Override (step
                    // 6c), and the vml Default covers only a `.vml` name.
                    if !has_vml_extension(&path) {
                        content_types.add_override(format!("/{path}"), CT_VML_DRAWING);
                    }
                    generated_paths.insert(path.clone());
                    merged_vml.insert(path.clone());
                    entries.push(ZipEntry {
                        name: path,
                        data: merged,
                    });
                }
            } else if !cells.is_empty() {
                let path = names.alloc_preferring("xl/drawings/vmlDrawing", sheet_num, ".vml");
                let block = u32::try_from(sheet_num)
                    .unwrap_or(u32::MAX)
                    .max(next_vml_block);
                let (vml, next_block) = vml::note_drawing(&cells, block);
                next_vml_block = next_block;
                let r_id = make_rid(ws_rels.next_r_id());
                ws_rels.add(r_id.clone(), REL_VML_DRAWING, target_from_xl_folder(&path));
                content_types.add_default("vml", CT_VML_DRAWING);
                generated_paths.insert(path.clone());
                entries.push(ZipEntry {
                    name: path,
                    data: vml,
                });
                legacy_drawing_r_id = Some(r_id);
            }
        }

        // --- Worksheet XML (must come after table/chart/hyperlink rIds are computed) ---
        let ws_xml = sheet.worksheet.to_xml_with_legacy_drawings(
            Some(&sst_builder),
            &table_r_ids,
            drawing_r_id_str.as_deref(),
            legacy_drawing_r_id.as_deref(),
            legacy_drawing_hf_r_id.as_deref(),
            &hyperlink_r_ids,
        )?;
        entries.push(ZipEntry {
            name: format!("xl/worksheets/sheet{sheet_num}.xml"),
            data: ws_xml,
        });

        // --- Comments (the sheet's own, then the thread placeholders) ---
        if has_comments {
            debug!(
                "writing {} comments ({} thread placeholders) for sheet {}",
                sheet_comments.len(),
                placeholder_count,
                sheet_num
            );
            let comments_xml = comments::write_comments(&sheet_comments)?;
            let comments_path = names.alloc_preferring("xl/comments", sheet_num, ".xml");

            content_types.add_override(format!("/{comments_path}"), CT_COMMENTS);

            upsert_typed_rel(
                &mut ws_rels,
                REL_COMMENTS,
                target_from_xl_folder(&comments_path),
            );

            generated_paths.insert(comments_path.clone());
            entries.push(ZipEntry {
                name: comments_path,
                data: comments_xml,
            });
        }

        // --- Threaded Comments ---
        if has_threaded_comments {
            debug!(
                "writing {} threaded comments for sheet {}",
                sheet.worksheet.threaded_comments.len(),
                sheet_num
            );
            let tc_xml =
                threaded_comments::write_threaded_comments(&sheet.worksheet.threaded_comments)?;
            let tc_path =
                names.alloc_preferring("xl/threadedComments/threadedComment", sheet_num, ".xml");

            content_types.add_override(format!("/{tc_path}"), CT_THREADED_COMMENTS);

            upsert_typed_rel(
                &mut ws_rels,
                REL_THREADED_COMMENTS,
                target_from_xl_folder(&tc_path),
            );

            generated_paths.insert(tc_path.clone());
            entries.push(ZipEntry {
                name: tc_path,
                data: tc_xml,
            });
        }

        // --- Pivot Tables ---
        if has_pivot_tables {
            debug!(
                "writing {} pivot tables for sheet {}",
                sheet.worksheet.pivot_tables.len(),
                sheet_num
            );
            for pt in &sheet.worksheet.pivot_tables {
                // The cache the table's cacheId names (`pivot_cache_ids`
                // checked that one does) completes a table built in code,
                // and refreshes on open so Excel lays that table out.
                let cache = pivot_cache(workbook, &pivot_cache_index, pt.cache_id);
                if pt.has_axis_field_without_items() {
                    refresh_on_load.insert(pt.cache_id);
                }
                let pt_xml = pt.to_xml_with_cache(cache, &mut pivot_added)?;
                let pt_path = names.alloc("xl/pivotTables/pivotTable", ".xml");

                content_types.add_override(format!("/{pt_path}"), CT_PIVOT_TABLE);

                let rid_num = ws_rels.next_r_id();
                ws_rels.add(
                    make_rid(rid_num),
                    REL_PIVOT_TABLE,
                    target_from_xl_folder(&pt_path),
                );

                generated_paths.insert(pt_path.clone());
                pivot_table_paths.push((pt_path.clone(), pt.cache_id));
                entries.push(ZipEntry {
                    name: pt_path,
                    data: pt_xml,
                });
            }
        }

        // --- Slicers ---
        if has_slicers {
            debug!(
                "writing {} slicers for sheet {}",
                sheet.worksheet.slicers.len(),
                sheet_num
            );
            let slicer_xml = slicers::write_slicers(&sheet.worksheet.slicers)?;
            let slicer_path = names.alloc("xl/slicers/slicer", ".xml");

            content_types.add_override(format!("/{slicer_path}"), CT_SLICER);

            upsert_typed_rel(
                &mut ws_rels,
                REL_SLICER,
                target_from_xl_folder(&slicer_path),
            );

            generated_paths.insert(slicer_path.clone());
            entries.push(ZipEntry {
                name: slicer_path,
                data: slicer_xml,
            });
        }

        // --- Timelines ---
        if has_timelines {
            debug!(
                "writing {} timelines for sheet {}",
                sheet.worksheet.timelines.len(),
                sheet_num
            );
            let tl_xml = timelines::write_timelines(&sheet.worksheet.timelines)?;
            let tl_path = names.alloc("xl/timelines/timeline", ".xml");

            content_types.add_override(format!("/{tl_path}"), CT_TIMELINE);

            upsert_typed_rel(&mut ws_rels, REL_TIMELINE, target_from_xl_folder(&tl_path));

            generated_paths.insert(tl_path.clone());
            entries.push(ZipEntry {
                name: tl_path,
                data: tl_xml,
            });
        }

        // --- Write worksheet .rels ---
        // The preserved copy is always superseded; a sheet left with no
        // relationship (its last hyperlink removed) gets no rels part.
        if needs_rels {
            generated_paths.insert(sheet_rels_path.clone());
            if !ws_rels.is_empty() {
                entries.push(ZipEntry {
                    name: sheet_rels_path,
                    data: ws_rels.to_xml()?,
                });
            }
        }
    }

    // 5b. Write workbook-level pivot caches, each listed in workbook.xml
    // `<pivotCaches>` under its cacheId and workbook relationship. A cache
    // with records names them by `r:id` (in its .rels, `rId1` unless a kept
    // relationship holds it); one a table built in code uses refreshes on
    // load.
    let mut pivot_cache_refs: Vec<PivotCacheRef> = Vec::with_capacity(pivot_cache_ids.len());
    let mut cache_paths: HashMap<u32, String> = HashMap::with_capacity(pivot_cache_ids.len());
    for (ci, (cache, &cache_id)) in workbook
        .pivot_caches
        .iter()
        .zip(&pivot_cache_ids)
        .enumerate()
    {
        let part_num = ci + 1;
        let records = workbook.pivot_cache_records.get(ci);
        let records_r_id = records.map(|_| cache.records_r_id());
        let cache_xml =
            cache.write_xml(records_r_id.as_deref(), refresh_on_load.contains(&cache_id))?;
        let cache_path = names.alloc("xl/pivotCache/pivotCacheDefinition", ".xml");
        cache_paths.insert(cache_id, cache_path.clone());

        content_types.add_override(format!("/{cache_path}"), CT_PIVOT_CACHE_DEF);

        let r_id = make_rid(wb_rels.next_r_id());
        wb_rels.add(
            r_id.clone(),
            REL_PIVOT_CACHE_DEF,
            target_from_workbook(&cache_path).to_owned(),
        );
        pivot_cache_refs.push(PivotCacheRef { cache_id, r_id });
        let cache_rels_path = rels_path(&cache_path);
        generated_paths.insert(cache_path.clone());
        entries.push(ZipEntry {
            name: cache_path,
            data: cache_xml,
        });

        // The definition's .rels: the relationships kept from the source
        // under their own ids (its source names them by `r:id`), and the one
        // to its records, in the same folder, when present.
        let mut cache_rels = cache.kept_relationships();
        if let Some((records, records_r_id)) = records.zip(records_r_id) {
            let rec_xml = records.to_xml()?;
            let rec_path =
                names.alloc_preferring("xl/pivotCache/pivotCacheRecords", part_num, ".xml");

            content_types.add_override(format!("/{rec_path}"), CT_PIVOT_CACHE_REC);

            cache_rels.add(
                records_r_id,
                REL_PIVOT_CACHE_REC,
                rec_path
                    .rsplit_once('/')
                    .map_or(rec_path.as_str(), |(_, file)| file)
                    .to_owned(),
            );

            generated_paths.insert(rec_path.clone());
            entries.push(ZipEntry {
                name: rec_path,
                data: rec_xml,
            });
        }
        if !cache_rels.is_empty() {
            entries.push(ZipEntry {
                name: cache_rels_path.clone(),
                data: cache_rels.to_xml()?,
            });
            generated_paths.insert(cache_rels_path);
        }
    }

    // Each pivot table's .rels: rId1 to the definition of the cache its
    // cacheId names (`pivot_cache_ids` checked that one does).
    for (pt_path, cache_id) in &pivot_table_paths {
        let Some(cache_path) = cache_paths.get(cache_id) else {
            continue;
        };
        let file = cache_path
            .rsplit_once('/')
            .map_or(cache_path.as_str(), |(_, file)| file);
        let mut pt_rels = Relationships::new();
        pt_rels.add("rId1", REL_PIVOT_CACHE_DEF, format!("../pivotCache/{file}"));
        let pt_rels_path = rels_path(pt_path);
        entries.push(ZipEntry {
            name: pt_rels_path.clone(),
            data: pt_rels.to_xml()?,
        });
        generated_paths.insert(pt_rels_path);
    }

    // 5c. Write workbook-level persons list (for threaded comments). It
    // supersedes a preserved copy (one the reader could not parse).
    if !workbook.persons.is_empty() {
        debug!("writing {} persons", workbook.persons.len());
        let persons_xml = threaded_comments::write_persons(&workbook.persons)?;
        let persons_path = "xl/persons/person.xml";

        content_types.add_override(format!("/{persons_path}"), CT_PERSONS);

        entries.push(ZipEntry {
            name: persons_path.to_string(),
            data: persons_xml,
        });
        generated_paths.insert(persons_path.to_owned());
        wb_rels.add(
            make_rid(wb_rels.next_r_id()),
            REL_PERSONS,
            "persons/person.xml",
        );
    }

    // 5d. Write workbook-level slicer caches, each with its workbook
    // relationship and Override. The x14 `slicerCaches` entry of the workbook
    // `extLst` is not written.
    for cache in &workbook.slicer_caches {
        let cache_xml = cache.to_xml()?;
        let cache_path = names.alloc("xl/slicerCaches/slicerCache", ".xml");

        content_types.add_override(format!("/{cache_path}"), CT_SLICER_CACHE);
        wb_rels.add(
            make_rid(wb_rels.next_r_id()),
            REL_SLICER_CACHE,
            target_from_workbook(&cache_path).to_owned(),
        );

        generated_paths.insert(cache_path.clone());
        entries.push(ZipEntry {
            name: cache_path,
            data: cache_xml,
        });
    }

    // 5e. Write workbook-level timeline caches, each with its workbook
    // relationship and Override. The x15 `timelineCacheRefs` entry of the
    // workbook `extLst` is not written.
    for cache in &workbook.timeline_caches {
        let cache_xml = cache.to_xml()?;
        let cache_path = names.alloc("xl/timelineCaches/timelineCache", ".xml");

        content_types.add_override(format!("/{cache_path}"), CT_TIMELINE_CACHE);
        wb_rels.add(
            make_rid(wb_rels.next_r_id()),
            REL_TIMELINE_CACHE,
            target_from_workbook(&cache_path).to_owned(),
        );

        generated_paths.insert(cache_path.clone());
        entries.push(ZipEntry {
            name: cache_path,
            data: cache_xml,
        });
    }

    // 5f. Cell metadata (`<c cm>`): xl/metadata.xml needs its workbook
    // relationship and Override. A preserved part keeps its bytes (appended in
    // step 6); otherwise the dynamic-array part is generated when a cell
    // carries cm.
    let metadata_part = generated_metadata_part(workbook)?;
    if metadata_part.is_some() || workbook.preserved_entries.contains_key(METADATA_PATH) {
        if let Some(xml) = metadata_part {
            entries.push(ZipEntry {
                name: METADATA_PATH.to_owned(),
                data: xml.as_bytes().to_vec(),
            });
        }
        content_types.add_override(format!("/{METADATA_PATH}"), CT_SHEET_METADATA);
        wb_rels.add(
            make_rid(wb_rels.next_r_id()),
            REL_SHEET_METADATA,
            "metadata.xml",
        );
    }

    // 5g. External links stay preserved entries (step 6); each gets a
    // workbook relationship and an `<externalReference>`, in
    // `external_link_order`: those `external_references` lists first, then the
    // other preserved external link parts by number.
    let external_references: Vec<String> = external_link_order(workbook)
        .into_iter()
        .map(|path| {
            let r_id = make_rid(wb_rels.next_r_id());
            wb_rels.add(
                r_id.clone(),
                REL_EXTERNAL_LINK,
                target_from_workbook(path).to_owned(),
            );
            r_id
        })
        .collect();

    // Custom XML data parts (`is_custom_xml_item`) also stay preserved entries
    // (step 6); each gets a workbook relationship, in path order.
    for path in workbook
        .preserved_entries
        .keys()
        .filter(|path| is_custom_xml_item(path))
    {
        wb_rels.add(
            make_rid(wb_rels.next_r_id()),
            REL_CUSTOM_XML,
            format!("../{path}"),
        );
    }

    // 5h. workbook.xml and its .rels, now that every relationship it lists exists.
    let wb_xml = WorkbookXml {
        sheets: workbook
            .sheets
            .iter()
            .enumerate()
            .map(|(i, s)| SheetInfo {
                name: s.name.clone(),
                sheet_id: (i + 1) as u32,
                r_id: make_rid(i + 1),
                state: match s.state.as_deref() {
                    Some("hidden") => SheetState::Hidden,
                    Some("veryHidden") => SheetState::VeryHidden,
                    _ => SheetState::Visible,
                },
            })
            .collect(),
        date_system: workbook.date_system,
        defined_names: workbook.defined_names.clone(),
        workbook_views: workbook.workbook_views.clone(),
        protection: workbook.protection.clone(),
        calc_pr: workbook.calc_pr.clone(),
        file_version: workbook.file_version.clone(),
        workbook_pr: workbook.workbook_pr.clone(),
        external_references,
        pivot_caches: pivot_cache_refs,
    };
    entries.push(ZipEntry {
        name: "xl/workbook.xml".to_string(),
        data: wb_xml.to_xml()?,
    });
    entries.push(ZipEntry {
        name: "xl/_rels/workbook.xml.rels".to_string(),
        data: wb_rels.to_xml()?,
    });

    // 6. Append preserved entries (drawings, media, charts, etc.), except those
    // at a path in `generated_paths`, such as a worksheet .rels rebuilt with
    // its merged relationships or a drawing merged with the sheet's charts.
    // Every part written so far, and [Content_Types].xml (pushed in step 7),
    // joins `generated_paths` first: a preserved entry the source package left
    // unreferenced can sit at a name the writer generates (a worksheet part no
    // `<sheet>` names, at the name of a sheet the writer numbers), and a ZIP
    // holds each name once, so the generated part wins and the entry is
    // dropped (here, and from the content types of steps 6b and 6c). OPC part
    // names are equal ignoring ASCII case, so an entry gives way to a
    // generated name in any letter case: one lookup of its lower-cased name
    // per entry. A relationship that names the dropped spelling then resolves
    // to the generated part.
    for entry in &entries {
        if !generated_paths.contains(&entry.name) {
            generated_paths.insert(entry.name.clone());
        }
    }
    generated_paths.insert(CONTENT_TYPES_PATH.to_owned());
    let generated_lower: HashSet<String> = generated_paths
        .iter()
        .map(|path| path.to_ascii_lowercase())
        .collect();
    // The entries dropped for a generated name in another letter case.
    let mut case_dropped: HashSet<&str> = HashSet::new();
    for (path, data) in &workbook.preserved_entries {
        if generated_paths.contains(path) {
            continue;
        }
        if generated_lower.contains(&path.to_ascii_lowercase()) {
            debug!("preserved entry {path} gives way to a generated part of that name");
            case_dropped.insert(path.as_str());
            continue;
        }
        trace!("writing preserved ZIP entry: {path}");
        entries.push(ZipEntry {
            name: path.clone(),
            data: data.clone(),
        });
    }

    // A chart part left preserved keeps its .rels preserved beside it; the
    // parts it names are typed like those of the written charts.
    for (path, data) in &workbook.preserved_entries {
        if let Some(file) = path
            .strip_prefix("xl/charts/_rels/")
            .and_then(|file| file.strip_suffix(".rels"))
            && !generated_paths.contains(path)
            && !case_dropped.contains(path.as_str())
        {
            collect_chart_part_types(&format!("xl/charts/{file}"), data, &mut chart_part_types);
        }
    }

    // 6b. Auto-detect content types from preserved entries.
    // Media, VML and embedded OOXML packages need extension Defaults,
    // registered under the lowercased extension (OPC matches extensions
    // case-insensitively, so `PNG` and `png` would be one duplicated Default);
    // other parts need Overrides.
    // A part a chart .rels names takes the type that relationship implies. A
    // userShapes part is usually named like a sheet drawing
    // (`xl/drawings/drawing{n}.xml`) but takes the chartshapes type; it is always a
    // preserved entry, which `PartNames` never hands to a generated sheet drawing.
    for path in workbook.preserved_entries.keys() {
        if case_dropped.contains(path.as_str()) {
            continue;
        }
        let file_name = path
            .rsplit_once('/')
            .map_or(path.as_str(), |(_, name)| name);
        if let Some((_, ext)) = file_name.rsplit_once('.') {
            let ext = ext.to_ascii_lowercase();
            let content_type = match ext.as_str() {
                "png" => Some("image/png"),
                "jpg" | "jpeg" => Some("image/jpeg"),
                "gif" => Some("image/gif"),
                "emf" => Some("image/x-emf"),
                "wmf" => Some("image/x-wmf"),
                "svg" => Some("image/svg+xml"),
                "bmp" => Some("image/bmp"),
                "tif" | "tiff" => Some("image/tiff"),
                "wdp" => Some(CT_HD_PHOTO),
                "vml" => Some(CT_VML_DRAWING),
                "xlsx" => Some(CT_XLSX_PACKAGE),
                "docx" => Some(CT_DOCX_PACKAGE),
                "pptx" => Some(CT_PPTX_PACKAGE),
                _ => None,
            };
            if let Some(content_type) = content_type {
                content_types.add_default(ext, content_type);
            }
        }
        if path.starts_with("xl/printerSettings/") && path.ends_with(".bin") {
            content_types.add_override(format!("/{path}"), CT_PRINTER_SETTINGS);
        }
        if path.starts_with("xl/embeddings/") && path.ends_with(".bin") {
            content_types.add_override(format!("/{path}"), CT_OLE_OBJECT);
        }
        if let Some(&content_type) = chart_part_types.get(path)
            && !generated_paths.contains(path)
        {
            content_types.add_override(format!("/{path}"), content_type);
        } else if path.starts_with("xl/drawings/drawing") && path.ends_with(".xml") {
            content_types.add_override(format!("/{path}"), CT_DRAWING);
        }
        if path.starts_with("xl/charts/chartEx") && path.ends_with(".xml") {
            content_types.add_override(format!("/{path}"), CT_CHART_EX);
        } else if path.starts_with("xl/charts/chart") && path.ends_with(".xml") {
            content_types.add_override(format!("/{path}"), CT_CHART);
        }
        // External links
        if is_external_link(path) {
            content_types.add_override(format!("/{path}"), CT_EXTERNAL_LINK);
        }
        // Custom XML items (plain XML, exclude itemProps files)
        if path.starts_with("customXml/item") && path.ends_with(".xml") && !path.contains("Props") {
            content_types.add_override(format!("/{path}"), CT_XML);
        }
        // Custom XML properties
        if path.starts_with("customXml/itemProps") && path.ends_with(".xml") {
            content_types.add_override(format!("/{path}"), CT_CUSTOM_XML_PROPS);
        }
    }

    // 6c. The content type the source package declared for a preserved part
    // (`preserved_content_types`) wins over the rules above, as an Override,
    // unless the Default for its extension already gives that type; an
    // Override set above is then kept. A part the writer generated at that
    // path, and a `.rels`, keep their own type.
    for (path, content_type) in &workbook.preserved_content_types {
        if !workbook.preserved_entries.contains_key(path)
            || generated_paths.contains(path)
            || case_dropped.contains(path.as_str())
            || path.ends_with(".rels")
        {
            continue;
        }
        let file_name = path
            .rsplit_once('/')
            .map_or(path.as_str(), |(_, name)| name);
        let by_default = file_name.rsplit_once('.').is_some_and(|(_, ext)| {
            content_types.defaults.get(&ext.to_ascii_lowercase()) == Some(content_type)
        });
        if !by_default {
            content_types.add_override(format!("/{path}"), content_type.clone());
        }
    }

    // 7. Write [Content_Types].xml, now that every Default and Override is
    // known, and rotate it to be the first entry of the ZIP archive.
    entries.push(ZipEntry {
        name: CONTENT_TYPES_PATH.to_owned(),
        data: content_types.to_xml()?,
    });
    entries.rotate_right(1);

    for entry in &entries {
        trace!("writing ZIP entry: {}", entry.name);
    }

    write_zip(&entries)
}

/// Write a password-protected XLSX file using Agile Encryption (AES-256-CBC, SHA-512).
///
/// First writes the workbook as a normal ZIP archive, then encrypts it into
/// an OLE2 compound document containing `EncryptionInfo` and `EncryptedPackage`
/// streams.
///
/// The encryption is [`crate::ole2::crypto::encrypt_file`]: Agile Encryption
/// with SHA-512, a 256-bit AES key and 100,000 spin-count rounds. The output
/// reads back through [`crate::reader::read_xlsx_json_with_password`], and
/// msoffcrypto-tool decrypts `encrypt_file` output in the ignored test of
/// `tests/interop/main.rs`, which the CI job `encryption-oracle` runs.
#[cfg(feature = "encryption")]
pub fn write_xlsx_with_password(workbook: &WorkbookData, password: &str) -> Result<Vec<u8>> {
    let zip_bytes = write_xlsx(workbook)?;
    crate::ole2::crypto::encrypt_file(&zip_bytes, password)
}

// ---------------------------------------------------------------------------
// Part naming and drawing merge helpers
// ---------------------------------------------------------------------------

/// The content type of the part a chart relationship of `rel_type` targets,
/// for the types whose part the writer must declare by Override.
fn chart_part_content_type(rel_type: &str) -> Option<&'static str> {
    match rel_type {
        REL_CHART_USER_SHAPES => Some(CT_CHART_SHAPES),
        REL_CHART_STYLE => Some(CT_CHART_STYLE),
        REL_CHART_COLOR_STYLE => Some(CT_CHART_COLOR_STYLE),
        _ => None,
    }
}

/// Adds to `out` the ZIP path and content type of every part with a
/// [`chart_part_content_type`] (user shapes, style, colour style) that `rels`,
/// the .rels of the chart at `chart_path`, names. A malformed .rels names none.
fn collect_chart_part_types(
    chart_path: &str,
    rels: &[u8],
    out: &mut HashMap<String, &'static str>,
) {
    match Relationships::parse(rels) {
        Ok(rels) => out.extend(rels.iter().filter_map(|rel| {
            chart_part_content_type(&rel.rel_type)
                .map(|content_type| (resolve_target(chart_path, &rel.target), content_type))
        })),
        Err(e) => {
            cold_path();
            warn!("{chart_path} .rels: {e}; the parts it names are not typed");
        }
    }
}

/// The relationship target of the `xl/` part at `path` from a part in another
/// folder directly under `xl/` (a worksheet, a drawing): `../charts/chart1.xml`.
fn target_from_xl_folder(path: &str) -> String {
    format!("../{}", target_from_workbook(path))
}

/// The relationship target of the `xl/` part at `path` from `xl/workbook.xml`:
/// the path without `xl/`.
fn target_from_workbook(path: &str) -> &str {
    path.strip_prefix("xl/").unwrap_or(path)
}

/// The part names one write has taken, and for each `{prefix}{n}{ext}` family
/// the number where the next free-name search resumes.
///
/// `used` holds ASCII-lowercased names, because OPC part names are equal
/// ignoring ASCII case while the ZIP writer rejects only byte-identical
/// duplicates. It starts as the preserved entries and only grows: no name is
/// released during a write.
///
/// Invariant: for every `(prefix, ext)` in `next`, each `n` in
/// `1..next[(prefix, ext)]` makes a name in `used`. [`Self::alloc`] moves a
/// cursor only past numbers it found taken or took itself, and as nothing
/// leaves `used`, the numbers below a cursor stay taken whatever
/// [`Self::alloc_preferring`] or [`Self::take`] adds later, above the cursor
/// or not. So the first free number at or above the cursor is the first free
/// number from 1: `alloc` returns the name a search from 1 would, and over a
/// whole write each family costs one probe per allocated or taken name rather
/// than one per earlier name on every call.
struct PartNames {
    used: HashSet<String>,
    next: HashMap<(&'static str, &'static str), usize>,
}

impl PartNames {
    /// Starts from the names already `taken` (the preserved entries).
    fn new<S: AsRef<str>>(taken: impl IntoIterator<Item = S>) -> Self {
        Self {
            used: taken
                .into_iter()
                .map(|name| name.as_ref().to_ascii_lowercase())
                .collect(),
            next: HashMap::new(),
        }
    }

    /// Marks `name` taken (a part kept at a path of its own).
    fn take(&mut self, name: &str) {
        self.used.insert(name.to_ascii_lowercase());
    }

    /// Claims the part name `{prefix}{n}{ext}` when it is free.
    fn claim(&mut self, prefix: &str, n: usize, ext: &str) -> Option<String> {
        let mut name = String::with_capacity(prefix.len() + 20 + ext.len());
        name.push_str(prefix);
        name.push_str(itoa::Buffer::new().format(n));
        name.push_str(ext);
        self.used.insert(name.to_ascii_lowercase()).then_some(name)
    }

    /// Allocates the first part name `{prefix}{n}{ext}` (`n` from 1) not
    /// taken and marks it taken, so a generated part never collides with a
    /// preserved entry or another generated part. The search starts at the
    /// family's cursor (see the invariant on [`PartNames`]).
    fn alloc(&mut self, prefix: &'static str, ext: &'static str) -> String {
        let mut n = self.next.get(&(prefix, ext)).copied().unwrap_or(1);
        loop {
            if let Some(name) = self.claim(prefix, n, ext) {
                self.next.insert((prefix, ext), n + 1);
                return name;
            }
            n += 1;
        }
    }

    /// Like [`Self::alloc`], but takes `{prefix}{preferred}{ext}` when it is
    /// free, so a part numbered after its sheet or cache (`comments2.xml` for
    /// sheet 2) keeps that name whenever nothing collides with it.
    fn alloc_preferring(
        &mut self,
        prefix: &'static str,
        preferred: usize,
        ext: &'static str,
    ) -> String {
        self.claim(prefix, preferred, ext)
            .unwrap_or_else(|| self.alloc(prefix, ext))
    }
}

/// The direct `<xdr:wsDr>` children a chart merge keeps from a preserved
/// drawing, and the namespace declarations on its root.
struct KeptDrawing<'a> {
    /// Verbatim XML of each kept child, in document order.
    children: Vec<&'a str>,
    /// `(prefix, uri)` of each declaration on `<xdr:wsDr>`.
    namespaces: Vec<(String, String)>,
}

/// Whether a direct `<xdr:wsDr>` child is kept by a merge: `Some(true)` for an
/// anchor (dropped later if it frames a chart), `Some(false)` for an
/// `mc:AlternateContent` block (always kept), `None` for anything else.
fn kept_child_kind(e: &BytesStart<'_>) -> Option<bool> {
    match e.local_name().as_ref() {
        "twoCellAnchor" | "oneCellAnchor" | "absoluteAnchor" => Some(true),
        "AlternateContent" => Some(false),
        _ => None,
    }
}

/// Scans a preserved drawing part in one forward pass over the direct
/// children of its root `<xdr:wsDr>`.
///
/// The part may come from any producer (`addImage`, Excel, other libraries), so
/// it is read as XML, not by text matching: comments, CDATA, processing
/// instructions and any namespace prefix, or none (openpyxl declares
/// SpreadsheetDrawing as the default namespace), are handled. The anchors
/// (`xdr:twoCellAnchor`, `xdr:oneCellAnchor`, `xdr:absoluteAnchor`) and
/// `mc:AlternateContent` blocks are kept verbatim in document order, which is
/// their z-order, except an anchor whose `xdr:graphicFrame` holds a `c:chart`:
/// the sheet's charts are regenerated from the model. Elements match by local
/// name. Text and comments between the children are dropped. The root's
/// namespace declarations, the default one included, are returned so the
/// merge can declare them on the generated root.
///
/// # Errors
///
/// `XmlParse` naming `path` when the part is not well-formed, including an
/// element still open at the end of the part (the quick-xml `Reader` reports
/// mismatched end tags but not unclosed elements at EOF, so that is checked
/// here).
fn kept_drawing_children<'a>(xml: &'a str, path: &str) -> Result<KeptDrawing<'a>> {
    let parse_err = |e: &dyn std::fmt::Display| {
        cold_path();
        ModernXlsxError::XmlParse(format!("{path}: {e}"))
    };
    let mut reader = Reader::from_str(xml);
    let offset = |reader: &Reader<&[u8]>| {
        usize::try_from(reader.buffer_position()).map_err(|e| parse_err(&e))
    };
    let mut kept = KeptDrawing {
        children: Vec::new(),
        namespaces: Vec::new(),
    };
    // Open elements; the root's children start at depth 1.
    let mut depth = 0usize;
    // Start offset of the kept child being scanned and whether it is an anchor.
    let mut child: Option<(usize, bool)> = None;
    // Depth of the open `xdr:graphicFrame`, and whether it holds a chart.
    let mut frame_depth: Option<usize> = None;
    let mut frames_chart = false;
    loop {
        let start = offset(&reader)?;
        match reader.read_event().map_err(|e| parse_err(&e))? {
            Event::Start(e) => {
                match depth {
                    0 => kept.namespaces = root_namespaces(&e)?,
                    1 => {
                        child = kept_child_kind(&e).map(|anchor| (start, anchor));
                        frames_chart = false;
                    }
                    _ if child.is_some() => match e.local_name().as_ref() {
                        "graphicFrame" => frame_depth = frame_depth.or(Some(depth)),
                        "chart" if frame_depth.is_some() => frames_chart = true,
                        _ => {}
                    },
                    _ => {}
                }
                depth += 1;
            }
            Event::Empty(e) => match depth {
                0 => kept.namespaces = root_namespaces(&e)?,
                1 => {
                    if kept_child_kind(&e).is_some() {
                        kept.children.push(&xml[start..offset(&reader)?]);
                    }
                }
                _ => {
                    if frame_depth.is_some() && e.local_name().as_ref() == "chart" {
                        frames_chart = true;
                    }
                }
            },
            Event::End(_) => {
                // The reader checks end names, so every End closes an open element.
                depth = depth.saturating_sub(1);
                if frame_depth == Some(depth) {
                    frame_depth = None;
                }
                if depth == 1
                    && let Some((from, anchor)) = child.take()
                    && !(anchor && frames_chart)
                {
                    kept.children.push(&xml[from..offset(&reader)?]);
                }
            }
            Event::Eof if depth != 0 => {
                cold_path();
                return Err(ModernXlsxError::XmlParse(format!(
                    "{path}: unclosed element at end of part"
                )));
            }
            Event::Eof => break,
            _ => {}
        }
    }
    Ok(kept)
}

/// For `tail` starting with an `r:<NCName>` attribute name followed by `=` and
/// a quoted value (XML allows whitespace around the `=`), the byte offset of
/// the value in `tail` and the value.
fn r_attr_value(tail: &str) -> Option<(usize, &str)> {
    let name = tail.strip_prefix("r:")?;
    if !name.starts_with(|c: char| c.is_alphabetic() || c == '_') {
        return None;
    }
    let name_end = name
        .find(|c: char| !(c.is_alphanumeric() || matches!(c, '_' | '-' | '.')))
        .unwrap_or(name.len());
    let quoted = name[name_end..]
        .trim_start_matches(crate::ooxml::xml::XML_S)
        .strip_prefix('=')?
        .trim_start_matches(crate::ooxml::xml::XML_S);
    let quote = quoted.chars().next().filter(|q| matches!(q, '"' | '\''))?;
    let inner = &quoted[1..];
    let end = inner.find(quote)?;
    Some((tail.len() - inner.len(), &inner[..end]))
}

/// Appends `xml` to `out` with the value of every `r:<NCName>` attribute
/// (`r:embed`, `r:link`, `r:id`, `r:pict`, the SmartArt `r:dm`, `r:lo`, `r:qs`
/// and `r:cs`, or any other relationship-namespace attribute) that is a key in
/// `remap` replaced, in one forward pass: every value is looked up once, so
/// `rId1 -> rId2` never chains into a later `rId2 -> rId3`. An attribute name
/// must follow whitespace, so `xr:id` and element names such as `<r:x>` are
/// left alone.
fn push_remapped_rel_ids(out: &mut String, xml: &str, remap: &HashMap<&str, &str>) {
    let mut rest = xml;
    while let Some(at) = rest.find("r:") {
        let (head, tail) = rest.split_at(at);
        out.push_str(head);
        // An attribute name follows whitespace inside a start tag.
        let value = if out.ends_with(crate::ooxml::xml::XML_S) {
            r_attr_value(tail)
        } else {
            None
        };
        if let Some((value_at, value)) = value {
            out.push_str(&tail[..value_at]);
            out.push_str(remap.get(value).copied().unwrap_or(value));
            rest = &tail[value_at + value.len()..];
        } else {
            out.push_str("r:");
            rest = &tail[2..];
        }
    }
    out.push_str(rest);
}

/// Merges the preserved drawing at `path` (with its `.rels`, if any) into
/// `generated`, the drawing just written for the sheet's charts, which already
/// lives at `path`.
///
/// The kept children (see [`kept_drawing_children`]) are appended after the
/// chart anchors, so they keep their relative order (z-order) and the
/// regenerated charts are drawn behind them. The preserved relationships,
/// chart ones excepted, move to `drawing_rels` under fresh rIds following the
/// chart rIds, and every `r:*` attribute in the kept children whose value is
/// one of the old rIds (`r:embed`, `r:link`, `r:id`, `r:pict`, SmartArt
/// `r:dm`/`r:lo`/`r:qs`/`r:cs`, ...) is rewritten to match (see
/// [`push_remapped_rel_ids`]). Namespace declarations on the preserved root
/// that the generated one lacks are declared on it: each missing prefix, and
/// the default namespace (openpyxl's unprefixed anchors rely on it; the
/// generated drawing uses only prefixed names, so a default declaration never
/// changes it).
///
/// A preserved drawing the merge cannot read (not UTF-8, not well-formed, or a
/// malformed preserved `.rels`) is a warning, not an error: `generated`
/// is returned as it is.
///
/// # Errors
///
/// `XmlParse` when `generated` is not UTF-8; `XmlWrite` when it has no
/// `<xdr:wsDr>` root.
fn merge_preserved_drawing(
    generated: Vec<u8>,
    path: &str,
    preserved: &[u8],
    preserved_rels: Option<&[u8]>,
    drawing_rels: &mut Relationships,
) -> Result<Vec<u8>> {
    let not_merged = |e: &dyn std::fmt::Display| {
        cold_path();
        warn!("{path}: {e}; preserved drawing not merged");
    };
    let Ok(preserved) = std::str::from_utf8(preserved) else {
        not_merged(&"not valid UTF-8");
        return Ok(generated);
    };
    let kept = match kept_drawing_children(preserved, path) {
        Ok(kept) => kept,
        Err(e) => {
            not_merged(&e);
            return Ok(generated);
        }
    };
    if kept.children.is_empty() {
        return Ok(generated);
    }
    debug!("{path}: merging {} preserved anchors", kept.children.len());

    let generated = String::from_utf8(generated).map_err(|_| {
        cold_path();
        ModernXlsxError::XmlParse(format!("{path}: generated drawing is not valid UTF-8"))
    })?;
    let root_start = generated.find("<xdr:wsDr");
    let root_end = root_start.and_then(|s| generated[s..].find('>').map(|e| s + e));
    let close = generated.rfind("</xdr:wsDr>");
    let (Some(root_start), Some(root_end), Some(close)) = (root_start, root_end, close) else {
        cold_path();
        return Err(ModernXlsxError::XmlWrite(format!(
            "{path}: generated drawing has no <xdr:wsDr> root"
        )));
    };

    let preserved_rels = match preserved_rels.map(Relationships::parse).transpose() {
        Ok(rels) => rels.unwrap_or_default(),
        Err(e) => {
            not_merged(&e);
            return Ok(generated.into_bytes());
        }
    };
    // Chart relationships belong to the chart frames the merge drops; an image
    // copied whole keeps an `r:link` TargetMode="External".
    let carried: Vec<&Relationship> = preserved_rels
        .iter()
        .filter(|rel| rel.rel_type != REL_CHART)
        .collect();
    let new_ids: Vec<String> = carried
        .iter()
        .map(|rel| {
            let id = make_rid(drawing_rels.next_r_id());
            drawing_rels.push(Relationship {
                id: id.clone(),
                ..(*rel).clone()
            });
            id
        })
        .collect();
    let remap: HashMap<&str, &str> = carried
        .iter()
        .map(|rel| rel.id.as_str())
        .zip(new_ids.iter().map(String::as_str))
        .collect();

    let root = &generated[root_start..root_end];
    let mut merged = String::with_capacity(generated.len() + preserved.len());
    merged.push_str(&generated[..root_end]);
    for (prefix, uri) in &kept.namespaces {
        let attr = if prefix.is_empty() {
            " xmlns=".to_owned()
        } else {
            format!(" xmlns:{prefix}=")
        };
        if !root.contains(&attr) {
            merged.push_str(&attr);
            merged.push('"');
            merged.push_str(&quick_xml::escape::escape(uri.as_str()));
            merged.push('"');
        }
    }
    merged.push_str(&generated[root_end..close]);
    for child in &kept.children {
        push_remapped_rel_ids(&mut merged, child, &remap);
    }
    merged.push_str(&generated[close..]);
    Ok(merged.into_bytes())
}

// ---------------------------------------------------------------------------
// Tests
// ---------------------------------------------------------------------------

#[cfg(test)]
mod tests;
