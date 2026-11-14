//! Package-level tests: an Excel-authored fixture read into a golden, and the
//! part graph (parts, relationships, content types) of written packages.
//!
//! The fixture is `tests/fixtures/msoffcrypto/example.xlsx` (origin and licence
//! in its `README.md`).

use std::collections::HashMap;

use modern_xlsx_core::ooxml::content_types::ContentTypes;
use modern_xlsx_core::ooxml::relationships::{Relationships, TargetMode};
use modern_xlsx_core::reader::read_xlsx;
use modern_xlsx_core::writer::write_xlsx;
use modern_xlsx_core::zip::reader::{ZipSecurityLimits, read_zip_entries};
use modern_xlsx_core::zip::writer::{ZipEntry, write_zip};
use modern_xlsx_core::{SheetData, WorkbookData};
use serde_json::json;

mod common;

use common::{golden_subtree, minimal_workbook, number_row, text_row};

const EXCEL_EXAMPLE: &[u8] = include_bytes!("fixtures/msoffcrypto/example.xlsx");

const REL_NS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const REL_THEME: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme";
const CT_THEME: &str = "application/vnd.openxmlformats-officedocument.theme+xml";

fn zip_map(bytes: &[u8]) -> HashMap<String, Vec<u8>> {
    read_zip_entries(bytes, &ZipSecurityLimits::default())
        .expect("a readable ZIP")
        .into_iter()
        .collect()
}

/// The part a `.rels` entry belongs to: `xl/_rels/workbook.xml.rels` gives
/// `xl/workbook.xml`, and `_rels/.rels` the package root (`""`).
fn rels_source(rels_path: &str) -> String {
    let (dir, file) = rels_path.rsplit_once("_rels/").expect("a .rels path");
    format!("{dir}{}", file.strip_suffix(".rels").expect("a .rels path"))
}

/// The ZIP entry an internal relationship target of `source` names.
fn resolve(source: &str, target: &str) -> String {
    if let Some(absolute) = target.strip_prefix('/') {
        return absolute.to_owned();
    }
    let mut parts: Vec<&str> = source.split('/').collect();
    parts.pop();
    for segment in target.split('/') {
        match segment {
            ".." => {
                parts.pop();
            }
            "" | "." => {}
            other => parts.push(other),
        }
    }
    parts.join("/")
}

/// Asserts every internal relationship target in every `.rels` entry of the
/// package is a ZIP entry.
fn assert_every_target_resolves(entries: &HashMap<String, Vec<u8>>) {
    let mut checked = 0;
    for (path, data) in entries.iter().filter(|(p, _)| p.ends_with(".rels")) {
        let source = rels_source(path);
        for rel in &Relationships::parse(data).expect("parsable .rels") {
            if rel.target_mode == TargetMode::External {
                continue;
            }
            let target = resolve(&source, &rel.target);
            assert!(
                entries.contains_key(&target),
                "{path}: {} ({}) targets missing {target}",
                rel.id,
                rel.rel_type
            );
            checked += 1;
        }
    }
    assert!(checked > 0, "no relationship checked");
}

/// The relationships of type `rel_type` in the `.rels` entry `rels_path`, as
/// (Id, Target) pairs.
fn rels_of_type(
    entries: &HashMap<String, Vec<u8>>,
    rels_path: &str,
    rel_type: &str,
) -> Vec<(String, String)> {
    let data = entries
        .get(rels_path)
        .unwrap_or_else(|| panic!("missing {rels_path}"));
    Relationships::parse(data)
        .unwrap()
        .find_by_type(rel_type)
        .map(|rel| (rel.id.clone(), rel.target.to_string()))
        .collect()
}

fn content_types(entries: &HashMap<String, Vec<u8>>) -> ContentTypes {
    ContentTypes::parse(&entries["[Content_Types].xml"]).unwrap()
}

// ---------------------------------------------------------------------------
// Excel-authored fixture
// ---------------------------------------------------------------------------

#[test]
fn excel_example_reads_to_its_golden() {
    let wb = read_xlsx(EXCEL_EXAMPLE).expect("the Excel fixture reads");
    let styles = &wb.styles;
    let projection = json!({
        "sheets": wb.sheets.iter().map(|s| json!({
            "name": s.name,
            "state": s.state,
            "rows": s.worksheet.rows.iter().take(5).collect::<Vec<_>>(),
        })).collect::<Vec<_>>(),
        "definedNames": wb.defined_names,
        "themeColors": wb.theme_colors,
        "styles": {
            "numFmts": styles.num_fmts.len(),
            "fonts": styles.fonts.len(),
            "fills": styles.fills.len(),
            "borders": styles.borders.len(),
            "cellStyleXfs": styles.cell_style_xfs.len(),
            "cellXfs": styles.cell_xfs.len(),
            "dxfs": styles.dxfs.len(),
            "cellStyles": styles.cell_styles.len(),
        },
    });
    golden_subtree(&projection, "excel_example");
}

#[test]
fn excel_example_round_trip_keeps_theme_and_resolves_every_target() {
    let input = zip_map(EXCEL_EXAMPLE);
    assert_eq!(
        rels_of_type(&input, "xl/_rels/workbook.xml.rels", REL_THEME).len(),
        1,
        "the fixture has a theme rel"
    );

    let wb = read_xlsx(EXCEL_EXAMPLE).unwrap();
    let out = write_xlsx(&wb).expect("the Excel fixture writes");
    let entries = zip_map(&out);

    let themes: Vec<String> = rels_of_type(&entries, "xl/_rels/workbook.xml.rels", REL_THEME)
        .into_iter()
        .map(|(_, target)| target)
        .collect();
    assert_eq!(themes, ["theme/theme1.xml"]);
    assert_eq!(
        content_types(&entries)
            .overrides
            .get("/xl/theme/theme1.xml")
            .map(String::as_str),
        Some(CT_THEME)
    );
    // The preserved theme keeps Excel's bytes.
    assert_eq!(entries["xl/theme/theme1.xml"], input["xl/theme/theme1.xml"]);
    assert_every_target_resolves(&entries);

    let wb2 = read_xlsx(&out).unwrap();
    assert_eq!(
        serde_json::to_value(&wb2.theme_colors).unwrap(),
        serde_json::to_value(&wb.theme_colors).unwrap()
    );
}

// ---------------------------------------------------------------------------
// Parts, rels and content types of written packages
// ---------------------------------------------------------------------------

/// One sheet with a pivot table (cache + records), threaded comments with
/// their persons, a slicer and a timeline with their caches, document
/// properties and a calculation chain over the formula cell A2.
fn full_graph_workbook() -> WorkbookData {
    let mut wb = minimal_workbook("Sheet1", Default::default());
    let ws = &mut wb.sheets[0].worksheet;
    let mut total = number_row(2, "A2", "1");
    total.cells[0].formula = Some("1+0".into());
    ws.rows = vec![text_row(1, "A1", "Region"), total];
    ws.pivot_tables = serde_json::from_value(json!([
        { "name": "PT1", "location": { "ref": "D1:E3" }, "cacheId": 0 },
    ]))
    .unwrap();
    ws.threaded_comments = serde_json::from_value(json!([{
        "id": "{00000000-0000-0000-0000-000000000001}",
        "refCell": "A1",
        "personId": "{00000000-0000-0000-0000-0000000000AA}",
        "text": "Check this",
        "timestamp": "2026-11-11T00:00:00.00",
    }]))
    .unwrap();
    ws.slicers =
        serde_json::from_value(json!([{ "name": "Slicer_Region", "cacheName": "Slicer_Region" }]))
            .unwrap();
    ws.timelines =
        serde_json::from_value(json!([{ "name": "Timeline_A", "cacheName": "NativeTimeline_A" }]))
            .unwrap();
    wb.pivot_caches = serde_json::from_value(json!([
        { "source": { "ref": "A1:B5", "sheet": "Sheet1" }, "fields": [{ "name": "Region" }] },
    ]))
    .unwrap();
    wb.pivot_cache_records =
        serde_json::from_value(json!([{ "records": [[{ "type": "string", "v": "East" }]] }]))
            .unwrap();
    wb.persons = serde_json::from_value(json!([
        { "id": "{00000000-0000-0000-0000-0000000000AA}", "displayName": "Ann" },
    ]))
    .unwrap();
    wb.slicer_caches =
        serde_json::from_value(json!([{ "name": "Slicer_Region", "sourceName": "Region" }]))
            .unwrap();
    wb.timeline_caches =
        serde_json::from_value(json!([{ "name": "NativeTimeline_A", "sourceName": "Date" }]))
            .unwrap();
    wb.doc_properties = serde_json::from_value(json!({
        "title": "Graph",
        "creator": "modern-xlsx",
        "application": "modern-xlsx",
    }))
    .unwrap();
    wb.calc_chain = serde_json::from_value(json!([{ "cellRef": "A2", "sheetId": 1 }])).unwrap();
    wb
}

#[test]
fn every_written_part_has_its_rels_and_content_type() {
    let entries = zip_map(&write_xlsx(&full_graph_workbook()).expect("the workbook writes"));
    let ct = content_types(&entries);

    // Each part, with the Override that types it.
    for (part, content_type) in [
        (
            "xl/pivotCache/pivotCacheDefinition1.xml",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml",
        ),
        (
            "xl/pivotCache/pivotCacheRecords1.xml",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml",
        ),
        (
            "xl/pivotTables/pivotTable1.xml",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml",
        ),
        (
            "xl/persons/person.xml",
            "application/vnd.ms-excel.person+xml",
        ),
        (
            "xl/threadedComments/threadedComment1.xml",
            "application/vnd.ms-excel.threadedcomments+xml",
        ),
        (
            "xl/slicers/slicer1.xml",
            "application/vnd.ms-excel.slicer+xml",
        ),
        (
            "xl/slicerCaches/slicerCache1.xml",
            "application/vnd.ms-excel.slicerCache+xml",
        ),
        (
            "xl/timelines/timeline1.xml",
            "application/vnd.ms-excel.timeline+xml",
        ),
        (
            "xl/timelineCaches/timelineCache1.xml",
            "application/vnd.ms-excel.timelineCache+xml",
        ),
        (
            "docProps/core.xml",
            "application/vnd.openxmlformats-package.core-properties+xml",
        ),
        (
            "docProps/app.xml",
            "application/vnd.openxmlformats-officedocument.extended-properties+xml",
        ),
        (
            "xl/calcChain.xml",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml",
        ),
        ("xl/theme/theme1.xml", CT_THEME),
    ] {
        assert!(entries.contains_key(part), "missing part {part}");
        assert_eq!(
            ct.overrides.get(&format!("/{part}")).map(String::as_str),
            Some(content_type),
            "Override of {part}"
        );
    }

    // Each part is reached through a relationship of its type.
    let pkg = "http://schemas.openxmlformats.org/package/2006/relationships";
    let ms17 = "http://schemas.microsoft.com/office/2017/10/relationships";
    let ms07 = "http://schemas.microsoft.com/office/2007/relationships";
    let ms11 = "http://schemas.microsoft.com/office/2011/relationships";
    for (rels_path, rel_type, target) in [
        (
            "_rels/.rels",
            format!("{pkg}/metadata/core-properties"),
            "docProps/core.xml",
        ),
        (
            "_rels/.rels",
            format!("{REL_NS}/extended-properties"),
            "docProps/app.xml",
        ),
        (
            "xl/_rels/workbook.xml.rels",
            format!("{REL_NS}/pivotCacheDefinition"),
            "pivotCache/pivotCacheDefinition1.xml",
        ),
        (
            "xl/_rels/workbook.xml.rels",
            format!("{ms17}/person"),
            "persons/person.xml",
        ),
        (
            "xl/_rels/workbook.xml.rels",
            format!("{ms07}/slicerCache"),
            "slicerCaches/slicerCache1.xml",
        ),
        (
            "xl/_rels/workbook.xml.rels",
            format!("{ms11}/timelineCache"),
            "timelineCaches/timelineCache1.xml",
        ),
        (
            "xl/_rels/workbook.xml.rels",
            format!("{REL_NS}/calcChain"),
            "calcChain.xml",
        ),
        (
            "xl/_rels/workbook.xml.rels",
            REL_THEME.to_owned(),
            "theme/theme1.xml",
        ),
        (
            "xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels",
            format!("{REL_NS}/pivotCacheRecords"),
            "pivotCacheRecords1.xml",
        ),
        (
            "xl/worksheets/_rels/sheet1.xml.rels",
            format!("{REL_NS}/pivotTable"),
            "../pivotTables/pivotTable1.xml",
        ),
        (
            "xl/worksheets/_rels/sheet1.xml.rels",
            format!("{ms17}/threadedComment"),
            "../threadedComments/threadedComment1.xml",
        ),
        (
            "xl/worksheets/_rels/sheet1.xml.rels",
            format!("{ms07}/slicer"),
            "../slicers/slicer1.xml",
        ),
        (
            "xl/worksheets/_rels/sheet1.xml.rels",
            format!("{ms11}/timeline"),
            "../timelines/timeline1.xml",
        ),
    ] {
        let targets: Vec<String> = rels_of_type(&entries, rels_path, &rel_type)
            .into_iter()
            .map(|(_, t)| t)
            .collect();
        assert_eq!(targets, [target], "{rels_path}: {rel_type}");
    }

    // The pivot cache is listed in workbook.xml under the table's cacheId.
    let wb_xml = String::from_utf8(entries["xl/workbook.xml"].clone()).unwrap();
    let (cache_r_id, _) = rels_of_type(
        &entries,
        "xl/_rels/workbook.xml.rels",
        &format!("{REL_NS}/pivotCacheDefinition"),
    )
    .remove(0);
    assert!(
        wb_xml.contains(&format!(r#"<pivotCache cacheId="0" r:id="{cache_r_id}"/>"#)),
        "{wb_xml}"
    );

    assert_every_target_resolves(&entries);
}

#[test]
fn shared_string_cell_without_value_is_an_invalid_cell_value() {
    let mut wb = minimal_workbook("Sheet1", Default::default());
    let mut row = text_row(1, "A1", "x");
    row.cells[0].value = None;
    wb.sheets[0].worksheet.rows = vec![row];
    let err = write_xlsx(&wb).unwrap_err();
    assert_eq!(err.code(), "INVALID_CELL_VALUE", "{err}");
}

#[test]
fn two_sheet_package_resolves_every_target() {
    let mut wb = full_graph_workbook();
    wb.sheets.push(SheetData {
        name: "Second".into(),
        worksheet: common::minimal_worksheet(vec![text_row(1, "B2", "y")]),
        ..Default::default()
    });
    let out = write_xlsx(&wb).unwrap();
    assert_every_target_resolves(&zip_map(&out));
    let wb2 = read_xlsx(&out).unwrap();
    assert_eq!(wb2.sheets.len(), 2);
    assert_eq!(wb2.pivot_caches[0].cache_id, Some(0));
    assert_eq!(wb2.calc_chain.len(), 1);
    assert!(wb2.doc_properties.is_some());
}

// ---------------------------------------------------------------------------
// Part names in another ASCII letter case
// ---------------------------------------------------------------------------

/// The Excel fixture with the entry `from` renamed `to` ([`respelled_in`]).
fn respelled(from: &str, to: &str) -> Vec<u8> {
    respelled_in(EXCEL_EXAMPLE, from, to)
}

/// The package `bytes` with the entry `from` renamed `to`; every reference to
/// it (the relationship targets, the content-type Overrides) is left as it is.
fn respelled_in(bytes: &[u8], from: &str, to: &str) -> Vec<u8> {
    let mut entries: Vec<ZipEntry> = zip_map(bytes)
        .into_iter()
        .map(|(name, data)| ZipEntry {
            name: if name == from { to.to_owned() } else { name },
            data,
        })
        .collect();
    assert!(
        entries.iter().any(|e| e.name == to),
        "{from} is an entry of the fixture"
    );
    entries.sort_by(|a, b| a.name.cmp(&b.name));
    write_zip(&entries).unwrap()
}

/// The entry names of `bytes` that equal `name` ignoring ASCII case.
fn spellings_of(bytes: &[u8], name: &str) -> Vec<String> {
    zip_map(bytes)
        .into_keys()
        .filter(|entry| entry.eq_ignore_ascii_case(name))
        .collect()
}

/// Asserts a package whose part `from` is spelled `to` reads to the model of
/// the fixture, keeps no spelling of `from` as a preserved entry other than
/// one the fixture keeps too, and writes the fixture's bytes, so `from` is
/// written once, in its canonical spelling.
fn assert_reads_as_canonical(from: &str, to: &str) {
    let canonical = read_xlsx(EXCEL_EXAMPLE).unwrap();
    let wb = read_xlsx(&respelled(from, to)).unwrap_or_else(|e| panic!("{to}: {e}"));
    assert_eq!(
        serde_json::to_value(&wb).unwrap(),
        serde_json::to_value(&canonical).unwrap(),
        "{to} reads to the fixture's model"
    );
    let out = write_xlsx(&wb).unwrap();
    assert_eq!(
        out,
        write_xlsx(&canonical).unwrap(),
        "{to} writes the fixture's bytes"
    );
    assert_eq!(spellings_of(&out, from), [from]);
}

#[test]
fn lower_case_content_types_part_is_read_and_written_once() {
    let input = respelled("[Content_Types].xml", "[content_types].xml");
    let wb = read_xlsx(&input).unwrap();
    assert!(
        !wb.preserved_entries
            .keys()
            .any(|path| path.eq_ignore_ascii_case("[Content_Types].xml")),
        "no spelling of the content types part is a preserved entry: {:?}",
        wb.preserved_entries.keys()
    );
    let out = write_xlsx(&wb).unwrap();
    assert_eq!(
        spellings_of(&out, "[Content_Types].xml"),
        ["[Content_Types].xml"]
    );
    let reread = read_xlsx(&out).unwrap();
    assert_eq!(
        serde_json::to_value(&reread).unwrap(),
        serde_json::to_value(&wb).unwrap()
    );
    assert_reads_as_canonical("[Content_Types].xml", "[content_types].xml");
}

#[test]
fn upper_case_content_types_part_is_read_and_written_once() {
    assert_reads_as_canonical("[Content_Types].xml", "[CONTENT_TYPES].XML");
}

#[test]
fn content_types_part_in_two_spellings_reads_the_canonical_one() {
    let mut entries: Vec<ZipEntry> = zip_map(EXCEL_EXAMPLE)
        .into_iter()
        .map(|(name, data)| ZipEntry { name, data })
        .collect();
    entries.push(ZipEntry {
        name: "[content_types].xml".into(),
        data: b"<Types/>".to_vec(),
    });
    let wb = read_xlsx(&write_zip(&entries).unwrap()).unwrap();
    let canonical = read_xlsx(EXCEL_EXAMPLE).unwrap();
    assert_eq!(
        serde_json::to_value(&wb).unwrap(),
        serde_json::to_value(&canonical).unwrap(),
        "the canonical spelling is read and the other one is not preserved"
    );
}

/// The JSON bridge: preserved entries a caller supplies at a case variant
/// (or the exact name) of a part the writer generates give way to it, the
/// content types part included.
#[test]
fn preserved_entries_at_generated_names_in_any_case_give_way_through_json() {
    let mut value = serde_json::to_value(minimal_workbook(
        "S",
        common::minimal_worksheet(vec![number_row(1, "A1", "1")]),
    ))
    .unwrap();
    let junk = json!(b"<x/>".to_vec());
    value["preservedEntries"] = json!({
        "XL/STYLES.XML": junk,
        "xl/Workbook.xml": junk,
        "xl/worksheets/SHEET1.xml": junk,
        "[content_types].xml": junk,
        "[CONTENT_TYPES].XML": junk,
        "[Content_Types].xml": junk,
        "_RELS/.rels": junk,
    });
    let wb: WorkbookData = serde_json::from_value(value).unwrap();
    let out = write_xlsx(&wb).unwrap();
    let names = zip_map(&out);
    let mut lower: Vec<String> = names.keys().map(|n| n.to_ascii_lowercase()).collect();
    lower.sort();
    lower.dedup();
    assert_eq!(
        lower.len(),
        names.len(),
        "no two entries equal ignoring case: {:?}",
        names.keys()
    );
    assert_eq!(
        spellings_of(&out, "[Content_Types].xml"),
        ["[Content_Types].xml"]
    );
    let types = content_types(&names);
    assert!(
        types
            .overrides
            .keys()
            .all(|part| !part.eq_ignore_ascii_case("/[Content_Types].xml")),
        "{:?}",
        types.overrides
    );
    assert_ne!(names["[Content_Types].xml"], b"<x/>");
    assert!(read_xlsx(&out).unwrap().preserved_entries.is_empty());
}

#[test]
fn workbook_part_in_another_case_reads_as_canonical() {
    assert_reads_as_canonical("xl/workbook.xml", "xl/Workbook.xml");
}

#[test]
fn workbook_rels_in_another_case_reads_as_canonical() {
    assert_reads_as_canonical("xl/_rels/workbook.xml.rels", "xl/_rels/workbook.xml.RELS");
}

#[test]
fn shared_strings_part_in_another_case_reads_as_canonical() {
    assert_reads_as_canonical("xl/sharedStrings.xml", "xl/SharedStrings.xml");
}

#[test]
fn styles_part_in_another_case_reads_as_canonical() {
    assert_reads_as_canonical("xl/styles.xml", "xl/Styles.xml");
}

#[test]
fn core_properties_part_in_another_case_reads_as_canonical() {
    assert_reads_as_canonical("docProps/core.xml", "docProps/Core.xml");
}

#[test]
fn app_properties_part_in_another_case_reads_as_canonical() {
    assert_reads_as_canonical("docProps/app.xml", "docProps/App.xml");
}

#[test]
fn theme_part_in_another_case_reads_as_canonical() {
    assert_reads_as_canonical("xl/theme/theme1.xml", "xl/theme/Theme1.xml");
}

#[test]
fn package_rels_in_another_case_reads_as_canonical() {
    assert_reads_as_canonical("_rels/.rels", "_RELS/.rels");
}

#[test]
fn calc_chain_part_in_another_case_reads_as_canonical() {
    let written = write_xlsx(&full_graph_workbook()).unwrap();
    let canonical = read_xlsx(&written).unwrap();
    assert_eq!(canonical.calc_chain.len(), 1);
    let wb = read_xlsx(&respelled_in(
        &written,
        "xl/calcChain.xml",
        "xl/CalcChain.xml",
    ))
    .unwrap();
    assert_eq!(
        serde_json::to_value(&wb).unwrap(),
        serde_json::to_value(&canonical).unwrap()
    );
}

#[test]
fn persons_part_in_another_case_reads_as_canonical() {
    let written = write_xlsx(&full_graph_workbook()).unwrap();
    let canonical = read_xlsx(&written).unwrap();
    assert_eq!(canonical.persons.len(), 1);
    let wb = read_xlsx(&respelled_in(
        &written,
        "xl/persons/person.xml",
        "xl/persons/Person.xml",
    ))
    .unwrap();
    assert_eq!(
        serde_json::to_value(&wb).unwrap(),
        serde_json::to_value(&canonical).unwrap()
    );
}
