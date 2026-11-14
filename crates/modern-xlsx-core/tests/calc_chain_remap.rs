//! calcChain sheetId remap: each entry's source `sheetId` becomes the
//! sheet's 1-based position, via one map built once (entries + sheets work, not
//! entries x sheets), and dropped entries are reported once, by count. The same
//! file pins the pivot cache `cacheId` lookup (one map, first `r:id` wins).

use std::fmt::Write as _;
use std::time::{Duration, Instant};

use modern_xlsx_core::reader::read_xlsx;
use modern_xlsx_core::writer::write_xlsx;
use modern_xlsx_core::zip::{ZipEntry, ZipSecurityLimits, read_zip_entries, write_zip};
use modern_xlsx_core::{SheetData, WorkbookData};

const MAIN_NS: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const SHEET_BODY: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>"#;

/// A package with one sheet per id in `sheet_ids` (tiny sheet parts) and the
/// given calcChain body.
fn package(sheet_ids: &[u32], chain_entries: &str) -> Vec<u8> {
    let base = write_xlsx(&WorkbookData {
        sheets: vec![SheetData {
            name: "S".into(),
            ..Default::default()
        }],
        ..Default::default()
    })
    .expect("write base");
    let mut parts = read_zip_entries(&base, &ZipSecurityLimits::default()).expect("unzip base");

    let mut workbook = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="{MAIN_NS}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>"#
    );
    let mut rels = String::from(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">"#,
    );
    for (n, id) in sheet_ids.iter().enumerate() {
        let n = n + 1;
        write!(
            workbook,
            r#"<sheet name="Sheet{n}" sheetId="{id}" r:id="rId{n}"/>"#
        )
        .unwrap();
        write!(
            rels,
            r#"<Relationship Id="rId{n}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{n}.xml"/>"#
        )
        .unwrap();
        parts.insert(
            format!("xl/worksheets/sheet{n}.xml"),
            SHEET_BODY.as_bytes().to_vec(),
        );
    }
    workbook.push_str("</sheets></workbook>");
    rels.push_str("</Relationships>");
    parts.insert("xl/workbook.xml".into(), workbook.into_bytes());
    parts.insert("xl/_rels/workbook.xml.rels".into(), rels.into_bytes());
    parts.insert(
        "xl/calcChain.xml".into(),
        format!(r#"<calcChain xmlns="{MAIN_NS}">{chain_entries}</calcChain>"#).into_bytes(),
    );
    let entries: Vec<ZipEntry> = parts
        .into_iter()
        .map(|(name, data)| ZipEntry { name, data })
        .collect();
    write_zip(&entries).expect("zip")
}

fn chain(xlsx: &[u8]) -> Vec<(String, u32)> {
    read_xlsx(xlsx)
        .expect("read")
        .calc_chain
        .into_iter()
        .map(|e| (e.cell_ref, e.sheet_id))
        .collect()
}

#[test]
fn source_sheet_ids_map_to_positions() {
    let xlsx = package(&[5, 9], r#"<c r="A1" i="9"/><c r="A2" i="5"/>"#);
    assert_eq!(chain(&xlsx), [("A1".into(), 2), ("A2".into(), 1)]);
}

#[test]
fn entry_naming_an_absent_sheet_id_is_dropped() {
    let xlsx = package(
        &[5, 9],
        r#"<c r="A1" i="9"/><c r="A2" i="77"/><c r="A3" i="5"/>"#,
    );
    assert_eq!(chain(&xlsx), [("A1".into(), 2), ("A3".into(), 1)]);
}

#[test]
fn first_sheet_with_a_shared_sheet_id_wins() {
    let xlsx = package(&[5, 5], r#"<c r="A1" i="5"/><c r="A2"/>"#);
    assert_eq!(chain(&xlsx), [("A1".into(), 1), ("A2".into(), 1)]);
}

/// Regression guard: the read must not be quadratic in sheets x entries. With
/// 20,000 sheets and 200,000 entries naming an absent sheetId, a per-entry scan
/// of the sheets would cost 4e9 comparisons plus one log call per dropped entry.
/// The test asserts the read finishes within 8 s.
#[test]
fn many_sheets_and_many_dropped_entries_read_quickly() {
    let ids: Vec<u32> = (1..=20_000).collect();
    let mut entries = String::with_capacity(200_000 * 24);
    entries.push_str(r#"<c r="A1" i="999999"/>"#);
    for row in 2..=200_000u32 {
        write!(entries, r#"<c r="A{row}"/>"#).unwrap();
    }
    let xlsx = package(&ids, &entries);

    let start = Instant::now();
    let wb = read_xlsx(&xlsx).expect("read");
    let elapsed = start.elapsed();
    assert!(wb.calc_chain.is_empty());
    assert_eq!(wb.sheets.len(), 20_000);
    assert!(
        elapsed < Duration::from_secs(8),
        "read took {elapsed:?}, bound 8 s"
    );
}

/// A package whose workbook lists `pivot_caches` (`(cacheId, r:id)` pairs, in
/// order) and whose workbook rels hold one pivotCacheDefinition relationship per
/// `rel_ids` entry, all targeting one real cache definition part.
fn pivot_package(pivot_caches: &[(u32, &str)], rel_ids: &[&str]) -> Vec<u8> {
    let seed = WorkbookData {
        sheets: vec![SheetData {
            name: "S".into(),
            ..Default::default()
        }],
        pivot_caches: serde_json::from_value(serde_json::json!([
            { "source": { "ref": "A1:B5", "sheet": "S" }, "fields": [{ "name": "A" }] }
        ]))
        .expect("seed cache"),
        pivot_cache_records: serde_json::from_value(serde_json::json!([
            { "records": [[{ "type": "number", "v": 1.0 }]] }
        ]))
        .expect("seed records"),
        ..Default::default()
    };
    let seed = write_xlsx(&seed).expect("write seed");
    let mut parts = read_zip_entries(&seed, &ZipSecurityLimits::default()).expect("unzip seed");
    let definition = parts
        .keys()
        .find(|name| name.starts_with("xl/pivotCache/pivotCacheDefinition"))
        .expect("seed holds a cache definition")
        .clone();
    let target = definition.trim_start_matches("xl/").to_owned();

    let mut workbook = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="{MAIN_NS}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets><pivotCaches>"#
    );
    for (cache_id, r_id) in pivot_caches {
        write!(
            workbook,
            r#"<pivotCache cacheId="{cache_id}" r:id="{r_id}"/>"#
        )
        .unwrap();
    }
    workbook.push_str("</pivotCaches></workbook>");
    let mut rels = String::from(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>"#,
    );
    for id in rel_ids {
        write!(
            rels,
            r#"<Relationship Id="{id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition" Target="{target}"/>"#
        )
        .unwrap();
    }
    rels.push_str("</Relationships>");
    parts.insert("xl/workbook.xml".into(), workbook.into_bytes());
    parts.insert("xl/_rels/workbook.xml.rels".into(), rels.into_bytes());
    parts.insert(
        "xl/worksheets/sheet1.xml".into(),
        SHEET_BODY.as_bytes().to_vec(),
    );
    let entries: Vec<ZipEntry> = parts
        .into_iter()
        .map(|(name, data)| ZipEntry { name, data })
        .collect();
    write_zip(&entries).expect("zip")
}

#[test]
fn pivot_cache_ids_follow_the_relationship_id_and_the_first_entry_wins() {
    // rId9 appears twice (first wins: 7, not 8). rId9, rId10 and rId11 all name
    // the one cache definition of `pivot_package`, which is read once, under
    // the first of them; the reader tests cover rId10 and rId11 over
    // distinct definitions.
    let xlsx = pivot_package(
        &[(7, "rId9"), (8, "rId9"), (3, "rId10")],
        &["rId9", "rId10", "rId11"],
    );
    let ids: Vec<Option<u32>> = read_xlsx(&xlsx)
        .expect("read")
        .pivot_caches
        .iter()
        .map(|cache| cache.cache_id)
        .collect();
    assert_eq!(ids, [Some(7)]);
}
