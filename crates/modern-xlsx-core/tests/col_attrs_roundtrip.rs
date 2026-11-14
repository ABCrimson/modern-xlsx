//! Column fidelity: `<col style>`, `bestFit` and `phonetic` survive read -> write -> read.
//!
//! Each case wraps the sheet XML under test in a minimal package from `write_xlsx` (six
//! `cellXfs` records, so `style="4"` is a real xf index), reads it through both public
//! readers, writes it back and inspects the written `<cols>` bytes.

use std::io::{Cursor, Read, Write};

use modern_xlsx_core::ooxml::styles::{CellXf, Styles};
use modern_xlsx_core::ooxml::worksheet::ColumnInfo;
use modern_xlsx_core::reader::{read_xlsx, read_xlsx_json};
use modern_xlsx_core::writer::write_xlsx;
use modern_xlsx_core::{SheetData, WorkbookData};
use serde_json::{Value, json};
use zip::write::SimpleFileOptions;
use zip::{ZipArchive, ZipWriter};

const MAIN_NS: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const SHEET_PART: &str = "xl/worksheets/sheet1.xml";

fn sheet(cols: &str) -> String {
    format!("<worksheet xmlns=\"{MAIN_NS}\"><cols>{cols}</cols><sheetData/></worksheet>")
}

/// A package from `write_xlsx` with six `cellXfs` records whose sheet part is `sheet_xml`.
fn package(sheet_xml: &str) -> Vec<u8> {
    let mut styles = Styles::default();
    styles.cell_xfs.extend((0..5).map(|i| CellXf {
        num_fmt_id: 10 + i,
        apply_number_format: true,
        ..CellXf::default()
    }));
    let base = write_xlsx(&WorkbookData {
        sheets: vec![SheetData {
            name: "Sheet1".into(),
            ..Default::default()
        }],
        styles,
        ..Default::default()
    })
    .unwrap();
    let mut archive = ZipArchive::new(Cursor::new(base)).unwrap();
    let mut out = ZipWriter::new(Cursor::new(Vec::new()));
    let mut replaced = false;
    for i in 0..archive.len() {
        let mut file = archive.by_index(i).unwrap();
        let name = file.name().unwrap().into_owned();
        let mut bytes = Vec::new();
        file.read_to_end(&mut bytes).unwrap();
        if name == SHEET_PART {
            bytes = sheet_xml.as_bytes().to_vec();
            replaced = true;
        }
        out.start_file(name, SimpleFileOptions::default()).unwrap();
        out.write_all(&bytes).unwrap();
    }
    assert!(replaced, "base package has no {SHEET_PART}");
    out.finish().unwrap().into_inner()
}

/// The written sheet part of `zip`.
fn sheet_part(zip: &[u8]) -> String {
    let mut archive = ZipArchive::new(Cursor::new(zip)).unwrap();
    let mut xml = String::new();
    archive
        .by_name(SHEET_PART)
        .unwrap()
        .read_to_string(&mut xml)
        .unwrap();
    xml
}

/// The text between `<cols>` and `</cols>` of a written sheet part.
fn cols_xml(xml: &str) -> &str {
    let start = xml.find("<cols>").expect("written sheet has <cols>") + "<cols>".len();
    let end = xml.find("</cols>").expect("written sheet has </cols>");
    &xml[start..end]
}

/// Read `zip` through both public readers, assert their columns agree on the wire, and
/// return the struct-path workbook with the columns as JSON.
fn read_both(zip: &[u8]) -> (WorkbookData, Value) {
    let wb = read_xlsx(zip).unwrap();
    let json: Value = serde_json::from_str(&read_xlsx_json(zip).unwrap()).unwrap();
    let streamed = json["sheets"][0]["worksheet"]["columns"].clone();
    let model = serde_json::to_value(&wb.sheets[0].worksheet.columns).unwrap();
    assert_eq!(streamed, model, "JSON reader and struct reader disagree");
    (wb, model)
}

fn columns(wb: &WorkbookData) -> &[ColumnInfo] {
    &wb.sheets[0].worksheet.columns
}

#[test]
fn style_best_fit_and_phonetic_round_trip() {
    let zip = package(&sheet(
        r#"<col min="2" max="3" width="12.5" style="4" bestFit="1" customWidth="1" phonetic="1"/>"#,
    ));
    let (wb, cols) = read_both(&zip);
    assert_eq!(
        cols,
        json!([{
            "min": 2, "max": 3, "width": 12.5, "hidden": false, "customWidth": true,
            "styleIndex": 4, "bestFit": true, "phonetic": true,
        }])
    );
    let col = &columns(&wb)[0];
    assert_eq!(col.style_index, Some(4));
    assert!(col.best_fit && col.phonetic && col.custom_width && !col.hidden);

    let written = write_xlsx(&wb).unwrap();
    let xml = sheet_part(&written);
    assert_eq!(
        cols_xml(&xml),
        r#"<col min="2" max="3" width="12.5" style="4" customWidth="1" bestFit="1" phonetic="1"/>"#
    );

    let (again, cols_again) = read_both(&written);
    assert_eq!(cols_again, cols);
    let col = &columns(&again)[0];
    assert_eq!(col.style_index, Some(4));
    assert!(col.best_fit && col.phonetic);
}

#[test]
fn style_zero_is_kept_and_written_back() {
    let zip = package(&sheet(r#"<col min="1" max="1" width="9" style="0"/>"#));
    let (wb, cols) = read_both(&zip);
    assert_eq!(cols[0]["styleIndex"], json!(0));
    assert_eq!(columns(&wb)[0].style_index, Some(0));
    let xml = sheet_part(&write_xlsx(&wb).unwrap());
    assert_eq!(
        cols_xml(&xml),
        r#"<col min="1" max="1" width="9" style="0"/>"#
    );
}

#[test]
fn false_flags_read_as_absent() {
    let zip = package(&sheet(
        r#"<col min="1" max="1" width="9" bestFit="0" phonetic="false"/>"#,
    ));
    let (wb, cols) = read_both(&zip);
    assert_eq!(cols[0].get("bestFit"), None);
    assert_eq!(cols[0].get("phonetic"), None);
    assert_eq!(cols[0].get("styleIndex"), None);
    let col = &columns(&wb)[0];
    assert!(!col.best_fit && !col.phonetic && col.style_index.is_none());
}

/// The `<cols>` bytes for columns without the new attributes, as the writer emitted
/// them before `style`, `bestFit` and `phonetic` were modelled.
#[test]
fn columns_without_new_attributes_write_unchanged_bytes() {
    let zip = package(&sheet(concat!(
        r#"<col min="1" max="1" width="15" hidden="1" customWidth="1" outlineLevel="1" collapsed="1"/>"#,
        r#"<col min="2" max="4" width="20.5"/>"#,
    )));
    let (wb, cols) = read_both(&zip);
    for col in cols.as_array().unwrap() {
        for key in ["styleIndex", "bestFit", "phonetic"] {
            assert_eq!(col.get(key), None, "{key} in {col}");
        }
    }
    let xml = sheet_part(&write_xlsx(&wb).unwrap());
    assert_eq!(
        cols_xml(&xml),
        concat!(
            r#"<col min="1" max="1" width="15" customWidth="1" hidden="1" outlineLevel="1" collapsed="1"/>"#,
            r#"<col min="2" max="4" width="20.5"/>"#,
        )
    );
}

#[test]
fn malformed_style_is_xml_parse_error() {
    let zip = package(&sheet(r#"<col min="1" max="1" width="9" style="x"/>"#));
    assert_eq!(read_xlsx(&zip).unwrap_err().code(), "XML_PARSE");
    assert_eq!(read_xlsx_json(&zip).unwrap_err().code(), "XML_PARSE");
}
