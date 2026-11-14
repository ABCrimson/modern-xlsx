//! Parity between the two worksheet read paths.
//!
//! `read_xlsx` builds the typed `WorksheetXml` (`parse_with_sst`); `read_xlsx_json` streams
//! JSON straight to the WASM bridge (`parse_to_json`). Both worksheet parsers are
//! crate-private, so each case is wrapped in a minimal package whose `sheet1.xml` (and, for
//! the shared-string case, `sharedStrings.xml`) is replaced with the XML under test. The JSON
//! path must deserialize into `WorkbookData` (no duplicate keys) and its worksheet must
//! serialize to the same `serde_json::Value` as the struct path. Parity alone passes when both
//! paths are equally wrong, so each case also asserts absolute values.

use std::fmt::Write as _;
use std::io::{Cursor, Read, Write};

use modern_xlsx_core::ooxml::worksheet::{Cell, WorksheetXml};
use modern_xlsx_core::reader::{read_xlsx, read_xlsx_json};
use modern_xlsx_core::writer::write_xlsx;
use modern_xlsx_core::{SheetData, WorkbookData};
use zip::write::SimpleFileOptions;
use zip::{ZipArchive, ZipWriter};

const MAIN_NS: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

fn sheet(body: &str) -> Vec<u8> {
    format!("<worksheet xmlns=\"{MAIN_NS}\"><sheetData>{body}</sheetData></worksheet>").into_bytes()
}

/// A package from `write_xlsx` whose `xl/worksheets/sheet1.xml` is `sheet_xml` and, when
/// given, whose `xl/sharedStrings.xml` is `sst_xml`; every other part is copied unchanged.
fn package(sheet_xml: &[u8], sst_xml: Option<&[u8]>) -> Vec<u8> {
    let base = write_xlsx(&WorkbookData {
        sheets: vec![SheetData {
            name: "Sheet1".into(),
            ..Default::default()
        }],
        ..Default::default()
    })
    .unwrap();
    let mut archive = ZipArchive::new(Cursor::new(base)).unwrap();
    let mut out = ZipWriter::new(Cursor::new(Vec::new()));
    let mut replaced_sheet = false;
    let mut replaced_sst = sst_xml.is_none();
    for i in 0..archive.len() {
        let mut file = archive.by_index(i).unwrap();
        let name = file.name().unwrap().into_owned();
        let mut bytes = Vec::new();
        file.read_to_end(&mut bytes).unwrap();
        match (name.as_str(), sst_xml) {
            ("xl/worksheets/sheet1.xml", _) => {
                bytes = sheet_xml.to_vec();
                replaced_sheet = true;
            }
            ("xl/sharedStrings.xml", Some(sst)) => {
                bytes = sst.to_vec();
                replaced_sst = true;
            }
            _ => {}
        }
        out.start_file(name, SimpleFileOptions::default()).unwrap();
        out.write_all(&bytes).unwrap();
    }
    assert!(
        replaced_sheet,
        "base package has no xl/worksheets/sheet1.xml"
    );
    assert!(replaced_sst, "base package has no xl/sharedStrings.xml");
    out.finish().unwrap().into_inner()
}

/// Read `xml` through both paths, assert they agree, and return the struct-path worksheet.
fn parity(xml: &[u8], sst: Option<&[u8]>) -> WorksheetXml {
    let zip = package(xml, sst);
    let mut s = read_xlsx(&zip).unwrap();
    let out = read_xlsx_json(&zip).unwrap();
    let j: WorkbookData =
        serde_json::from_str(&out).expect("JSON path must have no duplicate keys");
    assert_eq!(
        serde_json::to_value(&j.sheets[0].worksheet).unwrap(),
        serde_json::to_value(&s.sheets[0].worksheet).unwrap(),
        "JSON path and struct path diverge; JSON was: {out}"
    );
    s.sheets.swap_remove(0).worksheet
}

fn cell<'a>(ws: &'a WorksheetXml, reference: &str) -> &'a Cell {
    ws.rows
        .iter()
        .flat_map(|r| &r.cells)
        .find(|c| c.reference == reference)
        .unwrap_or_else(|| panic!("cell {reference} not found"))
}

#[test]
fn shared_formula_parity() {
    let xml = sheet(
        "<row r=\"1\"><c r=\"A1\"><f t=\"shared\" ref=\"A1:A2\" si=\"0\">B1*2</f><v>2</v></c></row>\
         <row r=\"2\"><c r=\"A2\"><f t=\"shared\" si=\"0\"/><v>4</v></c></row>",
    );
    let ws = parity(&xml, None);
    let a1 = cell(&ws, "A1");
    assert_eq!(a1.formula_type.as_deref(), Some("shared"));
    assert_eq!(a1.formula_ref.as_deref(), Some("A1:A2"));
    assert_eq!(a1.shared_index, Some(0));
    let a2 = cell(&ws, "A2");
    assert_eq!(a2.formula_type.as_deref(), Some("shared"));
    assert_eq!(a2.formula_ref, None);
    assert_eq!(a2.shared_index, Some(0));
}

#[test]
fn array_formula_parity() {
    let xml = sheet(
        "<row r=\"1\"><c r=\"C1\"><f t=\"array\" ref=\"C1:C2\">B1:B2*2</f><v>2</v></c></row>",
    );
    let ws = parity(&xml, None);
    let c1 = cell(&ws, "C1");
    assert_eq!(c1.formula_type.as_deref(), Some("array"));
    assert_eq!(c1.formula_ref.as_deref(), Some("C1:C2"));
    assert_eq!(c1.formula.as_deref(), Some("B1:B2*2"));
}

#[test]
fn inline_rich_string_parity() {
    let xml = sheet(
        "<row r=\"1\"><c r=\"A1\" t=\"inlineStr\"><is><r><rPr><b/></rPr><t>Bold</t></r>\
         <r><t xml:space=\"preserve\"> tail</t></r></is></c></row>",
    );
    let ws = parity(&xml, None);
    assert_eq!(cell(&ws, "A1").value.as_deref(), Some("Bold tail"));
}

#[test]
fn empty_row_parity() {
    let xml = sheet("<row r=\"3\" ht=\"20\" customHeight=\"1\"/>");
    let ws = parity(&xml, None);
    assert_eq!(ws.rows.len(), 1);
    assert_eq!(ws.rows[0].index, 3);
    assert_eq!(ws.rows[0].height, Some(20.0));
    assert!(ws.rows[0].custom_height);
    assert!(ws.rows[0].cells.is_empty());
}

#[test]
fn styled_empty_cell_parity() {
    let xml = sheet("<row r=\"1\"><c r=\"A1\" s=\"3\"/></row>");
    let ws = parity(&xml, None);
    assert_eq!(cell(&ws, "A1").style_index, Some(3));
}

/// The `cellType` strings of the TypeScript `CellType` union
/// (`packages/modern-xlsx/src/model/types/cells.ts`).
const TS_CELL_TYPES: [&str; 7] = [
    "sharedString",
    "number",
    "boolean",
    "error",
    "formulaStr",
    "inlineStr",
    "stub",
];

/// Every sequence of two children among a date `<v>`, three non-date `<v>`
/// (two also SST indexes: 0 names a plain entry, 1 a rich one), an empty
/// `<v>`, an `<f>` and an `<is>`, on a cell of each `t`, and every sequence
/// of three of them on a `t="s"` cell: the JSON path gives each cell a
/// `cellType` of the TypeScript union and, once its repeated keys are folded
/// (the last one wins, as in `JSON.parse`), the same cell as the struct
/// path. A `t="d"` cell takes its `cellType` and `value` from its last
/// `<v>`: a date gives "number" and the serial, other text "formulaStr" and
/// the text, asserted absolutely on the JSON path for three sequences. A
/// `t="s"` cell takes its `richText` from its last `<v>` alone (an `<is>`
/// after it keeps it), asserted absolutely for five sequences; the JSON
/// text holds one `richText` key for each cell of the struct path that has
/// runs, so no replaced value leaves its runs behind.
#[test]
fn two_part_cell_parity() {
    let parts = [
        ("date", "<v>2024-01-15</v>"),
        ("text", "<v>bad</v>"),
        ("zero", "<v>0</v>"),
        ("rich", "<v>1</v>"),
        ("empty", "<v></v>"),
        ("formula", "<f>A1*2</f>"),
        ("inline", "<is><t>in</t></is>"),
    ];
    let types = ["d", "n", "str", "s", "b", "e", "inlineStr"];
    let mut body = String::new();
    let mut cases = Vec::new();
    for (r, t) in types.iter().enumerate() {
        let row = r + 1;
        write!(body, "<row r=\"{row}\">").unwrap();
        let mut col = 0_u32;
        for (first_name, first) in parts {
            for (second_name, second) in parts {
                col += 1;
                let reference = format!("{}{row}", col_letters(col));
                write!(body, "<c r=\"{reference}\" t=\"{t}\">{first}{second}</c>").unwrap();
                cases.push((reference, format!("t={t} {first_name} then {second_name}")));
            }
        }
        body.push_str("</row>");
    }
    let row = types.len() + 1;
    write!(body, "<row r=\"{row}\">").unwrap();
    let mut col = 0_u32;
    for (first_name, first) in parts {
        for (second_name, second) in parts {
            for (third_name, third) in parts {
                col += 1;
                let reference = format!("{}{row}", col_letters(col));
                write!(
                    body,
                    "<c r=\"{reference}\" t=\"s\">{first}{second}{third}</c>"
                )
                .unwrap();
                cases.push((
                    reference,
                    format!("t=s {first_name} then {second_name} then {third_name}"),
                ));
            }
        }
    }
    body.push_str("</row>");
    let sst = format!(
        "<sst xmlns=\"{MAIN_NS}\" count=\"2\" uniqueCount=\"2\"><si><t>s0</t></si>\
         <si><r><rPr><b/></rPr><t>ri</t></r><r><t>ch</t></r></si></sst>"
    );
    let zip = package(&sheet(&body), Some(sst.as_bytes()));
    let model = read_xlsx(&zip).unwrap();
    let out = read_xlsx_json(&zip).unwrap();
    let json: serde_json::Value =
        serde_json::from_str(&out).unwrap_or_else(|e| panic!("JSON path is not JSON: {e}"));
    let json_cells: Vec<&serde_json::Value> = json["sheets"][0]["worksheet"]["rows"]
        .as_array()
        .unwrap()
        .iter()
        .flat_map(|r| r["cells"].as_array().unwrap())
        .collect();
    let model_ws = &model.sheets[0].worksheet;
    assert_eq!(json_cells.len(), cases.len(), "JSON was: {out}");
    for (reference, label) in &cases {
        let j = json_cells
            .iter()
            .find(|c| c["reference"] == reference.as_str())
            .unwrap_or_else(|| panic!("{label}: cell {reference} missing from the JSON path"));
        let cell_type = j["cellType"].as_str().unwrap_or_default();
        assert!(
            TS_CELL_TYPES.contains(&cell_type),
            "{label}: cellType {cell_type:?} is not a CellType; cell: {j}"
        );
        let m = serde_json::to_value(cell(model_ws, reference)).unwrap();
        assert_eq!(**j, m, "{label}: JSON path and struct path diverge");
    }
    let rich_cells = model_ws
        .rows
        .iter()
        .flat_map(|r| &r.cells)
        .filter(|c| c.rich_text.is_some())
        .count();
    assert_eq!(
        out.matches("\"richText\":").count(),
        rich_cells,
        "one richText key per rich cell; JSON was: {out}"
    );
    let runs = serde_json::json!([{ "bold": true, "text": "ri" }, { "text": "ch" }]);
    for (label, expected) in [
        (
            "t=s rich then text",
            serde_json::json!({ "cellType": "sharedString", "value": "bad" }),
        ),
        (
            "t=s rich then zero",
            serde_json::json!({ "cellType": "sharedString", "value": "s0" }),
        ),
        (
            "t=s zero then rich",
            serde_json::json!({ "cellType": "sharedString", "value": "rich", "richText": runs }),
        ),
        (
            "t=s rich then inline",
            serde_json::json!({
                "cellType": "sharedString", "value": "in", "inlineString": "in", "richText": runs,
            }),
        ),
        (
            "t=s rich then zero then rich",
            serde_json::json!({ "cellType": "sharedString", "value": "rich", "richText": runs }),
        ),
    ] {
        let (reference, _) = cases
            .iter()
            .find(|(_, l)| l == label)
            .unwrap_or_else(|| panic!("no case {label}"));
        let mut expected = expected;
        expected["reference"] = reference.as_str().into();
        let j = json_cells
            .iter()
            .find(|c| c["reference"] == reference.as_str())
            .unwrap();
        assert_eq!(**j, expected, "{label}");
    }
    for (label, cell_type, value) in [
        ("t=d text then text", "formulaStr", "bad"),
        ("t=d text then date", "number", "45306"),
        ("t=d date then text", "formulaStr", "bad"),
    ] {
        let (reference, _) = cases
            .iter()
            .find(|(_, l)| l == label)
            .unwrap_or_else(|| panic!("no case {label}"));
        let j = json_cells
            .iter()
            .find(|c| c["reference"] == reference.as_str())
            .unwrap();
        assert_eq!(
            (j["cellType"].as_str(), j["value"].as_str()),
            (Some(cell_type), Some(value)),
            "{label}: cell {j}"
        );
    }
}

/// The column letters of the 1-based column `col` (at most `ZZ`).
fn col_letters(col: u32) -> String {
    let letter = |n: u32| char::from(b'A' + u8::try_from(n).unwrap());
    if col <= 26 {
        letter(col - 1).to_string()
    } else {
        format!("{}{}", letter((col - 1) / 26 - 1), letter((col - 1) % 26))
    }
}

#[test]
fn shared_string_cell_parity() {
    let sst =
        format!("<sst xmlns=\"{MAIN_NS}\" count=\"1\" uniqueCount=\"1\"><si><t>hi</t></si></sst>");
    let xml = sheet("<row r=\"1\"><c r=\"A1\" t=\"s\"><v>0</v></c></row>");
    let ws = parity(&xml, Some(sst.as_bytes()));
    assert_eq!(cell(&ws, "A1").value.as_deref(), Some("hi"));
}
