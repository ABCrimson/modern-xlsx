//! `<sheetProtection>` (`CT_SheetProtection`, ECMA-376 Part 1 §18.3.1.85) on the wire.
//!
//! The schema defaults (transitional `sml.xsd`): `sheet`, `objects`, `scenarios`,
//! `selectLockedCells` and `selectUnlockedCells` default to false; `formatCells`,
//! `formatColumns`, `formatRows`, `insertColumns`, `insertRows`, `insertHyperlinks`,
//! `deleteColumns`, `deleteRows`, `sort`, `autoFilter` and `pivotTables` default to true.
//! A flag that is true locks its action while the sheet is protected. The reader starts
//! from those defaults; the writer emits an attribute exactly when it differs from its
//! default, in schema attribute order.
//!
//! `password` is the legacy hash (`ST_UnsignedShortHex`, 1 to 4 hex digits): any other
//! value fails the write with `INVALID_WORKBOOK` unless `plain_password` replaces it, and
//! `plain_password` never shows in `Debug` output.

use std::io::{Cursor, Read, Write};

use modern_xlsx_core::ooxml::worksheet::{SheetProtection, WorksheetXml, hash_password_legacy};
use modern_xlsx_core::reader::{read_xlsx, read_xlsx_json};
use modern_xlsx_core::validate::{IssueCategory, Severity, repair_workbook, validate_workbook};
use modern_xlsx_core::writer::write_xlsx;
use modern_xlsx_core::{SheetData, WorkbookData};
use zip::write::SimpleFileOptions;
use zip::{ZipArchive, ZipWriter};

const MAIN_NS: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

/// The element Excel (Microsoft 365, version 16.0 build 20527) wrote on 2026-11-11 for
/// `Worksheet.Protect("secret", DrawingObjects:=True, Contents:=True, Scenarios:=True,
/// AllowInsertingHyperlinks:=True, AllowSorting:=True, AllowFiltering:=True,
/// AllowUsingPivotTables:=True)` followed by `EnableSelection = xlNoSelection`.
const EXCEL_ELEMENT: &str = r#"<sheetProtection algorithmName="SHA-512" hashValue="hB8/iyYvob8HgRJ+NO+YKBgn7lnRwNEUYDX4lmGWZmg0yTpuJMi4yst12SW/XzmmOoQsVFc/OzVezcVyHPe16Q==" saltValue="ss5X428MLP42vrdGaMjWBQ==" spinCount="100000" sheet="1" objects="1" scenarios="1" insertHyperlinks="0" selectLockedCells="1" sort="0" autoFilter="0" pivotTables="0" selectUnlockedCells="1"/>"#;
const EXCEL_HASH: &str =
    "hB8/iyYvob8HgRJ+NO+YKBgn7lnRwNEUYDX4lmGWZmg0yTpuJMi4yst12SW/XzmmOoQsVFc/OzVezcVyHPe16Q==";
const EXCEL_SALT: &str = "ss5X428MLP42vrdGaMjWBQ==";

fn sheet_xml(protection: &str) -> String {
    format!("<worksheet xmlns=\"{MAIN_NS}\"><sheetData/>{protection}</worksheet>")
}

fn parse(protection: &str) -> SheetProtection {
    WorksheetXml::parse(sheet_xml(protection).as_bytes())
        .unwrap()
        .sheet_protection
        .expect("sheetProtection parsed")
}

/// The `<sheetProtection .../>` element `sp` is written as.
fn written(sp: SheetProtection) -> String {
    let ws = WorksheetXml {
        sheet_protection: Some(sp),
        ..Default::default()
    };
    let xml = String::from_utf8(ws.to_xml().unwrap()).unwrap();
    let start = xml.find("<sheetProtection").expect("element written");
    let end = start + xml[start..].find("/>").unwrap() + 2;
    xml[start..end].to_owned()
}

/// A package from `write_xlsx` whose `xl/worksheets/sheet1.xml` is `sheet`.
fn package(sheet: &str) -> Vec<u8> {
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
    for i in 0..archive.len() {
        let mut file = archive.by_index(i).unwrap();
        let name = file.name().unwrap().into_owned();
        let mut bytes = Vec::new();
        file.read_to_end(&mut bytes).unwrap();
        if name == "xl/worksheets/sheet1.xml" {
            bytes = sheet.as_bytes().to_vec();
        }
        out.start_file(name, SimpleFileOptions::default()).unwrap();
        out.write_all(&bytes).unwrap();
    }
    out.finish().unwrap().into_inner()
}

#[test]
fn absent_flags_read_as_schema_defaults_and_write_back_unchanged() {
    let element = r#"<sheetProtection sheet="1" sort="0" autoFilter="0"/>"#;
    let sp = parse(element);
    assert!(sp.sheet);
    assert!(!sp.sort);
    assert!(!sp.auto_filter);
    assert!(
        sp.format_cells,
        "absent formatCells is the schema default true"
    );
    assert!(sp.format_columns && sp.format_rows);
    assert!(sp.insert_columns && sp.insert_rows && sp.insert_hyperlinks);
    assert!(sp.delete_columns && sp.delete_rows && sp.pivot_tables);
    assert!(!sp.objects && !sp.scenarios);
    assert!(!sp.select_locked_cells && !sp.select_unlocked_cells);
    assert_eq!(written(sp), element);
}

#[test]
fn schema_defaults_with_sheet_on_write_a_bare_element() {
    let sp = SheetProtection {
        sheet: true,
        ..Default::default()
    };
    assert_eq!(written(sp), r#"<sheetProtection sheet="1"/>"#);
}

#[test]
fn every_flag_that_differs_from_its_default_is_written_in_schema_order() {
    let sp = SheetProtection {
        sheet: true,
        objects: true,
        scenarios: true,
        format_cells: false,
        format_columns: false,
        format_rows: false,
        insert_columns: false,
        insert_rows: false,
        insert_hyperlinks: false,
        delete_columns: false,
        delete_rows: false,
        select_locked_cells: true,
        sort: false,
        auto_filter: false,
        pivot_tables: false,
        select_unlocked_cells: true,
        password: Some("DAA7".into()),
        ..Default::default()
    };
    assert_eq!(
        written(sp),
        r#"<sheetProtection password="DAA7" sheet="1" objects="1" scenarios="1" formatCells="0" formatColumns="0" formatRows="0" insertColumns="0" insertRows="0" insertHyperlinks="0" deleteColumns="0" deleteRows="0" selectLockedCells="1" sort="0" autoFilter="0" pivotTables="0" selectUnlockedCells="1"/>"#
    );
}

#[test]
fn excel_element_survives_read_write_read() {
    let sp = parse(EXCEL_ELEMENT);
    assert_eq!(sp.algorithm_name.as_deref(), Some("SHA-512"));
    assert_eq!(sp.hash_value.as_deref(), Some(EXCEL_HASH));
    assert_eq!(sp.salt_value.as_deref(), Some(EXCEL_SALT));
    assert_eq!(sp.spin_count, Some(100_000));
    assert!(sp.select_locked_cells && sp.select_unlocked_cells);
    assert!(!sp.insert_hyperlinks && !sp.pivot_tables);
    assert!(sp.password.is_none());
    let out = written(sp.clone());
    assert_eq!(out, EXCEL_ELEMENT);
    let again = parse(&out);
    assert_eq!(
        serde_json::to_value(&again).unwrap(),
        serde_json::to_value(&sp).unwrap()
    );
}

#[test]
fn struct_reader_and_json_reader_agree() {
    for element in [
        EXCEL_ELEMENT,
        r#"<sheetProtection sheet="1" sort="0" autoFilter="0"/>"#,
        r#"<sheetProtection password="DAA7" sheet="true" formatCells="false"/>"#,
    ] {
        let zip = package(&sheet_xml(element));
        let wb = read_xlsx(&zip).unwrap();
        let json = read_xlsx_json(&zip).unwrap();
        let from_json: WorkbookData = serde_json::from_str(&json).unwrap();
        let a = wb.sheets[0].worksheet.sheet_protection.as_ref().unwrap();
        let b = from_json.sheets[0]
            .worksheet
            .sheet_protection
            .as_ref()
            .unwrap();
        assert_eq!(
            serde_json::to_value(a).unwrap(),
            serde_json::to_value(b).unwrap(),
            "{element}"
        );
        assert_eq!(
            serde_json::to_value(a).unwrap(),
            serde_json::to_value(parse(element)).unwrap(),
            "{element}"
        );
    }
}

#[test]
fn json_without_the_new_keys_deserializes_to_schema_defaults() {
    // The shape modern-xlsx 1.2.x serialized.
    let json = r#"{"sheet":true,"objects":false,"scenarios":false,"password":"DAA7",
        "formatCells":true,"formatColumns":true,"formatRows":true,"insertColumns":true,
        "insertRows":true,"deleteColumns":true,"deleteRows":true,"sort":false,"autoFilter":false}"#;
    let sp: SheetProtection = serde_json::from_str(json).unwrap();
    assert!(sp.insert_hyperlinks && sp.pivot_tables);
    assert!(!sp.select_locked_cells && !sp.select_unlocked_cells);
    assert!(sp.algorithm_name.is_none() && sp.hash_value.is_none());
    assert!(sp.salt_value.is_none() && sp.spin_count.is_none());
    assert!(sp.plain_password.is_none());
    // Keys at their schema default are not serialized; the 1.2.x keys always are.
    let back = serde_json::to_value(&sp).unwrap();
    let keys: Vec<&str> = back
        .as_object()
        .unwrap()
        .keys()
        .map(String::as_str)
        .collect();
    for key in [
        "insertHyperlinks",
        "pivotTables",
        "selectLockedCells",
        "selectUnlockedCells",
        "algorithmName",
        "hashValue",
        "saltValue",
        "spinCount",
        "plainPassword",
    ] {
        assert!(!keys.contains(&key), "{key} serialized at its default");
    }
    assert!(keys.contains(&"formatCells") && keys.contains(&"sort"));
}

#[test]
fn plain_password_is_hashed_and_never_echoed() {
    let sp = SheetProtection {
        sheet: true,
        plain_password: Some("secret".into()),
        ..Default::default()
    };
    let json = serde_json::to_string(&sp).unwrap();
    assert!(
        !json.contains("secret"),
        "plainPassword is write-only: {json}"
    );
    let out = written(sp);
    assert!(!out.contains("secret"), "{out}");
    let back = parse(&out);
    assert!(back.plain_password.is_none());
    assert!(back.sheet);
    #[cfg(feature = "encryption")]
    {
        use modern_xlsx_core::ooxml::worksheet::hash_password_sha512;
        assert_eq!(back.algorithm_name.as_deref(), Some("SHA-512"));
        assert_eq!(back.spin_count, Some(100_000));
        assert!(back.password.is_none());
        let salt = back.salt_value.as_deref().unwrap();
        assert_eq!(salt.len(), 24, "16-byte salt in base64: {salt}");
        assert_eq!(
            back.hash_value.as_deref(),
            Some(hash_password_sha512("secret", &decode_base64(salt), 100_000).as_str())
        );
    }
    #[cfg(not(feature = "encryption"))]
    {
        assert_eq!(back.password.as_deref(), Some("DAA7"));
        assert!(back.algorithm_name.is_none() && back.hash_value.is_none());
    }
}

#[test]
fn plain_password_replaces_a_hash_read_from_a_file() {
    let mut sp = parse(EXCEL_ELEMENT);
    sp.password = Some("ABCD".into());
    sp.plain_password = Some("other".into());
    let out = written(sp);
    assert!(!out.contains(EXCEL_HASH), "{out}");
    assert!(!out.contains("ABCD"), "{out}");
    assert!(out.contains(r#"sort="0" autoFilter="0""#), "{out}");
}

/// A one-sheet workbook named `name` whose sheet carries `sp`.
fn workbook_with(name: &str, sp: SheetProtection) -> WorkbookData {
    WorkbookData {
        sheets: vec![SheetData {
            name: name.into(),
            worksheet: WorksheetXml {
                sheet_protection: Some(sp),
                ..Default::default()
            },
            ..Default::default()
        }],
        ..Default::default()
    }
}

#[test]
fn debug_output_redacts_plain_password() {
    let sp = SheetProtection {
        sheet: true,
        password: Some("DAA7".into()),
        sort: false,
        spin_count: Some(100_000),
        plain_password: Some("hunter2".into()),
        ..Default::default()
    };
    let out = format!("{sp:?}");
    assert!(!out.contains("hunter2"), "{out}");
    assert!(
        out.contains(r#"plain_password: Some("<redacted>")"#),
        "{out}"
    );
    for field in [
        "sheet: true",
        "objects: false",
        "scenarios: false",
        r#"password: Some("DAA7")"#,
        "format_cells: true",
        "format_columns: true",
        "format_rows: true",
        "insert_columns: true",
        "insert_rows: true",
        "delete_columns: true",
        "delete_rows: true",
        "sort: false",
        "auto_filter: true",
        "insert_hyperlinks: true",
        "pivot_tables: true",
        "select_locked_cells: false",
        "select_unlocked_cells: false",
        "algorithm_name: None",
        "hash_value: None",
        "salt_value: None",
        "spin_count: Some(100000)",
    ] {
        assert!(out.contains(field), "{field} missing from {out}");
    }
    let unset = format!(
        "{:?}",
        SheetProtection {
            plain_password: None,
            ..sp.clone()
        }
    );
    assert!(unset.contains("plain_password: None"), "{unset}");
    // Every type that embeds a SheetProtection prints it through the same impl.
    let wb = format!("{:#?}", workbook_with("Sheet1", sp));
    assert!(!wb.contains("hunter2"), "{wb}");
    assert!(wb.contains("<redacted>"), "{wb}");
}

#[test]
fn write_rejects_a_password_that_is_not_a_legacy_hash() {
    for bad in ["secret", "", "12345", "DAA7 ", "0x1F", "GHIJ"] {
        let sp = SheetProtection {
            sheet: true,
            password: Some(bad.into()),
            ..Default::default()
        };
        let err = write_xlsx(&workbook_with("Budget 2026", sp.clone())).unwrap_err();
        assert_eq!(err.code(), "INVALID_WORKBOOK", "{bad:?}: {err}");
        let msg = err.to_string();
        assert!(msg.contains("'Budget 2026'"), "{bad:?}: {msg}");
        assert!(msg.contains("plainPassword"), "{bad:?}: {msg}");
        assert!(msg.contains("hash"), "{bad:?}: {msg}");
        if !bad.is_empty() {
            assert!(!msg.contains(bad), "the value may be a password: {msg}");
        }
        // The worksheet writer on its own rejects it too.
        let ws = WorksheetXml {
            sheet_protection: Some(sp),
            ..Default::default()
        };
        let err = ws.to_xml().unwrap_err();
        assert_eq!(err.code(), "INVALID_WORKBOOK", "{bad:?}: {err}");
        assert!(err.to_string().contains("plainPassword"), "{err}");
    }
}

#[test]
fn legacy_hashes_are_written_verbatim() {
    for hash in ["DAA7", "daa7", "1", "CE4B", "0000", "fFfF"] {
        let sp = SheetProtection {
            sheet: true,
            password: Some(hash.into()),
            ..Default::default()
        };
        assert_eq!(
            written(sp.clone()),
            format!(r#"<sheetProtection password="{hash}" sheet="1"/>"#)
        );
        let bytes = write_xlsx(&workbook_with("Sheet1", sp)).unwrap();
        let back = read_xlsx(&bytes).unwrap();
        let read = back.sheets[0].worksheet.sheet_protection.as_ref().unwrap();
        assert_eq!(read.password.as_deref(), Some(hash));
    }
}

#[test]
fn plain_password_replaces_a_non_hash_password() {
    let sp = SheetProtection {
        sheet: true,
        password: Some("secret".into()),
        plain_password: Some("secret".into()),
        ..Default::default()
    };
    let bytes = write_xlsx(&workbook_with("Sheet1", sp.clone())).unwrap();
    let back = read_xlsx(&bytes).unwrap();
    let read = back.sheets[0].worksheet.sheet_protection.as_ref().unwrap();
    assert_ne!(read.password.as_deref(), Some("secret"));
    assert!(!written(sp).contains("secret"));
    let report = validate_workbook(&workbook_with(
        "Sheet1",
        SheetProtection {
            sheet: true,
            password: Some("secret".into()),
            plain_password: Some("secret".into()),
            ..Default::default()
        },
    ));
    assert!(
        report
            .issues
            .iter()
            .all(|i| i.category != IssueCategory::SheetProtection),
        "{:?}",
        report.issues
    );
}

#[test]
fn validate_reports_a_password_that_is_not_a_legacy_hash() {
    let bad = workbook_with(
        "Budget",
        SheetProtection {
            sheet: true,
            password: Some("secret".into()),
            ..Default::default()
        },
    );
    let report = validate_workbook(&bad);
    let issues: Vec<_> = report
        .issues
        .iter()
        .filter(|i| i.category == IssueCategory::SheetProtection)
        .collect();
    assert_eq!(issues.len(), 1, "{:?}", report.issues);
    let issue = issues[0];
    assert_eq!(issue.severity, Severity::Error);
    assert!(!report.is_valid);
    assert_eq!(issue.location, "Budget!sheetProtection");
    assert!(issue.message.contains("'Budget'"), "{}", issue.message);
    assert!(issue.message.contains("plainPassword"), "{}", issue.message);
    assert!(!issue.message.contains("secret"), "{}", issue.message);
    assert!(
        issue.suggestion.contains("plainPassword"),
        "{}",
        issue.suggestion
    );
    assert!(!issue.auto_fixable);
    // The write fails with the same text the report gives.
    let err = write_xlsx(&bad).unwrap_err();
    assert!(err.to_string().contains(&issue.message), "{err}");
    // Repair does not guess: the value stays and the issue stays.
    let mut repaired = bad.clone();
    repair_workbook(&mut repaired);
    let sp = repaired.sheets[0]
        .worksheet
        .sheet_protection
        .as_ref()
        .unwrap();
    assert_eq!(sp.password.as_deref(), Some("secret"));
    assert!(
        validate_workbook(&repaired)
            .issues
            .iter()
            .any(|i| i.category == IssueCategory::SheetProtection)
    );

    let good = workbook_with(
        "Budget",
        SheetProtection {
            sheet: true,
            password: Some("DAA7".into()),
            ..Default::default()
        },
    );
    assert!(
        validate_workbook(&good)
            .issues
            .iter()
            .all(|i| i.category != IssueCategory::SheetProtection)
    );
}

#[test]
fn legacy_hash_matches_excel() {
    // Excel opens a sheet carrying password="DAA7" and Unprotect("secret") succeeds
    // (checked through COM, Excel 16.0 build 20527, 2026-11-11).
    assert_eq!(hash_password_legacy("secret").unwrap(), "DAA7");
    assert_eq!(hash_password_legacy("test").unwrap(), "CBEB");
    assert_eq!(hash_password_legacy("").unwrap(), "CE4B");
    assert!(hash_password_legacy("pässword").is_err());
}

#[cfg(feature = "encryption")]
#[test]
fn sha512_hash_matches_the_value_excel_wrote() {
    use modern_xlsx_core::ooxml::worksheet::hash_password_sha512;
    let salt = decode_base64(EXCEL_SALT);
    assert_eq!(hash_password_sha512("secret", &salt, 100_000), EXCEL_HASH);
    assert_ne!(hash_password_sha512("wrong", &salt, 100_000), EXCEL_HASH);
}

/// Minimal standard-alphabet base64 decoder for the fixtures above.
#[cfg(feature = "encryption")]
fn decode_base64(s: &str) -> Vec<u8> {
    const ALPHABET: &[u8; 64] = b"ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/";
    let mut out = Vec::new();
    let (mut acc, mut bits) = (0u32, 0u32);
    for c in s.bytes().filter(|&c| c != b'=') {
        let v = ALPHABET.iter().position(|&a| a == c).expect("base64 digit") as u32;
        acc = (acc << 6) | v;
        bits += 6;
        if bits >= 8 {
            bits -= 8;
            out.push((acc >> bits) as u8);
            acc &= (1 << bits) - 1;
        }
    }
    out
}
