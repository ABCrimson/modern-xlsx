//! Golden tests for the workbook-level parts: `workbook.xml`, `docProps` and
//! `calcChain.xml`. Each serialises only the subtree under test (see
//! `crate::common::golden_subtree`) from the workbook re-read after a write.

use modern_xlsx_core::SheetData;
use modern_xlsx_core::WorkbookData;
use modern_xlsx_core::dates::DateSystem;
use modern_xlsx_core::ooxml::calc_chain::CalcChainEntry;
use modern_xlsx_core::ooxml::doc_props::DocProperties;
use modern_xlsx_core::ooxml::workbook::{CalcPr, DefinedName, FileVersion, WorkbookPr};
use serde_json::json;

use crate::common::{golden_subtree, number_row, roundtrip};

fn two_sheets() -> Vec<SheetData> {
    ["Data", "R&D"]
        .into_iter()
        .map(|name| SheetData {
            name: name.into(),
            ..Default::default()
        })
        .collect()
}

/// Defined names: a formula with entities, a hidden `_xlnm._FilterDatabase`,
/// and a name with every optional attribute; in a 1904 workbook.
#[test]
fn golden_workbook_defined_names() {
    let wb = WorkbookData {
        sheets: two_sheets(),
        date_system: DateSystem::Date1904,
        defined_names: vec![
            DefinedName {
                name: "Cap".into(),
                value: r#"IF(Data!$A$1<1,"a & b",0)"#.into(),
                ..Default::default()
            },
            DefinedName {
                name: "_xlnm._FilterDatabase".into(),
                value: "Data!$A$1:$C$9".into(),
                sheet_id: Some(0),
                hidden: true,
                ..Default::default()
            },
            DefinedName {
                name: "Macro".into(),
                value: "Data!$B$2".into(),
                comment: Some("R&D <macro>".into()),
                custom_menu: Some("Menu".into()),
                description: Some("Described".into()),
                help: Some("Help".into()),
                status_bar: Some("Status".into()),
                function: true,
                vb_procedure: true,
                xlm: true,
                function_group_id: Some(3),
                shortcut_key: Some("m".into()),
                publish_to_server: true,
                workbook_parameter: true,
                ..Default::default()
            },
        ],
        ..Default::default()
    };
    let back = roundtrip(&wb);
    golden_subtree(
        &json!({
            "dateSystem": back.date_system,
            "definedNames": back.defined_names,
        }),
        "workbook_defined_names",
    );
}

/// docProps (new fields included), calcPr, fileVersion, workbookPr and a
/// calcChain whose sheet ids are sheet positions, over the formula cells
/// Data!A1 and R&D!B2.
#[test]
fn golden_workbook_props() {
    let mut sheets = two_sheets();
    for (sheet, (index, cell_ref)) in sheets.iter_mut().zip([(1, "A1"), (2, "B2")]) {
        let mut row = number_row(index, cell_ref, "1");
        row.cells[0].formula = Some("1+0".into());
        sheet.worksheet.rows = vec![row];
    }
    let wb = WorkbookData {
        sheets,
        doc_properties: Some(DocProperties {
            title: Some("AT&T report".into()),
            creator: Some("Ada".into()),
            last_printed: Some("2024-01-02T03:04:05Z".into()),
            language: Some("en-US".into()),
            identifier: Some("id-1".into()),
            version: Some("1.2".into()),
            application: Some("Microsoft Excel".into()),
            company: Some("R&D <x>".into()),
            total_time: Some("5".into()),
            doc_security: Some("0".into()),
            ..Default::default()
        }),
        calc_pr: Some(CalcPr {
            calc_id: Some(191_029),
            calc_mode: Some("manual".into()),
            full_calc_on_load: Some(true),
            iterate: Some(true),
            iterate_count: Some(50),
            iterate_delta: Some(0.01),
            full_precision: Some(false),
            ..Default::default()
        }),
        file_version: Some(FileVersion {
            app_name: Some("xl".into()),
            last_edited: Some("7".into()),
            lowest_edited: Some("7".into()),
            rup_build: Some("22228".into()),
            ..Default::default()
        }),
        workbook_pr: Some(WorkbookPr {
            code_name: Some("ThisWorkbook".into()),
            default_theme_version: Some(166_925),
            filter_privacy: Some(true),
            ..Default::default()
        }),
        calc_chain: vec![
            CalcChainEntry {
                cell_ref: "A1".into(),
                sheet_id: 1,
                ..Default::default()
            },
            CalcChainEntry {
                cell_ref: "B2".into(),
                sheet_id: 2,
                l: true,
                ..Default::default()
            },
        ],
        ..Default::default()
    };
    let back = roundtrip(&wb);
    golden_subtree(
        &json!({
            "docProperties": back.doc_properties,
            "calcPr": back.calc_pr,
            "fileVersion": back.file_version,
            "workbookPr": back.workbook_pr,
            "calcChain": back.calc_chain,
        }),
        "workbook_props",
    );
}
