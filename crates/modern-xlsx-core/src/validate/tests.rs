use super::*;
use crate::SheetData;
use crate::ooxml::charts::{
    ChartAnchor, ChartData, ChartSeries, ChartTitle, ChartType, WorksheetChart,
};
use crate::ooxml::workbook::DefinedName;
use crate::ooxml::worksheet::{Cell, CellType, Row, WorksheetXml};
use pretty_assertions::assert_eq;

fn empty_worksheet() -> WorksheetXml {
    WorksheetXml {
        ..Default::default()
    }
}

fn minimal_workbook() -> WorkbookData {
    WorkbookData {
        sheets: vec![SheetData {
            name: "Sheet1".into(),
            worksheet: empty_worksheet(),
            ..Default::default()
        }],
        theme_colors: Some(ThemeColors::default()),
        ..Default::default()
    }
}

#[test]
fn test_validate_minimal_workbook() {
    let wb = minimal_workbook();
    let report = validate_workbook(&wb);
    assert!(report.is_valid);
    assert_eq!(report.error_count, 0);
}

#[test]
fn test_validate_empty_workbook() {
    let mut wb = minimal_workbook();
    wb.sheets.clear();
    let report = validate_workbook(&wb);
    assert!(!report.is_valid);
    assert_eq!(report.error_count, 1);
    assert_eq!(report.issues[0].category, IssueCategory::Structure);
}

#[test]
fn test_validate_bad_sheet_name() {
    let mut wb = minimal_workbook();
    wb.sheets[0].name = "Sheet/With*Bad:Chars".into();
    let report = validate_workbook(&wb);
    assert!(!report.is_valid);
    assert!(
        report
            .issues
            .iter()
            .any(|i| i.category == IssueCategory::SheetName)
    );
}

#[test]
fn test_validate_long_sheet_name() {
    let mut wb = minimal_workbook();
    wb.sheets[0].name = "A".repeat(50);
    let report = validate_workbook(&wb);
    assert!(!report.is_valid);
    assert!(
        report
            .issues
            .iter()
            .any(|i| { i.category == IssueCategory::SheetName && i.message.contains("exceeds") })
    );
}

#[test]
fn test_validate_duplicate_sheet_names() {
    let mut wb = minimal_workbook();
    wb.sheets.push(SheetData {
        name: "sheet1".into(), // case-insensitive duplicate
        worksheet: empty_worksheet(),
        ..Default::default()
    });
    let report = validate_workbook(&wb);
    assert!(!report.is_valid);
    assert!(
        report
            .issues
            .iter()
            .any(|i| { i.category == IssueCategory::SheetName && i.message.contains("Duplicate") })
    );
}

#[test]
fn test_validate_dangling_style_index() {
    let mut wb = minimal_workbook();
    wb.sheets[0].worksheet.rows.push(Row {
        index: 1,
        cells: vec![Cell {
            reference: "A1".into(),
            cell_type: CellType::Number,
            style_index: Some(999),
            value: Some("42".into()),
            ..Default::default()
        }],
        ..Default::default()
    });
    let report = validate_workbook(&wb);
    assert!(!report.is_valid);
    assert!(
        report
            .issues
            .iter()
            .any(|i| i.category == IssueCategory::StyleIndex)
    );
}

#[test]
fn test_validate_dangling_font_id() {
    let mut wb = minimal_workbook();
    wb.styles.cell_xfs[0].font_id = 99;
    let report = validate_workbook(&wb);
    assert!(!report.is_valid);
    assert!(
        report
            .issues
            .iter()
            .any(|i| { i.category == IssueCategory::StyleIndex && i.message.contains("fontId") })
    );
}

#[test]
fn test_validate_overlapping_merges() {
    let mut wb = minimal_workbook();
    wb.sheets[0].worksheet.merge_cells = vec!["A1:B2".into(), "B2:C3".into()];
    let report = validate_workbook(&wb);
    assert!(!report.is_valid);
    assert!(
        report
            .issues
            .iter()
            .any(|i| { i.category == IssueCategory::MergeCell && i.message.contains("overlaps") })
    );
}

fn merge_issues(report: &ValidationReport) -> Vec<&ValidationIssue> {
    report
        .issues
        .iter()
        .filter(|i| i.category == IssueCategory::MergeCell)
        .collect()
}

fn with_merges(merges: &[&str]) -> WorkbookData {
    let mut wb = minimal_workbook();
    wb.sheets[0].worksheet.merge_cells = merges.iter().map(|m| (*m).to_owned()).collect();
    wb
}

#[test]
fn test_validate_full_sheet_merge_is_fast() {
    let wb = with_merges(&["A1:XFD1048576"]);
    let start = std::time::Instant::now();
    let report = validate_workbook(&wb);
    let elapsed = start.elapsed();
    assert!(
        elapsed < std::time::Duration::from_secs(2),
        "full-sheet merge validate took {elapsed:?}"
    );
    assert!(merge_issues(&report).is_empty());
}

#[test]
fn test_validate_full_sheet_merge_overlap_once() {
    let report = validate_workbook(&with_merges(&["A1:XFD1048576", "C3:D4"]));
    let issues = merge_issues(&report);
    assert_eq!(issues.len(), 1);
    assert_eq!(issues[0].severity, Severity::Error);
    assert!(issues[0].message.contains("overlaps"));
}

#[test]
fn test_validate_reversed_merge_is_warning() {
    let report = validate_workbook(&with_merges(&["B2:A1"]));
    let issues = merge_issues(&report);
    assert_eq!(issues.len(), 1);
    assert_eq!(issues[0].severity, Severity::Warning);
    assert!(issues[0].auto_fixable);
    assert!(report.is_valid);
}

#[test]
fn test_validate_reversed_and_overlapping_merge() {
    let report = validate_workbook(&with_merges(&["A1:B2", "B2:A1"]));
    let issues = merge_issues(&report);
    assert_eq!(issues.len(), 2);
    assert!(issues.iter().any(|i| i.severity == Severity::Warning));
    assert!(
        issues
            .iter()
            .any(|i| i.severity == Severity::Error && i.message.contains("overlaps"))
    );
}

#[test]
fn test_validate_out_of_sheet_merge_endpoint() {
    let report = validate_workbook(&with_merges(&["A1:XFE1"]));
    let issues = merge_issues(&report);
    assert_eq!(issues.len(), 1);
    assert_eq!(issues[0].severity, Severity::Error);
    assert!(
        issues[0]
            .message
            .starts_with("Cannot parse merge range endpoints in 'A1:XFE1': ")
    );
}

#[test]
fn test_validate_out_of_sheet_cell_reference() {
    let mut wb = minimal_workbook();
    wb.sheets[0].worksheet.rows.push(Row {
        index: 1,
        cells: vec![Cell {
            reference: "XFE1".into(),
            cell_type: CellType::Number,
            value: Some("1".into()),
            ..Default::default()
        }],
        ..Default::default()
    });
    let report = validate_workbook(&wb);
    assert!(report.issues.iter().any(|i| {
        i.category == IssueCategory::CellReference
            && i.severity == Severity::Error
            && i.message.starts_with("Invalid cell reference 'XFE1': ")
            && i.message.contains("exceeds maximum XFD (16384 columns)")
    }));
}

#[test]
fn test_repair_reversed_merge() {
    let mut wb = with_merges(&["B2:A1"]);
    let repairs = repair_workbook(&mut wb);
    assert!(repairs >= 1);
    assert_eq!(wb.sheets[0].worksheet.merge_cells, vec!["A1:B2".to_owned()]);
}

#[test]
fn test_validate_shared_string_no_value() {
    let mut wb = minimal_workbook();
    wb.sheets[0].worksheet.rows.push(Row {
        index: 1,
        cells: vec![Cell {
            reference: "A1".into(),
            cell_type: CellType::SharedString,
            style_index: None,
            value: None,
            ..Default::default()
        }],
        ..Default::default()
    });
    let report = validate_workbook(&wb);
    assert!(!report.is_valid);
    assert!(
        report
            .issues
            .iter()
            .any(|i| i.category == IssueCategory::SharedString)
    );
}

#[test]
fn test_validate_missing_theme() {
    let mut wb = minimal_workbook();
    wb.theme_colors = None;
    let report = validate_workbook(&wb);
    // Missing theme is info-level, not an error.
    assert!(report.is_valid);
    assert!(
        report
            .issues
            .iter()
            .any(|i| i.category == IssueCategory::Theme)
    );
}

#[test]
fn test_validate_bad_date_metadata() {
    let mut wb = minimal_workbook();
    wb.doc_properties = Some(crate::ooxml::doc_props::DocProperties {
        created: Some("not-a-date".into()),
        ..Default::default()
    });
    let report = validate_workbook(&wb);
    assert!(report.is_valid); // Warning, not error.
    assert!(
        report
            .issues
            .iter()
            .any(|i| { i.category == IssueCategory::Metadata && i.message.contains("ISO-8601") })
    );
}

// -- Repair tests --

#[test]
fn test_repair_empty_sheet_name() {
    let mut wb = minimal_workbook();
    wb.sheets[0].name = String::new();
    let repairs = repair_workbook(&mut wb);
    assert!(repairs > 0);
    assert_eq!(wb.sheets[0].name, "Sheet1");
}

#[test]
fn test_repair_forbidden_chars() {
    let mut wb = minimal_workbook();
    wb.sheets[0].name = "My/Sheet*1".into();
    let repairs = repair_workbook(&mut wb);
    assert!(repairs > 0);
    assert_eq!(wb.sheets[0].name, "MySheet1");
}

#[test]
fn test_repair_dangling_style_index() {
    let mut wb = minimal_workbook();
    wb.sheets[0].worksheet.rows.push(Row {
        index: 1,
        cells: vec![Cell {
            reference: "A1".into(),
            cell_type: CellType::Number,
            style_index: Some(999),
            value: Some("1".into()),
            ..Default::default()
        }],
        ..Default::default()
    });
    let repairs = repair_workbook(&mut wb);
    assert!(repairs > 0);
    assert_eq!(wb.sheets[0].worksheet.rows[0].cells[0].style_index, Some(0));
}

#[test]
fn test_repair_missing_styles() {
    let mut wb = minimal_workbook();
    wb.styles.fonts.clear();
    wb.styles.fills.clear();
    wb.styles.borders.clear();
    wb.styles.cell_xfs.clear();
    let repairs = repair_workbook(&mut wb);
    assert!(repairs > 0);
    assert!(!wb.styles.fonts.is_empty());
    assert!(wb.styles.fills.len() >= 2);
    assert!(!wb.styles.borders.is_empty());
    assert!(!wb.styles.cell_xfs.is_empty());
}

#[test]
fn test_repair_missing_theme() {
    let mut wb = minimal_workbook();
    wb.theme_colors = None;
    let repairs = repair_workbook(&mut wb);
    assert!(repairs > 0);
    assert!(wb.theme_colors.is_some());
}

#[test]
fn test_repair_bad_metadata() {
    let mut wb = minimal_workbook();
    wb.doc_properties = Some(crate::ooxml::doc_props::DocProperties {
        created: Some("garbage".into()),
        title: Some("   ".into()),
        ..Default::default()
    });
    let repairs = repair_workbook(&mut wb);
    assert!(repairs > 0);
    assert!(wb.doc_properties.as_ref().unwrap().created.is_none());
    assert!(wb.doc_properties.as_ref().unwrap().title.is_none());
}

#[test]
fn test_repair_row_ordering() {
    let mut wb = minimal_workbook();
    wb.sheets[0].worksheet.rows = vec![
        Row {
            index: 5,
            ..Default::default()
        },
        Row {
            index: 2,
            ..Default::default()
        },
        Row {
            index: 8,
            ..Default::default()
        },
    ];
    let repairs = repair_workbook(&mut wb);
    assert!(repairs > 0);
    let indices: Vec<u32> = wb.sheets[0]
        .worksheet
        .rows
        .iter()
        .map(|r| r.index)
        .collect();
    assert_eq!(indices, vec![2, 5, 8]);
}

#[test]
fn test_w3cdtf_check() {
    for ok in [
        "2024",
        "2024-01",
        "2024-01-15",
        "2024-02-29",
        "2024-01-15T10:30Z",
        "2024-01-15T10:30:00Z",
        "2024-01-15T10:30:00.123Z",
        "2024-01-15T10:30:00+05:30",
        "2024-01-15T23:59:59-08:00",
    ] {
        assert!(is_w3cdtf(ok), "{ok} should be valid");
    }
    for bad in [
        "not-a-date",
        "",
        "20",
        "2024-13",
        "2024-00-10",
        "2024-02-30",
        "2023-02-29",
        "1900-02-29x",
        "2024-01-15 10:30:00Z",
        "2024-01-15t10:30:00Z",
        "2024-01-15T24:00:00Z",
        "2024-01-15T10:60:00Z",
        "2024-01-15T10:30:60Z",
        "2024-01-15T10:30:00.Z",
        "2024-01-15T10:30:00",
        "2024-01-15T10:30:00+5:30",
        "2024-01-15T10",
    ] {
        assert!(!is_w3cdtf(bad), "{bad} should be invalid");
    }
}

fn props_with(created: &str) -> WorkbookData {
    let mut wb = minimal_workbook();
    wb.doc_properties = Some(crate::ooxml::doc_props::DocProperties {
        created: Some(created.into()),
        ..Default::default()
    });
    wb
}

fn metadata_warnings(wb: &WorkbookData) -> usize {
    validate_workbook(wb)
        .issues
        .iter()
        .filter(|i| i.category == IssueCategory::Metadata)
        .count()
}

#[test]
fn test_validate_date_calendar_and_shape() {
    assert_eq!(metadata_warnings(&props_with("2024-02-30")), 1);
    assert_eq!(metadata_warnings(&props_with("2024-01-15 10:30:00Z")), 1);
    assert_eq!(metadata_warnings(&props_with("2024-01-15")), 0);
    assert_eq!(metadata_warnings(&props_with("2024-01-15T10:30:00Z")), 0);
}

#[test]
fn test_repair_date_appends_utc_designator_instead_of_clearing() {
    let mut wb = props_with("2024-01-15T10:30:00");
    assert_eq!(repair_workbook(&mut wb), 1);
    assert_eq!(
        wb.doc_properties.as_ref().unwrap().created.as_deref(),
        Some("2024-01-15T10:30:00Z")
    );
    // Not calendar-valid: cleared.
    let mut wb = props_with("2024-02-30T10:30:00");
    assert!(repair_workbook(&mut wb) > 0);
    assert!(wb.doc_properties.as_ref().unwrap().created.is_none());
}

#[test]
fn test_repair_date_normalises_separator_and_zone_case() {
    for (input, want) in [
        ("2024-01-15 10:30:00Z", "2024-01-15T10:30:00Z"),
        ("2024-01-15t10:30:00", "2024-01-15T10:30:00Z"),
        ("2024-01-15T10:30:00z", "2024-01-15T10:30:00Z"),
    ] {
        let mut wb = props_with(input);
        assert_eq!(repair_workbook(&mut wb), 1, "{input}");
        assert_eq!(
            wb.doc_properties.as_ref().unwrap().created.as_deref(),
            Some(want)
        );
        assert_eq!(repair_workbook(&mut wb), 0, "{input} second pass");
    }
}

#[test]
fn test_validate_whitespace_fields_match_repair() {
    // Every text field validate checks is one repair clears (and vice versa).
    let mut wb = minimal_workbook();
    wb.doc_properties = Some(crate::ooxml::doc_props::DocProperties {
        keywords: Some("   ".into()),
        language: Some(" ".into()),
        total_time: Some(" ".into()),
        ..Default::default()
    });
    assert_eq!(metadata_warnings(&wb), 3);
    assert_eq!(repair_workbook(&mut wb), 3);
    assert_eq!(metadata_warnings(&wb), 0);
}

fn row_with(index: u32, refs: &[(&str, &str)]) -> Row {
    Row {
        index,
        cells: refs
            .iter()
            .map(|(r, v)| Cell {
                reference: (*r).into(),
                cell_type: CellType::Number,
                value: Some((*v).into()),
                ..Default::default()
            })
            .collect(),
        ..Default::default()
    }
}

#[test]
fn test_repair_merges_duplicate_rows() {
    let mut wb = minimal_workbook();
    wb.sheets[0].worksheet.rows = vec![
        row_with(1, &[("B1", "2")]),
        row_with(1, &[("A1", "1"), ("B1", "3")]),
    ];
    assert!(validate_workbook(&wb).warning_count > 0);
    assert!(repair_workbook(&mut wb) > 0);
    let rows = &wb.sheets[0].worksheet.rows;
    assert_eq!(rows.len(), 1);
    let cells: Vec<(&str, Option<&str>)> = rows[0]
        .cells
        .iter()
        .map(|c| (c.reference.as_str(), c.value.as_deref()))
        .collect();
    assert_eq!(cells, vec![("A1", Some("1")), ("B1", Some("3"))]);
    assert_eq!(repair_workbook(&mut wb), 0);
}

#[test]
fn test_validate_merge_overlap_with_full_sheet() {
    let report = validate_workbook(&with_merges(&["A1:XFD1048576", "B2:C3"]));
    assert_eq!(merge_issues(&report).len(), 1);
}

#[test]
fn test_validate_list_dv_without_formula1() {
    let mut wb = minimal_workbook();
    wb.sheets[0]
        .worksheet
        .data_validations
        .push(crate::ooxml::worksheet::DataValidation {
            sqref: "A1".into(),
            validation_type: Some("list".into()),
            operator: None,
            formula1: None,
            formula2: None,
            allow_blank: None,
            show_error_message: None,
            error_title: None,
            error_message: None,
            show_input_message: None,
            prompt_title: None,
            prompt: None,
            error_style: None,
            ime_mode: None,
            show_drop_down: None,
        });
    let report = validate_workbook(&wb);
    let dv: Vec<_> = report
        .issues
        .iter()
        .filter(|i| i.category == IssueCategory::DataValidation)
        .collect();
    assert_eq!(dv.len(), 1);
}

#[test]
fn test_validate_dangling_num_fmt() {
    let mut wb = minimal_workbook();
    wb.styles.cell_xfs[0].num_fmt_id = 170;
    let report = validate_workbook(&wb);
    assert_eq!(report.error_count, 1);
    assert_eq!(report.issues[0].category, IssueCategory::StyleIndex);
    assert!(!report.issues[0].auto_fixable);
    // Built-in ids below 164 never need a numFmt record; a declared one is fine.
    wb.styles.cell_xfs[0].num_fmt_id = 14;
    assert_eq!(validate_workbook(&wb).error_count, 0);
    wb.styles.cell_xfs[0].num_fmt_id = 170;
    wb.styles.num_fmts.push(crate::ooxml::styles::NumFmt {
        id: 170,
        format_code: "0.0".into(),
    });
    assert_eq!(validate_workbook(&wb).error_count, 0);
}

#[test]
fn test_validate_dangling_dxf_ids() {
    use crate::ooxml::worksheet::{ConditionalFormatting, ConditionalFormattingRule};
    let mut wb = minimal_workbook();
    wb.styles.dxfs.push(Default::default());
    wb.sheets[0]
        .worksheet
        .conditional_formatting
        .push(ConditionalFormatting {
            sqref: "A1".into(),
            rules: vec![ConditionalFormattingRule {
                rule_type: "expression".into(),
                priority: 1,
                operator: None,
                formula: Some("TRUE".into()),
                additional_formulas: Vec::new(),
                dxf_id: Some(5),
                stop_if_true: None,
                above_average: None,
                percent: None,
                bottom: None,
                text: None,
                time_period: None,
                rank: None,
                std_dev: None,
                equal_average: None,
                color_scale: None,
                data_bar: None,
                icon_set: None,
            }],
        });
    let report = validate_workbook(&wb);
    assert_eq!(report.error_count, 1);
    assert_eq!(report.issues[0].category, IssueCategory::StyleIndex);
    wb.sheets[0].worksheet.conditional_formatting[0].rules[0].dxf_id = Some(0);
    assert_eq!(validate_workbook(&wb).error_count, 0);

    wb.sheets[0]
        .worksheet
        .tables
        .push(crate::ooxml::tables::TableDefinition {
            id: 1,
            name: None,
            display_name: "T".into(),
            ref_range: "A1:A2".into(),
            header_row_count: 1,
            totals_row_count: 0,
            totals_row_shown: false,
            columns: vec![crate::ooxml::tables::TableColumn {
                id: 1,
                name: "a".into(),
                data_dxf_id: Some(1),
                ..Default::default()
            }],
            style_info: None,
            auto_filter_ref: None,
        });
    assert_eq!(validate_workbook(&wb).error_count, 1);
}

#[test]
fn test_repair_sheet_names_then_valid_and_stable() {
    let mut wb = minimal_workbook();
    let mk = |name: String| SheetData {
        name,
        worksheet: empty_worksheet(),
        ..Default::default()
    };
    wb.sheets = [
        "???".to_owned(),
        "A".repeat(31),
        "A".repeat(31),
        "History".to_owned(),
        "'x'".to_owned(),
        "a\u{1}b".to_owned(),
    ]
    .into_iter()
    .map(mk)
    .collect();
    wb.doc_properties = Some(crate::ooxml::doc_props::DocProperties {
        created: Some("2024-02-30".into()),
        keywords: Some("   ".into()),
        ..Default::default()
    });
    wb.sheets[0].worksheet.rows = vec![row_with(1, &[("A1", "1")]), row_with(1, &[("B1", "2")])];

    assert!(!validate_workbook(&wb).is_valid);
    assert!(repair_workbook(&mut wb) > 0);
    let report = validate_workbook(&wb);
    assert_eq!(report.error_count, 0, "{:?}", report.issues);
    let mut seen = HashSet::new();
    for s in &wb.sheets {
        assert!(s.name.chars().count() <= MAX_SHEET_NAME_LEN, "{}", s.name);
        assert!(seen.insert(s.name.to_lowercase()), "dup {}", s.name);
    }
    assert_eq!(wb.sheets[0].worksheet.rows.len(), 1);
    assert_eq!(wb.sheets[3].name, "History_1");
    assert_eq!(wb.sheets[4].name, "x");
    assert_eq!(wb.sheets[5].name, "ab");
    assert_eq!(repair_workbook(&mut wb), 0);
}

#[test]
fn test_sheet_name_problem_rules() {
    assert!(sheet_name_problem("History").is_some());
    assert!(sheet_name_problem("history").is_some());
    assert!(sheet_name_problem("a\u{7f}b").is_some());
    assert!(sheet_name_problem("it's").is_none());
    assert!(sheet_name_problem("ok name").is_none());
}

#[test]
fn test_repair_default_font_is_calibri() {
    let mut wb = minimal_workbook();
    wb.styles.fonts.clear();
    repair_workbook(&mut wb);
    assert_eq!(wb.styles.fonts[0].name.as_deref(), Some("Calibri"));
}

#[test]
fn test_validate_then_repair_roundtrip() {
    let mut wb = minimal_workbook();
    wb.sheets[0].name = String::new();
    wb.styles.fonts.clear();
    wb.theme_colors = None;

    let report_before = validate_workbook(&wb);
    assert!(!report_before.is_valid);

    let repairs = repair_workbook(&mut wb);
    assert!(repairs > 0);

    let report_after = validate_workbook(&wb);
    assert!(
        report_after.is_valid,
        "After repair: {:?}",
        report_after.issues
    );
}

#[test]
fn test_validate_many_tall_merges_scale() {
    use crate::ooxml::cell::col_to_letters;
    let mut merges: Vec<String> = (0..16_383u32)
        .map(|c| {
            let l = col_to_letters(c);
            format!("{l}1:{l}1048576")
        })
        .collect();
    merges.extend((1..=200_000u32).map(|i| format!("XFD{i}:XFD{i}")));
    let mut wb = minimal_workbook();
    wb.sheets[0].worksheet.merge_cells = merges;
    let start = std::time::Instant::now();
    let report = validate_workbook(&wb);
    assert!(start.elapsed() < std::time::Duration::from_secs(2));
    // XFD (col 16383) is not among the tall columns, so nothing overlaps.
    assert!(merge_issues(&report).is_empty());
}

#[test]
fn test_validate_overlap_reported_in_document_order() {
    let report = validate_workbook(&with_merges(&["C3:D4", "A1:B2", "B2:C3", "E5:E5"]));
    let msgs: Vec<&str> = merge_issues(&report)
        .iter()
        .map(|i| i.message.as_str())
        .collect();
    assert_eq!(msgs.len(), 2);
    assert!(msgs[0].contains("'C3:D4'"));
    assert!(msgs[1].contains("'B2:C3'"));
}

#[test]
fn test_repair_many_duplicate_sheet_names_scales() {
    let mut wb = minimal_workbook();
    wb.sheets = (0..20_000)
        .map(|_| SheetData {
            name: "a".into(),
            worksheet: empty_worksheet(),
            ..Default::default()
        })
        .collect();
    let start = std::time::Instant::now();
    let repairs = repair_workbook(&mut wb);
    assert!(start.elapsed() < std::time::Duration::from_secs(2));
    assert_eq!(repairs, 19_999);
    assert_eq!(wb.sheets[0].name, "a");
    assert_eq!(wb.sheets[1].name, "a_2");
    assert_eq!(wb.sheets[19_999].name, "a_20000");
}

#[test]
fn test_repair_sheet_names_memo_respects_case_folded_stems() {
    // U+0130 lower-cases to "i" + U+0307, so both names share one lower-cased
    // key although their truncated stems fold differently.
    let a = format!("\u{130}{}", "x".repeat(29));
    let b = format!("i\u{307}{}", "x".repeat(29));
    let mut wb = minimal_workbook();
    wb.sheets = [&a, &b, &a]
        .into_iter()
        .map(|name| SheetData {
            name: name.clone(),
            worksheet: empty_worksheet(),
            ..Default::default()
        })
        .collect();
    repair_workbook(&mut wb);
    assert_eq!(wb.sheets[2].name, format!("\u{130}{}_2", "x".repeat(28)));
}

// -- Series c:order across a combo chart --

fn chart_series(idx: u32, order: u32) -> ChartSeries {
    ChartSeries {
        idx,
        order,
        val_ref: "Sheet1!$B$2:$B$5".into(),
        ..Default::default()
    }
}

/// `minimal_workbook` with one column chart titled "Sales" holding `primary`,
/// and a line chart holding `secondary` as its secondary chart.
fn with_combo(primary: Vec<ChartSeries>, secondary: Vec<ChartSeries>) -> WorkbookData {
    let mut wb = minimal_workbook();
    wb.sheets[0].worksheet.charts.push(WorksheetChart {
        chart: ChartData {
            chart_type: ChartType::Column,
            title: Some(ChartTitle {
                text: "Sales".into(),
                ..Default::default()
            }),
            series: primary,
            secondary_chart: Some(Box::new(ChartData {
                chart_type: ChartType::Line,
                series: secondary,
                ..Default::default()
            })),
            ..Default::default()
        },
        anchor: ChartAnchor::default(),
    });
    wb
}

fn order_issues(report: &ValidationReport) -> Vec<&ValidationIssue> {
    report
        .issues
        .iter()
        .filter(|i| i.message.contains("c:order"))
        .collect()
}

#[test]
fn test_duplicate_series_order_across_combo_fails_write_and_validate() {
    let wb = with_combo(vec![chart_series(0, 0)], vec![chart_series(1, 0)]);
    let expected = "the chart at charts[0] of sheet 'Sheet1' (\"Sales\"): series[0] and secondaryChart.series[0] both have c:order 0, which must be unique across the chart's series, including its secondary chart's";

    let err = check_writable(&wb).unwrap_err();
    assert_eq!(err.code(), "INVALID_WORKBOOK");
    assert_eq!(
        err.to_string(),
        format!("invalid workbook: Write failed: {expected}")
    );

    let report = validate_workbook(&wb);
    let issues = order_issues(&report);
    assert_eq!(issues.len(), 1);
    assert_eq!(issues[0].severity, Severity::Error);
    assert_eq!(issues[0].category, IssueCategory::Structure);
    assert_eq!(issues[0].message, expected);
    assert_eq!(issues[0].location, "Sheet1!charts[0]");
    assert!(!report.is_valid);
}

#[test]
fn test_duplicate_series_order_within_one_chart_is_located() {
    let mut wb = with_combo(
        vec![chart_series(0, 0), chart_series(1, 2), chart_series(2, 2)],
        vec![chart_series(3, 3)],
    );
    wb.sheets[0].worksheet.charts[0].chart.title = None;
    let err = check_writable(&wb).unwrap_err();
    assert_eq!(
        err.to_string(),
        "invalid workbook: Write failed: the chart at charts[0] of sheet 'Sheet1': series[1] and series[2] both have c:order 2, which must be unique across the chart's series, including its secondary chart's"
    );
}

#[test]
fn test_unique_series_orders_pass_even_with_a_shared_idx() {
    // Excel opens a chart whose series share only a c:idx (excel-check.ps1, 2026-11-12).
    for wb in [
        with_combo(vec![chart_series(0, 0)], vec![chart_series(1, 1)]),
        with_combo(vec![chart_series(0, 0)], vec![chart_series(0, 1)]),
    ] {
        assert!(check_writable(&wb).is_ok());
        let report = validate_workbook(&wb);
        assert!(order_issues(&report).is_empty());
        assert!(report.is_valid, "{:?}", report.issues);
    }
}

#[test]
fn test_raw_xml_chart_is_not_checked_for_series_order() {
    // The writer emits `raw_xml` verbatim and ignores the model.
    let mut wb = with_combo(vec![chart_series(0, 0)], vec![chart_series(1, 0)]);
    wb.sheets[0].worksheet.charts[0].chart.raw_xml = Some("<c:chartSpace/>".into());
    assert!(check_writable(&wb).is_ok());
    assert!(order_issues(&validate_workbook(&wb)).is_empty());
}

#[test]
fn test_series_order_check_is_linear() {
    // 200,000 series: a pairwise scan would take minutes, a set takes milliseconds.
    let primary = (0..100_000).map(|i| chart_series(i, i)).collect();
    let secondary = (100_000..200_000).map(|i| chart_series(i, i)).collect();
    let mut wb = with_combo(primary, secondary);
    let start = std::time::Instant::now();
    assert!(check_writable(&wb).is_ok());
    assert!(order_issues(&validate_workbook(&wb)).is_empty());
    assert!(start.elapsed() < std::time::Duration::from_secs(2));

    // The last series repeats the first one's order.
    let last = wb.sheets[0].worksheet.charts[0]
        .chart
        .secondary_chart
        .as_mut()
        .and_then(|c| c.series.last_mut())
        .unwrap();
    last.order = 0;
    let err = check_writable(&wb).unwrap_err();
    assert!(
        err.to_string()
            .contains("series[0] and secondaryChart.series[99999] both have c:order 0"),
        "{err}"
    );
}

/// Every colour site the colour check visits, each holding a non-hex `rgb`:
/// the report lists them in traversal order with these exact locations, and
/// the write names the first.
#[test]
fn test_bad_colour_locations_name_every_site() {
    use crate::ooxml::shared_strings::RichTextRun;
    use crate::ooxml::styles::{
        Border, BorderSide, DxfStyle, Fill, Font, GradientFill, GradientStop,
    };
    use crate::ooxml::worksheet::{ConditionalFormatting, SparklineGroup};
    const BAD: &str = "<x>";
    let bad = || Some(String::from(BAD));
    let side = || {
        Some(BorderSide {
            style: "thin".into(),
            color: bad(),
            color_spec: None,
        })
    };
    let font = Font {
        color: bad(),
        ..Default::default()
    };
    let pattern = Fill {
        pattern_type: "solid".into(),
        fg_color: bad(),
        bg_color: bad(),
        ..Default::default()
    };
    let border = Border {
        left: side(),
        right: side(),
        top: side(),
        bottom: side(),
        diagonal: side(),
        ..Default::default()
    };

    let mut wb = minimal_workbook();
    wb.styles.fonts = vec![font.clone()];
    wb.styles.fills = vec![
        pattern.clone(),
        Fill {
            gradient_fill: Some(GradientFill {
                stops: vec![
                    GradientStop {
                        position: 0.0,
                        color: "FF000000".into(),
                        ..Default::default()
                    },
                    GradientStop {
                        position: 1.0,
                        color: BAD.into(),
                        ..Default::default()
                    },
                ],
                ..Default::default()
            }),
            ..Default::default()
        },
    ];
    wb.styles.borders = vec![border.clone()];
    wb.styles.dxfs = vec![DxfStyle {
        font: Some(font),
        fill: Some(pattern),
        border: Some(border),
        ..Default::default()
    }];
    let ws = &mut wb.sheets[0].worksheet;
    ws.tab_color = bad();
    ws.conditional_formatting.push(
        serde_json::from_value::<ConditionalFormatting>(serde_json::json!({
            "sqref": "A1:B2",
            "rules": [
                {
                    "ruleType": "colorScale",
                    "priority": 1,
                    "colorScale": { "cfvos": [], "colors": ["FF000000", BAD] }
                },
                {
                    "ruleType": "dataBar",
                    "priority": 2,
                    "dataBar": { "cfvos": [], "color": BAD }
                }
            ]
        }))
        .unwrap(),
    );
    ws.sparkline_groups.push(SparklineGroup {
        color_series: bad(),
        color_negative: bad(),
        color_axis: bad(),
        color_markers: bad(),
        color_first: bad(),
        color_last: bad(),
        color_high: bad(),
        color_low: bad(),
        ..Default::default()
    });
    ws.rows.push(Row {
        index: 3,
        cells: vec![Cell {
            reference: "C3".into(),
            cell_type: CellType::SharedString,
            value: Some("ab".into()),
            rich_text: Some(vec![
                RichTextRun {
                    text: "a".into(),
                    ..Default::default()
                },
                RichTextRun {
                    text: "b".into(),
                    color: bad(),
                    ..Default::default()
                },
            ]),
            ..Default::default()
        }],
        ..Default::default()
    });

    let expected = [
        "styles.fonts[0].color",
        "styles.fills[0].fgColor",
        "styles.fills[0].bgColor",
        "styles.fills[1].gradientFill.stops[1].color",
        "styles.borders[0].left.color",
        "styles.borders[0].right.color",
        "styles.borders[0].top.color",
        "styles.borders[0].bottom.color",
        "styles.borders[0].diagonal.color",
        "styles.dxfs[0].font.color",
        "styles.dxfs[0].fill.fgColor",
        "styles.dxfs[0].fill.bgColor",
        "styles.dxfs[0].border.left.color",
        "styles.dxfs[0].border.right.color",
        "styles.dxfs[0].border.top.color",
        "styles.dxfs[0].border.bottom.color",
        "styles.dxfs[0].border.diagonal.color",
        "Sheet1!tabColor",
        "Sheet1!A1:B2 cfRule[0].colorScale.colors[1]",
        "Sheet1!A1:B2 cfRule[1].dataBar.color",
        "Sheet1!sparklineGroups[0].colorSeries",
        "Sheet1!sparklineGroups[0].colorNegative",
        "Sheet1!sparklineGroups[0].colorAxis",
        "Sheet1!sparklineGroups[0].colorMarkers",
        "Sheet1!sparklineGroups[0].colorFirst",
        "Sheet1!sparklineGroups[0].colorLast",
        "Sheet1!sparklineGroups[0].colorHigh",
        "Sheet1!sparklineGroups[0].colorLow",
        "Sheet1!C3 richText[1].color",
    ];
    let report = validate_workbook(&wb);
    let colour_issues: Vec<_> = report
        .issues
        .iter()
        .filter(|i| i.message == not_argb_message(BAD, None))
        .collect();
    let locations: Vec<&str> = colour_issues.iter().map(|i| i.location.as_str()).collect();
    assert_eq!(locations, expected);
    let categories: Vec<IssueCategory> = colour_issues.iter().map(|i| i.category).collect();
    let mut want = vec![IssueCategory::StyleIndex; 17];
    want.push(IssueCategory::Structure);
    want.extend([IssueCategory::ConditionalFormatting; 2]);
    want.extend([IssueCategory::Structure; 9]);
    assert_eq!(categories, want);
    let err = check_writable(&wb).unwrap_err();
    assert_eq!(
        err.to_string(),
        format!(
            "invalid workbook: Write failed: {}",
            not_argb_message(BAD, Some("styles.fonts[0].color"))
        )
    );
}

fn with_defined_name(name: &str, value: &str, sheet_id: Option<u32>) -> WorkbookData {
    let mut wb = minimal_workbook();
    wb.defined_names.push(DefinedName {
        name: name.into(),
        value: value.into(),
        sheet_id,
        ..Default::default()
    });
    wb
}

fn defined_name_errors(report: &ValidationReport) -> Vec<&ValidationIssue> {
    report
        .issues
        .iter()
        .filter(|i| i.category == IssueCategory::DefinedName && i.severity == Severity::Error)
        .collect()
}

#[test]
fn test_validate_unqualified_print_area_is_an_error() {
    for value in ["A1:B2", "$A$1:$B$2", "Sheet1!$A$1:$B$2,D1:E2"] {
        let report = validate_workbook(&with_defined_name("_xlnm.Print_Area", value, Some(0)));
        assert!(!report.is_valid, "{value}");
        let errors = defined_name_errors(&report);
        assert_eq!(errors.len(), 1, "{value}");
        assert_eq!(errors[0].location, "definedNames[0]");
        assert!(
            errors[0].message.contains("Print_Area"),
            "{}",
            errors[0].message
        );
    }
}

#[test]
fn test_validate_unqualified_print_titles_is_an_error() {
    let report = validate_workbook(&with_defined_name("_xlnm.Print_Titles", "$1:$2", Some(0)));
    assert!(!report.is_valid);
    assert_eq!(defined_name_errors(&report).len(), 1);
}

#[test]
fn test_validate_qualified_print_names_are_clean() {
    for (name, value) in [
        ("_xlnm.Print_Area", "Sheet1!$A$1:$B$2"),
        ("_xlnm.Print_Area", "'My, Sheet'!A1:B2,Sheet1!$D$1:$E$2"),
        ("_xlnm.Print_Area", "'O''Brien'!A1"),
        ("_xlnm.Print_Titles", "Sheet1!$A:$B,Sheet1!$1:$2"),
    ] {
        let report = validate_workbook(&with_defined_name(name, value, Some(0)));
        assert!(report.is_valid, "{value}: {:?}", report.issues);
    }
}

#[test]
fn test_validate_print_names_match_in_any_letter_case() {
    // Excel matches the built-in names without regard to case: an unqualified
    // variant spelling is refused like the canonical one.
    for name in [
        "_xlnm.print_area",
        "_XLNM.PRINT_AREA",
        "_xlnm.Print_area",
        "_xlnm.print_titles",
        "_XLNM.Print_titles",
    ] {
        let report = validate_workbook(&with_defined_name(name, "A1:B2", Some(0)));
        let errors = defined_name_errors(&report);
        assert_eq!(errors.len(), 1, "{name}");
        assert_eq!(errors[0].location, "definedNames[0]", "{name}");
        assert!(errors[0].message.contains(name), "{}", errors[0].message);
        let qualified = validate_workbook(&with_defined_name(name, "Sheet1!$A$1:$B$2", Some(0)));
        assert!(qualified.is_valid, "{name}: {:?}", qualified.issues);
    }
}

#[test]
fn test_validate_unqualified_other_defined_name_is_not_flagged() {
    let report = validate_workbook(&with_defined_name("Total", "A1:B2", None));
    assert!(defined_name_errors(&report).is_empty());
}

// -- Table names --
//
// The rows below are an Excel 365 oracle table: one workbook
// per name, written by modern-xlsx and opened through COM; every refused row
// failed to open and every accepted row opened. `__tests__/table-names.test.ts`
// runs the same table against `addTable` and `validate()`.

/// Rule fragments of the messages.
const NAME_EMPTY: &str = "empty";
const NAME_LONG: &str = "longer than 255";
const NAME_CHAR: &str = "character Excel refuses";
const NAME_FIRST: &str = "starts with";
const NAME_BOOL: &str = "TRUE and FALSE";
const NAME_A1: &str = "cell reference";
const NAME_R1C1: &str = "R1C1";

/// Every printable ASCII character a table name may not hold anywhere.
const FORBIDDEN_ASCII: &str = " !\"#$%&'()*+,-/:;<=>@[]^`{|}~";

fn refused_table_names() -> Vec<(String, &'static str)> {
    let mut rows: Vec<(String, &'static str)> = vec![
        (String::new(), NAME_EMPTY),
        (format!("T{}", "a".repeat(255)), NAME_LONG),
        ("é".repeat(256), NAME_LONG),
        // UTF-16 code units: 1 + 2 x 128 = 257.
        (format!("T{}", "𝐀".repeat(128)), NAME_LONG),
    ];
    for ch in FORBIDDEN_ASCII.chars() {
        rows.push((format!("Ta{ch}b"), NAME_CHAR));
        rows.push((format!("{ch}Tab"), NAME_CHAR));
    }
    for name in [
        "My Table",
        "My-Table",
        "Ta\tb",
        "Ta\nb",
        "Ta\rb",
        "Ta\u{7f}b",
    ] {
        rows.push((name.into(), NAME_CHAR));
    }
    for name in ["1Table", "123", ".Table", "?Table"] {
        rows.push((name.into(), NAME_FIRST));
    }
    for name in ["TRUE", "FALSE", "true"] {
        rows.push((name.into(), NAME_BOOL));
    }
    for name in [
        "T2",
        "Q1",
        "A1",
        "a1",
        "AB12",
        "xfd1",
        "XFD1048576",
        "A01",
        "xA1",
        "Rx1",
        "Tab123",
        "R1",
        "C1",
        "R01",
        "RC1",
        "C16385",
    ] {
        rows.push((name.into(), NAME_A1));
    }
    for name in [
        "R",
        "C",
        "r",
        "c",
        "RC",
        "rc",
        "R1C",
        "R1C1",
        "r2c3",
        "R1048576C16384",
        "R1C16385",
        "R1C0",
        "R1x",
        "R1_",
        "R1é",
        "R1\\",
        "R1?",
        "R10x",
        "R0001x",
        "R1048576x",
        "R1Cx",
        "R1C.",
        "R1C0x",
        "R1C1x",
        "R1C1_",
        "R1C1.a",
        "R1C1R",
        "R1C16384x",
        "R1C16385x",
        "R1C1048576x",
        "R1R",
        "RC1x",
        "RC16384x",
        "rC1x",
        "Rc1x",
        "C1x",
        "C1_",
        "C01x",
        "C1C",
        "C1R",
        "C1R1",
        "C16384x",
        "r1x",
        "c1x",
    ] {
        rows.push((name.into(), NAME_R1C1));
    }
    rows
}

fn accepted_table_names() -> Vec<String> {
    let mut rows: Vec<String> = vec![
        format!("T{}", "a".repeat(254)),
        "é".repeat(255),
        format!("T{}", "𝐀".repeat(127)),
    ];
    rows.extend(
        [
            "Table1",
            "Sales_2",
            "Ta.b",
            "Ta?b",
            "Ta\\b",
            "Ta_b",
            "My.Table",
            "_Table",
            "\\Table",
            "Tab\\le",
            "\\",
            "Table?",
            "TRUE1",
            "TRUEx",
            "True_",
            "XFE1",
            "ABCD1",
            "A1048577",
            "A0",
            "A1x",
            "A1_",
            "T2x",
            "A1.b",
            "XFD1x",
            "A1é",
            "A1\\",
            "A1?",
            "R0",
            "C0",
            "R0x",
            "C0x",
            "R1048577",
            "R1048577x",
            "R1048577C1",
            "R99999999999x",
            "C16385x",
            "C99999x",
            "RC0",
            "RC16385x",
            "R1.",
            "R1.a",
            "R1048576.x",
            "C1.",
            "C1.x",
            "RC1.",
            "RC1.x",
            "R.",
            "R\\",
            "C?",
            "RC_",
            "RC.",
            "RC\\",
            "RC?",
            "RC.x",
            "RCx",
            "RCC",
            "Rx",
            "Rr",
            "RR",
            "RRx",
            "R_1",
            "CC",
            "CR",
            "CRx",
            "CR1x",
            "Cat",
            "Données",
            "Таблица",
            "売上",
            "Étude",
            "بTab",
            "中Tab",
            "１Tab",
            "Ⅷb",
            "T𝐀b",
            "Tae\u{301}b",
            "Ta€b",
            "Ta·b",
            "·Tab",
            "Ta–b",
            "–Tab",
            "Ta²b",
            "²Tab",
            "Ta٣b",
            "Ta😀b",
            "😀Tab",
            "Ta\u{3000}b",
            "\u{3000}Tab",
            "Ta\u{2002}b",
            "Ta\u{2009}b",
            "Ta\u{200b}b",
            "Ta\u{202f}b",
            "Ta\u{2028}b",
            "Ta\u{feff}b",
            "Ta\u{1680}b",
            "Ta\u{205f}b",
            "Ta\u{ad}b",
            "Ta\u{d7}b",
            "Ta\u{bf}b",
        ]
        .map(String::from),
    );
    rows
}

/// A one-column table `display_name` over `ref_range`, with `name` as its name attribute.
fn named_table(
    id: u32,
    display_name: &str,
    name: Option<&str>,
    ref_range: &str,
) -> crate::ooxml::tables::TableDefinition {
    crate::ooxml::tables::TableDefinition {
        id,
        name: name.map(Into::into),
        display_name: display_name.into(),
        ref_range: ref_range.into(),
        header_row_count: 1,
        totals_row_count: 0,
        totals_row_shown: false,
        columns: vec![crate::ooxml::tables::TableColumn {
            id: 1,
            name: "h".into(),
            ..Default::default()
        }],
        style_info: None,
        auto_filter_ref: None,
    }
}

fn table_errors(report: &ValidationReport) -> Vec<&ValidationIssue> {
    report
        .issues
        .iter()
        .filter(|i| i.severity == Severity::Error && i.location.contains("!table("))
        .collect()
}

#[test]
fn test_validate_table_names_excel_refuses() {
    for (name, rule) in refused_table_names() {
        let mut wb = minimal_workbook();
        wb.sheets[0].worksheet.tables = vec![named_table(1, &name, None, "A1:A2")];
        let report = validate_workbook(&wb);
        let errors = table_errors(&report);
        assert_eq!(errors.len(), 1, "{name:?}: {:?}", report.issues);
        assert_eq!(errors[0].category, IssueCategory::Structure, "{name:?}");
        assert!(
            errors[0].message.contains(rule),
            "{name:?}: {}",
            errors[0].message
        );
        assert_eq!(errors[0].location, format!("Sheet1!table({name})"));
    }
}

#[test]
fn test_validate_table_names_excel_accepts() {
    for name in accepted_table_names() {
        let mut wb = minimal_workbook();
        wb.sheets[0].worksheet.tables = vec![named_table(1, &name, None, "A1:A2")];
        let report = validate_workbook(&wb);
        assert!(report.is_valid, "{name:?}: {:?}", report.issues);
    }
}

#[test]
fn test_validate_table_name_attribute() {
    // Excel opened each of these as the name attribute beside a valid display name.
    let long_ok = format!("T{}", "a".repeat(254));
    for name in [
        "T2",
        "My Table",
        "R",
        "",
        "A&B",
        "1x",
        "TRUE",
        long_ok.as_str(),
    ] {
        let mut wb = minimal_workbook();
        wb.sheets[0].worksheet.tables = vec![named_table(1, "GoodName", Some(name), "A1:A2")];
        let report = validate_workbook(&wb);
        assert!(report.is_valid, "{name:?}: {:?}", report.issues);
    }
    let long = format!("T{}", "a".repeat(255));
    let mut wb = minimal_workbook();
    wb.sheets[0].worksheet.tables = vec![named_table(1, "GoodName", Some(&long), "A1:A2")];
    let report = validate_workbook(&wb);
    let errors = table_errors(&report);
    assert_eq!(errors.len(), 1, "{:?}", report.issues);
    assert!(errors[0].message.contains(NAME_LONG));
}

#[test]
fn test_validate_table_names_must_be_unique() {
    // (first, second, second on another sheet, rule) Excel refused.
    let cases = [
        (
            ("Sales", None),
            ("Sales", None),
            false,
            "already used by table",
        ),
        (
            ("Sales", None),
            ("SALES", None),
            false,
            "already used by table",
        ),
        (
            ("Sales", None),
            ("sales", None),
            true,
            "already used by table",
        ),
        (
            ("Täble", None),
            ("TÄBLE", None),
            false,
            "already used by table",
        ),
        (
            ("First", Some("Same")),
            ("Second", Some("Same")),
            false,
            "already used by name attribute",
        ),
        (
            ("First", Some("same")),
            ("Second", Some("SAME")),
            false,
            "already used by name attribute",
        ),
    ];
    for ((dn1, n1), (dn2, n2), other_sheet, rule) in cases {
        let mut wb = minimal_workbook();
        let second = named_table(2, dn2, n2, "C1:C2");
        wb.sheets[0].worksheet.tables = vec![named_table(1, dn1, n1, "A1:A2")];
        if other_sheet {
            wb.sheets.push(SheetData {
                name: "Sheet2".into(),
                worksheet: WorksheetXml {
                    tables: vec![second],
                    ..Default::default()
                },
                ..Default::default()
            });
        } else {
            wb.sheets[0].worksheet.tables.push(second);
        }
        let report = validate_workbook(&wb);
        let errors = table_errors(&report);
        assert_eq!(errors.len(), 1, "{dn2}: {:?}", report.issues);
        assert!(errors[0].message.contains(rule), "{}", errors[0].message);
        assert_eq!(errors[0].category, IssueCategory::Structure);
    }
}

#[test]
fn test_validate_table_named_like_a_defined_name() {
    for (defined, sheet_id) in [("Rates", None), ("RATES", None), ("Rates", Some(0))] {
        let mut wb = with_defined_name(defined, "Sheet1!$B$1", sheet_id);
        wb.sheets[0].worksheet.tables = vec![named_table(1, "Rates", None, "A1:A2")];
        let report = validate_workbook(&wb);
        let errors = table_errors(&report);
        assert_eq!(errors.len(), 1, "{defined}: {:?}", report.issues);
        assert!(errors[0].message.contains("already used by defined name"));
    }
}

#[test]
fn test_validate_table_names_excel_opened_side_by_side() {
    // A name attribute equal to a defined name, to a later table's display name
    // or (in another case) to its own display name, and a table named like its sheet.
    let mut wb = with_defined_name("Rates", "Sheet1!$B$1", None);
    wb.sheets[0].worksheet.tables = vec![
        named_table(1, "GoodName", Some("Rates"), "A1:A2"),
        named_table(2, "First", Some("Other"), "C1:C2"),
        named_table(3, "Other", None, "E1:E2"),
        named_table(4, "Sheet1", None, "G1:G2"),
        named_table(5, "Own", Some("OWN"), "I1:I2"),
    ];
    let report = validate_workbook(&wb);
    assert!(report.is_valid, "{:?}", report.issues);
}

// -- Table totals cells (checked against Excel 365) --

/// A table column with an optional `totalsRowFunction`, `totalsRowLabel` and `totalsRowFormula`.
fn totals_column(
    id: u32,
    name: &str,
    function: Option<&str>,
    label: Option<&str>,
    formula: Option<&str>,
) -> crate::ooxml::tables::TableColumn {
    crate::ooxml::tables::TableColumn {
        id,
        name: name.into(),
        totals_row_function: function.map(Into::into),
        totals_row_label: label.map(Into::into),
        totals_row_formula: formula.map(Into::into),
        ..Default::default()
    }
}

/// `Item` with the label `Total` (column A of the oracle workbooks).
fn item_total() -> crate::ooxml::tables::TableColumn {
    totals_column(1, "Item", None, Some("Total"), None)
}

/// `Amount` with `totalsRowFunction` `function` (column B of the oracle workbooks).
fn amount(function: &str) -> crate::ooxml::tables::TableColumn {
    totals_column(2, "Amount", Some(function), None, None)
}

/// `Amount` with `totalsRowFunction="custom"` and `formula`.
fn amount_custom(formula: Option<&str>) -> crate::ooxml::tables::TableColumn {
    totals_column(2, "Amount", Some("custom"), None, formula)
}

fn text_cell(reference: &str, text: &str) -> Cell {
    Cell {
        reference: reference.into(),
        cell_type: CellType::SharedString,
        value: Some(text.into()),
        ..Default::default()
    }
}

fn number_cell(reference: &str, number: &str) -> Cell {
    Cell {
        reference: reference.into(),
        cell_type: CellType::Number,
        value: Some(number.into()),
        ..Default::default()
    }
}

fn bool_cell(reference: &str, value: bool) -> Cell {
    Cell {
        reference: reference.into(),
        cell_type: CellType::Boolean,
        value: Some(if value { "1" } else { "0" }.into()),
        ..Default::default()
    }
}

fn formula_cell(reference: &str, formula: &str) -> Cell {
    Cell {
        reference: reference.into(),
        cell_type: CellType::Number,
        formula: Some(formula.into()),
        ..Default::default()
    }
}

/// The oracle workbook: sheet `Sheet1`, table `Table1` over A1:B4 with a totals
/// row, columns `item` and `amount`, and the totals row 4 holding `cells`.
fn totals_workbook(
    item: crate::ooxml::tables::TableColumn,
    amount: crate::ooxml::tables::TableColumn,
    cells: Vec<Cell>,
) -> WorkbookData {
    let mut wb = minimal_workbook();
    let mut table = named_table(1, "Table1", Some("Table1"), "A1:B4");
    table.totals_row_count = 1;
    table.totals_row_shown = true;
    table.columns = vec![item, amount];
    wb.sheets[0].worksheet.tables = vec![table];
    wb.sheets[0].worksheet.rows = vec![Row {
        index: 4,
        cells,
        ..Default::default()
    }];
    wb
}

fn totals_errors(report: &ValidationReport) -> Vec<&ValidationIssue> {
    report
        .issues
        .iter()
        .filter(|i| i.severity == Severity::Error && i.message.contains("totals cell"))
        .collect()
}

const TOTALS_SUBTOTAL: &str = "a formula other than the column's SUBTOTAL";
const TOTALS_CUSTOM: &str = "no formula in a custom column";
const TOTALS_FORMULA: &str = "a formula in a column without a totals function";
const TOTALS_VALUE: &str = "a number or boolean in a column without a totals function";
const TOTALS_LABEL: &str = "text other than the column's totalsRowLabel";

#[test]
fn test_validate_totals_cells_excel_refuses() {
    // Each workbook Excel 365 refused to open, with the cell and rule reported.
    let sum = || amount("sum");
    let none = || totals_column(2, "Amount", None, None, None);
    let item_bare = || totals_column(1, "Item", None, None, None);
    let custom = || amount_custom(Some("SUM(Table1[Amount])*2"));
    let mut cases = vec![
        (item_total(), sum(), text_cell("A4", "mine"), TOTALS_LABEL),
        (item_total(), sum(), text_cell("A4", "total"), TOTALS_LABEL),
        (item_total(), sum(), text_cell("A4", "Total "), TOTALS_LABEL),
        (item_total(), sum(), number_cell("A4", "5"), TOTALS_VALUE),
        (item_total(), sum(), bool_cell("A4", true), TOTALS_VALUE),
        (
            item_total(),
            sum(),
            formula_cell("A4", "\"Total\""),
            TOTALS_FORMULA,
        ),
        (item_bare(), sum(), number_cell("A4", "3"), TOTALS_VALUE),
        (
            item_bare(),
            sum(),
            formula_cell("A4", "\"x\""),
            TOTALS_FORMULA,
        ),
        (
            item_total(),
            none(),
            formula_cell("B4", "SUM(B2:B3)"),
            TOTALS_FORMULA,
        ),
        (
            item_total(),
            amount("none"),
            formula_cell("B4", "SUM(B2:B3)"),
            TOTALS_FORMULA,
        ),
        (item_total(), none(), number_cell("B4", "5"), TOTALS_VALUE),
        (item_total(), none(), bool_cell("B4", true), TOTALS_VALUE),
        (
            item_total(),
            totals_column(2, "Amount", Some("none"), Some("Sum"), None),
            text_cell("B4", "Other"),
            TOTALS_LABEL,
        ),
        (
            item_total(),
            custom(),
            number_cell("B4", "9"),
            TOTALS_CUSTOM,
        ),
        (item_total(), custom(), text_cell("B4", "x"), TOTALS_CUSTOM),
        (
            item_total(),
            custom(),
            bool_cell("B4", false),
            TOTALS_CUSTOM,
        ),
        (
            item_total(),
            amount("average"),
            formula_cell("B4", "SUBTOTAL(109,Table1[Amount])"),
            TOTALS_SUBTOTAL,
        ),
        (
            item_total(),
            totals_column(2, "Q#1", Some("sum"), None, None),
            formula_cell("B4", "SUBTOTAL(109,Table1[[Q#1]])"),
            TOTALS_SUBTOTAL,
        ),
    ];
    for formula in [
        "B2+B3",
        "SUM(B2:B3)",
        "subtotal(109,table1[amount])",
        "SUBTOTAL(109, Table1[Amount])",
        "SUBTOTAL( 109,Table1[Amount] )",
        "SUBTOTAL(101,Table1[Amount])",
        "SUBTOTAL(109,TABLE1[Amount])",
        "SUBTOTAL(109,table1[Amount])",
        "SUBTOTAL(109,Table1[[#Data],[Amount]])",
        "SUBTOTAL(9,Table1[Amount])",
        "SUBTOTAL(109,Table1[Item])",
    ] {
        cases.push((
            item_total(),
            sum(),
            formula_cell("B4", formula),
            TOTALS_SUBTOTAL,
        ));
    }
    for (item, amount, cell, rule) in cases {
        let what = format!("{} {:?} {:?}", cell.reference, cell.value, cell.formula);
        let reference = cell.reference.clone();
        let report = validate_workbook(&totals_workbook(item, amount, vec![cell]));
        let errors = totals_errors(&report);
        assert_eq!(errors.len(), 1, "{what}: {:?}", report.issues);
        assert_eq!(errors[0].category, IssueCategory::Structure, "{what}");
        assert_eq!(errors[0].location, format!("Sheet1!{reference}"), "{what}");
        assert!(
            errors[0].message.contains(rule),
            "{what}: {}",
            errors[0].message
        );
    }
    // A custom column with a totalsRowFormula and no cell at all.
    let report = validate_workbook(&totals_workbook(item_total(), custom(), Vec::new()));
    let errors = totals_errors(&report);
    assert_eq!(errors.len(), 1, "{:?}", report.issues);
    assert_eq!(errors[0].location, "Sheet1!B4");
    assert!(errors[0].message.contains(TOTALS_CUSTOM));
}

#[test]
fn test_validate_totals_cells_excel_accepts() {
    // Each workbook Excel 365 opened: no totals error.
    let sum = || amount("sum");
    let none = || totals_column(2, "Amount", None, None, None);
    let custom = || amount_custom(Some("SUM(Table1[Amount])*2"));
    let mut cases = vec![
        (item_total(), sum(), vec![text_cell("A4", "Total")]),
        (item_total(), sum(), Vec::new()),
        (item_total(), sum(), vec![text_cell("A4", "")]),
        (item_total(), sum(), vec![number_cell("B4", "7")]),
        (item_total(), sum(), vec![text_cell("B4", "n/a")]),
        (item_total(), sum(), vec![bool_cell("B4", true)]),
        (item_total(), sum(), vec![text_cell("B4", "")]),
        (
            totals_column(1, "Item", None, None, None),
            sum(),
            vec![text_cell("A4", "x")],
        ),
        (
            totals_column(1, "Item", Some("count"), Some("Total"), None),
            sum(),
            vec![text_cell("A4", "x")],
        ),
        (item_total(), none(), vec![text_cell("B4", "x")]),
        (item_total(), none(), vec![text_cell("B4", "")]),
        (item_total(), amount("none"), vec![text_cell("B4", "x")]),
        (
            item_total(),
            totals_column(2, "Amount", Some("none"), Some("Sum"), None),
            vec![text_cell("B4", "Sum")],
        ),
        (item_total(), amount_custom(None), Vec::new()),
        (
            item_total(),
            amount_custom(None),
            vec![formula_cell("B4", "SUM(Table1[Amount])")],
        ),
        (
            item_total(),
            amount_custom(None),
            vec![number_cell("B4", "4")],
        ),
        (
            item_total(),
            amount_custom(None),
            vec![text_cell("B4", "x")],
        ),
        (
            item_total(),
            amount("average"),
            vec![formula_cell("B4", "SUBTOTAL(101,Table1[Amount])")],
        ),
        (
            item_total(),
            totals_column(2, "Q#1", Some("sum"), None, None),
            vec![formula_cell("B4", "SUBTOTAL(109,Table1[Q'#1])")],
        ),
        (
            item_total(),
            totals_column(2, "Unit Price", Some("sum"), None, None),
            vec![formula_cell("B4", "SUBTOTAL(109,Table1[Unit Price])")],
        ),
        (
            item_total(),
            totals_column(2, "Unit Price", Some("sum"), None, None),
            vec![formula_cell("B4", "SUBTOTAL(109,Table1[[Unit Price]])")],
        ),
    ];
    for formula in [
        "SUM(Table1[Amount])*2",
        "SUM(Table1[Amount])*3",
        "sum(table1[amount])*2",
        "SUM(Table1[Amount]) * 2",
        "=SUM(Table1[Amount])*2",
    ] {
        cases.push((item_total(), custom(), vec![formula_cell("B4", formula)]));
    }
    for formula in [
        "SUBTOTAL(109,Table1[Amount])",
        "=SUBTOTAL(109,Table1[Amount])",
        "==SUBTOTAL(109,Table1[Amount])",
        "SUBTOTAL(109,[Amount])",
        "SUBTOTAL(109,Table1[amount])",
        "SUBTOTAL(109,[amount])",
        "subtotal(109,Table1[Amount])",
        "SUBTOTAL(109,Table1[AMOUNT])",
        "SUBTOTAL(109,Table1[[Amount]])",
        "SUBTOTAL(109,[[Amount]])",
        "SUBTOTAL(109,Table1[Amount]) ",
        " SUBTOTAL(109,Table1[Amount])",
        "SUBTOTAL(0109,Table1[Amount])",
        "_xlfn.SUBTOTAL(109,Table1[Amount])",
    ] {
        cases.push((item_total(), sum(), vec![formula_cell("B4", formula)]));
    }
    for (item, amount, cells) in cases {
        let what = format!("{cells:?}");
        let report = validate_workbook(&totals_workbook(item, amount, cells));
        assert!(
            totals_errors(&report).is_empty(),
            "{what}: {:?}",
            report.issues
        );
    }
}

#[test]
fn test_validate_totals_cells_excel_made_tables() {
    // The totals rows Excel 365 saved through COM: one column per
    // totalsRowFunction, a custom formula and escaped column names (table
    // Sales), and the cells a user typed into totals cells, which Excel turns
    // into labels and custom formulas (table Edits).
    let mut wb = minimal_workbook();
    let mut sales = named_table(1, "Sales", Some("Sales"), "A1:N4");
    sales.totals_row_count = 1;
    sales.totals_row_shown = true;
    sales.columns = vec![
        totals_column(1, "Item", None, Some("Total"), None),
        totals_column(2, "Sum", Some("sum"), None, None),
        totals_column(3, "Average", Some("average"), None, None),
        totals_column(4, "Count", Some("count"), None, None),
        totals_column(5, "CountNums", Some("countNums"), None, None),
        totals_column(6, "Min", Some("min"), None, None),
        totals_column(7, "Max", Some("max"), None, None),
        totals_column(8, "StdDev", Some("stdDev"), None, None),
        totals_column(9, "Var", Some("var"), None, None),
        totals_column(
            10,
            "Custom",
            Some("custom"),
            None,
            Some("SUM(Sales[Custom])*2"),
        ),
        totals_column(11, "Space Name", Some("sum"), None, None),
        totals_column(12, "Br[ack]et", Some("sum"), None, None),
        totals_column(13, "Hash#", Some("sum"), None, None),
        totals_column(14, "Apos'trophe", Some("sum"), None, None),
    ];
    let sales_row = vec![
        text_cell("A4", "Total"),
        formula_cell("B4", "SUBTOTAL(109,Sales[Sum])"),
        formula_cell("C4", "SUBTOTAL(101,Sales[Average])"),
        formula_cell("D4", "SUBTOTAL(103,Sales[Count])"),
        formula_cell("E4", "SUBTOTAL(102,Sales[CountNums])"),
        formula_cell("F4", "SUBTOTAL(105,Sales[Min])"),
        formula_cell("G4", "SUBTOTAL(104,Sales[Max])"),
        formula_cell("H4", "SUBTOTAL(107,Sales[StdDev])"),
        formula_cell("I4", "SUBTOTAL(110,Sales[Var])"),
        formula_cell("J4", "SUM(Sales[Custom])*2"),
        formula_cell("K4", "SUBTOTAL(109,Sales[Space Name])"),
        formula_cell("L4", "SUBTOTAL(109,Sales[Br'[ack']et])"),
        formula_cell("M4", "SUBTOTAL(109,Sales[Hash'#])"),
        formula_cell("N4", "SUBTOTAL(109,Sales[Apos''trophe])"),
    ];
    let mut edits = named_table(2, "Edits", Some("Edits"), "A6:H9");
    edits.totals_row_count = 1;
    edits.totals_row_shown = true;
    edits.columns = vec![
        totals_column(1, "Label", None, Some("Total"), None),
        totals_column(2, "NoFnNum", None, Some("5"), None),
        totals_column(3, "NoFnBool", None, Some("TRUE"), None),
        totals_column(4, "NoFnFormula", Some("custom"), None, Some("SUM(D7:D8)")),
        totals_column(5, "SumNum", None, Some("7"), None),
        totals_column(6, "SumOther", Some("custom"), None, Some("MAX(F7:F8)")),
        totals_column(7, "SumText", None, Some("n/a"), None),
        totals_column(8, "LabelNum", None, Some("9"), None),
    ];
    let edits_row = vec![
        text_cell("A9", "Total"),
        text_cell("B9", "5"),
        text_cell("C9", "TRUE"),
        formula_cell("D9", "SUM(D7:D8)"),
        text_cell("E9", "7"),
        formula_cell("F9", "MAX(F7:F8)"),
        text_cell("G9", "n/a"),
        text_cell("H9", "9"),
    ];
    let ws = &mut wb.sheets[0].worksheet;
    ws.tables = vec![sales, edits];
    ws.rows = vec![
        Row {
            index: 4,
            cells: sales_row,
            ..Default::default()
        },
        Row {
            index: 9,
            cells: edits_row,
            ..Default::default()
        },
    ];
    let report = validate_workbook(&wb);
    assert!(report.is_valid, "{:?}", report.issues);
}

#[test]
fn test_validate_totals_cells_decode_xstring_escapes() {
    // Excel 365 saves a column name or label with a line break as `_x000a_` in
    // the table part and the real character in the formula and the shared
    // strings; it opens both files, so validate() compares the decoded forms.
    let line_break = totals_workbook(
        item_total(),
        totals_column(2, "Line_x000a_Break", Some("average"), None, None),
        vec![formula_cell("B4", "SUBTOTAL(101,Table1[Line\nBreak])")],
    );
    let grand_total = totals_workbook(
        totals_column(1, "Item", None, Some("Grand_x000a_Total"), None),
        amount("sum"),
        vec![text_cell("A4", "Grand\nTotal")],
    );
    // `_x005f_` decodes to `_` in one left-to-right pass: both sides read `Odd_x0041_Name`.
    let mut odd = totals_workbook(
        item_total(),
        totals_column(2, "Odd_x005f_x0041_Name", Some("sum"), None, None),
        vec![formula_cell("B4", "SUBTOTAL(109,Lb[Odd_x005F_x0041_Name])")],
    );
    let table = &mut odd.sheets[0].worksheet.tables[0];
    table.display_name = "Lb".into();
    table.name = Some("Lb".into());
    for wb in [line_break, grand_total, odd] {
        let report = validate_workbook(&wb);
        assert!(report.is_valid, "{:?}", report.issues);
    }
    // The decoded label still has to match the text.
    let report = validate_workbook(&totals_workbook(
        totals_column(1, "Item", None, Some("Grand_x000a_Total"), None),
        amount("sum"),
        vec![text_cell("A4", "Other")],
    ));
    let errors = totals_errors(&report);
    assert_eq!(errors.len(), 1, "{:?}", report.issues);
    assert_eq!(errors[0].location, "Sheet1!A4");
    assert!(
        errors[0].message.contains(TOTALS_LABEL),
        "{}",
        errors[0].message
    );
}

#[test]
fn test_validate_totals_cells_skip_tables_without_totals_row() {
    // No totals row (or an unparseable ref): the last row's cells are data, never checked.
    for (count, ref_range) in [(0, "A1:B4"), (1, "not a range")] {
        let mut wb = totals_workbook(
            item_total(),
            amount("sum"),
            vec![formula_cell("B4", "B2+B3")],
        );
        let table = &mut wb.sheets[0].worksheet.tables[0];
        table.totals_row_count = count;
        table.ref_range = ref_range.into();
        let report = validate_workbook(&wb);
        assert!(totals_errors(&report).is_empty(), "{:?}", report.issues);
    }
}

#[test]
fn test_validate_totals_cells_lookup_matches_a_search_in_file_order() {
    // The totals cell is the first cell with its reference in the first row
    // with its index, whatever order a hostile file gives rows and cells in.
    let row = |index: u32, cells: Vec<Cell>| Row {
        index,
        cells,
        ..Default::default()
    };
    let cases = [
        // A later row 4 is never read.
        (
            vec![row(4, vec![]), row(4, vec![text_cell("A4", "mine")])],
            vec![],
        ),
        // A later A4 in row 4 is never read.
        (
            vec![row(
                4,
                vec![text_cell("A4", "Total"), text_cell("A4", "mine")],
            )],
            vec![],
        ),
        // Cells in reverse column order, rows out of order: found, reported by column.
        (
            vec![
                row(5, vec![]),
                row(
                    4,
                    vec![formula_cell("B4", "B2+B3"), text_cell("A4", "mine")],
                ),
                row(3, vec![]),
            ],
            vec![("Sheet1!A4", TOTALS_LABEL), ("Sheet1!B4", TOTALS_SUBTOTAL)],
        ),
        // A cell is matched by its reference text, so `a4` is not the totals cell.
        (vec![row(4, vec![text_cell("a4", "mine")])], vec![]),
    ];
    for (rows, expected) in cases {
        let mut wb = totals_workbook(item_total(), amount("sum"), vec![]);
        wb.sheets[0].worksheet.rows = rows;
        let report = validate_workbook(&wb);
        let found: Vec<_> = totals_errors(&report)
            .iter()
            .map(|i| (i.location.as_str(), i.message.as_str()))
            .collect();
        assert_eq!(found.len(), expected.len(), "{found:?}");
        for ((location, message), (want_location, want_message)) in found.iter().zip(&expected) {
            assert_eq!(location, want_location);
            assert!(message.contains(want_message), "{message}");
        }
    }
}

#[test]
fn test_validate_totals_tables_sharing_a_totals_row() {
    // Two tables over one totals row (as a hostile file may give): each is
    // checked against the same cells, in table order.
    let mut wb = totals_workbook(item_total(), amount("sum"), vec![text_cell("A4", "mine")]);
    let mut second = wb.sheets[0].worksheet.tables[0].clone();
    second.id = 2;
    second.display_name = "Table2".into();
    second.name = Some("Table2".into());
    wb.sheets[0].worksheet.tables.push(second);
    let report = validate_workbook(&wb);
    let found: Vec<_> = totals_errors(&report)
        .iter()
        .map(|i| i.message.as_str())
        .collect();
    assert_eq!(found.len(), 2, "{found:?}");
    assert!(
        found[0].starts_with("Table 'Table1' totals cell A4"),
        "{found:?}"
    );
    assert!(
        found[1].starts_with("Table 'Table2' totals cell A4"),
        "{found:?}"
    );
}

// -- Report bounds --

/// The `omittedIssueCount` the report serialises; 0 when the field is absent.
fn omitted(report: &ValidationReport) -> u64 {
    serde_json::to_value(report)
        .unwrap()
        .get("omittedIssueCount")
        .map_or(0, |v| v.as_u64().unwrap())
}

/// Bytes of `message` and `location` text the report stores.
fn stored_text(report: &ValidationReport) -> usize {
    report
        .issues
        .iter()
        .map(|i| i.message.len() + i.location.len())
        .sum()
}

/// `minimal_workbook` with `cells` cells, one per row, each with a dangling style index.
fn dangling_styles(cells: u32) -> WorkbookData {
    let mut wb = minimal_workbook();
    wb.sheets[0].worksheet.rows = (1..=cells)
        .map(|r| Row {
            index: r,
            cells: vec![Cell {
                reference: format!("A{r}"),
                style_index: Some(999),
                ..Default::default()
            }],
            ..Default::default()
        })
        .collect();
    wb
}

#[test]
fn test_report_stores_at_most_10000_issues_and_counts_every_one() {
    // 50,000 problems: 30,000 dangling style indices (errors) and 20,000 data
    // validations without a sqref (warnings).
    let mut wb = dangling_styles(30_000);
    let no_sqref: crate::ooxml::worksheet::DataValidation =
        serde_json::from_str(r#"{"sqref":""}"#).unwrap();
    wb.sheets[0].worksheet.data_validations = vec![no_sqref; 20_000];
    let report = validate_workbook(&wb);
    assert_eq!(
        (report.error_count, report.warning_count, report.info_count),
        (30_000, 20_000, 0)
    );
    assert_eq!(report.issues.len(), MAX_REPORTED_ISSUES);
    assert_eq!(report.issues.len(), 10_000);
    assert_eq!(report.omitted_issue_count, 40_000);
    assert_eq!(omitted(&report), 40_000);
    assert!(!report.is_valid);
    // The first issues found are the ones stored.
    assert_eq!(report.issues[0].location, "Sheet1!A1");
    assert_eq!(report.issues[9_999].location, "Sheet1!A10000");
}

#[test]
fn test_report_is_invalid_when_its_only_error_is_left_out() {
    // 10,000 reversed merge ranges (warnings) fill the report before the
    // empty defined name (an error) is found: isValid still comes from the count.
    let mut wb = minimal_workbook();
    wb.sheets[0].worksheet.merge_cells = (1..=10_000).map(|r| format!("B{r}:A{r}")).collect();
    wb.defined_names = vec![DefinedName {
        name: String::new(),
        value: "Sheet1!$A$1".into(),
        ..Default::default()
    }];
    let report = validate_workbook(&wb);
    assert_eq!((report.error_count, report.warning_count), (1, 10_000));
    assert_eq!(report.issues.len(), 10_000);
    assert!(
        report
            .issues
            .iter()
            .all(|i| i.severity == Severity::Warning)
    );
    assert_eq!(report.omitted_issue_count, 1);
    assert!(!report.is_valid);
}

#[test]
fn test_report_stops_at_the_text_bound() {
    // A sheet name of 100,000 characters that every cell issue repeats.
    let mut wb = dangling_styles(1_000);
    wb.sheets[0].name = "x".repeat(100_000);
    let report = validate_workbook(&wb);
    // The sheet name issue and 1,000 dangling style indices.
    assert_eq!(report.error_count, 1_001);
    let stored = report.issues.len();
    assert!(stored < 1_001, "{stored}");
    assert_eq!(stored as u64 + omitted(&report), 1_001);
    let text = stored_text(&report);
    assert!(text <= MAX_REPORT_TEXT_BYTES, "{text}");
    assert_eq!(MAX_REPORT_TEXT_BYTES, 16 * 1024 * 1024);
    // It stopped at the text bound: one more cell issue would not fit.
    assert!(text + 100_000 > MAX_REPORT_TEXT_BYTES, "{text}");
    assert!(!report.is_valid);
}

#[test]
fn test_report_under_the_bounds_is_unchanged() {
    // The wire shape of a report that stores every issue has no new field.
    let report = validate_workbook(&dangling_styles(3));
    let value = serde_json::to_value(&report).unwrap();
    let keys: Vec<&str> = value
        .as_object()
        .unwrap()
        .keys()
        .map(String::as_str)
        .collect();
    let mut sorted = keys.clone();
    sorted.sort_unstable();
    assert_eq!(
        sorted,
        [
            "errorCount",
            "infoCount",
            "isValid",
            "issues",
            "warningCount"
        ]
    );
    assert_eq!(report.issues.len(), 3);
    assert_eq!(omitted(&report), 0);
    // A report without the field reads back with nothing left out.
    let back: ValidationReport = serde_json::from_value(value).unwrap();
    assert_eq!(back.omitted_issue_count, 0);
    assert_eq!(back.issues.len(), 3);
}

// -- Defined names Excel refuses to open (checked against Excel 365) --
//
// One workbook per name, written by modern-xlsx with the name stored as given
// and opened through COM with alerts off (Excel 365, build 20527,
// 2026-11-12). `__tests__/defined-names.test.ts` runs the same lists against
// `validate()` and, for the names `Names.Add` refused, `addNamedRange`.

/// [name, rule fragment] for each defined name Excel refused to open.
fn refused_defined_names() -> Vec<(String, &'static str)> {
    let mut rows: Vec<(String, &'static str)> = [
        "R",
        "C",
        "r",
        "c",
        "RC",
        "R1C1",
        "R1x",
        "C1_",
        "C1_total",
        "RC2x",
        "R2D2",
        "C3PO",
        "R12abc",
        "RC1data",
        "R01x",
        "C001_",
        "c16384x",
        "r1048576x",
        "R1048576x",
        "R1C",
    ]
    .into_iter()
    .map(|name| (name.to_owned(), NAME_R1C1))
    .collect();
    for name in ["TRUE", "FALSE", "true", "False"] {
        rows.push((name.to_owned(), NAME_BOOL));
    }
    for name in ["1a", "?a", ".a"] {
        rows.push((name.to_owned(), NAME_FIRST));
    }
    // Every ASCII punctuation character but `_ . \ ? :` (`a:b` opened), and the whitespace controls.
    for ch in " !\"#$%&'()*+,-/;<=>@[]^`{|}~\t\n\r\u{7f}".chars() {
        rows.push((format!("a{ch}b"), NAME_CHAR));
    }
    rows.push(("x".repeat(256), NAME_LONG));
    rows
}

/// Defined names Excel opened: an A1 cell reference among them (Excel keeps the name).
fn opened_defined_names() -> Vec<String> {
    let mut rows: Vec<String> = [
        "R1",
        "C1",
        "A1",
        "A01",
        "XFD1",
        "C16385",
        "RC16385",
        "T2",
        "R0x",
        "R1048577",
        "C16385x",
        "A0",
        "XFE1",
        "AAAA1",
        "TRUEx",
        "TRUE.x",
        "R1.a",
        "C1.x",
        "a?b",
        "a\\b",
        "_a",
        "\\a",
        "a.b",
        "Rx",
        "Cx",
        "R_1",
        "C5.x",
        "rc5.x",
        "a€b",
        "Año2024",
        "表1",
        "Таблица1",
        "a·b",
        "a°b",
        "RC0x",
        "C0x",
        "r1048576.x",
        "Rate1",
        "a:b",
        "\\1",
        "\\R1",
        "\\",
        "_",
        "a\u{1}b",
        "_xlnm._FilterDatabase",
        "_xlnm.Criteria",
    ]
    .map(String::from)
    .to_vec();
    rows.push("x".repeat(255));
    rows
}

#[test]
fn test_validate_defined_names_excel_refuses() {
    for (name, rule) in refused_defined_names() {
        let report = validate_workbook(&with_defined_name(&name, "Sheet1!$A$1", None));
        let errors = defined_name_errors(&report);
        assert_eq!(errors.len(), 1, "{name:?}: {:?}", report.issues);
        assert_eq!(errors[0].location, "definedNames[0]", "{name:?}");
        assert!(
            errors[0].message.contains(rule),
            "{name:?}: {}",
            errors[0].message
        );
        assert!(
            errors[0]
                .message
                .starts_with(&format!("Invalid defined name '{name}': ")),
            "{}",
            errors[0].message
        );
    }
}

#[test]
fn test_validate_defined_names_excel_opens() {
    for name in opened_defined_names() {
        let report = validate_workbook(&with_defined_name(&name, "Sheet1!$A$1", Some(0)));
        assert!(report.is_valid, "{name:?}: {:?}", report.issues);
    }
}

#[test]
fn test_validate_defined_name_messages() {
    // One case per message.
    for (name, message) in [
        (
            "R2D2",
            "Invalid defined name 'R2D2': it is an R1C1 reference or starts with one",
        ),
        (
            "TRUE",
            "Invalid defined name 'TRUE': TRUE and FALSE are reserved",
        ),
        (
            "1a",
            "Invalid defined name '1a': it starts with a digit, '.' or '?'",
        ),
        (
            "a b",
            "Invalid defined name 'a b': it holds a character Excel refuses (a space, tab, line break, DEL, or ASCII punctuation other than '_', '.', '\\' and '?')",
        ),
    ] {
        let report = validate_workbook(&with_defined_name(name, "Sheet1!$A$1", None));
        let errors = defined_name_errors(&report);
        assert_eq!(errors.len(), 1, "{name}");
        assert_eq!(errors[0].message, message);
        assert_eq!(
            errors[0].suggestion,
            "Rename the defined name to a name Excel opens"
        );
    }
    let long = "x".repeat(256);
    let report = validate_workbook(&with_defined_name(&long, "Sheet1!$A$1", None));
    assert_eq!(
        defined_name_errors(&report)[0].message,
        format!("Invalid defined name '{long}': it is longer than 255 characters")
    );
    // The empty name keeps its own message, once.
    let report = validate_workbook(&with_defined_name("", "Sheet1!$A$1", None));
    let errors = defined_name_errors(&report);
    assert_eq!(errors.len(), 1);
    assert_eq!(errors[0].message, "Defined name has empty name");
}

// -- The `@` this-row shorthand in a table definition (checked against Excel 365) --
//
// One workbook per form, written by modern-xlsx with the form stored as given,
// opened through COM with alerts off (Excel 365, build 20527, 2026-11-12):
// every form with `[@` in `<calculatedColumnFormula>` or `<totalsRowFormula>`
// was refused, the others opened. `__tests__/table.test.ts` runs the same
// lists against `addTable` and `validate()`.

/// [formula, shorthand reference, the file form Excel writes] refused as a calculatedColumnFormula.
const CALC_REFUSED: [(&str, &str, &str); 8] = [
    ("[@Qty]*[@Price]", "[@Qty]", "CalcT[[#This Row],[Qty]]"),
    (
        "[@[Unit Cost]]*2",
        "[@[Unit Cost]]",
        "CalcT[[#This Row],[Unit Cost]]",
    ),
    ("CalcT[@Qty]*2", "CalcT[@Qty]", "CalcT[[#This Row],[Qty]]"),
    (
        "CalcT[@[Unit Cost]]*2",
        "CalcT[@[Unit Cost]]",
        "CalcT[[#This Row],[Unit Cost]]",
    ),
    ("COLUMNS([@])", "[@]", "CalcT[#This Row]"),
    ("COLUMNS(CalcT[@])", "CalcT[@]", "CalcT[#This Row]"),
    ("[@qty]*2", "[@qty]", "CalcT[[#This Row],[qty]]"),
    ("\"[@x]\"&[@Qty]", "[@Qty]", "CalcT[[#This Row],[Qty]]"),
];

/// calculatedColumnFormula forms Excel opened.
const CALC_ACCEPTED: [&str; 10] = [
    "SUM([Qty])",
    "[[#This Row],[Qty]]*2",
    "CalcT[[#This Row],[Qty]]*2",
    "CalcT[[#This Row],[Qty]]*CalcT[[#This Row],[Price]]",
    "SUM(CalcT[Qty])",
    "\"[@Qty]\"&CalcT[[#This Row],[Qty]]",
    "calct[[#this row],[qty]]*2",
    "COLUMNS([#This Row])",
    "COLUMNS(CalcT[#This Row])",
    "=CalcT[[#This Row],[Qty]]*CalcT[[#This Row],[Price]]",
];

/// Table `CalcT` on A1:D3 whose `Total` column has `calculated` and whose `Qty` column has a custom `totals`.
fn calc_table(
    calculated: Option<&str>,
    totals: Option<&str>,
) -> crate::ooxml::tables::TableDefinition {
    let mut table = named_table(1, "CalcT", Some("CalcT"), "A1:D3");
    table.columns = vec![
        crate::ooxml::tables::TableColumn {
            id: 1,
            name: "Qty".into(),
            totals_row_function: totals.map(|_| "custom".into()),
            totals_row_formula: totals.map(Into::into),
            ..Default::default()
        },
        crate::ooxml::tables::TableColumn {
            id: 2,
            name: "Price".into(),
            ..Default::default()
        },
        crate::ooxml::tables::TableColumn {
            id: 3,
            name: "Unit Cost".into(),
            ..Default::default()
        },
        crate::ooxml::tables::TableColumn {
            id: 4,
            name: "Total".into(),
            calculated_column_formula: calculated.map(Into::into),
            ..Default::default()
        },
    ];
    table
}

#[test]
fn test_validate_calculated_column_shorthand_excel_refuses() {
    for (formula, reference, file_form) in CALC_REFUSED {
        let mut wb = minimal_workbook();
        wb.sheets[0].worksheet.tables = vec![calc_table(Some(formula), None)];
        let report = validate_workbook(&wb);
        let errors = table_errors(&report);
        assert_eq!(errors.len(), 1, "{formula}: {:?}", report.issues);
        assert_eq!(errors[0].category, IssueCategory::Structure);
        assert_eq!(errors[0].location, "Sheet1!table(CalcT)");
        assert_eq!(
            errors[0].message,
            format!(
                "Invalid calculatedColumnFormula of column 'Total' in table 'CalcT': Excel refuses the @ shorthand {reference} in a table definition; write {file_form} instead"
            )
        );
    }
}

#[test]
fn test_validate_calculated_column_forms_excel_opens() {
    for formula in CALC_ACCEPTED {
        let mut wb = minimal_workbook();
        wb.sheets[0].worksheet.tables = vec![calc_table(Some(formula), None)];
        let report = validate_workbook(&wb);
        assert!(report.is_valid, "{formula}: {:?}", report.issues);
    }
}

#[test]
fn test_validate_totals_row_formula_shorthand() {
    // Refused as a totalsRowFormula: [@Qty] and CalcT[@Qty]; opened: the other forms.
    for (formula, reference) in [
        ("SUM([@Qty])", "[@Qty]"),
        ("SUM(CalcT[@Qty])", "CalcT[@Qty]"),
    ] {
        let mut wb = minimal_workbook();
        wb.sheets[0].worksheet.tables = vec![calc_table(None, Some(formula))];
        let report = validate_workbook(&wb);
        let errors: Vec<_> = table_errors(&report)
            .into_iter()
            .filter(|e| e.message.contains("shorthand"))
            .collect();
        assert_eq!(errors.len(), 1, "{formula}: {:?}", report.issues);
        assert_eq!(
            errors[0].message,
            format!(
                "Invalid totalsRowFormula of column 'Qty' in table 'CalcT': Excel refuses the @ shorthand {reference} in a table definition; write CalcT[[#This Row],[Qty]] instead"
            )
        );
    }
    for formula in [
        "SUM([Qty])",
        "SUM(CalcT[Qty])",
        "SUM([[#This Row],[Qty]])",
        "CalcT[[#Totals],[Price]]*2",
        "[[#Totals],[Price]]*2",
        "SUM([[#Data],[Qty]])",
    ] {
        let mut wb = minimal_workbook();
        wb.sheets[0].worksheet.tables = vec![calc_table(None, Some(formula))];
        let report = validate_workbook(&wb);
        assert!(
            !report
                .issues
                .iter()
                .any(|i| i.message.contains("shorthand")),
            "{formula}: {:?}",
            report.issues
        );
    }
}

#[test]
fn test_this_row_shorthand_scan() {
    let none = |_: &str| -> &'static [crate::ooxml::tables::TableColumn] { &[] };
    // Escaped brackets inside a column item, quoted names and an unclosed quote.
    assert!(this_row_shorthand("Tbl[[#This Row],[a'[@b]]", none).is_none());
    assert!(this_row_shorthand("'Sheet [@x]'!A1", none).is_none());
    assert!(this_row_shorthand("\"[@x]", none).is_none());
    let found = this_row_shorthand("SUM('My Sheet'!Täble[@[a']b]])+1", none).expect("shorthand");
    assert_eq!(found.reference, "Täble[@[a']b]]");
    assert_eq!(found.table, "Täble");
    assert_eq!(found.item, "[a']b]");
    // A long run of string literals is one linear pass.
    let long = format!("{}Tbl[[#This Row],[Qty]]", "\"[@x]\"&".repeat(20_000));
    assert!(this_row_shorthand(&long, none).is_none());
    // A `[@...]` naming a column of its table (escapes removed, any case) is a
    // column reference; the scan goes on past it.
    let columns = [
        crate::ooxml::tables::TableColumn {
            name: "@Home".into(),
            ..Default::default()
        },
        crate::ooxml::tables::TableColumn {
            name: "@a]b".into(),
            ..Default::default()
        },
    ];
    let at_t = |table: &str| -> &[crate::ooxml::tables::TableColumn] {
        if table == "AtT" { &columns } else { &[] }
    };
    assert!(this_row_shorthand("SUM(AtT[@home])+AtT[@a']b]", at_t).is_none());
    let found = this_row_shorthand("AtT[@Home]+Other[@Home]", at_t).expect("shorthand");
    assert_eq!(found.reference, "Other[@Home]");
    assert_eq!(found.table, "Other");
    assert_eq!(found.item, "Home");
}

/// Table `AtT` on F1:G3 with the columns `@Home` and `Share`, whose `Share`
/// column has the calculated and totals formulas Excel 365 saved for it.
fn at_home_table() -> crate::ooxml::tables::TableDefinition {
    let mut table = named_table(2, "AtT", Some("AtT"), "F1:G3");
    table.columns = vec![
        crate::ooxml::tables::TableColumn {
            id: 1,
            name: "@Home".into(),
            ..Default::default()
        },
        crate::ooxml::tables::TableColumn {
            id: 2,
            name: "Share".into(),
            calculated_column_formula: Some("AtT[[#This Row],[@Home]]/SUM(AtT[@Home])".into()),
            totals_row_function: Some("custom".into()),
            totals_row_formula: Some("SUM(AtT[@Home])*2".into()),
            ..Default::default()
        },
    ];
    table
}

#[test]
fn test_validate_at_column_reference_excel_saved() {
    // Excel writes a whole-column reference to the column `@Home` as
    // `AtT[@Home]`, and opens the file it saved.
    let mut wb = minimal_workbook();
    wb.sheets[0].worksheet.tables = vec![at_home_table()];
    let report = validate_workbook(&wb);
    assert!(report.is_valid, "{:?}", report.issues);
    assert!(table_errors(&report).is_empty(), "{:?}", report.issues);

    // Without a table name, in another case, and from another table.
    let mut wb = minimal_workbook();
    let mut at_t = at_home_table();
    at_t.columns[1].calculated_column_formula = Some("[@home]*2".into());
    let calc = calc_table(Some("CalcT[[#This Row],[Qty]]*SUM(att[@HOME])"), None);
    wb.sheets[0].worksheet.tables = vec![calc, at_t];
    let report = validate_workbook(&wb);
    assert!(table_errors(&report).is_empty(), "{:?}", report.issues);
}

#[test]
fn test_validate_at_reference_without_such_column_refused() {
    // `[@Home]` in a table with no `@Home` column is still the shorthand.
    for (formula, reference, file_form) in [
        ("[@Home]*2", "[@Home]", "CalcT[[#This Row],[Home]]"),
        (
            "CalcT[@Home]*2",
            "CalcT[@Home]",
            "CalcT[[#This Row],[Home]]",
        ),
        ("Nope[@Home]*2", "Nope[@Home]", "Nope[[#This Row],[Home]]"),
    ] {
        let mut wb = minimal_workbook();
        wb.sheets[0].worksheet.tables = vec![calc_table(Some(formula), None), at_home_table()];
        let report = validate_workbook(&wb);
        let errors = table_errors(&report);
        assert_eq!(errors.len(), 1, "{formula}: {:?}", report.issues);
        assert_eq!(
            errors[0].message,
            format!(
                "Invalid calculatedColumnFormula of column 'Total' in table 'CalcT': Excel refuses the @ shorthand {reference} in a table definition; write {file_form} instead"
            )
        );
    }
}
