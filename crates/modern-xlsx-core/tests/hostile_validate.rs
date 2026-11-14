//! Hostile models for `validate_workbook`, built in memory: each must
//! validate within a time bound ([`LIMIT`]) and a peak-heap bound, and its
//! report must store at most 10,000 issues and 16 MiB of `message` and
//! `location` text while its counts stay exact (the issues left out are
//! counted in `omittedIssueCount`).
//!
//! The cases are the hostile files measured before the report was bounded,
//! at sizes a debug test run holds: many totals tables over many rows (the totals row absent, and the
//! sheet's last row), totals rows of 16,384 cells in reverse column order,
//! a table display name of 100,000 characters repeated by 16,384 totals-cell
//! issues, a sheet name of 100,000 characters repeated by 10,000 cell issues,
//! 150,000 threaded comment issues, and `[@...]` references in the table
//! formulas of 16,384-column tables (the check must stay linear in the
//! references plus the columns).
//!
//! Every case runs sequentially inside one `#[test]`, under a counting
//! `#[global_allocator]` (`peak_alloc::PeakAlloc`), so a timing or peak
//! measures only its own work. Models are built before the measurement. The
//! other tests here pin which `[@...]` references are the `@` this-row
//! shorthand, on models of a few columns (the shared case table
//! `fixtures/this-row-shorthand-cases.json`, which the TypeScript tests run
//! too, and the formulas of a table part Excel saved).

use std::time::{Duration, Instant};

use modern_xlsx_core::ooxml::cell::CellRef;
use modern_xlsx_core::ooxml::tables::{TableColumn, TableDefinition};
use modern_xlsx_core::ooxml::theme::ThemeColors;
use modern_xlsx_core::ooxml::threaded_comments::ThreadedCommentData;
use modern_xlsx_core::ooxml::worksheet::{Cell, CellType, Row, WorksheetXml};
use modern_xlsx_core::validate::{Severity, ValidationReport, validate_workbook};
use modern_xlsx_core::{SheetData, WorkbookData};
use peak_alloc::PeakAlloc;

#[global_allocator]
static PEAK: PeakAlloc = PeakAlloc;

const MIB: usize = 1 << 20;
const LIMIT: Duration = Duration::from_secs(1);

/// Most issues a report stores.
const MAX_ISSUES: usize = 10_000;
/// Most bytes of `message` and `location` text a report stores.
const MAX_TEXT: usize = 16 * MIB;

/// Validates `wb` and asserts it took less than [`LIMIT`] and grew the heap
/// by less than `peak_limit` bytes.
fn validated(label: &str, wb: &WorkbookData, peak_limit: usize) -> ValidationReport {
    PEAK.reset_peak_usage();
    let base = PEAK.current_usage();
    let start = Instant::now();
    let report = validate_workbook(wb);
    let elapsed = start.elapsed();
    let peak = PEAK.peak_usage().saturating_sub(base);
    assert!(
        elapsed < LIMIT,
        "{label}: took {elapsed:?}, limit {LIMIT:?}"
    );
    assert!(
        peak < peak_limit,
        "{label}: peak heap growth {peak} bytes, limit {peak_limit} bytes"
    );
    report
}

/// The `omittedIssueCount` the report serialises; 0 when the field is absent.
fn omitted(report: &ValidationReport) -> usize {
    serde_json::to_value(report)
        .expect("the report serialises")
        .get("omittedIssueCount")
        .map_or(0, |v| {
            usize::try_from(v.as_u64().expect("a count")).expect("fits")
        })
}

/// Asserts the report stores at most [`MAX_ISSUES`] issues and [`MAX_TEXT`]
/// bytes of text, and that its counts add up to `errors` issues in all, of
/// which those not stored are counted as omitted.
fn assert_bounded(label: &str, report: &ValidationReport, errors: usize) {
    let text: usize = report
        .issues
        .iter()
        .map(|i| i.message.len() + i.location.len())
        .sum();
    assert!(report.issues.len() <= MAX_ISSUES, "{label}");
    assert!(text <= MAX_TEXT, "{label}: {text} bytes of text");
    assert_eq!(
        (report.error_count, report.warning_count, report.info_count),
        (errors, 0, 0),
        "{label}"
    );
    assert_eq!(report.issues.len() + omitted(report), errors, "{label}");
    assert_eq!(report.is_valid, errors == 0, "{label}");
}

/// A workbook of one sheet `name` over `worksheet`, with a theme (so that
/// the report has no info issue).
fn workbook(name: String, worksheet: WorksheetXml) -> WorkbookData {
    WorkbookData {
        sheets: vec![SheetData {
            name,
            worksheet,
            ..Default::default()
        }],
        theme_colors: Some(ThemeColors::default()),
        ..Default::default()
    }
}

fn number_cell(reference: String) -> Cell {
    Cell {
        reference,
        cell_type: CellType::Number,
        value: Some("1".into()),
        ..Default::default()
    }
}

/// A table `display_name` over `ref_range` with a totals row and `columns`
/// columns of `totalsRowFunction` `function`.
fn totals_table(
    id: u32,
    display_name: String,
    ref_range: &str,
    columns: u32,
    function: Option<&str>,
) -> TableDefinition {
    TableDefinition {
        id,
        name: Some(display_name.clone()),
        display_name,
        ref_range: ref_range.into(),
        header_row_count: 1,
        totals_row_count: 1,
        totals_row_shown: true,
        columns: (0..columns)
            .map(|c| TableColumn {
                id: c + 1,
                name: format!("C{c}"),
                totals_row_function: function.map(Into::into),
                ..Default::default()
            })
            .collect(),
        style_info: None,
        auto_filter_ref: None,
    }
}

/// Row 2 holding a number in each of the 16,384 columns, in `order`.
fn wide_row(order: impl Iterator<Item = u32>) -> Row {
    Row {
        index: 2,
        cells: order
            .map(|col| number_cell(CellRef { col, row: 1 }.to_a1()))
            .collect(),
        ..Default::default()
    }
}

/// Evidence t06/t06b: 10,000 `sum` tables down to row 1,048,575 over a sheet
/// of 200,000 rows, without and with that last row.
fn many_totals_tables_over_many_rows_case() {
    for last_row in [false, true] {
        let label = format!("10,000 totals tables x 200,000 rows (last row {last_row})");
        let mut rows: Vec<Row> = (1..=200_000)
            .map(|r| Row {
                index: r,
                cells: vec![number_cell(format!("B{r}"))],
                ..Default::default()
            })
            .collect();
        if last_row {
            rows.push(Row {
                index: 1_048_575,
                cells: vec![number_cell("B1048575".into())],
                ..Default::default()
            });
        }
        let worksheet = WorksheetXml {
            tables: (0..10_000)
                .map(|i| totals_table(i + 1, format!("T{i}x"), "A1:A1048575", 1, Some("sum")))
                .collect(),
            rows,
            ..Default::default()
        };
        let wb = workbook("S".into(), worksheet);
        let report = validated(&label, &wb, 16 * MIB);
        assert_bounded(&label, &report, 0);
    }
}

/// Evidence t07/t07b: 20 tables of 16,384 `sum` columns over one totals row
/// of 16,384 numbers in reverse column order.
fn wide_totals_rows_in_reverse_order_case() {
    let label = "20 tables x 16,384 columns, totals row in reverse order";
    let worksheet = WorksheetXml {
        tables: (0..20)
            .map(|i| totals_table(i + 1, format!("T{i}x"), "A1:XFD2", 16_384, Some("sum")))
            .collect(),
        rows: vec![wide_row((0..16_384).rev())],
        ..Default::default()
    };
    let wb = workbook("S".into(), worksheet);
    let report = validated(label, &wb, 16 * MIB);
    assert_bounded(label, &report, 0);
}

/// Evidence t08 (table name): one table whose display name is 100,000
/// characters, over a totals row of 16,384 numbers in columns without a
/// function. Each totals-cell issue repeats the name.
fn long_table_name_over_a_wide_totals_row_case() {
    let label = "table name of 100,000 characters x 16,384 totals cells";
    let name = format!("T{}", "x".repeat(100_000));
    let worksheet = WorksheetXml {
        tables: vec![totals_table(1, name, "A1:XFD2", 16_384, None)],
        rows: vec![wide_row(0..16_384)],
        ..Default::default()
    };
    let wb = workbook("S".into(), worksheet);
    let report = validated(label, &wb, 48 * MIB);
    // The display name and the name attribute are too long, and each of the
    // 16,384 totals cells holds a number in a column without a function.
    assert_bounded(label, &report, 2 + 16_384);
    assert!(
        report.issues.len() < 1_000,
        "{label}: stopped at the text bound"
    );
}

/// Evidence t08 (sheet name): a sheet name of 100,000 characters over
/// 10,000 cells with a dangling style index. Each cell issue repeats the name.
fn long_sheet_name_over_many_bad_styles_case() {
    let label = "sheet name of 100,000 characters x 10,000 dangling styles";
    let worksheet = WorksheetXml {
        rows: (1..=10_000)
            .map(|r| Row {
                index: r,
                cells: vec![Cell {
                    style_index: Some(999),
                    ..number_cell(format!("A{r}"))
                }],
                ..Default::default()
            })
            .collect(),
        ..Default::default()
    };
    let wb = workbook("x".repeat(100_000), worksheet);
    let report = validated(label, &wb, 48 * MIB);
    // The sheet name and the 10,000 style indices.
    assert_bounded(label, &report, 1 + 10_000);
    assert!(
        report.issues.len() < 1_000,
        "{label}: stopped at the text bound"
    );
}

/// The 2,000,000-thread-root file, scaled down: 100,000 thread roots without a
/// personId, two on each of 50,000 cells: 150,000 errors.
fn many_thread_issues_case() {
    let label = "100,000 thread roots, 50,000 second roots";
    let worksheet = WorksheetXml {
        threaded_comments: (0..100_000u32)
            .map(|i| ThreadedCommentData {
                id: format!("{{{i:08X}-0000-4000-8000-000000000000}}"),
                ref_cell: format!("A{}", i / 2 + 1),
                person_id: String::new(),
                text: String::new(),
                timestamp: String::new(),
                parent_id: None,
                done: None,
                mentions: Vec::new(),
            })
            .collect(),
        ..Default::default()
    };
    let wb = workbook("S".into(), worksheet);
    let report = validated(label, &wb, 48 * MIB);
    assert_bounded(label, &report, 150_000);
    assert_eq!(report.issues.len(), MAX_ISSUES, "{label}");
    assert!(
        report.issues.iter().all(|i| i.severity == Severity::Error),
        "{label}"
    );
}

/// A table `display_name` over `ref_range` without a totals row, with a
/// column named by `name` for each of `columns` columns, and the
/// `calculatedColumnFormula` `formula(c)` on column `c`.
fn formula_table(
    id: u32,
    display_name: String,
    ref_range: &str,
    columns: u32,
    name: impl Fn(u32) -> String,
    formula: impl Fn(u32) -> Option<String>,
) -> TableDefinition {
    TableDefinition {
        id,
        name: Some(display_name.clone()),
        display_name,
        ref_range: ref_range.into(),
        header_row_count: 1,
        totals_row_count: 0,
        totals_row_shown: false,
        columns: (0..columns)
            .map(|c| TableColumn {
                id: c + 1,
                name: name(c),
                calculated_column_formula: formula(c),
                ..Default::default()
            })
            .collect(),
        style_info: None,
        auto_filter_ref: None,
    }
}

/// Four tables of 16,384
/// columns with `[@x]` as the `calculatedColumnFormula` of every column, once
/// without a column `@x` (each formula is the shorthand: 65,536 errors) and
/// once with a last column named `@x` (each is a column reference: none).
fn this_row_shorthand_in_every_column_case() {
    for at_x in [false, true] {
        let label = format!("4 tables x 16,384 columns of [@x] (column @x {at_x})");
        let name = |c: u32| {
            if at_x && c == 16_383 {
                "@x".to_owned()
            } else {
                format!("C{c}")
            }
        };
        let worksheet = WorksheetXml {
            tables: (0..4)
                .map(|i| {
                    formula_table(i + 1, format!("W{i}x"), "A1:XFD2", 16_384, name, |_| {
                        Some("[@x]".into())
                    })
                })
                .collect(),
            ..Default::default()
        };
        let wb = workbook("S".into(), worksheet);
        let report = validated(&label, &wb, 48 * MIB);
        assert_bounded(&label, &report, if at_x { 0 } else { 4 * 16_384 });
    }
}

/// One table of
/// 16,384 columns, the last sixteen named `@a` to `@p`, whose first twenty
/// columns hold formulas of up to 8,190 bytes of `[@a]+[@b]+...+[@p]+[@a]...`
/// (Excel's formula limit is 8,192 characters). Every reference is a column
/// reference.
fn long_formulas_of_at_columns_case() {
    let label = "16,384 columns, 20 formulas of up to 8,190 bytes of [@a]..[@p]";
    let mut formula = String::new();
    for letter in (b'a'..=b'p').cycle() {
        if formula.len() + 5 > 8_190 {
            break;
        }
        formula.push_str("[@");
        formula.push(char::from(letter));
        formula.push_str("]+");
    }
    formula.pop();
    assert!(formula.len() <= 8_190, "{label}: {} bytes", formula.len());
    let name = |c: u32| match c.checked_sub(16_384 - 16) {
        Some(at) => format!("@{}", char::from(b'a' + u8::try_from(at).expect("fits"))),
        None => format!("C{c}"),
    };
    let worksheet = WorksheetXml {
        tables: vec![formula_table(
            1,
            "V5f".into(),
            "A1:XFD2",
            16_384,
            name,
            |c| (c < 20).then(|| formula.clone()),
        )],
        ..Default::default()
    };
    let wb = workbook("S".into(), worksheet);
    let report = validated(label, &wb, 48 * MIB);
    assert_bounded(label, &report, 0);
}

#[test]
fn hostile_models_validate_within_bounds() {
    many_totals_tables_over_many_rows_case();
    wide_totals_rows_in_reverse_order_case();
    long_table_name_over_a_wide_totals_row_case();
    long_sheet_name_over_many_bad_styles_case();
    many_thread_issues_case();
    this_row_shorthand_in_every_column_case();
    long_formulas_of_at_columns_case();
}

/// The issues of `wb`'s report that name the `@` this-row shorthand.
fn shorthand_issues(wb: &WorkbookData) -> Vec<String> {
    validate_workbook(wb)
        .issues
        .into_iter()
        .filter(|issue| issue.message.contains("@ shorthand"))
        .map(|issue| issue.message)
        .collect()
}

/// The shared case table `fixtures/this-row-shorthand-cases.json`, which the
/// TypeScript `addTable` and `validate()` tests run too: table `T` with the
/// case's columns and a last column `Out` holding the formula, beside one
/// table per entry of `tables`.
#[test]
fn this_row_shorthand_shared_cases() {
    let fixture: serde_json::Value =
        serde_json::from_str(include_str!("fixtures/this-row-shorthand-cases.json"))
            .expect("the fixture parses");
    let cases = fixture["cases"].as_array().expect("cases");
    assert!(cases.len() >= 30, "{} cases", cases.len());
    let names = |value: &serde_json::Value| -> Vec<String> {
        value
            .as_array()
            .expect("column names")
            .iter()
            .map(|name| name.as_str().expect("a column name").to_owned())
            .collect()
    };
    let table = |id: u32, display_name: &str, row: u32, columns: Vec<String>| {
        let count = u32::try_from(columns.len()).expect("fits");
        let end = CellRef {
            col: count - 1,
            row: row + 1,
        };
        formula_table(
            id,
            display_name.into(),
            &format!("A{}:{}", row + 1, end.to_a1()),
            count,
            |c| columns[c as usize].clone(),
            |_| None,
        )
    };
    let mut wrong = Vec::new();
    for case in cases {
        let formula = case["formula"].as_str().expect("formula");
        let mut columns = names(&case["columns"]);
        columns.push("Out".into());
        let mut own = table(1, "T", 0, columns);
        own.columns
            .last_mut()
            .expect("Out")
            .calculated_column_formula = Some(formula.into());
        let mut tables = vec![own];
        for (i, (name, columns)) in case["tables"]
            .as_object()
            .expect("tables")
            .iter()
            .enumerate()
        {
            let id = u32::try_from(i).expect("fits") + 2;
            tables.push(table(id, name, 10 * (id - 1), names(columns)));
        }
        let wb = workbook(
            "S".into(),
            WorksheetXml {
                tables,
                ..Default::default()
            },
        );
        let issues = shorthand_issues(&wb);
        if issues.len() != usize::from(case["reported"].as_bool().expect("reported")) {
            wrong.push(format!("{case}: {issues:?}"));
        }
    }
    assert!(
        wrong.is_empty(),
        "{} cases disagree:\n{}",
        wrong.len(),
        wrong.join("\n")
    );
}

/// The table part Excel 16.0 (build 20527) saved for a table `LocT` beside a
/// link to another workbook:
/// `SUM([1]!Ext[@Home])` names the column `@Home` of table `Ext` in
/// another workbook, and `LocT[@a_x005F_x000a_b]` the column whose name is the
/// text `@a_x000a_b`, which the table part escapes in both places. Excel
/// opens the file; no formula is the shorthand.
#[test]
fn excel_saved_cross_workbook_and_escaped_references_pass() {
    let ccf = |c: u32| match c {
        4 => Some("SUM([1]!Ext[@Home])".to_owned()),
        5 => Some("LocT[[#This Row],[Qty]]+SUM([1]!Ext[Qty])".into()),
        6 => Some("SUM(LocT[@a_x005F_x000a_b])".into()),
        7 => Some("SUM([1]!Ext[[#All],[@Home]])".into()),
        8 => Some("SUM(LocT[@t\tb])".into()),
        _ => None,
    };
    let names = [
        "Item",
        "Qty",
        "@a_x005f_x000a_b",
        "@t_x0009_b",
        "E1",
        "E2",
        "E3",
        "E4",
        "E5",
    ];
    let mut table = formula_table(
        1,
        "LocT".into(),
        "A1:I5",
        9,
        |c| names[c as usize].into(),
        ccf,
    );
    table.totals_row_count = 1;
    table.columns[1].totals_row_function = Some("custom".into());
    table.columns[1].totals_row_formula = Some("SUM([1]!Ext[@Home])".into());
    let wb = workbook(
        "Loc".into(),
        WorksheetXml {
            tables: vec![table],
            ..Default::default()
        },
    );
    assert_eq!(shorthand_issues(&wb), Vec::<String>::new());
}
