//! Workbook validation and repair.
//!
//! - **Validation:** [`validate_workbook`] reports issues without changing the
//!   workbook, each with a severity, a location and a suggestion, in a
//!   [`ValidationReport`]. It checks sheet names (no sheets, empty, longer
//!   than 31 characters, forbidden or control characters, a leading or
//!   trailing apostrophe, the reserved name `History`, duplicates); styles (dangling
//!   cellXf font, fill and border ids, a custom `numFmtId` without its
//!   record, fewer than two fills, no font, no cellXf); conditional-format and
//!   table `dxfId`s past the dxf list; table names Excel refuses to open (a
//!   display name that is empty, longer than 255 UTF-16 code units, holds a
//!   space or ASCII punctuation other than `_ . \ ?`, starts with a digit,
//!   `.` or `?`, is `TRUE`, `FALSE`, a cell reference or an R1C1 reference
//!   or starts with one; a `name` longer than 255; a display name another
//!   table or a defined name has, or a `name` another table's `name` has,
//!   ignoring case); table totals-row cells Excel refuses to open (a formula
//!   other than the column's `SUBTOTAL` in a function column, no formula in a
//!   `custom` column with a `totalsRowFormula`, and a formula, a number, a
//!   boolean or text other than the `totalsRowLabel` in a column without a
//!   function); a `calculatedColumnFormula` or `totalsRowFormula` holding
//!   the `@` this-row shorthand (`[@Qty]`, `Sales[@Qty]`, `[@]`), which Excel
//!   refuses to open in a table definition; cells (unparseable references,
//!   dangling style indices, shared-string cells without a value); merge
//!   ranges (unparseable, reversed, overlapping); defined names with an empty
//!   name or value, a name Excel refuses to open (other than an `_xlnm.`
//!   built-in name: the table-name rules, except that `:` and an A1 cell
//!   reference are allowed), and a `_xlnm.Print_Area` or `_xlnm.Print_Titles` value
//!   (the name in any letter case, as Excel matches it) with an area that has
//!   no sheet qualifier (Excel refuses the file); data
//!   validations with an empty `sqref` or a list without `formula1`; a missing theme; document properties (`created` and
//!   `modified` not W3CDTF, the ISO 8601 profile `dcterms:created` uses,
//!   and whitespace-only text fields); row order (row indices not strictly
//!   ascending); threaded comment `id`, `personId` and `parentId` values and
//!   person `id`s that are not GUIDs in braces (Excel refuses the file, or
//!   drops the reply), a reply whose `parentId` names another reply on its
//!   sheet (Excel refuses the file), a second thread root on a cell of its
//!   sheet (Excel shows only the first thread), and a threaded comment text
//!   longer than 32,767 UTF-16 code units (Excel refuses the file); a pivot table whose `cacheId` names
//!   no pivot cache (a warning; the write fails), and the first pivot table
//!   the writer would add too many items or pivot fields to from its cache
//!   (a warning; the write fails with `SIZE_LIMIT`); and four values the writer cannot write, which also fail
//!   the write (`check_writable`) and which repair leaves alone: a sheet
//!   protection `password` that is not a legacy hash (1 to 4 hex digits) and
//!   would be written, a conditional-format rule with more than three
//!   formulas, a colour (styles, tab, conditional format, sparkline or
//!   rich-text run) whose written `rgb` is not 6 or 8 hex digits, and a
//!   chart written from the model two of whose series share a `c:order`
//!   (within one chart-type element or across its primary and secondary
//!   chart).
//! - **Repair:** [`repair_workbook`] fixes in place: invalid and duplicate
//!   sheet names; missing default styles (a Calibri 11 font, the `none` and
//!   `gray125` fills, a border, a cellXf) and dangling cellXf font, fill and
//!   border ids (set to 0); dangling cell style indices (set to 0);
//!   shared-string cells without a value (set to the empty string); a missing
//!   theme (set to `ThemeColors::default()`); document property dates
//!   (normalised, given a `Z`, or cleared) and whitespace-only text fields
//!   (cleared); row order (rows sorted, rows sharing an index merged); and
//!   reversed merge ranges (normalised).

use core::hint::cold_path;
use std::borrow::Cow;
use std::cell::OnceCell;
use std::collections::hash_map::Entry;
use std::collections::{HashMap, HashSet};

use serde::{Deserialize, Serialize};

use crate::ooxml::cell::{CellRef, MAX_COL, MAX_ROW};
use crate::ooxml::charts::{ChartData, WorksheetChart};
use crate::ooxml::color::{Color, ColorOut, is_argb, not_argb_message, resolve_color};
use crate::ooxml::styles::{Border, CellXf, Fill, Font};
use crate::ooxml::theme::ThemeColors;
use crate::ooxml::threaded_comments;
use crate::ooxml::worksheet::{
    Cell, CellType, ConditionalFormatting, ConditionalFormattingRule, MAX_CF_RULE_FORMULAS, Row,
    SheetProtection, cf_formula_count_message, non_hash_password_message,
};
use crate::{SheetData, WorkbookData};

// ---------------------------------------------------------------------------
// OOXML limits
// ---------------------------------------------------------------------------

/// Maximum sheet name length Excel accepts.
const MAX_SHEET_NAME_LEN: usize = 31;
/// First `numFmtId` that is a custom format (ids below it are built in).
const FIRST_CUSTOM_NUM_FMT_ID: u32 = 164;
/// Characters forbidden in sheet names.
const FORBIDDEN_SHEET_CHARS: &[char] = &['\\', '/', '*', '?', ':', '[', ']'];

// ---------------------------------------------------------------------------
// Severity & Issue types
// ---------------------------------------------------------------------------

/// Severity level for a validation issue.
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum Severity {
    /// Informational finding — no action required.
    Info,
    /// A problem a reader can usually work around (non-W3CDTF dates, empty defined-name values or sqref, a list validation without formula1, unsorted rows, reversed merges, fewer than two fills, a pivot table whose `cacheId` names no cache or that the writer would add too many items or pivot fields to).
    Warning,
    /// A broken name, reference or rule (no sheets, invalid or duplicate sheet name, no fonts or no cellXfs, dangling style, dxf or numFmt id, bad cell reference, bad or overlapping merge range, a shared-string cell without a value, empty defined name, a defined name Excel refuses (an R1C1 reference, `TRUE`, `FALSE`, a refused character, a leading digit, `.` or `?`, or more than 255 characters), a `_xlnm.Print_Area` or `_xlnm.Print_Titles` (any letter case) area without a sheet qualifier, a sheet protection password that is not a stored hash, a conditional-format rule with more than three formulas, a colour that is not 6 or 8 hex digits, chart series that share a `c:order`, a threaded comment or person id that is not a GUID in braces, a threaded comment reply whose `parentId` names another reply, a second threaded comment root on a cell, a threaded comment text longer than 32,767 UTF-16 code units, a table name Excel refuses or another table or a defined name already has, a table totals-row cell that disagrees with its column, a `calculatedColumnFormula` or `totalsRowFormula` holding the `@` this-row shorthand).
    Error,
}

/// Category of a validation issue.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum IssueCategory {
    CellReference,
    StyleIndex,
    MergeCell,
    SharedString,
    SheetName,
    DefinedName,
    DataValidation,
    ConditionalFormatting,
    Theme,
    Metadata,
    Structure,
    SheetProtection,
}

/// A single validation issue with an actionable suggestion.
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ValidationIssue {
    pub severity: Severity,
    pub category: IssueCategory,
    /// Human-readable description of the issue.
    pub message: String,
    /// Where the issue was found (e.g. `Sheet1!A1`, `styles.cellXfs[3]`).
    pub location: String,
    /// Suggested fix (e.g. "Clamp style index to 0").
    pub suggestion: String,
    /// Whether `repair_workbook` can auto-fix this issue.
    pub auto_fixable: bool,
}

/// Most issues a [`ValidationReport`] stores in `issues`.
const MAX_REPORTED_ISSUES: usize = 10_000;
/// Most bytes of `message` and `location` text the issues a
/// [`ValidationReport`] stores hold together (16 MiB).
const MAX_REPORT_TEXT_BYTES: usize = 16 << 20;

/// Result of validating a workbook.
///
/// The report is bounded: `issues` holds the first issues found, at most
/// 10,000 of them and at most 16 MiB of `message` and `location` text (a
/// name a file controls, such as a sheet or table name, can be repeated in
/// every issue). Each issue past either bound is counted, not stored:
/// `error_count`, `warning_count` and `info_count` count every issue found,
/// `is_valid` comes from `error_count`, and `omitted_issue_count` is the
/// number left out, so `issues.len() + omitted_issue_count` is the sum of
/// the three counts.
#[derive(Debug, Clone, Default, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ValidationReport {
    pub issues: Vec<ValidationIssue>,
    pub error_count: usize,
    pub warning_count: usize,
    pub info_count: usize,
    /// `true` if no errors were found.
    pub is_valid: bool,
    /// The issues found but not stored in `issues` because the report
    /// reached its bound; 0, and absent from the JSON, when every issue is
    /// stored.
    #[serde(default, skip_serializing_if = "is_zero")]
    pub omitted_issue_count: usize,
    /// Bytes of `message` and `location` text in `issues`.
    #[serde(skip)]
    text_bytes: usize,
}

/// `true` for 0 (serde's `skip_serializing_if` passes a reference).
const fn is_zero(n: &usize) -> bool {
    *n == 0
}

impl ValidationReport {
    /// Count `issue` and store it while the report holds fewer than
    /// [`MAX_REPORTED_ISSUES`] issues, its text fits within
    /// [`MAX_REPORT_TEXT_BYTES`] and no earlier issue was left out; otherwise
    /// count it as omitted. Every check reports through here.
    #[inline]
    fn push(&mut self, issue: ValidationIssue) {
        self.count(issue.severity);
        let text = issue.message.len() + issue.location.len();
        if self.omitted_issue_count == 0
            && self.issues.len() < MAX_REPORTED_ISSUES
            && self.text_bytes + text <= MAX_REPORT_TEXT_BYTES
        {
            self.text_bytes += text;
            self.issues.push(issue);
        } else {
            self.omitted_issue_count += 1;
        }
    }

    /// When the report stores no more issues, count an issue of `severity`
    /// as omitted and return `true`, so that the caller skips building it:
    /// a check whose message or location repeats a name the file controls
    /// calls this before formatting them.
    #[inline]
    fn omit_if_full(&mut self, severity: Severity) -> bool {
        let full = self.omitted_issue_count > 0
            || self.issues.len() >= MAX_REPORTED_ISSUES
            || self.text_bytes >= MAX_REPORT_TEXT_BYTES;
        if full {
            self.count(severity);
            self.omitted_issue_count += 1;
        }
        full
    }

    #[inline]
    const fn count(&mut self, severity: Severity) {
        match severity {
            Severity::Error => self.error_count += 1,
            Severity::Warning => self.warning_count += 1,
            Severity::Info => self.info_count += 1,
        }
    }

    fn finalize(&mut self) {
        self.is_valid = self.error_count == 0;
        self.issues
            .sort_by_key(|issue| std::cmp::Reverse(issue.severity));
    }
}

// ---------------------------------------------------------------------------
// Validate
// ---------------------------------------------------------------------------

/// Validate a workbook and return a report of the issues found: every issue
/// is counted, and the first ones are stored up to the report's bound (see
/// [`ValidationReport`]).
///
/// This does **not** modify the workbook. Use [`repair_workbook`] to auto-fix
/// repairable issues.
pub fn validate_workbook(wb: &WorkbookData) -> ValidationReport {
    let mut report = ValidationReport::default();

    validate_sheets(&mut report, wb);
    validate_styles(&mut report, wb);
    validate_dxf_references(&mut report, wb);
    validate_table_names(&mut report, wb);
    validate_table_totals(&mut report, wb);
    validate_table_formulas(&mut report, wb);
    validate_cells(&mut report, wb);
    validate_merge_cells(&mut report, wb);
    validate_defined_names(&mut report, wb);
    validate_data_validations(&mut report, wb);
    validate_theme(&mut report, wb);
    validate_metadata(&mut report, wb);
    validate_structure(&mut report, wb);
    validate_thread_ids(&mut report, wb);
    validate_sheet_protection(&mut report, wb);
    validate_cf_formula_counts(&mut report, wb);
    validate_colors(&mut report, wb);
    validate_chart_series_orders(&mut report, wb);
    validate_pivot_cache_ids(&mut report, wb);
    validate_pivot_additions(&mut report, wb);

    report.finalize();
    report
}

// -- Sheet names --

/// The reserved sheet name Excel keeps for its change-tracking history.
const RESERVED_SHEET_NAME: &str = "History";

/// First rule a single sheet name breaks, as a short reason; `None` when the
/// name is acceptable. Checks the name rules only (empty, longer than 31
/// characters, forbidden or control characters, leading or trailing
/// apostrophe, the reserved name "History", compared case-insensitively);
/// duplicate detection needs the whole workbook and stays with the caller.
/// One predicate serves validation, repair and the streaming writer, so a
/// repaired name always validates.
pub(crate) fn sheet_name_problem(name: &str) -> Option<&'static str> {
    if name.is_empty() {
        return Some("Sheet name is empty");
    }
    if name.chars().count() > MAX_SHEET_NAME_LEN {
        return Some("Sheet name exceeds 31 character limit");
    }
    if name
        .chars()
        .any(|c| FORBIDDEN_SHEET_CHARS.contains(&c) || c.is_control())
    {
        return Some("Sheet name contains forbidden characters (\\ / * ? : [ ] or controls)");
    }
    if name.starts_with('\'') || name.ends_with('\'') {
        return Some("Sheet name must not start or end with an apostrophe");
    }
    if name.eq_ignore_ascii_case(RESERVED_SHEET_NAME) {
        return Some("Sheet name 'History' is reserved");
    }
    None
}

fn validate_sheets(report: &mut ValidationReport, wb: &WorkbookData) {
    if wb.sheets.is_empty() {
        report.push(ValidationIssue {
            severity: Severity::Error,
            category: IssueCategory::Structure,
            message: "Workbook contains no sheets".into(),
            location: "workbook".into(),
            suggestion: "Add at least one worksheet".into(),
            auto_fixable: false,
        });
        return;
    }

    let mut seen_names: HashSet<String> = HashSet::with_capacity(wb.sheets.len());

    for (i, sheet) in wb.sheets.iter().enumerate() {
        if let Some(problem) = sheet_name_problem(&sheet.name) {
            report.push(ValidationIssue {
                severity: Severity::Error,
                category: IssueCategory::SheetName,
                message: format!("Invalid sheet name '{}': {problem}", sheet.name),
                location: format!("sheets[{i}]"),
                suggestion: format!("Rename the sheet (repair falls back to 'Sheet{}')", i + 1),
                auto_fixable: true,
            });
        }

        let lower = sheet.name.to_lowercase();
        if !seen_names.insert(lower) {
            report.push(ValidationIssue {
                severity: Severity::Error,
                category: IssueCategory::SheetName,
                message: format!("Duplicate sheet name '{}'", sheet.name),
                location: format!("sheets[{i}]"),
                suggestion: "Rename to a unique name".into(),
                auto_fixable: true,
            });
        }
    }
}

// -- Styles --

fn validate_styles(report: &mut ValidationReport, wb: &WorkbookData) {
    let styles = &wb.styles;
    let font_count = styles.fonts.len();
    let fill_count = styles.fills.len();
    let border_count = styles.borders.len();

    for (i, xf) in styles.cell_xfs.iter().enumerate() {
        let loc = || format!("styles.cellXfs[{i}]");

        if xf.font_id as usize >= font_count {
            report.push(ValidationIssue {
                severity: Severity::Error,
                category: IssueCategory::StyleIndex,
                message: format!(
                    "cellXf[{i}].fontId={} exceeds fonts count ({font_count})",
                    xf.font_id
                ),
                location: loc(),
                suggestion: "Clamp fontId to 0".into(),
                auto_fixable: true,
            });
        }

        if xf.fill_id as usize >= fill_count {
            report.push(ValidationIssue {
                severity: Severity::Error,
                category: IssueCategory::StyleIndex,
                message: format!(
                    "cellXf[{i}].fillId={} exceeds fills count ({fill_count})",
                    xf.fill_id
                ),
                location: loc(),
                suggestion: "Clamp fillId to 0".into(),
                auto_fixable: true,
            });
        }

        if xf.border_id as usize >= border_count {
            report.push(ValidationIssue {
                severity: Severity::Error,
                category: IssueCategory::StyleIndex,
                message: format!(
                    "cellXf[{i}].borderId={} exceeds borders count ({border_count})",
                    xf.border_id
                ),
                location: loc(),
                suggestion: "Clamp borderId to 0".into(),
                auto_fixable: true,
            });
        }

        // Ids below 164 are built-in formats; a custom id needs its numFmt record.
        if xf.num_fmt_id >= FIRST_CUSTOM_NUM_FMT_ID
            && !styles.num_fmts.iter().any(|f| f.id == xf.num_fmt_id)
        {
            report.push(ValidationIssue {
                severity: Severity::Error,
                category: IssueCategory::StyleIndex,
                message: format!(
                    "cellXf[{i}].numFmtId={} has no numFmt definition",
                    xf.num_fmt_id
                ),
                location: loc(),
                suggestion: "Add the missing numFmt or point the xf at a built-in format".into(),
                auto_fixable: false,
            });
        }
    }

    // Verify minimum required fills (none + gray125).
    if fill_count < 2 {
        report.push(ValidationIssue {
            severity: Severity::Warning,
            category: IssueCategory::StyleIndex,
            message: format!(
                "Styles has only {fill_count} fill(s); Excel requires at least 2 (none + gray125)"
            ),
            location: "styles.fills".into(),
            suggestion: "Inject default 'none' and 'gray125' fills".into(),
            auto_fixable: true,
        });
    }

    // Verify minimum required fonts (at least one).
    if font_count == 0 {
        report.push(ValidationIssue {
            severity: Severity::Error,
            category: IssueCategory::StyleIndex,
            message: "Styles has no fonts; Excel requires at least one default font".into(),
            location: "styles.fonts".into(),
            suggestion: "Inject default Calibri 11pt font".into(),
            auto_fixable: true,
        });
    }

    // Verify at least one cellXf.
    if styles.cell_xfs.is_empty() {
        report.push(ValidationIssue {
            severity: Severity::Error,
            category: IssueCategory::StyleIndex,
            message: "Styles has no cellXfs; Excel requires at least one default xf".into(),
            location: "styles.cellXfs".into(),
            suggestion: "Inject default cellXf".into(),
            auto_fixable: true,
        });
    }
}

// -- Differential formats --

/// Report a conditional-format rule or table column whose `dxfId` points past
/// `styles.dxfs`. Not auto-fixable: there is no right dxf to point at.
fn validate_dxf_references(report: &mut ValidationReport, wb: &WorkbookData) {
    let dxf_count = wb.styles.dxfs.len();
    let mut dangling = |dxf_id: Option<u32>, what: &str, location: &dyn Fn() -> String| {
        if let Some(id) = dxf_id
            && id as usize >= dxf_count
            && !report.omit_if_full(Severity::Error)
        {
            report.push(ValidationIssue {
                severity: Severity::Error,
                category: IssueCategory::StyleIndex,
                message: format!("{what} dxfId={id} exceeds dxfs count ({dxf_count})"),
                location: location(),
                suggestion: "Point the dxfId at an existing dxf or remove it".into(),
                auto_fixable: false,
            });
        }
    };

    for sheet in &wb.sheets {
        let ws = &sheet.worksheet;
        for (i, cf) in ws.conditional_formatting.iter().enumerate() {
            for (j, rule) in cf.rules.iter().enumerate() {
                dangling(rule.dxf_id, "Conditional-format rule", &|| {
                    format!("{}!conditionalFormatting[{i}].rules[{j}]", sheet.name)
                });
            }
        }
        for table in &ws.tables {
            for col in &table.columns {
                for (dxf_id, part) in [
                    (col.header_row_dxf_id, "Table header"),
                    (col.data_dxf_id, "Table data"),
                    (col.totals_row_dxf_id, "Table totals"),
                ] {
                    dangling(dxf_id, part, &|| {
                        format!("{}!table({})", sheet.name, table.display_name)
                    });
                }
            }
        }
    }
}

// -- Tables --

/// Longest table `displayName` or `name` Excel opens, in UTF-16 code units.
const MAX_TABLE_NAME_LEN: usize = 255;

/// The kind of name [`name_problem`] checks. Excel 365 (build 20527,
/// 2026-11-12) opened a file holding each kind by the same rules but two:
/// a defined name may hold `:` after its first character (`a:b`; `ab:` and
/// `a::b` were refused, so `:` is not checked at all), and a defined name
/// that is an A1 cell reference inside the grid (`A1`, `R1`, `C16385`) opens
/// and keeps its name, which also exempts it from the R1C1 rule.
#[derive(Clone, Copy, PartialEq, Eq)]
enum NameKind {
    /// A table display name.
    Table,
    /// A defined name other than an `_xlnm.` built-in name.
    Defined,
}

/// `true` for a byte Excel refuses anywhere in a table name: a space, tab,
/// line break, DEL, or ASCII punctuation other than `_ . \ ?`.
const fn is_forbidden_name_byte(b: u8) -> bool {
    matches!(b, b'\t' | b'\n' | b'\r' | b' ' | 0x7f)
        || (b.is_ascii_punctuation() && !matches!(b, b'_' | b'.' | b'\\' | b'?'))
}

/// First rule Excel applies to a name of `kind` that `name` breaks, as a
/// short reason; `None` when Excel opens it. The rules are the ones an Excel
/// 365 probe proved: 1 to 255 UTF-16 code units; no space, tab, line break,
/// DEL or ASCII punctuation other than `_ . \ ?` (and `:` in a defined name);
/// no leading digit, `.` or `?`; not `TRUE` or `FALSE`; not a cell reference
/// inside the grid (a table name only); and not an R1C1 reference or a name
/// that starts with one ([`is_r1c1_name`]), case ignored. The TypeScript
/// `Worksheet.addTable` applies the table rules.
fn name_problem(name: &str, kind: NameKind) -> Option<&'static str> {
    if name.is_empty() {
        return Some("it is empty");
    }
    if name.encode_utf16().count() > MAX_TABLE_NAME_LEN {
        return Some("it is longer than 255 characters");
    }
    if name
        .bytes()
        .any(|b| is_forbidden_name_byte(b) && !(b == b':' && kind == NameKind::Defined))
    {
        return Some(
            "it holds a character Excel refuses (a space, tab, line break, DEL, or ASCII punctuation other than '_', '.', '\\' and '?')",
        );
    }
    if name.starts_with(|c: char| c.is_ascii_digit() || c == '.' || c == '?') {
        return Some("it starts with a digit, '.' or '?'");
    }
    if name.eq_ignore_ascii_case("TRUE") || name.eq_ignore_ascii_case("FALSE") {
        return Some("TRUE and FALSE are reserved");
    }
    if CellRef::parse(name).is_ok() {
        return match kind {
            NameKind::Table => Some("it is a cell reference"),
            NameKind::Defined => None,
        };
    }
    if is_r1c1_name(name) {
        return Some("it is an R1C1 reference or starts with one");
    }
    None
}

/// `true` for the names Excel reads as R1C1 references, ignoring case: `R`,
/// `C` and `RC` alone, and any name starting `R<row>`, `C<column>` or
/// `RC<column>` (row 1 to 1 048 576, column 1 to 16 384) unless a `.`
/// follows the number (`R1x` and `C1_` are refused, `R1.a` and `R0x` open).
fn is_r1c1_name(name: &str) -> bool {
    let b = name.as_bytes();
    let (start, max) = match b {
        [r, c, ..] if r.eq_ignore_ascii_case(&b'r') && c.eq_ignore_ascii_case(&b'c') => {
            (2, MAX_COL + 1)
        }
        [r, ..] if r.eq_ignore_ascii_case(&b'r') => (1, MAX_ROW + 1),
        [c, ..] if c.eq_ignore_ascii_case(&b'c') => (1, MAX_COL + 1),
        _ => return false,
    };
    let end = start + b[start..].iter().take_while(|d| d.is_ascii_digit()).count();
    if end == start {
        return end == b.len();
    }
    name[start..end]
        .parse::<u32>()
        .is_ok_and(|v| (1..=max).contains(&v))
        && b.get(end) != Some(&b'.')
}

/// Report the table names Excel refuses to open: a display name breaking
/// [`name_problem`], a `name` longer than 255 code units, a display
/// name another table or a defined name (any scope) already has, and a
/// non-empty `name` another table's `name` already has, all ignoring case.
fn validate_table_names(report: &mut ValidationReport, wb: &WorkbookData) {
    let defined: HashMap<String, &str> = wb
        .defined_names
        .iter()
        .map(|d| (d.name.to_lowercase(), d.name.as_str()))
        .collect();
    let mut display_names: HashMap<String, &str> = HashMap::new();
    let mut names: HashMap<String, (&str, &str)> = HashMap::new();
    for sheet in &wb.sheets {
        for table in &sheet.worksheet.tables {
            let dn = table.display_name.as_str();
            let mut error = |message: String| {
                report.push(ValidationIssue {
                    severity: Severity::Error,
                    category: IssueCategory::Structure,
                    message,
                    location: format!("{}!table({dn})", sheet.name),
                    suggestion: "Rename the table to a name Excel opens".into(),
                    auto_fixable: false,
                });
            };
            if let Some(problem) = name_problem(dn, NameKind::Table) {
                error(format!("Invalid table name '{dn}': {problem}"));
            }
            let key = dn.to_lowercase();
            if let Some(d) = defined.get(&key) {
                error(format!(
                    "Table name '{dn}' is already used by defined name '{d}' (names ignore case)"
                ));
            }
            match display_names.entry(key) {
                Entry::Occupied(other) => error(format!(
                    "Table name '{dn}' is already used by table '{}' (table names ignore case)",
                    other.get()
                )),
                Entry::Vacant(slot) => {
                    slot.insert(dn);
                }
            }
            let Some(name) = table.name.as_deref().filter(|n| !n.is_empty()) else {
                continue;
            };
            if name.encode_utf16().count() > MAX_TABLE_NAME_LEN {
                error(format!(
                    "Invalid table name attribute '{name}': it is longer than 255 characters"
                ));
            }
            match names.entry(name.to_lowercase()) {
                Entry::Occupied(other) => {
                    let (other_name, other_table) = other.get();
                    error(format!(
                        "Table name attribute '{name}' is already used by name attribute '{other_name}' of table '{other_table}' (table names ignore case)"
                    ));
                }
                Entry::Vacant(slot) => {
                    slot.insert((name, dn));
                }
            }
        }
    }
}

/// The `SUBTOTAL` function number Excel writes in the totals cell of a column
/// whose `totalsRowFunction` is `function`; `None` for `none`, `custom` and
/// values ECMA-376 does not define.
fn subtotal_code(function: &str) -> Option<u32> {
    Some(match function {
        "average" => 101,
        "countNums" => 102,
        "count" => 103,
        "max" => 104,
        "min" => 105,
        "stdDev" => 107,
        "sum" => 109,
        "var" => 110,
        _ => return None,
    })
}

/// `s` without `prefix`, matched ignoring ASCII case.
fn strip_prefix_ignore_case<'a>(s: &'a str, prefix: &str) -> Option<&'a str> {
    s.get(..prefix.len())
        .filter(|head| head.eq_ignore_ascii_case(prefix))
        .map(|_| &s[prefix.len()..])
}

/// A structured-reference column item (brackets removed) without Excel's `'`
/// escapes; `None` when a `[`, `]` or `#` in it is not escaped.
fn unescape_column_item(item: &str) -> Option<String> {
    let mut name = String::with_capacity(item.len());
    let mut chars = item.chars();
    while let Some(c) = chars.next() {
        match c {
            '\'' => name.push(chars.next()?),
            '[' | ']' | '#' => return None,
            _ => name.push(c),
        }
    }
    Some(name)
}

/// `s` with each `_xHHHH_` escape (four hex digits, any case) turned into
/// U+HHHH in one left-to-right pass; an escape that names no character (a
/// surrogate) stays as written. Excel writes the characters XML cannot hold
/// in a table part's column names and labels this way, and the text `_xHHHH_`
/// itself as `_x005F_xHHHH_`, in the column names and in the table formulas
/// alike (a column named `@a_x000a_b` is `@a_x005f_x000a_b` in both), while
/// the shared strings hold the characters. Borrowed when `s` has no `_x`.
fn decode_xstring(s: &str) -> Cow<'_, str> {
    if !s.contains("_x") {
        return Cow::Borrowed(s);
    }
    let mut out = String::with_capacity(s.len());
    let mut rest = s;
    while let Some(at) = rest.find("_x") {
        out.push_str(&rest[..at]);
        let tail = &rest[at..];
        let decoded = tail
            .get(2..6)
            .filter(|hex| hex.bytes().all(|b| b.is_ascii_hexdigit()))
            .filter(|_| tail.as_bytes().get(6) == Some(&b'_'))
            .and_then(|hex| u32::from_str_radix(hex, 16).ok())
            .and_then(char::from_u32);
        if let Some(c) = decoded {
            out.push(c);
            rest = &tail[7..];
        } else {
            out.push('_');
            rest = &tail[1..];
        }
    }
    out.push_str(rest);
    Cow::Owned(out)
}

/// `true` when `formula` is a totals formula Excel opens for column `column`
/// of table `table`: `SUBTOTAL(<code>,<table>[<column>])` as Excel writes it.
/// An Excel 365 probe also opened leading `=` signs, surrounding whitespace,
/// an `_xlfn.` prefix, any case in the function and column names, leading
/// zeros in the number, no table name, and `[[<column>]]`; it refused the
/// table name in another case, a space inside, and any other item. The item
/// and `column` are compared after [`decode_xstring`], since the table part
/// keeps `_xHHHH_` escapes that the formula holds as characters.
fn is_column_subtotal(formula: &str, code: u32, table: &str, column: &str) -> bool {
    let f = formula.trim().trim_start_matches('=');
    let f = strip_prefix_ignore_case(f, "_xlfn.").unwrap_or(f);
    let Some((number, reference)) = strip_prefix_ignore_case(f, "SUBTOTAL(")
        .and_then(|args| args.strip_suffix(')'))
        .and_then(|args| args.split_once(','))
    else {
        return false;
    };
    let reference = reference.strip_prefix(table).unwrap_or(reference);
    let item = reference
        .strip_prefix("[[")
        .and_then(|r| r.strip_suffix("]]"))
        .or_else(|| reference.strip_prefix('[')?.strip_suffix(']'));
    number.parse() == Ok(code)
        && item.and_then(unescape_column_item).is_some_and(|name| {
            decode_xstring(&name).to_lowercase() == decode_xstring(column).to_lowercase()
        })
}

/// What a totals cell of `column` in table `table` holds that Excel refuses
/// to open, as a short reason (`cell` is `None` when the sheet has none);
/// `None` when an Excel 365 probe opened it. A `SUBTOTAL` column takes a
/// value or its own formula ([`is_column_subtotal`]); a `custom` column with
/// a `totalsRowFormula` takes any formula and nothing else; a column without
/// a function takes text, which must be its `totalsRowLabel` when it has one,
/// or nothing; the text and the label are compared after [`decode_xstring`],
/// since the table part keeps `_xHHHH_` escapes that the cell holds as
/// characters. Excel saves what a user types into a totals cell in those
/// forms. A shared-formula child (its text is in another cell) is not checked.
fn totals_cell_problem(
    table: &str,
    column: &crate::ooxml::tables::TableColumn,
    cell: Option<&Cell>,
) -> Option<&'static str> {
    let formula = cell
        .and_then(|c| c.formula.as_deref())
        .filter(|f| !f.is_empty());
    if formula.is_none() && cell.is_some_and(|c| c.formula_type.is_some()) {
        return None;
    }
    let function = column.totals_row_function.as_deref();
    if let Some(code) = function.and_then(subtotal_code) {
        return formula
            .filter(|f| !is_column_subtotal(f, code, table, &column.name))
            .map(|_| "a formula other than the column's SUBTOTAL");
    }
    if function == Some("custom") {
        let has_formula = column
            .totals_row_formula
            .as_deref()
            .is_some_and(|f| !f.is_empty());
        return (has_formula && formula.is_none())
            .then_some("no formula in a custom column that has a totalsRowFormula");
    }
    if function.is_some_and(|f| f != "none") {
        return None;
    }
    let cell = cell?;
    if formula.is_some() {
        return Some("a formula in a column without a totals function");
    }
    let value = decode_xstring(
        cell.value
            .as_deref()
            .or(cell.inline_string.as_deref())
            .unwrap_or_default(),
    );
    let label = decode_xstring(column.totals_row_label.as_deref().unwrap_or_default());
    match cell.cell_type {
        CellType::Number | CellType::Boolean if !value.is_empty() => {
            Some("a number or boolean in a column without a totals function")
        }
        CellType::SharedString | CellType::InlineStr | CellType::FormulaStr
            if !value.is_empty() && !label.is_empty() && value != label =>
        {
            Some("text other than the column's totalsRowLabel")
        }
        _ => None,
    }
}

/// The cells of the sheet's rows whose 1-based index is in `wanted`, by
/// (row index, cell reference text), from one pass over `rows`: for each
/// index only the first row with it, and in that row the first cell with
/// each reference. That is what a search of the rows and cells in file order
/// finds, so a hostile file whose rows or cells are out of order or repeated
/// gets the same answer, and each lookup costs O(1).
fn cells_of_rows(rows: &[Row], mut wanted: HashSet<u32>) -> HashMap<(u32, &str), &Cell> {
    let mut cells = HashMap::new();
    for row in rows {
        if wanted.remove(&row.index) {
            for cell in &row.cells {
                cells
                    .entry((row.index, cell.reference.as_str()))
                    .or_insert(cell);
            }
        }
    }
    cells
}

/// Report the totals-row cells (the last row of a table with
/// `totalsRowCount` 1) that Excel refuses to open; see [`totals_cell_problem`].
/// The sheet's totals-row cells are indexed once ([`cells_of_rows`]), so each
/// table costs time proportional to its own columns.
fn validate_table_totals(report: &mut ValidationReport, wb: &WorkbookData) {
    for sheet in &wb.sheets {
        let ws = &sheet.worksheet;
        let tables: Vec<_> = ws
            .tables
            .iter()
            .filter(|t| t.totals_row_count > 0)
            .filter_map(|table| {
                let (a, b) = table.ref_range.split_once(':')?;
                Some((table, CellRef::parse(a).ok()?, CellRef::parse(b).ok()?))
            })
            .collect();
        if tables.is_empty() {
            continue;
        }
        let cells = cells_of_rows(
            &ws.rows,
            tables.iter().map(|&(_, _, end)| end.row + 1).collect(),
        );
        for (table, start, end) in tables {
            for (col, column) in (start.col..=end.col).zip(&table.columns) {
                let reference = CellRef { col, row: end.row }.to_a1();
                let cell = cells.get(&(end.row + 1, reference.as_str())).copied();
                if let Some(problem) = totals_cell_problem(&table.display_name, column, cell)
                    && !report.omit_if_full(Severity::Error)
                {
                    report.push(ValidationIssue {
                        severity: Severity::Error,
                        category: IssueCategory::Structure,
                        message: format!(
                            "Table '{}' totals cell {reference} holds {problem}",
                            table.display_name
                        ),
                        location: format!("{}!{reference}", sheet.name),
                        suggestion: "Make the totals cell agree with its column, or clear it"
                            .into(),
                        auto_fixable: false,
                    });
                }
            }
        }
    }
}

/// A `@` this-row reference found by [`this_row_shorthand`].
struct ThisRowShorthand<'a> {
    /// The reference as written, table name included: `[@Qty]`, `Sales[@[Unit Cost]]`, `[@]`.
    reference: &'a str,
    /// The table name before the `[`; empty when the reference names none.
    table: &'a str,
    /// What follows the `@`: `Qty`, `[Unit Cost]`, or empty for `[@]`.
    item: &'a str,
}

/// The first `@` this-row reference in `formula` (`[@Qty]`, `[@[Unit Cost]]`,
/// `Sales[@Qty]`, `[@]`) outside `"..."` string literals and `'...'` quoted
/// names; `None` when there is none or a quote or bracket is never closed.
/// The table name is the run of bytes a table name may hold (none that
/// [`is_forbidden_name_byte`] matches) just before the `[`. A `[@...]` that
/// names a column of that table (`columns_of(table)`, which gets `""` for no
/// table name; see [`names_column`]) is a column reference, not the
/// shorthand: Excel writes a whole-column reference to a column named
/// `@Home` as `AtT[@Home]`. A table name right after `]!` names a table of
/// another workbook (Excel stores `Book.xlsx!Ext[@Home]` as
/// `[1]!Ext[@Home]`, whether that workbook is open or closed and whatever its
/// path), whose columns the workbook does not hold, so its `[@...]` is never
/// reported. One pass over `formula`, with one `columns_of` call and one
/// lookup per `[@`; every slice starts and ends at an ASCII byte.
fn this_row_shorthand<'f, 'c, C: ColumnKeys + ?Sized + 'c>(
    formula: &'f str,
    columns_of: impl Fn(&str) -> &'c C,
) -> Option<ThisRowShorthand<'f>> {
    let b = formula.as_bytes();
    let mut run_start = 0;
    let mut i = 0;
    while i < b.len() {
        let end = match b[i] {
            quote @ (b'"' | b'\'') => quoted_end(b, i, quote)?,
            b'[' => {
                let end = bracket_end(b, i)?;
                let table = &formula[run_start..i];
                if b.get(i + 1) == Some(&b'@')
                    && !(run_start < i && run_start >= 2 && b[run_start - 2..run_start] == *b"]!")
                    && !names_column(&formula[i + 1..end - 1], columns_of(table))
                {
                    return Some(ThisRowShorthand {
                        reference: &formula[run_start..end],
                        table,
                        item: &formula[i + 2..end - 1],
                    });
                }
                end
            }
            c if !is_forbidden_name_byte(c) => {
                i += 1;
                continue;
            }
            _ => i + 1,
        };
        i = end;
        run_start = end;
    }
    None
}

/// The column names of a table as [`names_column`] looks them up: each
/// decoded ([`decode_xstring`]) and lower-cased.
trait ColumnKeys {
    /// `true` when one of the columns has the decoded, lower-cased name `key`.
    fn has_key(&self, key: &str) -> bool;
}

/// The set [`validate_table_formulas`] builds once per table: one lookup.
impl ColumnKeys for HashSet<String> {
    fn has_key(&self, key: &str) -> bool {
        self.contains(key)
    }
}

/// A column list, decoded and lower-cased on each lookup (the scan's unit tests).
#[cfg(test)]
impl ColumnKeys for [crate::ooxml::tables::TableColumn] {
    fn has_key(&self, key: &str) -> bool {
        self.iter()
            .any(|column| decode_xstring(&column.name).to_lowercase() == key)
    }
}

/// `true` when `inner`, the text of a structured-reference bracket without
/// the brackets, is the name of one of `columns`, ignoring case. `inner` is
/// decoded ([`decode_xstring`]) as the column names are, since the table part
/// escapes both alike, and then loses its `'` escapes; one lookup.
fn names_column<C: ColumnKeys + ?Sized>(inner: &str, columns: &C) -> bool {
    let inner = decode_xstring(inner);
    let mut name = String::with_capacity(inner.len());
    let mut chars = inner.chars();
    while let Some(c) = chars.next() {
        // `bracket_end` closed the bracket, so a `'` is last only when it was
        // decoded from `_x0027_`, and is dropped then.
        name.extend(if c == '\'' { chars.next() } else { Some(c) });
    }
    columns.has_key(&name.to_lowercase())
}

/// Index just past the `quote` closing the one at `pos` (a doubled quote
/// stays inside), or `None`.
fn quoted_end(b: &[u8], pos: usize, quote: u8) -> Option<usize> {
    let mut i = pos + 1;
    while i < b.len() {
        if b[i] == quote {
            if b.get(i + 1) != Some(&quote) {
                return Some(i + 1);
            }
            i += 1;
        }
        i += 1;
    }
    None
}

/// Index just past the `]` matching the `[` at `pos` (`'` escapes the next
/// byte), or `None`.
fn bracket_end(b: &[u8], pos: usize) -> Option<usize> {
    let mut depth = 0_usize;
    let mut i = pos;
    while i < b.len() {
        match b[i] {
            b'\'' => i += 1,
            b'[' => depth += 1,
            b']' => {
                depth -= 1;
                if depth == 0 {
                    return Some(i + 1);
                }
            }
            _ => {}
        }
        i += 1;
    }
    None
}

/// Report a table column whose `calculatedColumnFormula` or
/// `totalsRowFormula` uses the `@` this-row shorthand
/// ([`this_row_shorthand`]): Excel 365 refuses to open such a table
/// definition. The message gives the form Excel writes there instead
/// (`Sales[[#This Row],[Qty]]` for `[@Qty]`, `Sales[#This Row]` for `[@]`).
/// A `[@...]` naming a column of its table (the table being checked when the
/// reference names none, else the first workbook table with that display
/// name, ignoring case), or one of a table in another workbook, is a column
/// reference and passes. The cost is linear in the formulas' length plus the
/// columns: the display names are mapped once, and a table's column names
/// are decoded and lower-cased into a set once, the first time a formula
/// names it. The TypeScript `Worksheet.addTable` throws a message with the
/// same text, except that it quotes the column and table names with JSON
/// double quotes.
fn validate_table_formulas(report: &mut ValidationReport, wb: &WorkbookData) {
    let tables: Vec<_> = wb
        .sheets
        .iter()
        .flat_map(|sheet| &sheet.worksheet.tables)
        .collect();
    let mut by_name: HashMap<String, usize> = HashMap::with_capacity(tables.len());
    for (index, table) in tables.iter().enumerate() {
        by_name
            .entry(table.display_name.to_lowercase())
            .or_insert(index);
    }
    let keys: Vec<OnceCell<HashSet<String>>> = tables.iter().map(|_| OnceCell::new()).collect();
    let column_keys = |index: usize| {
        keys[index].get_or_init(|| {
            tables[index]
                .columns
                .iter()
                .map(|column| decode_xstring(&column.name).to_lowercase())
                .collect()
        })
    };
    let no_columns = HashSet::new();
    // `tables` lists the tables in this order, so `own` is each one's index.
    let mut own = 0;
    for sheet in &wb.sheets {
        for table in &sheet.worksheet.tables {
            let dn = table.display_name.as_str();
            let dn_key = dn.to_lowercase();
            let columns_of = |name: &str| {
                let key = name.to_lowercase();
                let index = if key.is_empty() || key == dn_key {
                    Some(own)
                } else {
                    by_name.get(&key).copied()
                };
                index.map_or(&no_columns, column_keys)
            };
            for column in &table.columns {
                for (field, formula) in [
                    ("calculatedColumnFormula", &column.calculated_column_formula),
                    ("totalsRowFormula", &column.totals_row_formula),
                ] {
                    let Some(found) = formula
                        .as_deref()
                        .and_then(|formula| this_row_shorthand(formula, columns_of))
                    else {
                        continue;
                    };
                    if report.omit_if_full(Severity::Error) {
                        continue;
                    }
                    let name = if found.table.is_empty() {
                        dn
                    } else {
                        found.table
                    };
                    let file_form = match found.item {
                        "" => format!("{name}[#This Row]"),
                        item if item.starts_with('[') => format!("{name}[[#This Row],{item}]"),
                        item => format!("{name}[[#This Row],[{item}]]"),
                    };
                    report.push(ValidationIssue {
                        severity: Severity::Error,
                        category: IssueCategory::Structure,
                        message: format!(
                            "Invalid {field} of column '{}' in table '{dn}': Excel refuses the @ shorthand {} in a table definition; write {file_form} instead",
                            column.name, found.reference
                        ),
                        location: format!("{}!table({dn})", sheet.name),
                        suggestion: "Write each this-row reference as Table[[#This Row],[Column]]"
                            .into(),
                        auto_fixable: false,
                    });
                }
            }
            own += 1;
        }
    }
}

// -- Cells --

fn validate_cells(report: &mut ValidationReport, wb: &WorkbookData) {
    let xf_count = wb.styles.cell_xfs.len();

    for sheet in &wb.sheets {
        for row in &sheet.worksheet.rows {
            for cell in &row.cells {
                let loc = || format!("{}!{}", sheet.name, cell.reference);

                // Validate the cell reference (the parser enforces the XFD1048576 bound).
                if let Err(e) = CellRef::parse(&cell.reference)
                    && !report.omit_if_full(Severity::Error)
                {
                    report.push(ValidationIssue {
                        severity: Severity::Error,
                        category: IssueCategory::CellReference,
                        message: format!("Invalid cell reference '{}': {e}", cell.reference),
                        location: loc(),
                        suggestion: "Fix cell reference format".into(),
                        auto_fixable: false,
                    });
                }

                // Validate style index.
                if let Some(si) = cell.style_index
                    && si as usize >= xf_count
                    && !report.omit_if_full(Severity::Error)
                {
                    report.push(ValidationIssue {
                        severity: Severity::Error,
                        category: IssueCategory::StyleIndex,
                        message: format!(
                            "Cell {} styleIndex={si} exceeds cellXfs count ({xf_count})",
                            cell.reference
                        ),
                        location: loc(),
                        suggestion: "Clamp styleIndex to 0".into(),
                        auto_fixable: true,
                    });
                }

                // SharedString cell must have a value.
                if cell.cell_type == CellType::SharedString
                    && cell.value.is_none()
                    && !report.omit_if_full(Severity::Error)
                {
                    report.push(ValidationIssue {
                        severity: Severity::Error,
                        category: IssueCategory::SharedString,
                        message: format!("SharedString cell {} has no value", cell.reference),
                        location: loc(),
                        suggestion: "Set value to an empty string or change cell type".into(),
                        auto_fixable: true,
                    });
                }
            }
        }
    }
}

// -- Merge cells --

/// A merge range as a normalized rectangle (zero-based, inclusive corners).
#[derive(Clone, Copy)]
struct MergeRect {
    r1: u32,
    c1: u32,
    r2: u32,
    c2: u32,
    /// Index of the range in the sheet's `merge_cells` (for reporting).
    idx: usize,
}

/// Columns a sheet has (`XFD` is the last, column 16383).
const SHEET_COLUMNS: usize = 16_384;

/// A segment tree over the sheet's columns holding, per column, the lowest row
/// any raised range has reached: range "raise to at least" and range maximum,
/// both O(log columns). Tags stay on the nodes (no push-down), so a query adds
/// the tags of the nodes it passes.
struct ColumnReach {
    /// Largest `r2 + 1` raised on a node's whole range (0 = none).
    tag: Vec<u32>,
    /// Largest `r2 + 1` raised anywhere in a node's range (0 = none).
    best: Vec<u32>,
}

impl ColumnReach {
    fn new() -> Self {
        Self {
            tag: vec![0; 2 * SHEET_COLUMNS],
            best: vec![0; 2 * SHEET_COLUMNS],
        }
    }

    /// Raise columns `c1..=c2` (clamped to the sheet) to at least row `r2`.
    fn raise(&mut self, c1: u32, c2: u32, r2: u32) {
        let (lo, hi) = Self::clamp(c1, c2);
        self.raise_in(1, 0, SHEET_COLUMNS - 1, lo, hi, r2.saturating_add(1));
    }

    /// The lowest row reached on any of columns `c1..=c2`, if any.
    fn max(&self, c1: u32, c2: u32) -> Option<u32> {
        let (lo, hi) = Self::clamp(c1, c2);
        self.max_in(1, 0, SHEET_COLUMNS - 1, lo, hi, 0)
            .checked_sub(1)
    }

    fn clamp(c1: u32, c2: u32) -> (usize, usize) {
        let last = SHEET_COLUMNS - 1;
        ((c1 as usize).min(last), (c2 as usize).min(last))
    }

    fn raise_in(&mut self, node: usize, nl: usize, nr: usize, lo: usize, hi: usize, v: u32) {
        self.best[node] = self.best[node].max(v);
        if lo <= nl && nr <= hi {
            self.tag[node] = self.tag[node].max(v);
            return;
        }
        let mid = usize::midpoint(nl, nr);
        if lo <= mid {
            self.raise_in(2 * node, nl, mid, lo, hi, v);
        }
        if hi > mid {
            self.raise_in(2 * node + 1, mid + 1, nr, lo, hi, v);
        }
    }

    fn max_in(&self, node: usize, nl: usize, nr: usize, lo: usize, hi: usize, above: u32) -> u32 {
        if lo <= nl && nr <= hi {
            return above.max(self.best[node]);
        }
        let above = above.max(self.tag[node]);
        let mid = usize::midpoint(nl, nr);
        let mut out = above;
        if lo <= mid {
            out = out.max(self.max_in(2 * node, nl, mid, lo, hi, above));
        }
        if hi > mid {
            out = out.max(self.max_in(2 * node + 1, mid + 1, nr, lo, hi, above));
        }
        out
    }
}

/// Parse a `A1:B2` merge range into its two corners, without normalizing.
///
/// `None` means the text has no `:`; `Some(Err)` means an endpoint is not a
/// valid in-sheet cell reference.
fn parse_merge_corners(merge_ref: &str) -> Option<crate::Result<(CellRef, CellRef)>> {
    let (start, end) = merge_ref.split_once(':')?;
    Some(CellRef::parse(start).and_then(|s| CellRef::parse(end).map(|e| (s, e))))
}

const fn is_reversed(start: CellRef, end: CellRef) -> bool {
    start.row > end.row || start.col > end.col
}

fn normalized_corners(start: CellRef, end: CellRef) -> (CellRef, CellRef) {
    (
        CellRef {
            col: start.col.min(end.col),
            row: start.row.min(end.row),
        },
        CellRef {
            col: start.col.max(end.col),
            row: start.row.max(end.row),
        },
    )
}

/// Validate merge ranges; the cost depends on the number of ranges, never on
/// the area they cover.
fn validate_merge_cells(report: &mut ValidationReport, wb: &WorkbookData) {
    for sheet in &wb.sheets {
        let mut rects: Vec<MergeRect> = Vec::with_capacity(sheet.worksheet.merge_cells.len());

        for (idx, merge_ref) in sheet.worksheet.merge_cells.iter().enumerate() {
            let loc = || format!("{}!mergeCell({})", sheet.name, merge_ref);

            let (start, end) = match parse_merge_corners(merge_ref) {
                None => {
                    if report.omit_if_full(Severity::Error) {
                        continue;
                    }
                    report.push(ValidationIssue {
                        severity: Severity::Error,
                        category: IssueCategory::MergeCell,
                        message: format!("Invalid merge range '{merge_ref}'"),
                        location: loc(),
                        suggestion: "Use format 'A1:B2'".into(),
                        auto_fixable: false,
                    });
                    continue;
                }
                Some(Err(e)) => {
                    if report.omit_if_full(Severity::Error) {
                        continue;
                    }
                    report.push(ValidationIssue {
                        severity: Severity::Error,
                        category: IssueCategory::MergeCell,
                        message: format!(
                            "Cannot parse merge range endpoints in '{merge_ref}': {e}"
                        ),
                        location: loc(),
                        suggestion: "Fix cell references".into(),
                        auto_fixable: false,
                    });
                    continue;
                }
                Some(Ok(corners)) => corners,
            };

            let (lo, hi) = normalized_corners(start, end);
            if is_reversed(start, end) && !report.omit_if_full(Severity::Warning) {
                report.push(ValidationIssue {
                    severity: Severity::Warning,
                    category: IssueCategory::MergeCell,
                    message: format!(
                        "Merge range '{merge_ref}' is reversed; normalized it is '{}:{}'",
                        lo.to_a1(),
                        hi.to_a1()
                    ),
                    location: loc(),
                    suggestion: "Write the range as top-left:bottom-right".into(),
                    auto_fixable: true,
                });
            }
            rects.push(MergeRect {
                r1: lo.row,
                c1: lo.col,
                r2: hi.row,
                c2: hi.col,
                idx,
            });
        }

        // Sweep the rectangles in (top row, left column) order. Every earlier
        // rectangle has a top row at or above the current one, so two overlap
        // exactly when an earlier rectangle on a shared column reaches down to
        // the current top row. `reach` keeps, per column, the lowest row any
        // earlier rectangle covers, so each range costs O(log columns) however
        // many rectangles there are. The later range of an overlapping pair is
        // reported once.
        if rects.len() < 2 {
            continue;
        }
        rects.sort_unstable_by_key(|r| (r.r1, r.c1, r.idx));
        let mut reach = ColumnReach::new();
        let mut overlapping: Vec<usize> = Vec::new();
        for rect in rects {
            if reach.max(rect.c1, rect.c2).is_some_and(|r2| r2 >= rect.r1) {
                overlapping.push(rect.idx);
            }
            reach.raise(rect.c1, rect.c2, rect.r2);
        }

        // Report in document order.
        overlapping.sort_unstable();
        for idx in overlapping {
            if report.omit_if_full(Severity::Error) {
                continue;
            }
            let merge_ref = &sheet.worksheet.merge_cells[idx];
            report.push(ValidationIssue {
                severity: Severity::Error,
                category: IssueCategory::MergeCell,
                message: format!("Merge range '{merge_ref}' overlaps with another merge region"),
                location: format!("{}!mergeCell({})", sheet.name, merge_ref),
                suggestion: "Remove overlapping merge or adjust ranges".into(),
                auto_fixable: false,
            });
        }
    }
}

// -- Defined names --

/// Report the defined names Excel refuses to open: an empty name, a name
/// breaking [`name_problem`] for [`NameKind::Defined`] (the `_xlnm.` built-in
/// names are not checked), and a print name with an unqualified area; warn
/// about an empty value.
fn validate_defined_names(report: &mut ValidationReport, wb: &WorkbookData) {
    for (i, dn) in wb.defined_names.iter().enumerate() {
        let loc = || format!("definedNames[{i}]");

        if dn.name.is_empty() {
            report.push(ValidationIssue {
                severity: Severity::Error,
                category: IssueCategory::DefinedName,
                message: "Defined name has empty name".into(),
                location: loc(),
                suggestion: "Set a valid name".into(),
                auto_fixable: false,
            });
        } else if strip_prefix_ignore_case(&dn.name, "_xlnm.").is_none()
            && let Some(problem) = name_problem(&dn.name, NameKind::Defined)
            && !report.omit_if_full(Severity::Error)
        {
            report.push(ValidationIssue {
                severity: Severity::Error,
                category: IssueCategory::DefinedName,
                message: format!("Invalid defined name '{}': {problem}", dn.name),
                location: loc(),
                suggestion: "Rename the defined name to a name Excel opens".into(),
                auto_fixable: false,
            });
        }

        if dn.value.is_empty() {
            report.push(ValidationIssue {
                severity: Severity::Warning,
                category: IssueCategory::DefinedName,
                message: format!("Defined name '{}' has empty value/reference", dn.name),
                location: loc(),
                suggestion: "Set a valid cell range reference".into(),
                auto_fixable: false,
            });
        } else if is_print_name(&dn.name) && has_unqualified_area(&dn.value) {
            report.push(ValidationIssue {
                severity: Severity::Error,
                category: IssueCategory::DefinedName,
                message: format!(
                    "Defined name '{}' has an area without a sheet qualifier: '{}'",
                    dn.name, dn.value
                ),
                location: loc(),
                suggestion: "Qualify every area with its sheet, e.g. Sheet1!$A$1:$B$2".into(),
                auto_fixable: false,
            });
        }
    }
}

/// `_xlnm.Print_Area` and `_xlnm.Print_Titles` in any ASCII letter case (Excel
/// matches built-in names without regard to case): the built-in names Excel
/// refuses to open when a value is not sheet-qualified.
fn is_print_name(name: &str) -> bool {
    name.eq_ignore_ascii_case("_xlnm.Print_Area") || name.eq_ignore_ascii_case("_xlnm.Print_Titles")
}

/// Whether a comma-separated list of areas has one that does not start with a
/// sheet prefix (`Sheet1!` or `'My Sheet'!`). A comma inside a quoted sheet
/// name does not split.
fn has_unqualified_area(value: &str) -> bool {
    let mut in_quote = false;
    let mut start = 0;
    let mut areas = Vec::new();
    for (i, ch) in value.char_indices() {
        match ch {
            '\'' => in_quote = !in_quote,
            ',' if !in_quote => {
                areas.push(&value[start..i]);
                start = i + 1;
            }
            _ => {}
        }
    }
    areas.push(&value[start..]);
    areas.iter().any(|area| !has_sheet_prefix(area.trim()))
}

/// Whether `area` starts with a quoted sheet name (`''` is a literal
/// apostrophe) or a bare one, followed by `!`. A leading `#` is an error
/// literal such as `#REF!`, not a sheet.
fn has_sheet_prefix(area: &str) -> bool {
    if let Some(rest) = area.strip_prefix('\'') {
        let mut chars = rest.char_indices().peekable();
        while let Some((i, ch)) = chars.next() {
            if ch != '\'' {
                continue;
            }
            if chars.next_if(|&(_, next)| next == '\'').is_some() {
                continue;
            }
            return rest[i + 1..].starts_with('!');
        }
        false
    } else {
        area.find('!')
            .is_some_and(|bang| bang > 0 && !area.starts_with('#'))
    }
}

// -- Data validations --

fn validate_data_validations(report: &mut ValidationReport, wb: &WorkbookData) {
    for sheet in &wb.sheets {
        for (i, dv) in sheet.worksheet.data_validations.iter().enumerate() {
            let loc = || format!("{}!dataValidation[{i}]", sheet.name);

            if dv.sqref.is_empty() && !report.omit_if_full(Severity::Warning) {
                report.push(ValidationIssue {
                    severity: Severity::Warning,
                    category: IssueCategory::DataValidation,
                    message: "Data validation has empty sqref".into(),
                    location: loc(),
                    suggestion: "Set target cell range".into(),
                    auto_fixable: false,
                });
            }

            // List validations should have formula1.
            if dv.validation_type.as_deref() == Some("list")
                && dv.formula1.is_none()
                && !report.omit_if_full(Severity::Warning)
            {
                report.push(ValidationIssue {
                    severity: Severity::Warning,
                    category: IssueCategory::DataValidation,
                    message: "List validation missing formula1 (source list)".into(),
                    location: loc(),
                    suggestion: "Set formula1 to a comma-separated list or range".into(),
                    auto_fixable: false,
                });
            }
        }
    }
}

// -- Theme --

fn validate_theme(report: &mut ValidationReport, wb: &WorkbookData) {
    if wb.theme_colors.is_none() {
        report.push(ValidationIssue {
            severity: Severity::Info,
            category: IssueCategory::Theme,
            message: "No theme colors present; default Office Theme (Excel 2016-2022) will be used"
                .into(),
            location: "themeColors".into(),
            suggestion: "No action needed — defaults are injected automatically".into(),
            auto_fixable: true,
        });
    }
}

// -- Metadata --

/// The `DocProperties` text fields whose value must not be only white space.
/// One list serves [`validate_metadata`] (`&` borrows) and [`repair_metadata`]
/// (`&mut` borrows), so the two can never disagree. The date fields
/// (`created`, `modified`) are checked by the W3CDTF rule instead.
macro_rules! doc_text_fields {
    ($props:expr, $($borrow:tt)+) => {
        [
            ("title", $($borrow)+ $props.title),
            ("subject", $($borrow)+ $props.subject),
            ("creator", $($borrow)+ $props.creator),
            ("keywords", $($borrow)+ $props.keywords),
            ("description", $($borrow)+ $props.description),
            ("lastModifiedBy", $($borrow)+ $props.last_modified_by),
            ("category", $($borrow)+ $props.category),
            ("contentStatus", $($borrow)+ $props.content_status),
            ("lastPrinted", $($borrow)+ $props.last_printed),
            ("language", $($borrow)+ $props.language),
            ("identifier", $($borrow)+ $props.identifier),
            ("version", $($borrow)+ $props.version),
            ("application", $($borrow)+ $props.application),
            ("company", $($borrow)+ $props.company),
            ("manager", $($borrow)+ $props.manager),
            ("appVersion", $($borrow)+ $props.app_version),
            ("hyperlinkBase", $($borrow)+ $props.hyperlink_base),
            ("revision", $($borrow)+ $props.revision),
            ("totalTime", $($borrow)+ $props.total_time),
            ("docSecurity", $($borrow)+ $props.doc_security),
        ]
    };
}

/// `true` for a non-empty value that is only white space.
fn is_blank(s: &str) -> bool {
    !s.is_empty() && s.trim().is_empty()
}

fn validate_metadata(report: &mut ValidationReport, wb: &WorkbookData) {
    if let Some(ref props) = wb.doc_properties {
        for (field_name, value) in [("created", &props.created), ("modified", &props.modified)] {
            if let Some(date_str) = value
                && !is_w3cdtf(date_str)
            {
                report.push(ValidationIssue {
                    severity: Severity::Warning,
                    category: IssueCategory::Metadata,
                    message: format!(
                        "Document property '{field_name}' has non-ISO-8601 (W3CDTF) value '{date_str}'"
                    ),
                    location: format!("docProperties.{field_name}"),
                    suggestion: "Use ISO 8601 format: YYYY-MM-DDTHH:MM:SSZ".into(),
                    auto_fixable: true,
                });
            }
        }

        for (field_name, value) in doc_text_fields!(props, &) {
            if let Some(s) = value
                && is_blank(s)
            {
                report.push(ValidationIssue {
                    severity: Severity::Info,
                    category: IssueCategory::Metadata,
                    message: format!("Document property '{field_name}' is whitespace-only"),
                    location: format!("docProperties.{field_name}"),
                    suggestion: "Clear the field or set meaningful content".into(),
                    auto_fixable: true,
                });
            }
        }
    }
}

/// How a string fares as a W3CDTF date-time (the ISO 8601 profile of
/// `dcterms:created`, W3C NOTE-datetime).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum W3cdtf {
    /// `YYYY`, `YYYY-MM`, `YYYY-MM-DD`, or `YYYY-MM-DDThh:mm[:ss[.s+]]` with a
    /// `Z` or `+hh:mm` / `-hh:mm` designator, calendar-valid.
    Valid,
    /// A calendar-valid date-time that lacks the required time-zone designator.
    MissingZone,
    Invalid,
}

const fn is_leap_year(year: u32) -> bool {
    (year.is_multiple_of(4) && !year.is_multiple_of(100)) || year.is_multiple_of(400)
}

const fn days_in_month(year: u32, month: u32) -> u32 {
    match month {
        2 if is_leap_year(year) => 29,
        2 => 28,
        4 | 6 | 9 | 11 => 30,
        _ => 31,
    }
}

/// The `n`-digit decimal number at `at`, or `None` when those bytes are not all
/// ASCII digits (or run past the end).
fn digits(b: &[u8], at: usize, n: usize) -> Option<u32> {
    let d = b.get(at..at.checked_add(n)?)?;
    d.iter()
        .all(u8::is_ascii_digit)
        .then(|| d.iter().fold(0, |acc, &c| acc * 10 + u32::from(c - b'0')))
}

fn w3cdtf_status(s: &str) -> W3cdtf {
    parse_w3cdtf(s.as_bytes()).unwrap_or(W3cdtf::Invalid)
}

fn parse_w3cdtf(b: &[u8]) -> Option<W3cdtf> {
    let year = digits(b, 0, 4)?;
    if b.len() == 4 {
        return Some(W3cdtf::Valid);
    }
    if b[4] != b'-' {
        return None;
    }
    let month = digits(b, 5, 2).filter(|m| (1..=12).contains(m))?;
    if b.len() == 7 {
        return Some(W3cdtf::Valid);
    }
    if b[7] != b'-' {
        return None;
    }
    digits(b, 8, 2).filter(|d| (1..=days_in_month(year, month)).contains(d))?;
    if b.len() == 10 {
        return Some(W3cdtf::Valid);
    }
    if b[10] != b'T' {
        return None;
    }
    digits(b, 11, 2).filter(|h| *h < 24)?;
    if b.get(13) != Some(&b':') {
        return None;
    }
    digits(b, 14, 2).filter(|m| *m < 60)?;
    let mut i = 16;
    if b.get(i) == Some(&b':') {
        digits(b, i + 1, 2).filter(|s| *s < 60)?;
        i += 3;
        if b.get(i) == Some(&b'.') {
            let frac = b[i + 1..].iter().take_while(|c| c.is_ascii_digit()).count();
            if frac == 0 {
                return None;
            }
            i += 1 + frac;
        }
    }
    match &b[i..] {
        [] => Some(W3cdtf::MissingZone),
        [b'Z'] => Some(W3cdtf::Valid),
        [b'+' | b'-', ..] if b.len() == i + 6 && b[i + 3] == b':' => {
            digits(b, i + 1, 2).filter(|h| *h < 24)?;
            digits(b, i + 4, 2).filter(|m| *m < 60)?;
            Some(W3cdtf::Valid)
        }
        _ => None,
    }
}

/// `true` when `s` is a valid W3CDTF date or date-time (see [`W3cdtf::Valid`]).
fn is_w3cdtf(s: &str) -> bool {
    w3cdtf_status(s) == W3cdtf::Valid
}

// -- Structure --

fn validate_structure(report: &mut ValidationReport, wb: &WorkbookData) {
    // Check row ordering within sheets.
    for sheet in &wb.sheets {
        let mut prev_row_idx: u32 = 0;
        for (i, row) in sheet.worksheet.rows.iter().enumerate() {
            if i > 0 && row.index <= prev_row_idx && !report.omit_if_full(Severity::Warning) {
                report.push(ValidationIssue {
                    severity: Severity::Warning,
                    category: IssueCategory::Structure,
                    message: format!(
                        "Row index {} is not strictly ascending (prev={})",
                        row.index, prev_row_idx
                    ),
                    location: format!("{}!row[{i}]", sheet.name),
                    suggestion: "Sort rows by index".into(),
                    auto_fixable: true,
                });
            }
            prev_row_idx = row.index;
        }
    }
}

// -- Threaded comments (ids, one thread per cell, text length) and person ids --

/// Longest threaded comment text Excel 16 opens, in UTF-16 code units: it
/// refuses a file with a 32,768-unit root or reply, and counts an astral
/// character as two units (Excel 16.0 build 20527, probed 2026-11-12).
const MAX_THREADED_COMMENT_TEXT: usize = 32_767;

/// What an issue says of a threaded-comment or person id Excel cannot use:
/// `field` names the attribute.
fn thread_id_message(field: &str, id: &str) -> String {
    format!(
        "{field} '{id}' is not a GUID in braces ({{XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX}}): Excel refuses the file, or for a parentId drops the reply"
    )
}

/// The cell a threaded comment `ref` names, as the TS API stores it: without
/// `$` and in ASCII upper case, so `a1`, `$A$1` and `A1` are one cell.
fn thread_cell_key(ref_cell: &str) -> String {
    ref_cell.replace('$', "").to_ascii_uppercase()
}

/// The length of `text` in UTF-16 code units when it is longer than
/// [`MAX_THREADED_COMMENT_TEXT`], the unit Excel counts; `None` otherwise.
fn thread_text_overflow(text: &str) -> Option<usize> {
    // A UTF-8 string never has more UTF-16 code units than bytes.
    if text.len() <= MAX_THREADED_COMMENT_TEXT {
        return None;
    }
    let units = text.encode_utf16().count();
    (units > MAX_THREADED_COMMENT_TEXT).then_some(units)
}

/// Reports each threaded comment `id`, `personId` and `parentId`, and each
/// person `id`, that is not a GUID in braces
/// ([`threaded_comments::is_braced_guid`]): Excel refuses a file with such an
/// id and drops a reply with such a `parentId`. Also reports a reply whose
/// `parentId` names a comment on its sheet that is itself a reply, not a
/// thread root (ids compared ASCII case-insensitively, as Excel does): Excel
/// refuses a reply to a reply. Also reports a thread root on a cell that an
/// earlier root on its sheet already holds (`ref` compared without `$` and
/// ignoring ASCII case): Excel shows only the first thread on a cell and the
/// later one is lost. Also reports a comment whose text is longer than
/// [`MAX_THREADED_COMMENT_TEXT`] UTF-16 code units: Excel refuses the file.
/// The writer writes ids, refs and texts as given (an id is the caller's
/// key), so a raw model learns it here.
fn validate_thread_ids(report: &mut ValidationReport, wb: &WorkbookData) {
    const GUID_FORM: &str = "Use the {UPPER-CASE} GUID form Excel writes, and keep every personId, parentId and tc= comment author naming it in step";
    // Each issue repeats a sheet name or an id the file controls, so it is
    // built only when the report still stores issues.
    let issue = |message: String, location: String, suggestion: &str| ValidationIssue {
        severity: Severity::Error,
        category: IssueCategory::Structure,
        message,
        location,
        suggestion: suggestion.into(),
        auto_fixable: false,
    };
    for sheet in &wb.sheets {
        let threads = &sheet.worksheet.threaded_comments;
        // The sheet's thread roots and replies, by id in ASCII upper case.
        let (mut roots, mut replies) = (HashSet::new(), HashSet::new());
        for comment in threads {
            let set = if comment.parent_id.is_none() {
                &mut roots
            } else {
                &mut replies
            };
            set.insert(comment.id.to_ascii_uppercase());
        }
        // The index of the first thread root on each cell, by `thread_cell_key`.
        let mut root_cells: HashMap<String, usize> = HashMap::new();
        for (i, comment) in threads.iter().enumerate() {
            if comment.parent_id.is_none() {
                match root_cells.entry(thread_cell_key(&comment.ref_cell)) {
                    Entry::Vacant(slot) => {
                        slot.insert(i);
                    }
                    Entry::Occupied(_) if report.omit_if_full(Severity::Error) => {}
                    Entry::Occupied(first) => report.push(issue(
                        format!(
                            "Threaded comment is a second thread root on cell {} (the first is threadedComments[{}]): Excel shows only the first thread on a cell, and this one is lost",
                            first.key(),
                            first.get()
                        ),
                        format!("{}!threadedComments[{i}]", sheet.name),
                        "Make it a reply: set parentId to the id of the first thread root on the cell",
                    )),
                }
            }
            if let Some(units) = thread_text_overflow(&comment.text)
                && !report.omit_if_full(Severity::Error)
            {
                report.push(issue(
                    format!(
                        "Threaded comment text is {units} UTF-16 code units long: Excel refuses a file with a threaded comment longer than {MAX_THREADED_COMMENT_TEXT} UTF-16 code units"
                    ),
                    format!("{}!threadedComments[{i}]", sheet.name),
                    "Shorten the text to at most 32767 UTF-16 code units (an astral character such as an emoji counts as two)",
                ));
            }
            let fields = [
                ("Threaded comment id", Some(&comment.id)),
                ("Threaded comment personId", Some(&comment.person_id)),
                ("Threaded comment parentId", comment.parent_id.as_ref()),
            ];
            for (field, id) in fields {
                if let Some(id) = id
                    && !threaded_comments::is_braced_guid(id)
                    && !report.omit_if_full(Severity::Error)
                {
                    report.push(issue(
                        thread_id_message(field, id),
                        format!("{}!threadedComments[{i}]", sheet.name),
                        GUID_FORM,
                    ));
                }
            }
            if let Some(parent) = &comment.parent_id {
                let parent_upper = parent.to_ascii_uppercase();
                if replies.contains(&parent_upper)
                    && !roots.contains(&parent_upper)
                    && !report.omit_if_full(Severity::Error)
                {
                    report.push(issue(
                        format!(
                            "Threaded comment parentId '{parent}' names a reply, not a thread root: Excel refuses a reply to a reply"
                        ),
                        format!("{}!threadedComments[{i}]", sheet.name),
                        "Set parentId to the id of the thread root (the comment without a parentId) the reply belongs to",
                    ));
                }
            }
        }
    }
    for (i, person) in wb.persons.iter().enumerate() {
        if !threaded_comments::is_braced_guid(&person.id) && !report.omit_if_full(Severity::Error) {
            report.push(issue(
                thread_id_message("Person id", &person.id),
                format!("persons[{i}]"),
                GUID_FORM,
            ));
        }
    }
}

// -- Sheet protection --

/// The sheets whose `sheetProtection.password` the writer would emit
/// although it is not a legacy hash
/// ([`SheetProtection::has_non_hash_password`]).
fn sheets_with_non_hash_password(wb: &WorkbookData) -> impl Iterator<Item = &SheetData> {
    wb.sheets.iter().filter(|sheet| {
        sheet
            .worksheet
            .sheet_protection
            .as_ref()
            .is_some_and(SheetProtection::has_non_hash_password)
    })
}

fn validate_sheet_protection(report: &mut ValidationReport, wb: &WorkbookData) {
    for sheet in sheets_with_non_hash_password(wb) {
        report.push(ValidationIssue {
            severity: Severity::Error,
            category: IssueCategory::SheetProtection,
            message: non_hash_password_message(Some(&sheet.name)),
            location: format!("{}!sheetProtection", sheet.name),
            suggestion: "Move the password to plainPassword and clear password, or set password to the legacy hash a file stored (1 to 4 hex digits)".into(),
            auto_fixable: false,
        });
    }
}

// -- Conditional-format formula count --

/// The conditional-format rules with more `<formula>` children than
/// `CT_CfRule` allows ([`MAX_CF_RULE_FORMULAS`]), with their sheet and block.
fn oversized_cf_rules(
    wb: &WorkbookData,
) -> impl Iterator<
    Item = (
        &SheetData,
        &ConditionalFormatting,
        &ConditionalFormattingRule,
    ),
> {
    wb.sheets.iter().flat_map(|sheet| {
        sheet
            .worksheet
            .conditional_formatting
            .iter()
            .flat_map(move |cf| {
                cf.rules
                    .iter()
                    .filter(|rule| rule.formula_count() > MAX_CF_RULE_FORMULAS)
                    .map(move |rule| (sheet, cf, rule))
            })
    })
}

fn validate_cf_formula_counts(report: &mut ValidationReport, wb: &WorkbookData) {
    for (sheet, cf, rule) in oversized_cf_rules(wb) {
        if report.omit_if_full(Severity::Error) {
            continue;
        }
        report.push(ValidationIssue {
            severity: Severity::Error,
            category: IssueCategory::ConditionalFormatting,
            message: cf_formula_count_message(
                Some(&sheet.name),
                &cf.sqref,
                rule.priority,
                rule.formula_count(),
            ),
            location: format!("{}!{}", sheet.name, cf.sqref),
            suggestion:
                "Keep at most three formulas on a rule: formula plus up to two additionalFormulas"
                    .into(),
            auto_fixable: false,
        });
    }
}

// -- Colours --

/// A colour the writer cannot emit: where it sits, its issue category and
/// the `rgb` it would write.
struct BadColor<'a> {
    at: ColorAt<'a>,
    category: IssueCategory,
    rgb: &'a str,
}

/// A style record a colour sits in: `styles.{list}[{index}]{member}`, with
/// `member` empty or, in a dxf, `.font`, `.fill` or `.border`.
#[derive(Clone, Copy)]
struct StyleAt {
    list: &'static str,
    index: usize,
    member: &'static str,
}

/// Where a colour sits, held in parts: [`ColorAt::location`] formats it only
/// for a bad colour the report stores or the write error names.
#[derive(Clone, Copy)]
enum ColorAt<'a> {
    /// A colour field of a style record: `field` is `.color`, `.fgColor`,
    /// `.bgColor` or `.<side>.color`.
    Style {
        at: StyleAt,
        field: &'static str,
    },
    /// A gradient stop of a style fill.
    GradientStop {
        at: StyleAt,
        stop: usize,
    },
    TabColor {
        sheet: &'a str,
    },
    ColorScale {
        sheet: &'a str,
        sqref: &'a str,
        rule: usize,
        color: usize,
    },
    DataBar {
        sheet: &'a str,
        sqref: &'a str,
        rule: usize,
    },
    /// A sparkline group colour; `field` is its JSON name (`colorSeries`).
    Sparkline {
        sheet: &'a str,
        group: usize,
        field: &'static str,
    },
    RichText {
        sheet: &'a str,
        cell: &'a str,
        run: usize,
    },
}

impl ColorAt<'_> {
    /// The location string of the report issue and the write error.
    #[cold]
    #[inline(never)]
    fn location(self) -> String {
        match self {
            Self::Style { at, field } => {
                format!("styles.{}[{}]{}{field}", at.list, at.index, at.member)
            }
            Self::GradientStop { at, stop } => format!(
                "styles.{}[{}]{}.gradientFill.stops[{stop}].color",
                at.list, at.index, at.member
            ),
            Self::TabColor { sheet } => format!("{sheet}!tabColor"),
            Self::ColorScale {
                sheet,
                sqref,
                rule,
                color,
            } => format!("{sheet}!{sqref} cfRule[{rule}].colorScale.colors[{color}]"),
            Self::DataBar { sheet, sqref, rule } => {
                format!("{sheet}!{sqref} cfRule[{rule}].dataBar.color")
            }
            Self::Sparkline {
                sheet,
                group,
                field,
            } => format!("{sheet}!sparklineGroups[{group}].{field}"),
            Self::RichText { sheet, cell, run } => {
                format!("{sheet}!{cell} richText[{run}].color")
            }
        }
    }
}

/// Collects the colours whose written `rgb` is not 6 or 8 hex digits.
#[derive(Default)]
struct ColorCheck<'a> {
    found: Vec<BadColor<'a>>,
}

impl<'a> ColorCheck<'a> {
    /// Check the colour the writer picks from a legacy rgb string and its
    /// full colour ([`resolve_color`]), and keep a bad one with where it sits.
    fn check(
        &mut self,
        category: IssueCategory,
        legacy: Option<&'a str>,
        spec: Option<&'a Color>,
        at: ColorAt<'a>,
    ) {
        if let Some(rgb) = resolve_color(legacy, spec).and_then(ColorOut::rgb)
            && !is_argb(rgb)
        {
            self.found.push(BadColor { at, category, rgb });
        }
    }

    fn font(&mut self, font: &'a Font, at: StyleAt) {
        self.check(
            IssueCategory::StyleIndex,
            font.color.as_deref(),
            font.color_spec.as_ref(),
            ColorAt::Style {
                at,
                field: ".color",
            },
        );
    }

    /// A pattern fill's `fgColor` and `bgColor`, or a gradient fill's stops
    /// (the writer emits one or the other).
    fn fill(&mut self, fill: &'a Fill, at: StyleAt) {
        let category = IssueCategory::StyleIndex;
        if let Some(ref gradient) = fill.gradient_fill {
            for (stop, gradient_stop) in gradient.stops.iter().enumerate() {
                self.check(
                    category,
                    Some(&gradient_stop.color),
                    gradient_stop.color_spec.as_ref(),
                    ColorAt::GradientStop { at, stop },
                );
            }
        } else {
            self.check(
                category,
                fill.fg_color.as_deref(),
                fill.fg_color_spec.as_ref(),
                ColorAt::Style {
                    at,
                    field: ".fgColor",
                },
            );
            self.check(
                category,
                fill.bg_color.as_deref(),
                fill.bg_color_spec.as_ref(),
                ColorAt::Style {
                    at,
                    field: ".bgColor",
                },
            );
        }
    }

    fn border(&mut self, border: &'a Border, at: StyleAt) {
        for (field, side) in [
            (".left.color", &border.left),
            (".right.color", &border.right),
            (".top.color", &border.top),
            (".bottom.color", &border.bottom),
            (".diagonal.color", &border.diagonal),
        ] {
            if let Some(side) = side {
                self.check(
                    IssueCategory::StyleIndex,
                    side.color.as_deref(),
                    side.color_spec.as_ref(),
                    ColorAt::Style { at, field },
                );
            }
        }
    }

    /// The colours of one sheet: tab colour, conditional-format colour scales
    /// and data bars, sparkline group colours, and the rich-text runs of its
    /// shared-string cells.
    fn sheet(&mut self, sheet: &'a SheetData) {
        let name = sheet.name.as_str();
        let ws = &sheet.worksheet;
        self.check(
            IssueCategory::Structure,
            ws.tab_color.as_deref(),
            ws.tab_color_spec.as_ref(),
            ColorAt::TabColor { sheet: name },
        );
        for cf in &ws.conditional_formatting {
            let sqref = cf.sqref.as_str();
            for (r, rule) in cf.rules.iter().enumerate() {
                let category = IssueCategory::ConditionalFormatting;
                if let Some(ref cs) = rule.color_scale {
                    for i in 0..cs.colors.len().max(cs.color_specs.len()) {
                        self.check(
                            category,
                            cs.colors.get(i).map(String::as_str),
                            cs.color_specs.get(i),
                            ColorAt::ColorScale {
                                sheet: name,
                                sqref,
                                rule: r,
                                color: i,
                            },
                        );
                    }
                }
                if let Some(ref db) = rule.data_bar {
                    self.check(
                        category,
                        Some(&db.color),
                        db.color_spec.as_ref(),
                        ColorAt::DataBar {
                            sheet: name,
                            sqref,
                            rule: r,
                        },
                    );
                }
            }
        }
        for (g, group) in ws.sparkline_groups.iter().enumerate() {
            for (field, rgb, spec) in [
                ("colorSeries", &group.color_series, &group.color_series_spec),
                (
                    "colorNegative",
                    &group.color_negative,
                    &group.color_negative_spec,
                ),
                ("colorAxis", &group.color_axis, &group.color_axis_spec),
                (
                    "colorMarkers",
                    &group.color_markers,
                    &group.color_markers_spec,
                ),
                ("colorFirst", &group.color_first, &group.color_first_spec),
                ("colorLast", &group.color_last, &group.color_last_spec),
                ("colorHigh", &group.color_high, &group.color_high_spec),
                ("colorLow", &group.color_low, &group.color_low_spec),
            ] {
                self.check(
                    IssueCategory::Structure,
                    rgb.as_deref(),
                    spec.as_ref(),
                    ColorAt::Sparkline {
                        sheet: name,
                        group: g,
                        field,
                    },
                );
            }
        }
        for row in &ws.rows {
            for cell in &row.cells {
                let Some(ref runs) = cell.rich_text else {
                    continue;
                };
                if cell.cell_type != CellType::SharedString || cell.value.is_none() {
                    continue;
                }
                for (k, run) in runs.iter().enumerate() {
                    self.check(
                        IssueCategory::Structure,
                        run.color.as_deref(),
                        run.color_spec.as_ref(),
                        ColorAt::RichText {
                            sheet: name,
                            cell: &cell.reference,
                            run: k,
                        },
                    );
                }
            }
        }
    }
}

/// Every colour of `wb` that the writer would emit with an `rgb` that is not
/// 6 or 8 hex digits ([`is_argb`]): such a string cannot be a colour, and the
/// write fails on it.
fn bad_colors(wb: &WorkbookData) -> Vec<BadColor<'_>> {
    let mut check = ColorCheck::default();
    let styles = &wb.styles;
    let at = |list, index, member| StyleAt {
        list,
        index,
        member,
    };
    for (i, font) in styles.fonts.iter().enumerate() {
        check.font(font, at("fonts", i, ""));
    }
    for (i, fill) in styles.fills.iter().enumerate() {
        check.fill(fill, at("fills", i, ""));
    }
    for (i, border) in styles.borders.iter().enumerate() {
        check.border(border, at("borders", i, ""));
    }
    for (i, dxf) in styles.dxfs.iter().enumerate() {
        if let Some(ref font) = dxf.font {
            check.font(font, at("dxfs", i, ".font"));
        }
        if let Some(ref fill) = dxf.fill {
            check.fill(fill, at("dxfs", i, ".fill"));
        }
        if let Some(ref border) = dxf.border {
            check.border(border, at("dxfs", i, ".border"));
        }
    }
    for sheet in &wb.sheets {
        check.sheet(sheet);
    }
    check.found
}

fn validate_colors(report: &mut ValidationReport, wb: &WorkbookData) {
    for bad in bad_colors(wb) {
        if report.omit_if_full(Severity::Error) {
            continue;
        }
        report.push(ValidationIssue {
            severity: Severity::Error,
            category: bad.category,
            message: not_argb_message(bad.rgb, None),
            location: bad.at.location(),
            suggestion: "Set the colour to RRGGBB or AARRGGBB hex digits, such as FF0000 or FFFF0000".into(),
            auto_fixable: false,
        });
    }
}

// -- Chart series order --

/// A series of a chart: its index in the chart's `series`, or with
/// `secondary` in its secondary chart's `series`.
#[derive(Clone, Copy)]
struct SeriesAt {
    secondary: bool,
    index: usize,
}

/// Every series of a chart with its `c:order`: the primary chart's series,
/// then the secondary chart's, in document order.
fn series_orders(chart: &ChartData) -> impl Iterator<Item = (SeriesAt, u32)> {
    let primary = chart.series.iter().enumerate().map(|(index, s)| {
        let at = SeriesAt {
            secondary: false,
            index,
        };
        (at, s.order)
    });
    let secondary = chart.secondary_chart.iter().flat_map(|c| {
        c.series.iter().enumerate().map(|(index, s)| {
            let at = SeriesAt {
                secondary: true,
                index,
            };
            (at, s.order)
        })
    });
    primary.chain(secondary)
}

/// The first two series of `chart` that share a `c:order`, with that order.
/// Excel refuses such a chart, whether the two share one chart-type element
/// or one sits in the primary and one in the secondary chart. Linear in the
/// number of series.
fn duplicate_series_order(chart: &ChartData) -> Option<(SeriesAt, SeriesAt, u32)> {
    let mut seen = HashSet::new();
    let (second, order) = series_orders(chart).find(|&(_, order)| !seen.insert(order))?;
    let (first, _) = series_orders(chart).find(|&(_, o)| o == order)?;
    Some((first, second, order))
}

/// The charts written from the model (no `raw_xml`) whose series share a
/// `c:order` ([`duplicate_series_order`]), with their sheet and position in
/// the sheet's `charts`.
fn charts_with_duplicate_series_order(
    wb: &WorkbookData,
) -> impl Iterator<
    Item = (
        &SheetData,
        usize,
        &WorksheetChart,
        (SeriesAt, SeriesAt, u32),
    ),
> {
    wb.sheets.iter().flat_map(|sheet| {
        sheet
            .worksheet
            .charts
            .iter()
            .enumerate()
            .filter(|(_, c)| c.chart.raw_xml.is_none())
            .filter_map(move |(i, c)| {
                duplicate_series_order(&c.chart).map(|dup| (sheet, i, c, dup))
            })
    })
}

/// The message for two series of a chart that share a `c:order`, naming the
/// sheet, the chart's position in `charts` and its title, and both series.
fn duplicate_series_order_message(
    sheet: &str,
    position: usize,
    chart: &WorksheetChart,
    (first, second, order): (SeriesAt, SeriesAt, u32),
) -> String {
    let title = chart
        .chart
        .title
        .as_ref()
        .map_or_else(String::new, |t| format!(" ({:?})", t.text));
    let name = |at: SeriesAt| {
        let owner = if at.secondary { "secondaryChart." } else { "" };
        format!("{owner}series[{}]", at.index)
    };
    format!(
        "the chart at charts[{position}] of sheet '{sheet}'{title}: {} and {} both have c:order {order}, which must be unique across the chart's series, including its secondary chart's",
        name(first),
        name(second)
    )
}

fn validate_chart_series_orders(report: &mut ValidationReport, wb: &WorkbookData) {
    for (sheet, i, chart, dup) in charts_with_duplicate_series_order(wb) {
        report.push(ValidationIssue {
            severity: Severity::Error,
            category: IssueCategory::Structure,
            message: duplicate_series_order_message(&sheet.name, i, chart, dup),
            location: format!("{}!charts[{i}]", sheet.name),
            suggestion: "Give every series of the chart and of its secondaryChart its own order, such as 0, 1, 2 across both".into(),
            auto_fixable: false,
        });
    }
}

// -- Pivot tables --

/// Reports each pivot table whose `cacheId` names no pivot cache: a cache is
/// listed under its own `cacheId`, else its index in `pivotCaches`, as the
/// writer lists it. The write of such a workbook fails (`INVALID_WORKBOOK`).
/// Shared items are not required: the writer completes an axis field from
/// its cache field's shared items when there are any, and Excel refreshes a
/// field whose cache field has none.
fn validate_pivot_cache_ids(report: &mut ValidationReport, wb: &WorkbookData) {
    let ids: HashSet<u32> = wb
        .pivot_caches
        .iter()
        .enumerate()
        .map(|(i, cache)| {
            cache
                .cache_id
                .unwrap_or_else(|| u32::try_from(i).unwrap_or(u32::MAX))
        })
        .collect();
    for sheet in &wb.sheets {
        for (i, table) in sheet.worksheet.pivot_tables.iter().enumerate() {
            if ids.contains(&table.cache_id) {
                continue;
            }
            report.push(ValidationIssue {
                severity: Severity::Warning,
                category: IssueCategory::Structure,
                message: format!(
                    "Pivot table '{}' has cacheId {}, which no pivot cache has: the write fails",
                    table.name, table.cache_id
                ),
                location: format!("{}!pivotTables[{i}]", sheet.name),
                suggestion: "Add its cache with addPivotCache, or set the table's cacheId to a cache's cacheId (its index in pivotCaches when that is unset)".into(),
                auto_fixable: false,
            });
        }
    }
}

/// Reports the first pivot table the write fails on (`SIZE_LIMIT`) because
/// of what the writer adds to it from the cache its `cacheId` names (see
/// `PivotTableData::count_added_items`): a row, column or page field without
/// items whose cache field has more than
/// [`MAX_FIELD_ITEMS`](crate::ooxml::pivot_table::MAX_FIELD_ITEMS) shared
/// items, or more than
/// [`MAX_ADDED_ITEMS`](crate::ooxml::pivot_table::MAX_ADDED_ITEMS) completed
/// items and appended pivot fields
/// across the workbook. Items and pivot fields the model gives are not
/// counted.
fn validate_pivot_additions(report: &mut ValidationReport, wb: &WorkbookData) {
    let ids: Vec<u32> = wb
        .pivot_caches
        .iter()
        .enumerate()
        .map(|(i, cache)| {
            cache
                .cache_id
                .unwrap_or_else(|| u32::try_from(i).unwrap_or(u32::MAX))
        })
        .collect();
    let mut added = 0;
    for sheet in &wb.sheets {
        for (i, table) in sheet.worksheet.pivot_tables.iter().enumerate() {
            let cache = ids
                .iter()
                .position(|&id| id == table.cache_id)
                .and_then(|c| wb.pivot_caches.get(c));
            if let Err(crate::ModernXlsxError::SizeLimit(message)) =
                table.count_added_items(cache, &mut added)
            {
                report.push(ValidationIssue {
                    severity: Severity::Warning,
                    category: IssueCategory::Structure,
                    message,
                    location: format!("{}!pivotTables[{i}]", sheet.name),
                    suggestion: "Give the table's row, column and page fields explicit items, and the table one pivot field per cache field: what the model gives is written as given".into(),
                    auto_fixable: false,
                });
                return;
            }
        }
    }
}

// -- Write-time checks --

/// The write-time form of the [`validate_workbook`] errors that the writer
/// cannot write around: a sheet protection `password` that is not a legacy
/// hash, a conditional-format rule with more than three formulas, a
/// colour whose `rgb` is not 6 or 8 hex digits, and a chart two of whose
/// series share a `c:order` (within one chart-type element or across its
/// primary and secondary chart).
///
/// # Errors
///
/// `InvalidWorkbook` for the first such value, naming its sheet (or, for a
/// colour, its location), with the message the report gives.
pub(crate) fn check_writable(wb: &WorkbookData) -> crate::Result<()> {
    let why = if let Some(sheet) = sheets_with_non_hash_password(wb).next() {
        non_hash_password_message(Some(&sheet.name))
    } else if let Some((sheet, cf, rule)) = oversized_cf_rules(wb).next() {
        cf_formula_count_message(
            Some(&sheet.name),
            &cf.sqref,
            rule.priority,
            rule.formula_count(),
        )
    } else if let Some(bad) = bad_colors(wb).first() {
        not_argb_message(bad.rgb, Some(&bad.at.location()))
    } else if let Some((sheet, i, chart, dup)) = charts_with_duplicate_series_order(wb).next() {
        duplicate_series_order_message(&sheet.name, i, chart, dup)
    } else {
        return Ok(());
    };
    cold_path();
    Err(crate::ModernXlsxError::InvalidWorkbook(format!(
        "Write failed: {why}"
    )))
}

// ---------------------------------------------------------------------------
// Repair
// ---------------------------------------------------------------------------

/// Auto-repair repairable issues in a workbook.
///
/// Returns the number of repairs applied. The workbook is modified in-place.
pub fn repair_workbook(wb: &mut WorkbookData) -> usize {
    let mut repairs = 0;

    repairs += repair_sheet_names(wb);
    repairs += repair_styles(wb);
    repairs += repair_cell_style_indices(wb);
    repairs += repair_shared_string_cells(wb);
    repairs += repair_theme(wb);
    repairs += repair_metadata(wb);
    repairs += repair_row_ordering(wb);
    repairs += repair_merge_ranges(wb);

    repairs
}

/// A name that passes [`sheet_name_problem`] (apart from duplicates), made from
/// `name` in a fixed order so that no step undoes an earlier one: drop
/// forbidden and control characters, truncate to 31 characters, trim
/// apostrophes, fall back to `Sheet{i+1}` when nothing is left, and move off
/// the reserved name.
///
/// Repair renames only the sheet; formulas and defined names that refer to the
/// old name by text are not rewritten.
fn sanitized_sheet_name(name: &str, i: usize) -> String {
    let kept: String = name
        .chars()
        .filter(|c| !FORBIDDEN_SHEET_CHARS.contains(c) && !c.is_control())
        .take(MAX_SHEET_NAME_LEN)
        .collect();
    let trimmed = kept.trim_matches('\'');
    if trimmed.is_empty() {
        format!("Sheet{}", i + 1)
    } else if trimmed.eq_ignore_ascii_case(RESERVED_SHEET_NAME) {
        format!("{RESERVED_SHEET_NAME}_1")
    } else {
        trimmed.to_owned()
    }
}

fn repair_sheet_names(wb: &mut WorkbookData) -> usize {
    let mut repairs = 0;

    for (i, sheet) in wb.sheets.iter_mut().enumerate() {
        if sheet_name_problem(&sheet.name).is_some() {
            sheet.name = sanitized_sheet_name(&sheet.name, i);
            repairs += 1;
        }
    }

    // Deduplicate case-insensitively; the suffix is kept whole and the stem is
    // shortened so that the result still fits in 31 characters.
    // `next` remembers, per (lower-cased stem, suffix length), where the suffix
    // search stopped, so N sheets sharing one name cost O(N) probes instead of
    // O(N^2). Keying by the truncated stem is exact: the `_digits` suffix has no
    // cased letter, so equal keys give equal candidates for every `n`.
    let mut seen: HashSet<String> = HashSet::with_capacity(wb.sheets.len());
    let mut next: HashMap<(String, usize), u32> = HashMap::new();
    for sheet in &mut wb.sheets {
        if seen.insert(sheet.name.to_lowercase()) {
            continue;
        }
        let mut n = 2;
        loop {
            let suffix = format!("_{n}");
            let stem: String = sheet
                .name
                .chars()
                .take(MAX_SHEET_NAME_LEN - suffix.len())
                .collect();
            let memo = (stem.to_lowercase(), suffix.len());
            if let Some(&m) = next.get(&memo)
                && m > n
            {
                n = m;
                continue;
            }
            let candidate = format!("{stem}{suffix}");
            if seen.insert(candidate.to_lowercase()) {
                sheet.name = candidate;
                next.insert(memo, n + 1);
                repairs += 1;
                break;
            }
            n += 1;
        }
    }

    repairs
}

fn repair_styles(wb: &mut WorkbookData) -> usize {
    let mut repairs = 0;
    let styles = &mut wb.styles;

    // Ensure at least one font.
    if styles.fonts.is_empty() {
        styles.fonts.push(Font {
            name: Some("Calibri".to_owned()),
            size: Some(11.0),
            ..Font::default()
        });
        repairs += 1;
    }

    // Ensure at least 2 fills (none + gray125).
    if styles.fills.len() < 2 {
        if styles.fills.is_empty() {
            styles.fills.push(Fill {
                pattern_type: "none".to_owned(),
                ..Fill::default()
            });
        }
        styles.fills.push(Fill {
            pattern_type: "gray125".to_owned(),
            ..Fill::default()
        });
        repairs += 1;
    }

    // Ensure at least one border.
    if styles.borders.is_empty() {
        styles.borders.push(Border::default());
        repairs += 1;
    }

    // Ensure at least one cellXf.
    if styles.cell_xfs.is_empty() {
        styles.cell_xfs.push(CellXf::default());
        repairs += 1;
    }

    // Clamp dangling indices in cellXfs.
    let font_max = (styles.fonts.len() - 1) as u32;
    let fill_max = (styles.fills.len() - 1) as u32;
    let border_max = (styles.borders.len() - 1) as u32;

    for xf in &mut styles.cell_xfs {
        if xf.font_id > font_max {
            xf.font_id = 0;
            repairs += 1;
        }
        if xf.fill_id > fill_max {
            xf.fill_id = 0;
            repairs += 1;
        }
        if xf.border_id > border_max {
            xf.border_id = 0;
            repairs += 1;
        }
    }

    repairs
}

fn repair_cell_style_indices(wb: &mut WorkbookData) -> usize {
    let mut repairs = 0;
    let xf_count = wb.styles.cell_xfs.len();

    for sheet in &mut wb.sheets {
        for row in &mut sheet.worksheet.rows {
            for cell in &mut row.cells {
                if let Some(si) = cell.style_index
                    && si as usize >= xf_count
                {
                    cell.style_index = Some(0);
                    repairs += 1;
                }
            }
        }
    }

    repairs
}

fn repair_shared_string_cells(wb: &mut WorkbookData) -> usize {
    let mut repairs = 0;

    for sheet in &mut wb.sheets {
        for row in &mut sheet.worksheet.rows {
            for cell in &mut row.cells {
                if cell.cell_type == CellType::SharedString && cell.value.is_none() {
                    cell.value = Some(String::new());
                    repairs += 1;
                }
            }
        }
    }

    repairs
}

fn repair_theme(wb: &mut WorkbookData) -> usize {
    if wb.theme_colors.is_none() {
        wb.theme_colors = Some(ThemeColors::default());
        return 1;
    }
    0
}

fn repair_metadata(wb: &mut WorkbookData) -> usize {
    let mut repairs = 0;

    if let Some(props) = &mut wb.doc_properties {
        // Normalise a ` ` / `t` date-time separator and a lowercase `z`, then
        // classify: a calendar-valid date-time without a time-zone designator
        // gets `Z` (UTC is assumed) so the value survives; only a value that
        // is not calendar-valid is cleared.
        for date_field in [&mut props.created, &mut props.modified] {
            let Some(original) = date_field.as_deref() else {
                continue;
            };
            let mut fixed = original.to_owned();
            if fixed.len() > 10
                && fixed.is_char_boundary(11)
                && matches!(fixed.as_bytes()[10], b' ' | b't')
            {
                fixed.replace_range(10..11, "T");
            }
            if fixed.ends_with('z') {
                fixed.replace_range(fixed.len() - 1.., "Z");
            }
            match w3cdtf_status(&fixed) {
                W3cdtf::Valid => {}
                W3cdtf::MissingZone => fixed.push('Z'),
                W3cdtf::Invalid => {
                    *date_field = None;
                    repairs += 1;
                    continue;
                }
            }
            if fixed != original {
                *date_field = Some(fixed);
                repairs += 1;
            }
        }

        // Clear whitespace-only string fields.
        for (_, field) in doc_text_fields!(props, &mut) {
            if field.as_deref().is_some_and(is_blank) {
                *field = None;
                repairs += 1;
            }
        }
    }

    repairs
}

/// Sort rows by index and merge rows that share one: the first row's
/// attributes win, the cells are concatenated, the last cell for a reference
/// wins, and the cells end up in column order. One repair per changed sheet.
fn repair_row_ordering(wb: &mut WorkbookData) -> usize {
    let mut repairs = 0;

    for sheet in &mut wb.sheets {
        let rows = &mut sheet.worksheet.rows;
        if rows.is_sorted_by(|a, b| a.index < b.index) {
            continue;
        }
        rows.sort_by_key(|r| r.index);

        let mut merged: Vec<Row> = Vec::with_capacity(rows.len());
        for row in std::mem::take(rows) {
            match merged.last_mut() {
                Some(prev) if prev.index == row.index => prev.cells.extend(row.cells),
                _ => merged.push(row),
            }
        }
        for row in &mut merged {
            dedupe_cells(&mut row.cells);
        }
        *rows = merged;
        repairs += 1;
    }

    repairs
}

/// Keep the last cell for each reference and order the cells by column
/// (unparseable references last, in their original order).
fn dedupe_cells(cells: &mut Vec<Cell>) {
    let mut position: HashMap<String, usize> = HashMap::with_capacity(cells.len());
    let mut unique: Vec<Cell> = Vec::with_capacity(cells.len());
    for cell in std::mem::take(cells) {
        if let Some(&at) = position.get(&cell.reference) {
            unique[at] = cell;
        } else {
            position.insert(cell.reference.clone(), unique.len());
            unique.push(cell);
        }
    }
    unique.sort_by_cached_key(|c| CellRef::parse(&c.reference).map_or(u32::MAX, |r| r.col));
    *cells = unique;
}

/// Rewrite each reversed merge range (`B2:A1`) to its normalized form (`A1:B2`).
fn repair_merge_ranges(wb: &mut WorkbookData) -> usize {
    let mut repairs = 0;

    for sheet in &mut wb.sheets {
        for merge_ref in &mut sheet.worksheet.merge_cells {
            if let Some(Ok((start, end))) = parse_merge_corners(merge_ref)
                && is_reversed(start, end)
            {
                let (lo, hi) = normalized_corners(start, end);
                *merge_ref = format!("{}:{}", lo.to_a1(), hi.to_a1());
                repairs += 1;
            }
        }
    }

    repairs
}

// ---------------------------------------------------------------------------
// Tests
// ---------------------------------------------------------------------------

#[cfg(test)]
mod tests;

#[cfg(test)]
mod pivot_tests {
    use super::*;
    use crate::ooxml::pivot_table::{PivotCacheDefinitionData, PivotTableData};
    use pretty_assertions::assert_eq;

    /// Sheet `S` with one pivot table over `table_cache_id`, and one cache per
    /// entry of `cache_ids`.
    fn workbook(cache_ids: &[Option<u32>], table_cache_id: u32) -> WorkbookData {
        let table = PivotTableData::parse(
            br#"<pivotTableDefinition name="P" cacheId="0"><location ref="A1:B2"/></pivotTableDefinition>"#,
        )
        .unwrap();
        let cache = PivotCacheDefinitionData::parse(
            br#"<pivotCacheDefinition><cacheSource type="worksheet"><worksheetSource ref="A1:B2" sheet="S"/></cacheSource></pivotCacheDefinition>"#,
        )
        .unwrap();
        WorkbookData {
            sheets: vec![SheetData {
                name: "S".into(),
                worksheet: crate::ooxml::worksheet::WorksheetXml {
                    pivot_tables: vec![PivotTableData {
                        cache_id: table_cache_id,
                        ..table
                    }],
                    ..Default::default()
                },
                ..Default::default()
            }],
            pivot_caches: cache_ids
                .iter()
                .map(|&cache_id| PivotCacheDefinitionData {
                    cache_id,
                    ..cache.clone()
                })
                .collect(),
            theme_colors: Some(ThemeColors::default()),
            ..Default::default()
        }
    }

    #[test]
    fn a_pivot_table_whose_cache_id_names_no_cache_is_a_structure_warning() {
        // (caches' cacheIds, the table's cacheId, reported)
        let cases: [(&[Option<u32>], u32, bool); 6] = [
            (&[Some(7)], 7, false),
            // A cache without an id is listed under its index.
            (&[None], 0, false),
            (&[Some(7), None], 1, false),
            (&[Some(7)], 0, true),
            (&[None], 4, true),
            (&[], 0, true),
        ];
        for (cache_ids, table_cache_id, reported) in cases {
            let report = validate_workbook(&workbook(cache_ids, table_cache_id));
            let label = format!("{cache_ids:?} / {table_cache_id}");
            let issues: Vec<_> = report
                .issues
                .iter()
                .filter(|issue| issue.location.contains("pivotTables"))
                .collect();
            assert_eq!(issues.len(), usize::from(reported), "{label}: {issues:?}");
            if let Some(issue) = issues.first() {
                assert_eq!(
                    (issue.severity, issue.category, issue.location.as_str()),
                    (
                        Severity::Warning,
                        IssueCategory::Structure,
                        "S!pivotTables[0]"
                    ),
                    "{label}"
                );
                assert!(
                    issue.message.contains(&format!("cacheId {table_cache_id}")),
                    "{label}: {}",
                    issue.message
                );
                assert!(!issue.auto_fixable, "{label}");
            }
        }
    }

    #[test]
    fn a_pivot_table_the_writer_would_add_too_much_to_is_a_structure_warning() {
        use crate::ooxml::pivot_table::{
            CacheFieldData, CacheValue, MAX_FIELD_ITEMS, PivotAxis, PivotFieldData, PivotItem,
        };

        /// `tables` tables on cacheId 0, each with one row field without
        /// items, over a cache field of `shared` shared items.
        fn completed(tables: usize, shared: usize) -> WorkbookData {
            let mut wb = workbook(&[Some(0)], 0);
            wb.pivot_caches[0].fields = vec![CacheFieldData {
                name: "F".into(),
                shared_items: vec![CacheValue::Missing; shared],
                ..CacheFieldData::default()
            }];
            let table = &mut wb.sheets[0].worksheet.pivot_tables[0];
            table.pivot_fields = vec![PivotFieldData {
                axis: Some(PivotAxis::AxisRow),
                name: None,
                items: Vec::new(),
                subtotals: Vec::new(),
                compact: true,
                outline: true,
                data_field: false,
                default_subtotal: None,
            }];
            let table = table.clone();
            wb.sheets[0].worksheet.pivot_tables = vec![table; tables];
            wb
        }
        let pivot_issues = |wb: &WorkbookData| -> Vec<ValidationIssue> {
            validate_workbook(wb)
                .issues
                .into_iter()
                .filter(|issue| issue.location.contains("pivotTables"))
                .collect()
        };

        assert!(pivot_issues(&completed(4, 1_000_000)).is_empty());
        // (label, workbook, the table reported, a number the message names)
        let cases = [
            (
                "one field past the cap",
                completed(1, MAX_FIELD_ITEMS + 1),
                0,
                "1048576",
            ),
            (
                "past the workbook ceiling",
                completed(5, 1_000_000),
                4,
                "4194304",
            ),
        ];
        for (label, wb, table, limit) in cases {
            let issues = pivot_issues(&wb);
            assert_eq!(issues.len(), 1, "{label}: {issues:?}");
            let issue = &issues[0];
            assert_eq!(
                (issue.severity, issue.category, issue.location.as_str()),
                (
                    Severity::Warning,
                    IssueCategory::Structure,
                    format!("S!pivotTables[{table}]").as_str()
                ),
                "{label}"
            );
            assert!(issue.message.contains(limit), "{label}: {}", issue.message);
            assert_eq!(
                crate::writer::write_xlsx(&wb).unwrap_err().code(),
                "SIZE_LIMIT",
                "{label}: the write fails"
            );
        }
        // Items the model gives are not counted.
        let mut wb = completed(1, MAX_FIELD_ITEMS + 1);
        wb.sheets[0].worksheet.pivot_tables[0].pivot_fields[0].items = vec![PivotItem {
            t: None,
            x: Some(0),
        }];
        assert!(pivot_issues(&wb).is_empty());
    }
}

#[cfg(test)]
mod thread_id_tests {
    use super::*;
    use crate::ooxml::threaded_comments::{PersonData, ThreadedCommentData};
    use pretty_assertions::assert_eq;

    const GOOD: &str = "{5FA6E5A7-6BD7-4105-A288-53EB03A03DCF}";

    /// A sheet `S` with one thread (root on A1 and a reply) and its person,
    /// with `edit` applied to the root, the reply and the person.
    fn threaded(
        edit: impl FnOnce(&mut ThreadedCommentData, &mut ThreadedCommentData, &mut PersonData),
    ) -> WorkbookData {
        let comment = |id: &str, parent: Option<&str>| ThreadedCommentData {
            id: id.into(),
            ref_cell: "A1".into(),
            person_id: GOOD.into(),
            text: "t".into(),
            timestamp: "2026-11-12T00:00:00.00".into(),
            parent_id: parent.map(Into::into),
            done: None,
            mentions: Vec::new(),
        };
        let mut root = comment("{87D66092-362F-4C38-9C8D-1611ABA57577}", None);
        let mut reply = comment(
            "{87D66092-362F-4C39-9C8D-1611ABA57577}",
            Some("{87D66092-362F-4C38-9C8D-1611ABA57577}"),
        );
        let mut person = PersonData {
            id: GOOD.into(),
            display_name: "Ann".into(),
            user_id: None,
            provider_id: None,
        };
        edit(&mut root, &mut reply, &mut person);
        WorkbookData {
            sheets: vec![SheetData {
                name: "S".into(),
                worksheet: crate::ooxml::worksheet::WorksheetXml {
                    threaded_comments: vec![root, reply],
                    ..Default::default()
                },
                ..Default::default()
            }],
            persons: vec![person],
            theme_colors: Some(ThemeColors::default()),
            ..Default::default()
        }
    }

    /// The location of each issue of `wb`.
    fn issue_locations(wb: &WorkbookData) -> Vec<String> {
        validate_workbook(wb)
            .issues
            .into_iter()
            .map(|issue| issue.location)
            .collect()
    }

    #[test]
    fn excel_form_ids_in_either_case_are_valid() {
        // Excel 16.0 opens braced GUIDs in upper or lower case.
        let lower = GOOD.to_ascii_lowercase();
        let wb = threaded(|root, reply, person| {
            root.person_id.clone_from(&lower);
            reply.person_id.clone_from(&lower);
            person.id.clone_from(&lower);
        });
        assert_eq!(issue_locations(&wb), Vec::<String>::new());
        assert_eq!(
            issue_locations(&threaded(|_, _, _| {})),
            Vec::<String>::new()
        );
    }

    #[test]
    fn ids_excel_refuses_are_structure_errors() {
        const BARE: &str = "5fa6e5a7-6bd7-4105-a288-53eb03a03dcf";
        type Edit = fn(&mut ThreadedCommentData, &mut ThreadedCommentData, &mut PersonData);
        const FORM: &str = "{XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX}";
        let cases: [(&str, Edit, &str, &str); 8] = [
            (
                "root id",
                |root, _, _| root.id = BARE.into(),
                "S!threadedComments[0]",
                FORM,
            ),
            (
                "reply personId",
                |_, reply, _| reply.person_id = BARE.into(),
                "S!threadedComments[1]",
                FORM,
            ),
            (
                "reply parentId",
                |_, reply, _| reply.parent_id = Some(BARE.into()),
                "S!threadedComments[1]",
                FORM,
            ),
            (
                "person id",
                |_, _, person| person.id = BARE.into(),
                "persons[0]",
                FORM,
            ),
            (
                "upper case without braces",
                |root, _, _| root.id = BARE.to_ascii_uppercase(),
                "S!threadedComments[0]",
                FORM,
            ),
            (
                "braces without a GUID",
                |root, _, _| root.id = "{1}".into(),
                "S!threadedComments[0]",
                FORM,
            ),
            (
                "a GUID with a non-hex digit",
                |root, _, _| root.id = "{5FA6E5A7-6BD7-4105-A288-53EB03A03DCG}".into(),
                "S!threadedComments[0]",
                FORM,
            ),
            (
                // The first comment becomes a reply (to a comment not on the
                // sheet), so the second, naming it in lower case, replies to
                // a reply.
                "parentId naming a reply, in another case",
                |root, reply, _| {
                    root.parent_id = Some(GOOD.into());
                    reply.parent_id = Some(root.id.to_ascii_lowercase());
                },
                "S!threadedComments[1]",
                "Excel refuses a reply to a reply",
            ),
        ];
        for (label, edit, location, message) in cases {
            let wb = threaded(edit);
            let found = validate_workbook(&wb);
            assert!(!found.is_valid, "{label}");
            assert_eq!(found.error_count, 1, "{label}: {:?}", found.issues);
            let issue = &found.issues[0];
            assert_eq!(
                (issue.severity, issue.category, issue.location.as_str()),
                (Severity::Error, IssueCategory::Structure, location),
                "{label}"
            );
            assert!(
                issue.message.contains(message),
                "{label}: {}",
                issue.message
            );
            assert!(!issue.auto_fixable, "{label}");
        }
    }

    /// The errors of `wb` as (location, message).
    fn errors(wb: &WorkbookData) -> Vec<(String, String)> {
        validate_workbook(wb)
            .issues
            .into_iter()
            .filter(|issue| issue.severity == Severity::Error)
            .inspect(|issue| assert_eq!(issue.category, IssueCategory::Structure))
            .map(|issue| (issue.location, issue.message))
            .collect()
    }

    #[test]
    fn two_thread_roots_on_one_cell_are_an_error() {
        // Excel shows only the first thread on a cell.
        // The reply becomes a second root on A1, spelled as a raw model may.
        for spelling in ["A1", "a1", "$A$1", "$a1"] {
            let wb = threaded(|_, reply, _| {
                reply.parent_id = None;
                reply.ref_cell = spelling.into();
            });
            let found = errors(&wb);
            assert_eq!(found.len(), 1, "{spelling}: {found:?}");
            let (location, message) = &found[0];
            assert_eq!(location, "S!threadedComments[1]", "{spelling}");
            assert!(
                message.contains("A1") && message.contains("threadedComments[0]"),
                "{spelling}: {message}"
            );
        }
        // Roots on different cells, and replies sharing their root's cell, are clean.
        let wb = threaded(|_, reply, _| {
            reply.parent_id = None;
            reply.ref_cell = "B1".into();
        });
        assert_eq!(errors(&wb), Vec::new());
    }

    #[test]
    fn a_text_over_32767_utf16_code_units_is_an_error() {
        // Excel 16 opens 32,767 UTF-16 code units in a root or a reply and
        // refuses 32,768, counting an astral character as two (Excel 16.0 build 20527, probed 2026-11-12).
        const ASTRAL: &str = "\u{1F600}";
        let ascii_max = "a".repeat(MAX_THREADED_COMMENT_TEXT);
        let astral_max = format!("{}a", ASTRAL.repeat(16_383));
        let wb = threaded(|root, reply, _| {
            root.text.clone_from(&ascii_max);
            reply.text.clone_from(&astral_max);
        });
        assert_eq!(errors(&wb), Vec::new());

        let ascii_over = "a".repeat(MAX_THREADED_COMMENT_TEXT + 1);
        let astral_over = ASTRAL.repeat(16_384);
        let wb = threaded(|root, reply, _| {
            root.text.clone_from(&astral_over);
            reply.text.clone_from(&ascii_over);
        });
        let found = errors(&wb);
        assert_eq!(
            found.iter().map(|(l, _)| l.as_str()).collect::<Vec<_>>(),
            ["S!threadedComments[0]", "S!threadedComments[1]"]
        );
        for (_, message) in &found {
            assert!(
                message.contains("32768 UTF-16 code units") && message.contains("32767"),
                "{message}"
            );
        }
    }
}
