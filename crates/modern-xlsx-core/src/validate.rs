//! OOXML Validation & Compliance enforcement.
//!
//! Provides structural validation, auto-repair, color normalization,
//! metadata sanitization, and compliance reporting for workbooks.
//!
//! # Features
//!
//! - **Validation:** Detects structural issues (invalid cell refs, dangling
//!   style indices, merge overlaps, missing required parts).
//! - **Repair:** Auto-fixes repairable issues (clamps indices, injects
//!   missing default styles, fills missing theme colors).
//! - **Color normalization:** Resolves indexed/theme colors to 6-char RGB hex.
//! - **Metadata sanitization:** Cleans docProps, ensures valid UTF-8 dates.
//! - **Theme injection:** Generates a complete Office 2016 theme when absent.
//! - **Compliance report:** Returns a structured `ValidationReport` with
//!   actionable fix suggestions.

use std::collections::HashSet;

use serde::{Deserialize, Serialize};

use crate::ooxml::cell::CellRef;
use crate::ooxml::styles::{Border, CellXf, Fill, Font};
use crate::ooxml::theme::ThemeColors;
use crate::ooxml::worksheet::CellType;
use crate::WorkbookData;

// ---------------------------------------------------------------------------
// OOXML limits
// ---------------------------------------------------------------------------

/// Maximum column index (0-based): XFD = 16383.
const MAX_COL: u32 = 16_383;
/// Maximum row index (0-based): 1048576 rows → 0..1048575.
const MAX_ROW: u32 = 1_048_575;
/// Maximum sheet name length per OOXML spec.
const MAX_SHEET_NAME_LEN: usize = 31;
/// Characters forbidden in sheet names.
const FORBIDDEN_SHEET_CHARS: &[char] = &['\\', '/', '*', '?', ':', '[', ']'];

// ---------------------------------------------------------------------------
// Indexed color palette (legacy Excel 97-2003)
// ---------------------------------------------------------------------------

/// The 56 standard indexed colors used by Excel for color indices 8-63.
/// Index 0-7 are the same as 8-15 (duplicated for legacy reasons).
/// Indices 64 and 65 are special (System Foreground/Background).
const INDEXED_COLORS: &[&str; 56] = &[
    "000000", "FFFFFF", "FF0000", "00FF00", "0000FF", "FFFF00", "FF00FF", "00FFFF",
    "000000", "FFFFFF", "FF0000", "00FF00", "0000FF", "FFFF00", "FF00FF", "00FFFF",
    "800000", "008000", "000080", "808000", "800080", "008080", "C0C0C0", "808080",
    "9999FF", "993366", "FFFFCC", "CCFFFF", "660066", "FF8080", "0066CC", "CCCCFF",
    "000080", "FF00FF", "FFFF00", "00FFFF", "800080", "800000", "008080", "0000FF",
    "00CCFF", "CCFFFF", "CCFFCC", "FFFF99", "99CCFF", "FF99CC", "CC99FF", "FFCC99",
    "3366FF", "33CCCC", "99CC00", "FFCC00", "FF9900", "FF6600", "666699", "969696",
];

// ---------------------------------------------------------------------------
// Severity & Issue types
// ---------------------------------------------------------------------------

/// Severity level for a validation issue.
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum Severity {
    /// Informational finding — no action required.
    Info,
    /// Potential problem that may cause display differences.
    Warning,
    /// Structural error that will likely cause failures in Excel.
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
}

/// A single validation issue with an actionable suggestion.
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ValidationIssue {
    pub severity: Severity,
    pub category: IssueCategory,
    /// Human-readable description of the issue.
    pub message: String,
    /// Where the issue was found (e.g. "Sheet1!A1", "styles.cellXfs[3]").
    pub location: String,
    /// Suggested fix (e.g. "Clamp style index to 0").
    pub suggestion: String,
    /// Whether `repair_workbook` can auto-fix this issue.
    pub auto_fixable: bool,
}

/// Result of validating a workbook.
#[derive(Debug, Clone, Default, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ValidationReport {
    pub issues: Vec<ValidationIssue>,
    pub error_count: usize,
    pub warning_count: usize,
    pub info_count: usize,
    /// `true` if no errors were found.
    pub is_valid: bool,
}

impl ValidationReport {
    fn push(&mut self, issue: ValidationIssue) {
        match issue.severity {
            Severity::Error => self.error_count += 1,
            Severity::Warning => self.warning_count += 1,
            Severity::Info => self.info_count += 1,
        }
        self.issues.push(issue);
    }

    fn finalize(&mut self) {
        self.is_valid = self.error_count == 0;
        self.issues.sort_by(|a, b| b.severity.cmp(&a.severity));
    }
}

// ---------------------------------------------------------------------------
// Validate
// ---------------------------------------------------------------------------

/// Validate a workbook and return a report of all issues found.
///
/// This does **not** modify the workbook. Use [`repair_workbook`] to auto-fix
/// repairable issues.
pub fn validate_workbook(wb: &WorkbookData) -> ValidationReport {
    let mut report = ValidationReport::default();

    validate_sheets(&mut report, wb);
    validate_styles(&mut report, wb);
    validate_cells(&mut report, wb);
    validate_merge_cells(&mut report, wb);
    validate_defined_names(&mut report, wb);
    validate_data_validations(&mut report, wb);
    validate_theme(&mut report, wb);
    validate_metadata(&mut report, wb);
    validate_structure(&mut report, wb);

    report.finalize();
    report
}

// -- Sheet names --

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
        let loc = format!("sheets[{i}]");

        if sheet.name.is_empty() {
            report.push(ValidationIssue {
                severity: Severity::Error,
                category: IssueCategory::SheetName,
                message: "Sheet name is empty".into(),
                location: loc.clone(),
                suggestion: format!("Set sheet name to 'Sheet{}'", i + 1),
                auto_fixable: true,
            });
        }

        if sheet.name.chars().count() > MAX_SHEET_NAME_LEN {
            report.push(ValidationIssue {
                severity: Severity::Error,
                category: IssueCategory::SheetName,
                message: format!(
                    "Sheet name '{}' exceeds {} character limit",
                    sheet.name, MAX_SHEET_NAME_LEN
                ),
                location: loc.clone(),
                suggestion: format!(
                    "Truncate to '{}'",
                    sheet.name.chars().take(MAX_SHEET_NAME_LEN).collect::<String>()
                ),
                auto_fixable: true,
            });
        }

        if sheet.name.chars().any(|c| FORBIDDEN_SHEET_CHARS.contains(&c)) {
            report.push(ValidationIssue {
                severity: Severity::Error,
                category: IssueCategory::SheetName,
                message: format!(
                    "Sheet name '{}' contains forbidden characters",
                    sheet.name
                ),
                location: loc.clone(),
                suggestion: "Remove \\ / * ? : [ ] from sheet name".into(),
                auto_fixable: true,
            });
        }

        let lower = sheet.name.to_lowercase();
        if !seen_names.insert(lower) {
            report.push(ValidationIssue {
                severity: Severity::Error,
                category: IssueCategory::SheetName,
                message: format!("Duplicate sheet name '{}'", sheet.name),
                location: loc,
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
        let loc = format!("styles.cellXfs[{i}]");

        if xf.font_id as usize >= font_count {
            report.push(ValidationIssue {
                severity: Severity::Error,
                category: IssueCategory::StyleIndex,
                message: format!(
                    "cellXf[{i}].fontId={} exceeds fonts count ({font_count})",
                    xf.font_id
                ),
                location: loc.clone(),
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
                location: loc.clone(),
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
                location: loc,
                suggestion: "Clamp borderId to 0".into(),
                auto_fixable: true,
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
            suggestion: "Inject default Aptos 11pt font".into(),
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

// -- Cells --

fn validate_cells(report: &mut ValidationReport, wb: &WorkbookData) {
    let xf_count = wb.styles.cell_xfs.len();

    for sheet in &wb.sheets {
        for row in &sheet.worksheet.rows {
            for cell in &row.cells {
                let loc = format!("{}!{}", sheet.name, cell.reference);

                // Validate cell reference bounds.
                if let Ok(cr) = CellRef::parse(&cell.reference) {
                    if cr.col > MAX_COL {
                        report.push(ValidationIssue {
                            severity: Severity::Error,
                            category: IssueCategory::CellReference,
                            message: format!(
                                "Cell {} column {} exceeds max ({})",
                                cell.reference, cr.col, MAX_COL
                            ),
                            location: loc.clone(),
                            suggestion: "Remove or relocate cell".into(),
                            auto_fixable: false,
                        });
                    }
                    if cr.row > MAX_ROW {
                        report.push(ValidationIssue {
                            severity: Severity::Error,
                            category: IssueCategory::CellReference,
                            message: format!(
                                "Cell {} row {} exceeds max ({})",
                                cell.reference, cr.row, MAX_ROW
                            ),
                            location: loc.clone(),
                            suggestion: "Remove or relocate cell".into(),
                            auto_fixable: false,
                        });
                    }
                } else {
                    report.push(ValidationIssue {
                        severity: Severity::Error,
                        category: IssueCategory::CellReference,
                        message: format!(
                            "Invalid cell reference '{}'",
                            cell.reference
                        ),
                        location: loc.clone(),
                        suggestion: "Fix cell reference format".into(),
                        auto_fixable: false,
                    });
                }

                // Validate style index.
                if let Some(si) = cell.style_index
                    && si as usize >= xf_count
                {
                    report.push(ValidationIssue {
                        severity: Severity::Error,
                        category: IssueCategory::StyleIndex,
                        message: format!(
                            "Cell {} styleIndex={si} exceeds cellXfs count ({xf_count})",
                            cell.reference
                        ),
                        location: loc.clone(),
                        suggestion: "Clamp styleIndex to 0".into(),
                        auto_fixable: true,
                    });
                }

                // SharedString cell must have a value.
                if cell.cell_type == CellType::SharedString && cell.value.is_none() {
                    report.push(ValidationIssue {
                        severity: Severity::Error,
                        category: IssueCategory::SharedString,
                        message: format!(
                            "SharedString cell {} has no value",
                            cell.reference
                        ),
                        location: loc,
                        suggestion: "Set value to an empty string or change cell type".into(),
                        auto_fixable: true,
                    });
                }
            }
        }
    }
}

// -- Merge cells --

fn validate_merge_cells(report: &mut ValidationReport, wb: &WorkbookData) {
    for sheet in &wb.sheets {
        let mut occupied: HashSet<(u32, u32)> = HashSet::new();

        for merge_ref in &sheet.worksheet.merge_cells {
            let loc = format!("{}!mergeCell({})", sheet.name, merge_ref);

            let Some((start_str, end_str)) = merge_ref.split_once(':') else {
                report.push(ValidationIssue {
                    severity: Severity::Error,
                    category: IssueCategory::MergeCell,
                    message: format!("Invalid merge range '{merge_ref}'"),
                    location: loc,
                    suggestion: "Use format 'A1:B2'".into(),
                    auto_fixable: false,
                });
                continue;
            };

            let (start, end) = match (CellRef::parse(start_str), CellRef::parse(end_str)) {
                (Ok(s), Ok(e)) => (s, e),
                _ => {
                    report.push(ValidationIssue {
                        severity: Severity::Error,
                        category: IssueCategory::MergeCell,
                        message: format!(
                            "Cannot parse merge range endpoints in '{merge_ref}'"
                        ),
                        location: loc,
                        suggestion: "Fix cell references".into(),
                        auto_fixable: false,
                    });
                    continue;
                }
            };

            // Check for overlapping merge regions.
            let mut overlap = false;
            for r in start.row..=end.row {
                for c in start.col..=end.col {
                    if !occupied.insert((r, c)) {
                        overlap = true;
                    }
                }
            }

            if overlap {
                report.push(ValidationIssue {
                    severity: Severity::Error,
                    category: IssueCategory::MergeCell,
                    message: format!(
                        "Merge range '{merge_ref}' overlaps with another merge region"
                    ),
                    location: loc,
                    suggestion: "Remove overlapping merge or adjust ranges".into(),
                    auto_fixable: false,
                });
            }
        }
    }
}

// -- Defined names --

fn validate_defined_names(report: &mut ValidationReport, wb: &WorkbookData) {
    for (i, dn) in wb.defined_names.iter().enumerate() {
        let loc = format!("definedNames[{i}]");

        if dn.name.is_empty() {
            report.push(ValidationIssue {
                severity: Severity::Error,
                category: IssueCategory::DefinedName,
                message: "Defined name has empty name".into(),
                location: loc.clone(),
                suggestion: "Set a valid name".into(),
                auto_fixable: false,
            });
        }

        if dn.value.is_empty() {
            report.push(ValidationIssue {
                severity: Severity::Warning,
                category: IssueCategory::DefinedName,
                message: format!("Defined name '{}' has empty value/reference", dn.name),
                location: loc,
                suggestion: "Set a valid cell range reference".into(),
                auto_fixable: false,
            });
        }
    }
}

// -- Data validations --

fn validate_data_validations(report: &mut ValidationReport, wb: &WorkbookData) {
    for sheet in &wb.sheets {
        for (i, dv) in sheet.worksheet.data_validations.iter().enumerate() {
            let loc = format!("{}!dataValidation[{i}]", sheet.name);

            if dv.sqref.is_empty() {
                report.push(ValidationIssue {
                    severity: Severity::Warning,
                    category: IssueCategory::DataValidation,
                    message: "Data validation has empty sqref".into(),
                    location: loc.clone(),
                    suggestion: "Set target cell range".into(),
                    auto_fixable: false,
                });
            }

            // List validations should have formula1.
            if dv.validation_type.as_deref() == Some("list") && dv.formula1.is_none() {
                report.push(ValidationIssue {
                    severity: Severity::Warning,
                    category: IssueCategory::DataValidation,
                    message: "List validation missing formula1 (source list)".into(),
                    location: loc,
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
            message: "No theme colors present; default Office 2016 theme will be used".into(),
            location: "themeColors".into(),
            suggestion: "No action needed — defaults are injected automatically".into(),
            auto_fixable: true,
        });
    }
}

// -- Metadata --

fn validate_metadata(report: &mut ValidationReport, wb: &WorkbookData) {
    if let Some(ref props) = wb.doc_properties {
        // Check for potentially invalid date formats.
        for (field_name, value) in [
            ("created", &props.created),
            ("modified", &props.modified),
        ] {
            if let Some(date_str) = value
                && !is_valid_iso8601(date_str)
            {
                report.push(ValidationIssue {
                    severity: Severity::Warning,
                    category: IssueCategory::Metadata,
                    message: format!(
                        "Document property '{field_name}' has non-ISO-8601 value '{date_str}'"
                    ),
                    location: format!("docProperties.{field_name}"),
                    suggestion: "Use ISO 8601 format: YYYY-MM-DDTHH:MM:SSZ".into(),
                    auto_fixable: true,
                });
            }
        }

        // Check for non-UTF-8 safe strings (shouldn't happen in Rust, but check for empty/whitespace-only).
        for (field_name, value) in [
            ("title", &props.title),
            ("creator", &props.creator),
            ("subject", &props.subject),
            ("description", &props.description),
        ] {
            if let Some(s) = value
                && s.trim().is_empty() && !s.is_empty()
            {
                report.push(ValidationIssue {
                    severity: Severity::Info,
                    category: IssueCategory::Metadata,
                    message: format!(
                        "Document property '{field_name}' is whitespace-only"
                    ),
                    location: format!("docProperties.{field_name}"),
                    suggestion: "Clear the field or set meaningful content".into(),
                    auto_fixable: true,
                });
            }
        }
    }
}

/// Quick check for ISO 8601 datetime format.
const fn is_valid_iso8601(s: &str) -> bool {
    // Accept common ISO 8601 patterns:
    // YYYY-MM-DD, YYYY-MM-DDTHH:MM:SS, YYYY-MM-DDTHH:MM:SSZ
    let bytes = s.as_bytes();
    if bytes.len() < 10 {
        return false;
    }
    // Check date portion: YYYY-MM-DD
    bytes[0].is_ascii_digit()
        && bytes[1].is_ascii_digit()
        && bytes[2].is_ascii_digit()
        && bytes[3].is_ascii_digit()
        && bytes[4] == b'-'
        && bytes[5].is_ascii_digit()
        && bytes[6].is_ascii_digit()
        && bytes[7] == b'-'
        && bytes[8].is_ascii_digit()
        && bytes[9].is_ascii_digit()
}

// -- Structure --

fn validate_structure(report: &mut ValidationReport, wb: &WorkbookData) {
    // Check row ordering within sheets.
    for sheet in &wb.sheets {
        let mut prev_row_idx: u32 = 0;
        for (i, row) in sheet.worksheet.rows.iter().enumerate() {
            if i > 0 && row.index <= prev_row_idx {
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

    repairs
}

fn repair_sheet_names(wb: &mut WorkbookData) -> usize {
    let mut repairs = 0;

    for (i, sheet) in wb.sheets.iter_mut().enumerate() {
        // Fix empty names.
        if sheet.name.is_empty() {
            sheet.name = format!("Sheet{}", i + 1);
            repairs += 1;
        }

        // Truncate overlong names.
        if sheet.name.chars().count() > MAX_SHEET_NAME_LEN {
            sheet.name = sheet.name.chars().take(MAX_SHEET_NAME_LEN).collect();
            repairs += 1;
        }

        // Remove forbidden characters.
        if sheet.name.chars().any(|c| FORBIDDEN_SHEET_CHARS.contains(&c)) {
            sheet.name = sheet
                .name
                .chars()
                .filter(|c| !FORBIDDEN_SHEET_CHARS.contains(c))
                .collect();
            repairs += 1;
        }
    }

    // Deduplicate names.
    let mut seen: HashSet<String> = HashSet::with_capacity(wb.sheets.len());
    for sheet in &mut wb.sheets {
        let lower = sheet.name.to_lowercase();
        if !seen.insert(lower) {
            let mut suffix = 2u32;
            loop {
                let candidate = format!("{}_{suffix}", sheet.name);
                if seen.insert(candidate.to_lowercase()) {
                    sheet.name = candidate;
                    repairs += 1;
                    break;
                }
                suffix += 1;
            }
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
            name: Some("Aptos".to_owned()),
            size: Some(11.0),
            ..Font::default()
        });
        repairs += 1;
    }

    // Ensure at least 2 fills (none + gray125).
    while styles.fills.len() < 2 {
        if styles.fills.is_empty() {
            styles.fills.push(Fill {
                pattern_type: "none".to_owned(),
                ..Fill::default()
            });
        }
        if styles.fills.len() < 2 {
            styles.fills.push(Fill {
                pattern_type: "gray125".to_owned(),
                ..Fill::default()
            });
        }
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
        // Fix non-ISO-8601 dates.
        for date_field in [&mut props.created, &mut props.modified] {
            let should_clear = date_field
                .as_deref()
                .is_some_and(|d| !is_valid_iso8601(d));
            if should_clear {
                *date_field = None;
                repairs += 1;
            }
        }

        // Clear whitespace-only string fields.
        for field in [
            &mut props.title,
            &mut props.creator,
            &mut props.subject,
            &mut props.description,
            &mut props.keywords,
            &mut props.last_modified_by,
            &mut props.category,
            &mut props.content_status,
            &mut props.application,
            &mut props.company,
            &mut props.manager,
        ] {
            let should_clear = field
                .as_deref()
                .is_some_and(|s| s.trim().is_empty() && !s.is_empty());
            if should_clear {
                *field = None;
                repairs += 1;
            }
        }
    }

    repairs
}

fn repair_row_ordering(wb: &mut WorkbookData) -> usize {
    let mut repairs = 0;

    for sheet in &mut wb.sheets {
        let rows = &mut sheet.worksheet.rows;
        let was_sorted = rows.windows(2).all(|w| w[0].index < w[1].index);
        if !was_sorted {
            rows.sort_by_key(|r| r.index);
            repairs += 1;
        }
    }

    repairs
}

// ---------------------------------------------------------------------------
// Color normalization
// ---------------------------------------------------------------------------

/// Resolve an indexed color (legacy Excel 97-2003) to a 6-char RGB hex string.
///
/// Returns `None` for out-of-range indices or special system colors (64, 65).
pub fn resolve_indexed_color(index: u32) -> Option<&'static str> {
    if (8..64).contains(&index) {
        Some(INDEXED_COLORS[(index - 8) as usize])
    } else if index < 8 {
        // Indices 0-7 map to same as 8-15.
        Some(INDEXED_COLORS[index as usize])
    } else {
        None // 64 = system foreground, 65 = system background
    }
}

/// Resolve a theme color index to a 6-char RGB hex string using the given theme.
///
/// Theme indices: 0=dk1, 1=lt1, 2=dk2, 3=lt2, 4-9=accent1-6, 10=hlink, 11=folHlink.
pub fn resolve_theme_color(theme_index: u32, theme: &ThemeColors) -> Option<&str> {
    match theme_index {
        0 => Some(&theme.dk1),
        1 => Some(&theme.lt1),
        2 => Some(&theme.dk2),
        3 => Some(&theme.lt2),
        4 => Some(&theme.accent1),
        5 => Some(&theme.accent2),
        6 => Some(&theme.accent3),
        7 => Some(&theme.accent4),
        8 => Some(&theme.accent5),
        9 => Some(&theme.accent6),
        10 => Some(&theme.hlink),
        11 => Some(&theme.fol_hlink),
        _ => None,
    }
}

/// Apply a tint value to a base RGB color.
///
/// Tint ranges from -1.0 (fully dark) to +1.0 (fully light).
/// Follows OOXML tint algorithm (ECMA-376 §18.8.19).
pub fn apply_tint(rgb_hex: &str, tint: f64) -> String {
    if rgb_hex.len() < 6 {
        return rgb_hex.to_owned();
    }

    let r = u8::from_str_radix(&rgb_hex[0..2], 16).unwrap_or(0);
    let g = u8::from_str_radix(&rgb_hex[2..4], 16).unwrap_or(0);
    let b = u8::from_str_radix(&rgb_hex[4..6], 16).unwrap_or(0);

    let tint_channel = |c: u8| -> u8 {
        let c_f = c as f64 / 255.0;
        let result = if tint < 0.0 {
            c_f * (1.0 + tint)
        } else {
            c_f * (1.0 - tint) + tint
        };
        (result.clamp(0.0, 1.0) * 255.0).round() as u8
    };

    format!("{:02X}{:02X}{:02X}", tint_channel(r), tint_channel(g), tint_channel(b))
}

// ---------------------------------------------------------------------------
// Theme XML generation
// ---------------------------------------------------------------------------

/// Generate a complete Office 2016 `theme1.xml` document.
///
/// This produces a valid `xl/theme/theme1.xml` that Excel expects, using the
/// given theme colors (or defaults if `None`).
pub fn generate_theme_xml(colors: Option<&ThemeColors>) -> Vec<u8> {
    let c = colors.cloned().unwrap_or_default();

    let mut xml = Vec::with_capacity(4096);
    xml.extend_from_slice(b"<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n");
    xml.extend_from_slice(
        b"<a:theme xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" name=\"Office Theme\">"
    );

    // themeElements
    xml.extend_from_slice(b"<a:themeElements>");

    // clrScheme
    xml.extend_from_slice(b"<a:clrScheme name=\"Office\">");
    write_theme_color(&mut xml, "dk1", &c.dk1, true);
    write_theme_color(&mut xml, "lt1", &c.lt1, true);
    write_theme_color(&mut xml, "dk2", &c.dk2, false);
    write_theme_color(&mut xml, "lt2", &c.lt2, false);
    write_theme_color(&mut xml, "accent1", &c.accent1, false);
    write_theme_color(&mut xml, "accent2", &c.accent2, false);
    write_theme_color(&mut xml, "accent3", &c.accent3, false);
    write_theme_color(&mut xml, "accent4", &c.accent4, false);
    write_theme_color(&mut xml, "accent5", &c.accent5, false);
    write_theme_color(&mut xml, "accent6", &c.accent6, false);
    write_theme_color(&mut xml, "hlink", &c.hlink, false);
    write_theme_color(&mut xml, "folHlink", &c.fol_hlink, false);
    xml.extend_from_slice(b"</a:clrScheme>");

    // fontScheme
    xml.extend_from_slice(b"<a:fontScheme name=\"Office\">");
    xml.extend_from_slice(
        b"<a:majorFont><a:latin typeface=\"Aptos Display\"/><a:ea typeface=\"\"/><a:cs typeface=\"\"/></a:majorFont>"
    );
    xml.extend_from_slice(
        b"<a:minorFont><a:latin typeface=\"Aptos\"/><a:ea typeface=\"\"/><a:cs typeface=\"\"/></a:minorFont>"
    );
    xml.extend_from_slice(b"</a:fontScheme>");

    // fmtScheme
    xml.extend_from_slice(b"<a:fmtScheme name=\"Office\">");

    // fillStyleLst
    xml.extend_from_slice(b"<a:fillStyleLst>");
    xml.extend_from_slice(b"<a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill>");
    xml.extend_from_slice(
        b"<a:gradFill rotWithShape=\"1\"><a:gsLst>"
    );
    xml.extend_from_slice(
        b"<a:gs pos=\"0\"><a:schemeClr val=\"phClr\"><a:lumMod val=\"110000\"/><a:satMod val=\"105000\"/><a:tint val=\"67000\"/></a:schemeClr></a:gs>"
    );
    xml.extend_from_slice(
        b"<a:gs pos=\"50000\"><a:schemeClr val=\"phClr\"><a:lumMod val=\"105000\"/><a:satMod val=\"103000\"/><a:tint val=\"73000\"/></a:schemeClr></a:gs>"
    );
    xml.extend_from_slice(
        b"<a:gs pos=\"100000\"><a:schemeClr val=\"phClr\"><a:lumMod val=\"105000\"/><a:satMod val=\"109000\"/><a:tint val=\"81000\"/></a:schemeClr></a:gs>"
    );
    xml.extend_from_slice(b"</a:gsLst><a:lin ang=\"5400000\" scaled=\"0\"/></a:gradFill>");
    xml.extend_from_slice(
        b"<a:gradFill rotWithShape=\"1\"><a:gsLst>"
    );
    xml.extend_from_slice(
        b"<a:gs pos=\"0\"><a:schemeClr val=\"phClr\"><a:satMod val=\"103000\"/><a:lumMod val=\"102000\"/><a:tint val=\"94000\"/></a:schemeClr></a:gs>"
    );
    xml.extend_from_slice(
        b"<a:gs pos=\"50000\"><a:schemeClr val=\"phClr\"><a:satMod val=\"110000\"/><a:lumMod val=\"100000\"/><a:shade val=\"100000\"/></a:schemeClr></a:gs>"
    );
    xml.extend_from_slice(
        b"<a:gs pos=\"100000\"><a:schemeClr val=\"phClr\"><a:lumMod val=\"99000\"/><a:satMod val=\"120000\"/><a:shade val=\"78000\"/></a:schemeClr></a:gs>"
    );
    xml.extend_from_slice(b"</a:gsLst><a:lin ang=\"5400000\" scaled=\"0\"/></a:gradFill>");
    xml.extend_from_slice(b"</a:fillStyleLst>");

    // lnStyleLst
    xml.extend_from_slice(b"<a:lnStyleLst>");
    xml.extend_from_slice(
        b"<a:ln w=\"6350\" cap=\"flat\" cmpd=\"sng\" algn=\"ctr\"><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill><a:prstDash val=\"solid\"/><a:miter lim=\"800000\"/></a:ln>"
    );
    xml.extend_from_slice(
        b"<a:ln w=\"12700\" cap=\"flat\" cmpd=\"sng\" algn=\"ctr\"><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill><a:prstDash val=\"solid\"/><a:miter lim=\"800000\"/></a:ln>"
    );
    xml.extend_from_slice(
        b"<a:ln w=\"19050\" cap=\"flat\" cmpd=\"sng\" algn=\"ctr\"><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill><a:prstDash val=\"solid\"/><a:miter lim=\"800000\"/></a:ln>"
    );
    xml.extend_from_slice(b"</a:lnStyleLst>");

    // effectStyleLst
    xml.extend_from_slice(b"<a:effectStyleLst>");
    xml.extend_from_slice(b"<a:effectStyle><a:effectLst/></a:effectStyle>");
    xml.extend_from_slice(b"<a:effectStyle><a:effectLst/></a:effectStyle>");
    xml.extend_from_slice(
        b"<a:effectStyle><a:effectLst><a:outerShdw blurRad=\"57150\" dist=\"19050\" dir=\"5400000\" algn=\"ctr\" rotWithShape=\"0\"><a:srgbClr val=\"000000\"><a:alpha val=\"63000\"/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle>"
    );
    xml.extend_from_slice(b"</a:effectStyleLst>");

    // bgFillStyleLst
    xml.extend_from_slice(b"<a:bgFillStyleLst>");
    xml.extend_from_slice(b"<a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill>");
    xml.extend_from_slice(b"<a:solidFill><a:schemeClr val=\"phClr\"><a:tint val=\"95000\"/><a:satMod val=\"170000\"/></a:schemeClr></a:solidFill>");
    xml.extend_from_slice(
        b"<a:gradFill rotWithShape=\"1\"><a:gsLst>"
    );
    xml.extend_from_slice(
        b"<a:gs pos=\"0\"><a:schemeClr val=\"phClr\"><a:tint val=\"93000\"/><a:satMod val=\"150000\"/><a:shade val=\"98000\"/><a:lumMod val=\"102000\"/></a:schemeClr></a:gs>"
    );
    xml.extend_from_slice(
        b"<a:gs pos=\"50000\"><a:schemeClr val=\"phClr\"><a:tint val=\"98000\"/><a:satMod val=\"130000\"/><a:shade val=\"90000\"/><a:lumMod val=\"103000\"/></a:schemeClr></a:gs>"
    );
    xml.extend_from_slice(
        b"<a:gs pos=\"100000\"><a:schemeClr val=\"phClr\"><a:shade val=\"63000\"/><a:satMod val=\"120000\"/></a:schemeClr></a:gs>"
    );
    xml.extend_from_slice(b"</a:gsLst><a:lin ang=\"5400000\" scaled=\"0\"/></a:gradFill>");
    xml.extend_from_slice(b"</a:bgFillStyleLst>");

    xml.extend_from_slice(b"</a:fmtScheme>");
    xml.extend_from_slice(b"</a:themeElements>");

    // objectDefaults + extraClrSchemeLst (required, but empty)
    xml.extend_from_slice(b"<a:objectDefaults/>");
    xml.extend_from_slice(b"<a:extraClrSchemeLst/>");

    xml.extend_from_slice(b"</a:theme>");
    xml
}

/// Write a single theme color element.
fn write_theme_color(xml: &mut Vec<u8>, name: &str, color: &str, use_sys_clr: bool) {
    xml.extend_from_slice(b"<a:");
    xml.extend_from_slice(name.as_bytes());
    xml.push(b'>');

    if use_sys_clr {
        let sys_val = if name == "dk1" { "windowText" } else { "window" };
        xml.extend_from_slice(b"<a:sysClr val=\"");
        xml.extend_from_slice(sys_val.as_bytes());
        xml.extend_from_slice(b"\" lastClr=\"");
        xml.extend_from_slice(color.as_bytes());
        xml.extend_from_slice(b"\"/>");
    } else {
        xml.extend_from_slice(b"<a:srgbClr val=\"");
        xml.extend_from_slice(color.as_bytes());
        xml.extend_from_slice(b"\"/>");
    }

    xml.extend_from_slice(b"</a:");
    xml.extend_from_slice(name.as_bytes());
    xml.push(b'>');
}

// ---------------------------------------------------------------------------
// Tests
// ---------------------------------------------------------------------------

#[cfg(test)]
mod tests {
    use super::*;
    use crate::ooxml::styles::Styles;
    use crate::ooxml::worksheet::{Cell, CellType, Row, WorksheetXml};
    use crate::SheetData;
    use pretty_assertions::assert_eq;

    fn empty_worksheet() -> WorksheetXml {
        WorksheetXml {
            dimension: None,
            rows: Vec::new(),
            merge_cells: Vec::new(),
            auto_filter: None,
            frozen_pane: None,
            split_pane: None,
            pane_selections: vec![],
            sheet_view: None,
            columns: Vec::new(),
            data_validations: Vec::new(),
            conditional_formatting: Vec::new(),
            hyperlinks: Vec::new(),
            page_setup: None,
            sheet_protection: None,
            comments: Vec::new(),
            tab_color: None,
            tables: Vec::new(),
            header_footer: None,
            outline_properties: None,
            sparkline_groups: Vec::new(),
            charts: Vec::new(),
            preserved_extensions: Vec::new(),
        }
    }

    fn minimal_workbook() -> WorkbookData {
        WorkbookData {
            sheets: vec![SheetData {
                name: "Sheet1".into(),
                state: None,
                worksheet: empty_worksheet(),
            }],
            date_system: crate::dates::DateSystem::Date1900,
            styles: Styles::default_styles(),
            defined_names: Vec::new(),
            shared_strings: None,
            doc_properties: None,
            theme_colors: Some(ThemeColors::default()),
            calc_chain: Vec::new(),
            workbook_views: Vec::new(),
            protection: None,
            preserved_entries: Default::default(),
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
        assert!(report.issues.iter().any(|i| i.category == IssueCategory::SheetName));
    }

    #[test]
    fn test_validate_long_sheet_name() {
        let mut wb = minimal_workbook();
        wb.sheets[0].name = "A".repeat(50);
        let report = validate_workbook(&wb);
        assert!(!report.is_valid);
        assert!(report.issues.iter().any(|i| {
            i.category == IssueCategory::SheetName && i.message.contains("exceeds")
        }));
    }

    #[test]
    fn test_validate_duplicate_sheet_names() {
        let mut wb = minimal_workbook();
        wb.sheets.push(SheetData {
            name: "sheet1".into(), // case-insensitive duplicate
            state: None,
            worksheet: empty_worksheet(),
        });
        let report = validate_workbook(&wb);
        assert!(!report.is_valid);
        assert!(report.issues.iter().any(|i| {
            i.category == IssueCategory::SheetName && i.message.contains("Duplicate")
        }));
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
            height: None,
            hidden: false,
            outline_level: None,
            collapsed: false,
        });
        let report = validate_workbook(&wb);
        assert!(!report.is_valid);
        assert!(report.issues.iter().any(|i| i.category == IssueCategory::StyleIndex));
    }

    #[test]
    fn test_validate_dangling_font_id() {
        let mut wb = minimal_workbook();
        wb.styles.cell_xfs[0].font_id = 99;
        let report = validate_workbook(&wb);
        assert!(!report.is_valid);
        assert!(report.issues.iter().any(|i| {
            i.category == IssueCategory::StyleIndex && i.message.contains("fontId")
        }));
    }

    #[test]
    fn test_validate_overlapping_merges() {
        let mut wb = minimal_workbook();
        wb.sheets[0].worksheet.merge_cells = vec!["A1:B2".into(), "B2:C3".into()];
        let report = validate_workbook(&wb);
        assert!(!report.is_valid);
        assert!(report.issues.iter().any(|i| {
            i.category == IssueCategory::MergeCell && i.message.contains("overlaps")
        }));
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
            height: None,
            hidden: false,
            outline_level: None,
            collapsed: false,
        });
        let report = validate_workbook(&wb);
        assert!(!report.is_valid);
        assert!(report.issues.iter().any(|i| i.category == IssueCategory::SharedString));
    }

    #[test]
    fn test_validate_missing_theme() {
        let mut wb = minimal_workbook();
        wb.theme_colors = None;
        let report = validate_workbook(&wb);
        // Missing theme is info-level, not an error.
        assert!(report.is_valid);
        assert!(report.issues.iter().any(|i| i.category == IssueCategory::Theme));
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
        assert!(report.issues.iter().any(|i| {
            i.category == IssueCategory::Metadata && i.message.contains("ISO-8601")
        }));
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
            height: None,
            hidden: false,
            outline_level: None,
            collapsed: false,
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
            Row { index: 5, cells: Vec::new(), height: None, hidden: false, outline_level: None, collapsed: false },
            Row { index: 2, cells: Vec::new(), height: None, hidden: false, outline_level: None, collapsed: false },
            Row { index: 8, cells: Vec::new(), height: None, hidden: false, outline_level: None, collapsed: false },
        ];
        let repairs = repair_workbook(&mut wb);
        assert!(repairs > 0);
        let indices: Vec<u32> = wb.sheets[0].worksheet.rows.iter().map(|r| r.index).collect();
        assert_eq!(indices, vec![2, 5, 8]);
    }

    // -- Color normalization tests --

    #[test]
    fn test_resolve_indexed_color() {
        assert_eq!(resolve_indexed_color(8), Some("000000"));
        assert_eq!(resolve_indexed_color(9), Some("FFFFFF"));
        assert_eq!(resolve_indexed_color(10), Some("FF0000"));
        assert_eq!(resolve_indexed_color(64), None);
    }

    #[test]
    fn test_resolve_theme_color() {
        let theme = ThemeColors::default();
        assert_eq!(resolve_theme_color(0, &theme), Some("000000"));
        assert_eq!(resolve_theme_color(1, &theme), Some("FFFFFF"));
        assert_eq!(resolve_theme_color(4, &theme), Some("4472C4"));
        assert_eq!(resolve_theme_color(12, &theme), None);
    }

    #[test]
    fn test_apply_tint() {
        // No tint.
        assert_eq!(apply_tint("4472C4", 0.0), "4472C4");
        // Full white tint.
        assert_eq!(apply_tint("000000", 1.0), "FFFFFF");
        // Full dark tint.
        assert_eq!(apply_tint("FFFFFF", -1.0), "000000");
        // Partial tint.
        let result = apply_tint("4472C4", 0.5);
        assert_eq!(result.len(), 6);
        // Verify it's lighter.
        let r = u8::from_str_radix(&result[0..2], 16).unwrap();
        assert!(r > 0x44);
    }

    #[test]
    fn test_iso8601_check() {
        assert!(is_valid_iso8601("2024-01-15"));
        assert!(is_valid_iso8601("2024-01-15T10:30:00Z"));
        assert!(!is_valid_iso8601("not-a-date"));
        assert!(!is_valid_iso8601("2024"));
    }

    // -- Theme XML generation --

    #[test]
    fn test_generate_theme_xml() {
        let xml = generate_theme_xml(None);
        let s = std::str::from_utf8(&xml).unwrap();
        assert!(s.contains("a:theme"));
        assert!(s.contains("a:clrScheme"));
        assert!(s.contains("a:fontScheme"));
        assert!(s.contains("a:fmtScheme"));
        assert!(s.contains("windowText"));
        assert!(s.contains("000000"));
        assert!(s.contains("Aptos"));
    }

    #[test]
    fn test_generate_theme_xml_custom_colors() {
        let colors = ThemeColors {
            dk1: "111111".into(),
            lt1: "EEEEEE".into(),
            ..ThemeColors::default()
        };
        let xml = generate_theme_xml(Some(&colors));
        let s = std::str::from_utf8(&xml).unwrap();
        assert!(s.contains("111111"));
        assert!(s.contains("EEEEEE"));
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
        assert!(report_after.is_valid, "After repair: {:?}", report_after.issues);
    }
}
