//! Rust core of modern-xlsx: reads and writes SpreadsheetML packages (`.xlsx`)
//! through the [`WorkbookData`] model.
//!
//! # Reading
//!
//! [`reader::read_xlsx`] unzips the package under the default
//! [`ZipSecurityLimits`](crate::zip::ZipSecurityLimits), parses the parts it
//! models and returns a [`WorkbookData`]; the entries it keeps unparsed are in
//! [`WorkbookData::preserved_entries`]. [`reader::read_xlsx_json`] returns the
//! same model as a JSON string for the WASM bridge: worksheet rows are written
//! straight to JSON while each sheet is parsed, and the other fields go through
//! `serde_json`.
//!
//! [`reader::read_xlsx_with_options`] and [`reader::read_xlsx_json_with_options`]
//! take custom limits. [`reader::read_xlsx_json_with_password`] and
//! [`reader::read_xlsx_json_with_password_and_options`] also decrypt a
//! password-protected file.
//!
//! # Writing
//!
//! [`writer::write_xlsx`] builds the shared string table from the text of the
//! shared-string cells, writes each worksheet with its string indices filled
//! in during serialization, and zips the parts.
//!
//! # Other modules
//!
//! - [`validate`] checks a [`WorkbookData`] ([`validate::validate_workbook`])
//!   and repairs what it can ([`validate::repair_workbook`]).
//! - [`streaming_writer`] writes rows straight into the ZIP
//!   ([`streaming_writer::StreamingWriterCore`]).
//! - [`ole2`] reads and writes OLE2 compound files: it tells encrypted packages
//!   from legacy `.xls` files and, with `encryption`, holds the ECMA-376
//!   encryption.
//!
//! # Features
//!
//! - `encryption` (default) compiles `ole2::crypto` and
//!   `writer::write_xlsx_with_password`, and lets the password readers decrypt.
//!   Without it the password readers still exist, but an encrypted file fails
//!   with [`ModernXlsxError::PasswordProtected`].
//! - `parallel` parses the worksheets of [`reader::read_xlsx`] and
//!   [`reader::read_xlsx_with_options`] on rayon's thread pool; the JSON
//!   readers parse sheets one after another. The `modern-xlsx-wasm` crate does
//!   not enable it.

pub mod dates;
pub mod errors;
pub mod ole2;
pub mod ooxml;
pub mod reader;
pub mod streaming_writer;
pub mod validate;
pub mod writer;
pub mod zip;

pub use errors::{ModernXlsxError, Result};

use std::collections::BTreeMap;

use serde::{Deserialize, Serialize};

use crate::dates::DateSystem;
use crate::ooxml::calc_chain::CalcChainEntry;
use crate::ooxml::doc_props::DocProperties;
use crate::ooxml::shared_strings::SharedStringTable;
use crate::ooxml::styles::Styles;
use crate::ooxml::theme::ThemeColors;
use crate::ooxml::workbook::{
    CalcPr, DefinedName, FileVersion, WorkbookPr, WorkbookProtection, WorkbookView,
};
use crate::ooxml::worksheet::WorksheetXml;

// ---------------------------------------------------------------------------
// Canonical WorkbookData — shared by reader and writer
// ---------------------------------------------------------------------------

/// Top-level representation of a workbook.
///
/// The reader populates `shared_strings` with the parsed SST from the XLSX
/// archive. The writer ignores this field and builds the SST internally from
/// the cell values.
#[derive(Debug, Clone, Default, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct WorkbookData {
    /// The sheets in this workbook.
    pub sheets: Vec<SheetData>,
    /// The date system used (1900 or 1904).
    pub date_system: DateSystem,
    /// The styles object for this workbook.
    pub styles: Styles,
    /// Named ranges / defined names.
    #[serde(default)]
    pub defined_names: Vec<DefinedName>,
    /// The shared string table (populated by the reader, ignored by the writer).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub shared_strings: Option<SharedStringTable>,
    /// Document properties (docProps/core.xml + docProps/app.xml).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub doc_properties: Option<DocProperties>,
    /// Theme colors extracted from xl/theme/theme1.xml. A workbook without a
    /// preserved `xl/theme/theme1.xml` is written with a generated Office
    /// theme in these colors (the defaults when `None`). The reader keeps a
    /// file's theme in `preserved_entries`, except one byte-identical to that
    /// generated theme.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub theme_colors: Option<ThemeColors>,
    /// Calculation chain entries from xl/calcChain.xml. Each `sheet_id` is the
    /// 1-based position of the sheet in `sheets`.
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub calc_chain: Vec<CalcChainEntry>,
    /// Workbook view settings from `<bookViews>` in workbook.xml.
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub workbook_views: Vec<WorkbookView>,
    /// Workbook-level protection settings from `<workbookProtection>`.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub protection: Option<WorkbookProtection>,
    /// Calculation settings from `<calcPr>` in workbook.xml.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub calc_pr: Option<CalcPr>,
    /// Application version from `<fileVersion>` in workbook.xml.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub file_version: Option<FileVersion>,
    /// `<workbookPr>` attributes other than `date1904` (see `date_system`).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub workbook_pr: Option<WorkbookPr>,
    /// Pivot cache definitions from xl/pivotCache/.
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub pivot_caches: Vec<ooxml::pivot_table::PivotCacheDefinitionData>,
    /// Pivot cache records from xl/pivotCache/.
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub pivot_cache_records: Vec<ooxml::pivot_table::PivotCacheRecordsData>,
    /// Persons list for threaded comments (xl/persons/person.xml).
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub persons: Vec<ooxml::threaded_comments::PersonData>,
    /// Slicer cache definitions from xl/slicerCaches/.
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub slicer_caches: Vec<ooxml::slicers::SlicerCacheData>,
    /// Timeline cache definitions from xl/timelineCaches/.
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub timeline_caches: Vec<ooxml::timelines::TimelineCacheData>,
    /// ZIP paths of the external link parts (`xl/externalLinks/externalLink{n}.xml`,
    /// kept in `preserved_entries`) in workbook.xml `<externalReferences>`
    /// order, which formulas use as `[n]`. Preserved links missing from it are
    /// written after the listed ones, in numeric order.
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub external_references: Vec<String>,
    /// ZIP entries the reader keeps as raw bytes because it does not consume
    /// them into the model: every part it does not model (media, VML, custom
    /// XML, external links, printer settings, VBA and the like), every
    /// worksheet `.rels`, a drawing unless the reader takes it as chart-only
    /// (one or more anchors, all charts, its `.rels` parsed and no chart
    /// malformed), a theme other than the one the writer generates, and any
    /// comment, table, drawing, chart, pivot, threaded-comment, person, slicer
    /// or timeline part that fails to parse (a malformed chart keeps every
    /// chart of its drawing here, and a malformed pivot part keeps every pivot
    /// part here).
    ///
    /// A worksheet `.rels` is keyed by the sheet's position, as
    /// `xl/worksheets/_rels/sheet{n}.xml.rels` (n is the sheet's 1-based
    /// position, the name the writer gives the sheet part), whatever the
    /// source part was named, when the source part is directly in
    /// `xl/worksheets/`.
    ///
    /// The writer copies each entry unchanged, except where it writes a part
    /// at the same path. A sheet that needs generated relationships has the
    /// kept `.rels` at `xl/worksheets/_rels/sheet{n}.xml.rels` rebuilt:
    /// hyperlink relationships are regenerated from the model, and comment,
    /// threaded-comment, slicer and timeline relationships point at the
    /// written part or are dropped when they name a part that is neither
    /// written nor kept. If no relationship is left, that `.rels` is dropped.
    /// A drawing at the path of a sheet's chart drawing, with its `.rels`, is
    /// merged with the regenerated chart frames. A kept custom XML data part
    /// (`customXml/item{n}.xml`) gets a workbook relationship, and a kept
    /// `docProps/custom.xml` a package relationship and its Override.
    #[serde(default, skip_serializing_if = "BTreeMap::is_empty")]
    pub preserved_entries: BTreeMap<String, Vec<u8>>,
    /// The content type the source `[Content_Types].xml` declared for each
    /// entry of `preserved_entries` but a `.rels`, keyed by its ZIP path: its
    /// Override, else the Default for its extension (matched ignoring ASCII
    /// case). Paths the source declared no type for are absent.
    ///
    /// The writer declares each recorded type, after its own rules, as an
    /// Override (none when the Default for the extension already gives it;
    /// an Override the writer's own rules set is then kept), except for a
    /// `.rels` and a part it writes itself at that path.
    #[serde(default, skip_serializing_if = "BTreeMap::is_empty")]
    pub preserved_content_types: BTreeMap<String, String>,
}

/// A single sheet inside a workbook.
#[derive(Debug, Clone, Default, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct SheetData {
    /// The user-visible sheet name.
    pub name: String,
    /// Sheet visibility: `"hidden"` or `"veryHidden"`. Omitted (`None`) means visible.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub state: Option<String>,
    /// The parsed worksheet content (rows, cells, merges, etc.).
    pub worksheet: WorksheetXml,
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::ooxml::charts::ChartType;
    use crate::zip::reader::{ZipSecurityLimits, read_zip_entries};

    #[test]
    fn workbook_data_default_is_valid() {
        let wb = WorkbookData::default();
        assert_eq!(wb.date_system, DateSystem::Date1900);
        assert_eq!(wb.styles.fonts.len(), 1);
        assert_eq!(wb.styles.fills.len(), 2);
        assert_eq!(wb.styles.borders.len(), 1);
        assert_eq!(wb.styles.cell_xfs.len(), 1);
        assert!(wb.sheets.is_empty());
    }

    #[test]
    fn default_workbook_writes_valid_styles() {
        let wb = WorkbookData {
            sheets: vec![SheetData {
                name: "S".into(),
                worksheet: WorksheetXml::default(),
                ..Default::default()
            }],
            ..Default::default()
        };
        let bytes = writer::write_xlsx(&wb).unwrap();
        let entries = read_zip_entries(&bytes, &ZipSecurityLimits::default()).unwrap();
        let styles = String::from_utf8(entries["xl/styles.xml"].clone()).unwrap();
        assert!(styles.contains("<cellXfs count=\"1\">"), "{styles}");
        assert_eq!(
            reader::read_xlsx(&bytes).unwrap().date_system,
            DateSystem::Date1900
        );
    }

    #[test]
    fn chart_type_default_is_bar() {
        assert_eq!(ChartType::default(), ChartType::Bar);
    }
}
