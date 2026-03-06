pub mod dates;
pub mod errors;
pub mod number_format;
#[cfg(feature = "encryption")]
pub mod ole2;
pub mod ooxml;
pub mod reader;
pub mod streaming;
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
use crate::ooxml::workbook::{DefinedName, WorkbookProtection, WorkbookView};
use crate::ooxml::worksheet::WorksheetXml;

// ---------------------------------------------------------------------------
// Canonical WorkbookData — shared by reader and writer
// ---------------------------------------------------------------------------

/// Top-level representation of a workbook.
///
/// The reader populates `shared_strings` with the parsed SST from the XLSX
/// archive. The writer ignores this field and builds the SST internally from
/// the cell values.
#[derive(Debug, Clone, Serialize, Deserialize)]
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
    /// Theme colors extracted from xl/theme/theme1.xml (read-only, not written).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub theme_colors: Option<ThemeColors>,
    /// Calculation chain entries from xl/calcChain.xml.
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub calc_chain: Vec<CalcChainEntry>,
    /// Workbook view settings from <bookViews> in workbook.xml.
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub workbook_views: Vec<WorkbookView>,
    /// Workbook-level protection settings from `<workbookProtection>`.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub protection: Option<WorkbookProtection>,
    /// Pivot cache definitions from xl/pivotCache/.
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub pivot_caches: Vec<crate::ooxml::pivot_table::PivotCacheDefinitionData>,
    /// Pivot cache records from xl/pivotCache/.
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub pivot_cache_records: Vec<crate::ooxml::pivot_table::PivotCacheRecordsData>,
    /// Persons list for threaded comments (xl/persons/person.xml).
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub persons: Vec<crate::ooxml::threaded_comments::PersonData>,
    /// Slicer cache definitions from xl/slicerCaches/.
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub slicer_caches: Vec<crate::ooxml::slicers::SlicerCacheData>,
    /// Timeline cache definitions from xl/timelineCaches/.
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub timeline_caches: Vec<crate::ooxml::timelines::TimelineCacheData>,
    /// Opaque ZIP entries not parsed by the reader (drawings, media, charts, etc.)
    /// Preserved through roundtrip without modification.
    #[serde(default, skip_serializing_if = "BTreeMap::is_empty")]
    pub preserved_entries: BTreeMap<String, Vec<u8>>,
}

/// A single sheet inside a workbook.
#[derive(Debug, Clone, Serialize, Deserialize)]
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
