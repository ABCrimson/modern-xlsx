//! Full XLSX read orchestrator.
//!
//! Decompresses a `.xlsx` ZIP archive (decrypting an encrypted package first
//! when a password is given and the `encryption` feature is on), parses the
//! OPC / SpreadsheetML parts it models, and assembles them into a
//! [`WorkbookData`] struct; the entries it keeps unparsed go to
//! [`WorkbookData::preserved_entries`].
//!
//! Entry points: [`read_xlsx`] and [`read_xlsx_with_options`] build the
//! struct; [`read_xlsx_json`], [`read_xlsx_json_with_options`],
//! [`read_xlsx_json_with_password`] and
//! [`read_xlsx_json_with_password_and_options`] return it as JSON.

use core::hint::cold_path;
use std::collections::{BTreeMap, HashMap, HashSet};
use std::fmt::{self, Write as _};

use log::{Level, debug, log, warn};

#[cfg(feature = "parallel")]
use rayon::prelude::*;

use crate::dates::DateSystem;
use crate::ole2::cfb::reader::Cfb;
use crate::ole2::detect::{FileFormat, Ole2Kind, classify_ole2_with_cfb, detect_format};
use crate::ole2::encryption_info::build_encrypted_error_with_cfb;
use crate::ooxml::{
    calc_chain, charts, comments,
    content_types::ContentTypes,
    doc_props,
    pivot_table::{
        CacheRelationship, PivotCacheDefinitionData, PivotCacheRecordsData, PivotTableData,
    },
    relationships::{
        REL_COMMENTS, REL_DRAWING, REL_PIVOT_CACHE_DEF, REL_PIVOT_CACHE_REC, REL_PIVOT_TABLE,
        REL_SLICER, REL_SLICER_CACHE, REL_TABLE, REL_THREADED_COMMENTS, REL_TIMELINE,
        REL_TIMELINE_CACHE, Relationships, rels_path, resolve_target,
    },
    shared_strings::{ExpansionBudget, SharedStringTable},
    slicers::{self, SlicerCacheData, SlicerData},
    styles::Styles,
    tables::TableDefinition,
    theme,
    threaded_comments::{self, PersonData, ThreadedCommentData},
    timelines::{self, TimelineCacheData, TimelineData},
    workbook::{SheetState, WorkbookXml},
    worksheet::{SheetContext, WorksheetXml},
};
use crate::zip::reader::{ZipSecurityLimits, read_zip_entries};
use crate::{ModernXlsxError, Result, SheetData, WorkbookData};

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

/// Convert a [`SheetState`] to the JSON-bridge string representation.
///
/// `Visible` maps to `None` (omitted via `skip_serializing_if`),
/// while `Hidden` and `VeryHidden` map to their camelCase strings.
#[inline]
const fn sheet_state_to_str(state: SheetState) -> Option<&'static str> {
    match state {
        SheetState::Visible => None,
        SheetState::Hidden => Some("hidden"),
        SheetState::VeryHidden => Some("veryHidden"),
    }
}

// ---------------------------------------------------------------------------
// Log hygiene
// ---------------------------------------------------------------------------

/// Longest file-derived text a reader log line prints: 160 bytes after
/// escaping, not counting the quotes and the truncation note.
const LOG_TEXT_LIMIT: usize = 160;

/// File-derived text made safe for one log line: quoted, every control,
/// non-printable and grapheme-extending character escaped as by `Debug` (a
/// newline becomes `\n`, ESC `\u{1b}`), and cut at [`LOG_TEXT_LIMIT`] bytes
/// with a note of the full length.
fn log_text(raw: impl fmt::Display) -> String {
    let raw = raw.to_string();
    let mut out = String::with_capacity(raw.len().min(LOG_TEXT_LIMIT) + 2);
    out.push('"');
    for c in raw.chars() {
        let start = out.len();
        out.extend(c.escape_debug());
        if out.len() - 1 > LOG_TEXT_LIMIT {
            out.truncate(start);
            let _ = write!(out, "\" (truncated, {} bytes)", raw.len());
            return out;
        }
    }
    out.push('"');
    out
}

/// One kind of per-part log event, aggregated over a read.
struct DigestEntry {
    level: Level,
    kind: &'static str,
    count: usize,
    /// The first offending value, already passed through [`log_text`].
    first: String,
}

/// Aggregates the reader's per-part log events (a missing part, a malformed
/// part, a part found) into one line per kind: the kind, how many times it
/// happened and the first example. A package with a million dangling
/// relationships therefore logs one line, not a million lines of file text.
///
/// The lines are emitted once, when the digest is dropped at the end of the
/// read, on success and on error alike.
#[derive(Default)]
struct LogDigest {
    /// In order of first occurrence; there are at most a few dozen kinds, so a
    /// linear scan beats hashing.
    entries: Vec<DigestEntry>,
}

impl LogDigest {
    /// Counts one `kind` event; `value` is formatted only for the first one.
    fn note(&mut self, level: Level, kind: &'static str, value: impl fmt::Display) {
        if let Some(entry) = self.entries.iter_mut().find(|e| e.kind == kind) {
            entry.count += 1;
        } else {
            self.entries.push(DigestEntry {
                level,
                kind,
                count: 1,
                first: log_text(value),
            });
        }
    }
}

impl Drop for LogDigest {
    fn drop(&mut self) {
        for e in &self.entries {
            log!(e.level, "{}: {} (first: {})", e.kind, e.count, e.first);
        }
    }
}

// ---------------------------------------------------------------------------
// Shared reader context
// ---------------------------------------------------------------------------

/// The package's content types part.
const CONTENT_TYPES_PATH: &str = "[Content_Types].xml";

/// Well-known ZIP entry paths that are parsed (not preserved verbatim).
const KNOWN_STATIC_PATHS: &[&str] = &[
    CONTENT_TYPES_PATH,
    "_rels/.rels",
    "xl/workbook.xml",
    "xl/sharedStrings.xml",
    "xl/styles.xml",
    "xl/_rels/workbook.xml.rels",
    "docProps/core.xml",
    "docProps/app.xml",
    "xl/calcChain.xml",
];

/// The theme part's ZIP path.
const THEME_PATH: &str = "xl/theme/theme1.xml";

/// The persons part's ZIP path (the authors of threaded comments).
const PERSONS_PATH: &str = "xl/persons/person.xml";

/// Gives each part the reader looks up by a fixed name
/// ([`KNOWN_STATIC_PATHS`], [`THEME_PATH`], [`PERSONS_PATH`]) that name when
/// the package holds it in another ASCII letter case only: OPC part names are
/// equal ignoring ASCII case, and Excel opens a package whose
/// `[Content_Types].xml` or `xl/styles.xml` is spelled `[content_types].xml`
/// or `xl/Styles.xml`. The entry is then read like the canonical one, never
/// preserved under its own spelling. An entry of the exact name always wins;
/// of two other spellings, the first in byte order is taken, and the second
/// stays an entry of its own.
fn canonicalize_fixed_names(entries: &mut HashMap<String, Vec<u8>>) {
    let fixed_names = || {
        KNOWN_STATIC_PATHS
            .iter()
            .copied()
            .chain([THEME_PATH, PERSONS_PATH])
    };
    let mut renames: Vec<(String, &'static str)> = entries
        .keys()
        .filter_map(|name| {
            fixed_names()
                .find(|fixed| fixed.eq_ignore_ascii_case(name) && *fixed != name.as_str())
                .filter(|fixed| !entries.contains_key(*fixed))
                .map(|fixed| (name.clone(), fixed))
        })
        .collect();
    renames.sort_unstable();
    for (name, fixed) in renames {
        if !entries.contains_key(fixed)
            && let Some(data) = entries.remove(&name)
        {
            debug!("{} read as {fixed}", log_text(&name));
            entries.insert(fixed.to_owned(), data);
        }
    }
}

/// Common parsed data shared between the struct-based and JSON-streaming readers.
struct ReaderContext {
    entries: HashMap<String, Vec<u8>>,
    workbook_xml: WorkbookXml,
    sst: SharedStringTable,
    styles: Styles,
    /// The read's one [`ExpansionBudget`], made before `xl/workbook.xml` is
    /// parsed: it holds the defined names, shared string and style records
    /// counted, and the comments, the charts, the threaded comments and the
    /// sheets count on after them.
    budget: ExpansionBudget,
    doc_properties: Option<doc_props::DocProperties>,
    theme_colors: Option<theme::ThemeColors>,
    calc_chain: Vec<calc_chain::CalcChainEntry>,
    /// Workbook-level relationships (they resolve the pivot, slicer and
    /// timeline caches and the external links).
    wb_rels: Relationships,
    /// (sheet_name, zip_path, visibility_state) triples in workbook order.
    sheet_targets: Vec<(String, String, SheetState)>,
    /// Each sheet's parsed `.rels`, parallel to `sheet_targets`; `None` when
    /// the sheet has none or it is malformed (then it stays preserved).
    sheet_rels: Vec<Option<Relationships>>,
    /// Paths of parsed entries (used to compute preserved_entries).
    known_dynamic: HashSet<String>,
    /// The source `[Content_Types].xml`; `None` when it is absent or
    /// malformed. It types the preserved entries
    /// ([`preserved_content_types`]).
    content_types: Option<ContentTypes>,
    /// Per-part log events of this read, emitted as summary lines on drop.
    log: LogDigest,
}

/// Parse all shared XLSX metadata from the ZIP entries.
///
/// When `password` is `Some`, encrypted OLE2 files are decrypted before parsing.
/// When `password` is `None`, encrypted files produce a descriptive error. The lite
/// build (no `encryption` feature) cannot decrypt: an encrypted file is always
/// `PasswordProtected`, whether or not a password is given.
fn parse_common(
    data: &[u8],
    limits: &ZipSecurityLimits,
    password: Option<&str>,
) -> Result<ReaderContext> {
    // One sniffing path in every build; only decryption needs the `encryption` feature.
    // `decrypted` owns the decrypted bytes; `actual_data` borrows either `data` or it.
    #[cfg(feature = "encryption")]
    let decrypted: Vec<u8>;
    let actual_data: &[u8];

    {
        use crate::ole2::detect::{ERR_LEGACY_XLS, ERR_NOT_XLSX, ERR_OLE2_UNKNOWN};

        match detect_format(data) {
            FileFormat::Zip => {
                actual_data = data;
            }
            // The compound file is parsed once (header, FAT, directory); the
            // classification and every stream read below share that parse.
            FileFormat::Ole2 => {
                let cfb = Cfb::parse(data)?;
                match classify_ole2_with_cfb(&cfb) {
                    Ole2Kind::EncryptedXlsx => {
                        #[cfg(feature = "encryption")]
                        if let Some(pw) = password {
                            decrypted = crate::ole2::crypto::decrypt_file_with_cfb(cfb, pw)?;
                            actual_data = &decrypted;
                        } else {
                            cold_path();
                            return Err(build_encrypted_error_with_cfb(&cfb));
                        }
                        // The lite build cannot decrypt, with or without a password.
                        #[cfg(not(feature = "encryption"))]
                        {
                            let _ = password;
                            cold_path();
                            return Err(build_encrypted_error_with_cfb(&cfb));
                        }
                    }
                    Ole2Kind::LegacyXls => {
                        cold_path();
                        return Err(ModernXlsxError::LegacyFormat(ERR_LEGACY_XLS.into()));
                    }
                    Ole2Kind::Unknown => {
                        cold_path();
                        return Err(ModernXlsxError::UnrecognizedFormat(ERR_OLE2_UNKNOWN.into()));
                    }
                }
            }
            FileFormat::Unknown => {
                cold_path();
                return Err(ModernXlsxError::UnrecognizedFormat(ERR_NOT_XLSX.into()));
            }
        }
    }

    let mut entries = read_zip_entries(actual_data, limits)?;
    canonicalize_fixed_names(&mut entries);

    // Parse workbook (required part).
    debug!("parsing workbook.xml");
    let workbook_data = entries
        .get("xl/workbook.xml")
        .ok_or_else(|| {
            cold_path();
            ModernXlsxError::MissingPart(
                "ZIP entry 'xl/workbook.xml' not found in archive — the file may be corrupt or not a valid XLSX".into()
            )
        })?;
    if workbook_data.is_empty() {
        cold_path();
        return Err(ModernXlsxError::MissingPart(
            "ZIP entry 'xl/workbook.xml' is empty — the file may be corrupt or not a valid XLSX"
                .into(),
        ));
    }
    // The read's one budget, in this order: the workbook's defined names, the
    // shared strings, the styles, the comments (the author copied into each),
    // the charts, the threaded comments and the sheets.
    let mut budget = ExpansionBudget::new(limits.max_json_length);
    let workbook_xml = WorkbookXml::parse_within(workbook_data, &mut budget)?;

    // Parse shared strings (optional — some workbooks have no strings).
    let sst = entries
        .get("xl/sharedStrings.xml")
        .map(|d| SharedStringTable::parse_within(d, &mut budget))
        .transpose()?
        .unwrap_or_else(|| {
            warn!("shared strings table not found");
            SharedStringTable::empty()
        });

    // Parse styles (optional — use defaults when absent), counting each
    // record kept in the read's budget.
    let styles = entries
        .get("xl/styles.xml")
        .map(|d| Styles::parse_within(d, &mut budget))
        .transpose()?
        .unwrap_or_else(Styles::default_styles);

    // Parse document properties (optional, never fatal: a malformed core.xml
    // drops the core properties, a malformed app.xml only the extended ones,
    // and the read goes on). Only a security limit stays an error.
    let mut doc_properties = match entries
        .get("docProps/core.xml")
        .map(|d| doc_props::parse_core(d))
    {
        None => None,
        Some(Ok(props)) => Some(props),
        Some(Err(e @ ModernXlsxError::Security(_))) => return Err(e),
        Some(Err(e)) => {
            warn!(
                "docProps/core.xml: {}; document properties dropped",
                log_text(&e)
            );
            None
        }
    };
    if let Some(app_data) = entries.get("docProps/app.xml") {
        // A copy, so a failed parse leaves the core properties as they were
        // and never keeps half-applied extended fields.
        let mut props = doc_properties.clone().unwrap_or_default();
        match doc_props::parse_app(&mut props, app_data) {
            Ok(()) => doc_properties = Some(props),
            Err(e @ ModernXlsxError::Security(_)) => return Err(e),
            Err(e) => warn!(
                "docProps/app.xml: {}; extended properties dropped",
                log_text(&e)
            ),
        }
    }

    // Parse theme colors (optional — xl/theme/theme1.xml). A malformed theme
    // is not fatal: its colours are dropped and, with `theme_colors` None, the
    // part stays in preserved_entries verbatim. Only a security limit is an error.
    let theme_colors = match entries.get(THEME_PATH).map(|d| theme::parse(d)) {
        None => None,
        Some(Ok(c)) => Some(c),
        Some(Err(e @ ModernXlsxError::Security(_))) => return Err(e),
        Some(Err(e)) => {
            warn!(
                "{THEME_PATH}: {}; theme colours dropped, part kept verbatim",
                log_text(&e)
            );
            None
        }
    };

    // Parse calculation chain (optional — xl/calcChain.xml). Each entry's `i`
    // is a source sheetId, but the writer renumbers sheetId to 1..n by
    // position, so map it to that 1-based position now.
    // The part is an optional cache: any failure but a security limit drops it.
    let mut calc_chain = match entries
        .get("xl/calcChain.xml")
        .map(|d| calc_chain::parse(d))
    {
        None => Vec::new(),
        Some(Ok(chain)) => chain,
        Some(Err(e @ ModernXlsxError::Security(_))) => return Err(e),
        Some(Err(e)) => {
            warn!("xl/calcChain.xml: {}; calcChain dropped", log_text(&e));
            Vec::new()
        }
    };
    // One map, built once: a linear scan per entry would cost entries x sheets,
    // and both counts come from the file. The first sheet with a sheetId wins.
    let mut position_by_sheet_id: HashMap<u32, u32> =
        HashMap::with_capacity(workbook_xml.sheets.len());
    for (index, sheet) in workbook_xml.sheets.iter().enumerate() {
        position_by_sheet_id
            .entry(sheet.sheet_id)
            .or_insert_with(|| u32::try_from(index + 1).unwrap_or(u32::MAX));
    }
    let mut dropped = 0usize;
    calc_chain.retain_mut(|entry| {
        if let Some(&position) = position_by_sheet_id.get(&entry.sheet_id) {
            entry.sheet_id = position;
            true
        } else {
            dropped += 1;
            false
        }
    });
    if dropped > 0 {
        warn!("calcChain: {dropped} entries name a sheetId which no sheet has; entries dropped");
    }

    // Parse workbook relationships (optional — needed to resolve sheet targets).
    let wb_rels = entries
        .get("xl/_rels/workbook.xml.rels")
        .map(|d| Relationships::parse(d))
        .transpose()?
        .unwrap_or_else(Relationships::new);

    // Resolve sheet paths from workbook relationships.
    let sheet_count = workbook_xml.sheets.len();
    let mut known_dynamic: HashSet<String> = HashSet::with_capacity(sheet_count * 4);
    // A theme is preserved verbatim so its fonts, effects and anything else
    // survive the round trip; only one byte-identical to the theme the writer
    // generates from its colours is consumed, since the writer regenerates it.
    if let (Some(colors), Some(data)) = (&theme_colors, entries.get(THEME_PATH))
        && *data == theme::generate_theme_xml(Some(colors))
    {
        known_dynamic.insert(THEME_PATH.to_owned());
    }
    let mut sheet_targets: Vec<(String, String, SheetState)> = Vec::with_capacity(sheet_count);

    for sheet_info in &workbook_xml.sheets {
        let rel = wb_rels.get_by_id(&sheet_info.r_id).ok_or_else(|| {
            cold_path();
            warn!(
                "could not resolve sheet target for rId: {}",
                log_text(&sheet_info.r_id)
            );
            ModernXlsxError::MissingPart(format!(
                "Failed to resolve sheet '{}': relationship '{}' not found in xl/_rels/workbook.xml.rels",
                sheet_info.name, sheet_info.r_id
            ))
        })?;

        // Normalize target path: handle both relative ("worksheets/sheet1.xml")
        // and absolute ("/xl/worksheets/sheet1.xml") targets.
        let sheet_path = workbook_part_path(&rel.target);

        // A worksheet part is one sheet: rows sharing one would each parse it,
        // and the model cannot hold one part as two sheets. `known_dynamic`
        // holds every sheet path resolved so far (and at most the theme
        // besides), so only a failed insert scans for the earlier row, and
        // the read fails before any sheet .rels or worksheet is parsed.
        if !known_dynamic.insert(sheet_path.clone())
            && let Some((first, ..)) = sheet_targets
                .iter()
                .find(|(_, path, _)| *path == sheet_path)
        {
            cold_path();
            return Err(ModernXlsxError::Security(format!(
                "sheets {} and {} both name worksheet part {}; a worksheet part belongs to one sheet",
                log_text(first),
                log_text(&sheet_info.name),
                log_text(&sheet_path)
            )));
        }
        sheet_targets.push((sheet_info.name.clone(), sheet_path, sheet_info.state));
    }

    // Each sheet's .rels, parsed once for every resolver and the worksheet
    // parse (hyperlink r:id). The file itself is never consumed: it stays in
    // preserved_entries, keyed by the sheet's position (`collect_preserved`),
    // for the writer to merge.
    let mut log = LogDigest::default();
    let sheet_rels = sheet_targets
        .iter()
        .map(|(_name, sheet_path, _state)| {
            let sheet_rels_path = rels_path(sheet_path);
            match entries.get(&sheet_rels_path) {
                Some(data) => optional_part(&mut log, &sheet_rels_path, Relationships::parse(data)),
                None => Ok(None),
            }
        })
        .collect::<Result<Vec<_>>>()?;

    // The source content types, kept for the preserved entries; a malformed
    // part is logged as dropped (it is never preserved: the writer generates
    // [Content_Types].xml) and the writer's own rules type those entries.
    let content_types = match entries.get(CONTENT_TYPES_PATH) {
        Some(data) => match ContentTypes::parse(data) {
            Ok(ct) => Some(ct),
            Err(e @ ModernXlsxError::Security(_)) => return Err(e),
            Err(e) => {
                cold_path();
                log.note(
                    Level::Warn,
                    "malformed content types part dropped",
                    format_args!("{CONTENT_TYPES_PATH}: {e}"),
                );
                None
            }
        },
        None => None,
    };

    Ok(ReaderContext {
        entries,
        workbook_xml,
        sst,
        styles,
        budget,
        doc_properties,
        theme_colors,
        calc_chain,
        wb_rels,
        sheet_targets,
        sheet_rels,
        known_dynamic,
        content_types,
        log,
    })
}

/// Applies the optional-part policy to the parse result of the part at `path`:
/// a malformed part is logged and kept verbatim, and a security or size
/// limit error stays fatal.
///
/// `Ok(Some)` when it parsed. A malformed part is counted in `log` (one
/// summary warning per read, naming the first such part and its error) and
/// gives `Ok(None)`: the caller then leaves `path` out of `known_dynamic`, so
/// [`collect_preserved`] keeps its bytes verbatim.
/// [`ModernXlsxError::Security`] and [`ModernXlsxError::SizeLimit`] (the
/// read's [`ExpansionBudget`]) stay fatal.
fn optional_part<T>(log: &mut LogDigest, path: &str, parsed: Result<T>) -> Result<Option<T>> {
    match parsed {
        Ok(value) => Ok(Some(value)),
        Err(e @ (ModernXlsxError::Security(_) | ModernXlsxError::SizeLimit(_))) => Err(e),
        Err(e) => {
            cold_path();
            log.note(
                Level::Warn,
                "malformed part preserved verbatim",
                format_args!("{path}: {e}"),
            );
            Ok(None)
        }
    }
}

/// The parts one resolver has read, keyed by resolved ZIP path, each with the
/// owner that read it, so that a part is parsed at most once per read and
/// what it gives enters the model once, for that owner. A crafted package can
/// name one part from thousands of relationships or owners, and parsing it
/// (or copying what it gave) for each would cost references x part size, far
/// past what the ZIP limits bound; honest packages never share these parts.
///
/// An owner is what names a part (a sheet, a drawing, a pivot cache
/// definition or the workbook), numbered by the resolver in the order it
/// visits them. A part is recorded once read, whether it parsed or not, so a
/// malformed part is not parsed again; a target missing from the package is
/// not recorded.
///
/// `ReaderContext::known_dynamic` cannot serve here: it is one set per read
/// rather than per resolver, and it never holds a malformed part, which would
/// then be parsed again for every relationship naming it.
struct PartLedger {
    /// Each part read so far, with the owner that read it.
    readers: HashMap<String, usize>,
}

/// What a resolver does with a relationship target ([`PartLedger::visit`]).
enum Visit {
    /// Not read before: read the part, then [`PartLedger::record`] it.
    First,
    /// Read before by the same owner: skipped, and counted in the log digest.
    Repeat,
    /// Read before by an earlier owner, which keeps what it gave: skipped,
    /// and counted in the log digest under a kind of its own.
    Shared,
}

impl PartLedger {
    fn new() -> Self {
        Self {
            readers: HashMap::new(),
        }
    }

    /// Says what `owner` does with the part at `path`.
    fn visit(&self, path: &str, owner: usize) -> Visit {
        match self.readers.get(path) {
            None => Visit::First,
            Some(&reader) if reader == owner => Visit::Repeat,
            Some(_) => Visit::Shared,
        }
    }

    /// Records that `owner` read the part at `path`.
    fn record(&mut self, path: String, owner: usize) {
        self.readers.insert(path, owner);
    }
}

/// The log digest kinds one per-sheet part type is counted under.
struct PartKinds {
    /// A part found in the package.
    found: &'static str,
    /// A target missing from the package.
    missing: &'static str,
    /// A part its sheet names more than once.
    repeated: &'static str,
    /// A part an earlier sheet read.
    shared: &'static str,
}

/// Resolve the parts of type `rel_type` that each sheet's .rels points at.
///
/// Every part found is counted in the read's log digest under `kinds.found`
/// and parsed under the optional-part policy ([`optional_part`]): a parsed
/// part's items are appended to its sheet's list, so a sheet with several
/// such parts keeps all of them, and the part is consumed. A target missing
/// from the package is counted under `kinds.missing`.
///
/// Each part is parsed once per read, for the first sheet naming it
/// ([`PartLedger`]): that sheet naming it again is counted under
/// `kinds.repeated`, and a later sheet naming it under `kinds.shared`; either
/// relationship is skipped and gives its sheet nothing. The result is
/// parallel to `ctx.sheet_targets`.
///
/// Only this shell is generic (a copy per caller's item type and parser):
/// the walk itself ([`walk_sheet_parts`]) is one non-generic function,
/// compiled once and shared by every part type.
fn resolve_per_sheet<T>(
    ctx: &mut ReaderContext,
    rel_type: &str,
    kinds: &PartKinds,
    mut parse: impl FnMut(&[u8]) -> Result<Vec<T>>,
) -> Result<Vec<Vec<T>>> {
    let mut per_sheet: Vec<Vec<T>> = std::iter::repeat_with(Vec::new)
        .take(ctx.sheet_targets.len())
        .collect();
    walk_sheet_parts(ctx, rel_type, kinds, &mut |sheet, data| {
        per_sheet[sheet].extend(parse(data)?);
        Ok(())
    })?;
    Ok(per_sheet)
}

/// The walk of [`resolve_per_sheet`]: `read(sheet, data)` parses one part
/// and appends its items to `sheet`'s list, and appends nothing when the
/// part does not parse.
fn walk_sheet_parts(
    ctx: &mut ReaderContext,
    rel_type: &str,
    kinds: &PartKinds,
    read: &mut dyn FnMut(usize, &[u8]) -> Result<()>,
) -> Result<()> {
    let mut ledger = PartLedger::new();

    for (i, ((_name, sheet_path, _state), ws_rels)) in
        ctx.sheet_targets.iter().zip(&ctx.sheet_rels).enumerate()
    {
        let Some(ws_rels) = ws_rels else {
            continue;
        };
        for rel in ws_rels.find_by_type(rel_type) {
            let part_path = resolve_target(sheet_path, &rel.target);

            match ledger.visit(&part_path, i) {
                Visit::First => {}
                Visit::Repeat => {
                    ctx.log.note(Level::Warn, kinds.repeated, &part_path);
                    continue;
                }
                Visit::Shared => {
                    ctx.log.note(Level::Warn, kinds.shared, &part_path);
                    continue;
                }
            }

            let Some(data) = ctx.entries.get(&part_path) else {
                ctx.log.note(Level::Warn, kinds.missing, &part_path);
                continue;
            };
            ctx.log.note(Level::Debug, kinds.found, &part_path);
            if optional_part(&mut ctx.log, &part_path, read(i, data))?.is_some() {
                ctx.known_dynamic.insert(part_path.clone());
            }
            ledger.record(part_path, i);
        }
    }

    Ok(())
}

/// Resolve comments for each sheet from worksheet .rels files, counting each
/// comment's copy of its author in `budget`. A part that does not parse
/// gives its comments nothing, so what it counted is given back.
fn resolve_comments(
    ctx: &mut ReaderContext,
    budget: &mut ExpansionBudget,
) -> Result<Vec<Vec<comments::Comment>>> {
    resolve_per_sheet(
        ctx,
        REL_COMMENTS,
        &PartKinds {
            found: "comments part found",
            missing: "comments part not found",
            repeated: "comments part named again, skipped",
            shared: "comments part shared with an earlier sheet, skipped",
        },
        |data| {
            let spent = budget.spent;
            comments::parse_comments_within(data, budget).inspect_err(|_| budget.spent = spent)
        },
    )
}

/// Resolve table definitions for each sheet from worksheet .rels files.
fn resolve_tables(ctx: &mut ReaderContext) -> Result<Vec<Vec<TableDefinition>>> {
    resolve_per_sheet(
        ctx,
        REL_TABLE,
        &PartKinds {
            found: "table part found",
            missing: "table part not found",
            repeated: "table part named again, skipped",
            shared: "table part shared with an earlier sheet, skipped",
        },
        |d| TableDefinition::parse(d).map(|t| vec![t]),
    )
}

/// Resolve chart definitions for each sheet.
///
/// Charts require a two-step lookup: sheet .rels → drawing rel → drawing.xml
/// (which contains chart anchors) → drawing .rels → chart XML. A drawing whose
/// XML, .rels or any chart is malformed contributes no chart, and the drawing,
/// its .rels and all its charts stay in `preserved_entries`.
///
/// A chart the model cannot represent comes back with `raw_xml` set. Every
/// chart's .rels, modelled or verbatim, is consumed into its `raw_rels`, so the
/// writer can move both together when it renumbers the charts.
///
/// The drawing and its .rels are consumed only when every anchor in it is a
/// chart. A drawing that also holds pictures, shapes or `mc:AlternateContent`
/// stays in `preserved_entries` beside its charts, which still enter the
/// model; the writer regenerates the chart frames and keeps the rest.
///
/// Each drawing and each chart part is parsed once per read ([`PartLedger`]).
/// A drawing goes to the first sheet naming it: that sheet naming it again,
/// or a later sheet naming it, is counted and skipped. A chart part goes to
/// the first anchor naming it: a later anchor naming it, in the same drawing
/// or another, is counted and skipped with that anchor, and the rest of its
/// drawing is handled as above.
///
/// Each series and axis a chart keeps is counted in `ctx.budget`; a drawing
/// that contributes no chart gives back what its charts counted.
fn resolve_charts(ctx: &mut ReaderContext) -> Result<Vec<Vec<charts::WorksheetChart>>> {
    let mut sheet_charts = vec![Vec::new(); ctx.sheet_targets.len()];
    // Owned by sheets.
    let mut drawings = PartLedger::new();
    // Owned by drawings, numbered in reading order.
    let mut chart_parts = PartLedger::new();
    let mut drawings_read = 0usize;

    for (i, ((_name, sheet_path, _state), ws_rels)) in
        ctx.sheet_targets.iter().zip(&ctx.sheet_rels).enumerate()
    {
        let Some(ws_rels) = ws_rels else {
            continue;
        };
        // Find drawing relationships.
        for rel in ws_rels.find_by_type(REL_DRAWING) {
            let drawing_path = resolve_target(sheet_path, &rel.target);

            match drawings.visit(&drawing_path, i) {
                Visit::First => {}
                Visit::Repeat => {
                    ctx.log.note(
                        Level::Warn,
                        "drawing part named again, skipped",
                        &drawing_path,
                    );
                    continue;
                }
                Visit::Shared => {
                    ctx.log.note(
                        Level::Warn,
                        "drawing part shared with an earlier sheet, skipped",
                        &drawing_path,
                    );
                    continue;
                }
            }

            let Some(drawing_data) = ctx.entries.get(&drawing_path) else {
                ctx.log
                    .note(Level::Warn, "drawing part not found", &drawing_path);
                continue;
            };
            ctx.log
                .note(Level::Debug, "drawing part found", &drawing_path);
            let drawing = drawings_read;
            drawings_read += 1;

            // The drawing's charts when they enter the model, else `None`.
            let stored = ctx.budget.stored;
            let read = 'drawing: {
                let Some(anchors) = optional_part(
                    &mut ctx.log,
                    &drawing_path,
                    charts::parse_drawing_anchors(drawing_data),
                )?
                else {
                    break 'drawing None;
                };

                // Resolve each chart rId via drawing .rels.
                let drawing_rels_path = rels_path(&drawing_path);
                let Some(dr) = ctx.entries.get(&drawing_rels_path) else {
                    break 'drawing None;
                };
                let Some(drawing_rels) =
                    optional_part(&mut ctx.log, &drawing_rels_path, Relationships::parse(dr))?
                else {
                    break 'drawing None;
                };

                let chart_only = anchors.total > 0 && anchors.charts.len() == anchors.total;
                let mut drawing_charts = Vec::with_capacity(anchors.charts.len());
                let mut chart_paths = Vec::with_capacity(anchors.charts.len());
                let mut intact = true;

                for (anchor, chart_r_id) in &anchors.charts {
                    let Some(chart_rel) = drawing_rels.get_by_id(chart_r_id) else {
                        continue;
                    };
                    let chart_path = resolve_target(&drawing_path, &chart_rel.target);
                    match chart_parts.visit(&chart_path, drawing) {
                        Visit::First => {}
                        Visit::Repeat => {
                            ctx.log.note(
                                Level::Warn,
                                "chart part named again, skipped",
                                &chart_path,
                            );
                            continue;
                        }
                        Visit::Shared => {
                            ctx.log.note(
                                Level::Warn,
                                "chart part shared with an earlier drawing, skipped",
                                &chart_path,
                            );
                            continue;
                        }
                    }

                    let Some(chart_data) = ctx.entries.get(&chart_path) else {
                        ctx.log
                            .note(Level::Warn, "chart part not found", &chart_path);
                        continue;
                    };
                    ctx.log.note(Level::Debug, "chart part found", &chart_path);
                    chart_parts.record(chart_path.clone(), drawing);
                    match optional_part(
                        &mut ctx.log,
                        &chart_path,
                        charts::ChartData::parse_within(chart_data, &mut ctx.budget),
                    )? {
                        Some(mut chart) => {
                            // Every chart, modelled or verbatim, keeps its .rels,
                            // so the writer places both at the chart's new path.
                            let chart_rels_path = rels_path(&chart_path);
                            if let Some(rels) = ctx.entries.get(&chart_rels_path) {
                                match std::str::from_utf8(rels) {
                                    Ok(rels) => {
                                        chart.raw_rels = Some(rels.to_owned());
                                        chart_paths.push(chart_rels_path);
                                    }
                                    Err(e) => ctx.log.note(
                                        Level::Warn,
                                        "chart .rels not UTF-8, preserved verbatim",
                                        format_args!("{chart_rels_path}: {e}"),
                                    ),
                                }
                            }
                            drawing_charts.push(charts::WorksheetChart {
                                chart,
                                anchor: anchor.clone(),
                            });
                            chart_paths.push(chart_path);
                        }
                        None => intact = false,
                    }
                }

                if !intact {
                    break 'drawing None;
                }
                ctx.known_dynamic.extend(chart_paths);
                if chart_only {
                    ctx.known_dynamic.insert(drawing_path.clone());
                    ctx.known_dynamic.insert(drawing_rels_path);
                }
                Some(drawing_charts)
            };

            if let Some(charts) = read {
                sheet_charts[i].extend(charts);
            } else {
                ctx.budget.stored = stored;
            }
            drawings.record(drawing_path, i);
        }
    }

    Ok(sheet_charts)
}

/// Pivot tables per sheet plus the workbook's pivot cache definitions and records.
struct PivotFamily {
    sheet_tables: Vec<Vec<PivotTableData>>,
    caches: Vec<PivotCacheDefinitionData>,
    records: Vec<PivotCacheRecordsData>,
}

impl PivotFamily {
    fn empty(sheet_count: usize) -> Self {
        Self {
            sheet_tables: vec![Vec::new(); sheet_count],
            caches: Vec::new(),
            records: Vec::new(),
        }
    }
}

/// Resolve pivot tables from worksheet .rels, and pivot cache definitions and
/// records from workbook .rels.
///
/// The writer pairs records with caches by position and pivot tables point at
/// caches by `cacheId`, so the family enters the model whole or not at all: one
/// malformed pivot table, cache definition, cache .rels or records part keeps
/// every pivot table, cache and record out, and every pivot part stays in
/// `preserved_entries`.
///
/// Records belong to the cache definition whose .rels names them, and the
/// model pairs them by position (`pivot_cache_records[i]` with
/// `pivot_caches[i]`): the definitions with records come first and those
/// without (saved with `saveData="0"`) follow, each group in workbook .rels
/// order. A cache read with a `cacheId` keeps it wherever it lands. A
/// definition whose .rels names two records parts that are both read
/// (present, parsed and not owned by an earlier definition) keeps the family
/// out.
///
/// Each part is parsed once per read ([`PartLedger`]), for the first owner
/// naming it. A sheet naming a pivot table again, the workbook naming a cache
/// definition again and a cache definition naming a records part again are
/// counted and skipped, and so are a later sheet naming a pivot table an
/// earlier sheet read and a later cache definition naming a records part an
/// earlier definition read. Records read after such a skip keep the family
/// out. A pivot table bound to the `cacheId` of a skipped definition repeat,
/// which no kept cache carries, also keeps the family out.
///
/// The by-id associations stay as they are: pivot tables name their cache by
/// `cacheId`, and a cache definition several pivot tables use is parsed once
/// and held once in `caches`, never copied per table.
fn resolve_pivots(ctx: &mut ReaderContext) -> Result<PivotFamily> {
    let mut family = PivotFamily::empty(ctx.sheet_targets.len());
    // Consumed only once the whole family has parsed.
    let mut parsed_paths: Vec<String> = Vec::new();
    let mut intact = true;
    // Owned by sheets.
    let mut tables = PartLedger::new();

    for (i, ((_name, sheet_path, _state), ws_rels)) in
        ctx.sheet_targets.iter().zip(&ctx.sheet_rels).enumerate()
    {
        let Some(ws_rels) = ws_rels else {
            continue;
        };
        for rel in ws_rels.find_by_type(REL_PIVOT_TABLE) {
            let pivot_path = resolve_target(sheet_path, &rel.target);

            match tables.visit(&pivot_path, i) {
                Visit::First => {}
                Visit::Repeat => {
                    ctx.log.note(
                        Level::Warn,
                        "pivot table part named again, skipped",
                        &pivot_path,
                    );
                    continue;
                }
                Visit::Shared => {
                    ctx.log.note(
                        Level::Warn,
                        "pivot table part shared with an earlier sheet, skipped",
                        &pivot_path,
                    );
                    continue;
                }
            }

            let Some(pivot_data) = ctx.entries.get(&pivot_path) else {
                ctx.log
                    .note(Level::Warn, "pivot table part not found", &pivot_path);
                continue;
            };
            ctx.log
                .note(Level::Debug, "pivot table part found", &pivot_path);
            if let Some(table) =
                optional_part(&mut ctx.log, &pivot_path, PivotTableData::parse(pivot_data))?
            {
                family.sheet_tables[i].push(table);
                // Its .rels (to the cache definition) is rebuilt
                // by the writer under the table's new part name.
                let table_rels_path = rels_path(&pivot_path);
                if ctx.entries.contains_key(&table_rels_path) {
                    parsed_paths.push(table_rels_path);
                }
                parsed_paths.push(pivot_path.clone());
            } else {
                intact = false;
            }
            tables.record(pivot_path, i);
        }
    }

    // Each cache's workbook.xml `cacheId`, found through its relationship Id. One
    // map, built once: a scan of the caches per relationship would cost
    // relationships x caches, and both counts come from the file. The first
    // `<pivotCache>` with an `r:id` wins.
    let mut cache_id_by_r_id: HashMap<&str, u32> =
        HashMap::with_capacity(ctx.workbook_xml.pivot_caches.len());
    for cache in &ctx.workbook_xml.pivot_caches {
        cache_id_by_r_id
            .entry(cache.r_id.as_str())
            .or_insert(cache.cache_id);
    }
    let cache_targets: Vec<(String, Option<u32>)> = ctx
        .wb_rels
        .find_by_type(REL_PIVOT_CACHE_DEF)
        .map(|rel| {
            let cache_id = cache_id_by_r_id.get(rel.id.as_str()).copied();
            (workbook_part_path(&rel.target), cache_id)
        })
        .collect();

    // The workbook is the one owner of the cache definitions.
    let mut cache_parts = PartLedger::new();
    // Owned by cache definitions, numbered by position in `cache_targets`.
    let mut records_parts = PartLedger::new();
    // The `cacheId`s of skipped definition repeats; a pivot table may name one.
    let mut skipped_cache_ids: HashSet<u32> = HashSet::new();
    // Set once a cache definition is left without a records part it named
    // (one an earlier definition read); records read after it keep the
    // family out.
    let mut records_gap = false;
    // The definitions without records, in workbook .rels order; they follow
    // the definitions with records in `family.caches`.
    let mut caches_without_records: Vec<PivotCacheDefinitionData> = Vec::new();

    for (cache, (cache_path, cache_id)) in cache_targets.into_iter().enumerate() {
        if !matches!(cache_parts.visit(&cache_path, 0), Visit::First) {
            ctx.log.note(
                Level::Warn,
                "pivot cache definition part named again, skipped",
                &cache_path,
            );
            skipped_cache_ids.extend(cache_id);
            continue;
        }
        let Some(cache_data) = ctx.entries.get(&cache_path) else {
            ctx.log.note(
                Level::Warn,
                "pivot cache definition part not found",
                &cache_path,
            );
            continue;
        };
        ctx.log.note(
            Level::Debug,
            "pivot cache definition part found",
            &cache_path,
        );
        cache_parts.record(cache_path.clone(), 0);
        let mut definition = optional_part(
            &mut ctx.log,
            &cache_path,
            PivotCacheDefinitionData::parse(cache_data),
        )?
        .map(|cache| PivotCacheDefinitionData { cache_id, ..cache });
        if definition.is_none() {
            intact = false;
        }
        // The records part this definition's own .rels names.
        let mut own_records: Option<PivotCacheRecordsData> = None;

        // Look for corresponding records via cache definition .rels.
        let cache_rels_path = rels_path(&cache_path);
        if let Some(cr_data) = ctx.entries.get(&cache_rels_path) {
            match optional_part(
                &mut ctx.log,
                &cache_rels_path,
                Relationships::parse(cr_data),
            )? {
                Some(cache_rels) => {
                    for rel in cache_rels.find_by_type(REL_PIVOT_CACHE_REC) {
                        let records_path = resolve_target(&cache_path, &rel.target);

                        match records_parts.visit(&records_path, cache) {
                            Visit::First => {}
                            Visit::Repeat => {
                                ctx.log.note(
                                    Level::Warn,
                                    "pivot cache records part named again, skipped",
                                    &records_path,
                                );
                                continue;
                            }
                            Visit::Shared => {
                                ctx.log.note(
                                    Level::Warn,
                                    "pivot cache records part shared with an earlier cache definition, skipped",
                                    &records_path,
                                );
                                records_gap = true;
                                continue;
                            }
                        }

                        let Some(rec_data) = ctx.entries.get(&records_path) else {
                            ctx.log.note(
                                Level::Warn,
                                "pivot cache records part not found",
                                &records_path,
                            );
                            continue;
                        };
                        ctx.log.note(
                            Level::Debug,
                            "pivot cache records part found",
                            &records_path,
                        );
                        if let Some(records) = optional_part(
                            &mut ctx.log,
                            &records_path,
                            PivotCacheRecordsData::parse(rec_data),
                        )? {
                            if records_gap {
                                intact = false;
                            }
                            if own_records.is_some() {
                                // A cache holds one records part: which one
                                // belongs to it is not known.
                                ctx.log.note(
                                    Level::Warn,
                                    "pivot cache definition names a second records part, pivot family kept out",
                                    &records_path,
                                );
                                intact = false;
                            }
                            own_records = Some(records);
                            parsed_paths.push(records_path.clone());
                        } else {
                            intact = false;
                        }
                        records_parts.record(records_path, cache);
                    }
                    // Every other relationship (such as the externalLinkPath
                    // a source in another workbook names by `r:id`) is kept
                    // on the definition, which the writer gives back its id.
                    if let Some(definition) = definition.as_mut() {
                        definition.relationships = cache_rels
                            .iter()
                            .filter(|rel| rel.rel_type != REL_PIVOT_CACHE_REC)
                            .map(CacheRelationship::from_relationship)
                            .collect();
                    }
                    parsed_paths.push(cache_rels_path);
                }
                None => intact = false,
            }
        }
        match (definition, own_records) {
            (Some(cache), Some(records)) => {
                family.caches.push(cache);
                family.records.push(records);
            }
            (Some(cache), None) => caches_without_records.push(cache),
            (None, _) => {}
        }
        parsed_paths.push(cache_path);
    }
    family.caches.append(&mut caches_without_records);

    // A pivot table bound to a skipped repeat's `cacheId` that no kept cache
    // carries would name a cache the model lacks, and the writer rejects it:
    // keep the family out. A pure repeat leaves no such table. Sets,
    // not scans: tables, caches and repeats all come from the file.
    if !skipped_cache_ids.is_empty() {
        let kept: HashSet<u32> = family
            .caches
            .iter()
            .filter_map(|cache| cache.cache_id)
            .collect();
        let orphaned = family.sheet_tables.iter().flatten().any(|table| {
            skipped_cache_ids.contains(&table.cache_id) && !kept.contains(&table.cache_id)
        });
        if orphaned {
            intact = false;
        }
    }

    if !intact {
        debug!("pivot family kept out of the model; every pivot part is preserved verbatim");
        return Ok(PivotFamily::empty(ctx.sheet_targets.len()));
    }
    ctx.known_dynamic.extend(parsed_paths);
    Ok(family)
}

/// Resolve threaded comments for each sheet from worksheet .rels files,
/// counting each comment and mention kept in `ctx.budget`. A part that does
/// not parse gives its comments nothing, so what it counted is given back.
fn resolve_threaded_comments(ctx: &mut ReaderContext) -> Result<Vec<Vec<ThreadedCommentData>>> {
    let mut budget = ctx.budget;
    let resolved = resolve_per_sheet(
        ctx,
        REL_THREADED_COMMENTS,
        &PartKinds {
            found: "threaded comments part found",
            missing: "threaded comments part not found",
            repeated: "threaded comments part named again, skipped",
            shared: "threaded comments part shared with an earlier sheet, skipped",
        },
        |data| {
            let stored = budget.stored;
            threaded_comments::parse_threaded_comments_within(data, &mut budget)
                .inspect_err(|_| budget.stored = stored)
        },
    );
    ctx.budget = budget;
    resolved
}

/// Resolve persons list from the well-known path [`PERSONS_PATH`].
fn resolve_persons(ctx: &mut ReaderContext) -> Result<Vec<PersonData>> {
    let Some(persons_data) = ctx.entries.get(PERSONS_PATH) else {
        return Ok(Vec::new());
    };
    debug!("parsing persons from: {PERSONS_PATH}");
    let Some(persons) = optional_part(
        &mut ctx.log,
        PERSONS_PATH,
        threaded_comments::parse_persons(persons_data),
    )?
    else {
        return Ok(Vec::new());
    };
    ctx.known_dynamic.insert(PERSONS_PATH.to_owned());
    Ok(persons)
}

/// Resolve slicers per worksheet via REL_SLICER in each sheet's .rels.
fn resolve_slicers(ctx: &mut ReaderContext) -> Result<Vec<Vec<SlicerData>>> {
    resolve_per_sheet(
        ctx,
        REL_SLICER,
        &PartKinds {
            found: "slicer part found",
            missing: "slicer part not found",
            repeated: "slicer part named again, skipped",
            shared: "slicer part shared with an earlier sheet, skipped",
        },
        slicers::parse_slicers,
    )
}

/// Resolve slicer cache definitions from workbook .rels; a part the
/// workbook names again is counted and skipped ([`PartLedger`]).
fn resolve_slicer_caches(ctx: &mut ReaderContext) -> Result<Vec<SlicerCacheData>> {
    let mut caches: Vec<SlicerCacheData> = Vec::new();

    let cache_targets: Vec<String> = ctx
        .wb_rels
        .find_by_type(REL_SLICER_CACHE)
        .map(|rel| workbook_part_path(&rel.target))
        .collect();

    // The workbook is the one owner of the caches.
    let mut cache_parts = PartLedger::new();
    for cache_path in cache_targets {
        if !matches!(cache_parts.visit(&cache_path, 0), Visit::First) {
            ctx.log.note(
                Level::Warn,
                "slicer cache part named again, skipped",
                &cache_path,
            );
            continue;
        }
        if let Some(cache_data) = ctx.entries.get(&cache_path) {
            ctx.log
                .note(Level::Debug, "slicer cache part found", &cache_path);
            cache_parts.record(cache_path.clone(), 0);
            if let Some(cache) = optional_part(
                &mut ctx.log,
                &cache_path,
                SlicerCacheData::parse(cache_data),
            )? {
                caches.push(cache);
                ctx.known_dynamic.insert(cache_path);
            }
        } else {
            ctx.log
                .note(Level::Warn, "slicer cache part not found", &cache_path);
        }
    }

    Ok(caches)
}

/// Resolve timelines per worksheet via REL_TIMELINE in each sheet's .rels.
fn resolve_timelines(ctx: &mut ReaderContext) -> Result<Vec<Vec<TimelineData>>> {
    resolve_per_sheet(
        ctx,
        REL_TIMELINE,
        &PartKinds {
            found: "timeline part found",
            missing: "timeline part not found",
            repeated: "timeline part named again, skipped",
            shared: "timeline part shared with an earlier sheet, skipped",
        },
        timelines::parse_timelines,
    )
}

/// Resolve timeline cache definitions from workbook .rels; a part the
/// workbook names again is counted and skipped ([`PartLedger`]).
fn resolve_timeline_caches(ctx: &mut ReaderContext) -> Result<Vec<TimelineCacheData>> {
    let mut caches: Vec<TimelineCacheData> = Vec::new();

    let cache_targets: Vec<String> = ctx
        .wb_rels
        .find_by_type(REL_TIMELINE_CACHE)
        .map(|rel| workbook_part_path(&rel.target))
        .collect();

    // The workbook is the one owner of the caches.
    let mut cache_parts = PartLedger::new();
    for cache_path in cache_targets {
        if !matches!(cache_parts.visit(&cache_path, 0), Visit::First) {
            ctx.log.note(
                Level::Warn,
                "timeline cache part named again, skipped",
                &cache_path,
            );
            continue;
        }
        if let Some(cache_data) = ctx.entries.get(&cache_path) {
            ctx.log
                .note(Level::Debug, "timeline cache part found", &cache_path);
            cache_parts.record(cache_path.clone(), 0);
            if let Some(cache) = optional_part(
                &mut ctx.log,
                &cache_path,
                TimelineCacheData::parse(cache_data),
            )? {
                caches.push(cache);
                ctx.known_dynamic.insert(cache_path);
            }
        } else {
            ctx.log
                .note(Level::Warn, "timeline cache part not found", &cache_path);
        }
    }

    Ok(caches)
}

/// Collect all ZIP entries that were not parsed into preserved_entries.
/// No spelling of [`CONTENT_TYPES_PATH`] is one: the writer always generates
/// that part.
///
/// Takes ownership of entries via `drain()` to avoid cloning large byte
/// vectors (e.g. embedded images, charts).
///
/// Worksheet `.rels` are keyed by the sheet's 1-based position
/// (`xl/worksheets/_rels/sheet{n}.xml.rels`), the name the writer looks them
/// up by, whenever the sheet's source part is directly in `xl/worksheets/`:
/// a sheet whose part is `sheet2.xml` at position 1 has its `.rels` under
/// `sheet1.xml.rels`. Each sheet has a source part of its own
/// ([`parse_common`] rejects two on one part), so each `.rels` moves once.
/// The `.rels` of a sheet whose part is in another folder keeps its source
/// path.
fn collect_preserved(ctx: &mut ReaderContext) -> BTreeMap<String, Vec<u8>> {
    let mut preserved = BTreeMap::new();
    for (path, data) in ctx.entries.drain() {
        if !KNOWN_STATIC_PATHS.contains(&path.as_str())
            && !ctx.known_dynamic.contains(&path)
            && !path.eq_ignore_ascii_case(CONTENT_TYPES_PATH)
        {
            ctx.log
                .note(Level::Debug, "unparsed ZIP entry preserved", &path);
            preserved.insert(path, data);
        }
    }

    // Every worksheet .rels is taken out before any is put back, so a .rels
    // already at another sheet's position name is never overwritten first.
    let in_worksheets = |sheet_path: &str| {
        sheet_path
            .rsplit_once('/')
            .is_some_and(|(dir, _)| dir == "xl/worksheets")
    };
    let mut sheet_rels: HashMap<&str, Vec<u8>> = HashMap::new();
    for (_name, sheet_path, _state) in &ctx.sheet_targets {
        if in_worksheets(sheet_path)
            && let Some(data) = preserved.remove(&rels_path(sheet_path))
        {
            sheet_rels.insert(sheet_path.as_str(), data);
        }
    }
    for (i, (_name, sheet_path, _state)) in ctx.sheet_targets.iter().enumerate() {
        if let Some(data) = sheet_rels.remove(sheet_path.as_str()) {
            preserved.insert(format!("xl/worksheets/_rels/sheet{}.xml.rels", i + 1), data);
        }
    }
    preserved
}

/// The content type the source `[Content_Types].xml` declares for each path
/// of `preserved`: its Override (`/{path}`), else the Default for its
/// extension, matched ignoring ASCII case. A `.rels` is left out (it is
/// always a relationships part, and the writer types it itself). Empty when
/// the package has no parsable `[Content_Types].xml`.
fn preserved_content_types(
    content_types: Option<&ContentTypes>,
    preserved: &BTreeMap<String, Vec<u8>>,
) -> BTreeMap<String, String> {
    let Some(content_types) = content_types else {
        return BTreeMap::new();
    };
    let defaults: HashMap<String, &String> = content_types
        .defaults
        .iter()
        .map(|(ext, content_type)| (ext.to_ascii_lowercase(), content_type))
        .collect();
    // `preserved` iterates in key order: inserting in turn builds the map
    // without the stable sort a `collect` into a `BTreeMap` runs first.
    let mut declared = BTreeMap::new();
    for path in preserved.keys().filter(|path| !path.ends_with(".rels")) {
        let content_type = content_types
            .overrides
            .get(&format!("/{path}"))
            .or_else(|| {
                let file_name = path.rsplit_once('/').map_or(path.as_str(), |(_, f)| f);
                let (_, ext) = file_name.rsplit_once('.')?;
                defaults.get(&ext.to_ascii_lowercase()).copied()
            });
        if let Some(content_type) = content_type {
            declared.insert(path.clone(), content_type.clone());
        }
    }
    declared
}

/// The ZIP path of a workbook relationship target: relative to `xl/`
/// (`theme/theme1.xml`), or package-absolute with a leading `/`.
#[inline]
fn workbook_part_path(target: &str) -> String {
    match target.strip_prefix('/') {
        Some(absolute) => absolute.to_owned(),
        None => format!("xl/{target}"),
    }
}

/// The external link part paths in workbook.xml `<externalReferences>` order;
/// an `r:id` the workbook rels do not resolve is skipped and counted in the
/// read's log digest.
fn external_reference_paths(ctx: &mut ReaderContext) -> Vec<String> {
    ctx.workbook_xml
        .external_references
        .iter()
        .filter_map(|r_id| {
            let rel = ctx.wb_rels.get_by_id(r_id);
            if rel.is_none() {
                ctx.log.note(
                    Level::Warn,
                    "externalReference without a workbook relationship, skipped",
                    r_id,
                );
            }
            rel.map(|rel| workbook_part_path(&rel.target))
        })
        .collect()
}

// ---------------------------------------------------------------------------
// Public API — struct-based reader
// ---------------------------------------------------------------------------

/// Read an XLSX file from bytes using default security limits.
pub fn read_xlsx(data: &[u8]) -> Result<WorkbookData> {
    read_xlsx_with_options(data, &ZipSecurityLimits::default())
}

/// Read an XLSX file from bytes with custom ZIP security limits.
///
/// Fails with [`ModernXlsxError::SizeLimit`] when the text the read copies
/// from a table into every place naming it (shared strings into cells,
/// comment authors into comments, hyperlink targets into hyperlinks, and the
/// worksheet namespace declarations re-declared on each kept `<ext>`) passes
/// `limits.max_json_length` (`ExpansionBudget`): the model would not
/// serialize to JSON within that limit. It also fails with `SizeLimit` when
/// the records the read keeps in its lists (the workbook's defined names,
/// the entries and rich text runs of `xl/sharedStrings.xml`, the records of
/// `xl/styles.xml`, the series and axes of the charts, the threaded comments
/// and their mentions, then the worksheet list elements), each counted at
/// the size of its struct plus its strings (not the plain text of a shared string entry), pass the same limit in memory.
pub fn read_xlsx_with_options(data: &[u8], limits: &ZipSecurityLimits) -> Result<WorkbookData> {
    let mut ctx = parse_common(data, limits, None)?;
    let mut budget = ctx.budget;
    let sheet_comments = resolve_comments(&mut ctx, &mut budget)?;
    ctx.budget = budget;
    let sheet_tables = resolve_tables(&mut ctx)?;
    let sheet_charts = resolve_charts(&mut ctx)?;
    let PivotFamily {
        sheet_tables: sheet_pivots,
        caches: pivot_caches,
        records: pivot_cache_records,
    } = resolve_pivots(&mut ctx)?;
    let sheet_threaded = resolve_threaded_comments(&mut ctx)?;
    let persons = resolve_persons(&mut ctx)?;
    let sheet_slicers = resolve_slicers(&mut ctx)?;
    let slicer_caches = resolve_slicer_caches(&mut ctx)?;
    let sheet_timelines = resolve_timelines(&mut ctx)?;
    let timeline_caches = resolve_timeline_caches(&mut ctx)?;
    // The sheets count after the charts and threaded comments.
    budget = ctx.budget;

    // Parse each worksheet XML.
    // When the "parallel" feature is enabled, sheets are parsed concurrently.
    let mut sheets = parse_sheets(
        &ctx.entries,
        &ctx.sheet_targets,
        &ctx.sheet_rels,
        &ctx.sst,
        ctx.workbook_xml.date_system,
        &mut budget,
    )?;

    // Attach comments to their respective worksheets.
    for (sheet, comments) in sheets.iter_mut().zip(sheet_comments) {
        if !comments.is_empty() {
            sheet.worksheet.comments = comments;
        }
    }

    // Attach tables to their respective worksheets.
    for (sheet, tables) in sheets.iter_mut().zip(sheet_tables) {
        if !tables.is_empty() {
            sheet.worksheet.tables = tables;
        }
    }

    // Attach charts to their respective worksheets.
    for (sheet, charts_vec) in sheets.iter_mut().zip(sheet_charts) {
        if !charts_vec.is_empty() {
            sheet.worksheet.charts = charts_vec;
        }
    }

    // Attach pivot tables to their respective worksheets.
    for (sheet, pivots) in sheets.iter_mut().zip(sheet_pivots) {
        if !pivots.is_empty() {
            sheet.worksheet.pivot_tables = pivots;
        }
    }

    // Attach threaded comments to their respective worksheets.
    for (sheet, tc) in sheets.iter_mut().zip(sheet_threaded) {
        if !tc.is_empty() {
            sheet.worksheet.threaded_comments = tc;
        }
    }

    // Attach slicers to their respective worksheets.
    for (sheet, sl) in sheets.iter_mut().zip(sheet_slicers) {
        if !sl.is_empty() {
            sheet.worksheet.slicers = sl;
        }
    }

    // Attach timelines to their respective worksheets.
    for (sheet, tl) in sheets.iter_mut().zip(sheet_timelines) {
        if !tl.is_empty() {
            sheet.worksheet.timelines = tl;
        }
    }

    let external_references = external_reference_paths(&mut ctx);
    let preserved_entries = collect_preserved(&mut ctx);
    let preserved_content_types =
        preserved_content_types(ctx.content_types.as_ref(), &preserved_entries);

    Ok(WorkbookData {
        sheets,
        date_system: ctx.workbook_xml.date_system,
        styles: ctx.styles,
        defined_names: ctx.workbook_xml.defined_names,
        shared_strings: Some(ctx.sst),
        doc_properties: ctx.doc_properties,
        theme_colors: ctx.theme_colors,
        calc_chain: ctx.calc_chain,
        workbook_views: ctx.workbook_xml.workbook_views,
        protection: ctx.workbook_xml.protection,
        calc_pr: ctx.workbook_xml.calc_pr,
        file_version: ctx.workbook_xml.file_version,
        workbook_pr: ctx.workbook_xml.workbook_pr,
        pivot_caches,
        pivot_cache_records,
        persons,
        slicer_caches,
        timeline_caches,
        external_references,
        preserved_entries,
        preserved_content_types,
    })
}

// ---------------------------------------------------------------------------
// Public API — streaming JSON reader
// ---------------------------------------------------------------------------

/// Read an XLSX file and return the result directly as a JSON string.
///
/// Worksheet rows and cells are written straight to JSON while each sheet's
/// XML is parsed, with no `Row` or `Cell` values in between; the other fields
/// (styles, defined names, charts and the rest) go through `serde_json`. The
/// WASM bridge reads through this path (`read_xlsx_json_with_options`), and
/// `benches/read_write.rs` compares it with [`read_xlsx`] plus
/// `serde_json::to_string`.
///
/// The JSON output is compatible with `WorkbookData`'s serde format, so the
/// TypeScript side can use `JSON.parse()` to get the same structure.
pub fn read_xlsx_json(data: &[u8]) -> Result<String> {
    read_xlsx_json_with_options(data, &ZipSecurityLimits::default())
}

/// Read an XLSX file as JSON with custom ZIP security limits.
///
/// Fails with [`ModernXlsxError::SizeLimit`] when the JSON grows past
/// `limits.max_json_length`, or when the text the read copies from a table
/// into every place naming it (shared strings into cells, comment authors
/// into comments, hyperlink targets into hyperlinks, and the worksheet
/// namespace declarations re-declared on each kept `<ext>`) passes that
/// limit first (`ExpansionBudget`), or when the records the read keeps in its
/// lists (the defined names, shared strings, styles, chart series and axes
/// and threaded comments, then the worksheet list elements) pass it in
/// memory, as [`read_xlsx_with_options`] says. The JSON is checked after each shared
/// string a `<v>` names and each cell, row, comment, hyperlink, kept
/// extension and sheet, and after each workbook-level field, so a copy
/// written escaped (up to 6 bytes for a control character) passes the limit
/// by at most the item written last (one `<v>`'s entry, a cell's other
/// parts, a comment, hyperlink, extension or field) before the read fails.
pub fn read_xlsx_json_with_options(data: &[u8], limits: &ZipSecurityLimits) -> Result<String> {
    build_json_from_context(
        parse_common(data, limits, None)?,
        data.len(),
        limits.max_json_length,
    )
}

/// Read an XLSX file (possibly encrypted) from bytes with a password.
///
/// If the file is a plain ZIP (not encrypted), the password is ignored
/// and the file is read normally. If the file is OLE2-encrypted and the
/// password is empty, a descriptive error is returned.
///
/// The lite build (no `encryption` feature) cannot decrypt: an OLE2-encrypted file
/// is `PasswordProtected` even when a password is given.
pub fn read_xlsx_json_with_password(data: &[u8], password: &str) -> Result<String> {
    read_xlsx_json_with_password_and_options(data, password, &ZipSecurityLimits::default())
}

/// Like [`read_xlsx_json_with_password`], with custom security and size limits.
/// The lite build cannot decrypt, so an encrypted file is `PasswordProtected`.
pub fn read_xlsx_json_with_password_and_options(
    data: &[u8],
    password: &str,
    limits: &ZipSecurityLimits,
) -> Result<String> {
    let pw = if password.is_empty() {
        None
    } else {
        Some(password)
    };
    build_json_from_context(
        parse_common(data, limits, pw)?,
        data.len(),
        limits.max_json_length,
    )
}

/// Upper bound (64 MiB) on the JSON buffer reserved up front.
const JSON_PREALLOC_CAP: usize = 64 * 1024 * 1024;

/// Initial capacity for the JSON output: 15x the input, held between 4 KiB and
/// [`JSON_PREALLOC_CAP`]. The 15x is only a growth hint; the string still grows
/// past it.
#[inline]
fn json_capacity_hint(data_len: usize) -> usize {
    data_len.saturating_mul(15).clamp(4096, JSON_PREALLOC_CAP)
}

/// Fails with `SizeLimit` once the JSON built so far is longer than `max` bytes.
pub(crate) fn check_json_len(out: &str, max: usize) -> Result<()> {
    if out.len() > max {
        cold_path();
        return Err(ModernXlsxError::SizeLimit(format!(
            "workbook JSON reached {} bytes, above the {max}-byte limit (ReadOptions.limits.maxJsonLength / ZipSecurityLimits.max_json_length)",
            out.len()
        )));
    }
    Ok(())
}

/// Room kept past the JSON limit once the JSON buffer nears it (8 MiB), so the
/// item that takes the JSON over the limit is most often written without
/// another reallocation.
const JSON_LIMIT_SLACK: usize = 8 * 1024 * 1024;

/// Keeps `out`'s capacity from doubling past the target, `max` plus
/// [`JSON_LIMIT_SLACK`]: when less than the slack is free and the next
/// doubling would pass the target, reserves exactly up to the target
/// instead. A buffer with the slack or more free is left alone, so a read
/// whose JSON fits the room it has does not reallocate. Run after each check
/// of the JSON length, it holds the buffer of a read that fails at the limit
/// to at most the target, and the old and the new buffer, both live while it
/// grows, to under twice the target (doubling alone could reach three times
/// the limit). An item larger than the room left (at least the slack when
/// the buffer was left alone) still makes the buffer double.
pub(crate) fn clamp_json_growth(out: &mut String, max: usize) {
    let target = max.saturating_add(JSON_LIMIT_SLACK);
    let cap = out.capacity();
    if cap < target && cap.saturating_mul(2) > target && cap - out.len() < JSON_LIMIT_SLACK {
        out.reserve_exact(target - out.len());
    }
}

/// Adds `"key":items` to the JSON object that ends `out`, before its closing
/// `}`, then checks `out` against `max` ([`check_json_len`],
/// [`clamp_json_growth`]); does nothing
/// when `items` is empty.
fn inject_field<T: serde::Serialize>(
    out: &mut String,
    key: &str,
    items: &[T],
    max: usize,
) -> Result<()> {
    if items.is_empty() {
        return Ok(());
    }
    debug_assert_eq!(out.as_bytes().last(), Some(&b'}'));
    out.pop();
    out.push_str(",\"");
    out.push_str(key);
    out.push_str("\":");
    out.push_str(&serde_json::to_string(items)?);
    out.push('}');
    check_json_len(out, max)?;
    clamp_json_growth(out, max);
    Ok(())
}

/// Appends `,"key":value` to `out`, then checks `out` against `max`
/// ([`check_json_len`], [`clamp_json_growth`]), so each workbook-level field is held to the limit
/// as it is written rather than once at the end.
fn push_field<T: serde::Serialize + ?Sized>(
    out: &mut String,
    key: &str,
    value: &T,
    max: usize,
) -> Result<()> {
    out.push_str(",\"");
    out.push_str(key);
    out.push_str("\":");
    out.push_str(&serde_json::to_string(value)?);
    check_json_len(out, max)?;
    clamp_json_growth(out, max);
    Ok(())
}

/// Shared JSON builder used by the `read_xlsx_json*` entry points.
fn build_json_from_context(
    mut ctx: ReaderContext,
    data_len: usize,
    max_json_length: usize,
) -> Result<String> {
    let mut budget = ctx.budget;
    let sheet_comments = resolve_comments(&mut ctx, &mut budget)?;
    ctx.budget = budget;
    let sheet_tables = resolve_tables(&mut ctx)?;
    let sheet_charts = resolve_charts(&mut ctx)?;
    let PivotFamily {
        sheet_tables: sheet_pivots,
        caches: pivot_caches,
        records: pivot_cache_records,
    } = resolve_pivots(&mut ctx)?;
    let sheet_threaded = resolve_threaded_comments(&mut ctx)?;
    let persons = resolve_persons(&mut ctx)?;
    let sheet_slicers = resolve_slicers(&mut ctx)?;
    let slicer_caches = resolve_slicer_caches(&mut ctx)?;
    let sheet_timelines = resolve_timelines(&mut ctx)?;
    let timeline_caches = resolve_timeline_caches(&mut ctx)?;
    // The sheets count after the charts and threaded comments.
    budget = ctx.budget;

    // --- Build JSON output ---
    // Bounded growth hint; the String still grows past it.
    let mut out = String::with_capacity(json_capacity_hint(data_len));

    out.push_str("{\"sheets\":[");

    for (i, (name, path, state)) in ctx.sheet_targets.iter().enumerate() {
        if i > 0 {
            out.push(',');
        }
        out.push_str("{\"name\":\"");
        crate::ooxml::worksheet::json_escape_to(&mut out, name);
        out.push('"');

        // Emit state only for non-visible sheets (skip_serializing_if equivalent).
        if let Some(s) = sheet_state_to_str(*state) {
            out.push_str(",\"state\":\"");
            out.push_str(s);
            out.push('"');
        }

        out.push_str(",\"worksheet\":");

        let ws_data = ctx.entries.get(path).ok_or_else(|| {
            cold_path();
            ModernXlsxError::MissingPart(format!(
                "ZIP entry '{path}' not found for sheet '{name}' — check xl/_rels/workbook.xml.rels targets"
            ))
        })?;
        WorksheetXml::parse_to_json_within(
            ws_data,
            Some(&ctx.sst),
            &sheet_comments[i],
            &sheet_tables[i],
            &SheetContext {
                date_system: ctx.workbook_xml.date_system,
                rels: ctx.sheet_rels[i].as_ref(),
            },
            &mut out,
            &mut budget,
        )?;

        // Inject the sheet's charts, pivot tables, threaded comments, slicers
        // and timelines into the worksheet JSON, before its closing '}'.
        inject_field(&mut out, "charts", &sheet_charts[i], max_json_length)?;
        inject_field(&mut out, "pivotTables", &sheet_pivots[i], max_json_length)?;
        inject_field(
            &mut out,
            "threadedComments",
            &sheet_threaded[i],
            max_json_length,
        )?;
        inject_field(&mut out, "slicers", &sheet_slicers[i], max_json_length)?;
        inject_field(&mut out, "timelines", &sheet_timelines[i], max_json_length)?;

        out.push('}');
        check_json_len(&out, max_json_length)?;
    }

    out.push(']');

    // Drain entries AFTER sheets are parsed to avoid cloning preserved data.
    let preserved_entries = collect_preserved(&mut ctx);

    // dateSystem
    push_field(
        &mut out,
        "dateSystem",
        &ctx.workbook_xml.date_system,
        max_json_length,
    )?;

    // styles
    push_field(&mut out, "styles", &ctx.styles, max_json_length)?;

    // definedNames
    if !ctx.workbook_xml.defined_names.is_empty() {
        push_field(
            &mut out,
            "definedNames",
            &ctx.workbook_xml.defined_names,
            max_json_length,
        )?;
    }

    // sharedStrings
    push_field(&mut out, "sharedStrings", &ctx.sst, max_json_length)?;

    // docProperties
    if let Some(dp) = &ctx.doc_properties {
        push_field(&mut out, "docProperties", dp, max_json_length)?;
    }

    // themeColors
    if let Some(tc) = &ctx.theme_colors {
        push_field(&mut out, "themeColors", tc, max_json_length)?;
    }

    // calcChain
    if !ctx.calc_chain.is_empty() {
        push_field(&mut out, "calcChain", &ctx.calc_chain, max_json_length)?;
    }

    // workbookViews
    if !ctx.workbook_xml.workbook_views.is_empty() {
        push_field(
            &mut out,
            "workbookViews",
            &ctx.workbook_xml.workbook_views,
            max_json_length,
        )?;
    }

    // protection
    if let Some(prot) = &ctx.workbook_xml.protection {
        push_field(&mut out, "protection", prot, max_json_length)?;
    }

    // calcPr
    if let Some(cp) = &ctx.workbook_xml.calc_pr {
        push_field(&mut out, "calcPr", cp, max_json_length)?;
    }

    // fileVersion
    if let Some(fv) = &ctx.workbook_xml.file_version {
        push_field(&mut out, "fileVersion", fv, max_json_length)?;
    }

    // workbookPr
    if let Some(pr) = &ctx.workbook_xml.workbook_pr {
        push_field(&mut out, "workbookPr", pr, max_json_length)?;
    }

    // pivotCaches
    if !pivot_caches.is_empty() {
        push_field(&mut out, "pivotCaches", &pivot_caches, max_json_length)?;
    }

    // pivotCacheRecords
    if !pivot_cache_records.is_empty() {
        push_field(
            &mut out,
            "pivotCacheRecords",
            &pivot_cache_records,
            max_json_length,
        )?;
    }

    // persons
    if !persons.is_empty() {
        push_field(&mut out, "persons", &persons, max_json_length)?;
    }

    // slicerCaches
    if !slicer_caches.is_empty() {
        push_field(&mut out, "slicerCaches", &slicer_caches, max_json_length)?;
    }

    // timelineCaches
    if !timeline_caches.is_empty() {
        push_field(
            &mut out,
            "timelineCaches",
            &timeline_caches,
            max_json_length,
        )?;
    }

    // externalReferences
    let external_references = external_reference_paths(&mut ctx);
    if !external_references.is_empty() {
        push_field(
            &mut out,
            "externalReferences",
            &external_references,
            max_json_length,
        )?;
    }

    // preservedEntries
    if !preserved_entries.is_empty() {
        // The entries print as JSON number arrays, several bytes of JSON per
        // byte of data, so a lower bound on the final length is checked before
        // they are serialized: `,"preservedEntries":{` and `}` plus the
        // workbook's closing `}` (23 bytes), and per entry its quoted key,
        // colon and brackets (key + 5) and at least one digit per byte with a
        // comma between bytes. Being a lower bound, it never rejects a JSON
        // that fits.
        let floor = out.len()
            + 23
            + preserved_entries
                .iter()
                .map(|(k, v)| k.len() + 5 + (2 * v.len()).saturating_sub(1))
                .sum::<usize>();
        if floor > max_json_length {
            cold_path();
            return Err(ModernXlsxError::SizeLimit(format!(
                "workbook JSON would reach at least {floor} bytes, above the {max_json_length}-byte limit (ReadOptions.limits.maxJsonLength / ZipSecurityLimits.max_json_length)"
            )));
        }
        push_field(
            &mut out,
            "preservedEntries",
            &preserved_entries,
            max_json_length,
        )?;
    }

    // preservedContentTypes
    let preserved_content_types =
        preserved_content_types(ctx.content_types.as_ref(), &preserved_entries);
    if !preserved_content_types.is_empty() {
        push_field(
            &mut out,
            "preservedContentTypes",
            &preserved_content_types,
            max_json_length,
        )?;
    }

    out.push('}');
    check_json_len(&out, max_json_length)?;

    Ok(out)
}

/// Parse worksheet XML for each (name, path, state) triple.
///
/// Sequential by default. With the `parallel` feature, parsing runs on
/// rayon's thread-pool via `par_iter()`; `entries` is a `&HashMap`, which is
/// `Sync`, so it is shared across threads. Shared string indices are
/// resolved inline during parsing via `sst`, text and rich runs by the same
/// index. `t="d"` cells become serials in `date_system`, and hyperlink `r:id`s
/// resolve through the sheet's own `.rels` in `sheet_rels` (parallel to
/// `sheet_targets`).
///
/// The text copied into cells and hyperlinks, and the worksheet namespace
/// declarations re-declared on each kept `<ext>`, is counted in `budget`, one
/// sheet after another, after the defined name, shared string, style, chart
/// and threaded comment records it already holds. With `parallel`, each
/// sheet counts from what `budget` held before the sheets, and only the text
/// each sheet copied is added to it once all are parsed: each sheet's list
/// elements are held to the limit beside the records counted before the
/// sheets, not together with the other sheets'.
fn parse_sheets(
    entries: &HashMap<String, Vec<u8>>,
    sheet_targets: &[(String, String, SheetState)],
    sheet_rels: &[Option<Relationships>],
    sst: &SharedStringTable,
    date_system: DateSystem,
    budget: &mut ExpansionBudget,
) -> Result<Vec<SheetData>> {
    let parse_one = |((name, path, state), rels): (
        &(String, String, SheetState),
        &Option<Relationships>,
    ),
                     budget: &mut ExpansionBudget|
     -> Result<SheetData> {
        let sheet_data = entries.get(path).ok_or_else(|| {
            cold_path();
            ModernXlsxError::MissingPart(format!(
                "ZIP entry '{path}' not found for sheet '{name}' — check xl/_rels/workbook.xml.rels targets"
            ))
        })?;
        let sheet_ctx = SheetContext {
            date_system,
            rels: rels.as_ref(),
        };
        let worksheet =
            WorksheetXml::parse_with_sst_within(sheet_data, Some(sst), &sheet_ctx, budget)?;
        Ok(SheetData {
            name: name.clone(),
            state: sheet_state_to_str(*state).map(str::to_owned),
            worksheet,
        })
    };

    // One line per read: a line per sheet would print a file-chosen name for
    // every `<sheet>`, however many there are.
    debug!("parsing {} worksheets", sheet_targets.len());

    #[cfg(feature = "parallel")]
    {
        let start = *budget;
        let parsed = sheet_targets
            .par_iter()
            .zip(sheet_rels.par_iter())
            .map(|target| {
                let mut own = start;
                parse_one(target, &mut own).map(|sheet| (sheet, own.spent - start.spent))
            })
            .collect::<Result<Vec<_>>>()?;
        parsed
            .into_iter()
            .map(|(sheet, spent)| budget.spend(spent).map(|()| sheet))
            .collect()
    }
    #[cfg(not(feature = "parallel"))]
    sheet_targets
        .iter()
        .zip(sheet_rels)
        .map(|target| parse_one(target, budget))
        .collect()
}

#[cfg(test)]
mod tests;

#[cfg(test)]
mod json_growth_tests {
    use super::{JSON_LIMIT_SLACK, clamp_json_growth};

    const MIB: usize = 1024 * 1024;

    /// A buffer past half the target (limit plus slack) is left alone while
    /// it has at least the slack free, and reserved up to the target once
    /// less than the slack is free.
    #[test]
    fn clamp_reserves_only_when_less_than_the_slack_is_free() {
        let max = 16 * MIB;
        let target = max + JSON_LIMIT_SLACK;

        let mut roomy = String::with_capacity(target / 2 + MIB);
        roomy.push_str("[]");
        let cap = roomy.capacity();
        assert!(cap - roomy.len() >= JSON_LIMIT_SLACK);
        clamp_json_growth(&mut roomy, max);
        assert_eq!(
            roomy.capacity(),
            cap,
            "free room above the slack: no reserve"
        );

        let mut nearly_full = String::with_capacity(target / 2 + MIB);
        let cap = nearly_full.capacity();
        nearly_full.extend(std::iter::repeat_n('x', cap - JSON_LIMIT_SLACK + 1));
        assert_eq!(nearly_full.capacity(), cap);
        clamp_json_growth(&mut nearly_full, max);
        assert!(
            nearly_full.capacity() >= target,
            "less than the slack free: reserved up to the target"
        );
        assert!(nearly_full.capacity() < 2 * cap, "reserved, not doubled");
    }
}
