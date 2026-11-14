//! Reader log hygiene: a package full of relationships whose target part is missing, or whose calcChain names a sheetId no sheet has, yields one summary line per kind of problem, not one line per relationship or entry, and no log line carries a raw control character or a raw cell reference from the file.
//!
//! The `log` facade takes one global logger per process, so this file is its
//! own test binary: the capturing logger here never sees another file's tests.
//! The tests in this binary share it, so each one holds [`TEST_LOCK`] while it
//! clears, reads and inspects the captured records.

use std::fmt::Write as _;
use std::sync::{Mutex, Once, PoisonError};

use log::{Level, LevelFilter, Log, Metadata, Record};
use modern_xlsx_core::reader::{read_xlsx, read_xlsx_json};
use modern_xlsx_core::writer::write_xlsx;
use modern_xlsx_core::zip::{ZipEntry, ZipSecurityLimits, read_zip_entries, write_zip};
use modern_xlsx_core::{SheetData, WorkbookData};

/// Every record the reader logs, at every level, as (level, formatted message).
struct CaptureLogger {
    records: Mutex<Vec<(Level, String)>>,
}

impl Log for CaptureLogger {
    fn enabled(&self, _metadata: &Metadata<'_>) -> bool {
        true
    }

    fn log(&self, record: &Record<'_>) {
        self.records
            .lock()
            .unwrap_or_else(PoisonError::into_inner)
            .push((record.level(), record.args().to_string()));
    }

    fn flush(&self) {}
}

static LOGGER: CaptureLogger = CaptureLogger {
    records: Mutex::new(Vec::new()),
};
static INSTALL: Once = Once::new();
/// Serializes the tests of this binary around the shared logger.
static TEST_LOCK: Mutex<()> = Mutex::new(());

/// Runs `read` with the capturing logger at `Trace` and returns what it logged.
fn captured(read: impl FnOnce()) -> Vec<(Level, String)> {
    INSTALL.call_once(|| {
        log::set_logger(&LOGGER).expect("no other logger in this binary");
        log::set_max_level(LevelFilter::Trace);
    });
    LOGGER
        .records
        .lock()
        .unwrap_or_else(PoisonError::into_inner)
        .clear();
    read();
    std::mem::take(
        &mut *LOGGER
            .records
            .lock()
            .unwrap_or_else(PoisonError::into_inner),
    )
}

/// Records allowed per read: the fixed once-per-read lines plus one summary
/// line per kind. A read of the dangling-table packages emits four (`read_xlsx`) or three
/// (`read_xlsx_json`), and a read of the dropped-calcChain package three or two; before the digest `read_xlsx` emitted 50,003. The bound leaves room
/// for a few more kinds and does not grow with the relationship count.
const MAX_RECORDS: usize = 8;

/// Longest message allowed: the reader caps logged file text at 160 bytes,
/// plus the fixed text of the line (kind, count, quotes and the truncation note).
const MAX_RECORD_LEN: usize = 160 + 120;

/// A one-sheet workbook whose sheet `.rels` holds `count` table relationships,
/// each pointing at a part that is not in the package. Every target carries an
/// encoded newline and ESC byte, which a log line must never print raw.
fn dangling_tables_xlsx(count: usize) -> Vec<u8> {
    let workbook = WorkbookData {
        sheets: vec![SheetData {
            name: "S".into(),
            ..Default::default()
        }],
        ..Default::default()
    };
    let plain = write_xlsx(&workbook).expect("write base workbook");
    let mut parts: Vec<ZipEntry> = read_zip_entries(&plain, &ZipSecurityLimits::default())
        .expect("read base workbook")
        .into_iter()
        .map(|(name, data)| ZipEntry { name, data })
        .collect();

    let mut rels = String::from(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">"#,
    );
    for i in 1..=count {
        write!(
            rels,
            r#"<Relationship Id="rId{i}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/evil&#10;&#27;{i}.xml"/>"#
        )
        .expect("write to String");
    }
    rels.push_str("</Relationships>");

    let rels_path = "xl/worksheets/_rels/sheet1.xml.rels";
    parts.retain(|p| p.name != rels_path);
    parts.push(ZipEntry {
        name: rels_path.into(),
        data: rels.into_bytes(),
    });
    write_zip(&parts).expect("write package")
}

/// Fails unless every record is short and free of control characters.
fn assert_clean(records: &[(Level, String)]) {
    for (level, message) in records {
        assert!(
            message.len() <= MAX_RECORD_LEN,
            "{level} record of {} bytes: {message:?}",
            message.len()
        );
        assert!(
            !message.chars().any(char::is_control),
            "{level} record carries a control character: {message:?}"
        );
    }
}

#[test]
fn many_dangling_relationships_log_one_summary_line() {
    let _guard = TEST_LOCK.lock().unwrap_or_else(PoisonError::into_inner);
    let xlsx = dangling_tables_xlsx(50_000);

    let struct_records = captured(|| {
        read_xlsx(&xlsx).expect("struct read");
    });
    let json_records = captured(|| {
        read_xlsx_json(&xlsx).expect("JSON read");
    });

    for (path, records) in [
        ("read_xlsx", &struct_records),
        ("read_xlsx_json", &json_records),
    ] {
        assert!(
            records.len() <= MAX_RECORDS,
            "{path} emitted {} log records (bound {MAX_RECORDS}); first: {:?}",
            records.len(),
            records.first()
        );
        assert_clean(records);
        let summaries: Vec<&String> = records
            .iter()
            .filter(|(level, message)| {
                *level == Level::Warn && message.starts_with("table part not found")
            })
            .map(|(_, message)| message)
            .collect();
        assert_eq!(summaries.len(), 1, "{path}: {records:?}");
        assert!(
            summaries[0].contains(": 50000 "),
            "{path} summary must carry the count: {}",
            summaries[0]
        );
    }
}

#[test]
fn one_dangling_relationship_logs_one_escaped_warning() {
    let _guard = TEST_LOCK.lock().unwrap_or_else(PoisonError::into_inner);
    let xlsx = dangling_tables_xlsx(1);

    let records = captured(|| {
        read_xlsx(&xlsx).expect("struct read");
    });

    assert_clean(&records);
    let warnings: Vec<&String> = records
        .iter()
        .filter(|(level, _)| *level == Level::Warn)
        .map(|(_, message)| message)
        .collect();
    assert_eq!(warnings.len(), 1, "{records:?}");
    assert!(
        warnings[0].starts_with("table part not found: 1 "),
        "{}",
        warnings[0]
    );
    assert!(
        warnings[0].contains(r#""xl/tables/evil\n\u{1b}1.xml""#),
        "the warning names the escaped target: {}",
        warnings[0]
    );
}

/// A two-sheet workbook (sheetIds 1 and 2) plus an `xl/calcChain.xml` of
/// `count` entries, all naming sheetId 77, which no sheet has. Each cell ref
/// starts with `QQ`, a column no other once-per-read line can contain.
fn dropped_calc_chain_xlsx(count: usize) -> Vec<u8> {
    let sheet = |name: &str| SheetData {
        name: name.into(),
        ..Default::default()
    };
    let workbook = WorkbookData {
        sheets: vec![sheet("One"), sheet("Two")],
        ..Default::default()
    };
    let plain = write_xlsx(&workbook).expect("write base workbook");
    let mut parts: Vec<ZipEntry> = read_zip_entries(&plain, &ZipSecurityLimits::default())
        .expect("read base workbook")
        .into_iter()
        .map(|(name, data)| ZipEntry { name, data })
        .collect();

    let mut chain = String::from(
        r#"<calcChain xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">"#,
    );
    for n in 1..=count {
        write!(chain, r#"<c r="QQ{n}" i="77"/>"#).expect("write to String");
    }
    chain.push_str("</calcChain>");
    parts.retain(|p| p.name != "xl/calcChain.xml");
    parts.push(ZipEntry {
        name: "xl/calcChain.xml".into(),
        data: chain.into_bytes(),
    });
    write_zip(&parts).expect("write package")
}

#[test]
fn dropped_calc_chain_entries_log_one_count_line() {
    let _guard = TEST_LOCK.lock().unwrap_or_else(PoisonError::into_inner);
    let xlsx = dropped_calc_chain_xlsx(1_000);

    let struct_records = captured(|| {
        let wb = read_xlsx(&xlsx).expect("struct read");
        assert!(wb.calc_chain.is_empty(), "{:?}", wb.calc_chain);
    });
    let json_records = captured(|| {
        read_xlsx_json(&xlsx).expect("JSON read");
    });

    for (path, records) in [
        ("read_xlsx", &struct_records),
        ("read_xlsx_json", &json_records),
    ] {
        assert!(
            records.len() <= MAX_RECORDS,
            "{path} emitted {} log records (bound {MAX_RECORDS}); first: {:?}",
            records.len(),
            records.first()
        );
        assert_clean(records);
        let summaries: Vec<&String> = records
            .iter()
            .filter(|(level, message)| *level == Level::Warn && message.starts_with("calcChain:"))
            .map(|(_, message)| message)
            .collect();
        assert_eq!(summaries.len(), 1, "{path}: {records:?}");
        assert_eq!(
            summaries[0],
            "calcChain: 1000 entries name a sheetId which no sheet has; entries dropped",
            "{path}"
        );
        assert!(
            !records.iter().any(|(_, message)| message.contains("QQ")),
            "{path} logged a raw cell ref: {records:?}"
        );
    }
}
