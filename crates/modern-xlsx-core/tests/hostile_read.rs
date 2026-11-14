//! Hostile reads that copy text stored once into every place that names it:
//! one shared string into every cell, one comment author into every comment,
//! one hyperlink target into every hyperlink, one namespace declaration onto
//! every worksheet extension using its prefix. The ZIP limits bound the parts,
//! not these copies, so a read counts the bytes it copies against
//! `ZipSecurityLimits::max_json_length` and fails with `SizeLimit` once they
//! pass it: the JSON a read returns holds every copy at no less than its byte
//! length, so such a read could not have returned JSON within the limit.
//!
//! The packages are built in memory in the shapes of the hostile files
//! measured before the budget existed: a 100,000-character shared string
//! named by 10,000 and by 30,000 cells, a 100,000-character comment author on
//! 10,000 and on 30,000 comments, and the same with one hyperlink target.
//! Each must fail on both read paths within [`LIMIT`] and with a bounded peak
//! heap; a workbook just under the budget reads unchanged, and a custom lower
//! limit is honoured.
//!
//! The JSON path writes each copy escaped: a backslash or a quote takes 2
//! bytes, a control character 6 (`\u00XX`). So the JSON it writes is also
//! held to the limit after each shared string a `<v>` names and each cell,
//! comment and hyperlink, not only at row and sheet ends: one shared string
//! of 40,000 control characters, or of backslashes and quotes, named by the
//! 16,384 cells of one row, or by the 13,000 `<v>` of one cell; a comment
//! author or a hyperlink target of control characters on thousands of
//! comments or hyperlinks (12,000 comments come near the budget, with their
//! copies held besides the JSON); and a rich entry of one-letter
//! control-character runs must each fail on the JSON path within the same
//! bounds. A workbook whose JSON is exactly the limit reads, and one byte
//! less fails.
//!
//! On the model path a cell with several `<v>` keeps the last one, so its
//! shared string is counted once: a workbook whose model fits reads at the
//! limit.
//!
//! A few bytes of XML can also take a hundred times their size in memory:
//! one cell of millions of children, or a list of millions of tiny elements.
//! A cell reaches the sink one child at a time, so one cell of 25,000,000
//! `<f/>`, 12,500,000 `<v>`, millions of `<is>` or of `<f>x</f>` reads
//! within the bounds; each element a worksheet list keeps is counted at the
//! size of its struct plus its strings against the same limit, so 9,000,000
//! `<cfRule/>` (in one `<conditionalFormatting>` or spread over many) and
//! millions of `<dataValidation/>`, `<selection/>` and `<hyperlink/>`
//! end in `SizeLimit`, while millions of `<col/>` and `<mergeCell/>`,
//! whose records are small, read. A generous real sheet of each list
//! (1,000,000 data validations, conditional-format rules, hyperlinks and
//! merged ranges, 16,384 columns) reads at the default limit. The lists the
//! JSON path writes inside one value are checked after each element too:
//! 15,000,000 page breaks, whose records fit the limit but whose JSON does
//! not, end in `SizeLimit` on the JSON path with the JSON held to the limit,
//! and millions of filter columns, filter values and custom filters in one
//! column read.
//!
//! The records of `xl/styles.xml` are counted the same way, before the
//! sheets and in the same count: millions of empty `<dxf/>` (also spelled
//! `<dxf></dxf>`), `<border/>`, `<fill/>`, `<font/>`, `<xf/>` (in `cellXfs`
//! and in `cellStyleXfs`) and gradient `<stop/>` end in `SizeLimit`, while
//! millions of `<numFmt/>`, `<cellStyle/>` and explicitly-off font tags in
//! one differential format, whose records are small, read (and end in
//! `SizeLimit` at a count past the limit). A generous real styles part
//! (64,000 cell formats, 10,000 fonts, fills, borders and differential
//! formats, 2,000 number formats) reads at the default limit.
//!
//! The entries and rich text runs of `xl/sharedStrings.xml` are counted
//! first, in the same count: millions of empty `<r></r>` (in one `<si>` or
//! one per `<si>`, with or without an empty `<rPr/>`) and of `<si/>` end in
//! `SizeLimit`, while 2,000,000 plain entries, and 200,000 rich entries of
//! five runs with the generous styles part and 500,000 data validations,
//! read. One count holds the shared strings, the styles and the sheets: a
//! styles part or a shared string entry that takes 0.6 of the limit with a
//! sheet that takes 0.6 of it ends in `SizeLimit`, while each alone reads.
//!
//! The workbook's defined names, the series and axes of a chart and the
//! threaded comments and their mentions are counted in the same count:
//! millions of empty `<ser></ser>` (in the first and in the second chart
//! type), `<catAx></catAx>`, `<valAx></valAx>`, `<threadedComment/>`,
//! `<mention/>` in one comment and `<definedName name="a"/>` end in
//! `SizeLimit`, while 200,000 series, 2,000,000 threaded comments and
//! 1,000,000 defined names read.
//!
//! A part whose root declares tens of thousands of namespaces and that
//! keeps tens of thousands of verbatim elements (a `<colors>` of
//! `xl/styles.xml`, a worksheet `<ext>`, a pivot `<cacheSource>`) reads
//! within the flood bounds: each element costs its own length and the
//! prefixes it uses, not the number of declarations. A root that binds one
//! prefix to a 400,000-byte URI, with 50,000 `<colors>` or pivot
//! `<cacheSource>` using it, reads within the same bounds: the last element
//! is kept and declares the URI once, not once per element.
//!
//! The charge is each element's size in memory, which depends on the target:
//! the endings above (`SizeLimit` or a read) are those of a 64-bit target. A
//! wasm32 build stores the same records in from half to about four fifths
//! of the bytes (4-byte pointers and lengths; a `ChartAxis` is 216 bytes on
//! x86_64 and 176 on wasm32, measured 2026-11-13), so it keeps more
//! elements before `SizeLimit`, and some floods that end in `SizeLimit`
//! here read there.
//!
//! Every case runs sequentially inside one `#[test]` under a counting
//! `#[global_allocator]` (`peak_alloc::PeakAlloc`, as in
//! `tests/hostile_inputs.rs`), so a timing or peak measures only its own work.

use std::fmt::Write as _;
use std::time::{Duration, Instant};

use modern_xlsx_core::ModernXlsxError;
use modern_xlsx_core::ooxml::cell::col_to_letters;
use modern_xlsx_core::ooxml::worksheet::{Cell, CellType, Row, WorksheetXml};
use modern_xlsx_core::reader::{
    read_xlsx, read_xlsx_json, read_xlsx_json_with_options, read_xlsx_with_options,
};
use modern_xlsx_core::writer::write_xlsx;
use modern_xlsx_core::zip::{ZipEntry, ZipSecurityLimits, read_zip_entries, write_zip};
use modern_xlsx_core::{SheetData, WorkbookData};
use peak_alloc::PeakAlloc;

#[global_allocator]
static PEAK: PeakAlloc = PeakAlloc;

const MIB: usize = 1 << 20;

/// Time bound of each hostile read. The read copies up to the default
/// `max_json_length` (512 MiB) of text before it stops, which an
/// unoptimized test build does more slowly: the slowest case, r1 on the
/// JSON path, took 8.2 s unoptimized and 0.6 s optimized on the test host
/// (an Intel Core Ultra 9 290HX Plus laptop, Windows 11, measured
/// 2026-11-13), so the unoptimized bound is 40 s, about 5 times that, for
/// slower CI runners, and the optimized bound is 1 s.
const LIMIT: Duration = if cfg!(debug_assertions) {
    Duration::from_secs(40)
} else {
    Duration::from_secs(1)
};

/// Peak heap bound of each hostile model read: the 512 MiB of copies the
/// budget lets through plus the package, below the 956 MiB and 964 MiB that
/// the 10,000-copy shapes reached before the budget (and the 2.9 GB of copies
/// of the 30,000-copy shapes).
const MODEL_PEAK_LIMIT: usize = 640 * MIB;

/// Peak heap bound of each hostile JSON read that holds no more than a
/// small part of its copies besides the JSON: the JSON buffer. The per-item
/// JSON checks (after each shared string a `<v>` names, cell, comment,
/// hyperlink and extension) hold it to the limit (512 MiB), and its last
/// growth is clamped to the limit plus 8 MiB, with the old and the new
/// buffer both live during that growth; the counted copies come on top
/// (a read whose budget refuses it before the JSON grows holds only those).
/// Below the 1.7 GiB and 3.8 GiB that the 10,000-copy shapes reached before
/// the budget.
const JSON_PEAK_LIMIT: usize = 1024 * MIB;

/// Time bound of each flood case. The read parses the whole flood, up to
/// 100 MB of XML, which takes an optimized build about 0.9 s on the test
/// host (measured 2026-11-12; 1.2.1's WASM build took 0.85 s to 1.15 s on
/// the same files), past [`LIMIT`]. An unoptimized build reads a tenth of each flood against a
/// tenth of the limit ([`FLOOD_SCALE`]), the same counts per element.
const FLOOD_LIMIT: Duration = if cfg!(debug_assertions) {
    Duration::from_secs(10)
} else {
    Duration::from_secs(3)
};

/// The divisor of each flood's element count and of its limit: 10 in an
/// unoptimized build, whose XML reader alone took about 19 s for 100 MB on
/// the test host (2026-11-12).
const FLOOD_SCALE: usize = if cfg!(debug_assertions) { 10 } else { 1 };

/// Peak heap bound of a read that keeps list elements until their count
/// passes the limit (512 MiB): a list's last doubling holds its old buffer
/// and one twice the counted length live, three times the limit, 1,536
/// MiB; plus 256 MiB for the part's XML and the rest of the read.
const LIST_PEAK_LIMIT: usize = 1792 * MIB;

/// Time bound of a flood of 180 MB to 585 MB of XML (15,000,000 page breaks
/// of 39 bytes each; 12,400,000 merged ranges; 15,000,000 named cell styles;
/// 22,000,000 explicitly-off font tags), which takes longer to parse than
/// [`FLOOD_LIMIT`] allows. The slowest case, 15,000,000 page breaks on the
/// JSON path, took 3.9 s optimized and 8.0 s to 8.6 s for a tenth of the
/// flood unoptimized on the test host (an Intel Core Ultra 9 290HX Plus
/// laptop, Windows 11, measured 2026-11-13), so the unoptimized bound is
/// 40 s, about 4.6 times that, and the optimized bound is 12 s, about 3
/// times.
const BREAK_FLOOD_LIMIT: Duration = if cfg!(debug_assertions) {
    Duration::from_secs(40)
} else {
    Duration::from_secs(12)
};

/// Peak heap bound of the JSON read of 15,000,000 page breaks: the part's
/// 585 MB of XML and the list of breaks (1,604 MiB at the model read's
/// peak), plus the JSON buffer, held to the limit plus 8 MiB with the old
/// buffer live while it grows. The JSON read peaked at 2,440 MiB on the
/// test host on 2026-11-13; a temporary string of the whole list's JSON
/// (570 MB, 38 bytes each with the separating comma) beside the buffer
/// would pass this bound.
const BREAK_PEAK_LIMIT: usize = 2688 * MIB;

/// Peak heap bound of a hostile JSON read whose copies are also held in the
/// parsed sheet (a comment author copied into each comment): the counted
/// copies, at most the limit (512 MiB), on top of the JSON buffer of
/// [`JSON_PEAK_LIMIT`] (at most the limit plus 8 MiB, and the old buffer,
/// smaller than that, while it grows): 1,552 MiB, plus the package.
const JSON_COPIES_PEAK_LIMIT: usize = 1600 * MIB;

const NS: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const REL_NS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const PKG_REL_NS: &str = "http://schemas.openxmlformats.org/package/2006/relationships";

/// `len` characters from a four-letter alphabet (deflate gets about 4:1).
fn text(len: usize) -> String {
    let mut x = 0x9E37_79B9_7F4A_7C15_u64;
    let mut s = String::with_capacity(len);
    while s.len() < len {
        x ^= x << 13;
        x ^= x >> 7;
        x ^= x << 17;
        for k in 0..32 {
            if s.len() == len {
                break;
            }
            s.push(char::from(b"acgt"[((x >> (2 * k)) & 3) as usize]));
        }
    }
    s
}

/// The parts of a written one-sheet workbook with a shared string in A1, to
/// graft the hostile parts onto.
fn base_parts() -> Vec<ZipEntry> {
    let wb = WorkbookData {
        sheets: vec![SheetData {
            name: "S1".into(),
            worksheet: WorksheetXml {
                rows: vec![Row {
                    index: 1,
                    cells: vec![Cell {
                        reference: "A1".into(),
                        cell_type: CellType::SharedString,
                        value: Some("x".into()),
                        ..Default::default()
                    }],
                    ..Default::default()
                }],
                ..Default::default()
            },
            ..Default::default()
        }],
        ..Default::default()
    };
    let zip = write_xlsx(&wb).expect("base write");
    let mut parts: Vec<ZipEntry> = read_zip_entries(&zip, &ZipSecurityLimits::default())
        .expect("base read")
        .into_iter()
        .map(|(name, data)| ZipEntry { name, data })
        .collect();
    parts.sort_by(|a, b| a.name.cmp(&b.name));
    parts
}

/// `base_parts` with each `(name, data)` set (replaced or added), zipped.
fn package(set: Vec<(&str, String)>) -> Vec<u8> {
    let mut parts = base_parts();
    for (name, data) in set {
        parts.retain(|p| p.name != name);
        parts.push(ZipEntry {
            name: name.into(),
            data: data.into_bytes(),
        });
    }
    write_zip(&parts).expect("package write")
}

/// A sheet of `cells` rows, each with one `t="s"` cell naming entry 0.
fn sst_sheet(cells: usize) -> String {
    let mut ws = format!(r#"<worksheet xmlns="{NS}"><sheetData>"#);
    for i in 1..=cells {
        let _ = write!(ws, r#"<row r="{i}"><c r="A{i}" t="s"><v>0</v></c></row>"#);
    }
    ws.push_str("</sheetData></worksheet>");
    ws
}

/// One shared string of `len` characters named by `cells` cells.
fn sst_package(len: usize, cells: usize) -> Vec<u8> {
    let s = text(len);
    package(vec![
        (
            "xl/sharedStrings.xml",
            format!(r#"<sst xmlns="{NS}" count="1" uniqueCount="1"><si><t>{s}</t></si></sst>"#),
        ),
        ("xl/worksheets/sheet1.xml", sst_sheet(cells)),
    ])
}

/// `n` items drawn from `choices` by a xorshift generator.
fn mix(n: usize, choices: &[&str]) -> String {
    let mut x = 0x9E37_79B9_7F4A_7C15_u64;
    let mut s = String::new();
    for _ in 0..n {
        x ^= x << 13;
        x ^= x >> 7;
        x ^= x << 17;
        s.push_str(choices[(x % choices.len() as u64) as usize]);
    }
    s
}

/// `n` control characters as XML character references (`&#1;` to `&#31;`
/// without tab, line feed and carriage return), each written as 6 bytes of
/// JSON (`\u00XX`).
fn control_refs(n: usize) -> String {
    let refs: Vec<String> = (1_u32..32)
        .filter(|c| ![9, 10, 13].contains(c))
        .map(|c| format!("&#{c};"))
        .collect();
    mix(n, &refs.iter().map(String::as_str).collect::<Vec<_>>())
}

/// One shared string `<si>{si_xml}</si>` named by `cells` cells of one row,
/// each cell holding `values` `<v>` elements.
fn sst_row_package(si_xml: &str, cells: u32, values: usize) -> Vec<u8> {
    let mut ws = format!(r#"<worksheet xmlns="{NS}"><sheetData><row r="1">"#);
    for c in 0..cells {
        let _ = write!(ws, r#"<c r="{}1" t="s">"#, col_to_letters(c));
        for _ in 0..values {
            ws.push_str("<v>0</v>");
        }
        ws.push_str("</c>");
    }
    ws.push_str("</row></sheetData></worksheet>");
    package(vec![
        (
            "xl/sharedStrings.xml",
            format!(r#"<sst xmlns="{NS}" count="1" uniqueCount="1"><si>{si_xml}</si></sst>"#),
        ),
        ("xl/worksheets/sheet1.xml", ws),
    ])
}

/// One shared string `<si><t>{t_xml}</t></si>` named by one cell holding
/// `values` `<v>` elements. A 40,000-character XML comment of random text
/// pads the sheet so its part stays under the 100:1 compression ratio.
fn sst_padded_cell_package(t_xml: &str, values: usize) -> Vec<u8> {
    let mut ws = format!(
        r#"<worksheet xmlns="{NS}"><!--{}--><sheetData><row r="1"><c r="A1" t="s">"#,
        text(40_000)
    );
    for _ in 0..values {
        ws.push_str("<v>0</v>");
    }
    ws.push_str("</c></row></sheetData></worksheet>");
    package(vec![
        (
            "xl/sharedStrings.xml",
            format!(r#"<sst xmlns="{NS}" count="1" uniqueCount="1"><si><t>{t_xml}</t></si></sst>"#),
        ),
        ("xl/worksheets/sheet1.xml", ws),
    ])
}

/// One comment author of `len` characters on `comments` comments.
fn author_package(len: usize, comments: usize) -> Vec<u8> {
    author_xml_package(&text(len), comments)
}

/// One comment author written as `author_xml` on `comments` comments.
fn author_xml_package(author_xml: &str, comments: usize) -> Vec<u8> {
    let mut x = format!(
        r#"<comments xmlns="{NS}"><authors><author>{author_xml}</author></authors><commentList>"#
    );
    for i in 1..=comments {
        let _ = write!(
            x,
            r#"<comment ref="A{i}" authorId="0"><text><t>n</t></text></comment>"#
        );
    }
    x.push_str("</commentList></comments>");
    package(vec![
        ("xl/comments1.xml", x),
        (
            "xl/worksheets/_rels/sheet1.xml.rels",
            format!(
                r#"<Relationships xmlns="{PKG_REL_NS}"><Relationship Id="rId1" Type="{REL_NS}/comments" Target="../comments1.xml"/></Relationships>"#
            ),
        ),
    ])
}

/// One external hyperlink target of `len` characters named by `links`
/// hyperlinks.
fn hyperlink_package(len: usize, links: usize) -> Vec<u8> {
    hyperlink_xml_package(&text(len), links)
}

/// A `<worksheet>` that binds prefix `p` to a namespace of `len` characters
/// and keeps `exts` `<p:ext/>` elements, each of which the read makes
/// self-contained with its own copy of the declaration.
fn ext_package(len: usize, exts: usize) -> Vec<u8> {
    let ws = format!(
        r#"<worksheet xmlns="{NS}" xmlns:p="urn:{}"><sheetData/><extLst>{}</extLst></worksheet>"#,
        text(len),
        "<p:ext/>".repeat(exts)
    );
    package(vec![("xl/worksheets/sheet1.xml", ws)])
}

/// One external hyperlink target `https://example.com/{target_xml}` named by
/// `links` hyperlinks.
fn hyperlink_xml_package(target_xml: &str, links: usize) -> Vec<u8> {
    let mut ws = format!(r#"<worksheet xmlns="{NS}" xmlns:r="{REL_NS}"><sheetData/><hyperlinks>"#);
    for i in 1..=links {
        let _ = write!(ws, r#"<hyperlink ref="A{i}" r:id="rId1"/>"#);
    }
    ws.push_str("</hyperlinks></worksheet>");
    package(vec![
        ("xl/worksheets/sheet1.xml", ws),
        (
            "xl/worksheets/_rels/sheet1.xml.rels",
            format!(
                r#"<Relationships xmlns="{PKG_REL_NS}"><Relationship Id="rId1" Type="{REL_NS}/hyperlink" Target="https://example.com/{target_xml}" TargetMode="External"/></Relationships>"#
            ),
        ),
    ])
}

/// Runs `f` and returns its result, its time and its peak heap growth.
fn measured<T>(f: impl FnOnce() -> T) -> (T, Duration, usize) {
    PEAK.reset_peak_usage();
    let base = PEAK.current_usage();
    let start = Instant::now();
    let out = f();
    let elapsed = start.elapsed();
    (out, elapsed, PEAK.peak_usage().saturating_sub(base))
}

/// Asserts `result` is `SizeLimit` naming `limit`. The model path has only
/// the shared-text budget; on the JSON path the JSON written after a row can
/// pass the limit before the copies counted so far do, and either message is
/// right.
fn assert_shared_text_limit<T>(label: &str, result: Result<T, ModernXlsxError>, limit: usize) {
    match result {
        Ok(_) => panic!("{label}: expected SIZE_LIMIT, got Ok"),
        Err(e) => {
            assert_eq!(e.code(), "SIZE_LIMIT", "{label}: {e}");
            let msg = e.to_string();
            assert!(
                msg.contains("shared text")
                    || (label.ends_with("(json)") && msg.contains("workbook JSON reached")),
                "{label}: {msg}"
            );
            assert!(
                msg.contains(&format!("{limit}-byte limit")),
                "{label}: {msg}"
            );
            assert!(
                msg.contains(
                    "ReadOptions.limits.maxJsonLength / ZipSecurityLimits.max_json_length"
                ),
                "{label}: {msg}"
            );
        }
    }
}

/// Both read paths of `zip` fail with `SizeLimit` within [`LIMIT`] and
/// [`MODEL_PEAK_LIMIT`] or [`JSON_PEAK_LIMIT`].
fn hostile_case(label: &str, zip: &[u8]) {
    let limit = ZipSecurityLimits::default().max_json_length;
    for (path, read, peak_limit) in [
        (
            "model",
            &(|z: &[u8]| read_xlsx(z).map(drop)) as &dyn Fn(&[u8]) -> _,
            MODEL_PEAK_LIMIT,
        ),
        (
            "json",
            &|z: &[u8]| read_xlsx_json(z).map(drop),
            JSON_PEAK_LIMIT,
        ),
    ] {
        let (result, elapsed, peak) = measured(|| read(zip));
        println!(
            "CASE {label} ({path}): zip {} B; {elapsed:?}, peak {} MiB",
            zip.len(),
            peak / MIB
        );
        assert_shared_text_limit(&format!("{label} ({path})"), result, limit);
        assert!(elapsed < LIMIT, "{label} ({path}): took {elapsed:?}");
        assert!(peak < peak_limit, "{label} ({path}): peak {peak} bytes");
    }
}

/// The JSON read path of `zip` fails with `SizeLimit` within [`LIMIT`] and
/// [`JSON_PEAK_LIMIT`]: for the shapes whose copies fit the budget but whose
/// escaped JSON does not (the model path reads them, or, for a rich entry,
/// is bounded by the budget alone).
fn json_case(label: &str, zip: &[u8]) {
    json_case_within(label, zip, JSON_PEAK_LIMIT);
}

/// [`json_case`] under the peak bound `peak_limit`.
fn json_case_within(label: &str, zip: &[u8], peak_limit: usize) {
    let limit = ZipSecurityLimits::default().max_json_length;
    let (result, elapsed, peak) = measured(|| read_xlsx_json(zip).map(drop));
    println!(
        "CASE {label} (json): zip {} B; {elapsed:?}, peak {} MiB",
        zip.len(),
        peak / MIB
    );
    assert_shared_text_limit(&format!("{label} (json)"), result, limit);
    assert!(elapsed < LIMIT, "{label} (json): took {elapsed:?}");
    assert!(peak < peak_limit, "{label} (json): peak {peak} bytes");
}

/// 64-bit FNV-1a over `bytes`, to pin a JSON output without storing it.
fn fnv1a(bytes: &[u8]) -> u64 {
    bytes.iter().fold(0xCBF2_9CE4_8422_2325, |h, &b| {
        (h ^ u64::from(b)).wrapping_mul(0x0100_0000_01B3)
    })
}

/// One string of 1,000 bytes named by 100,000 cells (100 MB of copies)
/// reads on both paths, and the JSON is the one the read returned before the
/// budget (length and FNV-1a hash taken on the 1.3.0 development tree of
/// 2026-11-12, UTC).
fn under_budget_case() {
    let zip = sst_package(1_000, 100_000);
    let json = read_xlsx_json(&zip).expect("under the budget, the JSON read succeeds");
    assert_eq!(
        (json.len(), fnv1a(json.as_bytes())),
        (UNDER_BUDGET_JSON_LEN, UNDER_BUDGET_JSON_FNV),
        "the JSON differs from the base read"
    );
    let wb = read_xlsx(&zip).expect("under the budget, the model read succeeds");
    let rows = &wb.sheets[0].worksheet.rows;
    assert_eq!(rows.len(), 100_000);
    let s = text(1_000);
    assert!(
        rows.iter()
            .all(|r| r.cells[0].value.as_deref() == Some(s.as_str()))
    );
}

const UNDER_BUDGET_JSON_LEN: usize = 108_579_574;
const UNDER_BUDGET_JSON_FNV: u64 = 0x391C_78A3_27D3_B7AF;

/// A limit below the default is the budget: 100 MB of copies against 50 MB.
fn custom_limit_case() {
    let zip = sst_package(1_000, 100_000);
    let low = ZipSecurityLimits {
        max_json_length: 50_000_000,
        ..ZipSecurityLimits::default()
    };
    let (result, elapsed, _) = measured(|| read_xlsx_json_with_options(&zip, &low));
    assert_shared_text_limit("custom limit (json)", result, 50_000_000);
    assert!(elapsed < LIMIT, "custom limit (json): took {elapsed:?}");
    let (result, elapsed, _) = measured(|| read_xlsx_with_options(&zip, &low));
    assert_shared_text_limit("custom limit (model)", result, 50_000_000);
    assert!(elapsed < LIMIT, "custom limit (model): took {elapsed:?}");
}

/// `limits` with `max_json_length` set to `len`.
fn json_limit(len: usize) -> ZipSecurityLimits {
    ZipSecurityLimits {
        max_json_length: len,
        ..ZipSecurityLimits::default()
    }
}

/// A workbook whose JSON is exactly the limit reads to the same JSON, and one
/// byte less fails, for escaped copies in cells, rich runs, comments and
/// hyperlinks (each written under the checks that follow every copy).
fn exact_limit_case() {
    for (label, zip) in [
        (
            "cells",
            sst_row_package(&format!("<t>{}</t>", control_refs(100)), 50, 1),
        ),
        (
            "rich runs",
            sst_row_package("<r><t>&#1;\\</t></r><r><t>&quot;</t></r>", 50, 1),
        ),
        ("comments", author_xml_package(&control_refs(100), 20)),
        ("hyperlinks", hyperlink_xml_package(&control_refs(100), 20)),
    ] {
        let json = read_xlsx_json(&zip).expect(label);
        let at = read_xlsx_json_with_options(&zip, &json_limit(json.len()))
            .unwrap_or_else(|e| panic!("{label}: at the limit: {e}"));
        assert_eq!(at, json, "{label}");
        let err = read_xlsx_json_with_options(&zip, &json_limit(json.len() - 1))
            .expect_err("one byte under the JSON length must fail");
        assert_eq!(err.code(), "SIZE_LIMIT", "{label}: {err}");
    }
}

/// A cell with three `<v>` keeps the last one on the model path, so its
/// shared string is counted once: 600 cells naming a 1,000-byte string read
/// at a 600,000-byte limit and fail one byte under it.
fn several_values_case() {
    let s = "a".repeat(1_000);
    let zip = sst_row_package(&format!("<t>{s}</t>"), 600, 3);
    let wb = read_xlsx_with_options(&zip, &json_limit(600_000))
        .expect("each cell keeps one copy, so 600,000 bytes fit");
    let cells = &wb.sheets[0].worksheet.rows[0].cells;
    assert_eq!(cells.len(), 600);
    assert!(cells.iter().all(|c| c.value.as_deref() == Some(s.as_str())));
    let err = read_xlsx_with_options(&zip, &json_limit(599_999)).unwrap_err();
    assert_eq!(err.code(), "SIZE_LIMIT", "{err}");
    assert!(err.to_string().contains("shared text"), "{err}");
}

/// A one-sheet package whose `<worksheet>` holds `body` after a comment of
/// random text a 25th of its size, which keeps the part under the 100:1
/// compression ratio.
fn flood_package(body: &str) -> Vec<u8> {
    let ws = format!(
        r#"<worksheet xmlns="{NS}"><!--{}-->{body}</worksheet>"#,
        text(body.len() / 25 + 1_000)
    );
    package(vec![("xl/worksheets/sheet1.xml", ws)])
}

/// One cell A1 holding `children`.
fn one_cell(children: &str) -> String {
    format!(r#"<sheetData><row r="1"><c r="A1">{children}</c></row></sheetData>"#)
}

/// `n` elements, each written by `f` from its 1-based number.
fn each(n: usize, f: impl Fn(&mut String, usize)) -> String {
    let mut s = String::new();
    for i in 1..=n {
        f(&mut s, i);
    }
    s
}

/// How a flood case must end on both read paths.
#[derive(Clone, Copy, Debug)]
enum Flood {
    /// The read succeeds.
    Reads,
    /// The read fails with `SizeLimit`: the stored list elements pass the
    /// limit.
    ListLimit,
    /// The model path reads; the JSON path fails with `SizeLimit` once the
    /// JSON it writes passes the limit.
    JsonLimit,
}

/// The limit of a flood read: the default divided by [`FLOOD_SCALE`].
fn flood_limits() -> ZipSecurityLimits {
    json_limit(ZipSecurityLimits::default().max_json_length / FLOOD_SCALE)
}

/// `n` divided by [`FLOOD_SCALE`].
const fn scaled(n: usize) -> usize {
    n / FLOOD_SCALE
}

/// Both read paths of `zip` end as `expect` says within [`FLOOD_LIMIT`]
/// and `peaks` (model, JSON), under [`flood_limits`].
fn flood_case(label: &str, zip: &[u8], expect: Flood, peaks: (usize, usize)) {
    flood_case_within(label, zip, expect, peaks, FLOOD_LIMIT);
}

/// [`flood_case`] within the time bound `time_limit`.
fn flood_case_within(
    label: &str,
    zip: &[u8],
    expect: Flood,
    peaks: (usize, usize),
    time_limit: Duration,
) {
    flood_case_counted(label, zip, expect, peaks, time_limit, FLOOD_SCALE);
}

/// [`flood_case_within`] for a flood whose element count, as `label`
/// states it, was divided by `divisor` (1: the count `label` states).
fn flood_case_counted(
    label: &str,
    zip: &[u8],
    expect: Flood,
    peaks: (usize, usize),
    time_limit: Duration,
    divisor: usize,
) {
    let count = if divisor == 1 {
        "full count".to_owned()
    } else {
        format!("count / {divisor}")
    };
    let limits = flood_limits();
    let limit = limits.max_json_length;
    for (path, read, peak_limit) in [
        (
            "model",
            &(|z: &[u8]| read_xlsx_with_options(z, &limits).map(drop)) as &dyn Fn(&[u8]) -> _,
            peaks.0,
        ),
        (
            "json",
            &|z: &[u8]| read_xlsx_json_with_options(z, &limits).map(drop),
            peaks.1,
        ),
    ] {
        let (result, elapsed, peak) = measured(|| read(zip));
        println!(
            "CASE {label} ({count}, {path}): zip {} B; {elapsed:?}, peak {} MiB",
            zip.len(),
            peak / MIB
        );
        match (expect, result) {
            (Flood::Reads, Ok(())) => {}
            (Flood::JsonLimit, Ok(())) if path == "model" => {}
            (Flood::JsonLimit, Err(e)) if path == "json" => {
                assert_eq!(e.code(), "SIZE_LIMIT", "{label} ({path}): {e}");
                let msg = e.to_string();
                assert!(
                    msg.contains("workbook JSON reached")
                        && msg.contains(&format!("{limit}-byte limit")),
                    "{label} ({path}): {msg}"
                );
            }
            (Flood::ListLimit, Err(e)) => {
                assert_eq!(e.code(), "SIZE_LIMIT", "{label} ({path}): {e}");
                let msg = e.to_string();
                assert!(
                    msg.contains("lists take more memory"),
                    "{label} ({path}): {msg}"
                );
                assert!(
                    msg.contains(&format!("{limit}-byte limit")),
                    "{label} ({path}): {msg}"
                );
            }
            (expect, result) => {
                panic!("{label} ({path}): expected {expect:?}, got {result:?}")
            }
        }
        assert!(elapsed < time_limit, "{label} ({path}): took {elapsed:?}");
        assert!(peak < peak_limit, "{label} ({path}): peak {peak} bytes");
    }
}

/// One cell of millions of children reads one child at a time, and lists of
/// millions of tiny elements end in `SizeLimit` (or read, when their records
/// are small), each within the time and peak bounds.
fn element_flood_cases() {
    let cells = (MODEL_PEAK_LIMIT, JSON_PEAK_LIMIT);
    let lists = (LIST_PEAK_LIMIT, LIST_PEAK_LIMIT);
    flood_case(
        "one cell of 25,000,000 <f/>",
        &flood_package(&one_cell(&"<f/>".repeat(scaled(25_000_000)))),
        Flood::Reads,
        cells,
    );
    flood_case(
        "one cell of 12,500,000 <v>1</v>",
        &flood_package(&one_cell(&"<v>1</v>".repeat(scaled(12_500_000)))),
        Flood::Reads,
        cells,
    );
    flood_case(
        "one cell of 8,000,000 <is><t/></is>",
        &flood_package(&one_cell(&"<is><t/></is>".repeat(scaled(8_000_000)))),
        Flood::Reads,
        cells,
    );
    flood_case(
        "one cell of 4,000,000 <is><r><t>a</t></r></is>",
        &flood_package(&one_cell(
            &"<is><r><t>a</t></r></is>".repeat(scaled(4_000_000)),
        )),
        Flood::Reads,
        cells,
    );
    flood_case(
        "one cell of 8,000,000 <f>x</f>",
        &flood_package(&one_cell(&"<f>x</f>".repeat(scaled(8_000_000)))),
        Flood::Reads,
        cells,
    );
    flood_case(
        "9,000,000 <cfRule/> in one <conditionalFormatting>",
        &flood_package(&format!(
            r#"<sheetData/><conditionalFormatting sqref="A1">{}</conditionalFormatting>"#,
            "<cfRule/>".repeat(scaled(9_000_000))
        )),
        Flood::ListLimit,
        lists,
    );
    let thousand_rules = format!(
        r#"<conditionalFormatting sqref="A1">{}</conditionalFormatting>"#,
        "<cfRule/>".repeat(1_000)
    );
    flood_case(
        "9,000,000 <cfRule/> over 9,000 <conditionalFormatting>",
        &flood_package(&format!(
            "<sheetData/>{}",
            thousand_rules.repeat(scaled(9_000))
        )),
        Flood::ListLimit,
        lists,
    );
    flood_case(
        "3,000,000 <conditionalFormatting> of one <cfRule/>",
        &flood_package(&format!(
            "<sheetData/>{}",
            "<conditionalFormatting><cfRule/></conditionalFormatting>".repeat(scaled(3_000_000))
        )),
        Flood::ListLimit,
        lists,
    );
    flood_case(
        "9,000,000 <dataValidation/>",
        &flood_package(&format!(
            "<sheetData/><dataValidations>{}</dataValidations>",
            "<dataValidation/>".repeat(scaled(9_000_000))
        )),
        Flood::ListLimit,
        lists,
    );
    flood_case(
        "9,000,000 <selection/>",
        &flood_package(&format!(
            "<sheetViews><sheetView>{}</sheetView></sheetViews><sheetData/>",
            "<selection/>".repeat(scaled(9_000_000))
        )),
        Flood::ListLimit,
        lists,
    );
    flood_case(
        "5,000,000 <hyperlink ref=A1/>",
        &flood_package(&format!(
            "<sheetData/><hyperlinks>{}</hyperlinks>",
            r#"<hyperlink ref="A1"/>"#.repeat(scaled(5_000_000))
        )),
        Flood::ListLimit,
        lists,
    );
    flood_case(
        "5,000,000 <col min=1 max=1/>",
        &flood_package(&format!(
            "<cols>{}</cols><sheetData/>",
            r#"<col min="1" max="1"/>"#.repeat(scaled(5_000_000))
        )),
        Flood::Reads,
        lists,
    );
    flood_case(
        "5,000,000 <mergeCell ref=A1/>",
        &flood_package(&format!(
            "<sheetData/><mergeCells>{}</mergeCells>",
            r#"<mergeCell ref="A1"/>"#.repeat(scaled(5_000_000))
        )),
        Flood::Reads,
        lists,
    );
}

/// The lists the JSON path writes inside a nested value (the page breaks,
/// the auto filter's columns, and the values and custom filters of one
/// column) are checked after each element: 15,000,000 page breaks, whose
/// records fit the limit but whose JSON (38 bytes each with the separating
/// comma) does not, end in
/// `SizeLimit` on the JSON path within [`BREAK_FLOOD_LIMIT`] and
/// [`BREAK_PEAK_LIMIT`], and millions of filter columns, filter values and
/// custom filters in one column read within the flood bounds.
fn sink_list_cases() {
    let lists = (LIST_PEAK_LIMIT, LIST_PEAK_LIMIT);
    let brk = r#"<brk id="1048575" max="16383" man="1"/>"#;
    let breaks = brk.repeat(scaled(15_000_000));
    for list in ["rowBreaks", "colBreaks"] {
        flood_case_within(
            &format!("15,000,000 <brk/> in <{list}>"),
            &flood_package(&format!("<sheetData/><{list}>{breaks}</{list}>")),
            Flood::JsonLimit,
            (LIST_PEAK_LIMIT, BREAK_PEAK_LIMIT),
            BREAK_FLOOD_LIMIT,
        );
    }
    drop(breaks);
    flood_case(
        "3,000,000 <filterColumn></filterColumn>",
        &flood_package(&format!(
            r#"<sheetData/><autoFilter ref="A1">{}</autoFilter>"#,
            "<filterColumn></filterColumn>".repeat(scaled(3_000_000))
        )),
        Flood::Reads,
        lists,
    );
    flood_case(
        "9,000,000 <filter val=a/> in one column",
        &flood_package(&format!(
            r#"<sheetData/><autoFilter ref="A1"><filterColumn><filters>{}</filters></filterColumn></autoFilter>"#,
            r#"<filter val="a"/>"#.repeat(scaled(9_000_000))
        )),
        Flood::Reads,
        lists,
    );
    flood_case(
        "6,000,000 <customFilter val=a/> in one column",
        &flood_package(&format!(
            r#"<sheetData/><autoFilter ref="A1"><filterColumn><customFilters>{}</customFilters></filterColumn></autoFilter>"#,
            r#"<customFilter val="a"/>"#.repeat(scaled(6_000_000))
        )),
        Flood::Reads,
        lists,
    );
}

/// A package whose `xl/styles.xml` holds `inner` after a comment of random
/// text a 25th of its size, which keeps the part under the 100:1
/// compression ratio.
fn styles_package(inner: &str) -> Vec<u8> {
    let styles = format!(
        r#"<styleSheet xmlns="{NS}"><!--{}-->{inner}</styleSheet>"#,
        text(inner.len() / 25 + 1_000)
    );
    package(vec![("xl/styles.xml", styles)])
}

/// Each record of `xl/styles.xml` is a list element too: millions of
/// records of 5 to 37 bytes of XML each (empty records in the self-closing
/// and the start/end spelling, an explicitly-off font tag, a number format
/// with its attributes) end in `SizeLimit`, or read when their records are
/// small, within the time and peak bounds. The small records end in
/// `SizeLimit` too at a count past the limit: 15,000,000 named cell styles
/// (40 bytes each on a 64-bit target) and 22,000,000 explicitly-off font
/// tags (25 bytes each), within [`BREAK_FLOOD_LIMIT`].
fn style_flood_cases() {
    let lists = (LIST_PEAK_LIMIT, LIST_PEAK_LIMIT);
    let cases: [(&str, &str, &str, usize, Flood); 13] = [
        ("<dxfs>", "<dxf/>", "</dxfs>", 3_300_000, Flood::ListLimit),
        (
            "<dxfs>",
            "<dxf></dxf>",
            "</dxfs>",
            3_300_000,
            Flood::ListLimit,
        ),
        (
            "<borders>",
            "<border/>",
            "</borders>",
            6_000_000,
            Flood::ListLimit,
        ),
        (
            "<fills>",
            "<fill/>",
            "</fills>",
            9_000_000,
            Flood::ListLimit,
        ),
        (
            "<fonts>",
            "<font/>",
            "</fonts>",
            5_000_000,
            Flood::ListLimit,
        ),
        (
            "<cellXfs>",
            "<xf/>",
            "</cellXfs>",
            9_000_000,
            Flood::ListLimit,
        ),
        (
            "<cellStyleXfs>",
            "<xf/>",
            "</cellStyleXfs>",
            9_000_000,
            Flood::ListLimit,
        ),
        (
            "<fills><fill><gradientFill>",
            "<stop/>",
            "</gradientFill></fill></fills>",
            9_000_000,
            Flood::ListLimit,
        ),
        (
            "<dxfs><dxf><fill><gradientFill>",
            "<stop/>",
            "</gradientFill></fill></dxf></dxfs>",
            9_000_000,
            Flood::ListLimit,
        ),
        (
            "<numFmts>",
            "<numFmt/>",
            "</numFmts>",
            9_000_000,
            Flood::Reads,
        ),
        (
            "<cellStyles>",
            "<cellStyle/>",
            "</cellStyles>",
            9_000_000,
            Flood::Reads,
        ),
        (
            "<dxfs><dxf><font>",
            r#"<b val="0"/>"#,
            "</font></dxf></dxfs>",
            9_000_000,
            Flood::Reads,
        ),
        (
            "<numFmts>",
            r#"<numFmt numFmtId="1" formatCode="0"/>"#,
            "</numFmts>",
            3_000_000,
            Flood::Reads,
        ),
    ];
    for (open, record, close, n, expect) in cases {
        flood_case(
            &format!("{n} {record} in {open}"),
            &styles_package(&format!("{open}{}{close}", record.repeat(scaled(n)))),
            expect,
            lists,
        );
    }
    let small: [(&str, &str, &str, usize); 2] = [
        ("<cellStyles>", "<cellStyle/>", "</cellStyles>", 15_000_000),
        (
            "<dxfs><dxf><font>",
            r#"<b val="0"/>"#,
            "</font></dxf></dxfs>",
            22_000_000,
        ),
    ];
    for (open, record, close, n) in small {
        flood_case_within(
            &format!("{n} {record} in {open}"),
            &styles_package(&format!("{open}{}{close}", record.repeat(scaled(n)))),
            Flood::ListLimit,
            lists,
            BREAK_FLOOD_LIMIT,
        );
    }
}

/// `<{root} xmlns=main{decls}>` holding `inner` after a comment of random
/// text a 25th of its size, which keeps the part under the 100:1
/// compression ratio.
fn padded_part(root: &str, decls: &str, inner: &str) -> String {
    format!(
        r#"<{root} xmlns="{NS}"{decls}><!--{}-->{inner}</{root}>"#,
        text((decls.len() + inner.len()) / 25 + 1_000)
    )
}

/// `n` namespace declarations `xmlns:p1="urn:ns:1"` to `xmlns:pn`.
fn root_decls(n: usize) -> String {
    each(n, |s, i| {
        let _ = write!(s, r#" xmlns:p{i}="urn:ns:{i}""#);
    })
}

/// The text of the part `name` of [`base_parts`].
fn base_part(name: &str) -> String {
    let part = base_parts()
        .into_iter()
        .find(|p| p.name == name)
        .unwrap_or_else(|| panic!("{name} in the base parts"));
    String::from_utf8(part.data).expect("UTF-8 part")
}

/// A package whose workbook names one pivot cache definition, `definition`.
fn pivot_cache_package(definition: String) -> Vec<u8> {
    let rels = base_part("xl/_rels/workbook.xml.rels").replace(
        "</Relationships>",
        &format!(
            r#"<Relationship Id="rIdPC1" Type="{REL_NS}/pivotCacheDefinition" Target="pivotCache/pivotCacheDefinition1.xml"/></Relationships>"#
        ),
    );
    package(vec![
        ("xl/_rels/workbook.xml.rels", rels),
        ("xl/pivotCache/pivotCacheDefinition1.xml", definition),
    ])
}

/// A part's root that declares many namespaces costs each verbatim element
/// of the part (a depth-1 `<colors>` of `xl/styles.xml`, a kept worksheet
/// `<ext>`, a pivot `<cacheSource>` kept whole) its own length plus the
/// prefixes it uses, not the number of declarations: each shape reads on
/// both paths within the flood bounds, with the element kept.
fn namespace_cases() {
    let quick = (MODEL_PEAK_LIMIT, JSON_PEAK_LIMIT);
    let limits = flood_limits();
    for (k, m, record) in [
        (50_000, 20_000, "<colors/>"),
        (10_000, 100_000, "<colors><p1:a/></colors>"),
    ] {
        let zip = package(vec![(
            "xl/styles.xml",
            padded_part("styleSheet", &root_decls(k), &record.repeat(m)),
        )]);
        flood_case(
            &format!("{k} root declarations, {m} {record} in xl/styles.xml"),
            &zip,
            Flood::Reads,
            quick,
        );
        let wb = read_xlsx_with_options(&zip, &limits).expect("styles model");
        let colors = wb.styles.colors_xml.as_deref().expect("colors kept");
        assert_eq!(
            colors.contains("xmlns:p1="),
            record.contains("p1:"),
            "{colors}"
        );
    }
    for (k, m, record) in [
        (50_000, 20_000, r#"<ext uri="a"/>"#),
        (10_000, 100_000, r#"<ext uri="a"><p1:a/></ext>"#),
    ] {
        let zip = package(vec![(
            "xl/worksheets/sheet1.xml",
            padded_part(
                "worksheet",
                &root_decls(k),
                &format!("<sheetData/><extLst>{}</extLst>", record.repeat(m)),
            ),
        )]);
        flood_case(
            &format!("{k} root declarations, {m} kept {record} in a worksheet"),
            &zip,
            Flood::Reads,
            quick,
        );
        let wb = read_xlsx_with_options(&zip, &limits).expect("worksheet model");
        let kept = &wb.sheets[0].worksheet.preserved_extensions;
        assert_eq!(kept.len(), m);
        assert_eq!(
            kept[0].contains("xmlns:p1="),
            record.contains("p1:"),
            "{}",
            kept[0]
        );
    }
    for (k, m, record) in [
        (50_000, 20_000, r#"<cacheSource type="external"/>"#),
        (
            10_000,
            100_000,
            r#"<cacheSource type="external"><p1:a/></cacheSource>"#,
        ),
    ] {
        let zip = pivot_cache_package(padded_part(
            "pivotCacheDefinition",
            &format!(r#" xmlns:r="{REL_NS}"{}"#, root_decls(k)),
            &format!(r#"{}<cacheFields count="0"/>"#, record.repeat(m)),
        ));
        flood_case(
            &format!("{k} root declarations, {m} {record} in a pivot cache definition"),
            &zip,
            Flood::Reads,
            quick,
        );
        let wb = read_xlsx_with_options(&zip, &limits).expect("pivot model");
        let raw = wb
            .pivot_caches
            .first()
            .and_then(|c| c.source.raw_xml.as_deref())
            .expect("cacheSource kept");
        assert_eq!(raw.contains("xmlns:p1="), record.contains("p1:"), "{raw}");
    }
    long_uri_cases();
    inherited_declaration_case();
}

/// A part whose root binds `p1` to a 400,000-byte URI and that repeats
/// 50,000 times a verbatim element using it (a depth-1 `<colors>` of
/// `xl/styles.xml`, a pivot `<cacheSource>` kept whole) keeps only the last
/// one, made self-contained once: each repeat costs its own length, not a
/// copy of the URI, and each shape reads on both paths within the flood
/// bounds, with the URI declared on the element kept. The count is not
/// divided by [`FLOOD_SCALE`]: each part is about 3 MB of XML at most
/// (1.7 MB for the styles part, 3.0 MB for the pivot part), and a copy of
/// the URI per element took an unoptimized build over 60 s for 50,000
/// elements on the test host (2026-11-13), but about 6 s for 5,000, within
/// [`FLOOD_LIMIT`].
fn long_uri_cases() {
    let quick = (MODEL_PEAK_LIMIT, JSON_PEAK_LIMIT);
    let limits = flood_limits();
    let uri = text(400_000);
    let decl = format!(r#" xmlns:p1="{uri}""#);
    let m = 50_000;
    let record = "<colors><p1:a/></colors>";
    let zip = package(vec![(
        "xl/styles.xml",
        padded_part("styleSheet", &decl, &record.repeat(m)),
    )]);
    flood_case_counted(
        &format!("p1 of 400,000 bytes, {m} {record} in xl/styles.xml"),
        &zip,
        Flood::Reads,
        quick,
        FLOOD_LIMIT,
        1,
    );
    let wb = read_xlsx_with_options(&zip, &limits).expect("styles model");
    let colors = wb.styles.colors_xml.as_deref().expect("colors kept");
    assert!(colors.contains(&decl), "colors declares p1");
    let record = r#"<cacheSource type="external"><p1:a/></cacheSource>"#;
    let zip = pivot_cache_package(padded_part(
        "pivotCacheDefinition",
        &format!(r#" xmlns:r="{REL_NS}"{decl}"#),
        &format!(r#"{}<cacheFields count="0"/>"#, record.repeat(m)),
    ));
    flood_case_counted(
        &format!("p1 of 400,000 bytes, {m} {record} in a pivot cache definition"),
        &zip,
        Flood::Reads,
        quick,
        FLOOD_LIMIT,
        1,
    );
    let wb = read_xlsx_with_options(&zip, &limits).expect("pivot model");
    let raw = wb
        .pivot_caches
        .first()
        .and_then(|c| c.source.raw_xml.as_deref())
        .expect("cacheSource kept");
    assert!(raw.contains(&decl), "cacheSource declares p1");
}

/// A package whose `xl/sharedStrings.xml` holds `inner`, padded as
/// [`padded_part`] pads.
fn sst_flood_package(inner: &str) -> Vec<u8> {
    package(vec![(
        "xl/sharedStrings.xml",
        padded_part("sst", "", inner),
    )])
}

/// Each shared string entry and rich text run is a list element too
/// (48 bytes, its plain text not counted, and 248 bytes plus its strings on a
/// 64-bit target): millions
/// of empty runs in one entry or one per entry, with or without an empty
/// `<rPr/>`, and millions of empty entries end in `SizeLimit` within the
/// time and peak bounds.
fn sst_flood_cases() {
    let lists = (LIST_PEAK_LIMIT, LIST_PEAK_LIMIT);
    let cases: [(&str, &str, &str, usize); 4] = [
        ("<si>", "<r></r>", "</si>", 10_000_000),
        ("", "<si><r></r></si>", "", 10_000_000),
        ("", "<si/>", "", 15_000_000),
        ("<si>", "<r><rPr/></r>", "</si>", 4_000_000),
    ];
    for (open, record, close, n) in cases {
        flood_case(
            &format!("{n} {record} in xl/sharedStrings.xml {open}"),
            &sst_flood_package(&format!("{open}{}{close}", record.repeat(scaled(n)))),
            Flood::ListLimit,
            lists,
        );
    }
}

/// Real shared string tables read on both paths at the default limit (in
/// an unoptimized build a tenth of each at a tenth of the limit): 2,000,000
/// plain entries, and 200,000 rich entries of five runs each together with
/// the generous real styles part and a sheet of 500,000 data validations.
fn real_sst_cases() {
    let limits = flood_limits();
    let plain = each(scaled(2_000_000), |s, i| {
        let _ = write!(s, "<si><t>Item {i}</t></si>");
    });
    let zip = package(vec![(
        "xl/sharedStrings.xml",
        padded_part("sst", "", &plain),
    )]);
    drop(plain);
    let wb = read_xlsx_with_options(&zip, &limits)
        .unwrap_or_else(|e| panic!("2,000,000 plain entries (model): {e}"));
    let table = wb.shared_strings.as_ref().expect("table kept");
    assert_eq!(table.strings.len(), scaled(2_000_000));
    drop(wb);
    read_xlsx_json_with_options(&zip, &limits)
        .unwrap_or_else(|e| panic!("2,000,000 plain entries (json): {e}"));
    drop(zip);

    let rich = each(scaled(200_000), |s, i| {
        s.push_str("<si>");
        for k in 0..5 {
            let _ = write!(
                s,
                r#"<r><rPr><b/><sz val="11"/><color theme="1"/><rFont val="Calibri"/><family val="2"/><scheme val="minor"/></rPr><t xml:space="preserve">Word {i}.{k} </t></r>"#
            );
        }
        s.push_str("</si>");
    });
    let dvs = each(scaled(500_000), |s, i| {
        let _ = write!(
            s,
            r#"<dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="A{i}"><formula1>"a,b,c"</formula1></dataValidation>"#
        );
    });
    let zip = package(vec![
        ("xl/sharedStrings.xml", padded_part("sst", "", &rich)),
        (
            "xl/styles.xml",
            padded_part("styleSheet", "", &real_styles(FLOOD_SCALE)),
        ),
        (
            "xl/worksheets/sheet1.xml",
            padded_part(
                "worksheet",
                "",
                &format!("<sheetData/><dataValidations>{dvs}</dataValidations>"),
            ),
        ),
    ]);
    drop((rich, dvs));
    let label = "200,000 rich entries of five runs, real styles, 500,000 data validations";
    let wb =
        read_xlsx_with_options(&zip, &limits).unwrap_or_else(|e| panic!("{label} (model): {e}"));
    let table = wb.shared_strings.as_ref().expect("table kept");
    assert_eq!(table.rich_runs.len(), scaled(200_000), "{label}");
    assert!(
        table
            .rich_runs
            .iter()
            .all(|r| r.as_ref().map(Vec::len) == Some(5))
    );
    assert_eq!(
        wb.sheets[0].worksheet.data_validations.len(),
        scaled(500_000),
        "{label}"
    );
    drop(wb);
    read_xlsx_json_with_options(&zip, &limits).unwrap_or_else(|e| panic!("{label} (json): {e}"));
}

/// The read counts the shared strings, the styles, the charts and the sheets
/// in one budget: a styles part whose `<xf/>` records take about 0.6 of the
/// limit (120 bytes each on a 64-bit target), a shared string entry whose
/// `<r></r>` runs do (248 bytes each), or a chart whose `<ser></ser>` series
/// do (384 bytes each), together with a sheet whose `<mergeCell/>` elements
/// take about 0.6 of it (26 bytes each) end in `SizeLimit` on both paths,
/// while each of them alone reads.
fn budget_carry_cases() {
    let lists = (LIST_PEAK_LIMIT, LIST_PEAK_LIMIT);
    let styles = (
        "xl/styles.xml",
        padded_part(
            "styleSheet",
            "",
            &format!("<cellXfs>{}</cellXfs>", "<xf/>".repeat(scaled(2_700_000))),
        ),
    );
    let sst = (
        "xl/sharedStrings.xml",
        // Entry 0, which the base sheet's one cell names, is plain.
        padded_part(
            "sst",
            "",
            &format!(
                "<si><t>x</t></si><si>{}</si>",
                "<r></r>".repeat(scaled(1_300_000))
            ),
        ),
    );
    let merges = (
        "xl/worksheets/sheet1.xml",
        padded_part(
            "worksheet",
            "",
            &format!(
                "<sheetData/><mergeCells>{}</mergeCells>",
                r#"<mergeCell ref="A1"/>"#.repeat(scaled(12_400_000))
            ),
        ),
    );
    let cases = [
        ("2,700,000 <xf/> alone", vec![styles.clone()], Flood::Reads),
        (
            "1,300,000 <r></r> in one <si> alone",
            vec![sst.clone()],
            Flood::Reads,
        ),
        (
            "12,400,000 <mergeCell/> alone",
            vec![merges.clone()],
            Flood::Reads,
        ),
        (
            "2,700,000 <xf/> with 12,400,000 <mergeCell/>",
            vec![styles, merges.clone()],
            Flood::ListLimit,
        ),
        (
            "1,300,000 <r></r> in one <si> with 12,400,000 <mergeCell/>",
            vec![sst, merges.clone()],
            Flood::ListLimit,
        ),
    ];
    for (label, parts, expect) in cases {
        flood_case_within(label, &package(parts), expect, lists, BREAK_FLOOD_LIMIT);
    }
    let series = format!(
        "<barChart>{}</barChart>",
        "<ser></ser>".repeat(scaled(840_000))
    );
    for (label, parts, expect) in [
        ("840,000 <ser></ser> in a chart alone", vec![], Flood::Reads),
        (
            "840,000 <ser></ser> in a chart with 12,400,000 <mergeCell/>",
            vec![merges],
            Flood::ListLimit,
        ),
    ] {
        flood_case_within(
            label,
            &chart_package_with(&series, parts),
            expect,
            lists,
            BREAK_FLOOD_LIMIT,
        );
    }
}

/// A generous real `xl/styles.xml`: 64,000 cell formats, 10,000 fonts, fills,
/// borders and differential formats, and 2,000 number formats (Excel's own
/// limit is 64,000 cell formats), in the shapes Excel writes.
fn real_styles(scale: usize) -> String {
    let (xfs, records, fmts) = (64_000 / scale, 10_000 / scale, 2_000 / scale);
    let mut s = format!(r#"<numFmts count="{fmts}">"#);
    s += &each(fmts, |s, i| {
        let _ = write!(
            s,
            r##"<numFmt numFmtId="{}" formatCode="#,##0.{i:0>3} &quot;EUR&quot;;[Red]\-#,##0.00 &quot;EUR&quot;"/>"##,
            163 + i
        );
    });
    let _ = write!(s, r#"</numFmts><fonts count="{records}">"#);
    s += &each(records, |s, i| {
        let _ = write!(
            s,
            r#"<font><b/><i/><sz val="{}"/><color theme="1" tint="-0.249977111117893"/><name val="Calibri Light {i}"/><family val="2"/><scheme val="minor"/></font>"#,
            8 + i % 30
        );
    });
    let _ = write!(s, r#"</fonts><fills count="{records}">"#);
    s += &each(records, |s, i| {
        if i % 10 == 0 {
            s.push_str(r#"<fill><gradientFill degree="90"><stop position="0"><color theme="0"/></stop><stop position="1"><color theme="4"/></stop></gradientFill></fill>"#);
        } else {
            let _ = write!(
                s,
                r#"<fill><patternFill patternType="solid"><fgColor rgb="FF{:06X}"/><bgColor indexed="64"/></patternFill></fill>"#,
                i * 97
            );
        }
    });
    let _ = write!(s, r#"</fills><borders count="{records}">"#);
    s += &each(records, |s, i| {
        let _ = write!(
            s,
            r#"<border><left style="thin"><color rgb="FF{i:06X}"/></left><right style="thin"><color rgb="FF{i:06X}"/></right><top style="medium"><color indexed="64"/></top><bottom style="double"><color auto="1"/></bottom><diagonal/></border>"#
        );
    });
    let _ = write!(
        s,
        r#"</borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="{xfs}">"#
    );
    s += &each(xfs, |s, i| {
        let _ = write!(
            s,
            r#"<xf numFmtId="{}" fontId="{}" fillId="{}" borderId="{}" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1" applyProtection="1"><alignment horizontal="center" vertical="center" wrapText="1"/><protection locked="0"/></xf>"#,
            164 + i % fmts,
            i % records,
            i % records,
            i % records
        );
    });
    let _ = write!(
        s,
        r#"</cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles><dxfs count="{records}">"#
    );
    s += &each(records, |s, i| {
        let _ = write!(
            s,
            r#"<dxf><font><b val="0"/><i/><color rgb="FF9C{i:04X}"/></font><numFmt numFmtId="{}" formatCode="0.00%"/><fill><patternFill><bgColor rgb="FFFFC7CE"/></patternFill></fill><alignment horizontal="left"/><border><left style="thin"><color rgb="FF000000"/></left><right style="thin"><color rgb="FF000000"/></right><top style="thin"><color rgb="FF000000"/></top><bottom style="thin"><color rgb="FF000000"/></bottom></border></dxf>"#,
            3_000 + i
        );
    });
    s.push_str(r#"</dxfs><tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"/>"#);
    s
}

/// The generous real styles part reads on both paths at the default limit
/// (in an unoptimized build a tenth of it at a tenth of the limit), with
/// every record kept.
fn real_styles_case() {
    let limits = flood_limits();
    let zip = styles_package(&real_styles(FLOOD_SCALE));
    let wb = read_xlsx_with_options(&zip, &limits)
        .unwrap_or_else(|e| panic!("real styles part (model): {e}"));
    let s = &wb.styles;
    let (xfs, records, fmts) = (scaled(64_000), scaled(10_000), scaled(2_000));
    assert_eq!(
        [
            s.cell_xfs.len(),
            s.fonts.len(),
            s.fills.len(),
            s.borders.len(),
            s.dxfs.len(),
            s.num_fmts.len()
        ],
        [xfs, records, records, records, records, fmts],
        "real styles part"
    );
    assert_eq!(
        s.fills[9].gradient_fill.as_ref().map(|g| g.stops.len()),
        Some(2)
    );
    assert_eq!(
        s.dxfs[0].font.as_ref().map(|f| f.explicit_off.len()),
        Some(1)
    );
    drop(wb);
    read_xlsx_json_with_options(&zip, &limits)
        .unwrap_or_else(|e| panic!("real styles part (json): {e}"));
}

/// The elements a real-sheet case keeps, counted from the parsed sheet.
type Count = fn(&WorksheetXml) -> usize;

/// A generous real sheet of each list reads on both paths at the default
/// limit (in an unoptimized build a tenth of each sheet at a tenth of the
/// limit), with every element kept.
fn real_list_cases() {
    let n = scaled(1_000_000);
    let limits = flood_limits();
    let cases: [(&str, String, Count, usize); 5] = [
        (
            "1,000,000 list data validations",
            format!(
                "<sheetData/><dataValidations>{}</dataValidations>",
                each(n, |s, i| {
                    let _ = write!(
                        s,
                        r#"<dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="A{i}"><formula1>"a,b,c"</formula1></dataValidation>"#
                    );
                })
            ),
            |w| w.data_validations.len(),
            n,
        ),
        (
            "1,000,000 conditional formats of one cellIs rule",
            format!(
                "<sheetData/>{}",
                each(n, |s, i| {
                    let _ = write!(
                        s,
                        r#"<conditionalFormatting sqref="A{i}"><cfRule type="cellIs" dxfId="0" priority="{i}" operator="greaterThan"><formula>0</formula></cfRule></conditionalFormatting>"#
                    );
                })
            ),
            |w| w.conditional_formatting.iter().map(|c| c.rules.len()).sum(),
            n,
        ),
        (
            "16,384 columns",
            format!(
                "<cols>{}</cols><sheetData/>",
                each(16_384, |s, i| {
                    let _ = write!(
                        s,
                        r#"<col min="{i}" max="{i}" width="12.7109375" style="1" customWidth="1"/>"#
                    );
                })
            ),
            |w| w.columns.len(),
            16_384,
        ),
        (
            "1,000,000 internal hyperlinks",
            format!(
                "<sheetData/><hyperlinks>{}</hyperlinks>",
                each(n, |s, i| {
                    let _ = write!(
                        s,
                        r#"<hyperlink ref="A{i}" location="'S1'!B{i}" display="Go to B{i}"/>"#
                    );
                })
            ),
            |w| w.hyperlinks.len(),
            n,
        ),
        (
            "1,000,000 merged ranges",
            format!(
                "<sheetData/><mergeCells>{}</mergeCells>",
                each(n, |s, i| {
                    let _ = write!(s, r#"<mergeCell ref="A{i}:B{i}"/>"#);
                })
            ),
            |w| w.merge_cells.len(),
            n,
        ),
    ];
    for (label, body, count, expected) in cases {
        let zip = flood_package(&body);
        drop(body);
        let wb = read_xlsx_with_options(&zip, &limits)
            .unwrap_or_else(|e| panic!("{label} (model): {e}"));
        assert_eq!(count(&wb.sheets[0].worksheet), expected, "{label}");
        drop(wb);
        read_xlsx_json_with_options(&zip, &limits)
            .unwrap_or_else(|e| panic!("{label} (json): {e}"));
    }
}

/// `inner` after a comment of random text a 25th of its size, which keeps
/// the part under the 100:1 compression ratio.
fn padded(inner: &str) -> String {
    format!("<!--{}-->{inner}", text(inner.len() / 25 + 1_000))
}

/// A package whose sheet names, by its `.rels`, one part of the
/// relationship type `rel_type` at `target` (relative to the sheet), with
/// the parts `set` added.
fn sheet_part_package(rel_type: &str, target: &str, set: Vec<(&str, String)>) -> Vec<u8> {
    let rels = format!(
        r#"<Relationships xmlns="{PKG_REL_NS}"><Relationship Id="rIdX1" Type="{rel_type}" Target="{target}"/></Relationships>"#
    );
    let mut set = set;
    set.push(("xl/worksheets/_rels/sheet1.xml.rels", rels));
    package(set)
}

/// A package whose sheet holds one drawing with one chart, `chart_inner`
/// inside the `<plotArea>` of `xl/charts/chart1.xml`.
fn chart_package(chart_inner: &str) -> Vec<u8> {
    chart_package_with(chart_inner, vec![])
}

/// [`chart_package`] with the parts `set` added.
fn chart_package_with(chart_inner: &str, set: Vec<(&str, String)>) -> Vec<u8> {
    const CHART_NS: &str = "http://schemas.openxmlformats.org/drawingml/2006/chart";
    const DRAWING_NS: &str = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
    let drawing = format!(
        r#"<wsDr xmlns="{DRAWING_NS}" xmlns:r="{REL_NS}"><twoCellAnchor><from><col>0</col><colOff>0</colOff><row>0</row><rowOff>0</rowOff></from><to><col>5</col><colOff>0</colOff><row>10</row><rowOff>0</rowOff></to><graphicFrame><chart r:id="rIdC1"/></graphicFrame><clientData/></twoCellAnchor></wsDr>"#
    );
    let drawing_rels = format!(
        r#"<Relationships xmlns="{PKG_REL_NS}"><Relationship Id="rIdC1" Type="{REL_NS}/chart" Target="../charts/chart1.xml"/></Relationships>"#
    );
    let chart = format!(
        r#"<chartSpace xmlns="{CHART_NS}"><chart><plotArea>{}</plotArea></chart></chartSpace>"#,
        padded(chart_inner)
    );
    let mut set = set;
    set.extend([
        ("xl/drawings/drawing1.xml", drawing),
        ("xl/drawings/_rels/drawing1.xml.rels", drawing_rels),
        ("xl/charts/chart1.xml", chart),
    ]);
    sheet_part_package(
        &format!("{REL_NS}/drawing"),
        "../drawings/drawing1.xml",
        set,
    )
}

/// A package whose sheet names one threaded comments part holding `inner`.
fn threaded_package(inner: &str) -> Vec<u8> {
    let part = format!(
        r#"<ThreadedComments xmlns="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments">{}</ThreadedComments>"#,
        padded(inner)
    );
    sheet_part_package(
        "http://schemas.microsoft.com/office/2017/10/relationships/threadedComment",
        "../threadedComments/threadedComment1.xml",
        vec![("xl/threadedComments/threadedComment1.xml", part)],
    )
}

/// A package whose `xl/workbook.xml` holds `<definedNames>{inner}</definedNames>`.
fn defined_names_package(inner: &str) -> Vec<u8> {
    let workbook = base_part("xl/workbook.xml").replacen(
        "</sheets>",
        &format!(
            "</sheets>{}",
            padded(&format!("<definedNames>{inner}</definedNames>"))
        ),
        1,
    );
    package(vec![("xl/workbook.xml", workbook)])
}

/// The records of a chart part, of a threaded comments part and of the
/// workbook's defined names are list elements too: millions of empty
/// series (in the first and in the second chart type), category and value
/// axes, threaded comments, mentions in one threaded comment and defined
/// names end in `SizeLimit`, each at its record size plus its strings
/// (384, 216, 176, 56 and 216 bytes on a 64-bit target), within the time
/// and peak bounds, defined names with text as well as empty ones; a tenth
/// of the series, half of the threaded comments and 1,000,000 defined names
/// read.
fn part_flood_cases() {
    let lists = (LIST_PEAK_LIMIT, LIST_PEAK_LIMIT);
    let charts: [(&str, &str, &str, usize, Flood); 5] = [
        (
            "<barChart>",
            "<ser></ser>",
            "</barChart>",
            2_000_000,
            Flood::ListLimit,
        ),
        (
            "<barChart></barChart><lineChart>",
            "<ser></ser>",
            "</lineChart>",
            2_000_000,
            Flood::ListLimit,
        ),
        (
            "<barChart>",
            "<ser><tx/></ser>",
            "</barChart>",
            200_000,
            Flood::Reads,
        ),
        (
            "<barChart/>",
            "<catAx></catAx>",
            "",
            3_000_000,
            Flood::ListLimit,
        ),
        (
            "<barChart/>",
            "<valAx></valAx>",
            "",
            3_000_000,
            Flood::ListLimit,
        ),
    ];
    for (open, record, close, n, expect) in charts {
        flood_case(
            &format!("{n} {record} in a chart's {open}"),
            &chart_package(&format!("{open}{}{close}", record.repeat(scaled(n)))),
            expect,
            lists,
        );
    }
    for (n, expect) in [(4_000_000, Flood::ListLimit), (2_000_000, Flood::Reads)] {
        flood_case(
            &format!("{n} <threadedComment/>"),
            &threaded_package(&"<threadedComment/>".repeat(scaled(n))),
            expect,
            lists,
        );
    }
    flood_case_within(
        "12000000 <mention/> in one <threadedComment>",
        &threaded_package(&format!(
            "<threadedComment>{}</threadedComment>",
            "<mention/>".repeat(scaled(12_000_000))
        )),
        Flood::ListLimit,
        lists,
        BREAK_FLOOD_LIMIT,
    );
    for (n, expect) in [(3_000_000, Flood::ListLimit), (1_000_000, Flood::Reads)] {
        flood_case(
            &format!(r#"{n} <definedName name="a"/>"#),
            &defined_names_package(&r#"<definedName name="a"/>"#.repeat(scaled(n))),
            expect,
            lists,
        );
    }
    // A defined name with text, the form every real one takes, is counted
    // when it ends rather than when it starts.
    let n = 3_000_000;
    flood_case(
        &format!(r#"{n} <definedName name="a">A1</definedName>"#),
        &defined_names_package(&r#"<definedName name="a">A1</definedName>"#.repeat(scaled(n))),
        Flood::ListLimit,
        lists,
    );
}

#[test]
fn shared_text_expansion_ends_in_size_limit_within_bounds() {
    hostile_case(
        "r2: shared string of 40,000 control characters x 16,384 cells of one row",
        &sst_row_package(&format!("<t>{}</t>", control_refs(40_000)), 16_384, 1),
    );
    hostile_case(
        "r1: shared string of 40,000 backslashes and quotes x 16,384 cells of one row",
        &sst_row_package(&format!("<t>{}</t>", mix(40_000, &["\\", "\""])), 16_384, 1),
    );
    json_case(
        "author of 40,000 control characters x 2,500 comments",
        &author_xml_package(&control_refs(40_000), 2_500),
    );
    json_case(
        "hyperlink target of 40,000 control characters x 2,500 hyperlinks",
        &hyperlink_xml_package(&control_refs(40_000), 2_500),
    );
    json_case_within(
        "author of 40,000 control characters x 12,000 comments",
        &author_xml_package(&control_refs(40_000), 12_000),
        JSON_COPIES_PEAK_LIMIT,
    );
    json_case(
        "one cell with 13,000 <v> naming a 40,000-control-char string",
        &sst_padded_cell_package(&control_refs(40_000), 13_000),
    );
    let font = control_refs(1_000);
    let runs: Vec<String> = (1_u32..32)
        .filter(|c| ![9, 10, 13].contains(c))
        .map(|c| format!(r#"<r><rPr><rFont val="{font}"/></rPr><t>&#{c};</t></r>"#))
        .collect();
    json_case(
        "rich entry of 100 one-control-character runs, each with a font name of 1,000 control characters, x 16,384 cells of one row",
        &sst_row_package(
            &mix(100, &runs.iter().map(String::as_str).collect::<Vec<_>>()),
            16_384,
            1,
        ),
    );
    hostile_case(
        "namespace of 100,000 chars re-declared on 6,000 extensions",
        &ext_package(100_000, 6_000),
    );
    exact_limit_case();
    several_values_case();
    hostile_case(
        "t05: shared string of 100,000 chars x 10,000 cells",
        &sst_package(100_000, 10_000),
    );
    hostile_case(
        "t10: shared string of 100,000 chars x 30,000 cells",
        &sst_package(100_000, 30_000),
    );
    hostile_case(
        "t04: author of 100,000 chars x 10,000 comments",
        &author_package(100_000, 10_000),
    );
    hostile_case(
        "t10: author of 100,000 chars x 30,000 comments",
        &author_package(100_000, 30_000),
    );
    hostile_case(
        "hyperlink target of 100,000 chars x 10,000 hyperlinks",
        &hyperlink_package(100_000, 10_000),
    );
    under_budget_case();
    custom_limit_case();
    element_flood_cases();
    sink_list_cases();
    real_list_cases();
    style_flood_cases();
    real_styles_case();
    namespace_cases();
    sst_flood_cases();
    real_sst_cases();
    budget_carry_cases();
    part_flood_cases();
}

/// The `<styleSheet>` of [`repeated_verbatim_elements_keep_the_last`]: two
/// `<tableStyles>`, two `<colors>` and two `<extLst>`, the first of each
/// using `p1` and the second `p2`.
const REPEATED_STYLES: &str = concat!(
    r#"<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:p1="urn:first" xmlns:p2="urn:second">"#,
    r#"<tableStyles count="0" p1:a="1"/>"#,
    r#"<colors><p1:a/></colors>"#,
    r#"<extLst><ext uri="{first}"><p1:a/></ext></extLst>"#,
    "\n  ",
    r#"<tableStyles count="1"><p2:b/></tableStyles>"#,
    "\n  ",
    r#"<colors><indexedColors/><p2:b/></colors>"#,
    r#"<extLst><ext uri="{second}"><p2:b/></ext></extLst>"#,
    "</styleSheet>",
);

/// A pivot cache definition with two `<cacheSource>` kept whole, the first
/// using `p1` and the second `p2`.
fn repeated_cache_sources() -> String {
    format!(
        concat!(
            r#"<pivotCacheDefinition xmlns="{NS}" xmlns:r="{REL_NS}" xmlns:p1="urn:first" xmlns:p2="urn:second">"#,
            r#"<cacheSource type="external" connectionId="1"><p1:a/></cacheSource>"#,
            r#"<cacheSource type="consolidation"><p2:b r:id="rId9"/></cacheSource>"#,
            r#"<cacheFields count="0"/></pivotCacheDefinition>"#,
        ),
        NS = NS,
        REL_NS = REL_NS,
    )
}

/// A `<styleSheet>` holding a `<colors>` that uses `p1`, followed by a
/// second root that binds `p1` to another URI and holds no `<colors>`.
const TWO_ROOT_STYLES: &str = concat!(
    r#"<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:p1="urn:first">"#,
    r#"<colors><p1:a/></colors></styleSheet>"#,
    r#"<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:p1="urn:second">"#,
    "</styleSheet>",
);

/// A pivot cache definition with a `<cacheSource>` kept whole that uses
/// `p1`, followed by a second root that binds `p1` to another URI and holds
/// no `<cacheSource>`.
fn two_root_cache_sources() -> String {
    format!(
        concat!(
            r#"<pivotCacheDefinition xmlns="{NS}" xmlns:r="{REL_NS}" xmlns:p1="urn:first">"#,
            r#"<cacheSource type="external" connectionId="1"><p1:a/></cacheSource>"#,
            r#"<cacheFields count="0"/></pivotCacheDefinition>"#,
            r#"<pivotCacheDefinition xmlns="{NS}" xmlns:p1="urn:second"></pivotCacheDefinition>"#,
        ),
        NS = NS,
        REL_NS = REL_NS,
    )
}

/// `json` holds each of `kept` as a JSON string and no trace of the first
/// occurrences (`urn:first`).
fn assert_json_keeps(label: &str, json: &str, kept: &[&str]) {
    for fragment in kept {
        let quoted = serde_json::to_string(fragment).expect("JSON string");
        assert!(json.contains(&quoted), "{label}: JSON keeps {fragment}");
    }
    assert!(!json.contains("urn:first"), "{label}: JSON drops the first");
}

/// A part that repeats a verbatim element keeps the last occurrence, made
/// self-contained against its root, on both read paths, with the strings a
/// build that made every occurrence self-contained gives (the same FNV-1a of
/// each JSON, checked 2026-11-13). An element kept under a first root is made
/// self-contained against that root, not against a second root that rebinds
/// its prefix.
#[test]
fn repeated_verbatim_elements_keep_the_last() {
    let limits = ZipSecurityLimits::default();
    let table_styles = r#"<tableStyles xmlns:p2="urn:second" count="1"><p2:b/></tableStyles>"#;
    let colors = r#"<colors xmlns:p2="urn:second"><indexedColors/><p2:b/></colors>"#;
    let ext_lst = r#"<extLst xmlns:p2="urn:second"><ext uri="{second}"><p2:b/></ext></extLst>"#;
    let zip = package(vec![("xl/styles.xml", REPEATED_STYLES.to_owned())]);
    let wb = read_xlsx_with_options(&zip, &limits).expect("styles model");
    assert_eq!(wb.styles.table_styles_xml.as_deref(), Some(table_styles));
    assert_eq!(wb.styles.colors_xml.as_deref(), Some(colors));
    assert_eq!(wb.styles.ext_lst_xml.as_deref(), Some(ext_lst));
    let json = read_xlsx_json_with_options(&zip, &limits).expect("styles json");
    assert_json_keeps("styles", &json, &[table_styles, colors, ext_lst]);

    let source = r#"<cacheSource xmlns:p2="urn:second" type="consolidation"><p2:b r:id="rId9"/></cacheSource>"#;
    let zip = pivot_cache_package(repeated_cache_sources());
    let wb = read_xlsx_with_options(&zip, &limits).expect("pivot model");
    let raw = wb
        .pivot_caches
        .first()
        .and_then(|c| c.source.raw_xml.as_deref());
    assert_eq!(raw, Some(source));
    let json = read_xlsx_json_with_options(&zip, &limits).expect("pivot json");
    assert_json_keeps("pivot", &json, &[source]);

    let colors = r#"<colors xmlns:p1="urn:first"><p1:a/></colors>"#;
    let zip = package(vec![("xl/styles.xml", TWO_ROOT_STYLES.to_owned())]);
    let wb = read_xlsx_with_options(&zip, &limits).expect("two-root styles model");
    assert_eq!(wb.styles.colors_xml.as_deref(), Some(colors));
    let json = read_xlsx_json_with_options(&zip, &limits).expect("two-root styles json");
    let quoted = serde_json::to_string(colors).expect("JSON string");
    assert!(
        json.contains(&quoted),
        "two-root styles: JSON keeps {colors}"
    );
    assert!(
        !json.contains("urn:second"),
        "two-root styles: no second URI"
    );

    let source = r#"<cacheSource xmlns:p1="urn:first" type="external" connectionId="1"><p1:a/></cacheSource>"#;
    let zip = pivot_cache_package(two_root_cache_sources());
    let wb = read_xlsx_with_options(&zip, &limits).expect("two-root pivot model");
    let raw = wb
        .pivot_caches
        .first()
        .and_then(|c| c.source.raw_xml.as_deref());
    assert_eq!(raw, Some(source));
    let json = read_xlsx_json_with_options(&zip, &limits).expect("two-root pivot json");
    let quoted = serde_json::to_string(source).expect("JSON string");
    assert!(
        json.contains(&quoted),
        "two-root pivot: JSON keeps {source}"
    );
    assert!(
        !json.contains("urn:second"),
        "two-root pivot: no second URI"
    );
}

/// The sheet of a package whose `<worksheet>` declares `k` prefixes
/// `xmlns:p0="u"` to `xmlns:p{k-1}` and whose `<extLst>` holds `m` distinct
/// `<ext uri="u{i}"/>`, each using every prefix in an attribute: each kept
/// `<ext>` inherits `k` declarations. With `own`, each `<ext>` declares the
/// prefixes itself, so nothing is inserted and the kept text is the same.
fn inherited_ext_package(k: usize, m: usize, own: bool) -> Vec<u8> {
    let mut decls = String::new();
    let mut uses = String::new();
    for i in 0..k {
        let _ = write!(decls, r#" xmlns:p{i}="u""#);
        let _ = write!(uses, r#" p{i}:a="1""#);
    }
    let on_ext = if own { decls.as_str() } else { "" };
    let exts = each(m, |s, i| {
        let _ = write!(s, r#"<ext{on_ext} uri="u{i}"{uses}/>"#);
    });
    package(vec![(
        "xl/worksheets/sheet1.xml",
        padded_part(
            "worksheet",
            &decls,
            &format!("<sheetData/><extLst>{exts}</extLst>"),
        ),
    )])
}

/// Time bound of the read of [`inherited_ext_package`] with 1,000 root
/// prefixes used by every kept `<ext>`. Each inserted declaration is a binary
/// search of the root index and a copy of text it rendered once. On the test
/// host (an Intel Core Ultra 9 290HX Plus laptop, Windows 11, measured
/// 2026-11-13) 11,500 elements (11,500,000 inserted declarations) took 1.91 s
/// on the model path and 2.26 s on the JSON path optimized, where formatting
/// and escaping each declaration took 2.88 s and 3.7 s: the optimized bound,
/// 3 s, fails the JSON path of that. Unoptimized, the XML reader dominates
/// both, 1,150 elements took 1.7 s and 2.4 s (2.4 s and 3.4 s formatting each
/// declaration), so the unoptimized bound, 10 s, only guards against a
/// return to a cost per root declaration.
const INHERITED_EXT_LIMIT: Duration = if cfg!(debug_assertions) {
    Duration::from_secs(10)
} else {
    Duration::from_secs(3)
};

/// Kept extensions that each inherit 1,000 root declarations read on both
/// paths within [`INHERITED_EXT_LIMIT`], each declaration the text the
/// element would declare itself. Run in the sequence of
/// [`shared_text_expansion_ends_in_size_limit_within_bounds`], so that the
/// peak it measures is its own.
fn inherited_declaration_case() {
    // The JSON path holds the JSON (343 MB optimized) beside the kept text.
    let peaks = (MODEL_PEAK_LIMIT, JSON_COPIES_PEAK_LIMIT);
    let (k, m) = (1_000, scaled(11_500));
    let zip = inherited_ext_package(k, m, false);
    flood_case_counted(
        &format!(
            "{k} root prefixes used by each of {m} kept <ext>: {} inserted declarations",
            k * m
        ),
        &zip,
        Flood::Reads,
        peaks,
        INHERITED_EXT_LIMIT,
        1,
    );
    let wb = read_xlsx_with_options(&zip, &flood_limits()).expect("worksheet model");
    let kept = &wb.sheets[0].worksheet.preserved_extensions;
    assert_eq!(kept.len(), m);
    let own = inherited_ext_package(k, 1, true);
    let own_wb = read_xlsx(&own).expect("own declarations");
    assert_eq!(
        kept[0], own_wb.sheets[0].worksheet.preserved_extensions[0],
        "an inherited declaration is the text the element would declare itself"
    );
}

/// The `xl/worksheets/sheet1.xml` of `zip`.
fn sheet_xml(zip: &[u8]) -> String {
    read_zip_entries(zip, &ZipSecurityLimits::default())
        .expect("written package")
        .into_iter()
        .find(|(n, _)| n == "xl/worksheets/sheet1.xml")
        .map(|(_, d)| String::from_utf8(d).expect("UTF-8 sheet"))
        .expect("sheet1.xml")
}

/// Every element and attribute prefix of `xml` is bound (Namespaces in XML
/// 1.0, §5, Namespace Constraint: Prefix Declared).
fn assert_prefixes_bound(label: &str, xml: &str) {
    use quick_xml::NsReader;
    use quick_xml::events::Event;
    use quick_xml::name::ResolveResult;
    let mut reader = NsReader::from_str(xml);
    loop {
        match reader.read_event().expect("well-formed sheet") {
            Event::Start(e) | Event::Empty(e) => {
                let (ns, _) = reader.resolver().resolve_element(e.name());
                assert!(
                    !matches!(ns, ResolveResult::Unknown(_)),
                    "{label}: <{}> has an unbound prefix in {xml}",
                    e.name().as_ref()
                );
                for a in e.attributes() {
                    let a = a.expect("attribute");
                    let (ns, _) = reader.resolver().resolve_attribute(a.key);
                    assert!(
                        !matches!(ns, ResolveResult::Unknown(_)),
                        "{label}: @{} has an unbound prefix in {xml}",
                        a.key.as_ref()
                    );
                }
            }
            Event::Eof => break,
            _ => {}
        }
    }
}

/// A declaration on `<extLst>`, the one element between the root and a kept
/// `<ext>`, takes part in each kept `<ext>` as the root's do and overrides
/// the root's binding of the same prefix, on both read paths; the default
/// namespace, declared on either, is never inserted (unprefixed names take
/// the default namespace of the sheet they are written into). The written
/// sheet binds every prefix it uses.
#[test]
fn ext_lst_declarations_bind_in_each_kept_ext() {
    const X14: &str = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main";
    const CF_URI: &str = "{78C0D931-6437-407d-A8EE-F0AAD7539E65}";
    for (label, root_decls, ext_lst_decls, inner, kept) in [
        (
            "x14 declared on <extLst> only",
            String::new(),
            format!(r#" xmlns:x14="{X14}""#),
            "<x14:conditionalFormattings/>",
            format!(r#"<ext xmlns:x14="{X14}" uri="{CF_URI}"><x14:conditionalFormattings/></ext>"#),
        ),
        (
            "x14 declared on the root and rebound on <extLst>",
            r#" xmlns:x14="urn:root""#.to_owned(),
            format!(r#" xmlns:x14="{X14}""#),
            "<x14:conditionalFormattings/>",
            format!(r#"<ext xmlns:x14="{X14}" uri="{CF_URI}"><x14:conditionalFormattings/></ext>"#),
        ),
        (
            "the default namespace rebound on <extLst>",
            format!(r#" xmlns:x14="{X14}""#),
            r#" xmlns="urn:other" xmlns:xm="urn:xm""#.to_owned(),
            "<a/><x14:b xm:c=\"1\"/>",
            format!(
                r#"<ext xmlns:x14="{X14}" xmlns:xm="urn:xm" uri="{CF_URI}"><a/><x14:b xm:c="1"/></ext>"#
            ),
        ),
        (
            "no declaration on <extLst>",
            format!(r#" xmlns:x14="{X14}""#),
            String::new(),
            "<x14:conditionalFormattings/>",
            format!(r#"<ext xmlns:x14="{X14}" uri="{CF_URI}"><x14:conditionalFormattings/></ext>"#),
        ),
    ] {
        let xml = format!(
            r#"<worksheet xmlns="{NS}"{root_decls}><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData><extLst{ext_lst_decls}><ext uri="{CF_URI}">{inner}</ext></extLst></worksheet>"#
        );
        let zip = package(vec![("xl/worksheets/sheet1.xml", xml)]);
        let wb = read_xlsx(&zip).unwrap_or_else(|e| panic!("{label}: {e}"));
        assert_eq!(
            wb.sheets[0].worksheet.preserved_extensions,
            std::slice::from_ref(&kept),
            "{label} (model)"
        );
        let json: serde_json::Value =
            serde_json::from_str(&read_xlsx_json(&zip).unwrap_or_else(|e| panic!("{label}: {e}")))
                .expect("JSON");
        assert_eq!(
            json["sheets"][0]["worksheet"]["preservedExtensions"],
            serde_json::json!([kept]),
            "{label} (json)"
        );
        let sheet = sheet_xml(&write_xlsx(&wb).expect("write"));
        assert!(sheet.contains(&kept), "{label}: {sheet}");
        assert_prefixes_bound(label, &sheet);
    }
}

/// A kept `<ext>` with a name that begins with a colon (`<:a/>`, `:k="1"`,
/// which quick-xml gives the empty prefix) is kept as it is: no default
/// namespace declaration is inserted for it, so an `<ext>` that declares its
/// own default namespace keeps one `xmlns`. The library reads its own output,
/// and a second read and write gives the same sheet.
#[test]
fn colon_leading_names_in_a_kept_ext_round_trip() {
    for (label, ext_lst_decls, ext) in [
        (
            "own default namespace",
            "",
            r#"<ext xmlns="urn:x" uri="u1"><:a/></ext>"#,
        ),
        (
            "no own declaration",
            "",
            r#"<ext uri="u1"><:a :k="1"/></ext>"#,
        ),
        (
            "default namespace on <extLst>",
            r#" xmlns="urn:e""#,
            r#"<ext uri="u1"><:a/></ext>"#,
        ),
    ] {
        let xml = format!(
            r#"<worksheet xmlns="{NS}"><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData><extLst{ext_lst_decls}>{ext}</extLst></worksheet>"#
        );
        let zip = package(vec![("xl/worksheets/sheet1.xml", xml)]);
        let wb = read_xlsx(&zip).unwrap_or_else(|e| panic!("{label}: {e}"));
        assert_eq!(
            wb.sheets[0].worksheet.preserved_extensions,
            [ext],
            "{label} (model)"
        );
        let json: serde_json::Value =
            serde_json::from_str(&read_xlsx_json(&zip).unwrap_or_else(|e| panic!("{label}: {e}")))
                .expect("JSON");
        assert_eq!(
            json["sheets"][0]["worksheet"]["preservedExtensions"],
            serde_json::json!([ext]),
            "{label} (json)"
        );
        let written = write_xlsx(&wb).expect("write");
        let sheet = sheet_xml(&written);
        assert!(sheet.contains(ext), "{label}: {sheet}");
        let again = read_xlsx(&written).unwrap_or_else(|e| panic!("{label}, second read: {e}"));
        assert_eq!(
            again.sheets[0].worksheet.preserved_extensions,
            [ext],
            "{label} (second read)"
        );
        let rewritten = sheet_xml(&write_xlsx(&again).expect("second write"));
        assert_eq!(rewritten, sheet, "{label}: second write");
    }
}
