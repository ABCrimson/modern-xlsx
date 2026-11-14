//! Hostile-input regression suite: crafted ZIP, OLE2 and XML inputs go
//! through the public read APIs (the two VML drawing cases, the pivot cache
//! case, the two generated-note cases and the many-tables case also back through `write_xlsx`) and must end in the exact coded error (or a clean read and
//! write), never a panic; every case except the duplicate-name, colon-name and missing-part checks also runs under a time bound.
//!
//! Every case runs sequentially inside one `#[test]`, so a slow case cannot
//! overlap another and a timing assertion measures only its own work.
//!
//! A counting `#[global_allocator]` (`peak_alloc::PeakAlloc`, a safe API, so the
//! workspace `unsafe_code = "forbid"` stays in force) bounds the peak heap of
//! the size-driven cases: each OLE2 case and each forged central-directory size
//! stays under 64 MiB, the 64 MiB zero bomb under 8 MiB, a sheet naming one
//! comments part 10,000 times under 16 MiB (32 MiB through the JSON reader),
//! and under 16 MiB on both read paths 10,000 `<sheet>` rows on one worksheet
//! part, 200 sheets naming one comments part and 200 drawings naming one chart
//! part, the write of 2,000 pivot tables on one cache of 100,000 shared
//! items under 1 MiB, and the writes of 1,000,000 bare thread roots and of
//! 1,000,000 legacy notes without a drawing under 16 MiB (their reads are
//! not timed); the ZIP-name, HMAC-tamper and DOCTYPE cases, and the read and write
//! of the two VML drawing cases and of 3,000 tables on one sheet, are timed only. The peak delta
//! also bounds the largest single allocation, so a
//! `Vec::with_capacity(declared_size)` read from a hostile header fails here
//! even where 64-bit native would return at once.
//! The counter is process-wide, which is why every case stays sequential. The
//! wasm32 traps are covered by `packages/modern-xlsx/__tests__/hostile-input.test.ts`.
#![cfg(feature = "encryption")]

use std::fmt::Write as _;
use std::time::{Duration, Instant};

use modern_xlsx_core::ModernXlsxError;
use modern_xlsx_core::ole2::cfb::reader::read_stream;
use modern_xlsx_core::ole2::cfb::writer::write_ole2;
use modern_xlsx_core::ole2::crypto::encrypt_file;
use modern_xlsx_core::reader::{read_xlsx, read_xlsx_json, read_xlsx_json_with_password};
use modern_xlsx_core::writer::write_xlsx;
use modern_xlsx_core::zip::{ZipEntry, ZipSecurityLimits, read_zip_entries, write_zip};
use modern_xlsx_core::{SheetData, WorkbookData};
use peak_alloc::PeakAlloc;

#[global_allocator]
static PEAK: PeakAlloc = PeakAlloc;

const MIB: usize = 1 << 20;

const LIMIT: Duration = Duration::from_secs(1);
const ENDOFCHAIN: u32 = 0xFFFF_FFFE;
const FATSECT: u32 = 0xFFFF_FFFD;
const FREESECT: u32 = 0xFFFF_FFFF;

// ---------------------------------------------------------------------------
// Assertion helpers
// ---------------------------------------------------------------------------

/// Asserts `result` is an error with `code`, returning its message.
fn expect_code<T>(label: &str, result: Result<T, ModernXlsxError>, code: &str) -> String {
    match result {
        Ok(_) => panic!("{label}: expected {code}, got Ok"),
        Err(e) => {
            assert_eq!(e.code(), code, "{label}: wrong code, message: {e}");
            e.to_string()
        }
    }
}

/// Runs `f` and asserts it finished within [`LIMIT`].
fn timed<T>(label: &str, f: impl FnOnce() -> T) -> T {
    timed_within(label, LIMIT, f)
}

/// Runs `f` and asserts it finished within `limit`.
fn timed_within<T>(label: &str, limit: Duration, f: impl FnOnce() -> T) -> T {
    let start = Instant::now();
    let out = f();
    let elapsed = start.elapsed();
    assert!(
        elapsed < limit,
        "{label}: took {elapsed:?}, limit {limit:?}"
    );
    out
}

/// Runs `f` under [`timed`] and returns its result with the peak heap growth
/// (bytes above the usage at entry) it caused. Build inputs before calling it.
fn measured<T>(label: &str, f: impl FnOnce() -> T) -> (T, usize) {
    PEAK.reset_peak_usage();
    let base = PEAK.current_usage();
    let out = timed(label, f);
    (out, PEAK.peak_usage().saturating_sub(base))
}

/// Asserts a measured peak stayed under `limit` bytes.
fn assert_peak(label: &str, peak: usize, limit: usize) {
    assert!(
        peak < limit,
        "{label}: peak heap growth {peak} bytes, limit {limit} bytes"
    );
}

// ---------------------------------------------------------------------------
// ZIP builders
// ---------------------------------------------------------------------------

fn crc32(data: &[u8]) -> u32 {
    let mut crc = 0xFFFF_FFFF_u32;
    for &byte in data {
        crc ^= u32::from(byte);
        for _ in 0..8 {
            crc = if crc & 1 == 1 {
                (crc >> 1) ^ 0xEDB8_8320
            } else {
                crc >> 1
            };
        }
    }
    !crc
}

/// Builds a ZIP of stored (method 0) entries byte by byte, so a name may repeat.
fn stored_zip(entries: &[(&[u8], &[u8])]) -> Vec<u8> {
    let mut out = Vec::new();
    let mut central = Vec::new();
    for &(name, data) in entries {
        let offset = out.len() as u32;
        let crc = crc32(data);
        let size = data.len() as u32;
        out.extend_from_slice(&0x0403_4B50_u32.to_le_bytes());
        out.extend_from_slice(&[20, 0, 0, 0, 0, 0, 0, 0, 0, 0]); // version, flags, method, time, date
        out.extend_from_slice(&crc.to_le_bytes());
        out.extend_from_slice(&size.to_le_bytes());
        out.extend_from_slice(&size.to_le_bytes());
        out.extend_from_slice(&(name.len() as u16).to_le_bytes());
        out.extend_from_slice(&[0, 0]);
        out.extend_from_slice(name);
        out.extend_from_slice(data);

        central.extend_from_slice(&0x0201_4B50_u32.to_le_bytes());
        central.extend_from_slice(&[20, 0, 20, 0, 0, 0, 0, 0, 0, 0, 0, 0]);
        central.extend_from_slice(&crc.to_le_bytes());
        central.extend_from_slice(&size.to_le_bytes());
        central.extend_from_slice(&size.to_le_bytes());
        central.extend_from_slice(&(name.len() as u16).to_le_bytes());
        central.extend_from_slice(&[0; 8]); // extra, comment, disk, int attrs
        central.extend_from_slice(&[0; 4]); // external attrs
        central.extend_from_slice(&offset.to_le_bytes());
        central.extend_from_slice(name);
    }
    let cd_offset = out.len() as u32;
    out.extend_from_slice(&central);
    out.extend_from_slice(&0x0605_4B50_u32.to_le_bytes());
    out.extend_from_slice(&[0, 0, 0, 0]);
    out.extend_from_slice(&(entries.len() as u16).to_le_bytes());
    out.extend_from_slice(&(entries.len() as u16).to_le_bytes());
    out.extend_from_slice(&(central.len() as u32).to_le_bytes());
    out.extend_from_slice(&cd_offset.to_le_bytes());
    out.extend_from_slice(&[0, 0]);
    out
}

fn text_workbook(cells: usize) -> WorkbookData {
    use modern_xlsx_core::ooxml::worksheet::{Cell, CellType, Row, WorksheetXml};
    let rows = (1..=cells as u32)
        .map(|r| Row {
            index: r,
            cells: vec![Cell {
                reference: format!("A{r}"),
                cell_type: CellType::SharedString,
                value: Some(format!("row {r} hostile-input fixture text")),
                ..Default::default()
            }],
            ..Default::default()
        })
        .collect();
    WorkbookData {
        sheets: vec![SheetData {
            name: "S".into(),
            worksheet: WorksheetXml {
                rows,
                ..Default::default()
            },
            ..Default::default()
        }],
        ..Default::default()
    }
}

fn plain_xlsx() -> Vec<u8> {
    write_xlsx(&text_workbook(3)).expect("write fixture")
}

/// The parts of a real workbook, to graft hostile entries onto.
fn base_parts() -> Vec<ZipEntry> {
    let entries = read_zip_entries(&plain_xlsx(), &ZipSecurityLimits::default()).expect("read");
    let mut parts: Vec<ZipEntry> = entries
        .into_iter()
        .map(|(name, data)| ZipEntry { name, data })
        .collect();
    parts.sort_by(|a, b| a.name.cmp(&b.name));
    parts
}

// ---------------------------------------------------------------------------
// OLE2 builder
// ---------------------------------------------------------------------------

struct Stream<'a> {
    name: &'a str,
    data: Vec<u8>,
    /// Size written to the directory entry, when it should differ from `data.len()`.
    declared: Option<u64>,
}

/// Builds a one-FAT-sector compound file with a valid header (BOM FFFE, mini
/// shift 6, cutoff 4096) and no MiniFAT (the legacy layout, header@60 =
/// ENDOFCHAIN), so every stream is read through the regular FAT whatever its size.
fn build_ole2(major: u16, streams: &[Stream<'_>], self_loop_dir: bool) -> Vec<u8> {
    let (shift, ss) = if major == 4 {
        (12_u16, 4096)
    } else {
        (9_u16, 512)
    };
    let mut next = 2_u32;
    let chains: Vec<Vec<u32>> = streams
        .iter()
        .map(|s| {
            let count = s.data.len().div_ceil(ss).max(1);
            (0..count)
                .map(|_| {
                    next += 1;
                    next - 1
                })
                .collect()
        })
        .collect();
    let mut buf = vec![0_u8; ss * (1 + next as usize)];
    let put32 =
        |buf: &mut [u8], at: usize, v: u32| buf[at..at + 4].copy_from_slice(&v.to_le_bytes());

    buf[..8].copy_from_slice(&[0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1]);
    buf[24..26].copy_from_slice(&62_u16.to_le_bytes());
    buf[26..28].copy_from_slice(&major.to_le_bytes());
    buf[28..30].copy_from_slice(&0xFFFE_u16.to_le_bytes());
    buf[30..32].copy_from_slice(&shift.to_le_bytes());
    buf[32..34].copy_from_slice(&6_u16.to_le_bytes());
    put32(&mut buf, 44, 1); // FAT sector count
    put32(&mut buf, 48, 0); // first directory sector
    put32(&mut buf, 56, 4096); // mini stream cutoff
    put32(&mut buf, 60, ENDOFCHAIN); // no MiniFAT
    put32(&mut buf, 68, ENDOFCHAIN); // no DIFAT sector
    put32(&mut buf, 76, 1); // DIFAT[0] = the FAT sector
    for i in 1..109 {
        put32(&mut buf, 76 + i * 4, FREESECT);
    }

    let dir = ss; // sector 0
    let dir_entry = |buf: &mut [u8], index: usize, name: &str, kind: u8, start: u32, size: u64| {
        let off = dir + index * 128;
        for (i, unit) in name.encode_utf16().enumerate() {
            buf[off + i * 2..off + i * 2 + 2].copy_from_slice(&unit.to_le_bytes());
        }
        let name_len = ((name.encode_utf16().count() + 1) * 2) as u16;
        buf[off + 64..off + 66].copy_from_slice(&name_len.to_le_bytes());
        buf[off + 66] = kind;
        buf[off + 67] = 1;
        for field in [68, 72, 76] {
            buf[off + field..off + field + 4].copy_from_slice(&FREESECT.to_le_bytes());
        }
        buf[off + 116..off + 120].copy_from_slice(&start.to_le_bytes());
        buf[off + 120..off + 128].copy_from_slice(&size.to_le_bytes());
    };
    dir_entry(&mut buf, 0, "Root Entry", 5, ENDOFCHAIN, 0);
    put32(&mut buf, dir + 76, 1);
    for (i, s) in streams.iter().enumerate() {
        let size = s.declared.unwrap_or(s.data.len() as u64);
        dir_entry(&mut buf, i + 1, s.name, 2, chains[i][0], size);
        if i + 1 < streams.len() {
            put32(&mut buf, dir + (i + 1) * 128 + 72, (i + 2) as u32);
        }
    }

    let fat = 2 * ss; // sector 1
    for i in 0..ss / 4 {
        put32(&mut buf, fat + i * 4, FREESECT);
    }
    put32(&mut buf, fat, if self_loop_dir { 0 } else { ENDOFCHAIN });
    put32(&mut buf, fat + 4, FATSECT);
    for (s, chain) in streams.iter().zip(&chains) {
        for (j, &sid) in chain.iter().enumerate() {
            let link = chain.get(j + 1).copied().unwrap_or(ENDOFCHAIN);
            put32(&mut buf, fat + sid as usize * 4, link);
            let from = j * ss;
            let to = (from + ss).min(s.data.len());
            if from < to {
                let at = ss * (1 + sid as usize);
                buf[at..at + (to - from)].copy_from_slice(&s.data[from..to]);
            }
        }
    }
    buf
}

/// An Agile `EncryptionInfo` stream whose `encryptedKey` carries the given
/// `spinCount` and `keyBits`; the cipher text is placeholder base64.
fn agile_info(spin_count: &str, key_bits: &str) -> Vec<u8> {
    let xml = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><encryption xmlns="http://schemas.microsoft.com/office/2006/encryption" xmlns:p="http://schemas.microsoft.com/office/2006/keyEncryptor/password"><keyData saltSize="16" blockSize="16" keyBits="256" hashSize="64" cipherAlgorithm="AES" cipherChaining="ChainingModeCBC" hashAlgorithm="SHA512" saltValue="AAAAAAAAAAAAAAAAAAAAAA=="/><dataIntegrity encryptedHmacKey="{b64_64}" encryptedHmacValue="{b64_64}"/><keyEncryptors><keyEncryptor uri="http://schemas.microsoft.com/office/2006/keyEncryptor/password"><p:encryptedKey spinCount="{spin_count}" saltSize="16" blockSize="16" keyBits="{key_bits}" hashSize="64" cipherAlgorithm="AES" cipherChaining="ChainingModeCBC" hashAlgorithm="SHA512" saltValue="AAAAAAAAAAAAAAAAAAAAAA==" encryptedVerifierHashInput="AAAAAAAAAAAAAAAAAAAAAA==" encryptedVerifierHashValue="{b64_64}" encryptedKeyValue="AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA="/></keyEncryptor></keyEncryptors></encryption>"#,
        b64_64 = "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA==",
    );
    let mut out = vec![4, 0, 4, 0, 0x40, 0, 0, 0];
    out.extend_from_slice(xml.as_bytes());
    out
}

/// A Standard (4.2) `EncryptionInfo` stream whose `headerSize` field is hostile.
fn standard_info(header_size: u32) -> Vec<u8> {
    let mut out = vec![0_u8; 8 + 4 + 200];
    out[..4].copy_from_slice(&[4, 0, 2, 0]);
    out[4..8].copy_from_slice(&0x24_u32.to_le_bytes());
    out[8..12].copy_from_slice(&header_size.to_le_bytes());
    out
}

fn enc_streams<'a>(info: Vec<u8>, declared: Option<u64>) -> [Stream<'a>; 2] {
    [
        Stream {
            name: "EncryptionInfo",
            data: info,
            declared,
        },
        Stream {
            name: "EncryptedPackage",
            data: vec![0; 32],
            declared: None,
        },
    ]
}

// ---------------------------------------------------------------------------
// Cases
// ---------------------------------------------------------------------------

fn zip_name_cases() {
    let limits = ZipSecurityLimits::default();
    let names: [&[u8]; 8] = [
        b"../x",
        b"/x",
        b"\\x",
        b"a\\..\\b",
        b"C:/x",
        b"C:x",
        b"a\0b",
        b"a/../../b",
    ];
    for name in names {
        let label = format!("zip name {name:?}");
        let zip = stored_zip(&[(b"ok.txt", b"hi"), (name, b"x")]);
        timed(&label, || {
            expect_code(&label, read_zip_entries(&zip, &limits), "SECURITY");
        });
    }

    let dup = stored_zip(&[(b"a.txt", b"one"), (b"a.txt", b"two")]);
    expect_code(
        "duplicate name",
        read_zip_entries(&dup, &limits),
        "SECURITY",
    );

    // A colon past the first two bytes is a legal part name.
    let legal = stored_zip(&[(b"xl/a:b.xml", b"<x/>")]);
    let entries = read_zip_entries(&legal, &limits).expect("colon inside a name is legal");
    assert_eq!(entries["xl/a:b.xml"], b"<x/>");

    // A missing workbook part is MISSING_PART; a missing [Content_Types].xml still reads.
    let parts = base_parts();
    let without = |skip: &str| -> Vec<u8> {
        let kept: Vec<ZipEntry> = parts.iter().filter(|p| p.name != skip).cloned().collect();
        write_zip(&kept).expect("rezip")
    };
    expect_code(
        "no xl/workbook.xml",
        read_xlsx(&without("xl/workbook.xml")),
        "MISSING_PART",
    );
    read_xlsx(&without("[Content_Types].xml")).expect("missing [Content_Types].xml reads");
}

fn zip_bomb_case() {
    let mut parts = base_parts();
    parts.push(ZipEntry {
        name: "xl/media/zeros.bin".into(),
        data: vec![0; 64 << 20],
    });
    let bomb = write_zip(&parts).expect("zip bomb");
    drop(parts);
    let limits = ZipSecurityLimits {
        max_decompressed_size: 1 << 20,
        max_compression_ratio: f64::INFINITY,
        ..ZipSecurityLimits::default()
    };
    let (msg, peak) = measured("64 MiB zero bomb", || {
        expect_code(
            "64 MiB zero bomb",
            read_zip_entries(&bomb, &limits),
            "SECURITY",
        )
    });
    assert_peak("64 MiB zero bomb", peak, 8 * MIB);
    assert!(
        msg.contains("budget") || msg.contains("exceed"),
        "message: {msg}"
    );
}

fn forged_central_directory_size_case() {
    let mut zip = plain_xlsx();
    // The first central-directory record: signature, then the uncompressed size at +24.
    let cd = zip
        .windows(4)
        .position(|w| w == [0x50, 0x4B, 0x01, 0x02])
        .expect("central directory");
    for forged in [0x8000_0000_u32, 0xFFFF_FFFE] {
        zip[cd + 24..cd + 28].copy_from_slice(&forged.to_le_bytes());
        let label = format!("forged central-directory size {forged:#x}");
        let (result, peak) = measured(&label, || read_xlsx(&zip));
        assert_peak(&label, peak, 64 * MIB);
        // The size is a hint only: the entry inflates to its real length.
        result.unwrap_or_else(|e| panic!("forged size {forged:#x} must still read: {e}"));
    }
}

fn ole2_cases() {
    // Directory size 0xFFFFFFFF (v3) and u64::MAX (v4), read without a password.
    for (label, major, declared) in [
        ("dir size 0xFFFFFFFF v3", 3, 0xFFFF_FFFF_u64),
        ("dir size u64::MAX v4", 4, u64::MAX),
    ] {
        let file = build_ole2(
            major,
            &enc_streams(agile_info("1", "256"), Some(declared)),
            false,
        );
        let ((), peak) = measured(label, || {
            expect_code(label, read_xlsx_json(&file), "SECURITY");
        });
        assert_peak(label, peak, 64 * MIB);
    }

    // Standard headerSize 0xFFFFFFFC.
    let file = build_ole2(3, &enc_streams(standard_info(0xFFFF_FFFC), None), false);
    let ((), peak) = measured("headerSize 0xFFFFFFFC", || {
        expect_code(
            "headerSize 0xFFFFFFFC",
            read_xlsx_json(&file),
            "PASSWORD_PROTECTED",
        );
    });
    assert_peak("headerSize 0xFFFFFFFC", peak, 64 * MIB);

    // A directory chain that loops onto itself.
    let file = build_ole2(3, &enc_streams(agile_info("1", "256"), None), true);
    let ((), peak) = measured("self-looping directory chain", || {
        expect_code(
            "self-looping directory chain",
            read_xlsx_json(&file),
            "SECURITY",
        );
    });
    assert_peak("self-looping directory chain", peak, 64 * MIB);

    // spinCount u32::MAX with a password.
    let file = build_ole2(
        3,
        &enc_streams(agile_info("4294967295", "256"), None),
        false,
    );
    let ((), peak) = measured("spinCount u32::MAX", || {
        expect_code(
            "spinCount u32::MAX",
            read_xlsx_json_with_password(&file, "pw"),
            "SECURITY",
        );
    });
    assert_peak("spinCount u32::MAX", peak, 64 * MIB);

    // keyBits 0xFFFFFFF8 with a password.
    let file = build_ole2(3, &enc_streams(agile_info("1", "4294967288"), None), false);
    let ((), peak) = measured("keyBits 0xFFFFFFF8", || {
        expect_code(
            "keyBits 0xFFFFFFF8",
            read_xlsx_json_with_password(&file, "pw"),
            "PASSWORD_PROTECTED",
        );
    });
    assert_peak("keyBits 0xFFFFFFF8", peak, 64 * MIB);
}

fn hmac_tamper_case() {
    let zip = write_xlsx(&text_workbook(400)).expect("fixture");
    let encrypted = encrypt_file(&zip, "pw").expect("encrypt");
    let info = read_stream(&encrypted, "EncryptionInfo").expect("EncryptionInfo");
    let package = read_stream(&encrypted, "EncryptedPackage").expect("EncryptedPackage");
    assert!(
        package.len() >= 4096,
        "fixture package must leave the mini stream"
    );

    let rebuild = |package: &[u8]| {
        write_ole2(&[("EncryptionInfo", &info), ("EncryptedPackage", package)]).expect("ole2")
    };
    let intact = rebuild(&package);
    read_xlsx_json_with_password(&intact, "pw").expect("untampered control reads");

    let mut tampered = package.clone();
    tampered[8 + 100] ^= 0x01;
    let file = rebuild(&tampered);
    let msg = timed("HMAC tamper", || {
        expect_code(
            "HMAC tamper",
            read_xlsx_json_with_password(&file, "pw"),
            "PASSWORD_PROTECTED",
        )
    });
    assert!(msg.contains("HMAC"), "message: {msg}");
}

fn doctype_case() {
    const DTD: &str = r#"<!DOCTYPE x [<!ENTITY a "AAAA"><!ENTITY b "&a;&a;&a;&a;"><!ENTITY e SYSTEM "file:///etc/passwd">]>"#;
    let decl = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#;
    let sheet = format!(
        r#"{decl}{DTD}<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>&a;&b;&e;</t></is></c></row><row r="2"><c r="A2" t="s"><v>0</v></c></row></sheetData></worksheet>"#
    );
    let sst = format!(
        r#"{decl}{DTD}<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1"><si><t>&a;&b;&e;</t></si></sst>"#
    );
    let parts: Vec<ZipEntry> = base_parts()
        .into_iter()
        .map(|mut p| {
            match p.name.as_str() {
                "xl/worksheets/sheet1.xml" => p.data = sheet.clone().into_bytes(),
                "xl/sharedStrings.xml" => p.data = sst.clone().into_bytes(),
                _ => {}
            }
            p
        })
        .collect();
    assert!(parts.iter().any(|p| p.name == "xl/sharedStrings.xml"));
    let zip = write_zip(&parts).expect("zip");

    let json = timed("DOCTYPE read_xlsx_json", || read_xlsx_json(&zip)).expect("DOCTYPE reads");
    assert!(!json.contains("AAAA"), "an entity expanded: {json}");
    assert!(
        !json.contains("root:"),
        "an external entity resolved: {json}"
    );
    timed("DOCTYPE read_xlsx", || read_xlsx(&zip)).expect("DOCTYPE reads");

    // Unknown entities are dropped by the XML reader, so each cell is empty.
    let value: serde_json::Value = serde_json::from_str(&json).expect("json");
    let cells = &value["sheets"][0]["worksheet"]["rows"];
    for row in 0..2 {
        let cell = &cells[row]["cells"][0];
        assert_eq!(
            cell["reference"],
            format!("A{}", row + 1),
            "row {row}: {cell}"
        );
        assert_eq!(cell["value"], "", "row {row}: {cell}");
    }
}

/// A sheet .rels that names one 200-comment part 10,000 times: the part is
/// parsed once and the sheet holds one copy of each comment. Read once per
/// relationship, it was 2,000,000 comments, a 408 MB peak (release) and 15 s
/// (debug). The struct read stays under 16 MiB; the JSON read under 32 MiB,
/// most of it the 1.1 MB `.rels` kept in `preservedEntries` as a JSON number
/// array.
fn many_relationships_to_one_part_case() {
    const REPEATS: usize = 10_000;
    const COMMENTS: usize = 200;
    const REL_COMMENTS: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments";

    let mut comments = String::from(
        r#"<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><authors><author>Ann</author></authors><commentList>"#,
    );
    for n in 1..=COMMENTS {
        let _ = write!(
            comments,
            r#"<comment ref="B{n}" authorId="0"><text><t>hostile-input comment {n}</t></text></comment>"#
        );
    }
    comments.push_str("</commentList></comments>");
    let mut rels = String::from(
        r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">"#,
    );
    for n in 1..=REPEATS {
        let _ = write!(
            rels,
            r#"<Relationship Id="rId{n}" Type="{REL_COMMENTS}" Target="../comments1.xml"/>"#
        );
    }
    rels.push_str("</Relationships>");

    let mut parts = base_parts();
    parts.push(ZipEntry {
        name: "xl/comments1.xml".into(),
        data: comments.into_bytes(),
    });
    parts.push(ZipEntry {
        name: "xl/worksheets/_rels/sheet1.xml.rels".into(),
        data: rels.into_bytes(),
    });
    let zip = write_zip(&parts).expect("zip");
    drop(parts);

    let label = "10,000 relationships to one comments part";
    let (wb, peak) = measured(label, || read_xlsx(&zip));
    assert_peak(label, peak, 16 * MIB);
    let wb = wb.expect("the read succeeds");
    assert_eq!(wb.sheets[0].worksheet.comments.len(), COMMENTS, "{label}");

    let label = "10,000 relationships to one comments part, JSON";
    let (json, peak) = measured(label, || read_xlsx_json(&zip));
    assert_peak(label, peak, 32 * MIB);
    let json: serde_json::Value =
        serde_json::from_str(&json.expect("the read succeeds")).expect("json");
    assert_eq!(
        json["sheets"][0]["worksheet"]["comments"]
            .as_array()
            .map(Vec::len),
        Some(COMMENTS),
        "{label}"
    );
}

/// The relationship types the sharing cases name.
const REL_COMMENTS: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments";
const REL_DRAWING: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing";

/// A `.rels` part with one relationship of `rel_type` per target, in order.
fn rels_naming(rel_type: &str, targets: impl IntoIterator<Item = String>) -> Vec<u8> {
    let mut rels = String::from(
        r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">"#,
    );
    for (n, target) in targets.into_iter().enumerate() {
        let _ = write!(
            rels,
            r#"<Relationship Id="rId{}" Type="{rel_type}" Target="{target}"/>"#,
            n + 1
        );
    }
    rels.push_str("</Relationships>");
    rels.into_bytes()
}

/// The size of the ZIP entry `name` in `parts`.
fn part_len(parts: &[ZipEntry], name: &str) -> usize {
    parts
        .iter()
        .find(|p| p.name == name)
        .unwrap_or_else(|| panic!("no {name}"))
        .data
        .len()
}

/// 10,000 `<sheet>` rows of `xl/workbook.xml` resolving to one worksheet part
/// of about 1 MB. Each row was a parse of that part (struct and JSON reads
/// alike), 10,000 x the part; the read now fails with `SECURITY` while the
/// rows are resolved, before any worksheet is parsed, under 16 MiB on both
/// read paths.
fn many_sheet_rows_on_one_worksheet_part_case() {
    const ROWS: usize = 10_000;

    let mut parts = write_zip_parts(&text_workbook(22_000));
    let sheet_len = part_len(&parts, "xl/worksheets/sheet1.xml");
    assert!(sheet_len > 900_000, "worksheet part is {sheet_len} bytes");
    let workbook = parts
        .iter_mut()
        .find(|p| p.name == "xl/workbook.xml")
        .expect("xl/workbook.xml");
    let xml = String::from_utf8(std::mem::take(&mut workbook.data)).expect("utf-8");
    let (head, rest) = xml.split_once("<sheets>").expect("<sheets>");
    let (_, tail) = rest.split_once("</sheets>").expect("</sheets>");
    let mut rows = String::new();
    for n in 1..=ROWS {
        let _ = write!(rows, r#"<sheet name="S{n}" sheetId="{n}" r:id="rId1"/>"#);
    }
    workbook.data = format!("{head}<sheets>{rows}</sheets>{tail}").into_bytes();
    let zip = write_zip(&parts).expect("zip");
    drop(parts);

    let label = "10,000 sheet rows on one worksheet part";
    let (result, peak) = measured(label, || read_xlsx(&zip).map(drop));
    assert_peak(label, peak, 16 * MIB);
    let msg = expect_code(label, result, "SECURITY");
    assert!(
        msg.contains(r#""xl/worksheets/sheet1.xml""#),
        "message: {msg}"
    );

    let label = "10,000 sheet rows on one worksheet part, JSON";
    let (result, peak) = measured(label, || read_xlsx_json(&zip).map(drop));
    assert_peak(label, peak, 16 * MIB);
    expect_code(label, result, "SECURITY");
}

/// The entries of `write_xlsx(wb)`, sorted by name.
fn write_zip_parts(wb: &WorkbookData) -> Vec<ZipEntry> {
    let zip = write_xlsx(wb).expect("write fixture");
    let mut parts: Vec<ZipEntry> = read_zip_entries(&zip, &ZipSecurityLimits::default())
        .expect("read")
        .into_iter()
        .map(|(name, data)| ZipEntry { name, data })
        .collect();
    parts.sort_by(|a, b| a.name.cmp(&b.name));
    parts
}

/// 200 sheets, each on a worksheet part with a `.rels` of its own, all naming
/// one comments part of about 1 MB. Each sheet got a copy of every comment,
/// 200 x the part (a 236 MB peak); now the first sheet holds the comments and
/// the other 199 relationships are skipped, under 16 MiB on both read paths.
fn many_sheets_sharing_one_comments_part_case() {
    const SHEETS: usize = 200;
    const COMMENTS: usize = 11_000;

    let wb = WorkbookData {
        sheets: (1..=SHEETS)
            .map(|n| SheetData {
                name: format!("S{n}"),
                ..Default::default()
            })
            .collect(),
        ..Default::default()
    };
    let mut parts = write_zip_parts(&wb);
    let mut comments = String::from(
        r#"<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><authors><author>Ann</author></authors><commentList>"#,
    );
    for n in 1..=COMMENTS {
        let _ = write!(
            comments,
            r#"<comment ref="B{n}" authorId="0"><text><t>hostile-input comment {n}</t></text></comment>"#
        );
    }
    comments.push_str("</commentList></comments>");
    assert!(
        comments.len() > 900_000,
        "comments part is {} bytes",
        comments.len()
    );
    parts.push(ZipEntry {
        name: "xl/comments1.xml".into(),
        data: comments.into_bytes(),
    });
    for n in 1..=SHEETS {
        parts.push(ZipEntry {
            name: format!("xl/worksheets/_rels/sheet{n}.xml.rels"),
            data: rels_naming(REL_COMMENTS, ["../comments1.xml".to_owned()]),
        });
    }
    let zip = write_zip(&parts).expect("zip");
    drop(parts);

    let label = "200 sheets sharing one comments part";
    let (wb, peak) = measured(label, || read_xlsx(&zip));
    assert_peak(label, peak, 16 * MIB);
    let wb = wb.expect("the read succeeds");
    assert_eq!(wb.sheets[0].worksheet.comments.len(), COMMENTS, "{label}");
    assert!(
        wb.sheets[1..]
            .iter()
            .all(|s| s.worksheet.comments.is_empty()),
        "{label}"
    );
    drop(wb);

    let label = "200 sheets sharing one comments part, JSON";
    let (json, peak) = measured(label, || read_xlsx_json(&zip));
    assert_peak(label, peak, 16 * MIB);
    let json: serde_json::Value =
        serde_json::from_str(&json.expect("the read succeeds")).expect("json");
    let counts: Vec<usize> = json["sheets"]
        .as_array()
        .expect("sheets")
        .iter()
        .map(|s| s["worksheet"]["comments"].as_array().map_or(0, Vec::len))
        .collect();
    assert_eq!(counts[0], COMMENTS, "{label}");
    assert!(counts[1..].iter().all(|&c| c == 0), "{label}");
}

/// One sheet with 200 drawings, each a part with a `.rels` of its own, all
/// naming one chart part of about 1 MB (8,000 series). The chart was parsed
/// once per drawing and the sheet held 200 charts (18 s in a debug build); now
/// it is parsed once, the first drawing keeps it and the other 199 anchors are
/// skipped, under 16 MiB on both read paths.
fn many_drawings_sharing_one_chart_part_case() {
    const DRAWINGS: usize = 200;
    const SERIES: usize = 8_000;

    let series: Vec<serde_json::Value> = (0..SERIES)
        .map(|n| {
            serde_json::json!({
                "idx": n,
                "order": n,
                "valRef": format!("S!$B${}:$B${}", n + 1, n + 100),
            })
        })
        .collect();
    let mut wb = text_workbook(1);
    wb.sheets[0].worksheet.charts = serde_json::from_value(serde_json::json!([{
        "chart": { "chartType": "line", "series": series },
        "anchor": { "fromCol": 0, "fromRow": 0, "toCol": 5, "toRow": 10 },
    }]))
    .expect("chart fixture");
    let mut parts = write_zip_parts(&wb);
    let chart_len = part_len(&parts, "xl/charts/chart1.xml");
    assert!(chart_len > 900_000, "chart part is {chart_len} bytes");
    let copy = |name: &str| {
        parts
            .iter()
            .find(|p| p.name == name)
            .unwrap_or_else(|| panic!("no {name}"))
            .data
            .clone()
    };
    let drawing = copy("xl/drawings/drawing1.xml");
    let drawing_rels = copy("xl/drawings/_rels/drawing1.xml.rels");
    parts.retain(|p| p.name != "xl/worksheets/_rels/sheet1.xml.rels");
    for n in 2..=DRAWINGS {
        parts.push(ZipEntry {
            name: format!("xl/drawings/drawing{n}.xml"),
            data: drawing.clone(),
        });
        parts.push(ZipEntry {
            name: format!("xl/drawings/_rels/drawing{n}.xml.rels"),
            data: drawing_rels.clone(),
        });
    }
    parts.push(ZipEntry {
        name: "xl/worksheets/_rels/sheet1.xml.rels".into(),
        data: rels_naming(
            REL_DRAWING,
            (1..=DRAWINGS).map(|n| format!("../drawings/drawing{n}.xml")),
        ),
    });
    let zip = write_zip(&parts).expect("zip");
    drop(parts);

    let label = "200 drawings sharing one chart part";
    let (wb, peak) = measured(label, || read_xlsx(&zip));
    assert_peak(label, peak, 16 * MIB);
    let wb = wb.expect("the read succeeds");
    assert_eq!(wb.sheets[0].worksheet.charts.len(), 1, "{label}");
    assert_eq!(
        wb.sheets[0].worksheet.charts[0].chart.series.len(),
        SERIES,
        "{label}"
    );
    drop(wb);

    let label = "200 drawings sharing one chart part, JSON";
    let (json, peak) = measured(label, || read_xlsx_json(&zip));
    assert_peak(label, peak, 16 * MIB);
    let json: serde_json::Value =
        serde_json::from_str(&json.expect("the read succeeds")).expect("json");
    assert_eq!(
        json["sheets"][0]["worksheet"]["charts"]
            .as_array()
            .map(Vec::len),
        Some(1),
        "{label}"
    );
}

const REL_VML_DRAWING: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing";

/// The size of the VML drawing the two VML cases name.
const BIG_VML_LEN: usize = 2 * MIB;

/// A VML drawing of about [`BIG_VML_LEN`] bytes: a root declaring the VML,
/// Office and Excel prefixes, `head`, then empty shapes whose titles are
/// pseudo-random (xorshift64), so the part deflates to about a tenth of its
/// size, under the default `ZipSecurityLimits` ratio of 100. The root is
/// closed when `closed`. Only `head` can hold `ClientData`.
fn big_vml(head: &str, closed: bool) -> Vec<u8> {
    let mut vml = String::with_capacity(BIG_VML_LEN + 256);
    vml.push_str(r#"<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel">"#);
    vml.push_str(head);
    let mut x: u64 = 0x9E37_79B9_7F4A_7C15;
    let mut n = 2_048_u32;
    while vml.len() < BIG_VML_LEN {
        x ^= x << 13;
        x ^= x >> 7;
        x ^= x << 17;
        n += 1;
        let _ = write!(
            vml,
            r#"<v:shape id="_x0000_s{n}" o:title="{x:016x}" style="position:absolute;visibility:hidden"/>"#
        );
    }
    if closed {
        vml.push_str("</xml>");
    }
    vml.into_bytes()
}

/// One sheet whose `.rels` names one 2 MiB VML drawing without notes or form
/// controls (no `ClientData`) 2,000 times. The writer searched the drawing
/// for `ClientData` once per relationship, 2,000 x the part (the write took
/// 40 s in a debug build); it now classifies each preserved drawing once per
/// write. The same package with one relationship is the control. Either way
/// the drawing is written once, as it is, and named by `<legacyDrawingHF>`.
fn many_relationships_to_one_vml_drawing_case() {
    const PATH: &str = "xl/drawings/vmlDrawing1.vml";
    let vml = big_vml("", true);
    for repeats in [1, 2_000] {
        let mut parts = base_parts();
        parts.push(ZipEntry {
            name: PATH.into(),
            data: vml.clone(),
        });
        parts.push(ZipEntry {
            name: "xl/worksheets/_rels/sheet1.xml.rels".into(),
            data: rels_naming(
                REL_VML_DRAWING,
                std::iter::repeat_n("../drawings/vmlDrawing1.vml".to_owned(), repeats),
            ),
        });
        let zip = write_zip(&parts).expect("zip");
        drop(parts);

        let label = format!("{repeats} vmlDrawing relationships to one 2 MiB drawing");
        let wb = timed(&format!("{label}, read"), || read_xlsx(&zip)).expect("the read succeeds");
        let written =
            timed(&format!("{label}, write"), || write_xlsx(&wb)).expect("the write succeeds");
        let out: std::collections::HashMap<String, Vec<u8>> =
            read_zip_entries(&written, &ZipSecurityLimits::default())
                .expect("read")
                .into_iter()
                .collect();
        assert!(out[PATH] == vml, "{label}: the drawing changed");
        let sheet = String::from_utf8_lossy(&out["xl/worksheets/sheet1.xml"]);
        assert!(
            sheet.contains(r#"<legacyDrawingHF r:id="rId1"/>"#)
                && !sheet.contains("<legacyDrawing "),
            "{label}: {sheet}"
        );
    }
}

/// Time bound of the write of 300 sheets sharing one 2 MiB note drawing: 3 s
/// in a debug build, where it took 0.45 s (446 ms to 451 ms over three runs)
/// on the test host, an Intel Core Ultra 9 290HX Plus laptop, on 2026-11-13,
/// so a slower or busy runner does not fail it; [`LIMIT`] in an optimized
/// build (20 ms on the same host and day). A write that parses the drawing
/// once per sheet (26 s in a debug build on 2026-11-12) fails it.
const SHARED_NOTE_DRAWING_WRITE_LIMIT: Duration = if cfg!(debug_assertions) {
    Duration::from_secs(3)
} else {
    LIMIT
};

/// 300 sheets with one note each on B2, whose `.rels` all name one 2 MiB note
/// drawing that is not well-formed (its root is never closed), so the writer
/// keeps it as it is. The first sheet with notes examines it, every sheet
/// names it as it is, and the write runs within
/// [`SHARED_NOTE_DRAWING_WRITE_LIMIT`]; a merge that parsed the part for each
/// sheet (300 x the part) took 26 s in a debug build on 2026-11-12.
fn many_sheets_sharing_one_note_drawing_case() {
    use modern_xlsx_core::ooxml::comments::Comment;
    use modern_xlsx_core::ooxml::worksheet::WorksheetXml;
    const SHEETS: usize = 300;
    const PATH: &str = "xl/drawings/notes.vml";

    let wb = WorkbookData {
        sheets: (1..=SHEETS)
            .map(|n| SheetData {
                name: format!("S{n}"),
                worksheet: WorksheetXml {
                    comments: vec![Comment {
                        cell_ref: "B2".into(),
                        author: "Ann".into(),
                        text: format!("note {n}"),
                    }],
                    ..Default::default()
                },
                ..Default::default()
            })
            .collect(),
        ..Default::default()
    };
    // The generated note drawings (sheet n's is vmlDrawing{n}.vml) go, and
    // every sheet .rels names the big one instead.
    let mut parts = write_zip_parts(&wb);
    parts.retain(|p| !p.name.starts_with("xl/drawings/vmlDrawing"));
    for n in 1..=SHEETS {
        let name = format!("xl/worksheets/_rels/sheet{n}.xml.rels");
        let part = parts
            .iter_mut()
            .find(|p| p.name == name)
            .unwrap_or_else(|| panic!("no {name}"));
        let rels = String::from_utf8(std::mem::take(&mut part.data)).expect("utf-8");
        let old = format!(r#"Target="../drawings/vmlDrawing{n}.vml""#);
        assert!(rels.contains(&old), "{name}: {rels}");
        part.data = rels
            .replace(&old, r#"Target="../drawings/notes.vml""#)
            .into_bytes();
    }
    let vml = big_vml(
        r#"<v:shape id="_x0000_s1025"><x:ClientData ObjectType="Note"><x:Row>1</x:Row><x:Column>1</x:Column></x:ClientData></v:shape>"#,
        false,
    );
    parts.push(ZipEntry {
        name: PATH.into(),
        data: vml.clone(),
    });
    let zip = write_zip(&parts).expect("zip");
    drop(parts);

    let label = "300 sheets sharing one 2 MiB note drawing";
    let wb = timed(&format!("{label}, read"), || read_xlsx(&zip)).expect("the read succeeds");
    assert!(
        wb.sheets.iter().all(|s| s.worksheet.comments.len() == 1),
        "{label}"
    );
    let written = timed_within(
        &format!("{label}, write"),
        SHARED_NOTE_DRAWING_WRITE_LIMIT,
        || write_xlsx(&wb),
    )
    .expect("the write succeeds");
    let out: std::collections::HashMap<String, Vec<u8>> =
        read_zip_entries(&written, &ZipSecurityLimits::default())
            .expect("read")
            .into_iter()
            .collect();
    assert!(out[PATH] == vml, "{label}: the drawing changed");
    for n in 1..=SHEETS {
        let rels = String::from_utf8_lossy(&out[&format!("xl/worksheets/_rels/sheet{n}.xml.rels")]);
        let sheet = String::from_utf8_lossy(&out[&format!("xl/worksheets/sheet{n}.xml")]);
        let at = sheet
            .find(r#"<legacyDrawing r:id=""#)
            .unwrap_or_else(|| panic!("{label}: sheet{n}: {sheet}"))
            + r#"<legacyDrawing r:id=""#.len();
        let id = &sheet[at..at + sheet[at..].find('"').expect("id")];
        assert!(
            rels.contains(&format!(
                r#"Id="{id}" Type="{REL_VML_DRAWING}" Target="../drawings/notes.vml""#
            )),
            "{label}: sheet{n} names {id}: {rels}"
        );
    }
}

const REL_PIVOT_TABLE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable";

/// 2,000 pivot table parts on one sheet, each with a row field without
/// items, all on one cache of 100,000 shared items: a package of about
/// 1.7 MB that reads at once. The write completed each table's row field
/// from the cache, 200 million items (37 s, 500 MB, a 5.8 GB peak heap, past
/// wasm32's 4 GiB); now it fails with `SIZE_LIMIT` before any part is
/// written, under 1 MiB, naming the ceiling.
fn many_pivot_tables_on_one_large_cache_case() {
    use modern_xlsx_core::ooxml::pivot_table::{PivotCacheDefinitionData, PivotTableData};
    use modern_xlsx_core::ooxml::worksheet::WorksheetXml;
    const TABLES: usize = 2_000;
    const SHARED: usize = 100_000;
    const NS: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

    let mut items = String::new();
    for n in 0..SHARED {
        let _ = write!(items, r#"<s v="v{n}"/>"#);
    }
    let cache = PivotCacheDefinitionData::parse(
        format!(
            r#"<pivotCacheDefinition xmlns="{NS}"><cacheSource type="worksheet"><worksheetSource ref="A1:B2" sheet="S1"/></cacheSource><cacheFields count="2"><cacheField name="F"><sharedItems count="{SHARED}">{items}</sharedItems></cacheField><cacheField name="V"/></cacheFields></pivotCacheDefinition>"#
        )
        .as_bytes(),
    )
    .expect("cache");
    drop(items);
    let table = |row_items: &str| {
        format!(
            r#"<pivotTableDefinition xmlns="{NS}" name="P" cacheId="0"><location ref="D1:E4"/><pivotFields count="2"><pivotField axis="axisRow"{row_items}<pivotField dataField="1"/></pivotFields><rowFields count="1"><field x="0"/></rowFields><dataFields count="1"><dataField fld="1" subtotal="sum"/></dataFields></pivotTableDefinition>"#
        )
    };
    // Write one table with items (a small write), then give the package
    // 2,000 item-less table parts.
    let wb = WorkbookData {
        sheets: vec![SheetData {
            name: "S1".into(),
            worksheet: WorksheetXml {
                pivot_tables: vec![
                    PivotTableData::parse(
                        table(r#"><items count="1"><item x="0"/></items></pivotField>"#).as_bytes(),
                    )
                    .expect("table"),
                ],
                ..Default::default()
            },
            ..Default::default()
        }],
        pivot_caches: vec![cache],
        ..Default::default()
    };
    let mut parts = write_zip_parts(&wb);
    drop(wb);
    let bare = table("/>");
    parts.retain(|p| !p.name.starts_with("xl/pivotTables/"));
    for n in 1..=TABLES {
        parts.push(ZipEntry {
            name: format!("xl/pivotTables/pivotTable{n}.xml"),
            data: bare.clone().into_bytes(),
        });
    }
    let sheet_rels = parts
        .iter_mut()
        .find(|p| p.name == "xl/worksheets/_rels/sheet1.xml.rels")
        .expect("sheet rels");
    sheet_rels.data = rels_naming(
        REL_PIVOT_TABLE,
        (1..=TABLES).map(|n| format!("../pivotTables/pivotTable{n}.xml")),
    );
    let zip = write_zip(&parts).expect("zip");
    drop(parts);

    let wb = read_xlsx(&zip).expect("the read succeeds");
    assert_eq!(wb.sheets[0].worksheet.pivot_tables.len(), TABLES);
    assert_eq!(wb.pivot_caches[0].fields[0].shared_items.len(), SHARED);

    let label = "2,000 pivot tables on one cache of 100,000 shared items";
    let (result, peak) = measured(label, || write_xlsx(&wb).map(drop));
    assert_peak(label, peak, MIB);
    let msg = expect_code(label, result, "SIZE_LIMIT");
    assert!(
        msg.contains("4194304") && msg.contains("explicit items"),
        "{label}: {msg}"
    );
}

const REL_THREADED_COMMENT: &str =
    "http://schemas.microsoft.com/office/2017/10/relationships/threadedComment";

/// The notes the generated-note cases put on one sheet, on A1 down to
/// A1000000.
const NOTES: usize = 1_000_000;

/// `base_parts()` with `part` at `name`, named by sheet 1's `.rels` under
/// `rel_type`.
fn sheet_part_zip(name: &str, rel_type: &str, part: String) -> Vec<u8> {
    let mut parts = base_parts();
    parts.push(ZipEntry {
        name: name.into(),
        data: part.into_bytes(),
    });
    parts.push(ZipEntry {
        name: "xl/worksheets/_rels/sheet1.xml.rels".into(),
        data: rels_naming(rel_type, [format!("../{}", &name[3..])]),
    });
    write_zip(&parts).expect("zip")
}

/// Reads `zip`, then asserts its write fails with the `SIZE_LIMIT` of the
/// generated-note ceiling within [`LIMIT`] and under 16 MiB of peak heap
/// growth.
fn assert_note_ceiling_write(label: &str, zip: &[u8], check: impl Fn(&WorkbookData)) {
    let wb = read_xlsx(zip).expect("the read succeeds");
    check(&wb);
    let (result, peak) = measured(label, || write_xlsx(&wb).map(drop));
    assert_peak(label, peak, 16 * MIB);
    let msg = expect_code(label, result, "SIZE_LIMIT");
    assert!(msg.contains("more than 131072"), "{label}: {msg}");
}

/// 1,000,000 thread roots on distinct cells and no legacy comment, a
/// package of 5.5 MB. The writer gave each root a `tc=` placeholder note and
/// a VML note shape, about 1 KB of output each (in a release build 524,288
/// roots wrote in 2.4 s with a 973 MiB peak heap; 2,000,000 took 10.6 s and
/// 3.66 GiB, and trapped the WASM build). The write now fails with
/// `SIZE_LIMIT` at the root that passes the ceiling of 131,072 generated
/// notes, before any placeholder is built: what it holds then is the set of
/// the cells counted so far.
fn many_thread_roots_without_notes_case() {
    let mut part = String::from(
        r#"<ThreadedComments xmlns="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments">"#,
    );
    for n in 1..=NOTES {
        let _ = write!(
            part,
            r#"<threadedComment ref="A{n}" id="{{{n:08X}-0000-0000-0000-000000000000}}"/>"#
        );
    }
    part.push_str("</ThreadedComments>");
    let zip = sheet_part_zip(
        "xl/threadedComments/threadedComment1.xml",
        REL_THREADED_COMMENT,
        part,
    );
    assert_note_ceiling_write("1,000,000 bare thread roots", &zip, |wb| {
        assert_eq!(wb.sheets[0].worksheet.threaded_comments.len(), NOTES);
        assert!(wb.sheets[0].worksheet.comments.is_empty());
    });
}

/// 1,000,000 one-character legacy notes on distinct cells and no VML
/// drawing, a package of 2.6 MB. The writer generated a note shape for
/// each, about 640 bytes of VML (in a release build 524,288 notes wrote in
/// 0.87 s with a 535 MiB peak heap; 2,000,000 took 3.5 s and 2.0 GiB). The
/// write now fails with `SIZE_LIMIT` at the cell that passes the ceiling,
/// before the drawing is built: what it holds then is the list and set of
/// the cells counted so far.
fn many_legacy_notes_without_a_drawing_case() {
    let mut part = String::from(
        r#"<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><authors><author>A</author></authors><commentList>"#,
    );
    for n in 1..=NOTES {
        let _ = write!(
            part,
            r#"<comment ref="A{n}" authorId="0"><text><t>x</t></text></comment>"#
        );
    }
    part.push_str("</commentList></comments>");
    let zip = sheet_part_zip("xl/comments1.xml", REL_COMMENTS, part);
    assert_note_ceiling_write("1,000,000 legacy notes", &zip, |wb| {
        assert_eq!(wb.sheets[0].worksheet.comments.len(), NOTES);
    });
}

const REL_TABLE: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/table";

/// Time bound of the write of 3,000 tables on one sheet, six times what it
/// took on the test host on 2026-11-13 (0.47 s to 0.51 s in a debug build,
/// 30 ms to 32 ms in a release build), so a busy host does not fail it (one
/// debug run under another build took 1.1 s), while name probing that grows
/// with the square of the table count (4.4 s in a debug build on
/// 2026-11-13) still fails it.
const MANY_TABLES_WRITE_LIMIT: Duration = if cfg!(debug_assertions) {
    Duration::from_secs(3)
} else {
    Duration::from_millis(200)
};

/// One sheet whose `.rels` names 3,000 table parts of about 170 bytes each,
/// well inside the default ZIP limits. The read keeps all 3,000 tables. The
/// write named each one by probing `table1.xml`, `table2.xml`, ... from 1,
/// about 4.5 million probes in all (4.4 s in a debug build; the cost grows
/// with the square of the count); it now resumes where the last search for
/// that prefix ended, and the write runs within [`MANY_TABLES_WRITE_LIMIT`].
fn many_tables_on_one_sheet_case() {
    const TABLES: usize = 3_000;
    const NS: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

    let mut parts = base_parts();
    for n in 1..=TABLES {
        parts.push(ZipEntry {
            name: format!("xl/tables/table{n}.xml"),
            data: format!(
                r#"<table xmlns="{NS}" id="{n}" name="T{n}" displayName="T{n}" ref="A1:A2"><tableColumns count="1"><tableColumn id="1" name="C"/></tableColumns></table>"#
            )
            .into_bytes(),
        });
    }
    parts.push(ZipEntry {
        name: "xl/worksheets/_rels/sheet1.xml.rels".into(),
        data: rels_naming(
            REL_TABLE,
            (1..=TABLES).map(|n| format!("../tables/table{n}.xml")),
        ),
    });
    let zip = write_zip(&parts).expect("zip");
    drop(parts);

    let label = "3,000 tables on one sheet";
    let wb = timed(&format!("{label}, read"), || read_xlsx(&zip)).expect("the read succeeds");
    assert_eq!(wb.sheets[0].worksheet.tables.len(), TABLES, "{label}");
    let written = timed_within(&format!("{label}, write"), MANY_TABLES_WRITE_LIMIT, || {
        write_xlsx(&wb)
    })
    .expect("the write succeeds");
    let names: std::collections::HashSet<String> =
        read_zip_entries(&written, &ZipSecurityLimits::default())
            .expect("read")
            .into_keys()
            .collect();
    assert!(
        (1..=TABLES).all(|n| names.contains(&format!("xl/tables/table{n}.xml"))),
        "{label}"
    );
}

#[test]
fn hostile_inputs_end_in_exact_codes_within_bounds() {
    zip_name_cases();
    zip_bomb_case();
    forged_central_directory_size_case();
    ole2_cases();
    hmac_tamper_case();
    doctype_case();
    many_relationships_to_one_part_case();
    many_sheet_rows_on_one_worksheet_part_case();
    many_sheets_sharing_one_comments_part_case();
    many_drawings_sharing_one_chart_part_case();
    many_relationships_to_one_vml_drawing_case();
    many_sheets_sharing_one_note_drawing_case();
    many_pivot_tables_on_one_large_cache_case();
    many_thread_roots_without_notes_case();
    many_legacy_notes_without_a_drawing_case();
    many_tables_on_one_sheet_case();
}
