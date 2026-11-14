//! Reads a ZIP archive into a map of entry name to decompressed bytes.
//!
//! [`read_zip_entries`] applies [`ZipSecurityLimits`]: a cap on the total
//! decompressed size, on each entry's compression ratio and on the entry
//! count, and rejects unsafe or duplicate entry names. The same limits carry
//! the maximum JSON length the JSON readers enforce. Untrusted limits go
//! through [`ZipSecurityLimits::clamped`], which holds the size, entry-count
//! and JSON-length limits to their ceilings and resets a NaN or non-positive
//! ratio to the default.

use core::hint::cold_path;
use std::collections::HashMap;
use std::collections::hash_map::Entry;
use std::io::{Cursor, Read};

use zip::ZipArchive;

use crate::errors::{ModernXlsxError, Result};

/// Upper bound on the buffer reserved up front for one entry (64 MiB).
///
/// The reservation is only an optimisation: the buffer still grows as the
/// entry inflates, within the per-entry cap enforced by [`read_zip_entries`].
const MAX_CAPACITY_HINT: u64 = 64 << 20;

/// Security and size limits for reading a workbook.
///
/// Deserializes from the JSON object that `ReadOptions.limits` serializes to
/// (keys `maxDecompressedSize`, `maxCompressionRatio`, `maxEntries`,
/// `maxJsonLength`). Omitted keys keep their default; unknown keys are rejected.
/// Call [`ZipSecurityLimits::clamped`] on untrusted values before use.
#[derive(Debug, Clone, Copy, PartialEq, serde::Deserialize)]
#[serde(default, rename_all = "camelCase", deny_unknown_fields)]
pub struct ZipSecurityLimits {
    /// Maximum total decompressed size across all entries (default: 2 GiB).
    pub max_decompressed_size: u64,
    /// Maximum allowed ratio of decompressed size to compressed size per entry (default: 100.0).
    pub max_compression_ratio: f64,
    /// Maximum number of entries in the archive (default: 100,000).
    #[serde(deserialize_with = "saturating_usize")]
    pub max_entries: usize,
    /// Maximum length of the JSON that `read_xlsx_json*` produces (default and
    /// ceiling: 536,870,888, V8's maximum string length).
    ///
    /// The length is counted in UTF-8 bytes, which is an upper bound on UTF-16
    /// code units, so a workbook with non-ASCII text can be rejected slightly
    /// before V8's own limit would be reached.
    ///
    /// The same limit bounds, separately, the text a read copies from where it
    /// is stored once (shared strings, comment authors, hyperlink targets,
    /// namespace declarations), counted in bytes, and the memory of the
    /// records a read keeps in its lists (the defined names first, then the
    /// shared string entries and rich text runs, the records of
    /// `xl/styles.xml`, the chart series and axes, the threaded comments and
    /// mentions, then the worksheet list elements), counted at the size of
    /// each record on the target plus its strings, except the plain text of a
    /// shared string entry, which the part's size bounds (on wasm32 from half
    /// to about four fifths of the 64-bit size: a `DxfStyle` is 1,200 bytes on
    /// x86_64 and 784 on wasm32, a `ChartAxis` 216 and 176, measured
    /// 2026-11-13), so a read can end in
    /// `SizeLimit` although its JSON would fit. With the `parallel` Cargo
    /// feature the records counted before the sheets and each sheet's list
    /// memory are held to the limit together, so the sheets of one read can
    /// together keep up to their number times it.
    #[serde(deserialize_with = "saturating_usize")]
    pub max_json_length: usize,
}

/// Deserializes a JSON integer into a `usize`, saturating at `usize::MAX` so a value
/// above the 32-bit `usize` of wasm32 still reaches [`ZipSecurityLimits::clamped`]
/// instead of failing to parse. Negative numbers and non-integers stay errors.
fn saturating_usize<'de, D>(deserializer: D) -> std::result::Result<usize, D::Error>
where
    D: serde::Deserializer<'de>,
{
    let value = <u64 as serde::Deserialize>::deserialize(deserializer)?;
    Ok(usize::try_from(value).unwrap_or(usize::MAX))
}

impl ZipSecurityLimits {
    /// Ceiling for `max_decompressed_size` (2 GiB): wasm32 has 4 GiB of memory.
    pub const MAX_DECOMPRESSED_SIZE_CEILING: u64 = 2 * 1024 * 1024 * 1024;
    /// Ceiling for `max_entries`.
    pub const MAX_ENTRIES_CEILING: usize = 1_048_576;
    /// Ceiling for `max_json_length`: V8's maximum string length.
    pub const MAX_JSON_LENGTH_CEILING: usize = 536_870_888;

    /// Returns the limits with each value held to its ceiling.
    ///
    /// A NaN or non-positive `max_compression_ratio` becomes the default (100.0);
    /// the ratio has no ceiling.
    #[must_use]
    pub fn clamped(self) -> Self {
        let defaults = Self::default();
        Self {
            max_decompressed_size: self
                .max_decompressed_size
                .min(Self::MAX_DECOMPRESSED_SIZE_CEILING),
            max_compression_ratio: if self.max_compression_ratio > 0.0 {
                self.max_compression_ratio
            } else {
                defaults.max_compression_ratio
            },
            max_entries: self.max_entries.min(Self::MAX_ENTRIES_CEILING),
            max_json_length: self.max_json_length.min(Self::MAX_JSON_LENGTH_CEILING),
        }
    }
}

impl Default for ZipSecurityLimits {
    fn default() -> Self {
        Self {
            max_decompressed_size: Self::MAX_DECOMPRESSED_SIZE_CEILING,
            max_compression_ratio: 100.0,
            max_entries: 100_000,
            max_json_length: Self::MAX_JSON_LENGTH_CEILING,
        }
    }
}

/// Returns `true` when an entry name could escape the archive root or smuggle
/// a path past a consumer: a NUL byte, a leading `/` or `\`, a `..` component
/// (split on both separators) or a drive prefix such as `C:` at the start.
///
/// A `:` elsewhere is legal in OPC part names (`xl/a:b.xml`), and the reader never
/// touches a filesystem, so only a leading drive prefix is rejected.
fn is_unsafe_entry_name(name: &str) -> bool {
    name.contains('\0')
        || name.starts_with(['/', '\\'])
        || matches!(name.as_bytes(), [drive, b':', ..] if drive.is_ascii_alphabetic())
        || name.split(['/', '\\']).any(|component| component == "..")
}

/// Reads a ZIP archive from raw bytes, applying security checks, and returns a map
/// of entry name to decompressed bytes.
///
/// # Security
///
/// - Rejects entry names containing a NUL byte.
/// - Rejects entry names that start with a drive prefix (`C:`).
/// - Rejects entry names with a `..` component, split on `/` and `\`.
/// - Rejects entry names with a leading `/` or `\`.
/// - Rejects archives with duplicate entry names, both byte-identical names in
///   the central directory and names that decode to the same string.
/// - Rejects archives whose declared entry data ranges overlap each other or the
///   central directory.
/// - Rejects archives with more than `limits.max_entries` entries.
/// - Bounds every entry's inflate to `min(remaining total budget, compressed
///   size * limits.max_compression_ratio)` plus one byte, so a decompression
///   bomb is cut off while reading and never inflated in full; an entry that
///   exceeds the cap is rejected for its ratio or for the total budget.
/// - Rejects archives whose total decompressed size exceeds `limits.max_decompressed_size`.
///
/// Directory entries are skipped; empty (zero-byte) entries are kept.
pub fn read_zip_entries(
    data: &[u8],
    limits: &ZipSecurityLimits,
) -> Result<HashMap<String, Vec<u8>>> {
    let cursor = Cursor::new(data);
    let mut archive = ZipArchive::new(cursor)
        .map_err(|e| ModernXlsxError::ZipRead(format!("Failed to open ZIP archive: {e}")))?;

    // The zip crate collapses central-directory records with identical raw names
    // before iteration, so duplicates must be detected on the archive itself.
    if archive.has_duplicate_names() {
        cold_path();
        return Err(ModernXlsxError::Security(
            "duplicate entry names in ZIP central directory".into(),
        ));
    }
    // The per-entry ratio cap derives from the compressed size an entry declares, which a
    // hostile archive can forge to raise the cap; forged ranges overlap other data.
    if archive
        .has_overlapping_files()
        .map_err(|e| ModernXlsxError::ZipRead(format!("Failed to read ZIP entry headers: {e}")))?
    {
        cold_path();
        return Err(ModernXlsxError::Security(
            "ZIP entry data overlaps another entry or the central directory (forged compressed size)"
                .into(),
        ));
    }
    if archive.len() > limits.max_entries {
        cold_path();
        return Err(ModernXlsxError::Security(format!(
            "ZIP entry count {} exceeds limit {}",
            archive.len(),
            limits.max_entries
        )));
    }

    let mut entries = HashMap::with_capacity(archive.len());
    let mut total_decompressed: u64 = 0;

    for i in 0..archive.len() {
        let mut file = archive.by_index(i).map_err(|e| {
            ModernXlsxError::ZipEntry(format!("Failed to read ZIP entry at index {i}: {e}"))
        })?;

        // Skip directories
        if file.is_dir() {
            continue;
        }

        // `ZipFile::name()` is fallible: the raw entry name may fail to decode.
        let name = file
            .name()
            .map_err(|e| {
                ModernXlsxError::ZipEntry(format!("Invalid ZIP entry name at index {i}: {e}"))
            })?
            .into_owned();

        if is_unsafe_entry_name(&name) {
            cold_path();
            return Err(ModernXlsxError::Security(format!(
                "path traversal detected in ZIP entry: {name}"
            )));
        }

        let compressed_size = file.compressed_size();
        let declared_size = file.size();

        if compressed_size == 0 && declared_size > 0 {
            // compressed_size is 0 but there is data: suspicious
            cold_path();
            return Err(ModernXlsxError::Security(format!(
                "zero compressed size with non-zero data for entry: {name}"
            )));
        }

        // Per-entry cap: the entry may inflate to at most what is left of the total
        // budget and at most `compressed size * max ratio` bytes. A float-to-int
        // `as` cast saturates, so an infinite ratio gives `u64::MAX`.
        let remaining = limits
            .max_decompressed_size
            .saturating_sub(total_decompressed);
        let ratio_cap = if compressed_size > 0 {
            (compressed_size as f64 * limits.max_compression_ratio).ceil() as u64
        } else {
            0
        };
        let entry_cap = remaining.min(ratio_cap);

        // The declared size comes straight from the (attacker-controllable) ZIP
        // central directory and can be up to u64::MAX (e.g. a forged ZIP64 extra
        // field). The read below is capped at `entry_cap + 1` bytes and the
        // reservation is at most 64 MiB, so a forged declared size neither traps
        // wasm32 in `with_capacity` nor forces a large reserve; `read_to_end` grows
        // the buffer as the entry really inflates.
        let capacity_hint = declared_size.min(entry_cap).min(MAX_CAPACITY_HINT) as usize;
        let mut buf = Vec::with_capacity(capacity_hint);
        (&mut file)
            .take(entry_cap.saturating_add(1))
            .read_to_end(&mut buf)
            .map_err(|e| {
                ModernXlsxError::ZipRead(format!("Failed to decompress ZIP entry '{name}': {e}"))
            })?;

        let decompressed_size = buf.len() as u64;
        if decompressed_size > entry_cap {
            cold_path();
            return Err(ModernXlsxError::Security(
                if decompressed_size > remaining {
                    format!(
                        "total decompressed size exceeds limit {} at entry: {name}",
                        limits.max_decompressed_size
                    )
                } else if compressed_size == 0 {
                    format!("zero compressed size with non-zero data for entry: {name}")
                } else {
                    format!(
                        "compression ratio exceeds limit {} for entry: {name}",
                        limits.max_compression_ratio
                    )
                },
            ));
        }

        total_decompressed = total_decompressed.saturating_add(decompressed_size);

        // A name that decodes to the same string as an earlier entry (for example
        // UTF-8 and CP437 raw names) is a duplicate too.
        match entries.entry(name) {
            Entry::Occupied(slot) => {
                cold_path();
                return Err(ModernXlsxError::Security(format!(
                    "duplicate entry name: {}",
                    slot.key()
                )));
            }
            Entry::Vacant(slot) => {
                slot.insert(buf);
            }
        }
    }

    Ok(entries)
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::io::Write;
    use zip::write::SimpleFileOptions;
    use zip::{CompressionMethod, ZipWriter};

    /// Helper: build a ZIP archive from a list of (name, data) pairs.
    fn build_zip(files: &[(&str, &[u8])]) -> Vec<u8> {
        let buf = Vec::new();
        let cursor = Cursor::new(buf);
        let mut zw = ZipWriter::new(cursor);
        let opts = SimpleFileOptions::default().compression_method(CompressionMethod::Deflated);
        for (name, data) in files {
            zw.start_file(*name, opts).unwrap();
            zw.write_all(data).unwrap();
        }
        zw.finish().unwrap().into_inner()
    }

    #[test]
    fn test_read_zip_entries() {
        let zip_bytes = build_zip(&[
            ("hello.txt", b"Hello, world!"),
            ("nested/data.xml", b"<root>data</root>"),
        ]);

        let limits = ZipSecurityLimits::default();
        let result =
            read_zip_entries(&zip_bytes, &limits).expect("read_zip_entries should succeed");

        assert_eq!(result.len(), 2);
        assert_eq!(result.get("hello.txt").unwrap(), b"Hello, world!");
        assert_eq!(result.get("nested/data.xml").unwrap(), b"<root>data</root>");
    }

    #[test]
    fn test_rejects_path_traversal() {
        let zip_bytes = build_zip(&[("../evil.txt", b"malicious content")]);

        let limits = ZipSecurityLimits::default();
        let result = read_zip_entries(&zip_bytes, &limits);

        assert!(result.is_err());
        let err = result.unwrap_err();
        assert!(
            matches!(err, ModernXlsxError::Security(ref msg) if msg.contains("path traversal")),
            "expected Security error with path traversal message, got: {err}"
        );
    }

    #[test]
    fn test_rejects_oversized_entries() {
        let zip_bytes = build_zip(&[("big.txt", &vec![b'A'; 1024])]);

        let limits = ZipSecurityLimits {
            max_decompressed_size: 512, // only 512 bytes allowed
            max_compression_ratio: 100.0,
            ..Default::default()
        };

        let result = read_zip_entries(&zip_bytes, &limits);

        assert!(result.is_err());
        let err = result.unwrap_err();
        assert!(
            matches!(err, ModernXlsxError::Security(ref msg) if msg.contains("total decompressed size")),
            "expected Security error about total decompressed size, got: {err}"
        );
    }

    #[test]
    fn test_error_messages_have_context() {
        // Bad ZIP archive should mention "Failed to open"
        let err = read_zip_entries(b"not a zip", &ZipSecurityLimits::default()).unwrap_err();
        assert!(
            err.to_string().contains("Failed to open ZIP archive"),
            "expected 'Failed to open ZIP archive' context, got: {err}"
        );
        assert_eq!(err.code(), "ZIP_READ");

        // Path traversal should include the entry name
        let zip_bytes = build_zip(&[("../evil.txt", b"bad")]);
        let err = read_zip_entries(&zip_bytes, &ZipSecurityLimits::default()).unwrap_err();
        let msg = err.to_string();
        assert!(
            msg.contains("../evil.txt"),
            "expected entry name in error, got: {msg}"
        );
    }

    /// Regression (audit): a forged central-directory uncompressed-size must not
    /// drive an unbounded `Vec::with_capacity` (which would trap the wasm module
    /// or abort on overflow) before the size guards run. We patch a valid ZIP's
    /// central-directory uncompressed-size field to ~4 GiB; the reader must still
    /// complete quickly without panicking/aborting because the capacity hint is
    /// clamped to the decompression budget while the actual (tiny) data is read.
    #[test]
    fn forged_uncompressed_size_does_not_trigger_huge_allocation() {
        let mut zip_bytes = build_zip(&[("a.txt", b"hi")]);

        // Locate the central-directory file header (signature PK\x01\x02) and
        // overwrite its uncompressed-size field (offset +24, u32 LE) with a huge
        // sentinel. Leaving CRC/data intact keeps the entry otherwise readable.
        let sig = [0x50u8, 0x4B, 0x01, 0x02];
        let cd = zip_bytes
            .windows(4)
            .position(|w| w == sig)
            .expect("central directory header present");
        zip_bytes[cd + 24..cd + 28].copy_from_slice(&0xFFFF_FFFEu32.to_le_bytes());

        // The read is capped by the per-entry budget, so the forged size only
        // influences a clamped capacity hint and the real content is still read.
        let entries = read_zip_entries(&zip_bytes, &ZipSecurityLimits::default())
            .expect("forged declared size must not stop a small, honest entry");
        assert_eq!(entries["a.txt"], b"hi");
    }

    /// One hand-built entry: raw name bytes, general-purpose flags, stored data.
    struct Raw<'a> {
        name: &'a [u8],
        flags: u16,
        data: &'a [u8],
    }

    fn crc32(data: &[u8]) -> u32 {
        let mut c = 0xFFFF_FFFFu32;
        for &b in data {
            c ^= u32::from(b);
            for _ in 0..8 {
                c = if c & 1 != 0 {
                    (c >> 1) ^ 0xEDB8_8320
                } else {
                    c >> 1
                };
            }
        }
        !c
    }

    /// Builds a STORED archive by hand so duplicate and odd raw names are possible
    /// (`ZipWriter` refuses duplicates and validates names).
    fn raw_zip(entries: &[Raw<'_>]) -> Vec<u8> {
        let mut out = Vec::new();
        let mut cd = Vec::new();
        for e in entries {
            let off = out.len() as u32;
            let crc = crc32(e.data);
            let sz = e.data.len() as u32;
            let name_len = (e.name.len() as u16).to_le_bytes();
            out.extend_from_slice(&0x0403_4b50u32.to_le_bytes());
            out.extend_from_slice(&20u16.to_le_bytes());
            out.extend_from_slice(&e.flags.to_le_bytes());
            out.extend_from_slice(&0u16.to_le_bytes()); // stored
            out.extend_from_slice(&0u32.to_le_bytes()); // time/date
            out.extend_from_slice(&crc.to_le_bytes());
            out.extend_from_slice(&sz.to_le_bytes());
            out.extend_from_slice(&sz.to_le_bytes());
            out.extend_from_slice(&name_len);
            out.extend_from_slice(&0u16.to_le_bytes()); // extra
            out.extend_from_slice(e.name);
            out.extend_from_slice(e.data);

            cd.extend_from_slice(&0x0201_4b50u32.to_le_bytes());
            cd.extend_from_slice(&20u16.to_le_bytes()); // made by
            cd.extend_from_slice(&20u16.to_le_bytes()); // needed
            cd.extend_from_slice(&e.flags.to_le_bytes());
            cd.extend_from_slice(&0u16.to_le_bytes()); // stored
            cd.extend_from_slice(&0u32.to_le_bytes()); // time/date
            cd.extend_from_slice(&crc.to_le_bytes());
            cd.extend_from_slice(&sz.to_le_bytes());
            cd.extend_from_slice(&sz.to_le_bytes());
            cd.extend_from_slice(&name_len);
            cd.extend_from_slice(&[0; 2 + 2 + 2 + 2 + 4]); // extra, comment, disk, int, ext attrs
            cd.extend_from_slice(&off.to_le_bytes());
            cd.extend_from_slice(e.name);
        }
        let cd_off = out.len() as u32;
        let count = (entries.len() as u16).to_le_bytes();
        out.extend_from_slice(&cd);
        out.extend_from_slice(&0x0605_4b50u32.to_le_bytes());
        out.extend_from_slice(&[0; 4]); // disk numbers
        out.extend_from_slice(&count);
        out.extend_from_slice(&count);
        out.extend_from_slice(&(cd.len() as u32).to_le_bytes());
        out.extend_from_slice(&cd_off.to_le_bytes());
        out.extend_from_slice(&0u16.to_le_bytes()); // comment length
        out
    }

    fn security_message(result: Result<HashMap<String, Vec<u8>>>) -> String {
        match result {
            Err(ModernXlsxError::Security(m)) => m,
            other => panic!("expected Err(Security), got {other:?}"),
        }
    }

    #[test]
    fn zero_bomb_is_cut_off_by_the_total_budget() {
        let zip_bytes = build_zip(&[("zeros.bin", &vec![0u8; 64 << 20])]);
        let limits = ZipSecurityLimits {
            max_decompressed_size: 1 << 20,
            max_compression_ratio: f64::INFINITY,
            ..Default::default()
        };
        let msg = security_message(read_zip_entries(&zip_bytes, &limits));
        assert!(msg.contains("total decompressed size"), "got: {msg}");
    }

    #[test]
    fn zero_bomb_is_cut_off_by_the_default_ratio() {
        let zip_bytes = build_zip(&[("zeros.bin", &vec![0u8; 64 << 20])]);
        let msg = security_message(read_zip_entries(&zip_bytes, &ZipSecurityLimits::default()));
        assert!(msg.contains("compression ratio"), "got: {msg}");
    }

    /// The 64 MiB zero bomb with its declared compressed size forged to `0x7FFF_FFFF` in the
    /// central directory (offset + 20) and, when `patch_local` is set, in the local header
    /// (offset 18) too.
    fn forged_zero_bomb(patch_local: bool) -> Vec<u8> {
        let mut zip_bytes = build_zip(&[("zeros.bin", &vec![0u8; 64 << 20])]);
        let cd = zip_bytes
            .windows(4)
            .rposition(|w| w == [b'P', b'K', 1, 2])
            .expect("central directory record");
        let forged = 0x7FFF_FFFFu32.to_le_bytes();
        zip_bytes[cd + 20..cd + 24].copy_from_slice(&forged);
        if patch_local {
            zip_bytes[18..22].copy_from_slice(&forged);
        }
        zip_bytes
    }

    #[test]
    fn rejects_forged_compressed_size_in_both_headers() {
        let msg = security_message(read_zip_entries(
            &forged_zero_bomb(true),
            &ZipSecurityLimits::default(),
        ));
        assert!(msg.contains("overlaps"), "got: {msg}");
    }

    #[test]
    fn rejects_forged_compressed_size_in_the_central_directory_only() {
        let msg = security_message(read_zip_entries(
            &forged_zero_bomb(false),
            &ZipSecurityLimits::default(),
        ));
        assert!(msg.contains("overlaps"), "got: {msg}");
    }

    #[test]
    fn zero_bomb_is_cut_off_before_full_inflate() {
        let mut zip_bytes = build_zip(&[("zeros.bin", &vec![0u8; 64 << 20])]);
        // Corrupt the CRC in the local header (offset 14) and the central-directory record
        // (offset 16). A reader that inflated the whole entry would reach EOF and fail the
        // checksum with a ZipRead error; a bounded read stops at the cap first.
        let cd = zip_bytes
            .windows(4)
            .position(|w| w == b"PK")
            .expect("central directory record");
        for i in (14..18).chain(cd + 16..cd + 20) {
            zip_bytes[i] ^= 0xFF;
        }
        let msg = security_message(read_zip_entries(&zip_bytes, &ZipSecurityLimits::default()));
        assert!(msg.contains("compression ratio"), "got: {msg}");
    }

    #[test]
    fn rejects_duplicate_central_directory_names() {
        let mut zip_bytes = build_zip(&[("a.xml", b"<a/>"), ("b.xml", b"<b/>")]);
        // Rename b.xml to a.xml in both its local header and its central-directory record.
        let mut patched = 0;
        for i in 0..zip_bytes.len() - 5 {
            if &zip_bytes[i..i + 5] == b"b.xml" {
                zip_bytes[i..i + 5].copy_from_slice(b"a.xml");
                patched += 1;
            }
        }
        assert_eq!(
            patched, 2,
            "expected one local and one central-directory name"
        );
        let msg = security_message(read_zip_entries(&zip_bytes, &ZipSecurityLimits::default()));
        assert!(msg.contains("duplicate"), "got: {msg}");
    }

    #[test]
    fn rejects_names_that_decode_to_the_same_string() {
        // 0xC3 0xA9 with the UTF-8 flag and the lone CP437 byte 0x82 both decode to "é".
        let zip_bytes = raw_zip(&[
            Raw {
                name: &[0xC3, 0xA9],
                flags: 0x0800,
                data: b"UTF8",
            },
            Raw {
                name: &[0x82],
                flags: 0,
                data: b"CP437",
            },
        ]);
        let msg = security_message(read_zip_entries(&zip_bytes, &ZipSecurityLimits::default()));
        assert!(msg.contains("duplicate"), "got: {msg}");
    }

    #[test]
    fn rejects_too_many_entries() {
        let zip_bytes = build_zip(&[("a", b"1"), ("b", b"2"), ("c", b"3")]);
        let limits = ZipSecurityLimits {
            max_entries: 2,
            ..Default::default()
        };
        let msg = security_message(read_zip_entries(&zip_bytes, &limits));
        assert!(msg.contains("entry count"), "got: {msg}");
        let limits = ZipSecurityLimits {
            max_entries: 3,
            ..Default::default()
        };
        assert_eq!(read_zip_entries(&zip_bytes, &limits).unwrap().len(), 3);
    }

    #[test]
    fn rejects_unsafe_entry_names() {
        let names: [&[u8]; 7] = [
            b"C:/x.xml",
            b"C:x",
            b"a\0b.xml",
            b"\\x.xml",
            b"/x.xml",
            b"a\\..\\b",
            b"../x",
        ];
        for raw_name in names {
            let zip_bytes = raw_zip(&[Raw {
                name: raw_name,
                flags: 0,
                data: b"x",
            }]);
            let msg = security_message(read_zip_entries(&zip_bytes, &ZipSecurityLimits::default()));
            let name = String::from_utf8_lossy(raw_name);
            assert!(msg.contains("path traversal"), "{name:?}: {msg}");
            assert!(msg.contains(name.as_ref()), "{name:?}: {msg}");
        }
    }

    #[test]
    fn accepts_colons_that_are_not_drive_prefixes() {
        for name in ["xl/a:b.xml", "xl/workbook.xml"] {
            let zip_bytes = raw_zip(&[Raw {
                name: name.as_bytes(),
                flags: 0,
                data: b"x",
            }]);
            let entries = read_zip_entries(&zip_bytes, &ZipSecurityLimits::default())
                .unwrap_or_else(|e| panic!("{name}: {e}"));
            assert_eq!(entries[name], b"x");
        }
    }

    #[test]
    fn keeps_empty_entries_and_skips_directories() {
        let mut zw = ZipWriter::new(Cursor::new(Vec::new()));
        let opts = SimpleFileOptions::default().compression_method(CompressionMethod::Deflated);
        zw.add_directory("customXml/", opts).unwrap();
        zw.start_file("customXml/item1.xml", opts).unwrap();
        zw.start_file("a.xml", opts).unwrap();
        zw.write_all(b"<a/>").unwrap();
        let zip_bytes = zw.finish().unwrap().into_inner();

        let entries = read_zip_entries(&zip_bytes, &ZipSecurityLimits::default()).unwrap();
        assert_eq!(entries.len(), 2);
        assert_eq!(entries["customXml/item1.xml"], Vec::<u8>::new());
        assert!(!entries.contains_key("customXml/"));
    }

    #[test]
    fn test_limits_deserialize_partial_uses_defaults() {
        let limits: ZipSecurityLimits = serde_json::from_str(r#"{"maxEntries":5}"#).unwrap();
        assert_eq!(limits.max_entries, 5);
        assert_eq!(
            ZipSecurityLimits {
                max_entries: 100_000,
                ..limits
            },
            ZipSecurityLimits::default()
        );
        let all: ZipSecurityLimits = serde_json::from_str(
            r#"{"maxDecompressedSize":1,"maxCompressionRatio":2.5,"maxEntries":3,"maxJsonLength":4}"#,
        )
        .unwrap();
        assert_eq!(all.max_decompressed_size, 1);
        assert_eq!(all.max_compression_ratio, 2.5);
        assert_eq!(all.max_entries, 3);
        assert_eq!(all.max_json_length, 4);
    }

    #[test]
    fn test_limits_deserialize_saturates_huge_counts() {
        let limits: ZipSecurityLimits = serde_json::from_str(
            r#"{"maxEntries":18446744073709551615,"maxJsonLength":9007199254740991}"#,
        )
        .unwrap();
        assert_eq!(limits.clamped().max_entries, 1_048_576);
        assert_eq!(limits.clamped().max_json_length, 536_870_888);
        assert!(serde_json::from_str::<ZipSecurityLimits>(r#"{"maxEntries":-1}"#).is_err());
    }

    #[test]
    fn test_limits_deserialize_rejects_unknown_key() {
        assert!(serde_json::from_str::<ZipSecurityLimits>(r#"{"bogus":1}"#).is_err());
    }

    #[test]
    fn test_limits_clamped_applies_ceilings() {
        let c = ZipSecurityLimits {
            max_decompressed_size: 1 << 40,
            max_compression_ratio: 1e9,
            max_entries: usize::MAX,
            max_json_length: usize::MAX,
        }
        .clamped();
        assert_eq!(c.max_decompressed_size, 2 << 30);
        assert_eq!(c.max_compression_ratio, 1e9);
        assert_eq!(c.max_entries, 1_048_576);
        assert_eq!(c.max_json_length, 536_870_888);
        assert_eq!(
            ZipSecurityLimits::default().clamped(),
            ZipSecurityLimits::default()
        );
    }

    #[test]
    fn test_limits_clamped_resets_bad_ratio() {
        for bad in [f64::NAN, 0.0, -1.0, f64::NEG_INFINITY] {
            let c = ZipSecurityLimits {
                max_compression_ratio: bad,
                ..Default::default()
            }
            .clamped();
            assert_eq!(c.max_compression_ratio, 100.0);
        }
    }
}
