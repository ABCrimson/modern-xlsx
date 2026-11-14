//! MS-CFB compound-file reader, built in every build.
//!
//! The crate-internal `Cfb` parses a file once (the header with its DIFAT, the
//! FAT and the directory) and reads any number of its streams through the FAT
//! or the mini stream. [`read_stream`] is the one-shot form: one parse, one
//! stream. `ole2::detect` classifies a file by the directory of a `Cfb`, and
//! the read path hands that same `Cfb` on to the `EncryptionInfo` and
//! `EncryptedPackage` reads.
//!
//! The input is untrusted. The reader validates the header fields it depends
//! on (MS-CFB §2.2), caps the FAT sector list at the sectors the data can hold,
//! sizes its buffers from the bytes present rather than from header counts, and
//! saturates sector offsets on 32-bit targets. A FAT sector listed twice, a
//! sector chain that loops and a declared stream size above the bytes its chain
//! holds are `Security` errors.

use core::hint::cold_path;
use std::collections::HashSet;

use super::{ENDOFCHAIN, FREESECT, MAXREGSECT, MINI_SECTOR_SIZE};
use crate::Result;
use crate::errors::ModernXlsxError;

/// Directory entry object types (MS-CFB §2.6.1).
const STREAM_OBJECT: u8 = 2;
const ROOT_STORAGE_OBJECT: u8 = 5;

/// Streams shorter than 4096 bytes live in the mini stream (MS-CFB §2.2).
const MINI_STREAM_CUTOFF: u64 = 4096;

/// The OLE2 header fields the reader needs, with the FAT sector list collected
/// from the header DIFAT and the DIFAT chain.
struct Ole2Header {
    sector_size: usize,
    first_dir_sector: u32,
    first_minifat_sector: u32,
    fat_sectors: Vec<u32>,
}

impl Ole2Header {
    /// Validates the fixed header fields the reader depends on (MS-CFB §2.2):
    /// the byte order mark, the major version with its sector shift, the mini
    /// sector shift (6), and the mini stream cutoff (4096). Any other value is
    /// an `UnrecognizedFormat` "Invalid OLE2 header" error.
    fn parse(data: &[u8]) -> Result<Self> {
        if data.len() < 512 {
            cold_path();
            return Err(ModernXlsxError::UnrecognizedFormat(format!(
                "File too small for OLE2 header: got {} bytes, need at least 512",
                data.len()
            )));
        }
        let u16_at = |at: usize| u16::from_le_bytes([data[at], data[at + 1]]);
        let u32_at =
            |at: usize| u32::from_le_bytes([data[at], data[at + 1], data[at + 2], data[at + 3]]);

        if data[28..30] != [0xFE, 0xFF] {
            cold_path();
            return Err(invalid_header(format!(
                "byte order mark is {:02X?}, expected [FE, FF]",
                &data[28..30]
            )));
        }
        let (major_version, sector_shift) = (u16_at(26), u16_at(30));
        if !matches!((major_version, sector_shift), (3, 9) | (4, 12)) {
            cold_path();
            return Err(invalid_header(format!(
                "major version {major_version} with sector shift {sector_shift} \
                 (expected 3 with 9, or 4 with 12)"
            )));
        }
        let mini_sector_shift = u16_at(32);
        if mini_sector_shift != 6 {
            cold_path();
            return Err(invalid_header(format!(
                "mini sector shift {mini_sector_shift} (expected 6)"
            )));
        }
        let mini_cutoff = u32_at(56);
        if u64::from(mini_cutoff) != MINI_STREAM_CUTOFF {
            cold_path();
            return Err(invalid_header(format!(
                "mini stream cutoff {mini_cutoff} (expected 4096)"
            )));
        }

        let mut header = Self {
            sector_size: 1 << sector_shift,
            first_dir_sector: u32_at(48),
            first_minifat_sector: u32_at(60),
            fat_sectors: Vec::new(),
        };
        let header_difat = data[76..512]
            .as_chunks::<4>()
            .0
            .iter()
            .map(|c| u32::from_le_bytes(*c));
        header.fat_sectors = header.collect_fat_sectors(
            data,
            u32_at(44) as usize,
            header_difat,
            u32_at(68),
            u32_at(72) as usize,
        )?;
        Ok(header)
    }

    /// Collects the FAT sector IDs: the 109 header DIFAT entries, then the DIFAT
    /// chain from `first_difat`, where each sector holds `sector_size / 4 - 1` FAT
    /// sector IDs and the next DIFAT sector ID last. Free slots (any marker above
    /// `MAXREGSECT`) are skipped, and a FAT sector listed twice is an error. The
    /// walk stops at an ENDOFCHAIN/FREESECT next pointer, a revisited DIFAT
    /// sector, `num_difat` sectors walked, `num_fat_sectors` IDs collected, or a
    /// sector past the data. Header counts are attacker-controlled, so the list
    /// is capped at the number of sectors the data can hold.
    fn collect_fat_sectors(
        &self,
        data: &[u8],
        num_fat_sectors: usize,
        header_difat: impl Iterator<Item = u32>,
        first_difat: u32,
        num_difat: usize,
    ) -> Result<Vec<u32>> {
        let limit = num_fat_sectors.min(data.len() / self.sector_size);
        let mut fat_sectors = Vec::with_capacity(limit);
        let mut seen = HashSet::with_capacity(limit);
        // Pushes one DIFAT entry; returns false once `limit` IDs are collected.
        let mut push = |sid: u32| -> Result<bool> {
            if fat_sectors.len() == limit {
                return Ok(false);
            }
            if sid <= MAXREGSECT {
                // A conformant file never lists a FAT sector twice; a repeat
                // would copy one sector's entries into the FAT again.
                if !seen.insert(sid) {
                    cold_path();
                    return Err(ModernXlsxError::Security(
                        "duplicate FAT sector in OLE2 DIFAT".into(),
                    ));
                }
                fat_sectors.push(sid);
            }
            Ok(true)
        };

        for sid in header_difat {
            if !push(sid)? {
                break;
            }
        }

        let per_sector = self.sector_size / 4 - 1;
        let mut visited = HashSet::new();
        let mut next = first_difat;
        let mut walked = 0;
        'chain: while walked < num_difat && next <= MAXREGSECT && visited.insert(next) {
            let offset = self.sector_offset(next);
            let end = offset.saturating_add(self.sector_size).min(data.len());
            if offset >= end {
                break;
            }
            let words = data[offset..end].as_chunks::<4>().0;
            for word in words.iter().take(per_sector) {
                if !push(u32::from_le_bytes(*word))? {
                    break 'chain;
                }
            }
            next = words
                .get(per_sector)
                .map_or(ENDOFCHAIN, |w| u32::from_le_bytes(*w));
            walked += 1;
        }
        Ok(fat_sectors)
    }

    /// Returns the byte offset of sector `sid`: `(sid + 1) * sector_size`,
    /// because the header fills the first sector (512 bytes in v3, 4096 in v4).
    ///
    /// `sid` comes from the attacker-controlled FAT/DIFAT, so the arithmetic
    /// saturates in `usize`: on a 32-bit target (wasm) a large `sid` yields
    /// `usize::MAX` instead of an overflow panic. Every caller computes the
    /// slice end as `offset.saturating_add(sector_size).min(data.len())` and
    /// reads nothing when `offset` is not below it.
    #[inline]
    fn sector_offset(&self, sid: u32) -> usize {
        (sid as usize)
            .saturating_add(1)
            .saturating_mul(self.sector_size)
    }
}

fn invalid_header(detail: String) -> ModernXlsxError {
    ModernXlsxError::UnrecognizedFormat(format!("Invalid OLE2 header: {detail}"))
}

/// Reads the FAT (File Allocation Table) to follow sector chains. A FAT sector
/// past the end of the data reads as free entries, so later sectors keep their
/// indexes.
fn read_fat(data: &[u8], header: &Ole2Header) -> Vec<u32> {
    let entries_per_sector = header.sector_size / 4;
    let mut fat = Vec::with_capacity(header.fat_sectors.len() * entries_per_sector);
    for &sid in &header.fat_sectors {
        let offset = header.sector_offset(sid);
        let end = offset.saturating_add(header.sector_size).min(data.len());
        let sector_end = fat.len() + entries_per_sector;
        if offset < end {
            fat.extend(
                data[offset..end]
                    .as_chunks::<4>()
                    .0
                    .iter()
                    .map(|c| u32::from_le_bytes(*c)),
            );
        }
        fat.resize(sector_end, FREESECT);
    }
    fat
}

/// Follows a FAT (or MiniFAT) chain from `start` and collects its sector IDs.
/// The chain ends at a marker or at an ID outside `fat`; a sector visited twice
/// is a loop and gives a `Security` error, so a chain never repeats a sector.
fn follow_chain(fat: &[u32], start: u32) -> Result<Vec<u32>> {
    let mut chain = Vec::with_capacity(16);
    let mut visited = vec![0u64; fat.len().div_ceil(64)];
    let mut current = start;
    while current <= MAXREGSECT
        && let Some(&next) = fat.get(current as usize)
    {
        let (word, bit) = (current as usize / 64, 1u64 << (current % 64));
        if visited[word] & bit != 0 {
            cold_path();
            return Err(ModernXlsxError::Security(format!(
                "OLE2 sector chain loops at sector {current}"
            )));
        }
        visited[word] |= bit;
        chain.push(current);
        current = next;
    }
    Ok(chain)
}

/// Reads up to `size` bytes of a stream from the OLE2 file given a sector chain.
/// The buffer never reserves more than `data.len()` bytes, so a long chain of
/// sectors past the end of the file cannot request a huge allocation.
fn read_sectors(data: &[u8], header: &Ole2Header, chain: &[u32], size: usize) -> Vec<u8> {
    let mut result = Vec::with_capacity(
        size.min(chain.len().saturating_mul(header.sector_size))
            .min(data.len()),
    );
    for &sid in chain {
        let offset = header.sector_offset(sid);
        let end = offset.saturating_add(header.sector_size).min(data.len());
        if offset < end {
            result.extend_from_slice(&data[offset..end]);
        }
    }
    result.truncate(size);
    result
}

/// Reads every byte of a FAT chain (bounded by the data, since a chain never
/// repeats a sector).
fn read_chain(data: &[u8], header: &Ole2Header, chain: &[u32]) -> Vec<u8> {
    read_sectors(
        data,
        header,
        chain,
        chain.len().saturating_mul(header.sector_size),
    )
}

/// Checks a directory entry's declared size against the bytes its chain holds.
/// A larger declaration is hostile input and gives a `Security` error.
fn declared_size(name: &str, declared: u64, available: usize) -> Result<usize> {
    usize::try_from(declared).ok().filter(|&n| n <= available).ok_or_else(|| {
        cold_path();
        ModernXlsxError::Security(format!(
            "OLE2 stream '{name}' declares {declared} bytes but its sector chain holds {available}"
        ))
    })
}

/// Directory entry from OLE2 compound document.
pub(crate) struct DirEntry {
    name: String,
    entry_type: u8,
    start_sector: u32,
    size: u64,
}

impl DirEntry {
    /// A stream (not a storage) whose name equals `name`, ignoring ASCII case.
    /// MS-CFB compares names through an uppercase mapping (§2.6.4), which
    /// agrees with this for ASCII names.
    pub(crate) fn is_stream_named(&self, name: &str) -> bool {
        self.entry_type == STREAM_OBJECT && self.name.eq_ignore_ascii_case(name)
    }
}

/// Reads directory entries from the directory sector chain.
fn read_directory(data: &[u8], header: &Ole2Header, fat: &[u32]) -> Result<Vec<DirEntry>> {
    let chain = follow_chain(fat, header.first_dir_sector)?;
    let dir_bytes = read_chain(data, header, &chain);

    let mut entries = Vec::with_capacity(dir_bytes.len() / 128);
    for chunk in dir_bytes.as_chunks::<128>().0 {
        let entry_type = chunk[66];
        if entry_type == 0 {
            continue;
        } // empty entry

        // Read name: UTF-16LE in the first 64 bytes of the entry; length in
        // bytes at offset 64-65. Per MS-CFB §2.6.1 the name field is at most 64
        // bytes (32 UTF-16 code units) including the terminating NUL, so a
        // malformed length must be clamped — otherwise the loop below would index
        // past the 64-byte name field (and the 128-byte entry) and panic.
        let name_len = (u16::from_le_bytes([chunk[64], chunk[65]]) as usize).min(64);
        let name_units = name_len.saturating_sub(2) / 2; // drop the UTF-16 NUL terminator
        let name: String = (0..name_units)
            .map(|i| u16::from_le_bytes([chunk[i * 2], chunk[i * 2 + 1]]))
            .map(|c| char::from_u32(u32::from(c)).unwrap_or('\u{FFFD}'))
            .collect();

        let start_sector = u32::from_le_bytes(chunk[116..120].try_into().unwrap_or_default());
        let size = if header.sector_size == 4096 {
            u64::from_le_bytes(chunk[120..128].try_into().unwrap_or_default())
        } else {
            u32::from_le_bytes(chunk[120..124].try_into().unwrap_or_default()) as u64
        };

        entries.push(DirEntry {
            name,
            entry_type,
            start_sector,
            size,
        });
    }
    Ok(entries)
}

/// A compound file parsed once: its validated header, its FAT and its
/// directory. A caller that reads several streams of one file (classification,
/// then `EncryptionInfo`, then `EncryptedPackage`) parses it once with
/// [`Cfb::parse`] and reads each stream with [`Cfb::read_stream`]. The MiniFAT
/// and the mini stream stay lazy: each mini-stream read collects them again.
pub(crate) struct Cfb<'a> {
    data: &'a [u8],
    header: Ole2Header,
    fat: Vec<u32>,
    dir: Vec<DirEntry>,
}

impl<'a> Cfb<'a> {
    /// Parses the header (with the DIFAT), the FAT and the directory, in that
    /// order and with the errors [`read_stream`] gives before its name lookup.
    pub(crate) fn parse(data: &'a [u8]) -> Result<Self> {
        let header = Ole2Header::parse(data)?;
        let fat = read_fat(data, &header);
        let dir = read_directory(data, &header, &fat)?;
        Ok(Self {
            data,
            header,
            fat,
            dir,
        })
    }

    /// The directory entries, empty slots skipped, in directory order.
    pub(crate) fn entries(&self) -> &[DirEntry] {
        &self.dir
    }

    /// Reads a named stream. The name matches stream entries only, ignoring
    /// ASCII case.
    ///
    /// A stream shorter than 4096 bytes is read from the mini stream when the
    /// file has a MiniFAT and a non-empty root entry. Otherwise it is read
    /// through the regular FAT, which is where modern-xlsx 1.2.1 and earlier
    /// stored small streams (they wrote no MiniFAT).
    pub(crate) fn read_stream(&self, name: &str) -> Result<Vec<u8>> {
        let entry = self
            .dir
            .iter()
            .find(|e| e.is_stream_named(name))
            .ok_or_else(|| {
                cold_path();
                ModernXlsxError::MissingPart(format!(
                    "OLE2 stream '{name}' not found in compound document — required for decryption"
                ))
            })?;

        if entry.size < MINI_STREAM_CUTOFF && self.header.first_minifat_sector != ENDOFCHAIN {
            let root = self
                .dir
                .iter()
                .find(|e| e.entry_type == ROOT_STORAGE_OBJECT)
                .filter(|root| root.size > 0);
            if let Some(root) = root {
                return self.read_mini_stream(root, entry, name);
            }
        }

        let chain = follow_chain(&self.fat, entry.start_sector)?;
        let mut bytes = read_chain(self.data, &self.header, &chain);
        let size = declared_size(name, entry.size, bytes.len())?;
        bytes.truncate(size);
        Ok(bytes)
    }

    /// Reads a stream stored in the mini stream: the MiniFAT is the FAT chain
    /// from header @60, the container is the root entry's FAT chain bounded by
    /// its size, and the stream is the 64-byte mini sectors of its MiniFAT
    /// chain.
    fn read_mini_stream(&self, root: &DirEntry, entry: &DirEntry, name: &str) -> Result<Vec<u8>> {
        let (data, header, fat) = (self.data, &self.header, self.fat.as_slice());
        let minifat_chain = follow_chain(fat, header.first_minifat_sector)?;
        let minifat: Vec<u32> = read_chain(data, header, &minifat_chain)
            .as_chunks::<4>()
            .0
            .iter()
            .map(|c| u32::from_le_bytes(*c))
            .collect();

        let root_chain = follow_chain(fat, root.start_sector)?;
        let mut container = read_chain(data, header, &root_chain);
        container.truncate(declared_size(&root.name, root.size, container.len())?);

        let chain = follow_chain(&minifat, entry.start_sector)?;
        let mut bytes = Vec::with_capacity(
            chain
                .len()
                .saturating_mul(MINI_SECTOR_SIZE)
                .min(container.len()),
        );
        for &sid in &chain {
            let offset = (sid as usize).saturating_mul(MINI_SECTOR_SIZE);
            let end = offset.saturating_add(MINI_SECTOR_SIZE).min(container.len());
            if offset < end {
                bytes.extend_from_slice(&container[offset..end]);
            }
        }
        let size = declared_size(name, entry.size, bytes.len())?;
        bytes.truncate(size);
        Ok(bytes)
    }
}

/// Reads a named stream from the OLE2 document. The name matches stream entries
/// only, ignoring ASCII case.
///
/// A stream shorter than 4096 bytes is read from the mini stream when the file
/// has a MiniFAT and a non-empty root entry. Otherwise it is read through the
/// regular FAT, which is where modern-xlsx 1.2.1 and earlier stored small
/// streams (they wrote no MiniFAT).
///
/// Each call parses the header, the FAT and the directory again; reading
/// several streams of one file through one parse is the crate-internal `Cfb`.
pub fn read_stream(data: &[u8], name: &str) -> Result<Vec<u8>> {
    Cfb::parse(data)?.read_stream(name)
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::ole2::detect::{Ole2Kind, classify_ole2};

    // -----------------------------------------------------------------------
    // MS-CFB conformance: files built by the independent `cfb` crate
    // -----------------------------------------------------------------------

    use std::io::{Cursor, Write};

    const ENDOFCHAIN_U32: u32 = 0xFFFF_FFFE;
    const FREESECT_U32: u32 = 0xFFFF_FFFF;
    const FATSECT_U32: u32 = 0xFFFF_FFFD;
    const DIFSECT_U32: u32 = 0xFFFF_FFFC;

    /// Deterministic xorshift bytes, so a misplaced sector shows up as a mismatch.
    fn pattern(len: usize, seed: u32) -> Vec<u8> {
        let mut x = seed.wrapping_mul(2_654_435_761).wrapping_add(1);
        (0..len)
            .map(|_| {
                x ^= x << 13;
                x ^= x >> 17;
                x ^= x << 5;
                x.to_le_bytes()[0]
            })
            .collect()
    }

    fn u32_at(bytes: &[u8], offset: usize) -> u32 {
        u32::from_le_bytes(bytes[offset..offset + 4].try_into().unwrap())
    }

    /// A conformant compound file written by the `cfb` dev-dependency: streams
    /// under 4096 B go to the mini stream, and a FAT over 109 sectors gets DIFAT
    /// sectors.
    fn cfb_file(version: cfb::Version, streams: &[(&str, &[u8])]) -> Vec<u8> {
        let mut comp = cfb::CompoundFile::create_with_version(version, Cursor::new(Vec::new()))
            .expect("create compound file");
        for (name, bytes) in streams {
            let mut stream = comp
                .create_stream(format!("/{name}"))
                .expect("create stream");
            stream.write_all(bytes).expect("write stream");
        }
        comp.flush().expect("flush compound file");
        comp.into_inner().into_inner()
    }

    #[test]
    fn reads_mini_stream_of_conformant_v3_file() {
        let small = pattern(1024, 1);
        let big = pattern(10_000, 2);
        let file = cfb_file(
            cfb::Version::V3,
            &[("EncryptionInfo", &small), ("EncryptedPackage", &big)],
        );
        assert_ne!(u32_at(&file, 60), ENDOFCHAIN_U32, "cfb must emit a MiniFAT");
        assert_eq!(read_stream(&file, "EncryptionInfo").unwrap(), small);
        assert_eq!(read_stream(&file, "EncryptedPackage").unwrap(), big);
    }

    #[test]
    fn reads_difat_backed_8_mib_stream() {
        let big = pattern(8 * 1024 * 1024, 3);
        let file = cfb_file(cfb::Version::V3, &[("EncryptedPackage", &big)]);
        assert!(u32_at(&file, 44) > 109, "FAT must outgrow the header DIFAT");
        assert!(u32_at(&file, 72) >= 1, "cfb must emit DIFAT sectors");
        let got = read_stream(&file, "EncryptedPackage").unwrap();
        assert_eq!(got.len(), big.len());
        assert!(got == big, "8 MiB stream differs");
    }

    #[test]
    fn reads_v4_regular_and_mini_streams() {
        let big = pattern(10_000, 4);
        let small = pattern(1024, 5);
        let file = cfb_file(
            cfb::Version::V4,
            &[("EncryptedPackage", &big), ("EncryptionInfo", &small)],
        );
        assert_eq!(u16::from_le_bytes([file[26], file[27]]), 4);
        assert_eq!(read_stream(&file, "EncryptedPackage").unwrap(), big);
        assert_eq!(read_stream(&file, "EncryptionInfo").unwrap(), small);
    }

    #[test]
    fn stream_lookup_ignores_ascii_case() {
        let info = pattern(100, 6);
        let package = pattern(5000, 7);
        let file = cfb_file(
            cfb::Version::V3,
            &[("encryptioninfo", &info), ("ENCRYPTEDPACKAGE", &package)],
        );
        assert_eq!(read_stream(&file, "EncryptionInfo").unwrap(), info);
        assert_eq!(read_stream(&file, "EncryptedPackage").unwrap(), package);
        assert_eq!(classify_ole2(&file).unwrap(), Ole2Kind::EncryptedXlsx);
    }

    /// One `Cfb::parse` serves every stream read of a file, with the bytes and
    /// the errors the one-shot `read_stream` gives for each name.
    #[test]
    fn one_parse_reads_every_stream_like_read_stream() {
        let info = pattern(1024, 10);
        let package = pattern(10_000, 11);
        let file = cfb_file(
            cfb::Version::V3,
            &[("EncryptionInfo", &info), ("EncryptedPackage", &package)],
        );
        let cfb = Cfb::parse(&file).unwrap();
        assert_eq!(cfb.read_stream("EncryptionInfo").unwrap(), info);
        assert_eq!(cfb.read_stream("EncryptedPackage").unwrap(), package);
        assert_eq!(cfb.read_stream("EncryptionInfo").unwrap(), info);
        assert_eq!(
            cfb.read_stream("Workbook").unwrap_err().to_coded_string(),
            read_stream(&file, "Workbook")
                .unwrap_err()
                .to_coded_string()
        );
        let hostile = one_sector_stream(4, u64::MAX);
        assert_eq!(
            Cfb::parse(&hostile)
                .unwrap()
                .read_stream("EncryptionInfo")
                .unwrap_err()
                .to_coded_string(),
            read_stream(&hostile, "EncryptionInfo")
                .unwrap_err()
                .to_coded_string()
        );
    }

    #[test]
    fn storages_never_match_stream_names() {
        let mut comp =
            cfb::CompoundFile::create_with_version(cfb::Version::V3, Cursor::new(Vec::new()))
                .unwrap();
        comp.create_storage("/EncryptionInfo").unwrap();
        comp.create_storage("/Workbook").unwrap();
        comp.create_stream("/EncryptedPackage")
            .unwrap()
            .write_all(&[1; 64])
            .unwrap();
        comp.flush().unwrap();
        let file = comp.into_inner().into_inner();
        assert_eq!(classify_ole2(&file).unwrap(), Ole2Kind::Unknown);
        let err = read_stream(&file, "EncryptionInfo").unwrap_err();
        assert_eq!(err.code(), "MISSING_PART", "{err}");
    }

    #[test]
    fn mini_chain_loop_is_a_security_error() {
        let file = cfb_file(cfb::Version::V3, &[("EncryptionInfo", &pattern(1024, 8))]);
        // The only mini stream occupies mini sectors 0..16; close its chain on itself.
        let minifat = (u32_at(&file, 60) as usize + 1) * 512;
        let mut hostile = file.clone();
        hostile[minifat + 15 * 4..minifat + 16 * 4].copy_from_slice(&0u32.to_le_bytes());
        let err = read_stream(&hostile, "EncryptionInfo").unwrap_err();
        assert_eq!(err.code(), "SECURITY", "{err}");
    }

    // -----------------------------------------------------------------------
    // Interop fixtures (tests/fixtures/msoffcrypto/README.md, tests/fixtures/legacy)
    // -----------------------------------------------------------------------

    #[test]
    fn excel_agile_fixture_reads_encryption_info_from_mini_stream() {
        let file = include_bytes!(concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/tests/fixtures/msoffcrypto/example_password.xlsx"
        ));
        let info = read_stream(file, "EncryptionInfo").unwrap();
        assert_eq!(info.len(), 1289);
        assert_eq!(info[..8], [0x04, 0x00, 0x04, 0x00, 0x40, 0x00, 0x00, 0x00]);
        assert_eq!(classify_ole2(file).unwrap(), Ole2Kind::EncryptedXlsx);
    }

    #[test]
    fn standard_fixture_reads_encryption_info() {
        let file = include_bytes!(concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/tests/fixtures/msoffcrypto/ecma376standard_password.docx"
        ));
        let info = read_stream(file, "EncryptionInfo").unwrap();
        assert_eq!(info.len(), 224);
        assert_eq!(
            info[2..4],
            [2, 0],
            "Standard encryption has minor version 2"
        );
    }

    /// `modern-xlsx-1.2.1-agile.xlsx` is the output of modern-xlsx-core at commit
    /// 07b54a1 (release 1.2.1): `encrypt_file(example.xlsx, "Password1234_")`. That
    /// writer kept small streams in regular sectors and wrote no MiniFAT (header
    /// @60 = ENDOFCHAIN), so the reader must take the regular-FAT path for them.
    #[test]
    fn modern_xlsx_1_2_1_layout_reads_through_regular_fat() {
        let file = include_bytes!(concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/tests/fixtures/legacy/modern-xlsx-1.2.1-agile.xlsx"
        ));
        assert_eq!(u32_at(file, 60), ENDOFCHAIN_U32);
        assert_eq!(read_stream(file, "EncryptionInfo").unwrap().len(), 1342);
        assert_eq!(read_stream(file, "EncryptedPackage").unwrap().len(), 8392);
        assert_eq!(classify_ole2(file).unwrap(), Ole2Kind::EncryptedXlsx);
    }

    // -----------------------------------------------------------------------
    // Hostile inputs
    // -----------------------------------------------------------------------

    /// A hand-built compound file: a valid header (magic, byte order FE FF,
    /// (3, 9) or (4, 12), mini shift 6, cutoff 4096, no MiniFAT, no DIFAT, every
    /// header DIFAT slot free) and `sectors` zeroed sectors after it. Sector
    /// `sid` starts at `(sid + 1) * sector_size`.
    struct RawCfb {
        buf: Vec<u8>,
        sector_size: usize,
    }

    impl RawCfb {
        fn new(major: u16, sectors: usize) -> Self {
            let (shift, sector_size) = if major == 4 { (12, 4096) } else { (9, 512) };
            let mut raw = Self {
                buf: vec![0; sector_size * (sectors + 1)],
                sector_size,
            };
            raw.buf[..8].copy_from_slice(&crate::ole2::OLE2_MAGIC);
            raw.put_u16(24, 0x3E);
            raw.put_u16(26, major);
            raw.put_u16(28, 0xFFFE);
            raw.put_u16(30, shift);
            raw.put_u16(32, 6);
            raw.put_u32(56, 4096);
            raw.put_u32(60, ENDOFCHAIN_U32);
            raw.put_u32(68, ENDOFCHAIN_U32);
            for slot in 0..109 {
                raw.put_u32(76 + slot * 4, FREESECT_U32);
            }
            raw
        }

        fn put_u16(&mut self, offset: usize, value: u16) {
            self.buf[offset..offset + 2].copy_from_slice(&value.to_le_bytes());
        }

        fn put_u32(&mut self, offset: usize, value: u32) {
            self.buf[offset..offset + 4].copy_from_slice(&value.to_le_bytes());
        }

        fn sector(&self, sid: u32) -> usize {
            (sid as usize + 1) * self.sector_size
        }

        /// Fills sector `sid` with `words`, padded with FREESECT.
        fn words(&mut self, sid: u32, words: &[u32]) {
            let base = self.sector(sid);
            for i in 0..self.sector_size / 4 {
                self.put_u32(base + i * 4, words.get(i).copied().unwrap_or(FREESECT_U32));
            }
        }

        /// Makes `sid` the single FAT sector, listed in header slot 0.
        fn fat(&mut self, sid: u32, entries: &[u32]) {
            self.put_u32(44, 1);
            self.put_u32(76, sid);
            self.words(sid, entries);
        }

        fn dir_entry(&mut self, dir_sid: u32, slot: usize, entry: (&str, u8, u32, u64)) {
            let (name, entry_type, start, size) = entry;
            self.put_u32(48, dir_sid);
            let at = self.sector(dir_sid) + slot * 128;
            let units: Vec<u16> = name.encode_utf16().collect();
            for (i, unit) in units.iter().enumerate() {
                self.put_u16(at + i * 2, *unit);
            }
            self.put_u16(at + 64, ((units.len() + 1) * 2) as u16);
            self.buf[at + 66] = entry_type;
            self.put_u32(at + 68, FREESECT_U32);
            self.put_u32(at + 72, FREESECT_U32);
            self.put_u32(at + 76, FREESECT_U32);
            self.put_u32(at + 116, start);
            self.buf[at + 120..at + 128].copy_from_slice(&size.to_le_bytes());
        }
    }

    /// FAT in sector 0, directory in sector 1, one data sector 2 that the
    /// `EncryptionInfo` entry claims holds `info_size` bytes.
    fn one_sector_stream(major: u16, info_size: u64) -> Vec<u8> {
        let mut raw = RawCfb::new(major, 3);
        raw.fat(0, &[FATSECT_U32, ENDOFCHAIN_U32, ENDOFCHAIN_U32]);
        raw.dir_entry(1, 0, ("Root Entry", 5, ENDOFCHAIN_U32, 0));
        raw.dir_entry(1, 1, ("EncryptionInfo", 2, 2, info_size));
        raw.dir_entry(1, 2, ("EncryptedPackage", 2, ENDOFCHAIN_U32, 0));
        raw.buf
    }

    #[test]
    fn hand_built_file_is_valid_control() {
        let file = one_sector_stream(4, 100);
        assert_eq!(classify_ole2(&file).unwrap(), Ole2Kind::EncryptedXlsx);
        assert_eq!(read_stream(&file, "EncryptionInfo").unwrap(), vec![0; 100]);
    }

    #[test]
    fn v4_entry_size_u64_max_is_a_security_error() {
        let file = one_sector_stream(4, u64::MAX);
        assert_eq!(classify_ole2(&file).unwrap(), Ole2Kind::EncryptedXlsx);
        let err = read_stream(&file, "EncryptionInfo").unwrap_err();
        assert_eq!(err.code(), "SECURITY", "{err}");
        assert!(
            err.to_string().contains("'EncryptionInfo' declares"),
            "{err}"
        );
    }

    #[test]
    fn v3_size_over_one_sector_chain_is_a_security_error() {
        let file = one_sector_stream(3, 0xFFFF_FFFF);
        let err = read_stream(&file, "EncryptionInfo").unwrap_err();
        assert_eq!(err.code(), "SECURITY", "{err}");
    }

    #[test]
    fn self_looping_directory_is_a_security_error() {
        let mut raw = RawCfb::new(4, 2);
        raw.fat(0, &[FATSECT_U32, 1]);
        raw.dir_entry(1, 0, ("Root Entry", 5, ENDOFCHAIN_U32, 0));
        let err = classify_ole2(&raw.buf).unwrap_err();
        assert_eq!(err.code(), "SECURITY", "{err}");
        assert!(err.to_string().contains("loops at sector 1"), "{err}");
        assert_eq!(read_stream(&raw.buf, "x").unwrap_err().code(), "SECURITY");
    }

    /// All 109 header DIFAT slots name FAT sector 0. The second listing is a
    /// `Security` error ("duplicate FAT sector"), raised while the header is
    /// parsed and before any FAT is read or chain followed.
    #[test]
    fn repeated_header_fat_sector_is_rejected() {
        let mut raw = RawCfb::new(4, 2);
        raw.fat(0, &[FATSECT_U32, 1]);
        raw.dir_entry(1, 0, ("Root Entry", 5, ENDOFCHAIN_U32, 0));
        raw.put_u32(44, 109);
        for slot in 0..109 {
            raw.put_u32(76 + slot * 4, 0);
        }
        let err = classify_ole2(&raw.buf).unwrap_err();
        assert_eq!(err.code(), "SECURITY", "{err}");
        assert!(err.to_string().contains("duplicate FAT sector"), "{err}");
    }

    /// FAT in sector 0, a DIFAT sector 1 whose next pointer is itself, the
    /// directory in sector 2; the header claims 200 FAT and 200 DIFAT sectors.
    fn self_looping_difat(difat_entries: &[u32]) -> Vec<u8> {
        let mut raw = RawCfb::new(3, 3);
        raw.fat(0, &[FATSECT_U32, DIFSECT_U32, ENDOFCHAIN_U32]);
        raw.put_u32(44, 200);
        raw.put_u32(68, 1);
        raw.put_u32(72, 200);
        let mut difat = difat_entries.to_vec();
        difat.resize(127, FREESECT_U32);
        difat.push(1);
        raw.words(1, &difat);
        raw.dir_entry(2, 0, ("Root Entry", 5, ENDOFCHAIN_U32, 0));
        raw.dir_entry(2, 1, ("Workbook", 2, ENDOFCHAIN_U32, 0));
        raw.buf
    }

    #[test]
    fn self_looping_difat_chain_terminates() {
        let file = self_looping_difat(&[]);
        assert_eq!(classify_ole2(&file).unwrap(), Ole2Kind::LegacyXls);
        assert_eq!(read_stream(&file, "Workbook").unwrap(), Vec::<u8>::new());
    }

    #[test]
    fn fat_sector_repeated_in_difat_is_rejected() {
        let err = classify_ole2(&self_looping_difat(&[0])).unwrap_err();
        assert_eq!(err.code(), "SECURITY", "{err}");
    }

    #[test]
    fn invalid_headers_are_unrecognized_format() {
        let valid = cfb_file(cfb::Version::V3, &[("EncryptionInfo", &[1; 10])]);
        let patches: [(usize, u16); 4] = [
            (28, 0xFEFF), // byte order mark FF FE
            (30, 12),     // major 3 with a 4096-byte sector shift
            (26, 4),      // major 4 with a 512-byte sector shift
            (32, 7),      // mini sector shift other than 6
        ];
        for (offset, value) in patches {
            let mut file = valid.clone();
            file[offset..offset + 2].copy_from_slice(&value.to_le_bytes());
            let err = classify_ole2(&file).unwrap_err();
            assert_eq!(
                err.code(),
                "UNRECOGNIZED_FORMAT",
                "@{offset}={value:#x}: {err}"
            );
            assert!(err.to_string().contains("Invalid OLE2 header"), "{err}");
            let err = read_stream(&file, "EncryptionInfo").unwrap_err();
            assert_eq!(
                err.code(),
                "UNRECOGNIZED_FORMAT",
                "@{offset}={value:#x}: {err}"
            );
        }
        for cutoff in [4095_u32, 0] {
            let mut file = valid.clone();
            file[56..60].copy_from_slice(&cutoff.to_le_bytes());
            let err = classify_ole2(&file).unwrap_err();
            assert_eq!(err.code(), "UNRECOGNIZED_FORMAT", "cutoff {cutoff}: {err}");
            assert!(err.to_string().contains("Invalid OLE2 header"), "{err}");
            assert!(err.to_string().contains("mini stream cutoff"), "{err}");
            let err = read_stream(&file, "EncryptionInfo").unwrap_err();
            assert_eq!(err.code(), "UNRECOGNIZED_FORMAT", "cutoff {cutoff}: {err}");
            assert!(err.to_string().contains("mini stream cutoff"), "{err}");
        }
    }

    #[test]
    fn fuzz_crash_input_errors_without_panic() {
        let crash = include_bytes!(concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/tests/fixtures/fuzz_crash_ole2.bin"
        ));
        assert!(classify_ole2(crash).is_err());
        assert!(read_stream(crash, "EncryptionInfo").is_err());
    }

    #[test]
    fn far_sector_ids_read_nothing() {
        let file = one_sector_stream(3, 100);
        let header = Ole2Header::parse(&file).unwrap();
        assert!(read_sectors(&file, &header, &[0xFFFF_FFF0, 0xFFFF_FFFA], 1024).is_empty());
    }

    /// A v4 file of 16,384 bytes whose FAT chains sectors 2..=1023, almost all
    /// past the end of the data: reading the chain reserves at most the file's
    /// length, not 1022 * 4096 bytes, and the 5000 bytes the entry declares
    /// against the 4096 its chain holds are a `Security` error.
    #[test]
    fn chain_past_end_of_file_reserves_at_most_the_file() {
        let mut raw = RawCfb::new(4, 3);
        let mut fat = vec![FATSECT_U32, ENDOFCHAIN_U32];
        fat.extend((2..1023).map(|i: u32| i + 1));
        fat.push(ENDOFCHAIN_U32);
        raw.fat(0, &fat);
        raw.dir_entry(1, 0, ("Root Entry", 5, ENDOFCHAIN_U32, 0));
        raw.dir_entry(1, 1, ("EncryptedPackage", 2, 2, 5000));
        let file = raw.buf;
        let header = Ole2Header::parse(&file).unwrap();
        let fat = read_fat(&file, &header);
        let chain = follow_chain(&fat, 2).unwrap();
        assert_eq!(chain.len(), 1022);
        assert!(read_chain(&file, &header, &chain).capacity() <= file.len());
        let err = read_stream(&file, "EncryptedPackage").unwrap_err();
        assert_eq!(err.code(), "SECURITY", "{err}");
    }

    /// On a 32-bit target `(sid + 1) * sector_size` saturates, and the
    /// slice end must saturate too instead of overflowing.
    #[cfg(target_pointer_width = "32")]
    #[test]
    fn sector_offset_saturates_on_32_bit() {
        let file = one_sector_stream(3, 100);
        let header = Ole2Header::parse(&file).unwrap();
        assert_eq!(header.sector_offset(0xFFFF_FFF0), usize::MAX);
        assert!(read_sectors(&file, &header, &[0xFFFF_FFF0], 512).is_empty());
    }
}
