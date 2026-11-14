//! MS-CFB compound-file writer.
//!
//! [`write_ole2`] lays out a version 3 file (512-byte sectors): streams under
//! 4096 bytes go into the mini stream with a MiniFAT, larger ones into regular
//! sectors, and FAT sector IDs past the 109 header slots into DIFAT sectors.
//! A stream may hold at most `MAX_STREAM_SIZE` (0x8000_0000 bytes, the
//! version 3 limit of MS-CFB §2.6.1). A larger stream, or a document whose
//! sectors would pass MAXREGSECT or the platform's address space, is a
//! `SizeLimit` error.

use core::cmp::Ordering;
use core::hint::cold_path;

use super::{ENDOFCHAIN, FREESECT, MAXREGSECT, MINI_SECTOR_SIZE};
use crate::errors::{ModernXlsxError, Result};

/// Sector size of a version 3 file: sector shift 9, 2^9 = 512 bytes (MS-CFB §2.2).
const SECTOR_SIZE: usize = 512;

/// Streams shorter than the mini stream cutoff live in the mini stream (MS-CFB §2.2).
const MINI_STREAM_CUTOFF: u32 = 4096;

/// Largest stream a version 3 file can hold (MS-CFB §2.6.1, Stream Size).
const MAX_STREAM_SIZE: u32 = 0x8000_0000;

/// FAT, MiniFAT and DIFAT sectors each hold 512 / 4 = 128 little-endian u32 entries.
const ENTRIES_PER_SECTOR: u32 = 128;

/// Mini sectors per regular sector of the mini-stream container (512 / 64).
const MINI_SECTORS_PER_SECTOR: u32 = 8;

/// FAT sector IDs held by the header DIFAT array (MS-CFB §2.2, offsets 76..512).
const HEADER_DIFAT_ENTRIES: u32 = 109;

/// FAT sector IDs per DIFAT sector; its last u32 is the next DIFAT sector ID (MS-CFB §2.5).
const DIFAT_ENTRIES_PER_SECTOR: u32 = ENTRIES_PER_SECTOR - 1;

/// Directory entries are 128 bytes, four per 512-byte sector (MS-CFB §2.6).
const DIR_ENTRY_SIZE: usize = 128;
const DIR_ENTRIES_PER_SECTOR: u32 = 4;

/// Sector ID markers for DIFAT and FAT sectors (MS-CFB §2.1); the shared
/// MAXREGSECT, ENDOFCHAIN and FREESECT live in the parent `cfb` module.
const DIFSECT: u32 = 0xFFFF_FFFC;
const FATSECT: u32 = 0xFFFF_FFFD;

/// Stream ID markers (MS-CFB §2.6.1).
const MAXREGSID: u32 = 0xFFFF_FFFA;
const NOSTREAM: u32 = 0xFFFF_FFFF;

/// Directory object types (MS-CFB §2.6.1).
const DIR_TYPE_STREAM: u8 = 2;
const DIR_TYPE_ROOT: u8 = 5;

/// Color flag black (MS-CFB §2.6.1). Every node of the sibling tree is black,
/// which MS-CFB §2.6.4 names as the simplest valid red-black tree.
const COLOR_BLACK: u8 = 1;

/// Where one stream's bytes live.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum Placement {
    /// A zero-length stream: start sector ENDOFCHAIN, size 0.
    Empty,
    /// `count` mini sectors from mini sector `start`, chained in the MiniFAT.
    Mini { start: u32, count: u32 },
    /// `count` regular sectors from sector `start`, chained in the FAT.
    Regular { start: u32, count: u32 },
}

/// One stream's placement and its size in bytes.
#[derive(Clone, Copy, Debug)]
struct StreamPlan {
    placement: Placement,
    size: u32,
}

/// The sector plan of a compound file, computed from stream lengths alone.
///
/// After the header the file holds, in order: the regular stream sectors, the
/// mini-stream container, the MiniFAT, the directory, the FAT and the DIFAT
/// sectors. Every sector count is checked, so the region starts below never
/// overflow and every sector ID is at most MAXREGSECT.
#[derive(Debug)]
struct Layout {
    streams: Vec<StreamPlan>,
    /// Regular sectors holding streams of 4096 bytes or more.
    regular: u32,
    /// Bytes in the mini stream: its 64-byte mini sectors hold the streams
    /// under 4096 bytes.
    mini_stream_size: u32,
    /// Regular sectors of the mini-stream container.
    container: u32,
    minifat: u32,
    dir: u32,
    fat: u32,
    difat: u32,
    /// Header plus every sector, in bytes.
    file_len: usize,
}

impl Layout {
    /// Plans the sectors for streams of the given `(name, length)`.
    ///
    /// The FAT must also cover its own sectors and the DIFAT sectors that list
    /// them, so its size is the least fixed point of
    /// `fat = ceil((data + container + minifat + dir + fat + difat) / 128)`,
    /// where `difat = ceil((fat - 109) / 127)` once `fat` passes the 109 header
    /// DIFAT slots.
    fn plan(lens: &[(&str, usize)]) -> Result<Self> {
        let mut streams = Vec::with_capacity(lens.len());
        let (mut regular, mut mini) = (0u32, 0u32);
        for &(name, len) in lens {
            let size = u32::try_from(len)
                .ok()
                .filter(|&size| size <= MAX_STREAM_SIZE)
                .ok_or_else(|| {
                    cold_path();
                    ModernXlsxError::SizeLimit(format!(
                        "OLE2 stream '{name}' is {len} bytes; MS-CFB v3 allows at most 2147483648 bytes per stream"
                    ))
                })?;
            let placement = if size == 0 {
                Placement::Empty
            } else if size < MINI_STREAM_CUTOFF {
                let (start, count) = (mini, size.div_ceil(MINI_SECTOR_SIZE as u32));
                mini = mini.checked_add(count).ok_or_else(too_many_sectors)?;
                Placement::Mini { start, count }
            } else {
                let (start, count) = (regular, size.div_ceil(SECTOR_SIZE as u32));
                regular = regular.checked_add(count).ok_or_else(too_many_sectors)?;
                Placement::Regular { start, count }
            };
            streams.push(StreamPlan { placement, size });
        }

        // The mini stream is itself a stream, so the same size limit applies.
        let mini_stream_size = mini
            .checked_mul(MINI_SECTOR_SIZE as u32)
            .filter(|&size| size <= MAX_STREAM_SIZE)
            .ok_or_else(too_many_sectors)?;
        let container = mini.div_ceil(MINI_SECTORS_PER_SECTOR);
        let minifat = mini.div_ceil(ENTRIES_PER_SECTOR);
        // The root entry plus one entry per stream; stream IDs stop at MAXREGSID.
        let entries = u32::try_from(lens.len())
            .ok()
            .filter(|&n| n <= MAXREGSID)
            .ok_or_else(too_many_sectors)?
            + 1;
        let dir = entries.div_ceil(DIR_ENTRIES_PER_SECTOR);
        let fixed = [regular, container, minifat, dir]
            .into_iter()
            .try_fold(0u32, u32::checked_add)
            .ok_or_else(too_many_sectors)?;

        let mut fat = 0u32;
        let (difat, total) = loop {
            let difat = fat
                .saturating_sub(HEADER_DIFAT_ENTRIES)
                .div_ceil(DIFAT_ENTRIES_PER_SECTOR);
            let total = fixed
                .checked_add(fat)
                .and_then(|n| n.checked_add(difat))
                .filter(|&n| n <= MAXREGSECT + 1)
                .ok_or_else(too_many_sectors)?;
            let needed = total.div_ceil(ENTRIES_PER_SECTOR);
            if needed <= fat {
                break (difat, total);
            }
            fat = needed;
        };

        let file_len = usize::try_from(total)
            .ok()
            .and_then(|n| n.checked_add(1)?.checked_mul(SECTOR_SIZE))
            .ok_or_else(|| {
                cold_path();
                ModernXlsxError::SizeLimit(format!(
                    "OLE2 document of {total} sectors exceeds this platform's address space"
                ))
            })?;

        Ok(Self {
            streams,
            regular,
            mini_stream_size,
            container,
            minifat,
            dir,
            fat,
            difat,
            file_len,
        })
    }

    const fn container_start(&self) -> u32 {
        self.regular
    }

    const fn minifat_start(&self) -> u32 {
        self.container_start() + self.container
    }

    const fn dir_start(&self) -> u32 {
        self.minifat_start() + self.minifat
    }

    const fn fat_start(&self) -> u32 {
        self.dir_start() + self.dir
    }

    const fn difat_start(&self) -> u32 {
        self.fat_start() + self.fat
    }
}

/// The error for a document whose sector or stream IDs would pass MAXREGSECT.
#[cold]
fn too_many_sectors() -> ModernXlsxError {
    ModernXlsxError::SizeLimit(
        "OLE2 document needs more sectors than MS-CFB v3 can address (0xFFFFFFFA)".into(),
    )
}

/// A directory entry for writing (MS-CFB §2.6.1).
struct WriteDirEntry<'a> {
    name: &'a str,
    entry_type: u8,
    left_sibling_id: u32,
    right_sibling_id: u32,
    child_id: u32,
    start_sector: u32,
    size: u32,
}

impl WriteDirEntry<'_> {
    /// Serialize this directory entry to 128 bytes.
    #[inline]
    fn to_bytes(&self) -> [u8; DIR_ENTRY_SIZE] {
        let mut buf = [0u8; DIR_ENTRY_SIZE];

        // Name as UTF-16LE: at most 31 code units plus the null terminator,
        // which the zeroed buffer already holds.
        let mut name_units = 0u16;
        for (slot, unit) in buf[..62]
            .as_chunks_mut::<2>()
            .0
            .iter_mut()
            .zip(self.name.encode_utf16())
        {
            *slot = unit.to_le_bytes();
            name_units += 1;
        }

        // Name length in bytes, including the null terminator.
        buf[64..66].copy_from_slice(&((name_units + 1) * 2).to_le_bytes());
        buf[66] = self.entry_type;
        buf[67] = COLOR_BLACK;
        buf[68..72].copy_from_slice(&self.left_sibling_id.to_le_bytes());
        buf[72..76].copy_from_slice(&self.right_sibling_id.to_le_bytes());
        buf[76..80].copy_from_slice(&self.child_id.to_le_bytes());
        // CLSID (80..96), state bits (96..100), creation and modified times
        // (100..116) stay zero.
        buf[116..120].copy_from_slice(&self.start_sector.to_le_bytes());
        // Stream size: v3 keeps the high 32 bits (124..128) zero.
        buf[120..124].copy_from_slice(&self.size.to_le_bytes());
        buf
    }
}

/// An unallocated directory entry: all zeroes except the left sibling, right
/// sibling and child IDs, which are NOSTREAM (MS-CFB §2.6.3).
fn unused_dir_entry() -> [u8; DIR_ENTRY_SIZE] {
    let mut buf = [0u8; DIR_ENTRY_SIZE];
    buf[68..80].fill(0xFF);
    buf
}

/// Uppercases one UTF-16 code unit the way MS-CFB §2.6.4 compares names:
/// simple case mapping of a single code unit; surrogates stay unchanged.
fn cfb_upper(unit: u16) -> u16 {
    let Some(c) = char::from_u32(u32::from(unit)) else {
        return unit;
    };
    let mut upper = c.to_uppercase();
    match (upper.next(), upper.next()) {
        (Some(u), None) => u16::try_from(u32::from(u)).unwrap_or(unit),
        _ => unit,
    }
}

/// Orders two names per MS-CFB §2.6.4: the shorter name first, then code
/// unit by code unit after uppercasing. Only the 31 stored code units count.
fn compare_names(a: &str, b: &str) -> Ordering {
    let a = a.encode_utf16().take(31);
    let b = b.encode_utf16().take(31);
    a.clone()
        .count()
        .cmp(&b.clone().count())
        .then_with(|| a.map(cfb_upper).cmp(b.map(cfb_upper)))
}

/// Links `sorted` (stream IDs in name order) into a balanced binary search
/// tree through the left and right sibling IDs and returns its root, or
/// NOSTREAM when empty. All nodes are black (MS-CFB §2.6.4).
fn link_sibling_tree(entries: &mut [WriteDirEntry<'_>], sorted: &[u32]) -> u32 {
    if sorted.is_empty() {
        return NOSTREAM;
    }
    let mid = sorted.len() / 2;
    let left = link_sibling_tree(entries, &sorted[..mid]);
    let right = link_sibling_tree(entries, &sorted[mid + 1..]);
    let node = &mut entries[sorted[mid] as usize];
    node.left_sibling_id = left;
    node.right_sibling_id = right;
    sorted[mid]
}

/// Chains `count` consecutive sectors from `start` in a FAT or MiniFAT.
fn link_chain(table: &mut [u32], start: u32, count: u32) {
    let end = start + count;
    for sid in start..end {
        table[sid as usize] = if sid + 1 < end { sid + 1 } else { ENDOFCHAIN };
    }
}

/// Appends `words` as little-endian u32s.
fn push_u32s(out: &mut Vec<u8>, words: &[u32]) {
    out.extend(words.iter().flat_map(|w| w.to_le_bytes()));
}

/// Zero-pads `out` to the next multiple of `align` bytes.
fn pad_to(out: &mut Vec<u8>, align: usize) {
    out.resize(out.len().next_multiple_of(align), 0);
}

/// Writes streams into an MS-CFB compound file, version 3 with 512-byte sectors.
///
/// The layout follows MS-CFB (Compound File Binary File Format):
///
/// - the header (§2.2) declares mini stream cutoff 4096, the MiniFAT start and
///   count, the DIFAT start and count, and the first 109 FAT sector IDs;
/// - the FAT (§2.3) chains the stream, mini-stream container, MiniFAT and
///   directory sectors, and marks the FAT sectors FATSECT and the DIFAT
///   sectors DIFSECT;
/// - streams of 1 to 4095 bytes go into the mini stream: 64-byte mini sectors
///   chained in the MiniFAT (§2.4), stored in a regular-sector container that
///   the root entry points to;
/// - streams of 4096 bytes or more go into regular sectors;
/// - FAT sector IDs past the 109 header slots go into DIFAT sectors (§2.5), so
///   the file is not capped at the 7,143,424 bytes of sectors that 109 FAT
///   sectors address;
/// - the directory (§2.6) holds the root entry and one entry per stream, in
///   input order, linked as a name-ordered sibling tree.
///
/// # Errors
///
/// Returns `SizeLimit` when a stream exceeds 0x80000000 bytes, the most a
/// version 3 stream may hold (§2.6.1), or when the document would need more
/// sectors than the format or the platform can address.
pub fn write_ole2(streams: &[(&str, &[u8])]) -> Result<Vec<u8>> {
    let lens: Vec<(&str, usize)> = streams
        .iter()
        .map(|&(name, data)| (name, data.len()))
        .collect();
    let layout = Layout::plan(&lens)?;
    let mut out = Vec::with_capacity(layout.file_len);
    out.extend_from_slice(&header(&layout));

    // Regular stream sectors, each stream padded to a whole sector.
    for (&(_, data), plan) in streams.iter().zip(&layout.streams) {
        if let Placement::Regular { .. } = plan.placement {
            out.extend_from_slice(data);
            pad_to(&mut out, SECTOR_SIZE);
        }
    }

    // Mini-stream container: each small stream padded to a whole mini sector.
    // Every region starts sector-aligned, so file offsets align mini sectors too.
    for (&(_, data), plan) in streams.iter().zip(&layout.streams) {
        if let Placement::Mini { .. } = plan.placement {
            out.extend_from_slice(data);
            pad_to(&mut out, MINI_SECTOR_SIZE);
        }
    }
    pad_to(&mut out, SECTOR_SIZE);

    // MiniFAT.
    let mut minifat = vec![FREESECT; layout.minifat as usize * ENTRIES_PER_SECTOR as usize];
    for plan in &layout.streams {
        if let Placement::Mini { start, count } = plan.placement {
            link_chain(&mut minifat, start, count);
        }
    }
    push_u32s(&mut out, &minifat);

    // Directory: the root entry (stream ID 0), then one entry per stream.
    let mut entries = Vec::with_capacity(streams.len() + 1);
    entries.push(WriteDirEntry {
        name: "Root Entry",
        entry_type: DIR_TYPE_ROOT,
        left_sibling_id: NOSTREAM,
        right_sibling_id: NOSTREAM,
        child_id: NOSTREAM,
        // The mini-stream container, sized in whole mini sectors (MS-CFB §2.6.2).
        start_sector: if layout.container > 0 {
            layout.container_start()
        } else {
            ENDOFCHAIN
        },
        size: layout.mini_stream_size,
    });
    for (&(name, _), plan) in streams.iter().zip(&layout.streams) {
        entries.push(WriteDirEntry {
            name,
            entry_type: DIR_TYPE_STREAM,
            left_sibling_id: NOSTREAM,
            right_sibling_id: NOSTREAM,
            child_id: NOSTREAM,
            start_sector: match plan.placement {
                Placement::Empty => ENDOFCHAIN,
                Placement::Mini { start, .. } | Placement::Regular { start, .. } => start,
            },
            size: plan.size,
        });
    }
    // Stream IDs 1..=n; Layout::plan bounded n by MAXREGSID.
    let mut sorted: Vec<u32> = (1..=MAXREGSID).take(streams.len()).collect();
    sorted.sort_by(|&a, &b| compare_names(entries[a as usize].name, entries[b as usize].name));
    entries[0].child_id = link_sibling_tree(&mut entries, &sorted);
    for entry in &entries {
        out.extend_from_slice(&entry.to_bytes());
    }
    let unused = unused_dir_entry();
    while !out.len().is_multiple_of(SECTOR_SIZE) {
        out.extend_from_slice(&unused);
    }

    // FAT.
    let mut fat = vec![FREESECT; layout.fat as usize * ENTRIES_PER_SECTOR as usize];
    for plan in &layout.streams {
        if let Placement::Regular { start, count } = plan.placement {
            link_chain(&mut fat, start, count);
        }
    }
    link_chain(&mut fat, layout.container_start(), layout.container);
    link_chain(&mut fat, layout.minifat_start(), layout.minifat);
    link_chain(&mut fat, layout.dir_start(), layout.dir);
    let fat_sectors = layout.fat_start()..layout.difat_start();
    let difat_end = layout.difat_start() + layout.difat;
    let difat_sectors = layout.difat_start()..difat_end;
    for sid in fat_sectors.clone() {
        fat[sid as usize] = FATSECT;
    }
    for sid in difat_sectors.clone() {
        fat[sid as usize] = DIFSECT;
    }
    push_u32s(&mut out, &fat);

    // DIFAT sectors: the FAT sector IDs past the header's 109, 127 per sector
    // padded with FREESECT, then the next DIFAT sector ID (ENDOFCHAIN on the last).
    let mut overflow = fat_sectors.skip(HEADER_DIFAT_ENTRIES as usize);
    for sid in difat_sectors {
        let mut words = [FREESECT; ENTRIES_PER_SECTOR as usize];
        for (slot, fat_sid) in words
            .iter_mut()
            .zip(overflow.by_ref().take(DIFAT_ENTRIES_PER_SECTOR as usize))
        {
            *slot = fat_sid;
        }
        words[DIFAT_ENTRIES_PER_SECTOR as usize] = if sid + 1 < difat_end {
            sid + 1
        } else {
            ENDOFCHAIN
        };
        push_u32s(&mut out, &words);
    }

    debug_assert_eq!(out.len(), layout.file_len);
    Ok(out)
}

/// Builds the version 3 header (MS-CFB §2.2).
fn header(layout: &Layout) -> [u8; SECTOR_SIZE] {
    let mut buf = [0u8; SECTOR_SIZE];
    let mut put_u16 = |offset: usize, value: u16| {
        buf[offset..offset + 2].copy_from_slice(&value.to_le_bytes());
    };
    put_u16(24, 0x003E); // minor version
    put_u16(26, 0x0003); // major version 3
    put_u16(28, 0xFFFE); // byte order: little-endian
    put_u16(30, 9); // sector shift: 2^9 = 512
    put_u16(32, 6); // mini sector shift: 2^6 = 64

    let first_or_end = |count: u32, start: u32| if count > 0 { start } else { ENDOFCHAIN };
    // Number of directory sectors (40..44) stays zero, as version 3 requires,
    // and so does the transaction signature (52..56), as MS-CFB requires of a
    // writer without file transactions.
    let words = [
        (44, layout.fat),
        (48, layout.dir_start()),
        (56, MINI_STREAM_CUTOFF),
        (60, first_or_end(layout.minifat, layout.minifat_start())),
        (64, layout.minifat),
        (68, first_or_end(layout.difat, layout.difat_start())),
    ];
    buf[..8].copy_from_slice(&crate::ole2::OLE2_MAGIC);
    // CLSID (8..24) and reserved (34..40) stay zero.
    for (offset, value) in words {
        buf[offset..offset + 4].copy_from_slice(&value.to_le_bytes());
    }
    buf[72..76].copy_from_slice(&layout.difat.to_le_bytes());

    // Header DIFAT: the first 109 FAT sector IDs, the rest FREESECT.
    let mut fat_sids = layout.fat_start()..layout.difat_start();
    for slot in buf[76..].as_chunks_mut::<4>().0 {
        *slot = fat_sids.next().unwrap_or(FREESECT).to_le_bytes();
    }
    buf
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::ole2::cfb::reader::read_stream;
    use crate::ole2::detect::{FileFormat, Ole2Kind, classify_ole2, detect_format};

    #[test]
    fn test_write_ole2_roundtrip() {
        let stream1_data = b"Hello, OLE2 world!";
        let stream2_data = b"Second stream content here.";
        let streams: Vec<(&str, &[u8])> = vec![
            ("EncryptionInfo", stream1_data),
            ("EncryptedPackage", stream2_data),
        ];

        let ole2_bytes = write_ole2(&streams).unwrap();

        // Verify it's detected as OLE2.
        assert_eq!(detect_format(&ole2_bytes), FileFormat::Ole2);

        // Verify it classifies correctly (has EncryptionInfo + EncryptedPackage).
        assert_eq!(classify_ole2(&ole2_bytes).unwrap(), Ole2Kind::EncryptedXlsx);

        // Read back each stream and verify content.
        let read1 = read_stream(&ole2_bytes, "EncryptionInfo").unwrap();
        assert_eq!(read1, stream1_data);

        let read2 = read_stream(&ole2_bytes, "EncryptedPackage").unwrap();
        assert_eq!(read2, stream2_data);
    }

    #[test]
    fn test_write_ole2_empty_stream() {
        let streams: Vec<(&str, &[u8])> = vec![("EmptyStream", b"")];

        let ole2_bytes = write_ole2(&streams).unwrap();

        // Must be valid OLE2.
        assert_eq!(detect_format(&ole2_bytes), FileFormat::Ole2);

        // Classification: unknown (no EncryptionInfo/Workbook).
        assert_eq!(classify_ole2(&ole2_bytes).unwrap(), Ole2Kind::Unknown);

        // Read back the empty stream — size should be 0.
        let read_back = read_stream(&ole2_bytes, "EmptyStream").unwrap();
        assert!(read_back.is_empty());
    }

    #[test]
    fn test_write_ole2_large_stream() {
        // Create a stream larger than one sector (> 512 bytes).
        let large_data: Vec<u8> = (0..2000u32).map(|i| (i % 256) as u8).collect();
        let streams: Vec<(&str, &[u8])> = vec![("LargeStream", &large_data)];

        let ole2_bytes = write_ole2(&streams).unwrap();

        assert_eq!(detect_format(&ole2_bytes), FileFormat::Ole2);

        // Roundtrip: verify every byte matches.
        let read_back = read_stream(&ole2_bytes, "LargeStream").unwrap();
        assert_eq!(read_back.len(), large_data.len());
        assert_eq!(read_back, large_data);
    }

    #[test]
    fn test_ole2_sector_chain_integrity() {
        // Several streams under 4096 bytes, to test MiniFAT chain correctness.
        let data_a: Vec<u8> = vec![0xAA; 600]; // 10 mini sectors
        let data_b: Vec<u8> = vec![0xBB; 100]; // 2 mini sectors
        let data_c: Vec<u8> = vec![0xCC; 1500]; // 24 mini sectors

        let streams: Vec<(&str, &[u8])> = vec![
            ("StreamA", &data_a),
            ("StreamB", &data_b),
            ("StreamC", &data_c),
        ];

        let ole2_bytes = write_ole2(&streams).unwrap();

        // Read each stream back via read_stream (which follows MiniFAT chains).
        let read_a = read_stream(&ole2_bytes, "StreamA").unwrap();
        assert_eq!(read_a.len(), 600);
        assert_eq!(read_a, data_a);

        let read_b = read_stream(&ole2_bytes, "StreamB").unwrap();
        assert_eq!(read_b.len(), 100);
        assert_eq!(read_b, data_b);

        let read_c = read_stream(&ole2_bytes, "StreamC").unwrap();
        assert_eq!(read_c.len(), 1500);
        assert_eq!(read_c, data_c);
    }

    #[test]
    fn test_ole2_directory_utf16() {
        // Stream name with non-ASCII characters.
        let name = "Data\u{2122}"; // "Data(TM)" — trademark symbol
        let content = b"Unicode name test";
        let streams: Vec<(&str, &[u8])> = vec![(name, content.as_slice())];

        let ole2_bytes = write_ole2(&streams).unwrap();

        assert_eq!(detect_format(&ole2_bytes), FileFormat::Ole2);

        // Roundtrip: reader should decode UTF-16LE name correctly.
        let read_back = read_stream(&ole2_bytes, name).unwrap();
        assert_eq!(read_back, content);
    }

    // -----------------------------------------------------------------------
    // MS-CFB conformance: the independent `cfb` crate reads what we write
    // -----------------------------------------------------------------------

    use std::io::{Cursor, Read};

    const ENDOFCHAIN_U32: u32 = 0xFFFF_FFFE;

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

    /// Reads `name` through the `cfb` dev-dependency opened in strict mode,
    /// which rejects MS-CFB violations in the FAT, MiniFAT, DIFAT and
    /// directory tree.
    fn cfb_read(ole: &[u8], name: &str) -> Vec<u8> {
        let mut comp = cfb::CompoundFile::open_strict(Cursor::new(ole)).expect("cfb opens");
        let mut out = Vec::new();
        comp.open_stream(format!("/{name}"))
            .expect("cfb finds stream")
            .read_to_end(&mut out)
            .expect("cfb reads stream");
        out
    }

    #[test]
    fn difat_sectors_carry_an_8_mib_stream() {
        let big = pattern(8 * 1024 * 1024, 1);
        let ole = write_ole2(&[("Big", &big)]).unwrap();
        assert!(u32_at(&ole, 44) > 109, "FAT must outgrow the header DIFAT");
        assert_ne!(u32_at(&ole, 68), ENDOFCHAIN_U32, "first DIFAT sector");
        assert!(u32_at(&ole, 72) >= 1, "DIFAT sector count");
        assert!(
            read_stream(&ole, "Big").unwrap() == big,
            "read_stream differs"
        );
        assert!(cfb_read(&ole, "Big") == big, "cfb read differs");
    }

    #[test]
    fn small_streams_go_to_the_mini_stream() {
        let info = [7u8; 1024];
        let package = pattern(5000, 2);
        let ole = write_ole2(&[("EncryptionInfo", &info), ("EncryptedPackage", &package)]).unwrap();
        assert_ne!(u32_at(&ole, 60), ENDOFCHAIN_U32, "first MiniFAT sector");
        assert!(u32_at(&ole, 64) >= 1, "MiniFAT sector count");
        assert_eq!(read_stream(&ole, "EncryptionInfo").unwrap(), info);
        assert_eq!(read_stream(&ole, "EncryptedPackage").unwrap(), package);
        assert_eq!(cfb_read(&ole, "EncryptionInfo"), info);
        assert_eq!(cfb_read(&ole, "EncryptedPackage"), package);

        // The root entry (directory entry 0) holds the mini-stream container:
        // 16 mini sectors of 64 bytes.
        let dir = (u32_at(&ole, 48) as usize + 1) * 512;
        assert_ne!(u32_at(&ole, dir + 116), ENDOFCHAIN_U32, "root start sector");
        assert_eq!(u32_at(&ole, dir + 120), 16 * 64, "root size");
    }

    #[test]
    fn zero_length_stream_round_trips_empty() {
        let small = pattern(100, 3);
        let ole = write_ole2(&[("Empty", b""), ("Small", &small)]).unwrap();
        // Directory entry 1 is the empty stream: start ENDOFCHAIN, size 0.
        let entry = (u32_at(&ole, 48) as usize + 1) * 512 + 128;
        assert_eq!(u32_at(&ole, entry + 116), ENDOFCHAIN_U32);
        assert_eq!(u32_at(&ole, entry + 120), 0);
        assert!(read_stream(&ole, "Empty").unwrap().is_empty());
        assert!(cfb_read(&ole, "Empty").is_empty());
        assert_eq!(cfb_read(&ole, "Small"), small);
    }

    #[test]
    fn streams_in_any_name_order_open_in_cfb() {
        let streams: Vec<(String, Vec<u8>)> = ["Zeta", "Beta", "Alpha", "Longer name", "Mid", "a"]
            .iter()
            .zip(1u32..)
            .map(|(name, seed)| ((*name).to_owned(), pattern(300 * seed as usize, seed)))
            .collect();
        let refs: Vec<(&str, &[u8])> = streams
            .iter()
            .map(|(n, d)| (n.as_str(), d.as_slice()))
            .collect();
        let ole = write_ole2(&refs).unwrap();
        for (name, data) in &streams {
            assert_eq!(&cfb_read(&ole, name), data, "{name}");
            assert_eq!(&read_stream(&ole, name).unwrap(), data, "{name}");
        }
    }

    #[test]
    fn stream_over_2_gib_is_a_size_limit() {
        let err = Layout::plan(&[("Small", 10), ("Huge", 0x8000_0001)]).unwrap_err();
        assert_eq!(err.code(), "SIZE_LIMIT", "{err}");
        assert!(
            err.to_string().contains(
                "OLE2 stream 'Huge' is 2147483649 bytes; MS-CFB v3 allows at most 2147483648 bytes per stream"
            ),
            "{err}"
        );
    }

    #[test]
    fn stream_of_exactly_2_gib_is_planned() {
        let layout = Layout::plan(&[("Max", 0x8000_0000)]).unwrap();
        assert_eq!(layout.regular, 0x8000_0000 / 512);
        assert!(layout.fat > 109 && layout.difat >= 1);
        // The FAT covers every sector, including its own and the DIFAT sectors.
        let total = layout.difat_start() + layout.difat;
        assert!(layout.fat * 128 >= total);
        assert_eq!((layout.fat - 109).div_ceil(127), layout.difat);
    }

    #[test]
    fn existing_layouts_open_in_cfb() {
        let a = vec![0xAA; 600];
        let b = vec![0xBB; 100];
        let c = vec![0xCC; 1500];
        let ole = write_ole2(&[("StreamA", &a), ("StreamB", &b), ("StreamC", &c)]).unwrap();
        assert_eq!(cfb_read(&ole, "StreamA"), a);
        assert_eq!(cfb_read(&ole, "StreamB"), b);
        assert_eq!(cfb_read(&ole, "StreamC"), c);
        let empty = write_ole2(&[]).unwrap();
        cfb::CompoundFile::open_strict(Cursor::new(&empty)).expect("cfb opens empty file");
    }
}
