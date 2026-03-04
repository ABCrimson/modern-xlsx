use crate::errors::{ModernXlsxError, Result};

/// OLE2 Compound Document magic bytes.
const OLE2_MAGIC: [u8; 8] = [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1];

/// Sector size for v3 format (2^9 = 512 bytes).
const SECTOR_SIZE: usize = 512;

/// Number of FAT entries per sector (512 / 4 = 128).
const FAT_ENTRIES_PER_SECTOR: usize = SECTOR_SIZE / 4;

/// Number of directory entries per sector (512 / 128 = 4).
const DIR_ENTRIES_PER_SECTOR: usize = SECTOR_SIZE / 128;

/// Maximum FAT sector IDs in header DIFAT array.
const MAX_HEADER_DIFAT: usize = 109;

/// FAT special values.
const FAT_END_OF_CHAIN: u32 = 0xFFFFFFFE;
const FAT_FREE_SECTOR: u32 = 0xFFFFFFFF;
const FAT_FAT_SECTOR: u32 = 0xFFFFFFFD;

/// OLE2 directory entry types.
const DIR_TYPE_ROOT: u8 = 5;
const DIR_TYPE_STREAM: u8 = 2;

/// A directory entry for writing.
struct WriteDirEntry {
    name: String,
    entry_type: u8,
    child_id: u32,
    right_sibling_id: u32,
    start_sector: u32,
    size: u32,
}

impl WriteDirEntry {
    /// Create a root storage entry.
    fn root(child_id: u32) -> Self {
        Self {
            name: "Root Entry".into(),
            entry_type: DIR_TYPE_ROOT,
            child_id,
            right_sibling_id: FAT_FREE_SECTOR,
            start_sector: FAT_END_OF_CHAIN,
            size: 0,
        }
    }

    /// Create a stream entry.
    fn stream(name: &str, start_sector: u32, size: u32, right_sibling_id: u32) -> Self {
        Self {
            name: name.into(),
            entry_type: DIR_TYPE_STREAM,
            child_id: FAT_FREE_SECTOR,
            right_sibling_id,
            start_sector,
            size,
        }
    }

    /// Serialize this directory entry to 128 bytes.
    fn to_bytes(&self) -> [u8; 128] {
        let mut buf = [0u8; 128];

        // Name as UTF-16LE (max 31 chars + null = 64 bytes)
        let name_utf16: Vec<u16> = self.name.encode_utf16().collect();
        let name_chars = name_utf16.len().min(31);
        for (i, &c) in name_utf16[..name_chars].iter().enumerate() {
            let offset = i * 2;
            buf[offset] = c as u8;
            buf[offset + 1] = (c >> 8) as u8;
        }
        // Null terminator
        let null_pos = name_chars * 2;
        if null_pos + 1 < 64 {
            buf[null_pos] = 0;
            buf[null_pos + 1] = 0;
        }

        // Name length in bytes (including null terminator)
        let name_byte_len = ((name_chars + 1) * 2) as u16;
        buf[64..66].copy_from_slice(&name_byte_len.to_le_bytes());

        // Entry type
        buf[66] = self.entry_type;

        // Color: black (1)
        buf[67] = 1;

        // Left sibling: none
        buf[68..72].copy_from_slice(&FAT_FREE_SECTOR.to_le_bytes());

        // Right sibling
        buf[72..76].copy_from_slice(&self.right_sibling_id.to_le_bytes());

        // Child ID
        buf[76..80].copy_from_slice(&self.child_id.to_le_bytes());

        // CLSID (16 bytes at 80..96): all zeros
        // State bits (96..100): zero
        // Creation time (100..108): zero
        // Modified time (108..116): zero

        // Start sector
        buf[116..120].copy_from_slice(&self.start_sector.to_le_bytes());

        // Size (v3: 4-byte size)
        buf[120..124].copy_from_slice(&self.size.to_le_bytes());

        buf
    }
}

/// Generates an OLE2 compound document containing the given streams.
///
/// Each entry in `streams` is a `(name, data)` pair. The output is a valid
/// OLE2 v3 compound document (512-byte sectors) that can be read back by
/// our existing OLE2 reader.
///
/// # Errors
///
/// Returns an error if the resulting document would require more than 109
/// FAT sectors (DIFAT chains are not implemented).
pub fn write_ole2(streams: &[(&str, &[u8])]) -> Result<Vec<u8>> {
    // Step 1: Calculate data sectors needed per stream.
    let mut stream_sector_counts: Vec<usize> = Vec::with_capacity(streams.len());
    let mut total_data_sectors: usize = 0;
    for &(_, data) in streams {
        if data.is_empty() {
            stream_sector_counts.push(0);
        } else {
            let sectors = data.len().div_ceil(SECTOR_SIZE);
            stream_sector_counts.push(sectors);
            total_data_sectors += sectors;
        }
    }

    // Step 2: Calculate directory sectors needed.
    // Directory entries: 1 root + N stream entries.
    let total_dir_entries = 1 + streams.len();
    let dir_sectors = total_dir_entries.div_ceil(DIR_ENTRIES_PER_SECTOR);

    // Step 3: Calculate FAT sectors needed.
    // Total sectors in file = data + directory + FAT.
    // The FAT must account for all sectors including itself, so we iterate:
    //   total = data + dir + fat
    //   fat_needed = ceil(total / 128)
    // We solve this iteratively since adding FAT sectors increases the total.
    let mut fat_sectors = 1usize;
    loop {
        let total = total_data_sectors + dir_sectors + fat_sectors;
        let needed = total.div_ceil(FAT_ENTRIES_PER_SECTOR);
        if needed <= fat_sectors {
            break;
        }
        fat_sectors = needed;
    }

    if fat_sectors > MAX_HEADER_DIFAT {
        return Err(ModernXlsxError::ZipWrite(format!(
            "OLE2 document too large: requires {fat_sectors} FAT sectors (max {MAX_HEADER_DIFAT})"
        )));
    }

    let total_sectors = total_data_sectors + dir_sectors + fat_sectors;
    let file_size = SECTOR_SIZE + total_sectors * SECTOR_SIZE; // header + sectors
    let mut buf = vec![0u8; file_size];

    // Sector ID assignments:
    //   [0 .. total_data_sectors-1]          = data sectors
    //   [total_data_sectors .. +dir_sectors]  = directory sectors
    //   [total_data_sectors+dir_sectors .. ]  = FAT sectors
    let first_dir_sid = total_data_sectors as u32;
    let first_fat_sid = (total_data_sectors + dir_sectors) as u32;

    // Step 4: Write the header (sector 0 = file offset 0..512).
    write_header(
        &mut buf[..SECTOR_SIZE],
        first_dir_sid,
        fat_sectors as u32,
        first_fat_sid,
    );

    // Step 5: Write data sectors.
    let mut current_data_sid: u32 = 0;
    for (i, &(_, data)) in streams.iter().enumerate() {
        let count = stream_sector_counts[i];
        if count == 0 {
            continue;
        }
        for s in 0..count {
            let sector_offset = SECTOR_SIZE + (current_data_sid as usize) * SECTOR_SIZE;
            let data_start = s * SECTOR_SIZE;
            let data_end = (data_start + SECTOR_SIZE).min(data.len());
            let chunk = &data[data_start..data_end];
            buf[sector_offset..sector_offset + chunk.len()].copy_from_slice(chunk);
            // Remaining bytes in sector stay zero-filled (already initialized).
            current_data_sid += 1;
        }
    }

    // Step 6: Build directory entries and write directory sectors.
    let mut dir_entries: Vec<WriteDirEntry> = Vec::with_capacity(total_dir_entries);

    // Root entry: child points to entry 1 (first stream), or NOSTREAM if no streams.
    let root_child = if streams.is_empty() {
        FAT_FREE_SECTOR
    } else {
        1
    };
    dir_entries.push(WriteDirEntry::root(root_child));

    // Stream entries: linked via right sibling IDs (flat list).
    let mut data_sid_cursor: u32 = 0;
    for (i, &(name, data)) in streams.iter().enumerate() {
        let start_sector = if data.is_empty() {
            FAT_END_OF_CHAIN
        } else {
            data_sid_cursor
        };
        let right_sibling = if i + 1 < streams.len() {
            (i + 2) as u32 // directory entry index of next stream
        } else {
            FAT_FREE_SECTOR
        };

        dir_entries.push(WriteDirEntry::stream(
            name,
            start_sector,
            data.len() as u32,
            right_sibling,
        ));

        data_sid_cursor += stream_sector_counts[i] as u32;
    }

    // Write directory entries into directory sectors.
    for (i, entry) in dir_entries.iter().enumerate() {
        let sector_idx = i / DIR_ENTRIES_PER_SECTOR;
        let entry_in_sector = i % DIR_ENTRIES_PER_SECTOR;
        let sector_sid = first_dir_sid as usize + sector_idx;
        let offset = SECTOR_SIZE + sector_sid * SECTOR_SIZE + entry_in_sector * 128;
        let entry_bytes = entry.to_bytes();
        buf[offset..offset + 128].copy_from_slice(&entry_bytes);
    }

    // Pad remaining directory entry slots with empty entries (type 0, already zero-filled).

    // Step 7: Build and write the FAT.
    let mut fat = vec![FAT_FREE_SECTOR; total_sectors];

    // Data sector chains: consecutive sectors for each stream, last one = end-of-chain.
    let mut data_sid: u32 = 0;
    for count in &stream_sector_counts {
        for s in 0..*count {
            let sid = data_sid as usize;
            if s + 1 < *count {
                fat[sid] = data_sid + 1; // next sector in chain
            } else {
                fat[sid] = FAT_END_OF_CHAIN; // end of chain
            }
            data_sid += 1;
        }
    }

    // Directory sector chain: consecutive, last = end-of-chain.
    for d in 0..dir_sectors {
        let sid = first_dir_sid as usize + d;
        if d + 1 < dir_sectors {
            fat[sid] = first_dir_sid + (d as u32) + 1;
        } else {
            fat[sid] = FAT_END_OF_CHAIN;
        }
    }

    // FAT sector entries: marked as FAT_FAT_SECTOR.
    for f in 0..fat_sectors {
        let sid = first_fat_sid as usize + f;
        fat[sid] = FAT_FAT_SECTOR;
    }

    // Write FAT sectors.
    for f in 0..fat_sectors {
        let fat_sector_offset = SECTOR_SIZE + (first_fat_sid as usize + f) * SECTOR_SIZE;
        let entry_start = f * FAT_ENTRIES_PER_SECTOR;
        for e in 0..FAT_ENTRIES_PER_SECTOR {
            let fat_idx = entry_start + e;
            let value = if fat_idx < fat.len() {
                fat[fat_idx]
            } else {
                FAT_FREE_SECTOR
            };
            let offset = fat_sector_offset + e * 4;
            buf[offset..offset + 4].copy_from_slice(&value.to_le_bytes());
        }
    }

    Ok(buf)
}

/// Writes the OLE2 v3 header into the first 512 bytes.
fn write_header(buf: &mut [u8], first_dir_sid: u32, fat_sector_count: u32, first_fat_sid: u32) {
    // Magic signature
    buf[..8].copy_from_slice(&OLE2_MAGIC);

    // CLSID (8..24): all zeros

    // Minor version: 62
    buf[24..26].copy_from_slice(&62u16.to_le_bytes());

    // Major version: 3 (v3 format, 512-byte sectors)
    buf[26..28].copy_from_slice(&3u16.to_le_bytes());

    // Byte order: 0xFFFE (little-endian)
    buf[28..30].copy_from_slice(&0xFFFEu16.to_le_bytes());

    // Sector size shift: 9 (2^9 = 512)
    buf[30..32].copy_from_slice(&9u16.to_le_bytes());

    // Mini sector size shift: 6 (2^6 = 64)
    buf[32..34].copy_from_slice(&6u16.to_le_bytes());

    // Reserved (34..40): zeros

    // Total directory sectors: 0 (must be 0 for v3)
    buf[40..44].copy_from_slice(&0u32.to_le_bytes());

    // Number of FAT sectors
    buf[44..48].copy_from_slice(&fat_sector_count.to_le_bytes());

    // First directory sector SID
    buf[48..52].copy_from_slice(&first_dir_sid.to_le_bytes());

    // Transaction signature number (52..56): 0
    // Mini stream cutoff size (56..60): 0x00001000 (4096)
    buf[56..60].copy_from_slice(&0x1000u32.to_le_bytes());

    // First mini FAT sector SID: none
    buf[60..64].copy_from_slice(&FAT_END_OF_CHAIN.to_le_bytes());

    // Number of mini FAT sectors: 0
    buf[64..68].copy_from_slice(&0u32.to_le_bytes());

    // First DIFAT sector: none
    buf[68..72].copy_from_slice(&FAT_END_OF_CHAIN.to_le_bytes());

    // Number of DIFAT sectors: 0
    buf[72..76].copy_from_slice(&0u32.to_le_bytes());

    // DIFAT array (109 entries at offsets 76..512)
    for i in 0..MAX_HEADER_DIFAT {
        let offset = 76 + i * 4;
        if (i as u32) < fat_sector_count {
            let sid = first_fat_sid + i as u32;
            buf[offset..offset + 4].copy_from_slice(&sid.to_le_bytes());
        } else {
            buf[offset..offset + 4].copy_from_slice(&FAT_FREE_SECTOR.to_le_bytes());
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::ole2::detect::{classify_ole2, detect_format, read_stream, FileFormat, Ole2Kind};

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
        assert_eq!(
            classify_ole2(&ole2_bytes).unwrap(),
            Ole2Kind::EncryptedXlsx
        );

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
        // Multiple streams of varying sizes to test FAT chain correctness.
        let data_a: Vec<u8> = vec![0xAA; 600]; // 2 sectors
        let data_b: Vec<u8> = vec![0xBB; 100]; // 1 sector
        let data_c: Vec<u8> = vec![0xCC; 1500]; // 3 sectors

        let streams: Vec<(&str, &[u8])> = vec![
            ("StreamA", &data_a),
            ("StreamB", &data_b),
            ("StreamC", &data_c),
        ];

        let ole2_bytes = write_ole2(&streams).unwrap();

        // Read each stream back via read_stream (which follows FAT chains).
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
}
