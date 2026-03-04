use crate::errors::ModernXlsxError;

type Result<T> = std::result::Result<T, ModernXlsxError>;

/// OLE2 Compound Document magic bytes.
const OLE2_MAGIC: [u8; 8] = [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1];

/// ZIP archive magic bytes (PK\x03\x04).
const ZIP_MAGIC: [u8; 4] = [0x50, 0x4B, 0x03, 0x04];

// Error message constants shared between reader.rs and streaming.rs.
pub const ERR_ENCRYPTED: &str = "This file is password-protected (OLE2 compound document). \
    Provide password via readBuffer(data, { password: '...' }).";
pub const ERR_LEGACY_XLS: &str = "Legacy .xls format not supported. Convert to .xlsx first.";
pub const ERR_OLE2_UNKNOWN: &str = "Unrecognized OLE2 compound document.";
pub const ERR_NOT_XLSX: &str = "Not a valid XLSX file (expected ZIP or OLE2 header).";

/// Detected file format.
#[derive(Debug, Clone, PartialEq)]
pub enum FileFormat {
    Zip,
    Ole2,
    Unknown,
}

/// Classification of an OLE2 compound document.
#[derive(Debug, Clone, PartialEq)]
pub enum Ole2Kind {
    EncryptedXlsx,
    LegacyXls,
    Unknown,
}

/// Detect whether a byte slice begins with ZIP or OLE2 magic bytes.
pub fn detect_format(data: &[u8]) -> FileFormat {
    if data.len() >= 8 && data[..8] == OLE2_MAGIC {
        FileFormat::Ole2
    } else if data.len() >= 4 && data[..4] == ZIP_MAGIC {
        FileFormat::Zip
    } else {
        FileFormat::Unknown
    }
}

/// Minimal OLE2 header parsing — just enough to read directory entries.
struct Ole2Header {
    sector_size: usize,
    first_dir_sector: u32,
    fat_sectors: Vec<u32>,
}

impl Ole2Header {
    fn parse(data: &[u8]) -> Result<Self> {
        if data.len() < 512 {
            return Err(ModernXlsxError::UnrecognizedFormat(
                "File too small for OLE2 header".into(),
            ));
        }
        let major_version = u16::from_le_bytes([data[26], data[27]]);
        let sector_size = if major_version == 4 { 4096 } else { 512 };
        let first_dir_sector = u32::from_le_bytes(data[48..52].try_into().unwrap());

        // Read FAT sector IDs from header (109 entries starting at offset 76)
        let num_fat_sectors = u32::from_le_bytes(data[44..48].try_into().unwrap()) as usize;
        let mut fat_sectors = Vec::with_capacity(num_fat_sectors.min(109));
        for i in 0..num_fat_sectors.min(109) {
            let offset = 76 + i * 4;
            if offset + 4 > 512 {
                break;
            }
            let sid = u32::from_le_bytes(data[offset..offset + 4].try_into().unwrap());
            if sid != 0xFFFFFFFE && sid != 0xFFFFFFFF {
                fat_sectors.push(sid);
            }
        }

        Ok(Self {
            sector_size,
            first_dir_sector,
            fat_sectors,
        })
    }

    /// Returns the byte offset for a given sector ID.
    fn sector_offset(&self, sid: u32) -> usize {
        512 + (sid as usize) * self.sector_size
    }
}

/// Reads the FAT (File Allocation Table) to follow sector chains.
fn read_fat(data: &[u8], header: &Ole2Header) -> Vec<u32> {
    let entries_per_sector = header.sector_size / 4;
    let mut fat = Vec::new();
    for &sid in &header.fat_sectors {
        let offset = header.sector_offset(sid);
        for i in 0..entries_per_sector {
            let pos = offset + i * 4;
            if pos + 4 <= data.len() {
                fat.push(u32::from_le_bytes(data[pos..pos + 4].try_into().unwrap()));
            }
        }
    }
    fat
}

/// Follows a FAT chain from a starting sector, collecting all sectors.
fn follow_chain(fat: &[u32], start: u32) -> Vec<u32> {
    let mut chain = Vec::new();
    let mut current = start;
    // Safety limit to avoid infinite loops on corrupt files
    let max_sectors = fat.len();
    while (current as usize) < fat.len() && chain.len() < max_sectors {
        chain.push(current);
        let next = fat[current as usize];
        if next == 0xFFFFFFFE || next == 0xFFFFFFFF {
            break;
        } // end of chain
        current = next;
    }
    chain
}

/// Reads bytes of a stream from the OLE2 file given a sector chain.
fn read_sectors(data: &[u8], header: &Ole2Header, chain: &[u32], size: usize) -> Vec<u8> {
    let mut result = Vec::with_capacity(size);
    for &sid in chain {
        let offset = header.sector_offset(sid);
        let end = (offset + header.sector_size).min(data.len());
        if offset < data.len() {
            result.extend_from_slice(&data[offset..end]);
        }
    }
    result.truncate(size);
    result
}

/// Directory entry from OLE2 compound document.
struct DirEntry {
    name: String,
    #[allow(dead_code)]
    entry_type: u8,
    start_sector: u32,
    size: u64,
}

/// Reads directory entries from the directory sector chain.
fn read_directory(data: &[u8], header: &Ole2Header, fat: &[u32]) -> Vec<DirEntry> {
    let chain = follow_chain(fat, header.first_dir_sector);
    let dir_bytes = read_sectors(data, header, &chain, chain.len() * header.sector_size);

    let mut entries = Vec::new();
    for chunk in dir_bytes.chunks(128) {
        if chunk.len() < 128 {
            break;
        }
        let entry_type = chunk[66];
        if entry_type == 0 {
            continue;
        } // empty entry

        // Read name: UTF-16LE, length in bytes at offset 64-65
        let name_len = u16::from_le_bytes([chunk[64], chunk[65]]) as usize;
        let name_bytes = name_len.saturating_sub(2); // subtract null terminator
        let name: String = (0..name_bytes / 2)
            .map(|i| u16::from_le_bytes([chunk[i * 2], chunk[i * 2 + 1]]))
            .collect::<Vec<u16>>()
            .iter()
            .copied()
            .map(|c| char::from_u32(u32::from(c)).unwrap_or('\u{FFFD}'))
            .collect();

        let start_sector = u32::from_le_bytes(chunk[116..120].try_into().unwrap());
        let size = if header.sector_size == 4096 {
            u64::from_le_bytes(chunk[120..128].try_into().unwrap())
        } else {
            u32::from_le_bytes(chunk[120..124].try_into().unwrap()) as u64
        };

        entries.push(DirEntry {
            name,
            entry_type,
            start_sector,
            size,
        });
    }
    entries
}

/// Classifies an OLE2 compound document.
pub fn classify_ole2(data: &[u8]) -> Result<Ole2Kind> {
    let header = Ole2Header::parse(data)?;
    let fat = read_fat(data, &header);
    let entries = read_directory(data, &header, &fat);

    let has_encryption_info = entries.iter().any(|e| e.name == "EncryptionInfo");
    let has_encrypted_package = entries.iter().any(|e| e.name == "EncryptedPackage");
    let has_workbook = entries.iter().any(|e| e.name == "Workbook" || e.name == "Book");

    if has_encryption_info && has_encrypted_package {
        Ok(Ole2Kind::EncryptedXlsx)
    } else if has_workbook {
        Ok(Ole2Kind::LegacyXls)
    } else {
        Ok(Ole2Kind::Unknown)
    }
}

/// Reads a named stream from the OLE2 document.
pub fn read_stream(data: &[u8], name: &str) -> Result<Vec<u8>> {
    let header = Ole2Header::parse(data)?;
    let fat = read_fat(data, &header);
    let entries = read_directory(data, &header, &fat);

    let entry = entries.iter().find(|e| e.name == name).ok_or_else(|| {
        ModernXlsxError::MissingPart(format!("OLE2 stream '{name}' not found"))
    })?;

    let chain = follow_chain(&fat, entry.start_sector);
    Ok(read_sectors(data, &header, &chain, entry.size as usize))
}

/// Decrypts an OLE2-wrapped encrypted XLSX file.
/// Returns the decrypted ZIP bytes ready for normal reading.
pub fn decrypt_file(data: &[u8], password: &str) -> Result<Vec<u8>> {
    let enc_info_bytes = read_stream(data, "EncryptionInfo")?;
    let enc_info = super::encryption_info::EncryptionInfo::parse(&enc_info_bytes)?;

    match enc_info {
        super::encryption_info::EncryptionInfo::Agile(ref agile) => {
            let data_key = super::crypto::verify_password_agile(password, agile)?;
            let encrypted_package = read_stream(data, "EncryptedPackage")?;
            super::crypto::verify_hmac(&data_key, agile, &encrypted_package)?;
            super::crypto::decrypt_package(&data_key, agile, &encrypted_package)
        }
        super::encryption_info::EncryptionInfo::Standard(ref std_info) => {
            let data_key = super::crypto::verify_password_standard(password, std_info)?;
            let encrypted_package = read_stream(data, "EncryptedPackage")?;
            super::crypto::decrypt_standard_package(&data_key, std_info, &encrypted_package)
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_detect_zip_format() {
        let data = [0x50, 0x4B, 0x03, 0x04, 0x00, 0x00, 0x00, 0x00];
        assert_eq!(detect_format(&data), FileFormat::Zip);
    }

    #[test]
    fn test_detect_ole2_format() {
        let mut data = vec![0u8; 512];
        data[..8].copy_from_slice(&OLE2_MAGIC);
        assert_eq!(detect_format(&data), FileFormat::Ole2);
    }

    #[test]
    fn test_detect_unknown_format() {
        let data = [0x00, 0x01, 0x02, 0x03, 0x04, 0x05, 0x06, 0x07];
        assert_eq!(detect_format(&data), FileFormat::Unknown);
    }

    #[test]
    fn test_detect_empty_input() {
        assert_eq!(detect_format(&[]), FileFormat::Unknown);
        assert_eq!(detect_format(&[0x50]), FileFormat::Unknown);
    }

    #[test]
    fn test_encrypted_xlsx_classification() {
        // Build a minimal valid OLE2 with EncryptionInfo + EncryptedPackage directory entries
        let ole2 = build_test_ole2(&["EncryptionInfo", "EncryptedPackage"]);
        assert_eq!(classify_ole2(&ole2).unwrap(), Ole2Kind::EncryptedXlsx);
    }

    #[test]
    fn test_legacy_xls_classification() {
        let ole2 = build_test_ole2(&["Workbook"]);
        assert_eq!(classify_ole2(&ole2).unwrap(), Ole2Kind::LegacyXls);
    }

    #[test]
    fn test_legacy_xls_book_stream() {
        let ole2 = build_test_ole2(&["Book"]);
        assert_eq!(classify_ole2(&ole2).unwrap(), Ole2Kind::LegacyXls);
    }

    #[test]
    fn test_unknown_ole2_classification() {
        let ole2 = build_test_ole2(&["SomeRandomStream"]);
        assert_eq!(classify_ole2(&ole2).unwrap(), Ole2Kind::Unknown);
    }

    /// Builds a minimal OLE2 compound document with the given stream names
    /// in the directory. The streams have empty content (size 0).
    fn build_test_ole2(stream_names: &[&str]) -> Vec<u8> {
        let sector_size: usize = 512;
        // Layout:
        //   Offset 0: Header (512 bytes)
        //   Sector 0 (offset 512): Directory sector (contains root entry + stream entries)
        //   Sector 1 (offset 1024): FAT sector
        let total_sectors = 2; // directory + FAT
        let file_size = 512 + total_sectors * sector_size;
        let mut buf = vec![0u8; file_size];

        // --- Header ---
        // Magic
        buf[..8].copy_from_slice(&OLE2_MAGIC);
        // Minor version
        buf[24..26].copy_from_slice(&62u16.to_le_bytes());
        // Major version (3 = 512-byte sectors)
        buf[26..28].copy_from_slice(&3u16.to_le_bytes());
        // Byte order: 0xFFFE (little-endian)
        buf[28..30].copy_from_slice(&0xFFFEu16.to_le_bytes());
        // Sector size shift: 9 (2^9 = 512)
        buf[30..32].copy_from_slice(&9u16.to_le_bytes());
        // Mini sector size shift: 6 (2^6 = 64)
        buf[32..34].copy_from_slice(&6u16.to_le_bytes());
        // Total directory sectors (0 for v3)
        buf[40..44].copy_from_slice(&0u32.to_le_bytes());
        // Number of FAT sectors: 1
        buf[44..48].copy_from_slice(&1u32.to_le_bytes());
        // First directory sector SID: 0
        buf[48..52].copy_from_slice(&0u32.to_le_bytes());
        // First mini FAT sector SID: 0xFFFFFFFE (none)
        buf[60..64].copy_from_slice(&0xFFFFFFFEu32.to_le_bytes());
        // Number of mini FAT sectors: 0
        buf[64..68].copy_from_slice(&0u32.to_le_bytes());
        // First DIFAT sector: 0xFFFFFFFE (none)
        buf[68..72].copy_from_slice(&0xFFFFFFFEu32.to_le_bytes());
        // Number of DIFAT sectors: 0
        buf[72..76].copy_from_slice(&0u32.to_le_bytes());
        // DIFAT array: first entry is FAT sector SID = 1
        buf[76..80].copy_from_slice(&1u32.to_le_bytes());
        // Rest of DIFAT: 0xFFFFFFFF (free)
        for i in 1..109 {
            let offset = 76 + i * 4;
            buf[offset..offset + 4].copy_from_slice(&0xFFFFFFFFu32.to_le_bytes());
        }

        // --- Sector 0 (offset 512): Directory ---
        let dir_offset = 512;

        // Root Entry (entry 0)
        write_dir_entry(
            &mut buf[dir_offset..dir_offset + 128],
            "Root Entry",
            5,
            0xFFFFFFFE,
            0,
        );
        // Set root's child ID to 1 (first stream entry) if we have entries
        if !stream_names.is_empty() {
            buf[dir_offset + 76..dir_offset + 80].copy_from_slice(&1u32.to_le_bytes());
        }

        // Stream entries
        for (i, name) in stream_names.iter().enumerate() {
            let entry_offset = dir_offset + (i + 1) * 128;
            if entry_offset + 128 > buf.len() {
                break;
            }
            write_dir_entry(
                &mut buf[entry_offset..entry_offset + 128],
                name,
                2,
                0xFFFFFFFE,
                0,
            );

            // Set right sibling for red-black tree navigation
            let next_id = if i + 1 < stream_names.len() {
                (i + 2) as u32
            } else {
                0xFFFFFFFFu32
            };
            buf[entry_offset + 72..entry_offset + 76].copy_from_slice(&next_id.to_le_bytes());
        }

        // --- Sector 1 (offset 1024): FAT ---
        let fat_offset = 1024;
        // Sector 0 (directory): end of chain
        buf[fat_offset..fat_offset + 4].copy_from_slice(&0xFFFFFFFEu32.to_le_bytes());
        // Sector 1 (FAT): special FAT sector marker
        buf[fat_offset + 4..fat_offset + 8].copy_from_slice(&0xFFFFFFFDu32.to_le_bytes());
        // Rest: free
        for i in 2..(sector_size / 4) {
            let offset = fat_offset + i * 4;
            buf[offset..offset + 4].copy_from_slice(&0xFFFFFFFFu32.to_le_bytes());
        }

        buf
    }

    fn write_dir_entry(buf: &mut [u8], name: &str, entry_type: u8, start_sector: u32, size: u32) {
        // Name as UTF-16LE
        let name_utf16: Vec<u16> = name.encode_utf16().collect();
        for (i, &c) in name_utf16.iter().enumerate() {
            if i * 2 + 1 >= 64 {
                break;
            }
            buf[i * 2] = c as u8;
            buf[i * 2 + 1] = (c >> 8) as u8;
        }
        // Null terminator
        let null_pos = name_utf16.len() * 2;
        if null_pos + 1 < 64 {
            buf[null_pos] = 0;
            buf[null_pos + 1] = 0;
        }
        // Name length (including null, in bytes)
        let name_byte_len = ((name_utf16.len() + 1) * 2) as u16;
        buf[64..66].copy_from_slice(&name_byte_len.to_le_bytes());
        // Entry type
        buf[66] = entry_type;
        // Color: black (1)
        buf[67] = 1;
        // Left/right siblings: none
        buf[68..72].copy_from_slice(&0xFFFFFFFFu32.to_le_bytes());
        buf[72..76].copy_from_slice(&0xFFFFFFFFu32.to_le_bytes());
        // Child: none
        buf[76..80].copy_from_slice(&0xFFFFFFFFu32.to_le_bytes());
        // Start sector
        buf[116..120].copy_from_slice(&start_sector.to_le_bytes());
        // Size
        buf[120..124].copy_from_slice(&size.to_le_bytes());
    }
}
