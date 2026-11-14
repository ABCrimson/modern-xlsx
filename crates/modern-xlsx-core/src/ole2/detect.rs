//! File-format sniffing and OLE2 classification, built in every build.
//!
//! [`detect_format`] tells a ZIP package from an OLE2 compound file by its
//! magic bytes. [`classify_ole2`] names an OLE2 file an encrypted package
//! (`EncryptionInfo` and `EncryptedPackage` streams), a legacy `.xls`
//! (`Workbook` or `Book` stream) or unknown, reading the directory through the
//! MS-CFB reader in `ole2::cfb::reader`; `classify_ole2_with_cfb` does the same
//! for a file the caller has already parsed.

use super::cfb::reader::Cfb;
use crate::Result;

/// ZIP archive magic bytes (PK\x03\x04).
const ZIP_MAGIC: [u8; 4] = [0x50, 0x4B, 0x03, 0x04];

// Error message constants shared by reader.rs and encryption_info.rs.
pub const ERR_ENCRYPTED: &str = "This file is password-protected (OLE2 compound document). \
    Provide password via readBuffer(data, { password: '...' }).";
/// The lite build sniffs encrypted files but cannot decrypt them.
pub const ERR_ENCRYPTED_LITE: &str = "This file is password-protected (OLE2 compound document). \
    The lite build cannot decrypt; use the full modern-xlsx entry and pass \
    readBuffer(data, { password }).";
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
    if data.starts_with(&super::OLE2_MAGIC) {
        FileFormat::Ole2
    } else if data.starts_with(&ZIP_MAGIC) {
        FileFormat::Zip
    } else {
        FileFormat::Unknown
    }
}

/// Classifies an OLE2 compound document by its stream names (ASCII case
/// ignored; storages never match).
pub fn classify_ole2(data: &[u8]) -> Result<Ole2Kind> {
    Cfb::parse(data).map(|cfb| classify_ole2_with_cfb(&cfb))
}

/// Classifies an already parsed compound file, as [`classify_ole2`] does. The
/// read path classifies through this and keeps the `Cfb` for its stream reads.
pub(crate) fn classify_ole2_with_cfb(cfb: &Cfb<'_>) -> Ole2Kind {
    // Single pass over directory entries to classify the OLE2 document.
    let (mut has_encryption_info, mut has_encrypted_package, mut has_workbook) =
        (false, false, false);
    for e in cfb.entries() {
        has_encryption_info |= e.is_stream_named("EncryptionInfo");
        has_encrypted_package |= e.is_stream_named("EncryptedPackage");
        has_workbook |= e.is_stream_named("Workbook") || e.is_stream_named("Book");
    }

    if has_encryption_info && has_encrypted_package {
        Ole2Kind::EncryptedXlsx
    } else if has_workbook {
        Ole2Kind::LegacyXls
    } else {
        Ole2Kind::Unknown
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
        data[..8].copy_from_slice(&crate::ole2::OLE2_MAGIC);
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

    /// The parsed-once classifier names each file as `classify_ole2` does.
    #[test]
    fn classify_with_cfb_matches_classify_ole2() {
        for names in [
            &["EncryptionInfo", "EncryptedPackage"][..],
            &["Workbook"],
            &["Book"],
            &["EncryptionInfo"],
            &["SomeRandomStream"],
        ] {
            let ole2 = build_test_ole2(names);
            let cfb = Cfb::parse(&ole2).unwrap();
            assert_eq!(
                classify_ole2_with_cfb(&cfb),
                classify_ole2(&ole2).unwrap(),
                "{names:?}"
            );
        }
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
        buf[..8].copy_from_slice(&crate::ole2::OLE2_MAGIC);
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
        // Mini stream cutoff: 4096 (the only value MS-CFB allows)
        buf[56..60].copy_from_slice(&4096u32.to_le_bytes());
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
