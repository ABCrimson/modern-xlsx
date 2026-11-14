//! Writes [`ZipEntry`] values into a new in-memory ZIP archive.
//!
//! [`write_zip`] compresses every entry with Deflate at level 6 and fails with
//! `ZipWrite` on a duplicate entry name.

use std::io::{Cursor, Write};

use zip::write::SimpleFileOptions;
use zip::{CompressionMethod, ZipWriter};

use crate::errors::{ModernXlsxError, Result};

/// A single entry to be written into a ZIP archive.
#[derive(Debug, Clone)]
pub struct ZipEntry {
    /// The path/name of the entry within the archive.
    pub name: String,
    /// The uncompressed content of the entry.
    pub data: Vec<u8>,
}

/// Deflate at level 6: the options every ZIP entry is written with.
pub(crate) fn deflate_options() -> SimpleFileOptions {
    SimpleFileOptions::default()
        .compression_method(CompressionMethod::Deflated)
        .compression_level(Some(6))
}

/// Writes the given entries into a new ZIP archive and returns the raw bytes.
///
/// Each entry is compressed using DEFLATE at compression level 6.
pub fn write_zip(entries: &[ZipEntry]) -> Result<Vec<u8>> {
    // Compressed output is usually well under 1/4 of the XML; the Vec grows if not.
    // Per entry: 30-B local header + 46-B central record + the name twice; +22-B EOCD.
    let estimated = entries
        .iter()
        .map(|e| {
            (e.data.len() / 4)
                .saturating_add(76)
                .saturating_add(2 * e.name.len())
        })
        .fold(22usize, usize::saturating_add);
    let buf = Vec::with_capacity(estimated);
    let cursor = Cursor::new(buf);
    let mut zip_writer = ZipWriter::new(cursor);

    let options = deflate_options();

    for entry in entries {
        zip_writer.start_file(&entry.name, options).map_err(|e| {
            ModernXlsxError::ZipWrite(format!("Failed to start ZIP entry '{}': {e}", entry.name))
        })?;

        zip_writer.write_all(&entry.data).map_err(|e| {
            ModernXlsxError::ZipWrite(format!(
                "Failed to write {} bytes to ZIP entry '{}': {e}",
                entry.data.len(),
                entry.name
            ))
        })?;
    }

    let cursor = zip_writer.finish().map_err(|e| {
        ModernXlsxError::ZipFinalize(format!(
            "Failed to finalize ZIP archive ({} entries): {e}",
            entries.len()
        ))
    })?;

    Ok(cursor.into_inner())
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::zip::reader::{ZipSecurityLimits, read_zip_entries};

    #[test]
    fn write_zip_capacity_is_compressed_sized() {
        let name = "xl/worksheets/sheet1.xml";
        let cell = br#"<c r="A1"><v>1</v></c>"#;
        let data: Vec<u8> = cell.iter().copied().cycle().take(1 << 20).collect();
        let out = write_zip(&[ZipEntry {
            name: name.to_string(),
            data,
        }])
        .expect("write_zip should succeed");
        // 30-B local header + 46-B central record + the name twice, plus the 22-B EOCD.
        let overhead = 76 + 2 * name.len() + 22;
        assert!(out.capacity() <= (1 << 20) / 4 + overhead);
        let mut archive = zip::ZipArchive::new(Cursor::new(&out[..])).expect("valid archive");
        let compressed = archive.by_index(0).expect("entry 0").compressed_size();
        assert_eq!(out.len() as u64, compressed + overhead as u64);
    }

    #[test]
    fn test_write_and_read_roundtrip() {
        let entries = vec![
            ZipEntry {
                name: "file1.txt".to_string(),
                data: b"content of file one".to_vec(),
            },
            ZipEntry {
                name: "subdir/file2.xml".to_string(),
                data: b"<data>value</data>".to_vec(),
            },
        ];

        let zip_bytes = write_zip(&entries).expect("write_zip should succeed");
        let limits = ZipSecurityLimits::default();
        let result =
            read_zip_entries(&zip_bytes, &limits).expect("read_zip_entries should succeed");

        assert_eq!(result.len(), 2);
        assert_eq!(result.get("file1.txt").unwrap(), b"content of file one");
        assert_eq!(
            result.get("subdir/file2.xml").unwrap(),
            b"<data>value</data>"
        );
    }

    #[test]
    fn test_write_empty_zip() {
        let entries: Vec<ZipEntry> = vec![];

        let zip_bytes = write_zip(&entries).expect("write_zip should succeed");
        let limits = ZipSecurityLimits::default();
        let result =
            read_zip_entries(&zip_bytes, &limits).expect("read_zip_entries should succeed");

        assert!(result.is_empty());
    }

    #[test]
    fn test_write_rejects_duplicate_names() {
        let entries = vec![
            ZipEntry {
                name: "a.xml".to_string(),
                data: b"<a/>".to_vec(),
            },
            ZipEntry {
                name: "a.xml".to_string(),
                data: b"<b/>".to_vec(),
            },
        ];

        let err = write_zip(&entries).expect_err("duplicate names must be rejected");
        assert_eq!(err.code(), "ZIP_WRITE");
    }

    #[test]
    fn test_write_deflates_every_entry() {
        let entries = vec![
            ZipEntry {
                name: "a.xml".to_string(),
                data: b"<a>aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa</a>".to_vec(),
            },
            ZipEntry {
                name: "empty.xml".to_string(),
                data: Vec::new(),
            },
        ];

        let zip_bytes = write_zip(&entries).expect("write_zip should succeed");
        let mut archive =
            zip::ZipArchive::new(Cursor::new(zip_bytes)).expect("archive should open");
        assert_eq!(archive.len(), 2);
        for i in 0..archive.len() {
            let file = archive.by_index(i).expect("entry should open");
            assert_eq!(file.compression(), CompressionMethod::Deflated);
        }
    }
}
