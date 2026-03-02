use std::io::{Cursor, Write};

use zip::write::SimpleFileOptions;
use zip::{CompressionMethod, ZipWriter};

use crate::errors::{ModernXlsxError, Result};

/// A single entry to be written into a ZIP archive.
pub struct ZipEntry {
    /// The path/name of the entry within the archive.
    pub name: String,
    /// The uncompressed content of the entry.
    pub data: Vec<u8>,
}

/// Writes the given entries into a new ZIP archive and returns the raw bytes.
///
/// Each entry is compressed using DEFLATE at compression level 6.
pub fn write_zip(entries: &[ZipEntry]) -> Result<Vec<u8>> {
    let buf = Vec::new();
    let cursor = Cursor::new(buf);
    let mut zip_writer = ZipWriter::new(cursor);

    let options = SimpleFileOptions::default()
        .compression_method(CompressionMethod::Deflated)
        .compression_level(Some(6));

    for entry in entries {
        zip_writer
            .start_file(&entry.name, options)
            .map_err(|e| ModernXlsxError::ZipWrite(e.to_string()))?;

        zip_writer
            .write_all(&entry.data)
            .map_err(|e| ModernXlsxError::ZipWrite(e.to_string()))?;
    }

    let cursor = zip_writer
        .finish()
        .map_err(|e| ModernXlsxError::ZipFinalize(e.to_string()))?;

    Ok(cursor.into_inner())
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::zip::reader::{read_zip_entries, ZipSecurityLimits};

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
        let result = read_zip_entries(&zip_bytes, &limits).expect("read_zip_entries should succeed");

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
        let result = read_zip_entries(&zip_bytes, &limits).expect("read_zip_entries should succeed");

        assert!(result.is_empty());
    }
}
