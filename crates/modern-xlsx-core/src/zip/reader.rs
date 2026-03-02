use std::collections::HashMap;
use std::io::{Cursor, Read};

use log::warn;
use zip::ZipArchive;

use crate::errors::{ModernXlsxError, Result};

/// Security limits for ZIP archive extraction.
pub struct ZipSecurityLimits {
    /// Maximum total decompressed size across all entries (default: 2 GB).
    pub max_decompressed_size: u64,
    /// Maximum allowed ratio of decompressed size to compressed size per entry (default: 100.0).
    pub max_compression_ratio: f64,
}

impl Default for ZipSecurityLimits {
    fn default() -> Self {
        Self {
            max_decompressed_size: 2 * 1024 * 1024 * 1024, // 2 GB
            max_compression_ratio: 100.0,
        }
    }
}

/// Reads a ZIP archive from raw bytes, applying security checks, and returns a map
/// of entry name to decompressed bytes.
///
/// # Security
///
/// - Rejects entries with path traversal components (`..`, leading `/` or `\`).
/// - Rejects entries whose decompression ratio exceeds `limits.max_compression_ratio`.
/// - Rejects archives whose total decompressed size exceeds `limits.max_decompressed_size`.
/// - Skips directory entries.
pub fn read_zip_entries(data: &[u8], limits: &ZipSecurityLimits) -> Result<HashMap<String, Vec<u8>>> {
    let cursor = Cursor::new(data);
    let mut archive = ZipArchive::new(cursor).map_err(|e| ModernXlsxError::ZipRead(e.to_string()))?;

    let mut entries = HashMap::new();
    let mut total_decompressed: u64 = 0;

    for i in 0..archive.len() {
        let mut file = archive
            .by_index(i)
            .map_err(|e| ModernXlsxError::ZipEntry(e.to_string()))?;

        // Skip directories
        if file.is_dir() {
            continue;
        }

        let name = file.name().to_string();

        // Path traversal guard: reject `..` components
        if name.split('/').any(|component| component == "..") {
            return Err(ModernXlsxError::Security(format!(
                "path traversal detected in ZIP entry: {name}"
            )));
        }

        // Path traversal guard: reject absolute paths (leading `/` or `\`)
        if name.starts_with('/') || name.starts_with('\\') {
            return Err(ModernXlsxError::Security(format!(
                "absolute path detected in ZIP entry: {name}"
            )));
        }

        // Also check for backslash-based traversal
        if name.split('\\').any(|component| component == "..") {
            return Err(ModernXlsxError::Security(format!(
                "path traversal detected in ZIP entry: {name}"
            )));
        }

        let compressed_size = file.compressed_size();
        let declared_size = file.size();

        // Read the decompressed data (pre-allocate using declared size).
        let mut buf = Vec::with_capacity(declared_size as usize);
        file.read_to_end(&mut buf)
            .map_err(|e| ModernXlsxError::ZipRead(e.to_string()))?;

        let decompressed_size = buf.len() as u64;

        // Compression ratio guard (only meaningful when compressed size > 0)
        if compressed_size > 0 {
            let ratio = decompressed_size as f64 / compressed_size as f64;
            if ratio > limits.max_compression_ratio {
                return Err(ModernXlsxError::Security(format!(
                    "compression ratio {ratio:.1} exceeds limit {} for entry: {name}",
                    limits.max_compression_ratio
                )));
            }
        } else if declared_size > 0 {
            // compressed_size is 0 but there is data — suspicious
            return Err(ModernXlsxError::Security(format!(
                "zero compressed size with non-zero data for entry: {name}"
            )));
        }

        // Total decompressed size guard
        total_decompressed += decompressed_size;
        if total_decompressed > limits.max_decompressed_size {
            return Err(ModernXlsxError::Security(format!(
                "total decompressed size {} exceeds limit {}",
                total_decompressed, limits.max_decompressed_size
            )));
        }

        if buf.is_empty() {
            warn!("skipping empty ZIP entry: {}", name);
            continue;
        }

        entries.insert(name, buf);
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
        let opts = SimpleFileOptions::default()
            .compression_method(CompressionMethod::Deflated);
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
        let result = read_zip_entries(&zip_bytes, &limits).expect("read_zip_entries should succeed");

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
        };

        let result = read_zip_entries(&zip_bytes, &limits);

        assert!(result.is_err());
        let err = result.unwrap_err();
        assert!(
            matches!(err, ModernXlsxError::Security(ref msg) if msg.contains("total decompressed size")),
            "expected Security error about total decompressed size, got: {err}"
        );
    }
}
