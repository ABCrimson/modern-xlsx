//! The crate's error type, [`ModernXlsxError`], and its [`Result`] alias.
//!
//! Each variant carries a message and maps to a stable code through
//! [`ModernXlsxError::code`]. [`ModernXlsxError::to_coded_string`] formats an
//! error as `[CODE] message`, the form the WASM bridge sends to JavaScript,
//! where the code becomes `ModernXlsxError.code`.

use thiserror::Error;

#[derive(Error, Debug)]
#[non_exhaustive]
pub enum ModernXlsxError {
    #[error("ZIP read error: {0}")]
    ZipRead(String),

    #[error("ZIP write error: {0}")]
    ZipWrite(String),

    #[error("ZIP entry error: {0}")]
    ZipEntry(String),

    #[error("ZIP finalize error: {0}")]
    ZipFinalize(String),

    #[error("XML parse error: {0}")]
    XmlParse(String),

    #[error("XML write error: {0}")]
    XmlWrite(String),

    #[error("invalid cell reference: {0}")]
    InvalidCellRef(String),

    #[error("invalid cell value: {0}")]
    InvalidCellValue(String),

    #[error("invalid date serial number: {0}")]
    InvalidDate(String),

    #[error("missing required part: {0}")]
    MissingPart(String),

    #[error("security violation: {0}")]
    Security(String),

    /// A hard size limit was exceeded: the JSON of a read grew longer than the configured maximum, or a read copied text stored once (a shared string, a comment author, a hyperlink target) past that same maximum; an OLE2 document being written exceeds what MS-CFB v3 can address (a stream above 0x8000_0000 bytes, too many sectors) or this platform's address space; or the writer would complete a pivot field of a table built in code with more than [`MAX_FIELD_ITEMS`](crate::ooxml::pivot_table::MAX_FIELD_ITEMS) items, or add more than [`MAX_ADDED_ITEMS`](crate::ooxml::pivot_table::MAX_ADDED_ITEMS) items and pivot fields to the pivot tables of one workbook; or one write would generate more than `vml::MAX_GENERATED_NOTE_SHAPES` (131,072) notes, the VML note shapes for legacy comments and the `tc=` placeholders for threaded comments.
    #[error("size limit exceeded: {0}")]
    SizeLimit(String),

    /// A password-protected (encrypted) file cannot be handled: it was read
    /// without a password or in a build without `encryption`, the password is
    /// wrong, the integrity check fails, the encryption is unsupported or
    /// malformed, or a cipher step failed while encrypting or decrypting.
    #[error("Password protected: {0}")]
    PasswordProtected(String),

    /// The file is a legacy .xls (OLE2) format.
    #[error("Legacy format: {0}")]
    LegacyFormat(String),

    /// The file format is unrecognized.
    #[error("Unrecognized format: {0}")]
    UnrecognizedFormat(String),

    /// JSON could not be (de)serialized; a parse error's message ends with serde_json's "at line N column M".
    #[error("JSON error: {0}")]
    Json(String),

    /// The workbook cannot be written as given: it has no sheets, two pivot
    /// caches share a `cacheId` or a pivot table names one no cache has, a
    /// theme colour of a generated theme is not 6-digit hex, a colour `rgb` (a
    /// font, fill, border, dxf, tab, conditional-format, sparkline or rich-text
    /// run colour) is not 6 or 8 hex digits after an optional `#`, a
    /// conditional-formatting rule has more than three formulas (`formula` plus
    /// `additional_formulas`), two series of a chart share a `c:order` (within
    /// one chart-type element or across its primary and secondary chart), or a sheet's `sheetProtection` password is not a
    /// legacy hash (1 to 4 hex digits) and
    /// no `plain_password` replaces it, or, in a build without `encryption`, a
    /// `plain_password` holds a non-ASCII character (the legacy hash takes ASCII
    /// passwords only).
    #[error("invalid workbook: {0}")]
    InvalidWorkbook(String),
}

impl ModernXlsxError {
    /// Return a stable machine-readable error code for programmatic handling.
    ///
    /// These codes are forwarded through the WASM boundary and exposed on the
    /// TypeScript `ModernXlsxError.code` property. Codes follow the convention
    /// `UPPER_SNAKE_CASE` and are guaranteed stable across minor versions.
    #[inline]
    pub fn code(&self) -> &'static str {
        match self {
            Self::ZipRead(_) => "ZIP_READ",
            Self::ZipWrite(_) => "ZIP_WRITE",
            Self::ZipEntry(_) => "ZIP_ENTRY",
            Self::ZipFinalize(_) => "ZIP_FINALIZE",
            Self::XmlParse(_) => "XML_PARSE",
            Self::XmlWrite(_) => "XML_WRITE",
            Self::InvalidCellRef(_) => "INVALID_CELL_REF",
            Self::InvalidCellValue(_) => "INVALID_CELL_VALUE",
            Self::InvalidDate(_) => "INVALID_DATE",
            Self::MissingPart(_) => "MISSING_PART",
            Self::Security(_) => "SECURITY",
            Self::SizeLimit(_) => "SIZE_LIMIT",
            Self::PasswordProtected(_) => "PASSWORD_PROTECTED",
            Self::LegacyFormat(_) => "LEGACY_FORMAT",
            Self::UnrecognizedFormat(_) => "UNRECOGNIZED_FORMAT",
            Self::Json(_) => "JSON",
            Self::InvalidWorkbook(_) => "INVALID_WORKBOOK",
        }
    }

    /// Format the error as `"[CODE] message"` for the WASM boundary.
    ///
    /// The TypeScript layer parses this format to extract both the machine-readable
    /// code and the human-readable message.
    #[inline]
    pub fn to_coded_string(&self) -> String {
        format!("[{}] {}", self.code(), self)
    }
}

/// serde_json's `Display` already ends with "at line N column M" when the error has a position.
impl From<serde_json::Error> for ModernXlsxError {
    fn from(e: serde_json::Error) -> Self {
        Self::Json(e.to_string())
    }
}

pub type Result<T> = std::result::Result<T, ModernXlsxError>;

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_error_codes_are_stable() {
        assert_eq!(ModernXlsxError::ZipRead("x".into()).code(), "ZIP_READ");
        assert_eq!(ModernXlsxError::ZipWrite("x".into()).code(), "ZIP_WRITE");
        assert_eq!(ModernXlsxError::ZipEntry("x".into()).code(), "ZIP_ENTRY");
        assert_eq!(
            ModernXlsxError::ZipFinalize("x".into()).code(),
            "ZIP_FINALIZE"
        );
        assert_eq!(ModernXlsxError::XmlParse("x".into()).code(), "XML_PARSE");
        assert_eq!(ModernXlsxError::XmlWrite("x".into()).code(), "XML_WRITE");
        assert_eq!(
            ModernXlsxError::InvalidCellRef("x".into()).code(),
            "INVALID_CELL_REF"
        );
        assert_eq!(
            ModernXlsxError::InvalidCellValue("x".into()).code(),
            "INVALID_CELL_VALUE"
        );
        assert_eq!(
            ModernXlsxError::InvalidDate("x".into()).code(),
            "INVALID_DATE"
        );
        assert_eq!(
            ModernXlsxError::MissingPart("x".into()).code(),
            "MISSING_PART"
        );
        assert_eq!(ModernXlsxError::Security("x".into()).code(), "SECURITY");
        assert_eq!(ModernXlsxError::SizeLimit("x".into()).code(), "SIZE_LIMIT");
        assert_eq!(
            ModernXlsxError::PasswordProtected("x".into()).code(),
            "PASSWORD_PROTECTED"
        );
        assert_eq!(
            ModernXlsxError::LegacyFormat("x".into()).code(),
            "LEGACY_FORMAT"
        );
        assert_eq!(
            ModernXlsxError::UnrecognizedFormat("x".into()).code(),
            "UNRECOGNIZED_FORMAT"
        );
        assert_eq!(ModernXlsxError::Json("x".into()).code(), "JSON");
        assert_eq!(
            ModernXlsxError::InvalidWorkbook("x".into()).code(),
            "INVALID_WORKBOOK"
        );
    }

    #[test]
    fn test_size_limit_contract() {
        let err = ModernXlsxError::SizeLimit("x".into());
        assert_eq!(err.to_coded_string(), "[SIZE_LIMIT] size limit exceeded: x");
    }

    #[test]
    fn test_coded_string_format() {
        let err = ModernXlsxError::MissingPart("xl/workbook.xml".into());
        let coded = err.to_coded_string();
        assert_eq!(
            coded,
            "[MISSING_PART] missing required part: xl/workbook.xml"
        );
    }

    #[test]
    fn test_serde_json_error_includes_context() {
        let bad_json = r#"{"invalid": }"#;
        let err: ModernXlsxError = serde_json::from_str::<serde_json::Value>(bad_json)
            .unwrap_err()
            .into();
        assert_eq!(err.code(), "JSON");
        let msg = err.to_string();
        assert!(
            msg.starts_with("JSON error: "),
            "expected context prefix, got: {msg}"
        );
        assert_eq!(
            msg.matches("line").count(),
            1,
            "position printed once: {msg}"
        );
        assert_eq!(
            msg.matches("column").count(),
            1,
            "position printed once: {msg}"
        );
    }

    #[test]
    fn test_json_error_code_and_position() {
        let err = ModernXlsxError::from(serde_json::from_str::<u8>("x").unwrap_err());
        assert_eq!(err.code(), "JSON");
        assert_eq!(err.to_string().matches("line").count(), 1, "got: {err}");
        assert!(err.to_coded_string().starts_with("[JSON] JSON error: "));
    }

    #[test]
    fn test_empty_workbook_is_invalid_workbook() {
        let err = crate::writer::write_xlsx(&crate::WorkbookData::default())
            .expect_err("a workbook without sheets must not be written");
        assert_eq!(err.code(), "INVALID_WORKBOOK");

        let err = crate::streaming_writer::StreamingWriterCore::new()
            .finish()
            .expect_err("a streaming workbook without sheets must not be finished");
        assert_eq!(err.code(), "INVALID_WORKBOOK");
    }
}
