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

    #[error("invalid style: {0}")]
    InvalidStyle(String),

    #[error("invalid date serial number: {0}")]
    InvalidDate(String),

    #[error("invalid number format: {0}")]
    InvalidFormat(String),

    #[error("missing required part: {0}")]
    MissingPart(String),

    #[error("security violation: {0}")]
    Security(String),

    /// The file is a password-protected OLE2 compound document.
    #[error("Password protected: {0}")]
    PasswordProtected(String),

    /// The file is a legacy .xls (OLE2) format.
    #[error("Legacy format: {0}")]
    LegacyFormat(String),

    /// The file format is unrecognized.
    #[error("Unrecognized format: {0}")]
    UnrecognizedFormat(String),

    #[error("I/O error: {0}")]
    Io(#[from] std::io::Error),
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
            Self::InvalidStyle(_) => "INVALID_STYLE",
            Self::InvalidDate(_) => "INVALID_DATE",
            Self::InvalidFormat(_) => "INVALID_FORMAT",
            Self::MissingPart(_) => "MISSING_PART",
            Self::Security(_) => "SECURITY",
            Self::PasswordProtected(_) => "PASSWORD_PROTECTED",
            Self::LegacyFormat(_) => "LEGACY_FORMAT",
            Self::UnrecognizedFormat(_) => "UNRECOGNIZED_FORMAT",
            Self::Io(_) => "IO_ERROR",
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

impl From<serde_json::Error> for ModernXlsxError {
    fn from(e: serde_json::Error) -> Self {
        ModernXlsxError::XmlParse(format!(
            "Failed to deserialize JSON: {e} (line {}, column {})",
            e.line(),
            e.column()
        ))
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
            ModernXlsxError::InvalidStyle("x".into()).code(),
            "INVALID_STYLE"
        );
        assert_eq!(
            ModernXlsxError::InvalidDate("x".into()).code(),
            "INVALID_DATE"
        );
        assert_eq!(
            ModernXlsxError::InvalidFormat("x".into()).code(),
            "INVALID_FORMAT"
        );
        assert_eq!(
            ModernXlsxError::MissingPart("x".into()).code(),
            "MISSING_PART"
        );
        assert_eq!(ModernXlsxError::Security("x".into()).code(), "SECURITY");
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
    }

    #[test]
    fn test_coded_string_format() {
        let err = ModernXlsxError::MissingPart("xl/workbook.xml".into());
        let coded = err.to_coded_string();
        assert_eq!(coded, "[MISSING_PART] missing required part: xl/workbook.xml");
    }

    #[test]
    fn test_serde_json_error_includes_context() {
        let bad_json = r#"{"invalid": }"#;
        let err: ModernXlsxError = serde_json::from_str::<serde_json::Value>(bad_json)
            .unwrap_err()
            .into();
        let msg = err.to_string();
        assert!(
            msg.contains("Failed to deserialize JSON"),
            "expected context prefix, got: {msg}"
        );
        assert!(msg.contains("line"), "expected line info, got: {msg}");
        assert!(msg.contains("column"), "expected column info, got: {msg}");
    }
}
