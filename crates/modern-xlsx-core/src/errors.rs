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

    /// Actionable, human-facing guidance for resolving this error.
    ///
    /// Unlike [`Self::code`] (stable, machine-readable), this guidance text may be
    /// refined across versions. Pairs with [`Self::docs_url`] for a deep link.
    #[inline]
    pub fn help(&self) -> &'static str {
        match self {
            Self::ZipRead(_) => {
                "The bytes are not a valid ZIP/XLSX archive or are truncated. Verify you passed a complete .xlsx file."
            }
            Self::ZipWrite(_) | Self::ZipEntry(_) | Self::ZipFinalize(_) | Self::XmlWrite(_) => {
                "Failed while assembling the workbook — this usually indicates an internal bug; please report it with a reproducer."
            }
            Self::XmlParse(_) => {
                "An OOXML part could not be parsed. The file may be malformed or use a feature modern-xlsx does not yet read."
            }
            Self::InvalidCellRef(_) => {
                "Use A1-style references (e.g. \"B2\"): columns A-XFD, rows 1-1048576."
            }
            Self::InvalidCellValue(_) => "The value is out of range or does not match its cell type.",
            Self::InvalidStyle(_) => {
                "Build styles via Workbook.createStyle()/StyleBuilder and apply the returned index."
            }
            Self::InvalidDate(_) => "Excel date serials must be valid; convert dates with dateToSerial().",
            Self::InvalidFormat(_) => {
                "The number-format code is invalid; see the Styling guide for supported tokens."
            }
            Self::MissingPart(_) => {
                "The workbook is missing a required OOXML part and is likely corrupted or not a real .xlsx."
            }
            Self::Security(_) => {
                "A security limit (e.g. zip-bomb / entry-count) was exceeded. If the file is trusted, raise the limits via ZipSecurityLimits."
            }
            Self::PasswordProtected(_) => {
                "The file is encrypted - read it with the password-enabled API (e.g. readWithPassword)."
            }
            Self::LegacyFormat(_) => {
                "Legacy .xls (BIFF/OLE2) files are not supported; convert to .xlsx first."
            }
            Self::UnrecognizedFormat(_) => "The bytes are not a recognized XLSX/OOXML file.",
            Self::Io(_) => {
                "An underlying I/O operation failed; check the path/permissions or the provided reader."
            }
        }
    }

    /// Stable documentation deep-link for this error code (wiki error reference).
    #[inline]
    pub fn docs_url(&self) -> String {
        format!(
            "https://github.com/ABCrimson/modern-xlsx/wiki/FAQ#{}",
            self.code().to_lowercase()
        )
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
    fn test_every_error_has_help_and_docs_url() {
        let samples = [
            ModernXlsxError::ZipRead("x".into()),
            ModernXlsxError::ZipWrite("x".into()),
            ModernXlsxError::XmlParse("x".into()),
            ModernXlsxError::InvalidCellRef("x".into()),
            ModernXlsxError::InvalidCellValue("x".into()),
            ModernXlsxError::InvalidStyle("x".into()),
            ModernXlsxError::InvalidDate("x".into()),
            ModernXlsxError::InvalidFormat("x".into()),
            ModernXlsxError::MissingPart("x".into()),
            ModernXlsxError::Security("x".into()),
            ModernXlsxError::PasswordProtected("x".into()),
            ModernXlsxError::LegacyFormat("x".into()),
            ModernXlsxError::UnrecognizedFormat("x".into()),
        ];
        for err in &samples {
            assert!(!err.help().is_empty(), "{} has empty help", err.code());
            let url = err.docs_url();
            assert!(url.starts_with("https://"), "docs_url not a URL: {url}");
            assert!(
                url.ends_with(&err.code().to_lowercase()),
                "docs_url anchor mismatch: {url}"
            );
        }
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
