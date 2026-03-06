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

impl From<serde_json::Error> for ModernXlsxError {
    fn from(e: serde_json::Error) -> Self {
        ModernXlsxError::XmlParse(e.to_string())
    }
}

pub type Result<T> = std::result::Result<T, ModernXlsxError>;
