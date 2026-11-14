//! The ZIP container of an OPC package.
//!
//! [`read_zip_entries`] reads every file entry (directories are skipped) into
//! memory under [`ZipSecurityLimits`]; [`write_zip`] writes [`ZipEntry`]
//! values into a new archive.

pub mod reader;
pub mod writer;

pub use reader::{ZipSecurityLimits, read_zip_entries};
pub use writer::{ZipEntry, write_zip};
