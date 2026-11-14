//! MS-CFB compound files, built in every build (no Cargo feature gates them).
//!
//! - [`reader`]: the hostile-input-hardened reader behind
//!   `ole2::detect::classify_ole2`, `ole2::encryption_info` and package
//!   decryption in `ole2::crypto`: header validation, FAT and DIFAT
//!   collection, chain following and stream reading through the FAT or the
//!   mini stream.
//! - [`writer`]: a version 3 compound-file writer (512-byte sectors) with a
//!   MiniFAT for streams under 4096 bytes and DIFAT sectors once the FAT
//!   outgrows the 109 header slots, used by the encrypted writer in
//!   `ole2::crypto`.
//!
//! The constants below are the ones both halves share in name, type and
//! value. `MINI_STREAM_CUTOFF` stays in each file, because the reader compares
//! it with a `u64` stream size and the writer stores it as a `u32` header word.

pub mod reader;
pub mod writer;

/// Largest regular sector ID (MS-CFB §2.1). Every value above it is a marker,
/// never a sector ID.
pub(crate) const MAXREGSECT: u32 = 0xFFFF_FFFA;
/// Sector ID marker that ends a FAT or MiniFAT chain (MS-CFB §2.1).
pub(crate) const ENDOFCHAIN: u32 = 0xFFFF_FFFE;
/// Sector ID marker of an unallocated sector (MS-CFB §2.1).
pub(crate) const FREESECT: u32 = 0xFFFF_FFFF;

/// Mini sector size: mini sector shift 6, 2^6 = 64 bytes (MS-CFB §2.2).
pub(crate) const MINI_SECTOR_SIZE: usize = 64;
