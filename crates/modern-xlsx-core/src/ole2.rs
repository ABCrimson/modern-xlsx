//! OLE2 compound files (MS-CFB) and the MS-OFFCRYPTO password encryption that
//! travels inside them.
//!
//! - `base64`: the hand-written base64 codec for the Agile descriptor and
//!   sheet-protection hashes; the decoder is built in every build, the encoder
//!   only with `encryption`.
//! - [`cfb`]: the MS-CFB layer. [`cfb::reader`] validates a compound file and
//!   reads its streams through the FAT or the mini stream; [`cfb::writer`] is
//!   a version 3 compound-file writer (512-byte sectors) with a MiniFAT for
//!   streams under 4096 bytes and DIFAT sectors once the FAT outgrows the 109
//!   header slots. It is built in every build.
//! - [`detect`]: format sniffing (ZIP, OLE2 or neither) and OLE2
//!   classification (encrypted package, legacy `.xls` or unknown). It is built
//!   in every build, so the lite build names encrypted and `.xls` files
//!   precisely.
//! - [`encryption_info`]: the `EncryptionInfo` stream parser for Standard and
//!   Agile descriptors. It uses no crypto crate and is built in every build.
//! - `crypto`: key derivation, password verification, package decryption and
//!   Agile encryption, compiled only with the `encryption` Cargo feature.

/// OLE2 Compound Binary File magic signature (MS-CFB §2.2, Header Signature).
pub(crate) const OLE2_MAGIC: [u8; 8] = [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1];

pub(crate) mod base64;
pub mod cfb;
#[cfg(feature = "encryption")]
pub mod crypto;
pub mod detect;
pub mod encryption_info;
