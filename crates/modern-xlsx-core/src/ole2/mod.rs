/// OLE2 Compound Binary File magic signature (ECMA-376 §2.2).
pub(crate) const OLE2_MAGIC: [u8; 8] = [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1];

pub mod crypto;
pub mod detect;
pub mod encryption_info;
pub mod writer;
