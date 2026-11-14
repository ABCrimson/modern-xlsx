//! A1-style cell references: [`CellRef`] parsing and formatting, the
//! conversion between zero-based column indexes and column letters, and the
//! worksheet limits [`MAX_COL`] and [`MAX_ROW`].

use core::hint::cold_path;

use crate::{ModernXlsxError, Result};

/// Largest zero-based column index of a worksheet: `XFD` is column 16 384.
pub const MAX_COL: u32 = 16_383;
/// Largest zero-based row index of a worksheet: row 1 048 576.
pub const MAX_ROW: u32 = 1_048_575;

/// A cell reference (zero-based column and row).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct CellRef {
    pub col: u32,
    pub row: u32,
}

impl CellRef {
    /// Parse an A1-style cell reference such as "A1", "AZ100", or "XFD1048576".
    ///
    /// The column is converted from letters (A=0, B=1, ..., Z=25, AA=26, ...)
    /// and the row is converted from the 1-based number to a 0-based index.
    pub fn parse(s: &str) -> Result<Self> {
        if s.is_empty() {
            cold_path();
            return Err(ModernXlsxError::InvalidCellRef(
                "Failed to parse cell reference: input is an empty string".into(),
            ));
        }

        // Find the boundary between letters and digits.
        let first_digit = s
            .bytes()
            .position(|b| b.is_ascii_digit())
            .ok_or_else(|| ModernXlsxError::InvalidCellRef(format!(
                "Failed to parse cell reference '{s}': no row number found (expected format like 'A1')"
            )))?;

        if first_digit == 0 {
            cold_path();
            return Err(ModernXlsxError::InvalidCellRef(format!(
                "Failed to parse cell reference '{s}': no column letters found (expected format like 'A1')"
            )));
        }

        let col_part = &s[..first_digit];
        let row_part = &s[first_digit..];

        // Verify col_part is all ASCII letters.
        if !col_part.bytes().all(|b| b.is_ascii_alphabetic()) {
            cold_path();
            return Err(ModernXlsxError::InvalidCellRef(format!(
                "Failed to parse cell reference '{s}': column part contains non-alphabetic characters"
            )));
        }

        let col = letters_to_col(col_part)?;

        let row_1based: u32 = row_part.parse::<u32>().map_err(|_| {
            ModernXlsxError::InvalidCellRef(format!(
                "Failed to parse cell reference '{s}': row number is not a valid u32 integer"
            ))
        })?;

        if row_1based == 0 {
            cold_path();
            return Err(ModernXlsxError::InvalidCellRef(format!(
                "Failed to parse cell reference '{s}': row number must be >= 1, got 0 (XLSX rows are 1-based)"
            )));
        }

        if row_1based > MAX_ROW + 1 {
            cold_path();
            return Err(ModernXlsxError::InvalidCellRef(format!(
                "Failed to parse cell reference '{s}': row exceeds maximum 1048576"
            )));
        }

        Ok(CellRef {
            col,
            row: row_1based - 1,
        })
    }

    /// Convert this cell reference back to an A1-style string.
    #[inline]
    pub fn to_a1(&self) -> String {
        let mut buf = [0u8; 7];
        let col_len = col_to_letters_buf(self.col, &mut buf);
        let mut result = String::with_capacity(col_len + 7);
        result.extend(buf[..col_len].iter().copied().map(char::from));
        let mut itoa_buf = itoa::Buffer::new();
        result.push_str(itoa_buf.format(self.row + 1));
        result
    }
}

impl std::fmt::Display for CellRef {
    #[inline]
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        let mut buf = [0u8; 7];
        let col_len = col_to_letters_buf(self.col, &mut buf);
        for &b in &buf[..col_len] {
            std::fmt::Write::write_char(f, char::from(b))?;
        }
        let mut itoa_buf = itoa::Buffer::new();
        f.write_str(itoa_buf.format(self.row + 1))
    }
}

/// Write column letters into a 7-byte stack buffer.
///
/// Total over `u32`: 26^7 exceeds `u32::MAX`, so at most 7 letters are written
/// (`XFD`, the sheet maximum, needs 3). Returns the number of bytes written
/// (1–7); the caller can slice `buf[..len]`.
#[inline]
pub(crate) fn col_to_letters_buf(col: u32, buf: &mut [u8; 7]) -> usize {
    let mut n = col;
    let mut pos = 7usize;

    loop {
        pos -= 1;
        buf[pos] = b'A' + (n % 26) as u8;
        if n < 26 {
            break;
        }
        n = n / 26 - 1;
    }

    // Shift to front if needed
    let len = 7 - pos;
    if pos > 0 {
        buf.copy_within(pos..7, 0);
    }
    len
}

/// Convert a zero-based column index to A1-style column letters.
///
/// 0 → "A", 1 → "B", ..., 25 → "Z", 26 → "AA", 27 → "AB", ..., 16383 → "XFD"
#[inline]
pub fn col_to_letters(col: u32) -> String {
    let mut buf = [0u8; 7];
    let len = col_to_letters_buf(col, &mut buf);
    buf[..len].iter().copied().map(char::from).collect()
}

/// Convert A1-style column letters to a zero-based column index.
///
/// "A" → 0, "B" → 1, ..., "Z" → 25, "AA" → 26, ..., "XFD" → 16383
pub fn letters_to_col(s: &str) -> Result<u32> {
    if s.is_empty() {
        cold_path();
        return Err(ModernXlsxError::InvalidCellRef(
            "Failed to parse column letters: input is an empty string".into(),
        ));
    }

    let mut col: u32 = 0;
    for byte in s.bytes() {
        let c = byte.to_ascii_uppercase();
        if !c.is_ascii_uppercase() {
            cold_path();
            return Err(ModernXlsxError::InvalidCellRef(format!(
                "Failed to parse column letters '{s}': invalid character '{}' (expected A-Z only)",
                c as char
            )));
        }
        col = col
            .checked_mul(26)
            .and_then(|v| v.checked_add((c - b'A') as u32 + 1))
            .filter(|&v| v <= MAX_COL + 1)
            .ok_or_else(|| {
                ModernXlsxError::InvalidCellRef(format!(
                    "Failed to parse column letters '{s}': column exceeds maximum XFD (16384 columns)"
                ))
            })?;
    }

    // Convert from 1-based to 0-based.
    Ok(col - 1)
}

#[cfg(test)]
mod tests {
    use super::*;
    use pretty_assertions::assert_eq;

    #[test]
    fn col_letters_roundtrip_all() {
        for c in 0..16384u32 {
            assert_eq!(letters_to_col(&col_to_letters(c)).unwrap(), c);
        }
        assert_eq!(col_to_letters(16383), "XFD");
    }

    #[test]
    fn test_parse_cell_ref() {
        let r = CellRef::parse("A1").unwrap();
        assert_eq!(r.col, 0);
        assert_eq!(r.row, 0);
    }

    #[test]
    fn test_parse_cell_ref_multi_letter() {
        let r = CellRef::parse("AZ100").unwrap();
        assert_eq!(r.col, 51);
        assert_eq!(r.row, 99);
    }

    #[test]
    fn test_parse_cell_ref_max() {
        let r = CellRef::parse("XFD1048576").unwrap();
        assert_eq!(r.col, 16383);
        assert_eq!(r.row, 1048575);
    }

    #[test]
    fn test_cell_ref_to_string() {
        let r = CellRef { col: 0, row: 0 };
        assert_eq!(r.to_a1(), "A1");

        let r2 = CellRef { col: 51, row: 99 };
        assert_eq!(r2.to_a1(), "AZ100");

        let r3 = CellRef {
            col: 16383,
            row: 1048575,
        };
        assert_eq!(r3.to_a1(), "XFD1048576");
    }

    #[test]
    fn test_invalid_cell_ref() {
        // Empty string
        assert!(CellRef::parse("").is_err());
        // No letters (starts with digit)
        assert!(CellRef::parse("123").is_err());
        // No digits
        assert!(CellRef::parse("ABC").is_err());
        // Row = 0
        assert!(CellRef::parse("A0").is_err());
    }

    #[test]
    fn test_parse_rejects_out_of_sheet_refs() {
        for input in ["AAAA1", "XFE1", "A1048577"] {
            let err = CellRef::parse(input).unwrap_err();
            assert_eq!(err.code(), "INVALID_CELL_REF", "{input}");
        }
        assert_eq!(
            CellRef::parse("XFD1048576").unwrap(),
            CellRef {
                col: MAX_COL,
                row: MAX_ROW
            }
        );
    }

    #[test]
    fn test_col_to_letters_total_over_u32() {
        assert_eq!(col_to_letters(16383), "XFD");
        assert_eq!(col_to_letters(18278), "AAAA");
        assert_eq!(col_to_letters(u32::MAX).len(), 7);
        let r = CellRef {
            col: u32::MAX,
            row: 0,
        };
        assert_eq!(r.to_a1().len(), 8);
        assert_eq!(r.to_string(), r.to_a1());
    }

    #[test]
    fn test_letters_to_col_bounds_and_case() {
        assert!(letters_to_col("XFE").is_err());
        assert_eq!(letters_to_col("xfd").unwrap(), 16383);
        let err = letters_to_col("XFE").unwrap_err();
        assert!(
            err.to_string()
                .contains("exceeds maximum XFD (16384 columns)")
        );
    }

    #[test]
    fn test_error_messages_contain_context() {
        // Empty string error includes "empty"
        let err = CellRef::parse("").unwrap_err();
        assert!(
            err.to_string().contains("empty string"),
            "expected 'empty string' in: {err}"
        );

        // Non-alpha column includes the offending input
        let err = CellRef::parse("1A").unwrap_err();
        assert!(
            err.to_string().contains("'1A'"),
            "expected quoted input in: {err}"
        );

        // Column overflow includes "XFD" hint
        let err = letters_to_col("ZZZZZZZZZZ").unwrap_err();
        assert!(
            err.to_string().contains("XFD") || err.to_string().contains("overflow"),
            "expected XFD/overflow hint in: {err}"
        );

        // Row 0 includes "1-based" hint
        let err = CellRef::parse("A0").unwrap_err();
        assert!(
            err.to_string().contains("1-based") || err.to_string().contains(">= 1"),
            "expected 1-based hint in: {err}"
        );

        // Error code is INVALID_CELL_REF
        assert_eq!(err.code(), "INVALID_CELL_REF");
    }
}

#[cfg(test)]
mod prop_tests {
    use super::*;
    use proptest::prelude::*;

    proptest! {
        /// A1 cell reference parse -> to_a1 round-trips for any valid (col, row).
        #[test]
        fn cell_ref_roundtrip(col in 0u32..=16_383, row in 0u32..=1_048_575) {
            let a1 = format!("{}{}", col_to_letters(col), row + 1);
            let parsed = CellRef::parse(&a1).unwrap();
            prop_assert_eq!(parsed.to_a1(), a1);
        }
    }
}
