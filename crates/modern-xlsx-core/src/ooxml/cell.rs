use core::hint::cold_path;

use crate::{ModernXlsxError, Result};

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

        Ok(CellRef {
            col,
            row: row_1based - 1,
        })
    }

    /// Convert this cell reference back to an A1-style string.
    #[inline]
    pub fn to_a1(&self) -> String {
        let mut buf = [0u8; 3];
        let col_len = col_to_letters_buf(self.col, &mut buf);
        let mut result = String::with_capacity(col_len + 7);
        // SAFETY: col_to_letters_buf writes only ASCII uppercase letters
        result.push_str(unsafe { std::str::from_utf8_unchecked(&buf[..col_len]) });
        let mut itoa_buf = itoa::Buffer::new();
        result.push_str(itoa_buf.format(self.row + 1));
        result
    }
}

impl std::fmt::Display for CellRef {
    #[inline]
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        let mut buf = [0u8; 3];
        let col_len = col_to_letters_buf(self.col, &mut buf);
        // SAFETY: col_to_letters_buf writes only ASCII uppercase letters
        f.write_str(unsafe { std::str::from_utf8_unchecked(&buf[..col_len]) })?;
        let mut itoa_buf = itoa::Buffer::new();
        f.write_str(itoa_buf.format(self.row + 1))
    }
}

/// Write column letters into a 3-byte stack buffer (max XLSX column is "XFD").
///
/// Returns the number of bytes written (1–3). The caller can slice `buf[..len]`.
#[inline]
fn col_to_letters_buf(col: u32, buf: &mut [u8; 3]) -> usize {
    let mut n = col;
    let mut pos = 3usize;

    loop {
        pos -= 1;
        buf[pos] = b'A' + (n % 26) as u8;
        if n < 26 {
            break;
        }
        n = n / 26 - 1;
    }

    // Shift to front if needed
    let len = 3 - pos;
    if pos > 0 {
        buf.copy_within(pos..3, 0);
    }
    len
}

/// Convert a zero-based column index to A1-style column letters.
///
/// 0 → "A", 1 → "B", ..., 25 → "Z", 26 → "AA", 27 → "AB", ..., 16383 → "XFD"
#[inline]
pub fn col_to_letters(col: u32) -> String {
    let mut buf = [0u8; 3];
    let len = col_to_letters_buf(col, &mut buf);
    // SAFETY: col_to_letters_buf only writes ASCII uppercase letters — always valid UTF-8
    unsafe { std::str::from_utf8_unchecked(&buf[..len]) }.to_owned()
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
    fn test_col_letter_roundtrip() {
        for col in 0..16384 {
            let letters = col_to_letters(col);
            let back = letters_to_col(&letters).unwrap();
            assert_eq!(col, back, "roundtrip failed for col {col} → {letters}");
        }
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
