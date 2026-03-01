use crate::{IronsheetError, Result};

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
            return Err(IronsheetError::InvalidCellRef("empty string".into()));
        }

        // Find the boundary between letters and digits.
        let first_digit = s
            .bytes()
            .position(|b| b.is_ascii_digit())
            .ok_or_else(|| IronsheetError::InvalidCellRef(format!("no row number in '{s}'")))?;

        if first_digit == 0 {
            return Err(IronsheetError::InvalidCellRef(format!(
                "no column letters in '{s}'"
            )));
        }

        let col_part = &s[..first_digit];
        let row_part = &s[first_digit..];

        // Verify col_part is all ASCII letters.
        if !col_part.bytes().all(|b| b.is_ascii_alphabetic()) {
            return Err(IronsheetError::InvalidCellRef(format!(
                "invalid column letters in '{s}'"
            )));
        }

        let col = letters_to_col(col_part)?;

        let row_1based: u32 = row_part.parse::<u32>().map_err(|_| {
            IronsheetError::InvalidCellRef(format!("invalid row number in '{s}'"))
        })?;

        if row_1based == 0 {
            return Err(IronsheetError::InvalidCellRef(format!(
                "row number must be >= 1, got 0 in '{s}'"
            )));
        }

        Ok(CellRef {
            col,
            row: row_1based - 1,
        })
    }

    /// Convert this cell reference back to an A1-style string.
    pub fn to_a1(&self) -> String {
        let letters = col_to_letters(self.col);
        format!("{}{}", letters, self.row + 1)
    }
}

impl std::fmt::Display for CellRef {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        write!(f, "{}", self.to_a1())
    }
}

/// Convert a zero-based column index to A1-style column letters.
///
/// 0 → "A", 1 → "B", ..., 25 → "Z", 26 → "AA", 27 → "AB", ..., 16383 → "XFD"
pub fn col_to_letters(col: u32) -> String {
    let mut result = Vec::new();
    let mut n = col;

    loop {
        result.push(b'A' + (n % 26) as u8);
        if n < 26 {
            break;
        }
        n = n / 26 - 1;
    }

    result.reverse();
    String::from_utf8(result).expect("column letters are always valid ASCII")
}

/// Convert A1-style column letters to a zero-based column index.
///
/// "A" → 0, "B" → 1, ..., "Z" → 25, "AA" → 26, ..., "XFD" → 16383
pub fn letters_to_col(s: &str) -> Result<u32> {
    if s.is_empty() {
        return Err(IronsheetError::InvalidCellRef(
            "empty column letters".into(),
        ));
    }

    let mut col: u32 = 0;
    for byte in s.bytes() {
        let c = byte.to_ascii_uppercase();
        if !c.is_ascii_uppercase() {
            return Err(IronsheetError::InvalidCellRef(format!(
                "invalid character '{}' in column letters",
                c as char
            )));
        }
        col = col
            .checked_mul(26)
            .and_then(|v| v.checked_add((c - b'A') as u32 + 1))
            .ok_or_else(|| {
                IronsheetError::InvalidCellRef(format!("column overflow for '{s}'"))
            })?;
    }

    // Convert from 1-based to 0-based.
    Ok(col - 1)
}

#[cfg(test)]
mod tests {
    use super::*;

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
}
