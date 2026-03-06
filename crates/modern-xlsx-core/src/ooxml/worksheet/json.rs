use super::CellType;

/// Public wrapper for JSON escaping, used by `reader::read_xlsx_json`.
pub fn json_escape_to_pub(out: &mut String, s: &str) {
    json_escape_to(out, s);
}

/// Write a JSON-escaped string to the output buffer.
/// Handles `"`, `\`, and control characters (0x00-0x1F).
///
/// Scans bytes instead of chars for better performance on ASCII-heavy content.
#[inline]
pub(super) fn json_escape_to(out: &mut String, s: &str) {
    let bytes = s.as_bytes();
    let mut start = 0;
    for (i, &b) in bytes.iter().enumerate() {
        let escape = match b {
            b'"' => "\\\"",
            b'\\' => "\\\\",
            b'\n' => "\\n",
            b'\r' => "\\r",
            b'\t' => "\\t",
            0x00..=0x1f => {
                // Flush pending slice.
                if start < i {
                    // SAFETY: start..i are valid UTF-8 boundaries because we only
                    // break at ASCII bytes (< 0x80), which are always single-byte
                    // codepoints and thus valid boundaries.
                    out.push_str(unsafe { std::str::from_utf8_unchecked(&bytes[start..i]) });
                }
                start = i + 1;
                out.push_str("\\u00");
                let hi = b >> 4;
                let lo = b & 0x0F;
                out.push(char::from(if hi < 10 { b'0' + hi } else { b'a' + hi - 10 }));
                out.push(char::from(if lo < 10 { b'0' + lo } else { b'a' + lo - 10 }));
                continue;
            }
            _ => {
                continue;
            }
        };
        // Flush pending slice before the escape.
        if start < i {
            // SAFETY: same reasoning — only ASCII bytes trigger a break.
            out.push_str(unsafe { std::str::from_utf8_unchecked(&bytes[start..i]) });
        }
        out.push_str(escape);
        start = i + 1;
    }
    // Flush remaining.
    if start < bytes.len() {
        // SAFETY: start is at a valid UTF-8 boundary (set after an ASCII byte).
        out.push_str(unsafe { std::str::from_utf8_unchecked(&bytes[start..]) });
    }
}

/// Convert a `CellType` to its camelCase JSON string (matching serde rename).
#[inline]
pub(super) fn cell_type_json_str(ct: CellType) -> &'static str {
    match ct {
        CellType::SharedString => "sharedString",
        CellType::Number => "number",
        CellType::Boolean => "boolean",
        CellType::Error => "error",
        CellType::FormulaStr => "formulaStr",
        CellType::InlineStr => "inlineStr",
        CellType::Stub => "stub",
    }
}

/// Write an f64 to the JSON output buffer matching serde_json's formatting.
#[inline]
pub(super) fn write_f64_json(out: &mut String, v: f64) {
    if !v.is_finite() {
        out.push_str("null");
        return;
    }
    // serde_json formats floats without trailing zeros for integers.
    if v == v.floor() && v.abs() < 1e15 {
        out.push_str(itoa::Buffer::new().format(v as i64));
        out.push_str(".0");
    } else {
        out.push_str(ryu::Buffer::new().format(v));
    }
}
