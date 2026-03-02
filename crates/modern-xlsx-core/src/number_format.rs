/// The high-level category of a number format.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum FormatType {
    Number,
    Date,
    Time,
    DateTime,
    Text,
}

/// Classify a built-in Excel format by its numeric ID.
///
/// Returns the format type for well-known IDs, or `Number` for unknown IDs.
pub fn classify_format(id: u32) -> FormatType {
    match id {
        // General
        0 => FormatType::Number,
        // Number formats: 0, 0.00, #,##0, etc.
        1..=8 => FormatType::Number,
        // Percentage
        9 | 10 => FormatType::Number,
        // Fraction
        11 | 12 => FormatType::Number,
        // Scientific
        13 => FormatType::Number,
        // Date formats
        14..=17 => FormatType::Date,
        // Time formats
        18..=21 => FormatType::Time,
        // DateTime
        22 => FormatType::DateTime,
        // Time/date formats (locale-dependent, IDs 23-26)
        23..=26 => FormatType::Date,
        // CJK date formats
        27..=36 => FormatType::Date,
        // More number formats
        37..=44 => FormatType::Number,
        // Time formats
        45..=47 => FormatType::Time,
        // Scientific
        48 => FormatType::Number,
        // @ (Text)
        49 => FormatType::Text,
        // CJK date
        50..=58 => FormatType::Date,
        // Default for unknown IDs
        _ => FormatType::Number,
    }
}

/// Classify a custom format string into a format type.
///
/// The algorithm:
/// 1. Strip bracketed content (except elapsed-time markers `[h]`, `[m]`, `[s]`)
/// 2. Strip quoted strings between `"`
/// 3. Strip escape sequences (`\x`)
/// 4. Detect date tokens (y, d, and m-when-used-as-month)
/// 5. Detect time tokens (h, s, AM/PM, and m-when-used-as-minutes)
/// 6. `m` ambiguity: if preceded by `h` or followed by `s` → minutes, else → month.
pub fn classify_format_string(format: &str) -> FormatType {
    // Special cases.
    if format == "@" {
        return FormatType::Text;
    }
    if format.eq_ignore_ascii_case("General") {
        return FormatType::Number;
    }

    // clean_format_string returns already-lowercased bytes so that
    // has_date_tokens / has_time_tokens can scan without allocating.
    let cleaned = clean_format_string(format);

    let has_date = has_date_tokens(&cleaned);
    let has_time = has_time_tokens(&cleaned);

    if has_date && has_time {
        FormatType::DateTime
    } else if has_date {
        FormatType::Date
    } else if has_time {
        FormatType::Time
    } else {
        FormatType::Number
    }
}

/// Remove bracketed content (except [h], [m], [s]), quoted strings, and escape sequences.
///
/// Returns already-lowercased bytes so callers can scan without further allocation.
fn clean_format_string(format: &str) -> Vec<u8> {
    let bytes = format.as_bytes();
    let len = bytes.len();
    let mut result = Vec::with_capacity(len);
    let mut i = 0;

    while i < len {
        match bytes[i] {
            b'[' => {
                // Find closing bracket.
                if let Some(close) = bytes[i..].iter().position(|&b| b == b']') {
                    let content = &bytes[i + 1..i + close];
                    // Keep elapsed time markers [h], [m], [s], [hh], [mm], [ss].
                    if is_elapsed_time_marker(content) {
                        // Push the content lowercased.
                        for &b in content {
                            result.push(b.to_ascii_lowercase());
                        }
                    }
                    // Otherwise strip the bracketed content (colors, conditions, etc.)
                    i += close + 1;
                } else {
                    // No closing bracket; just emit the byte lowercased.
                    result.push(bytes[i].to_ascii_lowercase());
                    i += 1;
                }
            }
            b'"' => {
                // Skip quoted string.
                i += 1;
                while i < len && bytes[i] != b'"' {
                    i += 1;
                }
                if i < len {
                    i += 1; // skip closing quote
                }
            }
            b'\\' => {
                // Skip escape sequence.
                i += 2;
            }
            _ => {
                result.push(bytes[i].to_ascii_lowercase());
                i += 1;
            }
        }
    }

    result
}

/// Check whether a bracket's content matches an elapsed-time marker
/// (h, m, s, hh, mm, ss — case-insensitive) without allocating.
fn is_elapsed_time_marker(content: &[u8]) -> bool {
    match content.len() {
        1 => {
            let c = content[0].to_ascii_lowercase();
            c == b'h' || c == b'm' || c == b's'
        }
        2 => {
            let a = content[0].to_ascii_lowercase();
            let b = content[1].to_ascii_lowercase();
            a == b && (a == b'h' || a == b'm' || a == b's')
        }
        _ => false,
    }
}

/// Check whether the cleaned (already-lowercased) byte slice contains date tokens.
///
/// Date tokens are: y, d, and m-when-used-as-month.
/// m is a month token when it is NOT preceded by h and NOT followed by s.
fn has_date_tokens(cleaned: &[u8]) -> bool {
    // Check for y or d.
    if cleaned.iter().any(|&b| b == b'y' || b == b'd') {
        return true;
    }

    // Check for m used as month (not preceded by h and not followed by s).
    for (i, &b) in cleaned.iter().enumerate() {
        if b == b'm' && is_month_m(cleaned, i) {
            return true;
        }
    }

    false
}

/// Check whether the cleaned (already-lowercased) byte slice contains time tokens.
///
/// Time tokens are: h, s, AM/PM, and m-when-used-as-minutes.
fn has_time_tokens(cleaned: &[u8]) -> bool {
    // Check for h or s.
    if cleaned.iter().any(|&b| b == b'h' || b == b's') {
        return true;
    }

    // Check for AM/PM.
    if contains_subsequence(cleaned, b"am/pm") || contains_subsequence(cleaned, b"a/p") {
        return true;
    }

    // Check for m used as minutes (preceded by h or followed by s).
    for (i, &b) in cleaned.iter().enumerate() {
        if b == b'm' && !is_month_m(cleaned, i) {
            return true;
        }
    }

    false
}

/// Check whether `haystack` contains `needle` as a contiguous subsequence.
fn contains_subsequence(haystack: &[u8], needle: &[u8]) -> bool {
    haystack.windows(needle.len()).any(|w| w == needle)
}

/// Determine whether an `m` at position `i` in the (lowercased) byte slice
/// is a **month** token (returns true) or a **minutes** token (returns false).
///
/// Rule: if the nearest non-m token before is `h` → minutes.
///       if the nearest non-m token after is `s` → minutes.
///       otherwise → month.
fn is_month_m(bytes: &[u8], i: usize) -> bool {
    // Look backwards past any other 'm' characters for the nearest relevant token.
    let preceded_by_h = {
        let mut j = i;
        while j > 0 {
            j -= 1;
            match bytes[j] {
                b'm' => continue,
                b'h' => break,
                b if !b.is_ascii_alphabetic() => continue,
                _ => break,
            }
        }
        j < i && bytes[j] == b'h'
    };

    if preceded_by_h {
        return false; // minutes
    }

    // Look forward past any other 'm' characters for the nearest relevant token.
    let followed_by_s = {
        let mut j = i + 1;
        while j < bytes.len() {
            match bytes[j] {
                b'm' => {
                    j += 1;
                    continue;
                }
                b's' => break,
                b if !b.is_ascii_alphabetic() => {
                    j += 1;
                    continue;
                }
                _ => break,
            }
        }
        j < bytes.len() && bytes[j] == b's'
    };

    if followed_by_s {
        return false; // minutes
    }

    true // month
}

#[cfg(test)]
mod tests {
    use super::*;
    use pretty_assertions::assert_eq;

    // --- Built-in format ID tests ---

    #[test]
    fn test_builtin_general() {
        assert_eq!(classify_format(0), FormatType::Number);
    }

    #[test]
    fn test_builtin_date() {
        assert_eq!(classify_format(14), FormatType::Date);
    }

    #[test]
    fn test_builtin_time() {
        assert_eq!(classify_format(18), FormatType::Time);
    }

    #[test]
    fn test_builtin_datetime() {
        assert_eq!(classify_format(22), FormatType::DateTime);
    }

    #[test]
    fn test_builtin_text() {
        assert_eq!(classify_format(49), FormatType::Text);
    }

    // --- Custom format string tests ---

    #[test]
    fn test_custom_date() {
        assert_eq!(classify_format_string("yyyy-mm-dd"), FormatType::Date);
    }

    #[test]
    fn test_custom_time() {
        assert_eq!(classify_format_string("h:mm:ss"), FormatType::Time);
    }

    #[test]
    fn test_custom_datetime() {
        assert_eq!(
            classify_format_string("m/d/yyyy h:mm"),
            FormatType::DateTime
        );
    }

    #[test]
    fn test_custom_number() {
        assert_eq!(classify_format_string("#,##0.00"), FormatType::Number);
    }

    #[test]
    fn test_color_bracket_stripped() {
        // [Red] should be stripped; yyyy-mm-dd is a Date format.
        assert_eq!(classify_format_string("[Red]yyyy-mm-dd"), FormatType::Date);
    }

    #[test]
    fn test_condition_bracket_stripped() {
        assert_eq!(
            classify_format_string("[>100]#,##0;#,##0.00"),
            FormatType::Number
        );
    }

    #[test]
    fn test_quoted_strings_stripped() {
        // The "Date: " part is quoted and should be stripped.
        assert_eq!(
            classify_format_string("\"Date: \"yyyy-mm-dd"),
            FormatType::Date
        );
    }

    #[test]
    fn test_m_ambiguity_time() {
        // h:mm → m is after h, so it's minutes → Time
        assert_eq!(classify_format_string("h:mm"), FormatType::Time);
    }

    #[test]
    fn test_m_ambiguity_date() {
        // m/d → m is a month (no h before, no s after) → Date
        assert_eq!(classify_format_string("m/d"), FormatType::Date);
    }

    #[test]
    fn test_at_sign_text() {
        assert_eq!(classify_format_string("@"), FormatType::Text);
    }

    #[test]
    fn test_general_string() {
        assert_eq!(classify_format_string("General"), FormatType::Number);
    }
}
