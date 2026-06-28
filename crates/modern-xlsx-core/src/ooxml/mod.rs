pub mod calc_chain;
pub mod cell;
pub mod charts;
pub mod comments;
pub mod content_types;
pub mod doc_props;
pub mod pivot_table;
pub mod relationships;
pub mod shared_strings;
pub mod styles;
pub mod tables;
pub mod theme;
pub mod slicers;
pub mod threaded_comments;
pub mod timelines;
pub mod workbook;
pub mod worksheet;

pub(crate) const SPREADSHEET_NS: &str =
    "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

/// Append the resolved character(s) of a quick-xml `Event::GeneralRef` entity
/// to the provided buffer.
///
/// Handles the five predefined XML entities (`amp`, `lt`, `gt`, `quot`, `apos`)
/// and character references (`#N` for decimal, `#xN` for hex).
///
/// This avoids per-call `String` allocation for the common predefined entities.
#[inline]
pub(crate) fn push_entity(buf: &mut String, name: &[u8]) {
    match name {
        b"amp" => buf.push('&'),
        b"lt" => buf.push('<'),
        b"gt" => buf.push('>'),
        b"quot" => buf.push('"'),
        b"apos" => buf.push('\''),
        _ if name.starts_with(b"#x") || name.starts_with(b"#X") => {
            // Hex character reference: &#xHH;
            if let Some(c) = std::str::from_utf8(&name[2..])
                .ok()
                .and_then(|hex| u32::from_str_radix(hex, 16).ok())
                .and_then(char::from_u32)
            {
                buf.push(c);
            }
        }
        _ if name.starts_with(b"#") => {
            // Decimal character reference: &#NN;
            if let Some(c) = std::str::from_utf8(&name[1..])
                .ok()
                .and_then(|dec| dec.parse::<u32>().ok())
                .and_then(char::from_u32)
            {
                buf.push(c);
            }
        }
        _ => {} // Unknown entity — drop silently
    }
}

/// Returns `true` if `c` is forbidden in an XML 1.0 document: the C0 control
/// characters other than tab/LF/CR, plus the permanently-unassigned U+FFFE and
/// U+FFFF noncharacters. Emitting any of these produces a document that strict
/// parsers (including Excel) reject as corrupt.
#[inline]
pub(crate) fn is_illegal_xml_char(c: char) -> bool {
    let u = c as u32;
    (u < 0x20 && c != '\t' && c != '\n' && c != '\r') || u == 0xFFFE || u == 0xFFFF
}

/// Drop characters illegal in XML 1.0 from user-supplied text so the writer can
/// never emit a corrupt document. Returns the input borrowed (zero allocation)
/// when it is already clean, which is the overwhelmingly common case.
pub(crate) fn sanitize_xml_text(s: &str) -> std::borrow::Cow<'_, str> {
    if s.chars().any(is_illegal_xml_char) {
        std::borrow::Cow::Owned(s.chars().filter(|c| !is_illegal_xml_char(*c)).collect())
    } else {
        std::borrow::Cow::Borrowed(s)
    }
}

/// Serde helper: skip serializing if the value is `false`.
#[inline]
pub(crate) fn is_false(v: &bool) -> bool {
    !v
}

/// Serde helper: skip serializing if the value is `true`.
#[inline]
pub(crate) fn is_true(v: &bool) -> bool {
    *v
}

/// Serde default that returns `true`.
#[inline]
pub(crate) fn default_true() -> bool {
    true
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn sanitize_xml_text_borrows_when_clean() {
        // Clean text (including the legal control chars tab/LF/CR) is returned
        // borrowed — no allocation.
        let s = "hello\tworld\r\nok ✓ 日本語";
        assert!(matches!(sanitize_xml_text(s), std::borrow::Cow::Borrowed(_)));
        assert_eq!(sanitize_xml_text(s), s);
    }

    #[test]
    fn sanitize_xml_text_drops_illegal_chars() {
        // C0 control chars (other than tab/LF/CR) and the U+FFFE/U+FFFF
        // noncharacters are removed; legal whitespace is preserved.
        let dirty = "a\u{0001}b\u{0008}c\u{000B}d\u{001F}e\u{FFFE}f\u{FFFF}g\th\ni";
        assert_eq!(sanitize_xml_text(dirty), "abcdefg\th\ni");
    }

    #[test]
    fn is_illegal_xml_char_classification() {
        for c in ['\u{0000}', '\u{0001}', '\u{0008}', '\u{000B}', '\u{001F}', '\u{FFFE}', '\u{FFFF}']
        {
            assert!(is_illegal_xml_char(c), "{c:?} should be illegal");
        }
        for c in ['\t', '\n', '\r', ' ', 'a', '✓', '\u{FFFD}'] {
            assert!(!is_illegal_xml_char(c), "{c:?} should be legal");
        }
    }
}
