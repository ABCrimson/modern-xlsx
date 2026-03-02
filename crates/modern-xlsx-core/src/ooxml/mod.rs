pub mod calc_chain;
pub mod cell;
pub mod comments;
pub mod content_types;
pub mod doc_props;
pub mod relationships;
pub mod shared_strings;
pub mod styles;
pub mod theme;
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
