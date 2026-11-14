//! One module per OOXML part: workbook, worksheet, styles, shared strings,
//! comments, tables, charts, pivot tables and the other parts this crate reads
//! and writes, plus the `.rels` and `[Content_Types].xml` package plumbing and the shared
//! building blocks `cell` (A1 references and sheet limits), `color` (the
//! SpreadsheetML colour type) and `xml` (XML decoding and formatting helpers).
//!
//! The helpers defined in this file are: the serde predicates
//! `is_false` and `is_true` and the `default_true` default, the ARGB
//! colour normaliser `to_argb`, `write_xml_decl` for the XML prologue and
//! `xml_write_err` for writer I/O errors.

pub mod calc_chain;
pub mod cell;
pub mod charts;
pub mod color;
pub mod comments;
pub mod content_types;
pub mod doc_props;
pub mod pivot_table;
pub mod relationships;
pub mod shared_strings;
pub mod slicers;
pub mod styles;
pub mod tables;
pub mod theme;
pub mod threaded_comments;
pub mod timelines;
pub(crate) mod vml;
pub mod workbook;
pub mod worksheet;
pub(crate) mod xml;

pub(crate) use xml::{SPREADSHEET_NS, is_illegal_xml_char, sanitize_xml_text};

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

/// Map a quick-xml writer I/O error to `ModernXlsxError::XmlWrite` (code `XML_WRITE`).
#[inline]
pub(crate) fn xml_write_err(e: std::io::Error) -> crate::errors::ModernXlsxError {
    crate::errors::ModernXlsxError::XmlWrite(e.to_string())
}

/// Write the `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` prologue
/// that every OOXML part starts with.
pub(crate) fn write_xml_decl<W: std::io::Write>(
    writer: &mut quick_xml::Writer<W>,
) -> crate::errors::Result<()> {
    use quick_xml::events::{BytesDecl, Event};
    writer
        .write_event(Event::Decl(BytesDecl::new(
            "1.0",
            Some("UTF-8"),
            Some("yes"),
        )))
        .map_err(xml_write_err)
}

/// Normalise a SpreadsheetML colour to `ST_UnsignedIntHex` ARGB (`AARRGGBB`).
///
/// One leading `#` is stripped. Six hex digits (`RRGGBB`) become `FF` plus the
/// uppercase digits; eight hex digits become uppercase (borrowed when already
/// uppercase). Anything else is returned unchanged. Only for SpreadsheetML
/// `rgb` attributes, never for DrawingML `a:srgbClr`, which stays `RRGGBB`.
pub(crate) fn to_argb(s: &str) -> std::borrow::Cow<'_, str> {
    use std::borrow::Cow;
    let hex = s.strip_prefix('#').unwrap_or(s);
    if !hex.bytes().all(|b| b.is_ascii_hexdigit()) {
        return Cow::Borrowed(s);
    }
    match hex.len() {
        6 => Cow::Owned(format!("FF{}", hex.to_ascii_uppercase())),
        8 if !hex.bytes().any(|b| b.is_ascii_lowercase()) => Cow::Borrowed(hex),
        8 => Cow::Owned(hex.to_ascii_uppercase()),
        _ => Cow::Borrowed(s),
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn to_argb_normalises_hex_colours() {
        assert_eq!(to_argb("#ff0000"), "FFFF0000");
        assert_eq!(to_argb("ff0000"), "FFFF0000");
        assert_eq!(to_argb("80ff0000"), "80FF0000");
        assert_eq!(to_argb("FFFF0000"), "FFFF0000");
        assert!(matches!(to_argb("FFFF0000"), std::borrow::Cow::Borrowed(_)));
        assert!(matches!(
            to_argb("#FFFF0000"),
            std::borrow::Cow::Borrowed("FFFF0000")
        ));
        assert_eq!(to_argb("red"), "red");
        assert_eq!(to_argb(""), "");
    }

    #[test]
    fn xml_write_err_maps_to_xml_write() {
        let e = xml_write_err(std::io::Error::other("x"));
        assert_eq!(e.code(), "XML_WRITE");
        assert!(e.to_string().contains('x'));
    }

    #[test]
    fn write_xml_decl_emits_prologue() {
        let mut w = quick_xml::Writer::new(Vec::new());
        write_xml_decl(&mut w).unwrap();
        assert_eq!(
            w.into_inner(),
            b"<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
        );
    }
}
