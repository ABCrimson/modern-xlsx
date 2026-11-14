//! The SpreadsheetML colour, `CT_Color` (ECMA-376 Part 1, §A.2; its attributes are specified with `bgColor`, §18.8.3): one of an
//! ARGB value, a theme slot, a legacy palette index or the automatic colour,
//! with an optional tint. Fonts, fills, borders, gradient stops, rich-text runs
//! and the worksheet colour elements all share this shape.

use core::hint::cold_path;

use quick_xml::Writer;
use quick_xml::events::{BytesStart, Event};
use serde::{Deserialize, Serialize};

use super::xml::{XsdBuf, attr_str, fmt_xsd_double, parse_attr, xsd_bool};
use super::{to_argb, xml_write_err};
use crate::{ModernXlsxError, Result};

/// A `CT_Color` value. Every attribute is optional; serialized as camelCase
/// JSON with absent attributes omitted.
#[derive(Debug, Clone, Default, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct Color {
    /// `@rgb`: ARGB hex, e.g. `"FF4472C4"`. RGB (`RRGGBB`) and a leading `#`
    /// are accepted and written as uppercase ARGB; any other non-empty value
    /// fails the write with `InvalidWorkbook`.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub rgb: Option<String>,
    /// `@theme`: zero-based index into the theme's colour scheme.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub theme: Option<u32>,
    /// `@indexed`: index into the legacy indexed palette (64 is the system
    /// foreground, 65 the system background).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub indexed: Option<u32>,
    /// `@tint`: lightens (positive) or darkens (negative) the colour, in
    /// -1.0..=1.0.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tint: Option<f64>,
    /// `@auto`: the automatic (system) colour.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub auto: Option<bool>,
}

impl Color {
    /// Read the `CT_Color` attributes of `e` (`<color>`, `<fgColor>`,
    /// `<bgColor>`, ...). Returns `Ok(None)` when the element carries none of
    /// them. An `@auto` that is not an xsd:boolean is treated as absent.
    ///
    /// # Errors
    ///
    /// `XmlParse` for a malformed attribute list, a bad entity, or a `theme`,
    /// `indexed` or `tint` value that does not parse as its schema type.
    pub(crate) fn from_attrs(e: &BytesStart<'_>) -> Result<Option<Self>> {
        let name = e.name();
        let elem = name.as_ref();
        let mut c = Self::default();
        for attr in e.attributes() {
            let attr = attr.map_err(|err| {
                cold_path();
                ModernXlsxError::XmlParse(format!("<{elem}> attributes: {err}"))
            })?;
            match attr.key.local_name().as_ref() {
                "rgb" => c.rgb = Some(attr_str(&attr)?.into_owned()),
                "theme" => c.theme = Some(parse_attr(elem, &attr)?),
                "indexed" => c.indexed = Some(parse_attr(elem, &attr)?),
                "tint" => c.tint = Some(parse_attr(elem, &attr)?),
                "auto" => c.auto = xsd_bool(&attr_str(&attr)?),
                _ => {}
            }
        }
        Ok((c != Self::default()).then_some(c))
    }

    /// Whether any attribute besides `rgb` is set: the colour cannot be held
    /// by a legacy rgb-only string field.
    pub(crate) fn has_non_rgb(&self) -> bool {
        self.theme.is_some() || self.indexed.is_some() || self.tint.is_some() || self.auto.is_some()
    }

    /// Append the attributes to `e` in schema order: `rgb` (as ARGB, see
    /// [`checked_argb`]), `theme`, `indexed`, `tint`, `auto`. An empty `rgb`
    /// is never written. `theme` and `indexed` are integers and `tint` a
    /// finite double, so only `rgb` holds free text.
    ///
    /// # Errors
    ///
    /// `InvalidWorkbook` when `rgb` is not 6 or 8 hex digits; `XmlWrite` when
    /// `tint` is not finite.
    pub(crate) fn push_attrs(&self, e: &mut BytesStart<'_>, buf: &mut XsdBuf) -> Result<()> {
        if let Some(rgb) = self.rgb.as_deref().filter(|s| !s.is_empty()) {
            e.push_attribute(("rgb", &*checked_argb(rgb)?));
        }
        let mut ibuf = itoa::Buffer::new();
        if let Some(theme) = self.theme {
            e.push_attribute(("theme", ibuf.format(theme)));
        }
        if let Some(indexed) = self.indexed {
            e.push_attribute(("indexed", ibuf.format(indexed)));
        }
        if let Some(tint) = self.tint {
            e.push_attribute(("tint", fmt_xsd_double(tint, buf, "color tint")?));
        }
        if let Some(auto) = self.auto {
            e.push_attribute(("auto", if auto { "1" } else { "0" }));
        }
        Ok(())
    }
}

/// Whether `rgb` is a colour the writer can emit: 6 or 8 hex digits after an
/// optional leading `#`, which [`to_argb`] turns into `ST_UnsignedIntHex`
/// ARGB.
pub(crate) fn is_argb(rgb: &str) -> bool {
    let hex = rgb.strip_prefix('#').unwrap_or(rgb);
    matches!(hex.len(), 6 | 8) && hex.bytes().all(|b| b.is_ascii_hexdigit())
}

/// Why `rgb` cannot be written as a colour, naming where it sits when that
/// is known. The value is quoted with its control characters escaped.
pub(crate) fn not_argb_message(rgb: &str, location: Option<&str>) -> String {
    let at = location.map_or_else(String::new, |at| format!(" at {at}"));
    format!("the colour {rgb:?}{at} is not 6 or 8 hex digits (RRGGBB or AARRGGBB)")
}

/// `rgb` normalised by [`to_argb`] to eight uppercase hex digits.
///
/// # Errors
///
/// `InvalidWorkbook` when `rgb` is not 6 or 8 hex digits ([`is_argb`]): such
/// a string cannot be a colour.
pub(crate) fn checked_argb(rgb: &str) -> Result<std::borrow::Cow<'_, str>> {
    if is_argb(rgb) {
        Ok(to_argb(rgb))
    } else {
        cold_path();
        Err(ModernXlsxError::InvalidWorkbook(format!(
            "Write failed: {}",
            not_argb_message(rgb, None)
        )))
    }
}

/// The legacy rgb string and the full colour of a `CT_Color` element: the
/// full colour only when it holds more than an `rgb`.
pub(crate) fn split_color(color: Option<Color>) -> (Option<String>, Option<Color>) {
    match color {
        None => (None, None),
        Some(c) => (c.rgb.clone(), c.has_non_rgb().then_some(c)),
    }
}

/// The colour to write for a legacy rgb string and its sibling full colour,
/// borrowed from the model.
#[derive(Clone, Copy)]
pub(crate) enum ColorOut<'a> {
    Spec(&'a Color),
    Rgb(&'a str),
}

impl<'a> ColorOut<'a> {
    /// Whether this is a full colour with no attribute set. Styles write it as
    /// a bare element; rich-text runs and the worksheet colours write nothing.
    pub(crate) fn is_default_spec(&self) -> bool {
        matches!(self, Self::Spec(s) if **s == Color::default())
    }

    /// The `rgb` this colour writes, before ARGB normalisation; `None` when it
    /// writes none (an empty `rgb` is never written).
    pub(crate) fn rgb(self) -> Option<&'a str> {
        match self {
            Self::Spec(c) => c.rgb.as_deref().filter(|s| !s.is_empty()),
            Self::Rgb(rgb) => Some(rgb),
        }
    }
}

/// Pick the colour to write: the full colour when its `rgb` equals the
/// legacy string (an empty string counts as absent), else the legacy string
/// as plain `rgb`, else nothing. An empty `rgb` is never written.
pub(crate) fn resolve_color<'a>(
    legacy: Option<&'a str>,
    spec: Option<&'a Color>,
) -> Option<ColorOut<'a>> {
    let legacy = legacy.filter(|s| !s.is_empty());
    match spec {
        Some(spec) if spec.rgb.as_deref().filter(|s| !s.is_empty()) == legacy => {
            Some(ColorOut::Spec(spec))
        }
        _ => legacy.map(ColorOut::Rgb),
    }
}

/// Write `<tag .../>` for a resolved colour, its `rgb` normalised to ARGB.
///
/// # Errors
///
/// `InvalidWorkbook` when the `rgb` is not 6 or 8 hex digits
/// ([`checked_argb`]); `XmlWrite` when the writer fails or a `tint` is not
/// finite.
pub(crate) fn write_color<W: std::io::Write>(
    writer: &mut Writer<W>,
    tag: &str,
    color: ColorOut<'_>,
    xbuf: &mut XsdBuf,
) -> Result<()> {
    let mut elem = BytesStart::new(tag);
    match color {
        ColorOut::Spec(c) => c.push_attrs(&mut elem, xbuf)?,
        ColorOut::Rgb(rgb) => elem.push_attribute(("rgb", &*checked_argb(rgb)?)),
    }
    writer
        .write_event(Event::Empty(elem))
        .map_err(xml_write_err)
}

#[cfg(test)]
mod tests {
    use super::*;
    use quick_xml::Reader;
    use quick_xml::events::Event;

    fn first_tag(xml: &str) -> BytesStart<'static> {
        let mut r = Reader::from_str(xml);
        loop {
            match r.read_event().unwrap() {
                Event::Start(e) | Event::Empty(e) => return e.into_owned(),
                Event::Eof => panic!("no start tag in {xml}"),
                _ => {}
            }
        }
    }

    fn written(c: &Color) -> String {
        let mut e = BytesStart::new("color");
        c.push_attrs(&mut e, &mut XsdBuf::new()).unwrap();
        e.attributes_raw().to_owned()
    }

    #[test]
    fn from_attrs_reads_every_attribute() {
        let c = Color::from_attrs(&first_tag(
            r#"<color rgb="FF000000" theme=" 1 " indexed="64" tint="-0.25" auto="1"/>"#,
        ))
        .unwrap()
        .unwrap();
        assert_eq!(
            c,
            Color {
                rgb: Some("FF000000".into()),
                theme: Some(1),
                indexed: Some(64),
                tint: Some(-0.25),
                auto: Some(true),
            }
        );
        assert!(c.has_non_rgb());
    }

    #[test]
    fn from_attrs_none_without_colour_attributes() {
        assert_eq!(Color::from_attrs(&first_tag("<color/>")).unwrap(), None);
        assert_eq!(
            Color::from_attrs(&first_tag(r#"<color auto="maybe"/>"#)).unwrap(),
            None
        );
        let rgb_only = Color::from_attrs(&first_tag(r#"<color rgb="FF00FF00"/>"#))
            .unwrap()
            .unwrap();
        assert!(!rgb_only.has_non_rgb());
    }

    #[test]
    fn from_attrs_rejects_malformed_numbers() {
        for xml in [
            r#"<color theme="x"/>"#,
            r#"<color indexed="-1"/>"#,
            r#"<color tint="pale"/>"#,
            r#"<color rgb="a" rgb="b"/>"#,
        ] {
            let e = Color::from_attrs(&first_tag(xml)).unwrap_err();
            assert!(matches!(e, ModernXlsxError::XmlParse(_)), "{xml}: {e:?}");
        }
    }

    #[test]
    fn push_attrs_writes_schema_order_and_skips_empty_rgb() {
        let c = Color {
            rgb: Some("FF112233".into()),
            theme: Some(4),
            indexed: Some(64),
            tint: Some(0.4),
            auto: Some(true),
        };
        assert_eq!(
            written(&c),
            r#" rgb="FF112233" theme="4" indexed="64" tint="0.4" auto="1""#
        );
        let c = Color {
            rgb: Some(String::new()),
            tint: Some(-1.0),
            auto: Some(false),
            ..Color::default()
        };
        assert_eq!(written(&c), r#" tint="-1" auto="0""#);
    }

    #[test]
    fn push_attrs_rejects_non_finite_tint() {
        let c = Color {
            tint: Some(f64::NAN),
            ..Color::default()
        };
        let e = c
            .push_attrs(&mut BytesStart::new("color"), &mut XsdBuf::new())
            .unwrap_err();
        assert!(matches!(e, ModernXlsxError::XmlWrite(_)), "{e:?}");
    }

    /// `<tag .../>` as `write_color` writes it.
    fn color_xml(color: ColorOut<'_>) -> Result<String> {
        let mut w = Writer::new(Vec::new());
        write_color(&mut w, "color", color, &mut XsdBuf::new())?;
        Ok(String::from_utf8(w.into_inner()).unwrap())
    }

    /// Strings that are not 6 or 8 hex digits once a leading `#` is dropped.
    const NOT_ARGB: [&str; 6] = ["FF0000\u{1}", "<x>", "red", "FFF", "#FF00000", "FF\"0000"];

    #[test]
    fn write_color_writes_six_and_eight_hex_digits_as_argb() {
        for (rgb, argb) in [
            ("ff0000", "FFFF0000"),
            ("80FF0000", "80FF0000"),
            ("#80ff0000", "80FF0000"),
        ] {
            let want = format!(r#"<color rgb="{argb}"/>"#);
            assert_eq!(color_xml(ColorOut::Rgb(rgb)).unwrap(), want);
            let spec = Color {
                rgb: Some(rgb.into()),
                ..Color::default()
            };
            assert_eq!(color_xml(ColorOut::Spec(&spec)).unwrap(), want);
        }
    }

    /// A full colour's `rgb` is written as ARGB by `push_attrs` itself, with
    /// the other attributes untouched: six digits gain `FF`, eight stay, a
    /// leading `#` goes and lower case becomes upper case.
    #[test]
    fn write_color_spec_normalises_rgb_through_push_attrs() {
        for (rgb, argb) in [
            ("a1b2c3", "FFA1B2C3"),
            ("A1B2C3", "FFA1B2C3"),
            ("aabbccdd", "AABBCCDD"),
            ("AABBCCDD", "AABBCCDD"),
            ("#a1b2c3", "FFA1B2C3"),
            ("#AABBCCDD", "AABBCCDD"),
            ("#aabbccdd", "AABBCCDD"),
        ] {
            let spec = Color {
                rgb: Some(rgb.into()),
                theme: Some(2),
                tint: Some(-0.5),
                ..Color::default()
            };
            let want = format!(r#" rgb="{argb}" theme="2" tint="-0.5""#);
            assert_eq!(written(&spec), want, "{rgb}");
            assert_eq!(
                color_xml(ColorOut::Spec(&spec)).unwrap(),
                format!("<color{want}/>"),
                "{rgb}"
            );
        }
    }

    #[test]
    fn write_color_rejects_an_rgb_that_is_not_hex() {
        for bad in NOT_ARGB {
            let spec = Color {
                rgb: Some(bad.into()),
                theme: Some(1),
                ..Color::default()
            };
            for color in [ColorOut::Rgb(bad), ColorOut::Spec(&spec)] {
                let e = color_xml(color).unwrap_err();
                assert_eq!(e.code(), "INVALID_WORKBOOK", "{bad:?}: {e}");
            }
            let e = spec
                .push_attrs(&mut BytesStart::new("color"), &mut XsdBuf::new())
                .unwrap_err();
            assert_eq!(e.code(), "INVALID_WORKBOOK", "{bad:?}: {e}");
        }
    }

    /// A one-sheet workbook whose font 1 has the colour `rgb`.
    fn workbook_with_font_color(rgb: &str) -> crate::WorkbookData {
        let mut wb = crate::WorkbookData {
            sheets: vec![crate::SheetData {
                name: "Sheet1".into(),
                state: None,
                worksheet: crate::ooxml::worksheet::WorksheetXml::default(),
            }],
            ..Default::default()
        };
        wb.styles.fonts.push(crate::ooxml::styles::Font {
            color: Some(rgb.into()),
            ..Default::default()
        });
        wb
    }

    #[test]
    fn a_font_colour_that_is_not_hex_fails_the_write_and_validation() {
        for bad in ["FF0000\u{1}", "<x>"] {
            let wb = workbook_with_font_color(bad);
            let e = wb.styles.to_xml().unwrap_err();
            assert_eq!(e.code(), "INVALID_WORKBOOK", "{bad:?}: {e}");
            let e = crate::writer::write_xlsx(&wb).unwrap_err();
            assert_eq!(e.code(), "INVALID_WORKBOOK", "{bad:?}: {e}");
            assert!(e.to_string().contains("styles.fonts[1].color"), "{e}");
            let report = crate::validate::validate_workbook(&wb);
            assert!(!report.is_valid);
            let issue = report
                .issues
                .iter()
                .find(|i| i.location == "styles.fonts[1].color")
                .expect("a colour issue");
            assert_eq!(issue.severity, crate::validate::Severity::Error);
        }
        for (rgb, argb) in [("ff0000", "FFFF0000"), ("80FF0000", "80FF0000")] {
            let wb = workbook_with_font_color(rgb);
            let xml = String::from_utf8(wb.styles.to_xml().unwrap()).unwrap();
            assert!(xml.contains(&format!(r#"<color rgb="{argb}"/>"#)), "{xml}");
            crate::writer::write_xlsx(&wb).unwrap();
            let report = crate::validate::validate_workbook(&wb);
            assert!(
                report.issues.iter().all(|i| !i.location.contains("color")),
                "{:?}",
                report.issues
            );
        }
    }

    #[test]
    fn json_is_camel_case_and_omits_absent_fields() {
        let c = Color {
            theme: Some(1),
            tint: Some(-0.25),
            ..Color::default()
        };
        assert_eq!(
            serde_json::to_string(&c).unwrap(),
            r#"{"theme":1,"tint":-0.25}"#
        );
        let back: Color = serde_json::from_str(r#"{"indexed":64,"auto":true}"#).unwrap();
        assert_eq!(back.indexed, Some(64));
        assert_eq!(back.auto, Some(true));
    }
}
