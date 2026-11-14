//! Theme color parser.
//!
//! Parses `xl/theme/theme1.xml` to extract the color scheme from
//! `<a:clrScheme>`. The reader keeps the theme part verbatim in
//! `preserved_entries` for roundtrip fidelity, except one byte-identical to the
//! theme the writer generates from the extracted colors; this module reads
//! colors and also holds [`generate_theme_xml`], which builds that generated
//! theme part (the reader compares a theme part with it).

use core::hint::cold_path;

use quick_xml::Reader;
use quick_xml::events::{BytesStart, Event};

use serde::{Deserialize, Serialize};

use super::xml::attr_str;
use crate::{ModernXlsxError, Result};

/// The 12 standard theme colors extracted from `<a:clrScheme>`.
///
/// Each color is stored as a 6-character RGB hex string (e.g. `"000000"`,
/// `"FFFFFF"`). For `sysClr` elements the `lastClr` attribute is used;
/// for `srgbClr` elements the `val` attribute is used.
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ThemeColors {
    pub dk1: String,
    pub lt1: String,
    pub dk2: String,
    pub lt2: String,
    pub accent1: String,
    pub accent2: String,
    pub accent3: String,
    pub accent4: String,
    pub accent5: String,
    pub accent6: String,
    pub hlink: String,
    pub fol_hlink: String,
}

impl Default for ThemeColors {
    /// The Office theme colours the writer generates when none are set
    /// (`accent1` 4472C4).
    fn default() -> Self {
        Self {
            dk1: "000000".to_string(),
            lt1: "FFFFFF".to_string(),
            dk2: "44546A".to_string(),
            lt2: "E7E6E6".to_string(),
            accent1: "4472C4".to_string(),
            accent2: "ED7D31".to_string(),
            accent3: "A5A5A5".to_string(),
            accent4: "FFC000".to_string(),
            accent5: "5B9BD5".to_string(),
            accent6: "70AD47".to_string(),
            hlink: "0563C1".to_string(),
            fol_hlink: "954F72".to_string(),
        }
    }
}

/// One of the 12 `<a:clrScheme>` colour slots.
#[derive(Clone, Copy)]
enum Slot {
    Dk1,
    Lt1,
    Dk2,
    Lt2,
    Accent1,
    Accent2,
    Accent3,
    Accent4,
    Accent5,
    Accent6,
    Hlink,
    FolHlink,
}

impl Slot {
    /// The slot named by a `<a:clrScheme>` child element, if any.
    fn from_local_name(name: &str) -> Option<Self> {
        Some(match name {
            "dk1" => Self::Dk1,
            "lt1" => Self::Lt1,
            "dk2" => Self::Dk2,
            "lt2" => Self::Lt2,
            "accent1" => Self::Accent1,
            "accent2" => Self::Accent2,
            "accent3" => Self::Accent3,
            "accent4" => Self::Accent4,
            "accent5" => Self::Accent5,
            "accent6" => Self::Accent6,
            "hlink" => Self::Hlink,
            "folHlink" => Self::FolHlink,
            _ => return None,
        })
    }

    /// The `ThemeColors` field this slot fills.
    fn field(self, colors: &mut ThemeColors) -> &mut String {
        match self {
            Self::Dk1 => &mut colors.dk1,
            Self::Lt1 => &mut colors.lt1,
            Self::Dk2 => &mut colors.dk2,
            Self::Lt2 => &mut colors.lt2,
            Self::Accent1 => &mut colors.accent1,
            Self::Accent2 => &mut colors.accent2,
            Self::Accent3 => &mut colors.accent3,
            Self::Accent4 => &mut colors.accent4,
            Self::Accent5 => &mut colors.accent5,
            Self::Accent6 => &mut colors.accent6,
            Self::Hlink => &mut colors.hlink,
            Self::FolHlink => &mut colors.fol_hlink,
        }
    }
}

/// The colour held by a `<a:srgbClr val>` or `<a:sysClr lastClr>` element, or
/// `None` for any other element (or a `sysClr` without `lastClr`, whose `val`
/// is a system-colour name, not a colour).
///
/// Called from both the `Start` and the `Empty` arm: a colour element with
/// child transforms (`<a:lumMod/>`) is a `Start`, a bare one is an `Empty`.
fn extract_color(e: &BytesStart<'_>) -> Result<Option<String>> {
    let wanted = match e.local_name().as_ref() {
        "srgbClr" => "val",
        "sysClr" => "lastClr",
        _ => return Ok(None),
    };
    for attr in e.attributes() {
        let attr = attr.map_err(|err| {
            cold_path();
            ModernXlsxError::XmlParse(format!("error parsing theme XML: {err}"))
        })?;
        if attr.key.local_name().as_ref() == wanted {
            return Ok(Some(attr_str(&attr)?.into_owned()));
        }
    }
    Ok(None)
}

/// Parse `xl/theme/theme1.xml` and extract the color scheme.
///
/// The expected structure is:
/// ```xml
/// <a:theme>
///   <a:themeElements>
///     <a:clrScheme name="Office">
///       <a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>
///       <a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>
///       <a:dk2><a:srgbClr val="44546A"/></a:dk2>
///       ...
///     </a:clrScheme>
///   </a:themeElements>
/// </a:theme>
/// ```
///
/// Only the first `<a:clrScheme>` (the one in `<a:themeElements>`) is read;
/// the schemes of `<a:extraClrSchemeLst>` are ignored.
///
/// A colour element may carry child transforms (`<a:srgbClr val="123456">
/// <a:lumMod val="75000"/></a:srgbClr>`); the base colour is read either way.
///
/// # Errors
///
/// `XmlParse` for malformed XML or an attribute that does not decode.
pub fn parse(data: &[u8]) -> Result<ThemeColors> {
    let mut reader = Reader::from_reader(data);
    reader.config_mut().trim_text(true);

    let mut buf = Vec::with_capacity(512);
    let mut colors = ThemeColors::default();

    // The colour slot we are inside (dk1, lt1, etc.).
    let mut current_slot: Option<Slot> = None;
    let mut in_clr_scheme = false;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) => {
                let local = e.local_name();
                if local.as_ref() == "clrScheme" {
                    in_clr_scheme = true;
                } else if in_clr_scheme && let Some(slot) = Slot::from_local_name(local.as_ref()) {
                    current_slot = Some(slot);
                } else if let Some(slot) = current_slot
                    && let Some(color) = extract_color(e)?
                {
                    *slot.field(&mut colors) = color;
                }
            }
            Ok(Event::Empty(ref e)) => {
                if let Some(slot) = current_slot
                    && let Some(color) = extract_color(e)?
                {
                    *slot.field(&mut colors) = color;
                }
            }
            Ok(Event::End(ref e)) => {
                let local = e.local_name();
                if local.as_ref() == "clrScheme" {
                    break;
                } else if in_clr_scheme && Slot::from_local_name(local.as_ref()).is_some() {
                    current_slot = None;
                }
            }
            Ok(Event::Eof) => break,
            Err(err) => {
                cold_path();
                return Err(ModernXlsxError::XmlParse(format!(
                    "error parsing theme XML: {err}"
                )));
            }
            _ => {}
        }
        buf.clear();
    }

    Ok(colors)
}

// ---------------------------------------------------------------------------
// Theme XML generation
// ---------------------------------------------------------------------------

/// Builds the `xl/theme/theme1.xml` part the writer emits when a workbook
/// keeps no theme part of its own.
///
/// The theme is named "Office Theme", with "Office" colour, font and format
/// schemes: the twelve colours of `colors`, or of `ThemeColors::default()`
/// (accent1 4472C4) when `None`, written as given (`dk1` and `lt1` as
/// `sysClr` windowText and window with the colour as `lastClr`, the others as
/// `srgbClr`); Calibri Light as the major and Calibri as the minor Latin
/// font.
pub fn generate_theme_xml(colors: Option<&ThemeColors>) -> Vec<u8> {
    let c = colors.cloned().unwrap_or_default();

    let mut xml = Vec::with_capacity(4096);
    xml.extend_from_slice(b"<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n");
    xml.extend_from_slice(
        b"<a:theme xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" name=\"Office Theme\">"
    );

    // themeElements
    xml.extend_from_slice(b"<a:themeElements>");

    // clrScheme
    xml.extend_from_slice(b"<a:clrScheme name=\"Office\">");
    write_theme_color(&mut xml, "dk1", &c.dk1, true);
    write_theme_color(&mut xml, "lt1", &c.lt1, true);
    write_theme_color(&mut xml, "dk2", &c.dk2, false);
    write_theme_color(&mut xml, "lt2", &c.lt2, false);
    write_theme_color(&mut xml, "accent1", &c.accent1, false);
    write_theme_color(&mut xml, "accent2", &c.accent2, false);
    write_theme_color(&mut xml, "accent3", &c.accent3, false);
    write_theme_color(&mut xml, "accent4", &c.accent4, false);
    write_theme_color(&mut xml, "accent5", &c.accent5, false);
    write_theme_color(&mut xml, "accent6", &c.accent6, false);
    write_theme_color(&mut xml, "hlink", &c.hlink, false);
    write_theme_color(&mut xml, "folHlink", &c.fol_hlink, false);
    xml.extend_from_slice(b"</a:clrScheme>");

    // fontScheme
    xml.extend_from_slice(b"<a:fontScheme name=\"Office\">");
    xml.extend_from_slice(
        b"<a:majorFont><a:latin typeface=\"Calibri Light\"/><a:ea typeface=\"\"/><a:cs typeface=\"\"/></a:majorFont>"
    );
    xml.extend_from_slice(
        b"<a:minorFont><a:latin typeface=\"Calibri\"/><a:ea typeface=\"\"/><a:cs typeface=\"\"/></a:minorFont>"
    );
    xml.extend_from_slice(b"</a:fontScheme>");

    // fmtScheme
    xml.extend_from_slice(b"<a:fmtScheme name=\"Office\">");

    // fillStyleLst
    xml.extend_from_slice(b"<a:fillStyleLst>");
    xml.extend_from_slice(b"<a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill>");
    xml.extend_from_slice(b"<a:gradFill rotWithShape=\"1\"><a:gsLst>");
    xml.extend_from_slice(
        b"<a:gs pos=\"0\"><a:schemeClr val=\"phClr\"><a:lumMod val=\"110000\"/><a:satMod val=\"105000\"/><a:tint val=\"67000\"/></a:schemeClr></a:gs>"
    );
    xml.extend_from_slice(
        b"<a:gs pos=\"50000\"><a:schemeClr val=\"phClr\"><a:lumMod val=\"105000\"/><a:satMod val=\"103000\"/><a:tint val=\"73000\"/></a:schemeClr></a:gs>"
    );
    xml.extend_from_slice(
        b"<a:gs pos=\"100000\"><a:schemeClr val=\"phClr\"><a:lumMod val=\"105000\"/><a:satMod val=\"109000\"/><a:tint val=\"81000\"/></a:schemeClr></a:gs>"
    );
    xml.extend_from_slice(b"</a:gsLst><a:lin ang=\"5400000\" scaled=\"0\"/></a:gradFill>");
    xml.extend_from_slice(b"<a:gradFill rotWithShape=\"1\"><a:gsLst>");
    xml.extend_from_slice(
        b"<a:gs pos=\"0\"><a:schemeClr val=\"phClr\"><a:satMod val=\"103000\"/><a:lumMod val=\"102000\"/><a:tint val=\"94000\"/></a:schemeClr></a:gs>"
    );
    xml.extend_from_slice(
        b"<a:gs pos=\"50000\"><a:schemeClr val=\"phClr\"><a:satMod val=\"110000\"/><a:lumMod val=\"100000\"/><a:shade val=\"100000\"/></a:schemeClr></a:gs>"
    );
    xml.extend_from_slice(
        b"<a:gs pos=\"100000\"><a:schemeClr val=\"phClr\"><a:lumMod val=\"99000\"/><a:satMod val=\"120000\"/><a:shade val=\"78000\"/></a:schemeClr></a:gs>"
    );
    xml.extend_from_slice(b"</a:gsLst><a:lin ang=\"5400000\" scaled=\"0\"/></a:gradFill>");
    xml.extend_from_slice(b"</a:fillStyleLst>");

    // lnStyleLst
    xml.extend_from_slice(b"<a:lnStyleLst>");
    xml.extend_from_slice(
        b"<a:ln w=\"6350\" cap=\"flat\" cmpd=\"sng\" algn=\"ctr\"><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill><a:prstDash val=\"solid\"/><a:miter lim=\"800000\"/></a:ln>"
    );
    xml.extend_from_slice(
        b"<a:ln w=\"12700\" cap=\"flat\" cmpd=\"sng\" algn=\"ctr\"><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill><a:prstDash val=\"solid\"/><a:miter lim=\"800000\"/></a:ln>"
    );
    xml.extend_from_slice(
        b"<a:ln w=\"19050\" cap=\"flat\" cmpd=\"sng\" algn=\"ctr\"><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill><a:prstDash val=\"solid\"/><a:miter lim=\"800000\"/></a:ln>"
    );
    xml.extend_from_slice(b"</a:lnStyleLst>");

    // effectStyleLst
    xml.extend_from_slice(b"<a:effectStyleLst>");
    xml.extend_from_slice(b"<a:effectStyle><a:effectLst/></a:effectStyle>");
    xml.extend_from_slice(b"<a:effectStyle><a:effectLst/></a:effectStyle>");
    xml.extend_from_slice(
        b"<a:effectStyle><a:effectLst><a:outerShdw blurRad=\"57150\" dist=\"19050\" dir=\"5400000\" algn=\"ctr\" rotWithShape=\"0\"><a:srgbClr val=\"000000\"><a:alpha val=\"63000\"/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle>"
    );
    xml.extend_from_slice(b"</a:effectStyleLst>");

    // bgFillStyleLst
    xml.extend_from_slice(b"<a:bgFillStyleLst>");
    xml.extend_from_slice(b"<a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill>");
    xml.extend_from_slice(b"<a:solidFill><a:schemeClr val=\"phClr\"><a:tint val=\"95000\"/><a:satMod val=\"170000\"/></a:schemeClr></a:solidFill>");
    xml.extend_from_slice(b"<a:gradFill rotWithShape=\"1\"><a:gsLst>");
    xml.extend_from_slice(
        b"<a:gs pos=\"0\"><a:schemeClr val=\"phClr\"><a:tint val=\"93000\"/><a:satMod val=\"150000\"/><a:shade val=\"98000\"/><a:lumMod val=\"102000\"/></a:schemeClr></a:gs>"
    );
    xml.extend_from_slice(
        b"<a:gs pos=\"50000\"><a:schemeClr val=\"phClr\"><a:tint val=\"98000\"/><a:satMod val=\"130000\"/><a:shade val=\"90000\"/><a:lumMod val=\"103000\"/></a:schemeClr></a:gs>"
    );
    xml.extend_from_slice(
        b"<a:gs pos=\"100000\"><a:schemeClr val=\"phClr\"><a:shade val=\"63000\"/><a:satMod val=\"120000\"/></a:schemeClr></a:gs>"
    );
    xml.extend_from_slice(b"</a:gsLst><a:lin ang=\"5400000\" scaled=\"0\"/></a:gradFill>");
    xml.extend_from_slice(b"</a:bgFillStyleLst>");

    xml.extend_from_slice(b"</a:fmtScheme>");
    xml.extend_from_slice(b"</a:themeElements>");

    // Empty objectDefaults and extraClrSchemeLst.
    xml.extend_from_slice(b"<a:objectDefaults/>");
    xml.extend_from_slice(b"<a:extraClrSchemeLst/>");

    xml.extend_from_slice(b"</a:theme>");
    xml
}

/// Write a single theme color element.
fn write_theme_color(xml: &mut Vec<u8>, name: &str, color: &str, use_sys_clr: bool) {
    xml.extend_from_slice(b"<a:");
    xml.extend_from_slice(name.as_bytes());
    xml.push(b'>');

    if use_sys_clr {
        let sys_val = if name == "dk1" {
            "windowText"
        } else {
            "window"
        };
        xml.extend_from_slice(b"<a:sysClr val=\"");
        xml.extend_from_slice(sys_val.as_bytes());
        xml.extend_from_slice(b"\" lastClr=\"");
        xml.extend_from_slice(color.as_bytes());
        xml.extend_from_slice(b"\"/>");
    } else {
        xml.extend_from_slice(b"<a:srgbClr val=\"");
        xml.extend_from_slice(color.as_bytes());
        xml.extend_from_slice(b"\"/>");
    }

    xml.extend_from_slice(b"</a:");
    xml.extend_from_slice(name.as_bytes());
    xml.push(b'>');
}

#[cfg(test)]
mod tests {
    use super::*;
    use pretty_assertions::assert_eq;

    #[test]
    fn test_parse_ignores_extra_clr_scheme_list() {
        let xml = r#"<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="T">
  <a:themeElements>
    <a:clrScheme name="Office"><a:accent1><a:srgbClr val="4472C4"/></a:accent1></a:clrScheme>
  </a:themeElements>
  <a:extraClrSchemeLst><a:extraClrScheme><a:clrScheme name="x"><a:accent1><a:srgbClr val="00FF00"/></a:accent1></a:clrScheme></a:extraClrScheme></a:extraClrSchemeLst>
</a:theme>"#;
        assert_eq!(parse(xml.as_bytes()).unwrap().accent1, "4472C4");
    }

    #[test]
    fn test_parse_theme_colors_srgb() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme">
  <a:themeElements>
    <a:clrScheme name="Office">
      <a:dk1><a:srgbClr val="000000"/></a:dk1>
      <a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>
      <a:dk2><a:srgbClr val="44546A"/></a:dk2>
      <a:lt2><a:srgbClr val="E7E6E6"/></a:lt2>
      <a:accent1><a:srgbClr val="4472C4"/></a:accent1>
      <a:accent2><a:srgbClr val="ED7D31"/></a:accent2>
      <a:accent3><a:srgbClr val="A5A5A5"/></a:accent3>
      <a:accent4><a:srgbClr val="FFC000"/></a:accent4>
      <a:accent5><a:srgbClr val="5B9BD5"/></a:accent5>
      <a:accent6><a:srgbClr val="70AD47"/></a:accent6>
      <a:hlink><a:srgbClr val="0563C1"/></a:hlink>
      <a:folHlink><a:srgbClr val="954F72"/></a:folHlink>
    </a:clrScheme>
  </a:themeElements>
</a:theme>"#;

        let colors = parse(xml.as_bytes()).unwrap();
        assert_eq!(colors.dk1, "000000");
        assert_eq!(colors.lt1, "FFFFFF");
        assert_eq!(colors.dk2, "44546A");
        assert_eq!(colors.lt2, "E7E6E6");
        assert_eq!(colors.accent1, "4472C4");
        assert_eq!(colors.accent2, "ED7D31");
        assert_eq!(colors.accent3, "A5A5A5");
        assert_eq!(colors.accent4, "FFC000");
        assert_eq!(colors.accent5, "5B9BD5");
        assert_eq!(colors.accent6, "70AD47");
        assert_eq!(colors.hlink, "0563C1");
        assert_eq!(colors.fol_hlink, "954F72");
    }

    #[test]
    fn test_parse_theme_colors_sysclr() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme">
  <a:themeElements>
    <a:clrScheme name="Office">
      <a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>
      <a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>
      <a:dk2><a:srgbClr val="44546A"/></a:dk2>
      <a:lt2><a:srgbClr val="E7E6E6"/></a:lt2>
      <a:accent1><a:srgbClr val="4472C4"/></a:accent1>
      <a:accent2><a:srgbClr val="ED7D31"/></a:accent2>
      <a:accent3><a:srgbClr val="A5A5A5"/></a:accent3>
      <a:accent4><a:srgbClr val="FFC000"/></a:accent4>
      <a:accent5><a:srgbClr val="5B9BD5"/></a:accent5>
      <a:accent6><a:srgbClr val="70AD47"/></a:accent6>
      <a:hlink><a:srgbClr val="0563C1"/></a:hlink>
      <a:folHlink><a:srgbClr val="954F72"/></a:folHlink>
    </a:clrScheme>
  </a:themeElements>
</a:theme>"#;

        let colors = parse(xml.as_bytes()).unwrap();
        assert_eq!(colors.dk1, "000000");
        assert_eq!(colors.lt1, "FFFFFF");
    }

    #[test]
    fn test_parse_theme_colors_mixed() {
        // Minimal theme with only a few colors — the rest should keep defaults.
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <a:themeElements>
    <a:clrScheme name="Custom">
      <a:dk1><a:srgbClr val="111111"/></a:dk1>
      <a:lt1><a:srgbClr val="EEEEEE"/></a:lt1>
    </a:clrScheme>
  </a:themeElements>
</a:theme>"#;

        let colors = parse(xml.as_bytes()).unwrap();
        assert_eq!(colors.dk1, "111111");
        assert_eq!(colors.lt1, "EEEEEE");
        // Other colors should remain at defaults.
        assert_eq!(colors.dk2, "44546A");
        assert_eq!(colors.accent1, "4472C4");
    }

    #[test]
    fn test_default_theme_colors() {
        let colors = ThemeColors::default();
        assert_eq!(colors.dk1, "000000");
        assert_eq!(colors.lt1, "FFFFFF");
        assert_eq!(colors.fol_hlink, "954F72");
    }

    #[test]
    fn start_form_srgb_with_child_transform() {
        let xml = r#"<a:theme xmlns:a="x"><a:themeElements><a:clrScheme name="C"><a:accent1><a:srgbClr val="123456"><a:lumMod val="75000"/></a:srgbClr></a:accent1></a:clrScheme></a:themeElements></a:theme>"#;
        assert_eq!(parse(xml.as_bytes()).unwrap().accent1, "123456");
    }

    #[test]
    fn sysclr_last_clr_is_used() {
        let xml = r#"<a:theme xmlns:a="x"><a:themeElements><a:clrScheme name="C"><a:dk1><a:sysClr val="windowText" lastClr="010101"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FEFEFE"/></a:lt1></a:clrScheme></a:themeElements></a:theme>"#;
        let colors = parse(xml.as_bytes()).unwrap();
        assert_eq!(colors.dk1, "010101");
        assert_eq!(colors.lt1, "FEFEFE");
    }

    #[test]
    fn malformed_theme_is_xml_parse_error() {
        assert!(matches!(
            parse(b"<a:theme><a></b></a:theme>"),
            Err(ModernXlsxError::XmlParse(_))
        ));
    }

    // -- Theme XML generation --

    #[test]
    fn test_generate_theme_xml() {
        let xml = generate_theme_xml(None);
        let s = std::str::from_utf8(&xml).unwrap();
        assert!(s.contains("a:theme"));
        assert!(s.contains("a:clrScheme"));
        assert!(s.contains("a:fontScheme"));
        assert!(s.contains("a:fmtScheme"));
        assert!(s.contains("windowText"));
        assert!(s.contains("000000"));
        assert!(s.contains("<a:latin typeface=\"Calibri Light\"/>"));
        assert!(s.contains("<a:minorFont><a:latin typeface=\"Calibri\"/>"));
        assert!(s.contains("4472C4"));
        assert!(!s.contains("Aptos"));
    }

    #[test]
    fn test_generate_theme_xml_custom_colors() {
        let colors = ThemeColors {
            dk1: "111111".into(),
            lt1: "EEEEEE".into(),
            ..ThemeColors::default()
        };
        let xml = generate_theme_xml(Some(&colors));
        let s = std::str::from_utf8(&xml).unwrap();
        assert!(s.contains("111111"));
        assert!(s.contains("EEEEEE"));
    }
}
