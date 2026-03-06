//! Theme color parser.
//!
//! Parses `xl/theme/theme1.xml` to extract the color scheme from
//! `<a:clrScheme>`. The full theme XML is preserved verbatim through
//! `preserved_entries` for roundtrip fidelity — this module only reads colors.

use core::hint::cold_path;

use quick_xml::events::Event;
use quick_xml::Reader;

use serde::{Deserialize, Serialize};

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
    /// Office 2016 default theme colors.
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
pub fn parse(data: &[u8]) -> Result<ThemeColors> {
    let mut reader = Reader::from_reader(data);
    reader.config_mut().trim_text(true);

    let mut buf = Vec::with_capacity(512);
    let mut colors = ThemeColors::default();

    // We track which color slot we are inside (dk1, lt1, etc.).
    let mut current_slot: Option<&'static str> = None;
    let mut in_clr_scheme = false;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) => {
                let local = e.local_name();
                match local.as_ref() {
                    b"clrScheme" => {
                        in_clr_scheme = true;
                    }
                    b"dk1" if in_clr_scheme => current_slot = Some("dk1"),
                    b"lt1" if in_clr_scheme => current_slot = Some("lt1"),
                    b"dk2" if in_clr_scheme => current_slot = Some("dk2"),
                    b"lt2" if in_clr_scheme => current_slot = Some("lt2"),
                    b"accent1" if in_clr_scheme => current_slot = Some("accent1"),
                    b"accent2" if in_clr_scheme => current_slot = Some("accent2"),
                    b"accent3" if in_clr_scheme => current_slot = Some("accent3"),
                    b"accent4" if in_clr_scheme => current_slot = Some("accent4"),
                    b"accent5" if in_clr_scheme => current_slot = Some("accent5"),
                    b"accent6" if in_clr_scheme => current_slot = Some("accent6"),
                    b"hlink" if in_clr_scheme => current_slot = Some("hlink"),
                    b"folHlink" if in_clr_scheme => current_slot = Some("folHlink"),
                    _ => {}
                }
            }
            Ok(Event::Empty(ref e)) if current_slot.is_some() => {
                let local = e.local_name();
                let color_value = match local.as_ref() {
                    b"sysClr" => e
                        .attributes()
                        .flatten()
                        .find(|a| a.key.local_name().as_ref() == b"lastClr")
                        .map(|a| std::str::from_utf8(&a.value).unwrap_or_default().to_owned()),
                    b"srgbClr" => e
                        .attributes()
                        .flatten()
                        .find(|a| a.key.local_name().as_ref() == b"val")
                        .map(|a| std::str::from_utf8(&a.value).unwrap_or_default().to_owned()),
                    _ => None,
                };

                if let (Some(color), Some(slot)) = (color_value, current_slot) {
                    match slot {
                        "dk1" => colors.dk1 = color,
                        "lt1" => colors.lt1 = color,
                        "dk2" => colors.dk2 = color,
                        "lt2" => colors.lt2 = color,
                        "accent1" => colors.accent1 = color,
                        "accent2" => colors.accent2 = color,
                        "accent3" => colors.accent3 = color,
                        "accent4" => colors.accent4 = color,
                        "accent5" => colors.accent5 = color,
                        "accent6" => colors.accent6 = color,
                        "hlink" => colors.hlink = color,
                        "folHlink" => colors.fol_hlink = color,
                        _ => {}
                    }
                }
            }
            Ok(Event::End(ref e)) => {
                let local = e.local_name();
                match local.as_ref() {
                    b"clrScheme" => {
                        in_clr_scheme = false;
                        current_slot = None;
                    }
                    b"dk1" | b"lt1" | b"dk2" | b"lt2" | b"accent1" | b"accent2"
                    | b"accent3" | b"accent4" | b"accent5" | b"accent6" | b"hlink"
                    | b"folHlink"
                        if in_clr_scheme =>
                    {
                        current_slot = None;
                    }
                    _ => {}
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

#[cfg(test)]
mod tests {
    use super::*;
    use pretty_assertions::assert_eq;

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
}
