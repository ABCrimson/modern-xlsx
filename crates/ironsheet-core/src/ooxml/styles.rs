use quick_xml::events::{BytesDecl, BytesEnd, BytesStart, Event};
use quick_xml::{Reader, Writer};
use serde::{Deserialize, Serialize};

use crate::{IronsheetError, Result};

const SPREADSHEET_NS: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

// ---------------------------------------------------------------------------
// Data structures
// ---------------------------------------------------------------------------

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct NumFmt {
    pub id: u32,
    pub format_code: String,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct Font {
    pub name: Option<String>,
    pub size: Option<f64>,
    pub bold: bool,
    pub italic: bool,
    pub underline: bool,
    pub strike: bool,
    /// ARGB hex colour string, e.g. "FF000000".
    pub color: Option<String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct Fill {
    pub pattern_type: String,
    pub fg_color: Option<String>,
    pub bg_color: Option<String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct Border {
    pub left: Option<BorderSide>,
    pub right: Option<BorderSide>,
    pub top: Option<BorderSide>,
    pub bottom: Option<BorderSide>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct BorderSide {
    pub style: String,
    pub color: Option<String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct CellXf {
    pub num_fmt_id: u32,
    pub font_id: u32,
    pub fill_id: u32,
    pub border_id: u32,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct Styles {
    pub num_fmts: Vec<NumFmt>,
    pub fonts: Vec<Font>,
    pub fills: Vec<Fill>,
    pub borders: Vec<Border>,
    pub cell_xfs: Vec<CellXf>,
}

// ---------------------------------------------------------------------------
// Parser state
// ---------------------------------------------------------------------------

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum Section {
    None,
    NumFmts,
    Fonts,
    Fills,
    Borders,
    CellXfs,
}

/// Tracks which border child element we are inside.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum BorderChild {
    None,
    Left,
    Right,
    Top,
    Bottom,
}

impl Styles {
    // -----------------------------------------------------------------------
    // Factory
    // -----------------------------------------------------------------------

    /// Return the minimal valid styles that Excel requires.
    pub fn default_styles() -> Self {
        Styles {
            num_fmts: Vec::new(),
            fonts: vec![Font {
                name: Some("Aptos".to_owned()),
                size: Some(11.0),
                ..Font::default()
            }],
            fills: vec![
                Fill {
                    pattern_type: "none".to_owned(),
                    ..Fill::default()
                },
                Fill {
                    pattern_type: "gray125".to_owned(),
                    ..Fill::default()
                },
            ],
            borders: vec![Border::default()],
            cell_xfs: vec![CellXf::default()],
        }
    }

    // -----------------------------------------------------------------------
    // Parser
    // -----------------------------------------------------------------------

    /// Parse a `styles.xml` document from raw XML bytes.
    pub fn parse(data: &[u8]) -> Result<Self> {
        let mut reader = Reader::from_reader(data);
        reader.config_mut().trim_text(true);

        let mut buf = Vec::new();

        let mut num_fmts: Vec<NumFmt> = Vec::new();
        let mut fonts: Vec<Font> = Vec::new();
        let mut fills: Vec<Fill> = Vec::new();
        let mut borders: Vec<Border> = Vec::new();
        let mut cell_xfs: Vec<CellXf> = Vec::new();

        let mut section = Section::None;

        // Font parsing state.
        let mut current_font = Font::default();
        let mut in_font = false;

        // Fill parsing state.
        let mut current_fill = Fill::default();
        let mut in_fill = false;

        // Border parsing state.
        let mut current_border = Border::default();
        let mut in_border = false;
        let mut border_child = BorderChild::None;
        let mut current_border_side: Option<BorderSide> = None;

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) => {
                    let local = e.local_name();
                    match local.as_ref() {
                        // ----- section openers -----
                        b"numFmts" => section = Section::NumFmts,
                        b"fonts" => section = Section::Fonts,
                        b"fills" => section = Section::Fills,
                        b"borders" => section = Section::Borders,
                        b"cellXfs" => section = Section::CellXfs,

                        // ----- font children -----
                        b"font" if section == Section::Fonts => {
                            current_font = Font::default();
                            in_font = true;
                        }

                        // ----- fill children -----
                        b"fill" if section == Section::Fills => {
                            current_fill = Fill::default();
                            in_fill = true;
                        }
                        b"patternFill" if in_fill => {
                            for attr in e.attributes().flatten() {
                                if attr.key.local_name().as_ref() == b"patternType" {
                                    current_fill.pattern_type =
                                        std::str::from_utf8(&attr.value).unwrap_or_default().to_owned();
                                }
                            }
                        }

                        // ----- border children -----
                        b"border" if section == Section::Borders => {
                            current_border = Border::default();
                            in_border = true;
                        }
                        b"left" if in_border => {
                            border_child = BorderChild::Left;
                            current_border_side = parse_border_side_attrs(e);
                        }
                        b"right" if in_border => {
                            border_child = BorderChild::Right;
                            current_border_side = parse_border_side_attrs(e);
                        }
                        b"top" if in_border => {
                            border_child = BorderChild::Top;
                            current_border_side = parse_border_side_attrs(e);
                        }
                        b"bottom" if in_border => {
                            border_child = BorderChild::Bottom;
                            current_border_side = parse_border_side_attrs(e);
                        }

                        // colour inside a border side
                        b"color" if border_child != BorderChild::None => {
                            if let Some(ref mut side) = current_border_side {
                                for attr in e.attributes().flatten() {
                                    if attr.key.local_name().as_ref() == b"rgb" {
                                        side.color = Some(
                                            std::str::from_utf8(&attr.value).unwrap_or_default().to_owned(),
                                        );
                                    }
                                }
                            }
                        }

                        // ----- xf (start variant, unlikely but safe) -----
                        b"xf" if section == Section::CellXfs => {
                            cell_xfs.push(parse_xf_attrs(e));
                        }

                        _ => {}
                    }
                }

                Ok(Event::Empty(ref e)) => {
                    let local = e.local_name();
                    match local.as_ref() {
                        // ----- numFmt is always self-closing -----
                        b"numFmt" if section == Section::NumFmts => {
                            let mut id: u32 = 0;
                            let mut code = String::new();
                            for attr in e.attributes().flatten() {
                                match attr.key.local_name().as_ref() {
                                    b"numFmtId" => {
                                        id = std::str::from_utf8(&attr.value).unwrap_or_default()
                                            .parse()
                                            .unwrap_or(0);
                                    }
                                    b"formatCode" => {
                                        code = std::str::from_utf8(&attr.value).unwrap_or_default().to_owned();
                                    }
                                    _ => {}
                                }
                            }
                            num_fmts.push(NumFmt {
                                id,
                                format_code: code,
                            });
                        }

                        // ----- font child elements (self-closing) -----
                        b"b" if in_font => current_font.bold = true,
                        b"i" if in_font => current_font.italic = true,
                        b"u" if in_font => current_font.underline = true,
                        b"strike" if in_font => current_font.strike = true,
                        b"sz" if in_font => {
                            for attr in e.attributes().flatten() {
                                if attr.key.local_name().as_ref() == b"val" {
                                    current_font.size = std::str::from_utf8(&attr.value).unwrap_or_default()
                                        .parse()
                                        .ok();
                                }
                            }
                        }
                        b"name" if in_font => {
                            for attr in e.attributes().flatten() {
                                if attr.key.local_name().as_ref() == b"val" {
                                    current_font.name = Some(
                                        std::str::from_utf8(&attr.value).unwrap_or_default().to_owned(),
                                    );
                                }
                            }
                        }
                        b"color" if in_font => {
                            for attr in e.attributes().flatten() {
                                if attr.key.local_name().as_ref() == b"rgb" {
                                    current_font.color = Some(
                                        std::str::from_utf8(&attr.value).unwrap_or_default().to_owned(),
                                    );
                                }
                            }
                        }

                        // ----- patternFill as self-closing -----
                        b"patternFill" if in_fill => {
                            for attr in e.attributes().flatten() {
                                if attr.key.local_name().as_ref() == b"patternType" {
                                    current_fill.pattern_type =
                                        std::str::from_utf8(&attr.value).unwrap_or_default().to_owned();
                                }
                            }
                        }

                        // ----- fgColor / bgColor inside fill -----
                        b"fgColor" if in_fill => {
                            for attr in e.attributes().flatten() {
                                if attr.key.local_name().as_ref() == b"rgb" {
                                    current_fill.fg_color = Some(
                                        std::str::from_utf8(&attr.value).unwrap_or_default().to_owned(),
                                    );
                                }
                            }
                        }
                        b"bgColor" if in_fill => {
                            for attr in e.attributes().flatten() {
                                if attr.key.local_name().as_ref() == b"rgb" {
                                    current_fill.bg_color = Some(
                                        std::str::from_utf8(&attr.value).unwrap_or_default().to_owned(),
                                    );
                                }
                            }
                        }

                        // ----- empty border sides (e.g. <left/>) -----
                        b"left" if in_border => {
                            let side = parse_border_side_attrs(e);
                            if side.is_some() {
                                current_border.left = side;
                            }
                        }
                        b"right" if in_border => {
                            let side = parse_border_side_attrs(e);
                            if side.is_some() {
                                current_border.right = side;
                            }
                        }
                        b"top" if in_border => {
                            let side = parse_border_side_attrs(e);
                            if side.is_some() {
                                current_border.top = side;
                            }
                        }
                        b"bottom" if in_border => {
                            let side = parse_border_side_attrs(e);
                            if side.is_some() {
                                current_border.bottom = side;
                            }
                        }

                        // colour inside a border side (self-closing)
                        b"color" if border_child != BorderChild::None => {
                            if let Some(ref mut side) = current_border_side {
                                for attr in e.attributes().flatten() {
                                    if attr.key.local_name().as_ref() == b"rgb" {
                                        side.color = Some(
                                            std::str::from_utf8(&attr.value).unwrap_or_default().to_owned(),
                                        );
                                    }
                                }
                            }
                        }

                        // ----- xf (self-closing) -----
                        b"xf" if section == Section::CellXfs => {
                            cell_xfs.push(parse_xf_attrs(e));
                        }

                        _ => {}
                    }
                }

                Ok(Event::End(ref e)) => {
                    let local = e.local_name();
                    match local.as_ref() {
                        // ----- section closers -----
                        b"numFmts" => section = Section::None,
                        b"fonts" => section = Section::None,
                        b"fills" => section = Section::None,
                        b"borders" => section = Section::None,
                        b"cellXfs" => section = Section::None,

                        // ----- font end -----
                        b"font" if in_font => {
                            fonts.push(std::mem::take(&mut current_font));
                            in_font = false;
                        }

                        // ----- fill end -----
                        b"fill" if in_fill => {
                            fills.push(std::mem::take(&mut current_fill));
                            in_fill = false;
                        }

                        // ----- border side ends -----
                        b"left" if border_child == BorderChild::Left => {
                            current_border.left = current_border_side.take();
                            border_child = BorderChild::None;
                        }
                        b"right" if border_child == BorderChild::Right => {
                            current_border.right = current_border_side.take();
                            border_child = BorderChild::None;
                        }
                        b"top" if border_child == BorderChild::Top => {
                            current_border.top = current_border_side.take();
                            border_child = BorderChild::None;
                        }
                        b"bottom" if border_child == BorderChild::Bottom => {
                            current_border.bottom = current_border_side.take();
                            border_child = BorderChild::None;
                        }

                        // ----- border end -----
                        b"border" if in_border => {
                            borders.push(std::mem::take(&mut current_border));
                            in_border = false;
                        }

                        _ => {}
                    }
                }

                Ok(Event::Eof) => break,
                Err(err) => return Err(IronsheetError::XmlParse(err.to_string())),
                _ => {}
            }
            buf.clear();
        }

        Ok(Styles {
            num_fmts,
            fonts,
            fills,
            borders,
            cell_xfs,
        })
    }

    // -----------------------------------------------------------------------
    // Writer
    // -----------------------------------------------------------------------

    /// Serialize this `Styles` to a valid `styles.xml` string.
    pub fn to_xml(&self) -> Result<String> {
        let mut buf: Vec<u8> = Vec::with_capacity(2048);
        let mut writer = Writer::new(&mut buf);

        let map_err = |e: std::io::Error| IronsheetError::XmlWrite(e.to_string());

        // XML declaration.
        writer
            .write_event(Event::Decl(BytesDecl::new("1.0", Some("UTF-8"), Some("yes"))))
            .map_err(&map_err)?;

        // <styleSheet xmlns="...">
        let mut root = BytesStart::new("styleSheet");
        root.push_attribute(("xmlns", SPREADSHEET_NS));
        writer
            .write_event(Event::Start(root))
            .map_err(&map_err)?;

        // --- numFmts ---
        if !self.num_fmts.is_empty() {
            let mut nf_start = BytesStart::new("numFmts");
            nf_start.push_attribute(("count", self.num_fmts.len().to_string().as_str()));
            writer
                .write_event(Event::Start(nf_start))
                .map_err(&map_err)?;

            for nf in &self.num_fmts {
                let mut elem = BytesStart::new("numFmt");
                elem.push_attribute(("numFmtId", nf.id.to_string().as_str()));
                elem.push_attribute(("formatCode", nf.format_code.as_str()));
                writer
                    .write_event(Event::Empty(elem))
                    .map_err(&map_err)?;
            }

            writer
                .write_event(Event::End(BytesEnd::new("numFmts")))
                .map_err(&map_err)?;
        }

        // --- fonts ---
        {
            let mut f_start = BytesStart::new("fonts");
            f_start.push_attribute(("count", self.fonts.len().to_string().as_str()));
            writer
                .write_event(Event::Start(f_start))
                .map_err(&map_err)?;

            for font in &self.fonts {
                writer
                    .write_event(Event::Start(BytesStart::new("font")))
                    .map_err(&map_err)?;

                if font.bold {
                    writer
                        .write_event(Event::Empty(BytesStart::new("b")))
                        .map_err(&map_err)?;
                }
                if font.italic {
                    writer
                        .write_event(Event::Empty(BytesStart::new("i")))
                        .map_err(&map_err)?;
                }
                if font.underline {
                    writer
                        .write_event(Event::Empty(BytesStart::new("u")))
                        .map_err(&map_err)?;
                }
                if font.strike {
                    writer
                        .write_event(Event::Empty(BytesStart::new("strike")))
                        .map_err(&map_err)?;
                }
                if let Some(sz) = font.size {
                    let mut elem = BytesStart::new("sz");
                    // Format without trailing zeros for integers.
                    let val = if sz.fract() == 0.0 {
                        format!("{}", sz as i64)
                    } else {
                        sz.to_string()
                    };
                    elem.push_attribute(("val", val.as_str()));
                    writer
                        .write_event(Event::Empty(elem))
                        .map_err(&map_err)?;
                }
                if let Some(ref name) = font.name {
                    let mut elem = BytesStart::new("name");
                    elem.push_attribute(("val", name.as_str()));
                    writer
                        .write_event(Event::Empty(elem))
                        .map_err(&map_err)?;
                }
                if let Some(ref color) = font.color {
                    let mut elem = BytesStart::new("color");
                    elem.push_attribute(("rgb", color.as_str()));
                    writer
                        .write_event(Event::Empty(elem))
                        .map_err(&map_err)?;
                }

                writer
                    .write_event(Event::End(BytesEnd::new("font")))
                    .map_err(&map_err)?;
            }

            writer
                .write_event(Event::End(BytesEnd::new("fonts")))
                .map_err(&map_err)?;
        }

        // --- fills ---
        {
            let mut f_start = BytesStart::new("fills");
            f_start.push_attribute(("count", self.fills.len().to_string().as_str()));
            writer
                .write_event(Event::Start(f_start))
                .map_err(&map_err)?;

            for fill in &self.fills {
                writer
                    .write_event(Event::Start(BytesStart::new("fill")))
                    .map_err(&map_err)?;

                let has_children = fill.fg_color.is_some() || fill.bg_color.is_some();

                if has_children {
                    let mut pf = BytesStart::new("patternFill");
                    pf.push_attribute(("patternType", fill.pattern_type.as_str()));
                    writer
                        .write_event(Event::Start(pf))
                        .map_err(&map_err)?;

                    if let Some(ref fg) = fill.fg_color {
                        let mut elem = BytesStart::new("fgColor");
                        elem.push_attribute(("rgb", fg.as_str()));
                        writer
                            .write_event(Event::Empty(elem))
                            .map_err(&map_err)?;
                    }
                    if let Some(ref bg) = fill.bg_color {
                        let mut elem = BytesStart::new("bgColor");
                        elem.push_attribute(("rgb", bg.as_str()));
                        writer
                            .write_event(Event::Empty(elem))
                            .map_err(&map_err)?;
                    }

                    writer
                        .write_event(Event::End(BytesEnd::new("patternFill")))
                        .map_err(&map_err)?;
                } else {
                    let mut pf = BytesStart::new("patternFill");
                    pf.push_attribute(("patternType", fill.pattern_type.as_str()));
                    writer
                        .write_event(Event::Empty(pf))
                        .map_err(&map_err)?;
                }

                writer
                    .write_event(Event::End(BytesEnd::new("fill")))
                    .map_err(&map_err)?;
            }

            writer
                .write_event(Event::End(BytesEnd::new("fills")))
                .map_err(&map_err)?;
        }

        // --- borders ---
        {
            let mut b_start = BytesStart::new("borders");
            b_start.push_attribute(("count", self.borders.len().to_string().as_str()));
            writer
                .write_event(Event::Start(b_start))
                .map_err(&map_err)?;

            for border in &self.borders {
                writer
                    .write_event(Event::Start(BytesStart::new("border")))
                    .map_err(&map_err)?;

                write_border_side(&mut writer, "left", &border.left, &map_err)?;
                write_border_side(&mut writer, "right", &border.right, &map_err)?;
                write_border_side(&mut writer, "top", &border.top, &map_err)?;
                write_border_side(&mut writer, "bottom", &border.bottom, &map_err)?;

                // Always write <diagonal/>.
                writer
                    .write_event(Event::Empty(BytesStart::new("diagonal")))
                    .map_err(&map_err)?;

                writer
                    .write_event(Event::End(BytesEnd::new("border")))
                    .map_err(&map_err)?;
            }

            writer
                .write_event(Event::End(BytesEnd::new("borders")))
                .map_err(&map_err)?;
        }

        // --- cellXfs ---
        {
            let mut xf_start = BytesStart::new("cellXfs");
            xf_start.push_attribute(("count", self.cell_xfs.len().to_string().as_str()));
            writer
                .write_event(Event::Start(xf_start))
                .map_err(&map_err)?;

            for xf in &self.cell_xfs {
                let mut elem = BytesStart::new("xf");
                elem.push_attribute(("numFmtId", xf.num_fmt_id.to_string().as_str()));
                elem.push_attribute(("fontId", xf.font_id.to_string().as_str()));
                elem.push_attribute(("fillId", xf.fill_id.to_string().as_str()));
                elem.push_attribute(("borderId", xf.border_id.to_string().as_str()));
                writer
                    .write_event(Event::Empty(elem))
                    .map_err(&map_err)?;
            }

            writer
                .write_event(Event::End(BytesEnd::new("cellXfs")))
                .map_err(&map_err)?;
        }

        // </styleSheet>
        writer
            .write_event(Event::End(BytesEnd::new("styleSheet")))
            .map_err(&map_err)?;

        String::from_utf8(buf)
            .map_err(|e| IronsheetError::XmlWrite(format!("invalid UTF-8 in output: {e}")))
    }
}

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

/// Extract a `BorderSide` from attributes on a border child element (e.g. `<left style="thin">`).
/// Returns `None` if there is no `style` attribute.
fn parse_border_side_attrs(e: &BytesStart<'_>) -> Option<BorderSide> {
    let mut style: Option<String> = None;
    for attr in e.attributes().flatten() {
        if attr.key.local_name().as_ref() == b"style" {
            style = Some(std::str::from_utf8(&attr.value).unwrap_or_default().to_owned());
        }
    }
    style.map(|s| BorderSide {
        style: s,
        color: None,
    })
}

/// Parse attributes of an `<xf>` element into a `CellXf`.
fn parse_xf_attrs(e: &BytesStart<'_>) -> CellXf {
    let mut xf = CellXf::default();
    for attr in e.attributes().flatten() {
        match attr.key.local_name().as_ref() {
            b"numFmtId" => {
                xf.num_fmt_id = std::str::from_utf8(&attr.value).unwrap_or_default()
                    .parse()
                    .unwrap_or(0);
            }
            b"fontId" => {
                xf.font_id = std::str::from_utf8(&attr.value).unwrap_or_default()
                    .parse()
                    .unwrap_or(0);
            }
            b"fillId" => {
                xf.fill_id = std::str::from_utf8(&attr.value).unwrap_or_default()
                    .parse()
                    .unwrap_or(0);
            }
            b"borderId" => {
                xf.border_id = std::str::from_utf8(&attr.value).unwrap_or_default()
                    .parse()
                    .unwrap_or(0);
            }
            _ => {}
        }
    }
    xf
}

/// Write a single border side element (`<left/>`, `<left style="thin"><color rgb="..."/></left>`, etc.).
fn write_border_side<W: std::io::Write>(
    writer: &mut Writer<W>,
    tag: &str,
    side: &Option<BorderSide>,
    map_err: &dyn Fn(std::io::Error) -> IronsheetError,
) -> Result<()> {
    match side {
        None => {
            writer
                .write_event(Event::Empty(BytesStart::new(tag)))
                .map_err(map_err)?;
        }
        Some(bs) => {
            let mut elem = BytesStart::new(tag);
            elem.push_attribute(("style", bs.style.as_str()));

            if let Some(ref color) = bs.color {
                writer
                    .write_event(Event::Start(elem))
                    .map_err(map_err)?;

                let mut c = BytesStart::new("color");
                c.push_attribute(("rgb", color.as_str()));
                writer
                    .write_event(Event::Empty(c))
                    .map_err(map_err)?;

                writer
                    .write_event(Event::End(BytesEnd::new(tag)))
                    .map_err(map_err)?;
            } else {
                writer
                    .write_event(Event::Empty(elem))
                    .map_err(map_err)?;
            }
        }
    }
    Ok(())
}

// ---------------------------------------------------------------------------
// Tests
// ---------------------------------------------------------------------------

#[cfg(test)]
mod tests {
    use super::*;

    const MINIMAL_STYLES: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <numFmts count="1">
    <numFmt numFmtId="164" formatCode="yyyy-mm-dd"/>
  </numFmts>
  <fonts count="2">
    <font><sz val="11"/><name val="Aptos"/></font>
    <font><b/><sz val="11"/><name val="Aptos"/></font>
  </fonts>
  <fills count="2">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
  </fills>
  <borders count="1">
    <border><left/><right/><top/><bottom/><diagonal/></border>
  </borders>
  <cellXfs count="3">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
    <xf numFmtId="0" fontId="1" fillId="0" borderId="0" applyFont="1"/>
    <xf numFmtId="164" fontId="0" fillId="0" borderId="0" applyNumberFormat="1"/>
  </cellXfs>
</styleSheet>"#;

    #[test]
    fn test_parse_styles() {
        let styles = Styles::parse(MINIMAL_STYLES.as_bytes()).unwrap();

        // numFmts
        assert_eq!(styles.num_fmts.len(), 1);
        assert_eq!(styles.num_fmts[0].id, 164);
        assert_eq!(styles.num_fmts[0].format_code, "yyyy-mm-dd");

        // fonts
        assert_eq!(styles.fonts.len(), 2);
        assert!(!styles.fonts[0].bold);
        assert_eq!(styles.fonts[0].name.as_deref(), Some("Aptos"));
        assert_eq!(styles.fonts[0].size, Some(11.0));
        assert!(styles.fonts[1].bold);
        assert_eq!(styles.fonts[1].name.as_deref(), Some("Aptos"));
        assert_eq!(styles.fonts[1].size, Some(11.0));

        // fills
        assert_eq!(styles.fills.len(), 2);
        assert_eq!(styles.fills[0].pattern_type, "none");
        assert_eq!(styles.fills[1].pattern_type, "gray125");

        // borders
        assert_eq!(styles.borders.len(), 1);
        assert!(styles.borders[0].left.is_none());
        assert!(styles.borders[0].right.is_none());
        assert!(styles.borders[0].top.is_none());
        assert!(styles.borders[0].bottom.is_none());

        // cellXfs
        assert_eq!(styles.cell_xfs.len(), 3);
    }

    #[test]
    fn test_style_get_num_fmt_id() {
        let styles = Styles::parse(MINIMAL_STYLES.as_bytes()).unwrap();
        assert_eq!(styles.cell_xfs[0].num_fmt_id, 0);
        assert_eq!(styles.cell_xfs[2].num_fmt_id, 164);
    }

    #[test]
    fn test_styles_write_roundtrip() {
        let styles1 = Styles::parse(MINIMAL_STYLES.as_bytes()).unwrap();
        let xml = styles1.to_xml().unwrap();
        let styles2 = Styles::parse(xml.as_bytes()).unwrap();

        assert_eq!(styles2.num_fmts.len(), styles1.num_fmts.len());
        assert_eq!(styles2.fonts.len(), styles1.fonts.len());
        assert_eq!(styles2.fills.len(), styles1.fills.len());
        assert_eq!(styles2.borders.len(), styles1.borders.len());
        assert_eq!(styles2.cell_xfs.len(), styles1.cell_xfs.len());

        // Spot-check values survived the roundtrip.
        assert_eq!(styles2.num_fmts[0].id, 164);
        assert_eq!(styles2.num_fmts[0].format_code, "yyyy-mm-dd");
        assert!(styles2.fonts[1].bold);
        assert_eq!(styles2.fills[1].pattern_type, "gray125");
        assert_eq!(styles2.cell_xfs[2].num_fmt_id, 164);
    }

    #[test]
    fn test_default_styles() {
        let styles = Styles::default_styles();

        // Verify shape.
        assert_eq!(styles.num_fmts.len(), 0);
        assert_eq!(styles.fonts.len(), 1);
        assert_eq!(styles.fills.len(), 2);
        assert_eq!(styles.borders.len(), 1);
        assert_eq!(styles.cell_xfs.len(), 1);

        // Verify default font.
        assert_eq!(styles.fonts[0].name.as_deref(), Some("Aptos"));
        assert_eq!(styles.fonts[0].size, Some(11.0));

        // Verify fills.
        assert_eq!(styles.fills[0].pattern_type, "none");
        assert_eq!(styles.fills[1].pattern_type, "gray125");

        // Roundtrip through XML.
        let xml = styles.to_xml().unwrap();
        let styles2 = Styles::parse(xml.as_bytes()).unwrap();

        assert_eq!(styles2.num_fmts.len(), 0);
        assert_eq!(styles2.fonts.len(), 1);
        assert_eq!(styles2.fills.len(), 2);
        assert_eq!(styles2.borders.len(), 1);
        assert_eq!(styles2.cell_xfs.len(), 1);
        assert_eq!(styles2.fonts[0].name.as_deref(), Some("Aptos"));
        assert_eq!(styles2.fonts[0].size, Some(11.0));
    }
}
