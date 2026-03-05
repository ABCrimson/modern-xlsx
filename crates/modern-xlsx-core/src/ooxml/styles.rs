use quick_xml::events::{BytesDecl, BytesEnd, BytesStart, Event};
use quick_xml::{Reader, Writer};
use serde::{Deserialize, Serialize};

use crate::{ModernXlsxError, Result};

use super::SPREADSHEET_NS;

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
pub struct Alignment {
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub horizontal: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub vertical: Option<String>,
    #[serde(default, skip_serializing_if = "crate::ooxml::styles::is_false")]
    pub wrap_text: bool,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub text_rotation: Option<u16>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub indent: Option<u32>,
    #[serde(default, skip_serializing_if = "crate::ooxml::styles::is_false")]
    pub shrink_to_fit: bool,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct Protection {
    #[serde(default = "default_true")]
    pub locked: bool,
    #[serde(default)]
    pub hidden: bool,
}

impl Default for Protection {
    fn default() -> Self {
        Protection {
            locked: true,
            hidden: false,
        }
    }
}

fn default_true() -> bool {
    true
}

fn is_false(v: &bool) -> bool {
    !v
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
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub vert_align: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub family: Option<u32>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub charset: Option<u32>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub scheme: Option<String>,
    #[serde(default, skip_serializing_if = "is_false")]
    pub condense: bool,
    #[serde(default, skip_serializing_if = "is_false")]
    pub extend: bool,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct GradientStop {
    pub position: f64,
    pub color: String,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct GradientFill {
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub gradient_type: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub degree: Option<f64>,
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub stops: Vec<GradientStop>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct Fill {
    pub pattern_type: String,
    pub fg_color: Option<String>,
    pub bg_color: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub gradient_fill: Option<GradientFill>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct Border {
    pub left: Option<BorderSide>,
    pub right: Option<BorderSide>,
    pub top: Option<BorderSide>,
    pub bottom: Option<BorderSide>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub diagonal: Option<BorderSide>,
    #[serde(default, skip_serializing_if = "is_false")]
    pub diagonal_up: bool,
    #[serde(default, skip_serializing_if = "is_false")]
    pub diagonal_down: bool,
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
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub alignment: Option<Alignment>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub protection: Option<Protection>,
    #[serde(default, skip_serializing_if = "is_false")]
    pub apply_font: bool,
    #[serde(default, skip_serializing_if = "is_false")]
    pub apply_fill: bool,
    #[serde(default, skip_serializing_if = "is_false")]
    pub apply_border: bool,
    #[serde(default, skip_serializing_if = "is_false")]
    pub apply_number_format: bool,
    #[serde(default, skip_serializing_if = "is_false")]
    pub apply_alignment: bool,
    #[serde(default, skip_serializing_if = "is_false")]
    pub apply_protection: bool,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct DxfStyle {
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub font: Option<Font>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub fill: Option<Fill>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub border: Option<Border>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub num_fmt: Option<NumFmt>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct CellStyle {
    pub name: String,
    pub xf_id: u32,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub builtin_id: Option<u32>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct Styles {
    pub num_fmts: Vec<NumFmt>,
    pub fonts: Vec<Font>,
    pub fills: Vec<Fill>,
    pub borders: Vec<Border>,
    pub cell_xfs: Vec<CellXf>,
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub dxfs: Vec<DxfStyle>,
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub cell_styles: Vec<CellStyle>,
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
    Dxfs,
    CellStyles,
}

/// Tracks which border child element we are inside.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum BorderChild {
    None,
    Left,
    Right,
    Top,
    Bottom,
    Diagonal,
}

/// Tracks where we are inside a DXF element.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum DxfChild {
    None,
    Font,
    Fill,
    Border,
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
            dxfs: Vec::new(),
            cell_styles: Vec::new(),
        }
    }

    // -----------------------------------------------------------------------
    // Parser
    // -----------------------------------------------------------------------

    /// Parse a `styles.xml` document from raw XML bytes.
    pub fn parse(data: &[u8]) -> Result<Self> {
        let mut reader = Reader::from_reader(data);
        reader.config_mut().trim_text(true);

        let mut buf = Vec::with_capacity(1024);

        let mut num_fmts: Vec<NumFmt> = Vec::new();
        let mut fonts: Vec<Font> = Vec::new();
        let mut fills: Vec<Fill> = Vec::new();
        let mut borders: Vec<Border> = Vec::new();
        let mut cell_xfs: Vec<CellXf> = Vec::new();
        let mut dxfs: Vec<DxfStyle> = Vec::new();
        let mut cell_styles: Vec<CellStyle> = Vec::new();

        let mut section = Section::None;

        // Font parsing state.
        let mut current_font = Font::default();
        let mut in_font = false;

        // Fill parsing state.
        let mut current_fill = Fill::default();
        let mut in_fill = false;
        let mut in_gradient_fill = false;
        let mut current_gradient = GradientFill::default();
        let mut in_gradient_stop = false;
        let mut current_stop_position: f64 = 0.0;

        // Border parsing state.
        let mut current_border = Border::default();
        let mut in_border = false;
        let mut border_child = BorderChild::None;
        let mut current_border_side: Option<BorderSide> = None;

        // CellXf parsing state.
        let mut current_xf: Option<CellXf> = None;
        let mut in_xf = false;

        // DXF parsing state.
        let mut current_dxf = DxfStyle::default();
        let mut in_dxf = false;
        let mut dxf_child = DxfChild::None;
        let mut dxf_font = Font::default();
        let mut dxf_fill = Fill::default();
        let mut dxf_border = Border::default();
        let mut dxf_in_gradient_fill = false;
        let mut dxf_current_gradient = GradientFill::default();
        let mut dxf_in_gradient_stop = false;
        let mut dxf_current_stop_position: f64 = 0.0;

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
                        b"dxfs" => section = Section::Dxfs,
                        b"cellStyles" => section = Section::CellStyles,

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
                        b"patternFill" if in_fill && !in_dxf => {
                            for attr in e.attributes().flatten() {
                                if attr.key.local_name().as_ref() == b"patternType" {
                                    current_fill.pattern_type =
                                        std::str::from_utf8(&attr.value).unwrap_or_default().to_owned();
                                }
                            }
                        }
                        b"gradientFill" if in_fill && !in_dxf => {
                            in_gradient_fill = true;
                            current_gradient = GradientFill::default();
                            for attr in e.attributes().flatten() {
                                match attr.key.local_name().as_ref() {
                                    b"type" => {
                                        current_gradient.gradient_type = Some(
                                            std::str::from_utf8(&attr.value).unwrap_or_default().to_owned(),
                                        );
                                    }
                                    b"degree" => {
                                        current_gradient.degree = std::str::from_utf8(&attr.value)
                                            .unwrap_or_default()
                                            .parse()
                                            .ok();
                                    }
                                    _ => {}
                                }
                            }
                        }
                        b"stop" if in_gradient_fill && !in_dxf => {
                            in_gradient_stop = true;
                            current_stop_position = 0.0;
                            for attr in e.attributes().flatten() {
                                if attr.key.local_name().as_ref() == b"position" {
                                    current_stop_position = std::str::from_utf8(&attr.value)
                                        .unwrap_or_default()
                                        .parse()
                                        .unwrap_or(0.0);
                                }
                            }
                        }

                        // ----- border children -----
                        b"border" if section == Section::Borders => {
                            current_border = Border::default();
                            in_border = true;
                            for attr in e.attributes().flatten() {
                                match attr.key.local_name().as_ref() {
                                    b"diagonalUp" => {
                                        let v = std::str::from_utf8(&attr.value).unwrap_or_default();
                                        current_border.diagonal_up = v == "1" || v == "true";
                                    }
                                    b"diagonalDown" => {
                                        let v = std::str::from_utf8(&attr.value).unwrap_or_default();
                                        current_border.diagonal_down = v == "1" || v == "true";
                                    }
                                    _ => {}
                                }
                            }
                        }
                        b"left" if in_border && !in_dxf => {
                            border_child = BorderChild::Left;
                            current_border_side = parse_border_side_attrs(e);
                        }
                        b"right" if in_border && !in_dxf => {
                            border_child = BorderChild::Right;
                            current_border_side = parse_border_side_attrs(e);
                        }
                        b"top" if in_border && !in_dxf => {
                            border_child = BorderChild::Top;
                            current_border_side = parse_border_side_attrs(e);
                        }
                        b"bottom" if in_border && !in_dxf => {
                            border_child = BorderChild::Bottom;
                            current_border_side = parse_border_side_attrs(e);
                        }
                        b"diagonal" if in_border && !in_dxf => {
                            border_child = BorderChild::Diagonal;
                            current_border_side = parse_border_side_attrs(e);
                        }

                        // colour inside a border side
                        b"color" if border_child != BorderChild::None && !in_dxf => {
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

                        // ----- xf (start variant) -----
                        b"xf" if section == Section::CellXfs => {
                            let xf = parse_xf_attrs(e);
                            current_xf = Some(xf);
                            in_xf = true;
                        }

                        // ----- alignment inside xf -----
                        b"alignment" if in_xf => {
                            if let Some(ref mut xf) = current_xf {
                                xf.alignment = Some(parse_alignment_attrs(e));
                            }
                        }

                        // ----- protection inside xf -----
                        b"protection" if in_xf => {
                            if let Some(ref mut xf) = current_xf {
                                xf.protection = Some(parse_protection_attrs(e));
                            }
                        }

                        // ----- DXF section -----
                        b"dxf" if section == Section::Dxfs => {
                            current_dxf = DxfStyle::default();
                            in_dxf = true;
                        }

                        // DXF children
                        b"font" if in_dxf && dxf_child == DxfChild::None => {
                            dxf_font = Font::default();
                            dxf_child = DxfChild::Font;
                        }
                        b"fill" if in_dxf && dxf_child == DxfChild::None => {
                            dxf_fill = Fill::default();
                            dxf_child = DxfChild::Fill;
                        }
                        b"patternFill" if in_dxf && dxf_child == DxfChild::Fill => {
                            for attr in e.attributes().flatten() {
                                if attr.key.local_name().as_ref() == b"patternType" {
                                    dxf_fill.pattern_type =
                                        std::str::from_utf8(&attr.value).unwrap_or_default().to_owned();
                                }
                            }
                        }
                        b"gradientFill" if in_dxf && dxf_child == DxfChild::Fill => {
                            dxf_in_gradient_fill = true;
                            dxf_current_gradient = GradientFill::default();
                            for attr in e.attributes().flatten() {
                                match attr.key.local_name().as_ref() {
                                    b"type" => {
                                        dxf_current_gradient.gradient_type = Some(
                                            std::str::from_utf8(&attr.value).unwrap_or_default().to_owned(),
                                        );
                                    }
                                    b"degree" => {
                                        dxf_current_gradient.degree = std::str::from_utf8(&attr.value)
                                            .unwrap_or_default()
                                            .parse()
                                            .ok();
                                    }
                                    _ => {}
                                }
                            }
                        }
                        b"stop" if in_dxf && dxf_in_gradient_fill => {
                            dxf_in_gradient_stop = true;
                            dxf_current_stop_position = 0.0;
                            for attr in e.attributes().flatten() {
                                if attr.key.local_name().as_ref() == b"position" {
                                    dxf_current_stop_position = std::str::from_utf8(&attr.value)
                                        .unwrap_or_default()
                                        .parse()
                                        .unwrap_or(0.0);
                                }
                            }
                        }
                        b"border" if in_dxf && dxf_child == DxfChild::None => {
                            dxf_border = Border::default();
                            dxf_child = DxfChild::Border;
                            for attr in e.attributes().flatten() {
                                match attr.key.local_name().as_ref() {
                                    b"diagonalUp" => {
                                        let v = std::str::from_utf8(&attr.value).unwrap_or_default();
                                        dxf_border.diagonal_up = v == "1" || v == "true";
                                    }
                                    b"diagonalDown" => {
                                        let v = std::str::from_utf8(&attr.value).unwrap_or_default();
                                        dxf_border.diagonal_down = v == "1" || v == "true";
                                    }
                                    _ => {}
                                }
                            }
                        }
                        b"left" if in_dxf && dxf_child == DxfChild::Border => {
                            border_child = BorderChild::Left;
                            current_border_side = parse_border_side_attrs(e);
                        }
                        b"right" if in_dxf && dxf_child == DxfChild::Border => {
                            border_child = BorderChild::Right;
                            current_border_side = parse_border_side_attrs(e);
                        }
                        b"top" if in_dxf && dxf_child == DxfChild::Border => {
                            border_child = BorderChild::Top;
                            current_border_side = parse_border_side_attrs(e);
                        }
                        b"bottom" if in_dxf && dxf_child == DxfChild::Border => {
                            border_child = BorderChild::Bottom;
                            current_border_side = parse_border_side_attrs(e);
                        }
                        b"diagonal" if in_dxf && dxf_child == DxfChild::Border => {
                            border_child = BorderChild::Diagonal;
                            current_border_side = parse_border_side_attrs(e);
                        }
                        b"color" if in_dxf && border_child != BorderChild::None => {
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

                        _ => {}
                    }
                }

                Ok(Event::Empty(ref e)) => {
                    let local = e.local_name();
                    match local.as_ref() {
                        // ----- numFmt is always self-closing -----
                        b"numFmt" if section == Section::NumFmts || (in_dxf && dxf_child == DxfChild::None) => {
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
                            if section == Section::NumFmts {
                                num_fmts.push(NumFmt {
                                    id,
                                    format_code: code,
                                });
                            } else if in_dxf {
                                current_dxf.num_fmt = Some(NumFmt {
                                    id,
                                    format_code: code,
                                });
                            }
                        }

                        // ----- font child elements (self-closing) -----
                        b"b" if in_font || (in_dxf && dxf_child == DxfChild::Font) => {
                            if in_dxf && dxf_child == DxfChild::Font {
                                dxf_font.bold = true;
                            } else {
                                current_font.bold = true;
                            }
                        }
                        b"i" if in_font || (in_dxf && dxf_child == DxfChild::Font) => {
                            if in_dxf && dxf_child == DxfChild::Font {
                                dxf_font.italic = true;
                            } else {
                                current_font.italic = true;
                            }
                        }
                        b"u" if in_font || (in_dxf && dxf_child == DxfChild::Font) => {
                            if in_dxf && dxf_child == DxfChild::Font {
                                dxf_font.underline = true;
                            } else {
                                current_font.underline = true;
                            }
                        }
                        b"strike" if in_font || (in_dxf && dxf_child == DxfChild::Font) => {
                            if in_dxf && dxf_child == DxfChild::Font {
                                dxf_font.strike = true;
                            } else {
                                current_font.strike = true;
                            }
                        }
                        b"condense" if in_font || (in_dxf && dxf_child == DxfChild::Font) => {
                            if in_dxf && dxf_child == DxfChild::Font {
                                dxf_font.condense = true;
                            } else {
                                current_font.condense = true;
                            }
                        }
                        b"extend" if in_font || (in_dxf && dxf_child == DxfChild::Font) => {
                            if in_dxf && dxf_child == DxfChild::Font {
                                dxf_font.extend = true;
                            } else {
                                current_font.extend = true;
                            }
                        }
                        b"sz" if in_font || (in_dxf && dxf_child == DxfChild::Font) => {
                            for attr in e.attributes().flatten() {
                                if attr.key.local_name().as_ref() == b"val" {
                                    let val = std::str::from_utf8(&attr.value).unwrap_or_default()
                                        .parse()
                                        .ok();
                                    if in_dxf && dxf_child == DxfChild::Font {
                                        dxf_font.size = val;
                                    } else {
                                        current_font.size = val;
                                    }
                                }
                            }
                        }
                        b"name" if in_font || (in_dxf && dxf_child == DxfChild::Font) => {
                            for attr in e.attributes().flatten() {
                                if attr.key.local_name().as_ref() == b"val" {
                                    let val = Some(
                                        std::str::from_utf8(&attr.value).unwrap_or_default().to_owned(),
                                    );
                                    if in_dxf && dxf_child == DxfChild::Font {
                                        dxf_font.name = val;
                                    } else {
                                        current_font.name = val;
                                    }
                                }
                            }
                        }
                        b"color" if (in_font || (in_dxf && dxf_child == DxfChild::Font)) && border_child == BorderChild::None => {
                            for attr in e.attributes().flatten() {
                                if attr.key.local_name().as_ref() == b"rgb" {
                                    let val = Some(
                                        std::str::from_utf8(&attr.value).unwrap_or_default().to_owned(),
                                    );
                                    if in_dxf && dxf_child == DxfChild::Font {
                                        dxf_font.color = val;
                                    } else {
                                        current_font.color = val;
                                    }
                                }
                            }
                        }
                        b"vertAlign" if in_font || (in_dxf && dxf_child == DxfChild::Font) => {
                            for attr in e.attributes().flatten() {
                                if attr.key.local_name().as_ref() == b"val" {
                                    let val = Some(
                                        std::str::from_utf8(&attr.value).unwrap_or_default().to_owned(),
                                    );
                                    if in_dxf && dxf_child == DxfChild::Font {
                                        dxf_font.vert_align = val;
                                    } else {
                                        current_font.vert_align = val;
                                    }
                                }
                            }
                        }
                        b"family" if in_font || (in_dxf && dxf_child == DxfChild::Font) => {
                            for attr in e.attributes().flatten() {
                                if attr.key.local_name().as_ref() == b"val" {
                                    let val = std::str::from_utf8(&attr.value).unwrap_or_default()
                                        .parse()
                                        .ok();
                                    if in_dxf && dxf_child == DxfChild::Font {
                                        dxf_font.family = val;
                                    } else {
                                        current_font.family = val;
                                    }
                                }
                            }
                        }
                        b"charset" if in_font || (in_dxf && dxf_child == DxfChild::Font) => {
                            for attr in e.attributes().flatten() {
                                if attr.key.local_name().as_ref() == b"val" {
                                    let val = std::str::from_utf8(&attr.value).unwrap_or_default()
                                        .parse()
                                        .ok();
                                    if in_dxf && dxf_child == DxfChild::Font {
                                        dxf_font.charset = val;
                                    } else {
                                        current_font.charset = val;
                                    }
                                }
                            }
                        }
                        b"scheme" if in_font || (in_dxf && dxf_child == DxfChild::Font) => {
                            for attr in e.attributes().flatten() {
                                if attr.key.local_name().as_ref() == b"val" {
                                    let val = Some(
                                        std::str::from_utf8(&attr.value).unwrap_or_default().to_owned(),
                                    );
                                    if in_dxf && dxf_child == DxfChild::Font {
                                        dxf_font.scheme = val;
                                    } else {
                                        current_font.scheme = val;
                                    }
                                }
                            }
                        }

                        // ----- patternFill as self-closing -----
                        b"patternFill" if in_fill && !in_dxf => {
                            for attr in e.attributes().flatten() {
                                if attr.key.local_name().as_ref() == b"patternType" {
                                    current_fill.pattern_type =
                                        std::str::from_utf8(&attr.value).unwrap_or_default().to_owned();
                                }
                            }
                        }
                        b"patternFill" if in_dxf && dxf_child == DxfChild::Fill => {
                            for attr in e.attributes().flatten() {
                                if attr.key.local_name().as_ref() == b"patternType" {
                                    dxf_fill.pattern_type =
                                        std::str::from_utf8(&attr.value).unwrap_or_default().to_owned();
                                }
                            }
                        }

                        // ----- fgColor / bgColor inside fill -----
                        b"fgColor" if (in_fill && !in_dxf) || (in_dxf && dxf_child == DxfChild::Fill) => {
                            for attr in e.attributes().flatten() {
                                if attr.key.local_name().as_ref() == b"rgb" {
                                    let val = Some(
                                        std::str::from_utf8(&attr.value).unwrap_or_default().to_owned(),
                                    );
                                    if in_dxf && dxf_child == DxfChild::Fill {
                                        dxf_fill.fg_color = val;
                                    } else {
                                        current_fill.fg_color = val;
                                    }
                                }
                            }
                        }
                        b"bgColor" if (in_fill && !in_dxf) || (in_dxf && dxf_child == DxfChild::Fill) => {
                            for attr in e.attributes().flatten() {
                                if attr.key.local_name().as_ref() == b"rgb" {
                                    let val = Some(
                                        std::str::from_utf8(&attr.value).unwrap_or_default().to_owned(),
                                    );
                                    if in_dxf && dxf_child == DxfChild::Fill {
                                        dxf_fill.bg_color = val;
                                    } else {
                                        current_fill.bg_color = val;
                                    }
                                }
                            }
                        }

                        // ----- color inside gradient stop -----
                        b"color" if in_gradient_stop && !in_dxf => {
                            for attr in e.attributes().flatten() {
                                if attr.key.local_name().as_ref() == b"rgb" {
                                    let color = std::str::from_utf8(&attr.value).unwrap_or_default().to_owned();
                                    current_gradient.stops.push(GradientStop {
                                        position: current_stop_position,
                                        color,
                                    });
                                }
                            }
                        }
                        b"color" if dxf_in_gradient_stop && in_dxf => {
                            for attr in e.attributes().flatten() {
                                if attr.key.local_name().as_ref() == b"rgb" {
                                    let color = std::str::from_utf8(&attr.value).unwrap_or_default().to_owned();
                                    dxf_current_gradient.stops.push(GradientStop {
                                        position: dxf_current_stop_position,
                                        color,
                                    });
                                }
                            }
                        }

                        // ----- empty border sides (e.g. <left/>) -----
                        b"left" if in_border && !in_dxf => {
                            current_border.left = parse_border_side_attrs(e);
                        }
                        b"right" if in_border && !in_dxf => {
                            current_border.right = parse_border_side_attrs(e);
                        }
                        b"top" if in_border && !in_dxf => {
                            current_border.top = parse_border_side_attrs(e);
                        }
                        b"bottom" if in_border && !in_dxf => {
                            current_border.bottom = parse_border_side_attrs(e);
                        }
                        b"diagonal" if in_border && !in_dxf => {
                            current_border.diagonal = parse_border_side_attrs(e);
                        }

                        // DXF border sides (self-closing)
                        b"left" if in_dxf && dxf_child == DxfChild::Border => {
                            dxf_border.left = parse_border_side_attrs(e);
                        }
                        b"right" if in_dxf && dxf_child == DxfChild::Border => {
                            dxf_border.right = parse_border_side_attrs(e);
                        }
                        b"top" if in_dxf && dxf_child == DxfChild::Border => {
                            dxf_border.top = parse_border_side_attrs(e);
                        }
                        b"bottom" if in_dxf && dxf_child == DxfChild::Border => {
                            dxf_border.bottom = parse_border_side_attrs(e);
                        }
                        b"diagonal" if in_dxf && dxf_child == DxfChild::Border => {
                            dxf_border.diagonal = parse_border_side_attrs(e);
                        }

                        // colour inside a DXF border side (self-closing)
                        b"color" if in_dxf && border_child != BorderChild::None => {
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

                        // colour inside a border side (self-closing)
                        b"color" if border_child != BorderChild::None && !in_dxf => {
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

                        // ----- alignment (self-closing inside xf) -----
                        b"alignment" if in_xf => {
                            if let Some(ref mut xf) = current_xf {
                                xf.alignment = Some(parse_alignment_attrs(e));
                            }
                        }

                        // ----- protection (self-closing inside xf) -----
                        b"protection" if in_xf => {
                            if let Some(ref mut xf) = current_xf {
                                xf.protection = Some(parse_protection_attrs(e));
                            }
                        }

                        // ----- cellStyle (self-closing) -----
                        b"cellStyle" if section == Section::CellStyles => {
                            cell_styles.push(parse_cell_style_attrs(e));
                        }

                        // ----- numFmt inside DXF -----
                        b"numFmt" if in_dxf && dxf_child == DxfChild::None => {
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
                            current_dxf.num_fmt = Some(NumFmt {
                                id,
                                format_code: code,
                            });
                        }

                        _ => {}
                    }
                }

                Ok(Event::End(ref e)) => {
                    let local = e.local_name();
                    match local.as_ref() {
                        // ----- section closers -----
                        b"numFmts" => section = Section::None,
                        b"fonts" if section == Section::Fonts => section = Section::None,
                        b"fills" if section == Section::Fills => section = Section::None,
                        b"borders" if section == Section::Borders => section = Section::None,
                        b"cellXfs" => section = Section::None,
                        b"dxfs" => section = Section::None,
                        b"cellStyles" => section = Section::None,

                        // ----- font end -----
                        b"font" if in_font && !in_dxf => {
                            fonts.push(std::mem::take(&mut current_font));
                            in_font = false;
                        }
                        b"font" if in_dxf && dxf_child == DxfChild::Font => {
                            current_dxf.font = Some(std::mem::take(&mut dxf_font));
                            dxf_child = DxfChild::None;
                        }

                        // ----- fill end -----
                        b"fill" if in_fill && !in_dxf => {
                            fills.push(std::mem::take(&mut current_fill));
                            in_fill = false;
                        }
                        b"fill" if in_dxf && dxf_child == DxfChild::Fill => {
                            current_dxf.fill = Some(std::mem::take(&mut dxf_fill));
                            dxf_child = DxfChild::None;
                        }

                        // ----- gradientFill end -----
                        b"gradientFill" if in_gradient_fill && !in_dxf => {
                            current_fill.gradient_fill = Some(std::mem::take(&mut current_gradient));
                            in_gradient_fill = false;
                        }
                        b"gradientFill" if dxf_in_gradient_fill && in_dxf => {
                            dxf_fill.gradient_fill = Some(std::mem::take(&mut dxf_current_gradient));
                            dxf_in_gradient_fill = false;
                        }

                        // ----- stop end -----
                        b"stop" if in_gradient_stop && !in_dxf => {
                            in_gradient_stop = false;
                        }
                        b"stop" if dxf_in_gradient_stop && in_dxf => {
                            dxf_in_gradient_stop = false;
                        }

                        // ----- border side ends -----
                        b"left" if border_child == BorderChild::Left && !in_dxf => {
                            current_border.left = current_border_side.take();
                            border_child = BorderChild::None;
                        }
                        b"right" if border_child == BorderChild::Right && !in_dxf => {
                            current_border.right = current_border_side.take();
                            border_child = BorderChild::None;
                        }
                        b"top" if border_child == BorderChild::Top && !in_dxf => {
                            current_border.top = current_border_side.take();
                            border_child = BorderChild::None;
                        }
                        b"bottom" if border_child == BorderChild::Bottom && !in_dxf => {
                            current_border.bottom = current_border_side.take();
                            border_child = BorderChild::None;
                        }
                        b"diagonal" if border_child == BorderChild::Diagonal && !in_dxf => {
                            current_border.diagonal = current_border_side.take();
                            border_child = BorderChild::None;
                        }

                        // DXF border side ends
                        b"left" if border_child == BorderChild::Left && in_dxf => {
                            dxf_border.left = current_border_side.take();
                            border_child = BorderChild::None;
                        }
                        b"right" if border_child == BorderChild::Right && in_dxf => {
                            dxf_border.right = current_border_side.take();
                            border_child = BorderChild::None;
                        }
                        b"top" if border_child == BorderChild::Top && in_dxf => {
                            dxf_border.top = current_border_side.take();
                            border_child = BorderChild::None;
                        }
                        b"bottom" if border_child == BorderChild::Bottom && in_dxf => {
                            dxf_border.bottom = current_border_side.take();
                            border_child = BorderChild::None;
                        }
                        b"diagonal" if border_child == BorderChild::Diagonal && in_dxf => {
                            dxf_border.diagonal = current_border_side.take();
                            border_child = BorderChild::None;
                        }

                        // ----- border end -----
                        b"border" if in_border && !in_dxf => {
                            borders.push(std::mem::take(&mut current_border));
                            in_border = false;
                        }
                        b"border" if in_dxf && dxf_child == DxfChild::Border => {
                            current_dxf.border = Some(std::mem::take(&mut dxf_border));
                            dxf_child = DxfChild::None;
                        }

                        // ----- xf end -----
                        b"xf" if in_xf => {
                            if let Some(xf) = current_xf.take() {
                                cell_xfs.push(xf);
                            }
                            in_xf = false;
                        }

                        // ----- dxf end -----
                        b"dxf" if in_dxf => {
                            dxfs.push(std::mem::take(&mut current_dxf));
                            in_dxf = false;
                            dxf_child = DxfChild::None;
                        }

                        _ => {}
                    }
                }

                Ok(Event::Eof) => break,
                Err(err) => return Err(ModernXlsxError::XmlParse(err.to_string())),
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
            dxfs,
            cell_styles,
        })
    }

    // -----------------------------------------------------------------------
    // Writer
    // -----------------------------------------------------------------------

    /// Serialize this `Styles` to valid `styles.xml` bytes.
    pub fn to_xml(&self) -> Result<Vec<u8>> {
        let mut buf: Vec<u8> = Vec::with_capacity(4096);
        let mut writer = Writer::new(&mut buf);
        let mut ibuf = itoa::Buffer::new();

        let map_err = |e: std::io::Error| ModernXlsxError::XmlWrite(e.to_string());

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
            nf_start.push_attribute(("count", ibuf.format(self.num_fmts.len())));
            writer
                .write_event(Event::Start(nf_start))
                .map_err(&map_err)?;

            for nf in &self.num_fmts {
                let mut elem = BytesStart::new("numFmt");
                elem.push_attribute(("numFmtId", ibuf.format(nf.id)));
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
            f_start.push_attribute(("count", ibuf.format(self.fonts.len())));
            writer
                .write_event(Event::Start(f_start))
                .map_err(&map_err)?;

            for font in &self.fonts {
                write_font(&mut writer, font, &mut ibuf, &map_err)?;
            }

            writer
                .write_event(Event::End(BytesEnd::new("fonts")))
                .map_err(&map_err)?;
        }

        // --- fills ---
        {
            let mut f_start = BytesStart::new("fills");
            f_start.push_attribute(("count", ibuf.format(self.fills.len())));
            writer
                .write_event(Event::Start(f_start))
                .map_err(&map_err)?;

            for fill in &self.fills {
                write_fill(&mut writer, fill, &mut ibuf, &map_err)?;
            }

            writer
                .write_event(Event::End(BytesEnd::new("fills")))
                .map_err(&map_err)?;
        }

        // --- borders ---
        {
            let mut b_start = BytesStart::new("borders");
            b_start.push_attribute(("count", ibuf.format(self.borders.len())));
            writer
                .write_event(Event::Start(b_start))
                .map_err(&map_err)?;

            for border in &self.borders {
                write_border(&mut writer, border, &map_err)?;
            }

            writer
                .write_event(Event::End(BytesEnd::new("borders")))
                .map_err(&map_err)?;
        }

        // --- cellXfs ---
        {
            let mut xf_start = BytesStart::new("cellXfs");
            xf_start.push_attribute(("count", ibuf.format(self.cell_xfs.len())));
            writer
                .write_event(Event::Start(xf_start))
                .map_err(&map_err)?;

            for xf in &self.cell_xfs {
                let mut elem = BytesStart::new("xf");
                elem.push_attribute(("numFmtId", ibuf.format(xf.num_fmt_id)));
                elem.push_attribute(("fontId", ibuf.format(xf.font_id)));
                elem.push_attribute(("fillId", ibuf.format(xf.fill_id)));
                elem.push_attribute(("borderId", ibuf.format(xf.border_id)));

                if xf.apply_font {
                    elem.push_attribute(("applyFont", "1"));
                }
                if xf.apply_fill {
                    elem.push_attribute(("applyFill", "1"));
                }
                if xf.apply_border {
                    elem.push_attribute(("applyBorder", "1"));
                }
                if xf.apply_number_format {
                    elem.push_attribute(("applyNumberFormat", "1"));
                }
                if xf.apply_alignment {
                    elem.push_attribute(("applyAlignment", "1"));
                }
                if xf.apply_protection {
                    elem.push_attribute(("applyProtection", "1"));
                }

                let has_children = xf.alignment.is_some() || xf.protection.is_some();
                if has_children {
                    writer
                        .write_event(Event::Start(elem))
                        .map_err(&map_err)?;

                    if let Some(ref alignment) = xf.alignment {
                        write_alignment(&mut writer, alignment, &mut ibuf, &map_err)?;
                    }
                    if let Some(ref protection) = xf.protection {
                        write_protection(&mut writer, protection, &map_err)?;
                    }

                    writer
                        .write_event(Event::End(BytesEnd::new("xf")))
                        .map_err(&map_err)?;
                } else {
                    writer
                        .write_event(Event::Empty(elem))
                        .map_err(&map_err)?;
                }
            }

            writer
                .write_event(Event::End(BytesEnd::new("cellXfs")))
                .map_err(&map_err)?;
        }

        // --- cellStyles ---
        if !self.cell_styles.is_empty() {
            let mut cs_start = BytesStart::new("cellStyles");
            cs_start.push_attribute(("count", ibuf.format(self.cell_styles.len())));
            writer
                .write_event(Event::Start(cs_start))
                .map_err(&map_err)?;

            for cs in &self.cell_styles {
                let mut elem = BytesStart::new("cellStyle");
                elem.push_attribute(("name", cs.name.as_str()));
                elem.push_attribute(("xfId", ibuf.format(cs.xf_id)));
                if let Some(bid) = cs.builtin_id {
                    elem.push_attribute(("builtinId", ibuf.format(bid)));
                }
                writer
                    .write_event(Event::Empty(elem))
                    .map_err(&map_err)?;
            }

            writer
                .write_event(Event::End(BytesEnd::new("cellStyles")))
                .map_err(&map_err)?;
        }

        // --- dxfs ---
        if !self.dxfs.is_empty() {
            let mut d_start = BytesStart::new("dxfs");
            d_start.push_attribute(("count", ibuf.format(self.dxfs.len())));
            writer
                .write_event(Event::Start(d_start))
                .map_err(&map_err)?;

            for dxf in &self.dxfs {
                writer
                    .write_event(Event::Start(BytesStart::new("dxf")))
                    .map_err(&map_err)?;

                if let Some(ref font) = dxf.font {
                    write_font(&mut writer, font, &mut ibuf, &map_err)?;
                }
                if let Some(ref nf) = dxf.num_fmt {
                    let mut elem = BytesStart::new("numFmt");
                    elem.push_attribute(("numFmtId", ibuf.format(nf.id)));
                    elem.push_attribute(("formatCode", nf.format_code.as_str()));
                    writer
                        .write_event(Event::Empty(elem))
                        .map_err(&map_err)?;
                }
                if let Some(ref fill) = dxf.fill {
                    write_fill(&mut writer, fill, &mut ibuf, &map_err)?;
                }
                if let Some(ref border) = dxf.border {
                    write_border(&mut writer, border, &map_err)?;
                }

                writer
                    .write_event(Event::End(BytesEnd::new("dxf")))
                    .map_err(&map_err)?;
            }

            writer
                .write_event(Event::End(BytesEnd::new("dxfs")))
                .map_err(&map_err)?;
        }

        // </styleSheet>
        writer
            .write_event(Event::End(BytesEnd::new("styleSheet")))
            .map_err(&map_err)?;

        Ok(buf)
    }
}

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

/// Extract a `BorderSide` from attributes on a border child element (e.g. `<left style="thin">`).
/// Returns `None` if there is no `style` attribute.
fn parse_border_side_attrs(e: &BytesStart<'_>) -> Option<BorderSide> {
    e.attributes()
        .flatten()
        .find(|attr| attr.key.local_name().as_ref() == b"style")
        .map(|attr| BorderSide {
            style: std::str::from_utf8(&attr.value).unwrap_or_default().to_owned(),
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
            b"applyFont" => {
                xf.apply_font = parse_bool_attr(&attr.value);
            }
            b"applyFill" => {
                xf.apply_fill = parse_bool_attr(&attr.value);
            }
            b"applyBorder" => {
                xf.apply_border = parse_bool_attr(&attr.value);
            }
            b"applyNumberFormat" => {
                xf.apply_number_format = parse_bool_attr(&attr.value);
            }
            b"applyAlignment" => {
                xf.apply_alignment = parse_bool_attr(&attr.value);
            }
            b"applyProtection" => {
                xf.apply_protection = parse_bool_attr(&attr.value);
            }
            _ => {}
        }
    }
    xf
}

/// Parse attributes of an `<alignment>` element.
fn parse_alignment_attrs(e: &BytesStart<'_>) -> Alignment {
    let mut alignment = Alignment::default();
    for attr in e.attributes().flatten() {
        match attr.key.local_name().as_ref() {
            b"horizontal" => {
                alignment.horizontal = Some(
                    std::str::from_utf8(&attr.value).unwrap_or_default().to_owned(),
                );
            }
            b"vertical" => {
                alignment.vertical = Some(
                    std::str::from_utf8(&attr.value).unwrap_or_default().to_owned(),
                );
            }
            b"wrapText" => {
                alignment.wrap_text = parse_bool_attr(&attr.value);
            }
            b"textRotation" => {
                alignment.text_rotation = std::str::from_utf8(&attr.value).unwrap_or_default()
                    .parse()
                    .ok();
            }
            b"indent" => {
                alignment.indent = std::str::from_utf8(&attr.value).unwrap_or_default()
                    .parse()
                    .ok();
            }
            b"shrinkToFit" => {
                alignment.shrink_to_fit = parse_bool_attr(&attr.value);
            }
            _ => {}
        }
    }
    alignment
}

/// Parse attributes of a `<protection>` element.
fn parse_protection_attrs(e: &BytesStart<'_>) -> Protection {
    let mut protection = Protection::default();
    for attr in e.attributes().flatten() {
        match attr.key.local_name().as_ref() {
            b"locked" => {
                protection.locked = parse_bool_attr(&attr.value);
            }
            b"hidden" => {
                protection.hidden = parse_bool_attr(&attr.value);
            }
            _ => {}
        }
    }
    protection
}

/// Parse attributes of a `<cellStyle>` element.
fn parse_cell_style_attrs(e: &BytesStart<'_>) -> CellStyle {
    let mut cs = CellStyle::default();
    for attr in e.attributes().flatten() {
        match attr.key.local_name().as_ref() {
            b"name" => {
                cs.name = std::str::from_utf8(&attr.value).unwrap_or_default().to_owned();
            }
            b"xfId" => {
                cs.xf_id = std::str::from_utf8(&attr.value).unwrap_or_default()
                    .parse()
                    .unwrap_or(0);
            }
            b"builtinId" => {
                cs.builtin_id = std::str::from_utf8(&attr.value).unwrap_or_default()
                    .parse()
                    .ok();
            }
            _ => {}
        }
    }
    cs
}

/// Parse a boolean attribute value ("1", "true" -> true, everything else -> false).
fn parse_bool_attr(value: &[u8]) -> bool {
    let s = std::str::from_utf8(value).unwrap_or_default();
    s == "1" || s == "true"
}

/// Write a single font element.
fn write_font<W: std::io::Write>(
    writer: &mut Writer<W>,
    font: &Font,
    ibuf: &mut itoa::Buffer,
    map_err: &dyn Fn(std::io::Error) -> ModernXlsxError,
) -> Result<()> {
    writer
        .write_event(Event::Start(BytesStart::new("font")))
        .map_err(map_err)?;

    if font.bold {
        writer
            .write_event(Event::Empty(BytesStart::new("b")))
            .map_err(map_err)?;
    }
    if font.italic {
        writer
            .write_event(Event::Empty(BytesStart::new("i")))
            .map_err(map_err)?;
    }
    if font.underline {
        writer
            .write_event(Event::Empty(BytesStart::new("u")))
            .map_err(map_err)?;
    }
    if font.strike {
        writer
            .write_event(Event::Empty(BytesStart::new("strike")))
            .map_err(map_err)?;
    }
    if font.condense {
        writer
            .write_event(Event::Empty(BytesStart::new("condense")))
            .map_err(map_err)?;
    }
    if font.extend {
        writer
            .write_event(Event::Empty(BytesStart::new("extend")))
            .map_err(map_err)?;
    }
    if let Some(sz) = font.size {
        let mut elem = BytesStart::new("sz");
        // Format without trailing zeros for integers.
        let val = if sz.fract() == 0.0 {
            ibuf.format(sz as i64).to_owned()
        } else {
            sz.to_string()
        };
        elem.push_attribute(("val", val.as_str()));
        writer
            .write_event(Event::Empty(elem))
            .map_err(map_err)?;
    }
    if let Some(ref name) = font.name {
        let mut elem = BytesStart::new("name");
        elem.push_attribute(("val", name.as_str()));
        writer
            .write_event(Event::Empty(elem))
            .map_err(map_err)?;
    }
    if let Some(ref color) = font.color {
        let mut elem = BytesStart::new("color");
        elem.push_attribute(("rgb", color.as_str()));
        writer
            .write_event(Event::Empty(elem))
            .map_err(map_err)?;
    }
    if let Some(ref vert_align) = font.vert_align {
        let mut elem = BytesStart::new("vertAlign");
        elem.push_attribute(("val", vert_align.as_str()));
        writer
            .write_event(Event::Empty(elem))
            .map_err(map_err)?;
    }
    if let Some(family) = font.family {
        let mut elem = BytesStart::new("family");
        elem.push_attribute(("val", ibuf.format(family)));
        writer
            .write_event(Event::Empty(elem))
            .map_err(map_err)?;
    }
    if let Some(charset) = font.charset {
        let mut elem = BytesStart::new("charset");
        elem.push_attribute(("val", ibuf.format(charset)));
        writer
            .write_event(Event::Empty(elem))
            .map_err(map_err)?;
    }
    if let Some(ref scheme) = font.scheme {
        let mut elem = BytesStart::new("scheme");
        elem.push_attribute(("val", scheme.as_str()));
        writer
            .write_event(Event::Empty(elem))
            .map_err(map_err)?;
    }

    writer
        .write_event(Event::End(BytesEnd::new("font")))
        .map_err(map_err)?;

    Ok(())
}

/// Write a single fill element.
fn write_fill<W: std::io::Write>(
    writer: &mut Writer<W>,
    fill: &Fill,
    _ibuf: &mut itoa::Buffer,
    map_err: &dyn Fn(std::io::Error) -> ModernXlsxError,
) -> Result<()> {
    writer
        .write_event(Event::Start(BytesStart::new("fill")))
        .map_err(map_err)?;

    if let Some(ref gradient) = fill.gradient_fill {
        // Write gradient fill
        let mut gf = BytesStart::new("gradientFill");
        if let Some(ref gt) = gradient.gradient_type {
            gf.push_attribute(("type", gt.as_str()));
        }
        if let Some(degree) = gradient.degree {
            let deg_str = if degree.fract() == 0.0 {
                format!("{}", degree as i64)
            } else {
                degree.to_string()
            };
            gf.push_attribute(("degree", deg_str.as_str()));
        }

        if gradient.stops.is_empty() {
            writer
                .write_event(Event::Empty(gf))
                .map_err(map_err)?;
        } else {
            writer
                .write_event(Event::Start(gf))
                .map_err(map_err)?;

            for stop in &gradient.stops {
                let mut stop_elem = BytesStart::new("stop");
                let pos_str = stop.position.to_string();
                stop_elem.push_attribute(("position", pos_str.as_str()));
                writer
                    .write_event(Event::Start(stop_elem))
                    .map_err(map_err)?;

                let mut color_elem = BytesStart::new("color");
                color_elem.push_attribute(("rgb", stop.color.as_str()));
                writer
                    .write_event(Event::Empty(color_elem))
                    .map_err(map_err)?;

                writer
                    .write_event(Event::End(BytesEnd::new("stop")))
                    .map_err(map_err)?;
            }

            writer
                .write_event(Event::End(BytesEnd::new("gradientFill")))
                .map_err(map_err)?;
        }
    } else {
        let has_children = fill.fg_color.is_some() || fill.bg_color.is_some();

        if has_children {
            let mut pf = BytesStart::new("patternFill");
            pf.push_attribute(("patternType", fill.pattern_type.as_str()));
            writer
                .write_event(Event::Start(pf))
                .map_err(map_err)?;

            if let Some(ref fg) = fill.fg_color {
                let mut elem = BytesStart::new("fgColor");
                elem.push_attribute(("rgb", fg.as_str()));
                writer
                    .write_event(Event::Empty(elem))
                    .map_err(map_err)?;
            }
            if let Some(ref bg) = fill.bg_color {
                let mut elem = BytesStart::new("bgColor");
                elem.push_attribute(("rgb", bg.as_str()));
                writer
                    .write_event(Event::Empty(elem))
                    .map_err(map_err)?;
            }

            writer
                .write_event(Event::End(BytesEnd::new("patternFill")))
                .map_err(map_err)?;
        } else {
            let mut pf = BytesStart::new("patternFill");
            pf.push_attribute(("patternType", fill.pattern_type.as_str()));
            writer
                .write_event(Event::Empty(pf))
                .map_err(map_err)?;
        }
    }

    writer
        .write_event(Event::End(BytesEnd::new("fill")))
        .map_err(map_err)?;

    Ok(())
}

/// Write a single border element.
fn write_border<W: std::io::Write>(
    writer: &mut Writer<W>,
    border: &Border,
    map_err: &dyn Fn(std::io::Error) -> ModernXlsxError,
) -> Result<()> {
    let mut border_elem = BytesStart::new("border");
    if border.diagonal_up {
        border_elem.push_attribute(("diagonalUp", "1"));
    }
    if border.diagonal_down {
        border_elem.push_attribute(("diagonalDown", "1"));
    }
    writer
        .write_event(Event::Start(border_elem))
        .map_err(map_err)?;

    write_border_side(writer, "left", &border.left, map_err)?;
    write_border_side(writer, "right", &border.right, map_err)?;
    write_border_side(writer, "top", &border.top, map_err)?;
    write_border_side(writer, "bottom", &border.bottom, map_err)?;
    write_border_side(writer, "diagonal", &border.diagonal, map_err)?;

    writer
        .write_event(Event::End(BytesEnd::new("border")))
        .map_err(map_err)?;

    Ok(())
}

/// Write a single border side element (`<left/>`, `<left style="thin"><color rgb="..."/></left>`, etc.).
fn write_border_side<W: std::io::Write>(
    writer: &mut Writer<W>,
    tag: &str,
    side: &Option<BorderSide>,
    map_err: &dyn Fn(std::io::Error) -> ModernXlsxError,
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

/// Write an `<alignment>` element.
fn write_alignment<W: std::io::Write>(
    writer: &mut Writer<W>,
    alignment: &Alignment,
    ibuf: &mut itoa::Buffer,
    map_err: &dyn Fn(std::io::Error) -> ModernXlsxError,
) -> Result<()> {
    let mut elem = BytesStart::new("alignment");
    if let Some(ref h) = alignment.horizontal {
        elem.push_attribute(("horizontal", h.as_str()));
    }
    if let Some(ref v) = alignment.vertical {
        elem.push_attribute(("vertical", v.as_str()));
    }
    if alignment.wrap_text {
        elem.push_attribute(("wrapText", "1"));
    }
    if let Some(rot) = alignment.text_rotation {
        elem.push_attribute(("textRotation", ibuf.format(rot)));
    }
    if let Some(indent) = alignment.indent {
        elem.push_attribute(("indent", ibuf.format(indent)));
    }
    if alignment.shrink_to_fit {
        elem.push_attribute(("shrinkToFit", "1"));
    }
    writer
        .write_event(Event::Empty(elem))
        .map_err(map_err)?;
    Ok(())
}

/// Write a `<protection>` element.
fn write_protection<W: std::io::Write>(
    writer: &mut Writer<W>,
    protection: &Protection,
    map_err: &dyn Fn(std::io::Error) -> ModernXlsxError,
) -> Result<()> {
    let mut elem = BytesStart::new("protection");
    // Only write locked if it differs from default (true).
    if !protection.locked {
        elem.push_attribute(("locked", "0"));
    }
    if protection.hidden {
        elem.push_attribute(("hidden", "1"));
    }
    writer
        .write_event(Event::Empty(elem))
        .map_err(map_err)?;
    Ok(())
}

// ---------------------------------------------------------------------------
// Tests
// ---------------------------------------------------------------------------

#[cfg(test)]
mod tests {
    use super::*;
    use pretty_assertions::assert_eq;

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
        let styles2 = Styles::parse(&xml).unwrap();

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
        let styles2 = Styles::parse(&xml).unwrap();

        assert_eq!(styles2.num_fmts.len(), 0);
        assert_eq!(styles2.fonts.len(), 1);
        assert_eq!(styles2.fills.len(), 2);
        assert_eq!(styles2.borders.len(), 1);
        assert_eq!(styles2.cell_xfs.len(), 1);
        assert_eq!(styles2.fonts[0].name.as_deref(), Some("Aptos"));
        assert_eq!(styles2.fonts[0].size, Some(11.0));
    }

    #[test]
    fn test_parse_apply_flags() {
        let styles = Styles::parse(MINIMAL_STYLES.as_bytes()).unwrap();

        // xf[0] has no apply flags
        assert!(!styles.cell_xfs[0].apply_font);
        assert!(!styles.cell_xfs[0].apply_fill);
        assert!(!styles.cell_xfs[0].apply_number_format);

        // xf[1] has applyFont="1"
        assert!(styles.cell_xfs[1].apply_font);
        assert!(!styles.cell_xfs[1].apply_fill);

        // xf[2] has applyNumberFormat="1"
        assert!(styles.cell_xfs[2].apply_number_format);
        assert!(!styles.cell_xfs[2].apply_font);
    }

    #[test]
    fn test_parse_alignment() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1"><font><sz val="11"/><name val="Aptos"/></font></fonts>
  <fills count="1"><fill><patternFill patternType="none"/></fill></fills>
  <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
  <cellXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" applyAlignment="1">
      <alignment horizontal="center" vertical="bottom" wrapText="1" textRotation="90" indent="2" shrinkToFit="1"/>
    </xf>
  </cellXfs>
</styleSheet>"#;

        let styles = Styles::parse(xml.as_bytes()).unwrap();
        assert_eq!(styles.cell_xfs.len(), 1);

        let xf = &styles.cell_xfs[0];
        assert!(xf.apply_alignment);
        let align = xf.alignment.as_ref().unwrap();
        assert_eq!(align.horizontal.as_deref(), Some("center"));
        assert_eq!(align.vertical.as_deref(), Some("bottom"));
        assert!(align.wrap_text);
        assert_eq!(align.text_rotation, Some(90));
        assert_eq!(align.indent, Some(2));
        assert!(align.shrink_to_fit);

        // Roundtrip
        let xml_out = styles.to_xml().unwrap();
        let styles2 = Styles::parse(&xml_out).unwrap();
        let xf2 = &styles2.cell_xfs[0];
        assert!(xf2.apply_alignment);
        let align2 = xf2.alignment.as_ref().unwrap();
        assert_eq!(align2.horizontal.as_deref(), Some("center"));
        assert_eq!(align2.vertical.as_deref(), Some("bottom"));
        assert!(align2.wrap_text);
        assert_eq!(align2.text_rotation, Some(90));
        assert_eq!(align2.indent, Some(2));
        assert!(align2.shrink_to_fit);
    }

    #[test]
    fn test_parse_protection() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1"><font><sz val="11"/><name val="Aptos"/></font></fonts>
  <fills count="1"><fill><patternFill patternType="none"/></fill></fills>
  <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
  <cellXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" applyProtection="1">
      <protection locked="0" hidden="1"/>
    </xf>
  </cellXfs>
</styleSheet>"#;

        let styles = Styles::parse(xml.as_bytes()).unwrap();
        let xf = &styles.cell_xfs[0];
        assert!(xf.apply_protection);
        let prot = xf.protection.as_ref().unwrap();
        assert!(!prot.locked);
        assert!(prot.hidden);

        // Roundtrip
        let xml_out = styles.to_xml().unwrap();
        let styles2 = Styles::parse(&xml_out).unwrap();
        let prot2 = styles2.cell_xfs[0].protection.as_ref().unwrap();
        assert!(!prot2.locked);
        assert!(prot2.hidden);
    }

    #[test]
    fn test_parse_diagonal_border() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1"><font><sz val="11"/><name val="Aptos"/></font></fonts>
  <fills count="1"><fill><patternFill patternType="none"/></fill></fills>
  <borders count="1">
    <border diagonalUp="1" diagonalDown="1">
      <left/>
      <right/>
      <top/>
      <bottom/>
      <diagonal style="thin">
        <color rgb="FFFF0000"/>
      </diagonal>
    </border>
  </borders>
  <cellXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
  </cellXfs>
</styleSheet>"#;

        let styles = Styles::parse(xml.as_bytes()).unwrap();
        let border = &styles.borders[0];
        assert!(border.diagonal_up);
        assert!(border.diagonal_down);
        let diag = border.diagonal.as_ref().unwrap();
        assert_eq!(diag.style, "thin");
        assert_eq!(diag.color.as_deref(), Some("FFFF0000"));

        // Roundtrip
        let xml_out = styles.to_xml().unwrap();
        let styles2 = Styles::parse(&xml_out).unwrap();
        let border2 = &styles2.borders[0];
        assert!(border2.diagonal_up);
        assert!(border2.diagonal_down);
        let diag2 = border2.diagonal.as_ref().unwrap();
        assert_eq!(diag2.style, "thin");
        assert_eq!(diag2.color.as_deref(), Some("FFFF0000"));
    }

    #[test]
    fn test_parse_gradient_fill() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1"><font><sz val="11"/><name val="Aptos"/></font></fonts>
  <fills count="2">
    <fill><patternFill patternType="none"/></fill>
    <fill>
      <gradientFill type="linear" degree="90">
        <stop position="0">
          <color rgb="FF000000"/>
        </stop>
        <stop position="1">
          <color rgb="FFFFFFFF"/>
        </stop>
      </gradientFill>
    </fill>
  </fills>
  <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
  <cellXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
  </cellXfs>
</styleSheet>"#;

        let styles = Styles::parse(xml.as_bytes()).unwrap();
        assert_eq!(styles.fills.len(), 2);

        let grad_fill = &styles.fills[1];
        let gradient = grad_fill.gradient_fill.as_ref().unwrap();
        assert_eq!(gradient.gradient_type.as_deref(), Some("linear"));
        assert_eq!(gradient.degree, Some(90.0));
        assert_eq!(gradient.stops.len(), 2);
        assert_eq!(gradient.stops[0].position, 0.0);
        assert_eq!(gradient.stops[0].color, "FF000000");
        assert_eq!(gradient.stops[1].position, 1.0);
        assert_eq!(gradient.stops[1].color, "FFFFFFFF");

        // Roundtrip
        let xml_out = styles.to_xml().unwrap();
        let styles2 = Styles::parse(&xml_out).unwrap();
        let gradient2 = styles2.fills[1].gradient_fill.as_ref().unwrap();
        assert_eq!(gradient2.gradient_type.as_deref(), Some("linear"));
        assert_eq!(gradient2.degree, Some(90.0));
        assert_eq!(gradient2.stops.len(), 2);
        assert_eq!(gradient2.stops[0].color, "FF000000");
        assert_eq!(gradient2.stops[1].color, "FFFFFFFF");
    }

    #[test]
    fn test_parse_font_extended_properties() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1">
    <font>
      <b/>
      <sz val="11"/>
      <name val="Aptos"/>
      <color rgb="FF000000"/>
      <vertAlign val="superscript"/>
      <family val="2"/>
      <charset val="1"/>
      <scheme val="minor"/>
      <condense/>
      <extend/>
    </font>
  </fonts>
  <fills count="1"><fill><patternFill patternType="none"/></fill></fills>
  <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
  <cellXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
  </cellXfs>
</styleSheet>"#;

        let styles = Styles::parse(xml.as_bytes()).unwrap();
        let font = &styles.fonts[0];
        assert!(font.bold);
        assert_eq!(font.vert_align.as_deref(), Some("superscript"));
        assert_eq!(font.family, Some(2));
        assert_eq!(font.charset, Some(1));
        assert_eq!(font.scheme.as_deref(), Some("minor"));
        assert!(font.condense);
        assert!(font.extend);

        // Roundtrip
        let xml_out = styles.to_xml().unwrap();
        let styles2 = Styles::parse(&xml_out).unwrap();
        let font2 = &styles2.fonts[0];
        assert_eq!(font2.vert_align.as_deref(), Some("superscript"));
        assert_eq!(font2.family, Some(2));
        assert_eq!(font2.charset, Some(1));
        assert_eq!(font2.scheme.as_deref(), Some("minor"));
        assert!(font2.condense);
        assert!(font2.extend);
    }

    #[test]
    fn test_parse_dxf_styles() {
        let xml = r##"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1"><font><sz val="11"/><name val="Aptos"/></font></fonts>
  <fills count="1"><fill><patternFill patternType="none"/></fill></fills>
  <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
  <cellXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
  </cellXfs>
  <dxfs count="2">
    <dxf>
      <font><b/><color rgb="FFFF0000"/></font>
      <fill><patternFill patternType="solid"><fgColor rgb="FFFFFF00"/></patternFill></fill>
    </dxf>
    <dxf>
      <numFmt numFmtId="164" formatCode="#,##0.00"/>
      <border>
        <left style="thin"><color rgb="FF000000"/></left>
        <right/><top/><bottom/><diagonal/>
      </border>
    </dxf>
  </dxfs>
</styleSheet>"##;

        let styles = Styles::parse(xml.as_bytes()).unwrap();
        assert_eq!(styles.dxfs.len(), 2);

        // First DXF: font + fill
        let dxf0 = &styles.dxfs[0];
        let font = dxf0.font.as_ref().unwrap();
        assert!(font.bold);
        assert_eq!(font.color.as_deref(), Some("FFFF0000"));
        let fill = dxf0.fill.as_ref().unwrap();
        assert_eq!(fill.pattern_type, "solid");
        assert_eq!(fill.fg_color.as_deref(), Some("FFFFFF00"));

        // Second DXF: numFmt + border
        let dxf1 = &styles.dxfs[1];
        let nf = dxf1.num_fmt.as_ref().unwrap();
        assert_eq!(nf.id, 164);
        assert_eq!(nf.format_code, "#,##0.00");
        let border = dxf1.border.as_ref().unwrap();
        let left = border.left.as_ref().unwrap();
        assert_eq!(left.style, "thin");
        assert_eq!(left.color.as_deref(), Some("FF000000"));

        // Roundtrip
        let xml_out = styles.to_xml().unwrap();
        let styles2 = Styles::parse(&xml_out).unwrap();
        assert_eq!(styles2.dxfs.len(), 2);
        let dxf0_rt = &styles2.dxfs[0];
        assert!(dxf0_rt.font.as_ref().unwrap().bold);
        assert_eq!(dxf0_rt.fill.as_ref().unwrap().fg_color.as_deref(), Some("FFFFFF00"));
        let dxf1_rt = &styles2.dxfs[1];
        assert_eq!(dxf1_rt.num_fmt.as_ref().unwrap().format_code, "#,##0.00");
    }

    #[test]
    fn test_parse_cell_styles() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1"><font><sz val="11"/><name val="Aptos"/></font></fonts>
  <fills count="1"><fill><patternFill patternType="none"/></fill></fills>
  <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
  <cellXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
  </cellXfs>
  <cellStyles count="2">
    <cellStyle name="Normal" xfId="0" builtinId="0"/>
    <cellStyle name="Custom" xfId="1"/>
  </cellStyles>
</styleSheet>"#;

        let styles = Styles::parse(xml.as_bytes()).unwrap();
        assert_eq!(styles.cell_styles.len(), 2);
        assert_eq!(styles.cell_styles[0].name, "Normal");
        assert_eq!(styles.cell_styles[0].xf_id, 0);
        assert_eq!(styles.cell_styles[0].builtin_id, Some(0));
        assert_eq!(styles.cell_styles[1].name, "Custom");
        assert_eq!(styles.cell_styles[1].xf_id, 1);
        assert_eq!(styles.cell_styles[1].builtin_id, None);

        // Roundtrip
        let xml_out = styles.to_xml().unwrap();
        let styles2 = Styles::parse(&xml_out).unwrap();
        assert_eq!(styles2.cell_styles.len(), 2);
        assert_eq!(styles2.cell_styles[0].name, "Normal");
        assert_eq!(styles2.cell_styles[0].builtin_id, Some(0));
        assert_eq!(styles2.cell_styles[1].name, "Custom");
        assert_eq!(styles2.cell_styles[1].builtin_id, None);
    }

    #[test]
    fn test_apply_flags_roundtrip() {
        let styles1 = Styles::parse(MINIMAL_STYLES.as_bytes()).unwrap();

        // xf[1] has applyFont="1"
        assert!(styles1.cell_xfs[1].apply_font);
        // xf[2] has applyNumberFormat="1"
        assert!(styles1.cell_xfs[2].apply_number_format);

        let xml = styles1.to_xml().unwrap();
        let styles2 = Styles::parse(&xml).unwrap();

        assert!(styles2.cell_xfs[1].apply_font);
        assert!(!styles2.cell_xfs[1].apply_fill);
        assert!(styles2.cell_xfs[2].apply_number_format);
        assert!(!styles2.cell_xfs[2].apply_font);
    }

    #[test]
    fn test_protection_default_values() {
        let prot = Protection::default();
        assert!(prot.locked);
        assert!(!prot.hidden);
    }

    #[test]
    fn test_dxf_with_gradient_fill_roundtrip() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1"><font><sz val="11"/><name val="Aptos"/></font></fonts>
  <fills count="1"><fill><patternFill patternType="none"/></fill></fills>
  <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
  <cellXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
  </cellXfs>
  <dxfs count="1">
    <dxf>
      <fill>
        <gradientFill type="linear" degree="45">
          <stop position="0">
            <color rgb="FFFF0000"/>
          </stop>
          <stop position="1">
            <color rgb="FF00FF00"/>
          </stop>
        </gradientFill>
      </fill>
    </dxf>
  </dxfs>
</styleSheet>"#;

        let styles = Styles::parse(xml.as_bytes()).unwrap();
        assert_eq!(styles.dxfs.len(), 1);
        let fill = styles.dxfs[0].fill.as_ref().unwrap();
        let gradient = fill.gradient_fill.as_ref().unwrap();
        assert_eq!(gradient.gradient_type.as_deref(), Some("linear"));
        assert_eq!(gradient.degree, Some(45.0));
        assert_eq!(gradient.stops.len(), 2);
        assert_eq!(gradient.stops[0].color, "FFFF0000");
        assert_eq!(gradient.stops[1].color, "FF00FF00");

        // Roundtrip
        let xml_out = styles.to_xml().unwrap();
        let styles2 = Styles::parse(&xml_out).unwrap();
        let fill2 = styles2.dxfs[0].fill.as_ref().unwrap();
        let gradient2 = fill2.gradient_fill.as_ref().unwrap();
        assert_eq!(gradient2.stops.len(), 2);
        assert_eq!(gradient2.stops[0].color, "FFFF0000");
    }
}
