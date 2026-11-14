//! `xl/styles.xml` (ECMA-376 Part 1, §18.8): number formats, fonts, fills,
//! borders, the cell-style and cell formatting records, named cell styles and
//! differential formats. Read with a SAX parser and written back in
//! `CT_Stylesheet` element order; `tableStyles`, `colors` and the root
//! `extLst` are kept verbatim.

use core::hint::cold_path;

use quick_xml::events::{BytesEnd, BytesStart, Event};
use quick_xml::{Reader, Writer};
use serde::{Deserialize, Serialize};

use super::color::{Color, resolve_color, split_color, write_color};
use super::xml::{
    self, XsdBuf, attr_str, attrs, ct_boolean, fmt_xsd_double, parse_attr, val_attr, val_string,
    xml_attr, xsd_bool,
};
use crate::ooxml::shared_strings::ExpansionBudget;
use crate::ooxml::{write_xml_decl, xml_write_err};
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
    #[serde(default, skip_serializing_if = "crate::ooxml::is_false")]
    pub wrap_text: bool,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub text_rotation: Option<u16>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub indent: Option<u32>,
    #[serde(default, skip_serializing_if = "crate::ooxml::is_false")]
    pub shrink_to_fit: bool,
    /// `@readingOrder`: 0 context-dependent, 1 left-to-right, 2 right-to-left.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub reading_order: Option<u32>,
    /// `@justifyLastLine`: justify the last line of distributed text.
    #[serde(default, skip_serializing_if = "crate::ooxml::is_false")]
    pub justify_last_line: bool,
    /// `@relativeIndent`: indent change applied by a differential format.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub relative_indent: Option<i32>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct Protection {
    #[serde(default = "crate::ooxml::default_true")]
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

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct Font {
    pub name: Option<String>,
    pub size: Option<f64>,
    pub bold: bool,
    pub italic: bool,
    pub underline: bool,
    /// The `<u val>` underline type when it is not `single`: `double`,
    /// `singleAccounting` or `doubleAccounting`. Written as `<u val="..."/>`
    /// whenever set; with `None`, `underline` writes a single `<u/>`.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub underline_style: Option<String>,
    pub strike: bool,
    /// The `rgb` of the font colour, an ARGB hex string such as "FF000000"
    /// (RGB and a leading `#` are written as ARGB; any other non-empty value
    /// fails the write with `InvalidWorkbook`, as do the other colour strings).
    pub color: Option<String>,
    /// The full font colour, set when it has a theme, indexed, tint or auto
    /// attribute. On write it is used when its `rgb` equals `color` (an empty
    /// string counts as absent); otherwise `color` is written as plain `rgb`.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub color_spec: Option<Color>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub vert_align: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub family: Option<u32>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub charset: Option<u32>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub scheme: Option<String>,
    #[serde(default, skip_serializing_if = "crate::ooxml::is_false")]
    pub condense: bool,
    #[serde(default, skip_serializing_if = "crate::ooxml::is_false")]
    pub extend: bool,
    /// The boolean properties written explicitly off (`b`, `i`, `strike`,
    /// `condense`, `extend` as `<tag val="0"/>`, `u` as `<u val="none"/>`), so a
    /// dxf keeps its override instead of inheriting the cell's value. Filled
    /// when reading a dxf font; on write each listed tag whose flag is off is
    /// emitted explicitly.
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub explicit_off: Vec<String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct GradientStop {
    pub position: f64,
    /// The `rgb` of the stop colour, or "" when the colour has none.
    pub color: String,
    /// The full stop colour when it has a theme, indexed, tint or auto
    /// attribute; same write rule as [`Font::color_spec`].
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub color_spec: Option<Color>,
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
    /// `ST_PatternType`; "" when the fill has no `<patternFill>` or it has no
    /// `patternType` (written as `<patternFill/>`).
    pub pattern_type: String,
    /// The `rgb` of `<fgColor>`.
    pub fg_color: Option<String>,
    /// The full `<fgColor>`; same write rule as [`Font::color_spec`].
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub fg_color_spec: Option<Color>,
    /// The `rgb` of `<bgColor>`.
    pub bg_color: Option<String>,
    /// The full `<bgColor>`; same write rule as [`Font::color_spec`].
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub bg_color_spec: Option<Color>,
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
    #[serde(default, skip_serializing_if = "crate::ooxml::is_false")]
    pub diagonal_up: bool,
    #[serde(default, skip_serializing_if = "crate::ooxml::is_false")]
    pub diagonal_down: bool,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct BorderSide {
    pub style: String,
    /// The `rgb` of the side colour.
    pub color: Option<String>,
    /// The full side colour; same write rule as [`Font::color_spec`].
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub color_spec: Option<Color>,
}

/// One `<xf>` record of `cellXfs` or `cellStyleXfs`.
#[derive(Debug, Clone, Default, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct CellXf {
    pub num_fmt_id: u32,
    pub font_id: u32,
    pub fill_id: u32,
    pub border_id: u32,
    /// `@xfId`: the `cellStyleXfs` record a `cellXfs` record inherits from.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub xf_id: Option<u32>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub alignment: Option<Alignment>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub protection: Option<Protection>,
    #[serde(default, skip_serializing_if = "crate::ooxml::is_false")]
    pub apply_font: bool,
    #[serde(default, skip_serializing_if = "crate::ooxml::is_false")]
    pub apply_fill: bool,
    #[serde(default, skip_serializing_if = "crate::ooxml::is_false")]
    pub apply_border: bool,
    #[serde(default, skip_serializing_if = "crate::ooxml::is_false")]
    pub apply_number_format: bool,
    #[serde(default, skip_serializing_if = "crate::ooxml::is_false")]
    pub apply_alignment: bool,
    #[serde(default, skip_serializing_if = "crate::ooxml::is_false")]
    pub apply_protection: bool,
    /// CT_Xf quotePrefix: the cell text is shown as entered (leading apostrophe).
    #[serde(default, skip_serializing_if = "crate::ooxml::is_false")]
    pub quote_prefix: bool,
    /// CT_Xf pivotButton.
    #[serde(default, skip_serializing_if = "crate::ooxml::is_false")]
    pub pivot_button: bool,
}

/// A differential format (`<dxf>`), written in `CT_Dxf` order: font, numFmt,
/// fill, alignment, border, protection.
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
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub alignment: Option<Alignment>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub protection: Option<Protection>,
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
    /// The `cellStyleXfs` records that `cellStyles` and `CellXf::xf_id` point
    /// at. When empty, the writer emits one default record at index 0, the target
    /// of `xfId="0"`.
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub cell_style_xfs: Vec<CellXf>,
    pub cell_xfs: Vec<CellXf>,
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub dxfs: Vec<DxfStyle>,
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub cell_styles: Vec<CellStyle>,
    /// The source `<tableStyles>` element, verbatim.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub table_styles_xml: Option<String>,
    /// The source `<colors>` element (indexed palette, MRU colours), verbatim.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub colors_xml: Option<String>,
    /// The source `<styleSheet>`-level `<extLst>` element, verbatim, with the
    /// namespace declarations it inherited from `<styleSheet>` added to it.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ext_lst_xml: Option<String>,
}

// ---------------------------------------------------------------------------
// Parser state
// ---------------------------------------------------------------------------

#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
enum Section {
    #[default]
    None,
    NumFmts,
    Fonts,
    Fills,
    Borders,
    CellStyleXfs,
    CellXfs,
    CellStyles,
    Dxfs,
}

impl Section {
    /// The section a `<styleSheet>` list element opens.
    fn of(local: &str) -> Option<Self> {
        Some(match local {
            "numFmts" => Self::NumFmts,
            "fonts" => Self::Fonts,
            "fills" => Self::Fills,
            "borders" => Self::Borders,
            "cellStyleXfs" => Self::CellStyleXfs,
            "cellXfs" => Self::CellXfs,
            "cellStyles" => Self::CellStyles,
            "dxfs" => Self::Dxfs,
            _ => return None,
        })
    }
}

/// A side of a `<border>`.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum Side {
    Left,
    Right,
    Top,
    Bottom,
    Diagonal,
}

impl Side {
    fn of(local: &str) -> Option<Self> {
        Some(match local {
            "left" => Self::Left,
            "right" => Self::Right,
            "top" => Self::Top,
            "bottom" => Self::Bottom,
            "diagonal" => Self::Diagonal,
            _ => return None,
        })
    }

    fn slot(self, border: &mut Border) -> &mut Option<BorderSide> {
        match self {
            Self::Left => &mut border.left,
            Self::Right => &mut border.right,
            Self::Top => &mut border.top,
            Self::Bottom => &mut border.bottom,
            Self::Diagonal => &mut border.diagonal,
        }
    }
}

/// SAX state for [`Styles::parse`]: the open section and the entries being
/// built. A self-closing element runs [`StylesParser::open`] then
/// [`StylesParser::close`], so `<font/>` adds an entry exactly as
/// `<font></font>` does and every later index stays aligned.
///
/// Each record is counted in `budget` ([`ExpansionBudget::store`]) before it
/// is kept: a number format, cell style, font, fill, border, cell format,
/// differential format or gradient stop at the size of its struct plus the
/// heap bytes of its strings (a font, fill or border inside a differential
/// format is inline in it, so only its strings are added), and each tag a
/// differential font lists in [`Font::explicit_off`] at the size of a
/// `String` plus its bytes. An empty record takes from 5 bytes (`<xf/>`) to
/// 12 bytes (`<cellStyle/>`) of XML and up to 1,200 bytes of memory (a
/// `DxfStyle` on a 64-bit target, measured 2026-11-13).
struct StylesParser {
    out: Styles,
    budget: ExpansionBudget,
    section: Section,
    font: Option<Font>,
    fill: Option<Fill>,
    gradient: Option<GradientFill>,
    stop: Option<GradientStop>,
    border: Option<Border>,
    side: Option<(Side, Option<BorderSide>)>,
    xf: Option<CellXf>,
    dxf: Option<DxfStyle>,
}

/// Heap bytes of an optional string.
fn str_heap(s: Option<&String>) -> usize {
    s.map_or(0, String::capacity)
}

/// Heap bytes of an optional full colour: its `rgb`.
fn color_heap(c: Option<&Color>) -> usize {
    str_heap(c.and_then(|c| c.rgb.as_ref()))
}

/// Heap bytes of a font's strings. [`Font::explicit_off`] is counted tag by
/// tag as it grows.
fn font_heap(f: &Font) -> usize {
    str_heap(f.name.as_ref())
        + str_heap(f.underline_style.as_ref())
        + str_heap(f.color.as_ref())
        + color_heap(f.color_spec.as_ref())
        + str_heap(f.vert_align.as_ref())
        + str_heap(f.scheme.as_ref())
}

/// Heap bytes of a fill's strings. Its gradient stops are counted stop by
/// stop as they are kept.
fn fill_heap(f: &Fill) -> usize {
    f.pattern_type.capacity()
        + str_heap(f.fg_color.as_ref())
        + color_heap(f.fg_color_spec.as_ref())
        + str_heap(f.bg_color.as_ref())
        + color_heap(f.bg_color_spec.as_ref())
        + f.gradient_fill
            .as_ref()
            .map_or(0, |g| str_heap(g.gradient_type.as_ref()))
}

/// Heap bytes of a border's strings, its five sides together.
fn border_heap(b: &Border) -> usize {
    [&b.left, &b.right, &b.top, &b.bottom, &b.diagonal]
        .into_iter()
        .flatten()
        .map(|s| {
            s.style.capacity() + str_heap(s.color.as_ref()) + color_heap(s.color_spec.as_ref())
        })
        .sum()
}

/// Heap bytes of an optional alignment's strings.
fn alignment_heap(a: Option<&Alignment>) -> usize {
    a.map_or(0, |a| {
        str_heap(a.horizontal.as_ref()) + str_heap(a.vertical.as_ref())
    })
}

/// Heap bytes of a differential format's strings, its inline font, fill,
/// border, number format and alignment together.
fn dxf_heap(d: &DxfStyle) -> usize {
    d.font.as_ref().map_or(0, font_heap)
        + d.fill.as_ref().map_or(0, fill_heap)
        + d.border.as_ref().map_or(0, border_heap)
        + d.num_fmt.as_ref().map_or(0, |n| n.format_code.capacity())
        + alignment_heap(d.alignment.as_ref())
}

impl StylesParser {
    fn new(budget: ExpansionBudget) -> Self {
        Self {
            budget,
            out: Styles {
                num_fmts: Vec::new(),
                fonts: Vec::new(),
                fills: Vec::new(),
                borders: Vec::new(),
                cell_style_xfs: Vec::new(),
                cell_xfs: Vec::new(),
                dxfs: Vec::new(),
                cell_styles: Vec::new(),
                table_styles_xml: None,
                colors_xml: None,
                ext_lst_xml: None,
            },
            section: Section::None,
            font: None,
            fill: None,
            gradient: None,
            stop: None,
            border: None,
            side: None,
            xf: None,
            dxf: None,
        }
    }

    /// A `<dxf>` is open and none of its font, fill or border children is.
    fn in_dxf_body(&self) -> bool {
        self.dxf.is_some() && self.font.is_none() && self.fill.is_none() && self.border.is_none()
    }

    /// Handle a start or self-closing tag.
    fn open(&mut self, e: &BytesStart<'_>) -> Result<()> {
        let name = e.local_name();
        let local = name.as_ref();
        if let Some(section) = Section::of(local) {
            self.section = section;
            return Ok(());
        }
        match local {
            "numFmt" if self.section == Section::NumFmts => {
                let nf = parse_num_fmt(e)?;
                self.budget
                    .store(size_of::<NumFmt>() + nf.format_code.capacity())?;
                self.out.num_fmts.push(nf);
            }
            "numFmt" if self.in_dxf_body() => {
                let nf = parse_num_fmt(e)?;
                if let Some(dxf) = &mut self.dxf {
                    dxf.num_fmt = Some(nf);
                }
            }
            "font"
                if self.font.is_none()
                    && (self.section == Section::Fonts || self.in_dxf_body()) =>
            {
                self.font = Some(Font::default());
            }
            "fill"
                if self.fill.is_none()
                    && (self.section == Section::Fills || self.in_dxf_body()) =>
            {
                self.fill = Some(Fill::default());
            }
            "border"
                if self.border.is_none()
                    && (self.section == Section::Borders || self.in_dxf_body()) =>
            {
                self.border = Some(parse_border_attrs(e)?);
            }
            "dxf" if self.section == Section::Dxfs && self.dxf.is_none() => {
                self.dxf = Some(DxfStyle::default());
            }
            "xf" if matches!(self.section, Section::CellXfs | Section::CellStyleXfs) => {
                self.xf = Some(parse_xf_attrs(e)?);
            }
            "cellStyle" if self.section == Section::CellStyles => {
                let cs = parse_cell_style_attrs(e)?;
                self.budget
                    .store(size_of::<CellStyle>() + cs.name.capacity())?;
                self.out.cell_styles.push(cs);
            }
            "alignment" if self.xf.is_some() || self.in_dxf_body() => {
                let alignment = Some(parse_alignment_attrs(e)?);
                if let Some(xf) = &mut self.xf {
                    xf.alignment = alignment;
                } else if let Some(dxf) = &mut self.dxf {
                    dxf.alignment = alignment;
                }
            }
            "protection" if self.xf.is_some() || self.in_dxf_body() => {
                let protection = Some(parse_protection_attrs(e)?);
                if let Some(xf) = &mut self.xf {
                    xf.protection = protection;
                } else if let Some(dxf) = &mut self.dxf {
                    dxf.protection = protection;
                }
            }
            "patternFill" => {
                if let Some(fill) = &mut self.fill {
                    for attr in attrs(e) {
                        let attr = attr?;
                        if attr.key.local_name().as_ref() == "patternType" {
                            fill.pattern_type = attr_str(&attr)?.into_owned();
                        }
                    }
                }
            }
            "fgColor" | "bgColor" => {
                if let Some(fill) = &mut self.fill {
                    let (rgb, spec) = split_color(Color::from_attrs(e)?);
                    if local == "fgColor" {
                        fill.fg_color = rgb;
                        fill.fg_color_spec = spec;
                    } else {
                        fill.bg_color = rgb;
                        fill.bg_color_spec = spec;
                    }
                }
            }
            "gradientFill" if self.fill.is_some() => {
                self.gradient = Some(parse_gradient_attrs(e)?);
            }
            "stop" if self.gradient.is_some() => {
                let mut stop = GradientStop::default();
                for attr in attrs(e) {
                    let attr = attr?;
                    if attr.key.local_name().as_ref() == "position" {
                        stop.position = parse_attr("stop", &attr)?;
                    }
                }
                self.stop = Some(stop);
            }
            "color" => self.color(e)?,
            _ => {
                if let Some(side) = Side::of(local) {
                    if self.border.is_some() && self.side.is_none() {
                        self.side = Some((side, parse_border_side_attrs(e)?));
                    }
                } else if let Some(font) = &mut self.font {
                    font_property(font, local, e, self.dxf.is_some(), &mut self.budget)?;
                }
            }
        }
        Ok(())
    }

    /// A `<color>`: a gradient stop's, else a border side's, else a font's.
    fn color(&mut self, e: &BytesStart<'_>) -> Result<()> {
        if self.stop.is_none() && self.side.is_none() && self.font.is_none() {
            return Ok(());
        }
        let (rgb, spec) = split_color(Color::from_attrs(e)?);
        if let Some(stop) = &mut self.stop {
            stop.color = rgb.unwrap_or_default();
            stop.color_spec = spec;
        } else if let Some((_, side)) = &mut self.side {
            // A side without a style is no border; its colour is dropped.
            if let Some(side) = side {
                side.color = rgb;
                side.color_spec = spec;
            }
        } else if let Some(font) = &mut self.font {
            font.color = rgb;
            font.color_spec = spec;
        }
        Ok(())
    }

    /// Handle an end tag, or the end of a self-closing tag. A record kept in
    /// a list is counted in `budget` first.
    ///
    /// # Errors
    ///
    /// `SizeLimit` once the records counted pass the budget's `store_limit`.
    fn close(&mut self, local: &str) -> Result<()> {
        if Section::of(local).is_some_and(|s| s == self.section) {
            self.section = Section::None;
            return Ok(());
        }
        match local {
            "font" => {
                if let Some(font) = self.font.take() {
                    if let Some(dxf) = &mut self.dxf {
                        dxf.font = Some(font);
                    } else {
                        self.budget.store(size_of::<Font>() + font_heap(&font))?;
                        self.out.fonts.push(font);
                    }
                }
            }
            "fill" => {
                if let Some(fill) = self.fill.take() {
                    if let Some(dxf) = &mut self.dxf {
                        dxf.fill = Some(fill);
                    } else {
                        self.budget.store(size_of::<Fill>() + fill_heap(&fill))?;
                        self.out.fills.push(fill);
                    }
                }
            }
            "border" => {
                if let Some(border) = self.border.take() {
                    if let Some(dxf) = &mut self.dxf {
                        dxf.border = Some(border);
                    } else {
                        self.budget
                            .store(size_of::<Border>() + border_heap(&border))?;
                        self.out.borders.push(border);
                    }
                }
            }
            "gradientFill" => {
                if let Some(gradient) = self.gradient.take()
                    && let Some(fill) = &mut self.fill
                {
                    fill.gradient_fill = Some(gradient);
                }
            }
            "stop" => {
                if let Some(stop) = self.stop.take()
                    && let Some(gradient) = &mut self.gradient
                {
                    self.budget.store(
                        size_of::<GradientStop>()
                            + stop.color.capacity()
                            + color_heap(stop.color_spec.as_ref()),
                    )?;
                    gradient.stops.push(stop);
                }
            }
            "xf" => {
                if let Some(xf) = self.xf.take() {
                    self.budget
                        .store(size_of::<CellXf>() + alignment_heap(xf.alignment.as_ref()))?;
                    if self.section == Section::CellStyleXfs {
                        self.out.cell_style_xfs.push(xf);
                    } else {
                        self.out.cell_xfs.push(xf);
                    }
                }
            }
            "dxf" => {
                if let Some(dxf) = self.dxf.take() {
                    self.budget.store(size_of::<DxfStyle>() + dxf_heap(&dxf))?;
                    self.out.dxfs.push(dxf);
                }
            }
            _ => {
                if let Some(side) = Side::of(local)
                    && let Some(border) = &mut self.border
                    && let Some((open, value)) = self.side.take_if(|(s, _)| *s == side)
                {
                    *open.slot(border) = value;
                }
            }
        }
        Ok(())
    }
}

/// The index in [`VerbatimSpans`] of the `<styleSheet>` child kept verbatim,
/// by local name: `tableStyles`, `colors`, `extLst`.
fn verbatim_slot(local: &str) -> Option<usize> {
    match local {
        "tableStyles" => Some(0),
        "colors" => Some(1),
        "extLst" => Some(2),
        _ => None,
    }
}

/// The source text of the last `<tableStyles>`, `<colors>` and `<extLst>`
/// child of `<styleSheet>` read so far, indexed by [`verbatim_slot`]. A part
/// may repeat a child and the last one wins, so each is made self-contained
/// once per root, by [`flush_verbatim`], not per occurrence, which would
/// copy the root's declarations it uses once per occurrence.
type VerbatimSpans<'d> = [Option<&'d str>; 3];

/// The source text of one whole element, `data[start..end]` without the
/// white space before its tag.
///
/// # Errors
///
/// `XmlParse` when the span is out of range or not UTF-8.
fn verbatim(data: &[u8], start: u64, end: u64) -> Result<&str> {
    usize::try_from(start)
        .ok()
        .zip(usize::try_from(end).ok())
        .and_then(|(s, e)| data.get(s..e))
        .and_then(|bytes| std::str::from_utf8(bytes).ok())
        .map(|text| text.trim_start_matches(xml::XML_S))
        .ok_or_else(|| {
            cold_path();
            ModernXlsxError::XmlParse(format!(
                "styles.xml: element at bytes {start}..{end} is not in-range UTF-8"
            ))
        })
}

/// Keep each span of `spans` in its `styles` field, made self-contained
/// against the namespace declarations of `<styleSheet>` so it can be
/// written back verbatim, and empty `spans`.
fn flush_verbatim(
    styles: &mut Styles,
    spans: &mut VerbatimSpans<'_>,
    root_ns: &xml::RootNamespaces,
) {
    let slots = [
        &mut styles.table_styles_xml,
        &mut styles.colors_xml,
        &mut styles.ext_lst_xml,
    ];
    for (slot, span) in slots.into_iter().zip(spans.iter_mut()) {
        if let Some(text) = span.take() {
            *slot = Some(xml::self_contained_fragment(text, root_ns));
        }
    }
}

impl Default for Styles {
    /// The same styles as [`Styles::default_styles`].
    ///
    /// A derived (empty) `Styles` would hold no font, fill, border or cell
    /// format for `styles.xml` to write.
    fn default() -> Self {
        Self::default_styles()
    }
}

impl Styles {
    // -----------------------------------------------------------------------
    // Factory
    // -----------------------------------------------------------------------

    /// Return the default styles: one 11-point Calibri font, the `none` and
    /// `gray125` fills, an empty border and one cell format.
    /// It must equal the TypeScript `defaultStyles()`; the streaming writer's
    /// `MINIMAL_STYLES_XML` holds the same records plus one default `cellStyleXfs`
    /// entry and `xfId="0"` on its cell format (`__tests__/default-styles.test.ts`).
    pub fn default_styles() -> Self {
        Styles {
            num_fmts: Vec::new(),
            fonts: vec![Font {
                name: Some("Calibri".to_owned()),
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
            cell_style_xfs: Vec::new(),
            cell_xfs: vec![CellXf::default()],
            dxfs: Vec::new(),
            cell_styles: Vec::new(),
            table_styles_xml: None,
            colors_xml: None,
            ext_lst_xml: None,
        }
    }

    // -----------------------------------------------------------------------
    // Parser
    // -----------------------------------------------------------------------

    /// Parse a `styles.xml` document from raw XML bytes, with no limit on
    /// the memory of the records it keeps.
    ///
    /// # Errors
    ///
    /// `XmlParse` for malformed XML, a malformed attribute list, or a numeric
    /// attribute (an id, a size, a position) that does not parse.
    pub fn parse(data: &[u8]) -> Result<Self> {
        Self::parse_within(data, &mut ExpansionBudget::new(usize::MAX))
    }

    /// [`Styles::parse`], counting each record it keeps in `budget`
    /// ([`ExpansionBudget::store`]) before it is kept, so the list elements
    /// of the read's worksheets count after the styles in the same `stored`.
    ///
    /// # Errors
    ///
    /// As [`Styles::parse`], and `SizeLimit` once the records counted pass
    /// `budget.store_limit`.
    pub(crate) fn parse_within(data: &[u8], budget: &mut ExpansionBudget) -> Result<Self> {
        // quick-xml skips a UTF-8 BOM without counting it in buffer_position(),
        // so strip it here to keep positions aligned with `data` for the
        // verbatim spans.
        let data = data.strip_prefix(b"\xEF\xBB\xBF").unwrap_or(data);
        let mut reader = Reader::from_reader(data);
        reader.config_mut().trim_text(true);

        let mut buf = Vec::with_capacity(1024);
        let mut skip_buf = Vec::new();
        let mut p = StylesParser::new(*budget);
        let mut root_ns = xml::RootNamespaces::default();
        let mut spans: VerbatimSpans<'_> = [None; 3];
        let mut depth = 0usize;
        let parse_err = |err: quick_xml::Error| {
            cold_path();
            ModernXlsxError::XmlParse(format!("error parsing styles.xml: {err}"))
        };

        loop {
            let start = reader.buffer_position();
            match reader.read_event_into(&mut buf).map_err(parse_err)? {
                Event::Start(e) => {
                    let name = e.local_name();
                    let local = name.as_ref();
                    if depth == 0 {
                        // The spans so far belong to the root being replaced.
                        flush_verbatim(&mut p.out, &mut spans, &root_ns);
                        root_ns = xml::RootNamespaces::new(xml::root_namespaces(&e)?);
                    } else if depth == 1
                        && let Some(slot) = verbatim_slot(local)
                    {
                        reader
                            .read_to_end_into(e.name(), &mut skip_buf)
                            .map_err(parse_err)?;
                        spans[slot] = Some(verbatim(data, start, reader.buffer_position())?);
                        buf.clear();
                        continue;
                    } else if local == "extLst" {
                        // Extensions below the root are not modelled; skip them
                        // whole so their children are not read as styles.
                        reader
                            .read_to_end_into(e.name(), &mut skip_buf)
                            .map_err(parse_err)?;
                        buf.clear();
                        continue;
                    }
                    p.open(&e)?;
                    depth += 1;
                }
                Event::Empty(e) => {
                    let name = e.local_name();
                    let local = name.as_ref();
                    if depth == 1
                        && let Some(slot) = verbatim_slot(local)
                    {
                        spans[slot] = Some(verbatim(data, start, reader.buffer_position())?);
                    } else {
                        p.open(&e)?;
                        p.close(local)?;
                    }
                }
                Event::End(e) => {
                    depth = depth.saturating_sub(1);
                    p.close(e.local_name().as_ref())?;
                }
                Event::Eof => break,
                _ => {}
            }
            buf.clear();
        }

        flush_verbatim(&mut p.out, &mut spans, &root_ns);
        *budget = p.budget;
        Ok(p.out)
    }

    // -----------------------------------------------------------------------
    // Writer
    // -----------------------------------------------------------------------

    /// Serialize this `Styles` to `styles.xml` bytes in `CT_Stylesheet`
    /// order: numFmts, fonts, fills, borders, cellStyleXfs, cellXfs,
    /// cellStyles, dxfs, tableStyles, colors, extLst.
    ///
    /// # Errors
    ///
    /// `InvalidWorkbook` when a colour `rgb` is not 6 or 8 hex digits (see
    /// `color::checked_argb`); `XmlWrite` for a non-finite font size, gradient
    /// degree, stop position or colour tint.
    pub fn to_xml(&self) -> Result<Vec<u8>> {
        let mut buf: Vec<u8> = Vec::with_capacity(4096);
        let mut writer = Writer::new(&mut buf);
        let mut ibuf = itoa::Buffer::new();
        let mut xbuf = XsdBuf::new();

        write_xml_decl(&mut writer)?;

        // <styleSheet xmlns="...">
        let mut root = BytesStart::new("styleSheet");
        root.push_attribute(("xmlns", SPREADSHEET_NS));
        emit(&mut writer, Event::Start(root))?;

        // --- numFmts ---
        if !self.num_fmts.is_empty() {
            let mut nf_start = BytesStart::new("numFmts");
            nf_start.push_attribute(("count", ibuf.format(self.num_fmts.len())));
            emit(&mut writer, Event::Start(nf_start))?;
            for nf in &self.num_fmts {
                write_num_fmt(&mut writer, nf, &mut ibuf)?;
            }
            emit(&mut writer, Event::End(BytesEnd::new("numFmts")))?;
        }

        // --- fonts ---
        let mut f_start = BytesStart::new("fonts");
        f_start.push_attribute(("count", ibuf.format(self.fonts.len())));
        emit(&mut writer, Event::Start(f_start))?;
        for font in &self.fonts {
            write_font(&mut writer, font, &mut ibuf, &mut xbuf)?;
        }
        emit(&mut writer, Event::End(BytesEnd::new("fonts")))?;

        // --- fills ---
        let mut f_start = BytesStart::new("fills");
        f_start.push_attribute(("count", ibuf.format(self.fills.len())));
        emit(&mut writer, Event::Start(f_start))?;
        for fill in &self.fills {
            write_fill(&mut writer, fill, &mut xbuf)?;
        }
        emit(&mut writer, Event::End(BytesEnd::new("fills")))?;

        // --- borders ---
        let mut b_start = BytesStart::new("borders");
        b_start.push_attribute(("count", ibuf.format(self.borders.len())));
        emit(&mut writer, Event::Start(b_start))?;
        for border in &self.borders {
            write_border(&mut writer, border, &mut xbuf)?;
        }
        emit(&mut writer, Event::End(BytesEnd::new("borders")))?;

        // --- cellStyleXfs: always present; an empty list is written as one default record (index 0) ---
        let default_style_xf = [CellXf::default()];
        let style_xfs = if self.cell_style_xfs.is_empty() {
            &default_style_xf[..]
        } else {
            &self.cell_style_xfs
        };
        write_xfs(&mut writer, "cellStyleXfs", style_xfs, &mut ibuf)?;

        // --- cellXfs ---
        write_xfs(&mut writer, "cellXfs", &self.cell_xfs, &mut ibuf)?;

        // --- cellStyles ---
        if !self.cell_styles.is_empty() {
            let mut cs_start = BytesStart::new("cellStyles");
            cs_start.push_attribute(("count", ibuf.format(self.cell_styles.len())));
            emit(&mut writer, Event::Start(cs_start))?;
            for cs in &self.cell_styles {
                let mut elem = BytesStart::new("cellStyle");
                elem.push_attribute(xml_attr("name", &cs.name));
                elem.push_attribute(("xfId", ibuf.format(cs.xf_id)));
                if let Some(bid) = cs.builtin_id {
                    elem.push_attribute(("builtinId", ibuf.format(bid)));
                }
                emit(&mut writer, Event::Empty(elem))?;
            }
            emit(&mut writer, Event::End(BytesEnd::new("cellStyles")))?;
        }

        // --- dxfs ---
        if !self.dxfs.is_empty() {
            let mut d_start = BytesStart::new("dxfs");
            d_start.push_attribute(("count", ibuf.format(self.dxfs.len())));
            emit(&mut writer, Event::Start(d_start))?;
            for dxf in &self.dxfs {
                write_dxf(&mut writer, dxf, &mut ibuf, &mut xbuf)?;
            }
            emit(&mut writer, Event::End(BytesEnd::new("dxfs")))?;
        }

        // --- tableStyles, colors, extLst: verbatim from the source ---
        for raw in [&self.table_styles_xml, &self.colors_xml, &self.ext_lst_xml]
            .into_iter()
            .flatten()
        {
            writer.get_mut().extend_from_slice(raw.as_bytes());
        }

        // </styleSheet>
        emit(&mut writer, Event::End(BytesEnd::new("styleSheet")))?;

        Ok(buf)
    }
}

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

/// The numeric `val` of a property element.
fn val_number<T: std::str::FromStr>(elem: &str, e: &BytesStart<'_>) -> Result<Option<T>> {
    val_attr(e)?.map(|a| parse_attr(elem, &a)).transpose()
}

/// Apply one `CT_Font` child element (other than `<color>`) to `font`.
/// `CT_BooleanProperty` elements (`b`, `i`, `strike`, `condense`, `extend`)
/// are on unless `val` is an xsd:boolean false. In a dxf font (`in_dxf`) an
/// explicit off, and `<u val="none"/>`, is listed in [`Font::explicit_off`]:
/// there it overrides the cell's value, while in a `<fonts>` entry it equals
/// an absent element. Each tag listed is counted in `budget` first (a
/// `String` and its bytes), since one font can repeat the tag without end.
fn font_property(
    font: &mut Font,
    local: &str,
    e: &BytesStart<'_>,
    in_dxf: bool,
    budget: &mut ExpansionBudget,
) -> Result<()> {
    match local {
        "b" | "i" | "strike" | "condense" | "extend" => {
            let on = ct_boolean(e)?;
            match local {
                "b" => font.bold = on,
                "i" => font.italic = on,
                "strike" => font.strike = on,
                "condense" => font.condense = on,
                _ => font.extend = on,
            }
            if !on && in_dxf {
                budget.store(size_of::<String>() + local.len())?;
                font.explicit_off.push(local.to_owned());
            }
        }
        "u" => {
            // CT_UnderlineProperty: `val` defaults to "single".
            let val = val_attr(e)?.map(|a| attr_str(&a)).transpose()?;
            let style = val.as_deref().map_or("single", str::trim_ascii);
            font.underline = style != "none";
            font.underline_style =
                matches!(style, "double" | "singleAccounting" | "doubleAccounting")
                    .then(|| style.to_owned());
            if style == "none" && in_dxf {
                budget.store(size_of::<String>() + 1)?;
                font.explicit_off.push("u".to_owned());
            }
        }
        "sz" => font.size = val_number("sz", e)?,
        "name" => font.name = val_string(e)?,
        "vertAlign" => font.vert_align = val_string(e)?,
        "family" => font.family = val_number("family", e)?,
        "charset" => font.charset = val_number("charset", e)?,
        "scheme" => font.scheme = val_string(e)?,
        _ => {}
    }
    Ok(())
}

/// Parse a `<numFmt>` element.
fn parse_num_fmt(e: &BytesStart<'_>) -> Result<NumFmt> {
    let mut nf = NumFmt {
        id: 0,
        format_code: String::new(),
    };
    for attr in attrs(e) {
        let attr = attr?;
        match attr.key.local_name().as_ref() {
            "numFmtId" => nf.id = parse_attr("numFmt", &attr)?,
            "formatCode" => nf.format_code = attr_str(&attr)?.into_owned(),
            _ => {}
        }
    }
    Ok(nf)
}

/// Parse the attributes of a `<gradientFill>` element.
fn parse_gradient_attrs(e: &BytesStart<'_>) -> Result<GradientFill> {
    let mut gradient = GradientFill::default();
    for attr in attrs(e) {
        let attr = attr?;
        match attr.key.local_name().as_ref() {
            "type" => gradient.gradient_type = Some(attr_str(&attr)?.into_owned()),
            "degree" => gradient.degree = Some(parse_attr("gradientFill", &attr)?),
            _ => {}
        }
    }
    Ok(gradient)
}

/// Parse the attributes of a `<border>` element.
fn parse_border_attrs(e: &BytesStart<'_>) -> Result<Border> {
    let mut border = Border::default();
    for attr in attrs(e) {
        let attr = attr?;
        match attr.key.local_name().as_ref() {
            "diagonalUp" => border.diagonal_up = xsd_bool(&attr_str(&attr)?).unwrap_or(false),
            "diagonalDown" => {
                border.diagonal_down = xsd_bool(&attr_str(&attr)?).unwrap_or(false);
            }
            _ => {}
        }
    }
    Ok(border)
}

/// Extract a `BorderSide` from attributes on a border child element (e.g. `<left style="thin">`).
/// Returns `None` if there is no `style` attribute.
fn parse_border_side_attrs(e: &BytesStart<'_>) -> Result<Option<BorderSide>> {
    for attr in attrs(e) {
        let attr = attr?;
        if attr.key.local_name().as_ref() == "style" {
            return Ok(Some(BorderSide {
                style: attr_str(&attr)?.into_owned(),
                color: None,
                color_spec: None,
            }));
        }
    }
    Ok(None)
}

/// Parse attributes of an `<xf>` element into a `CellXf`.
fn parse_xf_attrs(e: &BytesStart<'_>) -> Result<CellXf> {
    let mut xf = CellXf::default();
    for attr in attrs(e) {
        let attr = attr?;
        let flag = || -> Result<bool> { Ok(xsd_bool(&attr_str(&attr)?).unwrap_or(false)) };
        match attr.key.local_name().as_ref() {
            "numFmtId" => xf.num_fmt_id = parse_attr("xf", &attr)?,
            "fontId" => xf.font_id = parse_attr("xf", &attr)?,
            "fillId" => xf.fill_id = parse_attr("xf", &attr)?,
            "borderId" => xf.border_id = parse_attr("xf", &attr)?,
            "xfId" => xf.xf_id = Some(parse_attr("xf", &attr)?),
            "applyFont" => xf.apply_font = flag()?,
            "applyFill" => xf.apply_fill = flag()?,
            "applyBorder" => xf.apply_border = flag()?,
            "applyNumberFormat" => xf.apply_number_format = flag()?,
            "applyAlignment" => xf.apply_alignment = flag()?,
            "applyProtection" => xf.apply_protection = flag()?,
            "quotePrefix" => xf.quote_prefix = flag()?,
            "pivotButton" => xf.pivot_button = flag()?,
            _ => {}
        }
    }
    Ok(xf)
}

/// Parse attributes of an `<alignment>` element.
fn parse_alignment_attrs(e: &BytesStart<'_>) -> Result<Alignment> {
    let mut alignment = Alignment::default();
    for attr in attrs(e) {
        let attr = attr?;
        let flag = || -> Result<bool> { Ok(xsd_bool(&attr_str(&attr)?).unwrap_or(false)) };
        match attr.key.local_name().as_ref() {
            "horizontal" => alignment.horizontal = Some(attr_str(&attr)?.into_owned()),
            "vertical" => alignment.vertical = Some(attr_str(&attr)?.into_owned()),
            "wrapText" => alignment.wrap_text = flag()?,
            "textRotation" => alignment.text_rotation = Some(parse_attr("alignment", &attr)?),
            "indent" => alignment.indent = Some(parse_attr("alignment", &attr)?),
            "shrinkToFit" => alignment.shrink_to_fit = flag()?,
            "readingOrder" => alignment.reading_order = Some(parse_attr("alignment", &attr)?),
            "justifyLastLine" => alignment.justify_last_line = flag()?,
            "relativeIndent" => {
                alignment.relative_indent = Some(parse_attr("alignment", &attr)?);
            }
            _ => {}
        }
    }
    Ok(alignment)
}

/// Parse attributes of a `<protection>` element (`locked` defaults to true).
fn parse_protection_attrs(e: &BytesStart<'_>) -> Result<Protection> {
    let mut protection = Protection::default();
    for attr in attrs(e) {
        let attr = attr?;
        match attr.key.local_name().as_ref() {
            "locked" => protection.locked = xsd_bool(&attr_str(&attr)?).unwrap_or(true),
            "hidden" => protection.hidden = xsd_bool(&attr_str(&attr)?).unwrap_or(false),
            _ => {}
        }
    }
    Ok(protection)
}

/// Parse attributes of a `<cellStyle>` element.
fn parse_cell_style_attrs(e: &BytesStart<'_>) -> Result<CellStyle> {
    let mut cs = CellStyle::default();
    for attr in attrs(e) {
        let attr = attr?;
        match attr.key.local_name().as_ref() {
            "name" => cs.name = attr_str(&attr)?.into_owned(),
            "xfId" => cs.xf_id = parse_attr("cellStyle", &attr)?,
            "builtinId" => cs.builtin_id = Some(parse_attr("cellStyle", &attr)?),
            _ => {}
        }
    }
    Ok(cs)
}

/// Write one event, mapping the writer's I/O error to `XmlWrite`.
#[inline]
fn emit<W: std::io::Write>(writer: &mut Writer<W>, event: Event<'_>) -> Result<()> {
    writer.write_event(event).map_err(xml_write_err)
}

/// Write a `<numFmt>` element.
fn write_num_fmt<W: std::io::Write>(
    writer: &mut Writer<W>,
    nf: &NumFmt,
    ibuf: &mut itoa::Buffer,
) -> Result<()> {
    let mut elem = BytesStart::new("numFmt");
    elem.push_attribute(("numFmtId", ibuf.format(nf.id)));
    elem.push_attribute(xml_attr("formatCode", &nf.format_code));
    emit(writer, Event::Empty(elem))
}

/// Write `<tag count="n">` with one `<xf>` per record (`cellStyleXfs` or `cellXfs`).
fn write_xfs<W: std::io::Write>(
    writer: &mut Writer<W>,
    tag: &str,
    xfs: &[CellXf],
    ibuf: &mut itoa::Buffer,
) -> Result<()> {
    let mut start = BytesStart::new(tag);
    start.push_attribute(("count", ibuf.format(xfs.len())));
    emit(writer, Event::Start(start))?;

    for xf in xfs {
        let mut elem = BytesStart::new("xf");
        elem.push_attribute(("numFmtId", ibuf.format(xf.num_fmt_id)));
        elem.push_attribute(("fontId", ibuf.format(xf.font_id)));
        elem.push_attribute(("fillId", ibuf.format(xf.fill_id)));
        elem.push_attribute(("borderId", ibuf.format(xf.border_id)));
        if let Some(xf_id) = xf.xf_id {
            elem.push_attribute(("xfId", ibuf.format(xf_id)));
        }
        for (set, name) in [
            (xf.apply_font, "applyFont"),
            (xf.apply_fill, "applyFill"),
            (xf.apply_border, "applyBorder"),
            (xf.apply_number_format, "applyNumberFormat"),
            (xf.apply_alignment, "applyAlignment"),
            (xf.apply_protection, "applyProtection"),
            (xf.quote_prefix, "quotePrefix"),
            (xf.pivot_button, "pivotButton"),
        ] {
            if set {
                elem.push_attribute((name, "1"));
            }
        }

        if xf.alignment.is_some() || xf.protection.is_some() {
            emit(writer, Event::Start(elem))?;
            if let Some(ref alignment) = xf.alignment {
                write_alignment(writer, alignment, ibuf)?;
            }
            if let Some(ref protection) = xf.protection {
                write_protection(writer, protection)?;
            }
            emit(writer, Event::End(BytesEnd::new("xf")))?;
        } else {
            emit(writer, Event::Empty(elem))?;
        }
    }

    emit(writer, Event::End(BytesEnd::new(tag)))
}

/// Write a `<dxf>` element in `CT_Dxf` order: font, numFmt, fill, alignment,
/// border, protection.
fn write_dxf<W: std::io::Write>(
    writer: &mut Writer<W>,
    dxf: &DxfStyle,
    ibuf: &mut itoa::Buffer,
    xbuf: &mut XsdBuf,
) -> Result<()> {
    emit(writer, Event::Start(BytesStart::new("dxf")))?;
    if let Some(ref font) = dxf.font {
        write_font(writer, font, ibuf, xbuf)?;
    }
    if let Some(ref nf) = dxf.num_fmt {
        write_num_fmt(writer, nf, ibuf)?;
    }
    if let Some(ref fill) = dxf.fill {
        write_fill(writer, fill, xbuf)?;
    }
    if let Some(ref alignment) = dxf.alignment {
        write_alignment(writer, alignment, ibuf)?;
    }
    if let Some(ref border) = dxf.border {
        write_border(writer, border, xbuf)?;
    }
    if let Some(ref protection) = dxf.protection {
        write_protection(writer, protection)?;
    }
    emit(writer, Event::End(BytesEnd::new("dxf")))
}

/// Write a single font element.
fn write_font<W: std::io::Write>(
    writer: &mut Writer<W>,
    font: &Font,
    ibuf: &mut itoa::Buffer,
    xbuf: &mut XsdBuf,
) -> Result<()> {
    emit(writer, Event::Start(BytesStart::new("font")))?;

    let listed_off = |tag: &str| font.explicit_off.iter().any(|t| t == tag);
    for (on, tag) in [(font.bold, "b"), (font.italic, "i")] {
        write_bool_property(writer, tag, on, listed_off(tag))?;
    }
    if let Some(ref style) = font.underline_style {
        let mut elem = BytesStart::new("u");
        elem.push_attribute(xml_attr("val", style));
        emit(writer, Event::Empty(elem))?;
    } else if font.underline {
        emit(writer, Event::Empty(BytesStart::new("u")))?;
    } else if listed_off("u") {
        let mut elem = BytesStart::new("u");
        elem.push_attribute(("val", "none"));
        emit(writer, Event::Empty(elem))?;
    }
    for (on, tag) in [
        (font.strike, "strike"),
        (font.condense, "condense"),
        (font.extend, "extend"),
    ] {
        write_bool_property(writer, tag, on, listed_off(tag))?;
    }
    if let Some(sz) = font.size {
        let mut elem = BytesStart::new("sz");
        elem.push_attribute(("val", fmt_xsd_double(sz, xbuf, "font sz")?));
        emit(writer, Event::Empty(elem))?;
    }
    if let Some(ref name) = font.name {
        let mut elem = BytesStart::new("name");
        elem.push_attribute(xml_attr("val", name));
        emit(writer, Event::Empty(elem))?;
    }
    if let Some(color) = resolve_color(font.color.as_deref(), font.color_spec.as_ref()) {
        write_color(writer, "color", color, xbuf)?;
    }
    if let Some(ref vert_align) = font.vert_align {
        let mut elem = BytesStart::new("vertAlign");
        elem.push_attribute(xml_attr("val", vert_align));
        emit(writer, Event::Empty(elem))?;
    }
    if let Some(family) = font.family {
        let mut elem = BytesStart::new("family");
        elem.push_attribute(("val", ibuf.format(family)));
        emit(writer, Event::Empty(elem))?;
    }
    if let Some(charset) = font.charset {
        let mut elem = BytesStart::new("charset");
        elem.push_attribute(("val", ibuf.format(charset)));
        emit(writer, Event::Empty(elem))?;
    }
    if let Some(ref scheme) = font.scheme {
        let mut elem = BytesStart::new("scheme");
        elem.push_attribute(xml_attr("val", scheme));
        emit(writer, Event::Empty(elem))?;
    }

    emit(writer, Event::End(BytesEnd::new("font")))
}

/// Write a `CT_BooleanProperty` font child: `<tag/>` when on, `<tag val="0"/>`
/// when off and listed in [`Font::explicit_off`], nothing otherwise.
fn write_bool_property<W: std::io::Write>(
    writer: &mut Writer<W>,
    tag: &str,
    on: bool,
    explicit_off: bool,
) -> Result<()> {
    if on {
        emit(writer, Event::Empty(BytesStart::new(tag)))
    } else if explicit_off {
        let mut elem = BytesStart::new(tag);
        elem.push_attribute(("val", "0"));
        emit(writer, Event::Empty(elem))
    } else {
        Ok(())
    }
}

/// Write a single fill element. An empty `pattern_type` writes a
/// `<patternFill>` without `patternType` (its schema default, `none`).
fn write_fill<W: std::io::Write>(
    writer: &mut Writer<W>,
    fill: &Fill,
    xbuf: &mut XsdBuf,
) -> Result<()> {
    emit(writer, Event::Start(BytesStart::new("fill")))?;

    if let Some(ref gradient) = fill.gradient_fill {
        let mut gf = BytesStart::new("gradientFill");
        if let Some(ref gt) = gradient.gradient_type {
            gf.push_attribute(xml_attr("type", gt));
        }
        if let Some(degree) = gradient.degree {
            gf.push_attribute((
                "degree",
                fmt_xsd_double(degree, xbuf, "gradientFill degree")?,
            ));
        }

        if gradient.stops.is_empty() {
            emit(writer, Event::Empty(gf))?;
        } else {
            emit(writer, Event::Start(gf))?;
            for stop in &gradient.stops {
                let mut stop_elem = BytesStart::new("stop");
                stop_elem.push_attribute((
                    "position",
                    fmt_xsd_double(stop.position, xbuf, "gradient stop position")?,
                ));
                emit(writer, Event::Start(stop_elem))?;
                // CT_GradientStop requires its colour, so a stop without one
                // still gets an (all-default) `<color/>`.
                match resolve_color(Some(stop.color.as_str()), stop.color_spec.as_ref()) {
                    Some(color) => write_color(writer, "color", color, xbuf)?,
                    None => emit(writer, Event::Empty(BytesStart::new("color")))?,
                }
                emit(writer, Event::End(BytesEnd::new("stop")))?;
            }
            emit(writer, Event::End(BytesEnd::new("gradientFill")))?;
        }
    } else {
        let mut pf = BytesStart::new("patternFill");
        if !fill.pattern_type.is_empty() {
            pf.push_attribute(xml_attr("patternType", &fill.pattern_type));
        }
        let fg = resolve_color(fill.fg_color.as_deref(), fill.fg_color_spec.as_ref());
        let bg = resolve_color(fill.bg_color.as_deref(), fill.bg_color_spec.as_ref());

        if fg.is_some() || bg.is_some() {
            emit(writer, Event::Start(pf))?;
            if let Some(fg) = fg {
                write_color(writer, "fgColor", fg, xbuf)?;
            }
            if let Some(bg) = bg {
                write_color(writer, "bgColor", bg, xbuf)?;
            }
            emit(writer, Event::End(BytesEnd::new("patternFill")))?;
        } else {
            emit(writer, Event::Empty(pf))?;
        }
    }

    emit(writer, Event::End(BytesEnd::new("fill")))
}

/// Write a single border element.
fn write_border<W: std::io::Write>(
    writer: &mut Writer<W>,
    border: &Border,
    xbuf: &mut XsdBuf,
) -> Result<()> {
    let mut border_elem = BytesStart::new("border");
    if border.diagonal_up {
        border_elem.push_attribute(("diagonalUp", "1"));
    }
    if border.diagonal_down {
        border_elem.push_attribute(("diagonalDown", "1"));
    }
    emit(writer, Event::Start(border_elem))?;

    write_border_side(writer, "left", border.left.as_ref(), xbuf)?;
    write_border_side(writer, "right", border.right.as_ref(), xbuf)?;
    write_border_side(writer, "top", border.top.as_ref(), xbuf)?;
    write_border_side(writer, "bottom", border.bottom.as_ref(), xbuf)?;
    write_border_side(writer, "diagonal", border.diagonal.as_ref(), xbuf)?;

    emit(writer, Event::End(BytesEnd::new("border")))
}

/// Write a single border side element (`<left/>`, `<left style="thin"><color rgb="..."/></left>`, etc.).
fn write_border_side<W: std::io::Write>(
    writer: &mut Writer<W>,
    tag: &str,
    side: Option<&BorderSide>,
    xbuf: &mut XsdBuf,
) -> Result<()> {
    let Some(bs) = side else {
        return emit(writer, Event::Empty(BytesStart::new(tag)));
    };
    let mut elem = BytesStart::new(tag);
    elem.push_attribute(xml_attr("style", &bs.style));

    if let Some(color) = resolve_color(bs.color.as_deref(), bs.color_spec.as_ref()) {
        emit(writer, Event::Start(elem))?;
        write_color(writer, "color", color, xbuf)?;
        emit(writer, Event::End(BytesEnd::new(tag)))
    } else {
        emit(writer, Event::Empty(elem))
    }
}

/// Write an `<alignment>` element.
fn write_alignment<W: std::io::Write>(
    writer: &mut Writer<W>,
    alignment: &Alignment,
    ibuf: &mut itoa::Buffer,
) -> Result<()> {
    let mut elem = BytesStart::new("alignment");
    if let Some(ref h) = alignment.horizontal {
        elem.push_attribute(xml_attr("horizontal", h));
    }
    if let Some(ref v) = alignment.vertical {
        elem.push_attribute(xml_attr("vertical", v));
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
    if let Some(relative_indent) = alignment.relative_indent {
        elem.push_attribute(("relativeIndent", ibuf.format(relative_indent)));
    }
    if alignment.justify_last_line {
        elem.push_attribute(("justifyLastLine", "1"));
    }
    if alignment.shrink_to_fit {
        elem.push_attribute(("shrinkToFit", "1"));
    }
    if let Some(reading_order) = alignment.reading_order {
        elem.push_attribute(("readingOrder", ibuf.format(reading_order)));
    }
    emit(writer, Event::Empty(elem))
}

/// Write a `<protection>` element.
fn write_protection<W: std::io::Write>(
    writer: &mut Writer<W>,
    protection: &Protection,
) -> Result<()> {
    let mut elem = BytesStart::new("protection");
    // Only write locked if it differs from default (true).
    if !protection.locked {
        elem.push_attribute(("locked", "0"));
    }
    if protection.hidden {
        elem.push_attribute(("hidden", "1"));
    }
    emit(writer, Event::Empty(elem))
}

// ---------------------------------------------------------------------------
// Tests
// ---------------------------------------------------------------------------

#[cfg(test)]
mod tests;
