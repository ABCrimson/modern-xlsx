//! `xl/sharedStrings.xml`: the workbook's shared string table.
//!
//! [`SharedStringTable`] is the table parsed from a file (plain text plus rich
//! text runs) and [`SharedStringTableBuilder`] is the deduplicating table the
//! writer fills with [`SharedStringTableBuilder::insert`] and
//! [`SharedStringTableBuilder::insert_rich`] and serializes with
//! [`SharedStringTableBuilder::to_xml`].

use core::hint::cold_path;
use std::borrow::Cow;
use std::collections::HashMap;
use std::hash::{BuildHasher, Hash, Hasher, RandomState};
use std::io::Write;
use std::rc::Rc;

use quick_xml::events::{BytesEnd, BytesStart, BytesText, Event};
use quick_xml::{Reader, Writer};

use serde::{Deserialize, Serialize};

use super::color::{Color, resolve_color, write_color};
use super::xml::{
    SPREADSHEET_NS, XsdBuf, ct_boolean, fmt_xsd_double, needs_preserve, parse_attr,
    push_general_ref, val_attr, val_string, xml_attr,
};

use crate::ooxml::{write_xml_decl, xml_write_err};
use crate::{ModernXlsxError, Result};

/// A single run of formatted text within a rich text shared string.
///
/// Each run carries its text content plus the run properties found in the
/// `<rPr>` element (`CT_RPrElt`, ECMA-376 Part 1, §18.4.7).
///
/// `color` stays an rgb-only string; a theme, indexed, tinted or automatic
/// colour is carried in `color_spec`. On write `color_spec` wins only while its
/// `rgb` equals `color` (an empty string counts as absent); otherwise `color`
/// is written as `rgb`.
#[derive(Debug, Clone, Default, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct RichTextRun {
    pub text: String,
    /// `<b>`: `Some(false)` is an explicit `<b val="0"/>`.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub bold: Option<bool>,
    /// `<i>`: `Some(false)` is an explicit `<i val="0"/>`.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub italic: Option<bool>,
    /// `<u>`: `Some(true)` for every underline style, `Some(false)` for
    /// `<u val="none"/>`. A style other than `single` and `none` is in
    /// `underline_style`. On write, `Some(false)` wins over any style.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub underline: Option<bool>,
    /// `<strike>`: `Some(false)` is an explicit `<strike val="0"/>`.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub strike: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub font_name: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub font_size: Option<f64>,
    /// `<color rgb>` only; see `color_spec` for the other colour kinds.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub color: Option<String>,
    /// `<vertAlign val>`: `baseline`, `superscript` or `subscript`.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub vert_align: Option<String>,
    /// `<family val>`: the font family class.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub family: Option<u32>,
    /// `<charset val>`: the font character set.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub charset: Option<u32>,
    /// `<scheme val>`: `major`, `minor` or `none`.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub scheme: Option<String>,
    /// The full `<color>` element when it is more than a plain rgb value.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub color_spec: Option<Color>,
    /// The `<u val>` token when it is not absent, `single` or `none`
    /// (`double`, `singleAccounting` or `doubleAccounting`).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub underline_style: Option<String>,
}

/// The numeric `val` of an `<rPr>` child, parsed as `T`.
fn val_num<T: std::str::FromStr>(e: &BytesStart<'_>) -> Result<Option<T>> {
    let name = e.name();
    val_attr(e)?
        .map(|a| parse_attr(name.as_ref(), &a))
        .transpose()
}

/// Apply one `<rPr>` child element (`CT_RPrElt`) to `run`. Children this model
/// does not carry are ignored.
fn apply_rpr_child(run: &mut RichTextRun, e: &BytesStart<'_>) -> Result<()> {
    match e.local_name().as_ref() {
        "b" => run.bold = Some(ct_boolean(e)?),
        "i" => run.italic = Some(ct_boolean(e)?),
        "strike" => run.strike = Some(ct_boolean(e)?),
        "u" => {
            let val = val_string(e)?;
            (run.underline, run.underline_style) = match val.as_deref() {
                None | Some("single") => (Some(true), None),
                Some("none") => (Some(false), None),
                Some(other) => (Some(true), Some(other.to_owned())),
            };
        }
        "rFont" => run.font_name = val_string(e)?,
        "sz" => run.font_size = val_num(e)?,
        "color" => {
            if let Some(c) = Color::from_attrs(e)? {
                run.color.clone_from(&c.rgb);
                run.color_spec = c.has_non_rgb().then_some(c);
            }
        }
        "vertAlign" => run.vert_align = val_string(e)?,
        "family" => run.family = val_num(e)?,
        "charset" => run.charset = val_num(e)?,
        "scheme" => run.scheme = val_string(e)?,
        _ => {}
    }
    Ok(())
}

/// Write `<name/>`, or `<name val="0"/>` for an explicit `Some(false)`.
fn write_flag<W: Write>(writer: &mut Writer<W>, name: &str, flag: Option<bool>) -> Result<()> {
    let Some(on) = flag else { return Ok(()) };
    let mut e = BytesStart::new(name);
    if !on {
        e.push_attribute(("val", "0"));
    }
    writer.write_event(Event::Empty(e)).map_err(xml_write_err)
}

/// Write `<name val="value"/>`.
fn write_val<W: Write>(writer: &mut Writer<W>, name: &str, value: &str) -> Result<()> {
    let mut e = BytesStart::new(name);
    e.push_attribute(xml_attr("val", value));
    writer.write_event(Event::Empty(e)).map_err(xml_write_err)
}

/// Write the `<rPr>` of `run`, or nothing when the run has no properties.
fn write_rpr<W: Write>(writer: &mut Writer<W>, xsd: &mut XsdBuf, run: &RichTextRun) -> Result<()> {
    // `color_spec` while its `rgb` agrees with `color`, else `color` as plain
    // rgb; an all-default `color_spec` writes no `<color>`.
    let color = resolve_color(run.color.as_deref(), run.color_spec.as_ref())
        .filter(|c| !c.is_default_spec());
    if run.bold.is_none()
        && run.italic.is_none()
        && run.underline.is_none()
        && run.underline_style.is_none()
        && run.strike.is_none()
        && run.font_name.is_none()
        && run.font_size.is_none()
        && color.is_none()
        && run.vert_align.is_none()
        && run.family.is_none()
        && run.charset.is_none()
        && run.scheme.is_none()
    {
        return Ok(());
    }
    writer
        .write_event(Event::Start(BytesStart::new("rPr")))
        .map_err(xml_write_err)?;

    write_flag(writer, "b", run.bold)?;
    write_flag(writer, "i", run.italic)?;
    // The edited boolean wins: `Some(false)` is `<u val="none"/>`; otherwise a
    // style token other than `none` is written, else `Some(true)` is `<u/>`.
    match (run.underline, run.underline_style.as_deref()) {
        (Some(false), _) => write_val(writer, "u", "none")?,
        (_, Some(style)) if style != "none" => write_val(writer, "u", style)?,
        (Some(true), _) => writer
            .write_event(Event::Empty(BytesStart::new("u")))
            .map_err(xml_write_err)?,
        _ => {}
    }
    write_flag(writer, "strike", run.strike)?;
    if let Some(size) = run.font_size {
        let mut sz = BytesStart::new("sz");
        sz.push_attribute(("val", fmt_xsd_double(size, xsd, "rPr sz")?));
        writer
            .write_event(Event::Empty(sz))
            .map_err(xml_write_err)?;
    }
    if let Some(color) = color {
        write_color(writer, "color", color, xsd)?;
    }
    if let Some(name) = run.font_name.as_deref() {
        write_val(writer, "rFont", name)?;
    }
    if let Some(family) = run.family {
        write_val(writer, "family", itoa::Buffer::new().format(family))?;
    }
    if let Some(charset) = run.charset {
        write_val(writer, "charset", itoa::Buffer::new().format(charset))?;
    }
    if let Some(scheme) = run.scheme.as_deref() {
        write_val(writer, "scheme", scheme)?;
    }
    if let Some(align) = run.vert_align.as_deref() {
        write_val(writer, "vertAlign", align)?;
    }

    writer
        .write_event(Event::End(BytesEnd::new("rPr")))
        .map_err(xml_write_err)
}

/// Write `<t>text</t>`, dropping characters that are illegal in XML 1.0 and
/// preserving the edge white space of what is left.
fn write_t<W: Write>(writer: &mut Writer<W>, text: &str) -> Result<()> {
    let text = crate::ooxml::sanitize_xml_text(text);
    let mut t_elem = BytesStart::new("t");
    if needs_preserve(&text) {
        t_elem.push_attribute(("xml:space", "preserve"));
    }
    writer
        .write_event(Event::Start(t_elem))
        .map_err(xml_write_err)?;
    writer
        .write_event(Event::Text(BytesText::new(&text)))
        .map_err(xml_write_err)?;
    writer
        .write_event(Event::End(BytesEnd::new("t")))
        .map_err(xml_write_err)
}

/// Write one `<si>`: the runs when the entry is rich text, else `plain` as a
/// single `<t>`.
fn write_si<W: Write>(
    writer: &mut Writer<W>,
    xsd: &mut XsdBuf,
    plain: &str,
    runs: Option<&[RichTextRun]>,
) -> Result<()> {
    writer
        .write_event(Event::Start(BytesStart::new("si")))
        .map_err(xml_write_err)?;
    if let Some(runs) = runs {
        for run in runs {
            writer
                .write_event(Event::Start(BytesStart::new("r")))
                .map_err(xml_write_err)?;
            write_rpr(writer, xsd, run)?;
            write_t(writer, &run.text)?;
            writer
                .write_event(Event::End(BytesEnd::new("r")))
                .map_err(xml_write_err)?;
        }
    } else {
        write_t(writer, plain)?;
    }
    writer
        .write_event(Event::End(BytesEnd::new("si")))
        .map_err(xml_write_err)
}

/// Write `<sst>` open tag; `count` is omitted when `None` (it is optional in
/// `CT_Sst`).
fn write_sst_start<W: Write>(
    writer: &mut Writer<W>,
    count: Option<usize>,
    unique: usize,
) -> Result<()> {
    let mut ibuf = itoa::Buffer::new();
    let mut sst_start = BytesStart::new("sst");
    sst_start.push_attribute(("xmlns", SPREADSHEET_NS));
    if let Some(count) = count {
        sst_start.push_attribute(("count", ibuf.format(count)));
    }
    let mut ibuf2 = itoa::Buffer::new();
    sst_start.push_attribute(("uniqueCount", ibuf2.format(unique)));
    writer
        .write_event(Event::Start(sst_start))
        .map_err(xml_write_err)
}

/// One read's running count of the text it copies out of a table by index: a
/// shared string (and a rich entry's runs) into every cell naming it, a
/// comment author into every comment naming it, a hyperlink target into
/// every hyperlink naming its relationship, a `<worksheet>` namespace
/// declaration onto every kept `<ext>` using its prefix. The ZIP limits bound the parts,
/// not these copies, so the count is held against
/// `ZipSecurityLimits::max_json_length`. The JSON a read returns holds every
/// copy counted here at no less than its byte length, so a read whose copies
/// pass that limit could not have returned JSON within it.
///
/// Beside the copies it counts, in `stored`, the memory of the elements the
/// reader keeps in its lists: first the workbook's defined names, then the
/// shared string table (each entry and each rich text run, not the text of
/// an entry, which the part's size bounds), then the records of
/// `xl/styles.xml` (number formats, fonts, fills and their gradient stops,
/// borders, cell formats, named cell styles, differential formats and the
/// tags a differential font lists as explicitly off), then the series and
/// axes of the charts, then the threaded comments and their mentions, then
/// the worksheet lists (data validations, conditional formats and their
/// rules, selections, columns, hyperlinks, merged ranges, page breaks,
/// filters, sparklines, kept extensions, root namespace declarations): each
/// at the size of its struct on the target (on wasm32
/// from half to about four fifths of the 64-bit size: a `DxfStyle` is 1,200
/// bytes on x86_64 and 784 on wasm32, a `ChartAxis` 216 and 176, measured
/// 2026-11-13) plus the heap
/// bytes of its strings. A string that is a copy counted in `spent` (a
/// hyperlink target, a declaration re-declared on an extension) is not
/// counted again. An element of a few bytes of XML can take a hundred times
/// that in memory (an empty `<dxf/>` of 6 bytes is a 1,200-byte record on a
/// 64-bit target, measured 2026-11-13), and the ZIP limits bound only the
/// bytes, so `stored` is held to a limit of its own, `store_limit`, which a
/// read sets to the same `max_json_length`.
///
/// A sequential read counts the defined names, the shared strings, the
/// styles, the charts, the threaded comments and then every sheet in one
/// budget, one sheet after another, so `stored` holds all of them together.
/// With the `parallel` feature each sheet counts from a copy of the budget
/// taken before the sheets (the defined names, shared strings, styles,
/// charts and threaded comments already counted), and only the copies in
/// `spent` are added back after them: the records counted before the sheets
/// with each sheet alone are held to `store_limit`, and the sheets of one
/// read together can keep up to their number times it.
#[derive(Debug, Clone, Copy)]
pub(crate) struct ExpansionBudget {
    /// Bytes counted so far.
    pub(crate) spent: usize,
    /// Bytes of kept list elements counted so far ([`Self::store`]).
    pub(crate) stored: usize,
    /// The most `spent` may reach; `usize::MAX` for no limit.
    pub(crate) limit: usize,
    /// The most `stored` may reach; `usize::MAX` for no limit.
    pub(crate) store_limit: usize,
}

impl ExpansionBudget {
    /// A budget of `limit` bytes of copies and `limit` bytes of stored
    /// list elements, with nothing spent or stored.
    pub(crate) const fn new(limit: usize) -> Self {
        Self {
            spent: 0,
            stored: 0,
            limit,
            store_limit: limit,
        }
    }

    /// Counts `bytes` more of memory for a record a parser keeps in a list (a
    /// defined name, a shared string entry or run, a styles.xml record, a
    /// chart series or axis, a threaded comment or mention, or a worksheet
    /// list element), before it is kept.
    ///
    /// # Errors
    ///
    /// `SizeLimit` once the count passes `store_limit`.
    #[inline]
    pub(crate) fn store(&mut self, bytes: usize) -> Result<()> {
        self.stored = self.stored.saturating_add(bytes);
        if self.stored > self.store_limit {
            cold_path();
            return Err(ModernXlsxError::SizeLimit(format!(
                "the defined names, shared strings, styles, chart series and axes, threaded comments and worksheet lists take more memory than the {}-byte limit (ReadOptions.limits.maxJsonLength / ZipSecurityLimits.max_json_length): each defined name, shared string entry, rich text run, font, fill, border, cell format, differential format, chart series or axis, threaded comment or mention, data validation, conditional format rule, selection, column, hyperlink, merged range or other list element kept is counted at the size of its record",
                self.store_limit
            )));
        }
        Ok(())
    }

    /// Counts `len` more bytes of copied text, before they are copied.
    ///
    /// # Errors
    ///
    /// `SizeLimit` once the count passes the limit.
    #[inline]
    pub(crate) fn spend(&mut self, len: usize) -> Result<()> {
        self.spent = self.spent.saturating_add(len);
        if self.spent > self.limit {
            cold_path();
            return Err(ModernXlsxError::SizeLimit(format!(
                "the workbook's shared text expands past the {}-byte limit (ReadOptions.limits.maxJsonLength / ZipSecurityLimits.max_json_length): a shared string, comment author or hyperlink target is copied into every cell, comment or hyperlink naming it, and a worksheet namespace declaration onto every extension using its prefix",
                self.limit
            )));
        }
        Ok(())
    }
}

/// The heap bytes of `run`'s strings: its text, font name, colour, vertical
/// alignment, scheme, underline style and the rgb of its colour spec.
fn run_heap(run: &RichTextRun) -> usize {
    run.text.capacity()
        + [
            &run.font_name,
            &run.color,
            &run.vert_align,
            &run.scheme,
            &run.underline_style,
        ]
        .into_iter()
        .flatten()
        .map(String::capacity)
        .sum::<usize>()
        + run
            .color_spec
            .as_ref()
            .and_then(|c| c.rgb.as_ref())
            .map_or(0, String::capacity)
}

/// The bytes `runs` take at least in a read's JSON: each run writes
/// `{"text":""}` (11 bytes) around its text, and its text and string
/// properties at no less than their byte length.
pub(crate) fn runs_json_floor(runs: &[RichTextRun]) -> usize {
    runs.iter()
        .map(|r| {
            11 + r.text.len()
                + [
                    &r.font_name,
                    &r.color,
                    &r.vert_align,
                    &r.scheme,
                    &r.underline_style,
                ]
                .into_iter()
                .flatten()
                .map(String::len)
                .sum::<usize>()
                + r.color_spec
                    .as_ref()
                    .and_then(|c| c.rgb.as_ref())
                    .map_or(0, String::len)
        })
        .sum()
}

/// A Shared String Table parsed from `sharedStrings.xml`.
///
/// In OOXML, cells that contain string values store an index into this table
/// rather than the string itself.
///
/// The `strings` vector always contains the plain-text value for every entry
/// (for rich text entries this is the concatenation of all run texts).
/// Phonetic runs (`<rPh>`, furigana) are not part of the text.
///
/// The `rich_runs` vector is parallel to `strings`: if entry *i* was a rich
/// text `<si>` with `<r>` elements, `rich_runs[i]` is `Some(runs)`;
/// otherwise it is `None`.
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct SharedStringTable {
    pub strings: Vec<String>,
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub rich_runs: Vec<Option<Vec<RichTextRun>>>,
}

impl SharedStringTable {
    /// Parse a Shared String Table from the raw XML bytes of `sharedStrings.xml`.
    ///
    /// # Errors
    ///
    /// `XmlParse` for malformed XML, a bad entity or character reference, or a
    /// run property whose attribute does not parse.
    pub fn parse(data: &[u8]) -> Result<Self> {
        Self::parse_within(data, &mut ExpansionBudget::new(usize::MAX))
    }

    /// [`Self::parse`], counting in `budget` (with [`ExpansionBudget::store`])
    /// each record the table keeps before it is kept: an entry at the size of
    /// its `String` and its `Option<Vec<RichTextRun>>`, and a rich text run at
    /// the size of its [`RichTextRun`] plus the capacity of its strings. The
    /// text of an entry is not counted: one byte of XML keeps at most one
    /// byte of it, so the part's size bounds it.
    ///
    /// # Errors
    ///
    /// As [`Self::parse`]; `SizeLimit` once the records pass
    /// `budget.store_limit`.
    pub(crate) fn parse_within(data: &[u8], budget: &mut ExpansionBudget) -> Result<Self> {
        /// The memory of one table entry beside its text.
        const ENTRY: usize = size_of::<String>() + size_of::<Option<Vec<RichTextRun>>>();
        let mut reader = Reader::from_reader(data);
        reader.config_mut().trim_text(false);

        let mut buf = Vec::with_capacity(512);
        let mut strings = Vec::new();
        let mut rich_runs: Vec<Option<Vec<RichTextRun>>> = Vec::new();

        // State tracking
        let mut in_si = false;
        let mut in_t = false;
        let mut in_r = false;
        let mut in_rpr = false;
        // Inside <rPh> (phonetic run): its <t> is furigana, not cell text.
        let mut in_rph = false;
        let mut current_text = String::new();

        // Track whether any rich text was found across all entries.
        let mut has_any_rich = false;

        // Rich text accumulation for the current <si>
        let mut runs: Vec<RichTextRun> = Vec::new();
        let mut run = RichTextRun::default();

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(e)) => {
                    let local = e.local_name();
                    match local.as_ref() {
                        "si" => {
                            in_si = true;
                            current_text.clear();
                            runs.clear();
                        }
                        "rPh" if in_si => in_rph = true,
                        "r" if in_si && !in_rph => {
                            in_r = true;
                            run = RichTextRun::default();
                        }
                        "rPr" if in_r => in_rpr = true,
                        "t" if in_si && !in_rph => in_t = true,
                        _ if in_rpr => apply_rpr_child(&mut run, &e)?,
                        _ => {}
                    }
                }
                // `<si/>` (every CT_Rst child is optional) is an empty entry;
                // skipping it would shift every later index.
                Ok(Event::Empty(e)) if !in_si && e.local_name().as_ref() == "si" => {
                    budget.store(ENTRY)?;
                    strings.push(String::new());
                    rich_runs.push(None);
                }
                Ok(Event::Empty(e)) if in_rpr => apply_rpr_child(&mut run, &e)?,
                Ok(Event::Text(e)) if in_t => {
                    let text = e.xml10_content();
                    current_text.push_str(&text);
                    if in_r {
                        run.text.push_str(&text);
                    }
                }
                Ok(Event::CData(e)) if in_t => {
                    let text = e.xml10_content();
                    current_text.push_str(&text);
                    if in_r {
                        run.text.push_str(&text);
                    }
                }
                Ok(Event::GeneralRef(e)) if in_t => {
                    // Resolve predefined XML entity references (&amp; &lt; &gt; &quot; &apos;)
                    // and character references (&#N; &#xN;) back to their characters.
                    push_general_ref(&mut current_text, &e)?;
                    if in_r {
                        push_general_ref(&mut run.text, &e)?;
                    }
                }
                Ok(Event::End(e)) => {
                    let local = e.local_name();
                    match local.as_ref() {
                        "t" => {
                            in_t = false;
                        }
                        "rPh" => {
                            in_rph = false;
                        }
                        "rPr" => {
                            in_rpr = false;
                        }
                        "r" if in_r => {
                            budget.store(size_of::<RichTextRun>() + run_heap(&run))?;
                            runs.push(std::mem::take(&mut run));
                            in_r = false;
                        }
                        "si" => {
                            budget.store(ENTRY)?;
                            strings.push(std::mem::take(&mut current_text));
                            if runs.is_empty() {
                                rich_runs.push(None);
                            } else {
                                has_any_rich = true;
                                // The runs move into a vector of their exact
                                // count, so the entry keeps no spare capacity
                                // beyond its counted runs, and `runs` keeps
                                // its buffer for the next entry.
                                let mut kept = Vec::with_capacity(runs.len());
                                kept.append(&mut runs);
                                rich_runs.push(Some(kept));
                            }
                            in_si = false;
                            in_rph = false;
                        }
                        _ => {}
                    }
                }
                Ok(Event::Eof) => break,
                Err(err) => {
                    cold_path();
                    return Err(ModernXlsxError::XmlParse(format!(
                        "error parsing sharedStrings.xml: {err}"
                    )));
                }
                _ => {}
            }
            buf.clear();
        }

        // If no rich text was found at all, keep rich_runs empty to avoid
        // serializing an all-None vector.
        if !has_any_rich {
            rich_runs.clear();
        }

        Ok(SharedStringTable { strings, rich_runs })
    }

    /// Return an empty Shared String Table.
    ///
    /// Used when the workbook does not contain a `sharedStrings.xml` part.
    pub fn empty() -> Self {
        SharedStringTable {
            strings: Vec::new(),
            rich_runs: Vec::new(),
        }
    }

    /// Return the rich text runs for the entry at `index`, if it is a rich
    /// text entry.
    #[cfg(test)]
    pub fn get_rich_runs(&self, index: usize) -> Option<&[RichTextRun]> {
        self.rich_runs.get(index).and_then(|opt| opt.as_deref())
    }

    /// Return the number of strings in the table.
    pub fn len(&self) -> usize {
        self.strings.len()
    }

    /// Return `true` if the table contains no strings.
    pub fn is_empty(&self) -> bool {
        self.strings.is_empty()
    }

    /// Look up a string by its zero-based index.
    pub fn get(&self, index: usize) -> Option<&str> {
        self.strings.get(index).map(String::as_str)
    }

    /// Serialize the shared string table to XML bytes.
    ///
    /// Rich text entries are written using `<r>` / `<rPr>` elements; plain
    /// text entries use a simple `<si><t>…</t></si>` form. A parsed table does
    /// not know how many cells reference each string, so `count` (the optional
    /// reference total) is omitted and only `uniqueCount` is written.
    #[cfg(test)]
    pub fn to_xml(&self) -> Result<Vec<u8>> {
        let mut buf: Vec<u8> = Vec::with_capacity(256 + self.strings.len() * 64);
        let mut writer = Writer::new(&mut buf);

        write_xml_decl(&mut writer)?;
        write_sst_start(&mut writer, None, self.strings.len())?;

        let mut xsd = XsdBuf::new();
        for (i, s) in self.strings.iter().enumerate() {
            let runs = self.rich_runs.get(i).and_then(|opt| opt.as_deref());
            write_si(&mut writer, &mut xsd, s, runs)?;
        }

        // </sst>
        writer
            .write_event(Event::End(BytesEnd::new("sst")))
            .map_err(xml_write_err)?;

        Ok(buf)
    }
}

/// A run list hashed for [`SharedStringTableBuilder::hash_runs`]: its length,
/// then every field of every run. Each run and colour is destructured without
/// `..`, so a field added later must be hashed here too. A float is hashed as
/// the bits of `v + 0.0`, which folds `-0.0` into `0.0` as `==` does.
struct RunsHash<'r>(&'r [RichTextRun]);

impl Hash for RunsHash<'_> {
    fn hash<H: Hasher>(&self, h: &mut H) {
        let float_bits = |v: &f64| (v + 0.0).to_bits();
        self.0.len().hash(h);
        for run in self.0 {
            let RichTextRun {
                text,
                bold,
                italic,
                underline,
                strike,
                font_name,
                font_size,
                color,
                vert_align,
                family,
                charset,
                scheme,
                color_spec,
                underline_style,
            } = run;
            (text, bold, italic, underline, strike, font_name).hash(h);
            (font_size.as_ref().map(float_bits), color, vert_align).hash(h);
            (family, charset, scheme, underline_style).hash(h);
            color_spec
                .as_ref()
                .map(
                    |Color {
                         rgb,
                         theme,
                         indexed,
                         tint,
                         auto,
                     }| {
                        (rgb, theme, indexed, tint.as_ref().map(float_bits), auto)
                    },
                )
                .hash(h);
        }
    }
}

/// A builder for constructing a Shared String Table when writing an XLSX file.
///
/// Entries are deduplicated, with plain and rich entries keyed separately: a
/// plain string is keyed by its text, a rich entry by its text and its runs.
/// Inserting the same plain string twice, or the same text with equal runs
/// twice, returns the same index both times. A plain and a rich entry with the
/// same text, or two rich entries with the same text and different runs, are
/// distinct entries, so no cell takes on another cell's formatting.
///
/// The ordered list and the lookup maps share one `Rc<str>` allocation per
/// distinct text. Rich text runs are borrowed from the workbook (`'a`) instead
/// of being cloned; an owned `Vec` is also accepted and becomes `Cow::Owned`.
///
/// Every `insert` and `insert_rich` call counts as one reference to an entry,
/// so `to_xml` writes `count` as the total number of references and
/// `uniqueCount` as the number of distinct entries.
///
/// `Rc` (not `Arc`) is deliberate: the builder is filled and serialized on one
/// thread. Switch to `Arc<str>` if writing ever goes parallel.
#[derive(Debug, Clone, Default)]
pub struct SharedStringTableBuilder<'a> {
    strings: Vec<Rc<str>>,
    /// Plain entries: text -> index.
    index: HashMap<Rc<str>, usize>,
    /// Rich entries: text -> run-list hash ([`Self::hash_runs`]) -> the
    /// indices of the rich entries with that text and hash (one per distinct
    /// run list), so a lookup compares runs only within one bucket.
    rich_index: HashMap<Rc<str>, HashMap<u64, Vec<usize>>>,
    /// Keys the run-list hash with the table's own `RandomState`. On native
    /// targets std draws the keys from the operating system's random source,
    /// so a file cannot choose runs that share a bucket ahead of time. In the
    /// `wasm32-unknown-unknown` build std derives them from a stack and a heap
    /// address, which are predictable, so there the keying only weakly stops
    /// such a file.
    run_hasher: RandomState,
    /// Rich text runs parallel to `strings`. If entry *i* has rich text,
    /// `rich_runs[i]` is `Some(runs)`, otherwise `None`.
    rich_runs: Vec<Option<Cow<'a, [RichTextRun]>>>,
    /// Total `insert` / `insert_rich` calls: the SST `count`.
    references: usize,
}

impl<'a> SharedStringTableBuilder<'a> {
    /// Create a new, empty builder.
    pub fn new() -> Self {
        Self::default()
    }

    /// Return the number of unique strings inserted so far.
    pub fn len(&self) -> usize {
        self.strings.len()
    }

    /// Return `true` if no strings have been inserted.
    pub fn is_empty(&self) -> bool {
        self.strings.is_empty()
    }

    /// Look up the index of a previously inserted entry: the plain entry for
    /// `text` when `runs` is `None`, else the rich entry for `text` whose runs
    /// equal `runs`. Runs are compared only against the entries whose run
    /// list hashes the same, so a lookup does not grow with the number of
    /// rich entries sharing `text`.
    ///
    /// Returns `None` if no such entry has been inserted.
    pub fn index_of(&self, text: &str, runs: Option<&[RichTextRun]>) -> Option<usize> {
        let Some(runs) = runs else {
            return self.index.get(text).copied();
        };
        self.find_rich(text, self.hash_runs(runs), runs)
    }

    /// The rich entry for `text` whose runs equal `runs`, searched only in
    /// the bucket of `h`, the [`Self::hash_runs`] of `runs`.
    fn find_rich(&self, text: &str, h: u64, runs: &[RichTextRun]) -> Option<usize> {
        self.rich_index
            .get(text)?
            .get(&h)?
            .iter()
            .copied()
            .find(|&i| self.rich_runs.get(i).and_then(Option::as_deref) == Some(runs))
    }

    /// The bucket key of a run list in `rich_index`: a keyed hash of its
    /// length and every field of every run, colour spec included. Run lists
    /// that compare equal hash equal: `-0.0` and `0.0` hash alike because a
    /// float is hashed as the bits of `v + 0.0`.
    fn hash_runs(&self, runs: &[RichTextRun]) -> u64 {
        self.run_hasher.hash_one(RunsHash(runs))
    }

    /// The shared `Rc<str>` for `text` when any entry already holds it, else a
    /// new one.
    fn key_for(&self, text: &str) -> Rc<str> {
        self.index
            .get_key_value(text)
            .map(|(k, _)| k)
            .or_else(|| self.rich_index.get_key_value(text).map(|(k, _)| k))
            .map_or_else(|| Rc::from(text), Rc::clone)
    }

    /// Insert a plain string and return its zero-based index.
    ///
    /// If the string has been inserted as plain text before, the existing
    /// index is returned without adding a duplicate entry or allocating; a
    /// rich entry with the same text is not reused. Either way the call counts
    /// as one reference.
    pub fn insert(&mut self, s: &str) -> usize {
        self.references += 1;
        if let Some(&i) = self.index.get(s) {
            return i;
        }
        let i = self.strings.len();
        let key = self.key_for(s);
        self.strings.push(Rc::clone(&key));
        self.rich_runs.push(None);
        self.index.insert(key, i);
        i
    }

    /// Insert a rich text entry and return its zero-based index.
    ///
    /// `key` is the plain text, the concatenation of all run texts. A rich
    /// entry with the same text and equal runs is reused; otherwise a new
    /// entry is added. A plain entry with the same text is never reused or
    /// changed. Either way the call counts as one reference.
    pub fn insert_rich(&mut self, key: &str, runs: impl Into<Cow<'a, [RichTextRun]>>) -> usize {
        self.references += 1;
        let runs = runs.into();
        let h = self.hash_runs(&runs);
        if let Some(i) = self.find_rich(key, h, &runs) {
            return i;
        }
        let i = self.strings.len();
        let k = self.key_for(key);
        self.strings.push(Rc::clone(&k));
        self.rich_runs.push(Some(runs));
        self.rich_index
            .entry(k)
            .or_default()
            .entry(h)
            .or_default()
            .push(i);
        i
    }

    /// Return the rich text runs for the entry at `index`, if it is a rich
    /// text entry.
    #[cfg(test)]
    pub fn get_rich_runs(&self, index: usize) -> Option<&[RichTextRun]> {
        self.rich_runs.get(index).and_then(|opt| opt.as_deref())
    }

    /// Serialize the shared string table to XML bytes suitable for
    /// inclusion in an XLSX archive as `xl/sharedStrings.xml`.
    ///
    /// `count` is the number of `insert` / `insert_rich` calls (the cell
    /// references), `uniqueCount` the number of distinct entries.
    ///
    /// # Errors
    ///
    /// `InvalidWorkbook` when a run colour's `rgb` is not 6 or 8 hex digits
    /// (see `color::checked_argb`); `XmlWrite` when a run's font size or
    /// colour tint is not finite.
    pub fn to_xml(&self) -> Result<Vec<u8>> {
        let mut buf: Vec<u8> = Vec::with_capacity(256 + self.strings.len() * 64);
        self.write_to(&mut buf)?;
        Ok(buf)
    }

    /// Write the same bytes as [`Self::to_xml`] straight into `out` (a ZIP
    /// entry, for the streaming writer) without buffering the whole part.
    ///
    /// # Errors
    ///
    /// `InvalidWorkbook` when a run colour's `rgb` is not 6 or 8 hex digits
    /// (see `color::checked_argb`); `XmlWrite` when `out` fails or a run's
    /// font size or colour tint is not finite.
    pub(crate) fn write_to(&self, out: &mut dyn Write) -> Result<()> {
        let mut writer = Writer::new(out);

        write_xml_decl(&mut writer)?;
        write_sst_start(&mut writer, Some(self.references), self.strings.len())?;

        // Each string: <si><t>text</t></si> or <si><r>...</r>...</si>
        let mut xsd = XsdBuf::new();
        for (i, s) in self.strings.iter().enumerate() {
            let runs = self.rich_runs.get(i).and_then(|o| o.as_deref());
            write_si(&mut writer, &mut xsd, s, runs)?;
        }

        // </sst>
        writer
            .write_event(Event::End(BytesEnd::new("sst")))
            .map_err(xml_write_err)
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use pretty_assertions::assert_eq;

    #[test]
    fn test_parse_sst() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="3" uniqueCount="3">
  <si><t>Hello</t></si>
  <si><t>World</t></si>
  <si><t>Rust</t></si>
</sst>"#;

        let sst = SharedStringTable::parse(xml.as_bytes()).unwrap();
        assert_eq!(sst.len(), 3);
        assert!(!sst.is_empty());
        assert_eq!(sst.get(0), Some("Hello"));
        assert_eq!(sst.get(1), Some("World"));
        assert_eq!(sst.get(2), Some("Rust"));
    }

    #[test]
    fn test_sst_builder() {
        let mut builder = SharedStringTableBuilder::new();
        let i0 = builder.insert("Hello");
        let i1 = builder.insert("World");
        let i2 = builder.insert("Hello"); // duplicate

        assert_eq!(i0, 0);
        assert_eq!(i1, 1);
        assert_eq!(i2, 0); // dedup: same index as first "Hello"
        assert_eq!(builder.len(), 2);
    }

    #[test]
    fn test_sst_write_roundtrip() {
        let mut builder = SharedStringTableBuilder::new();
        builder.insert("Alpha");
        builder.insert("Beta");
        builder.insert("Gamma");
        builder.insert("Alpha"); // duplicate — should not add a new entry

        assert_eq!(builder.len(), 3);

        let xml = builder.to_xml().unwrap();
        let sst = SharedStringTable::parse(&xml).unwrap();

        assert_eq!(sst.len(), 3);
        assert_eq!(sst.get(0), Some("Alpha"));
        assert_eq!(sst.get(1), Some("Beta"));
        assert_eq!(sst.get(2), Some("Gamma"));
    }

    #[test]
    fn test_sst_out_of_bounds() {
        let sst = SharedStringTable::empty();
        assert_eq!(sst.get(999), None);

        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1">
  <si><t>Only</t></si>
</sst>"#;
        let sst = SharedStringTable::parse(xml.as_bytes()).unwrap();
        assert_eq!(sst.get(0), Some("Only"));
        assert_eq!(sst.get(999), None);
    }

    #[test]
    fn test_sst_empty() {
        let builder = SharedStringTableBuilder::new();
        assert!(builder.is_empty());

        let xml = builder.to_xml().unwrap();
        let sst = SharedStringTable::parse(&xml).unwrap();

        assert!(sst.is_empty());
        assert_eq!(sst.len(), 0);
    }

    #[test]
    fn test_parse_rich_text() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="2" uniqueCount="2">
  <si><t>Plain text</t></si>
  <si>
    <r>
      <rPr><b/><sz val="14"/><color rgb="FF0000"/><rFont val="Arial"/></rPr>
      <t>Bold red text</t>
    </r>
    <r>
      <t> normal text</t>
    </r>
  </si>
</sst>"#;

        let sst = SharedStringTable::parse(xml.as_bytes()).unwrap();
        assert_eq!(sst.len(), 2);

        // First entry is plain text.
        assert_eq!(sst.get(0), Some("Plain text"));
        assert!(sst.get_rich_runs(0).is_none());

        // Second entry is rich text — plain text is the concatenation.
        assert_eq!(sst.get(1), Some("Bold red text normal text"));

        let runs = sst.get_rich_runs(1).expect("should have rich runs");
        assert_eq!(runs.len(), 2);

        // First run: bold, font size 14, red, Arial.
        assert_eq!(runs[0].text, "Bold red text");
        assert_eq!(runs[0].bold, Some(true));
        assert_eq!(runs[0].italic, None);
        assert_eq!(runs[0].font_size, Some(14.0));
        assert_eq!(runs[0].color.as_deref(), Some("FF0000"));
        assert_eq!(runs[0].font_name.as_deref(), Some("Arial"));

        // Second run: no formatting.
        assert_eq!(runs[1].text, " normal text");
        assert_eq!(runs[1].bold, None);
        assert_eq!(runs[1].italic, None);
        assert_eq!(runs[1].font_size, None);
        assert_eq!(runs[1].color, None);
        assert_eq!(runs[1].font_name, None);
    }

    #[test]
    fn test_parse_rich_text_italic() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1">
  <si>
    <r>
      <rPr><i/></rPr>
      <t>italic</t>
    </r>
  </si>
</sst>"#;

        let sst = SharedStringTable::parse(xml.as_bytes()).unwrap();
        assert_eq!(sst.len(), 1);
        assert_eq!(sst.get(0), Some("italic"));

        let runs = sst.get_rich_runs(0).expect("should have rich runs");
        assert_eq!(runs.len(), 1);
        assert_eq!(runs[0].italic, Some(true));
        assert_eq!(runs[0].bold, None);
    }

    #[test]
    fn test_rich_text_roundtrip_via_sst_to_xml() {
        // Build an SST with both plain and rich text entries.
        let sst = SharedStringTable {
            strings: vec!["Plain".to_string(), "Bold italic".to_string()],
            rich_runs: vec![
                None,
                Some(vec![
                    RichTextRun {
                        text: "Bold".to_string(),
                        bold: Some(true),
                        italic: None,
                        underline: None,
                        strike: None,
                        font_name: Some("Calibri".to_string()),
                        font_size: Some(12.0),
                        color: Some("FF0000".to_string()),
                        ..Default::default()
                    },
                    RichTextRun {
                        text: " italic".to_string(),
                        bold: None,
                        italic: Some(true),
                        underline: None,
                        strike: None,
                        font_name: None,
                        font_size: None,
                        color: None,
                        ..Default::default()
                    },
                ]),
            ],
        };

        // Write to XML and re-parse.
        let xml = sst.to_xml().unwrap();
        let sst2 = SharedStringTable::parse(&xml).unwrap();

        assert_eq!(sst2.len(), 2);

        // Plain entry preserved.
        assert_eq!(sst2.get(0), Some("Plain"));
        assert!(sst2.get_rich_runs(0).is_none());

        // Rich entry preserved.
        assert_eq!(sst2.get(1), Some("Bold italic"));
        let runs = sst2.get_rich_runs(1).expect("should have rich runs");
        assert_eq!(runs.len(), 2);
        assert_eq!(runs[0].text, "Bold");
        assert_eq!(runs[0].bold, Some(true));
        assert_eq!(runs[0].font_name.as_deref(), Some("Calibri"));
        assert_eq!(runs[0].font_size, Some(12.0));
        assert_eq!(runs[0].color.as_deref(), Some("FFFF0000"));
        assert_eq!(runs[1].text, " italic");
        assert_eq!(runs[1].italic, Some(true));
    }

    #[test]
    fn test_plain_only_sst_has_empty_rich_runs() {
        // When all entries are plain, rich_runs should be empty (not
        // a vector of Nones) so that serialization stays compact.
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="2" uniqueCount="2">
  <si><t>A</t></si>
  <si><t>B</t></si>
</sst>"#;

        let sst = SharedStringTable::parse(xml.as_bytes()).unwrap();
        assert!(
            sst.rich_runs.is_empty(),
            "rich_runs should be empty for plain-only SST"
        );
    }

    #[test]
    fn test_rich_text_serde_json_roundtrip() {
        let sst = SharedStringTable {
            strings: vec!["hello world".to_string()],
            rich_runs: vec![Some(vec![
                RichTextRun {
                    text: "hello".to_string(),
                    bold: Some(true),
                    italic: None,
                    underline: None,
                    strike: None,
                    font_name: None,
                    font_size: None,
                    color: None,
                    ..Default::default()
                },
                RichTextRun {
                    text: " world".to_string(),
                    bold: None,
                    italic: None,
                    underline: None,
                    strike: None,
                    font_name: None,
                    font_size: None,
                    color: None,
                    ..Default::default()
                },
            ])],
        };

        let json = serde_json::to_string(&sst).unwrap();
        let sst2: SharedStringTable = serde_json::from_str(&json).unwrap();

        assert_eq!(sst2.strings, sst.strings);
        assert_eq!(sst2.rich_runs.len(), 1);
        let runs = sst2.rich_runs[0].as_ref().unwrap();
        assert_eq!(runs.len(), 2);
        assert_eq!(runs[0].text, "hello");
        assert_eq!(runs[0].bold, Some(true));
        assert_eq!(runs[1].text, " world");
    }

    #[test]
    fn test_parse_rich_text_underline_strike() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1">
  <si>
    <r>
      <rPr><u/><strike/></rPr>
      <t>crossed out</t>
    </r>
  </si>
</sst>"#;

        let sst = SharedStringTable::parse(xml.as_bytes()).unwrap();
        assert_eq!(sst.len(), 1);
        assert_eq!(sst.get(0), Some("crossed out"));

        let runs = sst.get_rich_runs(0).expect("should have rich runs");
        assert_eq!(runs.len(), 1);
        assert_eq!(runs[0].underline, Some(true));
        assert_eq!(runs[0].strike, Some(true));
        assert_eq!(runs[0].bold, None);
        assert_eq!(runs[0].italic, None);
    }

    #[test]
    fn test_underline_strike_roundtrip_via_sst_to_xml() {
        let sst = SharedStringTable {
            strings: vec!["underline strike".to_string()],
            rich_runs: vec![Some(vec![RichTextRun {
                text: "underline strike".to_string(),
                bold: None,
                italic: None,
                underline: Some(true),
                strike: Some(true),
                font_name: None,
                font_size: None,
                color: None,
                ..Default::default()
            }])],
        };

        let xml = sst.to_xml().unwrap();
        let sst2 = SharedStringTable::parse(&xml).unwrap();

        let runs = sst2.get_rich_runs(0).expect("should have rich runs");
        assert_eq!(runs.len(), 1);
        assert_eq!(runs[0].underline, Some(true));
        assert_eq!(runs[0].strike, Some(true));
    }

    #[test]
    fn builder_insert_dedupes() {
        let mut builder = SharedStringTableBuilder::new();
        assert_eq!(builder.insert("a"), 0);
        assert_eq!(builder.insert("a"), 0);
        assert_eq!(builder.len(), 1);
    }

    #[test]
    fn builder_insert_rich_borrows() {
        let runs = vec![RichTextRun {
            text: "x".into(),
            bold: None,
            italic: None,
            underline: None,
            strike: None,
            font_name: None,
            font_size: None,
            color: None,
            ..Default::default()
        }];
        let mut builder = SharedStringTableBuilder::new();
        assert_eq!(builder.insert_rich("x", runs.as_slice()), 0);
        let got = builder.get_rich_runs(0).expect("rich runs");
        assert_eq!(got.len(), 1);
        assert_eq!(got[0].text, "x");
    }

    #[test]
    fn test_builder_insert_rich() {
        let mut builder = SharedStringTableBuilder::new();
        builder.insert("plain");
        let runs = vec![
            RichTextRun {
                text: "bold".to_string(),
                bold: Some(true),
                italic: None,
                underline: None,
                strike: None,
                font_name: None,
                font_size: None,
                color: None,
                ..Default::default()
            },
            RichTextRun {
                text: " normal".to_string(),
                bold: None,
                italic: None,
                underline: None,
                strike: None,
                font_name: None,
                font_size: None,
                color: None,
                ..Default::default()
            },
        ];
        let idx = builder.insert_rich("bold normal", runs);
        assert_eq!(idx, 1);
        assert_eq!(builder.len(), 2);

        // Verify rich text survives write → parse roundtrip.
        let xml = builder.to_xml().unwrap();
        let sst = SharedStringTable::parse(&xml).unwrap();

        assert_eq!(sst.get(0), Some("plain"));
        assert!(sst.get_rich_runs(0).is_none());

        assert_eq!(sst.get(1), Some("bold normal"));
        let parsed_runs = sst.get_rich_runs(1).expect("should have rich runs");
        assert_eq!(parsed_runs.len(), 2);
        assert_eq!(parsed_runs[0].text, "bold");
        assert_eq!(parsed_runs[0].bold, Some(true));
        assert_eq!(parsed_runs[1].text, " normal");
    }

    #[test]
    fn builder_keys_plain_and_rich_entries_apart() {
        let bold = vec![RichTextRun {
            text: "hello".to_string(),
            bold: Some(true),
            ..Default::default()
        }];
        let italic = vec![RichTextRun {
            text: "hello".to_string(),
            italic: Some(true),
            ..Default::default()
        }];
        let mut builder = SharedStringTableBuilder::new();
        assert_eq!(builder.insert("hello"), 0);
        // The same text as rich text is its own entry; the plain one stays plain.
        assert_eq!(builder.insert_rich("hello", bold.as_slice()), 1);
        assert!(builder.get_rich_runs(0).is_none());
        // Equal runs reuse the rich entry, different runs get a new one.
        assert_eq!(builder.insert_rich("hello", bold.clone()), 1);
        assert_eq!(builder.insert_rich("hello", italic.as_slice()), 2);
        assert_eq!(builder.insert("hello"), 0);
        assert_eq!(builder.len(), 3);

        let sst = SharedStringTable::parse(&builder.to_xml().unwrap()).unwrap();
        assert_eq!(sst.strings, ["hello", "hello", "hello"]);
        assert!(sst.get_rich_runs(0).is_none());
        assert_eq!(sst.get_rich_runs(1).unwrap()[0].bold, Some(true));
        assert_eq!(sst.get_rich_runs(2).unwrap()[0].italic, Some(true));
    }

    #[test]
    fn test_xml_entity_roundtrip() {
        // Verify XML special characters survive SST write → parse roundtrip.
        let mut builder = SharedStringTableBuilder::new();
        builder.insert("Tom & Jerry");
        builder.insert("2 < 3");
        builder.insert("5 > 4");
        builder.insert(r#"She said "hello""#);
        builder.insert("it's fine");

        let xml = builder.to_xml().unwrap();

        let sst = SharedStringTable::parse(&xml).unwrap();
        assert_eq!(sst.get(0), Some("Tom & Jerry"));
        assert_eq!(sst.get(1), Some("2 < 3"));
        assert_eq!(sst.get(2), Some("5 > 4"));
        assert_eq!(sst.get(3), Some(r#"She said "hello""#));
        assert_eq!(sst.get(4), Some("it's fine"));
    }

    const SST_HEAD: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1">"#;

    fn parse(body: &str) -> SharedStringTable {
        SharedStringTable::parse(format!("{SST_HEAD}{body}</sst>").as_bytes()).unwrap()
    }

    #[test]
    fn parse_within_counts_each_entry_and_run_before_keeping_it() {
        let xml = format!(
            r#"{SST_HEAD}<si/><si><t>plain text is not counted</t></si><si><r><t>ab</t></r><r><rPr><rFont val="Arial"/><color rgb="FF00FF00" theme="1"/><vertAlign val="superscript"/><scheme val="minor"/><u val="double"/></rPr><t>c</t></r><r></r></si></sst>"#
        );
        let table = SharedStringTable::parse(xml.as_bytes()).unwrap();
        let runs = table.rich_runs[2].as_deref().unwrap();
        assert_eq!(runs.len(), 3);
        let entry = size_of::<String>() + size_of::<Option<Vec<RichTextRun>>>();
        let expected = 3 * entry
            + runs
                .iter()
                .map(|r| size_of::<RichTextRun>() + run_heap(r))
                .sum::<usize>();
        // The second run's strings: "c", Arial, FF00FF00 twice (colour and
        // colour spec), superscript, minor, double.
        assert!(run_heap(&runs[1]) >= 1 + 5 + 8 + 8 + 11 + 5 + 6, "{runs:?}");
        let mut budget = ExpansionBudget::new(expected);
        let within = SharedStringTable::parse_within(xml.as_bytes(), &mut budget).unwrap();
        assert_eq!(within.strings, table.strings);
        assert_eq!(within.rich_runs, table.rich_runs);
        assert_eq!((budget.stored, budget.spent), (expected, 0));
        // Each entry keeps no spare run capacity beyond what was counted.
        assert_eq!(within.rich_runs[2].as_ref().map(Vec::capacity), Some(3));
        let err = SharedStringTable::parse_within(
            xml.as_bytes(),
            &mut ExpansionBudget::new(expected - 1),
        )
        .unwrap_err();
        assert_eq!(err.code(), "SIZE_LIMIT", "{err}");
        assert!(err.to_string().contains("shared string entry"), "{err}");
    }

    fn xml_str(sst: &SharedStringTable) -> String {
        String::from_utf8(sst.to_xml().unwrap()).unwrap()
    }

    #[test]
    fn phonetic_run_is_not_part_of_the_text() {
        let sst = parse(
            r#"<si><t>BASE</t><rPh sb="0" eb="1"><t>FURI</t></rPh><phoneticPr fontId="1"/></si>"#,
        );
        assert_eq!(sst.strings, ["BASE"]);
        assert!(sst.rich_runs.is_empty());
    }

    #[test]
    fn phonetic_run_next_to_rich_runs_keeps_only_run_texts() {
        let sst = parse(
            r#"<si><r><t>Hel</t></r><r><rPr><b/></rPr><t>lo</t></r><rPh sb="0" eb="1"><t>FURI</t></rPh><phoneticPr fontId="1"/></si><si><t>next</t></si>"#,
        );
        assert_eq!(sst.strings, ["Hello", "next"]);
        let runs = sst.get_rich_runs(0).unwrap();
        assert_eq!(runs.len(), 2);
        assert_eq!(runs[0].text, "Hel");
        assert_eq!(runs[1].text, "lo");
        assert_eq!(runs[1].bold, Some(true));
        assert!(sst.get_rich_runs(1).is_none());
    }

    #[test]
    fn boolean_run_properties_read_their_val() {
        let sst = parse(
            r#"<si><r><rPr><b val="0"/><i val="0"/><u val="double"/><strike val="false"/></rPr><t>x</t></r></si>"#,
        );
        let run = &sst.get_rich_runs(0).unwrap()[0];
        assert_eq!(run.bold, Some(false));
        assert_eq!(run.italic, Some(false));
        assert_eq!(run.underline, Some(true));
        assert_eq!(run.underline_style.as_deref(), Some("double"));
        assert_eq!(run.strike, Some(false));
        assert_eq!(
            xml_str(&sst)
                .split("<rPr>")
                .nth(1)
                .unwrap()
                .split("</rPr>")
                .next()
                .unwrap(),
            r#"<b val="0"/><i val="0"/><u val="double"/><strike val="0"/>"#
        );
    }

    #[test]
    fn underline_none_is_off_and_plain_underline_has_no_style() {
        let sst = parse(
            r#"<si><r><rPr><u val="none"/></rPr><t>a</t></r><r><rPr><u/></rPr><t>b</t></r><r><rPr><u val="single"/></rPr><t>c</t></r></si>"#,
        );
        let runs = sst.get_rich_runs(0).unwrap();
        assert_eq!(runs[0].underline, Some(false));
        assert_eq!(runs[0].underline_style, None);
        assert_eq!(runs[1].underline, Some(true));
        assert_eq!(runs[1].underline_style, None);
        assert_eq!(runs[2].underline, Some(true));
        assert_eq!(runs[2].underline_style, None);
        let xml = xml_str(&sst);
        assert!(xml.contains(r#"<rPr><u val="none"/></rPr>"#), "{xml}");
        assert!(xml.contains("<rPr><u/></rPr>"), "{xml}");
    }

    #[test]
    fn edited_underline_boolean_wins_over_style() {
        let run = |underline, style: Option<&str>| RichTextRun {
            text: "x".to_owned(),
            underline,
            underline_style: style.map(str::to_owned),
            ..Default::default()
        };
        let rpr = |r: RichTextRun| {
            let sst = SharedStringTable {
                strings: vec!["x".to_owned()],
                rich_runs: vec![Some(vec![r])],
            };
            xml_str(&sst)
        };
        let off = rpr(run(Some(false), Some("double")));
        assert!(off.contains(r#"<rPr><u val="none"/></rPr>"#), "{off}");
        let on = rpr(run(Some(true), Some("none")));
        assert!(on.contains("<rPr><u/></rPr>"), "{on}");
    }

    #[test]
    fn full_run_properties_round_trip() {
        let body = r#"<si><r><rPr><vertAlign val="superscript"/><color theme="1" tint="-0.25"/><family val="2"/><charset val="204"/><scheme val="minor"/><rFont val="Aptos"/><sz val="11"/></rPr><t>2</t></r></si>"#;
        let sst = parse(body);
        let run = &sst.get_rich_runs(0).unwrap()[0];
        assert_eq!(run.vert_align.as_deref(), Some("superscript"));
        assert_eq!(run.family, Some(2));
        assert_eq!(run.charset, Some(204));
        assert_eq!(run.scheme.as_deref(), Some("minor"));
        assert_eq!(run.color, None);
        assert_eq!(
            run.color_spec,
            Some(Color {
                theme: Some(1),
                tint: Some(-0.25),
                ..Color::default()
            })
        );
        let json = serde_json::to_value(run).unwrap();
        assert_eq!(json["vertAlign"], "superscript");
        assert_eq!(json["colorSpec"]["theme"], 1);
        assert_eq!(json["family"], 2);
        assert_eq!(json["charset"], 204);
        assert_eq!(json["scheme"], "minor");

        let xml = xml_str(&sst);
        for needle in [
            r#"<vertAlign val="superscript"/>"#,
            r#"<color theme="1" tint="-0.25"/>"#,
            r#"<family val="2"/>"#,
            r#"<charset val="204"/>"#,
            r#"<scheme val="minor"/>"#,
        ] {
            assert!(xml.contains(needle), "missing {needle} in {xml}");
        }
        let again = SharedStringTable::parse(xml.as_bytes()).unwrap();
        assert_eq!(
            serde_json::to_value(&again).unwrap(),
            serde_json::to_value(&sst).unwrap()
        );
    }

    #[test]
    fn start_form_run_properties_are_read() {
        let sst = parse(
            r#"<si><r><rPr><sz val="9.5"></sz><color rgb="FF00FF00"></color></rPr><t>a</t></r></si>"#,
        );
        let run = &sst.get_rich_runs(0).unwrap()[0];
        assert_eq!(run.font_size, Some(9.5));
        assert_eq!(run.color.as_deref(), Some("FF00FF00"));
        assert_eq!(run.color_spec, None);
    }

    #[test]
    fn color_spec_loses_to_a_changed_legacy_color() {
        let run = RichTextRun {
            text: "x".into(),
            color: Some("FF112233".into()),
            color_spec: Some(Color {
                theme: Some(4),
                ..Color::default()
            }),
            ..Default::default()
        };
        let sst = SharedStringTable {
            strings: vec!["x".into()],
            rich_runs: vec![Some(vec![run])],
        };
        let xml = xml_str(&sst);
        assert!(xml.contains(r#"<color rgb="FF112233"/>"#), "{xml}");
        assert!(!xml.contains("theme"), "{xml}");
    }

    #[test]
    fn run_color_is_argb_and_an_all_default_spec_writes_nothing() {
        let sst_xml = |run: RichTextRun| {
            let mut builder = SharedStringTableBuilder::new();
            builder.insert_rich("x", vec![run]);
            String::from_utf8(builder.to_xml().unwrap()).unwrap()
        };
        let red = sst_xml(RichTextRun {
            text: "x".into(),
            color: Some("FF0000".into()),
            ..Default::default()
        });
        assert!(
            red.contains(r#"<r><rPr><color rgb="FFFF0000"/></rPr><t>x</t></r>"#),
            "{red}"
        );
        let blank = sst_xml(RichTextRun {
            text: "x".into(),
            color_spec: Some(Color::default()),
            ..Default::default()
        });
        assert!(blank.contains("<si><r><t>x</t></r></si>"), "{blank}");
        assert!(!blank.contains("rPr"), "{blank}");
    }

    #[test]
    fn bad_run_property_value_is_an_error() {
        let xml = format!(r#"{SST_HEAD}<si><r><rPr><sz val="big"/></rPr><t>x</t></r></si></sst>"#);
        assert!(matches!(
            SharedStringTable::parse(xml.as_bytes()),
            Err(ModernXlsxError::XmlParse(_))
        ));
    }

    #[test]
    fn builder_count_is_the_reference_count() {
        let mut builder = SharedStringTableBuilder::new();
        for _ in 0..3 {
            builder.insert("x");
        }
        let xml = String::from_utf8(builder.to_xml().unwrap()).unwrap();
        assert!(xml.contains(r#"count="3" uniqueCount="1""#), "{xml}");
    }

    #[test]
    fn builder_counts_rich_inserts_too() {
        let mut builder = SharedStringTableBuilder::new();
        builder.insert("ab");
        builder.insert_rich("ab", vec![RichTextRun::default()]);
        builder.insert("cd");
        builder.insert_rich("ab", vec![RichTextRun::default()]);
        let xml = String::from_utf8(builder.to_xml().unwrap()).unwrap();
        // Plain "ab", rich "ab" and "cd" are three entries; four references.
        assert!(xml.contains(r#"count="4" uniqueCount="3""#), "{xml}");
    }

    #[test]
    fn parsed_table_omits_count() {
        let sst = parse("<si><t>a</t></si>");
        let xml = xml_str(&sst);
        assert!(xml.contains(r#"<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" uniqueCount="1">"#), "{xml}");
    }

    #[test]
    fn edge_spaces_are_preserved_and_illegal_chars_dropped() {
        let sst = parse(r#"<si><t xml:space="preserve">  a  </t></si>"#);
        assert_eq!(sst.get(0), Some("  a  "));
        let xml = xml_str(&sst);
        assert!(
            xml.contains(r#"<t xml:space="preserve">  a  </t>"#),
            "{xml}"
        );

        let mut builder = SharedStringTableBuilder::new();
        builder.insert("a\u{1}b");
        let out = String::from_utf8(builder.to_xml().unwrap()).unwrap();
        assert!(out.contains("<t>ab</t>"), "{out}");
        assert!(!out.contains('\u{1}'));
    }

    #[test]
    fn space_behind_a_dropped_control_char_is_preserved() {
        let mut builder = SharedStringTableBuilder::new();
        builder.insert("\u{1} total");
        let out = String::from_utf8(builder.to_xml().unwrap()).unwrap();
        assert!(
            out.contains(r#"<si><t xml:space="preserve"> total</t></si>"#),
            "{out}"
        );
    }

    #[test]
    fn cdata_text_is_read() {
        let sst = parse("<si><t><![CDATA[a<b]]></t></si>");
        assert_eq!(sst.get(0), Some("a<b"));
        let rich = parse("<si><r><t><![CDATA[x&y]]></t></r></si>");
        assert_eq!(rich.get_rich_runs(0).unwrap()[0].text, "x&y");
    }

    #[test]
    fn self_closing_si_is_an_empty_entry() {
        let sst = SharedStringTable::parse(
            br#"<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><si><t>a</t></si><si/><si><t>c</t></si></sst>"#,
        )
        .unwrap();
        assert_eq!(sst.strings, ["a", "", "c"]);
        // A plain-only table keeps `rich_runs` empty (no all-`None` vector).
        assert!(sst.rich_runs.is_empty());

        // Next to a rich entry, `rich_runs` stays parallel to `strings`.
        let rich = SharedStringTable::parse(
            br#"<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><si><r><rPr><b/></rPr><t>a</t></r></si><si/><si><t>c</t></si></sst>"#,
        )
        .unwrap();
        assert_eq!(rich.strings, ["a", "", "c"]);
        assert_eq!(rich.rich_runs.len(), 3);
        assert!(rich.rich_runs[0].is_some());
        assert!(rich.rich_runs[1].is_none());
        assert!(rich.rich_runs[2].is_none());
    }

    #[test]
    fn malformed_xml_is_xml_parse() {
        assert!(matches!(
            SharedStringTable::parse(b"<sst><a></b></sst>"),
            Err(ModernXlsxError::XmlParse(_))
        ));
    }

    #[test]
    fn malformed_xml_error_names_the_part() {
        match SharedStringTable::parse(b"<sst><si></sst>") {
            Err(ModernXlsxError::XmlParse(msg)) => {
                assert!(msg.contains("sharedStrings.xml"), "{msg}");
            }
            other => panic!("expected XmlParse, got {other:?}"),
        }
    }

    #[test]
    fn literal_line_ends_are_normalized_and_char_refs_kept() {
        let sst = parse(
            "<si><t>a\r\nb\rc</t></si><si><t>a&#13;&#10;b</t></si><si><r><t>x\r\ny</t></r></si>",
        );
        assert_eq!(sst.get(0), Some("a\nb\nc"));
        assert_eq!(sst.get(1), Some("a\r\nb"));
        assert_eq!(sst.get(2), Some("x\ny"));
        assert_eq!(sst.get_rich_runs(2).unwrap()[0].text, "x\ny");
    }

    /// One run of text `a` with font size `size` and a colour tinted by `tint`.
    fn sized_run(size: f64, tint: f64) -> Vec<RichTextRun> {
        vec![RichTextRun {
            text: "a".into(),
            font_size: Some(size),
            color_spec: Some(Color {
                theme: Some(1),
                tint: Some(tint),
                ..Color::default()
            }),
            ..RichTextRun::default()
        }]
    }

    /// N rich entries that share one text must not cost O(N^2): a linear
    /// scan of the text's entries makes ~1.25e9 run comparisons per pass here.
    #[test]
    fn many_rich_entries_with_one_text_insert_and_look_up_in_linear_time() {
        const N: usize = 50_000;
        let runs: Vec<Vec<RichTextRun>> = (1..=N)
            .map(|k| sized_run(k as f64, k as f64 / N as f64))
            .collect();
        let start = std::time::Instant::now();
        let mut builder = SharedStringTableBuilder::new();
        for (i, r) in runs.iter().enumerate() {
            assert_eq!(builder.insert_rich("a", r.as_slice()), i);
        }
        for (i, r) in runs.iter().enumerate() {
            assert_eq!(builder.index_of("a", Some(r)), Some(i));
        }
        let elapsed = start.elapsed();
        assert_eq!(builder.len(), N);
        assert!(elapsed < std::time::Duration::from_secs(2), "{elapsed:?}");
    }

    #[test]
    fn negative_and_positive_zero_runs_are_one_entry() {
        let mut builder = SharedStringTableBuilder::new();
        let i = builder.insert_rich("a", sized_run(-0.0, -0.0));
        assert_eq!(builder.insert_rich("a", sized_run(0.0, 0.0)), i);
        assert_eq!(builder.index_of("a", Some(&sized_run(0.0, -0.0))), Some(i));
        assert_eq!(builder.len(), 1);
    }

    #[test]
    fn expansion_budget_allows_the_limit_and_refuses_one_byte_more() {
        let mut budget = ExpansionBudget::new(10);
        budget.spend(4).unwrap();
        budget.spend(6).unwrap();
        let err = budget.spend(1).unwrap_err();
        assert_eq!(err.code(), "SIZE_LIMIT");
        assert!(err.to_string().contains("10-byte limit"), "{err}");
        let mut none = ExpansionBudget::new(usize::MAX);
        none.spend(usize::MAX).unwrap();
        none.spend(1).unwrap();
    }

    /// `stored` has the limit to itself: copies spent do not count
    /// against it, nor stored bytes against the copies.
    #[test]
    fn expansion_budget_stores_up_to_the_limit_beside_the_copies() {
        let mut budget = ExpansionBudget::new(10);
        budget.spend(10).unwrap();
        budget.store(4).unwrap();
        budget.store(6).unwrap();
        let err = budget.store(1).unwrap_err();
        assert_eq!(err.code(), "SIZE_LIMIT");
        let msg = err.to_string();
        assert!(
            msg.contains(
                "the defined names, shared strings, styles, chart series and axes, threaded comments and worksheet lists take more memory"
            ),
            "{msg}"
        );
        assert!(msg.contains("10-byte limit"), "{msg}");
        assert_eq!(budget.spent, 10);
        let mut none = ExpansionBudget::new(usize::MAX);
        none.store(usize::MAX).unwrap();
        none.store(1).unwrap();
    }

    /// The floor never passes what the runs serialize to.
    #[test]
    fn runs_json_floor_is_at_most_the_serialized_length() {
        let runs = vec![
            RichTextRun::default(),
            RichTextRun {
                text: "a\"b".into(),
                font_name: Some("Calibri".into()),
                color: Some("FF0000".into()),
                vert_align: Some("superscript".into()),
                scheme: Some("minor".into()),
                underline_style: Some("double".into()),
                color_spec: Some(Color {
                    rgb: Some("FF00FF00".into()),
                    ..Color::default()
                }),
                ..RichTextRun::default()
            },
        ];
        for r in &runs {
            let one = std::slice::from_ref(r);
            assert!(runs_json_floor(one) <= serde_json::to_string(r).unwrap().len());
        }
        assert_eq!(runs_json_floor(&runs[..1]), 11);
        assert!(runs_json_floor(&runs) <= serde_json::to_string(&runs).unwrap().len());
    }
}
