//! The note shapes of a sheet's legacy VML drawing
//! (`xl/drawings/vmlDrawing{n}.vml`; VML is the legacy drawing format of
//! ECMA-376 Part 4, Transitional).
//!
//! Excel shows a note (a legacy comment, `xl/comments{n}.xml`) only through
//! a shape in the VML drawing the sheet names by `<legacyDrawing>`, and it
//! refuses a sheet with notes and a `<drawing>` (a chart or a picture) but no
//! such drawing. [`note_drawing`] writes the drawing for a sheet that has
//! none: one hidden note shape per commented cell, in the markup Microsoft
//! Excel 16.0 saved on 2026-11-12. [`merge_note_drawing`] brings a preserved
//! drawing in line with the sheet's comments: it keeps every other shape and
//! every note shape whose cell still has a comment as they are, drops the
//! note shapes of cells without one, and appends a generated shape for each
//! commented cell it lacks.
//!
//! One write generates at most [`MAX_GENERATED_NOTE_SHAPES`] notes
//! ([`count_generated_notes`]): past it the write fails with `SizeLimit`.

use core::hint::cold_path;
use std::collections::HashSet;
use std::ops::Range;

use log::warn;
use quick_xml::Reader;
use quick_xml::events::{BytesStart, Event};

use super::cell::{CellRef, MAX_COL, MAX_ROW};
use super::comments::Comment;
use super::xml::{attr_str, attrs, root_namespaces};
use crate::{ModernXlsxError, Result};

/// The prefixes and namespaces the generated markup uses: VML, Office and
/// Excel, as Excel declares them on the drawing's root.
const NAMESPACES: [(&str, &str); 3] = [
    ("v", "urn:schemas-microsoft-com:vml"),
    ("o", "urn:schemas-microsoft-com:office:office"),
    ("x", "urn:schemas-microsoft-com:office:excel"),
];

/// A generated drawing's root start tag and shape layout, up to the
/// `o:idmap` `data` value.
const ROOT_OPEN: &str = concat!(
    r#"<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel">"#,
    r#"<o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data=""#,
);

/// The id of the text-box shape type every note shape is of.
const NOTE_SHAPE_TYPE_ID: &str = "_x0000_t202";

/// The note shape type, as Excel declares it once per drawing.
const NOTE_SHAPE_TYPE: &str = concat!(
    r#"<v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe">"#,
    r#"<v:stroke joinstyle="miter"/><v:path gradientshapeok="t" o:connecttype="rect"/></v:shapetype>"#,
);

/// The prefix of a shape id (`id` or `o:spid`): `_x0000_s` and its number.
const SHAPE_ID_PREFIX: &str = "_x0000_s";

/// The shape numbers an `o:idmap` block covers: block `b` covers
/// `1024 * b` to `1024 * b + 1023`, and Excel numbers a sheet's first shape
/// `1024 * b + 1`.
const ID_BLOCK: u32 = 1024;

/// Pixels per column and per row of the classic default grid (Calibri 11:
/// 8.43 characters, 15 pt), from which a generated shape's `margin-left` and
/// `margin-top` are worked out of its anchor. Excel works them out of the
/// sheet's actual widths and heights, so they can differ by a few points
/// from what Excel would write.
const COL_PX: u64 = 64;
const ROW_PX: u64 = 20;

/// The most notes one write generates across the workbook. A unit is a
/// note shape the writer adds for a legacy comment (in a generated drawing
/// or merged into a preserved one, about 640 bytes of VML) or a thread
/// placeholder (a `tc=` legacy comment, about 430 bytes of comments XML,
/// counted once with the note shape it gets: about 1,070 bytes together).
/// A preserved drawing written as it is, and the comments and shapes the
/// model gives, are not counted, so a workbook Excel saved is written back
/// whatever its number of notes. Measured on 2026-11-12 (release build,
/// peak heap growth of the write): 131,072 bare thread roots on distinct
/// cells wrote in 0.46 s with a 243 MiB peak natively and in 0.73 s with
/// 338 MiB of linear memory in the release WASM build (Node 26.13), and
/// the package reads back. The ceiling is the largest power of two whose
/// output reads back: a read keeps the generated note drawing as a
/// preserved entry, which the read JSON carries as an array of byte
/// values, so past about 207,000 generated notes reading the package
/// fails with `SizeLimit` at `max_json_length` (536,870,888 bytes, its
/// default and ceiling), although 262,144 wrote in 0.95 s with a 486 MiB
/// peak. Each unit costs a few dozen bytes of XML in the input (a few
/// bytes compressed), so without the ceiling a 6 MB package of 2,000,000
/// roots wrote for 10.6 s with a 3.66 GiB peak and trapped the WASM build.
/// Past it the write fails (`SizeLimit`).
pub const MAX_GENERATED_NOTE_SHAPES: usize = 1 << 17;

/// Adds `count` generated notes to `generated`, the running total of the
/// write (see [`MAX_GENERATED_NOTE_SHAPES`]), before they are built.
///
/// # Errors
///
/// `SizeLimit` when the total would pass [`MAX_GENERATED_NOTE_SHAPES`];
/// `generated` is then left as it was.
pub(crate) fn count_generated_notes(generated: &mut usize, count: usize) -> Result<()> {
    let total = generated.saturating_add(count);
    if total > MAX_GENERATED_NOTE_SHAPES {
        cold_path();
        return Err(ModernXlsxError::SizeLimit(format!(
            "Write failed: the notes the writer generates (a note shape for each legacy comment its VML drawing lacks, a tc= placeholder for each thread root without a legacy comment) reach {total} in one workbook, more than {MAX_GENERATED_NOTE_SHAPES} — write fewer comments per workbook; notes kept in the VML drawing they were read with are not counted"
        )));
    }
    *generated = total;
    Ok(())
}

/// The distinct cells of `comments`, in comment order. A comment whose
/// reference is not an A1 cell gets no shape (a warning is logged).
///
/// Each new cell among the first `counted` comments is a note shape the
/// writer generates: it is added to `generated` ([`count_generated_notes`])
/// as it is found, so a sheet past the ceiling stops there.
///
/// # Errors
///
/// `SizeLimit` when a counted cell would bring `generated` past
/// [`MAX_GENERATED_NOTE_SHAPES`].
pub(crate) fn note_cells(
    comments: &[Comment],
    counted: usize,
    generated: &mut usize,
) -> Result<Vec<CellRef>> {
    // Room for every cell up to the ceiling; a sheet with more (its notes
    // kept in a preserved drawing) grows past it.
    let room = comments.len().min(MAX_GENERATED_NOTE_SHAPES);
    let mut seen = HashSet::with_capacity(room);
    let mut cells = Vec::with_capacity(room);
    for (i, comment) in comments.iter().enumerate() {
        match CellRef::parse(&comment.cell_ref) {
            Ok(cell) => {
                if seen.insert(cell) {
                    if i < counted {
                        count_generated_notes(generated, 1)?;
                    }
                    cells.push(cell);
                }
            }
            Err(e) => {
                cold_path();
                warn!(
                    "comment on '{}': {e}; it gets no note shape",
                    comment.cell_ref
                );
            }
        }
    }
    Ok(cells)
}

/// The `x:Anchor` Excel gives a new note on `cell` on a sheet of default
/// column widths and row heights: left column, left offset, top row, top
/// offset, right column, right offset, bottom row, bottom offset (offsets in
/// pixels). The box sits one column right of the cell, or left of it within
/// the last three columns, and starts one row above it, from row 0 for the
/// first row and ending in the cell's row for the last three rows. Excel
/// 16.0 wrote exactly these anchors on 2026-11-12 for every cell the tests
/// name; in the last three rows its offsets vary by up to 2 pixels with the
/// row, and these are the ones it wrote for the last row.
pub(crate) fn note_anchor(cell: CellRef) -> [u32; 8] {
    let (c, r) = (cell.col, cell.row);
    let [left, left_off, right, right_off] = if c + 3 <= MAX_COL {
        [c + 1, 10, c + 3, 20]
    } else {
        [c.saturating_sub(3), 43, c.saturating_sub(1), 54]
    };
    let [top, top_off, bottom, bottom_off] = if r == 0 {
        [0, 1, 4, 6]
    } else if r + 3 <= MAX_ROW {
        [r - 1, 12, r + 3, 17]
    } else {
        [r.saturating_sub(4), 0, r, 5]
    };
    [
        left, left_off, top, top_off, right, right_off, bottom, bottom_off,
    ]
}

/// Appends `n` in decimal.
fn push_num(out: &mut String, n: u64) {
    out.push_str(itoa::Buffer::new().format(n));
}

/// Appends `px` pixels as points (`0.75 pt` each) with the unit: `55.5pt`.
fn push_pt(out: &mut String, px: u64) {
    let quarters = px * 3;
    push_num(out, quarters / 4);
    out.push_str(match quarters % 4 {
        1 => ".25pt",
        2 => ".5pt",
        3 => ".75pt",
        _ => "pt",
    });
}

/// Appends a hidden note shape for `cell`, numbered `id` and stacked at
/// `z_index`, in the markup Excel writes for a note that is not shown.
fn push_note_shape(out: &mut String, id: u32, z_index: u64, cell: CellRef) {
    let anchor = note_anchor(cell);
    out.push_str(r#"<v:shape id="_x0000_s"#);
    push_num(out, id.into());
    out.push_str(r##"" type="#_x0000_t202" style="position:absolute;margin-left:"##);
    push_pt(out, u64::from(anchor[0]) * COL_PX + u64::from(anchor[1]));
    out.push_str(";margin-top:");
    push_pt(out, u64::from(anchor[2]) * ROW_PX + u64::from(anchor[3]));
    out.push_str(";width:104pt;height:61.5pt;z-index:");
    push_num(out, z_index);
    out.push_str(concat!(
        r#";visibility:hidden" fillcolor="infoBackground [80]" strokecolor="none [81]" o:insetmode="auto">"#,
        r#"<v:fill color2="infoBackground [80]"/><v:shadow color="none [81]" obscured="t"/><v:path o:connecttype="none"/>"#,
        r#"<v:textbox style="mso-direction-alt:auto"><div style="text-align:left"></div></v:textbox>"#,
        r#"<x:ClientData ObjectType="Note"><x:MoveWithCells/><x:SizeWithCells/><x:Anchor>"#,
    ));
    for (i, n) in anchor.into_iter().enumerate() {
        if i > 0 {
            out.push_str(", ");
        }
        push_num(out, n.into());
    }
    out.push_str("</x:Anchor><x:AutoFill>False</x:AutoFill><x:Row>");
    push_num(out, cell.row.into());
    out.push_str("</x:Row><x:Column>");
    push_num(out, cell.col.into());
    out.push_str("</x:Column></x:ClientData></v:shape>");
}

/// A VML drawing holding a hidden note shape for each of `cells` (one or
/// more), numbered from the `o:idmap` block `first_block` up (Excel gives
/// sheet `n` block `n`), and the first block it leaves free.
pub(crate) fn note_drawing(cells: &[CellRef], first_block: u32) -> (Vec<u8>, u32) {
    let first_id = first_block.saturating_mul(ID_BLOCK).saturating_add(1);
    let count = u32::try_from(cells.len()).unwrap_or(u32::MAX);
    let last_block = first_id.saturating_add(count.saturating_sub(1)) / ID_BLOCK;
    let mut out = String::with_capacity(ROOT_OPEN.len() + 300 + cells.len() * 760);
    out.push_str(ROOT_OPEN);
    for block in first_block..=last_block {
        if block > first_block {
            out.push(',');
        }
        push_num(&mut out, block.into());
    }
    out.push_str(r#""/></o:shapelayout>"#);
    out.push_str(NOTE_SHAPE_TYPE);
    for (z_index, &cell) in (1u64..).zip(cells) {
        let id = first_id.saturating_add((z_index - 1) as u32);
        push_note_shape(&mut out, id, z_index, cell);
    }
    out.push_str("</xml>");
    (out.into_bytes(), last_block.saturating_add(1))
}

/// What a shape of a preserved drawing is.
#[derive(Clone, Copy)]
enum ShapeKind {
    /// A note (`x:ClientData ObjectType="Note"`), with its cell when it names
    /// one by `x:Row` and `x:Column`.
    Note(Option<CellRef>),
    /// Any other shape: a form control, a picture.
    Other,
}

/// A direct child of the drawing's root that is a shape.
struct Shape {
    /// Its bytes in the drawing.
    span: Range<usize>,
    kind: ShapeKind,
}

/// What [`merge_note_drawing`] needs to know of a preserved drawing.
struct ScannedDrawing {
    /// The root start tag, `<` to `>`.
    root_tag: Range<usize>,
    /// The offset of the root end tag.
    root_close: usize,
    /// The root's namespace declarations, as `(prefix, uri)`.
    namespaces: Vec<(String, String)>,
    shapes: Vec<Shape>,
    /// Whether it declares the note shape type ([`NOTE_SHAPE_TYPE_ID`]).
    has_note_type: bool,
    /// The highest `_x0000_sNNNN` shape number in it, 0 for none.
    max_shape_id: u32,
}

/// The `x:Row` or `x:Column` element whose text is being read.
#[derive(Clone, Copy)]
enum CellField {
    Row,
    Column,
}

/// Records in `scan` what the attributes of `e` say: a shape number in an
/// `id` or `o:spid`, and the note shape type.
fn scan_attrs(e: &BytesStart<'_>, scan: &mut ScannedDrawing) -> Result<()> {
    for attr in attrs(e) {
        let attr = attr?;
        if !matches!(attr.key.local_name().as_ref(), "id" | "spid") {
            continue;
        }
        let value = attr_str(&attr)?;
        if let Some(n) = value
            .strip_prefix(SHAPE_ID_PREFIX)
            .and_then(|n| n.parse::<u32>().ok())
        {
            scan.max_shape_id = scan.max_shape_id.max(n);
        } else if value == NOTE_SHAPE_TYPE_ID && e.local_name().as_ref() == "shapetype" {
            scan.has_note_type = true;
        }
    }
    Ok(())
}

/// Whether `e` is an `x:ClientData` of `ObjectType="Note"`.
fn is_note_client_data(e: &BytesStart<'_>) -> Result<bool> {
    if e.local_name().as_ref() != "ClientData" {
        return Ok(false);
    }
    for attr in attrs(e) {
        let attr = attr?;
        if attr.key.local_name().as_ref() == "ObjectType" {
            return Ok(attr_str(&attr)? == "Note");
        }
    }
    Ok(false)
}

/// Reads a preserved VML drawing in one forward pass with the quick-xml
/// reader: its root, the direct children of the root that are shapes (local
/// name `shape`), which of them are notes and on which cell, the note shape
/// type and the highest shape number. A leading U+FEFF is refused: quick-xml
/// drops it without counting it in `buffer_position()`, so the offsets would
/// miss its three bytes. The caller strips one byte order mark first, so a
/// second one is refused here.
///
/// Outside the root only ASCII whitespace, the XML declaration, comments,
/// processing instructions and a document type declaration are accepted
/// (they stay where they are when the drawing is merged).
///
/// # Errors
///
/// `XmlParse` when `xml` starts with U+FEFF, when the drawing is not
/// well-formed, including an element still open at its end, an element,
/// non-whitespace text, a character or entity reference or a CDATA section
/// outside the root, or when its root is an empty element.
fn scan_drawing(xml: &str) -> Result<ScannedDrawing> {
    let parse_err = |e: &dyn std::fmt::Display| {
        cold_path();
        ModernXlsxError::XmlParse(e.to_string())
    };
    if xml.starts_with('\u{feff}') {
        return Err(parse_err(&"content outside the root element"));
    }
    let mut reader = Reader::from_str(xml);
    let offset = |reader: &Reader<&[u8]>| {
        usize::try_from(reader.buffer_position()).map_err(|e| parse_err(&e))
    };
    let mut scan = ScannedDrawing {
        root_tag: 0..0,
        root_close: 0,
        namespaces: Vec::new(),
        shapes: Vec::new(),
        has_note_type: false,
        max_shape_id: 0,
    };
    let mut depth = 0usize;
    // The shape being read: its start, whether it is a note, its row and column.
    let mut shape: Option<(usize, bool, Option<u32>, Option<u32>)> = None;
    let mut field: Option<CellField> = None;
    loop {
        let start = offset(&reader)?;
        let event = reader.read_event().map_err(|e| parse_err(&e))?;
        if depth == 0 {
            // Only markup that holds no shape may stand outside the root: a
            // second root would move the root the shape offsets lie in.
            let outside = match &event {
                Event::Start(_) | Event::Empty(_) => !scan.root_tag.is_empty(),
                Event::Text(text) => !text.trim_ascii().is_empty(),
                Event::CData(_) | Event::GeneralRef(_) => true,
                _ => false,
            };
            if outside {
                return Err(parse_err(&"content outside the root element"));
            }
        }
        match event {
            Event::Start(e) => {
                scan_attrs(&e, &mut scan)?;
                match depth {
                    0 => {
                        scan.root_tag = start..offset(&reader)?;
                        scan.namespaces = root_namespaces(&e)?;
                    }
                    1 if e.local_name().as_ref() == "shape" => {
                        shape = Some((start, false, None, None));
                    }
                    _ => {
                        if let Some((_, note, _, _)) = shape.as_mut() {
                            *note |= is_note_client_data(&e)?;
                            field = match e.local_name().as_ref() {
                                "Row" => Some(CellField::Row),
                                "Column" => Some(CellField::Column),
                                _ => None,
                            };
                        }
                    }
                }
                depth += 1;
            }
            Event::Empty(e) => {
                scan_attrs(&e, &mut scan)?;
                match depth {
                    0 => return Err(parse_err(&"the root element is empty")),
                    1 if e.local_name().as_ref() == "shape" => scan.shapes.push(Shape {
                        span: start..offset(&reader)?,
                        kind: ShapeKind::Other,
                    }),
                    _ => {
                        if let Some((_, note, _, _)) = shape.as_mut() {
                            *note |= is_note_client_data(&e)?;
                        }
                    }
                }
            }
            Event::Text(text) => {
                if let (Some(field), Some((_, _, row, col))) = (field, shape.as_mut()) {
                    let value = text.trim().parse::<u32>().ok();
                    match field {
                        CellField::Row => *row = value,
                        CellField::Column => *col = value,
                    }
                }
            }
            Event::End(_) => {
                // The reader checks end names, so every End closes an open element.
                field = None;
                depth = depth.saturating_sub(1);
                match depth {
                    0 => scan.root_close = start,
                    1 => {
                        if let Some((from, note, row, col)) = shape.take() {
                            let cell = row.zip(col).map(|(row, col)| CellRef { col, row });
                            scan.shapes.push(Shape {
                                span: from..offset(&reader)?,
                                kind: if note {
                                    ShapeKind::Note(cell)
                                } else {
                                    ShapeKind::Other
                                },
                            });
                        }
                    }
                    _ => {}
                }
            }
            Event::Eof if depth != 0 => {
                return Err(parse_err(&"unclosed element at end of part"));
            }
            Event::Eof => break,
            _ => {}
        }
    }
    if scan.root_tag.is_empty() {
        return Err(parse_err(&"no root element"));
    }
    Ok(scan)
}

/// The preserved VML drawing at `path` brought in line with the sheet's
/// note `cells`, or `None` when it is written as it is.
///
/// When the cells of its note shapes are the sheet's commented cells, it is
/// kept byte for byte (`None`): it carries the boxes' sizes, positions and
/// formatting. Otherwise everything outside its note shapes (form controls,
/// shape types, the shape layout, text between shapes) and each note shape
/// whose cell still has a comment are kept as they are, each note shape on a
/// cell without one is dropped, and a hidden note shape is appended for each
/// commented cell it lacks, numbered above the highest shape number in it,
/// after the note shape type when it does not declare it. The VML, Office
/// and Excel namespaces the generated shapes use are declared on its root
/// when it lacks them.
///
/// A leading byte order mark stays in front. A drawing that is not UTF-8,
/// is not well-formed XML (see [`scan_drawing`]) or binds one of those
/// prefixes to another namespace is kept as it is, with a warning: no
/// drawing fails the write over its content or makes it panic.
///
/// The last `placeholders` of `cells` are the cells of the sheet's thread
/// placeholders, which the caller has counted already; the shapes appended
/// for the other cells are added to `generated` ([`count_generated_notes`])
/// before the merged drawing is built. A drawing kept as it is counts
/// nothing.
///
/// # Errors
///
/// `SizeLimit` when the appended shapes would bring `generated` past
/// [`MAX_GENERATED_NOTE_SHAPES`].
pub(crate) fn merge_note_drawing(
    preserved: &[u8],
    path: &str,
    cells: &[CellRef],
    placeholders: usize,
    generated: &mut usize,
) -> Result<Option<Vec<u8>>> {
    let kept_as_is = |e: &dyn std::fmt::Display| {
        cold_path();
        warn!("{path}: {e}; the note drawing is written as it is");
    };
    let Ok(xml) = std::str::from_utf8(preserved) else {
        kept_as_is(&"not valid UTF-8");
        return Ok(None);
    };
    // The offsets the scan takes leave out a BOM (see `scan_drawing`).
    let (bom, xml) = match xml.strip_prefix('\u{feff}') {
        Some(rest) => ("\u{feff}", rest),
        None => ("", xml),
    };
    let scan = match scan_drawing(xml) {
        Ok(scan) => scan,
        Err(e) => {
            kept_as_is(&e);
            return Ok(None);
        }
    };
    let wanted: HashSet<CellRef> = cells.iter().copied().collect();
    // The cells of its note shapes, and whether a note shape names none.
    let mut present: HashSet<CellRef> = HashSet::new();
    let mut cellless = false;
    for shape in &scan.shapes {
        match shape.kind {
            ShapeKind::Note(Some(cell)) => {
                present.insert(cell);
            }
            ShapeKind::Note(None) => cellless = true,
            ShapeKind::Other => {}
        }
    }
    if !cellless && present == wanted {
        return Ok(None);
    }

    let mut missing_namespaces = String::new();
    for (prefix, uri) in NAMESPACES {
        match scan.namespaces.iter().find(|(p, _)| p == prefix) {
            Some((_, declared)) if declared == uri => {}
            Some(_) => {
                kept_as_is(&format_args!(
                    "prefix {prefix} is bound to another namespace"
                ));
                return Ok(None);
            }
            None => {
                missing_namespaces.push_str(" xmlns:");
                missing_namespaces.push_str(prefix);
                missing_namespaces.push_str("=\"");
                missing_namespaces.push_str(uri);
                missing_namespaces.push('"');
            }
        }
    }

    // The shapes `splice` appends: one per cell without a note shape here.
    let added = cells[..cells.len().saturating_sub(placeholders)]
        .iter()
        .filter(|cell| !present.contains(cell))
        .count();
    count_generated_notes(generated, added)?;
    let merged = splice(xml, bom, &scan, &wanted, cells, &missing_namespaces);
    if merged.is_none() {
        kept_as_is(&"its shape offsets do not fit it");
        *generated -= added;
    }
    Ok(merged)
}

/// The merge [`merge_note_drawing`] describes, of `xml` (a drawing without
/// its byte order mark `bom`) as `scan` read it. Every piece of `xml` is
/// taken with `str::get`, so offsets that are out of order, past its end or
/// inside a character give `None` instead of a panic; the scan of a
/// well-formed drawing gives none of those.
fn splice(
    xml: &str,
    bom: &str,
    scan: &ScannedDrawing,
    wanted: &HashSet<CellRef>,
    cells: &[CellRef],
    missing_namespaces: &str,
) -> Option<Vec<u8>> {
    let mut out =
        String::with_capacity(bom.len() + xml.len() + NOTE_SHAPE_TYPE.len() + cells.len() * 760);
    out.push_str(bom);
    // The root start tag ends with `>`: the declarations go before it.
    let root_end = scan.root_tag.end.checked_sub(1)?;
    out.push_str(xml.get(..root_end)?);
    out.push_str(missing_namespaces);
    let mut copied = root_end;
    let mut have: HashSet<CellRef> = HashSet::with_capacity(cells.len());
    let mut kept_shapes = 0u64;
    for shape in &scan.shapes {
        match shape.kind {
            ShapeKind::Note(Some(cell)) if wanted.contains(&cell) && have.insert(cell) => {}
            ShapeKind::Note(_) => {
                out.push_str(xml.get(copied..shape.span.start)?);
                copied = shape.span.end;
                continue;
            }
            ShapeKind::Other => {}
        }
        kept_shapes += 1;
    }
    out.push_str(xml.get(copied..scan.root_close)?);
    if !scan.has_note_type {
        out.push_str(NOTE_SHAPE_TYPE);
    }
    let mut id = scan.max_shape_id.max(ID_BLOCK);
    for &cell in cells.iter().filter(|cell| !have.contains(cell)) {
        id = id.saturating_add(1);
        kept_shapes += 1;
        push_note_shape(&mut out, id, kept_shapes, cell);
    }
    out.push_str(xml.get(scan.root_close..)?);
    Some(out.into_bytes())
}

#[cfg(test)]
mod tests {
    use super::*;

    fn cell(a1: &str) -> CellRef {
        CellRef::parse(a1).unwrap()
    }

    /// `merge_note_drawing` of `preserved` at `d.vml` for `cells`, none of
    /// them a thread placeholder's, far under the ceiling.
    fn merge(preserved: &[u8], cells: &[CellRef]) -> Option<Vec<u8>> {
        merge_note_drawing(preserved, "d.vml", cells, 0, &mut 0).unwrap()
    }

    #[test]
    fn anchors_clamp_at_the_first_row_and_the_last_rows_and_columns() {
        // Microsoft Excel 16.0, 2026-11-12, default grid.
        let cases: [(&str, [u32; 8]); 7] = [
            ("A1", [1, 10, 0, 1, 3, 20, 4, 6]),
            ("C5", [3, 10, 3, 12, 5, 20, 7, 17]),
            ("AB200", [28, 10, 198, 12, 30, 20, 202, 17]),
            ("A1048573", [1, 10, 1_048_571, 12, 3, 20, 1_048_575, 17]),
            ("XFA13", [16381, 10, 11, 12, 16383, 20, 15, 17]),
            ("XFB3", [16378, 43, 1, 12, 16380, 54, 5, 17]),
            (
                "XFD1048576",
                [16380, 43, 1_048_571, 0, 16382, 54, 1_048_575, 5],
            ),
        ];
        for (a1, anchor) in cases {
            assert_eq!(note_anchor(cell(a1)), anchor, "{a1}");
        }
    }

    #[test]
    fn points_are_written_in_quarters() {
        let mut out = String::new();
        for px in [0, 1, 2, 3, 74] {
            push_pt(&mut out, px);
            out.push(' ');
        }
        assert_eq!(out, "0pt 0.75pt 1.5pt 2.25pt 55.5pt ");
    }

    #[test]
    fn a_drawing_of_more_than_1023_notes_maps_every_block_it_uses() {
        let cells: Vec<CellRef> = (0..1100).map(|row| CellRef { col: 0, row }).collect();
        let (xml, next) = note_drawing(&cells, 1);
        let xml = String::from_utf8(xml).unwrap();
        assert!(xml.contains(r#"<o:idmap v:ext="edit" data="1,2"/>"#));
        assert!(xml.contains(r#"id="_x0000_s1025""#));
        assert!(xml.contains(r#"id="_x0000_s2124""#));
        assert_eq!(next, 3);
        let (_, next) = note_drawing(&cells[..1], 3);
        assert_eq!(next, 4);
    }

    #[test]
    fn duplicate_and_unparseable_comment_cells_get_no_extra_shape() {
        let comment = |cell_ref: &str| Comment {
            cell_ref: cell_ref.into(),
            author: String::new(),
            text: String::new(),
        };
        let comments = [comment("B2"), comment("$B$2"), comment("B2"), comment("A1")];
        let mut generated = 0;
        let cells = note_cells(&comments, comments.len(), &mut generated).unwrap();
        assert_eq!(cells, [cell("B2"), cell("A1")]);
        // One generated shape per distinct cell, counted among the first
        // `counted` comments only.
        assert_eq!(generated, 2);
        let mut generated = 0;
        note_cells(&comments, 3, &mut generated).unwrap();
        assert_eq!(generated, 1);
        // The cell that passes the ceiling stops the search.
        let mut generated = MAX_GENERATED_NOTE_SHAPES - 1;
        let e = note_cells(&comments, comments.len(), &mut generated).unwrap_err();
        assert_eq!(e.code(), "SIZE_LIMIT");
        assert_eq!(generated, MAX_GENERATED_NOTE_SHAPES);
    }

    #[test]
    fn merge_declares_missing_namespaces_and_keeps_foreign_bindings_as_is() {
        let bare = "<xml><o:shapelayout/></xml>";
        let merged = String::from_utf8(merge(bare.as_bytes(), &[cell("A1")]).unwrap()).unwrap();
        assert!(merged.starts_with(r#"<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel"><o:shapelayout/><v:shapetype id="_x0000_t202""#), "{merged}");
        assert!(merged.contains(r#"<v:shape id="_x0000_s1025""#), "{merged}");
        assert!(merged.ends_with("</v:shape></xml>"), "{merged}");

        let foreign = r#"<xml xmlns:v="urn:other"><v:shape/></xml>"#;
        assert!(merge(foreign.as_bytes(), &[cell("A1")]).is_none());
        let empty_root = "<xml/>";
        assert!(merge(empty_root.as_bytes(), &[cell("A1")]).is_none());
        assert!(merge(b"<xml>\xff</xml>", &[cell("A1")]).is_none());
    }

    #[test]
    fn a_merge_counts_the_shapes_it_appends_for_legacy_comments() {
        let vml = format!("{ROOT}{A1_NOTE}</xml>");
        // A1 has a shape, B2 is a legacy comment without one, C3 a thread
        // placeholder (counted by the caller).
        let cells = [cell("A1"), cell("B2"), cell("C3")];
        let mut generated = 0;
        let out = merge_note_drawing(vml.as_bytes(), "d.vml", &cells, 1, &mut generated)
            .unwrap()
            .unwrap();
        assert_eq!(generated, 1);
        assert_eq!(
            String::from_utf8(out).unwrap().matches("<v:shape ").count(),
            3
        );

        // Exactly at the ceiling the merge is built.
        let mut generated = MAX_GENERATED_NOTE_SHAPES - 1;
        assert!(
            merge_note_drawing(vml.as_bytes(), "d.vml", &cells, 1, &mut generated)
                .unwrap()
                .is_some()
        );
        assert_eq!(generated, MAX_GENERATED_NOTE_SHAPES);
        // One past it fails, and the total stays where it was.
        let e = merge_note_drawing(vml.as_bytes(), "d.vml", &cells, 1, &mut generated).unwrap_err();
        assert_eq!(e.code(), "SIZE_LIMIT");
        assert!(e.to_string().contains("131072"), "{e}");
        assert_eq!(generated, MAX_GENERATED_NOTE_SHAPES);
        // A drawing whose notes match, or one kept as it is, counts nothing.
        let matching = [cell("A1")];
        assert!(
            merge_note_drawing(vml.as_bytes(), "d.vml", &matching, 0, &mut generated)
                .unwrap()
                .is_none()
        );
        let foreign = r#"<xml xmlns:v="urn:other"><v:shape/></xml>"#;
        assert!(
            merge_note_drawing(foreign.as_bytes(), "d.vml", &cells, 0, &mut generated)
                .unwrap()
                .is_none()
        );
        assert_eq!(generated, MAX_GENERATED_NOTE_SHAPES);
    }

    #[test]
    fn a_note_shape_without_a_cell_is_dropped() {
        let vml = concat!(
            r#"<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel">"#,
            r#"<v:shape id="_x0000_s1030"><x:ClientData ObjectType="Note"><x:Anchor>0</x:Anchor></x:ClientData></v:shape>"#,
            r#"</xml>"#,
        );
        let merged = String::from_utf8(merge(vml.as_bytes(), &[cell("A1")]).unwrap()).unwrap();
        assert!(!merged.contains("_x0000_s1030\""), "{merged}");
        assert!(merged.contains(r#"<v:shape id="_x0000_s1031""#), "{merged}");
        assert!(merged.contains("z-index:1;"), "{merged}");
    }

    /// A root start tag declaring the VML, Office and Excel namespaces.
    const ROOT: &str = r#"<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel">"#;

    /// A note shape on A1, numbered 1030.
    const A1_NOTE: &str = r#"<v:shape id="_x0000_s1030"><x:ClientData ObjectType="Note"><x:Row>0</x:Row><x:Column>0</x:Column></x:ClientData></v:shape>"#;

    /// `merge_note_drawing` of `vml` for the cells `a1`, as text.
    fn merged(vml: &str, a1: &[&str]) -> Option<String> {
        let cells: Vec<CellRef> = a1.iter().map(|a1| cell(a1)).collect();
        merge(vml.as_bytes(), &cells).map(|out| String::from_utf8(out).unwrap())
    }

    #[test]
    fn content_after_the_root_keeps_the_drawing_as_it_is() {
        // A1's shape is dropped and B2's added, so each of these would merge
        // if what follows the root were not refused.
        for tail in [
            "<xml></xml>",
            r#"<xml><v:shape id="_x0000_s1031"/></xml>"#,
            "<xml/>",
            "<o:shapelayout/>",
            "<x:a>1</x:a>",
            "text",
            "&amp;",
            "<![CDATA[x]]>",
        ] {
            let vml = format!("{ROOT}{A1_NOTE}</xml>{tail}");
            assert_eq!(merged(&vml, &["B2"]), None, "{tail}");
        }
        // Text before the root is no more well-formed than after it.
        assert_eq!(merged(&format!("x{ROOT}{A1_NOTE}</xml>"), &["B2"]), None);
    }

    #[test]
    fn comments_instructions_and_whitespace_after_the_root_are_kept() {
        let tail = "\r\n<!-- saved by hand -->\n<?app keep?>\n";
        let vml = format!("{ROOT}{A1_NOTE}</xml>{tail}");
        let out = merged(&vml, &["B2"]).unwrap();
        assert!(out.starts_with(ROOT), "{out}");
        assert!(!out.contains(A1_NOTE), "{out}");
        assert!(out.contains(r#"<v:shape id="_x0000_s1031""#), "{out}");
        assert!(out.ends_with(&format!("</v:shape></xml>{tail}")), "{out}");
    }

    #[test]
    fn a_byte_order_mark_is_kept_in_front_of_a_merged_drawing() {
        let (vml, _) = note_drawing(&[cell("A1")], 1);
        let vml = format!("\u{feff}{}", String::from_utf8(vml).unwrap());
        // (b) Notes that match: kept byte for byte.
        assert_eq!(merged(&vml, &["A1"]), None);
        // (a) A note added: every preserved byte up to the root end tag, the
        // BOM first, then the new shape.
        let out = merged(&vml, &["A1", "B2"]).unwrap();
        let body = vml.strip_suffix("</xml>").unwrap();
        assert!(out.starts_with(body), "{out}");
        let added = &out[body.len()..];
        assert!(added.starts_with(r#"<v:shape id="_x0000_s1026""#), "{out}");
        assert!(added.ends_with("</v:shape></xml>"), "{out}");
        assert_eq!(out.matches("<v:shape ").count(), 2, "{out}");
    }

    #[test]
    fn a_second_byte_order_mark_keeps_the_drawing_as_it_is() {
        // quick-xml drops a second U+FEFF uncounted, so its offsets would be
        // three bytes short.
        let vml = format!("\u{feff}\u{feff}{ROOT}{A1_NOTE}</xml>");
        assert_eq!(merged(&vml, &["B2"]), None);
    }

    #[test]
    fn non_ascii_text_beside_a_dropped_note_is_kept() {
        let check_box = r##"<v:shape id="_x0000_s1031" type="#_x0000_t201" alt="日本語 😀"><v:textbox><div>チェック 😀</div></v:textbox><x:ClientData ObjectType="Checkbox"/></v:shape>"##;
        for bom in ["", "\u{feff}"] {
            let vml = format!("{bom}{ROOT}{A1_NOTE}{check_box}</xml>");
            let out = merged(&vml, &["B2"]).unwrap();
            assert!(out.starts_with(&format!("{bom}{ROOT}{check_box}")), "{out}");
            assert!(!out.contains("_x0000_s1030"), "{out}");
            assert!(out.contains(r#"<v:shape id="_x0000_s1032""#), "{out}");
        }
    }

    #[test]
    fn no_truncated_or_mangled_drawing_makes_the_merge_panic() {
        // A drawing laid out as Excel lays it out (a BOM, the shape layout,
        // the note shape type, notes on A1 and B2, a check box with non-ASCII
        // text) without the boxes' styles and anchors, which keeps the sweep
        // short; A1's note is kept, B2's dropped and C3's added.
        let note = |id: u32, row: u32, col: u32| {
            format!(
                r##"<v:shape id="_x0000_s{id}" type="#_x0000_t202"><x:ClientData ObjectType="Note"><x:Row>{row}</x:Row><x:Column>{col}</x:Column></x:ClientData></v:shape>"##
            )
        };
        let check_box = r##"<v:shape id="_x0000_s1027" type="#_x0000_t201" alt="日本 😀"><v:textbox><div>チェック 😀</div></v:textbox><x:ClientData ObjectType="Checkbox"/></v:shape>"##;
        let vml = format!(
            r#"{}{ROOT}<o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="1"/></o:shapelayout>{NOTE_SHAPE_TYPE}{}{}{check_box}</xml>"#,
            '\u{feff}',
            note(1025, 0, 0),
            note(1026, 1, 1),
        );
        let cells = [cell("A1"), cell("C3")];
        let check = |bytes: &[u8]| {
            if let Some(out) = merge(bytes, &cells) {
                assert!(std::str::from_utf8(&out).is_ok());
            }
        };
        let started = std::time::Instant::now();
        let bytes = vml.as_bytes();
        assert!(merge(bytes, &cells).is_some());
        for len in 0..=bytes.len() {
            check(&bytes[..len]);
        }
        let mut mangled = bytes.to_vec();
        for at in 0..bytes.len() {
            for b in [b'<', b'>', b'/', b'"', b'&', 0xFF] {
                mangled[at] = b;
                check(&mangled);
            }
            mangled[at] = bytes[at];
        }
        eprintln!("sweep over {} bytes: {:?}", bytes.len(), started.elapsed());
    }
}
