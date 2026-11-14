//! Threaded comments — `xl/threadedComments/threadedComment{n}.xml` and the
//! persons list `xl/persons/person.xml` (MS-XLSX threaded comments, namespace
//! `.../spreadsheetml/2018/threadedcomments`).
//!
//! The reader accepts both `dT` (written by Excel and by this writer) and `dt`
//! (written by modern-xlsx 1.2.x) for the timestamp attribute, and keeps its
//! value as the file holds it. The writer always emits `dT`, in Excel's form
//! (`excel_dt`).
//!
//! How Microsoft Excel 16.0 reads `dT` and orders a thread (one file per
//! change through COM, 2026-11-12): Excel writes `dT` in UTC with two
//! fractional digits and no zone (`2026-11-12T22:17:43.41`). It reads the
//! fraction digits as a whole number of hundredths (`.41` is 0.41 s, `.5`
//! is 0.05 s, `.051` is 0.51 s) and leaves the date unread when that number
//! is above 99 (any three-digit fraction from `.100`) or when the value
//! has an offset such as `+01:00`; a trailing `Z` is ignored. It shows a
//! thread's replies by that date, and replies with the same date (or none)
//! by `id`, so a thread whose dates Excel cannot read shows its replies in
//! the order of their random GUIDs.
//!
//! What Microsoft Excel 16.0 needs to open and show a thread (tried one file
//! per change on 2026-11-12):
//! - Every comment `id`, `personId` and `parentId`, and every person `id`, is
//!   a GUID in braces, `{XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX}`
//!   (`is_braced_guid`); Excel writes it in upper case and reads either
//!   case. Excel refused each file with a comment `id`, or a person `id`
//!   (with the `personId`s naming it), without braces, and silently dropped
//!   a reply whose `parentId` had none.
//! - The thread's cell has a legacy comment (`ooxml::comments`), the
//!   placeholder older Excel versions display, whose author is
//!   `tc={root id}`; without one Excel opens the file but hides the thread,
//!   and when the cell's legacy comment is an ordinary note Excel shows the
//!   note, not the thread. Its text is not checked: Excel shows a thread whose
//!   placeholder text is stale or free. The package writer adds the missing
//!   placeholders (`placeholder_comments`) to the comments part it writes,
//!   never to the model, each counted against the notes one write may
//!   generate (`vml::MAX_GENERATED_NOTE_SHAPES`).
//! - A person needs no `userId` or `providerId`.

use core::hint::cold_path;
use std::borrow::Cow;
use std::collections::{HashMap, HashSet};
use std::fmt::Write as _;

use chrono::{Datelike, NaiveDate, NaiveDateTime, NaiveTime, TimeDelta, Timelike};
use quick_xml::events::{BytesEnd, BytesStart, BytesText, Event};
use quick_xml::{Reader, Writer};
use serde::{Deserialize, Serialize};

use crate::{ModernXlsxError, Result};

use crate::ooxml::{write_xml_decl, xml_write_err};

use super::cell::CellRef;
use super::comments::Comment;
use super::shared_strings::ExpansionBudget;
use super::vml::count_generated_notes;
use super::xml::{attr_str, attrs, parse_attr, push_general_ref, xml_attr, xsd_bool};

const THREADED_COMMENTS_NS: &str =
    "http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments";

/// A single threaded comment (Microsoft 365 modern comments).
///
/// The part's `<extLst>` extension lists are not modelled and are dropped on
/// read.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ThreadedCommentData {
    /// Unique comment identifier: a GUID in braces, `{XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX}`
    /// (the form Excel requires, `is_braced_guid`).
    pub id: String,
    /// Cell reference, e.g. "A1".
    pub ref_cell: String,
    /// The author's [`PersonData::id`] (a GUID in braces).
    pub person_id: String,
    /// Plain text content of the comment.
    pub text: String,
    /// ISO 8601 timestamp of the comment (the `dT` attribute), as the file
    /// holds it after a read. The writer writes it in Excel's form
    /// (`excel_dt`) without changing this field.
    pub timestamp: String,
    /// The `id` of the thread root this comment replies to (a GUID in
    /// braces; Excel drops a reply whose `parentId` has none).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub parent_id: Option<String>,
    /// Whether the thread is resolved (the `done` attribute).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub done: Option<bool>,
    /// `@mentions` inside the text, in document order.
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub mentions: Vec<MentionData>,
}

/// One `@mention` of a person inside a threaded comment's text.
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct MentionData {
    /// Person identifier (GUID) of the mentioned person (XML attribute
    /// `mentionpersonId`, with a lowercase `p`).
    pub mention_person_id: String,
    /// Unique identifier (GUID) of this mention.
    pub mention_id: String,
    /// Index of the first mentioned character in the comment text.
    pub start_index: u32,
    /// Number of mentioned characters.
    pub length: u32,
}

/// A person entry from the persons list (xl/persons/person.xml).
///
/// The `<extLst>` extension list is not modelled and is dropped on read.
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct PersonData {
    /// Unique person identifier: a GUID in braces (the form Excel requires,
    /// `is_braced_guid`).
    pub id: String,
    /// Display name of the person.
    pub display_name: String,
    /// Optional identity-provider user identifier.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub user_id: Option<String>,
    /// Optional provider identifier (e.g. "Windows Live").
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub provider_id: Option<String>,
}

// ---------------------------------------------------------------------------
// What Excel requires
// ---------------------------------------------------------------------------

/// Whether `id` has the form Excel requires of a threaded comment's `id`,
/// `personId` and `parentId` and of a person's `id`: a GUID in braces,
/// `{XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX}`, its hex digits in either case.
pub(crate) fn is_braced_guid(id: &str) -> bool {
    let [b'{', guid @ .., b'}'] = id.as_bytes() else {
        return false;
    };
    guid.len() == 36
        && guid.iter().enumerate().all(|(i, &b)| {
            if matches!(i, 8 | 13 | 18 | 23) {
                b == b'-'
            } else {
                b.is_ascii_hexdigit()
            }
        })
}

/// The start of a placeholder's text, up to the root's text: what Excel 16.0
/// wrote on 2026-11-12, with its CRLF line ends as LF (as the reader reads
/// them).
const PLACEHOLDER_HEAD: &str = "[Threaded comment]\n\nYour version of Excel allows you to read this threaded comment; however, any edits to it will get removed if the file is opened in a newer version of Excel. Learn more: https://go.microsoft.com/fwlink/?linkid=870924\n\nComment:\n    ";

/// The legacy placeholder comments a sheet's comments part needs beside
/// `threads` when its legacy comments are `notes`: one for each thread root
/// (a comment without `parent_id`) whose cell has no legacy comment yet, the
/// first root on a cell winning. Its author is `tc={root id}` and its text
/// the one Excel writes: [`PLACEHOLDER_HEAD`], the root's text, then for
/// each reply in order (a comment whose `parent_id` is the root's id, ASCII
/// case ignored as Excel does) `Reply:` and its text, each heading on its
/// own line and each text after four spaces. A root whose `ref` is not an
/// A1 cell gets none. Excel needs the placeholder to show the thread; the
/// writer adds these to the part it writes and leaves `notes` as they are.
///
/// Each placeholder is added to `generated`, the running total of the notes
/// the write generates (`vml::count_generated_notes`), as its root is found
/// and before any placeholder is built.
///
/// # Errors
///
/// `SizeLimit` when a placeholder would bring `generated` past
/// `vml::MAX_GENERATED_NOTE_SHAPES`: the search stops at the first one past
/// it, and none is built.
pub(crate) fn placeholder_comments(
    threads: &[ThreadedCommentData],
    notes: &[Comment],
    generated: &mut usize,
) -> Result<Vec<Comment>> {
    if threads.is_empty() {
        return Ok(Vec::new());
    }
    let mut taken: HashSet<CellRef> = notes
        .iter()
        .filter_map(|note| CellRef::parse(&note.cell_ref).ok())
        .collect();
    let mut roots: Vec<&ThreadedCommentData> = Vec::new();
    for root in threads.iter().filter(|c| c.parent_id.is_none()) {
        if CellRef::parse(&root.ref_cell).is_ok_and(|cell| taken.insert(cell)) {
            count_generated_notes(generated, 1)?;
            roots.push(root);
        }
    }
    let mut placeholders: Vec<Comment> = Vec::with_capacity(roots.len());
    // Each placeholder's index, by its root id in ASCII upper case.
    let mut by_root: HashMap<String, usize> = HashMap::with_capacity(roots.len());
    for root in roots {
        by_root.insert(root.id.to_ascii_uppercase(), placeholders.len());
        placeholders.push(Comment {
            cell_ref: root.ref_cell.clone(),
            author: format!("tc={}", root.id),
            text: format!("{PLACEHOLDER_HEAD}{}", root.text),
        });
    }
    if by_root.is_empty() {
        return Ok(placeholders);
    }
    for reply in threads {
        if let Some(parent) = &reply.parent_id
            && let Some(&i) = by_root.get(&parent.to_ascii_uppercase())
            && let Some(placeholder) = placeholders.get_mut(i)
        {
            placeholder.text.push_str("\nReply:\n    ");
            placeholder.text.push_str(&reply.text);
        }
    }
    Ok(placeholders)
}

/// Read the attributes of a `<threadedComment>` start or empty tag; the text
/// and mentions come from its children.
fn comment_from_attrs(e: &BytesStart<'_>) -> Result<ThreadedCommentData> {
    let mut c = ThreadedCommentData {
        id: String::new(),
        ref_cell: String::new(),
        person_id: String::new(),
        text: String::new(),
        timestamp: String::new(),
        parent_id: None,
        done: None,
        mentions: Vec::new(),
    };
    for attr in attrs(e) {
        let attr = attr?;
        match attr.key.local_name().as_ref() {
            "id" => c.id = attr_str(&attr)?.into_owned(),
            "ref" => c.ref_cell = attr_str(&attr)?.into_owned(),
            "personId" => c.person_id = attr_str(&attr)?.into_owned(),
            // Excel writes `dT`; 1.2.x wrote `dt`, which is read as well.
            "dT" | "dt" => c.timestamp = attr_str(&attr)?.into_owned(),
            "parentId" => c.parent_id = Some(attr_str(&attr)?.into_owned()),
            "done" => c.done = xsd_bool(&attr_str(&attr)?),
            _ => {}
        }
    }
    Ok(c)
}

/// Read the attributes of a `<mention>` tag.
fn mention_from_attrs(e: &BytesStart<'_>) -> Result<MentionData> {
    let mut m = MentionData {
        mention_person_id: String::new(),
        mention_id: String::new(),
        start_index: 0,
        length: 0,
    };
    for attr in attrs(e) {
        let attr = attr?;
        match attr.key.local_name().as_ref() {
            "mentionpersonId" | "mentionPersonId" => {
                m.mention_person_id = attr_str(&attr)?.into_owned();
            }
            "mentionId" => m.mention_id = attr_str(&attr)?.into_owned(),
            "startIndex" => m.start_index = parse_attr("mention", &attr)?,
            "length" => m.length = parse_attr("mention", &attr)?,
            _ => {}
        }
    }
    Ok(m)
}

// ---------------------------------------------------------------------------
// Parsers
// ---------------------------------------------------------------------------

/// Parse a threaded comments XML file from raw bytes.
///
/// The XML structure expected is:
/// ```xml
/// <ThreadedComments xmlns="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments">
///   <threadedComment ref="A1" dT="2024-01-15T10:30:00.00" personId="{GUID}" id="{GUID}" done="0">
///     <text>This is a comment</text>
///     <mentions>
///       <mention mentionpersonId="{GUID}" mentionId="{GUID}" startIndex="0" length="4"/>
///     </mentions>
///   </threadedComment>
/// </ThreadedComments>
/// ```
///
/// A self-closing `<threadedComment .../>` yields a comment with empty text.
///
/// # Errors
///
/// `XmlParse` for malformed XML, an attribute holding a bad entity, or a
/// numeric attribute that is not a valid number.
pub fn parse_threaded_comments(data: &[u8]) -> Result<Vec<ThreadedCommentData>> {
    parse_threaded_comments_within(data, &mut ExpansionBudget::new(usize::MAX))
}

/// The memory a kept comment takes beside its mentions: its
/// [`ThreadedCommentData`] plus the capacity of its strings.
fn comment_memory(c: &ThreadedCommentData) -> usize {
    size_of::<ThreadedCommentData>()
        + c.id.capacity()
        + c.ref_cell.capacity()
        + c.person_id.capacity()
        + c.text.capacity()
        + c.timestamp.capacity()
        + c.parent_id.as_ref().map_or(0, String::capacity)
}

/// The memory a kept mention takes: its [`MentionData`] plus the capacity of
/// its strings.
fn mention_memory(m: &MentionData) -> usize {
    size_of::<MentionData>() + m.mention_person_id.capacity() + m.mention_id.capacity()
}

/// [`parse_threaded_comments`], counting in `budget` (with
/// [`ExpansionBudget::store`]) each comment and mention the part keeps before
/// it is kept: a mention when it is read, a comment (without its mentions)
/// when it ends.
///
/// # Errors
///
/// As [`parse_threaded_comments`], and `SizeLimit` once the count passes
/// `budget.store_limit`.
pub(crate) fn parse_threaded_comments_within(
    data: &[u8],
    budget: &mut ExpansionBudget,
) -> Result<Vec<ThreadedCommentData>> {
    let mut reader = Reader::from_reader(data);
    reader.config_mut().trim_text(false);

    let mut buf = Vec::with_capacity(512);
    let mut comments: Vec<ThreadedCommentData> = Vec::new();

    // The comment being read, and whether the reader is inside its `<text>`.
    let mut current: Option<ThreadedCommentData> = None;
    let mut in_text = false;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) => match e.local_name().as_ref() {
                "threadedComment" => current = Some(comment_from_attrs(e)?),
                "text" if current.is_some() => in_text = true,
                "mention" => {
                    if let Some(c) = current.as_mut() {
                        let mention = mention_from_attrs(e)?;
                        budget.store(mention_memory(&mention))?;
                        c.mentions.push(mention);
                    }
                }
                _ => {}
            },
            Ok(Event::Empty(ref e)) => match e.local_name().as_ref() {
                "threadedComment" => {
                    let comment = comment_from_attrs(e)?;
                    budget.store(comment_memory(&comment))?;
                    comments.push(comment);
                }
                "mention" => {
                    if let Some(c) = current.as_mut() {
                        let mention = mention_from_attrs(e)?;
                        budget.store(mention_memory(&mention))?;
                        c.mentions.push(mention);
                    }
                }
                _ => {}
            },
            Ok(Event::Text(ref e)) => {
                if in_text && let Some(c) = current.as_mut() {
                    c.text.push_str(&e.xml10_content());
                }
            }
            Ok(Event::CData(ref e)) => {
                if in_text && let Some(c) = current.as_mut() {
                    c.text.push_str(&e.xml10_content());
                }
            }
            Ok(Event::GeneralRef(ref e)) => {
                if in_text && let Some(c) = current.as_mut() {
                    push_general_ref(&mut c.text, e)?;
                }
            }
            Ok(Event::End(ref e)) => match e.local_name().as_ref() {
                "threadedComment" => {
                    if let Some(comment) = current.take() {
                        budget.store(comment_memory(&comment))?;
                        comments.push(comment);
                    }
                    in_text = false;
                }
                "text" => in_text = false,
                _ => {}
            },
            Ok(Event::Eof) => break,
            Ok(_) => {}
            Err(e) => {
                cold_path();
                return Err(ModernXlsxError::XmlParse(format!(
                    "error parsing threaded comments XML: {e}"
                )));
            }
        }
        buf.clear();
    }

    Ok(comments)
}

/// Parse a persons XML file from raw bytes.
///
/// The XML structure expected is:
/// ```xml
/// <personList xmlns="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments">
///   <person displayName="John Doe" id="{GUID}" userId="S::john@example.com::{GUID}" providerId="AD"/>
/// </personList>
/// ```
///
/// # Errors
///
/// `XmlParse` for malformed XML or an attribute holding a bad entity.
pub fn parse_persons(data: &[u8]) -> Result<Vec<PersonData>> {
    let mut reader = Reader::from_reader(data);
    let mut buf = Vec::with_capacity(256);
    let mut persons: Vec<PersonData> = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Empty(ref e) | Event::Start(ref e)) => {
                if e.local_name().as_ref() == "person" {
                    let mut id = String::new();
                    let mut display_name = String::new();
                    let mut user_id: Option<String> = None;
                    let mut provider_id: Option<String> = None;

                    for attr in attrs(e) {
                        let attr = attr?;
                        match attr.key.local_name().as_ref() {
                            "id" => id = attr_str(&attr)?.into_owned(),
                            "displayName" => display_name = attr_str(&attr)?.into_owned(),
                            "userId" => user_id = Some(attr_str(&attr)?.into_owned()),
                            "providerId" => provider_id = Some(attr_str(&attr)?.into_owned()),
                            _ => {}
                        }
                    }

                    if !id.is_empty() {
                        persons.push(PersonData {
                            id,
                            display_name,
                            user_id,
                            provider_id,
                        });
                    }
                }
            }
            Ok(Event::Eof) => break,
            Ok(_) => {}
            Err(e) => {
                cold_path();
                return Err(ModernXlsxError::XmlParse(format!(
                    "error parsing persons XML: {e}"
                )));
            }
        }
        buf.clear();
    }

    Ok(persons)
}

// ---------------------------------------------------------------------------
// Writers
// ---------------------------------------------------------------------------

/// `timestamp` in the form Excel writes and reads as a date for `dT`: UTC,
/// `YYYY-MM-DDTHH:MM:SS.ff` (hundredths, truncated) and no zone.
///
/// Converts an ISO 8601 date-time `YYYY-MM-DDTHH:MM:SS[.f…][Z|±HH:MM]`; one
/// without a zone is UTC, the zone `dT` is in. Anything else, Excel's own
/// form, and one whose offset moves the UTC instant out of years 0000-9999
/// (which `YYYY` cannot write) is returned as it is.
pub(crate) fn excel_dt(timestamp: &str) -> Cow<'_, str> {
    match parse_iso_utc(timestamp) {
        Some(utc) if (0..=9999).contains(&utc.year()) => {
            let mut out = String::with_capacity(22);
            // Writing to a String cannot fail.
            let _ = write!(
                out,
                "{:04}-{:02}-{:02}T{:02}:{:02}:{:02}.{:02}",
                utc.year(),
                utc.month(),
                utc.day(),
                utc.hour(),
                utc.minute(),
                utc.second(),
                utc.nanosecond() / 10_000_000,
            );
            if out == timestamp {
                Cow::Borrowed(timestamp)
            } else {
                Cow::Owned(out)
            }
        }
        _ => Cow::Borrowed(timestamp),
    }
}

/// The UTC date-time of `YYYY-MM-DDTHH:MM:SS[.f…][Z|±HH:MM]` (ASCII digits;
/// no zone is UTC), or `None` when `s` is not that or names no valid instant.
fn parse_iso_utc(s: &str) -> Option<NaiveDateTime> {
    /// The value of the ASCII digits `b[range]`, `None` if any is not one.
    fn num(b: &[u8], range: core::ops::Range<usize>) -> Option<u32> {
        b.get(range)?.iter().try_fold(0u32, |n, &d| {
            d.is_ascii_digit().then(|| n * 10 + u32::from(d - b'0'))
        })
    }
    let b = s.as_bytes();
    if b.len() < 19
        || [(4, b'-'), (7, b'-'), (10, b'T'), (13, b':'), (16, b':')]
            .iter()
            .any(|&(i, c)| b[i] != c)
    {
        return None;
    }
    let date = NaiveDate::from_ymd_opt(
        i32::try_from(num(b, 0..4)?).ok()?,
        num(b, 5..7)?,
        num(b, 8..10)?,
    )?;
    let mut rest = &b[19..];
    let mut nanos = 0u32;
    if let [b'.', tail @ ..] = rest {
        let digits = tail.iter().take_while(|d| d.is_ascii_digit()).count();
        if digits == 0 {
            return None;
        }
        // Nanoseconds from the first nine digits; later ones are below Excel's hundredths.
        nanos = tail[..digits]
            .iter()
            .zip((0..9).rev())
            .fold(0, |n, (&d, power)| {
                n + u32::from(d - b'0') * 10u32.pow(power)
            });
        rest = &tail[digits..];
    }
    // Rejects hour 24 and second 60.
    let time =
        NaiveTime::from_hms_nano_opt(num(b, 11..13)?, num(b, 14..16)?, num(b, 17..19)?, nanos)?;
    let local = date.and_time(time);
    let offset_minutes = match rest {
        [] | [b'Z'] => 0,
        [sign @ (b'+' | b'-'), _, _, b':', _, _] => {
            let (hours, minutes) = (num(rest, 1..3)?, num(rest, 4..6)?);
            if hours > 23 || minutes > 59 {
                return None;
            }
            let total = i64::from(hours * 60 + minutes);
            if *sign == b'-' { -total } else { total }
        }
        _ => return None,
    };
    local.checked_sub_signed(TimeDelta::try_minutes(offset_minutes)?)
}

/// Serialize threaded comments to XML bytes.
pub fn write_threaded_comments(comments: &[ThreadedCommentData]) -> Result<Vec<u8>> {
    let mut buf: Vec<u8> = Vec::with_capacity(512 + comments.len() * 256);
    let mut writer = Writer::new(&mut buf);
    let mut ibuf = itoa::Buffer::new();

    // XML declaration.
    write_xml_decl(&mut writer)?;

    // <ThreadedComments xmlns="...">
    let mut root = BytesStart::new("ThreadedComments");
    root.push_attribute(("xmlns", THREADED_COMMENTS_NS));
    writer
        .write_event(Event::Start(root))
        .map_err(xml_write_err)?;

    for comment in comments {
        let mut elem = BytesStart::new("threadedComment");
        elem.push_attribute(xml_attr("id", &comment.id));
        elem.push_attribute(xml_attr("ref", &comment.ref_cell));
        elem.push_attribute(xml_attr("personId", &comment.person_id));
        // Excel misreads ISO 8601 in other forms and then misorders the
        // thread's replies (module docs).
        elem.push_attribute(xml_attr("dT", &excel_dt(&comment.timestamp)));

        if let Some(ref parent_id) = comment.parent_id {
            elem.push_attribute(xml_attr("parentId", parent_id));
        }
        if let Some(done) = comment.done {
            elem.push_attribute(("done", if done { "1" } else { "0" }));
        }

        writer
            .write_event(Event::Start(elem))
            .map_err(xml_write_err)?;

        // <text>...</text>
        writer
            .write_event(Event::Start(BytesStart::new("text")))
            .map_err(xml_write_err)?;
        writer
            .write_event(Event::Text(BytesText::new(
                crate::ooxml::sanitize_xml_text(&comment.text).as_ref(),
            )))
            .map_err(xml_write_err)?;
        writer
            .write_event(Event::End(BytesEnd::new("text")))
            .map_err(xml_write_err)?;

        // <mentions><mention .../>...</mentions>
        if !comment.mentions.is_empty() {
            writer
                .write_event(Event::Start(BytesStart::new("mentions")))
                .map_err(xml_write_err)?;
            for mention in &comment.mentions {
                let mut m = BytesStart::new("mention");
                m.push_attribute(xml_attr("mentionpersonId", &mention.mention_person_id));
                m.push_attribute(xml_attr("mentionId", &mention.mention_id));
                m.push_attribute(("startIndex", ibuf.format(mention.start_index)));
                m.push_attribute(("length", ibuf.format(mention.length)));
                writer.write_event(Event::Empty(m)).map_err(xml_write_err)?;
            }
            writer
                .write_event(Event::End(BytesEnd::new("mentions")))
                .map_err(xml_write_err)?;
        }

        // </threadedComment>
        writer
            .write_event(Event::End(BytesEnd::new("threadedComment")))
            .map_err(xml_write_err)?;
    }

    // </ThreadedComments>
    writer
        .write_event(Event::End(BytesEnd::new("ThreadedComments")))
        .map_err(xml_write_err)?;

    Ok(buf)
}

/// Serialize persons to XML bytes.
pub fn write_persons(persons: &[PersonData]) -> Result<Vec<u8>> {
    let mut buf: Vec<u8> = Vec::with_capacity(256 + persons.len() * 128);
    let mut writer = Writer::new(&mut buf);

    // XML declaration.
    write_xml_decl(&mut writer)?;

    // <personList xmlns="...">
    let mut root = BytesStart::new("personList");
    root.push_attribute(("xmlns", THREADED_COMMENTS_NS));
    writer
        .write_event(Event::Start(root))
        .map_err(xml_write_err)?;

    for person in persons {
        let mut elem = BytesStart::new("person");
        elem.push_attribute(xml_attr("id", &person.id));
        elem.push_attribute(xml_attr("displayName", &person.display_name));

        if let Some(ref user_id) = person.user_id {
            elem.push_attribute(xml_attr("userId", user_id));
        }
        if let Some(ref provider_id) = person.provider_id {
            elem.push_attribute(xml_attr("providerId", provider_id));
        }

        writer
            .write_event(Event::Empty(elem))
            .map_err(xml_write_err)?;
    }

    // </personList>
    writer
        .write_event(Event::End(BytesEnd::new("personList")))
        .map_err(xml_write_err)?;

    Ok(buf)
}

#[cfg(test)]
mod tests {
    use super::*;
    use pretty_assertions::assert_eq;

    #[test]
    fn parse_threaded_comment() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<ThreadedComments xmlns="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments">
  <threadedComment id="{A0000000-0000-4000-8000-000000000001}" ref="A1" personId="{D0000000-0000-4000-8000-0000000000A1}" dT="2024-01-15T10:30:00.000">
    <text>This is a comment</text>
  </threadedComment>
</ThreadedComments>"#;

        let comments = parse_threaded_comments(xml.as_bytes()).unwrap();
        assert_eq!(comments.len(), 1);
        assert_eq!(comments[0].id, "{A0000000-0000-4000-8000-000000000001}");
        assert_eq!(comments[0].ref_cell, "A1");
        assert_eq!(
            comments[0].person_id,
            "{D0000000-0000-4000-8000-0000000000A1}"
        );
        assert_eq!(comments[0].text, "This is a comment");
        assert_eq!(comments[0].timestamp, "2024-01-15T10:30:00.000");
        assert!(comments[0].parent_id.is_none());
    }

    #[test]
    fn parse_reply_chain() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<ThreadedComments xmlns="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments">
  <threadedComment id="{A0000000-0000-4000-8000-000000000001}" ref="A1" personId="{D0000000-0000-4000-8000-0000000000A1}" dT="2024-01-15T10:30:00.000">
    <text>Original comment</text>
  </threadedComment>
  <threadedComment id="{B0000000-0000-4000-8000-000000000002}" ref="A1" personId="{D0000000-0000-4000-8000-0000000000A2}" dT="2024-01-15T11:00:00.000" parentId="{A0000000-0000-4000-8000-000000000001}">
    <text>This is a reply</text>
  </threadedComment>
  <threadedComment id="{C0000000-0000-4000-8000-000000000003}" ref="A1" personId="{D0000000-0000-4000-8000-0000000000A1}" dT="2024-01-15T11:30:00.000" parentId="{A0000000-0000-4000-8000-000000000001}">
    <text>Another reply</text>
  </threadedComment>
</ThreadedComments>"#;

        let comments = parse_threaded_comments(xml.as_bytes()).unwrap();
        assert_eq!(comments.len(), 3);

        assert!(comments[0].parent_id.is_none());
        assert_eq!(
            comments[1].parent_id.as_deref(),
            Some("{A0000000-0000-4000-8000-000000000001}")
        );
        assert_eq!(
            comments[2].parent_id.as_deref(),
            Some("{A0000000-0000-4000-8000-000000000001}")
        );

        assert_eq!(comments[0].text, "Original comment");
        assert_eq!(comments[1].text, "This is a reply");
        assert_eq!(comments[2].text, "Another reply");
    }

    #[test]
    fn parse_persons_basic() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<personList xmlns="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments">
  <person id="{D0000000-0000-4000-8000-0000000000A1}" displayName="John Doe" providerId="Windows Live"/>
  <person id="{D0000000-0000-4000-8000-0000000000A2}" displayName="Jane Smith"/>
</personList>"#;

        let persons = parse_persons(xml.as_bytes()).unwrap();
        assert_eq!(persons.len(), 2);

        assert_eq!(persons[0].id, "{D0000000-0000-4000-8000-0000000000A1}");
        assert_eq!(persons[0].display_name, "John Doe");
        assert_eq!(persons[0].provider_id.as_deref(), Some("Windows Live"));

        assert_eq!(persons[1].id, "{D0000000-0000-4000-8000-0000000000A2}");
        assert_eq!(persons[1].display_name, "Jane Smith");
        assert!(persons[1].provider_id.is_none());
    }

    #[test]
    fn threaded_comments_roundtrip() {
        let comments = vec![
            ThreadedCommentData {
                id: "{A0000000-0000-4000-8000-000000000001}".to_string(),
                ref_cell: "A1".to_string(),
                person_id: "{D0000000-0000-4000-8000-0000000000A1}".to_string(),
                text: "First comment".to_string(),
                timestamp: "2024-01-15T10:30:00.00".to_string(),
                parent_id: None,
                done: None,
                mentions: Vec::new(),
            },
            ThreadedCommentData {
                id: "{B0000000-0000-4000-8000-000000000002}".to_string(),
                ref_cell: "A1".to_string(),
                person_id: "{D0000000-0000-4000-8000-0000000000A2}".to_string(),
                text: "Reply with <special> & \"chars\"".to_string(),
                timestamp: "2024-01-15T11:00:00.00".to_string(),
                parent_id: Some("{A0000000-0000-4000-8000-000000000001}".to_string()),
                done: None,
                mentions: Vec::new(),
            },
        ];

        let xml = write_threaded_comments(&comments).unwrap();
        let parsed = parse_threaded_comments(&xml).unwrap();

        assert_eq!(parsed.len(), 2);
        for (orig, round) in comments.iter().zip(parsed.iter()) {
            assert_eq!(orig.id, round.id);
            assert_eq!(orig.ref_cell, round.ref_cell);
            assert_eq!(orig.person_id, round.person_id);
            assert_eq!(orig.text, round.text);
            assert_eq!(orig.timestamp, round.timestamp);
            assert_eq!(orig.parent_id, round.parent_id);
        }
    }

    #[test]
    fn persons_roundtrip() {
        let persons = vec![
            PersonData {
                id: "{D0000000-0000-4000-8000-0000000000A1}".to_string(),
                display_name: "Alice O'Reilly".to_string(),
                user_id: Some("S::alice@example.com::guid".to_string()),
                provider_id: Some("Windows Live".to_string()),
            },
            PersonData {
                id: "{D0000000-0000-4000-8000-0000000000A2}".to_string(),
                display_name: "Bob & Carol".to_string(),
                user_id: None,
                provider_id: None,
            },
        ];

        let xml = write_persons(&persons).unwrap();
        let parsed = parse_persons(&xml).unwrap();

        assert_eq!(parsed.len(), 2);
        for (orig, round) in persons.iter().zip(parsed.iter()) {
            assert_eq!(orig.id, round.id);
            assert_eq!(orig.display_name, round.display_name);
            assert_eq!(orig.provider_id, round.provider_id);
        }
    }

    #[test]
    fn write_threaded_comments_empty() {
        let comments: Vec<ThreadedCommentData> = Vec::new();
        let xml = write_threaded_comments(&comments).unwrap();
        let parsed = parse_threaded_comments(&xml).unwrap();
        assert!(parsed.is_empty());
    }

    #[test]
    fn write_persons_empty() {
        let persons: Vec<PersonData> = Vec::new();
        let xml = write_persons(&persons).unwrap();
        let parsed = parse_persons(&xml).unwrap();
        assert!(parsed.is_empty());
    }

    const NS: &str = "http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments";
    const C1: &str = "{A0000000-0000-4000-8000-000000000001}";
    const P1: &str = "{D0000000-0000-4000-8000-0000000000A1}";
    const M1: &str = "{E0000000-0000-4000-8000-0000000000B1}";

    #[test]
    fn dt_done_and_mentions_parse_and_write_back() {
        let xml = format!(
            r#"<ThreadedComments xmlns="{NS}"><threadedComment ref="A1" dT="2024-01-15T10:30:00.00" personId="{P1}" id="{C1}" done="1"><text>Hi @Jo</text><mentions><mention mentionpersonId="{P1}" mentionId="{M1}" startIndex="3" length="3"/></mentions></threadedComment></ThreadedComments>"#
        );
        let c = parse_threaded_comments(xml.as_bytes()).unwrap();
        assert_eq!(c.len(), 1);
        assert_eq!(c[0].timestamp, "2024-01-15T10:30:00.00");
        assert_eq!(c[0].done, Some(true));
        assert_eq!(
            c[0].mentions,
            vec![MentionData {
                mention_person_id: P1.to_string(),
                mention_id: M1.to_string(),
                start_index: 3,
                length: 3,
            }]
        );
        assert_eq!(c[0].text, "Hi @Jo");

        let out = write_threaded_comments(&c).unwrap();
        let s = std::str::from_utf8(&out).unwrap();
        assert!(s.contains(r#"dT="2024-01-15T10:30:00.00""#), "{s}");
        assert!(s.contains(r#"done="1""#), "{s}");
        assert!(s.contains(&format!(r#"mentionpersonId="{P1}""#)), "{s}");
        assert!(!s.contains(" dt="), "{s}");
        assert!(
            s.find("</text>").unwrap() < s.find("<mentions>").unwrap(),
            "{s}"
        );
        assert_eq!(parse_threaded_comments(&out).unwrap(), c);
    }

    #[test]
    fn excel_dt_conversion_table() {
        for (timestamp, written) in [
            // The API's form: three digits and Z (Excel reads .051 as .51 s, .151 not at all).
            ("2026-11-12T22:00:00.051Z", "2026-11-12T22:00:00.05"),
            // Truncated to hundredths, never carried into the seconds.
            ("2026-11-12T23:59:59.999999Z", "2026-11-12T23:59:59.99"),
            // No zone is UTC, the zone Excel writes dT in.
            ("2024-01-15T10:30:00.000", "2024-01-15T10:30:00.00"),
            ("2026-11-12T22:00:01Z", "2026-11-12T22:00:01.00"),
            ("2026-11-12T22:00:01", "2026-11-12T22:00:01.00"),
            // One digit is tenths (Excel would read .5 as .05 s).
            ("2026-11-12T22:00:01.5", "2026-11-12T22:00:01.50"),
            // An offset (Excel reads no date at all) is converted to UTC.
            ("2026-11-12T23:00:01.50+01:00", "2026-11-12T22:00:01.50"),
            ("2026-11-12T19:30:00.25-05:00", "2026-11-13T00:30:00.25"),
            ("2024-03-01T00:30:00+01:00", "2024-02-29T23:30:00.00"),
        ] {
            assert_eq!(excel_dt(timestamp), written, "{timestamp}");
        }
        // Excel's own form, and anything that is not such a timestamp, is written as it is.
        for kept in [
            "2026-11-12T22:17:43.41",
            "t",
            "",
            "2026-11-12",
            "2026-11-12T22:00",
            "2026-13-01T00:00:00Z",
            "2026-02-29T00:00:00Z",
            "2026-11-12T24:00:00Z",
            "2026-11-12T22:00:60Z",
            "2026-11-12T22:60:00Z",
            "2026-11-12 22:00:00Z",
            "2026-11-12T22:00:00.Z",
            "2026-11-12T22:00:00+0100",
            "2026-11-12T22:00:00+24:00",
            "2026-11-12T22:00:00z",
            " 2026-11-12T22:00:00Z",
            "2026-11-12T22:00:00Zjunk",
            "+12026-11-12T22:00:00Z",
            "2026-11-12T22:00:00.\u{0661}Z",
            // The offset moves the instant out of years 0000-9999, which
            // `YYYY` cannot write.
            "0000-01-01T00:30:00+01:00",
            "9999-12-31T23:30:00-01:00",
        ] {
            let out = excel_dt(kept);
            assert!(
                matches!(out, Cow::Borrowed(s) if s == kept),
                "{kept:?} -> {out:?}"
            );
        }
    }

    #[test]
    fn dt_is_written_in_excel_form_and_read_back_as_written() {
        let model = vec![ThreadedCommentData {
            id: C1.to_string(),
            ref_cell: "A1".to_string(),
            person_id: P1.to_string(),
            text: "x".to_string(),
            timestamp: "2026-11-12T22:00:00.151Z".to_string(),
            parent_id: None,
            done: None,
            mentions: Vec::new(),
        }];
        let out = write_threaded_comments(&model).unwrap();
        let s = std::str::from_utf8(&out).unwrap();
        assert!(s.contains(r#"dT="2026-11-12T22:00:00.15""#), "{s}");
        // The model is not changed by the write; a read returns the file's dT.
        assert_eq!(model[0].timestamp, "2026-11-12T22:00:00.151Z");
        let read = parse_threaded_comments(&out).unwrap();
        assert_eq!(read[0].timestamp, "2026-11-12T22:00:00.15");
        assert_eq!(write_threaded_comments(&read).unwrap(), out);
    }

    #[test]
    fn legacy_lowercase_dt_is_still_read() {
        let xml = format!(
            r#"<ThreadedComments xmlns="{NS}"><threadedComment id="{C1}" ref="A1" personId="{P1}" dt="2024-01-15T10:30:00.000"><text>x</text></threadedComment></ThreadedComments>"#
        );
        let c = parse_threaded_comments(xml.as_bytes()).unwrap();
        assert_eq!(c[0].timestamp, "2024-01-15T10:30:00.000");
    }

    #[test]
    fn done_false_round_trips_as_zero() {
        let mut c = parse_threaded_comments(
            format!(
                r#"<ThreadedComments xmlns="{NS}"><threadedComment id="{C1}" ref="A1" personId="{P1}" dT="t" done="0"><text>x</text></threadedComment></ThreadedComments>"#
            )
            .as_bytes(),
        )
        .unwrap();
        assert_eq!(c[0].done, Some(false));
        let out = write_threaded_comments(&c).unwrap();
        assert!(std::str::from_utf8(&out).unwrap().contains(r#"done="0""#));
        c[0].done = None;
        let out = write_threaded_comments(&c).unwrap();
        assert!(!std::str::from_utf8(&out).unwrap().contains("done="));
    }

    #[test]
    fn self_closing_threaded_comment_yields_empty_text() {
        let xml = format!(
            r#"<ThreadedComments xmlns="{NS}"><threadedComment id="{C1}" ref="B2" personId="{P1}" dT="t"/></ThreadedComments>"#
        );
        let c = parse_threaded_comments(xml.as_bytes()).unwrap();
        assert_eq!(c.len(), 1);
        assert_eq!(c[0].ref_cell, "B2");
        assert_eq!(c[0].text, "");
    }

    #[test]
    fn person_user_id_is_kept() {
        let xml = format!(
            r#"<personList xmlns="{NS}"><person displayName="A &amp; B" id="{P1}" userId="S::a@b.c::guid" providerId="AD"/></personList>"#
        );
        let p = parse_persons(xml.as_bytes()).unwrap();
        assert_eq!(p[0].display_name, "A & B");
        assert_eq!(p[0].user_id.as_deref(), Some("S::a@b.c::guid"));
        let out = write_persons(&p).unwrap();
        let s = std::str::from_utf8(&out).unwrap();
        assert!(s.contains(r#"displayName="A &amp; B""#), "{s}");
        assert!(!s.contains("&amp;amp;"), "{s}");
        assert_eq!(parse_persons(&out).unwrap(), p);
    }

    #[test]
    fn bad_entity_is_xml_parse_error() {
        let persons = format!(
            r#"<personList xmlns="{NS}"><person displayName="A &bad; B" id="{P1}"/></personList>"#
        );
        assert!(matches!(
            parse_persons(persons.as_bytes()),
            Err(ModernXlsxError::XmlParse(_))
        ));
        let comments = format!(
            r#"<ThreadedComments xmlns="{NS}"><threadedComment id="{C1}" ref="A&bad;" personId="{P1}" dT="t"/></ThreadedComments>"#
        );
        assert!(matches!(
            parse_threaded_comments(comments.as_bytes()),
            Err(ModernXlsxError::XmlParse(_))
        ));
    }

    #[test]
    fn mismatched_end_tag_is_xml_parse_error() {
        assert!(matches!(
            parse_threaded_comments(b"<root><a></b></root>"),
            Err(ModernXlsxError::XmlParse(_))
        ));
        assert!(matches!(
            parse_persons(b"<root><a></b></root>"),
            Err(ModernXlsxError::XmlParse(_))
        ));
    }

    fn one_text(text_xml: &str) -> String {
        let xml = format!(
            r#"<ThreadedComments xmlns="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments"><threadedComment id="{{A0000000-0000-4000-8000-000000000001}}" ref="A1" personId="{{D0000000-0000-4000-8000-0000000000A1}}" dT="2024-01-15T10:30:00.000"><text>{text_xml}</text></threadedComment></ThreadedComments>"#
        );
        parse_threaded_comments(xml.as_bytes()).unwrap()[0]
            .text
            .clone()
    }

    #[test]
    fn text_line_ends_are_normalized_and_char_refs_kept() {
        assert_eq!(one_text("a\r\nb\rc"), "a\nb\nc");
        assert_eq!(one_text("a&#13;&#10;b"), "a\r\nb");
    }

    #[test]
    fn cdata_text_is_read() {
        assert_eq!(one_text("<![CDATA[x&y]]>"), "x&y");
    }
}
