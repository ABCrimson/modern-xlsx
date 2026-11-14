//! Legacy cell comments (notes) — `xl/comments{n}.xml` (ECMA-376 §18.7).
//!
//! Only the author list and each comment's cell reference, author and plain
//! text are modelled; rich-text formatting is flattened on read. This module
//! writes the comments part alone. Excel shows each note through a shape in
//! the sheet's VML drawing (`vmlDrawing`), which the package writer names in
//! `<legacyDrawing>`: it keeps the drawing a sheet was read with when its
//! note shapes match the comments, merges the comments into it otherwise,
//! and generates one when there is none (`ooxml::vml`). A drawing it cannot
//! merge (not UTF-8, not well-formed XML, or binding the `v`, `o` or `x`
//! prefix to another namespace) is written as it is, and a drawing several
//! sheets share is merged for the first of them with comments only.
//!
//! A threaded comment (`ooxml::threaded_comments`) shows in Excel only when
//! its cell also has a comment here whose author is `tc={thread root id}`,
//! the placeholder older Excel versions display. Excel writes one for each
//! thread and the reader reads it as an ordinary comment; the package writer
//! adds one to the part it writes for each thread root whose cell has no
//! comment (`threaded_comments::placeholder_comments`), and writes the
//! comments it was given as they are. The note shapes and placeholders it
//! generates count against one ceiling per write
//! (`vml::MAX_GENERATED_NOTE_SHAPES`; past it the write fails with
//! `SizeLimit`).

use core::hint::cold_path;

use quick_xml::events::{BytesEnd, BytesStart, BytesText, Event};
use quick_xml::{Reader, Writer};
use serde::{Deserialize, Serialize};

use crate::{ModernXlsxError, Result};

use super::shared_strings::ExpansionBudget;
use super::xml::{attr_str, attrs, needs_preserve, parse_attr, push_general_ref, xml_attr};
use super::{SPREADSHEET_NS, write_xml_decl, xml_write_err};

/// A single cell comment / note.
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct Comment {
    /// Cell reference, e.g. "A1".
    pub cell_ref: String,
    /// Author name.
    pub author: String,
    /// Plain text content.
    pub text: String,
}

/// Parse a `comments*.xml` file from raw XML bytes.
///
/// The XML structure expected is:
/// ```xml
/// <comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
///   <authors>
///     <author>Author Name</author>
///   </authors>
///   <commentList>
///     <comment ref="A1" authorId="0">
///       <text>
///         <t>Comment text here</t>
///       </text>
///     </comment>
///   </commentList>
/// </comments>
/// ```
///
/// Rich text runs inside `<text>` are flattened to plain text (all `<t>` runs
/// are concatenated). Phonetic runs (`<rPh>`) are ignored. An `authorId` with
/// no matching `<author>` reads as an empty author name. Text and CDATA are
/// read with XML 1.0 line-end normalization (a literal CRLF or CR reads as LF;
/// `&#13;` stays a CR).
///
/// # Errors
///
/// `XmlParse` for malformed XML, an attribute holding a bad entity, or an
/// `authorId` that is not a number.
pub fn parse_comments(data: &[u8]) -> Result<Vec<Comment>> {
    parse_comments_within(data, &mut ExpansionBudget::new(usize::MAX))
}

/// [`parse_comments`], counting in `budget` the author name copied into each
/// comment (one `<author>` may be named by any number of comments).
///
/// # Errors
///
/// As [`parse_comments`], and `SizeLimit` when the count passes the budget.
pub(crate) fn parse_comments_within(
    data: &[u8],
    budget: &mut ExpansionBudget,
) -> Result<Vec<Comment>> {
    let mut reader = Reader::from_reader(data);
    reader.config_mut().trim_text(false);

    let mut buf = Vec::with_capacity(512);
    let mut authors: Vec<String> = Vec::new();
    let mut comments: Vec<Comment> = Vec::new();

    // State tracking.
    let mut in_authors = false;
    let mut in_author = false;
    let mut in_comment = false;
    let mut in_text = false;
    let mut in_t = false;
    // Inside a phonetic run (`<rPh>`): its `<t>` holds the reading, not the text.
    let mut in_rph = false;

    let mut current_author_text = String::new();
    let mut current_ref = String::new();
    let mut current_author_id: usize = 0;
    let mut current_comment_text = String::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) => {
                let local = e.local_name();
                match local.as_ref() {
                    "authors" => {
                        in_authors = true;
                    }
                    "author" if in_authors => {
                        in_author = true;
                        current_author_text.clear();
                    }
                    "comment" => {
                        in_comment = true;
                        current_ref.clear();
                        current_author_id = 0;
                        current_comment_text.clear();
                        for attr in attrs(e) {
                            let attr = attr?;
                            match attr.key.local_name().as_ref() {
                                "ref" => {
                                    current_ref = attr_str(&attr)?.into_owned();
                                }
                                "authorId" => {
                                    current_author_id = parse_attr("comment", &attr)?;
                                }
                                _ => {}
                            }
                        }
                    }
                    "text" if in_comment => {
                        in_text = true;
                    }
                    "rPh" if in_text => {
                        in_rph = true;
                    }
                    "t" if in_text && !in_rph => {
                        in_t = true;
                    }
                    _ => {}
                }
            }
            Ok(Event::Text(ref e)) => {
                if in_author {
                    current_author_text.push_str(&e.xml10_content());
                } else if in_t {
                    current_comment_text.push_str(&e.xml10_content());
                }
            }
            Ok(Event::CData(ref e)) => {
                if in_author {
                    current_author_text.push_str(&e.xml10_content());
                } else if in_t {
                    current_comment_text.push_str(&e.xml10_content());
                }
            }
            Ok(Event::GeneralRef(ref e)) => {
                if in_author {
                    push_general_ref(&mut current_author_text, e)?;
                } else if in_t {
                    push_general_ref(&mut current_comment_text, e)?;
                }
            }
            Ok(Event::End(ref e)) => {
                let local = e.local_name();
                match local.as_ref() {
                    "authors" => {
                        in_authors = false;
                    }
                    "author" => {
                        authors.push(std::mem::take(&mut current_author_text));
                        in_author = false;
                    }
                    "comment" => {
                        let author = authors.get(current_author_id).map_or("", String::as_str);
                        budget.spend(author.len())?;
                        comments.push(Comment {
                            cell_ref: std::mem::take(&mut current_ref),
                            author: author.to_owned(),
                            text: std::mem::take(&mut current_comment_text),
                        });
                        in_comment = false;
                    }
                    "text" => {
                        in_text = false;
                    }
                    "rPh" => {
                        in_rph = false;
                    }
                    "t" => {
                        in_t = false;
                    }
                    _ => {}
                }
            }
            // A self-closing `<author/>` is an empty author name.
            Ok(Event::Empty(ref e)) if in_authors && e.local_name().as_ref() == "author" => {
                authors.push(String::new());
            }
            Ok(Event::Eof) => break,
            Ok(_) => {}
            Err(e) => {
                cold_path();
                return Err(ModernXlsxError::XmlParse(format!(
                    "error parsing comments XML: {e}"
                )));
            }
        }
        buf.clear();
    }

    Ok(comments)
}

/// Serialize a list of comments to XML bytes suitable for inclusion in an XLSX
/// archive as `xl/comments*.xml`.
pub fn write_comments(comments: &[Comment]) -> Result<Vec<u8>> {
    let mut buf: Vec<u8> = Vec::with_capacity(512 + comments.len() * 128);
    let mut writer = Writer::new(&mut buf);

    // XML declaration.
    write_xml_decl(&mut writer)?;

    // <comments xmlns="...">
    let mut root = BytesStart::new("comments");
    root.push_attribute(("xmlns", SPREADSHEET_NS));
    writer
        .write_event(Event::Start(root))
        .map_err(xml_write_err)?;

    // Build deduplicated author list, preserving insertion order.
    let mut author_list: Vec<&str> = Vec::with_capacity(comments.len().min(16));
    let mut author_indices: std::collections::HashMap<&str, usize> =
        std::collections::HashMap::with_capacity(comments.len().min(16));
    for comment in comments {
        let next_idx = author_list.len();
        if let std::collections::hash_map::Entry::Vacant(e) =
            author_indices.entry(comment.author.as_str())
        {
            e.insert(next_idx);
            author_list.push(comment.author.as_str());
        }
    }

    // <authors>
    writer
        .write_event(Event::Start(BytesStart::new("authors")))
        .map_err(xml_write_err)?;
    for author in author_list {
        writer
            .write_event(Event::Start(BytesStart::new("author")))
            .map_err(xml_write_err)?;
        writer
            .write_event(Event::Text(BytesText::new(
                crate::ooxml::sanitize_xml_text(author).as_ref(),
            )))
            .map_err(xml_write_err)?;
        writer
            .write_event(Event::End(BytesEnd::new("author")))
            .map_err(xml_write_err)?;
    }
    writer
        .write_event(Event::End(BytesEnd::new("authors")))
        .map_err(xml_write_err)?;

    // <commentList>
    writer
        .write_event(Event::Start(BytesStart::new("commentList")))
        .map_err(xml_write_err)?;

    let mut ibuf = itoa::Buffer::new();
    for comment in comments {
        let author_id = author_indices
            .get(comment.author.as_str())
            .copied()
            .unwrap_or(0);

        let mut elem = BytesStart::new("comment");
        elem.push_attribute(xml_attr("ref", &comment.cell_ref));
        elem.push_attribute(("authorId", ibuf.format(author_id)));
        writer
            .write_event(Event::Start(elem))
            .map_err(xml_write_err)?;

        // <text><t>...</t></text>
        writer
            .write_event(Event::Start(BytesStart::new("text")))
            .map_err(xml_write_err)?;
        let text = crate::ooxml::sanitize_xml_text(&comment.text);
        let mut t_elem = BytesStart::new("t");
        if needs_preserve(&text) {
            t_elem.push_attribute(("xml:space", "preserve"));
        }
        writer
            .write_event(Event::Start(t_elem))
            .map_err(xml_write_err)?;
        writer
            .write_event(Event::Text(BytesText::new(text.as_ref())))
            .map_err(xml_write_err)?;
        writer
            .write_event(Event::End(BytesEnd::new("t")))
            .map_err(xml_write_err)?;
        writer
            .write_event(Event::End(BytesEnd::new("text")))
            .map_err(xml_write_err)?;

        // </comment>
        writer
            .write_event(Event::End(BytesEnd::new("comment")))
            .map_err(xml_write_err)?;
    }

    // </commentList>
    writer
        .write_event(Event::End(BytesEnd::new("commentList")))
        .map_err(xml_write_err)?;

    // </comments>
    writer
        .write_event(Event::End(BytesEnd::new("comments")))
        .map_err(xml_write_err)?;

    Ok(buf)
}

#[cfg(test)]
mod tests {
    use super::*;
    use pretty_assertions::assert_eq;

    #[test]
    fn test_parse_comments_basic() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <authors>
    <author>John Doe</author>
  </authors>
  <commentList>
    <comment ref="A1" authorId="0">
      <text>
        <t>This is a comment</t>
      </text>
    </comment>
  </commentList>
</comments>"#;

        let comments = parse_comments(xml.as_bytes()).unwrap();
        assert_eq!(comments.len(), 1);
        assert_eq!(comments[0].cell_ref, "A1");
        assert_eq!(comments[0].author, "John Doe");
        assert_eq!(comments[0].text, "This is a comment");
    }

    #[test]
    fn test_parse_comments_multiple() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <authors>
    <author>Alice</author>
    <author>Bob</author>
  </authors>
  <commentList>
    <comment ref="A1" authorId="0">
      <text>
        <t>Comment from Alice</t>
      </text>
    </comment>
    <comment ref="B2" authorId="1">
      <text>
        <t>Comment from Bob</t>
      </text>
    </comment>
    <comment ref="C3" authorId="0">
      <text>
        <t>Another from Alice</t>
      </text>
    </comment>
  </commentList>
</comments>"#;

        let comments = parse_comments(xml.as_bytes()).unwrap();
        assert_eq!(comments.len(), 3);

        assert_eq!(comments[0].cell_ref, "A1");
        assert_eq!(comments[0].author, "Alice");
        assert_eq!(comments[0].text, "Comment from Alice");

        assert_eq!(comments[1].cell_ref, "B2");
        assert_eq!(comments[1].author, "Bob");
        assert_eq!(comments[1].text, "Comment from Bob");

        assert_eq!(comments[2].cell_ref, "C3");
        assert_eq!(comments[2].author, "Alice");
        assert_eq!(comments[2].text, "Another from Alice");
    }

    #[test]
    fn test_parse_comments_rich_text_flattened() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <authors>
    <author>User</author>
  </authors>
  <commentList>
    <comment ref="D4" authorId="0">
      <text>
        <r><rPr><b/></rPr><t>Bold</t></r>
        <r><t> and normal</t></r>
      </text>
    </comment>
  </commentList>
</comments>"#;

        let comments = parse_comments(xml.as_bytes()).unwrap();
        assert_eq!(comments.len(), 1);
        assert_eq!(comments[0].cell_ref, "D4");
        assert_eq!(comments[0].author, "User");
        assert_eq!(comments[0].text, "Bold and normal");
    }

    #[test]
    fn test_parse_comments_empty() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <authors/>
  <commentList/>
</comments>"#;

        let comments = parse_comments(xml.as_bytes()).unwrap();
        assert!(comments.is_empty());
    }

    #[test]
    fn test_write_comments_basic() {
        let comments = vec![Comment {
            cell_ref: "A1".to_string(),
            author: "Test Author".to_string(),
            text: "Hello comment".to_string(),
        }];

        let xml = write_comments(&comments).unwrap();
        let parsed = parse_comments(&xml).unwrap();

        assert_eq!(parsed.len(), 1);
        assert_eq!(parsed[0].cell_ref, "A1");
        assert_eq!(parsed[0].author, "Test Author");
        assert_eq!(parsed[0].text, "Hello comment");
    }

    #[test]
    fn test_write_comments_empty() {
        let comments: Vec<Comment> = Vec::new();
        let xml = write_comments(&comments).unwrap();
        let parsed = parse_comments(&xml).unwrap();
        assert!(parsed.is_empty());
    }

    #[test]
    fn test_comments_roundtrip() {
        let comments = vec![
            Comment {
                cell_ref: "A1".to_string(),
                author: "Alice".to_string(),
                text: "First comment".to_string(),
            },
            Comment {
                cell_ref: "B2".to_string(),
                author: "Bob".to_string(),
                text: "Second comment".to_string(),
            },
            Comment {
                cell_ref: "C3".to_string(),
                author: "Alice".to_string(),
                text: "Third comment from Alice".to_string(),
            },
        ];

        let xml = write_comments(&comments).unwrap();
        let parsed = parse_comments(&xml).unwrap();

        assert_eq!(parsed.len(), 3);
        for (orig, round) in comments.iter().zip(parsed.iter()) {
            assert_eq!(orig.cell_ref, round.cell_ref);
            assert_eq!(orig.author, round.author);
            assert_eq!(orig.text, round.text);
        }
    }

    #[test]
    fn test_comments_special_characters() {
        let comments = vec![Comment {
            cell_ref: "A1".to_string(),
            author: "O'Reilly & Sons".to_string(),
            text: "This has <special> \"characters\" & entities".to_string(),
        }];

        let xml = write_comments(&comments).unwrap();
        let parsed = parse_comments(&xml).unwrap();

        assert_eq!(parsed.len(), 1);
        assert_eq!(parsed[0].author, "O'Reilly & Sons");
        assert_eq!(
            parsed[0].text,
            "This has <special> \"characters\" & entities"
        );
    }

    #[test]
    fn test_write_deduplicates_authors() {
        let comments = vec![
            Comment {
                cell_ref: "A1".to_string(),
                author: "Alice".to_string(),
                text: "First".to_string(),
            },
            Comment {
                cell_ref: "B1".to_string(),
                author: "Alice".to_string(),
                text: "Second".to_string(),
            },
        ];

        let xml = write_comments(&comments).unwrap();
        let xml_str = std::str::from_utf8(&xml).unwrap();

        // "Alice" should appear only once in the <authors> section.
        let author_count = xml_str.matches("<author>").count();
        assert_eq!(author_count, 1, "author 'Alice' should be deduplicated");

        // Both comments should reference authorId="0".
        let parsed = parse_comments(&xml).unwrap();
        assert_eq!(parsed.len(), 2);
        assert_eq!(parsed[0].author, "Alice");
        assert_eq!(parsed[1].author, "Alice");
    }

    const NS: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

    fn one(text: &str) -> Vec<Comment> {
        vec![Comment {
            cell_ref: "A1".to_string(),
            author: "A".to_string(),
            text: text.to_string(),
        }]
    }

    #[test]
    fn edge_whitespace_text_is_written_with_xml_space_preserve() {
        let xml = write_comments(&one(" note ")).unwrap();
        let s = std::str::from_utf8(&xml).unwrap();
        assert!(s.contains(r#"<t xml:space="preserve"> note </t>"#), "{s}");
        assert_eq!(parse_comments(&xml).unwrap()[0].text, " note ");
    }

    #[test]
    fn plain_text_has_no_xml_space() {
        let xml = write_comments(&one("note")).unwrap();
        let s = std::str::from_utf8(&xml).unwrap();
        assert!(s.contains("<t>note</t>"), "{s}");
        assert!(!s.contains("xml:space"), "{s}");
    }

    #[test]
    fn phonetic_run_text_is_excluded() {
        let xml = format!(
            r#"<comments xmlns="{NS}"><authors><author>A</author></authors><commentList><comment ref="A1" authorId="0"><text><r><t>kanji</t></r><rPh sb="0" eb="1"><t>reading</t></rPh><phoneticPr fontId="0"/></text></comment></commentList></comments>"#
        );
        let c = parse_comments(xml.as_bytes()).unwrap();
        assert_eq!(c[0].text, "kanji");
    }

    #[test]
    fn author_id_out_of_range_reads_empty_author() {
        let xml = format!(
            r#"<comments xmlns="{NS}"><authors><author>A</author></authors><commentList><comment ref="A1" authorId="5"><text><t>x</t></text></comment></commentList></comments>"#
        );
        let c = parse_comments(xml.as_bytes()).unwrap();
        assert_eq!(c[0].author, "");
        assert_eq!(c[0].text, "x");
    }

    #[test]
    fn invalid_author_id_is_xml_parse_error() {
        let xml = format!(
            r#"<comments xmlns="{NS}"><authors><author>A</author></authors><commentList><comment ref="A1" authorId="-1"><text><t>x</t></text></comment></commentList></comments>"#
        );
        assert!(matches!(
            parse_comments(xml.as_bytes()),
            Err(ModernXlsxError::XmlParse(_))
        ));
    }

    #[test]
    fn bad_entity_in_ref_is_xml_parse_error() {
        let xml = format!(
            r#"<comments xmlns="{NS}"><authors><author>A</author></authors><commentList><comment ref="A&bad;1" authorId="0"><text><t>x</t></text></comment></commentList></comments>"#
        );
        assert!(matches!(
            parse_comments(xml.as_bytes()),
            Err(ModernXlsxError::XmlParse(_))
        ));
    }

    #[test]
    fn ref_attribute_is_escaped_once() {
        let comments = vec![Comment {
            cell_ref: "A&B\"1".to_string(),
            author: "A".to_string(),
            text: "x".to_string(),
        }];
        let xml = write_comments(&comments).unwrap();
        let s = std::str::from_utf8(&xml).unwrap();
        assert!(s.contains(r#"ref="A&amp;B&quot;1""#), "{s}");
        assert_eq!(parse_comments(&xml).unwrap()[0].cell_ref, "A&B\"1");
    }

    #[test]
    fn mismatched_end_tag_is_xml_parse_error() {
        assert!(matches!(
            parse_comments(b"<root><a></b></root>"),
            Err(ModernXlsxError::XmlParse(_))
        ));
    }

    #[test]
    fn self_closing_author_keeps_author_ids_aligned() {
        let xml = r#"<comments><authors><author/><author>Bob</author></authors><commentList><comment ref="A1" authorId="0"><text><t>x</t></text></comment><comment ref="A2" authorId="1"><text><t>y</t></text></comment></commentList></comments>"#;
        let comments = parse_comments(xml.as_bytes()).unwrap();
        assert_eq!(comments[0].author, "");
        assert_eq!(comments[1].author, "Bob");
    }

    #[test]
    fn literal_line_ends_are_normalized_and_char_refs_kept() {
        let xml = "<comments><authors><author>A\r\nB</author></authors><commentList>\
            <comment ref=\"A1\" authorId=\"0\"><text><t>Line1\r\nLine2\rLine3</t></text></comment>\
            <comment ref=\"A2\" authorId=\"0\"><text><t>a&#13;&#10;b</t></text></comment>\
            </commentList></comments>";
        let comments = parse_comments(xml.as_bytes()).unwrap();
        assert_eq!(comments[0].author, "A\nB");
        assert_eq!(comments[0].text, "Line1\nLine2\nLine3");
        assert_eq!(comments[1].text, "a\r\nb");
    }

    #[test]
    fn cdata_text_is_read() {
        let xml = r#"<comments><authors><author><![CDATA[A&B]]></author></authors><commentList><comment ref="A1" authorId="0"><text><t><![CDATA[a<b]]></t></text></comment></commentList></comments>"#;
        let comments = parse_comments(xml.as_bytes()).unwrap();
        assert_eq!(comments[0].author, "A&B");
        assert_eq!(comments[0].text, "a<b");
    }
}
