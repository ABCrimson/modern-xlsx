use quick_xml::events::{BytesDecl, BytesEnd, BytesStart, BytesText, Event};
use quick_xml::{Reader, Writer};
use serde::{Deserialize, Serialize};

use crate::{ModernXlsxError, Result};

use super::{push_entity, SPREADSHEET_NS};

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
/// are concatenated).
pub fn parse_comments(data: &[u8]) -> Result<Vec<Comment>> {
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

    let mut current_author_text = String::new();
    let mut current_ref = String::new();
    let mut current_author_id: usize = 0;
    let mut current_comment_text = String::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) => {
                let local = e.local_name();
                match local.as_ref() {
                    b"authors" => {
                        in_authors = true;
                    }
                    b"author" if in_authors => {
                        in_author = true;
                        current_author_text.clear();
                    }
                    b"comment" => {
                        in_comment = true;
                        current_ref.clear();
                        current_author_id = 0;
                        current_comment_text.clear();
                        for attr in e.attributes().flatten() {
                            match attr.key.local_name().as_ref() {
                                b"ref" => {
                                    current_ref = std::str::from_utf8(&attr.value)
                                        .unwrap_or_default()
                                        .to_owned();
                                }
                                b"authorId" => {
                                    current_author_id = std::str::from_utf8(&attr.value)
                                        .unwrap_or_default()
                                        .parse::<usize>()
                                        .unwrap_or(0);
                                }
                                _ => {}
                            }
                        }
                    }
                    b"text" if in_comment => {
                        in_text = true;
                    }
                    b"t" if in_text => {
                        in_t = true;
                    }
                    _ => {}
                }
            }
            Ok(Event::Text(ref e)) => {
                if in_author {
                    let text = std::str::from_utf8(e.as_ref()).unwrap_or_default();
                    current_author_text.push_str(text);
                } else if in_t {
                    let text = std::str::from_utf8(e.as_ref()).unwrap_or_default();
                    current_comment_text.push_str(text);
                }
            }
            Ok(Event::GeneralRef(ref e)) => {
                if in_author {
                    push_entity(&mut current_author_text, e.as_ref());
                } else if in_t {
                    push_entity(&mut current_comment_text, e.as_ref());
                }
            }
            Ok(Event::End(ref e)) => {
                let local = e.local_name();
                match local.as_ref() {
                    b"authors" => {
                        in_authors = false;
                    }
                    b"author" => {
                        authors.push(std::mem::take(&mut current_author_text));
                        in_author = false;
                    }
                    b"comment" => {
                        let author = authors
                            .get(current_author_id)
                            .cloned()
                            .unwrap_or_default();
                        comments.push(Comment {
                            cell_ref: std::mem::take(&mut current_ref),
                            author,
                            text: std::mem::take(&mut current_comment_text),
                        });
                        in_comment = false;
                    }
                    b"text" => {
                        in_text = false;
                    }
                    b"t" => {
                        in_t = false;
                    }
                    _ => {}
                }
            }
            Ok(Event::Eof) => break,
            Ok(_) => {}
            Err(e) => {
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

    let map_err = |e: std::io::Error| ModernXlsxError::XmlWrite(e.to_string());

    // XML declaration.
    writer
        .write_event(Event::Decl(BytesDecl::new("1.0", Some("UTF-8"), Some("yes"))))
        .map_err(map_err)?;

    // <comments xmlns="...">
    let mut root = BytesStart::new("comments");
    root.push_attribute(("xmlns", SPREADSHEET_NS));
    writer.write_event(Event::Start(root)).map_err(map_err)?;

    // Build deduplicated author list, preserving insertion order.
    let mut author_list: Vec<String> = Vec::new();
    let mut author_indices: std::collections::HashMap<String, usize> =
        std::collections::HashMap::new();
    for comment in comments {
        if !author_indices.contains_key(&comment.author) {
            let idx = author_list.len();
            author_indices.insert(comment.author.clone(), idx);
            author_list.push(comment.author.clone());
        }
    }

    // <authors>
    writer
        .write_event(Event::Start(BytesStart::new("authors")))
        .map_err(map_err)?;
    for author in &author_list {
        writer
            .write_event(Event::Start(BytesStart::new("author")))
            .map_err(map_err)?;
        writer
            .write_event(Event::Text(BytesText::new(author)))
            .map_err(map_err)?;
        writer
            .write_event(Event::End(BytesEnd::new("author")))
            .map_err(map_err)?;
    }
    writer
        .write_event(Event::End(BytesEnd::new("authors")))
        .map_err(map_err)?;

    // <commentList>
    writer
        .write_event(Event::Start(BytesStart::new("commentList")))
        .map_err(map_err)?;

    let mut ibuf = itoa::Buffer::new();
    for comment in comments {
        let author_id = author_indices
            .get(&comment.author)
            .copied()
            .unwrap_or(0);

        let mut elem = BytesStart::new("comment");
        elem.push_attribute(("ref", comment.cell_ref.as_str()));
        elem.push_attribute(("authorId", ibuf.format(author_id)));
        writer.write_event(Event::Start(elem)).map_err(map_err)?;

        // <text><t>...</t></text>
        writer
            .write_event(Event::Start(BytesStart::new("text")))
            .map_err(map_err)?;
        writer
            .write_event(Event::Start(BytesStart::new("t")))
            .map_err(map_err)?;
        writer
            .write_event(Event::Text(BytesText::new(&comment.text)))
            .map_err(map_err)?;
        writer
            .write_event(Event::End(BytesEnd::new("t")))
            .map_err(map_err)?;
        writer
            .write_event(Event::End(BytesEnd::new("text")))
            .map_err(map_err)?;

        // </comment>
        writer
            .write_event(Event::End(BytesEnd::new("comment")))
            .map_err(map_err)?;
    }

    // </commentList>
    writer
        .write_event(Event::End(BytesEnd::new("commentList")))
        .map_err(map_err)?;

    // </comments>
    writer
        .write_event(Event::End(BytesEnd::new("comments")))
        .map_err(map_err)?;

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
}
