use core::hint::cold_path;

use quick_xml::events::{BytesDecl, BytesEnd, BytesStart, BytesText, Event};
use quick_xml::{Reader, Writer};
use serde::{Deserialize, Serialize};

use crate::{ModernXlsxError, Result};

use super::push_entity;

const THREADED_COMMENTS_NS: &str =
    "http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments";

/// A single threaded comment (Microsoft 365 modern comments).
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ThreadedCommentData {
    /// Unique comment identifier (GUID).
    pub id: String,
    /// Cell reference, e.g. "A1".
    pub ref_cell: String,
    /// Person identifier (GUID) of the comment author.
    pub person_id: String,
    /// Plain text content of the comment.
    pub text: String,
    /// ISO 8601 timestamp of the comment.
    pub timestamp: String,
    /// Parent comment ID for replies (forms a thread).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub parent_id: Option<String>,
}

/// A person entry from the persons list (xl/persons/person.xml).
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct PersonData {
    /// Unique person identifier (GUID).
    pub id: String,
    /// Display name of the person.
    pub display_name: String,
    /// Optional provider identifier (e.g. "Windows Live").
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub provider_id: Option<String>,
}

// ---------------------------------------------------------------------------
// Parsers
// ---------------------------------------------------------------------------

/// Parse a threaded comments XML file from raw bytes.
///
/// The XML structure expected is:
/// ```xml
/// <ThreadedComments xmlns="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments">
///   <threadedComment id="{GUID}" ref="A1" personId="{GUID}" dt="2024-01-15T10:30:00.000">
///     <text>This is a comment</text>
///   </threadedComment>
/// </ThreadedComments>
/// ```
pub fn parse_threaded_comments(data: &[u8]) -> Result<Vec<ThreadedCommentData>> {
    let mut reader = Reader::from_reader(data);
    reader.config_mut().trim_text(false);

    let mut buf = Vec::with_capacity(512);
    let mut comments: Vec<ThreadedCommentData> = Vec::new();

    // State tracking.
    let mut in_comment = false;
    let mut in_text = false;

    let mut current_id = String::new();
    let mut current_ref = String::new();
    let mut current_person_id = String::new();
    let mut current_timestamp = String::new();
    let mut current_parent_id: Option<String> = None;
    let mut current_text = String::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) => {
                let local = e.local_name();
                match local.as_ref() {
                    b"threadedComment" => {
                        in_comment = true;
                        current_id.clear();
                        current_ref.clear();
                        current_person_id.clear();
                        current_timestamp.clear();
                        current_parent_id = None;
                        current_text.clear();

                        for attr in e.attributes().flatten() {
                            match attr.key.local_name().as_ref() {
                                b"id" => {
                                    current_id = std::str::from_utf8(&attr.value)
                                        .unwrap_or_default()
                                        .to_owned();
                                }
                                b"ref" => {
                                    current_ref = std::str::from_utf8(&attr.value)
                                        .unwrap_or_default()
                                        .to_owned();
                                }
                                b"personId" => {
                                    current_person_id = std::str::from_utf8(&attr.value)
                                        .unwrap_or_default()
                                        .to_owned();
                                }
                                b"dt" => {
                                    current_timestamp = std::str::from_utf8(&attr.value)
                                        .unwrap_or_default()
                                        .to_owned();
                                }
                                b"parentId" => {
                                    current_parent_id = Some(
                                        std::str::from_utf8(&attr.value)
                                            .unwrap_or_default()
                                            .to_owned(),
                                    );
                                }
                                _ => {}
                            }
                        }
                    }
                    b"text" if in_comment => {
                        in_text = true;
                    }
                    _ => {}
                }
            }
            Ok(Event::Text(ref e)) => {
                if in_text {
                    let text = std::str::from_utf8(e.as_ref()).unwrap_or_default();
                    current_text.push_str(text);
                }
            }
            Ok(Event::GeneralRef(ref e)) => {
                if in_text {
                    push_entity(&mut current_text, e.as_ref());
                }
            }
            Ok(Event::End(ref e)) => {
                let local = e.local_name();
                match local.as_ref() {
                    b"threadedComment" => {
                        comments.push(ThreadedCommentData {
                            id: std::mem::take(&mut current_id),
                            ref_cell: std::mem::take(&mut current_ref),
                            person_id: std::mem::take(&mut current_person_id),
                            text: std::mem::take(&mut current_text),
                            timestamp: std::mem::take(&mut current_timestamp),
                            parent_id: current_parent_id.take(),
                        });
                        in_comment = false;
                    }
                    b"text" => {
                        in_text = false;
                    }
                    _ => {}
                }
            }
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
///   <person id="{GUID}" displayName="John Doe" providerId="Windows Live"/>
/// </personList>
/// ```
pub fn parse_persons(data: &[u8]) -> Result<Vec<PersonData>> {
    let mut reader = Reader::from_reader(data);
    let mut buf = Vec::with_capacity(256);
    let mut persons: Vec<PersonData> = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Empty(ref e)) | Ok(Event::Start(ref e)) => {
                if e.local_name().as_ref() == b"person" {
                    let mut id = String::new();
                    let mut display_name = String::new();
                    let mut provider_id: Option<String> = None;

                    for attr in e.attributes().flatten() {
                        match attr.key.local_name().as_ref() {
                            b"id" => {
                                id = attr
                                    .unescape_value()
                                    .unwrap_or_default()
                                    .into_owned();
                            }
                            b"displayName" => {
                                display_name = attr
                                    .unescape_value()
                                    .unwrap_or_default()
                                    .into_owned();
                            }
                            b"providerId" => {
                                provider_id = Some(
                                    attr.unescape_value()
                                        .unwrap_or_default()
                                        .into_owned(),
                                );
                            }
                            _ => {}
                        }
                    }

                    if !id.is_empty() {
                        persons.push(PersonData {
                            id,
                            display_name,
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

/// Serialize threaded comments to XML bytes.
pub fn write_threaded_comments(comments: &[ThreadedCommentData]) -> Result<Vec<u8>> {
    let mut buf: Vec<u8> = Vec::with_capacity(512 + comments.len() * 256);
    let mut writer = Writer::new(&mut buf);

    let map_err = |e: std::io::Error| ModernXlsxError::XmlWrite(e.to_string());

    // XML declaration.
    writer
        .write_event(Event::Decl(BytesDecl::new("1.0", Some("UTF-8"), Some("yes"))))
        .map_err(map_err)?;

    // <ThreadedComments xmlns="...">
    let mut root = BytesStart::new("ThreadedComments");
    root.push_attribute(("xmlns", THREADED_COMMENTS_NS));
    writer.write_event(Event::Start(root)).map_err(map_err)?;

    for comment in comments {
        let mut elem = BytesStart::new("threadedComment");
        elem.push_attribute(("id", comment.id.as_str()));
        elem.push_attribute(("ref", comment.ref_cell.as_str()));
        elem.push_attribute(("personId", comment.person_id.as_str()));
        elem.push_attribute(("dt", comment.timestamp.as_str()));

        if let Some(ref parent_id) = comment.parent_id {
            elem.push_attribute(("parentId", parent_id.as_str()));
        }

        writer.write_event(Event::Start(elem)).map_err(map_err)?;

        // <text>...</text>
        writer
            .write_event(Event::Start(BytesStart::new("text")))
            .map_err(map_err)?;
        writer
            .write_event(Event::Text(BytesText::new(&comment.text)))
            .map_err(map_err)?;
        writer
            .write_event(Event::End(BytesEnd::new("text")))
            .map_err(map_err)?;

        // </threadedComment>
        writer
            .write_event(Event::End(BytesEnd::new("threadedComment")))
            .map_err(map_err)?;
    }

    // </ThreadedComments>
    writer
        .write_event(Event::End(BytesEnd::new("ThreadedComments")))
        .map_err(map_err)?;

    Ok(buf)
}

/// Serialize persons to XML bytes.
pub fn write_persons(persons: &[PersonData]) -> Result<Vec<u8>> {
    let mut buf: Vec<u8> = Vec::with_capacity(256 + persons.len() * 128);
    let mut writer = Writer::new(&mut buf);

    let map_err = |e: std::io::Error| ModernXlsxError::XmlWrite(e.to_string());

    // XML declaration.
    writer
        .write_event(Event::Decl(BytesDecl::new("1.0", Some("UTF-8"), Some("yes"))))
        .map_err(map_err)?;

    // <personList xmlns="...">
    let mut root = BytesStart::new("personList");
    root.push_attribute(("xmlns", THREADED_COMMENTS_NS));
    writer.write_event(Event::Start(root)).map_err(map_err)?;

    for person in persons {
        let mut elem = BytesStart::new("person");
        elem.push_attribute(("id", person.id.as_str()));
        elem.push_attribute(("displayName", person.display_name.as_str()));

        if let Some(ref provider_id) = person.provider_id {
            elem.push_attribute(("providerId", provider_id.as_str()));
        }

        writer.write_event(Event::Empty(elem)).map_err(map_err)?;
    }

    // </personList>
    writer
        .write_event(Event::End(BytesEnd::new("personList")))
        .map_err(map_err)?;

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
  <threadedComment id="{AAA-111}" ref="A1" personId="{PPP-001}" dt="2024-01-15T10:30:00.000">
    <text>This is a comment</text>
  </threadedComment>
</ThreadedComments>"#;

        let comments = parse_threaded_comments(xml.as_bytes()).unwrap();
        assert_eq!(comments.len(), 1);
        assert_eq!(comments[0].id, "{AAA-111}");
        assert_eq!(comments[0].ref_cell, "A1");
        assert_eq!(comments[0].person_id, "{PPP-001}");
        assert_eq!(comments[0].text, "This is a comment");
        assert_eq!(comments[0].timestamp, "2024-01-15T10:30:00.000");
        assert!(comments[0].parent_id.is_none());
    }

    #[test]
    fn parse_reply_chain() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<ThreadedComments xmlns="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments">
  <threadedComment id="{AAA-111}" ref="A1" personId="{PPP-001}" dt="2024-01-15T10:30:00.000">
    <text>Original comment</text>
  </threadedComment>
  <threadedComment id="{BBB-222}" ref="A1" personId="{PPP-002}" dt="2024-01-15T11:00:00.000" parentId="{AAA-111}">
    <text>This is a reply</text>
  </threadedComment>
  <threadedComment id="{CCC-333}" ref="A1" personId="{PPP-001}" dt="2024-01-15T11:30:00.000" parentId="{AAA-111}">
    <text>Another reply</text>
  </threadedComment>
</ThreadedComments>"#;

        let comments = parse_threaded_comments(xml.as_bytes()).unwrap();
        assert_eq!(comments.len(), 3);

        assert!(comments[0].parent_id.is_none());
        assert_eq!(comments[1].parent_id.as_deref(), Some("{AAA-111}"));
        assert_eq!(comments[2].parent_id.as_deref(), Some("{AAA-111}"));

        assert_eq!(comments[0].text, "Original comment");
        assert_eq!(comments[1].text, "This is a reply");
        assert_eq!(comments[2].text, "Another reply");
    }

    #[test]
    fn parse_persons_basic() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<personList xmlns="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments">
  <person id="{PPP-001}" displayName="John Doe" providerId="Windows Live"/>
  <person id="{PPP-002}" displayName="Jane Smith"/>
</personList>"#;

        let persons = parse_persons(xml.as_bytes()).unwrap();
        assert_eq!(persons.len(), 2);

        assert_eq!(persons[0].id, "{PPP-001}");
        assert_eq!(persons[0].display_name, "John Doe");
        assert_eq!(persons[0].provider_id.as_deref(), Some("Windows Live"));

        assert_eq!(persons[1].id, "{PPP-002}");
        assert_eq!(persons[1].display_name, "Jane Smith");
        assert!(persons[1].provider_id.is_none());
    }

    #[test]
    fn threaded_comments_roundtrip() {
        let comments = vec![
            ThreadedCommentData {
                id: "{AAA-111}".to_string(),
                ref_cell: "A1".to_string(),
                person_id: "{PPP-001}".to_string(),
                text: "First comment".to_string(),
                timestamp: "2024-01-15T10:30:00.000".to_string(),
                parent_id: None,
            },
            ThreadedCommentData {
                id: "{BBB-222}".to_string(),
                ref_cell: "A1".to_string(),
                person_id: "{PPP-002}".to_string(),
                text: "Reply with <special> & \"chars\"".to_string(),
                timestamp: "2024-01-15T11:00:00.000".to_string(),
                parent_id: Some("{AAA-111}".to_string()),
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
                id: "{PPP-001}".to_string(),
                display_name: "Alice O'Reilly".to_string(),
                provider_id: Some("Windows Live".to_string()),
            },
            PersonData {
                id: "{PPP-002}".to_string(),
                display_name: "Bob & Carol".to_string(),
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
}
