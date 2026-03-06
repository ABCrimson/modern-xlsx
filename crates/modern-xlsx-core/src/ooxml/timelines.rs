use core::hint::cold_path;

use quick_xml::events::{BytesDecl, BytesEnd, BytesStart, Event};
use quick_xml::{Reader, Writer};
use serde::{Deserialize, Serialize};

use crate::{ModernXlsxError, Result};

const TIMELINE_NS: &str = "http://schemas.microsoft.com/office/spreadsheetml/2014/timeline";

/// Timeline display level granularity.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum TimelineLevel {
    Years,
    Quarters,
    Months,
    Days,
}

impl TimelineLevel {
    #[inline]
    pub fn from_xml(s: &str) -> Option<Self> {
        match s {
            "years" => Some(Self::Years),
            "quarters" => Some(Self::Quarters),
            "months" => Some(Self::Months),
            "days" => Some(Self::Days),
            _ => None,
        }
    }

    #[inline]
    pub fn xml_val(self) -> &'static str {
        match self {
            Self::Years => "years",
            Self::Quarters => "quarters",
            Self::Months => "months",
            Self::Days => "days",
        }
    }
}

/// A timeline definition within a worksheet.
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct TimelineData {
    /// Display name of the timeline.
    pub name: String,
    /// Optional caption text.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub caption: Option<String>,
    /// Name of the associated timeline cache.
    pub cache_name: String,
    /// Source column name.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub source_name: Option<String>,
    /// Display level (years, quarters, months, days).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub level: Option<TimelineLevel>,
}

/// A timeline cache definition at the workbook level.
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct TimelineCacheData {
    /// Unique cache name.
    pub name: String,
    /// Source column name.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub source_name: Option<String>,
    /// Selection start date (ISO 8601 date string).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub selection_start: Option<String>,
    /// Selection end date (ISO 8601 date string).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub selection_end: Option<String>,
}

// ---------------------------------------------------------------------------
// Parsers
// ---------------------------------------------------------------------------

/// Parse a timelines XML file (`xl/timelines/timelineN.xml`) from raw bytes.
///
/// Expected XML:
/// ```xml
/// <timelines xmlns="http://schemas.microsoft.com/office/spreadsheetml/2014/timeline">
///   <timeline name="Timeline_Date" cache="NativeTimeline_Date"
///             caption="Date" sourceName="Date" level="months"/>
/// </timelines>
/// ```
pub fn parse_timelines(data: &[u8]) -> Result<Vec<TimelineData>> {
    let mut reader = Reader::from_reader(data);
    let mut buf = Vec::with_capacity(256);
    let mut timelines: Vec<TimelineData> = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Empty(ref e)) | Ok(Event::Start(ref e)) => {
                if e.local_name().as_ref() == b"timeline" {
                    let mut name = String::new();
                    let mut caption: Option<String> = None;
                    let mut cache_name = String::new();
                    let mut source_name: Option<String> = None;
                    let mut level: Option<TimelineLevel> = None;

                    for attr in e.attributes().flatten() {
                        match attr.key.local_name().as_ref() {
                            b"name" => {
                                name = attr
                                    .unescape_value()
                                    .unwrap_or_default()
                                    .into_owned();
                            }
                            b"cache" => {
                                cache_name = attr
                                    .unescape_value()
                                    .unwrap_or_default()
                                    .into_owned();
                            }
                            b"caption" => {
                                caption = Some(
                                    attr.unescape_value()
                                        .unwrap_or_default()
                                        .into_owned(),
                                );
                            }
                            b"sourceName" => {
                                source_name = Some(
                                    attr.unescape_value()
                                        .unwrap_or_default()
                                        .into_owned(),
                                );
                            }
                            b"level" => {
                                level = TimelineLevel::from_xml(
                                    std::str::from_utf8(&attr.value).unwrap_or_default(),
                                );
                            }
                            _ => {}
                        }
                    }

                    if !name.is_empty() {
                        timelines.push(TimelineData {
                            name,
                            caption,
                            cache_name,
                            source_name,
                            level,
                        });
                    }
                }
            }
            Ok(Event::Eof) => break,
            Ok(_) => {}
            Err(e) => {
                cold_path();
                return Err(ModernXlsxError::XmlParse(format!(
                    "error parsing timelines XML: {e}"
                )));
            }
        }
        buf.clear();
    }

    Ok(timelines)
}

impl TimelineCacheData {
    /// Parse a timeline cache definition XML file from raw bytes.
    ///
    /// Expected XML:
    /// ```xml
    /// <timelineCacheDefinition xmlns="..." name="NativeTimeline_Date" sourceName="Date">
    ///   <state>
    ///     <selection startDate="2024-01-01" endDate="2024-06-30"/>
    ///   </state>
    /// </timelineCacheDefinition>
    /// ```
    pub fn parse(data: &[u8]) -> Result<Self> {
        let mut reader = Reader::from_reader(data);
        let mut buf = Vec::with_capacity(256);

        let mut name = String::new();
        let mut source_name: Option<String> = None;
        let mut selection_start: Option<String> = None;
        let mut selection_end: Option<String> = None;

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) | Ok(Event::Empty(ref e)) => {
                    let local = e.local_name();
                    match local.as_ref() {
                        b"timelineCacheDefinition" => {
                            for attr in e.attributes().flatten() {
                                match attr.key.local_name().as_ref() {
                                    b"name" => {
                                        name = attr
                                            .unescape_value()
                                            .unwrap_or_default()
                                            .into_owned();
                                    }
                                    b"sourceName" => {
                                        source_name = Some(
                                            attr.unescape_value()
                                                .unwrap_or_default()
                                                .into_owned(),
                                        );
                                    }
                                    _ => {}
                                }
                            }
                        }
                        b"selection" => {
                            for attr in e.attributes().flatten() {
                                match attr.key.local_name().as_ref() {
                                    b"startDate" => {
                                        selection_start = Some(
                                            std::str::from_utf8(&attr.value)
                                                .unwrap_or_default()
                                                .to_owned(),
                                        );
                                    }
                                    b"endDate" => {
                                        selection_end = Some(
                                            std::str::from_utf8(&attr.value)
                                                .unwrap_or_default()
                                                .to_owned(),
                                        );
                                    }
                                    _ => {}
                                }
                            }
                        }
                        _ => {}
                    }
                }
                Ok(Event::Eof) => break,
                Ok(_) => {}
                Err(e) => {
                    cold_path();
                    return Err(ModernXlsxError::XmlParse(format!(
                        "error parsing timeline cache XML: {e}"
                    )));
                }
            }
            buf.clear();
        }

        Ok(Self {
            name,
            source_name,
            selection_start,
            selection_end,
        })
    }

    /// Serialize this timeline cache definition to XML bytes.
    pub fn to_xml(&self) -> Result<Vec<u8>> {
        let mut buf: Vec<u8> = Vec::with_capacity(512);
        let mut writer = Writer::new(&mut buf);

        let map_err = |e: std::io::Error| ModernXlsxError::XmlWrite(e.to_string());

        // XML declaration.
        writer
            .write_event(Event::Decl(BytesDecl::new("1.0", Some("UTF-8"), Some("yes"))))
            .map_err(map_err)?;

        // <timelineCacheDefinition xmlns="..." name="..." sourceName="...">
        let mut root = BytesStart::new("timelineCacheDefinition");
        root.push_attribute(("xmlns", TIMELINE_NS));
        root.push_attribute(("name", self.name.as_str()));
        if let Some(ref sn) = self.source_name {
            root.push_attribute(("sourceName", sn.as_str()));
        }

        let has_selection = self.selection_start.is_some() || self.selection_end.is_some();

        if has_selection {
            writer.write_event(Event::Start(root)).map_err(map_err)?;

            // <state>
            writer
                .write_event(Event::Start(BytesStart::new("state")))
                .map_err(map_err)?;

            // <selection startDate="..." endDate="..."/>
            let mut sel = BytesStart::new("selection");
            if let Some(ref start) = self.selection_start {
                sel.push_attribute(("startDate", start.as_str()));
            }
            if let Some(ref end) = self.selection_end {
                sel.push_attribute(("endDate", end.as_str()));
            }
            writer.write_event(Event::Empty(sel)).map_err(map_err)?;

            // </state>
            writer
                .write_event(Event::End(BytesEnd::new("state")))
                .map_err(map_err)?;

            // </timelineCacheDefinition>
            writer
                .write_event(Event::End(BytesEnd::new("timelineCacheDefinition")))
                .map_err(map_err)?;
        } else {
            // Self-closing if no selection state.
            writer.write_event(Event::Empty(root)).map_err(map_err)?;
        }

        Ok(buf)
    }
}

/// Serialize timelines to XML bytes.
pub fn write_timelines(timelines: &[TimelineData]) -> Result<Vec<u8>> {
    let mut buf: Vec<u8> = Vec::with_capacity(256 + timelines.len() * 128);
    let mut writer = Writer::new(&mut buf);

    let map_err = |e: std::io::Error| ModernXlsxError::XmlWrite(e.to_string());

    // XML declaration.
    writer
        .write_event(Event::Decl(BytesDecl::new("1.0", Some("UTF-8"), Some("yes"))))
        .map_err(map_err)?;

    // <timelines xmlns="...">
    let mut root = BytesStart::new("timelines");
    root.push_attribute(("xmlns", TIMELINE_NS));
    writer.write_event(Event::Start(root)).map_err(map_err)?;

    for tl in timelines {
        let mut elem = BytesStart::new("timeline");
        elem.push_attribute(("name", tl.name.as_str()));
        elem.push_attribute(("cache", tl.cache_name.as_str()));

        if let Some(ref caption) = tl.caption {
            elem.push_attribute(("caption", caption.as_str()));
        }
        if let Some(ref sn) = tl.source_name {
            elem.push_attribute(("sourceName", sn.as_str()));
        }
        if let Some(level) = tl.level {
            elem.push_attribute(("level", level.xml_val()));
        }

        writer.write_event(Event::Empty(elem)).map_err(map_err)?;
    }

    // </timelines>
    writer
        .write_event(Event::End(BytesEnd::new("timelines")))
        .map_err(map_err)?;

    Ok(buf)
}

#[cfg(test)]
mod tests {
    use super::*;
    use pretty_assertions::assert_eq;

    #[test]
    fn parse_timelines_basic() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<timelines xmlns="http://schemas.microsoft.com/office/spreadsheetml/2014/timeline">
  <timeline name="Timeline_Date" cache="NativeTimeline_Date" caption="Date" sourceName="Date" level="months"/>
  <timeline name="Timeline_Created" cache="NativeTimeline_Created"/>
</timelines>"#;

        let timelines = parse_timelines(xml.as_bytes()).unwrap();
        assert_eq!(timelines.len(), 2);

        assert_eq!(timelines[0].name, "Timeline_Date");
        assert_eq!(timelines[0].cache_name, "NativeTimeline_Date");
        assert_eq!(timelines[0].caption.as_deref(), Some("Date"));
        assert_eq!(timelines[0].source_name.as_deref(), Some("Date"));
        assert_eq!(timelines[0].level, Some(TimelineLevel::Months));

        assert_eq!(timelines[1].name, "Timeline_Created");
        assert_eq!(timelines[1].cache_name, "NativeTimeline_Created");
        assert!(timelines[1].caption.is_none());
        assert!(timelines[1].level.is_none());
    }

    #[test]
    fn parse_timeline_cache_basic() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<timelineCacheDefinition xmlns="http://schemas.microsoft.com/office/spreadsheetml/2014/timeline"
    name="NativeTimeline_Date" sourceName="Date">
  <state>
    <selection startDate="2024-01-01" endDate="2024-06-30"/>
  </state>
</timelineCacheDefinition>"#;

        let cache = TimelineCacheData::parse(xml.as_bytes()).unwrap();
        assert_eq!(cache.name, "NativeTimeline_Date");
        assert_eq!(cache.source_name.as_deref(), Some("Date"));
        assert_eq!(cache.selection_start.as_deref(), Some("2024-01-01"));
        assert_eq!(cache.selection_end.as_deref(), Some("2024-06-30"));
    }

    #[test]
    fn timelines_roundtrip() {
        let timelines = vec![
            TimelineData {
                name: "Timeline_Date".to_string(),
                caption: Some("Date".to_string()),
                cache_name: "NativeTimeline_Date".to_string(),
                source_name: Some("Date".to_string()),
                level: Some(TimelineLevel::Months),
            },
            TimelineData {
                name: "Timeline_Year".to_string(),
                caption: None,
                cache_name: "NativeTimeline_Year".to_string(),
                source_name: None,
                level: Some(TimelineLevel::Years),
            },
        ];

        let xml = write_timelines(&timelines).unwrap();
        let parsed = parse_timelines(&xml).unwrap();

        assert_eq!(parsed.len(), 2);
        for (orig, round) in timelines.iter().zip(parsed.iter()) {
            assert_eq!(orig.name, round.name);
            assert_eq!(orig.caption, round.caption);
            assert_eq!(orig.cache_name, round.cache_name);
            assert_eq!(orig.source_name, round.source_name);
            assert_eq!(orig.level, round.level);
        }
    }

    #[test]
    fn timeline_cache_roundtrip() {
        let cache = TimelineCacheData {
            name: "NativeTimeline_Date".to_string(),
            source_name: Some("Date".to_string()),
            selection_start: Some("2024-01-01".to_string()),
            selection_end: Some("2024-06-30".to_string()),
        };

        let xml = cache.to_xml().unwrap();
        let parsed = TimelineCacheData::parse(&xml).unwrap();

        assert_eq!(parsed.name, cache.name);
        assert_eq!(parsed.source_name, cache.source_name);
        assert_eq!(parsed.selection_start, cache.selection_start);
        assert_eq!(parsed.selection_end, cache.selection_end);
    }

    #[test]
    fn timeline_cache_no_selection_roundtrip() {
        let cache = TimelineCacheData {
            name: "EmptyCache".to_string(),
            source_name: None,
            selection_start: None,
            selection_end: None,
        };

        let xml = cache.to_xml().unwrap();
        let parsed = TimelineCacheData::parse(&xml).unwrap();

        assert_eq!(parsed.name, "EmptyCache");
        assert!(parsed.source_name.is_none());
        assert!(parsed.selection_start.is_none());
        assert!(parsed.selection_end.is_none());
    }

    #[test]
    fn write_timelines_empty() {
        let timelines: Vec<TimelineData> = Vec::new();
        let xml = write_timelines(&timelines).unwrap();
        let parsed = parse_timelines(&xml).unwrap();
        assert!(parsed.is_empty());
    }

    #[test]
    fn timeline_level_from_xml() {
        assert_eq!(TimelineLevel::from_xml("years"), Some(TimelineLevel::Years));
        assert_eq!(TimelineLevel::from_xml("quarters"), Some(TimelineLevel::Quarters));
        assert_eq!(TimelineLevel::from_xml("months"), Some(TimelineLevel::Months));
        assert_eq!(TimelineLevel::from_xml("days"), Some(TimelineLevel::Days));
        assert_eq!(TimelineLevel::from_xml("unknown"), None);
    }

    #[test]
    fn timeline_level_xml_val() {
        assert_eq!(TimelineLevel::Years.xml_val(), "years");
        assert_eq!(TimelineLevel::Quarters.xml_val(), "quarters");
        assert_eq!(TimelineLevel::Months.xml_val(), "months");
        assert_eq!(TimelineLevel::Days.xml_val(), "days");
    }
}
