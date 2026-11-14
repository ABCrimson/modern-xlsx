//! Timelines and timeline caches — `xl/timelines/timeline{n}.xml` and
//! `xl/timelineCaches/timelineCache{n}.xml`.
//!
//! **Custom dialect, not Excel-verified.** The writer emits its own namespace
//! (`.../2014/timeline`), string levels (`level="months"` and
//! `selectionLevel="months"`) and a `<state>` holding only a `<selection>`,
//! not the MS-XLSX x15 shapes (numeric levels, `<state>` attributes and
//! `<bounds>`), and the extLst links Excel needs to load a timeline are not
//! written. The reader accepts both this dialect and the
//! MS-XLSX x15 parts (namespace `.../2010/11/main`, numeric `level`,
//! `<state>` and `<bounds>`).

use core::hint::cold_path;

use quick_xml::events::{BytesEnd, BytesStart, Event};
use quick_xml::{Reader, Writer};
use serde::{Deserialize, Serialize};

use crate::{ModernXlsxError, Result};

use super::xml::{attr_str, attrs, parse_attr, xml_attr};
use crate::ooxml::{write_xml_decl, xml_write_err};

/// Namespace of the custom dialect the writer emits (not the MS-XLSX x15
/// namespace `http://schemas.microsoft.com/office/spreadsheetml/2010/11/main`).
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
    /// Read a level: the custom dialect's name (`years`, `quarters`,
    /// `months`, `days`) or the MS-XLSX x15 index (`0` Years, `1` Quarters,
    /// `2` Months, `3` Days).
    #[inline]
    pub fn from_xml(s: &str) -> Option<Self> {
        match s {
            "years" | "0" => Some(Self::Years),
            "quarters" | "1" => Some(Self::Quarters),
            "months" | "2" => Some(Self::Months),
            "days" | "3" => Some(Self::Days),
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
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
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
    /// Level of the current selection (x15 `selectionLevel`).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub selection_level: Option<TimelineLevel>,
}

/// A timeline cache definition at the workbook level.
///
/// `pivot_cache_id`, `filter_type`, `bounds_start` and `bounds_end` are read
/// from an MS-XLSX x15 cache (`<state>`, `<bounds>`); the writer's custom
/// dialect does not emit them.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
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
    /// Pivot cache the timeline filters (x15 `<state pivotCacheId>`).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub pivot_cache_id: Option<u32>,
    /// Filter type (x15 `<state filterType>`, for example `dateBetween`).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub filter_type: Option<String>,
    /// Start of the selectable date range (x15 `<bounds startDate>`).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub bounds_start: Option<String>,
    /// End of the selectable date range (x15 `<bounds endDate>`).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub bounds_end: Option<String>,
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
///
/// An MS-XLSX x15 timeline (numeric `level` and `selectionLevel`) is read too.
///
/// # Errors
///
/// `XmlParse` for malformed XML or an attribute holding a bad entity.
pub fn parse_timelines(data: &[u8]) -> Result<Vec<TimelineData>> {
    let mut reader = Reader::from_reader(data);
    let mut buf = Vec::with_capacity(256);
    let mut timelines: Vec<TimelineData> = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Empty(ref e) | Event::Start(ref e)) => {
                if e.local_name().as_ref() == "timeline" {
                    let mut name = String::new();
                    let mut caption: Option<String> = None;
                    let mut cache_name = String::new();
                    let mut source_name: Option<String> = None;
                    let mut level: Option<TimelineLevel> = None;
                    let mut selection_level: Option<TimelineLevel> = None;

                    for attr in attrs(e) {
                        let attr = attr?;
                        match attr.key.local_name().as_ref() {
                            "name" => name = attr_str(&attr)?.into_owned(),
                            "cache" => cache_name = attr_str(&attr)?.into_owned(),
                            "caption" => caption = Some(attr_str(&attr)?.into_owned()),
                            "sourceName" => source_name = Some(attr_str(&attr)?.into_owned()),
                            "level" => level = TimelineLevel::from_xml(&attr_str(&attr)?),
                            "selectionLevel" => {
                                selection_level = TimelineLevel::from_xml(&attr_str(&attr)?);
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
                            selection_level,
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
    ///
    /// An MS-XLSX x15 cache is read too: `<state pivotCacheId filterType>` and
    /// `<bounds startDate endDate>`.
    ///
    /// # Errors
    ///
    /// `XmlParse` for malformed XML, an attribute holding a bad entity, or a
    /// `pivotCacheId` that is not a number.
    pub fn parse(data: &[u8]) -> Result<Self> {
        let mut reader = Reader::from_reader(data);
        let mut buf = Vec::with_capacity(256);

        let mut cache = Self {
            name: String::new(),
            source_name: None,
            selection_start: None,
            selection_end: None,
            pivot_cache_id: None,
            filter_type: None,
            bounds_start: None,
            bounds_end: None,
        };

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e) | Event::Empty(ref e)) => {
                    let local = e.local_name();
                    match local.as_ref() {
                        "timelineCacheDefinition" => {
                            for attr in attrs(e) {
                                let attr = attr?;
                                match attr.key.local_name().as_ref() {
                                    "name" => cache.name = attr_str(&attr)?.into_owned(),
                                    "sourceName" => {
                                        cache.source_name = Some(attr_str(&attr)?.into_owned());
                                    }
                                    _ => {}
                                }
                            }
                        }
                        "state" => {
                            for attr in attrs(e) {
                                let attr = attr?;
                                match attr.key.local_name().as_ref() {
                                    "pivotCacheId" => {
                                        cache.pivot_cache_id = Some(parse_attr("state", &attr)?);
                                    }
                                    "filterType" => {
                                        cache.filter_type = Some(attr_str(&attr)?.into_owned());
                                    }
                                    _ => {}
                                }
                            }
                        }
                        "selection" => {
                            for attr in attrs(e) {
                                let attr = attr?;
                                match attr.key.local_name().as_ref() {
                                    "startDate" => {
                                        cache.selection_start = Some(attr_str(&attr)?.into_owned());
                                    }
                                    "endDate" => {
                                        cache.selection_end = Some(attr_str(&attr)?.into_owned());
                                    }
                                    _ => {}
                                }
                            }
                        }
                        "bounds" => {
                            for attr in attrs(e) {
                                let attr = attr?;
                                match attr.key.local_name().as_ref() {
                                    "startDate" => {
                                        cache.bounds_start = Some(attr_str(&attr)?.into_owned());
                                    }
                                    "endDate" => {
                                        cache.bounds_end = Some(attr_str(&attr)?.into_owned());
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

        Ok(cache)
    }

    /// Serialize this timeline cache definition to XML bytes, in the custom
    /// dialect (see the module docs). The x15-only fields (`pivot_cache_id`,
    /// `filter_type`, `bounds_start`, `bounds_end`) are not written.
    pub fn to_xml(&self) -> Result<Vec<u8>> {
        let mut buf: Vec<u8> = Vec::with_capacity(512);
        let mut writer = Writer::new(&mut buf);

        // XML declaration.
        write_xml_decl(&mut writer)?;

        // <timelineCacheDefinition xmlns="..." name="..." sourceName="...">
        let mut root = BytesStart::new("timelineCacheDefinition");
        root.push_attribute(("xmlns", TIMELINE_NS));
        root.push_attribute(xml_attr("name", &self.name));
        if let Some(ref sn) = self.source_name {
            root.push_attribute(xml_attr("sourceName", sn));
        }

        let has_selection = self.selection_start.is_some() || self.selection_end.is_some();

        if has_selection {
            writer
                .write_event(Event::Start(root))
                .map_err(xml_write_err)?;

            // <state>
            writer
                .write_event(Event::Start(BytesStart::new("state")))
                .map_err(xml_write_err)?;

            // <selection startDate="..." endDate="..."/>
            let mut sel = BytesStart::new("selection");
            if let Some(ref start) = self.selection_start {
                sel.push_attribute(xml_attr("startDate", start));
            }
            if let Some(ref end) = self.selection_end {
                sel.push_attribute(xml_attr("endDate", end));
            }
            writer
                .write_event(Event::Empty(sel))
                .map_err(xml_write_err)?;

            // </state>
            writer
                .write_event(Event::End(BytesEnd::new("state")))
                .map_err(xml_write_err)?;

            // </timelineCacheDefinition>
            writer
                .write_event(Event::End(BytesEnd::new("timelineCacheDefinition")))
                .map_err(xml_write_err)?;
        } else {
            // Self-closing if no selection state.
            writer
                .write_event(Event::Empty(root))
                .map_err(xml_write_err)?;
        }

        Ok(buf)
    }
}

/// Serialize timelines to XML bytes, in the custom dialect (see the module docs).
pub fn write_timelines(timelines: &[TimelineData]) -> Result<Vec<u8>> {
    let mut buf: Vec<u8> = Vec::with_capacity(256 + timelines.len() * 128);
    let mut writer = Writer::new(&mut buf);

    // XML declaration.
    write_xml_decl(&mut writer)?;

    // <timelines xmlns="...">
    let mut root = BytesStart::new("timelines");
    root.push_attribute(("xmlns", TIMELINE_NS));
    writer
        .write_event(Event::Start(root))
        .map_err(xml_write_err)?;

    for tl in timelines {
        let mut elem = BytesStart::new("timeline");
        elem.push_attribute(xml_attr("name", &tl.name));
        elem.push_attribute(xml_attr("cache", &tl.cache_name));

        if let Some(ref caption) = tl.caption {
            elem.push_attribute(xml_attr("caption", caption));
        }
        if let Some(ref sn) = tl.source_name {
            elem.push_attribute(xml_attr("sourceName", sn));
        }
        if let Some(level) = tl.level {
            elem.push_attribute(("level", level.xml_val()));
        }
        if let Some(level) = tl.selection_level {
            elem.push_attribute(("selectionLevel", level.xml_val()));
        }

        writer
            .write_event(Event::Empty(elem))
            .map_err(xml_write_err)?;
    }

    // </timelines>
    writer
        .write_event(Event::End(BytesEnd::new("timelines")))
        .map_err(xml_write_err)?;

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
                selection_level: None,
            },
            TimelineData {
                name: "Timeline_Year".to_string(),
                caption: None,
                cache_name: "NativeTimeline_Year".to_string(),
                source_name: None,
                level: Some(TimelineLevel::Years),
                selection_level: None,
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
            pivot_cache_id: None,
            filter_type: None,
            bounds_start: None,
            bounds_end: None,
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
            pivot_cache_id: None,
            filter_type: None,
            bounds_start: None,
            bounds_end: None,
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
        assert_eq!(
            TimelineLevel::from_xml("quarters"),
            Some(TimelineLevel::Quarters)
        );
        assert_eq!(
            TimelineLevel::from_xml("months"),
            Some(TimelineLevel::Months)
        );
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

    const X15: &str = "http://schemas.microsoft.com/office/spreadsheetml/2010/11/main";

    #[test]
    fn x15_numeric_level_is_read() {
        let xml = format!(
            r#"<timelines xmlns="{X15}"><timeline name="T" cache="NativeTimeline_D" caption="D" level="2" selectionLevel="3"/></timelines>"#
        );
        let t = parse_timelines(xml.as_bytes()).unwrap();
        assert_eq!(t[0].level, Some(TimelineLevel::Months));
        assert_eq!(t[0].selection_level, Some(TimelineLevel::Days));
    }

    #[test]
    fn numeric_levels_map_to_variants() {
        assert_eq!(TimelineLevel::from_xml("0"), Some(TimelineLevel::Years));
        assert_eq!(TimelineLevel::from_xml("1"), Some(TimelineLevel::Quarters));
        assert_eq!(TimelineLevel::from_xml("2"), Some(TimelineLevel::Months));
        assert_eq!(TimelineLevel::from_xml("3"), Some(TimelineLevel::Days));
        assert_eq!(TimelineLevel::from_xml("4"), None);
    }

    #[test]
    fn x15_cache_state_and_bounds_are_read() {
        let xml = format!(
            r#"<timelineCacheDefinition xmlns="{X15}" name="NativeTimeline_D" sourceName="Date"><state name="NativeTimeline_D" pivotCacheId="7" filterType="dateBetween" minimalRefreshVersion="6" lastRefreshVersion="6"><selection startDate="2024-01-01T00:00:00" endDate="2024-03-31T00:00:00"/><bounds startDate="2023-01-01T00:00:00" endDate="2025-01-01T00:00:00"/></state></timelineCacheDefinition>"#
        );
        let c = TimelineCacheData::parse(xml.as_bytes()).unwrap();
        assert_eq!(c.pivot_cache_id, Some(7));
        assert_eq!(c.filter_type.as_deref(), Some("dateBetween"));
        assert_eq!(c.selection_start.as_deref(), Some("2024-01-01T00:00:00"));
        assert_eq!(c.bounds_start.as_deref(), Some("2023-01-01T00:00:00"));
        assert_eq!(c.bounds_end.as_deref(), Some("2025-01-01T00:00:00"));
    }

    #[test]
    fn start_only_selection_round_trips() {
        let cache = TimelineCacheData {
            name: "C".into(),
            source_name: None,
            selection_start: Some("2024-01-01".into()),
            selection_end: None,
            pivot_cache_id: None,
            filter_type: None,
            bounds_start: None,
            bounds_end: None,
        };
        let parsed = TimelineCacheData::parse(&cache.to_xml().unwrap()).unwrap();
        assert_eq!(parsed, cache);
    }

    #[test]
    fn selection_level_round_trips_in_the_custom_dialect() {
        let tl = vec![TimelineData {
            name: "T".into(),
            caption: None,
            cache_name: "C".into(),
            source_name: None,
            level: Some(TimelineLevel::Years),
            selection_level: Some(TimelineLevel::Quarters),
        }];
        assert_eq!(parse_timelines(&write_timelines(&tl).unwrap()).unwrap(), tl);
    }

    #[test]
    fn bad_entity_is_xml_parse_error() {
        let xml = br#"<timelines><timeline name="A &bad; B" cache="c"/></timelines>"#;
        assert!(matches!(
            parse_timelines(xml),
            Err(ModernXlsxError::XmlParse(_))
        ));
        let xml = br#"<timelineCacheDefinition name="A &bad; B"/>"#;
        assert!(matches!(
            TimelineCacheData::parse(xml),
            Err(ModernXlsxError::XmlParse(_))
        ));
    }

    #[test]
    fn ampersand_name_reads_decoded_and_writes_one_amp() {
        let xml = br#"<timelines><timeline name="A &amp; B" cache="c"/></timelines>"#;
        let tl = parse_timelines(xml).unwrap();
        assert_eq!(tl[0].name, "A & B");
        let out = write_timelines(&tl).unwrap();
        let s = std::str::from_utf8(&out).unwrap();
        assert!(s.contains(r#"name="A &amp; B""#), "{s}");
        assert!(!s.contains("&amp;amp;"), "{s}");
    }

    #[test]
    fn mismatched_end_tag_is_xml_parse_error() {
        assert!(matches!(
            parse_timelines(b"<root><a></b></root>"),
            Err(ModernXlsxError::XmlParse(_))
        ));
        assert!(matches!(
            TimelineCacheData::parse(b"<root><a></b></root>"),
            Err(ModernXlsxError::XmlParse(_))
        ));
    }
}
