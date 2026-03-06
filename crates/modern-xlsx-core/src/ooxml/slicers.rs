use core::hint::cold_path;

use quick_xml::events::{BytesDecl, BytesEnd, BytesStart, Event};
use quick_xml::{Reader, Writer};
use serde::{Deserialize, Serialize};

use crate::{ModernXlsxError, Result};

const SLICER_NS: &str = "http://schemas.microsoft.com/office/spreadsheetml/2010/slicer";
const MC_NS: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";

/// Sort order for slicer items.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum SortOrder {
    Ascending,
    Descending,
}

impl SortOrder {
    #[inline]
    pub fn from_xml(s: &str) -> Option<Self> {
        match s {
            "ascending" => Some(Self::Ascending),
            "descending" => Some(Self::Descending),
            _ => None,
        }
    }

    #[inline]
    pub fn xml_val(self) -> &'static str {
        match self {
            Self::Ascending => "ascending",
            Self::Descending => "descending",
        }
    }
}

/// A slicer definition within a worksheet.
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct SlicerData {
    /// Display name of the slicer.
    pub name: String,
    /// Optional caption text.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub caption: Option<String>,
    /// Name of the associated slicer cache.
    pub cache_name: String,
    /// Source column name.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub column_name: Option<String>,
    /// Item sort order.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub sort_order: Option<SortOrder>,
    /// First visible item index.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub start_item: Option<u32>,
}

/// A slicer cache definition at the workbook level.
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct SlicerCacheData {
    /// Unique cache name.
    pub name: String,
    /// Source column name in the data source.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub source_name: Option<String>,
    /// Cached items with selection state.
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub items: Vec<SlicerItem>,
}

/// A single item in a slicer cache.
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct SlicerItem {
    /// Display name.
    pub n: String,
    /// Whether the item is selected.
    #[serde(default, skip_serializing_if = "crate::ooxml::is_false")]
    pub s: bool,
}

// ---------------------------------------------------------------------------
// Parsers
// ---------------------------------------------------------------------------

/// Parse a slicers XML file (`xl/slicers/slicerN.xml`) from raw bytes.
///
/// Expected XML:
/// ```xml
/// <slicers xmlns="http://schemas.microsoft.com/office/spreadsheetml/2010/slicer"
///          xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">
///   <slicer name="Slicer_Region" cache="Slicer_Region" caption="Region"
///           columnName="Region" sortOrder="ascending" startItem="0"/>
/// </slicers>
/// ```
pub fn parse_slicers(data: &[u8]) -> Result<Vec<SlicerData>> {
    let mut reader = Reader::from_reader(data);
    let mut buf = Vec::with_capacity(256);
    let mut slicers: Vec<SlicerData> = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Empty(ref e)) | Ok(Event::Start(ref e)) => {
                if e.local_name().as_ref() == b"slicer" {
                    let mut name = String::new();
                    let mut caption: Option<String> = None;
                    let mut cache_name = String::new();
                    let mut column_name: Option<String> = None;
                    let mut sort_order: Option<SortOrder> = None;
                    let mut start_item: Option<u32> = None;

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
                            b"columnName" => {
                                column_name = Some(
                                    attr.unescape_value()
                                        .unwrap_or_default()
                                        .into_owned(),
                                );
                            }
                            b"sortOrder" => {
                                sort_order = SortOrder::from_xml(
                                    std::str::from_utf8(&attr.value).unwrap_or_default(),
                                );
                            }
                            b"startItem" => {
                                start_item = std::str::from_utf8(&attr.value)
                                    .ok()
                                    .and_then(|s| s.parse().ok());
                            }
                            _ => {}
                        }
                    }

                    if !name.is_empty() {
                        slicers.push(SlicerData {
                            name,
                            caption,
                            cache_name,
                            column_name,
                            sort_order,
                            start_item,
                        });
                    }
                }
            }
            Ok(Event::Eof) => break,
            Ok(_) => {}
            Err(e) => {
                cold_path();
                return Err(ModernXlsxError::XmlParse(format!(
                    "error parsing slicers XML: {e}"
                )));
            }
        }
        buf.clear();
    }

    Ok(slicers)
}

impl SlicerCacheData {
    /// Parse a slicer cache definition XML file from raw bytes.
    ///
    /// Expected XML:
    /// ```xml
    /// <slicerCacheDefinition xmlns="..." name="Slicer_Region" sourceName="Region">
    ///   <data>
    ///     <tabularData>
    ///       <items>
    ///         <i n="East" s="1"/>
    ///         <i n="West"/>
    ///       </items>
    ///     </tabularData>
    ///   </data>
    /// </slicerCacheDefinition>
    /// ```
    pub fn parse(data: &[u8]) -> Result<Self> {
        let mut reader = Reader::from_reader(data);
        let mut buf = Vec::with_capacity(256);

        let mut name = String::new();
        let mut source_name: Option<String> = None;
        let mut items: Vec<SlicerItem> = Vec::new();

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) | Ok(Event::Empty(ref e)) => {
                    let local = e.local_name();
                    match local.as_ref() {
                        b"slicerCacheDefinition" => {
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
                        b"i" => {
                            let mut item_n = String::new();
                            let mut item_s = false;

                            for attr in e.attributes().flatten() {
                                match attr.key.local_name().as_ref() {
                                    b"n" => {
                                        item_n = attr
                                            .unescape_value()
                                            .unwrap_or_default()
                                            .into_owned();
                                    }
                                    b"s" => {
                                        let val =
                                            std::str::from_utf8(&attr.value).unwrap_or_default();
                                        item_s = val == "1" || val == "true";
                                    }
                                    _ => {}
                                }
                            }

                            if !item_n.is_empty() {
                                items.push(SlicerItem { n: item_n, s: item_s });
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
                        "error parsing slicer cache XML: {e}"
                    )));
                }
            }
            buf.clear();
        }

        Ok(Self {
            name,
            source_name,
            items,
        })
    }

    /// Serialize this slicer cache definition to XML bytes.
    pub fn to_xml(&self) -> Result<Vec<u8>> {
        let mut buf: Vec<u8> = Vec::with_capacity(512 + self.items.len() * 64);
        let mut writer = Writer::new(&mut buf);

        let map_err = |e: std::io::Error| ModernXlsxError::XmlWrite(e.to_string());

        // XML declaration.
        writer
            .write_event(Event::Decl(BytesDecl::new("1.0", Some("UTF-8"), Some("yes"))))
            .map_err(map_err)?;

        // <slicerCacheDefinition xmlns="..." name="..." sourceName="...">
        let mut root = BytesStart::new("slicerCacheDefinition");
        root.push_attribute(("xmlns", SLICER_NS));
        root.push_attribute(("name", self.name.as_str()));
        if let Some(ref sn) = self.source_name {
            root.push_attribute(("sourceName", sn.as_str()));
        }
        writer.write_event(Event::Start(root)).map_err(map_err)?;

        if !self.items.is_empty() {
            // <data><tabularData><items>
            writer
                .write_event(Event::Start(BytesStart::new("data")))
                .map_err(map_err)?;
            writer
                .write_event(Event::Start(BytesStart::new("tabularData")))
                .map_err(map_err)?;
            writer
                .write_event(Event::Start(BytesStart::new("items")))
                .map_err(map_err)?;

            for item in &self.items {
                let mut elem = BytesStart::new("i");
                elem.push_attribute(("n", item.n.as_str()));
                if item.s {
                    elem.push_attribute(("s", "1"));
                }
                writer.write_event(Event::Empty(elem)).map_err(map_err)?;
            }

            // </items></tabularData></data>
            writer
                .write_event(Event::End(BytesEnd::new("items")))
                .map_err(map_err)?;
            writer
                .write_event(Event::End(BytesEnd::new("tabularData")))
                .map_err(map_err)?;
            writer
                .write_event(Event::End(BytesEnd::new("data")))
                .map_err(map_err)?;
        }

        // </slicerCacheDefinition>
        writer
            .write_event(Event::End(BytesEnd::new("slicerCacheDefinition")))
            .map_err(map_err)?;

        Ok(buf)
    }
}

/// Serialize slicers to XML bytes.
pub fn write_slicers(slicers: &[SlicerData]) -> Result<Vec<u8>> {
    let mut buf: Vec<u8> = Vec::with_capacity(256 + slicers.len() * 128);
    let mut writer = Writer::new(&mut buf);

    let map_err = |e: std::io::Error| ModernXlsxError::XmlWrite(e.to_string());

    // XML declaration.
    writer
        .write_event(Event::Decl(BytesDecl::new("1.0", Some("UTF-8"), Some("yes"))))
        .map_err(map_err)?;

    // <slicers xmlns="..." xmlns:mc="...">
    let mut root = BytesStart::new("slicers");
    root.push_attribute(("xmlns", SLICER_NS));
    root.push_attribute(("xmlns:mc", MC_NS));
    writer.write_event(Event::Start(root)).map_err(map_err)?;

    let mut ibuf = itoa::Buffer::new();
    for slicer in slicers {
        let mut elem = BytesStart::new("slicer");
        elem.push_attribute(("name", slicer.name.as_str()));
        elem.push_attribute(("cache", slicer.cache_name.as_str()));

        if let Some(ref caption) = slicer.caption {
            elem.push_attribute(("caption", caption.as_str()));
        }
        if let Some(ref col) = slicer.column_name {
            elem.push_attribute(("columnName", col.as_str()));
        }
        if let Some(so) = slicer.sort_order {
            elem.push_attribute(("sortOrder", so.xml_val()));
        }
        if let Some(si) = slicer.start_item {
            elem.push_attribute(("startItem", ibuf.format(si)));
        }

        writer.write_event(Event::Empty(elem)).map_err(map_err)?;
    }

    // </slicers>
    writer
        .write_event(Event::End(BytesEnd::new("slicers")))
        .map_err(map_err)?;

    Ok(buf)
}

#[cfg(test)]
mod tests {
    use super::*;
    use pretty_assertions::assert_eq;

    #[test]
    fn parse_slicers_basic() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<slicers xmlns="http://schemas.microsoft.com/office/spreadsheetml/2010/slicer"
         xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">
  <slicer name="Slicer_Region" cache="Slicer_Region" caption="Region" columnName="Region" sortOrder="ascending" startItem="0"/>
  <slicer name="Slicer_Product" cache="Slicer_Product"/>
</slicers>"#;

        let slicers = parse_slicers(xml.as_bytes()).unwrap();
        assert_eq!(slicers.len(), 2);

        assert_eq!(slicers[0].name, "Slicer_Region");
        assert_eq!(slicers[0].cache_name, "Slicer_Region");
        assert_eq!(slicers[0].caption.as_deref(), Some("Region"));
        assert_eq!(slicers[0].column_name.as_deref(), Some("Region"));
        assert_eq!(slicers[0].sort_order, Some(SortOrder::Ascending));
        assert_eq!(slicers[0].start_item, Some(0));

        assert_eq!(slicers[1].name, "Slicer_Product");
        assert_eq!(slicers[1].cache_name, "Slicer_Product");
        assert!(slicers[1].caption.is_none());
        assert!(slicers[1].sort_order.is_none());
    }

    #[test]
    fn parse_slicer_cache_basic() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<slicerCacheDefinition xmlns="http://schemas.microsoft.com/office/spreadsheetml/2010/slicer"
    name="Slicer_Region" sourceName="Region">
  <data>
    <tabularData>
      <items>
        <i n="East" s="1"/>
        <i n="West"/>
        <i n="North" s="true"/>
      </items>
    </tabularData>
  </data>
</slicerCacheDefinition>"#;

        let cache = SlicerCacheData::parse(xml.as_bytes()).unwrap();
        assert_eq!(cache.name, "Slicer_Region");
        assert_eq!(cache.source_name.as_deref(), Some("Region"));
        assert_eq!(cache.items.len(), 3);
        assert_eq!(cache.items[0].n, "East");
        assert!(cache.items[0].s);
        assert_eq!(cache.items[1].n, "West");
        assert!(!cache.items[1].s);
        assert_eq!(cache.items[2].n, "North");
        assert!(cache.items[2].s);
    }

    #[test]
    fn slicers_roundtrip() {
        let slicers = vec![
            SlicerData {
                name: "Slicer_Region".to_string(),
                caption: Some("Region".to_string()),
                cache_name: "Slicer_Region".to_string(),
                column_name: Some("Region".to_string()),
                sort_order: Some(SortOrder::Descending),
                start_item: Some(2),
            },
            SlicerData {
                name: "Slicer_Product".to_string(),
                caption: None,
                cache_name: "Slicer_Product".to_string(),
                column_name: None,
                sort_order: None,
                start_item: None,
            },
        ];

        let xml = write_slicers(&slicers).unwrap();
        let parsed = parse_slicers(&xml).unwrap();

        assert_eq!(parsed.len(), 2);
        for (orig, round) in slicers.iter().zip(parsed.iter()) {
            assert_eq!(orig.name, round.name);
            assert_eq!(orig.caption, round.caption);
            assert_eq!(orig.cache_name, round.cache_name);
            assert_eq!(orig.column_name, round.column_name);
            assert_eq!(orig.sort_order, round.sort_order);
            assert_eq!(orig.start_item, round.start_item);
        }
    }

    #[test]
    fn slicer_cache_roundtrip() {
        let cache = SlicerCacheData {
            name: "Slicer_Region".to_string(),
            source_name: Some("Region".to_string()),
            items: vec![
                SlicerItem {
                    n: "East".to_string(),
                    s: true,
                },
                SlicerItem {
                    n: "West & South".to_string(),
                    s: false,
                },
            ],
        };

        let xml = cache.to_xml().unwrap();
        let parsed = SlicerCacheData::parse(&xml).unwrap();

        assert_eq!(parsed.name, cache.name);
        assert_eq!(parsed.source_name, cache.source_name);
        assert_eq!(parsed.items.len(), 2);
        for (orig, round) in cache.items.iter().zip(parsed.items.iter()) {
            assert_eq!(orig.n, round.n);
            assert_eq!(orig.s, round.s);
        }
    }

    #[test]
    fn write_slicers_empty() {
        let slicers: Vec<SlicerData> = Vec::new();
        let xml = write_slicers(&slicers).unwrap();
        let parsed = parse_slicers(&xml).unwrap();
        assert!(parsed.is_empty());
    }

    #[test]
    fn slicer_cache_empty_items() {
        let cache = SlicerCacheData {
            name: "Empty".to_string(),
            source_name: None,
            items: Vec::new(),
        };

        let xml = cache.to_xml().unwrap();
        let parsed = SlicerCacheData::parse(&xml).unwrap();
        assert_eq!(parsed.name, "Empty");
        assert!(parsed.source_name.is_none());
        assert!(parsed.items.is_empty());
    }

    #[test]
    fn sort_order_from_xml() {
        assert_eq!(SortOrder::from_xml("ascending"), Some(SortOrder::Ascending));
        assert_eq!(SortOrder::from_xml("descending"), Some(SortOrder::Descending));
        assert_eq!(SortOrder::from_xml("unknown"), None);
    }

    #[test]
    fn sort_order_xml_val() {
        assert_eq!(SortOrder::Ascending.xml_val(), "ascending");
        assert_eq!(SortOrder::Descending.xml_val(), "descending");
    }
}
