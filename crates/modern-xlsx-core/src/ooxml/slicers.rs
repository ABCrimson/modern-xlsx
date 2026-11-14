//! Slicers and slicer caches — `xl/slicers/slicer{n}.xml` and
//! `xl/slicerCaches/slicerCache{n}.xml`.
//!
//! **Custom dialect, not Excel-verified.** The writer emits its own namespace
//! (`.../2010/slicer`), `<tabularData>` and `<i n=..>` items (with `x`, `s` and
//! `nd` attributes when set), not the MS-XLSX x14 shapes, and the extLst,
//! defined-name and drawing links Excel needs to load a slicer are not
//! written. The reader accepts both this dialect and the
//! MS-XLSX x14 caches (namespace `.../2009/9/main`, `<tabular pivotCacheId>`,
//! items `<i x s nd>`).

use core::hint::cold_path;

use quick_xml::events::{BytesEnd, BytesStart, Event};
use quick_xml::{Reader, Writer};
use serde::{Deserialize, Serialize};

use crate::{ModernXlsxError, Result};

use super::xml::{attr_str, attrs, parse_attr, xml_attr, xsd_bool};
use crate::ooxml::{write_xml_decl, xml_write_err};

/// Namespace of the custom dialect the writer emits (not the MS-XLSX x14
/// namespace `http://schemas.microsoft.com/office/spreadsheetml/2009/9/main`).
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
///
/// `pivot_cache_id`, `sort_order`, `cross_filter` and `pivot_tables` are read
/// from an MS-XLSX x14 cache (`<tabular>`, `<pivotTables>`); the writer's
/// custom dialect does not emit them.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
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
    /// Pivot cache the slicer filters (x14 `<tabular pivotCacheId>`).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub pivot_cache_id: Option<u32>,
    /// Item sort order (x14 `<tabular sortOrder>`).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub sort_order: Option<SortOrder>,
    /// How items without data are shown (x14 `<tabular crossFilter>`, for
    /// example `showItemsWithDataAtTop`).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cross_filter: Option<String>,
    /// Pivot tables the cache is connected to (x14 `<pivotTables>`).
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub pivot_tables: Vec<SlicerPivotTable>,
}

/// A pivot table a slicer cache is connected to (x14 `<pivotTable>`).
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct SlicerPivotTable {
    /// Index of the worksheet that holds the pivot table.
    pub tab_id: u32,
    /// Name of the pivot table.
    pub name: String,
}

/// A single item in a slicer cache.
///
/// The custom dialect names items (`n`); MS-XLSX x14 caches index them into
/// the pivot field's shared items (`x`).
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct SlicerItem {
    /// Display name (empty for an x14 item that only carries `x`).
    #[serde(default)]
    pub n: String,
    /// Whether the item is selected.
    #[serde(default, skip_serializing_if = "crate::ooxml::is_false")]
    pub s: bool,
    /// Index of the item in the pivot field's shared items (x14 `x`).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub x: Option<u32>,
    /// Whether the item has no data (x14 `nd`).
    #[serde(default, skip_serializing_if = "crate::ooxml::is_false")]
    pub nd: bool,
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
///
/// # Errors
///
/// `XmlParse` for malformed XML, an attribute holding a bad entity, or a
/// `startItem` that is not a number.
pub fn parse_slicers(data: &[u8]) -> Result<Vec<SlicerData>> {
    let mut reader = Reader::from_reader(data);
    let mut buf = Vec::with_capacity(256);
    let mut slicers: Vec<SlicerData> = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Empty(ref e) | Event::Start(ref e)) => {
                if e.local_name().as_ref() == "slicer" {
                    let mut name = String::new();
                    let mut caption: Option<String> = None;
                    let mut cache_name = String::new();
                    let mut column_name: Option<String> = None;
                    let mut sort_order: Option<SortOrder> = None;
                    let mut start_item: Option<u32> = None;

                    for attr in attrs(e) {
                        let attr = attr?;
                        match attr.key.local_name().as_ref() {
                            "name" => name = attr_str(&attr)?.into_owned(),
                            "cache" => cache_name = attr_str(&attr)?.into_owned(),
                            "caption" => caption = Some(attr_str(&attr)?.into_owned()),
                            "columnName" => column_name = Some(attr_str(&attr)?.into_owned()),
                            "sortOrder" => sort_order = SortOrder::from_xml(&attr_str(&attr)?),
                            "startItem" => start_item = Some(parse_attr("slicer", &attr)?),
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
    ///
    /// An MS-XLSX x14 cache is read too: `<tabular pivotCacheId sortOrder
    /// crossFilter>` with `<i x s nd/>` items, and `<pivotTables>`.
    ///
    /// # Errors
    ///
    /// `XmlParse` for malformed XML, an attribute holding a bad entity, or a
    /// numeric attribute that is not a valid number.
    pub fn parse(data: &[u8]) -> Result<Self> {
        let mut reader = Reader::from_reader(data);
        let mut buf = Vec::with_capacity(256);

        let mut cache = Self {
            name: String::new(),
            source_name: None,
            items: Vec::new(),
            pivot_cache_id: None,
            sort_order: None,
            cross_filter: None,
            pivot_tables: Vec::new(),
        };

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e) | Event::Empty(ref e)) => {
                    let local = e.local_name();
                    match local.as_ref() {
                        "slicerCacheDefinition" => {
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
                        "tabular" => {
                            for attr in attrs(e) {
                                let attr = attr?;
                                match attr.key.local_name().as_ref() {
                                    "pivotCacheId" => {
                                        cache.pivot_cache_id = Some(parse_attr("tabular", &attr)?);
                                    }
                                    "sortOrder" => {
                                        cache.sort_order = SortOrder::from_xml(&attr_str(&attr)?);
                                    }
                                    "crossFilter" => {
                                        cache.cross_filter = Some(attr_str(&attr)?.into_owned());
                                    }
                                    _ => {}
                                }
                            }
                        }
                        "pivotTable" => {
                            let mut pt = SlicerPivotTable {
                                tab_id: 0,
                                name: String::new(),
                            };
                            for attr in attrs(e) {
                                let attr = attr?;
                                match attr.key.local_name().as_ref() {
                                    "tabId" => pt.tab_id = parse_attr("pivotTable", &attr)?,
                                    "name" => pt.name = attr_str(&attr)?.into_owned(),
                                    _ => {}
                                }
                            }
                            cache.pivot_tables.push(pt);
                        }
                        "i" => {
                            let mut item = SlicerItem {
                                n: String::new(),
                                s: false,
                                x: None,
                                nd: false,
                            };
                            for attr in attrs(e) {
                                let attr = attr?;
                                match attr.key.local_name().as_ref() {
                                    "n" => item.n = attr_str(&attr)?.into_owned(),
                                    "s" => item.s = xsd_bool(&attr_str(&attr)?).unwrap_or(false),
                                    "x" => item.x = Some(parse_attr("i", &attr)?),
                                    "nd" => item.nd = xsd_bool(&attr_str(&attr)?).unwrap_or(false),
                                    _ => {}
                                }
                            }
                            if !item.n.is_empty() || item.x.is_some() {
                                cache.items.push(item);
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

        Ok(cache)
    }

    /// Serialize this slicer cache definition to XML bytes, in the custom
    /// dialect (see the module docs). The x14-only fields (`pivot_cache_id`,
    /// `sort_order`, `cross_filter`, `pivot_tables`) are not written.
    pub fn to_xml(&self) -> Result<Vec<u8>> {
        let mut buf: Vec<u8> = Vec::with_capacity(512 + self.items.len() * 64);
        let mut writer = Writer::new(&mut buf);
        let mut ibuf = itoa::Buffer::new();

        // XML declaration.
        write_xml_decl(&mut writer)?;

        // <slicerCacheDefinition xmlns="..." name="..." sourceName="...">
        let mut root = BytesStart::new("slicerCacheDefinition");
        root.push_attribute(("xmlns", SLICER_NS));
        root.push_attribute(xml_attr("name", &self.name));
        if let Some(ref sn) = self.source_name {
            root.push_attribute(xml_attr("sourceName", sn));
        }
        writer
            .write_event(Event::Start(root))
            .map_err(xml_write_err)?;

        if !self.items.is_empty() {
            // <data><tabularData><items>
            writer
                .write_event(Event::Start(BytesStart::new("data")))
                .map_err(xml_write_err)?;
            writer
                .write_event(Event::Start(BytesStart::new("tabularData")))
                .map_err(xml_write_err)?;
            writer
                .write_event(Event::Start(BytesStart::new("items")))
                .map_err(xml_write_err)?;

            for item in &self.items {
                let mut elem = BytesStart::new("i");
                if !item.n.is_empty() {
                    elem.push_attribute(xml_attr("n", &item.n));
                }
                if let Some(x) = item.x {
                    elem.push_attribute(("x", ibuf.format(x)));
                }
                if item.s {
                    elem.push_attribute(("s", "1"));
                }
                if item.nd {
                    elem.push_attribute(("nd", "1"));
                }
                writer
                    .write_event(Event::Empty(elem))
                    .map_err(xml_write_err)?;
            }

            // </items></tabularData></data>
            writer
                .write_event(Event::End(BytesEnd::new("items")))
                .map_err(xml_write_err)?;
            writer
                .write_event(Event::End(BytesEnd::new("tabularData")))
                .map_err(xml_write_err)?;
            writer
                .write_event(Event::End(BytesEnd::new("data")))
                .map_err(xml_write_err)?;
        }

        // </slicerCacheDefinition>
        writer
            .write_event(Event::End(BytesEnd::new("slicerCacheDefinition")))
            .map_err(xml_write_err)?;

        Ok(buf)
    }
}

/// Serialize slicers to XML bytes, in the custom dialect (see the module docs).
pub fn write_slicers(slicers: &[SlicerData]) -> Result<Vec<u8>> {
    let mut buf: Vec<u8> = Vec::with_capacity(256 + slicers.len() * 128);
    let mut writer = Writer::new(&mut buf);

    // XML declaration.
    write_xml_decl(&mut writer)?;

    // <slicers xmlns="..." xmlns:mc="...">
    let mut root = BytesStart::new("slicers");
    root.push_attribute(("xmlns", SLICER_NS));
    root.push_attribute(("xmlns:mc", MC_NS));
    writer
        .write_event(Event::Start(root))
        .map_err(xml_write_err)?;

    let mut ibuf = itoa::Buffer::new();
    for slicer in slicers {
        let mut elem = BytesStart::new("slicer");
        elem.push_attribute(xml_attr("name", &slicer.name));
        elem.push_attribute(xml_attr("cache", &slicer.cache_name));

        if let Some(ref caption) = slicer.caption {
            elem.push_attribute(xml_attr("caption", caption));
        }
        if let Some(ref col) = slicer.column_name {
            elem.push_attribute(xml_attr("columnName", col));
        }
        if let Some(so) = slicer.sort_order {
            elem.push_attribute(("sortOrder", so.xml_val()));
        }
        if let Some(si) = slicer.start_item {
            elem.push_attribute(("startItem", ibuf.format(si)));
        }

        writer
            .write_event(Event::Empty(elem))
            .map_err(xml_write_err)?;
    }

    // </slicers>
    writer
        .write_event(Event::End(BytesEnd::new("slicers")))
        .map_err(xml_write_err)?;

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
            pivot_cache_id: None,
            sort_order: None,
            cross_filter: None,
            pivot_tables: Vec::new(),
            items: vec![
                SlicerItem {
                    n: "East".to_string(),
                    s: true,
                    x: None,
                    nd: false,
                },
                SlicerItem {
                    n: "West & South".to_string(),
                    s: false,
                    x: None,
                    nd: false,
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
            pivot_cache_id: None,
            sort_order: None,
            cross_filter: None,
            pivot_tables: Vec::new(),
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
        assert_eq!(
            SortOrder::from_xml("descending"),
            Some(SortOrder::Descending)
        );
        assert_eq!(SortOrder::from_xml("unknown"), None);
    }

    #[test]
    fn sort_order_xml_val() {
        assert_eq!(SortOrder::Ascending.xml_val(), "ascending");
        assert_eq!(SortOrder::Descending.xml_val(), "descending");
    }

    const X14: &str = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main";

    #[test]
    fn x14_cache_items_pivot_cache_and_tables_are_read() {
        let xml = format!(
            r#"<slicerCacheDefinition xmlns="{X14}" name="Slicer_Region" sourceName="Region"><pivotTables><pivotTable tabId="2" name="PivotTable1"/></pivotTables><data><tabular pivotCacheId="1" sortOrder="descending" crossFilter="showItemsWithDataAtTop"><items><i x="0" s="1"/><i x="1"/><i x="2" nd="1"/></items></tabular></data></slicerCacheDefinition>"#
        );
        let c = SlicerCacheData::parse(xml.as_bytes()).unwrap();
        assert_eq!(c.items.len(), 3);
        assert_eq!(
            c.items.iter().map(|i| i.x).collect::<Vec<_>>(),
            vec![Some(0), Some(1), Some(2)]
        );
        assert!(c.items[0].s);
        assert!(!c.items[1].s);
        assert!(c.items[2].nd);
        assert_eq!(c.pivot_cache_id, Some(1));
        assert_eq!(c.sort_order, Some(SortOrder::Descending));
        assert_eq!(c.cross_filter.as_deref(), Some("showItemsWithDataAtTop"));
        assert_eq!(
            c.pivot_tables,
            vec![SlicerPivotTable {
                tab_id: 2,
                name: "PivotTable1".into()
            }]
        );
    }

    #[test]
    fn x_only_items_round_trip_through_the_custom_dialect() {
        let cache = SlicerCacheData {
            name: "C".into(),
            source_name: None,
            items: vec![
                SlicerItem {
                    n: String::new(),
                    s: true,
                    x: Some(4),
                    nd: true,
                },
                SlicerItem {
                    n: "East".into(),
                    s: false,
                    x: None,
                    nd: false,
                },
            ],
            pivot_cache_id: None,
            sort_order: None,
            cross_filter: None,
            pivot_tables: Vec::new(),
        };
        let parsed = SlicerCacheData::parse(&cache.to_xml().unwrap()).unwrap();
        assert_eq!(parsed, cache);
    }

    #[test]
    fn slicer_item_without_n_deserializes() {
        let item: SlicerItem = serde_json::from_str(r#"{"x":3}"#).unwrap();
        assert_eq!(item.n, "");
        assert_eq!(item.x, Some(3));
    }

    #[test]
    fn bad_entity_is_xml_parse_error() {
        let xml = br#"<slicers><slicer name="A &bad; B" cache="c"/></slicers>"#;
        assert!(matches!(
            parse_slicers(xml),
            Err(ModernXlsxError::XmlParse(_))
        ));
        let xml = br#"<slicerCacheDefinition name="A &bad; B"/>"#;
        assert!(matches!(
            SlicerCacheData::parse(xml),
            Err(ModernXlsxError::XmlParse(_))
        ));
    }

    #[test]
    fn ampersand_name_reads_decoded_and_writes_one_amp() {
        let xml = br#"<slicers><slicer name="A &amp; B" cache="c"/></slicers>"#;
        let slicers = parse_slicers(xml).unwrap();
        assert_eq!(slicers[0].name, "A & B");
        let out = write_slicers(&slicers).unwrap();
        let s = std::str::from_utf8(&out).unwrap();
        assert!(s.contains(r#"name="A &amp; B""#), "{s}");
        assert!(!s.contains("&amp;amp;"), "{s}");
    }

    #[test]
    fn invalid_start_item_is_xml_parse_error() {
        let xml = br#"<slicers><slicer name="A" cache="c" startItem="x"/></slicers>"#;
        assert!(matches!(
            parse_slicers(xml),
            Err(ModernXlsxError::XmlParse(_))
        ));
    }

    #[test]
    fn mismatched_end_tag_is_xml_parse_error() {
        assert!(matches!(
            parse_slicers(b"<root><a></b></root>"),
            Err(ModernXlsxError::XmlParse(_))
        ));
        assert!(matches!(
            SlicerCacheData::parse(b"<root><a></b></root>"),
            Err(ModernXlsxError::XmlParse(_))
        ));
    }
}
