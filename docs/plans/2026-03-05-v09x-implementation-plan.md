# v0.9.x Advanced Features & Production Polish — Implementation Plan

> **For Claude:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task.

**Goal:** Implement pivot tables, threaded comments, slicers, timelines, performance optimization, WASM feature-gating, API audit, CLI tool, and RC preparation — then perform a comprehensive codebase audit.

**Architecture:** New OOXML parts follow the existing SAX parsing pattern (quick-xml). Each part type gets a new Rust module with `parse()` + `to_xml()` methods, reader/writer integration via relationships and content types, and TypeScript types that cross the WASM boundary via JSON serde. Performance optimizations target hot paths (XML buffer pre-allocation, SST hash dedup, deflate tuning). WASM feature-gating uses Cargo features to split encryption and charts into optional modules.

**Tech Stack:** Rust 1.95 (Edition 2024), quick-xml 0.39.2, serde, wasm-bindgen 0.2.114, TypeScript 6.0, Vitest 4.1, Biome 2.4

---

## Pair 1: Pivot Tables (0.9.0 + 0.9.1)

### Task 1: Pivot Table Rust Types & Constants

**Files:**
- Create: `crates/modern-xlsx-core/src/ooxml/pivot_table.rs`
- Modify: `crates/modern-xlsx-core/src/ooxml/mod.rs` (add `pub mod pivot_table;`)
- Modify: `crates/modern-xlsx-core/src/ooxml/relationships.rs` (add REL constants)
- Modify: `crates/modern-xlsx-core/src/ooxml/content_types.rs` (add CT constants)

**Step 1: Add relationship and content type constants**

In `relationships.rs`, after existing `REL_DRAWING` constant (~line 26), add:
```rust
pub(crate) const REL_PIVOT_TABLE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable";
pub(crate) const REL_PIVOT_CACHE_DEF: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition";
pub(crate) const REL_PIVOT_CACHE_REC: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheRecords";
```

In `content_types.rs`, after existing `CT_DRAWING` constant (~line 32), add:
```rust
pub(crate) const CT_PIVOT_TABLE: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml";
pub(crate) const CT_PIVOT_CACHE_DEF: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml";
pub(crate) const CT_PIVOT_CACHE_REC: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml";
```

**Step 2: Create pivot_table.rs with type definitions**

```rust
use serde::{Deserialize, Serialize};

// --- Enums ---

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum PivotAxis {
    AxisRow,
    AxisCol,
    AxisPage,
    AxisValues,
}

impl PivotAxis {
    #[inline]
    fn from_xml(s: &str) -> Option<Self> {
        match s {
            "axisRow" => Some(Self::AxisRow),
            "axisCol" => Some(Self::AxisCol),
            "axisPage" => Some(Self::AxisPage),
            "axisValues" => Some(Self::AxisValues),
            _ => None,
        }
    }

    #[inline]
    fn xml_val(self) -> &'static str {
        match self {
            Self::AxisRow => "axisRow",
            Self::AxisCol => "axisCol",
            Self::AxisPage => "axisPage",
            Self::AxisValues => "axisValues",
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum SubtotalFunction {
    Sum,
    Count,
    Average,
    Max,
    Min,
    Product,
    CountNums,
    StdDev,
    StdDevP,
    Var,
    VarP,
}

impl SubtotalFunction {
    #[inline]
    fn from_xml(s: &str) -> Self {
        match s {
            "sum" => Self::Sum,
            "count" => Self::Count,
            "average" => Self::Average,
            "max" => Self::Max,
            "min" => Self::Min,
            "product" => Self::Product,
            "countNums" => Self::CountNums,
            "stdDev" => Self::StdDev,
            "stdDevp" => Self::StdDevP,
            "var" => Self::Var,
            "varp" => Self::VarP,
            _ => Self::Sum,
        }
    }

    #[inline]
    fn xml_val(self) -> &'static str {
        match self {
            Self::Sum => "sum",
            Self::Count => "count",
            Self::Average => "average",
            Self::Max => "max",
            Self::Min => "min",
            Self::Product => "product",
            Self::CountNums => "countNums",
            Self::StdDev => "stdDev",
            Self::StdDevP => "stdDevp",
            Self::Var => "var",
            Self::VarP => "varp",
        }
    }
}

// --- Structs ---

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct PivotTableData {
    pub name: String,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub data_caption: Option<String>,
    pub location: PivotLocation,
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub pivot_fields: Vec<PivotFieldData>,
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub row_fields: Vec<PivotFieldRef>,
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub col_fields: Vec<PivotFieldRef>,
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub data_fields: Vec<PivotDataFieldData>,
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub page_fields: Vec<PivotPageFieldData>,
    pub cache_id: u32,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct PivotLocation {
    #[serde(rename = "ref")]
    pub ref_range: String,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub first_header_row: Option<u32>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub first_data_row: Option<u32>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub first_data_col: Option<u32>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct PivotFieldData {
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub axis: Option<PivotAxis>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub name: Option<String>,
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub items: Vec<PivotItem>,
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub subtotals: Vec<SubtotalFunction>,
    #[serde(default, skip_serializing_if = "crate::ooxml::is_false")]
    pub compact: bool,
    #[serde(default, skip_serializing_if = "crate::ooxml::is_false")]
    pub outline: bool,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct PivotItem {
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub t: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub x: Option<u32>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct PivotFieldRef {
    pub x: i32,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct PivotDataFieldData {
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub name: Option<String>,
    pub fld: u32,
    #[serde(default)]
    pub subtotal: SubtotalFunction,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub num_fmt_id: Option<u32>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct PivotPageFieldData {
    pub fld: i32,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub item: Option<u32>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub name: Option<String>,
}
```

**Step 3: Add module declaration to mod.rs**

In `crates/modern-xlsx-core/src/ooxml/mod.rs`, after `pub mod charts;` (line 3), add:
```rust
pub mod pivot_table;
```

**Step 4: Run `cargo test -p modern-xlsx-core` to verify compilation**

Expected: All 357 existing tests pass, no compilation errors.

**Step 5: Commit**
```bash
git add crates/modern-xlsx-core/src/ooxml/pivot_table.rs crates/modern-xlsx-core/src/ooxml/mod.rs crates/modern-xlsx-core/src/ooxml/relationships.rs crates/modern-xlsx-core/src/ooxml/content_types.rs
git commit -m "feat(pivot): add PivotTableData types, relationship and content type constants"
```

---

### Task 2: Pivot Table SAX Parser

**Files:**
- Modify: `crates/modern-xlsx-core/src/ooxml/pivot_table.rs`

**Step 1: Write failing test**

Add at bottom of `pivot_table.rs`:
```rust
#[cfg(test)]
mod tests {
    use super::*;

    fn sample_pivot_table_xml() -> &'static [u8] {
        br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    name="PivotTable1" cacheId="0" dataCaption="Values">
  <location ref="A3:C11" firstHeaderRow="1" firstDataRow="2" firstDataCol="1"/>
  <pivotFields count="3">
    <pivotField axis="axisRow" compact="0" outline="0">
      <items count="2">
        <item x="0"/>
        <item t="default"/>
      </items>
    </pivotField>
    <pivotField axis="axisCol">
      <items count="2">
        <item x="0"/>
        <item t="default"/>
      </items>
    </pivotField>
    <pivotField dataField="1"/>
  </pivotFields>
  <rowFields count="1"><field x="0"/></rowFields>
  <colFields count="1"><field x="1"/></colFields>
  <dataFields count="1">
    <dataField name="Sum of Amount" fld="2" subtotal="sum"/>
  </dataFields>
</pivotTableDefinition>"#
    }

    #[test]
    fn parse_pivot_table_definition() {
        let pt = PivotTableData::parse(sample_pivot_table_xml()).unwrap();
        assert_eq!(pt.name, "PivotTable1");
        assert_eq!(pt.cache_id, 0);
        assert_eq!(pt.data_caption.as_deref(), Some("Values"));
        assert_eq!(pt.location.ref_range, "A3:C11");
        assert_eq!(pt.location.first_header_row, Some(1));
        assert_eq!(pt.location.first_data_row, Some(2));
        assert_eq!(pt.location.first_data_col, Some(1));
        assert_eq!(pt.pivot_fields.len(), 3);
        assert_eq!(pt.pivot_fields[0].axis, Some(PivotAxis::AxisRow));
        assert_eq!(pt.pivot_fields[0].items.len(), 2);
        assert_eq!(pt.pivot_fields[1].axis, Some(PivotAxis::AxisCol));
        assert_eq!(pt.row_fields.len(), 1);
        assert_eq!(pt.row_fields[0].x, 0);
        assert_eq!(pt.col_fields.len(), 1);
        assert_eq!(pt.col_fields[0].x, 1);
        assert_eq!(pt.data_fields.len(), 1);
        assert_eq!(pt.data_fields[0].name.as_deref(), Some("Sum of Amount"));
        assert_eq!(pt.data_fields[0].fld, 2);
        assert_eq!(pt.data_fields[0].subtotal, SubtotalFunction::Sum);
    }
}
```

**Step 2: Run test to verify it fails**

Run: `cargo test -p modern-xlsx-core parse_pivot_table_definition`
Expected: FAIL — `parse` method does not exist.

**Step 3: Implement SAX parser**

Add to `pivot_table.rs` (before the `#[cfg(test)]` block):
```rust
use core::hint::cold_path;
use quick_xml::events::Event;
use quick_xml::Reader;
use crate::{ModernXlsxError, Result};

impl PivotTableData {
    pub fn parse(data: &[u8]) -> Result<Self> {
        let mut reader = Reader::from_reader(data);
        reader.config_mut().trim_text(true);
        let mut buf = Vec::with_capacity(512);

        let mut name = String::new();
        let mut data_caption: Option<String> = None;
        let mut cache_id: u32 = 0;
        let mut location: Option<PivotLocation> = None;
        let mut pivot_fields: Vec<PivotFieldData> = Vec::new();
        let mut row_fields: Vec<PivotFieldRef> = Vec::new();
        let mut col_fields: Vec<PivotFieldRef> = Vec::new();
        let mut data_fields: Vec<PivotDataFieldData> = Vec::new();
        let mut page_fields: Vec<PivotPageFieldData> = Vec::new();

        let mut in_pivot_fields = false;
        let mut in_row_fields = false;
        let mut in_col_fields = false;
        let mut in_data_fields = false;
        let mut in_page_fields = false;
        let mut current_field: Option<PivotFieldData> = None;
        let mut in_items = false;

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) | Ok(Event::Empty(ref e)) => {
                    let is_empty = matches!(reader.read_event_into(&mut Vec::new()), _)
                        && e.name().as_ref() == e.name().as_ref(); // Marker — we re-check below
                    let local = e.local_name();
                    let is_start = !matches!(buf.len(), _); // Need proper check
                    // Actually, let's use the standard pattern:
                    let tag = local.as_ref();
                    match tag {
                        b"pivotTableDefinition" => {
                            for attr in e.attributes().flatten() {
                                match attr.key.local_name().as_ref() {
                                    b"name" => name = String::from_utf8_lossy(&attr.value).into(),
                                    b"cacheId" => {
                                        cache_id = std::str::from_utf8(&attr.value)
                                            .unwrap_or("0")
                                            .parse()
                                            .unwrap_or(0);
                                    }
                                    b"dataCaption" => {
                                        data_caption = Some(String::from_utf8_lossy(&attr.value).into());
                                    }
                                    _ => {}
                                }
                            }
                        }
                        b"location" => {
                            let mut loc = PivotLocation {
                                ref_range: String::new(),
                                first_header_row: None,
                                first_data_row: None,
                                first_data_col: None,
                            };
                            for attr in e.attributes().flatten() {
                                match attr.key.local_name().as_ref() {
                                    b"ref" => loc.ref_range = String::from_utf8_lossy(&attr.value).into(),
                                    b"firstHeaderRow" => {
                                        loc.first_header_row = std::str::from_utf8(&attr.value)
                                            .ok().and_then(|s| s.parse().ok());
                                    }
                                    b"firstDataRow" => {
                                        loc.first_data_row = std::str::from_utf8(&attr.value)
                                            .ok().and_then(|s| s.parse().ok());
                                    }
                                    b"firstDataCol" => {
                                        loc.first_data_col = std::str::from_utf8(&attr.value)
                                            .ok().and_then(|s| s.parse().ok());
                                    }
                                    _ => {}
                                }
                            }
                            location = Some(loc);
                        }
                        b"pivotFields" => in_pivot_fields = true,
                        b"pivotField" if in_pivot_fields => {
                            let mut field = PivotFieldData {
                                axis: None,
                                name: None,
                                items: Vec::new(),
                                subtotals: Vec::new(),
                                compact: false,
                                outline: false,
                            };
                            for attr in e.attributes().flatten() {
                                match attr.key.local_name().as_ref() {
                                    b"axis" => {
                                        field.axis = PivotAxis::from_xml(
                                            std::str::from_utf8(&attr.value).unwrap_or_default()
                                        );
                                    }
                                    b"name" => {
                                        field.name = Some(String::from_utf8_lossy(&attr.value).into());
                                    }
                                    b"compact" => {
                                        field.compact = attr.value.as_ref() != b"0";
                                    }
                                    b"outline" => {
                                        field.outline = attr.value.as_ref() != b"0";
                                    }
                                    _ => {}
                                }
                            }
                            current_field = Some(field);
                        }
                        b"items" if current_field.is_some() => in_items = true,
                        b"item" if in_items => {
                            if let Some(ref mut field) = current_field {
                                let mut item = PivotItem { t: None, x: None };
                                for attr in e.attributes().flatten() {
                                    match attr.key.local_name().as_ref() {
                                        b"t" => {
                                            item.t = Some(String::from_utf8_lossy(&attr.value).into());
                                        }
                                        b"x" => {
                                            item.x = std::str::from_utf8(&attr.value)
                                                .ok().and_then(|s| s.parse().ok());
                                        }
                                        _ => {}
                                    }
                                }
                                field.items.push(item);
                            }
                        }
                        b"rowFields" => in_row_fields = true,
                        b"colFields" => in_col_fields = true,
                        b"dataFields" => in_data_fields = true,
                        b"pageFields" => in_page_fields = true,
                        b"field" if in_row_fields || in_col_fields => {
                            let mut x: i32 = 0;
                            for attr in e.attributes().flatten() {
                                if attr.key.local_name().as_ref() == b"x" {
                                    x = std::str::from_utf8(&attr.value)
                                        .unwrap_or("0")
                                        .parse()
                                        .unwrap_or(0);
                                }
                            }
                            if in_row_fields {
                                row_fields.push(PivotFieldRef { x });
                            } else {
                                col_fields.push(PivotFieldRef { x });
                            }
                        }
                        b"dataField" if in_data_fields => {
                            let mut df = PivotDataFieldData {
                                name: None,
                                fld: 0,
                                subtotal: SubtotalFunction::Sum,
                                num_fmt_id: None,
                            };
                            for attr in e.attributes().flatten() {
                                match attr.key.local_name().as_ref() {
                                    b"name" => {
                                        df.name = Some(String::from_utf8_lossy(&attr.value).into());
                                    }
                                    b"fld" => {
                                        df.fld = std::str::from_utf8(&attr.value)
                                            .unwrap_or("0")
                                            .parse()
                                            .unwrap_or(0);
                                    }
                                    b"subtotal" => {
                                        df.subtotal = SubtotalFunction::from_xml(
                                            std::str::from_utf8(&attr.value).unwrap_or("sum")
                                        );
                                    }
                                    b"numFmtId" => {
                                        df.num_fmt_id = std::str::from_utf8(&attr.value)
                                            .ok().and_then(|s| s.parse().ok());
                                    }
                                    _ => {}
                                }
                            }
                            data_fields.push(df);
                        }
                        b"pageField" if in_page_fields => {
                            let mut pf = PivotPageFieldData {
                                fld: 0,
                                item: None,
                                name: None,
                            };
                            for attr in e.attributes().flatten() {
                                match attr.key.local_name().as_ref() {
                                    b"fld" => {
                                        pf.fld = std::str::from_utf8(&attr.value)
                                            .unwrap_or("0")
                                            .parse()
                                            .unwrap_or(0);
                                    }
                                    b"item" => {
                                        pf.item = std::str::from_utf8(&attr.value)
                                            .ok().and_then(|s| s.parse().ok());
                                    }
                                    b"name" => {
                                        pf.name = Some(String::from_utf8_lossy(&attr.value).into());
                                    }
                                    _ => {}
                                }
                            }
                            page_fields.push(pf);
                        }
                        _ => {}
                    }
                }
                Ok(Event::End(ref e)) => {
                    match e.local_name().as_ref() {
                        b"pivotFields" => in_pivot_fields = false,
                        b"pivotField" if current_field.is_some() => {
                            pivot_fields.push(current_field.take().unwrap());
                        }
                        b"items" => in_items = false,
                        b"rowFields" => in_row_fields = false,
                        b"colFields" => in_col_fields = false,
                        b"dataFields" => in_data_fields = false,
                        b"pageFields" => in_page_fields = false,
                        _ => {}
                    }
                }
                Ok(Event::Eof) => break,
                Ok(_) => {}
                Err(e) => {
                    cold_path();
                    return Err(ModernXlsxError::XmlParse(
                        format!("error parsing pivot table: {e}")
                    ));
                }
            }
            buf.clear();
        }

        Ok(PivotTableData {
            name,
            data_caption,
            location: location.unwrap_or(PivotLocation {
                ref_range: String::new(),
                first_header_row: None,
                first_data_row: None,
                first_data_col: None,
            }),
            pivot_fields,
            row_fields,
            col_fields,
            data_fields,
            page_fields,
            cache_id,
        })
    }
}
```

**NOTE:** The above is a reference implementation. The actual SAX loop should properly distinguish `Event::Start` from `Event::Empty` — use two match arms. Follow the pattern in `comments.rs:43-185` exactly.

**Step 4: Run test to verify it passes**

Run: `cargo test -p modern-xlsx-core parse_pivot_table_definition`
Expected: PASS

**Step 5: Commit**
```bash
git add crates/modern-xlsx-core/src/ooxml/pivot_table.rs
git commit -m "feat(pivot): implement SAX parser for pivot table definitions"
```

---

### Task 3: Pivot Table XML Writer

**Files:**
- Modify: `crates/modern-xlsx-core/src/ooxml/pivot_table.rs`

**Step 1: Write failing test**

```rust
#[test]
fn pivot_table_roundtrip() {
    let pt = PivotTableData::parse(sample_pivot_table_xml()).unwrap();
    let xml = pt.to_xml().unwrap();
    let pt2 = PivotTableData::parse(&xml).unwrap();
    assert_eq!(pt2.name, pt.name);
    assert_eq!(pt2.cache_id, pt.cache_id);
    assert_eq!(pt2.location.ref_range, pt.location.ref_range);
    assert_eq!(pt2.pivot_fields.len(), pt.pivot_fields.len());
    assert_eq!(pt2.row_fields.len(), pt.row_fields.len());
    assert_eq!(pt2.col_fields.len(), pt.col_fields.len());
    assert_eq!(pt2.data_fields.len(), pt.data_fields.len());
}
```

**Step 2: Run test to verify it fails**

Run: `cargo test -p modern-xlsx-core pivot_table_roundtrip`
Expected: FAIL — `to_xml` method does not exist.

**Step 3: Implement to_xml()**

Add `to_xml` method to `PivotTableData` impl block. Use `quick_xml::Writer` following the pattern in `comments::write_comments` and `tables::TableDefinition::to_xml`. Write `<pivotTableDefinition>` with namespace, attributes, nested `<location>`, `<pivotFields>`, `<rowFields>`, `<colFields>`, `<dataFields>`, `<pageFields>`.

**Step 4: Run test to verify it passes**

Run: `cargo test -p modern-xlsx-core pivot_table_roundtrip`
Expected: PASS

**Step 5: Commit**
```bash
git add crates/modern-xlsx-core/src/ooxml/pivot_table.rs
git commit -m "feat(pivot): implement XML writer for pivot table definitions"
```

---

### Task 4: Pivot Cache Types, Parser & Writer

**Files:**
- Modify: `crates/modern-xlsx-core/src/ooxml/pivot_table.rs`

**Step 1: Add cache types**

```rust
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct PivotCacheDefinitionData {
    pub source: CacheSource,
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub fields: Vec<CacheFieldData>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub record_count: Option<u32>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct CacheSource {
    #[serde(rename = "ref")]
    pub ref_range: String,
    pub sheet: String,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct CacheFieldData {
    pub name: String,
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub shared_items: Vec<CacheValue>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", tag = "type")]
pub enum CacheValue {
    Number { v: f64 },
    String { v: String },
    Boolean { v: bool },
    DateTime { v: String },
    Missing,
    Error { v: String },
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct PivotCacheRecordsData {
    pub records: Vec<Vec<CacheValue>>,
}
```

**Step 2: Write failing tests**

```rust
#[test]
fn parse_pivot_cache_definition() {
    let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    recordCount="4">
  <cacheSource type="worksheet">
    <worksheetSource ref="A1:C5" sheet="Data"/>
  </cacheSource>
  <cacheFields count="3">
    <cacheField name="Region">
      <sharedItems count="2">
        <s v="East"/>
        <s v="West"/>
      </sharedItems>
    </cacheField>
    <cacheField name="Product">
      <sharedItems count="2">
        <s v="Widget"/>
        <s v="Gadget"/>
      </sharedItems>
    </cacheField>
    <cacheField name="Amount">
      <sharedItems containsSemiMixedTypes="0" containsString="0" containsNumber="1">
        <n v="100"/>
        <n v="200"/>
      </sharedItems>
    </cacheField>
  </cacheFields>
</pivotCacheDefinition>"#;
    let cache = PivotCacheDefinitionData::parse(xml).unwrap();
    assert_eq!(cache.source.sheet, "Data");
    assert_eq!(cache.source.ref_range, "A1:C5");
    assert_eq!(cache.fields.len(), 3);
    assert_eq!(cache.fields[0].name, "Region");
    assert_eq!(cache.fields[0].shared_items.len(), 2);
    assert_eq!(cache.record_count, Some(4));
}

#[test]
fn parse_pivot_cache_records() {
    let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotCacheRecords xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="4">
  <r><x v="0"/><x v="0"/><n v="100"/></r>
  <r><x v="0"/><x v="1"/><n v="200"/></r>
  <r><x v="1"/><x v="0"/><n v="150"/></r>
  <r><x v="1"/><x v="1"/><n v="250"/></r>
</pivotCacheRecords>"#;
    let records = PivotCacheRecordsData::parse(xml).unwrap();
    assert_eq!(records.records.len(), 4);
}

#[test]
fn pivot_cache_roundtrip() {
    // Parse → to_xml → parse again → compare
}
```

**Step 3: Implement `PivotCacheDefinitionData::parse()`, `to_xml()`, `PivotCacheRecordsData::parse()`, `to_xml()`**

Follow same SAX pattern as pivot table parser.

**Step 4: Run tests**

Run: `cargo test -p modern-xlsx-core pivot_cache`
Expected: All pass

**Step 5: Commit**
```bash
git add crates/modern-xlsx-core/src/ooxml/pivot_table.rs
git commit -m "feat(pivot): implement pivot cache definition and records parser/writer"
```

---

### Task 5: Pivot Table Reader Integration

**Files:**
- Modify: `crates/modern-xlsx-core/src/ooxml/worksheet.rs:130-177` (add fields to WorksheetXml)
- Modify: `crates/modern-xlsx-core/src/lib.rs:43-75` (add fields to WorkbookData)
- Modify: `crates/modern-xlsx-core/src/reader.rs` (add resolve_pivot_tables function)

**Step 1: Add pivot table fields to WorksheetXml**

In `worksheet.rs`, before the `preserved_extensions` field (line 175), add:
```rust
/// Pivot tables attached to this worksheet.
#[serde(default, skip_serializing_if = "Vec::is_empty")]
pub pivot_tables: Vec<super::pivot_table::PivotTableData>,
```

**Step 2: Add pivot cache fields to WorkbookData**

In `lib.rs`, before `preserved_entries` (line 73), add:
```rust
/// Pivot cache definitions from xl/pivotCache/.
#[serde(default, skip_serializing_if = "Vec::is_empty")]
pub pivot_caches: Vec<crate::ooxml::pivot_table::PivotCacheDefinitionData>,
/// Pivot cache records from xl/pivotCache/.
#[serde(default, skip_serializing_if = "Vec::is_empty")]
pub pivot_cache_records: Vec<crate::ooxml::pivot_table::PivotCacheRecordsData>,
```

**Step 3: Add resolve_pivot_tables() to reader.rs**

Follow the `resolve_comments` pattern (~lines 213-240 in reader.rs). Import `REL_PIVOT_TABLE` and `REL_PIVOT_CACHE_DEF`. Add the function to resolve pivot tables per sheet via worksheet rels, and resolve pivot caches at workbook level via workbook rels.

Call `resolve_pivot_tables()` from the main `read_xlsx_with_options()` function alongside existing `resolve_comments()`, `resolve_tables()`, `resolve_charts()` calls.

**Step 4: Run all Rust tests**

Run: `cargo test -p modern-xlsx-core`
Expected: All tests pass (existing + new)

**Step 5: Commit**
```bash
git add crates/modern-xlsx-core/src/ooxml/worksheet.rs crates/modern-xlsx-core/src/lib.rs crates/modern-xlsx-core/src/reader.rs
git commit -m "feat(pivot): integrate pivot table reader into XLSX pipeline"
```

---

### Task 6: Pivot Table Writer Integration

**Files:**
- Modify: `crates/modern-xlsx-core/src/writer.rs`

**Step 1: Add pivot table writing to the writer loop**

In `writer.rs`, import the new constants and types:
```rust
use crate::ooxml::content_types::{CT_PIVOT_TABLE, CT_PIVOT_CACHE_DEF, CT_PIVOT_CACHE_REC};
use crate::ooxml::relationships::{REL_PIVOT_TABLE, REL_PIVOT_CACHE_DEF, REL_PIVOT_CACHE_REC};
```

In the per-sheet loop (after comments writing, ~line 399), add pivot table writing following the tables pattern (~lines 176-205):
- Iterate `sheet.worksheet.pivot_tables`
- Generate `xl/pivotTables/pivotTable{id}.xml` entries
- Add content type overrides
- Add relationships to ws_rels

After the per-sheet loop, write workbook-level pivot caches:
- Iterate `workbook.pivot_caches`
- Generate `xl/pivotCache/pivotCacheDefinition{id}.xml` and records entries
- Add to workbook rels

**Step 2: Write a Rust roundtrip test**

```rust
#[test]
fn pivot_table_full_roundtrip() {
    // Build a WorkbookData with a pivot table
    // Write to XLSX bytes
    // Read back
    // Verify pivot table data preserved
}
```

**Step 3: Run tests**

Run: `cargo test -p modern-xlsx-core`
Expected: All pass

**Step 4: Commit**
```bash
git add crates/modern-xlsx-core/src/writer.rs
git commit -m "feat(pivot): integrate pivot table writer into XLSX pipeline"
```

---

### Task 7: Pivot Table TypeScript Types & API

**Files:**
- Modify: `packages/modern-xlsx/src/types.ts` (add interfaces)
- Modify: `packages/modern-xlsx/src/workbook.ts` (add getter)
- Modify: `packages/modern-xlsx/src/index.ts` (add exports)
- Create: `packages/modern-xlsx/__tests__/pivot-table.test.ts`

**Step 1: Add TypeScript interfaces to types.ts**

Before `WorksheetData` interface (~line 567), add:
```typescript
export type PivotAxis = 'axisRow' | 'axisCol' | 'axisPage' | 'axisValues';
export type SubtotalFunction = 'sum' | 'count' | 'average' | 'max' | 'min' | 'product' | 'countNums' | 'stdDev' | 'stdDevP' | 'var' | 'varP';

export interface PivotLocation {
  ref: string;
  firstHeaderRow?: number;
  firstDataRow?: number;
  firstDataCol?: number;
}

export interface PivotItem {
  t?: string;
  x?: number;
}

export interface PivotFieldData {
  axis?: PivotAxis;
  name?: string;
  items: PivotItem[];
  subtotals: SubtotalFunction[];
  compact: boolean;
  outline: boolean;
}

export interface PivotDataFieldData {
  name?: string;
  fld: number;
  subtotal: SubtotalFunction;
  numFmtId?: number;
}

export interface PivotPageFieldData {
  fld: number;
  item?: number;
  name?: string;
}

export interface PivotTableData {
  name: string;
  dataCaption?: string;
  location: PivotLocation;
  pivotFields: PivotFieldData[];
  rowFields: { x: number }[];
  colFields: { x: number }[];
  dataFields: PivotDataFieldData[];
  pageFields: PivotPageFieldData[];
  cacheId: number;
}
```

Add `pivotTables` to the `WorksheetData` interface:
```typescript
pivotTables?: PivotTableData[];
```

**Step 2: Add getter to Worksheet class in workbook.ts**

After the `charts` getter (~line 1037), add:
```typescript
get pivotTables(): readonly PivotTableData[] {
  return this.data.worksheet.pivotTables ?? [];
}
```

**Step 3: Export types from index.ts**

Add to the type exports section:
```typescript
export type { PivotTableData, PivotFieldData, PivotDataFieldData, PivotPageFieldData, PivotLocation, PivotAxis, SubtotalFunction } from './types.js';
```

**Step 4: Create test file**

```typescript
import { describe, expect, it } from 'vitest';
import { Workbook } from '../src/index.js';

describe('Pivot Tables', () => {
  it('empty sheet has no pivot tables', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    expect(ws.pivotTables).toHaveLength(0);
  });
});
```

**Step 5: Run tests**

Run: `pnpm -C packages/modern-xlsx test`
Expected: All pass (existing 1210 + new)

**Step 6: Commit**
```bash
git add packages/modern-xlsx/src/types.ts packages/modern-xlsx/src/workbook.ts packages/modern-xlsx/src/index.ts packages/modern-xlsx/__tests__/pivot-table.test.ts
git commit -m "feat(pivot): add TypeScript types and read-only API for pivot tables"
```

---

### Task 8: WASM Build & Version Bump (0.9.0 + 0.9.1)

**Step 1: Build WASM**

Run: `cd crates/modern-xlsx-wasm && wasm-pack build --target web --release --out-dir ../../packages/modern-xlsx/wasm --no-opt`

**Step 2: Build TypeScript**

Run: `pnpm -C packages/modern-xlsx build`

**Step 3: Run all tests**

Run: `cargo test -p modern-xlsx-core && pnpm -C packages/modern-xlsx test`
Expected: All pass

**Step 4: Lint & typecheck**

Run: `pnpm -C packages/modern-xlsx lint && pnpm -C packages/modern-xlsx typecheck && cargo clippy -p modern-xlsx-core -- -D warnings`

**Step 5: Commit version bump**

Do NOT change version numbers (per user preference). Commit the WASM artifacts.

```bash
git add packages/modern-xlsx/wasm/
git commit -m "chore: rebuild WASM with pivot table support"
```

---

## Pair 2: Threaded Comments (0.9.2)

### Task 9: Threaded Comment Rust Types

**Files:**
- Create: `crates/modern-xlsx-core/src/ooxml/threaded_comments.rs`
- Modify: `crates/modern-xlsx-core/src/ooxml/mod.rs`
- Modify: `crates/modern-xlsx-core/src/ooxml/relationships.rs`
- Modify: `crates/modern-xlsx-core/src/ooxml/content_types.rs`

Add relationship constants:
```rust
pub(crate) const REL_THREADED_COMMENTS: &str =
    "http://schemas.microsoft.com/office/2017/10/relationships/threadedComment";
pub(crate) const REL_PERSONS: &str =
    "http://schemas.microsoft.com/office/2017/10/relationships/person";
```

Content type constants:
```rust
pub(crate) const CT_THREADED_COMMENTS: &str =
    "application/vnd.ms-excel.threadedcomments+xml";
pub(crate) const CT_PERSONS: &str =
    "application/vnd.ms-excel.person+xml";
```

Types in `threaded_comments.rs`:
```rust
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ThreadedCommentData {
    pub id: String,
    pub ref_cell: String,
    pub person_id: String,
    pub text: String,
    pub timestamp: String,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub parent_id: Option<String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct PersonData {
    pub id: String,
    pub display_name: String,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub provider_id: Option<String>,
}
```

**Commit:** `feat(comments): add threaded comment and person types`

---

### Task 10: Threaded Comment Parser & Writer

**Files:**
- Modify: `crates/modern-xlsx-core/src/ooxml/threaded_comments.rs`

Implement `ThreadedCommentData::parse_all(data: &[u8]) -> Result<Vec<ThreadedCommentData>>` and `PersonData::parse_all(data: &[u8]) -> Result<Vec<PersonData>>`.

Implement `write_threaded_comments(comments: &[ThreadedCommentData]) -> Result<Vec<u8>>` and `write_persons(persons: &[PersonData]) -> Result<Vec<u8>>`.

Write 3 tests: parse threaded comments, parse persons, roundtrip.

**Commit:** `feat(comments): implement threaded comment SAX parser and writer`

---

### Task 11: Threaded Comment Reader/Writer Integration

**Files:**
- Modify: `crates/modern-xlsx-core/src/ooxml/worksheet.rs` (add `threaded_comments` field)
- Modify: `crates/modern-xlsx-core/src/lib.rs` (add `persons` field to WorkbookData)
- Modify: `crates/modern-xlsx-core/src/reader.rs` (resolve threaded comments)
- Modify: `crates/modern-xlsx-core/src/writer.rs` (write threaded comments)

Add to WorksheetXml:
```rust
#[serde(default, skip_serializing_if = "Vec::is_empty")]
pub threaded_comments: Vec<super::threaded_comments::ThreadedCommentData>,
```

Add to WorkbookData:
```rust
#[serde(default, skip_serializing_if = "Vec::is_empty")]
pub persons: Vec<crate::ooxml::threaded_comments::PersonData>,
```

**Commit:** `feat(comments): integrate threaded comments into reader/writer pipeline`

---

### Task 12: Threaded Comment TypeScript API

**Files:**
- Modify: `packages/modern-xlsx/src/types.ts`
- Modify: `packages/modern-xlsx/src/workbook.ts`
- Modify: `packages/modern-xlsx/src/index.ts`
- Create: `packages/modern-xlsx/__tests__/threaded-comments.test.ts`

TypeScript interfaces + Worksheet methods:
```typescript
get threadedComments(): readonly ThreadedCommentData[]
addThreadedComment(cell: string, text: string, author: string): string
replyToComment(commentId: string, text: string, author: string): string
```

The `addThreadedComment` method:
1. Checks if author exists in workbook-level `persons` array
2. If not, creates a `PersonData` with `crypto.randomUUID()` for id
3. Creates `ThreadedCommentData` with `crypto.randomUUID()` for id, current ISO timestamp
4. Pushes to `this.data.worksheet.threadedComments`
5. Returns the comment id

The `replyToComment` method:
1. Same as above but sets `parentId` to the given `commentId`

Write 10 tests: create comment, reply chain 3 levels deep, roundtrip, empty sheet, multiple sheets.

**Commit:** `feat(comments): add TypeScript threaded comment API with create/reply`

---

## Pair 3: Slicers + Timelines (0.9.3 + 0.9.4)

### Task 13: Slicer Rust Types, Parser, Writer

**Files:**
- Create: `crates/modern-xlsx-core/src/ooxml/slicers.rs`
- Modify: `crates/modern-xlsx-core/src/ooxml/mod.rs`
- Modify: `crates/modern-xlsx-core/src/ooxml/relationships.rs`
- Modify: `crates/modern-xlsx-core/src/ooxml/content_types.rs`

Add constants:
```rust
// relationships.rs
pub(crate) const REL_SLICER: &str =
    "http://schemas.microsoft.com/office/2007/relationships/slicer";
pub(crate) const REL_SLICER_CACHE: &str =
    "http://schemas.microsoft.com/office/2007/relationships/slicerCache";

// content_types.rs
pub(crate) const CT_SLICER: &str = "application/vnd.ms-excel.slicer+xml";
pub(crate) const CT_SLICER_CACHE: &str = "application/vnd.ms-excel.slicerCache+xml";
```

Types, parser, writer in `slicers.rs`. Write 5 tests.

**Commit:** `feat(slicers): add slicer types, SAX parser, and XML writer`

---

### Task 14: Slicer Reader/Writer Integration + TypeScript

**Files:**
- Modify: `crates/modern-xlsx-core/src/ooxml/worksheet.rs`
- Modify: `crates/modern-xlsx-core/src/lib.rs`
- Modify: `crates/modern-xlsx-core/src/reader.rs`
- Modify: `crates/modern-xlsx-core/src/writer.rs`
- Modify: `packages/modern-xlsx/src/types.ts`
- Modify: `packages/modern-xlsx/src/workbook.ts`
- Modify: `packages/modern-xlsx/src/index.ts`
- Create: `packages/modern-xlsx/__tests__/slicers.test.ts`

Add `slicers` Vec to WorksheetXml, `slicer_caches` Vec to WorkbookData. Reader/writer integration follows the same pattern as pivot tables. TypeScript: `get slicers(): readonly SlicerData[]`.

**Commit:** `feat(slicers): integrate slicers into full XLSX pipeline with TypeScript API`

---

### Task 15: Timeline Rust Types, Parser, Writer

**Files:**
- Create: `crates/modern-xlsx-core/src/ooxml/timelines.rs`
- Modify: `crates/modern-xlsx-core/src/ooxml/mod.rs`
- Modify: `crates/modern-xlsx-core/src/ooxml/relationships.rs`
- Modify: `crates/modern-xlsx-core/src/ooxml/content_types.rs`

Add constants:
```rust
// relationships.rs
pub(crate) const REL_TIMELINE: &str =
    "http://schemas.microsoft.com/office/2011/relationships/timeline";
pub(crate) const REL_TIMELINE_CACHE: &str =
    "http://schemas.microsoft.com/office/2011/relationships/timelineCache";

// content_types.rs
pub(crate) const CT_TIMELINE: &str = "application/vnd.ms-excel.timeline+xml";
pub(crate) const CT_TIMELINE_CACHE: &str = "application/vnd.ms-excel.timelineCache+xml";
```

Types, parser, writer in `timelines.rs`. Write 5 tests.

**Commit:** `feat(timelines): add timeline types, SAX parser, and XML writer`

---

### Task 16: Timeline Reader/Writer Integration + TypeScript

Same pattern as Task 14 but for timelines.

Add `timelines` Vec to WorksheetXml, `timeline_caches` Vec to WorkbookData.
TypeScript: `get timelines(): readonly TimelineData[]`.

**Commit:** `feat(timelines): integrate timelines into full XLSX pipeline with TypeScript API`

---

## Pair 4: Performance + WASM Size (0.9.5 + 0.9.6)

### Task 17: XML Writer Buffer Pre-allocation

**Files:**
- Modify: `crates/modern-xlsx-core/src/writer.rs`
- Modify: `crates/modern-xlsx-core/src/ooxml/worksheet.rs` (streaming writer)

Identify all `String::new()` used for XML output buffers. Replace with `String::with_capacity()` using estimated sizes:
- Worksheet XML: ~100 bytes per row + ~50 bytes per cell
- SST XML: ~30 bytes per string entry
- Styles XML: ~200 bytes per style

**Commit:** `perf: pre-allocate XML writer buffers based on data size estimates`

---

### Task 18: SST Hash-Based Dedup

**Files:**
- Modify: `crates/modern-xlsx-core/src/ooxml/shared_strings.rs`

In `SharedStringTableBuilder`, replace any linear-scan dedup with `HashMap<String, u32>` for O(1) string-to-index lookup during write. The builder already exists — optimize its internal data structure.

**Commit:** `perf: use HashMap for O(1) SST string dedup during write`

---

### Task 19: Streaming JSON Buffer Reuse

**Files:**
- Modify: `crates/modern-xlsx-core/src/streaming.rs`

Audit `parse_to_json` for unnecessary allocations. Ensure all temporary `String` buffers use `std::mem::take()` for reuse. Ensure `itoa` is used for all integer formatting. Pre-allocate the output JSON string.

**Commit:** `perf: optimize streaming JSON with buffer reuse and pre-allocation`

---

### Task 20: Benchmark Tests

**Files:**
- Create: `crates/modern-xlsx-core/tests/benchmarks.rs` (integration test)

Add benchmark-style tests that measure read/write timing for 10K, 100K rows. These are `#[test]` functions that print timing info (not criterion benchmarks — keep it simple).

**Commit:** `test: add performance benchmark tests for large workbooks`

---

### Task 21: Cargo Feature Gates

**Files:**
- Modify: `crates/modern-xlsx-core/Cargo.toml`
- Modify: `crates/modern-xlsx-wasm/Cargo.toml`
- Modify: `crates/modern-xlsx-core/src/lib.rs`
- Modify: `crates/modern-xlsx-core/src/ole2/mod.rs`
- Modify: `crates/modern-xlsx-core/src/ole2/crypto.rs`
- Modify: `crates/modern-xlsx-core/src/ole2/detect.rs`
- Modify: `crates/modern-xlsx-core/src/ole2/writer.rs`
- Modify: `crates/modern-xlsx-core/src/ole2/encryption_info.rs`
- Modify: `crates/modern-xlsx-core/src/ooxml/charts.rs`
- Modify: `crates/modern-xlsx-core/src/reader.rs`
- Modify: `crates/modern-xlsx-core/src/writer.rs`

In `crates/modern-xlsx-core/Cargo.toml`, add:
```toml
[features]
default = ["encryption", "charts"]
encryption = ["dep:sha2", "dep:sha1", "dep:aes", "dep:cbc", "dep:hmac", "dep:digest", "dep:zeroize", "dep:constant_time_eq", "dep:getrandom"]
charts = []
```

Mark all crypto deps as optional. Wrap all OLE2/encryption code with `#[cfg(feature = "encryption")]`. Wrap chart parsing/writing with `#[cfg(feature = "charts")]`.

Ensure `crates/modern-xlsx-wasm/Cargo.toml` depends on `modern-xlsx-core` with `default-features = true`.

Verify: `cargo test -p modern-xlsx-core --no-default-features` (core tests pass without encryption/charts)
Verify: `cargo test -p modern-xlsx-core` (all tests pass with default features)

**Commit:** `feat: add Cargo feature gates for encryption and charts modules`

---

### Task 22: WASM Size Optimization

**Files:**
- Modify: `Cargo.toml` (workspace profile)
- Modify build scripts or CI

Add wasm-opt to the build pipeline. After `wasm-pack build`, run:
```bash
wasm-opt -Oz --enable-bulk-memory --enable-nontrapping-float-to-int -o optimized.wasm input.wasm
```

Add CI size check script that measures gzipped WASM size and fails if exceeding thresholds.

**Commit:** `perf: add wasm-opt optimization pass and CI size check`

---

## Pair 5: API Audit + CLI + RC (0.9.7 + 0.9.8 + 0.9.9)

### Task 23: API Consistency Audit

**Files:**
- Modify: `packages/modern-xlsx/src/workbook.ts`
- Modify: `packages/modern-xlsx/src/types.ts`
- Modify: `packages/modern-xlsx/src/index.ts`

Audit checklist:
1. All getters return `readonly` arrays (not mutable arrays)
2. All optional returns use `null` (not `undefined`) for consistency
3. Methods with 3+ positional params get options object overloads
4. Standardize error class: `ModernXlsxError` with `code` property
5. Add `@deprecated` JSDoc to any APIs that will change in 1.0

**Commit:** `refactor: API consistency audit — readonly arrays, null returns, error codes`

---

### Task 24: CLI Tool

**Files:**
- Create: `packages/modern-xlsx/src/cli.ts`
- Modify: `packages/modern-xlsx/package.json` (add `bin` field)
- Create: `packages/modern-xlsx/__tests__/cli.test.ts`

Implement:
- `modern-xlsx info <file>` — reads file, prints sheet names, dimensions, row counts
- `modern-xlsx convert <file> <output>` — xlsx to JSON
- `modern-xlsx convert <file> <output> --sheet 0 --format csv` — single sheet to CSV

Uses `node:fs`, `node:process`, imports from the library itself. No external deps.

Add to package.json:
```json
"bin": {
  "modern-xlsx": "./dist/cli.mjs"
}
```

Write 5 tests using temp files.

**Commit:** `feat: add CLI tool for info and convert commands`

---

### Task 25: CDN Bundle Verification

**Files:**
- Modify: `packages/modern-xlsx/package.json` (verify fields)

Verify:
1. IIFE build produces working `modern-xlsx.min.js`
2. WASM is properly inlined or fetchable from CDN
3. `window.ModernXlsx` namespace available in browser

Add example to README.

**Commit:** `docs: verify and document CDN bundle usage`

---

### Task 26: CI Runtime Matrix

**Files:**
- Modify: `.github/workflows/ci.yml`

Add jobs for:
- Node.js 20, 22, 24
- Bun latest
- Deno latest
- Playwright: Chrome, Firefox, WebKit (via existing `@vitest/browser`)

**Commit:** `ci: add multi-runtime test matrix (Node 20/22/24, Bun, Deno, Playwright browsers)`

---

### Task 27: Comprehensive Test XLSX Generation

**Files:**
- Create: `packages/modern-xlsx/__tests__/generate-comprehensive.test.ts`

Create a test that generates a single XLSX file with every feature:
- Multiple sheets, styles, formulas, data validation, conditional formatting
- Charts, tables, merge cells, hyperlinks, comments, threaded comments
- Pivot tables, slicers, timelines
- Frozen panes, split panes, page setup, headers/footers

Save as `test-output/comprehensive-test.xlsx` for manual verification.

**Commit:** `test: generate comprehensive XLSX for manual verification in Excel/Sheets/LibreOffice`

---

### Task 28: WASM Rebuild, Full Test Suite, Version Prep

**Step 1:** Rebuild WASM: `cd crates/modern-xlsx-wasm && wasm-pack build --target web --release --out-dir ../../packages/modern-xlsx/wasm --no-opt`
**Step 2:** Build TypeScript: `pnpm -C packages/modern-xlsx build`
**Step 3:** Run all Rust tests: `cargo test -p modern-xlsx-core`
**Step 4:** Run all TypeScript tests: `pnpm -C packages/modern-xlsx test`
**Step 5:** Lint: `pnpm -C packages/modern-xlsx lint && cargo clippy -p modern-xlsx-core -- -D warnings`
**Step 6:** Typecheck: `pnpm -C packages/modern-xlsx typecheck`

All must pass before proceeding to the audit.

**Commit:** `chore: rebuild WASM and verify full test suite for 0.9.x features`

---

## Final: Comprehensive Codebase Audit

### Task 29: Rust Modernization Audit

**Scope:** Every `.rs` file in `crates/modern-xlsx-core/src/`

Checklist for every file:
- [ ] `cold_path()` on all error/panic branches
- [ ] `.find()` instead of single-attribute `for` loops
- [ ] `#[inline]` on hot-path helpers
- [ ] `std::mem::take()` for buffer reuse instead of `.clone()` + `.clear()`
- [ ] `String::with_capacity()` for known-size strings
- [ ] `Vec::with_capacity()` for known-size vectors
- [ ] No unnecessary allocations in loops
- [ ] Iterator combinators where cleaner than manual loops
- [ ] Edition 2024 idioms (let chains, etc.)
- [ ] No panics — all `unwrap()` replaced with `unwrap_or_default()` or `?`
- [ ] Proper `#[serde(rename_all = "camelCase")]` on all public types
- [ ] Skip serialization of defaults (`skip_serializing_if`)

**Commit:** `refactor: Rust modernization audit — Edition 2024 idioms, zero-copy, cold_path`

---

### Task 30: TypeScript Modernization Audit

**Scope:** Every `.ts` file in `packages/modern-xlsx/src/`

Checklist for every file:
- [ ] `satisfies` used where type narrowing helps
- [ ] No `as` casts that could be `satisfies` or type guards
- [ ] `using` keyword for disposable resources (if applicable)
- [ ] `readonly` on all array return types
- [ ] Consistent `null` (not `undefined`) for optional returns
- [ ] No `any` — use `unknown` + type guards
- [ ] Single quotes, trailing commas, semicolons (Biome)
- [ ] No unnecessary `!` non-null assertions
- [ ] Safe array access patterns (no unchecked indexing)

**Commit:** `refactor: TypeScript 6.0 modernization audit — satisfies, readonly, null consistency`

---

### Task 31: Performance Audit

**Scope:** Hot paths in reader, writer, streaming JSON, WASM bridge

Checklist:
- [ ] All XML output buffers pre-allocated
- [ ] SST dedup is O(1) via HashMap
- [ ] No redundant string copies in cell value serialization
- [ ] `itoa` used for all integer-to-string conversions
- [ ] Streaming JSON avoids intermediate allocations
- [ ] ZIP compression level tuned (verify zopfli vs deflate tradeoff)
- [ ] WASM binary size within targets

**Commit:** `perf: comprehensive performance audit — buffer allocation, dedup, streaming`

---

### Task 32: Documentation & Aesthetics Audit

**Scope:** README.md, CLAUDE.md, wiki pages, GitHub presence

Checklist:
- [ ] README: accurate feature matrix, benchmark table, badges, quick-start
- [ ] README: CDN usage example, CLI usage example
- [ ] CLAUDE.md: updated test counts, file lists, architecture
- [ ] Wiki: all 15 pages up to date with 0.9.x features
- [ ] Wiki: pivot tables, threaded comments, slicers, timelines documented
- [ ] CHANGELOG: entries for all 0.9.x releases
- [ ] GitHub: release notes drafted for 0.9.x

**Commit:** `docs: comprehensive documentation audit — README, wiki, changelog, feature matrix`

---

## Summary

| Task | Description | Pair |
|------|-------------|------|
| 1 | Pivot table Rust types & constants | 1 |
| 2 | Pivot table SAX parser | 1 |
| 3 | Pivot table XML writer | 1 |
| 4 | Pivot cache types, parser, writer | 1 |
| 5 | Pivot table reader integration | 1 |
| 6 | Pivot table writer integration | 1 |
| 7 | Pivot table TypeScript types & API | 1 |
| 8 | WASM build & verify (0.9.0+0.9.1) | 1 |
| 9 | Threaded comment Rust types | 2 |
| 10 | Threaded comment parser & writer | 2 |
| 11 | Threaded comment reader/writer integration | 2 |
| 12 | Threaded comment TypeScript API | 2 |
| 13 | Slicer Rust types, parser, writer | 3 |
| 14 | Slicer integration + TypeScript | 3 |
| 15 | Timeline Rust types, parser, writer | 3 |
| 16 | Timeline integration + TypeScript | 3 |
| 17 | XML writer buffer pre-allocation | 4 |
| 18 | SST hash-based dedup | 4 |
| 19 | Streaming JSON buffer reuse | 4 |
| 20 | Benchmark tests | 4 |
| 21 | Cargo feature gates | 4 |
| 22 | WASM size optimization | 4 |
| 23 | API consistency audit | 5 |
| 24 | CLI tool | 5 |
| 25 | CDN bundle verification | 5 |
| 26 | CI runtime matrix | 5 |
| 27 | Comprehensive test XLSX generation | 5 |
| 28 | WASM rebuild & full test suite | 5 |
| 29 | Rust modernization audit | Audit |
| 30 | TypeScript modernization audit | Audit |
| 31 | Performance audit | Audit |
| 32 | Documentation & aesthetics audit | Audit |
