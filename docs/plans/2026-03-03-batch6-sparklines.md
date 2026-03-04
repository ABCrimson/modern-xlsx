# Sparklines (0.5.0–0.5.4) — TDD Implementation Plan

> **For Claude:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task.

**Goal:** Add sparkline support — small in-cell charts (line, column, win/loss) stored in OOXML extension lists (`<extLst>` with x14 namespace). Parse, preserve, write, and create sparklines programmatically.

**Architecture:** Add `SparklineGroup` struct to worksheet.rs, parse `<extLst><ext uri="{05C60535...}"><x14:sparklineGroups>` using SAX parser with new `ParseState::ExtLst` / `ParseState::SparklineGroup` states, write after `<tableParts>` before `</worksheet>`. TypeScript fluent `SparklineBuilder` API. Preserve non-sparkline extension entries as opaque XML blobs.

**Tech Stack:** Rust 1.94 (quick-xml 0.39.2, serde camelCase), TypeScript 6.0, Vitest 4.1

---

## Task 1: Rust — Sparkline data model + preserved extensions

**Files:** `crates/modern-xlsx-core/src/ooxml/worksheet.rs`

### Structs:

```rust
/// A group of sparklines sharing the same type, style, and options.
#[derive(Debug, Clone, Serialize, Deserialize, PartialEq)]
#[serde(rename_all = "camelCase")]
pub struct SparklineGroup {
    /// Sparkline type: "line", "column", or "stacked" (win/loss).
    #[serde(default = "default_sparkline_type", skip_serializing_if = "is_default_sparkline_type")]
    pub sparkline_type: String,
    /// Individual sparklines in this group.
    pub sparklines: Vec<Sparkline>,
    // --- Style colors (RGB hex, e.g. "FF376092") ---
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub color_series: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub color_negative: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub color_axis: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub color_markers: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub color_first: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub color_last: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub color_high: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub color_low: Option<String>,
    // --- Options ---
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub line_weight: Option<f64>,
    #[serde(default, skip_serializing_if = "std::ops::Not::not")]
    pub date_axis: bool,
    #[serde(default, skip_serializing_if = "std::ops::Not::not")]
    pub show_empty_cells_as: bool,  // or Option<String> for "gap"/"zero"/"span"
    #[serde(default, skip_serializing_if = "std::ops::Not::not")]
    pub markers: bool,
    #[serde(default, skip_serializing_if = "std::ops::Not::not")]
    pub high: bool,
    #[serde(default, skip_serializing_if = "std::ops::Not::not")]
    pub low: bool,
    #[serde(default, skip_serializing_if = "std::ops::Not::not")]
    pub first: bool,
    #[serde(default, skip_serializing_if = "std::ops::Not::not")]
    pub last: bool,
    #[serde(default, skip_serializing_if = "std::ops::Not::not")]
    pub negative: bool,
    #[serde(default, skip_serializing_if = "std::ops::Not::not")]
    pub display_x_axis: bool,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub manual_min: Option<f64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub manual_max: Option<f64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub display_empty_cells_as: Option<String>,
    #[serde(default, skip_serializing_if = "std::ops::Not::not")]
    pub right_to_left: bool,
}

/// A single sparkline: data range → display cell.
#[derive(Debug, Clone, Serialize, Deserialize, PartialEq)]
#[serde(rename_all = "camelCase")]
pub struct Sparkline {
    /// Data range formula (e.g. "Sheet1!A1:A10").
    pub formula: String,
    /// Cell where the sparkline renders (e.g. "B1").
    pub sqref: String,
}
```

Add to `WorksheetXml`:
```rust
pub sparkline_groups: Vec<SparklineGroup>,
/// Non-sparkline extension XML preserved as raw bytes.
#[serde(default, skip_serializing_if = "Vec::is_empty")]
pub preserved_extensions: Vec<String>,
```

Add `state: None` / `sparkline_groups: Vec::new()` / `preserved_extensions: Vec::new()` to all struct literals.

### Helper functions:
```rust
fn default_sparkline_type() -> String { "line".to_string() }
fn is_default_sparkline_type(s: &str) -> bool { s == "line" }
```

---

## Task 2: Rust — Sparkline parser (both paths)

**Files:** `crates/modern-xlsx-core/src/ooxml/worksheet.rs`

### ParseState additions:
```rust
ExtLst,
SparklineGroups,
SparklineGroup,
Sparklines,
SparklineItem,
```

### Parser logic:

In both `parse_with_sst()` and `parse_to_json()`:

1. When `Event::Start` matches `b"extLst"` at `ParseState::Root` → push `ParseState::ExtLst`
2. When `Event::Start` matches `b"ext"` at `ParseState::ExtLst`:
   - Check `uri` attribute for `{05C60535-1F16-4fd2-B633-F4F36011B0BD}` (sparklines)
   - If sparkline URI → push `ParseState::SparklineGroups`
   - Otherwise → buffer raw XML into `preserved_extensions` until matching `Event::End`
3. When `Event::Start` matches `sparklineGroups` (local name, ignore x14 prefix) → enter SparklineGroups
4. When `Event::Start` matches `sparklineGroup` → create new SparklineGroup, parse attributes (type, colors, options)
5. When `Event::Start` matches `sparklines` → enter Sparklines state
6. When `Event::Start` matches `sparkline` → enter SparklineItem state
7. Parse `<xm:f>` and `<xm:sqref>` text content within sparkline
8. Handle namespace prefixes: strip `x14:` and `xm:` prefixes, match on local name only

**Key insight:** quick-xml's `local_name()` already strips namespace prefixes, so `<x14:sparklineGroup>` matches `b"sparklineGroup"`.

### Color elements:

Parse `<x14:colorSeries rgb="FF376092"/>` etc. inside sparklineGroup:
- Match local name `colorSeries`, `colorNegative`, `colorAxis`, `colorMarkers`, `colorFirst`, `colorLast`, `colorHigh`, `colorLow`
- Read `rgb` or `theme` attribute

### Tests: 4 Rust unit tests

1. Parse sparkline group with 2 sparklines (line type)
2. Parse column sparkline with all colors set
3. Parse win/loss sparkline with markers enabled
4. Empty extLst (no sparklines) produces empty vec

---

## Task 3: Rust — Sparkline writer + roundtrip tests

**Files:** `crates/modern-xlsx-core/src/ooxml/worksheet.rs`

### Writer logic:

In `to_xml_with_sst()`, after `<tableParts>` and before `</worksheet>`:

```rust
// Write <extLst> if sparklines or preserved extensions exist
if !ws.sparkline_groups.is_empty() || !ws.preserved_extensions.is_empty() {
    writer.write_event(Event::Start(BytesStart::new("extLst")))?;

    // Write sparkline extension
    if !ws.sparkline_groups.is_empty() {
        let mut ext = BytesStart::new("ext");
        ext.push_attribute(("uri", "{05C60535-1F16-4fd2-B633-F4F36011B0BD}"));
        ext.push_attribute(("xmlns:x14", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"));
        writer.write_event(Event::Start(ext))?;

        let mut groups = BytesStart::new("x14:sparklineGroups");
        groups.push_attribute(("xmlns:xm", "http://schemas.microsoft.com/office/excel/2006/main"));
        writer.write_event(Event::Start(groups))?;

        for group in &ws.sparkline_groups {
            // Write <x14:sparklineGroup> with attributes
            // Write color elements
            // Write <x14:sparklines> with <x14:sparkline> children
        }

        writer.write_event(Event::End(BytesEnd::new("x14:sparklineGroups")))?;
        writer.write_event(Event::End(BytesEnd::new("ext")))?;
    }

    // Write preserved non-sparkline extensions
    for ext_xml in &ws.preserved_extensions {
        writer.get_mut().extend_from_slice(ext_xml.as_bytes());
    }

    writer.write_event(Event::End(BytesEnd::new("extLst")))?;
}
```

### Tests: 3 roundtrip tests

1. Line sparkline roundtrip (parse → write → parse → compare)
2. Column sparkline with colors roundtrip
3. Multiple sparkline groups in same worksheet

---

## Task 4: WASM rebuild + TypeScript types + API

### WASM rebuild

### TypeScript types (types.ts):

```typescript
export interface SparklineGroupData {
  sparklineType?: 'line' | 'column' | 'stacked';
  sparklines: SparklineData[];
  colorSeries?: string | null;
  colorNegative?: string | null;
  colorAxis?: string | null;
  colorMarkers?: string | null;
  colorFirst?: string | null;
  colorLast?: string | null;
  colorHigh?: string | null;
  colorLow?: string | null;
  lineWeight?: number | null;
  markers?: boolean;
  high?: boolean;
  low?: boolean;
  first?: boolean;
  last?: boolean;
  negative?: boolean;
  displayXAxis?: boolean;
  manualMin?: number | null;
  manualMax?: number | null;
  displayEmptyCellsAs?: 'gap' | 'zero' | 'span' | null;
  rightToLeft?: boolean;
}

export interface SparklineData {
  formula: string;
  sqref: string;
}
```

Add to `WorksheetData`:
```typescript
sparklineGroups?: SparklineGroupData[];
```

### Worksheet API (workbook.ts):

```typescript
/** Returns all sparkline groups on this sheet. */
get sparklineGroups(): SparklineGroupData[] {
  return this.data.worksheet.sparklineGroups ?? [];
}

/** Adds a sparkline group to this sheet. */
addSparklineGroup(group: SparklineGroupData): void {
  if (!this.data.worksheet.sparklineGroups) {
    this.data.worksheet.sparklineGroups = [];
  }
  this.data.worksheet.sparklineGroups.push(group);
}
```

### Export types from index.ts.

---

## Task 5: TypeScript tests + verification

**File:** `packages/modern-xlsx/__tests__/sparkline.test.ts`

### Tests (8):

1. Add line sparkline group
2. Add column sparkline group with colors
3. Add win/loss sparkline with markers
4. Line sparkline roundtrip
5. Column sparkline with colors roundtrip
6. Multiple sparkline groups roundtrip
7. Clear sparkline groups
8. Sparkline with manual min/max roundtrip

### Full verification:

```bash
cargo test -p modern-xlsx-core
cargo clippy -p modern-xlsx-core -- -D warnings
pnpm -C packages/modern-xlsx lint
pnpm -C packages/modern-xlsx typecheck
pnpm -C packages/modern-xlsx build
pnpm -C packages/modern-xlsx test
```
