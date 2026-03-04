# Sheet Views (0.4.2 + 0.4.3) — TDD Implementation Plan

> **For Claude:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task.

**Goal:** Add full `<sheetView>` attribute support — grid lines, headers, zoom, RTL, view modes — completing the sheet view configuration surface.

**Architecture:** Add `SheetViewData` struct to `WorksheetXml`, parse all ECMA-376 §18.3.1.87 attributes in both parser paths, write non-default values, expose via TypeScript `Worksheet.view` getter/setter.

**Tech Stack:** Rust 1.94 (quick-xml 0.39.2, serde camelCase), TypeScript 6.0, Vitest 4.1, wasm-bindgen 0.2.114

---

## Task 1: Rust — SheetViewData Struct

**Files:** `crates/modern-xlsx-core/src/ooxml/worksheet.rs`

Add `SheetViewData` struct after `PaneSelection` (~line 286) and field to `WorksheetXml`.

Booleans with non-false defaults need serde helpers (`default_true`, `is_true`, `is_false` already exist in the file).

```rust
/// Sheet view configuration from `<sheetView>` attributes (ECMA-376 §18.3.1.87).
#[derive(Debug, Clone, Serialize, Deserialize, PartialEq)]
#[serde(rename_all = "camelCase")]
pub struct SheetViewData {
    #[serde(default = "default_true_hf", skip_serializing_if = "is_true")]
    pub show_grid_lines: bool,
    #[serde(default = "default_true_hf", skip_serializing_if = "is_true")]
    pub show_row_col_headers: bool,
    #[serde(default = "default_true_hf", skip_serializing_if = "is_true")]
    pub show_zeros: bool,
    #[serde(default, skip_serializing_if = "is_false")]
    pub right_to_left: bool,
    #[serde(default, skip_serializing_if = "is_false")]
    pub tab_selected: bool,
    #[serde(default = "default_true_hf", skip_serializing_if = "is_true")]
    pub show_ruler: bool,
    #[serde(default = "default_true_hf", skip_serializing_if = "is_true")]
    pub show_outline_symbols: bool,
    #[serde(default = "default_true_hf", skip_serializing_if = "is_true")]
    pub show_white_space: bool,
    #[serde(default = "default_true_hf", skip_serializing_if = "is_true")]
    pub default_grid_color: bool,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub zoom_scale: Option<u16>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub zoom_scale_normal: Option<u16>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub zoom_scale_page_layout_view: Option<u16>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub zoom_scale_sheet_layout_view: Option<u16>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub color_id: Option<u32>,
    /// View mode: `"normal"`, `"pageBreakPreview"`, or `"pageLayout"`.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub view: Option<String>,
}
```

Add `Default` impl (all bools to their ECMA defaults, all Options to None).

Add to `WorksheetXml`:
```rust
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub sheet_view: Option<SheetViewData>,
```

---

## Task 2: Rust — SheetView Parser

**Files:** `crates/modern-xlsx-core/src/ooxml/worksheet.rs`

### Tests first:

```rust
#[test]
fn test_parse_sheet_view_hide_gridlines_zoom() {
    let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetViews>
    <sheetView showGridLines="0" zoomScale="150" workbookViewId="0"/>
  </sheetViews>
  <sheetData/>
</worksheet>"#;
    let ws = WorksheetXml::parse(xml.as_bytes()).unwrap();
    let sv = ws.sheet_view.as_ref().unwrap();
    assert!(!sv.show_grid_lines);
    assert_eq!(sv.zoom_scale, Some(150));
}

#[test]
fn test_parse_sheet_view_rtl() {
    let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetViews>
    <sheetView rightToLeft="1" showRowColHeaders="0" workbookViewId="0"/>
  </sheetViews>
  <sheetData/>
</worksheet>"#;
    let ws = WorksheetXml::parse(xml.as_bytes()).unwrap();
    let sv = ws.sheet_view.as_ref().unwrap();
    assert!(sv.right_to_left);
    assert!(!sv.show_row_col_headers);
}

#[test]
fn test_parse_sheet_view_page_break_preview() {
    let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetViews>
    <sheetView view="pageBreakPreview" zoomScaleNormal="100" zoomScale="60" workbookViewId="0"/>
  </sheetViews>
  <sheetData/>
</worksheet>"#;
    let ws = WorksheetXml::parse(xml.as_bytes()).unwrap();
    let sv = ws.sheet_view.as_ref().unwrap();
    assert_eq!(sv.view.as_deref(), Some("pageBreakPreview"));
    assert_eq!(sv.zoom_scale, Some(60));
    assert_eq!(sv.zoom_scale_normal, Some(100));
}

#[test]
fn test_parse_sheet_view_page_layout() {
    let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetViews>
    <sheetView view="pageLayout" showRuler="0" showWhiteSpace="0" zoomScalePageLayoutView="75" workbookViewId="0"/>
  </sheetViews>
  <sheetData/>
</worksheet>"#;
    let ws = WorksheetXml::parse(xml.as_bytes()).unwrap();
    let sv = ws.sheet_view.as_ref().unwrap();
    assert_eq!(sv.view.as_deref(), Some("pageLayout"));
    assert!(!sv.show_ruler);
    assert!(!sv.show_white_space);
    assert_eq!(sv.zoom_scale_page_layout_view, Some(75));
}

#[test]
fn test_parse_sheet_view_defaults() {
    // No sheetView attributes beyond workbookViewId → sheet_view should be None or all-defaults
    let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetViews>
    <sheetView workbookViewId="0"/>
  </sheetViews>
  <sheetData/>
</worksheet>"#;
    let ws = WorksheetXml::parse(xml.as_bytes()).unwrap();
    // Either None or all-defaults is acceptable
    if let Some(sv) = &ws.sheet_view {
        assert!(sv.show_grid_lines);
        assert!(sv.show_row_col_headers);
        assert!(sv.show_zeros);
        assert!(!sv.right_to_left);
        assert!(sv.zoom_scale.is_none());
        assert!(sv.view.is_none());
    }
}

#[test]
fn test_parse_sheet_view_with_frozen_pane() {
    // sheetView attributes + frozen pane child → both parsed
    let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetViews>
    <sheetView showGridLines="0" zoomScale="120" workbookViewId="0">
      <pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>
    </sheetView>
  </sheetViews>
  <sheetData/>
</worksheet>"#;
    let ws = WorksheetXml::parse(xml.as_bytes()).unwrap();
    let sv = ws.sheet_view.as_ref().unwrap();
    assert!(!sv.show_grid_lines);
    assert_eq!(sv.zoom_scale, Some(120));
    assert!(ws.frozen_pane.is_some());
}
```

### Parser implementation:

In both full and streaming parser paths, where `(ParseState::SheetViews, b"sheetView")` transitions to `ParseState::SheetView`, parse all attributes from the `<sheetView>` element:

```rust
(ParseState::SheetViews, b"sheetView") => {
    state = ParseState::SheetView;
    let mut sv = SheetViewData::default();
    let mut has_non_default = false;
    for attr in e.attributes().flatten() {
        let ln = attr.key.local_name();
        let val = std::str::from_utf8(&attr.value).unwrap_or_default();
        match ln.as_ref() {
            b"showGridLines" => { sv.show_grid_lines = val != "0"; if val == "0" { has_non_default = true; } }
            b"showRowColHeaders" => { sv.show_row_col_headers = val != "0"; if val == "0" { has_non_default = true; } }
            b"showZeros" => { sv.show_zeros = val != "0"; if val == "0" { has_non_default = true; } }
            b"rightToLeft" => { sv.right_to_left = val == "1"; if val == "1" { has_non_default = true; } }
            b"tabSelected" => { sv.tab_selected = val == "1"; if val == "1" { has_non_default = true; } }
            b"showRuler" => { sv.show_ruler = val != "0"; if val == "0" { has_non_default = true; } }
            b"showOutlineSymbols" => { sv.show_outline_symbols = val != "0"; if val == "0" { has_non_default = true; } }
            b"showWhiteSpace" => { sv.show_white_space = val != "0"; if val == "0" { has_non_default = true; } }
            b"defaultGridColor" => { sv.default_grid_color = val != "0"; if val == "0" { has_non_default = true; } }
            b"zoomScale" => { sv.zoom_scale = val.parse().ok(); has_non_default = true; }
            b"zoomScaleNormal" => { sv.zoom_scale_normal = val.parse().ok(); has_non_default = true; }
            b"zoomScalePageLayoutView" => { sv.zoom_scale_page_layout_view = val.parse().ok(); has_non_default = true; }
            b"zoomScaleSheetLayoutView" => { sv.zoom_scale_sheet_layout_view = val.parse().ok(); has_non_default = true; }
            b"colorId" => { sv.color_id = val.parse().ok(); has_non_default = true; }
            b"view" if val != "normal" => { sv.view = Some(val.to_owned()); has_non_default = true; }
            _ => {}
        }
    }
    if has_non_default {
        sheet_view = Some(sv);
    }
}
```

Handle both `Event::Start` and `Event::Empty` for `<sheetView>` — it can be self-closing if no children.

---

## Task 3: Rust — SheetView Writer

**Files:** `crates/modern-xlsx-core/src/ooxml/worksheet.rs`

### Tests first:

```rust
#[test]
fn test_roundtrip_sheet_view_hide_gridlines_zoom() {
    let ws = WorksheetXml { sheet_view: Some(SheetViewData { show_grid_lines: false, zoom_scale: Some(150), ..SheetViewData::default() }), ..default_ws() };
    let xml = ws.to_xml().unwrap();
    let ws2 = WorksheetXml::parse(&xml).unwrap();
    let sv = ws2.sheet_view.as_ref().unwrap();
    assert!(!sv.show_grid_lines);
    assert_eq!(sv.zoom_scale, Some(150));
}

#[test]
fn test_roundtrip_sheet_view_rtl() {
    let ws = WorksheetXml { sheet_view: Some(SheetViewData { right_to_left: true, show_row_col_headers: false, ..SheetViewData::default() }), ..default_ws() };
    let xml = ws.to_xml().unwrap();
    let ws2 = WorksheetXml::parse(&xml).unwrap();
    let sv = ws2.sheet_view.as_ref().unwrap();
    assert!(sv.right_to_left);
    assert!(!sv.show_row_col_headers);
}

#[test]
fn test_roundtrip_sheet_view_page_break_preview() {
    let ws = WorksheetXml { sheet_view: Some(SheetViewData { view: Some("pageBreakPreview".to_owned()), zoom_scale: Some(60), zoom_scale_normal: Some(100), ..SheetViewData::default() }), ..default_ws() };
    let xml = ws.to_xml().unwrap();
    let ws2 = WorksheetXml::parse(&xml).unwrap();
    let sv = ws2.sheet_view.as_ref().unwrap();
    assert_eq!(sv.view.as_deref(), Some("pageBreakPreview"));
    assert_eq!(sv.zoom_scale, Some(60));
}

#[test]
fn test_roundtrip_sheet_view_with_frozen_pane() {
    let ws = WorksheetXml { sheet_view: Some(SheetViewData { show_grid_lines: false, zoom_scale: Some(120), ..SheetViewData::default() }), frozen_pane: Some(FrozenPane { rows: 1, cols: 0 }), ..default_ws() };
    let xml = ws.to_xml().unwrap();
    let ws2 = WorksheetXml::parse(&xml).unwrap();
    assert!(ws2.sheet_view.is_some());
    assert!(ws2.frozen_pane.is_some());
}

#[test]
fn test_sheet_view_defaults_minimal_xml() {
    // All-defaults → should NOT write sheetViews at all (or write minimal)
    let ws = WorksheetXml { sheet_view: Some(SheetViewData::default()), ..default_ws() };
    let xml = ws.to_xml().unwrap();
    let xml_str = std::str::from_utf8(&xml).unwrap();
    // All defaults = no non-default attributes, so sheetViews should be minimal
    assert!(!xml_str.contains("showGridLines"));
    assert!(!xml_str.contains("zoomScale"));
}
```

### Writer implementation:

Modify the `<sheetViews>` writer block. Currently it only writes if `has_pane`. Change to:

```rust
let has_sheet_views = self.sheet_view.is_some() || self.split_pane.is_some() || self.frozen_pane.is_some();
if has_sheet_views {
    writer.write_event(Event::Start(BytesStart::new("sheetViews")))?;

    let mut sv_elem = BytesStart::new("sheetView");

    // Write non-default sheetView attributes
    if let Some(ref sv) = self.sheet_view {
        if !sv.show_grid_lines { sv_elem.push_attribute(("showGridLines", "0")); }
        if !sv.show_row_col_headers { sv_elem.push_attribute(("showRowColHeaders", "0")); }
        if !sv.show_zeros { sv_elem.push_attribute(("showZeros", "0")); }
        if sv.right_to_left { sv_elem.push_attribute(("rightToLeft", "1")); }
        if sv.tab_selected { sv_elem.push_attribute(("tabSelected", "1")); }
        if !sv.show_ruler { sv_elem.push_attribute(("showRuler", "0")); }
        if !sv.show_outline_symbols { sv_elem.push_attribute(("showOutlineSymbols", "0")); }
        if !sv.show_white_space { sv_elem.push_attribute(("showWhiteSpace", "0")); }
        if !sv.default_grid_color { sv_elem.push_attribute(("defaultGridColor", "0")); }
        if let Some(z) = sv.zoom_scale { sv_elem.push_attribute(("zoomScale", ibuf.format(z))); }
        if let Some(z) = sv.zoom_scale_normal { sv_elem.push_attribute(("zoomScaleNormal", ibuf.format(z))); }
        if let Some(z) = sv.zoom_scale_page_layout_view { sv_elem.push_attribute(("zoomScalePageLayoutView", ibuf.format(z))); }
        if let Some(z) = sv.zoom_scale_sheet_layout_view { sv_elem.push_attribute(("zoomScaleSheetLayoutView", ibuf.format(z))); }
        if let Some(c) = sv.color_id { sv_elem.push_attribute(("colorId", ibuf.format(c))); }
        if let Some(ref v) = sv.view { sv_elem.push_attribute(("view", v.as_str())); }
    }

    sv_elem.push_attribute(("workbookViewId", "0"));

    // If pane children exist, write as Start element; otherwise Empty
    let has_children = self.split_pane.is_some() || self.frozen_pane.is_some() || !self.pane_selections.is_empty();
    if has_children {
        writer.write_event(Event::Start(sv_elem))?;
        // ... existing pane/selection writing ...
        writer.write_event(Event::End(BytesEnd::new("sheetView")))?;
    } else {
        writer.write_event(Event::Empty(sv_elem))?;
    }

    writer.write_event(Event::End(BytesEnd::new("sheetViews")))?;
}
```

---

## Task 4: WASM Rebuild + TypeScript Types & API

Rebuild WASM, add `SheetViewData` interface to types.ts, add `Worksheet.view` getter/setter and `Worksheet.viewMode` convenience getter/setter. Export types.

---

## Task 5: TypeScript Tests

10 TypeScript roundtrip tests: hide gridlines, zoom, RTL, page break preview, page layout, combined with frozen pane, defaults, clear view, viewMode convenience.

---

## Task 6: Full Verification + Audit

Lint, typecheck, build, clippy, full test suites. Audit for modernization.
