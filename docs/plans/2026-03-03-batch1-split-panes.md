# Split Panes (0.4.0 + 0.4.1) — TDD Implementation Plan

> **For Claude:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task.

**Goal:** Add full split pane support (horizontal, vertical, four-way) with per-pane selections, complementing the existing frozen pane feature.

**Architecture:** Extend the existing `<pane>` SAX parser/writer in `worksheet.rs` to distinguish `state="split"` from `state="frozen"`, add `SplitPane` + `PaneSelection` structs alongside `FrozenPane`, wire through WASM JSON bridge to TypeScript API. Frozen and split are mutually exclusive on any given sheet.

**Tech Stack:** Rust 1.94 (quick-xml 0.39.2, serde camelCase), TypeScript 6.0, Vitest 4.1, wasm-bindgen 0.2.114

---

## Task 1: Rust — Split Pane & PaneSelection Data Model

**Files:**
- Modify: `crates/modern-xlsx-core/src/ooxml/worksheet.rs` (lines ~239-247, ~133-168)

**Step 1:** Add `SplitPane` struct after `FrozenPane` (line ~248):

```rust
/// Split pane configuration — divides the sheet view into 2 or 4 scrollable regions.
/// `xSplit` / `ySplit` values are in **twips** (1/20th of a point) for split mode.
#[derive(Debug, Clone, Serialize, Deserialize, PartialEq)]
#[serde(rename_all = "camelCase")]
pub struct SplitPane {
    /// Horizontal split position in twips (ySplit). `None` or `0.0` means no horizontal split.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub horizontal: Option<f64>,
    /// Vertical split position in twips (xSplit). `None` or `0.0` means no vertical split.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub vertical: Option<f64>,
    /// Cell reference for the top-left cell in the bottom-right pane.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub top_left_cell: Option<String>,
    /// The active pane: `"topLeft"`, `"topRight"`, `"bottomLeft"`, `"bottomRight"`.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub active_pane: Option<String>,
}
```

**Step 2:** Add `PaneSelection` struct after `SplitPane`:

```rust
/// Per-pane selection state within a `<sheetView>`.
/// Each visible pane can have its own active cell and selection range.
#[derive(Debug, Clone, Serialize, Deserialize, PartialEq)]
#[serde(rename_all = "camelCase")]
pub struct PaneSelection {
    /// Which pane this selection belongs to: `"topLeft"`, `"topRight"`, `"bottomLeft"`, `"bottomRight"`.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub pane: Option<String>,
    /// The active (focused) cell reference, e.g. `"A1"`.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub active_cell: Option<String>,
    /// The selected range, e.g. `"A1:C5"`.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub sqref: Option<String>,
}
```

**Step 3:** Add fields to `WorksheetXml` struct (after `frozen_pane` at line ~145):

```rust
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub split_pane: Option<SplitPane>,
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub pane_selections: Vec<PaneSelection>,
```

**Step 4:** Run `cargo test -p modern-xlsx-core` — verify no regressions (all 217 tests pass). The new fields have `serde(default)` so existing JSON roundtrips are unaffected.

**Step 5:** Commit:
```bash
git add crates/modern-xlsx-core/src/ooxml/worksheet.rs
git commit -m "feat: add SplitPane and PaneSelection data model structs"
```

---

## Task 2: Rust — Split Pane Parser (Full + Streaming)

**Files:**
- Modify: `crates/modern-xlsx-core/src/ooxml/worksheet.rs`

**Step 1:** Write failing tests in the `#[cfg(test)]` module. Add after the existing `test_parse_frozen_pane` test:

```rust
#[test]
fn test_parse_horizontal_split_pane() {
    let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetViews>
    <sheetView workbookViewId="0">
      <pane ySplit="2400" topLeftCell="A5" activePane="bottomLeft" state="split"/>
      <selection pane="bottomLeft" activeCell="A5" sqref="A5"/>
    </sheetView>
  </sheetViews>
  <sheetData/>
</worksheet>"#;
    let ws = WorksheetXml::parse(xml.as_bytes()).unwrap();
    assert!(ws.frozen_pane.is_none(), "split pane must not set frozen_pane");
    let sp = ws.split_pane.as_ref().expect("split_pane should be Some");
    assert!((sp.horizontal.unwrap() - 2400.0).abs() < f64::EPSILON);
    assert!(sp.vertical.is_none());
    assert_eq!(sp.top_left_cell.as_deref(), Some("A5"));
    assert_eq!(sp.active_pane.as_deref(), Some("bottomLeft"));
    assert_eq!(ws.pane_selections.len(), 1);
    assert_eq!(ws.pane_selections[0].pane.as_deref(), Some("bottomLeft"));
    assert_eq!(ws.pane_selections[0].active_cell.as_deref(), Some("A5"));
    assert_eq!(ws.pane_selections[0].sqref.as_deref(), Some("A5"));
}

#[test]
fn test_parse_vertical_split_pane() {
    let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetViews>
    <sheetView workbookViewId="0">
      <pane xSplit="3000" topLeftCell="D1" activePane="topRight" state="split"/>
      <selection pane="topRight" activeCell="D1" sqref="D1"/>
    </sheetView>
  </sheetViews>
  <sheetData/>
</worksheet>"#;
    let ws = WorksheetXml::parse(xml.as_bytes()).unwrap();
    assert!(ws.frozen_pane.is_none());
    let sp = ws.split_pane.as_ref().unwrap();
    assert!(sp.horizontal.is_none());
    assert!((sp.vertical.unwrap() - 3000.0).abs() < f64::EPSILON);
    assert_eq!(sp.top_left_cell.as_deref(), Some("D1"));
    assert_eq!(sp.active_pane.as_deref(), Some("topRight"));
}

#[test]
fn test_parse_four_way_split_pane() {
    let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetViews>
    <sheetView workbookViewId="0">
      <pane xSplit="3000" ySplit="2400" topLeftCell="D5" activePane="bottomRight" state="split"/>
      <selection pane="topRight" activeCell="D1" sqref="D1"/>
      <selection pane="bottomLeft" activeCell="A5" sqref="A5"/>
      <selection pane="bottomRight" activeCell="D5" sqref="D5"/>
    </sheetView>
  </sheetViews>
  <sheetData/>
</worksheet>"#;
    let ws = WorksheetXml::parse(xml.as_bytes()).unwrap();
    assert!(ws.frozen_pane.is_none());
    let sp = ws.split_pane.as_ref().unwrap();
    assert!((sp.vertical.unwrap() - 3000.0).abs() < f64::EPSILON);
    assert!((sp.horizontal.unwrap() - 2400.0).abs() < f64::EPSILON);
    assert_eq!(sp.top_left_cell.as_deref(), Some("D5"));
    assert_eq!(sp.active_pane.as_deref(), Some("bottomRight"));
    assert_eq!(ws.pane_selections.len(), 3);
}

#[test]
fn test_frozen_and_split_mutually_exclusive() {
    // Frozen pane XML → split_pane is None
    let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetViews>
    <sheetView workbookViewId="0">
      <pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>
    </sheetView>
  </sheetViews>
  <sheetData/>
</worksheet>"#;
    let ws = WorksheetXml::parse(xml.as_bytes()).unwrap();
    assert!(ws.frozen_pane.is_some());
    assert!(ws.split_pane.is_none());
}

#[test]
fn test_parse_split_no_state_attr() {
    // No state attribute → treated as split (default per ECMA-376)
    let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetViews>
    <sheetView workbookViewId="0">
      <pane xSplit="1500" ySplit="900" topLeftCell="B2"/>
    </sheetView>
  </sheetViews>
  <sheetData/>
</worksheet>"#;
    let ws = WorksheetXml::parse(xml.as_bytes()).unwrap();
    assert!(ws.frozen_pane.is_none());
    let sp = ws.split_pane.as_ref().unwrap();
    assert!((sp.vertical.unwrap() - 1500.0).abs() < f64::EPSILON);
    assert!((sp.horizontal.unwrap() - 900.0).abs() < f64::EPSILON);
}
```

**Step 2:** Run `cargo test -p modern-xlsx-core` → expect 5 failures (struct fields not parsed yet).

**Step 3:** Modify the full parser's `<pane>` parsing block (~line 894). Replace the current parsing logic:

```rust
(ParseState::SheetView, b"pane") => {
    let mut y_split_raw = String::new();
    let mut x_split_raw = String::new();
    let mut state_val = String::new();
    let mut top_left_cell = None::<String>;
    let mut active_pane_val = None::<String>;

    for attr in e.attributes().flatten() {
        let ln = attr.key.local_name();
        match ln.as_ref() {
            b"ySplit" => {
                y_split_raw = std::str::from_utf8(&attr.value).unwrap_or_default().to_owned();
            }
            b"xSplit" => {
                x_split_raw = std::str::from_utf8(&attr.value).unwrap_or_default().to_owned();
            }
            b"state" => {
                state_val = std::str::from_utf8(&attr.value).unwrap_or_default().to_owned();
            }
            b"topLeftCell" => {
                top_left_cell = Some(std::str::from_utf8(&attr.value).unwrap_or_default().to_owned());
            }
            b"activePane" => {
                active_pane_val = Some(std::str::from_utf8(&attr.value).unwrap_or_default().to_owned());
            }
            _ => {}
        }
    }

    if state_val == "frozen" {
        // Frozen pane: xSplit/ySplit are integer cell counts.
        let y: u32 = y_split_raw.parse().unwrap_or(0);
        let x: u32 = x_split_raw.parse().unwrap_or(0);
        if y > 0 || x > 0 {
            frozen_pane = Some(FrozenPane { rows: y, cols: x });
        }
    } else {
        // Split pane (state="split" or missing) — xSplit/ySplit are f64 twips.
        let y: f64 = y_split_raw.parse().unwrap_or(0.0);
        let x: f64 = x_split_raw.parse().unwrap_or(0.0);
        if y > 0.0 || x > 0.0 {
            split_pane = Some(SplitPane {
                horizontal: if y > 0.0 { Some(y) } else { None },
                vertical: if x > 0.0 { Some(x) } else { None },
                top_left_cell,
                active_pane: active_pane_val,
            });
        }
    }
}
```

**Step 4:** Add `<selection>` parsing after the `<pane>` match arm:

```rust
(ParseState::SheetView, b"selection") => {
    let mut sel_pane = None::<String>;
    let mut sel_active = None::<String>;
    let mut sel_sqref = None::<String>;
    for attr in e.attributes().flatten() {
        let ln = attr.key.local_name();
        match ln.as_ref() {
            b"pane" => {
                sel_pane = Some(std::str::from_utf8(&attr.value).unwrap_or_default().to_owned());
            }
            b"activeCell" => {
                sel_active = Some(std::str::from_utf8(&attr.value).unwrap_or_default().to_owned());
            }
            b"sqref" => {
                sel_sqref = Some(std::str::from_utf8(&attr.value).unwrap_or_default().to_owned());
            }
            _ => {}
        }
    }
    pane_selections.push(PaneSelection {
        pane: sel_pane,
        active_cell: sel_active,
        sqref: sel_sqref,
    });
}
```

**Step 5:** Declare `split_pane` and `pane_selections` local variables alongside `frozen_pane` at the top of the parsing function. Assign them to the `WorksheetXml` return struct.

**Step 6:** Apply identical changes to the streaming parser path (~line 1950).

**Step 7:** Run `cargo test -p modern-xlsx-core` → all new + existing tests pass (222 total).

**Step 8:** Commit:
```bash
git add crates/modern-xlsx-core/src/ooxml/worksheet.rs
git commit -m "feat: parse split pane and per-pane selection elements"
```

---

## Task 3: Rust — Split Pane Writer

**Files:**
- Modify: `crates/modern-xlsx-core/src/ooxml/worksheet.rs` (writer section ~line 2714)

**Step 1:** Write failing roundtrip test:

```rust
#[test]
fn test_roundtrip_horizontal_split() {
    let mut ws = WorksheetXml::default_for_test();
    ws.split_pane = Some(SplitPane {
        horizontal: Some(2400.0),
        vertical: None,
        top_left_cell: Some("A5".to_owned()),
        active_pane: Some("bottomLeft".to_owned()),
    });
    ws.pane_selections = vec![PaneSelection {
        pane: Some("bottomLeft".to_owned()),
        active_cell: Some("A5".to_owned()),
        sqref: Some("A5".to_owned()),
    }];
    let xml = ws.to_xml().unwrap();
    let ws2 = WorksheetXml::parse(&xml).unwrap();
    assert!(ws2.frozen_pane.is_none());
    let sp = ws2.split_pane.as_ref().unwrap();
    assert!((sp.horizontal.unwrap() - 2400.0).abs() < f64::EPSILON);
    assert!(sp.vertical.is_none());
    assert_eq!(sp.top_left_cell.as_deref(), Some("A5"));
    assert_eq!(ws2.pane_selections.len(), 1);
}

#[test]
fn test_roundtrip_four_way_split() {
    let mut ws = WorksheetXml::default_for_test();
    ws.split_pane = Some(SplitPane {
        horizontal: Some(2400.0),
        vertical: Some(3000.0),
        top_left_cell: Some("D5".to_owned()),
        active_pane: Some("bottomRight".to_owned()),
    });
    ws.pane_selections = vec![
        PaneSelection { pane: Some("topRight".to_owned()), active_cell: Some("D1".to_owned()), sqref: Some("D1".to_owned()) },
        PaneSelection { pane: Some("bottomLeft".to_owned()), active_cell: Some("A5".to_owned()), sqref: Some("A5".to_owned()) },
        PaneSelection { pane: Some("bottomRight".to_owned()), active_cell: Some("D5".to_owned()), sqref: Some("D5".to_owned()) },
    ];
    let xml = ws.to_xml().unwrap();
    let ws2 = WorksheetXml::parse(&xml).unwrap();
    let sp = ws2.split_pane.as_ref().unwrap();
    assert!((sp.vertical.unwrap() - 3000.0).abs() < f64::EPSILON);
    assert!((sp.horizontal.unwrap() - 2400.0).abs() < f64::EPSILON);
    assert_eq!(ws2.pane_selections.len(), 3);
}

#[test]
fn test_split_pane_overrides_frozen_on_write() {
    // If both somehow set, split takes priority and frozen is ignored.
    let mut ws = WorksheetXml::default_for_test();
    ws.frozen_pane = Some(FrozenPane { rows: 1, cols: 0 });
    ws.split_pane = Some(SplitPane {
        horizontal: Some(2400.0),
        vertical: None,
        top_left_cell: Some("A5".to_owned()),
        active_pane: Some("bottomLeft".to_owned()),
    });
    let xml = ws.to_xml().unwrap();
    let xml_str = std::str::from_utf8(&xml).unwrap();
    assert!(xml_str.contains(r#"state="split""#));
    assert!(!xml_str.contains(r#"state="frozen""#));
}
```

**Step 2:** Run `cargo test -p modern-xlsx-core` → expect failures (writer doesn't handle split_pane yet).

**Step 3:** Rewrite the `<sheetViews>` writer block (~line 2714) to handle both frozen and split:

```rust
// <sheetViews> — if frozen_pane or split_pane is present.
// Split pane takes priority over frozen pane (mutually exclusive).
let has_pane = self.split_pane.is_some() || self.frozen_pane.is_some();
if has_pane {
    writer
        .write_event(Event::Start(BytesStart::new("sheetViews")))
        .map_err(map_err)?;

    let mut sv = BytesStart::new("sheetView");
    sv.push_attribute(("workbookViewId", "0"));
    writer.write_event(Event::Start(sv)).map_err(map_err)?;

    if let Some(ref sp) = self.split_pane {
        // --- Split pane ---
        let mut pane_elem = BytesStart::new("pane");
        if let Some(x) = sp.vertical {
            let s = format_f64(x);
            pane_elem.push_attribute(("xSplit", s.as_str()));
        }
        if let Some(y) = sp.horizontal {
            let s = format_f64(y);
            pane_elem.push_attribute(("ySplit", s.as_str()));
        }
        if let Some(ref tlc) = sp.top_left_cell {
            pane_elem.push_attribute(("topLeftCell", tlc.as_str()));
        }
        if let Some(ref ap) = sp.active_pane {
            pane_elem.push_attribute(("activePane", ap.as_str()));
        }
        pane_elem.push_attribute(("state", "split"));
        writer.write_event(Event::Empty(pane_elem)).map_err(map_err)?;
    } else if let Some(ref pane) = self.frozen_pane {
        // --- Frozen pane (existing logic) ---
        let mut pane_elem = BytesStart::new("pane");
        if pane.cols > 0 {
            pane_elem.push_attribute(("xSplit", ibuf.format(pane.cols)));
        }
        if pane.rows > 0 {
            pane_elem.push_attribute(("ySplit", ibuf.format(pane.rows)));
        }
        let mut top_left = col_index_to_letter(pane.cols + 1);
        top_left.push_str(ibuf.format(pane.rows + 1));
        pane_elem.push_attribute(("topLeftCell", top_left.as_str()));
        let active_pane = match (pane.rows > 0, pane.cols > 0) {
            (true, true) => "bottomRight",
            (true, false) => "bottomLeft",
            (false, true) => "topRight",
            (false, false) => "bottomLeft",
        };
        pane_elem.push_attribute(("activePane", active_pane));
        pane_elem.push_attribute(("state", "frozen"));
        writer.write_event(Event::Empty(pane_elem)).map_err(map_err)?;
    }

    // Write <selection> elements for each pane selection.
    for sel in &self.pane_selections {
        let mut sel_elem = BytesStart::new("selection");
        if let Some(ref p) = sel.pane {
            sel_elem.push_attribute(("pane", p.as_str()));
        }
        if let Some(ref ac) = sel.active_cell {
            sel_elem.push_attribute(("activeCell", ac.as_str()));
        }
        if let Some(ref sq) = sel.sqref {
            sel_elem.push_attribute(("sqref", sq.as_str()));
        }
        writer.write_event(Event::Empty(sel_elem)).map_err(map_err)?;
    }

    writer.write_event(Event::End(BytesEnd::new("sheetView"))).map_err(map_err)?;
    writer.write_event(Event::End(BytesEnd::new("sheetViews"))).map_err(map_err)?;
}
```

**Step 4:** Run `cargo test -p modern-xlsx-core` → all tests pass (~225 total).

**Step 5:** Commit:
```bash
git add crates/modern-xlsx-core/src/ooxml/worksheet.rs
git commit -m "feat: write split pane and selection elements in sheet XML"
```

---

## Task 4: WASM Rebuild

**Step 1:** Rebuild WASM to include the new struct fields in the JSON bridge:
```bash
cd crates/modern-xlsx-wasm && wasm-pack build --target web --release --out-dir ../../packages/modern-xlsx/wasm --no-opt
```

**Step 2:** Verify build succeeds, no errors.

**Step 3:** Commit:
```bash
git add packages/modern-xlsx/wasm/
git commit -m "build: rebuild WASM with split pane support"
```

---

## Task 5: TypeScript — Split Pane Types

**Files:**
- Modify: `packages/modern-xlsx/src/types.ts`

**Step 1:** Add types after `FrozenPane` (~line 80):

```typescript
export interface SplitPaneData {
  /** Horizontal split position in twips (ySplit). */
  horizontal?: number | null;
  /** Vertical split position in twips (xSplit). */
  vertical?: number | null;
  /** Cell reference for top-left cell in bottom-right pane. */
  topLeftCell?: string | null;
  /** Active pane: `"topLeft"` | `"topRight"` | `"bottomLeft"` | `"bottomRight"`. */
  activePane?: string | null;
}

export interface PaneSelectionData {
  /** Which pane this selection belongs to. */
  pane?: string | null;
  /** Active cell reference, e.g. `"A1"`. */
  activeCell?: string | null;
  /** Selected range, e.g. `"A1:C5"`. */
  sqref?: string | null;
}
```

**Step 2:** Add fields to `WorksheetData` (~line 278, after `frozenPane`):

```typescript
  splitPane?: SplitPaneData | null;
  paneSelections?: PaneSelectionData[];
```

**Step 3:** Run `pnpm -C packages/modern-xlsx typecheck` → passes.

**Step 4:** Commit:
```bash
git add packages/modern-xlsx/src/types.ts
git commit -m "feat: add SplitPaneData and PaneSelectionData TypeScript types"
```

---

## Task 6: TypeScript — Split Pane API

**Files:**
- Modify: `packages/modern-xlsx/src/workbook.ts`
- Modify: `packages/modern-xlsx/src/index.ts` (exports)

**Step 1:** Add imports in `workbook.ts` (line ~18):

```typescript
import type { SplitPaneData, PaneSelectionData } from './types.js';
```

**Step 2:** Add getter/setter on `Worksheet` class (after `frozenPane` at ~line 630):

```typescript
  // --- Split pane ---

  /** Returns the split pane configuration, or `null` if no split pane is active. */
  get splitPane(): SplitPaneData | null {
    return this.data.worksheet.splitPane ?? null;
  }

  /** Sets a split pane configuration. Setting a split pane clears any frozen pane, and vice versa. */
  set splitPane(pane: SplitPaneData | null) {
    this.data.worksheet.splitPane = pane;
    if (pane) {
      this.data.worksheet.frozenPane = null;
    }
  }

  /** Sets a frozen pane. Setting a frozen pane clears any split pane. */
  set frozenPane(pane: FrozenPane | null) {
    this.data.worksheet.frozenPane = pane;
    if (pane) {
      this.data.worksheet.splitPane = null;
    }
  }

  // --- Pane selections ---

  /** Returns per-pane selection state, or empty array. */
  get paneSelections(): PaneSelectionData[] {
    return this.data.worksheet.paneSelections ?? [];
  }

  /** Sets per-pane selection state. */
  set paneSelections(selections: PaneSelectionData[]) {
    this.data.worksheet.paneSelections = selections;
  }
```

> **Note:** The existing `frozenPane` setter (line ~628) must be replaced with the version above that clears `splitPane`.

**Step 3:** Export new types from `index.ts`:

```typescript
export type { SplitPaneData, PaneSelectionData } from './types.js';
```

**Step 4:** Run `pnpm -C packages/modern-xlsx typecheck` → passes.

**Step 5:** Commit:
```bash
git add packages/modern-xlsx/src/workbook.ts packages/modern-xlsx/src/index.ts packages/modern-xlsx/src/types.ts
git commit -m "feat: add splitPane and paneSelections API on Worksheet"
```

---

## Task 7: TypeScript — Split Pane Tests

**Files:**
- Modify: `packages/modern-xlsx/__tests__/api-setters.test.ts` or create `packages/modern-xlsx/__tests__/split-pane.test.ts`

**Step 1:** Write tests:

```typescript
import { describe, it, expect } from 'vitest';
import { Workbook, readBuffer } from '../src/index.js';

describe('Split Pane', () => {
  it('set and get horizontal split pane', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    ws.splitPane = { horizontal: 2400, topLeftCell: 'A5', activePane: 'bottomLeft' };
    expect(ws.splitPane).toEqual({ horizontal: 2400, topLeftCell: 'A5', activePane: 'bottomLeft' });
  });

  it('set and get vertical split pane', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    ws.splitPane = { vertical: 3000, topLeftCell: 'D1', activePane: 'topRight' };
    expect(ws.splitPane?.vertical).toBe(3000);
  });

  it('set and get four-way split pane', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    ws.splitPane = { horizontal: 2400, vertical: 3000, topLeftCell: 'D5', activePane: 'bottomRight' };
    ws.paneSelections = [
      { pane: 'topRight', activeCell: 'D1', sqref: 'D1' },
      { pane: 'bottomLeft', activeCell: 'A5', sqref: 'A5' },
      { pane: 'bottomRight', activeCell: 'D5', sqref: 'D5' },
    ];
    expect(ws.paneSelections).toHaveLength(3);
  });

  it('split pane clears frozen pane', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    ws.frozenPane = { rows: 1, cols: 0 };
    expect(ws.frozenPane).toBeTruthy();
    ws.splitPane = { horizontal: 2400 };
    expect(ws.frozenPane).toBeNull();
    expect(ws.splitPane).toBeTruthy();
  });

  it('frozen pane clears split pane', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    ws.splitPane = { horizontal: 2400 };
    expect(ws.splitPane).toBeTruthy();
    ws.frozenPane = { rows: 1, cols: 0 };
    expect(ws.splitPane).toBeNull();
    expect(ws.frozenPane).toBeTruthy();
  });

  it('horizontal split pane survives roundtrip', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    ws.cell('A1').value = 'data';
    ws.splitPane = { horizontal: 2400, topLeftCell: 'A5', activePane: 'bottomLeft' };
    ws.paneSelections = [{ pane: 'bottomLeft', activeCell: 'A5', sqref: 'A5' }];

    const buffer = await wb.toBuffer();
    const wb2 = await readBuffer(buffer);
    const ws2 = wb2.getSheet('Sheet1')!;
    expect(ws2.splitPane?.horizontal).toBe(2400);
    expect(ws2.splitPane?.topLeftCell).toBe('A5');
    expect(ws2.paneSelections).toHaveLength(1);
  });

  it('vertical split pane survives roundtrip', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    ws.cell('A1').value = 'data';
    ws.splitPane = { vertical: 3000, topLeftCell: 'D1', activePane: 'topRight' };

    const buffer = await wb.toBuffer();
    const wb2 = await readBuffer(buffer);
    const ws2 = wb2.getSheet('Sheet1')!;
    expect(ws2.splitPane?.vertical).toBe(3000);
    expect(ws2.splitPane?.topLeftCell).toBe('D1');
  });

  it('four-way split pane survives roundtrip', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    ws.cell('A1').value = 'data';
    ws.splitPane = { horizontal: 2400, vertical: 3000, topLeftCell: 'D5', activePane: 'bottomRight' };
    ws.paneSelections = [
      { pane: 'topRight', activeCell: 'D1', sqref: 'D1' },
      { pane: 'bottomLeft', activeCell: 'A5', sqref: 'A5' },
      { pane: 'bottomRight', activeCell: 'D5', sqref: 'D5' },
    ];

    const buffer = await wb.toBuffer();
    const wb2 = await readBuffer(buffer);
    const ws2 = wb2.getSheet('Sheet1')!;
    expect(ws2.splitPane?.horizontal).toBe(2400);
    expect(ws2.splitPane?.vertical).toBe(3000);
    expect(ws2.paneSelections).toHaveLength(3);
  });

  it('clear split pane', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    ws.splitPane = { horizontal: 2400 };
    ws.splitPane = null;
    expect(ws.splitPane).toBeNull();
  });

  it('default (no pane) → both null', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    expect(ws.frozenPane).toBeNull();
    expect(ws.splitPane).toBeNull();
    expect(ws.paneSelections).toEqual([]);
  });
});
```

**Step 2:** Run `pnpm -C packages/modern-xlsx test` → all tests pass.

**Step 3:** Commit:
```bash
git add packages/modern-xlsx/__tests__/
git commit -m "test: add split pane TypeScript roundtrip tests"
```

---

## Task 8: Lint, Build & Full Verification

**Step 1:** Run full lint + typecheck + build:
```bash
pnpm -C packages/modern-xlsx lint
pnpm -C packages/modern-xlsx typecheck
pnpm -C packages/modern-xlsx build
cargo clippy -p modern-xlsx-core -- -D warnings
```

**Step 2:** Run full test suites:
```bash
cargo test -p modern-xlsx-core
pnpm -C packages/modern-xlsx test
```

**Step 3:** Fix any issues found.

**Step 4:** Commit any fixes:
```bash
git commit -m "fix: address lint and clippy issues from split pane implementation"
```

---

## Task 9: Codebase Audit — Modernization & Performance Pass

Audit all touched files + surrounding code for:

1. **Rust modernization (1.94 Edition 2024):** let-else, if-let chains, `LazyLock`, explicit `use` paths
2. **Performance:** `Vec::with_capacity`, `std::str::from_utf8().unwrap_or_default()` not `from_utf8_lossy().into_owned()`, `itoa::Buffer` for int formatting, binary insert for ordered collections
3. **TypeScript modernization (6.0):** using declarations, `satisfies`, explicit return types where beneficial
4. **Serde:** all structs use `#[serde(rename_all = "camelCase")]`, `skip_serializing_if` for Option/Vec
5. **No regressions in existing tests**

This audit covers files modified in this batch + adjacent code in the same modules.
