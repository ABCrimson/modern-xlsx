# Advanced Features Roadmap — 0.4.x & 0.5.x Implementation Plan

> **For Claude:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task.

**Goal:** Implement split panes, sheet views, workbook protection, sheet management, sparklines, data tables, and external links — bringing modern-xlsx to full OOXML feature coverage.

**Architecture:** Extend existing Rust SAX parser/writer + TypeScript API pattern. Each feature follows: Rust struct → parser → writer → WASM bridge → TS type → TS API → tests.

**Tech Stack:** Rust 1.94 (quick-xml 0.39.2), TypeScript 6.0, Vitest 4.1, wasm-bindgen 0.2.114

---

## Batch 1: Split Panes (0.4.0 + 0.4.1)

### Task 1: Rust — Split Pane Data Model

**Files:**
- Modify: `crates/modern-xlsx-core/src/ooxml/worksheet.rs`

**Step 1:** Add `SplitPane` struct alongside existing `FrozenPane`:
```rust
#[derive(Debug, Clone, Serialize, Deserialize, PartialEq)]
#[serde(rename_all = "camelCase")]
pub struct SplitPane {
    pub horizontal: Option<f64>,   // ySplit in twips
    pub vertical: Option<f64>,     // xSplit in twips
    pub top_left_cell: Option<String>,
    pub active_pane: Option<String>, // topLeft, topRight, bottomLeft, bottomRight
}
```

**Step 2:** Add `PaneSelection` struct for per-pane selection:
```rust
#[derive(Debug, Clone, Serialize, Deserialize, PartialEq)]
#[serde(rename_all = "camelCase")]
pub struct PaneSelection {
    pub pane: Option<String>,
    pub active_cell: Option<String>,
    pub sqref: Option<String>,
}
```

**Step 3:** Add fields to `WorksheetXml`:
```rust
pub split_pane: Option<SplitPane>,
pub pane_selections: Vec<PaneSelection>,
```

### Task 2: Rust — Split Pane Parser

**Files:**
- Modify: `crates/modern-xlsx-core/src/ooxml/worksheet.rs`

- In `<pane>` parsing (both fast + streaming paths): check `state` attribute
  - `state="frozen"` → existing `FrozenPane` path
  - `state="split"` or no state → new `SplitPane` path with `xSplit`/`ySplit` as f64 twips
- Parse `topLeftCell` and `activePane` attributes for split panes
- Parse `<selection>` elements after `<pane>`: extract `pane`, `activeCell`, `sqref`
- Ensure frozen and split are mutually exclusive

### Task 3: Rust — Split Pane Writer

**Files:**
- Modify: `crates/modern-xlsx-core/src/ooxml/worksheet.rs`

- In `write_sheet_views()`: if `split_pane.is_some()`, write `<pane>` with:
  - `xSplit` / `ySplit` as float values (twips)
  - `topLeftCell` attribute
  - `activePane` attribute
  - `state="split"`
- Write `<selection>` elements for each `PaneSelection`
- Ensure split pane path is separate from frozen pane path

### Task 4: Rust — Split Pane Tests

**Files:**
- Modify: `crates/modern-xlsx-core/src/ooxml/worksheet.rs` (test module)

- Test: horizontal split (ySplit only) → parse → verify twip value
- Test: vertical split (xSplit only) → parse → verify
- Test: four-way split (both xSplit + ySplit) → parse → verify all 4 pane selections
- Test: split vs frozen mutually exclusive
- Test: roundtrip write → parse for each split type

### Task 5: TypeScript — Split Pane Types & API

**Files:**
- Modify: `packages/modern-xlsx/src/types.ts`
- Modify: `packages/modern-xlsx/src/workbook.ts`

- Add `SplitPane` and `PaneSelection` interfaces to types.ts
- Add `Worksheet.splitPane` getter/setter
- Add `Worksheet.paneSelections` getter
- Export from index.ts

### Task 6: TypeScript — Split Pane Tests

**Files:**
- Modify: `packages/modern-xlsx/__tests__/workbook.test.ts`

- Test: set horizontal split → write → read → verify twip value preserved
- Test: set vertical split → roundtrip
- Test: set four-way split → roundtrip with all selections
- Test: split pane and frozen pane mutually exclusive (setting one clears the other)
- Test: default (no pane) → verify both null

---

## Batch 2: Sheet Views (0.4.2 + 0.4.3)

### Task 7: Rust — SheetView Data Model

**Files:**
- Modify: `crates/modern-xlsx-core/src/ooxml/worksheet.rs`

- Add `SheetViewData` struct with all ECMA-376 attributes:
  - Booleans: `show_grid_lines`, `show_row_col_headers`, `show_zeros`, `right_to_left`, `tab_selected`, `show_ruler`, `show_outline_symbols`, `show_white_space`, `default_grid_color`
  - Numerics: `zoom_scale`, `zoom_scale_normal`, `zoom_scale_page_layout_view`, `zoom_scale_sheet_layout_view`, `color_id`
  - Enum: `view` ("normal" | "pageBreakPreview" | "pageLayout")
- Add `sheet_view: Option<SheetViewData>` to `WorksheetXml`

### Task 8: Rust — SheetView Parser & Writer

- Parse all `<sheetView>` attributes in both fast + streaming paths
- Write `<sheetView>` with all non-default attributes
- Always write `<sheetViews>` if `sheet_view` or `frozen_pane` or `split_pane` present
- Omit default values (showGridLines=true is default → don't write)

### Task 9: Rust — SheetView Tests

- Test: hide gridlines + zoom 150% → roundtrip
- Test: RTL view → roundtrip
- Test: pageBreakPreview mode → roundtrip
- Test: pageLayout mode with custom zoom → roundtrip
- Test: default view (all defaults) → minimal XML output

### Task 10: TypeScript — SheetView Types & API

- Add `SheetViewData` interface to types.ts
- `Worksheet.view` getter/setter → `SheetViewData | null`
- `Worksheet.viewMode` convenience getter/setter → `'normal' | 'pageBreakPreview' | 'pageLayout'`
- TypeScript roundtrip tests (+5)

---

## Batch 3: Workbook Protection (0.4.4 + 0.4.5)

### Task 11: Rust — WorkbookProtection Data Model & Parser

**Files:**
- Modify: `crates/modern-xlsx-core/src/ooxml/workbook.rs`
- Modify: `crates/modern-xlsx-core/src/lib.rs`

- Add `WorkbookProtection` struct: `lock_structure`, `lock_windows`, `lock_revision`, hash attributes
- Add to `WorkbookXml` struct
- Parse `<workbookProtection>` element in workbook.xml
- Write after `<sheets>` element

### Task 12: Rust — Password Hash + Tests

- Implement SHA-512 password hashing (salt + spin count) matching Excel format
- Tests for parsing, writing, password hashing, roundtrip

### Task 13: TypeScript — WorkbookProtection API

- `WorkbookProtectionData` interface
- `Workbook.protection` getter/setter
- Password option generates hash attributes
- Tests (+5)

---

## Batch 4: Sheet Management (0.4.6 + 0.4.7 + 0.4.8)

### Task 14: Sheet State Validation

- Sheet state already parsed/written (confirmed by exploration)
- Add validation: cannot hide all sheets (at least one must remain visible)
- TypeScript: `Worksheet.state` getter/setter (if not already exposed)

### Task 15: Sheet Ordering & Management

- `Workbook.moveSheet(fromIndex, toIndex)`
- `Workbook.cloneSheet(index, newName?)`
- `Workbook.renameSheet(index, newName)`
- Update defined names referencing moved/renamed sheets
- Tests (+10)

### Task 16: Custom Sheet Properties

- Parse `<sheetPr>` attributes: `codeName`, `filterMode`, `published`, etc.
- Parse `<pageSetUpPr>`: `autoPageBreaks`, `fitToPage`
- `Worksheet.properties` getter/setter
- Tests (+5)

---

## Batch 5: Integration Tests (0.4.9)

### Task 17: View & Layout Integration Tests

- Combined feature roundtrips (frozen + hidden gridlines + zoom + page layout)
- Multi-sheet: workbook protection + hidden sheets + sheet protection
- Print-ready: margins + headers/footers + print titles
- Cross-feature: split pane on one sheet, frozen on another

---

## Batch 6: Sparklines (0.5.0–0.5.4)

### Task 18: Sparkline Rust Model & Parser

- `SparklineGroupData` + `SparklineData` structs
- Parse `<extLst>` → `<ext uri="{05C60535...}">` → `<x14:sparklineGroups>`
- Handle x14 namespace

### Task 19: Sparkline Writer

- Write sparkline extension block in `<extLst>`
- Merge with existing extension entries
- Roundtrip tests

### Task 20: Sparkline TypeScript API

- `Worksheet.addSparkline()`, `Worksheet.sparklines` getter
- `SparklineBuilder` fluent API
- Style presets, marker toggles, line weight
- Tests (+20)

---

## Batch 7: Data Tables & External Links (0.5.5–0.5.8)

### Task 21: Data Tables (What-If)

- Parse `r1`/`r2` attributes on formula elements
- `DataTableData` struct
- Preservation through roundtrip
- Tests (+10)

### Task 22: External Links & Custom XML

- Parse `xl/externalLinks/externalLink{n}.xml`
- Preserve through roundtrip (opaque + parsed metadata)
- Custom XML passthrough
- Tests (+10)

---

## Batch 8: Data Feature Integration (0.5.9)

### Task 23: Integration Tests

- Sparklines + data tables + external links in same workbook
- Cross-sheet sparkline references
- Performance regression checks
