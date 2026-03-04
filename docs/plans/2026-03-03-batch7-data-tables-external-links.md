# Data Tables & External Links (0.5.5–0.5.8) — TDD Implementation Plan

> **For Claude:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task.

**Goal:** Add data table formula attributes (`r1`/`r2`/`dt2D`/`dtr1`/`dtr2`) to cell parsing/writing, and fix content type handling for external links and custom XML preservation.

**Architecture:** Extend `Cell` struct with 5 optional fields, update 4 SAX parser locations and the cell writer. External links/custom XML already roundtrip via `preserved_entries` — just need content type overrides for proper ZIP compliance.

**Tech Stack:** Rust 1.94 (quick-xml 0.39.2, serde camelCase), TypeScript 6.0, Vitest 4.1

---

## Task 1: Rust — Data table formula attributes

**Files:** `crates/modern-xlsx-core/src/ooxml/worksheet.rs`

### Cell struct changes:

Add to the `Cell` struct:

```rust
/// Data table 1st input cell reference.
#[serde(default, skip_serializing_if = "Option::is_none")]
pub formula_r1: Option<String>,
/// Data table 2nd input cell reference.
#[serde(default, skip_serializing_if = "Option::is_none")]
pub formula_r2: Option<String>,
/// 2D data table flag.
#[serde(default, skip_serializing_if = "Option::is_none")]
pub formula_dt2d: Option<bool>,
/// Data table row input deleted flag.
#[serde(default, skip_serializing_if = "Option::is_none")]
pub formula_dtr1: Option<bool>,
/// Data table column input deleted flag.
#[serde(default, skip_serializing_if = "Option::is_none")]
pub formula_dtr2: Option<bool>,
```

### Parser changes (4 locations):

In all 4 formula parsing blocks (`InCellFormula` state in parse_with_sst, parse_to_json, and their duplicates), add attribute handling:

```rust
b"r1" => cur_cell_formula_r1 = Some(val.to_owned()),
b"r2" => cur_cell_formula_r2 = Some(val.to_owned()),
b"dt2D" => cur_cell_formula_dt2d = if val == "1" { Some(true) } else { None },
b"dtr1" => cur_cell_formula_dtr1 = if val == "1" { Some(true) } else { None },
b"dtr2" => cur_cell_formula_dtr2 = if val == "1" { Some(true) } else { None },
```

Add temp variables alongside existing formula temps:
```rust
let mut cur_cell_formula_r1: Option<String> = None;
let mut cur_cell_formula_r2: Option<String> = None;
let mut cur_cell_formula_dt2d: Option<bool> = None;
let mut cur_cell_formula_dtr1: Option<bool> = None;
let mut cur_cell_formula_dtr2: Option<bool> = None;
```

Wire into Cell construction and reset after each cell.

### Writer changes:

In the `<f>` element writer section, add attributes:

```rust
if let Some(ref r1) = cell.formula_r1 { f_elem.push_attribute(("r1", r1.as_str())); }
if let Some(ref r2) = cell.formula_r2 { f_elem.push_attribute(("r2", r2.as_str())); }
if cell.formula_dt2d == Some(true) { f_elem.push_attribute(("dt2D", "1")); }
if cell.formula_dtr1 == Some(true) { f_elem.push_attribute(("dtr1", "1")); }
if cell.formula_dtr2 == Some(true) { f_elem.push_attribute(("dtr2", "1")); }
```

### All Cell struct literals:

Add the 5 new fields as `None` to every `Cell { ... }` literal across the codebase.

### Tests (4):

1. Parse data table formula with r1/r2
2. Parse 2D data table (dt2D + dtr1 + dtr2)
3. Data table formula roundtrip (write → parse)
4. Normal formula ignores data table attributes

---

## Task 2: Rust — Content type overrides for external links & custom XML

**Files:** `crates/modern-xlsx-core/src/writer.rs`

In the preserved entries content type auto-detection section, add:

```rust
// External links
if path.starts_with("xl/externalLinks/externalLink") && path.ends_with(".xml") {
    content_types.add_override(
        &format!("/{path}"),
        "application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml",
    );
}
// Custom XML
if path.starts_with("customXml/item") && path.ends_with(".xml") && !path.contains("Props") {
    content_types.add_override(
        &format!("/{path}"),
        "application/xml",
    );
}
if path.starts_with("customXml/itemProps") && path.ends_with(".xml") {
    content_types.add_override(
        &format!("/{path}"),
        "application/vnd.openxmlformats-officedocument.customXmlProperties+xml",
    );
}
```

### Tests (2):

1. External link preserved entry roundtrip
2. Custom XML preserved entry roundtrip (verify content types)

---

## Task 3: WASM rebuild + TypeScript types

### WASM rebuild

### TypeScript types (types.ts):

Add to `CellData`:
```typescript
formulaR1?: string | null;
formulaR2?: string | null;
formulaDt2d?: boolean | null;
formulaDtr1?: boolean | null;
formulaDtr2?: boolean | null;
```

No new API methods needed — cell properties already accessible via `cell().formula`.

---

## Task 4: TypeScript tests + verification

### Tests (6):

1. Data table formula r1/r2 roundtrip
2. Data table 2D (dt2D) roundtrip
3. Normal formula unaffected by data table attributes
4. External link preservation roundtrip (via preserved entries)
5. Clear data table formula
6. Full verification (lint, typecheck, build, all tests)
