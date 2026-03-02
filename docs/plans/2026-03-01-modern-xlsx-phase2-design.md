# modern-xlsx Phase 2 Design

<div align="center">

**Full Feature Implementation**

Dependencies, data model, features, performance, and distribution

</div>

---

|  |  |
|---|---|
| **Date** | 2026-03-01 |
| **Status** | Completed |
| **Approach** | Foundation-first — fix data model, then build features, then distribution |

---

## Goal

Restore 5 workspace dependencies with best-practice usage, unify the data model, add core XLSX features (formulas, rich text, data validation, named ranges, style builder), improve performance (wasm-opt, itoa, streaming, parallel parsing), and set up distribution (browser testing, npm publish pipeline).

---

## Tech Stack

| Category | Dependencies |
|---|---|
| **Core** | Rust 1.94.0 (Edition 2024), TypeScript 6.0, wasm-bindgen 0.2.114 |
| **Restored** | serde_json 1.0, log 0.4, pretty_assertions 1.4, js-sys 0.3, web-sys 0.3 |
| **New** | itoa 1.0 (zero-alloc integer formatting), console_log 1.0 (WASM log backend) |

---

## Layer 1 — Restore Dependencies + Infrastructure

### 1a. Workspace Dependencies

| Dependency | Purpose |
|---|---|
| `serde_json` | Golden file tests, optional JSON export |
| `log` | Structured logging facade (zero-cost when no backend) |
| `pretty_assertions` | Colored struct diffs in test failures |
| `js-sys` | `Uint8Array` for zero-copy write return, `Date` utilities |
| `web-sys` | `Blob` + `BlobPropertyBag` for browser download, `console` log backend |

### 1b. wasm-opt Configuration

```toml
[package.metadata.wasm-pack.profile.release]
wasm-opt = ["-Oz", "--enable-bulk-memory", "--enable-nontrapping-float-to-int"]
```

### 1c. itoa Integration

Zero-alloc integer formatting in `writer.rs`, `worksheet.rs`, `content_types.rs`, `relationships.rs`.

### 1d. number_format.rs Allocation Cleanup

Replace repeated string allocations with `write!` to a single buffer. Byte-level scanning, `Cow<str>` for format strings.

---

## Layer 2 — Data Model Unification

| Change | Details |
|---|---|
| **SST resolution** | `readBuffer()` auto-resolves shared string indices to strings |
| **Unified WorkbookData** | Single canonical type; reader populates `shared_strings`, writer ignores it |
| **to_xml() returns Vec\<u8\>** | Skip String intermediary — write bytes directly |

---

## Layer 3 — Core Features

| Feature | Description |
|---|---|
| **Formula preservation** | Read/write `<f>` elements — no evaluation, just preservation |
| **Rich text** | `<si><r><rPr>...</rPr><t>...</t></r></si>` in shared strings |
| **Data validation** | `<dataValidation>` elements: list, whole, decimal, date, textLength, custom |
| **Named ranges** | `definedNames` from workbook.xml with TypeScript API |
| **Style builder** | Fluent chainable API: `.font()`, `.fill()`, `.border()`, `.numberFormat()`, `.build()` |

---

## Layer 4 — Advanced Features

| Feature | Description |
|---|---|
| **Streaming read/write** | Row-by-row parsing/writing for 100K+ row files |
| **Images/charts** | Read as opaque blobs, write back unchanged — roundtrip preservation |
| **Conditional formatting** | `<conditionalFormatting>` rules: cellIs, colorScale, dataBar, iconSet |

---

## Layer 5 — Distribution + Performance

| Feature | Description |
|---|---|
| **Browser testing** | Vitest browser mode with Playwright |
| **npm publish** | GitHub Actions: build WASM → test → publish on tag push |
| **Parallel parsing** | Optional `rayon` feature flag for multi-sheet parallel read |
| **tsdown externals** | Properly externalize WASM imports in bundle |

---

## Testing Strategy

| Strategy | Tools |
|---|---|
| TDD | Tests before implementation for every feature |
| Golden files | `serde_json` snapshots for complex XML roundtrips |
| Colored diffs | `pretty_assertions` in all Rust test modules |
| Browser tests | WASM-specific features (Blob, Uint8Array) |
| Integration | TypeScript tests for every new API surface |
