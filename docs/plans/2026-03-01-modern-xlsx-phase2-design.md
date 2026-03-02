# modern-xlsx Phase 2 — Full Feature Implementation Design

## Goal

Restore 5 workspace dependencies with best-practice usage, unify the data model, add core XLSX features (formulas, rich text, data validation, named ranges, style builder), improve performance (wasm-opt, itoa, streaming, parallel parsing), and set up distribution (browser testing, npm publish pipeline).

## Architecture

Foundation-first approach: fix the data model and infrastructure first, then build features on top, then distribution. Each layer builds on the previous one with no rework.

## Tech Stack

- Rust 1.94.0 (edition 2024), TypeScript 6.0, wasm-bindgen 0.2.114
- Restored deps: serde_json 1.0, log 0.4, pretty_assertions 1.4, js-sys 0.3, web-sys 0.3
- New deps: itoa (zero-alloc integer formatting), console_log (WASM log backend)

---

## Layer 1: Restore Dependencies + Infrastructure

### 1a. Restore 5 workspace deps with exact versions

Add back to workspace Cargo.toml:
- `serde_json = "1.0"`
- `log = "0.4"`
- `pretty_assertions = "1.4"`
- `js-sys = "0.3"`
- `web-sys = { version = "0.3", features = ["Blob", "BlobPropertyBag", "console", "Performance"] }`

### 1b. Best usage for each restored dep

**serde_json:**
- modern-xlsx-core dev-dependency: Golden file tests for reader/writer/styles (serialize WorkbookData to JSON, compare against snapshots)
- modern-xlsx-core optional feature `json-export`: `WorkbookData::to_json()` for debugging/interop

**log:**
- modern-xlsx-core dependency: Structured logging facade in reader.rs (warn on malformed XML), writer.rs (debug sheet serialization), zip/reader.rs (warn on unknown entries)
- Zero runtime cost when no backend is configured (production WASM)

**pretty_assertions:**
- modern-xlsx-core dev-dependency: Replace `assert_eq!` with `pretty_assertions::assert_eq!` in reader, writer, and styles tests for colored struct diffs on failure

**js-sys:**
- modern-xlsx-wasm dependency: `js_sys::Uint8Array` for zero-copy write return (instead of Vec<u8> → JsValue copy), `js_sys::Date` for date conversion utilities

**web-sys:**
- modern-xlsx-wasm dependency: `web_sys::Blob` + `BlobPropertyBag` for `write_blob()` browser download helper, `web_sys::console` as log backend

### 1c. wasm-opt configuration

Add to `crates/modern-xlsx-wasm/Cargo.toml`:
```toml
[package.metadata.wasm-pack.profile.release]
wasm-opt = ["-Oz", "--enable-bulk-memory", "--enable-nontrapping-float-to-int"]
```

Expected: 601KB → ~480-510KB uncompressed.

### 1d. itoa for zero-alloc integer formatting

Add `itoa = "1.0"` to workspace. Use in writer.rs and worksheet.rs `to_xml()` methods where row/col numbers are formatted to strings.

### 1e. number_format.rs allocation cleanup

Replace repeated string allocations with write! to a single buffer. Use Cow<str> for format strings that may or may not need modification.

---

## Layer 2: Data Model Unification

### 2a. SST resolution in TypeScript read path

`readBuffer()` auto-resolves shared string indices to actual string values before constructing the Workbook. Cell values come back as strings, not SST indices.

### 2b. Unify reader/writer WorkbookData

Single canonical `WorkbookData` type with optional `shared_strings` field. Reader populates it, writer ignores it (builds SST internally). Remove any dual-type ambiguity.

### 2c. to_xml() returns Vec<u8> directly

All `to_xml()` methods in OOXML modules return `Vec<u8>` instead of `String`. Eliminates the `to_string().into_bytes()` copy. Writer already writes bytes to ZIP.

---

## Layer 3: Core Features

### 3a. Formula preservation

Read and write formula cells (`<f>` elements in worksheet XML). TypeScript Cell API already has `formula` getter/setter. Rust reader/writer need to handle the `<f>` tag. No evaluation — just preservation.

### 3b. Rich text / formatted runs

Support `<si><r><rPr>...</rPr><t>...</t></r></si>` in shared strings. TypeScript types for `RichTextRun { text, font?, bold?, italic?, color? }`. Reader parses runs, writer serializes them.

### 3c. Data validation rules

Read/write `<dataValidation>` elements in worksheet XML. TypeScript types for validation rules (list, whole, decimal, date, textLength, custom). Worksheet API: `ws.addValidation(ref, rule)`.

### 3d. Named ranges in TypeScript API

Expose `definedNames` from workbook.xml. Workbook API: `wb.namedRanges` getter, `wb.addNamedRange(name, ref)`, `wb.getNamedRange(name)`.

### 3e. Style builder API

Fluent chainable API for creating styles:
```typescript
const style = wb.createStyle()
  .font({ bold: true, size: 14, color: '#FF0000' })
  .fill({ pattern: 'solid', fgColor: '#FFFF00' })
  .border({ bottom: { style: 'thin', color: '#000000' } })
  .numberFormat('0.00%')
  .build();
cell.styleIndex = style;
```

---

## Layer 4: Advanced Features

### 4a. Streaming read/write

Streaming reader: parse worksheets row-by-row without loading entire XML into memory. Streaming writer: write rows incrementally to ZIP entry. For large files (100K+ rows).

### 4b. Images/charts read-preserve

Read `drawing*.xml` and media files from XLSX, store as opaque blobs, write them back unchanged. Preserves images/charts through round-trip without creating new ones.

### 4c. Conditional formatting rules

Read/write `<conditionalFormatting>` elements. Support common rule types: cellIs, colorScale, dataBar, iconSet. TypeScript types and worksheet API.

---

## Layer 5: Distribution + Performance

### 5a. Browser testing

Vitest browser mode with Playwright. Test WASM initialization, round-trip read/write, and Blob download helper in actual browser environment.

### 5b. npm publish pipeline

GitHub Actions workflow: build WASM, run tests, publish to npm. Automated version bumping and changelog generation.

### 5c. Parallel sheet parsing (rayon)

Optional `rayon` feature flag on modern-xlsx-core. Parse multiple worksheets in parallel during read. Not available in WASM (single-threaded), only for native Rust usage.

### 5d. tsdown external config fix

Configure tsdown to properly externalize WASM imports so the bundle doesn't try to inline the .wasm file.

---

## Testing Strategy

- Every feature gets tests before implementation (TDD)
- Golden file tests with serde_json for complex XML round-trips
- pretty_assertions in all Rust test modules
- Browser tests for WASM-specific features (Blob, Uint8Array)
- Integration tests in TypeScript for every new API surface
