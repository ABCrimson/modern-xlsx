# modern-xlsx Phase 2 Implementation Plan

> **For Claude:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task.

**Goal:** Restore 5 workspace dependencies with best-practice usage, unify the data model, add core XLSX features, optimize performance, and set up distribution pipeline.

**Architecture:** Foundation-first — fix data model and infrastructure, then build features on top, then distribution. Rust core handles all OOXML parsing/writing; TypeScript wraps with ergonomic API.

**Tech Stack:** Rust 1.94.0 (edition 2024), TypeScript 6.0, wasm-bindgen 0.2.114, serde_json 1.0, log 0.4, pretty_assertions 1.4, js-sys 0.3, web-sys 0.3, itoa 1.0, console_log 1.0

---

### Task 1: Restore 5 workspace dependencies

**Files:**
- Modify: `Cargo.toml` (workspace root, lines 15-36)
- Modify: `crates/modern-xlsx-core/Cargo.toml`
- Modify: `crates/modern-xlsx-wasm/Cargo.toml`

**Step 1: Add deps back to workspace Cargo.toml**

Add after the existing `thiserror` entry (line 29):

```toml
# JSON serialization (golden file tests + optional JSON export)
serde_json = "1.0"

# Logging facade (zero-cost when no backend configured)
log = "0.4"

# Zero-alloc integer formatting for XML output
itoa = "1.0"

# WASM JS interop
js-sys = "0.3"
web-sys = { version = "0.3", features = ["Blob", "BlobPropertyBag", "console", "Performance"] }
```

Add to dev section after `wasm-bindgen-test`:

```toml
pretty_assertions = "1.4"
```

**Step 2: Add deps to modern-xlsx-core Cargo.toml**

```toml
[dependencies]
# ... existing deps ...
log.workspace = true
itoa.workspace = true

[dev-dependencies]
serde_json.workspace = true
pretty_assertions.workspace = true
```

**Step 3: Add deps to modern-xlsx-wasm Cargo.toml**

```toml
[dependencies]
# ... existing deps ...
js-sys.workspace = true
web-sys.workspace = true
log.workspace = true
```

**Step 4: Verify compilation**

Run: `cargo check --workspace`
Expected: compiles with no errors

**Step 5: Commit**

```bash
git add Cargo.toml crates/modern-xlsx-core/Cargo.toml crates/modern-xlsx-wasm/Cargo.toml
git commit -m "chore: restore workspace deps (serde_json, log, pretty_assertions, js-sys, web-sys, itoa)"
```

---

### Task 2: wasm-opt configuration

**Files:**
- Modify: `crates/modern-xlsx-wasm/Cargo.toml`

**Step 1: Add wasm-pack metadata**

Append to `crates/modern-xlsx-wasm/Cargo.toml`:

```toml
[package.metadata.wasm-pack.profile.release]
wasm-opt = ["-Oz", "--enable-bulk-memory", "--enable-nontrapping-float-to-int"]
```

**Step 2: Commit**

```bash
git add crates/modern-xlsx-wasm/Cargo.toml
git commit -m "perf: configure wasm-opt with bulk-memory and nontrapping-float"
```

---

### Task 3: Integrate `log` crate into Rust reader/writer

**Files:**
- Modify: `crates/modern-xlsx-core/src/reader.rs`
- Modify: `crates/modern-xlsx-core/src/writer.rs`
- Modify: `crates/modern-xlsx-core/src/zip/reader.rs`

**Step 1: Add log imports and calls to reader.rs**

At top: `use log::{debug, warn};`

Add logging at key points:
- `debug!("parsing workbook.xml");` before workbook parse
- `warn!("missing shared strings table");` when SST is absent
- `debug!("parsing worksheet: {}", sheet_name);` in the loop
- `warn!("unknown relationship target: {}", target);` for unresolved sheets

**Step 2: Add log imports and calls to writer.rs**

At top: `use log::{debug, trace};`

Add logging:
- `debug!("collecting shared strings from {} sheets", data.sheets.len());`
- `debug!("SST built: {} unique strings", sst_builder.len());`
- `trace!("writing ZIP entry: {}", entry_name);` for each ZIP entry

**Step 3: Add log imports to zip/reader.rs**

- `warn!("skipping unknown ZIP entry: {}", name);` for unrecognized files

**Step 4: Run tests**

Run: `cargo test -p modern-xlsx-core`
Expected: All tests pass (log is a no-op without a backend)

**Step 5: Commit**

```bash
git add crates/modern-xlsx-core/src/reader.rs crates/modern-xlsx-core/src/writer.rs crates/modern-xlsx-core/src/zip/reader.rs
git commit -m "feat: add structured logging via log crate in reader/writer/zip"
```

---

### Task 4: Integrate `pretty_assertions` in Rust tests

**Files:**
- Modify: `crates/modern-xlsx-core/src/reader.rs` (test module)
- Modify: `crates/modern-xlsx-core/src/writer.rs` (test module)
- Modify: `crates/modern-xlsx-core/src/ooxml/styles.rs` (test module)
- Modify: `crates/modern-xlsx-core/src/ooxml/worksheet.rs` (test module)
- Modify: `crates/modern-xlsx-core/src/ooxml/shared_strings.rs` (test module)
- Modify: `crates/modern-xlsx-core/src/ooxml/content_types.rs` (test module)
- Modify: `crates/modern-xlsx-core/src/ooxml/relationships.rs` (test module)

**Step 1: Add pretty_assertions to each test module**

In each file's `#[cfg(test)] mod tests { ... }`, add:
```rust
use pretty_assertions::assert_eq;
```

This shadows the standard `assert_eq!` with a version that shows colored diffs.

**Step 2: Run tests**

Run: `cargo test -p modern-xlsx-core`
Expected: All tests pass with colored output on failures

**Step 3: Commit**

```bash
git add crates/modern-xlsx-core/src/
git commit -m "test: use pretty_assertions for colored diffs in all test modules"
```

---

### Task 5: Integrate `itoa` for zero-alloc integer formatting

**Files:**
- Modify: `crates/modern-xlsx-core/src/writer.rs`
- Modify: `crates/modern-xlsx-core/src/ooxml/worksheet.rs`
- Modify: `crates/modern-xlsx-core/src/ooxml/content_types.rs`
- Modify: `crates/modern-xlsx-core/src/ooxml/relationships.rs`

**Step 1: Replace format!() integer formatting with itoa**

Wherever we do `format!("{}", some_integer)` or embed integers in XML strings, use:
```rust
let mut buf = itoa::Buffer::new();
let s = buf.format(row_index);
```

Key locations:
- worksheet.rs `to_xml()`: row index `r="{}"`, cell reference column+row
- content_types.rs `to_xml()`: sheet numbering
- relationships.rs `to_xml()`: rId numbering
- writer.rs: sheet path numbering

**Step 2: Run tests**

Run: `cargo test -p modern-xlsx-core`
Expected: All tests pass

**Step 3: Commit**

```bash
git add crates/modern-xlsx-core/src/
git commit -m "perf: use itoa for zero-alloc integer formatting in XML output"
```

---

### Task 6: Integrate `js-sys` for zero-copy write return

**Files:**
- Modify: `crates/modern-xlsx-wasm/src/lib.rs`

**Step 1: Change write() to return Uint8Array**

```rust
use js_sys::Uint8Array;

#[wasm_bindgen]
pub fn write(val: JsValue) -> Result<Uint8Array, JsError> {
    let data: WorkbookData = serde_wasm_bindgen::from_value(val)
        .map_err(|e| JsError::new(&e.to_string()))?;
    let bytes = modern_xlsx_core::writer::write_xlsx(&data)
        .map_err(|e| JsError::new(&e.to_string()))?;
    let arr = Uint8Array::new_with_length(bytes.len() as u32);
    arr.copy_from(&bytes);
    Ok(arr)
}
```

**Step 2: Update TypeScript wasm-loader.ts if needed**

The Uint8Array return should work transparently since wasm-bindgen already handles this conversion, but verify the TypeScript types match.

**Step 3: Run WASM tests + TS tests**

Run: `cargo test -p modern-xlsx-wasm` and `cd packages/modern-xlsx && pnpm test`
Expected: All pass

**Step 4: Commit**

```bash
git add crates/modern-xlsx-wasm/src/lib.rs packages/modern-xlsx/src/
git commit -m "perf: return Uint8Array from WASM write for zero-copy"
```

---

### Task 7: Integrate `web-sys` for Blob download helper + console log backend

**Files:**
- Modify: `crates/modern-xlsx-wasm/src/lib.rs`

**Step 1: Add write_blob() function**

```rust
use web_sys::{Blob, BlobPropertyBag};
use js_sys::Array;

#[wasm_bindgen]
pub fn write_blob(val: JsValue) -> Result<Blob, JsError> {
    let arr = write(val)?;
    let parts = Array::new();
    parts.push(&arr.buffer());
    let mut opts = BlobPropertyBag::new();
    opts.type_("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    Blob::new_with_buffer_source_sequence_and_options(&parts, &opts)
        .map_err(|e| JsError::new(&format!("{e:?}")))
}
```

**Step 2: Add console_log initialization**

```rust
use web_sys::console;

#[wasm_bindgen]
pub fn init_logging() {
    // Simple console.log-based log backend
    log::set_max_level(log::LevelFilter::Debug);
    // Use console_log crate or manual implementation
}
```

Note: We may need to add `console_log = "1.0"` as a dep or implement a minimal logger manually.

**Step 3: Export from TypeScript**

Add `write_blob` and `init_logging` to the TypeScript re-exports in `index.ts`.

**Step 4: Run tests**

Run: `cargo test -p modern-xlsx-wasm` and `cd packages/modern-xlsx && pnpm test`
Expected: All pass

**Step 5: Commit**

```bash
git add crates/modern-xlsx-wasm/ packages/modern-xlsx/src/
git commit -m "feat: add write_blob() for browser download and init_logging()"
```

---

### Task 8: number_format.rs allocation cleanup

**Files:**
- Modify: `crates/modern-xlsx-core/src/number_format.rs`

**Step 1: Replace string allocations with write! to buffer**

In `classify_format_string()`:
- Use `Cow<str>` for format strings that may not need modification
- Replace `.to_string()` + char-by-char building with direct byte scanning
- Use stack-allocated buffer for intermediate format cleaning

**Step 2: Run tests**

Run: `cargo test -p modern-xlsx-core`
Expected: All tests pass

**Step 3: Commit**

```bash
git add crates/modern-xlsx-core/src/number_format.rs
git commit -m "perf: reduce allocations in number_format classify_format_string"
```

---

### Task 9: to_xml() returns Vec<u8> directly

**Files:**
- Modify: `crates/modern-xlsx-core/src/ooxml/content_types.rs`
- Modify: `crates/modern-xlsx-core/src/ooxml/relationships.rs`
- Modify: `crates/modern-xlsx-core/src/ooxml/shared_strings.rs`
- Modify: `crates/modern-xlsx-core/src/ooxml/styles.rs`
- Modify: `crates/modern-xlsx-core/src/ooxml/workbook.rs`
- Modify: `crates/modern-xlsx-core/src/ooxml/worksheet.rs`
- Modify: `crates/modern-xlsx-core/src/writer.rs`

**Step 1: Change all to_xml() signatures from String to Vec<u8>**

For each OOXML module, change:
```rust
pub fn to_xml(&self) -> String {
    let mut xml = String::with_capacity(...);
```
to:
```rust
pub fn to_xml(&self) -> Vec<u8> {
    let mut xml = Vec::with_capacity(...);
    // Use write!(xml, ...) or xml.extend_from_slice(b"...")
```

**Step 2: Update writer.rs to use Vec<u8> directly**

The writer currently does `to_xml().into_bytes()` for ZIP entries. Change to just `to_xml()`.

**Step 3: Update all tests that check to_xml() output**

Tests that compare XML strings need to convert: `String::from_utf8(module.to_xml()).unwrap()`

**Step 4: Run tests**

Run: `cargo test -p modern-xlsx-core`
Expected: All tests pass

**Step 5: Commit**

```bash
git add crates/modern-xlsx-core/src/
git commit -m "perf: to_xml() returns Vec<u8> directly, skip String intermediary"
```

---

### Task 10: SST resolution in TypeScript read path

**Files:**
- Modify: `packages/modern-xlsx/src/workbook.ts`
- Modify: `packages/modern-xlsx/src/index.ts` (readBuffer function)
- Modify: `packages/modern-xlsx/__tests__/round-trip.test.ts`

**Step 1: Update readBuffer to auto-resolve SST indices**

In `readBuffer()`, after receiving WorkbookData from WASM:
```typescript
function resolveSharedStrings(data: WorkbookData): void {
  const sst = data.sharedStrings?.strings;
  if (!sst) return;
  for (const sheet of data.sheets) {
    for (const row of sheet.worksheet.rows) {
      for (const cell of row.cells) {
        if (cell.cellType === 'sharedString' && cell.value != null) {
          const index = Number.parseInt(cell.value, 10);
          cell.value = sst[index] ?? cell.value;
          // Keep cellType as sharedString for round-trip fidelity
        }
      }
    }
  }
}
```

**Step 2: Update Cell.value getter**

Remove any special SST resolution logic from the getter since values are now pre-resolved.

**Step 3: Update tests**

The round-trip string test currently checks SST indices directly. Update to check resolved string values.

**Step 4: Run tests**

Run: `cd packages/modern-xlsx && pnpm test`
Expected: All pass

**Step 5: Commit**

```bash
git add packages/modern-xlsx/src/ packages/modern-xlsx/__tests__/
git commit -m "feat: auto-resolve SST indices in readBuffer for ergonomic string access"
```

---

### Task 11: Unify reader/writer WorkbookData

**Files:**
- Modify: `crates/modern-xlsx-core/src/reader.rs`
- Modify: `crates/modern-xlsx-core/src/writer.rs`
- Modify: `crates/modern-xlsx-core/src/lib.rs`
- Modify: `packages/modern-xlsx/src/types.ts`

**Step 1: Create single canonical WorkbookData in lib.rs or a new types.rs**

```rust
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct WorkbookData {
    pub sheets: Vec<SheetData>,
    pub date_system: DateSystem,
    pub styles: Styles,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub shared_strings: Option<SharedStringTable>,
}
```

**Step 2: Update reader.rs to use canonical type**

Reader already returns this shape. Ensure it uses the canonical type.

**Step 3: Update writer.rs to accept canonical type**

Writer accepts `WorkbookData` and ignores `shared_strings` (builds SST internally). The write function signature stays the same.

**Step 4: Remove any duplicate type definitions**

Ensure there's only one `WorkbookData` definition in the codebase.

**Step 5: Verify TypeScript types.ts matches**

Ensure `WorkbookData` in types.ts has `sharedStrings?: SharedStringsData | null`.

**Step 6: Run all tests**

Run: `cargo test --workspace` and `cd packages/modern-xlsx && pnpm test`
Expected: All pass

**Step 7: Commit**

```bash
git add crates/ packages/
git commit -m "refactor: unify reader/writer WorkbookData into single canonical type"
```

---

### Task 12: serde_json golden file tests

**Files:**
- Create: `crates/modern-xlsx-core/tests/golden/` directory
- Create: `crates/modern-xlsx-core/tests/golden_tests.rs`
- Create: golden JSON snapshot files

**Step 1: Create golden test infrastructure**

```rust
use modern_xlsx_core::reader::read_xlsx;
use pretty_assertions::assert_eq;
use serde_json;

#[test]
fn golden_simple_workbook() {
    let xlsx_bytes = include_bytes!("fixtures/simple.xlsx");
    let data = read_xlsx(xlsx_bytes).unwrap();
    let json = serde_json::to_string_pretty(&data).unwrap();

    // First run: write the golden file
    // Subsequent runs: compare against it
    let expected = include_str!("golden/simple.json");
    assert_eq!(json, expected);
}
```

**Step 2: Create test fixtures**

Generate minimal .xlsx files using the writer, save as fixtures. Create corresponding JSON golden files.

**Step 3: Run tests**

Run: `cargo test -p modern-xlsx-core`
Expected: All pass

**Step 4: Commit**

```bash
git add crates/modern-xlsx-core/tests/
git commit -m "test: add serde_json golden file tests for reader/writer"
```

---

### Task 13: Formula preservation (read + write)

**Files:**
- Modify: `crates/modern-xlsx-core/src/ooxml/worksheet.rs` (parser + writer)
- Modify: `packages/modern-xlsx/src/types.ts`
- Modify: `packages/modern-xlsx/__tests__/round-trip.test.ts`

**Step 1: Verify formula parsing in worksheet.rs**

Check that the `<f>` element is already parsed into `Cell.formula`. The parser's `InCellFormula` state should handle this.

**Step 2: Verify formula writing in worksheet.rs to_xml()**

Ensure the writer emits `<f>FORMULA</f>` when `cell.formula.is_some()`.

**Step 3: Add formula round-trip test in TypeScript**

```typescript
it('roundtrips formula cells', async () => {
  const wb = new Workbook();
  const ws = wb.addSheet('Formulas');
  ws.cell('A1').value = 10;
  ws.cell('A2').value = 20;
  ws.cell('A3').formula = 'SUM(A1:A2)';

  const buffer = await wb.toBuffer();
  const wb2 = await readBuffer(buffer);
  const ws2 = wb2.getSheet('Formulas');
  expect(ws2?.cell('A3').formula).toBe('SUM(A1:A2)');
});
```

**Step 4: Run tests**

Run: `cargo test --workspace` and `cd packages/modern-xlsx && pnpm test`
Expected: All pass

**Step 5: Commit**

```bash
git add crates/ packages/
git commit -m "feat: formula cell round-trip preservation"
```

---

### Task 14: Named ranges in TypeScript API

**Files:**
- Modify: `crates/modern-xlsx-core/src/ooxml/workbook.rs` (parse/write definedNames)
- Modify: `packages/modern-xlsx/src/types.ts`
- Modify: `packages/modern-xlsx/src/workbook.ts`
- Modify: `packages/modern-xlsx/__tests__/round-trip.test.ts`

**Step 1: Add DefinedName to Rust WorkbookXml**

```rust
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct DefinedName {
    pub name: String,
    pub value: String,
    pub local_sheet_id: Option<u32>,
}
```

**Step 2: Parse `<definedNames>` in workbook.rs**

Add state for parsing `<definedName>` elements in the workbook XML parser.

**Step 3: Write `<definedNames>` in workbook.rs to_xml()**

Serialize defined names back to XML.

**Step 4: Add TypeScript types and API**

```typescript
interface DefinedNameData {
  name: string;
  value: string;
  localSheetId: number | null;
}

// In Workbook class:
get namedRanges(): readonly DefinedNameData[] { ... }
addNamedRange(name: string, value: string, localSheetId?: number): void { ... }
getNamedRange(name: string): DefinedNameData | undefined { ... }
```

**Step 5: Add tests**

Test round-trip of named ranges.

**Step 6: Run tests**

Run: `cargo test --workspace` and `cd packages/modern-xlsx && pnpm test`

**Step 7: Commit**

```bash
git add crates/ packages/
git commit -m "feat: named ranges API (definedNames read/write)"
```

---

### Task 15: Style builder API

**Files:**
- Create: `packages/modern-xlsx/src/style-builder.ts`
- Modify: `packages/modern-xlsx/src/index.ts`
- Modify: `packages/modern-xlsx/src/workbook.ts`
- Modify: `packages/modern-xlsx/__tests__/round-trip.test.ts`

**Step 1: Create StyleBuilder class**

```typescript
export class StyleBuilder {
  private fontData: Partial<FontData> = {};
  private fillData: Partial<FillData> = {};
  private borderData: Partial<BorderData> = {};
  private numFmtCode: string | null = null;

  font(opts: Partial<FontData>): this { ... }
  fill(opts: Partial<FillData>): this { ... }
  border(opts: Partial<BorderData>): this { ... }
  numberFormat(code: string): this { ... }
  build(styles: StylesData): number { /* returns styleIndex */ ... }
}
```

**Step 2: Add createStyle() to Workbook**

```typescript
createStyle(): StyleBuilder {
  return new StyleBuilder();
}
```

**Step 3: Add tests for style builder**

**Step 4: Run tests**

**Step 5: Commit**

```bash
git add packages/modern-xlsx/src/ packages/modern-xlsx/__tests__/
git commit -m "feat: fluent StyleBuilder API for creating cell styles"
```

---

### Task 16: Rich text / formatted runs

**Files:**
- Modify: `crates/modern-xlsx-core/src/ooxml/shared_strings.rs`
- Modify: `packages/modern-xlsx/src/types.ts`
- Modify: `packages/modern-xlsx/src/workbook.ts`

**Step 1: Add RichTextRun to Rust SharedStringTable**

```rust
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct RichTextRun {
    pub text: String,
    pub bold: Option<bool>,
    pub italic: Option<bool>,
    pub font_name: Option<String>,
    pub font_size: Option<f64>,
    pub color: Option<String>,
}
```

Update `SharedStringEntry` to be an enum: `Plain(String)` or `Rich(Vec<RichTextRun>)`.

**Step 2: Parse `<r>` elements in shared_strings parser**

Handle `<si><r><rPr><b/><sz val="14"/></rPr><t>bold text</t></r></si>`.

**Step 3: Write rich text runs in to_xml()**

**Step 4: Add TypeScript types**

**Step 5: Test round-trip**

**Step 6: Commit**

```bash
git add crates/ packages/
git commit -m "feat: rich text / formatted runs in shared strings"
```

---

### Task 17: Data validation rules

**Files:**
- Modify: `crates/modern-xlsx-core/src/ooxml/worksheet.rs`
- Modify: `packages/modern-xlsx/src/types.ts`
- Modify: `packages/modern-xlsx/src/workbook.ts`

**Step 1: Add DataValidation to Rust WorksheetXml**

```rust
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct DataValidation {
    pub sqref: String,
    pub validation_type: Option<String>,  // list, whole, decimal, date, textLength, custom
    pub operator: Option<String>,
    pub formula1: Option<String>,
    pub formula2: Option<String>,
    pub allow_blank: Option<bool>,
    pub show_error_message: Option<bool>,
    pub error_title: Option<String>,
    pub error_message: Option<String>,
}
```

**Step 2: Parse `<dataValidation>` in worksheet parser**

**Step 3: Write `<dataValidations>` in to_xml()**

**Step 4: TypeScript types and Worksheet API**

```typescript
addValidation(ref: string, rule: DataValidationData): void { ... }
get validations(): readonly DataValidationData[] { ... }
```

**Step 5: Test round-trip**

**Step 6: Commit**

```bash
git add crates/ packages/
git commit -m "feat: data validation rules (read/write/API)"
```

---

### Task 18: Conditional formatting

**Files:**
- Modify: `crates/modern-xlsx-core/src/ooxml/worksheet.rs`
- Modify: `packages/modern-xlsx/src/types.ts`

**Step 1: Add ConditionalFormatting to WorksheetXml**

```rust
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ConditionalFormatting {
    pub sqref: String,
    pub rules: Vec<ConditionalFormattingRule>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ConditionalFormattingRule {
    pub rule_type: String,
    pub priority: u32,
    pub operator: Option<String>,
    pub formula: Option<String>,
    pub dxf_id: Option<u32>,
}
```

**Step 2: Parse `<conditionalFormatting>` elements**

**Step 3: Write back to XML**

**Step 4: TypeScript types**

**Step 5: Test round-trip**

**Step 6: Commit**

```bash
git add crates/ packages/
git commit -m "feat: conditional formatting rules (read/write)"
```

---

### Task 19: Images/charts read-preserve

**Files:**
- Modify: `crates/modern-xlsx-core/src/reader.rs`
- Modify: `crates/modern-xlsx-core/src/writer.rs`
- Modify: `crates/modern-xlsx-core/src/ooxml/relationships.rs`

**Step 1: Preserve unknown ZIP entries**

In reader.rs, store all ZIP entries that aren't explicitly parsed (drawings, media, etc.) as opaque byte blobs.

**Step 2: Write them back**

In writer.rs, include preserved entries in the output ZIP.

**Step 3: Preserve drawing relationships**

Read and write `xl/drawings/drawing*.xml` and `xl/worksheets/_rels/sheet*.xml.rels`.

**Step 4: Test with image-containing XLSX**

**Step 5: Commit**

```bash
git add crates/
git commit -m "feat: preserve images/charts/drawings through round-trip"
```

---

### Task 20: Streaming read/write for large files

**Files:**
- Create: `crates/modern-xlsx-core/src/streaming.rs`
- Modify: `crates/modern-xlsx-core/src/lib.rs`

**Step 1: Streaming reader**

```rust
pub struct StreamingReader<R: Read + Seek> {
    zip: ZipArchive<R>,
    shared_strings: SharedStringTable,
    styles: Styles,
}

impl StreamingReader {
    pub fn open(reader: R) -> Result<Self, modern-xlsxError> { ... }
    pub fn sheet_names(&self) -> Vec<String> { ... }
    pub fn read_sheet(&mut self, name: &str) -> impl Iterator<Item = Row> { ... }
}
```

**Step 2: Streaming writer**

```rust
pub struct StreamingWriter<W: Write + Seek> {
    zip: ZipWriter<W>,
    sst: SharedStringTableBuilder,
}

impl StreamingWriter {
    pub fn new(writer: W) -> Self { ... }
    pub fn start_sheet(&mut self, name: &str) -> SheetWriter { ... }
    pub fn finish(self) -> Result<(), modern-xlsxError> { ... }
}
```

**Step 3: Tests with large row counts**

**Step 4: Commit**

```bash
git add crates/modern-xlsx-core/src/
git commit -m "feat: streaming read/write API for large files"
```

---

### Task 21: Parallel sheet parsing (rayon feature flag)

**Files:**
- Modify: `crates/modern-xlsx-core/Cargo.toml`
- Modify: `crates/modern-xlsx-core/src/reader.rs`

**Step 1: Add rayon as optional dependency**

```toml
[dependencies]
rayon = { version = "1.10", optional = true }

[features]
parallel = ["rayon"]
```

**Step 2: Conditional parallel parsing**

```rust
#[cfg(feature = "parallel")]
{
    use rayon::prelude::*;
    sheets = sheet_info.par_iter()
        .map(|info| parse_worksheet(info, &entries))
        .collect::<Result<Vec<_>>>()?;
}

#[cfg(not(feature = "parallel"))]
{
    sheets = sheet_info.iter()
        .map(|info| parse_worksheet(info, &entries))
        .collect::<Result<Vec<_>>>()?;
}
```

**Step 3: Test with and without feature**

Run: `cargo test -p modern-xlsx-core` and `cargo test -p modern-xlsx-core --features parallel`

**Step 4: Commit**

```bash
git add crates/modern-xlsx-core/
git commit -m "feat: optional parallel sheet parsing with rayon feature flag"
```

---

### Task 22: Browser testing with Playwright

**Files:**
- Create: `packages/modern-xlsx/__tests__/browser.test.ts`
- Modify: `packages/modern-xlsx/vitest.config.ts`

**Step 1: Add browser test configuration**

Create separate vitest config or use workspace projects for browser vs node tests.

**Step 2: Write browser-specific tests**

Test WASM init in browser, Blob creation, round-trip in browser environment.

**Step 3: Run tests**

**Step 4: Commit**

```bash
git add packages/modern-xlsx/
git commit -m "test: add browser tests with Vitest browser mode"
```

---

### Task 23: npm publish pipeline

**Files:**
- Create: `.github/workflows/publish.yml`
- Modify: `packages/modern-xlsx/package.json`

**Step 1: Create GitHub Actions workflow**

Workflow that:
1. Installs Rust + wasm-pack
2. Builds WASM
3. Runs Rust tests
4. Installs Node deps
5. Runs TypeScript tests
6. Publishes to npm on tag push

**Step 2: Add prepublishOnly script**

**Step 3: Commit**

```bash
git add .github/ packages/modern-xlsx/package.json
git commit -m "ci: add npm publish pipeline with GitHub Actions"
```

---

### Task 24: tsdown external config fix

**Files:**
- Modify: `packages/modern-xlsx/tsdown.config.ts` (or equivalent build config)

**Step 1: Configure externals**

Ensure WASM files are externalized and not inlined by the bundler.

**Step 2: Verify build output**

**Step 3: Commit**

```bash
git add packages/modern-xlsx/
git commit -m "fix: tsdown external config for WASM imports"
```
