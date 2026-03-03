# Beat SheetJS — Feature Parity Design

**Goal:** Close all 12 categories where SheetJS currently beats or ties modern-xlsx, across 3 phased releases.

**Current scorecard:** modern-xlsx wins 20, SheetJS wins 8, Tie 4 (of 32 categories)
**Target scorecard:** modern-xlsx wins 28, SheetJS wins 2 (format count only), Tie 2

---

## Architecture Principles

- All new format parsers (CSV, ODS, XLSB) live in Rust core — same performance guarantees as XLSX
- Auto-format detection at the WASM `read()` boundary — TS API stays unchanged
- New TS utilities are pure functions — no WASM changes unless data needs to cross the boundary
- Backward-compatible: zero breaking changes across all 3 releases

---

## v0.5.0 — "Parity Sprint"

**Scope:** Close all easy gaps. Pure TS additions + minor Rust field additions. ~680 lines.

**Categories flipped:** 7 (cell ref utils, doc properties, sheet conversions, number formatting, worksheet ops, cell ops, formulas)

### 1. Cell Reference Utilities

**Files:** `packages/modern-xlsx/src/cell-ref.ts`, `packages/modern-xlsx/src/index.ts`

Add 3 functions:

| Function | Signature | Behavior |
|----------|-----------|----------|
| `encodeRow` | `(row: number) => string` | 0-based input → 1-based string. `encodeRow(0)` → `"1"` |
| `decodeRow` | `(rowStr: string) => number` | 1-based string → 0-based. `decodeRow("1")` → `0` |
| `splitCellRef` | `(ref: string) => { col: string, row: string, absCol: boolean, absRow: boolean }` | `splitCellRef("$A$1")` → `{ col: "A", row: "1", absCol: true, absRow: true }` |

Export all 3 from `index.ts`.

### 2. Document Properties

**Files:** `crates/modern-xlsx-core/src/ooxml/doc_props.rs`, `packages/modern-xlsx/src/types.ts`

Add fields to `DocProperties` (Rust) and `DocPropertiesData` (TS):

| Field | XML Source | Type |
|-------|-----------|------|
| `appVersion` | `<AppVersion>` in app.xml | `Option<String>` / `string \| null` |
| `hyperlinkBase` | `<HyperlinkBase>` in app.xml | `Option<String>` / `string \| null` |
| `revision` | `<cp:revision>` in core.xml | `Option<String>` / `string \| null` |

Update parser in `parse_app_properties()` and `parse_core_properties()`. Update writer to emit these elements.

### 3. Sheet Conversion Utilities

**Files:** `packages/modern-xlsx/src/utils.ts`, `packages/modern-xlsx/src/index.ts`

Add 2 functions:

| Function | Signature | Behavior |
|----------|-----------|----------|
| `sheetToTxt` | `(ws: Worksheet, opts?: SheetToTxtOptions) => string` | Tab-separated output. Reuse `sheetToCsv` internals with `\t` delimiter. |
| `sheetToFormulae` | `(ws: Worksheet) => string[]` | Scan all cells, return `["A1=SUM(B1:B10)", "B2=42"]`. Cells without formulas show `ref=value`. |

`SheetToTxtOptions`: same as `SheetToCsvOptions` but default delimiter is `\t`.

### 4. Number Formatting

**Files:** `packages/modern-xlsx/src/format-cell.ts`, `packages/modern-xlsx/src/index.ts`

Enhance `formatCell()` with:

| Feature | Description |
|---------|-------------|
| Conditional sections | Parse `[>100]#,##0;[<=100]0.00`. Extract bracket condition `[op value]`, evaluate against cell value, select matching section. |
| Bracket color codes | Parse `[Red]`, `[Blue]`, `[Green]`, `[Yellow]`, `[Magenta]`, `[Cyan]`, `[White]`, `[Black]`, `[Color1]`–`[Color56]`. Strip from format string, return color name in `FormatCellResult.color`. |
| `loadFormat` | `(fmt: string, id: number) => void` — register custom format code at runtime in a module-level `Map<number, string>`. |
| `loadFormatTable` | `(table: Record<number, string>) => void` — bulk register. |

Change `formatCell` return type from `string` to `FormatCellResult`:

```typescript
interface FormatCellResult {
  text: string;
  color?: string;  // e.g. "Red", "Color3"
}
```

Keep backward compat: `formatCell()` still returns `string` by default. Add `formatCellRich()` that returns the full result.

### 5. Worksheet Operations

**Files:** `packages/modern-xlsx/src/workbook.ts`, `packages/modern-xlsx/src/types.ts`, `crates/modern-xlsx-core/src/ooxml/worksheet.rs`

| Feature | API | Implementation |
|---------|-----|----------------|
| Used range | `ws.usedRange: string \| null` getter | Compute from `ws.rows` — find min/max row and col across all cells. Return as `"A1:Z100"` or `null` if empty. |
| Sheet tab color | `ws.tabColor: string \| null` getter/setter | Add `tabColor: Option<String>` to Rust `SheetData`. Parse from `<sheetPr><tabColor rgb="..."/>`. Write back. Add to TS `SheetData` type. |

### 6. Cell Operations

**Files:** `packages/modern-xlsx/src/workbook.ts`, `packages/modern-xlsx/src/types.ts`, `crates/modern-xlsx-core/src/ooxml/worksheet.rs`

| Feature | API | Implementation |
|---------|-----|----------------|
| Per-cell number format | `cell.numberFormat: string \| null` getter/setter | Getter: look up `styleIndex` → `CellXf.numFmtId` → `NumFmt.formatCode`. Setter: find or create a `CellXf` with that format code, set `styleIndex`. |
| Stub cell type | Add `'stub'` to `CellType` union | Rust: add `Stub` variant to `CellType` enum. TS: add to union. Represents explicitly empty cells (SheetJS type "z"). |
| Date value getter | `cell.dateValue: Date \| null` getter | If cell value is a number and the resolved format is a date format (`isDateFormatCode`), return `serialToDate(value)`. Otherwise `null`. |

### 7. Formulas

**Files:** `packages/modern-xlsx/src/types.ts`, `crates/modern-xlsx-core/src/ooxml/worksheet.rs`

| Feature | API | Implementation |
|---------|-----|----------------|
| Dynamic array flag | `CellData.dynamicArray?: boolean` | Rust: add `dynamic_array: Option<bool>` to `Cell`. Parse from `<f t="array" ... cm="1">` (the `cm` attribute indicates dynamic array). Write back. TS: add to `CellData`. |

---

## v0.6.0 — "CSV + Streaming"

**Scope:** CSV read/write in Rust, streaming TS API, sync wrapper. ~1200 lines.

**Categories flipped:** 2 (I/O operations, streaming)

### 8. CSV Reader (Rust)

**File:** `crates/modern-xlsx-core/src/csv.rs`

```rust
pub struct CsvOptions {
    pub delimiter: u8,        // default b','
    pub quote: u8,            // default b'"'
    pub has_header: bool,     // default true
    pub encoding: CsvEncoding, // Utf8 only for now
}

pub fn read_csv(data: &[u8], opts: &CsvOptions) -> Result<WorkbookData>
```

- RFC 4180 compliant quote handling
- Auto-detect types per cell: try parse as `f64`, then `bool` ("true"/"false"), else string
- Single-sheet output named "Sheet1"
- Set `dimension` from row/col count
- Use `memchr` for fast delimiter scanning (already available in quick-xml's deps)

### 9. CSV Writer (Rust)

**File:** `crates/modern-xlsx-core/src/csv.rs`

```rust
pub fn write_csv(workbook: &WorkbookData, opts: &CsvOptions) -> Result<Vec<u8>>
```

- Write first sheet (or named sheet via option)
- Auto-quote cells containing delimiter, newline, or quote char
- Resolve SST references to string values
- Numbers written with full precision

### 10. WASM Bridge

**File:** `crates/modern-xlsx-wasm/src/lib.rs`

Add:
```rust
#[wasm_bindgen(js_name = "readCsv")]
pub fn read_csv(data: &[u8], opts_json: &str) -> Result<String, JsError>

#[wasm_bindgen(js_name = "writeCsv")]
pub fn write_csv(json: &str, opts_json: &str) -> Result<Vec<u8>, JsError>
```

### 11. TypeScript CSV API

**File:** `packages/modern-xlsx/src/index.ts`

```typescript
export async function readCsv(data: Uint8Array, opts?: CsvReadOptions): Promise<Workbook>
export async function writeCsv(wb: Workbook, opts?: CsvWriteOptions): Promise<Uint8Array>
export async function readCsvFile(path: string, opts?: CsvReadOptions): Promise<Workbook>
```

### 12. Streaming API

**File:** `packages/modern-xlsx/src/streaming.ts`

| Function | Signature | Implementation |
|----------|-----------|----------------|
| `streamToJson` | `(data: Uint8Array, opts?) => AsyncGenerator<Record<string, unknown>>` | WASM reads sheet, TS yields row-by-row from parsed rows |
| `streamToCsv` | `(data: Uint8Array, opts?) => AsyncGenerator<string>` | Same, but yields CSV lines |
| Row iterator | `ws.rowIterator()` → `IterableIterator<RowData>` | Sync generator over `ws.rows` |

For v0.6.0, "streaming" means row-by-row iteration over already-parsed data. True chunk-based streaming (reading ZIP entries one at a time without loading full file) is a future optimization.

### 13. Sync API Wrapper

**File:** `packages/modern-xlsx/src/sync.ts`

```typescript
export function readBufferSync(data: Uint8Array): Workbook
export function writeBufferSync(wb: Workbook): Uint8Array
export function readFileSync(path: string): Workbook
```

Requires `initWasmSync(module)` to have been called first. Throws if WASM not initialized.

Guard: check `typeof globalThis.process !== 'undefined'` — only available in Node.js/Bun/Deno.

---

## v0.7.0 — "ODS + XLSB"

**Scope:** ODS read/write + XLSB read in Rust. ~2700 lines.

**Categories improved:** read formats (1→5), write formats (1→4)

### 14. Format Detection

**File:** `crates/modern-xlsx-core/src/lib.rs`

```rust
pub enum FileFormat { Xlsx, Xlsb, Ods, Csv, Unknown }

pub fn detect_format(data: &[u8]) -> FileFormat {
    if data.starts_with(b"PK") {
        // ZIP-based: check internal files
        // [Content_Types].xml → XLSX
        // xl/*.bin → XLSB
        // content.xml + mimetype → ODS
    } else {
        // Try CSV heuristics (printable ASCII, delimiter patterns)
    }
}
```

Update `read()` WASM export to auto-detect and dispatch.

### 15. ODS Reader (Rust)

**File:** `crates/modern-xlsx-core/src/ods.rs`

ODS is ZIP with XML (like XLSX but OpenDocument schema):

| XLSX Part | ODS Equivalent |
|-----------|---------------|
| `xl/workbook.xml` | `content.xml` `<office:spreadsheet>` |
| `xl/worksheets/sheet1.xml` | `<table:table>` in content.xml |
| `xl/styles.xml` | `styles.xml` + `<office:automatic-styles>` |
| `xl/sharedStrings.xml` | N/A (strings inline) |

Parse `content.xml` with quick-xml SAX:
- `<table:table table:name="...">` → new sheet
- `<table:table-row>` → new row (handle `table:number-rows-repeated`)
- `<table:table-cell office:value-type="float" office:value="42">` → number cell
- `<table:table-cell office:value-type="string"><text:p>hello</text:p>` → string cell
- `<table:table-cell table:number-columns-repeated="5"/>` → repeated empty cells

Map to existing `WorkbookData` — same output as XLSX reader.

### 16. ODS Writer (Rust)

**File:** `crates/modern-xlsx-core/src/ods.rs`

Generate:
- `mimetype` (uncompressed, must be first ZIP entry): `application/vnd.oasis.opendocument.spreadsheet`
- `META-INF/manifest.xml` — file manifest
- `content.xml` — sheets, rows, cells
- `styles.xml` — font/fill/border/alignment mapped to ODS style attributes
- `meta.xml` — document properties

### 17. XLSB Reader (Rust)

**File:** `crates/modern-xlsx-core/src/xlsb.rs`

XLSB is ZIP with binary record streams instead of XML. Each `.bin` file contains:

```
[record_type: variable-length int] [record_size: variable-length int] [payload: bytes]
```

Key records to parse:

| Record | Type ID | Payload |
|--------|---------|---------|
| `BrtBundleSh` | 0x009C | Sheet name + state + rId |
| `BrtRowHdr` | 0x0000 | Row index + height + flags |
| `BrtCellIsst` | 0x0007 | Col + SST index + style |
| `BrtCellSt` | 0x0006 | Col + inline string + style |
| `BrtCellReal` | 0x0005 | Col + f64 value + style |
| `BrtCellRk` | 0x0002 | Col + RK-encoded number + style |
| `BrtCellBool` | 0x0004 | Col + bool + style |
| `BrtCellBlank` | 0x0001 | Col + style |
| `BrtCellError` | 0x0003 | Col + error code + style |
| `BrtFmlaString` | 0x0008 | Col + formula + string result |
| `BrtFmlaNum` | 0x0009 | Col + formula + f64 result |
| `BrtSSTItem` | 0x0013 | SST string entry |
| `BrtBeginSst` | 0x009F | SST count + unique count |
| `BrtMergeCell` | 0x00B0 | Merge range |

Read-only — no XLSB writer. Modern workflows should write XLSX.

### 18. WASM + TS Updates

Update `read()` to auto-detect. No TS API changes — `readBuffer()` handles all formats transparently.

Add `writeOds()` export:

```typescript
export async function writeOds(wb: Workbook): Promise<Uint8Array>
```

---

## Scorecard Progression

| Category | v0.4.0 | v0.5.0 | v0.6.0 | v0.7.0 |
|----------|--------|--------|--------|--------|
| **modern-xlsx wins** | 20 | 27 | 29 | **28** |
| **SheetJS wins** | 8 | 1 | 1 | **2** |
| **Ties** | 4 | 4 | 2 | **2** |

After v0.5.0: SheetJS only wins on read/write format count.
After v0.6.0: CSV added, streaming added, I/O flipped.
After v0.7.0: ODS + XLSB added. SheetJS wins only on legacy format count (BIFF, SYLK, DIF, DBF, etc.).

## Remaining SheetJS-only Features (not planned)

These are deliberately excluded — declining/niche formats not worth the investment:

- XLS read/write (BIFF2/3/4/5/8) — legacy, declining usage
- XLML (SpreadsheetML 2003) — superseded by XLSX
- SYLK, DIF, PRN, ETH, DBF — niche interchange formats
- WK1/WK3/WKS (Lotus 1-2-3) — extinct
- QPW (Quattro Pro) — extinct
- Numbers (Apple) — requires reverse-engineered protobuf parser
- DOM table parsing (`table_to_sheet`) — browser-only, trivial for users to implement
- CFB container parsing — only needed for XLS/DOC

---

## Test Strategy

Each release adds tests to `feature-comparison.test.ts` that verify the gap is closed:

- v0.5.0: Update matrix counts, add tests for new utilities
- v0.6.0: Add CSV roundtrip tests, streaming tests, sync API tests
- v0.7.0: Add ODS roundtrip tests, XLSB read tests with golden files

Rust tests: co-located `#[cfg(test)]` modules for each new parser.
