<div align="center">

<picture>
  <source media="(prefers-color-scheme: dark)" srcset="https://raw.githubusercontent.com/ABCrimson/modern-xlsx/v1.3.0/docs/assets/readme/hero-dark.svg">
  <img alt="modern-xlsx" src="https://raw.githubusercontent.com/ABCrimson/modern-xlsx/v1.3.0/docs/assets/readme/hero-light.svg" width="100%">
</picture>

**Read, write, style and encrypt Excel `.xlsx` files in JavaScript and TypeScript, with a Rust core compiled to WebAssembly.**

[![npm](https://img.shields.io/npm/v/modern-xlsx?style=flat-square&label=npm&logo=npm&logoColor=white&labelColor=24292f&color=c8102e)](https://www.npmjs.com/package/modern-xlsx) [![CI](https://img.shields.io/github/actions/workflow/status/ABCrimson/modern-xlsx/ci.yml?branch=master&style=flat-square&label=CI&logo=githubactions&logoColor=white&labelColor=24292f)](https://github.com/ABCrimson/modern-xlsx/actions/workflows/ci.yml) [![License](https://img.shields.io/npm/l/modern-xlsx?style=flat-square&labelColor=24292f&color=59636e)](https://github.com/ABCrimson/modern-xlsx/blob/v1.3.0/packages/modern-xlsx/LICENSE) ![Types included](https://img.shields.io/badge/types-included-59636e?style=flat-square&logo=typescript&logoColor=white&labelColor=24292f) ![Zero dependencies](https://img.shields.io/badge/dependencies-0-59636e?style=flat-square&labelColor=24292f)

[API&nbsp;reference](https://abcrimson.github.io/modern-xlsx/api/)&emsp; [Playground](https://abcrimson.github.io/modern-xlsx/playground.html)&emsp; [Recipes](https://github.com/ABCrimson/modern-xlsx/blob/v1.3.0/docs/examples.md)&emsp; [Comparison](https://github.com/ABCrimson/modern-xlsx/blob/v1.3.0/docs/FEATURE-COMPARISON.md)&emsp; [SheetJS&nbsp;migration](https://github.com/ABCrimson/modern-xlsx/blob/v1.3.0/docs/migration-from-sheetjs.md)&emsp; [ExcelJS&nbsp;migration](https://github.com/ABCrimson/modern-xlsx/blob/v1.3.0/docs/migration-from-exceljs.md)&emsp; [Changelog](https://github.com/ABCrimson/modern-xlsx/blob/v1.3.0/packages/modern-xlsx/CHANGELOG.md)

</div>

modern-xlsx reads and writes `.xlsx` workbooks with cell styles, number formats, data validation, conditional formatting, frozen panes, hyperlinks, comments, tables, charts, pivot tables, images and sheet protection, encrypts and decrypts password-protected files, and ships a formula engine, a streaming writer and a CLI.

**Contents:** [Install](#install) · [Quick start](#quick-start) · [Common tasks](#common-tasks) · [Runtimes](#runtimes-and-entry-points) · [Features](#features) · [Performance](#performance) · [API overview](#api-overview) · [Limits](#limits) · [Security](#security) · [TypeScript](#typescript)

## Install

```bash
npm install modern-xlsx
```

### Requirements

- A runtime with WebAssembly: Node.js 24 or newer, Bun, Deno or a current browser.
- TypeScript 5.7 or newer to type-check against the bundled declarations. A project that type-checks library declarations (`skipLibCheck: false`) also needs the `esnext` and `dom` libs: see [TypeScript setup](#typescript-setup).

## Quick start

<!-- quick-start:begin. Keep identical to the Quick Start block in the root README.md; CI checks it with node scripts/check-readme-sync.mjs. -->
```typescript
import * as xlsx from 'modern-xlsx';

await xlsx.initWasm();

// Build a workbook
const wb = new xlsx.Workbook();
const ws = wb.addSheet('Scores');
ws.cell('A1').value = 'Name';
ws.cell('B1').value = 'Score';
ws.cell('A2').value = 'Alice';
ws.cell('B2').value = 95;

// Style the header row
const header = wb.createStyle();
header.font({ bold: true });
const bold = header.build(wb.styles);
ws.cell('A1').styleIndex = bold;
ws.cell('B1').styleIndex = bold;

// Write it and read it back (Node.js,
// Bun, Deno; in a browser, use
// writeBlob(wb) and readBuffer(bytes))
const file = 'scores.xlsx';
await wb.toFile(file);
const back = await xlsx.readFile(file);
const sheet = back.getSheet('Scores');
console.log(sheet?.cell('B2').value); // 95
```
<!-- quick-start:end -->

## Why modern-xlsx

- **One free package for the whole workbook.** Cell styles, conditional formatting, charts, images, pivot tables and password encryption are all in the MIT-licensed package. SheetJS sells them as SheetJS Pro.
- **Fast reads, small files.** In the [benchmark](#performance), reading 100,000 rows of 10 columns took 485 ms against SheetJS's 1,591 ms, and a 10,000-row file came out 2.6 times smaller. Cell-by-cell writes are slower than SheetJS's.
- **Runs where your code runs.** CI tests Node.js 24 and 26, Bun, Deno, Chromium and a Cloudflare Worker. TypeScript types are included, and there are no runtime dependencies.

## Common tasks

Call `await initWasm()` once before these examples, as in the Quick start. The save example uses the `wb` built in [Style a header row](#style-a-header-row).

### Style a header row

```typescript
import { initWasm, Workbook } from 'modern-xlsx';

await initWasm();

const wb = new Workbook();
const ws = wb.addSheet('Sales');

ws.cell('A1').value = 'Product';
ws.cell('B1').value = 'Revenue';
ws.cell('A2').value = 'Widget';
ws.cell('B2').value = 9999.99;

// Style headers
const header = wb
  .createStyle()
  .font({ bold: true, size: 14, color: '1F4E79' })
  .fill({ pattern: 'solid', fgColor: 'D6E4F0' })
  .alignment({ horizontal: 'center' })
  .border({ bottom: { style: 'medium', color: '1F4E79' } })
  .build(wb.styles);

ws.cell('A1').styleIndex = header;
ws.cell('B1').styleIndex = header;

await wb.toFile('sales.xlsx');
```

### Read a file

```typescript
import { initWasm, readBuffer, readFile } from 'modern-xlsx';

await initWasm();

const fromDisk = await readFile('report.xlsx'); // Node.js, Bun, Deno
const bytes = await fromDisk.toBuffer(); // a Uint8Array of a .xlsx file
const fromBytes = await readBuffer(bytes); // any environment

console.log(fromDisk.getSheet('Sheet1')?.cell('A1').value);
console.log(fromBytes.sheetNames);
```

### Save a file or a browser download

```typescript
import { writeBlob } from 'modern-xlsx';

await wb.toFile('output.xlsx'); // Node.js, Bun, Deno
const bytes = await wb.toBuffer(); // Uint8Array, any environment
const blob = writeBlob(wb); // Blob, for a browser download
```

`readFile` and `Workbook#toFile` need file access, which comes from the package's `node` import condition: they work in Node.js, Bun and Deno. In browsers, in workers and in the IIFE bundle they reject with a `ModernXlsxError` whose code is `IO_ERROR`; use `readBuffer`, `toBuffer` or `writeBlob` there. The browser bundles import no `node:` module ([#24](https://github.com/ABCrimson/modern-xlsx/issues/24)): after a build, `grep -c 'node:' dist/modern-xlsx.min.js dist/runtime-browser.mjs dist/index.mjs` prints 0 for each file.

### Convert rows to and from JSON and CSV

```typescript
import {
  aoaToSheet,
  jsonToSheet,
  sheetAddAoa,
  sheetAddJson,
  sheetToCsv,
  sheetToHtml,
  sheetToJson,
} from 'modern-xlsx';

// Array-of-arrays -> sheet
const ws = aoaToSheet([
  ['Name', 'Age'],
  ['Alice', 30],
  ['Bob', 25],
]);

// JSON -> sheet
const ws2 = jsonToSheet([
  { name: 'Alice', age: 30 },
  { name: 'Bob', age: 25 },
]);

// Sheet -> JSON / CSV / HTML
// [{ Name: 'Alice', Age: 30 }, { Name: 'Bob', Age: 25 }]
const data = sheetToJson(ws);
const csv = sheetToCsv(ws); // "Name,Age\nAlice,30\nBob,25"
const html = sheetToHtml(ws); // "<table><tr><td>Name</td><td>Age</td></tr>..."

// Append to existing sheet (skipHeader: the sheet already has a header row)
sheetAddAoa(ws, [['Charlie', 35]], { origin: 'A4' });
sheetAddJson(ws, [{ Name: 'Diana', Age: 28 }], { skipHeader: true });
sheetToCsv(ws); // "Name,Age\nAlice,30\nBob,25\nCharlie,35\nDiana,28"

// Limit rows
const first5 = sheetToJson(ws, { sheetRows: 5 });
```

### Encrypt a workbook

Read and write password-protected workbooks (MS-OFFCRYPTO). The reader decrypts Agile and Standard encryption with AES-128, AES-192 or AES-256 keys; the writer uses Agile encryption with AES-256-CBC and SHA-512. In CI, msoffcrypto-tool decrypts the written output, and the decryption tests use [files encrypted by other software](https://github.com/ABCrimson/modern-xlsx/blob/v1.3.0/crates/modern-xlsx-core/tests/fixtures/msoffcrypto/README.md). `modern-xlsx/lite` has no encryption: a `password` option there throws a `ModernXlsxError` with code `INVALID_ARGUMENT`.

```typescript
import { initWasm, readFile } from 'modern-xlsx';

await initWasm();

// Read encrypted file
const wb = await readFile('protected.xlsx', { password: 'secret' });

// Write encrypted file
await wb.toFile('output.xlsx', { password: 'secret' });
```

### Write a large file row by row

```typescript
import { initWasm, StreamingXlsxWriter } from 'modern-xlsx';

await initWasm();
const writer = StreamingXlsxWriter.create();
writer.startSheet('Data');
for (let i = 0; i < 100_000; i++) {
  writer.writeRow([
    { value: String(i), cellType: 'number' },
    { value: `row_${i}`, cellType: 'sharedString' },
  ]);
}
const xlsx: Uint8Array = writer.finish();
```

Each row is deflated into an in-memory ZIP as it arrives, and `finish()` returns the bytes, so memory grows with the compressed output (plus the shared string table), not with the uncompressed sheet XML. Cell values are strings: `writeRow` throws a `ModernXlsxError` with code `INVALID_CELL_VALUE` for a `'number'` cell (the default type) whose value is not a finite number.

## Runtimes and entry points

- **`modern-xlsx`:** the full API (ESM).
- **`modern-xlsx/lite`:** the same API on the smaller lite WASM build, which has no encryption. A `password` option throws `INVALID_ARGUMENT`; an encrypted file rejects with `PASSWORD_PROTECTED` and a legacy `.xls` with `LEGACY_FORMAT`. Lite has no standalone barcode exports (`generateBarcode`, `encodeQR`, `renderBarcodePNG`, ...) and no `createXlsxWorker`. `Workbook#addBarcode` still works in both entries, because barcodes are drawn in TypeScript, not in WASM.
- **`modern-xlsx/browser`:** a classic `<script>` tag. The IIFE bundle `dist/modern-xlsx.min.js` defines the global `ModernXlsx`, which holds every runtime export of `modern-xlsx`. With no argument, `initWasm()` fetches `modern-xlsx.wasm` from beside the script. Top-level `await` needs a module script. Importing this subpath as an ES module yields no exports.
- **`modern-xlsx/worker`:** the module-worker script `dist/modern-xlsx.worker.js` that `createXlsxWorker({ workerUrl })` loads. It loads `modern-xlsx.wasm` from beside itself unless you pass `wasmUrl`.
- **`modern-xlsx/wasm`:** the raw full WASM binary for `initWasmSync`, in runtimes that import `.wasm` as a module (Cloudflare Workers).

Neither `modern-xlsx/worker` nor `modern-xlsx/wasm` is importable in plain Node: the worker script needs a worker global scope, and the `.wasm` subpath a runtime that imports `.wasm` files as modules. In Node, import `modern-xlsx` and call `initWasm()`.

### Browsers with a bundler or import map

```html
<script type="module">
  // The bare 'modern-xlsx' specifier resolves through a bundler
  // or an import map.
  import { initWasm, Workbook, writeBlob } from 'modern-xlsx';

  await initWasm();

  const wb = new Workbook();
  wb.addSheet('Sheet1').cell('A1').value = 'Hello from browser!';

  const blob = writeBlob(wb);
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = 'output.xlsx';
  a.click();
</script>
```

### Browsers with a script tag

No bundler required: a classic `<script>` tag loads the bundle, which defines the global `ModernXlsx`, and a module script uses it with top-level `await`:

```html
<script
  src="https://cdn.jsdelivr.net/npm/modern-xlsx@1.3.0/dist/modern-xlsx.min.js"
></script>
<script type="module">
  // With no argument, initWasm() fetches modern-xlsx.wasm
  // from beside modern-xlsx.min.js
  await ModernXlsx.initWasm();
  const wb = new ModernXlsx.Workbook();
  wb.addSheet('Sheet1').cell('A1').value = 'Hello!';
  const blob = ModernXlsx.writeBlob(wb);
  // trigger download...
</script>
```

Also available via unpkg: `https://unpkg.com/modern-xlsx@1.3.0/dist/modern-xlsx.min.js`

Pin the exact version, as above. For Subresource Integrity, add an `integrity="sha384-..."` attribute with `crossorigin="anonymous"`. From 1.3.0 on, each [GitHub release](https://github.com/ABCrimson/modern-xlsx/releases) lists the SHA-384 value of its `dist/modern-xlsx.min.js` under "CDN integrity", computed from the published npm tarball. To compute it from the file itself, run `openssl dgst -sha384 -binary modern-xlsx.min.js | openssl base64 -A`.

### Web workers

Keep the main thread responsive by running XLSX operations in a Web Worker:

```typescript
import { createXlsxWorker, Workbook } from 'modern-xlsx';

// `using` terminates the worker when the scope ends (Symbol.dispose)
using worker = createXlsxWorker({
  // dist/modern-xlsx.worker.js, served by you
  workerUrl: '/modern-xlsx.worker.js',
  // Optional; a relative URL resolves against the page
  wasmUrl: '/modern-xlsx.wasm',
});

const data = await worker.readBuffer(xlsxBytes); // WorkbookData (plain JSON)
const wb = new Workbook(data); // wrap it to use the Workbook API
wb.addSheet('Summary').cell('A1').value = 'Edited on the main thread';
const output = await worker.writeBuffer(wb.toJSON()); // Uint8Array
```

- `readBuffer` resolves to `WorkbookData`, the plain JSON model; `new Workbook(data)` wraps it. Pass `{ password }` as the second argument for an encrypted file.
- When the `Uint8Array` spans its whole `ArrayBuffer`, that buffer is transferred to the worker and detached, so `xlsxBytes` is unusable afterwards. A view on part of a buffer, or on a `SharedArrayBuffer`, is copied instead.
- `writeBuffer` resolves to a `Uint8Array` whose buffer the worker transfers back.
- Without `wasmUrl`, the worker loads `modern-xlsx.wasm` from beside its own script.
- `terminate()` (or the end of a `using` scope) rejects pending calls with `WORKER_ERROR`.

### Cloudflare Workers

Workers import a `.wasm` file as a compiled `WebAssembly.Module`, and the `modern-xlsx/wasm` subpath points at the binary, so `initWasmSync` initializes it at module load:

```typescript
import { initWasmSync, Workbook } from 'modern-xlsx';
import wasmModule from 'modern-xlsx/wasm';

initWasmSync(wasmModule);
```

The full Worker, with its `wrangler` setup, is in [examples/edge/cloudflare-worker](https://github.com/ABCrimson/modern-xlsx/tree/v1.3.0/examples/edge/cloudflare-worker).

### Initialization

Call `await initWasm()` once before any read, write, `validate`, `repair` or `StreamingXlsxWriter` call; no API initializes WASM on its own.

```typescript
import { initWasm } from 'modern-xlsx';
await initWasm(); // call once before any operation
```

`ensureReady()` initializes on its first call and is a no-op afterwards, so a lazily loaded feature can await it at every entry point:

```typescript
import { ensureReady, readBuffer } from 'modern-xlsx';

await ensureReady(); // initializes on the first call, a no-op afterwards
const wb = await readBuffer(bytes);
```

Pass a URL, or a fetch `Response`, when the binary is not at its default location:

```typescript
import { initWasm } from 'modern-xlsx';

// Custom URL
await initWasm('https://my-cdn.com/modern-xlsx.wasm');

// From fetch Response
const res = await fetch('/wasm/modern-xlsx.wasm');
await initWasm(res);
```

## Features

| Feature                         | What you get                                                                                                            |
| ------------------------------- | ----------------------------------------------------------------------------------------------------------------------- |
| Read and write `.xlsx`          | Values, formulas, shared strings, rich text, named ranges, document properties.                                         |
| Cell styles                     | Fonts, fills, borders, alignment and text wrap, number formats, cell protection, through `StyleBuilder`.                |
| Formatted cell values           | `formatCell` renders a value with an Excel number format code.                                                          |
| Validation, conditional formats | Data validation and conditional formatting, read and written.                                                           |
| Sheet layout                    | Frozen and split panes, column widths, row heights, merged cells, auto filter, row and column grouping.                 |
| Printing                        | Page setup, margins, headers and footers, page breaks, print titles and print areas.                                    |
| Hyperlinks                      | External links are written as relationships (`r:id`); internal links use `location`.                                    |
| Comments                        | Notes and threaded comments with replies, written in the form Excel shows ([details](#threaded-comments)).              |
| Sheet and workbook protection   | Read and write.                                                                                                         |
| Excel tables                    | Names checked against Excel's rules, `SUBTOTAL` totals rows, renamed by `cloneSheet` ([details](#tables)).              |
| Images and barcodes             | PNG, JPEG and GIF images; QR, Code 128, EAN-13, UPC-A, Code 39, ITF-14, GS1-128 ([details](#barcodes-and-qr-codes)).    |
| Charts                          | 10 modelled types; other chart XML is preserved verbatim ([details](#charts)).                                          |
| Sparklines                      | Read and write.                                                                                                         |
| Pivot tables                    | Read and write with `PivotTableBuilder`; Excel computes the values when it opens the file ([details](#pivot-tables)).   |
| Slicers and timelines           | Read and write, in a custom dialect Excel has not verified ([details](#slicers-and-timelines)).                         |
| Formula engine                  | 54 functions with Excel semantics; 3-D and structured references evaluate to `#REF!`.                                   |
| Encryption                      | Reads Agile and Standard encryption; writes Agile encryption (AES-256, SHA-512) ([details](#encrypt-a-workbook)).       |
| Dates                           | `dateToSerial` takes a `Date` or a `TemporalLike`, such as a `Temporal.PlainDate`, and `serialToDate` returns a `Date`. |
| Validation and repair           | `validate()` reports structural problems; `repair()` fixes the repairable ones.                                         |
| Streaming writer                | `StreamingXlsxWriter` writes rows one at a time without building a `Workbook` (see [Limits](#limits)).                  |
| Command line                    | `modern-xlsx info` and `modern-xlsx convert` (CSV or JSON).                                                             |
| Lite build                      | `modern-xlsx/lite` loads a smaller WASM build without encryption ([entry points](#runtimes-and-entry-points)).          |

PDF417 and Data Matrix barcodes are experimental: their output is not yet standard-conformant and may not scan. For comparisons with other libraries, see the [feature comparison](https://github.com/ABCrimson/modern-xlsx/blob/v1.3.0/docs/FEATURE-COMPARISON.md).

## Performance

Measured on 2026-11-13 (UTC) on Intel Core Ultra 9 290HX Plus, with Node v26.13.0 and SheetJS 0.20.3.

| 10 columns               | modern-xlsx | SheetJS | Ratio |
| ------------------------ | ----------: | ------: | ----: |
| Read 10K rows (ms)       |        43.3 |   140.5 | 3.25x |
| Read 100K rows (ms)      |       484.8 | 1,591.3 | 3.28x |
| Write 10K rows (ms)      |       136.6 |   106.8 | 0.78x |
| File size, 10K rows (KB) |         479 |   1,245 | 2.60x |

Ratio is the SheetJS figure divided by the modern-xlsx figure: above 1x favours modern-xlsx.

The figures are indicative and come from a single machine; reproduce them with `pnpm -C packages/modern-xlsx bench`. In this run, modern-xlsx read 3.3x faster and wrote files 2.6x smaller, while its cell-by-cell write took longer (ratio 0.78x).

Sizes of the files in the 1.3.0 npm tarball:

| Build                     | Raw bytes | gzip -9 bytes |
| ------------------------- | --------: | ------------: |
| Full WASM                 | 2,484,949 |       917,228 |
| Lite WASM (no encryption) | 2,301,661 |       854,223 |
| IIFE bundle (JS)          |   173,160 |        54,719 |

Files: `wasm/modern_xlsx_wasm_bg.wasm`, `wasm-lite/modern_xlsx_wasm_bg.wasm` and `dist/modern-xlsx.min.js`, measured on 2026-11-14 (UTC) with `wc -c` and `gzip -9 -n -c FILE | wc -c` in the tarball that CI builds on its Linux runner from the 1.3.0 release tree. [WASM size and speed](https://github.com/ABCrimson/modern-xlsx/blob/v1.3.0/docs/perf-size.md#final-figures-and-budgets) records the toolchain and the commands.

## How it works

<p align="center">
<picture>
  <source media="(prefers-color-scheme: dark)" srcset="https://raw.githubusercontent.com/ABCrimson/modern-xlsx/v1.3.0/docs/assets/readme/architecture-dark.svg">
  <img alt="Architecture of modern-xlsx in three layers: the TypeScript API (the modern-xlsx package), the WASM boundary built with wasm-bindgen, and the Rust core (the modern-xlsx-core crate). Read path: readBuffer, or readFile which loads the bytes and calls readBuffer, passes the bytes to the WASM read(data, limits). The Rust core detects whether the file is ZIP or OLE2, unzips it within the ZIP limits, SAX-parses each part with quick-xml and streams a JSON string of at most maxJsonLength bytes back across the boundary, where JSON.parse turns it into a Workbook. Write path: Workbook.toBuffer() passes JSON.stringify output to the WASM write(json); the Rust core decodes it with serde_json into WorkbookData, writes the parts with the shared string table built inline, zips them with Deflate at level 6 and returns a Uint8Array. Encryption is a Cargo feature in the full build only, not in /lite: it gates the WASM exports readWithPassword and writeWithPassword, which decrypt and encrypt as listed." src="https://raw.githubusercontent.com/ABCrimson/modern-xlsx/v1.3.0/docs/assets/readme/architecture-light.svg" width="640">
</picture>
</p>

Data crosses the WASM boundary as a JSON string, which JavaScript parses with `JSON.parse`. The Rust core (Edition 2024) does the ZIP compression and decompression, parses and writes the XML parts with quick-xml, and builds the shared string table and the styles.

## API overview

### Workbook

```typescript
const wb = new Workbook();

wb.sheetNames; // readonly string[]
wb.sheetCount; // number
wb.dateSystem; // 'date1900' | 'date1904'
wb.styles; // StylesData

wb.addSheet('Name'); // Worksheet
wb.getSheet('Name'); // Worksheet | undefined
wb.getSheetByIndex(0); // Worksheet | undefined
wb.removeSheet('Name'); // boolean
wb.removeSheet(0); // boolean

// Named ranges
// A name Excel refuses (TRUE, A1, R2D2, or a table's name)
// throws INVALID_ARGUMENT
wb.addNamedRange('MyRange', 'Sheet1!$A$1:$D$10');
wb.getNamedRange('MyRange'); // DefinedNameData | undefined
wb.removeNamedRange('MyRange'); // boolean

// Document properties
wb.docProperties = { title: 'Report', creator: 'App' };

// Workbook views
wb.workbookViews = [
  {
    activeTab: 0,
    firstSheet: 0,
    showHorizontalScroll: true,
    showVerticalScroll: true,
    showSheetTabs: true,
  },
];
```

### Worksheet

#### Cells, columns and rows

```typescript
const ws = wb.addSheet('Sheet1');

// Cell access
ws.cell('A1').value = 'Hello';
ws.cell('B1').value = 42;
ws.cell('C1').value = true;
ws.cell('D1').formula = 'SUM(B1:B100)';

// Columns & rows (1-based indexes)
ws.setColumnWidth(1, 20); // column A, in character units
// Row 1, in points, written as a custom height (customHeight)
ws.setRowHeight(1, 30);
ws.setRowHidden(2, true);
```

#### Merged cells, frozen panes and auto filter

```typescript
// Merge cells
ws.addMergeCell('A1:D1');
ws.removeMergeCell('A1:D1');

// Frozen panes
ws.frozenPane = { rows: 1, cols: 0 };

// Auto filter
ws.autoFilter = 'A1:D100';
ws.autoFilter = {
  range: 'A1:D100',
  filterColumns: [{ colId: 0, filters: ['Yes'] }],
};
```

#### Hyperlinks

```typescript
// A value with a URI scheme (https:, mailto:, file:) is an external
// link, written as a relationship (r:id, TargetMode="External");
// anything else (Sheet2!A1, 'Sheet Name'!A1, a same-sheet range, a
// defined name) is a location in the workbook
const link = { display: 'Example', tooltip: 'Opens example.com' };
ws.addHyperlink('A1', 'https://example.com', link);
ws.addHyperlink('A2', 'Sheet2!A1', { display: 'Go to Sheet2' });
ws.removeHyperlink('A2');
```

#### Data validation and notes

```typescript
// Data validation
ws.addValidation('B2:B100', {
  validationType: 'list',
  operator: null,
  formula1: '"Yes,No,Maybe"',
  formula2: null,
  allowBlank: true,
  showErrorMessage: true,
  errorTitle: 'Invalid',
  errorMessage: 'Pick from list',
});

// Comments
ws.addComment('A1', 'Author', 'This is a comment');
ws.removeComment('A1');
```

Threaded comments, pivot tables, slicers and timelines have their own sections below.

#### Sheet protection

```typescript
// Sheet protection: `sheet: true` turns it on; for the other flags,
// true = locked and false = allowed while the sheet is protected
ws.sheetProtection = {
  sheet: true,
  objects: false,
  scenarios: false,
  formatCells: false,
  formatColumns: false,
  formatRows: false,
  insertColumns: false,
  insertRows: false,
  deleteColumns: false,
  deleteRows: false,
  sort: false,
  autoFilter: false,
};
```

### Tables

```typescript
// Excel Tables (ListObjects)
ws.addTable({
  id: 1,
  displayName: 'SalesData',
  ref: 'A1:B3',
  headerRowCount: 1,
  totalsRowCount: 0,
  totalsRowShown: false,
  columns: [
    { id: 1, name: 'Product' },
    { id: 2, name: 'Revenue' },
  ],
  styleInfo: {
    name: 'TableStyleMedium9',
    showFirstColumn: false,
    showLastColumn: false,
    showRowStripes: true,
    showColumnStripes: false,
  },
});
ws.tables; // readonly TableDefinitionData[]
ws.getTable('SalesData'); // TableDefinitionData | undefined (by display name)
ws.removeTable('SalesData'); // boolean
```

`addTable` refuses a table name Excel will not open and the `@` this-row shorthand in a table definition, and fills the empty cells of a totals row with the `SUBTOTAL` formulas Excel writes; `cloneSheet` renames a cloned table and the structured references to it in the copy.

### Page layout and printing

```typescript
import { HeaderFooterBuilder } from 'modern-xlsx';

// Page setup & margins
ws.pageSetup = { orientation: 'landscape', paperSize: 9 };
ws.pageMargins = {
  top: 0.75,
  bottom: 0.75,
  left: 0.7,
  right: 0.7,
  header: 0.3,
  footer: 0.3,
};

// Headers & Footers
ws.headerFooter = {
  oddHeader: new HeaderFooterBuilder()
    .left(HeaderFooterBuilder.date())
    .center(HeaderFooterBuilder.bold('Report'))
    .right(`Page ${HeaderFooterBuilder.pageNumber()}`)
    .build(),
};

// Row & Column Grouping
ws.groupRows(2, 10); // outline level 1
ws.groupRows(3, 5, 2); // nested level 2
ws.collapseRows(2, 10);
ws.expandRows(2, 10);
ws.groupColumns(1, 3); // columns A-C
ws.outlineProperties = { summaryBelow: true, summaryRight: true };

// Print Titles & Areas (0-based sheet index + range reference; an
// area without a sheet name is stored qualified with that sheet's
// name, so 'A1:G50' becomes 'Sheet1!A1:G50')
wb.setPrintTitles(0, 'Sheet1!$1:$1'); // repeat row 1 on every page
wb.setPrintArea(0, 'Sheet1!$A$1:$G$50'); // limit the printed area
```

### Styles

Fluent builder that produces a style index for any cell:

```typescript
const idx = wb
  .createStyle()
  .font({ name: 'Arial', size: 12, bold: true, color: 'FF0000' })
  .fill({ pattern: 'solid', fgColor: 'FFFF00' })
  .border({
    top: { style: 'thin', color: '000000' },
    bottom: { style: 'double', color: '000000' },
    left: { style: 'thin', color: '000000' },
    right: { style: 'thin', color: '000000' },
  })
  .alignment({
    horizontal: 'center',
    vertical: 'top',
    wrapText: true,
    textRotation: 45,
  })
  .protection({ locked: true, hidden: false })
  .numberFormat('#,##0.00')
  .build(wb.styles);

ws.cell('A1').styleIndex = idx;
```

### Rich text

```typescript
import { RichTextBuilder } from 'modern-xlsx';

const runs = new RichTextBuilder()
  .text('Normal ')
  .bold('Bold ')
  .italic('Italic ')
  .styled('Custom', { color: 'FF0000', fontSize: 14 })
  .build();
```

### Charts

Create Excel charts with the fluent ChartBuilder API:

```typescript
import { ChartBuilder } from 'modern-xlsx';

const ws = wb.addSheet('Sales');
ws.cell('A1').value = 'Quarter';
ws.cell('B1').value = 'Revenue';
ws.cell('A2').value = 'Q1';
ws.cell('A3').value = 'Q2';
ws.cell('B2').value = 100;
ws.cell('B3').value = 200;

// Create a bar chart via the builder callback.
// `name` is a cell-reference formula: the series is named by
// the text in Sales!B1
ws.addChart('bar', (b) => {
  b.title('Quarterly Sales')
    .addSeries({
      name: 'Sales!$B$1',
      catRef: 'Sales!$A$2:$A$3',
      valRef: 'Sales!$B$2:$B$3',
    })
    .catAxis({ title: 'Quarter' })
    .valAxis({ title: 'Amount ($)' })
    .legend('bottom')
    .anchor({ col: 3, row: 0 }, { col: 12, row: 18 });
});

// Or build the data model yourself and attach it.
// `nameLiteral` names the series with literal text
const chart = new ChartBuilder('line')
  .title('Trend')
  .addSeries({ nameLiteral: 'Revenue', valRef: 'Sales!$B$2:$B$3' })
  .build();
ws.addChartData(chart);

// Supported types: bar, column, line, pie, area, scatter, radar,
// doughnut, bubble, stock
```

`ChartBuilder#style(id)` writes `c:style` and accepts Excel's chart styles 1-48. Bundled colour palettes cover styles 1-8 only, and `getChartStylePalette` throws a `RangeError` for any other id. 3-D, surface, bar-of-pie and pie-of-pie charts read from a file are kept verbatim and written back unchanged.

### Formula engine

Evaluate formulas in memory with 54 built-in Excel functions. Results follow Excel's semantics for type coercion, error propagation and operators. 3-D references (`Sheet1:Sheet3!A1`) and structured references (`Table1[Column]`) evaluate to `#REF!`. `OFFSET` and `INDIRECT` are not implemented, so, like any function the registry lacks, they evaluate to `#NAME?`.

```typescript
import {
  createDefaultFunctions,
  decodeRange,
  type EvalContext,
  encodeCellRef,
  evaluateFormula,
} from 'modern-xlsx';

const ctx: EvalContext = {
  // getCell receives the sheet name, a 0-based column index, and a 1-based row
  getCell: (sheet, col, row) => {
    const ref = encodeCellRef(row - 1, col);
    return wb.getSheet(sheet)?.cell(ref).value ?? null;
  },
  currentSheet: 'Sheet1',
  functions: createDefaultFunctions(),
  // Optional: the populated extent of a sheet (1-based maxRow,
  // 0-based maxCol), so a whole-column range such as SUM(A:A) reads
  // only populated cells
  getBounds: (sheet) => {
    // e.g. 'A1:D20', as read from the file
    const dimension = wb.getSheet(sheet)?.dimension;
    if (!dimension?.includes(':')) return undefined;
    const { end } = decodeRange(dimension);
    return { maxRow: end.row + 1, maxCol: end.col };
  },
};

const result = evaluateFormula('SUM(A1:A10)', ctx);
// Supports: SUM, AVERAGE, COUNT, IF, VLOOKUP, INDEX/MATCH,
// CONCATENATE, and 46 more
```

`EvalContext` also takes `currentCell` (`{ col, row }`, 0-based column and 1-based row), which a multi-cell reference in a scalar position needs for Excel's implicit intersection, and `dateSystem`, which `TEXT` uses to render serial numbers as dates.

### Pivot tables

```typescript
const wb = await readBuffer(buffer);
const ws = wb.getSheet('PivotReport');
if (!ws) throw new Error('missing sheet PivotReport');

// Read pivot table metadata (ws.pivotTables: readonly PivotTableData[])
for (const pt of ws.pivotTables) {
  console.log(pt.name, pt.location);
}

// Add a pivot table
ws.addPivotTable({
  name: 'SalesPivot',
  cacheId: 0,
  location: { ref: 'A3:D20' },
  pivotFields: [
    {
      items: [],
      name: 'Region',
      axis: 'axisRow',
      subtotals: [],
      compact: true,
      outline: true,
    },
  ],
  rowFields: [{ x: 0 }],
  colFields: [],
  dataFields: [],
  pageFields: [],
});

// Workbook-level pivot caches
wb.addPivotCache({ source: { ref: 'A1:D100', sheet: 'Data' } });

ws.removePivotTable(0); // boolean
```

`ws.addPivotTableFromBuilder((b) => ...)` builds the same definition with `PivotTableBuilder` (`name`, `cacheId`, `location`, `addRowField`, `addDataField`, ...).

The writer completes what Excel needs and leaves your model as you built it. Each row, column or page field with `items: []` (as above, and as every `PivotTableBuilder` field) is written with one item per shared item of the cache field at the same index (`fields[i].sharedItems` of the cache; the cache above has none, so it gets only the subtotal item), plus a subtotal item. The cache its `cacheId` names is written with `refreshOnLoad="1"`, and `<pivotFields>` is extended to the number of fields in the cache. The writer does not compute the table's cells: Excel lays out the table and fills in the values when it opens the file. A table whose `cacheId` names no cache fails the write with `INVALID_WORKBOOK`.

A cache's `source` is a range (`ref` and `sheet`) or, as Excel writes it for a pivot table made from a table or a named range, a `name` with `ref` and `sheet` empty (`wb.addPivotCache({ source: { ref: '', sheet: '', name: 'Sales' } })`); either is written as read, and an empty `ref` or `sheet` is left out of the file. A source that is not a worksheet, such as multiple consolidation ranges, is read into `source.rawXml` and written verbatim. A source in another workbook keeps its relationship to that workbook.

### Threaded comments

```typescript
const ws = wb.getSheet('Sheet1');
if (!ws) throw new Error('missing sheet Sheet1');

// Add a threaded comment (returns its id, an upper-case GUID in
// braces, the form Excel needs)
const id = ws.addThreadedComment('A1', 'Review this value', 'Alice');
// A reply to a reply joins the thread root
ws.replyToComment(id, 'Looks correct to me', 'Bob');

// Read threaded comments (ws.threadedComments: readonly ThreadedCommentData[]):
// authors are workbook persons, referenced by personId
const persons = wb.toJSON().persons ?? [];
const authors = new Map(persons.map((p) => [p.id, p.displayName]));
for (const tc of ws.threadedComments) {
  console.log(tc.refCell, tc.text, authors.get(tc.personId));
}
// A1 Review this value Alice
// A1 Looks correct to me Bob
```

Excel shows a thread only when its cell also has a legacy note whose author is `tc={thread id}`. The writer adds that placeholder note, with its note drawing, for each thread whose cell has no note, without changing the model; reading the file back returns the placeholders in `ws.comments`, as for a file Excel saved. A thread on a cell that already has an ordinary note gets none, and Excel shows the note. Replies show in Excel in the order they were added.

A cell holds one thread: `addThreadedComment` on a cell that already has a thread root throws `INVALID_ARGUMENT` naming the root's id, so add to the thread with `replyToComment`. Comment text is limited to 32,767 UTF-16 code units (an emoji counts as two), Excel's limit; a longer one throws `INVALID_ARGUMENT`.

### Slicers and timelines

Slicers and timelines are written in a custom dialect that Excel has not verified, and the links Excel needs to load them are not written. The reader accepts both that dialect and the MS-XLSX x14/x15 parts Excel saves, so slicers and timelines read from an Excel file are rewritten in the custom dialect on write.

```typescript
const ws = wb.getSheet('Dashboard');
if (!ws) throw new Error('missing sheet Dashboard');

// Read slicers & timelines (readonly SlicerData[] and TimelineData[])
for (const slicer of ws.slicers) console.log(slicer.name, slicer.caption);
for (const tl of ws.timelines) console.log(tl.name, tl.caption);

// Add a slicer
ws.addSlicer({
  name: 'Slicer_Region',
  cacheName: 'Slicer_Region',
  caption: 'Region',
});
wb.addSlicerCache({
  name: 'Slicer_Region',
  sourceName: 'Region',
  items: [
    { n: 'East', s: true },
    { n: 'West', s: false },
  ],
});

// Add a timeline
ws.addTimeline({
  name: 'Timeline_Date',
  cacheName: 'NativeTimeline_Date',
  caption: 'Date',
  level: 'months',
});
wb.addTimelineCache({
  name: 'NativeTimeline_Date',
  sourceName: 'Date',
  selectionStart: '2024-01-01',
  selectionEnd: '2024-12-31',
});

// Remove by index
ws.removeSlicer(0); // boolean
ws.removeTimeline(0); // boolean
```

### Barcodes and QR codes

Generate barcodes and embed them as images — pure TypeScript, zero dependencies:

```typescript
import { Workbook, encodeQR, renderBarcodePNG } from 'modern-xlsx';

const wb = new Workbook();
wb.addSheet('Labels');

// High-level: generate + embed in one call
const anchor = { fromCol: 0, fromRow: 0, toCol: 4, toRow: 4 };
wb.addBarcode('Labels', anchor, 'https://example.com', {
  type: 'qr',
  ecLevel: 'M',
});

// Or use the low-level pipeline:
const matrix = encodeQR('Hello', { ecLevel: 'H' });
const png = renderBarcodePNG(matrix, {
  moduleSize: 6,
  showText: true,
  textValue: 'Hello',
});
wb.addImage('Labels', { fromCol: 5, fromRow: 0, toCol: 9, toRow: 4 }, png);
```

**Supported formats:** QR Code, Code 128, EAN-13, UPC-A, Code 39, ITF-14, GS1-128, and, experimentally, PDF417 and Data Matrix. The experimental encoders' output is not ISO/IEC 15438 (PDF417) or ISO/IEC 16022 (Data Matrix) conformant and may not scan. `modern-xlsx/lite` has no standalone encoders (see [Runtimes and entry points](#runtimes-and-entry-points)).

See the [Barcode Guide](https://github.com/ABCrimson/modern-xlsx/blob/v1.3.0/docs/guide/barcodes.md) for format comparison and usage details.

### Table layout

`drawTable(wb, ws, opts)` and `drawTableFromData(wb, ws, data, opts?)` write a styled table from headers and rows, with column widths, zebra stripes, merges and per-cell styles. See the [Table Layout Guide](https://github.com/ABCrimson/modern-xlsx/blob/v1.3.0/docs/guide/tables.md).

### Validation and repair

The Rust core checks a workbook for structural problems and repairs the fixable ones:

```typescript
const wb = new Workbook();
const ws = wb.addSheet('Sheet1');
ws.cell('A1').value = 'Hello';
ws.cell('A1').styleIndex = 999; // dangling!

// Validate — returns a structured report
const report = wb.validate();
console.log(report.isValid); // false
console.log(report.errorCount); // 1
// 2: the error plus an info issue (no theme colours)
console.log(report.issues.length);
console.log(report.issues[0]);
// {
//   severity: 'error',
//   category: 'styleIndex',
//   message: 'Cell A1 styleIndex=999 exceeds cellXfs count (1)',
//   location: 'Sheet1!A1',
//   suggestion: 'Clamp styleIndex to 0',
//   autoFixable: true
// }

// Repair — returns a repaired copy; `wb` itself is unchanged
const { workbook, report: postReport, repairCount } = wb.repair();
console.log(repairCount); // 2 (style index + missing theme)
console.log(postReport.isValid); // true
console.log(workbook.getSheet('Sheet1')?.cell('A1').styleIndex); // 0
```

Each issue has a `severity` (`error`, `warning` or `info`) and a `category` (`IssueCategory`): `cellReference`, `styleIndex`, `mergeCell`, `sharedString`, `sheetName`, `definedName`, `dataValidation`, `conditionalFormatting`, `theme`, `metadata`, `structure` or `sheetProtection`. The report keeps at most 10,000 issues and 16 MiB of message and location text, and `omittedIssueCount` says how many it left out. [What each category detects, and how the report counts](https://github.com/ABCrimson/modern-xlsx/blob/v1.3.0/docs/guide/validation.md).

### Errors

```typescript
import { ModernXlsxError } from 'modern-xlsx';

try {
  ws.replyToComment('nonexistent-id', 'text', 'author');
} catch (e) {
  if (e instanceof ModernXlsxError) {
    console.log(e.code); // 'COMMENT_NOT_FOUND'
    console.log(e.message); // 'Comment nonexistent-id not found'
  }
}
```

### Utilities

```typescript
import {
  dateToSerial,
  serialToDate,
  isDateFormatId,
  isDateFormatCode,
  formatCell,
  encodeCellRef,
  decodeCellRef,
  encodeRange,
  decodeRange,
  columnToLetter,
  letterToColumn,
} from 'modern-xlsx';

// Date conversion: dateToSerial takes a Date or a TemporalLike
// object ({ year, month, day, hour?, ... }), which
// Temporal.PlainDate and PlainDateTime satisfy.
// 45292 (a Date counts in UTC):
const serial = dateToSerial(new Date(Date.UTC(2024, 0, 1)));
const serial2 = dateToSerial({ year: 2024, month: 1, day: 1 }); // 45292
const date = serialToDate(45292); // Date for 2024-01-01T00:00:00.000Z

// Format cell value using Excel format code
formatCell(45292, 'yyyy-mm-dd'); // "2024-01-01"

// Cell references
decodeCellRef('B3'); // { col: 1, row: 2 } (0-based)
encodeCellRef(2, 1); // "B3" (row, col)
```

Sheet conversion (`aoaToSheet`, `jsonToSheet`, `sheetToJson`, `sheetToCsv`, `sheetToHtml`, `sheetAddAoa`, `sheetAddJson`) is under [Convert rows to and from JSON and CSV](#convert-rows-to-and-from-json-and-csv).

### Command line

```bash
# Show sheet names, row counts and dimensions
npx modern-xlsx info report.xlsx

# Convert every sheet to JSON (the default format)
npx modern-xlsx convert report.xlsx output.json

# Export one sheet (0-based index) as CSV
npx modern-xlsx convert report.xlsx sheet1.csv --sheet 0 --format csv
```

`npx modern-xlsx --help` prints the usage. `convert` takes two options:

- `--sheet <n>`: the 0-based sheet index (default: 0 for CSV, all sheets for JSON).
- `--format csv|json`: the output format (default: `json`).

The output format comes from `--format`, not from the file name: `convert in.xlsx out.csv` without `--format csv` writes JSON. Inferring it from the extension is planned for 1.10.0.

Exit codes: `0` success, `1` runtime failure, `2` usage error. Errors go to stderr as `modern-xlsx: <message>`.

## Limits

A read builds the whole workbook as one JSON string before JavaScript parses it. `ReadOptions.limits` bounds that work:

- `maxDecompressedSize`: default 2,147,483,648 (2 GiB), which is also the ceiling. Exceeding it fails with `SECURITY`.
- `maxCompressionRatio`: default 100 (per ZIP entry), no ceiling. Exceeding it fails with `SECURITY`.
- `maxEntries`: default 100,000, ceiling 1,048,576. Exceeding it fails with `SECURITY`.
- `maxJsonLength`: default 536,870,888 (UTF-8 bytes), which is also the ceiling. Exceeding it fails with `SIZE_LIMIT`.

A value above a ceiling is clamped to it. A value that is not a finite number greater than zero, or not a safe integer for the three integer fields, throws `INVALID_ARGUMENT`. Raise a limit only for files you trust:

```typescript
import { initWasm, readBuffer } from 'modern-xlsx';

await initWasm();
// bytes: a Uint8Array holding the .xlsx file
const wb = await readBuffer(bytes, { limits: { maxCompressionRatio: 200 } });
```

- A workbook whose JSON exceeds 536,870,888 bytes (V8's maximum string length) cannot be read yet. Split the file into smaller workbooks. A streaming read is planned ([#31](https://github.com/ABCrimson/modern-xlsx/issues/31)).
- `maxJsonLength` also bounds the text a read copies into cells, comments, hyperlinks and kept extensions, and the memory of the records it keeps, so a small file cannot expand to gigabytes in memory: such a read rejects with `SIZE_LIMIT`.
- Text that repeats one character thousands of times compresses past the default 100:1 ratio and rejects with `SECURITY`. For files you trust, raise it: `{ limits: { maxCompressionRatio: 10_000 } }`.
- `XlsxWorker.readBuffer` takes no `limits`; it reads with the defaults.
- Cell references are limited to `A1:XFD1048576` (16,384 columns, 1,048,576 rows); a reference outside that range throws `INVALID_CELL_REF`.
- An encrypted write fails with `SIZE_LIMIT` above 2 GiB, the most an MS-CFB version 3 stream can hold.
- `StreamingXlsxWriter` writes rows into a compressed ZIP as they arrive, but it holds that compressed output in memory until `finish()` returns it.

[What `maxJsonLength` counts, the repeated-text case and the encrypted-write limit, in full](https://github.com/ABCrimson/modern-xlsx/blob/v1.3.0/docs/guide/limits.md).

## Security

The reader treats every file as untrusted. The limits above bound each read: a ZIP that exceeds the size, compression-ratio or entry-count limit rejects with `SECURITY`, and a read whose workbook JSON exceeds `maxJsonLength` rejects with `SIZE_LIMIT`, checked as the JSON is built. How the reader defends against hostile ZIP, XML and OLE2 input is in the [threat model](https://github.com/ABCrimson/modern-xlsx/blob/v1.3.0/docs/security/threat-model.md).

To report a vulnerability, do not open a public issue: [SECURITY.md](https://github.com/ABCrimson/modern-xlsx/blob/v1.3.0/SECURITY.md) describes how to report it privately.

## TypeScript

The package ships its declarations, and every type is exported by name: `WorkbookData`, `CellData`, `StylesData`, `ReadOptions`, `ValidationReport`, `ChartDataModel`, `PivotTableData` and the rest. The [API reference](https://abcrimson.github.io/modern-xlsx/api/) documents them all, and [the list of main types](https://github.com/ABCrimson/modern-xlsx/blob/v1.3.0/docs/guide/types.md) groups them by feature.

### TypeScript setup

- The declarations need TypeScript 5.7 or newer: `Workbook.toBuffer()` resolves to `Uint8Array<ArrayBuffer>`.
- They use `Symbol.dispose` (`StreamingXlsxWriter` and `XlsxWorker` are disposable). With `skipLibCheck: false`, include the `esnext.disposable` lib, which `lib: ["esnext"]` and `target: "esnext"` include and `es2026` does not.
- They also need the `dom` lib, for `WebAssembly`, `BufferSource`, `URL`, `Response` and `Blob`. `@types/node` declares `Symbol.dispose`, `URL`, `Response` and `Blob` but not `WebAssembly` or `BufferSource`, so it does not replace `dom`.
- Checked with TypeScript 7.1.0-dev.20261114.1 on 2026-11-14: `lib: ["esnext", "dom"]` and `lib: ["es2024", "dom"]` with `types: ["node"]` pass; `lib: ["esnext"]` fails with or without `types: ["node"]`.

```json
{ "compilerOptions": { "lib": ["esnext", "dom"] } }
```

## License

[MIT](https://github.com/ABCrimson/modern-xlsx/blob/v1.3.0/packages/modern-xlsx/LICENSE)
