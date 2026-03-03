<p align="center">
  <h1 align="center">modern-xlsx</h1>
  <p align="center">
    High-performance XLSX read/write for JavaScript &amp; TypeScript, powered by Rust + WASM.
  </p>
</p>

<p align="center">
  <a href="https://www.npmjs.com/package/modern-xlsx"><img alt="npm" src="https://img.shields.io/npm/v/modern-xlsx?color=cb0000&label=npm&logo=npm"></a>
  <a href="https://github.com/ABCrimson/modern-xlsx/blob/main/packages/modern-xlsx/LICENSE"><img alt="License" src="https://img.shields.io/badge/license-MIT-blue"></a>
  <img alt="Types" src="https://img.shields.io/badge/types-included-blue?logo=typescript&logoColor=white">
  <img alt="Zero deps" src="https://img.shields.io/badge/dependencies-0-brightgreen">
</p>

---

Full cell styling, data validation, conditional formatting, frozen panes, hyperlinks, comments, sheet protection, and more — features that SheetJS locks behind a **paid Pro license** — all **free and open source**.

## Install

```bash
npm install modern-xlsx
```

> Requires a runtime with WASM support (Node.js 24+, Bun, Deno, modern browsers).

## Quick Start

```typescript
import { initWasm, Workbook, readFile } from 'modern-xlsx';

await initWasm();

// --- Create ---
const wb = new Workbook();
const ws = wb.addSheet('Sales');

ws.cell('A1').value = 'Product';
ws.cell('B1').value = 'Revenue';
ws.cell('A2').value = 'Widget';
ws.cell('B2').value = 9999.99;

// Style headers
const header = wb.createStyle()
  .font({ bold: true, size: 14, color: '1F4E79' })
  .fill({ pattern: 'solid', fgColor: 'D6E4F0' })
  .alignment({ horizontal: 'center' })
  .border({ bottom: { style: 'medium', color: '1F4E79' } })
  .build(wb.styles);

ws.cell('A1').styleIndex = header;
ws.cell('B1').styleIndex = header;

// Number format
const currency = wb.createStyle()
  .numberFormat('$#,##0.00')
  .build(wb.styles);
ws.cell('B2').styleIndex = currency;

await wb.toFile('sales.xlsx');

// --- Read ---
const existing = await readFile('report.xlsx');
console.log(existing.getSheet('Sheet1')?.cell('A1').value);
```

## Performance

Benchmarks on a 100,000-row workbook (Node.js, single thread):

| Operation | modern-xlsx | SheetJS CE | Factor |
|-----------|------------:|-----------:|-------:|
| **Read 100K rows** | 1,377 ms | 5,942 ms | **4.3x faster** |
| **Read 10K rows** | 156 ms | 732 ms | **4.7x faster** |
| **Write 10K** (cell-by-cell) | 469 ms | 645 ms | **1.4x faster** |
| **Write 100K** (batch) | 1,172 ms | 6,248 ms | **5.3x faster** |
| aoaToSheet (50K) | 242 ms | 508 ms | **2.1x faster** |
| sheetToCsv (10K) | 134 ms | 151 ms | **1.1x faster** |
| sheetToJson (10K) | 130 ms | 127 ms | ~1.0x |

> ESM 133 KB + IIFE 60 KB + WASM 1.1 MB. Zero runtime dependencies. Output files **8.4x smaller**.

## Feature Comparison

| Feature | modern-xlsx | SheetJS CE | SheetJS Pro |
|---------|:-----------:|:----------:|:-----------:|
| Read / write XLSX | Yes | Yes | Yes |
| Cell styling (fonts, fills, borders) | **Free** | No | Paid |
| Number formats | **Free** | Read only | Paid |
| Alignment & text wrap | **Free** | No | Paid |
| Cell protection | **Free** | No | Paid |
| Comments / notes | **Free** | Read only | Paid |
| Rich text (inline formatting) | **Free** | No | Paid |
| Data validation | **Free** | No | Paid |
| Conditional formatting | **Free** | No | Paid |
| Frozen panes | **Free** | Partial | Paid |
| Hyperlinks | **Free** | Yes | Yes |
| Auto filter | **Free** | Basic | Yes |
| Sheet protection | **Free** | Read only | Paid |
| Page setup & margins | **Free** | Margins only | Paid |
| Temporal API dates | **Yes** | No | No |
| Format cell values (SSF) | **Free** | Basic | Full |
| Available on npm | **Yes** | Yes | No |
| Tree-shakable ESM | **Yes** | No | No |
| Strict TypeScript types | **Yes** | Partial | Partial |
| WASM-accelerated I/O | **Yes** | No | No |
| OOXML validation & repair | **Yes** | No | No |
| Barcode & QR code generation | **Yes** | No | No |
| Image embedding | **Yes** | No | Paid |
| Excel Tables (ListObjects) | **Yes** | No | Paid |
| Headers & footers | **Yes** | No | Paid |
| Row/column grouping (outline) | **Yes** | No | Paid |
| Print titles & areas | **Yes** | No | Paid |

## How It Works

```
                    ┌────────────────────────────┐
  TypeScript API    │  Workbook / Worksheet / Cell │
                    └─────────────┬──────────────┘
                                  │ JSON
                    ┌─────────────▼──────────────┐
  WASM boundary     │  wasm-bindgen bridge        │
                    └─────────────┬──────────────┘
                                  │
                    ┌─────────────▼──────────────┐
  Rust core         │  OOXML parser & writer      │
                    │  (quick-xml + zip)          │
                    └────────────────────────────┘
```

Data crosses the WASM boundary as a JSON string — 8-13x faster than `serde_wasm_bindgen` for large workbooks. The Rust core handles all ZIP compression, XML parsing, shared string tables, and style resolution.

---

## API Reference

### Initialization

```typescript
import { initWasm } from 'modern-xlsx';
await initWasm(); // call once before any operation
```

### Reading

```typescript
import { readFile, readBuffer } from 'modern-xlsx';

const wb = await readFile('data.xlsx');        // Node.js / Bun / Deno
const wb = await readBuffer(uint8Array);       // any environment
```

### Writing

```typescript
await wb.toFile('output.xlsx');                // Node.js / Bun / Deno
const buffer = await wb.toBuffer();            // Uint8Array

import { writeBlob } from 'modern-xlsx';
const blob = writeBlob(wb);                    // browser Blob
```

### Workbook

```typescript
const wb = new Workbook();

wb.sheetNames;                     // string[]
wb.sheetCount;                     // number
wb.dateSystem;                     // 'date1900' | 'date1904'
wb.styles;                         // StylesData

wb.addSheet('Name');               // Worksheet
wb.getSheet('Name');               // Worksheet | undefined
wb.getSheetByIndex(0);             // Worksheet | undefined
wb.removeSheet('Name');            // boolean
wb.removeSheet(0);                 // boolean

// Named ranges
wb.addNamedRange('MyRange', 'Sheet1!$A$1:$D$10');
wb.getNamedRange('MyRange');
wb.removeNamedRange('MyRange');

// Document properties
wb.docProperties = { title: 'Report', creator: 'App' };

// Workbook views
wb.workbookViews = [{
  activeTab: 0,
  firstSheet: 0,
  showHorizontalScroll: true,
  showVerticalScroll: true,
  showSheetTabs: true,
}];
```

### Worksheet

```typescript
const ws = wb.addSheet('Sheet1');

// Cell access
ws.cell('A1').value = 'Hello';
ws.cell('B1').value = 42;
ws.cell('C1').value = true;
ws.cell('D1').formula = 'SUM(B1:B100)';

// Columns & rows
ws.setColumnWidth(1, 20);
ws.setRowHeight(1, 30);
ws.setRowHidden(2, true);

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

// Hyperlinks
ws.addHyperlink('A1', '#Sheet2!A1', {
  display: 'Go to Sheet2',
  tooltip: 'Click',
});
ws.removeHyperlink('A1');

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

// Page setup & margins
ws.pageSetup = { orientation: 'landscape', paperSize: 9 };
ws.pageMargins = {
  top: 0.75, bottom: 0.75,
  left: 0.7, right: 0.7,
  header: 0.3, footer: 0.3,
};

// Sheet protection
ws.sheetProtection = {
  sheet: true, objects: false, scenarios: false,
  formatCells: false, formatColumns: false, formatRows: false,
  insertColumns: false, insertRows: false,
  deleteColumns: false, deleteRows: false,
  sort: false, autoFilter: false,
};

// Excel Tables (ListObjects)
ws.addTable({
  name: 'SalesData', ref: 'A1:B3',
  columns: [{ name: 'Product' }, { name: 'Revenue' }],
  style: { name: 'TableStyleMedium9', showRowStripes: true },
});
ws.tables;                         // TableDefinitionData[]
ws.getTable('SalesData');          // TableDefinitionData | undefined
ws.removeTable('SalesData');       // boolean

// Headers & Footers
import { HeaderFooterBuilder } from 'modern-xlsx';
ws.headerFooter = {
  oddHeader: new HeaderFooterBuilder()
    .left(HeaderFooterBuilder.date())
    .center(HeaderFooterBuilder.bold('Report'))
    .right(`Page ${HeaderFooterBuilder.pageNumber()}`)
    .build(),
};

// Row & Column Grouping
ws.groupRows(2, 10);              // outline level 1
ws.groupRows(3, 5, 2);            // nested level 2
ws.collapseRows(2, 10);
ws.expandRows(2, 10);
ws.groupColumns(1, 3);            // columns A-C
ws.outlineProperties = { summaryBelow: true, summaryRight: true };

// Print Titles & Areas
wb.setPrintTitles('Sheet1', { rows: { start: 1, end: 1 } });
wb.setPrintArea('Sheet1', 'A1:G50');
```

### Styles

Fluent builder that produces a style index for any cell:

```typescript
const idx = wb.createStyle()
  .font({ name: 'Arial', size: 12, bold: true, color: 'FF0000' })
  .fill({ pattern: 'solid', fgColor: 'FFFF00' })
  .border({
    top:    { style: 'thin',   color: '000000' },
    bottom: { style: 'double', color: '000000' },
    left:   { style: 'thin',   color: '000000' },
    right:  { style: 'thin',   color: '000000' },
  })
  .alignment({ horizontal: 'center', vertical: 'top', wrapText: true, textRotation: 45 })
  .protection({ locked: true, hidden: false })
  .numberFormat('#,##0.00')
  .build(wb.styles);

ws.cell('A1').styleIndex = idx;
```

### Utilities

```typescript
import {
  aoaToSheet, jsonToSheet, sheetToJson, sheetToCsv, sheetToHtml,
  sheetAddAoa, sheetAddJson,
  dateToSerial, serialToDate, isDateFormatId, isDateFormatCode,
  formatCell,
  encodeCellRef, decodeCellRef, encodeRange, decodeRange,
  columnToLetter, letterToColumn,
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
const data = sheetToJson(ws);     // [{ Name: 'Alice', Age: 30 }, ...]
const csv  = sheetToCsv(ws);      // "Name,Age\nAlice,30\nBob,25"
const html = sheetToHtml(ws);     // "<table>..."

// Append to existing sheet
sheetAddAoa(ws, [['Charlie', 35]], { origin: 'A4' });
sheetAddJson(ws, [{ name: 'Diana', age: 28 }]);

// Limit rows
const first5 = sheetToJson(ws, { sheetRows: 5 });

// Date conversion (Date and Temporal API)
const serial = dateToSerial(new Date(2024, 0, 1));
const serial2 = dateToSerial({ year: 2024, month: 1, day: 1 }); // Temporal-like
const date = serialToDate(45292);

// Format cell value using Excel format code
formatCell(45292, 'yyyy-mm-dd'); // "2024-01-01"

// Cell references
decodeCellRef('B3');  // { row: 2, col: 1 }
encodeCellRef(2, 1);  // "B3"
```

### Validation & Repair

WASM-accelerated OOXML compliance checking and auto-repair:

```typescript
const wb = new Workbook();
const ws = wb.addSheet('Sheet1');
ws.cell('A1').value = 'Hello';
ws.cell('A1').styleIndex = 999; // dangling!

// Validate — returns a structured report
const report = wb.validate();
console.log(report.isValid);       // false
console.log(report.errorCount);    // 1
console.log(report.issues[0]);
// {
//   severity: 'error',
//   category: 'styleIndex',
//   message: 'Cell A1 styleIndex=999 exceeds cellXfs count (1)',
//   location: 'Sheet1!A1',
//   suggestion: 'Clamp styleIndex to 0',
//   autoFixable: true
// }

// Repair — auto-fixes all repairable issues
const { workbook, report: postReport, repairCount } = wb.repair();
console.log(repairCount);          // 2 (style index + missing theme)
console.log(postReport.isValid);   // true
```

Detects: dangling style/font/fill/border indices, overlapping merges, invalid
cell refs, duplicate sheet names, bad metadata dates, missing required styles,
missing theme colors, unsorted rows, SharedString cells without values.

### Barcode & QR Code Generation

Generate barcodes and embed them as images — pure TypeScript, zero dependencies:

```typescript
import { Workbook, generateBarcode, encodeQR, renderBarcodePNG } from 'modern-xlsx';

const wb = new Workbook();
wb.addSheet('Labels');

// High-level: generate + embed in one call
wb.addBarcode('Labels',
  { fromCol: 0, fromRow: 0, toCol: 4, toRow: 4 },
  'https://example.com',
  { type: 'qr', ecLevel: 'M' },
);

// Or use the low-level pipeline:
const matrix = encodeQR('Hello', { ecLevel: 'H' });
const png = renderBarcodePNG(matrix, { moduleSize: 6, showText: true, textValue: 'Hello' });
wb.addImage('Labels', { fromCol: 5, fromRow: 0, toCol: 9, toRow: 4 }, png);
```

**Supported formats:** QR Code, Code 128, EAN-13, UPC-A, Code 39, PDF417, Data Matrix, ITF-14, GS1-128.

See the [Barcode Guide](https://github.com/ABCrimson/modern-xlsx/blob/main/docs/guide/barcodes.md) for format comparison and usage details.

### Rich Text

```typescript
import { RichTextBuilder } from 'modern-xlsx';

const runs = new RichTextBuilder()
  .text('Normal ')
  .bold('Bold ')
  .italic('Italic ')
  .styled('Custom', { color: 'FF0000', fontSize: 14 })
  .build();
```

### Browser Usage (ESM)

```html
<script type="module">
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

### Browser Usage (CDN / IIFE)

Single `<script>` tag — no bundler required:

```html
<script src="https://cdn.jsdelivr.net/npm/modern-xlsx@0.5.0/dist/modern-xlsx.min.js"></script>
<script>
  (async () => {
    await ModernXlsx.initWasm();
    const wb = new ModernXlsx.Workbook();
    wb.addSheet('Sheet1').cell('A1').value = 'Hello!';
    const blob = ModernXlsx.writeBlob(wb);
    // trigger download...
  })();
</script>
```

Also available via unpkg: `https://unpkg.com/modern-xlsx@0.5.0/dist/modern-xlsx.min.js`

### Web Worker (Off-Thread)

Keep the main thread responsive by running XLSX operations in a Web Worker:

```typescript
import { createXlsxWorker } from 'modern-xlsx';

const worker = createXlsxWorker({
  workerUrl: '/modern-xlsx.worker.js',
  wasmUrl: '/modern-xlsx.wasm',  // optional
});

const data = await worker.readBuffer(xlsxBytes);
const output = await worker.writeBuffer(data);
worker.terminate();
```

### Lazy Initialization

Auto-initialize WASM on first use:

```typescript
import { ensureReady, Workbook } from 'modern-xlsx';

await ensureReady();  // no-op if already initialized
const wb = new Workbook();
```

### Custom WASM URL

For environments where auto-detection doesn't work:

```typescript
import { initWasm } from 'modern-xlsx';

// Custom URL
await initWasm('https://my-cdn.com/modern-xlsx.wasm');

// From fetch Response
const res = await fetch('/wasm/modern-xlsx.wasm');
await initWasm(res);
```

## Types

All types are exported and include full TypeScript definitions:

```typescript
import type {
  WorkbookData, WorksheetData, SheetData, CellData, RowData,
  StylesData, FontData, FillData, BorderData, AlignmentData, ProtectionData,
  CellXfData, NumFmt, ColumnInfo, FrozenPane,
  AutoFilterData, FilterColumnData, HyperlinkData,
  DataValidationData, ConditionalFormattingData,
  PageSetupData, SheetProtectionData,
  DocPropertiesData, ThemeColorsData, WorkbookViewData,
  RichTextRun, DefinedNameData, CalcChainEntryData,
  // Validation & compliance
  ValidationReport, ValidationIssue, Severity, IssueCategory, RepairResult,
  // Barcode & image embedding
  BarcodeMatrix, DrawBarcodeOptions, RenderOptions, ImageAnchor,
  // Tables & print layout
  TableDefinitionData, TableColumnData, TableStyleInfoData,
  HeaderFooterData, OutlinePropertiesData,
} from 'modern-xlsx';
```

## License

[MIT](./LICENSE)
