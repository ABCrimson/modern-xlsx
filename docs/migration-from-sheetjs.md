# Migrating from SheetJS (xlsx)

<div align="center">

**Equivalent patterns for common SheetJS operations**

</div>

Statements about SheetJS versions, dependencies and feature availability carry an "as of" date and a source link. They were read on 2026-11-11 and can have changed since. The API names on the left of each diff are shown for mapping only.

---

## Installation

```diff
- npm install xlsx
+ npm install modern-xlsx
```

The npm package `xlsx` stops at 0.18.5 (2022-03-24); SheetJS publishes its newer Community Edition builds on its own CDN and calls that CDN the authoritative source (as of 2026-11-11: [npm registry](https://registry.npmjs.org/xlsx), [SheetJS install docs](https://docs.sheetjs.com/docs/getting-started/installation/nodejs)).

---

## Initialization

modern-xlsx uses a WASM core that must be initialized once before use.

```typescript
import { initWasm } from 'modern-xlsx';

// Call once at app startup (idempotent — safe to call multiple times)
await initWasm();
```

When the WASM bytes are already in memory, `initWasmSync(module)` initializes synchronously. It takes a `WebAssembly.Module` or raw bytes (`BufferSource`).

---

## Reading Files

### Node.js

```diff
- import XLSX from 'xlsx';
- const wb = XLSX.readFile('data.xlsx');
+ import { initWasm, readFile } from 'modern-xlsx';
+ await initWasm();
+ const wb = await readFile('data.xlsx');
```

### From a Buffer / Uint8Array

```diff
- const wb = XLSX.read(buffer, { type: 'buffer' });
+ import { readBuffer } from 'modern-xlsx';
+ const wb = await readBuffer(new Uint8Array(buffer));
```

### Bun and Deno

The repository's runtime tests ([`bun.test.ts`](../examples/runtimes/bun.test.ts) and [`deno.test.ts`](../examples/runtimes/deno.test.ts)) call `initWasm`, `new Workbook()`, `addSheet`, `StyleBuilder`, `toBuffer` and `readBuffer` on those runtimes. For file I/O there, read and write the bytes with the runtime's own file API and pass them to `readBuffer` or take them from `toBuffer`.

---

## Writing Files

### To a file

```diff
- XLSX.writeFile(wb, 'output.xlsx');
+ await wb.toFile('output.xlsx');
```

### To a buffer

```diff
- const buf = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
+ const buf = await wb.toBuffer();
```

### Browser download

```diff
- XLSX.writeFile(wb, 'download.xlsx');
+ import { writeBlob } from 'modern-xlsx';
+ const blob = writeBlob(wb);
+ const url = URL.createObjectURL(blob);
+ const a = document.createElement('a');
+ a.href = url;
+ a.download = 'download.xlsx';
+ a.click();
+ URL.revokeObjectURL(url);
```

---

## Sheet Access

```diff
- const ws = wb.Sheets[wb.SheetNames[0]];
+ const ws = wb.getSheetByIndex(0);
```

```diff
- const ws = wb.Sheets['Sheet1'];
+ const ws = wb.getSheet('Sheet1');
```

`getSheetByIndex` and `getSheet` return `undefined` when the sheet does not exist.

### Sheet names

```diff
- wb.SheetNames
+ wb.sheetNames
```

### Adding a sheet

`book_append_sheet` becomes `wb.addSheet(name)`, which creates the sheet inside the workbook, followed by `sheetAddAoa` to fill it. On an empty sheet `sheetAddAoa` starts at `A1`.

```diff
- const ws = XLSX.utils.aoa_to_sheet([['A', 'B'], [1, 2]]);
- XLSX.utils.book_append_sheet(wb, ws, 'Data');
+ import { sheetAddAoa } from 'modern-xlsx';
+ const ws = wb.addSheet('Data');
+ sheetAddAoa(ws, [['A', 'B'], [1, 2]]);
```

Cell by cell works too:

```diff
+ const ws = wb.addSheet('Data');
+ ws.cell('A1').value = 'A';
+ ws.cell('B1').value = 'B';
+ ws.cell('A2').value = 1;
+ ws.cell('B2').value = 2;
```

---

## Cell Access

```diff
- const cell = ws['A1'];
- const value = cell ? cell.v : undefined;
+ const value = ws.cell('A1').value;
```

### Setting values

```diff
- ws['A1'] = { t: 's', v: 'Hello' };
- ws['A2'] = { t: 'n', v: 42 };
- ws['A3'] = { t: 'b', v: true };
+ ws.cell('A1').value = 'Hello';   // auto-detects string
+ ws.cell('A2').value = 42;        // auto-detects number
+ ws.cell('A3').value = true;      // auto-detects boolean
```

### Formulas

```diff
- ws['A4'] = { t: 'n', f: 'SUM(A1:A3)' };
+ ws.cell('A4').formula = 'SUM(A1:A3)';
```

---

## Sheet Utilities

### Sheet to JSON

```diff
- const data = XLSX.utils.sheet_to_json(ws);
+ import { sheetToJson } from 'modern-xlsx';
+ const data = sheetToJson(ws);
```

### Array of arrays to sheet

`aoa_to_sheet` plus `book_append_sheet` maps to `wb.addSheet` plus `sheetAddAoa`:

```diff
- const ws = XLSX.utils.aoa_to_sheet(aoa);
- XLSX.utils.book_append_sheet(wb, ws, 'Data');
+ import { sheetAddAoa } from 'modern-xlsx';
+ const ws = wb.addSheet('Data');
+ sheetAddAoa(ws, aoa);
```

`sheetAddAoa(ws, aoa, { origin: 'C3' })` starts at a given cell. Without `origin` it appends below the existing data.

The package also exports `aoaToSheet(aoa)`. It returns a standalone `Worksheet` (named `Sheet1`) that no `Workbook` holds, and no API attaches it to a workbook, so it cannot be saved as an .xlsx file (`toFile` and `toBuffer` are `Workbook` methods). Use it only for in-memory work with the read-side helpers:

```typescript
import { aoaToSheet, sheetToCsv } from 'modern-xlsx';

const detached = aoaToSheet([
  ['A', 'B'],
  [1, 2],
]);
const csv = sheetToCsv(detached);
```

### JSON to sheet

`json_to_sheet` maps to `sheetAddJson`. The object keys of the first row become the header row.

```diff
- const ws = XLSX.utils.json_to_sheet(jsonData);
- XLSX.utils.book_append_sheet(wb, ws, 'Data');
+ import { sheetAddJson } from 'modern-xlsx';
+ const ws = wb.addSheet('Data');
+ sheetAddJson(ws, jsonData);
```

`jsonToSheet(jsonData)` is the standalone form and has the same limit as `aoaToSheet`: the returned `Worksheet` belongs to no workbook.

### Sheet to CSV

```diff
- const csv = XLSX.utils.sheet_to_csv(ws);
+ import { sheetToCsv } from 'modern-xlsx';
+ const csv = sheetToCsv(ws);
```

### Sheet to TXT (tab-separated)

```diff
- const txt = XLSX.utils.sheet_to_txt(ws);
+ import { sheetToTxt } from 'modern-xlsx';
+ const txt = sheetToTxt(ws);
```

### Sheet to formulae

```diff
- const formulae = XLSX.utils.sheet_to_formulae(ws);
+ import { sheetToFormulae } from 'modern-xlsx';
+ const formulae = sheetToFormulae(ws);
```

---

## Cell Reference Utilities

All indexes are 0-based. `decodeCellRef` returns a `CellAddress` with `row` and `col` fields, and `encodeCellRef` takes `(row, col)`.

```diff
- XLSX.utils.encode_row(0)          // "1"
+ import { encodeRow } from 'modern-xlsx';
+ encodeRow(0)                       // "1"

- XLSX.utils.decode_row("1")        // 0
+ import { decodeRow } from 'modern-xlsx';
+ decodeRow("1")                     // 0

- XLSX.utils.decode_cell("B3")      // { c: 1, r: 2 }
+ import { decodeCellRef } from 'modern-xlsx';
+ decodeCellRef("B3")               // { col: 1, row: 2 }

- XLSX.utils.encode_cell({ c: 1, r: 2 })  // "B3"
+ import { encodeCellRef } from 'modern-xlsx';
+ encodeCellRef(2, 1)               // "B3"

- XLSX.utils.encode_col(0)          // "A"
+ import { columnToLetter } from 'modern-xlsx';
+ columnToLetter(0)                  // "A"

- XLSX.utils.decode_col("A")        // 0
+ import { letterToColumn } from 'modern-xlsx';
+ letterToColumn("A")               // 0
```

---

## Number Formatting

`formatCell` takes the value first and the format second. `loadFormat` takes the format code first and the id second.

```diff
- const SSF = require('ssf');
- SSF.format(fmt, value)
+ import { formatCell } from 'modern-xlsx';
+ formatCell(value, fmt)

- SSF.load(fmt, id)
+ import { loadFormat } from 'modern-xlsx';
+ loadFormat(fmt, id)

- SSF.load_table(table)
+ import { loadFormatTable } from 'modern-xlsx';
+ loadFormatTable(table)  // table: Record<number, string>, id -> format code
```

---

## Excel Tables (ListObjects)

modern-xlsx can add, find and remove Excel tables:

```typescript
ws.addTable({
  id: 1,
  displayName: 'SalesData',
  ref: 'A1:C10',
  headerRowCount: 1,
  totalsRowCount: 0,
  totalsRowShown: false,
  columns: [
    { id: 1, name: 'Product' },
    { id: 2, name: 'Revenue' },
    { id: 3, name: 'Units' },
  ],
  styleInfo: {
    name: 'TableStyleMedium9',
    showFirstColumn: false,
    showLastColumn: false,
    showRowStripes: true,
    showColumnStripes: false,
  },
});

ws.tables; // all tables
ws.getTable('SalesData'); // find by display name
ws.removeTable('SalesData'); // remove
```

The SheetJS Community Edition docs list no table (ListObject) feature page, and the SheetJS Pro page lists table styling under Pro Basic (as of 2026-11-11: [CE feature list](https://docs.sheetjs.com/docs/csf/features/), [SheetJS Pro](https://sheetjs.com/pro/)).

---

## Headers, Footers & Print Layout

```typescript
import { HeaderFooterBuilder } from 'modern-xlsx';

ws.headerFooter = {
  oddHeader: new HeaderFooterBuilder()
    .center(HeaderFooterBuilder.bold('Report'))
    .right(`Page ${HeaderFooterBuilder.pageNumber()}`)
    .build(),
};

// Row/column grouping
ws.groupRows(2, 10);
ws.collapseRows(2, 10);
ws.groupColumns(1, 3);

// Print titles & areas (0-based sheet index; value is a defined-name reference)
wb.setPrintTitles(0, 'Sheet1!$1:$1');
wb.setPrintArea(0, 'Sheet1!$A$1:$G$50');
```

Pass `null` as the value to clear a print title or print area.

---

## Styles

modern-xlsx writes cell styles with a fluent style builder:

```typescript
const wb = new Workbook();
const ws = wb.addSheet('Styled');

const boldRed = wb
  .createStyle()
  .font({ bold: true, color: 'FF0000' })
  .fill({ pattern: 'solid', fgColor: 'FFFF00' })
  .border({ bottom: { style: 'thin', color: '000000' } })
  .build(wb.styles);

ws.cell('A1').value = 'Bold Red on Yellow';
ws.cell('A1').styleIndex = boldRed;
```

SheetJS Community Edition exports row and column styling metadata when the `cellStyles` write option is true, and lists cell and text styling under SheetJS Pro (as of 2026-11-11: [SheetJS write options](https://docs.sheetjs.com/docs/api/write-options)).

---

## Key Differences

| Feature                                                                                                                                          | SheetJS                                                                                                                                                                                                            | modern-xlsx                                                                                                                              |
| ------------------------------------------------------------------------------------------------------------------------------------------------ | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ | ---------------------------------------------------------------------------------------------------------------------------------------- |
| Cell access                                                                                                                                      | `ws['A1']`                                                                                                                                                                                                         | `ws.cell('A1')`                                                                                                                          |
| Runtime dependencies (as of 2026-11-11, [npm](https://registry.npmjs.org/xlsx), [CDN](https://cdn.sheetjs.com/xlsx-latest/package/package.json)) | npm 0.18.5: 7. CDN 0.20.3: 0                                                                                                                                                                                       | 0 in `package.json`. The WASM binary ships as separate `.wasm` files in the package (`wasm/`, `wasm-lite/`, and `dist/modern-xlsx.wasm`) |
| Cell styling                                                                                                                                     | Row and column style metadata with `cellStyles`; cell and text styling listed under SheetJS Pro (as of 2026-11-11, [write options](https://docs.sheetjs.com/docs/api/write-options))                               | Included (`createStyle()`)                                                                                                               |
| Excel Tables                                                                                                                                     | No CE feature page; Pro Basic lists table styling (as of 2026-11-11, [CE features](https://docs.sheetjs.com/docs/csf/features/), [Pro](https://sheetjs.com/pro/))                                                  | `addTable`, `getTable`, `removeTable`                                                                                                    |
| Available on npm                                                                                                                                 | `xlsx` stops at 0.18.5 (2022-03-24); CE 0.20.3 is on the SheetJS CDN (as of 2026-11-11, [npm](https://registry.npmjs.org/xlsx), [SheetJS docs](https://docs.sheetjs.com/docs/getting-started/installation/nodejs)) | `modern-xlsx`                                                                                                                            |

On the modern-xlsx side:

- The core is Rust compiled to WASM, behind a TypeScript API.
- The `modern-xlsx` and `modern-xlsx/lite` entries are ES modules, and there is no CommonJS build. `modern-xlsx/browser` is an IIFE bundle for a `<script>` tag that sets the global `ModernXlsx`. Files are exchanged as `Uint8Array`.
- `dateToSerial` accepts a JS `Date` or a `TemporalLike` (a native `Temporal.PlainDate` or `Temporal.PlainDateTime`, or any object with numeric `year`, `month` and `day`).
- Initialization is `await initWasm()`, or `initWasmSync(module)` when the WASM bytes are already in memory.
- It also generates barcodes (`generateBarcode`, `renderBarcodePNG`) and writes headers and footers, row and column grouping, and print titles and areas, as shown above.
