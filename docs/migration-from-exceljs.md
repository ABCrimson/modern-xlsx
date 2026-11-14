# Migrating from ExcelJS

<div align="center">

**Equivalent patterns for common ExcelJS operations**

</div>

Statements about ExcelJS versions, dependencies and feature availability carry an "as of" date and a source link. They were read on 2026-11-11 and can have changed since. The API names on the left of each diff are shown for mapping only.

---

## Installation

```diff
- npm install exceljs
+ npm install modern-xlsx
```

---

## Initialization

```typescript
import { initWasm } from 'modern-xlsx';

await initWasm(); // Call once at startup
```

---

## Creating a Workbook

```diff
- import ExcelJS from 'exceljs';
- const wb = new ExcelJS.Workbook();
+ import { initWasm, Workbook } from 'modern-xlsx';
+ await initWasm();
+ const wb = new Workbook();
```

---

## Adding Sheets

```diff
- const ws = wb.addWorksheet('Sales');
+ const ws = wb.addSheet('Sales');
```

---

## Cell Values

```diff
- ws.getCell('A1').value = 'Hello';
- ws.getCell('A2').value = 42;
- ws.getCell('A3').value = true;
+ ws.cell('A1').value = 'Hello';
+ ws.cell('A2').value = 42;
+ ws.cell('A3').value = true;
```

### Formulas

```diff
- ws.getCell('A4').value = { formula: 'SUM(A1:A3)', result: 42 };
+ ws.cell('A4').formula = 'SUM(A1:A3)';
```

---

## Styling

```diff
- ws.getCell('A1').font = { bold: true, color: { argb: 'FFFF0000' } };
- ws.getCell('A1').fill = {
-   type: 'pattern',
-   pattern: 'solid',
-   fgColor: { argb: 'FFFFFF00' },
- };
+ const style = wb.createStyle()
+   .font({ bold: true, color: 'FF0000' })
+   .fill({ pattern: 'solid', fgColor: 'FFFF00' })
+   .build(wb.styles);
+ ws.cell('A1').styleIndex = style;
```

Colors are hex strings. The style builder accepts 6-digit RGB (`'FF0000'`) and 8-digit ARGB (`'FFFF0000'`), with or without a leading `#`, so the ExcelJS `argb` value can be passed through unchanged. Both are written as uppercase ARGB, with an `FF` alpha added to an RGB value, and read back in that form.

---

## Merge Cells

```diff
- ws.mergeCells('A1:C1');
+ ws.addMergeCell('A1:C1');
```

---

## Column Widths

```diff
- ws.getColumn('A').width = 20;
+ ws.setColumnWidth(1, 20); // 1-based column number
```

---

## Row Heights

```diff
- ws.getRow(1).height = 30;
+ ws.setRowHeight(1, 30);
```

---

## Frozen Panes

```diff
- ws.views = [{ state: 'frozen', xSplit: 0, ySplit: 1 }];
+ ws.frozenPane = { rows: 1, cols: 0 };
```

`rows` is the number of rows frozen from the top and `cols` the number of columns frozen from the left. Assign `null` to remove the pane.

---

## Auto Filter

```diff
- ws.autoFilter = 'A1:D10';
+ ws.autoFilter = 'A1:D10';                  // string form (same API)
+ ws.autoFilter = { range: 'A1:D10' };       // or object form
```

The setter accepts a string, an `AutoFilterData` object (`range`, plus optional `filterColumns`) or `null` to remove the filter.

---

## Data Validation

```diff
- ws.getCell('A1').dataValidation = {
-   type: 'list',
-   formulae: ['"Yes,No,Maybe"'],
- };
+ ws.addValidation('A1', {
+   validationType: 'list',
+   operator: null,
+   formula1: '"Yes,No,Maybe"',
+   formula2: null,
+   allowBlank: true,
+   showErrorMessage: null,
+   errorTitle: null,
+   errorMessage: null,
+ });
```

---

## Reading Files

```diff
- const wb = new ExcelJS.Workbook();
- await wb.xlsx.readFile('data.xlsx');
+ import { readFile } from 'modern-xlsx';
+ const wb = await readFile('data.xlsx');
```

### From buffer

```diff
- const wb = new ExcelJS.Workbook();
- await wb.xlsx.load(buffer);
+ import { readBuffer } from 'modern-xlsx';
+ const wb = await readBuffer(new Uint8Array(buffer));
```

---

## Writing Files

```diff
- await wb.xlsx.writeFile('output.xlsx');
+ await wb.toFile('output.xlsx');
```

### To buffer

```diff
- const buffer = await wb.xlsx.writeBuffer();
+ const buffer = await wb.toBuffer();
```

---

## Streaming writes

`StreamingXlsxWriter` replaces the ExcelJS `stream.xlsx.WorkbookWriter` pattern of `addRow(...).commit()`. Create it with `StreamingXlsxWriter.create()` after `initWasm()`, open a sheet with `startSheet(name)`, write rows with `writeRow(cells)`, and call `finish()`, which returns the whole file as a `Uint8Array`.

```diff
- const wb = new ExcelJS.stream.xlsx.WorkbookWriter({ filename: 'big.xlsx' });
- const ws = wb.addWorksheet('Data');
- ws.addRow(['id', 'name']).commit();
- await ws.commit();
- await wb.commit();
+ import { initWasm, StreamingXlsxWriter } from 'modern-xlsx';
+ import { writeFile } from 'node:fs/promises';
+ await initWasm();
+ const writer = StreamingXlsxWriter.create();
+ writer.startSheet('Data');
+ writer.writeRow([
+   { value: 'id', cellType: 'sharedString' },
+   { value: 'name', cellType: 'sharedString' },
+ ]);
+ writer.writeRow([
+   { value: '1' },                              // cellType defaults to 'number'
+   { value: 'Alice', cellType: 'sharedString' },
+ ]);
+ await writeFile('big.xlsx', writer.finish());
```

Each cell is `{ value?, cellType?, style? }`. `value` is a string; omit it or pass `null` to leave the cell empty (a `stub` cell with a `style` is still written). `cellType` defaults to `'number'`, so give text cells `cellType: 'sharedString'`. `finish()` consumes the writer.

Rows are compressed into the ZIP as they are written, but the compressed archive stays in WASM memory until `finish()` returns it, so the output is buffered in memory and not written to a file or stream as it goes.

For a normal workbook, `ws.addRow(values)` maps to `sheetAddAoa`. Without an `origin` option it writes below the existing data, so repeated calls append rows:

```diff
- ws.addRow(['Alice', 30]);
+ import { sheetAddAoa } from 'modern-xlsx';
+ sheetAddAoa(ws, [['Alice', 30]]);
```

Pass `{ origin: 'A10' }` to start at a given cell instead.

---

## Key Differences

| Feature              | ExcelJS                                                                                                    | modern-xlsx                                                                                                                              |
| -------------------- | ---------------------------------------------------------------------------------------------------------- | ---------------------------------------------------------------------------------------------------------------------------------------- |
| Styles               | Per-cell objects                                                                                           | Shared style table + index                                                                                                               |
| Cell access          | `ws.getCell('A1')`                                                                                         | `ws.cell('A1')`                                                                                                                          |
| Merge cells          | `ws.mergeCells(range)`                                                                                     | `ws.addMergeCell(range)`                                                                                                                 |
| Sheet creation       | `wb.addWorksheet(name)`                                                                                    | `wb.addSheet(name)`                                                                                                                      |
| Streaming write      | `stream.xlsx.WorkbookWriter` (as of 2026-11-11, [README 4.4.0](https://unpkg.com/exceljs@4.4.0/README.md)) | `StreamingXlsxWriter` writes row by row but buffers its compressed output in memory until `finish()`                                     |
| Streaming read       | `stream.xlsx.WorkbookReader` (as of 2026-11-11, [README 4.4.0](https://unpkg.com/exceljs@4.4.0/README.md)) | None. `readBuffer` and `readFile` read the whole file in one call                                                                        |
| Runtime dependencies | 9 (npm registry, exceljs 4.4.0, 2026-11-11, [source](https://registry.npmjs.org/exceljs))                  | 0 in `package.json`. The WASM binary ships as separate `.wasm` files in the package (`wasm/`, `wasm-lite/`, and `dist/modern-xlsx.wasm`) |

modern-xlsx's core is Rust compiled to WASM behind a TypeScript API. It also generates barcodes in 9 formats (`BarcodeType` in `barcode/common.ts`); PDF417 and Data Matrix are marked `@experimental` because their output is not standards-conformant and may not scan.
