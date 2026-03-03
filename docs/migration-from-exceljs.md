# Migrating from ExcelJS

<div align="center">

**Equivalent patterns for common ExcelJS operations**

</div>

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

---

## Auto Filter

```diff
- ws.autoFilter = 'A1:D10';
+ ws.autoFilter = 'A1:D10';                  // string form (same API)
+ ws.autoFilter = { range: 'A1:D10' };       // or object form
```

---

## Data Validation

```diff
- ws.getCell('A1').dataValidation = {
-   type: 'list',
-   formulae: ['"Yes,No,Maybe"'],
- };
+ ws.addValidation('A1', {
+   validationType: 'list',
+   formula1: '"Yes,No,Maybe"',
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

## Key Differences

| Feature | ExcelJS | modern-xlsx |
|---|---|---|
| Runtime | Pure JS (streaming) | Rust WASM |
| Styles | Per-cell objects | Shared style table + index |
| Cell access | `ws.getCell('A1')` | `ws.cell('A1')` |
| Merge cells | `ws.mergeCells(range)` | `ws.addMergeCell(range)` |
| Sheet creation | `wb.addWorksheet(name)` | `wb.addSheet(name)` |
| Read/write | Streaming in JS | Bulk WASM processing |
| Dependencies | 14+ | 0 (WASM bundled) |
| Barcode generation | No | **Yes (9 formats)** |
