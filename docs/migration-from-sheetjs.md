# Migrating from SheetJS (xlsx)

<div align="center">

**Drop-in patterns for common SheetJS operations**

</div>

---

## Installation

```diff
- npm install xlsx
+ npm install modern-xlsx
```

---

## Initialization

modern-xlsx uses a WASM core that must be initialized once before use.

```typescript
import { initWasm } from 'modern-xlsx';

// Call once at app startup (idempotent — safe to call multiple times)
await initWasm();
```

---

## Reading Files

### Node.js / Bun / Deno

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
+ const wb = await readBuffer(new Uint8Array(buffer));
```

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
// or
+ const ws = wb.getSheet('Sheet1');
```

### Sheet names

```diff
- wb.SheetNames
+ wb.sheetNames
```

### Adding a sheet

```diff
- const ws = XLSX.utils.aoa_to_sheet([['A', 'B'], [1, 2]]);
- XLSX.utils.book_append_sheet(wb, ws, 'Data');
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

```diff
- const ws = XLSX.utils.aoa_to_sheet(aoa);
+ import { aoaToSheet } from 'modern-xlsx';
+ const rows = aoaToSheet(aoa);
```

### JSON to sheet

```diff
- const ws = XLSX.utils.json_to_sheet(jsonData);
+ import { jsonToSheet } from 'modern-xlsx';
+ const rows = jsonToSheet(jsonData);
```

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

```diff
- XLSX.utils.encode_row(0)          // "1"
+ import { encodeRow } from 'modern-xlsx';
+ encodeRow(0)                       // "1"

- XLSX.utils.decode_row("1")        // 0
+ import { decodeRow } from 'modern-xlsx';
+ decodeRow("1")                     // 0

- XLSX.utils.decode_cell("A1")      // { c: 0, r: 0 }
+ import { decodeCellRef } from 'modern-xlsx';
+ decodeCellRef("A1")               // { row: 0, col: 0 }

- XLSX.utils.encode_cell({ c: 0, r: 0 })  // "A1"
+ import { encodeCellRef } from 'modern-xlsx';
+ encodeCellRef(0, 0)               // "A1"

- XLSX.utils.encode_col(0)          // "A"
+ import { columnToLetter } from 'modern-xlsx';
+ columnToLetter(0)                  // "A"

- XLSX.utils.decode_col("A")        // 0
+ import { letterToColumn } from 'modern-xlsx';
+ letterToColumn("A")               // 0
```

---

## Number Formatting

```diff
- const SSF = require('ssf');
- SSF.format(fmt, value)
+ import { formatCell } from 'modern-xlsx';
+ formatCell(value, fmt)

- SSF.load(fmt, id)
+ import { loadFormat } from 'modern-xlsx';
+ loadFormat(id, fmt)

- SSF.load_table(table)
+ import { loadFormatTable } from 'modern-xlsx';
+ loadFormatTable(table)
```

---

## Excel Tables (ListObjects)

SheetJS Pro only. modern-xlsx includes full table CRUD for free:

```typescript
ws.addTable({
  name: 'SalesData',
  ref: 'A1:C10',
  columns: [{ name: 'Product' }, { name: 'Revenue' }, { name: 'Units' }],
  style: { name: 'TableStyleMedium9', showRowStripes: true },
});

ws.tables;                    // all tables
ws.getTable('SalesData');     // find by name
ws.removeTable('SalesData');  // remove
```

---

## Headers, Footers & Print Layout

SheetJS Pro only. modern-xlsx includes all print features for free:

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

// Print titles & areas
wb.setPrintTitles('Sheet1', { rows: { start: 1, end: 1 } });
wb.setPrintArea('Sheet1', 'A1:G50');
```

---

## Styles

SheetJS community edition does not support styles. modern-xlsx includes a fluent style builder:

```typescript
const wb = new Workbook();
const ws = wb.addSheet('Styled');

const boldRed = wb.createStyle()
  .font({ bold: true, color: 'FF0000' })
  .fill({ pattern: 'solid', fgColor: 'FFFF00' })
  .border({ bottom: { style: 'thin', color: '000000' } })
  .build(wb.styles);

ws.cell('A1').value = 'Bold Red on Yellow';
ws.cell('A1').styleIndex = boldRed;
```

---

## Key Differences

| Feature | SheetJS | modern-xlsx |
|---|---|---|
| Runtime | Pure JS | Rust WASM + TypeScript |
| Module format | CJS + ESM | ESM only |
| Styles | Pro only | Included |
| Buffer type | Node Buffer | `Uint8Array` |
| Date handling | JS `Date` | Temporal API |
| Initialization | Sync | Async (`initWasm()`) |
| Cell access | `ws['A1']` | `ws.cell('A1')` |
| Dependencies | 0 | 0 (WASM bundled) |
| Barcode generation | N/A | **Yes** |
| Excel Tables | Pro only | **Included** |
| Headers & footers | Pro only | **Included** |
| Row/column grouping | Pro only | **Included** |
| Print titles & areas | Pro only | **Included** |
| Available on npm | Yes | **Yes** |
