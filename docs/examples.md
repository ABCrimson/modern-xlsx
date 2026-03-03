# Usage Examples

<div align="center">

**Common patterns and recipes for modern-xlsx**

</div>

---

## Quick Start

```typescript
import { initWasm, Workbook, readBuffer } from 'modern-xlsx';

// Initialize WASM (once per process)
await initWasm();

// Create a workbook
const wb = new Workbook();
const ws = wb.addSheet('Hello');
ws.cell('A1').value = 'Hello, World!';
ws.cell('A2').value = 42;

// Write to file
await wb.toFile('hello.xlsx');
```

---

## Reading and Modifying

```typescript
import { initWasm, readFile } from 'modern-xlsx';

await initWasm();

const wb = await readFile('input.xlsx');
const ws = wb.getSheet('Sheet1');

if (ws) {
  // Read values
  console.log(ws.cell('A1').value);

  // Modify values
  ws.cell('A1').value = 'Updated!';

  // Save changes
  await wb.toFile('output.xlsx');
}
```

---

## Styling Cells

```typescript
const wb = new Workbook();
const ws = wb.addSheet('Styled');

// Create a header style
const headerStyle = wb.createStyle()
  .font({ name: 'Arial', size: 14, bold: true, color: 'FFFFFF' })
  .fill({ pattern: 'solid', fgColor: '4472C4' })
  .alignment({ horizontal: 'center', vertical: 'center' })
  .border({
    bottom: { style: 'medium', color: '000000' },
  })
  .build(wb.styles);

// Apply to cells
const headers = ['Name', 'Age', 'City'];
headers.forEach((h, i) => {
  const cell = ws.cell(`${String.fromCharCode(65 + i)}1`);
  cell.value = h;
  cell.styleIndex = headerStyle;
});
```

---

## Formulas

```typescript
const ws = wb.addSheet('Formulas');

// Data
ws.cell('A1').value = 100;
ws.cell('A2').value = 200;
ws.cell('A3').value = 300;

// Formulas
ws.cell('A4').formula = 'SUM(A1:A3)';
ws.cell('B1').formula = 'AVERAGE(A1:A3)';
ws.cell('C1').formula = 'MAX(A1:A3)';
```

---

## Merged Cells

```typescript
const ws = wb.addSheet('Merges');

ws.cell('A1').value = 'This spans three columns';
ws.addMergeCell('A1:C1');

ws.cell('A2').value = 'This spans two rows';
ws.addMergeCell('A2:A3');
```

---

## Column Widths and Row Heights

```typescript
const ws = wb.addSheet('Layout');

// Set column widths (1-based column number)
ws.setColumnWidth(1, 30); // Column A = 30 units
ws.setColumnWidth(2, 15); // Column B = 15 units

// Set row height (1-based row number, in points)
ws.setRowHeight(1, 40);

// Hide a row
ws.setRowHidden(5, true);
```

---

## Frozen Panes

```typescript
const ws = wb.addSheet('Frozen');

// Freeze top row
ws.frozenPane = { rows: 1, cols: 0 };

// Freeze first column
ws.frozenPane = { rows: 0, cols: 1 };

// Freeze both
ws.frozenPane = { rows: 1, cols: 1 };
```

---

## Data Validation

```typescript
const ws = wb.addSheet('Validation');

// Dropdown list
ws.addValidation('B2', {
  validationType: 'list',
  formula1: '"Yes,No,Maybe"',
  prompt: 'Select an option',
  promptTitle: 'Choice',
});

// Number range
ws.addValidation('C2', {
  validationType: 'whole',
  operator: 'between',
  formula1: '1',
  formula2: '100',
  errorTitle: 'Invalid',
  error: 'Enter a number between 1 and 100',
});
```

---

## Hyperlinks

```typescript
const ws = wb.addSheet('Links');

// External URL
ws.cell('A1').value = 'Visit Example';
ws.addHyperlink('A1', 'https://example.com', {
  display: 'Visit Example',
  tooltip: 'Opens example.com',
});

// Internal reference
ws.cell('A2').value = 'Go to Sheet2';
ws.addHyperlink('A2', 'Sheet2!A1', {
  display: 'Go to Sheet2',
});
```

---

## Comments

```typescript
const ws = wb.addSheet('Comments');

ws.cell('A1').value = 'Hover for comment';
ws.addComment('A1', 'Author Name', 'This is a comment on A1');
```

---

## Named Ranges

```typescript
const wb = new Workbook();
const ws = wb.addSheet('Data');

ws.cell('A1').value = 100;
ws.cell('A2').value = 200;

// Global named range
wb.addNamedRange('SalesTotal', 'Data!$A$1:$A$2');

// Access later
const range = wb.getNamedRange('SalesTotal');
console.log(range?.value); // "Data!$A$1:$A$2"
```

---

## Document Properties

```typescript
wb.docProperties = {
  title: 'Sales Report Q4',
  creator: 'Finance Team',
  description: 'Quarterly sales data',
  created: '2026-01-01T00:00:00Z',
  modified: '2026-03-01T00:00:00Z',
};
```

---

## Rich Text

```typescript
import { RichTextBuilder } from 'modern-xlsx';

const richText = new RichTextBuilder()
  .bold('Important: ')
  .text('This is normal text. ')
  .colored('Red text', 'FF0000')
  .styled('Custom', { bold: true, italic: true, fontSize: 14, fontName: 'Arial' })
  .build();

// Apply via raw data manipulation
const data = wb.toJSON();
if (!data.sharedStrings) {
  data.sharedStrings = { strings: [], richRuns: [] };
}
data.sharedStrings.strings.push(richText.plainText);
data.sharedStrings.richRuns.push(richText.runs);
```

---

## Sheet Conversion Utilities

### JSON to Sheet

```typescript
import { jsonToSheet, Workbook } from 'modern-xlsx';

const data = [
  { name: 'Alice', age: 30, city: 'NYC' },
  { name: 'Bob', age: 25, city: 'LA' },
];

const wb = new Workbook();
const ws = wb.addSheet('People');
const rows = jsonToSheet(data);
// rows is RowData[] that you can assign to sheet data
```

### Sheet to CSV

```typescript
import { sheetToCsv } from 'modern-xlsx';

const ws = wb.getSheet('Sheet1');
if (ws) {
  const csv = sheetToCsv(ws);
  console.log(csv);
}
```

---

## Auto Filter

```typescript
const ws = wb.addSheet('Filtered');

// Add data
ws.cell('A1').value = 'Name';
ws.cell('B1').value = 'Score';
ws.cell('A2').value = 'Alice';
ws.cell('B2').value = 95;
ws.cell('A3').value = 'Bob';
ws.cell('B3').value = 87;

// Enable auto filter on the range
ws.autoFilter = 'A1:B3';
```

---

## Page Setup

```typescript
ws.pageSetup = {
  orientation: 'landscape',
  paperSize: 1, // Letter
  fitToWidth: 1,
  fitToHeight: 0,
};

ws.pageMargins = {
  top: 0.75,
  bottom: 0.75,
  left: 0.7,
  right: 0.7,
  header: 0.3,
  footer: 0.3,
};
```

---

## Sheet Protection

```typescript
ws.sheetProtection = {
  sheet: true,
  selectLockedCells: false,
  selectUnlockedCells: false,
};
```

---

## Cell Reference Utilities

```typescript
import { columnToLetter, letterToColumn, decodeCellRef, encodeCellRef } from 'modern-xlsx';

columnToLetter(1);   // 'A'
columnToLetter(27);  // 'AA'
letterToColumn('A'); // 1
letterToColumn('AA'); // 27

decodeCellRef('B3');  // { row: 2, col: 1 } (0-based)
encodeCellRef(2, 1); // 'B3'
```

---

## Date Handling

```typescript
import { dateToSerial, serialToDate, isDateFormatCode } from 'modern-xlsx';

// Convert date to Excel serial number
dateToSerial({ year: 2026, month: 3, day: 1 }); // 46113

// Convert serial back to date components
serialToDate(46113); // { year: 2026, month: 3, day: 1 }

// Check if a format code is a date format
isDateFormatCode('yyyy-mm-dd'); // true
isDateFormatCode('#,##0.00');   // false
```

---

## Barcode & QR Code Generation

### Embed a QR code

```typescript
import { initWasm, Workbook } from 'modern-xlsx';

await initWasm();

const wb = new Workbook();
wb.addSheet('Labels');

wb.addBarcode('Labels',
  { fromCol: 1, fromRow: 0, toCol: 5, toRow: 4 },
  'https://example.com/product/12345',
  { type: 'qr', ecLevel: 'M' },
);

await wb.toFile('labels.xlsx');
```

### Embed a Code 128 barcode

```typescript
wb.addBarcode('Sheet1',
  { fromCol: 0, fromRow: 0, toCol: 4, toRow: 2 },
  'SKU-12345',
  { type: 'code128', showText: true },
);
```

### Low-level barcode pipeline

```typescript
import { encodeQR, renderBarcodePNG } from 'modern-xlsx';

const matrix = encodeQR('Hello World', { ecLevel: 'H' });
const png = renderBarcodePNG(matrix, {
  moduleSize: 6,
  quietZone: 4,
  showText: true,
  textValue: 'Hello World',
});

// Use PNG bytes however you want
wb.addImage('Sheet1', { fromCol: 0, fromRow: 0, toCol: 3, toRow: 3 }, png);
```

---

## Image Embedding

```typescript
import { Workbook, encodeQR, renderBarcodePNG } from 'modern-xlsx';

const wb = new Workbook();
wb.addSheet('Images');

// Any Uint8Array PNG bytes
const pngBytes = renderBarcodePNG(encodeQR('Test'), { moduleSize: 4 });

wb.addImage('Images',
  { fromCol: 0, fromRow: 0, toCol: 4, toRow: 4 },
  pngBytes,
  'png', // format: 'png' | 'jpeg' | 'gif'
);
```
