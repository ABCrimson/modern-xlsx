# Table Layout Engine

Generate styled XLSX tables from declarative options — no manual cell coordinate math.

## Quick Start

```typescript
import { initWasm, Workbook, drawTable } from 'modern-xlsx';

await initWasm();

const wb = new Workbook();
const ws = wb.addSheet('Report');

const result = drawTable(wb, ws, {
  headers: ['Name', 'Department', 'Salary'],
  rows: [
    ['Alice', 'Engineering', 95000],
    ['Bob', 'Marketing', 72000],
    ['Carol', 'Engineering', 105000],
  ],
  columnWidths: [20, 18, 12],
});

console.log(result.range); // "A1:C4"
```

## API Reference

### `drawTable(wb, ws, opts): TableResult`

Draws a styled table on a worksheet.

**Parameters:**

| Param | Type | Description |
|-------|------|-------------|
| `wb` | `Workbook` | Workbook instance (needed for styles) |
| `ws` | `Worksheet` | Target worksheet |
| `opts` | `DrawTableOptions` | Table configuration |

**Returns:** `TableResult` with layout metadata:

```typescript
interface TableResult {
  range: string;       // "A1:D10" — full table range
  rowCount: number;    // Total rows (header + data)
  colCount: number;    // Number of columns
  firstDataRow: number; // 0-based first data row index
  lastDataRow: number;  // 0-based last data row index
}
```

### `drawTableFromData(wb, ws, data, opts?): TableResult`

Creates a table from a JSON array, auto-extracting headers from object keys.

```typescript
const data = [
  { name: 'Alice', age: 30, city: 'NYC' },
  { name: 'Bob', age: 25, city: 'LA' },
];

drawTableFromData(wb, ws, data, {
  headerMap: { name: 'Full Name', age: 'Age', city: 'City' },
  autoWidth: true,
});
```

## Options

### Layout

| Option | Type | Default | Description |
|--------|------|---------|-------------|
| `headers` | `string[]` | required | Header labels |
| `rows` | `(string\|number\|boolean\|null)[][]` | required | Data rows |
| `origin` | `string` | `"A1"` | A1-style origin cell |
| `columns` | `TableColumn[]` | — | Per-column config |
| `columnWidths` | `number[]` | — | Fixed column widths |
| `autoWidth` | `boolean` | `false` | Auto-calculate widths |
| `freezeHeader` | `boolean` | `false` | Freeze the header row |
| `autoFilter` | `boolean` | `false` | Add filter dropdowns |
| `wrapText` | `boolean` | `false` | Enable text wrapping |

### Styling

| Option | Type | Default | Description |
|--------|------|---------|-------------|
| `headerFont` | `Partial<FontData>` | bold white | Header font |
| `headerColor` | `string` | `'4472C4'` | Header background (hex) |
| `bodyFont` | `Partial<FontData>` | — | Body font override |
| `borderStyle` | `BorderStyle\|null` | `'thin'` | Border style (`null` = none) |
| `borderColor` | `string` | `'000000'` | Border color (hex) |
| `alternateRowColor` | `string\|null` | — | Zebra stripe color (hex) |
| `headerAlign` | `'left'\|'center'\|'right'` | `'center'` | Header alignment |
| `bodyAlign` | `'left'\|'center'\|'right'` | — | Body alignment |
| `verticalAlign` | `'top'\|'center'\|'bottom'` | — | Vertical alignment |

### Per-Column Configuration

```typescript
interface TableColumn {
  header?: string;        // Header label override
  width?: number;         // Fixed width (character units)
  align?: 'left' | 'center' | 'right';
  numberFormat?: string;  // e.g. '#,##0.00'
}
```

## Recipes

### Invoice Table

```typescript
drawTable(wb, ws, {
  headers: ['Item', 'Qty', 'Unit Price', 'Total'],
  rows: [
    ['Widget', 10, 25.5, 255],
    ['Gadget', 5, 42.0, 210],
    ['Doohickey', 2, 99.99, 199.98],
  ],
  columnWidths: [20, 8, 12, 12],
  headerColor: '2F5496',
  alternateRowColor: 'D6E4F0',
  freezeHeader: true,
  columns: [
    { align: 'left' },
    { align: 'center' },
    { align: 'right', numberFormat: '$#,##0.00' },
    { align: 'right', numberFormat: '$#,##0.00' },
  ],
});
```

### Zebra Striping

```typescript
drawTable(wb, ws, {
  headers: ['ID', 'Name', 'Status'],
  rows: data,
  alternateRowColor: 'F2F2F2',
});
```

### Per-Cell Styling

Override individual cell styles using `"row,col"` keys (0-based, relative to data area):

```typescript
drawTable(wb, ws, {
  headers: ['Name', 'Score', 'Grade'],
  rows: [
    ['Alice', 95, 'A'],
    ['Bob', 42, 'F'],
  ],
  cellStyles: {
    '1,1': { font: { color: 'FF0000', bold: true } },
    '1,2': { fill: { pattern: 'solid', fgColor: 'FFCCCC' } },
  },
});
```

### Merge Cells

Merge cells in the data area (0-based row/col, relative to first data row):

```typescript
drawTable(wb, ws, {
  headers: ['Category', 'Product', 'Price'],
  rows: [
    ['Electronics', 'Phone', 999],
    ['', 'Laptop', 1299],
    ['Clothing', 'Shirt', 49],
  ],
  merges: [
    { row: 0, col: 0, rowSpan: 2 }, // Merge "Electronics" across 2 rows
  ],
});
```

### Nested Tables

Use `TableResult` metadata to position sequential tables:

```typescript
const result1 = drawTable(wb, ws, {
  headers: ['Q1 Summary'],
  rows: [['Revenue: $1M'], ['Profit: $200K']],
});

// Place second table below the first, with a gap row
const nextRow = result1.lastDataRow + 2;
drawTable(wb, ws, {
  headers: ['Q2 Summary'],
  rows: [['Revenue: $1.2M'], ['Profit: $250K']],
  origin: `A${nextRow + 1}`,
});
```

### Side-by-Side Tables

```typescript
drawTable(wb, ws, {
  headers: ['Team A'],
  rows: [['Alice'], ['Bob']],
  origin: 'A1',
});

drawTable(wb, ws, {
  headers: ['Team B'],
  rows: [['Carol'], ['Dave']],
  origin: 'D1', // Start 3 columns over
});
```

### Auto-Width with CJK Support

Auto-width calculation handles CJK characters as double-width:

```typescript
drawTable(wb, ws, {
  headers: ['Name', 'Description'],
  rows: [
    ['Widget', 'Standard component'],
    ['部品', '日本語の説明'],
  ],
  autoWidth: true,
});
```
