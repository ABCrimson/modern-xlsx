# Export Sales Report

Creates a professional sales report XLSX file using modern-xlsx.

## What It Builds

- **Title** -- merged across all columns, 20pt font
- **Data table** -- 6 regions x 7 columns (Q1-Q4, Total, Growth)
- **Zebra striping** -- alternating blue-tinted rows
- **Number formats** -- `$#,##0` for currency, `0.0%` for percentages
- **Formulas** -- SUM totals row, growth percentage calculations
- **Conditional styling** -- negative growth highlighted in red
- **Auto-filter** -- dropdown filters on every header
- **Frozen header** -- header row stays visible while scrolling
- **Column widths** -- sized to fit content

## Usage

```bash
npm install
node index.mjs
```

This produces `sales-report.xlsx` in the current directory. Open it in
Excel, Google Sheets, or LibreOffice to see the full styling.

## Key APIs Used

| API | Purpose |
|-----|---------|
| `Workbook` / `addSheet()` | Create workbook structure |
| `drawTable()` | High-level table layout with automatic styling |
| `StyleBuilder` | Build custom styles (font, fill, border, number format) |
| `ws.cell(ref)` | Get/set cell values and formulas |
| `ws.frozenPane` | Freeze rows/columns |
| `ws.autoFilter` | Add filter dropdowns |
| `ws.addMergeCell()` | Merge cell ranges |
| `wb.toBuffer()` | Serialize to XLSX bytes |
