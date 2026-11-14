# Export Sales Report

Creates a professional sales report XLSX file using modern-xlsx.

## What It Builds

- **Title** -- merged across all columns, 20pt font
- **Data table** -- six regions across seven columns (Region, Q1-Q4, Total, Growth), written as values computed in JavaScript
- **Zebra striping** -- alternating blue-tinted rows
- **Number formats** -- `$#,##0` for currency, `0.0%` for percentages
- **Formulas** -- a totals row below the table: `SUM` for the four quarters and the Total column, `AVERAGE` for Growth
- **Negative-growth style** -- the script styles any negative growth cell red, but the sample data has none, so no cell is red
- **Auto-filter** -- dropdown filters on the table header
- **Frozen header** -- header row stays visible while scrolling
- **Column widths** -- fixed per column in the table configuration

## Usage

```bash
npm install
node index.mjs
```

This produces `sales-report.xlsx` in the current directory. Open it in
Excel, Google Sheets, or LibreOffice to see the full styling.

## Key APIs Used

| API                                       | Purpose                                                 |
| ----------------------------------------- | ------------------------------------------------------- |
| `Workbook` / `addSheet()`                 | Create workbook structure                               |
| `drawTable()`                             | High-level table layout with automatic styling          |
| `StyleBuilder`                            | Build custom styles (font, fill, border, number format) |
| `ws.cell(ref)`                            | Get/set cell values and formulas                        |
| `drawTable({ freezeHeader, autoFilter })` | Freeze the header row and add filter dropdowns          |
| `ws.addMergeCell()`                       | Merge cell ranges                                       |
| `wb.toFile(path)`                         | Write the XLSX file                                     |
