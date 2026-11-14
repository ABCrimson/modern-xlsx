# CSV to Styled XLSX

Reads a CSV file, auto-detects column types, and writes a professionally
styled XLSX workbook.

## What It Does

1. **Parses CSV** -- reads the included `data.csv` (or any CSV you point it to) with a small built-in RFC 4180 parser: quoted fields, doubled quotes, CRLF or LF line ends, and newlines inside quotes
2. **Detects types** -- samples all values per column to infer number, date, boolean, or string
3. **Applies formatting** -- numbers get `#,##0` (`#,##0.0#` when a column holds any non-integer value), dates get `yyyy-mm-dd`, booleans center-align
4. **Styles the output** -- header row, zebra striping, thin borders, column widths estimated from the longest value (value length + 3 characters, between 8 and 40)
5. **Adds features** -- frozen header row, auto-filter dropdowns

## Usage

```bash
npm install
node index.mjs                 # uses the included data.csv
node index.mjs my-data.csv     # use your own CSV file
```

The output file will be named the same as the input with a `.xlsx` extension
(e.g., `data.xlsx`).

## Sample Data

The included `data.csv` contains employee records with mixed types:

| Column      | Detected Type | Format Applied |
| ----------- | ------------- | -------------- |
| Name        | string        | left-aligned   |
| Department  | string        | left-aligned   |
| Start Date  | date          | yyyy-mm-dd     |
| Salary      | number        | #,##0          |
| Performance | number        | #,##0.0#       |
| Active      | boolean       | centered       |

## Key APIs Used

| API                       | Purpose                                          |
| ------------------------- | ------------------------------------------------ |
| `Workbook` / `addSheet()` | Create workbook structure                        |
| `StyleBuilder`            | Build header, even-row, and odd-row styles       |
| `dateToSerial()`          | Turn `YYYY-MM-DD` text into an Excel date serial |
| `ws.cell(ref).value`      | Set typed cell values                            |
| `ws.setColumnWidth()`     | Estimated widths                                 |
| `ws.frozenPane`           | Freeze header row                                |
| `ws.autoFilter`           | Filter dropdowns                                 |
| `wb.toFile(path)`         | Write the XLSX file next to the input CSV        |
