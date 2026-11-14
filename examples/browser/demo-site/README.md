# modern-xlsx Demo Site

A single-page application demonstrating modern-xlsx in the browser. Uses the IIFE
build (`modern-xlsx.min.js`) loaded from the jsDelivr CDN — no bundler required.

## Features

- **Create Workbook** — generates a styled sales report with formulas, frozen header, and auto-filter
- **Read File** — drag-and-drop or browse for an XLSX file; contents are rendered as an HTML table
- **Style Demo** — showcases bold, italic, fill colors, border styles, and number formats
- **Chart Demo** — creates a workbook with an embedded bar chart (open in Excel to view)

## Running

Open `index.html` directly in a modern browser, or serve it with any static file server:

```bash
# Python
python -m http.server 8080

# Node.js (npx)
npx serve .

# Then open http://localhost:8080
```

No build step, no `npm install`, no framework. `initWasm()` takes no argument:
the browser bundle records its own script URL and fetches `modern-xlsx.wasm`
from beside it on the CDN (`packages/modern-xlsx/src/browser-entry.ts`).

## How It Works

The page loads `modern-xlsx.min.js` via a `<script>` tag which exposes
`window.ModernXlsx`. On load it calls `ModernXlsx.initWasm()` to fetch and
compile the WebAssembly module. Once it resolves, creating a workbook and
writing it with `writeBlob()` are synchronous, while `readBuffer()` is
asynchronous and the page awaits it.

Key APIs used:

| API                     | Purpose                                               |
| ----------------------- | ----------------------------------------------------- |
| `initWasm()`            | Initialize the WASM engine                            |
| `new Workbook()`        | Create an empty workbook                              |
| `wb.addSheet(name)`     | Add a worksheet                                       |
| `ws.cell(ref)`          | Get/create a cell                                     |
| `StyleBuilder`          | Build cell styles (font, fill, border, number format) |
| `readBuffer(bytes)`     | Parse an XLSX file from `Uint8Array`                  |
| `writeBlob(wb)`         | Serialize a workbook to a `Blob` for download         |
| `sheetToHtml(ws)`       | Convert a worksheet to an HTML table string           |
| `ws.addChart(type, cb)` | Add a chart via the fluent ChartBuilder               |
