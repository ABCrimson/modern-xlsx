# Feature Comparison: modern-xlsx vs SheetJS (xlsx)

> **Exhaustive capability matrix** — every feature either library supports, compared side-by-side.

| Legend | Meaning |
|--------|---------|
| :white_check_mark: | Fully supported |
| :star: | Supported + superior implementation |
| :lock: | SheetJS Pro only (paid license) |
| :x: | Not supported |
| :construction: | Partial / limited support |

---

## 1. File Format Support

### Read Formats

| Format | modern-xlsx | SheetJS | Notes |
|--------|:-----------:|:-------:|-------|
| XLSX (Office Open XML) | :white_check_mark: | :white_check_mark: | Both full support |
| XLSM (Macro-enabled) | :x: | :white_check_mark: | SheetJS reads data (not macros) |
| XLSB (Binary) | :x: | :white_check_mark: | SheetJS proprietary binary format |
| XLS (BIFF8/5/4/3/2) | :x: | :white_check_mark: | Legacy Excel 97-2003 |
| XLML (SpreadsheetML 2003) | :x: | :white_check_mark: | XML-based legacy format |
| ODS (OpenDocument) | :x: | :white_check_mark: | LibreOffice/OpenOffice native |
| FODS (Flat ODS) | :x: | :white_check_mark: | Single-file ODS variant |
| CSV | :x: | :white_check_mark: | SheetJS reads CSV as workbook |
| TSV / TXT | :x: | :white_check_mark: | Tab-separated values |
| HTML tables | :x: | :white_check_mark: | Parse `<table>` from HTML |
| SYLK | :x: | :white_check_mark: | Symbolic Link format |
| DIF | :x: | :white_check_mark: | Data Interchange Format |
| PRN | :x: | :white_check_mark: | Lotus Formatted Text |
| ETH | :x: | :white_check_mark: | EtherCalc record format |
| DBF (dBASE) | :x: | :white_check_mark: | dBASE III/IV |
| WK1/WK3/WKS | :x: | :white_check_mark: | Lotus 1-2-3 workbooks |
| QPW | :x: | :white_check_mark: | Quattro Pro |
| Numbers (Apple) | :x: | :white_check_mark: | Requires `numbers` option |

### Write Formats

| Format | modern-xlsx | SheetJS | Notes |
|--------|:-----------:|:-------:|-------|
| XLSX | :star: | :white_check_mark: | modern-xlsx: 8x smaller output via WASM-native ZIP |
| XLSM | :x: | :white_check_mark: | |
| XLSB | :x: | :white_check_mark: | |
| XLS (BIFF8) | :x: | :white_check_mark: | |
| XLS (BIFF5) | :x: | :white_check_mark: | |
| XLS (BIFF4/3/2) | :x: | :white_check_mark: | |
| XLML | :x: | :white_check_mark: | |
| ODS | :x: | :white_check_mark: | |
| FODS | :x: | :white_check_mark: | |
| CSV | :x: | :white_check_mark: | modern-xlsx converts to CSV string but doesn't write file |
| TXT | :x: | :white_check_mark: | |
| SYLK | :x: | :white_check_mark: | |
| HTML | :x: | :white_check_mark: | modern-xlsx converts to HTML string but doesn't write file |
| DIF | :x: | :white_check_mark: | |
| RTF | :x: | :white_check_mark: | |
| PRN | :x: | :white_check_mark: | |
| ETH | :x: | :white_check_mark: | |
| DBF | :x: | :white_check_mark: | |
| WK3 | :x: | :white_check_mark: | |
| Numbers | :x: | :construction: | Requires external `numbers` option |

---

## 2. I/O Operations

| Feature | modern-xlsx | SheetJS | Notes |
|---------|:-----------:|:-------:|-------|
| Read from `Uint8Array` / Buffer | :white_check_mark: `readBuffer()` | :white_check_mark: `XLSX.read()` | |
| Read from file path | :white_check_mark: `readFile()` | :white_check_mark: `XLSX.readFile()` | |
| Read from file (sync) | :x: | :white_check_mark: `readFileSync()` | modern-xlsx is async-only (WASM) |
| Write to `Uint8Array` / Buffer | :white_check_mark: `wb.toBuffer()` | :white_check_mark: `XLSX.write()` | |
| Write to file | :white_check_mark: `wb.toFile()` | :white_check_mark: `XLSX.writeFile()` | |
| Write to file (sync) | :x: | :white_check_mark: `writeFileSync()` | modern-xlsx is async-only |
| Write to file (async) | :white_check_mark: `wb.toFile()` | :white_check_mark: `writeFileAsync()` | |
| Write to Blob (browser) | :star: `writeBlob()` | :white_check_mark: via `type:'array'` | modern-xlsx: native WASM Blob |
| Parse CFB containers | :x: | :white_check_mark: `parse_xlscfb()` | Compound File Binary |
| Parse ZIP directly | :x: | :white_check_mark: `parse_zip()` | |
| Set custom FS module | :x: | :white_check_mark: `set_fs()` | |
| Set codepage tables | :x: | :white_check_mark: `set_cptable()` | For legacy encoding |
| WASM initialization | :star: `initWasm()` | :x: | Zero-copy WASM acceleration |
| WASM sync init | :white_check_mark: `initWasmSync()` | :x: | |

---

## 3. Workbook Operations

| Feature | modern-xlsx | SheetJS | Notes |
|---------|:-----------:|:-------:|-------|
| Create new workbook | :white_check_mark: `new Workbook()` | :white_check_mark: `utils.book_new()` | |
| Get sheet names | :white_check_mark: `.sheetNames` | :white_check_mark: `.SheetNames` | |
| Get sheet count | :white_check_mark: `.sheetCount` | :construction: `.SheetNames.length` | |
| Get sheet by name | :white_check_mark: `.getSheet(name)` | :white_check_mark: `.Sheets[name]` | |
| Get sheet by index | :white_check_mark: `.getSheetByIndex(i)` | :construction: `.Sheets[.SheetNames[i]]` | |
| Add sheet | :white_check_mark: `.addSheet(name)` | :white_check_mark: `utils.book_append_sheet()` | |
| Remove sheet | :white_check_mark: `.removeSheet(name)` | :x: | Manual splice required in SheetJS |
| Sheet visibility | :x: | :white_check_mark: `utils.book_set_sheet_visibility()` | Hidden / very hidden |
| Date system (1900/1904) | :white_check_mark: `.dateSystem` | :white_check_mark: `.Workbook.WBProps.date1904` | |
| Workbook views | :star: `.workbookViews` | :construction: `.Workbook.Views` | modern-xlsx: typed interface |
| Document properties | :star: `.docProperties` | :construction: `.Props` | modern-xlsx: typed `DocPropertiesData` |
| Serialize to JSON | :white_check_mark: `.toJSON()` | :x: | WASM boundary serialization |
| Validate workbook | :star: `.validate()` | :x: | Returns `ValidationReport` with issues |
| Repair workbook | :star: `.repair()` | :x: | Auto-fix common issues |

---

## 4. Worksheet Operations

| Feature | modern-xlsx | SheetJS | Notes |
|---------|:-----------:|:-------:|-------|
| Create new sheet | :white_check_mark: `ws = wb.addSheet()` | :white_check_mark: `utils.sheet_new()` | |
| Sheet name | :white_check_mark: `.name` | :x: | Tracked via `SheetNames[]` in SheetJS |
| Access rows | :star: `.rows` | :x: | Typed `RowData[]`; SheetJS uses flat cell refs |
| Access columns | :white_check_mark: `.columns` | :white_check_mark: `ws['!cols']` | |
| Set column width | :white_check_mark: `.setColumnWidth(col, w)` | :construction: `ws['!cols'][i].wch` | |
| Set row height | :white_check_mark: `.setRowHeight(row, h)` | :construction: `ws['!rows'][i].hpt` | |
| Hide row | :white_check_mark: `.setRowHidden(row, b)` | :construction: `ws['!rows'][i].hidden` | |
| Hide column | :construction: via `.columns` | :construction: `ws['!cols'][i].hidden` | |
| Sheet used range | :x: | :white_check_mark: `ws['!ref']` | Auto-computed from data |
| Outline / grouping | :x: | :construction: `ws['!outline']` | Row/column grouping |
| Sheet tab color | :x: | :construction: via `Workbook.Sheets` | |

---

## 5. Cell Operations

| Feature | modern-xlsx | SheetJS | Notes |
|---------|:-----------:|:-------:|-------|
| Read/write by reference | :white_check_mark: `ws.cell('A1').value` | :white_check_mark: `ws['A1'].v` | |
| Get cell (safe) | :white_check_mark: `ws.cell('A1')` | :white_check_mark: `utils.sheet_get_cell()` | |
| String values | :white_check_mark: | :white_check_mark: | |
| Number values | :white_check_mark: | :white_check_mark: | |
| Boolean values | :white_check_mark: | :white_check_mark: | |
| Error values | :white_check_mark: type `'error'` | :white_check_mark: type `'e'` | |
| Date values (native) | :construction: via serial + format | :white_check_mark: type `'d'` | SheetJS has native date type |
| Stub/empty cells | :x: | :white_check_mark: type `'z'` | SheetJS can represent empty cells |
| Inline strings | :white_check_mark: type `'inlineStr'` | :x: | Stored directly in cell XML |
| Formula strings | :white_check_mark: type `'formulaStr'` | :x: | Formula returning string |
| Cell type field | :star: `.type` (getter) | :white_check_mark: `.t` | modern-xlsx: semantic types |
| Cell reference field | :white_check_mark: `.reference` | :x: | Implicit from key in SheetJS |
| Style index | :star: `.styleIndex` | :lock: `.s` | SheetJS styling is Pro only |
| Formatted text | :white_check_mark: via `formatCell()` | :white_check_mark: `.w` | |
| Rich text (in cell) | :construction: via SST `richRuns` | :white_check_mark: `.r` (HTML) | SheetJS stores as HTML |
| Cell HTML rendering | :x: | :white_check_mark: `.h` | |
| Number format override | :construction: via style index | :white_check_mark: `.z` | SheetJS: per-cell `.z` property |

---

## 6. Formulas

| Feature | modern-xlsx | SheetJS | Notes |
|---------|:-----------:|:-------:|-------|
| Simple formulas | :white_check_mark: `.formula = 'SUM(A1:A3)'` | :white_check_mark: `.f = 'SUM(...)'` | |
| Array formulas | :white_check_mark: `formulaType: 'array'` | :white_check_mark: `.F` range + `sheet_set_array_formula()` | |
| Shared formulas | :white_check_mark: `formulaType: 'shared'` + `sharedIndex` | :x: | |
| Dynamic array formulas | :x: | :white_check_mark: `.D = true` | SPILL formulas |
| Formula reference | :white_check_mark: `.formulaRef` | :white_check_mark: `.F` | Range for array formulas |
| Sheet to formulae | :x: | :white_check_mark: `utils.sheet_to_formulae()` | Extract all formulas as strings |
| Calc chain | :star: `.calcChain` | :construction: parsed but not exposed | modern-xlsx: typed `CalcChainEntryData[]` |

---

## 7. Merge Cells

| Feature | modern-xlsx | SheetJS | Notes |
|---------|:-----------:|:-------:|-------|
| Get merge ranges | :white_check_mark: `.mergeCells` | :white_check_mark: `ws['!merges']` | |
| Add merge range | :white_check_mark: `.addMergeCell('A1:C1')` | :construction: `ws['!merges'].push()` | |
| Remove merge range | :white_check_mark: `.removeMergeCell('A1:C1')` | :x: | Manual splice in SheetJS |

---

## 8. Styling

| Feature | modern-xlsx | SheetJS | Notes |
|---------|:-----------:|:-------:|-------|
| **Style system** | :star: `StyleBuilder` API | :lock: Pro only | **modern-xlsx free; SheetJS paid** |
| Font (name, size, bold, italic) | :star: | :lock: | |
| Font color | :star: | :lock: | |
| Font underline | :star: | :lock: | |
| Font strikethrough | :star: | :lock: | |
| Fill (solid pattern) | :star: | :lock: | |
| Fill (pattern types — 18 patterns) | :star: | :lock: | gray125, darkGrid, lightTrellis, etc. |
| Fill (gradient — linear/path) | :star: | :lock: | `GradientFillData` with stops |
| Border (top/bottom/left/right) | :star: | :lock: | |
| Border (diagonal) | :star: | :lock: | |
| Border styles (13 styles) | :star: | :lock: | thin, medium, thick, dashed, dotted, etc. |
| Border colors | :star: | :lock: | |
| Alignment (horizontal) | :star: | :lock: | left, center, right, fill, justify, etc. |
| Alignment (vertical) | :star: | :lock: | top, center, bottom, justify, distributed |
| Alignment (wrap text) | :star: | :lock: | |
| Alignment (text rotation) | :star: | :lock: | |
| Alignment (indent) | :star: | :lock: | |
| Alignment (shrink to fit) | :star: | :lock: | |
| Cell protection (locked) | :star: | :lock: | |
| Cell protection (hidden) | :star: | :lock: | |
| Number format (custom) | :star: `numberFormat()` | :white_check_mark: `cell_set_number_format()` | Community edition has per-cell format |
| Number format (built-in table) | :star: `getBuiltinFormat()` | :white_check_mark: `SSF.get_table()` | |
| DXF styles (differential) | :star: | :x: | For conditional formatting |
| Cell styles (named) | :star: `CellStyleData` | :x: | Normal, Title, Heading, etc. |
| CellXf (cell format records) | :star: | :lock: | Full OOXML xf records |
| NumFmt records | :star: | :construction: | Custom format records |
| Theme colors | :star: `.themeColors` | :x: | Parsed `ThemeColorsData` |
| Style builder (fluent API) | :star: `wb.createStyle().font().fill().build()` | :x: | Chainable builder pattern |

---

## 9. Frozen Panes

| Feature | modern-xlsx | SheetJS | Notes |
|---------|:-----------:|:-------:|-------|
| Freeze rows | :white_check_mark: `.frozenPane = { rows: 1, cols: 0 }` | :white_check_mark: `ws['!freeze']` | |
| Freeze columns | :white_check_mark: `.frozenPane = { rows: 0, cols: 1 }` | :white_check_mark: `ws['!freeze']` | |
| Freeze both | :white_check_mark: | :white_check_mark: | |
| Split panes | :x: | :x: | Neither supports split panes |

---

## 10. Auto Filter

| Feature | modern-xlsx | SheetJS | Notes |
|---------|:-----------:|:-------:|-------|
| Set auto filter range | :white_check_mark: `.autoFilter = 'A1:C10'` | :white_check_mark: `ws['!autofilter'] = { ref: '...' }` | |
| Filter column definitions | :star: `FilterColumnData` | :x: | modern-xlsx: typed filter conditions |
| Remove auto filter | :white_check_mark: `.autoFilter = null` | :construction: `delete ws['!autofilter']` | |

---

## 11. Data Validation

| Feature | modern-xlsx | SheetJS | Notes |
|---------|:-----------:|:-------:|-------|
| Add validation rule | :star: `.addValidation()` | :lock: Pro only | **modern-xlsx free** |
| List validation | :star: `validationType: 'list'` | :lock: | Dropdown lists |
| Whole number validation | :star: `validationType: 'whole'` | :lock: | |
| Decimal validation | :star: `validationType: 'decimal'` | :lock: | |
| Date validation | :star: `validationType: 'date'` | :lock: | |
| Time validation | :star: `validationType: 'time'` | :lock: | |
| Text length validation | :star: `validationType: 'textLength'` | :lock: | |
| Custom formula validation | :star: `validationType: 'custom'` | :lock: | |
| Validation operators | :star: between, notBetween, equal, etc. | :lock: | |
| Input prompt | :star: `prompt` / `promptTitle` | :lock: | |
| Error alert | :star: `errorMessage` / `errorTitle` | :lock: | |
| Remove validation | :white_check_mark: `.removeValidation()` | :lock: | |
| Get validations | :white_check_mark: `.validations` | :lock: | |

---

## 12. Conditional Formatting

| Feature | modern-xlsx | SheetJS | Notes |
|---------|:-----------:|:-------:|-------|
| Conditional format rules | :star: `ConditionalFormattingData` | :lock: Pro only | **modern-xlsx free** |
| Color scales | :star: `ColorScaleData` | :lock: | 2-color and 3-color |
| Data bars | :star: `DataBarData` | :lock: | |
| Icon sets | :star: `IconSetData` | :lock: | |
| Cell value rules | :star: | :lock: | |
| DXF style references | :star: `DxfStyleData` | :lock: | Differential formatting |
| CFVO (conditional format value objects) | :star: `CfvoData` | :lock: | min/max/percentile/formula |

---

## 13. Hyperlinks

| Feature | modern-xlsx | SheetJS | Notes |
|---------|:-----------:|:-------:|-------|
| Add external hyperlink | :white_check_mark: `.addHyperlink()` | :white_check_mark: `cell_set_hyperlink()` | |
| Add internal link | :white_check_mark: | :white_check_mark: `cell_set_internal_link()` | |
| Link display text | :white_check_mark: `display` | :construction: | |
| Link tooltip | :white_check_mark: `tooltip` | :x: | |
| Remove hyperlink | :white_check_mark: `.removeHyperlink()` | :x: | |
| Get all hyperlinks | :white_check_mark: `.hyperlinks` | :construction: cell `.l` property | |

---

## 14. Comments

| Feature | modern-xlsx | SheetJS | Notes |
|---------|:-----------:|:-------:|-------|
| Add comment | :white_check_mark: `.addComment()` | :white_check_mark: `cell_add_comment()` | |
| Remove comment | :white_check_mark: `.removeComment()` | :x: | |
| Get all comments | :white_check_mark: `.comments` | :construction: cell `.c` array | |
| Comment author | :white_check_mark: | :construction: `.c[].a` | |
| Comment text | :white_check_mark: | :construction: `.c[].t` | |
| Threaded comments | :x: | :x: | Neither supports |

---

## 15. Named Ranges

| Feature | modern-xlsx | SheetJS | Notes |
|---------|:-----------:|:-------:|-------|
| Get all named ranges | :star: `.namedRanges` | :construction: `.Workbook.Names` | modern-xlsx: typed `DefinedNameData[]` |
| Add named range | :white_check_mark: `.addNamedRange()` | :x: | |
| Get named range by name | :white_check_mark: `.getNamedRange()` | :x: | |
| Remove named range | :white_check_mark: `.removeNamedRange()` | :x: | |
| Scoped named ranges | :white_check_mark: `localSheetId` | :construction: | |

---

## 16. Page Setup & Print

| Feature | modern-xlsx | SheetJS | Notes |
|---------|:-----------:|:-------:|-------|
| Page orientation | :white_check_mark: `pageSetup.orientation` | :x: | |
| Paper size | :white_check_mark: `pageSetup.paperSize` | :x: | |
| Fit to width/height | :white_check_mark: `pageSetup.fitToWidth/Height` | :x: | |
| Scale | :white_check_mark: `pageSetup.scale` | :x: | |
| Page margins | :white_check_mark: `pageMargins` | :white_check_mark: `ws['!margins']` | |
| Header/footer margins | :white_check_mark: `.header` / `.footer` | :white_check_mark: | |
| Print area | :x: | :x: | Neither directly supports |
| Page breaks | :x: | :x: | |
| Print titles (repeat rows) | :x: | :x: | |

---

## 17. Sheet Protection

| Feature | modern-xlsx | SheetJS | Notes |
|---------|:-----------:|:-------:|-------|
| Protect sheet | :star: `.sheetProtection` | :white_check_mark: `ws['!protect']` | |
| Protect objects | :star: `.objects` | :construction: | |
| Protect scenarios | :star: `.scenarios` | :construction: | |
| Protect format cells | :star: `.formatCells` | :construction: | |
| Protect format columns | :star: `.formatColumns` | :construction: | |
| Protect format rows | :star: `.formatRows` | :construction: | |
| Protect insert columns | :star: `.insertColumns` | :construction: | |
| Protect insert rows | :star: `.insertRows` | :construction: | |
| Protect delete columns | :star: `.deleteColumns` | :construction: | |
| Protect delete rows | :star: `.deleteRows` | :construction: | |
| Protect sort | :star: `.sort` | :construction: | |
| Protect auto filter | :star: `.autoFilter` | :construction: | |
| Select locked cells | :star: | :construction: | |
| Select unlocked cells | :star: | :construction: | |
| Password protection | :x: | :construction: | |

---

## 18. Document Properties

| Feature | modern-xlsx | SheetJS | Notes |
|---------|:-----------:|:-------:|-------|
| Title | :white_check_mark: `.docProperties.title` | :white_check_mark: `.Props.Title` | |
| Creator / Author | :white_check_mark: `.creator` | :white_check_mark: `.Props.Author` | |
| Description / Subject | :white_check_mark: `.description` | :white_check_mark: `.Props.Subject` | |
| Created date | :white_check_mark: `.created` | :white_check_mark: `.Props.CreatedDate` | |
| Modified date | :white_check_mark: `.modified` | :white_check_mark: `.Props.ModifiedDate` | |
| Keywords | :white_check_mark: `.keywords` | :white_check_mark: `.Props.Keywords` | |
| Category | :white_check_mark: `.category` | :white_check_mark: `.Props.Category` | |
| Last modified by | :white_check_mark: `.lastModifiedBy` | :white_check_mark: `.Props.LastAuthor` | |
| Application name | :x: | :white_check_mark: `.Props.Application` | |
| App version | :x: | :white_check_mark: `.Props.AppVersion` | |
| Company | :x: | :white_check_mark: `.Props.Company` | |
| Manager | :x: | :white_check_mark: `.Props.Manager` | |

---

## 19. Cell Reference Utilities

| Feature | modern-xlsx | SheetJS | Notes |
|---------|:-----------:|:-------:|-------|
| Column index → letter | :white_check_mark: `columnToLetter(0)` → `'A'` | :white_check_mark: `encode_col(0)` → `'A'` | |
| Letter → column index | :white_check_mark: `letterToColumn('A')` → `0` | :white_check_mark: `decode_col('A')` → `0` | |
| Decode cell ref | :white_check_mark: `decodeCellRef('B3')` → `{row:2,col:1}` | :white_check_mark: `decode_cell('B3')` → `{r:2,c:1}` | |
| Encode cell ref | :white_check_mark: `encodeCellRef(2,1)` → `'B3'` | :white_check_mark: `encode_cell({r:2,c:1})` → `'B3'` | |
| Decode range | :white_check_mark: `decodeRange('A1:C3')` | :white_check_mark: `decode_range('A1:C3')` | |
| Encode range | :white_check_mark: `encodeRange(...)` | :white_check_mark: `encode_range(...)` | |
| Encode row | :x: | :white_check_mark: `encode_row(0)` → `'1'` | |
| Decode row | :x: | :white_check_mark: `decode_row('1')` → `0` | |
| Split cell ref | :x: | :white_check_mark: `split_cell('$A$1')` | Separates col/row parts |

---

## 20. Date Utilities

| Feature | modern-xlsx | SheetJS | Notes |
|---------|:-----------:|:-------:|-------|
| Date → serial number | :star: `dateToSerial({year,month,day})` | :x: | Temporal-like input |
| Serial number → Date | :star: `serialToDate(serial)` | :x: | Returns JS Date (UTC) |
| Is date format code | :star: `isDateFormatCode('yyyy-mm-dd')` | :white_check_mark: `SSF.is_date(format, value)` | |
| Is date format ID | :star: `isDateFormatId(14)` | :construction: via `SSF._table` | |
| Parse date code | :x: | :white_check_mark: `SSF.parse_date_code(serial)` | Returns `{y,m,d,H,M,S}` |
| Is Temporal-like | :white_check_mark: `isTemporalLike(obj)` | :x: | |
| Lotus 1-2-3 bug handling | :white_check_mark: serial 60 handling | :white_check_mark: | Feb 29, 1900 quirk |

---

## 21. Number Formatting (SSF)

| Feature | modern-xlsx | SheetJS | Notes |
|---------|:-----------:|:-------:|-------|
| Format cell value | :white_check_mark: `formatCell(42, '#,##0')` | :white_check_mark: `SSF.format(fmt, val)` | |
| Built-in format table | :white_check_mark: `getBuiltinFormat(14)` | :white_check_mark: `SSF.get_table()` | |
| Load custom format | :x: | :white_check_mark: `SSF.load(fmt, idx)` | Register custom format |
| Load format table | :x: | :white_check_mark: `SSF.load_table(table)` | Bulk register |
| General format | :white_check_mark: | :white_check_mark: | |
| Number formats (#,##0) | :white_check_mark: | :white_check_mark: | |
| Percentage (0%) | :white_check_mark: | :white_check_mark: | |
| Scientific (0.00E+00) | :white_check_mark: | :white_check_mark: | |
| Fraction (# ?/?) | :white_check_mark: | :white_check_mark: | |
| Date formats | :white_check_mark: | :white_check_mark: | |
| Time formats | :white_check_mark: | :white_check_mark: | |
| Multi-section formats | :white_check_mark: | :white_check_mark: | pos;neg;zero;text |
| Color codes ([Red], etc.) | :white_check_mark: | :white_check_mark: | |
| Locale codes ([$-409]) | :construction: | :white_check_mark: | |
| Conditional sections | :construction: | :white_check_mark: | [>100]#,##0 |

---

## 22. Sheet Conversion Utilities

| Feature | modern-xlsx | SheetJS | Notes |
|---------|:-----------:|:-------:|-------|
| Array-of-arrays → sheet | :white_check_mark: `aoaToSheet(data)` | :white_check_mark: `aoa_to_sheet(data)` | |
| JSON → sheet | :white_check_mark: `jsonToSheet(data)` | :white_check_mark: `json_to_sheet(data)` | |
| Sheet → JSON | :white_check_mark: `sheetToJson(ws)` | :white_check_mark: `sheet_to_json(ws)` | |
| Sheet → CSV | :white_check_mark: `sheetToCsv(ws)` | :white_check_mark: `sheet_to_csv(ws)` | |
| Sheet → HTML | :white_check_mark: `sheetToHtml(ws)` | :white_check_mark: `sheet_to_html(ws)` | |
| Sheet → TXT | :x: | :white_check_mark: `sheet_to_txt(ws)` | Tab-separated |
| Sheet → formulae | :x: | :white_check_mark: `sheet_to_formulae(ws)` | Formula strings |
| Sheet → row objects | :x: | :white_check_mark: `sheet_to_row_object_array(ws)` | Alias for sheet_to_json |
| Add AoA to existing sheet | :white_check_mark: `sheetAddAoa(ws, data)` | :white_check_mark: `sheet_add_aoa(ws, data)` | |
| Add JSON to existing sheet | :white_check_mark: `sheetAddJson(ws, data)` | :white_check_mark: `sheet_add_json(ws, data)` | |
| DOM table → sheet | :x: | :white_check_mark: `sheet_add_dom(ws, el)` | Browser DOM element |
| DOM table → book | :x: | :white_check_mark: `table_to_book(el)` | |
| DOM table → sheet | :x: | :white_check_mark: `table_to_sheet(el)` | |
| Format cell utility | :white_check_mark: `formatCell()` | :white_check_mark: `format_cell()` | |
| JSON → sheet (options) | :white_check_mark: `JsonToSheetOptions` | :white_check_mark: options object | |
| AoA → sheet (options) | :white_check_mark: `AoaToSheetOptions` | :white_check_mark: options object | |
| CSV options (separator, etc.) | :white_check_mark: `SheetToCsvOptions` | :white_check_mark: options object | |
| HTML options | :white_check_mark: `SheetToHtmlOptions` | :white_check_mark: options object | |

---

## 23. Streaming

| Feature | modern-xlsx | SheetJS | Notes |
|---------|:-----------:|:-------:|-------|
| Stream to JSON | :x: | :white_check_mark: `stream.to_json()` | Node.js Readable stream |
| Stream to CSV | :x: | :white_check_mark: `stream.to_csv()` | |
| Stream to HTML | :x: | :white_check_mark: `stream.to_html()` | |
| Stream to XLML | :x: | :white_check_mark: `stream.to_xlml()` | |
| Set Readable impl | :x: | :white_check_mark: `stream.set_readable()` | Custom stream impl |
| WASM streaming reader | :star: via WASM SAX parser | :x: | Rust SAX-style XML parsing |
| Parallel sheet parsing | :star: rayon feature flag | :x: | Multi-threaded WASM |

---

## 24. Rich Text

| Feature | modern-xlsx | SheetJS | Notes |
|---------|:-----------:|:-------:|-------|
| Rich text builder | :star: `RichTextBuilder` | :x: | Fluent API |
| Bold text runs | :star: `.bold('text')` | :x: | |
| Italic text runs | :star: `.italic('text')` | :x: | |
| Bold+italic runs | :star: `.boldItalic('text')` | :x: | |
| Colored text runs | :star: `.colored('text', 'FF0000')` | :x: | |
| Custom styled runs | :star: `.styled('text', opts)` | :x: | font, size, bold, italic, color |
| Build rich text array | :star: `.build()` → `RichTextRun[]` | :x: | |
| Plain text extraction | :star: `.plainText()` | :x: | |
| Rich text in cells | :construction: via SST `richRuns` | :white_check_mark: `.r` (HTML string) | SheetJS: HTML-based |
| Rich text roundtrip | :white_check_mark: | :white_check_mark: | |

---

## 25. Images & Charts

| Feature | modern-xlsx | SheetJS | Notes |
|---------|:-----------:|:-------:|-------|
| Add image (PNG) | :star: `wb.addImage()` | :lock: Pro only | **modern-xlsx free** |
| Add image (JPEG) | :star: | :lock: | |
| Add image (GIF) | :star: | :lock: | |
| Image anchor (from/to cell) | :star: `ImageAnchor` | :lock: | |
| Charts (read) | :construction: preserved blob | :lock: | Passthrough roundtrip |
| Charts (create) | :x: | :lock: | |
| Drawings (preserved roundtrip) | :white_check_mark: `preservedEntries` | :lock: | |
| Tables (read) | :x: | :lock: | |
| Tables (create) | :x: | :lock: | |

---

## 26. Barcode & QR Code Generation

| Feature | modern-xlsx | SheetJS | Notes |
|---------|:-----------:|:-------:|-------|
| QR code encoding | :star: `encodeQR()` | :x: | **Unique to modern-xlsx** |
| Code 128 encoding | :star: `encodeCode128()` | :x: | |
| EAN-13 encoding | :star: `encodeEAN13()` | :x: | |
| UPC-A encoding | :star: `encodeUPCA()` | :x: | |
| Code 39 encoding | :star: `encodeCode39()` | :x: | |
| PDF417 encoding | :star: `encodePDF417()` | :x: | |
| DataMatrix encoding | :star: `encodeDataMatrix()` | :x: | |
| ITF-14 encoding | :star: `encodeITF14()` | :x: | |
| GS1-128 encoding | :star: `encodeGS1128()` | :x: | |
| Render barcode to PNG | :star: `renderBarcodePNG()` | :x: | |
| Embed barcode in sheet | :star: `wb.addBarcode()` | :x: | One-call API |
| Barcode options (ecLevel, showText) | :star: `DrawBarcodeOptions` | :x: | |
| Generate drawing XML | :star: `generateDrawingXml()` | :x: | Low-level API |
| Generate drawing rels | :star: `generateDrawingRels()` | :x: | Low-level API |

---

## 27. Table Layout Engine

| Feature | modern-xlsx | SheetJS | Notes |
|---------|:-----------:|:-------:|-------|
| Draw styled table | :star: `drawTable(wb, ws, opts)` | :x: | **Unique to modern-xlsx** |
| Draw table from data | :star: `drawTableFromData(wb, ws, data)` | :x: | Auto-generate from 2D array |
| Column definitions | :star: `TableColumn` with width, align, format | :x: | |
| Header styling | :star: auto-styled headers | :x: | |
| Zebra striping | :star: alternating row colors | :x: | |
| Auto column widths | :star: | :x: | |
| Custom cell styles | :star: `CellStyle` overrides | :x: | |
| Table result metadata | :star: `TableResult` (range, dimensions) | :x: | |

---

## 28. Shared Strings (SST)

| Feature | modern-xlsx | SheetJS | Notes |
|---------|:-----------:|:-------:|-------|
| SST read | :white_check_mark: auto-resolved | :white_check_mark: auto-resolved | |
| SST write | :white_check_mark: built internally by WASM | :white_check_mark: | |
| Rich text runs in SST | :white_check_mark: `SharedStringsData.richRuns` | :white_check_mark: | |
| SST deduplication | :white_check_mark: | :white_check_mark: | |
| SST index resolution | :star: automatic in `readBuffer()` | :white_check_mark: automatic | modern-xlsx: transparent |

---

## 29. Validation & Repair

| Feature | modern-xlsx | SheetJS | Notes |
|---------|:-----------:|:-------:|-------|
| Validate workbook | :star: `wb.validate()` | :x: | **Unique to modern-xlsx** |
| Validation report | :star: `ValidationReport` | :x: | Issues with severity & category |
| Issue categories | :star: `IssueCategory` enum | :x: | |
| Issue severity levels | :star: `Severity` enum | :x: | |
| Auto-repair | :star: `wb.repair()` | :x: | Returns `RepairResult` |
| Repair summary | :star: `RepairResult` | :x: | Changes made + report |

---

## 30. Web Worker Support

| Feature | modern-xlsx | SheetJS | Notes |
|---------|:-----------:|:-------:|-------|
| Dedicated worker API | :star: `createXlsxWorker()` | :x: | **Unique to modern-xlsx** |
| Worker options | :star: `XlsxWorkerOptions` | :x: | |
| Off-main-thread processing | :star: | :x: | WASM in Web Worker |
| Transferable buffers | :star: | :x: | Zero-copy to/from worker |

---

## 31. Performance & Architecture

| Feature | modern-xlsx | SheetJS | Notes |
|---------|:-----------:|:-------:|-------|
| Core language | :star: Rust (WASM) | JavaScript | Rust: memory-safe, fast |
| XML parsing | :star: SAX (quick-xml) | SAX (custom JS) | |
| ZIP handling | :star: Native Rust (zip crate) | JSZip / CFB | |
| Output file size | :star: ~8x smaller | Baseline | WASM-native compression |
| Read speed (10K rows) | :star: ~4.6x faster | Baseline | Benchmarked |
| Write speed (10K rows) | :star: ~1.3x faster | Baseline | Cell-by-cell API |
| aoaToSheet speed (50K rows) | :star: ~2x faster | Baseline | Batch API |
| sheetToJson speed (10K rows) | :star: ~2x faster | Baseline | |
| sheetToCsv speed (10K rows) | :star: ~2.4x faster | Baseline | |
| Tree-shakeable | :white_check_mark: ESM | :construction: CJS + ESM | |
| Zero runtime deps | :star: | :x: | SheetJS bundles CFB, SSF |
| TypeScript types | :star: Full generics + interfaces | :construction: `@types/xlsx` | |
| Type safety | :star: 109+ exported types | :construction: | |
| Bundle size (minified) | :construction: ~300KB (WASM) | :white_check_mark: ~200KB (JS) | WASM has fixed overhead |
| Async-first | :white_check_mark: | :x: | All I/O is async |
| Sync API | :x: | :white_check_mark: | SheetJS has sync methods |
| Browser support | :white_check_mark: | :white_check_mark: | Both work in browser |
| Node.js support | :white_check_mark: | :white_check_mark: | |
| Deno support | :white_check_mark: | :construction: | |
| Bun support | :white_check_mark: | :construction: | |
| Multi-threading (rayon) | :star: optional feature flag | :x: | Parallel sheet parsing |

---

## 32. API Design

| Feature | modern-xlsx | SheetJS | Notes |
|---------|:-----------:|:-------:|-------|
| API style | :star: Class-based OOP | POJO-based | |
| Cell access | :star: `ws.cell('A1').value` | `ws['A1'].v` | Getter/setter vs raw property |
| Method chaining | :star: StyleBuilder, RichTextBuilder | :x: | Fluent API |
| Null safety | :star: `getSheet()` returns `undefined` | :x: | |
| Error types | :star: `ModernXlsxError` (Rust thiserror) | Generic errors | |
| Versioning | :white_check_mark: `VERSION` constant | :white_check_mark: `version` property | |

---

## Summary Scorecard

| Category | modern-xlsx | SheetJS | Winner |
|----------|:-----------:|:-------:|--------|
| **Format support** (read) | 1 format | 20+ formats | SheetJS |
| **Format support** (write) | 1 format | 23 formats | SheetJS |
| **Styling** (free) | 30+ style features | 1 (number format) | modern-xlsx |
| **Styling** (with Pro) | 30+ style features | 30+ style features | Tie |
| **Data validation** | Full support (free) | Pro only | modern-xlsx |
| **Conditional formatting** | Full support (free) | Pro only | modern-xlsx |
| **Images** | Full support (free) | Pro only | modern-xlsx |
| **Barcode/QR generation** | 9 barcode types | None | modern-xlsx |
| **Table layout engine** | Full support | None | modern-xlsx |
| **Rich text builder** | Fluent API | None | modern-xlsx |
| **Validation & repair** | Full support | None | modern-xlsx |
| **Web Worker support** | Built-in API | None | modern-xlsx |
| **Performance (read)** | ~4.6x faster | Baseline | modern-xlsx |
| **Performance (write)** | ~1.3x faster | Baseline | modern-xlsx |
| **Output file size** | ~8x smaller | Baseline | modern-xlsx |
| **Type safety** | 109+ types | Basic typings | modern-xlsx |
| **Streaming export** | None | 4 stream formats | SheetJS |
| **DOM integration** | None | table_to_book, etc. | SheetJS |
| **Sync API** | None | Full sync support | SheetJS |
| **Legacy format support** | None | XLS, BIFF, WK, etc. | SheetJS |
| **Cell reference utils** | 6 functions | 9 functions | SheetJS |
| **Date utilities** | 5 functions | 3 functions | modern-xlsx |
| **Named ranges** | Full CRUD API | Read-only | modern-xlsx |
| **Sheet protection** | 14 granular fields | Basic | modern-xlsx |
| **Theme colors** | Full parsing | None | modern-xlsx |
| **Calc chain** | Typed access | Parsed but hidden | modern-xlsx |

### Overall

- **modern-xlsx wins on:** Performance, type safety, styling (free), data validation, conditional formatting, images, barcodes, table layout, rich text API, validation/repair, Web Workers, output size, API ergonomics
- **SheetJS wins on:** Multi-format support (20+ read, 23 write), streaming exports, DOM integration, sync API, legacy format compatibility, broader ecosystem maturity
- **Tie on:** Core XLSX read/write, cell operations, formulas, merge cells, frozen panes, auto filter, hyperlinks, comments, page setup, document properties, number formatting, sheet conversion utilities
