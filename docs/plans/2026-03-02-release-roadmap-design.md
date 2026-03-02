# modern-xlsx Release Roadmap

<div align="center">

**0.1.0 → 1.0.0**

From functional prototype to production-grade XLSX library

</div>

---

|  |  |
|---|---|
| **Date** | 2026-03-02 |
| **Status** | Draft |
| **Approach** | Feature-driven — each minor version unlocks a new capability milestone |
| **Total releases** | 90 (9 minors × 10 versions each) + 1.0.0 |

---

## Guiding Principles

| Principle | Rule |
|---|---|
| **SemVer** | 0.x allows breaking changes; 1.0.0 = API freeze |
| **TDD** | Every feature ships with tests written first |
| **CI** | Green on every push — no exceptions |
| **Scope** | Each patch is small, focused, and independently shippable |
| **Changelog** | Every release gets a changelog entry |

---

## Current State — 0.1.0

### Fully Implemented

| Category | Features |
|---|---|
| **Cell types** | Numbers, booleans, strings (SST), formulas (normal/array/shared), inline strings |
| **Styles** | Fonts, fills, borders, alignment, protection, number formats, gradient fills, diagonal borders, DXF styles, cell styles |
| **Rich text** | Formatted runs in shared strings, RichTextBuilder API |
| **Sheet features** | Merge cells, frozen panes, auto-filter (with filter columns), named ranges, data validation (with prompts), conditional formatting (colorScale/dataBar/iconSet), hyperlinks, comments/notes |
| **Protection** | Sheet protection (11 boolean flags + password hash) |
| **Page setup** | Paper size, orientation, scale, DPI, fit-to (read + write) |
| **Metadata** | Document properties (Dublin Core + app), theme colors (read-only), calc chain, workbook views |
| **I/O** | WASM JSON bridge (8–13x faster), streaming reader/writer, parallel parsing (rayon), browser Blob, Node.js file I/O |
| **Utilities** | Cell references, dates, formatCell, sheetToJson, jsonToSheet, aoaToSheet, sheetToCsv, sheetToHtml, StyleBuilder, RichTextBuilder |
| **Preservation** | Images/charts roundtrip via opaque blob passthrough |

### Not Implemented

| Priority | Features |
|---|---|
| **P0** | Tables/ListObjects, headers/footers, row/column grouping, print titles/areas, page margins (write) |
| **P1** | Sheet tab colors, workbook protection, split panes, file encryption |
| **P2** | Sparklines, data tables (what-if), chart creation, formula evaluation, pivot table creation, external links, custom XML, slicers/timelines |

### Test Coverage Gaps

Conditional formatting roundtrip, rich text cell roundtrip, hyperlink roundtrip, comments roundtrip, calc chain, theme colors, array/shared formulas, inline strings, gradient fills, DXF styles, diagonal borders, streaming, preserved entries

---

## 0.1.x — Hardening & Test Coverage

> Make what exists bulletproof. Zero new features — only correctness, coverage, and documentation.

### 0.1.1 — Style Roundtrip Tests

| Scope | Details |
|---|---|
| **Rust** | Bug fixes discovered during test writing |
| **TypeScript** | Bug fixes discovered during test writing |
| **Tests** | +15–20 new roundtrip tests |

- Gradient fill: write gradient → read back → verify type, stops, degree
- Diagonal border: write diagonal → read back → verify up/down/style/color
- DXF styles: write differential format → read back → verify font/fill/border overrides
- Cell styles (named styles): write xfId reference → read back → verify applied style
- Number format roundtrip: custom format string → write → read → verify formatCode preserved
- Alignment roundtrip: all 7 properties (horizontal, vertical, wrapText, textRotation, indent, shrinkToFit, readingOrder)
- Protection roundtrip: locked/hidden flags on individual cells
- Font roundtrip: bold, italic, underline, strikethrough, color, size, name, family, scheme

### 0.1.2 — Feature Roundtrip Tests

| Scope | Details |
|---|---|
| **Rust** | Bug fixes discovered during test writing |
| **TypeScript** | Bug fixes discovered during test writing |
| **Tests** | +15–20 new roundtrip tests |

- Conditional formatting: colorScale with 2-stop and 3-stop gradients
- Conditional formatting: dataBar with min/max settings
- Conditional formatting: iconSet with custom thresholds
- Rich text in cells: write RichTextBuilder output as cell value → read back → verify formatted runs preserved in SST
- Comments/notes roundtrip: write comment with author + text → read back → verify
- Hyperlink roundtrip: write external URL + tooltip → read back → verify
- Theme colors: read from real Excel file → verify parsed ARGB values

### 0.1.3 — Formula & Metadata Roundtrip Tests

| Scope | Details |
|---|---|
| **Rust** | Bug fixes discovered during test writing |
| **TypeScript** | Bug fixes discovered during test writing |
| **Tests** | +10–15 new roundtrip tests |

- Array formula roundtrip: write formula with formulaRef → read back → verify type + ref
- Shared formula roundtrip: write shared formula with si index → read back → verify
- Inline string roundtrip: write inline string cell → read back → verify value
- Calc chain roundtrip: write workbook with formulas → read back → verify calcChain entries
- Preserved entries roundtrip: write workbook with images/charts → read back → verify binary blobs unchanged
- Document properties roundtrip: write all Dublin Core fields → read back → verify

### 0.1.4 — Streaming & Parallel Tests

| Scope | Details |
|---|---|
| **Rust** | None |
| **TypeScript** | None |
| **Tests** | +10–15 new tests |

- Streaming reader: read 100K-row file via streaming API → verify row-by-row output matches standard reader
- Streaming writer: write 100K-row file via streaming API → read back with standard reader → verify
- Streaming roundtrip: stream-write → stream-read → compare
- Parallel parsing (rayon feature): read multi-sheet file → verify all sheets parse correctly
- Parallel vs. sequential: verify identical output regardless of parsing mode
- Benchmark baselines: standardized timing for 1K, 10K, 100K rows (read + write)

### 0.1.5 — Error Handling & Edge Cases (Part 1)

| Scope | Details |
|---|---|
| **Rust** | Improved error messages, graceful degradation |
| **TypeScript** | Clear error propagation from WASM |
| **Tests** | +10–15 new tests |

- Encrypted file detection: OLE2 compound document → throw descriptive error (not crash)
- Corrupted ZIP: truncated file, invalid headers → graceful XlsxError, not panic
- Malformed XML: unclosed tags, invalid attributes → best-effort parse with warnings
- Empty workbook (zero sheets): read and write without error
- Empty sheet (zero rows, zero cells): roundtrip preserves sheet metadata
- Sheet with only merged cells (no data cells): roundtrip preserves merges

### 0.1.6 — Error Handling & Edge Cases (Part 2)

| Scope | Details |
|---|---|
| **Rust** | Boundary validation, unicode handling |
| **TypeScript** | Input validation on public API |
| **Tests** | +10–15 new tests |

- Maximum dimensions: 1,048,576 rows × 16,384 columns (XFD) — verify cell ref encoding at limits
- Column letter edge cases: A, Z, AA, AZ, ZZ, AAA, XFD boundary
- Unicode in cell values: emoji, CJK characters, RTL text, combining characters
- Unicode in sheet names: verify XML escaping roundtrip
- Very large SST: 100K+ unique strings — verify no performance cliff
- Negative serial dates, zero serial, extremely large serial numbers
- NaN, Infinity, -Infinity as cell values — graceful handling

### 0.1.7 — TypeScript API Type Safety

| Scope | Details |
|---|---|
| **Rust** | None |
| **TypeScript** | Remove all `as` assertions, add validation |
| **Tests** | +5–10 new tests |

- Audit and remove every `as` type assertion — replace with type narrowing or runtime checks
- Add runtime schema validation at WASM boundary: validate JSON.parse output before casting to WorkbookData
- Strict null checks: audit all public API return types for possible undefined/null
- Input validation on all public setters: throw TypeError for invalid inputs instead of silently producing bad XLSX
- Exhaustive switch/case on CellType discriminated union (never default)
- Validate cell references in all APIs that accept A1-style strings

### 0.1.8 — JSDoc & API Documentation

| Scope | Details |
|---|---|
| **Rust** | None |
| **TypeScript** | 100% JSDoc coverage on public API |
| **Tests** | None |

- JSDoc on every exported function: description, @param, @returns, @throws, @example
- JSDoc on every exported type/interface: description, @property for each field
- JSDoc on every exported class method: description, @param, @returns, @example
- Generate API reference site (TypeDoc or similar)
- README quick-start section with copy-paste examples
- Add @since tags to all exports (0.1.0 for existing, future version for new)

### 0.1.9 — Migration Guides & Examples

| Scope | Details |
|---|---|
| **Rust** | None |
| **TypeScript** | Example files |
| **Tests** | None |

- Migration guide: SheetJS → modern-xlsx (function-by-function mapping table)
- Migration guide: ExcelJS → modern-xlsx (class/method mapping table)
- Usage example: read file → extract data → console output
- Usage example: create styled workbook from scratch
- Usage example: data validation + conditional formatting
- Usage example: browser usage (WASM init + Blob download)
- Usage example: streaming large file processing
- Performance comparison table vs. SheetJS, ExcelJS (read/write/bundle size)

---

## 0.2.x — Tables & Structured References

> Excel Tables (ListObjects) — the most-requested missing feature for data-heavy workbooks.

### 0.2.0 — Table Rust Types

| Scope | Details |
|---|---|
| **Rust** | New `ooxml/tables.rs` module with types |
| **TypeScript** | New type exports |
| **Tests** | +5 Rust unit tests |

- Define `TableData` struct: id, name, displayName, ref range, totalsRowShown
- Define `TableColumnData` struct: id, name, totalsRowFunction, calculatedColumnFormula
- Define `TableStyleInfoData` struct: name, showFirstColumn, showLastColumn, showRowStripes, showColumnStripes
- Define `TableAutoFilterData` for table-scoped auto-filter
- Add `tables: Vec<TableData>` to `WorksheetXml` Rust struct
- Export TypeScript types: `TableData`, `TableColumnData`, `TableStyleInfoData`
- Serde attributes: `#[serde(rename_all = "camelCase")]` on all new types

### 0.2.1 — Table XML Reader

| Scope | Details |
|---|---|
| **Rust** | SAX-style table XML parser |
| **TypeScript** | None |
| **Tests** | +10 Rust tests |

- Parse `xl/tables/table{n}.xml` parts from ZIP
- Extract table columns, style info, auto-filter, sort state
- Handle `totalsRowCount`, `headerRowCount` attributes
- Parse table relationships from `xl/worksheets/_rels/sheet{n}.xml.rels`
- Register table content types in content_types.rs
- Populate `WorksheetXml.tables` during sheet read
- Test: parse Excel-generated table XML → verify all fields

### 0.2.2 — Table Relationship Wiring

| Scope | Details |
|---|---|
| **Rust** | Relationship and content type management |
| **TypeScript** | None |
| **Tests** | +5 Rust tests |

- Auto-detect table parts in ZIP entries
- Build relationship entries for table ↔ worksheet links
- Register `application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml` override
- Handle multiple tables per sheet (separate relationship IDs)
- Handle table ID uniqueness across workbook
- Test: workbook with 3 sheets, 5 tables → verify all relationships correct

### 0.2.3 — TypeScript Table Read API

| Scope | Details |
|---|---|
| **Rust** | None |
| **TypeScript** | `Worksheet.tables` getter |
| **Tests** | +5–10 TypeScript roundtrip tests |

- `Worksheet.tables` → readonly array of table definitions
- `Worksheet.getTable(name)` → find table by display name
- Type-safe table column access
- Integration test: read Excel file with tables → verify TS API returns correct structure
- Test: table with calculated columns → verify formula preserved
- Test: table with totals row → verify function type preserved

### 0.2.4 — Table XML Writer

| Scope | Details |
|---|---|
| **Rust** | Table XML generation |
| **TypeScript** | None |
| **Tests** | +5–10 Rust tests |

- Generate `xl/tables/table{n}.xml` from `TableData`
- Write table columns with IDs, names, optional formulas
- Write `tableStyleInfo` element with all attributes
- Write auto-filter element scoped to table range
- Handle table ID allocation (unique across workbook)
- Test: write table XML → parse back → verify roundtrip

### 0.2.5 — Table Writer Integration

| Scope | Details |
|---|---|
| **Rust** | Writer pipeline integration |
| **TypeScript** | None |
| **Tests** | +5 Rust integration tests |

- Add table XML generation to main writer pipeline
- Generate table relationships in worksheet `.rels` files
- Add table content type overrides to `[Content_Types].xml`
- Handle `tableParts` element in worksheet XML (references to table parts)
- Test: full workbook write with tables → verify ZIP contains correct table parts
- Test: written file opens in Excel without repair dialog

### 0.2.6 — TypeScript Table Write API

| Scope | Details |
|---|---|
| **Rust** | None |
| **TypeScript** | `Worksheet.addTable()` method |
| **Tests** | +10 TypeScript tests |

- `Worksheet.addTable(ref, columns, options?)` → create a new table
- Options: name, displayName, style, showHeaderRow, showTotalsRow, showRowStripes, showColumnStripes
- Auto-generate column IDs and table ID
- Validate: ref range matches column count, no overlapping tables
- Test: create table → write → read back → verify all fields roundtrip
- Test: create multiple tables on same sheet → verify IDs unique

### 0.2.7 — Built-in Table Styles

| Scope | Details |
|---|---|
| **Rust** | Style name validation |
| **TypeScript** | Style constants and helpers |
| **Tests** | +5 tests |

- Export table style name constants: `TableStyleLight1`–`TableStyleLight21`, `TableStyleMedium1`–`TableStyleMedium28`, `TableStyleDark1`–`TableStyleDark11`
- Validate style name in `addTable()` — warn on unrecognized names
- Header row styling: bold font, bottom border (auto-applied by Excel)
- Totals row: SUM, COUNT, AVERAGE, MIN, MAX, STDDEV, VAR, COUNTNUMS functions
- Test: each style category (light, medium, dark) roundtrips correctly
- Test: totals row with different aggregate functions

### 0.2.8 — Calculated Columns & Structured References

| Scope | Details |
|---|---|
| **Rust** | Structured reference parsing in table context |
| **TypeScript** | Formula helpers |
| **Tests** | +5–10 tests |

- Parse structured references: `[@Column1]`, `[#Headers]`, `[#Totals]`, `[#Data]`, `[#All]`
- Write `calculatedColumnFormula` on table columns
- Preserve structured reference formulas through roundtrip
- TypeScript: `table.addCalculatedColumn(name, formula)` helper
- Test: table with `=[@Price]*[@Quantity]` → write → read → verify formula intact
- Test: structured reference syntax edge cases (special characters in column names)

### 0.2.9 — Table Utilities & Integration

| Scope | Details |
|---|---|
| **Rust** | None |
| **TypeScript** | Utility integrations |
| **Tests** | +5–10 tests |

- `Worksheet.removeTable(name)` — remove table definition, keep underlying cell data
- Table-aware `sheetToJson()` — use table column headers as keys when table spans the range
- `tableToJson(table)` — extract table data as array of objects
- `jsonToTable(data, options)` — create a table from JSON array
- Test: removeTable preserves cell values but removes table XML
- Test: sheetToJson detects table and uses headers automatically
- Test: jsonToTable end-to-end roundtrip

---

## 0.3.x — Print & Page Layout

> Everything needed to produce print-ready spreadsheets from code.

### 0.3.0 — Page Margins Read

| Scope | Details |
|---|---|
| **Rust** | Parse `<pageMargins>` element |
| **TypeScript** | Type exports |
| **Tests** | +5 Rust tests |

- Parse `<pageMargins>` attributes: top, bottom, left, right, header, footer (all in inches)
- Add `page_margins: Option<PageMarginsData>` to `WorksheetXml` (if not already present)
- Handle default margins when element is absent (Excel defaults: 0.75 top/bottom, 0.7 left/right, 0.3 header/footer)
- Test: read Excel file with custom margins → verify parsed values
- Test: read file with default margins → verify None (not present)

### 0.3.1 — Page Margins Write

| Scope | Details |
|---|---|
| **Rust** | Write `<pageMargins>` element |
| **TypeScript** | `Worksheet.pageMargins` getter/setter |
| **Tests** | +5–10 roundtrip tests |

- Generate `<pageMargins>` XML with all 6 attributes
- TypeScript: `Worksheet.pageMargins` getter returns `PageMarginsData | null`
- TypeScript: `Worksheet.pageMargins` setter accepts `PageMarginsData | null`
- Validate margin values: must be non-negative numbers
- Test: set all 6 margins → write → read back → verify values match
- Test: set margins to null → write → verify element omitted

### 0.3.2 — Headers & Footers Reader

| Scope | Details |
|---|---|
| **Rust** | Parse `<headerFooter>` element |
| **TypeScript** | Type exports |
| **Tests** | +5–10 Rust tests |

- Parse `<headerFooter>` with child elements: oddHeader, oddFooter, evenHeader, evenFooter, firstHeader, firstFooter
- Parse `differentOddEven` and `differentFirst` attributes
- Define `HeaderFooterData` struct with all 6 sections + boolean flags
- Handle Excel formatting codes within header/footer text: `&L` (left), `&C` (center), `&R` (right)
- Test: read file with complex header/footer → verify all sections parsed
- Test: read file with no header/footer → verify None

### 0.3.3 — Headers & Footers Writer

| Scope | Details |
|---|---|
| **Rust** | Write `<headerFooter>` element |
| **TypeScript** | `Worksheet.headerFooter` getter/setter |
| **Tests** | +5–10 roundtrip tests |

- Generate `<headerFooter>` XML with conditional child elements
- Only write child elements that have content (skip empty strings)
- TypeScript: `Worksheet.headerFooter` getter/setter with typed `HeaderFooterData`
- Test: set odd header/footer only → write → read back → verify
- Test: set all 6 sections with different/first flags → roundtrip
- Test: set header to null → verify element omitted from output

### 0.3.4 — Header/Footer Formatting Codes

| Scope | Details |
|---|---|
| **Rust** | None (codes are plain text, preserved as-is) |
| **TypeScript** | Helper builders |
| **Tests** | +5 TypeScript tests |

- TypeScript helper: `HeaderFooterBuilder` with fluent API
- Methods: `.left(text)`, `.center(text)`, `.right(text)`
- Built-in codes: `.pageNumber()` → `&P`, `.totalPages()` → `&N`, `.date()` → `&D`, `.time()` → `&T`, `.fileName()` → `&F`, `.sheetName()` → `&A`, `.filePath()` → `&Z`
- Font formatting: `.bold(text)` → `&B{text}&B`, `.italic(text)` → `&I{text}&I`, `.fontSize(n)` → `&{n}`
- Test: build complex header with page number and date → verify output string
- Test: builder output roundtrips through write → read

### 0.3.5 — Print Titles

| Scope | Details |
|---|---|
| **Rust** | Parse/write `_xlnm.Print_Titles` defined name |
| **TypeScript** | `Worksheet.printTitles` getter/setter |
| **Tests** | +5–10 roundtrip tests |

- Parse `<definedName name="_xlnm.Print_Titles">` from workbook.xml
- Format: `Sheet1!$1:$3` (repeat rows) or `Sheet1!$A:$B` (repeat columns) or combined
- Write print titles as defined name during workbook serialization
- TypeScript: `Worksheet.printTitles` → `{ rows?: [start, end], columns?: [start, end] } | null`
- Test: set repeat rows 1–3 → write → read → verify
- Test: set repeat columns A–B → write → read → verify
- Test: set both rows and columns → roundtrip

### 0.3.6 — Print Areas

| Scope | Details |
|---|---|
| **Rust** | Parse/write `_xlnm.Print_Area` defined name |
| **TypeScript** | `Worksheet.printArea` getter/setter |
| **Tests** | +5 roundtrip tests |

- Parse `<definedName name="_xlnm.Print_Area">` from workbook.xml
- Format: `Sheet1!$A$1:$H$50` (single range) or `Sheet1!$A$1:$D$25,Sheet1!$F$1:$H$25` (multiple)
- Write print area as defined name during workbook serialization
- TypeScript: `Worksheet.printArea` → `string | null` (A1-style range reference)
- Test: single print area range → roundtrip
- Test: multiple print area ranges (comma-separated) → roundtrip
- Test: clear print area (set null) → verify defined name removed

### 0.3.7 — Row Grouping (Outline Levels)

| Scope | Details |
|---|---|
| **Rust** | Parse/write `outlineLevel` on `<row>` elements |
| **TypeScript** | `Worksheet.groupRows()` method |
| **Tests** | +10 roundtrip tests |

- Parse `outlineLevel` attribute on `<row>` elements (1–7)
- Parse `collapsed` and `hidden` attributes for group state
- Write `outlineLevel`, `collapsed`, `hidden` on row elements
- Parse/write `<sheetFormatPr outlineLevelRow>` for max outline level
- Parse/write `<sheetPr><outlinePr summaryBelow>` for summary row position
- TypeScript: `Worksheet.groupRows(startRow, endRow, options?)` — options: level (default 1), collapsed
- Test: group rows 2–5 at level 1 → roundtrip
- Test: nested groups (rows 2–10 level 1, rows 3–5 level 2) → roundtrip
- Test: collapsed group → verify hidden + collapsed attributes

### 0.3.8 — Column Grouping (Outline Levels)

| Scope | Details |
|---|---|
| **Rust** | Parse/write `outlineLevel` on `<col>` elements |
| **TypeScript** | `Worksheet.groupColumns()` method |
| **Tests** | +10 roundtrip tests |

- Parse `outlineLevel` attribute on `<col>` elements (1–7)
- Parse `collapsed` and `hidden` attributes for group state
- Write `outlineLevel`, `collapsed`, `hidden` on col elements
- Parse/write `<sheetFormatPr outlineLevelCol>` for max outline level
- Parse/write `<sheetPr><outlinePr summaryRight>` for summary column position
- TypeScript: `Worksheet.groupColumns(startCol, endCol, options?)` — options: level (default 1), collapsed
- Test: group columns B–D at level 1 → roundtrip
- Test: nested column groups → roundtrip
- Test: collapsed column group → verify attributes

### 0.3.9 — Sheet Tab Colors

| Scope | Details |
|---|---|
| **Rust** | Parse/write `<sheetPr><tabColor>` element |
| **TypeScript** | `Worksheet.tabColor` getter/setter |
| **Tests** | +5 roundtrip tests |

- Parse `<tabColor>` element within `<sheetPr>`: rgb, theme, indexed, tint attributes
- Write `<tabColor>` during sheet serialization
- Ensure `<sheetPr>` is created if absent when tabColor is set
- TypeScript: `Worksheet.tabColor` → `{ rgb?: string; theme?: number; indexed?: number; tint?: number } | null`
- Test: set RGB tab color → roundtrip
- Test: set theme color with tint → roundtrip
- Test: set indexed color → roundtrip

---

## 0.4.x — Advanced Sheet Features

> Pane management, workbook protection, view configuration, and sheet management.

### 0.4.0 — Split Panes (Horizontal)

| Scope | Details |
|---|---|
| **Rust** | Extend `<pane>` parsing for `split` state |
| **TypeScript** | `Worksheet.splitPane` getter/setter |
| **Tests** | +5 roundtrip tests |

- Parse `<pane>` with `state="split"` (distinct from `state="frozen"`)
- `xSplit`/`ySplit` values in twips (1/20th of a point) for split mode, not cell coordinates
- Parse `topLeftCell` for bottom-right pane quadrant
- Write split pane XML — ensure state attribute is `split` not `frozen`
- TypeScript: `Worksheet.splitPane` → `{ horizontal?: number; vertical?: number; topLeftCell?: string } | null`
- Test: horizontal split → write → read → verify twip value preserved

### 0.4.1 — Split Panes (Vertical & Four-Way)

| Scope | Details |
|---|---|
| **Rust** | Complete split pane support |
| **TypeScript** | API extension |
| **Tests** | +5 roundtrip tests |

- Vertical split: `xSplit` in twips with `ySplit` = 0
- Four-way split: both `xSplit` and `ySplit` non-zero
- Parse `activePane` attribute: `topLeft`, `topRight`, `bottomLeft`, `bottomRight`
- Parse `<selection>` elements per pane (each pane has its own active cell)
- Test: vertical split → roundtrip
- Test: four-way split → verify all 4 pane selections preserved
- Test: split pane coexistence with frozen pane is mutually exclusive — verify only one applies

### 0.4.2 — Sheet View Attributes

| Scope | Details |
|---|---|
| **Rust** | Parse/write `<sheetView>` attributes |
| **TypeScript** | `Worksheet.view` getter/setter |
| **Tests** | +5–10 roundtrip tests |

- Parse boolean attributes: `showGridLines`, `showRowColHeaders`, `showZeros`, `rightToLeft`, `showRuler`, `showOutlineSymbols`, `showWhiteSpace`
- Parse numeric attributes: `zoomScale` (10–400), `zoomScaleNormal`, `zoomScalePageLayoutView`, `zoomScaleSheetLayoutView`
- Write all attributes during sheet serialization
- TypeScript: `Worksheet.view` → `SheetViewData` with all boolean/numeric fields
- Test: hide gridlines + set zoom to 150% → roundtrip
- Test: RTL view → roundtrip
- Test: default view (no explicit attributes) → verify defaults applied

### 0.4.3 — Sheet View Modes

| Scope | Details |
|---|---|
| **Rust** | Parse/write `view` attribute on `<sheetView>` |
| **TypeScript** | View mode enum |
| **Tests** | +5 roundtrip tests |

- Parse `view` attribute: `normal` (default), `pageBreakPreview`, `pageLayout`
- Parse `colorId` (legacy) and `defaultGridColor` attributes
- Write view mode during serialization
- Handle relationship between view mode and zoom attributes (each mode has its own zoom)
- TypeScript: `Worksheet.viewMode` → `'normal' | 'pageBreakPreview' | 'pageLayout'`
- Test: page break preview mode → roundtrip
- Test: page layout mode with custom zoom → roundtrip

### 0.4.4 — Workbook Protection (Structure & Windows)

| Scope | Details |
|---|---|
| **Rust** | Parse/write `<workbookProtection>` element |
| **TypeScript** | `Workbook.protection` getter/setter |
| **Tests** | +5 roundtrip tests |

- Parse `<workbookProtection>` attributes: `lockStructure`, `lockWindows`, `lockRevision`
- Parse `revisionsAlgorithmName`, `revisionsHashValue`, `revisionsSaltValue`, `revisionsSpinCount`
- Write `<workbookProtection>` in workbook.xml after `<sheets>` element
- TypeScript: `Workbook.protection` → `WorkbookProtectionData | null`
- Test: lock structure only → roundtrip
- Test: lock structure + windows → roundtrip
- Test: clear protection (set null) → verify element removed

### 0.4.5 — Workbook Protection (Password)

| Scope | Details |
|---|---|
| **Rust** | Password hash algorithm |
| **TypeScript** | Password option in protection setter |
| **Tests** | +5 roundtrip tests |

- Implement Excel password hash: SHA-512 with salt and spin count (same as sheet protection)
- Parse `workbookAlgorithmName`, `workbookHashValue`, `workbookSaltValue`, `workbookSpinCount` attributes
- Write hashed password attributes when password provided
- TypeScript: `Workbook.protection = { lockStructure: true, password: '...' }`
- Test: set password protection → write → read → verify hash attributes present
- Test: different passwords produce different hashes
- Test: protection without password → verify no hash attributes

### 0.4.6 — Sheet State & Visibility

| Scope | Details |
|---|---|
| **Rust** | Parse/write `state` attribute on `<sheet>` element |
| **TypeScript** | `Worksheet.state` getter/setter |
| **Tests** | +5 roundtrip tests |

- Parse `state` attribute on `<sheet>` in workbook.xml: `visible` (default), `hidden`, `veryHidden`
- Write state attribute during workbook serialization (omit for `visible` since it's the default)
- Validation: at least one sheet must remain visible — throw error if hiding last visible sheet
- TypeScript: `Worksheet.state` → `'visible' | 'hidden' | 'veryHidden'`
- Test: hide sheet → roundtrip → verify state preserved
- Test: veryHidden → roundtrip
- Test: attempt to hide all sheets → expect error

### 0.4.7 — Sheet Ordering & Management

| Scope | Details |
|---|---|
| **Rust** | None |
| **TypeScript** | Sheet management methods |
| **Tests** | +5–10 TypeScript tests |

- `Workbook.moveSheet(fromIndex, toIndex)` — reorder sheets
- `Workbook.cloneSheet(index, newName?)` — deep copy sheet with unique name
- `Workbook.renameSheet(index, newName)` — rename with uniqueness check
- Update all defined names, print titles, print areas that reference moved/renamed sheets
- Validate sheet name: max 31 chars, no `[]:*?/\` characters
- Test: move sheet from position 0 to 2 → verify order
- Test: clone sheet → verify data independent (modify clone, original unchanged)
- Test: rename sheet → verify defined names updated

### 0.4.8 — Custom Sheet Properties

| Scope | Details |
|---|---|
| **Rust** | Parse/write `<sheetPr>` attributes |
| **TypeScript** | Property accessors |
| **Tests** | +5 roundtrip tests |

- Parse `<sheetPr>` attributes: `codeName`, `enableFormatConditionsCalculation`, `filterMode`, `published`, `syncHorizontal`, `syncVertical`, `transitionEntry`, `transitionEvaluation`
- Parse `<pageSetUpPr>` child: `autoPageBreaks`, `fitToPage`
- Write all attributes during serialization
- TypeScript: `Worksheet.properties` → `SheetPropertiesData`
- Test: codeName preservation → roundtrip
- Test: fitToPage combined with page setup scale → roundtrip

### 0.4.9 — View & Layout Integration Tests

| Scope | Details |
|---|---|
| **Rust** | None |
| **TypeScript** | None |
| **Tests** | +10 integration tests |

- Full integration: frozen pane + hidden gridlines + zoom + page layout mode → single sheet roundtrip
- Full integration: workbook protection + hidden sheets + sheet protection → multi-sheet roundtrip
- Full integration: page margins + headers/footers + print titles + print area → print-ready sheet roundtrip
- Full integration: row grouping + column grouping + tab colors → grouped sheet roundtrip
- Cross-feature: split pane on one sheet, frozen pane on another → same workbook roundtrip
- Compatibility: written files open in Excel, Google Sheets, LibreOffice without warnings
- Performance: view/layout features do not regress read/write performance (benchmark comparison)

---

## 0.5.x — Data Features

> Sparklines, data tables, external links, and custom XML for data-rich workbooks.

### 0.5.0 — Sparkline Types & Model

| Scope | Details |
|---|---|
| **Rust** | Sparkline data structures |
| **TypeScript** | Type exports |
| **Tests** | +5 Rust unit tests |

- Define `SparklineGroupData` struct: type (line/column/winLoss), colorSeries, colorNegative, colorAxis, colorMarkers, colorFirst, colorLast, colorHigh, colorLow
- Define `SparklineData` struct: formula (data range), sqref (display cell)
- Add `sparkline_groups: Vec<SparklineGroupData>` to `WorksheetXml`
- Axis settings: minAxisType, maxAxisType (individual/group/custom), manualMin, manualMax
- Display options: displayEmptyCellsAs (gap/zero/span), rightToLeft, displayHidden
- Serde attributes and TypeScript type exports

### 0.5.1 — Sparkline Reader

| Scope | Details |
|---|---|
| **Rust** | Parse `<x14:sparklineGroups>` in `<extLst>` |
| **TypeScript** | None |
| **Tests** | +10 Rust tests |

- Parse `<ext uri="{05C60535-1F16-4fd2-B633-F4F36F0B64E0}">` extension block
- Parse `<x14:sparklineGroup>` elements with all attributes
- Parse `<x14:sparklines>` child elements (formula + sqref pairs)
- Parse color elements (theme, rgb, indexed references)
- Handle xmlns:x14 namespace declaration
- Test: line sparkline group → verify parsed fields
- Test: column sparkline group with negative colors → verify
- Test: multiple sparkline groups per sheet → verify all parsed

### 0.5.2 — Sparkline Writer

| Scope | Details |
|---|---|
| **Rust** | Generate sparkline XML in extension list |
| **TypeScript** | None |
| **Tests** | +5–10 Rust tests |

- Write `<extLst>` with sparkline extension block
- Generate namespace declarations for x14
- Write sparkline groups with all attributes and child elements
- Handle empty sparkline groups (omit extension block entirely)
- Merge with other extension list entries (preserve existing `<ext>` blocks)
- Test: write sparkline → parse back → roundtrip verification
- Test: sparkline alongside existing extLst entries → both preserved

### 0.5.3 — Sparkline TypeScript API

| Scope | Details |
|---|---|
| **Rust** | None |
| **TypeScript** | `Worksheet.addSparkline()` method |
| **Tests** | +10 TypeScript roundtrip tests |

- `Worksheet.addSparkline(type, dataRange, locationCell, options?)` → create sparkline
- `Worksheet.sparklines` → readonly array of sparkline group definitions
- Options: all color overrides, axis settings, display options
- Validate data range and location cell references
- Test: create line sparkline → write → read → verify
- Test: create column sparkline with custom colors → roundtrip
- Test: create win/loss sparkline → roundtrip
- Test: multiple sparklines on same sheet → roundtrip

### 0.5.4 — Sparkline Formatting & Styles

| Scope | Details |
|---|---|
| **Rust** | None |
| **TypeScript** | Style helpers |
| **Tests** | +5 tests |

- Sparkline style presets (matching Excel's built-in sparkline styles)
- Helper: `SparklineBuilder` with fluent API for constructing sparkline options
- Marker toggles: showMarkers, showFirst, showLast, showHigh, showLow, showNegative
- Line weight: `lineWeight` attribute (in points, default 0.75)
- Test: style presets produce expected color combinations
- Test: marker toggles roundtrip correctly
- Test: custom line weight roundtrip

### 0.5.5 — Data Tables (What-If) Reader

| Scope | Details |
|---|---|
| **Rust** | Parse `<dataConsolidate>` and what-if structures |
| **TypeScript** | Read-only type exports |
| **Tests** | +5 Rust tests |

- Parse one-variable data tables: `r1` (row input cell) and `r2` (column input cell) attributes on `<f>` elements
- Parse two-variable data tables: both `r1` and `r2` populated
- Differentiate from ListObject tables (different XML structure entirely)
- Define `DataTableData` struct: inputCell1, inputCell2, ref range
- Test: read file with one-variable data table → verify
- Test: read file with two-variable data table → verify

### 0.5.6 — Data Tables (What-If) Writer

| Scope | Details |
|---|---|
| **Rust** | Write data table formula references |
| **TypeScript** | Read-only (preservation, not creation) |
| **Tests** | +5 roundtrip tests |

- Write data table formula references with `r1`/`r2` attributes
- Preserve data table structure through read → write cycle
- TypeScript: `Worksheet.dataTables` read-only getter for inspection
- Test: read file with data table → write → read → verify structure intact
- Test: data table cells preserve calculated values

### 0.5.7 — External Links Reader

| Scope | Details |
|---|---|
| **Rust** | Parse `xl/externalLinks/externalLink{n}.xml` |
| **TypeScript** | Read-only type exports |
| **Tests** | +5 Rust tests |

- Parse external link XML parts from ZIP
- Extract external workbook file path, sheet names, defined names
- Parse external link relationships
- Parse cached cell values (for offline display when linked file unavailable)
- Define `ExternalLinkData` struct: fileName, sheetNames, definedNames, cachedValues
- Test: read file with external workbook reference → verify parsed path + cache
- Test: read file with multiple external links → verify all parsed

### 0.5.8 — External Links Writer & Custom XML

| Scope | Details |
|---|---|
| **Rust** | Preserve external links + parse custom XML |
| **TypeScript** | Read-only getters |
| **Tests** | +5–10 roundtrip tests |

- Write external link XML parts back during save (preservation)
- Parse `customXml/` entries and their relationships
- Preserve custom XML parts through roundtrip (opaque passthrough, like preserved_entries)
- Register content types for external link and custom XML parts
- TypeScript: `Workbook.externalLinks` and `Workbook.customXml` read-only getters
- Test: external links survive roundtrip
- Test: custom XML parts survive roundtrip

### 0.5.9 — Data Feature Integration Tests

| Scope | Details |
|---|---|
| **Rust** | None |
| **TypeScript** | None |
| **Tests** | +10 integration tests |

- Full integration: sparklines + data tables + external links in same workbook → roundtrip
- Sparkline referencing data from another sheet → verify reference preserved
- Data table within a ListObject table → both preserved correctly
- External links with broken file paths → graceful handling (warn, don't error)
- Custom XML with custom namespace → roundtrip preservation
- Performance: sparkline/data features don't regress read/write speed
- Compatibility: written files open in Excel without warnings

---

## 0.6.x — Security & Encryption

> File-level security for enterprise deployments.

### 0.6.0 — OLE2 Compound Document Detection

| Scope | Details |
|---|---|
| **Rust** | OLE2 magic byte detection |
| **TypeScript** | Descriptive error message |
| **Tests** | +5 tests |

- Detect OLE2 header magic bytes: `D0 CF 11 E0 A1 B1 1A E1`
- Distinguish from ZIP header (`PK\x03\x04`)
- Return descriptive error: "This file is password-protected (OLE2 compound document). Decryption not yet supported in this version."
- Also detect: legacy .xls format (OLE2 without encryption) → "Legacy .xls format not supported. Convert to .xlsx first."
- Test: encrypted XLSX → descriptive error
- Test: legacy .xls → appropriate error

### 0.6.1 — Encryption Method Identification

| Scope | Details |
|---|---|
| **Rust** | Parse OLE2 directory + EncryptionInfo stream |
| **TypeScript** | Encryption info in error message |
| **Tests** | +5 tests |

- Parse OLE2 compound document directory structure
- Locate and read `EncryptionInfo` stream
- Identify encryption version: Standard (2.3.6.1), Agile (2.3.6.2), Extensible (2.3.6.3)
- For Agile: parse XML descriptor → extract algorithm (AES-128, AES-256), hash (SHA-1, SHA-256, SHA-384, SHA-512), salt, spin count
- Test: AES-128 encrypted file → identify method correctly
- Test: AES-256 encrypted file → identify method correctly
- Test: Standard encryption → identify version

### 0.6.2 — Key Derivation (SHA-512)

| Scope | Details |
|---|---|
| **Rust** | ECMA-376 key derivation implementation |
| **TypeScript** | None |
| **Tests** | +5 Rust tests |

- Implement password-to-key derivation per ECMA-376 2.3.6.2
- UTF-16LE password encoding
- SHA-512 iterative hashing with salt and spin count (default 100,000)
- Derive encryption key, verification key, integrity key from password
- Block-based key derivation for AES-CBC segments
- Test: known password → expected derived key (test vectors from spec)
- Test: empty password → valid derived key
- Test: unicode password (non-ASCII characters) → valid derived key

### 0.6.3 — AES-256-CBC Decryption

| Scope | Details |
|---|---|
| **Rust** | AES-256-CBC decrypt + HMAC verify |
| **TypeScript** | None |
| **Tests** | +5 Rust tests |

- AES-256-CBC decryption with derived key
- PKCS#7 padding removal and validation
- HMAC-SHA-512 integrity verification (compare computed vs. stored hash)
- Decrypt `EncryptedPackage` stream → extract original ZIP bytes
- Handle block boundaries: 4096-byte segments, each with its own IV derived from block index
- Test: decrypt known test file → verify HMAC passes
- Test: wrong key → HMAC verification fails gracefully
- Test: corrupted encrypted data → clear error

### 0.6.4 — File Decryption Integration (Read Path)

| Scope | Details |
|---|---|
| **Rust** | Integrate decryption into reader pipeline |
| **TypeScript** | `password` option on `readBuffer()` |
| **Tests** | +5 integration tests |

- When OLE2 detected and password provided: decrypt → extract ZIP → proceed with normal reader
- When OLE2 detected and no password: throw descriptive error with encryption method info
- Pipeline: OLE2 parse → EncryptionInfo → derive key → decrypt → ZIP → normal read
- TypeScript: `readBuffer(data, { password: '...' })` option
- Test: read password-protected file with correct password → get Workbook
- Test: read with wrong password → clear "incorrect password" error
- Test: read with no password → clear "password required" error

### 0.6.5 — Decryption Compatibility & Edge Cases

| Scope | Details |
|---|---|
| **Rust** | Edge case handling |
| **TypeScript** | None |
| **Tests** | +5 tests |

- Handle Standard encryption (older method, pre-Agile) — AES-128 or RC4
- Handle files encrypted by LibreOffice (may use different defaults)
- Handle files encrypted by Google Sheets export
- Handle read-only recommended files (not encrypted, just a flag)
- Handle files with both workbook protection AND file encryption
- Test: LibreOffice-encrypted file → successful decrypt
- Test: read-only recommended file without password → normal read

### 0.6.6 — OLE2 Compound Document Writer

| Scope | Details |
|---|---|
| **Rust** | OLE2 binary format writer |
| **TypeScript** | None |
| **Tests** | +5 Rust tests |

- Generate OLE2 compound document structure:
  - Header (512 bytes): magic, sector size, directory chain
  - Directory entries: Root, EncryptionInfo, EncryptedPackage
  - FAT (file allocation table) chains
  - Mini FAT for small streams
- Write EncryptionInfo stream with Agile encryption XML descriptor
- Test: generated OLE2 → re-parse directory → verify structure valid
- Test: sector chain integrity for large encrypted packages

### 0.6.7 — File Encryption (Write Path)

| Scope | Details |
|---|---|
| **Rust** | Encrypt ZIP into OLE2 |
| **TypeScript** | `password` option on write |
| **Tests** | +5 integration tests |

- Pipeline: serialize WorkbookData → ZIP bytes → AES-256-CBC encrypt → HMAC → OLE2 package
- Generate random salt, compute spin count (100,000 default)
- Derive encryption, verification, and integrity keys
- Encrypt in 4096-byte blocks with per-block IV
- Compute and store HMAC-SHA-512 integrity hash
- TypeScript: `Workbook.toBuffer({ password: '...' })` and `writeBlob(wb, { password: '...' })`
- Test: encrypt → decrypt → verify roundtrip produces identical workbook

### 0.6.8 — Encryption Roundtrip & Compatibility

| Scope | Details |
|---|---|
| **Rust** | None |
| **TypeScript** | None |
| **Tests** | +10 roundtrip/compat tests |

- Encrypt with modern-xlsx → open in Excel → verify no repair dialog
- Encrypt with modern-xlsx → decrypt with modern-xlsx → verify data integrity
- Encrypt with Excel → decrypt with modern-xlsx → verify data integrity
- Password change: read with old password → write with new password → read with new password
- Large file encryption: 100K rows → verify performance acceptable
- Test: encrypted file with styles, formulas, tables → full feature roundtrip
- Test: empty password (not the same as no password) → roundtrip

### 0.6.9 — Security Hardening

| Scope | Details |
|---|---|
| **Rust** | Security audit |
| **TypeScript** | None |
| **Tests** | +5 security tests |

- Audit: password material never logged (even at debug level)
- Audit: derived keys zeroed from memory after use (zeroize crate)
- Audit: salt generation uses cryptographically secure PRNG (getrandom crate)
- Audit: timing-safe comparison for HMAC verification (constant_time_eq)
- Audit: no information leakage in error messages (don't reveal partial decryption state)
- Fuzz testing: malformed EncryptionInfo → no panics
- Fuzz testing: truncated EncryptedPackage → graceful error

---

## 0.7.x — Formulas & Calculation

> Formula intelligence — parsing, rewriting, and basic evaluation.

### 0.7.0 — Formula Tokenizer (Lexer)

| Scope | Details |
|---|---|
| **Rust** | Formula token types and lexer |
| **TypeScript** | None |
| **Tests** | +10 Rust tests |

- Define token enum: CellRef, RangeRef, NamedRef, Number, String, Bool, Error, Function, Operator, Paren, Comma, Colon, Semicolon, Whitespace
- Single-pass byte-level tokenizer (efficient, no regex)
- Handle: `$A$1`, `A1:B5`, `Sheet1!A1`, `'Sheet Name'!A1:B5`, `[1]Sheet1!A1` (external)
- Handle: `#REF!`, `#VALUE!`, `#NAME?`, `#NULL!`, `#N/A`, `#DIV/0!`, `#NUM!`
- Handle: `{1,2;3,4}` array constants
- Test: tokenize `=SUM(A1:B5)+IF(C1>0,D1,"N/A")` → verify token stream
- Test: tokenize formula with sheet reference, error values, array constants

### 0.7.1 — Formula Parser (AST)

| Scope | Details |
|---|---|
| **Rust** | Recursive descent parser |
| **TypeScript** | None |
| **Tests** | +10 Rust tests |

- Define AST node types: BinaryOp, UnaryOp, FunctionCall, CellRef, RangeRef, Literal, Array, Parenthesized
- Operator precedence: `^` > `*` `/` > `+` `-` > `&` > `=` `<>` `<` `>` `<=` `>=`
- Handle: nested function calls, mixed references, implicit intersection (`@`)
- Handle: R1C1 notation (R[-1]C[0]) alongside A1 notation
- Parse error recovery: return partial AST with error nodes for malformed formulas
- Test: complex nested formula → AST → verify structure
- Test: operator precedence edge cases
- Test: malformed formula → partial AST + error

### 0.7.2 — Formula Serializer (AST → String)

| Scope | Details |
|---|---|
| **Rust** | AST to formula string conversion |
| **TypeScript** | `parseFormula()` and `formulaToString()` exports |
| **Tests** | +5–10 roundtrip tests |

- Convert AST back to Excel formula string (inverse of parser)
- Minimal parentheses: only add parens where precedence requires them
- Preserve original formatting where possible (whitespace, case)
- TypeScript WASM exports: `parseFormula(formula: string)` → AST, `formulaToString(ast)` → string
- Test: parse → serialize → compare to original for 50+ formulas
- Test: serialize → parse → serialize → idempotent

### 0.7.3 — Reference Resolver

| Scope | Details |
|---|---|
| **Rust** | Cell reference extraction and classification |
| **TypeScript** | Reference utility functions |
| **Tests** | +10 tests |

- Extract all cell/range references from a formula AST
- Classify each reference: absolute row, absolute column, relative row, relative column, mixed
- Resolve named ranges to their target references
- Cross-sheet reference detection and resolution
- Dependency graph: given a cell, return all cells it depends on (direct dependencies)
- TypeScript: `getFormulaDependencies(formula, sheetName)` → array of cell references
- Test: formula with mixed absolute/relative refs → correct classification
- Test: dependency extraction for complex formulas

### 0.7.4 — Reference Rewriter (Insert/Delete Rows)

| Scope | Details |
|---|---|
| **Rust** | Row-based reference shifting |
| **TypeScript** | None |
| **Tests** | +10 Rust tests |

- Shift row references when rows are inserted above
- Shift row references when rows are deleted
- Handle absolute rows (don't shift `$5`)
- Handle ranges that expand (insert within range → range grows)
- Handle ranges that contract (delete within range → range shrinks)
- Handle references that become invalid (deleted range → `#REF!`)
- Test: insert row at 3 → `A5` becomes `A6`, `$A$5` becomes `$A$6`, `A$5` row stays
- Test: delete row 3 → `A5` becomes `A4`
- Test: delete row within range → range contracts

### 0.7.5 — Reference Rewriter (Insert/Delete Columns)

| Scope | Details |
|---|---|
| **Rust** | Column-based reference shifting |
| **TypeScript** | `Worksheet.insertRows()` and `Worksheet.insertColumns()` |
| **Tests** | +10 tests |

- Shift column references when columns are inserted/deleted
- Same logic as rows but for column letters (handle A→Z→AA transitions)
- Handle absolute columns (don't shift `$C`)
- TypeScript API: `Worksheet.insertRows(at, count)` — inserts blank rows and shifts all formulas
- TypeScript API: `Worksheet.deleteRows(at, count)` — deletes rows and shifts formulas
- TypeScript API: `Worksheet.insertColumns(at, count)` and `Worksheet.deleteColumns(at, count)`
- Test: insert column at C → `D1` becomes `E1`, `$D$1` stays
- Test: end-to-end: insert row via API → all formulas on sheet update correctly

### 0.7.6 — Shared Formula Expansion

| Scope | Details |
|---|---|
| **Rust** | Expand `si` shared formulas to explicit formulas |
| **TypeScript** | None |
| **Tests** | +5 Rust tests |

- When shared formula master cell (has `ref` + `si` + formula text) encountered, store template
- For child cells (have `si` only, no formula text), compute explicit formula by shifting references
- Option to expand all shared formulas to explicit during read (normalizes the data)
- Option to re-share common formulas during write (compression, reduces file size)
- Test: shared formula across 100 rows → expand → verify each cell has correct shifted formula
- Test: expand → re-share → verify `si` indices match original

### 0.7.7 — Arithmetic & Comparison Evaluation

| Scope | Details |
|---|---|
| **Rust** | Basic expression evaluator |
| **TypeScript** | `evaluateFormula()` export |
| **Tests** | +15 tests |

- Evaluate arithmetic operators: `+`, `-`, `*`, `/`, `^` (power)
- Evaluate unary: `-` (negation), `+` (no-op)
- Evaluate comparison: `=`, `<>`, `<`, `>`, `<=`, `>=` → returns TRUE/FALSE
- Evaluate string concatenation: `&`
- Cell value lookup from worksheet data (resolve `A1` to its value)
- Type coercion rules: string-to-number in arithmetic context, number-to-string in `&` context
- Error propagation: `#DIV/0!` from division by zero, `#VALUE!` from type mismatch
- TypeScript: `evaluateFormula(formula, worksheet)` → computed value
- Test: `=1+2*3` → 7, `=2^10` → 1024, `="Hello"&" "&"World"` → "Hello World"

### 0.7.8 — String & Logical Function Evaluation

| Scope | Details |
|---|---|
| **Rust** | String and logical function implementations |
| **TypeScript** | None |
| **Tests** | +15 tests |

- String functions: `LEN`, `LEFT`, `RIGHT`, `MID`, `TRIM`, `UPPER`, `LOWER`, `PROPER`, `CONCATENATE`, `TEXTJOIN`, `SUBSTITUTE`, `REPLACE`, `FIND`, `SEARCH`, `REPT`, `TEXT`, `VALUE`
- Logical functions: `IF`, `AND`, `OR`, `NOT`, `TRUE`, `FALSE`, `IFERROR`, `IFNA`, `IFS`, `SWITCH`
- IS-functions: `ISBLANK`, `ISERROR`, `ISNA`, `ISNUMBER`, `ISTEXT`, `ISLOGICAL`
- Type conversion: `N`, `T`, `TYPE`
- Test: each function with normal input, edge cases, error handling
- Test: nested `IF(AND(A1>0, B1<10), "Yes", "No")` → correct evaluation

### 0.7.9 — Math, Statistical & Lookup Evaluation

| Scope | Details |
|---|---|
| **Rust** | Math/stat/lookup function implementations |
| **TypeScript** | None |
| **Tests** | +20 tests |

- Math: `SUM`, `SUMIF`, `SUMIFS`, `SUMPRODUCT`, `PRODUCT`, `ABS`, `ROUND`, `ROUNDUP`, `ROUNDDOWN`, `CEILING`, `FLOOR`, `MOD`, `INT`, `SIGN`, `SQRT`, `POWER`, `LOG`, `LOG10`, `LN`, `EXP`, `PI`, `RAND`
- Statistical: `AVERAGE`, `AVERAGEIF`, `AVERAGEIFS`, `MIN`, `MAX`, `COUNT`, `COUNTA`, `COUNTBLANK`, `COUNTIF`, `COUNTIFS`, `MEDIAN`, `LARGE`, `SMALL`
- Lookup: `VLOOKUP`, `HLOOKUP`, `INDEX`, `MATCH`, `XLOOKUP`, `CHOOSE`
- Range evaluation: SUM(A1:B5) iterates all cells in range
- Test: `SUM(A1:A10)` with mixed types (numbers, blanks, strings) → correct sum
- Test: `VLOOKUP` exact and approximate match modes
- Test: `XLOOKUP` with match mode and search mode options
- Test: circular reference detection → `#REF!` error (no infinite loop)

---

## 0.8.x — Charts & Visualizations

> From chart preservation to chart creation — DrawingML generation.

### 0.8.0 — Chart Data Model

| Scope | Details |
|---|---|
| **Rust** | Chart type hierarchy |
| **TypeScript** | Type exports |
| **Tests** | +5 Rust unit tests |

- Define chart type enum: Bar, Column, Line, Pie, Scatter, Area, Doughnut, Radar, Bubble, Stock
- Define `ChartData` struct: type, title, series, axes, legend, plotArea, style
- Define `ChartSeriesData`: name, categories (cat), values (val), formatting
- Define `ChartAxisData`: id, title, numFmt, scaling (min/max/logBase), gridlines, tickLblPos
- Define `ChartLegendData`: position (top/bottom/left/right/overlay), entries
- Define `ChartTitleData`: text, formatting, overlay
- Serde attributes for all types

### 0.8.1 — Chart Axis & Legend Types

| Scope | Details |
|---|---|
| **Rust** | Detailed axis/legend/formatting types |
| **TypeScript** | Type exports |
| **Tests** | +5 Rust unit tests |

- Axis types: CategoryAxis, ValueAxis, DateAxis, SeriesAxis
- Axis scaling: min, max, major/minor unit, logBase, orientation (minMax/maxMin)
- Axis formatting: number format, font, line style, tick marks (cross/in/out/none)
- Grid lines: major/minor, line style, color
- Legend entry: idx, delete (hide specific series from legend)
- Chart area formatting: fill, border, shadow, 3D rotation
- Plot area formatting: fill, border, layout (manual positioning x/y/w/h)

### 0.8.2 — DrawingML Chart Writer (Bar, Column, Line)

| Scope | Details |
|---|---|
| **Rust** | Chart XML generation for 3 common types |
| **TypeScript** | None |
| **Tests** | +10 Rust tests |

- Generate `xl/charts/chart{n}.xml` with proper namespace declarations (c:, a:, r:)
- Write `<c:chartSpace>` → `<c:chart>` → `<c:plotArea>` structure
- Bar chart: `<c:barChart>` with barDir (bar/col), grouping (clustered/stacked/percentStacked)
- Line chart: `<c:lineChart>` with grouping, marker styles, smooth lines
- Series: `<c:ser>` with `<c:cat>` (categories) and `<c:val>` (values) referencing cell ranges
- Axes: category axis + value axis with crossing, position, formatting
- Test: bar chart with 3 series → valid XML that Excel accepts
- Test: line chart with markers → valid XML

### 0.8.3 — DrawingML Chart Writer (Pie, Scatter, Area)

| Scope | Details |
|---|---|
| **Rust** | Chart XML generation for 3 more types |
| **TypeScript** | None |
| **Tests** | +10 Rust tests |

- Pie chart: `<c:pieChart>` with `<c:ser>` and explosion settings — no axes
- Doughnut chart: `<c:doughnutChart>` with holeSize
- Scatter chart: `<c:scatterChart>` with scatterStyle, xVal/yVal references
- Area chart: `<c:areaChart>` with grouping (standard/stacked/percentStacked)
- Radar chart: `<c:radarChart>` with radarStyle (radar/filled)
- Test: pie chart with exploded slice → valid XML
- Test: scatter chart with XY data → valid XML
- Test: area chart stacked → valid XML

### 0.8.4 — Chart Drawing Anchors

| Scope | Details |
|---|---|
| **Rust** | Drawing + anchor XML generation |
| **TypeScript** | None |
| **Tests** | +5 Rust tests |

- Generate `xl/drawings/drawing{n}.xml` with `<xdr:twoCellAnchor>`
- Anchor positioning: from cell + offset (EMUs), to cell + offset
- Generate `<xdr:graphicFrame>` referencing chart part
- Generate `xl/drawings/_rels/drawing{n}.xml.rels` for chart relationship
- Generate `xl/worksheets/_rels/sheet{n}.xml.rels` for drawing relationship
- Update `[Content_Types].xml` with drawing and chart overrides
- Test: chart anchored from A1 to F15 → verify EMU calculations

### 0.8.5 — Chart Titles, Labels & Legend

| Scope | Details |
|---|---|
| **Rust** | Title and label XML generation |
| **TypeScript** | None |
| **Tests** | +5 Rust tests |

- Chart title: `<c:title>` with rich text (formatted runs) or plain string
- Axis titles: same structure on category and value axes
- Data labels: `<c:dLbls>` with showVal, showCatName, showPercent, showSerName, numFmt
- Individual point data labels (override series-level setting)
- Legend: `<c:legend>` with position and overlay settings
- Test: chart with title + axis labels + data labels → valid XML
- Test: legend positioning (all 4 positions) → roundtrip

### 0.8.6 — TypeScript Chart Creation API

| Scope | Details |
|---|---|
| **Rust** | None |
| **TypeScript** | `Worksheet.addChart()` method + ChartBuilder |
| **Tests** | +10 TypeScript tests |

- `Worksheet.addChart(type, options)` → add chart to sheet
- `ChartBuilder` fluent API: `.title()`, `.addSeries()`, `.setXAxis()`, `.setYAxis()`, `.legend()`, `.anchor()`
- Type-safe options per chart type (pie has no axes, scatter has xVal/yVal)
- Test: create bar chart with ChartBuilder → write → open in Excel
- Test: create pie chart → write → verify
- Test: create scatter chart with XY data → write → verify
- Test: create chart with full formatting (title, labels, legend, colors) → write

### 0.8.7 — DrawingML Chart Reader

| Scope | Details |
|---|---|
| **Rust** | Chart XML parser — replace opaque blob passthrough |
| **TypeScript** | `Worksheet.charts` getter |
| **Tests** | +10 Rust + TypeScript tests |

- Parse `xl/charts/chart{n}.xml` into ChartData structs
- Detect chart type from XML element names
- Parse series data references (cell ranges)
- Parse axis configuration, title, legend
- Replace opaque passthrough with structured data (for charts — keep passthrough for images)
- TypeScript: `Worksheet.charts` → array of typed chart objects
- Test: read Excel-generated chart → verify all fields parsed correctly
- Test: chart with custom formatting → verify styles parsed

### 0.8.8 — Chart Roundtrip & Styles

| Scope | Details |
|---|---|
| **Rust** | Chart style system |
| **TypeScript** | Style presets |
| **Tests** | +10 roundtrip tests |

- Chart style presets: numbered styles matching Excel's built-in chart styles (1–48)
- Color scheme: use theme colors or explicit RGB
- Series formatting: fill color, border, marker shape/size
- Apply chart style → write → read → verify style attributes preserved
- Read chart from Excel → write back → compare (fidelity test)
- Test: each chart type (bar, line, pie, scatter, area) full roundtrip
- Test: chart with custom colors → roundtrip

### 0.8.9 — Advanced Chart Features

| Scope | Details |
|---|---|
| **Rust** | Advanced DrawingML features |
| **TypeScript** | API extensions |
| **Tests** | +10 tests |

- Combined/combo charts: bar + line on same plot area with secondary axis
- Trendlines: linear, exponential, logarithmic, polynomial, moving average
- Error bars: fixed, percentage, standard deviation, custom
- Data table below chart (display values in tabular form)
- Secondary axis: add second value axis for different scale
- 3D settings: rotX, rotY, perspective, rAngAx (for 3D chart variants)
- Test: combo chart (bar + line) → write → open in Excel
- Test: chart with trendline → roundtrip
- Test: chart with error bars → roundtrip

---

## 0.9.x — Advanced Features & Production Polish

> Final features, performance tuning, API stabilization, and ecosystem readiness.

### 0.9.0 — Pivot Table Reader (Definitions)

| Scope | Details |
|---|---|
| **Rust** | Parse `xl/pivotTables/pivotTable{n}.xml` |
| **TypeScript** | Read-only type exports |
| **Tests** | +5 Rust tests |

- Parse pivot table definition XML: name, dataCaption, location (ref range)
- Parse pivot fields: axis (axisRow/axisCol/axisPage/axisValues), items, subtotals
- Parse data fields: name, fld (source field index), subtotal function (sum/count/average/etc.)
- Parse pivot table relationships and content types
- Define `PivotTableData`, `PivotFieldData`, `PivotDataFieldData` Rust types
- Test: read Excel pivot table → verify fields, data fields, location

### 0.9.1 — Pivot Table Reader (Cache)

| Scope | Details |
|---|---|
| **Rust** | Parse `xl/pivotCache/` parts |
| **TypeScript** | `Worksheet.pivotTables` read-only getter |
| **Tests** | +5–10 roundtrip tests |

- Parse `pivotCacheDefinition{n}.xml`: source range, fields with shared items
- Parse `pivotCacheRecords{n}.xml`: cached data values
- Link cache to pivot table via relationship
- Preserve cache through roundtrip (write back during save)
- TypeScript: `Worksheet.pivotTables` → readonly array of pivot table metadata
- Test: pivot table with cache → roundtrip → verify cache intact
- Test: multiple pivot tables sharing same cache → roundtrip

### 0.9.2 — Threaded Comments

| Scope | Details |
|---|---|
| **Rust** | Parse/write `xl/threadedComments/` |
| **TypeScript** | Threaded comment API |
| **Tests** | +10 roundtrip tests |

- Parse `xl/threadedComments/threadedComment{n}.xml`
- ThreadedComment: id, ref (cell), personId, text, timestamp, parentId (for replies)
- Parse `xl/persons/person.xml`: Person entries (id, displayName, providerId)
- Write threaded comment XML parts during save
- Maintain relationship between legacy comments and threaded comments
- TypeScript: `Worksheet.addThreadedComment(cell, text, author)`, `.replyToComment(commentId, text, author)`
- Test: create threaded comment → write → read → verify
- Test: reply chain (3 levels) → roundtrip

### 0.9.3 — Slicers

| Scope | Details |
|---|---|
| **Rust** | Parse `xl/slicers/slicer{n}.xml` |
| **TypeScript** | Read-only getter |
| **Tests** | +5 roundtrip tests |

- Parse slicer XML: name, caption, source (table/pivot), column name
- Parse slicer cache definition: source data, items, selected items
- Parse slicer drawing anchor (position/size on sheet)
- Preserve slicer state through roundtrip
- TypeScript: `Worksheet.slicers` → readonly array of slicer metadata
- Test: table slicer → roundtrip preservation
- Test: pivot slicer → roundtrip preservation

### 0.9.4 — Timelines

| Scope | Details |
|---|---|
| **Rust** | Parse `xl/timelines/timeline{n}.xml` |
| **TypeScript** | Read-only getter |
| **Tests** | +5 roundtrip tests |

- Parse timeline XML: name, caption, source pivot table, source column
- Parse timeline cache definition: selected date range, level (years/quarters/months/days)
- Parse timeline drawing anchor
- Preserve timeline state through roundtrip
- TypeScript: `Worksheet.timelines` → readonly array of timeline metadata
- Test: timeline with date range selection → roundtrip
- Test: timeline at different zoom levels → roundtrip

### 0.9.5 — Performance Optimization Pass

| Scope | Details |
|---|---|
| **Rust** | Hot path profiling and optimization |
| **TypeScript** | None |
| **Tests** | +5 benchmark tests |

- Profile read/write hot paths with large files (1M rows)
- Optimize XML writer: pre-allocate buffer sizes, reduce string allocations
- Optimize shared string table: hash-based dedup during write
- Optimize ZIP compression: tune deflate level for size vs. speed
- Memory usage audit: peak memory during 1M-row write/read
- Target benchmarks: 2x faster than ExcelJS, within 80% of SheetJS on read
- Publish benchmark comparison table

### 0.9.6 — WASM Binary Size Optimization

| Scope | Details |
|---|---|
| **Rust** | Binary size reduction |
| **TypeScript** | Lazy loading patterns |
| **Tests** | +3 size regression tests |

- Audit WASM binary for unused code paths (wasm-snip)
- Feature-gate optional modules: encryption, formula eval, chart creation
- Tree-shake unused Rust code (ensure LTO + codegen-units=1)
- wasm-opt flags: `-Oz --enable-bulk-memory --enable-nontrapping-float-to-int`
- TypeScript: document lazy-loading pattern for optional features
- Target: core WASM < 300KB gzipped, full WASM < 500KB gzipped
- Add CI check: fail if WASM size exceeds threshold

### 0.9.7 — Public API Audit & Deprecations

| Scope | Details |
|---|---|
| **Rust** | None |
| **TypeScript** | API consistency review |
| **Tests** | None |

- Audit every public export for naming consistency (get/set vs. property, add vs. create)
- Audit parameter types: prefer options objects over positional args for 3+ params
- Audit return types: ensure consistency (null vs. undefined, array vs. readonly array)
- Audit error types: standardized error classes with codes
- Deprecate any APIs that will change in 1.0.0 (add `@deprecated` JSDoc + console.warn)
- Document all breaking changes planned for 1.0.0
- Publish migration guide for deprecated APIs

### 0.9.8 — Ecosystem Adapters

| Scope | Details |
|---|---|
| **Rust** | None |
| **TypeScript** | Framework adapters, CLI tool |
| **Tests** | +5–10 tests |

- CLI tool: `modern-xlsx info file.xlsx` (sheet names, row counts, dimensions)
- CLI tool: `modern-xlsx convert file.xlsx output.json` (extract as JSON)
- CLI tool: `modern-xlsx convert file.xlsx output.csv --sheet 0` (export sheet as CSV)
- React hook: `useXlsx(url)` → `{ workbook, loading, error }`
- Worker/ServiceWorker example: process XLSX in background thread
- CDN bundle: IIFE build for `<script>` tag usage
- Test: CLI commands produce expected output

### 0.9.9 — Release Candidate Preparation

| Scope | Details |
|---|---|
| **Rust** | Final polish |
| **TypeScript** | Final polish |
| **Tests** | Full regression suite |

- Run complete test suite: all Rust + TypeScript tests pass
- Run full benchmark suite: publish final performance numbers
- Verify compatibility: Excel 2016+, Google Sheets, LibreOffice 7+
- Verify browser compatibility: Chrome 120+, Firefox 120+, Safari 17+, Edge 120+
- Verify runtime compatibility: Node.js 20+, Bun 1.0+, Deno 1.40+
- CHANGELOG: complete entries for all 0.x releases
- README: badges, quick-start, feature matrix, benchmark table, migration links
- Publish 1.0.0-rc.1 to npm with `--tag next`
- Collect feedback on RC, fix critical issues → 1.0.0-rc.2 if needed

---

## 1.0.0 — Stable Release

> API freeze. Production commitment. Migration-ready.

### Prerequisites

All must be true before publishing 1.0.0:

| Requirement | Status |
|---|---|
| All 0.x features implemented and tested | |
| Zero known correctness bugs | |
| 100% JSDoc coverage on public API | |
| API reference documentation published | |
| Migration guide: SheetJS → modern-xlsx | |
| Migration guide: ExcelJS → modern-xlsx | |
| Performance benchmarks published | |
| Browser compatibility matrix verified | |
| Runtime compatibility verified (Node/Bun/Deno) | |
| Security audit passed | |
| CHANGELOG complete for all 0.x releases | |
| README with badges, quick start, feature matrix | |

### What 1.0.0 Means

| Commitment | Description |
|---|---|
| **API stability** | No breaking changes until 2.0.0 |
| **SemVer strict** | Patch = bugfix only, minor = additive features only |
| **Production-ready** | Validated in real-world applications |
| **Feature-complete** | Covers 95%+ of common XLSX use cases |
| **Documented** | Every public API has docs, examples, and types |

### Release Process

1. `1.0.0-rc.1` — Feature freeze, bug fixes only
2. `1.0.0-rc.2` — Fixes from RC feedback (if needed)
3. `1.0.0` — Stable release with announcement

---

## Summary

| Minor | Theme | Patches | Focus |
|---|---|---|---|
| **0.1.x** | Hardening & Test Coverage | 9 | Tests, edge cases, type safety, docs, migration guides |
| **0.2.x** | Tables & Structured References | 10 | Table read/write, styles, calculated columns, utilities |
| **0.3.x** | Print & Page Layout | 10 | Margins, headers/footers, print titles/areas, grouping, tab colors |
| **0.4.x** | Advanced Sheet Features | 10 | Split panes, sheet views, workbook protection, sheet management |
| **0.5.x** | Data Features | 10 | Sparklines, data tables, external links, custom XML |
| **0.6.x** | Security & Encryption | 10 | OLE2 detection, AES-256 decrypt/encrypt, security hardening |
| **0.7.x** | Formulas & Calculation | 10 | Tokenizer, parser, serializer, ref rewriter, evaluation engine |
| **0.8.x** | Charts & Visualizations | 10 | Chart model, DrawingML writer/reader, styles, advanced features |
| **0.9.x** | Advanced & Polish | 10 | Pivots, threaded comments, slicers, timelines, perf, API audit, ecosystem |
| **1.0.0** | Stable Release | 1 | API freeze, production commitment |
| | | **90** | |

### Estimated Test Counts at 1.0.0

| Layer | Current | Projected |
|---|---|---|
| **Rust** | ~157 | ~550–600 |
| **TypeScript** | ~123 | ~400–450 |
| **Total** | ~280 | ~950–1050 |

### Risk Matrix

| Risk | Impact | Likelihood | Mitigation |
|---|---|---|---|
| Encryption complexity (0.6.x) | High | Medium | Research early, leverage established crypto crates |
| DrawingML chart scope (0.8.x) | High | Medium | Cap at 10 chart types, skip 3D initially |
| Formula eval scope creep (0.7.x) | Medium | High | Cap at ~50 functions, no circular ref resolution |
| Late-breaking API changes | Medium | Low | 0.9.7 API audit catches issues pre-1.0 |
| WASM binary size growth | Low | Medium | Track per-release, feature-gate optional modules |
| Competitor catches up on styling | Low | Low | Speed + open source + modern API = differentiated |
