# Changelog

All notable changes to this project are documented here.

The format is based on [Keep a Changelog](https://keepachangelog.com/), and this project adheres to [Semantic Versioning](https://semver.org/).

## [1.0.0-rc.1] - 2026-03-06

Release candidate 1 — API frozen, production-ready feature set.

### Added
- **PivotTableBuilder:** Fluent builder API for pivot tables (follows ChartBuilder pattern)
- **Auto-Filter Enhancements:** Custom filters (`CustomFilterData`, `CustomFiltersData`) with full roundtrip
- **Page Breaks:** `PageBreaksData` type with row/column break parsing and writing
- **Rich Text Cells:** `Cell.richText` getter/setter, `RichTextBuilder` fluent API, full roundtrip preservation
- **Error System:** `ModernXlsxError` class with `fromWasmError()` parser, 18 machine-readable error codes
- **Image Embedding:** PNG/JPEG image support in worksheets
- New test files: `rich-text.test.ts`, expanded `api-setters.test.ts`, `tables-print-layout.test.ts`

### Changed
- **API Audit:** Naming consistency, parameter types, return types across all public APIs
- **WASM Binary:** Rebuilt with auto-filter, page break, rich text, and error improvements
- **README:** 5-line quick start, improved onboarding, CDN references updated
- **JSDoc Coverage:** Expanded documentation across exported functions and types
- Error messages now context-rich and actionable across Rust core (ZIP, XML, OLE2, cell, style errors)
- ChartBuilder, StyleBuilder expanded with additional JSDoc
- Workbook API expanded with `addPivotTableFromBuilder()` convenience method

### Fixed
- Biome lint: all non-null assertions in test files replaced with optional chaining
- Import ordering: all imports moved to file top per Biome rules
- `exactOptionalPropertyTypes` compliance in PivotTableBuilder

### Stats
- Rust: 424 tests (402 unit + 12 golden + 5 security + 4 bench + 1 doctest), clippy clean
- TypeScript: 1287 tests across 58 files, typecheck clean
- WASM binary: ~2.0 MB

## [0.9.2] - 2026-03-06

Streaming writer, write APIs for pivot tables/slicers/timelines, parser refactoring, security audit, examples, and production polish.

### Added
- **Streaming XLSX Writer:** True streaming writer (`StreamingXlsxWriter`) for 100K+ rows with O(unique strings) memory usage — backed by Rust/WASM `StreamingWriterCore` with incremental ZIP entry writing
- **Pivot Table Write API:** `ws.addPivotTable()`, `ws.removePivotTable()`, `wb.addPivotCache()`, `wb.addPivotCacheRecords()` — full read/write support
- **Slicer Write API:** `ws.addSlicer()`, `ws.removeSlicer()`, `wb.addSlicerCache()` — full read/write support
- **Timeline Write API:** `ws.addTimeline()`, `ws.removeTimeline()`, `wb.addTimelineCache()` — full read/write support
- **Pivot Cache Types:** `PivotCacheDefinitionData`, `PivotCacheRecordsData`, discriminated union `CacheValue` types
- **WASM Lite Build:** Feature-gated encryption for smaller WASM binary via `./lite` export
- **Security Audit:** Comprehensive ECMA-376 crypto audit (`docs/SECURITY-AUDIT.md`)
- **Demo Site:** Interactive dark-themed SPA (`examples/demo-site/`)
- **Example Projects:** Sales report, CSV converter, encrypted payroll, chart dashboard
- **Deno/Bun Compatibility:** Test harnesses for Deno and Bun runtimes
- **CHANGELOG.md:** Full version history from 0.1.0

### Changed
- **CLI Modernization:** Node 24 `parseArgs()` + async `node:fs/promises`
- **Parser Refactoring:** worksheet/parser.rs reduced from 2,891 → 2,202 lines; 20 extracted helper functions for attribute parsing
- README updated: pivot tables, slicers, timelines now "read/write", streaming writer documented

## [0.9.1] - 2026-03-06

35-point performance, quality, and maintainability overhaul.

### Changed
- Comprehensive modernization for latest dependency versions
- Rust modernization audit: `cold_path()` hints on all error branches, `#[inline]` on hot paths, removal of remaining `unwrap()` calls
- TypeScript modernization audit: eliminated all `any` types, added `satisfies` operators, `readonly` array annotations throughout

## [0.9.0] - 2026-03-06

Major feature release: pivot tables, threaded comments, slicers, timelines, CLI tool.

### Added
- **Pivot Tables:** full data model types, SAX parser, XML writer, pivot cache definitions/records, reader/writer pipeline integration, TypeScript read-only API
- **Threaded Comments:** persons and threaded comment types, parser, writer, reader/writer integration, TypeScript API
- **Slicers & Timelines:** slicer and timeline types, parser, writer, reader/writer integration, TypeScript API
- **CLI Tool:** `info` and `convert` commands for command-line XLSX inspection and conversion
- **Cargo Feature Gates:** encryption module behind `encryption` feature flag for smaller builds
- Performance benchmark tests for large workbooks
- Pre-allocated XML writer buffers based on data size estimates
- API consistency audit: readonly arrays, `ModernXlsxError`, null consistency

### Changed
- Comprehensive TypeScript modernization: no `any`, `satisfies`, `readonly`
- Comprehensive Rust modernization: `cold_path`, `#[inline]`, `unwrap` removal
- Rebuilt WASM with all 0.9.x features

## [0.8.6] - 2026-03-05

Rust 1.95 modernization release.

### Changed
- Bumped `rust-version` to 1.95.0
- Modernized WASM bridge helpers (`parse_workbook()`, `to_js_err()`)
- `.find()` iterator pattern for single-attribute XML loops (58 conversions across OOXML parsers)
- `#[inline]` on hot-path helpers (worksheet streaming JSON, chart enum `xml_val` methods)
- `zip()` iteration in reader.rs, `next_r_id()` method on Relationships
- Extracted shared serde helpers (`is_false`, `is_true`, `default_true`) to `ooxml/mod.rs`
- `cold_path()` hints on all error branches (Rust 1.95 stabilization)

## [0.8.5] - 2026-03-05

Comprehensive audit and modernization release.

### Added
- WASM boundary validation for chart data (`validateChartData()`)
- Support for `oneCellAnchor` charts in drawing XML reader/writer
- Chart API and chart type roundtrip tests
- ChartAxis.fontSize with proper `<c:txPr>` writer/parser
- Expanded golden file suite from 2 to 12 scenarios

### Fixed
- Merged image + chart drawing XML to prevent silent data loss (barcode+chart collision)
- Eliminated all `noNonNullAssertion` lint errors with safe array accessors
- Rust modernization: eliminated panics, optimized allocations, Edition 2024 idioms

### Changed
- TypeScript 6.0 modernization: `satisfies`, cached `NumberFormat`, single-pass escaping, worker leak fix
- Deep Rust modernization: iterator combinators, dead code removal, pattern cleanup

## [0.8.1] - 2026-03-04

Audit patch release for v0.8.0.

### Fixed
- ECMA-376 schema ordering compliance in chart XML output
- NaN guard for numeric cell values
- Dead code removal across Rust and TypeScript
- Lint error resolution: non-null assertions, formatting, complexity

## [0.8.0] - 2026-03-04

Charts & Visualizations release.

### Added
- **10 chart types:** bar, column, line, pie, doughnut, scatter, area, radar, bubble, stock
- **ChartBuilder** fluent API in TypeScript for programmatic chart creation
- Chart data model types (`ChartData`, `ChartSeries`, `ChartAxis`)
- Chart XML writer for all chart types with ECMA-376 compliance
- Chart drawing anchors (twoCellAnchor, oneCellAnchor) and ZIP packaging
- Chart XML reader with full roundtrip support
- Chart roundtrip tests and 8 style presets
- Trendlines (6 types), error bars (5 types), 3D rotation (`View3D`)
- Combo charts with secondary chart and secondary value axis
- Data tables, axis font size control

## [0.7.1] - 2026-03-04

Audit patch release for v0.7.0.

### Fixed
- Comprehensive Rust audit: correctness, security, and performance improvements
- Comprehensive TypeScript audit: correctness and precision improvements
- JSON escaping edge cases
- `xml:space` preservation in shared strings
- ROUND function precision
- Formula tokenizer backtrack behavior

## [0.7.0] - 2026-03-04

Formula Engine release.

### Added
- **Formula tokenizer** (lexer) with full Excel formula syntax support
- **Recursive descent parser** producing an AST representation
- **Formula serializer** (AST to string) for roundtrip
- **Cell reference resolver** for evaluating cell references
- **Reference rewriter** for row/column insert/delete operations
- **Shared formula expansion** for array and shared formula support
- **Arithmetic evaluator** with operator precedence
- **54 built-in functions** via `createDefaultFunctions()`:
  - String & logical: IF, AND, OR, NOT, LEN, LEFT, RIGHT, MID, TRIM, UPPER, LOWER, CONCATENATE, etc.
  - Math & statistical: SUM, AVERAGE, COUNT, MIN, MAX, ROUND, ABS, SQRT, POWER, MOD, etc.
  - Lookup: VLOOKUP, HLOOKUP, INDEX, MATCH, OFFSET, INDIRECT, ROW, COLUMN, etc.
- `FormulaFunction` interface with lazy evaluation (IF short-circuit)
- `EvalContext` interface decoupled from Workbook class

## [0.6.1] - 2026-03-04

Post-release fixes and CI improvements.

### Fixed
- Added `digest` dependency to Cargo.toml for CI
- Extracted `compute_hmac` helper to deduplicate HMAC dispatch
- Consolidated `OLE2_MAGIC` constant into `ole2/mod.rs`
- Biome format errors in encryption roundtrip tests

### Changed
- CI triggers on tag pushes for npm publish
- pnpm version resolution from `packageManager` field
- Documented DIFAT sector chain limitation in OLE2 writer

## [0.6.0] - 2026-03-04

Full ECMA-376 Encryption release (read + write).

### Added
- **OLE2 detection:** magic byte detection with descriptive errors for encrypted/legacy files
- **EncryptionInfo parser:** Agile (version 4.4 XML) and Standard (version 2.2/3.2/4.2 binary) detection
- **Key derivation:** SHA-512/SHA-256/SHA-1 per ECMA-376 SS2.3.6.2 with configurable iteration count
- **AES decryption:** AES-256-CBC and AES-128-CBC with PKCS#7 and no-padding modes
- **Password verification:** Agile (hash-based) and Standard (ECB verifier) verification
- **Segment-based decryption:** 4096-byte segment decryption with per-segment IV derivation
- **HMAC integrity verification:** SHA-512/SHA-256 HMAC over encrypted package
- **Standard Encryption decryption:** AES-128-ECB verifier + AES-128-CBC package decryption
- **Decryption pipeline:** `readBuffer(data, { password: '...' })` API with WASM bridge
- **OLE2 compound document writer:** v3 format, 512-byte sectors, FAT/directory/DIFAT
- **File encryption write path:** Agile AES-256-CBC + SHA-512 + 100,000 iterations
- **Security hardening:** `SensitiveKey` RAII zeroization wrapper with `ZeroizeOnDrop`
- **Encryption roundtrip & compatibility test suite**
- Password/encryption support in Worker API

### Security
- `constant_time_eq` for all password and HMAC verification
- `getrandom` (with `wasm_js` feature) for CSPRNG
- `zeroize` with `ZeroizeOnDrop` derive for automatic key material cleanup

## [0.5.1] - 2026-03-03

Audit patch release for v0.5.0.

### Added
- Split pane and per-pane selection support (Rust types, parser, writer, TypeScript API)
- Sheet view attributes per ECMA-376 (SheetViewData, parser, writer, TypeScript API)
- Workbook protection support (struct, parser, writer, TypeScript API)
- Sheet management API: state, move, clone, rename, hide/unhide
- Sparkline groups: data model, parser, writer, TypeScript API
- Data table formula attributes (`r1`/`r2`/`dt2D`/`dtr1`/`dtr2`)
- Content type overrides for external links and custom XML

### Fixed
- Comprehensive audit: correctness, performance, version consistency
- Non-null assertion removal in sheet management methods

## [0.5.0] - 2026-03-03

SheetJS parity sprint release.

### Added
- `usedRange`, `tabColor`, `cell.numberFormat`, `cell.dateValue`, `dynamicArray` support
- `encodeRow`, `decodeRow`, `splitCellRef` utility functions
- `sheetToTxt` and `sheetToFormulae` utility functions
- Conditional sections, color codes, `loadFormat` for number formatting
- Rust struct additions: document properties, tab color, stub types, dynamic array formula support
- Exhaustive feature comparison: modern-xlsx vs SheetJS (117 tests)
- Tables: Rust `TableDefinition` types, SAX parser, reader wiring, relationship constants
- Headers/footers writer, outline level parsing/writing for rows and columns

### Fixed
- Full codebase audit: bugs, performance, docs accuracy

## [0.4.0] - 2026-03-02

Table Layout Engine release.

### Added
- Table layout engine for automatic column width and row height calculation
- Barcode & QR code generation: 9 formats, PNG renderer, XLSX embedding

### Fixed
- Full codebase audit: critical bugs, performance, modernization

## [0.3.0] - 2026-03-02

OOXML Validation & Repair release.

### Added
- OOXML validation engine with automatic repair
- Performance audit and optimization pass
- Documentation accuracy improvements

## [0.2.0] - 2026-03-02

Browser distribution release.

### Added
- Browser distribution bundle (ESM)
- Web Worker support for background XLSX processing
- Interactive playground / demo page
- GitHub Pages deployment

### Fixed
- Rust toolchain CI compatibility (beta toolchain, rust-version pinning)

## [0.1.0] - 2026-03-01

Initial release.

### Added
- **Rust WASM core:** ZIP reader (with path traversal, zip bomb, size limit guards) and ZIP writer (DEFLATE)
- **Cell reference parser:** A1-style with roundtrip support
- **Date serial numbers:** 1900 leap year bug compatibility and 1904 date system
- **Number format classifier:** date/time/number detection
- **Shared String Table:** parser and builder with deduplication
- **OPC Relationships:** parser and writer
- **Styles:** numFmts, fonts, fills, borders, cellXfs parser and writer
- **Workbook.xml:** parser and writer
- **Worksheet XML:** parser (all cell types, merged cells, frozen panes) and writer
- **Read/write orchestrators:** full XLSX read and write pipeline
- **WASM bridge:** read/write via serde-wasm-bindgen (later upgraded to JSON bridge)
- **TypeScript API:** `readBuffer()`, `writeBuffer()`, `Workbook`, `Worksheet` classes
- **Integration tests:** read, write, roundtrip, date handling
- CI pipeline with GitHub Actions
- Hardening, test coverage, and performance audit (v0.1.9)
