# modern-xlsx Phase 1 Design: Core Read/Write MVP

**Date:** 2026-03-01
**Status:** Approved
**Approach:** Bottom-Up Layered (Rust core first, then WASM bridge, then TypeScript glue)

## Architecture

Hybrid Rust WASM + TypeScript library. Rust handles CPU-intensive operations (ZIP, XML parsing, SST, compression). TypeScript provides the developer API, type safety, and orchestration.

## Monorepo Structure

```
modern-xlsx/
├── Cargo.toml                    # Rust workspace root
├── package.json                  # pnpm workspace root
├── pnpm-workspace.yaml           # pnpm 11 catalogs
├── biome.json                    # Biome 2.x
├── tsconfig.base.json
├── .cargo/config.toml            # WASM rustflags
├── crates/
│   ├── modern-xlsx-core/           # Pure Rust: data models, XML, ZIP
│   └── modern-xlsx-wasm/           # WASM bindings via wasm-bindgen
├── packages/
│   └── modern-xlsx/              # Main npm package (TS + WASM)
├── benchmarks/
└── docs/
```

## Toolchain (Exact Versions)

| Tool | Version |
|------|---------|
| Rust | 1.94.0 (Edition 2024) |
| wasm-pack | 0.14.0 |
| wasm-bindgen | 0.2.114 |
| Node.js | 25.x |
| TypeScript | 6.0.0-dev.20260301 |
| pnpm | 11.0.0-alpha.11 |
| tsdown | 0.21.0-beta.2 |
| Vitest | 4.1.0-beta.5 |
| Biome | 2.4.4 |

## Implementation Sequence

### Layer 1: Project Scaffolding
- Rust workspace Cargo.toml with both crates
- pnpm workspace with packages/modern-xlsx
- All config files (biome.json, tsconfig.base.json, .cargo/config.toml)
- Git init with .gitignore

### Layer 2: Error Types
- `modern-xlsxError` enum with thiserror 2.0.18
- Error variants: ZipRead, ZipWrite, ZipEntry, ZipFinalize, XmlParse, InvalidCell, InvalidStyle, InvalidDate, InvalidFormat, Wasm

### Layer 3: ZIP Layer
- Read ZIP archives using zip 8.1.0 (lazy metadata parsing)
- Write ZIP archives using zip 8.1.0 SimpleFileOptions API
- libdeflater 1.25.2 freestanding for buffer compression
- flate2 1.1.9 rust_backend for streaming
- Security: 100:1 decompression ratio limit, 2GB max, path traversal rejection

### Layer 4: OPC Layer
- [Content_Types].xml parser/writer (Default + Override entries)
- .rels relationship file parser/writer
- Package traversal: root rels → workbook → per-part rels

### Layer 5: SpreadsheetML Types + Parsing
- Workbook: sheet list, defined names, date system, calc properties
- Shared String Table: SST with string_cache 0.9 interning, rich text runs
- Styles: numFmts, fonts, fills, borders, cellXfs, cellStyleXfs, cellStyles
- Worksheet: sheetData (rows/cells), merged cells, frozen panes, auto-filter, column widths
- Cell types: string (s), number (n), boolean (b), error (e), formula (str), inline string (inlineStr)

### Layer 6: Date Handling
- Serial number ↔ date conversion (chrono 0.4.44 no-std)
- 1900 system with Lotus 1-2-3 leap year bug
- 1904 system (Mac legacy)
- Number format classifier: parse format strings for date tokens

### Layer 7: WASM Bridge
- wasm-bindgen 0.2.114 exports
- Bulk data transfer via Uint8Array (no per-cell calls)
- Read: JS passes Uint8Array → WASM decompresses/parses → returns structured data
- Write: JS passes workbook state → WASM generates XML/ZIP → returns Uint8Array

### Layer 8: TypeScript API
- WASM loader (auto-detect runtime: Node/Bun/Deno/browser/Worker)
- Workbook, Worksheet, Cell classes
- Style builder (fluent API)
- Temporal API for dates (PlainDate, PlainTime, PlainDateTime)
- readFile(), readBuffer(), Workbook.toBuffer(), Workbook.toFile()

## Key Design Decisions

1. **ESM-only** — No CJS. `"type": "module"`.
2. **Zero runtime deps** — WASM binary bundled, nothing else.
3. **Uint8Array everywhere** — No Node Buffer.
4. **Temporal for dates** — Not Date objects (available in compat layer).
5. **Transitional conformance first** — What Excel actually writes. Strict read in Phase 2.
6. **Both namespaces** — Handle both Transitional and Strict XML namespaces in reader.

## Validation Criteria

- Files produced open without repair prompts in Excel 365, Google Sheets, LibreOffice 25.x
- Read/write round-trip preserves all Phase 1 features
- WASM binary < 600KB uncompressed
- Read 100K rows x 10 cols < 200ms
