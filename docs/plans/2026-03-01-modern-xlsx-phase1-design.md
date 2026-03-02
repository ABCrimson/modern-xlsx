# modern-xlsx Phase 1 Design

<div align="center">

**Core Read/Write MVP**

Rust WASM core + TypeScript API for XLSX read/write

</div>

---

|  |  |
|---|---|
| **Date** | 2026-03-01 |
| **Status** | Completed |
| **Approach** | Bottom-up layered — Rust core first, then WASM bridge, then TypeScript glue |

---

## Architecture

Hybrid Rust WASM + TypeScript library. Rust handles CPU-intensive operations (ZIP, XML parsing, SST, compression). TypeScript provides the developer API, type safety, and orchestration.

```
modern-xlsx/
├── Cargo.toml                    # Rust workspace root
├── package.json                  # pnpm workspace root
├── pnpm-workspace.yaml           # pnpm 11 catalogs
├── biome.json                    # Biome 2.x
├── tsconfig.base.json
├── .cargo/config.toml            # WASM rustflags
├── crates/
│   ├── modern-xlsx-core/         # Pure Rust: data models, XML, ZIP
│   └── modern-xlsx-wasm/         # WASM bindings via wasm-bindgen
├── packages/
│   └── modern-xlsx/              # Main npm package (TS + WASM)
├── benchmarks/
└── docs/
```

---

## Toolchain

| Tool | Version |
|---|---|
| Rust | 1.94.0 (Edition 2024) |
| wasm-pack | 0.14.0 |
| wasm-bindgen | 0.2.114 |
| Node.js | 25.x |
| TypeScript | 6.0.0-dev |
| pnpm | 11.0.0-alpha.11 |
| tsdown | 0.21.0-beta.2 |
| Vitest | 4.1.0-beta.5 |
| Biome | 2.4.4 |

---

## Implementation Layers

### Layer 1 — Project Scaffolding

- Rust workspace `Cargo.toml` with both crates
- pnpm workspace with `packages/modern-xlsx`
- Config files: `biome.json`, `tsconfig.base.json`, `.cargo/config.toml`
- Git init with `.gitignore`

### Layer 2 — Error Types

- `ModernXlsxError` enum with `thiserror` 2.0
- Variants: ZipRead, ZipWrite, ZipEntry, ZipFinalize, XmlParse, InvalidCell, InvalidStyle, InvalidDate, InvalidFormat, Wasm

### Layer 3 — ZIP Layer

- Read/write ZIP archives using `zip` 8.1.0
- Security: 100:1 decompression ratio limit, 2GB max, path traversal rejection

### Layer 4 — OPC Layer

- `[Content_Types].xml` parser/writer (Default + Override entries)
- `.rels` relationship file parser/writer
- Package traversal: root rels → workbook → per-part rels

### Layer 5 — SpreadsheetML Types + Parsing

- Workbook: sheet list, defined names, date system, calc properties
- Shared String Table: SST with interning, rich text runs
- Styles: numFmts, fonts, fills, borders, cellXfs, cellStyleXfs, cellStyles
- Worksheet: sheetData (rows/cells), merged cells, frozen panes, auto-filter, column widths
- Cell types: string, number, boolean, error, formula, inline string

### Layer 6 — Date Handling

- Serial number ↔ date conversion (chrono 0.4 no-std)
- 1900 system with Lotus 1-2-3 leap year bug
- 1904 system (Mac legacy)
- Number format classifier

### Layer 7 — WASM Bridge

- `wasm-bindgen` 0.2.114 exports
- Bulk data transfer via `Uint8Array` (no per-cell calls)
- Read: JS → WASM decompresses/parses → structured data
- Write: JS → WASM generates XML/ZIP → `Uint8Array`

### Layer 8 — TypeScript API

- WASM loader (auto-detect runtime: Node/Bun/Deno/browser/Worker)
- `Workbook`, `Worksheet`, `Cell` classes
- `StyleBuilder` (fluent API)
- `readFile()`, `readBuffer()`, `Workbook.toBuffer()`, `Workbook.toFile()`

---

## Design Decisions

| Decision | Rationale |
|---|---|
| ESM-only | No CJS. `"type": "module"` |
| Zero runtime deps | WASM binary bundled, nothing else |
| `Uint8Array` everywhere | No Node Buffer |
| Temporal for dates | Not Date objects (available in compat layer) |
| Transitional conformance first | What Excel actually writes |
| Both namespaces | Handle Transitional and Strict XML namespaces in reader |

---

## Validation Criteria

| Metric | Target |
|---|---|
| Excel compatibility | Opens without repair prompts in Excel 365, Google Sheets, LibreOffice 25.x |
| Roundtrip fidelity | All Phase 1 features preserved through read → write |
| WASM binary size | < 600KB uncompressed |
| Read performance | 100K rows × 10 cols < 200ms |
