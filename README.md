<p align="center">
  <h1 align="center">modern-xlsx</h1>
  <p align="center">
    High-performance XLSX read/write for JavaScript &amp; TypeScript, powered by Rust + WASM.
  </p>
</p>

<p align="center">
  <a href="https://www.npmjs.com/package/modern-xlsx"><img alt="npm" src="https://img.shields.io/npm/v/modern-xlsx?color=cb0000&label=npm&logo=npm"></a>
  <a href="https://github.com/ABCrimson/modern-xlsx/blob/main/packages/modern-xlsx/LICENSE"><img alt="License" src="https://img.shields.io/badge/license-MIT-blue"></a>
  <img alt="Types" src="https://img.shields.io/badge/types-included-blue?logo=typescript&logoColor=white">
  <img alt="Zero deps" src="https://img.shields.io/badge/dependencies-0-brightgreen">
</p>

---

Full cell styling, data validation, conditional formatting, frozen panes, hyperlinks, comments, sheet protection, and more — features that SheetJS locks behind a **paid Pro license** — all **free and open source**.

```typescript
import { initWasm, Workbook } from 'modern-xlsx';

await initWasm();

const wb = new Workbook();
const ws = wb.addSheet('Sheet1');
ws.cell('A1').value = 'Hello';
ws.cell('B1').value = 42;

const bold = wb.createStyle().font({ bold: true }).build(wb.styles);
ws.cell('A1').styleIndex = bold;

await wb.toFile('output.xlsx');
```

## Performance

100,000-row workbook, Node.js, single thread:

| Operation | modern-xlsx | SheetJS CE | |
|-----------|------------:|-----------:|---:|
| **Read** | 1,155 ms | 4,927 ms | **4.3x faster** |
| **Write** | 5,048 ms | 5,048 ms | 1.0x |
| sheetToJson (10K) | 54 ms | 103 ms | **1.9x faster** |

> ~29 KB JS + ~1.6 MB WASM. Zero runtime dependencies.

## Install

```bash
npm install modern-xlsx
```

> Full API documentation: **[packages/modern-xlsx/README.md](./packages/modern-xlsx/README.md)**

## Repository Structure

```
crates/
  modern-xlsx-core/       Rust core — OOXML parsing, XML generation, ZIP I/O
  modern-xlsx-wasm/       WASM bridge — wasm-bindgen exports

packages/
  modern-xlsx/            npm package — TypeScript API, tests, benchmarks
```

### Architecture

```
  TypeScript API    Workbook / Worksheet / Cell
                              │ JSON
  WASM boundary     wasm-bindgen bridge
                              │
  Rust core         OOXML parser & writer (quick-xml + zip)
```

Data crosses the WASM boundary as JSON strings for maximum throughput. The Rust core handles ZIP compression, SAX-style XML parsing, shared string table construction, and style resolution.

## Development

```bash
# Rust tests (424 tests)
cargo test -p modern-xlsx-core

# WASM build
cd crates/modern-xlsx-wasm && wasm-pack build --target web --release \
  --out-dir ../../packages/modern-xlsx/wasm --no-opt

# TypeScript build + tests (1287 tests)
pnpm -C packages/modern-xlsx build
pnpm -C packages/modern-xlsx test

# Lint
cargo clippy -p modern-xlsx-core -- -D warnings
pnpm -C packages/modern-xlsx lint
```

**Toolchain:** Rust 1.95.0 (Edition 2024) / TypeScript 6.0 / pnpm 11 / Biome 2.4

## License

[MIT](./packages/modern-xlsx/LICENSE)
