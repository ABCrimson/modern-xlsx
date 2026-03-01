# Ironsheet

XLSX read/write library: Rust WASM core + TypeScript API.

## Project Structure

```
crates/ironsheet-core/     # Rust core library (all OOXML parsing/writing)
crates/ironsheet-wasm/     # WASM bridge (wasm-bindgen exports)
packages/ironsheet/        # TypeScript package (public API)
  src/                     # Source: index.ts, workbook.ts, types.ts, wasm-loader.ts
  wasm/                    # Built WASM artifacts (gitignored, built by wasm-pack)
  __tests__/               # Vitest integration tests
  dist/                    # Built JS output (gitignored)
docs/plans/                # Design and implementation plans
```

## Build Commands

```bash
# Rust tests (90 tests)
cargo test -p ironsheet-core

# WASM build (from repo root)
cd crates/ironsheet-wasm && wasm-pack build --target web --release --out-dir ../../packages/ironsheet/wasm --no-opt

# TypeScript build
pnpm -C packages/ironsheet build

# TypeScript tests (26 tests)
pnpm -C packages/ironsheet test

# Lint
pnpm -C packages/ironsheet lint

# Type check
pnpm -C packages/ironsheet typecheck
```

## Architecture

- **Reader path:** `Uint8Array` -> WASM `read()` -> ZIP decompress -> parse XML parts -> `WorkbookData` (serde) -> JS
- **Writer path:** JS `WorkbookData` -> WASM `write()` -> build SST + XML parts -> ZIP compress -> `Uint8Array`
- **SST behavior:** Writer expects `cellType: "sharedString"` with actual text in `value`. Reader returns SST index in `value` and includes `sharedStrings.strings[]` array for lookup.
- **Serde:** All Rust types use `#[serde(rename_all = "camelCase")]` for JS interop.

## Toolchain Versions

- Rust 1.94.0 Edition 2024, wasm-bindgen 0.2.114, quick-xml 0.39.2, zip 8.1 (feature: `deflate`)
- TypeScript 6.0.0-dev, Vitest 4.1.0-beta.5, tsdown 0.21.0-beta.2, Biome 2.4.4, pnpm 11

## Code Conventions

- Rust: `thiserror` for errors, `quick-xml` SAX-style parsing, co-located `#[cfg(test)]` modules
- TypeScript: ESM-only, single quotes, trailing commas, semicolons, 2-space indent (Biome)
- WASM bridge: `serde-wasm-bindgen` for structured data, `JsError` for error conversion
