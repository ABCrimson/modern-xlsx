# modern-xlsx agent guide

modern-xlsx is an XLSX read/write library with a Rust core compiled to WebAssembly and a TypeScript API, published to npm as `modern-xlsx`.

## Layout

Repository root:

- `Cargo.toml`: the Cargo workspace (core and wasm crates) with shared package metadata, exact dependency pins, `[workspace.lints]` and the release profile. `Cargo.lock` is committed.
- `rust-toolchain.toml`: the dated nightly channel with clippy, rustfmt and the `wasm32-unknown-unknown` target.
- `clippy.toml`: clippy settings.
- `deny.toml`: the cargo-deny policy (advisories, licenses, bans, sources).
- `package.json`: the private root manifest with `packageManager`, the WASM build scripts and Biome.
- `pnpm-workspace.yaml`: the workspace packages and the `minimumReleaseAge` gate with its exclusions. `pnpm-lock.yaml` is committed.
- `biome.json`: formatter and linter settings for the whole repository.

Rust:

- `crates/modern-xlsx-core/`: the Rust core.
  - `crates/modern-xlsx-core/src/reader.rs` and `crates/modern-xlsx-core/src/writer.rs`: the workbook read and write entry points.
  - `crates/modern-xlsx-core/src/ooxml/`: one module per OOXML part; `crates/modern-xlsx-core/src/ooxml/worksheet/` holds the worksheet SAX driver (`driver.rs`), the shared SAX helpers and read entry points (`parser.rs`), the struct and JSON row sinks (`sink_struct.rs`, `sink_json.rs`), the JSON escaping and number helpers (`json.rs`) and the XML writer (`writer.rs`), and `crates/modern-xlsx-core/src/ooxml/charts/` the chart parser, writer and types.
  - `crates/modern-xlsx-core/src/zip/`: the ZIP reader (with `ZipSecurityLimits`) and writer.
  - `crates/modern-xlsx-core/src/ole2/`: format detection, `EncryptionInfo` parsing and ECMA-376 encryption; `crates/modern-xlsx-core/src/ole2/cfb/` holds the MS-CFB compound-file reader and writer, and `crates/modern-xlsx-core/src/ole2/crypto/` the Agile (`agile.rs`) and Standard (`standard.rs`) cipher paths.
  - `crates/modern-xlsx-core/src/streaming_writer.rs`: the streaming writer that writes rows straight into ZIP entries.
  - `crates/modern-xlsx-core/src/validate.rs`: validation and repair.
  - `crates/modern-xlsx-core/tests/`: integration tests, with goldens in `crates/modern-xlsx-core/tests/golden/`, fixtures in `crates/modern-xlsx-core/tests/fixtures/`, shared helpers in `crates/modern-xlsx-core/tests/common/mod.rs` and the msoffcrypto-tool oracle in `crates/modern-xlsx-core/tests/interop/`.
  - `crates/modern-xlsx-core/benches/read_write.rs`: the criterion benchmarks.
- `crates/modern-xlsx-wasm/`: the wasm-bindgen bridge (`crates/modern-xlsx-wasm/src/lib.rs`) and its tests (`crates/modern-xlsx-wasm/tests/bridge.rs`).
- `fuzz/`: a separate cargo workspace for cargo-fuzz with its own committed `fuzz/Cargo.lock`. `fuzz/fuzz_targets/` holds six targets, and `fuzz/seeds/` holds one seed directory per target.

TypeScript:

- `packages/modern-xlsx/`: the npm package.
  - `packages/modern-xlsx/src/`: the public API; `packages/modern-xlsx/src/index.ts` and `packages/modern-xlsx/src/index-lite.ts` are the full and lite entries.
  - `packages/modern-xlsx/src/model/`: the workbook model: `Workbook`, `Worksheet` and `Cell`, the public types (the `types.ts` barrel over `model/types/`), `ModernXlsxError` with its codes (`errors.ts`), and the Excel-form GUID helper for threaded comment and person ids (`guid.ts`).
  - `packages/modern-xlsx/src/format/`: number-format rendering (`format-cell.ts`), date serials (`dates.ts`), A1 cell references (`cell-ref.ts`) and Excel number text (`number-text.ts`).
  - `packages/modern-xlsx/src/authoring/`: the builders: charts (`chart-builder.ts`, `validate-chart.ts`, `chart-styles.ts`), pivot tables (`pivot-builder.ts`), styles (`style-builder.ts`), rich text (`rich-text.ts`), headers and footers (`header-footer.ts`) and drawn tables (`table.ts`, `table-styles.ts`).
  - `packages/modern-xlsx/src/io/`: the sheet conversion helpers such as `aoaToSheet`, `sheetToJson` and `sheetToCsv` (`sheet-utils.ts`) and `StreamingXlsxWriter` (`streaming-writer.ts`).
  - `packages/modern-xlsx/src/formula/`: the formula tokenizer, parser, serializer, rewriter and evaluator.
  - `packages/modern-xlsx/src/barcode/`: the barcode encoders and renderer.
  - `packages/modern-xlsx/__tests__/`: the Vitest suites (Node and browser); `packages/modern-xlsx/__tests__/bench/` holds the Vitest benchmarks.
  - `packages/modern-xlsx/scripts/`: `finalize-dist.mjs` (post-build step), `smoke-tarball.mjs` (CI tarball smoke test) and `sw-smoke.mjs` (service-worker example test).
  - `packages/modern-xlsx/CHANGELOG.md`: the release history.
  - Generated and untracked (`.gitignore`): `packages/modern-xlsx/wasm/` and `packages/modern-xlsx/wasm-lite/` (wasm-pack output, full and lite) and `packages/modern-xlsx/dist/` (tsdown output).
- `examples/`: grouped by target: Node scripts (`examples/node/chart-dashboard/`, `examples/node/csv-to-styled-xlsx/`, `examples/node/encrypted-payroll/`, `examples/node/export-sales-report/`), edge examples (`examples/edge/cloudflare-worker/`, `examples/edge/deno-deploy/`, `examples/edge/service-worker/`), framework snippets (`examples/frameworks/react/`, `examples/frameworks/angular/`, `examples/frameworks/vue/`, `examples/frameworks/svelte/`), the browser demo (`examples/browser/demo-site/`) and runtime tests (`examples/runtimes/`).
- `scripts/`: `scripts/examples-smoke.mjs` runs the examples against a packed tarball, `scripts/examples-typecheck.mjs` type-checks the framework snippets against it, `scripts/sync-example-pins.mjs` pins the examples, `docs/site/*.html` and both READMEs to the package version, `scripts/check-readme-sync.mjs` checks the Quick Start block the two READMEs share, and `scripts/prepare-release.mjs` makes the release version bump.

Docs and CI:

- `docs/guide/`: feature guides. `docs/security/threat-model.md`: the threat model. `docs/ROADMAP.md`: the roadmap.
- `docs/site/`: the GitHub Pages site. `docs/site/api/` is the TypeDoc output, generated at deploy time and untracked (`.gitignore`).
- `.github/workflows/ci.yml`: every gate and the npm publish job. `.github/workflows/pages.yml`: the Pages deploy.

## Commands

Every command in the code blocks below is copied from `.github/workflows/ci.yml` or names a `package.json` script. Run them from the repository root.

Setup (the first line installs the toolchain `rust-toolchain.toml` names):

```bash
rustup toolchain install
pnpm install --frozen-lockfile
```

Rust gates, in CI order:

```bash
cargo fmt --all --check
cargo fmt --manifest-path fuzz/Cargo.toml --all --check
cargo clippy --workspace --all-targets --locked -- -D warnings
cargo clippy --workspace --all-targets --no-default-features --locked -- -D warnings
cargo clippy -p modern-xlsx-core --features parallel --all-targets --locked -- -D warnings
RUSTDOCFLAGS='-D warnings' cargo doc -p modern-xlsx-core -p modern-xlsx-wasm --no-deps
cargo test --workspace --locked
cargo test --workspace --no-default-features --locked
cargo test -p modern-xlsx-core --features parallel --locked
wasm-pack test --node crates/modern-xlsx-wasm
wasm-pack test --node crates/modern-xlsx-wasm --no-default-features
```

The `msrv` job runs `cargo check --workspace --all-targets --locked` on the `rust-version` toolchain. The `encryption-oracle` job runs `cargo test -p modern-xlsx-core --release --locked --test interop -- --ignored` with `MX_MSOFFCRYPTO_PYTHON` naming a Python that has msoffcrypto-tool installed.

WASM build (full into `packages/modern-xlsx/wasm/`, lite into `packages/modern-xlsx/wasm-lite/`; wasm-opt runs with `-Oz` from `[package.metadata.wasm-pack.profile.release]` in `crates/modern-xlsx-wasm/Cargo.toml`; both scripts pass `--config "profile.release.trim-paths='all'"`, as CI's build-wasm job does, so the `.wasm` carries no host paths, and need the pinned nightly because the key is stable only from Cargo 1.101):

```bash
pnpm build:wasm
pnpm build:wasm:lite
```

The `lint` job runs Biome over the whole repository and needs no WASM build:

```bash
pnpm exec biome check .
```

TypeScript gates, in CI order (they need the WASM build):

```bash
pnpm -C packages/modern-xlsx typecheck
pnpm -C packages/modern-xlsx typecheck:tests
pnpm -C packages/modern-xlsx lint
pnpm -C packages/modern-xlsx build
pnpm -C packages/modern-xlsx test
pnpm -C packages/modern-xlsx test:coverage   # writes coverage/ (gitignored); not a CI gate
```

Browser tests, after the build:

```bash
pnpm -C packages/modern-xlsx exec playwright install --with-deps chromium
pnpm -C packages/modern-xlsx test:browser
node packages/modern-xlsx/scripts/sw-smoke.mjs
```

Goldens: run a golden test with `UPDATE_GOLDEN=1` to rewrite its file instead of comparing (`crates/modern-xlsx-core/tests/common/mod.rs`). Review every regenerated diff, and never set `UPDATE_GOLDEN` in CI.

Fuzzing runs on the nightly from `rust-toolchain.toml` and needs cargo-fuzz. List the targets, then run one with `cargo fuzz run <target>`:

```bash
cargo fuzz list
```

Benchmarks: `pnpm -C packages/modern-xlsx bench` builds the package, then runs the Vitest benchmarks; `cargo bench -p modern-xlsx-core --bench read_write` runs the criterion benchmarks.

Release metadata: the version lives only in `packages/modern-xlsx/package.json` (the `VERSION` export reads it at build time). `node scripts/prepare-release.mjs <version> [--date YYYY-MM-DD]` sets it, re-pins the examples, `docs/site/*.html` and both READMEs, and turns the CHANGELOG `## [Unreleased]` heading into `## [<version>] - <date>` below a fresh one; it never runs git. CI checks that the pins, `VERSION` and the CHANGELOG agree and that the READMEs' Quick Start blocks match:

```bash
node scripts/sync-example-pins.mjs --check
node scripts/check-readme-sync.mjs
node scripts/prepare-release.mjs --check
```

API reference (TypeDoc into `docs/site/api/`, after the WASM build):

```bash
pnpm -C packages/modern-xlsx docs:api
```

Advisory CI parts: the `typescript-nightly` job and the dev-toolchain `pnpm audit --audit-level high` step run with `continue-on-error`, and the `fuzz` job is outside the `needs` of the `publish` job.

## Architecture

- Reader path: `readBuffer` or `readFile` (`packages/modern-xlsx/src/index.ts`) calls the wasm `read(data, limits)` (`crates/modern-xlsx-wasm/src/lib.rs`), which calls `reader::read_xlsx_json_with_options`. The reader detects the format (`ole2::detect::detect_format`), unpacks the archive with `zip::read_zip_entries` under `ZipSecurityLimits`, SAX-parses each part with quick-xml and streams a JSON string capped at `max_json_length` bytes. `packages/modern-xlsx/src/wasm-backend.ts` turns that string into objects with `JSON.parse`.
- Writer path: `Workbook.toBuffer` passes `JSON.stringify` output to the wasm `write`, which deserializes it with serde_json and calls `writer::write_xlsx`. The writer expects `cellType: "sharedString"` with the text in `value`: it collects those strings into a `SharedStringTableBuilder`, and `to_xml_with_legacy_drawings` (`crates/modern-xlsx-core/src/ooxml/worksheet/writer.rs`) writes each cell's table index while it writes the sheet XML, with no worksheet copy, and names the sheet's VML drawings (`<legacyDrawing>`, `<legacyDrawingHF>`). The parts are then zipped.
- JSON bridge: workbook data crosses the WASM boundary as JSON strings in both directions. Errors cross as `JsError` messages of the form `[CODE] message`, which `ModernXlsxError.fromWasmError` (`packages/modern-xlsx/src/model/errors.ts`) turns back into a code and a message.
- Encryption and the lite build: the default `encryption` Cargo feature gates only `ole2::crypto` (`crates/modern-xlsx-core/src/ole2.rs`), the crypto dependencies (`[features]` in `crates/modern-xlsx-core/Cargo.toml`) and the wasm exports `readWithPassword` and `writeWithPassword`. The MS-CFB layer `ole2::cfb` (`ole2::cfb::reader` and `ole2::cfb::writer`), `ole2::detect` and `ole2::encryption_info` compile in every build, so the lite build (`--no-default-features`) still reports a legacy .xls as `LEGACY_FORMAT` and an encrypted file as `PASSWORD_PROTECTED` (`crates/modern-xlsx-core/tests/lite_detection.rs`). The lite TypeScript entry rejects a password with `INVALID_ARGUMENT`.
- Entry points (`exports` in `packages/modern-xlsx/package.json`): `.` (full build), `./lite` (lite build, no encryption), `./browser` (the IIFE bundle that exposes `window.ModernXlsx`), `./worker` (the Web Worker script) and `./wasm` (the full WASM binary).
- Streaming writer: the wasm `StreamingWriter` wraps `StreamingWriterCore` (`crates/modern-xlsx-core/src/streaming_writer.rs`), and `StreamingXlsxWriter` (`packages/modern-xlsx/src/io/streaming-writer.ts`) is its TypeScript API.
- Formatting: integers go through `itoa`, floats through `zmij`, and relationship ids through `make_rid` (`crates/modern-xlsx-core/src/ooxml/relationships.rs`), which avoids `format!`. Relationship types and targets are `Cow<'static, str>`.

## Conventions

- Rust: the edition and the MSRV come from `[workspace.package]` in `Cargo.toml`. `[workspace.lints]` forbids `unsafe_code`, warns on clippy pedantic (CI denies warnings) and denies `allow_attributes` and `allow_attributes_without_reason`, so suppressions are `#[expect(lint, reason = "...")]`. `clippy.toml` sets `lint-commented-code`. Errors are a thiserror enum, `ModernXlsxError`, whose `code()` returns a stable code (`crates/modern-xlsx-core/src/errors.rs`). Unit tests sit in co-located `#[cfg(test)]` modules. Every serde type in the core crate uses `#[serde(rename_all = "camelCase")]`.
- TypeScript: ESM only. Biome (`biome.json`) sets single quotes, trailing commas everywhere, semicolons always, 2-space indent, line width 100 and LF line endings.
- WASM bridge: JSON strings in and out; errors go through `JsError`.
- Docs and comments: present tense. History lives only in `packages/modern-xlsx/CHANGELOG.md`. Write no test counts and no version numbers; point at the files below instead.

## Versions

The versions live in these files; read them there:

- `rust-toolchain.toml`: the dated nightly channel, clippy, rustfmt and the `wasm32-unknown-unknown` target.
- `rust-version` in `Cargo.toml`: the MSRV, which the `msrv` CI job checks.
- `[workspace.dependencies]` in `Cargo.toml` (exact `=` pins) and `Cargo.lock`.
- `packageManager` and `devDependencies` in the root `package.json`.
- `packages/modern-xlsx/package.json` and `pnpm-lock.yaml`.
- `pnpm-workspace.yaml`: `minimumReleaseAge` and its exclusions.
- `.github/workflows/ci.yml`: `env.WASM_PACK_VERSION` and the Node matrix of the `typescript` job.
