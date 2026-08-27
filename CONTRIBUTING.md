# Contributing to modern-xlsx

Thanks for your interest in contributing! This guide covers everything needed to go from a fresh clone to a passing pull request.

## Development Setup

### Prerequisites

| Tool | Version | Notes |
|------|---------|-------|
| **Rust** | beta channel (MSRV 1.96) | `rust-toolchain.toml` pins the beta channel, clippy/rustfmt, and the `wasm32-unknown-unknown` target automatically |
| **Node.js** | 24+ | CI tests Node 24 and 26 |
| **pnpm** | 12+ | The only supported package manager (single `pnpm-lock.yaml`) |
| **wasm-pack** | latest | Builds the WASM bridge crate |

```bash
# Install wasm-pack
cargo install wasm-pack

# Clone and install
git clone https://github.com/ABCrimson/modern-xlsx.git
cd modern-xlsx
pnpm install
```

> [!NOTE]
> `rustup` reads `rust-toolchain.toml` on first `cargo` invocation and installs the pinned beta toolchain plus the `wasm32-unknown-unknown` target for you — no manual `rustup target add` needed.

### Build

> [!IMPORTANT]
> The WASM artifact must exist before the TypeScript build or tests can run. Build it first (or run the all-in-one `pnpm build`).

```bash
# All-in-one from the repo root: WASM + TypeScript
pnpm build
```

Individual steps and variants (from the repo root):

| Command | What it does |
|---------|--------------|
| `pnpm build:wasm` | Release WASM build into `packages/modern-xlsx/wasm/` |
| `pnpm build:wasm:lite` | Encryption-free lite WASM into `packages/modern-xlsx/wasm-lite/` (powers the `modern-xlsx/lite` entry point) |
| `pnpm build:wasm:debug` | Fast unoptimized WASM build for iteration |
| `pnpm build:ts` | TypeScript build via tsdown into `packages/modern-xlsx/dist/` |

### Test

```bash
# Rust tests (441: unit + golden + security + benchmark + doc)
cargo test -p modern-xlsx-core

# TypeScript tests (1,290 across 59 files; needs the WASM build)
pnpm -C packages/modern-xlsx test

# Browser tests (Vitest + Playwright Chromium)
pnpm -C packages/modern-xlsx exec playwright install chromium
pnpm -C packages/modern-xlsx test:browser

# Lint
cargo clippy --workspace --all-targets -- -D warnings
pnpm -C packages/modern-xlsx lint

# Format
cargo fmt
pnpm fmt

# Type check
pnpm -C packages/modern-xlsx typecheck
```

<details>
<summary>Fuzzing the reader (optional, nightly-only)</summary>

The `fuzz/` directory is a separate cargo-fuzz workspace with two libFuzzer targets over the untrusted-input read path. CI runs a 30-second smoke pass of each on every push; locally you can run them longer:

```bash
cargo install cargo-fuzz --locked
cargo +nightly fuzz run read_xlsx        # raw reader
cargo +nightly fuzz run read_xlsx_json   # reader + JSON bridge
```

(Nightly is required because cargo-fuzz needs `-Z` sanitizer flags; the `+nightly` override bypasses the repo's beta pin.)

</details>

## Project Structure

```
crates/
  modern-xlsx-core/       Rust core — OOXML parsing, XML generation, ZIP I/O, encryption (OLE2)
  modern-xlsx-wasm/       WASM bridge — wasm-bindgen exports
packages/
  modern-xlsx/            npm package — TypeScript API, tests, benchmarks
fuzz/                     cargo-fuzz harness for the reader (separate workspace)
examples/                 Runnable examples — Node, Bun, Deno, browser, workers, frameworks
docs/                     Guides, migration docs, feature comparison, plans, GitHub Pages site
```

## Making Changes

1. **Fork** the repository and create a branch from `master`
2. **Write tests** for any new functionality
3. **Run the full test suite** before submitting
4. **Keep commits focused** — one logical change per commit
5. **Follow existing code style** — Biome handles TypeScript formatting, `cargo fmt` handles Rust

## Pull Requests

- Keep PRs focused on a single concern
- Include a clear description of what changed and why
- Reference any related issues
- Ensure CI passes — every gate below runs on each push/PR:

| CI gate | What it checks |
|---------|----------------|
| Rust | `cargo clippy --workspace --all-targets -D warnings` + `cargo test -p modern-xlsx-core` |
| TypeScript matrix | typecheck, lint, tests, build on Node 24/26 × Ubuntu/Windows (with WASM size tracking) |
| Browser tests | Vitest browser mode in Playwright Chromium |
| Security audit | `cargo audit` + `npm audit` |
| cargo-deny | License, duplicate-crate, and registry-source policy (`deny.toml`) |
| Fuzz smoke | 30s libFuzzer pass on both reader targets |
| Package smoke | Packs the tarball, installs it in a clean project, imports the main and `./lite` entry points |

## Reporting Issues

When filing a bug report, please include:

- A minimal reproduction (code snippet or `.xlsx` file)
- Expected vs actual behavior
- Node.js version and OS
- modern-xlsx version

> [!CAUTION]
> Security vulnerabilities should **not** be filed as public issues — see [SECURITY.md](./SECURITY.md) for private reporting via GitHub Security Advisories.

## Architecture Notes

- **Rust core** handles all OOXML parsing/writing, ZIP I/O, and shared string table management
- **WASM bridge** serializes data as JSON strings (faster than `serde_wasm_bindgen` for large workbooks)
- **TypeScript API** provides the developer-facing classes (`Workbook`, `Worksheet`, `Cell`)
- The writer builds the shared string table inline during XML generation (no worksheet clone)
- The `encryption` Cargo feature (default-on) gates the OLE2 module and all crypto dependencies; the lite build disables it

## Code of Conduct

This project follows the [Contributor Covenant](./CODE_OF_CONDUCT.md). Please be respectful and constructive in all interactions.

## License

By contributing, you agree that your contributions will be licensed under the [MIT License](./LICENSE).
