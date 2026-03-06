# Contributing to modern-xlsx

Thanks for your interest in contributing! This guide will help you get started.

## Development Setup

### Prerequisites

- **Rust** 1.95.0+ with `wasm32-unknown-unknown` target
- **Node.js** 25+
- **pnpm** 11+
- **wasm-pack** (for building the WASM bridge)

```bash
# Install Rust target
rustup target add wasm32-unknown-unknown

# Install wasm-pack
cargo install wasm-pack

# Clone and install
git clone https://github.com/ABCrimson/modern-xlsx.git
cd modern-xlsx
pnpm install
```

### Build

```bash
# Build WASM (required before TypeScript tests)
cd crates/modern-xlsx-wasm
wasm-pack build --target web --release --out-dir ../../packages/modern-xlsx/wasm --no-opt

# Build TypeScript
pnpm -C packages/modern-xlsx build
```

### Test

```bash
# Rust tests (389 tests)
cargo test -p modern-xlsx-core

# TypeScript tests (1,230 tests)
pnpm -C packages/modern-xlsx test

# Lint
cargo clippy -p modern-xlsx-core -- -D warnings
pnpm -C packages/modern-xlsx lint

# Type check
pnpm -C packages/modern-xlsx typecheck
```

## Project Structure

```
crates/
  modern-xlsx-core/       Rust core — OOXML parsing, XML generation, ZIP I/O
  modern-xlsx-wasm/       WASM bridge — wasm-bindgen exports
packages/
  modern-xlsx/            npm package — TypeScript API, tests, benchmarks
```

## Making Changes

1. **Fork** the repository and create a branch from `main`
2. **Write tests** for any new functionality
3. **Run the full test suite** before submitting
4. **Keep commits focused** — one logical change per commit
5. **Follow existing code style** — Biome handles TypeScript formatting, `cargo fmt` handles Rust

## Pull Requests

- Keep PRs focused on a single concern
- Include a clear description of what changed and why
- Reference any related issues
- Ensure CI passes (Rust tests, TypeScript tests, lint, type check)

## Reporting Issues

When filing a bug report, please include:

- A minimal reproduction (code snippet or `.xlsx` file)
- Expected vs actual behavior
- Node.js version and OS
- modern-xlsx version

## Architecture Notes

- **Rust core** handles all OOXML parsing/writing, ZIP I/O, and shared string table management
- **WASM bridge** serializes data as JSON strings (faster than `serde_wasm_bindgen` for large workbooks)
- **TypeScript API** provides the developer-facing classes (`Workbook`, `Worksheet`, `Cell`)
- The writer builds the shared string table inline during XML generation (no worksheet clone)

## Code of Conduct

This project follows the [Contributor Covenant](./CODE_OF_CONDUCT.md). Please be respectful and constructive in all interactions.

## License

By contributing, you agree that your contributions will be licensed under the [MIT License](./LICENSE).
