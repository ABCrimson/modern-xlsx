# Contributing to modern-xlsx

Thanks for your interest in contributing! This guide covers everything needed to go from a fresh clone to a passing pull request.

## Development Setup

### Prerequisites

| Tool          | Where the version lives                                           | Notes                                                                                                                                  |
| ------------- | ----------------------------------------------------------------- | -------------------------------------------------------------------------------------------------------------------------------------- |
| **Rust**      | `rust-toolchain.toml`                                             | Pins a dated nightly with clippy, rustfmt and the `wasm32-unknown-unknown` target. The MSRV is `rust-version` in the root `Cargo.toml` |
| **Node.js**   | `engines` in the root `package.json`                              | CI tests Node 24 and 26 on Ubuntu and Windows, plus an advisory Node nightly job                                                       |
| **pnpm**      | `packageManager` in the root `package.json`                       | Inside the repo pnpm runs that pinned version. The only supported package manager (single `pnpm-lock.yaml`)                            |
| **wasm-pack** | `WASM_PACK_VERSION` in `.github/workflows/ci.yml` and `pages.yml` | Builds the WASM bridge crate; CI asserts the exact version                                                                             |

Install each tool at the version its file names, then:

```bash
# Clone
git clone https://github.com/ABCrimson/modern-xlsx.git
cd modern-xlsx

# Install the pinned toolchain; run it in the repo root, as CI does
rustup toolchain install

# Install wasm-pack 0.15.1, the version CI pins (WASM_PACK_VERSION in .github/workflows/ci.yml
# and pages.yml; change this line with them)
cargo install wasm-pack --locked --version 0.15.1

# Install JS dependencies
pnpm install
```

> [!NOTE]
> `rustup toolchain install` with no arguments reads `rust-toolchain.toml` and installs the pinned nightly together with its components and the `wasm32-unknown-unknown` target. CI also runs a beta job (`rust-beta`, with `RUSTUP_TOOLCHAIN=beta`) and an `msrv` job that checks the workspace on the `rust-version` toolchain, so keep the code building on all three.

### Build

> [!IMPORTANT]
> Both WASM artifacts (`wasm/` and `wasm-lite/`) must exist before the TypeScript type check, build or tests can run. `pnpm build` builds both of them and then the TypeScript.

```bash
# All-in-one from the repo root: full + lite WASM, then TypeScript
pnpm build
```

Individual steps and variants (from the repo root):

| Command                 | What it does                                                                                                                                                                               |
| ----------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ |
| `pnpm build:wasm`       | Release WASM build into `packages/modern-xlsx/wasm/`; the release profile runs `wasm-opt -Oz` (see `[package.metadata.wasm-pack.profile.release]` in `crates/modern-xlsx-wasm/Cargo.toml`) |
| `pnpm build:wasm:lite`  | Encryption-free lite WASM into `packages/modern-xlsx/wasm-lite/` (powers the `modern-xlsx/lite` entry point)                                                                               |
| `pnpm build:wasm:debug` | Fast unoptimized WASM build for iteration                                                                                                                                                  |
| `pnpm build:ts`         | TypeScript build via tsdown into `packages/modern-xlsx/dist/`                                                                                                                              |

`pnpm build:wasm` and `pnpm build:wasm:lite` pass `--config "profile.release.trim-paths='all'"` to Cargo, so the `.wasm` holds relative source paths (`crates/modern-xlsx-core/src/...`, `/cargo/registry/<hash>/<crate>-<version>/src/...`) instead of paths on your machine. The option needs Cargo 1.101 or newer, which the nightly pinned in `rust-toolchain.toml` provides; CI's build-wasm job passes the same flag.

### Test and check

These are the commands CI runs, in CI's order. Run them from the repo root.

```bash
# Format check (Rust, including the fuzz workspace)
cargo fmt --all --check
cargo fmt --manifest-path fuzz/Cargo.toml --all --check

# Clippy: default features, no default features, and the parallel feature
cargo clippy --workspace --all-targets --locked -- -D warnings
cargo clippy --workspace --all-targets --no-default-features --locked -- -D warnings
cargo clippy -p modern-xlsx-core --features parallel --all-targets --locked -- -D warnings

# Rustdoc, warnings are errors
RUSTDOCFLAGS="-D warnings" cargo doc -p modern-xlsx-core -p modern-xlsx-wasm --no-deps

# Rust tests: the same three feature sets
cargo test --workspace --locked
cargo test --workspace --no-default-features --locked
cargo test -p modern-xlsx-core --features parallel --locked

# WASM bridge tests (Node), with and without default features
wasm-pack test --node crates/modern-xlsx-wasm
wasm-pack test --node crates/modern-xlsx-wasm --no-default-features
```

The TypeScript checks need both WASM builds (`pnpm build:wasm` and `pnpm build:wasm:lite`; CI's build-wasm job builds both), and some tests read `dist/`, so build before you test:

```bash
pnpm lint                                    # biome check . over the whole repository, as CI's lint job
pnpm -C packages/modern-xlsx typecheck
pnpm -C packages/modern-xlsx typecheck:tests
pnpm -C packages/modern-xlsx lint
pnpm -C packages/modern-xlsx build
pnpm -C packages/modern-xlsx test
pnpm -C packages/modern-xlsx test:coverage   # writes coverage/ (gitignored); not a CI gate
```

Browser tests (Vitest browser mode in Playwright Chromium) and the service worker example:

```bash
pnpm -C packages/modern-xlsx exec playwright install --with-deps chromium
pnpm -C packages/modern-xlsx build
pnpm -C packages/modern-xlsx test:browser
node packages/modern-xlsx/scripts/sw-smoke.mjs
```

To apply Biome formatting across the repo, run `pnpm fmt` (root script: `biome format --write .`). Run `cargo fmt --all` to apply Rust formatting.

### Golden files

Golden JSON files live in `crates/modern-xlsx-core/tests/golden/`, next to the `golden` test binary that checks most of them (`main.rs` plus one module per area); `tests/package_tests.rs` checks one more (`excel_example`). Both compare through `crates/modern-xlsx-core/tests/common/mod.rs`. To rewrite them after an intended output change, set `UPDATE_GOLDEN=1`:

```bash
# Only the golden test binary (not package_tests)
UPDATE_GOLDEN=1 cargo test -p modern-xlsx-core --test golden

# Every golden test: leave out --test
UPDATE_GOLDEN=1 cargo test -p modern-xlsx-core
```

A missing golden file fails the test; it is never created silently without `UPDATE_GOLDEN=1`. Review every regenerated diff before you commit it, and never set `UPDATE_GOLDEN` in CI.

### Fuzzing (optional)

The `fuzz/` directory is a separate cargo-fuzz workspace with six libFuzzer targets in `fuzz/fuzz_targets/`. The repo toolchain is already a nightly, which cargo-fuzz needs for its sanitizer flags, so `cargo fuzz run <target>` works from the repo root with no toolchain override.

```bash
# Install cargo-fuzz 0.13.2, the version CI pins (tool: cargo-fuzz@0.13.2 in .github/workflows/ci.yml;
# change this line with it)
cargo install cargo-fuzz --locked --version 0.13.2

# List the targets, then run one (here read_xlsx) for as long as you like; Ctrl+C stops it
cargo fuzz list
cargo fuzz run read_xlsx
```

CI runs every target from `cargo fuzz list` for 30 seconds each (`-max_total_time=30 -rss_limit_mb=4096`) on every push and pull request. A crash fails the `fuzz` job, but that job is outside the `needs` of the publish job.

## Project Structure

```
crates/
  modern-xlsx-core/       Rust core: OOXML parsing, XML generation, ZIP I/O, encryption (OLE2)
  modern-xlsx-wasm/       WASM bridge: wasm-bindgen exports
packages/
  modern-xlsx/            npm package: TypeScript API, tests, benchmarks
fuzz/                     cargo-fuzz harness: reader, writer, validator and streaming-writer targets (separate workspace)
examples/                 Runnable examples: Node, Bun, Deno, browser, workers, frameworks
scripts/                  Repo tooling: version pins, example smoke tests and type checks, README Quick Start sync check, release prep
.github/workflows/        CI (ci.yml) and the GitHub Pages deploy (pages.yml)
docs/                     Guides, migration docs, feature comparison, roadmap, threat model, GitHub Pages site
```

## Making Changes

1. **Fork** the repository and create a branch from `master`
2. **Write tests** for any new functionality
3. **Run the checks above** before submitting
4. **Keep commits focused**: one logical change per commit
5. **Follow existing code style**: Biome handles TypeScript formatting, `cargo fmt` handles Rust
6. **Record user-visible changes**: add an entry under `## [Unreleased]` in `packages/modern-xlsx/CHANGELOG.md`, the only changelog in the repo. The text at the top of that section, above its `<!-- release-notes:end -->` line (a summary, highlights, upgrade notes and the install line), becomes the GitHub Release body; the full list goes below the line

### Releases and version pins

The version lives only in `packages/modern-xlsx/package.json`; the `VERSION` export reads it at build time. The examples under `examples/`, `docs/site/*.html` and any CDN or `npm:` reference in the two READMEs pin `modern-xlsx` to exactly that version. The release version-bump commit is one command, `node scripts/prepare-release.mjs <version> [--date YYYY-MM-DD]`: it sets the version, rewrites the pins (`scripts/sync-example-pins.mjs`), renames the CHANGELOG `## [Unreleased]` heading to `## [<version>] - <date>` under a fresh `## [Unreleased]` that holds only a `<!-- release-notes:end -->` line, and sets the install line and the tag-pinned repository links above that section's `<!-- release-notes:end -->` line to the new version (it refuses to run when the section has no such line or no text above it), and prints the files it changed and the steps left (build, test, commit, tag `v<version>`, push the tag). It never runs git or touches the network.

CI runs `node scripts/sync-example-pins.mjs --check`, `node scripts/check-readme-sync.mjs` (the root README and the package README share one Quick Start block, byte for byte) and `node scripts/prepare-release.mjs --check`. For a `v*` tag, the `publish` job builds the GitHub Release body from the `## [<version>]` section: the text above `<!-- release-notes:end -->`, a link to the whole CHANGELOG at the tag, a compare link from the release below it and the SHA-384 integrity value of the published `dist/modern-xlsx.min.js`. It fails before `npm publish` when the section, the marker or the text above it is missing, or when the body reaches 100,000 characters (GitHub refuses a release body of about 125,000).

After those checks the job publishes the tarball the `smoke` job tested, creates the GitHub Release, and checks last that the registry shows the provenance attestation. A failed run can be re-run: the publish step skips `npm publish` when the registry already holds the version with the sha512 integrity of that tarball, and fails when it holds the version with another integrity; the release step replaces the title and notes of a Release that already exists for the tag.

## Pull Requests

- Keep PRs focused on a single concern
- Include a clear description of what changed and why
- Reference any related issues
- Ensure CI passes. The jobs in `.github/workflows/ci.yml` are:

| Job id               | What it checks                                                                                                                                                             | Gates publish        |
| -------------------- | -------------------------------------------------------------------------------------------------------------------------------------------------------------------------- | -------------------- |
| `rust`               | Format, three clippy runs, rustdoc with warnings as errors, three Rust test runs, and the WASM bridge tests (with and without default features) on the pinned nightly      | Yes                  |
| `rust-beta`          | Clippy and `cargo test -p modern-xlsx-core --locked` on the beta toolchain                                                                                                 | Yes                  |
| `rust-i686`          | Core lib tests on 32-bit Linux (`i686-unknown-linux-gnu`, gcc-multilib)                                                                                                    | Yes                  |
| `msrv`               | `cargo check --workspace --all-targets --locked` on the `rust-version` toolchain                                                                                           | Yes                  |
| `security`           | `cargo audit`; `pnpm audit --prod` blocks; `pnpm audit --audit-level high` is **advisory** (`continue-on-error`)                                                           | Yes (blocking parts) |
| `deny`               | `cargo deny --locked check`: advisories, bans, licenses and sources (`deny.toml`)                                                                                          | Yes                  |
| `encryption-oracle`  | msoffcrypto-tool decrypts the crate's output (`cargo test -p modern-xlsx-core --release --locked --test interop -- --ignored`)                                             | Yes                  |
| `fuzz`               | 30 s libFuzzer pass on every target; a crash fails the job                                                                                                                 | No                   |
| `build-wasm`         | Full and lite WASM builds with gzip size budgets                                                                                                                           | Yes                  |
| `smoke`              | Packs the tarball; `publint --strict`; `attw --profile esm-only`; no `node:` imports in the browser bundles; installs the tarball in a clean project and runs a smoke test | Yes                  |
| `lint`               | `pnpm exec biome check .` over the whole repository (root `biome.json`): docs, examples, config and scripts                                                                | Yes                  |
| `typescript`         | typecheck, typecheck:tests, lint, build, test on Node 24 and 26, Ubuntu and Windows                                                                                        | Yes                  |
| `typescript-nightly` | The same steps on a Node nightly; **advisory** (`continue-on-error`)                                                                                                       | No                   |
| `browser`            | Vitest browser mode in Playwright Chromium, then the service worker example                                                                                                | Yes                  |
| `examples`           | Pin, README-sync and release-metadata checks, then the Node, Bun, Deno and Workers examples against the packed tarball                                                     | Yes                  |
| `examples-types`     | Type-checks the framework example snippets against the tarball                                                                                                             | Yes                  |
| `publish`            | Runs only for a `v*` tag; every job marked Yes above is in its `needs`; release notes from the CHANGELOG, size-checked, plus the CDN integrity value                       | n/a                  |

## Reporting Issues

When filing a bug report, please include:

- A minimal reproduction (code snippet or `.xlsx` file)
- Expected vs actual behavior
- Node.js version and OS
- modern-xlsx version

> [!CAUTION]
> Security vulnerabilities should **not** be filed as public issues. See [SECURITY.md](./SECURITY.md) for private reporting through GitHub private vulnerability reporting.

## Architecture Notes

- **Rust core** handles all OOXML parsing/writing, ZIP I/O, and shared string table management
- **WASM bridge** serializes data as JSON strings
- **TypeScript API** provides the developer-facing classes (`Workbook`, `Worksheet`, `Cell`)
- The writer builds the shared string table inline during XML generation (no worksheet clone)
- The `encryption` Cargo feature (default-on) gates `ole2::crypto`, the crypto dependencies, and the `readWithPassword` and `writeWithPassword` WASM exports; the lite build disables it. OLE2 detection always compiles, so the lite build still reports encrypted and legacy `.xls` files precisely

## Code of Conduct

This project follows the [Contributor Covenant](./CODE_OF_CONDUCT.md). Please be respectful and constructive in all interactions.

## License

By contributing, you agree that your contributions will be licensed under the [MIT License](./LICENSE).
