# WASM size and speed: the 2026-11-12 measurement

This page records how the size of the two WebAssembly builds was measured for 1.3.0, which source changes were kept to shrink them, and why the release profile keeps `opt-level = 3` instead of the smaller `opt-level = "s"`. The CI size budgets in `.github/workflows/ci.yml` (`WASM_GZIP_BUDGET_FULL`, `WASM_GZIP_BUDGET_LITE`) are set from the final figures here.

## Setup

- **Date:** 2026-11-12, on the 1.3.0 release branch. Dates on this page are UTC calendar days.
- **Machine:** a Windows 11 laptop with an Intel Core Ultra 9 290HX Plus. Nothing else was building or benchmarking while it ran.
- **Toolchain:** `nightly-2026-11-12` (rustc 1.102.0-nightly 3ef0c22cd), wasm-bindgen 0.2.130, wasm-pack 0.15.1 (Binaryen version_130, `wasm-opt -Oz` from `crates/modern-xlsx-wasm/Cargo.toml`), Node v26.13.0, Vitest 5.1.0. The root `[profile.release]` sets `lto = "fat"`, `codegen-units = 1`, `strip = "debuginfo"` and `panic = "abort"`.

## How size is measured

Size is always measured with the two commands of the CI `build-wasm` job, run from the repository root, followed by `gzip`:

```sh
wasm-pack build crates/modern-xlsx-wasm --target web --release --out-dir ../../packages/modern-xlsx/wasm -- --config "profile.release.trim-paths='all'"
wasm-pack build crates/modern-xlsx-wasm --target web --release --out-dir ../../packages/modern-xlsx/wasm-lite -- --no-default-features --config "profile.release.trim-paths='all'"
gzip -9 -n -c packages/modern-xlsx/wasm/modern_xlsx_wasm_bg.wasm | wc -c
gzip -9 -n -c packages/modern-xlsx/wasm-lite/modern_xlsx_wasm_bg.wasm | wc -c
```

To try another optimisation level, append a second flag to the cargo arguments, for example `--config "profile.release.opt-level='s'"`.

## Source changes

Each change was measured on the full build against the commit before it. A change was kept only if it saved at least 400 bytes gzip and the workspace tests and `cargo clippy --workspace --all-targets --locked -- -D warnings` passed.

| Change                                                                                                                                                   | Full, raw bytes | Full, gzip bytes | Saving (gzip) | Kept                                                       |
| -------------------------------------------------------------------------------------------------------------------------------------------------------- | --------------: | ---------------: | ------------: | ---------------------------------------------------------- |
| Starting point                                                                                                                                           |       2,330,423 |          857,509 |               |                                                            |
| The reader's per-sheet part walk is compiled once, through a `&mut dyn FnMut` seam, instead of once for each of five item types                          |       2,317,296 |          855,087 |         2,422 | yes                                                        |
| The map of preserved content types is built with `insert` instead of `collect`, so it no longer compiles a stable sort                                   |       2,309,538 |          851,121 |         3,966 | yes                                                        |
| Worksheet children that are not read are skipped with quick-xml's buffered `read_to_end_into`, which the chart, pivot and styles parsers already compile |       2,304,269 |          849,343 |         1,778 | yes                                                        |
| The colour check formats a location string only when a colour is bad, in one cold function                                                               |       2,303,759 |          848,482 |           861 | yes                                                        |
| `SharedStringTableBuilder::write_to` writes through `&mut dyn Write`, so one body serves both writers                                                    |       2,302,471 |          848,044 |           438 | yes                                                        |
| `#[inline(never)]` on `attr_str`                                                                                                                         |       2,291,553 |          847,833 |           211 | no: under 400 bytes, and the 100K-row read was 2.8% slower |

The lite build went from 794,487 to 783,937 bytes gzip with these changes.

## Speed: opt-level 3, "s" and "z"

The two benchmark files that run the WASM read and write paths are `packages/modern-xlsx/__tests__/bench/read-write.bench.ts` and `packages/modern-xlsx/__tests__/bench/workbook-throughput.bench.ts`. They were run against the built package (`tsdown`, then `packages/modern-xlsx/scripts/finalize-dist.mjs`) with:

```sh
pnpm -C packages/modern-xlsx exec vitest bench --run --reporter=verbose __tests__/bench/<file>.bench.ts
```

Five builds of the full WASM were compared. They differ only in `modern_xlsx_wasm_bg.wasm`, and the JavaScript glue is the same for all of them.

- **3:** the final tree at `opt-level = 3`.
- **s and z:** the same tree at `opt-level = "s"` and `"z"`.
- **no (d):** the tree before the `write_to` change.
- **+ (e):** the tree with the rejected `attr_str` change added.

Each round ran both files once for every build. The build order was rotated by one place each round, so that drift on the host affected all builds alike. There were five rounds. The table shows the median operations per second (hz) of each case over the five rounds. Each Vitest run took 5 samples of each case, after one warm-up.

| Case                                                   | Runs WASM    |     3 |     s |     z | no (d) | + (e) | s vs 3 | z vs 3 | no (d) vs 3 | + (e) vs 3 |
| ------------------------------------------------------ | ------------ | ----: | ----: | ----: | -----: | ----: | -----: | -----: | ----------: | ---------: |
| Write 10K rows x 10 cols (cell by cell)                | yes          | 7.519 | 6.763 | 6.285 |  7.387 | 7.313 | -10.1% | -16.4% |       -1.8% |      -2.7% |
| Read 10K rows x 10 cols                                | yes          | 23.15 | 22.19 | 16.66 |  22.63 | 23.15 |  -4.2% | -28.1% |       -2.3% |       0.0% |
| Read 100K rows x 10 cols                               | yes          | 2.058 | 1.960 | 1.534 |  1.943 | 2.000 |  -4.8% | -25.5% |       -5.6% |      -2.8% |
| Write 100K rows x 10 cols (`sheetAddAoa` + `toBuffer`) | yes          | 0.667 | 0.618 | 0.550 |  0.659 | 0.665 |  -7.4% | -17.6% |       -1.2% |      -0.4% |
| Write 1K rows x 5 cols                                 | yes          | 248.4 | 222.6 | 177.0 |  248.2 | 231.4 | -10.4% | -28.8% |       -0.1% |      -6.8% |
| Read 1K rows x 5 cols                                  | yes          | 312.0 | 266.6 | 256.6 |  299.6 | 308.5 | -14.6% | -17.8% |       -4.0% |      -1.1% |
| Write 10K rows x 5 cols                                | yes          | 22.99 | 21.56 | 18.56 |  22.96 | 23.34 |  -6.2% | -19.2% |       -0.1% |       1.5% |
| Read 10K rows x 5 cols                                 | yes          | 44.07 | 42.88 | 32.27 |  41.90 | 44.22 |  -2.7% | -26.8% |       -4.9% |       0.3% |
| `aoaToSheet` 50K rows x 5 cols                         | no (JS only) | 12.99 | 12.78 | 12.10 |  13.26 | 13.32 |  -1.6% |  -6.9% |        2.0% |       2.5% |
| `sheetToCsv` 10K rows x 10 cols                        | no (JS only) | 38.32 | 39.55 | 39.20 |  41.34 | 40.83 |   3.2% |   2.3% |        7.9% |       6.6% |
| `sheetToJson` 10K rows x 10 cols                       | no (JS only) | 43.66 | 42.25 | 44.46 |  40.83 | 43.67 |  -3.2% |   1.8% |       -6.5% |       0.0% |

The three JavaScript-only cases do not call into WASM. They show how much the numbers move between runs on this host, up to about 8%. A difference of a few percent in a single case, such as the "no (d)" column, is within that noise. The opt-level losses are not: under `"s"` all eight WASM cases are slower, and under `"z"` they are slower by 16-29%.

Over the eight WASM read and write cases:

| Build vs 3 |                       Worst case | Geometric mean |
| ---------- | -------------------------------: | -------------: |
| s          |   -14.6% (read 1K rows x 5 cols) |          -7.6% |
| z          |  -28.8% (write 1K rows x 5 cols) |         -22.7% |
| no (d)     | -5.6% (read 100K rows x 10 cols) |          -2.5% |
| + (e)      |   -6.8% (write 1K rows x 5 cols) |          -1.5% |

Size of the same tree at each optimisation level (CI commands, gzip -9 -n):

| opt-level | Full, raw bytes | Full, gzip bytes | Lite, raw bytes | Lite, gzip bytes |
| --------- | --------------: | ---------------: | --------------: | ---------------: |
| 3         |       2,302,471 |          848,044 |       2,120,867 |          783,937 |
| "s"       |       2,027,602 | 750,930 (-11.5%) |       1,924,646 |  708,928 (-9.6%) |
| "z"       |       1,716,376 | 658,171 (-22.4%) |       1,615,477 | 619,668 (-21.0%) |

## Decision

The rule set beforehand was to adopt `opt-level = "s"` only if no read or write case lost more than 5% of its median throughput against `opt-level = 3`, and the geometric mean lost no more than 3%. `"z"` was measured but was never a candidate.

`"s"` loses 14.6% in its worst case and 7.6% on the geometric mean, so it fails both limits. **The release profile keeps `opt-level = 3`.** The `write_to` change was kept: no write case slowed by more than 2%. The `attr_str` change was dropped.

## Final figures and budgets

These are the final sizes at `opt-level = 3`, measured on 2026-11-14 (a UTC date) with the commands above, `nightly-2026-11-14` (rustc 1.102.0-nightly 5f4effd57) and wasm-pack 0.15.1 on a clean copy of the final 1.3.0 release tree (`git archive`, fresh target directory), with the five kept changes and the Rust fixes merged after them (the hardening of the preserved note-drawing merge, the row-height, print-area, threaded-comment and pivot-table fixes, the default sheet view and the part-name allocator, and the pivot cache-source, table-name, table totals-cell, threaded-comment limit and generated-note ceiling fixes, and the last ones: the shared-formula writer check, the threaded-comment `dT` form, the bounded validation report and linear totals check, the shared-text read budget, the defined-name rules and table-formula check, and the kept pivot-cache relationships and part-name letter case, and then the JSON read checks after each cell, shared string, comment, hyperlink, kept extension and field with the table-driven JSON escape, the namespace declarations counted on kept extensions, the linear table-formula shorthand check, the combining-mark and canonical-equivalence defined-name rules, the shared-formula master rule, the pivot relationship and kept cache-source entity checks, the threaded-comment year edges, and the fixed-name parts found in any letter case, and last the cell children handed to the sink as they end, the worksheet list memory charge, the JSON buffer clamp and the plain-string fast path of the JSON escape, and after them the style-record charge, the page breaks, auto filter and tables written to the read JSON one element at a time, and the cell type of a date cell with several values, and last the shared-string, chart series and axis, threaded-comment and defined-name charges, the namespace index of a part's root and the rich-text runs of a replaced shared-string value on the JSON read path, and after them the inherited `<ext>` declarations copied from a rendered index of the root, the `<extLst>` declarations bound in each kept `<ext>` and the streaming writer's one-lookup sheet-name check, and the toolchain pin moved from `nightly-2026-11-13` to `nightly-2026-11-14`, and last the kept-fragment namespace fixes: a colon-leading name gains no default namespace, a fragment's own declarations are looked up in a sorted list first, and open elements are counted only when a use needs them):

- **Full:** 2,485,128 bytes raw, 917,326 bytes gzip.
- **Lite:** 2,301,840 bytes raw, 854,256 bytes gzip.

Earlier on 2026-11-14 (UTC), before the kept-fragment namespace fixes, the same commands and toolchain gave 2,467,293 bytes raw and 909,311 bytes gzip for the full build and 2,283,982 and 846,316 for the lite build. On the release tree of 2026-11-13 (UTC), before the toolchain pin moved to `nightly-2026-11-14`, the same commands on `nightly-2026-11-13` (rustc 1.102.0-nightly 8eb778d04) gave 2,467,339 bytes raw and 909,561 bytes gzip for the full build and 2,284,040 and 846,414 for the lite build; the earlier figures below are from `nightly-2026-11-13` too. Earlier on 2026-11-13, before the inherited-declaration index, the `<extLst>` declarations and the streaming writer's sheet-name lookup, the same commands and toolchain gave 2,456,722 bytes raw and 905,163 bytes gzip for the full build and 2,273,393 and 841,798 for the lite build. Earlier on 2026-11-13, before the shared-string, chart, threaded-comment and defined-name charges and the fixes after them, the same commands and toolchain gave 2,444,121 bytes raw and 899,119 bytes gzip for the full build and 2,260,791 and 836,693 for the lite build. Earlier on 2026-11-13, before the style-record charge and the fixes after it, the same commands and toolchain gave 2,441,465 bytes raw and 897,139 bytes gzip for the full build and 2,258,078 and 833,954 for the lite build. Earlier on 2026-11-13, before the group that ends with the plain-string fast path, the same commands and toolchain gave 2,433,092 bytes raw and 895,957 bytes gzip for the full build and 2,249,709 and 832,662 for the lite build. The table of source changes, the speed tables and the size-per-opt-level table above were measured with `nightly-2026-11-12` at the commit before those fixes. With that toolchain the final tree as it stood on 2026-11-13 (UTC), with those fixes, measured 2,433,201 bytes raw and 895,805 bytes gzip for the full build and 2,249,814 and 832,586 for the lite build, so the fixes made the full build 130,730 bytes raw and 47,761 bytes gzip larger and the lite build 128,947 bytes raw and 48,649 bytes gzip larger. Earlier measurements of the final tree, both on 2026-11-12, gave 2,393,777 bytes raw and 882,046 bytes gzip for the full build and 2,212,145 and 818,150 for the lite build before the last group of fixes, and 2,351,370 and 865,954 for the full build and 2,169,731 and 802,377 for the lite build before the group before it.

The CI budgets are `ceil(gzip size x 1.05)`: 963,193 bytes for the full build and 896,969 bytes for the lite build. A budget is a tripwire that catches a regression. It is set from this measurement and is not a size target.
