# Security Policy

## Supported Versions

Security fixes go into the newest minor release, the one on the npm `latest` dist-tag. Older minors do not receive fixes; upgrade to the latest release.

## Reporting a Vulnerability

If you discover a security vulnerability, please report it responsibly:

1. **Do not** open a public issue.
2. Report it privately to the maintainer, Albert Badalov ([ABCrimson](https://github.com/ABCrimson)), through [GitHub private vulnerability reporting](https://github.com/ABCrimson/modern-xlsx/security/advisories/new) (on the repository page: the **Security** tab, then **Report a vulnerability**).
3. Include a description of the vulnerability and steps to reproduce.

We will acknowledge receipt within 48 hours and aim to release a fix within 7 days for critical issues.

## Scope

modern-xlsx processes untrusted `.xlsx` files. The full list of defences, with the test that proves each, is in the [threat model](docs/security/threat-model.md). In short:

- **ZIP decompression.** `read_zip_entries` in `crates/modern-xlsx-core/src/zip/reader.rs` enforces the size, ratio and entry-count limits below with `SECURITY` errors; the workbook JSON length is checked by `check_json_len` in `crates/modern-xlsx-core/src/reader.rs` as the JSON is built, and exceeding it is a `SIZE_LIMIT` error. File entry names containing a NUL byte, a leading `/` or `\`, a `..` component or a drive prefix, and duplicate entry names, are rejected with a `SECURITY` error (directory entries are skipped and never read).
- **XML parsing.** SAX-style parsing with `quick-xml`. A DOCTYPE is accepted, but the entity declarations in its internal subset are never expanded and external entities are never resolved. Only the five predefined XML entities and character references are decoded. In element text any other named entity reference is dropped (`push_general_ref`); in an attribute value it is an `XML_PARSE` error (`attr_str`), both in `crates/modern-xlsx-core/src/ooxml/xml.rs`. `doctype_case` in `crates/modern-xlsx-core/tests/hostile_inputs.rs` and `packages/modern-xlsx/__tests__/hostile-input.test.ts` cover this.
- **Encrypted files (OLE2).** Sector chain loops and oversized declared stream sizes give `SECURITY` errors (`crates/modern-xlsx-core/src/ole2/cfb/reader.rs`). An Agile `spinCount` above 10,000,000 (`MAX_SPIN_COUNT` in `ole2/encryption_info.rs`) is rejected.
- **Memory safety.** The workspace sets `unsafe_code = "forbid"` (`[workspace.lints.rust]` in `Cargo.toml`), and the Rust core runs as WebAssembly.
- **Formula engine (TypeScript).** The parser limits formula length and nesting depth (`MAX_FORMULA_LENGTH`, `MAX_NESTING_DEPTH` and `MAX_FUNCTION_DEPTH` in `packages/modern-xlsx/src/formula/parser.ts`). CSV export escapes formula-like cells only when the `escapeFormulas` option is set (`packages/modern-xlsx/src/io/sheet-utils.ts`); set it when the CSV will be opened in a spreadsheet.

## Default read limits

These defaults come from `ZipSecurityLimits::default` in `crates/modern-xlsx-core/src/zip/reader.rs`.

| Limit                       | Default                     | Ceiling                                                                   |
| --------------------------- | --------------------------- | ------------------------------------------------------------------------- |
| Total decompressed size     | 2 GiB (2,147,483,648 bytes) | 2 GiB (`MAX_DECOMPRESSED_SIZE_CEILING`)                                   |
| Per-entry compression ratio | 100                         | none                                                                      |
| Number of entries           | 100,000                     | 1,048,576 (`MAX_ENTRIES_CEILING`)                                         |
| Workbook JSON length        | 536,870,888 bytes           | 536,870,888 bytes (`MAX_JSON_LENGTH_CEILING`, V8's maximum string length) |

Callers change the limits through `ReadOptions.limits` (`maxDecompressedSize`, `maxCompressionRatio`, `maxEntries` and `maxJsonLength`; see `ReadLimits` in `packages/modern-xlsx/src/model/types/cells.ts` and `packages/modern-xlsx/src/read-limits.ts`). A value above a ceiling is clamped to it (`ZipSecurityLimits::clamped`); the TypeScript API throws `INVALID_ARGUMENT` for an unknown key, for a value that is not a finite number greater than zero, or for a `maxDecompressedSize`, `maxEntries` or `maxJsonLength` that is not a safe integer. `ZipSecurityLimits::clamped`, which the WASM bridge applies, also turns a non-positive or NaN compression ratio into the default.
