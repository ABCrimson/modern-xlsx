# msoffcrypto-tool interop fixtures

These files are byte-exact copies from
[nolze/msoffcrypto-tool](https://github.com/nolze/msoffcrypto-tool) v6.0.0,
MIT License, Copyright (c) 2015 nolze. The license text is in [`LICENSE`](LICENSE).

modern-xlsx uses them as interop fixtures: they were
encrypted by other software, not by modern-xlsx, so they test the reader against a
second implementation instead of a self round trip.

The password of both encrypted files is `Password1234_`.

| File                                  | Upstream path                                       | Content                                                                                 | SHA-256                                                            |
| ------------------------------------- | --------------------------------------------------- | --------------------------------------------------------------------------------------- | ------------------------------------------------------------------ |
| `example_password.xlsx`               | `tests/inputs/example_password.xlsx`                | Agile encryption (AES-256, SHA-512); `EncryptionInfo` (1289 B) lives in the mini stream | `3f792e3902a615bf0e91771f3f3016b80d59098b8e492efcdc0752748e15b997` |
| `example.xlsx`                        | `tests/outputs/example.xlsx`                        | The decrypted ZIP of `example_password.xlsx`                                            | `4dd9dd0ccbfc7fb8769f1f3307830d3cc4c5042e32d619f4b2835fada89d13c6` |
| `ecma376standard_password.docx`       | `tests/inputs/ecma376standard_password.docx`        | Standard encryption (AES-128, SHA-1); `EncryptionInfo` is 224 B                         | `d265dcf02f7d552486229b8c67a627ef3752fe6802b02fd2d2d485fcfbbac5de` |
| `ecma376standard_password_plain.docx` | `tests/outputs/ecma376standard_password_plain.docx` | The decrypted ZIP of `ecma376standard_password.docx`                                    | `ca1c0ebb465553361b9034e696d4081df0a2d41918f820060325b3ca634eb69b` |

`example.xlsx` is an Excel-authored workbook. It carries Excel's default
`xl/styles.xml`: theme font colours, `cellStyleXfs` with `xfId` references from
`cellXfs`, a `tableStyles` element and an `extLst` with x14/x15 extensions. The
styles tests in `src/ooxml/styles/tests.rs` and the package golden in
`tests/package_tests.rs` read that file.

Check a copy with `sha256sum <file>`.
