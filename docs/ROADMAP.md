# Roadmap

This page lists the work planned for modern-xlsx after the current release, grouped by the minor release it is aimed at. These are plans, not commitments: an item can move between releases, and nothing here carries a date.

What has shipped is recorded in the [changelog](../packages/modern-xlsx/CHANGELOG.md), newest release section first at the top of that file; this page does not repeat it.

## Policy for 1.x

- **Additive shape changes.** Every change to the workbook JSON and the TypeScript types in 1.x is additive. New fields are optional, and a field that is replaced keeps being read and written beside its replacement until 2.0: `dynamicArray` beside `cellMetadata`, `formulaDtr1` and `formulaDtr2` beside `formulaDel1` and `formulaDel2`.
- **Deprecations name their removal.** A deprecated field or type member says so in its documentation and names 2.0 as the release that removes it.
- **Releases come from tags.** CI publishes each release from a `v*` tag through npm trusted publishing with provenance, then creates the GitHub Release from that version's changelog section.
- **Size budget.** CI fails when the gzip size of the full or the lite WASM build exceeds its budget.
- **Hostile input.** The hostile-input and security test suites run in every CI build, and that job gates every release.
- **Very large workbooks ([issue #31](https://github.com/ABCrimson/modern-xlsx/issues/31)).** `ReadOptions.limits` bounds a read, and a workbook whose JSON would pass `maxJsonLength` fails with a `SIZE_LIMIT` error. The full fix, a streaming read that never builds one JSON string for the whole workbook, is planned for 1.4.0.

## 1.4.0: Big data, streaming read and a lean bridge

- An async row iterator per sheet over the Rust parser, with JSON built per sheet or per chunk, so no single string holds the whole workbook. This closes [issue #31](https://github.com/ABCrimson/modern-xlsx/issues/31). `streamToJson` and `streamToCsv` sit on top of it.
- `XlsxWorker.readBuffer` accepts `ReadOptions.limits`, and the worker's `readBuffer` and `writeBuffer` take `ReadOptions` and `WriteOptions` instead of an inline `{ password }` object.
- Lazy sheet loading: a sheet is parsed when it is first used.
- An opt-in byte transport for preserved parts, so they stop crossing the WASM boundary as JSON number arrays.
- An opt-in way to leave the shared-string table out of the read JSON; the writer ignores it.
- The shared strings and styles parts found through the workbook's relationships instead of their fixed names, and a part a relationship names found in any ASCII letter case, as Excel finds it; today a package whose relationship names another shared strings part (`xl/sst.xml`) reads every shared-string cell as its index, and one whose relationship names `xl/worksheets/sheet1.xml` while the part is stored as `xl/worksheets/Sheet1.xml` fails with `MISSING_PART`. With it, a fixed-name part stored in another letter case (`xl/theme/Theme1.xml`) keeps the content type its package declares in `preservedContentTypes`; today it gets `application/xml`.
- A `preservedEntries` name that is not a valid part name (percent-encoded, holding a backslash, ending in a dot or a space, starting with `/` or `../`) skipped on write with a warning; today it is written as given.
- A configurable ZIP compression level for writing.
- `StreamingXlsxWriter` output to a Web or Node stream instead of one buffer at `finish()`.
- A row-height option for `StreamingXlsxWriter`; today a streamed row's height follows the font size of its style.
- One copy of the full WASM in the npm tarball: today it ships `wasm/modern_xlsx_wasm_bg.wasm` and, for the worker and the IIFE bundle, the same 2.5 MB file again as `dist/modern-xlsx.wasm`.
- `validate()` finds each pivot table's cache by one lookup, as the writer does, instead of scanning every cache for every table.
- An auto filter's colour, icon, top-ten, dynamic (above-average and date-period) and date-group filters and its `sortState` kept through a read and a write; today they are dropped, and Excel shows those columns' filters off in the written file.
- The read JSON written one element at a time for every workbook-level field (the styles, shared strings and defined names among them) and for each sheet's charts, pivot tables, threaded comments, slicers and timelines, each measured against a flood first; today each is serialized whole into a temporary string before the length check, so one field of control characters (6 bytes of JSON per byte of XML) can still trap the WebAssembly build.
- A read that runs out of memory in the WebAssembly build investigated, so that it fails with an error and leaves the module usable, instead of each parser being given its own memory charge.
- Pivot cache fields, pivot fields, table columns, persons and the chart elements the model does not hold counted against the read's one memory budget, as chart series and axes are; today they are bounded only by their part's decompressed size, and pivot cache fields beside a shared-string flood brought the WebAssembly memory within 94 MiB of its 4 GiB ceiling at about 2.0 MB (measured on 2026-11-13).
- Rich shared strings read through the streaming read and counted at a smaller run record, so a workbook of a million rich text cells of four runs each, which 1.2.1 read, reads again; today 765,000 or more such cells fail with `SIZE_LIMIT` (measured on 2026-11-14; the first refused count depends on the runs' formatting), and `maxJsonLength` cannot raise the limit.
- The relationship Id index (`id_index` in `relationships.rs`) and the other read-path maps over part names and Ids (the content-type defaults, the sheet relationships, the pivot cache Ids, the known parts and the threaded-comment roots) moved off the standard hasher, whose keys are not secret in the WebAssembly build.
- The WebAssembly memory grown in larger steps during a read, so millions of rich shared-string entries and kept worksheet extensions that inherit many namespace declarations read in linear time; today 2,800,000 one-run entries take 3.5 s where a read with the memory grown first took 0.66 s (measured on 2026-11-13), and 1,000 inherited prefixes on each of 11,500 kept `<ext>` take about 3.7 s where the same JSON without insertions takes 3.0 s (measured on 2026-11-14).
- A default namespace declared on a worksheet's `<worksheet>` root or `<extLst>` declared on each kept `<ext>` whose unprefixed names use it; today such a name takes the default namespace of the sheet it is written into.
- A decision on whether a crafted cell with an `<is>` after a rich shared-string `<v>` should keep that entry's rich text runs; today both read paths keep them.
- Internal, no public API change: the ZIP writer's entry type gets a name that no longer clashes with the `ZipEntry` error variant, the test modules stop mixing `pretty_assertions` and the standard assertion macros, and the read JSON escape writes U+0008 and U+000C as `\b` and `\f`, as serde_json does.
- Carried over from 1.3.0:
  - Text holding a lone UTF-16 surrogate makes `toBuffer()` fail with a `JSON` error; it will be written as U+FFFD.
  - Text longer than Excel's per-cell limit is written without any error or warning; it will be reported.
  - Internal test and lint work: a golden file for the encrypted-package structure, part and relationship assertions for password-protected writes, error assertions for malformed `EncryptionInfo` input, a base64 decode round trip, the `[JSON]` error code checked at the WASM bridge, and `Option<&T>` parameters with Clippy's `ref_option` lint enabled.

## 1.5.0: Formats in, CSV/TSV and ODS

- CSV and TSV read into a workbook, with RFC 4180 quoting and delimiter and encoding detection.
- HTML table import.
- ODS read and write in Rust, keeping unknown parts as preserved blobs.
- Input format detection across XLSX, ODS and CSV.
- `sheetToAoa`, plus `tableToJson` and `jsonToTable` for a single Excel table.
- `_xHHHH_` escapes in cell text decoded on read and encoded on write.

## 1.6.0: Formula engine 2

- A much larger function library: `SUMIFS`, `COUNTIFS`, `MEDIAN`, `XLOOKUP`, `FILTER`, `SORT`, `UNIQUE`, more text functions, the date functions, `NA`, `IFNA`, `LOG10` and `CHAR`.
- Defined names in evaluation, through the evaluation context.
- Structured (table) references and 3-D references evaluate; today they tokenize, parse and serialize, and evaluate to `#REF!`.
- Dynamic-array spill, and functions that return references: `OFFSET`, `INDIRECT`, and `INDEX` with a zero row or column inside an aggregate such as `SUM`.
- `VALUE` and text-to-number coercion accept currency, percent and date or time text.
- `compileFormula()` with a parsed-formula cache.
- `calculateWorkbook()`, which evaluates every formula in dependency order and writes the cached values back, and `getFormulaDependencies()`.
- A table helper that adds a calculated column and fills its formula.
- The formula tokenizer accepts every character Excel accepts in a table or defined name, so a formula that uses a name holding `\`, `?` or a non-ASCII symbol (`a\b`, `a?b`, `Ventes€`) parses, evaluates, and is shifted by `rewriteFormula` and `expandSharedFormula`; today they return it unchanged.
- `addNamedRange` accepting every character Excel accepts in a defined name (a symbol such as `a€b`, a number character such as `a²b`, an emoji, a soft hyphen, a zero-width joiner or non-joiner inside a word, a Thai mark first) and `addTable` accepting a table name that starts with a Thai mark, `validate()` reporting every name Excel refuses (`ab:`, a leading `€`, a no-break space), and the R1C1 name rule checked against Excel in a language whose R1C1 letters differ from the English `R` and `C`.
- `validate()` comparing table and defined names ignoring canonical equivalence as well as case, as `addTable` and `addNamedRange` do, and reporting a table or defined name that starts with a combining mark (U+0301 + `abc`), which Excel refuses for some scripts; `validate()` and `repair()` comparing sheet names the same way, with the Greek final sigma folded, so sheets `café` (U+00E9) and `cafe` + U+0301, or `ΑΣ` and `ασ`, are reported and renamed, and `validate()` finding the table a `[@...]` names that way, so `Cafe` + U+0301 + `[@Home]` for a column `@Home` of table `Café` is not reported as the `@` shorthand.
- `addSheet`, `renameSheet`, the `Worksheet.name` setter, `cloneSheet`, `addTable`, `addNamedRange` and `validate()` comparing names as Excel does beyond case and canonical equivalence, so a pair such as `Straße` and `STRASSE`, U+FB01 + `x` and `fix`, or U+03F5 and U+03B5 is one name, and a Hangul syllable beside its conjoining jamo (U+D55C and U+1112 U+1161 U+11AB), which Excel opens as two sheets and keeps apart as defined names, is two sheet names and two defined names (Excel keeps only one of two such tables, so `addTable` keeps refusing the second), and the pairs whose lower case alone joins them (Cherokee U+13A0 and U+AB70, Georgian U+10D0 and U+1C90, Deseret U+10400 and U+10428, Adlam U+1E900 and U+1E922, Osage U+104B0 and U+104D8, Greek U+03F4 beside U+03B8), which Excel opens as two sheets, are two sheet names; today the first pairs are taken as two names and the others as one. With it, `cloneSheet` of a sheet holding two tables named by such a Hangul pair points each copied table's references at its own copy.
- The lookups by name (`getSheet`, `removeSheet`, `hideSheet`, `unhideSheet`, the sheet-name argument of `renameSheet` and `addImage`, `getTable`, `removeTable`, `getNamedRange`, `removeNamedRange`) matching a name as the add methods compare it, ignoring case and canonical equivalence; today they match it exactly, so `getSheet('DATA')` does not find `Data` although `addSheet('DATA')` is refused as taken.
- `getSheet` answering from the sheet data, so a sheet whose name was changed by a direct write to the object `toJSON()` hands out is found by its new name; today it is found by its old one, while `addSheet` already refuses the new one as taken.
- The duplicate checks of `addSheet`, `addTable` and `addNamedRange` answered from their name index without reading every stored name; today each call computes one key but still compares every stored name once, so many calls take quadratic time with a small constant (50,000 `addTable` calls 4,066 ms, 50,000 `addNamedRange` calls 3,373 ms, Node v26.13.0, measured on 2026-11-13), and a removal other than of the last name, a reorder or the first check after a read rebuilds the whole index (2,000 alternating removals from the head and additions on 10,000 defined names take 2.3 s).
- `cloneSheet` giving the copy its own sheet-scoped copies of the source sheet's sheet-scoped names and of the workbook names that refer to the sheet, as Excel's own sheet copy does; today it copies no defined names. A bare table name used as an operand (`ROWS(Table1)`) keeps naming the source table, as in Excel's copy.
- Internal, no public API change: one sheet-name predicate shared by `validateSheetName` and the quoting of `renameSheet` spans, and that quoting rule shared with `src/formula/serializer.ts`.
- `renameSheet` rewriting every formula-bearing part that names the sheet (cell formulas, chart series, pivot cache sources, data-validation and conditional-format formulas, hyperlink locations), and the tokenizer reading `Sheet1!myName` and `[1]!srcName`, so defined-name values that hold them are rewritten too; today `renameSheet` rewrites defined names, print areas and print titles only.
- `moveSheet` and `removeSheet` rewriting a 3-D span in a defined name as Excel's own move does (`Sheet1:Sheet3!$A$2` becomes `Sheet1:Sheet2!$A$2` when `Sheet3` moves to the front); today the span keeps its text, so it covers other sheets or names a removed one.
- A calculated column or custom totals formula given in the `@` this-row shorthand (`[@Qty]*[@Price]`) rewritten into the form Excel writes in a file (`Sales[[#This Row],[Qty]]*Sales[[#This Row],[Price]]`), in `addTable` and in a cell formula inside a table; today `addTable` refuses the shorthand and a cell formula keeps it, which Excel reads as `#REF!`.
- A shared-formula group re-homed onto another master when its master is edited or cleared, so its children keep their formulas; today they follow the master. With it, a child that precedes its master in row order, and two masters on one `sharedIndex`, are written in a form Excel opens; today both are written as given, Excel refuses the file, and `validate()` does not report them.
- `validate()` reporting a shared-formula master whose `formulaRef` is blank or not a range, or whose formula text is only U+200B, U+FEFF or `=` signs and spaces; today such a master is written as given.

## 1.7.0: Drawings and images, first-class

- Typed images (PNG, JPEG, GIF, BMP and SVG) with one-cell, two-cell and absolute anchors, stored on the sheet data and numbered when the workbook is written, so moving or removing sheets keeps every image on its sheet.
- New images merge into a drawing that the file already has.
- Shapes, text boxes, checkboxes and form controls.
- `generateDrawingXml` and `generateDrawingRels` move to generic drawing helpers outside the barcode module; the old exports stay, deprecated, until 2.0.
- One anchor type shared by images and charts in place of `ImageAnchor` and `ChartAnchorData`.
- Pictures added with `addImage` named `Picture N` in the drawing; today they are named `Barcode N`.
- Preserved sheet relationships keyed to the sheet instead of its position, so `removeSheet`, `moveSheet` and `cloneSheet` with an insert index keep every preserved drawing, form control and header or footer picture on its sheet.
- The sheet's `<controls>` element kept and written for form controls read from a file, so Excel no longer also loads their DrawingML fallback shapes as hidden shapes.
- A note drawing read from a file gets its `o:idmap` extended when added note shapes pass its id block.

## 1.8.0: Charts complete

- 3-D bar, line, pie and area charts, surface charts, and pie-of-pie and bar-of-pie charts, modelled instead of kept verbatim.
- Bundled palettes for chart styles 9 to 48; `getChartStylePalette` has palettes for styles 1 to 8.
- Chartsheets.
- More than one secondary chart type in a combo chart.
- A `ChartBuilder` method that draws the secondary chart on its own axes; today `secondaryChart()` shares the primary axes unless `secondaryCatAxis` and `secondaryValAxis` are set on the built model.
- High-low lines and up/down bars.
- Chart fixtures authored in Excel, read, written and read again.
- An SVG chart preview.

## 1.9.0: Pivot and BI parity

- Pivot tables model show-values-as, field subtotals and more data-field options, and keep what a round trip drops today: `showAll`, `sortType` and `numFmtId` on fields, the row and column item caches, the location's `rowPageCount` and `colPageCount`, and the style info.
- The writer computes a pivot table's layout and cell values, so a table built in code shows them without a refresh in Excel.
- A cache field's `sharedItemsAttrs` derived from its shared items, so number and date items no longer need them set by hand.
- Writer-side pivot validation; for example, a field placed in two areas is rejected.
- Slicers and timelines written in the MS-XLSX x14 and x15 form that Excel loads, with the workbook and worksheet extension links, and checked against Excel-authored files. Today both are a custom dialect that is not Excel-verified. Files in that dialect stay readable.
- Pivot refresh on load as a model field, read and written; today the writer sets `refreshOnLoad` only on a cache whose table it completes, and a read drops it.
- `removePivotCache`, `removeSlicerCache` and `removeTimelineCache`.
- The `id` of a kept pivot cache relationship limited in length on write, so the written `.rels` part stays within the reader's default `maxCompressionRatio`; today a very long, highly repetitive id gives a file that fails to read with `SECURITY`.

## 1.10.0: DX, schemas, templates and formatting

- Schema-driven typed row parsing compatible with Standard Schema, with per-cell errors.
- Template fill over an existing workbook.
- Number-format rendering on par with SSF: `*` fill to the column width, text sections, a defined result for negative date serials, and a way to remove formats registered with `loadFormat`.
- Autofit column widths.
- Synchronous read and write once `initWasmSync` has run.
- Typed error kinds with source chains.
- PDF417 (ISO/IEC 15438) and DataMatrix (ISO/IEC 16022, ECC 200) encoders that conform to their standards, replacing today's experimental ones, with CI decoding their output.
- Top-level `writeBuffer` and `writeFile`, a `dateSystem` setter, `removeSparklineGroup` and `removeThreadedComment`.
- `modern-xlsx convert` infers the output format from the output file's extension (`out.csv` writes CSV) when `--format` is not given; today it writes JSON unless `--format csv` is passed.
- Workbook protection set from a plain password, as sheet protection already is through `plainPassword`.
- `workbookPassword` and `revisionsPassword` checked like a sheet protection `password`: a value that is not a legacy hash fails the write, and `validate()` reports it.
- A placeholder note read from a file rewritten with its thread's replies when a reply is added, and a `validate()` warning for a thread on a cell that holds an ordinary note, which Excel shows instead of the thread.
- A `tc=` placeholder note whose thread is no longer in the model dropped on write, so a thread added on that cell gets its own placeholder.
- Every `<sheetView>` of a sheet kept through a read and a write, so a workbook saved with several windows on a sheet keeps each window's selection and scroll position; today only the first sheet view is kept.
- A `validate()` error for a `splitPane` with neither `horizontal` nor `vertical` set, which is written as a split pane without sizes that Excel refuses to open.
- `addTable` refusing, and `validate()` reporting, a table name with a control character other than tab, line feed and carriage return; today it is accepted and the writer drops the character.
- `addTable` writing each column name into an empty header cell of the table; today a table over empty header cells gives a file Excel refuses to open.
- `cloneSheet` giving a copied thread ids in the source thread's id order, so replies that share one `dT` keep their order.

## 1.11.0: Legacy formats in, XLS, XLSB and Strict OOXML

- XLS (BIFF8) and XLSB read in Rust, on the shared compound-file layer.
- Encrypted XLS (RC4 CryptoAPI) read, with a precise error where a scheme is unsupported.
- Strict OOXML read: the Strict namespaces and relationship types.

## 1.12.0: Render and runtime breadth, 2.0 runway

- SVG, HTML and PDF sheet rendering: cells, styles, merges, and charts from 1.8.0.
- Browser tests in Firefox and WebKit as well as Chromium.
- A blocking CI leg on the newest Node release line; today CI blocks on Node 24 and 26 and runs a Node nightly leg as advisory only.
- A size-budget redesign: whether lite becomes the default entry, optional Cargo features such as one for charts, and WASM `build-std` experiments.
- Deprecation warnings and codemods for every 2.0 removal.

## 2.0.0: Clean bridge and API

- The deprecated 1.x fields go: `dynamicArray` (use `cellMetadata`), and `formulaDtr1` and `formulaDtr2` (use `formulaDel1` and `formulaDel2`). Formulas are then read through the `cm`, `dtr`, `del1` and `del2` attributes only, and the `dtr1` and `dtr2` attributes that 1.2.x wrote are no longer accepted.
- `TokenType` loses `'range'`, which the tokenizer never emits.
- `RewriteAction` uses one base for rows and columns; today `start` is 1-based for rows and 0-based for columns.
- `ColorData` everywhere; the rgb-only colour strings beside the `*Spec` fields are removed.
- A date cell type.
- Preserved parts travel as bytes by default.
- An opaque Rust error type with `kind()` and `source()`.
- Sheet state typed as an enum in the workbook model.
- Lowercase error messages.
- `patternType` accepts only `PatternType` values.
- `XlsxWorker.readBuffer` returns a `Workbook` instead of `WorkbookData`.
- Readonly arrays across the public data types (`WorksheetData`, `WorkbookData`, `StylesData`, the formula token and syntax-tree types, and the barcode matrix and render options), and a readonly `toJSON()` result.
- Lookups return `null` instead of `undefined`: `getSheet`, `getSheetByIndex`, `getNamedRange`, `getTable` and `getBuiltinFormat`.
- `validateChartData` throws `ModernXlsxError` with a code instead of `RangeError`.
- Every `remove*` method takes the same kind of key.
- Node 24 support is dropped once Node 24 reaches end of life.
- The minimum Rust version rises so integer formatting comes from the standard library instead of the `itoa` crate.
- Lite becomes the default entry, and encryption an add-on import.

### Conditions for 2.0

2.0.0 ships only when all of these hold:

- No critical or high-severity defect is open.
- The fixture corpus of files authored in Excel and LibreOffice passes in CI.
- The encryption interop corpus passes against msoffcrypto-tool and Excel-authored files.
- The runtime matrix is green.
- A 2.0 migration guide exists, with codemods for every removed field and alias.

1.12.0 is the last 1.x feature release. After 2.0.0 ships, 1.12.x receives security and fidelity fixes only, for six months.

## Waiting on input

- An Excel-authored workbook encrypted with Agile AES-128 and SHA-1 for the encryption interop corpus. Until one is available, that combination is covered by round trips of the library's own output.

## Not planned

- **wasm-snip.** Release builds already run `wasm-opt -Oz`, and CI enforces the gzip budget.
- **A console logging backend for WASM.** Errors reach JavaScript as coded `ModernXlsxError` messages, and debug builds already send panics to the console.
- **A separate publish workflow.** Publishing is the `publish` job of the main CI workflow.
- **CI on Node 20 and 22.** The package requires Node 24 or newer.
- **A `modern-xlsx/core` entry.** The reduced build ships as `modern-xlsx/lite`.
- **A 1.0 migration guide.** 1.0.0 has shipped; the migration guide that remains planned is the one for 2.0.
- **Reordering `loadFormat(formatCode, id)`.** Swapping the arguments would break every caller for no gain.
- **Merging `ensureReady` into `initWasm`.** `ensureReady` is the idempotent form of `initWasm`; both stay.
- **Exporting `wasmVersion` from the entries.** The WASM crate has its own version, independent of the npm package version, which `VERSION` reports.
- **`@since` tags.** The changelog records when each API arrived.
- **Getters for external links, custom XML parts and data tables.** The data is already on the model: `WorkbookData.externalReferences` lists the external link parts, the data-table flags are on `CellData`, and custom XML parts round-trip in `preservedEntries`.
