# API Surface Audit for 1.0.0 Stability

**Date:** 2026-03-06
**Scope:** All public exports from `packages/modern-xlsx/src/index.ts`
**Goal:** Identify naming, consistency, and type-safety issues before locking the 1.0 API.

---

## Table of Contents

1. [Workbook Class](#1-workbook-class)
2. [Worksheet Class](#2-worksheet-class)
3. [Cell Class](#3-cell-class)
4. [Types (types.ts)](#4-types-typests)
5. [Utility Functions (utils.ts)](#5-utility-functions-utilsts)
6. [Cell Reference Utilities (cell-ref.ts)](#6-cell-reference-utilities-cell-refts)
7. [Date Utilities (dates.ts)](#7-date-utilities-datests)
8. [Format Cell (format-cell.ts)](#8-format-cell-format-cellts)
9. [ChartBuilder (chart-builder.ts)](#9-chartbuilder-chart-builderts)
10. [StyleBuilder (style-builder.ts)](#10-stylebuilder-style-builderts)
11. [HeaderFooterBuilder (header-footer.ts)](#11-headerfooterbuilder-header-footerts)
12. [RichTextBuilder (rich-text.ts)](#12-richtextbuilder-rich-textts)
13. [StreamingXlsxWriter (streaming-writer.ts)](#13-streamingxlsxwriter-streaming-writerts)
14. [Worker API (worker-api.ts)](#14-worker-api-worker-apits)
15. [Table Layout Engine (table.ts)](#15-table-layout-engine-tablets)
16. [Chart Styles (chart-styles.ts)](#16-chart-styles-chart-stylests)
17. [Table Styles (table-styles.ts)](#17-table-styles-table-stylests)
18. [Barcode Module (barcode/)](#18-barcode-module-barcode)
19. [Formula Engine (formula/)](#19-formula-engine-formula)
20. [Errors (errors.ts)](#20-errors-errorsts)
21. [Top-Level Functions (index.ts)](#21-top-level-functions-indexts)
22. [WASM Loader (wasm-loader.ts)](#22-wasm-loader-wasm-loaderts)
23. [Summary of Findings](#23-summary-of-findings)

---

## 1. Workbook Class

**File:** `workbook.ts`, lines 100-665

### Naming

All method and property names are camelCase -- consistent.

### Return Type Consistency (null vs undefined)

| Member | Returns | Issue |
|---|---|---|
| `getSheet(name)` | `Worksheet \| undefined` | **INCONSISTENT** -- getters should return `T \| null` |
| `getSheetByIndex(index)` | `Worksheet \| undefined` | **INCONSISTENT** -- getters should return `T \| null` |
| `getNamedRange(name)` | `DefinedNameData \| undefined` | **INCONSISTENT** -- getters should return `T \| null` |
| `docProperties` | `DocPropertiesData \| null` | OK |
| `themeColors` | `ThemeColorsData \| null` | OK |
| `protection` | `WorkbookProtectionData \| null` | OK |
| `getPrintTitles(sheetIndex)` | `string \| null` | OK |
| `getPrintArea(sheetIndex)` | `string \| null` | OK |

**Recommendation:** Change `getSheet`, `getSheetByIndex`, and `getNamedRange` to return `T | null` instead of `T | undefined`.

### Readonly Arrays

| Member | Returns | Issue |
|---|---|---|
| `sheetNames` | `readonly string[]` | OK |
| `namedRanges` | `readonly DefinedNameData[]` | OK |
| `calcChain` | `readonly CalcChainEntryData[]` | OK |
| `workbookViews` (getter) | `readonly WorkbookViewData[]` | OK |
| `pivotCaches` | `readonly PivotCacheDefinitionData[]` | OK |
| `pivotCacheRecords` | `readonly PivotCacheRecordsData[]` | OK |
| `slicerCaches` | `readonly SlicerCacheData[]` | OK |
| `timelineCaches` | `readonly TimelineCacheData[]` | OK |

All array getters return `readonly T[]` -- correct.

### Parameter Order

All methods follow "target first, options last" except:

| Method | Signature | Issue |
|---|---|---|
| `addImage(sheetName, anchor, imageBytes, format)` | target, config, data, option | OK (format is optional trailing param) |
| `addBarcode(sheetName, anchor, value, options)` | target, config, data, options | OK |
| `cloneSheet(sourceIndex, newName, insertIndex?)` | source, name, position | OK |

Parameter order is consistent.

### Method Naming Pairs

| Add | Remove | Issue |
|---|---|---|
| `addSheet` | `removeSheet` | OK |
| `addNamedRange` | `removeNamedRange` | OK |
| `addPivotCache` | -- | **MISSING** `removePivotCache` |
| `addPivotCacheRecords` | -- | **MISSING** `removePivotCacheRecords` |
| `addSlicerCache` | -- | **MISSING** `removeSlicerCache` |
| `addTimelineCache` | -- | **MISSING** `removeTimelineCache` |

**Recommendation:** Add `removePivotCache`, `removePivotCacheRecords`, `removeSlicerCache`, and `removeTimelineCache` methods for symmetry, or document why remove is not needed.

### Get/Set Pairs

| Getter | Setter | Issue |
|---|---|---|
| `docProperties` (get) | `docProperties` (set) | OK |
| `themeColors` (get) | -- | **NO SETTER** -- intentional (read from theme.xml)? Should be documented |
| `workbookViews` (get) | `workbookViews` (set) | OK |
| `protection` (get) | `protection` (set) | OK |
| `styles` (get) | -- | **NO SETTER** -- OK (mutation via StyleBuilder) |
| `dateSystem` (get) | -- | **NO SETTER** -- should there be one? Changing date system is a valid operation |

**Recommendation:** Consider adding a `dateSystem` setter or a `setDateSystem()` method. If `themeColors` is intentionally read-only, add a JSDoc note.

### Overload Opportunities

| Method | Current | Potential Overload |
|---|---|---|
| `removeSheet(nameOrIndex)` | `string \| number` | Already accepts both -- good |
| `renameSheet(nameOrIndex, newName)` | `string \| number` | Already accepts both -- good |
| `hideSheet(nameOrIndex)` | `string \| number` | Already accepts both -- good |
| `unhideSheet(nameOrIndex)` | `string \| number` | Already accepts both -- good |

Union parameter approach is fine; TypeScript overloads would add no value here.

### Other Issues

- **`toJSON()` return type:** Returns mutable `WorkbookData` -- should return `Readonly<WorkbookData>` or `DeepReadonly<WorkbookData>` to prevent accidental mutation of internal state.
- **`repair()` return type:** Returns `{ workbook: Workbook; report: ValidationReport; repairCount: number }` as an anonymous object. Consider defining a named `RepairOutput` interface for discoverability.

---

## 2. Worksheet Class

**File:** `workbook.ts`, lines 672-1455

### Return Type Consistency (null vs undefined)

| Member | Returns | Issue |
|---|---|---|
| `dimension` | `string \| null` | OK |
| `autoFilter` | `AutoFilterData \| null` | OK |
| `pageSetup` | `PageSetupData \| null` | OK |
| `sheetProtection` | `SheetProtectionData \| null` | OK |
| `frozenPane` | `FrozenPane \| null` | OK |
| `splitPane` | `SplitPaneData \| null` | OK |
| `view` | `SheetViewData \| null` | OK |
| `tabColor` | `string \| null` | OK |
| `usedRange` | `string \| null` | OK |
| `headerFooter` | `HeaderFooterData \| null` | OK |
| `outlineProperties` | `OutlinePropertiesData \| null` | OK |
| `pageMargins` | `PageMarginsData \| null` | OK |
| `getTable(displayName)` | `TableDefinitionData \| undefined` | **INCONSISTENT** -- should be `T \| null` |

**Recommendation:** Change `getTable` to return `T | null`.

### Readonly Arrays

| Member | Returns | Issue |
|---|---|---|
| `rows` | `readonly RowData[]` | OK |
| `columns` (getter) | `readonly ColumnInfo[]` | OK |
| `mergeCells` | `readonly string[]` | OK |
| `hyperlinks` | `readonly HyperlinkData[]` | OK |
| `validations` | `readonly DataValidationData[]` | OK |
| `comments` | `readonly CommentData[]` | OK |
| `threadedComments` | `readonly ThreadedCommentData[]` | OK |
| `tables` | `readonly TableDefinitionData[]` | OK |
| `charts` | `readonly WorksheetChartData[]` | OK |
| `pivotTables` | `readonly PivotTableData[]` | OK |
| `slicers` | `readonly SlicerData[]` | OK |
| `timelines` | `readonly TimelineData[]` | OK |
| `sparklineGroups` | `readonly SparklineGroupData[]` | OK |
| `paneSelections` (getter) | `readonly PaneSelectionData[]` | OK |

All correct.

### Method Naming Pairs

| Add | Remove | Issue |
|---|---|---|
| `addMergeCell` | `removeMergeCell` | OK |
| `addHyperlink` | `removeHyperlink` | OK |
| `addValidation` | `removeValidation` | OK |
| `addComment` | `removeComment` | OK |
| `addTable` | `removeTable` | OK |
| `addChart` | `removeChart` | OK |
| `addChartData` | `removeChart` (by index) | OK (two add variants, one remove) |
| `addPivotTable` | `removePivotTable` | OK |
| `addSlicer` | `removeSlicer` | OK |
| `addTimeline` | `removeTimeline` | OK |
| `addSparklineGroup` | `clearSparklineGroups` | **ASYMMETRIC** -- `add` singular vs `clear` plural. No `removeSparklineGroup(index)` |
| `addThreadedComment` | -- | **MISSING** `removeThreadedComment` |
| `groupRows` | `ungroupRows` | OK (verb pair, not add/remove) |
| `groupColumns` | `ungroupColumns` | OK |
| `collapseRows` | `expandRows` | OK |

**Recommendation:**
- Add `removeSparklineGroup(index): boolean` for consistency with other remove methods.
- Add `removeThreadedComment(id: string): boolean`.

### Parameter Inconsistency: Index-Based Remove

Some `remove*` methods accept a name/key, others accept a numeric index:

| Method | Parameter | Convention |
|---|---|---|
| `removeMergeCell(range)` | by value (string) | OK |
| `removeHyperlink(cellRef)` | by key (string) | OK |
| `removeComment(cellRef)` | by key (string) | OK |
| `removeValidation(ref)` | by key (string) | OK |
| `removeTable(displayName)` | by key (string) | OK |
| `removeChart(index)` | by numeric index | **DIFFERENT** |
| `removePivotTable(index)` | by numeric index | **DIFFERENT** |
| `removeSlicer(index)` | by numeric index | **DIFFERENT** |
| `removeTimeline(index)` | by numeric index | **DIFFERENT** |

**Recommendation:** Charts, pivot tables, slicers, and timelines use index-based removal while other collections use key-based removal. This is inconsistent. Consider adding name-based overloads for `removeChart`, `removePivotTable`, `removeSlicer`, and `removeTimeline`, or accept `string | number` like `removeSheet`.

### Setter Type Widening

| Setter | Accepts | Issue |
|---|---|---|
| `autoFilter` (set) | `AutoFilterData \| string \| null` | Widens to accept string -- good ergonomics but setter type is wider than getter. Consider an overloaded signature or separate `setAutoFilter(filter)` method. |

### Other Issues

- **`addThreadedComment` parameter order**: `(cell, text, author)` -- text before author is unusual. Most APIs put `author` before `text` since the author is a property of the comment, not content. However, this matches a "what, then who" pattern that is also defensible.
- **`replyToComment` parameter order**: `(commentId, text, author)` -- consistent with `addThreadedComment` so this is internally consistent.
- **`setColumnWidth(col, width)`**: `col` is documented as "1-based" but the convention in `cell-ref.ts` is 0-based. This is a potential confusion point. Should be documented prominently.
- **`setRowHeight(rowIndex, height)` and `setRowHidden(rowIndex, hidden)`**: `rowIndex` is 1-based (matches Excel row numbers). Consistent with internal `RowData.index`.
- **`groupRows(startRow, endRow, level)`**: 1-based. Consistent.
- **`groupColumns(startCol, endCol, level)`**: 1-based. Consistent with `setColumnWidth`.
- **Constructor visibility**: `Worksheet` constructor is public but takes internal types (`SheetData`, `StylesData`, `WorkbookData`). For 1.0, consider making it `@internal` or narrowing.

---

## 3. Cell Class

**File:** `workbook.ts`, lines 1462-1576

### Return Type Consistency

| Member | Returns | Issue |
|---|---|---|
| `value` (getter) | `string \| number \| boolean \| null` | OK (null for empty) |
| `formula` (getter) | `string \| null` | OK |
| `styleIndex` (getter) | `number \| null` | OK |
| `numberFormat` (getter) | `string \| null` | OK |
| `dateValue` (getter) | `Date \| null` | OK |
| `reference` (getter) | `string` | OK (always set) |
| `type` (getter) | `CellType` | OK |

All consistent with `T | null` for optional values.

### Other Issues

- **`type` vs `cellType`**: The Cell getter is named `type` but the underlying `CellData.cellType` uses `cellType`. The public API name `type` is cleaner and correct. No issue.
- **Constructor visibility**: `Cell` constructor is public but takes internal types. Same recommendation as Worksheet -- mark `@internal` or restrict.

---

## 4. Types (types.ts)

**File:** `types.ts`, 1091 lines

### Naming Conventions

All interfaces use PascalCase with `Data` suffix -- consistent pattern.
All type aliases use PascalCase -- correct.
All string literal union types use camelCase values -- correct.

### null vs undefined Consistency in Interfaces

The codebase uses a **mixed strategy**:
- `T | null` for explicitly nullable fields (e.g., `value: string | null`)
- `T?` (optional) with `| null` for fields that may be absent AND null (e.g., `formulaType?: string | null`)
- `T?` alone for truly optional fields with no null semantics (e.g., `outlineLevel?: number`)

This is actually reasonable -- the pattern maps to serde behavior where:
- Required fields with `null` -> always present, may be null
- Optional fields `?` -> may be missing from JSON
- Optional + null `? | null` -> may be missing or present-as-null

However, there are inconsistencies:

| Interface | Field | Current | Issue |
|---|---|---|---|
| `RowData` | `cells` | `readonly CellData[]` | Has `readonly` -- good |
| `WorksheetData` | `rows` | `RowData[]` | **MISSING `readonly`** |
| `WorksheetData` | `mergeCells` | `string[]` | **MISSING `readonly`** |
| `WorksheetData` | `columns` | `ColumnInfo[]` | **MISSING `readonly`** |
| `WorksheetData` | `dataValidations?` | `DataValidationData[]` | **MISSING `readonly`** |
| `WorksheetData` | `conditionalFormatting?` | `ConditionalFormattingData[]` | **MISSING `readonly`** |
| `WorksheetData` | `hyperlinks?` | `HyperlinkData[]` | **MISSING `readonly`** |
| `WorksheetData` | `comments?` | `CommentData[]` | **MISSING `readonly`** |
| `WorksheetData` | `tables?` | `TableDefinitionData[]` | **MISSING `readonly`** |
| `WorksheetData` | `sparklineGroups?` | `SparklineGroupData[]` | **MISSING `readonly`** |
| `WorksheetData` | `charts?` | `WorksheetChartData[]` | **MISSING `readonly`** |
| `WorksheetData` | `pivotTables?` | `PivotTableData[]` | **MISSING `readonly`** |
| `WorksheetData` | `threadedComments?` | `ThreadedCommentData[]` | **MISSING `readonly`** |
| `WorksheetData` | `slicers?` | `SlicerData[]` | **MISSING `readonly`** |
| `WorksheetData` | `timelines?` | `TimelineData[]` | **MISSING `readonly`** |
| `WorkbookData` | `sheets` | `SheetData[]` | **MISSING `readonly`** |
| `StylesData` | `numFmts` | `NumFmt[]` | **MISSING `readonly`** |
| `StylesData` | `fonts` | `FontData[]` | **MISSING `readonly`** |
| `StylesData` | `fills` | `FillData[]` | **MISSING `readonly`** |
| `StylesData` | `borders` | `BorderData[]` | **MISSING `readonly`** |
| `StylesData` | `cellXfs` | `CellXfData[]` | **MISSING `readonly`** |
| `StylesData` | `dxfs?` | `DxfStyleData[]` | **MISSING `readonly`** |
| `StylesData` | `cellStyles?` | `CellStyleData[]` | **MISSING `readonly`** |
| `PivotFieldData` | `items` | `PivotItem[]` | **MISSING `readonly`** |
| `PivotFieldData` | `subtotals` | `SubtotalFunction[]` | **MISSING `readonly`** |
| `PivotTableData` | `pivotFields` | `PivotFieldData[]` | **MISSING `readonly`** |
| `PivotTableData` | `rowFields` | `PivotFieldRef[]` | **MISSING `readonly`** |
| `PivotTableData` | `colFields` | `PivotFieldRef[]` | **MISSING `readonly`** |
| `PivotTableData` | `dataFields` | `PivotDataFieldData[]` | **MISSING `readonly`** |
| `PivotTableData` | `pageFields` | `PivotPageFieldData[]` | **MISSING `readonly`** |
| `PivotCacheRecordsData` | `records` | `CacheValue[][]` | **MISSING `readonly`** |
| `WorkbookData` | `definedNames?` | `DefinedNameData[]` | **MISSING `readonly`** |
| `WorkbookData` | `workbookViews?` | `WorkbookViewData[]` | **MISSING `readonly`** |
| `WorkbookData` | `persons?` | `PersonData[]` | **MISSING `readonly`** |
| `WorkbookData` | `pivotCaches?` | `PivotCacheDefinitionData[]` | **MISSING `readonly`** |
| `WorkbookData` | `pivotCacheRecords?` | `PivotCacheRecordsData[]` | **MISSING `readonly`** |
| `WorkbookData` | `slicerCaches?` | `SlicerCacheData[]` | **MISSING `readonly`** |
| `WorkbookData` | `timelineCaches?` | `TimelineCacheData[]` | **MISSING `readonly`** |

**Recommendation:** This is the single largest issue in the API. Almost all array fields in `WorkbookData`, `WorksheetData`, `StylesData`, and `PivotTableData` are missing `readonly`. Some arrays DO have `readonly` (e.g., `FilterColumnData.filters`, `ConditionalFormattingData.rules`, `SharedStringsData.strings`, `ChartDataModel.series`, `ValidationReport.issues`), making the omission in peer interfaces a clear inconsistency.

**Caveat:** Adding `readonly` to `WorkbookData`/`WorksheetData`/`StylesData` arrays is a breaking change if user code pushes directly to these arrays. Since the Workbook/Worksheet classes provide `add*`/`remove*` methods, the `Data` interfaces should be `readonly` to enforce using those methods.

### BarcodeMatrix Inconsistency

`BarcodeMatrix.modules` is typed as `boolean[][]` (mutable 2D array). Should be `readonly (readonly boolean[])[]` for a read-only output type.

### CacheFieldData

`CacheFieldData.sharedItems` is `CacheValue[]` (optional, mutable). Should be `readonly CacheValue[]`.

### SlicerCacheData

`SlicerCacheData.items` is `SlicerItem[]`. Should be `readonly SlicerItem[]`.

---

## 5. Utility Functions (utils.ts)

**File:** `utils.ts`, 630 lines

### Naming

All function names are camelCase -- correct.
All option interfaces are PascalCase -- correct.

### Parameter Order

| Function | Signature | Issue |
|---|---|---|
| `sheetToJson(ws, opts?)` | target, options | OK |
| `sheetToCsv(ws, opts?)` | target, options | OK |
| `sheetToTxt(ws, opts?)` | target, options | OK |
| `sheetToHtml(ws, opts?)` | target, options | OK |
| `sheetToFormulae(ws)` | target | OK |
| `jsonToSheet(data, opts?)` | data, options | OK |
| `aoaToSheet(data, opts?)` | data, options | OK |
| `sheetAddAoa(ws, data, opts?)` | target, data, options | OK |
| `sheetAddJson(ws, data, opts?)` | target, data, options | OK |

All consistent: target first, data next, options last.

### Return Types

| Function | Returns | Issue |
|---|---|---|
| `sheetToJson` | `T[]` | OK |
| `jsonToSheet` | `Worksheet` | OK |
| `aoaToSheet` | `Worksheet` | OK |
| `sheetToCsv` | `string` | OK |
| `sheetToTxt` | `string` | OK |
| `sheetToHtml` | `string` | OK |
| `sheetToFormulae` | `string[]` | OK |
| `sheetAddAoa` | `void` | OK |
| `sheetAddJson` | `void` | OK |

### Issues

- **`sheetToJson` generic constraint**: `<T extends Record<string, unknown>>` -- good type safety.
- **`sheetAddAoa` / `sheetAddJson` return void**: Correct for mutation methods.
- **No `sheetToAoa` function**: There is `sheetToJson` and `sheetToCsv` but no `sheetToAoa` (sheet to array-of-arrays). This is a common conversion in SheetJS. Consider adding for migration completeness.

---

## 6. Cell Reference Utilities (cell-ref.ts)

**File:** `cell-ref.ts`, 146 lines

### Naming

All functions camelCase, all types PascalCase -- correct.

### Return Types

| Function | Returns | Issue |
|---|---|---|
| `columnToLetter(col)` | `string` | OK |
| `letterToColumn(letter)` | `number` | OK |
| `encodeCellRef(row, col)` | `string` | OK |
| `decodeCellRef(ref)` | `CellAddress` | OK (throws on invalid) |
| `encodeRange(start, end)` | `string` | OK |
| `decodeRange(range)` | `CellRange` | OK (throws on invalid) |
| `encodeRow(row)` | `string` | OK |
| `decodeRow(rowStr)` | `number` | OK (throws on invalid) |
| `splitCellRef(ref)` | `SplitCellRef` | OK (throws on invalid) |

### Issues

- **Error handling**: `decodeCellRef`, `decodeRange`, `decodeRow`, and `splitCellRef` throw plain `Error` instead of `ModernXlsxError` with `INVALID_CELL_REF` code. This is inconsistent with the project's error handling strategy.
- **`CellAddress` readonly fields**: Both `row` and `col` are `readonly` -- good.
- **`CellRange` readonly fields**: `start` and `end` are `readonly` -- good.
- **`SplitCellRef` readonly fields**: All fields are `readonly` -- good.

**Recommendation:** Change thrown errors in cell-ref.ts to use `ModernXlsxError` with `INVALID_CELL_REF` code.

---

## 7. Date Utilities (dates.ts)

**File:** `dates.ts`, 203 lines

### Naming

All functions camelCase -- correct.

### Parameter Order

| Function | Signature | Issue |
|---|---|---|
| `dateToSerial(date, system?)` | data, option | OK |
| `serialToDate(serial, system?)` | data, option | OK |
| `isDateFormatId(numFmtId)` | data | OK |
| `isDateFormatCode(formatCode)` | data | OK |
| `isTemporalLike(value)` | data | OK (type guard) |

### Issues

- **`isTemporalLike` export style**: Defined as a function declaration inside the module, then re-exported via `export { isTemporalLike }` on line 128. This is fine but unusual -- most other exports use `export function` directly.
- **`TemporalLike` interface not exported**: The duck-typed interface `TemporalLike` used by `isTemporalLike` and `dateToSerial` is not exported. Users who want to pass a `TemporalLike` object cannot import the type. **Recommendation:** Export `TemporalLike` as a public type.
- **`dateToSerial` input type**: Accepts `Date | TemporalLike` but `TemporalLike` is not exported, making the type opaque.

---

## 8. Format Cell (format-cell.ts)

**File:** `format-cell.ts`, 601 lines

### Naming

All functions camelCase -- correct.

### Return Types

| Function | Returns | Issue |
|---|---|---|
| `formatCell(value, format, opts?)` | `string` | OK |
| `formatCellRich(value, format, opts?)` | `FormatCellResult` | OK |
| `getBuiltinFormat(id)` | `string \| undefined` | **INCONSISTENT** -- should return `string \| null` |
| `loadFormat(formatCode, id)` | `void` | OK |
| `loadFormatTable(table)` | `void` | OK |

**Recommendation:** Change `getBuiltinFormat` to return `string | null`.

### Parameter Order

| Function | Signature | Issue |
|---|---|---|
| `formatCell(value, format, opts?)` | data, format, options | OK |
| `formatCellRich(value, format, opts?)` | data, format, options | OK |
| `getBuiltinFormat(id)` | key | OK |
| `loadFormat(formatCode, id)` | data, key | **REVERSED** -- convention is key/id first, then data. `loadFormat(id, formatCode)` would be more natural |
| `loadFormatTable(table)` | data | OK |

**Recommendation:** Consider swapping `loadFormat` parameter order to `(id, formatCode)` for consistency with map-like APIs (key, value).

### FormatCellResult Interface

```typescript
interface FormatCellResult {
  text: string;
  color?: string;  // optional, not T | null
}
```

The `color` field uses `?` (truly optional) rather than `string | null`. This is correct -- the color is only present when the format code specifies one.

---

## 9. ChartBuilder (chart-builder.ts)

**File:** `chart-builder.ts`, 399 lines

### Naming

All methods camelCase, class PascalCase -- correct.

### Fluent API

All setter methods return `this` -- correct for a fluent builder.

### Issues

- **`build()` return type**: Returns `WorksheetChartData` -- good, named type.
- **`title()` method name clashes with property concept**: The `title()` method is both a setter and could be confused with a getter. For builders this is standard practice. No change needed.
- **`style(id)` method name**: Very generic. Could be confused with CSS styling. However, in context of ChartBuilder it is clear. The JSDoc clarifies "chart style ID (1-48)".
- **`anchor()` parameter types**: Uses anonymous object literals `{ col: number; row: number; colOff?: number; rowOff?: number }`. Consider extracting to a named interface (e.g., `AnchorPoint`) for reuse and documentation.
- **`AddSeriesOptions.dataLabels`**: Typed as `Partial<DataLabelsData>`. This is fine.
- **`AxisOptions.title`**: Accepts `string | { text: string; ... }` -- good overloaded input type.

---

## 10. StyleBuilder (style-builder.ts)

**File:** `style-builder.ts`, 172 lines

### Naming

All methods camelCase, class PascalCase -- correct.

### Fluent API

All setter methods return `this` -- correct.

### Issues

- **`build(styles)` takes a `StylesData` parameter**: The builder mutates the passed `StylesData` object. This is a side-effecting `build()` method, which is unusual. Most builders are pure and return data. Consider renaming to `register(styles)` to communicate the mutation, or return both the index and the mutated styles.
- **`fill()` method parameter**: Uses `pattern` key name instead of `patternType` (which is the field name in `FillData`). This is intentional shorthand for ergonomics but creates a naming mismatch between builder input and data output.
- **`protection()` method name clashes with `Worksheet.sheetProtection`**: Different concepts (cell-level vs sheet-level protection). The naming is clear in context.

---

## 11. HeaderFooterBuilder (header-footer.ts)

**File:** `header-footer.ts`, 120 lines

### Naming

All methods and static methods camelCase -- correct.

### Issues

- **Static methods for codes**: `pageNumber()`, `totalPages()`, `date()`, `time()`, etc. These are essentially constants that return format code strings. Could be `static readonly` constants instead:
  ```typescript
  static readonly PAGE_NUMBER = '&P';
  ```
  However, the method approach allows future parameterization. Current approach is acceptable.

- **`build()` returns `string`**: Correct for a format string builder.

---

## 12. RichTextBuilder (rich-text.ts)

**File:** `rich-text.ts`, 75 lines

### Naming

All methods camelCase, class PascalCase -- correct.

### Issues

- **`build()` return type**: Returns `readonly RichTextRun[]` -- correct, uses readonly.
- **`plainText()` method**: Utility method, not part of builder chain. Returns `string`. OK.
- **`colored()` method name**: Unconventional. `color()` or `withColor()` would be more standard. However, `colored` reads well as "add colored text".

---

## 13. StreamingXlsxWriter (streaming-writer.ts)

**File:** `streaming-writer.ts`, 111 lines

### Naming

Class PascalCase, methods camelCase -- correct.

### Issues

- **Static factory**: `StreamingXlsxWriter.create()` -- good pattern for WASM-backed objects.
- **`writeRow(cells)` parameter type**: `StreamingCellInput[]` -- not `readonly StreamingCellInput[]`. Since the array is serialized to JSON immediately, `readonly` would be more correct.
- **`finish()` return type**: `Uint8Array` -- correct.
- **No `addSheet` method**: Uses `startSheet` instead. This is correct for a streaming writer where "starting" a sheet is the semantic action, not "adding" one. Naming is intentional.
- **`setStylesXml(xml)` accepts raw XML string**: This is a low-level API that bypasses the StyleBuilder. Consider whether this should be documented as advanced/internal.

---

## 14. Worker API (worker-api.ts)

**File:** `worker-api.ts`, 112 lines

### Naming

All camelCase -- correct.

### Issues

- **`XlsxWorker.readBuffer` return type**: Returns `Promise<WorkbookData>` (raw data), not `Promise<Workbook>`. This is **inconsistent** with the top-level `readBuffer()` which returns `Promise<Workbook>`. The worker returns raw data because the `Workbook` class needs WASM which is in the worker thread. This difference should be prominently documented.
- **`XlsxWorker.writeBuffer` parameter**: Takes `WorkbookData` (raw data), not `Workbook`. Same reasoning. Document the difference.
- **`XlsxWorker.readBuffer` options**: Uses `{ password?: string }` inline instead of `ReadOptions`. Should reuse `ReadOptions` type for consistency.
- **`XlsxWorker.writeBuffer` options**: Uses `{ password?: string }` inline instead of `WriteOptions`. Should reuse `WriteOptions` type.
- **`terminate()` return type**: Returns `void`, not `Promise<void>`. This is correct -- termination is synchronous.

**Recommendation:** Reuse `ReadOptions` and `WriteOptions` types in the worker API signatures.

---

## 15. Table Layout Engine (table.ts)

**File:** `table.ts`, 457 lines

### Naming

Functions camelCase, types PascalCase -- correct.

### Parameter Order

| Function | Signature | Issue |
|---|---|---|
| `drawTable(wb, ws, opts)` | workbook, worksheet, options | OK |
| `drawTableFromData(wb, ws, data, opts?)` | workbook, worksheet, data, options | OK |

### Issues

- **`drawTable` and `drawTableFromData` require both `wb` and `ws`**: The workbook is needed for style registration. This is a slightly awkward API -- the user must pass both. Consider whether the worksheet could expose a reference back to its parent workbook.
- **`CellStyle` interface naming**: Conflicts conceptually with `CellStyleData` in types.ts. They serve different purposes (`CellStyle` is a builder input, `CellStyleData` is a stored named style). Consider renaming to `CellStyleOverride` or `TableCellStyle`.
- **`DrawTableOptions.cellStyles` key format**: Uses `"row,col"` string keys (e.g., `"2,3"`). This is fragile and not type-safe. A `Map<string, CellStyle>` or array-based approach would be safer, but this is a common pattern for sparse overrides.
- **`TableResult` fields**: Uses 0-based row indices (`firstDataRow`, `lastDataRow`). Consistent with `decodeCellRef` conventions.
- **`DrawTableFromDataOptions` extends `Omit<DrawTableOptions, 'headers' | 'rows'>`**: Good use of Omit to avoid conflicting fields.

---

## 16. Chart Styles (chart-styles.ts)

**File:** `chart-styles.ts`, 33 lines

### Naming

- `CHART_STYLE_PALETTES` -- UPPER_SNAKE_CASE constant. Correct.
- `getChartStylePalette` -- camelCase function. Correct.

### Issues

- **`CHART_STYLE_PALETTES` type**: `ReadonlyMap<number, readonly string[]>` -- correctly readonly at both levels.
- **Only 8 of 48 style IDs defined**: The JSDoc says "style IDs 1-48" but only IDs 1-8 are in the map. The fallback to default palette handles this, but it may surprise users. Should be documented.

---

## 17. Table Styles (table-styles.ts)

**File:** `table-styles.ts`, 31 lines

### Naming

- `TABLE_STYLES` -- UPPER_SNAKE_CASE. Correct.
- `VALID_TABLE_STYLES` -- UPPER_SNAKE_CASE. Correct.
- `TotalsRowFunction` -- PascalCase type. Correct.

### Issues

- **`TABLE_STYLES` type annotation**: Uses explicit `{ readonly light: readonly string[]; ... }`. Correct.
- **`VALID_TABLE_STYLES` type**: `ReadonlySet<string>`. Correct.

---

## 18. Barcode Module (barcode/)

**File:** `barcode/index.ts` (re-exports), `barcode/common.ts` (types)

### Naming

All encoder functions follow `encode<Format>` pattern -- consistent:
- `encodeQR`, `encodeCode128`, `encodeEAN13`, `encodeUPCA`, `encodeCode39`
- `encodePDF417`, `encodeDataMatrix`, `encodeITF14`, `encodeGS1128`

### Issues

- **`generateBarcode` function**: Top-level convenience function. OK.
- **`generateDrawingXml` and `generateDrawingRels`**: These are internal helpers for OOXML XML generation. They are exported publicly but are implementation details. Consider marking as `@internal` or moving to a separate entry point.
- **`renderBarcodePNG` function**: Public API. OK.
- **`BarcodeMatrix.modules`**: `boolean[][]` -- should be `readonly (readonly boolean[])[]` for an output type.
- **`RenderOptions.foreground` and `background`**: `number[]` -- should be `readonly number[]` and documented as `[R, G, B]` tuple. Consider using `[number, number, number]` tuple type.
- **`ImageAnchor` vs `ChartAnchorData`**: These are structurally similar (both have from/to col/row with offsets) but are separate types with slightly different field names. `ImageAnchor` uses `fromCol`/`toCol` and `ChartAnchorData` also uses `fromCol`/`toCol` but adds `extCx`/`extCy`. Consider making `ImageAnchor` extend or alias `ChartAnchorData` (minus the ext fields).

---

## 19. Formula Engine (formula/)

**File:** `formula/index.ts` (re-exports)

### Naming

All functions camelCase, all types PascalCase -- correct.

### Exported Functions

| Function | Signature Pattern | Issue |
|---|---|---|
| `tokenize(formula)` | data | OK |
| `parseFormula(formula)` | data | OK |
| `parseCellRefValue(value)` | data | OK |
| `serializeFormula(ast)` | data | OK |
| `evaluateFormula(formula, ctx)` | data, context | OK |
| `evaluateNode(node, ctx)` | data, context | OK |
| `resolveRef(ref, ctx)` | data, context | OK |
| `resolveRange(range, ctx)` | data, context | OK |
| `rewriteFormula(formula, action)` | data, action | OK |
| `expandSharedFormula(master, masterRef, childRef)` | formula, source, target | OK |
| `createDefaultFunctions()` | factory | OK |

### Issues

- **`ParseResult.ast`**: `ASTNode | null` -- correctly uses null for failure case.
- **`TokenizeResult.tokens`**: `Token[]` -- **MISSING `readonly`**.
- **`TokenizeResult.errors`**: `string[]` -- **MISSING `readonly`**.
- **`ParseResult.errors`**: `string[]` -- **MISSING `readonly`**.
- **`FunctionCallNode.args`**: `ASTNode[]` -- **MISSING `readonly`**.
- **`ArrayNode.rows`**: `ASTNode[][]` -- **MISSING `readonly`**.
- **`EvalContext.functions`**: `Map<string, FormulaFunction>` -- should be `ReadonlyMap` since the evaluator only reads from it.

---

## 20. Errors (errors.ts)

**File:** `errors.ts`, 35 lines

### Naming

- `ModernXlsxError` -- PascalCase class. Correct.
- Error code constants -- UPPER_SNAKE_CASE. Correct: `INVALID_CELL_REF`, `WASM_INIT_FAILED`, `SHEET_NOT_FOUND`, `COMMENT_NOT_FOUND`, `INVALID_ARGUMENT`.

### Issues

- **Error code type**: Each constant is typed via `as const` (e.g., `'INVALID_CELL_REF' as const`). Good.
- **`ModernXlsxError.code` type**: `string` -- could be narrowed to a union of the exported code constants for better type safety. E.g., `type ErrorCode = typeof INVALID_CELL_REF | typeof WASM_INIT_FAILED | ...`.
- **Missing error codes**: `cell-ref.ts` throws plain `Error` for invalid references instead of `ModernXlsxError` with `INVALID_CELL_REF`. The `validate-chart.ts` throws `RangeError` and `Error` instead of `ModernXlsxError`.

**Recommendation:** Ensure all public-facing throws use `ModernXlsxError` with appropriate codes.

---

## 21. Top-Level Functions (index.ts)

**File:** `index.ts`, 327 lines

### Exported Functions

| Function | Signature | Issue |
|---|---|---|
| `readBuffer(data, options?)` | data, options | OK |
| `readFile(path, options?)` | path, options | OK |
| `writeBlob(wb)` | target | OK |

### Issues

- **`readBuffer` is async but the WASM call is synchronous**: The function is `async` and returns `Promise<Workbook>`, but the actual WASM read is synchronous (`wasmRead` is not async). The async is only needed for the password path (also synchronous). This appears to be future-proofing or API consistency.
- **`writeBlob` is synchronous but `toBuffer` is async**: `writeBlob(wb)` is sync, but `Workbook.toBuffer()` is async. This is inconsistent. Both underlying WASM calls are synchronous.
- **`VERSION` constant**: `'1.0.0' as const` -- UPPER_SNAKE_CASE, correct.
- **No `writeBuffer` top-level function**: There is `readBuffer` but the write equivalent is `Workbook.toBuffer()` (instance method) or `writeBlob()` (browser-only). Consider adding a top-level `writeBuffer(wb)` for symmetry with `readBuffer`.
- **No `writeFile` top-level function**: There is `readFile` but the write equivalent is `Workbook.toFile()`. Consider adding `writeFile(wb, path, options?)` for symmetry.

---

## 22. WASM Loader (wasm-loader.ts)

**File:** `wasm-loader.ts`, 188 lines

### Exported Functions

| Function | Issue |
|---|---|
| `initWasm(wasmSource?)` | OK |
| `initWasmSync(module)` | OK |
| `ensureReady(wasmSource?)` | OK |

### Issues

- **`ensureReady` vs `initWasm`**: Both initialize WASM. `ensureReady` is a convenience wrapper. The naming distinction is unclear. Consider deprecating one or renaming `ensureReady` to `autoInit` or similar.
- **`wasmVersion` re-export**: Exported from wasm-loader but NOT re-exported from index.ts. This means users cannot access the WASM module version. Consider re-exporting.

---

## 23. Summary of Findings

### Critical (should fix before 1.0)

| # | Category | Location | Description |
|---|---|---|---|
| 1 | Return type | `Workbook.getSheet`, `getSheetByIndex`, `getNamedRange` | Returns `undefined` instead of `null` |
| 2 | Return type | `Worksheet.getTable` | Returns `undefined` instead of `null` |
| 3 | Return type | `getBuiltinFormat` | Returns `undefined` instead of `null` |
| 4 | Readonly arrays | `WorksheetData`, `WorkbookData`, `StylesData` | ~30 array fields missing `readonly` modifier |
| 5 | Error handling | `cell-ref.ts` | Throws plain `Error` instead of `ModernXlsxError` |
| 6 | Error handling | `validate-chart.ts` | Throws `RangeError`/`Error` instead of `ModernXlsxError` |

### Important (strong recommendation for 1.0)

| # | Category | Location | Description |
|---|---|---|---|
| 7 | Missing pair | `Workbook` | No `removePivotCache`, `removeSlicerCache`, `removeTimelineCache` |
| 8 | Missing pair | `Worksheet` | No `removeSparklineGroup(index)`, no `removeThreadedComment(id)` |
| 9 | Asymmetric removal | `Worksheet` | `removeChart`/`removePivotTable`/`removeSlicer`/`removeTimeline` use index; other removes use name/key |
| 10 | Missing type export | `dates.ts` | `TemporalLike` interface not exported |
| 11 | Worker API types | `worker-api.ts` | Inline `{ password?: string }` instead of reusing `ReadOptions`/`WriteOptions` |
| 12 | Worker API return | `worker-api.ts` | Returns `WorkbookData` not `Workbook` -- needs prominent documentation |
| 13 | Error code typing | `errors.ts` | `ModernXlsxError.code` is `string`, could be a union type |

### Minor (nice to have for 1.0)

| # | Category | Location | Description |
|---|---|---|---|
| 14 | Naming | `format-cell.ts` `loadFormat` | Parameter order `(formatCode, id)` is reversed from convention |
| 15 | Missing API | `utils.ts` | No `sheetToAoa` function (array-of-arrays output) |
| 16 | Missing API | `index.ts` | No top-level `writeBuffer`/`writeFile` (asymmetric with `readBuffer`/`readFile`) |
| 17 | Readonly | `formula/` types | `TokenizeResult.tokens`, `ParseResult.errors`, `FunctionCallNode.args` missing `readonly` |
| 18 | Readonly | `barcode/common.ts` | `BarcodeMatrix.modules`, `RenderOptions.foreground/background` missing `readonly` |
| 19 | Missing getter/setter | `Workbook` | No `dateSystem` setter |
| 20 | Internal exposure | `barcode/` | `generateDrawingXml`, `generateDrawingRels` are implementation details exported publicly |
| 21 | Duplicate concept | `ImageAnchor` vs `ChartAnchorData` | Structurally similar types not unified |
| 22 | Naming | `wasm-loader.ts` | `ensureReady` vs `initWasm` overlap |
| 23 | Missing re-export | `index.ts` | `wasmVersion` not re-exported |
| 24 | Type safety | `Workbook.toJSON()` | Returns mutable `WorkbookData` -- should return readonly or deep-frozen |

### Statistics

- **Total public exports:** ~180 (types, functions, classes, constants)
- **Critical issues:** 6
- **Important issues:** 7
- **Minor issues:** 11
- **Clean passes:** Naming conventions (camelCase methods, PascalCase types, UPPER_SNAKE_CASE constants), parameter ordering (target first, options last), fluent builder patterns, get/set property pairs on Workbook/Worksheet
