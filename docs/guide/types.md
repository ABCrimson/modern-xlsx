# Exported types

The package ships its TypeScript declarations, and every type is exported by name from `modern-xlsx`. The [API reference](https://abcrimson.github.io/modern-xlsx/api/) documents each one. This list groups the main ones. The [package README](../../packages/modern-xlsx/README.md#typescript) covers the TypeScript setup.

```typescript
import type {
  WorkbookData,
  WorksheetData,
  SheetData,
  CellData,
  RowData,
  StylesData,
  FontData,
  FillData,
  BorderData,
  AlignmentData,
  ProtectionData,
  CellXfData,
  NumFmt,
  ColumnInfo,
  FrozenPane,
  AutoFilterData,
  FilterColumnData,
  HyperlinkData,
  DataValidationData,
  ConditionalFormattingData,
  PageSetupData,
  SheetProtectionData,
  DocPropertiesData,
  ThemeColorsData,
  WorkbookViewData,
  RichTextRun,
  DefinedNameData,
  CalcChainEntryData,
  // Read and write options, errors
  ReadOptions,
  ReadLimits,
  WriteOptions,
  ModernXlsxErrorCode,
  // Validation & compliance
  ValidationReport,
  ValidationIssue,
  Severity,
  IssueCategory,
  RepairResult,
  // Barcode & image embedding
  BarcodeMatrix,
  DrawBarcodeOptions,
  RenderOptions,
  ImageAnchor,
  // Tables & print layout
  TableDefinitionData,
  TableColumnData,
  TableStyleInfoData,
  HeaderFooterData,
  OutlinePropertiesData,
  // Charts
  ChartDataModel,
  ChartSeriesData,
  ChartAxisData,
  ChartType,
  ChartAnchorData,
  ChartGrouping,
  ChartLegendData,
  // Formula engine
  ASTNode,
  CellValue,
  EvalContext,
  FormulaFunction,
  // Pivot tables & caches
  PivotTableData,
  PivotFieldData,
  PivotDataFieldData,
  PivotAxis,
  SubtotalFunction,
  PivotCacheDefinitionData,
  PivotCacheRecordsData,
  CacheFieldData,
  CacheSource,
  CacheValue,
  // Threaded comments
  ThreadedCommentData,
  PersonData,
  // Slicers & timelines
  SlicerData,
  SlicerCacheData,
  TimelineData,
  TimelineCacheData,
  // Streaming writer
  StreamingXlsxWriter,
  StreamingCellInput,
} from 'modern-xlsx';
```
