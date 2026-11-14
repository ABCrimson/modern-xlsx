/**
 * Exports shared by the `modern-xlsx` and `modern-xlsx/lite` entries.
 *
 * Both entries `export *` from here. Each entry adds what is bound to its WASM
 * build: `Workbook`, `StreamingXlsxWriter`, `initWasm`, `initWasmSync`,
 * `ensureReady`, `readBuffer`, `readFile` and `writeBlob`. The full entry also
 * adds the standalone barcode and drawing functions with their types, and
 * `createXlsxWorker` with its types. Nothing here imports a WASM glue module.
 */

// The version's only source. The bundler inlines just the `version` field.
import pkg from '../package.json' with { type: 'json' };

// Builders
export type { AddSeriesOptions, AxisOptions } from './authoring/chart-builder.js';
export { ChartBuilder } from './authoring/chart-builder.js';
// Chart style presets
export { CHART_STYLE_PALETTES, getChartStylePalette } from './authoring/chart-styles.js';
// Builders: header/footer, pivot table, rich text and style
export { HeaderFooterBuilder } from './authoring/header-footer.js';
export type {
  PivotDataFieldOptions,
  PivotFieldOptions,
  PivotPageFieldOptions,
} from './authoring/pivot-builder.js';
export { PivotTableBuilder } from './authoring/pivot-builder.js';
export { RichTextBuilder } from './authoring/rich-text.js';
export { StyleBuilder } from './authoring/style-builder.js';
// Table layout engine
export type {
  CellStyle,
  DrawTableFromDataOptions,
  DrawTableOptions,
  TableColumn,
  TableResult,
} from './authoring/table.js';
export { drawTable, drawTableFromData } from './authoring/table.js';
// Table styles
export { TABLE_STYLES, VALID_TABLE_STYLES } from './authoring/table-styles.js';
// Chart validation
export { validateChartData } from './authoring/validate-chart.js';
// Cell reference utilities
export type { CellAddress, CellRange, SplitCellRef } from './format/cell-ref.js';
export {
  columnToLetter,
  decodeCellRef,
  decodeRange,
  decodeRow,
  encodeCellRef,
  encodeRange,
  encodeRow,
  letterToColumn,
  splitCellRef,
} from './format/cell-ref.js';
// Date utilities
export type { TemporalLike } from './format/dates.js';
export {
  dateToSerial,
  isDateFormatCode,
  isDateFormatId,
  isTemporalLike,
  serialToDate,
} from './format/dates.js';
// Formatting
export type { FormatCellOptions, FormatCellResult } from './format/format-cell.js';
export {
  formatCell,
  formatCellRich,
  getBuiltinFormat,
  loadFormat,
  loadFormatTable,
} from './format/format-cell.js';
// Formula engine
export type {
  ArrayNode,
  ASTNode,
  BinaryOpNode,
  BooleanNode,
  CellRefNode,
  CellValue,
  ColumnRangeNode,
  ErrorNode,
  EvalContext,
  FormulaFunction,
  FunctionCallNode,
  MissingArgNode,
  NameNode,
  NumberNode,
  ParseResult,
  PercentNode,
  RangeNode,
  RewriteAction,
  RowRangeNode,
  StringNode,
  StructuredRefNode,
  Token,
  TokenizeResult,
  TokenType,
  UnaryOpNode,
} from './formula/index.js';
export {
  createDefaultFunctions,
  evaluateFormula,
  evaluateNode,
  expandSharedFormula,
  parseCellRefValue,
  parseFormula,
  resolveRange,
  resolveRef,
  rewriteFormula,
  serializeFormula,
  tokenize,
} from './formula/index.js';
// Sheet conversion utilities
export type {
  AoaToSheetOptions,
  JsonToSheetOptions,
  SheetAddAoaOptions,
  SheetAddJsonOptions,
  SheetToCsvOptions,
  SheetToHtmlOptions,
  SheetToJsonOptions,
  SheetToTxtOptions,
} from './io/sheet-utils.js';
export {
  aoaToSheet,
  jsonToSheet,
  sheetAddAoa,
  sheetAddJson,
  sheetToCsv,
  sheetToFormulae,
  sheetToHtml,
  sheetToJson,
  sheetToTxt,
} from './io/sheet-utils.js';
// Streaming writer (the loader-bound StreamingXlsxWriter class comes from each entry)
export type { StreamingCellInput } from './io/streaming-writer.js';
// Core classes: Cell here, Worksheet below (each entry exports the loader-bound Workbook)
export { Cell } from './model/cell.js';
// Errors
export type { KnownModernXlsxErrorCode, ModernXlsxErrorCode } from './model/errors.js';
export {
  COMMENT_NOT_FOUND,
  INVALID_ARGUMENT,
  INVALID_CELL_REF,
  INVALID_CELL_VALUE,
  INVALID_DATE,
  INVALID_FORMAT,
  INVALID_STYLE,
  INVALID_WORKBOOK,
  IO_ERROR,
  JSON_ERROR,
  LEGACY_FORMAT,
  MISSING_PART,
  ModernXlsxError,
  PASSWORD_PROTECTED,
  SECURITY,
  SHEET_NOT_FOUND,
  SIZE_LIMIT,
  UNRECOGNIZED_FORMAT,
  WASM_ERROR,
  WASM_INIT_FAILED,
  WORKER_ERROR,
  XML_PARSE,
  XML_WRITE,
  ZIP_ENTRY,
  ZIP_FINALIZE,
  ZIP_READ,
  ZIP_WRITE,
} from './model/errors.js';
// Types
export type {
  AlignmentData,
  AutoFilterData,
  AxisPosition,
  BorderData,
  BorderSideData,
  BorderStyle,
  CacheFieldData,
  CacheSource,
  CacheValue,
  CacheValueBoolean,
  CacheValueDateTime,
  CacheValueError,
  CacheValueIndex,
  CacheValueMissing,
  CacheValueNumber,
  CacheValueString,
  CacheValueType,
  CalcChainEntryData,
  CalcPrData,
  CellData,
  CellStyleData,
  CellType,
  CellXfData,
  CfvoData,
  ChartAnchorData,
  ChartAxisData,
  ChartDataModel,
  ChartGrouping,
  ChartLegendData,
  ChartSeriesData,
  ChartTitleData,
  ChartType,
  ColorData,
  ColorScaleData,
  ColumnInfo,
  CommentData,
  ConditionalFormattingData,
  ConditionalFormattingRuleData,
  CustomFilterData,
  CustomFiltersData,
  DataBarData,
  DataLabelsData,
  DataValidationData,
  DateSystem,
  DefinedNameData,
  DocPropertiesData,
  DxfStyleData,
  ErrorBarDirection,
  ErrorBarsData,
  ErrorBarType,
  FileVersionData,
  FillData,
  FilterColumnData,
  FontData,
  FrozenPane,
  GradientFillData,
  GradientStopData,
  HeaderFooterData,
  HyperlinkData,
  IconSetData,
  IssueCategory,
  LegendPosition,
  ManualLayoutData,
  MarkerStyleType,
  MentionData,
  NumFmt,
  OutlinePropertiesData,
  PageBreakData,
  PageBreaksData,
  PageMarginsData,
  PageSetupData,
  PaneSelectionData,
  PatternType,
  PersonData,
  PivotAxis,
  PivotCacheDefinitionData,
  PivotCacheRecordsData,
  PivotDataFieldData,
  PivotFieldData,
  PivotFieldRef,
  PivotItem,
  PivotLocation,
  PivotPageFieldData,
  PivotTableData,
  ProtectionData,
  RadarStyle,
  ReadLimits,
  ReadOptions,
  RepairResult,
  RichTextRun,
  RowData,
  ScatterStyle,
  Severity,
  SharedItemsAttrs,
  SharedStringsData,
  SheetData,
  SheetProtectionData,
  SheetState,
  SheetViewData,
  SlicerCacheData,
  SlicerData,
  SlicerItem,
  SlicerPivotTable,
  SortOrder,
  SparklineData,
  SparklineGroupData,
  SparklineType,
  SplitPaneData,
  StylesData,
  SubtotalFunction,
  TableColumnData,
  TableDefinitionData,
  TableStyleInfoData,
  ThemeColorsData,
  ThreadedCommentData,
  TickLabelPosition,
  TickMark,
  TimelineCacheData,
  TimelineData,
  TimelineLevel,
  TotalsRowFunction,
  TrendlineData,
  TrendlineType,
  ValidationIssue,
  ValidationReport,
  View3DData,
  ViewMode,
  WorkbookData,
  WorkbookPrData,
  WorkbookProtectionData,
  WorkbookViewData,
  WorksheetChartData,
  WorksheetData,
  WriteOptions,
} from './model/types.js';
// Core classes: Worksheet
export { Worksheet } from './model/worksheet.js';

/**
 * The current library version string: the `version` field of the package's package.json,
 * inlined at build time. The published declarations type it as that exact string literal.
 */
export const VERSION: string = pkg.version;
