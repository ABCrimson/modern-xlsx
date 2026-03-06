// Barcode & QR code generation
export type {
  BarcodeMatrix,
  BarcodeType,
  DrawBarcodeOptions,
  ImageAnchor,
  RenderOptions,
} from './barcode/index.js';
export {
  encodeCode39,
  encodeCode128,
  encodeDataMatrix,
  encodeEAN13,
  encodeGS1128,
  encodeITF14,
  encodePDF417,
  encodeQR,
  encodeUPCA,
  generateBarcode,
  generateDrawingRels,
  generateDrawingXml,
  renderBarcodePNG,
} from './barcode/index.js';
// Cell reference utilities
export type { CellAddress, CellRange, SplitCellRef } from './cell-ref.js';
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
} from './cell-ref.js';
// Builders
export type { AddSeriesOptions, AxisOptions } from './chart-builder.js';
export { ChartBuilder } from './chart-builder.js';
// Chart style presets
export { CHART_STYLE_PALETTES, getChartStylePalette } from './chart-styles.js';
// Date utilities
export {
  dateToSerial,
  isDateFormatCode,
  isDateFormatId,
  isTemporalLike,
  serialToDate,
} from './dates.js';
// Errors
export {
  COMMENT_NOT_FOUND,
  INVALID_ARGUMENT,
  INVALID_CELL_REF,
  ModernXlsxError,
  SHEET_NOT_FOUND,
  WASM_INIT_FAILED,
} from './errors.js';
// Formatting
export type { FormatCellOptions, FormatCellResult } from './format-cell.js';
export {
  formatCell,
  formatCellRich,
  getBuiltinFormat,
  loadFormat,
  loadFormatTable,
} from './format-cell.js';
// Formula engine
export type {
  ArrayNode,
  ASTNode,
  BinaryOpNode,
  BooleanNode,
  CellRefNode,
  CellValue,
  ErrorNode,
  EvalContext,
  FormulaFunction,
  FunctionCallNode,
  NameNode,
  NumberNode,
  ParseResult,
  PercentNode,
  RangeNode,
  RewriteAction,
  StringNode,
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
export { HeaderFooterBuilder } from './header-footer.js';
export { RichTextBuilder } from './rich-text.js';
export { StyleBuilder } from './style-builder.js';
// Table layout engine
export type {
  CellStyle,
  DrawTableFromDataOptions,
  DrawTableOptions,
  TableColumn,
  TableResult,
} from './table.js';
export { drawTable, drawTableFromData } from './table.js';
// Table styles
export type { TotalsRowFunction } from './table-styles.js';
export { TABLE_STYLES, VALID_TABLE_STYLES } from './table-styles.js';
// Types
export type {
  AlignmentData,
  AutoFilterData,
  AxisPosition,
  BorderData,
  BorderSideData,
  BorderStyle,
  CalcChainEntryData,
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
  ColorScaleData,
  ColumnInfo,
  CommentData,
  ConditionalFormattingData,
  ConditionalFormattingRuleData,
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
  NumFmt,
  OutlinePropertiesData,
  PageMarginsData,
  PageSetupData,
  PaneSelectionData,
  PatternType,
  PersonData,
  PivotAxis,
  PivotDataFieldData,
  PivotFieldData,
  PivotFieldRef,
  PivotItem,
  PivotLocation,
  PivotPageFieldData,
  PivotTableData,
  ProtectionData,
  RadarStyle,
  ReadOptions,
  RepairResult,
  RichTextRun,
  RowData,
  ScatterStyle,
  Severity,
  SharedStringsData,
  SheetData,
  SheetProtectionData,
  SheetState,
  SheetViewData,
  SlicerCacheData,
  SlicerData,
  SlicerItem,
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
  TrendlineData,
  TrendlineType,
  ValidationIssue,
  ValidationReport,
  View3DData,
  ViewMode,
  WorkbookData,
  WorkbookProtectionData,
  WorkbookViewData,
  WorksheetChartData,
  WorksheetData,
  WriteOptions,
} from './types.js';
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
} from './utils.js';
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
} from './utils.js';
// Chart validation
export { validateChartData } from './validate-chart.js';
// WASM initialization
export { ensureReady, initWasm, initWasmSync } from './wasm-loader.js';
// Core classes
export { Cell, Workbook, Worksheet } from './workbook.js';

// Web Worker support
export type { XlsxWorker, XlsxWorkerOptions } from './worker-api.js';
export { createXlsxWorker } from './worker-api.js';

import type { ReadOptions } from './types.js';
// Internal imports for readBuffer, writeBlob, readFile
import { ensureInitialized, wasmRead, wasmReadWithPassword, wasmWriteBlob } from './wasm-loader.js';
import { Workbook as _Workbook } from './workbook.js';

/**
 * Read an XLSX file buffer and return a Workbook instance.
 * WASM must be initialized first via `initWasm()`.
 *
 * Data crosses the WASM boundary as a JSON string (serde_json in Rust,
 * JSON.parse in JS) for optimal performance — 8-13x faster than
 * serde_wasm_bindgen for large workbooks (100K+ rows).
 *
 * @param data - Raw XLSX bytes (possibly encrypted OLE2 container).
 * @param options - Optional read options. Pass `{ password: '...' }` for encrypted files.
 */
export async function readBuffer(data: Uint8Array, options?: ReadOptions): Promise<_Workbook> {
  ensureInitialized();
  const raw = options?.password ? wasmReadWithPassword(data, options.password) : wasmRead(data);
  return new _Workbook(raw);
}

/**
 * Write a Workbook directly to a Blob for browser download.
 * WASM must be initialized first via `initWasm()`.
 * Only available in browser environments that support the Blob API.
 */
export function writeBlob(wb: _Workbook): Blob {
  ensureInitialized();
  return wasmWriteBlob(wb.toJSON());
}

/**
 * Read an XLSX file from disk and return a Workbook instance.
 * Only available in Node.js, Bun, and Deno environments.
 * WASM must be initialized first via `initWasm()`.
 *
 * @param path - File path to read.
 * @param options - Optional read options. Pass `{ password: '...' }` for encrypted files.
 */
export async function readFile(path: string, options?: ReadOptions): Promise<_Workbook> {
  const { readFile: fsReadFile } = await import('node:fs/promises');
  const buffer = await fsReadFile(path);
  return readBuffer(new Uint8Array(buffer), options);
}

export const VERSION = '0.8.6' as const;
