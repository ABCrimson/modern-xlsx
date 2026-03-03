// Barcode & QR code generation
export type {
  BarcodeMatrix,
  BarcodeType,
  DrawBarcodeOptions,
  ImageAnchor,
  RenderOptions,
} from './barcode.js';
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
} from './barcode.js';
// Cell reference utilities
export type { CellAddress, CellRange } from './cell-ref.js';
export {
  columnToLetter,
  decodeCellRef,
  decodeRange,
  encodeCellRef,
  encodeRange,
  letterToColumn,
} from './cell-ref.js';
// Date utilities
export {
  dateToSerial,
  isDateFormatCode,
  isDateFormatId,
  isTemporalLike,
  serialToDate,
} from './dates.js';
// Formatting
export type { FormatCellOptions } from './format-cell.js';
export { formatCell, getBuiltinFormat } from './format-cell.js';
// Builders
export { RichTextBuilder } from './rich-text.js';
export { StyleBuilder } from './style-builder.js';
// Types
export type {
  AlignmentData,
  AutoFilterData,
  BorderData,
  BorderSideData,
  BorderStyle,
  CalcChainEntryData,
  CellData,
  CellStyleData,
  CellType,
  CellXfData,
  CfvoData,
  ColorScaleData,
  ColumnInfo,
  CommentData,
  ConditionalFormattingData,
  ConditionalFormattingRuleData,
  DataBarData,
  DataValidationData,
  DateSystem,
  DefinedNameData,
  DocPropertiesData,
  DxfStyleData,
  FillData,
  FilterColumnData,
  FontData,
  FrozenPane,
  GradientFillData,
  GradientStopData,
  HyperlinkData,
  IconSetData,
  IssueCategory,
  NumFmt,
  PageMarginsData,
  PageSetupData,
  PatternType,
  ProtectionData,
  RepairResult,
  RichTextRun,
  RowData,
  Severity,
  SharedStringsData,
  SheetData,
  SheetProtectionData,
  StylesData,
  ThemeColorsData,
  ValidationIssue,
  ValidationReport,
  WorkbookData,
  WorkbookViewData,
  WorksheetData,
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
} from './utils.js';
export {
  aoaToSheet,
  jsonToSheet,
  sheetAddAoa,
  sheetAddJson,
  sheetToCsv,
  sheetToHtml,
  sheetToJson,
} from './utils.js';
// WASM initialization
export { ensureReady, initWasm, initWasmSync } from './wasm-loader.js';
// Core classes
export { Cell, Workbook, Worksheet } from './workbook.js';

// Web Worker support
export type { XlsxWorker, XlsxWorkerOptions } from './worker-api.js';
export { createXlsxWorker } from './worker-api.js';

// Internal imports for readBuffer, writeBlob, readFile
import { ensureInitialized, wasmRead, wasmWriteBlob } from './wasm-loader.js';
import { Workbook as _Workbook } from './workbook.js';

/**
 * Read an XLSX file buffer and return a Workbook instance.
 * WASM must be initialized first via `initWasm()`.
 *
 * Data crosses the WASM boundary as a JSON string (serde_json in Rust,
 * JSON.parse in JS) for optimal performance — 8-13x faster than
 * serde_wasm_bindgen for large workbooks (100K+ rows).
 */
export async function readBuffer(data: Uint8Array): Promise<_Workbook> {
  ensureInitialized();
  const raw = wasmRead(data);
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
 */
export async function readFile(path: string): Promise<_Workbook> {
  const { readFile: fsReadFile } = await import('node:fs/promises');
  const buffer = await fsReadFile(path);
  return readBuffer(new Uint8Array(buffer));
}

export const VERSION = '0.3.0' as const;
