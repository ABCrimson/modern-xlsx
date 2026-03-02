// Core classes

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
export type { FormatCellOptions } from './format-cell.js';
export { formatCell, getBuiltinFormat } from './format-cell.js';
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
  NumFmt,
  PageMarginsData,
  PageSetupData,
  PatternType,
  ProtectionData,
  RichTextRun,
  RowData,
  SharedStringsData,
  SheetData,
  SheetProtectionData,
  StylesData,
  ThemeColorsData,
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
export { initWasm } from './wasm-loader.js';
export { Cell, Workbook, Worksheet } from './workbook.js';

// Internal imports for readBuffer and writeBlob
import { ensureInitialized, wasmRead, wasmWriteBlob } from './wasm-loader.js';
import { Workbook } from './workbook.js';

/**
 * Read an XLSX file buffer and return a Workbook instance.
 * WASM must be initialized first via `initWasm()`.
 *
 * Data crosses the WASM boundary as a JSON string (serde_json in Rust,
 * JSON.parse in JS) for optimal performance — 8-13x faster than
 * serde_wasm_bindgen for large workbooks (100K+ rows).
 */
export async function readBuffer(data: Uint8Array): Promise<Workbook> {
  ensureInitialized();
  const raw = wasmRead(data);
  return new Workbook(raw);
}

/**
 * Write a Workbook directly to a Blob for browser download.
 * WASM must be initialized first via `initWasm()`.
 * Only available in browser environments that support the Blob API.
 */
export function writeBlob(wb: Workbook): Blob {
  ensureInitialized();
  return wasmWriteBlob(wb.toJSON());
}

/**
 * Read an XLSX file from disk and return a Workbook instance.
 * Only available in Node.js, Bun, and Deno environments.
 * WASM must be initialized first via `initWasm()`.
 */
export async function readFile(path: string): Promise<Workbook> {
  const { readFile: fsReadFile } = await import('node:fs/promises');
  const buffer = await fsReadFile(path);
  return readBuffer(new Uint8Array(buffer));
}

export const VERSION = '0.1.0' as const;
