/**
 * Browser IIFE entry point.
 *
 * Bundles everything into a single `modern-xlsx.min.js` file that exposes
 * `window.ModernXlsx` with the full API. The WASM binary is fetched
 * automatically from a sibling `modern-xlsx.wasm` file.
 *
 * Usage:
 * ```html
 * <script src="https://cdn.jsdelivr.net/npm/modern-xlsx/dist/modern-xlsx.min.js"></script>
 * <script>
 *   await ModernXlsx.initWasm();
 *   const wb = new ModernXlsx.Workbook();
 * </script>
 * ```
 */

// Re-export the full public API
export {
  // Sheet utilities
  aoaToSheet,
  // Core classes
  Cell,
  // Cell references
  columnToLetter,
  // Worker
  createXlsxWorker,
  // Dates
  dateToSerial,
  decodeCellRef,
  decodeRange,
  encodeCellRef,
  encodeRange,
  // Formatting
  formatCell,
  getBuiltinFormat,
  // WASM init
  initWasm,
  initWasmSync,
  isDateFormatCode,
  isDateFormatId,
  isTemporalLike,
  jsonToSheet,
  letterToColumn,
  RichTextBuilder,
  // I/O
  readBuffer,
  // Styles
  StyleBuilder,
  serialToDate,
  sheetAddAoa,
  sheetAddJson,
  sheetToCsv,
  sheetToHtml,
  sheetToJson,
  // Version
  VERSION,
  Workbook,
  Worksheet,
  writeBlob,
} from './index.js';
