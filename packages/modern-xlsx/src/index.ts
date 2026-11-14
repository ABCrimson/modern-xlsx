/**
 * modern-xlsx — the full build: every export of `modern-xlsx/lite` plus
 * password-protected read/write, the standalone barcode and drawing functions and the
 * Web Worker client.
 */

import { readFileBytes } from '#runtime';
import { StreamingXlsxWriter as StreamingXlsxWriterBase } from './io/streaming-writer.js';
import type { ReadOptions } from './model/types.js';
import { Workbook as WorkbookBase } from './model/workbook.js';
import { serializeReadLimits } from './read-limits.js';
import type { WasmLoader } from './wasm-backend.js';
import { loader } from './wasm-loader.js';

// Standalone barcode, QR code and drawing generation
export type {
  BarcodeMatrix,
  BarcodeType,
  DrawBarcodeOptions,
  ImageAnchor,
  QrEcLevel,
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
// Everything the lite entry exports too
export * from './exports-shared.js';
// WASM initialization
export { ensureReady, initWasm, initWasmSync } from './wasm-loader.js';
// Web Worker support
export type { XlsxWorker, XlsxWorkerOptions } from './worker-api.js';
export { createXlsxWorker } from './worker-api.js';

/**
 * Represents an Excel workbook containing sheets, styles, and metadata.
 *
 * This is the root object for reading, writing, and manipulating XLSX files.
 * Obtain an instance via {@link readBuffer}, {@link readFile}, or construct
 * a new empty workbook with `new Workbook()`. Its WASM-backed methods use the
 * full build, which supports password-protected files.
 *
 * @example
 * ```ts
 * const wb = new Workbook();
 * const ws = wb.addSheet('Data');
 * ws.cell('A1').value = 'Hello';
 * const bytes = await wb.toBuffer();
 * ```
 */
export class Workbook extends WorkbookBase {
  protected static override loader: WasmLoader = loader;
}

/**
 * Streaming XLSX writer backed by the full Rust/WASM build.
 *
 * Call {@link StreamingXlsxWriter.create} to instantiate (WASM must be
 * initialized first). Then call `startSheet`, `writeRow` (repeatedly), and
 * finally `finish` to obtain the XLSX bytes.
 */
export class StreamingXlsxWriter extends StreamingXlsxWriterBase {
  protected static override loader: WasmLoader = loader;
}

/**
 * Read an XLSX file buffer and return a Workbook instance.
 * WASM must be initialized first via `initWasm()`.
 *
 * Data crosses the WASM boundary as a JSON string (serde_json in Rust,
 * JSON.parse in JS).
 *
 * @param data - Raw XLSX bytes (possibly encrypted OLE2 container).
 * @param options - Optional read options. Pass `{ password: '...' }` for encrypted files and
 *   `{ limits: { ... } }` to raise or tighten the security and size limits (see {@link ReadLimits}).
 * @returns A fully-parsed Workbook instance ready for reading/writing.
 *
 * @example
 * ```ts
 * import { initWasm, readBuffer } from 'modern-xlsx';
 *
 * await initWasm();
 * const wb = await readBuffer(fileBytes);
 * const ws = wb.getSheet('Sheet1');
 * ```
 */
export async function readBuffer(data: Uint8Array, options?: ReadOptions): Promise<Workbook> {
  loader.ensureInitialized();
  const limitsJson = serializeReadLimits(options?.limits);
  const raw = options?.password
    ? loader.wasmReadWithPassword(data, options.password, limitsJson)
    : loader.wasmRead(data, limitsJson);
  return new Workbook(raw);
}

/**
 * Write a Workbook directly to a Blob for browser download.
 * WASM must be initialized first via `initWasm()`.
 * Only available in browser environments that support the Blob API.
 *
 * @param wb - The Workbook to serialize.
 * @returns A Blob containing the XLSX file bytes.
 * @throws ModernXlsxError with code `INVALID_WORKBOOK` for the causes listed with that code in `errors.ts`, and `IO_ERROR` if the Blob cannot be constructed.
 *
 * @example
 * ```ts
 * const blob = writeBlob(wb);
 * const url = URL.createObjectURL(blob);
 * anchor.href = url;
 * ```
 */
export function writeBlob(wb: WorkbookBase): Blob {
  loader.ensureInitialized();
  return loader.wasmWriteBlob(wb.toJSON());
}

/**
 * Read an XLSX file from disk and return a Workbook instance.
 * Only available in Node.js, Bun, and Deno environments.
 * WASM must be initialized first via `initWasm()`.
 *
 * @param path - File path to read.
 * @param options - Optional read options. Pass `{ password: '...' }` for encrypted files and
 *   `{ limits: { ... } }` to raise or tighten the security and size limits (see {@link ReadLimits}).
 * @returns A fully-parsed Workbook instance ready for reading/writing.
 * @throws ModernXlsxError with code `IO_ERROR` outside Node.js-compatible runtimes
 *   (browsers, workers, the IIFE build).
 *
 * @example
 * ```ts
 * import { initWasm, readFile } from 'modern-xlsx';
 *
 * await initWasm();
 * const wb = await readFile('./report.xlsx');
 * console.log(wb.sheetNames);
 * ```
 */
export async function readFile(path: string, options?: ReadOptions): Promise<Workbook> {
  return readBuffer(await readFileBytes(path), options);
}
