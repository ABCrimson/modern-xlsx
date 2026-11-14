/**
 * modern-xlsx/lite — the entry bound to the WASM build with encryption compiled out.
 *
 * Differences from the `modern-xlsx` entry:
 * - Password-protected read/write: a `password` option throws a
 *   `ModernXlsxError` with code `INVALID_ARGUMENT`.
 * - The standalone barcode and drawing exports are absent: `generateBarcode`,
 *   `renderBarcodePNG`, `encodeQR`, `encodeCode128`, `encodeCode39`, `encodeDataMatrix`,
 *   `encodeEAN13`, `encodeGS1128`, `encodeITF14`, `encodePDF417`, `encodeUPCA`,
 *   `generateDrawingXml` and `generateDrawingRels`. `Workbook#addBarcode` is
 *   inherited from the shared `Workbook` and works here.
 * - The Web Worker client (`createXlsxWorker`) is absent.
 * - The barcode and worker types are absent too: `BarcodeMatrix`, `BarcodeType`,
 *   `DrawBarcodeOptions`, `ImageAnchor`, `QrEcLevel`, `RenderOptions`, `XlsxWorker` and
 *   `XlsxWorkerOptions`.
 *
 * Everything else exported by `modern-xlsx` is the same, and the lite
 * `.wasm` binary is smaller.
 */

import { readFileBytes } from '#runtime';
import { StreamingXlsxWriter as StreamingXlsxWriterBase } from './io/streaming-writer.js';
import { INVALID_ARGUMENT, ModernXlsxError } from './model/errors.js';
import type { ReadOptions } from './model/types.js';
import { Workbook as WorkbookBase } from './model/workbook.js';
import { serializeReadLimits } from './read-limits.js';
import { LITE_PASSWORD_MESSAGE, type WasmLoader } from './wasm-backend.js';
import { loader } from './wasm-loader-lite.js';

// Everything the full entry exports except the standalone barcode and drawing exports
// (functions and types) and the worker client
export * from './exports-shared.js';
// WASM initialization (lite loader — no encryption)
export { ensureReady, initWasm, initWasmSync } from './wasm-loader-lite.js';

/**
 * Represents an Excel workbook containing sheets, styles, and metadata.
 *
 * This is the root object for reading, writing, and manipulating XLSX files.
 * Obtain an instance via {@link readBuffer}, {@link readFile}, or construct
 * a new empty workbook with `new Workbook()`. Its WASM-backed methods use the
 * lite build: `toBuffer({ password })` throws `INVALID_ARGUMENT`.
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
 * Streaming XLSX writer backed by the lite Rust/WASM build.
 *
 * Call {@link StreamingXlsxWriter.create} to instantiate (WASM must be
 * initialized first). Then call `startSheet`, `writeRow` (repeatedly), and
 * finally `finish` to obtain the XLSX bytes.
 */
export class StreamingXlsxWriter extends StreamingXlsxWriterBase {
  protected static override loader: WasmLoader = loader;
}

/** Reject a password, which the lite build cannot handle. */
function rejectPassword(options: ReadOptions | undefined): void {
  if (options?.password) {
    throw new ModernXlsxError(INVALID_ARGUMENT, LITE_PASSWORD_MESSAGE);
  }
}

/**
 * Read an XLSX file buffer and return a Workbook instance (lite build).
 * WASM must be initialized first via `initWasm()`.
 *
 * The lite build cannot decrypt. Input that is not a plain ZIP is still classified:
 * an encrypted file rejects with `PASSWORD_PROTECTED`, a legacy .xls with
 * `LEGACY_FORMAT`, and any other OLE2 or unrecognized bytes with `UNRECOGNIZED_FORMAT`.
 * Use the full `modern-xlsx` entry to read encrypted files.
 *
 * @param data - Raw XLSX bytes (must be a plain ZIP, not encrypted OLE2).
 * @param options - Optional read options: `{ limits: { ... } }` raises or tightens the
 *   security and size limits (see {@link ReadLimits}); `password` is not supported.
 * @returns A fully-parsed Workbook instance ready for reading/writing.
 * @throws ModernXlsxError with code `INVALID_ARGUMENT` if `options.password` is set,
 *   before any WASM call; `PASSWORD_PROTECTED`, `LEGACY_FORMAT` or `UNRECOGNIZED_FORMAT`
 *   for non-ZIP input as above.
 */
export async function readBuffer(data: Uint8Array, options?: ReadOptions): Promise<Workbook> {
  rejectPassword(options);
  loader.ensureInitialized();
  const raw = loader.wasmRead(data, serializeReadLimits(options?.limits));
  return new Workbook(raw);
}

/**
 * Write a Workbook directly to a Blob for browser download (lite build).
 * WASM must be initialized first via `initWasm()`.
 * Only available in browser environments that support the Blob API.
 *
 * @param wb - The Workbook to serialize.
 * @returns A Blob containing the XLSX file bytes.
 * @throws ModernXlsxError with code `INVALID_WORKBOOK` for the causes listed with that code in `errors.ts`, and `IO_ERROR` if the Blob cannot be constructed.
 */
export function writeBlob(wb: WorkbookBase): Blob {
  loader.ensureInitialized();
  return loader.wasmWriteBlob(wb.toJSON());
}

/**
 * Read an XLSX file from disk and return a Workbook instance (lite build).
 * Only available in Node.js, Bun, and Deno environments.
 * WASM must be initialized first via `initWasm()`.
 *
 * The lite build cannot decrypt. Input that is not a plain ZIP is still classified:
 * an encrypted file rejects with `PASSWORD_PROTECTED`, a legacy .xls with
 * `LEGACY_FORMAT`, and any other OLE2 or unrecognized bytes with `UNRECOGNIZED_FORMAT`.
 * Use the full `modern-xlsx` entry to read encrypted files.
 *
 * @param path - File path to read.
 * @param options - Optional read options: `limits` is honoured, `password` is not supported.
 * @returns A fully-parsed Workbook instance ready for reading/writing.
 * @throws ModernXlsxError with code `INVALID_ARGUMENT` if `options.password` is set,
 *   before the file is read, and `IO_ERROR` outside Node.js-compatible runtimes
 *   (browsers, workers).
 */
export async function readFile(path: string, options?: ReadOptions): Promise<Workbook> {
  rejectPassword(options);
  return readBuffer(await readFileBytes(path), options);
}
