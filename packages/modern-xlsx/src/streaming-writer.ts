/**
 * Streaming XLSX writer for large files (100K+ rows).
 *
 * Writes rows directly to ZIP entries inside WASM, so peak memory is
 * proportional to the number of unique strings — not the total row count.
 *
 * @example
 * ```ts
 * import { initWasm, StreamingXlsxWriter } from 'modern-xlsx';
 *
 * await initWasm();
 * const writer = StreamingXlsxWriter.create();
 * writer.startSheet('Sheet1');
 * for (let i = 0; i < 100_000; i++) {
 *   writer.writeRow([
 *     { value: String(i), cellType: 'number' },
 *     { value: `row_${i}`, cellType: 'sharedString' },
 *   ]);
 * }
 * const xlsx: Uint8Array = writer.finish();
 * ```
 */

import { StreamingWriter as WasmStreamingWriter } from '../wasm/modern_xlsx_wasm.js';
import type { CellType } from './types.js';
import { ensureInitialized } from './wasm-loader.js';

// ---------------------------------------------------------------------------
// Public types
// ---------------------------------------------------------------------------

/** A cell value for the streaming writer. */
export interface StreamingCellInput {
  /** Cell value as a string. Omit or set to `undefined`/`null` to skip. */
  value?: string | null;
  /** Cell type. Defaults to `'number'` when omitted. */
  cellType?: CellType;
  /** Style index (0-based, indexes into the cellXfs array). */
  style?: number;
}

// ---------------------------------------------------------------------------
// StreamingXlsxWriter
// ---------------------------------------------------------------------------

/**
 * Streaming XLSX writer backed by the Rust/WASM core.
 *
 * Call {@link create} to instantiate (WASM must be initialized first).
 * Then call {@link startSheet}, {@link writeRow} (repeatedly), and
 * finally {@link finish} to obtain the XLSX bytes.
 */
export class StreamingXlsxWriter {
  #inner: WasmStreamingWriter;

  private constructor(inner: WasmStreamingWriter) {
    this.#inner = inner;
  }

  /**
   * Create a new streaming writer.
   *
   * WASM must be initialized before calling this (via `initWasm()`).
   */
  static create(): StreamingXlsxWriter {
    ensureInitialized();
    return new StreamingXlsxWriter(new WasmStreamingWriter());
  }

  /**
   * Set custom styles XML (the full `xl/styles.xml` content).
   *
   * Must be called before {@link startSheet}. When omitted a minimal
   * default stylesheet (1 font, 2 fills, 1 border, 1 cell format) is used.
   */
  setStylesXml(xml: string): void {
    this.#inner.setStylesXml(xml);
  }

  /**
   * Start a new worksheet with the given name.
   *
   * If another sheet is already open it is closed automatically.
   */
  startSheet(name: string): void {
    this.#inner.startSheet(name);
  }

  /**
   * Write a single row of cells to the currently open sheet.
   *
   * Cells are placed left-to-right starting at column A. Cells with
   * `value: undefined | null` are skipped (sparse row support).
   */
  writeRow(cells: StreamingCellInput[]): void {
    this.#inner.writeRow(JSON.stringify(cells));
  }

  /**
   * Finish writing and return the complete XLSX as a `Uint8Array`.
   *
   * This closes any open sheet, writes the shared string table and
   * metadata parts, and finalizes the ZIP archive.
   *
   * The writer is consumed — calling any method after `finish()` throws.
   */
  finish(): Uint8Array {
    return this.#inner.finish();
  }
}
