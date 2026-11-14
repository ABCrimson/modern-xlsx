import {
  INVALID_ARGUMENT,
  INVALID_CELL_VALUE,
  ModernXlsxError,
  WASM_INIT_FAILED,
} from '../model/errors.js';
import { nameKey } from '../model/name-rules.js';
import type { CellType } from '../model/types.js';
import type { StreamingWriterHandle, WasmLoader } from '../wasm-backend.js';

// ---------------------------------------------------------------------------
// Public types
// ---------------------------------------------------------------------------

/**
 * A cell value for the streaming writer.
 *
 * {@link StreamingXlsxWriter.writeRow} validates every cell before writing
 * anything and throws a {@link ModernXlsxError} for invalid input:
 *
 * - `INVALID_CELL_VALUE`: a `'number'` value (also the default type) that is not
 *   a finite number (including `''` and text with `cellType` omitted), a
 *   `'boolean'` value other than `true`/`false` (any case), `1` or `0`, an
 *   `'error'` value that does not start with `#`, or any `'formulaStr'` cell.
 * - `INVALID_CELL_REF`: a cell past column XFD (the 16,385th) or a row past 1,048,576.
 * - `JSON`: a cell the WASM bridge cannot deserialize, such as a `cellType` outside {@link CellType}.
 */
export interface StreamingCellInput {
  /**
   * Cell value as a string. Omit or set to `undefined`/`null` to skip the cell
   * (except a `'stub'` with a `style`, which is still written).
   */
  value?: string | null;
  /**
   * Cell type. Defaults to `'number'` when omitted. A `'stub'` is written as an
   * empty cell (its `value` is ignored); a stub without a `value` is written only
   * when it carries a `style`.
   * `'error'` is written as a real error cell, `'inlineStr'` as an inline string.
   * `'formulaStr'` is not supported and throws.
   */
  cellType?: CellType;
  /** Style index (0-based, indexes into the cellXfs array). */
  style?: number;
}

// ---------------------------------------------------------------------------
// StreamingXlsxWriter
// ---------------------------------------------------------------------------

/**
 * Streaming XLSX writer backed by the Rust/WASM core, for large files (100K+ rows).
 *
 * Call {@link create} to instantiate (WASM must be initialized first).
 * Then call {@link startSheet}, {@link writeRow} (repeatedly), and
 * finally {@link finish} to obtain the XLSX bytes.
 *
 * Each row is serialized and compressed into the ZIP as it is written, so memory
 * holds the compressed output and the shared-string table, not every row's cell
 * data. The whole compressed archive stays in WASM memory until `finish()`
 * returns it.
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
 *
 * This class is the unbound base: the `modern-xlsx` and `modern-xlsx/lite`
 * entries each export a subclass bound to their WASM build, and `create()` on
 * the base itself throws `WASM_INIT_FAILED`.
 */
export class StreamingXlsxWriter {
  /** The WASM build an entry subclass binds; unset on the base. */
  protected static loader?: WasmLoader;

  readonly #inner: StreamingWriterHandle;
  #disposed = false;
  /** Set by the first {@link finish} call, which spends the Rust-side writer whether it succeeds or not. */
  #spent = false;
  /** The sheets started so far: each name's {@link nameKey} mapped to the name. */
  readonly #sheetKeys = new Map<string, string>();

  /**
   * Open a Rust-side writer through the loader of the class being constructed
   * (`new.target`), so each entry subclass writes through its own WASM build.
   */
  protected constructor() {
    const { loader } = new.target;
    if (!loader) {
      throw new ModernXlsxError(
        WASM_INIT_FAILED,
        "StreamingXlsxWriter is not bound to a WASM build; import it from 'modern-xlsx' or 'modern-xlsx/lite'.",
      );
    }
    loader.ensureInitialized();
    this.#inner = new loader.StreamingWriter();
  }

  /**
   * Create a new streaming writer.
   *
   * WASM must be initialized before calling this (via `initWasm()`).
   *
   * @returns A new writer, an instance of the class `create` is called on.
   * @throws {@link ModernXlsxError} with code `WASM_INIT_FAILED` if WASM is not
   *   initialized or the class is the unbound base.
   *
   * @example
   * ```ts
   * await initWasm();
   * const writer = StreamingXlsxWriter.create();
   * ```
   */
  static create(): StreamingXlsxWriter {
    return new this();
  }

  /** The Rust-side handle, or an `INVALID_ARGUMENT` error once the writer is disposed. */
  #live(): StreamingWriterHandle {
    if (this.#disposed) {
      throw new ModernXlsxError(INVALID_ARGUMENT, 'StreamingXlsxWriter has been disposed');
    }
    return this.#inner;
  }

  /**
   * Release the Rust-side writer. Idempotent; every other method throws
   * `INVALID_ARGUMENT` afterwards. Enables `using writer = StreamingXlsxWriter.create()`.
   * Calling it after {@link finish} is safe.
   */
  [Symbol.dispose](): void {
    if (this.#disposed) return;
    this.#disposed = true;
    this.#inner.free();
  }

  /**
   * Set custom styles XML (the full `xl/styles.xml` content).
   *
   * Must be called before {@link startSheet}. When omitted a minimal
   * default stylesheet (1 font, 2 fills, 1 border, 1 cell format) is used.
   *
   * @param xml - The complete `xl/styles.xml` content string.
   * @throws {@link ModernXlsxError} with code `INVALID_ARGUMENT` (message
   * 'StreamingWriter already finished') after {@link finish}.
   */
  setStylesXml(xml: string): void {
    const inner = this.#live();
    try {
      inner.setStylesXml(xml);
    } catch (err) {
      throw ModernXlsxError.fromWasmError(err);
    }
  }

  /**
   * Start a new worksheet with the given name.
   *
   * If another sheet is already open it is closed automatically.
   *
   * @param name - The sheet tab name.
   * @throws {@link ModernXlsxError} with code `INVALID_CELL_VALUE` when the name is
   * empty, longer than 31 characters, contains `\ / * ? : [ ]` or a control
   * character, starts or ends with an apostrophe, is the reserved name `History`
   * (any case), or duplicates an earlier sheet name, compared as addSheet
   * compares (ignoring case and canonical equivalence, final sigma folded:
   * `café` with U+00E9 repeats `cafe` + U+0301, and Excel refuses a file
   * holding both). The WASM core also refuses as a duplicate a name that
   * differs from an earlier one only in one of 20 letters whose case pairs are
   * new in Unicode 18 (such as U+A7DD beside U+0277), which addSheet accepts:
   * the WASM core lower-cases with the Unicode 18.0.0 tables of Rust
   * nightly-2026-11-14, and Node 26.13.0's `toLowerCase` follows Unicode 17.0
   * (ICU 78.3), where these 20 capital letters are not assigned.
   * The writer is left unchanged, so a currently open sheet stays open, and a
   * refused name does not count as started. A name that is
   * not a string (message 'invalid cell value: Invalid sheet name: not a
   * string') is refused the same way.
   * @throws {@link ModernXlsxError} with code `INVALID_ARGUMENT` (message
   * 'StreamingWriter already finished') after {@link finish} for a string name
   * (a name that is not a string is refused first and still gets
   * `INVALID_CELL_VALUE`), and (message 'StreamingXlsxWriter has been
   * disposed') after `[Symbol.dispose]()`.
   */
  startSheet(name: string): void {
    const inner = this.#live();
    if (typeof name !== 'string') {
      throw new ModernXlsxError(
        INVALID_CELL_VALUE,
        'invalid cell value: Invalid sheet name: not a string',
      );
    }
    // After finish() the Rust side refuses every call as finished; leave that error to it.
    const key = this.#spent ? undefined : nameKey(name);
    const earlier = key === undefined ? undefined : this.#sheetKeys.get(key);
    if (earlier !== undefined) {
      throw new ModernXlsxError(
        INVALID_CELL_VALUE,
        `invalid cell value: Invalid sheet name '${name}': duplicate of '${earlier}'`,
      );
    }
    try {
      inner.startSheet(name);
    } catch (err) {
      throw ModernXlsxError.fromWasmError(err);
    }
    if (key !== undefined) this.#sheetKeys.set(key, name);
  }

  /**
   * Write a single row of cells to the currently open sheet.
   *
   * Cells are placed left-to-right starting at column A. Cells with
   * `value: undefined | null` are skipped (sparse row support), except a
   * `'stub'` cell with a `style`, which is written as an empty styled cell.
   *
   * The whole row is validated before anything is written, so a thrown error
   * leaves the writer untouched and you may keep writing rows.
   *
   * @param cells - Array of cell inputs for the row.
   * @throws {@link ModernXlsxError} with code `INVALID_CELL_VALUE` for a non-numeric or
   * non-finite `'number'` value (including `''` and text with `cellType` omitted),
   * an unknown `'boolean'`, an `'error'` value without a leading `#`, or a
   * `'formulaStr'` cell.
   * @throws {@link ModernXlsxError} with code `INVALID_CELL_REF` for a cell past
   * column XFD or a row past 1,048,576.
   * @throws {@link ModernXlsxError} with code `JSON` for a cell the WASM bridge cannot
   * deserialize, such as a `cellType` outside {@link CellType}.
   * @throws {@link ModernXlsxError} with code `XML_WRITE` when no sheet is open (call
   * {@link startSheet} first).
   * @throws {@link ModernXlsxError} with code `INVALID_ARGUMENT` (message
   * 'StreamingWriter already finished') after {@link finish}, for cells the bridge can deserialize.
   */
  writeRow(cells: StreamingCellInput[]): void {
    const inner = this.#live();
    try {
      inner.writeRow(JSON.stringify(cells));
    } catch (err) {
      throw ModernXlsxError.fromWasmError(err);
    }
  }

  /**
   * Finish writing and return the complete XLSX as a `Uint8Array`.
   *
   * This closes any open sheet, writes the shared string table and
   * metadata parts, and finalizes the ZIP archive.
   *
   * The writer is consumed -- calling any method other than `[Symbol.dispose]()` after `finish()` throws.
   *
   * @returns The complete XLSX file as a Uint8Array.
   * @throws {@link ModernXlsxError} with code `INVALID_WORKBOOK` if no sheet was started,
   * with code `INVALID_ARGUMENT` if the writer is already finished, or another code if it
   * fails to finalize.
   */
  finish(): Uint8Array {
    const inner = this.#live();
    this.#spent = true;
    try {
      return inner.finish();
    } catch (err) {
      throw ModernXlsxError.fromWasmError(err);
    }
  }
}
