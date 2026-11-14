/**
 * The WASM backend shared by the full and lite builds.
 *
 * `wasm-loader.ts` and `wasm-loader-lite.ts` each pass their wasm-bindgen glue
 * module to {@link createLoader}; the entries bind the resulting
 * {@link WasmLoader} to their `Workbook` and `StreamingXlsxWriter` subclasses.
 * Nothing here imports a glue module, so the emitted declarations never
 * reference `../wasm/` or `../wasm-lite/`.
 */

import { loadDefaultWasm } from '#runtime';
import {
  INVALID_ARGUMENT,
  isError,
  ModernXlsxError,
  WASM_ERROR,
  WASM_INIT_FAILED,
} from './model/errors.js';
import type { RepairResult, ValidationReport, WorkbookData } from './model/types.js';

/**
 * Message of the `INVALID_ARGUMENT` error a password raises in the lite build.
 * Internal: not exported from the entries.
 */
export const LITE_PASSWORD_MESSAGE =
  "Password-protected workbooks need the full build: import from 'modern-xlsx', not 'modern-xlsx/lite'.";

/** What the async glue init accepts as `module_or_path` (wasm-bindgen `InitInput`). */
type WasmInitInput = RequestInfo | URL | Response | BufferSource | WebAssembly.Module;

/** What the sync glue init accepts as `module` (wasm-bindgen `SyncInitInput`). */
type WasmSyncInitInput = BufferSource | WebAssembly.Module;

/** An instance of the glue's `StreamingWriter` class. */
export interface StreamingWriterHandle {
  /** Release the Rust-side writer. A second call throws. */
  free(): void;
  [Symbol.dispose](): void;
  /** Close the open sheet, finalize the ZIP and return it; the writer is consumed. */
  finish(): Uint8Array;
  setStylesXml(xml: string): void;
  startSheet(name: string): void;
  /** Write one row given as a JSON array of streaming cells. */
  writeRow(cells_json: string): void;
}

/**
 * The subset of a wasm-bindgen glue module (`wasm/` or `wasm-lite/`
 * `modern_xlsx_wasm.js`) the loader uses, typed structurally after the glue's
 * `.d.ts`. The lite glue has no password functions.
 */
export interface WasmGlue {
  /** Async init; the loader passes the object form, or no argument so the glue resolves the binary via import.meta.url. */
  default(module_or_path?: {
    module_or_path: WasmInitInput | Promise<WasmInitInput>;
  }): Promise<unknown>;
  /** Sync init; the loader passes only the object form. */
  initSync(module: { module: WasmSyncInitInput }): unknown;
  read(data: Uint8Array | number[], limits?: string | null): string;
  readWithPassword?(data: Uint8Array | number[], password: string, limits?: string | null): string;
  write(json: string): Uint8Array;
  writeWithPassword?(json: string, password: string): Uint8Array;
  writeBlob(json: string): Blob;
  validate(json: string): string;
  repair(json: string): string;
  version(): string;
  StreamingWriter: new () => StreamingWriterHandle;
}

/** The loader one glue build gets from {@link createLoader}. Every member is a plain function. */
export interface WasmLoader {
  /**
   * Initialize the WASM module.
   *
   * Call once at application startup. Idempotent -- safe to call multiple times,
   * and concurrent calls share one initialization.
   *
   * @param wasmSource - The WASM binary as an http(s) or relative URL (a `URL` or string the
   *   wasm-bindgen glue fetches), or a fetch `Response`. Node.js-compatible runtimes cannot fetch
   *   a file path or `file:` URL: there, omit the argument or pass the bytes to `initWasmSync`.
   *   If omitted, the browser IIFE fetches `modern-xlsx.wasm` beside its
   *   `<script>` (or against the document base URL); Node.js-compatible runtimes
   *   read the package's own binary from disk; elsewhere wasm-bindgen resolves the
   *   binary next to the glue via `import.meta.url`.
   * @throws ModernXlsxError with code `WASM_INIT_FAILED` (the original error as
   *   `cause`) if the binary cannot be read, fetched or instantiated; a later call retries.
   *
   * @example
   * ```ts
   * import { initWasm } from 'modern-xlsx';
   * await initWasm();
   * ```
   */
  initWasm(wasmSource?: string | URL | Response): Promise<void>;
  /**
   * Initialize WASM synchronously from a pre-loaded buffer.
   * Useful in Node.js test environments or when WASM bytes are already available.
   *
   * @param module - A WebAssembly.Module or raw bytes (Uint8Array/ArrayBuffer).
   */
  initWasmSync(module: WebAssembly.Module | BufferSource): void;
  /**
   * Initialize WASM if it is not initialized yet; a no-op afterwards.
   *
   * @param wasmSource - An http(s) or relative URL of the `.wasm` binary (a `URL` or string); on
   *   Node.js-compatible runtimes omit it. Unlike {@link WasmLoader.initWasm} it takes no
   *   `Response`. If omitted, the default resolution
   *   of `initWasm()` applies. Ignored when WASM is already initialized.
   *
   * @example
   * ```ts
   * await ensureReady();
   * const wb = await readBuffer(data);
   * ```
   */
  ensureReady(wasmSource?: string | URL): Promise<void>;
  /**
   * Guard that throws if WASM has not been initialized.
   *
   * @throws ModernXlsxError with code `WASM_INIT_FAILED` if `initWasm()` has not been called.
   */
  ensureInitialized(): void;
  /**
   * Read XLSX bytes and return parsed WorkbookData.
   * WASM returns a JSON string, which is parsed with JSON.parse().
   *
   * @throws ModernXlsxError with code `WASM_ERROR` if WASM returns no WorkbookData.
   */
  wasmRead(data: Uint8Array, limitsJson?: string): WorkbookData;
  /**
   * Read encrypted XLSX bytes with a password and return parsed WorkbookData.
   * If the file is not encrypted, the password is ignored.
   *
   * @throws ModernXlsxError with code `INVALID_ARGUMENT` in the lite build.
   */
  wasmReadWithPassword(data: Uint8Array, password: string, limitsJson?: string): WorkbookData;
  /** Write WorkbookData to XLSX bytes (serialized to a JSON string for the WASM boundary). */
  wasmWrite(data: WorkbookData): Uint8Array<ArrayBuffer>;
  /**
   * Write WorkbookData to an encrypted XLSX (OLE2 container with Agile Encryption).
   *
   * @throws ModernXlsxError with code `INVALID_ARGUMENT` in the lite build.
   */
  wasmWriteWithPassword(data: WorkbookData, password: string): Uint8Array<ArrayBuffer>;
  /** Write WorkbookData to a Blob (browser). */
  wasmWriteBlob(data: WorkbookData): Blob;
  /**
   * Validate a workbook and return a structured report.
   *
   * @throws ModernXlsxError with code `WASM_ERROR` if WASM returns no ValidationReport.
   */
  wasmValidate(data: WorkbookData): ValidationReport;
  /**
   * Validate and auto-repair a workbook. Returns the repaired workbook,
   * a post-repair validation report, and the number of repairs applied.
   *
   * @throws ModernXlsxError with code `WASM_ERROR` if WASM returns no RepairResult.
   */
  wasmRepair(data: WorkbookData): RepairResult;
  /** The WASM crate version. */
  version(): string;
  /** The glue's streaming writer class. */
  readonly StreamingWriter: new () => StreamingWriterHandle;
  /**
   * Set the URL `initWasm()` uses when called without a source. The first call
   * wins; later calls are ignored. Internal: only the browser IIFE calls it, and
   * the entries do not export it.
   */
  setDefaultWasmUrl(url: URL): void;
}

/** Lightweight structural check for WorkbookData crossing the WASM or worker boundary. */
export function isWorkbookData(v: unknown): v is WorkbookData {
  if (typeof v !== 'object' || v === null) return false;
  if (!('sheets' in v) || !('styles' in v)) return false;
  return Array.isArray(v.sheets) && typeof v.styles === 'object' && v.styles !== null;
}

/** Lightweight structural check for ValidationReport from WASM. */
function isValidationReport(v: unknown): v is ValidationReport {
  if (typeof v !== 'object' || v === null) return false;
  return 'issues' in v && 'isValid' in v && Array.isArray(v.issues);
}

/** Lightweight structural check for RepairResult from WASM. */
function isRepairResult(v: unknown): v is RepairResult {
  if (typeof v !== 'object' || v === null) return false;
  return 'workbook' in v && 'report' in v && 'repairCount' in v;
}

/**
 * Run a glue call, turning its `"[CODE] message"` errors into `ModernXlsxError`
 * (falling back to `WASM_ERROR` when the message has no code).
 */
function callWasm<T>(call: () => T): T {
  try {
    return call();
  } catch (err) {
    throw ModernXlsxError.fromWasmError(err);
  }
}

/** Parse a JSON string returned by WASM and check its shape. */
function parseChecked<T>(json: string, guard: (v: unknown) => v is T, name: string): T {
  const parsed: unknown = JSON.parse(json);
  if (!guard(parsed)) {
    throw new ModernXlsxError(WASM_ERROR, `WASM returned invalid ${name} structure`);
  }
  return parsed;
}

/**
 * Build the loader for one glue build. The initialization state lives in this
 * closure, so each build initializes independently and the returned functions
 * can be called unbound.
 *
 * @param glue - The namespace of one glue module (`import * as glue` in a loader module).
 * @param kind - `'lite'` for the build without encryption: its password calls
 *   throw `INVALID_ARGUMENT` with {@link LITE_PASSWORD_MESSAGE}.
 */
export function createLoader(glue: WasmGlue, kind: 'full' | 'lite'): WasmLoader {
  let initPromise: Promise<void> | null = null;
  let initialized = false;
  let defaultWasmUrl: URL | undefined;

  function passwordCall<F>(fn: F | undefined): F {
    if (kind === 'lite' || fn === undefined) {
      throw new ModernXlsxError(INVALID_ARGUMENT, LITE_PASSWORD_MESSAGE);
    }
    return fn;
  }

  function ensureInitialized(): void {
    if (!initialized) {
      throw new ModernXlsxError(WASM_INIT_FAILED, 'WASM not initialized. Call initWasm() first.');
    }
  }

  async function initWasm(wasmSource?: string | URL | Response): Promise<void> {
    if (initialized) return;
    initPromise ??= (async () => {
      const moduleOrPath = wasmSource ?? defaultWasmUrl ?? (await loadDefaultWasm(kind));
      await glue.default(moduleOrPath === undefined ? undefined : { module_or_path: moduleOrPath });
      initialized = true;
    })().catch((err: unknown) => {
      initPromise = null;
      throw new ModernXlsxError(WASM_INIT_FAILED, isError(err) ? err.message : String(err), {
        cause: err,
      });
    });
    return initPromise;
  }

  return {
    initWasm,
    initWasmSync(module: WebAssembly.Module | BufferSource): void {
      if (initialized) return;
      glue.initSync({ module });
      initialized = true;
      initPromise = Promise.resolve(); // prevent concurrent initWasm from re-initializing
    },
    async ensureReady(wasmSource?: string | URL): Promise<void> {
      if (!initialized) {
        await initWasm(wasmSource);
      }
    },
    ensureInitialized,
    wasmRead(data: Uint8Array, limitsJson?: string): WorkbookData {
      const json = callWasm(() => glue.read(data, limitsJson));
      return parseChecked(json, isWorkbookData, 'WorkbookData');
    },
    wasmReadWithPassword(data: Uint8Array, password: string, limitsJson?: string): WorkbookData {
      const readWithPassword = passwordCall(glue.readWithPassword);
      const json = callWasm(() => readWithPassword(data, password, limitsJson));
      return parseChecked(json, isWorkbookData, 'WorkbookData');
    },
    // The WASM writers return `new Uint8Array(len)` filled from WASM memory, so the bytes
    // own a plain ArrayBuffer; the generated glue types them as Uint8Array<ArrayBufferLike>.
    wasmWrite(data: WorkbookData): Uint8Array<ArrayBuffer> {
      return callWasm(() => glue.write(JSON.stringify(data))) as Uint8Array<ArrayBuffer>;
    },
    wasmWriteWithPassword(data: WorkbookData, password: string): Uint8Array<ArrayBuffer> {
      const writeWithPassword = passwordCall(glue.writeWithPassword);
      return callWasm(() =>
        writeWithPassword(JSON.stringify(data), password),
      ) as Uint8Array<ArrayBuffer>;
    },
    wasmWriteBlob(data: WorkbookData): Blob {
      return callWasm(() => glue.writeBlob(JSON.stringify(data)));
    },
    wasmValidate(data: WorkbookData): ValidationReport {
      const json = callWasm(() => glue.validate(JSON.stringify(data)));
      return parseChecked(json, isValidationReport, 'ValidationReport');
    },
    wasmRepair(data: WorkbookData): RepairResult {
      const json = callWasm(() => glue.repair(JSON.stringify(data)));
      return parseChecked(json, isRepairResult, 'RepairResult');
    },
    version(): string {
      return glue.version();
    },
    StreamingWriter: glue.StreamingWriter,
    setDefaultWasmUrl(url: URL): void {
      defaultWasmUrl ??= url;
    },
  };
}
