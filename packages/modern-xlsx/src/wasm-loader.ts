import init, {
  initSync as _initSync,
  read as _wasmReadJson,
  readWithPassword as _wasmReadWithPasswordJson,
  repair as _wasmRepairJson,
  validate as _wasmValidateJson,
  writeBlob as _wasmWriteBlobJson,
  write as _wasmWriteJson,
  writeWithPassword as _wasmWriteWithPasswordJson,
  version as wasmVersion,
} from '../wasm/modern_xlsx_wasm.js';

import { ModernXlsxError, WASM_INIT_FAILED } from './errors.js';
import type { RepairResult, ValidationReport, WorkbookData } from './types.js';

/**
 * Re-throw a WASM error as a `ModernXlsxError` with a parsed error code.
 *
 * WASM errors arrive as `Error` objects whose `message` follows the
 * `"[CODE] human-readable message"` format. This helper extracts the code
 * and message, falling back to `WASM_ERROR` if the format doesn't match.
 */
function rethrowWasm(err: unknown): never {
  throw ModernXlsxError.fromWasmError(err);
}

let initPromise: Promise<void> | null = null;
let initialized = false;

/**
 * Detect the WASM binary URL from the current environment.
 *
 * - IIFE/script tag: derives URL relative to `document.currentScript.src`
 * - CDN: constructs versioned CDN URL
 * - ESM: returns `undefined` to let wasm-bindgen use `import.meta.url`
 */
function detectWasmUrl(): string | URL | undefined {
  // In a <script> tag context (IIFE bundle), derive from script src
  if (typeof document !== 'undefined' && document.currentScript instanceof HTMLScriptElement) {
    const { src } = document.currentScript;
    if (src) {
      return new URL('modern-xlsx.wasm', src);
    }
  }
  // Let wasm-bindgen use its default import.meta.url resolution
  return undefined;
}

/**
 * Initialize the WASM module.
 *
 * Call once at application startup. Idempotent -- safe to call multiple times.
 * Accepts an optional URL/path to the `.wasm` binary for environments where
 * auto-detection doesn't work (e.g., custom CDN, Service Workers).
 *
 * @param wasmSource - URL, string path, or fetch Response for the WASM binary.
 *   If omitted, auto-detects: script src, import.meta.url, or CDN fallback.
 *
 * @example
 * ```ts
 * import { initWasm } from 'modern-xlsx';
 * await initWasm();
 * ```
 */
export async function initWasm(wasmSource?: string | URL | Response): Promise<void> {
  if (initialized) return;
  initPromise ??= init(wasmSource ?? detectWasmUrl()).then(
    () => {
      initialized = true;
    },
    (err) => {
      initPromise = null;
      throw err;
    },
  );
  return initPromise;
}

/**
 * Initialize WASM synchronously from a pre-loaded buffer.
 * Useful in Node.js test environments or when WASM bytes are already available.
 *
 * @param module - A WebAssembly.Module or raw bytes (Uint8Array/ArrayBuffer).
 */
export function initWasmSync(module: WebAssembly.Module | BufferSource): void {
  if (initialized) return;
  _initSync({ module });
  initialized = true;
  initPromise = Promise.resolve(); // prevent concurrent initWasm from re-initializing
}

/**
 * Auto-initialize WASM on first use. Wraps an async function to
 * transparently call `initWasm()` before the first operation.
 * Subsequent calls skip initialization (cached Promise pattern).
 *
 * @param wasmSource - Optional URL/path to the `.wasm` binary.
 *
 * @example
 * ```ts
 * await ensureReady();
 * const wb = await readBuffer(data);
 * ```
 */
export async function ensureReady(wasmSource?: string | URL): Promise<void> {
  if (!initialized) {
    await initWasm(wasmSource);
  }
}

/**
 * Guard that throws if WASM has not been initialized.
 *
 * @throws ModernXlsxError with code `WASM_INIT_FAILED` if `initWasm()` has not been called.
 */
export function ensureInitialized(): void {
  if (!initialized) {
    throw new ModernXlsxError(WASM_INIT_FAILED, 'WASM not initialized. Call initWasm() first.');
  }
}

/**
 * Read XLSX bytes and return parsed WorkbookData.
 * WASM returns a JSON string; we use the V8-native JSON.parse()
 * which is 8-13x faster than serde_wasm_bindgen for large workbooks.
 */
export function wasmRead(data: Uint8Array): WorkbookData {
  let json: string;
  try {
    json = _wasmReadJson(data);
  } catch (err) {
    rethrowWasm(err);
  }
  const parsed: unknown = JSON.parse(json);
  if (!isWorkbookData(parsed)) {
    throw new Error('WASM returned invalid WorkbookData structure');
  }
  return parsed;
}

/**
 * Read encrypted XLSX bytes with a password and return parsed WorkbookData.
 * If the file is not encrypted, the password is ignored.
 */
export function wasmReadWithPassword(data: Uint8Array, password: string): WorkbookData {
  let json: string;
  try {
    json = _wasmReadWithPasswordJson(data, password);
  } catch (err) {
    rethrowWasm(err);
  }
  const parsed: unknown = JSON.parse(json);
  if (!isWorkbookData(parsed)) {
    throw new Error('WASM returned invalid WorkbookData structure');
  }
  return parsed;
}

/** Lightweight structural check for the WASM boundary. */
function isWorkbookData(v: unknown): v is WorkbookData {
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
 * Write WorkbookData to XLSX bytes.
 * Serializes to JSON string for transfer across the WASM boundary.
 */
export function wasmWrite(data: WorkbookData): Uint8Array {
  try {
    return _wasmWriteJson(JSON.stringify(data));
  } catch (err) {
    rethrowWasm(err);
  }
}

/**
 * Write WorkbookData to an encrypted XLSX (OLE2 container with Agile Encryption).
 * Serializes to JSON string for transfer across the WASM boundary.
 */
export function wasmWriteWithPassword(data: WorkbookData, password: string): Uint8Array {
  try {
    return _wasmWriteWithPasswordJson(JSON.stringify(data), password);
  } catch (err) {
    rethrowWasm(err);
  }
}

/**
 * Write WorkbookData to a Blob (browser).
 * Serializes to JSON string for transfer across the WASM boundary.
 */
export function wasmWriteBlob(data: WorkbookData): Blob {
  try {
    return _wasmWriteBlobJson(JSON.stringify(data));
  } catch (err) {
    rethrowWasm(err);
  }
}

/**
 * Validate a workbook and return a structured report.
 * Uses WASM-accelerated validation for structural compliance checking.
 */
export function wasmValidate(data: WorkbookData): ValidationReport {
  let json: string;
  try {
    json = _wasmValidateJson(JSON.stringify(data));
  } catch (err) {
    rethrowWasm(err);
  }
  const parsed: unknown = JSON.parse(json);
  if (!isValidationReport(parsed)) {
    throw new Error('WASM returned invalid ValidationReport structure');
  }
  return parsed;
}

/**
 * Validate and auto-repair a workbook. Returns the repaired workbook,
 * a post-repair validation report, and the number of repairs applied.
 */
export function wasmRepair(data: WorkbookData): RepairResult {
  let json: string;
  try {
    json = _wasmRepairJson(JSON.stringify(data));
  } catch (err) {
    rethrowWasm(err);
  }
  const parsed: unknown = JSON.parse(json);
  if (!isRepairResult(parsed)) {
    throw new Error('WASM returned invalid RepairResult structure');
  }
  return parsed;
}

export { wasmVersion };
