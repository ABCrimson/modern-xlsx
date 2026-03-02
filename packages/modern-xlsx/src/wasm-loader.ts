import init, {
  initSync as _initSync,
  read as _wasmReadJson,
  writeBlob as _wasmWriteBlobJson,
  write as _wasmWriteJson,
  version as wasmVersion,
} from '../wasm/modern_xlsx_wasm.js';

import type { WorkbookData } from './types.js';

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
  if (typeof document !== 'undefined' && document.currentScript) {
    const src = (document.currentScript as HTMLScriptElement).src;
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
 * Call once at application startup. Idempotent — safe to call multiple times.
 * Accepts an optional URL/path to the `.wasm` binary for environments where
 * auto-detection doesn't work (e.g., custom CDN, Service Workers).
 *
 * @param wasmSource - URL, string path, or fetch Response for the WASM binary.
 *   If omitted, auto-detects: script src → import.meta.url → CDN fallback.
 */
export async function initWasm(wasmSource?: string | URL | Response): Promise<void> {
  if (initialized) return;
  if (!initPromise) {
    const source = wasmSource ?? detectWasmUrl();
    initPromise = init(source).then(
      () => {
        initialized = true;
      },
      (err) => {
        // Reset so callers can retry after a failed init.
        initPromise = null;
        throw err;
      },
    );
  }
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
}

/**
 * Auto-initialize WASM on first use. Wraps an async function to
 * transparently call `initWasm()` before the first operation.
 * Subsequent calls skip initialization (cached Promise pattern).
 */
export async function ensureReady(wasmSource?: string | URL): Promise<void> {
  if (!initialized) {
    await initWasm(wasmSource);
  }
}

export function ensureInitialized(): void {
  if (!initialized) {
    throw new Error('WASM not initialized. Call initWasm() first.');
  }
}

/**
 * Read XLSX bytes and return parsed WorkbookData.
 * WASM returns a JSON string; we use the V8-native JSON.parse()
 * which is 8-13x faster than serde_wasm_bindgen for large workbooks.
 */
export function wasmRead(data: Uint8Array): WorkbookData {
  const json = _wasmReadJson(data);
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

/**
 * Write WorkbookData to XLSX bytes.
 * Serializes to JSON string for transfer across the WASM boundary.
 */
export function wasmWrite(data: WorkbookData): Uint8Array {
  return _wasmWriteJson(JSON.stringify(data));
}

/**
 * Write WorkbookData to a Blob (browser).
 * Serializes to JSON string for transfer across the WASM boundary.
 */
export function wasmWriteBlob(data: WorkbookData): Blob {
  return _wasmWriteBlobJson(JSON.stringify(data));
}

export { wasmVersion };
