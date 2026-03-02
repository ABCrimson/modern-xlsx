import init, {
  read as _wasmReadJson,
  version as wasmVersion,
  write as _wasmWriteJson,
  writeBlob as _wasmWriteBlobJson,
} from '../wasm/modern_xlsx_wasm.js';

import type { WorkbookData } from './types.js';

let initPromise: Promise<void> | null = null;
let initialized = false;

export async function initWasm(): Promise<void> {
  if (initialized) return;
  if (!initPromise) {
    initPromise = init().then(
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
  return JSON.parse(json) as WorkbookData;
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
