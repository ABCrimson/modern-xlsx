import init, {
  read as wasmRead,
  version as wasmVersion,
  write as wasmWrite,
} from '../wasm/ironsheet_wasm.js';

let initPromise: Promise<void> | null = null;
let initialized = false;

export async function initWasm(): Promise<void> {
  if (initialized) return;
  if (!initPromise) {
    initPromise = init().then(() => {
      initialized = true;
    });
  }
  return initPromise;
}

export function ensureInitialized(): void {
  if (!initialized) {
    throw new Error('WASM not initialized. Call initWasm() first.');
  }
}

export { wasmRead, wasmWrite, wasmVersion };
