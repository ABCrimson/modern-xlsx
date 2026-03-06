/**
 * Web Worker entry point for off-thread XLSX operations.
 *
 * Handles messages from the main thread to perform XLSX read/write
 * without blocking the UI. The WASM module is initialized inside the
 * worker on first use.
 *
 * Build as a separate entry: `dist/modern-xlsx.worker.js`
 */

import init, {
  read as _wasmReadJson,
  readWithPassword as _wasmReadWithPasswordJson,
  write as _wasmWriteJson,
  writeWithPassword as _wasmWriteWithPasswordJson,
} from '../wasm/modern_xlsx_wasm.js';

let initPromise: Promise<void> | null = null;

function ensureInit(wasmUrl?: string): Promise<void> {
  initPromise ??= init(wasmUrl ? new URL(wasmUrl) : undefined).then(
    () => {},
    (err) => {
      initPromise = null;
      throw err;
    },
  );
  return initPromise;
}

export interface WorkerRequest {
  id: number;
  type: 'init' | 'read' | 'write';
  wasmUrl?: string;
  data?: Uint8Array;
  json?: string;
  password?: string;
}

export interface WorkerResponse {
  id: number;
  type: 'result' | 'error';
  data?: Uint8Array;
  json?: string;
  error?: string;
}

// Use globalThis directly — worker scripts run in DedicatedWorkerGlobalScope
// but we cast to avoid requiring lib.webworker.d.ts in the main tsconfig.
const _self = globalThis as unknown as {
  addEventListener(type: string, listener: (event: MessageEvent<WorkerRequest>) => void): void;
  postMessage(message: WorkerResponse, transfer?: Transferable[]): void;
};

_self.addEventListener('message', (event: MessageEvent<WorkerRequest>) => {
  const { id, type, wasmUrl, data, json, password } = event.data;
  handleMessage(id, type, wasmUrl, data, json, password).catch((err) => {
    const response: WorkerResponse = {
      id,
      type: 'error',
      error: err instanceof Error ? err.message : String(err),
    };
    _self.postMessage(response);
  });
});

async function handleMessage(
  id: number,
  type: string,
  wasmUrl?: string,
  data?: Uint8Array,
  json?: string,
  password?: string,
): Promise<void> {
  switch (type) {
    case 'init': {
      await ensureInit(wasmUrl);
      const response: WorkerResponse = { id, type: 'result' };
      _self.postMessage(response);
      break;
    }

    case 'read': {
      await ensureInit(wasmUrl);
      if (!data) throw new Error('read requires data (Uint8Array)');
      const resultJson = password ? _wasmReadWithPasswordJson(data, password) : _wasmReadJson(data);
      const response: WorkerResponse = { id, type: 'result', json: resultJson };
      _self.postMessage(response);
      break;
    }

    case 'write': {
      await ensureInit(wasmUrl);
      if (!json) throw new Error('write requires json (WorkbookData)');
      const resultData = password
        ? _wasmWriteWithPasswordJson(json, password)
        : _wasmWriteJson(json);
      const response: WorkerResponse = { id, type: 'result', data: resultData };
      _self.postMessage(response, [resultData.buffer]);
      break;
    }

    default:
      throw new Error(`Unknown message type: ${type}`);
  }
}
