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

export type WorkerRequest =
  | { id: number; type: 'init'; wasmUrl?: string }
  | { id: number; type: 'read'; data: Uint8Array; wasmUrl?: string; password?: string }
  | { id: number; type: 'write'; json: string; wasmUrl?: string; password?: string };

export type WorkerResponse =
  | { id: number; type: 'result'; data?: Uint8Array; json?: string }
  | { id: number; type: 'error'; error: string };

// Use globalThis directly — worker scripts run in DedicatedWorkerGlobalScope
// but we cast to avoid requiring lib.webworker.d.ts in the main tsconfig.
const _self = globalThis as unknown as {
  addEventListener(type: string, listener: (event: MessageEvent<WorkerRequest>) => void): void;
  postMessage(message: WorkerResponse, transfer?: Transferable[]): void;
};

_self.addEventListener('message', (event: MessageEvent<WorkerRequest>) => {
  const msg = event.data;
  handleMessage(msg).catch((err) => {
    const response: WorkerResponse = {
      id: msg.id,
      type: 'error',
      error: err instanceof Error ? err.message : String(err),
    };
    _self.postMessage(response);
  });
});

async function handleMessage(msg: WorkerRequest): Promise<void> {
  switch (msg.type) {
    case 'init': {
      await ensureInit(msg.wasmUrl);
      const response: WorkerResponse = { id: msg.id, type: 'result' };
      _self.postMessage(response);
      break;
    }

    case 'read': {
      await ensureInit(msg.wasmUrl);
      const resultJson = msg.password
        ? _wasmReadWithPasswordJson(msg.data, msg.password)
        : _wasmReadJson(msg.data);
      const response: WorkerResponse = { id: msg.id, type: 'result', json: resultJson };
      _self.postMessage(response);
      break;
    }

    case 'write': {
      await ensureInit(msg.wasmUrl);
      const resultData = msg.password
        ? _wasmWriteWithPasswordJson(msg.json, msg.password)
        : _wasmWriteJson(msg.json);
      const response: WorkerResponse = { id: msg.id, type: 'result', data: resultData };
      _self.postMessage(response, [resultData.buffer]);
      break;
    }
  }
}
