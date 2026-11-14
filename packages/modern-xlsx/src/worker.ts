/**
 * Web Worker entry point for off-thread XLSX operations.
 *
 * Handles messages from the main thread to perform XLSX read/write
 * without blocking the UI. The WASM module is initialized inside the
 * worker on first use.
 *
 * Built as a separate entry, `dist/modern-xlsx.worker.js` (see `tsdown.config.ts`).
 */

import init, {
  read as _wasmReadJson,
  readWithPassword as _wasmReadWithPasswordJson,
  write as _wasmWriteJson,
  writeWithPassword as _wasmWriteWithPasswordJson,
} from '../wasm/modern_xlsx_wasm.js';
import { isError, ModernXlsxError, WASM_INIT_FAILED } from './model/errors.js';

let initPromise: Promise<void> | null = null;

function ensureInit(wasmUrl?: string): Promise<void> {
  initPromise ??= init({
    module_or_path: wasmUrl
      ? new URL(wasmUrl, globalThis.location.href)
      : new URL('modern-xlsx.wasm', import.meta.url),
  }).then(
    () => {},
    (err) => {
      initPromise = null;
      throw new ModernXlsxError(WASM_INIT_FAILED, isError(err) ? err.message : String(err), {
        cause: err,
      });
    },
  );
  return initPromise;
}

/** A message from the main thread: `id` pairs it with its {@link WorkerResponse}. */
export type WorkerRequest =
  | { id: number; type: 'init'; wasmUrl?: string }
  | { id: number; type: 'read'; data: Uint8Array; wasmUrl?: string; password?: string }
  | { id: number; type: 'write'; json: string; wasmUrl?: string; password?: string };

/** The worker's reply to the request with the same `id`: a result or an error message. */
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
      error:
        err instanceof ModernXlsxError
          ? `[${err.code}] ${err.message}`
          : isError(err)
            ? err.message
            : String(err),
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
