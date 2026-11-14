/**
 * Client-side API for off-thread XLSX operations via Web Workers.
 *
 * Usage:
 * ```ts
 * using worker = createXlsxWorker({ workerUrl: '/modern-xlsx.worker.js' });
 * const data = await worker.readBuffer(bytes);
 * const out = await worker.writeBuffer(data);
 * ```
 */

import { isError, ModernXlsxError, WASM_ERROR, WORKER_ERROR } from './model/errors.js';
import type { WorkbookData } from './model/types.js';
import { isWorkbookData } from './wasm-backend.js';
import type { WorkerRequest, WorkerResponse } from './worker.js';

/** Options for {@link createXlsxWorker}: where the worker script and the WASM binary live. */
export interface XlsxWorkerOptions {
  /** URL to the worker script (modern-xlsx.worker.js). */
  workerUrl: string | URL;
  /**
   * URL to the WASM binary. A relative URL is resolved against the page
   * (`document.baseURI`, else `location.href`), not against the worker script.
   * If omitted, the worker loads `modern-xlsx.wasm` next to itself.
   */
  wasmUrl?: string | URL;
}

/** Client for the XLSX worker: reads and writes workbooks on a worker thread. */
export interface XlsxWorker {
  /**
   * Read XLSX bytes into parsed WorkbookData (runs in worker thread).
   *
   * When `data` spans its whole `ArrayBuffer`, that buffer is transferred to the
   * worker and detached (`data` becomes unusable). When `data` is a view on part
   * of a buffer, or on a `SharedArrayBuffer`, it is copied instead and the
   * caller's buffer is left intact.
   */
  readBuffer(data: Uint8Array, options?: { password?: string }): Promise<WorkbookData>;
  /**
   * Write WorkbookData to XLSX bytes (runs in worker thread).
   *
   * @throws ModernXlsxError with code `INVALID_WORKBOOK` for the causes listed with that
   *   code in `errors.ts`, and `WORKER_ERROR` when the worker crashed, was terminated or
   *   returned no data.
   */
  writeBuffer(data: WorkbookData, options?: { password?: string }): Promise<Uint8Array>;
  /** Terminate the worker. Pending calls reject with `WORKER_ERROR`. */
  terminate(): void;
  /** Alias for {@link terminate}, so `using worker = createXlsxWorker(...)` works. */
  [Symbol.dispose](): void;
}

/**
 * Create an XLSX Web Worker for off-thread XLSX operations.
 * All WASM operations run in the worker — the main thread stays responsive.
 */
export function createXlsxWorker(options: XlsxWorkerOptions): XlsxWorker {
  const worker = new Worker(options.workerUrl, { type: 'module' });
  const base = globalThis.document?.baseURI ?? globalThis.location?.href;
  const wasmUrl =
    options.wasmUrl === undefined
      ? undefined
      : base
        ? new URL(options.wasmUrl, base).href
        : String(options.wasmUrl);
  let nextId = 0;
  let terminated = false;
  const pending = new Map<
    number,
    { resolve: (v: WorkerResponse) => void; reject: (e: Error) => void }
  >();

  worker.addEventListener('message', (event: MessageEvent<WorkerResponse>) => {
    const resp = event.data;
    const handler = pending.get(resp.id);
    if (!handler) return;
    pending.delete(resp.id);
    if (resp.type === 'error') {
      handler.reject(ModernXlsxError.fromWasmError(resp.error));
    } else {
      handler.resolve(resp);
    }
  });

  worker.addEventListener('error', (event) => {
    terminated = true;
    for (const handler of pending.values()) {
      handler.reject(new ModernXlsxError(WORKER_ERROR, event.message ?? 'Worker error'));
    }
    pending.clear();
  });

  worker.addEventListener('messageerror', () => {
    for (const handler of pending.values()) {
      handler.reject(new ModernXlsxError(WORKER_ERROR, 'Worker message could not be deserialized'));
    }
    pending.clear();
  });

  function send(request: WorkerRequest, transfer?: Transferable[]): Promise<WorkerResponse> {
    if (terminated) {
      return Promise.reject(
        new ModernXlsxError(WORKER_ERROR, 'Worker has crashed or been terminated'),
      );
    }
    const { promise, resolve, reject } = Promise.withResolvers<WorkerResponse>();
    pending.set(request.id, { resolve, reject });
    try {
      if (transfer) {
        worker.postMessage(request, transfer);
      } else {
        worker.postMessage(request);
      }
    } catch (err) {
      pending.delete(request.id);
      throw new ModernXlsxError(WORKER_ERROR, isError(err) ? err.message : String(err), {
        cause: err,
      });
    }
    return promise;
  }

  function terminate(): void {
    terminated = true;
    for (const handler of pending.values()) {
      handler.reject(new ModernXlsxError(WORKER_ERROR, 'Worker terminated'));
    }
    pending.clear();
    worker.terminate();
  }

  return {
    async readBuffer(data: Uint8Array, options?: { password?: string }): Promise<WorkbookData> {
      const id = nextId++;
      // Transfer zero-copy only when the view spans a whole ArrayBuffer; a
      // partial view or a SharedArrayBuffer view is copied so the caller's
      // buffer is never detached.
      const exact =
        data.buffer instanceof ArrayBuffer &&
        data.byteOffset === 0 &&
        data.byteLength === data.buffer.byteLength;
      const payload = exact ? data : data.slice();
      const msg: WorkerRequest = { id, type: 'read', data: payload };
      if (wasmUrl != null) msg.wasmUrl = wasmUrl;
      if (options?.password != null) msg.password = options.password;
      const response = await send(msg, [payload.buffer]);
      if (response.type !== 'result' || !response.json) {
        throw new ModernXlsxError(WORKER_ERROR, 'Worker returned no data');
      }
      const parsed: unknown = JSON.parse(response.json);
      if (!isWorkbookData(parsed)) {
        throw new ModernXlsxError(WASM_ERROR, 'Worker returned an invalid WorkbookData structure');
      }
      return parsed;
    },

    async writeBuffer(data: WorkbookData, options?: { password?: string }): Promise<Uint8Array> {
      const id = nextId++;
      const json = JSON.stringify(data);
      const msg: WorkerRequest = { id, type: 'write', json };
      if (wasmUrl != null) msg.wasmUrl = wasmUrl;
      if (options?.password != null) msg.password = options.password;
      const response = await send(msg);
      if (response.type !== 'result' || !response.data) {
        throw new ModernXlsxError(WORKER_ERROR, 'Worker returned no data');
      }
      return response.data;
    },

    terminate,
    [Symbol.dispose]: terminate,
  };
}
