/**
 * Client-side API for off-thread XLSX operations via Web Workers.
 *
 * Usage:
 * ```typescript
 * const worker = createXlsxWorker({ workerUrl: '/modern-xlsx.worker.js' });
 * const wb = await worker.readBuffer(data);
 * const buffer = await worker.writeBuffer(wb.toJSON());
 * worker.terminate();
 * ```
 */

import type { WorkbookData } from './types.js';
import type { WorkerRequest, WorkerResponse } from './worker.js';

export interface XlsxWorkerOptions {
  /** URL to the worker script (modern-xlsx.worker.js). */
  workerUrl: string | URL;
  /** URL to the WASM binary. If omitted, the worker auto-detects. */
  wasmUrl?: string | URL;
}

export interface XlsxWorker {
  /** Read XLSX bytes into parsed WorkbookData (runs in worker thread). */
  readBuffer(data: Uint8Array, options?: { password?: string }): Promise<WorkbookData>;
  /** Write WorkbookData to XLSX bytes (runs in worker thread). */
  writeBuffer(data: WorkbookData, options?: { password?: string }): Promise<Uint8Array>;
  /** Terminate the worker. */
  terminate(): void;
}

/**
 * Create an XLSX Web Worker for off-thread XLSX operations.
 * All WASM operations run in the worker — the main thread stays responsive.
 */
export function createXlsxWorker(options: XlsxWorkerOptions): XlsxWorker {
  const worker = new Worker(options.workerUrl, { type: 'module' });
  const wasmUrl = options.wasmUrl?.toString();
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
      handler.reject(new Error(resp.error));
    } else {
      handler.resolve(resp);
    }
  });

  worker.addEventListener('error', (event) => {
    terminated = true;
    for (const handler of pending.values()) {
      handler.reject(new Error(event.message ?? 'Worker error'));
    }
    pending.clear();
  });

  function send(request: WorkerRequest, transfer?: Transferable[]): Promise<WorkerResponse> {
    if (terminated) {
      return Promise.reject(new Error('Worker has crashed or been terminated'));
    }
    const { promise, resolve, reject } = Promise.withResolvers<WorkerResponse>();
    pending.set(request.id, { resolve, reject });
    if (transfer) {
      worker.postMessage(request, transfer);
    } else {
      worker.postMessage(request);
    }
    return promise;
  }

  return {
    async readBuffer(data: Uint8Array, options?: { password?: string }): Promise<WorkbookData> {
      const id = nextId++;
      // Transfer the buffer to the worker (zero-copy)
      const msg: WorkerRequest = { id, type: 'read', data };
      if (wasmUrl != null) msg.wasmUrl = wasmUrl;
      if (options?.password != null) msg.password = options.password;
      const response = await send(msg, [data.buffer]);
      if (response.type !== 'result' || !response.json) throw new Error('Worker returned no data');
      return JSON.parse(response.json) as WorkbookData;
    },

    async writeBuffer(data: WorkbookData, options?: { password?: string }): Promise<Uint8Array> {
      const id = nextId++;
      const json = JSON.stringify(data);
      const msg: WorkerRequest = { id, type: 'write', json };
      if (wasmUrl != null) msg.wasmUrl = wasmUrl;
      if (options?.password != null) msg.password = options.password;
      const response = await send(msg);
      if (response.type !== 'result' || !response.data) throw new Error('Worker returned no data');
      return response.data;
    },

    terminate(): void {
      terminated = true;
      for (const handler of pending.values()) {
        handler.reject(new Error('Worker terminated'));
      }
      pending.clear();
      worker.terminate();
    },
  };
}
