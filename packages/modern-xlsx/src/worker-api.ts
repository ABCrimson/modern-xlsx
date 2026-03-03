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
  readBuffer(data: Uint8Array): Promise<WorkbookData>;
  /** Write WorkbookData to XLSX bytes (runs in worker thread). */
  writeBuffer(data: WorkbookData): Promise<Uint8Array>;
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
  const pending = new Map<
    number,
    { resolve: (v: WorkerResponse) => void; reject: (e: Error) => void }
  >();

  worker.addEventListener('message', (event: MessageEvent<WorkerResponse>) => {
    const { id, type, error } = event.data;
    const handler = pending.get(id);
    if (!handler) return;
    pending.delete(id);
    if (type === 'error') {
      handler.reject(new Error(error ?? 'Unknown worker error'));
    } else {
      handler.resolve(event.data);
    }
  });

  worker.addEventListener('error', (event) => {
    for (const handler of pending.values()) {
      handler.reject(new Error(event.message ?? 'Worker error'));
    }
    pending.clear();
  });

  function send(
    request: Omit<WorkerRequest, 'id'>,
    transfer?: Transferable[],
  ): Promise<WorkerResponse> {
    return new Promise((resolve, reject) => {
      const id = nextId++;
      pending.set(id, { resolve, reject });
      const msg: WorkerRequest = { ...request, id, ...(wasmUrl != null && { wasmUrl }) };
      if (transfer) {
        worker.postMessage(msg, transfer);
      } else {
        worker.postMessage(msg);
      }
    });
  }

  return {
    async readBuffer(data: Uint8Array): Promise<WorkbookData> {
      // Transfer the buffer to the worker (zero-copy)
      const response = await send({ type: 'read', data }, [data.buffer]);
      if (!response.json) throw new Error('Worker returned no data');
      return JSON.parse(response.json) as WorkbookData;
    },

    async writeBuffer(data: WorkbookData): Promise<Uint8Array> {
      const json = JSON.stringify(data);
      const response = await send({ type: 'write', json });
      if (!response.data) throw new Error('Worker returned no data');
      return response.data;
    },

    terminate(): void {
      for (const handler of pending.values()) {
        handler.reject(new Error('Worker terminated'));
      }
      pending.clear();
      worker.terminate();
    },
  };
}
