import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { ModernXlsxError } from '../src/model/errors.js';
import type { WorkbookData } from '../src/model/types.js';
import { createXlsxWorker } from '../src/worker-api.js';

interface Posted {
  msg: { id: number; type: string; wasmUrl?: string; data?: Uint8Array; json?: string };
  transfer: Transferable[] | undefined;
}

class FakeWorker {
  static last: FakeWorker;
  posted: Posted[] = [];
  terminated = false;
  throwOnPost = false;
  #listeners = new Map<string, ((event: unknown) => void)[]>();

  constructor() {
    FakeWorker.last = this;
  }

  addEventListener(type: string, fn: (event: unknown) => void): void {
    const list = this.#listeners.get(type) ?? [];
    list.push(fn);
    this.#listeners.set(type, list);
  }

  postMessage(msg: Posted['msg'], transfer?: Transferable[]): void {
    if (this.throwOnPost) throw new Error('DataCloneError');
    if (transfer) structuredClone(msg, { transfer });
    this.posted.push({ msg, transfer });
  }

  terminate(): void {
    this.terminated = true;
  }

  emit(type: string, event: unknown): void {
    for (const fn of this.#listeners.get(type) ?? []) fn(event);
  }

  reply(data: unknown): void {
    this.emit('message', { data });
  }
}

const MINIMAL = JSON.stringify({ sheets: [], styles: {} });

beforeEach(() => {
  vi.stubGlobal('Worker', FakeWorker);
});

afterEach(() => {
  vi.unstubAllGlobals();
});

describe('createXlsxWorker hardening', () => {
  it('maps a coded error reply to ModernXlsxError', async () => {
    const w = createXlsxWorker({ workerUrl: 'w.js' });
    const p = w.readBuffer(new Uint8Array(4));
    FakeWorker.last.reply({
      id: 0,
      type: 'error',
      error: '[PASSWORD_PROTECTED] Incorrect password.',
    });
    const err = await p.catch((e: unknown) => e);
    expect(err).toBeInstanceOf(ModernXlsxError);
    expect((err as ModernXlsxError).code).toBe('PASSWORD_PROTECTED');
  });

  it('rejects every in-flight call on messageerror and keeps working', async () => {
    const w = createXlsxWorker({ workerUrl: 'w.js' });
    const a = w.readBuffer(new Uint8Array(4)).catch((e: unknown) => e);
    const b = w.readBuffer(new Uint8Array(4)).catch((e: unknown) => e);
    FakeWorker.last.emit('messageerror', {});
    for (const err of await Promise.all([a, b])) {
      expect((err as ModernXlsxError).code).toBe('WORKER_ERROR');
    }
    const c = w.readBuffer(new Uint8Array(4));
    FakeWorker.last.reply({ id: 2, type: 'result', json: MINIMAL });
    await expect(c).resolves.toHaveProperty('sheets');
  });

  it('rejects a result that is not WorkbookData with WASM_ERROR', async () => {
    const w = createXlsxWorker({ workerUrl: 'w.js' });
    const p = w.readBuffer(new Uint8Array(4)).catch((e: unknown) => e);
    FakeWorker.last.reply({ id: 0, type: 'result', json: '{}' });
    expect(((await p) as ModernXlsxError).code).toBe('WASM_ERROR');
  });

  it('does not detach the parent buffer of a subarray', async () => {
    const w = createXlsxWorker({ workerUrl: 'w.js' });
    const big = new Uint8Array(16);
    const p = w.readBuffer(big.subarray(4));
    expect(big.buffer.detached).toBe(false);
    FakeWorker.last.reply({ id: 0, type: 'result', json: MINIMAL });
    await p;
  });

  it('transfers an exact input and detaches it', async () => {
    const w = createXlsxWorker({ workerUrl: 'w.js' });
    const exact = new Uint8Array(16);
    const p = w.readBuffer(exact);
    expect(exact.buffer.detached).toBe(true);
    FakeWorker.last.reply({ id: 0, type: 'result', json: MINIMAL });
    await p;
  });

  it('copies a SharedArrayBuffer view instead of transferring it', async () => {
    const w = createXlsxWorker({ workerUrl: 'w.js' });
    const shared = new Uint8Array(new SharedArrayBuffer(8));
    const p = w.readBuffer(shared);
    const { transfer } = FakeWorker.last.posted[0] as Posted;
    expect(transfer?.[0]).not.toBe(shared.buffer);
    FakeWorker.last.reply({ id: 0, type: 'result', json: MINIMAL });
    await p;
  });

  it('clears the pending entry when postMessage throws', async () => {
    const w = createXlsxWorker({ workerUrl: 'w.js' });
    FakeWorker.last.throwOnPost = true;
    await expect(w.readBuffer(new Uint8Array(4))).rejects.toMatchObject({ code: 'WORKER_ERROR' });
    FakeWorker.last.throwOnPost = false;
    const onUnhandled = vi.fn();
    process.on('unhandledRejection', onUnhandled);
    try {
      w.terminate();
      await new Promise((r) => setTimeout(r, 0));
      expect(onUnhandled).not.toHaveBeenCalled();
    } finally {
      process.off('unhandledRejection', onUnhandled);
    }
  });

  it('rejects calls after terminate with WORKER_ERROR', async () => {
    const w = createXlsxWorker({ workerUrl: 'w.js' });
    w.terminate();
    const err = await w.readBuffer(new Uint8Array(4)).catch((e: unknown) => e);
    expect((err as ModernXlsxError).code).toBe('WORKER_ERROR');
  });

  it('terminates the worker when disposed with using', () => {
    {
      using _w = createXlsxWorker({ workerUrl: 'w.js' });
    }
    expect(FakeWorker.last.terminated).toBe(true);
  });

  it('rejects pending calls on a worker error event and marks the worker terminated', async () => {
    const w = createXlsxWorker({ workerUrl: 'w.js' });
    const p = w.readBuffer(new Uint8Array(4)).catch((e: unknown) => e);
    FakeWorker.last.emit('error', { message: 'boom' });
    expect(await p).toMatchObject({ code: 'WORKER_ERROR', message: 'boom' });
    await expect(w.readBuffer(new Uint8Array(4))).rejects.toMatchObject({
      code: 'WORKER_ERROR',
      message: 'Worker has crashed or been terminated',
    });
  });

  it('rejects an in-flight call when terminate() is called', async () => {
    const w = createXlsxWorker({ workerUrl: 'w.js' });
    const p = w.readBuffer(new Uint8Array(4)).catch((e: unknown) => e);
    w.terminate();
    expect(await p).toMatchObject({ code: 'WORKER_ERROR', message: 'Worker terminated' });
  });

  it('writeBuffer posts the JSON and returns the reply bytes', async () => {
    const w = createXlsxWorker({ workerUrl: 'w.js' });
    const data = { sheets: [], styles: {} } as unknown as WorkbookData;
    const first = w.writeBuffer(data);
    expect((FakeWorker.last.posted[0] as Posted).msg).toEqual({
      id: 0,
      type: 'write',
      json: JSON.stringify(data),
    });
    FakeWorker.last.reply({ id: 0, type: 'result' });
    await expect(first).rejects.toThrow('Worker returned no data');

    const second = w.writeBuffer(data);
    FakeWorker.last.reply({ id: 1, type: 'result', data: Uint8Array.of(1) });
    await expect(second).resolves.toEqual(Uint8Array.of(1));
  });

  it('resolves a relative wasmUrl against the page base', async () => {
    vi.stubGlobal('document', { baseURI: 'https://h/a/b.html' });
    const w = createXlsxWorker({ workerUrl: 'w.js', wasmUrl: 'x/y.wasm' });
    const p = w.readBuffer(new Uint8Array(4)).catch(() => {});
    expect((FakeWorker.last.posted[0] as Posted).msg.wasmUrl).toBe('https://h/a/x/y.wasm');
    FakeWorker.last.terminate();
    w.terminate();
    await p;
  });
});
