import { describe, expect, it } from 'vitest';
import { initWasm, ModernXlsxError, PASSWORD_PROTECTED, Workbook } from '../src/index.js';
import { createXlsxWorker } from '../src/worker-api.js';

const WORKER_URL = '/dist/modern-xlsx.worker.js';

async function sampleBytes(text = 'hi'): Promise<Uint8Array> {
  await initWasm();
  const wb = new Workbook();
  wb.addSheet('Sheet1').cell('A1').value = text;
  return wb.toBuffer();
}

describe('createXlsxWorker (module worker in Chromium)', () => {
  const urls: [string, string | undefined][] = [
    ['no wasmUrl (wasm beside the worker)', undefined],
    ['an absolute-path wasmUrl', '/dist/modern-xlsx.wasm'],
    ['a relative wasmUrl resolved against document.baseURI', 'dist/modern-xlsx.wasm'],
  ];

  for (const [label, wasmUrl] of urls) {
    it(`readBuffer resolves with ${label}`, async () => {
      const worker = createXlsxWorker({
        workerUrl: WORKER_URL,
        ...(wasmUrl === undefined ? {} : { wasmUrl }),
      });
      try {
        const data = await worker.readBuffer(await sampleBytes());
        expect(data.sheets[0]?.name).toBe('Sheet1');
      } finally {
        worker.terminate();
      }
    });
  }

  it('round-trips a password-protected workbook and rejects a wrong password', async () => {
    const worker = createXlsxWorker({ workerUrl: WORKER_URL });
    try {
      const plain = await worker.readBuffer(await sampleBytes('Secret'));
      const buf = await worker.writeBuffer(plain, { password: 'pw' });
      // Exact buffers are transferred, so pass slices.
      const data = await worker.readBuffer(buf.slice(), { password: 'pw' });
      expect(new Workbook(data).getSheet('Sheet1')?.cell('A1').value).toBe('Secret');

      const err = await worker.readBuffer(buf.slice(), { password: 'bad' }).catch((e) => e);
      expect(err).toBeInstanceOf(ModernXlsxError);
      expect((err as ModernXlsxError).code).toBe(PASSWORD_PROTECTED);
    } finally {
      worker.terminate();
    }
  });

  it('reports a failed WASM init as WASM_INIT_FAILED', async () => {
    const worker = createXlsxWorker({ workerUrl: WORKER_URL, wasmUrl: '/does-not-exist.wasm' });
    try {
      const err = await worker.readBuffer(await sampleBytes()).catch((e) => e);
      expect(err).toBeInstanceOf(ModernXlsxError);
      expect((err as ModernXlsxError).code).toBe('WASM_INIT_FAILED');
    } finally {
      worker.terminate();
    }
  });
});
