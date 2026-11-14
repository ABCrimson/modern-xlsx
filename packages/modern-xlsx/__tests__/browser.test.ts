import { describe, expect, it } from 'vitest';
import { initWasm, readBuffer, Workbook } from '../src/index.js';

describe('Browser runtime (Chromium)', () => {
  it('has WebAssembly and Blob support', () => {
    expect(typeof WebAssembly.instantiate).toBe('function');
    expect(new Blob(['hello'], { type: 'text/plain' }).size).toBe(5);
  });

  it('initializes WASM with no arguments and round-trips a workbook', async () => {
    await initWasm();
    const wb = new Workbook();
    wb.addSheet('Sheet1').cell('A1').value = 'hi';
    const buf = await wb.toBuffer();
    const back = await readBuffer(buf);
    expect(back.getSheet('Sheet1')?.cell('A1').value).toBe('hi');
  });
});
