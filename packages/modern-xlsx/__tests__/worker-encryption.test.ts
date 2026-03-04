import { describe, expect, it } from 'vitest';
import { initWasm, readBuffer, Workbook } from '../src/index.js';

describe('Worker encryption protocol', () => {
  it('readBuffer with password option works on main thread', async () => {
    await initWasm();
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    ws.cell('A1').value = 'Secret';
    const encrypted = await wb.toBuffer({ password: 'test123' });
    const wb2 = await readBuffer(encrypted, { password: 'test123' });
    expect(wb2.getSheet('Sheet1')?.cell('A1').value).toBe('Secret');
  });
});
