import { createHash } from 'node:crypto';
import { describe, expect, it } from 'vitest';
import { readBuffer, Workbook } from '../src/index.js';

describe('Workbook Protection', () => {
  it('set and get lock structure', () => {
    const wb = new Workbook();
    wb.protection = { lockStructure: true };
    expect(wb.protection).toEqual({ lockStructure: true });
  });

  it('set and get all lock flags', () => {
    const wb = new Workbook();
    wb.protection = {
      lockStructure: true,
      lockWindows: true,
      lockRevision: true,
    };
    expect(wb.protection?.lockStructure).toBe(true);
    expect(wb.protection?.lockWindows).toBe(true);
    expect(wb.protection?.lockRevision).toBe(true);
  });

  it('lock structure survives roundtrip', async () => {
    const wb = new Workbook();
    wb.addSheet('Sheet1').cell('A1').value = 'data';
    wb.protection = { lockStructure: true };

    const buffer = await wb.toBuffer();
    const wb2 = await readBuffer(buffer);
    expect(wb2.protection?.lockStructure).toBe(true);
  });

  it('all protection fields survive roundtrip', async () => {
    const h = createHash('sha512').update('x').digest('base64');
    const s = Buffer.alloc(16, 7).toString('base64');
    const input = {
      lockStructure: true,
      lockWindows: true,
      lockRevision: true,
      workbookAlgorithmName: 'SHA-512',
      workbookHashValue: h,
      workbookSaltValue: s,
      workbookSpinCount: 100000,
      revisionsAlgorithmName: 'SHA-512',
      revisionsHashValue: h,
      revisionsSaltValue: s,
      revisionsSpinCount: 100000,
    };
    const wb = new Workbook();
    wb.addSheet('Sheet1').cell('A1').value = 'data';
    wb.protection = input;

    const wb2 = await readBuffer(await wb.toBuffer());
    expect(wb2.protection).toStrictEqual(input);
  });

  it('clear protection', () => {
    const wb = new Workbook();
    wb.protection = { lockStructure: true };
    expect(wb.protection).toBeTruthy();
    wb.protection = null;
    expect(wb.protection).toBeNull();
  });
});
