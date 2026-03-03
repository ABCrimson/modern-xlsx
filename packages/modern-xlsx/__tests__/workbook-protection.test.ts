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

  it('hash attributes survive roundtrip', async () => {
    const wb = new Workbook();
    wb.addSheet('Sheet1').cell('A1').value = 'data';
    wb.protection = {
      lockStructure: true,
      workbookAlgorithmName: 'SHA-512',
      workbookHashValue: 'abc123',
      workbookSaltValue: 'salt456',
      workbookSpinCount: 100000,
    };

    const buffer = await wb.toBuffer();
    const wb2 = await readBuffer(buffer);
    expect(wb2.protection?.lockStructure).toBe(true);
    expect(wb2.protection?.workbookAlgorithmName).toBe('SHA-512');
    expect(wb2.protection?.workbookHashValue).toBe('abc123');
    expect(wb2.protection?.workbookSaltValue).toBe('salt456');
    expect(wb2.protection?.workbookSpinCount).toBe(100000);
  });

  it('clear protection', () => {
    const wb = new Workbook();
    wb.protection = { lockStructure: true };
    expect(wb.protection).toBeTruthy();
    wb.protection = null;
    expect(wb.protection).toBeNull();
  });
});
