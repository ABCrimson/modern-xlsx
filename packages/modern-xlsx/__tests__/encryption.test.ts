import { assert, describe, expect, expectTypeOf, it } from 'vitest';
import type { ReadLimits, ReadOptions, WriteOptions } from '../src/index.js';
import { readBuffer, StyleBuilder, Workbook } from '../src/index.js';
import { PASSWORD_PROTECTED } from '../src/model/errors.js';
import { agileEncInfo, buildOle2, standardEncInfo } from './helpers/ole2.js';

describe('Encryption: OLE2 Detection', () => {
  it('encrypted XLSX shows descriptive error', async () => {
    const ole2 = buildOle2(
      ['EncryptionInfo', 'EncryptedPackage'].map((name) => ({ name, data: new Uint8Array(0) })),
    );
    await expect(readBuffer(ole2)).rejects.toThrow(/password.protected/i);
  });

  it('legacy .xls shows appropriate error', async () => {
    const ole2 = buildOle2([{ name: 'Workbook', data: new Uint8Array(0) }]);
    await expect(readBuffer(ole2)).rejects.toThrow(/legacy.*xls/i);
  });

  it('unknown format shows clear error', async () => {
    const random = new Uint8Array([0x01, 0x02, 0x03, 0x04, 0x05]);
    await expect(readBuffer(random)).rejects.toThrow(/not a valid xlsx/i);
  });

  it('empty file shows error', async () => {
    await expect(readBuffer(new Uint8Array(0))).rejects.toThrow();
  });
});

describe('Encryption: Method Identification', () => {
  it('encrypted AES-256 shows method in error', async () => {
    const encInfoData = agileEncInfo();
    const ole2 = buildOle2([
      { name: 'EncryptionInfo', data: encInfoData },
      { name: 'EncryptedPackage', data: new Uint8Array(0) },
    ]);
    await expect(readBuffer(ole2)).rejects.toThrow(/AES-256/);
  });

  it('encrypted AES-128 shows method in error', async () => {
    const encInfoData = agileEncInfo({ keyBits: 128, hashSize: 32, hashAlgorithm: 'SHA256' });
    const ole2 = buildOle2([
      { name: 'EncryptionInfo', data: encInfoData },
      { name: 'EncryptedPackage', data: new Uint8Array(0) },
    ]);
    await expect(readBuffer(ole2)).rejects.toThrow(/AES-128/);
  });

  it('standard encryption shows key size in error', async () => {
    const ole2 = buildOle2([
      { name: 'EncryptionInfo', data: standardEncInfo() },
      { name: 'EncryptedPackage', data: new Uint8Array(0) },
    ]);
    await expect(readBuffer(ole2)).rejects.toThrow(/Standard.*128/i);
  });

  it('encryption error includes usage hint', async () => {
    const encInfoData = agileEncInfo();
    const ole2 = buildOle2([
      { name: 'EncryptionInfo', data: encInfoData },
      { name: 'EncryptedPackage', data: new Uint8Array(0) },
    ]);
    await expect(readBuffer(ole2)).rejects.toThrow(
      "Provide password via readBuffer(data, { password: '...' }).",
    );
  });

  it('malformed EncryptionInfo graceful fallback', async () => {
    // OLE2 with EncryptionInfo + EncryptedPackage but EncryptionInfo contains garbage
    const ole2 = buildOle2([
      { name: 'EncryptionInfo', data: new Uint8Array([0x01, 0x02, 0x03]) },
      { name: 'EncryptedPackage', data: new Uint8Array(0) },
    ]);
    // Should still show password-protected error (graceful fallback)
    await expect(readBuffer(ole2)).rejects.toThrow(/password.protected/i);
  });
});

describe('Encryption: Decryption Integration', () => {
  it('normal XLSX with password option still works', async () => {
    const wb = new Workbook();
    wb.addSheet('Sheet1').cell('A1').value = 'hello';
    const buffer = await wb.toBuffer();
    // Passing password option on non-encrypted file should work (password ignored)
    const wb2 = await readBuffer(buffer, { password: 'test' });
    expect(wb2.getSheet('Sheet1')?.cell('A1').value).toBe('hello');
  });

  it('encrypted file with wrong password shows error', async () => {
    const encInfoData = agileEncInfo({ spinCount: 1, long: true });
    const ole2 = buildOle2([
      { name: 'EncryptionInfo', data: encInfoData },
      { name: 'EncryptedPackage', data: new Uint8Array(16) },
    ]);
    await expect(readBuffer(ole2, { password: 'wrong' })).rejects.toMatchObject({
      name: 'ModernXlsxError',
      code: PASSWORD_PROTECTED,
    });
  });

  it('readBuffer treats { password: undefined } as unencrypted', async () => {
    const wb = new Workbook();
    wb.addSheet('Test');
    const buf = await wb.toBuffer();
    expectTypeOf<ReadOptions>().toEqualTypeOf<{ password?: string; limits?: ReadLimits }>();
    expectTypeOf(readBuffer).parameter(1).toEqualTypeOf<ReadOptions | undefined>();
    // @ts-expect-error exactOptionalPropertyTypes rejects explicit undefined; JS callers may still pass it
    const opts: ReadOptions = { password: undefined };
    const wb2 = await readBuffer(buf, opts);
    expect(wb2.sheetCount).toBe(1);
  });
});

describe('Encryption: Standard Encryption Edge Cases', () => {
  it('encrypted file with workbook protection has both layers', async () => {
    // File-level encryption wraps the package; the inner workbook.xml still carries
    // <workbookProtection>. Both layers must survive an encrypted round trip.
    const wb = new Workbook();
    wb.addSheet('Protected').cell('A1').value = 'safe';
    wb.protection = { lockStructure: true, lockWindows: true };

    const encrypted = await wb.toBuffer({ password: 'pw' });
    const wb2 = await readBuffer(encrypted, { password: 'pw' });
    expect(wb2.protection).toEqual({ lockStructure: true, lockWindows: true });
    expect(wb2.getSheet('Protected')?.cell('A1').value).toBe('safe');
  });

  it('Standard encrypted OLE2 with wrong password reports PASSWORD_PROTECTED', async () => {
    // Build a Standard encryption OLE2 file (version 4.2) and verify that
    // decrypting it with a wrong password rejects with PASSWORD_PROTECTED.
    const encInfoData = standardEncInfo();

    // Build a minimal EncryptedPackage: an 8-byte StreamSize header, then 16 bytes of placeholder ciphertext
    const encPkgData = new Uint8Array(8 + 16);
    const pkgDv = new DataView(encPkgData.buffer);
    pkgDv.setBigUint64(0, 10n, true); // original size = 10
    for (let i = 0; i < 16; i++) encPkgData[8 + i] = 0xdd; // fake encrypted data

    const ole2 = buildOle2([
      { name: 'EncryptionInfo', data: encInfoData },
      { name: 'EncryptedPackage', data: encPkgData },
    ]);

    await expect(readBuffer(ole2, { password: 'wrongpass' })).rejects.toMatchObject({
      name: 'ModernXlsxError',
      code: PASSWORD_PROTECTED,
    });
  });
});

/**
 * Returns the start sector id of the `EncryptedPackage` stream in a 512-byte-sector
 * compound file, plus its stream size. Walks the directory chain through the FAT
 * named by the header DIFAT.
 */
function locateEncryptedPackage(cfb: Uint8Array): { start: number; size: number } {
  const view = new DataView(cfb.buffer, cfb.byteOffset, cfb.byteLength);
  const sectorAt = (sid: number): number => (sid + 1) * 512;
  const fat: number[] = [];
  for (let i = 0; i < 109; i++) {
    const fatSector = view.getUint32(76 + i * 4, true);
    if (fatSector >= 0xfffffffa) break;
    for (let j = 0; j < 128; j++) fat.push(view.getUint32(sectorAt(fatSector) + j * 4, true));
  }
  let sid = view.getUint32(48, true);
  while (sid < 0xfffffffa) {
    const base = sectorAt(sid);
    for (let entry = 0; entry < 4; entry++) {
      const at = base + entry * 128;
      const nameLength = view.getUint16(at + 64, true);
      let name = '';
      for (let c = 0; c < nameLength / 2 - 1; c++) {
        name += String.fromCharCode(view.getUint16(at + c * 2, true));
      }
      if (name === 'EncryptedPackage') {
        return { start: view.getUint32(at + 116, true), size: view.getUint32(at + 120, true) };
      }
    }
    sid = fat[sid] ?? 0xfffffffe;
  }
  throw new Error('EncryptedPackage not found');
}

/** Flips one bit of the ciphertext, 24 bytes into the first sector of `EncryptedPackage`. */
function flipInEncryptedPackage(cfb: Uint8Array): Uint8Array {
  const { start, size } = locateEncryptedPackage(cfb);
  expect(size).toBeGreaterThanOrEqual(4096);
  const tampered = cfb.slice();
  const at = (start + 1) * 512 + 24;
  const byte = tampered[at];
  assert.isDefined(byte);
  tampered[at] = byte ^ 0x01;
  return tampered;
}

const latin1 = (bytes: Uint8Array): string => Buffer.from(bytes).toString('latin1');

describe('Encryption: Write Path', () => {
  it('encrypt then decrypt roundtrip', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('EncTest');
    ws.cell('A1').value = 'hello encrypted';
    ws.cell('B2').value = 42;
    ws.cell('C3').value = true;

    const encrypted = await wb.toBuffer({ password: 'test123' });

    // Verify the output is an OLE2 file (starts with magic bytes)
    expect(Array.from(encrypted.subarray(0, 8))).toEqual([
      0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1,
    ]);

    // Decrypt and verify data
    const wb2 = await readBuffer(encrypted, { password: 'test123' });
    expect(wb2.getSheet('EncTest')?.cell('A1').value).toBe('hello encrypted');
    expect(wb2.getSheet('EncTest')?.cell('B2').value).toBe(42);
    expect(wb2.getSheet('EncTest')?.cell('C3').value).toBe(true);
  });

  it('encrypted file with wrong password fails', async () => {
    const wb = new Workbook();
    wb.addSheet('Sheet1').cell('A1').value = 'secret';
    const encrypted = await wb.toBuffer({ password: 'correct' });

    await expect(readBuffer(encrypted, { password: 'wrong' })).rejects.toMatchObject({
      name: 'ModernXlsxError',
      code: PASSWORD_PROTECTED,
      message: 'Password protected: Incorrect password.',
    });
  });

  it('encrypted file without password shows helpful error', async () => {
    const wb = new Workbook();
    wb.addSheet('Sheet1').cell('A1').value = 'secret';
    const encrypted = await wb.toBuffer({ password: 'mypass' });

    // Reading without password should show error mentioning password
    await expect(readBuffer(encrypted)).rejects.toMatchObject({
      name: 'ModernXlsxError',
      code: PASSWORD_PROTECTED,
    });
  });

  it('empty password produces normal unencrypted output', async () => {
    const wb = new Workbook();
    wb.addSheet('Sheet1').cell('A1').value = 'normal';

    // Empty string password should be treated as "no encryption"
    const buffer = await wb.toBuffer({ password: '' });

    // Should be a normal ZIP (starts with PK)
    expect(buffer[0]).toBe(0x50); // P
    expect(buffer[1]).toBe(0x4b); // K

    // Should read without password
    const wb2 = await readBuffer(buffer);
    expect(wb2.getSheet('Sheet1')?.cell('A1').value).toBe('normal');
  });

  it('encryption preserves styles', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('StyledSheet');

    const boldStyle = new StyleBuilder().font({ bold: true, size: 14 }).build(wb.styles);
    ws.cell('A1').value = 'Bold Text';
    ws.cell('A1').styleIndex = boldStyle;

    const encrypted = await wb.toBuffer({ password: 'styles' });
    const wb2 = await readBuffer(encrypted, { password: 'styles' });

    const cell = wb2.getSheet('StyledSheet')?.cell('A1');
    expect(cell?.value).toBe('Bold Text');
    expect(cell?.styleIndex).toBeGreaterThan(0);

    // Verify the font is bold
    const style = wb2.styles.cellXfs?.[cell?.styleIndex ?? 0];
    expect(style?.fontId).toBeDefined();
    expect(wb2.styles.fonts?.[style?.fontId ?? -1]?.bold).toBe(true);
  });

  it('toBuffer treats { password: undefined } as unencrypted', async () => {
    const wb = new Workbook();
    wb.addSheet('Test');

    expectTypeOf<WriteOptions>().toEqualTypeOf<{ password?: string }>();
    expectTypeOf<Workbook['toBuffer']>().parameter(0).toEqualTypeOf<WriteOptions | undefined>();
    // @ts-expect-error exactOptionalPropertyTypes rejects explicit undefined; JS callers may still pass it
    const opts: WriteOptions = { password: undefined };
    const buffer = await wb.toBuffer(opts);
    expect(buffer.length).toBeGreaterThan(0);
  });

  it('plaintext markers are absent from encrypted output', async () => {
    const wb = new Workbook();
    wb.addSheet('Sheet1').cell('A1').value = 'marker';
    const plain = latin1(await wb.toBuffer());
    expect(plain).toContain('PK\x03\x04');
    expect(plain).toContain('xl/workbook.xml');

    const encrypted = latin1(await wb.toBuffer({ password: 'pw' }));
    expect(encrypted).not.toContain('PK\x03\x04');
    expect(encrypted).not.toContain('xl/workbook.xml');
  });

  it('two encryptions of one workbook differ and both decrypt to the same value', async () => {
    const wb = new Workbook();
    wb.addSheet('Sheet1').cell('A1').value = 'same';
    const a = await wb.toBuffer({ password: 'pw' });
    const b = await wb.toBuffer({ password: 'pw' });
    expect(Buffer.from(a).equals(Buffer.from(b))).toBe(false);
    for (const enc of [a, b]) {
      const back = await readBuffer(enc, { password: 'pw' });
      expect(back.getSheet('Sheet1')?.cell('A1').value).toBe('same');
    }
  });

  it('a flipped ciphertext byte fails the HMAC check', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Big');
    for (let r = 1; r <= 400; r++) ws.cell(`A${r}`).value = `row ${r} ${r * 7919}`;
    const encrypted = await wb.toBuffer({ password: 'pw' });
    await expect(readBuffer(flipInEncryptedPackage(encrypted), { password: 'pw' })).rejects.toThrow(
      /HMAC/i,
    );
  });
});
