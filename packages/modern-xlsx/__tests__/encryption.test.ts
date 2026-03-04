import { describe, expect, it } from 'vitest';
import { readBuffer, Workbook } from '../src/index.js';

/**
 * Builds a minimal OLE2 compound document with the given stream names.
 * Used to create test fixtures for encrypted/legacy file detection.
 */
function buildMinimalOle2(streamNames: string[]): Uint8Array {
  const sectorSize = 512;
  const totalSectors = 2; // directory + FAT
  const fileSize = 512 + totalSectors * sectorSize;
  const buf = new Uint8Array(fileSize);
  const view = new DataView(buf.buffer);

  // Magic bytes
  buf.set([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]);
  // Minor version
  view.setUint16(24, 62, true);
  // Major version 3
  view.setUint16(26, 3, true);
  // Byte order 0xFFFE
  view.setUint16(28, 0xfffe, true);
  // Sector size shift 9
  view.setUint16(30, 9, true);
  // Mini sector size shift 6
  view.setUint16(32, 6, true);
  // FAT sectors: 1
  view.setUint32(44, 1, true);
  // First dir sector: 0
  view.setUint32(48, 0, true);
  // First mini FAT: none
  view.setUint32(60, 0xfffffffe, true);
  // First DIFAT: none
  view.setUint32(68, 0xfffffffe, true);
  // DIFAT[0] = sector 1 (FAT sector)
  view.setUint32(76, 1, true);
  // Rest of DIFAT: free
  for (let i = 1; i < 109; i++) {
    view.setUint32(76 + i * 4, 0xffffffff, true);
  }

  const dirOffset = 512;

  // Root Entry
  writeDirEntry(buf, view, dirOffset, 'Root Entry', 5, 0xfffffffe, 0);
  if (streamNames.length > 0) {
    view.setUint32(dirOffset + 76, 1, true); // child = entry 1
  }

  // Stream entries
  for (let i = 0; i < streamNames.length; i++) {
    const entryOffset = dirOffset + (i + 1) * 128;
    const name = streamNames[i];
    if (!name) continue;
    writeDirEntry(buf, view, entryOffset, name, 2, 0xfffffffe, 0);
    // Right sibling
    if (i + 1 < streamNames.length) {
      view.setUint32(entryOffset + 72, i + 2, true);
    }
  }

  // FAT sector at offset 1024
  const fatOffset = 1024;
  view.setUint32(fatOffset, 0xfffffffe, true); // sector 0: end of dir chain
  view.setUint32(fatOffset + 4, 0xfffffffd, true); // sector 1: FAT sector
  for (let i = 2; i < sectorSize / 4; i++) {
    view.setUint32(fatOffset + i * 4, 0xffffffff, true); // free
  }

  return buf;
}

function writeDirEntry(
  buf: Uint8Array,
  view: DataView,
  offset: number,
  name: string,
  entryType: number,
  startSector: number,
  size: number,
): void {
  // UTF-16LE name
  for (let i = 0; i < name.length && i < 31; i++) {
    view.setUint16(offset + i * 2, name.charCodeAt(i), true);
  }
  // Null terminator
  view.setUint16(offset + name.length * 2, 0, true);
  // Name length (bytes, including null)
  view.setUint16(offset + 64, (name.length + 1) * 2, true);
  // Entry type
  buf[offset + 66] = entryType;
  // Color: black
  buf[offset + 67] = 1;
  // Siblings: none
  view.setUint32(offset + 68, 0xffffffff, true);
  view.setUint32(offset + 72, 0xffffffff, true);
  // Child: none
  view.setUint32(offset + 76, 0xffffffff, true);
  // Start sector
  view.setUint32(offset + 116, startSector, true);
  // Size
  view.setUint32(offset + 120, size, true);
}

describe('Encryption: OLE2 Detection', () => {
  it('encrypted XLSX shows descriptive error', async () => {
    const ole2 = buildMinimalOle2(['EncryptionInfo', 'EncryptedPackage']);
    await expect(readBuffer(ole2)).rejects.toThrow(/password.protected/i);
  });

  it('legacy .xls shows appropriate error', async () => {
    const ole2 = buildMinimalOle2(['Workbook']);
    await expect(readBuffer(ole2)).rejects.toThrow(/legacy.*xls/i);
  });

  it('unknown format shows clear error', async () => {
    const random = new Uint8Array([0x01, 0x02, 0x03, 0x04, 0x05]);
    await expect(readBuffer(random)).rejects.toThrow(/not a valid xlsx/i);
  });

  it('normal XLSX still works', async () => {
    const wb = new Workbook();
    wb.addSheet('Sheet1').cell('A1').value = 'hello';
    const buffer = await wb.toBuffer();
    const wb2 = await readBuffer(buffer);
    expect(wb2.getSheet('Sheet1')?.cell('A1').value).toBe('hello');
  });

  it('empty file shows error', async () => {
    await expect(readBuffer(new Uint8Array(0))).rejects.toThrow();
  });
});
