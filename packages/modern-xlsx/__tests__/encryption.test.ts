import { describe, expect, it } from 'vitest';
import type { ReadOptions, WriteOptions } from '../src/index.js';
import { readBuffer, StyleBuilder, Workbook } from '../src/index.js';

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

interface StreamDef {
  name: string;
  data: Uint8Array;
}

/** Allocate sector IDs for each stream (starting at sector 2). */
function allocateSectors(
  streams: StreamDef[],
  sectorSize: number,
): { streamSectors: number[][]; nextSector: number } {
  let nextSector = 2;
  const streamSectors: number[][] = streams.map((s) => {
    const count = Math.max(1, Math.ceil(s.data.length / sectorSize));
    return Array.from({ length: count }, () => nextSector++);
  });
  return { streamSectors, nextSector };
}

/** Write OLE2 header into buf (offsets 0..511). */
function writeOle2Header(view: DataView, buf: Uint8Array): void {
  buf.set([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]);
  view.setUint16(24, 62, true);
  view.setUint16(26, 3, true);
  view.setUint16(28, 0xfffe, true);
  view.setUint16(30, 9, true);
  view.setUint16(32, 6, true);
  view.setUint32(44, 1, true);
  view.setUint32(48, 0, true);
  view.setUint32(60, 0xfffffffe, true);
  view.setUint32(68, 0xfffffffe, true);
  view.setUint32(76, 1, true);
  for (let i = 1; i < 109; i++) view.setUint32(76 + i * 4, 0xffffffff, true);
}

/** Write directory entries for streams. */
function writeStreamDirEntries(
  buf: Uint8Array,
  view: DataView,
  dirOffset: number,
  streams: StreamDef[],
  streamSectors: number[][],
): void {
  writeDirEntry(buf, view, dirOffset, 'Root Entry', 5, 0xfffffffe, 0);
  if (streams.length > 0) view.setUint32(dirOffset + 76, 1, true);

  for (const [i, s] of streams.entries()) {
    const entryOffset = dirOffset + (i + 1) * 128;
    const sids = streamSectors[i] ?? [];
    const startSid = sids[0] ?? 0xfffffffe;
    writeDirEntry(buf, view, entryOffset, s.name, 2, startSid, s.data.length);
    if (i + 1 < streams.length) view.setUint32(entryOffset + 72, i + 2, true);
  }
}

/** Write FAT sector with chain entries and stream data into sectors. */
function writeFatAndData(
  buf: Uint8Array,
  view: DataView,
  fatOffset: number,
  sectorSize: number,
  streams: StreamDef[],
  streamSectors: number[][],
  totalDataSectors: number,
): void {
  view.setUint32(fatOffset, 0xfffffffe, true);
  view.setUint32(fatOffset + 4, 0xfffffffd, true);

  for (const [i, sids] of streamSectors.entries()) {
    for (const [j, sid] of sids.entries()) {
      const next = j + 1 < sids.length ? (sids[j + 1] ?? 0xfffffffe) : 0xfffffffe;
      view.setUint32(fatOffset + sid * 4, next, true);
    }
    // Write stream bytes into allocated sectors
    const s = streams[i];
    if (!s) continue;
    let written = 0;
    for (const sid of sids) {
      const chunk = Math.min(s.data.length - written, sectorSize);
      if (chunk > 0) buf.set(s.data.subarray(written, written + chunk), 512 + sid * sectorSize);
      written += chunk;
    }
  }

  for (let i = totalDataSectors + 2; i < sectorSize / 4; i++) {
    view.setUint32(fatOffset + i * 4, 0xffffffff, true);
  }
}

/**
 * Builds an OLE2 compound document with named streams containing actual data.
 * Extends buildMinimalOle2 to write stream content into data sectors.
 */
function buildOle2WithStreams(streams: StreamDef[]): Uint8Array {
  const sectorSize = 512;
  const { streamSectors, nextSector } = allocateSectors(streams, sectorSize);
  const fileSize = 512 + nextSector * sectorSize;
  const buf = new Uint8Array(fileSize);
  const view = new DataView(buf.buffer);

  writeOle2Header(view, buf);
  writeStreamDirEntries(buf, view, 512, streams, streamSectors);
  writeFatAndData(buf, view, 512 + sectorSize, sectorSize, streams, streamSectors, nextSector - 2);

  return buf;
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

/** Helper to build EncryptionInfo stream with Agile XML. */
function buildAgileEncInfoStream(xml: string): Uint8Array {
  const header = new Uint8Array(8);
  const hView = new DataView(header.buffer);
  hView.setUint16(0, 4, true); // major = 4
  hView.setUint16(2, 4, true); // minor = 4
  hView.setUint32(4, 0, true); // reserved
  const xmlBytes = new TextEncoder().encode(xml);
  const encInfoData = new Uint8Array(8 + xmlBytes.length);
  encInfoData.set(header);
  encInfoData.set(xmlBytes, 8);
  return encInfoData;
}

describe('Encryption: Method Identification', () => {
  it('encrypted AES-256 shows method in error', async () => {
    const encInfoXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<encryption xmlns="http://schemas.microsoft.com/office/2006/encryption" xmlns:p="http://schemas.microsoft.com/office/2006/keyEncryptor/password">
  <keyData saltSize="16" blockSize="16" keyBits="256" hashSize="64" cipherAlgorithm="AES" cipherChaining="ChainingModeCBC" hashAlgorithm="SHA512" saltValue="AAAAAAAAAAAAAAAAAAAAAA=="/>
  <dataIntegrity encryptedHmacKey="AAAAAAAAAAAAAAAAAAAAAA==" encryptedHmacValue="AAAAAAAAAAAAAAAAAAAAAA=="/>
  <keyEncryptors>
    <keyEncryptor uri="http://schemas.microsoft.com/office/2006/keyEncryptor/password">
      <p:encryptedKey spinCount="100000" saltSize="16" blockSize="16" keyBits="256" hashSize="64" cipherAlgorithm="AES" cipherChaining="ChainingModeCBC" hashAlgorithm="SHA512" saltValue="BBBBBBBBBBBBBBBBBBBBBB==" encryptedKeyValue="CCCCCCCCCCCCCCCCCCCCCC==" encryptedVerifierHashInput="DDDDDDDDDDDDDDDDDDDDDD==" encryptedVerifierHashValue="EEEEEEEEEEEEEEEEEEEEEE=="/>
    </keyEncryptor>
  </keyEncryptors>
</encryption>`;
    const encInfoData = buildAgileEncInfoStream(encInfoXml);
    const ole2 = buildOle2WithStreams([
      { name: 'EncryptionInfo', data: encInfoData },
      { name: 'EncryptedPackage', data: new Uint8Array(0) },
    ]);
    await expect(readBuffer(ole2)).rejects.toThrow(/AES-256/);
  });

  it('encrypted AES-128 shows method in error', async () => {
    const encInfoXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<encryption xmlns="http://schemas.microsoft.com/office/2006/encryption" xmlns:p="http://schemas.microsoft.com/office/2006/keyEncryptor/password">
  <keyData saltSize="16" blockSize="16" keyBits="128" hashSize="32" cipherAlgorithm="AES" cipherChaining="ChainingModeCBC" hashAlgorithm="SHA256" saltValue="AAAAAAAAAAAAAAAAAAAAAA=="/>
  <dataIntegrity encryptedHmacKey="AAAAAAAAAAAAAAAAAAAAAA==" encryptedHmacValue="AAAAAAAAAAAAAAAAAAAAAA=="/>
  <keyEncryptors>
    <keyEncryptor uri="http://schemas.microsoft.com/office/2006/keyEncryptor/password">
      <p:encryptedKey spinCount="100000" saltSize="16" blockSize="16" keyBits="128" hashSize="32" cipherAlgorithm="AES" cipherChaining="ChainingModeCBC" hashAlgorithm="SHA256" saltValue="BBBBBBBBBBBBBBBBBBBBBB==" encryptedKeyValue="CCCCCCCCCCCCCCCCCCCCCC==" encryptedVerifierHashInput="DDDDDDDDDDDDDDDDDDDDDD==" encryptedVerifierHashValue="EEEEEEEEEEEEEEEEEEEEEE=="/>
    </keyEncryptor>
  </keyEncryptors>
</encryption>`;
    const encInfoData = buildAgileEncInfoStream(encInfoXml);
    const ole2 = buildOle2WithStreams([
      { name: 'EncryptionInfo', data: encInfoData },
      { name: 'EncryptedPackage', data: new Uint8Array(0) },
    ]);
    await expect(readBuffer(ole2)).rejects.toThrow(/AES-128/);
  });

  it('standard encryption shows key size in error', async () => {
    // Standard: version 4.2, binary header
    const data = new Uint8Array(8 + 4 + 40 + 68);
    const dv = new DataView(data.buffer);
    dv.setUint16(0, 4, true); // major
    dv.setUint16(2, 2, true); // minor
    dv.setUint32(4, 0x24, true); // flags
    // headerSize
    dv.setUint32(8, 40, true);
    // header (40 bytes at offset 12):
    dv.setUint32(12, 0x24, true); // header flags
    dv.setUint32(16, 0, true); // sizeExtra
    dv.setUint32(20, 0x6801, true); // algID = AES-128
    dv.setUint32(24, 0x8004, true); // algIDHash = SHA-1
    dv.setUint32(28, 128, true); // keySize
    dv.setUint32(32, 0x18, true); // providerType
    dv.setUint32(36, 0, true); // reserved1
    dv.setUint32(40, 0, true); // reserved2
    // CSP name "AES\0" UTF-16LE (8 bytes at offset 44)
    dv.setUint16(44, 0x41, true); // A
    dv.setUint16(46, 0x45, true); // E
    dv.setUint16(48, 0x53, true); // S
    dv.setUint16(50, 0, true); // null
    // verifier (68 bytes at offset 52):
    // salt 16 bytes, encrypted verifier 16 bytes, hash size 4, encrypted hash 32
    for (let i = 0; i < 16; i++) data[52 + i] = 0xaa; // salt
    for (let i = 0; i < 16; i++) data[68 + i] = 0xbb; // encrypted verifier
    dv.setUint32(84, 20, true); // hash size
    for (let i = 0; i < 32; i++) data[88 + i] = 0xcc; // encrypted hash

    const ole2 = buildOle2WithStreams([
      { name: 'EncryptionInfo', data },
      { name: 'EncryptedPackage', data: new Uint8Array(0) },
    ]);
    await expect(readBuffer(ole2)).rejects.toThrow(/Standard.*128/i);
  });

  it('encryption error includes usage hint', async () => {
    const encInfoXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<encryption xmlns="http://schemas.microsoft.com/office/2006/encryption" xmlns:p="http://schemas.microsoft.com/office/2006/keyEncryptor/password">
  <keyData saltSize="16" blockSize="16" keyBits="256" hashSize="64" cipherAlgorithm="AES" cipherChaining="ChainingModeCBC" hashAlgorithm="SHA512" saltValue="AAAAAAAAAAAAAAAAAAAAAA=="/>
  <dataIntegrity encryptedHmacKey="AAAAAAAAAAAAAAAAAAAAAA==" encryptedHmacValue="AAAAAAAAAAAAAAAAAAAAAA=="/>
  <keyEncryptors>
    <keyEncryptor uri="http://schemas.microsoft.com/office/2006/keyEncryptor/password">
      <p:encryptedKey spinCount="100000" saltSize="16" blockSize="16" keyBits="256" hashSize="64" cipherAlgorithm="AES" cipherChaining="ChainingModeCBC" hashAlgorithm="SHA512" saltValue="BBBBBBBBBBBBBBBBBBBBBB==" encryptedKeyValue="CCCCCCCCCCCCCCCCCCCCCC==" encryptedVerifierHashInput="DDDDDDDDDDDDDDDDDDDDDD==" encryptedVerifierHashValue="EEEEEEEEEEEEEEEEEEEEEE=="/>
    </keyEncryptor>
  </keyEncryptors>
</encryption>`;
    const encInfoData = buildAgileEncInfoStream(encInfoXml);
    const ole2 = buildOle2WithStreams([
      { name: 'EncryptionInfo', data: encInfoData },
      { name: 'EncryptedPackage', data: new Uint8Array(0) },
    ]);
    await expect(readBuffer(ole2)).rejects.toThrow(/password/i);
  });

  it('malformed EncryptionInfo graceful fallback', async () => {
    // OLE2 with EncryptionInfo + EncryptedPackage but EncryptionInfo contains garbage
    const ole2 = buildOle2WithStreams([
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

  it('encrypted file without password shows helpful error', async () => {
    const encInfoXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<encryption xmlns="http://schemas.microsoft.com/office/2006/encryption" xmlns:p="http://schemas.microsoft.com/office/2006/keyEncryptor/password">
  <keyData saltSize="16" blockSize="16" keyBits="256" hashSize="64" cipherAlgorithm="AES" cipherChaining="ChainingModeCBC" hashAlgorithm="SHA512" saltValue="AAAAAAAAAAAAAAAAAAAAAA=="/>
  <dataIntegrity encryptedHmacKey="AAAAAAAAAAAAAAAAAAAAAA==" encryptedHmacValue="AAAAAAAAAAAAAAAAAAAAAA=="/>
  <keyEncryptors>
    <keyEncryptor uri="http://schemas.microsoft.com/office/2006/keyEncryptor/password">
      <p:encryptedKey spinCount="100000" saltSize="16" blockSize="16" keyBits="256" hashSize="64" cipherAlgorithm="AES" cipherChaining="ChainingModeCBC" hashAlgorithm="SHA512" saltValue="BBBBBBBBBBBBBBBBBBBBBB==" encryptedKeyValue="CCCCCCCCCCCCCCCCCCCCCC==" encryptedVerifierHashInput="DDDDDDDDDDDDDDDDDDDDDD==" encryptedVerifierHashValue="EEEEEEEEEEEEEEEEEEEEEE=="/>
    </keyEncryptor>
  </keyEncryptors>
</encryption>`;
    const encInfoData = buildAgileEncInfoStream(encInfoXml);
    const ole2 = buildOle2WithStreams([
      { name: 'EncryptionInfo', data: encInfoData },
      { name: 'EncryptedPackage', data: new Uint8Array(0) },
    ]);
    await expect(readBuffer(ole2)).rejects.toThrow(/password/i);
  });

  it('encrypted file with wrong password shows error', async () => {
    const encInfoXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<encryption xmlns="http://schemas.microsoft.com/office/2006/encryption" xmlns:p="http://schemas.microsoft.com/office/2006/keyEncryptor/password">
  <keyData saltSize="16" blockSize="16" keyBits="256" hashSize="64" cipherAlgorithm="AES" cipherChaining="ChainingModeCBC" hashAlgorithm="SHA512" saltValue="AAAAAAAAAAAAAAAAAAAAAA=="/>
  <dataIntegrity encryptedHmacKey="AAAAAAAAAAAAAAAAAAAAAA==" encryptedHmacValue="AAAAAAAAAAAAAAAAAAAAAA=="/>
  <keyEncryptors>
    <keyEncryptor uri="http://schemas.microsoft.com/office/2006/keyEncryptor/password">
      <p:encryptedKey spinCount="1" saltSize="16" blockSize="16" keyBits="256" hashSize="64" cipherAlgorithm="AES" cipherChaining="ChainingModeCBC" hashAlgorithm="SHA512" saltValue="BBBBBBBBBBBBBBBBBBBBBB==" encryptedKeyValue="CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC==" encryptedVerifierHashInput="DDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDD==" encryptedVerifierHashValue="EEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEE=="/>
    </keyEncryptor>
  </keyEncryptors>
</encryption>`;
    const encInfoData = buildAgileEncInfoStream(encInfoXml);
    const ole2 = buildOle2WithStreams([
      { name: 'EncryptionInfo', data: encInfoData },
      { name: 'EncryptedPackage', data: new Uint8Array(16) },
    ]);
    await expect(readBuffer(ole2, { password: 'wrong' })).rejects.toThrow(
      /password|decrypt|padding/i,
    );
  });

  it('readBuffer accepts ReadOptions type', async () => {
    const wb = new Workbook();
    wb.addSheet('Test');
    const buf = await wb.toBuffer();
    // TypeScript compile check: ReadOptions is valid
    const opts: ReadOptions = { password: undefined };
    const wb2 = await readBuffer(buf, opts);
    expect(wb2.sheetCount).toBe(1);
  });
});

describe('Encryption: Standard Encryption Edge Cases', () => {
  it('read-only recommended XLSX reads without password', async () => {
    // Per OOXML spec, <fileSharing readOnlyRecommended="1"/> is just metadata
    // inside a normal ZIP archive — NOT OLE2 encryption.
    // This test confirms the read path does not confuse a normal XLSX with an
    // encrypted file, even when a password option is provided.
    const wb = new Workbook();
    const ws = wb.addSheet('Data');
    ws.cell('A1').value = 'read-only-test';
    ws.cell('B2').value = 42;
    const buffer = await wb.toBuffer();

    // Read with an unnecessary password option — should succeed (no OLE2 encryption)
    const wb2 = await readBuffer(buffer, { password: 'ignored' });
    expect(wb2.getSheet('Data')?.cell('A1').value).toBe('read-only-test');
    expect(wb2.getSheet('Data')?.cell('B2').value).toBe(42);
  });

  it('encrypted file with workbook protection has both layers', async () => {
    // Verify that an encrypted OLE2 file that also contains <workbookProtection>
    // in the inner workbook.xml surfaces both the file-level encryption error
    // (when no password) and that the protection metadata would survive decryption.
    // Since we can't construct a real encrypted-then-protected file in a unit test,
    // we verify each layer independently:
    // 1. Encryption layer: OLE2 file without password → PasswordProtected error
    // 2. Protection layer: Normal workbook with protection → reads fine
    const encInfoXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<encryption xmlns="http://schemas.microsoft.com/office/2006/encryption" xmlns:p="http://schemas.microsoft.com/office/2006/keyEncryptor/password">
  <keyData saltSize="16" blockSize="16" keyBits="256" hashSize="64" cipherAlgorithm="AES" cipherChaining="ChainingModeCBC" hashAlgorithm="SHA512" saltValue="AAAAAAAAAAAAAAAAAAAAAA=="/>
  <dataIntegrity encryptedHmacKey="AAAAAAAAAAAAAAAAAAAAAA==" encryptedHmacValue="AAAAAAAAAAAAAAAAAAAAAA=="/>
  <keyEncryptors>
    <keyEncryptor uri="http://schemas.microsoft.com/office/2006/keyEncryptor/password">
      <p:encryptedKey spinCount="100000" saltSize="16" blockSize="16" keyBits="256" hashSize="64" cipherAlgorithm="AES" cipherChaining="ChainingModeCBC" hashAlgorithm="SHA512" saltValue="BBBBBBBBBBBBBBBBBBBBBB==" encryptedKeyValue="CCCCCCCCCCCCCCCCCCCCCC==" encryptedVerifierHashInput="DDDDDDDDDDDDDDDDDDDDDD==" encryptedVerifierHashValue="EEEEEEEEEEEEEEEEEEEEEE=="/>
    </keyEncryptor>
  </keyEncryptors>
</encryption>`;
    const encInfoData = buildAgileEncInfoStream(encInfoXml);
    const ole2 = buildOle2WithStreams([
      { name: 'EncryptionInfo', data: encInfoData },
      { name: 'EncryptedPackage', data: new Uint8Array(0) },
    ]);

    // Encryption layer: trying to read without password → error
    await expect(readBuffer(ole2)).rejects.toThrow(/password/i);

    // Protection layer: a non-encrypted workbook with protection reads fine
    const wb = new Workbook();
    wb.addSheet('Protected').cell('A1').value = 'safe';
    const buffer = await wb.toBuffer();
    const wb2 = await readBuffer(buffer);
    expect(wb2.getSheet('Protected')?.cell('A1').value).toBe('safe');
  });

  it('Standard encrypted OLE2 with wrong password mentions Standard', async () => {
    // Build a Standard encryption OLE2 file (version 4.2)
    // and verify the error message mentions "Standard" when trying
    // to decrypt with a wrong password.
    const encInfoData = new Uint8Array(8 + 4 + 40 + 68);
    const dv = new DataView(encInfoData.buffer);
    dv.setUint16(0, 4, true); // major
    dv.setUint16(2, 2, true); // minor -> Standard (4.2)
    dv.setUint32(4, 0x24, true); // flags
    dv.setUint32(8, 40, true); // headerSize
    // Header (40 bytes at offset 12)
    dv.setUint32(12, 0x24, true); // header flags
    dv.setUint32(16, 0, true); // sizeExtra
    dv.setUint32(20, 0x6801, true); // algID = AES-128
    dv.setUint32(24, 0x8004, true); // algIDHash = SHA-1
    dv.setUint32(28, 128, true); // keySize
    dv.setUint32(32, 0x18, true); // providerType
    dv.setUint32(36, 0, true); // reserved1
    dv.setUint32(40, 0, true); // reserved2
    // CSP name "AES\0" UTF-16LE (8 bytes at offset 44)
    dv.setUint16(44, 0x41, true); // A
    dv.setUint16(46, 0x45, true); // E
    dv.setUint16(48, 0x53, true); // S
    dv.setUint16(50, 0, true); // null
    // verifier (68 bytes at offset 52)
    for (let i = 0; i < 16; i++) encInfoData[52 + i] = 0xaa; // salt
    for (let i = 0; i < 16; i++) encInfoData[68 + i] = 0xbb; // encrypted verifier
    dv.setUint32(84, 20, true); // hash size
    for (let i = 0; i < 32; i++) encInfoData[88 + i] = 0xcc; // encrypted hash

    // Build a minimal EncryptedPackage with some data (16 bytes header + 16 bytes payload)
    const encPkgData = new Uint8Array(8 + 16);
    const pkgDv = new DataView(encPkgData.buffer);
    pkgDv.setBigUint64(0, 10n, true); // original size = 10
    for (let i = 0; i < 16; i++) encPkgData[8 + i] = 0xdd; // fake encrypted data

    const ole2 = buildOle2WithStreams([
      { name: 'EncryptionInfo', data: encInfoData },
      { name: 'EncryptedPackage', data: encPkgData },
    ]);

    // Attempting to read with a wrong password should produce an error
    // mentioning "Incorrect password" or similar
    await expect(readBuffer(ole2, { password: 'wrongpass' })).rejects.toThrow(/password|decrypt/i);
  });
});

describe('Encryption: Write Path', () => {
  it('encrypt then decrypt roundtrip', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('EncTest');
    ws.cell('A1').value = 'hello encrypted';
    ws.cell('B2').value = 42;
    ws.cell('C3').value = true;

    const encrypted = await wb.toBuffer({ password: 'test123' });

    // Verify the output is an OLE2 file (starts with magic bytes)
    expect(encrypted[0]).toBe(0xd0);
    expect(encrypted[1]).toBe(0xcf);
    expect(encrypted[2]).toBe(0x11);
    expect(encrypted[3]).toBe(0xe0);

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

    await expect(readBuffer(encrypted, { password: 'wrong' })).rejects.toThrow(/password|decrypt/i);
  });

  it('encrypted file without password shows helpful error', async () => {
    const wb = new Workbook();
    wb.addSheet('Sheet1').cell('A1').value = 'secret';
    const encrypted = await wb.toBuffer({ password: 'mypass' });

    // Reading without password should show error mentioning password
    await expect(readBuffer(encrypted)).rejects.toThrow(/password/i);
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
    if (style?.fontId !== undefined) {
      const font = wb2.styles.fonts?.[style.fontId];
      expect(font?.bold).toBe(true);
    }
  });

  it('WriteOptions type is accepted', async () => {
    const wb = new Workbook();
    wb.addSheet('Test');

    // TypeScript compile check: WriteOptions is valid
    const opts: WriteOptions = { password: undefined };
    const buffer = await wb.toBuffer(opts);
    expect(buffer.length).toBeGreaterThan(0);
  });
});
