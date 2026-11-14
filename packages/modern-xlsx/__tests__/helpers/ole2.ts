/**
 * OLE2 (CFB) fixtures for the encryption tests: a tiny compound-file builder
 * plus factories for the `EncryptionInfo` stream in its Agile (4.4) and
 * Standard (4.2) layouts. The ciphertext inside is placeholder bytes, so the
 * files are good for detection and error-path tests only.
 */

const ENDOFCHAIN = 0xfffffffe;
const FATSECT = 0xfffffffd;
const FREESECT = 0xffffffff;

/** One named stream of a compound file. */
export interface Ole2Stream {
  name: string;
  data: Uint8Array;
  /** Size written to the directory entry when it should differ from `data.length`. */
  declared?: bigint;
}

/** Options of {@link buildOle2}: the hostile knobs; the default is a version 3 file. */
export interface Ole2Options {
  /** Compound-file major version: 3 (512-byte sectors, the default) or 4 (4096-byte sectors). */
  major?: 3 | 4;
  /** Chain the directory sector to itself in the FAT instead of ending it. */
  selfLoopDirectory?: boolean;
}

function writeDirEntry(
  buf: Uint8Array,
  view: DataView,
  offset: number,
  name: string,
  entryType: number,
  startSector: number,
  size: bigint,
): void {
  // UTF-16LE name, then a null terminator
  for (let i = 0; i < name.length && i < 31; i++) {
    view.setUint16(offset + i * 2, name.charCodeAt(i), true);
  }
  view.setUint16(offset + name.length * 2, 0, true);
  // Name length in bytes, including the terminator
  view.setUint16(offset + 64, (name.length + 1) * 2, true);
  buf[offset + 66] = entryType;
  // Colour: black
  buf[offset + 67] = 1;
  // Siblings and child: none
  view.setUint32(offset + 68, FREESECT, true);
  view.setUint32(offset + 72, FREESECT, true);
  view.setUint32(offset + 76, FREESECT, true);
  view.setUint32(offset + 116, startSector, true);
  view.setBigUint64(offset + 120, size, true);
}

/** Allocates sector ids for each stream, starting at sector 2 (0 = directory, 1 = FAT). */
function allocateSectors(
  streams: Ole2Stream[],
  sectorSize: number,
): { streamSectors: number[][]; nextSector: number } {
  let nextSector = 2;
  const streamSectors: number[][] = streams.map((s) => {
    const count = Math.max(1, Math.ceil(s.data.length / sectorSize));
    return Array.from({ length: count }, () => nextSector++);
  });
  return { streamSectors, nextSector };
}

/** Writes the header: one FAT sector at 1, directory at 0, no MiniFAT, no DIFAT sector. */
function writeOle2Header(view: DataView, buf: Uint8Array, major: 3 | 4): void {
  buf.set([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]);
  view.setUint16(24, 62, true); // minor version
  view.setUint16(26, major, true); // major version
  view.setUint16(28, 0xfffe, true); // byte order
  view.setUint16(30, major === 4 ? 12 : 9, true); // sector shift
  view.setUint16(32, 6, true); // mini sector shift
  view.setUint32(44, 1, true); // FAT sector count
  view.setUint32(48, 0, true); // first directory sector
  view.setUint32(56, 4096, true); // mini stream cutoff
  view.setUint32(60, ENDOFCHAIN, true); // first mini FAT sector: none
  view.setUint32(68, ENDOFCHAIN, true); // first DIFAT sector: none
  view.setUint32(76, 1, true); // DIFAT[0] = the FAT sector
  for (let i = 1; i < 109; i++) view.setUint32(76 + i * 4, FREESECT, true);
}

/** Writes the root entry and one sibling-chained directory entry per stream. */
function writeStreamDirEntries(
  buf: Uint8Array,
  view: DataView,
  dirOffset: number,
  streams: Ole2Stream[],
  streamSectors: number[][],
): void {
  writeDirEntry(buf, view, dirOffset, 'Root Entry', 5, ENDOFCHAIN, 0n);
  if (streams.length > 0) view.setUint32(dirOffset + 76, 1, true);

  for (const [i, s] of streams.entries()) {
    const entryOffset = dirOffset + (i + 1) * 128;
    const startSid = streamSectors[i]?.[0] ?? ENDOFCHAIN;
    writeDirEntry(buf, view, entryOffset, s.name, 2, startSid, s.declared ?? BigInt(s.data.length));
    if (i + 1 < streams.length) view.setUint32(entryOffset + 72, i + 2, true);
  }
}

/** Writes the FAT chains and copies each stream's bytes into its sectors. */
function writeFatAndData(
  buf: Uint8Array,
  view: DataView,
  fatOffset: number,
  streams: Ole2Stream[],
  streamSectors: number[][],
  totalDataSectors: number,
  sectorSize: number,
): void {
  view.setUint32(fatOffset, ENDOFCHAIN, true);
  view.setUint32(fatOffset + 4, FATSECT, true);

  for (const [i, sids] of streamSectors.entries()) {
    for (const [j, sid] of sids.entries()) {
      const next = j + 1 < sids.length ? (sids[j + 1] ?? ENDOFCHAIN) : ENDOFCHAIN;
      view.setUint32(fatOffset + sid * 4, next, true);
    }
    const s = streams[i];
    if (!s) continue;
    let written = 0;
    for (const sid of sids) {
      const chunk = Math.min(s.data.length - written, sectorSize);
      if (chunk > 0) {
        buf.set(s.data.subarray(written, written + chunk), sectorSize + sid * sectorSize);
      }
      written += chunk;
    }
  }

  for (let i = totalDataSectors + 2; i < sectorSize / 4; i++) {
    view.setUint32(fatOffset + i * 4, FREESECT, true);
  }
}

/**
 * Builds a single-FAT-sector OLE2 compound document holding the given streams.
 * Pass an empty `data` for a stream that only has to exist. With the default
 * options the file is version 3 (512-byte sectors).
 */
export function buildOle2(streams: Ole2Stream[], opts: Ole2Options = {}): Uint8Array {
  const major = opts.major ?? 3;
  const sectorSize = major === 4 ? 4096 : 512;
  const { streamSectors, nextSector } = allocateSectors(streams, sectorSize);
  const buf = new Uint8Array(sectorSize + nextSector * sectorSize);
  const view = new DataView(buf.buffer);

  writeOle2Header(view, buf, major);
  writeStreamDirEntries(buf, view, sectorSize, streams, streamSectors);
  writeFatAndData(buf, view, 2 * sectorSize, streams, streamSectors, nextSector - 2, sectorSize);
  if (opts.selfLoopDirectory) view.setUint32(2 * sectorSize, 0, true);

  return buf;
}

/** Options for {@link agileEncInfo}. */
export interface AgileEncInfoOptions {
  /** AES key size in bits (default 256). */
  keyBits?: number;
  /** Digest size in bytes (default 64). */
  hashSize?: number;
  /** Hash algorithm name (default 'SHA512'). */
  hashAlgorithm?: string;
  /** Key-derivation iteration count (default 100000). */
  spinCount?: number;
  /** Use the long encrypted key and verifier values (default false). */
  long?: boolean;
}

/**
 * Builds an `EncryptionInfo` stream in the Agile layout (version 4.4): the
 * 8-byte header followed by the XML descriptor. The salts and ciphertext are
 * placeholder base64, so decryption fails by design.
 */
export function agileEncInfo(opts: AgileEncInfoOptions = {}): Uint8Array {
  const { keyBits = 256, hashSize = 64, hashAlgorithm = 'SHA512', spinCount = 100000 } = opts;
  const encryptedKeyValue = opts.long ? 'C'.repeat(68) : 'C'.repeat(22);
  const encryptedVerifierHashInput = opts.long ? 'D'.repeat(48) : 'D'.repeat(22);
  const encryptedVerifierHashValue = opts.long ? 'E'.repeat(92) : 'E'.repeat(22);
  const cipher = `cipherAlgorithm="AES" cipherChaining="ChainingModeCBC" hashAlgorithm="${hashAlgorithm}"`;
  const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<encryption xmlns="http://schemas.microsoft.com/office/2006/encryption" xmlns:p="http://schemas.microsoft.com/office/2006/keyEncryptor/password">
  <keyData saltSize="16" blockSize="16" keyBits="${keyBits}" hashSize="${hashSize}" ${cipher} saltValue="AAAAAAAAAAAAAAAAAAAAAA=="/>
  <dataIntegrity encryptedHmacKey="AAAAAAAAAAAAAAAAAAAAAA==" encryptedHmacValue="AAAAAAAAAAAAAAAAAAAAAA=="/>
  <keyEncryptors>
    <keyEncryptor uri="http://schemas.microsoft.com/office/2006/keyEncryptor/password">
      <p:encryptedKey spinCount="${spinCount}" saltSize="16" blockSize="16" keyBits="${keyBits}" hashSize="${hashSize}" ${cipher} saltValue="BBBBBBBBBBBBBBBBBBBBBB==" encryptedKeyValue="${encryptedKeyValue}==" encryptedVerifierHashInput="${encryptedVerifierHashInput}==" encryptedVerifierHashValue="${encryptedVerifierHashValue}=="/>
    </keyEncryptor>
  </keyEncryptors>
</encryption>`;
  const xmlBytes = new TextEncoder().encode(xml);
  const data = new Uint8Array(8 + xmlBytes.length);
  const view = new DataView(data.buffer);
  view.setUint16(0, 4, true); // major
  view.setUint16(2, 4, true); // minor
  view.setUint32(4, 0, true); // reserved
  data.set(xmlBytes, 8);
  return data;
}

/**
 * Builds the 124-byte `EncryptionInfo` stream of Standard encryption
 * (version 4.2, MS-OFFCRYPTO 2.3.4.5): AES-128 with SHA-1 and a placeholder
 * verifier.
 */
export function standardEncInfo(): Uint8Array {
  const data = new Uint8Array(8 + 4 + 40 + 72);
  const dv = new DataView(data.buffer);
  dv.setUint16(0, 4, true); // major
  dv.setUint16(2, 2, true); // minor -> Standard (4.2)
  dv.setUint32(4, 0x24, true); // flags
  dv.setUint32(8, 40, true); // headerSize
  // EncryptionHeader (40 bytes at offset 12)
  dv.setUint32(12, 0x24, true); // header flags
  dv.setUint32(16, 0, true); // sizeExtra
  dv.setUint32(20, 0x660e, true); // algID = AES-128
  dv.setUint32(24, 0x8004, true); // algIDHash = SHA-1
  dv.setUint32(28, 128, true); // keySize
  dv.setUint32(32, 0x18, true); // providerType
  dv.setUint32(36, 0, true); // reserved1
  dv.setUint32(40, 0, true); // reserved2
  // CSP name "AES\0" in UTF-16LE (8 bytes at offset 44)
  dv.setUint16(44, 0x41, true); // A
  dv.setUint16(46, 0x45, true); // E
  dv.setUint16(48, 0x53, true); // S
  dv.setUint16(50, 0, true); // null
  // EncryptionVerifier (72 bytes at offset 52, MS-OFFCRYPTO 2.3.3):
  // saltSize 4, salt 16, encrypted verifier 16, verifierHashSize 4, encrypted hash 32
  dv.setUint32(52, 16, true); // saltSize
  data.fill(0xaa, 56, 72); // salt
  data.fill(0xbb, 72, 88); // encrypted verifier
  dv.setUint32(88, 20, true); // verifierHashSize
  data.fill(0xcc, 92, 124); // encrypted hash
  return data;
}
