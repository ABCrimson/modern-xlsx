import { crc32, deflateRawSync, inflateRawSync } from 'node:zlib';

const EOCD_SIGNATURE = 0x06054b50;
const CENTRAL_SIGNATURE = 0x02014b50;
const LOCAL_SIGNATURE = 0x04034b50;
const EOCD_MIN_SIZE = 22;
const METHOD_STORED = 0;
const METHOD_DEFLATE = 8;

const decoder = new TextDecoder();

/**
 * Reads every entry of a ZIP archive into a map from entry name to decoded text.
 *
 * Sizes and offsets come from the central directory (a local header may defer
 * them to a data descriptor). Method 0 (stored) entries are copied and method 8
 * (deflate) entries are inflated with `inflateRawSync`; any other method throws.
 *
 * Every entry is decoded as UTF-8, so binary entries such as `xl/media/*.png`
 * come back as lossy text (invalid sequences become U+FFFD). Use the map to
 * check that such a part exists, never to check its bytes. ZIP64 archives are
 * not supported.
 */
export function readZipText(bytes: Uint8Array): Record<string, string> {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);

  let eocd = -1;
  for (let i = bytes.length - EOCD_MIN_SIZE; i >= 0; i--) {
    if (view.getUint32(i, true) === EOCD_SIGNATURE) {
      eocd = i;
      break;
    }
  }
  if (eocd < 0) throw new Error('readZipText: end of central directory not found');

  const entryCount = view.getUint16(eocd + 10, true);
  let cursor = view.getUint32(eocd + 16, true);
  const out: Record<string, string> = {};

  for (let n = 0; n < entryCount; n++) {
    if (view.getUint32(cursor, true) !== CENTRAL_SIGNATURE) {
      throw new Error('readZipText: bad central directory header');
    }
    const method = view.getUint16(cursor + 10, true);
    const compressedSize = view.getUint32(cursor + 20, true);
    const nameLength = view.getUint16(cursor + 28, true);
    const extraLength = view.getUint16(cursor + 30, true);
    const commentLength = view.getUint16(cursor + 32, true);
    const localOffset = view.getUint32(cursor + 42, true);
    const name = decoder.decode(bytes.subarray(cursor + 46, cursor + 46 + nameLength));
    cursor += 46 + nameLength + extraLength + commentLength;

    if (view.getUint32(localOffset, true) !== LOCAL_SIGNATURE) {
      throw new Error(`readZipText: bad local header for ${name}`);
    }
    const dataStart =
      localOffset +
      30 +
      view.getUint16(localOffset + 26, true) +
      view.getUint16(localOffset + 28, true);
    const data = bytes.subarray(dataStart, dataStart + compressedSize);

    if (method === METHOD_STORED) {
      out[name] = decoder.decode(data);
    } else if (method === METHOD_DEFLATE) {
      out[name] = decoder.decode(inflateRawSync(data));
    } else {
      throw new Error(`readZipText: unsupported compression method ${method} for ${name}`);
    }
  }
  return out;
}

/** One entry of {@link buildZip}: its name and raw (uncompressed) bytes. */
export type ZipPart = readonly [name: string, data: Uint8Array];

/**
 * Builds a ZIP archive from raw entries, byte by byte, so a name may repeat.
 *
 * `method` is `'deflate'` (the default, deflate-raw, method 8) or `'stored'`
 * (method 0). Stored entries carry general-purpose flag 0x0800 (UTF-8 names)
 * and the DOS date 0x21. Sizes and the CRC are always written in the headers
 * (no data descriptor); ZIP64 is not produced.
 */
export function buildZip(
  parts: readonly ZipPart[],
  opts: { method?: 'stored' | 'deflate' } = {},
): Uint8Array {
  const stored = opts.method === 'stored';
  const flags = stored ? 0x0800 : 0;
  const date = stored ? 0x21 : 0;
  const chunks: Buffer[] = [];
  const central: Buffer[] = [];
  let offset = 0;
  for (const [name, data] of parts) {
    const nameBytes = Buffer.from(name);
    const packed = stored ? data : deflateRawSync(data);
    const crc = crc32(data);
    const local = Buffer.alloc(30);
    local.writeUInt32LE(LOCAL_SIGNATURE, 0);
    local.writeUInt16LE(20, 4);
    local.writeUInt16LE(flags, 6);
    local.writeUInt16LE(stored ? METHOD_STORED : METHOD_DEFLATE, 8);
    local.writeUInt16LE(date, 12);
    local.writeUInt32LE(crc, 14);
    local.writeUInt32LE(packed.length, 18);
    local.writeUInt32LE(data.length, 22);
    local.writeUInt16LE(nameBytes.length, 26);
    const record = Buffer.alloc(46);
    record.writeUInt32LE(CENTRAL_SIGNATURE, 0);
    record.writeUInt16LE(20, 4);
    record.writeUInt16LE(20, 6);
    record.writeUInt16LE(flags, 8);
    record.writeUInt16LE(stored ? METHOD_STORED : METHOD_DEFLATE, 10);
    record.writeUInt16LE(date, 14);
    record.writeUInt32LE(crc, 16);
    record.writeUInt32LE(packed.length, 20);
    record.writeUInt32LE(data.length, 24);
    record.writeUInt16LE(nameBytes.length, 28);
    record.writeUInt32LE(offset, 42);
    chunks.push(local, nameBytes, Buffer.from(packed));
    central.push(record, nameBytes);
    offset += 30 + nameBytes.length + packed.length;
  }
  const directory = Buffer.concat(central);
  const end = Buffer.alloc(EOCD_MIN_SIZE);
  end.writeUInt32LE(EOCD_SIGNATURE, 0);
  end.writeUInt16LE(parts.length, 8);
  end.writeUInt16LE(parts.length, 10);
  end.writeUInt32LE(directory.length, 12);
  end.writeUInt32LE(offset, 16);
  return new Uint8Array(Buffer.concat([...chunks, directory, end]));
}
