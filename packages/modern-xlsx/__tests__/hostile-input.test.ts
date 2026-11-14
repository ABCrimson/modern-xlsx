/**
 * Hostile-input regression suite: crafted ZIP, OLE2 and
 * XML inputs must reject with the exact `ModernXlsxError` code, never a WASM
 * trap (`WASM_ERROR` "unreachable"). Every trap-class case has its own `it()`
 * so one failure cannot hide another, and the fixtures are built byte by byte
 * here (a duplicate entry name, a forged size) rather than through the writer.
 */
import { describe, expect, it } from 'vitest';
import { readBuffer, Workbook } from '../src/index.js';
import { buildOle2 } from './helpers/ole2.js';
import { buildZip, readZipText, type ZipPart } from './helpers/zip.js';

const enc = new TextEncoder();

/** The parts of a real one-cell workbook, as bytes. */
async function baseParts(): Promise<ZipPart[]> {
  const wb = new Workbook();
  wb.addSheet('S').cell('A1').value = 'x';
  const text = readZipText(await wb.toBuffer());
  return Object.entries(text).map(([name, body]) => [name, enc.encode(body)] as const);
}

async function zipWith(name: string, data: Uint8Array = enc.encode('x')): Promise<Uint8Array> {
  return buildZip([...(await baseParts()), [name, data]]);
}

function rejection(input: unknown, options?: { password: string }) {
  return {
    toMatchObject: (match: object) =>
      expect(readBuffer(input as Uint8Array, options)).rejects.toMatchObject(match),
  };
}

const PACKAGE = new Uint8Array(32);
const B64_16 = 'AAAAAAAAAAAAAAAAAAAAAA==';
const B64_64 = `${'A'.repeat(86)}==`;

/** An Agile `EncryptionInfo` whose password key encryptor carries the given values. */
function agileInfo(spinCount: number, keyBits = 256): Uint8Array {
  const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><encryption xmlns="http://schemas.microsoft.com/office/2006/encryption" xmlns:p="http://schemas.microsoft.com/office/2006/keyEncryptor/password"><keyData saltSize="16" blockSize="16" keyBits="256" hashSize="64" cipherAlgorithm="AES" cipherChaining="ChainingModeCBC" hashAlgorithm="SHA512" saltValue="${B64_16}"/><dataIntegrity encryptedHmacKey="${B64_64}" encryptedHmacValue="${B64_64}"/><keyEncryptors><keyEncryptor uri="http://schemas.microsoft.com/office/2006/keyEncryptor/password"><p:encryptedKey spinCount="${spinCount}" saltSize="16" blockSize="16" keyBits="${keyBits}" hashSize="64" cipherAlgorithm="AES" cipherChaining="ChainingModeCBC" hashAlgorithm="SHA512" saltValue="${B64_16}" encryptedVerifierHashInput="${B64_16}" encryptedVerifierHashValue="${B64_64}" encryptedKeyValue="AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA="/></keyEncryptor></keyEncryptors></encryption>`;
  const body = enc.encode(xml);
  const out = new Uint8Array(8 + body.length);
  new DataView(out.buffer).setUint16(0, 4, true);
  new DataView(out.buffer).setUint16(2, 4, true);
  new DataView(out.buffer).setUint32(4, 0x40, true);
  out.set(body, 8);
  return out;
}

/** A Standard (4.2) `EncryptionInfo` whose `headerSize` field is hostile. */
function standardInfo(headerSize: number): Uint8Array {
  const out = new Uint8Array(8 + 4 + 200);
  const v = new DataView(out.buffer);
  v.setUint16(0, 4, true);
  v.setUint16(2, 2, true);
  v.setUint32(4, 0x24, true);
  v.setUint32(8, headerSize, true);
  return out;
}

const encrypted = (info: Uint8Array, declared?: bigint): Uint8Array =>
  buildOle2([
    declared === undefined
      ? { name: 'EncryptionInfo', data: info }
      : { name: 'EncryptionInfo', data: info, declared },
    { name: 'EncryptedPackage', data: PACKAGE },
  ]);

// ---------------------------------------------------------------------------
// ZIP
// ---------------------------------------------------------------------------

describe('hostile ZIP archives', () => {
  const unsafeNames: Array<[string, string]> = [
    ['parent traversal', '../evil.txt'],
    ['absolute path', '/evil.txt'],
    ['leading backslash', '\\evil.txt'],
    ['backslash traversal', 'a\\..\\b'],
    ['drive letter with slash', 'C:/evil.txt'],
    ['drive letter without slash', 'C:evil.txt'],
    ['NUL in the name', 'a\0b.txt'],
  ];

  for (const [label, name] of unsafeNames) {
    it(`${label} gives SECURITY`, async () => {
      await rejection(await zipWith(name)).toMatchObject({
        name: 'ModernXlsxError',
        code: 'SECURITY',
      });
    });
  }

  it('a duplicate entry name gives SECURITY', async () => {
    const parts = await baseParts();
    const workbook = parts.find(([name]) => name === 'xl/workbook.xml');
    expect(workbook).toBeDefined();
    await rejection(
      buildZip([...parts, ['xl/workbook.xml', enc.encode('<garbage')]]),
    ).toMatchObject({
      name: 'ModernXlsxError',
      code: 'SECURITY',
    });
  });

  it('a 64 MiB zero bomb gives SECURITY', async () => {
    await rejection(await zipWith('xl/media/zeros.bin', new Uint8Array(64 << 20))).toMatchObject({
      name: 'ModernXlsxError',
      code: 'SECURITY',
    });
  });

  it('a missing xl/workbook.xml gives MISSING_PART', async () => {
    const parts = (await baseParts()).filter(([name]) => name !== 'xl/workbook.xml');
    await rejection(buildZip(parts)).toMatchObject({
      name: 'ModernXlsxError',
      code: 'MISSING_PART',
    });
  });

  it('a missing [Content_Types].xml still reads (current contract)', async () => {
    const parts = (await baseParts()).filter(([name]) => name !== '[Content_Types].xml');
    const wb = await readBuffer(buildZip(parts));
    expect(wb.getSheet('S')?.cell('A1').value).toBe('x');
  });

  for (const forged of [0x80000000, 0xfffffffe]) {
    it(`a central-directory size of ${forged.toString(16)} resolves, never WASM_ERROR`, async () => {
      const wb = new Workbook();
      wb.addSheet('S').cell('A1').value = 'x';
      const zip = await wb.toBuffer();
      const view = new DataView(zip.buffer, zip.byteOffset, zip.byteLength);
      let at = 0;
      while (view.getUint32(at, true) !== 0x02014b50) at++;
      view.setUint32(at + 24, forged, true);
      const read = await readBuffer(zip);
      expect(read.getSheet('S')?.cell('A1').value).toBe('x');
    });
  }
});

// ---------------------------------------------------------------------------
// XML
// ---------------------------------------------------------------------------

describe('hostile XML', () => {
  it('a DOCTYPE with internal and external entities is never expanded or resolved', async () => {
    const dtd =
      '<!DOCTYPE x [<!ENTITY a "AAAA"><!ENTITY b "&a;&a;&a;&a;"><!ENTITY e SYSTEM "file:///etc/passwd">]>';
    const decl = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
    const main = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"';
    const sheet = `${decl}${dtd}<worksheet ${main}><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>&a;&b;&e;</t></is></c></row><row r="2"><c r="A2" t="s"><v>0</v></c></row></sheetData></worksheet>`;
    const sst = `${decl}${dtd}<sst ${main} count="1" uniqueCount="1"><si><t>&a;&b;&e;</t></si></sst>`;
    const parts = (await baseParts()).map(([name, data]): ZipPart => {
      if (name === 'xl/worksheets/sheet1.xml') return [name, enc.encode(sheet)];
      if (name === 'xl/sharedStrings.xml') return [name, enc.encode(sst)];
      return [name, data];
    });
    const wb = await readBuffer(buildZip(parts));
    // Unknown entities are dropped, so both cells are empty.
    for (const ref of ['A1', 'A2']) {
      const value = wb.getSheet('S')?.cell(ref).value;
      expect(value, ref).toBe('');
      expect(String(value)).not.toMatch(/AAAA|root:/);
    }
  });
});

// ---------------------------------------------------------------------------
// Styles
// ---------------------------------------------------------------------------

describe('hostile styles part', () => {
  const main = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"';

  /** A stored (uncompressed) package whose `xl/styles.xml` holds `count` copies of `record` inside `list`. */
  async function stylesFlood(list: string, record: string, count: number): Promise<Uint8Array> {
    const styles = `<styleSheet ${main}><${list}>${record.repeat(count)}</${list}></styleSheet>`;
    const parts = (await baseParts()).map(
      ([name, data]): ZipPart =>
        name === 'xl/styles.xml' ? [name, enc.encode(styles)] : [name, data],
    );
    return buildZip(parts, { method: 'stored' });
  }

  // Each record is counted against the read's list limit before it is kept,
  // so millions of empty records end in SIZE_LIMIT instead of trapping the
  // WASM instance (out of memory, then `unreachable`).
  for (const [list, record, count] of [
    ['dxfs', '<dxf/>', 3_300_000],
    ['dxfs', '<dxf></dxf>', 3_300_000],
    ['borders', '<border/>', 6_000_000],
    ['fills', '<fill/>', 9_000_000],
  ] as const) {
    it(`${count.toLocaleString('en-US')} ${record} give SIZE_LIMIT, never WASM_ERROR`, {
      timeout: 60_000,
    }, async () => {
      await rejection(await stylesFlood(list, record, count)).toMatchObject({
        name: 'ModernXlsxError',
        code: 'SIZE_LIMIT',
      });
    });
  }
});

// ---------------------------------------------------------------------------
// Shared strings
// ---------------------------------------------------------------------------

describe('hostile shared strings part', () => {
  const main = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"';

  // Each entry and rich text run is counted against the read's list limit
  // before it is kept, so one entry of millions of empty runs ends in
  // SIZE_LIMIT instead of trapping the WASM instance (out of memory, then
  // `unreachable`). Entry 0, which the sheet's one cell names, stays plain.
  it('one entry of 10,000,000 <r></r> gives SIZE_LIMIT, never WASM_ERROR', {
    timeout: 60_000,
  }, async () => {
    const sst = `<sst ${main}><si><t>x</t></si><si>${'<r></r>'.repeat(10_000_000)}</si></sst>`;
    const parts = (await baseParts()).map(
      ([name, data]): ZipPart =>
        name === 'xl/sharedStrings.xml' ? [name, enc.encode(sst)] : [name, data],
    );
    await rejection(buildZip(parts, { method: 'stored' })).toMatchObject({
      name: 'ModernXlsxError',
      code: 'SIZE_LIMIT',
    });
  });
});

// ---------------------------------------------------------------------------
// Charts, defined names and worksheet extensions
// ---------------------------------------------------------------------------

describe('hostile chart, workbook and worksheet parts', () => {
  const main = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"';
  const relNs = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const pkgRels = 'xmlns="http://schemas.openxmlformats.org/package/2006/relationships"';

  /** A stored package: the base parts with each of `set` replaced or added. */
  async function packageWith(set: Record<string, string>): Promise<Uint8Array> {
    const parts: ZipPart[] = (await baseParts()).filter(([name]) => !(name in set));
    for (const [name, body] of Object.entries(set)) parts.push([name, enc.encode(body)]);
    return buildZip(parts, { method: 'stored' });
  }

  /** The SIZE_LIMIT of a flood, after which the same import reads a normal file. */
  async function refusedThenReads(zip: Uint8Array): Promise<void> {
    await rejection(zip).toMatchObject({ name: 'ModernXlsxError', code: 'SIZE_LIMIT' });
    const wb = await readBuffer(await packageWith({}));
    expect(wb.getSheet('S')?.cell('A1').value).toBe('x');
  }

  // Each chart series and each defined name is counted against the read's
  // list limit before it is kept, so millions of empty records end in
  // SIZE_LIMIT instead of trapping the WASM instance. The counts are past
  // the first refused ones through the WASM build (measured 2026-11-13 by
  // bisecting the record count of a generated package; the CHANGELOG's
  // Reading bullet lists them): 2,164,798 series and 4,436,941 defined names.
  it('2,200,000 empty chart series give SIZE_LIMIT, never WASM_ERROR', {
    timeout: 60_000,
  }, async () => {
    const chartNs = 'http://schemas.openxmlformats.org/drawingml/2006/chart';
    const drawingNs = 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing';
    await refusedThenReads(
      await packageWith({
        'xl/worksheets/_rels/sheet1.xml.rels': `<Relationships ${pkgRels}><Relationship Id="rIdX1" Type="${relNs}/drawing" Target="../drawings/drawing1.xml"/></Relationships>`,
        'xl/drawings/drawing1.xml': `<wsDr xmlns="${drawingNs}" xmlns:r="${relNs}"><twoCellAnchor><from><col>0</col><colOff>0</colOff><row>0</row><rowOff>0</rowOff></from><to><col>5</col><colOff>0</colOff><row>10</row><rowOff>0</rowOff></to><graphicFrame><chart r:id="rIdC1"/></graphicFrame><clientData/></twoCellAnchor></wsDr>`,
        'xl/drawings/_rels/drawing1.xml.rels': `<Relationships ${pkgRels}><Relationship Id="rIdC1" Type="${relNs}/chart" Target="../charts/chart1.xml"/></Relationships>`,
        'xl/charts/chart1.xml': `<chartSpace xmlns="${chartNs}"><chart><plotArea><barChart>${'<ser></ser>'.repeat(2_200_000)}</barChart></plotArea></chart></chartSpace>`,
      }),
    );
  });

  it('4,500,000 defined names give SIZE_LIMIT, never WASM_ERROR', {
    timeout: 60_000,
  }, async () => {
    const workbook = new TextDecoder().decode(
      (await baseParts()).find(([name]) => name === 'xl/workbook.xml')?.[1],
    );
    await refusedThenReads(
      await packageWith({
        'xl/workbook.xml': workbook.replace(
          '</sheets>',
          `</sheets><definedNames>${'<definedName name="a"/>'.repeat(4_500_000)}</definedNames>`,
        ),
      }),
    );
  });

  // A kept worksheet <ext> is written with the declarations of the root
  // prefixes it uses, each a copy of text rendered once per sheet.
  it('2,000 kept <ext> using 200 root prefixes each read and are written namespace-well-formed', {
    timeout: 60_000,
  }, async () => {
    const k = 200;
    const m = 2_000;
    const decls = Array.from({ length: k }, (_, i) => ` xmlns:p${i}="urn:p${i}"`).join('');
    const uses = Array.from({ length: k }, (_, i) => ` p${i}:a="1"`).join('');
    const exts = Array.from({ length: m }, (_, i) => `<ext uri="u${i}"${uses}/>`).join('');
    const wb = await readBuffer(
      await packageWith({
        'xl/worksheets/sheet1.xml': `<worksheet ${main}${decls}><sheetData/><extLst>${exts}</extLst></worksheet>`,
      }),
    );
    expect(wb.toJSON().sheets[0]?.worksheet.preservedExtensions).toHaveLength(m);
    const sheet = readZipText(await wb.toBuffer())['xl/worksheets/sheet1.xml'] ?? '';
    const tags = sheet.match(/<ext\b[^>]*>/g) ?? [];
    expect(tags).toHaveLength(m);
    for (const tag of tags) {
      const declared = new Map(
        [...tag.matchAll(/ xmlns:(\w+)="([^"]*)"/g)].map(([, p, uri]) => [p, uri]),
      );
      const used = [...tag.matchAll(/ (\w+):a="1"/g)].map(([, p]) => p ?? '');
      expect(used).toHaveLength(k);
      for (const p of used) expect(declared.get(p)).toBe(`urn:${p}`);
    }
  });
});

// ---------------------------------------------------------------------------
// Input types
// ---------------------------------------------------------------------------

describe('input that is not a Uint8Array', () => {
  // readBuffer passes its argument to the WASM reader without a type check (src/index.ts:98-105), so these are the codes that reader reports.
  it('null gives WASM_ERROR', async () => {
    await rejection(null).toMatchObject({ name: 'ModernXlsxError', code: 'WASM_ERROR' });
  });

  it('an ArrayBuffer gives UNRECOGNIZED_FORMAT', async () => {
    await rejection(new ArrayBuffer(16)).toMatchObject({
      name: 'ModernXlsxError',
      code: 'UNRECOGNIZED_FORMAT',
    });
  });
});

// ---------------------------------------------------------------------------
// OLE2
// ---------------------------------------------------------------------------

describe('hostile OLE2 compound files', () => {
  it('directory size 0xFFFFFFFF (v3) gives SECURITY', async () => {
    await rejection(encrypted(agileInfo(1), 0xffffffffn)).toMatchObject({
      name: 'ModernXlsxError',
      code: 'SECURITY',
    });
  });

  it('directory size u64::MAX (v4) gives SECURITY', async () => {
    const file = buildOle2(
      [
        { name: 'EncryptionInfo', data: agileInfo(1), declared: 0xffffffffffffffffn },
        { name: 'EncryptedPackage', data: PACKAGE },
      ],
      { major: 4 },
    );
    await rejection(file).toMatchObject({ name: 'ModernXlsxError', code: 'SECURITY' });
  });

  it('Standard headerSize 0xFFFFFFFC gives PASSWORD_PROTECTED', async () => {
    await rejection(encrypted(standardInfo(0xfffffffc))).toMatchObject({
      name: 'ModernXlsxError',
      code: 'PASSWORD_PROTECTED',
    });
  });

  it('a self-looping directory chain gives SECURITY', async () => {
    const file = buildOle2(
      [
        { name: 'EncryptionInfo', data: agileInfo(1) },
        { name: 'EncryptedPackage', data: PACKAGE },
      ],
      { selfLoopDirectory: true },
    );
    await rejection(file).toMatchObject({ name: 'ModernXlsxError', code: 'SECURITY' });
  });

  it('spinCount 4294967295 gives SECURITY within 2 s', { timeout: 2000 }, async () => {
    await rejection(encrypted(agileInfo(4294967295)), { password: 'pw' }).toMatchObject({
      name: 'ModernXlsxError',
      code: 'SECURITY',
    });
  });

  it('keyBits 0xFFFFFFF8 gives PASSWORD_PROTECTED', async () => {
    await rejection(encrypted(agileInfo(1, 0xfffffff8)), { password: 'pw' }).toMatchObject({
      name: 'ModernXlsxError',
      code: 'PASSWORD_PROTECTED',
    });
  });

  it('one tampered byte in an encrypted package gives PASSWORD_PROTECTED (HMAC)', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('S');
    for (let r = 1; r <= 400; r++) ws.cell(`A${r}`).value = `row ${r} hostile-input fixture text`;
    const file = await wb.toBuffer({ password: 'pw' });
    // The package ciphertext fills most of the file, so flip one bit in the middle
    // of it; the integrity check (HMAC) runs before the package is decrypted.
    const tampered = new Uint8Array(file);
    const at = Math.floor(tampered.length / 2);
    tampered[at] = (tampered[at] ?? 0) ^ 0x01;
    await rejection(tampered, { password: 'pw' }).toMatchObject({
      name: 'ModernXlsxError',
      code: 'PASSWORD_PROTECTED',
      message: expect.stringMatching(/HMAC/),
    });
    // The untampered control still opens.
    const ok = await readBuffer(file, { password: 'pw' });
    expect(ok.getSheet('S')?.cell('A400').value).toBe('row 400 hostile-input fixture text');
  });
});

// ---------------------------------------------------------------------------
// Passwords
// ---------------------------------------------------------------------------

describe('unusual passwords', () => {
  for (const [label, password] of [
    ['a 10,000-character password', 'p'.repeat(10_000)],
    ['a password with an embedded NUL', 'a\0b'],
  ] as const) {
    it(`${label} round-trips`, async () => {
      const wb = new Workbook();
      wb.addSheet('S').cell('A1').value = 'secret';
      const file = await wb.toBuffer({ password });
      const read = await readBuffer(file, { password });
      expect(read.getSheet('S')?.cell('A1').value).toBe('secret');
    });
  }
});
