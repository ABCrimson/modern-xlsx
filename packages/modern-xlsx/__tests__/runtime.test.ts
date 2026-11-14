import { execFile } from 'node:child_process';
import { existsSync, readFileSync } from 'node:fs';
import { mkdtemp, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { promisify } from 'node:util';
import { afterAll, beforeAll, describe, expect, it, vi } from 'vitest';
import { readBuffer, readFile, Workbook } from '../src/index.js';
import { IO_ERROR, ModernXlsxError } from '../src/model/errors.js';
import * as browserRuntime from '../src/runtime-browser.js';
import * as nodeRuntime from '../src/runtime-node.js';

const execFileAsync = promisify(execFile);

/** `\0asm`, the first four bytes of every WebAssembly binary. */
const WASM_MAGIC = [0x00, 0x61, 0x73, 0x6d];
/** `PK`, the first two bytes of a ZIP (and so of every XLSX). */
const ZIP_MAGIC = [0x50, 0x4b];
/** An Excel-authored workbook (msoffcrypto-tool test input, MIT). */
const FIXTURE = fileURLToPath(
  new URL(
    '../../../crates/modern-xlsx-core/tests/fixtures/msoffcrypto/example.xlsx',
    import.meta.url,
  ),
);
const DIST = new URL('../dist/', import.meta.url);

let tmp: string;
beforeAll(async () => {
  tmp = await mkdtemp(join(tmpdir(), 'modern-xlsx-runtime-'));
});
afterAll(async () => {
  await rm(tmp, { recursive: true, force: true });
});

describe('#runtime node twin', () => {
  it.each(['full', 'lite'] as const)(
    'loadDefaultWasm(%s) reads the package WASM binary',
    async (kind) => {
      const bytes = await nodeRuntime.loadDefaultWasm(kind);
      expect(bytes).toBeInstanceOf(Uint8Array);
      expect(Array.from(bytes?.subarray(0, 4) ?? [])).toEqual(WASM_MAGIC);
    },
  );

  it('readFileBytes returns the file bytes and writeFileBytes writes them', async () => {
    const path = join(tmp, 'bytes.bin');
    await nodeRuntime.writeFileBytes(path, Uint8Array.of(1, 2, 3));
    expect(Array.from(await nodeRuntime.readFileBytes(path))).toEqual([1, 2, 3]);
  });
});

describe('#runtime browser twin', () => {
  it('loadDefaultWasm resolves undefined so the glue resolves the binary itself', async () => {
    await expect(browserRuntime.loadDefaultWasm('full')).resolves.toBeUndefined();
    await expect(browserRuntime.loadDefaultWasm('lite')).resolves.toBeUndefined();
  });

  it('readFileBytes rejects with IO_ERROR', async () => {
    const err = await browserRuntime.readFileBytes('x.xlsx').catch((e: unknown) => e);
    expect(err).toBeInstanceOf(ModernXlsxError);
    expect(err).toMatchObject({
      code: IO_ERROR,
      message:
        'readFile is only available in Node.js-compatible runtimes; pass bytes to readBuffer instead.',
    });
  });

  it('writeFileBytes rejects with IO_ERROR', async () => {
    const err = await browserRuntime
      .writeFileBytes('x.xlsx', new Uint8Array())
      .catch((e: unknown) => e);
    expect(err).toBeInstanceOf(ModernXlsxError);
    expect(err).toMatchObject({
      code: IO_ERROR,
      message:
        'Workbook#toFile is only available in Node.js-compatible runtimes; use toBuffer instead.',
    });
  });
});

/**
 * `vi.resetModules()` before each dynamic import gives the entry a fresh,
 * uninitialized glue: `__tests__/setup.ts` has already initialized the shared one.
 */
const ENTRIES = [
  { name: 'modern-xlsx', load: () => import('../src/index.js') },
  { name: 'modern-xlsx/lite', load: () => import('../src/index-lite.js') },
] as const;

describe.each(ENTRIES)('$name initWasm() with no arguments', ({ name, load }) => {
  it('initializes from the package WASM binary on Node', async () => {
    vi.resetModules();
    const entry = await load();
    await expect(entry.initWasm()).resolves.toBeUndefined();
    const wb = new entry.Workbook();
    wb.addSheet('Data').cell('A1').value = 'node';
    const bytes = await wb.toBuffer();
    expect(Array.from(bytes.subarray(0, 2))).toEqual(ZIP_MAGIC);

    const path = join(tmp, `${name.replace('/', '-')}.xlsx`);
    await wb.toFile(path);
    const back = await entry.readFile(path);
    expect(back.sheetNames).toEqual(['Data']);
    expect(back.getSheet('Data')?.cell('A1').value).toBe('node');
  });
});

describe('Node file I/O', () => {
  it('readFile parses an Excel-authored file like readBuffer does', async () => {
    const fromFile = await readFile(FIXTURE);
    const fromBuffer = await readBuffer(readFileSync(FIXTURE));
    expect(fromFile.sheetNames).toEqual(fromBuffer.sheetNames);
    expect(fromFile.sheetNames.length).toBeGreaterThan(0);
  });

  it('toFile then readFile round-trips a workbook', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Round');
    ws.cell('A1').value = 'trip';
    ws.cell('B2').value = 42;
    const path = join(tmp, 'round-trip.xlsx');
    await wb.toFile(path);
    expect(Array.from(readFileSync(path).subarray(0, 2))).toEqual(ZIP_MAGIC);
    const back = await readFile(path);
    expect(back.sheetNames).toEqual(['Round']);
    expect(back.getSheet('Round')?.cell('A1').value).toBe('trip');
    expect(back.getSheet('Round')?.cell('B2').value).toBe(42);
  });
});

/**
 * The built package, run by a plain Node child process: `#runtime` resolves
 * through the package.json `imports` map, not the Vitest alias. Needs the build
 * (`tsdown && node scripts/finalize-dist.mjs`) to have run first.
 */
describe.each(['index.mjs', 'index-lite.mjs'])('dist/%s in a child Node process', (file) => {
  it('initWasm() with no arguments resolves and writes a workbook', async () => {
    const url = new URL(file, DIST);
    if (!existsSync(url)) {
      throw new Error(`dist/${file} is missing: build the package before running the tests`);
    }
    const script = [
      'const m = await import(process.argv[1]);',
      'await m.initWasm();',
      'const wb = new m.Workbook();',
      "wb.addSheet('S').cell('A1').value = 1;",
      'const b = await wb.toBuffer();',
      'if (b[0] !== 0x50 || b[1] !== 0x4b) process.exit(3);',
    ].join('\n');
    const { stderr } = await execFileAsync(process.execPath, [
      '--input-type=module',
      '-e',
      script,
      url.href,
    ]);
    expect(stderr).toBe('');
  });
});
