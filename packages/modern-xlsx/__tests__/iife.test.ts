import { existsSync, readFileSync } from 'node:fs';
import vm from 'node:vm';
import { beforeAll, describe, expect, it } from 'vitest';
import type * as ModernXlsxApi from '../src/index.js';

/**
 * The browser IIFE (`dist/modern-xlsx.min.js`) run as a classic `<script>` in a
 * `node:vm` context with a minimal DOM. Needs the build
 * (`tsdown && node scripts/finalize-dist.mjs`) to have run first.
 */
const IIFE = new URL('../dist/modern-xlsx.min.js', import.meta.url);
const IIFE_WASM = new URL('../dist/modern-xlsx.wasm', import.meta.url);
const SCRIPT_SRC = 'https://cdn.example/pkg/dist/modern-xlsx.min.js';
const PAGE = 'https://site.example/app/index.html';

class HTMLScriptElement {
  readonly src: string;
  constructor(src: string) {
    this.src = src;
  }
}

interface Page {
  api: typeof ModernXlsxApi;
  /** Every URL the IIFE fetched, in order. */
  fetched: string[];
  document: { currentScript: HTMLScriptElement | null; baseURI: string };
}

/** The vm context's global object: the browser globals the IIFE uses, plus what it defines. */
interface Sandbox {
  [name: string]: unknown;
  ModernXlsx?: typeof ModernXlsxApi;
}

let code: string;
let wasmBytes: Uint8Array<ArrayBuffer>;

beforeAll(() => {
  if (!existsSync(IIFE) || !existsSync(IIFE_WASM)) {
    throw new Error(
      'dist/modern-xlsx.min.js is missing: build the package before running the tests',
    );
  }
  code = readFileSync(IIFE, 'utf8');
  wasmBytes = readFileSync(IIFE_WASM);
});

/**
 * Run the IIFE in a fresh context. `currentScript` is what the browser exposes
 * while the classic script runs; afterwards it is reset to `null`, as it is for
 * every later script, callback or module.
 */
function loadPage(currentScript: HTMLScriptElement | null): Page {
  const fetched: string[] = [];
  const document = { currentScript, baseURI: PAGE };
  const fetch = async (input: string | URL): Promise<Response> => {
    const url = String(input);
    fetched.push(url);
    return url.endsWith('/modern-xlsx.wasm')
      ? new Response(wasmBytes, { headers: { 'Content-Type': 'application/wasm' } })
      : new Response('not found', { status: 404 });
  };
  const ctx: Sandbox = {
    document,
    HTMLScriptElement,
    URL,
    fetch,
    Response,
    Blob,
    TextDecoder,
    TextEncoder,
    WebAssembly,
    console,
  };
  Object.assign(ctx, { globalThis: ctx, self: ctx, window: ctx });
  vm.createContext(ctx);
  vm.runInContext(code, ctx);
  document.currentScript = null;
  if (ctx.ModernXlsx === undefined) throw new Error('the IIFE did not define ModernXlsx');
  return { api: ctx.ModernXlsx, fetched, document };
}

describe('browser IIFE', () => {
  it('exposes every runtime export of the modern-xlsx entry', async () => {
    const { api } = loadPage(new HTMLScriptElement(SCRIPT_SRC));
    const esm = await import('../src/index.js');
    expect(Object.keys(api).sort()).toEqual(Object.keys(esm).sort());
  });

  it('exposes encodeQR and renderBarcodePNG', () => {
    const { api } = loadPage(new HTMLScriptElement(SCRIPT_SRC));
    expect(typeof api.encodeQR).toBe('function');
    expect(typeof api.renderBarcodePNG).toBe('function');
  });

  it('contains no node: specifier', () => {
    expect(code).not.toContain('node:');
  });

  it('initWasm() from a later script fetches modern-xlsx.wasm beside the IIFE', async () => {
    const page = loadPage(new HTMLScriptElement(SCRIPT_SRC));
    await page.api.initWasm();
    expect(page.fetched).toEqual(['https://cdn.example/pkg/dist/modern-xlsx.wasm']);
    const wb = new page.api.Workbook();
    wb.addSheet('Data').cell('A1').value = 'iife';
    const bytes = await wb.toBuffer();
    expect(Array.from(bytes.subarray(0, 2))).toEqual([0x50, 0x4b]);
  });

  it('initWasm() falls back to the document base URL without a current script', async () => {
    const page = loadPage(null);
    await page.api.initWasm();
    expect(page.fetched).toEqual(['https://site.example/app/modern-xlsx.wasm']);
  });

  it('Node file I/O throws IO_ERROR in the browser build', async () => {
    const { api } = loadPage(new HTMLScriptElement(SCRIPT_SRC));
    await api.initWasm();
    await expect(api.readFile('x.xlsx')).rejects.toMatchObject({ code: api.IO_ERROR });
    const wb = new api.Workbook();
    wb.addSheet('Data');
    await expect(wb.toFile('x.xlsx')).rejects.toMatchObject({ code: api.IO_ERROR });
  });
});
