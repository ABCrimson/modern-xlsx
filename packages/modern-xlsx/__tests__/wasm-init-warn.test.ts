import { readFile } from 'node:fs/promises';
import { afterEach, describe, expect, it, vi } from 'vitest';

/**
 * The wasm-bindgen glue logs a deprecation `console.warn` when its init
 * functions get a positional argument instead of `{ module_or_path }` /
 * `{ module }`. `vi.resetModules()` before each dynamic import gives every case
 * a fresh, uninitialized glue: `__tests__/setup.ts` has already initialized the
 * shared one, and an initialized glue returns before it would warn.
 */
const CASES = [
  { name: 'wasm-loader', dir: 'wasm', load: () => import('../src/wasm-loader.js') },
  { name: 'wasm-loader-lite', dir: 'wasm-lite', load: () => import('../src/wasm-loader-lite.js') },
] as const;

describe.each(CASES)('$name init', ({ dir, load }) => {
  afterEach(() => {
    vi.restoreAllMocks();
  });

  async function fresh(): Promise<{
    loader: Awaited<ReturnType<typeof load>>;
    bytes: Uint8Array<ArrayBuffer>;
  }> {
    vi.resetModules();
    const loader = await load();
    const bytes = await readFile(new URL(`../${dir}/modern_xlsx_wasm_bg.wasm`, import.meta.url));
    return { loader, bytes };
  }

  it('initWasm(Response) does not log a deprecation warning', async () => {
    const { loader, bytes } = await fresh();
    const warn = vi.spyOn(console, 'warn');
    await loader.initWasm(new Response(bytes, { headers: { 'Content-Type': 'application/wasm' } }));
    expect(warn).not.toHaveBeenCalled();
  });

  it('initWasmSync(bytes) does not log a deprecation warning', async () => {
    const { loader, bytes } = await fresh();
    const warn = vi.spyOn(console, 'warn');
    loader.initWasmSync(bytes);
    expect(warn).not.toHaveBeenCalled();
  });
});
