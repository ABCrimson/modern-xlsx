/**
 * The `#runtime` module for Node.js-compatible runtimes (the `node` condition of
 * the package.json `imports` map). Its browser twin is `runtime-browser.ts`;
 * both export the same three functions, and only this one imports `node:` modules.
 */

import { readFile, writeFile } from 'node:fs/promises';

/**
 * Read the WASM binary of one build, so `initWasm()` with no arguments works on
 * Node, whose `fetch` cannot load the `file:` URL the wasm-bindgen glue would use.
 *
 * This module sits one level below the package root both as `src/runtime-node.ts`
 * and as the built `dist/runtime-node.mjs`, so `../wasm/` and `../wasm-lite/`
 * point at the package's `wasm/` and `wasm-lite/` directories from either.
 *
 * @param kind - `'full'` for `wasm/`, `'lite'` for `wasm-lite/`.
 */
export async function loadDefaultWasm(kind: 'full' | 'lite'): Promise<Uint8Array | undefined> {
  return readFile(
    kind === 'full'
      ? new URL('../wasm/modern_xlsx_wasm_bg.wasm', import.meta.url)
      : new URL('../wasm-lite/modern_xlsx_wasm_bg.wasm', import.meta.url),
  );
}

/** Read a file; the returned `Buffer` is already a `Uint8Array`, so it is not copied. */
export async function readFileBytes(path: string | URL): Promise<Uint8Array> {
  return readFile(path);
}

/** Write bytes to a file, replacing it if it exists. */
export async function writeFileBytes(path: string | URL, data: Uint8Array): Promise<void> {
  await writeFile(path, data);
}
