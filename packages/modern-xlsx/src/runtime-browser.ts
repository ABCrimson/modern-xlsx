/**
 * The `#runtime` module for browsers, workers and bundlers (the `default`
 * condition of the package.json `imports` map; the IIFE build resolves `#runtime`
 * to it with a resolveId hook, and the browser Vitest project with an alias).
 * It has the signatures of `runtime-node.ts` and imports no Node.js built-in, so
 * browser bundles never pull one in (issue #24).
 */

import { IO_ERROR, ModernXlsxError } from './model/errors.js';

/**
 * Resolve `undefined`: without file access the loader lets the wasm-bindgen glue
 * fetch the binary itself, next to the glue (`import.meta.url`).
 *
 * @param _kind - Unused; kept so both `#runtime` twins share one signature.
 */
export async function loadDefaultWasm(_kind: 'full' | 'lite'): Promise<Uint8Array | undefined> {
  return undefined;
}

/**
 * File reads need Node.js file access.
 *
 * @throws ModernXlsxError with code `IO_ERROR`, always.
 */
export async function readFileBytes(_path: string | URL): Promise<Uint8Array> {
  throw new ModernXlsxError(
    IO_ERROR,
    'readFile is only available in Node.js-compatible runtimes; pass bytes to readBuffer instead.',
  );
}

/**
 * File writes need Node.js file access.
 *
 * @throws ModernXlsxError with code `IO_ERROR`, always.
 */
export async function writeFileBytes(_path: string | URL, _data: Uint8Array): Promise<void> {
  throw new ModernXlsxError(
    IO_ERROR,
    'Workbook#toFile is only available in Node.js-compatible runtimes; use toBuffer instead.',
  );
}
