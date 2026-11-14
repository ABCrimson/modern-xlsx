/**
 * Browser IIFE entry point.
 *
 * Bundles the full build into a single `modern-xlsx.min.js` file whose global
 * `window.ModernXlsx` exposes every runtime export of `modern-xlsx`. With no
 * argument, `initWasm()` fetches `modern-xlsx.wasm` from beside the script (or
 * against the document base URL when the script's `src` is unknown). Node file
 * I/O (`readFile`, `Workbook#toFile`) throws a `ModernXlsxError` with code
 * `IO_ERROR` here; use `readBuffer` and `toBuffer` or `writeBlob` instead.
 *
 * Usage (top-level `await` needs a module script; a classic script cannot use it):
 * ```html
 * <script src="https://cdn.jsdelivr.net/npm/modern-xlsx/dist/modern-xlsx.min.js"></script>
 * <script type="module">
 *   await ModernXlsx.initWasm();
 *   const wb = new ModernXlsx.Workbook();
 * </script>
 * ```
 */

import { setDefaultWasmUrl } from './wasm-loader.js';

// `document.currentScript` is only set while this classic script runs
// synchronously; by the time a later script calls initWasm() it is null or
// another script, so the URL is captured here, at evaluation time.
if (typeof document !== 'undefined') {
  const s = document.currentScript instanceof HTMLScriptElement ? document.currentScript.src : '';
  setDefaultWasmUrl(new URL('modern-xlsx.wasm', s || document.baseURI));
}

export * from './index.js';
