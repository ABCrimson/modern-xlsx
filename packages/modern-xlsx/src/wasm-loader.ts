/**
 * The loader of the full WASM build (`wasm/`, with encryption). Every function
 * is documented on {@link WasmLoader}.
 */

import * as glue from '../wasm/modern_xlsx_wasm.js';
import { createLoader, type WasmLoader } from './wasm-backend.js';

/** The full build's loader, bound by the `modern-xlsx` entry to its `Workbook` and `StreamingXlsxWriter`. */
export const loader: WasmLoader = createLoader(glue, 'full');

/** Initialize the full WASM build; see {@link WasmLoader.initWasm}. */
export const initWasm: WasmLoader['initWasm'] = loader.initWasm;
/** Initialize the full WASM build synchronously; see {@link WasmLoader.initWasmSync}. */
export const initWasmSync: WasmLoader['initWasmSync'] = loader.initWasmSync;
/** Initialize the full WASM build unless it already is; see {@link WasmLoader.ensureReady}. */
export const ensureReady: WasmLoader['ensureReady'] = loader.ensureReady;
/**
 * Set the full build's no-argument `initWasm()` URL; see {@link WasmLoader.setDefaultWasmUrl}.
 * Internal: the browser IIFE calls it, the entries do not export it.
 */
export const setDefaultWasmUrl: WasmLoader['setDefaultWasmUrl'] = loader.setDefaultWasmUrl;
