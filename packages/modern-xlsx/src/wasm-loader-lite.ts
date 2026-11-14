/**
 * The loader of the lite WASM build (`wasm-lite/`, no encryption): its password
 * calls throw `INVALID_ARGUMENT`. Every function is documented on {@link WasmLoader}.
 */

import * as glue from '../wasm-lite/modern_xlsx_wasm.js';
import { createLoader, type WasmLoader } from './wasm-backend.js';

/** The lite build's loader, bound by the `modern-xlsx/lite` entry to its `Workbook` and `StreamingXlsxWriter`. */
export const loader: WasmLoader = createLoader(glue, 'lite');

/** Initialize the lite WASM build; see {@link WasmLoader.initWasm}. */
export const initWasm: WasmLoader['initWasm'] = loader.initWasm;
/** Initialize the lite WASM build synchronously; see {@link WasmLoader.initWasmSync}. */
export const initWasmSync: WasmLoader['initWasmSync'] = loader.initWasmSync;
/** Initialize the lite WASM build unless it already is; see {@link WasmLoader.ensureReady}. */
export const ensureReady: WasmLoader['ensureReady'] = loader.ensureReady;
