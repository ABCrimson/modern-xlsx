import { readFile } from 'node:fs/promises';
import { initWasmSync } from '../src/index.js';

initWasmSync(await readFile(new URL('../wasm/modern_xlsx_wasm_bg.wasm', import.meta.url)));
