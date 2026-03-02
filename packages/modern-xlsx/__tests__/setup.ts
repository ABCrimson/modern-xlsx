import { readFile } from 'node:fs/promises';
import { beforeAll } from 'vitest';
import { initWasm } from '../src/index.js';
import { initSync } from '../wasm/modern_xlsx_wasm.js';

beforeAll(async () => {
  const wasmBytes = await readFile(new URL('../wasm/modern_xlsx_wasm_bg.wasm', import.meta.url));
  initSync({ module: wasmBytes });
  await initWasm();
});
