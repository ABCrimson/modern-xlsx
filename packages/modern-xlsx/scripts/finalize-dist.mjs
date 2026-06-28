// Post-build finalization for the published package.
import { copyFileSync, existsSync, renameSync, rmSync } from 'node:fs';

// 1. Copy the full WASM binary next to the browser bundle.
copyFileSync('wasm/modern_xlsx_wasm_bg.wasm', 'dist/modern-xlsx.wasm');

// 2. Rename the IIFE browser bundle to its published name.
renameSync('dist/modern-xlsx.min.iife.js', 'dist/modern-xlsx.min.js');
renameSync('dist/modern-xlsx.min.iife.js.map', 'dist/modern-xlsx.min.js.map');

// 3. Strip wasm-pack's generated package metadata from the WASM output dirs.
//    npm pack skips any nested directory that contains its own package.json
//    (it treats it as a separate package), and wasm-pack also writes a
//    `.gitignore` containing `*`. Either one causes the published tarball to
//    OMIT wasm/ and wasm-lite/, so `import 'modern-xlsx'` and 'modern-xlsx/lite'
//    fail at module resolution with ERR_MODULE_NOT_FOUND. Removing these makes
//    the `files` allowlist authoritative and ships the WASM glue.
for (const dir of ['wasm', 'wasm-lite']) {
  for (const name of ['package.json', '.gitignore', 'README.md', '.cargo-ok']) {
    const p = `${dir}/${name}`;
    if (existsSync(p)) rmSync(p);
  }
}
