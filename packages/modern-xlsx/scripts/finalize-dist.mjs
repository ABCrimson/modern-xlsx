// Post-build finalization for the published package. Paths resolve against the
// package root, so the script works from any working directory.
import { copyFileSync, rmSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';

const pkg = join(import.meta.dirname, '..');

// 1. Copy the full WASM binary next to the browser bundle. tsdown already names the
//    IIFE bundle (dist/modern-xlsx.min.js and its map), so it needs no rename.
copyFileSync(join(pkg, 'wasm', 'modern_xlsx_wasm_bg.wasm'), join(pkg, 'dist', 'modern-xlsx.wasm'));

// 2. Strip wasm-pack's generated package metadata from the WASM output dirs.
//    wasm-pack writes a .gitignore containing `*` into each output dir, which the
//    packer (pnpm pack and npm pack alike) honours, so the tarball would OMIT wasm/
//    and wasm-lite/ and `import 'modern-xlsx'` / 'modern-xlsx/lite' would fail with
//    ERR_MODULE_NOT_FOUND. package.json, README.md and .cargo-ok are wasm-pack
//    metadata the package does not ship. Removing all four makes the `files`
//    allowlist authoritative.
for (const dir of ['wasm', 'wasm-lite']) {
  for (const name of ['package.json', '.gitignore', 'README.md', '.cargo-ok']) {
    rmSync(join(pkg, dir, name), { force: true });
  }
}

// 3. Hand-written declaration files for two subpaths tsdown cannot type. They are
//    written as plain strings into dist/ (never a src/ file), so tsc and tsdown never
//    see them. The package.json "exports" map points its "types" conditions at them.
//    The IIFE bundle declares the ModernXlsx global and exports nothing.
writeFileSync(
  join(pkg, 'dist', 'modern-xlsx.min.d.ts'),
  `/** Types for the IIFE build (modern-xlsx/browser). Loading it with a classic <script> or CDN tag defines globalThis.ModernXlsx. Importing this subpath as an ES module yields no exports and defines no global; import 'modern-xlsx' instead. */
import type * as Api from './index.mjs';
declare global {
  var ModernXlsx: typeof Api;
}
export {};
`,
);

//    The ./wasm subpath is the raw WASM binary itself: a compiled module on runtimes
//    and bundlers that import .wasm files, such as Cloudflare Workers.
writeFileSync(
  join(pkg, 'dist', 'wasm-module.d.ts'),
  `/**
 * The full modern-xlsx WASM binary, for runtimes and bundlers that import .wasm as a
 * compiled module (Cloudflare Workers via wrangler's CompiledWasm rule). Pass it to
 * initWasmSync().
 */
declare const wasmModule: WebAssembly.Module;
export default wasmModule;
`,
);
