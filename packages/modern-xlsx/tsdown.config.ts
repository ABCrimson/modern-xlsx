import { defineConfig } from 'tsdown';

export default defineConfig([
  // ESM build — primary module for bundlers and Node.js
  {
    entry: ['src/index.ts'],
    format: 'esm',
    target: 'esnext',
    dts: true,
    clean: true,
    sourcemap: true,
    treeshake: true,
    deps: {
      neverBundle: [
        /\.wasm$/,
        /\/wasm\/modern_xlsx_wasm/,
      ],
    },
    outDir: 'dist',
  },
  // IIFE browser bundle — single <script> tag, exposes window.ModernXlsx
  {
    entry: { 'modern-xlsx.min': 'src/browser-entry.ts' },
    format: 'iife',
    globalName: 'ModernXlsx',
    target: 'esnext',
    minify: true,
    sourcemap: true,
    treeshake: true,
    platform: 'browser',
    outDir: 'dist',
  },
  // Web Worker script — off-thread XLSX operations
  {
    entry: { 'modern-xlsx.worker': 'src/worker.ts' },
    format: 'esm',
    target: 'esnext',
    minify: true,
    sourcemap: true,
    treeshake: true,
    platform: 'browser',
    outDir: 'dist',
  },
]);
