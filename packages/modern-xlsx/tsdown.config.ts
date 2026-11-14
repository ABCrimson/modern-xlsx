import { fileURLToPath } from 'node:url';
import { defineConfig, type Rolldown, type TsdownPlugin } from 'tsdown';
import pkg from './package.json' with { type: 'json' };

/** The browser twin of the '#runtime' module, bundled into the IIFE. */
const RUNTIME_BROWSER = fileURLToPath(new URL('./src/runtime-browser.ts', import.meta.url));

/**
 * src/exports-shared.ts reads VERSION from package.json, and TypeScript types a JSON field as
 * `string`. The published declaration keeps the exact literal (`declare const VERSION: "x.y.z";`):
 * this stamps the one declaration chunk that declares it, and the build fails unless exactly one
 * does. It edits the finished chunk in generateBundle, after its declaration map is final; the
 * edit only lengthens the type at the end of that line, so the map still points at the source.
 */
const VERSION_DECLARATION = 'declare const VERSION: string;';
const versionLiteralDts: TsdownPlugin = {
  name: 'modern-xlsx:version-literal-dts',
  generateBundle(_options, bundle) {
    const chunks = Object.values(bundle).filter(
      (output): output is Rolldown.OutputChunk =>
        output.type === 'chunk' &&
        output.fileName.endsWith('.d.mts') &&
        output.code.includes(VERSION_DECLARATION),
    );
    if (chunks.length !== 1) {
      this.error(`expected one '${VERSION_DECLARATION}' to stamp, found ${chunks.length}`);
    }
    for (const chunk of chunks) {
      chunk.code = chunk.code.replace(
        VERSION_DECLARATION,
        `declare const VERSION: ${JSON.stringify(pkg.version)};`,
      );
    }
  },
};

export default defineConfig([
  // ESM build — primary module for bundlers and Node.js. Both #runtime twins are
  // entries of their own: dist/ imports '#runtime' as written, and the package.json
  // `imports` map picks dist/runtime-node.mjs (`node`) or dist/runtime-browser.mjs.
  {
    entry: {
      index: 'src/index.ts',
      'index-lite': 'src/index-lite.ts',
      cli: 'src/cli.ts',
      'runtime-node': 'src/runtime-node.ts',
      'runtime-browser': 'src/runtime-browser.ts',
    },
    plugins: [versionLiteralDts],
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
        /\/wasm-lite\/modern_xlsx_wasm/,
        /^#runtime$/,
      ],
    },
    outDir: 'dist',
  },
  // IIFE browser bundle — single <script> tag, exposes window.ModernXlsx. A classic
  // script has no `imports` map, so '#runtime' is bundled as the browser twin. This is
  // a resolveId hook, not `alias`: rolldown applies the tsconfig `paths` entry (the
  // node twin, for tsc) before `resolve.alias`.
  {
    entry: { 'modern-xlsx.min': 'src/browser-entry.ts' },
    plugins: [
      {
        name: 'modern-xlsx:runtime-browser',
        resolveId: { filter: { id: /^#runtime$/ }, handler: () => RUNTIME_BROWSER },
      },
    ],
    format: 'iife',
    globalName: 'ModernXlsx',
    target: 'esnext',
    minify: true,
    sourcemap: true,
    treeshake: true,
    platform: 'browser',
    outDir: 'dist',
    // An IIFE has no `import.meta`; the wasm-bindgen glue's `new URL(..., import.meta.url)` default
    // would make rolldown warn (EMPTY_IMPORT_META) and substitute `{}`. State that substitution here.
    // The fallback is not reached in normal use: browser-entry.ts sets the default URL from
    // `document.currentScript`, and without a document (a worker) the caller must pass the URL to
    // initWasm(); otherwise `new URL(..., undefined)` throws.
    define: { 'import.meta': '{}' },
    // tsdown appends '.iife' to the file name by default; keep the published name.
    outputOptions: { entryFileNames: '[name].js' },
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
