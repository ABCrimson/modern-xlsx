import { fileURLToPath } from 'node:url';
import { configDefaults, defineConfig } from 'vitest/config';

export default defineConfig({
  resolve: {
    // Tests run src/ on Node: '#runtime' is the node twin (dist/ resolves it through
    // the package.json `imports` map instead).
    alias: { '#runtime': fileURLToPath(new URL('./src/runtime-node.ts', import.meta.url)) },
  },
  test: {
    // Only `vitest bench` (pnpm bench) collects these. Never set `benchmark.enabled`: it would pull
    // the benches into `vitest run` and gate the test suite on wall-clock time.
    benchmark: { include: ['__tests__/bench/**/*.bench.ts'] },
    // The benches import 'modern-xlsx', which Vite resolves as a package self-reference through
    // the package.json `exports` to dist/index.mjs (`pnpm bench` builds it first). This makes Node
    // load dist/ natively: through Vite's module runner every cross-module call goes through an
    // export getter, which skews the modern-xlsx timings (Vitest's "accessed module export getters
    // too many times" warning; guide: Benchmarking > Module Runner Overhead).
    server: { deps: { external: [/\/packages\/modern-xlsx\/dist\//] } },
    include: ['__tests__/**/*.test.ts'],
    exclude: [
      ...configDefaults.exclude,
      '__tests__/browser.test.ts',
      '__tests__/*.browser.test.ts',
    ],
    testTimeout: 10_000,
    pool: 'threads',
    setupFiles: ['__tests__/setup.ts'],
    // Persist module transforms on disk across runs.
    fsModuleCache: true,
    typecheck: {
      enabled: true,
      tsconfig: './tsconfig.test.json',
      include: ['__tests__/**/*.test-d.ts'],
    },
  },
});
