import { spawnSync } from 'node:child_process';
import {
  cpSync,
  existsSync,
  mkdirSync,
  mkdtempSync,
  readdirSync,
  readFileSync,
  rmSync,
} from 'node:fs';
import { tmpdir } from 'node:os';
import { dirname, join, resolve } from 'node:path';
import { describe, expect, it } from 'vitest';

const pkg = resolve(import.meta.dirname, '..');
const dist = join(pkg, 'dist');
const built = existsSync(join(dist, 'index-lite.mjs'));

const SPECIFIER = /(?:from|import)\s*\(?\s*['"]([^'"]+)['"]/g;

/** The dist files that the import specifier `target`, found in `file`, can load. */
function resolveSpecifier(file: string, target: string): string[] {
  if (target === '#runtime')
    return [join(dist, 'runtime-node.mjs'), join(dist, 'runtime-browser.mjs')];
  if (!target.startsWith('.')) return [];
  const p = resolve(dirname(file), target);
  return /\.(mjs|js)$/.test(p) && existsSync(p) ? [p] : [];
}

/** Every file reachable from `entry` through relative imports and '#runtime'. */
function reach(entry: string): Set<string> {
  const seen = new Set<string>();
  const queue = [entry];
  while (queue.length > 0) {
    const file = queue.pop() as string;
    if (seen.has(file)) continue;
    seen.add(file);
    for (const m of readFileSync(file, 'utf8').matchAll(SPECIFIER)) {
      queue.push(...resolveSpecifier(file, m[1] as string));
    }
  }
  return seen;
}

// These tests inspect the output of `pnpm build`; they are skipped before a build.
describe.skipIf(!built)('dist layout', () => {
  it('has no file named *.iife.*', () => {
    expect(readdirSync(dist).filter((f) => f.includes('.iife.'))).toEqual([]);
  });

  it('leaves no import.meta in the IIFE bundle', () => {
    expect(readFileSync(join(dist, 'modern-xlsx.min.js'), 'utf8')).not.toContain('import.meta');
  });

  it('names the IIFE sourcemap after the published file', () => {
    const js = readFileSync(join(dist, 'modern-xlsx.min.js'), 'utf8').trimEnd();
    expect(js.endsWith('//# sourceMappingURL=modern-xlsx.min.js.map')).toBe(true);
    const map = JSON.parse(readFileSync(join(dist, 'modern-xlsx.min.js.map'), 'utf8'));
    expect(map.file).toBe('modern-xlsx.min.js');
  });

  it('ships the full WASM binary byte-for-byte', () => {
    const shipped = readFileSync(join(dist, 'modern-xlsx.wasm'));
    const source = readFileSync(join(pkg, 'wasm', 'modern_xlsx_wasm_bg.wasm'));
    expect(shipped.equals(source)).toBe(true);
  });

  // The glue is the .js module; '#runtime' only reads the binaries (_bg.wasm) of both builds.
  it('keeps the lite entry from reaching the full WASM glue', () => {
    for (const file of reach(join(dist, 'index-lite.mjs'))) {
      expect(readFileSync(file, 'utf8')).not.toMatch(/\.\.\/wasm\/modern_xlsx_wasm\.js/);
    }
  });

  // Runs a copy of the script in a scratch package: finalize-dist rewrites dist/modern-xlsx.wasm
  // and the hand-written .d.ts files, which other suites read in parallel workers (on Windows the
  // in-place copy locks the file and a concurrent read fails with EBUSY).
  it('finalize-dist runs from any working directory', () => {
    const scratch = mkdtempSync(join(tmpdir(), 'modern-xlsx-finalize-'));
    try {
      cpSync(
        join(pkg, 'scripts', 'finalize-dist.mjs'),
        join(scratch, 'scripts', 'finalize-dist.mjs'),
      );
      cpSync(join(pkg, 'wasm'), join(scratch, 'wasm'), { recursive: true });
      mkdirSync(join(scratch, 'wasm-lite'));
      mkdirSync(join(scratch, 'dist'));
      const r = spawnSync(process.execPath, [join(scratch, 'scripts', 'finalize-dist.mjs')], {
        cwd: resolve(pkg, '..', '..'),
      });
      expect(r.status).toBe(0);
      const copied = readFileSync(join(scratch, 'dist', 'modern-xlsx.wasm'));
      expect(copied.equals(readFileSync(join(pkg, 'wasm', 'modern_xlsx_wasm_bg.wasm')))).toBe(true);
    } finally {
      rmSync(scratch, { recursive: true, force: true });
    }
  });
});
