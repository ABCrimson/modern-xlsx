import { existsSync, globSync, readFileSync } from 'node:fs';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import pkg from '../package.json' with { type: 'json' };
import { VERSION } from '../src/index.js';
import { VERSION as LITE_VERSION } from '../src/index-lite.js';

const PKG_ROOT = fileURLToPath(new URL('..', import.meta.url));
const DIST = new URL('../dist/', import.meta.url);

/** Text of every built JavaScript and declaration file under dist/ (needs the build). */
function distFiles(pattern: string): { file: string; text: string }[] {
  if (!existsSync(new URL('index.mjs', DIST))) {
    throw new Error('dist/index.mjs is missing: build the package before running the tests');
  }
  return globSync(pattern, { cwd: fileURLToPath(DIST) }).map((file) => ({
    file,
    text: readFileSync(new URL(file, DIST), 'utf8'),
  }));
}

describe('VERSION', () => {
  it('equals the package.json version from both entries', () => {
    expect(VERSION).toBe(pkg.version);
    expect(LITE_VERSION).toBe(pkg.version);
  });

  it('has package.json as its only source: no file under src/ spells the version', () => {
    const literal = new RegExp(`['"\`]${pkg.version.replaceAll('.', '\\.')}['"\`]`);
    const offenders = globSync('src/**/*.ts', { cwd: PKG_ROOT })
      .filter((file) => literal.test(readFileSync(join(PKG_ROOT, file), 'utf8')))
      .map((file) => file.replaceAll('\\', '/'));
    expect(offenders).toEqual([]);
  });

  it('is the same string in the built dist/index.mjs and dist/index-lite.mjs', async () => {
    distFiles('index.mjs');
    const full: { VERSION: unknown } = await import(new URL('index.mjs', DIST).href);
    const lite: { VERSION: unknown } = await import(new URL('index-lite.mjs', DIST).href);
    expect(full.VERSION).toBe(pkg.version);
    expect(lite.VERSION).toBe(pkg.version);
  });

  it('is declared as the exact version literal in the published types', () => {
    const declarations = distFiles('*.d.mts')
      .map(({ text }) => /^declare const VERSION: (.*);$/m.exec(text)?.[1])
      .filter((type) => type !== undefined);
    expect(declarations).toEqual([JSON.stringify(pkg.version)]);
  });

  it('inlines only the version: no other package.json field reaches dist/', () => {
    const leaked = distFiles('*.{mjs,js}').filter(
      ({ text }) => text.includes('devDependencies') || text.includes('@arethetypeswrong/cli'),
    );
    expect(leaked.map(({ file }) => file)).toEqual([]);
  });
});
