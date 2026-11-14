#!/usr/bin/env node
/**
 * Type-checks the framework snippets in examples/frameworks/{react,angular,vue,svelte}
 * against a packed modern-xlsx tarball (the CI `npm-tarball` artifact) and exact framework pins.
 *
 * usage: node scripts/examples-typecheck.mjs --tgz <path to modern-xlsx-<version>.tgz>
 *
 * Two temp projects, each with its own node_modules:
 *   ts7  The repo's TypeScript (the packages/modern-xlsx devDependency) runs `tsc --noEmit`
 *        over the React .ts/.tsx and Angular .ts snippets.
 *   ts6  vue-tsc and svelte-check run isolated on TypeScript 6.0.x: svelte-check 4.7.8 peers
 *        typescript ^5 || ^6 (vue-tsc 3.3.14 peers >=5.0.0).
 * Any diagnostic (svelte-check warnings included) fails the run.
 *
 * Type checks cannot see runtime reactivity bugs such as destructuring a Svelte getter; those
 * need a component test or a review.
 */
import assert from 'node:assert/strict';
import { spawnSync } from 'node:child_process';
import { cpSync, mkdirSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { basename, join, resolve } from 'node:path';
import { parseArgs } from 'node:util';

/**
 * Exact pins, the newest release on any channel resolved from the npm registry on 2026-11-12.
 * React 19.3.1 (latest) is newer than the 19.3.0 canaries; rxjs stays on 7.x, the range
 * @angular/core peers (9.0.0-beta.1 is outside it).
 */
const PINS = {
  react: '19.3.1',
  '@types/react': '19.3.0',
  '@angular/core': '22.3.0-rc.0',
  rxjs: '7.8.2',
  vue: '3.6.0-rc.14',
  'vue-tsc': '3.3.14',
  svelte: '5.57.3',
  'svelte-check': '4.7.8',
  /** Newest 6.0.x: svelte-check 4.7.8 peers typescript ^5 || ^6. */
  typescript6: '6.0.3',
};

const root = join(import.meta.dirname, '..');
const WINDOWS = process.platform === 'win32';

const { values } = parseArgs({ options: { tgz: { type: 'string' } } });
if (values.tgz === undefined) {
  console.error('usage: node scripts/examples-typecheck.mjs --tgz <modern-xlsx tgz>');
  process.exit(2);
}
const tgz = `file:${resolve(values.tgz).replaceAll('\\', '/')}`;
const repoTypeScript = JSON.parse(
  readFileSync(join(root, 'packages/modern-xlsx/package.json'), 'utf8'),
).devDependencies.typescript;

/** Quotes one argument for cmd.exe, which runs npm and npx (.cmd shims) on Windows. */
const quote = (arg) => (/^[\w@./:=+-]+$/.test(arg) ? arg : `"${arg}"`);

/** Runs a command to completion with inherited stdio; a non-zero exit throws. */
function run(cwd, command, ...args) {
  console.log(`$ ${[command, ...args].join(' ')}  (in ${basename(cwd)})`);
  const result = WINDOWS
    ? spawnSync([command, ...args].map(quote).join(' '), { cwd, stdio: 'inherit', shell: true })
    : spawnSync(command, args, { cwd, stdio: 'inherit' });
  if (result.error) throw result.error;
  assert.equal(result.status, 0, `${command} ${args.join(' ')} exited with ${result.status}`);
}

const writeJson = (path, value) => writeFileSync(path, `${JSON.stringify(value, null, 2)}\n`);

const COMPILER_OPTIONS = {
  target: 'esnext',
  module: 'esnext',
  moduleResolution: 'bundler',
  lib: ['esnext', 'dom'],
  jsx: 'react-jsx',
  strict: true,
  noEmit: true,
  skipLibCheck: true,
  types: [],
};

/** Creates <base>/<name> with the snippet folders, a package.json and the given tsconfigs. */
function project(base, name, { snippets, dependencies, tsconfigs }) {
  const dir = join(base, name);
  mkdirSync(dir, { recursive: true });
  for (const folder of snippets) {
    cpSync(join(root, 'examples', 'frameworks', folder), join(dir, folder), { recursive: true });
  }
  writeJson(join(dir, 'package.json'), {
    name: `modern-xlsx-examples-${name}`,
    private: true,
    type: 'module',
    dependencies: { 'modern-xlsx': tgz, ...dependencies },
  });
  for (const [file, include] of Object.entries(tsconfigs)) {
    writeJson(join(dir, file), { compilerOptions: COMPILER_OPTIONS, include });
  }
  run(dir, 'npm', 'install', '--no-audit', '--no-fund');
  return dir;
}

const base = mkdtempSync(join(process.env.RUNNER_TEMP ?? tmpdir(), 'modern-xlsx-types-'));

const ts7 = project(base, 'ts7', {
  snippets: ['react', 'angular'],
  dependencies: {
    typescript: repoTypeScript,
    react: PINS.react,
    '@types/react': PINS['@types/react'],
    '@angular/core': PINS['@angular/core'],
    rxjs: PINS.rxjs,
  },
  tsconfigs: { 'tsconfig.json': ['react', 'angular'] },
});
run(ts7, 'npx', 'tsc', '--noEmit', '-p', 'tsconfig.json');
console.log(`ok tsc ${repoTypeScript}: react + angular snippets`);

const ts6 = project(base, 'ts6', {
  snippets: ['vue', 'svelte'],
  dependencies: {
    typescript: PINS.typescript6,
    vue: PINS.vue,
    'vue-tsc': PINS['vue-tsc'],
    svelte: PINS.svelte,
    'svelte-check': PINS['svelte-check'],
  },
  tsconfigs: { 'tsconfig.vue.json': ['vue'], 'tsconfig.svelte.json': ['svelte'] },
});
run(ts6, 'npx', 'vue-tsc', '--noEmit', '-p', 'tsconfig.vue.json');
console.log(`ok vue-tsc ${PINS['vue-tsc']} (typescript ${PINS.typescript6}): vue snippets`);
run(ts6, 'npx', 'svelte-check', '--tsconfig', './tsconfig.svelte.json', '--fail-on-warnings');
console.log(
  `ok svelte-check ${PINS['svelte-check']} (typescript ${PINS.typescript6}): svelte snippets`,
);

rmSync(base, { recursive: true, force: true, maxRetries: 10, retryDelay: 500 });
console.log('ok examples-typecheck');
