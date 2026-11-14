#!/usr/bin/env node
/**
 * Pins every modern-xlsx reference under examples/, in docs/site/*.html (the playground's
 * CDN URL; one level only, so the generated docs/site/api/ is never scanned) and in the two
 * READMEs (README.md and packages/modern-xlsx/README.md) to exactly the version in
 * packages/modern-xlsx/package.json (no ^, ~ or = prefix):
 *
 *   - the `modern-xlsx` entry of each dependency block in every package.json under examples/
 *   - `npm:modern-xlsx@…` specifiers (deno.json imports and source files); a bare
 *     `npm:modern-xlsx` or `npm:modern-xlsx/subpath` specifier gets the version inserted
 *   - `cdn.jsdelivr.net/npm/modern-xlsx@…` and `unpkg.com/modern-xlsx@…` URLs; a bare
 *     `…/modern-xlsx/…` URL gets the version inserted
 *
 * usage: node scripts/sync-example-pins.mjs [--check]
 *   (no flag)  rewrite the files in place and list each change; running it again changes nothing
 *   --check    write nothing; list each mismatch and exit 1 if there is any
 *
 * scripts/prepare-release.mjs imports syncPins() for the release version-bump commit; CI runs
 * --check.
 */
import { globSync, readFileSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { parseArgs } from 'node:util';

const root = join(import.meta.dirname, '..');

const DEPENDENCY_BLOCKS = [
  'dependencies',
  'devDependencies',
  'peerDependencies',
  'optionalDependencies',
];
const SKIPPED_DIRS = new Set(['node_modules', '.wrangler', 'dist']);

/**
 * One rewrite rule per specifier form. `pattern` matches the whole specifier head with the
 * version (if any) in group 1; `prefix` is the head that the pinned form appends `@version` to.
 */
const SPECIFIER_RULES = [
  {
    pattern: /npm:modern-xlsx(?:@([^\s/'"`]*))?(?=[\s/'"`]|$)/g,
    prefix: 'npm:modern-xlsx',
  },
  {
    pattern: /cdn\.jsdelivr\.net\/npm\/modern-xlsx(?:@([^\s/'"`]*))?(?=[\s/'"`]|$)/g,
    prefix: 'cdn.jsdelivr.net/npm/modern-xlsx',
  },
  {
    pattern: /unpkg\.com\/modern-xlsx(?:@([^\s/'"`]*))?(?=[\s/'"`]|$)/g,
    prefix: 'unpkg.com/modern-xlsx',
  },
];

/** Line number (1-based) of a string offset. */
function lineOf(text, offset) {
  let line = 1;
  for (let i = text.indexOf('\n'); i !== -1 && i < offset; i = text.indexOf('\n', i + 1)) line++;
  return line;
}

/** Rewrites one file's text; returns the new text and the mismatches it fixed. */
function pinText(file, text, version) {
  const mismatches = [];
  let out = text;

  if (file.endsWith('package.json')) {
    const manifest = JSON.parse(text);
    const pinned = DEPENDENCY_BLOCKS.some(
      (block) => manifest[block]?.['modern-xlsx'] !== undefined,
    );
    if (pinned) {
      out = out.replace(/("modern-xlsx"\s*:\s*")([^"]*)(")/g, (match, open, found, close, at) => {
        if (found === version) return match;
        mismatches.push({ line: lineOf(text, at), found: `"modern-xlsx": "${found}"` });
        return `${open}${version}${close}`;
      });
    }
  }

  for (const { pattern, prefix } of SPECIFIER_RULES) {
    out = out.replace(pattern, (match, found, at) => {
      if (found === version) return match;
      mismatches.push({ line: lineOf(text, at), found: match });
      return `${prefix}@${version}`;
    });
  }

  return { out, mismatches };
}

/** The version in packages/modern-xlsx/package.json, read from disk at call time. */
export function packageVersion() {
  return JSON.parse(readFileSync(join(root, 'packages/modern-xlsx/package.json'), 'utf8')).version;
}

/** Every file the scan covers, repository-relative with forward slashes, sorted. */
function scannedFiles() {
  return [
    ...globSync('examples/**/*.{json,jsonc,ts,tsx,mts,js,jsx,mjs,html,vue,svelte}', {
      cwd: root,
      exclude: (entry) => SKIPPED_DIRS.has(typeof entry === 'string' ? entry : entry.name),
    }),
    ...globSync('docs/site/*.html', { cwd: root }),
    'README.md',
    'packages/modern-xlsx/README.md',
  ]
    .map((file) => file.replaceAll('\\', '/'))
    .sort();
}

/**
 * Pins every scanned reference to `version`. With `check`, writes nothing. Returns each
 * mismatch found (`file:line: found`) and the files it rewrote (none with `check`).
 */
export function syncPins({ check = false, version = packageVersion() } = {}) {
  const mismatches = [];
  const changed = [];
  for (const file of scannedFiles()) {
    const path = join(root, file);
    const text = readFileSync(path, 'utf8');
    const result = pinText(file, text, version);
    if (result.mismatches.length === 0) continue;
    for (const { line, found } of result.mismatches) {
      mismatches.push(`${file}:${line}: ${found} (expected ${version})`);
    }
    if (!check) {
      writeFileSync(path, result.out);
      changed.push(file);
    }
  }
  return { version, mismatches, changed };
}

if (import.meta.main ?? process.argv[1] === import.meta.filename) {
  const { values } = parseArgs({ options: { check: { type: 'boolean', default: false } } });
  const { version, mismatches } = syncPins({ check: values.check });
  for (const line of mismatches) console.log(line);
  if (mismatches.length === 0) {
    console.log(`examples, docs/site and READMEs: every modern-xlsx pin is ${version}`);
  } else if (values.check) {
    console.error(
      `${mismatches.length} modern-xlsx pin(s) differ from ${version}; run node scripts/sync-example-pins.mjs`,
    );
    process.exitCode = 1;
  } else {
    console.log(`pinned ${mismatches.length} reference(s) to ${version}`);
  }
}
