#!/usr/bin/env node
/**
 * Prepares the release version-bump commit in one step, or checks that the tree is consistent.
 *
 * usage:
 *   node scripts/prepare-release.mjs <version> [--date YYYY-MM-DD]
 *     1. sets "version" in packages/modern-xlsx/package.json (exact semver x.y.z or
 *        x.y.z-pre, no build metadata, strictly greater than the current version);
 *     2. pins every modern-xlsx reference in examples/, docs/site/*.html and both READMEs
 *        (scripts/sync-example-pins.mjs), and sets the release tag of every link and image into
 *        this repository in packages/modern-xlsx/README.md (`/blob/v<version>/`, `/tree/v…/`,
 *        `raw.githubusercontent.com/ABCrimson/modern-xlsx/v…/`) to v<version>;
 *     3. renames the `## [Unreleased]` heading of packages/modern-xlsx/CHANGELOG.md to
 *        `## [<version>] - <date>` (default date: today, UTC), inserts above it a fresh
 *        `## [Unreleased]` that holds only a `<!-- release-notes:end -->` line, and updates the
 *        `[Unreleased]:` compare link if the file has link reference definitions. The section
 *        must hold one `<!-- release-notes:end -->` line with text above it: that text is the
 *        GitHub Release body (ci.yml, publish job), and in it every `npm install modern-xlsx@…`
 *        line and every release tag of a link into this repository are set to <version> /
 *        v<version>;
 *     4. prints every file it changed and the steps that remain.
 *   node scripts/prepare-release.mjs --check [<version>]
 *     Writes nothing. Exits 1 unless: VERSION in src/exports-shared.ts is read from package.json
 *     (no hard-coded version string under src/); every pin equals the package.json version;
 *     every link and image of packages/modern-xlsx/README.md into this repository names the tag
 *     v<package version> (not master, another branch or another version);
 *     the CHANGELOG has exactly one `## [Unreleased]` heading, holding one
 *     `<!-- release-notes:end -->` line whenever it has any text, and a `## [<package version>]`
 *     section, whose release notes (the text above its marker, when it has one) install and link
 *     only that version; and, when <version> is given, package.json holds that version.
 *
 * The script never runs git and never touches the network. The VERSION export needs no edit:
 * the package reads it from package.json at build time.
 */
import { globSync, readFileSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { parseArgs } from 'node:util';
import { packageVersion, syncPins } from './sync-example-pins.mjs';

const root = join(import.meta.dirname, '..');
const PACKAGE_JSON = 'packages/modern-xlsx/package.json';
const CHANGELOG = 'packages/modern-xlsx/CHANGELOG.md';
const PACKAGE_README = 'packages/modern-xlsx/README.md';
const UNRELEASED = '## [Unreleased]';

/** semver.org's regular expression for a version, without the build-metadata suffix. */
const SEMVER =
  /^(0|[1-9]\d*)\.(0|[1-9]\d*)\.(0|[1-9]\d*)(?:-((?:0|[1-9]\d*|\d*[a-zA-Z-][0-9a-zA-Z-]*)(?:\.(?:0|[1-9]\d*|\d*[a-zA-Z-][0-9a-zA-Z-]*))*))?$/;

const read = (file) => readFileSync(join(root, file), 'utf8');
const write = (file, text) => writeFileSync(join(root, file), text);

/** One pre-release identifier against another: numeric ones numerically and below alphanumeric ones. */
function compareIdentifiers(x, y) {
  const nx = /^\d+$/.test(x);
  const ny = /^\d+$/.test(y);
  if (nx && ny) return Number(x) - Number(y);
  if (nx !== ny) return nx ? -1 : 1;
  return x === y ? 0 : x < y ? -1 : 1;
}

/** Semver precedence: negative when a < b, 0 when equal, positive when a > b. */
function compareVersions(a, b) {
  const [, ...pa] = SEMVER.exec(a);
  const [, ...pb] = SEMVER.exec(b);
  for (let i = 0; i < 3; i++) {
    const d = Number(pa[i]) - Number(pb[i]);
    if (d !== 0) return d;
  }
  // A version without a pre-release ranks above the same version with one.
  if (pa[3] === undefined || pb[3] === undefined) {
    return (pa[3] === undefined ? 1 : 0) - (pb[3] === undefined ? 1 : 0);
  }
  const ia = pa[3].split('.');
  const ib = pb[3].split('.');
  for (let i = 0; i < Math.min(ia.length, ib.length); i++) {
    const d = compareIdentifiers(ia[i], ib[i]);
    if (d !== 0) return d;
  }
  return ia.length - ib.length;
}

/**
 * A link or image into this repository's files, as the package README writes them: a `blob`,
 * `tree` or `raw` URL on github.com, or a raw.githubusercontent.com URL. Group 1 is everything
 * before the git ref, group 2 the ref itself. npm renders the README once per publish and has no
 * repository context, so these URLs are absolute and name the release tag, never a branch.
 */
const REPO_FILE_URL =
  /(https:\/\/(?:github\.com\/ABCrimson\/modern-xlsx\/(?:blob|tree|raw)|raw\.githubusercontent\.com\/ABCrimson\/modern-xlsx)\/)([^/\s"'()<>]+)\//g;

/** A release tag as the repository writes it: `v` and a version (`v1.2.1`, `v1.3.0-rc.1`). */
const RELEASE_TAG = /^v\d+\.\d+\.\d+(?:-[0-9A-Za-z.-]+)?$/;

/** `text` with every release-tag segment of its repository URLs set to `v<version>`. */
function pinRepoUrls(text, version) {
  let count = 0;
  const out = text.replace(REPO_FILE_URL, (match, head, ref) => {
    if (!RELEASE_TAG.test(ref) || ref === `v${version}`) return match;
    count++;
    return `${head}v${version}/`;
  });
  return { out, count };
}

/** Each repository URL of the package README that does not name the tag `v<version>`. */
function readmeUrlProblems(text, version) {
  return [...text.matchAll(REPO_FILE_URL)]
    .filter(([, , ref]) => ref !== `v${version}`)
    .map(({ 0: match, 2: ref, index }) => {
      const line = text.slice(0, index).split('\n').length;
      return `${PACKAGE_README}:${line}: ${match} names '${ref}', expected the tag v${version}`;
    });
}

/** Line indexes of every `## [Unreleased]` heading. */
function unreleasedHeadings(lines) {
  return lines.flatMap((line, i) => (line.trimEnd() === UNRELEASED ? [i] : []));
}

/**
 * The line that ends a section's release notes: the publish job of ci.yml makes the text above
 * it the GitHub Release body, and the full list below it stays in the CHANGELOG.
 */
const NOTES_END = '<!-- release-notes:end -->';

/** An install line of the release notes: group 1 is everything before the version, group 2 the version. */
const INSTALL_LINE = /^(\s*npm install modern-xlsx@)(\S+)/;

/**
 * The section whose heading line is at `heading`: `body` is the index of its first line, `end`
 * the index of the next `## [` heading (or the line count), `markers` the indexes of its
 * NOTES_END lines.
 */
function sectionAt(lines, heading) {
  let end = lines.findIndex((line, i) => i > heading && line.startsWith('## ['));
  if (end < 0) end = lines.length;
  const markers = [];
  for (let i = heading + 1; i < end; i++) if (lines[i] === NOTES_END) markers.push(i);
  return { body: heading + 1, end, markers };
}

/** The release notes of the section at `heading`, or every problem with them, as messages. */
function notesProblems(lines, heading, version) {
  const { body, markers } = sectionAt(lines, heading);
  if (markers.length > 1) {
    return [`${CHANGELOG}:${markers[1] + 1}: a second '${NOTES_END}' line in one section`];
  }
  if (markers.length === 0) return [];
  const problems = [];
  for (let i = body; i < markers[0]; i++) {
    const install = INSTALL_LINE.exec(lines[i]);
    if (install && install[2] !== version) {
      problems.push(
        `${CHANGELOG}:${i + 1}: the install line names ${install[2]}, expected ${version}`,
      );
    }
    for (const { 0: match, 2: ref } of lines[i].matchAll(REPO_FILE_URL)) {
      if (RELEASE_TAG.test(ref) && ref !== `v${version}`) {
        problems.push(
          `${CHANGELOG}:${i + 1}: ${match} names '${ref}', expected the tag v${version}`,
        );
      }
    }
  }
  return problems;
}

/**
 * The CHANGELOG with the release heading, a fresh Unreleased section that holds only the
 * NOTES_END line (so the first entry added under it already sits above the marker), and updated
 * links.
 */
function releaseChangelog(text, version, date) {
  const lines = text.split('\n');
  const [at] = unreleasedHeadings(lines);
  lines.splice(at, 1, UNRELEASED, '', NOTES_END, '', `## [${version}] - ${date}`);
  // The release notes name the version being released: install line and tag-pinned links.
  const heading = at + 4;
  const [marker] = sectionAt(lines, heading).markers;
  for (let i = heading + 1; i < marker; i++) {
    lines[i] = pinRepoUrls(lines[i].replace(INSTALL_LINE, `$1${version}`), version).out;
  }
  let out = lines.join('\n');
  // Keep a Changelog link reference definitions, when present:
  //   [Unreleased]: <repo>/compare/v1.2.1...HEAD  becomes  <repo>/compare/v<version>...HEAD
  //   followed by [<version>]: <repo>/compare/v1.2.1...v<version>
  out = out.replace(
    /^\[Unreleased\]:\s*(\S*\/compare\/)(\S*?)\.\.\.HEAD[ \t]*$/m,
    (_match, base, from) => {
      const tag = `${from.startsWith('v') ? 'v' : ''}${version}`;
      return `[Unreleased]: ${base}${tag}...HEAD\n[${version}]: ${base}${from}...${tag}`;
    },
  );
  return out;
}

/** Every inconsistency --check reports, as messages. */
function checkTree(expected) {
  const problems = [];
  const version = packageVersion();
  if (expected !== undefined && expected !== version) {
    problems.push(`${PACKAGE_JSON}: version is ${version}, expected ${expected}`);
  }

  const exportsShared = read('packages/modern-xlsx/src/exports-shared.ts');
  if (!/^export const VERSION: string = pkg\.version;$/m.test(exportsShared)) {
    problems.push(
      'packages/modern-xlsx/src/exports-shared.ts: VERSION is not `pkg.version` (package.json)',
    );
  }
  const literal = new RegExp(`['"\`]${version.replaceAll('.', '\\.')}['"\`]`);
  for (const file of globSync('packages/modern-xlsx/src/**/*.ts', { cwd: root })) {
    if (literal.test(readFileSync(join(root, file), 'utf8'))) {
      problems.push(`${file.replaceAll('\\', '/')}: hard-codes the version ${version}`);
    }
  }

  problems.push(...syncPins({ check: true, version }).mismatches);
  problems.push(...readmeUrlProblems(read(PACKAGE_README), version));

  const lines = read(CHANGELOG).split('\n');
  const unreleased = unreleasedHeadings(lines).length;
  if (unreleased !== 1) {
    problems.push(`${CHANGELOG}: ${unreleased} '${UNRELEASED}' headings, expected exactly 1`);
  } else {
    const { body, end, markers } = sectionAt(lines, unreleasedHeadings(lines)[0]);
    const hasText = lines.slice(body, end).some((line) => line.trim() !== '');
    if (hasText && markers.length !== 1) {
      problems.push(
        `${CHANGELOG}: the '${UNRELEASED}' section has ${markers.length} '${NOTES_END}' lines, expected 1 (the publish job makes the text above it the release body)`,
      );
    }
  }
  const released = lines.findIndex((line) => line.startsWith(`## [${version}]`));
  if (released < 0) {
    problems.push(`${CHANGELOG}: no '## [${version}]' section for the package.json version`);
  } else {
    problems.push(...notesProblems(lines, released, version));
  }
  return { version, problems };
}

/** Applies the bump and returns the files it changed. */
function bump(version, date) {
  if (!SEMVER.test(version)) {
    throw new Error(`'${version}' is not an exact semver version (x.y.z or x.y.z-pre)`);
  }
  const day = new Date(`${date}T00:00:00Z`);
  if (
    !/^\d{4}-\d{2}-\d{2}$/.test(date) ||
    Number.isNaN(day.getTime()) ||
    day.toISOString().slice(0, 10) !== date
  ) {
    throw new Error(`--date '${date}' is not a calendar date in YYYY-MM-DD form`);
  }
  const current = packageVersion();
  if (compareVersions(version, current) <= 0) {
    throw new Error(`${version} is not greater than the current version ${current}`);
  }

  // Validate every file before writing any.
  const manifestText = read(PACKAGE_JSON);
  const versionField = /^(\s*"version"\s*:\s*")([^"]*)(")/m;
  if (versionField.exec(manifestText)?.[2] !== current) {
    throw new Error(`${PACKAGE_JSON}: the first "version" field is not ${current}`);
  }
  const changelog = read(CHANGELOG);
  const lines = changelog.split('\n');
  if (unreleasedHeadings(lines).length !== 1) {
    throw new Error(`${CHANGELOG}: expected exactly one '${UNRELEASED}' heading`);
  }
  if (lines.some((line) => line.startsWith(`## [${version}]`))) {
    throw new Error(`${CHANGELOG}: a '## [${version}]' section already exists`);
  }
  const { body, markers } = sectionAt(lines, unreleasedHeadings(lines)[0]);
  if (markers.length !== 1) {
    throw new Error(
      `${CHANGELOG}: the '${UNRELEASED}' section has ${markers.length} '${NOTES_END}' lines, expected 1 (the publish job makes the text above it the release body)`,
    );
  }
  if (!lines.slice(body, markers[0]).some((line) => line.trim() !== '')) {
    throw new Error(
      `${CHANGELOG}: nothing above '${NOTES_END}' in the '${UNRELEASED}' section (the publish job makes that text the release body)`,
    );
  }

  const manifest = manifestText.replace(versionField, `$1${version}$3`);
  const parsed = JSON.parse(manifest);
  if (parsed.version !== version) throw new Error(`${PACKAGE_JSON}: version rewrite failed`);
  write(PACKAGE_JSON, manifest);

  const pins = syncPins({ version });
  // After syncPins, which may have rewritten the README's CDN pins.
  const readme = pinRepoUrls(read(PACKAGE_README), version);
  if (readme.count > 0) write(PACKAGE_README, readme.out);
  write(CHANGELOG, releaseChangelog(changelog, version, date));
  const changed = [PACKAGE_JSON, ...pins.changed];
  if (readme.count > 0 && !changed.includes(PACKAGE_README)) changed.push(PACKAGE_README);
  return {
    current,
    changed: [...changed, CHANGELOG],
    pinCount: pins.mismatches.length + readme.count,
  };
}

const { values, positionals } = parseArgs({
  allowPositionals: true,
  options: {
    check: { type: 'boolean', default: false },
    date: { type: 'string' },
  },
});

if (positionals.length > 1 || (!values.check && positionals.length !== 1)) {
  console.error(
    'usage: node scripts/prepare-release.mjs <version> [--date YYYY-MM-DD]\n' +
      '       node scripts/prepare-release.mjs --check [<version>]',
  );
  process.exit(2);
}

if (values.check) {
  const { version, problems } = checkTree(positionals[0]);
  if (problems.length === 0) {
    console.log(
      `release check ok: package.json, VERSION and every pin are ${version}; every repository URL in ${PACKAGE_README} names v${version}; ${CHANGELOG} has one ${UNRELEASED} and a [${version}] section whose release notes name ${version}`,
    );
  } else {
    for (const problem of problems) console.error(problem);
    console.error(`${problems.length} release inconsistency(ies)`);
    process.exitCode = 1;
  }
} else {
  const version = positionals[0];
  const date = values.date ?? new Date().toISOString().slice(0, 10);
  try {
    const { current, changed, pinCount } = bump(version, date);
    console.log(`prepared ${current} -> ${version} (${date}); ${pinCount} pin(s) rewritten`);
    console.log('files changed:');
    for (const file of changed) console.log(`  ${file}`);
    console.log(`remaining steps:
  1. Review the diff, then build and test: pnpm build && pnpm -C packages/modern-xlsx test
  2. Run node scripts/prepare-release.mjs --check ${version}
  3. Commit the bump (release: v${version}), tag it v${version} and push the tag
  4. CI publishes to npm and creates the GitHub Release from the CHANGELOG [${version}] section
     up to ${NOTES_END}, with links to the full CHANGELOG and the compare view and the SHA-384
     integrity value of dist/modern-xlsx.min.js`);
  } catch (error) {
    console.error(`prepare-release: ${error.message}`);
    process.exitCode = 1;
  }
}
