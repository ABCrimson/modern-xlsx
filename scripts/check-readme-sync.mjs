#!/usr/bin/env node
/**
 * The root README.md and packages/modern-xlsx/README.md (the npm page) share one Quick Start
 * block, byte for byte. Each README marks it with a line starting `<!-- quick-start:begin`
 * and a line `<!-- quick-start:end -->`; the block is everything between the two marker lines.
 *
 * It also keeps every line inside a fenced code block of either README to at most 80
 * characters (UTF-16 code units), so that no code block scrolls sideways in the npm and GitHub
 * columns on a desktop: a longer line is an error, reported as `file:line: N characters`.
 *
 * usage: node scripts/check-readme-sync.mjs [repository root]
 *   Writes nothing. Exits 0 when the two blocks are identical and no fenced line is longer
 *   than 80 characters; otherwise prints a line diff (`-` root README, `+` package README)
 *   and each long line, and exits 1. A missing, repeated or unclosed marker or code fence is
 *   also an error. The root defaults to this script's parent directory. CI runs it; fix a
 *   difference by editing both READMEs, and a long line by moving a trailing comment onto its
 *   own line above the statement or by naming an intermediate value.
 */
import { readFileSync } from 'node:fs';
import { join, resolve } from 'node:path';

const root = process.argv[2] ? resolve(process.argv[2]) : join(import.meta.dirname, '..');
const READMES = ['README.md', 'packages/modern-xlsx/README.md'];
const BEGIN = '<!-- quick-start:begin';
const END = '<!-- quick-start:end -->';
const MAX_FENCED_LINE = 80;

/** The lines inside fenced code blocks longer than MAX_FENCED_LINE, as error messages. */
function longFencedLines(file) {
  const errors = [];
  let fence = null;
  for (const [i, line] of readFileSync(join(root, file), 'utf8').split('\n').entries()) {
    const marker = /^\s*(`{3,}|~{3,})/.exec(line)?.[1];
    if (fence === null) {
      if (marker) fence = marker;
    } else if (
      marker?.[0] === fence[0] &&
      marker.length >= fence.length &&
      line.trim() === marker
    ) {
      fence = null;
    } else if (line.trimEnd().length > MAX_FENCED_LINE) {
      errors.push(
        `${file}:${i + 1}: ${line.trimEnd().length} characters in a code block (at most ${MAX_FENCED_LINE})`,
      );
    }
  }
  if (fence !== null) errors.push(`${file}: a code block opened with ${fence} is never closed`);
  return errors;
}

/** The lines between the begin and end markers of one README, or an error message. */
function quickStart(file) {
  const lines = readFileSync(join(root, file), 'utf8').split('\n');
  const begins = lines.flatMap((line, i) => (line.startsWith(BEGIN) ? [i] : []));
  const ends = lines.flatMap((line, i) => (line.trimEnd() === END ? [i] : []));
  if (begins.length !== 1 || ends.length !== 1 || ends[0] <= begins[0]) {
    return {
      error: `${file}: expected one '${BEGIN} ...' line followed by one '${END}' line (found ${begins.length} begin, ${ends.length} end)`,
    };
  }
  return { lines: lines.slice(begins[0] + 1, ends[0]) };
}

/** lcs[i][j] = length of the longest common subsequence of a.slice(i) and b.slice(j). */
function lcsTable(a, b) {
  const lcs = Array.from({ length: a.length + 1 }, () => new Array(b.length + 1).fill(0));
  for (let i = a.length - 1; i >= 0; i--) {
    for (let j = b.length - 1; j >= 0; j--) {
      lcs[i][j] = a[i] === b[j] ? lcs[i + 1][j + 1] + 1 : Math.max(lcs[i + 1][j], lcs[i][j + 1]);
    }
  }
  return lcs;
}

/** A line diff of a against b (longest common subsequence): ' ' same, '-' only a, '+' only b. */
function diffLines(a, b) {
  const lcs = lcsTable(a, b);
  const out = [];
  let i = 0;
  let j = 0;
  while (i < a.length || j < b.length) {
    if (i < a.length && j < b.length && a[i] === b[j]) {
      out.push(`  ${a[i++]}`);
      j++;
    } else if (j < b.length && (i === a.length || lcs[i][j + 1] >= lcs[i + 1][j])) {
      out.push(`+ ${b[j++]}`);
    } else {
      out.push(`- ${a[i++]}`);
    }
  }
  return out;
}

const long = READMES.flatMap(longFencedLines);
if (long.length > 0) {
  for (const error of long) console.error(error);
  process.exitCode = 1;
} else {
  console.log(`Code block lines are at most ${MAX_FENCED_LINE} characters: ${READMES.join(', ')}`);
}

const [rootBlock, packageBlock] = READMES.map(quickStart);
const errors = [rootBlock.error, packageBlock.error].filter((e) => e !== undefined);
if (errors.length > 0) {
  for (const error of errors) console.error(error);
  process.exitCode = 1;
} else if (rootBlock.lines.join('\n') === packageBlock.lines.join('\n')) {
  console.log(`Quick Start blocks match (${rootBlock.lines.length} lines): ${READMES.join(', ')}`);
} else {
  console.error(`Quick Start blocks differ (- ${READMES[0]}, + ${READMES[1]}):`);
  for (const line of diffLines(rootBlock.lines, packageBlock.lines)) console.error(line);
  process.exitCode = 1;
}
