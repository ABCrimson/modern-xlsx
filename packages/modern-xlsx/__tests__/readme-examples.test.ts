/**
 * The npm README (packages/modern-xlsx/README.md) as npm shows it. npm renders the file through
 * GitHub's markdown API and only when a version is published, with no repository context: a
 * relative link or image breaks there, and a link to `master` drifts from the published version.
 * These tests pin what the page must keep for the whole life of a release: absolute links and
 * images pinned to the release tag, constructs that npm renders, table cells a phone can show,
 * in-page anchors that resolve, and the section order a newcomer reads. They also run
 * scripts/check-readme-sync.mjs, which keeps code-block lines of both READMEs within 80
 * characters, and pin the order of the root README's first sections, the labels of its link
 * row and the one page its Documentation list calls recipes.
 */
import { spawnSync } from 'node:child_process';
import { existsSync, readFileSync } from 'node:fs';
import { mkdir, mkdtempDisposable, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';

const packageDir = new URL('../', import.meta.url);
const repoDir = new URL('../../../', import.meta.url);
const readme = readFileSync(new URL('README.md', packageDir), 'utf8');
const { version } = JSON.parse(readFileSync(new URL('package.json', packageDir), 'utf8')) as {
  version: string;
};

/** The README lines outside fenced code blocks, with their 1-based line numbers. */
function proseLines(text: string): { line: number; text: string }[] {
  const out: { line: number; text: string }[] = [];
  let fenced = false;
  for (const [i, line] of text.split('\n').entries()) {
    if (/^\s*```/.test(line)) {
      fenced = !fenced;
      continue;
    }
    if (!fenced) out.push({ line: i + 1, text: line });
  }
  return out;
}

const prose = proseLines(readme);

/** Every link and image target in the prose: Markdown `](...)`, and `href`, `src`, `srcset`. */
function targets(): { line: number; url: string }[] {
  const found: { line: number; url: string }[] = [];
  for (const { line, text } of prose) {
    const withoutCode = text.replace(/`[^`]*`/g, '');
    for (const m of withoutCode.matchAll(/\]\(([^)\s]+)\)/g)) found.push({ line, url: m[1] ?? '' });
    for (const m of withoutCode.matchAll(/\b(?:href|src|srcset)="([^"]+)"/g)) {
      found.push({ line, url: m[1] ?? '' });
    }
  }
  return found;
}

/** GitHub's heading anchor: lower case, punctuation dropped, spaces to hyphens. */
function slug(heading: string): string {
  return heading
    .trim()
    .toLowerCase()
    .replace(/<[^>]+>/g, '')
    .replace(/[^\p{L}\p{N}\s_-]/gu, '')
    .replace(/\s/g, '-');
}

/** The ATX headings of the prose, in order. */
const headings = prose.flatMap(({ line, text }) => {
  const m = /^(#{1,6})\s+(.+?)\s*#*$/.exec(text);
  return m?.[1] && m[2] ? [{ line, level: m[1].length, title: m[2] }] : [];
});

describe('npm README: links and images', () => {
  it('uses absolute URLs or in-page anchors only', () => {
    const relative = targets().filter(({ url }) => !/^(?:https:\/\/|#)/.test(url));
    expect(relative).toEqual([]);
  });

  it(`pins every link and image into this repository to the tag v${version}`, () => {
    const repoFile =
      /^https:\/\/(?:github\.com\/ABCrimson\/modern-xlsx\/(?:blob|tree)|raw\.githubusercontent\.com\/ABCrimson\/modern-xlsx)\/([^/]+)\//;
    const unpinned = targets().filter(({ url }) => {
      const ref = repoFile.exec(url)?.[1];
      return ref !== undefined && ref !== `v${version}`;
    });
    expect(unpinned).toEqual([]);
    expect(readme).not.toMatch(/ABCrimson\/modern-xlsx\/(?:blob\/)?master\//);
  });

  it('points every pinned link and image at a file of this tree', () => {
    const pinned = new RegExp(
      `^https://(?:github\\.com/ABCrimson/modern-xlsx/(?:blob|tree)|raw\\.githubusercontent\\.com/ABCrimson/modern-xlsx)/v${version.replaceAll('.', '\\.')}/([^#?]+)`,
    );
    const paths = targets().flatMap(({ url }) => pinned.exec(url)?.[1] ?? []);
    expect(paths.length).toBeGreaterThan(10);
    const missing = paths.filter((path) => !existsSync(new URL(path, repoDir)));
    expect(missing).toEqual([]);
  });

  it('resolves every in-page anchor to a heading', () => {
    const anchors = new Set(headings.map(({ title }) => slug(title)));
    const dangling = targets().filter(
      ({ url }) => url.startsWith('#') && !anchors.has(url.slice(1)),
    );
    expect(dangling).toEqual([]);
  });

  it('shows the hero and the architecture picture with a light fallback image', () => {
    const base = `https://raw.githubusercontent.com/ABCrimson/modern-xlsx/v${version}/docs/assets/readme/`;
    for (const name of ['hero', 'architecture']) {
      expect(readme).toContain(`srcset="${base}${name}-dark.svg"`);
      expect(readme).toContain(`src="${base}${name}-light.svg"`);
    }
    const alt = readFileSync(new URL('docs/assets/readme/architecture-alt.txt', repoDir), 'utf8');
    expect(readme).toContain(`alt="${alt.trim()}"`);
  });
});

describe('npm README: constructs npm renders', () => {
  it('has no GitHub alert, Mermaid block, <details> or <h1> inside <p>', () => {
    expect(readme).not.toMatch(/^>\s*\[!/m);
    expect(readme).not.toMatch(/^```mermaid/m);
    expect(readme).not.toMatch(/<details/);
    expect(readme).not.toMatch(/<p[^>]*>\s*<h1/);
  });

  it('keeps every table cell to at most 160 characters', () => {
    const long = prose
      .filter(({ text }) => text.startsWith('|'))
      .flatMap(({ line, text }) =>
        text
          .replace(/\\\|/g, '')
          .split('|')
          .slice(1, -1)
          .map((cell) => ({ line, length: cell.trim().length }))
          .filter(({ length }) => length > 160),
      );
    expect(long).toEqual([]);
  });
});

describe('npm README: structure', () => {
  it('reaches Install and Quick start before any table, blockquote or reference section', () => {
    const quickStart = headings.find(({ title }) => title === 'Quick start');
    const install = headings.find(({ title }) => title === 'Install');
    expect(install?.level).toBe(2);
    expect(quickStart?.level).toBe(2);
    if (!quickStart || !install) return;
    expect(install.line).toBeLessThan(quickStart.line);
    expect(quickStart.line).toBeLessThan(50);
    const before = prose.filter(({ line }) => line < quickStart.line);
    expect(before.filter(({ text }) => /^\s*>/.test(text) || text.startsWith('|'))).toEqual([]);
  });

  it('files the usage sections under the API overview, not under a types heading', () => {
    const parent = (title: string) => {
      const at = headings.findIndex((h) => h.title === title);
      return headings
        .slice(0, at)
        .reverse()
        .find((h) => h.level === 2)?.title;
    };
    for (const title of [
      'Pivot tables',
      'Threaded comments',
      'Slicers and timelines',
      'Tables',
      'Page layout and printing',
    ]) {
      expect(parent(title), title).toBe('API overview');
    }
    expect(parent('Write a large file row by row')).toBe('Common tasks');
    expect(parent('Cloudflare Workers')).toBe('Runtimes and entry points');
    expect(headings.filter(({ title }) => /^Types$/.test(title))).toEqual([]);
  });

  it('prints the benchmark in the format the bench generator emits', () => {
    // Padded as Biome formats Markdown tables; the bench generator pads the same way.
    // The unit sits in the row label, so a phone column never breaks a figure from its unit.
    expect(readme).toMatch(/^\| 10 columns +\| modern-xlsx \| +SheetJS \| Ratio \|$/m);
    expect(readme).toContain(
      'Ratio is the SheetJS figure divided by the modern-xlsx figure: above 1x favours modern-xlsx.',
    );
    expect(readme).toMatch(/^\| Write 10K rows \(ms\) +\| +[\d,.]+ \| +[\d,.]+ \| 0\.\d\dx \|$/m);
    expect(readme).toMatch(
      /^\| File size, 10K rows \(KB\) \| +[\d,]+ \| +[\d,]+ \| \d\.\d\dx \|$/m,
    );
    expect(readme).not.toMatch(/\d (?:ms|KB) +\|/);
    expect(readme).not.toMatch(/^As of \d{4}-/m);
    expect(readme).not.toMatch(/\bcommit [0-9a-f]{7,40}\b/);
  });

  it('holds the header, row labels and ratio line the bench generator prints', () => {
    const bench = readFileSync(new URL('__tests__/bench/read-write.bench.ts', packageDir), 'utf8');
    const ratio =
      'Ratio is the SheetJS figure divided by the modern-xlsx figure: above 1x favours modern-xlsx.';
    expect(bench).toContain("['10 columns', 'modern-xlsx', 'SheetJS', 'Ratio']");
    expect(bench).toContain(`'${ratio}'`);
    expect(readme).toContain(ratio);
    for (const label of [
      'Read 10K rows (ms)',
      'Read 100K rows (ms)',
      'Write 10K rows (ms)',
      'File size, 10K rows (KB)',
    ]) {
      expect(bench, label).toContain(`'${label}'`);
      expect(readme, label).toMatch(new RegExp(`^\\| ${RegExp.escape(label)} +\\|`, 'm'));
    }
  });

  it('lists the entry points as bullets, so no table cell splits a specifier on a phone', () => {
    const at = prose.findIndex(({ text }) => text === '## Runtimes and entry points');
    const end = prose.findIndex(({ text }, i) => i > at && text.startsWith('### '));
    const section = prose.slice(at, end).map(({ text }) => text);
    expect(section.filter((text) => text.startsWith('|'))).toEqual([]);
    for (const entry of ['', '/lite', '/browser', '/worker', '/wasm']) {
      expect(section.some((text) => text.startsWith(`- **\`modern-xlsx${entry}\`:**`))).toBe(true);
    }
  });

  it('initializes WASM in the read and encryption examples, which run as copied', () => {
    const blockAfter = (heading: string) => {
      const from = readme.indexOf(`### ${heading}\n`);
      return /```typescript\n([\s\S]*?)```/.exec(readme.slice(from))?.[1] ?? '';
    };
    for (const heading of ['Read a file', 'Encrypt a workbook']) {
      expect(blockAfter(heading), heading).toMatch(
        /^import \{ initWasm, [^}]+\} from 'modern-xlsx';$/m,
      );
      expect(blockAfter(heading), heading).toMatch(/^await initWasm\(\);$/m);
    }
  });

  it('says which examples initialize WASM themselves and where the save example gets wb', () => {
    const at = readme.indexOf('## Common tasks\n');
    const intro = readme.slice(at).split('\n')[2];
    expect(intro).toBe(
      'Call `await initWasm()` once before these examples, as in the Quick start. ' +
        'The save example uses the `wb` built in [Style a header row](#style-a-header-row).',
    );
    const save = /### Save a file[^\n]*\n\n```typescript\n([\s\S]*?)```/.exec(readme)?.[1] ?? '';
    expect(save).toMatch(/\bwb\./);
    expect(save).not.toMatch(/\bconst wb\b/);
    const style = /### Style a header row\n\n```typescript\n([\s\S]*?)```/.exec(readme)?.[1] ?? '';
    expect(style).toMatch(/^const wb = new Workbook\(\);$/m);
  });

  it('says a ZIP or JSON that exceeds a limit rejects, with the code each limit throws', () => {
    expect(readme).toContain(
      'a ZIP that exceeds the size, compression-ratio or entry-count limit rejects with ' +
        '`SECURITY`, and a read whose workbook JSON exceeds `maxJsonLength` rejects with ' +
        '`SIZE_LIMIT`',
    );
    expect(readme).not.toMatch(/\bpasses (?:the size|`maxJsonLength`)/);
  });

  it('keeps imports at the top of every TypeScript example', () => {
    const blocks = readme.matchAll(/```typescript\n([\s\S]*?)```/g);
    const late = [...blocks].flatMap(([, code = '']) => {
      const lines = code.split('\n');
      const lastImport = lines.findLastIndex((l) => /^import\b/.test(l));
      const firstCode = lines.findIndex(
        (l) => l.trim() !== '' && !/^(?:import\b|\s|\}|\/\/)/.test(l),
      );
      return firstCode !== -1 && lastImport > firstCode ? [lines[lastImport]] : [];
    });
    expect(late).toEqual([]);
  });
});

describe('scripts/check-readme-sync.mjs', () => {
  const script = fileURLToPath(new URL('scripts/check-readme-sync.mjs', repoDir));
  const check = (root?: string) =>
    spawnSync(process.execPath, root ? [script, root] : [script], { encoding: 'utf8' });

  it('passes on this tree', () => {
    const r = check();
    expect(r.stderr).toBe('');
    expect(r.status).toBe(0);
    expect(r.stdout).toContain('Code block lines are at most 80 characters');
  });

  it('fails on an 81-character line in a code block and passes on an 80-character one', async () => {
    await using dir = await mkdtempDisposable(join(tmpdir(), 'modern-xlsx-readme-'));
    const block = [
      '<!-- quick-start:begin -->',
      '```ts',
      'quick();',
      '```',
      '<!-- quick-start:end -->',
    ];
    const write = async (codeLine: string) => {
      await mkdir(join(dir.path, 'packages', 'modern-xlsx'), { recursive: true });
      await writeFile(join(dir.path, 'README.md'), `${block.join('\n')}\n`);
      const pkg = [...block, '', '```js', codeLine, '```', ''].join('\n');
      await writeFile(join(dir.path, 'packages', 'modern-xlsx', 'README.md'), pkg);
    };
    await write(`// ${'x'.repeat(77)}`);
    expect(check(dir.path).status).toBe(0);
    await write(`// ${'x'.repeat(78)}`);
    const r = check(dir.path);
    expect(r.status).toBe(1);
    expect(r.stderr).toContain('packages/modern-xlsx/README.md:8: 81 characters in a code block');
  });
});

describe('root README: structure', () => {
  const root = readFileSync(new URL('README.md', repoDir), 'utf8');
  const order = [...root.matchAll(/^## (.+)$/gm)].map((m) => m[1]);

  it('orders Install, Quick start and Why as the package README does', () => {
    const at = (title: string) => order.indexOf(title);
    expect(at('Install')).toBeGreaterThanOrEqual(0);
    expect(at('Install')).toBeLessThan(at('Why modern-xlsx'));
    expect(headings.findIndex(({ title }) => title === 'Install')).toBeLessThan(
      headings.findIndex(({ title }) => title === 'Why modern-xlsx'),
    );
  });

  it('links the recipes page and the runnable examples separately', () => {
    expect(root).toContain('[Recipes](./docs/examples.md)');
    expect(root).toContain('[runnable examples](./examples/README.md)');
  });

  it('labels the link to docs/examples.md Recipes in both link rows', () => {
    const linkRow = (text: string) => text.split('\n').find((l) => l.includes('&emsp; [')) ?? '';
    const labels = (row: string) => [...row.matchAll(/\[([^\]]+)\]\(/g)].map((m) => m[1]);
    const rootRow = linkRow(root);
    const packageRow = linkRow(readme);
    expect(rootRow).toContain('[Recipes](./docs/examples.md)');
    expect(packageRow).toContain(
      `[Recipes](https://github.com/ABCrimson/modern-xlsx/blob/v${version}/docs/examples.md)`,
    );
    expect(labels(rootRow).slice(1)).toEqual(labels(packageRow).slice(0, -1));
  });

  it('uses the word recipes for one page only in the Documentation list', () => {
    const at = root.indexOf('## Documentation\n');
    const list = root.slice(at, root.indexOf('\n## ', at));
    const recipes = list.split('\n').filter((l) => /recipes/i.test(l));
    expect(recipes).toEqual([expect.stringMatching(/^- \[Recipes\]\(\.\/docs\/examples\.md\):/)]);
  });
});
