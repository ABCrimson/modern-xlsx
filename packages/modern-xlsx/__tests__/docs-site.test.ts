import { readFileSync } from 'node:fs';
import vm from 'node:vm';
import { describe, expect, it, vi } from 'vitest';
import * as ModernXlsx from '../src/index.js';

/**
 * The static docs site (docs/site/) and the service-worker example
 * (examples/edge/service-worker/sw.js) are not part of the package build, so their
 * checks read the files from the repository. The DOM and layout checks live in
 * docs-site.browser.test.ts.
 */
const repoFile = (path: string): string =>
  readFileSync(new URL(`../../../${path}`, import.meta.url), 'utf8');

const playground = repoFile('docs/site/playground.html');
const index = repoFile('docs/site/index.html');

/** A playground example's source, as the page puts it in the editor. */
function exampleSource(name: string): string {
  const match = new RegExp(`\\n {4}${name}: \`([^\`]*)\``).exec(playground);
  if (!match?.[1]) throw new Error(`playground example '${name}' not found`);
  return match[1];
}

/** WCAG 2.x relative luminance of a #rrggbb colour. */
function luminance(hex: string): number {
  const channels = [1, 3, 5].map((i) => Number.parseInt(hex.slice(i, i + 2), 16) / 255);
  const [r = 0, g = 0, b = 0] = channels.map((c) =>
    c <= 0.04045 ? c / 12.92 : ((c + 0.055) / 1.055) ** 2.4,
  );
  return 0.2126 * r + 0.7152 * g + 0.0722 * b;
}

function contrast(a: string, b: string): number {
  const [hi, lo] = [luminance(a), luminance(b)].sort((x, y) => y - x);
  return ((hi ?? 0) + 0.05) / ((lo ?? 0) + 0.05);
}

/** The light and dark values of a `--name: light-dark(#rrggbb, #rrggbb);` token in a page's `:root`. */
function token(html: string, name: string): { light: string; dark: string } {
  const match = new RegExp(`--${name}: light-dark\\((#[0-9a-f]{6}), (#[0-9a-f]{6})\\);`).exec(html);
  if (!match?.[1] || !match[2]) throw new Error(`token --${name} not found`);
  return { light: match[1], dark: match[2] };
}

describe('docs/site colour contrast (WCAG AA, 4.5:1)', () => {
  it.each([
    ['index.html', index],
    ['playground.html', playground],
  ])('--text-dim in %s reads on every surface of both themes', (_page, html) => {
    const dim = token(html, 'text-dim');
    for (const surface of ['#0a0a0a', '#111111', '#1a1a1a', '#0d0d0d']) {
      expect(contrast(dim.dark, surface), `${dim.dark} on ${surface}`).toBeGreaterThanOrEqual(4.5);
    }
    for (const surface of ['#ffffff', '#f8fafc', '#eef2f7', '#f1f5f9']) {
      expect(contrast(dim.light, surface), `${dim.light} on ${surface}`).toBeGreaterThanOrEqual(
        4.5,
      );
    }
  });

  it.each([
    ['index.html', index],
    ['playground.html', playground],
  ])('the primary button label in %s reads on the crimson accent and its hover', (_page, html) => {
    expect(html).toContain(
      '.btn-primary { background: var(--button); color: var(--on-accent); border-color: var(--button); }',
    );
    expect(html).toContain('.btn-primary:hover { background: var(--button-hover);');
    const label = token(html, 'on-accent');
    expect(label).toEqual({ light: '#ffffff', dark: '#0a0a0a' });
    for (const name of ['button', 'button-hover']) {
      const fill = token(html, name);
      expect(contrast(label.light, fill.light), `${name} light`).toBeGreaterThanOrEqual(4.5);
      expect(contrast(label.dark, fill.dark), `${name} dark`).toBeGreaterThanOrEqual(4.5);
    }
  });

  it.each([
    ['index.html', index],
    ['playground.html', playground],
  ])('accent text and link text in %s read on the page in both themes', (_page, html) => {
    for (const name of ['accent', 'link']) {
      const colour = token(html, name);
      expect(contrast(colour.light, '#ffffff'), `${name} light`).toBeGreaterThanOrEqual(4.5);
      expect(contrast(colour.dark, '#0a0a0a'), `${name} dark`).toBeGreaterThanOrEqual(4.5);
    }
  });
});

describe('docs/site identity: one crimson accent, as in docs/assets/readme/README.md', () => {
  it.each([
    ['index.html', index],
    ['playground.html', playground],
  ])('%s takes its accent and primary button from the brand crimson', (_page, html) => {
    expect(token(html, 'accent')).toEqual({ light: '#c8102e', dark: '#ff6b7a' });
    expect(token(html, 'button')).toEqual({ light: '#c8102e', dark: '#ff6b7a' });
    // Links keep their own colour, so the accent stays for the primary action.
    expect(token(html, 'link')).not.toEqual(token(html, 'accent'));
    expect(html).toContain('a { color: var(--link);');
  });

  it.each([
    ['index.html', index],
    ['playground.html', playground],
  ])('%s has no gradient text and no third button colour', (_page, html) => {
    expect(html).not.toMatch(/background-clip:\s*text/);
    expect(html).not.toMatch(/text-fill-color/);
    expect(html).not.toContain('--run:');
    expect(html).not.toContain('.gradient');
  });

  it('the stat numerals are ink, not three colours', () => {
    expect(index).toContain('.stat .number {');
    expect(index).not.toMatch(/class="number [a-z]/);
    expect(index).not.toMatch(/\.stat \.number\.[a-z]+ \{/);
  });

  it('the playground runs from the primary button and downloads from a secondary one', () => {
    expect(playground).toMatch(/<button type="button" id="run-btn" class="btn btn-primary"/);
    expect(playground).toContain(
      '.download-link {\n      display: inline-flex;\n      align-items: center;\n      gap: 0.4rem;\n      margin-top: 0.5rem;\n      padding: 0.4rem 0.8rem;\n      background: var(--surface-2);\n      color: var(--text);\n      border: 1px solid var(--border);',
    );
  });

  it('feature, section and table titles are in sentence case, as in the READMEs', () => {
    const features = [...index.matchAll(/<h3>([^<]+)<\/h3>/g)].map((m) => m[1]);
    expect(features).toEqual([
      'Cell styling',
      'Number formats',
      'Data validation',
      'Conditional formatting',
      'Charts',
      'Images and barcodes',
      'Pivot tables',
      'Formula engine',
      'Streaming writer',
      'Frozen panes',
      'Hyperlinks',
      'Comments and notes',
      'Sheet protection',
      'Encryption',
      'Page setup',
      'Dates and Temporal',
      'Rust and WebAssembly core',
      'Tree-shakable ESM',
    ]);
    const options = [...playground.matchAll(/<option value="[a-z]+">([^<]+)<\/option>/g)].map(
      (m) => m[1],
    );
    expect(options).toEqual([
      'Hello world',
      'Styled workbook',
      'Read and inspect',
      'Batch data (1,000 rows)',
      'JSON to sheet',
      'Formulas and dates',
      'Barcodes and QR codes',
    ]);
    expect(index).toContain('<td>Cell references</td>');
    expect(index).toContain('<td>Rich text</td>');
  });
});

describe('docs/site/playground.html start-up and preview', () => {
  const script = playground.slice(playground.indexOf('(async () => {'));

  it('puts the Hello example in the editor and wires the menu before it waits for the library', () => {
    const init = script.indexOf('await ModernXlsx.initWasm()');
    expect(init).toBeGreaterThan(0);
    expect(script.indexOf("loadExample('hello');")).toBeGreaterThan(0);
    expect(script.indexOf("loadExample('hello');")).toBeLessThan(init);
    expect(script.indexOf("examplesSelect.addEventListener('change'")).toBeLessThan(init);
    // The first automatic run happens only while the editor still holds the untouched example.
    expect(script).toContain(
      'if (!runBtn.disabled && editor.value === examples.hello) await runCode(false);',
    );
  });

  it('builds the preview table with DOM methods and textContent, never innerHTML from cell text', () => {
    // From the cell text helpers to the end of window.preview.
    const preview = script.slice(
      script.indexOf('function numberText('),
      script.indexOf('// Open a local .xlsx'),
    );
    expect(preview).toContain('window.preview = (ws, maxRows = 1000) => {');
    expect(preview).toContain('cell.textContent = ');
    expect(preview).not.toContain('sheetToHtml');
    expect(playground).not.toContain('innerHTML');
    expect(playground).not.toContain('insertAdjacentHTML');
  });

  it("scrolls a result into view only on the stylesheet's phone layout, without motion when reduced motion is preferred", () => {
    // The script's breakpoint is the stylesheet's one-column breakpoint.
    expect(playground).toContain('@media (max-width: 768px) {');
    expect(script).toContain("const narrowLayout = matchMedia('(max-width: 768px)');");
    expect(script).toContain(
      "const reducedMotion = matchMedia('(prefers-reduced-motion: reduce)');",
    );
    const reveal = script.slice(script.indexOf('function revealNewest()'));
    expect(reveal).toMatch(
      /^function revealNewest\(\) \{\n {4}if \(!narrowLayout\.matches\) return;/,
    );
    expect(reveal).toContain("behavior: reducedMotion.matches ? 'auto' : 'smooth',");
  });
});

describe('docs/site architecture picture', () => {
  const alt = repoFile('docs/assets/readme/architecture-alt.txt').trim();

  it('shows the committed picture in the visitor theme, sized and with the full alt text', () => {
    const picture = /<picture class="arch-picture">([\s\S]*?)<\/picture>/.exec(index)?.[1] ?? '';
    expect(picture).toContain(
      '<source media="(prefers-color-scheme: dark)" srcset="architecture-dark.svg" width="960" height="1040">',
    );
    expect(picture).toContain(
      `<img class="arch-img" src="architecture-light.svg" width="960" height="1040" loading="lazy" decoding="async" alt="${alt}">`,
    );
    // No HTML box diagram (class="layer") remains beside the picture.
    expect(index).not.toContain('class="layer');
  });
});

describe('docs/site brand files', () => {
  /** pages.yml publishes only docs/site/, so these are committed copies of docs/assets/readme/. */
  const brandFiles = [
    'favicon.svg',
    'og.png',
    'mark-light.svg',
    'mark-dark.svg',
    'architecture-light.svg',
    'architecture-dark.svg',
  ];
  const bytes = (path: string): Buffer =>
    readFileSync(new URL(`../../../${path}`, import.meta.url));

  it.each(brandFiles)('docs/site/%s is byte-identical to docs/assets/readme/', (file) => {
    expect(bytes(`docs/site/${file}`).equals(bytes(`docs/assets/readme/${file}`))).toBe(true);
  });
});

describe('playground JSON to Sheet example', () => {
  const employees = [
    { name: 'Alice', department: 'Engineering', salary: 120000, startDate: '2023-01-15' },
    { name: 'Bob', department: 'Marketing', salary: 95000, startDate: '2023-03-22' },
    { name: 'Carol', department: 'Engineering', salary: 135000, startDate: '2022-11-01' },
    { name: 'Dave', department: 'Sales', salary: 88000, startDate: '2024-06-10' },
    { name: 'Eve', department: 'Engineering', salary: 142000, startDate: '2021-08-05' },
  ];

  it('uses sheetAddJson instead of hand-written cell loops', () => {
    const source = exampleSource('json');
    expect(source).toContain('ModernXlsx.sheetAddJson(ws, employees);');
    expect(source).not.toContain('encodeCellRef');
    expect(source).not.toContain('Object.keys');
  });

  it('runs as the playground runs it and round-trips 5 rows with the same keys', async () => {
    const logged: string[] = [];
    const download = vi.fn();
    const preview = vi.fn();
    // The playground's runner: the example body inside an async function with these four
    // parameters (runCode in playground.html).
    const run = vm.compileFunction(`return (async () => {\n${exampleSource('json')}\n})();`, [
      'ModernXlsx',
      'log',
      'download',
      'preview',
    ]) as (
      api: typeof ModernXlsx,
      log: (msg: string) => void,
      dl: unknown,
      pv: unknown,
    ) => Promise<void>;
    await run(ModernXlsx, (msg) => logged.push(msg), download, preview);
    expect(preview).toHaveBeenCalledOnce();

    expect(logged).toContain('Converted 5 records to sheet');
    expect(logged).toContain('Round-trip JSON: 5 rows');
    expect(logged.filter((l) => l.startsWith('  ')).map((l) => JSON.parse(l))).toEqual(employees);
    expect(download).toHaveBeenCalledOnce();
    expect(download.mock.calls[0]?.[1]).toBe('employees.xlsx');
  });
});

describe('examples/edge/service-worker/sw.js', () => {
  const XLSX_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

  /** Runs sw.js in a vm context with a stub ModernXlsx whose initWasm fails `failures` times. */
  function loadWorker(failures: number) {
    const listeners = new Map<string, (event: unknown) => void>();
    const initWasm = vi.fn((_url: string) =>
      initWasm.mock.calls.length <= failures
        ? Promise.reject(new Error('network error'))
        : Promise.resolve(),
    );
    const api = {
      initWasm,
      Workbook: class {
        addSheet() {
          return { cell: () => ({ value: undefined as unknown }) };
        }
        async toBuffer() {
          return new Uint8Array([0x50, 0x4b]);
        }
      },
    };
    const context = vm.createContext({ URL, Response, Date });
    Object.assign(context, {
      self: {
        addEventListener: (type: string, fn: (event: unknown) => void) => listeners.set(type, fn),
        clients: { claim: () => Promise.resolve() },
      },
      importScripts: () => Object.assign(context, { ModernXlsx: api }),
    });
    vm.runInContext(repoFile('examples/edge/service-worker/sw.js'), context);

    async function generate(): Promise<Response> {
      let response: Promise<Response> | undefined;
      listeners.get('fetch')?.({
        request: { url: 'https://site.example/api/xlsx/generate' },
        respondWith: (r: Promise<Response>) => {
          response = r;
        },
      });
      if (!response) throw new Error('the worker did not respond');
      return response;
    }
    return { generate, initWasm, listeners };
  }

  it('answers 500 when WASM init fails and retries the init on the next request', async () => {
    const { generate, initWasm } = loadWorker(1);

    const failed = await generate();
    expect(failed.status).toBe(500);
    expect(failed.headers.get('Content-Type')).toBe('text/plain');
    expect(await failed.text()).toBe('XLSX generation failed');

    const ok = await generate();
    expect(ok.status).toBe(200);
    expect(ok.headers.get('Content-Type')).toBe(XLSX_MIME);
    expect(initWasm).toHaveBeenCalledTimes(2);
    expect(initWasm.mock.calls[1]?.[0]).toMatch(/\/dist\/modern-xlsx\.wasm$/);
  });

  it('still pre-initializes WASM on install', async () => {
    const { initWasm, listeners } = loadWorker(0);
    const waited: Promise<unknown>[] = [];
    listeners.get('install')?.({ waitUntil: (p: Promise<unknown>) => waited.push(p) });
    expect(waited).toHaveLength(1);
    await expect(waited[0]).resolves.toBeUndefined();
    expect(initWasm).toHaveBeenCalledOnce();
  });
});
