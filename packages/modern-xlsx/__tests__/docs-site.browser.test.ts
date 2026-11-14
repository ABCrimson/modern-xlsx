import { afterEach, describe, expect, it } from 'vitest';
import { commands } from 'vitest/browser';

/**
 * How a rendered playground gets the library: the package's own IIFE bundle (dist/, built
 * before the browser tests run, as for iife.browser.test.ts), whose initWasm waits for `hold`
 * when one is given. The bundle runs inside the page's frame, so its values share the page's realm.
 */
interface Library {
  hold?: Promise<void>;
}

/**
 * The static docs site (docs/site/*.html) rendered in Chromium. The pages load their fonts from
 * Google Fonts, and the playground loads vendor/modern-xlsx.min.js, which pages.yml copies in at
 * deploy time; the tests drop those tags so they run offline. Without the IIFE the playground's
 * WASM init fails inside its try/catch and the page still wires its editor. Given a `library`,
 * the playground loads /dist/modern-xlsx.min.js instead (a srcdoc frame resolves URLs against
 * this window's), and a `hold` reaches its initWasm through this window (it is same-origin).
 */
async function page(name: string, library?: Library): Promise<string> {
  const html = await commands.readFile(`../../docs/site/${name}`);
  Object.assign(globalThis, { playgroundHold: library?.hold });
  const hold =
    '<script>{ const init = ModernXlsx.initWasm; window.ModernXlsx = { ...ModernXlsx, initWasm: () => parent.playgroundHold.then(() => init()) }; }</script>';
  const iife = `<script src="/dist/modern-xlsx.min.js"></script>${library?.hold ? hold : ''}`;
  return html
    .replace(/<link[^>]+fonts\.(googleapis|gstatic)\.com[^>]*>/g, '')
    .replace(/<script src="vendor\/modern-xlsx\.min\.js"><\/script>/g, library ? iife : '');
}

const frames: HTMLIFrameElement[] = [];

afterEach(() => {
  for (const frame of frames.splice(0)) frame.remove();
});

async function render(name: string, width: number, library?: Library): Promise<Document> {
  const frame = document.createElement('iframe');
  frame.style.width = `${width}px`;
  frame.style.height = '800px';
  frame.style.border = '0';
  frames.push(frame);
  const loaded = new Promise<void>((resolve) =>
    frame.addEventListener('load', () => resolve(), { once: true }),
  );
  frame.srcdoc = await page(name, library);
  document.body.append(frame);
  await loaded;
  const doc = frame.contentDocument;
  if (!doc) throw new Error(`${name} did not render`);
  return doc;
}

function element<T extends Element>(doc: Document, selector: string): T {
  const found = doc.querySelector<T>(selector);
  if (!found) throw new Error(`${selector} not found`);
  return found;
}

describe('docs/site/playground.html', () => {
  /** Dispatches a keydown on the editor and reports whether the page cancelled it. */
  function press(editor: HTMLTextAreaElement, key: string, shiftKey = false): boolean {
    const event = new KeyboardEvent('keydown', { key, shiftKey, bubbles: true, cancelable: true });
    editor.dispatchEvent(event);
    return event.defaultPrevented;
  }

  it('names the editor and shows its keyboard focus', async () => {
    const doc = await render('playground.html', 1280);
    const editor = element<HTMLTextAreaElement>(doc, '#editor');
    expect(editor.getAttribute('aria-label')).toBe('Code');
    editor.focus();
    expect(editor.matches(':focus-visible')).toBe(true);
    const style = doc.defaultView?.getComputedStyle(editor);
    expect(style?.outlineStyle).toBe('solid');
    expect(style?.outlineWidth).toBe('2px');
    // --accent of the light theme, light-dark(#c8102e, #ff6b7a): the test browser prefers light.
    expect(style?.outlineColor).toBe('rgb(200, 16, 46)');
    expect(style?.outlineOffset).toBe('-2px');
  });

  it('inserts two spaces on Tab, and lets Shift+Tab or Esc then Tab leave the editor', async () => {
    const doc = await render('playground.html', 1280);
    const editor = element<HTMLTextAreaElement>(doc, '#editor');
    editor.value = 'ab';
    editor.setSelectionRange(1, 1);

    expect(press(editor, 'Tab')).toBe(true);
    expect(editor.value).toBe('a  b');
    expect(editor.selectionStart).toBe(3);

    expect(press(editor, 'Tab', true)).toBe(false);
    expect(editor.value).toBe('a  b');

    expect(press(editor, 'Escape')).toBe(false);
    expect(press(editor, 'Tab')).toBe(false);
    expect(editor.value).toBe('a  b');

    // The escape lasts for one Tab only, and any other key cancels it.
    expect(press(editor, 'Tab')).toBe(true);
    expect(press(editor, 'Escape')).toBe(false);
    press(editor, 'x');
    expect(press(editor, 'Tab')).toBe(true);
  });

  it('tells keyboard users how to leave the editor', async () => {
    const doc = await render('playground.html', 1280);
    expect(element(doc, '.toolbar .hint').textContent).toContain('Esc then Tab leaves the editor');
  });

  it('announces output and WASM status, and marks up its landmark and icons', async () => {
    const doc = await render('playground.html', 1280);
    const output = element(doc, '#output');
    expect(output.getAttribute('role')).toBe('log');
    expect(output.getAttribute('aria-live')).toBe('polite');
    expect(output.getAttribute('aria-label')).toBe('Output');
    expect(element(doc, '#wasm-status').getAttribute('role')).toBe('status');
    expect(doc.querySelectorAll('main')).toHaveLength(1);
    expect(element(doc, 'main').classList.contains('playground')).toBe(true);
    const icons = [...doc.querySelectorAll('.panel-header svg.i')];
    expect(icons).toHaveLength(2);
    for (const icon of icons) expect(icon.getAttribute('aria-hidden')).toBe('true');
  });

  it('wraps a long unbroken log line instead of widening the page on a phone', async () => {
    const doc = await render('playground.html', 390);
    const log = (
      doc.defaultView as (Window & { log?: (msg: string, level: string) => void }) | null
    )?.log;
    expect(log).toBeTypeOf('function');
    // A stack frame as runCode logs it: one long URL with no break opportunity.
    log?.(`at runCode (http://127.0.0.1:30177/${'playground'.repeat(12)}.html:891:13)`, 'error');
    expect(doc.documentElement.scrollWidth).toBeLessThanOrEqual(doc.documentElement.clientWidth);
  });

  it('marks a failed load with an icon and words, not with a colour that matches the accent', async () => {
    // Without the IIFE (dropped by page()), initWasm throws and the page shows its error state.
    const doc = await render('playground.html', 1280);
    await expect.poll(() => doc.querySelector('#output .log-line.error')).not.toBeNull();
    const status = element(doc, '#wasm-status');
    expect(status.classList.contains('error')).toBe(true);
    expect(status.textContent?.trim()).toBe('Failed to load');
    expect(status.querySelector('svg use')?.getAttribute('href')).toBe('#i-circle-alert');
    const line = element(doc, '#output .log-line.error');
    expect(line.textContent).toContain('Failed to load modern-xlsx');
    const icon = line.querySelector('svg');
    expect(icon?.getAttribute('aria-hidden')).toBe('true');
    expect(icon?.querySelector('use')?.getAttribute('href')).toBe('#i-circle-alert');
    expect(element(doc, '#output .log-line.warn svg use').getAttribute('href')).toBe(
      '#i-triangle-alert',
    );
    const view = doc.defaultView;
    const accent = view?.getComputedStyle(element(doc, '#run-btn')).backgroundColor;
    expect(accent).toBe('rgb(200, 16, 46)');
    expect(view?.getComputedStyle(line).color).not.toBe(accent);
  });

  /** The text of each preview table row, cells joined by ' | ', of the last preview shown. */
  function previewRows(doc: Document): string[] {
    const tables = doc.querySelectorAll('#output .preview table');
    const table = tables[tables.length - 1];
    return [...(table?.querySelectorAll('tr') ?? [])].map((tr) =>
      [...tr.cells].map((c) => c.textContent).join(' | '),
    );
  }

  /** Runs the editor's code with the Run button and waits for its last log line. */
  async function run(doc: Document, code?: string): Promise<string[]> {
    const editor = element<HTMLTextAreaElement>(doc, '#editor');
    if (code !== undefined) editor.value = code;
    const output = element(doc, '#output');
    output.replaceChildren();
    element<HTMLButtonElement>(doc, '#run-btn').click();
    await expect
      .poll(() => output.lastElementChild?.textContent ?? '')
      .toMatch(/Done in|Error:|of your code/);
    return [...output.querySelectorAll('.log-line')].map((l) => l.textContent ?? '');
  }

  async function ready(width = 1280): Promise<Document> {
    const doc = await render('playground.html', width, {});
    await expect.poll(() => element(doc, '#wasm-status').className).toBe('status ready');
    return doc;
  }

  it('shows the Hello example while the library loads, and keeps what the visitor types then', async () => {
    let finish = (): void => {};
    const hold = new Promise<void>((resolve) => {
      finish = resolve;
    });
    const doc = await render('playground.html', 1280, { hold });
    const editor = element<HTMLTextAreaElement>(doc, '#editor');
    expect(element(doc, '#wasm-status').className).toBe('status loading');
    expect(editor.value).toMatch(/^\/\/ Create a simple workbook and download it\n/);
    // The menu works during the load too.
    const menu = element<HTMLSelectElement>(doc, '#examples');
    menu.value = 'styled';
    menu.dispatchEvent(new Event('change'));
    expect(editor.value).toMatch(/^\/\/ Styled workbook/);
    editor.value = "log('my own code');";
    finish();
    await expect.poll(() => element(doc, '#wasm-status').className).toBe('status ready');
    await new Promise((resolve) => setTimeout(resolve, 50));
    expect(editor.value).toBe("log('my own code');");
    const lines = [...doc.querySelectorAll('#output .log-line')].map((l) => l.textContent ?? '');
    expect(lines.some((l) => l.includes('Executing'))).toBe(false);
  });

  it('runs the untouched Hello example once the library is ready', async () => {
    const doc = await ready();
    await expect.poll(() => previewRows(doc)).toEqual(['Hello | World', '42 | TRUE']);
    expect(element(doc, '#output .preview thead').querySelectorAll('th')).toHaveLength(2);
  });

  it('previews numbers through their number format and formulas without a cached value as text', async () => {
    const doc = await ready();
    const menu = element<HTMLSelectElement>(doc, '#examples');

    menu.value = 'styled';
    menu.dispatchEvent(new Event('change'));
    await run(doc);
    expect(previewRows(doc)[1]).toBe(
      'Widget A | $12,500.00 | $15,800.00 | $18,200.00 | $22,100.00 | 76.8%',
    );

    menu.value = 'formulas';
    menu.dispatchEvent(new Event('change'));
    expect((await run(doc)).at(-1)).toMatch(/Done in/);
    const rows = previewRows(doc);
    expect(rows[0]).toMatch(/^Invoice Date: \| \d{4}-\d{2}-\d{2} \| {2}\| $/);
    expect(rows[3]).toBe('Laptop | 2 | $1,299.99 | =B4*C4');
    expect(rows[7]).toBe(' |  | Grand Total: | =SUM(D4:D7)');
  });

  it("previews a 1904 workbook's elapsed times and dates in its own date system", async () => {
    const doc = await ready();
    await run(
      doc,
      [
        "const wb = new ModernXlsx.Workbook({ dateSystem: 'date1904' });",
        "const ws = wb.addSheet('S');",
        "const elapsed = new ModernXlsx.StyleBuilder().numberFormat('[h]:mm').build(wb.styles);",
        "ws.cell('A1').value = 1.5;",
        "ws.cell('A1').styleIndex = elapsed;",
        // formatCell reads serial 60 as 1-Mar-1904 in the 1904 system and as 28-Feb-1900 in
        // the 1900 system, so this cell shows which system the preview used.
        "ws.cell('A2').value = 60;",
        "ws.cell('A2').styleIndex = new ModernXlsx.StyleBuilder().numberFormat('d-mmm-yyyy').build(wb.styles);",
        'preview(ws);',
      ].join('\n'),
    );
    expect(previewRows(doc)).toEqual(['36:00', '1-Mar-1904']);
  });

  /** The tag of each cell in the last preview's header row. */
  function headerTags(doc: Document): string[] {
    const heads = doc.querySelectorAll('#output .preview thead tr');
    return [...(heads[heads.length - 1]?.children ?? [])].map((c) => c.tagName);
  }

  it('puts an empty cell of the header row in a td, so no header cell is empty', async () => {
    const doc = await ready();
    const menu = element<HTMLSelectElement>(doc, '#examples');
    menu.value = 'formulas';
    menu.dispatchEvent(new Event('change'));
    await run(doc);
    // Row 1 of the example: 'Invoice Date:', the date, then two empty cells.
    expect(headerTags(doc)).toEqual(['TH', 'TH', 'TD', 'TD']);
    expect(doc.querySelectorAll('#output .preview th:empty')).toHaveLength(0);
  });

  it('previews the first 50 columns of a sheet too wide to build, says so, and stays quick', async () => {
    const doc = await ready();
    const lines = await run(
      doc,
      [
        "const ws = new ModernXlsx.Workbook().addSheet('Wide');",
        "ws.cell('A1').value = 'first';",
        "ws.cell('XFD300').value = 'last';",
        'const t0 = performance.now();',
        'preview(ws);',
        "log('preview took ' + (performance.now() - t0));",
      ].join('\n'),
    );
    expect(lines).toContainEqual(
      expect.stringMatching(/Preview shows the first 50 of 16384 columns\.$/),
    );
    const took = Number(/preview took ([\d.]+)/.exec(lines.join('\n'))?.[1]);
    expect(took).toBeLessThan(1000);
    const rows = [...doc.querySelectorAll('#output .preview tr')];
    expect(rows).toHaveLength(300);
    for (const row of [rows[0], rows[299]]) expect(row?.children).toHaveLength(50);
    expect(rows[0]?.textContent).toBe('first');
  });

  /** Whether an element lies wholly inside its frame's viewport, to the nearest pixel. */
  function inView(el: Element | null | undefined): boolean {
    const view = el?.ownerDocument.defaultView;
    if (!el || !view) return false;
    const { top, bottom } = el.getBoundingClientRect();
    return Math.round(top) >= 0 && Math.round(bottom) <= view.innerHeight;
  }

  it('brings the newest result into view on a phone, after a run and after an opened file', async () => {
    const doc = await ready(390);
    const output = element(doc, '#output');
    const menu = element<HTMLSelectElement>(doc, '#examples');
    menu.value = 'formulas';
    menu.dispatchEvent(new Event('change'));
    doc.defaultView?.scrollTo(0, 0);
    await run(doc);
    await expect.poll(() => inView(output.lastElementChild)).toBe(true);

    // A file opened through the file input: the page reads it, previews it and logs it.
    doc.defaultView?.scrollTo(0, 0);
    await run(
      doc,
      [
        'const wb = new ModernXlsx.Workbook();',
        "const ws = wb.addSheet('Gaps');",
        "ws.cell('A1').value = 'Name'; ws.cell('C1').value = 'Qty';",
        "ws.cell('A2').value = 'Pen'; ws.cell('B2').value = 'blue'; ws.cell('C2').value = 3;",
        'const files = new DataTransfer();',
        "files.items.add(new File([await wb.toBuffer()], 'gaps.xlsx'));",
        "const input = document.getElementById('open-file');",
        'input.files = files.files;',
        "input.dispatchEvent(new Event('change'));",
      ].join('\n'),
    );
    await expect.poll(() => previewRows(doc)).toEqual(['Name |  | Qty', 'Pen | blue | 3']);
    expect(headerTags(doc)).toEqual(['TH', 'TD', 'TH']);
    await expect.poll(() => inView(output.lastElementChild)).toBe(true);
  });

  it('never moves the page on a desktop layout', async () => {
    const doc = await ready();
    const menu = element<HTMLSelectElement>(doc, '#examples');
    menu.value = 'formulas';
    menu.dispatchEvent(new Event('change'));
    await run(doc);
    await new Promise((resolve) => setTimeout(resolve, 100));
    expect(doc.scrollingElement?.scrollTop).toBe(0);
  });

  it('puts cell text into the preview as text, never as markup', async () => {
    const doc = await ready();
    const markup = '<img src="x" onerror="parent.injected = true">';
    await run(
      doc,
      `const ws = new ModernXlsx.Workbook().addSheet('S');\nws.cell('A1').value = 'Head';\nws.cell('A2').value = ${JSON.stringify(markup)};\npreview(ws);`,
    );
    expect(doc.querySelector('#output .preview img')).toBeNull();
    expect(previewRows(doc)).toEqual(['Head', markup]);
  });

  it("names the line of the visitor's own code when it throws, without the page's frames", async () => {
    const doc = await ready();
    const lines = await run(doc, "log('one');\nconst sheet = null;\nsheet.cell('A1');");
    expect(lines.at(-2)).toMatch(/^\d\d:\d\d:\d\d\.\d{3}Error: Cannot read properties of null/);
    expect(lines.at(-1)).toMatch(/at line 3, column \d+ of your code$/);
    expect(lines.join('\n')).not.toMatch(/runCode|<anonymous>|playground/);
  });

  it('gives the download link a 44px tap target on a phone', async () => {
    const doc = await ready(390);
    await expect.poll(() => doc.querySelector('#output .download-link')).not.toBeNull();
    const link = element(doc, '#output .download-link');
    expect(link.getBoundingClientRect().height).toBeGreaterThanOrEqual(44);
  });

  it('puts its nav in a banner landmark and closes with a footer', async () => {
    const doc = await render('playground.html', 1280);
    expect(doc.querySelectorAll('header')).toHaveLength(1);
    expect(doc.querySelectorAll('body > header > nav[aria-label="Primary"]')).toHaveLength(1);
    expect(doc.querySelectorAll('footer')).toHaveLength(1);
    expect(element(doc, 'main').nextElementSibling?.tagName).toBe('FOOTER');
  });
});

describe('docs/site/index.html', () => {
  it('fits a 375px phone without horizontal scroll', async () => {
    const doc = await render('index.html', 375);
    expect(doc.documentElement.scrollWidth).toBeLessThanOrEqual(375);
  });

  it('puts its nav in a banner landmark', async () => {
    const doc = await render('index.html', 1280);
    expect(doc.querySelectorAll('header')).toHaveLength(1);
    expect(doc.querySelectorAll('body > header > nav[aria-label="Primary"]')).toHaveLength(1);
  });

  it('gives the Copy buttons and the version badge 44px tap targets on a phone', async () => {
    const doc = await render('index.html', 390);
    const targets = [...doc.querySelectorAll('.copy, .hero .badge')];
    expect(targets).toHaveLength(3);
    for (const target of targets) {
      expect(target.getBoundingClientRect().height).toBeGreaterThanOrEqual(44);
    }
  });

  it('keeps the four stat cards one height on a phone', async () => {
    const doc = await render('index.html', 390);
    const heights = [...doc.querySelectorAll('.stat')].map((s) => s.getBoundingClientRect().height);
    expect(heights).toHaveLength(4);
    for (const h of heights) expect(h).toBe(heights[0]);
  });

  it('scrolls every in-page anchor target to below the sticky header', async () => {
    const doc = await render('index.html', 1280);
    const header = element(doc, 'body > header');
    const ids = [...doc.querySelectorAll('main [id]')].map((e) => e.id);
    expect(ids).toEqual(expect.arrayContaining(['features', 'architecture', 'api-title']));
    for (const id of ids) {
      const target = element(doc, `#${id}`);
      target.scrollIntoView({ behavior: 'instant' });
      const heading = target.matches('h2') ? target : (target.querySelector('h2') ?? target);
      const below = heading.getBoundingClientRect().top - header.getBoundingClientRect().bottom;
      expect(below, `#${id}`).toBeGreaterThanOrEqual(16);
    }
  });

  it('keeps the desktop nav on one 56px row', async () => {
    const doc = await render('index.html', 1280);
    expect(element(doc, 'nav .inner').getBoundingClientRect().height).toBe(56);
  });

  it('wraps the content in one main landmark and hides the decorative icons', async () => {
    const doc = await render('index.html', 1280);
    const mains = doc.querySelectorAll('main');
    expect(mains).toHaveLength(1);
    const main = element(doc, 'main');
    expect(main.id).toBe('main');
    expect(main.firstElementChild?.classList.contains('hero')).toBe(true);
    expect(main.querySelectorAll('section').length).toBe(doc.querySelectorAll('section').length);
    expect(main.nextElementSibling?.tagName).toBe('FOOTER');
    const icons = [...doc.querySelectorAll('.features .icon')];
    expect(icons).toHaveLength(18);
    for (const icon of icons) {
      expect(icon.getAttribute('aria-hidden')).toBe('true');
      expect(icon.querySelector('svg')?.getAttribute('aria-hidden')).toBe('true');
    }
  });

  it('sets the headline in ink with the crimson accent on its second line, and the stats in ink', async () => {
    const doc = await render('index.html', 1280);
    const view = doc.defaultView;
    const h1 = element(doc, '.hero h1');
    const ink = view?.getComputedStyle(h1).color;
    expect(ink).toBe(view?.getComputedStyle(doc.body).color);
    const accent = element(doc, '.hero h1 .accent');
    expect(accent.textContent).toBe('handled right.');
    const style = view?.getComputedStyle(accent);
    expect(style?.color).toBe('rgb(200, 16, 46)');
    expect(style?.backgroundImage).toBe('none');
    expect(style?.webkitTextFillColor).toBe('rgb(200, 16, 46)');
    const numbers = [...doc.querySelectorAll('.stat .number')];
    expect(numbers).toHaveLength(4);
    for (const n of numbers) expect(view?.getComputedStyle(n).color).toBe(ink);
  });

  it('gives every scrolling region a keyboard stop and a name', async () => {
    const doc = await render('index.html', 1280);
    const regions = [...doc.querySelectorAll('.table-wrap, .code-scroll')];
    expect(regions).toHaveLength(3);
    for (const region of regions) {
      expect(region.tagName).toBe('SECTION');
      expect(region.getAttribute('tabindex')).toBe('0');
      const label =
        region.getAttribute('aria-label') ??
        doc.getElementById(region.getAttribute('aria-labelledby') ?? '')?.textContent;
      expect(label).toBeTruthy();
    }
  });

  it('shows the architecture picture 640px wide on a desktop, and with 11px lettering on a phone', async () => {
    const desktop = await render('index.html', 1280);
    const img = element<HTMLImageElement>(desktop, '#architecture picture img');
    expect(img.getBoundingClientRect().width).toBe(640);
    expect(img.getBoundingClientRect().height).toBeCloseTo((640 * 1040) / 960, 0);
    // A 390px phone viewport: the test browser's iframe draws a classic scrollbar, which phones
    // overlay, so the frame is widened by the scrollbar's width.
    const probe = await render('index.html', 390);
    const bar = 390 - probe.documentElement.clientWidth;
    const phone = bar > 0 ? await render('index.html', 390 + bar) : probe;
    expect(phone.documentElement.clientWidth).toBe(390);
    const small = element<HTMLImageElement>(phone, '#architecture picture img');
    // The smallest label is 30 units of the 960-unit viewBox.
    expect((small.getBoundingClientRect().width * 30) / 960).toBeGreaterThanOrEqual(11);
    expect(phone.documentElement.scrollWidth).toBeLessThanOrEqual(390);
  });
});
