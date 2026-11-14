// node render.mjs jobs.json — jobs: [{svg, width, bg, out, pad?, scale?, browser?} | {html, out, width, scale?, full?}]
// Renders each SVG through <img> (as GitHub and npm show it) on a page colour, or an HTML page (#c or full page).
import { readFileSync } from 'node:fs';
import { pathToFileURL } from 'node:url';
import { chromium, firefox } from 'playwright';

const engines = { chromium, firefox };

const jobs = JSON.parse(readFileSync(process.argv[2], 'utf8'));
const browsers = new Map();
const pages = new Map();
for (const j of jobs) {
  const kind = j.browser ?? 'chromium';
  if (!browsers.has(kind)) {
    const b = await engines[kind].launch(
      kind === 'chromium' ? { args: ['--use-angle=d3d11'] } : {},
    );
    browsers.set(kind, b);
    console.log(kind, b.version());
  }
  const scale = j.scale ?? 1;
  const vw = j.viewport ?? 1800;
  const key = `${kind}|${scale}|${vw}`;
  let page = pages.get(key);
  if (!page) {
    page = await browsers
      .get(kind)
      .newPage({ deviceScaleFactor: scale, viewport: { width: vw, height: 1000 } });
    pages.set(key, page);
  }
  if (j.html) {
    await page.goto(pathToFileURL(j.html).href, { waitUntil: 'load' });
    if (j.full) await page.screenshot({ path: j.out, fullPage: true, clip: j.clip });
    else await page.locator('#c').screenshot({ path: j.out });
  } else {
    const src = `data:image/svg+xml;base64,${readFileSync(j.svg).toString('base64')}`;
    const pad = j.pad ?? 24;
    const px = j.pixelated ? ';image-rendering:pixelated' : '';
    const inner = j.zoom
      ? `<img src="${src}" style="display:block;width:${j.width}px;height:auto" id="i">`
      : `<img src="${src}" style="display:block;width:${j.width}px;height:auto${px}">`;
    await page.setContent(
      `<!doctype html><html><body style="margin:0;background:${j.bg}"><div id="c" style="display:inline-block;padding:${pad}px;background:${j.bg}">${inner}</div></body></html>`,
      { waitUntil: 'load' },
    );
    await page.locator('#c').screenshot({ path: j.out });
  }
  console.log('wrote', j.out);
}
for (const b of browsers.values()) await b.close();
