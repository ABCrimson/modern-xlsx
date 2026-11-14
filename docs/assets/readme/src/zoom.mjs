// node zoom.mjs <svg> <displayWidth> <bg> <out.png> <x> <y> <w> <h> [scale]
// Renders the SVG through <img> at displayWidth CSS px and clips a region given in viewBox units, at `scale` device px per CSS px.
import { readFileSync } from 'node:fs';
import { chromium } from 'playwright';

const [svg, dw, bg, out, x, y, w, h, sc = '4'] = process.argv.slice(2);
const src = readFileSync(svg, 'utf8');
const vbW = +src.match(/viewBox="0 0 (\d+)/)[1];
const k = +dw / vbW;
const b = await chromium.launch({ args: ['--use-angle=d3d11'] });
const p = await b.newPage({ deviceScaleFactor: +sc, viewport: { width: +dw + 40, height: 2000 } });
await p.setContent(
  `<body style="margin:0;background:${bg}"><img style="display:block;width:${dw}px" src="data:image/svg+xml;base64,${Buffer.from(src).toString('base64')}"></body>`,
  { waitUntil: 'load' },
);
await p.screenshot({ path: out, clip: { x: +x * k, y: +y * k, width: +w * k, height: +h * k } });
await b.close();
console.log('wrote', out);
