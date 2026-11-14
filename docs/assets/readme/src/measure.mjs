// Raster checks in Chromium: wordmark ink gaps by column ink profile, and hero corner leak.
// Usage: node measure.mjs <hero-light.svg> <hero-dark.svg>
import { readFileSync } from 'node:fs';
import { chromium } from 'playwright';

const browser = await chromium.launch();
const page = await browser.newPage();
const res = {};
for (const f of process.argv.slice(2)) {
  const uri = `data:image/svg+xml;base64,${readFileSync(f).toString('base64')}`;
  res[f] = await page.evaluate(async (uri) => {
    const load = (w, h) =>
      new Promise((ok) => {
        const i = new Image(w, h);
        i.onload = () => ok(i);
        i.src = uri;
      });
    // 1) ink gaps at 1280: columns inside the range interior (y 120..240) that hold any pixel far from the range fill.
    let img = await load(1280, 320);
    let cv = new OffscreenCanvas(1280, 320);
    let cx = cv.getContext('2d');
    cx.drawImage(img, 0, 0, 1280, 320);
    const d = cx.getImageData(0, 0, 1280, 320).data;
    const at = (x, y) => [
      d[(y * 1280 + x) * 4],
      d[(y * 1280 + x) * 4 + 1],
      d[(y * 1280 + x) * 4 + 2],
    ];
    const bg = at(640, 116);
    const inkCol = [];
    for (let x = 212; x < 1068; x++) {
      let ink = false;
      for (let y = 120; y < 240; y++) {
        const p = at(x, y);
        if (Math.abs(p[0] - bg[0]) + Math.abs(p[1] - bg[1]) + Math.abs(p[2] - bg[2]) > 60) {
          ink = true;
          break;
        }
      }
      inkCol.push(ink);
    }
    const runs = [];
    let s = null;
    inkCol.forEach((v, i) => {
      if (v && s === null) s = i;
      if (!v && s !== null) {
        runs.push([s + 212, i + 211]);
        s = null;
      }
    });
    const gaps = runs.slice(1).map((r, i) => r[0] - runs[i][1] - 1);
    // 2) corner leak at 16x on #0D1117: any band-coloured pixel outside the border along the diagonal.
    img = await load(1280 * 16, 320 * 16);
    cv = new OffscreenCanvas(48 * 16, 48 * 16);
    cx = cv.getContext('2d');
    cx.fillStyle = '#0D1117';
    cx.fillRect(0, 0, cv.width, cv.height);
    cx.drawImage(img, 0, 0, 1280 * 16, 320 * 16);
    const e = cx.getImageData(0, 0, cv.width, cv.height).data;
    const diag = [];
    for (let i = 60; i < 100; i++) {
      const o = (i * cv.width + i) * 4;
      diag.push(
        `${i}:${[e[o], e[o + 1], e[o + 2]].map((v) => v.toString(16).padStart(2, '0')).join('')}`,
      );
    }
    return {
      inkRuns: runs.length,
      ink: [runs[0]?.[0], runs.at(-1)?.[1]],
      gaps,
      diag: diag.join(' '),
    };
  }, uri);
}
console.log(JSON.stringify(res, null, 1));
await browser.close();
