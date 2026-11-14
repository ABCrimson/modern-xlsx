// node diff.mjs dir name...  — rasterise raw/<name> and <name> in Chromium at 2x; count pixels that differ by > 8/255
// (anti-aliasing from svgo's rounding) and by > 96/255 (a lost or moved shape would show here).
import { readFileSync } from 'node:fs';
import { chromium } from 'playwright';

const [dir, ...names] = process.argv.slice(2);
const browser = await chromium.launch();
const page = await browser.newPage();
for (const name of names) {
  const a = readFileSync(`${dir}/raw/${name}`, 'utf8');
  const b = readFileSync(`${dir}/${name}`, 'utf8');
  const res = await page.evaluate(
    async ([a, b]) => {
      const load = (s) =>
        new Promise((ok, bad) => {
          const i = new Image();
          i.onload = () => ok(i);
          i.onerror = bad;
          i.src = `data:image/svg+xml;charset=utf-8,${encodeURIComponent(s)}`;
        });
      const [ia, ib] = await Promise.all([load(a), load(b)]);
      const w = ia.naturalWidth * 2;
      const h = ia.naturalHeight * 2;
      const px = (img) => {
        const c = new OffscreenCanvas(w, h);
        const x = c.getContext('2d');
        x.drawImage(img, 0, 0, w, h);
        return x.getImageData(0, 0, w, h).data;
      };
      const da = px(ia);
      const db = px(ib);
      let diff = 0;
      let major = 0;
      for (let i = 0; i < da.length; i += 4)
        if (Math.max(...[0, 1, 2, 3].map((k) => Math.abs(da[i + k] - db[i + k]))) > 8) {
          diff++;
          if (Math.max(...[0, 1, 2].map((k) => Math.abs(da[i + k] - db[i + k]))) > 96) major++;
        }
      return { w, h, diff, major };
    },
    [a, b],
  );
  console.log(name, JSON.stringify(res));
}
await browser.close();
