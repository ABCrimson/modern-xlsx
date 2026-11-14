// node sheet.mjs out.html maxWidth "Heading" "label|file|width|bg|pixelated?" ...  (groups split by a bare "--")
import { readFileSync, writeFileSync } from 'node:fs';

const [out, maxW, heading, ...items] = process.argv.slice(2);
const uri = (f) =>
  `data:${f.endsWith('.png') ? 'image/png' : 'image/svg+xml'};base64,${readFileSync(f).toString('base64')}`;
const rows = [[]];
for (const it of items) {
  if (it === '--') {
    rows.push([]);
    continue;
  }
  rows.at(-1).push(it.split('|'));
}
const cell = ([label, f, w, bg, px]) =>
  `<figure style="margin:0"><div style="background:${bg};padding:16px;display:inline-block;border:1px solid #d0d7de"><img src="${uri(f)}" style="display:block;width:${w}px;height:auto${px ? ';image-rendering:pixelated' : ''}"></div><figcaption style="font:500 13px/1.4 system-ui,sans-serif;color:#1f2328;margin-top:6px">${label}</figcaption></figure>`;
const html = `<!doctype html><html><body style="margin:0;background:#ffffff"><div id="c" style="width:${maxW}px;box-sizing:border-box;padding:24px;background:#ffffff">
<h1 style="font:600 20px/1.3 system-ui,sans-serif;color:#1f2328;margin:0 0 16px">${heading}</h1>
${rows.map((r) => `<div style="display:flex;flex-wrap:wrap;gap:20px;align-items:flex-end;margin-bottom:22px">${r.map(cell).join('')}</div>`).join('\n')}
</div></body></html>`;
writeFileSync(out, html);
