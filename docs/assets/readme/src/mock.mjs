// README mock: the first screen of a GitHub README with the assets in place, styled by github-markdown-css.
// Usage: node mock.mjs <assetDir> <outDir>   -> writes readme-{light,dark}.html (render at 1280 and 390).
import { writeFileSync } from 'node:fs';
import { findPackageJSON } from 'node:module';
import { dirname, join } from 'node:path';
import { pathToFileURL } from 'node:url';

const [A, out] = process.argv.slice(2);
const css = dirname(findPackageJSON('github-markdown-css', import.meta.url));
const url = (f) => pathToFileURL(join(A, f)).href;
// Five flat-square badges as plain two-tone rectangles (label #24292f plus value colour, in the README's scheme), as shields.io draws them.
const badges = [
  [38, 52, '#c8102e'],
  [44, 64, '#4c1'],
  [52, 40, '#59636e'],
  [38, 70, '#59636e'],
  [56, 46, '#59636e'],
];
for (const theme of ['light', 'dark']) {
  const page = theme === 'light' ? '#ffffff' : '#0d1117';
  const line = theme === 'light' ? '#d1d9e0' : '#3d444d';
  const badgeRow = badges
    .map(
      ([l, r, c]) =>
        `<span style="display:inline-flex;height:20px;margin-right:4px;border-radius:2px;overflow:hidden;vertical-align:middle"><span style="width:${l}px;background:#24292f"></span><span style="width:${r}px;background:${c}"></span></span>`,
    )
    .join('');
  writeFileSync(
    join(out, `readme-${theme}.html`),
    `<!doctype html><html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width">
<link rel="stylesheet" href="${pathToFileURL(join(css, `github-markdown-${theme}.css`)).href}">
<style>
body{margin:0;background:${page}}
.wrap{max-width:904px;margin:24px auto;border:1px solid ${line};border-radius:6px}
.markdown-body{padding:32px;box-sizing:border-box}
@media (max-width:600px){.wrap{margin:0;border:0;border-radius:0}.markdown-body{padding:16px}}
</style></head><body><div class="wrap" id="c"><article class="markdown-body">
<p><img src="${url(`hero-${theme}.svg`)}" alt="modern-xlsx" width="100%"></p>
<p><strong>Read and write Excel .xlsx files from JavaScript and TypeScript, on a Rust core compiled to WebAssembly.</strong></p>
<p>${badgeRow}</p>
<p><a href="#">Documentation</a> · <a href="#">API reference</a> · <a href="#">Examples</a> · <a href="#">Changelog</a></p>
<h2>How it works</h2>
<p><img src="${url(`architecture-${theme}.svg`)}" alt="architecture" width="640"></p>
</article></div></body></html>`,
  );
}
