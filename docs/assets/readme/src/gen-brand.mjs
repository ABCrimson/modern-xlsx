// Brand family: mark-light.svg, mark-dark.svg (A1), favicon.svg (16 px form), hero-light.svg, hero-dark.svg (A2).
// Usage: node gen-brand.mjs <outDir>   -> prints a JSON report (rule checks, sizes, wordmark gaps).
import { join } from 'node:path';
import { emit, n2, PAL, paths, run } from './lib.mjs';

const outDir = process.argv[2] ?? join(import.meta.dirname, 'out');

// ---------- A1 mark, 64x64 ----------
// Every edge lands on an even unit, so the mark is pixel-crisp at 64 px and at 32 px.
// Frame 6..58 (stroke 4: outer 4, inner 8, inner radius 8); grid 2 wide at 19 and 45;
// active cell = the first data cell (A1), stroke 4 on 20..44; fill handle red 38..50 with a 2-unit ring.
export const MARK = {
  page: (c) => `<rect x="6" y="6" width="52" height="52" rx="10" fill="${c.page}"/>`,
  // Header bands as one L that follows the frame's inner corner arcs (no clipPath).
  bands: (c) =>
    `<path fill="${c.band}" d="M8 16a8 8 0 0 1 8-8h32a8 8 0 0 1 8 8v2H18v38h-2a8 8 0 0 1-8-8z"/>`,
  // x=45 and y=45 stop under the handle's knockout ring, so no stubs show beside it.
  grid: (c) =>
    `<path stroke="${c.grid}" stroke-width="2" d="M19 7.5v49M45 7.5V36M7.5 19h49M7.5 45H36"/>`,
  frame: (_c, frame) =>
    `<rect x="6" y="6" width="52" height="52" rx="10" fill="none" stroke="${frame}" stroke-width="4"/>`,
  cell: (c) => `<path fill="none" stroke="${c.accent}" stroke-width="4" d="M20 20h24v24H20z"/>`,
  handle: (c) =>
    `<path fill="${c.accent}" stroke="${c.page}" stroke-width="2" d="M37 37h14v14H37z"/>`,
};
// Dark frame is muted (#9198A1), not ink: the frame must weigh less than the accent cell.
export const markFrame = (theme) => (theme === 'dark' ? PAL.dark.muted : PAL.light.ink);
export function markBody(theme) {
  const c = PAL[theme];
  return [
    MARK.page(c),
    MARK.bands(c),
    MARK.grid(c),
    MARK.frame(c, markFrame(theme)),
    MARK.cell(c),
    MARK.handle(c),
  ].join('');
}
// The mark placed at (x, y) scaled by s, for the cards.
export const markAt = (theme, x, y, s) =>
  `<g transform="translate(${x} ${y}) scale(${s})">${markBody(theme)}</g>`;

// ---------- favicon, 16x16 ----------
// One file for light and dark tabs: an opaque white tile with a 1 px ink border, a 2 px crimson
// cell outline (outer 3..12) and a 4 px handle with a 1 px white knockout. No bands, no grid.
function favicon() {
  const c = PAL.light;
  return [
    `<rect x=".5" y=".5" width="15" height="15" rx="3.5" fill="${c.page}" stroke="${c.ink}"/>`,
    `<path fill="none" stroke="${c.accent}" stroke-width="2" d="M4 4h7v7H4z"/>`,
    `<path fill="${c.accent}" stroke="${c.page}" d="M8.5 8.5h5v5h-5z"/>`,
  ].join('');
}

// ---------- A2 hero, 1280x320 ----------
// The grid is the field; the wordmark is the content of the active merged range B2:G3.
export const WORDMARK = { size: 112, tracking: -0.01, nudges: { 5: -6, 6: -5 } }; // n|- and -|x pair kerning
export function wordmark(c, size = WORDMARK.size, nudges = WORDMARK.nudges) {
  return run(
    [
      { text: 'modern', font: 'mono', colour: c.ink },
      { text: '-', font: 'mono', colour: c.accent },
      { text: 'xlsx', font: 'mono', colour: c.ink },
    ],
    size,
    {
      tracking: WORDMARK.tracking,
      nudges: Object.fromEntries(
        Object.entries(nudges).map(([k, v]) => [k, (v * size) / WORDMARK.size]),
      ),
    },
  );
}

function hero(theme) {
  const c = PAL[theme];
  const W = 1280;
  const H = 320;
  const RH = 56;
  const CH = 40;
  const COL = 146;
  const ROW = 70;
  const xs = [];
  for (let x = RH; x < W; x += COL) xs.push(x);
  const ys = [];
  for (let y = CH; y < H; y += ROW) ys.push(y);
  const r = { x1: xs[1], x2: xs[7], y1: ys[1], y2: ys[3] }; // B2:G3 = x 202..1078, y 110..250, centred on 640
  const grid = [
    ...xs.map((x) => `M${x + 0.5} 0V${H}`),
    ...ys.map((y) => `M0 ${y + 0.5}H${W}`),
  ].join('');
  const wm = wordmark(c);
  const baseline = (r.y1 + r.y2) / 2 + wm.box.top / 2 - 2; // ascender box centred, raised 2 for the x-height mass
  const placed = wm.place(640, baseline, 'ink-centre');
  const B = 8;
  const HS = 28;
  const RING = 4; // the mark x2: border 8, handle red 24 + ring 4
  return {
    meta: {
      baseline: n2(baseline),
      ink: [n2(placed.x1), n2(placed.x2)],
      centre: n2((placed.x1 + placed.x2) / 2),
      gaps: wm.gaps,
      range: r,
    },
    body: [
      // Clip to the panel's inner edge so nothing leaks past the rounded border.
      `<clipPath id="p"><rect x="1" y="1" width="${W - 2}" height="${H - 2}" rx="15.5"/></clipPath>`,
      `<rect x=".5" y=".5" width="${W - 1}" height="${H - 1}" rx="16" fill="${c.panel}"/>`,
      `<g clip-path="url(#p)">`,
      `<path fill="${c.band}" d="M0 0H${W}V${CH}H${RH}V${H}H0z"/>`,
      // Selected range marked in the headers, as real sheets do (panel-border colour, not accent).
      `<path fill="${c.border}" d="M${r.x1} 0H${r.x2}V${CH}H${r.x1}zM0 ${r.y1}H${RH}V${r.y2}H0z"/>`,
      // No shape-rendering="crispEdges": Firefox 159 drops crispEdges lines thinner than a device pixel
      // (5 of 9 columns vanished at 358 px, 3 at 838 px); anti-aliased 1-unit lines show in every browser.
      `<path stroke="${c.grid}" d="${grid}"/>`,
      `<rect x="${r.x1}" y="${r.y1}" width="${r.x2 - r.x1}" height="${r.y2 - r.y1}" fill="${c.page}"/>`,
      `</g>`,
      `<rect x=".5" y=".5" width="${W - 1}" height="${H - 1}" rx="16" fill="none" stroke="${c.border}"/>`,
      `<rect x="${r.x1}" y="${r.y1}" width="${r.x2 - r.x1}" height="${r.y2 - r.y1}" fill="none" stroke="${c.accent}" stroke-width="${B}"/>`,
      `<path fill="${c.accent}" stroke="${c.page}" stroke-width="${RING}" d="M${r.x2 - HS / 2} ${r.y2 - HS / 2}h${HS}v${HS}h-${HS}z"/>`,
      paths(placed.parts),
    ].join(''),
  };
}

// Only when run directly; gen-cards.mjs imports the mark and the wordmark from here.
if (import.meta.main) {
  const report = {};
  report['mark-light.svg'] = emit(
    outDir,
    'mark-light.svg',
    64,
    64,
    'modern-xlsx',
    markBody('light'),
    3072,
  );
  report['mark-dark.svg'] = emit(
    outDir,
    'mark-dark.svg',
    64,
    64,
    'modern-xlsx',
    markBody('dark'),
    3072,
  );
  report['favicon.svg'] = emit(outDir, 'favicon.svg', 16, 16, 'modern-xlsx', favicon(), 3072);
  const hl = hero('light');
  const hd = hero('dark');
  report['hero-light.svg'] = emit(
    outDir,
    'hero-light.svg',
    1280,
    320,
    'modern-xlsx',
    hl.body,
    25600,
  );
  report['hero-dark.svg'] = emit(outDir, 'hero-dark.svg', 1280, 320, 'modern-xlsx', hd.body, 25600);
  console.log(JSON.stringify({ hero: hl.meta, report }, null, 1));
}
