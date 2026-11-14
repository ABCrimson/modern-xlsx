// Card family (A4 + docs site): social-preview.svg (1280x640) and og.svg (1200x630), dark, solid background.
// Mark above the wordmark on the left, a small active-cell sheet on the right. The PNG cards are exported by build.sh (render.mjs).
// Usage: node gen-cards.mjs <outDir>
import { join } from 'node:path';
import { markAt, wordmark } from './gen-brand.mjs';
import { emit, n2, PAL, paths, run } from './lib.mjs';

const outDir = process.argv[2] ?? join(import.meta.dirname, 'out');
const c = PAL.dark;

// L: x0 left edge of the content, y0 top of the mark frame, card [cx1, cx2], dy shifts all baselines.
function card(W, H, L) {
  const o = [];
  const T = (r, x, y, align = 'pen') => o.push(paths(r.place(x, y, align).parts));
  o.push(`<rect width="${W}" height="${H}" fill="${c.page}"/>`);
  // The mark at x2 (frame 108 px outside, every edge on a whole pixel); the frame's outer edge sits at (x0, y0).
  o.push(markAt('dark', L.x0 - 8, L.y0 - 8, 2));
  // Wordmark, Geist Mono SemiBold 84, hyphen in accent, ink box flush with the mark's frame.
  const wm = wordmark(c, 84);
  T(wm, L.x0, L.y0 + 210, 'ink-left');
  // Tagline, Geist Sans Medium 40, two lines; meta line, Geist Mono 28 muted.
  T(
    run([{ text: 'Read and write .xlsx from', font: 'sans', colour: c.ink }], 40),
    L.x0,
    L.y0 + 290,
    'ink-left',
  );
  T(
    run([{ text: 'JavaScript and TypeScript.', font: 'sans', colour: c.ink }], 40),
    L.x0,
    L.y0 + 342,
    'ink-left',
  );
  T(
    run(
      [
        {
          text: 'npm i modern-xlsx · Rust → WebAssembly · MIT',
          font: 'monoMedium',
          colour: c.muted,
        },
      ],
      28,
    ),
    L.x0,
    L.y0 + 410,
    'ink-left',
  );

  // The sheet card: name box + formula bar, column headers, 4 rows; A1 active and holding "Hello".
  const [x1, x2] = L.card;
  const y1 = L.y0;
  const FB = 64;
  const CH = 40;
  const RW = 48;
  const ROW = 64;
  const COLW = L.colw;
  const gy = y1 + FB;
  const ry = gy + CH;
  const y2 = ry + 4 * ROW;
  const cols = [x1 + RW, x1 + RW + COLW, x1 + RW + 2 * COLW];
  o.push(
    `<clipPath id="k"><rect x="${x1 + 1}" y="${y1 + 1}" width="${x2 - x1 - 2}" height="${y2 - y1 - 2}" rx="15.5"/></clipPath>`,
  );
  o.push(
    `<rect x="${x1 + 0.5}" y="${y1 + 0.5}" width="${x2 - x1 - 1}" height="${y2 - y1 - 1}" rx="16" fill="${c.panel}"/>`,
  );
  o.push(`<g clip-path="url(#k)">`);
  o.push(`<path fill="${c.band}" d="M${x1} ${gy}H${x2}V${ry}H${x1 + RW}V${y2}H${x1}z"/>`);
  o.push(
    `<path fill="${c.border}" d="M${cols[0]} ${gy}h${COLW}v${CH}h-${COLW}zM${x1} ${ry}h${RW}v${ROW}h-${RW}z"/>`,
  ); // selected column A and row 1 in the headers
  o.push(
    `<rect x="${cols[0]}" y="${ry}" width="${x2 - cols[0]}" height="${y2 - ry}" fill="${c.page}"/>`,
  );
  const grid = [
    ...cols.map((x) => `M${x + 0.5} ${gy}V${y2}`),
    ...[0, 1, 2, 3].map((i) => `M${x1} ${ry + i * ROW + 0.5}H${x2}`),
    `M${x1 + 108.5} ${y1}V${gy}`,
  ].join('');
  o.push(`<path stroke="${c.grid}" d="${grid}"/>`);
  o.push(`<path stroke="${c.border}" d="M${x1} ${gy + 0.5}H${x2}"/>`);
  o.push(`</g>`);
  o.push(
    `<rect x="${x1 + 0.5}" y="${y1 + 0.5}" width="${x2 - x1 - 1}" height="${y2 - y1 - 1}" rx="16" fill="none" stroke="${c.border}"/>`,
  );
  // Formula bar: name box "A1", then the cell's content.
  const fbBase = y1 + FB / 2 + 10;
  T(run([{ text: 'A1', font: 'monoMedium', colour: c.muted }], 28), x1 + 54, fbBase, 'ink-centre');
  T(run([{ text: 'fx', font: 'sans', colour: c.muted }], 28), x1 + 128, fbBase, 'ink-left');
  T(run([{ text: 'Hello', font: 'monoMedium', colour: c.ink }], 28), x1 + 170, fbBase, 'ink-left');
  // Headers (28), centred in their cells. The selected column A and row 1 sit on the #3D444D highlight, so they are
  // drawn in ink (9.05:1 there; muted would be 3.38:1); the others stay muted on the band (4.82:1).
  const hdr = (selected) => (selected ? c.ink : c.muted);
  for (const [i, l] of ['A', 'B'].entries()) {
    T(
      run([{ text: l, font: 'sans', colour: hdr(i === 0) }], 28),
      cols[i] + COLW / 2,
      gy + CH / 2 + 10,
      'ink-centre',
    );
  }
  T(
    run([{ text: 'C', font: 'sans', colour: c.muted }], 28),
    (cols[2] + x2) / 2,
    gy + CH / 2 + 10,
    'ink-centre',
  );
  for (const r of [1, 2, 3, 4]) {
    T(
      run([{ text: String(r), font: 'sans', colour: hdr(r === 1) }], 28),
      x1 + RW / 2,
      ry + (r - 1) * ROW + ROW / 2 + 10,
      'ink-centre',
    );
  }
  // Cell content, Geist Sans SemiBold 40 (11 px at a 358 px wide card): Hello left, 42 right with 16 padding.
  const cb = ry + ROW / 2 + 14;
  T(run([{ text: 'Hello', font: 'sansSemi', colour: c.ink }], 40), cols[0] + 16, cb, 'ink-left');
  T(run([{ text: '42', font: 'sansSemi', colour: c.ink }], 40), cols[2] - 16, cb, 'ink-right');
  // Active A1 = the mark x1: border 4, handle red 12 with a 2 px page-colour ring.
  o.push(
    `<rect x="${cols[0]}" y="${ry}" width="${COLW}" height="${ROW}" fill="none" stroke="${c.accent}" stroke-width="4"/>`,
  );
  o.push(
    `<path fill="${c.accent}" stroke="${c.page}" stroke-width="2" d="M${cols[1] - 7} ${ry + ROW - 7}h14v14h-14z"/>`,
  );
  return {
    body: o.join(''),
    meta: { card: [x1, x2, y1, y2], wordmark: [n2(wm.box.x1), n2(wm.box.x2)] },
  };
}

const s = card(1280, 640, { x0: 96, y0: 120, card: [760, 1184], colw: 160 });
const g = card(1200, 630, { x0: 88, y0: 116, card: [712, 1112], colw: 152 });
const report = {
  'social-preview.svg': emit(
    outDir,
    'social-preview.svg',
    1280,
    640,
    'modern-xlsx: read and write .xlsx from JavaScript and TypeScript',
    s.body,
  ),
  'og.svg': emit(
    outDir,
    'og.svg',
    1200,
    630,
    'modern-xlsx: read and write .xlsx from JavaScript and TypeScript',
    g.body,
  ),
};
console.log(JSON.stringify({ social: s.meta, og: g.meta, report }, null, 1));
