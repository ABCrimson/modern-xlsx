// Architecture family (A3): architecture-light.svg, architecture-dark.svg, 960x1040, shown at width 640.
// Every label was checked against the 1.3.0 code: the WASM exports in crates/modern-xlsx-wasm/src/lib.rs,
// reader.rs and writer.rs in crates/modern-xlsx-core/src, and the TypeScript read and write paths.
// Usage: node gen-architecture.mjs <outDir>
import { writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { emit, n2, PAL, paths, run } from './lib.mjs';

const outDir = process.argv[2] ?? join(import.meta.dirname, 'out');
const W = 960;
const H = 1040;
const S = 30; // minimum label size: 30 units (20 px at 640, 11.2 px at 358)
const M = 56;
const COLW = 404;
const PAD = 20;
const COL = { read: M, write: W - M - COLW }; // read 56..460, write 500..904
const CALL = { read: COL.read + COLW - 24, write: COL.write + 24 }; // inner lanes: calls go down
const RET = { read: M / 2, write: W - M / 2 }; // outer lanes: results come back up
const BAND = { ts: [0, 248], wasm: [248, 408], rust: [408, H] };
const TSBOX = { y: 124, h: 100 };
const RBOX = { y: 480, h: 64, gap: 28 };
const DASH = { y: 852, h: 156 };

const seg = (text, font, colour) => ({ text, font, colour });
const sans = (t, c) => seg(t, 'sans', c);
const mono = (t, c) => seg(t, 'monoMedium', c);

// First option that fits the box (inner width = COLW - 2*PAD); report the widths.
const fits = [];
function pick(name, options, max = COLW - 2 * PAD) {
  for (const o of options) {
    const r = run(o, S);
    if (r.advance <= max) {
      fits.push([name, n2(r.advance), max]);
      return r;
    }
  }
  const r = run(options.at(-1), S);
  fits.push([name, n2(r.advance), max, 'OVERFLOW']);
  return r;
}

function build(theme) {
  const c = PAL[theme];
  const out = [];
  const arrowHead = (x, y, dir) =>
    ({
      down: `M${x} ${y}l-7-12h14z`,
      up: `M${x} ${y}l-7 12h14z`,
      right: `M${x} ${y}l-12-7v14z`,
      left: `M${x} ${y}l12-7v14z`,
    })[dir];
  const lines = [];
  const heads = [];
  const down = (x, y1, y2) => {
    lines.push(`M${x} ${y1}V${y2 - 12}`);
    heads.push(arrowHead(x, y2, 'down'));
  };
  // Return lane: out of a box side at y1, along the outer lane, up, and into the TS box side at y2.
  // The shaft stops at the head's base (12 units outside the box edge), as the down arrows do.
  const ret = (side, boxEdge, lane, y1, y2) => {
    const s = side === 'read' ? -1 : 1;
    const r = 8;
    lines.push(
      `M${boxEdge} ${y1}H${lane - s * r}a${r} ${r} 0 0 ${s < 0 ? 1 : 0} ${s * r}-${r}V${y2 + r}a${r} ${r} 0 0 ${s < 0 ? 1 : 0} ${-s * r}-${r}H${boxEdge + s * 12}`,
    );
    heads.push(arrowHead(boxEdge, y2, s < 0 ? 'right' : 'left'));
  };
  const text = (r, x, baseline, align = 'pen') =>
    out.push(paths(r.place(x, baseline, align).parts));
  const box = (x, y, w, h, stroke) =>
    out.push(
      `<rect x="${x + 1}" y="${y + 1}" width="${w - 2}" height="${h - 2}" rx="8" fill="${c.page}" stroke="${stroke}" stroke-width="2"/>`,
    );
  const label = (x, y, r) => text(r, x + PAD, y + 42);

  // Panel, band separators.
  out.push(
    `<rect x=".5" y=".5" width="${W - 1}" height="${H - 1}" rx="16" fill="${c.panel}" stroke="${c.border}"/>`,
  );
  out.push(
    `<path stroke="${c.border}" d="M1 ${BAND.wasm[0] + 0.5}H${W - 1}M1 ${BAND.rust[0] + 0.5}H${W - 1}"/>`,
  );

  // Band labels (Geist Sans SemiBold 30, +4 %), in the layer colour.
  // The layer on the left, the artefact that implements it on the right, so no call lane crosses a label.
  const bandLabel = (t, artefact, colour, y) => {
    text(run([seg(t, 'sansSemi', colour)], S, { tracking: 0.04 }), M, y);
    text(run([seg(artefact, 'monoMedium', colour)], S), W - M, y, 'ink-right');
  };
  bandLabel('TypeScript API', 'modern-xlsx', c.ts, 56);
  bandLabel('WASM boundary', 'wasm-bindgen', c.wasm, BAND.wasm[0] + 48);
  bandLabel('Rust core', 'modern-xlsx-core', c.rust, BAND.rust[0] + 48);
  // Column titles.
  text(run([seg('READ', 'sansSemi', c.muted)], S, { tracking: 0.08 }), COL.read, 104);
  text(run([seg('WRITE', 'sansSemi', c.muted)], S, { tracking: 0.08 }), COL.write, 104);

  // TypeScript boxes: two lines each.
  for (const k of ['read', 'write']) box(COL[k], TSBOX.y, COLW, TSBOX.h, c.ts);
  const L1 = TSBOX.y + 42;
  const L2 = TSBOX.y + 80; // cap tops ~21 below the box top, ~20 from the last baseline to the box bottom
  text(
    pick('ts.read.1', [[mono('readBuffer', c.ink), sans(' · ', c.muted), mono('readFile', c.ink)]]),
    COL.read + PAD,
    L1,
  );
  text(
    pick('ts.read.2', [[mono('JSON.parse', c.ink), sans(' → ', c.muted), mono('Workbook', c.ink)]]),
    COL.read + PAD,
    L2,
  );
  text(pick('ts.write.1', [[mono('Workbook.toBuffer()', c.ink)]]), COL.write + PAD, L1);
  text(
    pick('ts.write.2', [[sans('→ ', c.muted), mono('JSON.stringify', c.ink)]]),
    COL.write + PAD,
    L2,
  );

  // Rust boxes.
  const readSteps = [
    [[sans('detect ZIP or OLE2', c.ink)]],
    [[sans('unzip within ZIP limits', c.ink)]],
    [
      [sans('SAX-parse each part, ', c.ink), mono('quick-xml', c.ink)],
      [sans('SAX-parse, ', c.ink), mono('quick-xml', c.ink)],
    ],
    [
      [sans('stream JSON ≤ ', c.ink), mono('maxJsonLength', c.ink)],
      [sans('JSON ≤ ', c.ink), mono('maxJsonLength', c.ink)],
    ],
  ];
  const writeSteps = [
    [
      [mono('serde_json', c.ink), sans(' → ', c.muted), mono('WorkbookData', c.ink)],
      [mono('serde_json', c.ink), sans(' decode', c.ink)],
    ],
    [[sans('write parts, SST built inline', c.ink)], [sans('write parts, SST inline', c.ink)]],
    [[sans('zip, Deflate level 6', c.ink)]],
  ];
  const by = (i) => RBOX.y + i * (RBOX.h + RBOX.gap);
  readSteps.forEach((o, i) => {
    box(COL.read, by(i), COLW, RBOX.h, c.rust);
    label(COL.read, by(i), pick(`rust.read.${i}`, o));
  });
  writeSteps.forEach((o, i) => {
    box(COL.write, by(i), COLW, RBOX.h, c.rust);
    label(COL.write, by(i), pick(`rust.write.${i}`, o));
  });

  // Calls: down the inner lanes, through the boundary.
  for (const k of ['read', 'write']) {
    const n = k === 'read' ? readSteps.length : writeSteps.length;
    down(CALL[k], TSBOX.y + TSBOX.h, by(0));
    for (let i = 0; i < n - 1; i++) down(CALL[k], by(i) + RBOX.h, by(i + 1));
  }
  // Results: up the outer lanes, into the TS box at the line they belong to.
  ret('read', COL.read, RET.read, by(readSteps.length - 1) + RBOX.h / 2, L2 - 10);
  ret('write', COL.write + COLW, RET.write, by(writeSteps.length - 1) + RBOX.h / 2, L1 - 10);
  out.push(`<path fill="none" stroke="${c.muted}" stroke-width="2" d="${lines.join('')}"/>`);
  out.push(`<path fill="${c.muted}" d="${heads.join('')}"/>`);

  // What crosses the boundary (code names in Geist Mono).
  // Baselines 296 (band label), 340, 382: pitches 44 and 42, and 26 from the last baseline to the Rust separator,
  // matching the 26.7 from the WASM separator to the band label's cap top. Results share the content margins (M, W - M).
  const R1 = BAND.wasm[0] + 92;
  const R2 = BAND.wasm[0] + 134;
  text(run([mono('read(data, limits)', c.ink)], S), CALL.read - 16, R1, 'ink-right');
  text(run([mono('JSON string', c.ink)], S), M, R2, 'ink-left');
  text(run([mono('write(json)', c.ink)], S), CALL.write + 16, R1, 'ink-left');
  text(run([mono('Uint8Array', c.ink)], S), W - M, R2, 'ink-right');

  // The encryption feature, dashed, across both columns.
  out.push(
    `<rect x="${M + 1}" y="${DASH.y + 1}" width="${W - 2 * M - 2}" height="${DASH.h - 2}" rx="8" fill="none" stroke="${c.rust}" stroke-width="2" stroke-dasharray="6 4"/>`,
  );
  const DW = W - 2 * M - 2 * PAD;
  text(
    pick(
      'dash.1',
      [
        [
          seg('encryption', 'monoMedium', c.ink),
          sans(' feature · full build only, not in ', c.ink),
          mono('/lite', c.ink),
        ],
      ],
      DW,
    ),
    M + PAD,
    DASH.y + 46,
  );
  text(
    pick(
      'dash.2',
      [
        [
          sans('gates ', c.muted),
          mono('readWithPassword', c.muted),
          sans(' · ', c.muted),
          mono('writeWithPassword', c.muted),
        ],
      ],
      DW,
    ),
    M + PAD,
    DASH.y + 88,
  );
  text(
    pick('dash.3', [[sans('decrypts Agile and Standard · encrypts Agile AES-256', c.muted)]], DW),
    M + PAD,
    DASH.y + 130,
  );
  return out.join('');
}

export const ALT =
  'Architecture of modern-xlsx in three layers: the TypeScript API (the modern-xlsx package), the WASM boundary built with wasm-bindgen, and the Rust core (the modern-xlsx-core crate). Read path: readBuffer, or readFile which loads the bytes and calls readBuffer, passes the bytes to the WASM read(data, limits). The Rust core detects whether the file is ZIP or OLE2, unzips it within the ZIP limits, SAX-parses each part with quick-xml and streams a JSON string of at most maxJsonLength bytes back across the boundary, where JSON.parse turns it into a Workbook. Write path: Workbook.toBuffer() passes JSON.stringify output to the WASM write(json); the Rust core decodes it with serde_json into WorkbookData, writes the parts with the shared string table built inline, zips them with Deflate at level 6 and returns a Uint8Array. Encryption is a Cargo feature in the full build only, not in /lite: it gates the WASM exports readWithPassword and writeWithPassword, which decrypt and encrypt as listed.';

// Budget 80 KB raw (about 15 KB gzipped): every label is outlined Geist, so the file is larger than a text-based SVG.
const BUDGET = 81920;
const report = {};
report['architecture-light.svg'] = emit(
  outDir,
  'architecture-light.svg',
  W,
  H,
  'modern-xlsx architecture: read and write paths',
  build('light'),
  BUDGET,
  1,
);
report['architecture-dark.svg'] = emit(
  outDir,
  'architecture-dark.svg',
  W,
  H,
  'modern-xlsx architecture: read and write paths',
  build('dark'),
  BUDGET,
  1,
);
writeFileSync(
  join(outDir, 'architecture-alt.txt'),
  `${ALT}
`,
);
console.log(JSON.stringify({ fits: fits.slice(0, fits.length / 2), report }, null, 1));
