// Shared pieces for the modern-xlsx README / social / docs-site assets:
// palette, Geist outlining (text -> <path>), SVG document shell, svgo pass and rule checks.
import { mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import { findPackageJSON } from 'node:module';
import { dirname, join } from 'node:path';
import { parse } from 'opentype.js/dist/opentype.mjs';
import { optimize } from 'svgo';

// Fonts come from the installed `geist` package next to this script (pinned in package.json).
// (geist has an exports map without ./package.json, so node:module findPackageJSON locates it.)
const FONTS = join(dirname(findPackageJSON('geist', import.meta.url)), 'dist', 'fonts');
const load = (p) => {
  const b = readFileSync(join(FONTS, p));
  return parse(b.buffer.slice(b.byteOffset, b.byteOffset + b.byteLength));
};
export const FONT = {
  mono: load('geist-mono/GeistMono-SemiBold.ttf'),
  monoMedium: load('geist-mono/GeistMono-Medium.ttf'),
  sans: load('geist-sans/Geist-Medium.ttf'),
  sansSemi: load('geist-sans/Geist-SemiBold.ttf'),
};

// The palette (../README.md, "Palette"), including the A1 header-band fills.
export const PAL = {
  light: {
    page: '#FFFFFF',
    panel: '#F6F8FA',
    border: '#D1D9E0',
    ink: '#1F2328',
    muted: '#59636E',
    grid: '#AFB8C1',
    accent: '#C8102E',
    band: '#EAEEF2',
    ts: '#2F6FB3',
    wasm: '#5A46D8',
    rust: '#B7410E',
  },
  dark: {
    page: '#0D1117',
    panel: '#151B23',
    border: '#3D444D',
    ink: '#F0F6FC',
    muted: '#9198A1',
    grid: '#656C76',
    accent: '#FF6B7A',
    band: '#262C36',
    ts: '#6CB6FF',
    wasm: '#A594FF',
    rust: '#F0883E',
  },
};
const ALLOWED = new Set(
  Object.values(PAL)
    .flatMap((p) => Object.values(p))
    .map((h) => h.toUpperCase()),
);

export const n2 = (v) => +v.toFixed(2);
const r1 = (v) => +v.toFixed(1);

// opentype.js 2.0.0: Glyph.getPath() transforms a cached path in place (a second call on the same
// glyph comes back mirrored), and Geist's glyf contours carry no Z, which made svgo 4.1.0 drop
// subpaths. So transform glyph.path.commands (font units, y up) by hand and close every contour.
function glyphD(glyph, x, baseline, scale) {
  const cmds = glyph.path.commands;
  const X = (v) => r1(x + v * scale);
  const Y = (v) => r1(baseline - v * scale);
  let d = '';
  cmds.forEach((c, i) => {
    if (c.type === 'M' && i > 0 && cmds[i - 1].type !== 'Z') d += 'Z';
    if (c.type === 'M' || c.type === 'L') d += `${c.type}${X(c.x)} ${Y(c.y)}`;
    else if (c.type === 'Q') d += `Q${X(c.x1)} ${Y(c.y1)} ${X(c.x)} ${Y(c.y)}`;
    else if (c.type === 'C')
      d += `C${X(c.x1)} ${Y(c.y1)} ${X(c.x2)} ${Y(c.y2)} ${X(c.x)} ${Y(c.y)}`;
    else if (c.type === 'Z') d += 'Z';
  });
  return cmds.length && cmds.at(-1).type !== 'Z' ? `${d}Z` : d;
}

const kern = (font, a, b) => {
  try {
    return font.getKerningValue(a, b) || 0;
  } catch {
    return 0;
  }
};

// Lay out a run of segments [{text, font, colour}] at `size`; tracking in em; nudges = {index: units}
// added after glyph `index` (optical pair kerning). Returns a placer plus measurements.
export function run(segments, size, { tracking = 0, nudges = {} } = {}) {
  const glyphs = [];
  let pen = 0;
  let idx = 0;
  segments.forEach((s) => {
    const font = FONT[s.font];
    const scale = size / font.unitsPerEm;
    let prev = null;
    for (const ch of s.text) {
      const g = font.charToGlyph(ch);
      if (prev) pen += kern(font, prev, g) * scale;
      const m = g.getMetrics();
      const ink = g.path.commands.length > 0;
      glyphs.push({
        g,
        ch,
        x: pen,
        scale,
        colour: s.colour,
        x1: pen + m.xMin * scale,
        x2: pen + m.xMax * scale,
        top: m.yMax * scale,
        bottom: m.yMin * scale,
        ink,
      });
      pen += g.advanceWidth * scale + tracking * size + (nudges[idx] ?? 0);
      prev = g;
      idx++;
    }
  });
  pen -= tracking * size; // no tracking after the last glyph
  const inked = glyphs.filter((g) => g.ink);
  const box = {
    x1: Math.min(...inked.map((g) => g.x1)),
    x2: Math.max(...inked.map((g) => g.x2)),
    top: Math.max(...inked.map((g) => g.top)),
    bottom: Math.min(...inked.map((g) => g.bottom)),
  };
  const gaps = inked.slice(1).map((g, i) => n2(g.x1 - inked[i].x2));
  return {
    advance: pen,
    box,
    gaps,
    inkWidth: box.x2 - box.x1,
    // Place with the pen origin at x (or the ink box left/right/centre with align) and baseline y.
    place(x, baseline, align = 'pen') {
      const dx =
        align === 'ink-left'
          ? x - box.x1
          : align === 'ink-right'
            ? x - box.x2
            : align === 'ink-centre'
              ? x - (box.x1 + box.x2) / 2
              : x;
      const parts = {};
      glyphs.forEach((g) => {
        if (g.ink)
          parts[g.colour] = (parts[g.colour] ?? '') + glyphD(g.g, g.x + dx, baseline, g.scale);
      });
      return { parts, x1: box.x1 + dx, x2: box.x2 + dx };
    },
  };
}
export const paths = (parts) =>
  Object.entries(parts)
    .map(([c, d]) => `<path fill="${c}" d="${d}"/>`)
    .join('');

export function svgDoc(w, h, title, body) {
  return `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ${w} ${h}" width="${w}" height="${h}" role="img" aria-labelledby="t"><title id="t">${title}</title>${body}</svg>`;
}

const svgoCfg = (floatPrecision) => ({
  multipass: true,
  floatPrecision,
  plugins: [
    {
      name: 'preset-default',
      params: {
        overrides: {
          cleanupIds: false,
          removeUnknownsAndDefaults: { keepRoleAttr: true, keepAriaAttrs: true },
          convertColors: { convertCase: false, shorthex: false },
        },
      },
    },
  ],
});

// Optimise, write raw + optimised, and check every rule ../README.md sets for these files.
// precision: decimals svgo keeps (1 = 0.1 unit, under 0.07 px at every size these files are shown).
export function emit(outDir, name, w, h, title, body, budget, precision = 2) {
  mkdirSync(join(outDir, 'raw'), { recursive: true });
  const raw = svgDoc(w, h, title, body);
  const out = optimize(raw, svgoCfg(precision)).data;
  writeFileSync(join(outDir, name), out);
  writeFileSync(join(outDir, 'raw', name), raw);
  const hexes = [...new Set(out.match(/#[0-9A-Fa-f]{3,8}\b/g) ?? [])];
  const bytes = Buffer.byteLength(out);
  const r = {
    raw: Buffer.byteLength(raw),
    bytes,
    budget,
    withinBudget: budget ? bytes < budget : null,
    viewBox: out.includes(`viewBox="0 0 ${w} ${h}"`),
    wh: out.includes(`width="${w}"`) && out.includes(`height="${h}"`),
    title: out.includes(`<title id="t">${title}</title>`),
    role: out.includes('role="img"'),
    labelledby: out.includes('aria-labelledby="t"'),
    xmlns: out.includes('xmlns="http://www.w3.org/2000/svg"'),
    forbidden:
      out.match(
        /<text|<style|style=|class=|<image|<script|<foreignObject|filter|Gradient|href|var\(|@media|font-family|@font-face/g,
      ) ?? [],
    offPalette: hexes.filter((x) => !ALLOWED.has(x.toUpperCase())),
  };
  r.ok =
    r.viewBox &&
    r.wh &&
    r.title &&
    r.role &&
    r.labelledby &&
    r.xmlns &&
    !r.forbidden.length &&
    !r.offPalette.length &&
    r.withinBudget !== false;
  return r;
}
