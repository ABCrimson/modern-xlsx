#!/usr/bin/env bash
# Regenerate every asset, check it, copy it to the deliverable dir, export the PNG cards and render the previews.
# Run from this directory after `npm install --ignore-scripts` and `npx playwright install chromium firefox`
# (package.json pins every package exactly; playwright is a devDependency and renders the previews and PNG cards).
# Usage: bash build.sh [target-dir]   (default: the parent of this directory, which is where the assets live; FINAL=<dir> also works)
set -euo pipefail
S=$(cd "$(dirname "$0")" && pwd); cd "$S"
F=${1:-${FINAL:-$(dirname "$S")}}; P=$F/preview
O=out/final; mkdir -p $O "$P"
node gen-brand.mjs $O > out/report-brand.json
node gen-architecture.mjs $O > out/report-architecture.json
node gen-cards.mjs $O > out/report-cards.json
node -e 'for (const f of process.argv.slice(1)) { const j = JSON.parse(require("fs").readFileSync(f, "utf8")); for (const [k, v] of Object.entries(j.report)) console.log(v.ok ? "ok  " : "FAIL", k, v.bytes, "B (raw", v.raw + ", budget", (v.budget ?? "-") + ")", v.forbidden.join(","), v.offPalette.join(",")); }' out/report-*.json
node diff.mjs $O mark-light.svg mark-dark.svg favicon.svg hero-light.svg hero-dark.svg architecture-light.svg architecture-dark.svg social-preview.svg og.svg
for f in mark-light mark-dark favicon hero-light hero-dark architecture-light architecture-dark social-preview og; do cp $O/$f.svg "$F/$f.svg"; done
cp $O/architecture-alt.txt "$F/architecture-alt.txt"
node mock.mjs "$F" out
node -e '
const [F, P] = process.argv.slice(1); const j = [];
const add = (svg, width, bg, out, o = {}) => j.push({ svg: F + "/" + svg, width, bg, out: P + "/" + out, pad: 16, ...o });
for (const t of ["light", "dark"]) { const bg = t === "light" ? "#FFFFFF" : "#0D1117"; for (const s of [64, 32, 16]) add(`mark-${t}.svg`, s, bg, `mark-${t}-${s}.png`); add(`mark-${t}.svg`, 32, bg, `mark-${t}-32@2x.png`, { scale: 2 }); }
for (const bg of ["#FFFFFF", "#1F2328", "#DFE1E5", "#35363A"]) { const n = bg.slice(1); add("favicon.svg", 16, bg, `favicon-16-on-${n}.png`, { pad: 8 }); add("favicon.svg", 32, bg, `favicon-32-on-${n}.png`, { pad: 8 }); add("favicon.svg", 16, bg, `favicon-16@2x-on-${n}.png`, { pad: 8, scale: 2 }); }
for (const w of [838, 358]) { add("hero-light.svg", w, "#FFFFFF", `hero-light-${w}-on-white.png`); add("hero-dark.svg", w, "#0D1117", `hero-dark-${w}-on-dark.png`); add("hero-light.svg", w, "#0D1117", `hero-light-${w}-on-dark.png`); }
add("hero-light.svg", 358, "#FFFFFF", "hero-light-358@3x-on-white.png", { scale: 3 }); add("hero-dark.svg", 358, "#0D1117", "hero-dark-358@3x-on-dark.png", { scale: 3 });
for (const w of [1012, 1280]) { add("hero-light.svg", w, "#FFFFFF", `hero-light-${w}-on-white.png`); add("hero-dark.svg", w, "#0D1117", `hero-dark-${w}-on-dark.png`); }
for (const w of [838, 358]) { add("hero-light.svg", w, "#FFFFFF", `firefox-hero-light-${w}-on-white.png`, { browser: "firefox" }); add("hero-dark.svg", w, "#0D1117", `firefox-hero-dark-${w}-on-dark.png`, { browser: "firefox" }); }
for (const w of [640, 358]) { add("architecture-light.svg", w, "#FFFFFF", `architecture-light-${w}-on-white.png`); add("architecture-dark.svg", w, "#0D1117", `architecture-dark-${w}-on-dark.png`); }
add("architecture-light.svg", 358, "#FFFFFF", "architecture-light-358@3x-on-white.png", { scale: 3 });
// The PNG cards: viewport-size screenshots at deviceScaleFactor 1, no padding.
j.push({ svg: F + "/social-preview.svg", width: 1280, bg: "#0D1117", pad: 0, out: F + "/social-preview.png" });
j.push({ svg: F + "/og.svg", width: 1200, bg: "#0D1117", pad: 0, out: F + "/og.png" });
for (const t of ["light", "dark"]) { j.push({ html: "out/readme-" + t + ".html", viewport: 1280, full: true, out: P + "/readme-mock-" + t + "-1280.png" }); j.push({ html: "out/readme-" + t + ".html", viewport: 390, full: true, out: P + "/readme-mock-" + t + "-390.png", scale: 2 }); }
require("fs").writeFileSync("out/jobs.json", JSON.stringify(j));' "$F" "$P"
node render.mjs out/jobs.json > out/render.log
grep -v '^wrote' out/render.log
# 2x crops of the return-arrow entries and the WASM band (zoom.mjs takes a region in viewBox units).
node zoom.mjs "$F/architecture-light.svg" 640 '#FFFFFF' "$P/architecture-light-640@2x-top-on-white.png" 0 0 960 420 2
node zoom.mjs "$F/architecture-dark.svg" 640 '#0D1117' "$P/architecture-dark-640@2x-top-on-dark.png" 0 0 960 420 2
node -e 'const fs = require("fs"); for (const f of process.argv.slice(1)) { const b = fs.readFileSync(f); console.log(f.split("/").pop(), b.readUInt32BE(16) + "x" + b.readUInt32BE(20), b.length, "B", b.length < 300 * 1024 ? "< 300 KB" : "OVER 300 KB"); }' "$F/social-preview.png" "$F/og.png"
node measure.mjs "$F/hero-light.svg" "$F/hero-dark.svg" > out/measure.json
node -e 'const j = JSON.parse(require("fs").readFileSync("out/measure.json", "utf8")); for (const [k, v] of Object.entries(j)) console.log(k.split("/").pop(), "ink", v.ink.join("-"), "gaps", v.gaps.join(" "), "| corner", v.diag.split(" ").slice(14, 20).join(" "));'
