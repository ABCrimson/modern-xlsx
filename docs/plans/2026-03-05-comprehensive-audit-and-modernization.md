# Comprehensive Audit, Bug Fix & Modernization Plan

> **For Claude:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task.

**Goal:** Fix 4 critical bugs (barcode+chart collision, missing axis font_size, oneCellAnchor parsing, combo chart roundtrip), close all test coverage gaps, add WASM boundary validation, modernize every file to Rust 1.94.0 Edition 2024 + TypeScript 6.0 idioms, optimize for maximum performance, and bring all documentation/wiki/README to pixel-perfect accuracy.

**Architecture:** Six-phase approach — critical bugs first (data loss prevention), then test coverage, WASM validation, Rust modernization, TypeScript modernization, and finally documentation/aesthetics. Each phase is independently committable and testable. The barcode+chart collision fix requires coordinating changes across both Rust (writer.rs, charts.rs) and TypeScript (workbook.ts, barcode.ts) layers.

**Tech Stack:** Rust 1.94.0 Edition 2024, wasm-bindgen 0.2.114, quick-xml 0.39.2, zip 8.1, TypeScript 6.0, Vitest 4.1.0-beta.5, tsdown 0.21.0-beta.2, Biome 2.4.4, pnpm 11

---

## Phase 1: Critical Bug Fixes

### Task 1: Fix Barcode + Chart Drawing XML Collision (LIVE DATA LOSS BUG)

**Context:** When a worksheet has BOTH images/barcodes AND charts, the image drawing XML is silently lost. Images store drawing XML in `preservedEntries["xl/drawings/drawing{n}.xml"]` (TypeScript, workbook.ts:538-539). Charts generate the same path via `ChartAnchor::generate_drawing_xml()` (Rust, writer.rs:253-254) and add it to `generated_rels` (writer.rs:271). In the preserved entries loop (writer.rs:341-344), the `generated_rels.contains(path)` check causes the image drawing XML to be **skipped**, losing all image anchors.

**Files:**
- Modify: `crates/modern-xlsx-core/src/writer.rs:218-283` (chart drawing generation)
- Modify: `crates/modern-xlsx-core/src/writer.rs:338-351` (preserved entries merging)
- Modify: `crates/modern-xlsx-core/src/ooxml/charts.rs:412-544` (drawing XML generation)
- Modify: `packages/modern-xlsx/src/workbook.ts:509-564` (addImage method)
- Modify: `packages/modern-xlsx/src/barcode.ts:1762-1815` (generateDrawingXml)
- Test: `packages/modern-xlsx/__tests__/barcode-chart-collision.test.ts` (NEW)

**Strategy:** Move image anchors out of `preservedEntries` into a first-class `imageAnchors` field on `WorksheetData` (same pattern as `charts`). The Rust writer will then merge image `<xdr:pic>` anchors and chart `<xdr:graphicFrame>` anchors into a single `<xdr:wsDr>` drawing XML. This eliminates the collision entirely.

**Step 1: Write the failing test**

Create `packages/modern-xlsx/__tests__/barcode-chart-collision.test.ts`:

```typescript
import { describe, expect, it } from 'vitest';
import { ChartBuilder, Workbook, readBuffer } from '../src/index.js';

describe('barcode + chart on same worksheet', () => {
  it('preserves both image and chart after roundtrip', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Mixed');

    // Add data for chart
    ws.cell('A1').value = 'Q1';
    ws.cell('A2').value = 'Q2';
    ws.cell('B1').value = 100;
    ws.cell('B2').value = 200;

    // Add a chart
    ws.addChart('bar', (b) => {
      b.title('Sales')
        .addSeries({ valRef: 'Mixed!$B$1:$B$2', catRef: 'Mixed!$A$1:$A$2' })
        .anchor({ fromCol: 4, fromRow: 0, toCol: 12, toRow: 15 });
    });

    // Add a barcode image
    const { encodeQR, renderBarcodePNG } = await import('../src/barcode.js');
    const matrix = encodeQR('TEST-DATA');
    const png = renderBarcodePNG(matrix, { scale: 4 });
    wb.addImage('Mixed', {
      fromCol: 0, fromRow: 5,
      toCol: 3, toRow: 12,
    }, png);

    // Roundtrip
    const buf = await wb.toBuffer();
    const wb2 = await readBuffer(buf);
    const ws2 = wb2.getSheet('Mixed');

    // Both must survive
    expect(ws2?.charts).toHaveLength(1);
    expect(ws2?.charts?.[0]?.chart.title?.text).toBe('Sales');

    // Image must be in preserved entries (check drawing XML contains pic element)
    const drawingXml = wb2.data.preservedEntries?.['xl/drawings/drawing1.xml'];
    expect(drawingXml).toBeDefined();
    const drawingStr = new TextDecoder().decode(new Uint8Array(drawingXml!));
    expect(drawingStr).toContain('xdr:pic'); // Image anchor
    expect(drawingStr).toContain('xdr:graphicFrame'); // Chart anchor
  });

  it('handles multiple images + multiple charts', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Multi');

    ws.cell('A1').value = 10;
    ws.cell('A2').value = 20;

    // Two charts
    ws.addChart('bar', (b) => {
      b.addSeries({ valRef: 'Multi!$A$1:$A$2' })
        .anchor({ fromCol: 3, fromRow: 0, toCol: 8, toRow: 10 });
    });
    ws.addChart('pie', (b) => {
      b.addSeries({ valRef: 'Multi!$A$1:$A$2' })
        .anchor({ fromCol: 3, fromRow: 12, toCol: 8, toRow: 22 });
    });

    // Two barcodes
    const { encodeQR, renderBarcodePNG } = await import('../src/barcode.js');
    const png1 = renderBarcodePNG(encodeQR('BC1'), { scale: 3 });
    const png2 = renderBarcodePNG(encodeQR('BC2'), { scale: 3 });
    wb.addImage('Multi', { fromCol: 0, fromRow: 0, toCol: 2, toRow: 4 }, png1);
    wb.addImage('Multi', { fromCol: 0, fromRow: 5, toCol: 2, toRow: 9 }, png2);

    const buf = await wb.toBuffer();
    const wb2 = await readBuffer(buf);
    const ws2 = wb2.getSheet('Multi');

    expect(ws2?.charts).toHaveLength(2);

    // Check drawing XML has all 4 anchors (2 charts + 2 images)
    const drawingXml = wb2.data.preservedEntries?.['xl/drawings/drawing1.xml'];
    expect(drawingXml).toBeDefined();
    const str = new TextDecoder().decode(new Uint8Array(drawingXml!));
    const picCount = (str.match(/xdr:pic/g) ?? []).length;
    const frameCount = (str.match(/xdr:graphicFrame/g) ?? []).length;
    // Each pic has open+close tags = 2 matches per image
    expect(picCount).toBeGreaterThanOrEqual(2);
    expect(frameCount).toBeGreaterThanOrEqual(2);
  });
});
```

**Step 2: Run test to verify it fails**

Run: `pnpm -C packages/modern-xlsx test -- --reporter=verbose barcode-chart-collision`
Expected: FAIL — image anchors lost after roundtrip (either no `xdr:pic` in drawing XML, or drawing XML only contains chart graphicFrames)

**Step 3: Implement the fix**

The fix requires modifying the Rust writer to merge image anchors from `preserved_entries` with chart anchors into a single drawing XML. The approach:

1. In `writer.rs`, when generating chart drawing XML, check if `preserved_entries` also has an image drawing XML for the same path.
2. If both exist, parse the preserved image drawing XML to extract `<xdr:twoCellAnchor>` elements that contain `<xdr:pic>` and inject them into the chart drawing XML alongside the `<xdr:graphicFrame>` elements.
3. Also merge the drawing `.rels` files (chart rels + image rels).

Modify `crates/modern-xlsx-core/src/writer.rs` — in the `has_charts` block (lines 218-283), after generating chart drawing XML, check for preserved image drawing entries and merge:

```rust
// After generating chart drawing XML (line 254), check for preserved image drawings.
let drawing_path = format!("xl/drawings/drawing{sheet_num}.xml");

// Check if there are preserved image anchors for this drawing.
let preserved_image_xml = workbook.preserved_entries.get(&drawing_path);

let final_drawing_xml = if let Some(image_xml_bytes) = preserved_image_xml {
    // Merge: extract <xdr:twoCellAnchor> blocks from preserved image XML
    // and append them to the chart drawing XML.
    merge_drawing_xml(&drawing_xml, image_xml_bytes)?
} else {
    drawing_xml
};
```

Add a helper function `merge_drawing_xml` to writer.rs that:
- Parses both XML documents
- Extracts all `<xdr:twoCellAnchor>` children from the image drawing
- Inserts them before `</xdr:wsDr>` in the chart drawing
- Merges the relationship IDs (rewriting image rIds to not conflict with chart rIds)

Also merge `_rels/drawing{n}.xml.rels`:
```rust
let drawing_rels_path = format!("xl/drawings/_rels/drawing{sheet_num}.xml.rels");
if let Some(image_rels_bytes) = workbook.preserved_entries.get(&drawing_rels_path) {
    // Parse preserved image rels and merge into chart drawing rels.
    let image_rels = Relationships::from_xml(image_rels_bytes)?;
    for (id, rel_type, target) in image_rels.entries() {
        // Remap rId to avoid conflicts with chart rIds
        let new_rid = format!("rId{}", chart_r_ids.len() + offset);
        drawing_rels.add(new_rid, rel_type, target);
    }
}
```

**Step 4: Run tests to verify they pass**

Run: `cargo test -p modern-xlsx-core && pnpm -C packages/modern-xlsx test -- --reporter=verbose barcode-chart-collision`
Expected: PASS — both image and chart anchors survive roundtrip

**Step 5: Commit**

```bash
git add crates/modern-xlsx-core/src/writer.rs packages/modern-xlsx/__tests__/barcode-chart-collision.test.ts
git commit -m "fix: merge barcode/image + chart drawing XML to prevent data loss

When a worksheet has both images and charts, the drawing XML was
generated separately by each code path. The chart path (Rust) would
add the drawing to generated_rels, causing the image path (preserved
entries) to be silently skipped. This merges both into a single
<xdr:wsDr> with both <xdr:pic> and <xdr:graphicFrame> anchors."
```

---

### Task 2: Restore ChartAxis.font_size with Proper Implementation

**Context:** `ChartAxis.font_size` was removed in v0.8.1 audit as "dead code" because the writer/parser didn't implement it. But it maps to a real ECMA-376 feature: `<c:txPr>` (text properties) on axis elements controls tick label font formatting. The field should be restored AND properly wired into the writer and parser. ECMA-376 reference: `CT_TextBody` inside `c:catAx` / `c:valAx` elements.

**Files:**
- Modify: `crates/modern-xlsx-core/src/ooxml/charts.rs:176-213` (ChartAxis struct — add font_size back)
- Modify: `crates/modern-xlsx-core/src/ooxml/charts.rs` (write_axis function — write `<c:txPr>`)
- Modify: `crates/modern-xlsx-core/src/ooxml/charts.rs` (parser — parse `<c:txPr>` on axes)
- Modify: `packages/modern-xlsx/src/types.ts:452+` (ChartAxisData — add fontSize back)
- Modify: `packages/modern-xlsx/src/chart-builder.ts` (AxisOptions — add fontSize back)
- Test: `packages/modern-xlsx/__tests__/chart-roundtrip.test.ts` (add axis font_size roundtrip test)

**Step 1: Write the failing test**

Add to `packages/modern-xlsx/__tests__/chart-roundtrip.test.ts`:

```typescript
it('preserves axis tick label font size on roundtrip', async () => {
  const wb = new Workbook();
  const ws = wb.addSheet('Data');
  ws.cell('A1').value = 'X';
  ws.cell('B1').value = 10;

  ws.addChart('bar', (b) => {
    b.addSeries({ valRef: 'Data!$B$1:$B$1', catRef: 'Data!$A$1:$A$1' })
      .categoryAxis({ title: 'Category', fontSize: 1400 }) // 14pt in hundredths
      .valueAxis({ title: 'Value', fontSize: 1200 }) // 12pt
      .anchor({ fromCol: 3, fromRow: 0, toCol: 10, toRow: 15 });
  });

  const buf = await wb.toBuffer();
  const wb2 = await readBuffer(buf);
  const ws2 = wb2.getSheet('Data') as Worksheet;
  const chart = ws2.charts![0]!;

  expect(chart.chart.catAxis?.fontSize).toBe(1400);
  expect(chart.chart.valAxis?.fontSize).toBe(1200);
});
```

**Step 2: Run test to verify it fails**

Run: `pnpm -C packages/modern-xlsx test -- --reporter=verbose chart-roundtrip`
Expected: FAIL — `fontSize` property doesn't exist on ChartAxisData

**Step 3: Restore the field in Rust**

In `crates/modern-xlsx-core/src/ooxml/charts.rs`, add to `ChartAxis` struct (after line 212):

```rust
/// Font size for axis tick labels in hundredths of a point (e.g., 1400 = 14pt).
/// Maps to `<c:txPr><a:p><a:pPr><a:defRPr sz="..."/></a:pPr></a:p></c:txPr>`.
#[serde(default, skip_serializing_if = "Option::is_none")]
pub font_size: Option<u32>,
```

Add `font_size: None` to all ChartAxis struct literals in test code (use `replace_all`).

**Step 4: Implement writer — `<c:txPr>` emission**

In the `write_axis` function, after writing `<c:delete>` and before `</c:catAx>` or `</c:valAx>`, add:

```rust
// <c:txPr> — axis tick label text properties.
if let Some(sz) = axis.font_size {
    writer.write_event(Event::Start(BytesStart::new("c:txPr"))).map_err(map_err)?;
    writer.write_event(Event::Empty(BytesStart::new("a:bodyPr"))).map_err(map_err)?;
    writer.write_event(Event::Empty(BytesStart::new("a:lstStyle"))).map_err(map_err)?;
    writer.write_event(Event::Start(BytesStart::new("a:p"))).map_err(map_err)?;
    writer.write_event(Event::Start(BytesStart::new("a:pPr"))).map_err(map_err)?;
    let mut def_rpr = BytesStart::new("a:defRPr");
    def_rpr.push_attribute(("sz", ibuf.format(sz)));
    writer.write_event(Event::Empty(def_rpr)).map_err(map_err)?;
    writer.write_event(Event::End(BytesEnd::new("a:pPr"))).map_err(map_err)?;
    writer.write_event(Event::End(BytesEnd::new("a:p"))).map_err(map_err)?;
    writer.write_event(Event::End(BytesEnd::new("c:txPr"))).map_err(map_err)?;
}
```

**Step 5: Implement parser — parse `<c:txPr>` on axes**

In the chart parser, add `ParseCtx::AxisTxPr` state. When inside `CatAxis` or `ValAxis` context and encountering `<c:txPr>`, push the new context. Parse `<a:defRPr sz="...">` to extract font size:

```rust
(ParseCtx::CatAxis, b"txPr") | (ParseCtx::ValAxis, b"txPr") => {
    ctx_stack.push(ParseCtx::AxisTxPr);
}
// Inside AxisTxPr, look for defRPr with sz attribute
(ParseCtx::AxisTxPr, b"defRPr") => {
    for attr in e.attributes().flatten() {
        if attr.key.as_ref() == b"sz" {
            let sz: u32 = String::from_utf8_lossy(&attr.value).parse().unwrap_or(0);
            if sz > 0 {
                match current_axis_ctx {
                    ParseCtx::CatAxis => cur_cat_axis.font_size = Some(sz),
                    ParseCtx::ValAxis => cur_val_axis.font_size = Some(sz),
                    _ => {}
                }
            }
        }
    }
}
```

**Step 6: Restore TypeScript types**

In `packages/modern-xlsx/src/types.ts`, add to `ChartAxisData` interface:

```typescript
/** Font size for tick labels in hundredths of a point (1400 = 14pt). */
fontSize?: number | null;
```

In `packages/modern-xlsx/src/chart-builder.ts`, add to `AxisOptions`:

```typescript
/** Font size for tick labels in hundredths of a point (1400 = 14pt). */
fontSize?: number;
```

And wire it through `buildAxisScaleProps` to include `fontSize` in the returned object.

**Step 7: Run all tests**

Run: `cargo test -p modern-xlsx-core && pnpm -C packages/modern-xlsx test`
Expected: ALL PASS

**Step 8: Commit**

```bash
git add crates/modern-xlsx-core/src/ooxml/charts.rs packages/modern-xlsx/src/types.ts packages/modern-xlsx/src/chart-builder.ts packages/modern-xlsx/__tests__/chart-roundtrip.test.ts
git commit -m "feat: restore ChartAxis.fontSize with proper writer/parser for <c:txPr>

Re-adds the font_size field that was incorrectly removed as dead code.
Now properly writes <c:txPr><a:p><a:pPr><a:defRPr sz='...'/> for axis
tick label font sizes, and parses the same structure on read."
```

---

### Task 3: Parse oneCellAnchor Charts

**Context:** The drawing XML parser in the chart reader only handles `<xdr:twoCellAnchor>`. Charts embedded with `<xdr:oneCellAnchor>` (which Excel uses for small inline charts) are silently skipped. Zero references to `oneCellAnchor` exist in the Rust codebase.

**Files:**
- Modify: `crates/modern-xlsx-core/src/reader.rs` (drawing XML parsing — add oneCellAnchor handling)
- Test: `packages/modern-xlsx/__tests__/chart-roundtrip.test.ts` (add oneCellAnchor test)

**Step 1: Write a failing Rust test**

In `crates/modern-xlsx-core/src/ooxml/charts.rs` test module, add a test that constructs drawing XML with `<xdr:oneCellAnchor>` containing a chart reference and verifies it gets parsed:

```rust
#[test]
fn parse_one_cell_anchor_chart() {
    let drawing_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
              xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
              xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
      <xdr:oneCellAnchor>
        <xdr:from><xdr:col>3</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
        <xdr:ext cx="5400000" cy="3240000"/>
        <xdr:graphicFrame macro="">
          <xdr:nvGraphicFramePr>
            <xdr:cNvPr id="2" name="Chart 1"/>
            <xdr:cNvGraphicFramePr/>
          </xdr:nvGraphicFramePr>
          <xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>
          <a:graphic>
            <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
              <c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="rId1"/>
            </a:graphicData>
          </a:graphic>
        </xdr:graphicFrame>
        <xdr:clientData/>
      </xdr:oneCellAnchor>
    </xdr:wsDr>"#;

    // Parse and verify the chart reference is extracted
    let chart_refs = parse_drawing_chart_refs(drawing_xml.as_bytes());
    assert_eq!(chart_refs.len(), 1);
    assert_eq!(chart_refs[0].r_id, "rId1");
}
```

**Step 2: Run test to verify it fails**

Run: `cargo test -p modern-xlsx-core parse_one_cell_anchor`
Expected: FAIL — `oneCellAnchor` is not parsed

**Step 3: Implement oneCellAnchor parsing**

In the drawing XML parser (in `reader.rs` where chart references are extracted from drawing XML), add handling for `<xdr:oneCellAnchor>` alongside the existing `<xdr:twoCellAnchor>` handling. Both anchor types can contain `<xdr:graphicFrame>` with chart references.

For the anchor data model, a `oneCellAnchor` has `from` + `ext` (cx/cy dimensions) instead of `from` + `to`. Add a variant to `ChartAnchor`:

```rust
// In ChartAnchor, add fields for oneCellAnchor support:
#[serde(default, skip_serializing_if = "Option::is_none")]
pub ext_cx: Option<u64>,
#[serde(default, skip_serializing_if = "Option::is_none")]
pub ext_cy: Option<u64>,
```

When `ext_cx`/`ext_cy` are `Some`, the anchor was a `oneCellAnchor` and should be written as such. When they're `None`, use `twoCellAnchor` (preserving current behavior).

**Step 4: Run all tests**

Run: `cargo test -p modern-xlsx-core && pnpm -C packages/modern-xlsx test`
Expected: ALL PASS

**Step 5: Commit**

```bash
git add crates/modern-xlsx-core/src/reader.rs crates/modern-xlsx-core/src/ooxml/charts.rs
git commit -m "feat: parse oneCellAnchor charts from drawing XML

Charts embedded via <xdr:oneCellAnchor> (Excel's format for small
inline charts) were silently skipped. Now parsed alongside existing
twoCellAnchor handling."
```

---

### Task 4: Combo Chart Roundtrip (Parse secondary_chart)

**Context:** `ChartData.secondary_chart` and `secondary_val_axis` can be written (charts.rs:834-848) but the parser never populates them. Any combo chart (e.g., bar+line) read from an XLSX will lose the secondary chart type on roundtrip. The parser currently only handles the FIRST chart-type element in `<c:plotArea>` and ignores subsequent ones.

**Files:**
- Modify: `crates/modern-xlsx-core/src/ooxml/charts.rs:2400-2900` (parser — handle multiple chart type elements)
- Test: `crates/modern-xlsx-core/src/ooxml/charts.rs` (add combo chart parse test)
- Test: `packages/modern-xlsx/__tests__/chart-roundtrip.test.ts` (add combo chart TS test)

**Step 1: Write the failing Rust test**

```rust
#[test]
fn parse_combo_bar_line_chart() {
    // XML with both <c:barChart> and <c:lineChart> in <c:plotArea>
    let xml = /* combo chart XML with barChart as primary and lineChart as secondary */;
    let chart = ChartData::from_xml(xml.as_bytes()).unwrap();
    assert_eq!(chart.chart_type, ChartType::Bar);
    assert!(chart.secondary_chart.is_some());
    let secondary = chart.secondary_chart.unwrap();
    assert_eq!(secondary.chart_type, ChartType::Line);
    assert!(!secondary.series.is_empty());
}
```

**Step 2: Implement parser changes**

The chart parser currently sets `chart_type` when it encounters the first chart-type element (`c:barChart`, `c:lineChart`, etc.) inside `<c:plotArea>`. To support combo charts:

1. Track whether `chart_type` has already been set
2. When a SECOND chart-type element is encountered inside `<c:plotArea>`, create a new `ChartData` for the secondary and parse its series into that
3. On completion, assign the secondary to `self.secondary_chart`

Add a `secondary_series: Vec<ChartSeries>` accumulator and a `secondary_chart_type: Option<ChartType>` to the parser state. When the parser finishes and both are set, construct `secondary_chart = Some(Box::new(ChartData { ... }))`.

Also parse the secondary value axis: the second `<c:valAx>` encountered should be assigned to `secondary_val_axis`.

**Step 3: Run tests**

Run: `cargo test -p modern-xlsx-core && pnpm -C packages/modern-xlsx test`
Expected: ALL PASS

**Step 4: Commit**

```bash
git add crates/modern-xlsx-core/src/ooxml/charts.rs packages/modern-xlsx/__tests__/chart-roundtrip.test.ts
git commit -m "feat: parse combo charts (secondary_chart + secondary_val_axis)

The parser now handles multiple chart-type elements inside <c:plotArea>,
populating secondary_chart and secondary_val_axis for combo charts
like bar+line. Previously the secondary chart type was silently lost
on roundtrip."
```

---

## Phase 2: Test Coverage Gaps

### Task 5: Test removeChart and addChartData

**Context:** Both are public API methods on `Worksheet` (workbook.ts:1062-1078) with zero test coverage.

**Files:**
- Test: `packages/modern-xlsx/__tests__/chart-api.test.ts` (NEW)

**Step 1: Write tests**

```typescript
import { describe, expect, it } from 'vitest';
import { Workbook, readBuffer } from '../src/index.js';
import type { WorksheetChartData } from '../src/types.js';

describe('chart API methods', () => {
  describe('addChartData', () => {
    it('adds a pre-built chart definition', async () => {
      const wb = new Workbook();
      const ws = wb.addSheet('S1');
      ws.cell('A1').value = 10;

      const chartData: WorksheetChartData = {
        chart: {
          chartType: 'bar',
          series: [{ idx: 0, order: 0, valRef: 'S1!$A$1' }],
        },
        anchor: {
          fromCol: 2, fromRow: 0, fromColOff: 0, fromRowOff: 0,
          toCol: 8, toRow: 10, toColOff: 0, toRowOff: 0,
        },
      };
      ws.addChartData(chartData);

      expect(ws.charts).toHaveLength(1);
      expect(ws.charts![0]!.chart.chartType).toBe('bar');

      // Verify it survives roundtrip
      const buf = await wb.toBuffer();
      const wb2 = await readBuffer(buf);
      expect(wb2.getSheet('S1')?.charts).toHaveLength(1);
    });

    it('initializes charts array if absent', () => {
      const wb = new Workbook();
      const ws = wb.addSheet('S1');
      expect(ws.charts).toBeUndefined();
      ws.addChartData({
        chart: { chartType: 'line', series: [{ idx: 0, order: 0, valRef: 'S1!$A$1' }] },
        anchor: { fromCol: 0, fromRow: 0, fromColOff: 0, fromRowOff: 0, toCol: 5, toRow: 5, toColOff: 0, toRowOff: 0 },
      });
      expect(ws.charts).toHaveLength(1);
    });
  });

  describe('removeChart', () => {
    it('removes chart by index and returns true', () => {
      const wb = new Workbook();
      const ws = wb.addSheet('S1');
      ws.addChart('bar', (b) => b.addSeries({ valRef: 'S1!$A$1' }).anchor({ fromCol: 0, fromRow: 0, toCol: 5, toRow: 5 }));
      ws.addChart('line', (b) => b.addSeries({ valRef: 'S1!$A$2' }).anchor({ fromCol: 6, fromRow: 0, toCol: 11, toRow: 5 }));

      expect(ws.charts).toHaveLength(2);
      const removed = ws.removeChart(0);
      expect(removed).toBe(true);
      expect(ws.charts).toHaveLength(1);
      expect(ws.charts![0]!.chart.chartType).toBe('line');
    });

    it('returns false for out-of-range index', () => {
      const wb = new Workbook();
      const ws = wb.addSheet('S1');
      expect(ws.removeChart(0)).toBe(false);
      expect(ws.removeChart(-1)).toBe(false);
      expect(ws.removeChart(99)).toBe(false);
    });

    it('returns false for negative index', () => {
      const wb = new Workbook();
      const ws = wb.addSheet('S1');
      ws.addChart('bar', (b) => b.addSeries({ valRef: 'S1!$A$1' }).anchor({ fromCol: 0, fromRow: 0, toCol: 5, toRow: 5 }));
      expect(ws.removeChart(-1)).toBe(false);
      expect(ws.charts).toHaveLength(1);
    });
  });

  describe('charts getter on empty sheet', () => {
    it('returns undefined when no charts added', () => {
      const wb = new Workbook();
      const ws = wb.addSheet('S1');
      expect(ws.charts).toBeUndefined();
    });
  });
});
```

**Step 2: Run tests**

Run: `pnpm -C packages/modern-xlsx test -- --reporter=verbose chart-api`
Expected: PASS

**Step 3: Commit**

```bash
git add packages/modern-xlsx/__tests__/chart-api.test.ts
git commit -m "test: add coverage for removeChart, addChartData, and charts getter"
```

---

### Task 6: Bubble Chart Tests

**Context:** Bubble charts exist in the type system and writer/parser but have zero test coverage. Key differentiator: `bubbleSizeRef` on series.

**Files:**
- Test: `packages/modern-xlsx/__tests__/chart-roundtrip.test.ts` (add bubble chart test)

**Step 1: Write bubble chart roundtrip test**

```typescript
it('roundtrips a bubble chart with bubbleSizeRef', async () => {
  const wb = new Workbook();
  const ws = wb.addSheet('Data');
  ws.cell('A1').value = 1;  // X
  ws.cell('A2').value = 2;
  ws.cell('B1').value = 10; // Y
  ws.cell('B2').value = 20;
  ws.cell('C1').value = 5;  // Bubble size
  ws.cell('C2').value = 15;

  ws.addChart('bubble', (b) => {
    b.title('Bubble Chart')
      .addSeries({
        xValRef: 'Data!$A$1:$A$2',
        valRef: 'Data!$B$1:$B$2',
        bubbleSizeRef: 'Data!$C$1:$C$2',
      })
      .anchor({ fromCol: 4, fromRow: 0, toCol: 12, toRow: 15 });
  });

  const buf = await wb.toBuffer();
  const wb2 = await readBuffer(buf);
  const ws2 = wb2.getSheet('Data') as Worksheet;
  const chart = ws2.charts![0]!;

  expect(chart.chart.chartType).toBe('bubble');
  expect(chart.chart.series[0]?.bubbleSizeRef).toBe('Data!$C$1:$C$2');
  expect(chart.chart.series[0]?.xValRef).toBe('Data!$A$1:$A$2');
  expect(chart.chart.series[0]?.valRef).toBe('Data!$B$1:$B$2');
});
```

**Step 2: Run and verify**

Run: `pnpm -C packages/modern-xlsx test -- --reporter=verbose chart-roundtrip`
Expected: PASS

**Step 3: Commit**

```bash
git add packages/modern-xlsx/__tests__/chart-roundtrip.test.ts
git commit -m "test: add bubble chart roundtrip with bubbleSizeRef"
```

---

### Task 7: Stock Chart Tests

**Context:** Stock charts exist in the type system but have zero tests. Stock charts (high-low-close or OHLC) use multiple series with specific ordering.

**Files:**
- Test: `packages/modern-xlsx/__tests__/chart-roundtrip.test.ts` (add stock chart test)

**Step 1: Write stock chart roundtrip test**

```typescript
it('roundtrips a stock chart with HLC data', async () => {
  const wb = new Workbook();
  const ws = wb.addSheet('Stock');
  // Date | High | Low | Close
  ws.cell('A1').value = 'Day 1';
  ws.cell('A2').value = 'Day 2';
  ws.cell('B1').value = 110; // High
  ws.cell('B2').value = 115;
  ws.cell('C1').value = 95;  // Low
  ws.cell('C2').value = 100;
  ws.cell('D1').value = 105; // Close
  ws.cell('D2').value = 108;

  ws.addChart('stock', (b) => {
    b.title('Stock Prices')
      .addSeries({ name: 'High', valRef: 'Stock!$B$1:$B$2', catRef: 'Stock!$A$1:$A$2' })
      .addSeries({ name: 'Low', valRef: 'Stock!$C$1:$C$2', catRef: 'Stock!$A$1:$A$2' })
      .addSeries({ name: 'Close', valRef: 'Stock!$D$1:$D$2', catRef: 'Stock!$A$1:$A$2' })
      .anchor({ fromCol: 5, fromRow: 0, toCol: 13, toRow: 15 });
  });

  const buf = await wb.toBuffer();
  const wb2 = await readBuffer(buf);
  const ws2 = wb2.getSheet('Stock') as Worksheet;
  const chart = ws2.charts![0]!;

  expect(chart.chart.chartType).toBe('stock');
  expect(chart.chart.series).toHaveLength(3);
  expect(chart.chart.series[0]?.name).toBe('High');
  expect(chart.chart.series[1]?.name).toBe('Low');
  expect(chart.chart.series[2]?.name).toBe('Close');
});
```

**Step 2: Run and verify**

Run: `pnpm -C packages/modern-xlsx test -- --reporter=verbose chart-roundtrip`
Expected: PASS

**Step 3: Commit**

```bash
git add packages/modern-xlsx/__tests__/chart-roundtrip.test.ts
git commit -m "test: add stock chart roundtrip with HLC series"
```

---

## Phase 3: WASM Boundary Validation

### Task 8: Chart Data Validation Layer

**Context:** Fields like `holeSize`, `styleId`, `explosion`, `lineWidth`, `rotX`, `rotY` are `number` in TypeScript but `u32`/`i32` in Rust. Passing negative values or floats causes opaque serde errors at the WASM boundary. This task adds a TypeScript validation layer that produces clear, actionable error messages.

**Files:**
- Create: `packages/modern-xlsx/src/validate-chart.ts` (NEW)
- Modify: `packages/modern-xlsx/src/workbook.ts:1050-1057` (call validator before chart push)
- Test: `packages/modern-xlsx/__tests__/chart-validation.test.ts` (NEW)

**Step 1: Write the failing test**

```typescript
import { describe, expect, it } from 'vitest';
import { Workbook } from '../src/index.js';

describe('chart validation', () => {
  it('rejects negative holeSize', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('S');
    expect(() => {
      ws.addChart('doughnut', (b) => {
        b.addSeries({ valRef: 'S!$A$1' })
          .holeSize(-5) // Invalid: must be 0-90
          .anchor({ fromCol: 0, fromRow: 0, toCol: 5, toRow: 5 });
      });
    }).toThrow(/holeSize must be 0–90/);
  });

  it('rejects float lineWidth', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('S');
    expect(() => {
      ws.addChart('line', (b) => {
        b.addSeries({ valRef: 'S!$A$1', lineWidth: 1.5 }) // must be integer
          .anchor({ fromCol: 0, fromRow: 0, toCol: 5, toRow: 5 });
      });
    }).toThrow(/lineWidth must be a non-negative integer/);
  });

  it('rejects rotX out of range', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('S');
    expect(() => {
      ws.addChart('bar', (b) => {
        b.addSeries({ valRef: 'S!$A$1' })
          .view3D({ rotX: 200 }) // must be -90 to 90
          .anchor({ fromCol: 0, fromRow: 0, toCol: 5, toRow: 5 });
      });
    }).toThrow(/rotX must be -90–90/);
  });

  it('accepts valid chart data', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('S');
    ws.cell('A1').value = 10;
    expect(() => {
      ws.addChart('doughnut', (b) => {
        b.addSeries({ valRef: 'S!$A$1' })
          .holeSize(50)
          .anchor({ fromCol: 0, fromRow: 0, toCol: 5, toRow: 5 });
      });
    }).not.toThrow();
  });
});
```

**Step 2: Run test to verify failure**

Run: `pnpm -C packages/modern-xlsx test -- --reporter=verbose chart-validation`
Expected: FAIL — no validation exists

**Step 3: Create validation module**

Create `packages/modern-xlsx/src/validate-chart.ts`:

```typescript
import type { WorksheetChartData } from './types.js';

function assertU32(name: string, value: number | null | undefined): void {
  if (value == null) return;
  if (!Number.isInteger(value) || value < 0 || value > 4294967295) {
    throw new RangeError(`${name} must be a non-negative integer (u32), got ${value}`);
  }
}

function assertI32Range(name: string, value: number | null | undefined, min: number, max: number): void {
  if (value == null) return;
  if (!Number.isInteger(value) || value < min || value > max) {
    throw new RangeError(`${name} must be ${min}\u2013${max}, got ${value}`);
  }
}

export function validateChartData(data: WorksheetChartData): void {
  const c = data.chart;

  // holeSize: u32, valid 0-90 for doughnut
  if (c.holeSize != null) {
    assertI32Range('holeSize', c.holeSize, 0, 90);
  }

  // styleId: u32
  assertU32('styleId', c.styleId);

  // View3D ranges (ECMA-376 spec)
  if (c.view3D) {
    assertI32Range('rotX', c.view3D.rotX, -90, 90);
    assertI32Range('rotY', c.view3D.rotY, 0, 360);
    assertI32Range('perspective', c.view3D.perspective, 0, 240);
  }

  // Series validation
  for (const ser of c.series) {
    assertU32('series.lineWidth', ser.lineWidth);
    assertU32('series.explosion', ser.explosion);
    if (ser.valRef === '') {
      throw new Error('series.valRef must not be empty');
    }
  }

  // Anchor validation
  const a = data.anchor;
  assertU32('anchor.fromCol', a.fromCol);
  assertU32('anchor.fromRow', a.fromRow);
  assertU32('anchor.toCol', a.toCol);
  assertU32('anchor.toRow', a.toRow);
}
```

**Step 4: Wire validation into Worksheet.addChart and addChartData**

In `workbook.ts:1050-1067`:

```typescript
import { validateChartData } from './validate-chart.js';

addChart(type: ChartType, configure: (builder: ChartBuilder) => void): void {
  const builder = new ChartBuilder(type);
  configure(builder);
  const chartData = builder.build();
  validateChartData(chartData); // NEW
  if (!this.data.worksheet.charts) {
    this.data.worksheet.charts = [];
  }
  this.data.worksheet.charts.push(chartData);
}

addChartData(chart: WorksheetChartData): void {
  validateChartData(chart); // NEW
  if (!this.data.worksheet.charts) {
    this.data.worksheet.charts = [];
  }
  this.data.worksheet.charts.push(chart);
}
```

**Step 5: Run all tests**

Run: `pnpm -C packages/modern-xlsx test`
Expected: ALL PASS

**Step 6: Commit**

```bash
git add packages/modern-xlsx/src/validate-chart.ts packages/modern-xlsx/src/workbook.ts packages/modern-xlsx/__tests__/chart-validation.test.ts
git commit -m "feat: add WASM boundary validation for chart data

Validates integer fields (holeSize, styleId, lineWidth, explosion) and
range constraints (rotX -90..90, rotY 0..360, holeSize 0..90) at the
TypeScript layer before data crosses to WASM. Produces clear error
messages instead of opaque serde deserialization failures."
```

---

## Phase 4: Rust Modernization (1.94.0 Edition 2024)

### Task 9: Convert panic! to Result in encryption_info.rs

**Context:** Three `panic!()` calls in `encryption_info.rs` (lines 403, 432, 476) can crash the WASM module if called with unexpected encryption variants. These should be `Result::Err` for proper error handling.

**Files:**
- Modify: `crates/modern-xlsx-core/src/ole2/encryption_info.rs:403,432,476`
- Test: existing encryption tests should continue passing

**Step 1: Replace panics with Result returns**

```rust
// Line 403: Change from:
_ => panic!("Expected Agile encryption info")
// To:
_ => return Err(ModernXlsxError::Encryption("expected Agile encryption info".into()))

// Line 432: Change from:
_ => panic!("Expected Agile encryption")
// To:
_ => return Err(ModernXlsxError::Encryption("expected Agile encryption".into()))

// Line 476: Change from:
_ => panic!("Expected Standard encryption")
// To:
_ => return Err(ModernXlsxError::Encryption("expected Standard encryption".into()))
```

Update function signatures to return `Result<T>` if not already.

**Step 2: Run tests**

Run: `cargo test -p modern-xlsx-core`
Expected: ALL PASS

**Step 3: Commit**

```bash
git add crates/modern-xlsx-core/src/ole2/encryption_info.rs
git commit -m "fix: replace panic! with Result::Err in encryption info parsing

Three panic! calls could crash WASM if called with unexpected encryption
variants. Now returns proper errors instead."
```

---

### Task 10: Optimize number_format.rs Allocation

**Context:** `clean_format_string()` (number_format.rs:90) allocates a `Vec<u8>` for every format string parse call. This can be optimized to use an iterator-based scan that avoids heap allocation for small format strings.

**Files:**
- Modify: `crates/modern-xlsx-core/src/number_format.rs:90-130`

**Step 1: Optimize with stack buffer**

Replace `Vec<u8>` with a `SmallVec<[u8; 64]>` or use a reusable buffer pattern. Most Excel format strings are under 64 bytes, so a stack allocation avoids heap:

```rust
use std::borrow::Cow;

fn clean_format_string(format: &str) -> SmallVec<[u8; 64]> {
    let bytes = format.as_bytes();
    let len = bytes.len();
    let mut result = SmallVec::with_capacity(len);
    // ... rest of logic unchanged, just use SmallVec instead of Vec
}
```

Alternatively, since `smallvec` is not a dependency, use a fixed-size array buffer with fallback:

```rust
fn clean_format_string(format: &str) -> Vec<u8> {
    let bytes = format.as_bytes();
    let len = bytes.len();
    // Pre-check: if no special characters, return lowercased bytes directly
    if !bytes.iter().any(|&b| b == b'[' || b == b'"' || b == b'\\' || b == b'_' || b == b'*') {
        return bytes.iter().map(|b| b.to_ascii_lowercase()).collect();
    }
    let mut result = Vec::with_capacity(len);
    // ... existing logic
    result
}
```

**Step 2: Run tests and benchmarks**

Run: `cargo test -p modern-xlsx-core`
Expected: ALL PASS

**Step 3: Commit**

```bash
git add crates/modern-xlsx-core/src/number_format.rs
git commit -m "perf: fast-path number format classification for simple strings

Skip allocation-heavy cleaning when format string has no special
characters (brackets, quotes, escapes). Most built-in formats take
the fast path."
```

---

### Task 11: Audit All Clone Calls

**Context:** 48 `.clone()` calls across the Rust codebase. Most are necessary for XML parsing state machines, but some may be replaceable with references or `Cow<str>`.

**Files:**
- Audit: all `.rs` files in `crates/modern-xlsx-core/src/`

**Step 1: Audit and categorize clones**

Search for all `.clone()` calls and categorize:
- **Necessary:** XML event data that goes out of scope (keep)
- **Removable:** Clones into structs that could take references (replace with `Cow<'_, str>` or `&str`)
- **Optimizable:** Clones that could use `std::mem::take()` instead (already used extensively at 64 call sites)

**Step 2: Replace unnecessary clones**

Focus on hot paths: `worksheet.rs` parser and `writer.rs` are the most performance-critical.

**Step 3: Run tests**

Run: `cargo test -p modern-xlsx-core`
Expected: ALL PASS

**Step 4: Commit**

```bash
git add crates/modern-xlsx-core/src/
git commit -m "perf: reduce unnecessary clones in XML parsing hot paths"
```

---

## Phase 5: TypeScript Modernization (6.0)

### Task 12: Add satisfies Operator and Readonly Modifiers

**Context:** TypeScript 6.0 provides `satisfies` for safer constant validation and `readonly` for stricter array contracts. These should be applied across the codebase.

**Files:**
- Modify: `packages/modern-xlsx/src/format-cell.ts:14` (BUILTIN_FORMATS)
- Modify: `packages/modern-xlsx/src/chart-styles.ts:2` (CHART_STYLE_PALETTES)
- Modify: `packages/modern-xlsx/src/types.ts` (add readonly to array properties)

**Step 1: Apply satisfies to constant objects**

In `format-cell.ts`:
```typescript
// Before:
const BUILTIN_FORMATS: Record<number, string> = { ... };
// After:
const BUILTIN_FORMATS = { ... } as const satisfies Record<number, string>;
```

In `chart-styles.ts`:
```typescript
// Before:
export const CHART_STYLE_PALETTES: ReadonlyMap<number, readonly string[]> = new Map([...]);
// After: (already good, but verify all entries are readonly)
```

**Step 2: Add readonly to types.ts array properties**

In `types.ts`, for data model interfaces that cross the WASM boundary, arrays should be `readonly` to prevent accidental mutation:

```typescript
// Example changes:
export interface WorksheetData {
  readonly cells: readonly CellData[];
  readonly rows: readonly RowData[];
  // ... other readonly arrays
}
```

Note: Only apply `readonly` where it doesn't break existing mutable APIs. The `Worksheet.cell()` setter pattern requires mutable access, so internal data remains mutable. Apply `readonly` to exported type interfaces that represent parsed data.

**Step 3: Run lint and tests**

Run: `pnpm -C packages/modern-xlsx lint && pnpm -C packages/modern-xlsx test`
Expected: ALL PASS

**Step 4: Commit**

```bash
git add packages/modern-xlsx/src/
git commit -m "refactor: apply TypeScript 6.0 satisfies and readonly modifiers"
```

---

### Task 13: Optimize format-cell.ts toLocaleString Caching

**Context:** `formatCell()` in `format-cell.ts` calls `toLocaleString()` for number formatting. In V8, ICU-based locale formatting is expensive (~10x slower than manual formatting). For sheets with 100K+ numeric cells, this becomes a bottleneck.

**Files:**
- Modify: `packages/modern-xlsx/src/format-cell.ts:541-545`

**Step 1: Cache Intl.NumberFormat instance**

```typescript
// Module-level cache
let cachedNumberFormatter: Intl.NumberFormat | null = null;

function getNumberFormatter(): Intl.NumberFormat {
  if (!cachedNumberFormatter) {
    cachedNumberFormatter = new Intl.NumberFormat(undefined, {
      useGrouping: true,
      maximumFractionDigits: 15,
    });
  }
  return cachedNumberFormatter;
}
```

Replace `value.toLocaleString()` calls with `getNumberFormatter().format(value)`.

**Step 2: Run benchmarks and tests**

Run: `pnpm -C packages/modern-xlsx test`
Expected: ALL PASS

**Step 3: Commit**

```bash
git add packages/modern-xlsx/src/format-cell.ts
git commit -m "perf: cache Intl.NumberFormat for 10x faster numeric cell formatting"
```

---

### Task 14: Optimize utils.ts String Building

**Context:** `sheetToCsv()`, `sheetToHtml()`, and `sheetToTxt()` use multiple `replaceAll()` calls per cell for escaping. For large sheets, this is O(n*m) where m is the number of escape patterns.

**Files:**
- Modify: `packages/modern-xlsx/src/utils.ts:315,481-485`

**Step 1: Consolidate HTML escaping to single pass**

```typescript
// Before (multiple replaceAll):
function escapeHtml(s: string): string {
  return s.replaceAll('&', '&amp;').replaceAll('<', '&lt;').replaceAll('>', '&gt;').replaceAll('"', '&quot;');
}

// After (single pass):
const HTML_ESCAPE_RE = /[&<>"]/g;
const HTML_ESCAPE_MAP: Record<string, string> = { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' };
function escapeHtml(s: string): string {
  return s.replace(HTML_ESCAPE_RE, (ch) => HTML_ESCAPE_MAP[ch]!);
}
```

**Step 2: Consolidate CSV escaping**

```typescript
// Single regex test instead of two replaceAll:
function csvQuote(s: string): string {
  if (/[",\n\r]/.test(s)) {
    return `"${s.replaceAll('"', '""')}"`;
  }
  return s;
}
```

**Step 3: Run tests**

Run: `pnpm -C packages/modern-xlsx test`
Expected: ALL PASS

**Step 4: Commit**

```bash
git add packages/modern-xlsx/src/utils.ts
git commit -m "perf: single-pass HTML/CSV escaping for large sheet exports"
```

---

### Task 15: Fix worker-api.ts Memory Leak

**Context:** `worker-api.ts` has a `pending` Map that stores unresolved Promises. If the worker crashes or is terminated while requests are pending, these promises leak (never resolved/rejected).

**Files:**
- Modify: `packages/modern-xlsx/src/worker-api.ts`

**Step 1: Add cleanup on worker error/close**

```typescript
// In createXlsxWorker():
worker.addEventListener('error', () => {
  // Reject all pending promises
  for (const [id, { reject }] of pending) {
    reject(new Error('Worker terminated unexpectedly'));
  }
  pending.clear();
});
```

**Step 2: Run tests**

Run: `pnpm -C packages/modern-xlsx test`
Expected: ALL PASS

**Step 3: Commit**

```bash
git add packages/modern-xlsx/src/worker-api.ts
git commit -m "fix: reject pending promises on worker crash to prevent memory leaks"
```

---

## Phase 6: Golden File Tests & Documentation

### Task 16: Expand Golden File Test Suite

**Context:** Only 2 golden file tests exist (simple_roundtrip.json, multi_sheet_roundtrip.json). Each major feature should have a golden file test that verifies the actual ZIP output matches a known-good reference.

**Files:**
- Create: `crates/modern-xlsx-core/tests/golden/` (new golden files)
- Modify: `crates/modern-xlsx-core/tests/golden_tests.rs` (add new golden tests)

**Step 1: Create golden tests for each major feature**

Add golden tests for:
1. **Styled workbook** — fonts, fills, borders, number formats
2. **Charts** — bar chart with axis titles
3. **Conditional formatting** — color scales, data bars, icon sets
4. **Data validation** — dropdown lists, numeric ranges
5. **Comments** — cell comments with authors
6. **Frozen panes** — split pane configurations
7. **Hyperlinks** — URL and internal references
8. **Tables** — table definitions with styles
9. **Encryption** — encrypted file structure (verify salt randomization doesn't affect comparison)

Each golden test:
1. Creates a workbook with the feature
2. Writes to buffer
3. Parses the JSON representation
4. Compares against a committed `.json` golden file

Use `GOLDEN_UPDATE=1 cargo test` env var pattern to regenerate golden files when format changes.

**Step 2: Generate initial golden files**

Run each test with `GOLDEN_UPDATE=1` to create the reference files. Commit them.

**Step 3: Verify all golden tests pass**

Run: `cargo test -p modern-xlsx-core golden`
Expected: ALL PASS

**Step 4: Commit**

```bash
git add crates/modern-xlsx-core/tests/golden/ crates/modern-xlsx-core/tests/golden_tests.rs
git commit -m "test: expand golden file suite to 10 tests covering all major features"
```

---

### Task 17: Update All Version References

**Context:** Multiple files have stale version numbers: README test counts, CLAUDE.md test counts, example configs referencing v0.5.0, WASM size claims.

**Files:**
- Modify: `README.md` (root) — test counts, performance numbers
- Modify: `CLAUDE.md` — test counts
- Modify: `packages/modern-xlsx/README.md` — file sizes, CDN URLs
- Modify: `examples/cloudflare-worker/package.json` — version reference
- Modify: `examples/deno-deploy/deno.json` — version reference

**Step 1: Update test counts**

Run the actual test counts:
```bash
cargo test -p modern-xlsx-core 2>&1 | tail -1  # Get Rust test count
pnpm -C packages/modern-xlsx test 2>&1 | grep "Tests"  # Get TS test count
```

Update all files with accurate numbers.

**Step 2: Update file sizes**

```bash
ls -la packages/modern-xlsx/dist/  # Get actual sizes
```

Update README claims to match actual sizes.

**Step 3: Update example configs**

Change `"modern-xlsx": "^0.5.0"` to `"modern-xlsx": "^0.8.1"` in example package.json files.

**Step 4: Commit**

```bash
git add README.md CLAUDE.md packages/modern-xlsx/README.md examples/
git commit -m "docs: update all version references, test counts, and file sizes to v0.8.1"
```

---

### Task 18: Wiki Accuracy Review

**Context:** Wiki pages were updated rapidly during feature sprints. A systematic pass is needed to verify every code snippet runs correctly and all version/feature claims are accurate.

**Files:**
- All wiki `.md` files in `C:\Users\alber\AppData\Local\Temp\modern-xlsx-wiki\`

**Step 1: Review each wiki page**

For each of the 15 wiki pages:
1. Verify all code snippets compile/run with current API
2. Check version references match v0.8.1
3. Verify feature claims are accurate (especially Feature-Comparison.md)
4. Fix any stale examples

**Step 2: Feature-Comparison honesty review**

Current charts claim: `:star: 10 types, ChartBuilder API, trendlines, error bars, 3D, combo`

After this plan's fixes (combo chart parsing, oneCellAnchor, axis fonts), this claim becomes more accurate. However, review whether `:star:` (superior) or `:white_check_mark:` (fully supported) is the honest assessment given remaining gaps (8/48 style palettes, limited combo chart support).

**Step 3: Update CDN URLs in all wiki pages**

All `@0.8.0` references should be `@0.8.1` (or whatever the latest version is after this plan).

**Step 4: Commit and push wiki**

```bash
cd /tmp/modern-xlsx-wiki
git add -A && git commit -m "docs: wiki accuracy review — fix stale examples, update versions"
git push
```

---

### Task 19: Update Memory File

**Context:** The auto-memory file should reflect the final state after all changes.

**Files:**
- Modify: `C:\Users\alber\.claude\projects\C--Users-alber-Desktop-Projects-modern-xlsx\memory\MEMORY.md`

**Step 1: Update version, test counts, and recent release info**

After all tasks complete, update:
- Version number
- Test counts (Rust + TypeScript)
- Recent releases list
- Any new architectural patterns (drawing XML merge, validation layer)
- Wiki page count if changed

**Step 2: Save**

---

## Phase 7: Final Build, Test & Release

### Task 20: Rebuild WASM and Full Test Suite

**Step 1: Rebuild WASM**

```bash
cd crates/modern-xlsx-wasm && wasm-pack build --target web --release --out-dir ../../packages/modern-xlsx/wasm --no-opt
```

**Step 2: Build TypeScript**

```bash
pnpm -C packages/modern-xlsx build
```

**Step 3: Run full test suite**

```bash
cargo test -p modern-xlsx-core
pnpm -C packages/modern-xlsx test
pnpm -C packages/modern-xlsx lint
pnpm -C packages/modern-xlsx typecheck
cargo clippy -p modern-xlsx-core -- -D warnings
```

Expected: ALL PASS, ZERO WARNINGS

**Step 4: Version bump**

Bump to next patch version (v0.9.0 if the changes are significant enough for a minor, or v0.8.2 for patch).

**Step 5: Commit, tag, push**

```bash
git add -A
git commit -m "chore: v0.9.0 — comprehensive audit, bug fixes, and modernization

Critical fixes:
- Barcode + chart drawing XML merge (prevented data loss)
- Restored ChartAxis.fontSize with proper <c:txPr> writer/parser
- oneCellAnchor chart parsing
- Combo chart roundtrip (secondary_chart parsing)

New:
- WASM boundary validation for chart data
- 10 golden file tests (up from 2)
- removeChart/addChartData/bubble/stock chart tests

Modernization:
- Rust: eliminated panic!, optimized allocations, reduced clones
- TypeScript: satisfies, readonly, cached Intl.NumberFormat, single-pass escaping
- Fixed worker-api memory leak on crash

Docs:
- Updated all version references and test counts
- Wiki accuracy review"

git tag v0.9.0
git push origin master --tags
```

---

## Summary

| Phase | Tasks | Priority | Impact |
|-------|-------|----------|--------|
| **1: Critical Bugs** | 4 tasks | P0 | Fix live data loss bug, restore lost features, improve roundtrip fidelity |
| **2: Test Coverage** | 3 tasks | P1 | Close coverage gaps for public API, bubble/stock charts |
| **3: WASM Validation** | 1 task | P1 | Prevent opaque serde errors with clear validation messages |
| **4: Rust Modernization** | 3 tasks | P2 | Eliminate panics, optimize allocations, reduce clones |
| **5: TS Modernization** | 4 tasks | P2 | TypeScript 6.0 idioms, performance optimizations |
| **6: Docs & Golden Tests** | 4 tasks | P2 | Documentation accuracy, expanded golden test suite |
| **7: Release** | 1 task | P3 | Build, test, version bump, publish |

**Total: 20 tasks across 7 phases**

**Estimated new test count after completion:**
- Rust: ~350 tests (from 337)
- TypeScript: ~1200 tests (from 1171)
- Golden files: 10 (from 2)
