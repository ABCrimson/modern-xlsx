import { describe, expect, it } from 'vitest';
import { encodeQR, readBuffer, renderBarcodePNG, Workbook } from '../src/index.js';
import { readZipText } from './helpers/zip.js';

const REL_DRAWING = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing';

function png(): Uint8Array {
  return renderBarcodePNG(encodeQR('IMAGE-DRAWING'), { moduleSize: 2 });
}

const anchor = { fromCol: 0, fromRow: 5, toCol: 3, toRow: 12 };

/** Returns the `r:id` of every `<drawing>` element in a sheet part. */
function drawingIds(sheetXml: string): string[] {
  return [...sheetXml.matchAll(/<drawing\b[^>]*\br:id="([^"]+)"/g)].map((m) => m[1] as string);
}

/** Returns the ids of the drawing relationships in a rels part. */
function drawingRelIds(relsXml: string): string[] {
  return [...relsXml.matchAll(/<Relationship\b[^>]*>/g)]
    .map((m) => m[0])
    .filter((tag) => tag.includes(`Type="${REL_DRAWING}"`))
    .map((tag) => /\bId="([^"]+)"/.exec(tag)?.[1] as string);
}

/** Resolves the drawing part that relationship `id` of sheet1 points to. */
function resolveDrawingPart(parts: Record<string, string>, id: string): string | undefined {
  const rels = parts['xl/worksheets/_rels/sheet1.xml.rels'] as string;
  const tag = [...rels.matchAll(/<Relationship\b[^>]*>/g)]
    .map((m) => m[0])
    .find((t) => t.includes(`Type="${REL_DRAWING}"`) && t.includes(`Id="${id}"`));
  const target = /\bTarget="([^"]+)"/.exec(tag ?? '')?.[1];
  if (target === undefined) return undefined;
  return parts[new URL(target, 'http://x/xl/worksheets/').pathname.slice(1)];
}

describe('image-only sheet drawing reference', () => {
  it('writes <drawing r:id> for a chart-free sheet with an image', async () => {
    const wb = new Workbook();
    wb.addSheet('Pics').cell('A1').value = 'x';
    wb.addImage('Pics', anchor, png());

    const parts = readZipText(await wb.toBuffer());
    const ids = drawingIds(parts['xl/worksheets/sheet1.xml'] as string);
    expect(ids).toHaveLength(1);
    expect(drawingRelIds(parts['xl/worksheets/_rels/sheet1.xml.rels'] as string)).toEqual(ids);
    expect(parts['xl/drawings/drawing1.xml']).toContain('<xdr:pic>');
  });

  it('never writes a dangling <drawing> after a read and rewrite round trip', async () => {
    const wb = new Workbook();
    wb.addSheet('Pics').cell('A1').value = 'x';
    wb.addImage('Pics', anchor, png());

    const parts = readZipText(await (await readBuffer(await wb.toBuffer())).toBuffer());
    // Every <drawing r:id> the writer emits must resolve to a part that is in
    // the package.
    const ids = drawingIds(parts['xl/worksheets/sheet1.xml'] as string);
    expect(ids.length).toBeLessThanOrEqual(1);
    for (const id of ids) {
      expect(resolveDrawingPart(parts, id)).toBeDefined();
    }
  });

  // The reader preserves every drawing that is not chart-only.
  it('round trip keeps exactly one <drawing>', async () => {
    const wb = new Workbook();
    wb.addSheet('Pics').cell('A1').value = 'x';
    wb.addImage('Pics', anchor, png());

    const parts = readZipText(await (await readBuffer(await wb.toBuffer())).toBuffer());
    const ids = drawingIds(parts['xl/worksheets/sheet1.xml'] as string);
    expect(ids).toHaveLength(1);
    expect(resolveDrawingPart(parts, ids[0] as string)).toContain('<xdr:pic>');
  });

  it('writes exactly one <drawing> for a sheet with a chart and an image', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Mixed');
    ws.cell('A1').value = '1';
    ws.addChart('bar', (b) => {
      b.addSeries({ valRef: 'Mixed!$A$1:$A$1' }).anchor({ col: 4, row: 0 }, { col: 12, row: 15 });
    });
    wb.addImage('Mixed', anchor, png());

    const parts = readZipText(await wb.toBuffer());
    expect(drawingIds(parts['xl/worksheets/sheet1.xml'] as string)).toHaveLength(1);
    expect(parts['xl/drawings/drawing1.xml']).toContain('<xdr:pic>');
    expect(parts['xl/drawings/drawing1.xml']).toContain('<xdr:graphicFrame');
  });
});
