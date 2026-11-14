import { assert, describe, expect, it } from 'vitest';
import type { Worksheet } from '../src/index.js';
import { encodeQR, readBuffer, renderBarcodePNG, Workbook } from '../src/index.js';
import { readZipText } from './helpers/zip.js';

const VML_REL_TYPE =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing';
const VML_CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.vmlDrawing';

/** The sheet named `name`, narrowed with `assert.isDefined`. */
function sheet(wb: Workbook, name: string): Worksheet {
  const ws = wb.getSheet(name);
  assert.isDefined(ws);
  return ws;
}

/** A sheet with data in A1:B5 and a note on D1. */
function notedWorkbook(): Workbook {
  const wb = new Workbook();
  const ws = wb.addSheet('Sheet1');
  ws.cell('A1').value = 'Name';
  ws.cell('B1').value = 'Value';
  for (const [i, name] of ['alpha', 'beta', 'gamma', 'delta'].entries()) {
    ws.cell(`A${i + 2}`).value = name;
    ws.cell(`B${i + 2}`).value = (i + 1) * 10;
  }
  ws.addComment('D1', 'Ann', 'note');
  return wb;
}

/** {@link notedWorkbook} with a bar chart of B2:B5 (case (b)). */
function withChart(): Workbook {
  const wb = notedWorkbook();
  sheet(wb, 'Sheet1').addChart('bar', (b) => {
    b.title('Values')
      .addSeries({ name: 'Value', catRef: 'Sheet1!$A$2:$A$5', valRef: 'Sheet1!$B$2:$B$5' })
      .anchor({ col: 4, row: 6 }, { col: 12, row: 20 });
  });
  return wb;
}

/** {@link notedWorkbook} with a picture over E7:I15 (case (c)). */
function withImage(): Workbook {
  const wb = notedWorkbook();
  const png = renderBarcodePNG(encodeQR('NV-01'), { moduleSize: 2 });
  wb.addImage('Sheet1', { fromCol: 4, fromRow: 6, toCol: 8, toRow: 14 }, png);
  return wb;
}

/** The `(Row, Column)` of each note shape in a VML drawing, zero-based, in order. */
function noteCells(vml: string): [number, number][] {
  return [
    ...vml.matchAll(
      /<x:ClientData ObjectType="Note">[\s\S]*?<x:Row>(\d+)<\/x:Row><x:Column>(\d+)<\/x:Column>/g,
    ),
  ].map((m) => [Number(m[1]), Number(m[2])]);
}

/**
 * Asserts sheet1 of `bytes` names its VML drawing by `<legacyDrawing>` after `<drawing>`, through
 * a vmlDrawing relationship, with the `vml` content type, and returns the drawing's text.
 */
function legacyVml(bytes: Uint8Array): string {
  const parts = readZipText(bytes);
  const sheetXml = parts['xl/worksheets/sheet1.xml'];
  const rels = parts['xl/worksheets/_rels/sheet1.xml.rels'];
  const types = parts['[Content_Types].xml'];
  assert.isDefined(sheetXml);
  assert.isDefined(rels);
  assert.isDefined(types);

  const id = /<legacyDrawing r:id="([^"]+)"\/>/.exec(sheetXml)?.[1];
  assert.isDefined(id, sheetXml);
  const drawingAt = sheetXml.indexOf('<drawing ');
  expect(drawingAt).toBeGreaterThan(-1);
  expect(sheetXml.indexOf('<legacyDrawing ')).toBeGreaterThan(drawingAt);

  const rel = [...rels.matchAll(/<Relationship\b[^>]*>/g)]
    .map((m) => m[0])
    .find((tag) => tag.includes(`Id="${id}"`));
  assert.isDefined(rel, rels);
  expect(rel).toContain(`Type="${VML_REL_TYPE}"`);
  const target = /\bTarget="\.\.\/drawings\/([^"]+\.vml)"/.exec(rel)?.[1];
  assert.isDefined(target, rel);
  expect(types).toContain(`<Default Extension="vml" ContentType="${VML_CONTENT_TYPE}"/>`);

  const vml = parts[`xl/drawings/${target}`];
  assert.isDefined(vml);
  return vml;
}

describe('a note beside a chart or a picture gets a VML note drawing', () => {
  for (const [label, build] of [
    ['a chart', withChart],
    ['a picture', withImage],
  ] as const) {
    it(`names a generated note drawing for a note beside ${label}`, async () => {
      const bytes = await build().toBuffer();
      const vml = legacyVml(bytes);
      expect(vml).toContain('<o:idmap v:ext="edit" data="1"/>');
      expect(vml).toContain('id="_x0000_s1025"');
      expect(vml).toContain('visibility:hidden');
      expect(noteCells(vml)).toEqual([[0, 3]]);

      const read = sheet(await readBuffer(bytes), 'Sheet1');
      expect(read.comments).toEqual([{ cellRef: 'D1', author: 'Ann', text: 'note' }]);
    });
  }

  it('keeps the note drawing through a round trip and follows an edit of the notes', async () => {
    const first = await withChart().toBuffer();
    const firstVml = legacyVml(first);

    // Read and written unchanged: the drawing is kept byte for byte.
    const unchanged = await (await readBuffer(first)).toBuffer();
    expect(legacyVml(unchanged)).toBe(firstVml);

    // A note added and the first one removed: the drawing follows the comments.
    const wb = await readBuffer(first);
    const ws = sheet(wb, 'Sheet1');
    ws.addComment('F10', 'Bob', 'added');
    expect(ws.removeComment('D1')).toBe(true);
    const edited = await wb.toBuffer();
    const vml = legacyVml(edited);
    expect(noteCells(vml)).toEqual([[9, 5]]);
    expect(vml).toContain('id="_x0000_s1026"');
    expect(sheet(await readBuffer(edited), 'Sheet1').comments).toEqual([
      { cellRef: 'F10', author: 'Bob', text: 'added' },
    ]);
  });
});
