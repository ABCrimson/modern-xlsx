import { assert, describe, expect, it } from 'vitest';
import type { Worksheet } from '../src/index.js';
import { readBuffer, Workbook } from '../src/index.js';
import { buildZip, readZipText } from './helpers/zip.js';

/** Element `index` of `items`, narrowed with `assert.isDefined` (noUncheckedIndexedAccess). */
function at<T>(items: readonly T[] | undefined, index: number): T {
  const item = items?.[index];
  assert.isDefined(item);
  return item;
}

/** The sheet named `name`, narrowed with `assert.isDefined`. */
function sheet(wb: Workbook, name: string): Worksheet {
  const ws = wb.getSheet(name);
  assert.isDefined(ws);
  return ws;
}

/** One sheet with a table over A1:B5, a note on D1 and a bar chart of the table's values. */
function mixedWorkbook(): Workbook {
  const wb = new Workbook();
  const ws = wb.addSheet('Sheet1');
  ws.cell('A1').value = 'Name';
  ws.cell('B1').value = 'Value';
  for (const [i, name] of ['alpha', 'beta', 'gamma', 'delta'].entries()) {
    ws.cell(`A${i + 2}`).value = name;
    ws.cell(`B${i + 2}`).value = (i + 1) * 10;
  }
  ws.addTable({
    id: 1,
    displayName: 'Table1',
    ref: 'A1:B5',
    headerRowCount: 1,
    totalsRowCount: 0,
    totalsRowShown: true,
    columns: [
      { id: 1, name: 'Name' },
      { id: 2, name: 'Value' },
    ],
  });
  ws.addComment('D1', 'Ann', 'note');
  ws.addChart('bar', (b) => {
    b.title('Values')
      .addSeries({ name: 'Value', catRef: 'Sheet1!$A$2:$A$5', valRef: 'Sheet1!$B$2:$B$5' })
      .anchor({ col: 4, row: 0 }, { col: 12, row: 15 });
  });
  return wb;
}

/** Asserts the table, the note, the chart and the data of {@link mixedWorkbook} are all there. */
function expectAllThree(ws: Worksheet): void {
  expect(ws.tables).toHaveLength(1);
  const table = at(ws.tables, 0);
  expect(table.displayName).toBe('Table1');
  expect(table.ref).toBe('A1:B5');
  expect(table.columns.map((c) => c.name)).toEqual(['Name', 'Value']);

  expect(ws.comments).toHaveLength(1);
  expect(at(ws.comments, 0)).toMatchObject({ cellRef: 'D1', author: 'Ann', text: 'note' });

  expect(ws.charts).toHaveLength(1);
  const chart = at(ws.charts, 0);
  expect(chart.chart.chartType).toBe('bar');
  expect(chart.chart.title?.text).toBe('Values');
  expect(at(chart.chart.series, 0).catRef).toBe('Sheet1!$A$2:$A$5');
  expect(at(chart.chart.series, 0).valRef).toBe('Sheet1!$B$2:$B$5');
  expect(chart.anchor.fromCol).toBe(4);

  expect(ws.cell('A5').value).toBe('delta');
  expect(ws.cell('B5').value).toBe(40);
}

const encoder = new TextEncoder();

/** The VML drawing Excel saves for one note on D1 (shape `_x0000_s1025`, anchored by `x:ClientData`). */
const NOTE_VML =
  '<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel">' +
  '<o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="1"/></o:shapelayout>' +
  '<v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe"><v:stroke joinstyle="miter"/><v:path gradientshapeok="t" o:connecttype="rect"/></v:shapetype>' +
  '<v:shape id="_x0000_s1025" type="#_x0000_t202" style="position:absolute;margin-left:199.5pt;margin-top:1pt;width:104pt;height:61.5pt;z-index:1;visibility:hidden" fillcolor="infoBackground [80]" o:insetmode="auto">' +
  '<v:textbox style="mso-direction-alt:auto"><div style="text-align:left"></div></v:textbox>' +
  '<x:ClientData ObjectType="Note"><x:MoveWithCells/><x:SizeWithCells/><x:Anchor>4, 10, 0, 1, 6, 20, 4, 6</x:Anchor><x:AutoFill>False</x:AutoFill><x:Row>0</x:Row><x:Column>3</x:Column></x:ClientData>' +
  '</v:shape></xml>';

const VML_REL_TYPE =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing';

/**
 * `bytes` with sheet1 given the note's VML drawing the way Excel saves it: the part, a
 * `vmlDrawing` relationship `rIdVml`, `<legacyDrawing r:id="rIdVml"/>` after `<drawing>` and
 * the `vml` content type.
 */
function withExcelNoteVml(bytes: Uint8Array): Uint8Array {
  const parts = readZipText(bytes);
  const relsPath = 'xl/worksheets/_rels/sheet1.xml.rels';
  const sheetPath = 'xl/worksheets/sheet1.xml';
  const rels = parts[relsPath];
  const sheetXml = parts[sheetPath];
  const types = parts['[Content_Types].xml'];
  assert.isDefined(rels);
  assert.isDefined(sheetXml);
  assert.isDefined(types);
  expect(sheetXml).toMatch(/<drawing r:id="[^"]+"\/>/);
  parts[relsPath] = rels.replace(
    '</Relationships>',
    `<Relationship Id="rIdVml" Type="${VML_REL_TYPE}" Target="../drawings/vmlDrawing1.vml"/></Relationships>`,
  );
  parts[sheetPath] = sheetXml.replace(
    /<drawing r:id="[^"]+"\/>/,
    '$&<legacyDrawing r:id="rIdVml"/>',
  );
  parts['[Content_Types].xml'] = types.includes('Extension="vml"')
    ? types
    : types.replace(
        '<Default Extension="xml"',
        '<Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/><Default Extension="xml"',
      );
  parts['xl/drawings/vmlDrawing1.vml'] = NOTE_VML;
  return buildZip(Object.entries(parts).map(([name, text]) => [name, encoder.encode(text)]));
}

describe('a sheet with a table, a note and a chart', () => {
  it('keeps all three through write, read, write and read', async () => {
    const first = await readBuffer(await mixedWorkbook().toBuffer());
    expectAllThree(sheet(first, 'Sheet1'));
    const second = await readBuffer(await first.toBuffer());
    expectAllThree(sheet(second, 'Sheet1'));
  });

  it("names the note's VML drawing in <legacyDrawing> after a round trip, as Excel saved it", async () => {
    const excelSaved = withExcelNoteVml(await mixedWorkbook().toBuffer());
    const wb = await readBuffer(excelSaved);
    expectAllThree(sheet(wb, 'Sheet1'));

    const out = readZipText(await wb.toBuffer());
    const sheetXml = out['xl/worksheets/sheet1.xml'];
    const rels = out['xl/worksheets/_rels/sheet1.xml.rels'];
    assert.isDefined(sheetXml);
    assert.isDefined(rels);
    expect(out['xl/drawings/vmlDrawing1.vml']).toBe(NOTE_VML);

    const vmlId =
      /<Relationship Id="([^"]+)" Type="[^"]*\/vmlDrawing" Target="\.\.\/drawings\/vmlDrawing1\.vml"\/>/.exec(
        rels,
      )?.[1];
    assert.isDefined(vmlId);
    expect(sheetXml).toContain(`<legacyDrawing r:id="${vmlId}"/>`);
    // CT_Worksheet order: <drawing>, then <legacyDrawing>, then <tableParts>.
    const drawingAt = sheetXml.indexOf('<drawing ');
    const legacyAt = sheetXml.indexOf('<legacyDrawing ');
    const tablesAt = sheetXml.indexOf('<tableParts');
    expect(drawingAt).toBeGreaterThan(-1);
    expect(legacyAt).toBeGreaterThan(drawingAt);
    expect(tablesAt).toBeGreaterThan(legacyAt);

    expectAllThree(sheet(await readBuffer(await wb.toBuffer()), 'Sheet1'));
  });
});
