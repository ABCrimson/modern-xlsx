import { describe, expect, it } from 'vitest';
import type { Worksheet } from '../src/index.js';
import { encodeQR, readBuffer, renderBarcodePNG, Workbook } from '../src/index.js';
import { readZipText } from './helpers/zip.js';

describe('barcode + chart on same worksheet (drawing collision)', () => {
  it('preserves both image and chart after roundtrip', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Mixed');

    ws.cell('A1').value = 'Q1';
    ws.cell('B1').value = '100';
    ws.cell('A2').value = 'Q2';
    ws.cell('B2').value = '200';

    // Add a chart
    ws.addChart('bar', (b) => {
      b.title('Sales')
        .addSeries({
          valRef: 'Mixed!$B$1:$B$2',
          catRef: 'Mixed!$A$1:$A$2',
        })
        .anchor({ col: 4, row: 0 }, { col: 12, row: 15 });
    });

    // Add a barcode image
    const matrix = encodeQR('TEST-DATA');
    const png = renderBarcodePNG(matrix, { moduleSize: 4 });
    wb.addImage(
      'Mixed',
      {
        fromCol: 0,
        fromRow: 5,
        toCol: 3,
        toRow: 12,
      },
      png,
    );

    // Before roundtrip, verify both are set up in the workbook data
    const preJson = wb.toJSON();
    const prePe = preJson.preservedEntries ?? {};
    expect(prePe['xl/drawings/drawing1.xml']).toBeDefined();
    expect(prePe['xl/drawings/_rels/drawing1.xml.rels']).toBeDefined();

    // Roundtrip: write then read back
    const buf = await wb.toBuffer();
    const wb2 = await readBuffer(buf);
    const ws2 = wb2.getSheet('Mixed') as Worksheet;

    expect(ws2).toBeDefined();

    // Chart must survive roundtrip
    expect(ws2.charts).toHaveLength(1);
    expect(ws2.charts[0]?.chart.title?.text).toBe('Sales');
    expect(ws2.charts[0]?.chart.series).toHaveLength(1);
    expect(ws2.charts[0]?.chart.series[0]?.valRef).toBe('Mixed!$B$1:$B$2');

    // Image must survive — the image media bytes should be in preserved entries.
    const json = wb2.toJSON();
    const pe = json.preservedEntries ?? {};

    // The image bytes should survive in xl/media/
    const mediaKeys = Object.keys(pe).filter((k) => k.startsWith('xl/media/'));
    expect(mediaKeys.length).toBeGreaterThanOrEqual(1);

    // A drawing that holds a picture beside its chart stays in preservedEntries:
    // the reader consumes only chart-only drawings.
    expect(pe['xl/drawings/drawing1.xml']).toBeDefined();
    expect(pe['xl/drawings/_rels/drawing1.xml.rels']).toBeDefined();
  });

  it('preserves chart when no image is present (no regression)', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('ChartOnly');

    ws.cell('A1').value = 'X';
    ws.cell('B1').value = '50';

    ws.addChart('line', (b) => {
      b.title('Trend')
        .addSeries({ valRef: 'ChartOnly!$B$1:$B$1' })
        .anchor({ col: 3, row: 0 }, { col: 10, row: 10 });
    });

    const buf = await wb.toBuffer();
    const wb2 = await readBuffer(buf);
    const ws2 = wb2.getSheet('ChartOnly') as Worksheet;

    expect(ws2.charts).toHaveLength(1);
    expect(ws2.charts[0]?.chart.title?.text).toBe('Trend');
  });

  it('preserves image when no chart is present (no regression)', async () => {
    const wb = new Workbook();
    wb.addSheet('ImageOnly');

    const matrix = encodeQR('IMG-ONLY');
    const png = renderBarcodePNG(matrix, { moduleSize: 2 });
    wb.addImage(
      'ImageOnly',
      {
        fromCol: 0,
        fromRow: 0,
        toCol: 4,
        toRow: 6,
      },
      png,
    );

    const buf = await wb.toBuffer();
    const wb2 = await readBuffer(buf);
    const json = wb2.toJSON();
    const pe = json.preservedEntries ?? {};

    // Image media should survive in preserved entries
    const mediaKeys = Object.keys(pe).filter((k) => k.startsWith('xl/media/'));
    expect(mediaKeys.length).toBeGreaterThanOrEqual(1);
  });

  it('handles multiple images + chart on the same sheet', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Multi');

    ws.cell('A1').value = 'Data';
    ws.cell('B1').value = '42';

    ws.addChart('pie', (b) => {
      b.title('Pie')
        .addSeries({ valRef: 'Multi!$B$1:$B$1' })
        .anchor({ col: 5, row: 0 }, { col: 13, row: 12 });
    });

    // Two images
    const matrix1 = encodeQR('IMG-1');
    const png1 = renderBarcodePNG(matrix1, { moduleSize: 2 });
    wb.addImage('Multi', { fromCol: 0, fromRow: 0, toCol: 2, toRow: 3 }, png1);

    const matrix2 = encodeQR('IMG-2');
    const png2 = renderBarcodePNG(matrix2, { moduleSize: 2 });
    wb.addImage('Multi', { fromCol: 0, fromRow: 4, toCol: 2, toRow: 7 }, png2);

    const buf = await wb.toBuffer();

    // Each picture keeps its own image: the chart takes rId1, the pictures rId2
    // and rId3, and each id resolves to a different media part.
    const zip = readZipText(buf);
    const drawingXml = zip['xl/drawings/drawing1.xml'] ?? '';
    const embeds = [...drawingXml.matchAll(/r:embed="([^"]+)"/g)].map((m) => m[1]);
    expect(embeds).toEqual(['rId2', 'rId3']);
    const drawingRels = zip['xl/drawings/_rels/drawing1.xml.rels'] ?? '';
    const targets = embeds.map(
      (id) => new RegExp(`Id="${id}"[^>]*Target="([^"]+)"`).exec(drawingRels)?.[1],
    );
    expect(targets).toEqual(['../media/image1.png', '../media/image2.png']);
    for (const target of targets) {
      expect(zip[`xl/${target?.slice('../'.length)}`]).toBeDefined();
    }

    const wb2 = await readBuffer(buf);
    const ws2 = wb2.getSheet('Multi') as Worksheet;

    // Chart should survive (exactly 1, not duplicated)
    expect(ws2.charts).toHaveLength(1);
    expect(ws2.charts[0]?.chart.title?.text).toBe('Pie');

    // Both image media files should survive
    const json = wb2.toJSON();
    const pe = json.preservedEntries ?? {};
    const mediaKeys = Object.keys(pe).filter((k) => k.startsWith('xl/media/'));
    expect(mediaKeys.length).toBeGreaterThanOrEqual(2);
  });

  it('chart drawing XML includes merged image anchors (ZIP inspection)', async () => {
    // Opens the written ZIP and checks the merged drawing XML and its relationships.
    const wb = new Workbook();
    const ws = wb.addSheet('Inspect');

    ws.cell('A1').value = 'Val';
    ws.cell('B1').value = '10';

    ws.addChart('bar', (b) => {
      b.title('Test')
        .addSeries({ valRef: 'Inspect!$B$1:$B$1' })
        .anchor({ col: 4, row: 0 }, { col: 10, row: 10 });
    });

    const matrix = encodeQR('INSPECT');
    const png = renderBarcodePNG(matrix, { moduleSize: 2 });
    wb.addImage(
      'Inspect',
      {
        fromCol: 0,
        fromRow: 0,
        toCol: 3,
        toRow: 5,
      },
      png,
    );

    const buf = await wb.toBuffer();

    // Open the written ZIP: one drawing part with one chart frame and one picture.
    const zip = readZipText(buf);
    const drawingXml = zip['xl/drawings/drawing1.xml'];
    expect(drawingXml).toBeDefined();
    expect(drawingXml?.match(/<xdr:graphicFrame\b/g)).toHaveLength(1);
    expect(drawingXml?.match(/<xdr:pic\b/g)).toHaveLength(1);
    const drawingRels = zip['xl/drawings/_rels/drawing1.xml.rels'];
    expect(drawingRels).toBeDefined();
    expect(drawingRels).toContain('/relationships/chart"');
    expect(drawingRels).toContain('/relationships/image"');

    const wb2 = await readBuffer(buf);
    const ws2 = wb2.getSheet('Inspect') as Worksheet;

    // The key assertion: exactly 1 chart, not 0 (lost) or 2 (duplicated)
    expect(ws2.charts).toHaveLength(1);
    expect(ws2.charts[0]?.chart.title?.text).toBe('Test');
  });
});
