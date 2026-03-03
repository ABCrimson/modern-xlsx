import { beforeAll, describe, expect, it } from 'vitest';
import {
  encodeCode39,
  encodeCode128,
  encodeDataMatrix,
  encodeEAN13,
  encodeGS1128,
  encodeITF14,
  encodePDF417,
  encodeQR,
  encodeUPCA,
  generateBarcode,
  generateDrawingRels,
  generateDrawingXml,
  initWasm,
  renderBarcodePNG,
  Workbook,
} from '../src/index';

beforeAll(async () => {
  await initWasm();
});

describe('Barcode Encoders', () => {
  it('encodes QR code with valid matrix', () => {
    const matrix = encodeQR('Hello');
    expect(matrix.width).toBeGreaterThan(0);
    expect(matrix.height).toBe(matrix.width); // QR codes are square
    expect(matrix.modules.length).toBe(matrix.height);
    expect(matrix.modules[0]?.length).toBe(matrix.width);
  });

  it('encodes QR with different EC levels', () => {
    const l = encodeQR('Test', { ecLevel: 'L' });
    const h = encodeQR('Test', { ecLevel: 'H' });
    // Higher EC level = larger matrix
    expect(h.width).toBeGreaterThanOrEqual(l.width);
  });

  it('encodes Code 128', () => {
    const matrix = encodeCode128('ABC-123');
    expect(matrix.width).toBeGreaterThan(0);
    expect(matrix.height).toBeGreaterThan(0);
    // Code 128 is 1D — height should be uniform
    expect(matrix.modules.length).toBe(matrix.height);
  });

  it('encodes EAN-13', () => {
    const matrix = encodeEAN13('5901234123457');
    expect(matrix.width).toBe(95); // EAN-13 is always 95 modules wide
    expect(matrix.height).toBeGreaterThan(0);
  });

  it('encodes EAN-13 with auto check digit', () => {
    // 12 digits — check digit should be auto-calculated
    const matrix = encodeEAN13('590123412345');
    expect(matrix.width).toBe(95);
  });

  it('encodes UPC-A', () => {
    const matrix = encodeUPCA('012345678905');
    expect(matrix.width).toBe(95); // UPC-A wraps EAN-13
  });

  it('encodes Code 39', () => {
    const matrix = encodeCode39('HELLO');
    expect(matrix.width).toBeGreaterThan(0);
    expect(matrix.height).toBeGreaterThan(0);
  });

  it('encodes ITF-14', () => {
    const matrix = encodeITF14('00012345600012');
    expect(matrix.width).toBeGreaterThan(0);
  });

  it('encodes GS1-128', () => {
    const matrix = encodeGS1128('(01)09501101530003');
    expect(matrix.width).toBeGreaterThan(0);
  });

  it('encodes PDF417', () => {
    const matrix = encodePDF417('Hello World');
    expect(matrix.width).toBeGreaterThan(0);
    expect(matrix.height).toBeGreaterThan(0);
    // PDF417 is 2D stacked
    expect(matrix.height).toBeGreaterThan(1);
  });

  it('encodes Data Matrix', () => {
    const matrix = encodeDataMatrix('Hello');
    expect(matrix.width).toBeGreaterThan(0);
    expect(matrix.height).toBe(matrix.width); // Data Matrix is square
  });
});

describe('PNG Renderer', () => {
  it('renders barcode to valid PNG bytes', () => {
    const matrix = encodeQR('Test');
    const png = renderBarcodePNG(matrix);
    expect(png).toBeInstanceOf(Uint8Array);
    expect(png.length).toBeGreaterThan(0);
    // Check PNG magic bytes
    expect(png[0]).toBe(0x89);
    expect(png[1]).toBe(0x50); // P
    expect(png[2]).toBe(0x4e); // N
    expect(png[3]).toBe(0x47); // G
  });

  it('respects moduleSize option', () => {
    const matrix = encodeQR('A');
    const small = renderBarcodePNG(matrix, { moduleSize: 2 });
    const large = renderBarcodePNG(matrix, { moduleSize: 8 });
    expect(large.length).toBeGreaterThan(small.length);
  });

  it('renders with showText option', () => {
    const matrix = encodeCode128('123');
    const withText = renderBarcodePNG(matrix, { showText: true, textValue: '123' });
    const withoutText = renderBarcodePNG(matrix, { showText: false });
    expect(withText.length).toBeGreaterThan(withoutText.length);
  });
});

describe('generateBarcode helper', () => {
  it('generates QR code PNG', () => {
    const png = generateBarcode('https://example.com', { type: 'qr' });
    expect(png[0]).toBe(0x89); // PNG magic
  });

  it('generates Code 128 PNG', () => {
    const png = generateBarcode('ABC-123', { type: 'code128' });
    expect(png[0]).toBe(0x89);
  });

  it('generates with fit dimensions', () => {
    const png = generateBarcode('Test', {
      type: 'qr',
      fitWidth: 200,
      fitHeight: 200,
    });
    expect(png.length).toBeGreaterThan(0);
  });
});

describe('Drawing XML Generation', () => {
  it('generates valid drawing XML', () => {
    const xml = generateDrawingXml([
      {
        anchor: { fromCol: 0, fromRow: 0, toCol: 3, toRow: 5 },
        imageIndex: 1,
      },
    ]);
    expect(xml).toContain('xdr:wsDr');
    expect(xml).toContain('xdr:twoCellAnchor');
    expect(xml).toContain('xdr:from');
    expect(xml).toContain('xdr:to');
  });

  it('generates valid drawing rels', () => {
    const xml = generateDrawingRels([{ rId: 'rId1', target: '../media/image1.png' }]);
    expect(xml).toContain('Relationships');
    expect(xml).toContain('rId1');
    expect(xml).toContain('../media/image1.png');
  });
});

describe('Workbook barcode integration', () => {
  it('adds barcode image to workbook preservedEntries', () => {
    const wb = new Workbook();
    wb.addSheet('Sheet1');
    wb.addBarcode('Sheet1', { fromCol: 0, fromRow: 0, toCol: 4, toRow: 4 }, 'Hello', {
      type: 'qr',
    });

    const data = wb.toJSON();
    expect(data.preservedEntries).toBeDefined();

    // Should have media, drawing, rels entries
    const entries = Object.keys(data.preservedEntries ?? {});
    expect(entries.some((e) => e.startsWith('xl/media/'))).toBe(true);
    expect(entries.some((e) => e.startsWith('xl/drawings/'))).toBe(true);
    expect(entries.some((e) => e.includes('_rels'))).toBe(true);
  });

  it('writes workbook with barcode to valid XLSX', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Barcodes');
    ws.cell('A1').value = 'QR Code:';
    wb.addBarcode(
      'Barcodes',
      { fromCol: 1, fromRow: 0, toCol: 5, toRow: 4 },
      'https://example.com',
      {
        type: 'qr',
      },
    );

    const buffer = await wb.toBuffer();
    expect(buffer).toBeInstanceOf(Uint8Array);
    expect(buffer.length).toBeGreaterThan(0);
  });

  it('addImage embeds raw PNG bytes', () => {
    const wb = new Workbook();
    wb.addSheet('Sheet1');

    // Minimal valid 1x1 PNG (hand-crafted)
    const qr = encodeQR('Test');
    const png = renderBarcodePNG(qr, { moduleSize: 1 });

    wb.addImage('Sheet1', { fromCol: 0, fromRow: 0, toCol: 2, toRow: 2 }, png);

    const data = wb.toJSON();
    const entries = Object.keys(data.preservedEntries ?? {});
    expect(entries.some((e) => e.includes('image1.png'))).toBe(true);
  });
});
