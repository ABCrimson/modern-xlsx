# Barcode & QR Code Guide

Generate barcodes and QR codes directly in XLSX files — pure JavaScript, zero external dependencies.

## Quick Start

```typescript
import { initWasm, Workbook } from 'modern-xlsx';

await initWasm();

const wb = new Workbook();
const ws = wb.addSheet('Labels');

ws.cell('A1').value = 'Product QR:';

// Add a QR code anchored to cells B1:F5
wb.addBarcode(
  'Labels',
  { fromCol: 1, fromRow: 0, toCol: 5, toRow: 4 },
  'https://example.com/product/12345',
  { type: 'qr', ecLevel: 'M' },
);

await wb.toFile('labels.xlsx');
```

## Supported Formats

| Format | Function | Use Case | Capacity |
|--------|----------|----------|----------|
| **QR Code** | `encodeQR()` | URLs, tickets, payments | ~4,296 chars |
| **Code 128** | `encodeCode128()` | SKUs, shipping, logistics | 48 chars |
| **EAN-13** | `encodeEAN13()` | Retail products (international) | 13 digits |
| **UPC-A** | `encodeUPCA()` | Retail products (North America) | 12 digits |
| **Code 39** | `encodeCode39()` | Military, healthcare, automotive | 43 chars |
| **PDF417** | `encodePDF417()` | IDs, boarding passes, licenses | ~1.1 KB |
| **Data Matrix** | `encodeDataMatrix()` | Electronics, pharma, small items | ~2,335 chars |
| **ITF-14** | `encodeITF14()` | Carton-level logistics | 14 digits |
| **GS1-128** | `encodeGS1128()` | Supply chain (structured data) | Variable |

## Which Format Should I Use?

- **URLs, app links, payments** → QR Code
- **Product SKUs, inventory** → Code 128
- **Retail product packaging** → EAN-13 (international) or UPC-A (North America)
- **Shipping cartons** → ITF-14
- **Supply chain with dates/lots** → GS1-128
- **IDs, boarding passes** → PDF417
- **Small electronic components** → Data Matrix
- **Military, healthcare** → Code 39

## Low-Level API

For fine control, use the encoder + renderer pipeline directly:

```typescript
import {
  encodeQR,
  encodeCode128,
  renderBarcodePNG,
  generateDrawingXml,
  generateDrawingRels,
} from 'modern-xlsx';

// 1. Encode data to a boolean matrix
const matrix = encodeQR('Hello World', { ecLevel: 'H' });
console.log(`${matrix.width}x${matrix.height} modules`);

// 2. Render to PNG with options
const png = renderBarcodePNG(matrix, {
  moduleSize: 6,       // 6px per module
  quietZone: 4,        // 4-module white border
  foreground: [0, 0, 0],      // black
  background: [255, 255, 255], // white
  showText: true,
  textValue: 'Hello World',
});

// 3. Use the PNG bytes however you want
// - Embed in XLSX via wb.addImage()
// - Save to disk
// - Display in browser
```

## Rendering Options

| Option | Type | Default | Description |
|--------|------|---------|-------------|
| `moduleSize` | `number` | `4` | Pixels per barcode module |
| `quietZone` | `number` | `4` | Modules of white space around barcode |
| `foreground` | `[R,G,B]` | `[0,0,0]` | Barcode color (black) |
| `background` | `[R,G,B]` | `[255,255,255]` | Background color (white) |
| `showText` | `boolean` | `false` | Show human-readable text below |
| `textValue` | `string` | encoded value | Text to display |

## Auto-Sizing

The `generateBarcode()` helper auto-calculates module size to fit target dimensions:

```typescript
import { generateBarcode } from 'modern-xlsx';

const png = generateBarcode('SKU-12345', {
  type: 'code128',
  fitWidth: 300,   // fit within 300px wide
  fitHeight: 100,  // fit within 100px tall
  showText: true,
});
```

## Embedding in XLSX

### Using `addBarcode()` (recommended)

```typescript
const wb = new Workbook();
wb.addSheet('Sheet1');

// QR code in cells B2:E5
wb.addBarcode('Sheet1',
  { fromCol: 1, fromRow: 1, toCol: 4, toRow: 4 },
  'https://example.com',
  { type: 'qr', ecLevel: 'M' },
);
```

### Using `addImage()` (raw PNG bytes)

```typescript
const wb = new Workbook();
wb.addSheet('Sheet1');

// Generate PNG yourself
const matrix = encodeCode128('ITEM-001');
const png = renderBarcodePNG(matrix, { moduleSize: 3, showText: true, textValue: 'ITEM-001' });

// Embed as image
wb.addImage('Sheet1',
  { fromCol: 0, fromRow: 0, toCol: 3, toRow: 2 },
  png,
);
```

## Error Correction (QR Only)

| Level | Recovery | Best For |
|-------|----------|----------|
| `L` (Low) | ~7% | Clean environments, smallest size |
| `M` (Medium) | ~15% | General purpose (default) |
| `Q` (Quartile) | ~25% | Industrial, some damage expected |
| `H` (High) | ~30% | Harsh environments, logos overlaid |

```typescript
encodeQR('data', { ecLevel: 'H' }); // Maximum error correction
```

## GS1-128 Application Identifiers

GS1-128 encodes structured data using Application Identifiers (AIs):

```typescript
// AI (01) = GTIN, AI (17) = Expiry date, AI (10) = Batch
const barcode = encodeGS1128('(01)09501101530003(17)250301(10)LOT123');
```

| AI | Description | Example |
|----|-------------|---------|
| `01` | GTIN (product ID) | `(01)09501101530003` |
| `10` | Batch/lot number | `(10)ABC123` |
| `17` | Expiry date (YYMMDD) | `(17)250301` |
| `21` | Serial number | `(21)SN12345` |
| `37` | Quantity | `(37)100` |
