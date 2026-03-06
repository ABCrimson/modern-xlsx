/**
 * Barcode & QR Code generation module for modern-xlsx.
 *
 * Exports encoders for QR Code, Code 128, EAN-13, UPC-A, Code 39, PDF417,
 * Data Matrix, ITF-14, and GS1-128. Includes a minimal PNG renderer and
 * OOXML drawing XML generator for embedding images in XLSX worksheets.
 *
 * Zero external dependencies. Tree-shakable named exports only.
 */

export { encodeCode39 } from './code39.js';
export { encodeCode128 } from './code128.js';
// Types & common utilities
export type {
  BarcodeMatrix,
  BarcodeType,
  DrawBarcodeOptions,
  ImageAnchor,
  RenderOptions,
} from './common.js';
export {
  generateDrawingRels,
  generateDrawingXml,
  renderBarcodePNG,
} from './common.js';
export { encodeDataMatrix } from './datamatrix.js';
export { encodeEAN13, encodeUPCA } from './ean13.js';
// Auto-sizing helper & generateBarcode
export { generateBarcode } from './generate.js';
export { encodeGS1128 } from './gs1128.js';
export { encodeITF14 } from './itf14.js';
export { encodePDF417 } from './pdf417.js';
// Per-codec exports
export { encodeQR } from './qr.js';
