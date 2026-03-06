/**
 * Auto-sizing helper and top-level barcode generation dispatcher.
 *
 * @module barcode/generate
 */

import { encodeCode39 } from './code39.js';
import { encodeCode128 } from './code128.js';
import type { BarcodeMatrix, DrawBarcodeOptions, RenderOptions } from './common.js';
import { renderBarcodePNG } from './common.js';
import { encodeDataMatrix } from './datamatrix.js';
import { encodeEAN13, encodeUPCA } from './ean13.js';
import { encodeGS1128 } from './gs1128.js';
import { encodeITF14 } from './itf14.js';
import { encodePDF417 } from './pdf417.js';
import { encodeQR } from './qr.js';

export function generateBarcode(value: string, options: DrawBarcodeOptions): Uint8Array {
  let matrix: BarcodeMatrix;

  switch (options.type) {
    case 'qr':
      matrix = encodeQR(value, options.ecLevel ? { ecLevel: options.ecLevel } : undefined);
      break;
    case 'code128':
      matrix = encodeCode128(value);
      break;
    case 'ean13':
      matrix = encodeEAN13(value);
      break;
    case 'upca':
      matrix = encodeUPCA(value);
      break;
    case 'code39':
      matrix = encodeCode39(value);
      break;
    case 'pdf417': {
      const pOpts: { ecLevel?: number; columns?: number } = {};
      if (options.pdf417EcLevel !== undefined) pOpts.ecLevel = options.pdf417EcLevel;
      if (options.pdf417Columns !== undefined) pOpts.columns = options.pdf417Columns;
      matrix = encodePDF417(value, pOpts);
      break;
    }
    case 'datamatrix':
      matrix = encodeDataMatrix(value);
      break;
    case 'itf14':
      matrix = encodeITF14(value);
      break;
    case 'gs1128':
      matrix = encodeGS1128(value);
      break;
    default:
      throw new Error(`Unknown barcode type: ${options.type as string}`);
  }

  let modSz = options.moduleSize ?? 4;
  const qz = options.quietZone ?? (options.type === 'qr' || options.type === 'datamatrix' ? 4 : 10);

  if (options.fitWidth || options.fitHeight) {
    const tmw = matrix.width + 2 * qz;
    const tmh = matrix.height + 2 * qz;
    if (options.fitWidth && options.fitHeight) {
      modSz = Math.max(
        1,
        Math.min(Math.floor(options.fitWidth / tmw), Math.floor(options.fitHeight / tmh)),
      );
    } else if (options.fitWidth) {
      modSz = Math.max(1, Math.floor(options.fitWidth / tmw));
    } else if (options.fitHeight) {
      modSz = Math.max(1, Math.floor(options.fitHeight / tmh));
    }
  }

  const renderOpts: RenderOptions = {
    moduleSize: modSz,
    quietZone: qz,
    textValue: options.textValue ?? value,
  };
  if (options.foreground !== undefined) renderOpts.foreground = options.foreground;
  if (options.background !== undefined) renderOpts.background = options.background;
  if (options.showText !== undefined) renderOpts.showText = options.showText;
  return renderBarcodePNG(matrix, renderOpts);
}
