/**
 * Auto-sizing helper and top-level barcode generation dispatcher.
 *
 * @module barcode/generate
 */

import { encodeCode39 } from './code39.js';
import { encodeCode128 } from './code128.js';
import type { BarcodeMatrix, BarcodeType, DrawBarcodeOptions, RenderOptions } from './common.js';
import { renderBarcodePNG } from './common.js';
import { encodeDataMatrix } from './datamatrix.js';
import { encodeEAN13, encodeUPCA } from './ean13.js';
import { encodeGS1128 } from './gs1128.js';
import { encodeITF14 } from './itf14.js';
import { encodePDF417 } from './pdf417.js';
import { encodeQR } from './qr.js';

/** One encoder per barcode type; each maps the relevant `DrawBarcodeOptions` fields to its encoder. */
const ENCODERS: Readonly<
  Record<BarcodeType, (value: string, options: DrawBarcodeOptions) => BarcodeMatrix>
> = {
  qr: (value, o) => encodeQR(value, o.ecLevel ? { ecLevel: o.ecLevel } : undefined),
  code128: (value) => encodeCode128(value),
  ean13: (value) => encodeEAN13(value),
  upca: (value) => encodeUPCA(value),
  code39: (value) => encodeCode39(value),
  pdf417: (value, o) => {
    const pOpts: { ecLevel?: number; columns?: number } = {};
    if (o.pdf417EcLevel !== undefined) pOpts.ecLevel = o.pdf417EcLevel;
    if (o.pdf417Columns !== undefined) pOpts.columns = o.pdf417Columns;
    return encodePDF417(value, pOpts);
  },
  datamatrix: (value) => encodeDataMatrix(value),
  itf14: (value) => encodeITF14(value),
  gs1128: (value) => encodeGS1128(value),
};

function encodeByType(value: string, options: DrawBarcodeOptions): BarcodeMatrix {
  const encoder = Object.hasOwn(ENCODERS, options.type) ? ENCODERS[options.type] : undefined;
  if (encoder === undefined) throw new Error(`Unknown barcode type: ${options.type as string}`);
  return encoder(value, options);
}

/** Throws a RangeError unless `v` is undefined or a positive finite number. */
function checkFitSize(name: string, v: number | undefined): void {
  if (v !== undefined && !(Number.isFinite(v) && v > 0)) {
    throw new RangeError(`${name} must be a positive finite number`);
  }
}

/** The module size in pixels: when `fitWidth` or `fitHeight` is set, the largest whole size whose symbol plus quiet zone fits the given bounds (at least 1, replacing `modSz`); otherwise `modSz`. */
function fitModuleSize(
  matrix: BarcodeMatrix,
  qz: number,
  modSz: number,
  options: DrawBarcodeOptions,
): number {
  const { fitWidth, fitHeight } = options;
  const byWidth =
    fitWidth === undefined ? undefined : Math.floor(fitWidth / (matrix.width + 2 * qz));
  const byHeight =
    fitHeight === undefined ? undefined : Math.floor(fitHeight / (matrix.height + 2 * qz));
  if (byWidth === undefined && byHeight === undefined) return modSz;
  return Math.max(
    1,
    Math.min(byWidth ?? Number.POSITIVE_INFINITY, byHeight ?? Number.POSITIVE_INFINITY),
  );
}

/**
 * Encodes `value` and renders it as a PNG, as configured by `options` (size, colours, text, fit).
 *
 * @param value - Data to encode, in the format the chosen `type` accepts.
 * @param options - Symbology, sizing, colours and caption; see {@link DrawBarcodeOptions} for defaults.
 * @returns The PNG file bytes.
 * @throws {RangeError} If `moduleSize` is not a positive safe integer, `quietZone` is not a non-negative safe integer, `fitWidth` or `fitHeight` is not a positive finite number, or `pdf417EcLevel` or `pdf417Columns` is not a finite number, or a foreground or background channel is not an integer 0-255.
 * @throws {Error} If `type` is unknown, or the encoder rejects `value` (see the `encode*` functions).
 */
export function generateBarcode(value: string, options: DrawBarcodeOptions): Uint8Array {
  const matrix = encodeByType(value, options);

  const modSz = options.moduleSize ?? 4;
  const qz = options.quietZone ?? (options.type === 'qr' || options.type === 'datamatrix' ? 4 : 10);
  if (options.moduleSize !== undefined && (!Number.isSafeInteger(modSz) || modSz < 1)) {
    throw new RangeError('moduleSize must be a positive integer');
  }
  if (options.quietZone !== undefined && (!Number.isSafeInteger(qz) || qz < 0)) {
    throw new RangeError('quietZone must be a non-negative integer');
  }
  checkFitSize('fitWidth', options.fitWidth);
  checkFitSize('fitHeight', options.fitHeight);

  const renderOpts: RenderOptions = {
    moduleSize: fitModuleSize(matrix, qz, modSz, options),
    quietZone: qz,
    textValue: options.textValue ?? value,
  };
  if (options.foreground !== undefined) renderOpts.foreground = options.foreground;
  if (options.background !== undefined) renderOpts.background = options.background;
  if (options.showText !== undefined) renderOpts.showText = options.showText;
  return renderBarcodePNG(matrix, renderOpts);
}
