/**
 * GS1-128 Encoder.
 *
 * @module barcode/gs1128
 */

import { code128Bars, code128Checksum } from './code128.js';
import type { BarcodeMatrix } from './common.js';
import { barsToMatrix } from './linear.js';

// ---------------------------------------------------------------------------
// GS1-128 Encoder
// ---------------------------------------------------------------------------

/**
 * Fixed-length Application Identifiers, keyed by the AI's first two digits, with the value
 * length excluding the AI. A value of such an AI is not followed by FNC1.
 *
 * Source: gs1-syntax-engine `ai.c` (`fixedAIprefixLengths`) and GS1 General Specifications
 * section 7.8.5 (AIs whose data does not need FNC1 separation).
 */
const PREDEF: Readonly<Record<string, number>> = {
  '00': 18,
  '01': 14,
  '02': 14,
  '03': 14,
  '04': 16,
  '11': 6,
  '12': 6,
  '13': 6,
  '14': 6,
  '15': 6,
  '16': 6,
  '17': 6,
  '18': 6,
  '19': 6,
  '20': 2,
  '31': 6,
  '32': 6,
  '33': 6,
  '34': 6,
  '35': 6,
  '36': 6,
  '41': 13,
};

const AI_FORM = /^(\(\d{2,4}\)[^()]+)+$/;
const AI_ELEMENT = /\((\d{2,4})\)([^()]+)/g;

const FNC1 = 102;

/** Code C symbol values for a segment of an even number of digits (two digits per symbol). */
function codeCPairs(seg: string): number[] {
  const out: number[] = [];
  for (let i = 0; i < seg.length; i += 2) {
    out.push((seg.charCodeAt(i) - 48) * 10 + (seg.charCodeAt(i + 1) - 48));
  }
  return out;
}

/** Code B symbol values for a segment of ASCII character codes 32 (space) to 127 (DEL). */
function codeBValues(seg: string): number[] {
  const out: number[] = [];
  for (let i = 0; i < seg.length; i++) {
    const ch = seg.charCodeAt(i);
    if (ch < 32 || ch > 127) throw new Error(`GS1-128: unsupported character code ${ch}`);
    out.push(ch - 32);
  }
  return out;
}

/**
 * Encodes a GS1-128 symbol.
 *
 * With parenthesised input such as `(01)09501101530003(10)ABC`, each Application Identifier is
 * parsed, fixed-length AIs are length-checked, and FNC1 separators are inserted after every
 * variable-length element that is not last. Input without `(` is encoded as one raw run: any `)` is
 * removed and no FNC1 separators are inserted. Element values must
 * not contain parentheses. GS1 values may legitimately contain them (CSET 82), but this parser
 * cannot tell them from AI delimiters, so such input is either rejected as malformed
 * (`(10)A(B)`) or silently split into extra elements (`(10)AB(21)CD`).
 *
 * @param data - Parenthesised AI input or a raw string of ASCII characters 32-127.
 * @returns The symbol matrix (60 rows), without quiet zone.
 * @throws {Error} If the parenthesised syntax is malformed, a fixed-length AI value has the wrong
 *   length, the data is empty, or a character is outside ASCII 32-127.
 */
export function encodeGS1128(data: string): BarcodeMatrix {
  // Each segment is an FNC1-delimited run of characters; FNC1 separates consecutive segments.
  const segments: string[] = [];
  if (data.includes('(')) {
    if (!AI_FORM.test(data)) throw new Error('GS1-128: malformed AI syntax');
    const elements = [...data.matchAll(AI_ELEMENT)];
    let cur = '';
    elements.forEach((m, idx) => {
      const ai = m[1] as string;
      const value = m[2] as string;
      const want = PREDEF[ai.slice(0, 2)];
      if (want !== undefined && value.length !== want) {
        throw new Error(`GS1-128: AI (${ai}) value must be ${want} characters`);
      }
      cur += ai + value;
      if (want === undefined && idx < elements.length - 1) {
        segments.push(cur);
        cur = '';
      }
    });
    segments.push(cur);
  } else {
    segments.push(data.replaceAll(')', ''));
  }
  if (segments.join('').length === 0) throw new Error('GS1-128: data must not be empty');

  const codeC = segments.every((s) => /^\d+$/.test(s) && s.length % 2 === 0);
  const vals: number[] = [codeC ? 105 : 104, FNC1];
  segments.forEach((seg, idx) => {
    if (idx > 0) vals.push(FNC1);
    for (const v of codeC ? codeCPairs(seg) : codeBValues(seg)) vals.push(v);
  });

  vals.push(code128Checksum(vals));

  return barsToMatrix(code128Bars(vals));
}
