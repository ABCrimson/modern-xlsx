# v0.5.0 "Parity Sprint" Implementation Plan

> **For Claude:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task.

**Goal:** Close all easy SheetJS gaps in one release — flip 7 of 12 losing/tied categories to wins.

**Architecture:** Pure TypeScript additions + minor Rust struct field additions. No new parsers, no new WASM exports. All changes are backward-compatible additions to existing modules.

**Tech Stack:** TypeScript 6.0, Rust 1.94.0, Vitest 4.1.0-beta.5, quick-xml 0.39.2, serde

---

## Build & Test Commands

```bash
# Rust tests
cargo test -p modern-xlsx-core

# WASM build (after Rust changes)
cd crates/modern-xlsx-wasm && wasm-pack build --target web --release --out-dir ../../packages/modern-xlsx/wasm --no-opt

# TypeScript tests
pnpm -C packages/modern-xlsx test

# Typecheck
pnpm -C packages/modern-xlsx typecheck

# Lint
pnpm -C packages/modern-xlsx lint
```

---

## Task 1: Cell Reference Utilities — `encodeRow`, `decodeRow`, `splitCellRef`

**Files:**
- Modify: `packages/modern-xlsx/src/cell-ref.ts`
- Modify: `packages/modern-xlsx/src/index.ts`
- Create: `packages/modern-xlsx/__tests__/cell-ref-extended.test.ts`

**Step 1: Write the failing tests**

```typescript
// packages/modern-xlsx/__tests__/cell-ref-extended.test.ts
import { describe, expect, it } from 'vitest';
import { decodeRow, encodeRow, splitCellRef } from '../src/index.js';

describe('extended cell reference utilities', () => {
  describe('encodeRow', () => {
    it('converts 0-based index to 1-based string', () => {
      expect(encodeRow(0)).toBe('1');
      expect(encodeRow(9)).toBe('10');
      expect(encodeRow(1048575)).toBe('1048576');
    });
  });

  describe('decodeRow', () => {
    it('converts 1-based string to 0-based index', () => {
      expect(decodeRow('1')).toBe(0);
      expect(decodeRow('10')).toBe(9);
      expect(decodeRow('1048576')).toBe(1048575);
    });

    it('throws on invalid input', () => {
      expect(() => decodeRow('')).toThrow();
      expect(() => decodeRow('0')).toThrow();
      expect(() => decodeRow('abc')).toThrow();
    });
  });

  describe('splitCellRef', () => {
    it('splits simple reference', () => {
      expect(splitCellRef('A1')).toEqual({
        col: 'A', row: '1', absCol: false, absRow: false,
      });
    });

    it('splits fully absolute reference', () => {
      expect(splitCellRef('$A$1')).toEqual({
        col: 'A', row: '1', absCol: true, absRow: true,
      });
    });

    it('splits mixed references', () => {
      expect(splitCellRef('$A1')).toEqual({
        col: 'A', row: '1', absCol: true, absRow: false,
      });
      expect(splitCellRef('A$1')).toEqual({
        col: 'A', row: '1', absCol: false, absRow: true,
      });
    });

    it('handles multi-letter columns', () => {
      expect(splitCellRef('$XFD$1048576')).toEqual({
        col: 'XFD', row: '1048576', absCol: true, absRow: true,
      });
    });

    it('throws on invalid ref', () => {
      expect(() => splitCellRef('')).toThrow();
      expect(() => splitCellRef('123')).toThrow();
    });
  });
});
```

**Step 2: Run tests to verify they fail**

Run: `pnpm -C packages/modern-xlsx test -- __tests__/cell-ref-extended.test.ts`
Expected: FAIL — `encodeRow`, `decodeRow`, `splitCellRef` are not exported

**Step 3: Implement**

Add to `packages/modern-xlsx/src/cell-ref.ts` at the end of the file:

```typescript
/** Result of splitting a cell reference into its components. */
export interface SplitCellRef {
  /** Column letters (e.g., "A", "XFD"). */
  readonly col: string;
  /** Row number as string (1-based, e.g., "1", "1048576"). */
  readonly row: string;
  /** Whether the column is absolute ($A). */
  readonly absCol: boolean;
  /** Whether the row is absolute ($1). */
  readonly absRow: boolean;
}

/**
 * Convert a 0-based row index to a 1-based row string.
 * `encodeRow(0)` → `"1"`
 */
export function encodeRow(row: number): string {
  return String(row + 1);
}

/**
 * Convert a 1-based row string to a 0-based row index.
 * `decodeRow("1")` → `0`
 */
export function decodeRow(rowStr: string): number {
  const n = Number.parseInt(rowStr, 10);
  if (!Number.isFinite(n) || n < 1) {
    throw new Error(`Invalid row string: ${rowStr}`);
  }
  return n - 1;
}

const SPLIT_RE = /^(\$?)([A-Z]+)(\$?)(\d+)$/;

/**
 * Split a cell reference into column/row parts with absolute flags.
 * `splitCellRef("$A$1")` → `{ col: "A", row: "1", absCol: true, absRow: true }`
 */
export function splitCellRef(ref: string): SplitCellRef {
  const match = ref.toUpperCase().match(SPLIT_RE);
  if (!match?.[2] || !match[4]) {
    throw new Error(`Invalid cell reference: ${ref}`);
  }
  return {
    col: match[2],
    row: match[4],
    absCol: match[1] === '$',
    absRow: match[3] === '$',
  };
}
```

Add exports to `packages/modern-xlsx/src/index.ts` — in the cell-ref section:

```typescript
export type { CellAddress, CellRange, SplitCellRef } from './cell-ref.js';
export {
  columnToLetter,
  decodeCellRef,
  decodeRange,
  decodeRow,
  encodeCellRef,
  encodeRange,
  encodeRow,
  letterToColumn,
  splitCellRef,
} from './cell-ref.js';
```

**Step 4: Run tests to verify they pass**

Run: `pnpm -C packages/modern-xlsx test -- __tests__/cell-ref-extended.test.ts`
Expected: PASS (all 9 tests)

**Step 5: Commit**

```bash
git add packages/modern-xlsx/src/cell-ref.ts packages/modern-xlsx/src/index.ts packages/modern-xlsx/__tests__/cell-ref-extended.test.ts
git commit -m "feat: add encodeRow, decodeRow, splitCellRef utilities"
```

---

## Task 2: Document Properties — `appVersion`, `hyperlinkBase`, `revision`

**Files:**
- Modify: `crates/modern-xlsx-core/src/ooxml/doc_props.rs`
- Modify: `packages/modern-xlsx/src/types.ts`
- Create: `packages/modern-xlsx/__tests__/doc-props-extended.test.ts`

**Step 1: Write the failing test**

```typescript
// packages/modern-xlsx/__tests__/doc-props-extended.test.ts
import { describe, expect, it } from 'vitest';
import { readBuffer, Workbook } from '../src/index.js';

describe('extended document properties', () => {
  it('roundtrips appVersion', async () => {
    const wb = new Workbook();
    wb.addSheet('Sheet1').cell('A1').value = 1;
    wb.docProperties = {
      title: 'Test',
      appVersion: '16.0300',
    };
    const buf = await wb.toBuffer();
    const wb2 = await readBuffer(buf);
    expect(wb2.docProperties?.appVersion).toBe('16.0300');
  });

  it('roundtrips hyperlinkBase', async () => {
    const wb = new Workbook();
    wb.addSheet('Sheet1').cell('A1').value = 1;
    wb.docProperties = {
      title: 'Test',
      hyperlinkBase: 'https://example.com/',
    };
    const buf = await wb.toBuffer();
    const wb2 = await readBuffer(buf);
    expect(wb2.docProperties?.hyperlinkBase).toBe('https://example.com/');
  });

  it('roundtrips revision', async () => {
    const wb = new Workbook();
    wb.addSheet('Sheet1').cell('A1').value = 1;
    wb.docProperties = {
      title: 'Test',
      revision: '4',
    };
    const buf = await wb.toBuffer();
    const wb2 = await readBuffer(buf);
    expect(wb2.docProperties?.revision).toBe('4');
  });
});
```

**Step 2: Run tests to verify they fail**

Run: `pnpm -C packages/modern-xlsx test -- __tests__/doc-props-extended.test.ts`
Expected: FAIL — TypeScript type error on `appVersion`, `hyperlinkBase`, `revision`

**Step 3: Add Rust fields**

In `crates/modern-xlsx-core/src/ooxml/doc_props.rs`, add 3 fields to `DocProperties` struct after `manager`:

```rust
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub app_version: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub hyperlink_base: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub revision: Option<String>,
```

In `parse_app()`, add parsing inside the match arm for app.xml elements:

```rust
b"AppVersion" => {
    props.app_version = Some(reader.read_text(e.name())?.into_owned());
}
b"HyperlinkBase" => {
    props.hyperlink_base = Some(reader.read_text(e.name())?.into_owned());
}
```

In `parse_core()`, add parsing for the `cp:revision` element:

```rust
b"cp:revision" => {
    props.revision = Some(reader.read_text(e.name())?.into_owned());
}
```

In `to_app_xml()`, add writing after the `Manager` element:

```rust
if let Some(ref v) = self.app_version {
    writer.create_element("AppVersion").write_text_content(BytesText::new(v))?;
}
if let Some(ref v) = self.hyperlink_base {
    writer.create_element("HyperlinkBase").write_text_content(BytesText::new(v))?;
}
```

In `to_core_xml()`, add writing after the last `dcterms` element:

```rust
if let Some(ref v) = self.revision {
    writer.create_element("cp:revision").write_text_content(BytesText::new(v))?;
}
```

Update `has_app()` and `has_core()` to check new fields.

**Step 4: Run Rust tests**

Run: `cargo test -p modern-xlsx-core`
Expected: PASS

**Step 5: Add TypeScript types**

In `packages/modern-xlsx/src/types.ts`, add to `DocPropertiesData`:

```typescript
  appVersion?: string | null;
  hyperlinkBase?: string | null;
  revision?: string | null;
```

**Step 6: Rebuild WASM**

Run: `cd crates/modern-xlsx-wasm && wasm-pack build --target web --release --out-dir ../../packages/modern-xlsx/wasm --no-opt`

**Step 7: Run TS tests**

Run: `pnpm -C packages/modern-xlsx test -- __tests__/doc-props-extended.test.ts`
Expected: PASS

**Step 8: Commit**

```bash
git add crates/modern-xlsx-core/src/ooxml/doc_props.rs packages/modern-xlsx/src/types.ts packages/modern-xlsx/__tests__/doc-props-extended.test.ts
git commit -m "feat: add appVersion, hyperlinkBase, revision to doc properties"
```

---

## Task 3: Sheet Conversion Utilities — `sheetToTxt`, `sheetToFormulae`

**Files:**
- Modify: `packages/modern-xlsx/src/utils.ts`
- Modify: `packages/modern-xlsx/src/index.ts`
- Create: `packages/modern-xlsx/__tests__/utils-extended.test.ts`

**Step 1: Write the failing tests**

```typescript
// packages/modern-xlsx/__tests__/utils-extended.test.ts
import { describe, expect, it } from 'vitest';
import { sheetToFormulae, sheetToTxt, Workbook } from '../src/index.js';

describe('extended sheet conversion utilities', () => {
  describe('sheetToTxt', () => {
    it('produces tab-separated output', () => {
      const wb = new Workbook();
      const ws = wb.addSheet('S');
      ws.cell('A1').value = 'Name';
      ws.cell('B1').value = 'Age';
      ws.cell('A2').value = 'Alice';
      ws.cell('B2').value = 30;

      const txt = sheetToTxt(ws);
      const lines = txt.split('\n');
      expect(lines[0]).toBe('Name\tAge');
      expect(lines[1]).toBe('Alice\t30');
    });

    it('respects sheetRows limit', () => {
      const wb = new Workbook();
      const ws = wb.addSheet('S');
      ws.cell('A1').value = 'r1';
      ws.cell('A2').value = 'r2';
      ws.cell('A3').value = 'r3';

      const txt = sheetToTxt(ws, { sheetRows: 2 });
      expect(txt.split('\n')).toHaveLength(2);
    });
  });

  describe('sheetToFormulae', () => {
    it('returns cell references with values and formulas', () => {
      const wb = new Workbook();
      const ws = wb.addSheet('S');
      ws.cell('A1').value = 100;
      ws.cell('A2').value = 200;
      ws.cell('A3').formula = 'SUM(A1:A2)';

      const formulae = sheetToFormulae(ws);
      expect(formulae).toContain('A1=100');
      expect(formulae).toContain('A2=200');
      expect(formulae).toContain("A3='SUM(A1:A2)");
    });

    it('handles string values', () => {
      const wb = new Workbook();
      const ws = wb.addSheet('S');
      ws.cell('A1').value = 'hello';

      const formulae = sheetToFormulae(ws);
      expect(formulae).toContain("A1='hello");
    });

    it('handles empty sheet', () => {
      const wb = new Workbook();
      const ws = wb.addSheet('S');
      expect(sheetToFormulae(ws)).toEqual([]);
    });
  });
});
```

**Step 2: Run tests to verify they fail**

Run: `pnpm -C packages/modern-xlsx test -- __tests__/utils-extended.test.ts`
Expected: FAIL — `sheetToTxt` and `sheetToFormulae` not exported

**Step 3: Implement**

Add to `packages/modern-xlsx/src/utils.ts` after `sheetToCsv`:

```typescript
// ---------------------------------------------------------------------------
// sheetToTxt — tab-separated text output
// ---------------------------------------------------------------------------

/** Options for the {@link sheetToTxt} function. */
export interface SheetToTxtOptions {
  /** Maximum number of rows to include. */
  sheetRows?: number;
}

/**
 * Convert a Worksheet to a tab-separated text string.
 */
export function sheetToTxt(ws: Worksheet, opts?: SheetToTxtOptions): string {
  return sheetToCsv(ws, { separator: '\t', sheetRows: opts?.sheetRows });
}

// ---------------------------------------------------------------------------
// sheetToFormulae — extract all cell values and formulas
// ---------------------------------------------------------------------------

/**
 * Extract all cell values and formulas as an array of strings.
 *
 * Format: `"A1=100"` for values, `"A3='SUM(A1:A2)"` for formulas.
 * String values are prefixed with `'` to distinguish from numbers.
 */
export function sheetToFormulae(ws: Worksheet): string[] {
  const result: string[] = [];
  for (const row of ws.rows) {
    for (const cell of row.cells) {
      if (cell.formula) {
        result.push(`${cell.reference}='${cell.formula}`);
      } else if (cell.value != null) {
        const val = cell.cellType === 'number' || cell.cellType === 'boolean'
          ? cell.value
          : `'${cell.value}`;
        result.push(`${cell.reference}=${val}`);
      }
    }
  }
  return result;
}
```

Add exports to `packages/modern-xlsx/src/index.ts`:

```typescript
export type {
  AoaToSheetOptions,
  JsonToSheetOptions,
  SheetAddAoaOptions,
  SheetAddJsonOptions,
  SheetToCsvOptions,
  SheetToHtmlOptions,
  SheetToJsonOptions,
  SheetToTxtOptions,
} from './utils.js';
export {
  aoaToSheet,
  jsonToSheet,
  sheetAddAoa,
  sheetAddJson,
  sheetToCsv,
  sheetToFormulae,
  sheetToHtml,
  sheetToJson,
  sheetToTxt,
} from './utils.js';
```

**Step 4: Run tests**

Run: `pnpm -C packages/modern-xlsx test -- __tests__/utils-extended.test.ts`
Expected: PASS

**Step 5: Commit**

```bash
git add packages/modern-xlsx/src/utils.ts packages/modern-xlsx/src/index.ts packages/modern-xlsx/__tests__/utils-extended.test.ts
git commit -m "feat: add sheetToTxt and sheetToFormulae utilities"
```

---

## Task 4: Number Formatting — Conditional Sections, Color Codes, `loadFormat`

**Files:**
- Modify: `packages/modern-xlsx/src/format-cell.ts`
- Modify: `packages/modern-xlsx/src/index.ts`
- Create: `packages/modern-xlsx/__tests__/format-cell-extended.test.ts`

**Step 1: Write the failing tests**

```typescript
// packages/modern-xlsx/__tests__/format-cell-extended.test.ts
import { describe, expect, it } from 'vitest';
import { formatCellRich, loadFormat, loadFormatTable } from '../src/index.js';

describe('extended number formatting', () => {
  describe('conditional sections', () => {
    it('applies [>100] condition', () => {
      const fmt = '[>100]"High";[<=100]"Low"';
      expect(formatCellRich(200, fmt).text).toBe('High');
      expect(formatCellRich(50, fmt).text).toBe('Low');
    });

    it('applies [Red] condition with number', () => {
      const fmt = '[Red][>100]#,##0;[Blue]#,##0';
      const r1 = formatCellRich(200, fmt);
      expect(r1.text).toBe('200');
      expect(r1.color).toBe('Red');

      const r2 = formatCellRich(50, fmt);
      expect(r2.text).toBe('50');
      expect(r2.color).toBe('Blue');
    });
  });

  describe('bracket color codes', () => {
    it('extracts named colors', () => {
      expect(formatCellRich(42, '[Red]0').color).toBe('Red');
      expect(formatCellRich(42, '[Blue]0.00').color).toBe('Blue');
      expect(formatCellRich(42, '[Green]#,##0').color).toBe('Green');
    });

    it('extracts indexed colors', () => {
      expect(formatCellRich(42, '[Color3]0').color).toBe('Color3');
      expect(formatCellRich(42, '[Color56]0').color).toBe('Color56');
    });

    it('returns no color when none specified', () => {
      expect(formatCellRich(42, '0.00').color).toBeUndefined();
    });
  });

  describe('loadFormat / loadFormatTable', () => {
    it('registers and uses a custom format', () => {
      loadFormat('#,##0.000', 200);
      // The format should be available via getBuiltinFormat or internal lookup
      expect(formatCellRich(1234.5, 200).text).toBe('1,234.500');
    });

    it('bulk-registers formats', () => {
      loadFormatTable({ 201: '0.0%', 202: '#,##0.00' });
      expect(formatCellRich(0.456, 201).text).toBe('45.6%');
      expect(formatCellRich(1234, 202).text).toBe('1,234.00');
    });
  });
});
```

**Step 2: Run tests to verify they fail**

Run: `pnpm -C packages/modern-xlsx test -- __tests__/format-cell-extended.test.ts`
Expected: FAIL — `formatCellRich`, `loadFormat`, `loadFormatTable` not exported

**Step 3: Implement**

In `packages/modern-xlsx/src/format-cell.ts`:

1. Add custom format registry at module level (after BUILTIN_FORMATS):

```typescript
/** Runtime-registered custom format codes. */
const CUSTOM_FORMATS = new Map<number, string>();

/**
 * Register a custom number format code at runtime.
 */
export function loadFormat(formatCode: string, id: number): void {
  CUSTOM_FORMATS.set(id, formatCode);
}

/**
 * Bulk-register multiple format codes at runtime.
 */
export function loadFormatTable(table: Record<number, string>): void {
  for (const [id, code] of Object.entries(table)) {
    CUSTOM_FORMATS.set(Number(id), code);
  }
}
```

2. Add `FormatCellResult` type and `formatCellRich` function:

```typescript
/** Result of rich cell formatting, including optional color metadata. */
export interface FormatCellResult {
  /** The formatted text string. */
  text: string;
  /** Color name from bracket directive (e.g., "Red", "Color3"). */
  color?: string;
}

/**
 * Format a cell value and return rich result with color metadata.
 */
export function formatCellRich(
  value: string | number | boolean | null,
  format: string | number,
  opts?: FormatCellOptions,
): FormatCellResult {
  if (value === null || value === undefined) return { text: '' };
  if (typeof value === 'boolean') return { text: value ? 'TRUE' : 'FALSE' };

  let formatCode: string;
  if (typeof format === 'number') {
    formatCode = BUILTIN_FORMATS[format] ?? CUSTOM_FORMATS.get(format) ?? 'General';
  } else {
    formatCode = format;
  }

  if (formatCode === 'General' || formatCode === '' || formatCode === '@') {
    return { text: String(value) };
  }

  const numVal = typeof value === 'number' ? value : Number.parseFloat(String(value));
  if (Number.isNaN(numVal)) return { text: String(value) };

  return dispatchFormatRich(numVal, formatCode, opts?.dateSystem ?? 'date1900');
}
```

3. Add conditional section parser and color extraction:

```typescript
const COLOR_RE = /\[(Red|Blue|Green|Yellow|Magenta|Cyan|White|Black|Color\d{1,2})\]/i;
const CONDITION_RE = /\[([<>=!]+)([\d.]+)\]/;

function dispatchFormatRich(numVal: number, code: string, system: DateSystem): FormatCellResult {
  // Split into sections, resolve conditional
  const { section, value } = resolveSectionConditional(code, numVal);

  // Extract color
  const colorMatch = section.match(COLOR_RE);
  const color = colorMatch?.[1];

  // Strip all bracket directives (colors + conditions)
  const cleaned = section
    .replace(COLOR_RE, '')
    .replace(CONDITION_RE, '')
    .trim();

  if (cleaned === '' || cleaned === 'General') {
    return { text: String(value), color };
  }

  const text = isDateFormatCode(cleaned)
    ? formatDate(value, cleaned, system)
    : cleaned.includes('%')
      ? formatPercentage(value, cleaned)
      : cleaned.includes('E+') || cleaned.includes('E-') || cleaned.includes('e+')
        ? formatScientific(value, cleaned)
        : cleaned.includes('?/') || cleaned.includes('#/')
          ? formatFraction(value)
          : applyNumberFormat(value, cleaned.replace(/\[(?:Red|Blue|Green|Yellow|Magenta|Cyan|White|Black|Color\d+)\]/gi, ''));

  return { text, color };
}

function resolveSectionConditional(code: string, value: number): { section: string; value: number } {
  const sections = splitSections(code);

  // Check for explicit conditions [>100], [<=50], etc.
  for (const section of sections) {
    const condMatch = section.match(CONDITION_RE);
    if (condMatch?.[1] && condMatch[2]) {
      const op = condMatch[1];
      const threshold = Number.parseFloat(condMatch[2]);
      if (evaluateCondition(value, op, threshold)) {
        return { section, value };
      }
    }
  }

  // If no explicit conditions, fall back to standard pos;neg;zero;text
  if (sections.length >= 3 && value === 0) {
    return { section: sections[2] ?? sections[0] ?? 'General', value };
  }
  if (sections.length >= 2 && value < 0) {
    return { section: sections[1] ?? sections[0] ?? 'General', value: Math.abs(value) };
  }
  return { section: sections[0] ?? 'General', value };
}

function evaluateCondition(value: number, op: string, threshold: number): boolean {
  switch (op) {
    case '>': return value > threshold;
    case '<': return value < threshold;
    case '>=': return value >= threshold;
    case '<=': return value <= threshold;
    case '=': return value === threshold;
    case '<>': case '!=': return value !== threshold;
    default: return false;
  }
}
```

4. Update `formatCell` to also resolve from CUSTOM_FORMATS:

```typescript
// In formatCell(), change line 69:
const formatCode = typeof format === 'number'
  ? (BUILTIN_FORMATS[format] ?? CUSTOM_FORMATS.get(format) ?? 'General')
  : format;
```

5. Add exports to `packages/modern-xlsx/src/index.ts`:

```typescript
export type { FormatCellOptions, FormatCellResult } from './format-cell.js';
export { formatCell, formatCellRich, getBuiltinFormat, loadFormat, loadFormatTable } from './format-cell.js';
```

**Step 4: Run tests**

Run: `pnpm -C packages/modern-xlsx test -- __tests__/format-cell-extended.test.ts`
Expected: PASS

**Step 5: Commit**

```bash
git add packages/modern-xlsx/src/format-cell.ts packages/modern-xlsx/src/index.ts packages/modern-xlsx/__tests__/format-cell-extended.test.ts
git commit -m "feat: add conditional sections, color codes, loadFormat to number formatting"
```

---

## Task 5: Worksheet Operations — `usedRange`, `tabColor`

**Files:**
- Modify: `packages/modern-xlsx/src/workbook.ts`
- Modify: `packages/modern-xlsx/src/types.ts`
- Modify: `crates/modern-xlsx-core/src/ooxml/worksheet.rs`
- Modify: `crates/modern-xlsx-core/src/lib.rs` (SheetData)
- Create: `packages/modern-xlsx/__tests__/worksheet-extended.test.ts`

**Step 1: Write the failing tests**

```typescript
// packages/modern-xlsx/__tests__/worksheet-extended.test.ts
import { describe, expect, it } from 'vitest';
import { readBuffer, Workbook } from '../src/index.js';

describe('extended worksheet operations', () => {
  describe('usedRange', () => {
    it('computes used range from cells', () => {
      const wb = new Workbook();
      const ws = wb.addSheet('S');
      ws.cell('B2').value = 'hello';
      ws.cell('D5').value = 42;

      expect(ws.usedRange).toBe('B2:D5');
    });

    it('returns null for empty sheet', () => {
      const wb = new Workbook();
      const ws = wb.addSheet('S');
      expect(ws.usedRange).toBeNull();
    });

    it('handles single cell', () => {
      const wb = new Workbook();
      const ws = wb.addSheet('S');
      ws.cell('C3').value = 'only';
      expect(ws.usedRange).toBe('C3:C3');
    });
  });

  describe('tabColor', () => {
    it('roundtrips tab color', async () => {
      const wb = new Workbook();
      const ws = wb.addSheet('Colored');
      ws.cell('A1').value = 1;
      ws.tabColor = 'FF0000';

      expect(ws.tabColor).toBe('FF0000');

      const buf = await wb.toBuffer();
      const wb2 = await readBuffer(buf);
      expect(wb2.getSheet('Colored')?.tabColor).toBe('FF0000');
    });

    it('defaults to null', () => {
      const wb = new Workbook();
      const ws = wb.addSheet('S');
      expect(ws.tabColor).toBeNull();
    });

    it('can be cleared', () => {
      const wb = new Workbook();
      const ws = wb.addSheet('S');
      ws.tabColor = 'FF0000';
      ws.tabColor = null;
      expect(ws.tabColor).toBeNull();
    });
  });
});
```

**Step 2: Run tests to verify they fail**

Run: `pnpm -C packages/modern-xlsx test -- __tests__/worksheet-extended.test.ts`
Expected: FAIL — `usedRange` and `tabColor` do not exist

**Step 3a: Add Rust fields**

In `crates/modern-xlsx-core/src/lib.rs`, add `tab_color` to `SheetData`:

```rust
pub struct SheetData {
    pub name: String,
    pub worksheet: WorksheetXml,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tab_color: Option<String>,
}
```

In `crates/modern-xlsx-core/src/ooxml/worksheet.rs`, update the parser to extract `<sheetPr><tabColor rgb="..."/>` and the writer to emit it.

In the worksheet parser (`parse` method), look for `<sheetPr>` → `<tabColor>` and extract the `rgb` attribute. Store it on a new field passed back to the caller (or store on `WorksheetXml` temporarily — but since `SheetData` is the parent, pass it back via return type or add a `tab_color` field to `WorksheetXml`).

Simplest approach: add `tab_color: Option<String>` to `WorksheetXml`:

```rust
// In WorksheetXml struct:
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tab_color: Option<String>,
```

Parse in the worksheet reader:
```rust
// Inside parse(), when encountering <sheetPr>:
b"tabColor" => {
    if let Some(rgb) = e.try_get_attribute(b"rgb")? {
        worksheet.tab_color = Some(std::str::from_utf8(&rgb.value)?.to_string());
    }
}
```

Write in the worksheet writer:
```rust
// At start of to_xml() or to_xml_with_sst(), before <sheetData>:
if let Some(ref color) = self.tab_color {
    writer.create_element("sheetPr")
        .write_inner_content(|w| {
            w.create_element("tabColor")
                .with_attribute(("rgb", color.as_str()))
                .write_empty()?;
            Ok(())
        })?;
}
```

**Step 3b: Run Rust tests**

Run: `cargo test -p modern-xlsx-core`
Expected: PASS

**Step 3c: Add TypeScript properties**

In `packages/modern-xlsx/src/types.ts`, add to `WorksheetData`:

```typescript
  tabColor?: string | null;
```

In `packages/modern-xlsx/src/workbook.ts`, add to `Worksheet` class:

```typescript
  /** Returns the computed used range (e.g., "A1:D10") or null if the sheet is empty. */
  get usedRange(): string | null {
    const rows = this.data.worksheet.rows;
    if (rows.length === 0) return null;

    let minRow = Number.MAX_SAFE_INTEGER;
    let maxRow = 0;
    let minCol = Number.MAX_SAFE_INTEGER;
    let maxCol = 0;
    let hasCell = false;

    for (const row of rows) {
      for (const cell of row.cells) {
        const { row: r, col: c } = decodeCellRef(cell.reference);
        hasCell = true;
        if (r < minRow) minRow = r;
        if (r > maxRow) maxRow = r;
        if (c < minCol) minCol = c;
        if (c > maxCol) maxCol = c;
      }
    }

    if (!hasCell) return null;
    return `${columnToLetter(minCol)}${minRow + 1}:${columnToLetter(maxCol)}${maxRow + 1}`;
  }

  /** Returns the sheet tab color as an RGB hex string (e.g., "FF0000") or null. */
  get tabColor(): string | null {
    return this.data.worksheet.tabColor ?? null;
  }

  /** Sets the sheet tab color. Pass null to clear. */
  set tabColor(color: string | null) {
    this.data.worksheet.tabColor = color ?? undefined;
  }
```

**Step 3d: Rebuild WASM**

Run: `cd crates/modern-xlsx-wasm && wasm-pack build --target web --release --out-dir ../../packages/modern-xlsx/wasm --no-opt`

**Step 4: Run tests**

Run: `pnpm -C packages/modern-xlsx test -- __tests__/worksheet-extended.test.ts`
Expected: PASS

**Step 5: Commit**

```bash
git add crates/modern-xlsx-core/src/ooxml/worksheet.rs crates/modern-xlsx-core/src/lib.rs packages/modern-xlsx/src/workbook.ts packages/modern-xlsx/src/types.ts packages/modern-xlsx/__tests__/worksheet-extended.test.ts
git commit -m "feat: add usedRange and tabColor to Worksheet"
```

---

## Task 6: Cell Operations — `numberFormat`, `dateValue`, stub type

**Files:**
- Modify: `packages/modern-xlsx/src/workbook.ts`
- Modify: `packages/modern-xlsx/src/types.ts`
- Modify: `crates/modern-xlsx-core/src/ooxml/worksheet.rs`
- Create: `packages/modern-xlsx/__tests__/cell-extended.test.ts`

**Step 1: Write the failing tests**

```typescript
// packages/modern-xlsx/__tests__/cell-extended.test.ts
import { describe, expect, it } from 'vitest';
import { Workbook } from '../src/index.js';

describe('extended cell operations', () => {
  describe('cell.dateValue', () => {
    it('returns Date for date-formatted number', () => {
      const wb = new Workbook();
      const ws = wb.addSheet('S');

      // Apply a date style
      const styleIdx = wb.createStyle().numberFormat('yyyy-mm-dd').build(wb.styles);
      const cell = ws.cell('A1');
      cell.value = 46082; // 2026-03-01
      cell.styleIndex = styleIdx;

      const d = cell.dateValue;
      expect(d).toBeInstanceOf(Date);
      expect(d!.getUTCFullYear()).toBe(2026);
      expect(d!.getUTCMonth()).toBe(2); // March = 2
      expect(d!.getUTCDate()).toBe(1);
    });

    it('returns null for non-date cell', () => {
      const wb = new Workbook();
      const ws = wb.addSheet('S');
      ws.cell('A1').value = 42;
      expect(ws.cell('A1').dateValue).toBeNull();
    });

    it('returns null for string cell', () => {
      const wb = new Workbook();
      const ws = wb.addSheet('S');
      ws.cell('A1').value = 'hello';
      expect(ws.cell('A1').dateValue).toBeNull();
    });
  });

  describe('cell.numberFormat', () => {
    it('reads the number format from style', () => {
      const wb = new Workbook();
      const ws = wb.addSheet('S');
      const styleIdx = wb.createStyle().numberFormat('#,##0.00').build(wb.styles);
      const cell = ws.cell('A1');
      cell.value = 42;
      cell.styleIndex = styleIdx;

      expect(cell.numberFormat).toBe('#,##0.00');
    });

    it('returns null when no custom format', () => {
      const wb = new Workbook();
      const ws = wb.addSheet('S');
      ws.cell('A1').value = 42;
      expect(ws.cell('A1').numberFormat).toBeNull();
    });
  });
});
```

**Step 2: Run tests to verify they fail**

Run: `pnpm -C packages/modern-xlsx test -- __tests__/cell-extended.test.ts`
Expected: FAIL — `dateValue` and `numberFormat` do not exist on Cell

**Step 3a: Add stub to Rust CellType**

In `crates/modern-xlsx-core/src/ooxml/worksheet.rs`, add `Stub` variant to `CellType`:

```rust
pub enum CellType {
    SharedString,
    Number,
    Boolean,
    Error,
    FormulaStr,
    InlineStr,
    /// Explicitly empty cell (equivalent to SheetJS type "z").
    Stub,
}
```

Update `cell_type_json_str` to handle `Stub`:
```rust
CellType::Stub => "stub",
```

Update the XML parser — if a cell has no value and no formula, it can be treated as Stub. But this is only needed if writing; for reading, the default of `Number` with no value is fine.

**Step 3b: Add TypeScript types**

In `packages/modern-xlsx/src/types.ts`, update `CellType`:

```typescript
export type CellType = 'sharedString' | 'number' | 'boolean' | 'error' | 'formulaStr' | 'inlineStr' | 'stub';
```

**Step 3c: Add Cell properties**

In `packages/modern-xlsx/src/workbook.ts`, the Cell class needs access to styles to resolve number format. The Cell class currently only has `private readonly data: CellData`. We need to pass styles reference. Add an optional second constructor parameter:

```typescript
export class Cell {
  private readonly data: CellData;
  private readonly styles?: StylesData;

  constructor(data: CellData, styles?: StylesData) {
    this.data = data;
    this.styles = styles;
  }
```

Update `Worksheet.cell()` to pass styles:

```typescript
// In Worksheet class — cell() method needs access to workbook styles.
// The Worksheet already has this.data (SheetData). We need a reference to styles.
// Add a styles field to Worksheet:

export class Worksheet {
  private readonly data: SheetData;
  /** @internal */ styles?: StylesData;
  // ... existing code

  cell(ref: string): Cell {
    // ... existing cell lookup code ...
    return new Cell(cellData, this.styles);
  }
}
```

In `Workbook.addSheet()` and `Workbook.getSheet()`, set `ws.styles = this.data.styles`:

```typescript
addSheet(name: string): Worksheet {
  // ... existing validation and creation ...
  const ws = new Worksheet(sheetData);
  ws.styles = this.data.styles;
  return ws;
}

getSheet(name: string): Worksheet | undefined {
  // ... existing lookup ...
  const ws = new Worksheet(sheet);
  ws.styles = this.data.styles;
  return ws;
}
```

Add the new getters to `Cell`:

```typescript
  /**
   * Returns the resolved number format code for this cell (e.g., "#,##0.00"),
   * or null if using the default "General" format.
   */
  get numberFormat(): string | null {
    if (!this.styles || this.data.styleIndex == null) return null;
    const xf = this.styles.cellXfs?.[this.data.styleIndex];
    if (!xf?.numFmtId) return null;
    if (xf.numFmtId === 0) return null;

    // Check built-in formats
    const builtin = getBuiltinFormat(xf.numFmtId);
    if (builtin && builtin !== 'General') return builtin;

    // Check custom numFmts
    const custom = this.styles.numFmts?.find((nf) => nf.id === xf.numFmtId);
    return custom?.formatCode ?? null;
  }

  /**
   * Returns a Date object if this cell contains a date-formatted number, or null otherwise.
   * Uses the cell's number format to determine if the value represents a date.
   */
  get dateValue(): Date | null {
    if (this.data.cellType !== 'number' || this.data.value == null) return null;
    const fmt = this.numberFormat;
    if (!fmt || !isDateFormatCode(fmt)) return null;
    return serialToDate(Number.parseFloat(this.data.value));
  }
```

Add imports at top of workbook.ts:
```typescript
import { isDateFormatCode, serialToDate } from './dates.js';
import { getBuiltinFormat } from './format-cell.js';
```

**Step 4: Rebuild WASM (if Rust changed), run tests**

Run: `cargo test -p modern-xlsx-core` (for stub type)
Run WASM build if Rust changed.
Run: `pnpm -C packages/modern-xlsx test -- __tests__/cell-extended.test.ts`
Expected: PASS

**Step 5: Commit**

```bash
git add crates/modern-xlsx-core/src/ooxml/worksheet.rs packages/modern-xlsx/src/workbook.ts packages/modern-xlsx/src/types.ts packages/modern-xlsx/__tests__/cell-extended.test.ts
git commit -m "feat: add cell.dateValue, cell.numberFormat, and stub cell type"
```

---

## Task 7: Formulas — `dynamicArray` flag

**Files:**
- Modify: `crates/modern-xlsx-core/src/ooxml/worksheet.rs`
- Modify: `packages/modern-xlsx/src/types.ts`
- Create: `packages/modern-xlsx/__tests__/formula-extended.test.ts`

**Step 1: Write the failing test**

```typescript
// packages/modern-xlsx/__tests__/formula-extended.test.ts
import { describe, expect, it } from 'vitest';
import { readBuffer, Workbook } from '../src/index.js';

describe('dynamic array formulas', () => {
  it('roundtrips dynamicArray flag', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('S');
    ws.cell('A1').value = 1;
    ws.cell('A2').value = 2;

    // Set up the raw cell data to have dynamicArray flag
    const cellData = ws.rows[0]?.cells[0];
    if (cellData) {
      cellData.formula = 'SORT(A1:A2)';
      cellData.formulaType = 'array';
      cellData.dynamicArray = true;
    }

    const buf = await wb.toBuffer();
    const wb2 = await readBuffer(buf);
    const cell2 = wb2.getSheet('S')?.rows[0]?.cells[0];
    expect(cell2?.dynamicArray).toBe(true);
  });
});
```

**Step 2: Run tests to verify it fails**

Run: `pnpm -C packages/modern-xlsx test -- __tests__/formula-extended.test.ts`
Expected: FAIL — `dynamicArray` not in type

**Step 3: Implement**

In `crates/modern-xlsx-core/src/ooxml/worksheet.rs`, add to `Cell` struct:

```rust
    /// Whether this is a dynamic array formula (CSE/SPILL).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub dynamic_array: Option<bool>,
```

In the parser, when parsing `<f>` element attributes, check for `cm="1"`:

```rust
// When parsing <f> attributes:
if let Some(cm) = e.try_get_attribute(b"cm")? {
    if cm.value.as_ref() == b"1" {
        cell.dynamic_array = Some(true);
    }
}
```

In the writer, when emitting `<f>`, add `cm="1"` if `dynamic_array` is true:

```rust
if cell.dynamic_array == Some(true) {
    f_elem = f_elem.with_attribute(("cm", "1"));
}
```

In `packages/modern-xlsx/src/types.ts`, add to `CellData`:

```typescript
  dynamicArray?: boolean | null;
```

**Step 4: Rebuild WASM, run tests**

```bash
cargo test -p modern-xlsx-core
cd crates/modern-xlsx-wasm && wasm-pack build --target web --release --out-dir ../../packages/modern-xlsx/wasm --no-opt
pnpm -C packages/modern-xlsx test -- __tests__/formula-extended.test.ts
```
Expected: PASS

**Step 5: Commit**

```bash
git add crates/modern-xlsx-core/src/ooxml/worksheet.rs packages/modern-xlsx/src/types.ts packages/modern-xlsx/__tests__/formula-extended.test.ts
git commit -m "feat: add dynamicArray flag for SPILL formulas"
```

---

## Task 8: Update Feature Comparison Tests

**Files:**
- Modify: `packages/modern-xlsx/__tests__/feature-comparison.test.ts`

**Step 1: Update the exhaustive feature matrix**

Update the matrix counts and status icons in section 26 to reflect the new features. Change all flipped categories from Loss/Tie to Win. Update the FINAL SCORECARD test.

**Step 2: Run full test suite**

Run: `pnpm -C packages/modern-xlsx test`
Expected: ALL PASS

**Step 3: Run typecheck and lint**

Run: `pnpm -C packages/modern-xlsx typecheck && pnpm -C packages/modern-xlsx lint`
Expected: Clean

**Step 4: Commit**

```bash
git add packages/modern-xlsx/__tests__/feature-comparison.test.ts
git commit -m "feat: update feature comparison for v0.5.0 parity wins"
```

---

## Task 9: Version Bump, Changelog, Wiki, Push

**Files:**
- Modify: `packages/modern-xlsx/src/index.ts` (VERSION)
- Modify: `packages/modern-xlsx/package.json` (version)
- Modify: wiki: `Changelog.md`, `Feature-Comparison.md`, `Home.md`, `_Sidebar.md`, `API-Reference.md`
- Modify: `docs/FEATURE-COMPARISON.md`

**Step 1: Bump version**

In `packages/modern-xlsx/src/index.ts`:
```typescript
export const VERSION = '0.5.0' as const;
```

In `packages/modern-xlsx/package.json`:
```json
"version": "0.5.0"
```

**Step 2: Update changelog, wiki, docs**

Update all documentation to reflect new features.

**Step 3: Run full test suite one more time**

```bash
cargo test -p modern-xlsx-core
pnpm -C packages/modern-xlsx test
pnpm -C packages/modern-xlsx typecheck
pnpm -C packages/modern-xlsx lint
```

**Step 4: Commit and push**

```bash
git add -A
git commit -m "feat: v0.5.0 — Parity Sprint (beat SheetJS in 7 more categories)"
git push origin master
```

Push wiki separately:
```bash
cd ../modern-xlsx.wiki && git add -A && git commit -m "docs: update for v0.5.0" && git push origin master
```

---

## Summary

| Task | Description | Est. Lines | Depends On |
|------|-------------|-----------|------------|
| 1 | Cell ref utils (encodeRow, decodeRow, splitCellRef) | ~60 TS | — |
| 2 | Doc properties (appVersion, hyperlinkBase, revision) | ~80 Rust, ~20 TS | WASM build |
| 3 | Sheet conversions (sheetToTxt, sheetToFormulae) | ~50 TS | — |
| 4 | Number formatting (conditionals, colors, loadFormat) | ~200 TS | — |
| 5 | Worksheet ops (usedRange, tabColor) | ~60 TS, ~40 Rust | WASM build |
| 6 | Cell ops (dateValue, numberFormat, stub type) | ~80 TS, ~10 Rust | WASM build |
| 7 | Formulas (dynamicArray flag) | ~20 TS, ~20 Rust | WASM build |
| 8 | Update feature comparison tests | ~100 TS | Tasks 1-7 |
| 9 | Version bump, docs, push | docs | Task 8 |

Tasks 1, 3, 4 are pure TS — no WASM rebuild needed.
Tasks 2, 5, 6, 7 require Rust changes + WASM rebuild — batch the rebuild.
