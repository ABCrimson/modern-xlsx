# Sheet Management (0.4.6 + 0.4.7 + 0.4.8) — TDD Implementation Plan

> **For Claude:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task.

**Goal:** Expose sheet visibility state (hidden/veryHidden) through the full stack, add `moveSheet`/`cloneSheet`/`renameSheet` APIs, and verify tab color (already implemented).

**Architecture:** Add `state` field to `SheetData` (Rust lib.rs), wire through reader/writer, expose via TypeScript `Worksheet.state`. Sheet management APIs are TypeScript-only (they operate on the in-memory `WorkbookData` array). Tab color is already done — just needs integration test coverage.

**Tech Stack:** Rust 1.94 (quick-xml 0.39.2, serde camelCase), TypeScript 6.0, Vitest 4.1

---

## Task 1: Rust — Add `state` field to SheetData + wire reader/writer

**Files:**
- `crates/modern-xlsx-core/src/lib.rs`
- `crates/modern-xlsx-core/src/reader.rs`
- `crates/modern-xlsx-core/src/writer.rs`
- `crates/modern-xlsx-core/src/streaming.rs`
- `crates/modern-xlsx-core/src/validate.rs`
- `crates/modern-xlsx-core/tests/golden_tests.rs`
- `crates/modern-xlsx-core/examples/read_bench.rs`

### SheetData change (lib.rs):

Add `state` field to `SheetData`:

```rust
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct SheetData {
    pub name: String,
    /// Sheet visibility: "visible" (default), "hidden", or "veryHidden".
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub state: Option<String>,
    pub worksheet: WorksheetXml,
}
```

### Reader wiring (reader.rs):

In the standard reader path where `SheetData` is constructed from parsed sheets, set state from `SheetInfo`:

```rust
// In parse_sheets() or where SheetData is built:
// Convert SheetState enum to string for JSON bridge
fn sheet_state_to_str(state: SheetState) -> Option<String> {
    match state {
        SheetState::Visible => None, // default, omit
        SheetState::Hidden => Some("hidden".to_string()),
        SheetState::VeryHidden => Some("veryHidden".to_string()),
    }
}
```

Apply this when building `SheetData` from `sheet_targets` + `workbook_xml.sheets`.

In the streaming JSON reader path (`read_xlsx_json`), also emit the state field.

### Writer wiring (writer.rs):

Currently hardcodes `state: SheetState::Visible` on line 83. Change to read from `SheetData.state`:

```rust
fn str_to_sheet_state(s: &Option<String>) -> SheetState {
    match s.as_deref() {
        Some("hidden") => SheetState::Hidden,
        Some("veryHidden") => SheetState::VeryHidden,
        _ => SheetState::Visible,
    }
}
```

### All struct literals:

Add `state: None` to every `SheetData` struct literal in streaming.rs, validate.rs, golden_tests.rs, read_bench.rs.

### Tests: 3 Rust unit tests

1. **test_sheet_state_roundtrip_hidden**: Create WorkbookData with `state: Some("hidden".into())`, write, parse back workbook.xml, verify SheetInfo has `SheetState::Hidden`.
2. **test_sheet_state_roundtrip_very_hidden**: Same for "veryHidden".
3. **test_sheet_state_visible_omitted**: Verify `state: None` produces no `state` attribute in XML.

These can go in a new `#[cfg(test)]` block in `writer.rs` or as integration tests.

---

## Task 2: WASM rebuild + TypeScript types + API

### WASM rebuild:

```bash
cd crates/modern-xlsx-wasm && wasm-pack build --target web --release --out-dir ../../packages/modern-xlsx/wasm --no-opt
```

### TypeScript types (types.ts):

Add `state` field to `SheetData` interface:

```typescript
export interface SheetData {
  name: string;
  /** Sheet visibility: 'visible' (default/omitted), 'hidden', or 'veryHidden'. */
  state?: 'visible' | 'hidden' | 'veryHidden' | null;
  worksheet: WorksheetData;
}
```

Add `SheetState` type alias:

```typescript
export type SheetState = 'visible' | 'hidden' | 'veryHidden';
```

### Worksheet API (workbook.ts):

Add `state` getter/setter on the `Worksheet` class. The Worksheet class wraps a `SheetData` object (accessible as `this.data` — check the actual field name). Add:

```typescript
/** Returns the sheet visibility state. */
get state(): SheetState {
  return this.data.state ?? 'visible';
}

/** Sets the sheet visibility state. */
set state(value: SheetState) {
  this.data.state = value === 'visible' ? undefined : value;
}
```

### Workbook management APIs (workbook.ts):

Add to the `Workbook` class:

```typescript
/**
 * Moves a sheet from one position to another.
 * @param fromIndex 0-based source index.
 * @param toIndex 0-based destination index.
 */
moveSheet(fromIndex: number, toIndex: number): void {
  if (fromIndex < 0 || fromIndex >= this.data.sheets.length) {
    throw new Error(`Invalid source index: ${fromIndex}`);
  }
  if (toIndex < 0 || toIndex >= this.data.sheets.length) {
    throw new Error(`Invalid destination index: ${toIndex}`);
  }
  const [sheet] = this.data.sheets.splice(fromIndex, 1);
  this.data.sheets.splice(toIndex, 0, sheet);
}

/**
 * Clones a sheet and inserts the copy at the given position (default: end).
 * @param sourceIndex 0-based index of the sheet to clone.
 * @param newName Name for the cloned sheet (must be unique).
 * @param insertIndex Where to insert (default: end).
 * @returns The new Worksheet.
 */
cloneSheet(sourceIndex: number, newName: string, insertIndex?: number): Worksheet {
  if (sourceIndex < 0 || sourceIndex >= this.data.sheets.length) {
    throw new Error(`Invalid source index: ${sourceIndex}`);
  }
  validateSheetName(newName);
  if (this.data.sheets.some((s) => s.name === newName)) {
    throw new Error(`Sheet "${newName}" already exists`);
  }
  const source = this.data.sheets[sourceIndex];
  const clone: SheetData = structuredClone(source);
  clone.name = newName;
  const idx = insertIndex ?? this.data.sheets.length;
  this.data.sheets.splice(idx, 0, clone);
  return new Worksheet(clone, this.data.styles);
}

/**
 * Renames a sheet.
 * @param nameOrIndex Current name or 0-based index.
 * @param newName The new name (validated per ECMA-376).
 */
renameSheet(nameOrIndex: string | number, newName: string): void {
  const idx = typeof nameOrIndex === 'string'
    ? this.data.sheets.findIndex((s) => s.name === nameOrIndex)
    : nameOrIndex;
  if (idx < 0 || idx >= this.data.sheets.length) {
    throw new Error(`Sheet not found: ${nameOrIndex}`);
  }
  validateSheetName(newName);
  if (this.data.sheets.some((s, i) => s.name === newName && i !== idx)) {
    throw new Error(`Sheet "${newName}" already exists`);
  }
  this.data.sheets[idx].name = newName;
}

/**
 * Hides a sheet. At least one sheet must remain visible.
 * @param nameOrIndex Sheet name or 0-based index.
 */
hideSheet(nameOrIndex: string | number): void {
  const idx = typeof nameOrIndex === 'string'
    ? this.data.sheets.findIndex((s) => s.name === nameOrIndex)
    : nameOrIndex;
  if (idx < 0 || idx >= this.data.sheets.length) {
    throw new Error(`Sheet not found: ${nameOrIndex}`);
  }
  const visibleCount = this.data.sheets.filter(
    (s) => !s.state || s.state === 'visible',
  ).length;
  if (visibleCount <= 1) {
    throw new Error('Cannot hide the last visible sheet');
  }
  this.data.sheets[idx].state = 'hidden';
}

/**
 * Unhides a sheet (sets state to 'visible').
 * @param nameOrIndex Sheet name or 0-based index.
 */
unhideSheet(nameOrIndex: string | number): void {
  const idx = typeof nameOrIndex === 'string'
    ? this.data.sheets.findIndex((s) => s.name === nameOrIndex)
    : nameOrIndex;
  if (idx < 0 || idx >= this.data.sheets.length) {
    throw new Error(`Sheet not found: ${nameOrIndex}`);
  }
  this.data.sheets[idx].state = undefined;
}
```

### Export from index.ts:

Add `SheetState` to type exports.

---

## Task 3: TypeScript tests + verification

**File:** `packages/modern-xlsx/__tests__/sheet-management.test.ts`

### Tests (10):

```typescript
import { describe, expect, it } from 'vitest';
import { readBuffer, Workbook } from '../src/index.js';

describe('Sheet Management', () => {
  // --- Sheet State ---

  it('default sheet state is visible', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    expect(ws.state).toBe('visible');
  });

  it('set sheet state to hidden', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    wb.addSheet('Sheet2'); // need another visible sheet
    ws.state = 'hidden';
    expect(ws.state).toBe('hidden');
  });

  it('hidden sheet survives roundtrip', async () => {
    const wb = new Workbook();
    wb.addSheet('Sheet1').cell('A1').value = 'data';
    const ws2 = wb.addSheet('Sheet2');
    ws2.cell('A1').value = 'hidden';
    ws2.state = 'hidden';

    const buffer = await wb.toBuffer();
    const wb2 = await readBuffer(buffer);
    expect(wb2.getSheet('Sheet2')!.state).toBe('hidden');
    expect(wb2.getSheet('Sheet1')!.state).toBe('visible');
  });

  it('cannot hide last visible sheet', () => {
    const wb = new Workbook();
    wb.addSheet('Sheet1');
    expect(() => wb.hideSheet('Sheet1')).toThrow('last visible');
  });

  it('unhide restores visible state', () => {
    const wb = new Workbook();
    wb.addSheet('Sheet1');
    const ws2 = wb.addSheet('Sheet2');
    ws2.state = 'hidden';
    wb.unhideSheet('Sheet2');
    expect(wb.getSheet('Sheet2')!.state).toBe('visible');
  });

  // --- Move Sheet ---

  it('move sheet changes order', () => {
    const wb = new Workbook();
    wb.addSheet('A');
    wb.addSheet('B');
    wb.addSheet('C');
    wb.moveSheet(2, 0);
    expect(wb.sheetNames).toEqual(['C', 'A', 'B']);
  });

  // --- Clone Sheet ---

  it('clone sheet duplicates content', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Original');
    ws.cell('A1').value = 'hello';
    const clone = wb.cloneSheet(0, 'Copy');
    expect(clone.cell('A1').value).toBe('hello');
    expect(wb.sheetCount).toBe(2);
  });

  it('clone sheet survives roundtrip', async () => {
    const wb = new Workbook();
    wb.addSheet('Sheet1').cell('A1').value = 'original';
    wb.cloneSheet(0, 'Clone');

    const buffer = await wb.toBuffer();
    const wb2 = await readBuffer(buffer);
    expect(wb2.sheetCount).toBe(2);
    expect(wb2.getSheet('Clone')!.cell('A1').value).toBe('original');
  });

  // --- Rename Sheet ---

  it('rename sheet updates name', () => {
    const wb = new Workbook();
    wb.addSheet('OldName');
    wb.renameSheet('OldName', 'NewName');
    expect(wb.sheetNames).toEqual(['NewName']);
    expect(wb.getSheet('NewName')).toBeDefined();
  });

  it('rename rejects duplicate name', () => {
    const wb = new Workbook();
    wb.addSheet('A');
    wb.addSheet('B');
    expect(() => wb.renameSheet('A', 'B')).toThrow('already exists');
  });
});
```

### Full verification:

```bash
cargo test -p modern-xlsx-core
cargo clippy -p modern-xlsx-core -- -D warnings
pnpm -C packages/modern-xlsx lint
pnpm -C packages/modern-xlsx typecheck
pnpm -C packages/modern-xlsx build
pnpm -C packages/modern-xlsx test
```
