/**
 * The reference rules shared by a defined name (`Workbook.addNamedRange`) and
 * a table display name (`Worksheet.addTable`): not `TRUE` or `FALSE`, not an
 * A1 cell reference inside the grid, and not an R1C1 reference or a name that
 * starts with one, case ignored. Excel 365 (build 20527, 2026-11-12) refuses
 * each of these in both kinds of name, except that it opens a file holding a
 * defined name that is an A1 cell reference (only its `Names.Add` refuses
 * one), so the Rust `validate.rs` twin `name_problem` does not report that
 * case for a defined name; `is_r1c1_name` is the twin of `isR1C1Name`. The
 * character rules, which differ between the two kinds, stay with each caller.
 * {@link nameKey} is the form both kinds, and sheet names, are compared in
 * for duplicates.
 *
 * @module
 */

import { columnIndexOrNull, MAX_COL, MAX_ROW } from '../format/cell-ref.js';
import type { DefinedNameData, TableDefinitionData } from './types.js';

/** A name of 1-3 letters and digits: a cell reference when inside the grid. */
const A1_NAME = /^([A-Za-z]{1,3})(\d+)$/;

/** A leading `R`, `C` or `RC` and the digits after it. */
const R1C1_PREFIX = /^(rc|r|c)(\d*)/i;

/**
 * The key two sheet, table or defined names are compared by: NFC, then lower
 * case with each final sigma (U+03C2) made U+03C3, then NFC again. Excel
 * takes names equal ignoring case and canonical equivalence (`café` with
 * U+00E9 and `cafe` + U+0301) as one name. `toLowerCase` picks the final
 * sigma by position (U+03A3 at the end of a word lower-cases to U+03C2,
 * elsewhere to U+03C3) and Excel takes the two as one letter, so the key
 * folds U+03C2 into U+03C3; U+03C2 has no decomposition and takes part in no
 * composition, so the fold does not disturb NFC. The second NFC is needed
 * because lower-casing can make a sequence that composes: `t` + U+0308
 * composes to U+1E97, which has no capital, so `T` + U+0308 stays decomposed
 * under NFC and lower-cases to `t` + U+0308; without the second NFC, `T` +
 * U+0308 and U+1E97 (or `t` + U+0308) would get different keys although they
 * differ only in case. A name of ASCII characters only gets its lower case
 * at once: NFC leaves ASCII text as it is and its lower case is ASCII again,
 * with no sigma, so the full key would give the same string.
 */
export function nameKey(name: string): string {
  for (let i = 0; i < name.length; i++) {
    if (name.charCodeAt(i) > 0x7f) {
      return name.normalize('NFC').toLowerCase().replaceAll('\u03C2', '\u03C3').normalize('NFC');
    }
  }
  return name.toLowerCase();
}

/** The keys of `T` whose values are names: strings, possibly missing. */
type NameField<T> = {
  [P in keyof T]-?: T[P] extends string | null | undefined ? P : never;
}[keyof T];

/** No positions: what {@link NameKeyIndex.positionsOf} gives for a key no item has. */
const NO_POSITIONS: readonly number[] = [];

/**
 * The {@link nameKey} of every name in one or more lists, kept between
 * duplicate checks so that a check computes the key of the name it checks
 * and no other. The lists are the model's own arrays, and the public API
 * hands them and their items out (`Workbook.toJSON()`, `namedRanges`,
 * `getNamedRange`, `Worksheet.tables`, `getTable`, the data a `Workbook` is
 * constructed from), so a name can change without the index being told.
 * {@link sync} therefore walks the lists before every lookup and compares
 * each item (by identity) and its name (by string equality) with what it saw
 * last, computing no key while they agree; a new item at the end gets its
 * key computed (or taken from {@link hint}), and any other difference
 * (an item removed, moved, replaced or renamed, a list replaced) rebuilds the
 * index, reusing the key of every name it already knew. A lookup thus
 * answers exactly as a scan of the lists would, at the cost of one identity
 * and one string comparison per stored item.
 */
export class NameKeyIndex<T extends object> {
  readonly #field: NameField<T>;
  #items: (T | undefined)[] = [];
  #names: (string | null | undefined)[] = [];
  #keys: (string | undefined)[] = [];
  #byKey = new Map<string, number[]>();
  #hint: { name: string; key: string } | undefined;
  /** The keys of the names the index knew before a rebuild, read while the rebuild runs. */
  #known = new Map<string, string>();

  /**
   * @param field - The property that holds an item's name; `null`, `undefined`,
   *   `''` and a name that is neither a string nor a String object are not
   *   indexed.
   */
  constructor(field: NameField<T>) {
    this.#field = field;
  }

  /** Records that `name` has the key `key`, so that {@link sync} need not compute it again. */
  hint(name: string, key: string): void {
    this.#hint = { name, key };
  }

  /**
   * Brings the index up to date with `lists`, read in order as one list
   * (a missing list holds no items). Returns the index.
   */
  sync(lists: readonly (readonly (T | undefined)[] | null | undefined)[]): this {
    const total = this.#agreeingLength(lists);
    if (total === this.#items.length) return this;
    // An item that differs (-1), or fewer items that all agree: the lists lost
    // items at their end. Either way the index starts again.
    if (total < this.#items.length) this.#reset();
    // Append what lies past the items the index holds (all of them after a reset).
    let position = 0;
    for (const list of lists) {
      if (!list) continue;
      for (let j = Math.max(0, this.#items.length - position); j < list.length; j++) {
        this.#append(list[j]);
      }
      position += list.length;
    }
    this.#known = new Map();
    return this;
  }

  /**
   * The number of items in `lists` when every item the index holds is still
   * at its position with its name, else -1.
   */
  #agreeingLength(lists: readonly (readonly (T | undefined)[] | null | undefined)[]): number {
    let total = 0;
    for (const list of lists) {
      if (!list) continue;
      if (!this.#agrees(list, total)) return -1;
      total += list.length;
    }
    return total;
  }

  /** `true` when the held items from position `from` on are the first items of `list`, with the same names. */
  #agrees(list: readonly (T | undefined)[], from: number): boolean {
    const items = this.#items;
    const names = this.#names;
    const field = this.#field;
    const stop = Math.min(list.length, items.length - from);
    for (let j = 0; j < stop; j++) {
      const item = list[j];
      if (
        item !== items[from + j] ||
        (item == null ? undefined : item[field]) !== names[from + j]
      ) {
        return false;
      }
    }
    return true;
  }

  /** The positions (in the order {@link sync} read the lists) of the items whose name has `key`. */
  positionsOf(key: string): readonly number[] {
    return this.#byKey.get(key) ?? NO_POSITIONS;
  }

  /** The item at `position`, as {@link sync} last read it. */
  itemAt(position: number): T | undefined {
    return this.#items[position];
  }

  /**
   * The first item, in list order, whose name has `key` and that `accept`
   * (when given) takes; `accept` gets the item and its position.
   */
  first(key: string, accept?: (item: T, position: number) => boolean): T | undefined {
    for (const position of this.positionsOf(key)) {
      const item = this.#items[position];
      if (item !== undefined && (!accept || accept(item, position))) return item;
    }
    return undefined;
  }

  #name(item: T | undefined): string | null | undefined {
    return item == null ? undefined : (item[this.#field] as string | null | undefined);
  }

  /** Empties the index, keeping the keys it knew (in `#known`) for the rebuild. */
  #reset(): void {
    const known = new Map<string, string>();
    for (let i = 0; i < this.#items.length; i++) {
      const name = this.#names[i];
      const key = this.#keys[i];
      if (name && key !== undefined) known.set(name, key);
    }
    this.#known = known;
    this.#items = [];
    this.#names = [];
    this.#keys = [];
    this.#byKey = new Map();
  }

  #append(item: T | undefined): void {
    const position = this.#items.length;
    const name = this.#name(item);
    // A String object (data built by hand) is keyed by its string value.
    const text = (name as unknown) instanceof String ? String(name) : name;
    let key: string | undefined;
    // Any other name that is not a string (data built by hand) is left out, as an empty one is.
    if (typeof text === 'string' && text !== '') {
      const hint = this.#hint;
      key = hint?.name === text ? hint.key : (this.#known.get(text) ?? nameKey(text));
      const positions = this.#byKey.get(key);
      if (positions) positions.push(position);
      else this.#byKey.set(key, [position]);
    }
    this.#items.push(item);
    this.#names.push(name);
    this.#keys.push(key);
  }
}

/**
 * The first reference rule `name` breaks, as a short reason, or `undefined`:
 * `TRUE` and `FALSE` are reserved, then {@link isCellReferenceName}, then
 * {@link isR1C1Name}.
 */
export function referenceNameProblem(name: string): string | undefined {
  const upper = name.toUpperCase();
  if (upper === 'TRUE' || upper === 'FALSE') return 'TRUE and FALSE are reserved';
  if (isCellReferenceName(name)) return 'it is a cell reference';
  if (isR1C1Name(name)) return 'it is an R1C1 reference or starts with one';
  return undefined;
}

/**
 * `true` for an A1 cell reference inside the grid, ignoring case: 1-3 letters
 * naming a column up to `XFD`, then a row 1-1048576 (leading zeros allowed).
 * `XFE1`, `A0` and `AAAA1` are outside the grid.
 */
export function isCellReferenceName(name: string): boolean {
  const a1 = A1_NAME.exec(name);
  if (!a1) return false;
  const row = Number(a1[2]);
  return columnIndexOrNull(a1[1] ?? '') !== null && row >= 1 && row <= MAX_ROW + 1;
}

/**
 * `true` for the names Excel reads as R1C1 references, ignoring case: `R`,
 * `C` and `RC` alone, and any name starting `R<row>`, `C<column>` or
 * `RC<column>` (row 1-1048576, column 1-16384) unless a `.` follows the number.
 * `R1x`, `C1_` and `R2D2` are refused; `R1.a`, `R0x` and `R1048577` are names.
 */
export function isR1C1Name(name: string): boolean {
  const match = R1C1_PREFIX.exec(name);
  if (!match) return false;
  const [prefix, letters = '', digits = ''] = match;
  if (digits === '') return prefix.length === name.length;
  const value = Number(digits);
  const max = letters.toLowerCase() === 'r' ? MAX_ROW + 1 : MAX_COL + 1;
  return value >= 1 && value <= max && name.charAt(prefix.length) !== '.';
}

/**
 * The name indexes of one workbook: tables by display name and by `name`
 * attribute (over the tables of every sheet, in sheet order), and defined
 * names by name.
 */
export interface WorkbookNameIndexes {
  readonly tables: NameKeyIndex<TableDefinitionData>;
  readonly tableNames: NameKeyIndex<TableDefinitionData>;
  readonly definedNames: NameKeyIndex<DefinedNameData>;
}

/** The indexes of each workbook data object (or standalone sheet), shared by its `Workbook` and `Worksheet` views. */
const WORKBOOK_INDEXES = new WeakMap<object, WorkbookNameIndexes>();

/**
 * The {@link WorkbookNameIndexes} of `owner`, the data object of a
 * workbook (or of a sheet outside one); each view of the same data shares
 * them, and they live as long as the data does.
 */
export function workbookNameIndexes(owner: object): WorkbookNameIndexes {
  let indexes = WORKBOOK_INDEXES.get(owner);
  if (!indexes) {
    indexes = {
      tables: new NameKeyIndex<TableDefinitionData>('displayName'),
      tableNames: new NameKeyIndex<TableDefinitionData>('name'),
      definedNames: new NameKeyIndex<DefinedNameData>('name'),
    };
    WORKBOOK_INDEXES.set(owner, indexes);
  }
  return indexes;
}
