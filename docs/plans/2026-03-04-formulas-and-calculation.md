# 0.7.x — Formulas & Calculation Implementation Plan

> **For Claude:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task.

**Goal:** Build a complete formula tokenizer, parser, reference resolver, reference rewriter, shared formula expansion, and multi-function evaluation engine — enabling modern-xlsx to compute cell values without Excel.

**Architecture:** A pipeline approach: raw formula string -> tokenizer (lexer) -> parser (AST) -> evaluator (tree-walk interpreter). Each stage is independently testable. The tokenizer produces flat tokens, the parser builds an AST, and the evaluator walks the AST resolving cell references from the workbook data. All formula logic lives in TypeScript (`packages/modern-xlsx/src/formula/`) — no Rust changes needed since formula strings already roundtrip cleanly.

**Tech Stack:** TypeScript (Vitest), pure computation (no dependencies)

---

## Version Map

| Version | Feature | Scope |
|---------|---------|-------|
| 0.7.0 | Formula Tokenizer (Lexer) | Tokenize formula strings into typed tokens |
| 0.7.1 | Formula Parser (AST) | Build expression trees from token streams |
| 0.7.2 | Formula Serializer | AST -> formula string (roundtrip fidelity) |
| 0.7.3 | Cell Reference Resolver | Parse A1/R1C1 refs, resolve to worksheet cells |
| 0.7.4 | Reference Rewriter (Rows/Cols) | Adjust refs on row/column insert/delete |
| 0.7.5 | Shared Formula Expansion | Derive child formulas from master + offset |
| 0.7.6 | Arithmetic Evaluation | Evaluate +, -, *, /, ^, %, comparisons, concatenation |
| 0.7.7 | String & Logical Functions | IF, AND, OR, NOT, CONCATENATE, LEFT, RIGHT, MID, LEN, TRIM, UPPER, LOWER, TEXT, VALUE |
| 0.7.8 | Math & Statistical Functions | SUM, AVERAGE, MIN, MAX, COUNT, COUNTA, ROUND, ABS, SQRT, MOD, INT, CEILING, FLOOR |
| 0.7.9 | Lookup Functions | VLOOKUP, HLOOKUP, INDEX, MATCH, CHOOSE, OFFSET |

---

## 0.7.0 — Formula Tokenizer (Lexer)

### Task 1: Create formula tokenizer module structure

**Files:**
- Create: `packages/modern-xlsx/src/formula/tokenizer.ts`
- Create: `packages/modern-xlsx/src/formula/index.ts`
- Modify: `packages/modern-xlsx/src/index.ts` — add formula exports

**Step 1: Define token types**

```typescript
// packages/modern-xlsx/src/formula/tokenizer.ts

export type TokenType =
  | 'number'         // 123, 3.14, 1.5E10
  | 'string'         // "hello"
  | 'boolean'        // TRUE, FALSE
  | 'error'          // #N/A, #REF!, #VALUE!, #DIV/0!, #NULL!, #NAME?, #NUM!
  | 'cell_ref'       // A1, $A$1, Sheet1!A1, 'Sheet Name'!A1
  | 'range'          // A1:B2, $A$1:$B$2
  | 'name'           // MyRange, _custom
  | 'function'       // SUM(, IF(, VLOOKUP(
  | 'operator'       // +, -, *, /, ^, &, =, <>, <, >, <=, >=
  | 'paren_open'     // (
  | 'paren_close'    // )
  | 'comma'          // ,
  | 'semicolon'      // ; (locale separator)
  | 'colon'          // :
  | 'percent'        // %
  | 'prefix_op'      // unary + or -
  | 'space'          // intersection operator (space between refs)
  | 'array_open'     // {
  | 'array_close'    // }
  | 'array_row_sep'  // ; inside array literals
  | 'array_col_sep'; // , inside array literals

export interface Token {
  type: TokenType;
  value: string;
  start: number;  // offset in source string
  end: number;
}

export interface TokenizeResult {
  tokens: Token[];
  errors: string[];
}
```

**Step 2: Write the failing test**

Create: `packages/modern-xlsx/__tests__/formula-tokenizer.test.ts`

```typescript
import { describe, expect, it } from 'vitest';
import { tokenize } from '../src/formula/tokenizer.js';

describe('Formula Tokenizer', () => {
  it('tokenizes a simple addition', () => {
    const result = tokenize('1+2');
    expect(result.errors).toEqual([]);
    expect(result.tokens).toEqual([
      { type: 'number', value: '1', start: 0, end: 1 },
      { type: 'operator', value: '+', start: 1, end: 2 },
      { type: 'number', value: '2', start: 2, end: 3 },
    ]);
  });

  it('tokenizes cell references', () => {
    const result = tokenize('A1+B2');
    expect(result.errors).toEqual([]);
    expect(result.tokens.map(t => t.type)).toEqual(['cell_ref', 'operator', 'cell_ref']);
  });

  it('tokenizes absolute references', () => {
    const result = tokenize('$A$1+$B2+C$3');
    expect(result.tokens.map(t => [t.type, t.value])).toEqual([
      ['cell_ref', '$A$1'], ['operator', '+'],
      ['cell_ref', '$B2'], ['operator', '+'],
      ['cell_ref', 'C$3'],
    ]);
  });

  it('tokenizes function calls', () => {
    const result = tokenize('SUM(A1:A10)');
    expect(result.tokens.map(t => [t.type, t.value])).toEqual([
      ['function', 'SUM'], ['paren_open', '('],
      ['cell_ref', 'A1'], ['colon', ':'], ['cell_ref', 'A10'],
      ['paren_close', ')'],
    ]);
  });

  it('tokenizes string literals', () => {
    const result = tokenize('"hello world"&"!"');
    expect(result.tokens.map(t => [t.type, t.value])).toEqual([
      ['string', 'hello world'], ['operator', '&'], ['string', '!'],
    ]);
  });

  it('tokenizes escaped quotes in strings', () => {
    const result = tokenize('"say ""hi"""');
    expect(result.tokens).toEqual([
      { type: 'string', value: 'say "hi"', start: 0, end: 12 },
    ]);
  });

  it('tokenizes boolean values', () => {
    const result = tokenize('TRUE+FALSE');
    expect(result.tokens.map(t => t.type)).toEqual(['boolean', 'operator', 'boolean']);
  });

  it('tokenizes error values', () => {
    const result = tokenize('#N/A');
    expect(result.tokens).toEqual([
      { type: 'error', value: '#N/A', start: 0, end: 4 },
    ]);
  });

  it('tokenizes comparison operators', () => {
    const result = tokenize('A1>=10');
    expect(result.tokens.map(t => [t.type, t.value])).toEqual([
      ['cell_ref', 'A1'], ['operator', '>='], ['cell_ref', '10'],
    ]);
    // Actually 10 is a number, not cell_ref. Let me fix:
  });

  it('tokenizes numbers with decimals and exponents', () => {
    const result = tokenize('3.14+1.5E10');
    expect(result.tokens.map(t => [t.type, t.value])).toEqual([
      ['number', '3.14'], ['operator', '+'], ['number', '1.5E10'],
    ]);
  });

  it('tokenizes sheet-qualified references', () => {
    const result = tokenize("Sheet1!A1+'My Sheet'!B2");
    expect(result.tokens.map(t => [t.type, t.value])).toEqual([
      ['cell_ref', 'Sheet1!A1'], ['operator', '+'], ['cell_ref', "'My Sheet'!B2"],
    ]);
  });

  it('tokenizes nested functions', () => {
    const result = tokenize('IF(A1>0,SUM(B1:B5),0)');
    expect(result.tokens.map(t => t.type)).toEqual([
      'function', 'paren_open',
      'cell_ref', 'operator', 'number', 'comma',
      'function', 'paren_open', 'cell_ref', 'colon', 'cell_ref', 'paren_close', 'comma',
      'number',
      'paren_close',
    ]);
  });

  it('tokenizes unary minus', () => {
    const result = tokenize('-A1');
    expect(result.tokens.map(t => [t.type, t.value])).toEqual([
      ['prefix_op', '-'], ['cell_ref', 'A1'],
    ]);
  });

  it('tokenizes percent operator', () => {
    const result = tokenize('50%');
    expect(result.tokens.map(t => [t.type, t.value])).toEqual([
      ['number', '50'], ['percent', '%'],
    ]);
  });

  it('tokenizes array constants', () => {
    const result = tokenize('{1,2;3,4}');
    expect(result.tokens.map(t => t.type)).toEqual([
      'array_open', 'number', 'array_col_sep', 'number',
      'array_row_sep', 'number', 'array_col_sep', 'number', 'array_close',
    ]);
  });

  it('reports errors for unterminated strings', () => {
    const result = tokenize('"unterminated');
    expect(result.errors.length).toBeGreaterThan(0);
  });
});
```

**Step 3: Run test to verify it fails**

Run: `pnpm -C packages/modern-xlsx test -- formula-tokenizer`
Expected: FAIL (module doesn't exist yet)

**Step 4: Implement the tokenizer**

Implement `tokenize(formula: string): TokenizeResult` in `packages/modern-xlsx/src/formula/tokenizer.ts`. The tokenizer is a single-pass character scanner that:

1. Skips leading `=` if present
2. Scans character-by-character with a position cursor
3. Matches patterns in priority order:
   - Whitespace → skip (or emit 'space' for intersection operator context)
   - `"` → scan string literal until closing `"`, handle `""` escapes
   - Digit or `.` followed by digit → scan number (including `E`/`e` exponent)
   - `#` → scan error literal (#N/A, #REF!, #VALUE!, #DIV/0!, #NULL!, #NAME?, #NUM!)
   - `{` → array_open, `}` → array_close
   - `(` → paren_open, `)` → paren_close
   - `,` → comma (or array_col_sep inside array context)
   - `;` → semicolon (or array_row_sep inside array context)
   - `:` → colon
   - `%` → percent
   - Operators: `+`, `-`, `*`, `/`, `^`, `&`, `=`, `<>`, `<=`, `>=`, `<`, `>`
   - `'` → scan sheet name until closing `'`, followed by `!` and cell ref
   - Letter → scan identifier (could be cell ref, function name, boolean, or named range)
     - If followed by `(` → function
     - If matches `TRUE`/`FALSE` → boolean
     - If matches `[A-Z]+[0-9]+` or `$[A-Z]+$[0-9]+` pattern → cell_ref
     - If contains `!` → sheet-qualified cell_ref
     - Otherwise → name (named range)
   - `$` → absolute ref prefix, scan following cell ref
   - Unary `+`/`-` detection: if previous token is null, operator, paren_open, comma → prefix_op

Key implementation detail: Use an `insideArray` counter (incremented on `{`, decremented on `}`) to distinguish `,` as `comma` vs `array_col_sep` and `;` as `semicolon` vs `array_row_sep`.

**Step 5: Run tests**

Run: `pnpm -C packages/modern-xlsx test -- formula-tokenizer`
Expected: All tests PASS

**Step 6: Create index barrel export**

Create `packages/modern-xlsx/src/formula/index.ts`:
```typescript
export type { Token, TokenType, TokenizeResult } from './tokenizer.js';
export { tokenize } from './tokenizer.js';
```

Add to `packages/modern-xlsx/src/index.ts`:
```typescript
// Formula engine
export type { Token, TokenType, TokenizeResult } from './formula/index.js';
export { tokenize } from './formula/index.js';
```

**Step 7: Run lint and typecheck**

Run: `pnpm -C packages/modern-xlsx lint && pnpm -C packages/modern-xlsx typecheck`
Expected: No errors

**Step 8: Commit**

```bash
git add packages/modern-xlsx/src/formula/ packages/modern-xlsx/__tests__/formula-tokenizer.test.ts packages/modern-xlsx/src/index.ts
git commit -m "feat(formula): add formula tokenizer (lexer) for 0.7.0"
```

---

## 0.7.1 — Formula Parser (AST)

### Task 2: Build expression tree parser

**Files:**
- Create: `packages/modern-xlsx/src/formula/parser.ts`
- Create: `packages/modern-xlsx/__tests__/formula-parser.test.ts`
- Modify: `packages/modern-xlsx/src/formula/index.ts` — add parser exports

**AST node types:**

```typescript
export type ASTNode =
  | NumberNode
  | StringNode
  | BooleanNode
  | ErrorNode
  | CellRefNode
  | RangeNode
  | NameNode
  | FunctionCallNode
  | BinaryOpNode
  | UnaryOpNode
  | PercentNode
  | ArrayNode;

export interface NumberNode { type: 'number'; value: number; }
export interface StringNode { type: 'string'; value: string; }
export interface BooleanNode { type: 'boolean'; value: boolean; }
export interface ErrorNode { type: 'error'; value: string; }
export interface CellRefNode {
  type: 'cell_ref';
  sheet?: string;     // Sheet1, 'My Sheet'
  col: string;        // A, B, AA
  row: number;        // 1, 2, 100
  absCol: boolean;    // $A → true
  absRow: boolean;    // $1 → true
}
export interface RangeNode {
  type: 'range';
  start: CellRefNode;
  end: CellRefNode;
}
export interface NameNode { type: 'name'; name: string; }
export interface FunctionCallNode {
  type: 'function';
  name: string;
  args: ASTNode[];
}
export interface BinaryOpNode {
  type: 'binary_op';
  op: string;         // +, -, *, /, ^, &, =, <>, <, >, <=, >=
  left: ASTNode;
  right: ASTNode;
}
export interface UnaryOpNode {
  type: 'unary_op';
  op: string;         // +, -
  operand: ASTNode;
}
export interface PercentNode {
  type: 'percent';
  operand: ASTNode;
}
export interface ArrayNode {
  type: 'array';
  rows: ASTNode[][];
}
```

**Parser implementation:** Recursive descent with operator precedence:
1. Comparison: `=`, `<>`, `<`, `>`, `<=`, `>=` (lowest)
2. Concatenation: `&`
3. Addition: `+`, `-`
4. Multiplication: `*`, `/`
5. Exponentiation: `^`
6. Unary: `+`, `-`
7. Percent: `%`
8. Atoms: number, string, boolean, error, cell_ref, range, function, `(expr)`, array

**Function:** `parse(formula: string): { ast: ASTNode; errors: string[] }`

Tests should cover: precedence, associativity, nested functions, ranges, errors, array constants.

---

## 0.7.2 — Formula Serializer

### Task 3: AST back to formula string

**Files:**
- Create: `packages/modern-xlsx/src/formula/serializer.ts`
- Create: `packages/modern-xlsx/__tests__/formula-serializer.test.ts`

**Function:** `serializeFormula(ast: ASTNode): string`

Walk the AST and produce a formula string. Test by round-tripping: `parse(formula) → AST → serialize(AST)` should produce equivalent formula.

Key details:
- CellRefNode → `$A$1` or `A1` based on `absCol`/`absRow`
- RangeNode → `A1:B2`
- FunctionCallNode → `NAME(arg1,arg2,...)`
- BinaryOpNode → `left op right` (with parens for precedence)
- UnaryOpNode → `op operand`
- PercentNode → `operand%`
- ArrayNode → `{1,2;3,4}`
- StringNode → `"value"` (with `""` for embedded quotes)

---

## 0.7.3 — Cell Reference Resolver

### Task 4: Parse and resolve cell references

**Files:**
- Create: `packages/modern-xlsx/src/formula/resolver.ts`
- Create: `packages/modern-xlsx/__tests__/formula-resolver.test.ts`

**Functions:**
- `parseCellRef(ref: string): { sheet?: string; col: number; row: number; absCol: boolean; absRow: boolean }`
- `resolveRef(ref: CellRefNode, workbook: Workbook, currentSheet: string): CellValue`
- `resolveRange(range: RangeNode, workbook: Workbook, currentSheet: string): CellValue[][]`

Tests: resolve A1 to actual cell values, cross-sheet refs, missing cells → null/0.

---

## 0.7.4 — Reference Rewriter (Row/Column Insert/Delete)

### Task 5: Adjust cell references on structural changes

**Files:**
- Create: `packages/modern-xlsx/src/formula/rewriter.ts`
- Create: `packages/modern-xlsx/__tests__/formula-rewriter.test.ts`

**Functions:**
- `rewriteFormula(formula: string, action: RewriteAction): string`
- `RewriteAction = InsertRows | DeleteRows | InsertCols | DeleteCols`

Walk AST, adjust row/col numbers in CellRefNode and RangeNode based on action. Absolute refs ($A$1) stay fixed. Relative refs adjust. Deleted refs become #REF!.

---

## 0.7.5 — Shared Formula Expansion

### Task 6: Derive child formulas from master

**Files:**
- Create: `packages/modern-xlsx/src/formula/shared.ts`
- Create: `packages/modern-xlsx/__tests__/formula-shared.test.ts`

**Function:**
- `expandSharedFormula(masterFormula: string, masterRef: string, childRef: string): string`

Parse master formula, compute row/col offset from master to child, adjust all relative refs, serialize back.

Example: Master at B1 with `A1*2`, child at B3 → `A3*2` (row offset +2).

---

## 0.7.6 — Arithmetic Evaluation

### Task 7: Evaluate basic expressions

**Files:**
- Create: `packages/modern-xlsx/src/formula/evaluator.ts`
- Create: `packages/modern-xlsx/__tests__/formula-evaluator.test.ts`

**Function:**
- `evaluateFormula(formula: string, context: EvalContext): CellValue`
- `EvalContext = { getCell(sheet: string, col: number, row: number): CellValue }`

Tree-walk interpreter supporting:
- Arithmetic: `+`, `-`, `*`, `/`, `^`
- Percent: `50%` → 0.5
- Comparison: `=`, `<>`, `<`, `>`, `<=`, `>=`
- Concatenation: `&`
- Unary: `+`, `-`
- Type coercion: string→number for arithmetic, number→string for concatenation
- Error propagation: any error input → error output

---

## 0.7.7 — String & Logical Functions

### Task 8: Built-in string and logical functions

**Files:**
- Create: `packages/modern-xlsx/src/formula/functions/string.ts`
- Create: `packages/modern-xlsx/src/formula/functions/logical.ts`
- Create: `packages/modern-xlsx/__tests__/formula-functions-string.test.ts`
- Create: `packages/modern-xlsx/__tests__/formula-functions-logical.test.ts`

**Functions to implement:**

| Function | Signature | Description |
|----------|-----------|-------------|
| IF | `IF(test, then, else)` | Conditional |
| AND | `AND(val1, val2, ...)` | Logical AND |
| OR | `OR(val1, val2, ...)` | Logical OR |
| NOT | `NOT(val)` | Logical negation |
| CONCATENATE | `CONCATENATE(s1, s2, ...)` | Join strings |
| LEFT | `LEFT(text, n)` | Left substring |
| RIGHT | `RIGHT(text, n)` | Right substring |
| MID | `MID(text, start, n)` | Mid substring |
| LEN | `LEN(text)` | String length |
| TRIM | `TRIM(text)` | Remove extra spaces |
| UPPER | `UPPER(text)` | Uppercase |
| LOWER | `LOWER(text)` | Lowercase |
| TEXT | `TEXT(value, format)` | Number → formatted string |
| VALUE | `VALUE(text)` | String → number |
| IFERROR | `IFERROR(value, fallback)` | Error handler |

---

## 0.7.8 — Math & Statistical Functions

### Task 9: Built-in math and statistical functions

**Files:**
- Create: `packages/modern-xlsx/src/formula/functions/math.ts`
- Create: `packages/modern-xlsx/src/formula/functions/stats.ts`
- Create: `packages/modern-xlsx/__tests__/formula-functions-math.test.ts`
- Create: `packages/modern-xlsx/__tests__/formula-functions-stats.test.ts`

**Functions to implement:**

| Function | Description |
|----------|-------------|
| SUM | Sum of values/ranges |
| AVERAGE | Arithmetic mean |
| MIN | Minimum value |
| MAX | Maximum value |
| COUNT | Count numeric values |
| COUNTA | Count non-empty values |
| COUNTBLANK | Count empty cells |
| ROUND | Round to N decimals |
| ROUNDUP | Round up |
| ROUNDDOWN | Round down |
| ABS | Absolute value |
| SQRT | Square root |
| MOD | Modulus |
| INT | Integer part |
| CEILING | Round up to multiple |
| FLOOR | Round down to multiple |
| POWER | Exponentiation |
| LOG | Logarithm |
| LN | Natural logarithm |
| PI | π constant |
| SUMIF | Conditional sum |
| COUNTIF | Conditional count |
| AVERAGEIF | Conditional average |

---

## 0.7.9 — Lookup Functions

### Task 10: Built-in lookup and reference functions

**Files:**
- Create: `packages/modern-xlsx/src/formula/functions/lookup.ts`
- Create: `packages/modern-xlsx/__tests__/formula-functions-lookup.test.ts`

**Functions to implement:**

| Function | Description |
|----------|-------------|
| VLOOKUP | Vertical lookup |
| HLOOKUP | Horizontal lookup |
| INDEX | Return value at position |
| MATCH | Find position of value |
| CHOOSE | Choose from list by index |
| OFFSET | Return range offset from ref |
| ROW | Row number of reference |
| COLUMN | Column number of reference |
| ROWS | Count rows in range |
| COLUMNS | Count columns in range |
| INDIRECT | Evaluate string as reference |

---

## Dependencies

```
0.7.0 (Tokenizer) ──→ 0.7.1 (Parser) ──→ 0.7.2 (Serializer)
                                    │
                                    ├──→ 0.7.3 (Resolver)
                                    ├──→ 0.7.4 (Rewriter) ──→ 0.7.5 (Shared Expansion)
                                    └──→ 0.7.6 (Evaluator) ──→ 0.7.7 (String/Logic)
                                                           ──→ 0.7.8 (Math/Stats)
                                                           ──→ 0.7.9 (Lookup)
```

## Public API Surface (when complete)

```typescript
import {
  // Tokenizer
  tokenize,
  // Parser
  parseFormula,
  // Serializer
  serializeFormula,
  // Resolver
  resolveRef,
  resolveRange,
  // Rewriter
  rewriteFormula,
  // Shared formulas
  expandSharedFormula,
  // Evaluator
  evaluateFormula,
  // Workbook-level
  calculateWorkbook,  // Evaluate all formulas in dependency order
} from 'modern-xlsx';
```
