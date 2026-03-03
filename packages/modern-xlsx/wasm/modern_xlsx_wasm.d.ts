/* tslint:disable */
/* eslint-disable */

/**
 * Read an XLSX file and return parsed workbook data as a JSON string.
 *
 * Accepts a `Uint8Array` containing the raw `.xlsx` bytes.
 * Returns a JSON string representing the parsed workbook (sheets, styles, etc.).
 * The caller should use `JSON.parse()` on the JS side to deserialize.
 *
 * Uses `read_xlsx_json` which streams row/cell data directly as JSON during
 * XML parsing, avoiding millions of intermediate struct/String allocations.
 * This is critical for WASM performance where `memory.grow` calls are expensive.
 */
export function read(data: Uint8Array): string;

/**
 * Validate and auto-repair a workbook. Returns repaired workbook as JSON.
 *
 * Accepts a JSON string describing the workbook.
 * Returns a JSON object with `{ workbook, report, repairCount }`.
 */
export function repair(json: string): string;

/**
 * Validate a workbook and return a JSON report.
 *
 * Accepts a JSON string describing the workbook (same format as `write`).
 * Returns a JSON string containing the `ValidationReport`.
 */
export function validate(json: string): string;

/**
 * Get the library version.
 */
export function version(): string;

/**
 * Write XLSX file bytes from a JSON string describing the workbook.
 *
 * Accepts a JSON string (from `JSON.stringify()` on the JS side).
 * Returns a `Uint8Array` containing the resulting `.xlsx` bytes.
 *
 * Uses `serde_json::from_str()` for fast deserialization, matching the
 * JSON string approach used in `read()`.
 */
export function write(json: string): Uint8Array;

/**
 * Write XLSX and return as a `Blob` for browser download.
 *
 * Accepts a JSON string describing the workbook.
 * Returns a `Blob` with MIME type `application/vnd.openxmlformats-officedocument.spreadsheetml.sheet`.
 */
export function writeBlob(json: string): Blob;

export type InitInput = RequestInfo | URL | Response | BufferSource | WebAssembly.Module;

export interface InitOutput {
    readonly memory: WebAssembly.Memory;
    readonly read: (a: number, b: number, c: number) => void;
    readonly repair: (a: number, b: number, c: number) => void;
    readonly validate: (a: number, b: number, c: number) => void;
    readonly version: (a: number) => void;
    readonly write: (a: number, b: number, c: number) => void;
    readonly writeBlob: (a: number, b: number, c: number) => void;
    readonly __wbindgen_export: (a: number, b: number) => number;
    readonly __wbindgen_export2: (a: number, b: number, c: number, d: number) => number;
    readonly __wbindgen_export3: (a: number) => void;
    readonly __wbindgen_add_to_stack_pointer: (a: number) => number;
    readonly __wbindgen_export4: (a: number, b: number, c: number) => void;
}

export type SyncInitInput = BufferSource | WebAssembly.Module;

/**
 * Instantiates the given `module`, which can either be bytes or
 * a precompiled `WebAssembly.Module`.
 *
 * @param {{ module: SyncInitInput }} module - Passing `SyncInitInput` directly is deprecated.
 *
 * @returns {InitOutput}
 */
export function initSync(module: { module: SyncInitInput } | SyncInitInput): InitOutput;

/**
 * If `module_or_path` is {RequestInfo} or {URL}, makes a request and
 * for everything else, calls `WebAssembly.instantiate` directly.
 *
 * @param {{ module_or_path: InitInput | Promise<InitInput> }} module_or_path - Passing `InitInput` directly is deprecated.
 *
 * @returns {Promise<InitOutput>}
 */
export default function __wbg_init (module_or_path?: { module_or_path: InitInput | Promise<InitInput> } | InitInput | Promise<InitInput>): Promise<InitOutput>;
