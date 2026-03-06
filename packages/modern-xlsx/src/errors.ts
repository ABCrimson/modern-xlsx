/**
 * Standardized error class for the modern-xlsx public API.
 *
 * All public-facing error conditions throw `ModernXlsxError` with a
 * machine-readable `code` string for programmatic error handling.
 */
export class ModernXlsxError extends Error {
  readonly code: string;

  constructor(code: string, message: string) {
    super(message);
    this.name = 'ModernXlsxError';
    this.code = code;
  }
}

// ---------------------------------------------------------------------------
// Error codes
// ---------------------------------------------------------------------------

/** Invalid A1-style cell reference (e.g., empty string, malformed). */
export const INVALID_CELL_REF = 'INVALID_CELL_REF' as const;

/** WASM module failed to initialize or was not initialized before use. */
export const WASM_INIT_FAILED = 'WASM_INIT_FAILED' as const;

/** Sheet lookup by name or index found no matching sheet. */
export const SHEET_NOT_FOUND = 'SHEET_NOT_FOUND' as const;

/** Comment lookup by ID or cell reference found no matching comment. */
export const COMMENT_NOT_FOUND = 'COMMENT_NOT_FOUND' as const;

/** Generic invalid argument (wrong type, out of range, etc.). */
export const INVALID_ARGUMENT = 'INVALID_ARGUMENT' as const;
