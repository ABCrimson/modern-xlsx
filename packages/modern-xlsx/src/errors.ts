/**
 * Standardized error class for the modern-xlsx public API.
 *
 * All public-facing error conditions throw `ModernXlsxError` with a
 * machine-readable `code` string for programmatic error handling.
 *
 * Errors originating from the Rust/WASM core use the `[CODE] message` format
 * and are automatically parsed by {@link fromWasmError}.
 */
export class ModernXlsxError extends Error {
  readonly code: string;

  constructor(code: string, message: string) {
    super(message);
    this.name = 'ModernXlsxError';
    this.code = code;
  }

  /**
   * Parse a WASM error message in the `"[CODE] human message"` format.
   *
   * If the message matches the coded format, returns a `ModernXlsxError` with
   * the extracted code and message. Otherwise falls back to `WASM_ERROR`.
   */
  static fromWasmError(err: unknown): ModernXlsxError {
    const msg = err instanceof Error ? err.message : String(err);
    const match = /^\[([A-Z_]+)\]\s*(.*)$/.exec(msg);
    if (match && match[1] !== undefined && match[2] !== undefined) {
      return new ModernXlsxError(match[1], match[2]);
    }
    return new ModernXlsxError(WASM_ERROR, msg);
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

/** Unspecified WASM error (message did not contain a recognized code). */
export const WASM_ERROR = 'WASM_ERROR' as const;

// --- Codes forwarded from the Rust core (match ModernXlsxError::code()) ---

/** ZIP archive could not be read or decompressed. */
export const ZIP_READ = 'ZIP_READ' as const;

/** ZIP archive could not be written. */
export const ZIP_WRITE = 'ZIP_WRITE' as const;

/** A ZIP entry could not be accessed. */
export const ZIP_ENTRY = 'ZIP_ENTRY' as const;

/** ZIP archive finalization failed. */
export const ZIP_FINALIZE = 'ZIP_FINALIZE' as const;

/** XML parsing failed for an OOXML part. */
export const XML_PARSE = 'XML_PARSE' as const;

/** XML writing failed for an OOXML part. */
export const XML_WRITE = 'XML_WRITE' as const;

/** A cell value is invalid or missing. */
export const INVALID_CELL_VALUE = 'INVALID_CELL_VALUE' as const;

/** A style index or definition is invalid. */
export const INVALID_STYLE = 'INVALID_STYLE' as const;

/** A date serial number could not be converted. */
export const INVALID_DATE = 'INVALID_DATE' as const;

/** A number format string is invalid. */
export const INVALID_FORMAT = 'INVALID_FORMAT' as const;

/** A required XLSX part is missing from the archive. */
export const MISSING_PART = 'MISSING_PART' as const;

/** A security check failed (ZIP bomb, path traversal, etc.). */
export const SECURITY = 'SECURITY' as const;

/** The file is password-protected. */
export const PASSWORD_PROTECTED = 'PASSWORD_PROTECTED' as const;

/** The file is a legacy .xls format (not supported). */
export const LEGACY_FORMAT = 'LEGACY_FORMAT' as const;

/** The file format is unrecognized (not ZIP or OLE2). */
export const UNRECOGNIZED_FORMAT = 'UNRECOGNIZED_FORMAT' as const;

/** An I/O error occurred. */
export const IO_ERROR = 'IO_ERROR' as const;
