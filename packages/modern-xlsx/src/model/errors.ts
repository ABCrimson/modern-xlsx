/**
 * Whether `value` is an `Error`, including errors created in another realm
 * (`node:vm` contexts, iframes) that fail `instanceof Error`. Falls back to
 * `instanceof` on engines without `Error.isError`. Internal: not exported from the entries.
 */
export function isError(value: unknown): value is Error {
  return typeof Error.isError === 'function' ? Error.isError(value) : value instanceof Error;
}

/**
 * Standardized error class for the modern-xlsx public API.
 *
 * Workbook, worksheet and cell operations, cell-reference parsing, WASM loading
 * and worker transport throw `ModernXlsxError` with a machine-readable `code`
 * string for programmatic error handling. Builder and encoder argument
 * validation (the barcode encoders, the chart, table and header/footer
 * builders, and `validateChartData`) throws `RangeError` or `Error` instead.
 *
 * Errors originating from the Rust/WASM core use the `[CODE] message` format
 * and are automatically parsed by {@link ModernXlsxError.fromWasmError}.
 */
export class ModernXlsxError extends Error {
  /** Machine-readable error code; see {@link ModernXlsxErrorCode}. */
  readonly code: ModernXlsxErrorCode;

  constructor(code: ModernXlsxErrorCode, message: string, options?: ErrorOptions) {
    super(message, options);
    this.name = 'ModernXlsxError';
    this.code = code;
  }

  /**
   * Parse a WASM error message in the `"[CODE] human message"` format.
   *
   * If the message matches the coded format, returns a `ModernXlsxError` with
   * the extracted code and message. Otherwise falls back to `WASM_ERROR`.
   * The original error (or thrown value) is kept as the `cause`.
   */
  static fromWasmError(err: unknown): ModernXlsxError {
    const msg = isError(err) ? err.message : String(err);
    const match = /^\[([A-Z_]+)\]\s*(.*)$/s.exec(msg);
    if (match && match[1] !== undefined && match[2] !== undefined) {
      return new ModernXlsxError(match[1], match[2], { cause: err });
    }
    return new ModernXlsxError(WASM_ERROR, msg, { cause: err });
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

// --- Codes forwarded from the Rust core (match ModernXlsxError::code()); IO_ERROR, INVALID_STYLE and INVALID_FORMAT are noted exceptions ---

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

/**
 * A cell value is invalid or missing, or a streaming writer sheet name is invalid or a
 * duplicate.
 */
export const INVALID_CELL_VALUE = 'INVALID_CELL_VALUE' as const;

/** Reserved: nothing raises it (the Rust core has no such variant); kept for compatibility. */
export const INVALID_STYLE = 'INVALID_STYLE' as const;

/** A date serial number could not be converted. */
export const INVALID_DATE = 'INVALID_DATE' as const;

/** Reserved: nothing raises it (the Rust core has no such variant); kept for compatibility. */
export const INVALID_FORMAT = 'INVALID_FORMAT' as const;

/** A required XLSX part is missing from the archive. */
export const MISSING_PART = 'MISSING_PART' as const;

/** A security check failed (ZIP bomb, path traversal, etc.). */
export const SECURITY = 'SECURITY' as const;

/**
 * A hard size limit was exceeded. On read: the JSON length, the total text a read copies from
 * shared strings, comment authors, hyperlink targets and worksheet namespace declarations into
 * the cells, comments, hyperlinks and kept worksheet extensions that name them, or the memory
 * of the records a read keeps in its lists: the defined names, the shared string entries and
 * runs, the records of `xl/styles.xml` (fonts, fills, borders, cell formats, cell styles,
 * differential formats, number formats and the like), the chart series and axes, the threaded
 * comments and mentions, then the worksheet lists (data validations, conditional-format rules,
 * selections, columns, hyperlinks, merged ranges and the like), counted together over all of
 * them and all sheets at each record's in-memory size plus its strings, except the plain text
 * of a shared string entry, which is not counted (all three are set by
 * `ReadOptions.limits.maxJsonLength`). Otherwise: an OLE2 stream size on encrypted write, the
 * pivot items the writer adds when it completes pivot tables built in code (more than
 * 1,048,576 items in one field, or more than 4,194,304 items and pivot fields added in one
 * workbook), the notes the writer generates in one workbook (more than
 * 131,072 VML note shapes for legacy comments and `tc=` placeholders for threaded comments,
 * `MAX_GENERATED_NOTE_SHAPES`; notes kept in the VML drawing they were read with are not
 * counted), or a `sheetToCsv`, `sheetToTxt` or `sheetToHtml` grid whose text would exceed
 * V8's maximum string length.
 */
export const SIZE_LIMIT = 'SIZE_LIMIT' as const;

/** The file is password-protected. */
export const PASSWORD_PROTECTED = 'PASSWORD_PROTECTED' as const;

/** The file is a legacy .xls format (not supported). */
export const LEGACY_FORMAT = 'LEGACY_FORMAT' as const;

/**
 * The bytes are not a readable workbook: neither a ZIP nor an encrypted or legacy-.xls OLE2 file
 * (another OLE2 document or a malformed OLE2 header lands here too).
 */
export const UNRECOGNIZED_FORMAT = 'UNRECOGNIZED_FORMAT' as const;

/** readFile or Workbook#toFile ran outside a Node.js-compatible runtime, or writeBlob could not construct the Blob (raised by the WASM bridge); never raised by the Rust core. */
export const IO_ERROR = 'IO_ERROR' as const;

/**
 * JSON crossing the WASM bridge (workbook data or a streaming row) is malformed or does
 * not match the expected shape. The value is `'JSON'`; the constant is named
 * `JSON_ERROR` so it never shadows the global `JSON`.
 */
export const JSON_ERROR = 'JSON' as const;

/**
 * The workbook cannot be written as given: it has no sheets (or a streaming writer finished
 * before any startSheet), two pivot caches share a cacheId, a pivot table cacheId names no
 * cache, a theme colour is not 6 hex digits, a colour string (font, fill, border, tab,
 * conditional-format, sparkline or rich-text colour) is not 6 or 8 hex digits (RRGGBB or
 * AARRGGBB, optional leading #), a conditional-formatting rule has more than three formulas
 * (formula plus additionalFormulas), two series of a chart share an `order` (within one
 * chart-type element or across its primary and secondary chart), or a sheet's `sheetProtection.password` is not
 * a 1 to 4 hex digit legacy hash (set `plainPassword` instead), or, in `modern-xlsx/lite`, a
 * `plainPassword` holds a non-ASCII character.
 */
export const INVALID_WORKBOOK = 'INVALID_WORKBOOK' as const;

/**
 * Web Worker transport failure: the worker terminated or crashed, a message could
 * not be deserialised, or a reply carried no data.
 */
export const WORKER_ERROR = 'WORKER_ERROR' as const;

/** Every error code this package defines. */
export type KnownModernXlsxErrorCode =
  | typeof INVALID_CELL_REF
  | typeof WASM_INIT_FAILED
  | typeof SHEET_NOT_FOUND
  | typeof COMMENT_NOT_FOUND
  | typeof INVALID_ARGUMENT
  | typeof WASM_ERROR
  | typeof ZIP_READ
  | typeof ZIP_WRITE
  | typeof ZIP_ENTRY
  | typeof ZIP_FINALIZE
  | typeof XML_PARSE
  | typeof XML_WRITE
  | typeof INVALID_CELL_VALUE
  | typeof INVALID_STYLE
  | typeof INVALID_DATE
  | typeof INVALID_FORMAT
  | typeof MISSING_PART
  | typeof SECURITY
  | typeof SIZE_LIMIT
  | typeof PASSWORD_PROTECTED
  | typeof LEGACY_FORMAT
  | typeof UNRECOGNIZED_FORMAT
  | typeof IO_ERROR
  | typeof JSON_ERROR
  | typeof INVALID_WORKBOOK
  | typeof WORKER_ERROR;

/**
 * The code of a {@link ModernXlsxError}. The Rust core may add codes, so any string is
 * accepted; the `(string & {})` arm keeps editor completion for the known codes.
 */
export type ModernXlsxErrorCode = KnownModernXlsxErrorCode | (string & {});
