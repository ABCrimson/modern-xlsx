//! JSON escaping, cell-type names and float formatting for the worksheet
//! JSON path (`parse_to_json_within`), which writes rows and cells as JSON text
//! itself instead of through serde.

use super::CellType;

/// Bytes of escape sequences gathered on the stack before one `push_str`
/// writes them to the output.
const ESCAPE_CHUNK: usize = 256;

/// For each byte, the character after the backslash of its JSON escape
/// (`u` for `\u00XX`), or 0 when the byte needs none: `"`, `\`, and the
/// control characters 0x00-0x1F (`\n`, `\r` and `\t` by their short forms).
const ESCAPE: [u8; 256] = {
    let mut table = [0_u8; 256];
    let mut b = 0;
    while b < 0x20 {
        table[b] = b'u';
        b += 1;
    }
    table[b'"' as usize] = b'"';
    table[b'\\' as usize] = b'\\';
    table[b'\n' as usize] = b'n';
    table[b'\r' as usize] = b'r';
    table[b'\t' as usize] = b't';
    table
};

/// Lowercase hex digits, for `\u00XX`.
const HEX: &[u8; 16] = b"0123456789abcdef";

/// The JSON escapes of the control characters 0x00-0x1F, as [`ESCAPE`]
/// spells them.
const CONTROL_ESCAPES: [&str; 32] = [
    "\\u0000", "\\u0001", "\\u0002", "\\u0003", "\\u0004", "\\u0005", "\\u0006", "\\u0007",
    "\\u0008", "\\t", "\\n", "\\u000b", "\\u000c", "\\r", "\\u000e", "\\u000f", "\\u0010",
    "\\u0011", "\\u0012", "\\u0013", "\\u0014", "\\u0015", "\\u0016", "\\u0017", "\\u0018",
    "\\u0019", "\\u001a", "\\u001b", "\\u001c", "\\u001d", "\\u001e", "\\u001f",
];

/// For each byte, its whole JSON escape as a static string, or `""` when it
/// needs none: the same escapes as [`ESCAPE`].
const ESCAPED: [&str; 256] = {
    let mut table = [""; 256];
    let mut b = 0;
    while b < 0x20 {
        table[b] = CONTROL_ESCAPES[b];
        b += 1;
    }
    table[b'"' as usize] = "\\\"";
    table[b'\\' as usize] = "\\\\";
    table
};

/// Escapes after a run of plain bytes that [`escape_from`] writes one by
/// one from [`ESCAPED`] before it gathers the rest of the run on the stack.
const STATIC_RUN: usize = 2;

/// Write a JSON-escaped string to the output buffer.
/// Handles `"`, `\`, and control characters (0x00-0x1F).
///
/// Scans bytes, not chars, for the first one that needs an escape: a string
/// without one, as most cell text is, goes out with one `push_str`, and the
/// rest go to [`escape_from`].
#[inline]
pub(crate) fn json_escape_to(out: &mut String, s: &str) {
    match s.bytes().position(|b| ESCAPE[usize::from(b)] != 0) {
        None => out.push_str(s),
        Some(first) => escape_from(out, s, first),
    }
}

/// The part of [`json_escape_to`] for a string whose first byte that needs
/// an escape is at `first`.
///
/// Copies each run of bytes that need no escape with one `push_str`. The
/// first [`STATIC_RUN`] escapes after such a run go out with one `push_str`
/// each of their static string from [`ESCAPED`], the fastest way to write
/// the single escapes and pairs (`\r\n`) of most text. The rest of a longer
/// run of escapes is looked up in [`ESCAPE`] and written into an
/// [`ESCAPE_CHUNK`]-byte stack buffer without branching on which escape it
/// is (all six bytes of `\u00XX` are stored, and only its length depends on
/// the byte); the buffer goes out with one `push_str` when it is full or a
/// run of plain bytes follows. So a string made of escapes in random order
/// (each shared-string copy of a 40,000-character entry of backslashes and
/// quotes, say) costs neither a mispredicted branch nor a `push_str` per
/// byte.
fn escape_from(out: &mut String, s: &str, first: usize) {
    let bytes = s.as_bytes();
    let mut chunk = [0_u8; ESCAPE_CHUNK];
    let mut n = 0;
    // Escapes since the last plain byte (or the start of `s`).
    let mut run = 0;
    // `start` and `i` only ever sit at or after ASCII bytes, so both are char boundaries and the slices cannot panic.
    let mut start = 0;
    for (i, &b) in bytes.iter().enumerate().skip(first) {
        let short = ESCAPE[usize::from(b)];
        if short == 0 {
            continue;
        }
        // The bytes before this escape, after the escapes gathered so far.
        if start < i {
            if n > 0 {
                push_ascii(out, &chunk[..n]);
                n = 0;
            }
            out.push_str(&s[start..i]);
            run = 0;
        }
        start = i + 1;
        if run < STATIC_RUN {
            run += 1;
            out.push_str(ESCAPED[usize::from(b)]);
            continue;
        }
        if n > ESCAPE_CHUNK - 6 {
            push_ascii(out, &chunk[..n]);
            n = 0;
        }
        chunk[n..n + 6].copy_from_slice(&[
            b'\\',
            short,
            b'0',
            b'0',
            HEX[usize::from(b >> 4)],
            HEX[usize::from(b & 0x0F)],
        ]);
        n += 2 + 4 * usize::from(short == b'u');
    }
    if n > 0 {
        push_ascii(out, &chunk[..n]);
    }
    if start < bytes.len() {
        out.push_str(&s[start..]);
    }
}

/// Append `ascii`, a run of escape sequences built by [`escape_from`], to
/// `out`.
#[inline]
fn push_ascii(out: &mut String, ascii: &[u8]) {
    match std::str::from_utf8(ascii) {
        Ok(text) => out.push_str(text),
        Err(_) => unreachable!("JSON escape sequences are ASCII"),
    }
}

/// Convert a `CellType` to its camelCase JSON string (matching serde rename).
#[inline]
pub(super) fn cell_type_json_str(ct: CellType) -> &'static str {
    match ct {
        CellType::SharedString => "sharedString",
        CellType::Number => "number",
        CellType::Boolean => "boolean",
        CellType::Error => "error",
        CellType::FormulaStr => "formulaStr",
        CellType::InlineStr => "inlineStr",
        CellType::Stub => "stub",
    }
}

/// Write an f64 to the JSON output buffer, byte-equal to `serde_json::to_string`:
/// both format with zmij's shortest round-trip `format_finite` (`15.0`, `-0.0`,
/// `1e+16`). Non-finite values, which JSON cannot carry, are written as `null`,
/// so a part holding `ht="INF"` still reads.
#[inline]
pub(super) fn write_f64_json(out: &mut String, v: f64) {
    if v.is_finite() {
        out.push_str(zmij::Buffer::new().format_finite(v));
    } else {
        out.push_str("null");
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use proptest::prelude::*;

    #[test]
    fn json_escape_matches_serde_json() {
        let long_escapes = "\"\\\u{1f}".repeat(200);
        let mixed = "é\u{0}a\"".repeat(100);
        for s in [
            "plain",
            "a\"b\\c",
            "é€😀\n\t\u{1}",
            "",
            long_escapes.as_str(),
            mixed.as_str(),
        ] {
            let mut out = String::from("\"");
            json_escape_to(&mut out, s);
            out.push('"');
            assert_eq!(out, serde_json::to_string(s).unwrap());
        }
    }

    /// The escape of `b` as JSON writers spell it here: short forms for
    /// `"`, `\`, `\n`, `\r` and `\t`, `\u00XX` for the other control
    /// characters, nothing for the rest.
    fn reference_escape(b: u8) -> String {
        match b {
            b'"' => "\\\"".to_owned(),
            b'\\' => "\\\\".to_owned(),
            b'\n' => "\\n".to_owned(),
            b'\r' => "\\r".to_owned(),
            b'\t' => "\\t".to_owned(),
            0x00..=0x1f => format!("\\u{b:04x}"),
            _ => String::new(),
        }
    }

    #[test]
    fn escape_tables_agree() {
        for b in 0..=255_u8 {
            let want = reference_escape(b);
            assert_eq!(ESCAPED[usize::from(b)], want, "byte {b:#x}");
            let short = ESCAPE[usize::from(b)];
            assert_eq!(short == 0, want.is_empty(), "byte {b:#x}");
            if short != 0 {
                assert_eq!(want.as_bytes()[1], short, "byte {b:#x}");
            }
        }
    }

    /// Runs of every escape kind, of each length across [`STATIC_RUN`] and
    /// the [`ESCAPE_CHUNK`] edge, between plain and multi-byte text.
    #[test]
    fn escape_runs_match_the_reference() {
        for esc in ["\"", "\\", "\n", "\u{1}", "\u{1f}", "\"\u{0}"] {
            for k in 0..300 {
                for (head, tail) in [("", ""), ("a", "é"), ("€", "b\"c"), ("😀", "\u{2}")] {
                    let s = format!("{head}{}{tail}", esc.repeat(k));
                    let want: String = s
                        .chars()
                        .map(|c| match u8::try_from(c) {
                            Ok(b) if b < 0x80 && !reference_escape(b).is_empty() => {
                                reference_escape(b)
                            }
                            _ => c.to_string(),
                        })
                        .collect();
                    let mut out = String::new();
                    json_escape_to(&mut out, &s);
                    assert_eq!(out, want, "{s:?}");
                }
            }
        }
    }

    fn f64_json(v: f64) -> String {
        let mut out = String::new();
        write_f64_json(&mut out, v);
        out
    }

    #[test]
    fn write_f64_json_keeps_negative_zero_and_nulls_non_finite() {
        assert_eq!(f64_json(-0.0), "-0.0");
        assert_eq!(f64_json(0.0), "0.0");
        assert_eq!(f64_json(15.0), "15.0");
        assert_eq!(f64_json(1e16), "1e+16");
        for v in [f64::INFINITY, f64::NEG_INFINITY, f64::NAN] {
            assert_eq!(f64_json(v), "null");
        }
    }

    proptest! {
        /// Every finite f64 is written byte-equal to serde_json.
        #[test]
        fn write_f64_json_matches_serde_json(
            v in proptest::num::f64::ANY.prop_filter("finite", |v| v.is_finite())
        ) {
            prop_assert_eq!(f64_json(v), serde_json::to_string(&v).unwrap());
        }
    }

    #[test]
    fn f64_json_matches_serde_json() {
        for v in [
            1e16,
            1.5e300,
            5e-324,
            0.1,
            -1.5,
            1e15,
            123456.789,
            -2.5e-8,
            f64::MAX,
            f64::MIN_POSITIVE,
        ] {
            let mut s = String::new();
            write_f64_json(&mut s, v);
            assert_eq!(s, serde_json::to_string(&v).unwrap(), "v = {v:e}");
        }
        let mut s = String::new();
        write_f64_json(&mut s, 1e16);
        assert_eq!(s, "1e+16");
    }
}
