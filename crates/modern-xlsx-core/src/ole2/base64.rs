//! Base64 over the RFC 4648 standard alphabet (`A-Z a-z 0-9 + /`, `=` padding),
//! hand-written because the crate has no base64 dependency.
//!
//! The `EncryptionInfo` Agile descriptor carries its salts, keys and verifiers
//! as base64 attributes. [`decode_base64`] parses them in every build; it is
//! lenient: it trims leading and trailing whitespace, skips `=`, space, CR and
//! LF anywhere in the input, drops trailing bits short of a whole byte, and
//! rejects any other byte outside the alphabet. [`encode_base64`] writes them,
//! and the sheet-protection `saltValue`/`hashValue` (ECMA-376 Part 1 §18.2.29),
//! and is built only with the `encryption` Cargo feature, so the lite build has
//! no dead code.

use core::hint::cold_path;

use crate::Result;
use crate::errors::ModernXlsxError;

/// Base64 decoder over the standard alphabet. Leading and trailing whitespace
/// is trimmed first. Skips `=`, space, CR and LF; any other byte outside the
/// alphabet is an error. Trailing bits short of a whole byte are dropped.
pub(crate) fn decode_base64(input: &str) -> Result<Vec<u8>> {
    const DECODE: [u8; 256] = {
        let mut table = [255u8; 256];
        let chars = b"ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/";
        let mut i = 0;
        while i < 64 {
            table[chars[i] as usize] = i as u8;
            i += 1;
        }
        table
    };

    let input = input.trim();
    let mut output = Vec::with_capacity(input.len() * 3 / 4);
    let mut buf = 0u32;
    let mut bits = 0u32;

    for &b in input.as_bytes() {
        if b == b'=' || b == b'\n' || b == b'\r' || b == b' ' {
            continue;
        }
        let val = DECODE[b as usize];
        if val == 255 {
            cold_path();
            return Err(ModernXlsxError::PasswordProtected(format!(
                "Invalid base64 character: {}",
                b as char
            )));
        }
        buf = (buf << 6) | u32::from(val);
        bits += 6;
        if bits >= 8 {
            bits -= 8;
            output.push((buf >> bits) as u8);
        }
    }
    Ok(output)
}

/// Base64 encode helper (inline implementation — no external dep).
#[cfg(feature = "encryption")]
pub(crate) fn encode_base64(data: &[u8]) -> String {
    const CHARS: &[u8; 64] = b"ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/";

    let mut result = String::with_capacity(data.len().div_ceil(3) * 4);

    for chunk in data.chunks(3) {
        let b0 = chunk[0] as u32;
        let b1 = chunk.get(1).copied().unwrap_or(0) as u32;
        let b2 = chunk.get(2).copied().unwrap_or(0) as u32;
        let triple = (b0 << 16) | (b1 << 8) | b2;

        result.push(CHARS[((triple >> 18) & 0x3F) as usize] as char);
        result.push(CHARS[((triple >> 12) & 0x3F) as usize] as char);

        match chunk.len() {
            1 => {
                result.push('=');
                result.push('=');
            }
            2 => {
                result.push(CHARS[((triple >> 6) & 0x3F) as usize] as char);
                result.push('=');
            }
            _ => {
                result.push(CHARS[((triple >> 6) & 0x3F) as usize] as char);
                result.push(CHARS[(triple & 0x3F) as usize] as char);
            }
        }
    }

    result
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn decode_known_values_and_leniency() {
        assert_eq!(decode_base64("").unwrap(), b"");
        assert_eq!(decode_base64("Zm9vYmFy").unwrap(), b"foobar");
        assert_eq!(decode_base64("Zm8=").unwrap(), b"fo");
        assert_eq!(decode_base64(" Zm9v\r\nYmFy ").unwrap(), b"foobar");
        // Trailing bits short of a whole byte are dropped.
        assert_eq!(decode_base64("Zm9vYg").unwrap(), b"foob");
    }

    #[test]
    fn decode_rejects_bytes_outside_the_alphabet() {
        let err = decode_base64("Zm9v*").unwrap_err();
        assert!(matches!(&err, ModernXlsxError::PasswordProtected(m)
            if m == "Invalid base64 character: *"));
    }

    #[cfg(feature = "encryption")]
    #[test]
    fn test_encode_base64_known_values() {
        // Known base64 values.
        assert_eq!(encode_base64(b""), "");
        assert_eq!(encode_base64(b"f"), "Zg==");
        assert_eq!(encode_base64(b"fo"), "Zm8=");
        assert_eq!(encode_base64(b"foo"), "Zm9v");
        assert_eq!(encode_base64(b"foobar"), "Zm9vYmFy");
        assert_eq!(encode_base64(b"Hello, World!"), "SGVsbG8sIFdvcmxkIQ==");

        // Binary input: the output uses only the base64 alphabet and padding.
        let binary = vec![0u8, 1, 2, 255, 254, 128];
        let encoded = encode_base64(&binary);
        assert!(
            encoded
                .chars()
                .all(|c| c.is_ascii_alphanumeric() || c == '+' || c == '/' || c == '='),
            "Base64 output should contain only valid characters"
        );
    }

    #[cfg(feature = "encryption")]
    #[test]
    fn roundtrip_lengths_0_to_64() {
        for len in 0..=64usize {
            let data: Vec<u8> = (0..len).map(|i| (i * 37 + 11) as u8).collect();
            assert_eq!(
                decode_base64(&encode_base64(&data)).unwrap(),
                data,
                "len {len}"
            );
        }
    }
}
