#![no_main]
//! Fuzz the encrypted-read path (CFB parse, EncryptionInfo, key derivation, decrypt, then the JSON read).
//! Property: `read_xlsx_json_with_password` must never panic on arbitrary input.
//!
//! Input layout: byte 0 modulo 33 is the password length, then the password bytes, then the file.
//! An empty password means "no password" to the reader and would duplicate `read_xlsx_json`, so it becomes `"x"`.
use libfuzzer_sys::fuzz_target;
use modern_xlsx_core::reader::read_xlsx_json_with_password;

fuzz_target!(|data: &[u8]| {
    let n = usize::from(data.first().copied().unwrap_or(0) % 33);
    let rest = data.get(1..).unwrap_or(&[]);
    let (pw, body) = rest.split_at(n.min(rest.len()));
    let pw = String::from_utf8_lossy(pw);
    let pw = if pw.is_empty() { "x".into() } else { pw };
    let _ = read_xlsx_json_with_password(body, &pw);
});
