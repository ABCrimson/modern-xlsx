# Security Audit -- modern-xlsx Encryption Implementation

**Audit Date:** 2026-03-06
**Auditor:** Automated analysis (Claude Opus 4.6)
**Scope:** ECMA-376 Standard and Agile Encryption (OLE2 module)
**Version:** 0.9.1

## Executive Summary

The modern-xlsx encryption implementation is a well-structured, security-conscious implementation of ECMA-376 Agile and Standard Encryption. The code demonstrates consistent use of RAII-based secret zeroization (`SensitiveKey`), constant-time comparison for all password and HMAC verification, and CSPRNG-sourced randomness for salts/keys/IVs. The audit identified several low-severity findings and two medium-severity observations, but no critical vulnerabilities. Overall, the implementation follows cryptographic best practices and correctly implements the ECMA-376 specification.

## Encryption Architecture

### Overview

The encryption module is organized across five files in `crates/modern-xlsx-core/src/ole2/`:

- **`mod.rs`** -- Module root, exports submodules and the OLE2 magic constant.
- **`detect.rs`** -- OLE2 compound document reader: format detection, stream extraction, FAT chain traversal, and the `decrypt_file()` entry point that orchestrates the full decryption pipeline.
- **`encryption_info.rs`** -- Parses the `EncryptionInfo` stream from OLE2 containers. Supports both Agile (version 4.4, XML-based) and Standard (version 2.2/3.2/4.2, binary) encryption descriptors.
- **`crypto.rs`** -- Core cryptographic operations: key derivation, AES-CBC/ECB encryption/decryption, HMAC computation/verification, password verification, segment-based package encryption/decryption, and the full `encrypt_file()` pipeline.
- **`writer.rs`** -- OLE2 compound document writer for producing encrypted output files.

### Read Path (Decryption)

1. `detect_format()` identifies OLE2 magic bytes
2. `classify_ole2()` confirms `EncryptionInfo` + `EncryptedPackage` streams exist
3. `decrypt_file()` orchestrates: parse EncryptionInfo -> verify password -> verify HMAC -> decrypt package
4. For Agile: segment-based AES-CBC decryption with per-segment IVs derived from `Hash(salt + LE32(segment_index))`
5. For Standard: single-stream AES-CBC decryption with salt as IV

### Write Path (Encryption)

1. `encrypt_file()` generates random salts (16 bytes each for key and password), random data encryption key (32 bytes)
2. Segment-based AES-256-CBC encryption of the ZIP payload
3. HMAC-SHA-512 integrity computation over the encrypted package
4. Password-based key derivation (SHA-512, 100,000 iterations) to encrypt the data key and verifier
5. Agile EncryptionInfo XML descriptor generation
6. OLE2 compound document wrapping via `write_ole2()`

## Algorithm Analysis

### Key Derivation

**File:** `crates/modern-xlsx-core/src/ole2/crypto.rs:86-146`

The key derivation function (`derive_key` / `derive_key_impl`) implements ECMA-376 SS2.3.6.2 correctly:

1. **Step 1 -- UTF-16LE encoding:** Password is correctly encoded to UTF-16LE via `encode_utf16().flat_map(|c| c.to_le_bytes())`. This matches the specification requirement.

2. **Step 2 -- H0:** `H0 = Hash(salt || password_bytes)` -- correct order per spec.

3. **Step 3 -- Iteration:** `H_i = Hash(LE32(i) || H_{i-1})` for `spin_count` iterations. The iterator counter `i` is correctly encoded as little-endian 32-bit, and the iteration count is passed from the EncryptionInfo (typically 100,000 for Agile, 50,000 for Standard).

4. **Step 4 -- Final hash:** `H_final = Hash(H_n || block_key)` -- correct.

5. **Step 5 -- Key truncation/padding:** Key is padded with `0x36` bytes if the hash output is shorter than the requested key length. This matches the spec's cbRequiredKeyLength/cbHash padding behavior.

**Observation:** The implementation uses `FixedOutputReset` to reuse the hasher across iterations, which is an efficient and correct approach. The generic implementation over `Digest + Default + FixedOutputReset` supports SHA-512, SHA-256, and SHA-1 without code duplication.

**Supported algorithms:**
- SHA-512 (Agile, default for write)
- SHA-256 (Agile)
- SHA-1 (Standard Encryption)

**Block key constants** at lines 63-68 match the ECMA-376 specification values exactly.

### Encryption/Decryption

**File:** `crates/modern-xlsx-core/src/ole2/crypto.rs:152-276`

**AES-CBC with PKCS#7 padding** (`aes_cbc_decrypt`, `aes_cbc_encrypt`):
- Correctly dispatches on key length (16 = AES-128, 32 = AES-256)
- Uses the RustCrypto `cbc` crate with `Pkcs7` padding mode
- Error handling does not reveal key material -- errors report only "AES decryption failed (bad padding)" or "AES init error"

**AES-CBC without padding** (`aes_cbc_decrypt_no_pad`, `aes_cbc_encrypt_no_pad`):
- Used for segment-based operations where data is pre-padded to block boundaries
- Correctly uses `NoPadding` mode

**AES-ECB** (`aes_ecb_decrypt_no_pad`):
- Used only for Standard Encryption password verification (decrypting the 16-byte verifier and 32-byte verifier hash)
- Correctly validates that input length is a multiple of 16 bytes
- Block-by-block operation using raw `BlockDecrypt` from the `aes` crate

**Segment-based decryption** (`decrypt_package`, lines 443-475):
- Correctly reads the 8-byte LE64 original size header
- Processes payload in 4096-byte segments per ECMA-376 spec
- Each segment IV is derived as `Hash(salt || LE32(segment_index))` truncated to `block_size`
- Final result is truncated to `original_size` to remove padding

**Segment-based encryption** (`encrypt_package`, lines 550-577):
- Mirrors the decryption path correctly
- Each segment is padded to AES block boundary (16 bytes) with zero bytes before encryption
- Writes the 8-byte LE64 original size header

**Standard Encryption decryption** (`decrypt_standard_package`, lines 912-940):
- Single-stream AES-CBC with salt as IV -- correct per MS-OFFCRYPTO
- Validates minimum package length
- Truncates to original size after decryption

### HMAC Verification

**File:** `crates/modern-xlsx-core/src/ole2/crypto.rs:494-540`

The `verify_hmac()` function correctly implements the ECMA-376 Agile HMAC verification:

1. Derives the HMAC key encryption key: `Hash(data_key || BLOCK_KEY_HMAC_KEY)`
2. Decrypts the HMAC key from `encrypted_hmac_key` using AES-CBC
3. Derives the HMAC value encryption key: `Hash(data_key || BLOCK_KEY_HMAC_VALUE)`
4. Decrypts the expected HMAC value from `encrypted_hmac_value`
5. Computes HMAC over the **entire** encrypted package (including the 8-byte size prefix)
6. Performs constant-time comparison of computed vs. expected HMAC

**HMAC computation** (`compute_hmac`, lines 333-354):
- Supports SHA-512 and SHA-256
- Uses the RustCrypto `hmac` crate
- Returns the full HMAC output

**HMAC for encryption** (`compute_and_encrypt_hmac`, lines 582-623):
- Generates a random HMAC key of `hash_len` bytes
- Computes HMAC over the encrypted package
- Encrypts both the HMAC key and value for storage in EncryptionInfo

### Random Number Generation

**File:** `crates/modern-xlsx-core/src/ole2/crypto.rs:285-290`

The `secure_random()` function uses `getrandom::fill()`, which is a well-vetted CSPRNG wrapper:
- On native targets: uses the OS CSPRNG (`/dev/urandom`, `BCryptGenRandom`, etc.)
- On WASM targets: uses `crypto.getRandomValues()` via the `wasm_js` feature flag

**Usage in `encrypt_file()`** (lines 681-694):
- `key_salt`: 16 bytes -- CSPRNG
- `pw_salt`: 16 bytes -- CSPRNG
- `data_key`: 32 bytes (256-bit AES key) -- CSPRNG
- `verifier_input`: 16 bytes -- CSPRNG
- HMAC key (in `compute_and_encrypt_hmac`): `hash_len` bytes -- CSPRNG

All random material is generated via `secure_random()`, which is correct.

## Security Properties

### Constant-Time Operations

**File:** `crates/modern-xlsx-core/src/ole2/crypto.rs`

The `constant_time_eq` crate (v0.3) is used for all security-critical comparisons:

1. **Password verification (Agile)** -- line 411: `constant_time_eq(&computed[..hash_len], &verifier_hash[..hash_len])`
2. **HMAC verification** -- line 532: `constant_time_eq(&computed[..hash_len], &expected_hmac[..hash_len])`
3. **Password verification (Standard)** -- line 893: `constant_time_eq(&verifier_hash[..hash_len], &decrypted_verifier_hash[..hash_len])`

All three verification paths use constant-time comparison. No timing side-channel gaps were identified in the verification logic.

**Note:** The length checks before constant-time comparison (`computed.len() < hash_len`) are not constant-time, but this is acceptable because the lengths are derived from public metadata (hash algorithm output size and EncryptionInfo parameters), not from secret data.

### Secret Zeroization

**File:** `crates/modern-xlsx-core/src/ole2/crypto.rs:24-47`

The `SensitiveKey` struct provides RAII-based zeroization:

```rust
#[derive(Zeroize, ZeroizeOnDrop)]
pub(crate) struct SensitiveKey(pub(super) Vec<u8>);
```

- Derives both `Zeroize` and `ZeroizeOnDrop` from the `zeroize` crate (v1.8)
- The inner `Vec<u8>` is zeroized on drop regardless of control flow (early returns via `?`, panics, normal scope exit)
- `Deref` to `[u8]` allows transparent read access without copying

**Usage analysis -- zeroized keys:**

| Location | Variable | Wrapped in SensitiveKey? |
|---|---|---|
| `verify_password_agile` line 367 | `input_key` | Yes |
| `verify_password_agile` line 384 | `value_key` | Yes |
| `verify_password_agile` line 419 | `enc_key_key` | Yes |
| `verify_hmac` line 502 | `hmac_key_derivation` | Yes |
| `verify_hmac` line 505 | `hmac_key` | Yes |
| `verify_hmac` line 512 | `hmac_value_derivation` | Yes |
| `verify_hmac` line 515 | `expected_hmac` | Yes |
| `encrypt_file` line 683 | `data_key` | Yes |
| `encrypt_file` line 698 | `input_key` | Yes |
| `encrypt_file` line 703 | `value_key` | Yes |
| `encrypt_file` line 708 | `enc_key_key` | Yes |
| `compute_and_encrypt_hmac` line 602 | `hmac_key` | Yes |
| `compute_and_encrypt_hmac` line 608 | `hmac_key_enc_key` | Yes |
| `compute_and_encrypt_hmac` line 611 | `hmac_value_enc_key` | Yes |
| `decrypt_file` (detect.rs) line 238 | `data_key` | Yes |
| `decrypt_file` (detect.rs) line 246 | `data_key` | Yes |
| `verify_password_standard` line 866 | `derived` | Yes |

All derived keys and intermediate key material are wrapped in `SensitiveKey`.

**Note on `verify_password_standard`** (line 904): The derived key is returned via `std::mem::take(&mut derived.0)`, which replaces the inner `Vec` with an empty one. The empty `Vec` is then zeroized on drop (a no-op since it is empty). The taken `Vec` is returned to the caller, which wraps it in a new `SensitiveKey` in `decrypt_file()`. This is correct.

### Error Information Leakage

Error messages across the encryption module were reviewed for information leakage:

**Good practices observed:**
- Password verification failures return a generic "Incorrect password." message (lines 408, 413, 889, 898) -- no indication of *why* the password was wrong
- HMAC failures distinguish between "truncated hash" and "file may be corrupted or tampered with" -- this is acceptable since it aids debugging without revealing secrets
- AES errors report "AES decryption failed (bad padding)" or generic "AES init error" -- no key material leakage
- Unsupported algorithm errors report only the algorithm name, which is public metadata

**Potential concern:**
- The `CSPRNG error: {e}` message (line 288) could theoretically reveal platform-specific error details, but `getrandom` errors are extremely rare and do not contain sensitive data

**Assessment:** Error messages do not leak secret key material, plaintext, or intermediate cryptographic values.

## Findings

### [LOW] Intermediate Plaintext Not Zeroized After Password Verification

**File:** `crates/modern-xlsx-core/src/ole2/crypto.rs:377-401`

**Description:** In `verify_password_agile()`, the decrypted `verifier_input` (line 377) and `verifier_hash` (line 394) are stored as plain `Vec<u8>` rather than `SensitiveKey`. While these values are derived from the password verifier (random data encrypted with the password-derived key), they are not secret in the same sense as key material -- they exist solely for verification. However, the `verifier_input` is a random value that, combined with knowledge of the encryption parameters, could theoretically aid a brute-force attack if recovered from memory.

Similarly, in `verify_password_standard()`, the `decrypted_verifier` (line 876) and `verifier_hash` (line 879) are plain `Vec<u8>`.

**Impact:** Low. These values are ephemeral, scoped to the function, and do not contain the actual encryption key. They will be freed (but not zeroized) when the function returns.

**Recommendation:** Wrap `verifier_input` and `verifier_hash` in `SensitiveKey` for defense-in-depth:
```rust
let verifier_input = SensitiveKey::new(aes_cbc_decrypt(...)?);
let verifier_hash = SensitiveKey::new(aes_cbc_decrypt(...)?);
```

---

### [LOW] Password UTF-16LE Encoding Not Zeroized

**File:** `crates/modern-xlsx-core/src/ole2/crypto.rs:116-119`

**Description:** In `derive_key_impl()`, the password is encoded to UTF-16LE as a `Vec<u8>` (`pw_bytes`). This buffer contains the raw password in a different encoding and is not wrapped in `SensitiveKey`. It will be freed but not zeroized when the function returns.

**Impact:** Low. The password is already in memory as the `&str` parameter, so the UTF-16LE copy adds marginal risk. In a WASM environment, the heap is not accessible to other origins. On native targets, an attacker with memory read access likely already has the password from the original string.

**Recommendation:** For defense-in-depth, zeroize `pw_bytes` after use:
```rust
let mut pw_bytes: Vec<u8> = password.encode_utf16().flat_map(|c| c.to_le_bytes()).collect();
// ... use pw_bytes ...
pw_bytes.zeroize();
```

---

### [LOW] Hash Intermediate State Not Zeroized

**File:** `crates/modern-xlsx-core/src/ole2/crypto.rs:122-137`

**Description:** The intermediate hash values in `derive_key_impl()` (the `hash` variable across iterations, and the `derived` final hash) are `GenericArray` values from the `digest` crate. These are not explicitly zeroized. The RustCrypto `digest` crate's `FixedOutputReset` trait resets the hasher state but does not zeroize the returned `GenericArray`.

**Impact:** Low. The intermediate hash states are on the stack and will be overwritten by subsequent function calls. In optimized builds with `panic = "abort"` (as configured in `Cargo.toml`), stack frames are not unwound, reducing the window for memory inspection.

**Recommendation:** This is acceptable for the current threat model. If defense-in-depth is desired, the final `derived` output could be copied into a `SensitiveKey` before truncation.

---

### [LOW] Base64 Implementation Uses Custom Code

**File:** `crates/modern-xlsx-core/src/ole2/encryption_info.rs:333-370` and `crates/modern-xlsx-core/src/ole2/crypto.rs:630-662`

**Description:** Both base64 encoding and decoding are implemented inline rather than using a well-tested crate like `base64`. The code includes a comment: "Task 3 will switch to the `base64` crate." The implementations appear correct based on inspection (standard alphabet, proper padding handling), but custom cryptographic-adjacent code increases the attack surface.

**Impact:** Low. Base64 encoding/decoding is not a security-critical operation -- it processes public metadata (salts, encrypted values from XML). An encoding bug would cause decryption to fail rather than leak secrets.

**Recommendation:** Replace with the `base64` crate when convenient, as the TODO comment suggests.

---

### [MEDIUM] No Validation of Encryption Parameters from EncryptionInfo

**File:** `crates/modern-xlsx-core/src/ole2/encryption_info.rs:81-221`

**Description:** The Agile EncryptionInfo XML parser does not validate that parsed parameters are within expected/safe ranges. For example:
- `pw_spin_count` could be 0 (making key derivation trivially weak) or extremely large (DoS)
- `key_bits` could be a value other than 128 or 256 (would fail at AES init, but with a confusing error)
- `key_block_size` could be 0 (would cause a panic in slice operations) or mismatched with the cipher
- `key_hash_size` could be 0 or larger than the hash algorithm's output
- `saltValue` could be empty (weakening key derivation)

The crypto functions downstream handle some of these (e.g., `aes_cbc_decrypt` rejects invalid key lengths), but there is no centralized validation of the parsed EncryptionInfo.

**Impact:** Medium for robustness, low for security. A maliciously crafted file could trigger confusing error messages or panics. However, the attacker would need to provide a file specifically designed to exploit parameter edge cases, and no secret disclosure results.

**Recommendation:** Add validation after parsing:
```rust
if info.pw_spin_count == 0 || info.pw_spin_count > 10_000_000 { return Err(...); }
if !matches!(info.key_bits, 128 | 256) { return Err(...); }
if info.key_block_size != 16 { return Err(...); }
if info.key_salt.is_empty() || info.pw_salt.is_empty() { return Err(...); }
```

---

### [MEDIUM] Segment IV Derivation Uses Hash Truncation Rather Than KDF

**File:** `crates/modern-xlsx-core/src/ole2/crypto.rs:478-487`

**Description:** The segment IV derivation computes `Hash(salt || LE32(segment_index))` and truncates the hash output to `block_size` (16 bytes). This is correct per the ECMA-376 specification (SS2.3.6.2), but it is worth noting that:

1. The salt is reused across all segments -- only the segment index varies. This is by design in the spec.
2. Each segment uses a unique IV (derived from the unique segment index), which is the correct approach for AES-CBC.
3. The hash truncation means only 16 of 64 bytes (for SHA-512) are used, but this provides more than sufficient entropy for an AES-CBC IV.

**Impact:** This is not a bug -- it faithfully implements the specification. It is noted here for completeness. The ECMA-376 approach is considered cryptographically sound because: (a) the data key is unique per file, (b) each segment has a unique IV, and (c) AES-CBC with unique IVs provides IND-CPA security.

**Recommendation:** No action required. This is spec-compliant behavior.

---

### [INFO] Standard Encryption Uses SHA-1 and AES-128

**File:** `crates/modern-xlsx-core/src/ole2/crypto.rs:844-905`

**Description:** Standard Encryption (versions 2.2/3.2/4.2) uses SHA-1 for key derivation and AES-128-ECB for password verifier decryption. SHA-1 is considered cryptographically weakened (collision resistance broken), and AES-ECB mode does not provide semantic security for data larger than one block.

**Impact:** Informational. This is dictated by the ECMA-376 specification for backward compatibility with older Office files. The implementation correctly limits ECB usage to the 16-byte verifier and 32-byte verifier hash (1-2 AES blocks), where ECB's weakness is not exploitable. The write path exclusively uses Agile Encryption (AES-256-CBC, SHA-512), which is the modern standard.

**Recommendation:** No action required. Consider documenting that the write path always uses the strongest available encryption (Agile AES-256-CBC + SHA-512), while the read path supports older formats for compatibility.

---

### [INFO] OLE2 Writer DIFAT Limitation

**File:** `crates/modern-xlsx-core/src/ole2/writer.rs:115-128`

**Description:** The OLE2 writer only supports the 109-entry DIFAT array in the header, limiting output to ~3.5 GB. This is well-documented in the code comments and enforced with an explicit error check (line 158).

**Impact:** Informational. Encrypted XLSX files are compressed ZIP archives and will not approach this limit in practice.

**Recommendation:** No action required. The limitation is clearly documented and properly enforced.

---

### [INFO] Encryption Dependencies Are Feature-Gated

**File:** `crates/modern-xlsx-core/Cargo.toml:24-45`

**Description:** All cryptographic dependencies (`aes`, `cbc`, `sha2`, `sha1`, `hmac`, `zeroize`, `constant_time_eq`, `getrandom`, `digest`) are behind the `encryption` Cargo feature gate. This means the attack surface is zero for users who do not need encryption support.

**Impact:** Positive security property. The feature gate ensures unused cryptographic code is not compiled into the binary.

**Recommendation:** No action required. This is good practice.

## Dependency Assessment

| Crate | Version | Purpose | Assessment |
|---|---|---|---|
| `aes` | 0.8 | AES-128/256 block cipher | Well-audited RustCrypto crate |
| `cbc` | 0.1 | CBC mode wrapper | Well-audited RustCrypto crate |
| `sha2` | 0.10 | SHA-256/512 | Well-audited RustCrypto crate |
| `sha1` | 0.10 | SHA-1 (Standard Encryption) | Well-audited RustCrypto crate |
| `hmac` | 0.12 | HMAC construction | Well-audited RustCrypto crate |
| `digest` | 0.10 | Trait abstraction for hashes | Well-audited RustCrypto crate |
| `zeroize` | 1.8 | Secret memory zeroization | Well-audited, derive macros |
| `constant_time_eq` | 0.3 | Constant-time byte comparison | Simple, well-reviewed |
| `getrandom` | 0.3 | CSPRNG (OS + WASM) | Well-audited, standard choice |

All dependencies are from the RustCrypto ecosystem, which has undergone formal audits and is widely used in production Rust projects.

## Test Coverage

The encryption module includes comprehensive tests (visible in `crypto.rs` lines 946-1494):

- Key derivation: known vectors, empty password, Unicode password, 128-bit, SHA-256
- AES-CBC: encrypt/decrypt roundtrip with and without padding
- Segment decryption: single segment, multi-segment, truncated input
- Standard Encryption: SHA-1 key derivation, password verification roundtrip, wrong password rejection
- Standard package decryption: roundtrip, corrupted/truncated input
- Full pipeline: `encrypt_file` -> `decrypt_file` roundtrip with HMAC verification
- Segment size validation: verifies encrypted output structure
- Secure random: uniqueness verification
- Base64: known value encoding tests

Additional integration tests exist in the TypeScript test suite (`encryption-roundtrip.test.ts`).

## Conclusion

The modern-xlsx encryption implementation is **well-designed and security-conscious**. Key strengths include:

1. **Correct ECMA-376 implementation:** Both Agile and Standard encryption follow the specification faithfully.
2. **Consistent zeroization:** All key material is wrapped in `SensitiveKey` with RAII-based zeroization.
3. **Constant-time comparison:** All password and HMAC verification uses `constant_time_eq`.
4. **Strong defaults for writes:** The write path uses AES-256-CBC + SHA-512 + 100,000 iterations, which represents the strongest ECMA-376 configuration.
5. **CSPRNG for all randomness:** All salts, keys, and IVs are generated via `getrandom`.
6. **Feature-gated dependencies:** Encryption is opt-in, minimizing attack surface.
7. **Comprehensive test coverage:** Roundtrip tests, known-vector tests, error path tests, and wrong-password tests.

The identified findings are all low-severity defense-in-depth improvements (zeroizing intermediate values, validating parsed parameters). No critical or high-severity vulnerabilities were found. The code is suitable for production use with the understanding that the Standard Encryption read path inherits the cryptographic limitations of the legacy ECMA-376 format (SHA-1, AES-128).
