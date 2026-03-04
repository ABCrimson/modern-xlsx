# Security & Encryption (0.6.x) — TDD Implementation Plan

> **For Claude:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development to implement this plan task-by-task.

**Goal:** Add full ECMA-376 Agile Encryption support — detect OLE2 compound documents, parse encryption metadata, derive keys via SHA-512, decrypt/encrypt XLSX files with AES-256-CBC, and expose `password` option on both read and write paths.

**Architecture:** Insert a detection layer before the ZIP reader. OLE2 compound documents wrap an `EncryptionInfo` stream (XML-based metadata) and an `EncryptedPackage` stream (AES-256-CBC encrypted ZIP bytes). Decryption pipeline: OLE2 parse → read EncryptionInfo → derive keys from password → AES-256-CBC decrypt segments → extract ZIP → normal reader. Encryption pipeline: normal writer → ZIP bytes → AES-256-CBC encrypt segments → HMAC integrity → OLE2 package.

**Tech Stack:** Rust 1.94 (quick-xml 0.39.2, sha2 0.10, aes 0.8, cbc 0.1, hmac 0.12, getrandom 0.3 with `wasm_js` feature, zeroize 1.8, constant_time_eq 0.3), TypeScript 6.0, Vitest 4.1

**New Rust Dependencies:**
```toml
sha2 = "0.10"
aes = "0.8"
cbc = "0.1"
hmac = "0.12"
getrandom = { version = "0.3", features = ["wasm_js"] }
zeroize = { version = "1.8", features = ["derive"] }
constant_time_eq = "0.3"
base64 = "0.22"
```

All are pure-Rust RustCrypto crates — zero C dependencies, full WASM compatibility.

---

## Task 1 (0.6.0): OLE2 Magic Byte Detection + Descriptive Errors

**Files:**
- Create: `crates/modern-xlsx-core/src/ole2/mod.rs`
- Create: `crates/modern-xlsx-core/src/ole2/detect.rs`
- Modify: `crates/modern-xlsx-core/src/lib.rs` (add `mod ole2`)
- Modify: `crates/modern-xlsx-core/src/errors.rs` (add error variants)
- Modify: `crates/modern-xlsx-core/src/reader.rs` (add detection at entry)
- Modify: `crates/modern-xlsx-wasm/src/lib.rs` (no changes needed — errors propagate via JsError)
- Modify: `packages/modern-xlsx/src/wasm-loader.ts` (no changes needed — errors propagate)

### Error enum additions (`errors.rs`):

```rust
/// The file is a password-protected OLE2 compound document.
PasswordProtected(String),
/// The file is a legacy .xls (OLE2) format, not .xlsx.
LegacyFormat(String),
/// The file format is unrecognized (not ZIP, not OLE2).
UnrecognizedFormat(String),
```

### Detection logic (`ole2/detect.rs`):

```rust
/// OLE2 Compound Document magic bytes.
const OLE2_MAGIC: [u8; 8] = [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1];

/// ZIP archive magic bytes.
const ZIP_MAGIC: [u8; 4] = [0x50, 0x4B, 0x03, 0x04];

/// Detects the file format from the first few bytes.
pub enum FileFormat {
    /// Standard XLSX (ZIP archive).
    Zip,
    /// OLE2 compound document — either encrypted XLSX or legacy .xls.
    Ole2,
    /// Unrecognized format.
    Unknown,
}

pub fn detect_format(data: &[u8]) -> FileFormat {
    if data.len() >= 8 && data[..8] == OLE2_MAGIC {
        FileFormat::Ole2
    } else if data.len() >= 4 && data[..4] == ZIP_MAGIC {
        FileFormat::Zip
    } else {
        FileFormat::Unknown
    }
}
```

### OLE2 classification (`ole2/detect.rs`):

When OLE2 detected, read the directory to classify:

```rust
/// Classifies an OLE2 file as encrypted XLSX or legacy .xls.
pub enum Ole2Kind {
    /// Password-protected XLSX (has EncryptionInfo + EncryptedPackage streams).
    EncryptedXlsx,
    /// Legacy .xls binary format (has Workbook or Book stream).
    LegacyXls,
    /// Unknown OLE2 content.
    Unknown,
}

pub fn classify_ole2(data: &[u8]) -> Result<Ole2Kind> {
    // Parse OLE2 header (first 512 bytes)
    // Read directory sector chain
    // Check for stream names: "EncryptionInfo", "EncryptedPackage" → EncryptedXlsx
    // Check for stream names: "Workbook", "Book" → LegacyXls
    // Otherwise: Unknown
}
```

### OLE2 header parsing (minimal, just for directory traversal):

```rust
/// Reads OLE2 header. We only need enough to find directory entries.
pub struct Ole2Header {
    /// Sector size in bytes (typically 512).
    pub sector_size: u32,
    /// Mini sector size (typically 64).
    pub mini_sector_size: u32,
    /// First directory sector ID.
    pub first_dir_sector: u32,
    /// First FAT sector IDs (up to 109 in header).
    pub fat_sectors: Vec<u32>,
    /// First DIFAT sector (for files > ~7MB).
    pub first_difat_sector: u32,
    /// Number of DIFAT sectors.
    pub num_difat_sectors: u32,
}

impl Ole2Header {
    pub fn parse(data: &[u8]) -> Result<Self> {
        if data.len() < 512 {
            return Err(ModernXlsxError::UnrecognizedFormat("File too small for OLE2".into()));
        }
        // Byte 0-7: magic (already validated)
        // Byte 28-29: minor version
        // Byte 30-31: major version (3 or 4)
        // Byte 30 == 0x03 → sector_size = 512
        // Byte 30 == 0x04 → sector_size = 4096
        // Byte 44-47: first directory sector SID
        // Byte 60-63: first mini FAT sector SID
        // Byte 68-71: first DIFAT sector SID
        // Byte 72-75: number of DIFAT sectors
        // Byte 76-511: first 109 FAT sector IDs (4 bytes each)
        ...
    }
}
```

### Directory entry parsing:

```rust
/// OLE2 directory entry (128 bytes each).
pub struct DirEntry {
    pub name: String,
    pub entry_type: u8,      // 0=unknown, 1=storage, 2=stream, 5=root
    pub start_sector: u32,
    pub size: u64,            // v4: 64-bit, v3: 32-bit
}

/// Reads all directory entries from the directory sector chain.
pub fn read_directory(data: &[u8], header: &Ole2Header) -> Result<Vec<DirEntry>> {
    // Follow directory sector chain via FAT
    // Each sector contains sector_size / 128 entries
    // Read UTF-16LE names (first 64 bytes of each 128-byte entry)
    // Extract entry_type (byte 66), start_sector (bytes 116-119), size (bytes 120-127 or 120-123)
    ...
}
```

### Reader integration (`reader.rs`):

At the top of `read_xlsx_json()` and `read_xlsx()`, before `read_zip_entries()`:

```rust
use crate::ole2::detect::{detect_format, FileFormat, classify_ole2, Ole2Kind};

pub fn read_xlsx_json(data: &[u8]) -> Result<String> {
    match detect_format(data) {
        FileFormat::Zip => { /* existing path */ }
        FileFormat::Ole2 => {
            match classify_ole2(data)? {
                Ole2Kind::EncryptedXlsx => {
                    return Err(ModernXlsxError::PasswordProtected(
                        "This file is password-protected (OLE2 compound document). \
                         Decryption not yet supported in this version.".into()
                    ));
                }
                Ole2Kind::LegacyXls => {
                    return Err(ModernXlsxError::LegacyFormat(
                        "Legacy .xls format not supported. Convert to .xlsx first.".into()
                    ));
                }
                Ole2Kind::Unknown => {
                    return Err(ModernXlsxError::UnrecognizedFormat(
                        "Unrecognized OLE2 compound document.".into()
                    ));
                }
            }
        }
        FileFormat::Unknown => {
            return Err(ModernXlsxError::UnrecognizedFormat(
                "Not a valid XLSX file (expected ZIP or OLE2 header).".into()
            ));
        }
    }
    // ... existing read logic
}
```

### Tests (5 Rust unit tests in `ole2/detect.rs`):

1. **test_detect_zip_format**: `detect_format(&[0x50, 0x4B, 0x03, 0x04, ...])` → `FileFormat::Zip`
2. **test_detect_ole2_format**: `detect_format(&OLE2_MAGIC)` → `FileFormat::Ole2`
3. **test_detect_unknown_format**: `detect_format(&[0x00, 0x01, ...])` → `FileFormat::Unknown`
4. **test_encrypted_xlsx_error**: Build minimal OLE2 with EncryptionInfo stream name in directory → `PasswordProtected` error
5. **test_legacy_xls_error**: Build minimal OLE2 with "Workbook" stream name → `LegacyFormat` error

### TypeScript tests (5 in `packages/modern-xlsx/__tests__/encryption.test.ts`):

1. **test: encrypted XLSX shows descriptive error**: Construct OLE2 bytes with EncryptionInfo directory entry → `readBuffer()` throws "password-protected"
2. **test: legacy .xls shows appropriate error**: Construct OLE2 bytes with Workbook directory entry → `readBuffer()` throws "Legacy .xls"
3. **test: unknown format shows clear error**: Pass random bytes → `readBuffer()` throws "Not a valid XLSX"
4. **test: normal XLSX still works**: Standard workbook → `readBuffer()` succeeds
5. **test: empty file shows error**: Pass empty Uint8Array → `readBuffer()` throws

For tests 1-2, construct minimal synthetic OLE2 bytes:
```typescript
function buildMinimalOle2(streamNames: string[]): Uint8Array {
    // 512-byte header + 512-byte directory sector
    const buf = new Uint8Array(1024);
    // Magic bytes
    buf.set([0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1]);
    // Major version 3, minor version 62
    buf[28] = 62; buf[30] = 3;
    // Byte order: 0xFFFE
    buf[26] = 0xFE; buf[27] = 0xFF;
    // Sector size shift: 9 (512 bytes)
    buf[30] = 3; buf[31] = 0; // version
    buf[30] = 0x03; buf[31] = 0x00;
    // ... (fill header fields to make parseable)
    // Directory sector: write directory entries with UTF-16LE stream names
    // Root entry + stream entries
    ...
    return buf;
}
```

### WASM rebuild after Rust changes.

---

## Task 2 (0.6.1): Encryption Method Identification

**Files:**
- Create: `crates/modern-xlsx-core/src/ole2/encryption_info.rs`
- Modify: `crates/modern-xlsx-core/src/ole2/mod.rs` (add module)
- Modify: `crates/modern-xlsx-core/src/ole2/detect.rs` (add stream reading)

### Read OLE2 stream by name:

```rust
/// Reads the raw bytes of a named stream from the OLE2 compound document.
pub fn read_stream(data: &[u8], header: &Ole2Header, dir: &[DirEntry], name: &str) -> Result<Vec<u8>> {
    // Find directory entry by name
    // Follow FAT chain from start_sector
    // Concatenate sector data
    // Truncate to entry.size
    ...
}
```

### EncryptionInfo parsing (`ole2/encryption_info.rs`):

```rust
/// ECMA-376 encryption version.
#[derive(Debug, Clone, PartialEq)]
pub enum EncryptionVersion {
    /// Standard Encryption (2.3.6.1) — RC4 or AES-128.
    Standard { flags: u32, header_size: u32 },
    /// Agile Encryption (2.3.6.2) — AES-128/256 with SHA-1/256/512.
    Agile,
    /// Extensible Encryption (2.3.6.3).
    Extensible,
}

/// Parsed Agile encryption descriptor.
#[derive(Debug, Clone)]
pub struct AgileEncryptionInfo {
    // Key data
    pub key_salt: Vec<u8>,
    pub key_block_size: u32,      // typically 16
    pub key_bits: u32,            // 128 or 256
    pub key_hash_size: u32,       // 32 or 64
    pub key_cipher: String,       // "AES"
    pub key_chaining: String,     // "ChainingModeCBC"
    pub key_hash_alg: String,     // "SHA512"
    // Data integrity
    pub encrypted_hmac_key: Vec<u8>,
    pub encrypted_hmac_value: Vec<u8>,
    // Password key encryptor
    pub pw_spin_count: u32,       // typically 100000
    pub pw_salt: Vec<u8>,
    pub pw_block_size: u32,
    pub pw_key_bits: u32,
    pub pw_hash_size: u32,
    pub pw_cipher: String,
    pub pw_chaining: String,
    pub pw_hash_alg: String,
    pub pw_encrypted_key_value: Vec<u8>,
    pub pw_encrypted_verifier_hash_input: Vec<u8>,
    pub pw_encrypted_verifier_hash_value: Vec<u8>,
}

/// Standard Encryption info.
#[derive(Debug, Clone)]
pub struct StandardEncryptionInfo {
    pub alg_id: u32,           // 0x6801 = AES-128, 0x6802 = AES-192, 0x6803 = AES-256
    pub hash_alg_id: u32,     // 0x8004 = SHA-1
    pub key_size: u32,         // 128
    pub provider: String,      // e.g. "Microsoft Enhanced RSA and AES Cryptographic Provider"
    pub salt: Vec<u8>,
    pub encrypted_verifier: Vec<u8>,
    pub verifier_hash_size: u32,
    pub encrypted_verifier_hash: Vec<u8>,
}

/// Combined encryption info.
#[derive(Debug, Clone)]
pub enum EncryptionInfo {
    Standard(StandardEncryptionInfo),
    Agile(AgileEncryptionInfo),
}

impl EncryptionInfo {
    /// Parses the EncryptionInfo stream bytes.
    pub fn parse(stream: &[u8]) -> Result<Self> {
        if stream.len() < 8 {
            return Err(ModernXlsxError::PasswordProtected(
                "EncryptionInfo stream too short".into()
            ));
        }
        let version_major = u16::from_le_bytes([stream[0], stream[1]]);
        let version_minor = u16::from_le_bytes([stream[2], stream[3]]);

        match (version_major, version_minor) {
            (4, 4) => Self::parse_agile(&stream[8..]),  // Skip 8-byte header
            (2, 2) | (3, 2) | (4, 2) => Self::parse_standard(&stream[4..]),
            _ => Err(ModernXlsxError::PasswordProtected(
                format!("Unsupported encryption version {version_major}.{version_minor}")
            )),
        }
    }

    fn parse_agile(xml_bytes: &[u8]) -> Result<Self> {
        // Parse XML using quick-xml SAX parser
        // Extract keyData, dataIntegrity, keyEncryptors elements
        // Decode base64 attributes (salt, encrypted values)
        ...
    }

    fn parse_standard(data: &[u8]) -> Result<Self> {
        // Binary format: flags (4 bytes), header size (4 bytes)
        // Header: flags, sizeExtra, algID, algIDHash, keySize, providerType, reserved, cspName
        // Verifier: salt (16 bytes), encrypted verifier (16 bytes), hash size, encrypted hash
        ...
    }
}
```

### Enhanced error messages:

Update Task 1's error path to include encryption details:

```rust
Ole2Kind::EncryptedXlsx => {
    // Try to parse EncryptionInfo for better error message
    match read_and_parse_encryption_info(data) {
        Ok(info) => {
            let desc = match &info {
                EncryptionInfo::Agile(a) => format!(
                    "Agile encryption ({}-{}, {})",
                    a.key_cipher, a.key_bits, a.key_hash_alg
                ),
                EncryptionInfo::Standard(s) => format!(
                    "Standard encryption (key size: {} bits)", s.key_size
                ),
            };
            Err(ModernXlsxError::PasswordProtected(
                format!("Password-protected XLSX ({desc}). \
                         Provide password via readBuffer(data, {{ password: '...' }}).")
            ))
        }
        Err(_) => Err(ModernXlsxError::PasswordProtected(
            "Password-protected XLSX (could not identify encryption method).".into()
        ))
    }
}
```

### Tests (5 Rust tests in `ole2/encryption_info.rs`):

1. **test_parse_agile_aes256_sha512**: Construct EncryptionInfo stream bytes (version 4.4 + Agile XML) → verify parsed fields
2. **test_parse_agile_aes128**: Same with AES-128 → verify key_bits = 128
3. **test_parse_standard_encryption**: Construct version 4.2 binary → verify Standard variant with correct key_size
4. **test_unsupported_version_error**: Construct version 99.99 → error
5. **test_truncated_stream_error**: Pass 4 bytes → error

### TypeScript tests (5):

1. **test: encrypted AES-256 shows method in error**: Construct OLE2 with Agile EncryptionInfo → error contains "AES-256"
2. **test: encrypted AES-128 shows method in error**: Same → error contains "AES-128"
3. **test: standard encryption shows version**: Construct OLE2 with Standard EncryptionInfo → error contains "Standard"
4. **test: encryption error includes usage hint**: Error message mentions `readBuffer(data, { password: ... })`
5. **test: malformed EncryptionInfo → graceful fallback**: Construct OLE2 with corrupt stream → still shows password-protected error

### WASM rebuild after Rust changes.

---

## Task 3 (0.6.2): Key Derivation (SHA-512)

**Files:**
- Create: `crates/modern-xlsx-core/src/ole2/crypto.rs`
- Modify: `crates/modern-xlsx-core/src/ole2/mod.rs`
- Modify: `crates/modern-xlsx-core/Cargo.toml` (add sha2, zeroize)

### Key derivation per ECMA-376 2.3.6.2:

```rust
use sha2::{Sha512, Digest};
use zeroize::Zeroize;

/// Block key constants per ECMA-376 §2.3.6.2.
const BLOCK_KEY_VERIFIER_INPUT: [u8; 8] = [0xFE, 0xA7, 0xD2, 0x76, 0x3B, 0x4B, 0x9E, 0x79];
const BLOCK_KEY_VERIFIER_VALUE: [u8; 8] = [0xD7, 0xAA, 0x0F, 0x6D, 0x30, 0x61, 0x34, 0x4E];
const BLOCK_KEY_ENCRYPTED_KEY:  [u8; 8] = [0x14, 0x6E, 0x0B, 0xE7, 0xAB, 0xAC, 0xD0, 0xD6];
const BLOCK_KEY_HMAC_KEY:       [u8; 8] = [0x5F, 0xB2, 0xAD, 0x01, 0x0C, 0xB9, 0xE1, 0xF6];
const BLOCK_KEY_HMAC_VALUE:     [u8; 8] = [0xA0, 0x67, 0x7F, 0x02, 0xB2, 0x2C, 0x84, 0x33];

/// Derives a key from password per ECMA-376 Agile Encryption §2.3.6.2.
///
/// # Arguments
/// * `password` - User password (will be encoded to UTF-16LE)
/// * `salt` - Random salt from EncryptionInfo
/// * `spin_count` - Iteration count (typically 100,000)
/// * `key_bits` - Desired key length in bits (128 or 256)
/// * `block_key` - Block key constant (determines which key is derived)
/// * `hash_alg` - Hash algorithm name ("SHA512", "SHA256", "SHA1")
///
/// # Returns
/// Derived key bytes, truncated/padded to key_bits/8.
pub fn derive_key(
    password: &str,
    salt: &[u8],
    spin_count: u32,
    key_bits: u32,
    block_key: &[u8],
    hash_alg: &str,
) -> Result<Vec<u8>> {
    // Step 1: Encode password as UTF-16LE
    let pw_bytes: Vec<u8> = password.encode_utf16()
        .flat_map(|c| c.to_le_bytes())
        .collect();

    // Step 2: H0 = SHA-512(salt + pw_bytes)
    let mut hasher = Sha512::new();
    hasher.update(salt);
    hasher.update(&pw_bytes);
    let mut hash = hasher.finalize_reset();

    // Step 3: Iterate spin_count times
    // Hi = SHA-512(LE32(i) + Hi-1)
    for i in 0..spin_count {
        hasher.update(&i.to_le_bytes());
        hasher.update(&hash);
        hash = hasher.finalize_reset();
    }

    // Step 4: Derive final key
    // Hfinal = SHA-512(Hn + blockKey)
    hasher.update(&hash);
    hasher.update(block_key);
    let derived = hasher.finalize();

    // Step 5: Truncate or pad to key_bits/8
    let key_len = (key_bits / 8) as usize;
    let mut key = vec![0u8; key_len];
    if derived.len() >= key_len {
        key.copy_from_slice(&derived[..key_len]);
    } else {
        // Pad with 0x36 per spec
        key[..derived.len()].copy_from_slice(&derived);
        for b in &mut key[derived.len()..] {
            *b = 0x36;
        }
    }

    // Zeroize intermediate values
    let _ = &mut hash; // hash will be dropped; derived too

    Ok(key)
}

/// Verifies a password against the Agile encryption info.
/// Returns the data encryption key if password is correct.
pub fn verify_password_agile(password: &str, info: &AgileEncryptionInfo) -> Result<Vec<u8>> {
    // 1. Derive verifier hash input key
    let input_key = derive_key(
        password, &info.pw_salt, info.pw_spin_count,
        info.pw_key_bits, &BLOCK_KEY_VERIFIER_INPUT, &info.pw_hash_alg,
    )?;

    // 2. Decrypt verifier hash input
    let verifier_input = aes_cbc_decrypt(
        &input_key, &info.pw_salt, &info.pw_encrypted_verifier_hash_input,
    )?;

    // 3. Derive verifier hash value key
    let value_key = derive_key(
        password, &info.pw_salt, info.pw_spin_count,
        info.pw_key_bits, &BLOCK_KEY_VERIFIER_VALUE, &info.pw_hash_alg,
    )?;

    // 4. Decrypt verifier hash value
    let verifier_hash = aes_cbc_decrypt(
        &value_key, &info.pw_salt, &info.pw_encrypted_verifier_hash_value,
    )?;

    // 5. Compute hash of verifier input
    let mut hasher = Sha512::new();
    hasher.update(&verifier_input);
    let computed = hasher.finalize();

    // 6. Constant-time comparison (truncate to hash_size)
    let hash_len = info.pw_hash_size as usize;
    if !constant_time_eq::constant_time_eq(
        &computed[..hash_len],
        &verifier_hash[..hash_len],
    ) {
        return Err(ModernXlsxError::PasswordProtected(
            "Incorrect password.".into()
        ));
    }

    // 7. Derive data encryption key
    let enc_key = derive_key(
        password, &info.pw_salt, info.pw_spin_count,
        info.pw_key_bits, &BLOCK_KEY_ENCRYPTED_KEY, &info.pw_hash_alg,
    )?;

    // 8. Decrypt the actual data encryption key
    let data_key = aes_cbc_decrypt(
        &enc_key, &info.pw_salt, &info.pw_encrypted_key_value,
    )?;

    Ok(data_key[..info.pw_key_bits as usize / 8].to_vec())
}
```

### Tests (5 Rust tests in `ole2/crypto.rs`):

Use known test vectors from the ECMA-376 spec and real Excel-encrypted files.

1. **test_derive_key_known_vector**: Password "password", known salt → expected key bytes
2. **test_derive_key_empty_password**: Empty string "" → valid key (not an error)
3. **test_derive_key_unicode**: Password "пароль" (Russian) → valid UTF-16LE encoding → valid key
4. **test_derive_key_high_spin_count**: spin_count=1 → fast, verify deterministic output
5. **test_verify_password_wrong**: Known encrypted info + wrong password → "Incorrect password" error

**Test vector generation:** Use a Python script or Excel to create a file with known password, extract EncryptionInfo bytes, and hardcode expected derived keys in tests. Alternatively, generate test vectors using the reference implementation in the ECMA-376 spec appendix.

---

## Task 4 (0.6.3): AES-256-CBC Decryption

**Files:**
- Modify: `crates/modern-xlsx-core/src/ole2/crypto.rs`
- Modify: `crates/modern-xlsx-core/Cargo.toml` (add aes, cbc, hmac, constant_time_eq)

### AES-CBC primitives:

```rust
use aes::Aes256;
use cbc::{Decryptor, Encryptor};
use cbc::cipher::{BlockDecryptMut, KeyIvInit};

type Aes256CbcDec = Decryptor<Aes256>;
type Aes256CbcEnc = Encryptor<Aes256>;

/// Decrypts data using AES-256-CBC with PKCS#7 padding.
pub fn aes_cbc_decrypt(key: &[u8], iv: &[u8], data: &[u8]) -> Result<Vec<u8>> {
    let decryptor = Aes256CbcDec::new_from_slices(key, iv)
        .map_err(|e| ModernXlsxError::PasswordProtected(format!("AES init error: {e}")))?;

    let mut buf = data.to_vec();
    let plaintext = decryptor.decrypt_padded_mut::<aes::cipher::block_padding::Pkcs7>(&mut buf)
        .map_err(|_| ModernXlsxError::PasswordProtected("AES decryption failed (bad padding)".into()))?;

    Ok(plaintext.to_vec())
}

/// Decrypts data without removing padding (for raw blocks).
pub fn aes_cbc_decrypt_no_pad(key: &[u8], iv: &[u8], data: &[u8]) -> Result<Vec<u8>> {
    let decryptor = Aes256CbcDec::new_from_slices(key, iv)
        .map_err(|e| ModernXlsxError::PasswordProtected(format!("AES init error: {e}")))?;

    let mut buf = data.to_vec();
    decryptor.decrypt_padded_mut::<aes::cipher::block_padding::NoPadding>(&mut buf)
        .map_err(|_| ModernXlsxError::PasswordProtected("AES decryption failed".into()))?;

    Ok(buf)
}
```

### Segment-based decryption of EncryptedPackage:

```rust
/// Segment size for ECMA-376 Agile Encryption.
const SEGMENT_SIZE: usize = 4096;

/// Decrypts the EncryptedPackage stream.
///
/// Layout: [8 bytes: original size LE64] [encrypted segments of 4096 bytes each]
/// Each segment IV = SHA-512(salt + LE32(segment_index)) truncated to block_size.
pub fn decrypt_package(
    data_key: &[u8],
    info: &AgileEncryptionInfo,
    encrypted_package: &[u8],
) -> Result<Vec<u8>> {
    if encrypted_package.len() < 8 {
        return Err(ModernXlsxError::PasswordProtected(
            "EncryptedPackage too short".into()
        ));
    }

    // Read original size
    let original_size = u64::from_le_bytes(encrypted_package[..8].try_into().unwrap()) as usize;
    let payload = &encrypted_package[8..];

    let block_size = info.key_block_size as usize;
    let mut result = Vec::with_capacity(original_size);
    let mut segment_idx: u32 = 0;

    for chunk in payload.chunks(SEGMENT_SIZE) {
        // Derive per-segment IV
        let iv = derive_segment_iv(&info.key_salt, segment_idx, block_size);

        // Decrypt segment (no PKCS7 — last segment is truncated to original_size)
        let decrypted = aes_cbc_decrypt_no_pad(data_key, &iv, chunk)?;
        result.extend_from_slice(&decrypted);
        segment_idx += 1;
    }

    // Truncate to original size
    result.truncate(original_size);
    Ok(result)
}

/// Derives the IV for a given segment index.
fn derive_segment_iv(salt: &[u8], segment_index: u32, block_size: usize) -> Vec<u8> {
    let mut hasher = Sha512::new();
    hasher.update(salt);
    hasher.update(&segment_index.to_le_bytes());
    let hash = hasher.finalize();
    hash[..block_size].to_vec()
}
```

### HMAC integrity verification:

```rust
use hmac::{Hmac, Mac};
type HmacSha512 = Hmac<Sha512>;

/// Verifies the HMAC integrity of the encrypted package.
pub fn verify_hmac(
    data_key: &[u8],
    info: &AgileEncryptionInfo,
    encrypted_package: &[u8],
) -> Result<()> {
    // 1. Derive HMAC key
    let hmac_key_key = derive_key_raw(
        data_key, &info.key_salt, &BLOCK_KEY_HMAC_KEY,
        info.key_hash_size, info.key_block_size as usize,
    );
    let hmac_key = aes_cbc_decrypt(
        &hmac_key_key[..info.key_bits as usize / 8],
        &info.key_salt,
        &info.encrypted_hmac_key,
    )?;

    // 2. Derive HMAC value
    let hmac_value_key = derive_key_raw(
        data_key, &info.key_salt, &BLOCK_KEY_HMAC_VALUE,
        info.key_hash_size, info.key_block_size as usize,
    );
    let expected_hmac = aes_cbc_decrypt(
        &hmac_value_key[..info.key_bits as usize / 8],
        &info.key_salt,
        &info.encrypted_hmac_value,
    )?;

    // 3. Compute HMAC of encrypted package
    let mut mac = HmacSha512::new_from_slice(&hmac_key[..info.key_hash_size as usize])
        .map_err(|e| ModernXlsxError::PasswordProtected(format!("HMAC init: {e}")))?;
    mac.update(encrypted_package);
    let computed = mac.finalize().into_bytes();

    // 4. Constant-time compare
    let hash_len = info.key_hash_size as usize;
    if !constant_time_eq::constant_time_eq(
        &computed[..hash_len],
        &expected_hmac[..hash_len],
    ) {
        return Err(ModernXlsxError::PasswordProtected(
            "HMAC verification failed — file may be corrupted or tampered with.".into()
        ));
    }

    Ok(())
}
```

### Tests (5 Rust tests):

1. **test_aes_cbc_decrypt_known_vector**: Known AES-256-CBC test vector → correct plaintext
2. **test_decrypt_single_segment**: 4096 bytes encrypted → decrypt → verify content
3. **test_decrypt_multi_segment**: 12288 bytes (3 segments) → decrypt → verify
4. **test_hmac_verification_pass**: Valid HMAC → Ok(())
5. **test_hmac_verification_fail**: Corrupted data → HMAC fail error

---

## Task 5 (0.6.4): File Decryption Integration (Read Path)

**Files:**
- Modify: `crates/modern-xlsx-core/src/reader.rs`
- Modify: `crates/modern-xlsx-core/src/ole2/detect.rs` (add `decrypt_file` function)
- Modify: `crates/modern-xlsx-wasm/src/lib.rs` (add password parameter to read)
- Modify: `packages/modern-xlsx/src/wasm-loader.ts` (add ReadOptions)
- Modify: `packages/modern-xlsx/src/workbook.ts` (readBuffer password option)
- Modify: `packages/modern-xlsx/src/index.ts` (export ReadOptions)
- Modify: `packages/modern-xlsx/src/types.ts` (ReadOptions type)

### Full decryption pipeline (`ole2/detect.rs`):

```rust
/// Decrypts an OLE2-wrapped encrypted XLSX file.
/// Returns the decrypted ZIP bytes ready for normal reading.
pub fn decrypt_file(data: &[u8], password: &str) -> Result<Vec<u8>> {
    let header = Ole2Header::parse(data)?;
    let dir = read_directory(data, &header)?;

    // Read EncryptionInfo stream
    let enc_info_bytes = read_stream(data, &header, &dir, "EncryptionInfo")?;
    let enc_info = EncryptionInfo::parse(&enc_info_bytes)?;

    match enc_info {
        EncryptionInfo::Agile(ref agile) => {
            // Verify password and get data key
            let data_key = verify_password_agile(password, agile)?;

            // Read EncryptedPackage stream
            let encrypted_package = read_stream(data, &header, &dir, "EncryptedPackage")?;

            // Verify HMAC integrity
            verify_hmac(&data_key, agile, &encrypted_package)?;

            // Decrypt package → ZIP bytes
            decrypt_package(&data_key, agile, &encrypted_package)
        }
        EncryptionInfo::Standard(ref std_info) => {
            decrypt_standard(data, &header, &dir, password, std_info)
        }
    }
}
```

### Reader integration:

```rust
/// Options for reading XLSX files.
pub struct ReadOptions {
    pub password: Option<String>,
    pub limits: ZipSecurityLimits,
}

pub fn read_xlsx_json_with_password(data: &[u8], options: &ReadOptions) -> Result<String> {
    let actual_data: std::borrow::Cow<[u8]> = match detect_format(data) {
        FileFormat::Zip => std::borrow::Cow::Borrowed(data),
        FileFormat::Ole2 => {
            match classify_ole2(data)? {
                Ole2Kind::EncryptedXlsx => {
                    if let Some(ref password) = options.password {
                        let decrypted = decrypt_file(data, password)?;
                        std::borrow::Cow::Owned(decrypted)
                    } else {
                        let info = read_and_parse_encryption_info(data).ok();
                        let msg = format_password_required_message(info.as_ref());
                        return Err(ModernXlsxError::PasswordProtected(msg));
                    }
                }
                Ole2Kind::LegacyXls => {
                    return Err(ModernXlsxError::LegacyFormat(
                        "Legacy .xls format not supported. Convert to .xlsx first.".into()
                    ));
                }
                Ole2Kind::Unknown => {
                    return Err(ModernXlsxError::UnrecognizedFormat(
                        "Unrecognized OLE2 compound document.".into()
                    ));
                }
            }
        }
        FileFormat::Unknown => {
            return Err(ModernXlsxError::UnrecognizedFormat(
                "Not a valid XLSX file (expected ZIP or OLE2 header).".into()
            ));
        }
    };

    // Continue with normal read path using decrypted (or original) data
    read_xlsx_json_internal(&actual_data, &options.limits)
}
```

### WASM bridge update:

```rust
#[wasm_bindgen]
pub fn read_with_password(data: &[u8], password: &str) -> Result<String, JsError> {
    let options = ReadOptions {
        password: if password.is_empty() { None } else { Some(password.to_string()) },
        limits: ZipSecurityLimits::default(),
    };
    modern_xlsx_core::reader::read_xlsx_json_with_password(data, &options)
        .map_err(|e| JsError::new(&e.to_string()))
}
```

### TypeScript types:

```typescript
// types.ts
export interface ReadOptions {
    /** Password for encrypted XLSX files. */
    password?: string;
}
```

### TypeScript API:

```typescript
// index.ts
export async function readBuffer(
    data: Uint8Array,
    options?: ReadOptions,
): Promise<Workbook> {
    await ensureReady();
    const password = options?.password ?? '';
    const raw: WorkbookData = password
        ? wasmReadWithPassword(data, password)
        : wasmRead(data);
    return new Workbook(raw);
}
```

### Tests (5 integration tests):

These need actual encrypted XLSX test fixtures. Generate them:

**Test fixture generation (one-time setup):**
Use Python + `msoffcrypto-tool` or create a helper script:
```python
import msoffcrypto
# Create encrypted test files with known passwords
```

Alternatively, embed pre-generated encrypted bytes as base64 constants in tests.

**Simpler approach for CI:** Generate encrypted test fixtures in the test itself using the write path (once Task 8 is done). For now, use synthetic encrypted data or build fixtures from known byte patterns.

1. **test_read_encrypted_correct_password**: Encrypted XLSX + correct password → Workbook with expected data
2. **test_read_encrypted_wrong_password**: Encrypted XLSX + wrong password → "Incorrect password" error
3. **test_read_encrypted_no_password**: Encrypted XLSX + no password → "password required" error with method info
4. **test_read_normal_xlsx_with_password_ignored**: Normal XLSX + password option → reads normally (password ignored for non-encrypted)
5. **test_read_encrypted_preserves_all_features**: Encrypted XLSX with styles, formulas, merges → all preserved after decrypt

### WASM rebuild after Rust changes.

---

## Task 6 (0.6.5): Decryption Compatibility & Edge Cases

**Files:**
- Modify: `crates/modern-xlsx-core/src/ole2/crypto.rs` (Standard encryption support)
- Modify: `crates/modern-xlsx-core/src/ole2/detect.rs` (read-only flag handling)
- Modify: `crates/modern-xlsx-core/src/reader.rs`

### Standard encryption (pre-Agile):

```rust
/// Decrypts a file using Standard Encryption (ECMA-376 §2.3.6.1).
/// Uses AES-128-ECB for key derivation and AES-128-CBC for decryption.
pub fn decrypt_standard(
    data: &[u8],
    header: &Ole2Header,
    dir: &[DirEntry],
    password: &str,
    info: &StandardEncryptionInfo,
) -> Result<Vec<u8>> {
    // 1. SHA-1 based key derivation (different from Agile)
    // H0 = SHA1(salt + UTF-16LE(password))
    // X1 = SHA1(iterator + H0) — 50000 iterations default
    // X2 = SHA1(X1 + blockKey)
    // cbRequiredKeyLength = keySize/8
    // cbHash = 20 (SHA-1)
    // If cbRequiredKeyLength <= cbHash: key = X2[..cbRequiredKeyLength]
    // Else: derive via SHA1(X2 + 0x00...) || SHA1(X2 + 0x01...) etc.
    ...

    // 2. Verify password using encrypted verifier
    // Decrypt verifier with AES-128-ECB
    // Compute SHA-1(verifier)
    // Compare with decrypted verifier hash
    ...

    // 3. Read and decrypt EncryptedPackage
    let encrypted_package = read_stream(data, header, dir, "EncryptedPackage")?;
    // Single-stream AES-128-CBC decryption (no segments)
    ...
}
```

### Read-only recommended detection:

In workbook.xml, check for `<fileSharing>` element:
```xml
<fileSharing readOnlyRecommended="1" userName="Author"/>
```

This is NOT encryption — the file is a normal ZIP. The `readOnlyRecommended` flag is just metadata. No special handling needed in the reader beyond what already exists (it parses workbook.xml normally).

### Tests (5):

1. **test_standard_encryption_aes128**: Standard encrypted file → correct decrypt
2. **test_standard_encryption_wrong_password**: Wrong password → error
3. **test_read_only_recommended_no_password_needed**: Normal XLSX with `<fileSharing>` → reads without password
4. **test_encrypted_with_workbook_protection**: Encrypted + `<workbookProtection>` → both layers work
5. **test_corrupted_encrypted_package**: Truncated EncryptedPackage → clear error

### WASM rebuild after Rust changes.

---

## Task 7 (0.6.6): OLE2 Compound Document Writer

**Files:**
- Create: `crates/modern-xlsx-core/src/ole2/writer.rs`
- Modify: `crates/modern-xlsx-core/src/ole2/mod.rs`

### OLE2 writer:

```rust
/// Generates an OLE2 compound document containing the given streams.
pub fn write_ole2(streams: &[(&str, &[u8])]) -> Result<Vec<u8>> {
    let sector_size: usize = 512;

    // Calculate sizes
    // Header: 1 sector (512 bytes)
    // Directory: 1+ sectors (128 bytes per entry, need root + N streams)
    // FAT: 1+ sectors (128 entries per 512-byte sector)
    // Data: ceil(stream_size / sector_size) sectors per stream

    // 1. Build directory entries
    let mut entries = vec![DirEntry::root(total_mini_stream_size)];
    for (name, data) in streams {
        entries.push(DirEntry::stream(name, data.len(), start_sector));
    }

    // 2. Allocate sectors and build FAT
    let mut fat = Vec::new();
    // ... allocate data sectors, directory sectors, FAT sectors

    // 3. Write header
    let mut buf = Vec::with_capacity(estimated_size);
    write_ole2_header(&mut buf, &fat, first_dir_sector, ...)?;

    // 4. Write data sectors
    for (_, data) in streams {
        write_sectors(&mut buf, data, sector_size);
    }

    // 5. Write directory sectors
    write_directory_sectors(&mut buf, &entries, sector_size);

    // 6. Write FAT sectors
    write_fat_sectors(&mut buf, &fat, sector_size);

    Ok(buf)
}
```

### Directory entry builder:

```rust
impl DirEntry {
    fn root(mini_stream_size: u64) -> Self { ... }

    fn stream(name: &str, size: usize, start_sector: u32) -> Self { ... }

    /// Serializes a directory entry to its 128-byte representation.
    fn to_bytes(&self) -> [u8; 128] {
        let mut buf = [0u8; 128];
        // Bytes 0-63: name as UTF-16LE, null-padded
        let name_utf16: Vec<u16> = self.name.encode_utf16().collect();
        for (i, &c) in name_utf16.iter().enumerate() {
            buf[i * 2] = c as u8;
            buf[i * 2 + 1] = (c >> 8) as u8;
        }
        // Byte 64-65: name length (including null terminator, in bytes)
        let name_len = ((name_utf16.len() + 1) * 2) as u16;
        buf[64..66].copy_from_slice(&name_len.to_le_bytes());
        // Byte 66: entry type
        buf[66] = self.entry_type;
        // Byte 67: color (0 = red, 1 = black) — always black for simplicity
        buf[67] = 1;
        // Bytes 68-71: left sibling (0xFFFFFFFF = none)
        buf[68..72].copy_from_slice(&0xFFFFFFFFu32.to_le_bytes());
        // Bytes 72-75: right sibling
        buf[72..76].copy_from_slice(&0xFFFFFFFFu32.to_le_bytes());
        // Bytes 76-79: child (root entry only)
        buf[76..80].copy_from_slice(&self.child_id.to_le_bytes());
        // Bytes 116-119: start sector
        buf[116..120].copy_from_slice(&self.start_sector.to_le_bytes());
        // Bytes 120-127: size (64-bit for v4, 32-bit for v3)
        buf[120..124].copy_from_slice(&(self.size as u32).to_le_bytes());
        buf
    }
}
```

### Tests (5 Rust tests):

1. **test_write_ole2_roundtrip**: Write OLE2 with 2 streams → parse back → verify stream names and content
2. **test_write_ole2_empty_stream**: Single empty stream → valid OLE2 structure
3. **test_write_ole2_large_stream**: Stream > sector size → multiple sectors, verify FAT chain
4. **test_ole2_sector_chain_integrity**: Write → parse → follow FAT chain → all sectors accounted for
5. **test_ole2_directory_utf16**: Stream name with unicode → UTF-16LE encoding correct in directory

---

## Task 8 (0.6.7): File Encryption (Write Path)

**Files:**
- Modify: `crates/modern-xlsx-core/src/ole2/crypto.rs` (add encryption functions)
- Modify: `crates/modern-xlsx-core/src/writer.rs`
- Modify: `crates/modern-xlsx-wasm/src/lib.rs` (add password parameter to write)
- Modify: `packages/modern-xlsx/src/wasm-loader.ts` (add WriteOptions)
- Modify: `packages/modern-xlsx/src/workbook.ts` (toBuffer password option)
- Modify: `packages/modern-xlsx/src/types.ts` (WriteOptions type)
- Modify: `crates/modern-xlsx-core/Cargo.toml` (add getrandom, base64)

### Encryption primitives:

```rust
use getrandom::fill as getrandom_fill;

/// AES-256-CBC encryption with PKCS#7 padding.
pub fn aes_cbc_encrypt(key: &[u8], iv: &[u8], data: &[u8]) -> Result<Vec<u8>> {
    let encryptor = Aes256CbcEnc::new_from_slices(key, iv)
        .map_err(|e| ModernXlsxError::PasswordProtected(format!("AES init: {e}")))?;

    Ok(encryptor.encrypt_padded_vec::<aes::cipher::block_padding::Pkcs7>(data))
}

/// Generates cryptographically secure random bytes.
pub fn secure_random(len: usize) -> Result<Vec<u8>> {
    let mut buf = vec![0u8; len];
    getrandom_fill(&mut buf)
        .map_err(|e| ModernXlsxError::PasswordProtected(format!("RNG error: {e}")))?;
    Ok(buf)
}
```

### Full encryption pipeline:

```rust
/// Encrypts ZIP bytes into an OLE2 compound document.
pub fn encrypt_file(zip_bytes: &[u8], password: &str) -> Result<Vec<u8>> {
    // 1. Generate random salts
    let key_salt = secure_random(16)?;
    let pw_salt = secure_random(16)?;
    let spin_count: u32 = 100_000;
    let key_bits: u32 = 256;

    // 2. Generate random data encryption key
    let mut data_key = secure_random(key_bits as usize / 8)?;

    // 3. Derive password keys and encrypt the data key
    let enc_key = derive_key(
        password, &pw_salt, spin_count, key_bits,
        &BLOCK_KEY_ENCRYPTED_KEY, "SHA512",
    )?;
    let encrypted_key_value = aes_cbc_encrypt(&enc_key, &pw_salt, &data_key)?;

    // 4. Generate and encrypt verifier
    let verifier_input = secure_random(16)?;
    let verifier_hash = {
        let mut h = Sha512::new();
        h.update(&verifier_input);
        h.finalize().to_vec()
    };

    let input_key = derive_key(
        password, &pw_salt, spin_count, key_bits,
        &BLOCK_KEY_VERIFIER_INPUT, "SHA512",
    )?;
    let encrypted_verifier_input = aes_cbc_encrypt(&input_key, &pw_salt, &verifier_input)?;

    let value_key = derive_key(
        password, &pw_salt, spin_count, key_bits,
        &BLOCK_KEY_VERIFIER_VALUE, "SHA512",
    )?;
    let encrypted_verifier_hash = aes_cbc_encrypt(&value_key, &pw_salt, &verifier_hash)?;

    // 5. Encrypt the ZIP package in 4096-byte segments
    let encrypted_package = encrypt_package(&data_key, &key_salt, zip_bytes)?;

    // 6. Compute HMAC
    let (encrypted_hmac_key, encrypted_hmac_value) = compute_and_encrypt_hmac(
        &data_key, &key_salt, &encrypted_package, key_bits,
    )?;

    // 7. Build EncryptionInfo XML
    let enc_info = build_agile_encryption_info_xml(
        &key_salt, &pw_salt, spin_count, key_bits,
        &encrypted_key_value, &encrypted_verifier_input,
        &encrypted_verifier_hash, &encrypted_hmac_key, &encrypted_hmac_value,
    );

    // 8. Build EncryptionInfo stream (version header + XML)
    let mut enc_info_stream = Vec::new();
    enc_info_stream.extend_from_slice(&4u16.to_le_bytes());  // version major
    enc_info_stream.extend_from_slice(&4u16.to_le_bytes());  // version minor
    enc_info_stream.extend_from_slice(&0x40u32.to_le_bytes()); // flags (Agile)
    enc_info_stream.extend_from_slice(enc_info.as_bytes());

    // 9. Write OLE2 compound document
    let ole2 = write_ole2(&[
        ("EncryptionInfo", &enc_info_stream),
        ("EncryptedPackage", &encrypted_package),
    ])?;

    // Zeroize sensitive data
    data_key.zeroize();

    Ok(ole2)
}

fn encrypt_package(data_key: &[u8], salt: &[u8], zip_bytes: &[u8]) -> Result<Vec<u8>> {
    let mut result = Vec::with_capacity(8 + zip_bytes.len() + 4096);

    // Write original size
    result.extend_from_slice(&(zip_bytes.len() as u64).to_le_bytes());

    // Encrypt in 4096-byte segments
    let block_size = 16;
    for (idx, chunk) in zip_bytes.chunks(SEGMENT_SIZE).enumerate() {
        let iv = derive_segment_iv(salt, idx as u32, block_size);
        // Pad last chunk to 4096 if needed
        let padded = if chunk.len() < SEGMENT_SIZE {
            let mut p = chunk.to_vec();
            p.resize(SEGMENT_SIZE, 0);
            p
        } else {
            chunk.to_vec()
        };
        let encrypted = aes_cbc_encrypt(data_key, &iv, &padded)?;
        result.extend_from_slice(&encrypted);
    }

    Ok(result)
}
```

### Writer integration:

```rust
pub struct WriteOptions {
    pub password: Option<String>,
}

pub fn write_xlsx_with_password(workbook: &WorkbookData, options: &WriteOptions) -> Result<Vec<u8>> {
    let zip_bytes = write_xlsx(workbook)?;

    if let Some(ref password) = options.password {
        encrypt_file(&zip_bytes, password)
    } else {
        Ok(zip_bytes)
    }
}
```

### WASM bridge:

```rust
#[wasm_bindgen]
pub fn write_with_password(json: &str, password: &str) -> Result<Uint8Array, JsError> {
    let workbook: WorkbookData = serde_json::from_str(json)
        .map_err(|e| JsError::new(&format!("JSON parse error: {e}")))?;
    let options = WriteOptions {
        password: if password.is_empty() { None } else { Some(password.to_string()) },
    };
    let bytes = modern_xlsx_core::writer::write_xlsx_with_password(&workbook, &options)
        .map_err(|e| JsError::new(&e.to_string()))?;
    Ok(Uint8Array::from(&bytes[..]))
}
```

### TypeScript API:

```typescript
// types.ts
export interface WriteOptions {
    /** Password to encrypt the XLSX file. */
    password?: string;
}

// workbook.ts
async toBuffer(options?: WriteOptions): Promise<Uint8Array> {
    await ensureReady();
    const password = options?.password ?? '';
    return password
        ? wasmWriteWithPassword(this.data, password)
        : wasmWrite(this.data);
}
```

### Tests (5 integration tests):

1. **test_encrypt_roundtrip**: Create workbook → `toBuffer({ password: 'test' })` → `readBuffer(buf, { password: 'test' })` → verify data matches
2. **test_encrypt_wrong_password_fails**: Encrypt → read with wrong password → "Incorrect password"
3. **test_encrypt_no_password_shows_info**: Encrypt → `readBuffer(buf)` without password → "password required"
4. **test_encrypt_empty_password_works**: `toBuffer({ password: '' })` → normal unencrypted output (empty = no encryption)
5. **test_encrypt_preserves_blob_api**: `writeBlob(wb, { password: 'test' })` → encrypted Blob

### WASM rebuild.

---

## Task 9 (0.6.8): Encryption Roundtrip & Compatibility

**Files:**
- Test file: `packages/modern-xlsx/__tests__/encryption-roundtrip.test.ts`

### Tests (10):

1. **encrypt → decrypt roundtrip with styles**: Workbook with bold/italic/colors → encrypt → decrypt → styles preserved
2. **encrypt → decrypt roundtrip with formulas**: Formulas, named ranges → encrypt → decrypt → formulas preserved
3. **encrypt → decrypt roundtrip with tables**: Excel tables → encrypt → decrypt → tables preserved
4. **encrypt → decrypt roundtrip with sparklines**: Sparklines → encrypt → decrypt → sparklines preserved
5. **encrypt → decrypt roundtrip with merged cells**: Merge regions → encrypt → decrypt → merges preserved
6. **password change**: Read encrypted(pw1) → write encrypted(pw2) → read with pw2 → verify data
7. **large file encryption performance**: 10K rows → encrypt → decrypt → under 5 seconds
8. **encrypt → decrypt with unicode password**: Password "密码テスト" → roundtrip works
9. **encrypt → decrypt with long password**: 100-character password → roundtrip works
10. **multiple sheets encrypted**: 5 sheets with different features → encrypt → decrypt → all preserved

---

## Task 10 (0.6.9): Security Hardening

**Files:**
- Modify: `crates/modern-xlsx-core/src/ole2/crypto.rs` (zeroize all key material)
- Create: `crates/modern-xlsx-core/tests/security_tests.rs`

### Hardening changes:

```rust
use zeroize::Zeroize;

// Add Zeroize to all key types:
#[derive(Zeroize)]
#[zeroize(drop)]
struct SensitiveKey(Vec<u8>);

// Wrap all derived keys in SensitiveKey for auto-zeroize on drop
// Ensure no log statements can leak key material
// Use constant_time_eq for ALL comparisons of secret data
```

### Audit checklist:
1. Grep for `log::` and `println!` in ole2/ — none should reference keys, passwords, or derived values
2. All `Vec<u8>` holding keys use `Zeroize`
3. All hash comparisons use `constant_time_eq`
4. Salt generation uses `getrandom` (not `rand` or similar)
5. No partial decryption state in error messages

### Tests (5):

1. **test_no_key_in_error_messages**: Trigger various decryption errors → scan error strings for key-like patterns
2. **test_constant_time_comparison_used**: (Code audit, verified in review)
3. **test_malformed_encryption_info_no_panic**: Fuzz with truncated/corrupted EncryptionInfo → no panics, only errors
4. **test_truncated_encrypted_package_no_panic**: Truncated at various offsets → graceful error
5. **test_empty_encrypted_package_no_panic**: Zero-length EncryptedPackage → clear error

### Final WASM rebuild + full verification.

---

## Post-Implementation: Full Verification

```bash
cargo test -p modern-xlsx-core
cargo clippy -p modern-xlsx-core -- -D warnings
pnpm -C packages/modern-xlsx lint
pnpm -C packages/modern-xlsx typecheck
pnpm -C packages/modern-xlsx build
pnpm -C packages/modern-xlsx test
```

Expected: ~305 Rust tests, ~600 TypeScript tests.
