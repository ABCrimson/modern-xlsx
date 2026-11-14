//! `EncryptionInfo` stream parser (MS-OFFCRYPTO 2.3.4.5 Standard, 2.3.4.10 Agile).
//!
//! Parsing validates the descriptor fields listed on [`EncryptionInfo::parse`],
//! so an unsupported cipher, a bad size or a hostile spinCount fails here,
//! before any key derivation. The module uses no crypto crate.

use core::fmt;
use core::hint::cold_path;
use core::str::FromStr;
use std::borrow::Cow;

use super::base64::decode_base64;
use super::cfb::reader::Cfb;
use crate::Result;
use crate::errors::ModernXlsxError;
use quick_xml::Reader;
use quick_xml::events::{BytesStart, Event, attributes::Attribute};

/// Largest `spinCount` MS-OFFCRYPTO 2.3.4.10 allows (`ST_SpinCount` maxInclusive).
pub(crate) const MAX_SPIN_COUNT: u32 = 10_000_000;

/// `uri` of the password key encryptor (MS-OFFCRYPTO 2.3.4.10).
const PASSWORD_KEY_ENCRYPTOR: &str =
    "http://schemas.microsoft.com/office/2006/keyEncryptor/password";

/// The AES block size, the only `blockSize` accepted.
const AES_BLOCK_SIZE: u32 = 16;

/// `EncryptionHeader.Flags` bits (MS-OFFCRYPTO 2.3.1).
const F_CRYPTO_API: u32 = 0x04;
const F_EXTERNAL: u32 = 0x10;
const F_AES: u32 = 0x20;
/// Flags every supported Standard (AES CryptoAPI) header sets.
const F_REQUIRED: u32 = F_CRYPTO_API | F_AES;

/// `EncryptionHeader.AlgID` and `AlgIDHash` values (MS-OFFCRYPTO 2.3.2).
const ALG_ID_RC4: u32 = 0x6801;
const ALG_ID_AES_128: u32 = 0x660E;
const ALG_ID_AES_192: u32 = 0x660F;
const ALG_ID_AES_256: u32 = 0x6610;
const ALG_ID_HASH_SHA1: u32 = 0x8004;

/// Fixed part of `EncryptionHeader`: Flags, SizeExtra, AlgID, AlgIDHash,
/// KeySize, ProviderType, Reserved1, Reserved2 (4 bytes each).
const STANDARD_HEADER_FIXED_LEN: usize = 32;

/// AES `EncryptionVerifier` (MS-OFFCRYPTO 2.3.3): SaltSize (4), Salt (16),
/// EncryptedVerifier (16), VerifierHashSize (4), EncryptedVerifierHash (32).
const STANDARD_VERIFIER_LEN: usize = 72;
const STANDARD_SALT_SIZE: u32 = 16;
const STANDARD_VERIFIER_HASH_SIZE: u32 = 20;

/// Combined encryption info parsed from the EncryptionInfo stream.
#[derive(Debug, Clone)]
pub enum EncryptionInfo {
    Standard(StandardEncryptionInfo),
    Agile(Box<AgileEncryptionInfo>),
}

/// Standard Encryption info (version 2.2, 3.2, 4.2), AES only.
#[derive(Debug, Clone)]
pub struct StandardEncryptionInfo {
    /// `0x660E`, `0x660F` or `0x6610`; a stored 0 with fAES set reads as `0x660E`.
    pub alg_id: u32,
    /// Always `0x8004` (SHA-1); a stored 0 reads as `0x8004`.
    pub hash_alg_id: u32,
    /// 128, 192 or 256, matching `alg_id`.
    pub key_size: u32,
    pub salt: Vec<u8>,
    pub encrypted_verifier: Vec<u8>,
    pub verifier_hash_size: u32,
    pub encrypted_verifier_hash: Vec<u8>,
}

/// Parsed Agile encryption descriptor (version 4.4).
#[derive(Debug, Clone)]
pub struct AgileEncryptionInfo {
    // Key data
    pub key_salt: Vec<u8>,
    pub key_block_size: u32,
    pub key_bits: u32,
    pub key_hash_size: u32,
    pub key_hash_alg: HashAlg,
    // Data integrity
    pub encrypted_hmac_key: Vec<u8>,
    pub encrypted_hmac_value: Vec<u8>,
    // Password key encryptor
    pub pw_spin_count: u32,
    pub pw_salt: Vec<u8>,
    pub pw_key_bits: u32,
    pub pw_hash_size: u32,
    pub pw_hash_alg: HashAlg,
    pub pw_encrypted_key_value: Vec<u8>,
    pub pw_encrypted_verifier_hash_input: Vec<u8>,
    pub pw_encrypted_verifier_hash_value: Vec<u8>,
}

/// Hash algorithm named by an Agile `hashAlgorithm` attribute (MS-OFFCRYPTO 2.3.4.10).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum HashAlg {
    Sha1,
    Sha256,
    Sha384,
    Sha512,
}

impl HashAlg {
    /// Digest length in bytes: 20, 32, 48 or 64.
    pub const fn digest_len(self) -> usize {
        match self {
            Self::Sha1 => 20,
            Self::Sha256 => 32,
            Self::Sha384 => 48,
            Self::Sha512 => 64,
        }
    }

    /// The `hashAlgorithm` spelling this crate writes: `SHA1`, `SHA256`, `SHA384` or `SHA512`.
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::Sha1 => "SHA1",
            Self::Sha256 => "SHA256",
            Self::Sha384 => "SHA384",
            Self::Sha512 => "SHA512",
        }
    }
}

impl FromStr for HashAlg {
    type Err = ModernXlsxError;

    /// Accepts `SHA1`/`SHA-1`, `SHA256`/`SHA-256`, `SHA384`/`SHA-384` and
    /// `SHA512`/`SHA-512`; anything else is a `PasswordProtected` error.
    fn from_str(s: &str) -> Result<Self> {
        match s {
            "SHA1" | "SHA-1" => Ok(Self::Sha1),
            "SHA256" | "SHA-256" => Ok(Self::Sha256),
            "SHA384" | "SHA-384" => Ok(Self::Sha384),
            "SHA512" | "SHA-512" => Ok(Self::Sha512),
            _ => Err(unsupported(format_args!("hash algorithm '{s}'"))),
        }
    }
}

impl TryFrom<&str> for HashAlg {
    type Error = ModernXlsxError;

    /// Same as [`FromStr`]: `HashAlg::try_from(s)` and `s.parse::<HashAlg>()` agree.
    fn try_from(s: &str) -> Result<Self> {
        s.parse()
    }
}

impl fmt::Display for HashAlg {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.write_str(self.as_str())
    }
}

impl EncryptionInfo {
    /// Parse the raw EncryptionInfo stream bytes.
    ///
    /// Version 4.4 is Agile; versions 2.2, 3.2 and 4.2 are Standard. Validated:
    /// - Agile `keyData` and password `encryptedKey`: `keyBits` in {128, 192, 256},
    ///   `cipherAlgorithm` "AES", `cipherChaining` "ChainingModeCBC", `blockSize` 16,
    ///   `saltSize` equal to the decoded `saltValue` length, a known `hashAlgorithm`
    ///   and `hashSize` equal to its digest length. Only the encryptor whose `uri`
    ///   is the password key encryptor is read; at least one must be present.
    /// - Agile `spinCount`: present and at most `MAX_SPIN_COUNT` (10,000,000).
    /// - Standard `EncryptionHeader`: fCryptoAPI and fAES set, fExternal clear,
    ///   AES `AlgID` (0 reads as AES-128), `AlgIDHash` SHA-1 or 0, and `KeySize`
    ///   matching `AlgID`.
    /// - Standard `EncryptionVerifier`: `SaltSize` 16 and `VerifierHashSize` 20.
    ///
    /// A missing or non-numeric attribute is a `PasswordProtected` "malformed
    /// EncryptionInfo" error; invalid base64 or XML is a `PasswordProtected`
    /// error with its own message; an unsupported cipher is "unsupported
    /// encryption"; a `spinCount` above the cap, including a decimal past
    /// `u32`, is a `Security` error.
    pub fn parse(stream: &[u8]) -> Result<Self> {
        if stream.len() < 8 {
            cold_path();
            return Err(ModernXlsxError::PasswordProtected(format!(
                "EncryptionInfo stream too short: got {} bytes, need at least 8 — the file may be corrupt",
                stream.len()
            )));
        }
        let version_major = u16::from_le_bytes([stream[0], stream[1]]);
        let version_minor = u16::from_le_bytes([stream[2], stream[3]]);

        match (version_major, version_minor) {
            (4, 4) => Self::parse_agile(&stream[8..]),
            (2..=4, 2) => Self::parse_standard(&stream[8..]),
            _ => {
                cold_path();
                Err(ModernXlsxError::PasswordProtected(format!(
                    "Unsupported encryption version {version_major}.{version_minor} — only Standard (2.2/3.2/4.2) and Agile (4.4) are supported"
                )))
            }
        }
    }

    fn parse_agile(xml_bytes: &[u8]) -> Result<Self> {
        let mut key_data: Option<CipherAttrs> = None;
        let mut password: Option<PasswordAttrs> = None;
        let mut in_password_encryptor = false;
        let mut encrypted_hmac_key = Vec::new();
        let mut encrypted_hmac_value = Vec::new();

        let mut reader = Reader::from_reader(xml_bytes);
        reader.config_mut().trim_text(true);
        let mut buf = Vec::with_capacity(512);

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == "keyEncryptor" => {
                    in_password_encryptor = false;
                    for attr in e.attributes() {
                        let attr = attr.map_err(|err| malformed(format_args!("{err}")))?;
                        if attr.key.as_ref() == "uri" {
                            let val = attr_value(&attr)?;
                            if val == PASSWORD_KEY_ENCRYPTOR {
                                in_password_encryptor = true;
                            }
                        }
                    }
                }
                Ok(Event::End(ref e)) if e.local_name().as_ref() == "keyEncryptor" => {
                    in_password_encryptor = false;
                }
                Ok(Event::Empty(ref e) | Event::Start(ref e)) => {
                    match e.local_name().as_ref() {
                        "keyData" if key_data.is_none() => {
                            key_data = Some(CipherAttrs::read("keyData", e)?);
                        }
                        "dataIntegrity" => {
                            for attr in e.attributes() {
                                let attr = attr.map_err(|err| malformed(format_args!("{err}")))?;
                                match attr.key.as_ref() {
                                    "encryptedHmacKey" => {
                                        encrypted_hmac_key = decode_base64(&attr_value(&attr)?)?;
                                    }
                                    "encryptedHmacValue" => {
                                        encrypted_hmac_value = decode_base64(&attr_value(&attr)?)?;
                                    }
                                    _ => {}
                                }
                            }
                        }
                        // A certificate encryptor's `c:encryptedKey` shares the
                        // local name; only the password encryptor's is read.
                        "encryptedKey" if in_password_encryptor && password.is_none() => {
                            password = Some(PasswordAttrs::read(e)?);
                        }
                        _ => {}
                    }
                }
                Ok(Event::Eof) => break,
                Err(e) => {
                    cold_path();
                    return Err(ModernXlsxError::PasswordProtected(format!(
                        "Failed to parse Agile encryption XML: {e}"
                    )));
                }
                _ => {}
            }
            buf.clear();
        }

        let Some(password) = password else {
            return Err(malformed("no password key encryptor"));
        };
        // Checked before anything else so a hostile count surfaces as SECURITY.
        let pw_spin_count = required("encryptedKey", "spinCount", password.spin_count)?;
        if pw_spin_count > MAX_SPIN_COUNT {
            cold_path();
            return Err(ModernXlsxError::Security(format!(
                "EncryptionInfo spinCount {pw_spin_count} exceeds the MS-OFFCRYPTO maximum of {MAX_SPIN_COUNT}"
            )));
        }
        let key = key_data.unwrap_or_default().finish("keyData")?;
        let pw = password.cipher.finish("encryptedKey")?;

        Ok(EncryptionInfo::Agile(Box::new(AgileEncryptionInfo {
            key_salt: key.salt,
            key_block_size: key.block_size,
            key_bits: key.key_bits,
            key_hash_size: key.hash_size,
            key_hash_alg: key.hash_alg,
            encrypted_hmac_key,
            encrypted_hmac_value,
            pw_spin_count,
            pw_salt: pw.salt,
            pw_key_bits: pw.key_bits,
            pw_hash_size: pw.hash_size,
            pw_hash_alg: pw.hash_alg,
            pw_encrypted_key_value: password.encrypted_key_value,
            pw_encrypted_verifier_hash_input: password.encrypted_verifier_hash_input,
            pw_encrypted_verifier_hash_value: password.encrypted_verifier_hash_value,
        })))
    }

    fn parse_standard(data: &[u8]) -> Result<Self> {
        // After parse() skips 8 bytes (version 4 + flags 4), we receive:
        //   data[0..4]  = headerSize (u32)
        //   data[4..4+headerSize] = EncryptionHeader (MS-OFFCRYPTO 2.3.2)
        //     header[0..4]  = flags
        //     header[4..8]  = sizeExtra (must be 0)
        //     header[8..12] = algID
        //     header[12..16] = algIDHash
        //     header[16..20] = keySize (bits)
        //     header[20..24] = providerType
        //     header[24..28] = reserved1
        //     header[28..32] = reserved2
        //     header[32..]  = CSP name (UTF-16LE, null-terminated)
        //   data[4+headerSize..] = EncryptionVerifier (MS-OFFCRYPTO 2.3.3)
        //     saltSize (4), salt (16), encryptedVerifier (16),
        //     verifierHashSize (4), encryptedVerifierHash (32)
        let Some((size, rest)) = data.split_first_chunk::<4>() else {
            cold_path();
            return Err(ModernXlsxError::PasswordProtected(
                "Standard encryption header too short".into(),
            ));
        };
        // `usize::try_from` and `split_at_checked` keep a hostile headerSize near
        // u32::MAX from wrapping on wasm32; the verifier must follow in full.
        let Some((header, v)) = usize::try_from(u32::from_le_bytes(*size))
            .ok()
            .and_then(|header_size| rest.split_at_checked(header_size))
            .and_then(|(header, verifier)| {
                Some((header, verifier.first_chunk::<STANDARD_VERIFIER_LEN>()?))
            })
        else {
            cold_path();
            return Err(ModernXlsxError::PasswordProtected(
                "Standard encryption data too short for verifier".into(),
            ));
        };

        let Some((fixed, _)) = header.split_first_chunk::<STANDARD_HEADER_FIXED_LEN>() else {
            cold_path();
            return Err(ModernXlsxError::PasswordProtected(
                "Standard encryption header fields too short".into(),
            ));
        };
        let (words, _) = fixed.as_chunks::<4>();
        let flags = u32::from_le_bytes(words[0]);
        let alg_id = u32::from_le_bytes(words[2]);
        let hash_alg_id = u32::from_le_bytes(words[3]);
        let key_size = u32::from_le_bytes(words[4]);

        if flags & F_EXTERNAL != 0 {
            return Err(unsupported("extensible encryption (fExternal)"));
        }
        let alg_id = match alg_id {
            ALG_ID_AES_128 | ALG_ID_AES_192 | ALG_ID_AES_256 => alg_id,
            0 if flags & F_AES != 0 => ALG_ID_AES_128,
            0 | ALG_ID_RC4 => {
                return Err(unsupported(format_args!(
                    "RC4 CryptoAPI (algID {alg_id:#06X})"
                )));
            }
            other => {
                return Err(unsupported(format_args!("algID {other:#06X}")));
            }
        };
        if flags & F_REQUIRED != F_REQUIRED {
            return Err(malformed(format_args!(
                "Standard flags {flags:#04X} must set fCryptoAPI and fAES"
            )));
        }
        let hash_alg_id = match hash_alg_id {
            ALG_ID_HASH_SHA1 | 0 => ALG_ID_HASH_SHA1,
            other => {
                return Err(unsupported(format_args!("algIDHash {other:#06X}")));
            }
        };
        let expected_key_size = match alg_id {
            ALG_ID_AES_128 => 128,
            ALG_ID_AES_192 => 192,
            _ => 256,
        };
        if key_size != expected_key_size {
            return Err(malformed(format_args!(
                "keySize {key_size} does not match algID {alg_id:#06X}"
            )));
        }

        let (v_words, _) = v.as_chunks::<4>();
        let salt_size = u32::from_le_bytes(v_words[0]);
        if salt_size != STANDARD_SALT_SIZE {
            return Err(malformed(format_args!(
                "verifier saltSize {salt_size} (expected {STANDARD_SALT_SIZE})"
            )));
        }
        let verifier_hash_size = u32::from_le_bytes(v_words[9]);
        if verifier_hash_size != STANDARD_VERIFIER_HASH_SIZE {
            return Err(malformed(format_args!(
                "verifierHashSize {verifier_hash_size} (expected {STANDARD_VERIFIER_HASH_SIZE})"
            )));
        }

        Ok(EncryptionInfo::Standard(StandardEncryptionInfo {
            alg_id,
            hash_alg_id,
            key_size,
            salt: v[4..20].to_vec(),
            encrypted_verifier: v[20..36].to_vec(),
            verifier_hash_size,
            encrypted_verifier_hash: v[40..72].to_vec(),
        }))
    }
}

/// Raw attributes shared by `keyData` and the password `encryptedKey`
/// (`CT_KeyData`, `CT_PasswordKeyEncryptor`); [`CipherAttrs::finish`] validates them.
#[derive(Default)]
struct CipherAttrs {
    salt_size: Option<u32>,
    block_size: Option<u32>,
    key_bits: Option<u32>,
    hash_size: Option<u32>,
    cipher: Option<String>,
    chaining: Option<String>,
    hash_alg: Option<String>,
    salt: Option<Vec<u8>>,
}

/// Validated cipher parameters of one `keyData` or `encryptedKey` element.
struct CipherParams {
    salt: Vec<u8>,
    block_size: u32,
    key_bits: u32,
    hash_size: u32,
    hash_alg: HashAlg,
}

impl CipherAttrs {
    fn read(elem: &str, e: &BytesStart<'_>) -> Result<Self> {
        let mut attrs = Self::default();
        for attr in e.attributes() {
            let attr = attr.map_err(|err| malformed(format_args!("{err}")))?;
            attrs.set(elem, attr.key.as_ref(), &attr_value(&attr)?)?;
        }
        Ok(attrs)
    }

    /// Stores a shared cipher attribute; returns `false` for any other key.
    fn set(&mut self, elem: &str, key: &str, val: &str) -> Result<bool> {
        match key {
            "saltSize" => self.salt_size = Some(parse_u32(elem, key, val)?),
            "blockSize" => self.block_size = Some(parse_u32(elem, key, val)?),
            "keyBits" => self.key_bits = Some(parse_u32(elem, key, val)?),
            "hashSize" => self.hash_size = Some(parse_u32(elem, key, val)?),
            "cipherAlgorithm" => self.cipher = Some(val.to_owned()),
            "cipherChaining" => self.chaining = Some(val.to_owned()),
            "hashAlgorithm" => self.hash_alg = Some(val.to_owned()),
            "saltValue" => self.salt = Some(decode_base64(val)?),
            _ => return Ok(false),
        }
        Ok(true)
    }

    fn finish(self, elem: &str) -> Result<CipherParams> {
        let key_bits = required(elem, "keyBits", self.key_bits)?;
        if !matches!(key_bits, 128 | 192 | 256) {
            return Err(unsupported(format_args!("key size {key_bits} bits")));
        }
        let cipher = required(elem, "cipherAlgorithm", self.cipher)?;
        if cipher != "AES" {
            return Err(unsupported(format_args!("cipherAlgorithm '{cipher}'")));
        }
        let chaining = required(elem, "cipherChaining", self.chaining)?;
        if chaining != "ChainingModeCBC" {
            return Err(unsupported(&chaining));
        }
        let block_size = required(elem, "blockSize", self.block_size)?;
        if block_size != AES_BLOCK_SIZE {
            return Err(unsupported(format_args!(
                "blockSize {block_size} (expected {AES_BLOCK_SIZE})"
            )));
        }
        let hash_alg = HashAlg::try_from(required(elem, "hashAlgorithm", self.hash_alg)?.as_str())?;
        let hash_size = required(elem, "hashSize", self.hash_size)?;
        if usize::try_from(hash_size).ok() != Some(hash_alg.digest_len()) {
            return Err(malformed(format_args!(
                "hashSize {hash_size} on {elem} does not match {hash_alg} ({} bytes)",
                hash_alg.digest_len()
            )));
        }
        let salt = required(elem, "saltValue", self.salt)?;
        let salt_size = required(elem, "saltSize", self.salt_size)?;
        if usize::try_from(salt_size).ok() != Some(salt.len()) {
            return Err(malformed(format_args!(
                "saltSize {salt_size} on {elem} does not match the {}-byte saltValue",
                salt.len()
            )));
        }
        Ok(CipherParams {
            salt,
            block_size,
            key_bits,
            hash_size,
            hash_alg,
        })
    }
}

/// Raw attributes of the password key encryptor's `encryptedKey` element.
#[derive(Default)]
struct PasswordAttrs {
    cipher: CipherAttrs,
    spin_count: Option<u32>,
    encrypted_key_value: Vec<u8>,
    encrypted_verifier_hash_input: Vec<u8>,
    encrypted_verifier_hash_value: Vec<u8>,
}

impl PasswordAttrs {
    fn read(e: &BytesStart<'_>) -> Result<Self> {
        const ELEM: &str = "encryptedKey";
        let mut attrs = Self::default();
        for attr in e.attributes() {
            let attr = attr.map_err(|err| malformed(format_args!("{err}")))?;
            let key = attr.key.as_ref();
            let val = attr_value(&attr)?;
            let val: &str = &val;
            if attrs.cipher.set(ELEM, key, val)? {
                continue;
            }
            match key {
                "spinCount" => attrs.spin_count = Some(parse_spin_count(val)?),
                "encryptedKeyValue" => attrs.encrypted_key_value = decode_base64(val)?,
                "encryptedVerifierHashInput" => {
                    attrs.encrypted_verifier_hash_input = decode_base64(val)?;
                }
                "encryptedVerifierHashValue" => {
                    attrs.encrypted_verifier_hash_value = decode_base64(val)?;
                }
                _ => {}
            }
        }
        Ok(attrs)
    }
}

/// The attribute's value with XML attribute-value normalization applied
/// (references resolved, tab, CR and LF read as spaces); an undeclared entity
/// or a bad character reference is a malformed error.
fn attr_value<'a>(attr: &Attribute<'a>) -> Result<Cow<'a, str>> {
    crate::ooxml::xml::attr_str(attr).map_err(|e| malformed(format_args!("{e}")))
}

/// Parses an unsigned decimal attribute; anything else is a malformed error.
fn parse_u32(elem: &str, attr: &str, val: &str) -> Result<u32> {
    val.parse().map_err(|_| {
        malformed(format_args!(
            "{attr} '{val}' on {elem} is not an unsigned integer"
        ))
    })
}

/// Parses `spinCount`. A decimal too large for `u32` saturates to `u32::MAX`,
/// so it fails the [`MAX_SPIN_COUNT`] check as SECURITY like any other
/// over-cap value; anything that is not a decimal is a malformed error.
fn parse_spin_count(val: &str) -> Result<u32> {
    if !val.is_empty() && val.bytes().all(|b| b.is_ascii_digit()) {
        Ok(val.parse().unwrap_or(u32::MAX))
    } else {
        parse_u32("encryptedKey", "spinCount", val)
    }
}

/// Unwraps a required attribute; a missing one is a malformed error.
fn required<T>(elem: &str, attr: &str, value: Option<T>) -> Result<T> {
    value.ok_or_else(|| malformed(format_args!("{attr} missing on {elem}")))
}

#[cold]
fn unsupported(what: impl fmt::Display) -> ModernXlsxError {
    ModernXlsxError::PasswordProtected(format!("unsupported encryption: {what}"))
}

#[cold]
fn malformed(what: impl fmt::Display) -> ModernXlsxError {
    ModernXlsxError::PasswordProtected(format!("malformed EncryptionInfo: {what}"))
}

/// Describe the encryption method for error messages.
fn describe_encryption(info: &EncryptionInfo) -> String {
    match info {
        EncryptionInfo::Agile(a) => {
            format!("Agile encryption (AES-{}, {})", a.key_bits, a.key_hash_alg)
        }
        EncryptionInfo::Standard(s) => {
            format!("Standard encryption (key size: {} bits)", s.key_size)
        }
    }
}

/// Builds the error for an encrypted XLSX file that will not be decrypted: one read without a password, or any encrypted file in the lite build (no `encryption` feature), with or without a password.
/// In the lite build the advice says it cannot decrypt ([`crate::ole2::detect::ERR_ENCRYPTED_LITE`]).
///
/// Tries to parse the `EncryptionInfo` stream of the compound file `cfb` (the
/// read path's, which it classified the file with) for a detailed description.
/// A `Security` error (hostile metadata) is returned unchanged; any other
/// failure falls back to the generic message.
pub(crate) fn build_encrypted_error_with_cfb(cfb: &Cfb<'_>) -> ModernXlsxError {
    encrypted_error(
        cfb.read_stream("EncryptionInfo")
            .and_then(|s| EncryptionInfo::parse(&s)),
    )
}

/// Maps the outcome of the `EncryptionInfo` read to the error
/// [`build_encrypted_error_with_cfb`] documents.
fn encrypted_error(info: Result<EncryptionInfo>) -> ModernXlsxError {
    let msg = match info {
        Ok(info) => {
            let desc = describe_encryption(&info);
            #[cfg(feature = "encryption")]
            let advice = "Provide password via readBuffer(data, { password: '...' }).";
            #[cfg(not(feature = "encryption"))]
            let advice = "The lite build cannot decrypt; use the full modern-xlsx entry \
                          and pass readBuffer(data, { password }).";
            format!("Password-protected XLSX ({desc}). {advice}")
        }
        Err(e @ ModernXlsxError::Security(_)) => return e,
        #[cfg(feature = "encryption")]
        Err(_) => crate::ole2::detect::ERR_ENCRYPTED.into(),
        #[cfg(not(feature = "encryption"))]
        Err(_) => crate::ole2::detect::ERR_ENCRYPTED_LITE.into(),
    };
    ModernXlsxError::PasswordProtected(msg)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_parse_agile_aes256_sha512() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<encryption xmlns="http://schemas.microsoft.com/office/2006/encryption" xmlns:p="http://schemas.microsoft.com/office/2006/keyEncryptor/password">
  <keyData saltSize="16" blockSize="16" keyBits="256" hashSize="64" cipherAlgorithm="AES" cipherChaining="ChainingModeCBC" hashAlgorithm="SHA512" saltValue="AAAAAAAAAAAAAAAAAAAAAA=="/>
  <dataIntegrity encryptedHmacKey="AAAAAAAAAAAAAAAAAAAAAA==" encryptedHmacValue="AAAAAAAAAAAAAAAAAAAAAA=="/>
  <keyEncryptors>
    <keyEncryptor uri="http://schemas.microsoft.com/office/2006/keyEncryptor/password">
      <p:encryptedKey spinCount="100000" saltSize="16" blockSize="16" keyBits="256" hashSize="64" cipherAlgorithm="AES" cipherChaining="ChainingModeCBC" hashAlgorithm="SHA512" saltValue="BBBBBBBBBBBBBBBBBBBBBB==" encryptedKeyValue="CCCCCCCCCCCCCCCCCCCCCC==" encryptedVerifierHashInput="DDDDDDDDDDDDDDDDDDDDDD==" encryptedVerifierHashValue="EEEEEEEEEEEEEEEEEEEEEE=="/>
    </keyEncryptor>
  </keyEncryptors>
</encryption>"#;
        let mut stream = Vec::new();
        stream.extend_from_slice(&4u16.to_le_bytes());
        stream.extend_from_slice(&4u16.to_le_bytes());
        stream.extend_from_slice(&0u32.to_le_bytes());
        stream.extend_from_slice(xml.as_bytes());

        let info = EncryptionInfo::parse(&stream).unwrap();
        match info {
            EncryptionInfo::Agile(a) => {
                assert_eq!(a.key_bits, 256);
                assert_eq!(a.key_hash_alg, HashAlg::Sha512);
                assert_eq!(a.key_block_size, 16);
                assert_eq!(a.key_hash_size, 64);
                assert_eq!(a.pw_spin_count, 100000);
                assert_eq!(a.pw_key_bits, 256);
                assert_eq!(a.pw_hash_alg, HashAlg::Sha512);
                assert!(!a.key_salt.is_empty());
                assert!(!a.pw_salt.is_empty());
                assert!(!a.pw_encrypted_key_value.is_empty());
                assert!(!a.pw_encrypted_verifier_hash_input.is_empty());
                assert!(!a.pw_encrypted_verifier_hash_value.is_empty());
                assert!(!a.encrypted_hmac_key.is_empty());
                assert!(!a.encrypted_hmac_value.is_empty());
            }
            other => {
                cold_path();
                unreachable!("Expected Agile encryption, got {other:?}")
            }
        }
    }

    #[test]
    fn test_parse_agile_aes128() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<encryption xmlns="http://schemas.microsoft.com/office/2006/encryption" xmlns:p="http://schemas.microsoft.com/office/2006/keyEncryptor/password">
  <keyData saltSize="16" blockSize="16" keyBits="128" hashSize="32" cipherAlgorithm="AES" cipherChaining="ChainingModeCBC" hashAlgorithm="SHA256" saltValue="AAAAAAAAAAAAAAAAAAAAAA=="/>
  <dataIntegrity encryptedHmacKey="AAAAAAAAAAAAAAAAAAAAAA==" encryptedHmacValue="AAAAAAAAAAAAAAAAAAAAAA=="/>
  <keyEncryptors>
    <keyEncryptor uri="http://schemas.microsoft.com/office/2006/keyEncryptor/password">
      <p:encryptedKey spinCount="50000" saltSize="16" blockSize="16" keyBits="128" hashSize="32" cipherAlgorithm="AES" cipherChaining="ChainingModeCBC" hashAlgorithm="SHA256" saltValue="BBBBBBBBBBBBBBBBBBBBBB==" encryptedKeyValue="CCCCCCCCCCCCCCCCCCCCCC==" encryptedVerifierHashInput="DDDDDDDDDDDDDDDDDDDDDD==" encryptedVerifierHashValue="EEEEEEEEEEEEEEEEEEEEEE=="/>
    </keyEncryptor>
  </keyEncryptors>
</encryption>"#;
        let mut stream = Vec::new();
        stream.extend_from_slice(&4u16.to_le_bytes());
        stream.extend_from_slice(&4u16.to_le_bytes());
        stream.extend_from_slice(&0u32.to_le_bytes());
        stream.extend_from_slice(xml.as_bytes());

        let info = EncryptionInfo::parse(&stream).unwrap();
        match info {
            EncryptionInfo::Agile(a) => {
                assert_eq!(a.key_bits, 128);
                assert_eq!(a.key_hash_alg, HashAlg::Sha256);
                assert_eq!(a.pw_spin_count, 50000);
            }
            other => {
                cold_path();
                unreachable!("Expected Agile encryption, got {other:?}")
            }
        }
    }

    #[test]
    fn test_parse_standard_encryption() {
        let mut stream = Vec::new();
        stream.extend_from_slice(&4u16.to_le_bytes()); // major
        stream.extend_from_slice(&2u16.to_le_bytes()); // minor
        // After version header: flags(4) + headerSize(4) + header + verifier
        let flags: u32 = 0x24;
        stream.extend_from_slice(&flags.to_le_bytes());
        let header_size: u32 = 40;
        stream.extend_from_slice(&header_size.to_le_bytes());
        // Header (40 bytes):
        let header_flags: u32 = 0x24;
        stream.extend_from_slice(&header_flags.to_le_bytes()); // flags
        stream.extend_from_slice(&0u32.to_le_bytes()); // sizeExtra
        stream.extend_from_slice(&0x660Eu32.to_le_bytes()); // algID = AES-128
        stream.extend_from_slice(&0x8004u32.to_le_bytes()); // algIDHash = SHA-1
        stream.extend_from_slice(&128u32.to_le_bytes()); // keySize
        stream.extend_from_slice(&0x18u32.to_le_bytes()); // providerType
        stream.extend_from_slice(&0u32.to_le_bytes()); // reserved1
        stream.extend_from_slice(&0u32.to_le_bytes()); // reserved2
        // CSP name (UTF-16LE, null-terminated) — "AES" + null = 8 bytes
        stream.extend_from_slice(&[b'A', 0, b'E', 0, b'S', 0, 0, 0]);
        // Verifier (MS-OFFCRYPTO 2.3.3):
        stream.extend_from_slice(&16u32.to_le_bytes()); // salt size
        stream.extend_from_slice(&[0xAA; 16]); // salt
        stream.extend_from_slice(&[0xBB; 16]); // encrypted verifier
        stream.extend_from_slice(&20u32.to_le_bytes()); // hash size
        stream.extend_from_slice(&[0xCC; 32]); // encrypted verifier hash

        let info = EncryptionInfo::parse(&stream).unwrap();
        match info {
            EncryptionInfo::Standard(s) => {
                assert_eq!(s.alg_id, 0x660E);
                assert_eq!(s.salt, [0xAA; 16]);
                assert_eq!(s.encrypted_verifier, [0xBB; 16]);
                assert_eq!(s.encrypted_verifier_hash, [0xCC; 32]);
                assert_eq!(s.hash_alg_id, 0x8004);
                assert_eq!(s.key_size, 128);
                assert_eq!(s.salt.len(), 16);
                assert_eq!(s.encrypted_verifier.len(), 16);
                assert_eq!(s.verifier_hash_size, 20);
                assert_eq!(s.encrypted_verifier_hash.len(), 32);
            }
            other => {
                cold_path();
                unreachable!("Expected Standard encryption, got {other:?}")
            }
        }
    }

    #[test]
    fn test_unsupported_version_error() {
        let mut stream = Vec::new();
        stream.extend_from_slice(&99u16.to_le_bytes());
        stream.extend_from_slice(&99u16.to_le_bytes());
        stream.extend_from_slice(&0u32.to_le_bytes());

        let result = EncryptionInfo::parse(&stream);
        assert!(result.is_err());
        let err = result.unwrap_err().to_string();
        assert!(err.contains("99.99"), "Error should mention version: {err}");
    }

    #[test]
    fn test_truncated_stream_error() {
        let stream = vec![0u8; 4]; // Too short
        let result = EncryptionInfo::parse(&stream);
        assert!(result.is_err());
        assert!(result.unwrap_err().to_string().contains("too short"));
    }

    // -----------------------------------------------------------------------
    // Hostile and unsupported EncryptionInfo
    // -----------------------------------------------------------------------

    /// A valid Agile descriptor: AES-256, SHA-512, one password key encryptor.
    const AGILE_XML: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<encryption xmlns="http://schemas.microsoft.com/office/2006/encryption" xmlns:p="http://schemas.microsoft.com/office/2006/keyEncryptor/password" xmlns:c="http://schemas.microsoft.com/office/2006/keyEncryptor/certificate">
  <keyData saltSize="16" blockSize="16" keyBits="256" hashSize="64" cipherAlgorithm="AES" cipherChaining="ChainingModeCBC" hashAlgorithm="SHA512" saltValue="AAAAAAAAAAAAAAAAAAAAAA=="/>
  <dataIntegrity encryptedHmacKey="AAAAAAAAAAAAAAAAAAAAAA==" encryptedHmacValue="AAAAAAAAAAAAAAAAAAAAAA=="/>
  <keyEncryptors>
    <keyEncryptor uri="http://schemas.microsoft.com/office/2006/keyEncryptor/password">
      <p:encryptedKey spinCount="100000" saltSize="16" blockSize="16" keyBits="256" hashSize="64" cipherAlgorithm="AES" cipherChaining="ChainingModeCBC" hashAlgorithm="SHA512" saltValue="BBBBBBBBBBBBBBBBBBBBBB==" encryptedKeyValue="CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCA=" encryptedVerifierHashInput="DDDDDDDDDDDDDDDDDDDDDD==" encryptedVerifierHashValue="EEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEA=="/>
    </keyEncryptor>
  </keyEncryptors>
</encryption>"#;

    /// Wraps an Agile XML descriptor in the 8-byte version 4.4 stream header.
    fn agile_stream(xml: &str) -> Vec<u8> {
        let mut stream = Vec::with_capacity(8 + xml.len());
        stream.extend_from_slice(&4u16.to_le_bytes());
        stream.extend_from_slice(&4u16.to_le_bytes());
        stream.extend_from_slice(&0x40u32.to_le_bytes());
        stream.extend_from_slice(xml.as_bytes());
        stream
    }

    /// Sets (`Some`) or removes (`None`) attribute `attr` on the first element of
    /// `xml` whose start tag begins with `elem`.
    fn set_attr(xml: &str, elem: &str, attr: &str, value: Option<&str>) -> String {
        let start = xml.find(elem).expect("element present");
        let key = format!(" {attr}=\"");
        let at = start + xml[start..].find(&key).expect("attribute present");
        let value_start = at + key.len();
        let value_end = value_start + xml[value_start..].find('"').expect("closing quote");
        match value {
            Some(v) => format!("{}{v}{}", &xml[..value_start], &xml[value_end..]),
            None => format!("{}{}", &xml[..at], &xml[value_end + 1..]),
        }
    }

    fn parse_with(elem: &str, attr: &str, value: Option<&str>) -> Result<EncryptionInfo> {
        EncryptionInfo::parse(&agile_stream(&set_attr(AGILE_XML, elem, attr, value)))
    }

    fn to_hex(bytes: &[u8]) -> String {
        use std::fmt::Write as _;
        bytes.iter().fold(String::new(), |mut s, b| {
            let _ = write!(s, "{b:02x}");
            s
        })
    }

    fn agile(info: EncryptionInfo) -> Box<AgileEncryptionInfo> {
        match info {
            EncryptionInfo::Agile(a) => a,
            other => unreachable!("expected Agile encryption, got {other:?}"),
        }
    }

    #[test]
    fn agile_template_parses() {
        let a = agile(EncryptionInfo::parse(&agile_stream(AGILE_XML)).unwrap());
        assert_eq!(a.pw_spin_count, 100_000);
        assert_eq!(a.pw_encrypted_key_value.len(), 32);
    }

    #[test]
    fn spin_count_above_cap_is_security_error() {
        for spin in ["4294967295", "10000001"] {
            let err = parse_with("<p:encryptedKey", "spinCount", Some(spin)).unwrap_err();
            assert_eq!(err.code(), "SECURITY", "spinCount {spin}: {err}");
        }
    }

    #[test]
    fn spin_count_at_cap_and_zero_parse() {
        for (spin, expected) in [("10000000", 10_000_000), ("0", 0)] {
            let a = agile(parse_with("<p:encryptedKey", "spinCount", Some(spin)).unwrap());
            assert_eq!(a.pw_spin_count, expected);
        }
    }

    #[test]
    fn spin_count_unparsable_or_missing_is_password_protected() {
        for value in [Some("abc"), None] {
            let err = parse_with("<p:encryptedKey", "spinCount", value).unwrap_err();
            assert_eq!(
                err.code(),
                "PASSWORD_PROTECTED",
                "spinCount {value:?}: {err}"
            );
            assert!(err.to_string().contains("spinCount"), "{err}");
        }
    }

    #[test]
    fn encrypted_key_bits_out_of_range_rejected_at_parse() {
        let err = parse_with("<p:encryptedKey", "keyBits", Some("4294967288")).unwrap_err();
        assert_eq!(err.code(), "PASSWORD_PROTECTED");
        assert!(
            err.to_string().contains("key size 4294967288 bits"),
            "{err}"
        );
    }

    #[test]
    fn encrypted_key_bits_192_parses() {
        let a = agile(parse_with("<p:encryptedKey", "keyBits", Some("192")).unwrap());
        assert_eq!(a.pw_key_bits, 192);
    }

    #[test]
    fn key_data_bits_out_of_range_rejected() {
        let err = parse_with("<keyData", "keyBits", Some("1024")).unwrap_err();
        assert!(err.to_string().contains("key size 1024 bits"), "{err}");
    }

    #[test]
    fn cfb_chaining_is_unsupported() {
        for elem in ["<keyData", "<p:encryptedKey"] {
            let err = parse_with(elem, "cipherChaining", Some("ChainingModeCFB")).unwrap_err();
            assert_eq!(err.code(), "PASSWORD_PROTECTED");
            assert!(err.to_string().contains("CFB"), "{elem}: {err}");
        }
    }

    #[test]
    fn non_aes_cipher_is_unsupported() {
        for elem in ["<keyData", "<p:encryptedKey"] {
            let err = parse_with(elem, "cipherAlgorithm", Some("RC2")).unwrap_err();
            assert!(err.to_string().contains("RC2"), "{elem}: {err}");
        }
    }

    #[test]
    fn block_salt_and_hash_sizes_are_validated() {
        let cases = [
            ("<keyData", "blockSize", "8"),
            ("<p:encryptedKey", "blockSize", "32"),
            ("<keyData", "saltSize", "32"),
            ("<p:encryptedKey", "saltSize", "8"),
            ("<keyData", "hashSize", "32"),
            ("<p:encryptedKey", "hashSize", "20"),
            ("<keyData", "hashAlgorithm", "MD5"),
            ("<p:encryptedKey", "hashAlgorithm", "SHA3-256"),
            ("<keyData", "blockSize", "sixteen"),
        ];
        for (elem, attr, value) in cases {
            let err = parse_with(elem, attr, Some(value)).unwrap_err();
            assert_eq!(
                err.code(),
                "PASSWORD_PROTECTED",
                "{elem} {attr}={value}: {err}"
            );
        }
    }

    #[test]
    fn certificate_encryptor_does_not_override_password_encryptor() {
        let certificate = r#"<keyEncryptor uri="http://schemas.microsoft.com/office/2006/keyEncryptor/certificate">
      <c:encryptedKey encryptedKeyValue="ZZZZZZZZZZZZZZZZZZZZZZ==" x509Certificate="AAAA" certVerifier="AAAA"/>
    </keyEncryptor>
    "#;
        let expected = agile(EncryptionInfo::parse(&agile_stream(AGILE_XML)).unwrap());
        let before =
            AGILE_XML.replacen("<keyEncryptor ", &format!("{certificate}<keyEncryptor "), 1);
        let after = AGILE_XML.replacen(
            "</keyEncryptors>",
            &format!("{certificate}</keyEncryptors>"),
            1,
        );
        for xml in [before, after] {
            let a = agile(EncryptionInfo::parse(&agile_stream(&xml)).unwrap());
            assert_eq!(a.pw_encrypted_key_value, expected.pw_encrypted_key_value);
            assert_eq!(a.pw_spin_count, 100_000);
        }
    }

    #[test]
    fn missing_password_encryptor_is_error() {
        let xml = AGILE_XML.replace(
            "http://schemas.microsoft.com/office/2006/keyEncryptor/password\">",
            "http://schemas.microsoft.com/office/2006/keyEncryptor/certificate\">",
        );
        let err = EncryptionInfo::parse(&agile_stream(&xml)).unwrap_err();
        assert_eq!(err.code(), "PASSWORD_PROTECTED");
    }

    /// A Standard (version 4.2) stream: `header_flags`, `alg_id`, `alg_id_hash`,
    /// `key_size`, and a spec-shaped EncryptionVerifier (saltSize 16,
    /// verifierHashSize 20).
    fn standard_stream(header_flags: u32, alg_id: u32, alg_id_hash: u32, key_size: u32) -> Vec<u8> {
        let mut stream = Vec::new();
        stream.extend_from_slice(&4u16.to_le_bytes());
        stream.extend_from_slice(&2u16.to_le_bytes());
        stream.extend_from_slice(&header_flags.to_le_bytes());
        stream.extend_from_slice(&40u32.to_le_bytes()); // headerSize
        stream.extend_from_slice(&header_flags.to_le_bytes());
        stream.extend_from_slice(&0u32.to_le_bytes()); // sizeExtra
        stream.extend_from_slice(&alg_id.to_le_bytes());
        stream.extend_from_slice(&alg_id_hash.to_le_bytes());
        stream.extend_from_slice(&key_size.to_le_bytes());
        stream.extend_from_slice(&0x18u32.to_le_bytes()); // providerType
        stream.extend_from_slice(&0u32.to_le_bytes()); // reserved1
        stream.extend_from_slice(&0u32.to_le_bytes()); // reserved2
        stream.extend_from_slice(&[b'A', 0, b'E', 0, b'S', 0, 0, 0]); // CSP name
        stream.extend_from_slice(&16u32.to_le_bytes()); // saltSize
        stream.extend_from_slice(&[0xAA; 16]); // salt
        stream.extend_from_slice(&[0xBB; 16]); // encryptedVerifier
        stream.extend_from_slice(&20u32.to_le_bytes()); // verifierHashSize
        stream.extend_from_slice(&[0xCC; 32]); // encryptedVerifierHash
        stream
    }

    #[test]
    fn standard_rc4_cryptoapi_is_unsupported() {
        let err = EncryptionInfo::parse(&standard_stream(0x04, 0x6801, 0x8004, 128)).unwrap_err();
        assert_eq!(err.code(), "PASSWORD_PROTECTED");
        assert!(err.to_string().contains("RC4"), "{err}");
    }

    #[test]
    fn standard_header_flags_alg_and_key_size_are_validated() {
        let cases = [
            (0x14, 0x660E, 0x8004, 128), // fExternal set
            (0x34, 0x660E, 0x8004, 128), // fExternal set
            (0x24, 0x6610, 0x8004, 128), // AES-256 with a 128-bit key
            (0x24, 0x660E, 0x8003, 128), // MD5
            (0x24, 0x1234, 0x8004, 128), // unknown algID
        ];
        for (flags, alg_id, alg_id_hash, key_size) in cases {
            let err = EncryptionInfo::parse(&standard_stream(flags, alg_id, alg_id_hash, key_size));
            assert!(
                err.is_err(),
                "flags {flags:#x} algID {alg_id:#x} keySize {key_size}"
            );
        }
    }

    #[test]
    fn standard_header_requires_crypto_api_and_aes_flags() {
        let cases = [
            (0x04, 0x660E, 0x8004, 128), // fAES clear
            (0x20, 0x660E, 0x8004, 128), // fCryptoAPI clear
        ];
        for (flags, alg_id, alg_id_hash, key_size) in cases {
            let err = EncryptionInfo::parse(&standard_stream(flags, alg_id, alg_id_hash, key_size))
                .unwrap_err();
            assert_eq!(err.code(), "PASSWORD_PROTECTED", "flags {flags:#x}");
            assert!(
                err.to_string().contains("must set fCryptoAPI and fAES"),
                "flags {flags:#x}: {err}"
            );
        }
    }

    #[test]
    fn standard_alg_id_zero_with_aes_means_aes_128() {
        match EncryptionInfo::parse(&standard_stream(0x24, 0, 0, 128)).unwrap() {
            EncryptionInfo::Standard(s) => {
                assert_eq!(s.alg_id, 0x660E);
                assert_eq!(s.hash_alg_id, 0x8004);
            }
            other => unreachable!("expected Standard encryption, got {other:?}"),
        }
    }

    #[test]
    fn standard_verifier_sizes_are_validated() {
        let mut bad_salt = standard_stream(0x24, 0x660E, 0x8004, 128);
        bad_salt[52..56].copy_from_slice(&20u32.to_le_bytes()); // saltSize
        assert!(EncryptionInfo::parse(&bad_salt).is_err());
        let mut bad_hash = standard_stream(0x24, 0x660E, 0x8004, 128);
        bad_hash[88..92].copy_from_slice(&32u32.to_le_bytes()); // verifierHashSize
        assert!(EncryptionInfo::parse(&bad_hash).is_err());
    }

    #[test]
    fn standard_huge_header_size_is_error_not_panic() {
        let mut stream = Vec::new();
        stream.extend_from_slice(&4u16.to_le_bytes());
        stream.extend_from_slice(&2u16.to_le_bytes());
        stream.extend_from_slice(&0x24u32.to_le_bytes());
        stream.extend_from_slice(&0xFFFF_FFFCu32.to_le_bytes());
        stream.extend_from_slice(&[0u8; 120]);
        let err = EncryptionInfo::parse(&stream).unwrap_err();
        assert_eq!(err.code(), "PASSWORD_PROTECTED");
    }

    #[test]
    fn word_standard_fixture_parses_spec_verifier() {
        let file = include_bytes!("../../tests/fixtures/msoffcrypto/ecma376standard_password.docx");
        let stream = crate::ole2::cfb::reader::read_stream(file, "EncryptionInfo").unwrap();
        match EncryptionInfo::parse(&stream).unwrap() {
            EncryptionInfo::Standard(s) => {
                assert_eq!(to_hex(&s.salt), "e88266490c5bd1eebd2b4394e3f830ef");
                assert_eq!(s.alg_id, 0x660E);
                assert_eq!(s.key_size, 128);
                assert_eq!(s.verifier_hash_size, 20);
                assert_eq!(s.encrypted_verifier.len(), 16);
                assert_eq!(s.encrypted_verifier_hash.len(), 32);
            }
            other => unreachable!("expected Standard encryption, got {other:?}"),
        }
    }

    #[test]
    fn excel_agile_fixture_parses() {
        let file = include_bytes!("../../tests/fixtures/msoffcrypto/example_password.xlsx");
        let stream = crate::ole2::cfb::reader::read_stream(file, "EncryptionInfo").unwrap();
        let a = agile(EncryptionInfo::parse(&stream).unwrap());
        assert_eq!(a.key_hash_alg, HashAlg::Sha512);
        assert_eq!(a.key_bits, 256);
        assert_eq!(a.pw_hash_alg, HashAlg::Sha512);
        assert_eq!(a.pw_key_bits, 256);
        assert_eq!(a.pw_spin_count, 100_000);
        assert_eq!(a.pw_salt.len(), 16);
    }

    #[test]
    fn hash_alg_spellings() {
        for (s, alg) in [
            ("SHA1", HashAlg::Sha1),
            ("SHA-1", HashAlg::Sha1),
            ("SHA256", HashAlg::Sha256),
            ("SHA-256", HashAlg::Sha256),
            ("SHA384", HashAlg::Sha384),
            ("SHA-384", HashAlg::Sha384),
            ("SHA512", HashAlg::Sha512),
            ("SHA-512", HashAlg::Sha512),
        ] {
            let parsed = HashAlg::try_from(s).unwrap();
            assert_eq!(parsed, alg, "{s}");
            assert_eq!(HashAlg::try_from(parsed.as_str()).unwrap(), alg);
            assert_eq!(parsed.to_string(), parsed.as_str());
        }
        let err = HashAlg::try_from("MD5").unwrap_err();
        assert_eq!(err.code(), "PASSWORD_PROTECTED");
        assert!(
            err.to_string()
                .contains("unsupported encryption: hash algorithm 'MD5'"),
            "{err}"
        );
    }

    #[test]
    fn spin_count_beyond_u32_is_security_error() {
        let err = parse_with("<p:encryptedKey", "spinCount", Some("99999999999")).unwrap_err();
        assert_eq!(err.code(), "SECURITY", "{err}");
    }

    #[test]
    fn build_encrypted_error_surfaces_security() {
        let stream = agile_stream(&set_attr(
            AGILE_XML,
            "<p:encryptedKey",
            "spinCount",
            Some("4294967295"),
        ));
        let ole2 = crate::ole2::cfb::writer::write_ole2(&[
            ("EncryptionInfo", &stream),
            ("EncryptedPackage", &[0u8; 24]),
        ])
        .unwrap();
        let cfb = Cfb::parse(&ole2).unwrap();
        assert_eq!(build_encrypted_error_with_cfb(&cfb).code(), "SECURITY");
    }

    #[test]
    fn build_encrypted_error_keeps_description_for_valid_info() {
        let ole2 = crate::ole2::cfb::writer::write_ole2(&[
            ("EncryptionInfo", &agile_stream(AGILE_XML)),
            ("EncryptedPackage", &[0u8; 24]),
        ])
        .unwrap();
        let err = build_encrypted_error_with_cfb(&Cfb::parse(&ole2).unwrap());
        assert_eq!(err.code(), "PASSWORD_PROTECTED");
        assert!(
            err.to_string()
                .contains("Agile encryption (AES-256, SHA512)"),
            "{err}"
        );
    }

    /// XML attribute-value normalization applies to every attribute the parser
    /// reads: a character reference or a tab inside base64 or a decimal, and a
    /// reference inside the encryptor `uri`.
    #[test]
    fn attribute_values_are_normalized() {
        let base = agile(EncryptionInfo::parse(&agile_stream(AGILE_XML)).unwrap());
        for salt in [
            "AAAA&#xA;AAAAAAAAAAAAAAAAAA==",
            "AAAA\tAAAAAAAAAAAAAAAAAA==",
        ] {
            let a = agile(parse_with("<keyData", "saltValue", Some(salt)).unwrap());
            assert_eq!(a.key_salt, base.key_salt, "{salt:?}");
        }
        let a = agile(
            parse_with(
                "<p:encryptedKey",
                "saltValue",
                Some("BBBB&#10;BBBBBBBBBBBBBBBBBB=="),
            )
            .unwrap(),
        );
        assert_eq!(a.pw_salt, base.pw_salt);
        let xml = set_attr(
            AGILE_XML,
            "<dataIntegrity",
            "encryptedHmacKey",
            Some("AAAA&#xA;AAAAAAAAAAAAAAAAAA=="),
        );
        let a = agile(EncryptionInfo::parse(&agile_stream(&xml)).unwrap());
        assert_eq!(a.encrypted_hmac_key, base.encrypted_hmac_key);

        let a = agile(parse_with("<p:encryptedKey", "spinCount", Some("&#49;00000")).unwrap());
        assert_eq!(a.pw_spin_count, 100_000);

        let xml = AGILE_XML.replace("keyEncryptor/password\">", "keyEncryptor&#x2F;password\">");
        let a = agile(EncryptionInfo::parse(&agile_stream(&xml)).unwrap());
        assert_eq!(a.pw_spin_count, 100_000);
    }

    #[test]
    fn undeclared_entity_in_an_attribute_is_malformed() {
        let err = parse_with("<keyData", "saltValue", Some("AA&foo;")).unwrap_err();
        assert_eq!(err.code(), "PASSWORD_PROTECTED");
        assert!(
            err.to_string().contains("malformed EncryptionInfo"),
            "{err}"
        );
    }

    #[test]
    fn duplicate_attribute_is_malformed_encryption_info() {
        for (elem, dup) in [
            ("<keyData ", "saltSize=\"16\" "),
            (
                "<dataIntegrity ",
                "encryptedHmacKey=\"AAAAAAAAAAAAAAAAAAAAAA==\" ",
            ),
            ("<p:encryptedKey ", "spinCount=\"100000\" "),
            ("<keyEncryptor ", "x=\"1\" x=\"2\" "),
        ] {
            let xml = AGILE_XML.replacen(elem, &format!("{elem}{dup}"), 1);
            let err = EncryptionInfo::parse(&agile_stream(&xml)).unwrap_err();
            assert_eq!(err.code(), "PASSWORD_PROTECTED", "{elem}: {err}");
            assert!(
                err.to_string().contains("malformed EncryptionInfo"),
                "{elem}: {err}"
            );
        }
    }
}
