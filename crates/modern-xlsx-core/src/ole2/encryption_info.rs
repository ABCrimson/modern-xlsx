use core::hint::cold_path;

use crate::errors::ModernXlsxError;
use quick_xml::events::Event;
use quick_xml::Reader;

type Result<T> = std::result::Result<T, ModernXlsxError>;

/// Combined encryption info parsed from the EncryptionInfo stream.
#[derive(Debug, Clone)]
pub enum EncryptionInfo {
    Standard(StandardEncryptionInfo),
    Agile(Box<AgileEncryptionInfo>),
}

/// Standard Encryption info (version 2.2, 3.2, 4.2).
#[derive(Debug, Clone)]
pub struct StandardEncryptionInfo {
    pub alg_id: u32,
    pub hash_alg_id: u32,
    pub key_size: u32,
    pub provider: String,
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
    pub key_cipher: String,
    pub key_chaining: String,
    pub key_hash_alg: String,
    // Data integrity
    pub encrypted_hmac_key: Vec<u8>,
    pub encrypted_hmac_value: Vec<u8>,
    // Password key encryptor
    pub pw_spin_count: u32,
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

impl EncryptionInfo {
    /// Parse the raw EncryptionInfo stream bytes.
    pub fn parse(stream: &[u8]) -> Result<Self> {
        if stream.len() < 8 {
            cold_path();
            return Err(ModernXlsxError::PasswordProtected(
                "EncryptionInfo stream too short".into(),
            ));
        }
        let version_major = u16::from_le_bytes([stream[0], stream[1]]);
        let version_minor = u16::from_le_bytes([stream[2], stream[3]]);

        match (version_major, version_minor) {
            (4, 4) => Self::parse_agile(&stream[8..]),
            (2, 2) | (3, 2) | (4, 2) => Self::parse_standard(&stream[8..]),
            _ => Err(ModernXlsxError::PasswordProtected(format!(
                "Unsupported encryption version {version_major}.{version_minor}"
            ))),
        }
    }

    fn parse_agile(xml_bytes: &[u8]) -> Result<Self> {
        let mut key_salt = Vec::new();
        let mut key_block_size = 0u32;
        let mut key_bits = 0u32;
        let mut key_hash_size = 0u32;
        let mut key_cipher = String::new();
        let mut key_chaining = String::new();
        let mut key_hash_alg = String::new();
        let mut encrypted_hmac_key = Vec::new();
        let mut encrypted_hmac_value = Vec::new();
        let mut pw_spin_count = 0u32;
        let mut pw_salt = Vec::new();
        let mut pw_block_size = 0u32;
        let mut pw_key_bits = 0u32;
        let mut pw_hash_size = 0u32;
        let mut pw_cipher = String::new();
        let mut pw_chaining = String::new();
        let mut pw_hash_alg = String::new();
        let mut pw_encrypted_key_value = Vec::new();
        let mut pw_encrypted_verifier_hash_input = Vec::new();
        let mut pw_encrypted_verifier_hash_value = Vec::new();

        let mut reader = Reader::from_reader(xml_bytes);
        reader.config_mut().trim_text(true);
        let mut buf = Vec::with_capacity(512);

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Empty(ref e)) | Ok(Event::Start(ref e)) => {
                    let local = e.local_name();
                    let tag = std::str::from_utf8(local.as_ref()).unwrap_or_default();
                    match tag {
                        "keyData" => {
                            for attr in e.attributes().flatten() {
                                let key =
                                    std::str::from_utf8(attr.key.as_ref()).unwrap_or_default();
                                let val =
                                    std::str::from_utf8(&attr.value).unwrap_or_default();
                                match key {
                                    "saltValue" => key_salt = decode_base64(val)?,
                                    "blockSize" => key_block_size = val.parse().unwrap_or(0),
                                    "keyBits" => key_bits = val.parse().unwrap_or(0),
                                    "hashSize" => key_hash_size = val.parse().unwrap_or(0),
                                    "cipherAlgorithm" => key_cipher = val.to_owned(),
                                    "cipherChaining" => key_chaining = val.to_owned(),
                                    "hashAlgorithm" => key_hash_alg = val.to_owned(),
                                    _ => {}
                                }
                            }
                        }
                        "dataIntegrity" => {
                            for attr in e.attributes().flatten() {
                                let key =
                                    std::str::from_utf8(attr.key.as_ref()).unwrap_or_default();
                                let val =
                                    std::str::from_utf8(&attr.value).unwrap_or_default();
                                match key {
                                    "encryptedHmacKey" => {
                                        encrypted_hmac_key = decode_base64(val)?;
                                    }
                                    "encryptedHmacValue" => {
                                        encrypted_hmac_value = decode_base64(val)?;
                                    }
                                    _ => {}
                                }
                            }
                        }
                        "encryptedKey" => {
                            for attr in e.attributes().flatten() {
                                let key =
                                    std::str::from_utf8(attr.key.as_ref()).unwrap_or_default();
                                let val =
                                    std::str::from_utf8(&attr.value).unwrap_or_default();
                                match key {
                                    "spinCount" => {
                                        pw_spin_count = val.parse().unwrap_or(0);
                                    }
                                    "saltValue" => pw_salt = decode_base64(val)?,
                                    "blockSize" => {
                                        pw_block_size = val.parse().unwrap_or(0);
                                    }
                                    "keyBits" => pw_key_bits = val.parse().unwrap_or(0),
                                    "hashSize" => {
                                        pw_hash_size = val.parse().unwrap_or(0);
                                    }
                                    "cipherAlgorithm" => pw_cipher = val.to_owned(),
                                    "cipherChaining" => pw_chaining = val.to_owned(),
                                    "hashAlgorithm" => pw_hash_alg = val.to_owned(),
                                    "encryptedKeyValue" => {
                                        pw_encrypted_key_value = decode_base64(val)?;
                                    }
                                    "encryptedVerifierHashInput" => {
                                        pw_encrypted_verifier_hash_input =
                                            decode_base64(val)?;
                                    }
                                    "encryptedVerifierHashValue" => {
                                        pw_encrypted_verifier_hash_value =
                                            decode_base64(val)?;
                                    }
                                    _ => {}
                                }
                            }
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

        Ok(EncryptionInfo::Agile(Box::new(AgileEncryptionInfo {
            key_salt,
            key_block_size,
            key_bits,
            key_hash_size,
            key_cipher,
            key_chaining,
            key_hash_alg,
            encrypted_hmac_key,
            encrypted_hmac_value,
            pw_spin_count,
            pw_salt,
            pw_block_size,
            pw_key_bits,
            pw_hash_size,
            pw_cipher,
            pw_chaining,
            pw_hash_alg,
            pw_encrypted_key_value,
            pw_encrypted_verifier_hash_input,
            pw_encrypted_verifier_hash_value,
        })))
    }

    fn parse_standard(data: &[u8]) -> Result<Self> {
        // After parse() skips 8 bytes (version 4 + flags 4), we receive:
        //   data[0..4]  = headerSize (u32)
        //   data[4..4+headerSize] = EncryptionHeader
        //     header[0..4]  = flags
        //     header[4..8]  = sizeExtra (must be 0)
        //     header[8..12] = algID
        //     header[12..16] = algIDHash
        //     header[16..20] = keySize (bits)
        //     header[20..24] = providerType
        //     header[24..28] = reserved1
        //     header[28..32] = reserved2
        //     header[32..]  = CSP name (UTF-16LE, null-terminated)
        //   data[4+headerSize..] = EncryptionVerifier
        //     salt (16), encrypted verifier (16), hash size (4), encrypted hash (32)
        if data.len() < 4 {
            cold_path();
            return Err(ModernXlsxError::PasswordProtected(
                "Standard encryption header too short".into(),
            ));
        }

        let header_size = u32::from_le_bytes(data[0..4].try_into().unwrap_or_default()) as usize;

        if data.len() < 4 + header_size + 68 {
            cold_path();
            return Err(ModernXlsxError::PasswordProtected(
                "Standard encryption data too short for verifier".into(),
            ));
        }

        let header = &data[4..4 + header_size];
        if header.len() < 32 {
            cold_path();
            return Err(ModernXlsxError::PasswordProtected(
                "Standard encryption header fields too short".into(),
            ));
        }
        let alg_id = u32::from_le_bytes(header[8..12].try_into().unwrap_or_default());
        let hash_alg_id = u32::from_le_bytes(header[12..16].try_into().unwrap_or_default());
        let key_size = u32::from_le_bytes(header[16..20].try_into().unwrap_or_default());

        // CSP name: UTF-16LE from offset 32 to end of header
        let csp_bytes = &header[32..];
        let provider: String = csp_bytes
            .chunks_exact(2)
            .map(|c| u16::from_le_bytes([c[0], c[1]]))
            .take_while(|&c| c != 0)
            .map(|c| char::from_u32(u32::from(c)).unwrap_or('\u{FFFD}'))
            .collect();

        // Verifier starts after header
        let v = &data[4 + header_size..];
        let salt = v[0..16].to_vec();
        let encrypted_verifier = v[16..32].to_vec();
        let verifier_hash_size = u32::from_le_bytes(v[32..36].try_into().unwrap_or_default());
        let encrypted_verifier_hash = v[36..68].to_vec();

        Ok(EncryptionInfo::Standard(StandardEncryptionInfo {
            alg_id,
            hash_alg_id,
            key_size,
            provider,
            salt,
            encrypted_verifier,
            verifier_hash_size,
            encrypted_verifier_hash,
        }))
    }
}

/// Describe the encryption method for error messages.
pub fn describe_encryption(info: &EncryptionInfo) -> String {
    match info {
        EncryptionInfo::Agile(a) => format!(
            "Agile encryption ({}-{}, {})",
            a.key_cipher, a.key_bits, a.key_hash_alg
        ),
        EncryptionInfo::Standard(s) => {
            format!("Standard encryption (key size: {} bits)", s.key_size)
        }
    }
}

/// Convenience: read EncryptionInfo stream from OLE2 data and parse it.
pub fn read_and_parse_encryption_info(ole2_data: &[u8]) -> Result<EncryptionInfo> {
    let stream = crate::ole2::detect::read_stream(ole2_data, "EncryptionInfo")?;
    EncryptionInfo::parse(&stream)
}

/// Builds the error message for an encrypted XLSX file.
///
/// Tries to parse the `EncryptionInfo` stream for a detailed description;
/// falls back to the generic constant when parsing fails.
pub fn build_encrypted_error(ole2_data: &[u8]) -> crate::ModernXlsxError {
    let msg = match read_and_parse_encryption_info(ole2_data) {
        Ok(info) => {
            let desc = describe_encryption(&info);
            format!(
                "Password-protected XLSX ({desc}). \
                 Provide password via readBuffer(data, {{ password: '...' }})."
            )
        }
        Err(_) => crate::ole2::detect::ERR_ENCRYPTED.into(),
    };
    crate::ModernXlsxError::PasswordProtected(msg)
}

/// Base64 decode helper (inline implementation — no external dep).
/// Task 3 will switch to the `base64` crate.
fn decode_base64(input: &str) -> Result<Vec<u8>> {
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
                assert_eq!(a.key_hash_alg, "SHA512");
                assert_eq!(a.key_cipher, "AES");
                assert_eq!(a.key_chaining, "ChainingModeCBC");
                assert_eq!(a.key_block_size, 16);
                assert_eq!(a.key_hash_size, 64);
                assert_eq!(a.pw_spin_count, 100000);
                assert_eq!(a.pw_key_bits, 256);
                assert_eq!(a.pw_hash_alg, "SHA512");
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
                assert_eq!(a.key_hash_alg, "SHA256");
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
        stream.extend_from_slice(&0x6801u32.to_le_bytes()); // algID = AES-128
        stream.extend_from_slice(&0x8004u32.to_le_bytes()); // algIDHash = SHA-1
        stream.extend_from_slice(&128u32.to_le_bytes()); // keySize
        stream.extend_from_slice(&0x18u32.to_le_bytes()); // providerType
        stream.extend_from_slice(&0u32.to_le_bytes()); // reserved1
        stream.extend_from_slice(&0u32.to_le_bytes()); // reserved2
        // CSP name (UTF-16LE, null-terminated) — "AES" + null = 8 bytes
        stream.extend_from_slice(&[b'A', 0, b'E', 0, b'S', 0, 0, 0]);
        // Verifier:
        stream.extend_from_slice(&[0xAA; 16]); // salt
        stream.extend_from_slice(&[0xBB; 16]); // encrypted verifier
        stream.extend_from_slice(&20u32.to_le_bytes()); // hash size
        stream.extend_from_slice(&[0xCC; 32]); // encrypted verifier hash

        let info = EncryptionInfo::parse(&stream).unwrap();
        match info {
            EncryptionInfo::Standard(s) => {
                assert_eq!(s.alg_id, 0x6801);
                assert_eq!(s.hash_alg_id, 0x8004);
                assert_eq!(s.key_size, 128);
                assert_eq!(s.provider, "AES");
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
}
