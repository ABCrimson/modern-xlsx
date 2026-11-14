//! Standard Encryption (MS-OFFCRYPTO 2.3.4.5-2.3.4.9, read only): AES-ECB
//! with no IV and no padding, the password verifier check and the
//! EncryptedPackage decryption.
//!
//! The key comes from [`standard_key`] in the parent module, which also holds
//! the hash helper and the package size check shared with Agile Encryption.

use core::hint::cold_path;

use aes::{Aes128Dec, Aes192Dec, Aes256Dec};
#[cfg(test)]
use aes::{Aes128Enc, Aes192Enc, Aes256Enc};
#[cfg(test)]
use cbc::cipher::BlockCipherEncrypt;
use cbc::cipher::{BlockCipherDecrypt, KeyInit, consts::U16};
use zeroize::Zeroizing;

use super::{
    AES_BLOCK, declared_package_size, hash_bytes, require_block_aligned, standard_key,
    unsupported_aes_key,
};
use crate::Result;
use crate::errors::ModernXlsxError;
use crate::ole2::encryption_info::{HashAlg, StandardEncryptionInfo};

// ---------------------------------------------------------------------------
// AES-ECB primitives (Standard Encryption)
// ---------------------------------------------------------------------------

/// Decrypts data using AES-ECB with no padding removal.
///
/// Standard Encryption uses ECB mode, with no IV, for the password verifier and
/// for the EncryptedPackage (MS-OFFCRYPTO 2.3.4.7-2.3.4.9). Implemented block by
/// block with the raw [`BlockCipherDecrypt`] of the `aes` types. Selects AES-128,
/// AES-192 or AES-256 from the key length (16, 24 or 32 bytes); unaligned `data`
/// is an error.
pub(super) fn aes_ecb_decrypt_no_pad(key: &[u8], data: &[u8]) -> Result<Vec<u8>> {
    require_block_aligned(data, "AES-ECB")?;
    let mut buf = data.to_vec();
    let (blocks, _) = buf.as_chunks_mut::<AES_BLOCK>();
    match key.len() {
        16 => ecb_decrypt::<Aes128Dec>(key, blocks)?,
        24 => ecb_decrypt::<Aes192Dec>(key, blocks)?,
        32 => ecb_decrypt::<Aes256Dec>(key, blocks)?,
        n => return Err(unsupported_aes_key(n)),
    }
    Ok(buf)
}

/// ECB-decrypts `blocks` in place under the block cipher `C`.
fn ecb_decrypt<C>(key: &[u8], blocks: &mut [[u8; AES_BLOCK]]) -> Result<()>
where
    C: BlockCipherDecrypt<BlockSize = U16> + KeyInit,
{
    let cipher = C::new_from_slice(key)
        .map_err(|e| ModernXlsxError::PasswordProtected(format!("AES-ECB init: {e}")))?;
    for block in blocks {
        cipher.decrypt_block(block.into());
    }
    Ok(())
}

/// Encrypts data using AES-ECB with no padding (for test use).
///
/// Block by block with the raw [`BlockCipherEncrypt`]; same key sizes and
/// alignment rule as [`aes_ecb_decrypt_no_pad`].
#[cfg(test)]
pub(super) fn aes_ecb_encrypt_no_pad(key: &[u8], data: &[u8]) -> Result<Vec<u8>> {
    fn ecb_encrypt<C>(key: &[u8], blocks: &mut [[u8; AES_BLOCK]]) -> Result<()>
    where
        C: BlockCipherEncrypt<BlockSize = U16> + KeyInit,
    {
        let cipher = C::new_from_slice(key)
            .map_err(|e| ModernXlsxError::PasswordProtected(format!("AES-ECB init: {e}")))?;
        for block in blocks {
            cipher.encrypt_block(block.into());
        }
        Ok(())
    }

    require_block_aligned(data, "AES-ECB")?;
    let mut buf = data.to_vec();
    let (blocks, _) = buf.as_chunks_mut::<AES_BLOCK>();
    match key.len() {
        16 => ecb_encrypt::<Aes128Enc>(key, blocks)?,
        24 => ecb_encrypt::<Aes192Enc>(key, blocks)?,
        32 => ecb_encrypt::<Aes256Enc>(key, blocks)?,
        n => return Err(unsupported_aes_key(n)),
    }
    Ok(buf)
}

// ---------------------------------------------------------------------------
// Standard Encryption — password verification & decryption
// ---------------------------------------------------------------------------

/// Verifies a password against Standard encryption info (MS-OFFCRYPTO 2.3.4.9).
///
/// Returns the derived decryption key, zeroized on drop, if the password is
/// correct.
///
/// Key derivation: `standard_key` over the verifier salt and `KeySize`
/// (MS-OFFCRYPTO 2.3.4.7: SHA-1 chain of 50,000 rounds, H_final with block 0,
/// then X1/X2 with 0x36/0x5C).
///
/// Verification:
/// - Decrypt `encrypted_verifier` (16 bytes) with AES-ECB
/// - Compute SHA-1(decrypted_verifier)
/// - Decrypt `encrypted_verifier_hash` (32 bytes) with AES-ECB
/// - Compare first `verifier_hash_size` bytes in constant time
pub fn verify_password_standard(
    password: &str,
    info: &StandardEncryptionInfo,
) -> Result<Zeroizing<Vec<u8>>> {
    // 1. Derive the key. The Zeroizing key is wiped on every early return and
    //    handed to the caller on success.
    let derived = standard_key(password, &info.salt, info.key_size)?;

    // 2. Decrypt the encrypted verifier (16 bytes) with AES-ECB
    let decrypted_verifier = aes_ecb_decrypt_no_pad(&derived, &info.encrypted_verifier)?;

    // 3. Compute SHA-1 of the decrypted verifier
    let verifier_hash = hash_bytes(HashAlg::Sha1, &decrypted_verifier);

    // 4. Decrypt the encrypted verifier hash (32 bytes) with AES-ECB
    let decrypted_verifier_hash = aes_ecb_decrypt_no_pad(&derived, &info.encrypted_verifier_hash)?;

    // 5. Compare first verifier_hash_size bytes in constant time; a zero size
    //    would compare two empty slices, which constant_time_eq accepts.
    let hash_len = info.verifier_hash_size as usize;
    if hash_len == 0 || verifier_hash.len() < hash_len || decrypted_verifier_hash.len() < hash_len {
        cold_path();
        return Err(ModernXlsxError::PasswordProtected(
            "Incorrect password.".into(),
        ));
    }
    if !constant_time_eq::constant_time_eq(
        &verifier_hash[..hash_len],
        &decrypted_verifier_hash[..hash_len],
    ) {
        cold_path();
        return Err(ModernXlsxError::PasswordProtected(
            "Incorrect password.".into(),
        ));
    }

    // 6. Return the derived key (still zeroized on drop by the caller)
    Ok(derived)
}

/// Decrypts the EncryptedPackage stream for Standard Encryption.
///
/// Layout: `[8 bytes: original size LE64] [encrypted data]`. The data is one
/// AES-ECB stream (layout MS-OFFCRYPTO 2.3.4.4, ECB per 2.3.4.7): no IV, no
/// segments, no padding removal; the plaintext is truncated to the declared
/// size. `data_key` must be the `KeySize` that `info` declares
/// ([`verify_password_standard`] returns it).
///
/// # Errors
/// `PasswordProtected` for a stream shorter than its size prefix, one with no
/// encrypted data (checked in that order), a key whose length is not `KeySize`
/// or unaligned ciphertext, and `Security` when the prefix declares more bytes
/// than the stream holds.
pub fn decrypt_standard_package(
    data_key: &[u8],
    info: &StandardEncryptionInfo,
    encrypted_package: &[u8],
) -> Result<Vec<u8>> {
    let Some((size, payload)) = encrypted_package.split_first_chunk::<8>() else {
        cold_path();
        return Err(ModernXlsxError::PasswordProtected(
            "EncryptedPackage too short".into(),
        ));
    };

    if payload.is_empty() {
        cold_path();
        return Err(ModernXlsxError::PasswordProtected(
            "EncryptedPackage has no encrypted data".into(),
        ));
    }
    let original_size = declared_package_size(*size, payload.len())?;
    if data_key.len().checked_mul(8) != usize::try_from(info.key_size).ok() {
        cold_path();
        return Err(ModernXlsxError::PasswordProtected(format!(
            "Standard key is {} bytes but KeySize is {} bits",
            data_key.len(),
            info.key_size
        )));
    }

    // Standard Encryption decrypts the package with AES-ECB: no IV.
    let mut result = aes_ecb_decrypt_no_pad(data_key, payload)?;
    result.truncate(original_size);

    Ok(result)
}
