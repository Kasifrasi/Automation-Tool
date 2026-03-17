//! ECMA-376 workbook protection: SHA-512 hash + ZIP/XML injection.
//!
//! Implements password-based workbook structure protection per
//! ECMA-376 (4th Edition, Part 4, §3.2.29) with SHA-512 and 100,000 iterations.
//!
//! The protection prevents adding, deleting, renaming, and reordering
//! worksheets in Excel/LibreOffice.
//!
//! # Workflow
//!
//! 1. [`precompute_hash()`] — compute the SHA-512 hash upfront (~25ms)
//! 2. [`protect_workbook_precomputed()`] — inject the hash into a .xlsx via ZIP/XML rewriting

use base64::{engine::general_purpose, Engine as _};
use byteorder::{WriteBytesExt, LE};
use quick_xml::events::Event;
use quick_xml::reader::Reader;
use quick_xml::writer::Writer;
use sha2::{Digest, Sha512};
use std::fs::File;
use std::io::{Cursor, Read, Write};
use zip::write::FileOptions;
use zip::{ZipArchive, ZipWriter};

const DEFAULT_SPIN_COUNT: u32 = 100_000;
const SALT_SIZE: usize = 16;

// ============================================================================
// ProtectionError
// ============================================================================

/// Error during workbook protection operations.
///
/// Covers ZIP manipulation, XML parsing, and file I/O errors
/// that occur when injecting workbook protection into a .xlsx file.
#[derive(Debug)]
pub enum ProtectionError {
    /// I/O error (file open, read, write)
    Io(std::io::Error),
    /// ZIP archive error (read, write, copy entries)
    Zip(zip::result::ZipError),
    /// XML parsing or writing error
    Xml(quick_xml::Error),
    /// XML element name contains invalid UTF-8
    InvalidUtf8(std::str::Utf8Error),
}

impl std::fmt::Display for ProtectionError {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Io(e) => write!(f, "I/O error during workbook protection: {e}"),
            Self::Zip(e) => write!(f, "ZIP archive error during workbook protection: {e}"),
            Self::Xml(e) => write!(f, "XML error during workbook protection: {e}"),
            Self::InvalidUtf8(e) => write!(f, "Invalid UTF-8 in XML element name: {e}"),
        }
    }
}

impl std::error::Error for ProtectionError {
    fn source(&self) -> Option<&(dyn std::error::Error + 'static)> {
        match self {
            Self::Io(e) => Some(e),
            Self::Zip(e) => Some(e),
            Self::Xml(e) => Some(e),
            Self::InvalidUtf8(e) => Some(e),
        }
    }
}

impl From<std::io::Error> for ProtectionError {
    fn from(e: std::io::Error) -> Self {
        Self::Io(e)
    }
}

impl From<zip::result::ZipError> for ProtectionError {
    fn from(e: zip::result::ZipError) -> Self {
        Self::Zip(e)
    }
}

impl From<quick_xml::Error> for ProtectionError {
    fn from(e: quick_xml::Error) -> Self {
        Self::Xml(e)
    }
}

impl From<std::str::Utf8Error> for ProtectionError {
    fn from(e: std::str::Utf8Error) -> Self {
        Self::InvalidUtf8(e)
    }
}

/// Fixed salt for deterministic hash computation.
/// Same salt + same password = same hash (no random element).
const FIXED_SALT: [u8; 16] = [
    0x6b, 0x6d, 0x77, 0x66, 0x62, 0x5f, 0x72, 0x70, // "kmwfb_rp"
    0x74, 0x5f, 0x76, 0x31, 0x5f, 0x21, 0x21, 0x00, // "t_v1_!!\0"
];

/// Precomputed workbook protection hash.
///
/// Create once with [`precompute_hash`] or [`precompute_hash_with_spin_count`],
/// then reuse for any number of files via [`protect_workbook_precomputed`].
/// Saves the expensive SHA-512 iteration cost in batch operations (~25ms per 100,000 iterations).
pub struct PrecomputedHash {
    pub salt_b64: String,
    pub hash_b64: String,
    pub spin_count: u32,
}

/// Fast spin count (sufficient to deter casual users)
const FAST_SPIN_COUNT: u32 = 1_000;

/// Workbook-level protection settings.
///
/// Protects the workbook structure (prevents adding, deleting, moving, or renaming sheets).
/// Uses SHA-512 based password hashing according to ECMA-376 standard.
///
/// # Security
///
/// Workbook protection is a **UI-level deterrent**, not cryptographic security.
/// The password prevents casual users from modifying the sheet structure in Excel,
/// but the underlying XML in the .xlsx file can be edited by determined users.
/// Do not rely on this for sensitive data protection.
///
/// # Example
/// ```ignore
/// use fb_generator::WorkbookProtection;
///
/// // Standard (100,000 iterations, ~25ms per report)
/// let protection = WorkbookProtection::new("secret123");
///
/// // Fast (1,000 iterations, ~0.3ms per report)
/// let fast = WorkbookProtection::fast("secret123");
///
/// // Custom
/// let custom = WorkbookProtection::new("secret123").with_spin_count(10_000);
/// ```
#[derive(Debug, Clone)]
pub struct WorkbookProtection {
    /// Password for workbook protection
    pub password: String,
    /// Lock structure (prevent sheet add/delete/move/rename)
    pub lock_structure: bool,
    /// SHA-512 iteration count (higher = slower but more secure)
    ///
    /// - `100_000` (default): ECMA-376 compliant, ~25ms per report
    /// - `1_000` (fast): deters casual users, ~0.3ms per report
    pub spin_count: u32,
}

impl WorkbookProtection {
    /// Creates new workbook protection with the given password
    ///
    /// Uses the standard ECMA-376 spin count of 100,000 iterations.
    pub fn new(password: impl Into<String>) -> Self {
        Self {
            password: password.into(),
            lock_structure: true,
            spin_count: DEFAULT_SPIN_COUNT,
        }
    }

    /// Creates fast workbook protection (1,000 iterations instead of 100,000)
    ///
    /// ~80x faster than `new()`. Sufficient to deter casual users,
    /// but not secure against targeted brute-force attacks.
    pub fn fast(password: impl Into<String>) -> Self {
        Self {
            password: password.into(),
            lock_structure: true,
            spin_count: FAST_SPIN_COUNT,
        }
    }

    /// Sets a custom spin count for SHA-512 hashing
    pub fn with_spin_count(mut self, count: u32) -> Self {
        self.spin_count = count;
        self
    }

    /// Sets whether to lock workbook structure
    pub fn lock_structure(mut self, lock: bool) -> Self {
        self.lock_structure = lock;
        self
    }
}

/// Computes the workbook protection hash once with a fixed salt (100,000 iterations).
///
/// The hash is deterministic: same password always produces the same hash.
/// Call once for batch protection of many files with the same password,
/// then pass the result to [`protect_workbook_precomputed`].
pub fn precompute_hash(password: &str) -> PrecomputedHash {
    precompute_hash_with_spin_count(password, DEFAULT_SPIN_COUNT)
}

/// Computes the workbook protection hash once with a fixed salt and configurable spin count.
pub fn precompute_hash_with_spin_count(password: &str, spin_count: u32) -> PrecomputedHash {
    let hash_b64 = compute_hash(password, &FIXED_SALT, spin_count);
    PrecomputedHash {
        salt_b64: general_purpose::STANDARD.encode(FIXED_SALT),
        hash_b64,
        spin_count,
    }
}

/// Protects a workbook structure with a precomputed hash (no hashing overhead).
///
/// Significantly faster than [`protect_workbook`] for batch operations, since the
/// SHA-512 hash is not recomputed but injected directly into the XML.
pub fn protect_workbook_precomputed(
    input_path: &str,
    output_path: &str,
    hash: &PrecomputedHash,
) -> Result<(), ProtectionError> {
    write_protected_zip(input_path, output_path, &hash.salt_b64, &hash.hash_b64, hash.spin_count)
}

/// Protects the workbook structure of an existing Excel file with a password.
///
/// Prevents users from adding, deleting, moving, or renaming sheets.
/// Follows the ECMA-376 standard (SHA-512 hashing).
///
/// Uses the default spin count of 100,000 iterations.
pub fn protect_workbook(input_path: &str, output_path: &str, password: &str) -> Result<(), ProtectionError> {
    protect_workbook_with_spin_count(input_path, output_path, password, DEFAULT_SPIN_COUNT)
}

/// Protects the workbook structure with a configurable spin count.
pub fn protect_workbook_with_spin_count(
    input_path: &str,
    output_path: &str,
    password: &str,
    spin_count: u32,
) -> Result<(), ProtectionError> {
    let (salt_b64, hash_b64) = hash_password(password, spin_count);
    write_protected_zip(input_path, output_path, &salt_b64, &hash_b64, spin_count)
}

/// Core function: read ZIP, inject protection tag into workbook.xml, write back.
///
/// All entries except `xl/workbook.xml` are copied 1:1 via `raw_copy_file`
/// (no decompression/recompression). Only workbook.xml is read, modified,
/// and recompressed.
fn write_protected_zip(
    input_path: &str,
    output_path: &str,
    salt_b64: &str,
    hash_b64: &str,
    spin_count: u32,
) -> Result<(), ProtectionError> {
    let file = File::open(input_path)?;
    let mut archive = ZipArchive::new(file)?;

    let out_file = File::create(output_path)?;
    let mut zip_writer = ZipWriter::new(out_file);

    for i in 0..archive.len() {
        let name = archive.by_index(i)?.name().to_string();

        if name == "xl/workbook.xml" {
            let mut file = archive.by_index(i)?;
            let compression = file.compression();
            let unix_mode = file.unix_mode();
            let mut content = Vec::new();
            file.read_to_end(&mut content)?;

            let new_xml = inject_protection(&content, salt_b64, hash_b64, spin_count)?;
            let options = FileOptions::<()>::default()
                .compression_method(compression)
                .unix_permissions(unix_mode.unwrap_or(0o644));
            zip_writer.start_file(&name, options)?;
            zip_writer.write_all(&new_xml)?;
        } else {
            let file = archive.by_index(i)?;
            zip_writer.raw_copy_file(file)?;
        }
    }

    zip_writer.finish()?;
    Ok(())
}

/// Computes the SHA-512 hash for a given password and salt.
///
/// Algorithm:
/// 1. H0 = SHA512(salt + password(UTF-16LE))
/// 2. Loop `spin_count` times: H(n) = SHA512(H(n-1) + iterator(4 bytes LE))
fn compute_hash(password: &str, salt: &[u8], spin_count: u32) -> String {
    let pw_utf16: Vec<u8> = password
        .encode_utf16()
        .flat_map(|c| c.to_le_bytes())
        .collect();

    let mut hasher = Sha512::new();
    hasher.update(salt);
    hasher.update(&pw_utf16);
    let mut hash = hasher.finalize();

    for i in 0..spin_count {
        let mut iterator = [0u8; 4];
        let mut wtr = &mut iterator[..];
        wtr.write_u32::<LE>(i).unwrap();

        let mut next_hasher = Sha512::new();
        next_hasher.update(hash);
        next_hasher.update(iterator);
        hash = next_hasher.finalize();
    }

    general_purpose::STANDARD.encode(hash)
}

/// Generates the password hash with a random salt (for single-file protection).
fn hash_password(password: &str, spin_count: u32) -> (String, String) {
    let mut salt = [0u8; SALT_SIZE];
    getrandom::fill(&mut salt).expect("OS RNG unavailable");
    let hash = compute_hash(password, &salt, spin_count);
    (general_purpose::STANDARD.encode(salt), hash)
}

/// Inserts the `<workbookProtection>` element into workbook.xml.
///
/// Respects the strict XSD element order:
/// the element must be inserted BEFORE `<bookViews>` or `<sheets>`.
fn inject_protection(xml_content: &[u8], salt: &str, hash: &str, spin_count: u32) -> Result<Vec<u8>, ProtectionError> {
    let mut reader = Reader::from_reader(xml_content);

    let mut writer = Writer::new(Cursor::new(Vec::new()));
    let mut buf = Vec::new();
    let mut inserted = false;

    let protection_elem_str = format!(
        r#"<workbookProtection lockStructure="1" workbookAlgorithmName="SHA-512" workbookHashValue="{}" workbookSaltValue="{}" workbookSpinCount="{}"/>"#,
        hash, salt, spin_count
    );

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) => {
                let name_str = std::str::from_utf8(e.name().into_inner())?;

                // Insertion point: before sheets, bookViews, or functionGroups
                if !inserted
                    && (name_str == "sheets"
                        || name_str == "bookViews"
                        || name_str == "functionGroups")
                {
                    write_protection_tag(&mut writer, &protection_elem_str)?;
                    inserted = true;
                }

                // If an existing workbookProtection tag exists, replace it
                if name_str == "workbookProtection" {
                    let target_name = name_str.as_bytes().to_vec();
                    if !inserted {
                        write_protection_tag(&mut writer, &protection_elem_str)?;
                        inserted = true;
                    }

                    // Skip the existing tag and its content
                    let mut depth = 1;
                    let mut skip_buf = Vec::new();
                    loop {
                        match reader.read_event_into(&mut skip_buf) {
                            Ok(Event::Start(e)) if e.name().into_inner() == target_name => {
                                depth += 1
                            }
                            Ok(Event::End(e)) if e.name().into_inner() == target_name => {
                                depth -= 1;
                                if depth == 0 {
                                    break;
                                }
                            }
                            Ok(Event::Eof) => break,
                            Err(_) => break,
                            _ => {}
                        }
                        skip_buf.clear();
                    }
                    continue;
                }
                writer.write_event(Event::Start(e.clone()))?;
            }
            Ok(Event::Empty(ref e)) => {
                let name_str = std::str::from_utf8(e.name().into_inner())?;

                // Insertion point for self-closing tags
                if !inserted && (name_str == "sheets" || name_str == "bookViews") {
                    write_protection_tag(&mut writer, &protection_elem_str)?;
                    inserted = true;
                }

                // Replace existing workbookProtection tag
                if name_str == "workbookProtection" {
                    if !inserted {
                        write_protection_tag(&mut writer, &protection_elem_str)?;
                        inserted = true;
                    }
                    continue;
                }
                writer.write_event(Event::Empty(e.clone()))?;
            }
            Ok(Event::Eof) => break,
            Ok(e) => writer.write_event(e)?,
            Err(e) => return Err(e.into()),
        }
        buf.clear();
    }
    Ok(writer.into_inner().into_inner())
}

/// Helper to write the XML protection tag via quick_xml.
fn write_protection_tag<W: std::io::Write>(writer: &mut Writer<W>, tag: &str) -> Result<(), ProtectionError> {
    let mut temp_reader = Reader::from_str(tag);
    temp_reader.config_mut().trim_text_start = true;
    temp_reader.config_mut().trim_text_end = true;
    loop {
        match temp_reader.read_event() {
            Ok(Event::Eof) => break,
            Ok(e) => writer.write_event(e)?,
            Err(e) => return Err(e.into()),
        }
    }
    Ok(())
}
