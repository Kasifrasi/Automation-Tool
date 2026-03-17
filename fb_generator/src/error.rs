//! Crate-level error types for report generation.

use crate::workbook_protection::ProtectionError;
use rust_xlsxwriter::XlsxError;
use std::fmt;

/// Top-level error for report generation.
///
/// Returned by [`ReportConfig::write_to`](crate::config::ReportConfig::write_to),
/// [`ReportConfig::write_to_precomputed`](crate::config::ReportConfig::write_to_precomputed),
/// and the internal engine functions.
///
/// # Example
///
/// ```ignore
/// use fb_generator::ReportConfig;
///
/// let config = ReportConfig::default();
/// match config.write_to("report.xlsx") {
///     Ok(()) => println!("Report written"),
///     Err(e) => {
///         eprintln!("Failed: {e}");
///         if let Some(source) = std::error::Error::source(&e) {
///             eprintln!("Caused by: {source}");
///         }
///     }
/// }
/// ```
#[derive(Debug)]
pub enum ReportError {
    /// Excel write failure (from `rust_xlsxwriter`)
    Xlsx(XlsxError),
    /// Workbook protection failure (ZIP/XML manipulation)
    Protection(ProtectionError),
    /// I/O error (temp file, persist, etc.)
    Io(std::io::Error),
    /// Output path contains non-UTF-8 characters
    InvalidPath(String),
    /// Input validation error (positions, table index, panel index)
    Validation(String),
}

impl fmt::Display for ReportError {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Xlsx(e) => write!(f, "Excel error: {e}"),
            Self::Protection(e) => write!(f, "Workbook protection error: {e}"),
            Self::Io(e) => write!(f, "I/O error: {e}"),
            Self::InvalidPath(p) => write!(f, "Path contains non-UTF-8 characters: {p}"),
            Self::Validation(msg) => write!(f, "Validation error: {msg}"),
        }
    }
}

impl std::error::Error for ReportError {
    fn source(&self) -> Option<&(dyn std::error::Error + 'static)> {
        match self {
            Self::Xlsx(e) => Some(e),
            Self::Protection(e) => Some(e),
            Self::Io(e) => Some(e),
            Self::InvalidPath(_) => None,
            Self::Validation(_) => None,
        }
    }
}

impl From<XlsxError> for ReportError {
    fn from(e: XlsxError) -> Self {
        Self::Xlsx(e)
    }
}

impl From<ProtectionError> for ReportError {
    fn from(e: ProtectionError) -> Self {
        Self::Protection(e)
    }
}

impl From<std::io::Error> for ReportError {
    fn from(e: std::io::Error) -> Self {
        Self::Io(e)
    }
}

impl From<tempfile::PersistError> for ReportError {
    fn from(e: tempfile::PersistError) -> Self {
        Self::Io(e.error)
    }
}
