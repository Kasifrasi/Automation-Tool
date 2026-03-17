//! Type-safe types for the financial report API
//!
//! This module ensures only valid values can be passed to the API.
//! Invalid values are caught at compile time (enums) or at runtime with clear errors.

use crate::lang::data::TEXT_MATRIX;
use std::fmt;
use std::str::FromStr;

// ============================================================================
// Language — type-safe language selection
// ============================================================================

/// Supported languages for the financial report.
///
/// This enum ensures only valid languages can be used. The available
/// languages are derived from the internal text matrix.
///
/// # Example
/// ```ignore
/// use fb_generator::Language;
///
/// let lang = Language::Deutsch;
/// assert_eq!(lang.as_str(), "deutsch");
/// ```
///
/// ## Serde (requires `serde` feature)
///
/// Serialized with `rename_all = "lowercase"`:
///
/// ```json
/// "deutsch"
/// "english"
/// "francais"
/// "espanol"
/// "portugues"
/// ```
///
/// [`FromStr`] accepts additional aliases (`"german"`, `"de"`, `"fr"`, etc.),
/// but serde only accepts the exact lowercase variant names above.
///
/// Default: `Deutsch`.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
#[cfg_attr(feature = "serde", derive(serde::Serialize, serde::Deserialize))]
#[cfg_attr(feature = "serde", serde(rename_all = "lowercase"))]
pub enum Language {
    /// Deutsch
    #[default]
    Deutsch,
    /// English
    English,
    /// Français
    Francais,
    /// Español
    Espanol,
    /// Português
    Portugues,
}

impl Language {
    /// Returns the string value as used in `TEXT_MATRIX`
    pub fn as_str(&self) -> &'static str {
        match self {
            Language::Deutsch => "deutsch",
            Language::English => "english",
            Language::Francais => "français",
            Language::Espanol => "español",
            Language::Portugues => "português",
        }
    }

    /// Returns all available languages
    pub fn all() -> &'static [Language] {
        &[
            Language::Deutsch,
            Language::English,
            Language::Francais,
            Language::Espanol,
            Language::Portugues,
        ]
    }

    /// Validates that the language exists in `TEXT_MATRIX`
    pub fn validate(&self) -> bool {
        let lang_str = self.as_str();
        TEXT_MATRIX
            .iter()
            .any(|row| !row.is_empty() && row[0].eq_ignore_ascii_case(lang_str))
    }
}

impl fmt::Display for Language {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(f, "{}", self.as_str())
    }
}

impl FromStr for Language {
    type Err = String;

    /// Parses a language string (case-insensitive)
    ///
    /// Accepts various formats like "German", "de", "Deutsch", etc.
    fn from_str(s: &str) -> Result<Self, Self::Err> {
        let lower = s.to_lowercase();
        match lower.as_str() {
            "deutsch" | "german" | "de" => Ok(Language::Deutsch),
            "english" | "englisch" | "en" => Ok(Language::English),
            "français" | "francais" | "french" | "fr" => Ok(Language::Francais),
            "español" | "espanol" | "spanish" | "es" => Ok(Language::Espanol),
            "português" | "portugues" | "portuguese" | "pt" => Ok(Language::Portugues),
            _ => Err(format!("Unknown language: {}", s)),
        }
    }
}

// ============================================================================
// Currency — ISO 4217 enum (defined in `lang::data`, re-exported here)
// ============================================================================

pub use crate::lang::data::Currency;

// ============================================================================
// Category — cost categories
// ============================================================================

/// Cost categories of the financial report.
///
/// Represents the 8 main categories. Used with [`BodyConfig::with_cat_positions`](crate::BodyConfig::with_cat_positions)
/// for the typed Rust API.
///
/// # Example
/// ```ignore
/// use fb_generator::{BodyConfig, Category};
///
/// let config = BodyConfig::new()
///     .with_cat_positions(Category::Bauausgaben, 10)
///     .with_cat_positions(Category::Evaluierung, 0);
/// ```
///
/// ## Serde (requires `serde` feature)
///
/// Serialized using PascalCase variant names:
///
/// | JSON value | Index | Description |
/// |---|---|---|
/// | `"Bauausgaben"` | 1 | Building Expenses |
/// | `"Investitionen"` | 2 | Non-Recurring Expenses (Investments) |
/// | `"Personalausgaben"` | 3 | Personnel Expenses |
/// | `"Projektaktivitaeten"` | 4 | Ongoing Expenses for Project Activities |
/// | `"Projektverwaltung"` | 5 | Project Administration |
/// | `"Evaluierung"` | 6 | Evaluation |
/// | `"Audit"` | 7 | Audit |
/// | `"Reserve"` | 8 | Reserve |
///
/// **Note:** In [`ReportConfig`](crate::ReportConfig) JSON, categories are referenced
/// by numeric keys (`"1"` – `"8"`) as keys of the `positions` and `header_inputs` maps.
/// The `Category` enum is for the typed Rust API only.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[cfg_attr(feature = "serde", derive(serde::Serialize, serde::Deserialize))]
pub enum Category {
    /// Category 1: Building Expenses
    Bauausgaben = 1,
    /// Category 2: Non-Recurring Expenses (Investments)
    Investitionen = 2,
    /// Category 3: Personnel Expenses
    Personalausgaben = 3,
    /// Category 4: Ongoing Expenses for Project Activities
    Projektaktivitaeten = 4,
    /// Category 5: Project Administration
    Projektverwaltung = 5,
    /// Category 6: Evaluation
    Evaluierung = 6,
    /// Category 7: Audit
    Audit = 7,
    /// Category 8: Reserve
    Reserve = 8,
}

impl Category {
    /// Returns the numeric index (1–8)
    pub fn index(&self) -> u8 {
        *self as u8
    }

    /// Creates a category from a numeric index.
    ///
    /// Returns `None` if the index is not 1–8.
    pub fn from_index(idx: u8) -> Option<Self> {
        match idx {
            1 => Some(Category::Bauausgaben),
            2 => Some(Category::Investitionen),
            3 => Some(Category::Personalausgaben),
            4 => Some(Category::Projektaktivitaeten),
            5 => Some(Category::Projektverwaltung),
            6 => Some(Category::Evaluierung),
            7 => Some(Category::Audit),
            8 => Some(Category::Reserve),
            _ => None,
        }
    }

    /// Returns all 8 categories
    pub fn all() -> &'static [Category] {
        &[
            Category::Bauausgaben,
            Category::Investitionen,
            Category::Personalausgaben,
            Category::Projektaktivitaeten,
            Category::Projektverwaltung,
            Category::Evaluierung,
            Category::Audit,
            Category::Reserve,
        ]
    }

    /// German display name (matches TEXT_MATRIX labels)
    pub fn name_de(&self) -> &'static str {
        match self {
            Category::Bauausgaben => "Bauausgaben",
            Category::Investitionen => "Einmalige Ausgaben (Investitionen)",
            Category::Personalausgaben => "Personalkosten",
            Category::Projektaktivitaeten => "Laufende Ausgaben für Projektaktivitäten",
            Category::Projektverwaltung => "Projektverwaltung",
            Category::Evaluierung => "Evaluierung",
            Category::Audit => "Audit",
            Category::Reserve => "Reserve",
        }
    }

    /// English display name (matches TEXT_MATRIX labels)
    pub fn name_en(&self) -> &'static str {
        match self {
            Category::Bauausgaben => "Building Expenses",
            Category::Investitionen => "Non-Recurring Expenses (Investments)",
            Category::Personalausgaben => "Personnel Expenses",
            Category::Projektaktivitaeten => "Ongoing Expenses for Project Activities",
            Category::Projektverwaltung => "Project Administration",
            Category::Evaluierung => "Evaluation",
            Category::Audit => "Audit",
            Category::Reserve => "Reserve",
        }
    }
}

impl fmt::Display for Category {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(f, "{}", self.name_de())
    }
}

// ============================================================================
// ReportDate — validated date
// ============================================================================

/// Validated date for financial reports.
///
/// Ensures only valid dates are used. Supports multiple input formats.
/// Valid year range: 1900–2100.
///
/// # Example
/// ```ignore
/// use fb_generator::ReportDate;
///
/// let date = ReportDate::new(2024, 1, 15).unwrap();
/// assert_eq!(date.format_de(), "15.01.2024");
/// assert_eq!(date.format_iso(), "2024-01-15");
/// ```
///
/// ## Serde (requires `serde` feature)
///
/// Deserialized via `try_from = "String"` using [`ReportDate::parse`], which accepts:
/// - `"2024-06-15"` — ISO 8601
/// - `"15.06.2024"` — German (DD.MM.YYYY)
/// - `"15/06/2024"` — EU (DD/MM/YYYY)
/// - `"06/15/2024"` — US (MM/DD/YYYY, tried as fallback)
///
/// Always serialized to ISO format:
/// ```json
/// "2024-06-15"
/// ```
///
/// **Note:** In [`ReportConfig`](crate::ReportConfig), date fields (`project_start`,
/// `report_end`, etc.) are plain `Option<String>`, not `ReportDate`. The `ReportDate`
/// type is for the typed Rust API via [`ReportValues`](crate::ReportValues).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[cfg_attr(feature = "serde", derive(serde::Serialize, serde::Deserialize))]
#[cfg_attr(feature = "serde", serde(try_from = "String", into = "String"))]
pub struct ReportDate {
    year: u16,
    month: u8,
    day: u8,
}

/// Error from date construction.
///
/// Returned by [`ReportDate::new`] and [`ReportDate::parse`].
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum DateError {
    /// Invalid month (must be 1–12)
    InvalidMonth(u8),
    /// Invalid day for the given month
    InvalidDay { day: u8, month: u8, year: u16 },
    /// Invalid year (must be 1900–2100)
    InvalidYear(u16),
    /// Unparseable date string
    ParseError(String),
}

impl fmt::Display for DateError {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            DateError::InvalidMonth(m) => write!(f, "invalid month: {} (must be 1-12)", m),
            DateError::InvalidDay { day, month, year } => {
                write!(f, "invalid day: {}.{}.{}", day, month, year)
            }
            DateError::InvalidYear(y) => write!(f, "invalid year: {}", y),
            DateError::ParseError(s) => write!(f, "failed to parse date: {}", s),
        }
    }
}

impl std::error::Error for DateError {}

impl ReportDate {
    /// Creates a new validated date.
    pub fn new(year: u16, month: u8, day: u8) -> Result<Self, DateError> {
        // Year validation (reasonable range for financial reports)
        if !(1900..=2100).contains(&year) {
            return Err(DateError::InvalidYear(year));
        }

        // Month validation
        if !(1..=12).contains(&month) {
            return Err(DateError::InvalidMonth(month));
        }

        // Day validation
        let days_in_month = Self::days_in_month(year, month);
        if day < 1 || day > days_in_month {
            return Err(DateError::InvalidDay { day, month, year });
        }

        Ok(ReportDate { year, month, day })
    }

    /// Parses a date from various formats.
    ///
    /// Supported formats:
    /// - `DD.MM.YYYY` (German)
    /// - `YYYY-MM-DD` (ISO)
    /// - `DD/MM/YYYY` (EU, tried first)
    /// - `MM/DD/YYYY` (US, fallback)
    pub fn parse(s: &str) -> Result<Self, DateError> {
        let s = s.trim();

        // Try different formats
        if let Some(date) = Self::try_parse_de(s) {
            return Ok(date);
        }
        if let Some(date) = Self::try_parse_iso(s) {
            return Ok(date);
        }
        if let Some(date) = Self::try_parse_slash(s) {
            return Ok(date);
        }

        Err(DateError::ParseError(s.to_string()))
    }

    fn try_parse_de(s: &str) -> Option<Self> {
        // DD.MM.YYYY
        let parts: Vec<&str> = s.split('.').collect();
        if parts.len() != 3 {
            return None;
        }
        let day: u8 = parts[0].parse().ok()?;
        let month: u8 = parts[1].parse().ok()?;
        let year: u16 = parts[2].parse().ok()?;
        Self::new(year, month, day).ok()
    }

    fn try_parse_iso(s: &str) -> Option<Self> {
        // YYYY-MM-DD
        let parts: Vec<&str> = s.split('-').collect();
        if parts.len() != 3 {
            return None;
        }
        let year: u16 = parts[0].parse().ok()?;
        let month: u8 = parts[1].parse().ok()?;
        let day: u8 = parts[2].parse().ok()?;
        Self::new(year, month, day).ok()
    }

    fn try_parse_slash(s: &str) -> Option<Self> {
        // DD/MM/YYYY or MM/DD/YYYY
        let parts: Vec<&str> = s.split('/').collect();
        if parts.len() != 3 {
            return None;
        }
        let first: u8 = parts[0].parse().ok()?;
        let second: u8 = parts[1].parse().ok()?;
        let year: u16 = parts[2].parse().ok()?;

        // Try DD/MM/YYYY first
        if let Ok(date) = Self::new(year, second, first) {
            return Some(date);
        }
        // Then MM/DD/YYYY
        Self::new(year, first, second).ok()
    }

    fn days_in_month(year: u16, month: u8) -> u8 {
        match month {
            1 | 3 | 5 | 7 | 8 | 10 | 12 => 31,
            4 | 6 | 9 | 11 => 30,
            2 => {
                if Self::is_leap_year(year) {
                    29
                } else {
                    28
                }
            }
            _ => 0,
        }
    }

    fn is_leap_year(year: u16) -> bool {
        (year.is_multiple_of(4) && !year.is_multiple_of(100)) || year.is_multiple_of(400)
    }

    /// Formats the date in German format (DD.MM.YYYY)
    pub fn format_de(&self) -> String {
        format!("{:02}.{:02}.{}", self.day, self.month, self.year)
    }

    /// Formats the date in ISO format (YYYY-MM-DD)
    pub fn format_iso(&self) -> String {
        format!("{}-{:02}-{:02}", self.year, self.month, self.day)
    }

    /// Formats the date in US format (MM/DD/YYYY)
    pub fn format_us(&self) -> String {
        format!("{:02}/{:02}/{}", self.month, self.day, self.year)
    }

    /// Returns the year
    pub fn year(&self) -> u16 {
        self.year
    }

    /// Returns the month
    pub fn month(&self) -> u8 {
        self.month
    }

    /// Returns the day
    pub fn day(&self) -> u8 {
        self.day
    }
}

impl fmt::Display for ReportDate {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(f, "{}", self.format_de())
    }
}

#[cfg(feature = "serde")]
impl From<ReportDate> for String {
    fn from(d: ReportDate) -> Self {
        d.format_iso()
    }
}

#[cfg(feature = "serde")]
impl TryFrom<String> for ReportDate {
    type Error = DateError;
    fn try_from(s: String) -> Result<Self, Self::Error> {
        Self::parse(&s)
    }
}

// ============================================================================
// Tests
// ============================================================================

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_language_as_str() {
        assert_eq!(Language::Deutsch.as_str(), "deutsch");
        assert_eq!(Language::English.as_str(), "english");
        assert_eq!(Language::Francais.as_str(), "français");
    }

    #[test]
    fn test_language_from_str() {
        assert_eq!(Language::from_str("deutsch"), Ok(Language::Deutsch));
        assert_eq!(Language::from_str("German"), Ok(Language::Deutsch));
        assert_eq!(Language::from_str("english"), Ok(Language::English));
        assert_eq!(Language::from_str("englisch"), Ok(Language::English)); // German spelling
        assert!(Language::from_str("invalid").is_err());
    }

    #[test]
    fn test_language_validate() {
        for lang in Language::all() {
            assert!(lang.validate(), "Language {:?} should be valid", lang);
        }
    }

    #[test]
    fn test_currency_parse() {
        assert_eq!("EUR".parse::<Currency>(), Ok(Currency::EUR));
        assert_eq!("eur".parse::<Currency>(), Ok(Currency::EUR)); // case-insensitive
        assert_eq!("USD".parse::<Currency>(), Ok(Currency::USD));
        assert!("INVALID".parse::<Currency>().is_err());
        assert!("".parse::<Currency>().is_err());
    }

    #[test]
    fn test_currency_as_str() {
        assert_eq!(Currency::EUR.as_str(), "EUR");
        assert_eq!(Currency::USD.as_str(), "USD");
        assert_eq!(Currency::GBP.as_str(), "GBP");
    }

    #[test]
    fn test_currency_all() {
        let all = Currency::all();
        assert_eq!(all.len(), 153);
        assert!(all.contains(&Currency::EUR));
        assert!(all.contains(&Currency::USD));
    }

    #[test]
    fn test_category_index() {
        assert_eq!(Category::Bauausgaben.index(), 1);
        assert_eq!(Category::Reserve.index(), 8);
    }

    #[test]
    fn test_category_from_index() {
        assert_eq!(Category::from_index(1), Some(Category::Bauausgaben));
        assert_eq!(Category::from_index(8), Some(Category::Reserve));
        assert_eq!(Category::from_index(0), None);
        assert_eq!(Category::from_index(9), None);
    }

    #[test]
    fn test_date_new_valid() {
        assert!(ReportDate::new(2024, 1, 15).is_ok());
        assert!(ReportDate::new(2024, 2, 29).is_ok()); // leap year
        assert!(ReportDate::new(2023, 2, 28).is_ok());
    }

    #[test]
    fn test_date_new_invalid() {
        assert!(ReportDate::new(2024, 13, 1).is_err()); // invalid month
        assert!(ReportDate::new(2024, 2, 30).is_err()); // invalid day
        assert!(ReportDate::new(2023, 2, 29).is_err()); // not a leap year
        assert!(ReportDate::new(1800, 1, 1).is_err()); // year out of range
    }

    #[test]
    fn test_date_parse() {
        // German format
        let date = ReportDate::parse("15.01.2024").unwrap();
        assert_eq!(date.day(), 15);
        assert_eq!(date.month(), 1);
        assert_eq!(date.year(), 2024);

        // ISO Format
        let date = ReportDate::parse("2024-01-15").unwrap();
        assert_eq!(date.format_iso(), "2024-01-15");

        // invalid
        assert!(ReportDate::parse("invalid").is_err());
    }

    #[test]
    fn test_date_format() {
        let date = ReportDate::new(2024, 1, 15).unwrap();
        assert_eq!(date.format_de(), "15.01.2024");
        assert_eq!(date.format_iso(), "2024-01-15");
        assert_eq!(date.format_us(), "01/15/2024");
    }
}
