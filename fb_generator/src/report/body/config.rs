//! Body configuration
//!
//! Defines the number of position rows per category (1–8).
//!
//! ## Flexible categories
//!
//! Each category can operate in one of two modes:
//!
//! - **0 positions** (header-input mode):
//!   Single row — input directly in columns D, E, F, H. No footer.
//!
//! - **1+ positions** (multi-row mode):
//!   Header row (category label) + N position rows (input) + footer row (sums).

use crate::report::api::Category;
use std::collections::HashMap;

/// Start row of the dynamic body area (0-based). Excel row 27 = index 26.
pub const BODY_START_ROW: u32 = 26;

/// All category numbers (1–8)
pub const ALL_CATEGORIES: [u8; 8] = [1, 2, 3, 4, 5, 6, 7, 8];

/// Configuration for the dynamic body area.
///
/// ## Example
/// ```ignore
/// let config = BodyConfig::new()
///     .with_positions(1, 10)   // Category 1: 10 position rows under header
///     .with_positions(2, 5)    // Category 2: 5 positions
///     .with_positions(6, 0)    // Category 6: header-input (0 positions)
///     .with_positions(7, 1);   // Category 7: 1 position under header
/// ```
///
/// ## Serde (requires `serde` feature)
///
/// Serialized as a flat JSON object via `from`/`into` `HashMap<u8, u16>`:
///
/// ```json
/// { "1": 20, "2": 20, "3": 30, "4": 30, "5": 20, "6": 0, "7": 0, "8": 0 }
/// ```
///
/// Keys are category numbers (1–8) as strings, values are the number of position
/// rows. `0` means header-input mode (single-row, no footer).
///
/// **Note:** In [`ReportConfig`](crate::ReportConfig), this is derived automatically
/// from the `positions` and `header_inputs` fields — it is not exposed in the public API.
///
/// Defaults: see [`BodyConfig::default_positions`].
#[derive(Debug, Clone, PartialEq, Eq)]
#[cfg_attr(feature = "serde", derive(serde::Serialize, serde::Deserialize))]
#[cfg_attr(feature = "serde", serde(from = "HashMap<u8, u16>", into = "HashMap<u8, u16>"))]
pub struct BodyConfig {
    /// Number of position rows per category (`0` = header-input mode)
    positions: HashMap<u8, u16>,
}

impl BodyConfig {
    /// Creates a new configuration with default values
    pub fn new() -> Self {
        Self::default()
    }

    /// Sets the number of position rows for a category
    ///
    /// - `count = 0`: header-input (single-row mode)
    /// - `count >= 1`: multi-row mode with N positions under the header
    pub fn with_positions(mut self, category: u8, count: u16) -> Self {
        if ALL_CATEGORIES.contains(&category) {
            self.positions.insert(category, count);
        }
        self
    }

    /// Sets the number of position rows using the [`Category`] enum
    pub fn with_cat_positions(self, category: Category, count: u16) -> Self {
        self.with_positions(category.index(), count)
    }

    /// Returns the number of position rows for a category
    ///
    /// - `0`: header-input (no separate position rows)
    /// - `1+`: number of position rows under the header
    pub fn position_count(&self, category: u8) -> u16 {
        *self.positions.get(&category).unwrap_or(&0)
    }

    /// Returns `true` if the category is in header-input mode (0 positions)
    pub fn is_header_input(&self, category: u8) -> bool {
        self.position_count(category) == 0
    }

    /// Returns `true` if the category has position rows (1+ positions)
    pub fn has_positions(&self, category: u8) -> bool {
        self.position_count(category) > 0
    }
}

impl Default for BodyConfig {
    fn default() -> Self {
        Self {
            positions: Self::default_positions(),
        }
    }
}

impl BodyConfig {
    /// Returns the default position counts (cat 1–5: 20/30 rows, cat 6–8: header-input)
    pub fn default_positions() -> HashMap<u8, u16> {
        HashMap::from([
            (1, 20),
            (2, 20),
            (3, 30),
            (4, 30),
            (5, 20),
            (6, 0), // header-input
            (7, 0), // header-input
            (8, 0), // header-input
        ])
    }
}

#[cfg(feature = "serde")]
impl From<HashMap<u8, u16>> for BodyConfig {
    fn from(positions: HashMap<u8, u16>) -> Self {
        BodyConfig { positions }
    }
}

#[cfg(feature = "serde")]
impl From<BodyConfig> for HashMap<u8, u16> {
    fn from(config: BodyConfig) -> Self {
        config.positions
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_default_config() {
        let config = BodyConfig::default();

        // Multi-row categories
        assert_eq!(config.position_count(1), 20);
        assert_eq!(config.position_count(3), 30);
        assert!(config.has_positions(1));
        assert!(!config.is_header_input(1));

        // Header-input categories
        assert_eq!(config.position_count(6), 0);
        assert_eq!(config.position_count(7), 0);
        assert_eq!(config.position_count(8), 0);
        assert!(!config.has_positions(6));
        assert!(config.is_header_input(6));
    }

    #[test]
    fn test_custom_config() {
        let config = BodyConfig::new()
            .with_positions(1, 10)
            .with_positions(6, 5); // cat 6 now with positions

        assert_eq!(config.position_count(1), 10);
        assert_eq!(config.position_count(6), 5);
        assert!(config.has_positions(6));
        assert!(!config.is_header_input(6));
    }

    #[test]
    fn test_header_input_mode() {
        let config = BodyConfig::new()
            .with_positions(1, 0) // cat 1 in header-input mode
            .with_positions(2, 1); // cat 2 with 1 position

        assert!(config.is_header_input(1));
        assert!(!config.is_header_input(2));
        assert_eq!(config.position_count(2), 1);
    }

    #[test]
    fn test_flexible_categories() {
        // Any category can switch between header-input and multi-row mode
        let config = BodyConfig::new()
            .with_positions(1, 0)  // normally multi-row → now header-input
            .with_positions(6, 10) // normally header-input → now multi-row
            .with_positions(7, 1)  // 1 position
            .with_positions(8, 0); // header-input (default)

        assert!(config.is_header_input(1));
        assert!(config.has_positions(6));
        assert!(config.has_positions(7));
        assert!(config.is_header_input(8));
    }
}
