//! API cell definitions
//!
//! All API cells are defined centrally here. The macro generates:
//! - `ApiKey` enum with all variants
//! - `ApiKey::static_addr()` method for compile-time-known addresses
//! - `ApiKey::all_static_keys()` iterator
//!
//! To add a new API cell:
//! 1. Single cell: `NewField => (row, col),` in the `single` block
//! 2. Range cell:  `NewField => (start_row, col, count),` in the `range` block
//!    Done — everything else is generated automatically.
//!
//! ## Dynamic positions (cost categories)
//!
//! Cost positions use `ApiKey::Position`:
//!
//! - **position = 0**: Header-input mode
//!   - Only for categories with 0 positions in BodyConfig
//!   - Input directly in the header row (D, E, F, H)
//!   - Description not available (C is the VLOOKUP label)
//!
//! - **position = 1..N**: Position mode
//!   - For categories with 1+ positions in BodyConfig
//!   - Input in separate position rows
//!   - All fields available (C, D, E, F, H)

use crate::report::core::CellAddr;

// =============================================================================
// PositionField — fields of a cost position
// =============================================================================

/// Fields of a cost position in the body area.
///
/// Each position has 5 input fields in columns C, D, E, F, H.
/// (G is the ratio formula and is not exposed as an API field.)
///
/// ## Header-input mode (position=0)
///
/// In header-input mode (position=0), `Description` is NOT available
/// because column C contains the VLOOKUP label. `BodyLayout::position_addr()`
/// returns `None` for `Description` when position=0.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum PositionField {
    /// Description (column C) — only available when position >= 1!
    Description,
    /// Approved budget (column D)
    Approved,
    /// Income report period (column E)
    IncomeReport,
    /// Income total (column F)
    IncomeTotal,
    /// Reason / remark (column H)
    Remark,
}

impl PositionField {
    /// Returns the column number (0-based).
    pub const fn col(&self) -> u16 {
        match self {
            Self::Description => 2,  // C
            Self::Approved => 3,     // D
            Self::IncomeReport => 4, // E
            Self::IncomeTotal => 5,  // F
            Self::Remark => 7,       // H
        }
    }

    /// Returns all 5 PositionFields as an array.
    pub const fn all() -> [PositionField; 5] {
        [
            Self::Description,
            Self::Approved,
            Self::IncomeReport,
            Self::IncomeTotal,
            Self::Remark,
        ]
    }

    /// Returns all PositionFields available in header-input mode (position=0).
    ///
    /// `Description` is not available in header-input mode (C is the VLOOKUP label).
    pub const fn header_input_fields() -> [PositionField; 4] {
        [
            Self::Approved,
            Self::IncomeReport,
            Self::IncomeTotal,
            Self::Remark,
        ]
    }

    /// Returns whether this field is available in header-input mode (position=0).
    pub const fn available_at_header_input(&self) -> bool {
        !matches!(self, Self::Description)
    }
}

/// Macro for defining all API cells.
///
/// Generates:
/// - `ApiKey` enum
/// - `ApiKey::static_addr()` method
/// - `ApiKey::all_static_keys()` iterator
macro_rules! define_api_cells {
    (
        single {
            $( $single_name:ident => ($single_row:expr, $single_col:expr) ),* $(,)?
        }
        range {
            $( $range_name:ident => ($range_start_row:expr, $range_col:expr, $range_count:expr) ),* $(,)?
        }
    ) => {
        /// Key for API values.
        ///
        /// Each variant corresponds to exactly one input cell in the financial report.
        /// Addresses are 0-based (row, col).
        ///
        /// **Static keys** (single/range): address known at compile time.
        /// **Dynamic keys** (Position, Footer): address computed at runtime from BodyLayout/FooterLayout.
        #[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
        pub enum ApiKey {
            // Single cells (static)
            $( $single_name, )*
            // Range cells with index (static)
            $( $range_name(u8), )*
            /// Dynamic cost position.
            ///
            /// Address computed at runtime from `BodyLayout`.
            ///
            /// ## Position values
            ///
            /// - `position = 0`: Header-input mode
            ///   - Only for categories with 0 positions in BodyConfig
            ///   - `Description` not available!
            ///
            /// - `position = 1..N`: Normal position
            ///   - For categories with 1+ positions in BodyConfig
            ///   - All fields available
            Position {
                /// Category number (1-8)
                category: u8,
                /// Position: 0 = header-input, 1..N = position row
                position: u16,
                /// Which field of the position
                field: PositionField,
            },
            /// Footer input fields (Bank, Kasse, Sonstiges).
            ///
            /// Address computed at runtime from `FooterLayout`.
            Footer(FooterField),
        }

        impl ApiKey {
            /// Returns the cell address for static keys (0-based).
            ///
            /// Returns `None` for dynamic keys (`Position`, `Footer`) —
            /// these require a `BodyLayout`/`FooterLayout` for address resolution.
            pub const fn static_addr(&self) -> Option<CellAddr> {
                match self {
                    // Single cells
                    $( Self::$single_name => Some(CellAddr::new($single_row, $single_col)), )*
                    // Range cells
                    $( Self::$range_name(i) => Some(CellAddr::new($range_start_row + *i as u32, $range_col)), )*
                    // Dynamic keys have no fixed address
                    Self::Position { .. } => None,
                    Self::Footer(_) => None,
                }
            }

            /// Returns whether this key has a runtime-computed address.
            pub const fn is_dynamic(&self) -> bool {
                matches!(self, Self::Position { .. } | Self::Footer(_))
            }

            /// Returns all static API keys as an iterator.
            ///
            /// `Position` and `Footer` keys are not included as they are dynamic.
            pub fn all_static_keys() -> impl Iterator<Item = ApiKey> {
                let singles = [
                    $( Self::$single_name, )*
                ].into_iter();

                let ranges = [
                    $( (Self::$range_name as fn(u8) -> ApiKey, $range_count as u8), )*
                ].into_iter().flat_map(|(constructor, count)| {
                    (0..count).map(move |i| constructor(i))
                });

                singles.chain(ranges)
            }

            /// Total number of static API cells.
            pub const fn static_count() -> usize {
                // Count single cells
                let singles = [$( stringify!($single_name), )*].len();
                // Sum range cells
                let ranges = 0 $( + $range_count )*;
                singles + ranges
            }

        }

    };
}

// =============================================================================
// API cell definitions — single source of truth
// =============================================================================
//
// To add a new cell:
// - Single cell: Name => (row, col),
// - Range cell:  Name => (start_row, col, count),
//
// row/col are 0-based! E2 = (1, 4), D15 = (14, 3), L14 = (13, 11)

define_api_cells! {
    single {
        // Header area
        Language => (1, 4),        // E2: language selection
        Currency => (2, 4),        // E3: currency selection
        ProjectNumber => (4, 3),   // D5: project number
        ProjectTitle => (5, 3),    // D6: project title
        ProjectStart => (7, 4),    // E8: project start
        ProjectEnd => (7, 6),      // G8: project end
        ReportStart => (8, 4),     // E9: report period start
        ReportEnd => (8, 6),       // G9: report period end
    }
    range {
        // Table: rows 15-19 (index 0-4)
        ApprovedBudget => (14, 3, 5),      // D15-D19: approved budget
        IncomeReportPeriod => (14, 4, 5),  // E15-E19: income report period
        IncomeTotal => (14, 5, 5),         // F15-F19: income total
        IncomeReason => (14, 7, 5),        // H15-H19: reason

        // Left panel: rows 14-31 (index 0-17)
        LeftDate => (13, 11, 18),          // L14-L31: date
        LeftAmountEuro => (13, 12, 18),    // M14-M31: amount Euro
        LeftAmountLocal => (13, 13, 18),   // N14-N31: amount local

        // Right panel: rows 14-31 (index 0-17)
        RightDate => (13, 18, 18),         // S14-S31: date
        RightAmountEuro => (13, 19, 18),   // T14-T31: amount Euro
        RightAmountLocal => (13, 20, 18),  // U14-U31: amount local
    }
}

// =============================================================================
// Footer API keys (dynamic addresses, depend on body layout)
// =============================================================================

/// Footer input field in the balance reconciliation area.
///
/// These fields are located in the footer after the dynamic body area.
/// Exact row positions are computed at runtime from `FooterLayout`.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum FooterField {
    /// Bank balance (column E, footer row 7)
    Bank,
    /// Cash balance (column E, footer row 8)
    Kasse,
    /// Other balances (column E, footer row 9)
    /// (uncashed checks, advances, loans, etc.)
    Sonstiges,
}

impl FooterField {
    /// Returns all 3 FooterFields as an array.
    pub const fn all() -> [FooterField; 3] {
        [Self::Bank, Self::Kasse, Self::Sonstiges]
    }

    /// Returns the 0-based index in the `input_rows` array of FooterLayout.
    pub const fn index(&self) -> usize {
        match self {
            Self::Bank => 0,
            Self::Kasse => 1,
            Self::Sonstiges => 2,
        }
    }

    /// Returns the column number (always column E = index 4).
    pub const fn col(&self) -> u16 {
        4 // column E
    }
}

// =============================================================================
// Tests
// =============================================================================

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_single_cell_addresses() {
        assert_eq!(ApiKey::Language.static_addr().unwrap(), CellAddr::new(1, 4)); // E2
        assert_eq!(ApiKey::Currency.static_addr().unwrap(), CellAddr::new(2, 4)); // E3
        assert_eq!(ApiKey::ProjectNumber.static_addr().unwrap(), CellAddr::new(4, 3)); // D5
        assert_eq!(ApiKey::ProjectStart.static_addr().unwrap(), CellAddr::new(7, 4)); // E8
        assert_eq!(ApiKey::ReportEnd.static_addr().unwrap(), CellAddr::new(8, 6)); // G9
    }

    #[test]
    fn test_range_cell_addresses() {
        // ApprovedBudget: D15-D19
        assert_eq!(ApiKey::ApprovedBudget(0).static_addr().unwrap(), CellAddr::new(14, 3)); // D15
        assert_eq!(ApiKey::ApprovedBudget(4).static_addr().unwrap(), CellAddr::new(18, 3)); // D19

        // LeftDate: L14-L31
        assert_eq!(ApiKey::LeftDate(0).static_addr().unwrap(), CellAddr::new(13, 11)); // L14
        assert_eq!(ApiKey::LeftDate(17).static_addr().unwrap(), CellAddr::new(30, 11)); // L31

        // RightAmountLocal: U14-U31
        assert_eq!(ApiKey::RightAmountLocal(0).static_addr().unwrap(), CellAddr::new(13, 20)); // U14
        assert_eq!(ApiKey::RightAmountLocal(17).static_addr().unwrap(), CellAddr::new(30, 20)); // U31
    }

    #[test]
    fn test_all_keys_count() {
        let count = ApiKey::all_static_keys().count();
        // 8 single cells + 4*5 table + 6*18 panel = 8 + 20 + 108 = 136
        assert_eq!(count, 136);
        assert_eq!(ApiKey::static_count(), 136);
    }

    #[test]
    fn test_all_keys_iteration() {
        let keys: Vec<_> = ApiKey::all_static_keys().collect();

        // First 8 are single cells
        assert_eq!(keys[0], ApiKey::Language);
        assert_eq!(keys[7], ApiKey::ReportEnd);

        // Then come the range cells
        assert_eq!(keys[8], ApiKey::ApprovedBudget(0));
        assert_eq!(keys[12], ApiKey::ApprovedBudget(4));
    }

    #[test]
    fn test_position_key() {
        use PositionField::*;

        let key = ApiKey::Position {
            category: 1,
            position: 3,
            field: Description,
        };

        // Position keys are dynamic
        assert!(key.is_dynamic());
        assert_eq!(key.static_addr(), None);

        // Static keys are not dynamic
        assert!(!ApiKey::Language.is_dynamic());
        assert!(ApiKey::Language.static_addr().is_some());
    }

    #[test]
    fn test_position_field() {
        use PositionField::*;

        assert_eq!(Description.col(), 2); // C
        assert_eq!(Approved.col(), 3); // D
        assert_eq!(IncomeReport.col(), 4); // E
        assert_eq!(IncomeTotal.col(), 5); // F
        assert_eq!(Remark.col(), 7); // H

        // all() returns all 5 fields
        assert_eq!(PositionField::all().len(), 5);

        // header_input_fields() returns 4 fields (without Description)
        assert_eq!(PositionField::header_input_fields().len(), 4);
    }

    #[test]
    fn test_position_field_availability() {
        use PositionField::*;

        // Description not available in header-input mode
        assert!(!Description.available_at_header_input());

        // All others are available
        assert!(Approved.available_at_header_input());
        assert!(IncomeReport.available_at_header_input());
        assert!(IncomeTotal.available_at_header_input());
        assert!(Remark.available_at_header_input());
    }

    #[test]
    fn test_header_input_position() {
        use PositionField::*;

        // Header-input: position=0
        let key = ApiKey::Position {
            category: 6,
            position: 0,
            field: Approved,
        };

        assert!(key.is_dynamic());

        // Normal position: position=1
        let key = ApiKey::Position {
            category: 1,
            position: 1,
            field: Description,
        };

        assert!(key.is_dynamic());
    }
}
