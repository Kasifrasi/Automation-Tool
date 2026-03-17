//! Body layout computation
//!
//! Computes row positions for all 8 categories based on [`BodyConfig`].
//!
//! ## Layout modes
//!
//! Each category operates in one of two modes:
//!
//! ### Header-input mode (0 positions)
//! ```text
//! | B     | C (Label)      | D        | E        | F        | G     | H        |
//! |-------|----------------|----------|----------|----------|-------|----------|
//! | 6.    | VLOOKUP Label  | [Input]  | [Input]  | [Input]  | Ratio | [Input]  |
//! ```
//! Single row — input directly in D, E, F, H.
//! API: `Position { category: 6, position: 0, field }`
//!
//! ### Multi-row mode (1+ positions)
//! ```text
//! | B     | C              | D        | E        | F        | G     | H        |
//! |-------|----------------|----------|----------|----------|-------|----------|
//! | 1.    | VLOOKUP Label  |          |          |          |       |          | ← Header
//! | 1.1   | [Input]        | [Input]  | [Input]  | [Input]  | Ratio | [Input]  | ← Position 1
//! | 1.2   | [Input]        | [Input]  | [Input]  | [Input]  | Ratio | [Input]  | ← Position 2
//! | ...   | ...            | ...      | ...      | ...      | ...   | ...      |
//! |       | VLOOKUP Sum    | SUM      | SUM      | SUM      | Ratio |          | ← Footer
//! ```
//! Header row (label) + N position rows (input) + footer row (sums).
//! API: `Position { category: 1, position: 1..N, field }` (1-based)

use super::config::{BodyConfig, ALL_CATEGORIES, BODY_START_ROW};
use crate::report::api::PositionField;
use crate::report::core::CellAddr;

/// Category metadata (VLOOKUP indices into the Sprachversionen sheet)
#[derive(Debug, Clone, Copy)]
pub struct CategoryMeta {
    /// Category number (1–8)
    pub num: u8,
    /// VLOOKUP index for the header/category label
    pub label_index: usize,
    /// VLOOKUP index for the footer/subtotal label (multi-row mode only)
    pub sum_label_index: usize,
}

/// VLOOKUP indices for all 8 categories (from Sprachversionen sheet):
/// - Cat 1–5: label/subtotal pairs (29/30, 31/32, 33/34, 35/36, 37/38)
/// - Cat 6 (Evaluierung): 39/64
/// - Cat 7 (Audit): 40/65
/// - Cat 8 (Reserve): 41/66
/// - Total label: index 42
pub const CATEGORY_META: [CategoryMeta; 8] = [
    CategoryMeta {
        num: 1,
        label_index: 29,
        sum_label_index: 30,
    },
    CategoryMeta {
        num: 2,
        label_index: 31,
        sum_label_index: 32,
    },
    CategoryMeta {
        num: 3,
        label_index: 33,
        sum_label_index: 34,
    },
    CategoryMeta {
        num: 4,
        label_index: 35,
        sum_label_index: 36,
    },
    CategoryMeta {
        num: 5,
        label_index: 37,
        sum_label_index: 38,
    },
    CategoryMeta {
        num: 6,
        label_index: 39,
        sum_label_index: 64, // Subtotal Evaluierung
    },
    CategoryMeta {
        num: 7,
        label_index: 40,
        sum_label_index: 65, // Subtotal Audit
    },
    CategoryMeta {
        num: 8,
        label_index: 41,
        sum_label_index: 66, // Subtotal Reserve
    },
];

/// VLOOKUP index for the "Total" label
pub const TOTAL_LABEL_INDEX: usize = 42;

/// Computed body layout — row positions for all 8 categories plus the total row.
#[derive(Debug, Clone)]
pub struct BodyLayout {
    /// Layout for each category (index 0 = cat 1, …, index 7 = cat 8)
    pub categories: Vec<CategoryLayout>,
    /// Row of the grand total
    pub total_row: u32,
    /// Last written row
    pub last_row: u32,
}

/// Layout of a single category
#[derive(Debug, Clone)]
pub struct CategoryLayout {
    /// Category metadata (number, VLOOKUP indices)
    pub meta: CategoryMeta,
    /// Operating mode (header-input or multi-row with positions)
    pub mode: CategoryMode,
}

/// Operating mode of a category
#[derive(Debug, Clone)]
pub enum CategoryMode {
    /// Header-input mode (0 positions) — the header row is the input row
    HeaderInput {
        /// The single row (header = input)
        row: u32,
    },
    /// Multi-row mode (1+ positions) — header + N positions + footer
    WithPositions {
        /// Header row (category label, no input)
        header_row: u32,
        /// Range of position rows
        positions: PositionRange,
        /// Footer row (subtotals)
        footer_row: u32,
    },
}

/// Contiguous range of position rows within a category
#[derive(Debug, Clone, Copy)]
pub struct PositionRange {
    /// First position row (0-based)
    pub start_row: u32,
    /// Last position row (0-based)
    pub end_row: u32,
    /// Number of positions
    pub count: u16,
}

impl BodyLayout {
    /// Computes the layout from a [`BodyConfig`]
    pub fn compute(config: &BodyConfig) -> Self {
        let mut current_row = BODY_START_ROW;
        let mut categories = Vec::with_capacity(8);

        for &cat_num in &ALL_CATEGORIES {
            let meta = CATEGORY_META[(cat_num - 1) as usize];
            let num_positions = config.position_count(cat_num);

            if num_positions == 0 {
                // Header-input mode: single row
                categories.push(CategoryLayout {
                    meta,
                    mode: CategoryMode::HeaderInput { row: current_row },
                });
                current_row += 1;
            } else {
                // Multi-row mode: header + N positions + footer
                let header_row = current_row;
                current_row += 1;

                let positions_start = current_row;
                let positions_end = current_row + num_positions as u32 - 1;
                current_row += num_positions as u32;

                let footer_row = current_row;
                current_row += 1;

                categories.push(CategoryLayout {
                    meta,
                    mode: CategoryMode::WithPositions {
                        header_row,
                        positions: PositionRange {
                            start_row: positions_start,
                            end_row: positions_end,
                            count: num_positions,
                        },
                        footer_row,
                    },
                });
            }
        }

        let total_row = current_row;

        Self {
            categories,
            total_row,
            last_row: total_row,
        }
    }

    /// Returns the layout for a category by number (1–8)
    pub fn category(&self, num: u8) -> Option<&CategoryLayout> {
        self.categories.iter().find(|c| c.meta.num == num)
    }

    /// Returns all footer rows (multi-row categories only), used for the grand total formula
    pub fn footer_rows(&self) -> Vec<u32> {
        self.categories
            .iter()
            .filter_map(|c| c.footer_row())
            .collect()
    }

    /// Returns all header-input rows, used for the grand total formula
    pub fn header_input_rows(&self) -> Vec<u32> {
        self.categories
            .iter()
            .filter_map(|c| c.header_input_row())
            .collect()
    }

    /// Returns all rows that need a ratio formula (col G)
    pub fn ratio_rows(&self) -> Vec<u32> {
        let mut rows = Vec::new();

        for cat in &self.categories {
            match &cat.mode {
                CategoryMode::HeaderInput { row } => {
                    rows.push(*row);
                }
                CategoryMode::WithPositions {
                    positions,
                    footer_row,
                    ..
                } => {
                    for row in positions.start_row..=positions.end_row {
                        rows.push(row);
                    }
                    rows.push(*footer_row);
                }
            }
        }

        rows.push(self.total_row);
        rows
    }

    /// Total number of rows in the body area
    pub fn row_count(&self) -> u32 {
        self.last_row - BODY_START_ROW + 1
    }

    /// Resolves a position field to a cell address
    ///
    /// Returns `None` if:
    /// - category does not exist
    /// - `position=0` but category has positions (use `1..N` instead)
    /// - `position>0` but category is header-input (use `0`)
    /// - position is out of range
    /// - `Description` field at `position=0` (column C holds the VLOOKUP label)
    pub fn position_addr(
        &self,
        category: u8,
        position: u16,
        field: PositionField,
    ) -> Option<CellAddr> {
        let cat = self.categories.iter().find(|c| c.meta.num == category)?;

        match &cat.mode {
            CategoryMode::HeaderInput { row } => {
                if position != 0 {
                    return None;
                }
                if matches!(field, PositionField::Description) {
                    return None;
                }
                Some(CellAddr::new(*row, field.col()))
            }
            CategoryMode::WithPositions { positions, .. } => {
                if position < 1 || position > positions.count {
                    return None;
                }
                let row = positions.start_row + (position - 1) as u32;
                Some(CellAddr::new(row, field.col()))
            }
        }
    }
}

impl CategoryLayout {
    /// Returns the header row (multi-row mode only)
    pub fn header_row(&self) -> Option<u32> {
        match &self.mode {
            CategoryMode::WithPositions { header_row, .. } => Some(*header_row),
            CategoryMode::HeaderInput { .. } => None,
        }
    }

    /// Returns the footer/subtotal row (multi-row mode only)
    pub fn footer_row(&self) -> Option<u32> {
        match &self.mode {
            CategoryMode::WithPositions { footer_row, .. } => Some(*footer_row),
            CategoryMode::HeaderInput { .. } => None,
        }
    }

    /// Returns the header-input row (header-input mode only)
    pub fn header_input_row(&self) -> Option<u32> {
        match &self.mode {
            CategoryMode::HeaderInput { row } => Some(*row),
            CategoryMode::WithPositions { .. } => None,
        }
    }

    /// Returns the position range (multi-row mode only)
    pub fn positions(&self) -> Option<&PositionRange> {
        match &self.mode {
            CategoryMode::WithPositions { positions, .. } => Some(positions),
            CategoryMode::HeaderInput { .. } => None,
        }
    }

    /// Returns `true` if this category is in header-input mode
    pub fn is_header_input(&self) -> bool {
        matches!(self.mode, CategoryMode::HeaderInput { .. })
    }

    /// Returns `true` if this category has position rows
    pub fn has_positions(&self) -> bool {
        matches!(self.mode, CategoryMode::WithPositions { .. })
    }

    /// Returns the row referenced by the grand total formula.
    /// Footer row for multi-row categories, header row for header-input categories.
    pub fn sum_row(&self) -> u32 {
        match &self.mode {
            CategoryMode::HeaderInput { row } => *row,
            CategoryMode::WithPositions { footer_row, .. } => *footer_row,
        }
    }
}

/// Footer layout with computed row positions for the balance reconciliation section.
#[derive(Debug, Clone)]
pub struct FooterLayout {
    /// Start row of the footer area (0-based)
    pub start_row: u32,
    /// Row for the balance difference formula (col E)
    pub saldo_row: u32,
    /// Input rows for Bank / Cash / Other (col E)
    pub input_rows: [u32; 3],
    /// Last row of the footer area
    pub end_row: u32,
}

impl FooterLayout {
    /// Computes the footer layout. The footer starts 3 rows after the body total row.
    pub fn compute(total_row: u32) -> Self {
        let start_row = total_row + 3;

        Self {
            start_row,
            saldo_row: start_row + 4,  // row 4: balance difference
            input_rows: [
                start_row + 7, // row 7: Bank
                start_row + 8, // row 8: Cash (Kasse)
                start_row + 9, // row 9: Other (Sonstiges)
            ],
            end_row: start_row + 20, // 21 rows (0–20)
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_layout_with_defaults() {
        let config = BodyConfig::default();
        let layout = BodyLayout::compute(&config);

        // Cat 1: multi-row (20 positions)
        let cat1 = layout.category(1).unwrap();
        assert!(cat1.has_positions());
        assert!(!cat1.is_header_input());

        if let CategoryMode::WithPositions {
            header_row,
            positions,
            footer_row,
        } = &cat1.mode
        {
            assert_eq!(*header_row, 26);
            assert_eq!(positions.start_row, 27);
            assert_eq!(positions.end_row, 46);
            assert_eq!(positions.count, 20);
            assert_eq!(*footer_row, 47);
        }

        // Cat 6: header-input (0 positions)
        let cat6 = layout.category(6).unwrap();
        assert!(cat6.is_header_input());
        assert!(!cat6.has_positions());
    }

    #[test]
    fn test_layout_with_custom_config() {
        let config = BodyConfig::new()
            .with_positions(1, 5)
            .with_positions(2, 10)
            .with_positions(6, 3); // cat 6 now with positions

        let layout = BodyLayout::compute(&config);

        // Cat 1: 5 positions
        let cat1 = layout.category(1).unwrap();
        assert_eq!(cat1.positions().unwrap().count, 5);

        // Cat 6: now with 3 positions
        let cat6 = layout.category(6).unwrap();
        assert!(cat6.has_positions());
        assert_eq!(cat6.positions().unwrap().count, 3);
    }

    #[test]
    fn test_header_input_mode() {
        let config = BodyConfig::new()
            .with_positions(1, 0) // cat 1 in header-input mode
            .with_positions(2, 5);

        let layout = BodyLayout::compute(&config);

        // Cat 1: header-input
        let cat1 = layout.category(1).unwrap();
        assert!(cat1.is_header_input());
        assert_eq!(cat1.header_input_row(), Some(26));

        // Cat 2: multi-row, starts right after cat 1
        let cat2 = layout.category(2).unwrap();
        assert!(cat2.has_positions());
        assert_eq!(cat2.header_row(), Some(27));
    }

    #[test]
    fn test_position_addr_header_input() {
        use PositionField::*;

        let config = BodyConfig::new().with_positions(1, 0);
        let layout = BodyLayout::compute(&config);

        // position=0 at header-input returns address
        let addr = layout.position_addr(1, 0, Approved).unwrap();
        assert_eq!(addr.row, 26);
        assert_eq!(addr.col, 3); // D

        // Description not available at header-input (C = VLOOKUP label)
        assert!(layout.position_addr(1, 0, Description).is_none());

        // position>0 not valid at header-input
        assert!(layout.position_addr(1, 1, Approved).is_none());
    }

    #[test]
    fn test_position_addr_with_positions() {
        use PositionField::*;

        let config = BodyConfig::new().with_positions(1, 5);
        let layout = BodyLayout::compute(&config);

        // position=0 not valid in multi-row mode
        assert!(layout.position_addr(1, 0, Approved).is_none());

        // position=1..5 valid
        let addr = layout.position_addr(1, 1, Description).unwrap();
        assert_eq!(addr.row, 27);
        assert_eq!(addr.col, 2); // C

        let addr = layout.position_addr(1, 5, IncomeTotal).unwrap();
        assert_eq!(addr.row, 31);
        assert_eq!(addr.col, 5); // F

        // position=6 out of range (only 5 positions)
        assert!(layout.position_addr(1, 6, Approved).is_none());
    }

    #[test]
    fn test_sum_rows() {
        let config = BodyConfig::new()
            .with_positions(1, 2)
            .with_positions(2, 0)
            .with_positions(3, 3);

        let layout = BodyLayout::compute(&config);

        // Cat 1: footer row is the sum row
        let cat1 = layout.category(1).unwrap();
        assert_eq!(cat1.sum_row(), cat1.footer_row().unwrap());

        // Cat 2: header-input row is the sum row
        let cat2 = layout.category(2).unwrap();
        assert_eq!(cat2.sum_row(), cat2.header_input_row().unwrap());
    }

    #[test]
    fn test_ratio_rows() {
        let config = BodyConfig::new()
            .with_positions(1, 2)
            .with_positions(2, 0)
            .with_positions(3, 0)
            .with_positions(4, 1)
            .with_positions(5, 0)
            .with_positions(6, 0)
            .with_positions(7, 0)
            .with_positions(8, 0);

        let layout = BodyLayout::compute(&config);
        let ratio_rows = layout.ratio_rows();

        // Cat 1: 2 positions + 1 footer = 3
        // Cat 2: 1 header-input = 1
        // Cat 3: 1 header-input = 1
        // Cat 4: 1 position + 1 footer = 2
        // Cat 5–8: 1 header-input each = 4
        // + 1 total = 1
        // Sum: 3 + 1 + 1 + 2 + 4 + 1 = 12
        assert_eq!(ratio_rows.len(), 12);
    }
}
