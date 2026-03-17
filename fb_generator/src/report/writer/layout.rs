//! Layout configuration for the financial report
//!
//! Column widths, row heights, merge ranges, freeze panes, and page setup.

use rust_xlsxwriter::{Format, Worksheet, XlsxError};

// ============================================================================
// Merge range definition
// ============================================================================

/// A rectangular cell merge range, defined by inclusive row/column bounds.
pub struct MergeRange {
    pub first_row: u32,
    pub first_col: u16,
    pub last_row: u32,
    pub last_col: u16,
}

impl MergeRange {
    /// Creates a new `MergeRange` from the given inclusive bounds.
    pub const fn new(first_row: u32, first_col: u16, last_row: u32, last_col: u16) -> Self {
        Self {
            first_row,
            first_col,
            last_row,
            last_col,
        }
    }
}

// ============================================================================
// Sheet setup (column widths, row heights, zoom, etc.)
// ============================================================================

/// Configures the worksheet with zoom level, column widths, and row heights.
pub fn setup_sheet(ws: &mut Worksheet) -> Result<(), XlsxError> {
    // Page Setup
    ws.set_active(true);
    ws.set_zoom(85);
    ws.set_screen_gridlines(false);

    // Column widths
    setup_column_widths(ws)?;

    // Row heights
    setup_row_heights(ws)?;

    Ok(())
}

/// Column widths (A–Y) and default unlock for the first 50 columns.
///
/// All 50 columns (A–AX) are set to unlocked so that cells without an
/// explicit format are editable when sheet protection is enabled. Formulas
/// and static text override this with a locked format when they are written.
fn setup_column_widths(ws: &mut Worksheet) -> Result<(), XlsxError> {
    let unlocked = Format::new().set_unlocked();

    // Unlock first 50 columns (A–AX)
    for col in 0u16..50 {
        ws.set_column_format(col, &unlocked)?;
    }

    // Set explicit widths for report columns (A–Y)
    let widths: &[(u16, f64)] = &[
        (0, 2.60),   // A
        (1, 4.10),   // B
        (2, 56.0),   // C
        (3, 20.0),   // D
        (4, 18.0),   // E
        (5, 18.0),   // F
        (6, 12.0),   // G
        (7, 18.0),   // H
        (8, 4.0),    // I
        (9, 4.0),    // J
        (10, 19.14), // K
        (11, 10.85), // L
        (12, 15.0),  // M
        (13, 15.0),  // N
        (14, 15.0),  // O
        (15, 5.0),   // P
        (16, 4.0),   // Q
        (17, 19.14), // R
        (18, 10.85), // S
        (19, 15.0),  // T
        (20, 15.0),  // U
        (21, 15.0),  // V
        (22, 4.0),   // W
        (23, 36.71), // X
        (24, 10.85), // Y
    ];

    for &(col, width) in widths {
        ws.set_column_width(col, width)?;
    }

    Ok(())
}

fn setup_row_heights(ws: &mut Worksheet) -> Result<(), XlsxError> {
    ws.set_row_height(1, 15.0)?;
    ws.set_row_height(2, 15.0)?;
    ws.set_row_height(3, 12.0)?;
    ws.set_row_height(4, 15.0)?;
    ws.set_row_height(5, 15.0)?;
    ws.set_row_height(6, 15.0)?;
    ws.set_row_height(7, 15.0)?;
    ws.set_row_height(8, 15.0)?;
    ws.set_row_height(9, 13.5)?;
    ws.set_row_height(10, 12.6)?;
    Ok(())
}

// ============================================================================
// Freeze Panes
// ============================================================================

/// Freezes all rows above `row` so the header stays visible while scrolling.
pub fn setup_freeze_panes(ws: &mut Worksheet, row: u32) -> Result<(), XlsxError> {
    ws.set_freeze_panes(row, 0)?;
    Ok(())
}
