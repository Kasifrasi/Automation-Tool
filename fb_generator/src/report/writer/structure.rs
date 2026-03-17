//! Report structure: merges, blanks, static strings, validations
//!
//! Formulas and input values are written by `write_cells_from_bridge()` in `engine.rs`.
//! This module only writes the Excel structure that must exist before the content.

use super::layout::MergeRange;
use crate::report::api::Currency;
use crate::report::body::{BodyLayout, CategoryLayout, CategoryMode, FooterLayout};
use crate::report::styles::FormatMatrix;
use rust_xlsxwriter::{DataValidation, Worksheet, XlsxError};

/// Writes the complete report structure (merges, blanks, static strings).
pub fn write_structure(
    ws: &mut Worksheet,
    fmt: &FormatMatrix,
    body_layout: &BodyLayout,
    footer_layout: &FooterLayout,
    suffix: &str,
    language: &str,
) -> Result<(), XlsxError> {
    // Static header area (rows 0–25): header, income table, right panel, pre-body
    write_header(ws, fmt, suffix, language)?;
    write_table(ws, fmt)?;
    write_panel(ws, fmt)?;
    write_prebody(ws, fmt)?;

    // Dynamic body area (row 26+): cost categories
    write_body(ws, fmt, body_layout)?;

    // Dynamic footer area: balance reconciliation
    write_footer(ws, fmt, footer_layout)?;

    Ok(())
}

// ============================================================================
// Header (Rows 0-9)
// ============================================================================

const HEADER_MERGES: &[MergeRange] = &[
    MergeRange::new(0, 1, 0, 2),  // B1:C1
    MergeRange::new(1, 1, 1, 2),  // B2:C2
    MergeRange::new(1, 9, 2, 14), // J2:O3
    MergeRange::new(2, 1, 2, 2),  // B3:C3
    MergeRange::new(3, 9, 3, 14), // J4:O4
    MergeRange::new(4, 1, 4, 2),  // B5:C5
    MergeRange::new(5, 1, 6, 2),  // B6:C7
    MergeRange::new(5, 3, 6, 7),  // D6:H7
    MergeRange::new(7, 1, 7, 2),  // B8:C8
    MergeRange::new(7, 6, 7, 7),  // G8:H8
    MergeRange::new(8, 1, 8, 2),  // B9:C9
    MergeRange::new(8, 6, 8, 7),  // G9:H9
];

const HEADER_BLANKS: &[(u32, u16)] = &[
    (0, 10),
    (0, 11),
    (0, 12),
    (0, 13),
    (0, 14),
    (2, 4),
    (4, 3),
    (6, 9),
    (7, 4),
    (7, 9),
    (8, 4),
    (8, 6),
    (8, 9),
];

fn write_header(
    ws: &mut Worksheet,
    fmt: &FormatMatrix,
    suffix: &str,
    language: &str,
) -> Result<(), XlsxError> {
    write_merges(ws, fmt, HEADER_MERGES)?;
    write_blanks(ws, fmt, HEADER_BLANKS)?;

    // Suffix (B2)
    write_string(ws, fmt, 1, 1, suffix)?;

    // Language (E2)
    write_string(ws, fmt, 1, 4, language)?;

    // Data Validations
    let lang_validation =
        DataValidation::new().allow_list_formula("=Sprachversionen!$B$1:$B$5".into());
    ws.add_data_validation(1, 4, 1, 4, &lang_validation)?;

    let currency_count = Currency::all().len();
    let currency_formula = format!("=Sprachversionen!$A$1:$A${}", currency_count);
    let currency_validation =
        DataValidation::new().allow_list_formula(currency_formula.as_str().into());
    ws.add_data_validation(2, 4, 2, 4, &currency_validation)?;

    Ok(())
}

// ============================================================================
// Table (Rows 10-19)
// ============================================================================

const TABLE_MERGES: &[MergeRange] = &[
    // Header
    MergeRange::new(10, 1, 10, 2), // B11:C11
    MergeRange::new(10, 3, 13, 3), // D11:D14
    MergeRange::new(10, 4, 13, 4), // E11:E14
    MergeRange::new(10, 5, 13, 5), // F11:F14
    MergeRange::new(10, 6, 13, 6), // G11:G14
    MergeRange::new(10, 7, 13, 7), // H11:H14
    MergeRange::new(11, 1, 11, 2), // B12:C12
    MergeRange::new(12, 1, 12, 2), // B13:C13
    MergeRange::new(13, 1, 13, 2), // B14:C14
    // Body
    MergeRange::new(14, 1, 14, 2), // B15:C15
    MergeRange::new(15, 1, 15, 2), // B16:C16
    MergeRange::new(16, 1, 16, 2), // B17:C17
    MergeRange::new(17, 1, 17, 2), // B18:C18
    MergeRange::new(18, 1, 18, 2), // B19:C19
    MergeRange::new(19, 1, 19, 2), // B20:C20
];

const TABLE_BLANKS: &[(u32, u16)] = &[
    // Header
    (10, 1),
    (10, 2),
    (13, 1),
    (13, 2),
    // Body
    (14, 4),
    (14, 5),
    (14, 6),
    (14, 7),
    (15, 4),
    (15, 5),
    (15, 7),
    (16, 4),
    (16, 5),
    (16, 7),
    (17, 4),
    (17, 5),
    (17, 7),
    (18, 4),
    (18, 5),
    (18, 7),
    (19, 7),
];

fn write_table(ws: &mut Worksheet, fmt: &FormatMatrix) -> Result<(), XlsxError> {
    write_merges(ws, fmt, TABLE_MERGES)?;
    write_blanks(ws, fmt, TABLE_BLANKS)?;

    // D15-D19: Default 0.0
    for row in 14..=18 {
        if let Some(format) = fmt.get(row, 3) {
            ws.write_number_with_format(row, 3, 0.0, format)?;
        }
    }

    Ok(())
}

// ============================================================================
// Panel (rows 10-30, cols J-V)
// ============================================================================

const PANEL_MERGES: &[MergeRange] = &[
    MergeRange::new(10, 9, 10, 10),  // J11:K11
    MergeRange::new(10, 16, 10, 17), // Q11:R11
];

fn write_panel(ws: &mut Worksheet, fmt: &FormatMatrix) -> Result<(), XlsxError> {
    write_merges(ws, fmt, PANEL_MERGES)?;

    // Row 10: L-O and S-V blanks
    for col in [11u16, 12, 13, 14] {
        write_blank(ws, fmt, 10, col)?;
    }
    for col in [18u16, 19, 20, 21] {
        write_blank(ws, fmt, 10, col)?;
    }

    // Row 11: J-O and Q-V blanks
    for col in [9u16, 10, 11, 12, 13, 14] {
        write_blank(ws, fmt, 11, col)?;
    }
    for col in [16u16, 17, 18, 19, 20, 21] {
        write_blank(ws, fmt, 11, col)?;
    }

    // Row 12: J, K, Q, R blanks
    for col in [9u16, 10] {
        write_blank(ws, fmt, 12, col)?;
    }
    for col in [16u16, 17] {
        write_blank(ws, fmt, 12, col)?;
    }

    // Body: 18 rows (row 13-30) — index strings + blanks
    for i in 0..18u32 {
        let row = 13 + i;
        let left_num = 1 + i;
        let right_num = 19 + i;

        // Left: J index + L/M/N blanks
        write_string(ws, fmt, row, 9, &format!("{}. ", left_num))?;
        for col in [11u16, 12, 13] {
            write_blank(ws, fmt, row, col)?;
        }

        // Right: Q index + S/T/U blanks
        write_string(ws, fmt, row, 16, &format!("{}. ", right_num))?;
        for col in [18u16, 19, 20] {
            write_blank(ws, fmt, row, col)?;
        }
    }

    Ok(())
}

// ============================================================================
// Prebody (Rows 20-25)
// ============================================================================

fn write_prebody(ws: &mut Worksheet, fmt: &FormatMatrix) -> Result<(), XlsxError> {
    ws.set_row_height(20, 13.5)?;

    // D-H (cols 3-7): vertical merges rows 22-25 (locked)
    for col in 3u16..=7u16 {
        if let Some(format) = fmt.get_locked(22, col) {
            ws.merge_range(22, col, 25, col, "", format)?;
        }
    }

    // B22, C22: blanks
    write_blank(ws, fmt, 22, 1)?;
    write_blank(ws, fmt, 22, 2)?;

    // B:C merges rows 23-24 (locked, formulas via CalcBridge)
    for row in 23..=24 {
        if let Some(format) = fmt.get_locked(row, 1) {
            ws.merge_range(row, 1, row, 2, "", format)?;
        }
    }

    // B26:C26 merged blank
    if let Some(format) = fmt.get(25, 1) {
        ws.merge_range(25, 1, 25, 2, "", format)?;
    }

    Ok(())
}

// ============================================================================
// Body (dynamic)
// ============================================================================

fn write_body(
    ws: &mut Worksheet,
    fmt: &FormatMatrix,
    layout: &BodyLayout,
) -> Result<(), XlsxError> {
    for cat in &layout.categories {
        write_category(ws, fmt, cat)?;
    }

    // Total row: B:C merge (locked) + H blank
    if let Some(format) = fmt.get_locked(layout.total_row, 1) {
        ws.merge_range(layout.total_row, 1, layout.total_row, 2, "", format)?;
    }
    write_blank(ws, fmt, layout.total_row, 7)?;

    Ok(())
}

fn write_category(
    ws: &mut Worksheet,
    fmt: &FormatMatrix,
    cat: &CategoryLayout,
) -> Result<(), XlsxError> {
    match &cat.mode {
        CategoryMode::HeaderInput { row } => {
            write_string(ws, fmt, *row, 1, &format!("{}.", cat.meta.num))?;
        }
        CategoryMode::WithPositions {
            header_row,
            positions,
            footer_row,
        } => {
            // Header: number + D-H blanks
            write_string(ws, fmt, *header_row, 1, &format!("{}.", cat.meta.num))?;
            for col in 3..=7 {
                write_blank(ws, fmt, *header_row, col)?;
            }

            // Positions: numbers
            for i in 0..positions.count {
                let row = positions.start_row + i as u32;
                write_string(ws, fmt, row, 1, &format!("{}.{}", cat.meta.num, i + 1))?;
            }

            // Footer: B:C merge (locked) + H blank
            if let Some(format) = fmt.get_locked(*footer_row, 1) {
                ws.merge_range(*footer_row, 1, *footer_row, 2, "", format)?;
            }
            write_blank(ws, fmt, *footer_row, 7)?;
        }
    }
    Ok(())
}

// ============================================================================
// Footer (dynamic)
// ============================================================================

fn write_footer(
    ws: &mut Worksheet,
    fmt: &FormatMatrix,
    layout: &FooterLayout,
) -> Result<(), XlsxError> {
    let s = layout.start_row;

    // Row 0: blanks B-D + E:E+1 merge (locked)
    for col in 1..=3 {
        write_blank(ws, fmt, s, col)?;
    }
    if let Some(format) = fmt.get_locked(s, 4) {
        ws.merge_range(s, 4, s + 1, 4, "", format)?;
    }

    // Row 1: B:D merge (locked)
    if let Some(format) = fmt.get_locked(s + 1, 1) {
        ws.merge_range(s + 1, 1, s + 1, 3, "", format)?;
    }

    // Row 2: blanks B-D
    for col in 1..=3 {
        write_blank(ws, fmt, s + 2, col)?;
    }

    // Row 3: blanks B-E
    for col in 1..=4 {
        write_blank(ws, fmt, s + 3, col)?;
    }

    // Row 4: blank C
    write_blank(ws, fmt, s + 4, 2)?;

    // Row 5: blanks B-E
    for col in 1..=4 {
        write_blank(ws, fmt, s + 5, col)?;
    }

    // Row 6: blanks C, D
    write_blank(ws, fmt, s + 6, 2)?;
    write_blank(ws, fmt, s + 6, 3)?;

    // Rows 7-8: blanks C, D
    for i in 7..=8 {
        write_blank(ws, fmt, s + i, 2)?;
        write_blank(ws, fmt, s + i, 3)?;
    }

    // Row 9: blanks C, D
    write_blank(ws, fmt, s + 9, 2)?;
    write_blank(ws, fmt, s + 9, 3)?;

    // Row 19: blank C + blanks E-G
    write_blank(ws, fmt, s + 19, 2)?;
    for col in 4..=6 {
        write_blank(ws, fmt, s + 19, col)?;
    }

    Ok(())
}

// ============================================================================
// Generic helpers
// ============================================================================

/// Writes merge ranges with unlocked format from [`FormatMatrix`].
fn write_merges(
    ws: &mut Worksheet,
    fmt: &FormatMatrix,
    merges: &[MergeRange],
) -> Result<(), XlsxError> {
    for range in merges {
        if let Some(format) = fmt.get(range.first_row, range.first_col) {
            ws.merge_range(
                range.first_row,
                range.first_col,
                range.last_row,
                range.last_col,
                "",
                format,
            )?;
        }
    }
    Ok(())
}

/// Writes multiple blank cells from a const array.
fn write_blanks(
    ws: &mut Worksheet,
    fmt: &FormatMatrix,
    blanks: &[(u32, u16)],
) -> Result<(), XlsxError> {
    for &(row, col) in blanks {
        write_blank(ws, fmt, row, col)?;
    }
    Ok(())
}

/// Writes a single blank cell with format.
fn write_blank(
    ws: &mut Worksheet,
    fmt: &FormatMatrix,
    row: u32,
    col: u16,
) -> Result<(), XlsxError> {
    if let Some(format) = fmt.get(row, col) {
        ws.write_blank(row, col, format)?;
    }
    Ok(())
}

/// Writes a string with optional format.
fn write_string(
    ws: &mut Worksheet,
    fmt: &FormatMatrix,
    row: u32,
    col: u16,
    value: &str,
) -> Result<(), XlsxError> {
    if let Some(format) = fmt.get_locked(row, col) {
        ws.write_string_with_format(row, col, value, format)?;
    } else {
        ws.write_string(row, col, value)?;
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use crate::report::body::FooterLayout;

    #[test]
    fn test_footer_layout_compute() {
        let layout = FooterLayout::compute(100);
        assert_eq!(layout.start_row, 103);
        assert_eq!(layout.saldo_row, 107);
        assert_eq!(layout.input_rows[0], 110);
        assert_eq!(layout.input_rows[1], 111);
        assert_eq!(layout.input_rows[2], 112);
        assert_eq!(layout.end_row, 123);
    }
}
