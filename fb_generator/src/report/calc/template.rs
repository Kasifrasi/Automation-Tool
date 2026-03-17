//! Master template for the IronCalc model
//!
//! Contains all static formulas and the language sheet data.
//! Created once, then cloned per report via `to_bytes()` / `from_bytes()`.

use crate::lang::data::TEXT_MATRIX;
use crate::report::core::CellAddr;
use ironcalc_base::Model;

/// Sheet index for "Bericht" (the report sheet)
const SHEET_BERICHT: u32 = 0;
/// Sheet index for "Sprachversionen" (the language sheet)
const SHEET_LANG: u32 = 1;

/// Cached IronCalc model with all static formulas.
///
/// Contains:
/// - Sheet 0 "Bericht": all static formulas (header, table, panel, pre-body)
/// - Sheet 1 "Sprachversionen": full TEXT_MATRIX (5 languages × 66 terms)
///
/// Created once, then cloned per report via `to_bytes()` / `from_bytes()`.
pub(crate) struct ModelTemplate {
    bytes: Vec<u8>,
    /// Addresses of all static formula cells (consumed by `write_cells_from_bridge`)
    static_formula_cells: Vec<CellAddr>,
}

impl ModelTemplate {
    /// Creates the master template with all static formulas and the language sheet.
    pub(crate) fn new() -> Self {
        let mut model =
            Model::new_empty("Bericht", "en", "UTC", "en").expect("Failed to create IronCalc model");

        // Add and populate the Sprachversionen sheet
        model
            .add_sheet("Sprachversionen")
            .expect("Failed to add Sprachversionen sheet");
        populate_language_sheet(&mut model);

        // Register all static formulas on the Bericht sheet and collect addresses
        let mut static_formula_cells = Vec::new();
        register_static_formulas(&mut model, &mut static_formula_cells);

        let bytes = model.to_bytes();
        Self {
            bytes,
            static_formula_cells,
        }
    }

    /// Creates a fresh Model instance from the serialized template bytes.
    pub(crate) fn instantiate(&self) -> Model<'static> {
        Model::from_bytes(&self.bytes, "en").expect("Failed to deserialize template")
    }

    /// Returns the addresses of all static formula cells.
    pub(crate) fn static_formula_cells(&self) -> &[CellAddr] {
        &self.static_formula_cells
    }

}

// ============================================================================
// Language sheet
// ============================================================================

/// Populates the "Sprachversionen" sheet with TEXT_MATRIX data.
///
/// Mapping: `TEXT_MATRIX[lang_row][col_idx]` → Sheet 1, Row `lang_row+1`, Col `col_idx+2`
/// - col_idx=0 → Col B = language name
/// - col_idx=1 → Col C = first text (VLOOKUP index 2)
/// - col_idx=N → Col (N+2)
fn populate_language_sheet(model: &mut Model) {
    for (row_idx, row) in TEXT_MATRIX.iter().enumerate() {
        let ic_row = row_idx as i32 + 1; // 1-based
        for (col_idx, text) in row.iter().enumerate() {
            let ic_col = col_idx as i32 + 2; // B=2, C=3, …
            model
                .set_user_input(SHEET_LANG, ic_row, ic_col, text.to_string())
                .expect("Failed to set language cell");
        }
    }
}

// ============================================================================
// Static formulas
// ============================================================================

/// Registers all static formulas on the "Bericht" sheet and collects their addresses.
fn register_static_formulas(model: &mut Model, addrs: &mut Vec<CellAddr>) {
    // ========================================================================
    // VLOOKUP text lookups
    // ========================================================================

    // Row 0: B1 (title), J1 (fillable fields info)
    set_vlookup(model, 0, 1, 2, addrs); // B1
    set_vlookup(model, 0, 9, 60, addrs); // J1

    // Row 1: D2 (language label with default), J2 (tip)
    set_vlookup_default(model, 1, 3, 27, "Chose your language", addrs); // D2
    set_vlookup(model, 1, 9, 61, addrs); // J2

    // Row 2: B3 (overview), D3 (local currency label)
    set_vlookup(model, 2, 1, 3, addrs); // B3
    set_vlookup(model, 2, 3, 28, addrs); // D3

    // Row 3: J4 — VLOOKUP evaluated here (URL), wrapped in HYPERLINK() by engine.rs
    set_vlookup(model, 3, 9, 62, addrs); // J4

    // Row 4: B5 (project number label)
    set_vlookup(model, 4, 1, 4, addrs); // B5

    // Row 5: B6 (project title label)
    set_vlookup(model, 5, 1, 26, addrs); // B6

    // Row 7: B8 (project duration), D8 (from), F8 (to)
    set_vlookup(model, 7, 1, 5, addrs); // B8
    set_vlookup(model, 7, 3, 7, addrs); // D8
    set_vlookup(model, 7, 5, 8, addrs); // F8

    // Row 8: B9 (reporting period), D9 (from), F9 (to), K8 (exchange rate label)
    set_vlookup(model, 8, 1, 6, addrs); // B9
    set_vlookup(model, 8, 3, 7, addrs); // D9
    set_vlookup(model, 8, 5, 8, addrs); // F9
    set_vlookup(model, 7, 10, 59, addrs); // K8

    // Row 10: table header
    set_vlookup(model, 10, 9, 19, addrs); // J11
    set_vlookup(model, 10, 16, 19, addrs); // Q11
    set_vlookup(model, 10, 3, 11, addrs); // D11
    set_vlookup(model, 10, 4, 12, addrs); // E11
    set_vlookup(model, 10, 5, 13, addrs); // F11
    set_vlookup(model, 10, 6, 14, addrs); // G11
    set_vlookup(model, 10, 7, 15, addrs); // H11

    // Row 11: B12 (INCOME)
    set_vlookup(model, 11, 1, 9, addrs); // B12

    // Row 12: B13, panel headers
    set_vlookup(model, 12, 1, 10, addrs); // B13
    set_vlookup(model, 12, 11, 22, addrs); // L13: Date
    set_vlookup(model, 12, 12, 63, addrs); // M13: Euro
    set_currency_or_lookup(model, 12, 13, 28, addrs); // N13: Currency
    set_vlookup(model, 12, 14, 58, addrs); // O13: Exchange rate
    set_vlookup(model, 12, 18, 22, addrs); // S13: Date
    set_vlookup(model, 12, 19, 63, addrs); // T13: Euro
    set_currency_or_lookup(model, 12, 20, 28, addrs); // U13: Currency
    set_vlookup(model, 12, 21, 58, addrs); // V13: Exchange rate

    // Row 14-18: B15-B19 (income table labels)
    set_vlookup(model, 14, 1, 16, addrs); // B15: Carry-forward
    set_vlookup(model, 15, 1, 17, addrs); // B16: Local contribution
    set_vlookup(model, 16, 1, 18, addrs); // B17: Third-party contributions
    set_vlookup(model, 17, 1, 19, addrs); // B18: Grant funds
    set_vlookup(model, 18, 1, 20, addrs); // B19: Interest income

    // ========================================================================
    // IFERROR division (G16–G19)
    // ========================================================================

    for i in 1..5u32 {
        let row = 14 + i; // rows 15–18
        set_iferror_division(model, CellAddr::new(row, 6), CellAddr::new(row, 5), CellAddr::new(row, 3), addrs);
    }

    // ========================================================================
    // Row 19: B20 (TOTAL), D20–G20
    // ========================================================================

    set_vlookup(model, 19, 1, 21, addrs); // B20

    set_sum_range(model, 19, 3, 14, 18, 3, addrs); // D20 = SUM(D15:D19)
    set_sum_range(model, 19, 4, 14, 18, 4, addrs); // E20 = SUM(E15:E19)
    set_sum_range(model, 19, 5, 14, 18, 5, addrs); // F20 = SUM(F15:F19)
    set_iferror_division(model, CellAddr::new(19, 6), CellAddr::new(19, 5), CellAddr::new(19, 3), addrs); // G20

    // ========================================================================
    // Panel body (rows 13–30, left + right panels)
    // ========================================================================

    for i in 0..18u32 {
        let row = 13 + i;
        let excel_row = row + 1;

        // Col K (10): receipt number VLOOKUP
        set_vlookup(model, row, 10, 23, addrs);

        // Col O (14): exchange rate = N/M (left panel)
        set_formula(
            model,
            row,
            14,
            format!(r#"=IF(M{}="","",N{}/M{})"#, excel_row, excel_row, excel_row),
            addrs,
        );

        // Col R (17): receipt number VLOOKUP
        set_vlookup(model, row, 17, 23, addrs);

        // Col V (21): exchange rate = U/T (right panel)
        set_formula(
            model,
            row,
            21,
            format!(r#"=IF(T{}="","",U{}/T{})"#, excel_row, excel_row, excel_row),
            addrs,
        );
    }

    // ========================================================================
    // Pre-body formulas (rows 22–24)
    // ========================================================================

    set_vlookup(model, 22, 3, 11, addrs); // D23
    set_vlookup(model, 22, 4, 25, addrs); // E23
    set_vlookup(model, 22, 5, 55, addrs); // F23
    set_vlookup(model, 22, 6, 56, addrs); // G23
    set_vlookup(model, 22, 7, 15, addrs); // H23
    set_vlookup(model, 23, 1, 24, addrs); // B24
    set_vlookup(model, 24, 1, 10, addrs); // B25
}

// ============================================================================
// Formula Helpers
// ============================================================================

/// Sets a VLOOKUP formula: `=IF($E$2="","",VLOOKUP($E$2,Sprachversionen!$B:$CD,{index},FALSE))`
fn set_vlookup(model: &mut Model, row: u32, col: u16, index: usize, addrs: &mut Vec<CellAddr>) {
    let formula = format!(
        r#"=IF($E$2="","",VLOOKUP($E$2,Sprachversionen!$B:$CD,{},FALSE))"#,
        index
    );
    set_formula(model, row, col, formula, addrs);
}

/// Sets a VLOOKUP formula with a default value when no language is selected.
fn set_vlookup_default(
    model: &mut Model,
    row: u32,
    col: u16,
    index: usize,
    default: &str,
    addrs: &mut Vec<CellAddr>,
) {
    let formula = format!(
        r#"=IF($E$2="","{}",VLOOKUP($E$2,Sprachversionen!$B:$CD,{},FALSE))"#,
        default, index
    );
    set_formula(model, row, col, formula, addrs);
}

/// Sets a currency-or-lookup formula: `=IF(E3="",VLOOKUP(...),E3)`
fn set_currency_or_lookup(
    model: &mut Model,
    row: u32,
    col: u16,
    _index: usize,
    addrs: &mut Vec<CellAddr>,
) {
    let formula =
        r#"=IF(E3="",VLOOKUP($E$2,Sprachversionen!$B:$CD,28,FALSE),E3)"#.to_string();
    set_formula(model, row, col, formula, addrs);
}

/// Sets an IFERROR division formula: `=IFERROR(num/denom,0)`
fn set_iferror_division(
    model: &mut Model,
    target: CellAddr,
    numerator: CellAddr,
    denominator: CellAddr,
    addrs: &mut Vec<CellAddr>,
) {
    let formula = format!("=IFERROR({}/{},0)", numerator.to_excel(), denominator.to_excel());
    set_formula(model, target.row, target.col, formula, addrs);
}

/// Sets a SUM formula over a contiguous range of rows in a single column.
fn set_sum_range(
    model: &mut Model,
    target_row: u32,
    target_col: u16,
    start_row: u32,
    end_row: u32,
    col: u16,
    addrs: &mut Vec<CellAddr>,
) {
    let start = CellAddr::new(start_row, col).to_excel();
    let end = CellAddr::new(end_row, col).to_excel();
    let formula = format!("=SUM({}:{})", start, end);
    set_formula(model, target_row, target_col, formula, addrs);
}

/// Low-level: sets a formula in IronCalc and tracks the cell address.
fn set_formula(
    model: &mut Model,
    row: u32,
    col: u16,
    formula: String,
    addrs: &mut Vec<CellAddr>,
) {
    model
        .update_cell_with_formula(SHEET_BERICHT, row as i32 + 1, col as i32 + 1, formula)
        .unwrap_or_else(|e| panic!("Failed to set formula at ({}, {}): {}", row, col, e));
    addrs.push(CellAddr::new(row, col));
}

// ============================================================================
// Tests
// ============================================================================

#[cfg(test)]
mod tests {
    use super::*;
    use ironcalc_base::cell::CellValue;

    #[test]
    fn test_template_creation() {
        let template = ModelTemplate::new();
        assert!(!template.bytes.is_empty());
    }

    #[test]
    fn test_template_static_formula_cells_tracked() {
        let template = ModelTemplate::new();
        // ~80 static formulas should be tracked
        assert!(
            template.static_formula_cells().len() > 70,
            "Expected >70 static formula cells, got {}",
            template.static_formula_cells().len()
        );

        // Verify known cells are tracked
        let addrs = template.static_formula_cells();
        assert!(addrs.contains(&CellAddr::new(0, 1)), "B1 should be tracked");
        assert!(addrs.contains(&CellAddr::new(1, 3)), "D2 should be tracked");
        assert!(
            addrs.contains(&CellAddr::new(19, 3)),
            "D20 (SUM) should be tracked"
        );
        assert!(
            addrs.contains(&CellAddr::new(15, 6)),
            "G16 (ratio) should be tracked"
        );
    }

    #[test]
    fn test_template_instantiation() {
        let template = ModelTemplate::new();
        let mut model = template.instantiate();

        // Set language and evaluate
        model
            .set_user_input(SHEET_BERICHT, 2, 5, "deutsch".to_string())
            .unwrap(); // E2
        model.evaluate();

        // B1 VLOOKUP should return the title
        let result = model.get_cell_value_by_index(SHEET_BERICHT, 1, 2).unwrap();
        assert_eq!(
            result,
            CellValue::String("F I N A N Z B E R I C H T".to_string()),
            "B1 VLOOKUP should return title for 'deutsch'"
        );
    }

    #[test]
    fn test_template_vlookup_english() {
        let template = ModelTemplate::new();
        let mut model = template.instantiate();

        model
            .set_user_input(SHEET_BERICHT, 2, 5, "english".to_string())
            .unwrap();
        model.evaluate();

        let result = model.get_cell_value_by_index(SHEET_BERICHT, 1, 2).unwrap();
        assert_eq!(
            result,
            CellValue::String("F I N A N C I A L  R E P O R T ".to_string()),
            "B1 VLOOKUP should return title for 'english' (trailing space in data)"
        );
    }

    #[test]
    fn test_template_default_language_label() {
        let template = ModelTemplate::new();
        let mut model = template.instantiate();

        // No language set → default "Chose your language"
        model.evaluate();

        let result = model.get_cell_value_by_index(SHEET_BERICHT, 2, 4).unwrap();
        assert_eq!(
            result,
            CellValue::String("Chose your language".to_string()),
            "D2 should show default when no language set"
        );
    }

    #[test]
    fn test_template_sum_and_ratio() {
        let template = ModelTemplate::new();
        let mut model = template.instantiate();

        // Set values in D15-D19
        model
            .set_user_input(SHEET_BERICHT, 15, 4, "100".to_string())
            .unwrap(); // D15
        model
            .set_user_input(SHEET_BERICHT, 16, 4, "200".to_string())
            .unwrap(); // D16
        model
            .set_user_input(SHEET_BERICHT, 17, 4, "300".to_string())
            .unwrap(); // D17
        // F16: set value for ratio formula
        model
            .set_user_input(SHEET_BERICHT, 16, 6, "150".to_string())
            .unwrap(); // F16

        model.evaluate();

        // D20 = SUM(D15:D19) = 100 + 200 + 300 = 600
        let d20 = model.get_cell_value_by_index(SHEET_BERICHT, 20, 4).unwrap();
        assert_eq!(d20, CellValue::Number(600.0));

        // G16 = IFERROR(F16/D16, 0) = 150/200 = 0.75
        let g16 = model.get_cell_value_by_index(SHEET_BERICHT, 16, 7).unwrap();
        assert_eq!(g16, CellValue::Number(0.75));
    }

    #[test]
    fn test_template_clone_isolation() {
        let template = ModelTemplate::new();

        // Two separate instances
        let mut model1 = template.instantiate();
        let mut model2 = template.instantiate();

        // Set different languages
        model1
            .set_user_input(SHEET_BERICHT, 2, 5, "deutsch".to_string())
            .unwrap();
        model2
            .set_user_input(SHEET_BERICHT, 2, 5, "english".to_string())
            .unwrap();

        model1.evaluate();
        model2.evaluate();

        let b1_de = model1
            .get_cell_value_by_index(SHEET_BERICHT, 1, 2)
            .unwrap();
        let b1_en = model2
            .get_cell_value_by_index(SHEET_BERICHT, 1, 2)
            .unwrap();

        assert_ne!(b1_de, b1_en, "Templates should be isolated");
    }

    #[test]
    fn test_currency_or_lookup() {
        let template = ModelTemplate::new();
        let mut model = template.instantiate();

        // No currency set, language "deutsch" → VLOOKUP(28) = currency text from language sheet
        model
            .set_user_input(SHEET_BERICHT, 2, 5, "deutsch".to_string())
            .unwrap();
        model.evaluate();

        let n13_no_currency = model
            .get_cell_value_by_index(SHEET_BERICHT, 13, 14)
            .unwrap();
        // Should return the currency text from VLOOKUP(28)
        assert!(
            !matches!(n13_no_currency, CellValue::None),
            "N13 should have a value when no currency set"
        );

        // Now with currency "USD" → E3 is used directly
        let mut model2 = template.instantiate();
        model2
            .set_user_input(SHEET_BERICHT, 2, 5, "deutsch".to_string())
            .unwrap();
        model2
            .set_user_input(SHEET_BERICHT, 3, 5, "USD".to_string())
            .unwrap();
        model2.evaluate();

        let n13_with_currency = model2
            .get_cell_value_by_index(SHEET_BERICHT, 13, 14)
            .unwrap();
        assert_eq!(
            n13_with_currency,
            CellValue::String("USD".to_string()),
            "N13 should show currency when E3 is set"
        );
    }
}
