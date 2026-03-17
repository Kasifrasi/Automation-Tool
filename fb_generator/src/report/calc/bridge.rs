//! IronCalc bridge: dynamic formulas, input population, evaluation
//!
//! Per-report instance that wraps an IronCalc [`Model`]. Responsible for
//! registering dynamic body/footer formulas, populating input values from
//! [`ReportValues`], triggering evaluation, and exposing results for the writer.

use super::template::ModelTemplate;
use crate::report::api::{ApiKey, CellValue, FooterField, PositionField, ReportValues};
use crate::report::body::{
    BodyLayout, CategoryLayout, CategoryMode, FooterLayout, TOTAL_LABEL_INDEX,
};
use crate::report::core::CellAddr;
use ironcalc_base::cell::CellValue as IcCellValue;
use ironcalc_base::Model;

/// Sheet index for "Bericht" (the report sheet)
const SHEET: u32 = 0;

/// Bridge between [`ReportValues`] and the IronCalc [`Model`].
///
/// Responsible for:
/// 1. Adding dynamic formulas (body, footer)
/// 2. Populating input values
/// 3. Triggering evaluation
/// 4. Reading computed results
pub(crate) struct CalcBridge {
    model: Model<'static>,
    /// All cells containing formulas (consumed by `write_cells_from_bridge`)
    formula_cells: Vec<CellAddr>,
    /// Input values with their original data — written directly to Excel without
    /// IronCalc readback (IronCalc's `set_user_input` may auto-parse strings).
    input_values: Vec<(CellAddr, CellValue)>,
}

impl CalcBridge {
    /// Creates a bridge from a template.
    ///
    /// Copies the static formula addresses from the template so that
    /// `write_cells_from_bridge()` can write them into the Excel output.
    pub(crate) fn from_template(template: &ModelTemplate) -> Self {
        Self {
            model: template.instantiate(),
            formula_cells: template.static_formula_cells().to_vec(),
            input_values: Vec::new(),
        }
    }

    // ========================================================================
    // Dynamic body formulas
    // ========================================================================

    /// Registers dynamic body formulas based on the computed layout.
    pub(crate) fn register_body_formulas(&mut self, layout: &BodyLayout) {
        // 1. VLOOKUP formulas for category labels
        for cat in &layout.categories {
            match &cat.mode {
                CategoryMode::HeaderInput { row } => {
                    // Col C: VLOOKUP for category label
                    self.add_vlookup(*row, 2, cat.meta.label_index);
                }
                CategoryMode::WithPositions {
                    header_row,
                    footer_row,
                    ..
                } => {
                    // Header: col C
                    self.add_vlookup(*header_row, 2, cat.meta.label_index);
                    // Footer: col B
                    self.add_vlookup(*footer_row, 1, cat.meta.sum_label_index);
                }
            }
        }

        // Total: col B
        self.add_vlookup(layout.total_row, 1, TOTAL_LABEL_INDEX);

        // 2. Ratio formulas (col G)
        for row in layout.ratio_rows() {
            self.add_ratio(row);
        }

        // 3. SUM formulas for category footer rows (cols D, E, F)
        for cat in &layout.categories {
            if let CategoryMode::WithPositions {
                positions,
                footer_row,
                ..
            } = &cat.mode
            {
                for col in [3u16, 4, 5] {
                    self.add_sum_range(*footer_row, col, positions.start_row, positions.end_row);
                }
            }
        }

        // 4. SUM formulas for the total row (cols D, E, F)
        let sum_rows: Vec<u32> = layout
            .categories
            .iter()
            .map(|c: &CategoryLayout| c.sum_row())
            .collect();
        for col in [3u16, 4, 5] {
            self.add_sum_cells(layout.total_row, col, &sum_rows);
        }
    }

    // ========================================================================
    // Dynamic footer formulas
    // ========================================================================

    /// Registers dynamic footer formulas (labels, check formulas, balance).
    pub(crate) fn register_footer_formulas(
        &mut self,
        footer_layout: &FooterLayout,
        income_row: u32,
    ) {
        let s = footer_layout.start_row;
        let total_row = s - 3; // Footer starts 3 rows after Total

        // VLOOKUP labels
        self.add_vlookup(s, 4, 44); // E(s): "Balance for reporting period"
        self.add_vlookup(s + 1, 1, 43); // B(s+1): "CLOSING"
        self.add_vlookup(s + 2, 4, 10); // E(s+2): currency
        self.add_vlookup(s + 4, 1, 45); // B(s+4): "Balance..."
        self.add_vlookup(s + 6, 1, 46); // B(s+6): "Balance reconciliation:"
        self.add_vlookup(s + 7, 1, 47); // B(s+7): "Bank"
        self.add_vlookup(s + 8, 1, 48); // B(s+8): "Cash"
        self.add_vlookup(s + 9, 1, 49); // B(s+9): "Other"
        self.add_vlookup(s + 13, 1, 50); // B(s+13): confirmation text 1
        self.add_vlookup(s + 14, 1, 54); // B(s+14): confirmation text 2
        self.add_vlookup(s + 19, 1, 51); // B(s+19): "Place, date..."
        self.add_vlookup(s + 19, 3, 52); // D(s+19): "Signature..."
        self.add_vlookup(s + 20, 3, 53); // D(s+20): "Function..."

        // Check formulas
        // D(s+4): =IF(ROUND(E_balance,2)=(ROUND(F_income-F_total,2)),"✓","")
        let saldo_row_excel = s + 4 + 1;
        let income_row_excel = income_row + 1;
        let total_row_excel = total_row + 1;
        self.add_formula(
            s + 4,
            3,
            format!(
                r#"=IF(ROUND(E{},2)=(ROUND(F{}-F{},2)),"✓","")"#,
                saldo_row_excel, income_row_excel, total_row_excel
            ),
        );

        // E(s+4): difference = E_income - E_total
        self.add_formula(
            s + 4,
            4,
            format!("=E{}-E{}", income_row_excel, total_row_excel),
        );

        // E(s+6): =IF(E_balance=SUM(E_bank:E_other),"OK","")
        let bank_row_excel = footer_layout.input_rows[0] + 1;
        let sonstiges_row_excel = footer_layout.input_rows[2] + 1;
        self.add_formula(
            s + 6,
            4,
            format!(
                r#"=IF(E{}=SUM(E{}:E{}),"OK","")"#,
                saldo_row_excel, bank_row_excel, sonstiges_row_excel
            ),
        );
    }

    // ========================================================================
    // Input population
    // ========================================================================

    /// Populates all input values from [`ReportValues`] into the IronCalc model.
    pub(crate) fn populate(
        &mut self,
        values: &ReportValues,
        layout: &BodyLayout,
        footer_layout: &FooterLayout,
    ) {
        // Static API keys (header, table, panel)
        for key in ApiKey::all_static_keys() {
            let value = values.get(key);
            if let Some(addr) = key.static_addr() {
                self.set_cell_value(addr, value);
            }
        }

        // Dynamic position keys (body)
        for cat in &layout.categories {
            match &cat.mode {
                CategoryMode::HeaderInput { row } => {
                    for field in PositionField::header_input_fields() {
                        let key = ApiKey::Position {
                            category: cat.meta.num,
                            position: 0,
                            field,
                        };
                        let addr = CellAddr::new(*row, field.col());
                        self.set_cell_value(addr, values.get(key));
                    }
                }
                CategoryMode::WithPositions { positions, .. } => {
                    for i in 0..positions.count {
                        let row = positions.start_row + i as u32;
                        let pos = i + 1;
                        for field in PositionField::all() {
                            let key = ApiKey::Position {
                                category: cat.meta.num,
                                position: pos,
                                field,
                            };
                            let addr = CellAddr::new(row, field.col());
                            self.set_cell_value(addr, values.get(key));
                        }
                    }
                }
            }
        }

        // Footer keys (bank, cash, other)
        for field in FooterField::all() {
            let key = ApiKey::Footer(field);
            let addr = CellAddr::new(footer_layout.input_rows[field.index()], field.col());
            self.set_cell_value(addr, values.get(key));
        }
    }

    /// Evaluates all formulas in the model.
    pub(crate) fn evaluate(&mut self) {
        self.model.evaluate();
    }

    // ========================================================================
    // Result access
    // ========================================================================

    /// Reads a computed cell value (converts IronCalc → CellValue).
    pub(crate) fn get_value(&self, row: u32, col: u16) -> CellValue {
        let ic_val = self
            .model
            .get_cell_value_by_index(SHEET, row as i32 + 1, col as i32 + 1)
            .unwrap_or(IcCellValue::None);
        convert_cell_value(ic_val)
    }

    /// Reads a formula string (`None` if the cell has no formula).
    pub(crate) fn get_formula(&self, row: u32, col: u16) -> Option<String> {
        self.model
            .get_cell_formula(SHEET, row as i32 + 1, col as i32 + 1)
            .ok()
            .flatten()
    }

    /// Returns all cells containing formulas.
    pub(crate) fn formula_cells(&self) -> &[CellAddr] {
        &self.formula_cells
    }

    /// Returns all input values (address + original value).
    ///
    /// These are written directly to Excel without IronCalc readback.
    pub(crate) fn input_values(&self) -> &[(CellAddr, CellValue)] {
        &self.input_values
    }

    // ========================================================================
    // Private helpers
    // ========================================================================

    /// Sets a VLOOKUP formula and tracks the cell address.
    fn add_vlookup(&mut self, row: u32, col: u16, index: usize) {
        let formula = format!(
            r#"=IF($E$2="","",VLOOKUP($E$2,Sprachversionen!$B:$CD,{},FALSE))"#,
            index
        );
        self.add_formula(row, col, formula);
    }

    /// Sets a ratio formula: `=IFERROR(F{row}/D{row},0)`
    fn add_ratio(&mut self, row: u32) {
        let excel_row = row + 1;
        let formula = format!("=IFERROR(F{}/D{},0)", excel_row, excel_row);
        self.add_formula(row, 6, formula);
    }

    /// Sets a SUM range formula: `=SUM(D{start}:D{end})`
    fn add_sum_range(&mut self, target_row: u32, col: u16, start_row: u32, end_row: u32) {
        let col_letter = CellAddr::col_to_letter(col);
        let formula = format!(
            "=SUM({}{}:{}{})",
            col_letter,
            start_row + 1,
            col_letter,
            end_row + 1
        );
        self.add_formula(target_row, col, formula);
    }

    /// Sets a SUM formula over individual cells: `=SUM(D47+D69+...)`
    fn add_sum_cells(&mut self, target_row: u32, col: u16, source_rows: &[u32]) {
        let col_letter = CellAddr::col_to_letter(col);
        let refs: Vec<String> = source_rows
            .iter()
            .map(|r| format!("{}{}", col_letter, r + 1))
            .collect();
        let formula = format!("=SUM({})", refs.join("+"));
        self.add_formula(target_row, col, formula);
    }

    /// Low-level: sets a formula in IronCalc and tracks the cell address.
    fn add_formula(&mut self, row: u32, col: u16, formula: String) {
        self.model
            .update_cell_with_formula(SHEET, row as i32 + 1, col as i32 + 1, formula)
            .unwrap_or_else(|e| panic!("Failed to set formula at ({}, {}): {}", row, col, e));
        self.formula_cells.push(CellAddr::new(row, col));
    }

    /// Sets a [`CellValue`] in IronCalc and stores the original for the writer.
    ///
    /// Values are set in IronCalc so formulas can reference them, but the
    /// writer uses the stored originals (not IronCalc readback) to avoid
    /// auto-parsing side effects (e.g. date strings → numbers).
    fn set_cell_value(&mut self, addr: CellAddr, value: &CellValue) {
        let ic_row = addr.row as i32 + 1;
        let ic_col = addr.col as i32 + 1;

        match value {
            CellValue::Empty => {
                // Empty cells don't need to be set in IronCalc,
                // but we track them for the Excel output (blanks with format).
            }
            CellValue::Text(s) | CellValue::Date(s) => {
                self.model
                    .set_user_input(SHEET, ic_row, ic_col, s.clone())
                    .unwrap_or_else(|e| {
                        panic!("Failed to set text at ({}, {}): {}", addr.row, addr.col, e)
                    });
            }
            CellValue::Number(n) => {
                self.model
                    .set_user_input(SHEET, ic_row, ic_col, n.to_string())
                    .unwrap_or_else(|e| {
                        panic!("Failed to set number at ({}, {}): {}", addr.row, addr.col, e)
                    });
            }
        }

        self.input_values.push((addr, value.clone()));
    }
}

/// Converts IronCalc [`IcCellValue`] → project [`CellValue`].
fn convert_cell_value(ic: IcCellValue) -> CellValue {
    match ic {
        IcCellValue::None => CellValue::Empty,
        IcCellValue::String(s) => {
            if s.is_empty() {
                CellValue::Empty
            } else {
                CellValue::Text(s)
            }
        }
        IcCellValue::Number(n) => CellValue::Number(n),
        IcCellValue::Boolean(b) => CellValue::Text(if b { "TRUE" } else { "FALSE" }.to_string()),
    }
}

// ============================================================================
// Tests
// ============================================================================

#[cfg(test)]
mod tests {
    use super::*;
    use crate::report::body::config::BodyConfig;

    fn create_template() -> ModelTemplate {
        ModelTemplate::new()
    }

    #[test]
    fn test_bridge_basic_flow() {
        let template = create_template();
        let mut bridge = CalcBridge::from_template(&template);

        let config = BodyConfig::new()
            .with_positions(1, 2)
            .with_positions(2, 0)
            .with_positions(3, 0)
            .with_positions(4, 0)
            .with_positions(5, 0)
            .with_positions(6, 0)
            .with_positions(7, 0)
            .with_positions(8, 0);
        let layout = BodyLayout::compute(&config);
        let footer_layout = FooterLayout::compute(layout.total_row);

        bridge.register_body_formulas(&layout);
        bridge.register_footer_formulas(&footer_layout, 19);

        // Set values
        let values = ReportValues::new()
            .with_language("deutsch")
            .with_project_number("2025-001")
            .with_project_title("Testprojekt");

        bridge.populate(&values, &layout, &footer_layout);
        bridge.evaluate();

        // B1 should have the VLOOKUP title
        let b1 = bridge.get_value(0, 1);
        assert_eq!(
            b1,
            CellValue::Text("F I N A N Z B E R I C H T".to_string())
        );

        // D5 should contain the project number
        let d5 = bridge.get_value(4, 3);
        assert_eq!(d5, CellValue::Text("2025-001".to_string()));
    }

    #[test]
    fn test_bridge_body_formulas() {
        let template = create_template();
        let mut bridge = CalcBridge::from_template(&template);

        let config = BodyConfig::new()
            .with_positions(1, 3)
            .with_positions(2, 0)
            .with_positions(3, 0)
            .with_positions(4, 0)
            .with_positions(5, 0)
            .with_positions(6, 0)
            .with_positions(7, 0)
            .with_positions(8, 0);
        let layout = BodyLayout::compute(&config);
        let footer_layout = FooterLayout::compute(layout.total_row);

        bridge.register_body_formulas(&layout);
        bridge.register_footer_formulas(&footer_layout, 19);

        // Set positions for category 1
        let mut values = ReportValues::new().with_language("deutsch");
        values.set_position(1, 1, PositionField::Approved, 1000.0);
        values.set_position(1, 1, PositionField::IncomeTotal, 500.0);
        values.set_position(1, 2, PositionField::Approved, 2000.0);
        values.set_position(1, 2, PositionField::IncomeTotal, 1500.0);
        values.set_position(1, 3, PositionField::Approved, 3000.0);
        values.set_position(1, 3, PositionField::IncomeTotal, 2000.0);

        bridge.populate(&values, &layout, &footer_layout);
        bridge.evaluate();

        // Footer row for category 1: SUM(D27:D29) = 1000 + 2000 + 3000 = 6000
        let cat1_footer_row = layout.category(1).unwrap().footer_row().unwrap();
        let d_sum = bridge.get_value(cat1_footer_row, 3);
        assert_eq!(d_sum, CellValue::Number(6000.0));

        // F-Sum: 500 + 1500 + 2000 = 4000
        let f_sum = bridge.get_value(cat1_footer_row, 5);
        assert_eq!(f_sum, CellValue::Number(4000.0));

        // Ratio (G): IFERROR(F/D, 0) = 4000/6000 ≈ 0.6667
        let ratio = bridge.get_value(cat1_footer_row, 6);
        match ratio {
            CellValue::Number(n) => assert!((n - 4000.0 / 6000.0).abs() < 0.0001),
            other => panic!("Expected Number, got {:?}", other),
        }
    }

    #[test]
    fn test_bridge_formula_readback() {
        let template = create_template();
        let mut bridge = CalcBridge::from_template(&template);

        let config = BodyConfig::new().with_positions(1, 2);
        let layout = BodyLayout::compute(&config);

        bridge.register_body_formulas(&layout);

        // VLOOKUP formula for category 1 header (C27)
        let formula = bridge.get_formula(26, 2);
        assert!(formula.is_some(), "Should have formula at C27");
        let f = formula.unwrap();
        assert!(f.contains("VLOOKUP"), "Formula should contain VLOOKUP: {}", f);
    }

    #[test]
    fn test_bridge_footer_check() {
        let template = create_template();
        let mut bridge = CalcBridge::from_template(&template);

        let config = BodyConfig::new()
            .with_positions(1, 1)
            .with_positions(2, 0)
            .with_positions(3, 0)
            .with_positions(4, 0)
            .with_positions(5, 0)
            .with_positions(6, 0)
            .with_positions(7, 0)
            .with_positions(8, 0);
        let layout = BodyLayout::compute(&config);
        let footer_layout = FooterLayout::compute(layout.total_row);

        bridge.register_body_formulas(&layout);
        bridge.register_footer_formulas(&footer_layout, 19);

        // Income table: total income (F20)
        let mut values = ReportValues::new().with_language("deutsch");
        values.set(ApiKey::IncomeTotal(0), 1000.0); // F15

        // Body: category 1, position 1 expenses
        values.set_position(1, 1, PositionField::IncomeTotal, 800.0);

        // Footer: balance reconciliation
        values
            .set_footer_bank(200.0)
            .set_footer_kasse(0.0)
            .set_footer_sonstiges(0.0);

        bridge.populate(&values, &layout, &footer_layout);
        bridge.evaluate();

        // Difference E(s+4) = E_income(20) - E_total
        let diff = bridge.get_value(footer_layout.start_row + 4, 4);
        assert!(
            matches!(diff, CellValue::Number(_)),
            "Diff should be a number, got {:?}",
            diff
        );
    }

    #[test]
    fn test_bridge_cell_tracking() {
        let template = create_template();
        let mut bridge = CalcBridge::from_template(&template);

        let config = BodyConfig::new().with_positions(1, 1);
        let layout = BodyLayout::compute(&config);
        let footer_layout = FooterLayout::compute(layout.total_row);

        bridge.register_body_formulas(&layout);
        bridge.register_footer_formulas(&footer_layout, 19);

        assert!(
            !bridge.formula_cells().is_empty(),
            "Should have formula cells after registration"
        );

        let values = ReportValues::new().with_language("deutsch");
        bridge.populate(&values, &layout, &footer_layout);

        assert!(
            !bridge.input_values().is_empty(),
            "Should have input values after populate"
        );
    }
}
