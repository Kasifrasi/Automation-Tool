//! Input value store for the financial report
//!
//! All input values are stored in a single `HashMap<ApiKey, CellValue>`.
//! No redundant enums — `ApiKey` is the sole key type.

use super::keys::{ApiKey, FooterField, PositionField};
use super::types::{Category, Currency, Language, ReportDate};
use std::collections::HashMap;

// ============================================================================
// Cell Value Types
// ============================================================================

/// Possible cell values in a financial report
#[derive(Debug, Clone, PartialEq, Default)]
#[cfg_attr(feature = "serde", derive(serde::Serialize, serde::Deserialize))]
pub enum CellValue {
    #[default]
    Empty,
    Text(String),
    Number(f64),
    /// Date string (format: "YYYY-MM-DD" or "DD.MM.YYYY")
    Date(String),
}

impl CellValue {
    pub fn is_empty(&self) -> bool {
        matches!(self, CellValue::Empty)
    }

    pub fn as_text(&self) -> Option<&str> {
        match self {
            CellValue::Text(s) => Some(s),
            CellValue::Date(s) => Some(s), // dates are also strings
            _ => None,
        }
    }

    pub fn as_number(&self) -> Option<f64> {
        match self {
            CellValue::Number(n) => Some(*n),
            _ => None,
        }
    }
}

impl From<&str> for CellValue {
    fn from(s: &str) -> Self {
        if s.is_empty() {
            CellValue::Empty
        } else {
            CellValue::Text(s.to_string())
        }
    }
}

impl From<String> for CellValue {
    fn from(s: String) -> Self {
        if s.is_empty() {
            CellValue::Empty
        } else {
            CellValue::Text(s)
        }
    }
}

impl From<f64> for CellValue {
    fn from(n: f64) -> Self {
        CellValue::Number(n)
    }
}

// ============================================================================
// Report Values — all input values of the financial report
// ============================================================================

/// Stores all input values of a financial report
///
/// Uses `ApiKey` directly as key for type-safe access.
///
/// # Example
/// ```ignore
/// let mut values = ReportValues::new();
/// values.set(ApiKey::Language, "deutsch");
/// values.set(ApiKey::ApprovedBudget(0), 1000.0);
///
/// assert_eq!(values.get(ApiKey::Language).as_text(), Some("deutsch"));
/// ```
#[derive(Debug, Clone, Default)]
pub struct ReportValues {
    /// All cell values, indexed by `ApiKey`
    values: HashMap<ApiKey, CellValue>,
    /// Optional version/suffix text for cell B2
    version: Option<String>,
}

impl ReportValues {
    pub fn new() -> Self {
        Self::default()
    }

    /// Sets a value for an API cell
    pub fn set(&mut self, key: ApiKey, value: impl Into<CellValue>) -> &mut Self {
        self.values.insert(key, value.into());
        self
    }

    /// Returns a reference to the value for an API cell
    ///
    /// Returns `&CellValue::Empty` if not set.
    pub fn get(&self, key: ApiKey) -> &CellValue {
        self.values.get(&key).unwrap_or(&CellValue::Empty)
    }

    /// Returns an owned copy of the value for an API cell
    ///
    /// Returns `CellValue::Empty` if not set.
    pub fn get_owned(&self, key: ApiKey) -> CellValue {
        self.values.get(&key).cloned().unwrap_or(CellValue::Empty)
    }

    /// Checks whether a cell has a non-empty value
    pub fn has_value(&self, key: ApiKey) -> bool {
        self.values
            .get(&key)
            .map(|v| !v.is_empty())
            .unwrap_or(false)
    }

    // ========================================================================
    // Convenience getters
    // ========================================================================

    /// Returns the selected language (E2)
    pub fn language(&self) -> Option<&str> {
        self.get(ApiKey::Language).as_text()
    }

    /// Returns the selected currency (E3)
    pub fn currency(&self) -> Option<&str> {
        self.get(ApiKey::Currency).as_text()
    }

    /// Returns the project number (D5)
    pub fn project_number(&self) -> Option<&str> {
        self.get(ApiKey::ProjectNumber).as_text()
    }

    /// Returns the project title (D6)
    pub fn project_title(&self) -> Option<&str> {
        self.get(ApiKey::ProjectTitle).as_text()
    }

    /// Returns the version/suffix override (B2)
    pub fn version(&self) -> Option<&str> {
        self.version.as_deref()
    }

    // ========================================================================
    // String-based builder setters
    // ========================================================================

    /// Sets the language from a string (E2)
    pub fn with_language(mut self, lang: &str) -> Self {
        self.set(ApiKey::Language, lang);
        self
    }

    /// Sets the currency from a string (E3)
    pub fn with_currency(mut self, currency: &str) -> Self {
        self.set(ApiKey::Currency, currency);
        self
    }

    /// Sets the project number (D5)
    pub fn with_project_number(mut self, number: &str) -> Self {
        self.set(ApiKey::ProjectNumber, number);
        self
    }

    /// Sets the project title (D6)
    pub fn with_project_title(mut self, title: &str) -> Self {
        self.set(ApiKey::ProjectTitle, title);
        self
    }

    /// Sets the version text for cell B2
    pub fn with_version(mut self, version: &str) -> Self {
        self.version = if version.is_empty() { None } else { Some(version.to_string()) };
        self
    }

    /// Sets the project start date from a string (E8)
    pub fn with_project_start(mut self, date: &str) -> Self {
        self.set(ApiKey::ProjectStart, CellValue::Date(date.to_string()));
        self
    }

    /// Sets the project end date from a string (G8)
    pub fn with_project_end(mut self, date: &str) -> Self {
        self.set(ApiKey::ProjectEnd, CellValue::Date(date.to_string()));
        self
    }

    /// Sets the reporting period start from a string (E9)
    pub fn with_report_start(mut self, date: &str) -> Self {
        self.set(ApiKey::ReportStart, CellValue::Date(date.to_string()));
        self
    }

    /// Sets the reporting period end from a string (G9)
    pub fn with_report_end(mut self, date: &str) -> Self {
        self.set(ApiKey::ReportEnd, CellValue::Date(date.to_string()));
        self
    }

    // ========================================================================
    // Type-safe builder setters (recommended)
    // ========================================================================

    /// Sets the language using the [`Language`] enum (E2)
    pub fn with_lang(mut self, lang: Language) -> Self {
        self.set(ApiKey::Language, lang.as_str());
        self
    }

    /// Sets the currency using the [`Currency`] type (E3)
    pub fn with_curr(mut self, currency: Currency) -> Self {
        self.set(ApiKey::Currency, currency.as_str());
        self
    }

    /// Sets the project start date using a validated [`ReportDate`] (E8)
    pub fn with_project_start_date(mut self, date: ReportDate) -> Self {
        self.set(ApiKey::ProjectStart, CellValue::Date(date.format_de()));
        self
    }

    /// Sets the project end date using a validated [`ReportDate`] (G8)
    pub fn with_project_end_date(mut self, date: ReportDate) -> Self {
        self.set(ApiKey::ProjectEnd, CellValue::Date(date.format_de()));
        self
    }

    /// Sets the reporting period start using a validated [`ReportDate`] (E9)
    pub fn with_report_start_date(mut self, date: ReportDate) -> Self {
        self.set(ApiKey::ReportStart, CellValue::Date(date.format_de()));
        self
    }

    /// Sets the reporting period end using a validated [`ReportDate`] (G9)
    pub fn with_report_end_date(mut self, date: ReportDate) -> Self {
        self.set(ApiKey::ReportEnd, CellValue::Date(date.format_de()));
        self
    }

    // ========================================================================
    // Position methods (dynamic cost positions)
    // ========================================================================

    /// Sets a single position field
    ///
    /// # Arguments
    /// * `category` — category number (1–8)
    /// * `position` — `0` for header-input mode, `1..N` for individual rows
    /// * `field` — which column of the position
    /// * `value` — the cell value
    ///
    /// **Note:** At `position=0` (header-input), `PositionField::Description` is
    /// not available because column C holds the VLOOKUP label.
    pub fn set_position(
        &mut self,
        category: u8,
        position: u16,
        field: PositionField,
        value: impl Into<CellValue>,
    ) -> &mut Self {
        self.set(
            ApiKey::Position {
                category,
                position,
                field,
            },
            value,
        )
    }

    /// Sets all 5 fields of a position row at once
    ///
    /// Only for `position >= 1`. For header-input (`position=0`), use
    /// [`set_header_input()`](Self::set_header_input).
    #[allow(clippy::too_many_arguments)]
    pub fn set_position_row(
        &mut self,
        category: u8,
        position: u16,
        description: impl Into<CellValue>,
        approved: impl Into<CellValue>,
        income_report: impl Into<CellValue>,
        income_total: impl Into<CellValue>,
        remark: impl Into<CellValue>,
    ) -> &mut Self {
        use PositionField::*;
        self.set_position(category, position, Description, description);
        self.set_position(category, position, Approved, approved);
        self.set_position(category, position, IncomeReport, income_report);
        self.set_position(category, position, IncomeTotal, income_total);
        self.set_position(category, position, Remark, remark);
        self
    }

    /// Sets header-input values (`position=0`) for categories with 0 position rows
    ///
    /// No `Description` field — column C holds the VLOOKUP label.
    pub fn set_header_input(
        &mut self,
        category: u8,
        approved: impl Into<CellValue>,
        income_report: impl Into<CellValue>,
        income_total: impl Into<CellValue>,
        remark: impl Into<CellValue>,
    ) -> &mut Self {
        use PositionField::*;
        // position=0 for header-input mode
        self.set_position(category, 0, Approved, approved);
        self.set_position(category, 0, IncomeReport, income_report);
        self.set_position(category, 0, IncomeTotal, income_total);
        self.set_position(category, 0, Remark, remark);
        self
    }

    /// Returns the value of a position field
    pub fn get_position(&self, category: u8, position: u16, field: PositionField) -> &CellValue {
        self.get(ApiKey::Position {
            category,
            position,
            field,
        })
    }

    // ========================================================================
    // Type-safe position methods (using Category enum)
    // ========================================================================

    /// Sets a single position field using the [`Category`] enum
    pub fn set_cat_position(
        &mut self,
        category: Category,
        position: u16,
        field: PositionField,
        value: impl Into<CellValue>,
    ) -> &mut Self {
        self.set_position(category.index(), position, field, value)
    }

    /// Sets all 5 fields of a position row using the [`Category`] enum
    #[allow(clippy::too_many_arguments)]
    pub fn set_cat_position_row(
        &mut self,
        category: Category,
        position: u16,
        description: impl Into<CellValue>,
        approved: impl Into<CellValue>,
        income_report: impl Into<CellValue>,
        income_total: impl Into<CellValue>,
        remark: impl Into<CellValue>,
    ) -> &mut Self {
        self.set_position_row(
            category.index(),
            position,
            description,
            approved,
            income_report,
            income_total,
            remark,
        )
    }

    /// Sets header-input values using the [`Category`] enum
    pub fn set_cat_header_input(
        &mut self,
        category: Category,
        approved: impl Into<CellValue>,
        income_report: impl Into<CellValue>,
        income_total: impl Into<CellValue>,
        remark: impl Into<CellValue>,
    ) -> &mut Self {
        self.set_header_input(
            category.index(),
            approved,
            income_report,
            income_total,
            remark,
        )
    }

    // ========================================================================
    // Footer values (balance reconciliation)
    // ========================================================================

    /// Sets the bank balance (consuming builder)
    pub fn with_footer_bank(mut self, value: f64) -> Self {
        self.set(ApiKey::Footer(FooterField::Bank), value);
        self
    }

    /// Sets the cash balance (consuming builder)
    pub fn with_footer_kasse(mut self, value: f64) -> Self {
        self.set(ApiKey::Footer(FooterField::Kasse), value);
        self
    }

    /// Sets other balances (uncashed cheques, advances, loans, etc.)
    pub fn with_footer_sonstiges(mut self, value: f64) -> Self {
        self.set(ApiKey::Footer(FooterField::Sonstiges), value);
        self
    }

    /// Sets all three footer balances at once
    pub fn with_footer_salden(mut self, bank: f64, kasse: f64, sonstiges: f64) -> Self {
        self.set(ApiKey::Footer(FooterField::Bank), bank);
        self.set(ApiKey::Footer(FooterField::Kasse), kasse);
        self.set(ApiKey::Footer(FooterField::Sonstiges), sonstiges);
        self
    }

    /// Sets the bank balance (mutating)
    pub fn set_footer_bank(&mut self, value: f64) -> &mut Self {
        self.set(ApiKey::Footer(FooterField::Bank), value);
        self
    }

    /// Sets the cash balance (mutating)
    pub fn set_footer_kasse(&mut self, value: f64) -> &mut Self {
        self.set(ApiKey::Footer(FooterField::Kasse), value);
        self
    }

    /// Sets other balances (mutating)
    pub fn set_footer_sonstiges(&mut self, value: f64) -> &mut Self {
        self.set(ApiKey::Footer(FooterField::Sonstiges), value);
        self
    }

    /// Returns the bank balance
    pub fn footer_bank(&self) -> Option<f64> {
        self.get(ApiKey::Footer(FooterField::Bank)).as_number()
    }

    /// Returns the cash balance
    pub fn footer_kasse(&self) -> Option<f64> {
        self.get(ApiKey::Footer(FooterField::Kasse)).as_number()
    }

    /// Returns other balances
    pub fn footer_sonstiges(&self) -> Option<f64> {
        self.get(ApiKey::Footer(FooterField::Sonstiges)).as_number()
    }

    // ========================================================================
    // Validation (balance reconciliation)
    // ========================================================================

    /// Computes the sum of all footer balances
    pub fn footer_balance_total(&self) -> f64 {
        self.footer_bank().unwrap_or(0.0)
            + self.footer_kasse().unwrap_or(0.0)
            + self.footer_sonstiges().unwrap_or(0.0)
    }

    /// Returns `true` if at least bank and cash balances are set
    pub fn validate_footer_complete(&self) -> bool {
        self.footer_bank().is_some() && self.footer_kasse().is_some()
    }
}

// ============================================================================
// Tests
// ============================================================================

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_report_values_builder() {
        let values = ReportValues::new()
            .with_language("deutsch")
            .with_currency("EUR")
            .with_project_number("12345");

        assert_eq!(values.language(), Some("deutsch"));
        assert_eq!(values.currency(), Some("EUR"));
        assert_eq!(values.project_number(), Some("12345"));
    }

    #[test]
    fn test_generic_set_get() {
        let mut values = ReportValues::new();
        values.set(ApiKey::Language, "english");
        values.set(ApiKey::ApprovedBudget(0), 1000.0);

        assert_eq!(values.get(ApiKey::Language).as_text(), Some("english"));
        assert_eq!(
            values.get(ApiKey::ApprovedBudget(0)).as_number(),
            Some(1000.0)
        );
    }

    #[test]
    fn test_range_cells() {
        let mut values = ReportValues::new();

        // Set all 5 ApprovedBudget values
        for i in 0..5u8 {
            values.set(ApiKey::ApprovedBudget(i), 1000.0 * (i + 1) as f64);
        }

        assert_eq!(
            values.get(ApiKey::ApprovedBudget(0)).as_number(),
            Some(1000.0)
        );
        assert_eq!(
            values.get(ApiKey::ApprovedBudget(4)).as_number(),
            Some(5000.0)
        );
    }

    #[test]
    fn test_right_panel_cells() {
        let mut values = ReportValues::new();

        values.set(
            ApiKey::LeftDate(0),
            CellValue::Date("2024-01-15".to_string()),
        );
        values.set(ApiKey::LeftAmountEuro(0), 500.0);
        values.set(ApiKey::RightAmountLocal(17), 999.99);

        assert_eq!(
            values.get(ApiKey::LeftDate(0)).as_text(),
            Some("2024-01-15")
        );
        assert_eq!(
            values.get(ApiKey::LeftAmountEuro(0)).as_number(),
            Some(500.0)
        );
        assert_eq!(
            values.get(ApiKey::RightAmountLocal(17)).as_number(),
            Some(999.99)
        );
    }

    #[test]
    fn test_set_position() {
        use PositionField::*;

        let mut values = ReportValues::new();
        values.set_position(1, 1, Description, "Personalkosten");
        values.set_position(1, 1, Approved, 5000.0);

        assert_eq!(
            values.get_position(1, 1, Description).as_text(),
            Some("Personalkosten")
        );
        assert_eq!(
            values.get_position(1, 1, Approved).as_number(),
            Some(5000.0)
        );
    }

    #[test]
    fn test_set_position_row() {
        use PositionField::*;

        let mut values = ReportValues::new();
        values.set_position_row(
            1,                    // category
            1,                    // position
            "Reisekosten",        // Description
            2000.0,               // Approved
            1800.0,               // IncomeReport
            1800.0,               // IncomeTotal
            "Dienstreise Berlin", // Remark
        );

        assert_eq!(
            values.get_position(1, 1, Description).as_text(),
            Some("Reisekosten")
        );
        assert_eq!(
            values.get_position(1, 1, Approved).as_number(),
            Some(2000.0)
        );
        assert_eq!(
            values.get_position(1, 1, IncomeReport).as_number(),
            Some(1800.0)
        );
        assert_eq!(
            values.get_position(1, 1, IncomeTotal).as_number(),
            Some(1800.0)
        );
        assert_eq!(
            values.get_position(1, 1, Remark).as_text(),
            Some("Dienstreise Berlin")
        );
    }

    #[test]
    fn test_header_input() {
        use PositionField::*;

        let mut values = ReportValues::new();
        values.set_header_input(6, 4000.0, 2000.0, 2000.0, "Sonstiges");

        // position=0 for header-input mode
        assert_eq!(
            values.get_position(6, 0, Approved).as_number(),
            Some(4000.0)
        );
        assert_eq!(
            values.get_position(6, 0, IncomeReport).as_number(),
            Some(2000.0)
        );
        assert_eq!(
            values.get_position(6, 0, IncomeTotal).as_number(),
            Some(2000.0)
        );
        assert_eq!(
            values.get_position(6, 0, Remark).as_text(),
            Some("Sonstiges")
        );

        // Description not available in header-input mode
        assert!(values.get_position(6, 0, Description).is_empty());
    }

    #[test]
    fn test_multiple_positions() {
        use PositionField::*;

        let mut values = ReportValues::new();

        // Category 1, positions 1-3
        values.set_position_row(1, 1, "Personal", 5000.0, 2500.0, 2500.0, "");
        values.set_position_row(1, 2, "Reisen", 2000.0, 1800.0, 1800.0, "");
        values.set_position_row(1, 3, "Material", 1000.0, 500.0, 500.0, "");

        // Category 2, position 1
        values.set_position_row(2, 1, "Externe", 3000.0, 1500.0, 1500.0, "");

        // Verify different positions
        assert_eq!(
            values.get_position(1, 1, Description).as_text(),
            Some("Personal")
        );
        assert_eq!(
            values.get_position(1, 3, Approved).as_number(),
            Some(1000.0)
        );
        assert_eq!(
            values.get_position(2, 1, Description).as_text(),
            Some("Externe")
        );

        // Unset position returns Empty
        assert!(values.get_position(1, 4, Description).is_empty());
    }

    #[test]
    fn test_mixed_modes() {
        use PositionField::*;

        let mut values = ReportValues::new();

        // Category 1: position mode (position >= 1)
        values.set_position_row(1, 1, "Personal", 5000.0, 2500.0, 2500.0, "");

        // Category 6: header-input mode (position = 0)
        values.set_header_input(6, 4000.0, 2000.0, 2000.0, "");

        // Verify both modes
        assert_eq!(
            values.get_position(1, 1, Description).as_text(),
            Some("Personal")
        );
        assert_eq!(
            values.get_position(6, 0, Approved).as_number(),
            Some(4000.0)
        );
    }

    #[test]
    fn test_footer_values_in_hashmap() {
        let values = ReportValues::new()
            .with_footer_bank(1500.50)
            .with_footer_kasse(250.25)
            .with_footer_sonstiges(100.0);

        // Check footer values via convenience methods
        assert_eq!(values.footer_bank(), Some(1500.50));
        assert_eq!(values.footer_kasse(), Some(250.25));
        assert_eq!(values.footer_sonstiges(), Some(100.0));

        // Verify footer values are in the HashMap (uniform access)
        assert_eq!(
            values.get(ApiKey::Footer(FooterField::Bank)).as_number(),
            Some(1500.50)
        );
        assert_eq!(
            values.get(ApiKey::Footer(FooterField::Kasse)).as_number(),
            Some(250.25)
        );
        assert_eq!(
            values
                .get(ApiKey::Footer(FooterField::Sonstiges))
                .as_number(),
            Some(100.0)
        );

        // Verify balance reconciliation logic
        assert_eq!(values.footer_balance_total(), 1850.75);
        assert!(values.validate_footer_complete());
    }

    #[test]
    fn test_footer_mutating_setters() {
        let mut values = ReportValues::new();

        values
            .set_footer_bank(2000.0)
            .set_footer_kasse(500.0)
            .set_footer_sonstiges(50.0);

        assert_eq!(values.footer_balance_total(), 2550.0);
    }

    #[test]
    fn test_footer_validation_incomplete() {
        let values = ReportValues::new().with_footer_bank(1000.0);

        // Only bank set, cash missing
        assert!(!values.validate_footer_complete());

        let values = values.with_footer_kasse(500.0);
        assert!(values.validate_footer_complete());
    }

    #[test]
    fn test_unified_access_all_types() {
        // All value types (header, positions, footer) are accessed
        // through the same unified HashMap mechanism.

        let mut values = ReportValues::new()
            .with_language("deutsch")
            .with_footer_bank(1000.0);

        values.set_position(1, 1, PositionField::Approved, 5000.0);

        // Uniform access via get()
        assert!(values.get(ApiKey::Language).as_text().is_some());
        assert!(values
            .get(ApiKey::Footer(FooterField::Bank))
            .as_number()
            .is_some());
        assert!(values
            .get(ApiKey::Position {
                category: 1,
                position: 1,
                field: PositionField::Approved
            })
            .as_number()
            .is_some());
    }
}
