//! Sheet Protection and Data Validation
//!
//! Provides APIs for:
//! - Sheet protection settings (lock/unlock cells, allow specific actions)
//! - Data validation rules for input cells (numbers, dates, text, lists)
//!
//! ## Usage
//!
//! ```ignore
//! use fb_generator::report::options::{SheetProtection, ValidationRule, FieldValidation};
//! use fb_generator::report::api::ApiKey;
//!
//! // Protection with custom options
//! let protection = SheetProtection::new()
//!     .with_password("secret")
//!     .allow_format_cells(true)
//!     .allow_sort(true);
//!
//! // Validation for specific fields
//! let validation = FieldValidation::new()
//!     .add_rule(
//!         ValidationTarget::Field(ApiKey::ApprovedBudget(0)),
//!         ValidationRule::decimal_between(0.0, 1_000_000.0)
//!             .with_error_title("Invalid Amount")
//!             .with_error_message("Please enter a value between 0 and 1,000,000"),
//!     );
//! ```

use crate::lang::LANG_CONFIG;
use crate::report::api::{ApiKey, FooterField, PositionField};
use rust_xlsxwriter::{
    DataValidation, DataValidationErrorStyle, DataValidationRule, Formula, IntoDataValidationValue,
};
use std::collections::HashSet;

// ============================================================================
// Sheet Protection
// ============================================================================

/// Sheet protection settings
///
/// Maps directly to rust_xlsxwriter's ProtectionOptions with a builder API.
/// By default, all editing is locked; use `allow_*` methods to enable specific actions.
#[derive(Debug, Clone, PartialEq)]
#[cfg_attr(feature = "serde", derive(serde::Serialize, serde::Deserialize))]
#[cfg_attr(feature = "serde", serde(default))]
pub struct SheetProtection {
    /// Optional password (weak protection - easily bypassed)
    pub password: Option<String>,

    // --- Protection Flags (false = protected, true = allowed) ---
    /// Allow selecting locked cells (default: true)
    pub select_locked_cells: bool,
    /// Allow selecting unlocked cells (default: true)
    pub select_unlocked_cells: bool,
    /// Allow formatting cells (default: false)
    pub format_cells: bool,
    /// Allow formatting columns (default: false)
    pub format_columns: bool,
    /// Allow formatting rows (default: false)
    pub format_rows: bool,
    /// Allow inserting columns (default: false)
    pub insert_columns: bool,
    /// Allow inserting rows (default: false)
    pub insert_rows: bool,
    /// Allow inserting hyperlinks (default: false)
    pub insert_hyperlinks: bool,
    /// Allow deleting columns (default: false)
    pub delete_columns: bool,
    /// Allow deleting rows (default: false)
    pub delete_rows: bool,
    /// Allow sorting (default: false)
    pub sort: bool,
    /// Allow autofilter (default: false)
    pub autofilter: bool,
    /// Allow pivot tables (default: false)
    pub pivot_tables: bool,
    /// Allow editing objects (default: false)
    pub edit_objects: bool,
    /// Allow editing scenarios (default: false)
    pub edit_scenarios: bool,
    /// Protect contents of locked cells (default: false)
    pub contents: bool,
}

impl Default for SheetProtection {
    fn default() -> Self {
        Self {
            password: None,
            select_locked_cells: true,
            select_unlocked_cells: true,
            format_cells: false,
            format_columns: false,
            format_rows: false,
            insert_columns: false,
            insert_rows: false,
            insert_hyperlinks: false,
            delete_columns: false,
            delete_rows: false,
            sort: false,
            autofilter: false,
            pivot_tables: false,
            edit_objects: false,
            edit_scenarios: false,
            contents: false,
        }
    }
}

impl SheetProtection {
    /// Creates new protection with default settings (all locked except selection)
    pub fn new() -> Self {
        Self::default()
    }

    /// Creates protection with recommended default settings
    pub fn from_defaults() -> Self {
        Self {
            password: None,
            select_locked_cells: true,
            select_unlocked_cells: true,
            format_cells: true,
            format_columns: true,
            format_rows: true,
            insert_columns: false,
            insert_rows: false,
            insert_hyperlinks: true,
            delete_columns: false,
            delete_rows: false,
            sort: true,
            autofilter: true,
            pivot_tables: true,
            edit_objects: false,
            edit_scenarios: true,
            contents: false,
        }
    }

    /// Sets a password (weak protection)
    pub fn with_password(mut self, password: impl Into<String>) -> Self {
        self.password = Some(password.into());
        self
    }

    // --- Builder methods ---

    pub fn allow_select_locked_cells(mut self, allow: bool) -> Self {
        self.select_locked_cells = allow;
        self
    }

    pub fn allow_select_unlocked_cells(mut self, allow: bool) -> Self {
        self.select_unlocked_cells = allow;
        self
    }

    pub fn allow_format_cells(mut self, allow: bool) -> Self {
        self.format_cells = allow;
        self
    }

    pub fn allow_format_columns(mut self, allow: bool) -> Self {
        self.format_columns = allow;
        self
    }

    pub fn allow_format_rows(mut self, allow: bool) -> Self {
        self.format_rows = allow;
        self
    }

    pub fn allow_insert_columns(mut self, allow: bool) -> Self {
        self.insert_columns = allow;
        self
    }

    pub fn allow_insert_rows(mut self, allow: bool) -> Self {
        self.insert_rows = allow;
        self
    }

    pub fn allow_insert_hyperlinks(mut self, allow: bool) -> Self {
        self.insert_hyperlinks = allow;
        self
    }

    pub fn allow_delete_columns(mut self, allow: bool) -> Self {
        self.delete_columns = allow;
        self
    }

    pub fn allow_delete_rows(mut self, allow: bool) -> Self {
        self.delete_rows = allow;
        self
    }

    pub fn allow_sort(mut self, allow: bool) -> Self {
        self.sort = allow;
        self
    }

    pub fn allow_autofilter(mut self, allow: bool) -> Self {
        self.autofilter = allow;
        self
    }

    pub fn allow_pivot_tables(mut self, allow: bool) -> Self {
        self.pivot_tables = allow;
        self
    }

    pub fn allow_edit_objects(mut self, allow: bool) -> Self {
        self.edit_objects = allow;
        self
    }

    pub fn allow_edit_scenarios(mut self, allow: bool) -> Self {
        self.edit_scenarios = allow;
        self
    }

    pub fn allow_contents(mut self, allow: bool) -> Self {
        self.contents = allow;
        self
    }

    /// Converts to rust_xlsxwriter ProtectionOptions
    pub fn to_protection_options(&self) -> rust_xlsxwriter::ProtectionOptions {
        rust_xlsxwriter::ProtectionOptions {
            select_locked_cells: self.select_locked_cells,
            select_unlocked_cells: self.select_unlocked_cells,
            format_cells: self.format_cells,
            format_columns: self.format_columns,
            format_rows: self.format_rows,
            insert_columns: self.insert_columns,
            insert_rows: self.insert_rows,
            insert_links: self.insert_hyperlinks,
            delete_columns: self.delete_columns,
            delete_rows: self.delete_rows,
            sort: self.sort,
            use_autofilter: self.autofilter,
            use_pivot_tables: self.pivot_tables,
            edit_objects: self.edit_objects,
            edit_scenarios: self.edit_scenarios,
            contents: self.contents,
        }
    }
}

// ============================================================================
// Validation Target
// ============================================================================

/// Target for validation rules
///
/// Specifies which cells should receive a validation rule.
/// The system prevents overlapping validations (same cell, multiple rules).
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub enum ValidationTarget {
    /// Single API field (static address)
    Field(ApiKey),

    /// Range of API fields (e.g., all ApprovedBudget fields)
    FieldRange {
        /// Start key (e.g., ApprovedBudget(0))
        start: ApiKey,
        /// End key (e.g., ApprovedBudget(4))
        end: ApiKey,
    },

    /// All fields of a specific type in body positions
    PositionFieldAll {
        /// Which field (e.g., Approved, IncomeReport)
        field: PositionField,
    },

    /// Specific position field
    PositionFieldSingle {
        category: u8,
        position: u16,
        field: PositionField,
    },

    /// All footer fields
    FooterAll,

    /// Specific footer field
    FooterSingle(FooterField),

    /// Right panel: all date fields (L14:L31, S14:S31)
    RightPanelDates,

    /// Right panel: all amount fields (M14:N31, T14:U31)
    RightPanelAmounts,

    /// Custom cell range (e.g., "D15:D19")
    CustomRange(String),
}

// ============================================================================
// Validation Rule
// ============================================================================

/// Data validation rule with optional messages
///
/// Wraps rust_xlsxwriter's DataValidation with a cleaner builder API
/// and language-aware error messages.
#[derive(Debug, Clone)]
pub struct ValidationRule {
    /// The validation type
    pub rule_type: ValidationRuleType,

    /// Input help title (max 32 chars)
    pub input_title: Option<String>,
    /// Input help message (max 255 chars)
    pub input_message: Option<String>,

    /// Error dialog title (max 32 chars)
    pub error_title: Option<String>,
    /// Error dialog message (max 255 chars)
    pub error_message: Option<String>,
    /// Error style: Stop (default), Warning, or Information
    pub error_style: ValidationErrorStyle,

    /// Allow blank/empty cells (default: true)
    pub ignore_blank: bool,
    /// Show dropdown for list validations (default: true)
    pub show_dropdown: bool,
}

/// Type of validation rule
#[derive(Debug, Clone)]
pub enum ValidationRuleType {
    /// Whole number (integer) validation
    WholeNumber(NumericRule<i32>),

    /// Decimal number validation
    Decimal(NumericRule<f64>),

    /// Date validation
    Date(DateRule),

    /// Text length validation
    TextLength(NumericRule<u32>),

    /// List validation (dropdown)
    List(Vec<String>),

    /// List validation from cell range
    ListFormula(String),

    /// Custom formula validation
    Custom(String),

    /// Any value (removes restrictions but shows messages)
    Any,
}

/// Numeric comparison rule
#[derive(Debug, Clone)]
pub enum NumericRule<T> {
    Between(T, T),
    NotBetween(T, T),
    EqualTo(T),
    NotEqualTo(T),
    GreaterThan(T),
    GreaterThanOrEqualTo(T),
    LessThan(T),
    LessThanOrEqualTo(T),
}

/// Date comparison rule
#[derive(Debug, Clone)]
pub enum DateRule {
    Between(String, String),
    NotBetween(String, String),
    EqualTo(String),
    NotEqualTo(String),
    GreaterThan(String),
    GreaterThanOrEqualTo(String),
    LessThan(String),
    LessThanOrEqualTo(String),
}

/// Error dialog style
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub enum ValidationErrorStyle {
    /// Prevents entry (default)
    #[default]
    Stop,
    /// Warns but allows entry
    Warning,
    /// Information only
    Information,
}

impl ValidationRule {
    /// Creates a new validation rule
    pub fn new(rule_type: ValidationRuleType) -> Self {
        Self {
            rule_type,
            input_title: None,
            input_message: None,
            error_title: None,
            error_message: None,
            error_style: ValidationErrorStyle::Stop,
            ignore_blank: true,
            show_dropdown: true,
        }
    }

    // --- Convenience constructors ---

    /// Decimal validation: value must be between min and max
    pub fn decimal_between(min: f64, max: f64) -> Self {
        Self::new(ValidationRuleType::Decimal(NumericRule::Between(min, max)))
    }

    /// Decimal validation: value must be >= 0
    pub fn decimal_non_negative() -> Self {
        Self::new(ValidationRuleType::Decimal(
            NumericRule::GreaterThanOrEqualTo(0.0),
        ))
    }

    /// Decimal validation: value must be > 0
    pub fn decimal_positive() -> Self {
        Self::new(ValidationRuleType::Decimal(NumericRule::GreaterThan(0.0)))
    }

    /// Whole number validation: value must be between min and max
    pub fn whole_number_between(min: i32, max: i32) -> Self {
        Self::new(ValidationRuleType::WholeNumber(NumericRule::Between(
            min, max,
        )))
    }

    /// Text length validation: max characters
    pub fn text_max_length(max: u32) -> Self {
        Self::new(ValidationRuleType::TextLength(
            NumericRule::LessThanOrEqualTo(max),
        ))
    }

    /// Text length validation: between min and max characters
    pub fn text_length_between(min: u32, max: u32) -> Self {
        Self::new(ValidationRuleType::TextLength(NumericRule::Between(
            min, max,
        )))
    }

    /// Date validation: between two dates (format: YYYY-MM-DD)
    pub fn date_between(start: impl Into<String>, end: impl Into<String>) -> Self {
        Self::new(ValidationRuleType::Date(DateRule::Between(
            start.into(),
            end.into(),
        )))
    }

    /// List validation: dropdown with options
    pub fn list(options: Vec<String>) -> Self {
        Self::new(ValidationRuleType::List(options))
    }

    /// List validation: dropdown from cell range
    pub fn list_from_range(range: impl Into<String>) -> Self {
        Self::new(ValidationRuleType::ListFormula(range.into()))
    }

    /// Custom formula validation
    pub fn custom(formula: impl Into<String>) -> Self {
        Self::new(ValidationRuleType::Custom(formula.into()))
    }

    /// Decimal validation with max 2 decimal places (common for currency)
    ///
    /// Uses custom formula: =AND(D1>=0,D1=ROUND(D1,2))
    pub fn currency() -> Self {
        Self::new(ValidationRuleType::Custom(
            "=AND(INDIRECT(ADDRESS(ROW(),COLUMN()))>=0,INDIRECT(ADDRESS(ROW(),COLUMN()))=ROUND(INDIRECT(ADDRESS(ROW(),COLUMN())),2))".to_string()
        ))
    }

    // --- Builder methods ---

    pub fn with_input_title(mut self, title: impl Into<String>) -> Self {
        self.input_title = Some(title.into());
        self
    }

    pub fn with_input_message(mut self, message: impl Into<String>) -> Self {
        self.input_message = Some(message.into());
        self
    }

    pub fn with_error_title(mut self, title: impl Into<String>) -> Self {
        self.error_title = Some(title.into());
        self
    }

    pub fn with_error_message(mut self, message: impl Into<String>) -> Self {
        self.error_message = Some(message.into());
        self
    }

    pub fn with_error_style(mut self, style: ValidationErrorStyle) -> Self {
        self.error_style = style;
        self
    }

    pub fn with_ignore_blank(mut self, ignore: bool) -> Self {
        self.ignore_blank = ignore;
        self
    }

    pub fn with_show_dropdown(mut self, show: bool) -> Self {
        self.show_dropdown = show;
        self
    }

    /// Sets language-specific error message for currency validation
    pub fn with_currency_error(mut self, language: Option<&str>) -> Self {
        let lang_key = language.unwrap_or("Deutsch");
        if let Some(config) = LANG_CONFIG.get(lang_key) {
            self.error_message = Some(config.decimal_err.to_string());
        }
        self
    }

    /// Converts to rust_xlsxwriter DataValidation
    pub fn to_data_validation(&self) -> Result<DataValidation, &'static str> {
        let mut dv = DataValidation::new();

        // Apply rule type
        match &self.rule_type {
            ValidationRuleType::WholeNumber(rule) => {
                dv = dv.allow_whole_number(rule.to_xlsx_rule());
            }
            ValidationRuleType::Decimal(rule) => {
                dv = dv.allow_decimal_number(rule.to_xlsx_rule());
            }
            ValidationRuleType::Date(_rule) => {
                // Note: Would need ExcelDateTime conversion
                return Err("Date validation requires ExcelDateTime conversion");
            }
            ValidationRuleType::TextLength(rule) => {
                dv = dv.allow_text_length(rule.to_xlsx_rule());
            }
            ValidationRuleType::List(options) => {
                let refs: Vec<&str> = options.iter().map(|s| s.as_str()).collect();
                dv = dv
                    .allow_list_strings(&refs)
                    .map_err(|_| "List too long (max 255 chars)")?;
            }
            ValidationRuleType::ListFormula(range) => {
                dv = dv.allow_list_formula(Formula::new(range));
            }
            ValidationRuleType::Custom(formula) => {
                dv = dv.allow_custom(Formula::new(formula));
            }
            ValidationRuleType::Any => {
                dv = dv.allow_any_value();
            }
        }

        // Apply messages
        if let Some(ref title) = self.input_title {
            dv = dv
                .set_input_title(title)
                .map_err(|_| "Input title too long")?;
        }
        if let Some(ref msg) = self.input_message {
            dv = dv
                .set_input_message(msg)
                .map_err(|_| "Input message too long")?;
        }
        if let Some(ref title) = self.error_title {
            dv = dv
                .set_error_title(title)
                .map_err(|_| "Error title too long")?;
        }
        if let Some(ref msg) = self.error_message {
            dv = dv
                .set_error_message(msg)
                .map_err(|_| "Error message too long")?;
        }

        // Apply error style
        dv = dv.set_error_style(match self.error_style {
            ValidationErrorStyle::Stop => DataValidationErrorStyle::Stop,
            ValidationErrorStyle::Warning => DataValidationErrorStyle::Warning,
            ValidationErrorStyle::Information => DataValidationErrorStyle::Information,
        });

        // Apply options
        dv = dv.ignore_blank(self.ignore_blank);
        dv = dv.show_dropdown(self.show_dropdown);

        Ok(dv)
    }
}

impl<T: Copy + IntoDataValidationValue> NumericRule<T> {
    /// Converts to rust_xlsxwriter DataValidationRule
    pub fn to_xlsx_rule(&self) -> DataValidationRule<T> {
        match *self {
            NumericRule::Between(a, b) => DataValidationRule::Between(a, b),
            NumericRule::NotBetween(a, b) => DataValidationRule::NotBetween(a, b),
            NumericRule::EqualTo(v) => DataValidationRule::EqualTo(v),
            NumericRule::NotEqualTo(v) => DataValidationRule::NotEqualTo(v),
            NumericRule::GreaterThan(v) => DataValidationRule::GreaterThan(v),
            NumericRule::GreaterThanOrEqualTo(v) => DataValidationRule::GreaterThanOrEqualTo(v),
            NumericRule::LessThan(v) => DataValidationRule::LessThan(v),
            NumericRule::LessThanOrEqualTo(v) => DataValidationRule::LessThanOrEqualTo(v),
        }
    }
}

// ============================================================================
// Field Validation Container
// ============================================================================

/// Container for field-based validations
///
/// Collects validation rules for specific API fields and ensures
/// no overlapping rules (same cell with multiple validations).
#[derive(Debug, Clone, Default)]
pub struct FieldValidation {
    /// Rules by target
    rules: Vec<(ValidationTarget, ValidationRule)>,
    /// Track assigned cells to prevent overlaps (used during validation application)
    #[allow(dead_code)]
    assigned_cells: HashSet<String>,
}

impl FieldValidation {
    pub fn new() -> Self {
        Self::default()
    }

    /// Adds a validation rule for a target
    ///
    /// Returns error if target overlaps with existing rules.
    pub fn add_rule(
        mut self,
        target: ValidationTarget,
        rule: ValidationRule,
    ) -> Result<Self, ValidationError> {
        // For now, just add the rule - overlap checking happens at apply time
        self.rules.push((target, rule));
        Ok(self)
    }

    /// Adds a rule without overlap checking (builder pattern)
    pub fn with_rule(mut self, target: ValidationTarget, rule: ValidationRule) -> Self {
        self.rules.push((target, rule));
        self
    }

    /// Returns all rules
    pub fn rules(&self) -> &[(ValidationTarget, ValidationRule)] {
        &self.rules
    }

    /// Checks if any rules are defined
    pub fn is_empty(&self) -> bool {
        self.rules.is_empty()
    }
}

/// Validation configuration error
#[derive(Debug, Clone)]
pub enum ValidationError {
    /// Same cell has multiple validation rules
    OverlappingRules { cell: String, rules: Vec<String> },
    /// Invalid rule configuration
    InvalidRule(String),
}

impl std::fmt::Display for ValidationError {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            ValidationError::OverlappingRules { cell, rules } => {
                write!(f, "Cell {} has overlapping rules: {:?}", cell, rules)
            }
            ValidationError::InvalidRule(msg) => write!(f, "Invalid rule: {}", msg),
        }
    }
}

impl std::error::Error for ValidationError {}

// ============================================================================
// Row Grouping (Outlining)
// ============================================================================

/// Configuration for row grouping/outlining.
///
/// Allows grouping rows together with expand/collapse functionality.
/// Groups can be nested up to 7 levels deep (Excel limitation).
///
/// ## Serde (requires `serde` feature)
///
/// ```json
/// {
///   "groups": [
///     { "start_row": 10, "end_row": 20, "collapsed": false },
///     { "start_row": 25, "end_row": 30, "collapsed": true }
///   ],
///   "symbols_above": false
/// }
/// ```
///
/// Row indices are 0-based.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
#[cfg_attr(feature = "serde", derive(serde::Serialize, serde::Deserialize))]
pub struct RowGrouping {
    /// List of row groups
    groups: Vec<RowGroup>,
    /// Show outline symbols above grouped rows (default: false = below)
    pub symbols_above: bool,
}

/// A single row group definition.
///
/// ## JSON
///
/// ```json
/// { "start_row": 10, "end_row": 20, "collapsed": false }
/// ```
#[derive(Debug, Clone, PartialEq, Eq)]
#[cfg_attr(feature = "serde", derive(serde::Serialize, serde::Deserialize))]
pub struct RowGroup {
    /// First row of the group (0-indexed)
    pub start_row: u32,
    /// Last row of the group (0-indexed, inclusive)
    pub end_row: u32,
    /// Whether the group is initially collapsed
    pub collapsed: bool,
}

impl RowGroup {
    /// Creates a new row group (expanded by default)
    pub fn new(start_row: u32, end_row: u32) -> Self {
        Self {
            start_row,
            end_row,
            collapsed: false,
        }
    }

    /// Creates a new collapsed row group
    pub fn collapsed(start_row: u32, end_row: u32) -> Self {
        Self {
            start_row,
            end_row,
            collapsed: true,
        }
    }
}

impl RowGrouping {
    pub fn new() -> Self {
        Self::default()
    }

    /// Adds a row group (expanded by default)
    pub fn add_group(mut self, start_row: u32, end_row: u32) -> Self {
        self.groups.push(RowGroup::new(start_row, end_row));
        self
    }

    /// Adds a collapsed row group
    pub fn add_collapsed_group(mut self, start_row: u32, end_row: u32) -> Self {
        self.groups.push(RowGroup::collapsed(start_row, end_row));
        self
    }

    /// Adds a custom row group
    pub fn add_custom_group(mut self, group: RowGroup) -> Self {
        self.groups.push(group);
        self
    }

    /// Sets whether outline symbols appear above grouped rows
    pub fn with_symbols_above(mut self, above: bool) -> Self {
        self.symbols_above = above;
        self
    }

    /// Returns all row groups
    pub fn groups(&self) -> &[RowGroup] {
        &self.groups
    }

    /// Returns true if any groups are defined
    pub fn has_groups(&self) -> bool {
        !self.groups.is_empty()
    }

    /// Returns true if no groups are defined
    pub fn is_empty(&self) -> bool {
        self.groups.is_empty()
    }
}

// ============================================================================
// Hidden Ranges (Columns and Rows)
// ============================================================================

/// Configuration for hiding columns and rows
#[derive(Debug, Clone, Default)]
pub struct HiddenRanges {
    /// Columns to hide (0-indexed)
    columns: Vec<HiddenRange>,
    /// Rows to hide (0-indexed)
    rows: Vec<HiddenRange>,
}

/// A range to hide (inclusive start and end)
#[derive(Debug, Clone, Copy)]
pub struct HiddenRange {
    /// Start index (0-indexed)
    pub start: u32,
    /// End index (0-indexed, inclusive)
    pub end: u32,
}

impl HiddenRange {
    /// Creates a new range from start to end (inclusive)
    pub const fn new(start: u32, end: u32) -> Self {
        Self { start, end }
    }

    /// Creates a range for a single column/row
    pub const fn single(index: u32) -> Self {
        Self {
            start: index,
            end: index,
        }
    }

    /// Creates a range from Excel column letters (e.g., "Q", "V")
    pub fn from_col_letters(start: &str, end: &str) -> Self {
        Self {
            start: col_letter_to_index(start),
            end: col_letter_to_index(end),
        }
    }
}

/// Converts Excel column letter(s) to 0-indexed number
/// A=0, B=1, ..., Z=25, AA=26, AB=27, ...
fn col_letter_to_index(col: &str) -> u32 {
    let mut result = 0u32;
    for c in col.to_uppercase().chars() {
        if c.is_ascii_alphabetic() {
            result = result * 26 + (c as u32 - 'A' as u32 + 1);
        }
    }
    result.saturating_sub(1) // Convert to 0-indexed
}

impl HiddenRanges {
    pub fn new() -> Self {
        Self::default()
    }

    /// Hides a single column (0-indexed)
    pub fn hide_column(mut self, col: u16) -> Self {
        self.columns.push(HiddenRange::single(col as u32));
        self
    }

    /// Hides a range of columns (0-indexed, inclusive)
    pub fn hide_columns(mut self, start: u16, end: u16) -> Self {
        self.columns
            .push(HiddenRange::new(start as u32, end as u32));
        self
    }

    /// Hides columns by Excel letter notation (e.g., "Q", "V")
    pub fn hide_columns_by_letter(mut self, start: &str, end: &str) -> Self {
        self.columns.push(HiddenRange::from_col_letters(start, end));
        self
    }

    /// Hides a single row (0-indexed)
    pub fn hide_row(mut self, row: u32) -> Self {
        self.rows.push(HiddenRange::single(row));
        self
    }

    /// Hides a range of rows (0-indexed, inclusive)
    pub fn hide_rows(mut self, start: u32, end: u32) -> Self {
        self.rows.push(HiddenRange::new(start, end));
        self
    }

    /// Preset: Hide columns Q:V (16-21, right panel duplicate)
    pub fn preset_hide_qv() -> Self {
        Self::new().hide_columns(16, 21)
    }

    /// Returns all column ranges to hide
    pub fn column_ranges(&self) -> &[HiddenRange] {
        &self.columns
    }

    /// Returns all row ranges to hide
    pub fn row_ranges(&self) -> &[HiddenRange] {
        &self.rows
    }

    /// Returns true if any ranges are configured
    pub fn has_ranges(&self) -> bool {
        !self.columns.is_empty() || !self.rows.is_empty()
    }

    /// Returns true if no ranges are configured
    pub fn is_empty(&self) -> bool {
        self.columns.is_empty() && self.rows.is_empty()
    }
}

// ============================================================================
// Sheet Options (combines Protection + Validation + other settings)
// ============================================================================

/// Combined sheet-level options for protection, validation, and display settings
#[derive(Debug, Clone, Default)]
pub struct SheetOptions {
    /// Sheet protection settings
    pub protection: Option<SheetProtection>,

    /// Field validations
    pub validation: Option<FieldValidation>,

    /// Hidden columns and rows
    pub hidden: HiddenRanges,

    /// Row grouping/outlining
    pub row_grouping: RowGrouping,

    /// Language for error messages (uses LANG_CONFIG)
    pub language: Option<String>,
}

impl SheetOptions {
    pub fn new() -> Self {
        Self::default()
    }

    /// Creates options with default protection
    pub fn with_default_protection() -> Self {
        Self {
            protection: Some(SheetProtection::from_defaults()),
            ..Default::default()
        }
    }

    /// Sets custom protection
    pub fn with_protection(mut self, protection: SheetProtection) -> Self {
        self.protection = Some(protection);
        self
    }

    /// Sets field validations
    pub fn with_validation(mut self, validation: FieldValidation) -> Self {
        self.validation = Some(validation);
        self
    }

    /// Sets hidden ranges configuration
    pub fn with_hidden(mut self, hidden: HiddenRanges) -> Self {
        self.hidden = hidden;
        self
    }

    /// Convenience: Hides columns Q:V (right panel duplicate)
    pub fn with_hidden_columns_qv(mut self) -> Self {
        self.hidden = HiddenRanges::preset_hide_qv();
        self
    }

    /// Convenience: Hides specific columns by letter
    pub fn hide_columns(mut self, start: &str, end: &str) -> Self {
        self.hidden = self.hidden.hide_columns_by_letter(start, end);
        self
    }

    /// Convenience: Hides specific rows
    pub fn hide_rows(mut self, start: u32, end: u32) -> Self {
        self.hidden = self.hidden.hide_rows(start, end);
        self
    }

    /// Sets language for error messages
    pub fn with_language(mut self, language: impl Into<String>) -> Self {
        self.language = Some(language.into());
        self
    }

    /// Gets the language, defaulting to "Deutsch"
    pub fn language(&self) -> &str {
        self.language.as_deref().unwrap_or("Deutsch")
    }

    /// Sets row grouping configuration
    pub fn with_row_grouping(mut self, grouping: RowGrouping) -> Self {
        self.row_grouping = grouping;
        self
    }

    /// Convenience: Adds a row group (expanded)
    pub fn add_row_group(mut self, start_row: u32, end_row: u32) -> Self {
        self.row_grouping = self.row_grouping.add_group(start_row, end_row);
        self
    }

    /// Convenience: Adds a collapsed row group
    pub fn add_collapsed_row_group(mut self, start_row: u32, end_row: u32) -> Self {
        self.row_grouping = self.row_grouping.add_collapsed_group(start_row, end_row);
        self
    }

}

// ============================================================================
// Tests
// ============================================================================

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_sheet_protection_defaults() {
        let prot = SheetProtection::new();
        assert!(prot.select_locked_cells);
        assert!(prot.select_unlocked_cells);
        assert!(!prot.format_cells);
        assert!(!prot.insert_columns);
    }

    #[test]
    fn test_sheet_protection_builder() {
        let prot = SheetProtection::new()
            .with_password("test123")
            .allow_format_cells(true)
            .allow_sort(true);

        assert_eq!(prot.password, Some("test123".to_string()));
        assert!(prot.format_cells);
        assert!(prot.sort);
    }

    #[test]
    fn test_sheet_protection_from_defaults() {
        let prot = SheetProtection::from_defaults();
        // Check that it uses recommended default settings
        assert!(prot.format_cells);
        assert!(prot.sort);
        assert!(!prot.insert_columns);
    }

    #[test]
    fn test_validation_rule_decimal() {
        let rule = ValidationRule::decimal_between(0.0, 1000.0)
            .with_error_title("Invalid")
            .with_error_message("Enter 0-1000");

        assert!(matches!(
            rule.rule_type,
            ValidationRuleType::Decimal(NumericRule::Between(0.0, 1000.0))
        ));
        assert_eq!(rule.error_title, Some("Invalid".to_string()));
    }

    #[test]
    fn test_validation_rule_currency() {
        let rule = ValidationRule::currency().with_currency_error(Some("Deutsch"));

        assert!(matches!(rule.rule_type, ValidationRuleType::Custom(_)));
        assert!(rule.error_message.is_some());
    }

    #[test]
    fn test_field_validation_builder() {
        let validation = FieldValidation::new()
            .with_rule(
                ValidationTarget::Field(ApiKey::ApprovedBudget(0)),
                ValidationRule::decimal_non_negative(),
            )
            .with_rule(
                ValidationTarget::RightPanelAmounts,
                ValidationRule::decimal_between(0.0, 999999.99),
            );

        assert_eq!(validation.rules().len(), 2);
    }

    #[test]
    fn test_report_options() {
        let opts = SheetOptions::with_default_protection()
            .with_hidden_columns_qv()
            .with_language("Englisch");

        assert!(opts.protection.is_some());
        assert!(!opts.hidden.is_empty());
        assert_eq!(opts.hidden.column_ranges().len(), 1);
        assert_eq!(opts.language(), "Englisch");
    }

    #[test]
    fn test_numeric_rule_conversion() {
        let rule = NumericRule::Between(0.0f64, 100.0);
        let xlsx_rule = rule.to_xlsx_rule();
        assert!(matches!(xlsx_rule, DataValidationRule::Between(0.0, 100.0)));
    }

    #[test]
    fn test_row_grouping() {
        let grouping = RowGrouping::new()
            .add_group(10, 20)
            .add_collapsed_group(25, 30)
            .with_symbols_above(true);

        assert_eq!(grouping.groups().len(), 2);
        assert!(grouping.has_groups());
        assert!(grouping.symbols_above);

        let first = &grouping.groups()[0];
        assert_eq!(first.start_row, 10);
        assert_eq!(first.end_row, 20);
        assert!(!first.collapsed);

        let second = &grouping.groups()[1];
        assert_eq!(second.start_row, 25);
        assert_eq!(second.end_row, 30);
        assert!(second.collapsed);
    }

    #[test]
    fn test_report_options_row_grouping() {
        let opts = SheetOptions::new()
            .add_row_group(10, 20)
            .add_collapsed_row_group(25, 30);

        assert!(!opts.row_grouping.is_empty());
        assert_eq!(opts.row_grouping.groups().len(), 2);
    }

}
