//! Report generation engine.
//!
//! ## Module Structure
//!
//! - [`api`]: Public API types (ApiKey, Language, Currency, ReportValues, etc.)
//! - [`core`]: Core types (`CellAddr`)
//! - `calc`: IronCalc formula engine (`ModelTemplate`, `CalcBridge`)
//! - [`options`]: Sheet-level options (SheetOptions, SheetProtection, Validation, etc.)
//! - `styles`: Internal styles and format matrix (`ReportStyles`, `FormatMatrix`)
//! - [`writer`]: Excel writer (write_report_with_options, create_protected_report)
//! - [`body`]: Dynamic cost position sections (BodyConfig, BodyLayout, FooterLayout)

pub mod api;
pub mod body;
pub(crate) mod calc;
pub mod core;
pub mod options;
pub(crate) mod styles;
pub mod writer;

// Re-exports for convenient access
pub use api::{
    ApiKey, Category, CellValue, Currency, DateError, FooterField,
    Language, PositionField, ReportDate, ReportValues,
};
pub use body::{BodyConfig, BodyLayout, CategoryMode};
pub use options::{
    FieldValidation, HiddenRanges, SheetOptions, RowGroup, RowGrouping, SheetProtection,
    ValidationError, ValidationErrorStyle, ValidationRule, ValidationRuleType, ValidationTarget,
};
pub use writer::{
    apply_sheet_options, setup_freeze_panes, setup_sheet, write_report_with_options, BodyResult,
    MergeRange,
};
