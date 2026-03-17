//! # FB-Generator — Financial Report Generator
//!
//! High-performance Excel financial report generator with dynamic cost positions,
//! formulas, formatting, and multi-language support.
//!
//! ## Quick Start (Rust)
//!
//! [`ReportConfig`] is the main entry point. Construct it, then call [`write_to`](ReportConfig::write_to):
//!
//! ```ignore
//! use fb_generator::*;
//!
//! let config = ReportConfig::builder()
//!     .header(ReportHeader::builder().language(Language::Deutsch).currency(Currency::EUR).build())
//!     .options(ReportOptions::builder().sheet_protection(SheetProtection::from_defaults().with_password("geheim")).build())
//!     .build();  // infallible — no ? required
//! config.write_to("report.xlsx")?;
//! ```
//!
//! ### Batch with precomputed hash
//!
//! When generating multiple reports with the same workbook password, precompute the
//! SHA-512 hash once and reuse it (~25ms saved per file):
//!
//! ```ignore
//! use fb_generator::{precompute_hash, ReportConfig};
//!
//! let hash = precompute_hash("password");
//! for config in &configs {
//!     config.write_to_precomputed("output.xlsx", &hash)?;
//! }
//! ```
//!
//! ## Quick Start (Tauri / JSON)
//!
//! With the `serde` feature enabled, [`ReportConfig`] can be deserialized directly
//! from a Tauri command payload:
//!
//! ```json
//! {
//!   "header": {
//!     "language": "deutsch",
//!     "currency": "EUR",
//!     "project_number": "2025-001",
//!     "project_title": "Klimaschutzprojekt"
//!   },
//!   "options": { "sheet_protection": { "password": "geheim" } }
//! }
//! ```
//!
//! All sub-structs (`header`, `body`, `footer`, `options`) default when omitted.
//! `deny_unknown_fields` is active on each sub-struct — misspelled keys produce a
//! serde error.
//!
//! ## Feature Flags
//!
//! | Flag    | Effect |
//! |---------|--------|
//! | `serde` | Enables `Serialize`/`Deserialize` on all public API types. Required for Tauri integration. Activates `deny_unknown_fields` on [`ReportConfig`], enum validation on [`Language`], `try_from` validation on [`Currency`] and [`ReportDate`]. |
//!
//! ## Module Overview
//!
//! | Module | Purpose |
//! |--------|---------|
//! | [`config`] | [`ReportConfig`] — main entry point for report generation |
//! | [`report::api`] | Validated types: [`Language`], [`Currency`], [`Category`], [`ReportDate`] |
//! | [`report::body`] | [`BodyConfig`] — category layout (positions per category) |
//! | [`report::options`] | Sheet protection, data validation, row grouping |
//! | [`workbook_protection`] | Workbook-level password protection (ECMA-376 SHA-512) |
//! | [`error`] | [`ReportError`] — all error variants from report generation |

pub mod config;
pub mod error;
pub mod lang;
pub mod report;
pub mod workbook_protection;

pub use error::ReportError;
pub use workbook_protection::{
    precompute_hash, precompute_hash_with_spin_count,
    protect_workbook, protect_workbook_precomputed, protect_workbook_with_spin_count,
    PrecomputedHash, ProtectionError, WorkbookProtection,
};

// Public API re-exports
pub use config::{
    IncomeTable,
    PanelEntry, PanelEntryBuilder, PositionEntry, PositionEntryBuilder,
    ReportBody, ReportBodyBuilder, ReportConfig, ReportConfigBuilder,
    ReportFooter, ReportFooterBuilder, ReportHeader, ReportHeaderBuilder,
    ReportOptions, ReportOptionsBuilder, TableEntry, TableEntryBuilder,
};
pub use lang::{
    build_sheet as build_language_sheet,
    build_sheet_with_visibility as build_language_sheet_with_visibility,
};
pub use report::{
    api::{ApiKey, Category, Currency, DateError, Language, ReportDate, ReportValues},
    body::BodyConfig,
    options::{
        FieldValidation, HiddenRanges, SheetOptions, RowGroup, RowGrouping,
        SheetProtection, ValidationError, ValidationErrorStyle, ValidationRule, ValidationRuleType,
        ValidationTarget,
    },
    writer::write_report_with_options,
};
