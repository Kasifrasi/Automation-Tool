//! Unified report configuration for Tauri integration.
//!
//! This module provides [`ReportConfig`] as the main entry point, composed of
//! four domain sub-structs: [`ReportHeader`], [`ReportBody`], [`ReportFooter`],
//! and [`ReportOptions`].
//!
//! All structs use a hand-written fluent builder with infallible `.build()`.
//!
//! ## Tauri Example
//!
//! ```ignore
//! #[tauri::command]
//! fn generate_report(config: ReportConfig, output_path: String) -> Result<(), String> {
//!     config.write_to(output_path).map_err(|e| e.to_string())
//! }
//! ```
//!
//! ## TypeScript Example
//!
//! ```typescript
//! await invoke("generate_report", {
//!     config: {
//!         header: {
//!             language: "deutsch",
//!             currency: "EUR",
//!             project_number: "2025-001",
//!             project_title: "Klimaschutzprojekt",
//!         },
//!         body: {
//!             positions: { "1": [{ "description": "Baukosten", "approved": 50000.0 }] },
//!             header_inputs: { "6": { "approved": 3000.0 } },
//!         },
//!         footer: { "bank": 12000.0, "kasse": 500.0 },
//!         options: { "sheet_protection": { "password": "geheim" } },
//!     },
//!     outputPath: "/home/user/report.xlsx"
//! });
//! ```

use std::collections::HashMap;
use std::path::Path;

use crate::report::api::{ApiKey, CellValue, Currency, Language, ReportValues};
use crate::report::body::BodyConfig;
use crate::report::options::{RowGrouping, SheetOptions, SheetProtection};
use crate::report::writer::{create_protected_report, create_protected_report_precomputed};
use crate::workbook_protection::WorkbookProtection;

// ============================================================================
// Entry types
// ============================================================================

/// A single row in the upper income table (Excel rows 15–19).
///
/// Used as a named field in [`IncomeTable`] — each row has a fixed semantic identity
/// (`saldovortrag`, `eigenmittel`, `drittmittel`, `kmw_mittel`, `zinsertraege`).
///
/// ## JSON
///
/// ```json
/// { "approved_budget": 50000.0, "income_report": 25000.0, "income_total": 25000.0, "reason": "Spende" }
/// ```
#[derive(Debug, Clone, PartialEq, Default)]
#[cfg_attr(feature = "serde", derive(serde::Serialize, serde::Deserialize))]
#[cfg_attr(feature = "serde", serde(deny_unknown_fields))]
pub struct TableEntry {
    /// Column D: Approved budget
    pub approved_budget: Option<f64>,
    /// Column E: Income report period
    pub income_report: Option<f64>,
    /// Column F: Income total
    pub income_total: Option<f64>,
    /// Column H: Reason
    pub reason: Option<String>,
}

impl TableEntry {
    pub fn builder() -> TableEntryBuilder {
        TableEntryBuilder::default()
    }
}

// ============================================================================

/// The five fixed income rows of the upper table (Excel rows 15–19).
///
/// Each field corresponds to a named, fixed row in the Excel template.
/// Leaving a field as `None` produces an empty row. It is impossible to
/// add more than five rows or to set the same row twice.
///
/// ## JSON
///
/// ```json
/// {
///   "saldovortrag": { "approved_budget": 5000.0, "income_report": 2500.0 },
///   "kmw_mittel":   { "approved_budget": 80000.0, "income_report": 50000.0, "income_total": 50000.0 }
/// }
/// ```
#[derive(Debug, Clone, PartialEq, Default)]
#[cfg_attr(feature = "serde", derive(serde::Serialize, serde::Deserialize))]
#[cfg_attr(feature = "serde", serde(deny_unknown_fields))]
pub struct IncomeTable {
    /// Row 1: Carry-forward balance from previous period (Saldovortrag)
    pub saldovortrag: Option<TableEntry>,
    /// Row 2: Own funds / equity contribution (Eigenmittel)
    pub eigenmittel: Option<TableEntry>,
    /// Row 3: Third-party funds (Drittmittel)
    pub drittmittel: Option<TableEntry>,
    /// Row 4: KMW grant funds (KMW-Mittel)
    pub kmw_mittel: Option<TableEntry>,
    /// Row 5: Interest income (Zinserträge)
    pub zinsertraege: Option<TableEntry>,
}

/// Builder for [`TableEntry`].
#[derive(Default)]
pub struct TableEntryBuilder {
    approved_budget: Option<f64>,
    income_report: Option<f64>,
    income_total: Option<f64>,
    reason: Option<String>,
}

impl TableEntryBuilder {
    pub fn approved_budget(mut self, v: f64) -> Self {
        self.approved_budget = Some(v);
        self
    }
    pub fn income_report(mut self, v: f64) -> Self {
        self.income_report = Some(v);
        self
    }
    pub fn income_total(mut self, v: f64) -> Self {
        self.income_total = Some(v);
        self
    }
    pub fn reason(mut self, v: impl Into<String>) -> Self {
        self.reason = Some(v.into());
        self
    }
    pub fn build(self) -> TableEntry {
        TableEntry {
            approved_budget: self.approved_budget,
            income_report: self.income_report,
            income_total: self.income_total,
            reason: self.reason,
        }
    }
}

// ============================================================================

/// A single row in a cash book panel (left or right, max 18 rows each).
///
/// Row position is determined by insertion order — no manual index required.
///
/// ## JSON
///
/// ```json
/// { "date": "2025-01-15", "amount_euro": 1200.50, "amount_local": null }
/// ```
#[derive(Debug, Clone, PartialEq, Default)]
#[cfg_attr(feature = "serde", derive(serde::Serialize, serde::Deserialize))]
#[cfg_attr(feature = "serde", serde(deny_unknown_fields))]
pub struct PanelEntry {
    /// Date string (e.g. "2025-01-15" or "15.01.2025")
    pub date: Option<String>,
    /// Amount in Euro
    pub amount_euro: Option<f64>,
    /// Amount in local currency
    pub amount_local: Option<f64>,
}

impl PanelEntry {
    pub fn builder() -> PanelEntryBuilder {
        PanelEntryBuilder::default()
    }
}

/// Builder for [`PanelEntry`].
#[derive(Default)]
pub struct PanelEntryBuilder {
    date: Option<String>,
    amount_euro: Option<f64>,
    amount_local: Option<f64>,
}

impl PanelEntryBuilder {
    pub fn date(mut self, v: impl Into<String>) -> Self {
        self.date = Some(v.into());
        self
    }
    pub fn amount_euro(mut self, v: f64) -> Self {
        self.amount_euro = Some(v);
        self
    }
    pub fn amount_local(mut self, v: f64) -> Self {
        self.amount_local = Some(v);
        self
    }
    pub fn build(self) -> PanelEntry {
        PanelEntry {
            date: self.date,
            amount_euro: self.amount_euro,
            amount_local: self.amount_local,
        }
    }
}

// ============================================================================

/// A single cost position row in the body area.
///
/// Used both for normal positions (`add_position`) and header-input entries
/// (`set_header_input`). In header-input mode `description` is silently ignored.
///
/// ## JSON (normal position, inside `positions["1"]`)
///
/// ```json
/// { "description": "Personalkosten", "approved": 50000.0 }
/// ```
///
/// ## JSON (header-input, inside `header_inputs["6"]`)
///
/// ```json
/// { "approved": 3000.0, "remark": "Sonstiges" }
/// ```
#[derive(Debug, Clone, PartialEq, Default)]
#[cfg_attr(feature = "serde", derive(serde::Serialize, serde::Deserialize))]
#[cfg_attr(feature = "serde", serde(deny_unknown_fields))]
pub struct PositionEntry {
    /// Column C: Description (normal positions only; ignored in header-input mode)
    pub description: Option<String>,
    /// Column D: Approved budget
    pub approved: Option<f64>,
    /// Column E: Income report period
    pub income_report: Option<f64>,
    /// Column F: Income total
    pub income_total: Option<f64>,
    /// Column H: Remark
    pub remark: Option<String>,
}

impl PositionEntry {
    pub fn builder() -> PositionEntryBuilder {
        PositionEntryBuilder::default()
    }
}

/// Builder for [`PositionEntry`].
#[derive(Default)]
pub struct PositionEntryBuilder {
    description: Option<String>,
    approved: Option<f64>,
    income_report: Option<f64>,
    income_total: Option<f64>,
    remark: Option<String>,
}

impl PositionEntryBuilder {
    pub fn description(mut self, v: impl Into<String>) -> Self {
        self.description = Some(v.into());
        self
    }
    pub fn approved(mut self, v: f64) -> Self {
        self.approved = Some(v);
        self
    }
    pub fn income_report(mut self, v: f64) -> Self {
        self.income_report = Some(v);
        self
    }
    pub fn income_total(mut self, v: f64) -> Self {
        self.income_total = Some(v);
        self
    }
    pub fn remark(mut self, v: impl Into<String>) -> Self {
        self.remark = Some(v.into());
        self
    }
    pub fn build(self) -> PositionEntry {
        PositionEntry {
            description: self.description,
            approved: self.approved,
            income_report: self.income_report,
            income_total: self.income_total,
            remark: self.remark,
        }
    }
}

// ============================================================================
// ReportHeader
// ============================================================================

/// Header metadata for a financial report.
///
/// ## JSON
///
/// ```json
/// {
///   "language": "deutsch",
///   "currency": "EUR",
///   "project_number": "PROJ-2025-001",
///   "project_title": "Bildungsprojekt",
///   "project_start": "01.01.2025",
///   "project_end": "31.12.2025",
///   "report_start": "01.01.2025",
///   "report_end": "30.06.2025",
///   "version": "v2025-1"
/// }
/// ```
#[derive(Debug, Clone, PartialEq)]
#[cfg_attr(feature = "serde", derive(serde::Serialize, serde::Deserialize))]
#[cfg_attr(feature = "serde", serde(deny_unknown_fields))]
pub struct ReportHeader {
    /// Language (validated enum)
    pub language: Language,
    /// Currency (validated ISO 4217 code)
    pub currency: Currency,
    /// Project number (cell D5)
    pub project_number: Option<String>,
    /// Project title (cell D6)
    pub project_title: Option<String>,
    /// Project start date (cell E8)
    pub project_start: Option<String>,
    /// Project end date (cell G8)
    pub project_end: Option<String>,
    /// Report period start (cell E9)
    pub report_start: Option<String>,
    /// Report period end (cell G9)
    pub report_end: Option<String>,
    /// Version / identifier text written to cell B2
    pub version: Option<String>,
}

impl Default for ReportHeader {
    fn default() -> Self {
        Self {
            language: Language::Deutsch,
            currency: Currency::EUR,
            project_number: None,
            project_title: None,
            project_start: None,
            project_end: None,
            report_start: None,
            report_end: None,
            version: None,
        }
    }
}

impl ReportHeader {
    pub fn builder() -> ReportHeaderBuilder {
        ReportHeaderBuilder::default()
    }
}

/// Builder for [`ReportHeader`].
#[derive(Default)]
pub struct ReportHeaderBuilder {
    language: Option<Language>,
    currency: Option<Currency>,
    project_number: Option<String>,
    project_title: Option<String>,
    project_start: Option<String>,
    project_end: Option<String>,
    report_start: Option<String>,
    report_end: Option<String>,
    version: Option<String>,
}

impl ReportHeaderBuilder {
    pub fn language(mut self, v: Language) -> Self {
        self.language = Some(v);
        self
    }
    pub fn currency(mut self, v: Currency) -> Self {
        self.currency = Some(v);
        self
    }
    pub fn project_number(mut self, v: impl Into<String>) -> Self {
        self.project_number = Some(v.into());
        self
    }
    pub fn project_title(mut self, v: impl Into<String>) -> Self {
        self.project_title = Some(v.into());
        self
    }
    pub fn project_start(mut self, v: impl Into<String>) -> Self {
        self.project_start = Some(v.into());
        self
    }
    pub fn project_end(mut self, v: impl Into<String>) -> Self {
        self.project_end = Some(v.into());
        self
    }
    pub fn report_start(mut self, v: impl Into<String>) -> Self {
        self.report_start = Some(v.into());
        self
    }
    pub fn report_end(mut self, v: impl Into<String>) -> Self {
        self.report_end = Some(v.into());
        self
    }
    /// Version / identifier text for cell B2
    pub fn version(mut self, v: impl Into<String>) -> Self {
        self.version = Some(v.into());
        self
    }
    pub fn build(self) -> ReportHeader {
        let def = ReportHeader::default();
        ReportHeader {
            language: self.language.unwrap_or(def.language),
            currency: self.currency.unwrap_or(def.currency),
            project_number: self.project_number,
            project_title: self.project_title,
            project_start: self.project_start,
            project_end: self.project_end,
            report_start: self.report_start,
            report_end: self.report_end,
            version: self.version,
        }
    }
}

// ============================================================================
// ReportBody
// ============================================================================

/// Body data: income table, cash book panels, and cost positions.
///
/// Row positions are derived from insertion order — no manual indices required.
/// The income table uses named fields instead
/// of a positional array, preventing duplicates and overflow at the type level.
///
/// ## JSON
///
/// ```json
/// {
///   "table": { "kmw_mittel": { "approved_budget": 80000.0 }, "saldovortrag": { "approved_budget": 5000.0 } },
///   "left_panel": [{ "date": "15.01.2025", "amount_euro": 1000.0 }, null],
///   "right_panel": [],
///   "positions": { "1": [{ "description": "Personnel", "approved": 18000.0 }] },
///   "header_inputs": { "6": { "approved": 3000.0 } }
/// }
/// ```
#[derive(Debug, Clone, PartialEq, Default)]
#[cfg_attr(feature = "serde", derive(serde::Serialize, serde::Deserialize))]
#[cfg_attr(feature = "serde", serde(deny_unknown_fields))]
pub struct ReportBody {
    /// Income table: five named rows (Saldovortrag, Eigenmittel, Drittmittel, KMW-Mittel, Zinserträge)
    #[cfg_attr(feature = "serde", serde(default))]
    pub table: IncomeTable,
    /// Left cash book panel (max 18 rows); `None` = empty row
    #[cfg_attr(feature = "serde", serde(default))]
    pub left_panel: Vec<Option<PanelEntry>>,
    /// Right cash book panel (max 18 rows); `None` = empty row
    #[cfg_attr(feature = "serde", serde(default))]
    pub right_panel: Vec<Option<PanelEntry>>,
    /// Normal cost positions per category (1–8).
    /// Vec index + 1 = 1-based position number; `None` = empty row.
    /// Vec length determines the row count allocated for this category.
    #[cfg_attr(feature = "serde", serde(default))]
    pub positions: HashMap<u8, Vec<Option<PositionEntry>>>,
    /// Header-input entries: single aggregated value per category (no individual rows).
    #[cfg_attr(feature = "serde", serde(default))]
    pub header_inputs: HashMap<u8, Option<PositionEntry>>,
}

impl ReportBody {
    pub fn builder() -> ReportBodyBuilder {
        ReportBodyBuilder::default()
    }
}

/// Fluent builder for [`ReportBody`].
///
/// Supports `skip_*_row` methods to insert empty rows and batch
/// `add_*_entries` methods for bulk insertion.
#[derive(Default)]
pub struct ReportBodyBuilder {
    table: IncomeTable,
    left_panel: Vec<Option<PanelEntry>>,
    right_panel: Vec<Option<PanelEntry>>,
    positions: HashMap<u8, Vec<Option<PositionEntry>>>,
    header_inputs: HashMap<u8, Option<PositionEntry>>,
}

impl ReportBodyBuilder {
    // --- Income table (5 named rows — each can only be set once) ---

    /// Row 1: Carry-forward balance from previous period (Saldovortrag)
    pub fn saldovortrag(mut self, e: TableEntry) -> Self {
        self.table.saldovortrag = Some(e);
        self
    }
    /// Row 2: Own funds / equity contribution (Eigenmittel)
    pub fn eigenmittel(mut self, e: TableEntry) -> Self {
        self.table.eigenmittel = Some(e);
        self
    }
    /// Row 3: Third-party funds (Drittmittel)
    pub fn drittmittel(mut self, e: TableEntry) -> Self {
        self.table.drittmittel = Some(e);
        self
    }
    /// Row 4: KMW grant funds (KMW-Mittel)
    pub fn kmw_mittel(mut self, e: TableEntry) -> Self {
        self.table.kmw_mittel = Some(e);
        self
    }
    /// Row 5: Interest income (Zinserträge)
    pub fn zinsertraege(mut self, e: TableEntry) -> Self {
        self.table.zinsertraege = Some(e);
        self
    }

    // --- Left panel ---

    pub fn add_left_panel_entry(mut self, e: PanelEntry) -> Self {
        self.left_panel.push(Some(e));
        self
    }
    pub fn skip_left_panel_row(mut self) -> Self {
        self.left_panel.push(None);
        self
    }
    pub fn add_left_panel_entries(mut self, it: impl IntoIterator<Item = PanelEntry>) -> Self {
        for e in it {
            self.left_panel.push(Some(e));
        }
        self
    }

    // --- Right panel ---

    pub fn add_right_panel_entry(mut self, e: PanelEntry) -> Self {
        self.right_panel.push(Some(e));
        self
    }
    pub fn skip_right_panel_row(mut self) -> Self {
        self.right_panel.push(None);
        self
    }
    pub fn add_right_panel_entries(mut self, it: impl IntoIterator<Item = PanelEntry>) -> Self {
        for e in it {
            self.right_panel.push(Some(e));
        }
        self
    }

    // --- Normal positions (any category 1–8) ---

    pub fn add_position(mut self, category: u8, e: PositionEntry) -> Self {
        self.positions.entry(category).or_default().push(Some(e));
        self
    }
    pub fn skip_position_row(mut self, category: u8) -> Self {
        self.positions.entry(category).or_default().push(None);
        self
    }
    pub fn add_positions(
        mut self,
        category: u8,
        it: impl IntoIterator<Item = PositionEntry>,
    ) -> Self {
        for e in it {
            self.positions.entry(category).or_default().push(Some(e));
        }
        self
    }

    // --- Header-input (single aggregated row, no individual positions) ---

    pub fn set_header_input(mut self, category: u8, e: PositionEntry) -> Self {
        self.header_inputs.insert(category, Some(e));
        self
    }

    pub fn build(self) -> ReportBody {
        ReportBody {
            table: self.table,
            left_panel: self.left_panel,
            right_panel: self.right_panel,
            positions: self.positions,
            header_inputs: self.header_inputs,
        }
    }
}

// ============================================================================
// ReportFooter
// ============================================================================

/// Footer balances for the reconciliation section.
///
/// ## JSON
///
/// ```json
/// { "bank": 8500.0, "kasse": 250.50, "sonstiges": null }
/// ```
#[derive(Debug, Clone, PartialEq, Default)]
#[cfg_attr(feature = "serde", derive(serde::Serialize, serde::Deserialize))]
#[cfg_attr(feature = "serde", serde(deny_unknown_fields))]
pub struct ReportFooter {
    /// Bank balance
    pub bank: Option<f64>,
    /// Cash balance
    pub kasse: Option<f64>,
    /// Other balances (checks, advances, etc.)
    pub sonstiges: Option<f64>,
}

impl ReportFooter {
    pub fn builder() -> ReportFooterBuilder {
        ReportFooterBuilder::default()
    }
}

/// Builder for [`ReportFooter`].
#[derive(Default)]
pub struct ReportFooterBuilder {
    bank: Option<f64>,
    kasse: Option<f64>,
    sonstiges: Option<f64>,
}

impl ReportFooterBuilder {
    pub fn bank(mut self, v: f64) -> Self {
        self.bank = Some(v);
        self
    }
    pub fn kasse(mut self, v: f64) -> Self {
        self.kasse = Some(v);
        self
    }
    pub fn sonstiges(mut self, v: f64) -> Self {
        self.sonstiges = Some(v);
        self
    }
    pub fn build(self) -> ReportFooter {
        ReportFooter {
            bank: self.bank,
            kasse: self.kasse,
            sonstiges: self.sonstiges,
        }
    }
}

// ============================================================================
// ReportOptions
// ============================================================================

/// Output options: protection, visibility, and row grouping.
///
/// ## JSON
///
/// ```json
/// {
///   "sheet_protection": {
///     "password": "sheet_secret",
///     "sort": true,
///     "format_cells": true
///   },
///   "hide_columns_qv": true,
///   "hide_language_sheet": false,
///   "workbook_password": "wb_secret",
///   "row_grouping": {
///     "groups": [{ "start_row": 10, "end_row": 20, "collapsed": false }],
///     "symbols_above": false
///   }
/// }
/// ```
#[derive(Debug, Clone, PartialEq, Default)]
#[cfg_attr(feature = "serde", derive(serde::Serialize, serde::Deserialize))]
#[cfg_attr(feature = "serde", serde(deny_unknown_fields))]
pub struct ReportOptions {
    /// Sheet protection settings (`None` = no protection).
    /// Use [`SheetProtection`] to configure which operations are allowed/blocked.
    /// The `password` field inside [`SheetProtection`] sets the sheet password.
    #[cfg_attr(feature = "serde", serde(default))]
    pub sheet_protection: Option<SheetProtection>,
    /// Hide helper columns Q-V
    pub hide_columns_qv: bool,
    /// Hide language sheet
    pub hide_language_sheet: bool,
    /// Workbook protection password (`None` = no workbook protection)
    pub workbook_password: Option<String>,
    /// Row grouping for collapsible sections
    #[cfg_attr(feature = "serde", serde(default))]
    pub row_grouping: Option<RowGrouping>,
}

impl ReportOptions {
    pub fn builder() -> ReportOptionsBuilder {
        ReportOptionsBuilder::default()
    }
}

/// Builder for [`ReportOptions`].
#[derive(Default)]
pub struct ReportOptionsBuilder {
    sheet_protection: Option<SheetProtection>,
    hide_columns_qv: bool,
    hide_language_sheet: bool,
    workbook_password: Option<String>,
    row_grouping: Option<RowGrouping>,
}

impl ReportOptionsBuilder {
    pub fn sheet_protection(mut self, prot: SheetProtection) -> Self {
        self.sheet_protection = Some(prot);
        self
    }
    pub fn workbook_password(mut self, v: impl Into<String>) -> Self {
        self.workbook_password = Some(v.into());
        self
    }
    pub fn hide_columns_qv(mut self, v: bool) -> Self {
        self.hide_columns_qv = v;
        self
    }
    pub fn hide_language_sheet(mut self, v: bool) -> Self {
        self.hide_language_sheet = v;
        self
    }
    pub fn row_grouping(mut self, v: RowGrouping) -> Self {
        self.row_grouping = Some(v);
        self
    }
    pub fn build(self) -> ReportOptions {
        ReportOptions {
            sheet_protection: self.sheet_protection,
            hide_columns_qv: self.hide_columns_qv,
            hide_language_sheet: self.hide_language_sheet,
            workbook_password: self.workbook_password,
            row_grouping: self.row_grouping,
        }
    }
}

// ============================================================================
// ReportConfig — Top-level compositor
// ============================================================================

/// Unified configuration for a financial report.
///
/// Composes four domain sub-structs: [`ReportHeader`], [`ReportBody`],
/// [`ReportFooter`], and [`ReportOptions`].
///
/// ## Builder Example
///
/// ```ignore
/// use fb_generator::*;
///
/// let config = ReportConfig::builder()
///     .header(
///         ReportHeader::builder()
///             .language(Language::Deutsch)
///             .currency(Currency::EUR)
///             .project_number("PROJ-001")
///             .build()
///     )
///     .body(
///         ReportBody::builder()
///             .add_position(1, PositionEntry::builder().approved(5000.0).build())
///             .set_header_input(6, PositionEntry::builder().approved(1000.0).build())
///             .build()
///     )
///     .footer(ReportFooter::builder().bank(8500.0).build())
///     .options(ReportOptions::builder().sheet_protection(SheetProtection::from_defaults().with_password("geheim")).build())
///     .build();
/// config.write_to("report.xlsx")?;
/// ```
///
/// ## JSON (minimal)
///
/// ```json
/// {}
/// ```
///
/// All sub-structs default when omitted.
///
/// ## Errors
///
/// See [`ReportError`](crate::ReportError) for error variants returned by
/// [`write_to`](Self::write_to) and [`write_to_precomputed`](Self::write_to_precomputed).
#[derive(Debug, Clone, PartialEq, Default)]
#[cfg_attr(feature = "serde", derive(serde::Serialize, serde::Deserialize))]
#[cfg_attr(feature = "serde", serde(deny_unknown_fields))]
pub struct ReportConfig {
    /// Header metadata (language, currency, project info, dates)
    #[cfg_attr(feature = "serde", serde(default))]
    pub header: ReportHeader,
    /// Body data (income table, panels, cost positions)
    #[cfg_attr(feature = "serde", serde(default))]
    pub body: ReportBody,
    /// Footer balances (bank, kasse, sonstiges)
    #[cfg_attr(feature = "serde", serde(default))]
    pub footer: ReportFooter,
    /// Output options (protection, visibility, grouping)
    #[cfg_attr(feature = "serde", serde(default))]
    pub options: ReportOptions,
}

impl ReportConfig {
    pub fn builder() -> ReportConfigBuilder {
        ReportConfigBuilder::default()
    }
}

/// Builder for [`ReportConfig`].
#[derive(Default)]
pub struct ReportConfigBuilder {
    header: Option<ReportHeader>,
    body: Option<ReportBody>,
    footer: Option<ReportFooter>,
    options: Option<ReportOptions>,
}

impl ReportConfigBuilder {
    pub fn header(mut self, v: ReportHeader) -> Self {
        self.header = Some(v);
        self
    }
    pub fn body(mut self, v: ReportBody) -> Self {
        self.body = Some(v);
        self
    }
    pub fn footer(mut self, v: ReportFooter) -> Self {
        self.footer = Some(v);
        self
    }
    pub fn options(mut self, v: ReportOptions) -> Self {
        self.options = Some(v);
        self
    }
    pub fn build(self) -> ReportConfig {
        ReportConfig {
            header: self.header.unwrap_or_default(),
            body: self.body.unwrap_or_default(),
            footer: self.footer.unwrap_or_default(),
            options: self.options.unwrap_or_default(),
        }
    }
}

// ============================================================================
// ReportConfig methods
// ============================================================================

impl ReportConfig {
    /// Writes the financial report to the given file path.
    ///
    /// # Errors
    ///
    /// Returns [`ReportError`](crate::ReportError):
    /// - [`Xlsx`](crate::ReportError::Xlsx) — Excel write failure
    /// - [`Protection`](crate::ReportError::Protection) — workbook password injection failed
    /// - [`Io`](crate::ReportError::Io) — file system error
    /// - [`InvalidPath`](crate::ReportError::InvalidPath) — path contains non-UTF-8
    pub fn write_to(&self, output_path: impl AsRef<Path>) -> Result<(), crate::error::ReportError> {
        self.validate()?;
        let values = self.build_values();
        let body_config = self.build_body_config();
        let sheet_opts = self.build_sheet_options();
        let wb_prot = self
            .options
            .workbook_password
            .as_ref()
            .map(WorkbookProtection::new);
        create_protected_report(
            output_path,
            &values,
            &body_config,
            &sheet_opts,
            wb_prot.as_ref(),
            self.options.hide_language_sheet,
        )?;
        Ok(())
    }

    /// Batch-optimized: precompute the hash once, reuse for N files.
    ///
    /// Saves ~25ms SHA-512 overhead per file when using the same password.
    /// The `workbook_password` field is ignored — the precomputed hash is used directly.
    ///
    /// ```ignore
    /// use fb_generator::{precompute_hash, ReportConfig};
    ///
    /// let hash = precompute_hash("password");
    /// for config in &configs {
    ///     config.write_to_precomputed("output.xlsx", &hash).unwrap();
    /// }
    /// ```
    ///
    /// # Errors
    ///
    /// Same as [`write_to`](Self::write_to). See [`ReportError`](crate::ReportError).
    pub fn write_to_precomputed(
        &self,
        output_path: impl AsRef<Path>,
        hash: &crate::workbook_protection::PrecomputedHash,
    ) -> Result<(), crate::error::ReportError> {
        self.validate()?;
        let values = self.build_values();
        let body_config = self.build_body_config();
        let sheet_opts = self.build_sheet_options();
        create_protected_report_precomputed(
            output_path,
            &values,
            &body_config,
            &sheet_opts,
            self.options.hide_language_sheet,
            hash,
        )?;
        Ok(())
    }

    fn validate(&self) -> Result<(), crate::error::ReportError> {
        // Income table overflow is impossible at the type level (IncomeTable has exactly 5 fields).
        if self.body.left_panel.len() > 18 {
            return Err(crate::error::ReportError::Validation(format!(
                "left_panel has {} entries (max 18)",
                self.body.left_panel.len()
            )));
        }
        if self.body.right_panel.len() > 18 {
            return Err(crate::error::ReportError::Validation(format!(
                "right_panel has {} entries (max 18)",
                self.body.right_panel.len()
            )));
        }

        for &cat in self
            .body
            .positions
            .keys()
            .chain(self.body.header_inputs.keys())
        {
            if !(1..=8).contains(&cat) {
                return Err(crate::error::ReportError::Validation(format!(
                    "position category {} out of range (must be 1-8)",
                    cat
                )));
            }
        }

        for &cat in self.body.positions.keys() {
            if self.body.header_inputs.contains_key(&cat) {
                return Err(crate::error::ReportError::Validation(format!(
                    "category {cat} is used in both positions and header_inputs — use one or the other",
                )));
            }
        }

        Ok(())
    }

    fn build_values(&self) -> ReportValues {
        let h = &self.header;
        let mut v = ReportValues::new().with_lang(h.language).with_curr(h.currency);

        if let Some(ref s) = h.project_number {
            v = v.with_project_number(s);
        }
        if let Some(ref s) = h.project_title {
            v = v.with_project_title(s);
        }
        if let Some(ref s) = h.project_start {
            v = v.with_project_start(s);
        }
        if let Some(ref s) = h.project_end {
            v = v.with_project_end(s);
        }
        if let Some(ref s) = h.report_start {
            v = v.with_report_start(s);
        }
        if let Some(ref s) = h.report_end {
            v = v.with_report_end(s);
        }
        if let Some(ref s) = h.version {
            v = v.with_version(s);
        }

        // Table (rows 15-19): fixed order by named field
        let table_rows: [(u8, &Option<TableEntry>); 5] = [
            (0, &self.body.table.saldovortrag),
            (1, &self.body.table.eigenmittel),
            (2, &self.body.table.drittmittel),
            (3, &self.body.table.kmw_mittel),
            (4, &self.body.table.zinsertraege),
        ];
        for (idx, opt) in &table_rows {
            if let Some(e) = opt {
                if let Some(n) = e.approved_budget {
                    v.set(ApiKey::ApprovedBudget(*idx), n);
                }
                if let Some(n) = e.income_report {
                    v.set(ApiKey::IncomeReportPeriod(*idx), n);
                }
                if let Some(n) = e.income_total {
                    v.set(ApiKey::IncomeTotal(*idx), n);
                }
                if let Some(ref s) = e.reason {
                    v.set(ApiKey::IncomeReason(*idx), s.as_str());
                }
            }
            // None field = empty row: no value written, ApiKey slot stays empty
        }

        // Left panel
        for (i, opt) in self.body.left_panel.iter().enumerate() {
            let idx = i as u8;
            if let Some(e) = opt {
                if let Some(ref s) = e.date {
                    v.set(ApiKey::LeftDate(idx), CellValue::Date(s.clone()));
                }
                if let Some(n) = e.amount_euro {
                    v.set(ApiKey::LeftAmountEuro(idx), n);
                }
                if let Some(n) = e.amount_local {
                    v.set(ApiKey::LeftAmountLocal(idx), n);
                }
            }
        }

        // Right panel
        for (i, opt) in self.body.right_panel.iter().enumerate() {
            let idx = i as u8;
            if let Some(e) = opt {
                if let Some(ref s) = e.date {
                    v.set(ApiKey::RightDate(idx), CellValue::Date(s.clone()));
                }
                if let Some(n) = e.amount_euro {
                    v.set(ApiKey::RightAmountEuro(idx), n);
                }
                if let Some(n) = e.amount_local {
                    v.set(ApiKey::RightAmountLocal(idx), n);
                }
            }
        }

        // Normal positions: position = Vec index + 1 (1-based)
        for (&cat, rows) in &self.body.positions {
            for (i, opt) in rows.iter().enumerate() {
                let pos = (i + 1) as u16;
                if let Some(e) = opt {
                    v.set_position_row(
                        cat,
                        pos,
                        opt_str(e.description.clone()),
                        opt_num(e.approved),
                        opt_num(e.income_report),
                        opt_num(e.income_total),
                        opt_str(e.remark.clone()),
                    );
                }
                // None = empty row: position slot stays empty
            }
        }

        // Header-input entries
        for (&cat, opt) in &self.body.header_inputs {
            if let Some(e) = opt {
                v.set_header_input(
                    cat,
                    opt_num(e.approved),
                    opt_num(e.income_report),
                    opt_num(e.income_total),
                    opt_str(e.remark.clone()),
                );
            }
        }

        // Footer balances
        if let Some(n) = self.footer.bank {
            v.set_footer_bank(n);
        }
        if let Some(n) = self.footer.kasse {
            v.set_footer_kasse(n);
        }
        if let Some(n) = self.footer.sonstiges {
            v.set_footer_sonstiges(n);
        }

        v
    }

    fn build_body_config(&self) -> BodyConfig {
        let mut config = BodyConfig::new();
        for (&cat, rows) in &self.body.positions {
            config = config.with_positions(cat, rows.len() as u16);
        }
        for &cat in self.body.header_inputs.keys() {
            config = config.with_positions(cat, 0);
        }
        config
    }

    fn build_sheet_options(&self) -> SheetOptions {
        let mut opts = if let Some(prot) = self.options.sheet_protection.clone() {
            SheetOptions::new().with_protection(prot)
        } else {
            SheetOptions::new()
        };
        if self.options.hide_columns_qv {
            opts = opts.with_hidden_columns_qv();
        }
        if let Some(ref rg) = self.options.row_grouping {
            opts = opts.with_row_grouping(rg.clone());
        }
        opts
    }
}

// ============================================================================
// Helper functions
// ============================================================================

fn opt_num(v: Option<f64>) -> CellValue {
    match v {
        Some(n) => CellValue::Number(n),
        None => CellValue::Empty,
    }
}

fn opt_str(v: Option<String>) -> CellValue {
    match v {
        Some(s) if !s.is_empty() => CellValue::Text(s),
        _ => CellValue::Empty,
    }
}

// ============================================================================
// Tests
// ============================================================================

#[cfg(test)]
mod tests {
    use super::*;

    fn valid_config() -> ReportConfig {
        ReportConfig {
            body: ReportBody {
                table: IncomeTable {
                    kmw_mittel: Some(TableEntry::default()),
                    ..IncomeTable::default()
                },
                positions: [(1u8, vec![Some(PositionEntry::default())])].into_iter().collect(),
                header_inputs: [(6u8, Some(PositionEntry::default()))].into_iter().collect(),
                ..ReportBody::default()
            },
            ..ReportConfig::default()
        }
    }

    #[test]
    fn test_validate_valid_config() {
        assert!(valid_config().validate().is_ok());
    }

    #[test]
    fn test_validate_default_config() {
        assert!(ReportConfig::default().validate().is_ok());
    }

    #[test]
    fn test_validate_left_panel_too_many_entries() {
        let config = ReportConfig {
            body: ReportBody {
                left_panel: vec![None; 19],
                ..ReportBody::default()
            },
            ..ReportConfig::default()
        };
        let err = config.validate().unwrap_err();
        assert!(err.to_string().contains("left_panel has 19 entries"));
    }

    #[test]
    fn test_validate_right_panel_too_many_entries() {
        let config = ReportConfig {
            body: ReportBody {
                right_panel: vec![None; 20],
                ..ReportBody::default()
            },
            ..ReportConfig::default()
        };
        let err = config.validate().unwrap_err();
        assert!(err.to_string().contains("right_panel has 20 entries"));
    }

    #[test]
    fn test_validate_position_category_out_of_range() {
        let config = ReportConfig {
            body: ReportBody {
                positions: [(9u8, vec![Some(PositionEntry::default())])].into_iter().collect(),
                ..ReportBody::default()
            },
            ..ReportConfig::default()
        };
        let err = config.validate().unwrap_err();
        assert!(err.to_string().contains("category 9 out of range"));
    }

    #[test]
    fn test_validate_header_input_category_out_of_range() {
        let config = ReportConfig {
            body: ReportBody {
                header_inputs: [(0u8, None)].into_iter().collect(),
                ..ReportBody::default()
            },
            ..ReportConfig::default()
        };
        let err = config.validate().unwrap_err();
        assert!(err.to_string().contains("category 0 out of range"));
    }

    #[test]
    fn test_builder_infallible_all_defaults() {
        // No ? required anywhere
        let config = ReportConfig::builder().build();
        assert!(config.validate().is_ok());
    }

    #[test]
    fn test_builder_header_str_ergonomics() {
        let h = ReportHeader::builder()
            .project_number("PROJ-001")      // &str, no .to_string()
            .project_title("Test Project")
            .build();
        assert_eq!(h.project_number, Some("PROJ-001".to_string()));
        assert_eq!(h.project_title, Some("Test Project".to_string()));
    }

    #[test]
    fn test_builder_header_defaults() {
        let h = ReportHeader::builder().build();
        assert_eq!(h.language, Language::Deutsch);
        assert_eq!(h.currency, Currency::EUR);
        assert_eq!(h.project_number, None);
    }

    #[test]
    fn test_body_builder_named_table_rows() {
        let body = ReportBody::builder()
            .kmw_mittel(TableEntry::builder().approved_budget(80_000.0).income_report(50_000.0).build())
            .eigenmittel(TableEntry::builder().approved_budget(15_000.0).build())
            .build();
        // Named rows set
        assert!(body.table.kmw_mittel.is_some());
        assert!(body.table.eigenmittel.is_some());
        // Unset rows default to None (no overflow, no accidental data)
        assert!(body.table.saldovortrag.is_none());
        assert!(body.table.drittmittel.is_none());
        assert!(body.table.zinsertraege.is_none());
        assert_eq!(body.table.kmw_mittel.as_ref().unwrap().approved_budget, Some(80_000.0));
    }

    #[test]
    fn test_body_builder_named_table_no_duplicates() {
        // Setting the same row twice just overwrites (last call wins)
        let body = ReportBody::builder()
            .saldovortrag(TableEntry::builder().approved_budget(1_000.0).build())
            .saldovortrag(TableEntry::builder().approved_budget(9_999.0).build())
            .build();
        assert_eq!(body.table.saldovortrag.as_ref().unwrap().approved_budget, Some(9_999.0));
    }

    #[test]
    fn test_body_builder_panel_skip() {
        let body = ReportBody::builder()
            .add_left_panel_entry(PanelEntry::builder().amount_euro(100.0).build())
            .skip_left_panel_row()
            .add_left_panel_entry(PanelEntry::builder().amount_euro(200.0).build())
            .build();
        assert_eq!(body.left_panel.len(), 3);
        assert!(body.left_panel[1].is_none());
    }

    #[test]
    fn test_body_builder_positions_no_index_or_category() {
        let body = ReportBody::builder()
            .add_position(1, PositionEntry::builder().approved(500.0).build())
            .skip_position_row(1)
            .add_position(1, PositionEntry::builder().approved(600.0).build())
            .set_header_input(6, PositionEntry::builder().approved(100.0).build())
            .build();
        assert_eq!(body.positions[&1].len(), 3);
        assert!(body.positions[&1][1].is_none()); // skipped
        assert!(body.header_inputs[&6].is_some());
    }

    #[test]
    fn test_body_builder_batch_inserts() {
        let entries = vec![
            PanelEntry::builder().amount_euro(1.0).build(),
            PanelEntry::builder().amount_euro(2.0).build(),
            PanelEntry::builder().amount_euro(3.0).build(),
        ];
        let body = ReportBody::builder()
            .add_left_panel_entries(entries)
            .build();
        assert_eq!(body.left_panel.len(), 3);
    }

    #[test]
    fn test_build_body_config_derives_positions() {
        let config = ReportConfig {
            body: ReportBody {
                positions: [
                    (1u8, vec![Some(PositionEntry::default()); 5]),
                    (2u8, vec![Some(PositionEntry::default()); 3]),
                ]
                .into_iter()
                .collect(),
                header_inputs: [(6u8, Some(PositionEntry::default()))].into_iter().collect(),
                ..ReportBody::default()
            },
            ..ReportConfig::default()
        };
        let bc = config.build_body_config();
        assert_eq!(bc.position_count(1), 5);
        assert_eq!(bc.position_count(2), 3);
        assert_eq!(bc.position_count(6), 0); // header-input
    }

    #[test]
    fn test_build_body_config_skip_rows_count() {
        // Skipped rows (None) count toward the row count for this category
        let config = ReportConfig {
            body: ReportBody {
                positions: [(1u8, vec![Some(PositionEntry::default()), None, Some(PositionEntry::default())])]
                    .into_iter()
                    .collect(),
                ..ReportBody::default()
            },
            ..ReportConfig::default()
        };
        let bc = config.build_body_config();
        assert_eq!(bc.position_count(1), 3); // 2 real + 1 skip
    }

    #[test]
    fn test_typed_language_in_header() {
        let config = ReportConfig {
            header: ReportHeader {
                language: Language::English,
                ..ReportHeader::default()
            },
            ..ReportConfig::default()
        };
        assert!(config.validate().is_ok());
    }

    #[test]
    fn test_typed_currency_in_header() {
        let config = ReportConfig {
            header: ReportHeader {
                currency: Currency::USD,
                ..ReportHeader::default()
            },
            ..ReportConfig::default()
        };
        assert!(config.validate().is_ok());
    }

    #[test]
    fn test_position_entry_builder_no_category_or_position() {
        let e = PositionEntry::builder()
            .description("Test")
            .approved(1000.0)
            .income_report(500.0)
            .income_total(500.0)
            .remark("Bemerkung")
            .build();
        assert_eq!(e.description, Some("Test".to_string()));
        assert_eq!(e.approved, Some(1000.0));
    }

    #[test]
    fn test_table_entry_builder() {
        let e = TableEntry::builder()
            .approved_budget(10_000.0)
            .income_report(5_000.0)
            .income_total(5_000.0)
            .reason("Spende")
            .build();
        assert_eq!(e.approved_budget, Some(10_000.0));
        assert_eq!(e.reason, Some("Spende".to_string()));
    }

    #[test]
    fn test_panel_entry_builder_no_index() {
        let e = PanelEntry::builder()
            .date("15.01.2024")
            .amount_euro(9_000.0)
            .amount_local(8_500.0)
            .build();
        assert_eq!(e.date, Some("15.01.2024".to_string()));
        assert_eq!(e.amount_euro, Some(9_000.0));
    }

    #[test]
    fn test_full_builder_chain_no_question_mark() {
        // This test verifies the entire chain compiles without any ?
        let config = ReportConfig::builder()
            .header(
                ReportHeader::builder()
                    .language(Language::English)
                    .currency(Currency::USD)
                    .project_number("TEST-001")
                    .build(),
            )
            .body(
                ReportBody::builder()
                    .kmw_mittel(TableEntry::builder().approved_budget(5_000.0).build())
                    .add_left_panel_entry(
                        PanelEntry::builder().date("01.01.2024").amount_euro(1_000.0).build(),
                    )
                    .add_position(
                        1,
                        PositionEntry::builder()
                            .description("Kosten")
                            .approved(2_000.0)
                            .build(),
                    )
                    .set_header_input(
                        6,
                        PositionEntry::builder().approved(500.0).build(),
                    )
                    .build(),
            )
            .footer(ReportFooter::builder().bank(8_000.0).build())
            .options(
                ReportOptions::builder()
                    .sheet_protection(SheetProtection::from_defaults().with_password("secret"))
                    .hide_columns_qv(true)
                    .build(),
            )
            .build();
        assert!(config.validate().is_ok());
    }

    #[test]
    fn test_categories_6_7_8_multi_row_via_add_position() {
        let body = ReportBody::builder()
            .add_position(6, PositionEntry::builder().description("Cat6 Row1").approved(1_000.0).build())
            .add_position(6, PositionEntry::builder().description("Cat6 Row2").approved(2_000.0).build())
            .add_position(7, PositionEntry::builder().description("Cat7 Row1").approved(500.0).build())
            .add_position(8, PositionEntry::builder().description("Cat8 Row1").approved(300.0).build())
            .add_position(8, PositionEntry::builder().description("Cat8 Row2").approved(400.0).build())
            .add_position(8, PositionEntry::builder().description("Cat8 Row3").approved(500.0).build())
            .build();
        assert_eq!(body.positions[&6].len(), 2);
        assert_eq!(body.positions[&7].len(), 1);
        assert_eq!(body.positions[&8].len(), 3);
        assert!(body.header_inputs.is_empty());
    }

    #[test]
    fn test_build_body_config_categories_6_to_8_multi_row() {
        let config = ReportConfig::builder()
            .body(
                ReportBody::builder()
                    .add_position(6, PositionEntry::builder().approved(1_000.0).build())
                    .add_position(6, PositionEntry::builder().approved(2_000.0).build())
                    .add_position(7, PositionEntry::builder().approved(500.0).build())
                    .add_position(8, PositionEntry::builder().approved(300.0).build())
                    .build(),
            )
            .build();
        assert!(config.validate().is_ok());
        let bc = config.build_body_config();
        assert_eq!(bc.position_count(6), 2);
        assert_eq!(bc.position_count(7), 1);
        assert_eq!(bc.position_count(8), 1);
    }

    #[test]
    fn test_categories_6_7_8_header_input() {
        let body = ReportBody::builder()
            .set_header_input(6, PositionEntry::builder().approved(3_000.0).build())
            .set_header_input(7, PositionEntry::builder().approved(1_500.0).remark("Overhead").build())
            .set_header_input(8, PositionEntry::builder().approved(800.0).build())
            .build();
        assert!(body.header_inputs[&6].is_some());
        assert_eq!(body.header_inputs[&6].as_ref().unwrap().approved, Some(3_000.0));
        assert!(body.header_inputs[&7].is_some());
        assert_eq!(body.header_inputs[&7].as_ref().unwrap().remark, Some("Overhead".to_string()));
        assert!(body.header_inputs[&8].is_some());
        assert!(body.positions.is_empty());
    }

    #[test]
    fn test_mixed_positions_categories_1_through_8() {
        let config = ReportConfig::builder()
            .body(
                ReportBody::builder()
                    .add_position(1, PositionEntry::builder().description("Personal").approved(5_000.0).build())
                    .add_position(2, PositionEntry::builder().description("Reisen").approved(1_000.0).build())
                    .add_position(6, PositionEntry::builder().description("Overhead").approved(500.0).build())
                    .add_position(7, PositionEntry::builder().description("Reserve").approved(200.0).build())
                    .set_header_input(8, PositionEntry::builder().approved(100.0).build())
                    .build(),
            )
            .build();
        assert!(config.validate().is_ok());
        let bc = config.build_body_config();
        assert_eq!(bc.position_count(1), 1);
        assert_eq!(bc.position_count(2), 1);
        assert_eq!(bc.position_count(6), 1);
        assert_eq!(bc.position_count(7), 1);
        assert_eq!(bc.position_count(8), 0); // header-input → count 0
    }

    #[test]
    fn test_validate_rejects_category_in_both_positions_and_header_inputs() {
        let config = ReportConfig {
            body: ReportBody {
                positions: [(6u8, vec![Some(PositionEntry::default())])].into_iter().collect(),
                header_inputs: [(6u8, Some(PositionEntry::default()))].into_iter().collect(),
                ..ReportBody::default()
            },
            ..ReportConfig::default()
        };
        let err = config.validate().unwrap_err();
        assert!(err.to_string().contains("category 6"));
    }

    #[test]
    fn test_skip_position_row_categories_6_to_8() {
        let body = ReportBody::builder()
            .add_position(6, PositionEntry::builder().approved(1_000.0).build())
            .skip_position_row(6)
            .add_position(6, PositionEntry::builder().approved(2_000.0).build())
            .add_position(7, PositionEntry::builder().approved(500.0).build())
            .skip_position_row(8)
            .add_position(8, PositionEntry::builder().approved(300.0).build())
            .build();
        assert_eq!(body.positions[&6].len(), 3);
        assert!(body.positions[&6][1].is_none()); // skipped middle row
        assert_eq!(body.positions[&7].len(), 1);
        assert_eq!(body.positions[&8].len(), 2);
        assert!(body.positions[&8][0].is_none()); // first row is the skip
    }

    #[test]
    fn test_version_builder_setter() {
        let h = ReportHeader::builder().version("v2025-1").build();
        assert_eq!(h.version, Some("v2025-1".to_string()));
    }

    #[test]
    fn test_version_none_by_default() {
        let h = ReportHeader::builder().build();
        assert_eq!(h.version, None);
    }
}
