//! Report writer with IronCalc-based formula evaluation.
//!
//! Main API: [`write_report_with_options()`].

use super::layout::{self, setup_sheet};
use super::structure::write_structure;
use crate::report::api::{CellValue, ReportDate, ReportValues};
use crate::report::body::config::BODY_START_ROW;
use crate::report::body::{BodyConfig, BodyLayout, CategoryMode, FooterLayout};
use crate::report::calc::{CalcBridge, ModelTemplate};
use crate::report::core::CellAddr;
use crate::report::options::SheetOptions;
use crate::report::styles::{
    build_format_matrix, extend_format_matrix_with_body, extend_format_matrix_with_footer,
    extend_format_matrix_with_prebody, FormatMatrix, ReportStyles, SectionStyles,
};
use rust_xlsxwriter::{
    ExcelDateTime, Format, Formula, Image, ObjectMovement, Workbook, Worksheet, XlsxError,
};
use std::path::Path;
use std::sync::LazyLock;

/// Global master template: created once, reused for all reports.
///
/// Contains all static formulas + language sheet data.
/// Initialized on first access (`LazyLock`).
static MASTER_TEMPLATE: LazyLock<ModelTemplate> = LazyLock::new(ModelTemplate::new);

/// Kindermissionswerk logo, embedded in the binary.
const LOGO_BYTES: &[u8] =
    include_bytes!("../../../assets/Logo_transparent.png");

/// Result of the body generation step.
#[derive(Debug, Clone)]
pub struct BodyResult {
    /// The computed body layout
    pub layout: BodyLayout,
    /// Last written row
    pub last_row: u32,
    /// Row of the grand total
    pub total_row: u32,
    /// Col E total value (used by footer check formula)
    pub e_total: Option<f64>,
    /// Col F total value (used by footer check formula)
    pub f_total: Option<f64>,
}

/// Internal: writes the report including the dynamic body area.
fn write_report_with_body(
    ws: &mut Worksheet,
    suffix: &str,
    values: &ReportValues,
    body_config: &BodyConfig,
) -> Result<BodyResult, XlsxError> {
    let styles = ReportStyles::new();
    // Set default formula result to "" (instead of 0)
    ws.set_formula_result_default("");

    // 1. Create CalcBridge from template (contains static formulas + language sheet)
    let mut bridge = CalcBridge::from_template(&MASTER_TEMPLATE);

    // 2. Compute layouts
    let body_layout = BodyLayout::compute(body_config);
    let footer_layout = FooterLayout::compute(body_layout.total_row);
    let income_row = 19u32;

    // 3. Register dynamic formulas (body + footer)
    bridge.register_body_formulas(&body_layout);
    bridge.register_footer_formulas(&footer_layout, income_row);

    // 4. Populate input values + evaluate
    bridge.populate(values, &body_layout, &footer_layout);
    bridge.evaluate();

    // 5. Build complete FormatMatrix: static (header + prebody), then dynamic (body + footer)
    let sec = SectionStyles::new(&styles);
    let mut fmt = build_format_matrix(&styles, &sec);
    extend_format_matrix_with_prebody(&mut fmt, &sec);
    extend_format_matrix_with_body(&mut fmt, &styles, &body_layout);
    extend_format_matrix_with_footer(&mut fmt, &styles, &sec, footer_layout.start_row);

    // 6. Write complete structure (merges, blanks, static strings)
    let lang = values.language().unwrap_or("");
    write_structure(ws, &fmt, &body_layout, &footer_layout, suffix, lang)?;

    // 7. Write ALL cells from bridge (formulas + input values)
    write_cells_from_bridge(ws, &bridge, &fmt)?;

    // 8. Write HYPERLINK formula at J4 with evaluated VLOOKUP result as cache
    //    IronCalc evaluates the inner VLOOKUP → URL string; we wrap it in HYPERLINK()
    let hyperlink_cache = cell_value_to_string(&bridge.get_value(3, 9));
    let hyperlink_formula = Formula::new(
        r#"=HYPERLINK(VLOOKUP($E$2,Sprachversionen!$B:$BN,62,FALSE))"#,
    )
    .set_result(&hyperlink_cache);
    if let Some(format) = fmt.get_locked(3, 9) {
        ws.write_formula_with_format(3, 9, hyperlink_formula, format)?;
    } else {
        ws.write_formula_with_format(3, 9, hyperlink_formula, &Format::new().set_locked())?;
    }

    // 9. Insert logo in header area (mid-F to mid-H, starting at mid-row-0)
    insert_logo(ws)?;

    // 10. Freeze pane
    layout::setup_freeze_panes(ws, 9)?;

    // BodyResult
    let last_row = body_layout.last_row;
    let total_row = body_layout.total_row;
    let e_total = bridge.get_value(total_row, 4).as_number();
    let f_total = bridge.get_value(total_row, 5).as_number();

    Ok(BodyResult {
        layout: body_layout,
        last_row,
        total_row,
        e_total,
        f_total,
    })
}

/// Writes the complete financial report with dynamic body area and options.
///
/// Styles are created internally with default settings.
pub fn write_report_with_options(
    ws: &mut Worksheet,
    suffix: &str,
    values: &ReportValues,
    body_config: &BodyConfig,
    options: &SheetOptions,
) -> Result<BodyResult, XlsxError> {
    // Write base report
    let body_result = write_report_with_body(ws, suffix, values, body_config)?;

    // Apply options
    apply_sheet_options(ws, options, &body_result)?;

    Ok(body_result)
}

/// Applies [`SheetOptions`] to a worksheet.
///
/// Can be called after writing the report.
pub fn apply_sheet_options(
    ws: &mut Worksheet,
    options: &SheetOptions,
    _body_result: &BodyResult,
) -> Result<(), XlsxError> {
    // 1. Hide columns and rows
    apply_hidden_ranges(ws, &options.hidden)?;

    // 2. Apply row grouping (if configured)
    apply_row_grouping(ws, &options.row_grouping)?;

    // 3. Apply sheet protection (if configured)
    if let Some(ref protection) = options.protection {
        let prot_options = protection.to_protection_options();

        if let Some(ref password) = protection.password {
            ws.protect_with_password(password);
        }
        ws.protect_with_options(&prot_options);
    }

    Ok(())
}

// ============================================================================
// High-Level API: create_protected_report
// ============================================================================

/// Extracts the version string for cell B2 from ReportValues.
///
/// Returns the version if set, otherwise an empty string.
fn extract_version_from_values(values: &ReportValues) -> String {
    values.version().unwrap_or("").to_string()
}

// ============================================================================
// Print setup: print area + page breaks
// ============================================================================

/// Estimated rows per printed page (A4 portrait, fit_to_pages width=1).
const MAX_ROWS_PER_PAGE: u32 = 80;

/// Determines the sheet name based on the language setting.
///
/// Returns the localized sheet name (e.g. "Finanzbericht", "Financial Report")
/// or "Sheet1" as fallback.
fn determine_sheet_name(values: &ReportValues) -> String {
    values
        .get(crate::report::api::ApiKey::Language)
        .as_text()
        .and_then(|lang_text| {
            crate::lang::LANG_CONFIG.get(lang_text).or_else(|| {
                crate::lang::LANG_CONFIG
                    .values()
                    .find(|c| c.lang_val.eq_ignore_ascii_case(lang_text))
            })
        })
        .map(|c| c.fb_sheet.to_string())
        .unwrap_or_else(|| "Sheet1".to_string())
}

/// Computes horizontal page breaks at category boundaries.
///
/// Places breaks BEFORE category headers once the current page is full.
/// The first page has less space because the header section (rows 0–25) is consumed.
///
/// **Oversized category:** If a single category together with the current page content
/// exceeds `MAX_ROWS_PER_PAGE` rows, additional hard breaks are inserted every
/// `MAX_ROWS_PER_PAGE` rows within the category.
/// Subsequent categories follow the normal rule (break before header).
fn compute_page_breaks(layout: &BodyLayout, footer_layout: &FooterLayout) -> Vec<u32> {
    let mut breaks = Vec::new();
    let mut rows_on_page = BODY_START_ROW;

    for cat in &layout.categories {
        let is_cat_8 = cat.meta.num == 8;

        let cat_start = match &cat.mode {
            CategoryMode::HeaderInput { row } => *row,
            CategoryMode::WithPositions { header_row, .. } => *header_row,
        };
        let mut block_end = cat.sum_row();
        if is_cat_8 {
            // Category 8 is always kept together with the total row
            block_end = layout.total_row;
        }

        let block_rows = block_end - cat_start + 1;

        if rows_on_page + block_rows <= MAX_ROWS_PER_PAGE {
            // Fits on the current page
            rows_on_page += block_rows;
        } else {
            // Does NOT fit on the current page.
            // Rule: A category is only pushed to the next page if the current page has
            // at least ~65 rows (2/3 of MAX_ROWS_PER_PAGE) used. Otherwise, fill the
            // page and break within the category.
            // Category 8 is an exception — always moved as a block to the next page
            // (or broken there if it alone exceeds the page).
            let min_rows_to_keep = 65;
            let should_fill_page =
                !is_cat_8 && rows_on_page < min_rows_to_keep && cat_start > BODY_START_ROW;

            if should_fill_page {
                // Fill the page → hard break within the category
                let page_remaining = MAX_ROWS_PER_PAGE.saturating_sub(rows_on_page);
                let mut pos = cat_start + page_remaining;
                while pos <= block_end {
                    breaks.push(pos);
                    pos += MAX_ROWS_PER_PAGE;
                }
                let last_break = breaks.last().copied().unwrap_or(cat_start);
                rows_on_page = block_end.saturating_sub(last_break) + 1;
            } else {
                // Break BEFORE the block, unless we are at the start
                if cat_start > BODY_START_ROW {
                    breaks.push(cat_start);
                    rows_on_page = 0;
                }

                // Check if the block fits on the empty page
                if rows_on_page + block_rows > MAX_ROWS_PER_PAGE {
                    // Block alone exceeds the page — must be hard-broken
                    let start_of_page = if rows_on_page == 0 {
                        cat_start
                    } else {
                        cat_start.saturating_sub(rows_on_page)
                    };
                    let mut pos = start_of_page + MAX_ROWS_PER_PAGE;
                    while pos <= block_end {
                        breaks.push(pos);
                        pos += MAX_ROWS_PER_PAGE;
                    }
                    let last_break = breaks.last().copied().unwrap_or(cat_start);
                    rows_on_page = block_end.saturating_sub(last_break) + 1;
                } else {
                    rows_on_page += block_rows;
                }
            }
        }
    }

    // Footer block
    let footer_start = footer_layout.start_row;
    let footer_rows = footer_layout.end_row - footer_layout.start_row + 1;

    if rows_on_page + footer_rows > MAX_ROWS_PER_PAGE {
        // Footer does not fit on this page → break before footer
        breaks.push(footer_start);
    }

    breaks
}

/// Applies print settings to the worksheet (A4, portrait, page breaks).
fn apply_print_setup(
    ws: &mut Worksheet,
    body_result: &BodyResult,
    footer_layout: &FooterLayout,
) -> Result<(), XlsxError> {
    ws.set_paper_size(9); // A4
    ws.set_print_fit_to_pages(1, 0); // width = 1 page, height unlimited

    let breaks = compute_page_breaks(&body_result.layout, footer_layout);
    if !breaks.is_empty() {
        ws.set_page_breaks(&breaks)?;
    }

    Ok(())
}

/// Sets the print area as a non-contiguous range via `define_name`.
///
/// Range 1: B1:H{footer_end} (main report — may span multiple pages, portrait)
/// Range 2: J1:V31 (receipt panel — always 1 page, portrait)
fn setup_print_area(
    workbook: &mut Workbook,
    sheet_name: &str,
    footer_end_row: u32,
) -> Result<(), XlsxError> {
    let footer_excel = footer_end_row + 1; // 0-indexed → 1-indexed
    let q = format!("'{}'", sheet_name);

    let name = format!("{}!_xlnm.Print_Area", q);
    let range = format!("{}!$B$1:$H${},{}!$J$1:$V$31", q, footer_excel, q);
    workbook.define_name(name, &range)?;
    Ok(())
}

// ============================================================================
// High-Level API: create_protected_report
// ============================================================================

/// Creates a complete financial report with optional workbook protection.
///
/// Styles are created internally. Workbook protection and language sheet visibility
/// are passed as separate parameters (not part of [`SheetOptions`]).
pub(crate) fn create_protected_report(
    output_path: impl AsRef<Path>,
    values: &ReportValues,
    body_config: &BodyConfig,
    options: &SheetOptions,
    wb_protection: Option<&crate::workbook_protection::WorkbookProtection>,
    hide_language_sheet: bool,
) -> Result<(), crate::error::ReportError> {
    let output_path = output_path.as_ref();
    let sheet_name = determine_sheet_name(values);

    let mut workbook = Workbook::new();
    let ws = workbook.add_worksheet();
    ws.set_name(&sheet_name)?;

    setup_sheet(ws)?;
    let suffix = extract_version_from_values(values);
    let body_result = write_report_with_options(ws, &suffix, values, body_config, options)?;
    let footer_layout = FooterLayout::compute(body_result.total_row);
    apply_print_setup(ws, &body_result, &footer_layout)?;

    setup_print_area(&mut workbook, &sheet_name, footer_layout.end_row)?;
    crate::lang::build_sheet_with_visibility(&mut workbook, hide_language_sheet)?;

    if let Some(wb_prot) = wb_protection {
        let tmp = tempfile::NamedTempFile::new_in(output_path.parent().unwrap_or(Path::new(".")))?;
        workbook.save(tmp.path())?;

        crate::workbook_protection::protect_workbook_with_spin_count(
            tmp.path().to_str().ok_or_else(|| {
                crate::error::ReportError::InvalidPath(format!("{:?}", tmp.path()))
            })?,
            output_path.to_str().ok_or_else(|| {
                crate::error::ReportError::InvalidPath(format!("{:?}", output_path))
            })?,
            &wb_prot.password,
            wb_prot.spin_count,
        )?;
        // tmp is automatically deleted on drop — not persisted,
        // since output_path already contains the protected file.
    } else {
        workbook.save(output_path)?;
    }

    Ok(())
}

/// Like [`create_protected_report`], but with a precomputed hash (no SHA-512 overhead).
///
/// For batch operations: compute the hash once with `precompute_hash()`,
/// then reuse it for N files.
pub(crate) fn create_protected_report_precomputed(
    output_path: impl AsRef<Path>,
    values: &ReportValues,
    body_config: &BodyConfig,
    options: &SheetOptions,
    hide_language_sheet: bool,
    hash: &crate::workbook_protection::PrecomputedHash,
) -> Result<(), crate::error::ReportError> {
    let output_path = output_path.as_ref();
    let sheet_name = determine_sheet_name(values);

    let mut workbook = Workbook::new();
    let ws = workbook.add_worksheet();
    ws.set_name(&sheet_name)?;

    setup_sheet(ws)?;
    let suffix = extract_version_from_values(values);
    let body_result = write_report_with_options(ws, &suffix, values, body_config, options)?;
    let footer_layout = FooterLayout::compute(body_result.total_row);
    apply_print_setup(ws, &body_result, &footer_layout)?;

    setup_print_area(&mut workbook, &sheet_name, footer_layout.end_row)?;
    crate::lang::build_sheet_with_visibility(&mut workbook, hide_language_sheet)?;

    let tmp = tempfile::NamedTempFile::new_in(output_path.parent().unwrap_or(Path::new(".")))?;
    workbook.save(tmp.path())?;

    crate::workbook_protection::protect_workbook_precomputed(
        tmp.path()
            .to_str()
            .ok_or_else(|| crate::error::ReportError::InvalidPath(format!("{:?}", tmp.path())))?,
        output_path
            .to_str()
            .ok_or_else(|| crate::error::ReportError::InvalidPath(format!("{:?}", output_path)))?,
        hash,
    )?;
    // tmp is automatically deleted on drop — not persisted,
    // since output_path already contains the protected file.
    Ok(())
}

/// Applies [`HiddenRanges`](crate::report::options::HiddenRanges) to a worksheet.
fn apply_hidden_ranges(
    ws: &mut Worksheet,
    hidden: &crate::report::options::HiddenRanges,
) -> Result<(), XlsxError> {
    // Hide columns
    for range in hidden.column_ranges() {
        for col in range.start..=range.end {
            ws.set_column_hidden(col as u16)?;
        }
    }

    // Hide rows
    for range in hidden.row_ranges() {
        for row in range.start..=range.end {
            ws.set_row_hidden(row)?;
        }
    }

    Ok(())
}

/// Applies row grouping to a worksheet.
fn apply_row_grouping(
    ws: &mut Worksheet,
    grouping: &crate::report::options::RowGrouping,
) -> Result<(), XlsxError> {
    if grouping.is_empty() {
        return Ok(());
    }

    // Set outline symbols position
    if grouping.symbols_above {
        ws.group_symbols_above(true);
    }

    // Apply groups
    for group in grouping.groups() {
        if group.collapsed {
            ws.group_rows_collapsed(group.start_row, group.end_row)?;
        } else {
            ws.group_rows(group.start_row, group.end_row)?;
        }
    }

    Ok(())
}

/// Inserts the Kindermissionswerk logo into the header area.
///
/// Positioned from 1/3 of column F to 2/3 of column H, starting at mid-row-0.
fn insert_logo(ws: &mut Worksheet) -> Result<(), XlsxError> {
    let logo = Image::new_from_buffer(LOGO_BYTES)?
        .set_scale_to_size(263, 500, true)
        .set_object_movement(ObjectMovement::MoveButDontSizeWithCells)
        .set_alt_text("Die Sternsinger - Kindermissionswerk");

    // Cell F1 (row 0, col 5), offset to 1/3 of column F / middle of row 0 (~20px high → 10px)
    ws.insert_image_with_offset(0, 5, &logo, 44, 10)?;
    Ok(())
}

/// Writes all cells from the [`CalcBridge`] (input values and formulas with cached results).
fn write_cells_from_bridge(
    ws: &mut Worksheet,
    bridge: &CalcBridge,
    fmt: &FormatMatrix,
) -> Result<(), XlsxError> {
    // 1. Write input values directly (no IronCalc readback — avoids auto-parsing)
    for (addr, value) in bridge.input_values() {
        write_cell_value(ws, *addr, value, fmt)?;
    }

    // 2. Write formula cells (with cached results)
    for addr in bridge.formula_cells() {
        // Skip J4 — VLOOKUP evaluated here, but written as HYPERLINK() formula by engine
        if addr.row == 3 && addr.col == 9 {
            continue;
        }
        if let Some(formula_str) = bridge.get_formula(addr.row, addr.col) {
            let result = bridge.get_value(addr.row, addr.col);

            // Create formula with result (cache for Excel)
            let formula = Formula::new(&formula_str).set_result(cell_value_to_string(&result));

            // Write with format (locked)
            if let Some(format) = fmt.get_locked(addr.row, addr.col) {
                ws.write_formula_with_format(addr.row, addr.col, formula, format)?;
            } else {
                let locked = Format::new().set_locked();
                ws.write_formula_with_format(addr.row, addr.col, formula, &locked)?;
            }
        }
    }



    Ok(())
}

/// Converts a [`CellValue`] to a string for `Formula::set_result`.
fn cell_value_to_string(value: &CellValue) -> String {
    match value {
        CellValue::Empty => String::new(),
        CellValue::Text(s) => s.clone(),
        CellValue::Number(n) => n.to_string(),
        CellValue::Date(d) => d.clone(),
    }
}

/// Writes a cell value with optional format from the [`FormatMatrix`].
fn write_cell_value(
    ws: &mut Worksheet,
    addr: CellAddr,
    value: &CellValue,
    fmt: &FormatMatrix,
) -> Result<(), XlsxError> {
    let format = fmt.get(addr.row, addr.col);

    match value {
        CellValue::Empty => {
            if let Some(f) = format {
                ws.write_blank(addr.row, addr.col, f)?;
            }
        }
        CellValue::Text(s) => {
            if let Some(f) = format {
                ws.write_string_with_format(addr.row, addr.col, s, f)?;
            } else {
                ws.write_string(addr.row, addr.col, s)?;
            }
        }
        CellValue::Number(n) => {
            if let Some(f) = format {
                ws.write_number_with_format(addr.row, addr.col, *n, f)?;
            } else {
                ws.write_number(addr.row, addr.col, *n)?;
            }
        }
        CellValue::Date(d) => {
            // Try to parse and write as a real Excel date
            if let Ok(parsed) = ReportDate::parse(d) {
                if let Ok(dt) =
                    ExcelDateTime::from_ymd(parsed.year(), parsed.month(), parsed.day())
                {
                    if let Some(f) = format {
                        ws.write_datetime_with_format(addr.row, addr.col, &dt, f)?;
                    } else {
                        let default_fmt = Format::new().set_num_format_index(14);
                        ws.write_datetime_with_format(addr.row, addr.col, &dt, &default_fmt)?;
                    }
                } else {
                    // ExcelDateTime construction failed — fall back to string
                    if let Some(f) = format {
                        ws.write_string_with_format(addr.row, addr.col, d, f)?;
                    } else {
                        ws.write_string(addr.row, addr.col, d)?;
                    }
                }
            } else {
                // Parse failed — fall back to string
                if let Some(f) = format {
                    ws.write_string_with_format(addr.row, addr.col, d, f)?;
                } else {
                    ws.write_string(addr.row, addr.col, d)?;
                }
            }
        }
    }

    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;
    use rust_xlsxwriter::Workbook;

    fn test_worksheet() -> (Workbook, FormatMatrix) {
        (Workbook::new(), FormatMatrix::new())
    }

    // ── Date: valid formats written as real ExcelDateTime ─────────────

    #[test]
    fn test_write_date_german_format() {
        let (mut wb, fmt) = test_worksheet();
        let ws = wb.add_worksheet();
        let addr = CellAddr { row: 0, col: 0 };
        let value = CellValue::Date("15.01.2024".to_string());
        write_cell_value(ws, addr, &value, &fmt).unwrap();
    }

    #[test]
    fn test_write_date_iso_format() {
        let (mut wb, fmt) = test_worksheet();
        let ws = wb.add_worksheet();
        let addr = CellAddr { row: 0, col: 0 };
        let value = CellValue::Date("2024-06-30".to_string());
        write_cell_value(ws, addr, &value, &fmt).unwrap();
    }

    #[test]
    fn test_write_date_slash_format() {
        let (mut wb, fmt) = test_worksheet();
        let ws = wb.add_worksheet();
        let addr = CellAddr { row: 0, col: 0 };
        let value = CellValue::Date("25/12/2024".to_string());
        write_cell_value(ws, addr, &value, &fmt).unwrap();
    }

    // ── Date: unparseable string falls back to text ──────────────────

    #[test]
    fn test_write_date_invalid_falls_back_to_string() {
        let (mut wb, fmt) = test_worksheet();
        let ws = wb.add_worksheet();
        let addr = CellAddr { row: 0, col: 0 };
        let value = CellValue::Date("not-a-date".to_string());
        write_cell_value(ws, addr, &value, &fmt).unwrap();
    }

    #[test]
    fn test_write_date_empty_string_falls_back() {
        let (mut wb, fmt) = test_worksheet();
        let ws = wb.add_worksheet();
        let addr = CellAddr { row: 0, col: 0 };
        let value = CellValue::Date(String::new());
        write_cell_value(ws, addr, &value, &fmt).unwrap();
    }

    // ── Date: with format from FormatMatrix ──────────────────────────

    #[test]
    fn test_write_date_with_format() {
        let (mut wb, mut fmt) = test_worksheet();
        let ws = wb.add_worksheet();
        let addr = CellAddr { row: 0, col: 0 };
        fmt.set(0, 0, &Format::new().set_num_format_index(14));
        let value = CellValue::Date("01.01.2024".to_string());
        write_cell_value(ws, addr, &value, &fmt).unwrap();
    }

    #[test]
    fn test_write_date_invalid_with_format() {
        let (mut wb, mut fmt) = test_worksheet();
        let ws = wb.add_worksheet();
        let addr = CellAddr { row: 0, col: 0 };
        fmt.set(0, 0, &Format::new().set_num_format_index(14));
        let value = CellValue::Date("foobar".to_string());
        write_cell_value(ws, addr, &value, &fmt).unwrap();
    }

    // ── Date: edge cases ─────────────────────────────────────────────

    #[test]
    fn test_write_date_leap_year() {
        let (mut wb, fmt) = test_worksheet();
        let ws = wb.add_worksheet();
        let addr = CellAddr { row: 0, col: 0 };
        let value = CellValue::Date("29.02.2024".to_string());
        write_cell_value(ws, addr, &value, &fmt).unwrap();
    }

    #[test]
    fn test_write_date_end_of_year() {
        let (mut wb, fmt) = test_worksheet();
        let ws = wb.add_worksheet();
        let addr = CellAddr { row: 0, col: 0 };
        let value = CellValue::Date("31.12.2026".to_string());
        write_cell_value(ws, addr, &value, &fmt).unwrap();
    }

    #[test]
    fn test_write_date_multiple_formats_in_one_sheet() {
        let (mut wb, fmt) = test_worksheet();
        let ws = wb.add_worksheet();
        let dates = ["01.01.2024", "2024-06-15", "31.12.2026", "not-a-date"];
        for (i, d) in dates.iter().enumerate() {
            let addr = CellAddr {
                row: i as u32,
                col: 0,
            };
            let value = CellValue::Date(d.to_string());
            write_cell_value(ws, addr, &value, &fmt).unwrap();
        }
    }
}
