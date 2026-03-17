//! Test: Date handling in financial reports.
//!
//! Demonstrates that date values are written as real Excel dates:
//! - Various input formats (German, ISO, slash) → all become native Excel dates
//! - Excel displays them according to the user's system locale
//! - Dates are sortable and usable in formulas
//! - Invalid strings fall back to plain text

use fb_generator::*;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    println!("Creating date test reports...\n");

    // ── Report 1: German date format (DD.MM.YYYY) ────────────────────
    let config_de = ReportConfig::builder()
        .header(
            ReportHeader::builder()
                .language(Language::Deutsch)
                .currency(Currency::EUR)
                .project_number("DATE-TEST-001")
                .project_title("Datumstest — Deutsches Format")
                .project_start("01.01.2024")
                .project_end("31.12.2026")
                .report_start("01.07.2024")
                .report_end("31.12.2024")
                .build(),
        )
        .body(
            ReportBody::builder()
                .kmw_mittel(
                    TableEntry::builder()
                        .approved_budget(10_000.0)
                        .income_report(5_000.0)
                        .income_total(5_000.0)
                        .build(),
                )
                // Panel dates in German format
                .add_left_panel_entry(
                    PanelEntry::builder()
                        .date("15.01.2024")
                        .amount_euro(1_000.0)
                        .amount_local(1_000.0)
                        .build(),
                )
                .add_left_panel_entry(
                    PanelEntry::builder()
                        .date("28.02.2024")
                        .amount_euro(2_000.0)
                        .amount_local(2_000.0)
                        .build(),
                )
                .add_left_panel_entry(
                    PanelEntry::builder()
                        .date("29.02.2024") // leap year
                        .amount_euro(500.0)
                        .amount_local(500.0)
                        .build(),
                )
                .add_right_panel_entry(
                    PanelEntry::builder()
                        .date("01.06.2024")
                        .amount_euro(3_000.0)
                        .amount_local(3_000.0)
                        .build(),
                )
                .add_right_panel_entry(
                    PanelEntry::builder()
                        .date("31.12.2024")
                        .amount_euro(1_500.0)
                        .amount_local(1_500.0)
                        .build(),
                )
                .add_position(
                    1,
                    PositionEntry::builder()
                        .description("Test position")
                        .approved(5_000.0)
                        .income_report(2_500.0)
                        .income_total(2_500.0)
                        .build(),
                )
                .set_header_input(
                    6,
                    PositionEntry::builder()
                        .approved(1_000.0)
                        .income_report(500.0)
                        .income_total(500.0)
                        .build(),
                )
                .build(),
        )
        .footer(
            ReportFooter::builder()
                .bank(3_000.0)
                .kasse(500.0)
                .build(),
        )
        .build();

    let path_de = "examples/output/test_dates_german.xlsx";
    config_de.write_to(path_de)?;
    println!("  [1] German dates (DD.MM.YYYY):  {}", path_de);

    // ── Report 2: ISO date format (YYYY-MM-DD) ──────────────────────
    let config_iso = ReportConfig::builder()
        .header(
            ReportHeader::builder()
                .language(Language::English)
                .currency(Currency::USD)
                .project_number("DATE-TEST-002")
                .project_title("Date Test — ISO Format")
                .project_start("2024-01-01")
                .project_end("2026-12-31")
                .report_start("2024-07-01")
                .report_end("2024-12-31")
                .build(),
        )
        .body(
            ReportBody::builder()
                .kmw_mittel(
                    TableEntry::builder()
                        .approved_budget(10_000.0)
                        .income_report(5_000.0)
                        .income_total(5_000.0)
                        .build(),
                )
                // Panel dates in ISO format
                .add_left_panel_entry(
                    PanelEntry::builder()
                        .date("2024-01-15")
                        .amount_euro(1_000.0)
                        .amount_local(1_000.0)
                        .build(),
                )
                .add_left_panel_entry(
                    PanelEntry::builder()
                        .date("2024-06-30")
                        .amount_euro(2_000.0)
                        .amount_local(2_000.0)
                        .build(),
                )
                .add_right_panel_entry(
                    PanelEntry::builder()
                        .date("2024-03-15")
                        .amount_euro(1_500.0)
                        .amount_local(1_500.0)
                        .build(),
                )
                .add_right_panel_entry(
                    PanelEntry::builder()
                        .date("2024-12-25")
                        .amount_euro(800.0)
                        .amount_local(800.0)
                        .build(),
                )
                .add_position(
                    1,
                    PositionEntry::builder()
                        .description("Test position")
                        .approved(5_000.0)
                        .income_report(2_500.0)
                        .income_total(2_500.0)
                        .build(),
                )
                .set_header_input(
                    6,
                    PositionEntry::builder()
                        .approved(1_000.0)
                        .income_report(500.0)
                        .income_total(500.0)
                        .build(),
                )
                .build(),
        )
        .footer(
            ReportFooter::builder()
                .bank(3_500.0)
                .build(),
        )
        .build();

    let path_iso = "examples/output/test_dates_iso.xlsx";
    config_iso.write_to(path_iso)?;
    println!("  [2] ISO dates (YYYY-MM-DD):     {}", path_iso);

    // ── Report 3: Mixed formats + ReportDate type-safe API ───────────
    let config_mixed = ReportConfig::builder()
        .header(
            ReportHeader::builder()
                .language(Language::Francais)
                .currency(Currency::CHF)
                .project_number("DATE-TEST-003")
                .project_title("Test dates — Formats mixtes")
                // Type-safe API with ReportDate → format_de()
                .project_start(ReportDate::new(2024, 3, 1).unwrap().format_de())
                .project_end(ReportDate::new(2025, 2, 28).unwrap().format_de())
                // String API with different formats
                .report_start("01.03.2024")   // German
                .report_end("2024-08-31")     // ISO
                .build(),
        )
        .body(
            ReportBody::builder()
                .kmw_mittel(
                    TableEntry::builder()
                        .approved_budget(15_000.0)
                        .income_report(7_500.0)
                        .income_total(7_500.0)
                        .build(),
                )
                // Mixed date formats in panels
                .add_left_panel_entry(
                    PanelEntry::builder()
                        .date("01.03.2024")    // German
                        .amount_euro(2_000.0)
                        .amount_local(2_000.0)
                        .build(),
                )
                .add_left_panel_entry(
                    PanelEntry::builder()
                        .date("2024-04-15")    // ISO
                        .amount_euro(3_000.0)
                        .amount_local(3_000.0)
                        .build(),
                )
                .add_left_panel_entry(
                    PanelEntry::builder()
                        .date("15/05/2024")    // Slash (EU)
                        .amount_euro(1_500.0)
                        .amount_local(1_500.0)
                        .build(),
                )
                .add_right_panel_entry(
                    PanelEntry::builder()
                        .date("2024-06-01")
                        .amount_euro(2_500.0)
                        .amount_local(2_500.0)
                        .build(),
                )
                .add_position(
                    1,
                    PositionEntry::builder()
                        .description("Position test")
                        .approved(8_000.0)
                        .income_report(4_000.0)
                        .income_total(4_000.0)
                        .build(),
                )
                .set_header_input(
                    6,
                    PositionEntry::builder()
                        .approved(2_000.0)
                        .income_report(1_000.0)
                        .income_total(1_000.0)
                        .build(),
                )
                .build(),
        )
        .footer(
            ReportFooter::builder()
                .bank(5_000.0)
                .kasse(200.0)
                .build(),
        )
        .build();

    let path_mixed = "examples/output/test_dates_mixed.xlsx";
    config_mixed.write_to(path_mixed)?;
    println!("  [3] Mixed formats + ReportDate: {}", path_mixed);

    // ── Report 4: Unparseable dates → string fallback ────────────────
    let config_fallback = ReportConfig::builder()
        .header(
            ReportHeader::builder()
                .language(Language::Deutsch)
                .currency(Currency::EUR)
                .project_number("DATE-TEST-004")
                .project_title("Datumstest — Fallback auf Text")
                .project_start("Q1/2024")          // not a date
                .project_end("Ende 2026")           // not a date
                .report_start("ca. Januar 2024")    // not a date
                .report_end("30.06.2024")           // valid → real Excel date
                .build(),
        )
        .body(
            ReportBody::builder()
                .kmw_mittel(
                    TableEntry::builder()
                        .approved_budget(10_000.0)
                        .income_report(5_000.0)
                        .income_total(5_000.0)
                        .build(),
                )
                // Mix of valid and invalid date strings in panels
                .add_left_panel_entry(
                    PanelEntry::builder()
                        .date("15.01.2024")        // valid → real date
                        .amount_euro(1_000.0)
                        .amount_local(1_000.0)
                        .build(),
                )
                .add_left_panel_entry(
                    PanelEntry::builder()
                        .date("Jan 2024")          // invalid → string fallback
                        .amount_euro(2_000.0)
                        .amount_local(2_000.0)
                        .build(),
                )
                .add_left_panel_entry(
                    PanelEntry::builder()
                        .date("Woche 12")          // invalid → string fallback
                        .amount_euro(500.0)
                        .amount_local(500.0)
                        .build(),
                )
                .add_left_panel_entry(
                    PanelEntry::builder()
                        .date("")                  // empty → string fallback
                        .amount_euro(300.0)
                        .amount_local(300.0)
                        .build(),
                )
                .add_right_panel_entry(
                    PanelEntry::builder()
                        .date("2024-03-15")        // valid → real date
                        .amount_euro(1_500.0)
                        .amount_local(1_500.0)
                        .build(),
                )
                .add_right_panel_entry(
                    PanelEntry::builder()
                        .date("TBD")               // invalid → string fallback
                        .amount_euro(800.0)
                        .amount_local(800.0)
                        .build(),
                )
                .add_right_panel_entry(
                    PanelEntry::builder()
                        .date("31/02/2024")        // invalid day → string fallback
                        .amount_euro(600.0)
                        .amount_local(600.0)
                        .build(),
                )
                .add_position(
                    1,
                    PositionEntry::builder()
                        .description("Test position")
                        .approved(5_000.0)
                        .income_report(2_500.0)
                        .income_total(2_500.0)
                        .build(),
                )
                .set_header_input(
                    6,
                    PositionEntry::builder()
                        .approved(1_000.0)
                        .income_report(500.0)
                        .income_total(500.0)
                        .build(),
                )
                .build(),
        )
        .footer(
            ReportFooter::builder()
                .bank(2_000.0)
                .build(),
        )
        .build();

    let path_fallback = "examples/output/test_dates_fallback.xlsx";
    config_fallback.write_to(path_fallback)?;
    println!("  [4] Unparseable → string fallback: {}", path_fallback);

    println!();
    println!("Reports 1-3: All dates are real Excel date values (locale-adaptive).");
    println!("Report 4: Mix of real dates and string fallbacks for unparseable values.");
    println!();
    println!("Open the files and check:");
    println!("  - Real dates: right-click cell → Format → shows 'Date'");
    println!("  - Fallbacks:  right-click cell → Format → shows 'Text'");

    Ok(())
}
