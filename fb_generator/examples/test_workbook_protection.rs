//! Test: Workbook-Protection Integration
//!
//! Demonstriert die `ReportConfig` Builder-API mit verschiedenen Schutz-Stufen:
//! 1. Kein Schutz
//! 2. Nur Workbook-Protection (Passwort)
//! 3. Sheet-Protection + Workbook-Protection (getrennte Passwörter)
//!
//! **Builder-Pattern** — kein `Some()`, kein `.to_string()`, kein `?` beim Bauen.

use fb_generator::*;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    println!("Testing Workbook-Protection Integration...\n");

    // Test 1: Report ohne Schutz
    {
        println!("1. Creating report WITHOUT protection...");
        let config = ReportConfig::builder()
            .header(
                ReportHeader::builder()
                    .language(Language::Deutsch)
                    .currency(Currency::EUR)
                    .project_number("TEST-001")
                    .project_title("Test Project")
                    .project_start("01.01.2024")
                    .project_end("31.12.2024")
                    .report_start("01.01.2024")
                    .report_end("31.03.2024")
                    .build(),
            )
            .build();
        config.write_to("examples/output/test_no_wb_protection.xlsx")?;
        println!("   ✓ Created: examples/output/test_no_wb_protection.xlsx");
    }

    // Test 2: Report mit Workbook-Protection (kein Sheet-Schutz)
    {
        println!("\n2. Creating report WITH workbook protection...");
        let config = ReportConfig::builder()
            .header(
                ReportHeader::builder()
                    .language(Language::Deutsch)
                    .currency(Currency::EUR)
                    .project_number("TEST-002")
                    .project_title("Protected Project")
                    .project_start("01.01.2024")
                    .project_end("31.12.2024")
                    .report_start("01.01.2024")
                    .report_end("31.03.2024")
                    .build(),
            )
            .options(
                ReportOptions::builder()
                    .workbook_password("secret123")
                    .build(),
            )
            .build();
        config.write_to("examples/output/test_with_wb_protection.xlsx")?;
        println!("   ✓ Created: examples/output/test_with_wb_protection.xlsx");
        println!("   ℹ Workbook Password: secret123");
    }

    // Test 3: Report mit Sheet- + Workbook-Protection (getrennte Passwörter)
    {
        println!("\n3. Creating report WITH sheet AND workbook protection...");
        let config = ReportConfig::builder()
            .header(
                ReportHeader::builder()
                    .language(Language::Deutsch)
                    .currency(Currency::EUR)
                    .project_number("TEST-003")
                    .project_title("Fully Protected Project")
                    .project_start("01.01.2024")
                    .project_end("31.12.2024")
                    .report_start("01.01.2024")
                    .report_end("31.03.2024")
                    .build(),
            )
            .options(
                ReportOptions::builder()
                    .sheet_protection(SheetProtection::from_defaults().with_password("blatt_geheim"))
                    .workbook_password("wb_geheim")
                    .hide_columns_qv(true)
                    .hide_language_sheet(true)
                    .build(),
            )
            .build();
        config.write_to("examples/output/test_full_protection.xlsx")?;
        println!("   ✓ Created: examples/output/test_full_protection.xlsx");
        println!("   ℹ Sheet Password:    blatt_geheim");
        println!("   ℹ Workbook Password: wb_geheim");
    }

    println!("\n✅ All tests completed successfully!");
    println!("\nManual verification:");
    println!("1. Open test_no_wb_protection.xlsx -> should allow sheet operations");
    println!(
        "2. Open test_with_wb_protection.xlsx -> try to move/delete sheets (requires password)"
    );
    println!(
        "3. Open test_full_protection.xlsx -> both sheet cells and workbook structure protected"
    );

    Ok(())
}
