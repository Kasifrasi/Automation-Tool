//! Test: Mehrsprachige Finanzberichte
//!
//! Generiert 5 Excel-Dateien in verschiedenen Sprachen:
//! - Deutsch (D_Finanzbericht.xlsx)
//! - Englisch (E_Financial_report.xlsx)
//! - Französisch (F_Rapport_financier.xlsx)
//! - Spanisch (S_Informe_financiero.xlsx)
//! - Portugiesisch (P_Relatorio_financeiro.xlsx)
//!
//! **Builder-Pattern** — kein `Some()`, kein `.to_string()`, kein `?` beim Bauen.

use fb_generator::*;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    println!("Erstelle mehrsprachige Finanzberichte...\n");

    let reports: &[(Language, &str, &str)] = &[
        (Language::Deutsch,   "PROJ-2024-DE", "examples/output/D_Finanzbericht.xlsx"),
        (Language::English,   "PROJ-2024-EN", "examples/output/E_Financial_report.xlsx"),
        (Language::Francais,  "PROJ-2024-FR", "examples/output/F_Rapport_financier.xlsx"),
        (Language::Espanol,   "PROJ-2024-ES", "examples/output/S_Informe_financiero.xlsx"),
        (Language::Portugues, "PROJ-2024-PT", "examples/output/P_Relatorio_financeiro.xlsx"),
    ];

    for &(language, proj_nr, filename) in reports {
        create_report(language, proj_nr, filename)?;
        println!("  -> {} gespeichert", filename);
    }

    println!("\nAlle 5 Dateien erstellt.");
    Ok(())
}

fn create_report(
    language: Language,
    project_number: &str,
    output_path: &str,
) -> Result<(), Box<dyn std::error::Error>> {
    println!("Erstelle {}...", language);

    let config = ReportConfig::builder()
        .header(
            ReportHeader::builder()
                .language(language)
                .currency(Currency::EUR)
                .project_number(project_number)
                .project_title(format!("Test Project ({})", language))
                .project_start("01.01.2024")
                .project_end("31.12.2024")
                .report_start("01.01.2024")
                .report_end("30.06.2024")
                .build(),
        )
        .body(
            ReportBody::builder()
                // Einnahmen-Tabelle (5 benannte Zeilen)
                .kmw_mittel(  TableEntry::builder().approved_budget(50000.0).income_report(25000.0).income_total(25000.0).build())
                .eigenmittel( TableEntry::builder().approved_budget(40000.0).income_report(20000.0).income_total(20000.0).build())
                .drittmittel( TableEntry::builder().approved_budget(30000.0).income_report(15000.0).income_total(15000.0).build())
                .saldovortrag(TableEntry::builder().approved_budget(20000.0).income_report(10000.0).income_total(10000.0).build())
                .zinsertraege(TableEntry::builder().approved_budget(10000.0).income_report( 5000.0).income_total( 5000.0).build())
                // Kategorie 1: 5 Positionen
                .add_position(1, PositionEntry::builder().description("Position 1.1").approved(15000.0).income_report( 7500.0).income_total( 7500.0).build())
                .add_position(1, PositionEntry::builder().description("Position 1.2").approved(25000.0).income_report(12000.0).income_total(12000.0).build())
                .add_position(1, PositionEntry::builder().description("Position 1.3").approved( 5000.0).income_report( 2500.0).income_total( 2500.0).build())
                .add_position(1, PositionEntry::builder().description("Position 1.4").approved( 8000.0).income_report( 4000.0).income_total( 4000.0).build())
                .add_position(1, PositionEntry::builder().description("Position 1.5").approved( 2000.0).income_report( 1000.0).income_total( 1000.0).build())
                // Kategorie 2: 3 Positionen
                .add_position(2, PositionEntry::builder().description("Position 2.1").approved(1500.0).income_report( 800.0).income_total( 800.0).build())
                .add_position(2, PositionEntry::builder().description("Position 2.2").approved(3000.0).income_report(2000.0).income_total(2000.0).build())
                .add_position(2, PositionEntry::builder().description("Position 2.3").approved(2000.0).income_report(1500.0).income_total(1500.0).build())
                // Kategorie 3: 4 Positionen
                .add_position(3, PositionEntry::builder().description("Position 3.1").approved(2000.0).income_report(1200.0).income_total(1200.0).build())
                .add_position(3, PositionEntry::builder().description("Position 3.2").approved(5000.0).income_report(2500.0).income_total(2500.0).build())
                .add_position(3, PositionEntry::builder().description("Position 3.3").approved(1000.0).income_report( 500.0).income_total( 500.0).build())
                .add_position(3, PositionEntry::builder().description("Position 3.4").approved( 800.0).income_report( 400.0).income_total( 400.0).build())
                // Kategorie 4: 3 Positionen
                .add_position(4, PositionEntry::builder().description("Position 4.1").approved(10000.0).income_report(5000.0).income_total(5000.0).build())
                .add_position(4, PositionEntry::builder().description("Position 4.2").approved( 8000.0).income_report(4000.0).income_total(4000.0).build())
                .add_position(4, PositionEntry::builder().description("Position 4.3").approved( 2000.0).income_report(1000.0).income_total(1000.0).build())
                // Kategorie 5: 2 Positionen
                .add_position(5, PositionEntry::builder().description("Position 5.1").approved(3000.0).income_report(1500.0).income_total(1500.0).build())
                .add_position(5, PositionEntry::builder().description("Position 5.2").approved(1500.0).income_report( 750.0).income_total( 750.0).build())
                // Kategorien 6-8: Header-Eingabe
                .set_header_input(6, PositionEntry::builder().approved(4000.0).income_report(2000.0).income_total(2000.0).build())
                .set_header_input(7, PositionEntry::builder().approved(6000.0).income_report(3000.0).income_total(3000.0).build())
                .set_header_input(8, PositionEntry::builder().approved(2500.0).income_report(1250.0).income_total(1250.0).build())
                .build(),
        )
        .footer(
            ReportFooter::builder()
                .bank(10000.0)
                .kasse(2000.0)
                .sonstiges(500.0)
                .build(),
        )
        .options(
            ReportOptions::builder()
                .sheet_protection(SheetProtection::from_defaults().with_password("sheet_pw"))
                .workbook_password("wb_pw")
                .hide_columns_qv(true)
                .build(),
        )
        .build();

    config.write_to(output_path)?;
    Ok(())
}
