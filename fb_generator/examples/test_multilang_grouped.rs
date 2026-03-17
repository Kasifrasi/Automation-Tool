//! Test: Mehrsprachige Reports mit Gruppierung
//!
//! Erstellt Reports in allen 5 Sprachen mit:
//! - Bauausgaben: 20 Positionen (letzte 15 gruppiert)
//! - Investitionen: 20 Positionen (letzte 15 gruppiert)
//! - Personalausgaben: 30 Positionen (letzte 18 gruppiert)
//! - Projektaktivitäten: 30 Positionen (letzte 18 gruppiert)
//! - Projektverwaltung: 20 Positionen (letzte 15 gruppiert)
//! - Evaluierung: 0 Positionen (nur Header)
//! - Audit: 0 Positionen (nur Header)
//! - Reserve: 0 Positionen (nur Header)
//!
//! **Builder-Pattern** — kein `Some()`, kein `.to_string()`, kein `?` beim Bauen.

use fb_generator::*;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    for lang in Language::all() {
        create_report(*lang)?;
    }

    println!("\n✅ Alle 5 Reports erstellt!");
    Ok(())
}

fn create_report(language: Language) -> Result<(), Box<dyn std::error::Error>> {
    println!("Erstelle Report für: {}", language);

    // Zeilen-Berechnung für Gruppierung (Body startet bei Zeile 26, 0-basiert):
    //
    //   Kat 1 (20 Pos): Header=26, Pos=27-46, Footer=47 → Gruppe 32-46 (letzte 15)
    //   Kat 2 (20 Pos): Header=48, Pos=49-68, Footer=69 → Gruppe 54-68 (letzte 15)
    //   Kat 3 (30 Pos): Header=70, Pos=71-100, Footer=101 → Gruppe 83-100 (letzte 18)
    //   Kat 4 (30 Pos): Header=102, Pos=103-132, Footer=133 → Gruppe 115-132 (letzte 18)
    //   Kat 5 (20 Pos): Header=134, Pos=135-154, Footer=155 → Gruppe 140-154 (letzte 15)
    //   Kat 6-8: je 1 Zeile (Header-Eingabe), keine Gruppierung

    let config = ReportConfig::builder()
        .header(
            ReportHeader::builder()
                .language(language)
                .build(),
        )
        .options(
            ReportOptions::builder()
                .hide_language_sheet(true)
                .row_grouping(
                    RowGrouping::new()
                        .add_collapsed_group(32, 46)
                        .add_collapsed_group(54, 68)
                        .add_collapsed_group(83, 100)
                        .add_collapsed_group(115, 132)
                        .add_collapsed_group(140, 154),
                )
                .build(),
        )
        .build();

    let filename = format!("examples/output/test_grouped_{}.xlsx", language);
    config.write_to(&filename)?;
    println!("  ✓ {}", filename);

    Ok(())
}
