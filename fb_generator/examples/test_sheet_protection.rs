//! Test: Sheet Protection mit individuellen Checkboxen
//!
//! Generiert drei Reports mit unterschiedlichen Schutz-Konfigurationen:
//! 1. from_defaults() — empfohlene Voreinstellung (Formatierung + Sortieren erlaubt)
//! 2. new() — maximaler Schutz (alles gesperrt außer Zellauswahl)
//! 3. Individuelle Konfiguration — nur bestimmte Aktionen erlaubt

use fb_generator::*;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    println!("Sheet-Protection Beispiele\n");

    // ── 1. from_defaults(): Empfohlene Voreinstellung ──────────────────
    //
    // Erlaubt: Formatierung, Sortieren, AutoFilter, PivotTables, Hyperlinks, Szenarien
    // Gesperrt: Einfügen/Löschen von Zeilen/Spalten, Objekte bearbeiten
    let config_defaults = ReportConfig::builder()
        .header(
            ReportHeader::builder()
                .language(Language::Deutsch)
                .project_title("Test: from_defaults()")
                .build(),
        )
        .body(
            ReportBody::builder()
                .kmw_mittel(TableEntry::builder().approved_budget(50_000.0).build())
                .add_position(1, PositionEntry::builder().approved(10_000.0).build())
                .build(),
        )
        .options(
            ReportOptions::builder()
                .sheet_protection(
                    SheetProtection::from_defaults()
                        .with_password("defaults_pw"),
                )
                .build(),
        )
        .build();
    config_defaults.write_to("examples/output/test_protection_defaults.xlsx")?;
    println!("   OK: test_protection_defaults.xlsx (from_defaults + Passwort)");

    // ── 2. new(): Maximaler Schutz ─────────────────────────────────────
    //
    // Alles gesperrt, nur Zellauswahl (locked + unlocked) erlaubt.
    // Kein Passwort → Schutz kann per Klick entfernt werden.
    let config_locked = ReportConfig::builder()
        .header(
            ReportHeader::builder()
                .language(Language::English)
                .project_title("Test: new() - max lock")
                .build(),
        )
        .body(
            ReportBody::builder()
                .kmw_mittel(TableEntry::builder().approved_budget(25_000.0).build())
                .add_position(1, PositionEntry::builder().approved(5_000.0).build())
                .build(),
        )
        .options(
            ReportOptions::builder()
                .sheet_protection(SheetProtection::new())
                .build(),
        )
        .build();
    config_locked.write_to("examples/output/test_protection_max_lock.xlsx")?;
    println!("   OK: test_protection_max_lock.xlsx (alles gesperrt, kein Passwort)");

    // ── 3. Individuelle Konfiguration ──────────────────────────────────
    //
    // Startet mit new() (alles gesperrt) und schaltet gezielt frei:
    // - Sortieren
    // - AutoFilter
    // - Zeilen formatieren (Höhe anpassen)
    // Rest bleibt gesperrt.
    let config_custom = ReportConfig::builder()
        .header(
            ReportHeader::builder()
                .language(Language::Deutsch)
                .project_title("Test: Individuelle Checkboxen")
                .build(),
        )
        .body(
            ReportBody::builder()
                .kmw_mittel(TableEntry::builder().approved_budget(80_000.0).build())
                .add_position(1, PositionEntry::builder().approved(20_000.0).build())
                .set_header_input(6, PositionEntry::builder().approved(3_000.0).build())
                .build(),
        )
        .options(
            ReportOptions::builder()
                .sheet_protection(
                    SheetProtection::new()
                        .allow_sort(true)
                        .allow_autofilter(true)
                        .allow_format_rows(true)
                        .with_password("custom_pw"),
                )
                .hide_columns_qv(true)
                .build(),
        )
        .build();
    config_custom.write_to("examples/output/test_protection_custom.xlsx")?;
    println!("   OK: test_protection_custom.xlsx (nur Sortieren + AutoFilter + Zeilen-Format)");

    println!("\nAlle Sheet-Protection Beispiele erstellt.");
    Ok(())
}
