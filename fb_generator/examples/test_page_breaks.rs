//! Test: Seitenumbrüche bei verschiedenen Berichtsgrößen
//!
//! Erzeugt vier Berichte mit unterschiedlicher Anzahl von Kostenpositionen,
//! um das Verhalten der automatischen Druckumbrüche zu testen:
//!
//! 1. **Klein** — Alle Kategorien passen auf eine Seite (keine Umbrüche)
//! 2. **Normal** — Mehrere Kategorien verursachen einen Seitenumbruch an Kategorie-Grenze
//! 3. **Übergroß** — Kategorie 1 hat 80 Positionen (> MAX_ROWS_PER_PAGE),
//!    erzwingt harte Umbrüche innerhalb der Kategorie
//! 4. **Mehrere Große** — Kategorien 1 und 3 jeweils 60 Positionen,
//!    testen gestaffelte Umbrüche über mehrere Kategorien
//!
//! Manuelle Prüfung in Excel: Datei → Drucken → Druckvorschau
//!   - Umbrüche sollten an Kategorie-Grenzen liegen (wenn möglich)
//!   - Übergroße Kategorien brechen stumpf nach je 80 Zeilen um
//!   - Panel (J:V) ist immer auf der letzten Seite

use fb_generator::*;

/// Hilfsfunktion: Erzeugt N Positionen für eine Kategorie auf dem Builder
fn fill_positions(body: ReportBodyBuilder, category: u8, count: usize) -> ReportBodyBuilder {
    let mut b = body;
    for i in 1..=count {
        b = b.add_position(
            category,
            PositionEntry::builder()
                .description(format!("Position {category}.{i}"))
                .approved(1_000.0 * i as f64)
                .income_report(500.0 * i as f64)
                .income_total(500.0 * i as f64)
                .remark(format!("Zeile {i}"))
                .build(),
        );
    }
    b
}

fn main() -> Result<(), Box<dyn std::error::Error>> {
    std::fs::create_dir_all("examples/output")?;

    println!("Testing page break behavior with different report sizes...\n");

    // =========================================================================
    // Test 1: Klein — alles passt auf eine Seite
    // =========================================================================
    {
        println!("1. Small report (all fits on one page, no breaks expected)...");
        let body = ReportBody::builder()
            .kmw_mittel(
                TableEntry::builder()
                    .approved_budget(50_000.0)
                    .income_report(25_000.0)
                    .income_total(25_000.0)
                    .reason("1. Rate")
                    .build(),
            )
            .eigenmittel(
                TableEntry::builder()
                    .approved_budget(5_000.0)
                    .income_report(2_500.0)
                    .income_total(2_500.0)
                    .build(),
            );

        // Kat 1–5: je 3 Positionen, Kat 6–8: Header-Input
        let body = fill_positions(body, 1, 3);
        let body = fill_positions(body, 2, 3);
        let body = fill_positions(body, 3, 3);
        let body = fill_positions(body, 4, 2);
        let body = fill_positions(body, 5, 2);
        let body = body
            .set_header_input(
                6,
                PositionEntry::builder()
                    .approved(1_000.0)
                    .income_report(500.0)
                    .income_total(500.0)
                    .build(),
            )
            .set_header_input(
                7,
                PositionEntry::builder()
                    .approved(500.0)
                    .income_report(250.0)
                    .income_total(250.0)
                    .build(),
            )
            .set_header_input(
                8,
                PositionEntry::builder()
                    .approved(200.0)
                    .income_report(100.0)
                    .income_total(100.0)
                    .build(),
            );

        let config = ReportConfig::builder()
            .header(
                ReportHeader::builder()
                    .language(Language::Deutsch)
                    .currency(Currency::EUR)
                    .project_number("PB-001")
                    .project_title("Kleiner Bericht — Keine Seitenumbrüche")
                    .project_start("01.01.2025")
                    .project_end("31.12.2026")
                    .report_start("01.01.2025")
                    .report_end("30.06.2025")
                    .build(),
            )
            .body(body.build())
            .footer(
                ReportFooter::builder()
                    .bank(5_000.0)
                    .kasse(800.0)
                    .build(),
            )
            .build();

        config.write_to("examples/output/test_page_breaks_small.xlsx")?;
        // 8 Kategorien mit wenigen Positionen → Header(26) + ~30 Body-Zeilen = ~56 → passt auf 1 Seite
        println!("   OK: test_page_breaks_small.xlsx");
        println!("   ~56 Zeilen total → keine Umbrüche erwartet");
    }

    // =========================================================================
    // Test 2: Normal — Umbruch an Kategorie-Grenze
    // =========================================================================
    {
        println!("\n2. Normal report (break at category boundary)...");
        let body = ReportBody::builder()
            .kmw_mittel(
                TableEntry::builder()
                    .approved_budget(100_000.0)
                    .income_report(50_000.0)
                    .income_total(50_000.0)
                    .reason("Hauptfinanzierung")
                    .build(),
            )
            .eigenmittel(
                TableEntry::builder()
                    .approved_budget(15_000.0)
                    .income_report(7_500.0)
                    .income_total(7_500.0)
                    .build(),
            );

        // Kat 1: 20 Positionen (Header=26, Pos=27-46, Footer=47 → 22 Zeilen)
        // Kat 2: 20 Positionen (Header=48, Pos=49-68, Footer=69 → 22 Zeilen)
        //   → Nach Kat 1: rows_on_page = 26+22 = 48
        //   → Kat 2 braucht 22, total = 70 < 80 → passt noch
        // Kat 3: 15 Positionen (Header=70, Pos=71-85, Footer=86 → 17 Zeilen)
        //   → rows_on_page = 70+17 = 87 > 80 → Umbruch VOR Kat 3!
        let body = fill_positions(body, 1, 20);
        let body = fill_positions(body, 2, 20);
        let body = fill_positions(body, 3, 15);
        let body = fill_positions(body, 4, 10);
        let body = fill_positions(body, 5, 5);
        let body = body
            .set_header_input(
                6,
                PositionEntry::builder()
                    .approved(2_000.0)
                    .income_report(1_000.0)
                    .income_total(1_000.0)
                    .build(),
            )
            .set_header_input(
                7,
                PositionEntry::builder()
                    .approved(1_000.0)
                    .income_report(500.0)
                    .income_total(500.0)
                    .build(),
            )
            .set_header_input(
                8,
                PositionEntry::builder()
                    .approved(500.0)
                    .income_report(250.0)
                    .income_total(250.0)
                    .build(),
            );

        let config = ReportConfig::builder()
            .header(
                ReportHeader::builder()
                    .language(Language::English)
                    .currency(Currency::USD)
                    .project_number("PB-002")
                    .project_title("Normal Report — Category Boundary Breaks")
                    .project_start("01.01.2025")
                    .project_end("31.12.2027")
                    .report_start("01.01.2025")
                    .report_end("31.12.2025")
                    .build(),
            )
            .body(body.build())
            .footer(
                ReportFooter::builder()
                    .bank(12_000.0)
                    .kasse(1_500.0)
                    .sonstiges(300.0)
                    .build(),
            )
            .build();

        config.write_to("examples/output/test_page_breaks_normal.xlsx")?;
        println!("   OK: test_page_breaks_normal.xlsx");
        println!("   Kat 1 (22 Zl) + Kat 2 (22 Zl) = 70 Zeilen auf Seite 1");
        println!("   Kat 3 passt nicht mehr → Umbruch vor Kat 3");
    }

    // =========================================================================
    // Test 3: Übergroß — Kategorie 1 allein > 80 Zeilen
    // =========================================================================
    {
        println!("\n3. Oversized category (Cat 1 = 80 positions, hard breaks within)...");
        let body = ReportBody::builder()
            .kmw_mittel(
                TableEntry::builder()
                    .approved_budget(200_000.0)
                    .income_report(100_000.0)
                    .income_total(100_000.0)
                    .reason("Großprojekt")
                    .build(),
            )
            .eigenmittel(
                TableEntry::builder()
                    .approved_budget(30_000.0)
                    .income_report(15_000.0)
                    .income_total(15_000.0)
                    .build(),
            );

        // Kat 1: 80 Positionen → Header(26) + 80 Pos + Footer = 82 Zeilen
        // Zusammen mit Report-Header (26 Zeilen): 26 + 82 = 108 Zeilen
        //   → Seite 1 bis Zeile 79 (80 Zeilen), Umbruch innerhalb Kat 1
        //   → Seite 2 ab Zeile 80, restliche Positionen + Footer
        let body = fill_positions(body, 1, 80);
        let body = fill_positions(body, 2, 5);
        let body = fill_positions(body, 3, 5);
        let body = body
            .set_header_input(
                4,
                PositionEntry::builder()
                    .approved(3_000.0)
                    .income_report(1_500.0)
                    .income_total(1_500.0)
                    .remark("Pauschal")
                    .build(),
            )
            .set_header_input(
                5,
                PositionEntry::builder()
                    .approved(2_000.0)
                    .income_report(1_000.0)
                    .income_total(1_000.0)
                    .build(),
            )
            .set_header_input(
                6,
                PositionEntry::builder()
                    .approved(1_500.0)
                    .income_report(750.0)
                    .income_total(750.0)
                    .build(),
            )
            .set_header_input(
                7,
                PositionEntry::builder()
                    .approved(1_000.0)
                    .income_report(500.0)
                    .income_total(500.0)
                    .build(),
            )
            .set_header_input(
                8,
                PositionEntry::builder()
                    .approved(500.0)
                    .income_report(250.0)
                    .income_total(250.0)
                    .build(),
            );

        let config = ReportConfig::builder()
            .header(
                ReportHeader::builder()
                    .language(Language::Deutsch)
                    .currency(Currency::EUR)
                    .project_number("PB-003")
                    .project_title("Übergroße Kategorie — Harte Umbrüche innerhalb")
                    .project_start("01.01.2025")
                    .project_end("31.12.2028")
                    .report_start("01.01.2025")
                    .report_end("31.12.2025")
                    .build(),
            )
            .body(body.build())
            .footer(
                ReportFooter::builder()
                    .bank(25_000.0)
                    .kasse(3_000.0)
                    .sonstiges(500.0)
                    .build(),
            )
            .options(
                ReportOptions::builder()
                    .sheet_protection(SheetProtection::from_defaults().with_password("test"))
                    .build(),
            )
            .build();

        config.write_to("examples/output/test_page_breaks_oversized.xlsx")?;
        println!("   OK: test_page_breaks_oversized.xlsx");
        println!("   Kat 1 = 82 Zeilen → harter Umbruch innerhalb nach 80 Zeilen");
        println!("   Restliche Kategorien folgen mit normalen Umbruchregeln");
    }

    // =========================================================================
    // Test 4: Mehrere große Kategorien — gestaffelte Umbrüche
    // =========================================================================
    {
        println!("\n4. Multiple large categories (Cat 1=60, Cat 3=60, staggered breaks)...");
        let body = ReportBody::builder()
            .kmw_mittel(
                TableEntry::builder()
                    .approved_budget(300_000.0)
                    .income_report(150_000.0)
                    .income_total(150_000.0)
                    .reason("Multi-Projekt")
                    .build(),
            )
            .eigenmittel(
                TableEntry::builder()
                    .approved_budget(40_000.0)
                    .income_report(20_000.0)
                    .income_total(20_000.0)
                    .build(),
            )
            .add_left_panel_entry(
                PanelEntry::builder()
                    .date("01.01.2025")
                    .amount_euro(50_000.0)
                    .amount_local(50_000.0)
                    .build(),
            )
            .add_left_panel_entry(
                PanelEntry::builder()
                    .date("01.04.2025")
                    .amount_euro(50_000.0)
                    .amount_local(50_000.0)
                    .build(),
            )
            .add_left_panel_entry(
                PanelEntry::builder()
                    .date("01.07.2025")
                    .amount_euro(50_000.0)
                    .amount_local(50_000.0)
                    .build(),
            );

        // Kat 1: 60 Positionen (Header + 60 Pos + Footer = 62 Zeilen)
        // Kat 2: 10 Positionen (Header + 10 Pos + Footer = 12 Zeilen)
        // Kat 3: 60 Positionen (Header + 60 Pos + Footer = 62 Zeilen)
        // Kat 4: 10 Positionen
        // Kat 5-8: Header-Input
        //
        // Seite 1: Header(26) + Kat 1 beginnt bei Zeile 26
        //   26 + 62 = 88 > 80 → Seite überläuft bei erster Kategorie
        //   → Harter Umbruch bei Zeile 80 innerhalb Kat 1
        // Seite 2: Rest von Kat 1 (8 Zeilen) + Kat 2 (12 Zeilen) = 20 Zeilen
        //   + Kat 3 (62 Zeilen) = 82 > 80 → Umbruch VOR Kat 3
        // Seite 3: Kat 3 (62 Zeilen) → passt auf eine Seite
        //   + Kat 4 (12 Zeilen) = 74 → passt noch
        // Seite 3: Rest + Header-Inputs → passt
        let body = fill_positions(body, 1, 60);
        let body = fill_positions(body, 2, 10);
        let body = fill_positions(body, 3, 60);
        let body = fill_positions(body, 4, 10);
        let body = body
            .set_header_input(
                5,
                PositionEntry::builder()
                    .approved(5_000.0)
                    .income_report(2_500.0)
                    .income_total(2_500.0)
                    .remark("Direktkosten")
                    .build(),
            )
            .set_header_input(
                6,
                PositionEntry::builder()
                    .approved(3_000.0)
                    .income_report(1_500.0)
                    .income_total(1_500.0)
                    .build(),
            )
            .set_header_input(
                7,
                PositionEntry::builder()
                    .approved(2_000.0)
                    .income_report(1_000.0)
                    .income_total(1_000.0)
                    .build(),
            )
            .set_header_input(
                8,
                PositionEntry::builder()
                    .approved(1_000.0)
                    .income_report(500.0)
                    .income_total(500.0)
                    .build(),
            );

        let config = ReportConfig::builder()
            .header(
                ReportHeader::builder()
                    .language(Language::Francais)
                    .currency(Currency::EUR)
                    .project_number("PB-004")
                    .project_title("Plusieurs grandes catégories — Sauts échelonnés")
                    .project_start("01.01.2025")
                    .project_end("31.12.2028")
                    .report_start("01.01.2025")
                    .report_end("31.12.2025")
                    .version("v2025-Q4")
                    .build(),
            )
            .body(body.build())
            .footer(
                ReportFooter::builder()
                    .bank(40_000.0)
                    .kasse(5_000.0)
                    .sonstiges(1_000.0)
                    .build(),
            )
            .options(
                ReportOptions::builder()
                    .hide_language_sheet(true)
                    .build(),
            )
            .build();

        config.write_to("examples/output/test_page_breaks_multi_large.xlsx")?;
        println!("   OK: test_page_breaks_multi_large.xlsx");
        println!("   Kat 1 (62 Zl) → harter Umbruch innerhalb");
        println!("   Kat 3 (62 Zl) → Umbruch vor Header, dann passt sie auf eine Seite");
        println!("   Sprache: Französisch, Panel (J:V) auf letzter Seite");
    }

    println!("\nAll page break tests completed successfully!");
    println!("\nVerifizierung in Excel:");
    println!("  Datei → Drucken → Druckvorschau (oder Strg+P)");
    println!("  - Hochformat A4, Breite = 1 Seite");
    println!("  - Umbrüche an Kategorie-Grenzen (wenn < 80 Zeilen)");
    println!("  - Harte Umbrüche bei übergroßen Kategorien (alle 80 Zeilen)");
    println!("  - Panel (Spalten J:V) immer auf letzter Seite");

    Ok(())
}
