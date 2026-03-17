//! Test: Mixed category modes for categories 6–8
//!
//! Demonstriert drei verschiedene Kombinationen für die Kostenkategorien 6, 7 und 8:
//!
//! 1. **Standard** — Kategorien 1–5 als Multi-Row, Kategorien 6–8 als Header-Input
//!    (klassische Verwendung: aggregierte Einzelwerte ohne Positions-Zeilen)
//!
//! 2. **Multi-Row** — Kategorien 6–8 *ebenfalls* mit mehreren Zeilen via `add_position()`
//!    (z.B. wenn Kategorie 6 mehrere Einzel-Buchungen ausweisen soll)
//!
//! 3. **Gemischt** — Kategorien 1–4 Multi-Row, Kategorie 5 Header-Input,
//!    Kategorien 6–8 Multi-Row mit Skip-Zeilen
//!
//! **Builder-Pattern** — kein `Some()`, kein `.to_string()`, kein `?` beim Bauen.

use fb_generator::*;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    std::fs::create_dir_all("examples/output")?;

    println!("Testing mixed category modes for categories 6-8...\n");

    // =========================================================================
    // Test 1: Kategorien 6-8 als klassische Header-Inputs (aggregierte Werte)
    // =========================================================================
    {
        println!("1. Categories 6-8 as header-inputs (aggregated single values)...");
        let config = ReportConfig::builder()
            .header(
                ReportHeader::builder()
                    .language(Language::Deutsch)
                    .currency(Currency::EUR)
                    .project_number("MIX-001")
                    .project_title("Gemischte Kategorien — Header-Input")
                    .project_start("01.01.2025")
                    .project_end("31.12.2027")
                    .report_start("01.01.2025")
                    .report_end("30.06.2025")
                    .build(),
            )
            .body(
                ReportBody::builder()
                    // ========================================================
                    // Einnahmen-Tabelle
                    // ========================================================
                    .kmw_mittel(   TableEntry::builder().approved_budget(80_000.0).income_report(50_000.0).income_total(50_000.0).reason("1. Rate").build())
                    .eigenmittel(  TableEntry::builder().approved_budget(10_000.0).income_report( 5_000.0).income_total( 5_000.0).reason("Eigenanteil").build())
                    .saldovortrag( TableEntry::builder().approved_budget( 2_000.0).income_report( 2_000.0).income_total( 2_000.0).build())
                    // ========================================================
                    // Linkes Panel (Belege 1–6)
                    // ========================================================
                    .add_left_panel_entry(PanelEntry::builder().date("15.01.2025").amount_euro(20_000.0).amount_local(20_000.0).build())
                    .add_left_panel_entry(PanelEntry::builder().date("01.03.2025").amount_euro(15_000.0).amount_local(15_000.0).build())
                    .add_left_panel_entry(PanelEntry::builder().date("15.04.2025").amount_euro(10_000.0).amount_local(10_000.0).build())
                    .add_left_panel_entry(PanelEntry::builder().date("01.05.2025").amount_euro( 5_000.0).amount_local( 5_000.0).build())
                    .add_left_panel_entry(PanelEntry::builder().date("15.05.2025").amount_euro( 3_000.0).amount_local( 3_000.0).build())
                    .add_left_panel_entry(PanelEntry::builder().date("30.06.2025").amount_euro( 4_000.0).amount_local( 4_000.0).build())
                    // ========================================================
                    // Rechtes Panel (Belege 7–10)
                    // ========================================================
                    .add_right_panel_entry(PanelEntry::builder().date("10.02.2025").amount_euro( 8_000.0).amount_local( 8_000.0).build())
                    .add_right_panel_entry(PanelEntry::builder().date("20.03.2025").amount_euro( 4_500.0).amount_local( 4_500.0).build())
                    .add_right_panel_entry(PanelEntry::builder().date("10.05.2025").amount_euro( 2_000.0).amount_local( 2_000.0).build())
                    .add_right_panel_entry(PanelEntry::builder().date("25.06.2025").amount_euro( 1_500.0).amount_local( 1_500.0).build())
                    // ========================================================
                    // Kategorien 1–5: Multi-Row Positionen
                    // ========================================================
                    .add_position(1, PositionEntry::builder().description("Projektleitung")   .approved(18_000.0).income_report( 9_000.0).income_total( 9_000.0).remark("6 Monate").build())
                    .add_position(1, PositionEntry::builder().description("Buchhaltung")      .approved( 5_000.0).income_report( 2_500.0).income_total( 2_500.0).remark("Teilzeit").build())
                    .add_position(1, PositionEntry::builder().description("Sozialabgaben")    .approved( 3_500.0).income_report( 1_750.0).income_total( 1_750.0).remark("AG-Anteile").build())
                    .add_position(2, PositionEntry::builder().description("Dienstreisen")     .approved( 4_000.0).income_report( 2_000.0).income_total( 2_000.0).remark("2 Reisen").build())
                    .add_position(2, PositionEntry::builder().description("Lokale Fahrten")   .approved( 1_500.0).income_report(   750.0).income_total(   750.0).build())
                    .add_position(3, PositionEntry::builder().description("Büromaterial")      .approved( 2_000.0).income_report( 1_000.0).income_total( 1_000.0).build())
                    .add_position(3, PositionEntry::builder().description("Lehrmaterialien")  .approved( 3_000.0).income_report( 1_500.0).income_total( 1_500.0).remark("Schulbücher").build())
                    .add_position(4, PositionEntry::builder().description("Laptops")          .approved( 4_000.0).income_report( 4_000.0).income_total( 4_000.0).remark("3 Stück").build())
                    .add_position(5, PositionEntry::builder().description("Fortbildung")      .approved( 3_000.0).income_report( 1_500.0).income_total( 1_500.0).remark("2 Seminare").build())
                    .add_position(5, PositionEntry::builder().description("Supervision")      .approved( 1_500.0).income_report(   750.0).income_total(   750.0).build())
                    // ========================================================
                    // Kategorien 6–8: Header-Inputs (aggregierter Einzelwert)
                    // ========================================================
                    .set_header_input(6, PositionEntry::builder().approved( 3_000.0).income_report( 1_500.0).income_total( 1_500.0).remark("Verwaltung pauschal").build())
                    .set_header_input(7, PositionEntry::builder().approved( 1_500.0).income_report(   750.0).income_total(   750.0).remark("Externes Audit").build())
                    .set_header_input(8, PositionEntry::builder().approved(   800.0).income_report(   400.0).income_total(   400.0).remark("Reservefonds").build())
                    .build(),
            )
            .footer(
                ReportFooter::builder()
                    .bank(8_500.0)
                    .kasse(1_200.0)
                    .sonstiges(300.0)
                    .build(),
            )
            .build();

        config.write_to("examples/output/test_mixed_cat_header_inputs.xlsx")?;
        println!("   OK: examples/output/test_mixed_cat_header_inputs.xlsx");
        println!("   Kategorien 1-5 → Multi-Row | Kategorien 6/7/8 → Header-Input");
    }

    // =========================================================================
    // Test 2: Kategorien 6-8 als Multi-Row (mehrere Einzel-Positionen)
    // =========================================================================
    {
        println!("\n2. Categories 6-8 as multi-row positions (multiple entries each)...");
        let config = ReportConfig::builder()
            .header(
                ReportHeader::builder()
                    .language(Language::English)
                    .currency(Currency::USD)
                    .project_number("MIX-002")
                    .project_title("Multi-Row Categories 6-8")
                    .project_start("01.01.2025")
                    .project_end("31.12.2027")
                    .report_start("01.01.2025")
                    .report_end("30.06.2025")
                    .version("v2025-Q2")
                    .build(),
            )
            .body(
                ReportBody::builder()
                    // ========================================================
                    // Einnahmen-Tabelle
                    // ========================================================
                    .kmw_mittel(  TableEntry::builder().approved_budget(60_000.0).income_report(30_000.0).income_total(30_000.0).reason("1st installment").build())
                    .eigenmittel( TableEntry::builder().approved_budget(10_000.0).income_report( 5_000.0).income_total( 5_000.0).reason("Own contribution").build())
                    .drittmittel( TableEntry::builder().approved_budget( 5_000.0).income_report( 2_500.0).income_total( 2_500.0).reason("Foundation XY").build())
                    .zinsertraege(TableEntry::builder().approved_budget(   200.0).income_report(   100.0).income_total(   100.0).reason("Savings account").build())
                    // ========================================================
                    // Linkes Panel (Belege 1–4)
                    // ========================================================
                    .add_left_panel_entry(PanelEntry::builder().date("15.01.2025").amount_euro(12_000.0).amount_local(13_200.0).build())
                    .add_left_panel_entry(PanelEntry::builder().date("01.03.2025").amount_euro( 8_000.0).amount_local( 8_800.0).build())
                    .add_left_panel_entry(PanelEntry::builder().date("15.04.2025").amount_euro( 6_000.0).amount_local( 6_600.0).build())
                    .add_left_panel_entry(PanelEntry::builder().date("01.06.2025").amount_euro(11_600.0).amount_local(12_760.0).build())
                    // ========================================================
                    // Kategorien 1–5: Multi-Row
                    // ========================================================
                    .add_position(1, PositionEntry::builder().description("Project manager")   .approved(15_000.0).income_report( 7_500.0).income_total( 7_500.0).remark("50% FTE").build())
                    .add_position(1, PositionEntry::builder().description("Accountant")        .approved( 8_000.0).income_report( 4_000.0).income_total( 4_000.0).remark("Part-time").build())
                    .add_position(1, PositionEntry::builder().description("Social contributions").approved(3_500.0).income_report( 1_750.0).income_total( 1_750.0).build())
                    .add_position(2, PositionEntry::builder().description("International travel").approved(4_000.0).income_report( 2_000.0).income_total( 2_000.0).remark("2 trips").build())
                    .add_position(2, PositionEntry::builder().description("Local transport")   .approved( 1_500.0).income_report(   750.0).income_total(   750.0).build())
                    .add_position(3, PositionEntry::builder().description("Office supplies")   .approved( 1_500.0).income_report(   750.0).income_total(   750.0).build())
                    .add_position(3, PositionEntry::builder().description("Teaching materials") .approved( 2_500.0).income_report( 1_250.0).income_total( 1_250.0).remark("Textbooks").build())
                    .add_position(4, PositionEntry::builder().description("Laptops")           .approved( 4_000.0).income_report( 4_000.0).income_total( 4_000.0).remark("3 units").build())
                    .add_position(4, PositionEntry::builder().description("Projector")         .approved( 1_200.0).income_report( 1_200.0).income_total( 1_200.0).remark("1 unit").build())
                    .add_position(5, PositionEntry::builder().description("Teacher training")  .approved( 3_000.0).income_report( 1_500.0).income_total( 1_500.0).remark("2 workshops").build())
                    .add_position(5, PositionEntry::builder().description("Supervision")       .approved( 1_500.0).income_report(   750.0).income_total(   750.0).build())
                    // ========================================================
                    // Kategorie 6: Evaluierung — Multi-Row (3 Positionen)
                    // ========================================================
                    .add_position(6, PositionEntry::builder().description("Mid-term evaluation").approved( 2_000.0).income_report( 2_000.0).income_total( 2_000.0).remark("External").build())
                    .add_position(6, PositionEntry::builder().description("Baseline study")    .approved( 1_200.0).income_report( 1_200.0).income_total( 1_200.0).build())
                    .add_position(6, PositionEntry::builder().description("Data collection")   .approved(   600.0).income_report(   300.0).income_total(   300.0).remark("Surveys").build())
                    // ========================================================
                    // Kategorie 7: Audit — Multi-Row mit Skip-Zeile
                    // ========================================================
                    .add_position(7, PositionEntry::builder().description("Annual audit 2025") .approved( 2_500.0).income_report( 2_500.0).income_total( 2_500.0).remark("External auditor").build())
                    .skip_position_row(7)
                    .add_position(7, PositionEntry::builder().description("Audit preparation") .approved(   500.0).income_report(   250.0).income_total(   250.0).build())
                    // ========================================================
                    // Kategorie 8: Reserve — Einzelne Position
                    // ========================================================
                    .add_position(8, PositionEntry::builder().description("Contingency fund")  .approved( 1_000.0).income_report(   500.0).income_total(   500.0).remark("Approved").build())
                    .build(),
            )
            .footer(
                ReportFooter::builder()
                    .bank(12_000.0)
                    .kasse(800.0)
                    .sonstiges(200.0)
                    .build(),
            )
            .options(
                ReportOptions::builder()
                    .hide_columns_qv(true)
                    .build(),
            )
            .build();

        config.write_to("examples/output/test_mixed_cat_multi_row.xlsx")?;
        println!("   OK: examples/output/test_mixed_cat_multi_row.xlsx");
        println!("   Kat. 6 → 3 Positionen | Kat. 7 → 2 + 1 Skip | Kat. 8 → 1 Position");
    }

    // =========================================================================
    // Test 3: Gemischt — unterschiedliche Modi pro Kategorie + version-Feld
    // =========================================================================
    {
        println!("\n3. Mixed modes: some categories as multi-row, others as header-input...");
        let config = ReportConfig::builder()
            .header(
                ReportHeader::builder()
                    .language(Language::Deutsch)
                    .currency(Currency::EUR)
                    .project_number("MIX-003")
                    .project_title("Gemischte Modi — Kat. 5 Header, Kat. 6-8 Multi-Row")
                    .project_start("01.01.2025")
                    .project_end("31.12.2027")
                    .report_start("01.01.2025")
                    .report_end("30.06.2025")
                    .version("v2025-MIX")
                    .build(),
            )
            .body(
                ReportBody::builder()
                    // ========================================================
                    // Einnahmen-Tabelle (alle 5 Zeilen)
                    // ========================================================
                    .kmw_mittel(   TableEntry::builder().approved_budget(100_000.0).income_report(40_000.0).income_total(40_000.0).reason("1. Tranche").build())
                    .eigenmittel(  TableEntry::builder().approved_budget( 15_000.0).income_report( 7_500.0).income_total( 7_500.0).reason("Eigenanteil").build())
                    .drittmittel(  TableEntry::builder().approved_budget(  5_000.0).income_report( 2_500.0).income_total( 2_500.0).reason("Stiftung ABC").build())
                    .saldovortrag( TableEntry::builder().approved_budget(  3_000.0).income_report( 3_000.0).income_total( 3_000.0).build())
                    .zinsertraege( TableEntry::builder().approved_budget(    200.0).income_report(   100.0).income_total(   100.0).reason("Tagesgeld").build())
                    // ========================================================
                    // Linkes Panel (Belege 1–8)
                    // ========================================================
                    .add_left_panel_entry(PanelEntry::builder().date("10.01.2025").amount_euro(15_000.0).amount_local(15_000.0).build())
                    .add_left_panel_entry(PanelEntry::builder().date("01.02.2025").amount_euro(10_000.0).amount_local(10_000.0).build())
                    .add_left_panel_entry(PanelEntry::builder().date("15.02.2025").amount_euro( 5_000.0).amount_local( 5_000.0).build())
                    .add_left_panel_entry(PanelEntry::builder().date("01.03.2025").amount_euro( 3_000.0).amount_local( 3_000.0).build())
                    .add_left_panel_entry(PanelEntry::builder().date("15.04.2025").amount_euro( 7_500.0).amount_local( 7_500.0).build())
                    .add_left_panel_entry(PanelEntry::builder().date("01.05.2025").amount_euro( 4_000.0).amount_local( 4_000.0).build())
                    .add_left_panel_entry(PanelEntry::builder().date("15.05.2025").amount_euro( 2_500.0).amount_local( 2_500.0).build())
                    .add_left_panel_entry(PanelEntry::builder().date("30.06.2025").amount_euro( 6_100.0).amount_local( 6_100.0).build())
                    // ========================================================
                    // Rechtes Panel (Belege 9–14)
                    // ========================================================
                    .add_right_panel_entry(PanelEntry::builder().date("20.01.2025").amount_euro( 5_000.0).amount_local( 5_000.0).build())
                    .add_right_panel_entry(PanelEntry::builder().date("10.02.2025").amount_euro( 3_500.0).amount_local( 3_500.0).build())
                    .add_right_panel_entry(PanelEntry::builder().date("01.04.2025").amount_euro( 6_000.0).amount_local( 6_000.0).build())
                    .add_right_panel_entry(PanelEntry::builder().date("15.04.2025").amount_euro( 2_000.0).amount_local( 2_000.0).build())
                    .add_right_panel_entry(PanelEntry::builder().date("01.06.2025").amount_euro( 4_500.0).amount_local( 4_500.0).build())
                    .add_right_panel_entry(PanelEntry::builder().date("25.06.2025").amount_euro( 2_100.0).amount_local( 2_100.0).build())
                    // ========================================================
                    // Kategorien 1–4: Multi-Row Positionen
                    // ========================================================
                    .add_position(1, PositionEntry::builder().description("Koordinator/in")    .approved(24_000.0).income_report(12_000.0).income_total(12_000.0).remark("50% Stelle").build())
                    .add_position(1, PositionEntry::builder().description("Buchhalter/in")     .approved( 6_000.0).income_report( 3_000.0).income_total( 3_000.0).remark("Teilzeit").build())
                    .skip_position_row(1)
                    .add_position(1, PositionEntry::builder().description("Freelancer")        .approved( 4_000.0).income_report( 2_000.0).income_total( 2_000.0).remark("Workshops").build())
                    .add_position(2, PositionEntry::builder().description("Inlandsreisen")     .approved( 2_000.0).income_report( 1_000.0).income_total( 1_000.0).build())
                    .add_position(2, PositionEntry::builder().description("Auslandsreisen")    .approved( 5_000.0).income_report( 2_500.0).income_total( 2_500.0).remark("2 Dienstreisen").build())
                    .add_position(3, PositionEntry::builder().description("Büromaterial")      .approved(   500.0).income_report(   250.0).income_total(   250.0).build())
                    .add_position(3, PositionEntry::builder().description("Lehrmaterialien")   .approved( 2_000.0).income_report( 1_000.0).income_total( 1_000.0).remark("Schulbücher").build())
                    .add_position(4, PositionEntry::builder().description("IT-Ausstattung")    .approved( 3_000.0).income_report( 3_000.0).income_total( 3_000.0).remark("4 Laptops").build())
                    .add_position(4, PositionEntry::builder().description("Mobiliar")          .approved( 2_000.0).income_report( 2_000.0).income_total( 2_000.0).remark("Tische + Stühle").build())
                    // ========================================================
                    // Kategorie 5: Header-Input (Gesamtbetrag, keine Zeilen)
                    // ========================================================
                    .set_header_input(5, PositionEntry::builder().approved( 8_000.0).income_report( 4_000.0).income_total( 4_000.0).remark("Direktkosten gesamt").build())
                    // ========================================================
                    // Kategorie 6: Evaluierung — Multi-Row
                    // ========================================================
                    .add_position(6, PositionEntry::builder().description("Zwischenevaluierung").approved( 3_600.0).income_report( 1_800.0).income_total( 1_800.0).remark("Extern").build())
                    .add_position(6, PositionEntry::builder().description("Datenerhebung")     .approved(   480.0).income_report(   240.0).income_total(   240.0).remark("Umfragen").build())
                    // ========================================================
                    // Kategorie 7: Audit — Multi-Row
                    // ========================================================
                    .add_position(7, PositionEntry::builder().description("Jahresabschluss 2025").approved(3_000.0).income_report( 1_500.0).income_total( 1_500.0).remark("Wirtschaftsprüfer").build())
                    // ========================================================
                    // Kategorie 8: Reserve — Multi-Row
                    // ========================================================
                    .add_position(8, PositionEntry::builder().description("Rücklage genehmigt") .approved( 2_000.0).income_report( 1_000.0).income_total( 1_000.0).build())
                    .add_position(8, PositionEntry::builder().description("Notfallfonds")       .approved( 1_500.0).income_report(   750.0).income_total(   750.0).remark("COVID-Reserve").build())
                    .build(),
            )
            .footer(
                ReportFooter::builder()
                    .bank(15_000.0)
                    .kasse(2_500.0)
                    .sonstiges(500.0)
                    .build(),
            )
            .options(
                ReportOptions::builder()
                    .sheet_protection(SheetProtection::from_defaults().with_password("geheim"))
                    .hide_columns_qv(true)
                    .hide_language_sheet(true)
                    .build(),
            )
            .build();

        config.write_to("examples/output/test_mixed_cat_gemischt.xlsx")?;
        println!("   OK: examples/output/test_mixed_cat_gemischt.xlsx");
        println!("   Kat. 1 → 3 Positionen + 1 Skip | Kat. 5 → Header-Input");
        println!("   Kat. 6 → 2 Positionen | Kat. 7 → 1 Position | Kat. 8 → 2 Positionen");
        println!("   B2 → \"v2025-MIX\"");
    }

    println!("\nAll tests completed successfully!");
    Ok(())
}
