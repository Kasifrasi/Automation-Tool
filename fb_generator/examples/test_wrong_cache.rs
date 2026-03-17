//! Test: Echter Finanzbericht mit absichtlich FALSCHEN Formel-Cache-Werten
//!
//! Workflow:
//! 1. Generiert einen vollständigen Finanzbericht über die normale API
//! 2. Öffnet die erzeugte .xlsx (ZIP), findet das Sheet-XML
//! 3. Ersetzt alle `<v>...</v>` in Formelzellen durch falsche Werte
//! 4. Speichert die korrumpierte Version separat
//!
//! Damit lässt sich testen, ob Excel beim Öffnen automatisch
//! die Formeln neu berechnet und die korrekten Werte anzeigt.

use fb_generator::*;
use quick_xml::events::{BytesText, Event};
use quick_xml::reader::Reader;
use quick_xml::writer::Writer;
use std::fs::File;
use std::io::{Cursor, Read, Write};
use zip::write::FileOptions;
use zip::{ZipArchive, ZipWriter};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let normal_path = "examples/output/test_wrong_cache_normal.xlsx";
    let corrupt_path = "examples/output/test_wrong_cache_corrupt.xlsx";

    // ========================================================================
    // 1. Echten Finanzbericht generieren (wie test_all_fields)
    // ========================================================================
    println!("1. Generiere normalen Finanzbericht...");

    let config = ReportConfig::builder()
        .header(
            ReportHeader::builder()
                .language(Language::Deutsch)
                .currency(Currency::EUR)
                .project_number("CACHE-TEST-001")
                .project_title("Cache-Test: Falscher Formel-Cache")
                .project_start("01.01.2024")
                .project_end("31.12.2026")
                .report_start("01.01.2024")
                .report_end("30.06.2024")
                .build(),
        )
        .body(
            ReportBody::builder()
                .kmw_mittel(TableEntry::builder().approved_budget(80_000.0).income_report(50_000.0).income_total(50_000.0).reason("1. Rate").build())
                .eigenmittel(TableEntry::builder().approved_budget(15_000.0).income_report(7_500.0).income_total(7_500.0).reason("Eigenanteil").build())
                .drittmittel(TableEntry::builder().approved_budget(10_000.0).income_report(5_000.0).income_total(5_000.0).reason("Stiftung").build())
                .add_left_panel_entry(PanelEntry::builder().date("15.01.2024").amount_euro(9_000.0).amount_local(9_000.0).build())
                .add_left_panel_entry(PanelEntry::builder().date("15.02.2024").amount_euro(6_000.0).amount_local(6_000.0).build())
                .add_right_panel_entry(PanelEntry::builder().date("01.03.2024").amount_euro(3_000.0).amount_local(3_000.0).build())
                .add_position(1, PositionEntry::builder().description("Projektleitung").approved(18_000.0).income_report(9_000.0).income_total(9_000.0).remark("6 Monate").build())
                .add_position(1, PositionEntry::builder().description("Buchhaltung").approved(12_000.0).income_report(6_000.0).income_total(6_000.0).remark("Teilzeit").build())
                .add_position(2, PositionEntry::builder().description("Reisekosten").approved(5_000.0).income_report(2_500.0).income_total(2_500.0).remark("2 Reisen").build())
                .add_position(3, PositionEntry::builder().description("Büromaterial").approved(2_000.0).income_report(1_000.0).income_total(1_000.0).remark("Verbrauch").build())
                .set_header_input(6, PositionEntry::builder().approved(8_000.0).income_report(4_000.0).income_total(4_000.0).remark("Verwaltung").build())
                .build(),
        )
        .footer(
            ReportFooter::builder()
                .bank(8_500.0)
                .kasse(1_200.0)
                .sonstiges(300.0)
                .build(),
        )
        .options(ReportOptions::builder().build())
        .build();

    config.write_to(normal_path)?;
    println!("   Normaler Bericht: {normal_path}");

    // ========================================================================
    // 2. xlsx öffnen und Formel-Caches korrumpieren
    // ========================================================================
    println!("2. Korrumpiere Formel-Cache-Werte...");

    let corrupted = corrupt_formula_caches(normal_path)?;
    println!("   {corrupted} Formel-Caches durch falsche Werte ersetzt");

    println!("3. Fertig!");
    println!();
    println!("   Normal:     {normal_path}");
    println!("   Korrumpiert: {corrupt_path}");
    println!();
    println!("Beide Dateien in Excel öffnen und vergleichen:");
    println!("- Normal: alle Werte korrekt");
    println!("- Korrumpiert: Formelzellen zeigen '99999.99'");
    println!("  → Wenn Excel automatisch neu berechnet, werden die Werte korrigiert");

    Ok(())
}

/// Öffnet eine .xlsx, ersetzt alle Formel-Cache-Werte mit '99999.99',
/// speichert das Ergebnis als `*_corrupt.xlsx`.
fn corrupt_formula_caches(input_path: &str) -> Result<usize, Box<dyn std::error::Error>> {
    let output_path = input_path.replace("_normal.xlsx", "_corrupt.xlsx");

    let file = File::open(input_path)?;
    let mut archive = ZipArchive::new(file)?;

    let out_file = File::create(&output_path)?;
    let mut zip_writer = ZipWriter::new(out_file);

    let mut corrupted_count = 0;

    for i in 0..archive.len() {
        let name = archive.by_index(i)?.name().to_string();

        if name.starts_with("xl/worksheets/sheet") && name.ends_with(".xml") {
            // Sheet-XML lesen, Formel-Caches korrumpieren, schreiben
            let mut entry = archive.by_index(i)?;
            let compression = entry.compression();
            let unix_mode = entry.unix_mode();
            let mut content = Vec::new();
            entry.read_to_end(&mut content)?;

            let (new_xml, count) = corrupt_sheet_xml(&content)?;
            corrupted_count += count;

            let options = FileOptions::<()>::default()
                .compression_method(compression)
                .unix_permissions(unix_mode.unwrap_or(0o644));
            zip_writer.start_file(&name, options)?;
            zip_writer.write_all(&new_xml)?;
        } else {
            // Alle anderen Einträge 1:1 kopieren
            let entry = archive.by_index(i)?;
            zip_writer.raw_copy_file(entry)?;
        }
    }

    zip_writer.finish()?;
    Ok(corrupted_count)
}

/// Parst Sheet-XML mit quick_xml, ersetzt <v>-Inhalte in Formelzellen.
///
/// Logik: Wenn ein <c>-Element ein <f>-Kind hat, wird der <v>-Text
/// durch "99999.99" ersetzt.
fn corrupt_sheet_xml(xml: &[u8]) -> Result<(Vec<u8>, usize), Box<dyn std::error::Error>> {
    let mut reader = Reader::from_reader(xml);
    let mut writer = Writer::new(Cursor::new(Vec::new()));
    let mut buf = Vec::new();

    // State: sind wir in einem <c> mit einer <f>?
    let mut in_cell = false;
    let mut cell_has_formula = false;
    let mut in_value = false; // sind wir zwischen <v> und </v>?
    let mut corrupted_count = 0usize;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) => {
                let tag = e.name();
                if tag.as_ref() == b"c" {
                    in_cell = true;
                    cell_has_formula = false;
                } else if tag.as_ref() == b"f" && in_cell {
                    cell_has_formula = true;
                } else if tag.as_ref() == b"v" && in_cell && cell_has_formula {
                    in_value = true;
                }
                writer.write_event(Event::Start(e.clone()))?;
            }
            Ok(Event::End(ref e)) => {
                let tag = e.name();
                if tag.as_ref() == b"v" {
                    in_value = false;
                } else if tag.as_ref() == b"c" {
                    in_cell = false;
                    cell_has_formula = false;
                }
                writer.write_event(Event::End(e.clone()))?;
            }
            Ok(Event::Text(ref t)) => {
                if in_value && cell_has_formula {
                    // Cache-Wert durch falschen Wert ersetzen
                    let wrong = BytesText::new("99999.99");
                    writer.write_event(Event::Text(wrong))?;
                    corrupted_count += 1;
                } else {
                    writer.write_event(Event::Text(t.clone()))?;
                }
            }
            Ok(Event::Eof) => break,
            Ok(other) => writer.write_event(other)?,
            Err(e) => return Err(Box::new(e)),
        }
        buf.clear();
    }

    let result = writer.into_inner().into_inner();
    Ok((result, corrupted_count))
}
