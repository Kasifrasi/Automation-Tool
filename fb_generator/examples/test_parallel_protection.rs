//! Paralleler Workbook-Protection Benchmark
//!
//! Generiert eine Basis-Datei, erstellt Kopien und schützt sie parallel mit Rayon.
//!
//! **Builder-Pattern** — kein `Some()`, kein `.to_string()`, kein `?` beim Bauen.
//!
//! Usage:
//!   cargo run --example test_parallel_protection --release

use fb_generator::*;
use rayon::prelude::*;
use std::fs;
use std::path::PathBuf;
use std::time::Instant;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let file_count = 100;
    let dir = tempfile::TempDir::new()?;
    let input_dir = dir.path().join("input");
    let result_dir = dir.path().join("locked");
    fs::create_dir_all(&input_dir)?;
    fs::create_dir_all(&result_dir)?;

    // 1. Basis-Datei generieren
    let base_file = dir.path().join("base.xlsx");
    let config = ReportConfig::builder()
        .header(
            ReportHeader::builder()
                .language(Language::Deutsch)
                .currency(Currency::EUR)
                .project_number("BENCH-001")
                .project_title("Protection Benchmark")
                .build(),
        )
        .options(
            ReportOptions::builder()
                .sheet_protection(SheetProtection::from_defaults())
                .build(),
        )
        .build();
    config.write_to(&base_file)?;
    println!("Basis-Datei generiert: {:?}", base_file);

    // 2. Kopien erstellen
    println!("Erstelle {} Kopien...", file_count);
    let files: Vec<(PathBuf, PathBuf)> = (0..file_count)
        .map(|i| {
            let inp = input_dir.join(format!("bench_{:03}.xlsx", i));
            let out = result_dir.join(format!("LOCKED_{:03}.xlsx", i));
            fs::copy(&base_file, &inp).unwrap();
            (inp, out)
        })
        .collect();

    // 3. Paralleler Schutz mit Rayon
    println!("Starte parallelen Schutz fuer {} Dateien...", file_count);
    let start = Instant::now();

    files.par_iter().for_each(|(inp, out)| {
        if let Err(e) = protect_workbook(inp.to_str().unwrap(), out.to_str().unwrap(), "geheim123")
        {
            eprintln!("Fehler bei {:?}: {}", inp, e);
        }
    });

    let duration = start.elapsed();
    let avg_per_file = duration.as_secs_f64() * 1000.0 / file_count as f64;
    let throughput = file_count as f64 / duration.as_secs_f64();

    println!("---------------------------------------------------");
    println!("ERGEBNIS ({} Dateien, Parallel mit Rayon)", file_count);
    println!("Gesamtzeit:      {:.2?}", duration);
    println!("Durchschnitt:    {:.2} ms / Datei", avg_per_file);
    println!("Durchsatz:       {:.0} Dateien / Sekunde", throughput);
    println!("---------------------------------------------------");

    Ok(())
}
