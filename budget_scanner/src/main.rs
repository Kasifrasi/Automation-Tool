use calamine::{open_workbook_auto, Data, Reader};
use regex::Regex;
use std::env;
use std::path::Path;
use walkdir::WalkDir;

const SHEET_NAMES: &[&str] = &["Budget", "Presupuesto", "Plano de custos e financiamento"];
const STOP_WORDS: &[&str] = &["Summe", "Total", "TOTAL"];
const MAX_EMPTY_ROWS: usize = 100;

fn process_file(path: &Path, re: &Regex) {
    println!("\n=== {} ===", path.display());

    let mut wb = match open_workbook_auto(path) {
        Ok(wb) => wb,
        Err(e) => {
            eprintln!("  Fehler beim Öffnen: {}", e);
            return;
        }
    };

    let sheet_names = wb.sheet_names();

    let sheet_name = match SHEET_NAMES
        .iter()
        .find(|&&name| sheet_names.iter().any(|s| s == name))
        .copied()
    {
        Some(name) => {
            println!("  Sheet '{}' gefunden.", name);
            name
        }
        None => {
            println!("  Kein passendes Sheet gefunden. Vorhandene: {}", sheet_names.join(", "));
            return;
        }
    };

    let range = match wb.worksheet_range(sheet_name) {
        Ok(r) => r,
        Err(e) => {
            eprintln!("  Fehler beim Lesen des Sheets: {}", e);
            return;
        }
    };

    match range.get((1, 0)) {
        Some(cell) => println!("  A2: {}", cell),
        None => println!("  A2 ist leer."),
    }

    println!("  Gefundene Einträge in Spalte A:");

    let mut first_match_found = false;
    let mut empty_streak = 0usize;

    for row in range.rows() {
        let cell = &row[0];
        let text = match cell {
            Data::String(s) => s.clone(),
            Data::Float(f) => f.to_string(),
            Data::Int(i) => i.to_string(),
            _ => {
                if first_match_found {
                    empty_streak += 1;
                    if empty_streak >= MAX_EMPTY_ROWS {
                        break;
                    }
                }
                continue;
            }
        };

        let trimmed = text.trim();

        if trimmed.is_empty() {
            if first_match_found {
                empty_streak += 1;
                if empty_streak >= MAX_EMPTY_ROWS {
                    break;
                }
            }
            continue;
        }

        if STOP_WORDS.iter().any(|&w| trimmed.contains(w)) {
            break;
        }

        if re.is_match(trimmed) {
            first_match_found = true;
            empty_streak = 0;
            println!("    {}", trimmed);
        }
    }
}

fn main() {
    let args: Vec<String> = env::args().collect();
    if args.len() < 2 {
        eprintln!("Usage: {} <file_or_directory>", args[0]);
        std::process::exit(1);
    }

    let input = Path::new(&args[1]);
    let re = Regex::new(r"^[1-8]\.\d*").unwrap();

    if input.is_file() {
        process_file(input, &re);
    } else if input.is_dir() {
        let entries: Vec<_> = WalkDir::new(input)
            .into_iter()
            .filter_map(|e| e.ok())
            .filter(|e| {
                e.file_type().is_file()
                    && matches!(
                        e.path().extension().and_then(|s| s.to_str()),
                        Some("xlsx") | Some("xlsm")
                    )
            })
            .collect();

        if entries.is_empty() {
            println!("Keine .xlsx/.xlsm Dateien in '{}' gefunden.", input.display());
            return;
        }

        println!("{} Datei(en) gefunden.", entries.len());
        for entry in entries {
            process_file(entry.path(), &re);
        }
    } else {
        eprintln!("'{}' ist weder eine Datei noch ein Verzeichnis.", input.display());
        std::process::exit(1);
    }
}
