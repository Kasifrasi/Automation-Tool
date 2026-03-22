use calamine::{open_workbook_auto, Data, Range, Reader};
use comfy_table::{Attribute, Cell, ContentArrangement, Table};
use regex::Regex;
use std::env;
use std::path::Path;
use walkdir::WalkDir;

const SHEET_NAMES: &[&str] = &["Budget", "Presupuesto", "Plano de custos e financiamento"];
const STOP_WORDS: &[&str] = &["Summe", "Total", "TOTAL"];
const COST_TERMS: &[&str] = &[
    "Gesamtkosten",
    "Total Costs",
    "Coût total",
    "Gastos total",
    "Custos total",
];
const MAX_EMPTY_ROWS: usize = 100;

fn cell_text(cell: &Data) -> Option<String> {
    match cell {
        Data::String(s) => Some(s.clone()),
        Data::Float(f) => Some(f.to_string()),
        Data::Int(i) => Some(i.to_string()),
        _ => None,
    }
}

fn is_exact_cost_term(text: &str) -> bool {
    COST_TERMS.iter().any(|&t| text.trim() == t)
}

/// Sucht die erste und zweite Spalte, die exakt einen der Kostenbegriffe enthalten.
/// Standard: Spalte I (8) für die erste, Spalte N (13) für die zweite.
/// Fallback: A1:Z100 wird nach dem 1. und 2. Vorkommen durchsucht.
fn find_cost_columns(range: &Range<Data>) -> (Option<usize>, Option<usize>) {
    let row_count = range.rows().count().min(100);

    let col_contains_term = |col: usize| -> bool {
        for row in range.rows() {
            if let Some(cell) = row.get(col) {
                if let Some(text) = cell_text(cell) {
                    if is_exact_cost_term(&text) {
                        return true;
                    }
                }
            }
        }
        false
    };

    let first_col = if col_contains_term(8) { Some(8) } else { None };
    let second_col = if col_contains_term(13) { Some(13) } else { None };

    if first_col.is_some() && second_col.is_some() {
        return (first_col, second_col);
    }

    // Fallback: A1:Z100 nach dem 1. und 2. Vorkommen scannen
    let mut found: Vec<usize> = Vec::new();
    'outer: for row_idx in 0..row_count {
        for col_idx in 0usize..26 {
            if let Some(cell) = range.get((row_idx, col_idx)) {
                if let Some(text) = cell_text(cell) {
                    if is_exact_cost_term(&text) {
                        if !found.contains(&col_idx) {
                            found.push(col_idx);
                        }
                        if found.len() == 2 {
                            break 'outer;
                        }
                    }
                }
            }
        }
    }

    let resolved_first = first_col.or_else(|| found.first().copied());
    let resolved_second = second_col.or_else(|| {
        found.get(1).copied().filter(|&col| {
            resolved_first.map_or(true, |fc| col > fc)
        })
    });

    (resolved_first, resolved_second)
}

fn col_to_letter(col: usize) -> char {
    (b'A' + col as u8) as char
}

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

    let get = |row: usize, col: usize| -> String {
        range.get((row, col)).and_then(|c| cell_text(c)).unwrap_or_default()
    };

    println!("  Projekttitel:    {}", get(1, 2));  // C2
    println!("  Projektnummer:   {}", get(1, 8));  // I2
    println!("  Sprache:         {}", get(2, 8));  // I3
    println!("  Lokalwährung:    {}", get(3, 8));  // I4

    let (col1, col2) = find_cost_columns(&range);

    let col1_label = col1.map(|c| col_to_letter(c).to_string()).unwrap_or("-".into());
    let col2_label = col2.map(|c| col_to_letter(c).to_string()).unwrap_or("-".into());
    println!("  Kostenspalten: {} / {}", col1_label, col2_label);

    let mut table = Table::new();
    table.set_content_arrangement(ContentArrangement::Dynamic);
    table.set_header(vec![
        Cell::new("Nr.").add_attribute(Attribute::Bold),
        Cell::new("Bezeichnung").add_attribute(Attribute::Bold),
        Cell::new(&col1_label).add_attribute(Attribute::Bold),
        Cell::new(&col2_label).add_attribute(Attribute::Bold),
    ]);

    let mut first_match_found = false;
    let mut empty_streak = 0usize;

    for row in range.rows() {
        let cell = &row[0];
        let text = match cell_text(cell) {
            Some(t) => t,
            None => {
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

        if let Some(m) = re.find(trimmed) {
            first_match_found = true;
            empty_streak = 0;

            let number = m.as_str().to_string();
            let label = row.get(1).and_then(|c| cell_text(c)).unwrap_or_default();
            let v1 = col1.and_then(|c| row.get(c)).and_then(|c| cell_text(c)).unwrap_or_default();
            let v2 = col2.and_then(|c| row.get(c)).and_then(|c| cell_text(c)).unwrap_or_default();

            table.add_row(vec![number, label, v1, v2]);
        }
    }

    println!("{table}");
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
