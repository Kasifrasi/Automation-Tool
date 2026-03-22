use calamine::{open_workbook_auto, Reader};
use std::env;

fn main() {
    let args: Vec<String> = env::args().collect();
    if args.len() < 2 {
        eprintln!("Usage: {} <path_to_xlsx_or_xlsm>", args[0]);
        std::process::exit(1);
    }

    let path = &args[1];
    let workbook = open_workbook_auto(path);

    match workbook {
        Err(e) => {
            eprintln!("Fehler beim Öffnen der Datei '{}': {}", path, e);
            std::process::exit(1);
        }
        Ok(mut wb) => {
            let sheet_names = wb.sheet_names();
            if sheet_names.iter().any(|s| s == "Budget") {
                println!("Sheet 'Budget' existiert in der Datei.");
                match wb.worksheet_range("Budget") {
                    Ok(range) => match range.get((1, 0)) {
                        Some(cell) => println!("A2: {}", cell),
                        None => println!("A2 ist leer."),
                    },
                    Err(e) => eprintln!("Fehler beim Lesen des Sheets: {}", e),
                }
            } else {
                println!("Sheet 'Budget' wurde NICHT gefunden.");
                println!("Vorhandene Sheets: {}", sheet_names.join(", "));
            }
        }
    }
}
