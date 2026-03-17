//! API-Validierung: IronCalc als Formel-Engine für den Report-Generator
//!
//! Prüft alle kritischen IronCalc-Funktionen, die wir brauchen:
//! - VLOOKUP mit externem Sheet
//! - IFERROR, SUM, IF
//! - to_bytes/from_bytes Template-Pattern
//! - Formel-Strings zurücklesen

use ironcalc_base::cell::CellValue;
use ironcalc_base::Model;

/// Erstellt ein Basis-Modell mit 2 Sheets
fn create_test_model() -> Model<'static> {
    let mut model =
        Model::new_empty("Bericht", "en", "UTC", "en").expect("Failed to create model");

    // Sheet 1: "Sprachversionen" hinzufügen
    model.add_sheet("Sprachversionen").unwrap();

    model
}

#[test]
fn test_basic_formula_evaluation() {
    let mut model = create_test_model();

    // Werte setzen
    model.set_user_input(0, 1, 1, "100".to_string()).unwrap(); // A1 = 100
    model.set_user_input(0, 2, 1, "200".to_string()).unwrap(); // A2 = 200

    // SUM-Formel
    model
        .update_cell_with_formula(0, 3, 1, "=SUM(A1:A2)".to_string())
        .unwrap();

    model.evaluate();

    let result = model.get_cell_value_by_index(0, 3, 1).unwrap();
    assert_eq!(result, CellValue::Number(300.0));
}

#[test]
fn test_iferror_formula() {
    let mut model = create_test_model();

    model.set_user_input(0, 1, 1, "10".to_string()).unwrap(); // A1 = 10
    model.set_user_input(0, 2, 1, "0".to_string()).unwrap(); // A2 = 0

    // IFERROR(A1/A2, 0) — Division by zero → 0
    model
        .update_cell_with_formula(0, 3, 1, "=IFERROR(A1/A2,0)".to_string())
        .unwrap();

    // IFERROR(A1/5, 0) — Normal division → 2
    model
        .update_cell_with_formula(0, 4, 1, "=IFERROR(A1/5,0)".to_string())
        .unwrap();

    model.evaluate();

    let div_zero = model.get_cell_value_by_index(0, 3, 1).unwrap();
    assert_eq!(div_zero, CellValue::Number(0.0));

    let normal = model.get_cell_value_by_index(0, 4, 1).unwrap();
    assert_eq!(normal, CellValue::Number(2.0));
}

#[test]
fn test_if_formula_with_string_comparison() {
    let mut model = create_test_model();

    model
        .set_user_input(0, 1, 5, "deutsch".to_string())
        .unwrap(); // E1 = "deutsch"

    // IF(E1="","empty","found")
    model
        .update_cell_with_formula(0, 2, 1, r#"=IF(E1="","empty","found")"#.to_string())
        .unwrap();

    model.evaluate();

    let result = model.get_cell_value_by_index(0, 2, 1).unwrap();
    assert_eq!(result, CellValue::String("found".to_string()));
}

#[test]
fn test_if_formula_empty_string() {
    let mut model = create_test_model();

    // E2 bleibt leer (nicht gesetzt)
    // IF($E$2="","","something")
    model
        .update_cell_with_formula(0, 1, 1, r#"=IF($E$2="","","something")"#.to_string())
        .unwrap();

    model.evaluate();

    let result = model.get_cell_value_by_index(0, 1, 1).unwrap();
    // Leere Zelle → IF-Bedingung true → leerer String
    assert_eq!(result, CellValue::String("".to_string()));
}

#[test]
fn test_vlookup_with_external_sheet() {
    let mut model = create_test_model();

    // Sheet "Sprachversionen" (Index 1) befüllen
    model
        .set_user_input(1, 1, 2, "deutsch".to_string())
        .unwrap(); // B1
    model
        .set_user_input(1, 1, 3, "FINANZBERICHT".to_string())
        .unwrap(); // C1 (VLOOKUP index 2)
    model
        .set_user_input(1, 1, 4, "ÜBERSICHT".to_string())
        .unwrap(); // D1 (VLOOKUP index 3)

    // Row 2: english
    model
        .set_user_input(1, 2, 2, "english".to_string())
        .unwrap(); // B2
    model
        .set_user_input(1, 2, 3, "FINANCIAL REPORT".to_string())
        .unwrap(); // C2
    model
        .set_user_input(1, 2, 4, "OVERVIEW".to_string())
        .unwrap(); // D2

    // Sheet "Bericht" (Index 0): Sprache setzen
    model
        .set_user_input(0, 2, 5, "deutsch".to_string())
        .unwrap(); // E2 = "deutsch"

    // VLOOKUP-Formel: Suche Sprache in Sprachversionen, gib Index 2 zurück
    model
        .update_cell_with_formula(
            0,
            1,
            2,
            "=VLOOKUP($E$2,Sprachversionen!$B$1:$D$5,2,FALSE)".to_string(),
        )
        .unwrap();

    model.evaluate();

    let result = model.get_cell_value_by_index(0, 1, 2).unwrap();
    assert_eq!(
        result,
        CellValue::String("FINANZBERICHT".to_string()),
        "VLOOKUP should find 'deutsch' and return index 2 text"
    );
}

#[test]
fn test_vlookup_full_column_range() {
    let mut model = create_test_model();

    // Sprachversionen befüllen
    model
        .set_user_input(1, 1, 2, "deutsch".to_string())
        .unwrap();
    model
        .set_user_input(1, 1, 3, "FINANZBERICHT".to_string())
        .unwrap();
    model
        .set_user_input(1, 2, 2, "english".to_string())
        .unwrap();
    model
        .set_user_input(1, 2, 3, "FINANCIAL REPORT".to_string())
        .unwrap();

    model
        .set_user_input(0, 2, 5, "english".to_string())
        .unwrap(); // E2

    // Test mit ganzer Spalte $B:$C (ohne explizite Zeilen)
    model
        .update_cell_with_formula(
            0,
            1,
            1,
            "=VLOOKUP($E$2,Sprachversionen!$B:$C,2,FALSE)".to_string(),
        )
        .unwrap();

    model.evaluate();

    let result = model.get_cell_value_by_index(0, 1, 1).unwrap();
    assert_eq!(
        result,
        CellValue::String("FINANCIAL REPORT".to_string()),
        "VLOOKUP with full column range should work"
    );
}

#[test]
fn test_if_vlookup_combined() {
    let mut model = create_test_model();

    // Sprachversionen
    model
        .set_user_input(1, 1, 2, "deutsch".to_string())
        .unwrap();
    model
        .set_user_input(1, 1, 3, "FINANZBERICHT".to_string())
        .unwrap();

    model
        .set_user_input(0, 2, 5, "deutsch".to_string())
        .unwrap(); // E2

    // Die tatsächliche Formel aus dem Projekt
    model
        .update_cell_with_formula(
            0,
            1,
            2,
            r#"=IF($E$2="","",VLOOKUP($E$2,Sprachversionen!$B:$C,2,FALSE))"#.to_string(),
        )
        .unwrap();

    model.evaluate();

    let result = model.get_cell_value_by_index(0, 1, 2).unwrap();
    assert_eq!(result, CellValue::String("FINANZBERICHT".to_string()));
}

#[test]
fn test_get_cell_formula_readback() {
    let mut model = create_test_model();

    let formula = "=SUM(A1:A5)";
    model
        .update_cell_with_formula(0, 6, 1, formula.to_string())
        .unwrap();

    // Formel zurücklesen — returns Result<Option<String>, String>
    let readback = model.get_cell_formula(0, 6, 1).unwrap();
    assert_eq!(
        readback.as_deref(),
        Some(formula),
        "Formula should be readable back"
    );
}

#[test]
fn test_template_pattern_to_from_bytes() {
    // 1. Master-Model erstellen
    let mut master = create_test_model();

    // Sprachversionen befüllen
    master
        .set_user_input(1, 1, 2, "deutsch".to_string())
        .unwrap();
    master
        .set_user_input(1, 1, 3, "FINANZBERICHT".to_string())
        .unwrap();

    // Statische Formel
    master
        .update_cell_with_formula(
            0,
            1,
            2,
            r#"=IF($E$2="","",VLOOKUP($E$2,Sprachversionen!$B:$C,2,FALSE))"#.to_string(),
        )
        .unwrap();
    master
        .update_cell_with_formula(0, 3, 1, "=SUM(A1:A2)".to_string())
        .unwrap();

    // 2. Serialisieren
    let bytes = master.to_bytes();
    assert!(!bytes.is_empty(), "Template bytes should not be empty");

    // 3. Deserialisieren → neues Model
    let mut instance =
        Model::from_bytes(&bytes, "en").expect("Failed to deserialize template");

    // 4. Inputs setzen
    instance
        .set_user_input(0, 2, 5, "deutsch".to_string())
        .unwrap(); // E2
    instance
        .set_user_input(0, 1, 1, "100".to_string())
        .unwrap(); // A1
    instance
        .set_user_input(0, 2, 1, "50".to_string())
        .unwrap(); // A2

    // 5. Evaluieren
    instance.evaluate();

    // 6. Ergebnisse prüfen
    let vlookup = instance.get_cell_value_by_index(0, 1, 2).unwrap();
    assert_eq!(vlookup, CellValue::String("FINANZBERICHT".to_string()));

    let sum = instance.get_cell_value_by_index(0, 3, 1).unwrap();
    assert_eq!(sum, CellValue::Number(150.0));
}

#[test]
fn test_round_formula() {
    let mut model = create_test_model();

    model
        .set_user_input(0, 1, 1, "1.23456".to_string())
        .unwrap();
    model
        .update_cell_with_formula(0, 2, 1, "=ROUND(A1,2)".to_string())
        .unwrap();

    model.evaluate();

    let result = model.get_cell_value_by_index(0, 2, 1).unwrap();
    assert_eq!(result, CellValue::Number(1.23));
}

#[test]
fn test_complex_check_formula() {
    let mut model = create_test_model();

    // Simuliere: E_saldo, F_income, F_total
    model
        .set_user_input(0, 1, 5, "100".to_string())
        .unwrap(); // E1 = 100 (saldo)
    model
        .set_user_input(0, 2, 6, "500".to_string())
        .unwrap(); // F2 = 500 (income)
    model
        .set_user_input(0, 3, 6, "400".to_string())
        .unwrap(); // F3 = 400 (total)

    // Check: IF(ROUND(E1,2)=ROUND(F2-F3,2),"✓","")
    model
        .update_cell_with_formula(
            0,
            4,
            4,
            r#"=IF(ROUND(E1,2)=ROUND(F2-F3,2),"✓","")"#.to_string(),
        )
        .unwrap();

    model.evaluate();

    let result = model.get_cell_value_by_index(0, 4, 4).unwrap();
    assert_eq!(result, CellValue::String("✓".to_string()));
}
