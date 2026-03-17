//! Profiling: Zeitverteilung pro Phase der Report-Generierung
//!
//! Usage: cargo run --example profile --release

use fb_generator::lang::build_sheet as build_sprachversionen;
use fb_generator::lang::LANG_CONFIG;
use fb_generator::report::writer::setup_sheet;
use fb_generator::report::ApiKey;
use fb_generator::{write_report_with_options, BodyConfig, SheetOptions, ReportValues};
use rust_xlsxwriter::{Format, Workbook};
use std::time::Instant;

const USED_COLUMNS: u16 = 22;
const ITERATIONS: usize = 100;

fn build_test_values(index: usize) -> (ReportValues, BodyConfig) {
    let mut values = ReportValues::new();
    let config = LANG_CONFIG.get("Deutsch").unwrap();

    values.set(ApiKey::Language, config.lang_val);
    values.set(ApiKey::Currency, "EUR");
    values.set(ApiKey::ProjectNumber, format!("PROJ-{:05}", index));
    values.set(
        ApiKey::ProjectTitle,
        format!("Project {}", index),
    );
    values.set(ApiKey::ProjectStart, "01.01.2024");
    values.set(ApiKey::ProjectEnd, "31.12.2024");
    values.set(ApiKey::ReportStart, "01.01.2024");
    values.set(ApiKey::ReportEnd, "30.06.2024");

    let base_budget = 10000.0 + (index as f64 * 100.0);
    for i in 0..5u8 {
        values.set(ApiKey::ApprovedBudget(i), base_budget * (i + 1) as f64);
        values.set(ApiKey::IncomeReportPeriod(i), base_budget * 0.5 * (i + 1) as f64);
        values.set(ApiKey::IncomeTotal(i), base_budget * 0.5 * (i + 1) as f64);
    }

    let pos_offset = (index % 5) as u16;
    let body_config = BodyConfig::new()
        .with_positions(1, 5 + pos_offset)
        .with_positions(2, 3 + pos_offset)
        .with_positions(3, 4 + pos_offset)
        .with_positions(4, 3 + pos_offset)
        .with_positions(5, 2 + pos_offset)
        .with_positions(6, 0)
        .with_positions(7, 0)
        .with_positions(8, 0);

    let base_cost = 1000.0 + (index as f64 * 10.0);
    for cat in 1..=5u8 {
        let num_pos = match cat {
            1 => 5 + pos_offset,
            2 => 3 + pos_offset,
            3 => 4 + pos_offset,
            4 => 3 + pos_offset,
            5 => 2 + pos_offset,
            _ => 0,
        };
        for pos in 1..=num_pos {
            let cost = base_cost * (cat as f64) * (pos as f64 / 2.0);
            values.set_position_row(cat, pos, format!("Pos {}.{}", cat, pos), cost, cost * 0.5, cost * 0.5, "");
        }
    }
    values.set_header_input(6, base_cost * 4.0, base_cost * 2.0, base_cost * 2.0, "");
    values.set_header_input(7, base_cost * 6.0, base_cost * 3.0, base_cost * 3.0, "");
    values.set_header_input(8, base_cost * 2.5, base_cost * 1.25, base_cost * 1.25, "");
    values.set_footer_bank(base_cost * 10.0);
    values.set_footer_kasse(base_cost * 2.0);
    values.set_footer_sonstiges(base_cost * 0.5);

    (values, body_config)
}

fn main() {
    println!("\nPROFILING: {} Reports, Zeitmessung pro Phase\n", ITERATIONS);

    let mut t_workbook_setup = std::time::Duration::ZERO;
    let mut t_lang_sheet = std::time::Duration::ZERO;
    let mut t_ws_setup = std::time::Duration::ZERO;
    let mut t_col_format = std::time::Duration::ZERO;
    let mut t_values = std::time::Duration::ZERO;
    let mut t_write_report = std::time::Duration::ZERO;
    let mut t_save_buffer = std::time::Duration::ZERO;
    let mut t_write_disk = std::time::Duration::ZERO;

    let temp_dir = std::env::temp_dir().join("finanzbericht_profile");
    std::fs::create_dir_all(&temp_dir).unwrap();

    // Baseline: Report-Generierung ohne Protection
    let total_start = Instant::now();
    for i in 0..ITERATIONS {
        let t = Instant::now();
        let mut workbook = Workbook::new();
        t_workbook_setup += t.elapsed();

        let t = Instant::now();
        build_sprachversionen(&mut workbook).unwrap();
        t_lang_sheet += t.elapsed();

        let t = Instant::now();
        let config = LANG_CONFIG.get("Deutsch").unwrap();
        let suffix = "";
        let ws = workbook.add_worksheet();
        ws.set_name(config.fb_sheet).unwrap();
        setup_sheet(ws).unwrap();
        t_ws_setup += t.elapsed();

        let t = Instant::now();
        let unlocked = Format::new().set_font_name("Arial").set_font_size(10.0).set_unlocked();
        for col in 0..USED_COLUMNS {
            ws.set_column_format(col, &unlocked).ok();
        }
        t_col_format += t.elapsed();

        let t = Instant::now();
        let (values, body_config) = build_test_values(i);
        t_values += t.elapsed();

        let t = Instant::now();
        let options = SheetOptions::with_default_protection().with_hidden_columns_qv();
        write_report_with_options(ws, suffix, &values, &body_config, &options).unwrap();
        t_write_report += t.elapsed();

        let t = Instant::now();
        let buffer = workbook.save_to_buffer().unwrap();
        t_save_buffer += t.elapsed();

        let t = Instant::now();
        let filename = temp_dir.join(format!("report_{:05}.xlsx", i));
        std::fs::write(filename, buffer).ok();
        t_write_disk += t.elapsed();
    }
    let total_baseline = total_start.elapsed();

    println!("=== Baseline: Report ohne Protection ===");
    println!("  Gesamt:           {:>8.2?}  ({:.1} Reports/s)", total_baseline, ITERATIONS as f64 / total_baseline.as_secs_f64());
    print_phase("Workbook::new()", t_workbook_setup, total_baseline);
    print_phase("build_sprachversionen()", t_lang_sheet, total_baseline);
    print_phase("Worksheet Setup", t_ws_setup, total_baseline);
    print_phase("Column Formats", t_col_format, total_baseline);
    print_phase("Build Values", t_values, total_baseline);
    print_phase("write_report_with_options()", t_write_report, total_baseline);
    print_phase("save_to_buffer()", t_save_buffer, total_baseline);
    print_phase("fs::write() to disk", t_write_disk, total_baseline);

    // Cleanup
    std::fs::remove_dir_all(&temp_dir).ok();
    std::fs::create_dir_all(&temp_dir).unwrap();

    // Variante C: Mit Workbook-Protection (SHA-512 Hashing)
    let mut t_wb_setup = std::time::Duration::ZERO;
    let mut t_wb_write_report = std::time::Duration::ZERO;
    let mut t_wb_save = std::time::Duration::ZERO;
    let mut t_wb_protect = std::time::Duration::ZERO;
    let mut t_wb_cleanup = std::time::Duration::ZERO;

    let total_start = Instant::now();
    for i in 0..ITERATIONS {
        let t = Instant::now();
        let mut workbook = Workbook::new();
        build_sprachversionen(&mut workbook).unwrap();
        let config = LANG_CONFIG.get("Deutsch").unwrap();
        let suffix = "";
        let ws = workbook.add_worksheet();
        ws.set_name(config.fb_sheet).unwrap();
        let unlocked = Format::new().set_font_name("Arial").set_font_size(10.0).set_unlocked();
        for col in 0..USED_COLUMNS {
            ws.set_column_format(col, &unlocked).ok();
        }
        setup_sheet(ws).unwrap();
        t_wb_setup += t.elapsed();

        let t = Instant::now();
        let (values, body_config) = build_test_values(i);
        let options = SheetOptions::with_default_protection().with_hidden_columns_qv();
        write_report_with_options(ws, suffix, &values, &body_config, &options).unwrap();
        t_wb_write_report += t.elapsed();

        // Save to temp file (unprotected)
        let t = Instant::now();
        let temp_file = temp_dir.join(format!("report_{:05}_tmp.xlsx", i));
        workbook.save(&temp_file).unwrap();
        t_wb_save += t.elapsed();

        // Apply workbook protection (SHA-512 mit 100.000 Iterationen)
        let t = Instant::now();
        let final_file = temp_dir.join(format!("report_{:05}.xlsx", i));
        fb_generator::workbook_protection::protect_workbook_with_spin_count(
            temp_file.to_str().unwrap(),
            final_file.to_str().unwrap(),
            "benchmark_password",
            100_000,
        ).unwrap();
        t_wb_protect += t.elapsed();

        // Cleanup temp file
        let t = Instant::now();
        std::fs::remove_file(&temp_file).ok();
        t_wb_cleanup += t.elapsed();
    }
    let total_c = total_start.elapsed();

    println!("\n=== Variante C: Mit Workbook-Protection (SHA-512) ===");
    println!("  Gesamt:           {:>8.2?}  ({:.1} Reports/s)", total_c, ITERATIONS as f64 / total_c.as_secs_f64());
    print_phase("Workbook + Sheet Setup", t_wb_setup, total_c);
    print_phase("Values + write_report()", t_wb_write_report, total_c);
    print_phase("workbook.save() to disk", t_wb_save, total_c);
    print_phase("protect_workbook() SHA-512", t_wb_protect, total_c);
    print_phase("Temp-File Cleanup", t_wb_cleanup, total_c);

    println!("\n=== Vergleich Baseline vs C (Overhead durch Protection) ===");
    println!("  Ohne Protection:  {:>8.2?}  ({:.1} Reports/s)", total_baseline, ITERATIONS as f64 / total_baseline.as_secs_f64());
    println!("  Mit Protection:   {:>8.2?}  ({:.1} Reports/s)", total_c, ITERATIONS as f64 / total_c.as_secs_f64());
    println!("  Overhead:         {:.1}x langsamer", total_c.as_secs_f64() / total_baseline.as_secs_f64());
    println!("  protect_workbook() pro Report: {:.0}us ({:.2?})",
        t_wb_protect.as_micros() as f64 / ITERATIONS as f64,
        t_wb_protect / ITERATIONS as u32);

    // Cleanup
    std::fs::remove_dir_all(&temp_dir).ok();
    std::fs::create_dir_all(&temp_dir).unwrap();

    // Variante D: Fast Protection (1.000 statt 100.000 Iterationen)
    let mut t_fast_protect = std::time::Duration::ZERO;
    let mut t_fast_rest = std::time::Duration::ZERO;

    let total_start = Instant::now();
    for i in 0..ITERATIONS {
        let t = Instant::now();
        let mut workbook = Workbook::new();
        build_sprachversionen(&mut workbook).unwrap();
        let config = LANG_CONFIG.get("Deutsch").unwrap();
        let suffix = "";
        let ws = workbook.add_worksheet();
        ws.set_name(config.fb_sheet).unwrap();
        let unlocked = Format::new().set_font_name("Arial").set_font_size(10.0).set_unlocked();
        for col in 0..USED_COLUMNS {
            ws.set_column_format(col, &unlocked).ok();
        }
        setup_sheet(ws).unwrap();
        let (values, body_config) = build_test_values(i);
        let options = SheetOptions::with_default_protection().with_hidden_columns_qv();
        write_report_with_options(ws, suffix, &values, &body_config, &options).unwrap();
        let temp_file = temp_dir.join(format!("report_{:05}_tmp.xlsx", i));
        workbook.save(&temp_file).unwrap();
        t_fast_rest += t.elapsed();

        let t = Instant::now();
        let final_file = temp_dir.join(format!("report_{:05}.xlsx", i));
        fb_generator::workbook_protection::protect_workbook_with_spin_count(
            temp_file.to_str().unwrap(),
            final_file.to_str().unwrap(),
            "benchmark_password",
            1_000,
        ).unwrap();
        t_fast_protect += t.elapsed();

        std::fs::remove_file(&temp_file).ok();
    }
    let total_d = total_start.elapsed();

    println!("\n=== Variante D: Fast Protection (1.000 Iterationen) ===");
    println!("  Gesamt:           {:>8.2?}  ({:.1} Reports/s)", total_d, ITERATIONS as f64 / total_d.as_secs_f64());
    print_phase("Report-Generierung", t_fast_rest, total_d);
    print_phase("protect_workbook() FAST", t_fast_protect, total_d);

    // Cleanup
    std::fs::remove_dir_all(&temp_dir).ok();
    std::fs::create_dir_all(&temp_dir).unwrap();

    // Variante E: Nur ZIP-Overhead messen (spin_count=1, SHA-512 vernachlassigbar)
    let mut t_zip_only = std::time::Duration::ZERO;
    let mut t_hash_only = std::time::Duration::ZERO;
    let mut t_save_only = std::time::Duration::ZERO;

    for i in 0..ITERATIONS {
        let mut workbook = Workbook::new();
        build_sprachversionen(&mut workbook).unwrap();
        let config = LANG_CONFIG.get("Deutsch").unwrap();
        let suffix = "";
        let ws = workbook.add_worksheet();
        ws.set_name(config.fb_sheet).unwrap();
        let unlocked = Format::new().set_font_name("Arial").set_font_size(10.0).set_unlocked();
        for col in 0..USED_COLUMNS {
            ws.set_column_format(col, &unlocked).ok();
        }
        setup_sheet(ws).unwrap();
        let (values, body_config) = build_test_values(i);
        let options = SheetOptions::with_default_protection().with_hidden_columns_qv();
        write_report_with_options(ws, suffix, &values, &body_config, &options).unwrap();

        let t = Instant::now();
        let temp_file = temp_dir.join(format!("report_{:05}_tmp.xlsx", i));
        workbook.save(&temp_file).unwrap();
        t_save_only += t.elapsed();

        // ZIP-Roundtrip mit minimalem Hashing (spin_count=1)
        let t = Instant::now();
        let final_file = temp_dir.join(format!("report_{:05}.xlsx", i));
        fb_generator::workbook_protection::protect_workbook_with_spin_count(
            temp_file.to_str().unwrap(),
            final_file.to_str().unwrap(),
            "pw",
            1,
        ).unwrap();
        t_zip_only += t.elapsed();

        // Reines SHA-512 Hashing (ohne ZIP) zum Vergleich
        let t = Instant::now();
        fb_generator::workbook_protection::protect_workbook_with_spin_count(
            final_file.to_str().unwrap(),
            temp_file.to_str().unwrap(),
            "pw",
            100_000,
        ).unwrap();
        t_hash_only += t.elapsed();

        std::fs::remove_file(&temp_file).ok();
    }

    // Reines Hashing = (hash_only - zip_only), weil beide den ZIP-Roundtrip haben
    let zip_overhead = t_zip_only;
    let hash_time = t_hash_only - t_zip_only;

    println!("\n=== Variante E: Aufschluesselung protect_workbook() ===");
    println!("  (Gemessen ueber {} Reports)\n", ITERATIONS);
    println!("  workbook.save() to disk:        {:>8.2?}  ({:.0}us/report)",
        t_save_only, t_save_only.as_micros() as f64 / ITERATIONS as f64);
    println!("  ZIP Roundtrip (read+modify+write): {:>8.2?}  ({:.0}us/report)",
        zip_overhead, zip_overhead.as_micros() as f64 / ITERATIONS as f64);
    println!("  SHA-512 Hashing (100k Iter.):    {:>8.2?}  ({:.0}us/report)",
        hash_time, hash_time.as_micros() as f64 / ITERATIONS as f64);

    println!("\n=== Gesamtvergleich ===");
    println!("  Baseline (ohne Protection):  {:>8.2?}  ({:.1} Reports/s)", total_baseline, ITERATIONS as f64 / total_baseline.as_secs_f64());
    println!("  C) Standard Protection:      {:>8.2?}  ({:.1} Reports/s)", total_c, ITERATIONS as f64 / total_c.as_secs_f64());
    println!("  D) Fast Protection:          {:>8.2?}  ({:.1} Reports/s)", total_d, ITERATIONS as f64 / total_d.as_secs_f64());
    println!("  Speedup Fast vs Standard: {:.1}x", total_c.as_secs_f64() / total_d.as_secs_f64());

    std::fs::remove_dir_all(&temp_dir).ok();
}

fn print_phase(name: &str, duration: std::time::Duration, total: std::time::Duration) {
    let pct = duration.as_secs_f64() / total.as_secs_f64() * 100.0;
    let per_report = duration.as_micros() as f64 / ITERATIONS as f64;
    println!(
        "  {:<30} {:>8.2?}  ({:>5.1}%)  {:.0}us/report",
        name, duration, pct, per_report
    );
}
