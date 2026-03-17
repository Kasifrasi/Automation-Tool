//! Criterion-Benchmarks: Finanzbericht-Generierung
//!
//! Misst Throughput (Dateien/Sekunde) für Single- und Multi-threaded
//! Generierung mit vollständig ausgefüllten Reports (alle API-Features aktiv).
//!
//! Usage:
//!   cargo bench
//!   cargo bench -- --save-baseline main
//!   cargo bench -- --baseline main

use criterion::{criterion_group, criterion_main, BenchmarkId, Criterion, Throughput};
use fb_generator::{
    precompute_hash_with_spin_count, Currency, IncomeTable, Language, PanelEntry, PositionEntry,
    ReportBody, ReportConfig, ReportFooter, ReportHeader, ReportOptions, RowGrouping,
    SheetProtection, TableEntry,
};
use std::collections::HashMap;
use std::fs;
use std::sync::Arc;
use std::thread;
use std::time::{Duration, Instant};

// ============================================================================
// Realistische Test-Config: Alle API-Features aktiv
// ============================================================================

const LANGUAGES: [Language; 5] = [
    Language::Deutsch,
    Language::English,
    Language::Francais,
    Language::Espanol,
    Language::Portugues,
];
const BENCH_CURRENCIES: [Currency; 5] = [
    Currency::EUR, Currency::USD, Currency::GBP, Currency::CHF, Currency::BRL,
];
const BENCH_PASSWORD: &str = "bench_protection_pw";
const BENCH_SPIN_COUNT: u32 = 100_000;

/// Erstellt eine vollständig ausgefüllte ReportConfig.
///
/// Aktivierte Features pro Report:
/// - Alle Header-Felder (Projekt-Nr., Titel, Start/Ende, Berichtszeitraum)
/// - 5 Table-Einträge (Einnahmenabstimmung)
/// - 2×18 Panel-Einträge (Kassenbuch links + rechts)
/// - 50 Kostenpositionen (Kat. 1–5 je 10) + 3 Header-Input (Kat. 6–8)
/// - 3 Footer-Werte (Bank, Kasse, Sonstiges)
/// - Sheet Protection (sheet_password)
/// - Hidden columns (Q–V) + hidden language sheet
/// - Row Grouping mit collapsed sections
///
/// workbook_password wird NICHT gesetzt — stattdessen wird pro Batch
/// einmal precompute_hash aufgerufen und write_to_precomputed verwendet.
fn build_config(index: usize) -> ReportConfig {
    let lang_idx = index % 5;
    let base = 1000.0 + (index as f64 * 7.3);

    // Einnahmen-Tabelle (5 benannte Zeilen)
    let table = IncomeTable {
        kmw_mittel:   Some(TableEntry { approved_budget: Some(base * 10.0), income_report: Some(base * 5.0), income_total: Some(base * 5.0), reason: Some("KMW-Mittel".to_string()) }),
        eigenmittel:  Some(TableEntry { approved_budget: Some(base * 20.0), income_report: Some(base * 10.0), income_total: Some(base * 10.0), reason: Some("Eigenmittel".to_string()) }),
        drittmittel:  Some(TableEntry { approved_budget: Some(base * 30.0), income_report: Some(base * 15.0), income_total: Some(base * 15.0), reason: Some("Drittmittel".to_string()) }),
        saldovortrag: Some(TableEntry { approved_budget: Some(base * 40.0), income_report: Some(base * 20.0), income_total: Some(base * 20.0), reason: Some("Saldovortrag".to_string()) }),
        zinsertraege: Some(TableEntry { approved_budget: Some(base * 50.0), income_report: Some(base * 25.0), income_total: Some(base * 25.0), reason: Some("Zinserträge".to_string()) }),
    };

    // 18 Panel-Einträge links (Kassenbuch Euro-Konto)
    let left_panel: Vec<Option<PanelEntry>> = (0..18u8)
        .map(|i| {
            Some(PanelEntry {
                date: Some(format!("{:02}.{:02}.2024", (i % 28) + 1, (i % 12) + 1)),
                amount_euro: Some(base * (i + 1) as f64 * 0.5),
                amount_local: Some(base * (i + 1) as f64 * 0.6),
            })
        })
        .collect();

    // 18 Panel-Einträge rechts (Kassenbuch Lokal-Konto)
    let right_panel: Vec<Option<PanelEntry>> = (0..18u8)
        .map(|i| {
            Some(PanelEntry {
                date: Some(format!("{:02}.{:02}.2024", (i % 28) + 1, (i % 12) + 1)),
                amount_euro: Some(base * (i + 1) as f64 * 0.3),
                amount_local: Some(base * (i + 1) as f64 * 0.4),
            })
        })
        .collect();

    // Body: Kat. 1–5 je 10 Positionen, Kat. 6–8 Header-Input
    let normal_cats: [(u8, u16); 5] = [(1, 10), (2, 10), (3, 10), (4, 10), (5, 10)];
    let header_cats: [u8; 3] = [6, 7, 8];

    let mut positions: HashMap<u8, Vec<Option<PositionEntry>>> = HashMap::new();
    for &(cat, count) in &normal_cats {
        let rows = (1..=count)
            .map(|pos| {
                let cost = base * (cat as f64) * (pos as f64 / 3.0);
                Some(PositionEntry {
                    description: Some(format!("Kostenposition {}.{}", cat, pos)),
                    approved: Some(cost * 2.0),
                    income_report: Some(cost),
                    income_total: Some(cost),
                    remark: Some(format!("Bemerkung {}.{}", cat, pos)),
                })
            })
            .collect();
        positions.insert(cat, rows);
    }

    let mut header_inputs: HashMap<u8, Option<PositionEntry>> = HashMap::new();
    for &cat in &header_cats {
        header_inputs.insert(
            cat,
            Some(PositionEntry {
                description: None,
                approved: Some(base * cat as f64 * 2.0),
                income_report: Some(base * cat as f64),
                income_total: Some(base * cat as f64),
                remark: Some(format!("Pauschal Kat. {}", cat)),
            }),
        );
    }

    ReportConfig {
        header: ReportHeader {
            language: LANGUAGES[lang_idx],
            currency: BENCH_CURRENCIES[lang_idx],
            project_number: Some(format!("PROJ-{:05}", index)),
            project_title: Some(format!("Projekt {} ({})", index, LANGUAGES[lang_idx])),
            project_start: Some(format!("01.{:02}.2024", (index % 12) + 1)),
            project_end: Some("31.12.2026".to_string()),
            report_start: Some(format!("01.{:02}.2024", (index % 12) + 1)),
            report_end: Some("30.06.2024".to_string()),
            version: None,
        },
        body: ReportBody {
            table,
            left_panel,
            right_panel,
            positions,
            header_inputs,
        }, // table is now IncomeTable (named fields, no Vec)
        footer: ReportFooter {
            bank: Some(base * 8.0),
            kasse: Some(base * 1.5),
            sonstiges: Some(base * 0.3),
        },
        options: ReportOptions {
            sheet_protection: Some(SheetProtection::from_defaults().with_password("bench_pw")),
            hide_columns_qv: true,
            hide_language_sheet: true,
            row_grouping: Some(
                RowGrouping::new()
                    .add_collapsed_group(30, 40)
                    .add_collapsed_group(42, 52),
            ),
            // workbook_password NOT set — using precompute_hash + write_to_precomputed
            ..ReportOptions::default()
        },
    }
}

fn cleanup_dir(dir: &std::path::Path) {
    if let Ok(entries) = fs::read_dir(dir) {
        for entry in entries.flatten() {
            fs::remove_file(entry.path()).ok();
        }
    }
}

// ============================================================================
// Helper: Batch mit precomputed hash
// ============================================================================

fn bench_batch(
    dir: &std::path::Path,
    count: usize,
    num_threads: usize,
    iters: u64,
) -> Duration {
    let mut total = Duration::ZERO;

    for i in 0..iters as usize {
        let start = Instant::now();

        // 1× Hash pro Batch (mitgemessen, realistischer Workflow)
        let hash = precompute_hash_with_spin_count(BENCH_PASSWORD, BENCH_SPIN_COUNT);

        if num_threads <= 1 {
            // Single-threaded
            for j in 0..count {
                let config = build_config(i * count + j);
                let path = dir.join(format!("r_{}.xlsx", j));
                config.write_to_precomputed(&path, &hash).ok();
            }
        } else {
            // Multi-threaded
            let dir_arc = Arc::new(dir.to_path_buf());
            let hash_arc = Arc::new(hash);
            let chunk = count.div_ceil(num_threads);

            let handles: Vec<_> = (0..num_threads)
                .map(|tid| {
                    let d = Arc::clone(&dir_arc);
                    let h = Arc::clone(&hash_arc);
                    let start_j = tid * chunk;
                    let end_j = count.min(start_j + chunk);
                    let offset = i * count;
                    thread::spawn(move || {
                        for j in start_j..end_j {
                            let config = build_config(offset + j);
                            let path = d.join(format!("r_{}_{}.xlsx", tid, j));
                            config.write_to_precomputed(&path, &h).ok();
                        }
                    })
                })
                .collect();
            for h in handles {
                h.join().ok();
            }
        }

        total += start.elapsed();
        cleanup_dir(dir);
    }
    total
}

// ============================================================================
// Benchmarks
// ============================================================================

/// Single-threaded: 100 + 1.000 Dateien
fn bench_single_threaded(c: &mut Criterion) {
    let dir = tempfile::TempDir::new().unwrap();
    let mut group = c.benchmark_group("single_threaded");
    group
        .warm_up_time(Duration::from_secs(1))
        .measurement_time(Duration::from_secs(5))
        .sample_size(10);

    for &count in &[100usize, 1_000] {
        group.throughput(Throughput::Elements(count as u64));
        group.bench_with_input(
            BenchmarkId::from_parameter(count),
            &count,
            |b, &count| {
                b.iter_custom(|iters| bench_batch(dir.path(), count, 1, iters));
            },
        );
    }
    group.finish();
}

/// Multi-threaded: 8 Threads, 100 + 1.000 Dateien
fn bench_multi_threaded_8(c: &mut Criterion) {
    let dir = tempfile::TempDir::new().unwrap();
    let mut group = c.benchmark_group("multi_threaded_8");
    group
        .warm_up_time(Duration::from_secs(1))
        .measurement_time(Duration::from_secs(5))
        .sample_size(10);

    for &count in &[100usize, 1_000] {
        group.throughput(Throughput::Elements(count as u64));
        group.bench_with_input(
            BenchmarkId::from_parameter(count),
            &count,
            |b, &count| {
                b.iter_custom(|iters| bench_batch(dir.path(), count, 8, iters));
            },
        );
    }
    group.finish();
}

/// Multi-threaded: 16 Threads, 100 + 1.000 Dateien
fn bench_multi_threaded_16(c: &mut Criterion) {
    let dir = tempfile::TempDir::new().unwrap();
    let mut group = c.benchmark_group("multi_threaded_16");
    group
        .warm_up_time(Duration::from_secs(1))
        .measurement_time(Duration::from_secs(5))
        .sample_size(10);

    for &count in &[100usize, 1_000] {
        group.throughput(Throughput::Elements(count as u64));
        group.bench_with_input(
            BenchmarkId::from_parameter(count),
            &count,
            |b, &count| {
                b.iter_custom(|iters| bench_batch(dir.path(), count, 16, iters));
            },
        );
    }
    group.finish();
}

criterion_group!(
    benches,
    bench_single_threaded,
    bench_multi_threaded_8,
    bench_multi_threaded_16,
);
criterion_main!(benches);
