//! README-Updater: Liest criterion JSON-Ergebnisse und aktualisiert die Performance-Tabelle.
//!
//! Wird automatisch nach `cargo bench --bench report_generation` ausgeführt:
//!   cargo bench
//!
//! Oder einzeln:
//!   cargo bench --bench update_readme

use serde_json::Value;
use std::collections::BTreeMap;
use std::fs;
use std::path::{Path, PathBuf};
use sysinfo::System;

fn main() {
    let project_root = Path::new(env!("CARGO_MANIFEST_DIR"));
    let criterion_dir = project_root.join("target").join("criterion");
    let readme_path = project_root.join("README.md");

    if !criterion_dir.exists() {
        eprintln!("No criterion results found at {:?}", criterion_dir);
        eprintln!("Run `cargo bench --bench report_generation` first.");
        return;
    }

    // Find all benchmark.json files in "new/" directories
    let bench_dirs = find_bench_dirs(&criterion_dir);
    if bench_dirs.is_empty() {
        eprintln!("No benchmark results found under {:?}", criterion_dir);
        return;
    }

    // Parse each benchmark → (files, threads) → stats
    let mut results: BTreeMap<(usize, usize), BenchResult> = BTreeMap::new();
    for dir in &bench_dirs {
        match parse_bench_dir(dir) {
            Ok(Some((key, result))) => {
                results.insert(key, result);
            }
            Ok(None) => {} // Unknown group, skip
            Err(e) => eprintln!("Warning: {:?}: {}", dir, e),
        }
    }

    if results.is_empty() {
        eprintln!("No valid benchmark data found.");
        return;
    }

    let hw = hardware_info();
    let table = format_table(&results);
    let date = current_date();

    let section = format!(
        "{hw}\n\n\
         | Files | Threads | Mean | Std Dev | CV | Throughput |\n\
         |------:|--------:|-----:|--------:|---:|-----------:|\n\
         {table}\n\n\
         *Last updated: {date}*"
    );

    if let Err(e) = update_readme(&readme_path, &section) {
        eprintln!("Failed to update README: {}", e);
        std::process::exit(1);
    }

    println!("README.md updated with {} benchmark configurations.", results.len());
}

// ============================================================================
// Criterion JSON discovery + parsing
// ============================================================================

struct BenchResult {
    mean_ns: f64,
    std_dev_ns: f64,
    throughput_elements: u64,
}

/// Find all "new/" directories containing benchmark.json
fn find_bench_dirs(dir: &Path) -> Vec<PathBuf> {
    let mut result = Vec::new();
    walk_for_new_dirs(dir, &mut result);
    result
}

fn walk_for_new_dirs(dir: &Path, result: &mut Vec<PathBuf>) {
    let entries = match fs::read_dir(dir) {
        Ok(e) => e,
        Err(_) => return,
    };
    for entry in entries.flatten() {
        let path = entry.path();
        if path.is_dir() {
            if path.file_name().map(|f| f == "new").unwrap_or(false)
                && path.join("benchmark.json").exists()
            {
                result.push(path);
            } else {
                walk_for_new_dirs(&path, result);
            }
        }
    }
}

type BenchKey = (usize, usize);

/// Parse a single benchmark "new/" directory → (files, threads) key + stats
fn parse_bench_dir(dir: &Path) -> Result<Option<(BenchKey, BenchResult)>, String> {
    let bench: Value = read_json(&dir.join("benchmark.json"))?;
    let estimates: Value = read_json(&dir.join("estimates.json"))?;

    let group_id = bench["group_id"].as_str().unwrap_or("");
    let value_str = bench["value_str"].as_str().unwrap_or("");
    let throughput_elements = bench["throughput"]["Elements"]
        .as_u64()
        .unwrap_or(0);

    let mean_ns = estimates["mean"]["point_estimate"].as_f64().unwrap_or(0.0);
    let std_dev_ns = estimates["std_dev"]["point_estimate"].as_f64().unwrap_or(0.0);

    if mean_ns <= 0.0 || throughput_elements == 0 {
        return Ok(None);
    }

    // Map (group_id, value_str) → (files, threads)
    let key = match group_id {
        "single_threaded" => {
            let files: usize = value_str.parse().map_err(|e| format!("{}", e))?;
            (files, 1)
        }
        "multi_threaded_8" => {
            let files: usize = value_str.parse().map_err(|e| format!("{}", e))?;
            (files, 8)
        }
        "multi_threaded_16" => {
            let files: usize = value_str.parse().map_err(|e| format!("{}", e))?;
            (files, 16)
        }
        _ => return Ok(None),
    };

    Ok(Some((
        key,
        BenchResult {
            mean_ns,
            std_dev_ns,
            throughput_elements,
        },
    )))
}

fn read_json(path: &Path) -> Result<Value, String> {
    let content = fs::read_to_string(path).map_err(|e| format!("Read {:?}: {}", path, e))?;
    serde_json::from_str(&content).map_err(|e| format!("Parse {:?}: {}", path, e))
}

// ============================================================================
// Formatting
// ============================================================================

fn format_duration(ms: f64) -> String {
    if ms >= 1000.0 {
        format!("{:.2}s", ms / 1000.0)
    } else {
        format!("{}ms", ms.round() as u64)
    }
}

fn format_table(data: &BTreeMap<(usize, usize), BenchResult>) -> String {
    let mut rows = Vec::new();

    for (&(files, threads), result) in data {
        let duration_ms = result.mean_ns / 1e6;
        let std_dev_ms = result.std_dev_ns / 1e6;
        let cv = result.std_dev_ns / result.mean_ns * 100.0;
        let throughput = result.throughput_elements as f64 * 1e9 / result.mean_ns;

        let files_str = if files >= 1000 {
            format!("{},{:03}", files / 1000, files % 1000)
        } else {
            format!("{}", files)
        };

        rows.push(format!(
            "| {:>5} | {:>7} | {:>8} | {:>9} | {:>4.1}% | {:>10}/sec |",
            files_str,
            threads,
            format_duration(duration_ms),
            format!("±{}", format_duration(std_dev_ms)),
            cv,
            throughput.round() as u64,
        ));
    }

    rows.join("\n")
}

fn hardware_info() -> String {
    let sys = System::new_all();
    let cpu_name = sys
        .cpus()
        .first()
        .map(|c| c.brand().to_string())
        .unwrap_or_else(|| "Unknown CPU".to_string());
    let logical_cores = sys.cpus().len();
    let physical_cores = System::physical_core_count().unwrap_or(logical_cores);
    let ram_gb = sys.total_memory() as f64 / 1_073_741_824.0;
    let os_version = System::long_os_version().unwrap_or_else(|| "Unknown OS".to_string());

    format!(
        "**Environment:** {} \u{00b7} {} cores ({} logical) \u{00b7} {:.0} GB RAM \u{00b7} {}",
        cpu_name.trim(),
        physical_cores,
        logical_cores,
        ram_gb,
        os_version.trim(),
    )
}

fn current_date() -> String {
    std::process::Command::new("date")
        .arg("+%Y-%m-%d")
        .output()
        .ok()
        .and_then(|o| String::from_utf8(o.stdout).ok())
        .map(|s| s.trim().to_string())
        .unwrap_or_else(|| "unknown".to_string())
}

// ============================================================================
// README update
// ============================================================================

fn update_readme(path: &Path, section: &str) -> Result<(), String> {
    let content = fs::read_to_string(path).map_err(|e| format!("Read error: {}", e))?;

    let start_marker = "<!-- PERF_START -->";
    let end_marker = "<!-- PERF_END -->";

    let start = content
        .find(start_marker)
        .ok_or("Missing <!-- PERF_START --> marker in README.md")?;
    let end = content
        .find(end_marker)
        .ok_or("Missing <!-- PERF_END --> marker in README.md")?;

    let before = &content[..start + start_marker.len()];
    let after = &content[end..];

    let new_content = format!("{}\n{}\n{}", before, section, after);
    fs::write(path, new_content).map_err(|e| format!("Write error: {}", e))?;

    Ok(())
}
