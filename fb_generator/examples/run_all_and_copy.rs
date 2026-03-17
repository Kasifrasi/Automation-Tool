//! Runs all `test_*` examples and copies results to /home/ardit/VM/Share.
//!
//! Automatically discovers all `examples/test_*.rs` files — no manual list needed.
//!
//! Usage: cargo run -p fb_generator --example run_all_and_copy

use std::fs;
use std::path::Path;
use std::process::Command;

// Holt zur Kompilierzeit den absoluten Pfad zum fb_generator Ordner
const MANIFEST_DIR: &str = env!("CARGO_MANIFEST_DIR");
const TARGET_DIR: &str = "/home/ardit/VM/Share";

/// Discover all `test_*.rs` files in the examples directory.
fn discover_test_examples() -> Vec<String> {
    let mut examples = Vec::new();
    // Nutze den Manifest-Pfad als Basis
    let dir = Path::new(MANIFEST_DIR).join("examples");
    
    if let Ok(entries) = fs::read_dir(dir) {
        for entry in entries.flatten() {
            let name = entry.file_name();
            let name = name.to_string_lossy();
            if name.starts_with("test_") && name.ends_with(".rs") {
                examples.push(name.trim_end_matches(".rs").to_string());
            }
        }
    }
    examples.sort();
    examples
}

fn main() {
    let examples = discover_test_examples();
    println!("=== Running {} test examples, copying to {} ===\n", examples.len(), TARGET_DIR);

    if examples.is_empty() {
        println!("No test examples found. Exiting.");
        return;
    }

    // 1. Clean existing .xlsx files in target directory
    let target = Path::new(TARGET_DIR);
    if target.exists() {
        let mut deleted = 0;
        for entry in fs::read_dir(target).expect("Cannot read target directory") {
            let entry = entry.expect("Error reading directory entry");
            let path = entry.path();
            if path.extension().is_some_and(|ext| ext == "xlsx") {
                fs::remove_file(&path).expect("Cannot delete file");
                deleted += 1;
            }
        }
        if deleted > 0 {
            println!("  Cleaned {} existing .xlsx files from {}\n", deleted, TARGET_DIR);
        }
    } else {
        fs::create_dir_all(target).expect("Cannot create target directory");
        println!("  Created target directory {}\n", TARGET_DIR);
    }

    // 2. Run all discovered test examples
    let mut failed = Vec::new();
    for example in &examples {
        print!("  Running: {:<35} ", example);
        
        let output = Command::new("cargo")
            .current_dir(MANIFEST_DIR) // WICHTIG: Setzt das Arbeitsverzeichnis für Cargo
            .args(["run", "--example", example, "--release"])
            .output()
            .expect("Cannot run cargo");

        if output.status.success() {
            println!("OK");
        } else {
            println!("FAILED");
            let stderr = String::from_utf8_lossy(&output.stderr);
            eprintln!("    {}", stderr.lines().last().unwrap_or("Unknown error"));
            failed.push(example.as_str());
        }
    }

    if !failed.is_empty() {
        eprintln!("\n  FAILED: {}", failed.join(", "));
        std::process::exit(1);
    }

    // 3. Copy all .xlsx from examples/output/ to target
    let source = Path::new(MANIFEST_DIR).join("examples").join("output");
    let mut copied = 0;
    
    if source.exists() {
        for entry in fs::read_dir(source).expect("Cannot read output directory") {
            let entry = entry.expect("Error reading entry");
            let path = entry.path();
            if path.extension().is_some_and(|ext| ext == "xlsx") {
                let dest = target.join(entry.file_name());
                fs::copy(&path, &dest).expect("Cannot copy file");
                copied += 1;
            }
        }
    } else {
        println!("\n  No output directory found at {:?}, nothing to copy.", source);
    }

    println!("\n=== Done: {} examples run, {} files copied to {} ===", examples.len(), copied, TARGET_DIR);
}