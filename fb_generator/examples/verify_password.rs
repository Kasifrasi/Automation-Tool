//! Verify Workbook Protection Password Hash
//!
//! This tool helps verify that the password hashing is working correctly

use base64::{engine::general_purpose, Engine as _};
use byteorder::{WriteBytesExt, LE};
use sha2::{Digest, Sha512};

const SPIN_COUNT: u32 = 100_000;

fn hash_password_with_salt(password: &str, salt: &[u8], spin_count: u32) -> String {
    let pw_utf16: Vec<u8> = password
        .encode_utf16()
        .flat_map(|c| c.to_le_bytes())
        .collect();

    let mut hasher = Sha512::new();
    hasher.update(salt);
    hasher.update(&pw_utf16);
    let mut hash = hasher.finalize();

    for i in 0..spin_count {
        let mut iterator = [0u8; 4];
        let mut wtr = &mut iterator[..];
        wtr.write_u32::<LE>(i).unwrap();

        let mut next_hasher = Sha512::new();
        next_hasher.update(hash);
        next_hasher.update(iterator);
        hash = next_hasher.finalize();
    }

    general_purpose::STANDARD.encode(hash)
}

fn verify_against_known(label: &str, password: &str, salt_b64: &str, expected_hash_b64: &str, spin_count: u32) {
    println!("--- {} ---", label);
    println!("Passwort:         '{}'", password);
    println!("Salt (Base64):    {}", salt_b64);
    println!("Erwarteter Hash:  {}", expected_hash_b64);

    match general_purpose::STANDARD.decode(salt_b64) {
        Ok(salt) => {
            let computed = hash_password_with_salt(password, &salt, spin_count);
            println!("Berechneter Hash: {}", computed);
            if computed == expected_hash_b64 {
                println!("ERGEBNIS: OK – Algorithmus stimmt überein\n");
            } else {
                println!("ERGEBNIS: FEHLER – Hash stimmt NICHT überein!\n");
            }
        }
        Err(e) => println!("Salt-Dekodierung fehlgeschlagen: {}\n", e),
    }
}

fn main() {
    println!("Workbook Protection Password Hash Verifier\n");

    // -----------------------------------------------------------------------
    // Test 1: Bekannter OnlyOffice-Hash
    // Quelle: examples/output/xl/workbook.xml (von OnlyOffice 9.1.0.173 gesperrt)
    // Passwort: "thetool"
    // -----------------------------------------------------------------------
    verify_against_known(
        "Test 1: OnlyOffice-Hash (Passwort: thetool)",
        "thetool",
        "1sr8HKYlVqDYhNStPA4VYQ==",
        "3MRjsltOAYA9PBhJIOwEWwhaEJXlgkw3+LbtMdchZ39rXcpVIMEaNATZs0tiQSLiwoxCWOc4Zv0bCPsOeGUxHA==",
        SPIN_COUNT,
    );

    // -----------------------------------------------------------------------
    // Test 2: Rust-generierter Hash aus LOCKED_bench_000.xlsx
    // Passwort: "geheim123"
    // -----------------------------------------------------------------------
    verify_against_known(
        "Test 2: Rust-generierter Hash (Passwort: geheim123)",
        "geheim123",
        "Z2n5ecngkPJqNMk6aRpIsg==",
        "SnhhTw2en0FM+2pwjC76q7x01UNDiJeNx1We6AYAbvwpONwGGz5WXYBtzZ9SSDPegnBjRKOfZ78VnomISatVrA==",
        SPIN_COUNT,
    );

    // -----------------------------------------------------------------------
    // Test 3: Rust-generierte Datei aus Filesystem (falls vorhanden)
    // -----------------------------------------------------------------------
    let output_file = "examples/output/gesperrt/LOCKED_bench_000.xlsx";
    println!("--- Test 3: Datei aus Filesystem ({}) ---", output_file);

    let result = std::process::Command::new("unzip")
        .args(["-p", output_file, "xl/workbook.xml"])
        .output();

    if let Ok(output) = result {
        let xml = String::from_utf8_lossy(&output.stdout);

        let salt_b64 = extract_attr(&xml, "workbookSaltValue");
        let hash_b64 = extract_attr(&xml, "workbookHashValue");

        if let (Some(salt_b64), Some(hash_b64)) = (salt_b64, hash_b64) {
            println!("Salt:    {}", salt_b64);
            println!("Hash:    {}", hash_b64);

            if let Ok(salt) = general_purpose::STANDARD.decode(salt_b64) {
                let computed = hash_password_with_salt("geheim123", &salt, SPIN_COUNT);
                if computed == hash_b64 {
                    println!("ERGEBNIS: OK – 'geheim123' stimmt überein\n");
                } else {
                    println!("ERGEBNIS: FEHLER – 'geheim123' stimmt NICHT überein!\n");
                }
            }
        } else {
            println!("workbookSaltValue oder workbookHashValue nicht gefunden\n");
        }
    } else {
        println!("Datei nicht gefunden oder unzip nicht verfügbar\n");
    }
}

fn extract_attr<'a>(xml: &'a str, attr_name: &str) -> Option<&'a str> {
    let needle = format!("{}=\"", attr_name);
    let start = xml.find(needle.as_str())? + needle.len();
    let end = xml[start..].find('"')?;
    Some(&xml[start..start + end])
}
