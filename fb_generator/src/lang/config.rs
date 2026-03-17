//! Language configuration for all supported languages.
//!
//! Defines [`LanguageConfig`] with per-language sheet names and
//! error messages, plus the static map [`LANG_CONFIG`].

use std::collections::HashMap;
use std::sync::LazyLock;

/// Per-language configuration for sheet names and error messages.
#[derive(Debug, Clone)]
pub struct LanguageConfig {
    /// Base name for generated file names (e.g. `"D_Finanzbericht"`)
    pub base: &'static str,
    /// Language code as used in TEXT_MATRIX (e.g. `"deutsch"`)
    pub lang_val: &'static str,
    /// Localized sheet name (e.g. `"Finanzbericht"`)
    pub fb_sheet: &'static str,
    /// Decimal-place validation error message in the respective language
    pub decimal_err: &'static str,
}

/// Language configuration, keyed by German language name (e.g. `"Deutsch"`, `"Englisch"`).
pub static LANG_CONFIG: LazyLock<HashMap<&'static str, LanguageConfig>> = LazyLock::new(|| {
    HashMap::from([
        (
            "Deutsch",
            LanguageConfig {
                base: "D_Finanzbericht",
                lang_val: "deutsch",
                fb_sheet: "Finanzbericht",
                decimal_err: "Bitte geben Sie maximal 2 Nachkommastellen ein.",
            },
        ),
        (
            "Englisch",
            LanguageConfig {
                base: "E_Financial_report",
                lang_val: "english",
                fb_sheet: "Financial Report",
                decimal_err: "Please enter a maximum of 2 decimal places.",
            },
        ),
        (
            "Französisch",
            LanguageConfig {
                base: "F_Rapport_financier",
                lang_val: "français",
                fb_sheet: "Rapport financier",
                decimal_err: "Veuillez saisir un maximum de 2 décimales.",
            },
        ),
        (
            "Spanisch",
            LanguageConfig {
                base: "S_Informe_financiero",
                lang_val: "español",
                fb_sheet: "Informe financiero",
                decimal_err: "Por favor, introduzca un máximo de 2 decimales.",
            },
        ),
        (
            "Portugiesisch",
            LanguageConfig {
                base: "P_Relatorio_financeiro",
                lang_val: "português",
                fb_sheet: "Relatório financeiro",
                decimal_err: "Por favor, insira no máximo 2 casas decimais.",
            },
        ),
    ])
});

