//! Static language data: text matrix and currencies.
//!
//! - [`TEXT_MATRIX`]: 66 UI terms x 5 languages (German, English, French, Spanish, Portuguese)
//! - [`Currency`]: ISO 4217 enum with all 153 supported currency codes

use std::fmt;
use std::str::FromStr;
use std::sync::LazyLock;

pub static TEXT_MATRIX: LazyLock<Vec<Vec<&'static str>>> = LazyLock::new(|| {
    vec![
        vec![
            "deutsch",
            "F I N A N Z B E R I C H T",
            "ÜBERSICHT ÜBER EINNAHMEN UND AUSGABEN",
            "Projektnummer",
            "Projektlaufzeit",
            "Aktueller Berichtszeitraum",
            "von (TT MM JJJJ):",
            "bis (TT MM JJJJ):",
            "EINNAHMEN",
            "(in lokaler Währung)",
            "Bewilligtes Budget",
            "Einnahmen im Berichtszeitraum",
            "Einnahmen der gesamten Projektlaufzeit",
            "% erhaltene Einnahmen",
            "Begründung der Budgetüber- oder unterschreitung",
            "Saldovortrag (aus vorherigem Berichtszeitraum)",
            "Lokale Eigenleistung",
            "Beiträge von dritter Seite",
            "KMW Mittel*",
            "Zinserträge",
            "GESAMT",
            "Datum",
            "Rate",
            "AUSGABEN",
            "Ausgaben im Berichtszeitraum",
            "Projekttitel",
            "Sprache",
            "Lokalwährung",
            "Bauausgaben",
            "Zwischensumme Bauausgaben",
            "Einmalige Ausgaben (Investitionen)",
            "Zwischensumme Einmalige Ausgaben (Investitionen)",
            "Personalkosten",
            "Zwischensumme Personalkosten",
            "Laufende Ausgaben für Projektaktivitäten",
            "Zwischensumme Laufende Ausgaben für Projektaktivitäten",
            "Projektverwaltung",
            "Zwischensumme Projektverwaltung",
            "Evaluierung",
            "Audit",
            "Reserve - nur nach Genehmigung",
            "Summe",
            "ABSCHLUSS",
            "Saldo für den Berichtszeitraum",
            "Saldo (Einnahmen minus Ausgaben)",
            "Saldenabstimmung:",
            "Bank",
            "Kasse",
            "Sonstiges (noch nicht eingelöste Schecks, Vorschüsse, Darlehen, etc.)",
            "Ich/Wir bestätige/n die Richtigkeit des vorliegenden Berichts unter Einhaltung des von beiden Partnern unterzeichneten Projektvertrages.",
            "Ort, Datum                  Unterschrift des Projektverantwortlichen",
            "Unterschrift einer weiteren autorisierten Person",
            "Funktion in der Organisation: ",
            "Ich/Wir bestätige/n, dass die lokale Eigenleistung monetär ist. ",
            "Ausgaben der gesamten Projektlaufzeit",
            "% verwendetes Budget",
            "Wechselkurs",
            "WK",
            "Ausfüllbare Felder sind in den Farben links",
            "Tipp:",
            "Bitte nutzen Sie zum ausfüllen des Finanzberichts unsere Ausfüllhilfe auf unserer Webseite:",
            "https://www.kindermissionswerk.org/fuer-partner/service",
            "EUR",
            "Zwischensumme Evaluierung",
            "Zwischensumme Audit",
            "Zwischensumme Reserve",
        ],
        vec![
            "english",
            "F I N A N C I A L  R E P O R T ",
            "OVERVIEW - INCOME AND EXPENDITURE",
            "Project Number",
            "Project Duration",
            "Actual Reporting Period",
            "From (DD.MM.YYYY)",
            "To (DD.MM.YYYY)",
            "INCOME",
            "(In Local Currency)",
            "Approved Budget",
            "Income During Reporting Period",
            "Total project Income (Till date) ",
            "% Income received",
            "Justification of budget surplus or deficit",
            "Balance -  carried forward from previous reporting period",
            "Local Contribution",
            "Third Party Contribution",
            "KMW Funds*",
            "Interest gains",
            "Total",
            "Date",
            "Installment",
            "EXPENSES",
            "Expenses in Reporting Period",
            "Project Title",
            "Language",
            "Local Currrency",
            "Building Expenses",
            "Subtotal Building Expenses",
            "Non-Recurring Expenses (Investments)",
            "Subtotal Non Recurring Expenses (Investments)",
            "Personnel Expenses",
            "Subtotal Personnel Expenses",
            "Ongoing Expenses for Project Activities",
            "Subtotal Ongoing Expenses for Project Activities",
            "Project Administration",
            "Subtotal Project Administration",
            "Evaluation",
            "Audit",
            "Contingency reserve - after authorisation only",
            "Total",
            "Summary",
            "Balance for the Reporting Period",
            "Balance (Income Minus Expenses)",
            "Balance Reconciliation",
            "Bank",
            "Cash",
            "Others (Uncashed Cheque, advances, loans etc.)",
            "I/We confirm the accuracy of this report in compliance with the project contract signed by both partners.",
            "Place, date Signature of the person responsible for the project",
            "Signature of another authorized person",
            "Function within the organization:",
            "I/We confirm that the local contribution is monetary. ",
            "Total Project Expenditure",
            "% Used Budget",
            "Exchange Rate",
            "ER",
            "Fillable fields are in the colours on the left",
            "Tip:",
            "Please use our completion guide on our website to fill out the financial report.",
            "https://www.kindermissionswerk.org/en/partners/service",
            "EUR",
            "Subtotal Evaluation",
            "Subtotal Audit",
            "Subtotal Contingency Reserve",
        ],
        vec![
            "français",
            "R A P P O R T   F I N A N C I E R",
            "APERÇU DES RECETTES ET DÉPENSES",
            "Numéro du projet",
            "Durée totale du projet",
            "Période actuelle couverte par le rapport",
            "de (JJ/MM/AAAA) :",
            "de (JJ/MM/AAAA) :",
            "RECETTES",
            "(en monnaie locale)",
            "Budget apprové",
            "Recettes pendant la période du rapport",
            "Cumul recettes sur toute la période",
            "% des recettes perçues",
            "Explication du dépassement de budget",
            "Solde (de la période précédente)",
            "Apport local",
            "Contributions de tiers",
            "Fonds KMW*",
            "Intérêts bancaires",
            "TOTAL",
            "Date",
            "versement",
            "DÉPENSES",
            "Dépenses pendant la période du rapport",
            "Titre du projet",
            "Langue",
            "Monnaie locale",
            "Dépenses de construction",
            "Sous-total dépenses de construction",
            "Dépenses uniques (investissements)",
            "Sous-total dépenses uniques (investissements)",
            "Personnel",
            "Sous-total personnel",
            "Dépenses courantes pour activités de projet",
            "Sous-total dépenses courantes pour activités de projet",
            "Gestion de projet",
            "Sous-total gestion de projet",
            "Evaluation",
            "Audit",
            "Réserve - uniquement après accord",
            "TOTAL",
            "CLÔTURE DES COMPTES",
            "Solde à la fin de la période du rapport",
            "Solde (recettes moins dépenses)",
            "Rapprochement des soldes :",
            "Banque",
            "Caisse",
            "Divers (chèques non encaissés, avances, prêts, etc.)",
            "Je confirme / Nous certifions que le présent rapport est exact et conforme au contrat de projet signé par les deux partenaires.",
            "Lieu, date                  Signature du responsable du projet",
            "Signature d’une autre personne autorisée",
            "Fonction au sein de l’organisation : ",
            "Je confirme / Nous confirmons que l'apport local est monétaire.",
            "Cumul dépenses sur toute la période",
            "% budget consommé",
            "Taux de change",
            "TC",
            "Les champs remplissables sont dans les couleurs à gauche",
            "Astuce :",
            "Veuillez utiliser notre aide à la saisie disponible sur notre site web pour remplir le rapport financier:",
            "https://www.kindermissionswerk.org/fr/pour-les-partenaires/services",
            "EUR",
            "Sous-total Évaluation",
            "Sous-total Audit",
            "Sous-total Réserve",
        ],
        vec![
            "español",
            "I N F O R M E  F I N A N C I E R O",
            "Resumen de los ingresos y egresos",
            "Número del proyecto",
            "Duración del proyecto",
            "Período de referencia",
            "de (DD MM AAAA):",
            "hasta (DD MM AAAA)",
            "Ingresos",
            "(en moneda local)",
            "Presupuesto aprobado",
            "Ingresos del período de referencia",
            "Ingresos acumulados (total desde el início)",
            "% Ingresos recibidos",
            "Justificación del superávit / déficit presupuestario",
            "Saldo (procendente del período de referencia anterior)",
            "Aporte propio",
            "Aportes de terceros",
            "Fondos KMW *",
            "Intereses acreedores percibidos",
            "TOTAL",
            "Fecha",
            "Cuota",
            "EGRESOS",
            "Egresos del período de referencia",
            "Título del proyecto",
            "Idioma",
            "Moneda local",
            "Gastos de construcción",
            "Subtotal gastos de construcción",
            "Gastos extraordinarios (inversiones)",
            "Subtotal gastos extraorinarios (inversiones) ",
            "Gastos de personal",
            "Subtotal gastos de personal",
            "Gastos corrientes para actividades del proyecto",
            "Subtotal gastos corrientes para actividades",
            "Gastos administrativos del proyecto / Gestión del proyecto ",
            "Subtotal gastos administrativos del proyecto",
            "Evaluación",
            "Auditoría externa",
            "Reserva - solo previa autorización",
            "Total",
            "CIERRE DE CUENTAS",
            "Saldo en el período de referencia",
            "Saldo (ingresos menos egresos)",
            "Conciliación de saldos ",
            "Banco",
            "Caja",
            "Otros (cheques aún no cobrados, anticipos, créditos, etc.)",
            "Por la presente certifico/certificamos la corrección del presente informe, el cual cumple las disposiciones del contrato de proyecto ",
            "Lugar, fecha, firma del resonsable de proyecto",
            "Firma de otra persona autorizada",
            "Cargo  o Función dentro de la organización: ",
            "Confirmo/confirmamos que el aporte propio es monetario ",
            "Egresos acumulados desde el início del proyecto",
            "Ejecución presupuestaria (en %) ",
            "Tipo de cambio",
            "TC",
            "Los campos rellenables aparecen en los colores de la izquierda",
            "Consejo:",
            "Por favor, utilice nuestra guía para completar el informe financiero, disponible en nuestro sitio web:",
            "https://www.kindermissionswerk.org/es/para-contrapartes-de-proyectos/servicios",
            "EUR",
            "Subtotal Evaluación",
            "Subtotal Auditoría Externa",
            "Subtotal Reserva",
        ],
        vec![
            "português",
            "RELATÓRIO FINANCEIRO",
            "Resumo das receitas e despesas",
            "Número do projeto ",
            "Duração do projeto",
            "Período de referência",
            "de (DD MM AAAA):",
            "até (DD MM AAAA):",
            "RECEITAS",
            "(em moeda local)",
            "Plano de custos e financiamento aprovado",
            "Receitas do período de referência",
            "Receitas cumuladas (total desde o início)",
            "% Reiceitas recebidas",
            "Justificação do superávit / déficit orçamentário",
            "Saldo transportado (do período de referência anterior)",
            "Contribuição própria ",
            "Contribuções de terceiros",
            "Subsídio KMW *",
            "Juros credores auferidos",
            "TOTAL",
            "Data",
            "transferência parcial",
            "DESPESAS",
            "Despesas do período de referência",
            "Título do projeto",
            "Idioma",
            "Moeda local",
            "Despesas de construção (civil)",
            "Subtotal despesas de construção (civil) ",
            "Despesas únicas (investimentos)",
            "Subtotal Despesas únicas (investimentos)",
            "Despesas com pessoal / recursos humanos",
            "Subtotal Despesas com pessoal / recursos humanos",
            "Custos de atividades de projeto",
            "Subtotal Custos de atividades de projeto",
            "Administração do projeto",
            "Subtotal Administração do projeto",
            "Avaliação",
            "Auditoria externa",
            "Reserva - somente após autorização",
            "Total",
            "BALANÇO",
            "Saldo do período de referência",
            "Saldo (receitas menos despesas",
            "Conciliação de saldo ",
            "Banco",
            "Caixa",
            "Outros (cheques não descontados, adiantamentos, empréstimos, etc.).",
            "Declaro/declaramos a veracidade deste relatório em conformidade com o contrato firmado entre as partes.",
            "Local, data e assinatura do responsável do projeto",
            "Assinatura de uma outra pessoa devidamente autorizada",
            "Função / cargo na organização:",
            "Confirmo/confirmamos que a contribuição prória é monetária ",
            "Despesas acumuladas desde o início de projeto ",
            "Execução orçamentária (em %)",
            "Taxa de câmbio",
            "TC",
            "Os campos preenchíveis estão nas cores à esquerda",
            "Dica:",
            "Por favor, utilize nosso guia de preenchimento disponível em nosso site para preencher o relatório financeiro:",
            "https://www.kindermissionswerk.org/pt/para-os-parceiros/servico",
            "EUR",
            "Subtotal Avaliação",
            "Subtotal Auditoria Externa",
            "Subtotal Reserva",
        ],
    ]
});

// ============================================================================
// Currency - ISO 4217 Enum
// ============================================================================

macro_rules! define_currency {
    ($($variant:ident),* $(,)?) => {
        /// ISO 4217 currency code.
        ///
        /// All 153 standard codes are available as enum variants (`Currency::EUR`,
        /// `Currency::USD`, etc.). Parse from a string with `FromStr` (case-insensitive).
        ///
        /// # Example
        /// ```ignore
        /// use fb_generator::Currency;
        ///
        /// let eur = Currency::EUR;
        /// let usd: Currency = "USD".parse().unwrap();
        /// let invalid: Result<Currency, _> = "INVALID".parse();
        /// assert!(invalid.is_err());
        /// ```
        ///
        /// ## Serde (requires `serde` feature)
        ///
        /// Serializes to the uppercase code string. Deserializes case-insensitively.
        ///
        /// ```json
        /// "EUR"
        /// "USD"
        /// ```
        #[allow(clippy::upper_case_acronyms)]
        #[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
        #[cfg_attr(feature = "serde", derive(serde::Serialize, serde::Deserialize))]
        #[cfg_attr(feature = "serde", serde(try_from = "String", into = "String"))]
        pub enum Currency {
            $($variant),*
        }

        impl Currency {
            /// Returns the ISO 4217 code as a static string slice.
            pub fn as_str(&self) -> &'static str {
                match self {
                    $(Currency::$variant => stringify!($variant),)*
                }
            }

            /// Returns a slice of all 153 supported currency codes.
            pub fn all() -> &'static [Currency] {
                const ALL: &[Currency] = &[$(Currency::$variant),*];
                ALL
            }
        }

        impl FromStr for Currency {
            type Err = String;

            /// Parses a currency code (case-insensitive).
            fn from_str(s: &str) -> Result<Self, Self::Err> {
                match s.to_uppercase().as_str() {
                    $(stringify!($variant) => Ok(Currency::$variant),)*
                    _ => Err(format!("Unknown currency: {}", s)),
                }
            }
        }
    };
}

define_currency!(
    AED, AFN, ALL, AMD, ANG, AOA, ARS, AUD, AWG, AZN,
    BAM, BBD, BDT, BGN, BHD, BIF, BMD, BND, BOB, BRL, BSD, BTN, BWP, BYN, BZD,
    CAD, CDF, CHF, CLP, CNY, COP, CRC, CUP, CVE, CZK,
    DJF, DKK, DOP, DZD,
    EGP, ERN, ETB, EUR,
    FJD, FKP,
    GBP, GEL, GHS, GIP, GMD, GNF, GTQ, GYD,
    HKD, HNL, HTG, HUF,
    IDR, ILS, INR, IQD, IRR, ISK,
    JMD, JOD, JPY,
    KES, KGS, KHR, KMF, KRW, KWD, KYD, KZT,
    LAK, LBP, LKR, LRD, LSL, LYD,
    MAD, MDL, MGA, MKD, MMK, MNT, MOP, MRU, MUR, MVR, MWK, MXN, MYR, MZN,
    NAD, NGN, NIO, NOK, NPR, NZD,
    OMR,
    PAB, PEN, PGK, PHP, PKR, PLN, PYG,
    QAR,
    RON, RSD, RUB, RWF,
    SAR, SBD, SCR, SDG, SEK, SGD, SHP, SLE, SOS, SRD, SSP, STN, SYP, SZL,
    THB, TJS, TMT, TND, TOP, TRY, TTD, TWD, TZS,
    UAH, UGX, USD, UYU, UZS,
    VES, VND, VUV,
    WST,
    XAF, XCD, XOF, XPF,
    YER,
    ZAR, ZMW, ZWL,
);

impl fmt::Display for Currency {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.write_str(self.as_str())
    }
}

#[allow(clippy::derivable_impls)] // macro-generated enum: #[default] on EUR not possible in macro
impl Default for Currency {
    fn default() -> Self {
        Currency::EUR
    }
}

#[cfg(feature = "serde")]
impl TryFrom<String> for Currency {
    type Error = String;

    fn try_from(s: String) -> Result<Self, Self::Error> {
        s.parse()
    }
}

#[cfg(feature = "serde")]
impl From<Currency> for String {
    fn from(c: Currency) -> Self {
        c.as_str().to_string()
    }
}

