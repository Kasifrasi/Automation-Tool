#![windows_subsystem = "windows"]

slint::include_modules!();

use fb_generator::{
    Language, PositionEntry, ReportBody, ReportConfig, ReportHeader, ReportOptions, SheetProtection,
};
use std::path::PathBuf;

// ==========================================
// Defaults (Single Source of Truth)
// ==========================================

fn apply_fb_defaults(ui: &MainWindow) {
    let fb = ui.global::<FBState>();

    fb.set_langs(Languages {
        de: true,
        en: false,
        fr: false,
        es: false,
        pt: false,
    });

    fb.set_version("".into());
    fb.set_folder("".into());

    fb.set_categories(Categories {
        cat1: 20,
        cat2: 20,
        cat3: 30,
        cat4: 30,
        cat5: 20,
        cat6: 0,
        cat7: 0,
        cat8: 0,
    });

    fb.set_protect_sheet(true);
    fb.set_protect_workbook(true);
    fb.set_sheet_password("".into());
    fb.set_workbook_password("".into());
    fb.set_hide_columns(true);
    fb.set_hide_lang_sheet(true);

    fb.set_sheet_permissions(SheetPermissions {
        select_locked: true,
        select_unlocked: true,
        format_cells: true,
        format_columns: true,
        format_rows: true,
        insert_columns: false,
        insert_rows: false,
        insert_hyperlinks: true,
        delete_columns: true,
        delete_rows: true,
        sort: true,
        autofilter: true,
        pivot_tables: true,
        edit_objects: false,
        edit_scenarios: true,
        contents: false,
    });

    fb.set_status_type("idle".into());
    fb.set_status_message("".into());
}

fn apply_b2f_defaults(ui: &MainWindow) {
    let b2f = ui.global::<BudgetState>();
    b2f.set_src_folder("".into());
    b2f.set_out_folder("".into());
    b2f.set_filename("v2_scan".into());
    b2f.set_include_xlsm(true);
    b2f.set_export_txt(false);
    b2f.set_export_csv(false);
    b2f.set_status_type("idle".into());
    b2f.set_status_message("".into());
}

fn apply_folder_defaults(ui: &MainWindow) {
    let fs = ui.global::<FolderState>();
    fs.set_target_folder("".into());
    fs.set_project_name("".into());
    fs.set_status_type("idle".into());
    fs.set_status_message("".into());
}

// ==========================================
// Main
// ==========================================

fn main() -> Result<(), slint::PlatformError> {
    let ui = MainWindow::new()?;

    // Defaults setzen
    apply_fb_defaults(&ui);
    apply_b2f_defaults(&ui);
    apply_folder_defaults(&ui);

    // Dark Mode: System-Einstellung erkennen
    let system_dark = matches!(dark_light::detect(), Ok(dark_light::Mode::Dark));
    ui.set_dark_mode(system_dark);
    if system_dark {
        ui.global::<Palette>()
            .set_color_scheme(slint::language::ColorScheme::Dark);
    }

    // Dark Mode Toggle
    ui.on_toggle_dark_mode({
        let ui_handle = ui.as_weak();
        move |dark| {
            if let Some(ui) = ui_handle.upgrade() {
                let scheme = if dark {
                    slint::language::ColorScheme::Dark
                } else {
                    slint::language::ColorScheme::Light
                };
                ui.global::<Palette>().set_color_scheme(scheme);
            }
        }
    });

    // ==========================================
    // FB-Generator Callbacks
    // ==========================================

    ui.global::<FBState>().on_select_folder({
        let ui_handle = ui.as_weak();
        move || {
            if let Some(ui) = ui_handle.upgrade() {
                if let Some(path) = rfd::FileDialog::new().pick_folder() {
                    let fb = ui.global::<FBState>();
                    fb.set_folder(path.to_string_lossy().to_string().into());
                    fb.set_status_type("idle".into());
                    fb.set_status_message("".into());
                }
            }
        }
    });

    ui.global::<FBState>().on_reset({
        let ui_handle = ui.as_weak();
        move || {
            if let Some(ui) = ui_handle.upgrade() {
                apply_fb_defaults(&ui);
            }
        }
    });

    ui.global::<FBState>().on_dismiss_status({
        let ui_handle = ui.as_weak();
        move || {
            if let Some(ui) = ui_handle.upgrade() {
                let fb = ui.global::<FBState>();
                fb.set_status_type("idle".into());
                fb.set_status_message("".into());
            }
        }
    });

    ui.global::<FBState>().on_generate_report({
        let ui_handle = ui.as_weak();
        move || {
            if let Some(ui) = ui_handle.upgrade() {
                let fb = ui.global::<FBState>();

                let folder = fb.get_folder().to_string();
                if folder.is_empty() {
                    fb.set_status_type("error".into());
                    fb.set_status_message("Bitte Zielordner wählen.".into());
                    return;
                }

                let version = fb.get_version().to_string();
                if version.is_empty() {
                    fb.set_status_type("error".into());
                    fb.set_status_message("Bitte Version angeben.".into());
                    return;
                }

                let langs = fb.get_langs();
                let mut lang_list = Vec::new();
                if langs.de {
                    lang_list.push(Language::Deutsch);
                }
                if langs.en {
                    lang_list.push(Language::English);
                }
                if langs.fr {
                    lang_list.push(Language::Francais);
                }
                if langs.es {
                    lang_list.push(Language::Espanol);
                }
                if langs.pt {
                    lang_list.push(Language::Portugues);
                }

                if lang_list.is_empty() {
                    fb.set_status_type("error".into());
                    fb.set_status_message("Bitte mindestens eine Sprache wählen.".into());
                    return;
                }

                fb.set_status_type("pending".into());
                fb.set_status_message("Export läuft...".into());

                let cats = fb.get_categories();
                let counts = [
                    cats.cat1 as u16,
                    cats.cat2 as u16,
                    cats.cat3 as u16,
                    cats.cat4 as u16,
                    cats.cat5 as u16,
                    cats.cat6 as u16,
                    cats.cat7 as u16,
                    cats.cat8 as u16,
                ];

                let sheet_prot = if fb.get_protect_sheet() {
                    let sp = fb.get_sheet_permissions();
                    Some(
                        SheetProtection::new()
                            .with_password(fb.get_sheet_password().to_string())
                            .allow_select_locked_cells(sp.select_locked)
                            .allow_select_unlocked_cells(sp.select_unlocked)
                            .allow_format_cells(sp.format_cells)
                            .allow_format_columns(sp.format_columns)
                            .allow_format_rows(sp.format_rows)
                            .allow_insert_columns(sp.insert_columns)
                            .allow_insert_rows(sp.insert_rows)
                            .allow_insert_hyperlinks(sp.insert_hyperlinks)
                            .allow_delete_columns(sp.delete_columns)
                            .allow_delete_rows(sp.delete_rows)
                            .allow_sort(sp.sort)
                            .allow_autofilter(sp.autofilter)
                            .allow_pivot_tables(sp.pivot_tables)
                            .allow_edit_objects(sp.edit_objects)
                            .allow_edit_scenarios(sp.edit_scenarios)
                            .allow_contents(sp.contents),
                    )
                } else {
                    None
                };

                let workbook_pw = if fb.get_protect_workbook() {
                    Some(fb.get_workbook_password().to_string())
                } else {
                    None
                };

                match generate_excel(
                    lang_list,
                    counts,
                    sheet_prot,
                    workbook_pw.as_deref(),
                    fb.get_hide_columns(),
                    fb.get_hide_lang_sheet(),
                    &folder,
                    &version,
                ) {
                    Ok(count) => {
                        fb.set_status_type("success".into());
                        fb.set_status_message(
                            format!("{count} Datei(en) erfolgreich erstellt!").into(),
                        );
                    }
                    Err(e) => {
                        fb.set_status_type("error".into());
                        fb.set_status_message(format!("Fehler: {e}").into());
                    }
                }
            }
        }
    });

    // ==========================================
    // Budget-to-FB Callbacks
    // ==========================================

    ui.global::<BudgetState>().on_select_src({
        let ui_handle = ui.as_weak();
        move || {
            if let Some(ui) = ui_handle.upgrade() {
                if let Some(path) = rfd::FileDialog::new().pick_folder() {
                    ui.global::<BudgetState>()
                        .set_src_folder(path.to_string_lossy().to_string().into());
                }
            }
        }
    });

    ui.global::<BudgetState>().on_select_out({
        let ui_handle = ui.as_weak();
        move || {
            if let Some(ui) = ui_handle.upgrade() {
                if let Some(path) = rfd::FileDialog::new().pick_folder() {
                    ui.global::<BudgetState>()
                        .set_out_folder(path.to_string_lossy().to_string().into());
                }
            }
        }
    });

    ui.global::<BudgetState>().on_scan({
        let ui_handle = ui.as_weak();
        move || {
            if let Some(ui) = ui_handle.upgrade() {
                let b2f = ui.global::<BudgetState>();
                if b2f.get_src_folder().is_empty() {
                    b2f.set_status_type("error".into());
                    b2f.set_status_message("Bitte Quellordner wählen.".into());
                    return;
                }
                b2f.set_status_type("success".into());
                b2f.set_status_message("Scan-Funktion noch nicht implementiert.".into());
            }
        }
    });

    ui.global::<BudgetState>().on_do_export_txt({
        let ui_handle = ui.as_weak();
        move || {
            if let Some(ui) = ui_handle.upgrade() {
                let b2f = ui.global::<BudgetState>();
                b2f.set_status_type("success".into());
                b2f.set_status_message("TXT-Export noch nicht implementiert.".into());
            }
        }
    });

    ui.global::<BudgetState>().on_do_export_excel({
        let ui_handle = ui.as_weak();
        move || {
            if let Some(ui) = ui_handle.upgrade() {
                let b2f = ui.global::<BudgetState>();
                b2f.set_status_type("success".into());
                b2f.set_status_message("Excel-Export noch nicht implementiert.".into());
            }
        }
    });

    ui.global::<BudgetState>().on_dismiss_status({
        let ui_handle = ui.as_weak();
        move || {
            if let Some(ui) = ui_handle.upgrade() {
                let b2f = ui.global::<BudgetState>();
                b2f.set_status_type("idle".into());
                b2f.set_status_message("".into());
            }
        }
    });

    ui.global::<BudgetState>().on_do_reset({
        let ui_handle = ui.as_weak();
        move || {
            if let Some(ui) = ui_handle.upgrade() {
                apply_b2f_defaults(&ui);
            }
        }
    });

    // ==========================================
    // Folder-Creation Callbacks
    // ==========================================

    ui.global::<FolderState>().on_select_folder({
        let ui_handle = ui.as_weak();
        move || {
            if let Some(ui) = ui_handle.upgrade() {
                if let Some(path) = rfd::FileDialog::new().pick_folder() {
                    ui.global::<FolderState>()
                        .set_target_folder(path.to_string_lossy().to_string().into());
                }
            }
        }
    });

    ui.global::<FolderState>().on_reset({
        let ui_handle = ui.as_weak();
        move || {
            if let Some(ui) = ui_handle.upgrade() {
                apply_folder_defaults(&ui);
            }
        }
    });

    ui.global::<FolderState>().on_dismiss_status({
        let ui_handle = ui.as_weak();
        move || {
            if let Some(ui) = ui_handle.upgrade() {
                let fs = ui.global::<FolderState>();
                fs.set_status_type("idle".into());
                fs.set_status_message("".into());
            }
        }
    });

    ui.global::<FolderState>().on_create_folders({
        let ui_handle = ui.as_weak();
        move || {
            if let Some(ui) = ui_handle.upgrade() {
                let fs = ui.global::<FolderState>();
                if fs.get_target_folder().is_empty() {
                    fs.set_status_type("error".into());
                    fs.set_status_message("Bitte Zielordner wählen.".into());
                    return;
                }
                if fs.get_project_name().is_empty() {
                    fs.set_status_type("error".into());
                    fs.set_status_message("Bitte Projektnamen angeben.".into());
                    return;
                }

                fs.set_status_type("success".into());
                fs.set_status_message(
                    "Ordnerstruktur-Erstellung noch nicht implementiert.".into(),
                );
            }
        }
    });

    ui.run()
}

// ==========================================
// Excel-Generierung
// ==========================================

#[allow(clippy::too_many_arguments)]
fn generate_excel(
    langs: Vec<Language>,
    counts: [u16; 8],
    sheet_prot: Option<SheetProtection>,
    workbook_pw: Option<&str>,
    hide_columns: bool,
    hide_lang_sheet: bool,
    folder: &str,
    version: &str,
) -> Result<usize, fb_generator::ReportError> {
    let folder_path = PathBuf::from(folder);
    if !folder_path.exists() {
        std::fs::create_dir_all(&folder_path)?;
    }

    // Mappenschutz-Hash vorab berechnen (~25ms Ersparnis pro Datei)
    let precomputed_hash = workbook_pw
        .filter(|pw| !pw.is_empty())
        .map(fb_generator::precompute_hash);

    let mut count = 0;

    for lang in langs {
        let header = ReportHeader::builder()
            .language(lang)
            .version(version)
            .build();

        let mut body_builder = ReportBody::builder();
        for (i, &pos_count) in counts.iter().enumerate() {
            let category = (i + 1) as u8;
            if pos_count > 0 {
                let positions = (0..pos_count).map(|_| PositionEntry::builder().build());
                body_builder = body_builder.add_positions(category, positions);
            } else {
                body_builder =
                    body_builder.set_header_input(category, PositionEntry::builder().build());
            }
        }

        let mut options_builder = ReportOptions::builder();
        if let Some(ref prot) = sheet_prot {
            options_builder = options_builder.sheet_protection(prot.clone());
        }
        if let Some(pw) = workbook_pw {
            options_builder = options_builder.workbook_password(pw);
        }
        if hide_columns {
            options_builder = options_builder.hide_columns_qv(true);
        }
        if hide_lang_sheet {
            options_builder = options_builder.hide_language_sheet(true);
        }

        let config = ReportConfig::builder()
            .header(header)
            .body(body_builder.build())
            .options(options_builder.build())
            .build();

        let lang_code = match lang {
            Language::Deutsch => "de",
            Language::English => "en",
            Language::Francais => "fr",
            Language::Espanol => "es",
            Language::Portugues => "po",
        };

        let path = folder_path.join(format!("{version}_{lang_code}.xlsx"));

        if let Some(ref hash) = precomputed_hash {
            config.write_to_precomputed(&path, hash)?;
        } else {
            config.write_to(&path)?;
        }

        count += 1;
    }

    Ok(count)
}
