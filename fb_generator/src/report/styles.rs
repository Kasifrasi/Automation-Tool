//! Central format definitions for the financial report.
//!
//! - [`ReportStyles`]: Base styles (colors, borders, standard formats)
//! - [`SectionStyles`]: Derived styles for specific sections
//! - [`FormatMatrix`]: Per-cell format lookup
//! - [`BodyStyles`]: Styles for the dynamic body section

use crate::report::body::BodyLayout;
use rust_xlsxwriter::{Color, Format, FormatAlign, FormatBorder, FormatPattern};
use std::collections::HashMap;

// ============================================================================
// FormatMatrix
// ============================================================================

/// Per-cell format lookup for the report grid.
///
/// Default format is unlocked (Arial 10); only formula cells are locked.
/// Locked variants are pre-cached on insertion to avoid clones in the hot path.
pub struct FormatMatrix {
    formats: HashMap<(u32, u16), Format>,
    locked_formats: HashMap<(u32, u16), Format>,
}

impl FormatMatrix {
    pub fn new() -> Self {
        Self {
            formats: HashMap::new(),
            locked_formats: HashMap::new(),
        }
    }

    /// Sets the format for a cell (automatically caches both unlocked and locked variants).
    pub fn set(&mut self, row: u32, col: u16, format: &Format) {
        let key = (row, col);
        let unlocked = format.clone().set_unlocked();
        let locked = unlocked.clone().set_locked();
        self.formats.insert(key, unlocked);
        self.locked_formats.insert(key, locked);
    }

    /// Returns the unlocked format for a cell.
    pub fn get(&self, row: u32, col: u16) -> Option<&Format> {
        self.formats.get(&(row, col))
    }

    /// Returns the locked format for a cell (for formula cells).
    pub fn get_locked(&self, row: u32, col: u16) -> Option<&Format> {
        self.locked_formats.get(&(row, col))
    }
}

impl Default for FormatMatrix {
    fn default() -> Self {
        Self::new()
    }
}

// ============================================================================
// ReportStyles: Base styles
// ============================================================================

pub struct ReportStyles {
    // --- Colors ---
    pub fill_input: Color,
    pub fill_value: Color,
    pub fill_summary: Color,
    pub fl_orange: Color,

    // --- Borders ---
    pub border_thin: FormatBorder,
    pub border_medium: FormatBorder,
    pub border_dotted: FormatBorder,

    // --- Base ---
    pub base: Format,

    // --- General Alignments ---
    pub left_center: Format,
    pub left_center_bold: Format,
    pub center_center: Format,
    pub center_center_bold: Format,
    pub center_center_bold_right: Format,

    // --- Header Specifics ---
    pub header_label: Format,
    pub header_suffix: Format,
    pub input_orange_dashed: Format,
    pub header_right_box_top: Format,
    pub header_right_box_body: Format,
    pub link_style: Format,

    // --- Table Header ---
    pub table_header_base: Format,

    // --- Body ---
    pub body_label: Format,
    pub body_value: Format,
    pub body_input: Format,
    pub body_calc: Format,
    pub body_pct: Format,
    pub body_right: Format,

    // --- Table Income Total (Row 19) ---
    pub table_total_label: Format,
    pub table_total_value: Format,
    pub table_total_pct: Format,
    pub table_total_right: Format,

    // --- Right Panel Base ---
    pub rp_index: Format,
    pub rp_text: Format,
    pub rp_date: Format,
    pub rp_number: Format,
    pub rp_calc: Format,

    // --- Footer ---
    pub footer_center_wrap: Format,
    pub footer_gray_center: Format,
    pub footer_number_right: Format,
    pub footer_input: Format,
}

impl ReportStyles {
    pub fn new() -> Self {
        // Colors
        let fill_input = Color::RGB(0xFFFAE5);
        let fill_value = Color::RGB(0xF2F2F2);
        let fill_summary = Color::RGB(0xD9D9D9);
        let fl_orange = Color::RGB(0xFFCC00);

        // Borders
        let border_thin = FormatBorder::Thin;
        let border_medium = FormatBorder::Medium;
        let border_dotted = FormatBorder::Dotted;
        let border_dashed = FormatBorder::Dashed;

        // Base format (Arial 10, locked — formula cells are always protected)
        let base = Format::new().set_font_name("Arial").set_font_size(10.0);

        // Base Alignments
        let left_center = base
            .clone()
            .set_align(FormatAlign::Left)
            .set_align(FormatAlign::VerticalCenter);

        let left_center_bold = left_center.clone().set_bold();

        let center_center = base
            .clone()
            .set_align(FormatAlign::Center)
            .set_align(FormatAlign::VerticalCenter);

        let center_center_bold = center_center.clone().set_bold();
        let center_center_bold_right = center_center_bold.clone().set_border_right(border_thin);

        let left_top_wrap = base
            .clone()
            .set_align(FormatAlign::Left)
            .set_align(FormatAlign::Top)
            .set_text_wrap();

        // Header Specifics
        let header_label = left_center_bold.clone();
        let header_suffix = left_center.clone().set_font_color(Color::White);

        let input_orange_dashed = center_center
            .clone()
            .set_background_color(fl_orange)
            .set_border_bottom(border_dashed);

        let header_right_box_top = base
            .clone()
            .set_bold()
            .set_align(FormatAlign::Left)
            .set_border_top(border_medium)
            .set_border_left(border_medium);

        let header_right_box_body = left_top_wrap
            .clone()
            .set_border_left(border_medium)
            .set_border_right(border_medium);

        let link_style = base
            .clone()
            .set_align(FormatAlign::Left)
            .set_hyperlink()
            .set_border_left(border_medium)
            .set_border_right(border_medium)
            .set_border_bottom(border_medium);

        let table_header_base = base
            .clone()
            .set_align(FormatAlign::Center)
            .set_align(FormatAlign::VerticalCenter)
            .set_text_wrap()
            .set_border_top(border_medium);

        // Body Styles
        let body_label = base
            .clone()
            .set_bold()
            .set_align(FormatAlign::Left)
            .set_border_left(border_medium)
            .set_border_right(border_thin)
            .set_border_top(border_thin)
            .set_border_bottom(border_thin);

        let body_value = base
            .clone()
            .set_border(border_thin)
            .set_background_color(fill_value)
            .set_num_format("#,##0.00");

        let body_input = base
            .clone()
            .set_border(border_thin)
            .set_background_color(fill_input)
            .set_num_format("#,##0.00");

        let body_calc = base
            .clone()
            .set_border(border_thin)
            .set_background_color(fill_input)
            .set_num_format("#,##0.00");

        let body_pct = base.clone().set_border(border_thin).set_num_format("0%");

        let body_right = base
            .clone()
            .set_border(border_thin)
            .set_background_color(fill_input)
            .set_border_right(border_medium);

        // Table Income Total (Row 19) — aggregate row closing the 5-row static income table
        let table_total_label = base
            .clone()
            .set_background_color(fill_summary)
            .set_bold()
            .set_align(FormatAlign::Left)
            .set_border_left(border_medium)
            .set_border_right(border_thin)
            .set_border_top(border_thin)
            .set_border_bottom(border_medium);

        let table_total_value = base
            .clone()
            .set_background_color(fill_summary)
            .set_border(border_thin)
            .set_border_bottom(border_medium)
            .set_num_format("#,##0.00");

        let table_total_pct = base
            .clone()
            .set_background_color(fill_summary)
            .set_border(border_thin)
            .set_border_bottom(border_medium)
            .set_num_format("0%");

        let table_total_right = base
            .clone()
            .set_background_color(fill_summary)
            .set_border(border_thin)
            .set_border_bottom(border_medium)
            .set_border_right(border_medium)
            .set_num_format("#,##0.00");

        // Right Panel Base Styles
        let rp_index = base.clone().set_align(FormatAlign::Right);
        let rp_text = base.clone().set_align(FormatAlign::Left);

        let rp_date = base
            .clone()
            .set_align(FormatAlign::Center)
            .set_num_format_index(14)
            .set_background_color(fill_input);

        let rp_number = base
            .clone()
            .set_align(FormatAlign::Right)
            .set_num_format("#,##0.00")
            .set_background_color(fill_input);

        let rp_calc = base
            .clone()
            .set_align(FormatAlign::Right)
            .set_num_format("0.0000");

        // Footer Styles
        let footer_center_wrap = center_center_bold.clone().set_text_wrap();

        let footer_gray_center = center_center.clone().set_font_color(Color::RGB(0x808080));

        let footer_number_right = base
            .clone()
            .set_align(FormatAlign::Right)
            .set_align(FormatAlign::VerticalCenter)
            .set_num_format("#,##0.00");

        let footer_input = footer_number_right.clone().set_background_color(fill_input);

        Self {
            fill_input,
            fill_value,
            fill_summary,
            fl_orange,
            border_thin,
            border_medium,
            border_dotted,
            base,
            left_center,
            left_center_bold,
            center_center,
            center_center_bold,
            center_center_bold_right,
            header_label,
            header_suffix,
            input_orange_dashed,
            header_right_box_top,
            header_right_box_body,
            link_style,
            table_header_base,
            body_label,
            body_value,
            body_input,
            body_calc,
            body_pct,
            body_right,
            table_total_label,
            table_total_value,
            table_total_pct,
            table_total_right,
            rp_index,
            rp_text,
            rp_date,
            rp_number,
            rp_calc,
            footer_center_wrap,
            footer_gray_center,
            footer_number_right,
            footer_input,
        }
    }
}

impl Default for ReportStyles {
    fn default() -> Self {
        Self::new()
    }
}

// ============================================================================
// SectionStyles: Derived styles for specific sections
// ============================================================================

pub struct SectionStyles {
    // Header Section (Row 0-9)
    pub fmt_top_med: Format,
    pub fmt_top_right_med: Format,
    pub fmt_d5: Format,
    pub fmt_d6: Format,
    pub fmt_row7_base: Format,
    pub fmt_row7_date: Format,
    pub orange_dotted: Format,
    pub value_dotted: Format,
    pub input_dotted: Format,

    // Table Header Section (Row 10-13)
    pub th_b: Format,
    pub th_c_top: Format,
    pub th_d: Format,
    pub th_d_bold: Format,
    pub th_h: Format,
    pub th_side: Format,
    pub th_side_bold: Format,
    pub th_bot_side: Format,
    pub th_bot_right: Format,

    // Table Body Section (Row 14-19)
    pub body_label_top: Format,
    pub body_label_no_h: Format,

    // Right Panel Header (Row 10-12)
    pub rp_hdr_left: Format,
    pub rp_hdr_mid: Format,
    pub rp_hdr_mid_top: Format,
    pub rp_hdr_right: Format,
    pub rp_hdr_right_top: Format,

    // Right Panel Body (Row 13-30)
    pub rp_idx: Format,
    pub rp_idx_last: Format,
    pub rp_txt: Format,
    pub rp_txt_last: Format,
    pub rp_date: Format,
    pub rp_date_last: Format,
    pub rp_num: Format,
    pub rp_num_last: Format,
    pub rp_calc: Format,
    pub rp_calc_last: Format,

    // Footer Section
    pub ft_b_top_left: Format,           // B: top+left border
    pub ft_c_top: Format,                // C: top border
    pub ft_d_top_right: Format,          // D: top+right thin border
    pub ft_e_merged_top: Format,         // E merged: top+left thin+right medium, center wrap
    pub ft_bcd_merged: Format,           // B:D merged: left medium+right thin, center bold
    pub ft_e_center: Format,             // E: left thin+right medium, center
    pub ft_b_left: Format,               // B: left medium only
    pub ft_d_right: Format,              // D: right thin only
    pub ft_e_right: Format,              // E: right medium only
    pub ft_b_label_box: Format,          // B: left+top+bottom, bold (label in box)
    pub ft_c_box: Format,                // C: top+bottom
    pub ft_d_box: Format,                // D: top+bottom (no right - prevents E left)
    pub ft_e_number_box: Format,         // E: top+bottom+right medium, number format
    pub ft_e_gray_box: Format,           // E: right medium, gray center (OK check)
    pub ft_b_input_label_top: Format,    // B: left+top, label for input row
    pub ft_c_input_top: Format,          // C: top
    pub ft_d_input_top_right: Format,    // D: top+right thin
    pub ft_e_input_top: Format,          // E: input format with top+left thin+right medium
    pub ft_b_input_label_bottom: Format, // B: left+top+bottom medium
    pub ft_c_input_bottom: Format,       // C: top+bottom medium
    pub ft_d_input_bottom_right: Format, // D: top+bottom medium+right thin
    pub ft_e_input_bottom: Format,       // E: input format with all borders
    pub ft_signature: Format,            // Signature: bold with top thin
    pub ft_signature_top: Format,        // Signature row: normal with top thin

    // Pre-Body Section (Rows 22-25)
    pub pb_lbl_b: Format,    // B23: blank — medium-left+top, no right border
    pub pb_lbl_c: Format,    // C23: blank — medium-top, thin-right
    pub pb_val: Format,      // D23, F23, G23: merge head — center+wrap, medium-top, thin-left/right
    pub pb_val_bold: Format, // E23: like pb_val + bold
    pub pb_right: Format,    // H23: like pb_val + medium-right
    pub pb_mid: Format,      // B25: B:C merge — center+wrap, medium-left, thin-right
    pub pb_mid_bold: Format, // B24: like pb_mid + bold
    pub pb_bot: Format,      // B26: B:C merge (blank) — medium-left+bottom, thin-right
}

impl SectionStyles {
    pub fn new(s: &ReportStyles) -> Self {
        // Header Section
        let fmt_top_med = s.base.clone().set_border_top(s.border_medium);
        let fmt_top_right_med = s
            .base
            .clone()
            .set_border_top(s.border_medium)
            .set_border_right(s.border_medium);

        let fmt_d5 = s
            .base
            .clone()
            .set_background_color(s.fill_input)
            .set_border_bottom(s.border_thin);

        let fmt_d6 = s
            .base
            .clone()
            .set_background_color(s.fill_input)
            .set_align(FormatAlign::Left)
            .set_align(FormatAlign::Bottom)
            .set_text_wrap()
            .set_border_bottom(s.border_thin);

        let fmt_row7_base = s.base.clone().set_border_bottom(s.border_thin);
        let fmt_row7_date = fmt_row7_base
            .clone()
            .set_background_color(s.fill_input)
            .set_num_format_index(14)
            .set_align(FormatAlign::Left);

        let orange_dotted = s
            .base
            .clone()
            .set_background_color(s.fl_orange)
            .set_border(s.border_dotted);
        let value_dotted = s
            .base
            .clone()
            .set_background_color(s.fill_value)
            .set_border(s.border_dotted);
        let input_dotted = s
            .base
            .clone()
            .set_background_color(s.fill_input)
            .set_border(s.border_dotted);

        // Table Header Section
        let th_b = s.table_header_base.clone().set_border_left(s.border_medium);
        let th_c_top = s.table_header_base.clone().set_border_right(s.border_thin);

        let th_d = s
            .table_header_base
            .clone()
            .set_border_left(s.border_thin)
            .set_border_right(s.border_thin);

        let th_d_bold = th_d.clone().set_bold();

        let th_h = s
            .table_header_base
            .clone()
            .set_border_left(s.border_thin)
            .set_border_right(s.border_medium);

        let th_side = s
            .base
            .clone()
            .set_align(FormatAlign::Center)
            .set_align(FormatAlign::VerticalCenter)
            .set_border_left(s.border_medium)
            .set_border_right(s.border_thin);

        let th_side_bold = th_side.clone().set_bold();

        let th_bot_side = s
            .base
            .clone()
            .set_align(FormatAlign::Center)
            .set_align(FormatAlign::VerticalCenter)
            .set_border_left(s.border_medium)
            .set_border_bottom(s.border_thin);

        let th_bot_right = s
            .base
            .clone()
            .set_border_bottom(s.border_thin)
            .set_border_right(s.border_thin);

        // Table Body Section
        let body_label_top = s.body_label.clone().set_border_bottom(FormatBorder::None);
        let body_label_no_h = s
            .body_label
            .clone()
            .set_border_top(FormatBorder::None)
            .set_border_bottom(FormatBorder::None);

        // Right Panel Header
        let rp_hdr_left = s.base.clone().set_border_left(s.border_thin);
        let rp_hdr_mid = s.base.clone();
        let rp_hdr_mid_top = rp_hdr_mid.clone().set_border_top(s.border_thin);
        let rp_hdr_right = s.base.clone().set_border_right(s.border_thin);
        let rp_hdr_right_top = rp_hdr_right.clone().set_border_top(s.border_thin);

        // Right Panel Body
        let rp_idx = s.rp_index.clone().set_border_left(s.border_thin);
        let rp_idx_last = rp_idx.clone().set_border_bottom(s.border_thin);
        let rp_txt = s.rp_text.clone().set_border_right(s.border_thin);
        let rp_txt_last = rp_txt.clone().set_border_bottom(s.border_thin);

        let rp_date = s
            .rp_date
            .clone()
            .set_border_left(s.border_thin)
            .set_border_right(s.border_thin)
            .set_border_top(s.border_thin);
        let rp_date_last = rp_date.clone().set_border_bottom(s.border_thin);

        let rp_num = s
            .rp_number
            .clone()
            .set_border_left(s.border_thin)
            .set_border_right(s.border_thin)
            .set_border_top(s.border_thin);
        let rp_num_last = rp_num.clone().set_border_bottom(s.border_thin);

        let rp_calc = s
            .rp_calc
            .clone()
            .set_border_left(s.border_thin)
            .set_border_right(s.border_thin)
            .set_border_top(s.border_thin);
        let rp_calc_last = rp_calc.clone().set_border_bottom(s.border_thin);

        // Footer Section
        let ft_b_top_left = s
            .left_center
            .clone()
            .set_border_top(s.border_medium)
            .set_border_left(s.border_medium);

        let ft_c_top = s.left_center.clone().set_border_top(s.border_medium);

        let ft_d_top_right = s
            .left_center
            .clone()
            .set_border_top(s.border_medium)
            .set_border_right(s.border_thin);

        let ft_e_merged_top = s
            .footer_center_wrap
            .clone()
            .set_border_top(s.border_medium)
            .set_border_left(s.border_thin)
            .set_border_right(s.border_medium);

        let ft_bcd_merged = s
            .left_center_bold
            .clone()
            .set_align(FormatAlign::Center)
            .set_border_left(s.border_medium)
            .set_border_right(s.border_thin);

        let ft_e_center = s
            .center_center
            .clone()
            .set_border_left(s.border_thin)
            .set_border_right(s.border_medium);

        let ft_b_left = s.left_center.clone().set_border_left(s.border_medium);

        let ft_d_right = s.left_center.clone().set_border_right(s.border_thin);

        let ft_e_right = s.left_center.clone().set_border_right(s.border_medium);

        let ft_b_label_box = s
            .left_center_bold
            .clone()
            .set_border_left(s.border_medium)
            .set_border_top(s.border_thin)
            .set_border_bottom(s.border_thin);

        let ft_c_box = s
            .left_center
            .clone()
            .set_border_top(s.border_thin)
            .set_border_bottom(s.border_thin);

        let ft_d_box = s
            .center_center
            .clone()
            .set_border_top(s.border_thin)
            .set_border_bottom(s.border_thin);

        let ft_e_number_box = s
            .footer_number_right
            .clone()
            .set_border_top(s.border_thin)
            .set_border_bottom(s.border_thin)
            .set_border_right(s.border_medium);

        let ft_e_gray_box = s
            .footer_gray_center
            .clone()
            .set_border_right(s.border_medium);

        let ft_b_input_label_top = s
            .left_center
            .clone()
            .set_border_left(s.border_medium)
            .set_border_top(s.border_thin);

        let ft_c_input_top = s.left_center.clone().set_border_top(s.border_thin);

        let ft_d_input_top_right = s
            .left_center
            .clone()
            .set_border_top(s.border_thin)
            .set_border_right(s.border_thin);

        let ft_e_input_top = s
            .footer_input
            .clone()
            .set_border_top(s.border_thin)
            .set_border_left(s.border_thin)
            .set_border_right(s.border_medium);

        let ft_b_input_label_bottom = s
            .left_center
            .clone()
            .set_border_left(s.border_medium)
            .set_border_top(s.border_thin)
            .set_border_bottom(s.border_medium);

        let ft_c_input_bottom = s
            .left_center
            .clone()
            .set_border_top(s.border_thin)
            .set_border_bottom(s.border_medium);

        let ft_d_input_bottom_right = s
            .left_center
            .clone()
            .set_border_top(s.border_thin)
            .set_border_bottom(s.border_medium)
            .set_border_right(s.border_thin);

        let ft_e_input_bottom = s
            .footer_input
            .clone()
            .set_border_top(s.border_thin)
            .set_border_bottom(s.border_medium)
            .set_border_left(s.border_thin)
            .set_border_right(s.border_medium);

        let ft_signature = s.left_center_bold.clone().set_border_top(s.border_thin);

        let ft_signature_top = s.left_center.clone().set_border_top(s.border_thin);

        // Pre-Body Section (Rows 22-25)
        let pb_val = s
            .base
            .clone()
            .set_align(FormatAlign::Center)
            .set_align(FormatAlign::VerticalCenter)
            .set_text_wrap()
            .set_border_top(s.border_medium)
            .set_border_left(s.border_thin)
            .set_border_right(s.border_thin);

        let pb_val_bold = pb_val.clone().set_bold();

        let pb_right = pb_val.clone().set_border_right(s.border_medium);

        let pb_lbl_b = s
            .base
            .clone()
            .set_align(FormatAlign::Center)
            .set_align(FormatAlign::VerticalCenter)
            .set_text_wrap()
            .set_border_left(s.border_medium)
            .set_border_top(s.border_medium);

        let pb_lbl_c = s
            .base
            .clone()
            .set_align(FormatAlign::Center)
            .set_align(FormatAlign::VerticalCenter)
            .set_text_wrap()
            .set_border_top(s.border_medium)
            .set_border_right(s.border_thin);

        let pb_mid = s
            .base
            .clone()
            .set_align(FormatAlign::Center)
            .set_align(FormatAlign::VerticalCenter)
            .set_border_left(s.border_medium)
            .set_border_right(s.border_thin);

        let pb_mid_bold = pb_mid.clone().set_bold();

        let pb_bot = s
            .base
            .clone()
            .set_border_left(s.border_medium)
            .set_border_bottom(s.border_thin)
            .set_border_right(s.border_thin);

        Self {
            fmt_top_med,
            fmt_top_right_med,
            fmt_d5,
            fmt_d6,
            fmt_row7_base,
            fmt_row7_date,
            orange_dotted,
            value_dotted,
            input_dotted,
            th_b,
            th_c_top,
            th_d,
            th_d_bold,
            th_h,
            th_side,
            th_side_bold,
            th_bot_side,
            th_bot_right,
            body_label_top,
            body_label_no_h,
            rp_hdr_left,
            rp_hdr_mid,
            rp_hdr_mid_top,
            rp_hdr_right,
            rp_hdr_right_top,
            rp_idx,
            rp_idx_last,
            rp_txt,
            rp_txt_last,
            rp_date,
            rp_date_last,
            rp_num,
            rp_num_last,
            rp_calc,
            rp_calc_last,
            ft_b_top_left,
            ft_c_top,
            ft_d_top_right,
            ft_e_merged_top,
            ft_bcd_merged,
            ft_e_center,
            ft_b_left,
            ft_d_right,
            ft_e_right,
            ft_b_label_box,
            ft_c_box,
            ft_d_box,
            ft_e_number_box,
            ft_e_gray_box,
            ft_b_input_label_top,
            ft_c_input_top,
            ft_d_input_top_right,
            ft_e_input_top,
            ft_b_input_label_bottom,
            ft_c_input_bottom,
            ft_d_input_bottom_right,
            ft_e_input_bottom,
            ft_signature,
            ft_signature_top,
            pb_lbl_b,
            pb_lbl_c,
            pb_val,
            pb_val_bold,
            pb_right,
            pb_mid,
            pb_mid_bold,
            pb_bot,
        }
    }
}

// ============================================================================
// build_format_matrix: Static cell format assignment
// ============================================================================

pub fn build_format_matrix(styles: &ReportStyles, sec: &SectionStyles) -> FormatMatrix {
    let mut m = FormatMatrix::new();

    // ========================================================================
    // HEADER (Rows 0-9)
    // ========================================================================

    // Row 0
    m.set(0, 1, &styles.header_label);
    m.set(0, 9, &styles.header_right_box_top);
    m.set(0, 10, &sec.fmt_top_med);
    m.set(0, 11, &sec.fmt_top_med);
    m.set(0, 12, &sec.fmt_top_med);
    m.set(0, 13, &sec.fmt_top_med);
    m.set(0, 14, &sec.fmt_top_right_med);

    // Row 1
    m.set(1, 1, &styles.header_suffix);
    m.set(1, 3, &styles.left_center);
    m.set(1, 4, &styles.input_orange_dashed);
    m.set(1, 9, &styles.header_right_box_body);

    // Row 2
    m.set(2, 1, &styles.header_label);
    m.set(2, 3, &styles.left_center);
    m.set(2, 4, &styles.input_orange_dashed);

    // Row 3
    m.set(3, 9, &styles.link_style);

    // Row 4
    m.set(4, 1, &styles.left_center);
    m.set(4, 3, &sec.fmt_d5);

    // Row 5
    m.set(5, 1, &styles.left_center);
    m.set(5, 3, &sec.fmt_d6);

    // Row 6
    m.set(6, 9, &sec.orange_dotted);

    // Row 7
    m.set(7, 1, &styles.left_center);
    m.set(7, 3, &sec.fmt_row7_base);
    m.set(7, 4, &sec.fmt_row7_date);
    m.set(7, 5, &sec.fmt_row7_base);
    m.set(7, 6, &sec.fmt_row7_date);
    m.set(7, 9, &sec.value_dotted);

    // Row 8
    m.set(8, 1, &styles.left_center);
    m.set(8, 3, &sec.fmt_row7_base);
    m.set(8, 4, &sec.fmt_row7_date);
    m.set(8, 5, &sec.fmt_row7_base);
    m.set(8, 6, &sec.fmt_row7_date);
    m.set(8, 9, &sec.input_dotted);

    // ========================================================================
    // TABLE HEADER (Rows 10-13)
    // ========================================================================

    // Row 10
    m.set(10, 1, &sec.th_b);
    m.set(10, 2, &sec.th_c_top);
    m.set(10, 3, &sec.th_d);
    m.set(10, 4, &sec.th_d_bold);
    m.set(10, 5, &sec.th_d);
    m.set(10, 6, &sec.th_d);
    m.set(10, 7, &sec.th_h);

    // Row 11
    m.set(11, 1, &sec.th_side_bold);

    // Row 12
    m.set(12, 1, &sec.th_side);
    m.set(12, 11, &styles.center_center_bold);
    m.set(12, 12, &styles.center_center_bold);
    m.set(12, 13, &styles.center_center_bold);
    m.set(12, 14, &styles.center_center_bold_right);
    m.set(12, 18, &styles.center_center_bold);
    m.set(12, 19, &styles.center_center_bold);
    m.set(12, 20, &styles.center_center_bold);
    m.set(12, 21, &styles.center_center_bold_right);

    // Row 13
    m.set(13, 1, &sec.th_bot_side);
    m.set(13, 2, &sec.th_bot_right);

    // ========================================================================
    // TABLE BODY (Rows 14-18)
    // ========================================================================

    // Row 14 (first data row)
    m.set(14, 1, &sec.body_label_top);
    m.set(14, 3, &styles.body_value);
    m.set(14, 4, &styles.body_input);
    m.set(14, 5, &styles.body_calc);
    m.set(14, 6, &styles.body_pct);
    m.set(14, 7, &styles.body_right);

    // Rows 15-18
    for row in 15..=18 {
        m.set(row, 1, &sec.body_label_no_h);
        m.set(row, 3, &styles.body_value);
        m.set(row, 4, &styles.body_input);
        m.set(row, 5, &styles.body_calc);
        m.set(row, 6, &styles.body_pct);
        m.set(row, 7, &styles.body_right);
    }

    // ========================================================================
    // TABLE INCOME TOTAL (Row 19)
    // ========================================================================

    m.set(19, 1, &styles.table_total_label);
    m.set(19, 3, &styles.table_total_value);
    m.set(19, 4, &styles.table_total_value);
    m.set(19, 5, &styles.table_total_value);
    m.set(19, 6, &styles.table_total_pct);
    m.set(19, 7, &styles.table_total_right);

    // ========================================================================
    // RIGHT PANEL (Row 10-30, Cols J-O and Q-V)
    // ========================================================================

    // Row 10 (J11-O11, Q11-V11): header with top border
    let j11_fmt = styles
        .left_center_bold
        .clone()
        .set_border_left(styles.border_thin)
        .set_border_top(styles.border_thin);
    let q11_fmt = j11_fmt.clone();

    m.set(10, 9, &j11_fmt);
    m.set(10, 11, &sec.rp_hdr_mid_top);
    m.set(10, 12, &sec.rp_hdr_mid_top);
    m.set(10, 13, &sec.rp_hdr_mid_top);
    m.set(10, 14, &sec.rp_hdr_right_top);
    m.set(10, 16, &q11_fmt);
    m.set(10, 18, &sec.rp_hdr_mid_top);
    m.set(10, 19, &sec.rp_hdr_mid_top);
    m.set(10, 20, &sec.rp_hdr_mid_top);
    m.set(10, 21, &sec.rp_hdr_right_top);

    // Row 11 (J12-O12, Q12-V12): header without top border
    m.set(11, 9, &sec.rp_hdr_left);
    m.set(11, 10, &sec.rp_hdr_mid);
    m.set(11, 11, &sec.rp_hdr_mid);
    m.set(11, 12, &sec.rp_hdr_mid);
    m.set(11, 13, &sec.rp_hdr_mid);
    m.set(11, 14, &sec.rp_hdr_right);
    m.set(11, 16, &sec.rp_hdr_left);
    m.set(11, 17, &sec.rp_hdr_mid);
    m.set(11, 18, &sec.rp_hdr_mid);
    m.set(11, 19, &sec.rp_hdr_mid);
    m.set(11, 20, &sec.rp_hdr_mid);
    m.set(11, 21, &sec.rp_hdr_right);

    // Row 12 (J13, K13, Q13, R13): blanks only; L13-O13 and S13-V13 use center_center_bold
    m.set(12, 9, &sec.rp_hdr_left);
    m.set(12, 10, &sec.rp_hdr_mid);
    m.set(12, 16, &sec.rp_hdr_left);
    m.set(12, 17, &sec.rp_hdr_mid);

    // Rows 13-29: body without bottom border
    for row in 13..=29 {
        m.set(row, 9, &sec.rp_idx);
        m.set(row, 10, &sec.rp_txt);
        m.set(row, 11, &sec.rp_date);
        m.set(row, 12, &sec.rp_num);
        m.set(row, 13, &sec.rp_num);
        m.set(row, 14, &sec.rp_calc);
        m.set(row, 16, &sec.rp_idx);
        m.set(row, 17, &sec.rp_txt);
        m.set(row, 18, &sec.rp_date);
        m.set(row, 19, &sec.rp_num);
        m.set(row, 20, &sec.rp_num);
        m.set(row, 21, &sec.rp_calc);
    }

    // Row 30: body with bottom border
    m.set(30, 9, &sec.rp_idx_last);
    m.set(30, 10, &sec.rp_txt_last);
    m.set(30, 11, &sec.rp_date_last);
    m.set(30, 12, &sec.rp_num_last);
    m.set(30, 13, &sec.rp_num_last);
    m.set(30, 14, &sec.rp_calc_last);
    m.set(30, 16, &sec.rp_idx_last);
    m.set(30, 17, &sec.rp_txt_last);
    m.set(30, 18, &sec.rp_date_last);
    m.set(30, 19, &sec.rp_num_last);
    m.set(30, 20, &sec.rp_num_last);
    m.set(30, 21, &sec.rp_calc_last);

    m
}

// ============================================================================
// BodyStyles: Formats for the dynamic body section
// ============================================================================

/// Styles for the dynamic body section (cost categories).
pub struct BodyStyles {
    // === Category Header (bold, no fill) ===
    /// B: category number ("1."), left, bold, medium left border
    pub cat_header_b: Format,
    /// C: VLOOKUP label, left, bold
    pub cat_header_c: Format,
    /// D-F: value columns (empty in header)
    pub cat_header_value: Format,
    /// G: percent column (empty in header)
    pub cat_header_pct: Format,
    /// H: remarks, medium right border
    pub cat_header_h: Format,

    // === Position Row (input cells, yellow) ===
    /// B: position number ("1.1"), left
    pub pos_b: Format,
    /// C: description, left, input fill, wrap
    pub pos_c: Format,
    /// D: approved, gray fill (calculated or value)
    pub pos_d: Format,
    /// E-F: income, input fill
    pub pos_ef: Format,
    /// G: percent (formula)
    pub pos_g: Format,
    /// H: remarks, input fill, wrap, medium right
    pub pos_h: Format,

    // === Category Footer (gray fill, bold) ===
    /// B:C merged: sum label via VLOOKUP (no right border on B, merged)
    pub footer_bc: Format,
    /// D-F: SUM formulas
    pub footer_value: Format,
    /// G: ratio formula
    pub footer_pct: Format,
    /// H: empty, medium right
    pub footer_h: Format,

    // === Single-Row Category (like position, but with bold label) ===
    /// B: category number ("6."), left, bold
    pub single_b: Format,
    /// C: VLOOKUP label, bold
    pub single_c: Format,
    // D-H: same as pos_*

    // === Total Row (gray fill, bold, medium bottom) ===
    /// B:C merged: "Total" label
    pub total_bc: Format,
    /// D-F: SUM formulas
    pub total_value: Format,
    /// G: ratio formula
    pub total_pct: Format,
    /// H: empty, medium right + bottom
    pub total_h: Format,
}

impl BodyStyles {
    pub fn new(s: &ReportStyles) -> Self {
        let thin = s.border_thin;
        let medium = s.border_medium;

        // === Category Header ===
        let cat_header_b = s
            .base
            .clone()
            .set_align(FormatAlign::Left)
            .set_bold()
            .set_num_format("@")
            .set_border_left(medium)
            .set_border_top(thin)
            .set_border_bottom(thin)
            .set_border_right(thin);

        let cat_header_c = s
            .base
            .clone()
            .set_align(FormatAlign::Left)
            .set_bold()
            .set_border(thin);

        let cat_header_value = s
            .base
            .clone()
            .set_align(FormatAlign::Right)
            .set_bold()
            .set_text_wrap()
            .set_num_format("#,##0.00")
            .set_border(thin);

        let cat_header_pct = s
            .base
            .clone()
            .set_align(FormatAlign::Right)
            .set_bold()
            .set_num_format("0%")
            .set_border(thin);

        let cat_header_h = s
            .base
            .clone()
            .set_align(FormatAlign::Left)
            .set_bold()
            .set_border_left(thin)
            .set_border_top(thin)
            .set_border_bottom(thin)
            .set_border_right(medium);

        // === Position Row ===
        let pos_b = s
            .base
            .clone()
            .set_align(FormatAlign::Left)
            .set_num_format("@")
            .set_border_left(medium)
            .set_border_top(thin)
            .set_border_bottom(thin)
            .set_border_right(thin);

        let pos_c = s
            .base
            .clone()
            .set_background_color(s.fill_input)
            .set_text_wrap()
            .set_border(thin);

        let pos_d = s
            .base
            .clone()
            .set_background_color(s.fill_value)
            .set_text_wrap()
            .set_num_format("#,##0.00")
            .set_border(thin);

        let pos_ef = s
            .base
            .clone()
            .set_background_color(s.fill_input)
            .set_text_wrap()
            .set_num_format("#,##0.00")
            .set_border(thin);

        let pos_g = s
            .base
            .clone()
            .set_align(FormatAlign::Right)
            .set_num_format("0%")
            .set_border(thin);

        let pos_h = s
            .base
            .clone()
            .set_background_color(s.fill_input)
            .set_text_wrap()
            .set_border_left(thin)
            .set_border_top(thin)
            .set_border_bottom(thin)
            .set_border_right(medium);

        // === Category Footer (gray) ===
        let gray_bold = s
            .base
            .clone()
            .set_foreground_color(s.fill_summary)
            .set_pattern(FormatPattern::Solid)
            .set_bold();

        // B:C merged — left-aligned, left border (medium), right border (thin)
        let footer_bc = gray_bold
            .clone()
            .set_align(FormatAlign::Left)
            .set_border_left(medium)
            .set_border_top(thin)
            .set_border_bottom(thin)
            .set_border_right(thin);

        let footer_value = gray_bold
            .clone()
            .set_align(FormatAlign::Right)
            .set_num_format("#,##0.00")
            .set_border(thin);

        let footer_pct = gray_bold
            .clone()
            .set_align(FormatAlign::Right)
            .set_num_format("0%")
            .set_border(thin);

        let footer_h = gray_bold
            .clone()
            .set_border_left(thin)
            .set_border_top(thin)
            .set_border_bottom(thin)
            .set_border_right(medium);

        // === Single-Row Category ===
        let single_b = s
            .base
            .clone()
            .set_align(FormatAlign::Left)
            .set_bold()
            .set_num_format("@")
            .set_border_left(medium)
            .set_border_top(thin)
            .set_border_bottom(thin)
            .set_border_right(thin);

        // For category 8: gray text
        let single_c = s
            .base
            .clone()
            .set_align(FormatAlign::Left)
            .set_bold()
            .set_border(thin);

        // === Total Row (gray, medium bottom) ===
        // B:C merged — left-aligned, left border (medium), right border (thin)
        let total_bc = gray_bold
            .clone()
            .set_align(FormatAlign::Left)
            .set_border_left(medium)
            .set_border_top(thin)
            .set_border_bottom(medium)
            .set_border_right(thin);

        let total_value = gray_bold
            .clone()
            .set_align(FormatAlign::Right)
            .set_num_format("#,##0.00")
            .set_border_top(thin)
            .set_border_bottom(medium)
            .set_border_left(thin)
            .set_border_right(thin);

        let total_pct = gray_bold
            .clone()
            .set_align(FormatAlign::Right)
            .set_num_format("0%")
            .set_border_top(thin)
            .set_border_bottom(medium)
            .set_border_left(thin)
            .set_border_right(thin);

        let total_h = gray_bold
            .clone()
            .set_border_left(thin)
            .set_border_top(thin)
            .set_border_bottom(medium)
            .set_border_right(medium);

        Self {
            cat_header_b,
            cat_header_c,
            cat_header_value,
            cat_header_pct,
            cat_header_h,
            pos_b,
            pos_c,
            pos_d,
            pos_ef,
            pos_g,
            pos_h,
            footer_bc,
            footer_value,
            footer_pct,
            footer_h,
            single_b,
            single_c,
            total_bc,
            total_value,
            total_pct,
            total_h,
        }
    }
}

// ============================================================================
// apply_formats!: Declarative macro for column→format assignment
// ============================================================================

/// Sets a fixed column→format mapping for one or more rows.
///
/// # Variants
///
/// - `row $row` — single row
/// - `rows $range` — row range (for-loop)
///
/// # Example
///
/// ```ignore
/// apply_formats!(m, row header_row, {
///     1 => &body.cat_header_b,
///     2 => &body.cat_header_c,
/// });
/// apply_formats!(m, rows start..=end, {
///     1 => &body.pos_b,
///     2 => &body.pos_c,
/// });
/// ```
macro_rules! apply_formats {
    ($m:expr, row $row:expr, { $( $col:literal => $fmt:expr ),+ $(,)? }) => {
        $( $m.set($row, $col, $fmt); )+
    };
    ($m:expr, rows $rows:expr, { $( $col:literal => $fmt:expr ),+ $(,)? }) => {
        for _row in $rows { $( $m.set(_row, $col, $fmt); )+ }
    };
}

// ============================================================================
// extend_format_matrix_with_body
// ============================================================================

/// Extends the [`FormatMatrix`] with dynamic body formats.
pub fn extend_format_matrix_with_body(
    m: &mut FormatMatrix,
    styles: &ReportStyles,
    layout: &BodyLayout,
) {
    use crate::report::body::CategoryMode;

    let body = BodyStyles::new(styles);

    for cat in &layout.categories {
        match &cat.mode {
            // === Header-input mode (0 positions) ===
            CategoryMode::HeaderInput { row } => {
                // B: bold, C: bold (category 8 gray)
                m.set(*row, 1, &body.single_b);
                if cat.meta.num == 8 {
                    let gray_text = body.single_c.clone().set_font_color(Color::RGB(0xBFBFBF));
                    m.set(*row, 2, &gray_text);
                } else {
                    m.set(*row, 2, &body.single_c);
                }
                // D–H: input formats (same as position row)
                apply_formats!(m, row *row, {
                    3 => &body.pos_d,
                    4 => &body.pos_ef,
                    5 => &body.pos_ef,
                    6 => &body.pos_g,
                    7 => &body.pos_h,
                });
            }

            // === Positions mode (1+ positions) ===
            CategoryMode::WithPositions {
                header_row,
                positions,
                footer_row,
            } => {
                // Header row: B–H
                apply_formats!(m, row *header_row, {
                    1 => &body.cat_header_b,
                    2 => &body.cat_header_c,
                    3 => &body.cat_header_value,
                    4 => &body.cat_header_value,
                    5 => &body.cat_header_value,
                    6 => &body.cat_header_pct,
                    7 => &body.cat_header_h,
                });

                // Position rows: B–H
                apply_formats!(m, rows positions.start_row..=positions.end_row, {
                    1 => &body.pos_b,
                    2 => &body.pos_c,
                    3 => &body.pos_d,
                    4 => &body.pos_ef,
                    5 => &body.pos_ef,
                    6 => &body.pos_g,
                    7 => &body.pos_h,
                });

                // Footer row: B (B:C merged in writer), D–H
                apply_formats!(m, row *footer_row, {
                    1 => &body.footer_bc,
                    3 => &body.footer_value,
                    4 => &body.footer_value,
                    5 => &body.footer_value,
                    6 => &body.footer_pct,
                    7 => &body.footer_h,
                });
            }
        }
    }

    // === Total row: B (B:C merged in writer), D–H ===
    apply_formats!(m, row layout.total_row, {
        1 => &body.total_bc,
        3 => &body.total_value,
        4 => &body.total_value,
        5 => &body.total_value,
        6 => &body.total_pct,
        7 => &body.total_h,
    });
}

// ============================================================================
// extend_format_matrix_with_footer
// ============================================================================

/// Extends the [`FormatMatrix`] with footer formats.
pub fn extend_format_matrix_with_footer(
    m: &mut FormatMatrix,
    _styles: &ReportStyles,
    sec: &SectionStyles,
    start_row: u32,
) {
    let s = start_row;

    // Row 0: B–E (top borders, E starts merge with row 1)
    apply_formats!(m, row s, {
        1 => &sec.ft_b_top_left,
        2 => &sec.ft_c_top,
        3 => &sec.ft_d_top_right,
        4 => &sec.ft_e_merged_top,
    });

    // Row 1: B:D merged (E is in the merge from above)
    m.set(s + 1, 1, &sec.ft_bcd_merged);

    // Row 2: B–E
    apply_formats!(m, row s + 2, {
        1 => &sec.ft_b_left,
        2 => &_styles.center_center,
        3 => &sec.ft_d_right,
        4 => &sec.ft_e_center,
    });

    // Row 3: B–E (blanks)
    apply_formats!(m, row s + 3, {
        1 => &sec.ft_b_left,
        2 => &_styles.left_center,
        3 => &sec.ft_d_right,
        4 => &sec.ft_e_right,
    });

    // Row 4: balance box (B: formula, C: blank, D: check, E: difference)
    apply_formats!(m, row s + 4, {
        1 => &sec.ft_b_label_box,
        2 => &sec.ft_c_box,
        3 => &sec.ft_d_box,
        4 => &sec.ft_e_number_box,
    });

    // Row 5: B–E (blanks)
    apply_formats!(m, row s + 5, {
        1 => &sec.ft_b_left,
        2 => &_styles.left_center,
        3 => &_styles.left_center,
        4 => &sec.ft_e_right,
    });

    // Row 6: balance reconciliation (B: formula, C,D: blank, E: OK formula)
    apply_formats!(m, row s + 6, {
        1 => &sec.ft_b_left,
        2 => &_styles.left_center,
        3 => &_styles.left_center,
        4 => &sec.ft_e_gray_box,
    });

    // Rows 7–8: bank, cash (B: label, C,D: blank, E: input)
    apply_formats!(m, rows s + 7..=s + 8, {
        1 => &sec.ft_b_input_label_top,
        2 => &sec.ft_c_input_top,
        3 => &sec.ft_d_input_top_right,
        4 => &sec.ft_e_input_top,
    });

    // Row 9: other (B: label, C,D: blank, E: input with bottom border)
    apply_formats!(m, row s + 9, {
        1 => &sec.ft_b_input_label_bottom,
        2 => &sec.ft_c_input_bottom,
        3 => &sec.ft_d_input_bottom_right,
        4 => &sec.ft_e_input_bottom,
    });

    // Rows 13–14: confirmation text (B: formula)
    m.set(s + 13, 1, &_styles.left_center);
    m.set(s + 14, 1, &_styles.left_center);

    // Row 19: signatures (B, D: formulas; C, E–G: blanks with top border)
    apply_formats!(m, row s + 19, {
        1 => &sec.ft_signature,
        2 => &sec.ft_signature_top,
        3 => &sec.ft_signature,
        4 => &sec.ft_signature_top,
        5 => &sec.ft_signature_top,
        6 => &sec.ft_signature_top,
    });

    // Row 20: function title (D: formula)
    m.set(s + 20, 3, &_styles.left_center);
}

/// Extends the [`FormatMatrix`] with pre-body section formats (rows 22–25).
///
/// Must be called after [`build_format_matrix`], before writing the prebody section.
pub fn extend_format_matrix_with_prebody(m: &mut FormatMatrix, sec: &SectionStyles) {
    // Row 22 (Excel 23): column headers (D–H: vertical merge heads) + blanks B/C
    apply_formats!(m, row 22, {
        1 => &sec.pb_lbl_b,    // B23: blank
        2 => &sec.pb_lbl_c,    // C23: blank
        3 => &sec.pb_val,      // D23: merge-head D23:D26
        4 => &sec.pb_val_bold, // E23: merge-head E23:E26 (expenses, bold)
        5 => &sec.pb_val,      // F23: merge-head F23:F26
        6 => &sec.pb_val,      // G23: merge-head G23:G26
        7 => &sec.pb_right,    // H23: merge-head H23:H26 (medium-right)
    });

    // Rows 23–25 (Excel 24–26): B:C merged rows
    apply_formats!(m, row 23, { 1 => &sec.pb_mid_bold }); // B24:C24 (expenses, bold)
    apply_formats!(m, row 24, { 1 => &sec.pb_mid });       // B25:C25 (currency row)
    apply_formats!(m, row 25, { 1 => &sec.pb_bot });       // B26:C26 (thin bottom)
}

// ============================================================================
// Tests
// ============================================================================

#[cfg(test)]
mod format_tests {
    use super::*;
    use crate::report::body::{BodyConfig, BodyLayout, CategoryMode};

    /// Verifies that `extend_format_matrix_with_body()` sets formats for all
    /// expected row/column combinations.
    #[test]
    fn test_body_format_matrix_coverage() {
        let styles = ReportStyles::new();
        // Cat 1: 2 positions → header + 2 position rows + footer
        // Cat 2: 0 positions → header-input (single row)
        let config = BodyConfig::new()
            .with_positions(1, 2)
            .with_positions(2, 0);
        let layout = BodyLayout::compute(&config);
        let sec = SectionStyles::new(&styles);
        let mut fmt = build_format_matrix(&styles, &sec);
        extend_format_matrix_with_body(&mut fmt, &styles, &layout);

        // --- Category 1: WithPositions ---
        let cat1 = &layout.categories[0];
        if let CategoryMode::WithPositions { header_row, positions, footer_row } = &cat1.mode {
            // Header row: B–H all set
            for col in [1u16, 2, 3, 4, 5, 6, 7] {
                assert!(fmt.get(*header_row, col).is_some(),
                    "cat1 header row={header_row}, col={col} missing");
            }
            // Position rows: B–H all set
            for row in positions.start_row..=positions.end_row {
                for col in [1u16, 2, 3, 4, 5, 6, 7] {
                    assert!(fmt.get(row, col).is_some(),
                        "cat1 pos row={row}, col={col} missing");
                }
            }
            // Footer row: B, D–H set (no C: B:C is merged)
            for col in [1u16, 3, 4, 5, 6, 7] {
                assert!(fmt.get(*footer_row, col).is_some(),
                    "cat1 footer row={footer_row}, col={col} missing");
            }
            // Column C in footer row: explicitly NOT set
            assert!(fmt.get(*footer_row, 2).is_none(),
                "cat1 footer col=2 should not be set (B:C merged)");
        } else {
            panic!("cat1 is not WithPositions");
        }

        // --- Category 2: HeaderInput ---
        let cat2 = &layout.categories[1];
        if let CategoryMode::HeaderInput { row } = &cat2.mode {
            // B–H all set
            for col in [1u16, 2, 3, 4, 5, 6, 7] {
                assert!(fmt.get(*row, col).is_some(),
                    "cat2 header-input row={row}, col={col} missing");
            }
        } else {
            panic!("cat2 is not HeaderInput");
        }

        // --- Total row: B, D–H set (no C: B:C is merged) ---
        for col in [1u16, 3, 4, 5, 6, 7] {
            assert!(fmt.get(layout.total_row, col).is_some(),
                "total row={}, col={col} missing", layout.total_row);
        }
        assert!(fmt.get(layout.total_row, 2).is_none(),
            "total col=2 should not be set (B:C merged)");
    }
}
