//! Sheet-name sanitization and autofit column widths (port of `writerHelpers.ts`).

use crate::formats::{get_format, CellValue};

/// Excel sheet-name characters that must be replaced.
pub const INVALID_SHEET_NAME_CHARS: &[char] = &['\\', '/', '*', '?', '[', ']', ':'];

/// Sanitize an Excel sheet name (invalid chars, max 31 chars, non-empty).
///
/// `sheet_count` is the current sheet count before adding this sheet
/// (used for default names).
pub fn sanitize_sheet_name(name: &str, sheet_count: usize) -> String {
    if name.is_empty() {
        return format!("Sheet{}", sheet_count + 1);
    }
    let mut sanitized: String = name
        .chars()
        .map(|c| {
            if INVALID_SHEET_NAME_CHARS.contains(&c) {
                '_'
            } else {
                c
            }
        })
        .collect();
    if sanitized.chars().count() > 31 {
        sanitized = sanitized.chars().take(31).collect();
    }
    if sanitized.trim().is_empty() {
        sanitized = format!("Sheet{}", sheet_count + 1);
    }
    sanitized
}

/// Return a sanitized, case-insensitively unique worksheet name.
pub fn unique_sheet_name(
    name: &str,
    sheet_count: usize,
    is_taken: impl Fn(&str) -> bool,
) -> String {
    let base = sanitize_sheet_name(name, sheet_count);
    if !is_taken(&base) {
        return base;
    }
    for suffix in 2.. {
        let suffix_text = format!(" ({suffix})");
        let keep = 31usize.saturating_sub(suffix_text.chars().count());
        let prefix: String = base.chars().take(keep).collect();
        let candidate = format!("{prefix}{suffix_text}");
        if !is_taken(&candidate) {
            return candidate;
        }
    }
    unreachable!("worksheet name suffix search cannot exhaust usize")
}

/// Create column-width scratch array filled with -1 (unknown).
pub fn init_col_widths(column_count: usize) -> Vec<f64> {
    vec![-1.0; column_count]
}

/// Apply autofit widths from header labels.
pub fn apply_header_widths(col_widths: &mut [f64], headers: &[String], column_count: usize) {
    for (i, width_slot) in col_widths.iter_mut().enumerate().take(column_count) {
        let len = headers.get(i).map(|h| h.chars().count() + 1).unwrap_or(0) as f64;
        let mut width = 1.25 * len + 2.0;
        if width > 80.0 {
            width = 80.0;
        }
        if *width_slot < width {
            *width_slot = width;
        }
    }
}

fn display_len_of(value: &CellValue) -> usize {
    if get_format(value).is_some() {
        // formatted values are measured by their primitive payload below
    }
    match value {
        CellValue::Empty => 0,
        CellValue::DateTime(_) => 20,
        CellValue::Formatted(f) => match &f.value {
            crate::formats::PrimitiveCellValue::DateTime(_) => 20,
            crate::formats::PrimitiveCellValue::Empty => 0,
            crate::formats::PrimitiveCellValue::Text(s) => s.chars().count() + 1,
            crate::formats::PrimitiveCellValue::Number(n) => n.to_string().len() + 1,
            crate::formats::PrimitiveCellValue::Integer(n) => n.to_string().len() + 1,
            crate::formats::PrimitiveCellValue::Boolean(b) => b.to_string().len() + 1,
        },
        CellValue::Text(s) => s.chars().count() + 1,
        CellValue::Number(n) => n.to_string().len() + 1,
        CellValue::Integer(n) => n.to_string().len() + 1,
        CellValue::Boolean(b) => b.to_string().len() + 1,
    }
}

/// Sample up to `sample_limit` rows and expand column widths (autofit).
pub fn update_col_widths_from_rows(
    col_widths: &mut [f64],
    rows: &[Vec<CellValue>],
    sample_limit: usize,
) {
    for row in rows.iter().take(sample_limit) {
        for (c, cell) in row.iter().enumerate() {
            if c >= col_widths.len() {
                break;
            }
            if matches!(cell, CellValue::Empty) {
                continue;
            }
            let len = display_len_of(cell) as f64;
            let mut width = 1.25 * len + 2.0;
            if width > 80.0 {
                width = 80.0;
            }
            if col_widths[c] < width {
                col_widths[c] = width;
            }
        }
    }
}

/// Fixed ZIP entry timestamp (DOS epoch 1980-01-01 00:00:00).
/// `zip::write::SimpleFileOptions::default()` uses wall-clock time via
/// `DateTime::default_for_write()`, which leaks build/run time into every
/// generated XLSX/XLSB. Use this instead for reproducible output.
pub fn deterministic_zip_timestamp() -> zip::DateTime {
    zip::DateTime::from_date_and_time(1980, 1, 1, 0, 0, 0).expect("valid fixed ZIP timestamp")
}

/// Resolve a stored autofit width to the value written into the file.
/// XLSB uses floored integers; XLSX keeps the float.
pub fn default_col_width(width: f64, floor: bool) -> f64 {
    if width > 0.0 {
        if floor {
            width.floor()
        } else {
            width
        }
    } else {
        10.0
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn sanitize_rules() {
        assert_eq!(sanitize_sheet_name("", 0), "Sheet1");
        assert_eq!(sanitize_sheet_name("a/b:c", 0), "a_b_c");
        assert_eq!(sanitize_sheet_name("   ", 2), "Sheet3");
        let long = "x".repeat(40);
        assert_eq!(sanitize_sheet_name(&long, 0).len(), 31);
    }

    #[test]
    fn unique_names_are_case_insensitive_and_fit_limit() {
        let existing = ["Data".to_string(), "Data (2)".to_string()];
        let name = unique_sheet_name("data", 2, |candidate| {
            existing
                .iter()
                .any(|value| value.eq_ignore_ascii_case(candidate))
        });
        assert_eq!(name, "data (3)");

        let existing = ["x".repeat(31)];
        let name = unique_sheet_name(&existing[0], 1, |candidate| candidate == existing[0]);
        assert_eq!(name.chars().count(), 31);
        assert!(name.ends_with(" (2)"));
    }

    #[test]
    fn widths_autofit() {
        let mut w = init_col_widths(2);
        apply_header_widths(&mut w, &["ID".to_string(), "Name".to_string()], 2);
        assert!(w[0] > 0.0 && w[1] > w[0]);
        assert_eq!(default_col_width(-1.0, true), 10.0);
        assert_eq!(default_col_width(12.9, true), 12.0);
        assert_eq!(default_col_width(12.9, false), 12.9);
    }
}
