//! High-performance Excel XLSB/XLSX reader and writer.
//!
//! Rust port of `@justybase/spreadsheet-tasks` (TypeScript).
//! XLSB is the preferred format for large datasets — typically faster
//! and smaller than XLSX.
//!
//! # Quick start
//! ```no_run
//! use spreadsheet::{XlsxWriter, CellValue};
//! use std::path::Path;
//!
//! let mut w = XlsxWriter::create(Path::new("output.xlsx")).unwrap();
//! w.add_sheet("Sheet1", false);
//! w.write_sheet(
//!     vec![vec![CellValue::Text("Alice".into()), CellValue::Integer(30)]],
//!     Some(&["Name".to_string(), "Age".to_string()]),
//!     true,
//! )
//! .unwrap();
//! w.finalize().unwrap();
//! ```

mod xlsb_templates;

pub mod atomic_file;
pub mod biff12;
pub mod biff_reader;
pub mod big_buffer;
pub mod error;
pub mod formats;
pub mod reader_factory;
pub mod streaming_state;
pub mod writer_helpers;
pub mod xlsb_reader;
pub mod xlsb_updater;
pub mod xlsb_writer;
pub mod xlsx_reader;
pub mod xlsx_updater;
pub mod xlsx_writer;
pub mod xml_utils;

mod zip_store;

pub use atomic_file::{install_temporary_output, write_buffer_atomically, AtomicFileOperations};
pub use biff12::{
    build_record, read_record, read_utf16, read_vlq, try_read_utf16, uses_1904_date_system_bin,
    vlq_bytes, vlq_length, Biff12Record,
};
pub use big_buffer::BigBuffer;
pub use error::{SpreadsheetError, SpreadsheetResult};
pub use formats::{
    get_format, is_formatted_cell, unwrap_cell, CellRef, CellValue, CellValueRef, FormattedCell,
    PrimitiveCellValue, F,
};
pub use reader_factory::{create_reader, ReaderKind, SpreadsheetCellReader, SpreadsheetReader};
pub use streaming_state::StreamingSheetState;
pub use writer_helpers::{
    apply_header_widths, default_col_width, deterministic_zip_timestamp, init_col_widths,
    sanitize_sheet_name, update_col_widths_from_rows, INVALID_SHEET_NAME_CHARS,
};
pub use xlsb_reader::{XlsbCellReader, XlsbReader};
pub use xlsb_updater::{ReplaceSheetDataOptions, StyleFallback, XlsbUpdater};
pub use xlsb_writer::{XlsbSheetOptions, XlsbWriter};
pub use xlsx_reader::{XlsxCellReader, XlsxReader};
pub use xlsx_updater::{XlsmUpdater, XlsxUpdater};
pub use xlsx_writer::{SheetOptions, XlsxWriter};
pub use xml_utils::{
    column_index_to_letter, column_letter_to_index, escape_xml_text, parse_shared_strings_xml,
    trim_trailing_empty_rows, try_column_letter_to_index, unescape_xml, uses_1904_date_system,
};

/// OLE Automation epoch (1899-12-30 UTC), shared by readers and writers.
pub fn oa_epoch_naive() -> chrono::NaiveDateTime {
    chrono::NaiveDate::from_ymd_opt(1899, 12, 30)
        .expect("valid OA epoch")
        .and_hms_opt(0, 0, 0)
        .expect("valid OA epoch time")
}

/// Convert an OLE Automation date to a naive UTC datetime.
pub fn datetime_from_oa_date(oa: f64) -> chrono::NaiveDateTime {
    oa_epoch_naive() + chrono::Duration::milliseconds((oa * 86_400_000.0) as i64)
}

/// Convert a naive UTC datetime to an OLE Automation date.
pub fn datetime_to_oa_date(dt: &chrono::NaiveDateTime) -> f64 {
    (*dt - oa_epoch_naive()).num_milliseconds() as f64 / 86_400_000.0
}

/// Number of days between the 1900 and 1904 Excel date systems.
pub const EXCEL_1904_DATE_OFFSET_DAYS: f64 = 1462.0;

/// Maximum worksheet dimensions supported by the XLSX/XLSB formats.
pub const EXCEL_MAX_COLUMNS: usize = 16_384;
pub const EXCEL_MAX_ROWS: usize = 1_048_576;

/// Convert an Excel date serial to a datetime.
pub fn datetime_from_excel_serial(
    serial: f64,
    uses_1904_date_system: bool,
) -> chrono::NaiveDateTime {
    let oa = if uses_1904_date_system {
        serial + EXCEL_1904_DATE_OFFSET_DAYS
    } else {
        serial
    };
    datetime_from_oa_date(oa)
}

/// Convert a datetime to the serial used by an Excel workbook.
pub fn datetime_to_excel_serial(dt: &chrono::NaiveDateTime, uses_1904_date_system: bool) -> f64 {
    let oa = datetime_to_oa_date(dt);
    if uses_1904_date_system {
        oa - EXCEL_1904_DATE_OFFSET_DAYS
    } else {
        oa
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn excel_1904_serials_round_trip() {
        let dt = chrono::NaiveDate::from_ymd_opt(2025, 2, 3)
            .unwrap()
            .and_hms_opt(12, 30, 0)
            .unwrap();
        let serial_1900 = datetime_to_excel_serial(&dt, false);
        let serial_1904 = datetime_to_excel_serial(&dt, true);
        assert!((serial_1900 - serial_1904 - EXCEL_1904_DATE_OFFSET_DAYS).abs() < 1e-9);
        assert_eq!(datetime_from_excel_serial(serial_1900, false), dt);
        assert_eq!(datetime_from_excel_serial(serial_1904, true), dt);
    }
}
