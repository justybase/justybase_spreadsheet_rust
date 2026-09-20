//! Reader factory (port of `ReaderFactory.ts`).

use crate::error::{SpreadsheetError, SpreadsheetResult};
use crate::formats::CellRef;
use crate::xlsb_reader::XlsbCellReader;
use crate::xlsx_reader::XlsxCellReader;
use crate::{XlsbReader, XlsxReader};
use std::path::Path;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ReaderKind {
    Xlsx,
    Xlsb,
}

/// Detect the reader kind from the file extension.
pub fn detect_kind(path: &Path) -> SpreadsheetResult<ReaderKind> {
    match path
        .extension()
        .and_then(|e| e.to_str())
        .map(|e| e.to_ascii_lowercase())
    {
        Some(e) if e == "xlsx" => Ok(ReaderKind::Xlsx),
        Some(e) if e == "xlsb" => Ok(ReaderKind::Xlsb),
        other => Err(SpreadsheetError::UnsupportedExtension(format!(
            ".{}",
            other.unwrap_or_default()
        ))),
    }
}

/// Unified reader surface (mirrors `ExcelReaderAbstract`).
#[allow(clippy::large_enum_variant)]
pub enum SpreadsheetReader {
    Xlsx(XlsxReader),
    Xlsb(XlsbReader),
}

/// Unified streaming cell reader for XLSX and XLSB workbooks.
pub enum SpreadsheetCellReader<'a> {
    Xlsx(XlsxCellReader<'a>),
    Xlsb(XlsbCellReader<'a>),
}

impl<'a> SpreadsheetCellReader<'a> {
    pub fn next_cell<'b>(&'b mut self) -> SpreadsheetResult<Option<CellRef<'b>>> {
        match self {
            Self::Xlsx(reader) => reader.next_cell(),
            Self::Xlsb(reader) => reader.next_cell(),
        }
    }
}

impl SpreadsheetReader {
    /// Create a reader based on the file extension (`.xlsx` / `.xlsb`).
    pub fn create(path: &Path) -> SpreadsheetResult<Self> {
        match detect_kind(path)? {
            ReaderKind::Xlsx => Ok(SpreadsheetReader::Xlsx(XlsxReader::new())),
            ReaderKind::Xlsb => Ok(SpreadsheetReader::Xlsb(XlsbReader::new())),
        }
    }

    pub fn open(&mut self, path: &Path, read_shared_strings: bool) -> SpreadsheetResult<()> {
        match self {
            SpreadsheetReader::Xlsx(r) => r.open(path, read_shared_strings),
            SpreadsheetReader::Xlsb(r) => r.open(path, read_shared_strings),
        }
    }

    pub fn sheet_names(&self) -> &[String] {
        match self {
            SpreadsheetReader::Xlsx(r) => r.sheet_names(),
            SpreadsheetReader::Xlsb(r) => r.sheet_names(),
        }
    }

    /// Select a worksheet by name and reset row iteration.
    pub fn select_sheet(&mut self, name: &str) -> SpreadsheetResult<()> {
        match self {
            SpreadsheetReader::Xlsx(r) => r.select_sheet(name),
            SpreadsheetReader::Xlsb(r) => r.select_sheet(name),
        }
    }

    pub fn read(&mut self) -> SpreadsheetResult<bool> {
        match self {
            SpreadsheetReader::Xlsx(r) => r.read(),
            SpreadsheetReader::Xlsb(r) => r.read(),
        }
    }

    pub fn get_value(&self, i: usize) -> crate::formats::CellValue {
        match self {
            SpreadsheetReader::Xlsx(r) => r.get_value(i),
            SpreadsheetReader::Xlsb(r) => r.get_value(i),
        }
    }

    pub fn current_row(&self) -> &[crate::formats::CellValue] {
        match self {
            SpreadsheetReader::Xlsx(r) => r.current_row(),
            SpreadsheetReader::Xlsb(r) => r.current_row(),
        }
    }

    /// Create a low-allocation streaming reader for explicit worksheet cells.
    pub fn cell_reader<'a>(
        &'a mut self,
        name: &str,
    ) -> SpreadsheetResult<SpreadsheetCellReader<'a>> {
        match self {
            SpreadsheetReader::Xlsx(reader) => {
                Ok(SpreadsheetCellReader::Xlsx(reader.cell_reader(name)?))
            }
            SpreadsheetReader::Xlsb(reader) => {
                Ok(SpreadsheetCellReader::Xlsb(reader.cell_reader(name)?))
            }
        }
    }
}

/// Convenience alias matching the TS `ReaderFactory.create` name.
pub fn create_reader(path: &Path) -> SpreadsheetResult<SpreadsheetReader> {
    SpreadsheetReader::create(path)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn detects_by_extension() {
        assert_eq!(detect_kind(Path::new("a.xlsx")).unwrap(), ReaderKind::Xlsx);
        assert_eq!(detect_kind(Path::new("a.XLSB")).unwrap(), ReaderKind::Xlsb);
        assert!(detect_kind(Path::new("a.csv")).is_err());
    }
}
