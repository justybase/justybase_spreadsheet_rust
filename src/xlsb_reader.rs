//! XLSB reader (port of `XlsbReader.ts`).
//!
//! Reads the first worksheet by default and supports selecting any worksheet
//! by name without reopening the package.

use crate::biff12::uses_1904_date_system_bin;
use crate::biff_reader::BiffReaderWriter;
use crate::error::{SpreadsheetError, SpreadsheetResult};
use crate::formats::{CellRef, CellValue, CellValueRef};
use crate::{datetime_from_excel_serial, datetime_from_oa_date, oa_epoch_naive};
use std::collections::{HashMap, HashSet};
use std::fs::File;
use std::io::{BufReader, Read};
use std::path::{Path, PathBuf};
use zip::read::ZipFile;

type XlsbStreamParser<'a> = BiffReaderWriter<'a, BufReader<ZipFile<'a, File>>>;

#[ouroboros::self_referencing]
struct XlsbSheetStream {
    archive: zip::ZipArchive<File>,
    #[borrows(mut archive)]
    #[covariant]
    reader: XlsbStreamParser<'this>,
}

struct SheetInfo {
    name: String,
    r_id: String,
    path: Option<String>,
}

pub struct XlsbReader {
    zip: Option<zip::ZipArchive<File>>,
    source_path: Option<PathBuf>,
    shared_strings: Vec<String>,
    sheet_names: Vec<String>,
    sheets: Vec<SheetInfo>,

    current_sheet_index: i64,
    reader: Option<XlsbSheetStream>,
    current_row: Vec<CellValue>,
    pending_row_index: i32,
    eof: bool,

    pub field_count: usize,
    pub row_count: usize,
    pub actual_sheet_name: String,
    pub results_count: usize,

    pub xf_id_to_num_fmt_id: Vec<u16>,
    pub custom_date_formats: HashSet<u16>,
    /// Whether the workbook uses Excel's Macintosh 1904 date system.
    pub uses_1904_date_system: bool,
}

impl XlsbReader {
    pub fn new() -> Self {
        Self {
            zip: None,
            source_path: None,
            shared_strings: Vec::new(),
            sheet_names: Vec::new(),
            sheets: Vec::new(),
            current_sheet_index: -1,
            reader: None,
            current_row: Vec::new(),
            pending_row_index: -1,
            eof: false,
            field_count: 0,
            row_count: 0,
            actual_sheet_name: String::new(),
            results_count: 0,
            xf_id_to_num_fmt_id: Vec::new(),
            custom_date_formats: HashSet::new(),
            uses_1904_date_system: false,
        }
    }

    pub fn open(&mut self, path: &Path, read_shared_strings: bool) -> SpreadsheetResult<()> {
        self.reader = None;
        self.zip = None;
        self.source_path = None;
        self.shared_strings.clear();
        self.sheet_names.clear();
        self.sheets.clear();
        self.current_row.clear();
        self.pending_row_index = -1;
        self.eof = false;
        self.field_count = 0;
        self.row_count = 0;
        self.actual_sheet_name.clear();
        self.results_count = 0;
        self.xf_id_to_num_fmt_id.clear();
        self.custom_date_formats.clear();
        self.uses_1904_date_system = false;
        self.current_sheet_index = -1;

        let file = File::open(path)?;
        let mut zip = zip::ZipArchive::new(file)?;

        if let Some(wb) = read_zip_entry(&mut zip, "xl/workbook.bin")? {
            self.uses_1904_date_system = uses_1904_date_system_bin(&wb)?;
            let mut reader = BiffReaderWriter::new(&wb);
            // borrow issue: reader borrows wb (local). Restructure: leak via
            // storing bytes in self? Simplest: parse with owned copy here.
            while reader.read_workbook()? {
                if reader.is_sheet {
                    let name = reader.workbook_name.clone().unwrap_or_default();
                    let r_id = reader.rec_id.clone().unwrap_or_default();
                    self.sheet_names.push(name.clone());
                    self.sheets.push(SheetInfo {
                        name,
                        r_id,
                        path: None,
                    });
                }
            }
        }

        if let Some(rels) = read_zip_entry(&mut zip, "xl/_rels/workbook.bin.rels")? {
            let xml = String::from_utf8_lossy(&rels).into_owned();
            let r_id_to_target = parse_relationships(&xml);
            for sheet in &mut self.sheets {
                if let Some(mut target) = r_id_to_target.get(&sheet.r_id).cloned() {
                    if let Some(stripped) = target.strip_prefix('/') {
                        target = stripped.to_owned();
                    }
                    if !target.starts_with("xl/") {
                        target = format!("xl/{target}");
                    }
                    sheet.path = Some(target);
                }
            }
        }

        if read_shared_strings {
            match zip.by_name("xl/sharedStrings.bin") {
                Ok(entry) => {
                    let mut reader =
                        BiffReaderWriter::from_reader(BufReader::with_capacity(64 * 1024, entry));
                    while reader.read_shared_strings()? {
                        if let Some(value) = reader.shared_string_value.take() {
                            self.shared_strings.push(value);
                        }
                    }
                }
                Err(zip::result::ZipError::FileNotFound) => {}
                Err(error) => return Err(error.into()),
            }
        }

        match zip.by_name("xl/styles.bin") {
            Ok(entry) => {
                let mut reader =
                    BiffReaderWriter::from_reader(BufReader::with_capacity(16 * 1024, entry));
                while reader.read_styles()? {}
                self.xf_id_to_num_fmt_id = reader.xf_index_to_num_fmt_id;
                self.custom_date_formats = reader.custom_num_fmts;
            }
            Err(zip::result::ZipError::FileNotFound) => {}
            Err(error) => return Err(error.into()),
        }

        self.source_path = Some(path.canonicalize()?);
        self.zip = Some(zip);
        self.results_count = self.sheets.len();
        self.current_sheet_index = -1;
        Ok(())
    }

    pub fn sheet_names(&self) -> &[String] {
        &self.sheet_names
    }

    /// Create a streaming cell reader for a worksheet.
    pub fn cell_reader<'a>(&'a mut self, name: &str) -> SpreadsheetResult<XlsbCellReader<'a>> {
        let path = self
            .sheets
            .iter()
            .find(|sheet| sheet.name == name)
            .and_then(|sheet| sheet.path.clone())
            .ok_or_else(|| SpreadsheetError::SheetNotFound(name.to_owned()))?;
        let entry = self
            .zip
            .as_mut()
            .ok_or_else(|| SpreadsheetError::InvalidFormat("reader is not open".into()))?
            .by_name(&path)
            .map_err(|error| match error {
                zip::result::ZipError::FileNotFound => {
                    SpreadsheetError::InvalidFormat(format!("missing sheet part: {path}"))
                }
                error => error.into(),
            })?;
        XlsbCellReader::new(
            entry,
            &self.shared_strings,
            &self.xf_id_to_num_fmt_id,
            &self.custom_date_formats,
            self.uses_1904_date_system,
        )
    }

    /// Select a worksheet by its workbook name and reset row iteration.
    ///
    /// Calling [`XlsbReader::read`] without selecting a sheet retains the
    /// historical behavior of reading the first worksheet.
    pub fn select_sheet(&mut self, name: &str) -> SpreadsheetResult<()> {
        let index = self
            .sheets
            .iter()
            .position(|sheet| sheet.name == name)
            .ok_or_else(|| SpreadsheetError::SheetNotFound(name.to_owned()))?;
        self.select_sheet_index(index).map(|_| ())
    }

    /// Advance to the next row. Returns true when a row is available.
    pub fn read(&mut self) -> SpreadsheetResult<bool> {
        if self.current_sheet_index == -1 {
            if self.sheets.is_empty() {
                return Ok(false);
            }
            if !self.select_sheet_index(0)? {
                return Ok(false);
            }
        }
        if self.eof {
            return Ok(false);
        }
        self.current_row.clear();
        loop {
            let (
                has_record,
                row_index,
                read_cell,
                cell_type,
                column_num,
                xf_index,
                int_value,
                double_val,
                bool_value,
                string_value,
            ) = match self.reader.as_mut() {
                Some(stream) => stream.with_reader_mut(|reader| {
                    let has = reader.read_worksheet()?;
                    Ok::<_, SpreadsheetError>((
                        has,
                        reader.row_index,
                        reader.read_cell,
                        reader.cell_type,
                        reader.column_num,
                        reader.xf_index,
                        reader.int_value,
                        reader.double_val,
                        reader.bool_value,
                        reader.string_value.take(),
                    ))
                })?,
                None => return Ok(false),
            };
            if !has_record {
                self.eof = true;
                self.row_count += 1;
                return Ok(true);
            }
            if row_index != -1 && row_index != self.pending_row_index {
                if row_index < 0 || row_index as usize >= crate::EXCEL_MAX_ROWS {
                    return Err(SpreadsheetError::InvalidFormat(
                        "XLSB row index is outside Excel worksheet bounds".into(),
                    ));
                }
                self.pending_row_index = row_index;
                self.row_count += 1;
                return Ok(true);
            }
            if read_cell {
                if column_num < 0 || column_num as usize >= crate::EXCEL_MAX_COLUMNS {
                    return Err(SpreadsheetError::InvalidFormat(
                        "XLSB column index is outside Excel worksheet bounds".into(),
                    ));
                }
                let col = column_num as usize;
                let val: CellValue = match cell_type {
                    2 => self
                        .shared_strings
                        .get(int_value as usize)
                        .map(|s| CellValue::Text(s.clone()))
                        .unwrap_or(CellValue::Empty),
                    3 => {
                        let mut v = CellValue::Number(double_val);
                        let num_fmt_id =
                            self.xf_id_to_num_fmt_id.get(xf_index).copied().unwrap_or(0);
                        if is_date_num_fmt(num_fmt_id, &self.custom_date_formats) {
                            v = CellValue::DateTime(datetime_from_excel_serial(
                                double_val,
                                self.uses_1904_date_system,
                            ));
                        }
                        v
                    }
                    4 => CellValue::Boolean(bool_value),
                    5 => string_value
                        .map(CellValue::Text)
                        .unwrap_or(CellValue::Empty),
                    _ => CellValue::Empty,
                };
                if self.current_row.len() <= col {
                    self.current_row.resize(col + 1, CellValue::Empty);
                }
                self.current_row[col] = val;
                if col + 1 > self.field_count {
                    self.field_count = col + 1;
                }
            }
        }
    }

    fn init_sheet(&mut self, index: usize) -> SpreadsheetResult<bool> {
        if index >= self.sheets.len() {
            return Ok(false);
        }
        self.reader = None;
        let path = self.sheets[index]
            .path
            .clone()
            .ok_or_else(|| SpreadsheetError::InvalidFormat("sheet has no resolved path".into()))?;
        let source_path = self
            .source_path
            .as_ref()
            .ok_or_else(|| SpreadsheetError::InvalidFormat("reader is not open".into()))?;
        let archive = zip::ZipArchive::new(File::open(source_path)?)?;
        let sheet_path = path.clone();
        let mut reader = XlsbSheetStreamTryBuilder {
            archive,
            reader_builder: |archive| {
                let entry = archive.by_name(&sheet_path).map_err(|error| match error {
                    zip::result::ZipError::FileNotFound => {
                        SpreadsheetError::InvalidFormat(format!("missing sheet part: {sheet_path}"))
                    }
                    error => error.into(),
                })?;
                Ok::<_, SpreadsheetError>(BiffReaderWriter::from_reader(BufReader::with_capacity(
                    64 * 1024,
                    entry,
                )))
            },
        }
        .try_build()?;
        self.actual_sheet_name = self.sheets[index].name.clone();
        self.eof = false;
        self.pending_row_index = -1;
        let available = reader.with_reader_mut(|parser| {
            while parser.read_worksheet()? {
                if parser.row_index != -1 {
                    self.pending_row_index = parser.row_index;
                    return Ok::<_, SpreadsheetError>(true);
                }
            }
            Ok(false)
        })?;
        self.eof = !available;
        self.reader = Some(reader);
        Ok(available)
    }

    fn select_sheet_index(&mut self, index: usize) -> SpreadsheetResult<bool> {
        let available = self.init_sheet(index)?;
        self.current_sheet_index = index as i64;
        self.current_row.clear();
        self.field_count = 0;
        self.row_count = 0;
        Ok(available)
    }

    pub fn get_value(&self, i: usize) -> CellValue {
        self.current_row.get(i).cloned().unwrap_or(CellValue::Empty)
    }

    pub fn current_row(&self) -> &[CellValue] {
        &self.current_row
    }

    pub fn close(&mut self) {
        self.reader = None;
        self.current_sheet_index = -1;
    }

    pub fn datetime_from_oa(&self, oa: f64) -> chrono::NaiveDateTime {
        let _ = oa_epoch_naive();
        datetime_from_oa_date(oa)
    }
}

impl Default for XlsbReader {
    fn default() -> Self {
        Self::new()
    }
}

fn is_date_num_fmt(id: u16, custom: &HashSet<u16>) -> bool {
    (14..=22).contains(&id) || (45..=47).contains(&id) || custom.contains(&id)
}

fn read_zip_entry(
    zip: &mut zip::ZipArchive<File>,
    name: &str,
) -> SpreadsheetResult<Option<Vec<u8>>> {
    match zip.by_name(name) {
        Ok(mut entry) => {
            let size = usize::try_from(entry.size()).map_err(|_| {
                SpreadsheetError::InvalidFormat("ZIP entry is too large for this platform".into())
            })?;
            let mut bytes = Vec::with_capacity(size);
            entry.read_to_end(&mut bytes)?;
            Ok(Some(bytes))
        }
        Err(zip::result::ZipError::FileNotFound) => Ok(None),
        Err(error) => Err(error.into()),
    }
}

/// Streaming XLSB cell reader backed by the BIFF12 worksheet decoder.
pub struct XlsbCellReader<'a> {
    reader: XlsbStreamParser<'a>,
    shared_strings: &'a [String],
    formats: &'a [u16],
    custom_date_formats: &'a HashSet<u16>,
    uses_1904_date_system: bool,
    scratch: String,
    row: u32,
    active_row: Option<i32>,
}

impl<'a> XlsbCellReader<'a> {
    fn new(
        entry: ZipFile<'a, File>,
        shared_strings: &'a [String],
        formats: &'a [u16],
        custom_date_formats: &'a HashSet<u16>,
        uses_1904_date_system: bool,
    ) -> SpreadsheetResult<Self> {
        Ok(Self {
            reader: BiffReaderWriter::from_reader(BufReader::with_capacity(64 * 1024, entry)),
            shared_strings,
            formats,
            custom_date_formats,
            uses_1904_date_system,
            scratch: String::with_capacity(128),
            row: 0,
            active_row: None,
        })
    }

    /// Return the next explicit worksheet cell in BIFF12 document order.
    pub fn next_cell<'b>(&'b mut self) -> SpreadsheetResult<Option<CellRef<'b>>> {
        loop {
            let record = {
                let reader = &mut self.reader;
                let has_record = reader.read_worksheet()?;
                (
                    has_record,
                    reader.row_index,
                    reader.read_cell,
                    reader.cell_type,
                    reader.column_num,
                    reader.xf_index,
                    reader.int_value,
                    reader.double_val,
                    reader.bool_value,
                    reader.string_value.take(),
                )
            };
            let (
                has_record,
                row_index,
                read_cell,
                cell_type,
                column_num,
                xf_index,
                int_value,
                double_val,
                bool_value,
                string_value,
            ) = record;
            if !has_record {
                return Ok(None);
            }
            if row_index >= 0 && self.active_row != Some(row_index) {
                if row_index as usize >= crate::EXCEL_MAX_ROWS {
                    return Err(SpreadsheetError::InvalidFormat(
                        "XLSB row index is outside Excel worksheet bounds".into(),
                    ));
                }
                self.active_row = Some(row_index);
                self.row = row_index as u32;
                continue;
            }
            if !read_cell {
                continue;
            }

            if column_num < 0 || column_num as usize >= crate::EXCEL_MAX_COLUMNS {
                return Err(SpreadsheetError::InvalidFormat(
                    "XLSB column index is outside Excel worksheet bounds".into(),
                ));
            }
            let column = column_num as u32;
            let value = match cell_type {
                2 => self
                    .shared_strings
                    .get(int_value as usize)
                    .map(String::as_str)
                    .map(CellValueRef::Text)
                    .unwrap_or(CellValueRef::Empty),
                3 => {
                    let num_fmt_id = self.formats.get(xf_index).copied().unwrap_or(0);
                    if is_date_num_fmt(num_fmt_id, self.custom_date_formats) {
                        CellValueRef::DateTime(datetime_from_excel_serial(
                            double_val,
                            self.uses_1904_date_system,
                        ))
                    } else {
                        CellValueRef::Number(double_val)
                    }
                }
                4 => CellValueRef::Boolean(bool_value),
                5 => {
                    self.scratch = string_value.unwrap_or_default();
                    CellValueRef::Text(self.scratch.as_str())
                }
                _ => CellValueRef::Empty,
            };
            return Ok(Some(CellRef {
                row: self.row,
                column,
                value,
            }));
        }
    }
}

pub(crate) fn parse_relationships(xml: &str) -> HashMap<String, String> {
    let mut map = HashMap::new();
    let mut search = 0;
    while let Some((start, end)) = crate::xml_utils::find_open_tag(xml, search, "Relationship") {
        let tag = &xml[start..end];
        if let (Some(id), Some(target)) = (attr_value(tag, "Id"), attr_value(tag, "Target")) {
            map.insert(id, target);
        }
        search = end;
        if search >= xml.len() {
            break;
        }
    }
    map
}

fn attr_value(tag: &str, name: &str) -> Option<String> {
    crate::xml_utils::tag_attribute(tag, name)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn rels_parsing_order_independent() {
        let xml = r#"<Relationships><Relationship Target="worksheets/sheet1.bin" Id="rId1" Type="x"/></Relationships>"#;
        let m = parse_relationships(xml);
        assert_eq!(m.get("rId1").unwrap(), "worksheets/sheet1.bin");
    }
}
