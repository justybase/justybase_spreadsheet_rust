//! XLSX reader (port of `XlsxReader.ts`).
//!
//! Worksheet XML is consumed with a reusable streaming parser, mirroring the
//! TS `<row>` / `<c r= t= s=>` / `<v>` parser without retaining the sheet.

use crate::error::{SpreadsheetError, SpreadsheetResult};
use crate::formats::{is_date_format_code, CellValue, CellValueRef};
use crate::xlsb_reader::parse_relationships;
use crate::{datetime_from_excel_serial, xml_utils};
use quick_xml::events::{BytesStart, Event};
use quick_xml::Reader as XmlReader;
use std::collections::HashSet;
use std::fs::File;
use std::io::{BufRead, BufReader, Cursor, Read};
use std::path::Path;
use zip::read::ZipFile;

struct SheetInfo {
    name: String,
    #[allow(dead_code)]
    sheet_id: String,
    #[allow(dead_code)]
    r_id: String,
    path: String,
}

pub struct XlsxReader {
    zip: Option<zip::ZipArchive<File>>,
    shared_strings: Vec<String>,
    sheet_names: Vec<String>,
    sheets: Vec<SheetInfo>,

    current_sheet_index: i64,
    row_reader: Option<XmlReader<BufReader<Cursor<Vec<u8>>>>>,
    row_buf: Vec<u8>,
    current_row: Vec<CellValue>,

    pub field_count: usize,
    pub row_count: usize,
    pub actual_sheet_name: String,
    pub results_count: usize,

    pub xf_id_to_num_fmt_id: Vec<u16>,
    pub custom_date_formats: HashSet<u16>,
    /// Whether the workbook uses Excel's Macintosh 1904 date system.
    pub uses_1904_date_system: bool,
}

impl XlsxReader {
    pub fn new() -> Self {
        Self {
            zip: None,
            shared_strings: Vec::new(),
            sheet_names: Vec::new(),
            sheets: Vec::new(),
            current_sheet_index: -1,
            row_reader: None,
            row_buf: Vec::with_capacity(1024),
            current_row: Vec::new(),
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
        self.zip = None;
        self.shared_strings.clear();
        self.sheet_names.clear();
        self.sheets.clear();
        self.current_row.clear();
        self.row_reader = None;
        self.row_buf.clear();
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
        let rels = read_zip_entry(&mut zip, "xl/_rels/workbook.xml.rels")?;
        let r_id_to_target = rels
            .as_deref()
            .map(|b| parse_relationships(&String::from_utf8_lossy(b)))
            .unwrap_or_default();

        if let Some(wb) = read_zip_entry(&mut zip, "xl/workbook.xml")? {
            let xml = String::from_utf8_lossy(&wb).into_owned();
            self.uses_1904_date_system = xml_utils::uses_1904_date_system(&xml);
            for (name, sheet_id, r_id) in parse_sheet_entries(&xml) {
                let mut full_path = r_id_to_target.get(&r_id).cloned().unwrap_or_default();
                if let Some(stripped) = full_path.strip_prefix('/') {
                    full_path = stripped.to_owned();
                }
                if !full_path.starts_with("xl/") {
                    full_path = format!("xl/{full_path}");
                }
                self.sheet_names.push(name.clone());
                self.sheets.push(SheetInfo {
                    name,
                    sheet_id,
                    r_id,
                    path: full_path,
                });
            }
        }

        if read_shared_strings {
            if let Some(ss) = read_zip_entry(&mut zip, "xl/sharedStrings.xml")? {
                let xml = String::from_utf8_lossy(&ss).into_owned();
                self.shared_strings = xml_utils::parse_shared_strings_xml(&xml);
            }
        }

        if let Some(styles) = read_zip_entry(&mut zip, "xl/styles.xml")? {
            let xml = String::from_utf8_lossy(&styles).into_owned();
            let (xfs, customs) = parse_styles(&xml);
            self.xf_id_to_num_fmt_id = xfs;
            self.custom_date_formats = customs;
        }

        self.zip = Some(zip);
        self.results_count = self.sheets.len();
        self.current_sheet_index = -1;
        Ok(())
    }

    pub fn sheet_names(&self) -> &[String] {
        &self.sheet_names
    }

    /// Create a zero-copy cell reader for a worksheet.
    ///
    /// Text values borrow the shared-string table or the reader's reusable
    /// scratch buffer and remain valid until the next call to `next_cell`.
    pub fn cell_reader<'a>(&'a mut self, name: &str) -> SpreadsheetResult<XlsxCellReader<'a>> {
        let path = self
            .sheets
            .iter()
            .find(|sheet| sheet.name == name)
            .map(|sheet| sheet.path.clone())
            .ok_or_else(|| SpreadsheetError::SheetNotFound(name.to_owned()))?;
        let shared_strings = &self.shared_strings;
        let formats = &self.xf_id_to_num_fmt_id;
        let custom_date_formats = &self.custom_date_formats;
        let uses_1904_date_system = self.uses_1904_date_system;
        let zip = self
            .zip
            .as_mut()
            .ok_or_else(|| SpreadsheetError::InvalidFormat("reader is not open".into()))?;
        let entry = zip.by_name(&path)?;
        XlsxCellReader::new(
            entry,
            shared_strings,
            formats,
            custom_date_formats,
            uses_1904_date_system,
        )
    }

    /// Select a worksheet by its workbook name and reset row iteration.
    ///
    /// Calling [`XlsxReader::read`] without selecting a sheet retains the
    /// historical behavior of reading the first worksheet.
    pub fn select_sheet(&mut self, name: &str) -> SpreadsheetResult<()> {
        let index = self
            .sheets
            .iter()
            .position(|sheet| sheet.name == name)
            .ok_or_else(|| SpreadsheetError::SheetNotFound(name.to_owned()))?;
        self.select_sheet_index(index)
    }

    pub fn read(&mut self) -> SpreadsheetResult<bool> {
        if self.current_sheet_index == -1 {
            if self.sheets.is_empty() {
                return Ok(false);
            }
            self.select_sheet_index(0)?;
        }
        if self.read_next_row()? {
            self.row_count += 1;
            Ok(true)
        } else {
            Ok(false)
        }
    }

    fn init_sheet(&mut self, index: usize) -> SpreadsheetResult<()> {
        if index >= self.sheets.len() {
            return Err(SpreadsheetError::InvalidFormat(
                "sheet index out of range".into(),
            ));
        }
        let sheet = &self.sheets[index];
        self.actual_sheet_name = sheet.name.clone();
        let bytes = self
            .zip
            .as_mut()
            .ok_or_else(|| SpreadsheetError::InvalidFormat("reader is not open".into()))?
            .by_name(&sheet.path)
            .map_err(|_| {
                SpreadsheetError::InvalidFormat(format!("missing worksheet part: {}", sheet.path))
            })
            .and_then(|mut entry| {
                let size = usize::try_from(entry.size()).map_err(|_| {
                    SpreadsheetError::InvalidFormat(
                        "ZIP entry is too large for this platform".into(),
                    )
                })?;
                let mut bytes = Vec::with_capacity(size);
                entry.read_to_end(&mut bytes)?;
                Ok(bytes)
            })?;
        let mut reader =
            XmlReader::from_reader(BufReader::with_capacity(64 * 1024, Cursor::new(bytes)));
        reader.config_mut().trim_text(false);
        self.row_reader = Some(reader);
        Ok(())
    }

    fn select_sheet_index(&mut self, index: usize) -> SpreadsheetResult<()> {
        self.init_sheet(index)?;
        self.current_sheet_index = index as i64;
        self.current_row.clear();
        self.field_count = 0;
        self.row_count = 0;
        Ok(())
    }

    fn read_next_row(&mut self) -> SpreadsheetResult<bool> {
        self.current_row.clear();
        let mut next_col = 0usize;
        let mut in_row = false;
        loop {
            self.row_buf.clear();
            let event = self
                .row_reader
                .as_mut()
                .ok_or_else(|| SpreadsheetError::InvalidFormat("reader is not open".into()))?
                .read_event_into(&mut self.row_buf)
                .map_err(|e| SpreadsheetError::InvalidFormat(format!("XLSX XML: {e}")))?;
            match event {
                Event::Start(element) if element.local_name().as_ref() == b"row" => {
                    in_row = true;
                }
                Event::End(element) if element.local_name().as_ref() == b"row" => {
                    if in_row {
                        return Ok(true);
                    }
                }
                Event::Empty(element) if element.local_name().as_ref() == b"row" => {
                    return Ok(true);
                }
                Event::Start(element) if element.local_name().as_ref() == b"c" && in_row => {
                    let (column, cell_type, style) = parse_cell_attributes(&element);
                    let context = XlsxParseContext {
                        shared_strings: &self.shared_strings,
                        formats: &self.xf_id_to_num_fmt_id,
                        custom_date_formats: &self.custom_date_formats,
                        uses_1904_date_system: self.uses_1904_date_system,
                    };
                    let value = read_owned_cell_value(
                        self.row_reader.as_mut().expect("row reader exists"),
                        &mut self.row_buf,
                        cell_type,
                        style,
                        &context,
                    )?;
                    let col = column.map(|value| value as usize).unwrap_or(next_col);
                    if col >= crate::EXCEL_MAX_COLUMNS {
                        return Err(SpreadsheetError::InvalidFormat(
                            "XLSX column index is outside Excel worksheet bounds".into(),
                        ));
                    }
                    next_col = col + 1;
                    if self.current_row.len() <= col {
                        self.current_row.resize(col + 1, CellValue::Empty);
                    }
                    self.current_row[col] = value;
                    self.field_count = self.field_count.max(col + 1);
                }
                Event::Empty(element) if element.local_name().as_ref() == b"c" && in_row => {
                    let (column, _, _) = parse_cell_attributes(&element);
                    let col = column.map(|value| value as usize).unwrap_or(next_col);
                    if col >= crate::EXCEL_MAX_COLUMNS {
                        return Err(SpreadsheetError::InvalidFormat(
                            "XLSX column index is outside Excel worksheet bounds".into(),
                        ));
                    }
                    next_col = col + 1;
                    if self.current_row.len() <= col {
                        self.current_row.resize(col + 1, CellValue::Empty);
                    }
                    self.field_count = self.field_count.max(col + 1);
                }
                Event::Eof => return Ok(false),
                _ => {}
            }
        }
    }

    pub fn get_value(&self, i: usize) -> CellValue {
        self.current_row.get(i).cloned().unwrap_or(CellValue::Empty)
    }

    pub fn current_row(&self) -> &[CellValue] {
        &self.current_row
    }

    pub fn close(&mut self) {
        self.row_reader = None;
        self.current_sheet_index = -1;
    }
}

/// Streaming XLSX cell reader. The ZIP member and shared-string table are
/// borrowed from the parent reader; no worksheet-wide XML buffer is created.
pub struct XlsxCellReader<'a> {
    xml: XmlReader<BufReader<ZipFile<'a>>>,
    shared_strings: &'a [String],
    formats: &'a [u16],
    custom_date_formats: &'a HashSet<u16>,
    uses_1904_date_system: bool,
    buf: Vec<u8>,
    cell_buf: Vec<u8>,
    value_buf: Vec<u8>,
    scratch: String,
    row: u32,
    column: u32,
}

impl<'a> XlsxCellReader<'a> {
    fn new(
        entry: ZipFile<'a>,
        shared_strings: &'a [String],
        formats: &'a [u16],
        custom_date_formats: &'a HashSet<u16>,
        uses_1904_date_system: bool,
    ) -> SpreadsheetResult<Self> {
        // Worksheet XML is read sequentially from a compressed ZIP member;
        // a larger buffer reduces decompressor/read-call overhead on large
        // sheets without retaining the worksheet in memory.
        let mut xml = XmlReader::from_reader(BufReader::with_capacity(64 * 1024, entry));
        xml.config_mut().trim_text(false);
        Ok(Self {
            xml,
            shared_strings,
            formats,
            custom_date_formats,
            uses_1904_date_system,
            buf: Vec::with_capacity(1024),
            cell_buf: Vec::with_capacity(256),
            value_buf: Vec::with_capacity(256),
            scratch: String::with_capacity(128),
            row: 0,
            column: 0,
        })
    }

    /// Return the next explicit worksheet cell in document order.
    pub fn next_cell<'b>(&'b mut self) -> SpreadsheetResult<Option<crate::CellRef<'b>>> {
        loop {
            self.buf.clear();
            let event_kind = {
                let event = self
                    .xml
                    .read_event_into(&mut self.buf)
                    .map_err(|e| SpreadsheetError::InvalidFormat(format!("XLSX XML: {e}")))?;
                match event {
                    Event::Start(element) if element.local_name().as_ref() == b"row" => {
                        EventKind::Row(parse_row_number(&element))
                    }
                    Event::End(element) if element.local_name().as_ref() == b"row" => {
                        EventKind::EndRow
                    }
                    Event::Start(element) if element.local_name().as_ref() == b"c" => {
                        EventKind::Cell(parse_cell_attributes(&element))
                    }
                    Event::Empty(element) if element.local_name().as_ref() == b"c" => {
                        EventKind::EmptyCell(parse_cell_attributes(&element).0)
                    }
                    Event::End(element) if element.local_name().as_ref() == b"sheetData" => {
                        EventKind::EndSheet
                    }
                    Event::Eof => EventKind::Eof,
                    _ => EventKind::Other,
                }
            };

            match event_kind {
                EventKind::Row(value) => {
                    if let Some(row) = value {
                        self.row = row;
                    }
                }
                EventKind::EndRow => {
                    self.row = self.row.saturating_add(1);
                    self.column = 0;
                }
                EventKind::Cell((column, cell_type, style)) => {
                    let column = column.unwrap_or(self.column);
                    self.column = column.saturating_add(1);
                    let row = self.row;
                    let value = self.read_cell_value(cell_type, style)?;
                    return Ok(Some(crate::CellRef { row, column, value }));
                }
                EventKind::EmptyCell(column) => {
                    let column = column.unwrap_or(self.column);
                    self.column = column.saturating_add(1);
                    return Ok(Some(crate::CellRef {
                        row: self.row,
                        column,
                        value: CellValueRef::Empty,
                    }));
                }
                EventKind::EndSheet | EventKind::Eof => return Ok(None),
                EventKind::Other => {}
            }
        }
    }

    fn read_cell_value(
        &mut self,
        cell_type: XlsxCellType,
        style: usize,
    ) -> SpreadsheetResult<CellValueRef<'_>> {
        self.scratch.clear();
        let mut value = ParsedCellValue::Empty;
        let formats = self.formats;
        let custom_date_formats = self.custom_date_formats;
        let uses_1904_date_system = self.uses_1904_date_system;
        loop {
            self.cell_buf.clear();
            let event = self
                .xml
                .read_event_into(&mut self.cell_buf)
                .map_err(|e| SpreadsheetError::InvalidFormat(format!("XLSX cell XML: {e}")))?;
            match event {
                Event::Start(element) if element.local_name().as_ref() == b"v" => {
                    self.scratch.clear();
                    loop {
                        self.value_buf.clear();
                        let value_event =
                            self.xml.read_event_into(&mut self.value_buf).map_err(|e| {
                                SpreadsheetError::InvalidFormat(format!("XLSX value XML: {e}"))
                            })?;
                        match value_event {
                            Event::Text(text) => {
                                append_unescaped(&mut self.scratch, text.as_ref())?;
                            }
                            Event::CData(text) => {
                                append_unescaped(&mut self.scratch, text.as_ref())?;
                            }
                            Event::GeneralRef(reference) => {
                                self.scratch
                                    .push_str(&decode_xml_reference(reference.as_ref()));
                            }
                            Event::End(end) if end.local_name().as_ref() == b"v" => break,
                            Event::Eof => {
                                return Err(SpreadsheetError::InvalidFormat(
                                    "unterminated XLSX value".into(),
                                ))
                            }
                            _ => {}
                        }
                    }
                    let parsed = parse_value(
                        self.scratch.as_bytes(),
                        cell_type,
                        style,
                        formats,
                        custom_date_formats,
                        uses_1904_date_system,
                    )?;
                    value = match parsed {
                        ParsedCellValue::OwnedText(text) => {
                            self.scratch = text;
                            ParsedCellValue::ScratchText
                        }
                        other => other,
                    };
                }
                Event::Start(element) if element.local_name().as_ref() == b"is" => {
                    self.read_inline_string()?;
                    value = ParsedCellValue::ScratchText;
                }
                Event::Start(element) if element.local_name().as_ref() == b"t" => {
                    if cell_type == XlsxCellType::InlineString {
                        self.value_buf.clear();
                        if let Ok(Event::Text(text)) = self.xml.read_event_into(&mut self.value_buf)
                        {
                            append_unescaped(&mut self.scratch, text.as_ref())?;
                            value = ParsedCellValue::ScratchText;
                        }
                    }
                }
                Event::End(element) if element.local_name().as_ref() == b"c" => {
                    return Ok(match value {
                        ParsedCellValue::Empty => CellValueRef::Empty,
                        ParsedCellValue::SharedString(index) => self
                            .shared_strings
                            .get(index)
                            .map(String::as_str)
                            .map(CellValueRef::Text)
                            .unwrap_or(CellValueRef::Empty),
                        ParsedCellValue::ScratchText => CellValueRef::Text(self.scratch.as_str()),
                        ParsedCellValue::Number(number) => CellValueRef::Number(number),
                        ParsedCellValue::Boolean(boolean) => CellValueRef::Boolean(boolean),
                        ParsedCellValue::DateTime(date_time) => CellValueRef::DateTime(date_time),
                        ParsedCellValue::OwnedText(text) => {
                            self.scratch = text;
                            CellValueRef::Text(self.scratch.as_str())
                        }
                    });
                }
                Event::Empty(element) if element.local_name().as_ref() == b"c" => {
                    return Ok(CellValueRef::Empty)
                }
                Event::Eof => {
                    return Err(SpreadsheetError::InvalidFormat(
                        "unterminated XLSX cell".into(),
                    ))
                }
                _ => {}
            }
        }
    }

    fn read_inline_string(&mut self) -> SpreadsheetResult<()> {
        let mut in_text = false;
        loop {
            self.value_buf.clear();
            let event = self
                .xml
                .read_event_into(&mut self.value_buf)
                .map_err(|e| SpreadsheetError::InvalidFormat(format!("XLSX inline XML: {e}")))?;
            match event {
                Event::Start(element) if element.local_name().as_ref() == b"t" => {
                    in_text = true;
                }
                Event::End(element) if element.local_name().as_ref() == b"t" => {
                    in_text = false;
                }
                Event::Text(text) if in_text => {
                    append_unescaped(&mut self.scratch, text.as_ref())?;
                }
                Event::GeneralRef(reference) if in_text => {
                    self.scratch
                        .push_str(&decode_xml_reference(reference.as_ref()));
                }
                Event::CData(text) if in_text => {
                    append_unescaped(&mut self.scratch, text.as_ref())?;
                }
                Event::End(element) if element.local_name().as_ref() == b"is" => return Ok(()),
                Event::Eof => {
                    return Err(SpreadsheetError::InvalidFormat(
                        "unterminated inline string".into(),
                    ))
                }
                _ => {}
            }
        }
    }
}

struct XlsxParseContext<'a> {
    shared_strings: &'a [String],
    formats: &'a [u16],
    custom_date_formats: &'a HashSet<u16>,
    uses_1904_date_system: bool,
}

fn read_owned_cell_value<R: BufRead>(
    reader: &mut XmlReader<R>,
    buf: &mut Vec<u8>,
    cell_type: XlsxCellType,
    style: usize,
    context: &XlsxParseContext<'_>,
) -> SpreadsheetResult<CellValue> {
    let mut parsed = ParsedCellValue::Empty;
    let mut inline_string = false;
    let mut in_value = false;
    let mut in_text = false;
    let mut value_text = String::new();
    loop {
        buf.clear();
        let event = reader
            .read_event_into(buf)
            .map_err(|e| SpreadsheetError::InvalidFormat(format!("XLSX cell XML: {e}")))?;
        match event {
            Event::Start(element) if element.local_name().as_ref() == b"v" => {
                in_value = true;
                value_text.clear();
            }
            Event::Start(element) if element.local_name().as_ref() == b"is" => {
                inline_string = true;
                in_text = false;
            }
            Event::Start(element) if inline_string && element.local_name().as_ref() == b"t" => {
                in_text = true;
            }
            Event::End(element) if element.local_name().as_ref() == b"t" => {
                in_text = false;
            }
            Event::Text(text) if in_value => {
                append_unescaped(&mut value_text, text.as_ref())?;
            }
            Event::CData(text) if in_value => {
                append_unescaped(&mut value_text, text.as_ref())?;
            }
            Event::GeneralRef(reference) if in_value => {
                value_text.push_str(&decode_xml_reference(reference.as_ref()));
            }
            Event::End(element) if in_value && element.local_name().as_ref() == b"v" => {
                parsed = parse_value(
                    value_text.as_bytes(),
                    cell_type,
                    style,
                    context.formats,
                    context.custom_date_formats,
                    context.uses_1904_date_system,
                )?;
                in_value = false;
            }
            Event::Text(text) if inline_string && in_text => {
                let mut value = match parsed {
                    ParsedCellValue::OwnedText(value) => value,
                    _ => String::new(),
                };
                value.push_str(&String::from_utf8_lossy(text.as_ref()));
                parsed = ParsedCellValue::OwnedText(value);
            }
            Event::CData(text) if inline_string && in_text => {
                let mut value = match parsed {
                    ParsedCellValue::OwnedText(value) => value,
                    _ => String::new(),
                };
                value.push_str(&String::from_utf8_lossy(text.as_ref()));
                parsed = ParsedCellValue::OwnedText(value);
            }
            Event::GeneralRef(reference) if inline_string && in_text => {
                let mut value = match parsed {
                    ParsedCellValue::OwnedText(value) => value,
                    _ => String::new(),
                };
                value.push_str(&decode_xml_reference(reference.as_ref()));
                parsed = ParsedCellValue::OwnedText(value);
            }
            Event::End(element) if element.local_name().as_ref() == b"c" => {
                return Ok(match parsed {
                    ParsedCellValue::Empty => CellValue::Empty,
                    ParsedCellValue::SharedString(index) => context
                        .shared_strings
                        .get(index)
                        .map(|value| CellValue::Text(value.clone()))
                        .unwrap_or(CellValue::Empty),
                    ParsedCellValue::ScratchText => CellValue::Empty,
                    ParsedCellValue::Number(number) => CellValue::Number(number),
                    ParsedCellValue::Boolean(value) => CellValue::Boolean(value),
                    ParsedCellValue::DateTime(value) => CellValue::DateTime(value),
                    ParsedCellValue::OwnedText(value) => CellValue::Text(value),
                });
            }
            Event::End(element) if element.local_name().as_ref() == b"is" => {
                inline_string = false;
                in_text = false;
            }
            Event::Eof => {
                return Err(SpreadsheetError::InvalidFormat(
                    "unterminated XLSX cell".into(),
                ))
            }
            _ => {}
        }
    }
}

fn parse_value(
    raw: &[u8],
    cell_type: XlsxCellType,
    style: usize,
    formats: &[u16],
    custom_date_formats: &HashSet<u16>,
    uses_1904_date_system: bool,
) -> SpreadsheetResult<ParsedCellValue> {
    match cell_type {
        XlsxCellType::SharedString => {
            let index = atoi_simd::parse::<usize, true, false>(raw).unwrap_or(usize::MAX);
            Ok(ParsedCellValue::SharedString(index))
        }
        XlsxCellType::Boolean => Ok(ParsedCellValue::Boolean(raw == b"1" || raw == b"true")),
        XlsxCellType::Number | XlsxCellType::Other | XlsxCellType::InlineString => {
            let number = match fast_float2::parse::<f64, _>(raw) {
                Ok(number) => number,
                Err(_) if cell_type == XlsxCellType::Other => {
                    return Ok(ParsedCellValue::OwnedText(
                        String::from_utf8_lossy(raw).into_owned(),
                    ))
                }
                Err(_) => {
                    return Err(SpreadsheetError::InvalidFormat(
                        "invalid XLSX numeric value".into(),
                    ))
                }
            };
            if style > 0
                && is_date_num_fmt(
                    formats.get(style).copied().unwrap_or(0),
                    custom_date_formats,
                )
            {
                Ok(ParsedCellValue::DateTime(datetime_from_excel_serial(
                    number,
                    uses_1904_date_system,
                )))
            } else {
                Ok(ParsedCellValue::Number(number))
            }
        }
    }
}

enum ParsedCellValue {
    Empty,
    SharedString(usize),
    ScratchText,
    Number(f64),
    Boolean(bool),
    DateTime(chrono::NaiveDateTime),
    OwnedText(String),
}

fn parse_row_number(element: &BytesStart<'_>) -> Option<u32> {
    element
        .attributes()
        .flatten()
        .find(|attr| attr.key.as_ref() == b"r")
        .and_then(|attr| parse_cell_reference(attr.value.as_ref()).map(|(row, _)| row))
}

fn parse_cell_attributes(element: &BytesStart<'_>) -> (Option<u32>, XlsxCellType, usize) {
    let mut column = None;
    let mut cell_type = XlsxCellType::Number;
    let mut style = 0;
    for attr in element.attributes().flatten() {
        match attr.key.as_ref() {
            b"r" => column = parse_cell_reference(attr.value.as_ref()).map(|(_, col)| col),
            b"s" => {
                style = atoi_simd::parse::<usize, true, false>(attr.value.as_ref()).unwrap_or(0)
            }
            b"t" => {
                cell_type = match attr.value.as_ref() {
                    b"s" => XlsxCellType::SharedString,
                    b"b" => XlsxCellType::Boolean,
                    b"inlineStr" | b"inline_string" => XlsxCellType::InlineString,
                    _ => XlsxCellType::Other,
                }
            }
            _ => {}
        }
    }
    (column, cell_type, style)
}

fn append_unescaped(scratch: &mut String, raw: &[u8]) -> SpreadsheetResult<()> {
    let raw = std::str::from_utf8(raw)
        .map_err(|e| SpreadsheetError::InvalidFormat(format!("invalid XLSX text: {e}")))?;
    scratch.push_str(raw);
    Ok(())
}

fn decode_xml_reference(reference: &[u8]) -> String {
    let raw = format!("&{};", String::from_utf8_lossy(reference));
    quick_xml::escape::unescape(&raw)
        .map(|value| value.into_owned())
        .unwrap_or(raw)
}

enum EventKind {
    Row(Option<u32>),
    EndRow,
    Cell((Option<u32>, XlsxCellType, usize)),
    EmptyCell(Option<u32>),
    EndSheet,
    Eof,
    Other,
}

fn parse_cell_reference(value: &[u8]) -> Option<(u32, u32)> {
    let mut split = 0;
    while split < value.len() && value[split].is_ascii_alphabetic() {
        split += 1;
    }
    if split == 0 || split == value.len() {
        return None;
    }
    let mut column = 0u32;
    for byte in &value[..split] {
        column = column
            .checked_mul(26)?
            .checked_add(byte.to_ascii_uppercase().checked_sub(b'A')? as u32 + 1)?;
    }
    let row = atoi_simd::parse::<u32, true, false>(&value[split..]).ok()?;
    Some((row.checked_sub(1)?, column.checked_sub(1)?))
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

impl Default for XlsxReader {
    fn default() -> Self {
        Self::new()
    }
}

fn is_date_num_fmt(id: u16, custom: &HashSet<u16>) -> bool {
    (14..=22).contains(&id) || (45..=47).contains(&id) || custom.contains(&id)
}

fn parse_sheet_entries(xml: &str) -> Vec<(String, String, String)> {
    let mut reader = XmlReader::from_str(xml);
    reader.config_mut().trim_text(false);
    let mut out = Vec::new();
    loop {
        match reader.read_event() {
            Ok(Event::Start(element)) | Ok(Event::Empty(element))
                if element.local_name().as_ref() == b"sheet" =>
            {
                let mut name = None;
                let mut sheet_id = None;
                let mut r_id = None;
                for attribute in element.attributes().flatten() {
                    let value = attribute
                        .decoded_and_normalized_value(
                            quick_xml::XmlVersion::default(),
                            reader.decoder(),
                        )
                        .map(|value| value.into_owned())
                        .unwrap_or_else(|_| {
                            String::from_utf8_lossy(attribute.value.as_ref()).into_owned()
                        });
                    match attribute.key.local_name().as_ref() {
                        b"name" => name = Some(value),
                        b"sheetId" => sheet_id = Some(value),
                        b"id" => r_id = Some(value),
                        _ => {}
                    }
                }
                if let (Some(name), Some(sheet_id), Some(r_id)) = (name, sheet_id, r_id) {
                    out.push((name, sheet_id, r_id));
                }
            }
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
    }
    out
}

fn parse_styles(xml: &str) -> (Vec<u16>, HashSet<u16>) {
    let mut reader = XmlReader::from_str(xml);
    reader.config_mut().trim_text(false);
    let mut customs = HashSet::new();
    let mut xfs = Vec::new();
    let mut in_cell_xfs = false;
    loop {
        match reader.read_event() {
            Ok(Event::Start(element)) | Ok(Event::Empty(element)) => {
                match element.local_name().as_ref() {
                    b"numFmt" => {
                        let mut id = None;
                        let mut code = None;
                        for attribute in element.attributes().flatten() {
                            let value = attribute
                                .decoded_and_normalized_value(
                                    quick_xml::XmlVersion::default(),
                                    reader.decoder(),
                                )
                                .map(|value| value.into_owned())
                                .unwrap_or_else(|_| {
                                    String::from_utf8_lossy(attribute.value.as_ref()).into_owned()
                                });
                            match attribute.key.as_ref() {
                                b"numFmtId" => id = value.parse::<u16>().ok(),
                                b"formatCode" => code = Some(value),
                                _ => {}
                            }
                        }
                        if let (Some(id), Some(code)) = (id, code) {
                            if is_date_format_code(&code) {
                                customs.insert(id);
                            }
                        }
                    }
                    b"cellXfs" => in_cell_xfs = true,
                    b"xf" if in_cell_xfs => {
                        for attribute in element.attributes().flatten() {
                            if attribute.key.as_ref() == b"numFmtId" {
                                if let Ok(id) = std::str::from_utf8(attribute.value.as_ref())
                                    .unwrap_or_default()
                                    .parse::<u16>()
                                {
                                    xfs.push(id);
                                }
                            }
                        }
                    }
                    _ => {}
                }
            }
            Ok(Event::End(element)) if element.local_name().as_ref() == b"cellXfs" => {
                in_cell_xfs = false;
            }
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
    }
    (xfs, customs)
}

/// Cell storage type from the XLSX `t` attribute.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum XlsxCellType {
    Number,
    SharedString,
    Boolean,
    InlineString,
    Other,
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::{datetime_to_excel_serial, CellValue};
    use std::io::Write;

    #[test]
    fn cell_attrs_parse() {
        let mut reader = XmlReader::from_reader(Cursor::new(b"<c r=\"B2\" t=\"s\" s=\"3\"/>"));
        let mut buf = Vec::new();
        let event = reader.read_event_into(&mut buf).unwrap();
        let Event::Empty(element) = event else {
            panic!("expected empty cell element");
        };
        let (col, t, s) = parse_cell_attributes(&element);
        assert_eq!(col, Some(1));
        assert_eq!(t, XlsxCellType::SharedString);
        assert_eq!(s, 3);
    }

    #[test]
    fn parses_date_serial_using_1904_system() {
        let expected = chrono::NaiveDate::from_ymd_opt(2025, 2, 3)
            .unwrap()
            .and_hms_opt(12, 30, 0)
            .unwrap();
        let serial = datetime_to_excel_serial(&expected, true);
        let mut reader = XlsxReader::new();
        reader.uses_1904_date_system = true;
        reader.xf_id_to_num_fmt_id = vec![0, 14];
        let xml = format!(r#"<sheetData><row><c s="1"><v>{serial}</v></c></row></sheetData>"#);
        let mut row_reader = XmlReader::from_reader(BufReader::new(Cursor::new(xml.into_bytes())));
        row_reader.config_mut().trim_text(false);
        reader.row_reader = Some(row_reader);
        assert!(reader.read_next_row().unwrap());
        assert_eq!(reader.current_row, vec![CellValue::DateTime(expected)]);
    }

    #[test]
    fn reads_escaped_value_fragments_as_one_string() {
        let mut xml =
            XmlReader::from_reader(Cursor::new(br#"<c t="str"><v>A&amp;B</v></c>"#.to_vec()));
        let mut buf = Vec::new();
        let event = xml.read_event_into(&mut buf).unwrap();
        let Event::Start(_) = event else {
            panic!("expected cell start");
        };
        let context = XlsxParseContext {
            shared_strings: &[],
            formats: &[],
            custom_date_formats: &HashSet::new(),
            uses_1904_date_system: false,
        };
        assert_eq!(
            read_owned_cell_value(&mut xml, &mut buf, XlsxCellType::Other, 0, &context).unwrap(),
            CellValue::Text("A&B".into())
        );
    }

    #[test]
    fn sheet_entries_decode_names_and_local_relationship_names() {
        let entries = parse_sheet_entries(
            r#"<workbook><sheets><sheet name="A&amp;B" sheetId="1" rel:id="rId1"/></sheets></workbook>"#,
        );
        assert_eq!(entries, vec![("A&B".into(), "1".into(), "rId1".into())]);
    }

    #[test]
    fn escaped_literal_format_codes_are_not_dates() {
        let (xfs, custom_dates) = parse_styles(
            r#"<styleSheet><numFmts><numFmt numFmtId="200" formatCode="0 &quot;m&quot;"/></numFmts><cellXfs><xf numFmtId="200"/></cellXfs></styleSheet>"#,
        );
        assert_eq!(xfs, vec![200]);
        assert!(!custom_dates.contains(&200));
    }

    #[test]
    fn streaming_reader_accumulates_entities_and_skips_inline_indentation() {
        let dir = tempfile::tempdir().unwrap();
        let path = dir.path().join("escaped.xlsx");
        let file = File::create(&path).unwrap();
        let mut zip = zip::ZipWriter::new(file);
        let options = zip::write::SimpleFileOptions::default()
            .last_modified_time(crate::writer_helpers::deterministic_zip_timestamp());
        zip.start_file("xl/workbook.xml", options).unwrap();
        zip.write_all(
            br#"<workbook xmlns:r="urn:test"><sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets></workbook>"#,
        )
        .unwrap();
        zip.start_file("xl/_rels/workbook.xml.rels", options)
            .unwrap();
        zip.write_all(
            br#"<Relationships><Relationship Id="rId1" Target="worksheets/sheet1.xml"/></Relationships>"#,
        )
        .unwrap();
        zip.start_file("xl/worksheets/sheet1.xml", options).unwrap();
        zip.write_all(
            br#"<worksheet><sheetData><row r="1"><c r="A1" t="str"><v>A&amp;B</v></c><c r="B1" t="inlineStr"><is>
  <r><t>left</t></r>
  <r><t>right&amp;</t></r>
</is></c></row></sheetData></worksheet>"#,
        )
        .unwrap();
        zip.finish().unwrap();

        let mut reader = XlsxReader::new();
        reader.open(&path, false).unwrap();
        let mut cells = reader.cell_reader("Data").unwrap();
        let first = cells.next_cell().unwrap().unwrap();
        assert_eq!(first.row, 0);
        assert_eq!(first.column, 0);
        assert_eq!(first.value, CellValueRef::Text("A&B"));
        let second = cells.next_cell().unwrap().unwrap();
        assert_eq!(second.column, 1);
        assert_eq!(second.value, CellValueRef::Text("leftright&"));
        assert!(cells.next_cell().unwrap().is_none());
    }
}
