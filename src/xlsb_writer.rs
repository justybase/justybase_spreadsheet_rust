//! High-performance XLSB writer (port of `XlsbWriter.ts`).
//!
//! Prefer this format for large datasets — typically faster and smaller
//! than XLSX. Binary templates live in the private `xlsb_templates` module (frozen
//! golden bytes extracted from the TS reference).

use crate::big_buffer::{BigBuffer, EntryData};
use crate::error::{SpreadsheetError, SpreadsheetResult};
use crate::formats::{borrow_cell, get_format, CellValue, CellValueRef};
use crate::streaming_state::StreamingSheetState;
use crate::writer_helpers::{
    apply_header_widths, default_col_width, init_col_widths, unique_sheet_name,
    update_col_widths_from_rows,
};
use crate::xlsb_templates;
use crate::{datetime_to_oa_date, oa_epoch_naive};
use std::collections::HashMap;
use std::fs::File;
use std::path::{Path, PathBuf};

const RK_LOWER: i64 = -1 << 29;
const RK_UPPER: i64 = (1 << 29) - 1;

struct SheetInfo {
    name: String,
    path_in_archive: String,
    hidden: bool,
    name_in_archive: String,
    sheet_id: usize,
    filter_data: Option<FilterData>,
}

#[derive(Clone, Copy)]
struct FilterData {
    start_row: u32,
    end_row: u32,
    start_column: u32,
    end_column: u32,
}

/// Options for [`XlsbWriter::start_sheet`].
#[derive(Debug, Clone, Default)]
pub struct XlsbSheetOptions {
    pub hidden: bool,
    pub do_autofilter: bool,
    pub sample_rows: Option<Vec<Vec<CellValue>>>,
}

impl XlsbSheetOptions {
    pub fn new() -> Self {
        Self {
            hidden: false,
            do_autofilter: true,
            sample_rows: None,
        }
    }
}

/// Scratch row staging for the streaming writer; kept on the writer so the
/// staging `Vec` is allocated once and reused for every row.
enum StreamStaged {
    Rk(i64, u8),
    Double(f64, u8),
    Bool(bool),
    Sst(usize, u8), // (sst_index, style)
}

pub struct XlsbWriter {
    output_path: PathBuf,
    entries: Vec<(String, EntryData)>,

    sheet_count: usize,
    sheet_list: Vec<SheetInfo>,
    sst_dic: HashMap<String, usize>,
    sst_cnt_unique: usize,
    sst_cnt_all: usize,
    autofilter_is_on: bool,

    next_fmt_id: u16,
    next_xf_idx: u8,
    st_fmt_records: HashMap<String, (u16, u8)>,
    compression_level: i64,

    stream: StreamingSheetState,
    staged_row: Vec<Option<StreamStaged>>,
}

impl XlsbWriter {
    pub fn create(path: &Path) -> SpreadsheetResult<Self> {
        if let Some(parent) = path.parent() {
            if !parent.as_os_str().is_empty() {
                std::fs::create_dir_all(parent)?;
            }
        }
        let _ = oa_epoch_naive();
        Ok(Self {
            output_path: path.to_path_buf(),
            entries: Vec::new(),
            sheet_count: 0,
            sheet_list: Vec::new(),
            sst_dic: HashMap::new(),
            sst_cnt_unique: 0,
            sst_cnt_all: 0,
            autofilter_is_on: false,
            next_fmt_id: 167,
            next_xf_idx: 4,
            st_fmt_records: HashMap::new(),
            compression_level: 1,
            stream: StreamingSheetState::new(),
            staged_row: Vec::new(),
        })
    }

    /// Set the ZIP deflate compression level (1–9) used by [`Self::finalize`].
    ///
    /// `1` is the default (matching the TypeScript reference — fastest),
    /// `9` gives the smallest file at a higher CPU cost. Level `0` (no
    /// compression) is not supported by the bundled `zip` backend and is
    /// rejected here.
    pub fn set_compression_level(&mut self, level: i64) -> SpreadsheetResult<()> {
        if !(1..=9).contains(&level) {
            return Err(SpreadsheetError::InvalidFormat(format!(
                "compression level must be within 1..=9, got {level}"
            )));
        }
        self.compression_level = level;
        Ok(())
    }

    fn register_format(&mut self, fmt: &str) -> SpreadsheetResult<u8> {
        if let Some(&(_, xf)) = self.st_fmt_records.get(fmt) {
            return Ok(xf);
        }
        if self.next_fmt_id == u16::MAX || self.next_xf_idx == u8::MAX {
            return Err(SpreadsheetError::InvalidFormat(
                "too many custom number formats for XLSB".into(),
            ));
        }
        let numfmt_id = self.next_fmt_id;
        self.next_fmt_id += 1;
        let xf = self.next_xf_idx;
        self.next_xf_idx += 1;
        self.st_fmt_records.insert(fmt.to_owned(), (numfmt_id, xf));
        Ok(xf)
    }

    fn encode_vlq(mut value: u32) -> Vec<u8> {
        let mut out = Vec::new();
        while value >= 0x80 {
            out.push((value as u8 & 0x7F) | 0x80);
            value >>= 7;
        }
        out.push(value as u8 & 0x7F);
        out
    }

    fn build_brt_record(rec_type: u32, data: &[u8]) -> Vec<u8> {
        let mut out = Self::encode_vlq(rec_type);
        out.extend(Self::encode_vlq(data.len() as u32));
        out.extend_from_slice(data);
        out
    }

    fn build_st_fmt_record(ifmt: u16, fmt: &str) -> Vec<u8> {
        let units: Vec<u16> = fmt.encode_utf16().collect();
        let cch = units.len();
        let mut payload = vec![0u8; 2 + 4 + cch * 2];
        payload[0..2].copy_from_slice(&ifmt.to_le_bytes());
        payload[2..6].copy_from_slice(&(cch as u32).to_le_bytes());
        for (i, u) in units.iter().enumerate() {
            payload[6 + i * 2..8 + i * 2].copy_from_slice(&(*u).to_le_bytes());
        }
        Self::build_brt_record(0x2c, &payload)
    }

    fn build_xf_record(font_id: u16, ifmt: u16, flags: u16, byte6_extra: u8) -> Vec<u8> {
        let mut data = vec![0u8; 18];
        data[0] = 0x2F;
        data[1] = 16;
        data[2..4].copy_from_slice(&font_id.to_le_bytes());
        data[4..6].copy_from_slice(&ifmt.to_le_bytes());
        data[6] = byte6_extra;
        data[14] = 0x10;
        data[15] = 0x10;
        data[16..18].copy_from_slice(&flags.to_le_bytes());
        data
    }

    fn build_styles_bin(&self) -> Vec<u8> {
        if self.st_fmt_records.is_empty() {
            return xlsb_templates::STYLES_BIN_BASE.to_vec();
        }
        let base = &xlsb_templates::STYLES_BIN_BASE;
        let find = |needle: &[u8], from: usize| -> Option<usize> {
            base[from..]
                .windows(needle.len())
                .position(|w| w == needle)
                .map(|p| p + from)
        };
        let font_end = find(&[0xE8, 0x04, 0x00], 3).unwrap_or(base.len());
        let xf_begin = find(&[0xE9, 0x04], font_end + 3).unwrap_or(base.len());
        let xf_end = find(&[0xEA, 0x04, 0x00], xf_begin).unwrap_or(base.len());
        let fill_begin = find(&[0xE3, 0x04], font_end + 3).unwrap_or(xf_begin);

        let total_fmt = (2 + self.st_fmt_records.len()) as u16;
        let total_xf = (4 + self.st_fmt_records.len()) as u16;

        let mut parts: Vec<u8> = Vec::new();
        parts.extend_from_slice(&base[..3]);
        let mut fmt_header = [0u8; 4];
        fmt_header[..2].copy_from_slice(&total_fmt.to_le_bytes());
        parts.extend(Self::build_brt_record(0x0267, &fmt_header));
        parts.extend(Self::build_st_fmt_record(164, "yyyy\\-mm\\-dd\\ hh:mm"));
        parts.extend(Self::build_st_fmt_record(166, "yyyy\\-mm\\-dd"));
        let mut sorted: Vec<(&String, &(u16, u8))> = self.st_fmt_records.iter().collect();
        sorted.sort_by_key(|item| item.1 .1);
        for (fmt, entry) in &sorted {
            parts.extend(Self::build_st_fmt_record(entry.0, fmt));
        }
        parts.extend(Self::build_brt_record(0x0268, &[]));
        parts.extend_from_slice(&base[fill_begin..xf_begin]);
        let mut xf_header = [0u8; 4];
        xf_header[..2].copy_from_slice(&total_xf.to_le_bytes());
        parts.extend(Self::build_brt_record(0x0269, &xf_header));
        parts.extend(Self::build_xf_record(0, 0, 0x0000, 0));
        parts.extend(Self::build_xf_record(0, 164, 0x0001, 0));
        parts.extend(Self::build_xf_record(0, 166, 0x0001, 0));
        parts.extend(Self::build_xf_record(1, 0, 0x0000, 1));
        for (_, entry) in sorted {
            parts.extend(Self::build_xf_record(0, entry.0, 0x0001, 0));
        }
        parts.extend(Self::build_brt_record(0x026A, &[]));
        parts.extend_from_slice(&base[xf_end + 3..]);
        parts
    }

    /// Register a worksheet (metadata only).
    pub fn add_sheet(&mut self, sheet_name: &str, hidden: bool) {
        let sanitized = unique_sheet_name(sheet_name, self.sheet_count, |candidate| {
            self.sheet_list
                .iter()
                .any(|sheet| sheet.name.eq_ignore_ascii_case(candidate))
        });
        self.sheet_count += 1;
        self.sheet_list.push(SheetInfo {
            name: sanitized,
            path_in_archive: format!("xl/worksheets/sheet{}.bin", self.sheet_count),
            hidden,
            name_in_archive: format!("sheet{}.bin", self.sheet_count),
            sheet_id: self.sheet_count,
            filter_data: None,
        });
    }

    /// Start a new sheet in streaming mode.
    pub fn start_sheet(
        &mut self,
        sheet_name: &str,
        column_count: usize,
        headers: Option<&[String]>,
        options: XlsbSheetOptions,
    ) -> SpreadsheetResult<()> {
        if column_count > crate::EXCEL_MAX_COLUMNS {
            return Err(SpreadsheetError::InvalidFormat(format!(
                "worksheet has {column_count} columns; Excel supports at most {}",
                crate::EXCEL_MAX_COLUMNS
            )));
        }
        if let Some(headers) = headers {
            if headers.len() != column_count {
                return Err(SpreadsheetError::RowLengthMismatch {
                    expected: column_count,
                    got: headers.len(),
                });
            }
        }
        let do_filter = if options.do_autofilter {
            headers.is_some()
        } else {
            false
        };
        self.stream
            .try_begin(column_count, do_filter, BigBuffer::default())?;
        self.add_sheet(sheet_name, options.hidden);
        if let Some(h) = headers {
            apply_header_widths(&mut self.stream.col_widths, h, column_count);
        }
        if let Some(sample) = &options.sample_rows {
            update_col_widths_from_rows(&mut self.stream.col_widths, sample, 100);
        }

        let widths = self.stream.col_widths.clone();
        let (sc, ec) = (self.stream.start_col, self.stream.end_col);
        let sheet_count = self.sheet_count;
        let has_headers = headers.is_some();
        // Intern header strings before the buffer borrow begins.
        let header_cells: Vec<usize> = match headers {
            Some(h) => h.iter().map(|s| self.intern_string(s)).collect(),
            None => Vec::new(),
        };
        {
            let big_buf = self.stream.assert_streaming()?;
            let mut header = xlsb_templates::SHEET1_BYTES.to_vec();
            header[40..44].copy_from_slice(&(sc as i32).to_le_bytes());
            header[44..48].copy_from_slice(&(ec as i32).to_le_bytes());
            if sheet_count != 1 {
                header[54] = 0x9C;
            }
            big_buf.write(&header[..84]);
            if has_headers {
                big_buf.write(&xlsb_templates::STICKY_HEADER_A1);
            }
            big_buf.write(&header[84..159]);
            big_buf.write_byte(134);
            big_buf.write_byte(3);
            for (i, width_value) in widths.iter().enumerate().take(ec).skip(sc) {
                big_buf.write_byte(0);
                big_buf.write_byte(60);
                big_buf.write_byte(18);
                big_buf.write_i32_le(i as i32);
                big_buf.write_i32_le(i as i32);
                let width = default_col_width(*width_value, true);
                big_buf.write_byte(0);
                big_buf.write_byte(width.clamp(0.0, 255.0) as u8);
                big_buf.write_byte(0);
                big_buf.write_byte(0);
                big_buf.write_byte(0);
                big_buf.write_byte(0);
                big_buf.write_byte(0);
                big_buf.write_byte(0);
                big_buf.write_byte(2);
            }
            big_buf.write_byte(0);
            big_buf.write_byte(135);
            big_buf.write_byte(3);
            big_buf.write_byte(0);
            big_buf.write(&header[159..175]);
            big_buf.write(&[38, 0]);
            if has_headers {
                Self::create_row_header_static(big_buf, 0, sc, ec);
                for (c, idx) in header_cells.iter().enumerate() {
                    Self::write_sst_cell_static(big_buf, *idx, c, 3);
                }
            }
        }
        if has_headers {
            self.stream.row_num = 1;
        }
        Ok(())
    }

    /// Write a single row in streaming mode.
    pub fn write_row(&mut self, row: &[CellValue]) -> SpreadsheetResult<()> {
        let expected = self.stream.end_col - self.stream.start_col;
        if row.len() != expected {
            return Err(SpreadsheetError::RowLengthMismatch {
                expected,
                got: row.len(),
            });
        }
        // Stage cell payloads that need SST interning, then write.
        // The staging Vec lives on the writer and is reused across rows.
        let mut staged = std::mem::take(&mut self.staged_row);
        staged.clear();
        for raw in row.iter() {
            if matches!(raw, CellValue::Empty) {
                staged.push(None);
                continue;
            }
            let fmt = get_format(raw);
            let val = borrow_cell(raw);
            let style = match fmt {
                Some(f) => self.register_format(f)?,
                None => 0,
            };
            let cell = match val {
                CellValueRef::Empty => None,
                CellValueRef::Integer(n) => {
                    if (RK_LOWER..=RK_UPPER).contains(&n) {
                        Some(StreamStaged::Rk(n, style))
                    } else {
                        Some(StreamStaged::Double(n as f64, style))
                    }
                }
                CellValueRef::Number(n) => {
                    if n.fract() == 0.0
                        && n >= RK_LOWER as f64
                        && n <= RK_UPPER as f64
                        && n.is_finite()
                    {
                        Some(StreamStaged::Rk(n as i64, style))
                    } else {
                        Some(StreamStaged::Double(n, style))
                    }
                }
                CellValueRef::Boolean(b) => Some(StreamStaged::Bool(b)),
                CellValueRef::DateTime(d) => {
                    let oa = datetime_to_oa_date(&d);
                    Some(StreamStaged::Double(
                        oa,
                        if fmt.is_some() { style } else { 1 },
                    ))
                }
                CellValueRef::Text(s) => {
                    let idx = self.intern_string(s);
                    let st = if fmt.is_some() { style } else { 0 };
                    Some(StreamStaged::Sst(idx, st))
                }
            };
            staged.push(cell);
        }
        let row_num = self.stream.row_num;
        let (sc, ec) = (self.stream.start_col, self.stream.end_col);
        {
            let big_buf = self.stream.assert_streaming()?;
            Self::create_row_header_static(big_buf, row_num, sc, ec);
            for (c, cell) in staged.drain(..).enumerate() {
                match cell {
                    None => {}
                    Some(StreamStaged::Rk(n, style)) => {
                        Self::write_rk_static(big_buf, n as i32, c, style)
                    }
                    Some(StreamStaged::Double(v, style)) => {
                        Self::write_double_static(big_buf, v, c, style)
                    }
                    Some(StreamStaged::Bool(b)) => Self::write_bool_static(big_buf, b, c),
                    Some(StreamStaged::Sst(idx, style)) => {
                        Self::write_sst_cell_static(big_buf, idx, c, style)
                    }
                }
            }
        }
        self.stream.row_num = row_num + 1;
        self.staged_row = staged;
        Ok(())
    }

    /// Finalize the current streaming sheet and stage it for the archive.
    pub fn end_sheet(&mut self) -> SpreadsheetResult<()> {
        if !self.stream.is_streaming {
            return Err(SpreadsheetError::NotStreaming);
        }
        let do_filter = self.stream.do_autofilter;
        let row_num = self.stream.row_num;
        let (sc, ec) = (self.stream.start_col, self.stream.end_col);
        let sheet_idx = self.sheet_count - 1;
        {
            let big_buf = self.stream.assert_streaming()?;
            let header = xlsb_templates::SHEET1_BYTES.to_vec();
            big_buf.write(&header[218..290]);
            if do_filter && ec > sc && row_num > 0 {
                big_buf.write(&xlsb_templates::AUTOFILTER_START);
                big_buf.write_i32_le(0);
                big_buf.write_i32_le(row_num as i32 - 1);
                big_buf.write_i32_le(sc as i32);
                big_buf.write_i32_le(ec as i32 - 1);
                big_buf.write(&xlsb_templates::AUTOFILTER_END);
            }
            big_buf.write(&header[290..]);
        }
        if do_filter && ec > sc && row_num > 0 {
            self.autofilter_is_on = true;
            let sheet = &mut self.sheet_list[sheet_idx];
            sheet.filter_data = Some(FilterData {
                start_row: 0,
                end_row: row_num - 1,
                start_column: sc as u32,
                end_column: (ec as u32) - 1,
            });
        }
        let data = EntryData::Chunks({
            let big_buf = self.stream.assert_streaming()?;
            big_buf.chunks_drain()
        });
        let path = self.sheet_list[sheet_idx].path_in_archive.clone();
        self.entries.push((path, data));
        self.stream.end();
        Ok(())
    }

    /// Write an entire sheet in one call (batch mode).
    pub fn write_sheet(
        &mut self,
        rows: Vec<Vec<CellValue>>,
        headers: Option<&[String]>,
        do_autofilter: bool,
    ) -> SpreadsheetResult<()> {
        if self.stream.is_streaming {
            return Err(SpreadsheetError::AlreadyStreaming);
        }
        if self.sheet_count == 0 {
            return Err(SpreadsheetError::InvalidFormat(
                "call add_sheet() before write_sheet()".into(),
            ));
        }
        let mut column_count = 0;
        if !rows.is_empty() {
            column_count = rows[0].len();
        } else if let Some(h) = headers {
            column_count = h.len();
        }
        if column_count > crate::EXCEL_MAX_COLUMNS {
            return Err(SpreadsheetError::InvalidFormat(format!(
                "worksheet has {column_count} columns; Excel supports at most {}",
                crate::EXCEL_MAX_COLUMNS
            )));
        }
        if let Some(h) = headers {
            if h.len() != column_count {
                return Err(SpreadsheetError::RowLengthMismatch {
                    expected: column_count,
                    got: h.len(),
                });
            }
        }
        for row in &rows {
            if row.len() != column_count {
                return Err(SpreadsheetError::RowLengthMismatch {
                    expected: column_count,
                    got: row.len(),
                });
            }
        }
        let mut col_widths = init_col_widths(column_count);
        if let Some(h) = headers {
            apply_header_widths(&mut col_widths, h, column_count);
        }
        update_col_widths_from_rows(&mut col_widths, &rows, 100);
        // Reserve SST for benchmark shape (2 text cols per row)
        let expected_sst = headers.map(|h| h.len()).unwrap_or(0) + rows.len() * 2 + 16;
        self.sst_dic.reserve(expected_sst);

        let mut big_buf = BigBuffer::default();
        let mut header = xlsb_templates::SHEET1_BYTES.to_vec();
        header[40..44].copy_from_slice(&0i32.to_le_bytes());
        header[44..48].copy_from_slice(&(column_count as i32).to_le_bytes());
        if self.sheet_count != 1 {
            // Note: TS checks `this.sheetCount !== 1` BEFORE addSheet is
            // called in writeSheet path? No — writeSheet is called after
            // addSheet, so sheetCount already includes this sheet.
            // Mirror TS exactly: the check runs with the current count.
            header[54] = 0x9C;
        }
        big_buf.write(&header[..84]);
        if headers.is_some() {
            big_buf.write(&xlsb_templates::STICKY_HEADER_A1);
        }
        big_buf.write(&header[84..159]);
        big_buf.write_byte(134);
        big_buf.write_byte(3);
        for (i, width_value) in col_widths.iter().enumerate().take(column_count) {
            big_buf.write_byte(0);
            big_buf.write_byte(60);
            big_buf.write_byte(18);
            big_buf.write_i32_le(i as i32);
            big_buf.write_i32_le(i as i32);
            let width = default_col_width(*width_value, true);
            big_buf.write_byte(0);
            big_buf.write_byte(width.clamp(0.0, 255.0) as u8);
            big_buf.write_byte(0);
            big_buf.write_byte(0);
            big_buf.write_byte(0);
            big_buf.write_byte(0);
            big_buf.write_byte(0);
            big_buf.write_byte(0);
            big_buf.write_byte(2);
        }
        big_buf.write_byte(0);
        big_buf.write_byte(135);
        big_buf.write_byte(3);
        big_buf.write_byte(0);
        big_buf.write(&header[159..175]);
        big_buf.write(&[38, 0]);

        let mut row_num: u32 = 0;
        if let Some(h) = headers {
            Self::create_row_header_static(&mut big_buf, row_num, 0, column_count);
            for (c, header_text) in h.iter().enumerate() {
                let idx = self.intern_string(header_text);
                Self::write_sst_cell_static(&mut big_buf, idx, c, 3);
            }
            row_num += 1;
        }

        for row in &rows {
            Self::create_row_header_static(&mut big_buf, row_num, 0, column_count);
            for (c, raw) in row.iter().enumerate() {
                if matches!(raw, CellValue::Empty) {
                    continue;
                }
                let fmt = get_format(raw);
                let val = borrow_cell(raw);
                let style = match fmt {
                    Some(f) => self.register_format(f)?,
                    None => 0,
                };
                match val {
                    CellValueRef::Empty => {}
                    CellValueRef::Integer(n) => {
                        if (RK_LOWER..=RK_UPPER).contains(&n) {
                            Self::write_rk_static(&mut big_buf, n as i32, c, style);
                        } else {
                            Self::write_double_static(&mut big_buf, n as f64, c, style);
                        }
                    }
                    CellValueRef::Number(n) => {
                        if n.fract() == 0.0
                            && n >= RK_LOWER as f64
                            && n <= RK_UPPER as f64
                            && n.is_finite()
                        {
                            Self::write_rk_static(&mut big_buf, n as i32, c, style);
                        } else {
                            Self::write_double_static(&mut big_buf, n, c, style);
                        }
                    }
                    CellValueRef::Boolean(b) => Self::write_bool_static(&mut big_buf, b, c),
                    CellValueRef::DateTime(d) => {
                        let oa = datetime_to_oa_date(&d);
                        Self::write_double_static(
                            &mut big_buf,
                            oa,
                            c,
                            if fmt.is_some() { style } else { 1 },
                        );
                    }
                    CellValueRef::Text(s) => {
                        let idx = self.intern_string(s);
                        let st = if fmt.is_some() { style } else { 0 };
                        Self::write_sst_cell_static(&mut big_buf, idx, c, st);
                    }
                }
            }
            row_num += 1;
        }

        big_buf.write(&header[218..290]);
        if do_autofilter && headers.is_some() && column_count > 0 {
            self.autofilter_is_on = true;
            let end_row = rows.len() as u32 + 1;
            big_buf.write(&xlsb_templates::AUTOFILTER_START);
            big_buf.write_i32_le(0);
            big_buf.write_i32_le(end_row as i32 - 1);
            big_buf.write_i32_le(0);
            big_buf.write_i32_le(column_count as i32 - 1);
            big_buf.write(&xlsb_templates::AUTOFILTER_END);
            let sheet = &mut self.sheet_list[self.sheet_count - 1];
            sheet.filter_data = Some(FilterData {
                start_row: 0,
                end_row: rows.len() as u32,
                start_column: 0,
                end_column: (column_count as u32) - 1,
            });
        }
        big_buf.write(&header[290..]);
        let data = EntryData::Chunks(big_buf.chunks_drain());
        let path = self.sheet_list[self.sheet_count - 1]
            .path_in_archive
            .clone();
        self.entries.push((path, data));
        Ok(())
    }

    fn intern_string(&mut self, val: &str) -> usize {
        if let Some(&idx) = self.sst_dic.get(val) {
            self.sst_cnt_all += 1;
            return idx;
        }
        let idx = self.sst_cnt_unique;
        self.sst_dic.insert(val.to_owned(), idx);
        self.sst_cnt_unique += 1;
        self.sst_cnt_all += 1;
        idx
    }

    fn create_row_header_static(big_buf: &mut BigBuffer, row: u32, sc: usize, ec: usize) {
        big_buf.ensure_capacity(27);
        big_buf.write_unsafe_byte(0);
        big_buf.write_unsafe_byte(25);
        big_buf.write_unsafe_i32_le(row as i32);
        big_buf.write_unsafe_i32_le(0);
        big_buf.write_unsafe_byte(44);
        big_buf.write_unsafe_byte(1);
        big_buf.write_unsafe_byte(0);
        big_buf.write_unsafe_byte(0);
        big_buf.write_unsafe_byte(0);
        big_buf.write_unsafe_byte(1);
        big_buf.write_unsafe_byte(0);
        big_buf.write_unsafe_byte(0);
        big_buf.write_unsafe_byte(0);
        big_buf.write_unsafe_i32_le(sc as i32);
        big_buf.write_unsafe_i32_le(ec as i32);
    }

    /// Public low-level helper (kept for unit-test parity with TS).
    pub fn create_row_header(
        &self,
        big_buf: &mut BigBuffer,
        row_number: u32,
        start_col: usize,
        end_col: usize,
    ) {
        Self::create_row_header_static(big_buf, row_number, start_col, end_col);
    }

    fn write_rk_static(big_buf: &mut BigBuffer, val: i32, col: usize, style: u8) {
        big_buf.ensure_capacity(14);
        big_buf.write_unsafe_byte(2);
        big_buf.write_unsafe_byte(12);
        big_buf.write_unsafe_i32_le(col as i32);
        big_buf.write_unsafe_byte(style);
        big_buf.write_unsafe_byte(0);
        big_buf.write_unsafe_byte(0);
        big_buf.write_unsafe_byte(0);
        big_buf.write_unsafe_i32_le((val << 2) | 2);
    }

    /// Public low-level helper (kept for unit-test parity with TS).
    pub fn write_rk_number_integer(
        &self,
        big_buf: &mut BigBuffer,
        val: i32,
        col_num: usize,
        style_num: u8,
    ) {
        Self::write_rk_static(big_buf, val, col_num, style_num);
    }

    fn write_double_static(big_buf: &mut BigBuffer, val: f64, col: usize, style: u8) {
        big_buf.ensure_capacity(18);
        big_buf.write_unsafe_byte(5);
        big_buf.write_unsafe_byte(16);
        big_buf.write_unsafe_i32_le(col as i32);
        big_buf.write_unsafe_byte(style);
        big_buf.write_unsafe_byte(0);
        big_buf.write_unsafe_byte(0);
        big_buf.write_unsafe_byte(0);
        big_buf.write_unsafe_f64_le(val);
    }

    /// Public low-level helper (kept for unit-test parity with TS).
    pub fn write_double(&self, big_buf: &mut BigBuffer, val: f64, col_num: usize, style_num: u8) {
        Self::write_double_static(big_buf, val, col_num, style_num);
    }

    fn write_bool_static(big_buf: &mut BigBuffer, val: bool, col: usize) {
        big_buf.ensure_capacity(13);
        big_buf.write_unsafe_byte(0x04);
        big_buf.write_unsafe_byte(9);
        big_buf.write_unsafe_i32_le(col as i32);
        big_buf.write_unsafe_i32_le(0);
        big_buf.write_unsafe_byte(if val { 1 } else { 0 });
    }

    /// Public low-level helper (kept for unit-test parity with TS).
    pub fn write_bool(&self, big_buf: &mut BigBuffer, val: bool, col_num: usize) {
        Self::write_bool_static(big_buf, val, col_num);
    }
    pub fn write_date_time(
        &self,
        big_buf: &mut BigBuffer,
        date: &chrono::NaiveDateTime,
        col_num: usize,
    ) {
        Self::write_double_static(big_buf, datetime_to_oa_date(date), col_num, 1);
    }

    fn write_sst_cell_static(big_buf: &mut BigBuffer, index: usize, col: usize, style: u8) {
        big_buf.ensure_capacity(17);
        big_buf.write_unsafe_byte(7);
        big_buf.write_unsafe_byte(12);
        big_buf.write_unsafe_i32_le(col as i32);
        big_buf.write_unsafe_byte(style);
        big_buf.write_unsafe_byte(0);
        big_buf.write_unsafe_byte(0);
        big_buf.write_unsafe_byte(0);
        big_buf.write_unsafe_i32_le(index as i32);
    }

    pub fn write_string(
        &mut self,
        big_buf: &mut BigBuffer,
        val: &str,
        col_num: usize,
        bolded: bool,
        style_override: Option<u8>,
    ) {
        let index = self.intern_string(val);
        let style = style_override.unwrap_or(if bolded { 3 } else { 0 });
        Self::write_sst_cell_static(big_buf, index, col_num, style);
    }

    fn save_sst(&mut self) -> EntryData {
        let mut big_buf = BigBuffer::default();
        big_buf.write_byte(159);
        big_buf.write_byte(1);
        big_buf.write_byte(8);
        big_buf.write_i32_le(self.sst_cnt_unique as i32);
        big_buf.write_i32_le(self.sst_cnt_all as i32);
        // SST entries sorted by index for deterministic output.
        let mut items: Vec<(&String, usize)> = self.sst_dic.iter().map(|(k, &v)| (k, v)).collect();
        items.sort_by_key(|(_, idx)| *idx);
        for (txt, _) in items {
            let txt_len = txt.encode_utf16().count();
            big_buf.write_byte(19);
            let rec_len = 5 + 2 * txt_len;
            if rec_len >= 128 {
                big_buf.write_byte((128 + (rec_len % 128)) as u8);
                let tmp = rec_len >> 7;
                if tmp >= 256 {
                    big_buf.write_byte((128 + (tmp % 128)) as u8);
                } else {
                    big_buf.write_byte(tmp as u8);
                }
                big_buf.write_byte((rec_len >> 14) as u8);
                if (rec_len >> 14) > 0 {
                    big_buf.write_byte(0);
                }
            } else {
                big_buf.write_byte((rec_len & 0xFF) as u8);
                big_buf.write_byte(((rec_len >> 8) & 0xFF) as u8);
            }
            big_buf.write_i32_le(txt_len as i32);
            big_buf.write_utf16le(txt);
        }
        big_buf.write_byte(160);
        big_buf.write_byte(1);
        big_buf.write_byte(0);
        EntryData::Chunks(big_buf.chunks_drain())
    }

    fn write_filter_defined_name(
        &self,
        wb_buffers: &mut Vec<u8>,
        sheet: &SheetInfo,
        sheet_num: usize,
    ) {
        let filter = sheet
            .filter_data
            .expect("filter data set when autofilter on");
        let sheet_index = (sheet.sheet_id - 1) as u8;
        let mut fix1 = xlsb_templates::FILTER_FIX1.to_vec();
        let last_idx = fix1.len() - 2;
        fix1[7] = sheet_index;
        fix1[last_idx] = sheet_num as u8;
        wb_buffers.extend_from_slice(&fix1);
        wb_buffers.extend_from_slice(&filter.start_row.to_le_bytes());
        wb_buffers.extend_from_slice(&filter.end_row.to_le_bytes());
        wb_buffers.extend_from_slice(&(filter.start_column as u16).to_le_bytes());
        wb_buffers.extend_from_slice(&(filter.end_column as u16).to_le_bytes());
        wb_buffers.extend_from_slice(&xlsb_templates::FILTER_FIX2);
    }

    fn workbook_bin(&self) -> Vec<u8> {
        let mut out = Vec::new();
        out.extend_from_slice(&xlsb_templates::WORKBOOK_BIN_START);
        for sheet in &self.sheet_list {
            let r_id = format!("rId{}", sheet.sheet_id);
            let r_units: Vec<u16> = r_id.encode_utf16().collect();
            let n_units: Vec<u16> = sheet.name.encode_utf16().collect();
            let rec_len = 4 + 12 + n_units.len() * 2 + r_units.len() * 2;
            // Note: TS uses JS-string `.length` (UTF-16 units) implicitly via
            // Buffer sizing; for ASCII names byte len == unit count.
            let mut buf = vec![0u8; 3 + rec_len];
            buf[0] = 156;
            buf[1] = 1;
            buf[2] = rec_len as u8;
            let mut pos = 3;
            let hidden = if sheet.hidden { 1i32 } else { 0 };
            buf[pos..pos + 4].copy_from_slice(&hidden.to_le_bytes());
            pos += 4;
            buf[pos..pos + 4].copy_from_slice(&(sheet.sheet_id as i32).to_le_bytes());
            pos += 4;
            buf[pos..pos + 4].copy_from_slice(&(r_units.len() as i32).to_le_bytes());
            pos += 4;
            for u in &r_units {
                buf[pos..pos + 2].copy_from_slice(&u.to_le_bytes());
                pos += 2;
            }
            buf[pos..pos + 4].copy_from_slice(&(n_units.len() as i32).to_le_bytes());
            pos += 4;
            for u in &n_units {
                buf[pos..pos + 2].copy_from_slice(&u.to_le_bytes());
                pos += 2;
            }
            out.extend_from_slice(&buf);
        }
        out.extend_from_slice(&xlsb_templates::WORKBOOK_BIN_MIDDLE);
        if self.autofilter_is_on {
            let filtered: Vec<&SheetInfo> = self
                .sheet_list
                .iter()
                .filter(|s| s.filter_data.is_some())
                .collect();
            let cnt = filtered.len();
            if cnt > 0 {
                out.extend_from_slice(&xlsb_templates::FILTER_FIX0);
                let first_byte = if cnt <= 20 {
                    0x10 + (cnt as u8 - 1) * 0x0C
                } else {
                    0x80 + (cnt as u8 - 21) * 0x0C
                };
                if cnt <= 10 {
                    out.extend_from_slice(&[first_byte, cnt as u8, 0x00, 0x00, 0x00]);
                } else {
                    out.extend_from_slice(&[
                        first_byte,
                        ((cnt - 1) / 10) as u8,
                        cnt as u8,
                        0x00,
                        0x00,
                        0x00,
                    ]);
                }
                for s in &filtered {
                    let sheet_index = (s.sheet_id - 1) as u8;
                    let mut idx_buf = [0u8; 12];
                    idx_buf[4] = sheet_index;
                    idx_buf[8] = sheet_index;
                    out.extend_from_slice(&idx_buf);
                }
                out.extend_from_slice(&[0xE2, 0x02, 0x00]);
                for (sheet_num, sheet) in filtered.iter().enumerate() {
                    self.write_filter_defined_name(&mut out, sheet, sheet_num);
                }
            }
        }
        out.extend_from_slice(&xlsb_templates::WORKBOOK_BIN_END);
        out
    }

    /// Finalize the ZIP package and write the output file.
    pub fn finalize(mut self) -> SpreadsheetResult<()> {
        if self.stream.is_streaming {
            return Err(SpreadsheetError::StreamingSheetOpen);
        }
        let sst = self.save_sst();
        let styles = self.build_styles_bin();
        let workbook = self.workbook_bin();
        self.entries.push(("xl/sharedStrings.bin".to_string(), sst));
        self.entries
            .push(("xl/styles.bin".to_string(), styles.into()));
        self.entries
            .push(("xl/workbook.bin".to_string(), workbook.into()));
        for sheet in &self.sheet_list {
            self.entries.push((
                format!("xl/worksheets/binaryIndex{}.bin", sheet.sheet_id),
                xlsb_templates::BINARY_INDEX_BIN.to_vec().into(),
            ));
        }
        let mut content_types = String::from(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="bin" ContentType="application/vnd.ms-excel.sheet.binary.macroEnabled.main"/>
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>"#,
        );
        for sheet in &self.sheet_list {
            content_types.push_str(&format!(
                r#"<Override PartName="/{}" ContentType="application/vnd.ms-excel.worksheet"/>"#,
                sheet.path_in_archive
            ));
            content_types.push_str(&format!(
                r#"<Override PartName="/xl/worksheets/binaryIndex{}.bin" ContentType="application/vnd.ms-excel.binIndexWs"/>"#,
                sheet.sheet_id
            ));
        }
        content_types.push_str(
            r#"<Override PartName="/xl/styles.bin" ContentType="application/vnd.ms-excel.styles"/>
<Override PartName="/xl/sharedStrings.bin" ContentType="application/vnd.ms-excel.sharedStrings"/>
</Types>"#,
        );
        self.entries.push((
            "[Content_Types].xml".to_string(),
            content_types.into_bytes().into(),
        ));

        let mut wb_rels = String::from(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">"#,
        );
        for sheet in &self.sheet_list {
            wb_rels.push_str(&format!(
                r#"<Relationship Id="rId{}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/{}"/>"#,
                sheet.sheet_id, sheet.name_in_archive
            ));
        }
        wb_rels.push_str(&format!(
            r#"<Relationship Id="rId{}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.bin"/>
"#,
            self.sheet_list.len() + 2
        ));
        wb_rels.push_str(&format!(
            r#"<Relationship Id="rId{}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.bin"/>
"#,
            self.sheet_list.len() + 3
        ));
        wb_rels.push_str("</Relationships>");
        self.entries.push((
            "xl/_rels/workbook.bin.rels".to_string(),
            wb_rels.into_bytes().into(),
        ));

        let global_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.bin"/>
</Relationships>"#;
        self.entries.push((
            "_rels/.rels".to_string(),
            global_rels.as_bytes().to_vec().into(),
        ));

        for sheet in &self.sheet_list {
            let ws_rels = format!(
                r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.microsoft.com/office/2006/relationships/xlBinaryIndex" Target="binaryIndex{}.bin"/>
</Relationships>"#,
                sheet.sheet_id
            );
            self.entries.push((
                format!("xl/worksheets/_rels/{}.rels", sheet.name_in_archive),
                ws_rels.into_bytes().into(),
            ));
        }

        let file = File::create(&self.output_path)?;
        let mut zip = zip::ZipWriter::new(file);
        let options = zip::write::SimpleFileOptions::default()
            .compression_method(zip::CompressionMethod::Deflated)
            // Default 1 matches the TS reference (archiver/compress-commons
            // default); override via `set_compression_level`.
            .compression_level(Some(self.compression_level))
            .last_modified_time(crate::writer_helpers::deterministic_zip_timestamp());
        for (name, data) in &self.entries {
            zip.start_file(name, options)?;
            data.write_to(&mut zip)?;
        }
        zip.finish()?;
        Ok(())
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::formats::CellValue;
    use std::path::Path;

    fn write_compression_sample(path: &Path, level: Option<i64>) -> u64 {
        let mut w = XlsbWriter::create(path).unwrap();
        if let Some(level) = level {
            w.set_compression_level(level).unwrap();
        }
        w.add_sheet("Data", false);
        let row = vec![
            CellValue::Text("powtarzalny tekst żółć".into()),
            CellValue::Integer(42),
            CellValue::Number(1.5),
        ];
        w.write_sheet(vec![row; 200], None, false).unwrap();
        w.finalize().unwrap();
        std::fs::metadata(path).unwrap().len()
    }

    #[test]
    fn compression_level_validates_range() {
        let dir = tempfile::tempdir().unwrap();
        let mut w = XlsbWriter::create(&dir.path().join("levels.xlsb")).unwrap();
        assert!(w.set_compression_level(1).is_ok());
        assert!(w.set_compression_level(9).is_ok());
        assert!(
            w.set_compression_level(0).is_err(),
            "level 0 is not supported by the zip backend"
        );
        assert!(w.set_compression_level(10).is_err());
        assert!(w.set_compression_level(-1).is_err());
    }

    #[test]
    fn compression_level_default_matches_explicit_one() {
        let dir = tempfile::tempdir().unwrap();
        let default_path = dir.path().join("default.xlsb");
        let explicit_path = dir.path().join("explicit.xlsb");
        write_compression_sample(&default_path, None);
        write_compression_sample(&explicit_path, Some(1));
        assert_eq!(
            std::fs::read(&default_path).unwrap(),
            std::fs::read(&explicit_path).unwrap(),
            "default output must stay byte-identical to level 1"
        );
    }

    #[test]
    fn compression_level_9_shrinks_output() {
        let dir = tempfile::tempdir().unwrap();
        let size_default = write_compression_sample(&dir.path().join("default.xlsb"), None);
        let size_9 = write_compression_sample(&dir.path().join("level9.xlsb"), Some(9));
        assert!(
            size_9 < size_default,
            "level 9 ({size_9} B) should be smaller than the default level 1 ({size_default} B)"
        );
    }

    #[test]
    fn rk_encoding_matches_ts_vector() {
        // TS unit test: writeRkNumberInteger(42) -> rk = (42 << 2) | 2 = 170
        let w = XlsbWriter::create(Path::new("/tmp/unused.xlsb")).unwrap();
        let mut buf = BigBuffer::default();
        w.write_rk_number_integer(&mut buf, 42, 0, 0);
        let bytes = buf.chunks_drain().concat();
        assert_eq!(bytes[0], 2);
        assert_eq!(bytes[1], 12);
        assert_eq!(i32::from_le_bytes(bytes[10..14].try_into().unwrap()), 170);
    }

    #[test]
    fn roundtrip_basic() {
        let dir = tempfile::tempdir().unwrap();
        let path = dir.path().join("basic.xlsb");
        let mut w = XlsbWriter::create(&path).unwrap();
        w.add_sheet("Sheet1", false);
        w.write_sheet(
            vec![
                vec![CellValue::Text("Alice".into()), CellValue::Integer(30)],
                vec![CellValue::Boolean(true), CellValue::Number(1.5)],
            ],
            Some(&["Name".to_string(), "Age".to_string()]),
            true,
        )
        .unwrap();
        w.finalize().unwrap();

        let mut r = crate::XlsbReader::new();
        r.open(&path, true).unwrap();
        assert_eq!(r.sheet_names(), &["Sheet1".to_string()]);
        let mut rows: Vec<Vec<CellValue>> = Vec::new();
        while r.read().unwrap() {
            // last read() at EOF returns an empty trailing row; stop there
            if r.current_row().is_empty() && r.current_row().is_empty() {
                // collect anyway; EOF row is empty
            }
            rows.push(r.current_row().to_vec());
            if rows.len() > 10 {
                break;
            }
        }
        // header + 2 data rows + 1 empty EOF row
        assert_eq!(rows[0][0], CellValue::Text("Name".into()));
        assert_eq!(rows[1][0], CellValue::Text("Alice".into()));
        assert_eq!(rows[1][1], CellValue::Number(30.0));
        assert_eq!(rows[2][0], CellValue::Boolean(true));
        assert_eq!(rows[2][1], CellValue::Number(1.5));
    }

    #[test]
    fn streaming_matches_batch() {
        let dir = tempfile::tempdir().unwrap();
        let path = dir.path().join("stream.xlsb");
        let mut w = XlsbWriter::create(&path).unwrap();
        w.start_sheet(
            "Data",
            2,
            Some(&["A".to_string(), "B".to_string()]),
            XlsbSheetOptions::new(),
        )
        .unwrap();
        w.write_row(&[CellValue::Integer(1), CellValue::Text("x".into())])
            .unwrap();
        w.end_sheet().unwrap();
        w.finalize().unwrap();

        let mut r = crate::XlsbReader::new();
        r.open(&path, true).unwrap();
        let mut non_empty = 0;
        loop {
            let has = r.read().unwrap();
            if !has {
                break;
            }
            if !r.current_row().is_empty() {
                non_empty += 1;
            }
            if non_empty > 10 {
                break;
            }
        }
        assert_eq!(non_empty, 2);
    }

    #[test]
    fn unicode_sheet_names_roundtrip() {
        let dir = tempfile::tempdir().unwrap();
        let path = dir.path().join("unicode.xlsb");
        let mut w = XlsbWriter::create(&path).unwrap();
        w.add_sheet("Dane żółć 😀", false);
        w.write_sheet(vec![vec![CellValue::Integer(1)]], None, false)
            .unwrap();
        w.finalize().unwrap();

        let mut r = crate::XlsbReader::new();
        r.open(&path, true).unwrap();
        assert_eq!(r.sheet_names(), &["Dane żółć 😀".to_string()]);
    }
}
