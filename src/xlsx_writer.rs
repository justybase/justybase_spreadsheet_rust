//! XLSX (Office Open XML) writer (port of `XlsxWriter.ts`).

use crate::big_buffer::{BigBuffer, EntryData};
use crate::error::{SpreadsheetError, SpreadsheetResult};
use crate::formats::{borrow_cell, get_format, CellValue, CellValueRef};
use crate::streaming_state::StreamingSheetState;
use crate::writer_helpers::{
    apply_header_widths, default_col_width, init_col_widths, unique_sheet_name,
    update_col_widths_from_rows,
};
use crate::{datetime_to_oa_date, xml_utils};
use std::collections::HashMap;
use std::fmt::Write as FmtWrite;
use std::fs::File;
use std::path::{Path, PathBuf};

struct SheetInfo {
    name: String,
    path_in_archive: String,
    hidden: bool,
    name_in_archive: String,
    sheet_id: usize,
    r_id: String,
    filter_header_range: Option<String>,
}

/// Options for [`XlsxWriter::start_sheet`].
#[derive(Debug, Clone, Default)]
pub struct SheetOptions {
    pub hidden: bool,
    pub do_autofilter: bool,
    pub sample_rows: Option<Vec<Vec<CellValue>>>,
}

impl SheetOptions {
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
    Number(f64, Option<usize>),
    Integer(i64, Option<usize>),
    Bool(bool, Option<usize>),
    DateTime(f64, usize),       // oa, style
    Text(usize, Option<usize>), // sst idx, style
    NonFiniteSst(usize, Option<usize>),
}

pub struct XlsxWriter {
    output_path: PathBuf,
    entries: Vec<(String, EntryData)>,

    sheet_count: usize,
    sheet_list: Vec<SheetInfo>,
    sst_array: Vec<String>,
    sst_map: HashMap<String, usize>,
    sst_cnt_all: u64,
    autofilter_is_on: bool,

    format_registry: HashMap<String, u32>,
    format_xf_map: HashMap<String, usize>,
    next_numfmt_id: u32,
    next_xf_index: usize,
    compression_level: i64,

    stream: StreamingSheetState,
    current_col_letters: Vec<String>,
    staged_row: Vec<Option<StreamStaged>>,
    /// Sheet XML from `<sheetViews>` through the header row, staged until
    /// `end_sheet` so the `<dimension>` ref can carry the final row count.
    pending_sheet_head: String,
}

impl XlsxWriter {
    pub fn create(path: &Path) -> SpreadsheetResult<Self> {
        if let Some(parent) = path.parent() {
            if !parent.as_os_str().is_empty() {
                std::fs::create_dir_all(parent)?;
            }
        }
        Ok(Self {
            output_path: path.to_path_buf(),
            entries: Vec::new(),
            sheet_count: 0,
            sheet_list: Vec::new(),
            sst_array: Vec::new(),
            sst_map: HashMap::new(),
            sst_cnt_all: 0,
            autofilter_is_on: false,
            format_registry: HashMap::new(),
            format_xf_map: HashMap::new(),
            next_numfmt_id: 165,
            next_xf_index: 4,
            compression_level: 1,
            stream: StreamingSheetState::new(),
            current_col_letters: Vec::new(),
            staged_row: Vec::new(),
            pending_sheet_head: String::new(),
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

    fn get_column_letter(col: usize) -> String {
        xml_utils::column_index_to_letter(col)
    }

    /// Register a worksheet (metadata only).
    pub fn add_sheet(&mut self, sheet_name: &str, hidden: bool) {
        let sanitized = unique_sheet_name(sheet_name, self.sheet_count, |candidate| {
            self.sheet_list
                .iter()
                .any(|sheet| sheet.name.eq_ignore_ascii_case(candidate))
        });
        self.sheet_count += 1;
        let r_id = format!("rId{}", self.sheet_count);
        let file_name = format!("sheet{}.xml", self.sheet_count);
        self.sheet_list.push(SheetInfo {
            name: sanitized,
            path_in_archive: format!("xl/worksheets/{file_name}"),
            hidden,
            name_in_archive: file_name,
            sheet_id: self.sheet_count,
            r_id,
            filter_header_range: None,
        });
    }

    fn needs_escape(s: &str) -> bool {
        for c in s.chars() {
            let v = c as u32;
            if v == 38
                || v == 60
                || v == 62
                || v == 34
                || v == 39
                || (v <= 8)
                || v == 11
                || v == 12
                || (14..=31).contains(&v)
            {
                return true;
            }
        }
        false
    }

    fn escape(s: &str) -> String {
        if !Self::needs_escape(s) {
            return s.to_owned();
        }
        let mut out = s
            .replace('&', "&amp;")
            .replace('<', "&lt;")
            .replace('>', "&gt;")
            .replace('"', "&quot;")
            .replace('\'', "&apos;");
        out = out
            .chars()
            .filter(|&c| {
                let v = c as u32;
                !(v <= 8 || v == 11 || v == 12 || (14..=31).contains(&v))
            })
            .collect();
        out
    }

    fn escape_sheet_name_xml(name: &str) -> String {
        name.replace('&', "&amp;")
            .replace('<', "&lt;")
            .replace('>', "&gt;")
            .replace('"', "&quot;")
            .replace('\'', "&apos;")
    }

    fn format_sheet_name_for_formula(name: &str) -> String {
        let needs_quote = name.chars().any(|c| " \t-+=()!@#$%^&'".contains(c))
            || name
                .chars()
                .next()
                .map(|c| c.is_ascii_digit())
                .unwrap_or(false);
        if needs_quote {
            format!("'{}'", name.replace('\'', "''"))
        } else {
            name.to_owned()
        }
    }

    /// Start streaming a sheet. Call [`XlsxWriter::write_row`] per row, then
    /// [`XlsxWriter::end_sheet`].
    pub fn start_sheet(
        &mut self,
        sheet_name: &str,
        column_count: usize,
        headers: Option<&[String]>,
        options: SheetOptions,
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
        let col_letters: Vec<String> = (0..column_count).map(Self::get_column_letter).collect();
        self.add_sheet(sheet_name, options.hidden);
        self.current_col_letters = col_letters;
        if let Some(h) = headers {
            apply_header_widths(&mut self.stream.col_widths, h, column_count);
        }
        if let Some(sample) = &options.sample_rows {
            update_col_widths_from_rows(&mut self.stream.col_widths, sample, 100);
        }
        // Snapshot everything the buffer-writing phase needs so no borrow
        // of `self.stream` is held while `big_buf` is alive.
        let widths = self.stream.col_widths.clone();
        let is_first = self.sheet_count == 1;
        let has_headers = headers.is_some();
        let header_cells: Vec<usize> = match headers {
            Some(h) => h.iter().map(|s| self.intern_string(s)).collect(),
            None => Vec::new(),
        };
        // Everything from <worksheet> to the header row is fixed at this
        // point except <dimension>, which needs the final row count;
        // `end_sheet` joins the parts so the dimension ref is correct
        // (read-only consumers such as openpyxl and pandas size the sheet
        // from that element).
        let mut head = String::with_capacity(512);
        if has_headers {
            if is_first {
                head.push_str(r#"<sheetViews><sheetView tabSelected="1" workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen" /><selection pane="bottomLeft" /></sheetView></sheetViews><sheetFormatPr defaultRowHeight="15"/>"#);
            } else {
                head.push_str(r#"<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen" /><selection pane="bottomLeft" /></sheetView></sheetViews><sheetFormatPr defaultRowHeight="15"/>"#);
            }
        } else if is_first {
            head.push_str(r#"<sheetViews><sheetView tabSelected="1" workbookViewId="0"/></sheetViews><sheetFormatPr defaultRowHeight="15"/>"#);
        } else {
            head.push_str(r#"<sheetViews><sheetView workbookViewId="0"/></sheetViews><sheetFormatPr defaultRowHeight="15"/>"#);
        }
        head.push_str("<cols>");
        for (i, width_value) in widths.iter().enumerate().take(column_count) {
            let width = default_col_width(*width_value, false);
            head.push_str(&format!(
                r#"<col min="{}" max="{}" width="{width}" bestFit="1" customWidth="1" />"#,
                i + 1,
                i + 1
            ));
        }
        head.push_str("</cols><sheetData>");
        if has_headers {
            // Dense mode (like SpreadSheetTasks C#): omit r on <row> and <c>
            head.push_str(r#"<row>"#);
            for idx in header_cells.iter() {
                head.push_str(&format!(r#"<c t="s" s="3"><v>{idx}</v></c>"#));
            }
            head.push_str("</row>");
        }
        self.pending_sheet_head = head;
        if has_headers {
            self.stream.row_num = 1;
        }
        Ok(())
    }

    /// Write one data row in streaming mode.
    pub fn write_row(&mut self, row: &[CellValue]) -> SpreadsheetResult<()> {
        let expected = self.stream.end_col - self.stream.start_col;
        if row.len() != expected {
            return Err(SpreadsheetError::RowLengthMismatch {
                expected,
                got: row.len(),
            });
        }
        // Dense: omit r, write directly to BigBuffer to avoid per-cell allocs
        let row_num = self.stream.row_num + 1;
        // Stage sst indices that need interning, then stream.
        // We do interning first (needs &mut self) then borrow big_buf.
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
            let xf_idx: Option<usize> = fmt.map(|f| self.register_format(f));
            let cell = match val {
                CellValueRef::Empty => None,
                CellValueRef::Number(n) => {
                    if n.is_finite() {
                        Some(StreamStaged::Number(n, xf_idx))
                    } else {
                        let idx = self.intern_string(&n.to_string());
                        Some(StreamStaged::NonFiniteSst(idx, xf_idx))
                    }
                }
                CellValueRef::Integer(n) => Some(StreamStaged::Integer(n, xf_idx)),
                CellValueRef::Boolean(b) => Some(StreamStaged::Bool(b, xf_idx)),
                CellValueRef::DateTime(d) => {
                    let oa = datetime_to_oa_date(&d);
                    if oa.is_finite() {
                        let style = xf_idx.unwrap_or(1);
                        Some(StreamStaged::DateTime(oa, style))
                    } else {
                        let idx = self.intern_string(&d.to_string());
                        Some(StreamStaged::NonFiniteSst(idx, xf_idx))
                    }
                }
                CellValueRef::Text(s) => {
                    let idx = self.intern_string(s);
                    Some(StreamStaged::Text(idx, xf_idx))
                }
            };
            staged.push(cell);
        }
        let big_buf = self.stream.assert_streaming()?;
        big_buf.write_str(r#"<row>"#);
        for cell in staged.drain(..).flatten() {
            match cell {
                StreamStaged::Number(n, style) => {
                    if let Some(s) = style {
                        let _ = write!(big_buf, r#"<c s="{s}"><v>{n}</v></c>"#);
                    } else {
                        let _ = write!(big_buf, r#"<c><v>{n}</v></c>"#);
                    }
                }
                StreamStaged::Integer(n, style) => {
                    if let Some(s) = style {
                        let _ = write!(big_buf, r#"<c s="{s}"><v>{n}</v></c>"#);
                    } else {
                        let _ = write!(big_buf, r#"<c><v>{n}</v></c>"#);
                    }
                }
                StreamStaged::Bool(b, style) => {
                    let v = if b { 1 } else { 0 };
                    if let Some(s) = style {
                        let _ = write!(big_buf, r#"<c t="b" s="{s}"><v>{v}</v></c>"#);
                    } else {
                        let _ = write!(big_buf, r#"<c t="b"><v>{v}</v></c>"#);
                    }
                }
                StreamStaged::DateTime(oa, style) => {
                    let _ = write!(big_buf, r#"<c s="{style}"><v>{oa}</v></c>"#);
                }
                StreamStaged::NonFiniteSst(idx, style) => {
                    if let Some(s) = style {
                        let _ = write!(big_buf, r#"<c t="s" s="{s}"><v>{idx}</v></c>"#);
                    } else {
                        let _ = write!(big_buf, r#"<c t="s"><v>{idx}</v></c>"#);
                    }
                }
                StreamStaged::Text(idx, style) => {
                    if let Some(s) = style {
                        let _ = write!(big_buf, r#"<c t="s" s="{s}"><v>{idx}</v></c>"#);
                    } else {
                        let _ = write!(big_buf, r#"<c t="s"><v>{idx}</v></c>"#);
                    }
                }
            }
        }
        big_buf.write_str("</row>");
        self.stream.row_num = row_num;
        self.staged_row = staged;
        Ok(())
    }

    /// Finish the current streaming sheet and stage it for the archive.
    pub fn end_sheet(&mut self) -> SpreadsheetResult<()> {
        if !self.stream.is_streaming {
            return Err(crate::error::SpreadsheetError::NotStreaming);
        }
        let do_filter = self.stream.do_autofilter;
        let end_col = self.stream.end_col;
        let row_num = self.stream.row_num;
        let cols = self.current_col_letters.clone();
        let sheet_idx = self.sheet_count - 1;
        let sheet_name = self.sheet_list[sheet_idx].name.clone();
        {
            let big_buf = self.stream.assert_streaming()?;
            big_buf.write_str("</sheetData>");
            if do_filter && end_col > 0 && row_num > 0 {
                let last_col = cols[end_col - 1].clone();
                big_buf.write_str(&format!(r#"<autoFilter ref="A1:{last_col}{row_num}"/>"#));
            }
            big_buf.write_str("</worksheet>");
        }
        if do_filter && end_col > 0 && row_num > 0 {
            self.autofilter_is_on = true;
            let last_col = cols[end_col - 1].clone();
            let formula = Self::format_sheet_name_for_formula(&sheet_name);
            self.sheet_list[sheet_idx].filter_header_range =
                Some(format!("{formula}!$A$1:${last_col}${row_num}"));
        }
        // Sheet head with the final <dimension> (same rule as write_sheet),
        // prepended to the buffered rows as its own chunk.
        let mut head = String::with_capacity(self.pending_sheet_head.len() + 128);
        head.push_str(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#);
        head.push_str(r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">"#);
        if row_num > 0 && end_col > 0 {
            head.push_str(&format!(
                r#"<dimension ref="A1:{}{row_num}"/>"#,
                cols[end_col - 1]
            ));
        } else {
            head.push_str(r#"<dimension ref="A1"/>"#);
        }
        head.push_str(&std::mem::take(&mut self.pending_sheet_head));
        // Drain the buffer now that no other borrow is alive.
        let data = EntryData::Chunks({
            let big_buf = self.stream.assert_streaming()?;
            let mut chunks = big_buf.chunks_drain();
            let mut all = Vec::with_capacity(chunks.len() + 1);
            all.push(head.into_bytes());
            all.append(&mut chunks);
            all
        });
        let path = self.sheet_list[sheet_idx].path_in_archive.clone();
        self.entries.push((path, data));
        self.stream.end();
        self.current_col_letters.clear();
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
        let col_letters: Vec<String> = (0..column_count).map(Self::get_column_letter).collect();
        if let Some(h) = headers {
            apply_header_widths(&mut col_widths, h, column_count);
        }
        update_col_widths_from_rows(&mut col_widths, &rows, 100);

        // Reserve SST for ~2 text cells per row + headers (benchmark shape)
        let expected_sst = headers.map(|h| h.len()).unwrap_or(0) + rows.len() * 2 + 16;
        self.sst_array.reserve(expected_sst);
        self.sst_map.reserve(expected_sst);
        // Pre-allocate: ~30 bytes per cell + overhead, like C# dense
        let est = (rows.len() * column_count.saturating_mul(32) + 8192).max(8192);
        let mut xml = String::with_capacity(est);
        xml.push_str(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#);
        xml.push_str(r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">"#);
        let total_rows = rows.len() + headers.map(|_| 1).unwrap_or(0);
        if total_rows > 0 && column_count > 0 {
            xml.push_str(&format!(
                r#"<dimension ref="A1:{}{total_rows}"/>"#,
                col_letters[column_count - 1]
            ));
        } else {
            xml.push_str(r#"<dimension ref="A1"/>"#);
        }
        let is_first = self.sheet_count == 1;
        if headers.is_some() {
            if is_first {
                xml.push_str(r#"<sheetViews><sheetView tabSelected="1" workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen" /><selection pane="bottomLeft" /></sheetView></sheetViews><sheetFormatPr defaultRowHeight="15"/>"#);
            } else {
                xml.push_str(r#"<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen" /><selection pane="bottomLeft" /></sheetView></sheetViews><sheetFormatPr defaultRowHeight="15"/>"#);
            }
        } else if is_first {
            xml.push_str(r#"<sheetViews><sheetView tabSelected="1" workbookViewId="0"/></sheetViews><sheetFormatPr defaultRowHeight="15"/>"#);
        } else {
            xml.push_str(r#"<sheetViews><sheetView workbookViewId="0"/></sheetViews><sheetFormatPr defaultRowHeight="15"/>"#);
        }
        xml.push_str("<cols>");
        for (i, width_value) in col_widths.iter().enumerate().take(column_count) {
            let width = default_col_width(*width_value, false);
            xml.push_str(&format!(
                r#"<col min="{}" max="{}" width="{width}" bestFit="1" customWidth="1" />"#,
                i + 1,
                i + 1
            ));
        }
        xml.push_str("</cols><sheetData>");
        // Reuse buffers for number formatting to avoid per-cell alloc
        let mut itoa_buf = itoa::Buffer::new();
        let mut ryu_buf = ryu::Buffer::new();
        if let Some(h) = headers {
            xml.push_str(r#"<row>"#);
            for header in h.iter() {
                let idx = self.intern_string(header);
                xml.push_str(r#"<c t="s" s="3"><v>"#);
                xml.push_str(itoa_buf.format(idx));
                xml.push_str("</v></c>");
            }
            xml.push_str("</row>");
        }
        for row in &rows {
            xml.push_str(r#"<row>"#);
            for raw in row.iter() {
                if matches!(raw, CellValue::Empty) {
                    continue;
                }
                let fmt = get_format(raw);
                let val = borrow_cell(raw);
                let xf_idx: Option<usize> = fmt.map(|f| self.register_format(f));
                match val {
                    CellValueRef::Empty => {}
                    CellValueRef::Number(n) => {
                        if n.is_finite() {
                            if let Some(s) = xf_idx {
                                xml.push_str(r#"<c s=""#);
                                xml.push_str(itoa_buf.format(s));
                                xml.push_str(r#""><v>"#);
                                xml.push_str(ryu_buf.format(n));
                                xml.push_str("</v></c>");
                            } else {
                                xml.push_str("<c><v>");
                                xml.push_str(ryu_buf.format(n));
                                xml.push_str("</v></c>");
                            }
                        } else {
                            let idx = self.intern_string(&n.to_string());
                            if let Some(s) = xf_idx {
                                xml.push_str(r#"<c t="s" s=""#);
                                xml.push_str(itoa_buf.format(s));
                                xml.push_str(r#""><v>"#);
                                xml.push_str(itoa_buf.format(idx));
                                xml.push_str("</v></c>");
                            } else {
                                xml.push_str(r#"<c t="s"><v>"#);
                                xml.push_str(itoa_buf.format(idx));
                                xml.push_str("</v></c>");
                            }
                        }
                    }
                    CellValueRef::Integer(n) => {
                        if let Some(s) = xf_idx {
                            xml.push_str(r#"<c s=""#);
                            xml.push_str(itoa_buf.format(s));
                            xml.push_str(r#""><v>"#);
                            xml.push_str(itoa_buf.format(n));
                            xml.push_str("</v></c>");
                        } else {
                            xml.push_str("<c><v>");
                            xml.push_str(itoa_buf.format(n));
                            xml.push_str("</v></c>");
                        }
                    }
                    CellValueRef::Boolean(b) => {
                        let v = if b { 1 } else { 0 };
                        if let Some(s) = xf_idx {
                            xml.push_str(r#"<c t="b" s=""#);
                            xml.push_str(itoa_buf.format(s));
                            xml.push_str(r#""><v>"#);
                            xml.push_str(itoa_buf.format(v));
                            xml.push_str("</v></c>");
                        } else {
                            xml.push_str(r#"<c t="b"><v>"#);
                            xml.push_str(itoa_buf.format(v));
                            xml.push_str("</v></c>");
                        }
                    }
                    CellValueRef::DateTime(d) => {
                        let oa = datetime_to_oa_date(&d);
                        if oa.is_finite() {
                            let style = xf_idx.unwrap_or(1);
                            xml.push_str(r#"<c s=""#);
                            xml.push_str(itoa_buf.format(style));
                            xml.push_str(r#""><v>"#);
                            xml.push_str(ryu_buf.format(oa));
                            xml.push_str("</v></c>");
                        } else {
                            let idx = self.intern_string(&d.to_string());
                            if let Some(s) = xf_idx {
                                xml.push_str(r#"<c t="s" s=""#);
                                xml.push_str(itoa_buf.format(s));
                                xml.push_str(r#""><v>"#);
                                xml.push_str(itoa_buf.format(idx));
                                xml.push_str("</v></c>");
                            } else {
                                xml.push_str(r#"<c t="s"><v>"#);
                                xml.push_str(itoa_buf.format(idx));
                                xml.push_str("</v></c>");
                            }
                        }
                    }
                    CellValueRef::Text(s) => {
                        let idx = self.intern_string(s);
                        if let Some(st) = xf_idx {
                            xml.push_str(r#"<c t="s" s=""#);
                            xml.push_str(itoa_buf.format(st));
                            xml.push_str(r#""><v>"#);
                            xml.push_str(itoa_buf.format(idx));
                            xml.push_str("</v></c>");
                        } else {
                            xml.push_str(r#"<c t="s"><v>"#);
                            xml.push_str(itoa_buf.format(idx));
                            xml.push_str("</v></c>");
                        }
                    }
                }
            }
            xml.push_str("</row>");
        }
        xml.push_str("</sheetData>");
        if do_autofilter && headers.is_some() && column_count > 0 && total_rows > 0 {
            self.autofilter_is_on = true;
            let last = &col_letters[column_count - 1];
            xml.push_str(&format!(r#"<autoFilter ref="A1:{last}{total_rows}"/>"#));
            let sheet = &mut self.sheet_list[self.sheet_count - 1];
            let formula = Self::format_sheet_name_for_formula(&sheet.name);
            sheet.filter_header_range = Some(format!("{formula}!$A$1:${last}${total_rows}"));
        }
        xml.push_str("</worksheet>");
        let path = self.sheet_list[self.sheet_count - 1]
            .path_in_archive
            .clone();
        self.entries.push((path, xml.into_bytes().into()));
        Ok(())
    }

    fn intern_string(&mut self, val: &str) -> usize {
        if let Some(&idx) = self.sst_map.get(val) {
            self.sst_cnt_all += 1;
            return idx;
        }
        let idx = self.sst_array.len();
        let owned = val.to_owned();
        self.sst_array.push(owned.clone());
        self.sst_map.insert(owned, idx);
        self.sst_cnt_all += 1;
        idx
    }

    #[allow(dead_code)]
    fn write_string_cell_inner(
        &mut self,
        big_buf: &mut BigBuffer,
        val: &str,
        col_ref: &str,
        row_num: u32,
        style_override: Option<usize>,
    ) {
        let index = self.intern_string(val);
        let style_attr = style_override
            .map(|s| format!(r#" s="{s}""#))
            .unwrap_or_default();
        big_buf.write_str(&format!(
            r#"<c r="{col_ref}{row_num}" t="s"{style_attr}><v>{index}</v></c>"#
        ));
    }

    fn register_format(&mut self, fmt: &str) -> usize {
        if let Some(&xf) = self.format_xf_map.get(fmt) {
            return xf;
        }
        let numfmt_id = self.next_numfmt_id;
        self.next_numfmt_id += 1;
        self.format_registry.insert(fmt.to_owned(), numfmt_id);
        let xf = self.next_xf_index;
        self.next_xf_index += 1;
        self.format_xf_map.insert(fmt.to_owned(), xf);
        xf
    }

    fn shared_strings_xml(&self) -> Vec<u8> {
        // Reserve up front: the SST is the largest XML entry and doubling
        // from `String::new()` costs ~2x its final size in churn.
        let mut xml = String::with_capacity(
            self.sst_array
                .iter()
                .fold(256, |bytes, text| bytes + text.len() + 40),
        );
        xml.push_str(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#);
        xml.push_str(&format!(
            r#"<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="{}" uniqueCount="{}">"#,
            self.sst_cnt_all,
            self.sst_array.len()
        ));
        for txt in &self.sst_array {
            // Borrow the text when no escaping is needed: this skips one
            // allocation per unique string for plain data.
            let clean = if Self::needs_escape(txt) {
                std::borrow::Cow::Owned(Self::escape(txt))
            } else {
                std::borrow::Cow::Borrowed(txt.as_str())
            };
            if !clean.is_empty()
                && (clean.starts_with(' ')
                    || clean.ends_with(' ')
                    || clean.contains(['\t', '\n', '\r']))
            {
                xml.push_str(r#"<si><t xml:space="preserve">"#);
                xml.push_str(&clean);
                xml.push_str("</t></si>");
            } else {
                xml.push_str("<si><t>");
                xml.push_str(&clean);
                xml.push_str("</t></si>");
            }
        }
        xml.push_str("</sst>");
        xml.into_bytes()
    }

    fn styles_xml(&self) -> Vec<u8> {
        let mut num_fmts = String::new();
        let mut num_fmt_count = 0;
        let mut xf_entries = vec![
            r#"<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>"#.to_string(),
            r#"<xf numFmtId="14" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>"#
                .to_string(),
            r#"<xf numFmtId="22" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>"#
                .to_string(),
            r#"<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>"#
                .to_string(),
        ];
        // Deterministic order: sort by numfmt id (HashMap order is random).
        let mut reg: Vec<(&String, &u32)> = self.format_registry.iter().collect();
        reg.sort_by_key(|(_, id)| **id);
        for (fmt, numfmt_id) in reg {
            let escaped = xml_utils::escape_xml_text(fmt);
            num_fmts.push_str(&format!(
                r#"<numFmt numFmtId="{numfmt_id}" formatCode="{escaped}"/>"#
            ));
            num_fmt_count += 1;
        }
        let mut sorted: Vec<(&String, &usize)> = self.format_xf_map.iter().collect();
        sorted.sort_by_key(|(_, xf)| **xf);
        for (fmt, _) in sorted {
            if let Some(numfmt_id) = self.format_registry.get(fmt) {
                xf_entries.push(format!(
                    r#"<xf numFmtId="{numfmt_id}" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>"#
                ));
            }
        }
        let total_xf = xf_entries.len();
        let num_fmts_section = if num_fmt_count > 0 {
            format!(r#"<numFmts count="{num_fmt_count}">{num_fmts}</numFmts>"#)
        } else {
            String::new()
        };
        format!(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
{num_fmts_section}
<fonts count="2">
<font><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font>
<font><b/><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font>
</fonts>
<fills count="2">
<fill><patternFill patternType="none"/></fill>
<fill><patternFill patternType="gray125"/></fill>
</fills>
<borders count="1">
<border><left/><right/><top/><bottom/><diagonal/></border>
</borders>
<cellStyleXfs count="1">
<xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
</cellStyleXfs>
<cellXfs count="{total_xf}">
{}
</cellXfs>
<cellStyles count="1">
<cellStyle name="Normal" xfId="0" builtinId="0"/>
</cellStyles>
<dxfs count="0"/>
<tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"/>
</styleSheet>"#,
            xf_entries.join("\n")
        )
        .into_bytes()
    }

    fn workbook_xml(&self) -> Vec<u8> {
        let mut xml = String::from(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<fileVersion appName="xl" lastEdited="4" lowestEdited="4" rupBuild="4505"/>
<workbookPr defaultThemeVersion="124226"/>
<bookViews><workbookView xWindow="240" yWindow="15" windowWidth="16095" windowHeight="9660"/></bookViews>
<sheets>"#,
        );
        for sheet in &self.sheet_list {
            let state = if sheet.hidden {
                r#" state="hidden""#
            } else {
                ""
            };
            let escaped = Self::escape_sheet_name_xml(&sheet.name);
            xml.push_str(&format!(
                r#"<sheet name="{escaped}" sheetId="{}"{state} r:id="{}"/>"#,
                sheet.sheet_id, sheet.r_id
            ));
        }
        xml.push_str("</sheets>");
        if self.autofilter_is_on {
            xml.push_str("<definedNames>");
            for sheet in &self.sheet_list {
                if let Some(range) = &sheet.filter_header_range {
                    let local = sheet.sheet_id - 1;
                    let escaped = Self::escape(range);
                    xml.push_str(&format!(
                        r#"<definedName name="_xlnm._FilterDatabase" localSheetId="{local}" hidden="1">{escaped}</definedName>"#
                    ));
                }
            }
            xml.push_str("</definedNames>");
        }
        xml.push_str(r#"<calcPr calcId="124519" fullCalcOnLoad="1"/>"#);
        xml.push_str("</workbook>");
        xml.into_bytes()
    }

    fn content_types_xml(&self) -> Vec<u8> {
        let mut xml = String::from(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>"#,
        );
        for sheet in &self.sheet_list {
            xml.push_str(&format!(
                r#"<Override PartName="/{}" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>"#,
                sheet.path_in_archive
            ));
        }
        xml.push_str("</Types>");
        xml.into_bytes()
    }

    fn rels_entries(&self) -> Vec<(String, EntryData)> {
        let global = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>"#;
        let mut wb_rels = String::from(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">"#,
        );
        for sheet in &self.sheet_list {
            wb_rels.push_str(&format!(
                r#"<Relationship Id="{}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/{}"/>"#,
                sheet.r_id, sheet.name_in_archive
            ));
        }
        let mut next_id = self.sheet_list.len() + 1;
        wb_rels.push_str(&format!(
            r#"<Relationship Id="rId{next_id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>"#
        ));
        next_id += 1;
        wb_rels.push_str(&format!(
            r#"<Relationship Id="rId{next_id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>"#
        ));
        wb_rels.push_str("</Relationships>");
        vec![
            ("_rels/.rels".to_string(), global.as_bytes().to_vec().into()),
            (
                "xl/_rels/workbook.xml.rels".to_string(),
                wb_rels.into_bytes().into(),
            ),
        ]
    }

    /// Finalize the ZIP package and write the output file.
    pub fn finalize(mut self) -> SpreadsheetResult<()> {
        if self.stream.is_streaming {
            return Err(SpreadsheetError::StreamingSheetOpen);
        }
        let sst = self.shared_strings_xml();
        let styles = self.styles_xml();
        let workbook = self.workbook_xml();
        let content_types = self.content_types_xml();
        let rels = self.rels_entries();
        self.entries
            .push(("xl/sharedStrings.xml".to_string(), sst.into()));
        self.entries
            .push(("xl/styles.xml".to_string(), styles.into()));
        self.entries
            .push(("xl/workbook.xml".to_string(), workbook.into()));
        self.entries
            .push(("[Content_Types].xml".to_string(), content_types.into()));
        self.entries.extend(rels);

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

    fn write_compression_sample(path: &Path, level: Option<i64>) -> u64 {
        let mut w = XlsxWriter::create(path).unwrap();
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
        let mut w = XlsxWriter::create(&dir.path().join("levels.xlsx")).unwrap();
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
        let default_path = dir.path().join("default.xlsx");
        let explicit_path = dir.path().join("explicit.xlsx");
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
        let size_default = write_compression_sample(&dir.path().join("default.xlsx"), None);
        let size_9 = write_compression_sample(&dir.path().join("level9.xlsx"), Some(9));
        assert!(
            size_9 < size_default,
            "level 9 ({size_9} B) should be smaller than the default level 1 ({size_default} B)"
        );
    }

    #[test]
    fn roundtrip_basic() {
        let dir = tempfile::tempdir().unwrap();
        let path = dir.path().join("basic.xlsx");
        let mut w = XlsxWriter::create(&path).unwrap();
        w.add_sheet("Sheet1", false);
        w.write_sheet(
            vec![
                vec![CellValue::Text("Name".into()), CellValue::Integer(30)],
                vec![CellValue::Boolean(true), CellValue::Empty],
            ],
            Some(&["Name".to_string(), "Age".to_string()]),
            true,
        )
        .unwrap();
        w.finalize().unwrap();

        let mut r = crate::XlsxReader::new();
        r.open(&path, true).unwrap();
        assert_eq!(r.sheet_names(), &["Sheet1".to_string()]);
        assert!(r.read().unwrap());
        assert_eq!(r.get_value(0), CellValue::Text("Name".into()));
        assert!(r.read().unwrap());
        assert_eq!(r.get_value(0), CellValue::Text("Name".into()));
        assert_eq!(r.get_value(1), CellValue::Number(30.0));
    }

    #[test]
    fn streaming_matches_batch() {
        let dir = tempfile::tempdir().unwrap();
        let path = dir.path().join("stream.xlsx");
        let mut w = XlsxWriter::create(&path).unwrap();
        w.start_sheet(
            "Data",
            2,
            Some(&["A".to_string(), "B".to_string()]),
            SheetOptions::new(),
        )
        .unwrap();
        w.write_row(&[CellValue::Integer(1), CellValue::Text("x".into())])
            .unwrap();
        w.end_sheet().unwrap();
        w.finalize().unwrap();
        let mut r = crate::XlsxReader::new();
        r.open(&path, true).unwrap();
        let mut rows = 0;
        while r.read().unwrap() {
            rows += 1;
        }
        assert_eq!(rows, 2);
    }

    #[test]
    fn columns_beyond_zz_are_written_correctly() {
        let dir = tempfile::tempdir().unwrap();
        let path = dir.path().join("wide.xlsx");
        let mut w = XlsxWriter::create(&path).unwrap();
        w.add_sheet("Wide", false);
        w.write_sheet(vec![vec![CellValue::Integer(1); 703]], None, false)
            .unwrap();
        w.finalize().unwrap();

        let file = std::fs::File::open(path).unwrap();
        let mut zip = zip::ZipArchive::new(file).unwrap();
        let mut xml = String::new();
        use std::io::Read;
        zip.by_name("xl/worksheets/sheet1.xml")
            .unwrap()
            .read_to_string(&mut xml)
            .unwrap();
        assert!(xml.contains("<dimension ref=\"A1:AAA1\"/>"));
    }

    #[test]
    fn streaming_dimension_reflects_written_rows() {
        let dir = tempfile::tempdir().unwrap();
        let path = dir.path().join("stream_dim.xlsx");
        let mut w = XlsxWriter::create(&path).unwrap();
        let headers: Vec<String> = ["H1", "H2", "H3"].iter().map(|s| s.to_string()).collect();
        w.start_sheet("Stream", 3, Some(&headers), SheetOptions::new())
            .unwrap();
        for i in 0..4i64 {
            w.write_row(&[
                CellValue::Integer(i),
                CellValue::Text("x".into()),
                CellValue::Boolean(true),
            ])
            .unwrap();
        }
        w.end_sheet().unwrap();
        w.finalize().unwrap();

        let file = std::fs::File::open(path).unwrap();
        let mut zip = zip::ZipArchive::new(file).unwrap();
        let mut xml = String::new();
        use std::io::Read;
        zip.by_name("xl/worksheets/sheet1.xml")
            .unwrap()
            .read_to_string(&mut xml)
            .unwrap();
        // Header row + 4 data rows, 3 columns.
        assert!(xml.contains(r#"<dimension ref="A1:C5"/>"#));
        // The schema requires <dimension> before <sheetViews>.
        let dim = xml.find("<dimension").unwrap();
        let views = xml.find("<sheetViews").unwrap();
        assert!(dim < views);
    }

    #[test]
    fn duplicate_names_are_disambiguated_and_stream_state_is_atomic() {
        let dir = tempfile::tempdir().unwrap();
        let path = dir.path().join("state.xlsx");
        let mut w = XlsxWriter::create(&path).unwrap();
        w.add_sheet("Data", false);
        w.add_sheet("data", false);
        assert_eq!(w.sheet_list[1].name, "data (2)");
        w.start_sheet("Stream", 1, None, SheetOptions::new())
            .unwrap();
        assert!(matches!(
            w.start_sheet("Rejected", 1, None, SheetOptions::new()),
            Err(SpreadsheetError::AlreadyStreaming)
        ));
        assert_eq!(w.sheet_count, 3);
        assert!(matches!(
            w.finalize(),
            Err(SpreadsheetError::StreamingSheetOpen)
        ));
    }
}
