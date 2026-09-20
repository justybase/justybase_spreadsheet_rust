//! Update the data of a worksheet inside an existing XLSB file without
//! rebuilding the workbook (port of `XlsbUpdater.ts`).
//!
//! ```no_run
//! use spreadsheet::{XlsbUpdater, CellValue};
//! use std::path::Path;
//!
//! let mut updater = XlsbUpdater::open(Path::new("report.xlsb")).unwrap();
//! updater.replace_sheet_data("data1", vec![], None).unwrap();
//! updater.save(None).unwrap();
//! ```

use crate::biff12::{
    build_record, parse_shared_strings_bin, read_record, read_utf16, uses_1904_date_system_bin,
    Biff12Record,
};
use crate::error::{SpreadsheetError, SpreadsheetResult};
use crate::formats::{borrow_cell, CellValue, CellValueRef};
use crate::xlsb_reader::parse_relationships;
use crate::zip_store::ZipStore;
use crate::{datetime_to_excel_serial, xml_utils};
use std::collections::{HashMap, HashSet};
use std::fmt::Display;
use std::path::{Path, PathBuf};

pub use crate::xlsx_updater::{ReplaceSheetDataOptions, StyleFallback};

/// Worksheet cell records (BrtCellBlank … BrtFmlaError).
fn is_cell_record(id: u32) -> bool {
    (0x00..=0x0b).contains(&id)
}

const RK_INT_LOWER: i64 = -1 << 29;
const RK_INT_UPPER: i64 = (1 << 29) - 1;

pub struct XlsbUpdater {
    zip: ZipStore,
    sheet_name_to_path: HashMap<String, String>,
    sheet_order: Vec<String>,
    pivot_cache_def_paths: Vec<String>,

    shared_strings_buf: Option<Vec<u8>>,
    shared_strings_values: Vec<String>,
    shared_strings_total: u32,
    shared_strings_unique: u32,
    shared_strings_end_sst: usize,
    string_index_map: HashMap<String, usize>,
    append_count: usize,
    shared_strings_dirty: bool,
    uses_1904_date_system: bool,
}

impl XlsbUpdater {
    /// Open an existing `.xlsb` file.
    pub fn open(path: &Path) -> SpreadsheetResult<Self> {
        let zip = ZipStore::open(path).map_err(|e| match e {
            SpreadsheetError::Io(inner) => SpreadsheetError::Other(format!(
                "XlsbUpdater: file not found: {} ({inner})",
                path.display()
            )),
            other => other,
        })?;
        if !zip.has_entry("xl/workbook.bin") {
            return Err(SpreadsheetError::InvalidFormat(
                "XlsbUpdater: not an XLSB workbook (xl/workbook.bin missing). Convert the file to XLSB first."
                    .into(),
            ));
        }
        let mut updater = Self {
            zip,
            sheet_name_to_path: HashMap::new(),
            sheet_order: Vec::new(),
            pivot_cache_def_paths: Vec::new(),
            shared_strings_buf: None,
            shared_strings_values: Vec::new(),
            shared_strings_total: 0,
            shared_strings_unique: 0,
            shared_strings_end_sst: 0,
            string_index_map: HashMap::new(),
            append_count: 0,
            shared_strings_dirty: false,
            uses_1904_date_system: false,
        };
        updater.load_workbook_structure()?;
        Ok(updater)
    }

    /// Worksheet names in workbook order.
    pub fn sheet_names(&self) -> Vec<String> {
        self.sheet_order.clone()
    }

    fn sheet_path(&self, sheet: &str) -> SpreadsheetResult<String> {
        self.sheet_name_to_path.get(sheet).cloned().ok_or_else(|| {
            SpreadsheetError::SheetNotFound(format!(
                "XlsbUpdater: sheet \"{sheet}\" not found in the workbook."
            ))
        })
    }

    /// Clear the target sheet's cell data and write new rows in its place.
    pub fn replace_sheet_data(
        &mut self,
        sheet_name: &str,
        rows: Vec<Vec<CellValue>>,
        options: Option<ReplaceSheetDataOptions>,
    ) -> SpreadsheetResult<()> {
        let options = options.unwrap_or_default();
        let rows = xml_utils::trim_trailing_empty_rows(rows, |v| matches!(v, CellValue::Empty));
        let sheet_path = self.sheet_path(sheet_name)?;
        if !self.zip.has_entry(&sheet_path) {
            return Err(SpreadsheetError::InvalidFormat(format!(
                "XlsbUpdater: worksheet part missing for sheet \"{sheet_name}\"."
            )));
        }
        let sheet_buf = self.zip.entry_data(&sheet_path)?;
        self.ensure_shared_strings_loaded()?;
        if self.shared_strings_buf.is_some() {
            let refs = shared_string_ref_count(&sheet_buf)?;
            self.shared_strings_total = self
                .shared_strings_total
                .saturating_sub(refs.min(u32::MAX as usize) as u32);
            self.shared_strings_dirty = true;
        }

        let (mut data_styles, mut header_styles, mut date_xf) =
            (HashMap::new(), HashMap::new(), None);
        if options.style_fallback != StyleFallback::General {
            let collected = collect_styles(&sheet_buf)?;
            data_styles = collected.0;
            header_styles = collected.1;
            date_xf = self.find_date_xf()?;
        }

        let max_len = rows
            .iter()
            .map(|r| r.len())
            .chain(options.headers.as_ref().map(|h| h.len()))
            .max()
            .unwrap_or(0);
        let last_col = max_len.saturating_sub(1) as u32;

        let new_rows = self.build_rows(
            &rows,
            options.headers.as_deref(),
            &data_styles,
            &header_styles,
            date_xf,
        );
        let region = find_rows_region(&sheet_buf)?;

        let mut new_sheet = Vec::with_capacity(
            sheet_buf.len() - (region.rows_end - region.rows_start) + new_rows.len(),
        );
        new_sheet.extend_from_slice(&sheet_buf[..region.rows_start]);
        new_sheet.extend_from_slice(&new_rows);
        new_sheet.extend_from_slice(&sheet_buf[region.rows_end..]);

        let last_row_idx = (rows.len() + options.headers.as_ref().map(|_| 1).unwrap_or(0))
            .saturating_sub(1) as u32;
        let delta = new_rows.len() as i64 - (region.rows_end - region.rows_start) as i64;
        for af in &region.auto_filter_ranges {
            let new_af = (*af as i64 + delta) as usize;
            write_i32_le_at(&mut new_sheet, new_af + 12, last_col as i32)?;
            write_i32_le_at(&mut new_sheet, new_af + 4, last_row_idx as i32)?;
        }
        patch_dimension(&mut new_sheet, last_row_idx, last_col)?;

        self.zip.update_file(&sheet_path, new_sheet);

        if self.shared_strings_buf.is_some() {
            self.commit_shared_strings()?;
        }

        let rw_last = last_row_idx;
        self.patch_pivot_caches(sheet_name, rw_last, last_col)?;
        Ok(())
    }

    /// Replace worksheet rows from a one-pass source.
    pub fn replace_sheet_data_stream(
        &mut self,
        sheet_name: &str,
        rows: impl Iterator<Item = Vec<CellValue>>,
        options: Option<ReplaceSheetDataOptions>,
    ) -> SpreadsheetResult<()> {
        self.replace_sheet_data_result_stream(
            sheet_name,
            rows.map(Ok::<Vec<CellValue>, SpreadsheetError>),
            options,
        )
    }

    /// Replace worksheet rows from a fallible one-pass source.
    ///
    /// If the source returns an error after producing rows, all staged sheet
    /// and shared-string changes are rolled back so the updater can be reused.
    pub fn replace_sheet_data_fallible_stream<E: Display>(
        &mut self,
        sheet_name: &str,
        rows: impl Iterator<Item = Result<Vec<CellValue>, E>>,
        options: Option<ReplaceSheetDataOptions>,
    ) -> SpreadsheetResult<()> {
        self.replace_sheet_data_result_stream(
            sheet_name,
            rows.map(|row| row.map_err(|error| SpreadsheetError::Other(error.to_string()))),
            options,
        )
    }

    fn replace_sheet_data_result_stream(
        &mut self,
        sheet_name: &str,
        rows: impl Iterator<Item = SpreadsheetResult<Vec<CellValue>>>,
        options: Option<ReplaceSheetDataOptions>,
    ) -> SpreadsheetResult<()> {
        let options = options.unwrap_or_default();
        if !matches!(
            options.style_fallback,
            StyleFallback::Inherit | StyleFallback::General
        ) {
            return Err(SpreadsheetError::InvalidFormat(
                "XlsbUpdater: styleFallback must be 'inherit' or 'general'.".into(),
            ));
        }
        let sheet_path = self.sheet_path(sheet_name)?;
        if !self.zip.has_entry(&sheet_path) {
            return Err(SpreadsheetError::InvalidFormat(format!(
                "XlsbUpdater: worksheet part missing for sheet \"{sheet_name}\"."
            )));
        }

        let snap = self.snapshot_sst();
        let old_stage = self.zip.staged_path(&sheet_path).map(|p| p.to_path_buf());
        let before_sst_bytes = if self.zip.has_entry("xl/sharedStrings.bin") {
            Some(self.zip.entry_data("xl/sharedStrings.bin")?)
        } else {
            None
        };
        let before_sst_buf = self.shared_strings_buf.clone();
        let old_sheet = self.zip.entry_data(&sheet_path)?;
        let rows_path = self.zip.temp_path(".rows")?;
        let output_path = self.zip.temp_path(".bin")?;
        let mut created = vec![rows_path.clone(), output_path.clone()];

        let result = self.replace_stream_inner(
            sheet_name,
            &sheet_path,
            &old_sheet,
            rows,
            &options,
            &rows_path,
            &output_path,
        );
        match result {
            Ok(()) => {
                if let Some(old) = old_stage {
                    if old != output_path {
                        ZipStore::remove_temp_file(&old);
                    }
                }
                created.retain(|p| p != &output_path);
                for p in created {
                    ZipStore::remove_temp_file(&p);
                }
                Ok(())
            }
            Err(e) => {
                self.restore_sst(snap);
                self.shared_strings_buf = before_sst_buf;
                if let Some(bytes) = before_sst_bytes {
                    self.zip.update_file("xl/sharedStrings.bin", bytes);
                }
                self.zip
                    .rollback_staged(&sheet_path, old_stage, &output_path);
                for p in created {
                    ZipStore::remove_temp_file(&p);
                }
                Err(e)
            }
        }
    }

    #[allow(clippy::too_many_arguments)]
    fn replace_stream_inner(
        &mut self,
        sheet_name: &str,
        sheet_path: &str,
        old_sheet: &[u8],
        rows: impl Iterator<Item = SpreadsheetResult<Vec<CellValue>>>,
        options: &ReplaceSheetDataOptions,
        rows_path: &Path,
        output_path: &Path,
    ) -> SpreadsheetResult<()> {
        self.ensure_shared_strings_loaded()?;
        if self.shared_strings_buf.is_some() {
            let refs = shared_string_ref_count(old_sheet)?;
            self.shared_strings_total = self
                .shared_strings_total
                .saturating_sub(refs.min(u32::MAX as usize) as u32);
            self.shared_strings_dirty = true;
        }
        let (mut data_styles, mut header_styles, mut date_xf) =
            (HashMap::new(), HashMap::new(), None);
        if options.style_fallback != StyleFallback::General {
            let collected = collect_styles(old_sheet)?;
            data_styles = collected.0;
            header_styles = collected.1;
            date_xf = self.find_date_xf()?;
        }

        let rows_file = std::fs::File::create(rows_path)?;
        let mut writer = std::io::BufWriter::new(rows_file);
        use std::io::Write as _;

        let mut row_header_offsets: Vec<u64> = Vec::new();
        let mut written: u64 = 0;
        let mut last_kept_end: u64 = 0;
        let mut width: usize = options.headers.as_ref().map(|h| h.len()).unwrap_or(0);
        let mut data_rows_seen: usize = 0;
        let mut kept_data_rows: usize = 0;
        let mut pending_empty_width: usize = 0;
        let mut next_row: u32 = 0;

        let write_row = |row_number: u32,
                         cells: Vec<Vec<u8>>,
                         writer: &mut std::io::BufWriter<std::fs::File>,
                         written: &mut u64| {
            let mut header = vec![0u8; 27];
            header[0] = 0x00;
            header[1] = 25;
            header[2..6].copy_from_slice(&row_number.to_le_bytes());
            header[10] = 0x2c;
            header[11] = 0x01;
            header[15] = 0x01;
            // colFirst/colLast patched later (lastCol unknown up-front)
            writer.write_all(&header)?;
            *written += header.len() as u64;
            for cell in &cells {
                writer.write_all(cell)?;
                *written += cell.len() as u64;
            }
            Ok::<(), std::io::Error>(())
        };

        if let Some(headers) = &options.headers {
            row_header_offsets.push(written);
            let cells: Vec<Vec<u8>> = headers
                .iter()
                .enumerate()
                .map(|(col, value)| {
                    let style = if options.style_fallback == StyleFallback::General {
                        0
                    } else {
                        header_styles.get(&col).copied().unwrap_or(0)
                    };
                    self.string_cell_bytes(col as u32, style, value)
                })
                .collect();
            write_row(next_row, cells, &mut writer, &mut written)?;
            last_kept_end = written;
            next_row = 1;
        }

        for row in rows {
            let row = row?;
            row_header_offsets.push(written);
            let cells: Vec<Vec<u8>> = row
                .iter()
                .enumerate()
                .filter_map(|(col, value)| {
                    if matches!(value, CellValue::Empty) {
                        None
                    } else {
                        Some(self.value_cell_bytes(value, col as u32, &data_styles, date_xf))
                    }
                })
                .collect();
            write_row(next_row, cells, &mut writer, &mut written)?;
            data_rows_seen += 1;
            next_row += 1;
            if row.iter().all(|v| matches!(v, CellValue::Empty)) {
                pending_empty_width = pending_empty_width.max(row.len());
            } else {
                width = width.max(pending_empty_width).max(row.len());
                pending_empty_width = 0;
                kept_data_rows = data_rows_seen;
                last_kept_end = written;
            }
        }
        writer.flush()?;
        drop(writer);

        let f = std::fs::OpenOptions::new().write(true).open(rows_path)?;
        f.set_len(last_kept_end)?;
        drop(f);

        let last_col = width.saturating_sub(1) as u32;
        // Patch colLast (payload offset 23) of every kept row header.
        {
            let f = std::fs::OpenOptions::new().write(true).open(rows_path)?;
            use std::io::{Seek, SeekFrom};
            let mut f = f;
            let bytes = last_col.to_le_bytes();
            for off in &row_header_offsets {
                if *off >= last_kept_end {
                    continue;
                }
                f.seek(SeekFrom::Start(off + 23))?;
                f.write_all(&bytes)?;
            }
            f.sync_all()?;
        }

        let region = find_rows_region(old_sheet)?;
        let new_rows_len = std::fs::metadata(rows_path)?.len();
        let delta = new_rows_len as i64 - (region.rows_end - region.rows_start) as i64;

        // Splice: old[..rowsStart] + rows file + old[rowsEnd..]
        {
            use std::io::Read as _;
            let out = std::fs::File::create(output_path)?;
            let mut out = std::io::BufWriter::new(out);
            out.write_all(&old_sheet[..region.rows_start])?;
            let mut rf = std::fs::File::open(rows_path)?;
            let mut buf = [0u8; 1024 * 1024];
            let mut remaining = new_rows_len;
            while remaining > 0 {
                let take = remaining.min(buf.len() as u64) as usize;
                let n = rf.read(&mut buf[..take])?;
                if n == 0 {
                    return Err(SpreadsheetError::InvalidFormat(
                        "XlsbUpdater: failed to read staged rows.".into(),
                    ));
                }
                out.write_all(&buf[..n])?;
                remaining -= n as u64;
            }
            out.write_all(&old_sheet[region.rows_end..])?;
            out.flush()?;
        }

        let last_row_idx = (kept_data_rows + options.headers.as_ref().map(|_| 1).unwrap_or(0))
            .saturating_sub(1) as u32;
        patch_sheet_metadata_file(
            output_path,
            old_sheet,
            region.rows_end,
            delta,
            &region.auto_filter_ranges,
            last_row_idx,
            last_col,
        )?;
        self.zip.stage_part(sheet_path, output_path.to_path_buf());
        self.commit_shared_strings()?;
        self.patch_pivot_caches(sheet_name, last_row_idx, last_col)?;
        Ok(())
    }

    /// Write the updated workbook to disk (defaults to in-place overwrite).
    pub fn save(&mut self, output: Option<&Path>) -> SpreadsheetResult<()> {
        let target = output
            .map(|p| p.to_path_buf())
            .unwrap_or_else(|| self.zip.source_path().to_path_buf());
        self.zip.save(&target)
    }

    /// Save staged streaming replacements without materialising the ZIP.
    pub fn save_streaming(&mut self, output: Option<&Path>) -> SpreadsheetResult<()> {
        let target = output
            .map(|p| p.to_path_buf())
            .unwrap_or_else(|| self.zip.source_path().to_path_buf());
        self.zip.save_streaming(&target)
    }

    /// The updated workbook as an in-memory ZIP archive.
    pub fn buffer(&mut self) -> SpreadsheetResult<Vec<u8>> {
        self.zip.buffer()
    }

    /// The updated workbook as an in-memory ZIP archive.
    ///
    /// Use [`Self::buffer`] in new code. This compatibility spelling is kept
    /// for callers of the original public API.
    #[deprecated(note = "use buffer")]
    #[allow(clippy::wrong_self_convention)]
    pub fn to_buffer(&mut self) -> SpreadsheetResult<Vec<u8>> {
        self.buffer()
    }

    /// Release the staging directory and any unsaved staged parts.
    pub fn dispose(&mut self) {
        self.zip.dispose();
    }

    // -- workbook structure ---------------------------------------------------

    fn load_workbook_structure(&mut self) -> SpreadsheetResult<()> {
        let rels = self
            .zip
            .entry_data("xl/_rels/workbook.bin.rels")
            .unwrap_or_default();
        let r_id_to_target = parse_relationships(&String::from_utf8_lossy(&rels));
        let wb = self.zip.entry_data("xl/workbook.bin")?;
        self.uses_1904_date_system = uses_1904_date_system_bin(&wb)?;
        let mut rec = Biff12Record::default();
        let mut pos = 0;
        while read_record(&wb, pos, &mut rec)? {
            pos = rec.data_end;
            if rec.id != 0x009c || rec.len < 8 {
                continue;
            }
            // BrtBundleSh: hidden(4) + sheetId(4) + rId(string) + name(string)
            let mut o = rec.data_start + 8;
            let r_id_len = u32::from_le_bytes(copy4(&wb, o)) as usize;
            o += 4;
            let r_id = read_utf16(&wb, o, r_id_len);
            o += r_id_len * 2;
            let name_len = u32::from_le_bytes(copy4(&wb, o)) as usize;
            o += 4;
            let name = read_utf16(&wb, o, name_len);
            let target = r_id_to_target.get(&r_id).cloned().unwrap_or_default();
            let mut full = target;
            if let Some(stripped) = full.strip_prefix('/') {
                full = stripped.to_owned();
            } else if !full.starts_with("xl/") {
                full = format!("xl/{full}");
            }
            if !full.is_empty() {
                self.sheet_order.push(name.clone());
                self.sheet_name_to_path.insert(name, full);
            }
        }
        for name in self.zip.entry_names() {
            if is_pivot_cache_bin(&name) {
                self.pivot_cache_def_paths.push(name);
            }
        }
        Ok(())
    }

    // -- shared strings ----------------------------------------------------------

    fn ensure_shared_strings_loaded(&mut self) -> SpreadsheetResult<()> {
        if !self.string_index_map.is_empty() || self.shared_strings_buf.is_some() {
            return Ok(());
        }
        if !self.zip.has_entry("xl/sharedStrings.bin") {
            self.shared_strings_buf = None;
            return Ok(());
        }
        let buf = self.zip.entry_data("xl/sharedStrings.bin")?;
        let parsed = parse_shared_strings_bin(&buf)?;
        for (i, v) in parsed.values.iter().enumerate() {
            self.string_index_map.entry(v.clone()).or_insert(i);
        }
        self.shared_strings_values = parsed.values;
        self.shared_strings_total = parsed.total;
        self.shared_strings_unique = parsed.unique;
        self.shared_strings_end_sst = parsed.end_sst_offset;
        self.shared_strings_buf = Some(buf);
        self.append_count = 0;
        self.shared_strings_dirty = false;
        Ok(())
    }

    fn commit_shared_strings(&mut self) -> SpreadsheetResult<()> {
        if self.shared_strings_buf.is_none() || !self.shared_strings_dirty {
            return Ok(());
        }
        let buf = self.shared_strings_buf.clone().unwrap();
        let original_len = self.shared_strings_values.len() - self.append_count;
        let mut new_items = Vec::new();
        for text in self.shared_strings_values[original_len..].iter() {
            let units: Vec<u16> = text.encode_utf16().collect();
            let mut payload = vec![0u8; 1 + 4 + units.len() * 2];
            payload[0] = 0x00;
            payload[1..5].copy_from_slice(&(units.len() as u32).to_le_bytes());
            for (i, u) in units.iter().enumerate() {
                payload[5 + i * 2..7 + i * 2].copy_from_slice(&u.to_le_bytes());
            }
            new_items.extend(build_record(0x0013, &payload));
        }
        let inserted_end = self
            .shared_strings_end_sst
            .checked_add(new_items.len())
            .ok_or_else(|| {
                SpreadsheetError::InvalidFormat(
                    "XlsbUpdater: shared-string insertion offset overflows".into(),
                )
            })?;
        let output_len = buf.len().checked_add(new_items.len()).ok_or_else(|| {
            SpreadsheetError::InvalidFormat("XlsbUpdater: shared-string size overflows".into())
        })?;
        let mut out = vec![0u8; output_len];
        out[..self.shared_strings_end_sst].copy_from_slice(&buf[..self.shared_strings_end_sst]);
        out[self.shared_strings_end_sst..self.shared_strings_end_sst + new_items.len()]
            .copy_from_slice(&new_items);
        out[self.shared_strings_end_sst + new_items.len()..]
            .copy_from_slice(&buf[self.shared_strings_end_sst..]);

        // Patch cstTotal / cstUnique inside BrtBeginSst.
        let mut rec = Biff12Record::default();
        let mut pos = 0;
        while read_record(&out, pos, &mut rec)? {
            if rec.id == 0x009f {
                out[rec.data_start..rec.data_start + 4]
                    .copy_from_slice(&self.shared_strings_total.to_le_bytes());
                out[rec.data_start + 4..rec.data_start + 8]
                    .copy_from_slice(&self.shared_strings_unique.to_le_bytes());
                break;
            }
            pos = rec.data_end;
        }
        self.zip.update_file("xl/sharedStrings.bin", out.clone());
        self.shared_strings_buf = Some(out);
        self.shared_strings_end_sst = inserted_end;
        self.append_count = 0;
        self.shared_strings_dirty = false;
        Ok(())
    }

    // -- styles ---------------------------------------------------------------------

    fn find_date_xf(&self) -> SpreadsheetResult<Option<u32>> {
        if !self.zip.has_entry("xl/styles.bin") {
            return Ok(None);
        }
        let styles = self.zip.entry_data("xl/styles.bin")?;
        let mut reader = crate::biff_reader::BiffReaderWriter::new(&styles);
        while reader.read_styles()? {}
        for (xf_index, num_fmt) in reader.xf_index_to_num_fmt_id.iter().copied().enumerate() {
            if (14..=22).contains(&num_fmt)
                || (45..=47).contains(&num_fmt)
                || reader.custom_num_fmts.contains(&num_fmt)
            {
                return Ok(Some(xf_index as u32));
            }
        }
        Ok(None)
    }

    // -- row / cell serialization ------------------------------------------------------

    fn build_rows(
        &mut self,
        rows: &[Vec<CellValue>],
        headers: Option<&[String]>,
        data_styles: &HashMap<usize, u32>,
        header_styles: &HashMap<usize, u32>,
        date_xf: Option<u32>,
    ) -> Vec<u8> {
        let max_col = rows
            .iter()
            .map(|r| r.len())
            .chain(headers.map(|h| h.len()))
            .max()
            .unwrap_or(0);
        let col_last = max_col.saturating_sub(1) as u32;
        let mut chunks: Vec<Vec<u8>> = Vec::new();
        let push_row = |row_num: u32, cells: Vec<Vec<u8>>, chunks: &mut Vec<Vec<u8>>| {
            let mut header = vec![0u8; 27];
            header[0] = 0x00;
            header[1] = 25;
            header[2..6].copy_from_slice(&row_num.to_le_bytes());
            header[10] = 0x2c;
            header[11] = 0x01;
            header[15] = 0x01;
            header[19..23].copy_from_slice(&0u32.to_le_bytes());
            header[23..27].copy_from_slice(&col_last.to_le_bytes());
            chunks.push(header);
            chunks.extend(cells);
        };
        if let Some(h) = headers {
            let cells: Vec<Vec<u8>> = h
                .iter()
                .enumerate()
                .map(|(c, text)| {
                    let style = header_styles.get(&c).copied().unwrap_or(0);
                    self.string_cell_bytes(c as u32, style, text)
                })
                .collect();
            push_row(0, cells, &mut chunks);
        }
        let first_data_row = headers.map(|_| 1).unwrap_or(0);
        for (row_offset, row) in rows.iter().enumerate() {
            let row_num = (first_data_row + row_offset) as u32;
            let cells: Vec<Vec<u8>> = row
                .iter()
                .enumerate()
                .filter_map(|(c, raw)| {
                    if matches!(raw, CellValue::Empty) {
                        None
                    } else {
                        Some(self.value_cell_bytes(raw, c as u32, data_styles, date_xf))
                    }
                })
                .collect();
            push_row(row_num, cells, &mut chunks);
        }
        chunks.concat()
    }

    fn value_cell_bytes(
        &mut self,
        raw: &CellValue,
        col: u32,
        data_styles: &HashMap<usize, u32>,
        date_xf: Option<u32>,
    ) -> Vec<u8> {
        let val = borrow_cell(raw);
        let col_style = data_styles.get(&(col as usize)).copied().unwrap_or(0);
        match val {
            CellValueRef::Empty => self.string_cell_bytes(col, col_style, ""),
            CellValueRef::Integer(n) => {
                if (RK_INT_LOWER..=RK_INT_UPPER).contains(&n) {
                    let mut payload = vec![0u8; 12];
                    payload[0..4].copy_from_slice(&col.to_le_bytes());
                    payload[4..8].copy_from_slice(&col_style.to_le_bytes());
                    payload[8..12].copy_from_slice(&(((n as i32) << 2) | 2).to_le_bytes());
                    build_record(0x0002, &payload)
                } else {
                    let n = n as f64;
                    let mut payload = vec![0u8; 16];
                    payload[0..4].copy_from_slice(&col.to_le_bytes());
                    payload[4..8].copy_from_slice(&col_style.to_le_bytes());
                    payload[8..16].copy_from_slice(&n.to_le_bytes());
                    build_record(0x0005, &payload)
                }
            }
            CellValueRef::Number(n) => {
                if n.is_finite()
                    && n.fract() == 0.0
                    && n >= RK_INT_LOWER as f64
                    && n <= RK_INT_UPPER as f64
                {
                    let mut payload = vec![0u8; 12];
                    payload[0..4].copy_from_slice(&col.to_le_bytes());
                    payload[4..8].copy_from_slice(&col_style.to_le_bytes());
                    payload[8..12].copy_from_slice(&(((n as i32) << 2) | 2).to_le_bytes());
                    build_record(0x0002, &payload)
                } else if n.is_finite() {
                    let mut payload = vec![0u8; 16];
                    payload[0..4].copy_from_slice(&col.to_le_bytes());
                    payload[4..8].copy_from_slice(&col_style.to_le_bytes());
                    payload[8..16].copy_from_slice(&n.to_le_bytes());
                    build_record(0x0005, &payload)
                } else {
                    self.string_cell_bytes(col, col_style, &n.to_string())
                }
            }
            CellValueRef::Boolean(b) => {
                let mut payload = vec![0u8; 9];
                payload[0..4].copy_from_slice(&col.to_le_bytes());
                payload[4..8].copy_from_slice(&col_style.to_le_bytes());
                payload[8] = if b { 1 } else { 0 };
                build_record(0x0004, &payload)
            }
            CellValueRef::DateTime(d) => {
                let oa = datetime_to_excel_serial(&d, self.uses_1904_date_system);
                if oa.is_finite() {
                    let style = if col_style > 0 {
                        col_style
                    } else {
                        date_xf.unwrap_or(0)
                    };
                    let mut payload = vec![0u8; 16];
                    payload[0..4].copy_from_slice(&col.to_le_bytes());
                    payload[4..8].copy_from_slice(&style.to_le_bytes());
                    payload[8..16].copy_from_slice(&oa.to_le_bytes());
                    build_record(0x0005, &payload)
                } else {
                    self.string_cell_bytes(col, col_style, &d.to_string())
                }
            }
            CellValueRef::Text(s) => self.string_cell_bytes(col, col_style, s),
        }
    }

    fn string_cell_bytes(&mut self, col: u32, style: u32, text: &str) -> Vec<u8> {
        if self.shared_strings_buf.is_none() {
            // Inline string (BrtCellSt): col(4) + xf(4) + cch(4) + utf16.
            let units: Vec<u16> = text.encode_utf16().collect();
            let mut payload = vec![0u8; 12 + units.len() * 2];
            payload[0..4].copy_from_slice(&col.to_le_bytes());
            payload[4..8].copy_from_slice(&style.to_le_bytes());
            payload[8..12].copy_from_slice(&(units.len() as u32).to_le_bytes());
            for (i, u) in units.iter().enumerate() {
                payload[12 + i * 2..14 + i * 2].copy_from_slice(&u.to_le_bytes());
            }
            return build_record(0x0006, &payload);
        }
        let index = match self.string_index_map.get(text) {
            Some(&i) => i,
            None => {
                let i = self.shared_strings_values.len();
                let owned = text.to_owned();
                self.shared_strings_values.push(owned.clone());
                self.string_index_map.insert(owned, i);
                self.shared_strings_unique += 1;
                self.append_count += 1;
                i
            }
        };
        self.shared_strings_total += 1;
        self.shared_strings_dirty = true;
        let mut payload = vec![0u8; 12];
        payload[0..4].copy_from_slice(&col.to_le_bytes());
        payload[4..8].copy_from_slice(&style.to_le_bytes());
        payload[8..12].copy_from_slice(&(index as u32).to_le_bytes());
        build_record(0x0007, &payload)
    }

    // -- pivot caches ---------------------------------------------------------------------

    fn patch_pivot_caches(
        &mut self,
        sheet_name: &str,
        rw_last: u32,
        col_last: u32,
    ) -> SpreadsheetResult<()> {
        for path in self.pivot_cache_def_paths.clone() {
            let mut buf = match self.zip.entry_data(&path) {
                Ok(b) => b,
                Err(_) => continue,
            };
            let mut rec = Biff12Record::default();
            let mut pos = 0;
            let mut patched = false;
            let mut refresh_patched = false;
            while read_record(&buf, pos, &mut rec)? {
                pos = rec.data_end;
                if rec.id == 0x00b3 && rec.len >= 4 && !refresh_patched {
                    // BrtBeginPivotCacheDefinition flags: bit 0x04 of payload
                    // offset 3 = refreshOnLoad.
                    if rec.data_start + 4 <= buf.len() {
                        let v = buf[rec.data_start + 3];
                        if v & 0x04 == 0 {
                            buf[rec.data_start + 3] = v | 0x04;
                            patched = true;
                        }
                    }
                    refresh_patched = true;
                }
                if rec.id == 0x00bb && rec.len >= 23 {
                    // BrtPivotCacheSource: [3] + cch(4) + name(utf16) +
                    // rwFirst(4) rwLast(4) colFirst(4) colLast(4)
                    let cch = u32::from_le_bytes(copy4(&buf, rec.data_start + 3)) as usize;
                    let name_start = rec.data_start + 7;
                    let ref_start = name_start + cch * 2;
                    let name = read_utf16(&buf, name_start, cch);
                    if name == sheet_name && ref_start + 16 <= buf.len() {
                        buf[ref_start + 4..ref_start + 8].copy_from_slice(&rw_last.to_le_bytes());
                        buf[ref_start + 12..ref_start + 16]
                            .copy_from_slice(&col_last.to_le_bytes());
                        patched = true;
                    }
                }
            }
            if patched {
                self.zip.update_file(&path, buf);
            }
        }
        Ok(())
    }

    // -- SST snapshot / rollback ---------------------------------------------------------------

    fn snapshot_sst(&self) -> BiffSstSnapshot {
        BiffSstSnapshot {
            values_len: self.shared_strings_values.len(),
            append_count: self.append_count,
            total: self.shared_strings_total,
            unique: self.shared_strings_unique,
            end_sst: self.shared_strings_end_sst,
            dirty: self.shared_strings_dirty,
            index_map: self.string_index_map.clone(),
        }
    }

    fn restore_sst(&mut self, snap: BiffSstSnapshot) {
        self.shared_strings_values.truncate(snap.values_len);
        self.append_count = snap.append_count;
        self.shared_strings_total = snap.total;
        self.shared_strings_unique = snap.unique;
        self.shared_strings_end_sst = snap.end_sst;
        self.shared_strings_dirty = snap.dirty;
        self.string_index_map = snap.index_map;
    }
}

struct BiffSstSnapshot {
    values_len: usize,
    append_count: usize,
    total: u32,
    unique: u32,
    end_sst: usize,
    dirty: bool,
    index_map: HashMap<String, usize>,
}

impl Drop for XlsbUpdater {
    fn drop(&mut self) {
        self.zip.dispose();
    }
}

// -- free helpers -------------------------------------------------------------------

fn copy4(buf: &[u8], at: usize) -> [u8; 4] {
    [buf[at], buf[at + 1], buf[at + 2], buf[at + 3]]
}

fn write_i32_le_at(buf: &mut [u8], at: usize, value: i32) -> SpreadsheetResult<()> {
    if at + 4 > buf.len() {
        return Err(SpreadsheetError::InvalidFormat(
            "XlsbUpdater: autofilter patch out of range".into(),
        ));
    }
    buf[at..at + 4].copy_from_slice(&value.to_le_bytes());
    Ok(())
}

fn is_pivot_cache_bin(name: &str) -> bool {
    name.starts_with("xl/pivotCache/pivotCacheDefinition") && name.ends_with(".bin") && {
        let mid = &name["xl/pivotCache/pivotCacheDefinition".len()..name.len() - 4];
        mid.chars().all(|c| c.is_ascii_digit())
    }
}

fn shared_string_ref_count(sheet_buf: &[u8]) -> SpreadsheetResult<usize> {
    let mut rec = Biff12Record::default();
    let mut pos = 0;
    let mut count = 0;
    while read_record(sheet_buf, pos, &mut rec)? {
        if rec.id == 0x0007 {
            count += 1;
        }
        pos = rec.data_end;
    }
    Ok(count)
}

fn collect_styles(
    sheet_buf: &[u8],
) -> SpreadsheetResult<(HashMap<usize, u32>, HashMap<usize, u32>)> {
    let mut data_counts: HashMap<usize, HashMap<u32, usize>> = HashMap::new();
    let mut header_styles: HashMap<usize, u32> = HashMap::new();
    let mut rec = Biff12Record::default();
    let mut pos = 0;
    let mut row: i32 = -1;
    while read_record(sheet_buf, pos, &mut rec)? {
        pos = rec.data_end;
        if rec.id == 0x00 {
            row = i32::from_le_bytes(copy4(sheet_buf, rec.data_start));
            continue;
        }
        if !is_cell_record(rec.id) || rec.len < 8 {
            continue;
        }
        let col = u32::from_le_bytes(copy4(sheet_buf, rec.data_start)) as usize;
        let xf = u32::from_le_bytes(copy4(sheet_buf, rec.data_start + 4)) & 0xffffff;
        if xf == 0 {
            continue;
        }
        if row == 0 {
            header_styles.entry(col).or_insert(xf);
            continue;
        }
        data_counts
            .entry(col)
            .or_default()
            .entry(xf)
            .and_modify(|c| *c += 1)
            .or_insert(1);
    }
    let mut data_styles = HashMap::new();
    for (col, counts) in data_counts {
        if let Some((&best, _)) = counts.iter().max_by_key(|(_, &c)| c) {
            data_styles.insert(col, best);
        }
    }
    Ok((data_styles, header_styles))
}

struct RowsRegion {
    rows_start: usize,
    rows_end: usize,
    auto_filter_ranges: Vec<usize>,
}

fn find_rows_region(sheet_buf: &[u8]) -> SpreadsheetResult<RowsRegion> {
    let mut auto_filter_ranges = Vec::new();
    let mut rows_start: Option<usize> = None;
    let mut last_cell_end: Option<usize> = None;
    let mut last_w25: Option<usize> = None;
    let mut rec = Biff12Record::default();
    let mut pos = 0;
    while read_record(sheet_buf, pos, &mut rec)? {
        match rec.id {
            0x00 => {
                if rows_start.is_none() {
                    rows_start = Some(rec.header_start);
                }
            }
            0x01..=0x0b => last_cell_end = Some(rec.data_end),
            0x25 if rec.len == 6 => last_w25 = Some(rec.header_start),
            0x00a1 if rec.len >= 16 => auto_filter_ranges.push(rec.data_start),
            _ => {}
        }
        pos = rec.data_end;
    }
    match rows_start {
        None => {
            let insert_at = last_w25.unwrap_or(sheet_buf.len());
            Ok(RowsRegion {
                rows_start: insert_at,
                rows_end: insert_at,
                auto_filter_ranges,
            })
        }
        Some(start) => {
            // Preserve the row-block terminator records (0x92/0x217/...)
            // that Excel writes after the last row.
            let end = last_cell_end.or(last_w25).unwrap_or(sheet_buf.len());
            Ok(RowsRegion {
                rows_start: start,
                rows_end: end,
                auto_filter_ranges,
            })
        }
    }
}

/// Keep the sheet dimension record (0x98) in sync: rwLast at payload
/// offset 24, colLast at payload offset 32.
fn patch_dimension(
    sheet_buf: &mut [u8],
    last_row_idx: u32,
    last_col: u32,
) -> SpreadsheetResult<()> {
    let mut rec = Biff12Record::default();
    let mut pos = 0;
    while read_record(sheet_buf, pos, &mut rec)? {
        if rec.id == 0x0098 && rec.len >= 36 {
            sheet_buf[rec.data_start + 24..rec.data_start + 28]
                .copy_from_slice(&last_row_idx.to_le_bytes());
            sheet_buf[rec.data_start + 32..rec.data_start + 36]
                .copy_from_slice(&last_col.to_le_bytes());
            return Ok(());
        }
        pos = rec.data_end;
    }
    Ok(())
}

fn patch_sheet_metadata_file(
    file_path: &Path,
    original_sheet: &[u8],
    rows_end: usize,
    delta: i64,
    auto_filter_ranges: &[usize],
    last_row_idx: u32,
    last_col: u32,
) -> SpreadsheetResult<()> {
    use std::io::{Seek, SeekFrom, Write};
    let mut fd = std::fs::OpenOptions::new().write(true).open(file_path)?;
    let patch = |fd: &mut std::fs::File, offset: u64, value: u32| -> SpreadsheetResult<()> {
        fd.seek(SeekFrom::Start(offset))?;
        fd.write_all(&value.to_le_bytes())?;
        Ok(())
    };
    for &filter_offset in auto_filter_ranges {
        let new_offset = if filter_offset >= rows_end {
            (filter_offset as i64 + delta) as u64
        } else {
            filter_offset as u64
        };
        patch(&mut fd, new_offset + 4, last_row_idx)?;
        patch(&mut fd, new_offset + 12, last_col)?;
    }
    let mut rec = Biff12Record::default();
    let mut pos = 0;
    while read_record(original_sheet, pos, &mut rec)? {
        if rec.id == 0x0098 && rec.len >= 36 {
            let shift = if rec.data_start >= rows_end { delta } else { 0 };
            patch(
                &mut fd,
                (rec.data_start as i64 + shift + 24) as u64,
                last_row_idx,
            )?;
            patch(
                &mut fd,
                (rec.data_start as i64 + shift + 32) as u64,
                last_col,
            )?;
            break;
        }
        pos = rec.data_end;
    }
    fd.sync_all()?;
    Ok(())
}

#[allow(dead_code)]
fn unused_xlsb_helpers_anchor() {
    let _ = HashSet::<u8>::new();
    let _ = PathBuf::new();
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::{XlsbReader, XlsbWriter};

    fn sample_workbook(path: &Path) {
        let mut w = XlsbWriter::create(path).unwrap();
        w.add_sheet("data1", false);
        w.write_sheet(
            vec![
                vec![CellValue::Text("Alice".into()), CellValue::Integer(30)],
                vec![CellValue::Text("Bob".into()), CellValue::Integer(25)],
            ],
            Some(&["NAME".to_string(), "AGE".to_string()]),
            true,
        )
        .unwrap();
        w.add_sheet("other", false);
        w.write_sheet(
            vec![vec![CellValue::Text("keep".into())]],
            Some(&["COL".to_string()]),
            false,
        )
        .unwrap();
        w.finalize().unwrap();
    }

    fn date_workbook(path: &Path, date: chrono::NaiveDateTime) {
        let mut w = XlsbWriter::create(path).unwrap();
        w.add_sheet("data1", false);
        w.write_sheet(
            vec![vec![CellValue::DateTime(date)]],
            Some(&["DATE".to_string()]),
            false,
        )
        .unwrap();
        w.finalize().unwrap();
    }

    #[test]
    fn replace_preserves_other_sheets() {
        let dir = tempfile::tempdir().unwrap();
        let path = dir.path().join("report.xlsb");
        sample_workbook(&path);

        let mut u = XlsbUpdater::open(&path).unwrap();
        assert!(u.sheet_names().contains(&"data1".to_string()));
        u.replace_sheet_data(
            "data1",
            vec![vec![CellValue::Text("Zed".into()), CellValue::Integer(99)]],
            Some(ReplaceSheetDataOptions {
                headers: Some(vec!["NAME".to_string(), "AGE".to_string()]),
                style_fallback: StyleFallback::General,
            }),
        )
        .unwrap();
        let out = dir.path().join("report_new.xlsb");
        u.save(Some(&out)).unwrap();

        let mut r = XlsbReader::new();
        r.open(&out, true).unwrap();
        assert!(r.read().unwrap());
        assert_eq!(r.get_value(0), CellValue::Text("NAME".into()));
        assert!(r.read().unwrap());
        assert_eq!(r.get_value(0), CellValue::Text("Zed".into()));
        assert_eq!(r.get_value(1), CellValue::Number(99.0));
    }

    #[test]
    fn stream_replace_then_save_streaming() {
        let dir = tempfile::tempdir().unwrap();
        let path = dir.path().join("s.xlsb");
        sample_workbook(&path);
        let mut u = XlsbUpdater::open(&path).unwrap();
        let rows = vec![vec![CellValue::Integer(1)], vec![CellValue::Empty]];
        u.replace_sheet_data_stream("data1", rows.into_iter(), None)
            .unwrap();
        let out = dir.path().join("s_new.xlsb");
        u.save_streaming(Some(&out)).unwrap();
        let mut r = XlsbReader::new();
        r.open(&out, true).unwrap();
        let mut non_empty = 0;
        loop {
            if !r.read().unwrap() {
                break;
            }
            if !r.current_row().is_empty() {
                non_empty += 1;
            }
            if non_empty > 10 {
                break;
            }
        }
        // 1 data row (no headers passed; trailing empties trimmed by truncation)
        assert_eq!(non_empty, 1);
    }

    #[test]
    fn missing_sheet_errors() {
        let dir = tempfile::tempdir().unwrap();
        let path = dir.path().join("m.xlsb");
        sample_workbook(&path);
        let mut u = XlsbUpdater::open(&path).unwrap();
        assert!(u.replace_sheet_data("nope", vec![], None).is_err());
    }

    fn mark_xlsb_as_1904(path: &Path) {
        let mut zip = crate::zip_store::ZipStore::open(path).unwrap();
        let mut workbook = zip.entry_data("xl/workbook.bin").unwrap();
        let mut rec = Biff12Record::default();
        let mut pos = 0;
        let mut found = false;
        while read_record(&workbook, pos, &mut rec).unwrap() {
            if rec.id == 0x0099 && rec.len >= 4 {
                let mut flags = u32::from_le_bytes(copy4(&workbook, rec.data_start));
                flags |= 1;
                workbook[rec.data_start..rec.data_start + 4].copy_from_slice(&flags.to_le_bytes());
                found = true;
                break;
            }
            pos = rec.data_end;
        }
        assert!(found, "generated workbook has no BrtBeginBook record");
        zip.update_file("xl/workbook.bin", workbook);
        zip.save(path).unwrap();
    }

    #[test]
    fn updater_writes_1904_date_serials() {
        let dir = tempfile::tempdir().unwrap();
        let path = dir.path().join("date1904.xlsb");
        let expected = chrono::NaiveDate::from_ymd_opt(2025, 2, 3)
            .unwrap()
            .and_hms_opt(12, 30, 0)
            .unwrap();
        date_workbook(&path, expected);
        mark_xlsb_as_1904(&path);

        let mut updater = XlsbUpdater::open(&path).unwrap();
        assert!(updater.uses_1904_date_system);
        updater
            .replace_sheet_data(
                "data1",
                vec![vec![CellValue::DateTime(expected)]],
                Some(ReplaceSheetDataOptions {
                    headers: None,
                    style_fallback: StyleFallback::Inherit,
                }),
            )
            .unwrap();
        let output = dir.path().join("date1904-out.xlsb");
        updater.save(Some(&output)).unwrap();

        let zip = crate::zip_store::ZipStore::open(&output).unwrap();
        let sheet = zip.entry_data("xl/worksheets/sheet1.bin").unwrap();
        let expected_serial = crate::datetime_to_excel_serial(&expected, true);
        let mut rec = Biff12Record::default();
        let mut pos = 0;
        let mut found = false;
        while read_record(&sheet, pos, &mut rec).unwrap() {
            if rec.id == 0x0005 && rec.len >= 16 {
                let serial = f64::from_le_bytes(
                    sheet[rec.data_start + 8..rec.data_start + 16]
                        .try_into()
                        .unwrap(),
                );
                if (serial - expected_serial).abs() < 1e-9 {
                    found = true;
                    break;
                }
            }
            pos = rec.data_end;
        }
        assert!(
            found,
            "updated sheet did not contain serial {expected_serial}"
        );

        let mut reader = XlsbReader::new();
        reader.open(&output, true).unwrap();
        assert!(reader.uses_1904_date_system);
        assert!(reader.read().unwrap());
        assert_eq!(reader.get_value(0), CellValue::DateTime(expected));
    }
}
