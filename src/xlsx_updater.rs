//! Update the data of a worksheet inside an existing XLSX file without
//! rebuilding the workbook (port of `XlsxUpdater.ts`).
//!
//! Everything outside the target sheet's cell data — pivot tables and
//! their caches, other sheets, styles, themes, defined names — is
//! preserved.
//!
//! ```no_run
//! use spreadsheet::{XlsxUpdater, CellValue};
//! use std::path::Path;
//!
//! let mut updater = XlsxUpdater::open(Path::new("report.xlsx")).unwrap();
//! updater.replace_sheet_data("data1", vec![], None).unwrap();
//! updater.save(None).unwrap();
//! ```

use crate::error::{SpreadsheetError, SpreadsheetResult};
use crate::formats::{borrow_cell, CellValue, CellValueRef};
use crate::xlsb_reader::parse_relationships;
use crate::zip_store::ZipStore;
use crate::{datetime_to_excel_serial, xml_utils};
use std::collections::{HashMap, HashSet};
use std::fmt::Display;
use std::path::Path;

/// Options for [`XlsxUpdater::replace_sheet_data`].
#[derive(Debug, Clone, Default)]
pub struct ReplaceSheetDataOptions {
    /// Optional header row written into row 1 of the target sheet.
    /// When omitted, row 1 is filled with the data rows.
    pub headers: Option<Vec<String>>,
    /// Style strategy for the new cells.
    pub style_fallback: StyleFallback,
}

/// Style strategy for replacement cells.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub enum StyleFallback {
    /// Reuse the dominant cell style of each column (default).
    #[default]
    Inherit,
    /// All new cells use style 0 (General).
    General,
}

struct PivotCacheDef {
    path: String,
    xml: String,
}

pub struct XlsxUpdater {
    zip: ZipStore,
    sheet_name_to_path: HashMap<String, String>,
    sheet_order: Vec<String>,
    pivot_cache_defs: Vec<PivotCacheDef>,
    pivot_table_def_paths: Vec<String>,

    shared_strings_xml: Option<String>,
    shared_strings_values: Vec<String>,
    shared_strings_count: usize,
    shared_strings_pending: Vec<String>,
    shared_strings_dirty: bool,
    string_index_map: HashMap<String, usize>,
    uses_1904_date_system: bool,
}

impl XlsxUpdater {
    /// Open an existing `.xlsx` file.
    pub fn open(path: &Path) -> SpreadsheetResult<Self> {
        let zip = ZipStore::open(path).map_err(|e| match e {
            SpreadsheetError::Io(inner) => SpreadsheetError::Other(format!(
                "XlsxUpdater: file not found: {} ({inner})",
                path.display()
            )),
            other => other,
        })?;
        if zip.has_entry("xl/workbook.bin") {
            return Err(SpreadsheetError::InvalidFormat(
                "XlsxUpdater: XLSB files are not supported yet. Convert the workbook to XLSX first."
                    .into(),
            ));
        }
        if !zip.has_entry("xl/workbook.xml") {
            return Err(SpreadsheetError::InvalidFormat(
                "XlsxUpdater: not a valid XLSX workbook (xl/workbook.xml missing).".into(),
            ));
        }
        let mut updater = Self {
            zip,
            sheet_name_to_path: HashMap::new(),
            sheet_order: Vec::new(),
            pivot_cache_defs: Vec::new(),
            pivot_table_def_paths: Vec::new(),
            shared_strings_xml: None,
            shared_strings_values: Vec::new(),
            shared_strings_count: 0,
            shared_strings_pending: Vec::new(),
            shared_strings_dirty: false,
            string_index_map: HashMap::new(),
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
                "XlsxUpdater: sheet \"{sheet}\" not found in the workbook."
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
                "XlsxUpdater: worksheet part missing for sheet \"{sheet_name}\"."
            )));
        }
        let sheet_xml = String::from_utf8_lossy(&self.zip.entry_data(&sheet_path)?).into_owned();
        self.ensure_shared_strings_loaded()?;
        if self.shared_strings_xml.is_some() {
            let refs = shared_string_ref_count(&sheet_xml);
            self.shared_strings_count = self.shared_strings_count.saturating_sub(refs);
            self.shared_strings_dirty = true;
        }

        let (mut col_styles, mut header_styles, mut date_style) =
            (HashMap::new(), HashMap::new(), None);
        if options.style_fallback != StyleFallback::General {
            let collected = collect_existing_styles(&sheet_xml);
            col_styles = collected.0;
            header_styles = collected.1;
            date_style = self.find_date_style_index()?;
        }

        let max_len = rows
            .iter()
            .map(|r| r.len())
            .chain(options.headers.as_ref().map(|h| h.len()))
            .max()
            .unwrap_or(0);
        let last_col = max_len as i64 - 1;
        let total_rows = rows.len() + options.headers.as_ref().map(|_| 1).unwrap_or(0);

        let new_sheet_data = self.build_sheet_data(
            &rows,
            options.headers.as_deref(),
            &col_styles,
            &header_styles,
            date_style,
        );
        let new_xml = patch_sheet_xml(&sheet_xml, &new_sheet_data, total_rows, last_col)?;
        self.zip.update_file(&sheet_path, new_xml.into_bytes());

        if self.shared_strings_xml.is_some() {
            self.commit_shared_strings()?;
        }

        let dim_ref = dim_ref_for(total_rows, last_col);
        self.update_pivot_caches(sheet_name, &dim_ref, rows.len())?;
        self.add_refresh_on_load_to_pivot_tables()?;
        Ok(())
    }

    /// Replace worksheet rows from a one-pass source without materialising
    /// the complete result set. Persist with [`XlsxUpdater::save_streaming`]
    /// to keep the final ZIP write streaming as well.
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
                "XlsxUpdater: styleFallback must be 'inherit' or 'general'.".into(),
            ));
        }
        let sheet_path = self.sheet_path(sheet_name)?;
        if !self.zip.has_entry(&sheet_path) {
            return Err(SpreadsheetError::InvalidFormat(format!(
                "XlsxUpdater: worksheet part missing for sheet \"{sheet_name}\"."
            )));
        }

        // Snapshot SST state for rollback on error.
        let snap = self.snapshot_sst();
        let old_stage = self.zip.staged_path(&sheet_path).map(|p| p.to_path_buf());
        let before_sst_bytes = if self.zip.has_entry("xl/sharedStrings.xml") {
            Some(self.zip.entry_data("xl/sharedStrings.xml")?)
        } else {
            None
        };
        let old_sheet_xml =
            String::from_utf8_lossy(&self.zip.entry_data(&sheet_path)?).into_owned();
        let rows_path = self.zip.temp_path(".rows")?;
        let output_path = self.zip.temp_path(".xml")?;
        let mut created = vec![rows_path.clone(), output_path.clone()];

        let result = self.replace_stream_inner(
            sheet_name,
            &sheet_path,
            &old_sheet_xml,
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
                if let Some(bytes) = before_sst_bytes {
                    self.zip.update_file("xl/sharedStrings.xml", bytes);
                }
                // Restore the previous staged sheet (or drop the new one).
                // The operation can fail before installing output_path, so
                // rollback must not discard the existing snapshot.
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
        old_sheet_xml: &str,
        rows: impl Iterator<Item = SpreadsheetResult<Vec<CellValue>>>,
        options: &ReplaceSheetDataOptions,
        rows_path: &Path,
        output_path: &Path,
    ) -> SpreadsheetResult<()> {
        self.ensure_shared_strings_loaded()?;
        if self.shared_strings_xml.is_some() {
            let refs = shared_string_ref_count(old_sheet_xml);
            self.shared_strings_count = self.shared_strings_count.saturating_sub(refs);
            self.shared_strings_dirty = true;
        }
        let (mut col_styles, mut header_styles, mut date_style) =
            (HashMap::new(), HashMap::new(), None);
        if options.style_fallback != StyleFallback::General {
            let collected = collect_existing_styles(old_sheet_xml);
            col_styles = collected.0;
            header_styles = collected.1;
            date_style = self.find_date_style_index()?;
        }

        let prefix = find_tag(old_sheet_xml, 0, "sheetdata")
            .map(|t| t.prefix)
            .unwrap_or_default();

        let rows_file = std::fs::File::create(rows_path)?;
        let mut writer = std::io::BufWriter::new(rows_file);
        use std::io::Write as _;

        let mut next_row: usize = 1;
        let mut width: usize = options.headers.as_ref().map(|h| h.len()).unwrap_or(0);
        let mut data_rows_seen: usize = 0;
        let mut kept_data_rows: usize = 0;
        let mut last_kept_end: u64 = 0;
        let mut pending_empty_width: usize = 0;
        let mut written: u64 = 0;

        if let Some(headers) = &options.headers {
            let mut cells = String::new();
            for (col, value) in headers.iter().enumerate() {
                cells.push_str(&self.string_cell(
                    value,
                    col,
                    next_row,
                    *header_styles.get(&col).unwrap_or(&0),
                ));
            }
            let xml = format!("<{prefix}row r=\"{next_row}\">{cells}</{prefix}row>");
            writer.write_all(xml.as_bytes())?;
            written += xml.len() as u64;
            last_kept_end = written;
            next_row += 1;
        }

        for row in rows {
            let row = row?;
            let mut cells = String::new();
            for (col, value) in row.iter().enumerate() {
                if matches!(value, CellValue::Empty) {
                    continue;
                }
                cells.push_str(&self.value_cell(value, col, next_row, &col_styles, date_style));
            }
            let xml = format!("<{prefix}row r=\"{next_row}\">{cells}</{prefix}row>");
            writer.write_all(xml.as_bytes())?;
            written += xml.len() as u64;
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

        // Truncate trailing empty rows.
        let f = std::fs::OpenOptions::new().write(true).open(rows_path)?;
        f.set_len(last_kept_end)?;
        drop(f);

        let total_rows = kept_data_rows + options.headers.as_ref().map(|_| 1).unwrap_or(0);
        let last_col = width.saturating_sub(1) as i64;
        let last_col = if total_rows == 0 { -1 } else { last_col };
        let dim = dim_ref_for(total_rows, last_col);
        patch_sheet_file(old_sheet_xml, rows_path, output_path, &dim)?;

        self.zip.stage_part(sheet_path, output_path.to_path_buf());
        self.commit_shared_strings()?;
        self.update_pivot_caches(sheet_name, &dim, kept_data_rows)?;
        self.add_refresh_on_load_to_pivot_tables()?;
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

    // -- workbook structure ------------------------------------------------

    fn load_workbook_structure(&mut self) -> SpreadsheetResult<()> {
        let rels = self
            .zip
            .entry_data("xl/_rels/workbook.xml.rels")
            .unwrap_or_default();
        let r_id_to_target = parse_relationships(&String::from_utf8_lossy(&rels));
        let wb = self.zip.entry_data("xl/workbook.xml").unwrap_or_default();
        let wb_xml = String::from_utf8_lossy(&wb).into_owned();
        self.uses_1904_date_system = xml_utils::uses_1904_date_system(&wb_xml);
        let mut search = 0;
        while let Some((start, end)) = xml_utils::find_open_tag(&wb_xml, search, "sheet") {
            let tag = &wb_xml[start..end];
            if let (Some(name), Some(r_id)) = (tag_attr(tag, "name"), tag_attr(tag, "r:id")) {
                let name = xml_utils::unescape_xml(&name);
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
            search = end;
            if search >= wb_xml.len() {
                break;
            }
        }
        for name in self.zip.entry_names() {
            if is_pivot_cache_def(&name) {
                let xml = String::from_utf8_lossy(&self.zip.entry_data(&name)?).into_owned();
                self.pivot_cache_defs
                    .push(PivotCacheDef { path: name, xml });
            } else if is_pivot_table_def(&name) {
                self.pivot_table_def_paths.push(name);
            }
        }
        Ok(())
    }

    // -- shared strings -----------------------------------------------------

    fn ensure_shared_strings_loaded(&mut self) -> SpreadsheetResult<()> {
        if !self.string_index_map.is_empty() || self.shared_strings_xml.is_some() {
            return Ok(());
        }
        if !self.zip.has_entry("xl/sharedStrings.xml") {
            self.shared_strings_xml = None;
            return Ok(());
        }
        let xml =
            String::from_utf8_lossy(&self.zip.entry_data("xl/sharedStrings.xml")?).into_owned();
        self.shared_strings_values = xml_utils::parse_shared_strings_xml(&xml);
        // count="N" on the <sst> open tag; default to values length.
        let count = find_tag(&xml, 0, "sst")
            .and_then(|t| tag_attr(&xml[t.start..t.end], "count"))
            .and_then(|c| c.parse::<usize>().ok())
            .unwrap_or(self.shared_strings_values.len());
        self.shared_strings_count = count;
        for (i, v) in self.shared_strings_values.iter().enumerate() {
            self.string_index_map.entry(v.clone()).or_insert(i);
        }
        self.shared_strings_xml = Some(xml);
        self.shared_strings_pending.clear();
        self.shared_strings_dirty = false;
        Ok(())
    }

    fn commit_shared_strings(&mut self) -> SpreadsheetResult<()> {
        if self.shared_strings_xml.is_none() || !self.shared_strings_dirty {
            return Ok(());
        }
        let mut xml = self.shared_strings_xml.clone().unwrap();
        let tag = find_tag(&xml, 0, "sst").ok_or_else(|| {
            SpreadsheetError::InvalidFormat(
                "XlsxUpdater: sharedStrings.xml has no sst element.".into(),
            )
        })?;
        let old_open = xml[tag.start..tag.end].to_owned();
        let with_count = replace_or_add_xml_attribute(
            &old_open,
            "count",
            &self.shared_strings_count.to_string(),
        );
        let new_open = replace_or_add_xml_attribute(
            &with_count,
            "uniqueCount",
            &self.shared_strings_values.len().to_string(),
        );
        xml.replace_range(tag.start..tag.end, &new_open);

        // Closing </sst> (with optional prefix) at the end.
        let prefix = tag.prefix;
        let closing_pat = format!("</{prefix}sst");
        let rel = xml.rfind(&closing_pat).ok_or_else(|| {
            SpreadsheetError::InvalidFormat(
                "XlsxUpdater: sharedStrings.xml has no closing sst element.".into(),
            )
        })?;
        let mut appended = String::new();
        for txt in &self.shared_strings_pending {
            let clean = xml_utils::escape_xml_text(txt);
            let preserve = !clean.is_empty()
                && (clean.starts_with(' ')
                    || clean.ends_with(' ')
                    || clean.contains(['\t', '\n', '\r']));
            let attr = if preserve {
                r#" xml:space="preserve""#
            } else {
                ""
            };
            appended.push_str(&format!(
                "<{prefix}si><{prefix}t{attr}>{clean}</{prefix}t></{prefix}si>"
            ));
        }
        xml.insert_str(rel, &appended);
        self.zip
            .update_file("xl/sharedStrings.xml", xml.as_bytes().to_vec());
        self.shared_strings_xml = Some(xml);
        self.shared_strings_pending.clear();
        self.shared_strings_dirty = false;
        Ok(())
    }

    // -- styles ---------------------------------------------------------------

    fn find_date_style_index(&self) -> SpreadsheetResult<Option<usize>> {
        if !self.zip.has_entry("xl/styles.xml") {
            return Ok(None);
        }
        let xml = String::from_utf8_lossy(&self.zip.entry_data("xl/styles.xml")?).into_owned();
        let mut customs = HashSet::new();
        let mut search = 0;
        while let Some((start, end)) = xml_utils::find_open_tag(&xml, search, "numFmt") {
            let tag = &xml[start..end];
            if let (Some(id), Some(code)) = (tag_attr(tag, "numFmtId"), tag_attr(tag, "formatCode"))
            {
                if let Ok(id) = id.parse::<u16>() {
                    let lc = code.to_lowercase();
                    if lc.contains("yy")
                        || lc.contains("mm")
                        || lc.contains("dd")
                        || lc.contains("h:mm")
                    {
                        customs.insert(id);
                    }
                }
            }
            search = end;
            if search >= xml.len() {
                break;
            }
        }
        let xfs_start = match xml_utils::find_open_tag(&xml, 0, "cellXfs") {
            Some((_, i)) => i,
            None => return Ok(None),
        };
        let xfs_end = xml_utils::find_close_tag(&xml, xfs_start, "cellXfs").ok_or_else(|| {
            SpreadsheetError::InvalidFormat("XlsxUpdater: cellXfs is not closed.".into())
        })?;
        let section = &xml[xfs_start..xfs_end];
        let mut s = 0;
        let mut index = 0;
        while let Some((xs, xe)) = xml_utils::find_open_tag(section, s, "xf") {
            let nf = tag_attr(&section[xs..xe], "numFmtId")
                .and_then(|v| v.parse::<u16>().ok())
                .unwrap_or(0);
            if (14..=22).contains(&nf) || (45..=47).contains(&nf) || customs.contains(&nf) {
                return Ok(Some(index));
            }
            index += 1;
            s = xe;
            if s >= section.len() {
                break;
            }
        }
        Ok(None)
    }

    // -- cell serialization -----------------------------------------------------

    fn build_sheet_data(
        &mut self,
        rows: &[Vec<CellValue>],
        headers: Option<&[String]>,
        col_styles: &HashMap<usize, usize>,
        header_styles: &HashMap<usize, usize>,
        date_style: Option<usize>,
    ) -> String {
        let mut out = String::new();
        let mut _row_num = 0;
        if let Some(h) = headers {
            _row_num = 1;
            out.push_str(r#"<row>"#);
            for (c, header) in h.iter().enumerate() {
                let style = header_styles.get(&c).copied().unwrap_or(0);
                out.push_str(&self.string_cell(header, c, 0, style));
            }
            out.push_str("</row>");
        }
        for row in rows {
            _row_num += 1;
            out.push_str(r#"<row>"#);
            for (c, raw) in row.iter().enumerate() {
                if matches!(raw, CellValue::Empty) {
                    continue;
                }
                out.push_str(&self.value_cell(raw, c, 0, col_styles, date_style));
            }
            out.push_str("</row>");
        }
        out
    }

    fn value_cell(
        &mut self,
        raw: &CellValue,
        col: usize,
        _row_num: usize,
        col_styles: &HashMap<usize, usize>,
        date_style: Option<usize>,
    ) -> String {
        let val = borrow_cell(raw);
        let col_style = col_styles.get(&col).copied().unwrap_or(0);
        let style_attr = if col_style > 0 {
            format!(r#" s="{col_style}""#)
        } else {
            String::new()
        };
        match val {
            CellValueRef::Empty => String::new(),
            CellValueRef::Number(n) => {
                if n.is_finite() {
                    format!(r#"<c{style_attr}><v>{n}</v></c>"#)
                } else {
                    self.string_cell(&n.to_string(), col, 0, col_style)
                }
            }
            CellValueRef::Integer(n) => {
                format!(r#"<c{style_attr}><v>{n}</v></c>"#)
            }
            CellValueRef::Boolean(b) => {
                let v = if b { 1 } else { 0 };
                format!(r#"<c t="b"{style_attr}><v>{v}</v></c>"#)
            }
            CellValueRef::DateTime(d) => {
                let oa = datetime_to_excel_serial(&d, self.uses_1904_date_system);
                if oa.is_finite() {
                    let style = if col_style > 0 {
                        col_style
                    } else {
                        date_style.unwrap_or(0)
                    };
                    let attr = if style > 0 {
                        format!(r#" s="{style}""#)
                    } else {
                        String::new()
                    };
                    format!(r#"<c{attr}><v>{oa}</v></c>"#)
                } else {
                    self.string_cell(&d.to_string(), col, 0, col_style)
                }
            }
            CellValueRef::Text(s) => self.string_cell(s, col, 0, col_style),
        }
    }

    fn string_cell(&mut self, text: &str, _col: usize, _row_num: usize, style: usize) -> String {
        let style_attr = if style > 0 {
            format!(r#" s="{style}""#)
        } else {
            String::new()
        };
        if self.shared_strings_xml.is_none() {
            let clean = xml_utils::escape_xml_text(text);
            let preserve = !clean.is_empty()
                && (clean.starts_with(' ')
                    || clean.ends_with(' ')
                    || clean.contains(['\t', '\n', '\r']));
            let attr = if preserve {
                r#" xml:space="preserve""#
            } else {
                ""
            };
            return format!(r#"<c t="inlineStr"{style_attr}><is><t{attr}>{clean}</t></is></c>"#);
        }
        let index = match self.string_index_map.get(text) {
            Some(&i) => i,
            None => {
                let i = self.shared_strings_values.len();
                let owned = text.to_owned();
                self.shared_strings_values.push(owned.clone());
                self.string_index_map.insert(owned.clone(), i);
                self.shared_strings_pending.push(owned);
                i
            }
        };
        self.shared_strings_count += 1;
        self.shared_strings_dirty = true;
        format!(r#"<c t="s"{style_attr}><v>{index}</v></c>"#)
    }

    // -- pivot tables -------------------------------------------------------------

    fn update_pivot_caches(
        &mut self,
        sheet_name: &str,
        dim_ref: &str,
        record_count: usize,
    ) -> SpreadsheetResult<()> {
        for cache in std::mem::take(&mut self.pivot_cache_defs) {
            let mut xml = cache.xml;
            if let Some(ws_pos) = find_open_tag(&xml, "<worksheetSource") {
                let tag = xml[ws_pos.0..ws_pos.1].to_owned();
                if let Some(sheet) = tag_attr(&tag, "sheet") {
                    if sheet == sheet_name {
                        let new_tag = replace_ref_attr(&tag, dim_ref);
                        xml.replace_range(ws_pos.0..ws_pos.1, &new_tag);
                        if let Some(root) = find_open_tag(&xml, "<pivotCacheDefinition") {
                            let mut root_tag = xml[root.0..root.1].to_owned();
                            if root_tag.contains("recordCount=") {
                                root_tag = replace_int_attr(&root_tag, "recordCount", record_count);
                            } else if !root_tag.ends_with("/>") {
                                root_tag = root_tag.trim_end_matches('>').to_owned()
                                    + &format!(r#" recordCount="{record_count}">"#);
                            }
                            xml.replace_range(root.0..root.1, &root_tag);
                        }
                        self.zip.update_file(&cache.path, xml.as_bytes().to_vec());
                    }
                }
            }
            self.pivot_cache_defs.push(PivotCacheDef {
                path: cache.path,
                xml,
            });
        }
        Ok(())
    }

    fn add_refresh_on_load_to_pivot_tables(&mut self) -> SpreadsheetResult<()> {
        for path in self.pivot_table_def_paths.clone() {
            let bytes = match self.zip.entry_data(&path) {
                Ok(b) => b,
                Err(_) => continue,
            };
            let mut xml = String::from_utf8_lossy(&bytes).into_owned();
            let pos = match find_open_tag(&xml, "<pivotTableDefinition") {
                Some(p) => p,
                None => continue,
            };
            let tag = xml[pos.0..pos.1].to_owned();
            if tag.contains("refreshOnLoad") {
                continue;
            }
            let new_tag = if tag.ends_with("/>") {
                format!("{} refreshOnLoad=\"1\"/>", tag[..tag.len() - 2].to_owned())
            } else {
                format!("{} refreshOnLoad=\"1\">", tag[..tag.len() - 1].to_owned())
            };
            xml.replace_range(pos.0..pos.1, &new_tag);
            self.zip.update_file(&path, xml.as_bytes().to_vec());
        }
        Ok(())
    }

    // -- SST snapshot / rollback ----------------------------------------------------

    fn snapshot_sst(&self) -> SstSnapshot {
        SstSnapshot {
            values_len: self.shared_strings_values.len(),
            pending_len: self.shared_strings_pending.len(),
            count: self.shared_strings_count,
            dirty: self.shared_strings_dirty,
            xml: self.shared_strings_xml.clone(),
            index_map: self.string_index_map.clone(),
        }
    }

    fn restore_sst(&mut self, snap: SstSnapshot) {
        self.shared_strings_values.truncate(snap.values_len);
        self.shared_strings_pending.truncate(snap.pending_len);
        self.shared_strings_count = snap.count;
        self.shared_strings_dirty = snap.dirty;
        self.shared_strings_xml = snap.xml;
        self.string_index_map = snap.index_map;
    }
}

struct SstSnapshot {
    values_len: usize,
    pending_len: usize,
    count: usize,
    dirty: bool,
    xml: Option<String>,
    index_map: HashMap<String, usize>,
}

impl Drop for XlsxUpdater {
    fn drop(&mut self) {
        self.zip.dispose();
    }
}

/// Explicit name for updating macro-enabled XLSM packages (VBA parts pass through untouched).
pub type XlsmUpdater = XlsxUpdater;

// -- free helpers ---------------------------------------------------------------

fn dim_ref_for(total_rows: usize, last_col: i64) -> String {
    if total_rows > 0 && last_col >= 0 {
        format!(
            "A1:{}{}",
            xml_utils::column_index_to_letter(last_col as usize),
            total_rows
        )
    } else {
        "A1".to_string()
    }
}

fn is_pivot_cache_def(name: &str) -> bool {
    name.starts_with("xl/pivotCache/pivotCacheDefinition") && name.ends_with(".xml") && {
        let mid = &name["xl/pivotCache/pivotCacheDefinition".len()..name.len() - 4];
        mid.chars().all(|c| c.is_ascii_digit())
    }
}

fn is_pivot_table_def(name: &str) -> bool {
    name.starts_with("xl/pivotTables/pivotTable") && name.ends_with(".xml") && {
        let mid = &name["xl/pivotTables/pivotTable".len()..name.len() - 4];
        mid.chars().all(|c| c.is_ascii_digit())
    }
}

/// Attribute value (unescaped) of an XML open tag, or None.
fn tag_attr(tag: &str, name: &str) -> Option<String> {
    xml_utils::tag_attribute(tag, name)
}

struct FoundTag {
    start: usize,
    end: usize, // just past '>'
    prefix: String,
}

/// Find the first open (or self-closing) tag with the given local name,
/// case-insensitively, supporting namespace prefixes.
fn find_tag(xml: &str, from: usize, local: &str) -> Option<FoundTag> {
    let bytes = xml.as_bytes();
    let mut i = from.min(xml.len());
    while let Some(rel) = xml[i..].find('<') {
        let pos = i + rel;
        let next = bytes.get(pos + 1).copied().unwrap_or(0);
        if next == b'/' || next == b'!' || next == b'?' {
            i = pos + 1;
            continue;
        }
        let end = xml[pos..].find('>')? + pos;
        let inner = xml[pos + 1..end].trim();
        let raw = inner.split_whitespace().next().unwrap_or("");
        let raw = raw.trim_end_matches('/');
        let mut parts = raw.split(':');
        let (prefix, name) = match (parts.next(), parts.next()) {
            (Some(a), Some(b)) => (format!("{a}:"), b.to_owned()),
            (Some(a), None) => (String::new(), a.to_owned()),
            _ => {
                i = pos + 1;
                continue;
            }
        };
        if name.eq_ignore_ascii_case(local) {
            return Some(FoundTag {
                start: pos,
                end: end + 1,
                prefix,
            });
        }
        i = pos + 1;
        if i >= xml.len() {
            break;
        }
    }
    None
}

/// Find `<local ...>` open tag position (start, end) — byte offsets.
fn find_open_tag(xml: &str, pat: &str) -> Option<(usize, usize)> {
    let local = pat
        .trim_start_matches('<')
        .split_whitespace()
        .next()
        .unwrap_or(pat);
    find_tag(xml, 0, local).map(|t| (t.start, t.end))
}

fn replace_or_add_xml_attribute(tag: &str, name: &str, value: &str) -> String {
    let escaped = xml_utils::escape_xml_text(value);
    // existing attribute (case-insensitive)?
    let lower = tag.to_lowercase();
    let needle = format!(" {name}=");
    if lower.find(&needle.to_lowercase()).is_some() {
        // replace value preserving quote style
        if let Some(abs) = find_attr_value_span(tag, name) {
            let mut out = tag[..abs.0].to_owned();
            out.push_str(&escaped);
            out.push_str(&tag[abs.1..]);
            return out;
        }
    }
    let insertion = format!(r#" {name}="{escaped}""#);
    if let Some(stripped) = tag.strip_suffix("/>") {
        format!("{stripped}{insertion}/>")
    } else {
        format!("{}{insertion}>", &tag[..tag.len() - 1])
    }
}

fn find_attr_value_span(tag: &str, name: &str) -> Option<(usize, usize)> {
    let needle = format!("{name}=");
    let mut search = 0;
    while let Some(rel) = tag.to_lowercase()[search..].find(&needle.to_lowercase()) {
        let abs = search + rel;
        let rest = &tag[abs + needle.len()..];
        let quote = rest.chars().next()?;
        if quote != '"' && quote != '\'' {
            search = abs + needle.len();
            continue;
        }
        let inner = &rest[1..];
        let end = inner.find(quote)? + abs + needle.len() + 1;
        return Some((abs + needle.len() + 1, end));
    }
    None
}

fn replace_ref_attr(tag: &str, dim_ref: &str) -> String {
    replace_or_add_xml_attribute(tag, "ref", dim_ref)
}

fn replace_int_attr(tag: &str, name: &str, value: usize) -> String {
    // replace `name="digits"` occurrences
    let mut out = tag.to_owned();
    let lower = tag.to_lowercase();
    let needle = format!("{name}=\"");
    if let Some(rel) = lower.find(&needle.to_lowercase()) {
        let start = rel + needle.len();
        if let Some(end_rel) = out[start..].find('"') {
            out.replace_range(start..start + end_rel, &value.to_string());
        }
    }
    out
}

/// Count `<c ... t="s" ...>` cell tags (shared-string references).
fn shared_string_ref_count(sheet_xml: &str) -> usize {
    let mut count = 0;
    let mut search = 0;
    while let Some(rel) = sheet_xml[search..].find("<c") {
        let pos = search + rel;
        let after = sheet_xml.as_bytes().get(pos + 2).copied().unwrap_or(b' ');
        if after != b' ' && after != b'>' && after != b'/' {
            search = pos + 2;
            continue;
        }
        let end = match sheet_xml[pos..].find('>') {
            Some(e) => pos + e,
            None => break,
        };
        let tag = &sheet_xml[pos..end];
        if has_t_s_attr(tag) {
            count += 1;
        }
        search = end + 1;
        if search >= sheet_xml.len() {
            break;
        }
    }
    count
}

fn has_t_s_attr(tag: &str) -> bool {
    let lower = tag.to_lowercase();
    let mut search = 0;
    while let Some(rel) = lower[search..].find("t=") {
        let abs = search + rel;
        let rest = &tag[abs + 2..];
        let quote = rest.chars().next().unwrap_or(' ');
        if quote != '"' && quote != '\'' {
            search = abs + 2;
            continue;
        }
        let inner = &rest[1..];
        if let Some(end) = inner.find(quote) {
            if inner[..end].trim() == "s" {
                return true;
            }
        }
        search = abs + 2;
    }
    false
}

/// Dominant style per column (rows > 1) + original header styles (row 1).
fn collect_existing_styles(sheet_xml: &str) -> (HashMap<usize, usize>, HashMap<usize, usize>) {
    let mut data_counts: HashMap<usize, HashMap<usize, usize>> = HashMap::new();
    let mut header_styles: HashMap<usize, usize> = HashMap::new();
    let sd_start = match sheet_xml.find("<sheetData>") {
        Some(i) => i,
        None => return (HashMap::new(), header_styles),
    };
    let sd_end = match sheet_xml[sd_start..].find("</sheetData>") {
        Some(e) => sd_start + e,
        None => return (HashMap::new(), header_styles),
    };
    let region = &sheet_xml[sd_start..sd_end];
    let mut search = 0;
    while let Some(rel) = region[search..].find("<c") {
        let pos = search + rel;
        let after = region.as_bytes().get(pos + 2).copied().unwrap_or(b' ');
        if after != b' ' && after != b'>' && after != b'/' {
            search = pos + 2;
            continue;
        }
        let end = match region[pos..].find('>') {
            Some(e) => pos + e,
            None => break,
        };
        let tag = &region[pos..end];
        if let Some(cell_ref) = tag_attr(tag, "r") {
            let letters: String = cell_ref
                .chars()
                .take_while(|c| c.is_ascii_alphabetic())
                .collect();
            let digits: String = cell_ref
                .chars()
                .skip_while(|c| c.is_ascii_alphabetic())
                .collect();
            if let (Ok(col), Ok(row)) = (try_col_index(&letters), digits.parse::<usize>()) {
                if let Some(s) = tag_attr(tag, "s").and_then(|v| v.parse::<usize>().ok()) {
                    if s > 0 {
                        if row == 1 {
                            header_styles.entry(col).or_insert(s);
                        } else {
                            data_counts
                                .entry(col)
                                .or_default()
                                .entry(s)
                                .and_modify(|c| *c += 1)
                                .or_insert(1);
                        }
                    }
                }
            }
        }
        search = end + 1;
        if search >= region.len() {
            break;
        }
    }
    let mut data_styles = HashMap::new();
    for (col, counts) in data_counts {
        if let Some((&best, _)) = counts.iter().max_by_key(|(_, &c)| c) {
            data_styles.insert(col, best);
        }
    }
    (data_styles, header_styles)
}

fn try_col_index(letters: &str) -> Result<usize, ()> {
    if letters.is_empty() || !letters.chars().all(|c| c.is_ascii_alphabetic()) {
        return Err(());
    }
    Ok(xml_utils::column_letter_to_index(&letters.to_uppercase()))
}

/// Patch worksheet XML: dimension, sheetData content, autoFilter range.
fn patch_sheet_xml(
    sheet_xml: &str,
    new_sheet_data: &str,
    total_rows: usize,
    last_col: i64,
) -> SpreadsheetResult<String> {
    let mut xml = sheet_xml.to_owned();
    let dim_ref = dim_ref_for(total_rows, last_col);

    // Dimension element: update or insert after <worksheet ...>.
    if let Some(t) = find_tag(&xml, 0, "dimension") {
        let old = xml[t.start..t.end].to_owned();
        let new_tag = replace_or_add_xml_attribute(&old, "ref", &dim_ref);
        xml.replace_range(t.start..t.end, &new_tag);
    } else if let Some(t) = find_tag(&xml, 0, "worksheet") {
        xml.insert_str(t.end, &format!(r#"<dimension ref="{dim_ref}"/>"#));
    }

    // sheetData content replace (or expand self-closing).
    let sd_open = "<sheetData>";
    if let Some(start) = xml.find(sd_open) {
        let end = xml[start..].find("</sheetData>").ok_or_else(|| {
            SpreadsheetError::InvalidFormat(
                "XlsxUpdater: <sheetData> element not closed in worksheet XML.".into(),
            )
        })? + start;
        xml.replace_range(start + sd_open.len()..end, new_sheet_data);
    } else if let Some(t) = find_self_closing(&xml, "sheetdata") {
        let replacement = format!("<sheetData>{new_sheet_data}</sheetData>");
        xml.replace_range(t.0..t.1, &replacement);
    } else {
        // namespaced variants (prefix-aware)
        let t = find_tag(&xml, 0, "sheetdata").ok_or_else(|| {
            SpreadsheetError::InvalidFormat(
                "XlsxUpdater: <sheetData> element not found in worksheet XML.".into(),
            )
        })?;
        let open_tag = xml[t.start..t.end].to_owned();
        if open_tag.ends_with("/>") {
            let prefix = t.prefix.clone();
            let replacement = format!("<{prefix}sheetData>{new_sheet_data}</{prefix}sheetData>");
            xml.replace_range(t.start..t.end, &replacement);
        } else {
            let close_pat = format!("</{}sheetData", t.prefix);
            let rel = xml[t.end..].find(&close_pat).ok_or_else(|| {
                SpreadsheetError::InvalidFormat(
                    "XlsxUpdater: <sheetData> element not closed in worksheet XML.".into(),
                )
            })?;
            let close_start = t.end + rel;
            xml.replace_range(t.end..close_start, new_sheet_data);
        }
    }

    // autoFilter range sync.
    if total_rows > 0 && last_col >= 0 {
        if let Some(t) = find_tag(&xml, 0, "autoFilter") {
            let old = xml[t.start..t.end].to_owned();
            if old.contains("ref=") {
                let new_tag = replace_or_add_xml_attribute(&old, "ref", &dim_ref);
                xml.replace_range(t.start..t.end, &new_tag);
            }
        }
    }
    Ok(xml)
}

fn find_self_closing(xml: &str, local: &str) -> Option<(usize, usize)> {
    let mut search = 0;
    while let Some(t) = find_tag(xml, search, local) {
        if xml[t.start..t.end].ends_with("/>") {
            return Some((t.start, t.end));
        }
        search = t.end;
        if search >= xml.len() {
            break;
        }
    }
    None
}

/// Splice staged row fragments into a copy of the worksheet XML on disk.
fn patch_sheet_file(
    sheet_xml: &str,
    rows_path: &Path,
    output_path: &Path,
    dimension_ref: &str,
) -> SpreadsheetResult<()> {
    struct Replacement {
        start: usize,
        end: usize,
        data: Option<Vec<u8>>,
        rows: bool,
        suffix: Option<Vec<u8>>,
    }
    let mut replacements: Vec<Replacement> = Vec::new();

    // sheetData region (prefix-aware)
    if let Some(t) = find_tag(sheet_xml, 0, "sheetdata") {
        let open = sheet_xml[t.start..t.end].to_owned();
        if open.ends_with("/>") {
            let prefix = t.prefix.clone();
            replacements.push(Replacement {
                start: t.start,
                end: t.end,
                data: Some(format!("<{prefix}sheetData>").into_bytes()),
                rows: true,
                suffix: Some(format!("</{prefix}sheetData>").into_bytes()),
            });
        } else {
            let close_pat = format!("</{}sheetData", t.prefix);
            let rel = sheet_xml[t.end..].find(&close_pat).ok_or_else(|| {
                SpreadsheetError::InvalidFormat(
                    "XlsxUpdater: <sheetData> element not closed in worksheet XML.".into(),
                )
            })?;
            replacements.push(Replacement {
                start: t.end,
                end: t.end + rel,
                data: None,
                rows: true,
                suffix: None,
            });
        }
    } else {
        return Err(SpreadsheetError::InvalidFormat(
            "XlsxUpdater: <sheetData> element not found in worksheet XML.".into(),
        ));
    }

    if let Some(t) = find_tag(sheet_xml, 0, "dimension") {
        let old = sheet_xml[t.start..t.end].to_owned();
        replacements.push(Replacement {
            start: t.start,
            end: t.end,
            data: Some(replace_or_add_xml_attribute(&old, "ref", dimension_ref).into_bytes()),
            rows: false,
            suffix: None,
        });
    } else if let Some(t) = find_tag(sheet_xml, 0, "worksheet") {
        let prefix = t.prefix.clone();
        replacements.push(Replacement {
            start: t.end,
            end: t.end,
            data: Some(format!(r#"<{prefix}dimension ref="{dimension_ref}"/>"#).into_bytes()),
            rows: false,
            suffix: None,
        });
    }

    if let Some(t) = find_tag(sheet_xml, 0, "autoFilter") {
        let old = sheet_xml[t.start..t.end].to_owned();
        replacements.push(Replacement {
            start: t.start,
            end: t.end,
            data: Some(replace_or_add_xml_attribute(&old, "ref", dimension_ref).into_bytes()),
            rows: false,
            suffix: None,
        });
    }

    replacements.sort_by_key(|r| (r.start, r.end));
    let source = sheet_xml.as_bytes();
    let out = std::fs::File::create(output_path)?;
    let mut out = std::io::BufWriter::new(out);
    use std::io::Write as _;
    let mut cursor = 0;
    for r in &replacements {
        if r.start < cursor {
            return Err(SpreadsheetError::InvalidFormat(
                "XlsxUpdater: overlapping worksheet replacements.".into(),
            ));
        }
        out.write_all(&source[cursor..r.start])?;
        if let Some(data) = &r.data {
            out.write_all(data)?;
        }
        if r.rows {
            let mut f = std::fs::File::open(rows_path)?;
            std::io::copy(&mut f, &mut out)?;
        }
        if let Some(suffix) = &r.suffix {
            out.write_all(suffix)?;
        }
        cursor = r.end;
    }
    out.write_all(&source[cursor..])?;
    out.flush()?;
    Ok(())
}

#[allow(dead_code)]
fn unused_marker() {}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::{XlsxReader, XlsxWriter};

    fn sample_workbook(path: &Path) {
        let mut w = XlsxWriter::create(path).unwrap();
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

    #[test]
    fn replace_preserves_other_sheets() {
        let dir = tempfile::tempdir().unwrap();
        let path = dir.path().join("report.xlsx");
        sample_workbook(&path);

        let mut u = XlsxUpdater::open(&path).unwrap();
        let names = u.sheet_names();
        assert!(names.contains(&"data1".to_string()));
        u.replace_sheet_data(
            "data1",
            vec![vec![CellValue::Text("Zed".into()), CellValue::Integer(99)]],
            Some(ReplaceSheetDataOptions {
                headers: Some(vec!["NAME".to_string(), "AGE".to_string()]),
                style_fallback: StyleFallback::General,
            }),
        )
        .unwrap();
        let out = dir.path().join("report_new.xlsx");
        u.save(Some(&out)).unwrap();

        let mut r = XlsxReader::new();
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
        let path = dir.path().join("s.xlsx");
        sample_workbook(&path);
        let mut u = XlsxUpdater::open(&path).unwrap();
        let rows = vec![
            vec![CellValue::Integer(1)],
            vec![CellValue::Empty],
            vec![CellValue::Empty],
        ];
        u.replace_sheet_data_stream("data1", rows.into_iter(), None)
            .unwrap();
        let out = dir.path().join("s_new.xlsx");
        u.save_streaming(Some(&out)).unwrap();
        let mut r = XlsxReader::new();
        r.open(&out, true).unwrap();
        let mut n = 0;
        while r.read().unwrap() {
            n += 1;
        }
        // 1 data row (no headers passed; trailing empties trimmed)
        assert_eq!(n, 1);
    }

    #[test]
    fn missing_sheet_errors() {
        let dir = tempfile::tempdir().unwrap();
        let path = dir.path().join("m.xlsx");
        sample_workbook(&path);
        let mut u = XlsxUpdater::open(&path).unwrap();
        assert!(u.replace_sheet_data("nope", vec![], None).is_err());
    }

    fn mark_xlsx_as_1904(path: &Path) {
        let mut zip = crate::zip_store::ZipStore::open(path).unwrap();
        let mut workbook = String::from_utf8(zip.entry_data("xl/workbook.xml").unwrap()).unwrap();
        let marker = "<workbookPr ";
        let start = workbook.find(marker).unwrap();
        workbook.insert_str(start + marker.len(), "date1904=\"1\" ");
        zip.update_file("xl/workbook.xml", workbook.into_bytes());
        zip.save(path).unwrap();
    }

    #[test]
    fn updater_writes_1904_date_serials() {
        let dir = tempfile::tempdir().unwrap();
        let path = dir.path().join("date1904.xlsx");
        sample_workbook(&path);
        mark_xlsx_as_1904(&path);

        let expected = chrono::NaiveDate::from_ymd_opt(2025, 2, 3)
            .unwrap()
            .and_hms_opt(12, 30, 0)
            .unwrap();
        let mut updater = XlsxUpdater::open(&path).unwrap();
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
        let output = dir.path().join("date1904-out.xlsx");
        updater.save(Some(&output)).unwrap();

        let zip = crate::zip_store::ZipStore::open(&output).unwrap();
        let sheet = String::from_utf8(zip.entry_data("xl/worksheets/sheet1.xml").unwrap()).unwrap();
        let serial = crate::datetime_to_excel_serial(&expected, true).to_string();
        assert!(
            sheet.contains(&format!("<v>{serial}</v>")),
            "sheet XML: {sheet}"
        );

        let mut reader = XlsxReader::new();
        reader.open(&output, true).unwrap();
        assert!(reader.uses_1904_date_system);
        assert!(reader.read().unwrap());
        assert_eq!(reader.get_value(0), CellValue::DateTime(expected));
    }
}
