//! Shared ZIP backing store for the updaters.
//!
//! The TypeScript updaters pair `AdmZip` (random access) with staged
//! temporary files (streaming). This store mirrors that: entries live in
//! memory, while streaming replacements are staged as temp files and only
//! materialized on `save` / `to_buffer`, or streamed entry-by-entry on
//! `save_streaming`.

use crate::atomic_file::write_buffer_atomically;
use crate::error::{SpreadsheetError, SpreadsheetResult};
use std::collections::HashMap;
use std::fs::File;
use std::io::{Read, Write};
use std::path::{Path, PathBuf};

pub struct ZipStore {
    source_path: PathBuf,
    files: HashMap<String, Vec<u8>>,
    order: Vec<String>,
    staged: HashMap<String, PathBuf>,
    tempdir: Option<tempfile::TempDir>,
}

impl ZipStore {
    pub fn open(path: &Path) -> SpreadsheetResult<Self> {
        if !path.exists() {
            return Err(SpreadsheetError::Io(std::io::Error::new(
                std::io::ErrorKind::NotFound,
                format!("file not found: {}", path.display()),
            )));
        }
        let file = File::open(path)?;
        let mut zip = zip::ZipArchive::new(file)?;
        let mut files = HashMap::with_capacity(zip.len());
        let mut order = Vec::with_capacity(zip.len());
        for i in 0..zip.len() {
            let mut entry = zip.by_index(i)?;
            let name = entry.name().to_owned();
            let size = usize::try_from(entry.size()).map_err(|_| {
                SpreadsheetError::InvalidFormat("ZIP entry is too large for this platform".into())
            })?;
            let mut buf = Vec::with_capacity(size);
            entry.read_to_end(&mut buf)?;
            order.push(name.clone());
            files.insert(name, buf);
        }
        Ok(Self {
            source_path: path.to_path_buf(),
            files,
            order,
            staged: HashMap::new(),
            tempdir: None,
        })
    }

    pub fn source_path(&self) -> &Path {
        &self.source_path
    }

    pub fn has_entry(&self, name: &str) -> bool {
        self.files.contains_key(name) || self.staged.contains_key(name)
    }

    pub fn entry_names(&self) -> Vec<String> {
        let mut names = self.order.clone();
        for name in self.staged.keys() {
            if !self.files.contains_key(name) {
                names.push(name.clone());
            }
        }
        names
    }

    /// Current bytes of a member, preferring the staged replacement.
    pub fn entry_data(&self, name: &str) -> SpreadsheetResult<Vec<u8>> {
        if let Some(staged) = self.staged.get(name) {
            return Ok(std::fs::read(staged)?);
        }
        self.files
            .get(name)
            .cloned()
            .ok_or_else(|| SpreadsheetError::InvalidFormat(format!("ZIP member missing: {name}")))
    }

    /// Replace a member in memory and drop any staged file for it.
    pub fn update_file(&mut self, name: &str, data: Vec<u8>) {
        if !self.files.contains_key(name) && !self.order.iter().any(|entry| entry == name) {
            self.order.push(name.to_string());
        }
        self.files.insert(name.to_string(), data);
        self.discard_staged(name);
    }

    pub fn staged_path(&self, name: &str) -> Option<&Path> {
        self.staged.get(name).map(|p| p.as_path())
    }

    pub fn stage_part(&mut self, name: &str, file_path: PathBuf) {
        // Keep the previous path alive until the caller has committed the
        // replacement. Streaming updater operations use it as their rollback
        // snapshot when a later step fails.
        self.staged.insert(name.to_string(), file_path);
    }

    pub fn discard_staged(&mut self, name: &str) {
        if let Some(path) = self.staged.remove(name) {
            let _ = std::fs::remove_file(path);
        }
    }

    /// Restore a staged replacement after an updater operation fails.
    ///
    /// The replacement may fail before it reaches [`stage_part`], in which
    /// case the current path is already the previous snapshot and must not
    /// be removed.
    pub fn rollback_staged(&mut self, name: &str, previous: Option<PathBuf>, replacement: &Path) {
        let current = self.staged.get(name).cloned();
        if current.as_deref() == Some(replacement) {
            self.discard_staged(name);
        }
        if let Some(previous) = previous {
            if !self.staged.contains_key(name) {
                self.staged.insert(name.to_string(), previous);
            }
        }
    }

    /// Read staged files back into memory and delete them.
    pub fn materialize_staged(&mut self) -> SpreadsheetResult<()> {
        let staged = std::mem::take(&mut self.staged);
        let mut materialized = Vec::with_capacity(staged.len());
        for (name, path) in &staged {
            let data = match std::fs::read(path) {
                Ok(data) => data,
                Err(error) => {
                    self.staged = staged;
                    return Err(error.into());
                }
            };
            materialized.push((name.clone(), data));
        }
        for (name, data) in materialized {
            if !self.files.contains_key(&name) {
                self.order.push(name.clone());
            }
            self.files.insert(name, data);
        }
        for (_, path) in staged {
            let _ = std::fs::remove_file(path);
        }
        Ok(())
    }

    fn zip_options() -> zip::write::SimpleFileOptions {
        zip::write::SimpleFileOptions::default()
            .compression_method(zip::CompressionMethod::Deflated)
            // Match the TS reference (archiver/compress-commons default).
            .compression_level(Some(1))
            .last_modified_time(crate::writer_helpers::deterministic_zip_timestamp())
    }

    fn write_entries_to<W: Write + std::io::Seek>(
        &self,
        zip: &mut zip::ZipWriter<W>,
        stream_staged: bool,
    ) -> SpreadsheetResult<()> {
        let options = Self::zip_options();
        for name in self.entry_names() {
            zip.start_file(&name, options)?;
            if stream_staged {
                if let Some(staged) = self.staged.get(&name) {
                    let mut f = File::open(staged)?;
                    let mut buf = [0u8; 1024 * 1024];
                    loop {
                        let n = f.read(&mut buf)?;
                        if n == 0 {
                            break;
                        }
                        zip.write_all(&buf[..n])?;
                    }
                    continue;
                }
            }
            if let Some(data) = self.files.get(&name) {
                zip.write_all(data)?;
            }
        }
        Ok(())
    }

    /// The full workbook as an in-memory ZIP archive.
    pub fn buffer(&mut self) -> SpreadsheetResult<Vec<u8>> {
        self.materialize_staged()?;
        let cursor = std::io::Cursor::new(Vec::new());
        let mut zip = zip::ZipWriter::new(cursor);
        self.write_entries_to(&mut zip, false)?;
        let cursor = zip.finish()?;
        Ok(cursor.into_inner())
    }

    /// Materialize staged parts and atomically replace `target`.
    pub fn save(&mut self, target: &Path) -> SpreadsheetResult<()> {
        self.materialize_staged()?;
        let cursor = std::io::Cursor::new(Vec::new());
        let mut zip = zip::ZipWriter::new(cursor);
        self.write_entries_to(&mut zip, false)?;
        let cursor = zip.finish()?;
        write_buffer_atomically(cursor.get_ref(), target)
    }

    /// Rebuild the ZIP without materializing staged parts: unchanged
    /// members stream from memory one at a time, staged members stream
    /// from their temp files. The destination is replaced atomically.
    pub fn save_streaming(&mut self, target: &Path) -> SpreadsheetResult<()> {
        let target_abs: PathBuf = if target.is_absolute() {
            target.to_path_buf()
        } else {
            std::env::current_dir()
                .map(|c| c.join(target))
                .unwrap_or_else(|_| target.to_path_buf())
        };
        if let Some(parent) = target_abs.parent() {
            if !parent.as_os_str().is_empty() {
                std::fs::create_dir_all(parent)?;
            }
        }
        let tmp = target_abs.with_extension(format!("{}.tmp", unique_suffix()));
        let result: SpreadsheetResult<()> = (|| {
            let file = File::create(&tmp)?;
            let mut zip = zip::ZipWriter::new(file);
            self.write_entries_to(&mut zip, true)?;
            let file = zip.finish()?;
            file.sync_all()?;
            drop(file);
            std::fs::rename(&tmp, &target_abs)?;
            Ok(())
        })();
        match result {
            Ok(()) => {
                // Refresh the in-memory view so the updater stays reusable.
                self.materialize_staged()?;
                Ok(())
            }
            Err(e) => {
                let _ = std::fs::remove_file(&tmp);
                Err(e)
            }
        }
    }

    pub fn ensure_tempdir(&mut self) -> SpreadsheetResult<PathBuf> {
        if self.tempdir.is_none() {
            self.tempdir = Some(tempfile::TempDir::new()?);
        }
        Ok(self.tempdir.as_ref().unwrap().path().to_path_buf())
    }

    pub fn temp_path(&mut self, suffix: &str) -> SpreadsheetResult<PathBuf> {
        let dir = self.ensure_tempdir()?;
        Ok(dir.join(format!("part-{}{}", unique_suffix(), suffix)))
    }

    pub fn remove_temp_file(path: &Path) {
        match std::fs::remove_file(path) {
            Ok(()) => {}
            Err(e) if e.kind() == std::io::ErrorKind::NotFound => {}
            Err(_) => {}
        }
    }

    pub fn dispose(&mut self) {
        for path in self.staged.values() {
            let _ = std::fs::remove_file(path);
        }
        self.staged.clear();
        self.tempdir = None;
    }
}

fn unique_suffix() -> String {
    use std::time::{SystemTime, UNIX_EPOCH};
    let nanos = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .map(|d| d.as_nanos())
        .unwrap_or(0);
    format!("{}-{}", std::process::id(), nanos)
}
