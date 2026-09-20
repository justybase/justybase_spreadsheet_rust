//! Atomic file installation (port of `atomicFile.ts`).

use crate::error::{SpreadsheetError, SpreadsheetResult};
use std::fs;
use std::io::Write;
use std::path::{Path, PathBuf};

/// File-operations seam used to test replacement races without Excel.
pub trait AtomicFileOperations {
    fn exists(&self, path: &Path) -> bool;
    fn rename(&self, from: &Path, to: &Path) -> std::io::Result<()>;
    fn remove(&self, path: &Path) -> std::io::Result<()>;
}

pub struct RealFileOperations;

impl AtomicFileOperations for RealFileOperations {
    fn exists(&self, path: &Path) -> bool {
        path.exists()
    }
    fn rename(&self, from: &Path, to: &Path) -> std::io::Result<()> {
        fs::rename(from, to)
    }
    fn remove(&self, path: &Path) -> std::io::Result<()> {
        fs::remove_file(path)
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

/// Install a temporary result without deleting an unapproved or
/// concurrently-created destination.
pub fn install_temporary_output(
    temporary_path: &Path,
    destination_path: &Path,
    overwrite: bool,
    ops: &dyn AtomicFileOperations,
) -> SpreadsheetResult<()> {
    let backup_path =
        destination_path.with_extension(format!("{}.replacement-backup", unique_suffix()));
    let mut backup_created = false;
    let mut installed = false;

    if ops.exists(destination_path) {
        if !overwrite {
            return Err(SpreadsheetError::DestinationExists(
                destination_path.display().to_string(),
            ));
        }
        ops.rename(destination_path, &backup_path)?;
        backup_created = true;
    }

    let install_result: SpreadsheetResult<()> =
        (|| match ops.rename(temporary_path, destination_path) {
            Ok(()) => {
                installed = true;
                Ok(())
            }
            Err(e) => {
                if overwrite && !backup_created && ops.exists(destination_path) {
                    ops.rename(destination_path, &backup_path)?;
                    backup_created = true;
                    ops.rename(temporary_path, destination_path)?;
                    installed = true;
                    Ok(())
                } else {
                    Err(SpreadsheetError::Io(e))
                }
            }
        })();

    match install_result {
        Ok(()) => {
            if backup_created {
                let _ = ops.remove(&backup_path);
            }
            Ok(())
        }
        Err(e) => {
            if backup_created && !installed {
                if !ops.exists(destination_path) {
                    ops.rename(&backup_path, destination_path).map_err(|rb| {
                        SpreadsheetError::Other(format!(
                            "final destination install failed ({e}), and restoring the original destination also failed ({rb}). Original file remains at {}.",
                            backup_path.display()
                        ))
                    })?;
                } else {
                    return Err(SpreadsheetError::Other(format!(
                        "final destination install failed ({e}). A concurrent destination was preserved; the original file remains at {}.",
                        backup_path.display()
                    )));
                }
            }
            Err(e)
        }
    }
}

#[cfg(unix)]
fn destination_mode(target: &Path) -> Option<u32> {
    use std::os::unix::fs::PermissionsExt;
    fs::metadata(target)
        .ok()
        .map(|m| m.permissions().mode() & 0o777)
}

#[cfg(not(unix))]
fn destination_mode(_target: &Path) -> Option<u32> {
    None
}

/// Write a complete buffer through a same-directory temporary file and
/// replace the destination only after the write completes successfully.
pub fn write_buffer_atomically(buf: &[u8], destination_path: &Path) -> SpreadsheetResult<()> {
    let target = if destination_path.is_absolute() {
        destination_path.to_path_buf()
    } else {
        std::env::current_dir()
            .map(|c| c.join(destination_path))
            .unwrap_or_else(|_| destination_path.to_path_buf())
    };
    if let Some(parent) = target.parent() {
        if !parent.as_os_str().is_empty() {
            fs::create_dir_all(parent)?;
        }
    }
    let file_name = target
        .file_name()
        .map(|s| s.to_string_lossy().into_owned())
        .unwrap_or_else(|| "output".into());
    let parent: PathBuf = target
        .parent()
        .map(|p| p.to_path_buf())
        .unwrap_or_else(|| PathBuf::from("."));
    let tmp_path = parent.join(format!(".{}.{}.tmp", file_name, unique_suffix()));
    let mode = destination_mode(&target);

    let write_result: SpreadsheetResult<()> = (|| {
        let mut f = fs::File::create(&tmp_path)?;
        #[cfg(unix)]
        {
            use std::os::unix::fs::PermissionsExt;
            if let Some(m) = mode {
                let _ = f.set_permissions(fs::Permissions::from_mode(m));
            }
        }
        f.write_all(buf)?;
        f.sync_all()?;
        drop(f);
        #[cfg(unix)]
        {
            use std::os::unix::fs::PermissionsExt;
            if let Some(m) = mode {
                let _ = fs::set_permissions(&tmp_path, fs::Permissions::from_mode(m));
            }
        }
        // fsync the file once more via reopen for durability parity with TS
        let f2 = fs::OpenOptions::new().read(true).open(&tmp_path)?;
        f2.sync_all()?;
        Ok(())
    })();

    match write_result {
        Ok(()) => {
            let r = fs::rename(&tmp_path, &target);
            if r.is_err() && tmp_path.exists() {
                let _ = fs::remove_file(&tmp_path);
            }
            r?;
            Ok(())
        }
        Err(e) => {
            if tmp_path.exists() {
                let _ = fs::remove_file(&tmp_path);
            }
            Err(e)
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::collections::HashMap;

    struct MemOps {
        files: std::cell::RefCell<HashMap<PathBuf, Vec<u8>>>,
        fail_first_rename: std::cell::Cell<bool>,
    }

    impl AtomicFileOperations for MemOps {
        fn exists(&self, path: &Path) -> bool {
            self.files.borrow().contains_key(path)
        }
        fn rename(&self, from: &Path, to: &Path) -> std::io::Result<()> {
            if self.fail_first_rename.get() {
                self.fail_first_rename.set(false);
                return Err(std::io::Error::other("boom"));
            }
            let mut files = self.files.borrow_mut();
            let data = files
                .remove(from)
                .ok_or_else(|| std::io::Error::new(std::io::ErrorKind::NotFound, "missing tmp"))?;
            files.insert(to.to_path_buf(), data);
            Ok(())
        }
        fn remove(&self, path: &Path) -> std::io::Result<()> {
            self.files.borrow_mut().remove(path);
            Ok(())
        }
    }

    #[test]
    fn refuses_without_overwrite() {
        let ops = MemOps {
            files: std::cell::RefCell::new(HashMap::from([(PathBuf::from("dst"), vec![1])])),
            fail_first_rename: std::cell::Cell::new(false),
        };
        let r = install_temporary_output(Path::new("tmp"), Path::new("dst"), false, &ops);
        assert!(matches!(r, Err(SpreadsheetError::DestinationExists(_))));
    }

    #[test]
    fn atomic_write_roundtrip() {
        let dir = tempfile::tempdir().unwrap();
        let dst = dir.path().join("out.bin");
        write_buffer_atomically(b"hello", &dst).unwrap();
        assert_eq!(fs::read(&dst).unwrap(), b"hello");
        write_buffer_atomically(b"world", &dst).unwrap();
        assert_eq!(fs::read(&dst).unwrap(), b"world");
    }
}
