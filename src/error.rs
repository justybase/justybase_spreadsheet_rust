use thiserror::Error;

#[derive(Debug, Error)]
pub enum SpreadsheetError {
    #[error("I/O error: {0}")]
    Io(#[from] std::io::Error),
    #[error("ZIP error: {0}")]
    Zip(#[from] zip::result::ZipError),
    #[error("sheet not found: {0}")]
    SheetNotFound(String),
    #[error("invalid format: {0}")]
    InvalidFormat(String),
    #[error("unsupported extension: {0}")]
    UnsupportedExtension(String),
    #[error("row length mismatch: expected {expected} columns, got {got}")]
    RowLengthMismatch { expected: usize, got: usize },
    #[error("not in streaming mode; call start_sheet() first")]
    NotStreaming,
    #[error("already in streaming mode; call end_sheet() first")]
    AlreadyStreaming,
    #[error("a streaming worksheet is still open; call end_sheet() first")]
    StreamingSheetOpen,
    #[error("destination already exists (pass overwrite=true to replace it): {0}")]
    DestinationExists(String),
    #[error("operation not supported on this platform: {0}")]
    UnsupportedPlatform(String),
    #[error("{0}")]
    Other(String),
}

pub type SpreadsheetResult<T> = Result<T, SpreadsheetError>;
