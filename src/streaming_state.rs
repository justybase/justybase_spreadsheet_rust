//! Streaming-sheet session state (port of `StreamingSheetState.ts`).

use crate::big_buffer::BigBuffer;
use crate::error::{SpreadsheetError, SpreadsheetResult};
use crate::writer_helpers::init_col_widths;

#[derive(Debug, Default)]
pub struct StreamingSheetState {
    pub is_streaming: bool,
    pub row_num: u32,
    pub start_col: usize,
    pub end_col: usize,
    pub do_autofilter: bool,
    pub col_widths: Vec<f64>,
    pub buffer: Option<BigBuffer>,
}

impl StreamingSheetState {
    pub fn new() -> Self {
        Self::default()
    }

    /// Begin a streaming sheet session.
    pub fn begin(&mut self, column_count: usize, do_autofilter: bool, buffer: BigBuffer) {
        if self.is_streaming {
            // Mirror TS which throws; in Rust signal via panic-free error is
            // handled by callers checking first — keep TS message.
            panic!("Already in streaming mode. Call end_sheet() first.");
        }
        self.is_streaming = true;
        self.buffer = Some(buffer);
        self.row_num = 0;
        self.start_col = 0;
        self.end_col = column_count;
        self.do_autofilter = do_autofilter;
        self.col_widths = init_col_widths(column_count);
    }

    pub fn try_begin(
        &mut self,
        column_count: usize,
        do_autofilter: bool,
        buffer: BigBuffer,
    ) -> SpreadsheetResult<()> {
        if self.is_streaming {
            return Err(SpreadsheetError::AlreadyStreaming);
        }
        self.is_streaming = true;
        self.buffer = Some(buffer);
        self.row_num = 0;
        self.start_col = 0;
        self.end_col = column_count;
        self.do_autofilter = do_autofilter;
        self.col_widths = init_col_widths(column_count);
        Ok(())
    }

    /// Assert streaming mode and return the active buffer.
    pub fn assert_streaming(&mut self) -> SpreadsheetResult<&mut BigBuffer> {
        if !self.is_streaming || self.buffer.is_none() {
            return Err(SpreadsheetError::NotStreaming);
        }
        Ok(self.buffer.as_mut().unwrap())
    }

    /// Reset after `end_sheet()`.
    pub fn end(&mut self) {
        self.is_streaming = false;
        self.buffer = None;
        self.row_num = 0;
        self.start_col = 0;
        self.end_col = 0;
        self.do_autofilter = false;
        self.col_widths.clear();
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn begin_end_cycle() {
        let mut s = StreamingSheetState::new();
        assert!(s.assert_streaming().is_err());
        s.try_begin(3, true, BigBuffer::default()).unwrap();
        assert!(s.try_begin(3, true, BigBuffer::default()).is_err());
        assert_eq!(s.col_widths.len(), 3);
        s.end();
        assert!(!s.is_streaming);
    }
}
