//! Stateful forward-only BIFF12 parser (port of `BiffReaderWriter.ts`).

use std::collections::HashSet;

use crate::error::{SpreadsheetError, SpreadsheetResult};
use crate::formats::is_date_format_code;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum BiffCellType {
    Blank = 0,
    SharedString = 2,
    Number = 3,
    Bool = 4,
    Str = 5,
}

pub struct BiffReaderWriter<'a> {
    buffer: &'a [u8],
    pos: usize,
    length: usize,

    pub is_sheet: bool,
    pub workbook_id: u32,
    pub rec_id: Option<String>,
    pub workbook_name: Option<String>,

    in_cell_xf: bool,
    in_number_format: bool,

    pub shared_string_value: Option<String>,
    pub shared_string_unique_count: u32,

    pub cell_type: u8,
    pub int_value: u32,
    pub double_val: f64,
    pub bool_value: bool,
    pub string_value: Option<String>,
    pub column_num: i64,
    pub xf_index: usize,
    pub read_cell: bool,
    pub row_index: i32,

    record_start: usize,
    record_end: usize,

    pub xf_index_to_num_fmt_id: Vec<u16>,
    pub custom_num_fmts: HashSet<u16>,
}

impl<'a> BiffReaderWriter<'a> {
    pub fn new(buffer: &'a [u8]) -> Self {
        Self {
            buffer,
            pos: 0,
            length: buffer.len(),
            is_sheet: false,
            workbook_id: 0,
            rec_id: None,
            workbook_name: None,
            in_cell_xf: false,
            in_number_format: false,
            shared_string_value: None,
            shared_string_unique_count: 0,
            cell_type: 0,
            int_value: 0,
            double_val: 0.0,
            bool_value: false,
            string_value: None,
            column_num: -1,
            xf_index: 0,
            read_cell: false,
            row_index: -1,
            record_start: 0,
            record_end: 0,
            xf_index_to_num_fmt_id: Vec::new(),
            custom_num_fmts: HashSet::new(),
        }
    }

    #[inline(always)]
    fn try_read_variable_value(&mut self) -> SpreadsheetResult<Option<u32>> {
        if self.pos >= self.length {
            return Ok(None);
        }
        let mut value = 0u32;
        for index in 0..5 {
            let byte = *self.buffer.get(self.pos).ok_or_else(|| {
                SpreadsheetError::InvalidFormat("BIFF12 VLQ runs past end of buffer".into())
            })?;
            self.pos += 1;
            let payload = byte & 0x7f;
            if index == 4 && payload > 0x0f {
                return Err(SpreadsheetError::InvalidFormat(
                    "BIFF12 VLQ overflows u32".into(),
                ));
            }
            if index == 4 && byte & 0x80 != 0 {
                return Err(SpreadsheetError::InvalidFormat(
                    "BIFF12 VLQ is too long".into(),
                ));
            }
            value |= (payload as u32) << (index * 7);
            if byte & 0x80 == 0 {
                return Ok(Some(value));
            }
        }
        Err(SpreadsheetError::InvalidFormat(
            "BIFF12 VLQ is too long".into(),
        ))
    }

    #[inline(always)]
    fn field_range(
        &self,
        offset: usize,
        width: usize,
    ) -> SpreadsheetResult<std::ops::Range<usize>> {
        let start = self.record_start.checked_add(offset).ok_or_else(|| {
            SpreadsheetError::InvalidFormat("BIFF12 field offset overflows".into())
        })?;
        let end = start.checked_add(width).ok_or_else(|| {
            SpreadsheetError::InvalidFormat("BIFF12 field width overflows".into())
        })?;
        if end > self.record_end {
            return Err(SpreadsheetError::InvalidFormat(
                "BIFF12 record payload is truncated".into(),
            ));
        }
        Ok(start..end)
    }

    #[inline(always)]
    fn get_dword(&self, offset: usize) -> SpreadsheetResult<u32> {
        let range = self.field_range(offset, 4)?;
        Ok(u32::from_le_bytes(copy4(self.buffer, range.start)))
    }

    #[inline(always)]
    fn get_i32(&self, offset: usize) -> SpreadsheetResult<i32> {
        let range = self.field_range(offset, 4)?;
        Ok(i32::from_le_bytes(copy4(self.buffer, range.start)))
    }

    #[inline(always)]
    fn get_word(&self, offset: usize) -> SpreadsheetResult<u16> {
        let range = self.field_range(offset, 2)?;
        Ok(u16::from_le_bytes([
            self.buffer[range.start],
            self.buffer[range.start + 1],
        ]))
    }

    #[inline(always)]
    fn get_byte(&self, offset: usize) -> SpreadsheetResult<u8> {
        let range = self.field_range(offset, 1)?;
        Ok(self.buffer[range.start])
    }

    #[inline(always)]
    fn get_double(&self, offset: usize) -> SpreadsheetResult<f64> {
        let range = self.field_range(offset, 8)?;
        let mut b = [0u8; 8];
        b.copy_from_slice(&self.buffer[range]);
        Ok(f64::from_le_bytes(b))
    }

    #[inline(always)]
    fn get_string(&self, offset: usize, length: usize) -> SpreadsheetResult<String> {
        let width = length.checked_mul(2).ok_or_else(|| {
            SpreadsheetError::InvalidFormat("BIFF12 string length overflows".into())
        })?;
        let range = self.field_range(offset, width)?;
        let units = self.buffer[range]
            .as_chunks::<2>()
            .0
            .iter()
            .map(|chunk| u16::from_le_bytes(*chunk));
        let mut output = String::with_capacity(length);
        for unit in std::char::decode_utf16(units) {
            output.push(unit.unwrap_or(char::REPLACEMENT_CHARACTER));
        }
        Ok(output)
    }

    fn get_nullable_string(&self, offset: &mut usize) -> SpreadsheetResult<Option<String>> {
        let length = self.get_dword(*offset)?;
        *offset = (*offset).checked_add(4).ok_or_else(|| {
            SpreadsheetError::InvalidFormat("BIFF12 field offset overflows".into())
        })?;
        if length == 0xFFFF_FFFF {
            return Ok(None);
        }
        let length = length as usize;
        let s = self.get_string(*offset, length)?;
        *offset = (*offset)
            .checked_add(length.checked_mul(2).ok_or_else(|| {
                SpreadsheetError::InvalidFormat("BIFF12 string length overflows".into())
            })?)
            .ok_or_else(|| {
                SpreadsheetError::InvalidFormat("BIFF12 field offset overflows".into())
            })?;
        Ok(Some(s))
    }

    #[inline(always)]
    fn begin_record(&mut self) -> SpreadsheetResult<Option<u32>> {
        let record_id = match self.try_read_variable_value()? {
            Some(value) => value,
            None => return Ok(None),
        };
        let record_length = self.try_read_variable_value()?.ok_or_else(|| {
            SpreadsheetError::InvalidFormat("BIFF12 record length is truncated".into())
        })?;
        let start_pos = self.pos;
        let record_end = start_pos
            .checked_add(record_length as usize)
            .ok_or_else(|| {
                SpreadsheetError::InvalidFormat("BIFF12 record length overflows".into())
            })?;
        if record_end > self.length {
            return Err(SpreadsheetError::InvalidFormat(
                "BIFF12 record payload is truncated".into(),
            ));
        }
        self.record_start = start_pos;
        self.record_end = record_end;
        self.pos = record_end;
        Ok(Some(record_id))
    }

    pub fn read_workbook(&mut self) -> SpreadsheetResult<bool> {
        let record_id = match self.begin_record()? {
            Some(id) => id,
            None => return Ok(false),
        };
        self.is_sheet = false;
        if record_id == 0x9C {
            self.workbook_id = self.get_dword(4)?;
            let mut off = 8;
            self.rec_id = self.get_nullable_string(&mut off)?;
            let name_length = self.get_dword(off)? as usize;
            self.workbook_name = Some(self.get_string(off + 4, name_length)?);
            self.is_sheet = true;
        }
        Ok(true)
    }

    pub fn read_shared_strings(&mut self) -> SpreadsheetResult<bool> {
        let record_id = match self.begin_record()? {
            Some(id) => id,
            None => return Ok(false),
        };
        self.shared_string_value = None;
        if record_id == 0x13 {
            let length = self.get_dword(1)? as usize;
            self.shared_string_value = Some(self.get_string(5, length)?);
        } else if record_id == 159 {
            self.shared_string_unique_count = self.get_dword(4)?;
        }
        Ok(true)
    }

    pub fn read_styles(&mut self) -> SpreadsheetResult<bool> {
        let record_id = match self.begin_record()? {
            Some(id) => id,
            None => return Ok(false),
        };
        match record_id {
            0x269 => self.in_cell_xf = true,
            0x26a => self.in_cell_xf = false,
            0x267 => self.in_number_format = true,
            0x268 => self.in_number_format = false,
            0x2f if self.in_cell_xf => {
                let num_fmt = self.get_word(2)?;
                self.xf_index_to_num_fmt_id.push(num_fmt);
            }
            0x2c if self.in_number_format => {
                let fmt = self.get_word(0)?;
                let length = self.get_dword(2)? as usize;
                let fmt_string = self.get_string(6, length)?;
                if is_date_format_code(&fmt_string) {
                    self.custom_num_fmts.insert(fmt);
                }
            }
            _ => {}
        }
        Ok(true)
    }

    fn get_rk_number(&self, offset: usize) -> SpreadsheetResult<f64> {
        let flags = self.get_byte(offset)?;
        let mut result = if flags & 0x02 != 0 {
            (self.get_i32(offset)? >> 2) as f64
        } else {
            let raw = self.get_i32(offset)?;
            let high_bits = (raw as u32) & 0xFFFF_FFFC;
            let mut dbuf = [0u8; 8];
            dbuf[4..8].copy_from_slice(&high_bits.to_le_bytes());
            f64::from_le_bytes(dbuf)
        };
        if flags & 0x01 != 0 {
            result /= 100.0;
        }
        Ok(result)
    }

    #[inline(always)]
    pub fn read_worksheet(&mut self) -> SpreadsheetResult<bool> {
        let record_id = match self.begin_record()? {
            Some(id) => id,
            None => return Ok(false),
        };
        self.read_cell = false;
        self.column_num = -1;
        self.string_value = None;
        match record_id {
            0x00 => {
                self.row_index = self.get_i32(0)?;
            }
            0x01 | 0x03 | 0x0b => {
                self.read_cell = true;
                self.cell_type = 0;
            }
            0x02 => {
                self.double_val = self.get_rk_number(8)?;
                self.read_cell = true;
                self.cell_type = 3;
            }
            0x04 | 0x0a => {
                self.bool_value = self.get_byte(8)? == 1;
                self.read_cell = true;
                self.cell_type = 4;
            }
            0x09 | 0x05 => {
                self.double_val = self.get_double(8)?;
                self.read_cell = true;
                self.cell_type = 3;
            }
            0x06 | 0x08 => {
                let length = self.get_dword(8)? as usize;
                self.string_value = Some(self.get_string(12, length)?);
                self.read_cell = true;
                self.cell_type = 5;
            }
            0x07 => {
                self.int_value = self.get_dword(8)?;
                self.read_cell = true;
                self.cell_type = 2;
            }
            _ => {}
        }
        if self.read_cell {
            self.column_num = self.get_dword(0)? as i64;
            self.xf_index = (self.get_dword(4)? & 0xffffff) as usize;
        }
        Ok(true)
    }
}

fn copy4(buf: &[u8], at: usize) -> [u8; 4] {
    [buf[at], buf[at + 1], buf[at + 2], buf[at + 3]]
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::biff12::build_record;

    #[test]
    fn rk_integer_decodes() {
        // BrtCellRk for 42: rk = (42 << 2) | 2 = 170
        let mut payload = vec![0u8; 12];
        payload[0..4].copy_from_slice(&0u32.to_le_bytes()); // col
        payload[4..8].copy_from_slice(&0u32.to_le_bytes()); // xf
        payload[8..12].copy_from_slice(&170i32.to_le_bytes());
        let buf = build_record(0x02, &payload);
        let mut r = BiffReaderWriter::new(&buf);
        assert!(r.read_worksheet().unwrap());
        assert!(r.read_cell);
        assert_eq!(r.cell_type, 3);
        assert_eq!(r.double_val, 42.0);
    }

    #[test]
    fn row_and_sst_item() {
        let mut row_payload = vec![0u8; 25];
        row_payload[0..4].copy_from_slice(&7i32.to_le_bytes());
        let mut buf = build_record(0x00, &row_payload);
        let mut sst_payload = vec![0u8; 12];
        sst_payload[0..4].copy_from_slice(&3u32.to_le_bytes());
        sst_payload[4..8].copy_from_slice(&0u32.to_le_bytes());
        sst_payload[8..12].copy_from_slice(&9u32.to_le_bytes());
        buf.extend(build_record(0x07, &sst_payload));
        let mut r = BiffReaderWriter::new(&buf);
        assert!(r.read_worksheet().unwrap());
        assert_eq!(r.row_index, 7);
        assert!(r.read_worksheet().unwrap());
        assert_eq!(r.cell_type, 2);
        assert_eq!(r.int_value, 9);
        assert_eq!(r.column_num, 3);
    }

    #[test]
    fn truncated_cell_record_returns_error() {
        let record = build_record(0x05, &[0; 4]);
        let mut r = BiffReaderWriter::new(&record);
        let error = r.read_worksheet().unwrap_err();
        assert!(matches!(error, SpreadsheetError::InvalidFormat(_)));
    }
}
