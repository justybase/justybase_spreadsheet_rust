//! Low-level BIFF12 (XLSB) record helpers (port of `biff12Utils.ts`).
//!
//! A BIFF12 record is: VLQ(id) + VLQ(length) + payload of that length.

use crate::error::{SpreadsheetError, SpreadsheetResult};

/// Decode a 7-bit VLQ integer at `pos`. Returns value and next offset.
pub fn read_vlq(buf: &[u8], mut pos: usize) -> SpreadsheetResult<(u32, usize)> {
    let mut value: u32 = 0;
    for index in 0..5 {
        let byte = *buf
            .get(pos)
            .ok_or_else(|| SpreadsheetError::InvalidFormat("VLQ runs past end of buffer".into()))?;
        pos += 1;
        let payload = byte & 0x7f;
        if index == 4 && payload > 0x0f {
            return Err(SpreadsheetError::InvalidFormat("VLQ overflows u32".into()));
        }
        value |= (payload as u32) << (index * 7);
        if byte & 0x80 == 0 {
            return Ok((value, pos));
        }
    }
    Err(SpreadsheetError::InvalidFormat("VLQ too long".into()))
}

/// Number of bytes needed to VLQ-encode `value`.
pub fn vlq_length(mut value: u32) -> usize {
    let mut n = 1;
    while value >= 0x80 {
        n += 1;
        value >>= 7;
    }
    n
}

/// Encode `value` as 7-bit VLQ bytes.
pub fn vlq_bytes(mut value: u32) -> Vec<u8> {
    let mut out = Vec::new();
    while value >= 0x80 {
        out.push((value as u8 & 0x7f) | 0x80);
        value >>= 7;
    }
    out.push(value as u8 & 0x7f);
    out
}

/// A single BIFF12 record view (byte offsets into the source buffer).
#[derive(Debug, Clone, Copy, Default)]
pub struct Biff12Record {
    /// Byte offset of the record header (id + length).
    pub header_start: usize,
    /// Byte offset just past the header — payload starts here.
    pub data_start: usize,
    /// Byte offset just past the payload.
    pub data_end: usize,
    pub id: u32,
    pub len: u32,
}

/// Fill `rec` with the record starting at `pos`.
/// Returns `Ok(false)` only when the buffer is exhausted. Truncated or
/// overflowing record headers are reported as invalid input.
pub fn read_record(buf: &[u8], pos: usize, rec: &mut Biff12Record) -> SpreadsheetResult<bool> {
    if pos == buf.len() {
        return Ok(false);
    }
    if pos > buf.len() {
        return Err(SpreadsheetError::InvalidFormat(
            "record offset runs past end of buffer".into(),
        ));
    }
    let (id, p1) = read_vlq(buf, pos)?;
    let (len, p2) = read_vlq(buf, p1)?;
    let data_end = p2.checked_add(len as usize).ok_or_else(|| {
        SpreadsheetError::InvalidFormat("record length overflows buffer offset".into())
    })?;
    if data_end > buf.len() {
        return Err(SpreadsheetError::InvalidFormat(
            "record payload runs past end of buffer".into(),
        ));
    }
    rec.header_start = pos;
    rec.data_start = p2;
    rec.data_end = data_end;
    rec.id = id;
    rec.len = len;
    Ok(true)
}

/// Build a BIFF12 record from an id and payload.
pub fn build_record(id: u32, payload: &[u8]) -> Vec<u8> {
    let mut out = vlq_bytes(id);
    out.extend(vlq_bytes(payload.len() as u32));
    out.extend_from_slice(payload);
    out
}

/// Read a UTF-16LE string of `char_count` characters starting at `start`.
pub fn read_utf16(buf: &[u8], start: usize, char_count: usize) -> String {
    try_read_utf16(buf, start, char_count).unwrap_or_default()
}

/// Checked UTF-16LE reader used by format parsers.
pub fn try_read_utf16(buf: &[u8], start: usize, char_count: usize) -> SpreadsheetResult<String> {
    let byte_count = char_count.checked_mul(2).ok_or_else(|| {
        SpreadsheetError::InvalidFormat("UTF-16 character count overflows buffer offset".into())
    })?;
    let end = start.checked_add(byte_count).ok_or_else(|| {
        SpreadsheetError::InvalidFormat("UTF-16 range overflows buffer offset".into())
    })?;
    let bytes = buf.get(start..end).ok_or_else(|| {
        SpreadsheetError::InvalidFormat("UTF-16 string runs past record boundary".into())
    })?;
    let units = bytes
        .as_chunks::<2>()
        .0
        .iter()
        .map(|c| u16::from_le_bytes(*c));
    let mut output = String::with_capacity(char_count);
    for unit in std::char::decode_utf16(units) {
        output.push(unit.unwrap_or(char::REPLACEMENT_CHARACTER));
    }
    Ok(output)
}

/// Read the workbook-level 1904 date-system flag from `xl/workbook.bin`.
pub fn uses_1904_date_system_bin(buf: &[u8]) -> SpreadsheetResult<bool> {
    let mut rec = Biff12Record::default();
    let mut pos = 0;
    while read_record(buf, pos, &mut rec)? {
        if rec.id == 0x0099 && rec.len >= 4 {
            let flags = u32::from_le_bytes(copy4(buf, rec.data_start));
            return Ok(flags & 1 != 0);
        } else if rec.id == 0x0099 {
            return Err(SpreadsheetError::InvalidFormat(
                "workbook date-system record is truncated".into(),
            ));
        }
        pos = rec.data_end;
    }
    Ok(false)
}

/// Parsed contents of `xl/sharedStrings.bin`.
pub struct ParsedSharedStrings {
    pub values: Vec<String>,
    pub total: u32,
    pub unique: u32,
    /// Byte offset of the BrtEndSst record where new items are inserted.
    pub end_sst_offset: usize,
}

/// Parse the contents of `xl/sharedStrings.bin`.
pub fn parse_shared_strings_bin(buf: &[u8]) -> SpreadsheetResult<ParsedSharedStrings> {
    let mut values = Vec::new();
    let mut total = 0;
    let mut unique = 0;
    let mut end_sst_offset = buf.len();
    let mut rec = Biff12Record::default();
    let mut pos = 0;
    while read_record(buf, pos, &mut rec)? {
        pos = rec.data_end;
        match rec.id {
            0x009f => {
                if rec.len < 8 {
                    return Err(SpreadsheetError::InvalidFormat(
                        "shared-string header is truncated".into(),
                    ));
                }
                total = u32::from_le_bytes(copy4(buf, rec.data_start));
                unique = u32::from_le_bytes(copy4(buf, rec.data_start + 4));
            }
            0x0013 => {
                if rec.len < 5 {
                    return Err(SpreadsheetError::InvalidFormat(
                        "shared-string item is truncated".into(),
                    ));
                }
                let cch = u32::from_le_bytes(copy4(buf, rec.data_start + 1)) as usize;
                let byte_count = cch.checked_mul(2).ok_or_else(|| {
                    SpreadsheetError::InvalidFormat("shared-string length overflows".into())
                })?;
                if byte_count > rec.len as usize - 5 {
                    return Err(SpreadsheetError::InvalidFormat(
                        "shared-string UTF-16 payload is truncated".into(),
                    ));
                }
                let text = try_read_utf16(buf, rec.data_start + 5, cch)?;
                values.push(text);
            }
            0x00a0 => {
                end_sst_offset = rec.header_start;
            }
            _ => {}
        }
    }
    Ok(ParsedSharedStrings {
        values,
        total,
        unique,
        end_sst_offset,
    })
}

fn copy4(buf: &[u8], at: usize) -> [u8; 4] {
    [buf[at], buf[at + 1], buf[at + 2], buf[at + 3]]
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn vlq_roundtrip() {
        for v in [0u32, 1, 127, 128, 300, 16383, 1 << 21] {
            let enc = vlq_bytes(v);
            assert_eq!(enc.len(), vlq_length(v));
            let (dec, next) = read_vlq(&enc, 0).unwrap();
            assert_eq!(dec, v);
            assert_eq!(next, enc.len());
        }
    }

    #[test]
    fn record_build_and_read() {
        let payload = vec![1u8, 2, 3];
        let rec_bytes = build_record(0x13, &payload);
        let mut rec = Biff12Record::default();
        assert!(read_record(&rec_bytes, 0, &mut rec).unwrap());
        assert_eq!(rec.id, 0x13);
        assert_eq!(rec.len, 3);
        assert_eq!(&rec_bytes[rec.data_start..rec.data_end], &payload[..]);
    }

    #[test]
    fn reads_1904_workbook_flag() {
        let mut payload = vec![0u8; 4];
        payload[0] = 1;
        assert!(uses_1904_date_system_bin(&build_record(0x0099, &payload)).unwrap());
        assert!(!uses_1904_date_system_bin(&build_record(0x0099, &[0, 0, 0, 0])).unwrap());
    }

    #[test]
    fn truncated_record_is_not_treated_as_eof() {
        let mut rec = Biff12Record::default();
        let error = read_record(&[0x13, 0x04, 0x01], 0, &mut rec).unwrap_err();
        assert!(matches!(error, SpreadsheetError::InvalidFormat(_)));
    }

    #[test]
    fn overflowing_vlq_is_rejected() {
        let error = read_vlq(&[0xff, 0xff, 0xff, 0xff, 0x10], 0).unwrap_err();
        assert!(matches!(error, SpreadsheetError::InvalidFormat(_)));
    }
}
