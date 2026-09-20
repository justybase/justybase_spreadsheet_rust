//! Chunked byte accumulator (port of `BigBuffer.ts`).
//!
//! Avoids reallocating giant sheets: bytes are staged in fixed-size
//! chunks and handed to the ZIP writer as a list of slices.

#[derive(Debug)]
pub struct BigBuffer {
    chunk_size: usize,
    chunks: Vec<Vec<u8>>,
    current: Vec<u8>,
    cursor: usize,
}

impl BigBuffer {
    pub fn new(chunk_size: usize) -> Self {
        let chunk_size = chunk_size.max(8);
        Self {
            chunk_size,
            chunks: Vec::new(),
            current: vec![0u8; chunk_size],
            cursor: 0,
        }
    }

    fn flush(&mut self) {
        if self.cursor > 0 {
            // Hand the filled working chunk to `chunks` as-is instead of
            // copying it into a right-sized allocation; the fresh working
            // chunk below is the only new allocation per flush.
            self.current.truncate(self.cursor);
            let filled = std::mem::take(&mut self.current);
            self.chunks.push(filled);
            self.current = vec![0u8; self.chunk_size];
            self.cursor = 0;
        }
    }

    pub fn ensure_capacity(&mut self, size: usize) {
        if size > self.chunk_size {
            self.flush();
            self.chunk_size = size;
            self.current = vec![0u8; size];
            return;
        }
        if self.cursor + size > self.chunk_size {
            self.flush();
        }
    }

    pub fn write(&mut self, buf: &[u8]) {
        let mut len = buf.len();
        let mut offset = 0;
        while len > 0 {
            if self.chunk_size - self.cursor == 0 {
                self.flush();
            }
            let available = self.chunk_size - self.cursor;
            let to_write = len.min(available);
            self.current[self.cursor..self.cursor + to_write]
                .copy_from_slice(&buf[offset..offset + to_write]);
            self.cursor += to_write;
            offset += to_write;
            len -= to_write;
        }
    }

    pub fn write_byte(&mut self, val: u8) {
        self.ensure_capacity(1);
        self.current[self.cursor] = val;
        self.cursor += 1;
    }

    /// Caller must have called `ensure_capacity` first.
    pub fn write_unsafe_byte(&mut self, val: u8) {
        self.current[self.cursor] = val;
        self.cursor += 1;
    }

    pub fn write_i32_le(&mut self, val: i32) {
        self.ensure_capacity(4);
        self.current[self.cursor..self.cursor + 4].copy_from_slice(&val.to_le_bytes());
        self.cursor += 4;
    }

    /// Caller must have called `ensure_capacity` first.
    pub fn write_unsafe_i32_le(&mut self, val: i32) {
        self.current[self.cursor..self.cursor + 4].copy_from_slice(&val.to_le_bytes());
        self.cursor += 4;
    }

    pub fn write_u32_le(&mut self, val: u32) {
        self.write_i32_le(val as i32);
    }

    pub fn write_f64_le(&mut self, val: f64) {
        self.ensure_capacity(8);
        self.current[self.cursor..self.cursor + 8].copy_from_slice(&val.to_le_bytes());
        self.cursor += 8;
    }

    /// Caller must have called `ensure_capacity` first.
    pub fn write_unsafe_f64_le(&mut self, val: f64) {
        self.current[self.cursor..self.cursor + 8].copy_from_slice(&val.to_le_bytes());
        self.cursor += 8;
    }

    pub fn write_str(&mut self, s: &str) {
        let bytes = s.as_bytes();
        if self.cursor + bytes.len() <= self.chunk_size {
            self.current[self.cursor..self.cursor + bytes.len()].copy_from_slice(bytes);
            self.cursor += bytes.len();
            return;
        }
        self.flush();
        if bytes.len() > self.chunk_size {
            self.chunks.push(bytes.to_vec());
        } else {
            self.current[..bytes.len()].copy_from_slice(bytes);
            self.cursor = bytes.len();
        }
    }

    pub fn write_utf16le(&mut self, s: &str) {
        // Encode directly into the existing chunks.  This avoids both the
        // temporary Vec<u16> and a preliminary UTF-16 length pass for every
        // shared-string cell.
        if self.chunk_size < 2 {
            let mut buf = Vec::with_capacity(s.len().saturating_mul(2));
            for unit in s.encode_utf16() {
                buf.extend_from_slice(&unit.to_le_bytes());
            }
            self.flush();
            self.chunks.push(buf);
            return;
        }
        for unit in s.encode_utf16() {
            if self.cursor + 2 > self.chunk_size {
                self.flush();
            }
            self.current[self.cursor..self.cursor + 2].copy_from_slice(&unit.to_le_bytes());
            self.cursor += 2;
        }
    }

    pub fn chunks_drain(&mut self) -> Vec<Vec<u8>> {
        self.flush();
        std::mem::take(&mut self.chunks)
    }

    pub fn reset(&mut self) {
        self.chunks.clear();
        self.cursor = 0;
    }

    pub fn chunk_size(&self) -> usize {
        self.chunk_size
    }
}

impl std::fmt::Write for BigBuffer {
    fn write_str(&mut self, s: &str) -> std::fmt::Result {
        self.write_str(s);
        Ok(())
    }
}

impl Default for BigBuffer {
    fn default() -> Self {
        Self::new(65536)
    }
}

/// Output payload for one ZIP entry.
///
/// Sheets staged by [`BigBuffer`] are kept as their chunk list and streamed
/// to the ZIP writer directly, avoiding a full contiguous copy (`concat`).
pub enum EntryData {
    /// Flat buffer (single allocation, e.g. an XML `String`).
    Bytes(Vec<u8>),
    /// Chunk list from [`BigBuffer::chunks_drain`].
    Chunks(Vec<Vec<u8>>),
}

impl EntryData {
    /// Write the payload in order without concatenating the chunks.
    pub fn write_to<W: std::io::Write>(&self, writer: &mut W) -> std::io::Result<()> {
        match self {
            EntryData::Bytes(data) => writer.write_all(data),
            EntryData::Chunks(chunks) => {
                for chunk in chunks {
                    writer.write_all(chunk)?;
                }
                Ok(())
            }
        }
    }
}

impl From<Vec<u8>> for EntryData {
    fn from(data: Vec<u8>) -> Self {
        EntryData::Bytes(data)
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn chunking_and_utf16() {
        let mut b = BigBuffer::new(8);
        b.write(b"hello world, this is long");
        b.write_str("!");
        b.write_utf16le("AB");
        b.write_i32_le(-5);
        b.write_f64_le(1.5);
        let chunks = b.chunks_drain();
        let flat: Vec<u8> = chunks.concat();
        assert!(flat.windows(5).any(|w| w == b"hello"));
        // "AB" as UTF-16LE
        assert!(flat.windows(4).any(|w| w == [0x41, 0x00, 0x42, 0x00]));
        let mut utf16 = BigBuffer::new(4);
        utf16.write_utf16le("A😀B");
        let expected: Vec<u8> = "A😀B".encode_utf16().flat_map(u16::to_le_bytes).collect();
        assert_eq!(utf16.chunks_drain().concat(), expected);
        b.reset();
        assert!(b.chunks_drain().is_empty());
    }

    #[test]
    fn tiny_chunks_do_not_panic_on_fixed_width_writes() {
        let mut b = BigBuffer::new(0);
        b.write_i32_le(42);
        b.write_f64_le(1.5);
        assert_eq!(b.chunks_drain().concat().len(), 12);
    }
}
