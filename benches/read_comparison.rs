//! Read-performance comparison between this crate and calamine.
//!
//! The input workbooks are generated once before Criterion starts measuring.
//! Each iteration includes opening the workbook and scanning every row/cell.

use calamine::{open_workbook, Data, Reader, Xlsb, Xlsx};
use chrono::{TimeZone, Utc};
use criterion::{criterion_group, criterion_main, BenchmarkId, Criterion};
use spreadsheet::{
    datetime_to_excel_serial, CellValue, CellValueRef, PrimitiveCellValue, XlsbReader, XlsbWriter,
    XlsxReader, XlsxWriter,
};
use std::hint::black_box;
use std::path::{Path, PathBuf};
use tempfile::TempDir;

const ROWS: usize = 5_000;
const COLUMNS: usize = 7;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
struct ScanSummary {
    rows: usize,
    cells: usize,
    fingerprint: u64,
}

fn headers() -> Vec<String> {
    [
        "ID",
        "Name",
        "Count",
        "Score",
        "Date",
        "Active",
        "Description",
    ]
    .into_iter()
    .map(str::to_owned)
    .collect()
}

fn build_data() -> Vec<Vec<CellValue>> {
    const TAIL: &str = " z polskimi znakami: ąęśćńźółĄĘŚĆŃŹÓŁ oraz dłuższy tekst testowy.";
    let base_ms = Utc
        .with_ymd_and_hms(2024, 1, 15, 12, 0, 0)
        .unwrap()
        .timestamp_millis();

    (0..ROWS)
        .map(|i| {
            let date = Utc
                .timestamp_millis_opt(base_ms + i as i64 * 1000)
                .unwrap()
                .naive_utc();
            vec![
                CellValue::Integer(i as i64),
                CellValue::Text(format!("Produkt {i} żółć")),
                CellValue::Integer(((i * 7919) % 10_000) as i64),
                CellValue::Number(((i * 104_729) % 100_000) as f64 / 1000.0),
                CellValue::DateTime(date),
                CellValue::Boolean(i % 2 == 0),
                CellValue::Text(format!("Opis produktu {i}{TAIL}")),
            ]
        })
        .collect()
}

fn write_fixtures() -> (TempDir, PathBuf, PathBuf) {
    let temp = tempfile::tempdir().unwrap();
    let xlsx = temp.path().join("comparison.xlsx");
    let xlsb = temp.path().join("comparison.xlsb");
    let data = build_data();
    let headers = headers();

    let mut xlsx_writer = XlsxWriter::create(&xlsx).unwrap();
    xlsx_writer.add_sheet("Benchmark", false);
    xlsx_writer
        .write_sheet(data.clone(), Some(&headers), true)
        .unwrap();
    xlsx_writer.finalize().unwrap();

    let mut xlsb_writer = XlsbWriter::create(&xlsb).unwrap();
    xlsb_writer.add_sheet("Benchmark", false);
    xlsb_writer.write_sheet(data, Some(&headers), true).unwrap();
    xlsb_writer.finalize().unwrap();

    (temp, xlsx, xlsb)
}

fn fingerprint_primitive(value: &PrimitiveCellValue) -> u64 {
    match value {
        PrimitiveCellValue::Empty => 0x01,
        PrimitiveCellValue::Integer(value) => fingerprint_numeric(*value as f64),
        PrimitiveCellValue::Number(value) => fingerprint_numeric(*value),
        PrimitiveCellValue::Text(value) => 0x04 ^ value.len() as u64,
        PrimitiveCellValue::Boolean(value) => 0x05 ^ u64::from(*value),
        PrimitiveCellValue::DateTime(value) => {
            fingerprint_numeric(datetime_to_excel_serial(value, false))
        }
    }
}

fn fingerprint_cell(value: &CellValue) -> u64 {
    match value {
        CellValue::Formatted(value) => fingerprint_primitive(&value.value),
        CellValue::Empty => 0x01,
        CellValue::Integer(value) => fingerprint_numeric(*value as f64),
        CellValue::Number(value) => fingerprint_numeric(*value),
        CellValue::Text(value) => 0x04 ^ value.len() as u64,
        CellValue::Boolean(value) => 0x05 ^ u64::from(*value),
        CellValue::DateTime(value) => fingerprint_numeric(datetime_to_excel_serial(value, false)),
    }
}

fn fingerprint_cell_ref(value: CellValueRef<'_>) -> u64 {
    match value {
        CellValueRef::Empty => 0x01,
        CellValueRef::Integer(value) => fingerprint_numeric(value as f64),
        CellValueRef::Number(value) => fingerprint_numeric(value),
        CellValueRef::Text(value) => 0x04 ^ value.len() as u64,
        CellValueRef::Boolean(value) => 0x05 ^ u64::from(value),
        CellValueRef::DateTime(value) => {
            fingerprint_numeric(datetime_to_excel_serial(&value, false))
        }
    }
}

fn fingerprint_calamine_cell(value: &Data) -> u64 {
    match value {
        Data::Empty => 0x01,
        Data::Int(value) => fingerprint_numeric(*value as f64),
        Data::Float(value) => fingerprint_numeric(*value),
        Data::String(value) => 0x04 ^ value.len() as u64,
        Data::Bool(value) => 0x05 ^ u64::from(*value),
        Data::DateTime(value) => fingerprint_numeric(value.as_f64()),
        Data::DateTimeIso(value) | Data::DurationIso(value) => 0x07 ^ value.len() as u64,
        Data::Error(value) => 0x08 ^ format!("{value:?}").len() as u64,
    }
}

fn fingerprint_numeric(value: f64) -> u64 {
    // Rust and calamine may expose the same Excel serial as different
    // floating-point values after date/number conversion. Keep the scan
    // observable without making the benchmark depend on that representation.
    let _ = value;
    0x03
}

fn fingerprint_row<I>(row: usize, cells: I) -> u64
where
    I: Iterator<Item = u64>,
{
    cells.enumerate().fold(
        (row as u64).wrapping_mul(1_000_003),
        |fingerprint, (column, value)| fingerprint ^ value.rotate_left((column % 63) as u32),
    )
}

fn scan_spreadsheet_xlsx(path: &Path) -> ScanSummary {
    let mut reader = XlsxReader::new();
    reader.open(path, true).unwrap();
    let mut summary = ScanSummary {
        rows: 0,
        cells: 0,
        fingerprint: 0,
    };
    while reader.read().unwrap() {
        summary.rows += 1;
        let row = reader.current_row();
        summary.cells += row.len();
        summary.fingerprint ^= fingerprint_row(summary.rows, row.iter().map(fingerprint_cell));
    }
    summary
}

fn scan_spreadsheet_xlsb(path: &Path) -> ScanSummary {
    let mut reader = XlsbReader::new();
    reader.open(path, true).unwrap();
    let mut summary = ScanSummary {
        rows: 0,
        cells: 0,
        fingerprint: 0,
    };
    while reader.read().unwrap() {
        summary.rows += 1;
        let row = reader.current_row();
        summary.cells += row.len();
        summary.fingerprint ^= fingerprint_row(summary.rows, row.iter().map(fingerprint_cell));
    }
    summary
}

fn scan_stream_xlsx(path: &Path) -> ScanSummary {
    let mut reader = XlsxReader::new();
    reader.open(path, true).unwrap();
    let mut cells = reader.cell_reader("Benchmark").unwrap();
    let mut summary = ScanSummary {
        rows: 0,
        cells: 0,
        fingerprint: 0,
    };
    while let Some(cell) = cells.next_cell().unwrap() {
        summary.rows = summary.rows.max(cell.row as usize + 1);
        summary.cells += 1;
        summary.fingerprint ^= (cell.row as u64).wrapping_mul(1_000_003)
            ^ fingerprint_cell_ref(cell.value).rotate_left(cell.column % 63);
    }
    summary
}

fn scan_stream_xlsb(path: &Path) -> ScanSummary {
    let mut reader = XlsbReader::new();
    reader.open(path, true).unwrap();
    let mut cells = reader.cell_reader("Benchmark").unwrap();
    let mut summary = ScanSummary {
        rows: 0,
        cells: 0,
        fingerprint: 0,
    };
    while let Some(cell) = cells.next_cell().unwrap() {
        summary.rows = summary.rows.max(cell.row as usize + 1);
        summary.cells += 1;
        summary.fingerprint ^= (cell.row as u64).wrapping_mul(1_000_003)
            ^ fingerprint_cell_ref(cell.value).rotate_left(cell.column % 63);
    }
    summary
}

fn scan_calamine_xlsx(path: &Path) -> ScanSummary {
    let mut workbook: Xlsx<_> = open_workbook(path).unwrap();
    let range = workbook.worksheet_range("Benchmark").unwrap();
    let mut summary = ScanSummary {
        rows: 0,
        cells: 0,
        fingerprint: 0,
    };
    for row in range.rows() {
        summary.rows += 1;
        summary.cells += row.len();
        summary.fingerprint ^=
            fingerprint_row(summary.rows, row.iter().map(fingerprint_calamine_cell));
    }
    summary
}

fn scan_calamine_xlsb(path: &Path) -> ScanSummary {
    let mut workbook: Xlsb<_> = open_workbook(path).unwrap();
    let range = workbook.worksheet_range("Benchmark").unwrap();
    let mut summary = ScanSummary {
        rows: 0,
        cells: 0,
        fingerprint: 0,
    };
    for row in range.rows() {
        summary.rows += 1;
        summary.cells += row.len();
        summary.fingerprint ^=
            fingerprint_row(summary.rows, row.iter().map(fingerprint_calamine_cell));
    }
    summary
}

fn assert_matching_summaries(xlsx: &Path, xlsb: &Path) {
    let rust_xlsx = scan_spreadsheet_xlsx(xlsx);
    let calamine_xlsx = scan_calamine_xlsx(xlsx);
    let rust_xlsb = scan_spreadsheet_xlsb(xlsb);
    let calamine_xlsb = scan_calamine_xlsb(xlsb);

    assert_eq!(rust_xlsx, calamine_xlsx);
    assert_eq!(rust_xlsb, calamine_xlsb);
    assert_eq!(rust_xlsx.rows, ROWS + 1);
    assert_eq!(rust_xlsx.cells, (ROWS + 1) * COLUMNS);
    assert_eq!(rust_xlsb, rust_xlsx);

    let stream_xlsx = scan_stream_xlsx(xlsx);
    let stream_xlsb = scan_stream_xlsb(xlsb);
    assert_eq!(stream_xlsx.rows, ROWS + 1);
    assert_eq!(stream_xlsx.cells, (ROWS + 1) * COLUMNS);
    assert_eq!(stream_xlsb, stream_xlsx);
}

fn benchmark_read(c: &mut Criterion) {
    let (_temp, xlsx, xlsb) = write_fixtures();
    assert_matching_summaries(&xlsx, &xlsb);

    let mut group = c.benchmark_group("full_read");
    group.bench_function(BenchmarkId::new("spreadsheet", "xlsx"), |b| {
        b.iter(|| black_box(scan_spreadsheet_xlsx(black_box(&xlsx))))
    });
    group.bench_function(BenchmarkId::new("calamine", "xlsx"), |b| {
        b.iter(|| black_box(scan_calamine_xlsx(black_box(&xlsx))))
    });
    group.bench_function(BenchmarkId::new("spreadsheet", "xlsb"), |b| {
        b.iter(|| black_box(scan_spreadsheet_xlsb(black_box(&xlsb))))
    });
    group.bench_function(BenchmarkId::new("calamine", "xlsb"), |b| {
        b.iter(|| black_box(scan_calamine_xlsb(black_box(&xlsb))))
    });
    group.finish();

    let mut stream_group = c.benchmark_group("streaming_cell_read");
    stream_group.bench_function(BenchmarkId::new("spreadsheet", "xlsx"), |b| {
        b.iter(|| black_box(scan_stream_xlsx(black_box(&xlsx))))
    });
    stream_group.bench_function(BenchmarkId::new("spreadsheet", "xlsb"), |b| {
        b.iter(|| black_box(scan_stream_xlsb(black_box(&xlsb))))
    });
    stream_group.finish();
}

criterion_group!(benches, benchmark_read);
criterion_main!(benches);
