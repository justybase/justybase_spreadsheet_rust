//! XLSB write comparison: `spreadsheet` vs every published Rust XLSB writer.
//!
//! Cases:
//!  1. `spreadsheet batch` — `XlsbWriter`: `add_sheet` + `write_sheet` + `finalize`;
//!  2. `spreadsheet streaming` — `start_sheet` + `write_row` + `end_sheet` + `finalize`;
//!  3. `rxlsb batch` — `write_batch` with a cell supplier closure + `close`;
//!  4. `rxlsb streaming` — `start_sheet` + `write_rows` + `end_sheet` + `close`;
//!  5. `xlsb-writer (arrow)` — `write_sheet` from one Arrow `RecordBatch` + `finish`.
//!
//! Excluded: `SaelKimberly/rust-xlsb-writer` (GitHub-only pre-alpha, no
//! crates.io release / versioned API).
//!
//! Methodology (same as `write_comparison`): every iteration rebuilds the
//! input outside the measurement window; the window covers writer creation up
//! to `finalize()`/`close()`/`finish()` including the file write; 1 warmup +
//! N measured runs (median time, max peak heap, allocation total from a
//! counting `GlobalAllocator`); outputs re-read with `XlsbReader` and
//! row/cell counts asserted.
//!
//! Compression: this crate defaults to Deflate level 1 (configurable via the
//! optional `compression` argument — re-run with `6` for an equal-footing
//! cross-check); `xlsb-writer` hardcodes level 6; `rxlsb` uses the zip crate
//! default (~6).
//!
//! Usage:
//!   cargo run --release --example xlsb_write_comparison [outDir] [rows] [iters] [columns] [compression]

use std::alloc::{GlobalAlloc, Layout, System};
use std::fs::File;
use std::path::{Path, PathBuf};
use std::sync::atomic::{AtomicUsize, Ordering};
use std::sync::Arc;
use std::time::Instant;

use arrow_array::{BooleanArray, Date64Array, Float64Array, Int64Array, RecordBatch, StringArray};
use arrow_schema::{DataType, Field, Schema};
use chrono::{TimeZone, Utc};
use spreadsheet::{CellValue, XlsbReader, XlsbSheetOptions, XlsbWriter as OurXlsbWriter};

const SHEET_NAME: &str = "Benchmark";
const DEFAULT_ROWS: usize = 5_000;
const DEFAULT_ITERS: usize = 3;
const DEFAULT_COLUMNS: usize = 7;

// ---------------------------------------------------------------------------
// Allocation tracking: peak live heap and total allocated bytes.
// ---------------------------------------------------------------------------

struct TrackingAllocator;

static LIVE: AtomicUsize = AtomicUsize::new(0);
static PEAK: AtomicUsize = AtomicUsize::new(0);
static TOTAL: AtomicUsize = AtomicUsize::new(0);

#[global_allocator]
static GLOBAL: TrackingAllocator = TrackingAllocator;

impl TrackingAllocator {
    #[inline]
    fn on_alloc(size: usize) {
        let live = LIVE.fetch_add(size, Ordering::Relaxed) + size;
        PEAK.fetch_max(live, Ordering::Relaxed);
        TOTAL.fetch_add(size, Ordering::Relaxed);
    }

    #[inline]
    fn on_dealloc(size: usize) {
        LIVE.fetch_sub(size, Ordering::Relaxed);
    }
}

unsafe impl GlobalAlloc for TrackingAllocator {
    unsafe fn alloc(&self, layout: Layout) -> *mut u8 {
        // SAFETY: forwarded to the system allocator; null is propagated.
        let ptr = unsafe { System.alloc(layout) };
        if !ptr.is_null() {
            Self::on_alloc(layout.size());
        }
        ptr
    }

    unsafe fn alloc_zeroed(&self, layout: Layout) -> *mut u8 {
        // SAFETY: forwarded to the system allocator; null is propagated.
        let ptr = unsafe { System.alloc_zeroed(layout) };
        if !ptr.is_null() {
            Self::on_alloc(layout.size());
        }
        ptr
    }

    unsafe fn dealloc(&self, ptr: *mut u8, layout: Layout) {
        Self::on_dealloc(layout.size());
        // SAFETY: pointer/layout come from the matching alloc call.
        unsafe { System.dealloc(ptr, layout) };
    }

    unsafe fn realloc(&self, ptr: *mut u8, layout: Layout, new_size: usize) -> *mut u8 {
        // SAFETY: forwarded to the system allocator; null is propagated.
        let new_ptr = unsafe { System.realloc(ptr, layout, new_size) };
        if !new_ptr.is_null() {
            if new_size > layout.size() {
                Self::on_alloc(new_size - layout.size());
            } else {
                Self::on_dealloc(layout.size() - new_size);
            }
        }
        new_ptr
    }
}

struct Sample {
    ms: f64,
    peak: usize,
    alloc: usize,
}

/// Reset the counters and run `f` inside a measurement window.
fn sample<T>(f: impl FnOnce() -> T) -> (T, Sample) {
    PEAK.store(LIVE.load(Ordering::Relaxed), Ordering::Relaxed);
    TOTAL.store(0, Ordering::Relaxed);
    let start = Instant::now();
    let out = f();
    let sample = Sample {
        ms: start.elapsed().as_secs_f64() * 1000.0,
        peak: PEAK.load(Ordering::Relaxed),
        alloc: TOTAL.load(Ordering::Relaxed),
    };
    (out, sample)
}

// ---------------------------------------------------------------------------
// Deterministic dataset (same shape as `parity_bench` / `write_comparison`).
// ---------------------------------------------------------------------------

fn headers(columns: usize) -> Vec<String> {
    const BASE: [&str; 7] = [
        "ID",
        "Name",
        "Count",
        "Score",
        "Date",
        "Active",
        "Description",
    ];
    (0..columns)
        .map(|column| {
            BASE.get(column)
                .map(|name| (*name).to_owned())
                .unwrap_or_else(|| format!("Column {column}"))
        })
        .collect()
}

fn date_at(index: usize, base_ms: i64) -> chrono::NaiveDateTime {
    Utc.timestamp_millis_opt(base_ms + index as i64 * 1000)
        .unwrap()
        .naive_utc()
}

fn cell_value(index: usize, column: usize, base_ms: i64) -> CellValue {
    const TAIL: &str = " z polskimi znakami: ąęśćńźółĄĘŚĆŃŹÓŁ oraz dłuższy tekst testowy.";
    let dt = date_at(index, base_ms);
    match column {
        0 => CellValue::Integer(index as i64),
        1 => CellValue::Text(format!("Produkt {index} żółć")),
        2 => CellValue::Integer(((index * 7919) % 10000) as i64),
        3 => CellValue::Number(((index * 104729) % 100000) as f64 / 1000.0),
        4 => CellValue::DateTime(dt),
        5 => CellValue::Boolean(index.is_multiple_of(2)),
        6 => CellValue::Text(format!("Opis produktu {index}{TAIL}")),
        column if column % 4 == 0 => CellValue::Integer((index + column) as i64),
        column if column % 4 == 1 => CellValue::Text(format!("C{column}-{index} żółć")),
        column if column % 4 == 2 => {
            CellValue::Number(((index * (column + 3)) % 100000) as f64 / 100.0)
        }
        _ => CellValue::Boolean((index + column).is_multiple_of(2)),
    }
}

fn build_data(rows: usize, columns: usize, base_ms: i64) -> Vec<Vec<CellValue>> {
    (0..rows)
        .map(|index| {
            (0..columns)
                .map(|column| cell_value(index, column, base_ms))
                .collect()
        })
        .collect()
}

/// Arrow representation of the same dataset (for `xlsb-writer`).
fn build_arrow_batch(rows: usize, columns: usize, base_ms: i64, headers: &[String]) -> RecordBatch {
    let schema = Arc::new(Schema::new(
        headers
            .iter()
            .enumerate()
            .map(|(column, name)| {
                let data_type = match column {
                    0 | 2 => DataType::Int64,
                    3 => DataType::Float64,
                    4 => DataType::Date64,
                    5 => DataType::Boolean,
                    _ => DataType::Utf8,
                };
                Field::new(name, data_type, false)
            })
            .collect::<Vec<_>>(),
    ));
    let mut columns_data: Vec<Arc<dyn arrow_array::Array>> = Vec::with_capacity(columns);
    for column in 0..columns {
        match column {
            0 | 2 => columns_data.push(Arc::new(Int64Array::from(
                (0..rows)
                    .map(|index| match cell_value(index, column, base_ms) {
                        CellValue::Integer(value) => value,
                        _ => unreachable!(),
                    })
                    .collect::<Vec<_>>(),
            ))),
            3 => columns_data.push(Arc::new(Float64Array::from(
                (0..rows)
                    .map(|index| match cell_value(index, column, base_ms) {
                        CellValue::Number(value) => value,
                        _ => unreachable!(),
                    })
                    .collect::<Vec<_>>(),
            ))),
            4 => columns_data.push(Arc::new(Date64Array::from(
                (0..rows)
                    .map(|index| date_at(index, base_ms).and_utc().timestamp_millis())
                    .collect::<Vec<_>>(),
            ))),
            5 => columns_data.push(Arc::new(BooleanArray::from(
                (0..rows).map(|index| index % 2 == 0).collect::<Vec<_>>(),
            ))),
            _ => columns_data.push(Arc::new(StringArray::from(
                (0..rows)
                    .map(|index| match cell_value(index, column, base_ms) {
                        CellValue::Text(value) => value,
                        _ => unreachable!(),
                    })
                    .collect::<Vec<_>>(),
            ))),
        }
    }
    RecordBatch::try_new(schema, columns_data).unwrap()
}

// ---------------------------------------------------------------------------
// The five writer paths.
// ---------------------------------------------------------------------------

fn write_our_batch(path: &Path, data: Vec<Vec<CellValue>>, headers: &[String], compression: i64) {
    let mut writer = OurXlsbWriter::create(path).unwrap();
    writer.set_compression_level(compression).unwrap();
    writer.add_sheet(SHEET_NAME, false);
    writer.write_sheet(data, Some(headers), true).unwrap();
    writer.finalize().unwrap();
}

fn write_our_streaming(path: &Path, data: &[Vec<CellValue>], headers: &[String], compression: i64) {
    let column_count = data.first().map_or(headers.len(), |row| row.len());
    let mut writer = OurXlsbWriter::create(path).unwrap();
    writer.set_compression_level(compression).unwrap();
    writer
        .start_sheet(
            SHEET_NAME,
            column_count,
            Some(headers),
            XlsbSheetOptions::new(),
        )
        .unwrap();
    for row in data {
        writer.write_row(row).unwrap();
    }
    writer.end_sheet().unwrap();
    writer.finalize().unwrap();
}

/// Map one dataset cell onto `rxlsb`'s owned `CellData` (its API takes
/// `Fn(usize, usize) -> CellData`, so text cells are cloned per call — that
/// is the library's normal usage cost and is intentionally timed).
fn rxlsb_cell(
    row: usize,
    column: usize,
    headers: &[String],
    data: &[Vec<CellValue>],
) -> rxlsb::CellData {
    if row == 0 {
        return rxlsb::CellData::text(headers[column].clone());
    }
    match &data[row - 1][column] {
        CellValue::Integer(value) => rxlsb::CellData::Number(*value as f64),
        CellValue::Number(value) => rxlsb::CellData::Number(*value),
        CellValue::Text(value) => rxlsb::CellData::text(value.clone()),
        CellValue::Boolean(value) => rxlsb::CellData::Bool(*value),
        CellValue::DateTime(value) => rxlsb::CellData::date(value.and_utc()),
        CellValue::Empty => rxlsb::CellData::Blank,
        CellValue::Formatted(formatted) => match &formatted.value {
            spreadsheet::PrimitiveCellValue::Integer(value) => {
                rxlsb::CellData::Number(*value as f64)
            }
            spreadsheet::PrimitiveCellValue::Number(value) => rxlsb::CellData::Number(*value),
            spreadsheet::PrimitiveCellValue::Text(value) => rxlsb::CellData::text(value.clone()),
            spreadsheet::PrimitiveCellValue::Boolean(value) => rxlsb::CellData::Bool(*value),
            spreadsheet::PrimitiveCellValue::DateTime(value) => {
                rxlsb::CellData::date(value.and_utc())
            }
            spreadsheet::PrimitiveCellValue::Empty => rxlsb::CellData::Blank,
        },
    }
}

fn write_rxlsb_batch(path: &Path, headers: &[String], data: &[Vec<CellValue>]) {
    let mut writer = rxlsb::XlsbWriter::builder().path(path).build().unwrap();
    let total_rows = data.len() + 1;
    writer
        .write_batch(
            SHEET_NAME,
            |row, column| rxlsb_cell(row, column, headers, data),
            total_rows,
            headers.len(),
        )
        .unwrap();
    writer.close().unwrap();
}

fn write_rxlsb_streaming(path: &Path, headers: &[String], data: &[Vec<CellValue>]) {
    let mut writer = rxlsb::XlsbWriter::builder().path(path).build().unwrap();
    writer.start_sheet(SHEET_NAME, headers.len()).unwrap();
    let total_rows = data.len() + 1;
    writer
        .write_rows(
            |row, column| rxlsb_cell(row, column, headers, data),
            0,
            total_rows,
        )
        .unwrap();
    writer.end_sheet().unwrap();
    writer.close().unwrap();
}

fn write_arrow_xlsb(path: &Path, batch: RecordBatch) {
    let file = File::create(path).unwrap();
    // Mirror the datetime column formatting of the other writers.
    let options = xlsb_writer::SheetOptions {
        column_formats: batch
            .schema()
            .fields()
            .iter()
            .enumerate()
            .map(|(column, _)| (column == 4).then(|| "datetime".to_owned()))
            .collect(),
        ..Default::default()
    };
    let mut writer = xlsb_writer::XlsbWriter::new(file);
    writer
        .write_sheet(SHEET_NAME, std::iter::once(batch), options)
        .unwrap();
    writer.finish().unwrap();
}

// ---------------------------------------------------------------------------
// Measurement harness.
// ---------------------------------------------------------------------------

struct CaseResult {
    name: String,
    path: PathBuf,
    times_ms: Vec<f64>,
    size: u64,
    peak: usize,
    alloc: usize,
}

impl CaseResult {
    fn median_ms(&self) -> f64 {
        median(self.times_ms.clone())
    }

    fn alloc_f64(&self) -> f64 {
        self.alloc as f64
    }
}

fn median(mut values: Vec<f64>) -> f64 {
    values.sort_by(|a, b| a.partial_cmp(b).unwrap());
    values[values.len() / 2]
}

fn measure_case<F>(
    name: impl Into<String>,
    path: &Path,
    iters: usize,
    mut iteration: F,
) -> CaseResult
where
    F: FnMut() -> Sample,
{
    let _ = iteration(); // warmup (allocator/code paths, filesystem cache)
    let mut times_ms = Vec::with_capacity(iters);
    let mut peak = 0usize;
    let mut alloc = 0usize;
    let mut size = 0u64;
    for _ in 0..iters {
        let sample = iteration();
        times_ms.push(sample.ms);
        peak = peak.max(sample.peak);
        alloc = alloc.max(sample.alloc);
        size = std::fs::metadata(path).unwrap().len();
    }
    CaseResult {
        name: name.into(),
        path: path.to_path_buf(),
        times_ms,
        size,
        peak,
        alloc,
    }
}

fn count_rows_cells(path: &Path) -> (usize, usize) {
    let mut reader = XlsbReader::new();
    reader.open(path, true).unwrap();
    let mut rows = 0;
    let mut cells = 0;
    while reader.read().unwrap() {
        rows += 1;
        cells += reader.current_row().len();
    }
    (rows, cells)
}

// ---------------------------------------------------------------------------
// Report helpers.
// ---------------------------------------------------------------------------

fn group_digits(mut value: u64) -> String {
    let mut digits = Vec::new();
    while value > 0 {
        digits.push((value % 10) as u8);
        value /= 10;
    }
    let mut out = String::new();
    for (index, digit) in digits.iter().rev().enumerate() {
        if index > 0 && (digits.len() - index) % 3 == 0 {
            out.push(' ');
        }
        out.push(char::from(b'0' + digit));
    }
    if out.is_empty() {
        "0".to_owned()
    } else {
        out
    }
}

fn mib(bytes: f64) -> f64 {
    bytes / (1024.0 * 1024.0)
}

fn time_factor(ours: f64, theirs: f64) -> String {
    let factor = theirs / ours;
    if factor >= 1.0 {
        format!("{factor:.2}× faster")
    } else {
        format!("{:.2}× slower", 1.0 / factor)
    }
}

fn size_factor(ours: f64, theirs: f64) -> String {
    let delta = (ours / theirs - 1.0) * 100.0;
    if delta <= 0.0 {
        format!("{:.1}% smaller", -delta)
    } else {
        format!("{delta:.1}% larger")
    }
}

// ---------------------------------------------------------------------------
// Main.
// ---------------------------------------------------------------------------

fn main() {
    let mut args = std::env::args().skip(1);
    let out: PathBuf = args
        .next()
        .map(PathBuf::from)
        .unwrap_or_else(|| std::env::temp_dir().join("xlsb_write_comparison"));
    let rows: usize = args
        .next()
        .unwrap_or_else(|| DEFAULT_ROWS.to_string())
        .parse()
        .expect("rows argument");
    let iters: usize = args
        .next()
        .unwrap_or_else(|| DEFAULT_ITERS.to_string())
        .parse()
        .expect("iters argument");
    let columns: usize = args
        .next()
        .unwrap_or_else(|| DEFAULT_COLUMNS.to_string())
        .parse()
        .expect("columns argument");
    let compression: i64 = args
        .next()
        .unwrap_or_else(|| "1".to_owned())
        .parse()
        .expect("compression argument");
    let iters = iters.max(1);
    std::fs::create_dir_all(&out).unwrap();

    let base_ms = Utc
        .with_ymd_and_hms(2024, 1, 15, 12, 0, 0)
        .unwrap()
        .timestamp_millis();
    let header_row = headers(columns);

    // Probe both input representations once for the report header.
    let before = LIVE.load(Ordering::Relaxed);
    let probe = build_data(rows, columns, base_ms);
    let dataset_heap = LIVE.load(Ordering::Relaxed) - before;
    drop(probe);
    let before = LIVE.load(Ordering::Relaxed);
    let arrow_probe = build_arrow_batch(rows, columns, base_ms, &header_row);
    let arrow_heap = LIVE.load(Ordering::Relaxed) - before;
    drop(arrow_probe);

    let cells = (rows + 1) * columns;
    let our_batch_path = out.join("spreadsheet-batch.xlsb");
    let our_stream_path = out.join("spreadsheet-streaming.xlsb");
    let rxlsb_batch_path = out.join("rxlsb-batch.xlsb");
    let rxlsb_stream_path = out.join("rxlsb-streaming.xlsb");
    let arrow_path = out.join("xlsb-writer-arrow.xlsb");
    let lvl = compression;

    let our_batch = measure_case(
        format!("spreadsheet batch (deflate {lvl})"),
        &our_batch_path,
        iters,
        || {
            let data = build_data(rows, columns, base_ms);
            let _ = std::fs::remove_file(&our_batch_path);
            let (_, sample) = sample(|| write_our_batch(&our_batch_path, data, &header_row, lvl));
            sample
        },
    );

    let our_stream = measure_case(
        format!("spreadsheet streaming (deflate {lvl})"),
        &our_stream_path,
        iters,
        || {
            let data = build_data(rows, columns, base_ms);
            let _ = std::fs::remove_file(&our_stream_path);
            let (_, sample) =
                sample(|| write_our_streaming(&our_stream_path, &data, &header_row, lvl));
            sample
        },
    );

    let rxlsb_batch = measure_case("rxlsb batch", &rxlsb_batch_path, iters, || {
        let data = build_data(rows, columns, base_ms);
        let _ = std::fs::remove_file(&rxlsb_batch_path);
        let (_, sample) = sample(|| write_rxlsb_batch(&rxlsb_batch_path, &header_row, &data));
        sample
    });

    let rxlsb_stream = measure_case("rxlsb streaming", &rxlsb_stream_path, iters, || {
        let data = build_data(rows, columns, base_ms);
        let _ = std::fs::remove_file(&rxlsb_stream_path);
        let (_, sample) = sample(|| write_rxlsb_streaming(&rxlsb_stream_path, &header_row, &data));
        sample
    });

    let arrow = measure_case("xlsb-writer (arrow)", &arrow_path, iters, || {
        let batch = build_arrow_batch(rows, columns, base_ms, &header_row);
        let _ = std::fs::remove_file(&arrow_path);
        let (_, sample) = sample(|| write_arrow_xlsb(&arrow_path, batch));
        sample
    });

    let results = [our_batch, our_stream, rxlsb_batch, rxlsb_stream, arrow];
    let expected_rows = rows + 1;
    for result in &results {
        let (read_rows, read_cells) = count_rows_cells(&result.path);
        assert_eq!(
            (read_rows, read_cells),
            (expected_rows, cells),
            "output mismatch for {}",
            result.name
        );
    }

    // ----------------------------------------------------------------- report
    println!("{}", "=".repeat(96));
    println!("XLSB write comparison: spreadsheet vs rxlsb vs xlsb-writer");
    println!("{}", "=".repeat(96));
    println!(
        "Dataset : {rows} data rows + header × {columns} columns = {cells} cells (deterministic, Polish text)"
    );
    println!(
        "Runs    : 1 warmup + {iters} measured iterations per case (median time; input built outside the window)"
    );
    println!(
        "Input   : spreadsheet/CellData heap ≈ {:.2} MiB; Arrow RecordBatch heap ≈ {:.2} MiB (excluded from allocation totals)",
        mib(dataset_heap as f64),
        mib(arrow_heap as f64)
    );
    println!(
        "Notes   : spreadsheet deflate level {lvl} (set_compression_level); xlsb-writer hardcodes level 6;"
    );
    println!("          rxlsb uses the zip crate default (~6); rust-xlsb-writer (GitHub-only pre-alpha) excluded.");

    let name_width = results
        .iter()
        .map(|result| result.name.len())
        .max()
        .unwrap_or(1);
    println!();
    println!(
        "| {:<name_width$} | {:>11} | {:>13} | {:>9} | {:>14} | {:>13} |",
        "Case", "Time (ms)", "Size (B)", "Size (MiB)", "Peak heap (MiB)", "Allocated (MiB)",
    );
    println!(
        "|{:-<name_width$}-+|{:-<11}-+|{:-<13}-+|{:-<9}-+|{:-<14}-+|{:-<13}-|",
        "", "", "", "", "", ""
    );
    for result in &results {
        println!(
            "| {:<name_width$} | {:>11.1} | {:>13} | {:>9.2} | {:>14.2} | {:>13.1} |",
            result.name,
            result.median_ms(),
            group_digits(result.size),
            mib(result.size as f64),
            mib(result.peak as f64),
            mib(result.alloc_f64()),
        );
    }

    println!("\nLike-for-like (median of {iters} runs, lower is better):");
    println!(
        "  time  batch          {:>8.1} ms vs rxlsb batch      {:>8.1} ms -> {}",
        results[0].median_ms(),
        results[2].median_ms(),
        time_factor(results[0].median_ms(), results[2].median_ms())
    );
    println!(
        "  time  streaming      {:>8.1} ms vs rxlsb streaming  {:>8.1} ms -> {}",
        results[1].median_ms(),
        results[3].median_ms(),
        time_factor(results[1].median_ms(), results[3].median_ms())
    );
    println!(
        "  time  batch          {:>8.1} ms vs xlsb-writer      {:>8.1} ms -> {}",
        results[0].median_ms(),
        results[4].median_ms(),
        time_factor(results[0].median_ms(), results[4].median_ms())
    );
    println!(
        "  size  batch          {:>13} B vs rxlsb batch      {:>13} B -> {}",
        group_digits(results[0].size),
        group_digits(results[2].size),
        size_factor(results[0].size as f64, results[2].size as f64)
    );
    println!(
        "  size  batch          {:>13} B vs xlsb-writer      {:>13} B -> {}",
        group_digits(results[0].size),
        group_digits(results[4].size),
        size_factor(results[0].size as f64, results[4].size as f64)
    );
    println!(
        "  heap  batch          {:>8.2} MiB vs rxlsb batch      {:>8.2} MiB -> {}",
        mib(results[0].peak as f64),
        mib(results[2].peak as f64),
        size_factor(results[0].peak as f64, results[2].peak as f64)
    );
    println!(
        "  heap  batch          {:>8.2} MiB vs xlsb-writer      {:>8.2} MiB -> {}",
        mib(results[0].peak as f64),
        mib(results[4].peak as f64),
        size_factor(results[0].peak as f64, results[4].peak as f64)
    );

    println!(
        "\nValidation: all {} outputs re-read OK ({expected_rows} rows × {columns} cells)",
        results.len()
    );

    print!("JSON     : ");
    print!(
        "{{\"impl\":\"spreadsheet-vs-rxlsb-vs-xlsb-writer\",\"rows\":{rows},\"columns\":{columns},\"iters\":{iters},\"compression\":{lvl},\"cases\":["
    );
    for (index, result) in results.iter().enumerate() {
        if index > 0 {
            print!(",");
        }
        print!(
            "{{\"name\":\"{}\",\"times_ms\":{:?},\"median_ms\":{},\"size\":{},\"peak_bytes\":{},\"alloc_bytes\":{}}}",
            result.name,
            result.times_ms,
            result.median_ms(),
            result.size,
            result.peak,
            result.alloc
        );
    }
    println!("]}}");
}
