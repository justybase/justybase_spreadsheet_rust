//! XLSX write comparison: `spreadsheet` vs `rust_xlsxwriter`.
//!
//! Measures, for four writer paths, the median write time, the output file
//! size, and heap usage (peak live heap and total allocated bytes) on one
//! deterministic dataset:
//!
//!  1. `spreadsheet batch` — `add_sheet` + `write_sheet` + `finalize`;
//!  2. `spreadsheet streaming` — `start_sheet` + `write_row` + `end_sheet` + `finalize`;
//!  3. `rust_xlsxwriter default` — in-memory workbook + `save`;
//!  4. `rust_xlsxwriter constant_memory` — `add_worksheet_with_constant_memory` + `save`.
//!
//! Methodology:
//!
//! * every iteration builds a fresh dataset *outside* the measurement window,
//!   so exactly one dataset copy is live during each write and the build cost
//!   is neither timed nor counted;
//! * the window covers writer creation up to `finalize()`/`save()`, including
//!   the file write;
//! * one warmup run per case, then N runs (default 3); the median time is
//!   reported, the maximum peak heap, and the median allocation total;
//! * memory comes from a counting `GlobalAllocator` — peak live heap within
//!   the window and the sum of all bytes allocated inside it (the two atomic
//!   updates per allocation hit both libraries equally);
//! * all four outputs are re-read with `XlsxReader` and the row/cell counts
//!   are asserted before the report is printed.
//!
//! Fairness notes (mirrored in BENCHMARKS.md):
//!
//! * compression: this crate uses zip Deflate level 1 (project default),
//!   `rust_xlsxwriter` uses the zip crate default level 6 (not configurable),
//!   so its default-mode files compress slightly better at a higher CPU cost;
//! * equivalent workbook features are enabled on both sides: bold header,
//!   frozen first row, autofilter over the full range, builtin date format
//!   (`numFmtId` 14);
//! * `rust_xlsxwriter` writes full `r="A1"` cell references while this crate
//!   writes dense XML, and its `constant_memory` mode stores strings as
//!   inline strings instead of using the shared-string table;
//! * this crate additionally writes autofitted `<cols>` widths (small size
//!   bonus for `rust_xlsxwriter`).
//!
//! Usage:
//!
//! ```text
//! cargo run --release --example write_comparison [outDir] [rows] [iters] [columns]
//! ```
//!
//! Defaults: system temp dir, 5000 rows, 3 iterations, 7 columns. Pass larger
//! values explicitly when a full-size comparison is needed.

use std::alloc::{GlobalAlloc, Layout, System};
use std::path::{Path, PathBuf};
use std::sync::atomic::{AtomicUsize, Ordering};
use std::time::Instant;

use chrono::{TimeZone, Utc};
use rust_xlsxwriter::{Format, Workbook};
use spreadsheet::{CellValue, PrimitiveCellValue, SheetOptions, XlsxReader, XlsxWriter};

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

struct Sample {
    ms: f64,
    peak: usize,
    alloc: usize,
}

// ---------------------------------------------------------------------------
// Deterministic dataset (same shape as `parity_bench`).
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

fn build_data(rows: usize, columns: usize, base_ms: i64) -> Vec<Vec<CellValue>> {
    const TAIL: &str = " z polskimi znakami: ąęśćńźółĄĘŚĆŃŹÓŁ oraz dłuższy tekst testowy.";
    let mut data = Vec::with_capacity(rows);
    for i in 0..rows {
        let dt = Utc
            .timestamp_millis_opt(base_ms + i as i64 * 1000)
            .unwrap()
            .naive_utc();
        let mut row = Vec::with_capacity(columns);
        for column in 0..columns {
            let value = match column {
                0 => CellValue::Integer(i as i64),
                1 => CellValue::Text(format!("Produkt {i} żółć")),
                2 => CellValue::Integer(((i * 7919) % 10000) as i64),
                3 => CellValue::Number(((i * 104729) % 100000) as f64 / 1000.0),
                4 => CellValue::DateTime(dt),
                5 => CellValue::Boolean(i % 2 == 0),
                6 => CellValue::Text(format!("Opis produktu {i}{TAIL}")),
                column if column % 4 == 0 => CellValue::Integer((i + column) as i64),
                column if column % 4 == 1 => CellValue::Text(format!("C{column}-{i} żółć")),
                column if column % 4 == 2 => {
                    CellValue::Number(((i * (column + 3)) % 100000) as f64 / 100.0)
                }
                _ => CellValue::Boolean((i + column) % 2 == 0),
            };
            row.push(value);
        }
        data.push(row);
    }
    data
}

// ---------------------------------------------------------------------------
// The four writer paths.
// ---------------------------------------------------------------------------

fn write_spreadsheet_batch(path: &Path, data: Vec<Vec<CellValue>>, headers: &[String]) {
    let mut writer = XlsxWriter::create(path).unwrap();
    writer.add_sheet(SHEET_NAME, false);
    writer.write_sheet(data, Some(headers), true).unwrap();
    writer.finalize().unwrap();
}

fn write_spreadsheet_streaming(path: &Path, data: &[Vec<CellValue>], headers: &[String]) {
    let column_count = data.first().map_or(headers.len(), |row| row.len());
    let mut writer = XlsxWriter::create(path).unwrap();
    writer
        .start_sheet(SHEET_NAME, column_count, Some(headers), SheetOptions::new())
        .unwrap();
    for row in data {
        writer.write_row(row).unwrap();
    }
    writer.end_sheet().unwrap();
    writer.finalize().unwrap();
}

fn write_rust_xlsxwriter(
    path: &Path,
    data: &[Vec<CellValue>],
    headers: &[String],
    constant_memory: bool,
) {
    let mut workbook = Workbook::new();
    let worksheet = if constant_memory {
        workbook.add_worksheet_with_constant_memory()
    } else {
        workbook.add_worksheet()
    };
    worksheet.set_name(SHEET_NAME).unwrap();
    // Mirror this crate: frozen header row, bold headers, autofilter, and the
    // builtin short-date format (numFmtId 14) used for plain DateTime cells.
    worksheet.set_freeze_panes(1, 0).unwrap();
    let header_format = Format::new().set_bold();
    let date_format = Format::new().set_num_format_index(14);
    let blank_format = Format::new();
    for (column, header) in headers.iter().enumerate() {
        worksheet
            .write_with_format(0, column as u16, header, &header_format)
            .unwrap();
    }
    for (index, row) in data.iter().enumerate() {
        let r = index as u32 + 1;
        for (column, value) in row.iter().enumerate() {
            let c = column as u16;
            match value {
                CellValue::Empty => {
                    worksheet.write_blank(r, c, &blank_format).unwrap();
                }
                CellValue::Integer(value) => {
                    worksheet.write(r, c, *value).unwrap();
                }
                CellValue::Number(value) => {
                    worksheet.write(r, c, *value).unwrap();
                }
                CellValue::Text(value) => {
                    worksheet.write(r, c, value.as_str()).unwrap();
                }
                CellValue::Boolean(value) => {
                    worksheet.write(r, c, *value).unwrap();
                }
                CellValue::DateTime(value) => {
                    worksheet
                        .write_datetime_with_format(r, c, value, &date_format)
                        .unwrap();
                }
                CellValue::Formatted(formatted) => {
                    let format = Format::new().set_num_format(formatted.format.clone());
                    match &formatted.value {
                        PrimitiveCellValue::Empty => {
                            worksheet.write_blank(r, c, &format).unwrap();
                        }
                        PrimitiveCellValue::Integer(value) => {
                            worksheet.write_with_format(r, c, *value, &format).unwrap();
                        }
                        PrimitiveCellValue::Number(value) => {
                            worksheet.write_with_format(r, c, *value, &format).unwrap();
                        }
                        PrimitiveCellValue::Text(value) => {
                            worksheet
                                .write_with_format(r, c, value.as_str(), &format)
                                .unwrap();
                        }
                        PrimitiveCellValue::Boolean(value) => {
                            worksheet.write_with_format(r, c, *value, &format).unwrap();
                        }
                        PrimitiveCellValue::DateTime(value) => {
                            worksheet
                                .write_datetime_with_format(r, c, value, &format)
                                .unwrap();
                        }
                    }
                }
            }
        }
    }
    let last_row = data.len() as u32;
    let last_column = headers.len().saturating_sub(1) as u16;
    worksheet.autofilter(0, 0, last_row, last_column).unwrap();
    workbook.save(path).unwrap();
}

// ---------------------------------------------------------------------------
// Measurement harness.
// ---------------------------------------------------------------------------

struct CaseResult {
    name: &'static str,
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

    fn median_alloc(&self) -> f64 {
        median(vec![self.alloc as f64])
    }
}

fn median(mut values: Vec<f64>) -> f64 {
    values.sort_by(|a, b| a.partial_cmp(b).unwrap());
    values[values.len() / 2]
}

fn measure_case<F>(name: &'static str, path: &Path, iters: usize, mut iteration: F) -> CaseResult
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
        name,
        path: path.to_path_buf(),
        times_ms,
        size,
        peak,
        alloc,
    }
}

fn count_rows_cells(path: &Path) -> (usize, usize) {
    let mut reader = XlsxReader::new();
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
        .unwrap_or_else(|| std::env::temp_dir().join("write_comparison"));
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
    let iters = iters.max(1);
    std::fs::create_dir_all(&out).unwrap();

    let base_ms = Utc
        .with_ymd_and_hms(2024, 1, 15, 12, 0, 0)
        .unwrap()
        .timestamp_millis();
    let header_row = headers(columns);

    // Probe the dataset heap footprint once for the report header.
    let before = LIVE.load(Ordering::Relaxed);
    let probe = build_data(rows, columns, base_ms);
    let dataset_heap = LIVE.load(Ordering::Relaxed) - before;
    let cells = (rows + 1) * columns;
    drop(probe);

    let batch_path = out.join("spreadsheet-batch.xlsx");
    let streaming_path = out.join("spreadsheet-streaming.xlsx");
    let default_path = out.join("rust-xlsxwriter-default.xlsx");
    let constant_path = out.join("rust-xlsxwriter-constant-memory.xlsx");

    // Case 1: this crate, batch API. `write_sheet` takes ownership, so each
    // iteration builds its dataset outside the measurement window.
    let batch = measure_case("spreadsheet batch", &batch_path, iters, || {
        let data = build_data(rows, columns, base_ms);
        let _ = std::fs::remove_file(&batch_path);
        let (_, sample) = sample(|| write_spreadsheet_batch(&batch_path, data, &header_row));
        sample
    });

    // Case 2: this crate, streaming API.
    let streaming = measure_case("spreadsheet streaming", &streaming_path, iters, || {
        let data = build_data(rows, columns, base_ms);
        let _ = std::fs::remove_file(&streaming_path);
        let (_, sample) =
            sample(|| write_spreadsheet_streaming(&streaming_path, &data, &header_row));
        sample
    });

    // Case 3: rust_xlsxwriter, default in-memory mode.
    let default = measure_case("rust_xlsxwriter default", &default_path, iters, || {
        let data = build_data(rows, columns, base_ms);
        let _ = std::fs::remove_file(&default_path);
        let (_, sample) =
            sample(|| write_rust_xlsxwriter(&default_path, &data, &header_row, false));
        sample
    });

    // Case 4: rust_xlsxwriter, constant-memory worksheets.
    let constant = measure_case(
        "rust_xlsxwriter constant_memory",
        &constant_path,
        iters,
        || {
            let data = build_data(rows, columns, base_ms);
            let _ = std::fs::remove_file(&constant_path);
            let (_, sample) =
                sample(|| write_rust_xlsxwriter(&constant_path, &data, &header_row, true));
            sample
        },
    );

    // Validate every output before reporting anything.
    let results = [batch, streaming, default, constant];
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
    println!("XLSX write comparison: spreadsheet vs rust_xlsxwriter");
    println!("{}", "=".repeat(96));
    println!(
        "Dataset : {rows} data rows + header × {columns} columns = {cells} cells (deterministic, Polish text)"
    );
    println!(
        "Runs    : 1 warmup + {iters} measured iterations per case (median time; dataset built outside the window)"
    );
    println!(
        "Heap    : dataset ≈ {:.2} MiB — live during every write, excluded from allocation totals",
        mib(dataset_heap as f64)
    );
    println!(
        "Notes   : Deflate level 1 (spreadsheet default) vs level 6 (rust_xlsxwriter zip default,"
    );
    println!("          not configurable); dense XML (no r=\"A1\") vs full cell references; constant_memory");
    println!("          stores strings as inline strings instead of the shared-string table.");

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
            mib(result.median_alloc()),
        );
    }

    let batch_ms = results[0].median_ms();
    let streaming_ms = results[1].median_ms();
    let default_ms = results[2].median_ms();
    let constant_ms = results[3].median_ms();
    println!("\nLike-for-like (median of {iters} runs, lower is better):");
    println!(
        "  time      batch     {:>8.1} ms vs default          {:>8.1} ms -> {}",
        batch_ms,
        default_ms,
        time_factor(batch_ms, default_ms)
    );
    println!(
        "  time      streaming {:>8.1} ms vs constant_memory  {:>8.1} ms -> {}",
        streaming_ms,
        constant_ms,
        time_factor(streaming_ms, constant_ms)
    );
    println!(
        "  size      batch     {:>13} B vs default          {:>13} B -> {}",
        group_digits(results[0].size),
        group_digits(results[2].size),
        size_factor(results[0].size as f64, results[2].size as f64)
    );
    println!(
        "  size      streaming {:>13} B vs constant_memory  {:>13} B -> {}",
        group_digits(results[1].size),
        group_digits(results[3].size),
        size_factor(results[1].size as f64, results[3].size as f64)
    );
    println!(
        "  peak heap batch     {:>8.2} MiB vs default          {:>8.2} MiB -> {}",
        mib(results[0].peak as f64),
        mib(results[2].peak as f64),
        size_factor(results[0].peak as f64, results[2].peak as f64)
    );
    println!(
        "  peak heap streaming {:>8.2} MiB vs constant_memory  {:>8.2} MiB -> {}",
        mib(results[1].peak as f64),
        mib(results[3].peak as f64),
        size_factor(results[1].peak as f64, results[3].peak as f64)
    );

    println!(
        "\nValidation: all {} outputs re-read OK ({expected_rows} rows × {columns} cells)",
        results.len()
    );

    print!("JSON     : ");
    print!("{{\"impl\":\"spreadsheet-vs-rust_xlsxwriter\",\"rows\":{rows},\"columns\":{columns},\"iters\":{iters},\"cases\":[");
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
