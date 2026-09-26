use calamine::{open_workbook, Data, DataRef, Reader, Xlsb, Xlsx};
use chrono::{TimeZone, Utc};
use spreadsheet::{
    datetime_to_excel_serial, CellValue, CellValueRef, SheetOptions, XlsbReader, XlsbSheetOptions,
    XlsbWriter, XlsxReader, XlsxWriter,
};
use std::env;
use std::fs;
use std::path::Path;
use std::time::Instant;

const HEADERS: [&str; 7] = [
    "ID",
    "Name",
    "Count",
    "Score",
    "Date",
    "Active",
    "Description",
];
const TAIL: &str = " z polskimi znakami: ąęśćńźółĄĘŚĆŃŹÓŁ oraz dłuższy tekst testowy.";

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
struct Summary {
    rows: usize,
    cells: usize,
    hash: u64,
    column_hashes: [u64; 7],
    active_col: usize,
    date_max_delta_ms: i64,
    date_mismatches: usize,
}

impl Summary {
    fn new() -> Self {
        Self {
            rows: 0,
            cells: 0,
            hash: 0xcbf29ce484222325,
            column_hashes: [0xcbf29ce484222325; 7],
            active_col: 0,
            date_max_delta_ms: 0,
            date_mismatches: 0,
        }
    }

    fn byte(&mut self, byte: u8) {
        self.hash ^= u64::from(byte);
        self.hash = self.hash.wrapping_mul(0x100000001b3);
        if self.active_col < self.column_hashes.len() {
            self.column_hashes[self.active_col] ^= u64::from(byte);
            self.column_hashes[self.active_col] =
                self.column_hashes[self.active_col].wrapping_mul(0x100000001b3);
        }
    }

    fn bytes(&mut self, bytes: &[u8]) {
        for byte in bytes {
            self.byte(*byte);
        }
    }

    fn u32(&mut self, value: u32) {
        self.bytes(&value.to_le_bytes());
    }

    fn u64(&mut self, value: u64) {
        self.bytes(&value.to_le_bytes());
    }

    fn cell_prefix(&mut self, row: u32, col: u32, tag: u8) {
        self.active_col = col as usize;
        self.u32(row);
        self.u32(col);
        self.byte(tag);
        self.cells += 1;
    }

    fn number(&mut self, row: u32, col: u32, value: f64) {
        self.cell_prefix(row, col, 2);
        self.u64(value.to_bits());
    }

    fn text(&mut self, row: u32, col: u32, value: &str) {
        self.cell_prefix(row, col, 3);
        self.u64(value.len() as u64);
        self.bytes(value.as_bytes());
    }

    fn boolean(&mut self, row: u32, col: u32, value: bool) {
        self.cell_prefix(row, col, 4);
        self.byte(u8::from(value));
    }

    fn empty(&mut self, row: u32, col: u32) {
        self.cell_prefix(row, col, 1);
    }

    fn date(&mut self, row: u32, col: u32, serial_ms: i64) {
        self.cell_prefix(row, col, 5);
        const BASE_DATE_SERIAL_MS: i64 = 45_306 * 86_400_000 + 43_200_000;
        if row > 0 && col == 4 {
            let expected = BASE_DATE_SERIAL_MS + (row as i64 - 1) * 1000;
            let delta = (serial_ms - expected).abs();
            self.date_max_delta_ms = self.date_max_delta_ms.max(delta);
            self.date_mismatches += usize::from(delta > 1);
        }
        // Runtime date conversions can differ by a millisecond. Measure that
        // drift separately and use second precision for the cross-reader hash.
        self.u64(((serial_ms + 500) / 1000 * 1000) as u64);
    }
}

fn hash_owned(summary: &mut Summary, row: u32, col: u32, value: &CellValue) {
    match value {
        CellValue::Empty => summary.empty(row, col),
        CellValue::Text(value) => summary.text(row, col, value),
        CellValue::Number(value) => summary.number(row, col, *value),
        CellValue::Integer(value) => summary.number(row, col, *value as f64),
        CellValue::Boolean(value) => summary.boolean(row, col, *value),
        CellValue::DateTime(value) => summary.date(
            row,
            col,
            (datetime_to_excel_serial(value, false) * 86_400_000.0).round() as i64,
        ),
        CellValue::Formatted(_) => panic!("unexpected formatted cell"),
    }
}

fn hash_ref(summary: &mut Summary, row: u32, col: u32, value: CellValueRef<'_>) {
    match value {
        CellValueRef::Empty => summary.empty(row, col),
        CellValueRef::Text(value) => summary.text(row, col, value),
        CellValueRef::Number(value) => summary.number(row, col, value),
        CellValueRef::Integer(value) => summary.number(row, col, value as f64),
        CellValueRef::Boolean(value) => summary.boolean(row, col, value),
        CellValueRef::DateTime(value) => summary.date(
            row,
            col,
            (datetime_to_excel_serial(&value, false) * 86_400_000.0).round() as i64,
        ),
    }
}

fn hash_calamine_owned(summary: &mut Summary, row: u32, col: u32, value: &Data) {
    match value {
        Data::Empty => summary.empty(row, col),
        Data::Int(value) => summary.number(row, col, *value as f64),
        Data::Float(value) => summary.number(row, col, *value),
        Data::String(value) => summary.text(row, col, value),
        Data::Bool(value) => summary.boolean(row, col, *value),
        Data::DateTime(value) => {
            summary.date(row, col, (value.as_f64() * 86_400_000.0).round() as i64);
        }
        other => panic!("unexpected generated cell {other:?}"),
    }
}

fn hash_calamine_ref(summary: &mut Summary, row: u32, col: u32, value: &DataRef<'_>) {
    match value {
        DataRef::Empty => summary.empty(row, col),
        DataRef::Int(value) => summary.number(row, col, *value as f64),
        DataRef::Float(value) => summary.number(row, col, *value),
        DataRef::String(value) => summary.text(row, col, value),
        DataRef::SharedString(value) => summary.text(row, col, value),
        DataRef::Bool(value) => summary.boolean(row, col, *value),
        DataRef::DateTime(value) => {
            summary.date(row, col, (value.as_f64() * 86_400_000.0).round() as i64);
        }
        other => panic!("unexpected generated cell {other:?}"),
    }
}

fn build_row(i: usize, include_dates: bool) -> Vec<CellValue> {
    let base_ms = Utc
        .with_ymd_and_hms(2024, 1, 15, 12, 0, 0)
        .unwrap()
        .timestamp_millis();
    let dt = Utc
        .timestamp_millis_opt(base_ms + i as i64 * 1000)
        .unwrap()
        .naive_utc();
    vec![
        CellValue::Integer(i as i64),
        CellValue::Text(format!("Produkt {i} żółć")),
        CellValue::Integer(((i * 7919) % 10_000) as i64),
        CellValue::Number(((i * 104_729) % 100_000) as f64 / 1000.0),
        if include_dates {
            CellValue::DateTime(dt)
        } else {
            CellValue::Number(45_306.5 + i as f64 / 86_400.0)
        },
        CellValue::Boolean(i.is_multiple_of(2)),
        CellValue::Text(format!("Opis produktu {i}{TAIL}")),
    ]
}

fn generate(out: &Path, rows: usize, include_dates: bool) {
    fs::create_dir_all(out).unwrap();
    let headers: Vec<String> = HEADERS.iter().map(|s| (*s).to_owned()).collect();
    let suffix = if include_dates { "" } else { "-no-dates" };

    let xlsx = out.join(format!("rust-{rows}{suffix}.xlsx"));
    let mut writer = XlsxWriter::create(&xlsx).unwrap();
    writer
        .start_sheet(
            "Benchmark",
            HEADERS.len(),
            Some(&headers),
            SheetOptions::new(),
        )
        .unwrap();
    for i in 0..rows {
        writer.write_row(&build_row(i, include_dates)).unwrap();
    }
    writer.end_sheet().unwrap();
    writer.finalize().unwrap();

    let xlsb = out.join(format!("rust-{rows}{suffix}.xlsb"));
    let mut writer = XlsbWriter::create(&xlsb).unwrap();
    writer
        .start_sheet(
            "Benchmark",
            HEADERS.len(),
            Some(&headers),
            XlsbSheetOptions::new(),
        )
        .unwrap();
    for i in 0..rows {
        writer.write_row(&build_row(i, include_dates)).unwrap();
    }
    writer.end_sheet().unwrap();
    writer.finalize().unwrap();

    println!(
        "{{\"rows\":{rows},\"xlsx_bytes\":{},\"xlsb_bytes\":{}}}",
        fs::metadata(xlsx).unwrap().len(),
        fs::metadata(xlsb).unwrap().len()
    );
}

fn scan_rust_xlsx_row(path: &Path) -> Summary {
    let mut reader = XlsxReader::new();
    reader.open(path, true).unwrap();
    reader.select_sheet("Benchmark").unwrap();
    let mut summary = Summary::new();
    while reader.read().unwrap() {
        let row = reader.current_row();
        let row_index = summary.rows as u32;
        for (col, value) in row.iter().enumerate() {
            hash_owned(&mut summary, row_index, col as u32, value);
        }
        summary.rows += 1;
    }
    summary
}

fn scan_rust_xlsx_cell(path: &Path) -> Summary {
    let mut reader = XlsxReader::new();
    reader.open(path, true).unwrap();
    let mut cells = reader.cell_reader("Benchmark").unwrap();
    let mut summary = Summary::new();
    while let Some(cell) = cells.next_cell().unwrap() {
        hash_ref(&mut summary, cell.row, cell.column, cell.value);
        summary.rows = summary.rows.max(cell.row as usize + 1);
    }
    summary
}

fn scan_rust_xlsb_row(path: &Path) -> Summary {
    let mut reader = XlsbReader::new();
    reader.open(path, true).unwrap();
    reader.select_sheet("Benchmark").unwrap();
    let mut summary = Summary::new();
    while reader.read().unwrap() {
        let row = reader.current_row();
        let row_index = summary.rows as u32;
        for (col, value) in row.iter().enumerate() {
            hash_owned(&mut summary, row_index, col as u32, value);
        }
        summary.rows += 1;
    }
    summary
}

fn scan_rust_xlsb_cell(path: &Path) -> Summary {
    let mut reader = XlsbReader::new();
    reader.open(path, true).unwrap();
    let mut cells = reader.cell_reader("Benchmark").unwrap();
    let mut summary = Summary::new();
    while let Some(cell) = cells.next_cell().unwrap() {
        hash_ref(&mut summary, cell.row, cell.column, cell.value);
        summary.rows = summary.rows.max(cell.row as usize + 1);
    }
    summary
}

fn scan_calamine_range(path: &Path, xlsb: bool) -> Summary {
    let mut summary = Summary::new();
    if xlsb {
        let mut workbook: Xlsb<_> = open_workbook(path).unwrap();
        let range = workbook.worksheet_range("Benchmark").unwrap();
        for (row_idx, row) in range.rows().enumerate() {
            for (col_idx, value) in row.iter().enumerate() {
                hash_calamine_owned(&mut summary, row_idx as u32, col_idx as u32, value);
            }
            summary.rows += 1;
        }
    } else {
        let mut workbook: Xlsx<_> = open_workbook(path).unwrap();
        let range = workbook.worksheet_range("Benchmark").unwrap();
        for (row_idx, row) in range.rows().enumerate() {
            for (col_idx, value) in row.iter().enumerate() {
                hash_calamine_owned(&mut summary, row_idx as u32, col_idx as u32, value);
            }
            summary.rows += 1;
        }
    }
    summary
}

fn scan_calamine_cells(path: &Path, xlsb: bool) -> Summary {
    let mut summary = Summary::new();
    if xlsb {
        let mut workbook: Xlsb<_> = open_workbook(path).unwrap();
        let mut cells = workbook.worksheet_cells_reader("Benchmark").unwrap();
        while let Some(cell) = cells.next_cell().unwrap() {
            let (row, col) = cell.get_position();
            hash_calamine_ref(&mut summary, row, col, cell.get_value());
            summary.rows = summary.rows.max(row as usize + 1);
        }
    } else {
        let mut workbook: Xlsx<_> = open_workbook(path).unwrap();
        let mut cells = workbook.worksheet_cells_reader("Benchmark").unwrap();
        while let Some(cell) = cells.next_cell().unwrap() {
            let (row, col) = cell.get_position();
            hash_calamine_ref(&mut summary, row, col, cell.get_value());
            summary.rows = summary.rows.max(row as usize + 1);
        }
    }
    summary
}

fn scan(mode: &str, path: &Path) -> Summary {
    match mode {
        "rust-xlsx-row" => scan_rust_xlsx_row(path),
        "rust-xlsx-cell" => scan_rust_xlsx_cell(path),
        "rust-xlsb-row" => scan_rust_xlsb_row(path),
        "rust-xlsb-cell" => scan_rust_xlsb_cell(path),
        "calamine-xlsx-range" => scan_calamine_range(path, false),
        "calamine-xlsx-cell" => scan_calamine_cells(path, false),
        "calamine-xlsb-range" => scan_calamine_range(path, true),
        "calamine-xlsb-cell" => scan_calamine_cells(path, true),
        _ => panic!("unknown scan mode {mode}"),
    }
}

fn main() {
    let args: Vec<String> = env::args().collect();
    match args.get(1).map(String::as_str) {
        Some("generate") => {
            let out = Path::new(args.get(2).expect("output directory"));
            let rows = args.get(3).expect("row count").parse().unwrap();
            let include_dates = !args.get(4).is_some_and(|arg| arg == "--no-dates");
            generate(out, rows, include_dates);
        }
        Some("scan") => {
            let mode = args.get(2).expect("mode");
            let path = Path::new(args.get(3).expect("input path"));
            let once = args.get(4).is_some_and(|arg| arg == "--once");
            let repeats = if once {
                1
            } else {
                args.get(4).map(|s| s.parse().unwrap()).unwrap_or(5usize)
            };
            assert!(repeats > 0, "repeat count must be positive");
            let mut times_ms = Vec::new();
            let mut expected = None;
            if !once {
                let warm = scan(mode, path);
                expected = Some(warm);
            }
            for _ in 0..if once { 1 } else { repeats } {
                let start = Instant::now();
                let got = scan(mode, path);
                let elapsed = start.elapsed().as_secs_f64() * 1000.0;
                if let Some(want) = expected {
                    assert_eq!(got, want, "repeated scan changed output");
                } else {
                    expected = Some(got);
                }
                times_ms.push(elapsed);
            }
            let result = expected.unwrap();
            let median = {
                let mut sorted = times_ms.clone();
                sorted.sort_by(f64::total_cmp);
                sorted[sorted.len() / 2]
            };
            println!("{{\"mode\":\"{mode}\",\"runs_ms\":{:?},\"median_ms\":{median:.3},\"rows\":{},\"cells\":{},\"checksum\":\"{:016x}\",\"columns\":[{}],\"date_max_delta_ms\":{},\"date_mismatches\":{}}}",
                times_ms, result.rows, result.cells, result.hash,
                result.column_hashes.iter().map(|h| format!("\"{h:016x}\"")).collect::<Vec<_>>().join(","),
                result.date_max_delta_ms, result.date_mismatches);
        }
        _ => panic!("usage: read_scale_bench generate <outdir> <rows> [--no-dates] | scan <mode> <path> [--once|<repeat-count>]"),
    }
}
