//! Benchmarks for the Rust implementation (mirrors `parity/bench_ts.ts`).
//!
//! Deterministic dataset, batch mode (`add_sheet` + `write_sheet` +
//! `finalize`), full-scan reads. Prints one JSON document to stdout.
//!
//! Usage: `cargo run --release -p justybase-spreadsheet --example parity_bench <outDir> [rows] [iters] [columns]`

use chrono::{NaiveDate, TimeZone, Utc};
use spreadsheet::{
    CellValue, SheetOptions, XlsbReader, XlsbSheetOptions, XlsbWriter, XlsxReader, XlsxWriter,
};
use std::path::PathBuf;
use std::time::Instant;

const DEFAULT_ROWS: usize = 5_000;
const DEFAULT_ITERS: usize = 3;
const DEFAULT_COLUMNS: usize = 7;

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

fn median(mut xs: Vec<f64>) -> f64 {
    xs.sort_by(|a, b| a.partial_cmp(b).unwrap());
    xs[xs.len() / 2]
}

fn checksum_f64(v: &CellValue, acc: &mut f64) {
    match v {
        CellValue::Number(n) => *acc += *n,
        CellValue::Integer(n) => *acc += *n as f64,
        CellValue::Text(s) => *acc += s.len() as f64,
        CellValue::Boolean(b) => *acc += if *b { 1.0 } else { 0.0 },
        CellValue::DateTime(d) => {
            *acc += (d.and_utc().timestamp_millis() % 1000) as f64;
        }
        CellValue::Empty => {}
        CellValue::Formatted(f) => checksum_primitive(&f.value, acc),
    }
}

fn checksum_primitive(v: &spreadsheet::PrimitiveCellValue, acc: &mut f64) {
    use spreadsheet::PrimitiveCellValue as P;
    match v {
        P::Number(n) => *acc += *n,
        P::Integer(n) => *acc += *n as f64,
        P::Text(s) => *acc += s.len() as f64,
        P::Boolean(b) => *acc += if *b { 1.0 } else { 0.0 },
        P::DateTime(d) => *acc += (d.and_utc().timestamp_millis() % 1000) as f64,
        P::Empty => {}
    }
}

fn main() {
    let out = PathBuf::from(std::env::args().nth(1).expect("outDir argument"));
    let rows: usize = std::env::args()
        .nth(2)
        .unwrap_or_else(|| DEFAULT_ROWS.to_string())
        .parse()
        .unwrap();
    let iters: usize = std::env::args()
        .nth(3)
        .unwrap_or_else(|| DEFAULT_ITERS.to_string())
        .parse()
        .unwrap();
    let columns: usize = std::env::args()
        .nth(4)
        .unwrap_or_else(|| DEFAULT_COLUMNS.to_string())
        .parse()
        .unwrap();
    std::fs::create_dir_all(&out).unwrap();
    let _ = NaiveDate::from_ymd_opt(2024, 1, 1).unwrap();
    let base_ms = Utc
        .with_ymd_and_hms(2024, 1, 15, 12, 0, 0)
        .unwrap()
        .timestamp_millis();
    let data = build_data(rows, columns, base_ms);
    let headers = headers(columns);

    let mut w_xlsb = Vec::with_capacity(iters);
    for _ in 0..iters {
        let file = out.join("bench-rs.xlsb");
        let _ = std::fs::remove_file(&file);
        let t0 = Instant::now();
        let mut w = XlsbWriter::create(&file).unwrap();
        w.add_sheet("Benchmark", false);
        w.write_sheet(data.clone(), Some(&headers), true).unwrap();
        w.finalize().unwrap();
        w_xlsb.push(t0.elapsed().as_secs_f64() * 1000.0);
    }
    let mut w_xlsx = Vec::with_capacity(iters);
    for _ in 0..iters {
        let file = out.join("bench-rs.xlsx");
        let _ = std::fs::remove_file(&file);
        let t0 = Instant::now();
        let mut w = XlsxWriter::create(&file).unwrap();
        w.add_sheet("Benchmark", false);
        w.write_sheet(data.clone(), Some(&headers), true).unwrap();
        w.finalize().unwrap();
        w_xlsx.push(t0.elapsed().as_secs_f64() * 1000.0);
    }

    let mut ws_xlsb = Vec::with_capacity(iters);
    for _ in 0..iters {
        let file = out.join("bench-rs-stream.xlsb");
        let _ = std::fs::remove_file(&file);
        let t0 = Instant::now();
        let mut w = XlsbWriter::create(&file).unwrap();
        w.start_sheet(
            "Benchmark",
            columns,
            Some(&headers),
            XlsbSheetOptions::new(),
        )
        .unwrap();
        for row in &data {
            w.write_row(row).unwrap();
        }
        w.end_sheet().unwrap();
        w.finalize().unwrap();
        ws_xlsb.push(t0.elapsed().as_secs_f64() * 1000.0);
    }
    let mut ws_xlsx = Vec::with_capacity(iters);
    for _ in 0..iters {
        let file = out.join("bench-rs-stream.xlsx");
        let _ = std::fs::remove_file(&file);
        let t0 = Instant::now();
        let mut w = XlsxWriter::create(&file).unwrap();
        w.start_sheet("Benchmark", columns, Some(&headers), SheetOptions::new())
            .unwrap();
        for row in &data {
            w.write_row(row).unwrap();
        }
        w.end_sheet().unwrap();
        w.finalize().unwrap();
        ws_xlsx.push(t0.elapsed().as_secs_f64() * 1000.0);
    }

    let mut r_xlsb = Vec::with_capacity(iters);
    for _ in 0..iters {
        let t0 = Instant::now();
        let mut r = XlsbReader::new();
        r.open(&out.join("bench-rs.xlsb"), true).unwrap();
        let mut acc = 0.0;
        while r.read().unwrap() {
            for v in r.current_row() {
                checksum_f64(v, &mut acc);
            }
        }
        r_xlsb.push(t0.elapsed().as_secs_f64() * 1000.0);
        std::hint::black_box(acc);
    }
    let mut r_xlsx = Vec::with_capacity(iters);
    for _ in 0..iters {
        let t0 = Instant::now();
        let mut r = XlsxReader::new();
        r.open(&out.join("bench-rs.xlsx"), true).unwrap();
        let mut acc = 0.0;
        while r.read().unwrap() {
            for v in r.current_row() {
                checksum_f64(v, &mut acc);
            }
        }
        r_xlsx.push(t0.elapsed().as_secs_f64() * 1000.0);
        std::hint::black_box(acc);
    }

    let size_xlsb = std::fs::metadata(out.join("bench-rs.xlsb")).unwrap().len();
    let size_xlsx = std::fs::metadata(out.join("bench-rs.xlsx")).unwrap().len();
    let size_stream_xlsb = std::fs::metadata(out.join("bench-rs-stream.xlsb"))
        .unwrap()
        .len();
    let size_stream_xlsx = std::fs::metadata(out.join("bench-rs-stream.xlsx"))
        .unwrap()
        .len();
    let _ = XlsbSheetOptions::new();
    let m_w_xlsb = median(w_xlsb.clone());
    let m_w_xlsx = median(w_xlsx.clone());
    let m_ws_xlsb = median(ws_xlsb.clone());
    let m_ws_xlsx = median(ws_xlsx.clone());
    let m_r_xlsb = median(r_xlsb.clone());
    let m_r_xlsx = median(r_xlsx.clone());
    println!(
        "{{\"impl\":\"rust\",\
        \"rows\":{rows},\"columns\":{columns},\"iters\":{iters},\
        \"write_xlsb_ms\":{w_xlsb:?},\"write_xlsx_ms\":{w_xlsx:?},\
        \"write_stream_xlsb_ms\":{ws_xlsb:?},\"write_stream_xlsx_ms\":{ws_xlsx:?},\
        \"read_xlsb_ms\":{r_xlsb:?},\"read_xlsx_ms\":{r_xlsx:?},\
        \"write_xlsb_median_ms\":{m_w_xlsb},\"write_xlsx_median_ms\":{m_w_xlsx},\
        \"write_stream_xlsb_median_ms\":{m_ws_xlsb},\"write_stream_xlsx_median_ms\":{m_ws_xlsx},\
        \"read_xlsb_median_ms\":{m_r_xlsb},\"read_xlsx_median_ms\":{m_r_xlsx},\
        \"size_xlsb\":{size_xlsb},\"size_xlsx\":{size_xlsx},\
        \"size_stream_xlsb\":{size_stream_xlsb},\"size_stream_xlsx\":{size_stream_xlsx}}}",
    );
}
