//! Writes parity fixtures with the Rust implementation.
//!
//! Usage: `cargo run --release -p justybase-spreadsheet --example parity_write <outDir>`
//!
//! The data mirrors `parity/gen_fixtures.ts` (TypeScript reference) exactly:
//! same headers, rows, formats and sheet layout, so the two outputs can be
//! compared member-by-member.

use chrono::NaiveDate;
use spreadsheet::{CellValue, XlsbWriter, XlsxWriter, F};
use std::path::PathBuf;

fn dt(y: i32, mo: u32, d: u32, h: u32, mi: u32, s: u32) -> chrono::NaiveDateTime {
    NaiveDate::from_ymd_opt(y, mo, d)
        .unwrap()
        .and_hms_opt(h, mi, s)
        .unwrap()
}

fn headers() -> Vec<String> {
    [
        "ID", "Name", "Count", "Score", "Active", "Joined", "Amount", "Big", "Missing",
    ]
    .iter()
    .map(|s| s.to_string())
    .collect()
}

fn rows() -> Vec<Vec<CellValue>> {
    vec![
        vec![
            CellValue::Integer(1),
            CellValue::Text("Alicja żółć".into()),
            CellValue::Integer(42),
            CellValue::Number(12.5),
            CellValue::Boolean(true),
            CellValue::DateTime(dt(2024, 1, 15, 12, 30, 45)),
            CellValue::formatted(CellValue::Number(1234.5), F::TWO_DECIMALS),
            CellValue::Text("9007199254740993".into()),
            CellValue::Empty,
        ],
        vec![
            CellValue::Integer(-7),
            CellValue::Text("Bob & <Co> \"q\" 'x'".into()),
            CellValue::Integer(1 << 29),
            CellValue::Number(-0.001),
            CellValue::Boolean(false),
            CellValue::formatted(CellValue::DateTime(dt(2023, 12, 31, 8, 0, 0)), F::DATE_ISO),
            CellValue::Text("  spaced  ".into()),
            CellValue::Text("0".into()),
            CellValue::Empty,
        ],
        vec![
            CellValue::Integer(0),
            CellValue::Text(String::new()),
            CellValue::Integer(-(1 << 29)),
            CellValue::Number(f64::NAN),
            CellValue::Boolean(true),
            CellValue::DateTime(dt(1899, 12, 30, 0, 0, 0)),
            CellValue::formatted(CellValue::Integer(5), F::CURRENCY_PLN),
            CellValue::Text("-123456789012345678901234567890".into()),
            CellValue::Empty,
        ],
    ]
}

fn main() {
    let out = PathBuf::from(std::env::args().nth(1).expect("outDir argument"));
    std::fs::create_dir_all(&out).unwrap();

    // --- basic.xlsb ---
    let mut w = XlsbWriter::create(&out.join("rs-basic.xlsb")).unwrap();
    w.add_sheet("Sheet1", false);
    w.write_sheet(rows(), Some(&headers()), true).unwrap();
    w.finalize().unwrap();

    // --- basic.xlsx ---
    let mut w = XlsxWriter::create(&out.join("rs-basic.xlsx")).unwrap();
    w.add_sheet("Sheet1", false);
    w.write_sheet(rows(), Some(&headers()), true).unwrap();
    w.finalize().unwrap();

    // --- multi ---
    for (ext, is_xlsb) in [("xlsb", true), ("xlsx", false)] {
        let kv = vec!["K".to_string(), "V".to_string()];
        let xy = vec!["X".to_string(), "Y".to_string()];
        let h = vec!["H1".to_string(), "H2".to_string()];
        if is_xlsb {
            let mut w = XlsbWriter::create(&out.join(format!("rs-multi.{ext}"))).unwrap();
            w.add_sheet("data1", false);
            w.write_sheet(
                vec![
                    vec![CellValue::Text("a".into()), CellValue::Integer(1)],
                    vec![CellValue::Text("b".into()), CellValue::Integer(2)],
                ],
                Some(&kv),
                true,
            )
            .unwrap();
            w.add_sheet("data2", true);
            w.write_sheet(
                vec![vec![
                    CellValue::Text("only".into()),
                    CellValue::Boolean(false),
                ]],
                Some(&xy),
                true,
            )
            .unwrap();
            w.add_sheet("empty", false);
            w.write_sheet(vec![], Some(&h), false).unwrap();
            w.finalize().unwrap();
        } else {
            let mut w = XlsxWriter::create(&out.join(format!("rs-multi.{ext}"))).unwrap();
            w.add_sheet("data1", false);
            w.write_sheet(
                vec![
                    vec![CellValue::Text("a".into()), CellValue::Integer(1)],
                    vec![CellValue::Text("b".into()), CellValue::Integer(2)],
                ],
                Some(&kv),
                true,
            )
            .unwrap();
            w.add_sheet("data2", true);
            w.write_sheet(
                vec![vec![
                    CellValue::Text("only".into()),
                    CellValue::Boolean(false),
                ]],
                Some(&xy),
                true,
            )
            .unwrap();
            w.add_sheet("empty", false);
            w.write_sheet(vec![], Some(&h), false).unwrap();
            w.finalize().unwrap();
        }
    }
    println!("wrote Rust fixtures to {}", out.display());
}
