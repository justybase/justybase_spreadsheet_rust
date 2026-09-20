# justybase-spreadsheet

> **⚠️ PREVIEW — NOT PRODUCTION READY**
>
> This project is an early preview. APIs, file-format behavior, performance
> characteristics, and compatibility guarantees may change without notice.
> Do not use it for production data without independent validation and backups.

High-performance Rust readers and writers for Excel XLSX and XLSB workbooks.
The crate is a Rust port of `@justybase/spreadsheet-tasks` and focuses on
streaming-friendly access to workbook data.

## Current status

This repository is actively evolving and should be treated as experimental.
The current implementation includes:

- XLSX and XLSB reading;
- XLSX and XLSB writing;
- forward-only row iteration;
- borrowed forward-only explicit-cell iteration for XLSX and XLSB;
- batch and streaming worksheet writers;
- worksheet updates for XLSX, XLSM, and XLSB;
- date-system handling for Excel's 1900 and 1904 systems;
- parity and cross-runtime regression tests.

The public API is not yet considered stable. Not every Excel feature is
implemented, and compatibility should be tested against the exact workbooks
your application uses.

## Quick start

Add the crate from a local checkout while the project is in preview:

```toml
[dependencies]
justybase-spreadsheet = { path = "../justybase_spreadsheet_rust" }
```

> Note: the library target keeps the short name, so imports stay
> `use spreadsheet::...` even though the package is `justybase-spreadsheet`.

Create an XLSX workbook:

```rust,no_run
use spreadsheet::{CellValue, XlsxWriter};
use std::path::Path;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let path = Path::new("output.xlsx");
    let mut writer = XlsxWriter::create(path)?;
    writer.add_sheet("Sheet1", false);
    writer.write_sheet(
        vec![vec![CellValue::Text("Alice".into()), CellValue::Integer(30)]],
        Some(&["Name".to_string(), "Age".to_string()]),
        true,
    )?;
    writer.finalize()?;
    Ok(())
}
```

Read rows from either supported format:

```rust,no_run
use spreadsheet::{create_reader, SpreadsheetResult};
use std::path::Path;

fn read_workbook(path: &Path) -> SpreadsheetResult<usize> {
    let mut reader = create_reader(path)?;
    reader.open(path, true)?;

    let mut cells = 0;
    while reader.read()? {
        cells += reader.current_row().len();
    }
    Ok(cells)
}
```

For convenience, the row-oriented API remains available; avoid collecting the
entire workbook unless your application needs random access.

When the lowest allocation rate matters, use the explicit-cell streaming API.
It yields borrowed `CellRef` values in worksheet order; text is borrowed from
the shared-string table or a reusable reader buffer and is valid until the
next call to `next_cell`:

```rust,no_run
use spreadsheet::{create_reader, SpreadsheetResult};
use std::path::Path;

fn read_cells(path: &Path) -> SpreadsheetResult<usize> {
    let mut reader = create_reader(path)?;
    reader.open(path, true)?;
    let mut cells = reader.cell_reader("Sheet1")?;
    let mut count = 0;
    while let Some(cell) = cells.next_cell()? {
        println!("({}, {}) = {:?}", cell.row, cell.column, cell.value);
        count += 1;
    }
    Ok(count)
}
```

This is a forward-only API for explicit cells. It does not materialize a
worksheet-wide XML buffer or a `Vec<CellValue>` row, and therefore is the
preferred path for large scans. The legacy row API remains available for
compatibility and convenience.

## Development

Run the complete test suite:

```bash
cargo test
```

Run formatting checks:

```bash
cargo fmt --check
```

Run the read-performance comparison with
[`calamine`](https://github.com/tafia/calamine):

```bash
cargo bench --bench read_comparison
```

The benchmark creates one deterministic XLSX and one XLSB fixture. The
`full_read` group measures the compatible row API against calamine; the
`streaming_cell_read` group measures the new borrowed explicit-cell API.
The default fixture contains 5,000 rows and keeps the run short. Criterion
reports are written to `target/criterion/`; the available benchmark profiles
are listed in the repository's
[BENCHMARKS.md](https://github.com/justybase/justybase_spreadsheet_rust/blob/main/BENCHMARKS.md).

Run the XLSX write comparison against
[`rust_xlsxwriter`](https://github.com/jmcnamara/rust_xlsxwriter):

```bash
cargo run --release --example write_comparison
```

It prints a text report (median write time, output file size, peak heap,
total allocations) for the batch and streaming writers of this crate against
the default and `constant_memory` modes of `rust_xlsxwriter`; optional
arguments are `[outDir] [rows] [iters] [columns]`. The default is 5,000 rows
and 3 iterations; larger profiles are opt-in. Available checks are listed in
[the repository benchmark guide](https://github.com/justybase/justybase_spreadsheet_rust/blob/main/BENCHMARKS.md).

### Full write/read correctness process

End-to-end gate combining the suites above with an independent reader:

```bash
cargo fmt --check && cargo test && cargo test --release
cargo bench --bench read_comparison -- --quick   # asserts our reader == calamine
cargo run --release --example parity_bench /tmp/parity_check
cargo run --release --example write_comparison /tmp/parity_check
python3 scripts/validate_openpyxl.py \
  /tmp/parity_check/spreadsheet-batch.xlsx=5001x7 \
  /tmp/parity_check/spreadsheet-streaming.xlsx=5001x7
```

`scripts/validate_openpyxl.py` value-checks the `parity_write` fixtures with
[openpyxl](https://openpyxl.readthedocs.io) — a pure-Python reader with no
code shared with this crate — covering cell types, exact values, XML
escaping, number formats, datetimes, hidden sheets and empty sheets, then
verifies the sheet extent of the larger benchmark outputs against
`<dimension>` (read-only consumers such as pandas size sheets from that
element). It bootstraps a virtualenv under `target/ovenv` when openpyxl is
not installed. openpyxl reads `.xlsx` only; `.xlsb` is covered by the
roundtrip tests and the calamine cross-check.

## License

Licensed under the MIT License. See [LICENSE-MIT](LICENSE-MIT).
