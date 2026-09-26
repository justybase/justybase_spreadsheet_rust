# Performance benchmarks

The benchmark harnesses are intended for focused local checks. Their default
workload is deliberately small so routine runs do not generate large files or
long, noisy reports:

- 5,000 data rows plus one header row;
- 7 columns;
- 3 measured iterations (write comparisons also perform one warmup run).

The writer and parity harnesses accept explicit row, iteration, and column
counts when a larger comparison is needed. The Criterion reader comparison
uses the fixed lightweight profile. Benchmark results are machine-dependent
and are not stored in this document.

## Available checks

| Check | Command | Scope |
| --- | --- | --- |
| Reader comparison | `cargo bench --bench read_comparison` | Row and borrowed-cell reads against `calamine` for XLSX and XLSB |
| Rust parity | `cargo run --release --example parity_bench /tmp/bench` | Rust batch/streaming writes and full reads |
| XLSX writers | `cargo run --release --example write_comparison` | This crate against `rust_xlsxwriter` |
| XLSB writers | `cargo run --release --example xlsb_write_comparison` | This crate against `rxlsb` and `xlsb-writer` |
| Large reader scale | `cargo build --release --example read_scale_bench` | Large XLSX/XLSB row and cell scans against `calamine`, with checksums |

The large reader harness can generate an input and measure scans at larger
scales without building the full dataset in memory:

```bash
target/release/examples/read_scale_bench generate /tmp/read-scale 1000000
target/release/examples/read_scale_bench scan rust-xlsx-row /tmp/read-scale/rust-1000000.xlsx 5
target/release/examples/read_scale_bench scan calamine-xlsx-cell /tmp/read-scale/rust-1000000.xlsx 5
```

To isolate parser throughput from this crate's eager conversion of Excel date
serials into `chrono` values, generate the equivalent workload with a numeric
value in the date column:

```bash
target/release/examples/read_scale_bench generate /tmp/read-scale 1000000 --no-dates
target/release/examples/read_scale_bench scan rust-xlsx-cell /tmp/read-scale/rust-1000000-no-dates.xlsx 5
target/release/examples/read_scale_bench scan calamine-xlsx-cell /tmp/read-scale/rust-1000000-no-dates.xlsx 5
```

The no-date file has the same row/column shape and payloads, but the date
column is a number without a date style, so both readers return a numeric cell.

Supported scan modes are `rust-xlsx-row`, `rust-xlsx-cell`, `rust-xlsb-row`,
`rust-xlsb-cell`, and the equivalent `calamine-xlsx-range`,
`calamine-xlsx-cell`, `calamine-xlsb-range`, and `calamine-xlsb-cell` modes.
Each timed run opens the workbook and scans the full sheet; the harness first
does one untimed warmup scan, then checks that all measured scans return the
same row count and checksum. For peak RSS, run one scan in a fresh process and
repeat it three times:

```bash
/usr/bin/time -f 'peak_rss_kib=%M' \
  target/release/examples/read_scale_bench scan rust-xlsb-cell \
  /tmp/read-scale/rust-1000000.xlsb --once
```

Use the same generated files for every implementation being compared. A
representative scale set is 50,000, 250,000, and 1,000,000 data rows.

The comparison examples accept `[outDir] [rows] [iters] [columns]`; the XLSB
example also accepts `[compression]`. For example, an opt-in full-size run is:

```bash
cargo run --release --example parity_bench /tmp/bench 50000 5 7
cargo run --release --example write_comparison /tmp/bench 50000 5 7
```

## Cross-runtime parity

The TypeScript and .NET harnesses use the same deterministic seven-column
dataset and default to the same lightweight 5,000-row/3-iteration profile.
They remain available for explicit comparisons with the Rust example:

```bash
# TypeScript reference checkout is a sibling directory.
cd ../justybase_spreadsheet_tasks
npx ts-node --transpile-only --skip-project \
  --compiler-options '{"module":"commonjs","moduleResolution":"node"}' \
  ../justybase_spreadsheet_rust/parity/bench_ts.ts /tmp/bench

# Rust
cd ../justybase_spreadsheet_rust
cargo run --quiet --release --example parity_bench /tmp/bench

# .NET reference
dotnet run -c Release --project parity/dotnet/Bench/Bench.csproj -- /tmp/bench
```

For comparable full-size numbers, pass the same `50000 5` arguments to all
three harnesses. The JSON output contains per-iteration timings, medians, and
file sizes; the Rust writer comparisons additionally print allocation and
read-back validation summaries.

## Interpreting results

These are indicative local measurements, not capacity guarantees. CPU load,
filesystem cache, compiler/runtime versions, compression level, and dependency
versions can materially change timings. Use the default profile for quick
regression checks and opt into larger profiles only when investigating a
specific workload.
