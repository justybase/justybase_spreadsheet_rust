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
