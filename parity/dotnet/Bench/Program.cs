// Benchmarks for the original .NET implementation (SpreadSheetTasks 1.0.2 on .NET 10).
// Same deterministic dataset and typed flow as parity/bench_ts.ts and Rust parity_bench.
// Prints one JSON document to stdout. Usage: dotnet run -c Release -- <outDir> [rows] [iters]
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using SpreadSheetTasks;

string outDir = args.Length > 0 ? args[0] : "/tmp/bench";
int rowCount = args.Length > 1 ? int.Parse(args[1], CultureInfo.InvariantCulture) : 5000;
int iters = args.Length > 2 ? int.Parse(args[2], CultureInfo.InvariantCulture) : 3;
Directory.CreateDirectory(outDir);

string[] headers = ["ID", "Name", "Count", "Score", "Date", "Active", "Description"];
const string Tail = " z polskimi znakami: ąęśćńźółĄĘŚĆŃŹÓŁ oraz dłuższy tekst testowy.";
var baseDate = new DateTime(2024, 1, 15, 12, 0, 0, DateTimeKind.Utc);

DataTable BuildTable()
{
    var t = new DataTable();
    t.Columns.Add("ID", typeof(int));
    t.Columns.Add("Name", typeof(string));
    t.Columns.Add("Count", typeof(int));
    t.Columns.Add("Score", typeof(double));
    t.Columns.Add("Date", typeof(DateTime));
    t.Columns.Add("Active", typeof(bool));
    t.Columns.Add("Description", typeof(string));
    for (int i = 0; i < rowCount; i++)
        t.Rows.Add(
            i,
            $"Produkt {i} żółć",
            (int)((i * 7919L) % 10000),
            ((i * 104729L) % 100000) / 1000.0,
            baseDate.AddSeconds(i),
            i % 2 == 0,
            $"Opis produktu {i}{Tail}"
        );
    return t;
}

static double Median(List<double> xs) { xs.Sort(); return xs[xs.Count / 2]; }

var table = BuildTable();
string xlsb = Path.Combine(outDir, "bench-cs.xlsb");
string xlsx = Path.Combine(outDir, "bench-cs.xlsx");

// Unmeasured warmup (JIT) so medians reflect steady state.
{
    string w1 = Path.Combine(outDir, "warm.xlsb");
    string w2 = Path.Combine(outDir, "warm.xlsx");
    using (var w = new XlsbWriter(w1, CompressionLevel.Fastest)) { w.AddSheet("Benchmark", false); w.WriteSheet(table, true, 0, 1, 1, true); }
    using (var w = new XlsxWriter(w2, 1000, false, true, CompressionLevel.Fastest)) { w.AddSheet("Benchmark", false); w.WriteSheet(table, true, 0, 1, 1, true); }
    foreach (var f in new[] { w1, w2 })
    {
        using var r = new XlsxOrXlsbReadOrEdit();
        r.Open(f, true, false, System.Text.Encoding.UTF8);
        r.ActualSheetName = "Benchmark";
        while (r.Read()) { }
    }
    File.Delete(w1); File.Delete(w2);
}

static double Accumulate(object? v, double acc) => v switch
{
    double d => acc + d, float f => acc + f, int i => acc + i, long l => acc + l,
    decimal m => acc + (double)m, string s => acc + s.Length, bool b => acc + (b ? 1 : 0),
    DateTime dt => acc + dt.Ticks % 1000, null => acc, _ => acc + 1
};

var wXlsb = new List<double>();
for (int k = 0; k < iters; k++)
{
    if (File.Exists(xlsb)) File.Delete(xlsb);
    var sw = Stopwatch.StartNew();
    using (var w = new XlsbWriter(xlsb, CompressionLevel.Fastest)) { w.AddSheet("Benchmark", false); w.WriteSheet(table, true, 0, 1, 1, true); }
    sw.Stop(); wXlsb.Add(sw.Elapsed.TotalMilliseconds);
}

var wXlsx = new List<double>();
for (int k = 0; k < iters; k++)
{
    if (File.Exists(xlsx)) File.Delete(xlsx);
    var sw = Stopwatch.StartNew();
    using (var w = new XlsxWriter(xlsx, 1000, false, true, CompressionLevel.Fastest)) { w.AddSheet("Benchmark", false); w.WriteSheet(table, true, 0, 1, 1, true); }
    sw.Stop(); wXlsx.Add(sw.Elapsed.TotalMilliseconds);
}

var rXlsb = new List<double>();
for (int k = 0; k < iters; k++)
{
    var sw = Stopwatch.StartNew();
    using var r = new XlsxOrXlsbReadOrEdit();
    r.Open(xlsb, true, false, System.Text.Encoding.UTF8);
    r.ActualSheetName = "Benchmark";
    double acc = 0; while (r.Read()) for (int i = 0; i < r.FieldCount; i++) acc = Accumulate(r.GetValue(i), acc);
    sw.Stop(); rXlsb.Add(sw.Elapsed.TotalMilliseconds); GC.KeepAlive(acc);
}

var rXlsx = new List<double>();
for (int k = 0; k < iters; k++)
{
    var sw = Stopwatch.StartNew();
    using var r = new XlsxOrXlsbReadOrEdit();
    r.Open(xlsx, true, false, System.Text.Encoding.UTF8);
    r.ActualSheetName = "Benchmark";
    double acc = 0; while (r.Read()) for (int i = 0; i < r.FieldCount; i++) acc = Accumulate(r.GetValue(i), acc);
    sw.Stop(); rXlsx.Add(sw.Elapsed.TotalMilliseconds); GC.KeepAlive(acc);
}

string F(List<double> xs) => "[" + string.Join(",", xs.ConvertAll(x => x.ToString("G17", CultureInfo.InvariantCulture))) + "]";
Console.WriteLine(
    "{\"impl\":\"csharp\"," +
    $"\"rows\":{rowCount},\"iters\":{iters}," +
    $"\"write_xlsb_ms\":{F(wXlsb)},\"write_xlsx_ms\":{F(wXlsx)}," +
    $"\"read_xlsb_ms\":{F(rXlsb)},\"read_xlsx_ms\":{F(rXlsx)}," +
    $"\"write_xlsb_median_ms\":{Median(wXlsb).ToString("G17", CultureInfo.InvariantCulture)}," +
    $"\"write_xlsx_median_ms\":{Median(wXlsx).ToString("G17", CultureInfo.InvariantCulture)}," +
    $"\"read_xlsb_median_ms\":{Median(rXlsb).ToString("G17", CultureInfo.InvariantCulture)}," +
    $"\"read_xlsx_median_ms\":{Median(rXlsx).ToString("G17", CultureInfo.InvariantCulture)}," +
    $"\"size_xlsb\":{new FileInfo(xlsb).Length},\"size_xlsx\":{new FileInfo(xlsx).Length}" +
    "}");
