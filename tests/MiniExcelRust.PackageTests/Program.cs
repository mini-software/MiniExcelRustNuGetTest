using System.Diagnostics;
using System.Globalization;
using System.IO.Compression;
using System.Text;
using System.Text.Json;
using MiniExcelLib;
using MiniExcelLib.OpenXml;
using MiniExcelLibs;
using ManagedMiniExcel = MiniExcelLib.MiniExcel;

if (args.Length == 0)
  return RunSuite(lifecycleIterations: 1_000, maxPrivateGrowthMb: 32);

return args[0].ToLowerInvariant() switch
{
  "suite" => RunSuite(
    args.Length >= 2 ? int.Parse(args[1], CultureInfo.InvariantCulture) : 1_000,
    args.Length >= 3 ? int.Parse(args[2], CultureInfo.InvariantCulture) : 32),
  "verify" => VerifyFileParity(args),
  "generate" => GenerateBenchmarkWorkbook(args),
  "managed" => Benchmark(args, useRust: false),
  "rust" => Benchmark(args, useRust: true),
  _ => Usage()
};

static int RunSuite(int lifecycleIterations, int maxPrivateGrowthMb)
{
  var workbookPath = Path.Combine(Path.GetTempPath(), $"miniexcel-rust-{Guid.NewGuid():N}.xlsx");
  try
  {
    CreateWorkbook(workbookPath);
    VerifyParity(workbookPath);
    VerifyLifecycle(workbookPath, lifecycleIterations, maxPrivateGrowthMb);
    Console.WriteLine("MiniExcelRust parity and lifecycle suite passed.");
    return 0;
  }
  finally
  {
    if (File.Exists(workbookPath))
      File.Delete(workbookPath);
  }
}

static void VerifyParity(string path)
{
  var scenarios = new[]
  {
    new QueryScenario("header", true, "Sheet1", "A1"),
    new QueryScenario("headerless", false, "Sheet1", "A1"),
    new QueryScenario("sheet-and-start-cell", true, "Data", "C2")
  };

  foreach (var scenario in scenarios)
  {
    var managedRows = QueryManaged(path, scenario.UseHeaderRow, scenario.SheetName, scenario.StartCell).ToList();
    var rustRows = MiniExcelRust.Query(path, scenario.UseHeaderRow, scenario.SheetName, scenario.StartCell).ToList();
    CompareRows(managedRows, rustRows, scenario.Name);
  }

  var rows = MiniExcelRust.Query(path, useHeaderRow: true, sheetName: "Sheet1").ToList();
  Require(rows.Count == 3, $"Expected 3 rows, received {rows.Count}.");
  Require(Equals(rows[0]["Name"], "alpha"), "The first string value did not match.");
  Require(Equals(rows[0]["Value"], 42d), "The first numeric value did not match.");
  Require(Equals(rows[0]["Note"], "Taiwan 台灣"), "The Unicode value did not match.");
  Require(Equals(rows[1]["Name"], "beta"), "The second string value did not match.");
  Require(Equals(rows[1]["Value"], true), "The boolean value did not match.");
  Require(rows[1]["Note"] is null, "The empty value should be null.");
}

static void Require(bool condition, string message)
{
    if (!condition)
        throw new InvalidOperationException(message);
}

static IEnumerable<IDictionary<string, object?>> QueryManaged(
  string path,
  bool useHeaderRow,
  string? sheetName = null,
  string startCell = "A1")
{
  var importer = ManagedMiniExcel.Importers.GetOpenXmlImporter();
  foreach (IDictionary<string, object?> row in importer.Query(
    path,
    hasHeaderRow: useHeaderRow,
    sheetName: sheetName,
    startCell: startCell))
    yield return row;
}

static void CompareRows(
  IReadOnlyList<IDictionary<string, object?>> expected,
  IReadOnlyList<IDictionary<string, object?>> actual,
  string scenario)
{
  Require(expected.Count == actual.Count, $"{scenario}: row count differs: managed={expected.Count}, rust={actual.Count}.");
  for (var rowIndex = 0; rowIndex < expected.Count; rowIndex++)
    CompareRow(expected[rowIndex], actual[rowIndex], scenario, rowIndex);
}

static void CompareRow(
  IDictionary<string, object?> expected,
  IDictionary<string, object?> actual,
  string scenario,
  long rowIndex)
{
  Require(
    expected.Keys.SequenceEqual(actual.Keys, StringComparer.Ordinal),
    $"{scenario}: column order differs at row {rowIndex}.");
  foreach (var key in expected.Keys)
  {
    Require(
      Equals(expected[key], actual[key]),
      $"{scenario}: value differs at row {rowIndex}, column {key}: managed={expected[key] ?? "<null>"}, rust={actual[key] ?? "<null>"}.");
  }
}

static int VerifyFileParity(string[] arguments)
{
  if (arguments.Length is < 2 or > 3)
    return Usage();

  var path = Path.GetFullPath(arguments[1]);
  var useHeaderRow = arguments.Length == 3 && bool.Parse(arguments[2]);
  using var managed = QueryManaged(path, useHeaderRow).GetEnumerator();
  using var rust = MiniExcelRust.Query(path, useHeaderRow).GetEnumerator();
  long rowIndex = 0;

  while (true)
  {
    var hasManaged = managed.MoveNext();
    var hasRust = rust.MoveNext();
    Require(hasManaged == hasRust, $"benchmark-parity: row count differs after row {rowIndex}.");
    if (!hasManaged)
      break;
    CompareRow(managed.Current, rust.Current, "benchmark-parity", rowIndex);
    rowIndex++;
  }

  Console.WriteLine($"Verified {rowIndex} rows against the configured MiniExcel baseline.");
  return 0;
}

static void VerifyLifecycle(string path, int iterations, int maxPrivateGrowthMb)
{
  if (iterations < 100)
    throw new ArgumentOutOfRangeException(nameof(iterations), "Lifecycle iterations must be at least 100.");

  for (var iteration = 0; iteration < 50; iteration++)
    QueryAndDisposeEarly(path);

  ForceCollection();
  var baseline = CaptureResources();
  var samples = new List<ResourceSample> { baseline };
  var sampleInterval = Math.Max(100, iterations / 10);

  for (var iteration = 1; iteration <= iterations; iteration++)
  {
    QueryAndDisposeEarly(path);
    if (iteration % 25 == 0)
      _ = MiniExcelRust.Query(path, useHeaderRow: true).Count();
    if (iteration % sampleInterval == 0)
    {
      ForceCollection();
      samples.Add(CaptureResources());
    }
  }

  using (File.Open(path, FileMode.Open, FileAccess.ReadWrite, FileShare.None)) { }

  var final = samples[^1];
  var privateGrowth = final.PrivateBytes - baseline.PrivateBytes;
  var workingSetGrowth = final.WorkingSetBytes - baseline.WorkingSetBytes;
  var maxPrivateGrowth = maxPrivateGrowthMb * 1024L * 1024L;
  Require(
    privateGrowth <= maxPrivateGrowth,
    $"Private memory grew by {ToMb(privateGrowth):F2} MB after {iterations} early-disposal cycles; limit is {maxPrivateGrowthMb} MB.");

  if (baseline.ResourceCount is not null && final.ResourceCount is not null)
  {
    Require(
      final.ResourceCount - baseline.ResourceCount <= 4,
      $"Native handle/file-descriptor count grew from {baseline.ResourceCount} to {final.ResourceCount}.");
  }

  Console.WriteLine(
    $"Lifecycle: iterations={iterations}, private-growth={ToMb(privateGrowth):F2} MB, " +
    $"working-set-growth={ToMb(workingSetGrowth):F2} MB, resources={baseline.ResourceCount?.ToString() ?? "n/a"}->{final.ResourceCount?.ToString() ?? "n/a"}.");
}

static void QueryAndDisposeEarly(string path)
{
  using var rows = MiniExcelRust.Query(path).GetEnumerator();
  Require(rows.MoveNext(), "The lifecycle fixture did not contain a row.");
}

static ResourceSample CaptureResources()
{
  using var process = Process.GetCurrentProcess();
  process.Refresh();
  return new ResourceSample(process.PrivateMemorySize64, process.WorkingSet64, GetResourceCount(process));
}

static int? GetResourceCount(Process process)
{
  if (OperatingSystem.IsLinux() && Directory.Exists("/proc/self/fd"))
    return Directory.EnumerateFileSystemEntries("/proc/self/fd").Count();

  try
  {
    return process.HandleCount;
  }
  catch (PlatformNotSupportedException)
  {
    return null;
  }
}

static void ForceCollection()
{
  GC.Collect();
  GC.WaitForPendingFinalizers();
  GC.Collect();
}

static double ToMb(long bytes) => bytes / 1024d / 1024d;

static int GenerateBenchmarkWorkbook(string[] arguments)
{
  if (arguments.Length is < 2 or > 4)
    return Usage();

  var path = Path.GetFullPath(arguments[1]);
  var rows = arguments.Length >= 3 ? int.Parse(arguments[2], CultureInfo.InvariantCulture) : 100_000;
  var columns = arguments.Length >= 4 ? int.Parse(arguments[3], CultureInfo.InvariantCulture) : 10;
  CreateBenchmarkWorkbook(path, rows, columns);
  Console.WriteLine($"Generated {rows}x{columns} benchmark workbook at {path}.");
  return 0;
}

static int Benchmark(string[] arguments, bool useRust)
{
  if (arguments.Length is < 2 or > 4)
    return Usage();

  var path = Path.GetFullPath(arguments[1]);
  var passes = arguments.Length >= 3 ? int.Parse(arguments[2], CultureInfo.InvariantCulture) : 1;
  var warmupPasses = arguments.Length >= 4 ? int.Parse(arguments[3], CultureInfo.InvariantCulture) : 0;

  for (var pass = 0; pass < warmupPasses; pass++)
    Consume(path, useRust);

  ForceCollection();
  var allocatedBefore = GC.GetTotalAllocatedBytes(precise: true);
  var stopwatch = Stopwatch.StartNew();
  var firstRowMilliseconds = 0d;
  long rowCount = 0;
  long cellCount = 0;

  for (var pass = 0; pass < passes; pass++)
  {
    var rows = useRust ? MiniExcelRust.Query(path) : QueryManaged(path, useHeaderRow: false);
    foreach (var row in rows)
    {
      if (rowCount == 0)
        firstRowMilliseconds = stopwatch.Elapsed.TotalMilliseconds;
      rowCount++;
      cellCount += row.Count;
    }
  }

  stopwatch.Stop();
  Console.WriteLine(JsonSerializer.Serialize(new BenchmarkResult(
    useRust ? "MiniExcelRust" : "MiniExcelV2",
    System.Runtime.InteropServices.RuntimeInformation.FrameworkDescription,
    passes,
    rowCount,
    cellCount,
    stopwatch.Elapsed.TotalMilliseconds,
    firstRowMilliseconds,
    GC.GetTotalAllocatedBytes(precise: true) - allocatedBefore)));
  return 0;
}

static void Consume(string path, bool useRust)
{
  var rows = useRust ? MiniExcelRust.Query(path) : QueryManaged(path, useHeaderRow: false);
  foreach (var row in rows)
    _ = row.Count;
}

static void CreateWorkbook(string path)
{
    using var archive = ZipFile.Open(path, ZipArchiveMode.Create);
    AddEntry(archive, "[Content_Types].xml", """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
          <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
          <Default Extension="xml" ContentType="application/xml"/>
          <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
          <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
          <Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
        </Types>
        """);
    AddEntry(archive, "_rels/.rels", """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
        </Relationships>
        """);
    AddEntry(archive, "xl/workbook.xml", """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/><sheet name="Data" sheetId="2" r:id="rId2"/></sheets>
        </workbook>
        """);
    AddEntry(archive, "xl/_rels/workbook.xml.rels", """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
          <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>
        </Relationships>
        """);
    AddEntry(archive, "xl/worksheets/sheet1.xml", """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <sheetData>
            <row r="1"><c r="A1" t="inlineStr"><is><t>Name</t></is></c><c r="B1" t="inlineStr"><is><t>Value</t></is></c><c r="C1" t="inlineStr"><is><t>Note</t></is></c></row>
            <row r="2"><c r="A2" t="inlineStr"><is><t>alpha</t></is></c><c r="B2"><v>42</v></c><c r="C2" t="inlineStr"><is><t>Taiwan 台灣</t></is></c></row>
            <row r="3"><c r="A3" t="inlineStr"><is><t>beta</t></is></c><c r="B3" t="b"><v>1</v></c><c r="C3"/></row>
            <row r="4"><c r="A4" t="inlineStr"><is><t>gamma</t></is></c><c r="B4"><v>-7.5</v></c><c r="C4" t="inlineStr"><is><t>last</t></is></c></row>
          </sheetData>
        </worksheet>
        """);
    AddEntry(archive, "xl/worksheets/sheet2.xml", """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <sheetData>
            <row r="2"><c r="C2" t="inlineStr"><is><t>Code</t></is></c><c r="D2" t="inlineStr"><is><t>Amount</t></is></c></row>
            <row r="3"><c r="C3" t="inlineStr"><is><t>x</t></is></c><c r="D3"><v>3.5</v></c></row>
            <row r="4"><c r="C4" t="inlineStr"><is><t>y</t></is></c><c r="D4"><v>9</v></c></row>
          </sheetData>
        </worksheet>
        """);
}

static void CreateBenchmarkWorkbook(string path, int rows, int columns)
{
    if (rows < 1 || columns is < 1 or > 26)
        throw new ArgumentOutOfRangeException(nameof(rows), "Rows must be positive and columns must be between 1 and 26.");

    Directory.CreateDirectory(Path.GetDirectoryName(path)!);
    if (File.Exists(path))
      File.Delete(path);
    using var archive = ZipFile.Open(path, ZipArchiveMode.Create);
    AddEntry(archive, "[Content_Types].xml", """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
          <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
          <Default Extension="xml" ContentType="application/xml"/>
          <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
          <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
        </Types>
        """);
    AddEntry(archive, "_rels/.rels", """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
        </Relationships>
        """);
    AddEntry(archive, "xl/workbook.xml", """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>
        </workbook>
        """);
    AddEntry(archive, "xl/_rels/workbook.xml.rels", """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
        </Relationships>
        """);

    var entry = archive.CreateEntry("xl/worksheets/sheet1.xml", CompressionLevel.Fastest);
    using var writer = new StreamWriter(entry.Open(), new UTF8Encoding(false));
    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><sheetData>");
    for (var row = 1; row <= rows; row++)
    {
        writer.Write($"<row r=\"{row}\">");
        for (var column = 1; column <= columns; column++)
        {
            var reference = $"{(char)('A' + column - 1)}{row}";
            var value = (long)(row - 1) * columns + column;
            writer.Write($"<c r=\"{reference}\"><v>{value}</v></c>");
        }
        writer.Write("</row>");
    }
    writer.Write("</sheetData></worksheet>");
}

static void AddEntry(ZipArchive archive, string name, string contents)
{
    var entry = archive.CreateEntry(name, CompressionLevel.Fastest);
    using var writer = new StreamWriter(entry.Open(), new UTF8Encoding(false));
    writer.Write(contents);
}

  static int Usage()
  {
    Console.Error.WriteLine("Usage:");
    Console.Error.WriteLine("  PublicNuGetSmoke suite [lifecycle-iterations] [max-private-growth-mb]");
    Console.Error.WriteLine("  PublicNuGetSmoke verify <xlsx-path> [use-header-row]");
    Console.Error.WriteLine("  PublicNuGetSmoke generate <xlsx-path> [rows] [columns]");
    Console.Error.WriteLine("  PublicNuGetSmoke <managed|rust> <xlsx-path> [passes] [warmup-passes]");
    return 2;
  }

  internal sealed record QueryScenario(string Name, bool UseHeaderRow, string SheetName, string StartCell);

  internal sealed record ResourceSample(long PrivateBytes, long WorkingSetBytes, int? ResourceCount);

  internal sealed record BenchmarkResult(
    string Runtime,
    string DotNetRuntime,
    int Passes,
    long Rows,
    long Cells,
    double ElapsedMilliseconds,
    double FirstRowMilliseconds,
    long AllocatedBytes);