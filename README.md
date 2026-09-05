# MiniExcelRust

`MiniExcelRust` provides low-memory, high-performance XLSX queries for .NET by calling
[MiniExcel for Rust](https://github.com/mini-software/MiniExcel-Rust) through a small,
versioned C ABI.

> This repository and package are experimental. The current API supports Rust-backed dynamic
> XLSX and CSV reads, including path and stream inputs, bounded ranges, named tables, workbook
> metadata, and managed `DataTable`/`IDataReader` adapters.

## Install

```shell
dotnet add package MiniExcelRust --prerelease
```

The package is currently prerelease. Pin an exact version in production builds:

```shell
dotnet add package MiniExcelRust --version 0.1.0-preview.2
```

To check for and install a newer preview:

```shell
dotnet list package --outdated --include-prerelease
dotnet add package MiniExcelRust --prerelease
```

## Usage

```csharp
using MiniExcelLibs;

foreach (var row in MiniExcelRust.Query("input.xlsx", useHeaderRow: true))
{
    Console.WriteLine(row["Name"]);
}

var typedRows = MiniExcelRust.Query<MyRow>("input.xlsx");
```

`Query` accepts `path`, `useHeaderRow`, `sheetName`, and `startCell`. Each streamed row is
an `IDictionary<string, object?>`; cells are returned as strings, doubles, booleans, or nulls.
For example, a query against another sheet starting at `C2` is:

```csharp
var rows = MiniExcelRust.Query(
    "input.xlsx",
    useHeaderRow: true,
    sheetName: "Data",
    startCell: "C2");
```

Rows are streamed in bounded batches across the native boundary. Disposing the enumerator
early closes the native query handle. Normal `foreach` enumeration disposes it automatically;
code that manually obtains an enumerator should wrap it in `using`.

Additional read APIs include:

```csharp
var names = MiniExcelRust.GetSheetNames("input.xlsx");
var dimensions = MiniExcelRust.GetSheetDimensions("input.xlsx");
var comments = MiniExcelRust.RetrieveComments("input.xlsx", "Data");
var tableRows = MiniExcelRust.QueryTable("input.xlsx", "Data", "Table1");
var rangeRows = MiniExcelRust.QueryRange("input.xlsx", true, "Data", "C2", "F100");
var dataTable = MiniExcelRust.QueryAsDataTable("input.xlsx", hasHeaderRow: true);

var written = MiniExcelRust.SaveAs(
    "output.xlsx",
    new[]
    {
        new Dictionary<string, object?> { ["Name"] = "alpha", ["Value"] = 42d }
    });

var csvRows = MiniExcelRust.QueryCsv(
    "input.csv",
    useHeaderRow: true,
    new MiniExcelRustCsvReadOptions { Delimiter = ';' });

MiniExcelRust.SaveAsCsv(
    "output.csv",
    new[] { new Dictionary<string, object?> { ["Name"] = "alpha" } });

MiniExcelRust.FillTemplate(
    "report.xlsx",
    "template.xlsx",
    new { title = "Quarterly report", items = new[] { new { name = "Ada" } } });
```

Stream overloads stage input to a temporary file so the Rust engine can retain its bounded-memory
path iterator. They honor `leaveOpen` and remove the temporary file on completion, failure, or
early enumeration disposal. Native stream callbacks are planned to remove this staging step.

See [the live parity matrix](docs/parity-matrix.md) for verified APIs and known gaps.

## Supported Platforms

| .NET RID | Operating system | Architecture | C library |
| --- | --- | --- | --- |
| `win-x64` | Windows | x64 | MSVC |
| `win-arm64` | Windows | arm64 | MSVC |
| `linux-x64` | Linux | x64 | glibc |
| `linux-arm64` | Linux | arm64 | glibc |
| `linux-musl-x64` | Linux | x64 | musl |
| `linux-musl-arm64` | Linux | arm64 | musl |
| `osx-x64` | macOS | x64 | system |
| `osx-arm64` | macOS | arm64 | system |

The NuGet package follows the standard .NET runtime asset convention and places each native
library under `runtimes/{rid}/native/`. .NET selects the matching asset at publish or run time.

## Build and Test

Rust 1.85 and .NET SDK 8 or later are required.

Rust 1.85 is the source and release MSRV. The musl targets default to a static C runtime,
which cannot produce a `cdylib`, so their build disables `crt-static` and links dynamically
against musl before testing the package in Alpine.

```powershell
cargo test --workspace --all-targets --locked
dotnet build ./src/MiniExcelRust/MiniExcelRust.csproj -c Release
./build/Test-Package.ps1 -Rid win-x64
```

Use the local MiniExcel checkout as the read-only behavior oracle instead of the published package:

```powershell
./build/Test-Package.ps1 -Rid win-x64 -MiniExcelSourceRoot D:\git\MiniExcel
```

The comments contract can also be checked against the shared Rust fixture after packing:

```powershell
dotnet run --project .\tests\MiniExcelRust.PackageTests -c Release -- comments `
    D:\git\MiniExcel-Rust\tests\data\xlsx\TestCommentsAndNotes.xlsx sheet1
```

`Test-Package.ps1` builds the native library, packs `MiniExcelRust`, restores a separate
consumer from the local package feed, and verifies equivalent queries against MiniExcel.

GitHub CI runs header, headerless, sheet, range, named-table, metadata, stream, CSV, Unicode,
boolean, null, numeric, full-enumeration, and early-disposal queries on all eight supported RIDs.
Each platform also
runs 5,000 lifecycle iterations and fails when private memory grows by more than 32 MB, when
the native handle/file-descriptor count grows by more than four, or when the workbook cannot
be reopened exclusively. This is a bounded resource-growth regression test rather than a
mathematical proof that no leak can exist.

## Benchmark

The benchmark first compares every returned row and cell with the configured MiniExcel NuGet
baseline. It then measures both implementations in alternating fresh processes against the
same generated XLSX file and query options. The scheduled and manually dispatched GitHub
workflow runs on Windows, Linux, and macOS for x64 and Arm64; musl remains covered by the
Alpine correctness and lifecycle job because GitHub does not provide native musl runners.

### Latest Results

<!-- benchmark-summary:start -->
_Last updated (UTC): 2026-09-05 14:05:30_

| RID | Scenario | MiniExcel (ms) | MiniExcelRust (ms) | Speedup | Allocation reduction | Working-set reduction |
| --- | --- | ---: | ---: | ---: | ---: | ---: |
| win-x64 | Cold | 2891.79 | 1041.49 | 2.78x | 92.5% | 20.3% |
| win-x64 | Warm | 6136.27 | 2770.39 | 2.21x | 93.5% | 19.3% |

[Full reports and raw results](https://github.com/mini-software/MiniExcelRustNuGetTest/blob/main/docs/benchmarks/README.md)
<!-- benchmark-summary:end -->

Each full report includes elapsed time, first-row latency, managed allocation, peak process
memory, environment metadata, and a JSON file containing all raw iterations and hashes.

After all scheduled benchmarks pass on the default branch, the workflow updates this summary
and `docs/benchmarks/` through an `automation/benchmark-results` pull request. Repeated runs
refresh the same PR instead of committing directly to the protected branch. Repository settings
must allow GitHub Actions to create pull requests.

Run the same reproducible comparison locally, or override `-MiniExcelVersion` to test a newer
NuGet release:

```powershell
./build/Benchmark-Package.ps1 -Rid win-x64 -MiniExcelVersion 2.0.0-preview.4
```

## Release

Version tags use the form `v0.1.0-preview.2`. The release workflow builds all eight native
assets, verifies the assembled package, tests it on native GitHub-hosted runners, and publishes
to NuGet.org through the protected `release` environment.

## License

Apache-2.0
