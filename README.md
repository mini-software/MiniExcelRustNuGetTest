# MiniExcelRust

`MiniExcelRust` provides low-memory, high-performance XLSX queries for .NET by calling
[MiniExcel for Rust](https://github.com/mini-software/MiniExcel-Rust) through a small,
versioned C ABI.

> This repository and package are experimental. The initial API supports synchronous,
> path-based dynamic XLSX queries.

## Install

```shell
dotnet add package MiniExcelRust --prerelease
```

## Usage

```csharp
using MiniExcelLibs;

foreach (var row in MiniExcelRust.Query("input.xlsx", useHeaderRow: true))
{
    Console.WriteLine(row["Name"]);
}
```

Rows are streamed in bounded batches across the native boundary. Disposing the enumerator
early closes the native query handle.

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

`Test-Package.ps1` builds the native library, packs `MiniExcelRust`, restores a separate
consumer from the local package feed, and executes an XLSX query smoke test.

## Release

Version tags use the form `v0.1.0-preview.1`. The release workflow builds all eight native
assets, verifies the assembled package, tests it on native GitHub-hosted runners, and publishes
to NuGet.org through the protected `release` environment.

## License

Apache-2.0
