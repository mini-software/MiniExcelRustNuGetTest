# MiniExcelRust benchmark (osx-arm64)

- Date (UTC): 2026-11-16T11:13:15.6234470Z
- OS: macOS 26.7.2
- Architecture: Arm64
- .NET SDK: 10.0.404
- Target framework: net10.0
- .NET runtime: .NET 10.0.15
- Baseline: MiniExcel 2.0.0-preview.4
- Candidate: MiniExcelRust 0.1.0-preview.2
- Workbook: 100000 rows x 10 columns
- Iterations: 5 fresh processes per runtime and scenario

| Scenario | Runtime | Elapsed (ms) | First row (ms) | Managed allocation (MB) | Peak working set (MB) | Peak private (MB) | Speedup |
| --- | --- | ---: | ---: | ---: | ---: | ---: | ---: |
| Cold | MiniExcelV2 | 2831.22 | 1069.22 | 1532.25 | 76.9 | 0 | 1x |
| Cold | MiniExcelRust | 1272.47 | 327.16 | 115.68 | 53.52 | 0 | 2.22x |
| Warm | MiniExcelV2 | 5241.17 | 613.4 | 4675.22 | 77.42 | 0 | 1x |
| Warm | MiniExcelRust | 3638.15 | 356.24 | 304.99 | 53.27 | 0 | 1.44x |

Managed allocation excludes allocations made inside Rust. Peak working set and peak private bytes include the complete process.
