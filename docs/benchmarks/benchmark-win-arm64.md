# MiniExcelRust benchmark (win-arm64)

- Date (UTC): 2026-10-19T08:52:21.5026642Z
- OS: Microsoft Windows 10.0.26200
- Architecture: Arm64
- .NET SDK: 10.0.402
- Target framework: net10.0
- .NET runtime: .NET 10.0.13
- Baseline: MiniExcel 2.0.0-preview.4
- Candidate: MiniExcelRust 0.1.0-preview.2
- Workbook: 100000 rows x 10 columns
- Iterations: 5 fresh processes per runtime and scenario

| Scenario | Runtime | Elapsed (ms) | First row (ms) | Managed allocation (MB) | Peak working set (MB) | Peak private (MB) | Speedup |
| --- | --- | ---: | ---: | ---: | ---: | ---: | ---: |
| Cold | MiniExcelV2 | 3008.8 | 948.67 | 1551.83 | 46.17 | 17.12 | 1x |
| Cold | MiniExcelRust | 1081.36 | 345.37 | 111.4 | 34.32 | 12.92 | 2.78x |
| Warm | MiniExcelV2 | 6636.63 | 548.19 | 4677.76 | 49.65 | 20.57 | 1x |
| Warm | MiniExcelRust | 2989.89 | 339.31 | 304.99 | 37.07 | 16.12 | 2.22x |

Managed allocation excludes allocations made inside Rust. Peak working set and peak private bytes include the complete process.
