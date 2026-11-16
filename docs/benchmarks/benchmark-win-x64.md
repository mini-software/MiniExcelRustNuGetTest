# MiniExcelRust benchmark (win-x64)

- Date (UTC): 2026-11-16T11:14:23.4236502Z
- OS: Microsoft Windows 10.0.26100
- Architecture: X64
- .NET SDK: 10.0.404
- Target framework: net10.0
- .NET runtime: .NET 10.0.15
- Baseline: MiniExcel 2.0.0-preview.4
- Candidate: MiniExcelRust 0.1.0-preview.2
- Workbook: 100000 rows x 10 columns
- Iterations: 5 fresh processes per runtime and scenario

| Scenario | Runtime | Elapsed (ms) | First row (ms) | Managed allocation (MB) | Peak working set (MB) | Peak private (MB) | Speedup |
| --- | --- | ---: | ---: | ---: | ---: | ---: | ---: |
| Cold | MiniExcelV2 | 2762.31 | 1058.62 | 1551.1 | 56.49 | 28.95 | 1x |
| Cold | MiniExcelRust | 1572.84 | 496.51 | 110.46 | 46.06 | 25.84 | 1.76x |
| Warm | MiniExcelV2 | 5387.02 | 526.7 | 4674.54 | 60.32 | 33.26 | 1x |
| Warm | MiniExcelRust | 4312.56 | 477.37 | 304.99 | 49.03 | 28.53 | 1.25x |

Managed allocation excludes allocations made inside Rust. Peak working set and peak private bytes include the complete process.
