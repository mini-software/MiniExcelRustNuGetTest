# MiniExcelRust benchmark (win-x64)

- Date (UTC): 2026-10-19T08:53:23.9239752Z
- OS: Microsoft Windows 10.0.26100
- Architecture: X64
- .NET SDK: 10.0.402
- Target framework: net10.0
- .NET runtime: .NET 10.0.13
- Baseline: MiniExcel 2.0.0-preview.4
- Candidate: MiniExcelRust 0.1.0-preview.2
- Workbook: 100000 rows x 10 columns
- Iterations: 5 fresh processes per runtime and scenario

| Scenario | Runtime | Elapsed (ms) | First row (ms) | Managed allocation (MB) | Peak working set (MB) | Peak private (MB) | Speedup |
| --- | --- | ---: | ---: | ---: | ---: | ---: | ---: |
| Cold | MiniExcelV2 | 2808.55 | 1030.31 | 1552.21 | 56.42 | 28.83 | 1x |
| Cold | MiniExcelRust | 1472.99 | 451.42 | 110.79 | 45.36 | 24.99 | 1.91x |
| Warm | MiniExcelV2 | 5188.33 | 507.8 | 4674.55 | 60.01 | 33.58 | 1x |
| Warm | MiniExcelRust | 4099.78 | 442.7 | 304.99 | 49.21 | 28.57 | 1.27x |

Managed allocation excludes allocations made inside Rust. Peak working set and peak private bytes include the complete process.
