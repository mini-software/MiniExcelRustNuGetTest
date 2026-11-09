# MiniExcelRust benchmark (win-x64)

- Date (UTC): 2026-11-09T11:29:14.9257393Z
- OS: Microsoft Windows 10.0.26100
- Architecture: X64
- .NET SDK: 10.0.403
- Target framework: net10.0
- .NET runtime: .NET 10.0.14
- Baseline: MiniExcel 2.0.0-preview.4
- Candidate: MiniExcelRust 0.1.0-preview.2
- Workbook: 100000 rows x 10 columns
- Iterations: 5 fresh processes per runtime and scenario

| Scenario | Runtime | Elapsed (ms) | First row (ms) | Managed allocation (MB) | Peak working set (MB) | Peak private (MB) | Speedup |
| --- | --- | ---: | ---: | ---: | ---: | ---: | ---: |
| Cold | MiniExcelV2 | 2690.51 | 991 | 1546.96 | 56.16 | 28.81 | 1x |
| Cold | MiniExcelRust | 1492.2 | 457.28 | 110.98 | 45.48 | 25.01 | 1.8x |
| Warm | MiniExcelV2 | 5194.71 | 500.68 | 4674.55 | 60.25 | 33.31 | 1x |
| Warm | MiniExcelRust | 4189.04 | 447.09 | 304.99 | 49.21 | 28.64 | 1.24x |

Managed allocation excludes allocations made inside Rust. Peak working set and peak private bytes include the complete process.
