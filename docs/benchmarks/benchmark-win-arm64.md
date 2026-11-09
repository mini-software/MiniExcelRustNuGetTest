# MiniExcelRust benchmark (win-arm64)

- Date (UTC): 2026-11-09T11:29:23.7014212Z
- OS: Microsoft Windows 10.0.26200
- Architecture: Arm64
- .NET SDK: 10.0.403
- Target framework: net10.0
- .NET runtime: .NET 10.0.14
- Baseline: MiniExcel 2.0.0-preview.4
- Candidate: MiniExcelRust 0.1.0-preview.2
- Workbook: 100000 rows x 10 columns
- Iterations: 5 fresh processes per runtime and scenario

| Scenario | Runtime | Elapsed (ms) | First row (ms) | Managed allocation (MB) | Peak working set (MB) | Peak private (MB) | Speedup |
| --- | --- | ---: | ---: | ---: | ---: | ---: | ---: |
| Cold | MiniExcelV2 | 2898.45 | 925.58 | 1554.45 | 45.75 | 17.22 | 1x |
| Cold | MiniExcelRust | 1096.3 | 356.14 | 111.38 | 34.35 | 12.94 | 2.64x |
| Warm | MiniExcelV2 | 6603.97 | 544.09 | 4677.76 | 48.82 | 20.54 | 1x |
| Warm | MiniExcelRust | 2988.57 | 339.13 | 304.99 | 36.87 | 16.11 | 2.21x |

Managed allocation excludes allocations made inside Rust. Peak working set and peak private bytes include the complete process.
