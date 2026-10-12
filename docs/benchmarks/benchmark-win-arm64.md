# MiniExcelRust benchmark (win-arm64)

- Date (UTC): 2026-10-12T10:48:33.9983019Z
- OS: Microsoft Windows 10.0.26200
- Architecture: Arm64
- .NET SDK: 10.0.401
- Target framework: net10.0
- .NET runtime: .NET 10.0.12
- Baseline: MiniExcel 2.0.0-preview.4
- Candidate: MiniExcelRust 0.1.0-preview.2
- Workbook: 100000 rows x 10 columns
- Iterations: 5 fresh processes per runtime and scenario

| Scenario | Runtime | Elapsed (ms) | First row (ms) | Managed allocation (MB) | Peak working set (MB) | Peak private (MB) | Speedup |
| --- | --- | ---: | ---: | ---: | ---: | ---: | ---: |
| Cold | MiniExcelV2 | 2951.6 | 946.38 | 1554.54 | 45.9 | 17.48 | 1x |
| Cold | MiniExcelRust | 1093 | 347.07 | 111.23 | 34.63 | 12.92 | 2.7x |
| Warm | MiniExcelV2 | 6647.06 | 548.12 | 4677.76 | 48.72 | 20.51 | 1x |
| Warm | MiniExcelRust | 3116.98 | 357.93 | 304.99 | 36.55 | 16.14 | 2.13x |

Managed allocation excludes allocations made inside Rust. Peak working set and peak private bytes include the complete process.
