# MiniExcelRust benchmark (win-arm64)

- Date (UTC): 2026-11-16T11:14:34.8348345Z
- OS: Microsoft Windows 10.0.26200
- Architecture: Arm64
- .NET SDK: 11.0.100
- Target framework: net10.0
- .NET runtime: .NET 10.0.15
- Baseline: MiniExcel 2.0.0-preview.4
- Candidate: MiniExcelRust 0.1.0-preview.2
- Workbook: 100000 rows x 10 columns
- Iterations: 5 fresh processes per runtime and scenario

| Scenario | Runtime | Elapsed (ms) | First row (ms) | Managed allocation (MB) | Peak working set (MB) | Peak private (MB) | Speedup |
| --- | --- | ---: | ---: | ---: | ---: | ---: | ---: |
| Cold | MiniExcelV2 | 2897.79 | 926.37 | 1552.88 | 46.06 | 16.92 | 1x |
| Cold | MiniExcelRust | 1085.57 | 345.15 | 111.4 | 34.66 | 12.85 | 2.67x |
| Warm | MiniExcelV2 | 6670.97 | 544.81 | 4677.76 | 49.32 | 20.57 | 1x |
| Warm | MiniExcelRust | 3022.06 | 339.99 | 304.99 | 37.41 | 16.1 | 2.21x |

Managed allocation excludes allocations made inside Rust. Peak working set and peak private bytes include the complete process.
