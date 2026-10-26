# MiniExcelRust benchmark (win-arm64)

- Date (UTC): 2026-10-26T08:55:14.4392457Z
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
| Cold | MiniExcelV2 | 2953.73 | 929.03 | 1554.61 | 45.88 | 17.19 | 1x |
| Cold | MiniExcelRust | 1092.32 | 346.39 | 111.17 | 34.36 | 12.92 | 2.7x |
| Warm | MiniExcelV2 | 6650.61 | 544.61 | 4677.76 | 49.59 | 20.56 | 1x |
| Warm | MiniExcelRust | 3018.14 | 338.82 | 304.99 | 37.12 | 16.09 | 2.2x |

Managed allocation excludes allocations made inside Rust. Peak working set and peak private bytes include the complete process.
