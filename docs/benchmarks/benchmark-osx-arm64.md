# MiniExcelRust benchmark (osx-arm64)

- Date (UTC): 2026-10-26T08:53:33.2891290Z
- OS: macOS 26.6.2
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
| Cold | MiniExcelV2 | 2997.86 | 1076.67 | 1537.53 | 78.36 | 0 | 1x |
| Cold | MiniExcelRust | 1131.58 | 297.89 | 115.9 | 54.67 | 0 | 2.65x |
| Warm | MiniExcelV2 | 4913.64 | 540.24 | 4675.14 | 78.72 | 0 | 1x |
| Warm | MiniExcelRust | 3282.49 | 323.14 | 304.99 | 54.93 | 0 | 1.5x |

Managed allocation excludes allocations made inside Rust. Peak working set and peak private bytes include the complete process.
