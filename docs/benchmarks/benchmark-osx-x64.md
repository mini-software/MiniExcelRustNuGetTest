# MiniExcelRust benchmark (osx-x64)

- Date (UTC): 2026-10-26T08:57:51.4785880Z
- OS: macOS 15.7.9
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
| Cold | MiniExcelV2 | 6255.98 | 1924.61 | 1553.31 | 42.69 | 0 | 1x |
| Cold | MiniExcelRust | 2583.26 | 789.64 | 111.94 | 32.43 | 0 | 2.42x |
| Warm | MiniExcelV2 | 15069.95 | 1090.55 | 4675.23 | 43.47 | 0 | 1x |
| Warm | MiniExcelRust | 7729.16 | 943.4 | 304.99 | 31.5 | 0 | 1.95x |

Managed allocation excludes allocations made inside Rust. Peak working set and peak private bytes include the complete process.
