# MiniExcelRust benchmark (osx-x64)

- Date (UTC): 2026-09-14T08:58:18.2733120Z
- OS: macOS 15.7.9
- Architecture: X64
- .NET SDK: 10.0.401
- Target framework: net10.0
- .NET runtime: .NET 10.0.12
- Baseline: MiniExcel 2.0.0-preview.4
- Candidate: MiniExcelRust 0.1.0-preview.2
- Workbook: 100000 rows x 10 columns
- Iterations: 5 fresh processes per runtime and scenario

| Scenario | Runtime | Elapsed (ms) | First row (ms) | Managed allocation (MB) | Peak working set (MB) | Peak private (MB) | Speedup |
| --- | --- | ---: | ---: | ---: | ---: | ---: | ---: |
| Cold | MiniExcelV2 | 5246.29 | 1586.41 | 1556.35 | 42.9 | 0 | 1x |
| Cold | MiniExcelRust | 2316.14 | 700.44 | 110.8 | 32.2 | 0 | 2.27x |
| Warm | MiniExcelV2 | 12203.88 | 850.79 | 4675.23 | 43.38 | 0 | 1x |
| Warm | MiniExcelRust | 6352.62 | 676.38 | 304.99 | 31.34 | 0 | 1.92x |

Managed allocation excludes allocations made inside Rust. Peak working set and peak private bytes include the complete process.
