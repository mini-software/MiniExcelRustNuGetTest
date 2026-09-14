# MiniExcelRust benchmark (linux-x64)

- Date (UTC): 2026-09-14T08:55:01.3909402Z
- OS: Ubuntu 24.04.5 LTS
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
| Cold | MiniExcelV2 | 2801.97 | 1044.64 | 1551.94 | 75.81 | 132.53 | 1x |
| Cold | MiniExcelRust | 1248.99 | 415.53 | 114.85 | 60.33 | 125.86 | 2.24x |
| Warm | MiniExcelV2 | 5784.07 | 547 | 4674.55 | 79.27 | 132.62 | 1x |
| Warm | MiniExcelRust | 3225.5 | 400.47 | 304.99 | 64.52 | 130.06 | 1.79x |

Managed allocation excludes allocations made inside Rust. Peak working set and peak private bytes include the complete process.
