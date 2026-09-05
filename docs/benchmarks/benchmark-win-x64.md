# MiniExcelRust benchmark (win-x64)

- Date (UTC): 2026-09-05T14:05:30.0420829Z
- OS: Microsoft Windows 10.0.19045
- Architecture: X64
- .NET SDK: 10.0.103
- Target framework: net10.0
- .NET runtime: .NET 10.0.3
- Baseline: MiniExcel 2.0.0-preview.4
- Candidate: MiniExcelRust 0.1.0-preview.2
- Workbook: 100000 rows x 10 columns
- Iterations: 5 fresh processes per runtime and scenario

| Scenario | Runtime | Elapsed (ms) | First row (ms) | Managed allocation (MB) | Peak working set (MB) | Peak private (MB) | Speedup |
| --- | --- | ---: | ---: | ---: | ---: | ---: | ---: |
| Cold | MiniExcelV2 | 2891.79 | 756.86 | 1555 | 53.56 | 28.78 | 1x |
| Cold | MiniExcelRust | 1041.49 | 357.62 | 116.47 | 42.71 | 25.16 | 2.78x |
| Warm | MiniExcelV2 | 6136.27 | 393.94 | 4674.55 | 57.07 | 33.36 | 1x |
| Warm | MiniExcelRust | 2770.39 | 326.87 | 304.99 | 46.08 | 28.93 | 2.21x |

Managed allocation excludes allocations made inside Rust. Peak working set and peak private bytes include the complete process.
