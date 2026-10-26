# MiniExcelRust benchmark (win-x64)

- Date (UTC): 2026-10-26T08:55:42.0033262Z
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
| Cold | MiniExcelV2 | 3412.7 | 1129.18 | 1553.42 | 120.53 | 93.01 | 1x |
| Cold | MiniExcelRust | 1527.4 | 453.55 | 111.18 | 110.1 | 89.33 | 2.23x |
| Warm | MiniExcelV2 | 6891.14 | 559.64 | 4674.18 | 124.57 | 97.71 | 1x |
| Warm | MiniExcelRust | 3663.39 | 407.76 | 304.99 | 113.37 | 93.22 | 1.88x |

Managed allocation excludes allocations made inside Rust. Peak working set and peak private bytes include the complete process.
