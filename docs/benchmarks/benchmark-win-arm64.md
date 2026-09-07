# MiniExcelRust benchmark (win-arm64)

- Date (UTC): 2026-09-07T08:18:57.7571632Z
- OS: Microsoft Windows 10.0.26200
- Architecture: Arm64
- .NET SDK: 10.0.400
- Target framework: net10.0
- .NET runtime: .NET 10.0.11
- Baseline: MiniExcel 2.0.0-preview.4
- Candidate: MiniExcelRust 0.1.0-preview.2
- Workbook: 100000 rows x 10 columns
- Iterations: 5 fresh processes per runtime and scenario

| Scenario | Runtime | Elapsed (ms) | First row (ms) | Managed allocation (MB) | Peak working set (MB) | Peak private (MB) | Speedup |
| --- | --- | ---: | ---: | ---: | ---: | ---: | ---: |
| Cold | MiniExcelV2 | 3134.51 | 1013.24 | 1553.37 | 105.57 | 76.99 | 1x |
| Cold | MiniExcelRust | 1102.53 | 349.77 | 114.56 | 93.84 | 73.32 | 2.84x |
| Warm | MiniExcelV2 | 6597.02 | 536.83 | 4674.2 | 109.33 | 80.7 | 1x |
| Warm | MiniExcelRust | 2934.18 | 341.37 | 304.99 | 98.11 | 77.92 | 2.25x |

Managed allocation excludes allocations made inside Rust. Peak working set and peak private bytes include the complete process.
