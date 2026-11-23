# MiniExcelRust benchmark (linux-x64)

- Date (UTC): 2026-11-23T12:31:14.2887739Z
- OS: Ubuntu 26.04.1 LTS
- Architecture: X64
- .NET SDK: 11.0.100
- Target framework: net10.0
- .NET runtime: .NET 10.0.15
- Baseline: MiniExcel 2.0.0-preview.4
- Candidate: MiniExcelRust 0.1.0-preview.2
- Workbook: 100000 rows x 10 columns
- Iterations: 5 fresh processes per runtime and scenario

| Scenario | Runtime | Elapsed (ms) | First row (ms) | Managed allocation (MB) | Peak working set (MB) | Peak private (MB) | Speedup |
| --- | --- | ---: | ---: | ---: | ---: | ---: | ---: |
| Cold | MiniExcelV2 | 2408.56 | 860.29 | 1547.33 | 75.44 | 132.55 | 1x |
| Cold | MiniExcelRust | 1103.58 | 328.23 | 116.02 | 60.41 | 126.27 | 2.18x |
| Warm | MiniExcelV2 | 4722.85 | 450.2 | 4674.55 | 79.28 | 132.99 | 1x |
| Warm | MiniExcelRust | 2828.37 | 320.54 | 304.99 | 64.67 | 130.38 | 1.67x |

Managed allocation excludes allocations made inside Rust. Peak working set and peak private bytes include the complete process.
