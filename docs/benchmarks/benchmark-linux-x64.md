# MiniExcelRust benchmark (linux-x64)

- Date (UTC): 2026-09-07T08:17:47.9116934Z
- OS: Ubuntu 24.04.4 LTS
- Architecture: X64
- .NET SDK: 10.0.400
- Target framework: net10.0
- .NET runtime: .NET 10.0.11
- Baseline: MiniExcel 2.0.0-preview.4
- Candidate: MiniExcelRust 0.1.0-preview.2
- Workbook: 100000 rows x 10 columns
- Iterations: 5 fresh processes per runtime and scenario

| Scenario | Runtime | Elapsed (ms) | First row (ms) | Managed allocation (MB) | Peak working set (MB) | Peak private (MB) | Speedup |
| --- | --- | ---: | ---: | ---: | ---: | ---: | ---: |
| Cold | MiniExcelV2 | 2559.43 | 839.78 | 1544.23 | 139.47 | 193.51 | 1x |
| Cold | MiniExcelRust | 958.84 | 247.79 | 117.17 | 124.03 | 189.97 | 2.67x |
| Warm | MiniExcelV2 | 4754.54 | 380.31 | 4674.16 | 143.45 | 197.27 | 1x |
| Warm | MiniExcelRust | 2216.26 | 232.06 | 304.99 | 130.61 | 199.8 | 2.15x |

Managed allocation excludes allocations made inside Rust. Peak working set and peak private bytes include the complete process.
