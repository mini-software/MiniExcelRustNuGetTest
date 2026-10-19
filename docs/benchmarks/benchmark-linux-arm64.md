# MiniExcelRust benchmark (linux-arm64)

- Date (UTC): 2026-10-19T08:50:52.6730793Z
- OS: Ubuntu 24.04.5 LTS
- Architecture: Arm64
- .NET SDK: 10.0.402
- Target framework: net10.0
- .NET runtime: .NET 10.0.13
- Baseline: MiniExcel 2.0.0-preview.4
- Candidate: MiniExcelRust 0.1.0-preview.2
- Workbook: 100000 rows x 10 columns
- Iterations: 5 fresh processes per runtime and scenario

| Scenario | Runtime | Elapsed (ms) | First row (ms) | Managed allocation (MB) | Peak working set (MB) | Peak private (MB) | Speedup |
| --- | --- | ---: | ---: | ---: | ---: | ---: | ---: |
| Cold | MiniExcelV2 | 3596.66 | 1055.56 | 1555.66 | 123.97 | 176.19 | 1x |
| Cold | MiniExcelRust | 1056.86 | 330.4 | 113.29 | 108.22 | 177.14 | 3.4x |
| Warm | MiniExcelV2 | 8084.27 | 575.89 | 4674.19 | 127.85 | 183.18 | 1x |
| Warm | MiniExcelRust | 2785.07 | 323.11 | 304.99 | 112.17 | 181.1 | 2.9x |

Managed allocation excludes allocations made inside Rust. Peak working set and peak private bytes include the complete process.
