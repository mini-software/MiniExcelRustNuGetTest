# MiniExcelRust benchmark (linux-arm64)

- Date (UTC): 2026-09-14T08:55:12.7640532Z
- OS: Ubuntu 24.04.5 LTS
- Architecture: Arm64
- .NET SDK: 10.0.401
- Target framework: net10.0
- .NET runtime: .NET 10.0.12
- Baseline: MiniExcel 2.0.0-preview.4
- Candidate: MiniExcelRust 0.1.0-preview.2
- Workbook: 100000 rows x 10 columns
- Iterations: 5 fresh processes per runtime and scenario

| Scenario | Runtime | Elapsed (ms) | First row (ms) | Managed allocation (MB) | Peak working set (MB) | Peak private (MB) | Speedup |
| --- | --- | ---: | ---: | ---: | ---: | ---: | ---: |
| Cold | MiniExcelV2 | 3553.09 | 1056.5 | 1555.68 | 123.78 | 176.3 | 1x |
| Cold | MiniExcelRust | 1056.78 | 330.14 | 113.2 | 108.03 | 173.9 | 3.36x |
| Warm | MiniExcelV2 | 8110.62 | 574.97 | 4674.19 | 127.86 | 183.39 | 1x |
| Warm | MiniExcelRust | 2792.15 | 323.03 | 304.99 | 112.2 | 177.9 | 2.9x |

Managed allocation excludes allocations made inside Rust. Peak working set and peak private bytes include the complete process.
