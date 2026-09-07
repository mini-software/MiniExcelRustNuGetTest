# MiniExcelRust benchmark (linux-arm64)

- Date (UTC): 2026-09-07T08:18:15.0079026Z
- OS: Ubuntu 24.04.4 LTS
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
| Cold | MiniExcelV2 | 3524.58 | 1050.79 | 1555.73 | 123.49 | 176.23 | 1x |
| Cold | MiniExcelRust | 1052.69 | 330.52 | 113.28 | 108.11 | 173.88 | 3.35x |
| Warm | MiniExcelV2 | 8041.09 | 572.99 | 4674.19 | 127.31 | 179.89 | 1x |
| Warm | MiniExcelRust | 2768.08 | 323.06 | 304.99 | 111.84 | 177.85 | 2.9x |

Managed allocation excludes allocations made inside Rust. Peak working set and peak private bytes include the complete process.
