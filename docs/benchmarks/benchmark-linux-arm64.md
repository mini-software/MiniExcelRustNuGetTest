# MiniExcelRust benchmark (linux-arm64)

- Date (UTC): 2026-11-23T12:31:46.6236619Z
- OS: Ubuntu 24.04.5 LTS
- Architecture: Arm64
- .NET SDK: 11.0.100
- Target framework: net10.0
- .NET runtime: .NET 10.0.15
- Baseline: MiniExcel 2.0.0-preview.4
- Candidate: MiniExcelRust 0.1.0-preview.2
- Workbook: 100000 rows x 10 columns
- Iterations: 5 fresh processes per runtime and scenario

| Scenario | Runtime | Elapsed (ms) | First row (ms) | Managed allocation (MB) | Peak working set (MB) | Peak private (MB) | Speedup |
| --- | --- | ---: | ---: | ---: | ---: | ---: | ---: |
| Cold | MiniExcelV2 | 3528.25 | 1050.8 | 1555.76 | 123.93 | 176.4 | 1x |
| Cold | MiniExcelRust | 1061.69 | 330.36 | 113.05 | 108.66 | 177.4 | 3.32x |
| Warm | MiniExcelV2 | 8175.5 | 575.15 | 4674.19 | 127.78 | 180.16 | 1x |
| Warm | MiniExcelRust | 2770.9 | 323.13 | 304.99 | 111.59 | 177.5 | 2.95x |

Managed allocation excludes allocations made inside Rust. Peak working set and peak private bytes include the complete process.
