# MiniExcelRust benchmark (linux-arm64)

- Date (UTC): 2026-10-26T08:54:05.4754822Z
- OS: Ubuntu 24.04.5 LTS
- Architecture: Arm64
- .NET SDK: 10.0.403
- Target framework: net10.0
- .NET runtime: .NET 10.0.14
- Baseline: MiniExcel 2.0.0-preview.4
- Candidate: MiniExcelRust 0.1.0-preview.2
- Workbook: 100000 rows x 10 columns
- Iterations: 5 fresh processes per runtime and scenario

| Scenario | Runtime | Elapsed (ms) | First row (ms) | Managed allocation (MB) | Peak working set (MB) | Peak private (MB) | Speedup |
| --- | --- | ---: | ---: | ---: | ---: | ---: | ---: |
| Cold | MiniExcelV2 | 3546.78 | 1048.58 | 1555.66 | 123.88 | 179.48 | 1x |
| Cold | MiniExcelRust | 1046.15 | 330.21 | 113.44 | 108.34 | 173.91 | 3.39x |
| Warm | MiniExcelV2 | 8117.5 | 573.52 | 4674.19 | 128.04 | 183.39 | 1x |
| Warm | MiniExcelRust | 2750.09 | 323.36 | 304.99 | 112.38 | 181.1 | 2.95x |

Managed allocation excludes allocations made inside Rust. Peak working set and peak private bytes include the complete process.
