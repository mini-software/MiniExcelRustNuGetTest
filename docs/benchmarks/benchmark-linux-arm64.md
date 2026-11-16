# MiniExcelRust benchmark (linux-arm64)

- Date (UTC): 2026-11-16T11:13:00.0385685Z
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
| Cold | MiniExcelV2 | 3567.03 | 1054.38 | 1553.83 | 123.92 | 176.35 | 1x |
| Cold | MiniExcelRust | 1051.44 | 330.56 | 113.24 | 108.55 | 180.57 | 3.39x |
| Warm | MiniExcelV2 | 8168.68 | 574.36 | 4674.2 | 127.66 | 179.99 | 1x |
| Warm | MiniExcelRust | 2748.68 | 323.14 | 304.99 | 111.67 | 177.53 | 2.97x |

Managed allocation excludes allocations made inside Rust. Peak working set and peak private bytes include the complete process.
