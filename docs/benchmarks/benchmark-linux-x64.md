# MiniExcelRust benchmark (linux-x64)

- Date (UTC): 2026-10-12T10:47:21.9969806Z
- OS: Ubuntu 24.04.5 LTS
- Architecture: X64
- .NET SDK: 10.0.401
- Target framework: net10.0
- .NET runtime: .NET 10.0.12
- Baseline: MiniExcel 2.0.0-preview.4
- Candidate: MiniExcelRust 0.1.0-preview.2
- Workbook: 100000 rows x 10 columns
- Iterations: 5 fresh processes per runtime and scenario

| Scenario | Runtime | Elapsed (ms) | First row (ms) | Managed allocation (MB) | Peak working set (MB) | Peak private (MB) | Speedup |
| --- | --- | ---: | ---: | ---: | ---: | ---: | ---: |
| Cold | MiniExcelV2 | 2856.94 | 1052.29 | 1553.73 | 75.29 | 129.29 | 1x |
| Cold | MiniExcelRust | 1288.34 | 390.2 | 114.06 | 60.86 | 126.64 | 2.22x |
| Warm | MiniExcelV2 | 5906.86 | 539.11 | 4674.56 | 79.04 | 136.15 | 1x |
| Warm | MiniExcelRust | 3305.74 | 381.52 | 304.99 | 64.17 | 130.05 | 1.79x |

Managed allocation excludes allocations made inside Rust. Peak working set and peak private bytes include the complete process.
