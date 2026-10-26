# MiniExcelRust benchmark (linux-x64)

- Date (UTC): 2026-10-26T08:53:45.9056492Z
- OS: Ubuntu 24.04.5 LTS
- Architecture: X64
- .NET SDK: 10.0.403
- Target framework: net10.0
- .NET runtime: .NET 10.0.14
- Baseline: MiniExcel 2.0.0-preview.4
- Candidate: MiniExcelRust 0.1.0-preview.2
- Workbook: 100000 rows x 10 columns
- Iterations: 5 fresh processes per runtime and scenario

| Scenario | Runtime | Elapsed (ms) | First row (ms) | Managed allocation (MB) | Peak working set (MB) | Peak private (MB) | Speedup |
| --- | --- | ---: | ---: | ---: | ---: | ---: | ---: |
| Cold | MiniExcelV2 | 2842.14 | 1057.94 | 1554.45 | 75.94 | 129.28 | 1x |
| Cold | MiniExcelRust | 1330.67 | 402.54 | 112.53 | 60.4 | 129.05 | 2.14x |
| Warm | MiniExcelV2 | 5874.63 | 536.09 | 4674.55 | 79.57 | 136.06 | 1x |
| Warm | MiniExcelRust | 3211.41 | 380.77 | 304.99 | 64.59 | 133.27 | 1.83x |

Managed allocation excludes allocations made inside Rust. Peak working set and peak private bytes include the complete process.
