# MiniExcelRust benchmark (linux-x64)

- Date (UTC): 2026-11-09T11:28:19.3351476Z
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
| Cold | MiniExcelV2 | 2371.25 | 853.29 | 1549.91 | 75.67 | 129.47 | 1x |
| Cold | MiniExcelRust | 1097.49 | 325.12 | 117.15 | 60.23 | 125.94 | 2.16x |
| Warm | MiniExcelV2 | 4533.5 | 412.16 | 4674.55 | 79.35 | 136.09 | 1x |
| Warm | MiniExcelRust | 2816.59 | 315.98 | 304.99 | 64.38 | 133.3 | 1.61x |

Managed allocation excludes allocations made inside Rust. Peak working set and peak private bytes include the complete process.
