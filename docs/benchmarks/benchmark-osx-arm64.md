# MiniExcelRust benchmark (osx-arm64)

- Date (UTC): 2026-10-19T08:52:11.7990700Z
- OS: macOS 26.6.2
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
| Cold | MiniExcelV2 | 3170.19 | 1246.35 | 1541.45 | 77.57 | 0 | 1x |
| Cold | MiniExcelRust | 1262.87 | 363.74 | 115.26 | 52.78 | 0 | 2.51x |
| Warm | MiniExcelV2 | 4980.18 | 610.86 | 4675.22 | 77.83 | 0 | 1x |
| Warm | MiniExcelRust | 3217.08 | 336.83 | 304.99 | 54.75 | 0 | 1.55x |

Managed allocation excludes allocations made inside Rust. Peak working set and peak private bytes include the complete process.
