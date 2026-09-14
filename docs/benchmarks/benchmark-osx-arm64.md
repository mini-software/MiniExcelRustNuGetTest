# MiniExcelRust benchmark (osx-arm64)

- Date (UTC): 2026-09-14T08:55:16.1115700Z
- OS: macOS 26.6.2
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
| Cold | MiniExcelV2 | 2867.89 | 1021.93 | 1535.01 | 78.12 | 0 | 1x |
| Cold | MiniExcelRust | 1141.01 | 305.01 | 117.21 | 53.26 | 0 | 2.51x |
| Warm | MiniExcelV2 | 4859.98 | 563.74 | 4675.22 | 78.36 | 0 | 1x |
| Warm | MiniExcelRust | 2991.35 | 340.41 | 304.99 | 54.38 | 0 | 1.62x |

Managed allocation excludes allocations made inside Rust. Peak working set and peak private bytes include the complete process.
