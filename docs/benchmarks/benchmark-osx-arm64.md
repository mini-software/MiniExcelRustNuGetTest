# MiniExcelRust benchmark (osx-arm64)

- Date (UTC): 2026-11-09T11:28:17.2608080Z
- OS: macOS 26.7.2
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
| Cold | MiniExcelV2 | 2964.03 | 1019.15 | 1535.64 | 78.29 | 0 | 1x |
| Cold | MiniExcelRust | 1009.93 | 292.33 | 117.34 | 54.44 | 0 | 2.93x |
| Warm | MiniExcelV2 | 4214.26 | 500.78 | 4675.22 | 78.8 | 0 | 1x |
| Warm | MiniExcelRust | 2780.46 | 299.58 | 304.99 | 54.72 | 0 | 1.52x |

Managed allocation excludes allocations made inside Rust. Peak working set and peak private bytes include the complete process.
