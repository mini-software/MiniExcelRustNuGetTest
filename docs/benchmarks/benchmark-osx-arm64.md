# MiniExcelRust benchmark (osx-arm64)

- Date (UTC): 2026-09-07T08:17:27.4570550Z
- OS: macOS 26.6.2
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
| Cold | MiniExcelV2 | 2583.35 | 863.93 | 1530.26 | 78.15 | 0 | 1x |
| Cold | MiniExcelRust | 943 | 272.23 | 120.54 | 53.92 | 0 | 2.74x |
| Warm | MiniExcelV2 | 3868.17 | 441.08 | 4675.14 | 78.84 | 0 | 1x |
| Warm | MiniExcelRust | 2348.76 | 254.92 | 304.99 | 55.11 | 0 | 1.65x |

Managed allocation excludes allocations made inside Rust. Peak working set and peak private bytes include the complete process.
