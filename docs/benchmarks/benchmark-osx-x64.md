# MiniExcelRust benchmark (osx-x64)

- Date (UTC): 2026-10-19T08:55:13.3742660Z
- OS: macOS 15.7.9
- Architecture: X64
- .NET SDK: 10.0.402
- Target framework: net10.0
- .NET runtime: .NET 10.0.13
- Baseline: MiniExcel 2.0.0-preview.4
- Candidate: MiniExcelRust 0.1.0-preview.2
- Workbook: 100000 rows x 10 columns
- Iterations: 5 fresh processes per runtime and scenario

| Scenario | Runtime | Elapsed (ms) | First row (ms) | Managed allocation (MB) | Peak working set (MB) | Peak private (MB) | Speedup |
| --- | --- | ---: | ---: | ---: | ---: | ---: | ---: |
| Cold | MiniExcelV2 | 8604.59 | 2442.07 | 1555.11 | 43.55 | 0 | 1x |
| Cold | MiniExcelRust | 3493.68 | 1197.13 | 108.56 | 32.51 | 0 | 2.46x |
| Warm | MiniExcelV2 | 18943.96 | 1288.16 | 4675.23 | 43.6 | 0 | 1x |
| Warm | MiniExcelRust | 11782.38 | 1242.34 | 304.99 | 32.84 | 0 | 1.61x |

Managed allocation excludes allocations made inside Rust. Peak working set and peak private bytes include the complete process.
