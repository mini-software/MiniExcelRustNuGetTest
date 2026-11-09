# MiniExcelRust benchmark (osx-x64)

- Date (UTC): 2026-11-09T11:32:16.9021630Z
- OS: macOS 15.7.9
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
| Cold | MiniExcelV2 | 6842.79 | 1962.44 | 1555.32 | 42.99 | 0 | 1x |
| Cold | MiniExcelRust | 3367.88 | 1063.16 | 108.34 | 32.32 | 0 | 2.03x |
| Warm | MiniExcelV2 | 19819.13 | 1555.47 | 4675.22 | 43.51 | 0 | 1x |
| Warm | MiniExcelRust | 7578.01 | 874.36 | 304.99 | 32.42 | 0 | 2.62x |

Managed allocation excludes allocations made inside Rust. Peak working set and peak private bytes include the complete process.
