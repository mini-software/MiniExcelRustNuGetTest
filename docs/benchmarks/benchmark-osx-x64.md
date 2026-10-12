# MiniExcelRust benchmark (osx-x64)

- Date (UTC): 2026-10-12T10:49:17.1966060Z
- OS: macOS 15.7.9
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
| Cold | MiniExcelV2 | 5201.11 | 1558.38 | 1552.21 | 43.29 | 0 | 1x |
| Cold | MiniExcelRust | 2330.54 | 722.24 | 111.84 | 32.71 | 0 | 2.23x |
| Warm | MiniExcelV2 | 12091.95 | 901.1 | 4675.23 | 43.27 | 0 | 1x |
| Warm | MiniExcelRust | 7052.18 | 857.2 | 304.99 | 32.75 | 0 | 1.71x |

Managed allocation excludes allocations made inside Rust. Peak working set and peak private bytes include the complete process.
