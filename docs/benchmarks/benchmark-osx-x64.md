# MiniExcelRust benchmark (osx-x64)

- Date (UTC): 2026-11-16T11:17:21.6712750Z
- OS: macOS 15.7.9
- Architecture: X64
- .NET SDK: 10.0.404
- Target framework: net10.0
- .NET runtime: .NET 10.0.15
- Baseline: MiniExcel 2.0.0-preview.4
- Candidate: MiniExcelRust 0.1.0-preview.2
- Workbook: 100000 rows x 10 columns
- Iterations: 5 fresh processes per runtime and scenario

| Scenario | Runtime | Elapsed (ms) | First row (ms) | Managed allocation (MB) | Peak working set (MB) | Peak private (MB) | Speedup |
| --- | --- | ---: | ---: | ---: | ---: | ---: | ---: |
| Cold | MiniExcelV2 | 7219.99 | 2217.21 | 1555 | 43.15 | 0 | 1x |
| Cold | MiniExcelRust | 2772.03 | 944.55 | 109.19 | 32.37 | 0 | 2.6x |
| Warm | MiniExcelV2 | 15080.09 | 1040.27 | 4675.22 | 43.76 | 0 | 1x |
| Warm | MiniExcelRust | 7912.44 | 919.55 | 304.99 | 33.45 | 0 | 1.91x |

Managed allocation excludes allocations made inside Rust. Peak working set and peak private bytes include the complete process.
