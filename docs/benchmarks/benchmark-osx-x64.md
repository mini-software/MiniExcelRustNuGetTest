# MiniExcelRust benchmark (osx-x64)

- Date (UTC): 2026-11-23T12:33:32.2066660Z
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
| Cold | MiniExcelV2 | 5413.63 | 1647.94 | 1555.23 | 43.34 | 0 | 1x |
| Cold | MiniExcelRust | 2219.16 | 711.95 | 112.38 | 30.8 | 0 | 2.44x |
| Warm | MiniExcelV2 | 12755.7 | 909.56 | 4675.22 | 43.35 | 0 | 1x |
| Warm | MiniExcelRust | 6659.29 | 672.34 | 304.99 | 32.78 | 0 | 1.92x |

Managed allocation excludes allocations made inside Rust. Peak working set and peak private bytes include the complete process.
