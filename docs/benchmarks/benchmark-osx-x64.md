# MiniExcelRust benchmark (osx-x64)

- Date (UTC): 2026-09-07T08:23:09.5721600Z
- OS: macOS 15.7.9
- Architecture: X64
- .NET SDK: 10.0.400
- Target framework: net10.0
- .NET runtime: .NET 10.0.11
- Baseline: MiniExcel 2.0.0-preview.4
- Candidate: MiniExcelRust 0.1.0-preview.2
- Workbook: 100000 rows x 10 columns
- Iterations: 5 fresh processes per runtime and scenario

| Scenario | Runtime | Elapsed (ms) | First row (ms) | Managed allocation (MB) | Peak working set (MB) | Peak private (MB) | Speedup |
| --- | --- | ---: | ---: | ---: | ---: | ---: | ---: |
| Cold | MiniExcelV2 | 6952.98 | 1976.43 | 1555.78 | 44.44 | 0 | 1x |
| Cold | MiniExcelRust | 4133.51 | 1350.73 | 108.74 | 32.8 | 0 | 1.68x |
| Warm | MiniExcelV2 | 17775.09 | 1254.71 | 4675.22 | 43.65 | 0 | 1x |
| Warm | MiniExcelRust | 8440.59 | 962.89 | 304.99 | 33.77 | 0 | 2.11x |

Managed allocation excludes allocations made inside Rust. Peak working set and peak private bytes include the complete process.
