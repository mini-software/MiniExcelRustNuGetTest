# MiniExcelRust benchmark (linux-x64)

- Date (UTC): 2026-11-16T11:12:09.7108827Z
- OS: Ubuntu 26.04.1 LTS
- Architecture: X64
- .NET SDK: 11.0.100
- Target framework: net10.0
- .NET runtime: .NET 10.0.15
- Baseline: MiniExcel 2.0.0-preview.4
- Candidate: MiniExcelRust 0.1.0-preview.2
- Workbook: 100000 rows x 10 columns
- Iterations: 5 fresh processes per runtime and scenario

| Scenario | Runtime | Elapsed (ms) | First row (ms) | Managed allocation (MB) | Peak working set (MB) | Peak private (MB) | Speedup |
| --- | --- | ---: | ---: | ---: | ---: | ---: | ---: |
| Cold | MiniExcelV2 | 2182.83 | 622.15 | 1543.31 | 75.35 | 129.34 | 1x |
| Cold | MiniExcelRust | 790.43 | 236.22 | 123.49 | 60.28 | 126.29 | 2.76x |
| Warm | MiniExcelV2 | 4095.99 | 277.75 | 4674.55 | 78.7 | 132.55 | 1x |
| Warm | MiniExcelRust | 1931.4 | 216.27 | 304.99 | 64.36 | 130.3 | 2.12x |

Managed allocation excludes allocations made inside Rust. Peak working set and peak private bytes include the complete process.
