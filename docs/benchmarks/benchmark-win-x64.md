# MiniExcelRust benchmark (win-x64)

- Date (UTC): 2026-11-23T12:33:00.3043340Z
- OS: Microsoft Windows 10.0.26100
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
| Cold | MiniExcelV2 | 2636.85 | 984.37 | 1551.77 | 56.2 | 29.06 | 1x |
| Cold | MiniExcelRust | 1532.29 | 461.11 | 110.81 | 45.94 | 25.72 | 1.72x |
| Warm | MiniExcelV2 | 5230.34 | 505.92 | 4674.55 | 60.36 | 33.55 | 1x |
| Warm | MiniExcelRust | 4236.16 | 443.44 | 304.99 | 48.79 | 28.64 | 1.23x |

Managed allocation excludes allocations made inside Rust. Peak working set and peak private bytes include the complete process.
