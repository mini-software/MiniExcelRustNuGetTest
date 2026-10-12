# MiniExcelRust benchmark (win-x64)

- Date (UTC): 2026-10-12T10:47:28.4998382Z
- OS: Microsoft Windows 10.0.26100
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
| Cold | MiniExcelV2 | 2795.22 | 1035.25 | 1552.01 | 55.85 | 28.74 | 1x |
| Cold | MiniExcelRust | 1438.31 | 449.25 | 111.59 | 45.66 | 25.1 | 1.94x |
| Warm | MiniExcelV2 | 5116.05 | 492.97 | 4674.55 | 60.03 | 33.29 | 1x |
| Warm | MiniExcelRust | 4083.57 | 444.04 | 304.99 | 48.74 | 28.6 | 1.25x |

Managed allocation excludes allocations made inside Rust. Peak working set and peak private bytes include the complete process.
