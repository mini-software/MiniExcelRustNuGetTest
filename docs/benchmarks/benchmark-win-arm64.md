# MiniExcelRust benchmark (win-arm64)

- Date (UTC): 2026-11-23T12:32:01.9638536Z
- OS: Microsoft Windows 10.0.26200
- Architecture: Arm64
- .NET SDK: 11.0.100
- Target framework: net10.0
- .NET runtime: .NET 10.0.15
- Baseline: MiniExcel 2.0.0-preview.4
- Candidate: MiniExcelRust 0.1.0-preview.2
- Workbook: 100000 rows x 10 columns
- Iterations: 5 fresh processes per runtime and scenario

| Scenario | Runtime | Elapsed (ms) | First row (ms) | Managed allocation (MB) | Peak working set (MB) | Peak private (MB) | Speedup |
| --- | --- | ---: | ---: | ---: | ---: | ---: | ---: |
| Cold | MiniExcelV2 | 3038.12 | 947.49 | 1553.98 | 105.45 | 77.07 | 1x |
| Cold | MiniExcelRust | 1093.89 | 344.45 | 111.08 | 95.31 | 74 | 2.78x |
| Warm | MiniExcelV2 | 6537.5 | 529.74 | 4674.19 | 107.96 | 81.3 | 1x |
| Warm | MiniExcelRust | 2984.3 | 337.59 | 304.99 | 98.68 | 78.49 | 2.19x |

Managed allocation excludes allocations made inside Rust. Peak working set and peak private bytes include the complete process.
