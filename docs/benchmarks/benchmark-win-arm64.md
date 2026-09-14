# MiniExcelRust benchmark (win-arm64)

- Date (UTC): 2026-09-14T08:56:28.2987595Z
- OS: Microsoft Windows 10.0.26200
- Architecture: Arm64
- .NET SDK: 10.0.401
- Target framework: net10.0
- .NET runtime: .NET 10.0.12
- Baseline: MiniExcel 2.0.0-preview.4
- Candidate: MiniExcelRust 0.1.0-preview.2
- Workbook: 100000 rows x 10 columns
- Iterations: 5 fresh processes per runtime and scenario

| Scenario | Runtime | Elapsed (ms) | First row (ms) | Managed allocation (MB) | Peak working set (MB) | Peak private (MB) | Speedup |
| --- | --- | ---: | ---: | ---: | ---: | ---: | ---: |
| Cold | MiniExcelV2 | 3038.7 | 959.95 | 1553.68 | 106.32 | 77.14 | 1x |
| Cold | MiniExcelRust | 1080.11 | 348.29 | 111.19 | 95.85 | 74.06 | 2.81x |
| Warm | MiniExcelV2 | 6587.04 | 540.91 | 4674.2 | 109.88 | 81.34 | 1x |
| Warm | MiniExcelRust | 2947.6 | 341.97 | 304.99 | 99.66 | 78.49 | 2.23x |

Managed allocation excludes allocations made inside Rust. Peak working set and peak private bytes include the complete process.
