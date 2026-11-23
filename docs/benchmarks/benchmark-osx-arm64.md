# MiniExcelRust benchmark (osx-arm64)

- Date (UTC): 2026-11-23T12:32:10.2996840Z
- OS: macOS 26.7.2
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
| Cold | MiniExcelV2 | 2862.19 | 1021.14 | 1537.96 | 77.55 | 0 | 1x |
| Cold | MiniExcelRust | 1210.88 | 350.02 | 117.05 | 54.33 | 0 | 2.36x |
| Warm | MiniExcelV2 | 5645.68 | 681.91 | 4675.22 | 78.63 | 0 | 1x |
| Warm | MiniExcelRust | 3309.44 | 331.49 | 304.99 | 54.44 | 0 | 1.71x |

Managed allocation excludes allocations made inside Rust. Peak working set and peak private bytes include the complete process.
