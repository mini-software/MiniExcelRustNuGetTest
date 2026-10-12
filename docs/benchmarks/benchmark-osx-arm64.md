# MiniExcelRust benchmark (osx-arm64)

- Date (UTC): 2026-10-12T10:47:12.5403570Z
- OS: macOS 26.6.2
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
| Cold | MiniExcelV2 | 2863.33 | 1050.11 | 1547.4 | 78.1 | 0 | 1x |
| Cold | MiniExcelRust | 1162.95 | 333.44 | 117.63 | 54.54 | 0 | 2.46x |
| Warm | MiniExcelV2 | 4568.89 | 505.18 | 4675.22 | 78.67 | 0 | 1x |
| Warm | MiniExcelRust | 2709 | 288.64 | 304.99 | 54.38 | 0 | 1.69x |

Managed allocation excludes allocations made inside Rust. Peak working set and peak private bytes include the complete process.
