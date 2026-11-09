# MiniExcelRust benchmark (linux-arm64)

- Date (UTC): 2026-11-09T11:28:48.7457847Z
- OS: Ubuntu 24.04.5 LTS
- Architecture: Arm64
- .NET SDK: 10.0.403
- Target framework: net10.0
- .NET runtime: .NET 10.0.14
- Baseline: MiniExcel 2.0.0-preview.4
- Candidate: MiniExcelRust 0.1.0-preview.2
- Workbook: 100000 rows x 10 columns
- Iterations: 5 fresh processes per runtime and scenario

| Scenario | Runtime | Elapsed (ms) | First row (ms) | Managed allocation (MB) | Peak working set (MB) | Peak private (MB) | Speedup |
| --- | --- | ---: | ---: | ---: | ---: | ---: | ---: |
| Cold | MiniExcelV2 | 3501.67 | 1040.4 | 1555.59 | 123.93 | 182.61 | 1x |
| Cold | MiniExcelRust | 1059.57 | 330.71 | 113.15 | 108.04 | 173.89 | 3.3x |
| Warm | MiniExcelV2 | 8106.43 | 568.35 | 4674.2 | 127.56 | 183.04 | 1x |
| Warm | MiniExcelRust | 2775.82 | 322.99 | 304.99 | 112.39 | 177.91 | 2.92x |

Managed allocation excludes allocations made inside Rust. Peak working set and peak private bytes include the complete process.
