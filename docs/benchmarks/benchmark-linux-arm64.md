# MiniExcelRust benchmark (linux-arm64)

- Date (UTC): 2026-10-12T10:47:27.0854454Z
- OS: Ubuntu 24.04.5 LTS
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
| Cold | MiniExcelV2 | 3584.25 | 1055.64 | 1555.83 | 123.64 | 179.48 | 1x |
| Cold | MiniExcelRust | 1042.94 | 330.02 | 113.47 | 108.19 | 173.95 | 3.44x |
| Warm | MiniExcelV2 | 8103.99 | 569.06 | 4674.2 | 127.41 | 179.85 | 1x |
| Warm | MiniExcelRust | 2733.22 | 322.77 | 304.99 | 112.26 | 177.92 | 2.96x |

Managed allocation excludes allocations made inside Rust. Peak working set and peak private bytes include the complete process.
