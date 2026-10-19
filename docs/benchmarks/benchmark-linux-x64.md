# MiniExcelRust benchmark (linux-x64)

- Date (UTC): 2026-10-19T08:50:53.1864671Z
- OS: Ubuntu 24.04.5 LTS
- Architecture: X64
- .NET SDK: 10.0.402
- Target framework: net10.0
- .NET runtime: .NET 10.0.13
- Baseline: MiniExcel 2.0.0-preview.4
- Candidate: MiniExcelRust 0.1.0-preview.2
- Workbook: 100000 rows x 10 columns
- Iterations: 5 fresh processes per runtime and scenario

| Scenario | Runtime | Elapsed (ms) | First row (ms) | Managed allocation (MB) | Peak working set (MB) | Peak private (MB) | Speedup |
| --- | --- | ---: | ---: | ---: | ---: | ---: | ---: |
| Cold | MiniExcelV2 | 3095.75 | 1059.67 | 1554.31 | 140.1 | 193.35 | 1x |
| Cold | MiniExcelRust | 1259.23 | 342.33 | 113.59 | 124.4 | 190 | 2.46x |
| Warm | MiniExcelV2 | 6726.53 | 528.98 | 4674.17 | 144.11 | 200.58 | 1x |
| Warm | MiniExcelRust | 2959.33 | 327.93 | 304.99 | 128.67 | 197.46 | 2.27x |

Managed allocation excludes allocations made inside Rust. Peak working set and peak private bytes include the complete process.
