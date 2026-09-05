# Cross-platform benchmark results

Last updated (UTC): 2026-09-05 14:05:30

Each platform validates every returned row and cell before timing equivalent queries in alternating fresh processes.

| RID | Scenario | .NET runtime | MiniExcel | MiniExcel (ms) | MiniExcelRust (ms) | Speedup | Allocation reduction | Working-set reduction |
| --- | --- | --- | --- | ---: | ---: | ---: | ---: | ---: |
| [win-x64](benchmark-win-x64.md) | Cold | .NET 10.0.3 | 2.0.0-preview.4 | 2891.79 | 1041.49 | 2.78x | 92.5% | 20.3% |
| [win-x64](benchmark-win-x64.md) | Warm | .NET 10.0.3 | 2.0.0-preview.4 | 6136.27 | 2770.39 | 2.21x | 93.5% | 19.3% |

Managed allocation excludes allocations made inside Rust. Each linked report includes peak process memory and environment metadata; the adjacent JSON contains every raw iteration and input hash.
