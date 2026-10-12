# Cross-platform benchmark results

Last updated (UTC): 2026-10-12 10:49:17

Each platform validates every returned row and cell before timing equivalent queries in alternating fresh processes.

| RID | Scenario | .NET runtime | MiniExcel | MiniExcel (ms) | MiniExcelRust (ms) | Speedup | Allocation reduction | Working-set reduction |
| --- | --- | --- | --- | ---: | ---: | ---: | ---: | ---: |
| [linux-arm64](benchmark-linux-arm64.md) | Cold | .NET 10.0.12 | 2.0.0-preview.4 | 3584.25 | 1042.94 | 3.44x | 92.7% | 12.5% |
| [linux-arm64](benchmark-linux-arm64.md) | Warm | .NET 10.0.12 | 2.0.0-preview.4 | 8103.99 | 2733.22 | 2.96x | 93.5% | 11.9% |
| [linux-x64](benchmark-linux-x64.md) | Cold | .NET 10.0.12 | 2.0.0-preview.4 | 2856.94 | 1288.34 | 2.22x | 92.7% | 19.2% |
| [linux-x64](benchmark-linux-x64.md) | Warm | .NET 10.0.12 | 2.0.0-preview.4 | 5906.86 | 3305.74 | 1.79x | 93.5% | 18.8% |
| [osx-arm64](benchmark-osx-arm64.md) | Cold | .NET 10.0.12 | 2.0.0-preview.4 | 2863.33 | 1162.95 | 2.46x | 92.4% | 30.2% |
| [osx-arm64](benchmark-osx-arm64.md) | Warm | .NET 10.0.12 | 2.0.0-preview.4 | 4568.89 | 2709 | 1.69x | 93.5% | 30.9% |
| [osx-x64](benchmark-osx-x64.md) | Cold | .NET 10.0.12 | 2.0.0-preview.4 | 5201.11 | 2330.54 | 2.23x | 92.8% | 24.4% |
| [osx-x64](benchmark-osx-x64.md) | Warm | .NET 10.0.12 | 2.0.0-preview.4 | 12091.95 | 7052.18 | 1.71x | 93.5% | 24.3% |
| [win-arm64](benchmark-win-arm64.md) | Cold | .NET 10.0.12 | 2.0.0-preview.4 | 2951.6 | 1093 | 2.7x | 92.8% | 24.6% |
| [win-arm64](benchmark-win-arm64.md) | Warm | .NET 10.0.12 | 2.0.0-preview.4 | 6647.06 | 3116.98 | 2.13x | 93.5% | 25% |
| [win-x64](benchmark-win-x64.md) | Cold | .NET 10.0.12 | 2.0.0-preview.4 | 2795.22 | 1438.31 | 1.94x | 92.8% | 18.2% |
| [win-x64](benchmark-win-x64.md) | Warm | .NET 10.0.12 | 2.0.0-preview.4 | 5116.05 | 4083.57 | 1.25x | 93.5% | 18.8% |

Managed allocation excludes allocations made inside Rust. Each linked report includes peak process memory and environment metadata; the adjacent JSON contains every raw iteration and input hash.
