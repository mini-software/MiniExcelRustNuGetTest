# Cross-platform benchmark results

Last updated (UTC): 2026-11-23 12:33:32

Each platform validates every returned row and cell before timing equivalent queries in alternating fresh processes.

| RID | Scenario | .NET runtime | MiniExcel | MiniExcel (ms) | MiniExcelRust (ms) | Speedup | Allocation reduction | Working-set reduction |
| --- | --- | --- | --- | ---: | ---: | ---: | ---: | ---: |
| [linux-arm64](benchmark-linux-arm64.md) | Cold | .NET 10.0.15 | 2.0.0-preview.4 | 3528.25 | 1061.69 | 3.32x | 92.7% | 12.3% |
| [linux-arm64](benchmark-linux-arm64.md) | Warm | .NET 10.0.15 | 2.0.0-preview.4 | 8175.5 | 2770.9 | 2.95x | 93.5% | 12.7% |
| [linux-x64](benchmark-linux-x64.md) | Cold | .NET 10.0.15 | 2.0.0-preview.4 | 2408.56 | 1103.58 | 2.18x | 92.5% | 19.9% |
| [linux-x64](benchmark-linux-x64.md) | Warm | .NET 10.0.15 | 2.0.0-preview.4 | 4722.85 | 2828.37 | 1.67x | 93.5% | 18.4% |
| [osx-arm64](benchmark-osx-arm64.md) | Cold | .NET 10.0.15 | 2.0.0-preview.4 | 2862.19 | 1210.88 | 2.36x | 92.4% | 29.9% |
| [osx-arm64](benchmark-osx-arm64.md) | Warm | .NET 10.0.15 | 2.0.0-preview.4 | 5645.68 | 3309.44 | 1.71x | 93.5% | 30.8% |
| [osx-x64](benchmark-osx-x64.md) | Cold | .NET 10.0.15 | 2.0.0-preview.4 | 5413.63 | 2219.16 | 2.44x | 92.8% | 28.9% |
| [osx-x64](benchmark-osx-x64.md) | Warm | .NET 10.0.15 | 2.0.0-preview.4 | 12755.7 | 6659.29 | 1.92x | 93.5% | 24.4% |
| [win-arm64](benchmark-win-arm64.md) | Cold | .NET 10.0.15 | 2.0.0-preview.4 | 3038.12 | 1093.89 | 2.78x | 92.9% | 9.6% |
| [win-arm64](benchmark-win-arm64.md) | Warm | .NET 10.0.15 | 2.0.0-preview.4 | 6537.5 | 2984.3 | 2.19x | 93.5% | 8.6% |
| [win-x64](benchmark-win-x64.md) | Cold | .NET 10.0.15 | 2.0.0-preview.4 | 2636.85 | 1532.29 | 1.72x | 92.9% | 18.3% |
| [win-x64](benchmark-win-x64.md) | Warm | .NET 10.0.15 | 2.0.0-preview.4 | 5230.34 | 4236.16 | 1.23x | 93.5% | 19.2% |

Managed allocation excludes allocations made inside Rust. Each linked report includes peak process memory and environment metadata; the adjacent JSON contains every raw iteration and input hash.
