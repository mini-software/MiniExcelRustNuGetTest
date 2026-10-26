# Cross-platform benchmark results

Last updated (UTC): 2026-10-26 08:57:51

Each platform validates every returned row and cell before timing equivalent queries in alternating fresh processes.

| RID | Scenario | .NET runtime | MiniExcel | MiniExcel (ms) | MiniExcelRust (ms) | Speedup | Allocation reduction | Working-set reduction |
| --- | --- | --- | --- | ---: | ---: | ---: | ---: | ---: |
| [linux-arm64](benchmark-linux-arm64.md) | Cold | .NET 10.0.14 | 2.0.0-preview.4 | 3546.78 | 1046.15 | 3.39x | 92.7% | 12.5% |
| [linux-arm64](benchmark-linux-arm64.md) | Warm | .NET 10.0.14 | 2.0.0-preview.4 | 8117.5 | 2750.09 | 2.95x | 93.5% | 12.2% |
| [linux-x64](benchmark-linux-x64.md) | Cold | .NET 10.0.14 | 2.0.0-preview.4 | 2842.14 | 1330.67 | 2.14x | 92.8% | 20.5% |
| [linux-x64](benchmark-linux-x64.md) | Warm | .NET 10.0.14 | 2.0.0-preview.4 | 5874.63 | 3211.41 | 1.83x | 93.5% | 18.8% |
| [osx-arm64](benchmark-osx-arm64.md) | Cold | .NET 10.0.14 | 2.0.0-preview.4 | 2997.86 | 1131.58 | 2.65x | 92.5% | 30.2% |
| [osx-arm64](benchmark-osx-arm64.md) | Warm | .NET 10.0.14 | 2.0.0-preview.4 | 4913.64 | 3282.49 | 1.5x | 93.5% | 30.2% |
| [osx-x64](benchmark-osx-x64.md) | Cold | .NET 10.0.14 | 2.0.0-preview.4 | 6255.98 | 2583.26 | 2.42x | 92.8% | 24% |
| [osx-x64](benchmark-osx-x64.md) | Warm | .NET 10.0.14 | 2.0.0-preview.4 | 15069.95 | 7729.16 | 1.95x | 93.5% | 27.5% |
| [win-arm64](benchmark-win-arm64.md) | Cold | .NET 10.0.14 | 2.0.0-preview.4 | 2953.73 | 1092.32 | 2.7x | 92.8% | 25.1% |
| [win-arm64](benchmark-win-arm64.md) | Warm | .NET 10.0.14 | 2.0.0-preview.4 | 6650.61 | 3018.14 | 2.2x | 93.5% | 25.1% |
| [win-x64](benchmark-win-x64.md) | Cold | .NET 10.0.14 | 2.0.0-preview.4 | 3412.7 | 1527.4 | 2.23x | 92.8% | 8.7% |
| [win-x64](benchmark-win-x64.md) | Warm | .NET 10.0.14 | 2.0.0-preview.4 | 6891.14 | 3663.39 | 1.88x | 93.5% | 9% |

Managed allocation excludes allocations made inside Rust. Each linked report includes peak process memory and environment metadata; the adjacent JSON contains every raw iteration and input hash.
