# Cross-platform benchmark results

Last updated (UTC): 2026-11-09 11:32:16

Each platform validates every returned row and cell before timing equivalent queries in alternating fresh processes.

| RID | Scenario | .NET runtime | MiniExcel | MiniExcel (ms) | MiniExcelRust (ms) | Speedup | Allocation reduction | Working-set reduction |
| --- | --- | --- | --- | ---: | ---: | ---: | ---: | ---: |
| [linux-arm64](benchmark-linux-arm64.md) | Cold | .NET 10.0.14 | 2.0.0-preview.4 | 3501.67 | 1059.57 | 3.3x | 92.7% | 12.8% |
| [linux-arm64](benchmark-linux-arm64.md) | Warm | .NET 10.0.14 | 2.0.0-preview.4 | 8106.43 | 2775.82 | 2.92x | 93.5% | 11.9% |
| [linux-x64](benchmark-linux-x64.md) | Cold | .NET 10.0.14 | 2.0.0-preview.4 | 2371.25 | 1097.49 | 2.16x | 92.4% | 20.4% |
| [linux-x64](benchmark-linux-x64.md) | Warm | .NET 10.0.14 | 2.0.0-preview.4 | 4533.5 | 2816.59 | 1.61x | 93.5% | 18.9% |
| [osx-arm64](benchmark-osx-arm64.md) | Cold | .NET 10.0.14 | 2.0.0-preview.4 | 2964.03 | 1009.93 | 2.93x | 92.4% | 30.5% |
| [osx-arm64](benchmark-osx-arm64.md) | Warm | .NET 10.0.14 | 2.0.0-preview.4 | 4214.26 | 2780.46 | 1.52x | 93.5% | 30.6% |
| [osx-x64](benchmark-osx-x64.md) | Cold | .NET 10.0.14 | 2.0.0-preview.4 | 6842.79 | 3367.88 | 2.03x | 93% | 24.8% |
| [osx-x64](benchmark-osx-x64.md) | Warm | .NET 10.0.14 | 2.0.0-preview.4 | 19819.13 | 7578.01 | 2.62x | 93.5% | 25.5% |
| [win-arm64](benchmark-win-arm64.md) | Cold | .NET 10.0.14 | 2.0.0-preview.4 | 2898.45 | 1096.3 | 2.64x | 92.8% | 24.9% |
| [win-arm64](benchmark-win-arm64.md) | Warm | .NET 10.0.14 | 2.0.0-preview.4 | 6603.97 | 2988.57 | 2.21x | 93.5% | 24.5% |
| [win-x64](benchmark-win-x64.md) | Cold | .NET 10.0.14 | 2.0.0-preview.4 | 2690.51 | 1492.2 | 1.8x | 92.8% | 19% |
| [win-x64](benchmark-win-x64.md) | Warm | .NET 10.0.14 | 2.0.0-preview.4 | 5194.71 | 4189.04 | 1.24x | 93.5% | 18.3% |

Managed allocation excludes allocations made inside Rust. Each linked report includes peak process memory and environment metadata; the adjacent JSON contains every raw iteration and input hash.
