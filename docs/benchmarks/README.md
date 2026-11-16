# Cross-platform benchmark results

Last updated (UTC): 2026-11-16 11:17:21

Each platform validates every returned row and cell before timing equivalent queries in alternating fresh processes.

| RID | Scenario | .NET runtime | MiniExcel | MiniExcel (ms) | MiniExcelRust (ms) | Speedup | Allocation reduction | Working-set reduction |
| --- | --- | --- | --- | ---: | ---: | ---: | ---: | ---: |
| [linux-arm64](benchmark-linux-arm64.md) | Cold | .NET 10.0.15 | 2.0.0-preview.4 | 3567.03 | 1051.44 | 3.39x | 92.7% | 12.4% |
| [linux-arm64](benchmark-linux-arm64.md) | Warm | .NET 10.0.15 | 2.0.0-preview.4 | 8168.68 | 2748.68 | 2.97x | 93.5% | 12.5% |
| [linux-x64](benchmark-linux-x64.md) | Cold | .NET 10.0.15 | 2.0.0-preview.4 | 2182.83 | 790.43 | 2.76x | 92% | 20% |
| [linux-x64](benchmark-linux-x64.md) | Warm | .NET 10.0.15 | 2.0.0-preview.4 | 4095.99 | 1931.4 | 2.12x | 93.5% | 18.2% |
| [osx-arm64](benchmark-osx-arm64.md) | Cold | .NET 10.0.15 | 2.0.0-preview.4 | 2831.22 | 1272.47 | 2.22x | 92.5% | 30.4% |
| [osx-arm64](benchmark-osx-arm64.md) | Warm | .NET 10.0.15 | 2.0.0-preview.4 | 5241.17 | 3638.15 | 1.44x | 93.5% | 31.2% |
| [osx-x64](benchmark-osx-x64.md) | Cold | .NET 10.0.15 | 2.0.0-preview.4 | 7219.99 | 2772.03 | 2.6x | 93% | 25% |
| [osx-x64](benchmark-osx-x64.md) | Warm | .NET 10.0.15 | 2.0.0-preview.4 | 15080.09 | 7912.44 | 1.91x | 93.5% | 23.6% |
| [win-arm64](benchmark-win-arm64.md) | Cold | .NET 10.0.15 | 2.0.0-preview.4 | 2897.79 | 1085.57 | 2.67x | 92.8% | 24.7% |
| [win-arm64](benchmark-win-arm64.md) | Warm | .NET 10.0.15 | 2.0.0-preview.4 | 6670.97 | 3022.06 | 2.21x | 93.5% | 24.1% |
| [win-x64](benchmark-win-x64.md) | Cold | .NET 10.0.15 | 2.0.0-preview.4 | 2762.31 | 1572.84 | 1.76x | 92.9% | 18.5% |
| [win-x64](benchmark-win-x64.md) | Warm | .NET 10.0.15 | 2.0.0-preview.4 | 5387.02 | 4312.56 | 1.25x | 93.5% | 18.7% |

Managed allocation excludes allocations made inside Rust. Each linked report includes peak process memory and environment metadata; the adjacent JSON contains every raw iteration and input hash.
