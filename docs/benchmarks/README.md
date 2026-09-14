# Cross-platform benchmark results

Last updated (UTC): 2026-09-14 08:58:18

Each platform validates every returned row and cell before timing equivalent queries in alternating fresh processes.

| RID | Scenario | .NET runtime | MiniExcel | MiniExcel (ms) | MiniExcelRust (ms) | Speedup | Allocation reduction | Working-set reduction |
| --- | --- | --- | --- | ---: | ---: | ---: | ---: | ---: |
| [linux-arm64](benchmark-linux-arm64.md) | Cold | .NET 10.0.12 | 2.0.0-preview.4 | 3553.09 | 1056.78 | 3.36x | 92.7% | 12.7% |
| [linux-arm64](benchmark-linux-arm64.md) | Warm | .NET 10.0.12 | 2.0.0-preview.4 | 8110.62 | 2792.15 | 2.9x | 93.5% | 12.2% |
| [linux-x64](benchmark-linux-x64.md) | Cold | .NET 10.0.12 | 2.0.0-preview.4 | 2801.97 | 1248.99 | 2.24x | 92.6% | 20.4% |
| [linux-x64](benchmark-linux-x64.md) | Warm | .NET 10.0.12 | 2.0.0-preview.4 | 5784.07 | 3225.5 | 1.79x | 93.5% | 18.6% |
| [osx-arm64](benchmark-osx-arm64.md) | Cold | .NET 10.0.12 | 2.0.0-preview.4 | 2867.89 | 1141.01 | 2.51x | 92.4% | 31.8% |
| [osx-arm64](benchmark-osx-arm64.md) | Warm | .NET 10.0.12 | 2.0.0-preview.4 | 4859.98 | 2991.35 | 1.62x | 93.5% | 30.6% |
| [osx-x64](benchmark-osx-x64.md) | Cold | .NET 10.0.12 | 2.0.0-preview.4 | 5246.29 | 2316.14 | 2.27x | 92.9% | 24.9% |
| [osx-x64](benchmark-osx-x64.md) | Warm | .NET 10.0.12 | 2.0.0-preview.4 | 12203.88 | 6352.62 | 1.92x | 93.5% | 27.7% |
| [win-arm64](benchmark-win-arm64.md) | Cold | .NET 10.0.12 | 2.0.0-preview.4 | 3038.7 | 1080.11 | 2.81x | 92.8% | 9.8% |
| [win-arm64](benchmark-win-arm64.md) | Warm | .NET 10.0.12 | 2.0.0-preview.4 | 6587.04 | 2947.6 | 2.23x | 93.5% | 9.3% |
| [win-x64](benchmark-win-x64.md) | Cold | .NET 10.0.12 | 2.0.0-preview.4 | 2803.49 | 1528.12 | 1.83x | 92.8% | 18.4% |
| [win-x64](benchmark-win-x64.md) | Warm | .NET 10.0.12 | 2.0.0-preview.4 | 5267.52 | 4291.1 | 1.23x | 93.5% | 18.5% |

Managed allocation excludes allocations made inside Rust. Each linked report includes peak process memory and environment metadata; the adjacent JSON contains every raw iteration and input hash.
