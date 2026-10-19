# Cross-platform benchmark results

Last updated (UTC): 2026-10-19 08:55:13

Each platform validates every returned row and cell before timing equivalent queries in alternating fresh processes.

| RID | Scenario | .NET runtime | MiniExcel | MiniExcel (ms) | MiniExcelRust (ms) | Speedup | Allocation reduction | Working-set reduction |
| --- | --- | --- | --- | ---: | ---: | ---: | ---: | ---: |
| [linux-arm64](benchmark-linux-arm64.md) | Cold | .NET 10.0.13 | 2.0.0-preview.4 | 3596.66 | 1056.86 | 3.4x | 92.7% | 12.7% |
| [linux-arm64](benchmark-linux-arm64.md) | Warm | .NET 10.0.13 | 2.0.0-preview.4 | 8084.27 | 2785.07 | 2.9x | 93.5% | 12.3% |
| [linux-x64](benchmark-linux-x64.md) | Cold | .NET 10.0.13 | 2.0.0-preview.4 | 3095.75 | 1259.23 | 2.46x | 92.7% | 11.2% |
| [linux-x64](benchmark-linux-x64.md) | Warm | .NET 10.0.13 | 2.0.0-preview.4 | 6726.53 | 2959.33 | 2.27x | 93.5% | 10.7% |
| [osx-arm64](benchmark-osx-arm64.md) | Cold | .NET 10.0.13 | 2.0.0-preview.4 | 3170.19 | 1262.87 | 2.51x | 92.5% | 32% |
| [osx-arm64](benchmark-osx-arm64.md) | Warm | .NET 10.0.13 | 2.0.0-preview.4 | 4980.18 | 3217.08 | 1.55x | 93.5% | 29.6% |
| [osx-x64](benchmark-osx-x64.md) | Cold | .NET 10.0.13 | 2.0.0-preview.4 | 8604.59 | 3493.68 | 2.46x | 93% | 25.3% |
| [osx-x64](benchmark-osx-x64.md) | Warm | .NET 10.0.13 | 2.0.0-preview.4 | 18943.96 | 11782.38 | 1.61x | 93.5% | 24.7% |
| [win-arm64](benchmark-win-arm64.md) | Cold | .NET 10.0.13 | 2.0.0-preview.4 | 3008.8 | 1081.36 | 2.78x | 92.8% | 25.7% |
| [win-arm64](benchmark-win-arm64.md) | Warm | .NET 10.0.13 | 2.0.0-preview.4 | 6636.63 | 2989.89 | 2.22x | 93.5% | 25.3% |
| [win-x64](benchmark-win-x64.md) | Cold | .NET 10.0.13 | 2.0.0-preview.4 | 2808.55 | 1472.99 | 1.91x | 92.9% | 19.6% |
| [win-x64](benchmark-win-x64.md) | Warm | .NET 10.0.13 | 2.0.0-preview.4 | 5188.33 | 4099.78 | 1.27x | 93.5% | 18% |

Managed allocation excludes allocations made inside Rust. Each linked report includes peak process memory and environment metadata; the adjacent JSON contains every raw iteration and input hash.
