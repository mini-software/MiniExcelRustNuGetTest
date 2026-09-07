# Cross-platform benchmark results

Last updated (UTC): 2026-09-07 08:23:09

Each platform validates every returned row and cell before timing equivalent queries in alternating fresh processes.

| RID | Scenario | .NET runtime | MiniExcel | MiniExcel (ms) | MiniExcelRust (ms) | Speedup | Allocation reduction | Working-set reduction |
| --- | --- | --- | --- | ---: | ---: | ---: | ---: | ---: |
| [linux-arm64](benchmark-linux-arm64.md) | Cold | .NET 10.0.11 | 2.0.0-preview.4 | 3524.58 | 1052.69 | 3.35x | 92.7% | 12.5% |
| [linux-arm64](benchmark-linux-arm64.md) | Warm | .NET 10.0.11 | 2.0.0-preview.4 | 8041.09 | 2768.08 | 2.9x | 93.5% | 12.2% |
| [linux-x64](benchmark-linux-x64.md) | Cold | .NET 10.0.11 | 2.0.0-preview.4 | 2559.43 | 958.84 | 2.67x | 92.4% | 11.1% |
| [linux-x64](benchmark-linux-x64.md) | Warm | .NET 10.0.11 | 2.0.0-preview.4 | 4754.54 | 2216.26 | 2.15x | 93.5% | 8.9% |
| [osx-arm64](benchmark-osx-arm64.md) | Cold | .NET 10.0.11 | 2.0.0-preview.4 | 2583.35 | 943 | 2.74x | 92.1% | 31% |
| [osx-arm64](benchmark-osx-arm64.md) | Warm | .NET 10.0.11 | 2.0.0-preview.4 | 3868.17 | 2348.76 | 1.65x | 93.5% | 30.1% |
| [osx-x64](benchmark-osx-x64.md) | Cold | .NET 10.0.11 | 2.0.0-preview.4 | 6952.98 | 4133.51 | 1.68x | 93% | 26.2% |
| [osx-x64](benchmark-osx-x64.md) | Warm | .NET 10.0.11 | 2.0.0-preview.4 | 17775.09 | 8440.59 | 2.11x | 93.5% | 22.6% |
| [win-arm64](benchmark-win-arm64.md) | Cold | .NET 10.0.11 | 2.0.0-preview.4 | 3134.51 | 1102.53 | 2.84x | 92.6% | 11.1% |
| [win-arm64](benchmark-win-arm64.md) | Warm | .NET 10.0.11 | 2.0.0-preview.4 | 6597.02 | 2934.18 | 2.25x | 93.5% | 10.3% |
| [win-x64](benchmark-win-x64.md) | Cold | .NET 10.0.11 | 2.0.0-preview.4 | 2753.18 | 1513.59 | 1.82x | 92.7% | 22.8% |
| [win-x64](benchmark-win-x64.md) | Warm | .NET 10.0.11 | 2.0.0-preview.4 | 5210.33 | 4068.29 | 1.28x | 93.5% | 19.6% |

Managed allocation excludes allocations made inside Rust. Each linked report includes peak process memory and environment metadata; the adjacent JSON contains every raw iteration and input hash.
