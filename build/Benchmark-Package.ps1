[CmdletBinding()]
param(
    [ValidateSet('win-x64', 'win-arm64', 'linux-x64', 'linux-arm64', 'osx-x64', 'osx-arm64')]
    [string] $Rid = 'win-x64',

    [ValidateRange(1, 10)]
    [int] $Iterations = 5,

    [ValidateRange(1, 1000000)]
    [int] $Rows = 100000,

    [ValidateRange(1, 26)]
    [int] $Columns = 10,

    [string] $Version = '0.1.0-preview.2',

    [string] $MiniExcelVersion = '2.0.0-preview.4',

    [ValidateSet('net8.0', 'net9.0', 'net10.0')]
    [string] $TargetFramework = 'net10.0',

    [switch] $SkipPackageTest,

    [string] $OutputDirectory
)

$ErrorActionPreference = 'Stop'
$repositoryRoot = Split-Path $PSScriptRoot -Parent
if ([string]::IsNullOrWhiteSpace($OutputDirectory)) {
    $OutputDirectory = Join-Path $repositoryRoot 'artifacts/benchmarks'
}
$OutputDirectory = [IO.Path]::GetFullPath($OutputDirectory)

$consumerProject = Join-Path $repositoryRoot 'tests/MiniExcelRust.PackageTests/MiniExcelRust.PackageTests.csproj'
$consumerDll = Join-Path $repositoryRoot "tests/MiniExcelRust.PackageTests/bin/Release/$TargetFramework/MiniExcelRust.PackageTests.dll"
$workbook = Join-Path $OutputDirectory "benchmark-$Rows`x$Columns.xlsx"
$package = Join-Path $repositoryRoot "artifacts/packages/MiniExcelRust.$Version.nupkg"
$scenarios = @(
    [pscustomobject]@{ Name = 'Cold'; Passes = 1; WarmupPasses = 0 }
    [pscustomobject]@{ Name = 'Warm'; Passes = 3; WarmupPasses = 1 }
)

New-Item -ItemType Directory -Path $OutputDirectory -Force | Out-Null
if (-not $SkipPackageTest) {
    & (Join-Path $PSScriptRoot 'Test-Package.ps1') `
        -Rid $Rid `
        -Version $Version `
        -MiniExcelVersion $MiniExcelVersion
    if ($LASTEXITCODE -ne 0) {
        throw 'Package build and lifecycle suite failed.'
    }
}

if (-not (Test-Path $package)) {
    throw "Package not found: $package"
}

& dotnet restore $consumerProject `
    --force `
    --source (Split-Path $package -Parent) `
    -p:MiniExcelRustPackageVersion=$Version `
    -p:MiniExcelVersion=$MiniExcelVersion `
    -p:TargetFramework=$TargetFramework
if ($LASTEXITCODE -ne 0) {
    throw 'Benchmark consumer restore failed.'
}

& dotnet build $consumerProject `
    -c Release `
    --no-restore `
    -p:MiniExcelRustPackageVersion=$Version `
    -p:MiniExcelVersion=$MiniExcelVersion `
    -p:TargetFramework=$TargetFramework
if ($LASTEXITCODE -ne 0) {
    throw 'Benchmark consumer build failed.'
}

& dotnet $consumerDll generate $workbook $Rows $Columns
if ($LASTEXITCODE -ne 0) {
    throw 'Benchmark workbook generation failed.'
}
& dotnet $consumerDll verify $workbook false
if ($LASTEXITCODE -ne 0) {
    throw 'Managed and Rust benchmark parity failed.'
}

function Invoke-BenchmarkProcess {
    param(
        [string] $Runtime,
        [pscustomobject] $Scenario,
        [int] $Iteration
    )

    $startInfo = [Diagnostics.ProcessStartInfo]::new()
    $startInfo.FileName = 'dotnet'
    $startInfo.UseShellExecute = $false
    $startInfo.RedirectStandardOutput = $true
    $startInfo.RedirectStandardError = $true
    foreach ($argument in @($consumerDll, $Runtime, $workbook, "$($Scenario.Passes)", "$($Scenario.WarmupPasses)")) {
        $startInfo.ArgumentList.Add($argument)
    }

    $process = [Diagnostics.Process]::Start($startInfo)
    $peakWorkingSet = 0L
    $peakPrivateBytes = 0L
    while (-not $process.WaitForExit(10)) {
        $process.Refresh()
        $peakWorkingSet = [Math]::Max($peakWorkingSet, $process.WorkingSet64)
        $peakPrivateBytes = [Math]::Max($peakPrivateBytes, $process.PrivateMemorySize64)
    }

    $output = $process.StandardOutput.ReadToEnd().Trim()
    $errorOutput = $process.StandardError.ReadToEnd().Trim()
    if ($process.ExitCode -ne 0) {
        throw "$Runtime $($Scenario.Name) benchmark failed: $errorOutput"
    }

    $measurement = $output | ConvertFrom-Json
    [pscustomobject]@{
        Scenario = $Scenario.Name
        Runtime = $measurement.Runtime
        DotNetRuntime = $measurement.DotNetRuntime
        Iteration = $Iteration
        Passes = $measurement.Passes
        Rows = $measurement.Rows
        Cells = $measurement.Cells
        ElapsedMs = [Math]::Round($measurement.ElapsedMilliseconds, 2)
        FirstRowMs = [Math]::Round($measurement.FirstRowMilliseconds, 2)
        AllocatedMB = [Math]::Round($measurement.AllocatedBytes / 1MB, 2)
        PeakWorkingSetMB = [Math]::Round($peakWorkingSet / 1MB, 2)
        PeakPrivateMB = [Math]::Round($peakPrivateBytes / 1MB, 2)
    }
}

$results = [Collections.Generic.List[object]]::new()
foreach ($scenario in $scenarios) {
    foreach ($iteration in 1..$Iterations) {
        $order = if ($iteration % 2 -eq 1) { @('managed', 'rust') } else { @('rust', 'managed') }
        foreach ($runtime in $order) {
            $results.Add((Invoke-BenchmarkProcess -Runtime $runtime -Scenario $scenario -Iteration $iteration))
        }
    }
}

foreach ($scenario in $scenarios) {
    $scenarioResults = @($results | Where-Object Scenario -eq $scenario.Name)
    if (($scenarioResults.Rows | Select-Object -Unique).Count -ne 1 -or
        ($scenarioResults.Cells | Select-Object -Unique).Count -ne 1) {
        throw "$($scenario.Name): managed and Rust runners returned different row or cell counts."
    }
}

$summaries = foreach ($scenario in $scenarios) {
    $groups = @($results | Where-Object Scenario -eq $scenario.Name | Group-Object Runtime)
    $managed = $groups | Where-Object Name -eq 'MiniExcelV2'
    $rust = $groups | Where-Object Name -eq 'MiniExcelRust'
    $managedElapsed = ($managed.Group.ElapsedMs | Measure-Object -Average).Average
    $rustElapsed = ($rust.Group.ElapsedMs | Measure-Object -Average).Average
    $managedAllocated = ($managed.Group.AllocatedMB | Measure-Object -Average).Average
    $rustAllocated = ($rust.Group.AllocatedMB | Measure-Object -Average).Average
    $managedWorkingSet = ($managed.Group.PeakWorkingSetMB | Measure-Object -Average).Average
    $rustWorkingSet = ($rust.Group.PeakWorkingSetMB | Measure-Object -Average).Average

    [pscustomobject]@{
        Scenario = $scenario.Name
        Runtime = 'MiniExcelV2'
        ElapsedMs = [Math]::Round($managedElapsed, 2)
        FirstRowMs = [Math]::Round(($managed.Group.FirstRowMs | Measure-Object -Average).Average, 2)
        AllocatedMB = [Math]::Round($managedAllocated, 2)
        PeakWorkingSetMB = [Math]::Round($managedWorkingSet, 2)
        PeakPrivateMB = [Math]::Round(($managed.Group.PeakPrivateMB | Measure-Object -Average).Average, 2)
        Speedup = 1
        AllocationReductionPercent = 0
        WorkingSetReductionPercent = 0
    }
    [pscustomobject]@{
        Scenario = $scenario.Name
        Runtime = 'MiniExcelRust'
        ElapsedMs = [Math]::Round($rustElapsed, 2)
        FirstRowMs = [Math]::Round(($rust.Group.FirstRowMs | Measure-Object -Average).Average, 2)
        AllocatedMB = [Math]::Round($rustAllocated, 2)
        PeakWorkingSetMB = [Math]::Round($rustWorkingSet, 2)
        PeakPrivateMB = [Math]::Round(($rust.Group.PeakPrivateMB | Measure-Object -Average).Average, 2)
        Speedup = [Math]::Round($managedElapsed / $rustElapsed, 2)
        AllocationReductionPercent = [Math]::Round((1 - $rustAllocated / $managedAllocated) * 100, 1)
        WorkingSetReductionPercent = [Math]::Round((1 - $rustWorkingSet / $managedWorkingSet) * 100, 1)
    }
}

$metadata = [ordered]@{
    TimestampUtc = [DateTime]::UtcNow.ToString('O')
    OperatingSystem = [Runtime.InteropServices.RuntimeInformation]::OSDescription
    Architecture = [Runtime.InteropServices.RuntimeInformation]::OSArchitecture.ToString()
    Rid = $Rid
    DotNetSdk = (& dotnet --version).Trim()
    TargetFramework = $TargetFramework
    DotNetRuntime = (@($results.DotNetRuntime | Select-Object -Unique) -join ', ')
    MiniExcelVersion = $MiniExcelVersion
    MiniExcelRustVersion = $Version
    RustVersion = '1.85.0'
    Rows = $Rows
    Columns = $Columns
    WorkbookSha256 = (Get-FileHash $workbook -Algorithm SHA256).Hash.ToLowerInvariant()
    PackageSha256 = (Get-FileHash $package -Algorithm SHA256).Hash.ToLowerInvariant()
    Iterations = $Iterations
}

$jsonPath = Join-Path $OutputDirectory "benchmark-$Rid.json"
$markdownPath = Join-Path $OutputDirectory "benchmark-$Rid.md"
[ordered]@{ Metadata = $metadata; Results = $results; Summary = $summaries } |
    ConvertTo-Json -Depth 6 |
    Set-Content $jsonPath

$markdown = [Collections.Generic.List[string]]::new()
$markdown.Add("# MiniExcelRust benchmark ($Rid)")
$markdown.Add('')
$markdown.Add("- Date (UTC): $($metadata.TimestampUtc)")
$markdown.Add("- OS: $($metadata.OperatingSystem)")
$markdown.Add("- Architecture: $($metadata.Architecture)")
$markdown.Add("- .NET SDK: $($metadata.DotNetSdk)")
$markdown.Add("- Target framework: $($metadata.TargetFramework)")
$markdown.Add("- .NET runtime: $($metadata.DotNetRuntime)")
$markdown.Add("- Baseline: MiniExcel $($metadata.MiniExcelVersion)")
$markdown.Add("- Candidate: MiniExcelRust $($metadata.MiniExcelRustVersion)")
$markdown.Add("- Workbook: $Rows rows x $Columns columns")
$markdown.Add("- Iterations: $Iterations fresh processes per runtime and scenario")
$markdown.Add('')
$markdown.Add('| Scenario | Runtime | Elapsed (ms) | First row (ms) | Managed allocation (MB) | Peak working set (MB) | Peak private (MB) | Speedup |')
$markdown.Add('| --- | --- | ---: | ---: | ---: | ---: | ---: | ---: |')
foreach ($summary in $summaries) {
    $markdown.Add("| $($summary.Scenario) | $($summary.Runtime) | $($summary.ElapsedMs) | $($summary.FirstRowMs) | $($summary.AllocatedMB) | $($summary.PeakWorkingSetMB) | $($summary.PeakPrivateMB) | $($summary.Speedup)x |")
}
$markdown.Add('')
$markdown.Add('Managed allocation excludes allocations made inside Rust. Peak working set and peak private bytes include the complete process.')
$markdown | Set-Content $markdownPath

$results | Format-Table Scenario, Runtime, Iteration, ElapsedMs, FirstRowMs, AllocatedMB, PeakWorkingSetMB, PeakPrivateMB -AutoSize
$summaries | Format-Table Scenario, Runtime, ElapsedMs, FirstRowMs, AllocatedMB, PeakWorkingSetMB, Speedup, AllocationReductionPercent, WorkingSetReductionPercent -AutoSize
Write-Host "JSON: $jsonPath"
Write-Host "Markdown: $markdownPath"