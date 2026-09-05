[CmdletBinding()]
param(
    [string] $InputDirectory,

    [string] $OutputDirectory
)

$ErrorActionPreference = 'Stop'
$repositoryRoot = Split-Path $PSScriptRoot -Parent
if ([string]::IsNullOrWhiteSpace($InputDirectory)) {
    $InputDirectory = Join-Path $repositoryRoot 'artifacts/benchmarks'
}
if ([string]::IsNullOrWhiteSpace($OutputDirectory)) {
    $OutputDirectory = Join-Path $repositoryRoot 'docs/benchmarks'
}

$InputDirectory = [IO.Path]::GetFullPath($InputDirectory)
$OutputDirectory = [IO.Path]::GetFullPath($OutputDirectory)
$resultFiles = @(Get-ChildItem $InputDirectory -Filter 'benchmark-*.json' -File | Sort-Object Name)
if ($resultFiles.Count -eq 0) {
    throw "No benchmark JSON files found in $InputDirectory."
}

New-Item -ItemType Directory -Path $OutputDirectory -Force | Out-Null
$published = foreach ($resultFile in $resultFiles) {
    $result = Get-Content $resultFile.FullName -Raw | ConvertFrom-Json
    $rid = $result.Metadata.Rid
    if ([string]::IsNullOrWhiteSpace($rid)) {
        throw "$($resultFile.Name) does not contain Metadata.Rid."
    }

    $markdownSource = Join-Path $InputDirectory "benchmark-$rid.md"
    if (-not (Test-Path $markdownSource)) {
        throw "Missing Markdown report for ${rid}: $markdownSource"
    }

    Copy-Item $resultFile.FullName (Join-Path $OutputDirectory $resultFile.Name) -Force
    Copy-Item $markdownSource (Join-Path $OutputDirectory "benchmark-$rid.md") -Force

    foreach ($scenario in @('Cold', 'Warm')) {
        $baseline = $result.Summary | Where-Object { $_.Scenario -eq $scenario -and $_.Runtime -eq 'MiniExcelV2' }
        $candidate = $result.Summary | Where-Object { $_.Scenario -eq $scenario -and $_.Runtime -eq 'MiniExcelRust' }
        if ($null -eq $baseline -or $null -eq $candidate) {
            throw "$($resultFile.Name) is missing the $scenario summary."
        }

        [pscustomobject]@{
            Rid = $rid
            Scenario = $scenario
            DotNetRuntime = $result.Metadata.DotNetRuntime
            MiniExcelVersion = $result.Metadata.MiniExcelVersion
            TimestampUtc = ([DateTimeOffset]$result.Metadata.TimestampUtc).ToUniversalTime()
            BaselineElapsedMs = $baseline.ElapsedMs
            CandidateElapsedMs = $candidate.ElapsedMs
            Speedup = $candidate.Speedup
            AllocationReductionPercent = $candidate.AllocationReductionPercent
            WorkingSetReductionPercent = $candidate.WorkingSetReductionPercent
        }
    }
}

$latestTimestamp = ($published.TimestampUtc | Sort-Object -Descending | Select-Object -First 1)
$index = [Collections.Generic.List[string]]::new()
$index.Add('# Cross-platform benchmark results')
$index.Add('')
$index.Add("Last updated (UTC): $($latestTimestamp.ToString('yyyy-MM-dd HH:mm:ss'))")
$index.Add('')
$index.Add('Each platform validates every returned row and cell before timing equivalent queries in alternating fresh processes.')
$index.Add('')
$index.Add('| RID | Scenario | .NET runtime | MiniExcel | MiniExcel (ms) | MiniExcelRust (ms) | Speedup | Allocation reduction | Working-set reduction |')
$index.Add('| --- | --- | --- | --- | ---: | ---: | ---: | ---: | ---: |')
foreach ($row in $published | Sort-Object Rid, Scenario) {
    $index.Add("| [$($row.Rid)](benchmark-$($row.Rid).md) | $($row.Scenario) | $($row.DotNetRuntime) | $($row.MiniExcelVersion) | $($row.BaselineElapsedMs) | $($row.CandidateElapsedMs) | $($row.Speedup)x | $($row.AllocationReductionPercent)% | $($row.WorkingSetReductionPercent)% |")
}
$index.Add('')
$index.Add('Managed allocation excludes allocations made inside Rust. Each linked report includes peak process memory and environment metadata; the adjacent JSON contains every raw iteration and input hash.')
$index | Set-Content (Join-Path $OutputDirectory 'README.md')

Write-Host "Published $($resultFiles.Count) platform result set(s) to $OutputDirectory."