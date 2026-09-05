[CmdletBinding()]
param(
    [ValidateSet('win-x64', 'win-arm64', 'linux-x64', 'linux-arm64', 'linux-musl-x64', 'linux-musl-arm64', 'osx-x64', 'osx-arm64')]
    [string] $Rid = 'win-x64',

    [string] $Version = '0.1.0-preview.2',

    [string] $MiniExcelVersion = '2.0.0-preview.4',

    [switch] $SkipNativeBuild,

    [ValidateRange(100, 1000000)]
    [int] $LifecycleIterations = 1000,

    [ValidateRange(1, 1024)]
    [int] $MaxPrivateGrowthMb = 32
)

$ErrorActionPreference = 'Stop'
$repositoryRoot = Split-Path $PSScriptRoot -Parent
$packageDirectory = Join-Path $repositoryRoot 'artifacts/packages'
$consumerProject = Join-Path $repositoryRoot 'tests/MiniExcelRust.PackageTests/MiniExcelRust.PackageTests.csproj'

if (-not $SkipNativeBuild) {
    & (Join-Path $PSScriptRoot 'Build-Native.ps1') -Rid $Rid
}

New-Item -ItemType Directory -Path $packageDirectory -Force | Out-Null
& dotnet pack (Join-Path $repositoryRoot 'src/MiniExcelRust/MiniExcelRust.csproj') `
    -c Release `
    -o $packageDirectory `
    -p:PackageVersion=$Version `
    -p:MiniExcelRustRequireAllNativeAssets=false
if ($LASTEXITCODE -ne 0) {
    throw 'NuGet pack failed.'
}

$package = Join-Path $packageDirectory "MiniExcelRust.$Version.nupkg"
Add-Type -AssemblyName System.IO.Compression.FileSystem
$archive = [System.IO.Compression.ZipFile]::OpenRead($package)
try {
    $nativeEntry = $archive.Entries | Where-Object { $_.FullName -like "runtimes/$Rid/native/*" }
    if ($null -eq $nativeEntry) {
        throw "The package does not contain a native asset for $Rid."
    }
}
finally {
    $archive.Dispose()
}

& dotnet restore $consumerProject `
    --force `
    --source $packageDirectory `
    -p:MiniExcelRustPackageVersion=$Version `
    -p:MiniExcelVersion=$MiniExcelVersion
if ($LASTEXITCODE -ne 0) {
    throw 'Package consumer restore failed.'
}

& dotnet run --project $consumerProject -c Release --no-restore `
    -p:MiniExcelRustPackageVersion=$Version `
    -p:MiniExcelVersion=$MiniExcelVersion `
    -- suite $LifecycleIterations $MaxPrivateGrowthMb
if ($LASTEXITCODE -ne 0) {
    throw 'Package consumer smoke test failed.'
}
