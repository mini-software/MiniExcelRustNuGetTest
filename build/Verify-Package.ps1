[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string] $PackagePath
)

$ErrorActionPreference = 'Stop'
$expectedEntries = @(
    'runtimes/win-x64/native/miniexcel_ffi.dll',
    'runtimes/win-arm64/native/miniexcel_ffi.dll',
    'runtimes/linux-x64/native/libminiexcel_ffi.so',
    'runtimes/linux-arm64/native/libminiexcel_ffi.so',
    'runtimes/linux-musl-x64/native/libminiexcel_ffi.so',
    'runtimes/linux-musl-arm64/native/libminiexcel_ffi.so',
    'runtimes/osx-x64/native/libminiexcel_ffi.dylib',
    'runtimes/osx-arm64/native/libminiexcel_ffi.dylib'
)

Add-Type -AssemblyName System.IO.Compression.FileSystem
$archive = [System.IO.Compression.ZipFile]::OpenRead((Resolve-Path $PackagePath))
try {
    $actualEntries = @($archive.Entries | ForEach-Object FullName)
    foreach ($expectedEntry in $expectedEntries) {
        if ($expectedEntry -notin $actualEntries) {
            throw "The package is missing $expectedEntry."
        }
    }

    $unexpectedNativeEntries = @($actualEntries | Where-Object {
        $_ -like 'runtimes/*/native/*' -and $_ -notin $expectedEntries
    })
    if ($unexpectedNativeEntries.Count -ne 0) {
        throw "The package contains unexpected native assets: $($unexpectedNativeEntries -join ', ')."
    }
}
finally {
    $archive.Dispose()
}

Write-Host "Verified all $($expectedEntries.Count) native package assets."
