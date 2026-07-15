[CmdletBinding()]
param(
    [string]$Version,
    [string]$OutputDirectory = 'dist',
    [string]$CertificateThumbprint,
    [string]$TimestampServer
)

$ErrorActionPreference = 'Stop'
$workspace = Split-Path -Parent $PSScriptRoot
Push-Location $workspace
try {
    if (-not $Version) { $Version = (Get-Content Cargo.toml | Select-String '^version\s*=\s*"([^"]+)"').Matches[0].Groups[1].Value }
    if ($Version -notmatch '^\d+\.\d+\.\d+([-.+].*)?$') { throw "Version is not semver: $Version" }
    $shell = (Get-Process -Id $PID).Path
    & $shell -NoProfile -File scripts/build-minimal-xll.ps1 -Profile release
    if ($LASTEXITCODE -ne 0) { throw "build failed: $LASTEXITCODE" }
    & $shell -NoProfile -File scripts/inspect-minimal-xll-exports.ps1
    if ($LASTEXITCODE -ne 0) { throw "export inspection failed: $LASTEXITCODE" }
    $target = Join-Path $workspace "target/release"
    $name = "excel-api-$Version-windows-x64"
    $package = Join-Path $workspace "$OutputDirectory/$name"
    New-Item -ItemType Directory -Force -Path $package | Out-Null
    Copy-Item "$target/minimal_xll.xll" "$package/minimal_xll.xll" -Force
    Copy-Item "$target/minimal_xll.exports.txt" "$package/exports.txt" -Force
    Copy-Item NOTICES.txt $package -Force
    $sha = (git rev-parse HEAD).Trim()
    $manifest = [ordered]@{ version=$Version; git_sha=$sha; target='x86_64-pc-windows-msvc'; rust=(rustc --version); features='default'; sdk='tools/Excel2013XLLSDK'; files=@{} }
    Get-ChildItem $package -File | ForEach-Object { $manifest.files[$_.Name] = (Get-FileHash $_.FullName -Algorithm SHA256).Hash.ToLowerInvariant() }
    $manifest | ConvertTo-Json -Depth 4 | Set-Content "$package/manifest.json" -Encoding utf8
    if ($CertificateThumbprint) {
        if (-not (Get-Command signtool.exe -ErrorAction SilentlyContinue)) { throw 'signtool.exe is required for signing' }
        $args = @('sign','/sha1',$CertificateThumbprint)
        if ($TimestampServer) { $args += @('/tr',$TimestampServer,'/td','sha256') }
        $args += "$package/minimal_xll.xll"
        & signtool.exe @args
        if ($LASTEXITCODE -ne 0) { throw "signing failed: $LASTEXITCODE" }
        & signtool.exe verify /pa /all "$package/minimal_xll.xll"
        if ($LASTEXITCODE -ne 0) { throw "signature verification failed: $LASTEXITCODE" }
    }
    Remove-Item -LiteralPath "$package/SHA256SUMS" -Force -ErrorAction SilentlyContinue
    Get-ChildItem $package -File | Get-FileHash -Algorithm SHA256 | ForEach-Object { "$($_.Hash.ToLowerInvariant()) *$($_.Path | Split-Path -Leaf)" } | Set-Content "$package/SHA256SUMS" -Encoding ascii
    Write-Output $package
} finally { Pop-Location }
