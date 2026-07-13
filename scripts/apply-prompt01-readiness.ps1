$ErrorActionPreference = "Stop"
$repoRoot = Split-Path -Parent $PSScriptRoot
Set-Location $repoRoot
$obsolete = @("PACKAGE_README.md", "MANIFEST.md", "BOOK_REVISION_NOTES.md")
foreach ($path in $obsolete) {
    if (Test-Path $path) {
        Remove-Item -Force $path
        Write-Host "Removed obsolete file: $path"
    }
}
Write-Host "Prompt 01 readiness overlay applied. Review with git status and git diff."
