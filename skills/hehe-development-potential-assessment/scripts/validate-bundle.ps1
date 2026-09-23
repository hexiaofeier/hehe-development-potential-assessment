param()

$ErrorActionPreference = 'Stop'
$skillDir = Split-Path -Parent $PSScriptRoot
$manifestPath = Join-Path $skillDir 'package-manifest.json'
$manifest = Get-Content -LiteralPath $manifestPath -Raw -Encoding UTF8 | ConvertFrom-Json
$missing = @()

foreach ($relative in @($manifest.required_local_files)) {
    $path = Join-Path $skillDir ([string]$relative).Replace('/', [IO.Path]::DirectorySeparatorChar)
    if (-not (Test-Path -LiteralPath $path)) {
        $missing += [string]$relative
    }
}

$skillFile = Join-Path $skillDir 'SKILL.md'
$nameMatches = $false
if (Test-Path -LiteralPath $skillFile) {
    $nameMatches = [bool](Get-Content -LiteralPath $skillFile -Encoding UTF8 -TotalCount 12 |
        Where-Object { $_ -eq "name: $($manifest.entry_skill)" })
}

Write-Output "Entry=$($manifest.entry_skill)"
Write-Output "InstallPolicy=$($manifest.install_policy)"
Write-Output "LocalFiles=$(@($manifest.required_local_files).Count - $missing.Count)/$(@($manifest.required_local_files).Count)"
Write-Output "Missing=$($missing -join ',')"

if (-not $nameMatches -or $missing.Count) {
    exit 1
}

Write-Output 'Standalone package validation passed.'
