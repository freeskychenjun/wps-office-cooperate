# Publish wps-office-cooperate: copy ONLY the distributable files to the deploy target.
# Dev-only stuff (tests/, docs/, tools/, DESIGN-word.md, built exe) never ships.
#
#   .\publish.ps1                  -> deploy to $DefaultTarget
#   .\publish.ps1 -Target <dir>    -> deploy elsewhere (zip it yourself afterwards)
#
# Pure ASCII source on purpose (PS 5.1 parses it regardless of code page).

param(
    [string]$Target = "C:\Users\Administrator\.agents\skills\wps-office-cooperate"
)

$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $MyInvocation.MyCommand.Path

$includes = @(
    'SKILL.md',
    'README.md',
    'reference\encryption.md',
    'reference\word-quirks.md',
    'reference\excel-quirks.md',
    'reference\workflows.md',
    'reference\word-acceptance.md',
    'reference\excel-acceptance.md',
    'scripts\office.ps1',
    'scripts\word.ps1',
    'scripts\excel.ps1',
    'scripts\ppt.ps1'
)
# All helper sources ship (each wrapper compiles every *.cs under its own helper subdir).
$includes += @(Get-ChildItem (Join-Path $root 'scripts\helper') -Recurse -Filter *.cs | ForEach-Object { 'scripts\helper\' + $_.Directory.Name + '\' + $_.Name })

# Safety: refuse obviously wrong targets.
if (-not $Target -or $Target.Trim().Length -lt 10) { throw "Suspicious target: '$Target'" }

New-Item -ItemType Directory -Force -Path $Target | Out-Null
foreach ($rel in $includes) {
    $src = Join-Path $root $rel
    if (-not (Test-Path -LiteralPath $src)) { throw "Missing expected file: $src" }
    $dst = Join-Path $Target $rel
    $dstDir = Split-Path -Parent $dst
    if (-not (Test-Path -LiteralPath $dstDir)) { New-Item -ItemType Directory -Force -Path $dstDir | Out-Null }
    Copy-Item -LiteralPath $src -Destination $dst -Force
}

# Clean files in the target that are no longer part of the package (previous build junk etc.),
# but never touch anything outside $Target.
$shipped = $includes | ForEach-Object { $_.ToLowerInvariant() }
Get-ChildItem -LiteralPath $Target -Recurse -File | ForEach-Object {
    $rel = $_.FullName.Substring($Target.Length).TrimStart('\').ToLowerInvariant()
    if ($shipped -notcontains $rel) {
        Remove-Item -LiteralPath $_.FullName -Force
        Write-Output ("removed stale: " + $rel)
    }
}

$copied = Get-ChildItem -LiteralPath $Target -Recurse -File
$sum = ($copied | Measure-Object Length -Sum).Sum
Write-Output ("published " + $copied.Count + " files, " + [math]::Round($sum / 1KB) + " KB -> " + $Target)
