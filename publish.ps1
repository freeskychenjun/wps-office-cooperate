# Publish wps-office-cooperate: copy ONLY the distributable files to the deploy target.
# Dev-only stuff (tests/, docs/borrow, tools/, DESIGN*.md, OPTIMIZATION-PLAN.md, built exe) never ships.
#
#   .\publish.ps1                     -> deploy to $DefaultTarget
#   .\publish.ps1 -Target <dir>       -> deploy elsewhere (zip it yourself afterwards; needs -Force
#                                        unless the path looks like <...>\skills\<name>)
#   .\publish.ps1 -WhatIf             -> print what would be copied / removed, touch nothing
#
# Safety model (v1.34):
#   * the target must look like a skills install dir ("<...>\skills\<name>") unless -Force is given,
#     because this script recursively removes files that are not part of the package;
#   * the new package is staged next to the target and only moved into place AFTER every source file
#     was copied successfully, so a mid-way failure cannot leave a half-updated deployment;
#   * -WhatIf does a full dry run (resolve, validate, list copies and stale removals).
#
# Pure ASCII source on purpose (PS 5.1 parses it regardless of code page).

param(
    [string]$Target = "C:\Users\Administrator\.agents\skills\wps-office-cooperate",
    [switch]$Force,
    [switch]$WhatIf
)

$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $MyInvocation.MyCommand.Path

$includes = @(
    'SKILL.md',
    'README.md',
    'reference\encryption.md',
    'reference\changelog.md',
    'reference\ppt-quirks.md',
    'reference\ppt-presets.md',
    'reference\word-quirks.md',
    'reference\excel-quirks.md',
    'reference\workflows.md',
    'reference\multi-task.md',
    'reference\multi-task-excel.md',
    'reference\multi-task-ppt.md',
    'reference\multi-task-session.md',
    'reference\word-acceptance.md',
    'reference\excel-acceptance.md',
    'scripts\office.ps1',
    'scripts\woc.cmd',
    'scripts\woc.ps1',
    'scripts\param-check.ps1',
    'scripts\tasks.ps1',
    'scripts\tasks-excel.ps1',
    'scripts\tasks-ppt.ps1',
    'scripts\tasks-session.ps1',
    'scripts\woc-common.ps1',
    'scripts\woc-app-common.ps1',
    'scripts\word.ps1',
    'scripts\excel.ps1',
    'scripts\ppt.ps1'
)
# README links these screenshots; ship them or the deployed README shows broken images.
if (Test-Path -LiteralPath (Join-Path $root 'docs')) {
    $includes += @(Get-ChildItem (Join-Path $root 'docs') -Filter *.png | ForEach-Object { 'docs\' + $_.Name })
}
# All helper sources ship (each wrapper compiles every *.cs under its own helper subdir).
$includes += @(Get-ChildItem (Join-Path $root 'scripts\helper') -Recurse -Filter *.cs | ForEach-Object { 'scripts\helper\' + $_.Directory.Name + '\' + $_.Name })

# ---------------------------------------------------------------- target validation
if (-not $Target -or $Target.Trim().Length -lt 10) { throw "Suspicious target: '$Target'" }
$Target = [System.IO.Path]::GetFullPath($Target)
if ($Target -eq [System.IO.Path]::GetPathRoot($Target)) { throw "Refusing to publish into a drive root: $Target" }
if (Test-Path -LiteralPath $Target -PathType Leaf) { throw "Target is a file, not a directory: $Target" }
$parentName = Split-Path -Leaf (Split-Path -Parent $Target)
if (-not $Force -and $parentName -ne 'skills') {
    throw ("Refusing target '$Target': the parent folder is not 'skills' and this script removes files it does not ship. " +
           "Pass -Target <...\skills\<name>> or add -Force if you really mean it.")
}
$stage = $Target + '.__staging'
if (Test-Path -LiteralPath $stage) { Remove-Item -Recurse -Force -LiteralPath $stage }

Write-Output ("publish: " + $root + " -> " + $Target + $(if ($Force) { ' (forced)' } else { '' }))

# ---------------------------------------------------------------- preflight: every source must exist
$missing = @()
foreach ($rel in $includes) {
    $src = Join-Path $root $rel
    if (-not (Test-Path -LiteralPath $src)) { $missing += $src }
}
if ($missing.Count -gt 0) { throw ("Missing expected file(s):`n  " + ($missing -join "`n  ")) }

# ---------------------------------------------------------------- stage (no target writes yet)
if (-not $WhatIf) {
    New-Item -ItemType Directory -Force -Path $stage | Out-Null
    foreach ($rel in $includes) {
        $src = Join-Path $root $rel
        $dst = Join-Path $stage $rel
        $dstDir = Split-Path -Parent $dst
        if (-not (Test-Path -LiteralPath $dstDir)) { New-Item -ItemType Directory -Force -Path $dstDir | Out-Null }
        Copy-Item -LiteralPath $src -Destination $dst -Force
    }
}

# ---------------------------------------------------------------- stale files in the target
$shipped = $includes | ForEach-Object { $_.ToLowerInvariant() }
$stale = @()
if (Test-Path -LiteralPath $Target) {
    $stale = @(Get-ChildItem -LiteralPath $Target -Recurse -File | ForEach-Object {
        $rel = $_.FullName.Substring($Target.Length).TrimStart('\').ToLowerInvariant()
        if ($shipped -notcontains $rel) { $_ }
    })
}

if ($WhatIf) {
    Write-Output '--- WhatIf: copy list ---'
    foreach ($rel in $includes) { Write-Output ('  copy   ' + $rel) }
    Write-Output '--- WhatIf: stale removals ---'
    if ($stale.Count -eq 0) { Write-Output '  (none)' }
    foreach ($f in $stale) { Write-Output ('  remove ' + $f.FullName.Substring($Target.Length).TrimStart('\')) }
    Write-Output ("would publish " + $includes.Count + " files -> " + $Target)
    exit 0
}

foreach ($f in $stale) {
    Remove-Item -LiteralPath $f.FullName -Force
    Write-Output ('removed stale: ' + $f.FullName.Substring($Target.Length).TrimStart('\'))
}

# ---------------------------------------------------------------- swap in the staged package
if (-not (Test-Path -LiteralPath $Target)) { New-Item -ItemType Directory -Force -Path $Target | Out-Null }
foreach ($rel in $includes) {
    $src = Join-Path $stage $rel
    $dst = Join-Path $Target $rel
    $dstDir = Split-Path -Parent $dst
    if (-not (Test-Path -LiteralPath $dstDir)) { New-Item -ItemType Directory -Force -Path $dstDir | Out-Null }
    Copy-Item -LiteralPath $src -Destination $dst -Force
}
Remove-Item -Recurse -Force -LiteralPath $stage

# Prune directories left empty by the stale pass (only inside the target).
$dirs = Get-ChildItem -LiteralPath $Target -Recurse -Directory | Sort-Object { $_.FullName.Length } -Descending
foreach ($d in $dirs) {
    if (@(Get-ChildItem -LiteralPath $d.FullName -Force).Count -eq 0) { Remove-Item -LiteralPath $d.FullName -Force }
}

$copied = Get-ChildItem -LiteralPath $Target -Recurse -File
$sum = ($copied | Measure-Object Length -Sum).Sum
Write-Output ("published " + $copied.Count + " files, " + [math]::Round($sum / 1KB) + " KB -> " + $Target)
