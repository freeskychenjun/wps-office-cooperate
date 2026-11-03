# clean.ps1 - remove the local build/run residue of the test suite.
#
# Everything this script deletes is either gitignored or untracked; it never touches source.
# The only two places it is allowed to delete in are:
#
#   <repo>\tests\**\out*          generated fixtures / run outputs (run_all.ps1, run_*.ps1)
#   <repo>\tests\**\__pycache__   python byte code
#   <repo>\.tmp\impl              local scratch of the v1.33 acceptance runs
#
# A path outside those roots aborts the whole run (guard, not a warning).
#
#   .\clean.ps1            # delete and print every removed path
#   .\clean.ps1 -DryRun    # print what would be deleted
#
# Pure ASCII on purpose (Windows PowerShell 5.1 parses .ps1 as ANSI).
#
[CmdletBinding()]
param(
    [string]$Root = '',
    [switch]$DryRun
)

$ErrorActionPreference = 'Stop'
if (-not $Root) { $Root = Split-Path -Parent $PSScriptRoot }
$Root = [System.IO.Path]::GetFullPath($Root)
$Tests = [System.IO.Path]::GetFullPath((Join-Path $Root 'tests'))
$Tmp = [System.IO.Path]::GetFullPath((Join-Path $Root '.tmp'))
$Allowed = @($Tests, $Tmp)

function Get-Relative([string]$Path) {
    if ($Path.StartsWith($Root, [StringComparison]::OrdinalIgnoreCase)) {
        return $Path.Substring($Root.Length).TrimStart('\')
    }
    return $Path
}

function Assert-Allowed([string]$Path) {
    <# Refuse anything that is not strictly inside tests\ or .tmp\. #>
    $full = [System.IO.Path]::GetFullPath($Path)
    foreach ($root in $Allowed) {
        $trimmed = $root.TrimEnd('\')
        if ($full.Equals($trimmed, [StringComparison]::OrdinalIgnoreCase)) {
            throw ('refusing to delete a root directory: ' + $full)
        }
        if ($full.StartsWith($trimmed + '\', [StringComparison]::OrdinalIgnoreCase)) { return }
    }
    throw ('refusing to delete outside tests/ and .tmp/: ' + $full)
}

# ---------------------------------------------------------------- collect

$targets = New-Object System.Collections.ArrayList
$notes = New-Object System.Collections.ArrayList

if (Test-Path -LiteralPath $Tests) {
    # out* covers out/, out-fast/, out-check/, out-master/, out.stale-*/ ...
    foreach ($d in @(Get-ChildItem -LiteralPath $Tests -Directory -Recurse -ErrorAction SilentlyContinue |
                Where-Object { $_.Name -like 'out*' -or $_.Name -eq '__pycache__' })) {
        [void]$targets.Add($d.FullName)
    }
}

$impl = Join-Path $Tmp 'impl'
if (Test-Path -LiteralPath $impl) { [void]$targets.Add($impl) }
$fixture = Join-Path $Tests 'word\out\normalize-fixture.docx'
if (Test-Path -LiteralPath $fixture) { [void]$notes.Add('(covered by tests\word\out) ' + (Get-Relative $fixture)) }

# Drop children of a directory that is already being deleted, then delete deepest first.
$sorted = @($targets | Sort-Object Length)
$final = New-Object System.Collections.ArrayList
foreach ($path in $sorted) {
    $covered = $false
    foreach ($kept in $final) {
        if ($path.StartsWith($kept + '\', [StringComparison]::OrdinalIgnoreCase)) { $covered = $true; break }
    }
    if (-not $covered) { [void]$final.Add($path) }
}
$final = @($final | Sort-Object Length -Descending)

if ($final.Count -eq 0) {
    Write-Host 'clean: nothing to remove'
    exit 0
}

# ---------------------------------------------------------------- delete

$removed = 0
$items = 0
foreach ($path in $final) {
    Assert-Allowed $path
    if (-not (Test-Path -LiteralPath $path)) { continue }
    $count = @(Get-ChildItem -LiteralPath $path -Recurse -Force -ErrorAction SilentlyContinue).Count + 1
    if ($DryRun) {
        Write-Host ('would remove  {0}  ({1} item(s))' -f (Get-Relative $path), $count)
    }
    else {
        Remove-Item -LiteralPath $path -Recurse -Force
        Write-Host ('removed       {0}  ({1} item(s))' -f (Get-Relative $path), $count)
    }
    $removed++
    $items += $count
}
foreach ($note in $notes) { Write-Host ('note          ' + $note) }

$verb = 'removed'
if ($DryRun) { $verb = 'would remove' }
Write-Host ('clean: {0} {1} path(s), {2} item(s)' -f $verb, $removed, $items)
exit 0
