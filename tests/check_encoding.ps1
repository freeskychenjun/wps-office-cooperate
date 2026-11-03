# check_encoding.ps1 - the repository rule "every .ps1 is pure ASCII" as a gate.
#
# Windows PowerShell 5.1 parses a BOM-less .ps1 as ANSI, so a Chinese literal inside one is
# read as mojibake on a GBK machine: the JSON error message gibberishes and paths break.
# The repo therefore keeps its wrappers ASCII-only (pass Chinese/Unicode through
# -Encoding UTF8 data files or runtime arguments). A UTF-8 BOM is allowed - it makes the
# encoding unambiguous - but the bytes after it must still be ASCII.
#
#   .\check_encoding.ps1                 # scripts\*.ps1 + tests\**\*.ps1 + root *.ps1
#   .\check_encoding.ps1 -Path a.ps1,b.ps1
#   .\check_encoding.ps1 -Quiet          # one summary line only
#
# Exit 0 = clean, 1 = at least one non-ASCII byte, 2 = no files found (bad invocation).
#
[CmdletBinding()]
param(
    [string]$Root = '',
    [string[]]$Path = @(),
    [switch]$Quiet
)

$ErrorActionPreference = 'Stop'
if (-not $Root) { $Root = Split-Path -Parent $PSScriptRoot }

function Get-TargetFiles {
    $list = New-Object System.Collections.ArrayList
    if (-not [string]::IsNullOrEmpty($Path)) {
        foreach ($p in $Path) {
            if (-not (Test-Path -LiteralPath $p)) { throw ("no such file: " + $p) }
            [void]$list.Add((Resolve-Path -LiteralPath $p).Path)
        }
        return ,$list
    }
    # The three places the rule covers: repo root (publish.ps1), scripts\*.ps1 (the wrappers)
    # and everything under tests\. scripts\helper\** is C#, not PowerShell, so depth 1 only.
    $targets = @(
        @{ Folder = $Root; Filter = '*.ps1'; Recurse = $false },
        @{ Folder = (Join-Path $Root 'scripts'); Filter = '*.ps1'; Recurse = $false },
        @{ Folder = (Join-Path $Root 'tests'); Filter = '*.ps1'; Recurse = $true }
    )
    foreach ($target in $targets) {
        if (-not (Test-Path -LiteralPath $target.Folder)) { continue }
        $found = @(Get-ChildItem -LiteralPath $target.Folder -Filter $target.Filter -File -Recurse:$target.Recurse)
        foreach ($f in $found) {
            $full = $f.FullName
            if ($full -match '\\\.tmp\\') { continue }
            if ($full -match '\\__pycache__\\') { continue }
            if (-not $list.Contains($full)) { [void]$list.Add($full) }
        }
    }
    return ,$list
}

function Find-NonAscii {
    param([byte[]]$Bytes)
    $offset = 0
    # UTF-8 BOM is explicitly allowed (it is the documented workaround).
    if ($Bytes.Length -ge 3 -and $Bytes[0] -eq 0xEF -and $Bytes[1] -eq 0xBB -and $Bytes[2] -eq 0xBF) { $offset = 3 }
    # A UTF-16 file would parse as garbage even with a BOM; report it as one clear error.
    if ($Bytes.Length -ge 2 -and (($Bytes[0] -eq 0xFF -and $Bytes[1] -eq 0xFE) -or ($Bytes[0] -eq 0xFE -and $Bytes[1] -eq 0xFF))) {
        return @([pscustomobject]@{ Line = 0; Offset = 0; Byte = 0; Reason = 'UTF-16 file (write UTF-8 instead)' })
    }
    $bad = New-Object System.Collections.ArrayList
    $line = 1
    for ($i = $offset; $i -lt $Bytes.Length; $i++) {
        $b = $Bytes[$i]
        if ($b -eq 0x0A) { $line++ ; continue }
        if ($b -gt 0x7F) {
            [void]$bad.Add([pscustomobject]@{ Line = $line; Offset = $i; Byte = $b; Reason = ('non-ASCII byte 0x{0:X2}' -f $b) })
        }
    }
    return ,$bad
}

$files = Get-TargetFiles
if ($files.Count -eq 0) {
    [Console]::Error.WriteLine('check_encoding: no .ps1 files found under ' + $Root)
    exit 2
}

$violations = 0
$badFiles = 0
$checked = 0
foreach ($file in $files) {
    $checked++
    $relative = $file
    if ($file.StartsWith($Root)) { $relative = $file.Substring($Root.Length).TrimStart('\', '/') }
    $bytes = [System.IO.File]::ReadAllBytes($file)
    $bad = Find-NonAscii -Bytes $bytes
    if (-not $bad -or $bad.Count -eq 0) { continue }
    $badFiles++
    $violations += $bad.Count
    if (-not $Quiet) {
        Write-Host ("FAIL  {0}" -f $relative)
        $shown = 0
        foreach ($item in $bad) {
            if ($shown -ge 5) {
                Write-Host ("      ... {0} more non-ASCII byte(s)" -f ($bad.Count - 5))
                break
            }
            if ($item.Line -gt 0) {
                Write-Host ("      line {0}: {1}" -f $item.Line, $item.Reason)
            }
            else {
                Write-Host ("      {0}" -f $item.Reason)
            }
            $shown++
        }
    }
}

if ($badFiles -gt 0) {
    Write-Host ("encoding: FAIL - {0} non-ASCII byte(s) in {1} of {2} .ps1 file(s)" -f $violations, $badFiles, $checked)
    Write-Host '  rule: keep .ps1 pure ASCII (a UTF-8 BOM is allowed); pass Chinese through data or arguments'
    exit 1
}
Write-Host ("encoding: PASS - {0} .ps1 file(s) are pure ASCII" -f $checked)
exit 0
