# run_all.ps1 - one command to run the wps-office-cooperate test suite.
#
# Two layers, because half of this repo is COM automation and the other half is not:
#
#   fast   no WPS, no Office, no network.  Pure-Python fixture builders and oracles, the
#          static docx assertions, the tasks*.ps1 self-tests (pure logic) and the encoding
#          gate.  This is the layer that runs in CI on a clean windows-latest clone.
#   wps    the end-to-end drivers (tests\**\run_*.ps1 and the multi-task check_tasks.ps1).
#          They start WPS/Writer/ET/WPP through the skill wrapper, so they need WPS Office
#          12.1 installed, no other office document open, and 10+ minutes for the full set.
#
#   .\run_all.ps1                 # everything (fast first, stops nothing on failure)
#   .\run_all.ps1 -Layer fast     # CI mode (same as -SkipWps)
#   .\run_all.ps1 -Layer wps
#   .\run_all.ps1 -Layer fast -Only encoding,normalize
#   .\run_all.ps1 -List           # print the test names and exit
#
# Per-test PASS/FAIL/SKIP + elapsed time, a summary, and exit code 1 when anything failed.
# Generated fixtures go to tests\out-fast\ (gitignored; tests\clean.ps1 removes it).
#
# Pure ASCII on purpose (Windows PowerShell 5.1 parses .ps1 as ANSI).
#
[CmdletBinding()]
param(
    [ValidateSet('fast', 'wps', 'all')]
    [string]$Layer = 'all',
    [string[]]$Only = @(),
    [string]$PythonExe = 'python',
    [switch]$SkipWps,
    [switch]$List
)

$ErrorActionPreference = 'Stop'
try { [Console]::OutputEncoding = New-Object System.Text.UTF8Encoding($false) } catch { }

$Tests = $PSScriptRoot
$Root = Split-Path -Parent $Tests
$OutRoot = Join-Path $Tests 'out-fast'
$Woc = Join-Path $Root 'scripts\woc.ps1'
$NoBom = New-Object System.Text.UTF8Encoding($false)

if ($SkipWps) { $Layer = 'fast' }
$Only = @($Only | ForEach-Object { "$_" -split ',' } | ForEach-Object { $_.Trim() } | Where-Object { $_ })

# ---------------------------------------------------------------- helpers

function Invoke-Child {
    <# Run a native command from the repo root and capture stdout+stderr.
       Returns @{ Code; Output }.  Never throws for a non-zero exit code. #>
    param(
        [Parameter(Mandatory = $true)][string]$FilePath,
        [string[]]$Arguments = @(),
        [string]$WorkDir = ''
    )
    if (-not $WorkDir) { $WorkDir = $Root }
    Push-Location -LiteralPath $WorkDir
    try {
        $output = & $FilePath @Arguments 2>&1 | Out-String
        $code = $LASTEXITCODE
        if ($null -eq $code) { $code = 0 }
    }
    finally {
        Pop-Location
    }
    return [pscustomobject]@{ Code = [int]$code; Output = [string]$output }
}

function Invoke-Ps1 {
    param([string]$ScriptPath, [string[]]$Arguments = @())
    $all = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $ScriptPath) + $Arguments
    return Invoke-Child -FilePath 'powershell' -Arguments $all
}

function Invoke-Python {
    param([string[]]$Arguments)
    return Invoke-Child -FilePath $PythonExe -Arguments $Arguments
}

function Get-Tail {
    param([string]$Text, [int]$Lines = 12)
    if (-not $Text) { return '' }
    $rows = @($Text -split "`r?`n" | Where-Object { $_.Trim() })
    if ($rows.Count -le $Lines) { return ($rows -join "`n") }
    return (($rows | Select-Object -Last $Lines) -join "`n")
}

function Get-Prop {
    param($Object, [string]$Name, $Default = $null)
    if ($null -eq $Object) { return $Default }
    if ($Object -is [System.Collections.IDictionary]) {
        if ($Object.Contains($Name)) { return $Object[$Name] }
        return $Default
    }
    $property = $Object.PSObject.Properties[$Name]
    if ($null -eq $property) { return $Default }
    return $property.Value
}

function New-FastTest {
    param([string]$Name, [scriptblock]$Body)
    return [pscustomobject]@{ Name = $Name; Layer = 'fast'; Body = $Body }
}

function Get-Relative([string]$Path) {
    if ($Path.StartsWith($Root, [StringComparison]::OrdinalIgnoreCase)) {
        return $Path.Substring($Root.Length).TrimStart('\')
    }
    return $Path
}

function Write-Prefixed {
    param([string]$Text, [string]$Prefix)
    foreach ($line in ($Text -split "`r?`n")) { Write-Host ($Prefix + $line) }
}

# ---------------------------------------------------------------- fast tests

$fastTests = New-Object System.Collections.ArrayList

[void]$fastTests.Add((New-FastTest 'encoding' {
            $r = Invoke-Ps1 (Join-Path $Tests 'check_encoding.ps1')
            if ($r.Code -ne 0) { return @{ Status = 'FAIL'; Detail = 'non-ASCII .ps1 (see output)'; Output = $r.Output } }
            return @{ Status = 'PASS'; Detail = (Get-Tail $r.Output 1) }
        }))

[void]$fastTests.Add((New-FastTest 'ps1-syntax' {
            # Every .ps1 the encoding gate covers must also parse under the PS 5.1 grammar;
            # a botched edit in one wrapper would otherwise only surface at call time.
            $files = New-Object System.Collections.ArrayList
            foreach ($folder in @($Root, (Join-Path $Root 'scripts'), $Tests)) {
                if (-not (Test-Path -LiteralPath $folder)) { continue }
                $recurse = ($folder -eq $Tests)
                foreach ($f in @(Get-ChildItem -LiteralPath $folder -Filter '*.ps1' -File -Recurse:$recurse)) {
                    if (-not $files.Contains($f.FullName)) { [void]$files.Add($f.FullName) }
                }
            }
            $broken = New-Object System.Collections.ArrayList
            foreach ($file in $files) {
                $tokens = $null
                $errors = $null
                $null = [System.Management.Automation.Language.Parser]::ParseFile($file, [ref]$tokens, [ref]$errors)
                foreach ($err in @($errors)) {
                    [void]$broken.Add(((Get-Relative $file) + ':' + $err.Extent.StartLineNumber + ': ' + $err.Message))
                }
            }
            if ($broken.Count -gt 0) {
                return @{ Status = 'FAIL'; Detail = ($broken.Count.ToString() + ' syntax error(s)'); Output = ($broken -join "`n") }
            }
            return @{ Status = 'PASS'; Detail = ($files.Count.ToString() + ' .ps1 parsed') }
        }))

[void]$fastTests.Add((New-FastTest 'word-normalize-static' {
            # Makes the fixture itself (self-contained, no external corpus) and then runs the
            # static assertions: 14 tables / 4 captions in row 1 / 2 without tblBorders / 0 rev / 0 comments.
            $fixture = Join-Path $Tests 'word\out\normalize-fixture.docx'
            if (Test-Path -LiteralPath $fixture) { Remove-Item -LiteralPath $fixture -Force }
            $gen = Invoke-Python @((Join-Path $Tests 'word\make_normalize_fixture.py'))
            if ($gen.Code -ne 0) { return @{ Status = 'FAIL'; Detail = 'make_normalize_fixture.py failed'; Output = $gen.Output } }
            $check = Invoke-Python @((Join-Path $Tests 'word\check_normalize.py'), '--static-only')
            if ($check.Code -ne 0) { return @{ Status = 'FAIL'; Detail = 'check_normalize.py --static-only failed'; Output = $check.Output } }
            return @{ Status = 'PASS'; Detail = (Get-Tail $check.Output 1) }
        }))

[void]$fastTests.Add((New-FastTest 'ppt-fixtures' {
            # python-pptx is the independent oracle for the WPS layer: if a fixture builder
            # breaks, every ppt run_*.ps1 driver fails for the wrong reason. Build them all.
            $makers = @(Get-ChildItem -LiteralPath (Join-Path $Tests 'ppt') -Filter 'make_*.py' -File | Sort-Object Name)
            if ($makers.Count -eq 0) { return @{ Status = 'FAIL'; Detail = 'no make_*.py under tests\ppt' } }
            $failed = New-Object System.Collections.ArrayList
            $output = New-Object System.Collections.ArrayList
            foreach ($maker in $makers) {
                $out = Join-Path $OutRoot ('ppt\' + $maker.BaseName)
                $r = Invoke-Python @($maker.FullName, '--out', $out)
                [void]$output.Add(('  ' + $maker.Name + ' -> ' + $r.Code))
                if ($r.Code -ne 0) { [void]$failed.Add($maker.Name); [void]$output.Add($r.Output) }
            }
            if ($failed.Count -gt 0) {
                return @{ Status = 'FAIL'; Detail = ($failed -join ', '); Output = ($output -join "`n") }
            }
            return @{ Status = 'PASS'; Detail = ($makers.Count.ToString() + ' fixture decks built') }
        }))

[void]$fastTests.Add((New-FastTest 'ppt-oracle' {
            # check.py owns the ppt replace assertions and computes them with python-pptx.
            # --list only evaluates the pristine fixture, which proves the oracle still
            # agrees with the fixture (the deck has 16 slots carrying the token).
            $folder = Join-Path $OutRoot 'ppt-oracle\full'
            $r = Invoke-Python @((Join-Path $Tests 'ppt\make_deck.py'), '--out', $folder)
            if ($r.Code -ne 0) { return @{ Status = 'FAIL'; Detail = 'make_deck.py failed'; Output = $r.Output } }
            $pristine = Join-Path $folder 'replace-deck.pristine.pptx'
            if (-not (Test-Path -LiteralPath $pristine)) { return @{ Status = 'FAIL'; Detail = 'make_deck.py wrote no pristine deck'; Output = $r.Output } }
            Copy-Item -LiteralPath $pristine -Destination (Join-Path $folder 'pristine.pptx') -Force
            $check = Invoke-Python @((Join-Path $Tests 'ppt\check.py'), '--dir', (Join-Path $OutRoot 'ppt-oracle'), '--list')
            if ($check.Code -ne 0) { return @{ Status = 'FAIL'; Detail = 'check.py --list failed'; Output = $check.Output } }
            if ($check.Output -notmatch '"total":\s*16') {
                return @{ Status = 'FAIL'; Detail = 'expected 16 token hits in the replace fixture'; Output = $check.Output }
            }
            return @{ Status = 'PASS'; Detail = 'oracle sees 16 token slots' }
        }))

[void]$fastTests.Add((New-FastTest 'tasks-selftest' {
            # scripts\tasks.ps1 selftest over the committed chapter map: range coverage,
            # minimal-replace round trip, digit guard, similarity. Pure logic, no WPS.
            $ranges = Join-Path $Tests 'multi-task\ranges.json'
            if (-not (Test-Path -LiteralPath $ranges)) { return @{ Status = 'FAIL'; Detail = ('missing fixture: ' + (Get-Relative $ranges)) } }
            # PS 5.1 Get-Content defaults to the ANSI code page - always read fixtures as UTF-8.
            $rangesText = [System.IO.File]::ReadAllText($ranges, [System.Text.Encoding]::UTF8)
            $readFixtures = $rangesText | ConvertFrom-Json
            # One read item per paragraph up to the last chapter end; the self-test only needs
            # the indices to be consecutive and non-empty inside every range.
            $last = ($readFixtures | Measure-Object -Property to -Maximum).Maximum
            $items = New-Object System.Collections.ArrayList
            for ($i = 1; $i -le [int]$last; $i++) {
                [void]$items.Add([ordered]@{ i = $i; text = ('paragraph ' + $i + ' of the acceptance corpus'); style = 'Body' })
            }
            $readPath = Join-Path $OutRoot 'tasks\read.json'
            $readDir = Split-Path -Parent $readPath
            if (-not (Test-Path -LiteralPath $readDir)) { New-Item -ItemType Directory -Force -Path $readDir | Out-Null }
            $readJson = [ordered]@{ ok = $true; result = [ordered]@{ paragraphs = $items.Count; items = @($items) } }
            [System.IO.File]::WriteAllText($readPath, (ConvertTo-Json -InputObject $readJson -Depth 10 -Compress), $NoBom)
            $run = Invoke-Ps1 (Join-Path $Root 'scripts\tasks.ps1') @('selftest', '-Read', $readPath, '-Ranges', $ranges)
            if ($run.Code -ne 0 -or $run.Output -notmatch 'SELFTEST PASS') {
                return @{ Status = 'FAIL'; Detail = 'tasks.ps1 selftest did not pass'; Output = $run.Output }
            }
            return @{ Status = 'PASS'; Detail = (Get-Tail $run.Output 1) }
        }))

$fastTests.Add((New-FastTest 'tasks-excel-selftest' {
            $script = Join-Path $Root 'scripts\tasks-excel.ps1'
            if (-not (Test-Path -LiteralPath $script)) { return @{ Status = 'SKIP'; Detail = 'missing scripts\tasks-excel.ps1' } }
            $run = Invoke-Ps1 $script @('selftest')
            if ($run.Code -ne 0 -or $run.Output -notmatch 'SELFTEST PASS') {
                return @{ Status = 'FAIL'; Detail = 'tasks-excel.ps1 selftest did not pass'; Output = $run.Output }
            }
            return @{ Status = 'PASS'; Detail = (Get-Tail $run.Output 1) }
        })) | Out-Null

$fastTests.Add((New-FastTest 'tasks-ppt-selftest' {
            $script = Join-Path $Root 'scripts\tasks-ppt.ps1'
            if (-not (Test-Path -LiteralPath $script)) { return @{ Status = 'SKIP'; Detail = 'missing scripts\tasks-ppt.ps1' } }
            $run = Invoke-Ps1 $script @('selftest')
            if ($run.Code -ne 0 -or $run.Output -notmatch 'SELFTEST PASS') {
                return @{ Status = 'FAIL'; Detail = 'tasks-ppt.ps1 selftest did not pass'; Output = $run.Output }
            }
            return @{ Status = 'PASS'; Detail = (Get-Tail $run.Output 1) }
        })) | Out-Null

$fastTests.Add((New-FastTest 'tasks-session-selftest' {
            $script = Join-Path $Root 'scripts\tasks-session.ps1'
            if (-not (Test-Path -LiteralPath $script)) { return @{ Status = 'SKIP'; Detail = 'missing scripts\tasks-session.ps1' } }
            $run = Invoke-Ps1 $script @('selftest')
            if ($run.Code -ne 0 -or $run.Output -notmatch 'SELFTEST PASS') {
                return @{ Status = 'FAIL'; Detail = 'tasks-session.ps1 selftest did not pass'; Output = $run.Output }
            }
            return @{ Status = 'PASS'; Detail = (Get-Tail $run.Output 1) }
        })) | Out-Null

[void]$fastTests.Add((New-FastTest 'publish-whatif' {
            # publish.ps1 deletes stale files in the deploy target, so it grew a -WhatIf dry run.
            # It is a smoke test only: resolve + validate + list, touch nothing.
            $publish = Join-Path $Root 'publish.ps1'
            if (-not (Test-Path -LiteralPath $publish)) { return @{ Status = 'SKIP'; Detail = 'publish.ps1 not present' } }
            $source = [System.IO.File]::ReadAllText($publish)
            if ($source -notmatch 'WhatIf') { return @{ Status = 'SKIP'; Detail = 'publish.ps1 has no -WhatIf yet' } }
            $run = Invoke-Ps1 $publish @('-WhatIf')
            if ($run.Code -ne 0) { return @{ Status = 'FAIL'; Detail = 'publish.ps1 -WhatIf failed'; Output = $run.Output } }
            if ($run.Output -notmatch 'publish:') { return @{ Status = 'FAIL'; Detail = 'publish -WhatIf printed no plan'; Output = $run.Output } }
            return @{ Status = 'PASS'; Detail = (Get-Tail $run.Output 1) }
        }))

# ---------------------------------------------------------------- wps tests (discovered)

function Get-WpsTests {
    $drivers = New-Object System.Collections.ArrayList
    $self = $PSCommandPath
    if (-not $self) { $self = $MyInvocation.MyCommand.Path }
    foreach ($pattern in @('run_*.ps1', 'check_tasks.ps1')) {
        foreach ($f in @(Get-ChildItem -LiteralPath $Tests -Filter $pattern -File -Recurse)) {
            if ($f.FullName -eq $self) { continue }
            [void]$drivers.Add($f)
        }
    }
    $tests = New-Object System.Collections.ArrayList
    foreach ($driver in ($drivers | Sort-Object FullName)) {
        $name = (Get-Relative $driver.FullName).Substring('tests\'.Length) -replace '\.ps1$', '' -replace '\\', '/'
        $path = $driver.FullName
        $tests.Add([pscustomobject]@{
                Name  = $name
                Layer = 'wps'
                Body  = {
                    $run = Invoke-Ps1 $path
                    if ($run.Code -ne 0) { return @{ Status = 'FAIL'; Detail = ('exit ' + $run.Code); Output = $run.Output } }
                    return @{ Status = 'PASS'; Detail = (Get-Tail $run.Output 1) }
                }.GetNewClosure()
            }) | Out-Null
    }
    return ,$tests
}

# ---------------------------------------------------------------- run

$selected = New-Object System.Collections.ArrayList
if ($Layer -eq 'fast' -or $Layer -eq 'all') {
    foreach ($test in $fastTests) { [void]$selected.Add($test) }
}
if ($Layer -eq 'wps' -or $Layer -eq 'all') {
    foreach ($test in (Get-WpsTests)) { [void]$selected.Add($test) }
}
if ($Only.Count -gt 0) {
    $filtered = New-Object System.Collections.ArrayList
    foreach ($test in $selected) {
        $hit = $false
        foreach ($needle in $Only) { if ($test.Name -like ('*' + $needle + '*')) { $hit = $true } }
        if ($hit) { [void]$filtered.Add($test) }
    }
    $selected = $filtered
}

if ($List) {
    foreach ($test in $selected) { Write-Host ('{0,-8} {1}' -f $test.Layer, $test.Name) }
    Write-Host ('{0} test(s)' -f $selected.Count)
    exit 0
}

if ($selected.Count -eq 0) {
    [Console]::Error.WriteLine('run_all: no test matches (layer ' + $Layer + ', -Only ' + ($Only -join ',') + ')')
    exit 2
}

$started = Get-Date
$passed = 0
$failed = 0
$skipped = 0
$failedNames = New-Object System.Collections.ArrayList
$notRun = New-Object System.Collections.ArrayList

Write-Host ('wps-office-cooperate tests - layer {0}, {1} test(s), python {2}' -f $Layer, $selected.Count, $PythonExe)
Write-Host ''

foreach ($test in $selected) {
    $watch = [System.Diagnostics.Stopwatch]::StartNew()
    $status = 'FAIL'
    $detail = ''
    $output = ''
    try {
        $result = & $test.Body
        $status = [string](Get-Prop $result 'Status' 'FAIL')
        $detail = [string](Get-Prop $result 'Detail' '')
        $output = [string](Get-Prop $result 'Output' '')
    }
    catch {
        $status = 'FAIL'
        $detail = $_.Exception.Message
    }
    finally {
        $watch.Stop()
    }
    $seconds = $watch.Elapsed.TotalSeconds

    switch ($status) {
        'PASS' { $passed++ }
        'SKIP' { $skipped++ }
        default { $status = 'FAIL'; $failed++; [void]$failedNames.Add($test.Name) }
    }
    Write-Host ('{0,-4} {1,-34} {2,6:N1}s  {3}' -f $status, $test.Name, $seconds, $detail)
    if ($status -eq 'FAIL' -and $output) { Write-Prefixed -Text (Get-Tail $output 15) -Prefix '     | ' }
    if ($status -eq 'FAIL' -and $test.Layer -eq 'wps') { [void]$notRun.Add($test.Name) }
}

$elapsed = ((Get-Date) - $started).TotalSeconds
Write-Host ''
Write-Host ('== summary: {0} passed, {1} failed, {2} skipped in {3:N1}s (layer {4})' -f $passed, $failed, $skipped, $elapsed, $Layer)
if ($failed -gt 0) {
    Write-Host ('   failed: ' + ($failedNames -join ', '))
    exit 1
}
if ($skipped -gt 0) { Write-Host ('   skipped: see the SKIP reasons above (missing optional tooling)') }
exit 0
