# run_data.ps1 - Excel data-layer acceptance driver (wps-office-cooperate, excel side).
#
#   .\run_data.ps1                       # full suite (starts WPS, ~1-2 minutes)
#   .\run_data.ps1 -SkipWps              # list the cases only; no WPS, no fixture
#   .\run_data.ps1 -Only write,fill      # a subset of the cases
#   .\run_data.ps1 -SkipDecrypt          # assert on the skill JSON only (no file oracle)
#
# What it does, per case:
#   1. copy the pristine fixture (make_workbook.py) into <work>\cases\<case>\<case>.xlsx
#   2. drive the skill exactly the way an agent would: scripts\woc.ps1 excel <cmd> ...
#      every JSON response is stored under <work>\logs\<case>-<step>.json
#   3. save + close the workbook (close runs even when a step failed)
# Then, once: decrypt the saved workbooks (WPS writes them transparently encrypted on this
# machine, so openpyxl cannot read them), run check_data.py, which owns every assertion and
# prints the PASS/FAIL lines. This script only starts WPS, records what came back and turns
# check_data.py's verdict into its own exit code.
#
# The workbook under test is generated on the fly and lives under %TEMP% - no user document
# is ever opened or saved.
#
# Pure ASCII on purpose: Windows PowerShell 5.1 parses .ps1 as ANSI, so a non-ASCII literal in
# this file would be mangled (the Chinese folder name of the decrypt tool therefore comes from
# local.config.json or from directory enumeration, never from a literal).
[CmdletBinding()]
param(
    [string]$PythonExe = 'python',
    [string]$OutDir = '',
    [string[]]$Only = @(),
    # Path of the offline DGS transparent-encryption decryptor (a .ps1 taking -Paths/-OutDir).
    # Empty = local.config.json, then a filesystem search, then the "decrypt-office-pdf" skill.
    [string]$DecryptEngine = '',
    [switch]$SkipWps,
    [switch]$SkipDecrypt,
    [int]$StepTimeoutSec = 90,
    [switch]$KeepWork
)

$ErrorActionPreference = 'Stop'
try { [Console]::OutputEncoding = New-Object System.Text.UTF8Encoding($false) } catch { }

# `powershell -File run_data.ps1 -Only a,b` passes "a,b" as ONE string, so split it here.
$Only = @($Only | ForEach-Object { "$_" -split ',' } | ForEach-Object { $_.Trim() } | Where-Object { $_ })

$Root = $PSScriptRoot
$Repo = Split-Path -Parent (Split-Path -Parent $Root)
$Woc = Join-Path $Repo 'scripts\woc.ps1'
$MakeWorkbook = Join-Path $Root 'make_workbook.py'
$CheckData = Join-Path $Root 'check_data.py'
$Utf8 = New-Object System.Text.UTF8Encoding($false)

if (-not (Test-Path -LiteralPath $Woc)) { throw "skill entry point not found: $Woc" }
if (-not (Test-Path -LiteralPath $MakeWorkbook)) { throw "fixture builder not found: $MakeWorkbook" }
if (-not (Test-Path -LiteralPath $CheckData)) { throw "oracle not found: $CheckData" }

if (-not $OutDir) { $OutDir = Join-Path $env:TEMP 'woc-excel-regression' }
$OutDir = [System.IO.Path]::GetFullPath($OutDir)

# Transport: the resident daemon stays on (1 s per command vs 3.5 s for a one-shot helper), but
# every case starts and ends on a clean runtime - the daemon is stopped and any et.exe/EtHelper
# this run started is killed after the case. That removes the failure mode seen on 2026-11-04:
# a daemon answering for an app instance that had already quit, which left a document open in a
# half-dead instance and made the next case hang on WPS's "two workbooks in one instance" bug.
$script:Style = @{ Pass = 'Green'; Fail = 'Red'; Skip = 'Yellow'; Head = 'Cyan' }

function Write-Head([string]$text) { Write-Host $text -ForegroundColor $script:Style.Head }
function Write-Dim([string]$text) { Write-Host $text -ForegroundColor DarkGray }

# ---------------------------------------------------------------- process helpers

function Quote-Arg([string]$text) {
    if ($text -eq $null) { return '""' }
    if ($text.Length -eq 0) { return '""' }
    if ($text -match '[\s"]') { return '"' + ($text -replace '"', '\"') + '"' }
    return $text
}

function Get-ProcCount([string]$name) {
    $lines = & tasklist /FI ("IMAGENAME eq " + $name) /NH 2>$null
    return @($lines | Select-String -SimpleMatch $name).Count
}

# Kill only the automation processes THIS run started (start time after $StartedAt): a hung WPS
# instance or helper is reset between cases so one wedged case cannot poison the next. A process
# that predates the run belongs to the user or another skill and is never touched.
function Reset-ExcelRuntime {
    param([datetime]$StartedAt, [switch]$Quiet)
    $killed = 0
    foreach ($name in @('et', 'EtHelper')) {
        foreach ($one in @(Get-Process $name -ErrorAction SilentlyContinue)) {
            try {
                if ($one.StartTime -ge $StartedAt) {
                    Stop-Process -Id $one.Id -Force -ErrorAction SilentlyContinue
                    $killed++
                }
            } catch { }
        }
    }
    if ($killed -gt 0 -and -not $Quiet) {
        Write-Host ("    reset: killed " + $killed + " automation process(es) this run started") -ForegroundColor Yellow
    }
    return $killed
}

# One skill call. Returns a hashtable:
#   Json     parsed JSON response (or $null)
#   Text     raw stdout
#   Stderr   raw stderr
#   ExitCode child exit code
#   TimedOut $true when the step was killed
# The response is stored verbatim under <work>\logs\<case>-<step>.json so check_data.py can
# judge it; a non-JSON answer is stored as a synthetic ok:false envelope.
function Invoke-Step {
    param(
        $Ctx,
        [string]$Step,
        [string[]]$ExcelArgs,
        [int]$TimeoutSec = 0
    )
    if ($TimeoutSec -le 0) { $TimeoutSec = $StepTimeoutSec }
    if ($Ctx.Aborted) {
        # The case already hit a timeout; do not spend another watchdog interval on every step.
        $dead = @{ ok = $false; error = @{ code = 'CASE_ABORTED'; message = 'an earlier step in this case timed out' }; result = @{} }
        [System.IO.File]::WriteAllText((Join-Path $Ctx.Logs ($Ctx.Id + '-' + $Step + '.json')),
            ($dead | ConvertTo-Json -Depth 12 -Compress), $Utf8)
        [System.IO.File]::WriteAllText((Join-Path $Ctx.Logs ($Ctx.Id + '-' + $Step + '.ms.txt')), '0', $Utf8)
        Write-Host ("    {0,-18} (aborted: earlier timeout)" -f $Step) -ForegroundColor Yellow
        return $dead
    }
    $argLine = '-Mta -NoProfile -ExecutionPolicy Bypass -File ' + (Quote-Arg $Woc) + ' excel'
    foreach ($one in $ExcelArgs) { $argLine += ' ' + (Quote-Arg ([string]$one)) }
    # The wrapper kills the helper itself when -TimeoutSec expires and answers with TIMEOUT JSON;
    # that is what keeps a COM hang from costing the driver's whole wait.
    if (@($ExcelArgs).Count -gt 0 -and $ExcelArgs[0] -ne 'serve' -and $ExcelArgs[0] -ne 'ping') {
        $argLine += ' -TimeoutSec ' + [string]([int]$StepTimeoutSec)
    }

    $attempt = 0
    $result = $null
    $stepWatch = [System.Diagnostics.Stopwatch]::StartNew()
    while ($true) {
        $attempt++
        $psi = New-Object System.Diagnostics.ProcessStartInfo
        $psi.FileName = 'powershell.exe'
        $psi.Arguments = $argLine
        $psi.UseShellExecute = $false
        $psi.RedirectStandardOutput = $true
        $psi.RedirectStandardError = $true
        $psi.CreateNoWindow = $true
        $psi.StandardOutputEncoding = $Utf8
        $psi.StandardErrorEncoding = $Utf8
        $proc = $null
        $outTask = $null
        $errTask = $null
        $timedOut = $false
        $raw = ''
        $err = ''
        try {
            $proc = [System.Diagnostics.Process]::Start($psi)
            $outTask = $proc.StandardOutput.ReadToEndAsync()
            $errTask = $proc.StandardError.ReadToEndAsync()
            if (-not $proc.WaitForExit($TimeoutSec * 1000)) {
                $timedOut = $true
                try { $proc.Kill() } catch { }
            }
            try { $raw = $outTask.Result } catch { $raw = '' }
            try { $err = $errTask.Result } catch { $err = '' }
        } catch {
            $err = '' + $_.Exception.Message
        }
        $parsed = $null
        $trim = ($raw).Trim()
        if (-not $timedOut -and $trim.Length -gt 0) {
            try { $parsed = $trim | ConvertFrom-Json } catch { $parsed = $null }
        }
        $result = @{
            Json = $parsed; Text = $raw; Stderr = $err
            ExitCode = if ($proc) { $proc.ExitCode } else { -1 }
            TimedOut = $timedOut; Attempts = $attempt
        }
        # No automatic retry: several of the commands here mutate the workbook (write, rows,
        # dedupe ...), so a retry after a half-finished call would execute them twice. A wrapper
        # that is being edited shows up as a NOT_JSON FAIL instead - re-run once it is stable.
        break
    }

    if ($null -eq $result.Json) {
        $code = if ($result.TimedOut) { 'DRIVER_TIMEOUT' } else { 'NOT_JSON' }
        $message = if ($result.TimedOut) {
            'no answer within ' + $TimeoutSec + ' s; the PowerShell child was killed'
        } else {
            ('stdout was not JSON (exit ' + $result.ExitCode + '): ' + (($result.Text + $result.Stderr) -replace '\s+', ' ')).Trim()
        }
        $synthetic = @{ ok = $false; error = @{ code = $code; message = $message }; result = @{} }
    } else {
        $synthetic = $result.Json
    }
    $stepWatch.Stop()
    [System.IO.File]::WriteAllText((Join-Path $Ctx.Logs ($Ctx.Id + '-' + $Step + '.ms.txt')),
        [string][int]$stepWatch.Elapsed.TotalMilliseconds, $Utf8)
    $logPath = Join-Path $Ctx.Logs ($Ctx.Id + '-' + $Step + '.json')
    [System.IO.File]::WriteAllText($logPath, ($synthetic | ConvertTo-Json -Depth 24 -Compress), $Utf8)
    if ($result.Stderr -and $result.Stderr.Trim().Length -gt 0) {
        [System.IO.File]::WriteAllText((Join-Path $Ctx.Logs ($Ctx.Id + '-' + $Step + '.stderr.txt')), $result.Stderr, $Utf8)
    }

    # One live line per step: what came back, cut to something a human can scan.
    $ok = $synthetic.ok
    $short = (($synthetic | ConvertTo-Json -Depth 24 -Compress) -replace '\s+', ' ')
    if ($short.Length -gt 160) { $short = $short.Substring(0, 160) + '...' }
    $color = if ($ok -eq $true) { 'DarkGray' } else { 'Yellow' }
    Write-Host ("    {0,-18} {1}" -f $Step, $short) -ForegroundColor $color
    if ($result.TimedOut) { Write-Host ("    " + $Step + " TIMED OUT after " + $TimeoutSec + " s") -ForegroundColor Red }
    $errCode = ''
    if ($synthetic.error) { $errCode = [string]$synthetic.error.code }
    if ($result.TimedOut -or $errCode -eq 'TIMEOUT' -or $errCode -eq 'DRIVER_TIMEOUT') {
        # Abort the rest of this case and drop the instance it was driving: a WPS that did not
        # answer within the watchdog is not worth another command.
        $Ctx.Aborted = $true
        Reset-ExcelRuntime -StartedAt $script:RunStartedAt | Out-Null
    }
    return $synthetic
}

function New-ItemsFile {
    param($Ctx, [string]$Name, $Items)
    $path = Join-Path $Ctx.Dir $Name
    $json = ConvertTo-Json -InputObject @{ items = @($Items) } -Depth 8 -Compress
    [System.IO.File]::WriteAllText($path, $json, $Utf8)
    return $path
}

function Save-As {
    param($Ctx, [string]$FileName, [string]$Step)
    $target = Join-Path $Ctx.Dir $FileName
    $r = Invoke-Step $Ctx $Step @('save', '-Doc', $ctx.Doc, '-As', $target)
    if ($r.ok -eq $true) { $Ctx.Doc = $target }
    return $r
}

function Open-Doc {
    param($Ctx, [string]$Path, [string]$Step)
    $r = Invoke-Step $Ctx $Step @('open', '-Doc', $Path)
    if ($r.ok -ne $true) {
        # Everything after a failed open would just fail; drop the case and make sure whatever
        # half-open document it left behind does not reach the next case.
        $Ctx.Aborted = $true
        Reset-ExcelRuntime -StartedAt $script:RunStartedAt | Out-Null
    }
    return $r
}

# The daemon is a per-source-hash singleton of this skill; stopping it here only affects the
# helper this run talks to (a user document never lives in it, close already ran).
function Stop-ExcelDaemon {
    param([string]$Logs)
    try { Invoke-Step @{ Id = 'zz-daemon'; Logs = $Logs } 'stop' @('serve', '-Action', 'stop') | Out-Null } catch { }
}

# ---------------------------------------------------------------- the cases
#
# Every case is a scriptblock receiving $ctx: Start in a fresh copy of the fixture, run the
# commands, save. The driver closes the workbook in a finally block. All judgement lives in
# check_data.py; the numbers here are the same ones it re-derives from the pristine fixture.
$script:Cases = @(
    @{ Id = '01-write'; Title = 'write: plain values, a formula and a value block'; Run = {
        param($ctx)
        Open-Doc $ctx $ctx.Doc 'open' | Out-Null
        $items = New-ItemsFile $ctx 'items.json' @(
            @{ cell = 'Data!E1'; value = 'computed' },
            @{ cell = 'Data!E2'; formula = '=AVERAGE(B2:B4)' },
            @{ cell = 'Data!F2:G3'; values = @(@(1, 2), @(3, 4)) }
        )
        Invoke-Step $ctx 'write' @('write', '-Doc', $ctx.Doc, '-Items', $items) | Out-Null
        Invoke-Step $ctx 'save' @('save', '-Doc', $ctx.Doc) | Out-Null
    } }
    @{ Id = '02-write-anchor'; Title = 'write: expected-text mismatch -> ok:false, nothing written'; Run = {
        param($ctx)
        Open-Doc $ctx $ctx.Doc 'open' | Out-Null
        $items = New-ItemsFile $ctx 'items.json' @(
            @{ cell = 'Data!B2'; value = 777; expect = '10' },
            @{ cell = 'Data!H1'; value = 'x'; expect = 'WRONG' }
        )
        Invoke-Step $ctx 'write' @('write', '-Doc', $ctx.Doc, '-Items', $items) | Out-Null
        Invoke-Step $ctx 'save' @('save', '-Doc', $ctx.Doc) | Out-Null
    } }
    @{ Id = '03-fill'; Title = 'fill: relative formula down 1999 rows (acceptance case 17)'; Run = {
        param($ctx)
        Open-Doc $ctx $ctx.Doc 'open' | Out-Null
        Invoke-Step $ctx 'fill' @('fill', '-Doc', $ctx.Doc, '-Sheet', 'Fill', '-Range', 'B2:B2000', '-Formula', '=A2*2+1') | Out-Null
        Invoke-Step $ctx 'save' @('save', '-Doc', $ctx.Doc) | Out-Null
    } }
    @{ Id = '04-replace'; Title = 'replace: dry-run, single pass, numeric+formula guards, idempotency'; Run = {
        param($ctx)
        Open-Doc $ctx $ctx.Doc 'open' | Out-Null
        # numbers made of digits must never be rewritten as text
        Invoke-Step $ctx 'dryrun-number' @('replace', '-Doc', $ctx.Doc, '-Sheet', 'Text', '-Column', 'code', '-Find', '0', '-Replace', '9', '-DryRun') | Out-Null
        # a formula cell must never be overwritten by a value
        Invoke-Step $ctx 'formula-guard' @('replace', '-Doc', $ctx.Doc, '-Sheet', 'Text', '-Find', 'B2', '-Replace', 'B9') | Out-Null
        # the dry run itself
        Invoke-Step $ctx 'dryrun' @('replace', '-Doc', $ctx.Doc, '-Sheet', 'Text', '-Column', 'name', '-Find', 'P', '-Replace', 'P-', '-MatchCase', '-DryRun') | Out-Null
        Save-As $ctx '04-replace-dryrun.xlsx' 'save-dryrun' | Out-Null
        # the real thing: PP must become P-P- (one pass), not P--P--
        Invoke-Step $ctx 'apply' @('replace', '-Doc', $ctx.Doc, '-Sheet', 'Text', '-Column', 'name', '-Find', 'P', '-Replace', 'P-', '-MatchCase') | Out-Null
        # Default matching is case-insensitive: ALPHA must hit the lower-case "alpha" of A4.
        Invoke-Step $ctx 'case-insensitive' @('replace', '-Doc', $ctx.Doc, '-Sheet', 'Text', '-Column', 'name', '-Find', 'ALPHA', '-Replace', 'omega', '-DryRun') | Out-Null
        Invoke-Step $ctx 'apply-idem-1' @('replace', '-Doc', $ctx.Doc, '-Sheet', 'Text', '-Column', 'label', '-Find', 'alpha', '-Replace', 'omega') | Out-Null
        Invoke-Step $ctx 'apply-idem-2' @('replace', '-Doc', $ctx.Doc, '-Sheet', 'Text', '-Column', 'label', '-Find', 'alpha', '-Replace', 'omega') | Out-Null
        Save-As $ctx '04-replace-final.xlsx' 'save-final' | Out-Null
    } }
    @{ Id = '05-sort'; Title = 'sort: descending by a number column, formulas travel with the row'; Run = {
        param($ctx)
        Open-Doc $ctx $ctx.Doc 'open' | Out-Null
        Invoke-Step $ctx 'sort' @('sort', '-Doc', $ctx.Doc, '-Sheet', 'Sort', '-By', 'qty:desc') | Out-Null
        Invoke-Step $ctx 'save' @('save', '-Doc', $ctx.Doc) | Out-Null
    } }
    @{ Id = '06-deletewhere'; Title = 'deletewhere: drop rows by a text condition'; Run = {
        param($ctx)
        Open-Doc $ctx $ctx.Doc 'open' | Out-Null
        Invoke-Step $ctx 'delete' @('deletewhere', '-Doc', $ctx.Doc, '-Sheet', 'Where', '-Column', 'status', '-Value', 'drop') | Out-Null
        Invoke-Step $ctx 'save' @('save', '-Doc', $ctx.Doc) | Out-Null
    } }
    @{ Id = '07-fillblanks'; Title = 'fillblanks: -With above, then -With 0'; Run = {
        param($ctx)
        Open-Doc $ctx $ctx.Doc 'open' | Out-Null
        Invoke-Step $ctx 'above' @('fillblanks', '-Doc', $ctx.Doc, '-Sheet', 'Blanks', '-Column', 'qty', '-With', 'above') | Out-Null
        Invoke-Step $ctx 'zero' @('fillblanks', '-Doc', $ctx.Doc, '-Sheet', 'Blanks', '-Column', 'qty', '-With', '0') | Out-Null
        Invoke-Step $ctx 'save' @('save', '-Doc', $ctx.Doc) | Out-Null
    } }
    @{ Id = '08-dedupe'; Title = 'dedupe: duplicates by one column, first of each group kept'; Run = {
        param($ctx)
        Open-Doc $ctx $ctx.Doc 'open' | Out-Null
        Invoke-Step $ctx 'dedupe' @('dedupe', '-Doc', $ctx.Doc, '-Sheet', 'Dedupe', '-Column', 'site') | Out-Null
        Invoke-Step $ctx 'save' @('save', '-Doc', $ctx.Doc) | Out-Null
    } }
    @{ Id = '09-clean'; Title = 'clean: trim/case/no-spaces, then text -> number'; Run = {
        param($ctx)
        Open-Doc $ctx $ctx.Doc 'open' | Out-Null
        Invoke-Step $ctx 'upper' @('clean', '-Doc', $ctx.Doc, '-Sheet', 'Clean', '-Column', 'name', '-Casing', 'upper', '-NoSpaces') | Out-Null
        Invoke-Step $ctx 'number' @('clean', '-Doc', $ctx.Doc, '-Sheet', 'Clean', '-Column', 'qty', '-To', 'number') | Out-Null
        Invoke-Step $ctx 'save' @('save', '-Doc', $ctx.Doc) | Out-Null
    } }
    @{ Id = '10-table'; Title = 'table: create a ListObject, second call only restyles'; Run = {
        param($ctx)
        Open-Doc $ctx $ctx.Doc 'open' | Out-Null
        Invoke-Step $ctx 'create' @('table', '-Doc', $ctx.Doc, '-Sheet', 'Table', '-Range', 'A1:C7', '-Name', 'TblOne', '-Style', 'TableStyleMedium2') | Out-Null
        Invoke-Step $ctx 'again' @('table', '-Doc', $ctx.Doc, '-Sheet', 'Table', '-Range', 'A1:C7', '-Name', 'TblOne') | Out-Null
        Invoke-Step $ctx 'save' @('save', '-Doc', $ctx.Doc) | Out-Null
    } }
    @{ Id = '11-sheet'; Title = 'sheet: list / add / write / rename / delete'; Run = {
        param($ctx)
        Open-Doc $ctx $ctx.Doc 'open' | Out-Null
        Invoke-Step $ctx 'list' @('sheet', '-Doc', $ctx.Doc, '-Action', 'list') | Out-Null
        Invoke-Step $ctx 'add' @('sheet', '-Doc', $ctx.Doc, '-Action', 'add', '-Name', 'NewSheet') | Out-Null
        Invoke-Step $ctx 'write-new' @('write', '-Doc', $ctx.Doc, '-Cell', 'NewSheet!A1', '-Value', 'marker') | Out-Null
        Invoke-Step $ctx 'rename' @('sheet', '-Doc', $ctx.Doc, '-Action', 'rename', '-Name', 'Summary', '-NewName', 'Summary2') | Out-Null
        Invoke-Step $ctx 'delete' @('sheet', '-Doc', $ctx.Doc, '-Action', 'delete', '-Name', 'NewSheet') | Out-Null
        Invoke-Step $ctx 'save' @('save', '-Doc', $ctx.Doc) | Out-Null
    } }
    @{ Id = '12-rowscols'; Title = 'rows/cols: insert 2 rows then put them back (two sessions)'; Run = {
        param($ctx)
        Open-Doc $ctx $ctx.Doc 'open' | Out-Null
        Invoke-Step $ctx 'rows-insert' @('rows', '-Doc', $ctx.Doc, '-Action', 'insert', '-Sheet', 'Rows', '-Index', '2', '-Count', '2') | Out-Null
        Save-As $ctx '12-rowscols-inserted.xlsx' 'save-inserted' | Out-Null
        Invoke-Step $ctx 'close-inserted' @('close', '-Doc', $ctx.Doc) | Out-Null
        # Keep a copy of the shifted workbook: the second session saves over the same file, so
        # the shifted layout would otherwise be lost before check_data.py could look at it.
        Copy-Item -LiteralPath $ctx.Doc -Destination (Join-Path $ctx.Dir '12-rowscols-shifted.xlsx') -Force
        # second session on the saved, shifted workbook
        Open-Doc $ctx $ctx.Doc 'open-inserted' | Out-Null
        Invoke-Step $ctx 'rows-delete' @('rows', '-Doc', $ctx.Doc, '-Action', 'delete', '-Sheet', 'Rows', '-Index', '2', '-Count', '2') | Out-Null
        Invoke-Step $ctx 'cols-insert' @('cols', '-Doc', $ctx.Doc, '-Action', 'insert', '-Sheet', 'Rows', '-Index', '3', '-Count', '1') | Out-Null
        Invoke-Step $ctx 'cols-delete' @('cols', '-Doc', $ctx.Doc, '-Action', 'delete', '-Sheet', 'Rows', '-Index', '3', '-Count', '1') | Out-Null
        Invoke-Step $ctx 'save' @('save', '-Doc', $ctx.Doc) | Out-Null
    } }
    @{ Id = '13-snapshot'; Title = 'snapshot: sheet-scoped diff and restore isolation (case 19)'; Run = {
        param($ctx)
        Open-Doc $ctx $ctx.Doc 'open' | Out-Null
        Invoke-Step $ctx 'snapshot' @('snapshot', '-Doc', $ctx.Doc, '-Sheet', 'Data') | Out-Null
        $items = New-ItemsFile $ctx 'items-dirty.json' @(
            @{ cell = 'Data!C2'; value = 999 },
            @{ cell = 'Summary!B2'; value = 888 }
        )
        Invoke-Step $ctx 'dirty' @('write', '-Doc', $ctx.Doc, '-Items', $items) | Out-Null
        Invoke-Step $ctx 'diff' @('diff', '-Doc', $ctx.Doc) | Out-Null
        Invoke-Step $ctx 'restore' @('restore', '-Doc', $ctx.Doc) | Out-Null
        Invoke-Step $ctx 'save' @('save', '-Doc', $ctx.Doc) | Out-Null
    } }
    @{ Id = '14-restore'; Title = 'restore -Cells: selective rollback leaves the rest alone (case 28)'; Run = {
        param($ctx)
        Open-Doc $ctx $ctx.Doc 'open' | Out-Null
        Invoke-Step $ctx 'snapshot' @('snapshot', '-Doc', $ctx.Doc, '-Sheet', 'Data') | Out-Null
        $items = New-ItemsFile $ctx 'items-f.json' @(
            @{ cell = 'Data!F2'; value = 111 },
            @{ cell = 'Data!F3'; value = 222 },
            @{ cell = 'Data!F4'; value = 333 }
        )
        Invoke-Step $ctx 'dirty' @('write', '-Doc', $ctx.Doc, '-Items', $items) | Out-Null
        Invoke-Step $ctx 'restore-one' @('restore', '-Doc', $ctx.Doc, '-Cells', 'Data!F2') | Out-Null
        Invoke-Step $ctx 'save' @('save', '-Doc', $ctx.Doc) | Out-Null
    } }
    @{ Id = '15-chart'; Title = 'chart: add two, list, delete one, then read the drawing back'; Run = {
        param($ctx)
        Open-Doc $ctx $ctx.Doc 'open' | Out-Null
        Invoke-Step $ctx 'add-line' @('chart', '-Doc', $ctx.Doc, '-Action', 'add', '-Sheet', 'Series', '-Type', 'line', '-Range', 'A1:C21', '-Anchor', 'E2', '-Title', 'Q-H', '-Name', 'ChartQH') | Out-Null
        Invoke-Step $ctx 'add-column' @('chart', '-Doc', $ctx.Doc, '-Action', 'add', '-Sheet', 'Series', '-Type', 'column', '-Range', 'A1:B21', '-Anchor', 'E28', '-Title', 'QQ') | Out-Null
        Invoke-Step $ctx 'list-2' @('chart', '-Doc', $ctx.Doc, '-Action', 'list', '-Sheet', 'Series') | Out-Null
        Invoke-Step $ctx 'delete-1' @('chart', '-Doc', $ctx.Doc, '-Action', 'delete', '-Sheet', 'Series', '-Index', '1') | Out-Null
        Invoke-Step $ctx 'list-1' @('chart', '-Doc', $ctx.Doc, '-Action', 'list', '-Sheet', 'Series') | Out-Null
        Invoke-Step $ctx 'save' @('save', '-Doc', $ctx.Doc) | Out-Null
    } }
    @{ Id = '16-xls'; Title = 'xls round trip: SaveAs .xls, reopen, read, write expect, fill, save'; Run = {
        param($ctx)
        Open-Doc $ctx $ctx.Doc 'open' | Out-Null
        Save-As $ctx '16-xls.xls' 'save-as-xls' | Out-Null
        Invoke-Step $ctx 'read' @('read', '-Doc', $ctx.Doc, '-Sheet', 'Data', '-Range', 'A1:D3') | Out-Null
        Invoke-Step $ctx 'write' @('write', '-Doc', $ctx.Doc, '-Cell', 'Data!B2', '-Value', '777', '-Expect', '10') | Out-Null
        Invoke-Step $ctx 'fill' @('fill', '-Doc', $ctx.Doc, '-Sheet', 'Fill', '-Range', 'C2:C101', '-Formula', '=A2*2') | Out-Null
        Invoke-Step $ctx 'save' @('save', '-Doc', $ctx.Doc) | Out-Null
    } }
)

# ---------------------------------------------------------------- decryptor discovery

function Find-DecryptEngine {
    param([string]$Explicit)
    if ($Explicit) {
        if (Test-Path -LiteralPath $Explicit) { return (Resolve-Path -LiteralPath $Explicit).Path }
        Write-Host ("  decrypt engine not found: " + $Explicit) -ForegroundColor Yellow
    }
    $cfg = Join-Path $Root 'local.config.json'
    if (Test-Path -LiteralPath $cfg) {
        try {
            $o = [System.IO.File]::ReadAllText($cfg, [System.Text.Encoding]::UTF8) | ConvertFrom-Json
            if ($o.decryptEngine) {
                $p = [string]$o.decryptEngine
                if (Test-Path -LiteralPath $p) { return $p }
            }
        } catch { }
    }
    # Search the drives for <dir>\FastDycrypter\FastDycrypt\engine\dec.ps1 without spelling the
    # (non-ASCII) folder name in this file.
    foreach ($drive in @('D:\', 'C:\', 'E:\', 'F:\')) {
        if (-not (Test-Path -LiteralPath $drive)) { continue }
        $dirs = @()
        try { $dirs = @(Get-ChildItem -LiteralPath $drive -Directory -ErrorAction SilentlyContinue) } catch { }
        foreach ($d in $dirs) {
            $candidate = Join-Path $d.FullName 'FastDycrypter\FastDycrypt\engine\dec.ps1'
            if (Test-Path -LiteralPath $candidate) { return $candidate }
        }
    }
    # Last resort: the decrypt skill ships a (slower) reader that talks to a disguised app.
    $skill = Join-Path $env:USERPROFILE '.agents\skills\decrypt-office-pdf\decrypt-files.ps1'
    if (Test-Path -LiteralPath $skill) { return $skill }
    return $null
}

function Invoke-Decrypt {
    param([string]$Engine, [string]$Source, [string]$Target)
    $argLine = '-NoProfile -ExecutionPolicy Bypass -File ' + (Quote-Arg $Engine) + ' -Paths ' + (Quote-Arg $Source) +
        ' -OutDir ' + (Quote-Arg $Target) + ' -Recurse -Quiet -NoPause'
    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = 'powershell.exe'
    $psi.Arguments = $argLine
    $psi.UseShellExecute = $false
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError = $true
    $psi.CreateNoWindow = $true
    $psi.StandardOutputEncoding = $Utf8
    $psi.StandardErrorEncoding = $Utf8
    $proc = [System.Diagnostics.Process]::Start($psi)
    $outTask = $proc.StandardOutput.ReadToEndAsync()
    $errTask = $proc.StandardError.ReadToEndAsync()
    $done = $proc.WaitForExit(240000)
    $text = ''
    try { $text = $outTask.Result } catch { }
    $err = ''
    try { $err = $errTask.Result } catch { }
    if (-not $done) { try { $proc.Kill() } catch { } }
    return @{ TimedOut = (-not $done); ExitCode = $proc.ExitCode; Text = $text; Stderr = $err }
}

# Same capture, for the decrypt skill's in-place reader (different parameter names).
function Invoke-DecryptInPlace {
    param([string]$Engine, [string]$Source)
    $argLine = '-NoProfile -ExecutionPolicy Bypass -File ' + (Quote-Arg $Engine) + ' -Path ' + (Quote-Arg $Source) + ' -Recurse'
    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = 'powershell.exe'
    $psi.Arguments = $argLine
    $psi.UseShellExecute = $false
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError = $true
    $psi.CreateNoWindow = $true
    $psi.StandardOutputEncoding = $Utf8
    $psi.StandardErrorEncoding = $Utf8
    $proc = [System.Diagnostics.Process]::Start($psi)
    $outTask = $proc.StandardOutput.ReadToEndAsync()
    $errTask = $proc.StandardError.ReadToEndAsync()
    $done = $proc.WaitForExit(240000)
    $text = ''
    try { $text = $outTask.Result } catch { }
    $err = ''
    try { $err = $errTask.Result } catch { }
    if (-not $done) { try { $proc.Kill() } catch { } }
    return @{ TimedOut = (-not $done); ExitCode = $proc.ExitCode; Text = $text; Stderr = $err }
}

# ---------------------------------------------------------------- main

$script:RunStartedAt = Get-Date
$startedAt = $script:RunStartedAt
$wpsBefore = Get-ProcCount 'wps.exe'
$etBefore = Get-ProcCount 'et.exe'

$selected = @()
foreach ($case in $script:Cases) {
    if ($Only.Count -gt 0 -and -not ($Only | Where-Object { $_ -eq $case.Id })) { continue }
    $selected += $case
}

if ($selected.Count -eq 0) {
    Write-Host "no case matched -Only $($Only -join ',')" -ForegroundColor Red
    exit 1
}

Write-Head "== excel data-layer regression: $($selected.Count) case(s) =="
foreach ($case in $selected) { Write-Host ("  {0,-16} {1}" -f $case.Id, $case.Title) }

if ($SkipWps) {
    Write-Host ("listed {0} case(s); -SkipWps so WPS was not started" -f $selected.Count) -ForegroundColor Green
    exit 0
}

# ---- work area
$casesDir = Join-Path $OutDir 'cases'
$logsDir = Join-Path $OutDir 'logs'
$plainDir = Join-Path $OutDir 'plain'
foreach ($leaf in @('cases', 'logs', 'plain')) {
    $dir = Join-Path $OutDir $leaf
    if (Test-Path -LiteralPath $dir) { Remove-Item -Recurse -Force -LiteralPath $dir }
    New-Item -ItemType Directory -Force -Path $dir | Out-Null
}
$basePath = Join-Path $OutDir 'base.xlsx'
if (Test-Path -LiteralPath $basePath) { Remove-Item -Force -LiteralPath $basePath }

Write-Host ""
Write-Host ("work area: " + $OutDir)
Write-Host ("fixture:   " + $MakeWorkbook)
$makeOut = & $PythonExe $MakeWorkbook --out $basePath 2>&1
if ($LASTEXITCODE -ne 0) {
    Write-Host "fixture generation failed:" -ForegroundColor Red
    $makeOut | ForEach-Object { Write-Host ("  " + $_) }
    exit 1
}
Write-Dim ("  " + ($makeOut -join ' '))

# ---- drive each case on a clean runtime
Stop-ExcelDaemon $logsDir
Reset-ExcelRuntime -StartedAt $script:RunStartedAt -Quiet | Out-Null
foreach ($case in $selected) {
    $caseDir = Join-Path $casesDir $case.Id
    New-Item -ItemType Directory -Force -Path $caseDir | Out-Null
    $docPath = Join-Path $caseDir ($case.Id + '.xlsx')
    Copy-Item -LiteralPath $basePath -Destination $docPath -Force
    $ctx = @{
        Id = $case.Id; Title = $case.Title; Dir = $caseDir
        Doc = $docPath; Logs = $logsDir; Base = $basePath; Step = 0
    }
    Write-Head ("-- " + $case.Id + ": " + $case.Title)
    $caseWatch = [System.Diagnostics.Stopwatch]::StartNew()
    try {
        & $case.Run $ctx
    } catch {
        Write-Host ("    driver error: " + $_.Exception.Message) -ForegroundColor Red
        [System.IO.File]::WriteAllText((Join-Path $logsDir ($case.Id + '-driver-error.txt')), ('' + $_), $Utf8)
    } finally {
        $close = Invoke-Step $ctx 'close' @('close', '-Doc', $ctx.Doc)
        # Clean slate for the next case: the daemon is stopped (it would otherwise survive for an
        # app instance that close just quit) and any automation process this run started is
        # dropped, so a half-dead instance can never hold two documents at once (known WPS hang).
        Stop-ExcelDaemon $logsDir
        Reset-ExcelRuntime -StartedAt $script:RunStartedAt -Quiet | Out-Null
        $caseWatch.Stop()
        Write-Dim ("    case time: {0:N1} s" -f $caseWatch.Elapsed.TotalSeconds)
    }
}

# ---- file-level oracle: WPS writes transparently encrypted workbooks on this machine
$fileOracle = $true
$skipReason = ''
$fileRoot = $plainDir
if (-not $SkipDecrypt) {
    $engine = Find-DecryptEngine $DecryptEngine
    if (-not $engine) {
        $fileOracle = $false
        $skipReason = 'no transparent-encryption decryptor found (pass -DecryptEngine or -SkipDecrypt)'
    } else {
        Write-Host ""
        Write-Host ("decrypting saved workbooks with: " + $engine)
        # The offline DGS engine writes decrypted copies to -OutDir; the decrypt-office-pdf skill
        # script decrypts in place and takes -Path instead.
        $inPlace = ((Split-Path -Leaf $engine) -eq 'decrypt-files.ps1')
        if ($inPlace) {
            $dec = Invoke-DecryptInPlace $engine $casesDir
            $fileRoot = $casesDir
        } else {
            $dec = Invoke-Decrypt $engine $casesDir $plainDir
        }
        if ($dec.TimedOut) {
            $fileOracle = $false
            $skipReason = 'decryptor timed out after 240 s'
        }
        $plainCount = 0
        if (Test-Path -LiteralPath $fileRoot) {
            $plainCount = @(Get-ChildItem -LiteralPath $fileRoot -Recurse -File -ErrorAction SilentlyContinue).Count
        }
        Write-Dim ("  decrypted files: " + $plainCount)
        if ($plainCount -eq 0) {
            $fileOracle = $false
            $skipReason = 'decryptor produced no files (see its output above)'
        }
    }
} else {
    $fileOracle = $false
    $skipReason = '-SkipDecrypt'
}
if (-not $fileOracle) {
    Write-Host ("file-level oracle skipped: " + $skipReason) -ForegroundColor Yellow
}

# ---- assertions live in check_data.py
Write-Host ""
$checkArgs = @($CheckData, '--work', $OutDir)
if ($fileOracle) { $checkArgs += '--files'; $checkArgs += $fileRoot }
if (-not $fileOracle) { $checkArgs += '--skip-file' }
if ($skipReason) { $checkArgs += '--skip-reason'; $checkArgs += $skipReason }
& $PythonExe @checkArgs
$checkExit = $LASTEXITCODE

# ---- summary
$summary = $null
$summaryPath = Join-Path $OutDir 'summary.json'
if (Test-Path -LiteralPath $summaryPath) {
    try { $summary = [System.IO.File]::ReadAllText($summaryPath, [System.Text.Encoding]::UTF8) | ConvertFrom-Json } catch { }
}
# Stop the resident helper this run may have started, so no headless WPS instance is left behind.
Stop-ExcelDaemon $logsDir
Start-Sleep -Seconds 2
# A case that died before its close leaves an automation instance behind; an instance is this
# run's only when its et.exe started after this run began (a WPS window the user or another
# skill owns predates it), so that is the entire kill list. wps.exe is never touched.
$leaked = @()
try { $leaked = @(Get-Process et -ErrorAction SilentlyContinue | Where-Object { $_.StartTime -ge $startedAt }) } catch { $leaked = @() }
foreach ($one in $leaked) { try { Stop-Process -Id $one.Id -Force -ErrorAction SilentlyContinue } catch { } }
if ($leaked.Count -gt 0) {
    Write-Host ("cleaned up " + $leaked.Count + " leftover et.exe automation instance(s) started by this run") -ForegroundColor Yellow
    Start-Sleep -Milliseconds 800
}
$wpsAfter = Get-ProcCount 'wps.exe'
$etAfter = Get-ProcCount 'et.exe'

$elapsed = ((Get-Date) - $startedAt)
Write-Host ""
Write-Head "=================================================================="
if ($summary) {
    $line = "SUMMARY  PASS {0}  FAIL {1}  SKIP {2}   (cases {3}, commands run: see {4})" -f $summary.pass, $summary.fail, $summary.skip, $summary.cases, $logsDir
    if ($summary.fail -gt 0) { Write-Host $line -ForegroundColor Red } else { Write-Host $line -ForegroundColor Green }
} else {
    Write-Host "SUMMARY  missing - check_data.py did not write summary.json" -ForegroundColor Red
}
Write-Host ("ELAPSED  {0:N1} s" -f $elapsed.TotalSeconds) -ForegroundColor Cyan
Write-Host ("PROCESSES et.exe {0} -> {1} (ours are cleaned up), wps.exe {2} -> {3} (untouched)" -f $etBefore, $etAfter, $wpsBefore, $wpsAfter) -ForegroundColor DarkGray
Write-Head "=================================================================="

if (-not $summary -or $summary.fail -gt 0 -or $checkExit -ne 0) { exit 1 }
exit 0
