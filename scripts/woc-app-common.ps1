# Shared plumbing for the three app wrappers (scripts\word.ps1, excel.ps1, ppt.ps1), v1.34.
#
# Why this file exists: those three wrappers started as standalone copies and had already drifted
# - ppt.ps1 had lost -r:Microsoft.CSharp.dll from its csc line, and Start-Daemon looped 40 times
# in word/excel but 50 times in woc-common.ps1. Everything that is not "CLI surface or
# app-specific request building" lives here, once.
#
# A wrapper dot-sources this file right after its param block:
#
#   . "$PSScriptRoot\woc-app-common.ps1"
#   $exePath = Initialize-WocHelper -App 'word' -ScriptRoot $PSScriptRoot -ExeName 'WpsHelper.exe'
#
# Dot-sourcing (not Import-Module) is deliberate: the definitions land in the *calling wrapper's*
# script scope, so $script:daemonPipe below is the very variable Invoke-Daemon and the wrapper
# share (verified: a $script: assignment made inside a function defined in a dot-sourced file is
# visible to the dot-sourcing script), and $PSScriptRoot inside this file is scripts\ - the folder
# this file lives in, which also holds the wrappers and their helper\ sources.

# ---------------------------------------------------------------- shared helpers

function Write-Error-Json([string]$code, [string]$message) {
    $payload = @{ ok = $false; error = @{ code = $code; message = $message } }
    [Console]::Out.Write(($payload | ConvertTo-Json -Depth 6 -Compress))
    exit 1
}

function Resolve-FullPath([string]$p) {
    try { return [System.IO.Path]::GetFullPath($p) } catch { return $p }
}

# Did the helper answer ok:false (-> wrapper exit code 2)?
#
# Order matters, and it is not the obvious one: a RAW regex match first. Almost every response is
# ok:true and does not contain the literal at all, so the fast path decides those without the cost
# (and the 2 MB JavaScriptSerializer limit) of ConvertFrom-Json. Only when the literal occurs do we
# parse and ask the TOP-LEVEL ok - a response whose payload happens to contain a nested
# "ok": false (batch entries, document text that talks about JSON) is still a success.
# The regex is the fallback only when the response is not parseable JSON.
function Test-JsonOkFalse([string]$text) {
    if ($text -notmatch '"ok"\s*:\s*false') { return $false }
    try {
        $obj = $text | ConvertFrom-Json
        foreach ($prop in $obj.PSObject.Properties) {
            if ($prop.Name -eq 'ok') { return (-not [bool]$prop.Value) }
        }
        return $true
    } catch {
        return $true
    }
}

# Every command that takes a list file does the same steps; word.ps1 used to inline them six times.
#   -Keys         top-level keys to unwrap, in order (a bare array is used as-is)
#   -NoArrayWrap  hand back the value untouched (table -Data may be an object, not a list)
# The leading comma matters: it makes the function emit ONE object, so a caller's
# "$request['data'] = Read-Items-Json ..." stores exactly what the old inline
# "$request['data'] = $parsed.data" stored - including a one-row [[1,2]] staying nested.
function Read-Items-Json {
    param(
        [string]$Path,
        [string[]]$Keys = @('items'),
        [switch]$NoArrayWrap
    )
    if (-not (Test-Path -LiteralPath $Path)) { Write-Error-Json 'BAD_ARGS' ('Items file not found: ' + $Path) }
    $raw = [System.IO.File]::ReadAllText((Resolve-FullPath $Path), [System.Text.Encoding]::UTF8)
    $parsed = $raw | ConvertFrom-Json
    $names = @($parsed.PSObject.Properties.Name)
    foreach ($key in $Keys) {
        if ($names -contains $key) {
            if ($NoArrayWrap) { return ,$parsed.$key }
            return ,@($parsed.$key)
        }
    }
    if ($NoArrayWrap) { return ,$parsed }
    return ,@($parsed)
}

# ---------------------------------------------------------------- resident daemon (v1.14)
# A warm helper process holds the COM application and answers request JSON over a named pipe, so a
# command no longer pays "start exe + attach to WPS + find document" (measured ~2.5 s per command).
# The pipe name is derived from the compiled binary's hash, so a rebuilt helper simply gets a new
# pipe and the old daemon is left alone (it exits on idle).
$script:daemonPipe = $null   # set by Initialize-WocHelper, read by Invoke-Daemon

function Get-DaemonPipeName([string]$app, [string]$shaGen12) { return ('woc-' + $app + '-' + $shaGen12) }

# Ask the resident helper for one response.
#   -ReadTimeoutMs  how long the response may take before this returns an ok:false TIMEOUT envelope
#                   (0 = unbounded, the pre-v1.34 behaviour). Wrappers pass $TimeoutSec * 1000;
#                   internal probes pass a short value. The point is that a helper stuck inside a
#                   COM call blocked by a WPS dialog answers nothing at all, and an unbounded
#                   $pipe.Read would then wait exactly as long as the dialog stays open - forever.
function Invoke-Daemon([string]$payload, [int]$connectMs, [int]$ReadTimeoutMs = 0) {
    if ($script:daemonPipe -eq $null) { return $null }
    $pipe = $null
    try {
        $pipe = New-Object System.IO.Pipes.NamedPipeClientStream('.', $script:daemonPipe, [System.IO.Pipes.PipeDirection]::InOut)
        $pipe.Connect($connectMs)
        $bytes = [System.Text.Encoding]::UTF8.GetBytes($payload)
        $len = [System.BitConverter]::GetBytes([int]$bytes.Length)
        $pipe.Write($len, 0, 4)
        $pipe.Write($bytes, 0, $bytes.Length)
        $pipe.Flush()
        $watch = [System.Diagnostics.Stopwatch]::StartNew()
        $hdr = New-Object byte[] 4
        $got = 0
        while ($got -lt 4) {
            $n = Read-PipeBounded -Pipe $pipe -Buffer $hdr -Offset $got -Count (4 - $got) -Watch $watch -TimeoutMs $ReadTimeoutMs
            if ($n -eq -1) { return (New-DaemonTimeoutEnvelope $ReadTimeoutMs) }
            if ($n -le 0) { return $null }
            $got += $n
        }
        $total = [System.BitConverter]::ToInt32($hdr, 0)
        if ($total -le 0 -or $total -gt 33554432) { return $null }
        $buf = New-Object byte[] $total
        $got = 0
        while ($got -lt $total) {
            $n = Read-PipeBounded -Pipe $pipe -Buffer $buf -Offset $got -Count ($total - $got) -Watch $watch -TimeoutMs $ReadTimeoutMs
            if ($n -eq -1) { return (New-DaemonTimeoutEnvelope $ReadTimeoutMs) }
            if ($n -le 0) { return $null }
            $got += $n
        }
        return [System.Text.Encoding]::UTF8.GetString($buf)
    } catch {
        return $null
    } finally {
        if ($pipe) { try { $pipe.Dispose() } catch { } }
    }
}

# Read at least one byte from the pipe, but never longer than the caller allows.
# Returns the byte count, 0 when the helper closed the pipe (protocol failure -> one-shot fallback),
# or -1 when the deadline passed (helper alive but not answering: a blocked COM call).
# The read MUST stay async and be waited on in slices: a synchronous $pipe.Read blocks for as long
# as WPS keeps a modal dialog open, i.e. until the user closes it - "the tool hangs forever".
# TimeoutMs <= 0 means "no bound" (the pre-v1.34 behaviour), used only by internal callers that
# deliberately want to wait.
function Read-PipeBounded {
    param($Pipe, [byte[]]$Buffer, [int]$Offset, [int]$Count, $Watch, [int]$TimeoutMs)
    $task = $Pipe.ReadAsync($Buffer, $Offset, $Count)
    while ($true) {
        if ($TimeoutMs -le 0) {
            try { return $task.Result } catch { return 0 }
        }
        $remaining = $TimeoutMs - [int]$Watch.ElapsedMilliseconds
        if ($remaining -le 0) { return -1 }
        $slice = [Math]::Min(500, $remaining)
        try {
            if ($task.Wait($slice)) { return $task.Result }
        } catch {
            return 0
        }
    }
}

# The ok:false envelope a read timeout turns into: the caller prints it as-is and the ordinary
# "ok:false -> exit 2" rule applies, so a blocked WPS dialog becomes a normal, actionable error.
function New-DaemonTimeoutEnvelope([int]$TimeoutMs) {
    $seconds = [string]::Format([System.Globalization.CultureInfo]::InvariantCulture, '{0:0.###}', ($TimeoutMs / 1000.0))
    $payload = @{
        ok = $false
        error = @{
            code = 'TIMEOUT'
            message = ('the resident helper did not answer within ' + $seconds + ' s; the COM call is probably blocked by a WPS dialog - close it, then run serve -Action reset')
        }
    }
    return ($payload | ConvertTo-Json -Depth 6 -Compress)
}

# Start the resident helper if it is not already answering, then poll for it. 50 iterations (5 s)
# is the single number all three apps use now: word/excel used to stop at 40.
function Start-Daemon([string]$exePath2) {
    if ($env:WOC_NO_DAEMON -eq '1') { return $false }
    $probe = '{"cmd":"ping"}'
    if ($null -ne (Invoke-Daemon $probe 300 1000)) { return $true }
    try {
        Start-Process -FilePath $exePath2 -ArgumentList @('--serve', $script:daemonPipe) -WindowStyle Hidden | Out-Null
    } catch {
        return $false
    }
    for ($i = 0; $i -lt 50; $i++) {
        Start-Sleep -Milliseconds 100
        if ($null -ne (Invoke-Daemon $probe 300 1000)) { return $true }
    }
    return $false
}

# ---------------------------------------------------------------- helper binary
# Collect helper\<app>\*.cs + helper\shared\*.cs, hash them, and compile the exe into a cache dir
# named after the hash (first run only: csc.exe takes 1-2 s, afterwards the exe is reused until a
# source file changes). Returns the exe path; also sets $script:daemonPipe for this process.
#
# -Refs is spelled out per app even though it is currently identical: -r:Microsoft.CSharp.dll is
# required by the `dynamic` bindings all three helpers use (ppt: 372 occurrences, e.g.
# ppt\Canvas.cs:115 `private static dynamic KLine(...)`). Builds still succeed without the flag
# because csc.exe reads its default response file (Framework64\v4.0.30319\csc.rsp), which contains
# `/r:Microsoft.CSharp.dll` - so ppt.ps1's missing flag was harmless drift, not intent. Keep it
# explicit so the dependency does not silently die with a csc.rsp change.
function Initialize-WocHelper {
    param(
        [Parameter(Mandatory = $true)][string]$App,
        [Parameter(Mandatory = $true)][string]$ScriptRoot,
        [Parameter(Mandatory = $true)][string]$ExeName,
        [string]$Subdir = $App,
        [string[]]$Refs = @('-r:System.dll', '-r:System.Core.dll', '-r:System.Drawing.dll', '-r:System.Windows.Forms.dll', '-r:Microsoft.CSharp.dll', '/codepage:65001')
    )
    $sourceDir = Join-Path $ScriptRoot ('helper\' + $Subdir)
    if (-not (Test-Path -LiteralPath $sourceDir)) { Write-Error-Json 'NO_HELPER' ("Helper folder not found: " + $sourceDir) }
    $sources = @(Get-ChildItem -LiteralPath $sourceDir -Filter *.cs | Sort-Object Name)
    if ($sources.Count -eq 0) { Write-Error-Json 'NO_HELPER' ('No .cs in ' + $sourceDir) }
    # Shared sources (the tray daemon / hint plumbing) compile into every helper.
    $sharedDir = Join-Path $ScriptRoot 'helper\shared'
    if (Test-Path -LiteralPath $sharedDir) { $sources += @(Get-ChildItem -LiteralPath $sharedDir -Filter *.cs | Sort-Object Name) }

    # The cache key covers every source file, so editing any of them rebuilds. Get-FileHash over
    # every .cs file costs ~50 ms on each command (measured 2026-11-02), so woc-common.ps1 caches
    # it keyed on file count + sizes + newest mtime (any real edit changes one of them).
    if (Get-Command Get-CachedSourceHash -ErrorAction SilentlyContinue) {
        $sha = Get-CachedSourceHash -App $App -Sources $sources
    } else {
        $hashInput = ($sources | ForEach-Object { $_.Name + ':' + (Get-FileHash -LiteralPath $_.FullName -Algorithm SHA256).Hash }) -join ';'
        $shaGen = [System.Security.Cryptography.SHA256]::Create()
        $sha = ([System.BitConverter]::ToString($shaGen.ComputeHash([System.Text.Encoding]::UTF8.GetBytes($hashInput))) -replace '-', '').ToLowerInvariant().Substring(0, 12)
    }
    $cacheDir = Join-Path $env:LOCALAPPDATA ('wps-office-cooperate\' + $App + '-helper-' + $sha)
    $script:daemonPipe = Get-DaemonPipeName $App $sha
    $exePath = Join-Path $cacheDir $ExeName
    if (Test-Path -LiteralPath $exePath) { return $exePath }

    $csc = $null
    foreach ($arch in @('Framework64', 'Framework')) {
        $candidate = Join-Path $env:SystemRoot ("Microsoft.NET\" + $arch + "\v4.0.30319\csc.exe")
        if (Test-Path -LiteralPath $candidate) { $csc = $candidate; break }
    }
    if (-not $csc) { Write-Error-Json 'NO_COMPILER' 'The .NET Framework C# compiler (csc.exe) was not found. Enable ".NET Framework 4.8" in Windows Features.' }
    New-Item -ItemType Directory -Force -Path $cacheDir | Out-Null
    $tempExe = Join-Path $cacheDir ('build-' + $PID + '.exe')
    $compileLog = & $csc -nologo -optimize+ -target:exe ("-out:" + $tempExe) $Refs ($sources | ForEach-Object { $_.FullName }) 2>&1 | Out-String
    if (-not (Test-Path -LiteralPath $tempExe)) {
        Write-Error-Json 'COMPILE_FAILED' ("csc failed: " + $compileLog.Trim())
    }
    Move-Item -LiteralPath $tempExe -Destination $exePath -Force
    return $exePath
}

# ---------------------------------------------------------------- batch command (v1.19)
# Run many requests over one process start: the CLI path pays a PowerShell start + wrapper parse
# (~220 ms) per command, while a pipe round trip is 1-15 ms (measured 2026-11-02), so "one page =
# 15 commands" becomes one tool call instead of fifteen.
#   woc.ps1 word batch -Ops ops.json
#   ops.json: {"ops":[{"cmd":"...","doc":"...","sheet":"..."}, ...]}
# The envelope is ok:true when at least one op succeeded (okCount/failedCount say how the rest
# went); a batch where EVERY op failed is ok:false and exits 2, so a caller that only checks the
# envelope cannot mistake "the daemon is gone / the docs are missing" for success.
function Invoke-WocBatchCli {
    param([string]$Ops, [string]$ExePath)
    if (-not $Ops) { Write-Error-Json 'BAD_ARGS' 'batch needs -Ops <ops.json> (an object with an ops array)' }
    $opsPath = Resolve-FullPath $Ops
    if (-not (Test-Path -LiteralPath $opsPath)) { Write-Error-Json 'BAD_ARGS' ('ops file not found: ' + $opsPath) }
    $rawOps = [System.IO.File]::ReadAllText($opsPath, [System.Text.Encoding]::UTF8)
    try { $parsedOps = $rawOps | ConvertFrom-Json }
    catch { Write-Error-Json 'BAD_ARGS' ('ops file is not valid JSON: ' + $_.Exception.Message) }
    if ($parsedOps.PSObject.Properties.Name -contains 'ops') { $opList = @($parsedOps.ops) } else { $opList = @($parsedOps) }
    if ($opList.Count -eq 0) { Write-Error-Json 'BAD_ARGS' 'batch ops is empty' }
    if (-not (Get-Command Invoke-WocBatch -ErrorAction SilentlyContinue)) {
        $wocCommonPath = Join-Path $PSScriptRoot 'woc-common.ps1'
        if (Test-Path -LiteralPath $wocCommonPath) { . $wocCommonPath }
    }
    $batchResult = Invoke-WocBatch -PipeName $script:daemonPipe -ExePath $ExePath -Ops $opList
    [Console]::Out.Write(($batchResult | ConvertTo-Json -Depth 12 -Compress))
    if (-not [bool]$batchResult['ok']) { exit 2 }
    exit 0
}

# ---------------------------------------------------------------- transport
# Resident path first: same request JSON, same response JSON, just no process start and no COM
# attach. Anything unexpected (daemon gone, pipe error) falls through to the one-shot path.
#
# Returns the response text on ok:true and $null when the daemon did not answer at all, so the
# wrapper can run its own exit-code rules on it (ppt -ExitCode -> 3) and print each response once.
# ok:false still exits 2 here - that rule is the same for every app.
function Invoke-WocDaemonFirst {
    param([string]$Command, [string]$Json, [int]$TimeoutSec = 180)
    if ($Command -eq 'serve' -or $env:WOC_NO_DAEMON -eq '1') { return $null }
    $daemonText = Invoke-Daemon $Json 900 ($TimeoutSec * 1000)
    if ($null -eq $daemonText) { return $null }
    [Console]::Out.Write($daemonText)
    if (Test-JsonOkFalse $daemonText) { exit 2 }
    return $daemonText
}

# One-shot path: write the request to a temp file, run the helper exe, echo its response file.
# Returns the response text so a wrapper can add its own exit-code rules (ppt -ExitCode -> 3).
function Invoke-WocOneShot {
    param(
        [string]$App,
        [string]$Command,
        [string]$ExePath,
        [string]$Json,
        [int]$TimeoutSec
    )
    $work = Join-Path $env:TEMP ('wps-office-cooperate-' + $App + '-' + [guid]::NewGuid().ToString('N'))
    New-Item -ItemType Directory -Force -Path $work | Out-Null
    $requestPath = Join-Path $work 'request.json'
    $responsePath = Join-Path $work 'response.json'
    [System.IO.File]::WriteAllText($requestPath, $Json, (New-Object System.Text.UTF8Encoding($false)))

    $p = Start-Process -FilePath $ExePath -ArgumentList @($requestPath, $responsePath) -PassThru -NoNewWindow
    if (-not $p.WaitForExit($TimeoutSec * 1000)) {
        try { $p.Kill() } catch { }
        Remove-Item -Recurse -Force -LiteralPath $work -ErrorAction SilentlyContinue
        Write-Error-Json 'TIMEOUT' ('The helper did not answer within ' + $TimeoutSec + ' s. WPS may be showing a dialog; close it and try again.')
    }
    if (-not (Test-Path -LiteralPath $responsePath)) {
        Remove-Item -Recurse -Force -LiteralPath $work -ErrorAction SilentlyContinue
        Write-Error-Json 'INTERNAL' ('The helper exited (code ' + $p.ExitCode + ') without writing a response.')
    }
    $text = [System.IO.File]::ReadAllText($responsePath, [System.Text.Encoding]::UTF8)
    Remove-Item -Recurse -Force -LiteralPath $work -ErrorAction SilentlyContinue
    [Console]::Out.Write($text)
    if (Test-JsonOkFalse $text) { exit 2 }
    # First contact warms the resident helper for the commands that follow (open is always first in
    # a real session; it is also the cheapest place to pay the one-time start).
    if ($Command -eq 'open') { $null = Start-Daemon $ExePath }
    return $text
}

# ---------------------------------------------------------------- shared command bodies
# serve -Action start|stop|status|reset: control the resident helper (identical in all three apps).
# Never returns - every branch writes its JSON envelope and exits.
#   start   make sure a warm helper is running (no-op when one already answers)
#   stop    polite stop: the helper exits after answering (a helper stuck inside a blocked COM call
#           cannot answer it - use reset for that)
#   status  is a helper answering this build's pipe?
#   reset   recovery: kill THIS build's helper processes, so the next command starts a fresh one
function Invoke-WocServe {
    param([string]$Action, [string]$ExePath)
    $act = if ($Action) { $Action.ToLowerInvariant() } else { 'start' }
    if ($act -eq 'stop') {
        $null = Invoke-Daemon '{"cmd":"daemon-stop"}' 800 5000
        $payload = @{ ok = $true; cmd = 'serve'; result = @{ action = 'stop'; pipe = $script:daemonPipe; note = 'stop requested; the daemon exits after answering' } }
        [Console]::Out.Write(($payload | ConvertTo-Json -Depth 6 -Compress))
        exit 0
    }
    if ($act -eq 'status') {
        $alive = ($null -ne (Invoke-Daemon '{"cmd":"ping"}' 300 5000))
        $payload = @{ ok = $true; cmd = 'serve'; result = @{ action = 'status'; running = $alive; pipe = $script:daemonPipe } }
        [Console]::Out.Write(($payload | ConvertTo-Json -Depth 6 -Compress))
        exit 0
    }
    if ($act -eq 'reset') {
        # The recovery path for a helper stuck inside a blocked COM call (a WPS modal dialog): it
        # cannot answer daemon-stop, so it has to be killed. Match the EXECUTABLE PATH, not the image
        # name - WpsHelper.exe / EtHelper.exe / WppHelper.exe are shared by every cached build and by
        # other sessions, and killing by name would take out unrelated warm helpers.
        $killed = 0
        foreach ($proc in @(Get-Process -ErrorAction SilentlyContinue)) {
            $procPath = $null
            try { $procPath = $proc.Path } catch { }
            if ($procPath -and $procPath -eq $ExePath) {
                try { Stop-Process -Id $proc.Id -Force -ErrorAction Stop; $killed++ } catch { }
            }
        }
        $payload = @{ ok = $true; cmd = 'serve'; result = @{ action = 'reset'; killed = $killed; pipe = $script:daemonPipe; note = 'helper processes of this build were stopped; the next command starts a fresh daemon' } }
        [Console]::Out.Write(($payload | ConvertTo-Json -Depth 6 -Compress))
        exit 0
    }
    $okStart = Start-Daemon $ExePath
    if ($okStart) {
        $payload = @{ ok = $true; cmd = 'serve'; result = @{ action = 'start'; started = $true; pipe = $script:daemonPipe } }
    } else {
        $payload = @{ ok = $false; error = @{ code = 'DAEMON_FAILED'; message = 'the resident helper did not come up; commands still work through the one-shot path' } }
    }
    [Console]::Out.Write(($payload | ConvertTo-Json -Depth 6 -Compress))
    exit 0
}

# open: a takeover must be visible, so the window is shown by default (an earlier hidden open left
# the user staring at nothing while the helper edited a window they could not see). Pass
# -Show:$false explicitly if a hidden session is really what you want.
function Add-WocOpenFields {
    param($Request, $Bound, [string]$Doc, [switch]$ReadOnly, [switch]$NoHint)
    if (-not $Doc) { Write-Error-Json 'BAD_ARGS' 'open needs -Doc' }
    $Request['show'] = $true
    if ($Bound.ContainsKey('Show')) { $Request['show'] = [bool]$Bound['Show'] }
    $Request['readOnly'] = [bool]$ReadOnly
    if ($NoHint) { $Request['noHint'] = $true }
}

# save / close: identical in all three apps.
function Add-WocSaveFields {
    param($Request, [string]$Doc, [string]$As)
    if (-not $Doc) { Write-Error-Json 'BAD_ARGS' 'save needs -Doc' }
    if ($As) { $Request['as'] = Resolve-FullPath $As }
}

function Add-WocCloseFields {
    param($Request, [string]$Doc, [switch]$Save)
    if (-not $Doc) { Write-Error-Json 'BAD_ARGS' 'close needs -Doc' }
    $Request['save'] = [bool]$Save
}

# Copy the bound formatting switches into the request through an app-specific parameter->key map.
# word.ps1 also handles -Font and the -SplitWords inverse and keeps its own wrapper around this.
function Add-WocMappedFields {
    param($Bound, $Request, $Map)
    foreach ($key in $Map.Keys) {
        if (-not $Bound.ContainsKey($key)) { continue }
        $value = $Bound[$key]
        if ($value -is [System.Management.Automation.SwitchParameter]) { $value = [bool]$value }
        if ($null -ne $value -and "$value" -ne '') { $Request[$Map[$key]] = $value }
    }
}

# hint: the user-visible hint line / balloon / status bar / tray icon. Identical in all three apps;
# ppt adds the on-slide badge because that one is deck CONTENT, not a UI hint.
function Add-WocHintFields {
    param(
        $Request, [string]$Doc, [string]$HintTag, [string]$Bar, [string]$Text, [string]$Wrapper,
        [switch]$Clear, [switch]$Toast, [switch]$Sound, [switch]$Quiet, [switch]$Tray,
        [switch]$SlideMark, [switch]$NoSlideMark
    )
    if (-not $Doc) { Write-Error-Json 'BAD_ARGS' 'hint needs -Doc' }
    if ($HintTag) { $Request['tag'] = $HintTag }
    if ($Bar) { $Request['bar'] = $Bar }
    if ($Text) { $Request['text'] = $Text }
    if ($Clear) { $Request['clear'] = $true }
    if ($Toast) { $Request['toast'] = $true }
    if ($Sound) { $Request['sound'] = $true }
    if ($Quiet) { $Request['quiet'] = $true }
    # -Tray: keep a resident indicator icon; the daemon restores it via this same wrapper script.
    if ($Tray) { $Request['tray'] = $true; $Request['wrapper'] = $Wrapper }
    # ppt only: the on-slide badge is deck CONTENT (a manual Ctrl+S would save it), so it is
    # OPT-IN; -NoSlideMark explicitly removes one that is already showing.
    if ($SlideMark) { $Request['slideMark'] = $true }
    elseif ($NoSlideMark) { $Request['slideMark'] = $false }
}

# Run scripts\param-check.ps1 against the finished request (v1.18): a parameter that PowerShell
# bound but the command did not copy into the request warns on stderr instead of vanishing.
function Invoke-WocParamCheck {
    param([string]$ScriptRoot, [string]$Command, $Bound, $Request, $Aliases, $FormatFields, $CommandSkips, $DefaultValues)
    $paramCheckPath = Join-Path $ScriptRoot 'param-check.ps1'
    if (-not (Test-Path -LiteralPath $paramCheckPath)) { return }
    . $paramCheckPath
    Test-ParamConsumption -Command $Command -Bound $Bound -Request $Request -Aliases $Aliases `
        -FormatFields $FormatFields -CommandSkips $CommandSkips -DefaultValues $DefaultValues
}
