# wps-office-cooperate (excel): thin wrapper around EtHelper.exe.
# Pure ASCII source (Chinese only ever travels inside JSON files, never in this script).
#
#   excel.ps1 <command> [-Doc <path>] [options]
#
# The request is written to a UTF-8 file, the helper reads it and writes a UTF-8 JSON
# response, which is echoed back unchanged. No console code page is involved anywhere.

param(
    [Parameter(Mandatory = $true, Position = 0)]
    [ValidateSet('doctor', 'status', 'open', 'new', 'read', 'search', 'selection', 'write', 'fill', 'rows', 'cols', 'sheet', 'comment', 'render', 'chart', 'autofit', 'format',
                 'snapshot', 'diff', 'restore', 'calc', 'save', 'close', 'quit', 'batch', 'finish', 'hint', 'serve')]
    [string]$Command,

    [string]$Doc,
    [string]$Sheet,
    [string]$Range,
    [switch]$Profile,
    [string]$Items,
    # batch mode: path to a JSON file with an "ops" array of request objects
    [string]$Ops,
    [string]$Cell,
    [string]$Value,
    [string]$Formula,
    [string]$Expect,
    [string]$Action,
    [string]$Name,
    [string]$NewName,
    [string]$Text,
    [string]$TextFile,
    [int]$Index = 0,
    [switch]$All,
    [switch]$Delete,
    [switch]$Case,
    [switch]$Toast,
    [switch]$Sound,
    [switch]$Quiet,
    [switch]$Force,
    [int]$Count = 0,
    [string]$Type,
    [string]$Title,
    [string]$Anchor,
    [string]$XTitle,
    [string]$YTitle,
    [int]$Height = 0,
    [string]$Note,
    [int]$Max = 200,
    # format options
    [string]$Font,
    [string]$Size,
    [switch]$Bold,
    [switch]$Italic,
    [switch]$Underline,
    # Explicit "turn it off" switches: "-Bold:$false" does not survive `powershell -File` argument
    # parsing (the value arrives as the string "False"), so turning a format flag OFF needs its
    # own switch (reported 2026-11-01: there was no way to reset a cell back to regular).
    [switch]$NoBold,
    [switch]$NoItalic,
    [switch]$NoUnderline,
    [switch]$NoWrap,
    [string]$Color,
    [string]$FillColor,
    [string]$HAlign,
    [string]$VAlign,
    [string]$NumberFormat,
    [string]$ColWidth,
    [string]$RowHeight,
    [string]$Borders,
    [switch]$Wrap,
    # render options
    [string]$Out,
    [string]$Width,
    [switch]$IncludeCharts,
    # misc
    [string]$Since,
    [string]$Cells,
    [switch]$KeepExtras,
    [string]$As,
    [switch]$Show,
    [switch]$ReadOnly,
    [switch]$Overwrite,
    [switch]$Save,
    # takeover hint (see hint command)
    [string]$HintTag,
    [string]$Bar,
    [switch]$NoHint,
    [switch]$Clear,
    # resident tray indicator (opt-in)
    [switch]$Tray,
    [int]$TimeoutSec = 180
)

$ErrorActionPreference = 'Stop'
try { [Console]::OutputEncoding = New-Object System.Text.UTF8Encoding($false) } catch { }

function Write-Error-Json([string]$code, [string]$message) {
    $payload = @{ ok = $false; error = @{ code = $code; message = $message } }
    [Console]::Out.Write(($payload | ConvertTo-Json -Depth 6 -Compress))
    exit 1
}

function Resolve-FullPath([string]$p) {
    try { return [System.IO.Path]::GetFullPath($p) } catch { return $p }
}

# ---------------------------------------------------------------- helper binary

$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$sourceDir = Join-Path $scriptRoot 'helper\excel'
if (-not (Test-Path -LiteralPath $sourceDir)) { Write-Error-Json 'NO_HELPER' ("Helper folder not found: " + $sourceDir) }
$sources = @(Get-ChildItem -LiteralPath $sourceDir -Filter *.cs | Sort-Object Name)
if ($sources.Count -eq 0) { Write-Error-Json 'NO_HELPER' ('No .cs in ' + $sourceDir) }
# Shared sources (the tray daemon / hint plumbing) compile into every helper.
$sharedDir = Join-Path $scriptRoot 'helper\shared'
if (Test-Path -LiteralPath $sharedDir) { $sources += @(Get-ChildItem -LiteralPath $sharedDir -Filter *.cs | Sort-Object Name) }

# ---------------------------------------------------------------- resident daemon (v1.15)
# Same transport as the Word wrapper: a warm helper keeps the COM application and answers request
# JSON over a named pipe, so a command no longer pays "start exe + attach to WPS + find document".
# The pipe name is derived from the compiled binary's hash, so a rebuilt helper gets a new pipe.
$daemonPipe = $null

function Get-DaemonPipeName([string]$shaGen12) { return ('woc-excel-' + $shaGen12) }

function Invoke-Daemon([string]$payload, [int]$connectMs) {
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
        $hdr = New-Object byte[] 4
        $got = 0
        while ($got -lt 4) {
            $n = $pipe.Read($hdr, $got, 4 - $got)
            if ($n -le 0) { return $null }
            $got += $n
        }
        $total = [System.BitConverter]::ToInt32($hdr, 0)
        if ($total -le 0 -or $total -gt 33554432) { return $null }
        $buf = New-Object byte[] $total
        $got = 0
        while ($got -lt $total) {
            $n = $pipe.Read($buf, $got, $total - $got)
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

function Start-Daemon([string]$exePath2) {
    if ($env:WOC_NO_DAEMON -eq '1') { return $false }
    $probe = '{"cmd":"ping"}'
    if ($null -ne (Invoke-Daemon $probe 300)) { return $true }
    try {
        Start-Process -FilePath $exePath2 -ArgumentList @('--serve', $script:daemonPipe) -WindowStyle Hidden | Out-Null
    } catch {
        return $false
    }
    for ($i = 0; $i -lt 40; $i++) {
        Start-Sleep -Milliseconds 100
        if ($null -ne (Invoke-Daemon $probe 300)) { return $true }
    }
    return $false
}


# The cache key covers every source file, so editing any of them rebuilds.
# Hash of the helper sources -> compile cache dir + daemon pipe name. Get-FileHash over every
# .cs file costs ~50 ms on each command (measured 2026-11-02), so the shared helper caches it
# keyed on file count + sizes + newest mtime (any real edit changes one of them).
$wocCommonPath = Join-Path $PSScriptRoot 'woc-common.ps1'
if (Test-Path -LiteralPath $wocCommonPath) { . $wocCommonPath }
if (Get-Command Get-CachedSourceHash -ErrorAction SilentlyContinue) {
    $sha = Get-CachedSourceHash -App 'excel' -Sources $sources
} else {
    
$hashInput = ($sources | ForEach-Object { $_.Name + ':' + (Get-FileHash -LiteralPath $_.FullName -Algorithm SHA256).Hash }) -join ';'
    
$shaGen = [System.Security.Cryptography.SHA256]::Create()
    
$sha = ([System.BitConverter]::ToString($shaGen.ComputeHash([System.Text.Encoding]::UTF8.GetBytes($hashInput))) -replace '-', '').ToLowerInvariant().Substring(0, 12)
}
$cacheDir = Join-Path $env:LOCALAPPDATA ('wps-office-cooperate\excel-helper-' + $sha)
$script:daemonPipe = Get-DaemonPipeName $sha
$exePath = Join-Path $cacheDir 'EtHelper.exe'


if (-not (Test-Path -LiteralPath $exePath)) {
    $csc = $null
    foreach ($arch in @('Framework64', 'Framework')) {
        $candidate = Join-Path $env:SystemRoot ("Microsoft.NET\" + $arch + "\v4.0.30319\csc.exe")
        if (Test-Path -LiteralPath $candidate) { $csc = $candidate; break }
    }
    if (-not $csc) { Write-Error-Json 'NO_COMPILER' 'The .NET Framework C# compiler (csc.exe) was not found. Enable ".NET Framework 4.8" in Windows Features.' }
    New-Item -ItemType Directory -Force -Path $cacheDir | Out-Null
    $tempExe = Join-Path $cacheDir ('build-' + $PID + '.exe')
    $refs = @('-r:System.dll', '-r:System.Core.dll', '-r:System.Drawing.dll', '-r:System.Windows.Forms.dll', '-r:Microsoft.CSharp.dll', '/codepage:65001')
    $compileLog = & $csc -nologo -optimize+ -target:exe ("-out:" + $tempExe) $refs ($sources | ForEach-Object { $_.FullName }) 2>&1 | Out-String
    if (-not (Test-Path -LiteralPath $tempExe)) {
        Write-Error-Json 'COMPILE_FAILED' ("csc failed: " + $compileLog.Trim())
    }
    Move-Item -LiteralPath $tempExe -Destination $exePath -Force
}

# ---------------------------------------------------------------- request

# ---------------------------------------------------------------- batch (v1.19)
# Run many requests over ONE pipe connection from a single process start. The CLI path pays a
# PowerShell start + wrapper parse (~220 ms) per command, while a pipe round trip is 1-15 ms
# (measured 2026-11-02), so "one page = 15 commands" becomes one tool call instead of fifteen.
#   woc.ps1 excel batch -Ops ops.json
#   ops.json: {"ops":[{"cmd":"...","doc":"...","sheet":"..."}, ...]}
if ($Command -eq 'batch') {
    if (-not $Ops) { Write-Error-Json 'BAD_ARGS' 'batch needs -Ops <ops.json> (an object with an ops array)' }
    $opsPath = Resolve-FullPath $Ops
    if (-not (Test-Path -LiteralPath $opsPath)) { Write-Error-Json 'BAD_ARGS' ('ops file not found: ' + $opsPath) }
    $rawOps = [System.IO.File]::ReadAllText($opsPath, [System.Text.Encoding]::UTF8)
    try { $parsedOps = $rawOps | ConvertFrom-Json }
    catch { Write-Error-Json 'BAD_ARGS' ('ops file is not valid JSON: ' + $_.Exception.Message) }
    if ($parsedOps.PSObject.Properties.Name -contains 'ops') { $opList = @($parsedOps.ops) } else { $opList = @($parsedOps) }
    if ($opList.Count -eq 0) { Write-Error-Json 'BAD_ARGS' 'batch ops is empty' }
    if (-not (Get-Command Invoke-WocBatch -ErrorAction SilentlyContinue)) {
        $wocCommonPath2 = Join-Path $PSScriptRoot 'woc-common.ps1'
        if (Test-Path -LiteralPath $wocCommonPath2) { . $wocCommonPath2 }
    }
    $batchResult = Invoke-WocBatch -PipeName $script:daemonPipe -ExePath $exePath -Ops $opList
    [Console]::Out.Write(($batchResult | ConvertTo-Json -Depth 12 -Compress))
    exit 0
}

$request = @{}
$request['cmd'] = $Command
if ($Doc) { $request['doc'] = Resolve-FullPath $Doc }

$formatFields = @{
    Font = 'font'; Size = 'size'; Bold = 'bold'; Italic = 'italic'; Underline = 'underline';
    Color = 'color'; FillColor = 'fillColor'; HAlign = 'halign'; VAlign = 'valign';
    NumberFormat = 'numberFormat'; ColWidth = 'colWidth'; RowHeight = 'rowHeight'; Borders = 'borders'; Wrap = 'wrap'
}
function Add-Format-Fields($bound) {
    foreach ($key in $formatFields.Keys) {
        if (-not $bound.ContainsKey($key)) { continue }
        $value = $bound[$key]
        if ($value -is [System.Management.Automation.SwitchParameter]) { $value = [bool]$value }
        if ($null -ne $value -and "$value" -ne '') { $request[$formatFields[$key]] = $value }
    }
}

function Read-Items-Json([string]$path) {
    if (-not (Test-Path -LiteralPath $path)) { Write-Error-Json 'BAD_ARGS' ('Items file not found: ' + $path) }
    $raw = [System.IO.File]::ReadAllText((Resolve-FullPath $path), [System.Text.Encoding]::UTF8)
    $parsed = $raw | ConvertFrom-Json
    if ($parsed.PSObject.Properties.Name -contains 'items') { return ,@($parsed.items) }
    return ,@($parsed)
}

switch ($Command) {
    'search' {
        if (-not $Doc) { Write-Error-Json 'BAD_ARGS' 'search needs -Doc' }
        if (-not $Text) { Write-Error-Json 'BAD_ARGS' 'search needs -Text "<what to find>"' }
        $request['text'] = $Text
        if ($Sheet) { $request['sheet'] = $Sheet }
        if ($Max -gt 0) { $request['max'] = $Max }
        if ($Case) { $request['case'] = $true }
    }
    'selection' {
        if (-not $Doc) { Write-Error-Json 'BAD_ARGS' 'selection needs -Doc (the workbook whose window the user selected in)' }
    }
    'rows' {
        if (-not $Doc) { Write-Error-Json 'BAD_ARGS' 'rows needs -Doc' }
        if ($Index -lt 1) { Write-Error-Json 'BAD_ARGS' 'rows needs -Index <1-based row number>' }
        $request['action'] = if ($Action) { $Action.ToLowerInvariant() } else { 'insert' }
        $request['index'] = $Index
        if ($Count -gt 0) { $request['count'] = $Count } else { $request['count'] = 1 }
        if ($Sheet) { $request['sheet'] = $Sheet }
    }
    'cols' {
        if (-not $Doc) { Write-Error-Json 'BAD_ARGS' 'cols needs -Doc' }
        if ($Index -lt 1) { Write-Error-Json 'BAD_ARGS' 'cols needs -Index <1-based column number>' }
        $request['action'] = if ($Action) { $Action.ToLowerInvariant() } else { 'insert' }
        $request['index'] = $Index
        if ($Count -gt 0) { $request['count'] = $Count } else { $request['count'] = 1 }
        if ($Sheet) { $request['sheet'] = $Sheet }
    }
    'chart' {
        if (-not $Doc) { Write-Error-Json 'BAD_ARGS' 'chart needs -Doc' }
        $request['action'] = if ($Action) { $Action.ToLowerInvariant() } else { 'add' }
        # v1.16: -Sheet was silently dropped here, so the chart always landed on the ActiveSheet
        # and list/delete could not find it again (reported 2026-11-01). Forward it for every action.
        if ($Sheet) { $request['sheet'] = $Sheet }
        if ($request['action'] -eq 'add') {
            if (-not $Range) { Write-Error-Json 'BAD_ARGS' 'chart add needs -Range <data range>' }
            $request['range'] = $Range
            if ($Type) { $request['type'] = $Type }
            if ($Title) { $request['title'] = $Title }
            if ($Anchor) { $request['anchor'] = $Anchor }
            if ($Width) { $request['width'] = [int]$Width }
            if ($Height) { $request['height'] = [int]$Height }
            if ($XTitle) { $request['xtitle'] = $XTitle }
            if ($YTitle) { $request['ytitle'] = $YTitle }
        }
        elseif ($request['action'] -eq 'list') {
            if ($All) { $request['all'] = $true }
        }
        elseif ($request['action'] -eq 'delete' -or $request['action'] -eq 'move') {
            if ($Index -lt 1) { Write-Error-Json 'BAD_ARGS' ('chart ' + $request['action'] + ' needs -Index <n> (see chart -Action list)') }
            $request['index'] = $Index
            if ($request['action'] -eq 'move') {
                if ($Anchor) { $request['anchor'] = $Anchor }
                if ($Width) { $request['width'] = [int]$Width }
                if ($Height) { $request['height'] = [int]$Height }
            }
        }
    }
    'open' {
        if (-not $Doc) { Write-Error-Json 'BAD_ARGS' 'open needs -Doc' }
        # A takeover must be visible: default to showing the window (an earlier hidden open left the
        # user staring at nothing while the helper edited a window they could not see). Pass -Show:$false
        # explicitly if a hidden session is really what you want.
        $request['show'] = $true
        if ($PSBoundParameters.ContainsKey('Show')) { $request['show'] = [bool]$Show }
        $request['readOnly'] = [bool]$ReadOnly
        if ($NoHint) { $request['noHint'] = $true }
    }
    'new' {
        if (-not $Doc) { Write-Error-Json 'BAD_ARGS' 'new needs -Doc <path of the workbook to create>' }
        $request['show'] = [bool]$Show
        if ($Overwrite) { $request['overwrite'] = $true }
        if ($NoHint) { $request['noHint'] = $true }
    }
    'read' {
        if (-not $Doc) { Write-Error-Json 'BAD_ARGS' 'read needs -Doc' }
        if ($Sheet) { $request['sheet'] = $Sheet }
        if ($Range) { $request['range'] = $Range }
        if ($Profile) { $request['profile'] = $true }
    }
    'write' {
        if (-not $Doc) { Write-Error-Json 'BAD_ARGS' 'write needs -Doc' }
        if ($Items) { $request['items'] = Read-Items-Json $Items }
        elseif ($Cell) {
            $one = @{ cell = $Cell }
            if ($Sheet) { $one['sheet'] = $Sheet }
            if ($PSBoundParameters.ContainsKey('Formula')) { $one['formula'] = $Formula }
            elseif ($PSBoundParameters.ContainsKey('Value')) { $one['value'] = $Value }
            else { Write-Error-Json 'BAD_ARGS' 'write needs -Value or -Formula' }
            if ($Expect) { $one['expect'] = $Expect }
            $request['items'] = @($one)
        }
        else { Write-Error-Json 'BAD_ARGS' 'write needs -Items <json file> or -Cell with -Value/-Formula' }
    }
    'fill' {
        if (-not $Doc) { Write-Error-Json 'BAD_ARGS' 'fill needs -Doc' }
        if (-not $Range) { Write-Error-Json 'BAD_ARGS' 'fill needs -Range <single column, e.g. D2:D5000>' }
        if (-not $Formula) { Write-Error-Json 'BAD_ARGS' 'fill needs -Formula "=..."' }
        $request['range'] = $Range
        $request['formula'] = $Formula
        if ($Sheet) { $request['sheet'] = $Sheet }
    }
    'sheet' {
        if (-not $Doc) { Write-Error-Json 'BAD_ARGS' 'sheet needs -Doc' }
        $request['action'] = if ($Action) { $Action.ToLowerInvariant() } else { 'list' }
        if ($Name) { $request['name'] = $Name }
        if ($NewName) { $request['newName'] = $NewName }
    }
    'comment' {
        if (-not $Doc) { Write-Error-Json 'BAD_ARGS' 'comment needs -Doc' }
        $act = if ($Action) { $Action.ToLowerInvariant() } else { 'add' }
        $request['action'] = $act
        if ($act -eq 'add') {
            if (-not $Cell) { Write-Error-Json 'BAD_ARGS' 'comment -Action add needs -Cell <ref>' }
            $request['cell'] = $Cell
            if ($Sheet) { $request['sheet'] = $Sheet }
            if ($TextFile) {
                if (-not (Test-Path -LiteralPath $TextFile)) { Write-Error-Json 'BAD_ARGS' ('Text file not found: ' + $TextFile) }
                $request['text'] = [System.IO.File]::ReadAllText((Resolve-FullPath $TextFile), [System.Text.Encoding]::UTF8)
            }
            elseif ($Text) { $request['text'] = $Text }
            else { Write-Error-Json 'BAD_ARGS' 'comment needs -Text or -TextFile' }
        }
        elseif ($act -eq 'delete') {
            if ($All) { $request['all'] = $true }
            elseif ($Index -gt 0) { $request['index'] = $Index }
            else { Write-Error-Json 'BAD_ARGS' 'comment -Action delete needs -Index <n> or -All' }
        }
    }
    'render' {
        if (-not $Doc) { Write-Error-Json 'BAD_ARGS' 'render needs -Doc' }
        if (-not $Out) { Write-Error-Json 'BAD_ARGS' 'render needs -Out <directory>' }
        if ($Sheet) { $request['sheet'] = $Sheet }
        if ($Range) { $request['range'] = $Range }
        $request['out'] = Resolve-FullPath $Out
        $request['width'] = if ($Width) { [int]$Width } else { 1100 }
        if ($IncludeCharts) { $request['includeCharts'] = $true }
    }
    'autofit' {
        if (-not $Doc) { Write-Error-Json 'BAD_ARGS' 'autofit needs -Doc' }
        if ($Sheet) { $request['sheet'] = $Sheet }
        if ($Range) { $request['range'] = $Range }
    }
    'format' {
        if (-not $Doc) { Write-Error-Json 'BAD_ARGS' 'format needs -Doc' }
        if ($Items) {
            # Batch form (v1.18): each JSON item carries its own range + format options, so one
            # call can lay out a whole sheet with per-block differences.
            $request['items'] = Read-Items-Json $Items
            if ($Sheet) { $request['sheet'] = $Sheet }
        }
        else {
            if (-not $Range) { Write-Error-Json 'BAD_ARGS' 'format needs -Range <ref> (semicolon for several blocks) or -Items <json>' }
            $request['range'] = $Range
            if ($Sheet) { $request['sheet'] = $Sheet }
            Add-Format-Fields $PSBoundParameters
            if ($NoBold) { $request['bold'] = $false }
            if ($NoItalic) { $request['italic'] = $false }
            if ($NoUnderline) { $request['underline'] = $false }
            if ($NoWrap) { $request['wrap'] = $false }
        }
    }
    'snapshot' {
        if (-not $Doc) { Write-Error-Json 'BAD_ARGS' 'snapshot needs -Doc' }
        if ($Sheet) { $request['sheets'] = $Sheet }
        if ($Out) { $request['out'] = Resolve-FullPath $Out }
    }
    'diff' {
        if (-not $Doc) { Write-Error-Json 'BAD_ARGS' 'diff needs -Doc' }
        if ($Since) { $request['since'] = $Since } else { $request['since'] = 'last' }
    }
    'restore' {
        if (-not $Doc) { Write-Error-Json 'BAD_ARGS' 'restore needs -Doc' }
        if ($Since) { $request['since'] = $Since } else { $request['since'] = 'last' }
        if ($Cells) { $request['cells'] = $Cells }
        if ($KeepExtras) { $request['keepExtras'] = $true }
    }
    'hint' {
        if (-not $Doc) { Write-Error-Json 'BAD_ARGS' 'hint needs -Doc' }
        if ($HintTag) { $request['tag'] = $HintTag }
        if ($Bar) { $request['bar'] = $Bar }
        if ($Text) { $request['text'] = $Text }
        if ($Clear) { $request['clear'] = $true }
        if ($Toast) { $request['toast'] = $true }
        if ($Sound) { $request['sound'] = $true }
        if ($Quiet) { $request['quiet'] = $true }
        # -Tray: keep a resident indicator icon; the daemon restores via this same script.
        if ($Tray) { $request['tray'] = $true; $request['wrapper'] = $PSCommandPath }
    }
    'finish' {
        if (-not $Doc) { Write-Error-Json 'BAD_ARGS' 'finish needs -Doc' }
        if ($Note) { $request['note'] = $Note }
        if ($Cell) { $request['cell'] = $Cell }
        if ($Sheet) { $request['sheet'] = $Sheet }
        if ($Toast) { $request['toast'] = $true }
        if ($Sound) { $request['sound'] = $true }
        if ($Quiet) { $request['quiet'] = $true }
        if ($Force) { $request['force'] = $true }
    }    'calc' {
        if (-not $Doc) { Write-Error-Json 'BAD_ARGS' 'calc needs -Doc' }
    }
    'save' {
        if (-not $Doc) { Write-Error-Json 'BAD_ARGS' 'save needs -Doc' }
        if ($As) { $request['as'] = Resolve-FullPath $As }
    }
    'serve' {
        $act = if ($Action) { $Action.ToLowerInvariant() } else { 'start' }
        if ($act -eq 'stop') {
            $null = Invoke-Daemon '{"cmd":"daemon-stop"}' 800
            $serveOut = @{ ok = $true; cmd = 'serve'; result = @{ action = 'stop'; pipe = $script:daemonPipe; note = 'stop requested; the daemon exits after answering' } }
            [Console]::Out.Write(($serveOut | ConvertTo-Json -Depth 6 -Compress))
            exit 0
        }
        if ($act -eq 'status') {
            $alive = ($null -ne (Invoke-Daemon '{"cmd":"ping"}' 300))
            $serveOut = @{ ok = $true; cmd = 'serve'; result = @{ action = 'status'; running = $alive; pipe = $script:daemonPipe } }
            [Console]::Out.Write(($serveOut | ConvertTo-Json -Depth 6 -Compress))
            exit 0
        }
        $okStart = Start-Daemon $exePath
        if ($okStart) {
            $serveOut = @{ ok = $true; cmd = 'serve'; result = @{ action = 'start'; started = $true; pipe = $script:daemonPipe } }
        } else {
            $serveOut = @{ ok = $false; error = @{ code = 'DAEMON_FAILED'; message = 'the resident helper did not come up; commands still work through the one-shot path' } }
        }
        [Console]::Out.Write(($serveOut | ConvertTo-Json -Depth 6 -Compress))
        exit 0
    }
    'close' {
        if (-not $Doc) { Write-Error-Json 'BAD_ARGS' 'close needs -Doc' }
        $request['save'] = [bool]$Save
    }
}

# ---------------------------------------------------------------- parameter consumption check (v1.18)
# Shared with word.ps1 / ppt.ps1: any bound parameter that this command did not copy into the
# request warns on stderr instead of vanishing (the -Sheet drop that sent charts to the wrong
# sheet was exactly this failure mode). See scripts\param-check.ps1.
$paramAliases = @{
    TextFile = 'text'; HintTag = 'tag'
    NoBold = 'bold'; NoItalic = 'italic'; NoUnderline = 'underline'; NoWrap = 'wrap'
}
$paramCommandSkips = @{
    # write builds a nested item, so these never appear at the top level of the request
    'write' = @('Cell', 'Value', 'Formula', 'Expect', 'Sheet')
}
$paramCheckPath = Join-Path $PSScriptRoot 'param-check.ps1'
if (Test-Path -LiteralPath $paramCheckPath) {
    . $paramCheckPath
    $paramDefaultValues = @{ Index = 0; Count = 0; Max = 200; Height = 0 }
Test-ParamConsumption -Command $Command -Bound $PSBoundParameters -Request $request `
        -Aliases $paramAliases -FormatFields $formatFields -CommandSkips $paramCommandSkips -DefaultValues $paramDefaultValues
}
$json = $request | ConvertTo-Json -Depth 12 -Compress
if ($Command -ne 'serve' -and $env:WOC_NO_DAEMON -ne '1') {
    $daemonText = Invoke-Daemon $json 900
    if ($null -ne $daemonText) {
        [Console]::Out.Write($daemonText)
        if ($daemonText -match '"ok"\s*:\s*false') { exit 2 }
        exit 0
    }
}
$work = Join-Path $env:TEMP ('wps-office-cooperate-excel-' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Force -Path $work | Out-Null
$requestPath = Join-Path $work 'request.json'
$responsePath = Join-Path $work 'response.json'
[System.IO.File]::WriteAllText($requestPath, $json, (New-Object System.Text.UTF8Encoding($false)))

# ---------------------------------------------------------------- run

$p = Start-Process -FilePath $exePath -ArgumentList @($requestPath, $responsePath) -PassThru -NoNewWindow
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
if ($text -match '"ok"\s*:\s*false') { exit 2 }
if ($Command -eq 'open') { $null = Start-Daemon $exePath }
exit 0
