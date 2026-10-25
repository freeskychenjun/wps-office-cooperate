# wps-office-cooperate (excel): thin wrapper around EtHelper.exe.
# Pure ASCII source (Chinese only ever travels inside JSON files, never in this script).
#
#   excel.ps1 <command> [-Doc <path>] [options]
#
# The request is written to a UTF-8 file, the helper reads it and writes a UTF-8 JSON
# response, which is echoed back unchanged. No console code page is involved anywhere.

param(
    [Parameter(Mandatory = $true, Position = 0)]
    [ValidateSet('doctor', 'status', 'open', 'read', 'search', 'selection', 'write', 'fill', 'rows', 'cols', 'sheet', 'comment', 'render', 'chart', 'autofit', 'format',
                 'snapshot', 'diff', 'restore', 'calc', 'save', 'close', 'quit', 'finish')]
    [string]$Command,

    [string]$Doc,
    [string]$Sheet,
    [string]$Range,
    [switch]$Profile,
    [string]$Items,
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
    [string]$Note,
    [int]$Max = 200,
    # format options
    [string]$Font,
    [string]$Size,
    [switch]$Bold,
    [switch]$Italic,
    [switch]$Underline,
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
    # misc
    [string]$Since,
    [string]$Cells,
    [switch]$KeepExtras,
    [string]$As,
    [switch]$Show,
    [switch]$ReadOnly,
    [switch]$Save,
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

# The cache key covers every source file, so editing any of them rebuilds.
$hashInput = ($sources | ForEach-Object { $_.Name + ':' + (Get-FileHash -LiteralPath $_.FullName -Algorithm SHA256).Hash }) -join ';'
$shaGen = [System.Security.Cryptography.SHA256]::Create()
$sha = ([System.BitConverter]::ToString($shaGen.ComputeHash([System.Text.Encoding]::UTF8.GetBytes($hashInput))) -replace '-', '').ToLowerInvariant().Substring(0, 12)
$cacheDir = Join-Path $env:LOCALAPPDATA ('wps-office-cooperate\excel-helper-' + $sha)
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
        if ($request['action'] -eq 'add') {
            if (-not $Range) { Write-Error-Json 'BAD_ARGS' 'chart add needs -Range <data range>' }
            $request['range'] = $Range
            if ($Type) { $request['type'] = $Type }
            if ($Title) { $request['title'] = $Title }
            if ($Anchor) { $request['anchor'] = $Anchor }
        }
        elseif ($request['action'] -eq 'delete') {
            if ($Index -lt 1) { Write-Error-Json 'BAD_ARGS' 'chart delete needs -Index <n> (see chart -Action list)' }
            $request['index'] = $Index
        }
    }
    'open' {
        if (-not $Doc) { Write-Error-Json 'BAD_ARGS' 'open needs -Doc' }
        $request['show'] = [bool]$Show
        $request['readOnly'] = [bool]$ReadOnly
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
    }
    'autofit' {
        if (-not $Doc) { Write-Error-Json 'BAD_ARGS' 'autofit needs -Doc' }
        if ($Sheet) { $request['sheet'] = $Sheet }
        if ($Range) { $request['range'] = $Range }
    }
    'format' {
        if (-not $Doc) { Write-Error-Json 'BAD_ARGS' 'format needs -Doc' }
        if (-not $Range) { Write-Error-Json 'BAD_ARGS' 'format needs -Range <ref>' }
        $request['range'] = $Range
        if ($Sheet) { $request['sheet'] = $Sheet }
        Add-Format-Fields $PSBoundParameters
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
    'close' {
        if (-not $Doc) { Write-Error-Json 'BAD_ARGS' 'close needs -Doc' }
        $request['save'] = [bool]$Save
    }
}

$json = $request | ConvertTo-Json -Depth 12 -Compress
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
exit 0
