# wps-office-cooperate (ppt): thin wrapper around WppHelper.exe.
# Pure ASCII source (Chinese only ever travels inside JSON files, never in this script).
#
#   ppt.ps1 <command> [-Doc <path>] [options]
#
# The request is written to a UTF-8 file, the helper reads it and writes a UTF-8 JSON
# response, which is echoed back unchanged. No console code page is involved anywhere.

param(
    [Parameter(Mandatory = $true, Position = 0)]
    [ValidateSet('doctor', 'status', 'open', 'read', 'search', 'replace', 'selection', 'slide', 'split', 'merge', 'text', 'textbox', 'image', 'picture', 'comment', 'export', 'table', 'chart', 'shape',
                 'format', 'font', 'theme', 'master', 'audit', 'check', 'outline', 'align', 'zorder', 'notes', 'view', 'render', 'snapshot', 'diff', 'restore', 'save', 'close', 'quit', 'finish', 'hint')]
    [string]$Command,

    [string]$Doc,
    [int]$From = 1,
    [int]$To = 0,
    [int]$Count = 0,
    [string]$Items,
    [string]$Data,
    [int]$Slide = 0,
    [string]$Shape,
    [int]$ShapeIdx = 0,
    [string]$Expect,
    [string]$Text,
    [string]$TextFile,
    # replace options
    [string]$Find,
    [string]$Replace,
    [string]$Scope,
    [string]$Slides,
    [switch]$Regex,
    [switch]$DryRun,
    # check / outline options
    [string]$Profile,
    [string]$Rules,
    [string]$MinSeverity,
    [string]$Config,
    [switch]$ExitCode,
    # comment / export / picture options
    [string]$Format,
    [string]$New,
    [string]$Match,
    [string]$Ids,
    [string]$Author,
    # split / merge options
    [string]$Chunks,
    [string]$Sources,
    [string]$Prefix,
    [int]$After = -1,
    # master options
    [switch]$PageNum,
    [switch]$Date,
    [switch]$SkipCover,
    [switch]$KeepTemplatePages,
    # chart options
    [string]$XTitle,
    [string]$YTitle,
    [switch]$Legend,
    [int]$Quality = 88,
    [switch]$KeepCrop,
    [switch]$FitInside,
    [switch]$All,
    [string]$Preset,
    [string]$Title,
    [string]$Subtitle,
    [string]$Number,
    [int]$Layout = 0,
    [string]$Action,
    [int]$Index = 0,
    [int]$At = 0,
    [string]$Name,
    [string]$Path,
    [double]$L = -1,
    [double]$T = -1,
    [double]$W = 0,
    [double]$H = 0,
    [string]$Type,
    [string]$Fill,
    [string]$Line,
    [switch]$HeaderRow,
    [switch]$Clear,
    [switch]$Case,
    [switch]$Overview,
    [switch]$Apply,
    [switch]$Reset,
    [string]$Set,
    [string]$Mode,
    [string]$Shapes,
    [string]$Where,
    [int]$MinFont = 12,
    [switch]$Toast,
    [switch]$Sound,
    [switch]$Quiet,
    [switch]$Force,
    [int]$Max = 50,
    # format options (text shapes)
    [string]$Font,
    [string]$Size,
    [switch]$Bold,
    [switch]$Italic,
    [string]$Color,
    [string]$Align,
    # render options
    [string]$Out,
    [string]$Width,
    # misc
    [string]$Since,
    [string]$Tag,
    [string]$Note,
    [string]$As,
    [switch]$Show,
    [switch]$ReadOnly,
    [switch]$Save,
    # takeover hint (see hint command)
    [string]$HintTag,
    [string]$Bar,
    [switch]$NoHint,
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
    try { return [System.IO.Path]::GetFullPath((Join-Path (Get-Location) $p)) } catch { return $p }
}

# ---------------------------------------------------------------- helper binary

$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$sourceDir = Join-Path $scriptRoot 'helper\ppt'
if (-not (Test-Path -LiteralPath $sourceDir)) { Write-Error-Json 'NO_HELPER' ("Helper folder not found: " + $sourceDir) }
$helperSources = @(Get-ChildItem -LiteralPath $sourceDir -Filter *.cs | Sort-Object Name)
if ($helperSources.Count -eq 0) { Write-Error-Json 'NO_HELPER' ('No .cs in ' + $sourceDir) }
# Shared sources (the tray daemon / hint plumbing) compile into every helper.
$sharedDir = Join-Path $scriptRoot 'helper\shared'
if (Test-Path -LiteralPath $sharedDir) { $helperSources += @(Get-ChildItem -LiteralPath $sharedDir -Filter *.cs | Sort-Object Name) }

# The cache key covers every source file, so editing any of them rebuilds.
$hashInput = ($helperSources | ForEach-Object { $_.Name + ':' + (Get-FileHash -LiteralPath $_.FullName -Algorithm SHA256).Hash }) -join ';'
$shaGen = [System.Security.Cryptography.SHA256]::Create()
$sha = ([System.BitConverter]::ToString($shaGen.ComputeHash([System.Text.Encoding]::UTF8.GetBytes($hashInput))) -replace '-', '').ToLowerInvariant().Substring(0, 12)
$cacheDir = Join-Path $env:LOCALAPPDATA ('wps-office-cooperate\ppt-helper-' + $sha)
$exePath = Join-Path $cacheDir 'WppHelper.exe'

if (-not (Test-Path -LiteralPath $exePath)) {
    $csc = $null
    foreach ($arch in @('Framework64', 'Framework')) {
        $candidate = Join-Path $env:SystemRoot ("Microsoft.NET\" + $arch + "\v4.0.30319\csc.exe")
        if (Test-Path -LiteralPath $candidate) { $csc = $candidate; break }
    }
    if (-not $csc) { Write-Error-Json 'NO_COMPILER' 'The .NET Framework C# compiler (csc.exe) was not found. Enable ".NET Framework 4.8" in Windows Features.' }
    New-Item -ItemType Directory -Force -Path $cacheDir | Out-Null
    $tempExe = Join-Path $cacheDir ('build-' + $PID + '.exe')
    $refs = @('-r:System.dll', '-r:System.Core.dll', '-r:System.Drawing.dll', '-r:System.Windows.Forms.dll', '/codepage:65001')
    $compileLog = & $csc -nologo -optimize+ -target:exe ("-out:" + $tempExe) $refs ($helperSources | ForEach-Object { $_.FullName }) 2>&1 | Out-String
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
    Font = 'font'; Size = 'size'; Bold = 'bold'; Italic = 'italic';
    Color = 'color'; Align = 'align'; Fill = 'fill'
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
    'open' {
        if (-not $Doc) { Write-Error-Json 'BAD_ARGS' 'open needs -Doc' }
        $request['show'] = [bool]$Show
        $request['readOnly'] = [bool]$ReadOnly
        if ($NoHint) { $request['noHint'] = $true }
    }
    'read' {
        if (-not $Doc) { Write-Error-Json 'BAD_ARGS' 'read needs -Doc' }
        if ($From -gt 1) { $request['from'] = $From }
        if ($Count -gt 0) { $request['count'] = $Count }
    }
    'search' {
        if (-not $Doc) { Write-Error-Json 'BAD_ARGS' 'search needs -Doc' }
        if (-not $Text) { Write-Error-Json 'BAD_ARGS' 'search needs -Text "<what to find>"' }
        $request['text'] = $Text
        if ($Max -gt 0) { $request['max'] = $Max }
        if ($Case) { $request['case'] = $true }
    }
    'selection' {
        if (-not $Doc) { Write-Error-Json 'BAD_ARGS' 'selection needs -Doc' }
    }
    'replace' {
        # whole-deck find/replace; -Items <json> takes [{find,replace,regex?,case?,label?}]
        if (-not $Doc) { Write-Error-Json 'BAD_ARGS' 'replace needs -Doc' }
        if ($Items) { $request['items'] = Read-Items-Json $Items }
        else {
            if (-not $Find) { Write-Error-Json 'BAD_ARGS' 'replace needs -Find <text> (or -Items <json>)' }
            $request['find'] = $Find
        }
        # an omitted -Replace means "delete the match"; -Replace "" is the same thing
        if ($PSBoundParameters.ContainsKey('Replace')) { $request['replace'] = $Replace }
        if ($Scope) { $request['scope'] = $Scope }
        if ($Slides) { $request['slides'] = $Slides }
        if ($Regex) { $request['regex'] = $true }
        if ($Case) { $request['case'] = $true }
        if ($DryRun) { $request['dryRun'] = $true }
        if ($PSBoundParameters.ContainsKey('Max')) { $request['max'] = $Max }
    }
    'slide' {
        if (-not $Doc) { Write-Error-Json 'BAD_ARGS' 'slide needs -Doc' }
        $request['action'] = if ($Action) { $Action.ToLowerInvariant() } else { 'add' }
        if ($Preset) { $request['preset'] = $Preset.ToLowerInvariant() }
        if ($Title) { $request['title'] = $Title }
        if ($Subtitle) { $request['subtitle'] = $Subtitle }
        if ($Number) { $request['number'] = $Number }
        if ($Items) { $request['items'] = Read-Items-Json $Items }
        if ($Layout -gt 0) { $request['layout'] = $Layout }
        if ($At -gt 0) { $request['at'] = $At }
        if ($Index -gt 0) { $request['index'] = $Index }
        if ($To -gt 0) { $request['to'] = $To }
    }
    'text' {
        if (-not $Doc) { Write-Error-Json 'BAD_ARGS' 'text needs -Doc' }
        if ($Items) { $request['items'] = Read-Items-Json $Items }
        else {
            if ($Slide -lt 1) { Write-Error-Json 'BAD_ARGS' 'text needs -Slide <n> (or -Items <json>)' }
            if (-not $Text -and -not $TextFile) { Write-Error-Json 'BAD_ARGS' 'text needs -Text/-TextFile (or -Items <json>)' }
            $one = @{ slide = $Slide }
            if ($Shape) { $one['shape'] = $Shape }
            elseif ($ShapeIdx -gt 0) { $one['shapeIdx'] = $ShapeIdx }
            if ($Expect) { $one['expect'] = $Expect }
            if ($TextFile) { $one['textFile'] = (Resolve-FullPath $TextFile) }
            else { $one['text'] = $Text }
            $request['items'] = @($one)
        }
    }
    'textbox' {
        if (-not $Doc) { Write-Error-Json 'BAD_ARGS' 'textbox needs -Doc' }
        if ($Slide -lt 1) { Write-Error-Json 'BAD_ARGS' 'textbox needs -Slide <n>' }
        $request['slide'] = $Slide
        if ($TextFile) { $request['textFile'] = (Resolve-FullPath $TextFile) }
        elseif ($Text) { $request['text'] = $Text }
        else { Write-Error-Json 'BAD_ARGS' 'textbox needs -Text or -TextFile' }
        if ($Name) { $request['name'] = $Name }
        if ($L -ge 0) { $request['l'] = $L }
        if ($T -ge 0) { $request['t'] = $T }
        if ($W -gt 0) { $request['w'] = $W }
        if ($H -gt 0) { $request['h'] = $H }
        Add-Format-Fields $PSBoundParameters
    }
    'image' {
        if (-not $Doc) { Write-Error-Json 'BAD_ARGS' 'image needs -Doc' }
        if ($Slide -lt 1) { Write-Error-Json 'BAD_ARGS' 'image needs -Slide <n>' }
        if (-not $Path) { Write-Error-Json 'BAD_ARGS' 'image needs -Path <picture file>' }
        $request['slide'] = $Slide
        $request['path'] = Resolve-FullPath $Path
        if ($L -ge 0) { $request['l'] = $L }
        if ($T -ge 0) { $request['t'] = $T }
        if ($W -gt 0) { $request['w'] = $W }
        if ($H -gt 0) { $request['h'] = $H }
    }
    'split' {
        # 拆成分册：每个 chunk 存成一个独立文稿，源稿不动
        if (-not $Doc) { Write-Error-Json 'BAD_ARGS' 'split needs -Doc' }
        if (-not $Chunks) { Write-Error-Json 'BAD_ARGS' 'split needs -Chunks "1-20,21-45,46-80"' }
        if (-not $Out) { Write-Error-Json 'BAD_ARGS' 'split needs -Out <folder>' }
        $request['chunks'] = $Chunks
        $request['out'] = Resolve-FullPath $Out
        if ($Prefix) { $request['prefix'] = $Prefix }
        if ($DryRun) { $request['dryRun'] = $true }
    }
    'merge' {
        # 合并多稿：把 -Sources 的页依次插到目标稿（-After 之后，默认追加到末尾）
        if (-not $Doc) { Write-Error-Json 'BAD_ARGS' 'merge needs -Doc <target deck>' }
        if ($Items) {
            # a JSON file may hold either {"sources":[...]} or {"items":[{...}]}
            $rawItems = [System.IO.File]::ReadAllText((Resolve-FullPath $Items), [System.Text.Encoding]::UTF8)
            $parsedItems = $rawItems | ConvertFrom-Json
            if ($parsedItems.PSObject.Properties.Name -contains 'sources') {
                $request['sourcesList'] = @($parsedItems.sources)
            }
            else {
                $request['items'] = Read-Items-Json $Items
            }
        }
        elseif ($Sources) { $request['sources'] = $Sources }
        else { Write-Error-Json 'BAD_ARGS' 'merge needs -Sources "a.pptx,b.pptx" (or -Items <json: {"sources":[...]}>)' }
        if ($After -ge 0) { $request['after'] = $After }
        if ($DryRun) { $request['dryRun'] = $true }
    }
    'picture' {
        # list / replace every picture on the deck (unit logo, refreshed renderings)
        if (-not $Doc) { Write-Error-Json 'BAD_ARGS' 'picture needs -Doc' }
        $request['action'] = if ($Action) { $Action.ToLowerInvariant() } else { 'list' }
        if ($New) { $request['new'] = Resolve-FullPath $New }
        if ($Match) { $request['match'] = $Match }
        if ($Ids) { $request['ids'] = $Ids }
        if ($Slides) { $request['slides'] = $Slides }
        if ($DryRun) { $request['dryRun'] = $true }
        # the helper already defaults to "keep the old geometry, reset the crop";
        # these two switches are the opt-outs, not the defaults
        if ($FitInside) { $request['keepSize'] = $false }
        if ($KeepCrop) { $request['resetCrop'] = $false }
    }
    'comment' {
        # review remarks anchored to a shape; WPS keeps the text, not our Left/Top
        if (-not $Doc) { Write-Error-Json 'BAD_ARGS' 'comment needs -Doc' }
        $request['action'] = if ($Action) { $Action.ToLowerInvariant() } else { 'add' }
        if ($Items) { $request['items'] = Read-Items-Json $Items }
        else {
            if ($Slide -gt 0) { $request['slide'] = $Slide }
            if ($Shape) { $request['shape'] = $Shape }
            if ($TextFile) { $request['textFile'] = (Resolve-FullPath $TextFile) }
            elseif ($Text) { $request['text'] = $Text }
        }
        if ($Author) { $request['author'] = $Author }
        if ($Index -gt 0) { $request['index'] = $Index }
        if ($All) { $request['all'] = $true }
        if ($Clear) { $request['clear'] = $true }
    }
    'export' {
        # PDF / PNG / JPG / markdown hand-over
        if (-not $Doc) { Write-Error-Json 'BAD_ARGS' 'export needs -Doc' }
        if (-not $Out) { Write-Error-Json 'BAD_ARGS' 'export needs -Out <path or folder>' }
        $request['format'] = if ($Format) { $Format.ToLowerInvariant() } else { 'pdf' }
        $request['out'] = Resolve-FullPath $Out
        if ($Slides) { $request['slides'] = $Slides }
        if ($Width) { $request['width'] = [int]$Width }
        if ($Quality -ne 88) { $request['quality'] = $Quality }
    }
    'master' {
        # list / batch footer+page numbers / switch layout / rebase onto a template
        if (-not $Doc) { Write-Error-Json 'BAD_ARGS' 'master needs -Doc' }
        $request['action'] = if ($Action) { $Action.ToLowerInvariant() } else { 'list' }
        if ($Slides) { $request['slides'] = $Slides }
        if ($Text) { $request['text'] = $Text }
        if ($TextFile) { $request['text'] = [System.IO.File]::ReadAllText((Resolve-FullPath $TextFile), [System.Text.Encoding]::UTF8).TrimEnd() }
        if ($PageNum) { $request['pageNum'] = $true }
        if ($Date) { $request['date'] = $true }
        if ($SkipCover) { $request['skipCover'] = $true }
        if ($Clear) { $request['clear'] = $true }
        if ($Layout -gt 0) { $request['layout'] = $Layout }
        if ($Path) { $request['path'] = Resolve-FullPath $Path }
        if ($Out) { $request['out'] = Resolve-FullPath $Out }
        if ($KeepTemplatePages) { $request['keepTemplatePages'] = $true }
        if ($DryRun) { $request['dryRun'] = $true }
    }
    'chart' {
        # native editable chart; the data goes through the chart's embedded workbook
        if (-not $Doc) { Write-Error-Json 'BAD_ARGS' 'chart needs -Doc' }
        $request['action'] = if ($Action) { $Action.ToLowerInvariant() } else { 'add' }
        if ($request['action'] -eq 'add') {
            if ($Slide -lt 1) { Write-Error-Json 'BAD_ARGS' 'chart needs -Slide <n>' }
            if (-not $Data) { Write-Error-Json 'BAD_ARGS' 'chart needs -Data <json: {"categories":[..],"series":[{"name":..,"values":[..]}]}>' }
            $request['slide'] = $Slide
            $request['data'] = Resolve-FullPath $Data
            $request['type'] = if ($Type) { $Type.ToLowerInvariant() } else { 'column' }
            if ($Title) { $request['title'] = $Title }
            if ($XTitle) { $request['xtitle'] = $XTitle }
            if ($YTitle) { $request['ytitle'] = $YTitle }
            if ($Legend) { $request['legend'] = $true }
            if ($L -ge 0) { $request['l'] = $L }
            if ($T -ge 0) { $request['t'] = $T }
            if ($W -gt 0) { $request['w'] = $W }
            if ($H -gt 0) { $request['h'] = $H }
        }
        if ($Slides) { $request['slides'] = $Slides }
    }
    'table' {
        if (-not $Doc) { Write-Error-Json 'BAD_ARGS' 'table needs -Doc' }
        if ($Slide -lt 1) { Write-Error-Json 'BAD_ARGS' 'table needs -Slide <n>' }
        if (-not $Data) { Write-Error-Json 'BAD_ARGS' 'table needs -Data <json file: {"rows":[["h1","h2"],...]}>' }
        $request['slide'] = $Slide
        $request['data'] = Resolve-FullPath $Data
        if ($L -ge 0) { $request['l'] = $L }
        if ($T -ge 0) { $request['t'] = $T }
        if ($W -gt 0) { $request['w'] = $W }
        if ($H -gt 0) { $request['h'] = $H }
        $request['headerRow'] = [bool]$HeaderRow
        if ($Size) { $request['size'] = $Size }
    }
    'shape' {
        if (-not $Doc) { Write-Error-Json 'BAD_ARGS' 'shape needs -Doc' }
        if ($Slide -lt 1) { Write-Error-Json 'BAD_ARGS' 'shape needs -Slide <n>' }
        $request['slide'] = $Slide
        $request['action'] = if ($Action) { $Action.ToLowerInvariant() } else { 'add' }
        if ($request['action'] -eq 'delete') {
            if (-not $Name) { Write-Error-Json 'BAD_ARGS' 'shape -Action delete needs -Name <shape name>' }
            $request['name'] = $Name
        }
        else {
            if ($Type) { $request['type'] = $Type }
            if ($L -ge 0) { $request['l'] = $L }
            if ($T -ge 0) { $request['t'] = $T }
            if ($W -gt 0) { $request['w'] = $W }
            if ($H -gt 0) { $request['h'] = $H }
            if ($Line) { $request['line'] = $Line }
            if ($Text) { $request['text'] = $Text }
            Add-Format-Fields $PSBoundParameters
        }
    }
    'format' {
        if (-not $Doc) { Write-Error-Json 'BAD_ARGS' 'format needs -Doc' }
        if ($Slide -lt 1) { Write-Error-Json 'BAD_ARGS' 'format needs -Slide <n>' }
        if (-not $Shape -and $ShapeIdx -lt 1) { Write-Error-Json 'BAD_ARGS' 'format needs -Shape <name> or -ShapeIdx <n>' }
        $request['slide'] = $Slide
        if ($Shape) { $request['shape'] = $Shape }
        if ($ShapeIdx -gt 0) { $request['shapeIdx'] = $ShapeIdx }
        Add-Format-Fields $PSBoundParameters
    }
    'font' {
        # whole deck (or a slide range): every text run, groups and table cells included.
        if (-not $Doc) { Write-Error-Json 'BAD_ARGS' 'font needs -Doc' }
        if (-not $Font) { Write-Error-Json 'BAD_ARGS' 'font needs -Font <name>' }
        $request['font'] = $Font
        if ($From -gt 1) { $request['from'] = $From }
        if ($To -gt 0) { $request['to'] = $To }
    }
    'notes' {
        if (-not $Doc) { Write-Error-Json 'BAD_ARGS' 'notes needs -Doc' }
        if ($Slide -lt 1) { Write-Error-Json 'BAD_ARGS' 'notes needs -Slide <n>' }
        $request['slide'] = $Slide
        if ($Clear) { $request['clear'] = $true }
        elseif ($TextFile) { $request['textFile'] = (Resolve-FullPath $TextFile) }
        elseif ($Text) { $request['text'] = $Text }
        else { Write-Error-Json 'BAD_ARGS' 'notes needs -Text/-TextFile (or -Clear)' }
    }
    'view' {
        if (-not $Doc) { Write-Error-Json 'BAD_ARGS' 'view needs -Doc' }
        if ($Slide -lt 1) { Write-Error-Json 'BAD_ARGS' 'view needs -Slide <n>' }
        $request['slide'] = $Slide
    }
    'theme' {
        if (-not $Doc) { Write-Error-Json 'BAD_ARGS' 'theme needs -Doc' }
        if ($Set) { $request['set'] = Resolve-FullPath $Set }
        if ($Apply) { $request['apply'] = $true }
        if ($Reset) { $request['reset'] = $true }
    }
    'audit' {
        if (-not $Doc) { Write-Error-Json 'BAD_ARGS' 'audit needs -Doc' }
        if ($From -gt 1) { $request['from'] = $From }
        if ($To -gt 0) { $request['to'] = $To }
        if ($MinFont -ne 12) { $request['minFont'] = $MinFont }
    }
    'check' {
        # read-only delivery gate: severity-graded findings, configurable rule library
        if (-not $Doc) { Write-Error-Json 'BAD_ARGS' 'check needs -Doc' }
        if ($Profile) { $request['profile'] = $Profile }
        if ($Rules) { $request['rules'] = $Rules }
        if ($MinSeverity) { $request['minSeverity'] = $MinSeverity }
        if ($Config) { $request['config'] = Resolve-FullPath $Config }
        if ($Out) { $request['out'] = Resolve-FullPath $Out }
    }
    'outline' {
        # titles + bullets + notes -> markdown/json (feeds the "write me a script" flow)
        if (-not $Doc) { Write-Error-Json 'BAD_ARGS' 'outline needs -Doc' }
        if ($Out) { $request['out'] = Resolve-FullPath $Out }
    }
    'align' {
        if (-not $Doc) { Write-Error-Json 'BAD_ARGS' 'align needs -Doc' }
        if ($Slide -lt 1) { Write-Error-Json 'BAD_ARGS' 'align needs -Slide <n>' }
        if (-not $Mode) { Write-Error-Json 'BAD_ARGS' 'align needs -Mode left|right|top|bottom|center-h|same-width|same-height|distribute-h|distribute-v' }
        $request['slide'] = $Slide
        $request['mode'] = $Mode.ToLowerInvariant()
        if ($Shapes) { $request['shapes'] = $Shapes }
    }
    'zorder' {
        if (-not $Doc) { Write-Error-Json 'BAD_ARGS' 'zorder needs -Doc' }
        if ($Slide -lt 1) { Write-Error-Json 'BAD_ARGS' 'zorder needs -Slide <n>' }
        if (-not $Name -and $ShapeIdx -lt 1) { Write-Error-Json 'BAD_ARGS' 'zorder needs -Name <shape> or -ShapeIdx <n>' }
        $request['slide'] = $Slide
        if ($Name) { $request['name'] = $Name }
        if ($ShapeIdx -gt 0) { $request['shapeIdx'] = $ShapeIdx }
        $request['where'] = if ($Where) { $Where.ToLowerInvariant() } else { 'front' }
    }
    'restore' {
        if (-not $Doc) { Write-Error-Json 'BAD_ARGS' 'restore needs -Doc' }
        if ($Since) { $request['since'] = $Since } else { $request['since'] = 'last' }
        if ($Shapes) { $request['shapes'] = $Shapes }
    }
    'render' {
        if (-not $Doc) { Write-Error-Json 'BAD_ARGS' 'render needs -Doc' }
        if (-not $Out) { Write-Error-Json 'BAD_ARGS' 'render needs -Out <directory>' }
        if ($From -gt 1) { $request['from'] = $From }
        if ($To -gt 0) { $request['to'] = $To }
        $request['out'] = Resolve-FullPath $Out
        $request['width'] = if ($Width) { [int]$Width } else { 1100 }
        if ($Overview) { $request['overview'] = $true }
    }
    'snapshot' {
        if (-not $Doc) { Write-Error-Json 'BAD_ARGS' 'snapshot needs -Doc' }
        if ($Tag) { $request['tag'] = $Tag }
    }
    'diff' {
        if (-not $Doc) { Write-Error-Json 'BAD_ARGS' 'diff needs -Doc' }
        if ($Since) { $request['since'] = $Since } else { $request['since'] = 'last' }
    }
    'save' {
        if (-not $Doc) { Write-Error-Json 'BAD_ARGS' 'save needs -Doc' }
        if ($As) { $request['as'] = Resolve-FullPath $As }
    }
    'close' {
        if (-not $Doc) { Write-Error-Json 'BAD_ARGS' 'close needs -Doc' }
        $request['save'] = [bool]$Save
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
        if ($Note) { $request['note'] = $Note }
        if ($Toast) { $request['toast'] = $true }
        if ($Sound) { $request['sound'] = $true }
        if ($Quiet) { $request['quiet'] = $true }
    }
}

$json = $request | ConvertTo-Json -Depth 12 -Compress
$work = Join-Path $env:TEMP ('wps-office-cooperate-ppt-' + [guid]::NewGuid().ToString('N'))
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
# -ExitCode turns `check` into a gate: a non-zero blocker count fails the call, so a
# script (or CI) can refuse to ship the deck without parsing the JSON.
if ($ExitCode -and $text -match '"blocker"\s*:\s*[1-9]') { exit 3 }
exit 0
