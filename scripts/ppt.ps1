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
                 'format', 'font', 'theme', 'master', 'audit', 'check', 'outline', 'align', 'zorder', 'notes', 'view', 'render', 'snapshot', 'diff', 'restore', 'save', 'close', 'quit', 'batch', 'finish', 'hint', 'serve',
                 'crop', 'smartart', 'connect', 'bg', 'anim')]
    [string]$Command,

    [string]$Doc,
    [int]$From = 1,
    [int]$To = 0,
    [int]$Count = 0,
    [string]$Items,
    # batch mode: path to a JSON file with an "ops" array of request objects
    [string]$Ops,
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
    # chart extensions (v1.29): trend line, data labels, legend position
    [switch]$Trendline,
    [switch]$DataLabels,
    [string]$LegendPos,
    # crop: how a picture is fitted into its frame
    [string]$Focus,
    [string]$Fit,
    # split layout: which half the picture takes
    [string]$Side,
    # connect: the two shapes a connector joins (the plain -From/-To are slide ranges)
    [string]$FromShape,
    [string]$ToShape,
    [string]$FromSide,
    [string]$ToSide,
    [string]$Kind,
    [string]$Arrow,
    [string]$Label,
    # bg: texture name / background image
    [switch]$Gradient,
    [string]$Texture,
    [string]$Image,
    # anim: transition + entrance effect
    [string]$Transition,
    [string]$Effect,
    [switch]$OnClick,
    [double]$Duration = 0.6,
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
# theme -Probe: test whether this deck accepts Presentation.Tags (v1.32 theme mirror)
    [switch]$Probe,
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
    [switch]$SlideMark,
    [switch]$NoSlideMark,
    [int]$TimeoutSec = 180
)

$ErrorActionPreference = 'Stop'
try { [Console]::OutputEncoding = New-Object System.Text.UTF8Encoding($false) } catch { }

# Everything shared with word.ps1 / excel.ps1 lives in woc-app-common.ps1: the JSON error
# envelope, path resolution, the items-file reader, the resident-daemon client and probe, the
# source hash + csc build, the batch command and the daemon-first / one-shot transport tail.
# Dot-sourced (not imported) so its functions and $script:daemonPipe land in THIS script scope.
. "$PSScriptRoot\woc-app-common.ps1"

# ---------------------------------------------------------------- helper binary
# Collect helper\ppt\*.cs + helper\shared\*.cs, hash them, compile on first use; returns the
# exe path and sets $script:daemonPipe. Why the hash is cached and why the pipe name follows it
# is documented in woc-app-common.ps1.
$exePath = Initialize-WocHelper -App 'ppt' -ScriptRoot $PSScriptRoot -ExeName 'WppHelper.exe'

# ---------------------------------------------------------------- request

# ---------------------------------------------------------------- batch (v1.19)
# One process start for N requests: see Invoke-WocBatchCli in woc-app-common.ps1 for the why,
# the ops-file shape ({"ops":[...]}) and the ok/failed envelope rule.
if ($Command -eq 'batch') { Invoke-WocBatchCli -Ops $Ops -ExePath $exePath }

$request = @{}
$request['cmd'] = $Command
if ($Doc) { $request['doc'] = Resolve-FullPath $Doc }

$formatFields = @{
    Font = 'font'; Size = 'size'; Bold = 'bold'; Italic = 'italic';
    Color = 'color'; Align = 'align'; Fill = 'fill'
}
function Add-Format-Fields($bound) { Add-WocMappedFields -Bound $bound -Request $request -Map $formatFields }

function Read-Items-Doc([string]$path) {
    # v1.32: canvas / grid / reuse pass ONE document object (title + items/rows + edges +
    # texts/images). Read-Items-Json unwraps an "items" key and would lose the rest, so the
    # whole object travels as a one-element array and the helper picks the document out.
    if (-not (Test-Path -LiteralPath $path)) { Write-Error-Json 'BAD_ARGS' ('Items file not found: ' + $path) }
    $raw = [System.IO.File]::ReadAllText((Resolve-FullPath $path), [System.Text.Encoding]::UTF8)
    $parsed = $raw | ConvertFrom-Json
    return ,@($parsed)
}

switch ($Command) {
    'open' {
        Add-WocOpenFields -Request $request -Bound $PSBoundParameters -Doc $Doc -ReadOnly:$ReadOnly -NoHint:$NoHint
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
        if ($Items) {
            # v1.32: canvas / grid / reuse take a whole document object, so it must not be
            # unwrapped to its "items" array (that would drop title / rows / edges)
            $lowerPreset = if ($Preset) { $Preset.ToLowerInvariant() } else { '' }
            $lowerAction = if ($Action) { $Action.ToLowerInvariant() } else { 'add' }
            if ($lowerPreset -eq 'canvas' -or $lowerPreset -eq 'grid' -or $lowerAction -eq 'reuse') {
                $request['items'] = Read-Items-Doc $Items
            } else {
                $request['items'] = Read-Items-Json $Items
            }
        }
        # slide -Action reuse: -From is the SOURCE page here (read/render/export use it as a range start)
        if ($PSBoundParameters.ContainsKey('From')) { $request['from'] = $From }
        if ($Layout -gt 0) { $request['layout'] = $Layout }
        if ($At -gt 0) { $request['at'] = $At }
        if ($Index -gt 0) { $request['index'] = $Index }
        if ($To -gt 0) { $request['to'] = $To }
        # finished layouts (v1.31): gallery -Fit contain|cover, split -Path <image> -Side left|right
        if ($Fit) { $request['fit'] = $Fit.ToLowerInvariant() }
        if ($Path) { $request['path'] = Resolve-FullPath $Path }
        if ($Side) { $request['side'] = $Side.ToLowerInvariant() }
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
        # split into volumes: each chunk becomes its own deck, the source deck is left untouched
        if (-not $Doc) { Write-Error-Json 'BAD_ARGS' 'split needs -Doc' }
        if (-not $Chunks) { Write-Error-Json 'BAD_ARGS' 'split needs -Chunks "1-20,21-45,46-80"' }
        if (-not $Out) { Write-Error-Json 'BAD_ARGS' 'split needs -Out <folder>' }
        $request['chunks'] = $Chunks
        $request['out'] = Resolve-FullPath $Out
        if ($Prefix) { $request['prefix'] = $Prefix }
        if ($DryRun) { $request['dryRun'] = $true }
    }
    'merge' {
        # merge decks: insert the -Sources slides into the target after -After (default: append)
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
            if ($Trendline) { $request['trendline'] = $true }
            if ($DataLabels) { $request['dataLabels'] = $true }
            if ($LegendPos) { $request['legendPos'] = $LegendPos.ToLowerInvariant() }
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
            if ($Name) { $request['name'] = $Name }
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
        if ($Probe) { $request['probe'] = $true }
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
        Add-WocSaveFields -Request $request -Doc $Doc -As $As
    }
    'serve' {
        # Resident helper control: -Action start|stop|status (default start). Shared body.
        Invoke-WocServe -Action $Action -ExePath $exePath
    }
    'close' {
        Add-WocCloseFields -Request $request -Doc $Doc -Save:$Save
    }
    'hint' {
        # Shared body (woc-app-common.ps1); ppt adds the opt-in on-slide badge.
        Add-WocHintFields -Request $request -Doc $Doc -HintTag $HintTag -Bar $Bar -Text $Text -Wrapper $PSCommandPath -Clear:$Clear -Toast:$Toast -Sound:$Sound -Quiet:$Quiet -Tray:$Tray -SlideMark:$SlideMark -NoSlideMark:$NoSlideMark
    }
    'finish' {
        if ($Note) { $request['note'] = $Note }
        if ($Toast) { $request['toast'] = $true }
        if ($Sound) { $request['sound'] = $true }
        if ($Quiet) { $request['quiet'] = $true }
    }
    'crop' {
        # a picture cut to a frame: cover (fill the box, cut the overflow) / contain (show it whole)
        if (-not $Doc) { Write-Error-Json 'BAD_ARGS' 'crop needs -Doc' }
        if ($Slide -lt 1) { Write-Error-Json 'BAD_ARGS' 'crop needs -Slide <n>' }
        if (-not $Shape -and $ShapeIdx -lt 1 -and -not $Path) { Write-Error-Json 'BAD_ARGS' 'crop needs -Shape <picture> (or -Path <image> to place and cut one)' }
        $request['slide'] = $Slide
        if ($Shape) { $request['shape'] = $Shape }
        if ($ShapeIdx -gt 0) { $request['shapeIdx'] = $ShapeIdx }
        if ($Path) { $request['path'] = Resolve-FullPath $Path }
        if ($Name) { $request['name'] = $Name }
        if ($Focus) { $request['focus'] = $Focus.ToLowerInvariant() }
        if ($Fit) { $request['fit'] = $Fit.ToLowerInvariant() }
        if ($L -ge 0) { $request['l'] = $L }
        if ($T -ge 0) { $request['t'] = $T }
        if ($W -gt 0) { $request['w'] = $W }
        if ($H -gt 0) { $request['h'] = $H }
    }
    'smartart' {
        # native SmartArt: -Action list (134 layouts) / add (-Preset + -Items)
        if (-not $Doc) { Write-Error-Json 'BAD_ARGS' 'smartart needs -Doc' }
        $request['action'] = if ($Action) { $Action.ToLowerInvariant() } else { 'list' }
        if ($Preset) { $request['preset'] = $Preset }
        if ($Match) { $request['match'] = $Match }
        if ($request['action'] -eq 'add') {
            if ($Slide -lt 1) { Write-Error-Json 'BAD_ARGS' 'smartart -Action add needs -Slide <n>' }
            $request['slide'] = $Slide
            if ($Items) { $request['items'] = Read-Items-Json $Items }
            elseif ($Text) { $request['text'] = $Text }
            else { Write-Error-Json 'BAD_ARGS' 'smartart -Action add needs -Items <json> (or -Text)' }
            if ($Name) { $request['name'] = $Name }
            if ($L -ge 0) { $request['l'] = $L }
            if ($T -ge 0) { $request['t'] = $T }
            if ($W -gt 0) { $request['w'] = $W }
            if ($H -gt 0) { $request['h'] = $H }
        }
    }
    'connect' {
        # a connector between two shapes: -FromShape / -ToShape (-From/-To are slide ranges)
        if (-not $Doc) { Write-Error-Json 'BAD_ARGS' 'connect needs -Doc' }
        if ($Slide -lt 1) { Write-Error-Json 'BAD_ARGS' 'connect needs -Slide <n>' }
        if (-not $FromShape) { Write-Error-Json 'BAD_ARGS' 'connect needs -FromShape <name> (-From is the slide range)' }
        if (-not $ToShape) { Write-Error-Json 'BAD_ARGS' 'connect needs -ToShape <name> (-To is the slide range)' }
        $request['slide'] = $Slide
        $request['fromShape'] = $FromShape
        $request['toShape'] = $ToShape
        if ($Kind) { $request['kind'] = $Kind.ToLowerInvariant() }
        if ($Arrow) { $request['arrow'] = $Arrow.ToLowerInvariant() }
        if ($Label) { $request['label'] = $Label }
        if ($Color) { $request['color'] = $Color }
        if ($FromSide) { $request['fromSide'] = $FromSide.ToLowerInvariant() }
        if ($ToSide) { $request['toSide'] = $ToSide.ToLowerInvariant() }
    }
    'bg' {
        # slide background: colour / gradient / generated texture / picture / reset
        if (-not $Doc) { Write-Error-Json 'BAD_ARGS' 'bg needs -Doc' }
        if ($All) { $request['all'] = $true }
        elseif ($Slide -ge 1) { $request['slide'] = $Slide }
        else { Write-Error-Json 'BAD_ARGS' 'bg needs -Slide <n> or -All' }
        if ($Color) { $request['color'] = $Color }
        if ($Gradient) { $request['gradient'] = $true }
        if ($Texture) { $request['texture'] = $Texture.ToLowerInvariant() }
        if ($Image) { $request['path'] = Resolve-FullPath $Image }
        elseif ($Path) { $request['path'] = Resolve-FullPath $Path }
        if ($Reset) { $request['reset'] = $true }
    }
    'anim' {
        # slide transition and entrance effects; the response reports what WPS read back
        if (-not $Doc) { Write-Error-Json 'BAD_ARGS' 'anim needs -Doc' }
        if ($All) { $request['all'] = $true }
        elseif ($Slide -ge 1) { $request['slide'] = $Slide }
        else { Write-Error-Json 'BAD_ARGS' 'anim needs -Slide <n> or -All' }
        if ($Transition) { $request['transition'] = $Transition.ToLowerInvariant() }
        if ($Effect) { $request['effect'] = $Effect.ToLowerInvariant() }
        if ($OnClick) { $request['onClick'] = $true }
        if ($PSBoundParameters.ContainsKey('Duration')) { $request['duration'] = $Duration }
    }
}

# ---------------------------------------------------------------- parameter consumption check (v1.18)
# Shared with the other two wrappers: any bound parameter this command did not copy into the
# request warns on stderr instead of vanishing (see scripts\param-check.ps1).
$paramAliases = @{
    HintTag = 'tag'; KeepCrop = 'resetCrop'; FitInside = 'keepSize'; SlideMark = 'slideMark'; NoSlideMark = 'slideMark'
    Image = 'path'
}
$paramCommandSkips = @{
    '*' = @('ExitCode')
    'text' = @('Expect', 'Text', 'TextFile', 'Shape', 'ShapeIdx', 'Slide')
}
Invoke-WocParamCheck -ScriptRoot $PSScriptRoot -Command $Command -Bound $PSBoundParameters -Request $request `
    -Aliases $paramAliases -FormatFields $formatFields -CommandSkips $paramCommandSkips -DefaultValues @{ From = 1; Count = 0; Index = 0; Max = 200; Slide = 0; ShapeIdx = 0; L = -1; T = -1; W = 0; H = 0 }
$json = $request | ConvertTo-Json -Depth 12 -Compress

# Transport: resident daemon first, then one-shot - both live in woc-app-common.ps1 and both hand
# back the response text, so the exit-code rules below cover both transports.
# NOTE: this variable must NOT be called $text - PowerShell variable names are case-insensitive, so
# $text IS the [string]$Text parameter and assigning $null to it stores '' instead, which makes the
# $null test below false and silently skips the one-shot fallback.
$responseText = Invoke-WocDaemonFirst -Command $Command -Json $json -TimeoutSec $TimeoutSec
if ($null -eq $responseText) { $responseText = Invoke-WocOneShot -App 'ppt' -Command $Command -ExePath $exePath -Json $json -TimeoutSec $TimeoutSec }
# -ExitCode turns `check` into a gate: a non-zero blocker count fails the call, so a script (or
# CI) can refuse to ship the deck without parsing the JSON. Applying it AFTER the transport merge
# is the point: on the warm-daemon path the old code exited 0 before this line was ever reached.
if ($ExitCode -and $responseText -match '"blocker"\s*:\s*[1-9]') { exit 3 }
exit 0
