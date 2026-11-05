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
                 'snapshot', 'diff', 'restore', 'calc', 'save', 'close', 'quit', 'batch', 'finish', 'hint', 'serve',
                 'clean', 'fillblanks', 'dedupe', 'deletewhere', 'sort', 'filter', 'copy', 'addcol', 'pivot', 'freeze', 'cformat', 'validate', 'table',
                 'replace')]
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
    # ---- data work: clean / fillblanks / dedupe / deletewhere / sort / filter / copy / addcol
    [string]$Column,
    [string]$Columns,
    [string]$To,
    [string]$With,
    [string]$Cond,
    [string]$InValues,
    [string]$By,
    [string]$Order,
    [string]$Criteria,
    [string]$Criteria2,
    [string]$Join,
    [switch]$NoHeader,
    [string]$Casing,
    [string]$Remove,
    [int]$Year = 0,
    [string]$Map,
    [switch]$NoSpaces,
    [string]$Header,
    [string]$After,
    [switch]$Values,
    # ---- pivot
    [string]$Agg,
    [string]$Rows,
    [string]$Filters,
    [string]$Group,
    [string]$SortBy,
    # ---- replace (in-memory block replace; see reference/excel-quirks.md for why Range.Replace is not used)
    [string]$Find,
    [string]$Replace,
    [switch]$Whole,
    [switch]$MatchCase,
    [switch]$Regex,
    [switch]$DryRun,
    # ---- finished look + chart extras
    [string]$Rule,
    [string]$Value2,
    [string]$List,
    [string]$Style,
    [string]$Y2Title,
    [string]$Secondary,
    [string]$Labels,
    [string]$Legend,
    [string]$AxisMin,
    [string]$AxisMax,
    # ---- render -Overview
    [switch]$Overview,
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

# Everything shared with word.ps1 / ppt.ps1 lives in woc-app-common.ps1: the JSON error
# envelope, path resolution, the items-file reader, the resident-daemon client and probe, the
# source hash + csc build, the batch command and the daemon-first / one-shot transport tail.
# Dot-sourced (not imported) so its functions and $script:daemonPipe land in THIS script scope.
. "$PSScriptRoot\woc-app-common.ps1"

# ---------------------------------------------------------------- helper binary
# Collect helper\excel\*.cs + helper\shared\*.cs, hash them, compile on first use; returns the
# exe path and sets $script:daemonPipe. Why the hash is cached and why the pipe name follows it
# is documented in woc-app-common.ps1.
$exePath = Initialize-WocHelper -App 'excel' -ScriptRoot $PSScriptRoot -ExeName 'EtHelper.exe'

# ---------------------------------------------------------------- request

# ---------------------------------------------------------------- batch (v1.19)
# One process start for N requests: see Invoke-WocBatchCli in woc-app-common.ps1 for the why,
# the ops-file shape ({"ops":[...]}) and the ok/failed envelope rule.
if ($Command -eq 'batch') { Invoke-WocBatchCli -Ops $Ops -ExePath $exePath }

$request = @{}
$request['cmd'] = $Command
if ($Doc) { $request['doc'] = Resolve-FullPath $Doc }

$formatFields = @{
    Font = 'font'; Size = 'size'; Bold = 'bold'; Italic = 'italic'; Underline = 'underline';
    Color = 'color'; FillColor = 'fillColor'; HAlign = 'halign'; VAlign = 'valign';
    NumberFormat = 'numberFormat'; ColWidth = 'colWidth'; RowHeight = 'rowHeight'; Borders = 'borders'; Wrap = 'wrap'
}
function Add-Format-Fields($bound) { Add-WocMappedFields -Bound $bound -Request $request -Map $formatFields }

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
            # v1.20: combo / secondary axis / labels / legend / axis scale, and a name to find it by later.
            if ($Y2Title) { $request['y2title'] = $Y2Title }
            if ($Secondary) { $request['secondary'] = @($Secondary -split ',') }
            if ($Labels) { $request['labels'] = $Labels }
            if ($Legend) { $request['legend'] = $Legend }
            if ($NumberFormat) { $request['numberFormat'] = $NumberFormat }
            if ($AxisMin) { $request['min'] = $AxisMin }
            if ($AxisMax) { $request['max'] = $AxisMax }
            if ($Name) { $request['name'] = $Name }
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
    'clean' {
        if (-not $Doc) { Write-Error-Json 'BAD_ARGS' 'clean needs -Doc' }
        if ($Sheet) { $request['sheet'] = $Sheet }
        if ($Range) { $request['range'] = $Range }
        if ($Column) { $request['column'] = $Column }
        if ($Columns) { $request['columns'] = @($Columns -split ',') }
        if ($To) { $request['to'] = $To.ToLowerInvariant() }
        if ($Casing) { $request['case'] = $Casing.ToLowerInvariant() }
        if ($Remove) { $request['remove'] = @($Remove -split ',') }
        if ($Year -gt 0) { $request['year'] = $Year }
        if ($NoSpaces) { $request['noSpaces'] = $true }
        if ($Map) { $request['map'] = ($Map | ConvertFrom-Json) }
        if ($NumberFormat) { $request['numberFormat'] = $NumberFormat }
    }
    'fillblanks' {
        if (-not $Doc) { Write-Error-Json 'BAD_ARGS' 'fillblanks needs -Doc' }
        if (-not $With) { Write-Error-Json 'BAD_ARGS' 'fillblanks needs -With <value|above|mean|median|mode>' }
        if ($Sheet) { $request['sheet'] = $Sheet }
        if ($Range) { $request['range'] = $Range }
        if ($Column) { $request['column'] = $Column }
        if ($Columns) { $request['columns'] = @($Columns -split ',') }
        $fillNumber = 0.0
        if ([double]::TryParse($With, [ref]$fillNumber)) { $request['with'] = $fillNumber } else { $request['with'] = $With }
    }
    'dedupe' {
        if (-not $Doc) { Write-Error-Json 'BAD_ARGS' 'dedupe needs -Doc' }
        if ($Sheet) { $request['sheet'] = $Sheet }
        if ($Range) { $request['range'] = $Range }
        if ($Column) { $request['column'] = $Column }
        if ($Columns) { $request['columns'] = @($Columns -split ',') }
    }
    'deletewhere' {
        if (-not $Doc) { Write-Error-Json 'BAD_ARGS' 'deletewhere needs -Doc' }
        if ($Sheet) { $request['sheet'] = $Sheet }
        if ($Range) { $request['range'] = $Range }
        if ($Column) { $request['column'] = $Column }
        if ($Cond) { $request['is'] = $Cond }
        if ($PSBoundParameters.ContainsKey('Value')) { $request['value'] = $Value }
        if ($InValues) { $request['in'] = @($InValues -split ',') }
    }
    'sort' {
        if (-not $Doc) { Write-Error-Json 'BAD_ARGS' 'sort needs -Doc' }
        if (-not $By -and -not $Column) { Write-Error-Json 'BAD_ARGS' 'sort needs -By "column[:desc],column2"' }
        if ($Sheet) { $request['sheet'] = $Sheet }
        if ($Range) { $request['range'] = $Range }
        if ($Column) { $request['column'] = $Column }
        if ($By) { $request['by'] = @($By -split ',') }
        if ($Order) { $request['order'] = $Order.ToLowerInvariant() }
        if ($NoHeader) { $request['header'] = $false }
    }
    'filter' {
        if (-not $Doc) { Write-Error-Json 'BAD_ARGS' 'filter needs -Doc' }
        if ($Sheet) { $request['sheet'] = $Sheet }
        if ($Range) { $request['range'] = $Range }
        $act = if ($Action) { $Action.ToLowerInvariant() } else { 'apply' }
        if ($act -ne 'apply' -and $act -ne 'on') { $request['action'] = $act }
        if ($act -eq 'apply' -or $act -eq 'on') {
            if (-not $Column) { Write-Error-Json 'BAD_ARGS' 'filter needs -Column <the column to filter on>' }
            $request['column'] = $Column
            if ($InValues) { $request['in'] = @($InValues -split ',') }
            if ($Criteria) { $request['criteria'] = $Criteria }
            if ($Criteria2) { $request['criteria2'] = $Criteria2 }
            if ($Join) { $request['join'] = $Join.ToLowerInvariant() }
            if (-not $InValues -and -not $Criteria) { Write-Error-Json 'BAD_ARGS' 'filter needs -InValues "a,b" or -Criteria "<100"' }
        }
    }
    'copy' {
        if (-not $Doc) { Write-Error-Json 'BAD_ARGS' 'copy needs -Doc' }
        if (-not $Range) { Write-Error-Json 'BAD_ARGS' 'copy needs -Range <source>' }
        if (-not $To) { Write-Error-Json 'BAD_ARGS' 'copy needs -To <target cell or sheet!cell>' }
        if ($Sheet) { $request['sheet'] = $Sheet }
        $request['range'] = $Range
        $request['to'] = $To
        if ($Values) { $request['values'] = $true }
    }
    'addcol' {
        if (-not $Doc) { Write-Error-Json 'BAD_ARGS' 'addcol needs -Doc' }
        if (-not $Header) { Write-Error-Json 'BAD_ARGS' 'addcol needs -Header "<column name>"' }
        if (-not $Formula -and -not $PSBoundParameters.ContainsKey('Value')) { Write-Error-Json 'BAD_ARGS' 'addcol needs -Formula "=..." or -Value <x>' }
        if ($Sheet) { $request['sheet'] = $Sheet }
        if ($Range) { $request['range'] = $Range }
        $request['header'] = $Header
        if ($PSBoundParameters.ContainsKey('Formula')) { $request['formula'] = $Formula } else { $request['value'] = $Value }
        if ($After) { $request['after'] = $After }
        if ($NumberFormat) { $request['numberFormat'] = $NumberFormat }
        if ($Values) { $request['values'] = $true }
    }
    'pivot' {
        if (-not $Doc) { Write-Error-Json 'BAD_ARGS' 'pivot needs -Doc' }
        if (-not $Agg) { Write-Error-Json 'BAD_ARGS' 'pivot needs -Agg "field:sum[:caption],field2:count"' }
        if ($Sheet) { $request['sheet'] = $Sheet }
        if ($Range) { $request['range'] = $Range }
        if ($Rows) { $request['rows'] = @($Rows -split ',') }
        if ($Columns) { $request['columns'] = @($Columns -split ',') }
        if ($Filters) { $request['filters'] = @($Filters -split ',') }
        $aggList = @()
        foreach ($piece in ($Agg -split ',')) {
            if (-not $piece.Trim()) { continue }
            $bits = $piece -split ':'
            $item = @{ field = $bits[0].Trim() }
            if ($bits.Count -gt 1 -and $bits[1].Trim()) { $item['fn'] = $bits[1].Trim() }
            if ($bits.Count -gt 2 -and $bits[2].Trim()) { $item['name'] = $bits[2].Trim() }
            $aggList += $item
        }
        $request['values'] = $aggList
        if ($Name) { $request['name'] = $Name }
        if ($To) { $request['to'] = $To }
        if ($Group) {
            $bits = $Group -split ':'
            $request['group'] = @{ field = $bits[0].Trim(); by = if ($bits.Count -gt 1 -and $bits[1].Trim()) { $bits[1].Trim() } else { 'month' } }
        }
        if ($SortBy) {
            $bits = $SortBy -split ':'
            $sortSpec = @{ field = $bits[0].Trim() }
            if ($bits.Count -gt 1 -and $bits[1].Trim()) {
                $tail = $bits[1].Trim()
                if ($tail -match '^(a|d)') { $sortSpec['order'] = $tail.ToLowerInvariant() } else { $sortSpec['by'] = $tail }
            }
            $request['sort'] = $sortSpec
        }
    }
    'replace' {
        if (-not $Doc) { Write-Error-Json 'BAD_ARGS' 'replace needs -Doc' }
        if (-not $Find) { Write-Error-Json 'BAD_ARGS' 'replace needs -Find "<old text>"' }
        $request['find'] = $Find
        # -Replace omitted = delete the match. An empty string is a valid replacement, so the
        # presence of the parameter (not its value) decides.
        if ($PSBoundParameters.ContainsKey('Replace')) { $request['replace'] = $Replace }
        if ($Sheet) { $request['sheet'] = $Sheet }
        if ($Range) { $request['range'] = $Range }
        if ($Column) { $request['column'] = $Column }
        if ($Whole) { $request['whole'] = $true }
        # -Case is the search switch; -MatchCase is the replace spelling of the same thing.
        if ($MatchCase -or $Case) { $request['case'] = $true }
        if ($Regex) { $request['regex'] = $true }
        if ($DryRun) { $request['dryRun'] = $true }
        # -Max has a default of 200 for search, so only forward it when it was really given
        # (an unasked-for cap would silently stop a large replace).
        if ($PSBoundParameters.ContainsKey('Max')) { $request['max'] = $Max }
        if ($All) { $request['all'] = $true }
    }
    'freeze' {
        if (-not $Doc) { Write-Error-Json 'BAD_ARGS' 'freeze needs -Doc' }
        if ($Sheet) { $request['sheet'] = $Sheet }
        if ($Action) { $request['action'] = $Action.ToLowerInvariant() }
        if ($Cell) { $request['cell'] = $Cell }
    }
    'cformat' {
        if (-not $Doc) { Write-Error-Json 'BAD_ARGS' 'cformat needs -Doc' }
        if (-not $Range) { Write-Error-Json 'BAD_ARGS' 'cformat needs -Range' }
        if (-not $Rule) { Write-Error-Json 'BAD_ARGS' 'cformat needs -Rule (>, <, >=, <=, =, <>, between, top, bottom, duplicate, unique, blank, formula, colorscale, databar, clear)' }
        if ($Sheet) { $request['sheet'] = $Sheet }
        $request['range'] = $Range
        $request['rule'] = $Rule
        if ($PSBoundParameters.ContainsKey('Value')) { $request['value'] = $Value }
        if ($Value2) { $request['value2'] = $Value2 }
        if ($FillColor) { $request['fillColor'] = $FillColor }
        if ($Color) { $request['color'] = $Color }
        if ($PSBoundParameters.ContainsKey('Bold')) { $request['bold'] = [bool]$Bold }
    }
    'validate' {
        if (-not $Doc) { Write-Error-Json 'BAD_ARGS' 'validate needs -Doc' }
        if (-not $Range) { Write-Error-Json 'BAD_ARGS' 'validate needs -Range' }
        if (-not $List -and -not $InValues) { Write-Error-Json 'BAD_ARGS' 'validate needs -List "a,b,c"' }
        if ($Sheet) { $request['sheet'] = $Sheet }
        $request['range'] = $Range
        $listText = if ($List) { $List } else { $InValues }
        $request['list'] = @($listText -split ',')
    }
    'table' {
        if (-not $Doc) { Write-Error-Json 'BAD_ARGS' 'table needs -Doc' }
        if ($Sheet) { $request['sheet'] = $Sheet }
        if ($Range) { $request['range'] = $Range }
        if ($Name) { $request['name'] = $Name }
        if ($Style) { $request['style'] = $Style }
    }
    'open' {
        Add-WocOpenFields -Request $request -Bound $PSBoundParameters -Doc $Doc -ReadOnly:$ReadOnly -NoHint:$NoHint
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
        # -Overview: every sheet of the workbook as one thumbnail sheet (a workbook read at a glance).
        if ($Overview) { $request['overview'] = $true }
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
        # Shared body (woc-app-common.ps1); -Tray tells the daemon which wrapper to re-run.
        Add-WocHintFields -Request $request -Doc $Doc -HintTag $HintTag -Bar $Bar -Text $Text -Wrapper $PSCommandPath -Clear:$Clear -Toast:$Toast -Sound:$Sound -Quiet:$Quiet -Tray:$Tray
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
    }
    'calc' {
        if (-not $Doc) { Write-Error-Json 'BAD_ARGS' 'calc needs -Doc' }
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
}

# ---------------------------------------------------------------- parameter consumption check (v1.18)
# Shared with word.ps1 / ppt.ps1: any bound parameter that this command did not copy into the
# request warns on stderr instead of vanishing (the -Sheet drop that sent charts to the wrong
# sheet was exactly this failure mode). See scripts\param-check.ps1.
$paramAliases = @{
    TextFile = 'text'; HintTag = 'tag'
    NoBold = 'bold'; NoItalic = 'italic'; NoUnderline = 'underline'; NoWrap = 'wrap'
    # data work / finished look (v1.20): names that do not follow the "lower-case the first letter" rule
    InValues = 'in'; Cond = 'is'; Casing = 'case'; NoHeader = 'header'; Agg = 'values'
    SortBy = 'sort'; AxisMin = 'min'; AxisMax = 'max'
    MatchCase = 'case'
}
$paramCommandSkips = @{
    # write builds a nested item, so these never appear at the top level of the request
    'write' = @('Cell', 'Value', 'Formula', 'Expect', 'Sheet')
}
Invoke-WocParamCheck -ScriptRoot $PSScriptRoot -Command $Command -Bound $PSBoundParameters -Request $request `
    -Aliases $paramAliases -FormatFields $formatFields -CommandSkips $paramCommandSkips -DefaultValues @{ Index = 0; Count = 0; Max = 200; Height = 0 }
$json = $request | ConvertTo-Json -Depth 12 -Compress
# Transport: resident daemon first, then one-shot - both live in woc-app-common.ps1 and both hand
# back the response text, so any wrapper-level exit-code rule covers both transports.
# NOTE: this variable must NOT be called $text - PowerShell variable names are case-insensitive, so
# $text IS the [string]$Text parameter and assigning $null to it stores '' instead, which makes the
# $null test below false and silently skips the one-shot fallback.
# v1.36: if the resident helper exited with its hint lease (see shared/PipeServer.cs), start a
# fresh one now - a paused session must not fall back to the slow one-shot path for every command.
# The probe is a 300 ms pipe connect against a pipe that usually exists, so this costs ~nothing.
if ($env:WOC_NO_DAEMON -ne '1') { $null = Start-Daemon $exePath }
$responseText = Invoke-WocDaemonFirst -Command $Command -Json $json -TimeoutSec $TimeoutSec
if ($null -eq $responseText) { $responseText = Invoke-WocOneShot -App 'excel' -Command $Command -ExePath $exePath -Json $json -TimeoutSec $TimeoutSec }
exit 0
