# wps-office-cooperate (word): thin wrapper around WpsHelper.exe.
# Pure ASCII source (Chinese only ever travels inside JSON files, never in this script).
#
#   word.ps1 <command> [-Doc <path>] [options]
#
# The request is written to a UTF-8 file, the helper reads it and writes a UTF-8 JSON
# response, which is echoed back unchanged. No console code page is involved anywhere.

param(
    [Parameter(Mandatory = $true, Position = 0)]
    [ValidateSet('doctor', 'status', 'open', 'read', 'search', 'selection', 'undo', 'revise', 'comment', 'render', 'font',
                 'format', 'style', 'page', 'table', 'captions', 'caption', 'picture', 'toc', 'reference', 'view', 'fields', 'cell', 'insert', 'revisions', 'snapshot', 'diff', 'handoff', 'audit', 'unit', 'units',
                 'formula', 'footnote', 'fillgaps',
                 'save', 'close', 'quit', 'finish', 'hint', 'textdiag', 'batch', 'serve')]
    [string]$Command,

    [string]$Doc,
    [int]$From = 1,
    [int]$Count = 0,
    [switch]$Tables,
    [string]$Items,
    # batch mode: path to a JSON file with an "ops" array of request objects
    [string]$Ops,
    [int]$Para = 0,
    [int]$To = 0,
    [string]$Expect,
    [string]$Anchor,
    [string]$Text,
    [string]$TextFile,
    [string]$Name,
    [string]$Style,
    [string]$Match,
    [string]$Only,
    [switch]$SkipTableCells,
    [string]$Font,
    [string]$LatinFont,
    [string]$Size,
    [switch]$Bold,
    [switch]$Italic,
    [switch]$Underline,
    [string]$Color,
    [string]$Align,
    [string]$IndentChars,
    [string]$LeftIndent,
    [string]$FirstLineIndent,
    [string]$LineSpacing,
    [string]$SpaceBefore,
    [string]$SpaceAfter,
    [switch]$Super,
    [switch]$Sub,
    [switch]$PageBreakBefore,
    [switch]$Track,
    [switch]$NoNumbering,
    [switch]$NoSplitWords,
    # the negative of -NoSplitWords: a SwitchParameter cannot be given as "-NoSplitWords:$false" through woc.ps1
    [switch]$SplitWords,
    [int]$Row = 0,
    [int]$Col = 0,
    [string]$Where,
    [int]$Index = 0,
    [string]$Tag,
    [string]$Since,
    [string]$Note,
    [string]$HintTag,
    [string]$Bar,
    [switch]$NoHint,
    [switch]$Clear,
    [switch]$Toast,
    [switch]$Sound,
    [switch]$Quiet,
    [switch]$Delete,
    [int]$Max = 200,
    [switch]$Slow,
    [string]$Paper,
    [string]$Orientation,
    [string]$Top,
    [string]$Bottom,
    [string]$Left,
    [string]$Right,
    [string]$HeaderText,
    [string]$FooterText,
    [string]$PageNumber,
    [string]$PageNumberStart,
    [switch]$All,
    [switch]$Case,
    [string]$Data,
    [int]$Table = 0,
    [string]$Borders,
    [string]$Widths,
    [switch]$HeaderRow,
    [string]$Caption,
    [string]$Autofit,
    [string]$RowHeight,
    [switch]$KeepTogether,
    # table -Action plan|normalize (v1.33): renumber plan file, phase switches, caption spacing,
    # width policy after the borders were written, and a test hook for per-table isolation
    [string]$Plan,
    [switch]$NoCaptions,
    [switch]$NoBorders,
    [switch]$NoHeaderRepeat,
    [int]$Gap = -1,
    [string]$Fit,
    [switch]$Verify,
    # caption move: allow the cut-paste method on WPS (off by default: it froze WPS on 2026-11-03)
    [switch]$CutPaste,
    [int]$FailTable = 0,
    # fillgaps: report only / how many paragraphs may move in front of one block / empty band threshold
    [switch]$DryRun,
    [int]$MaxMoves = 8,
    [int]$MinGapPt = 0,
    [string]$Label,
    [string]$Title,
    [switch]$Above,
    [string]$Path,
    [string]$Image,
    [string]$Height,
    [int]$CropLeft = -1,
    [int]$CropRight = -1,
    [int]$CropTop = -1,
    [int]$CropBottom = -1,
    [string]$Action,
    # formula: LaTeX source, numbering and inline/display switch
    [string]$Latex,
    [switch]$Display,
    [string]$Number,
    [string]$Find,
    # render: one contact sheet per 6 pages instead of one picture per page
    [switch]$Overview,
    [int]$Passes = 0,
    [switch]$Force,
    [switch]$Apply,
    [switch]$Normalize,
    [string]$Units,
    [string]$Numbers,
    [string]$Heading,
    [int]$Levels = 3,
    [int]$AtPara = 0,
    [int]$Page = 1,
    [string]$Out,
    [string]$Width,
    [string]$As,
    [switch]$Show,
    [switch]$Save,
    [switch]$ReadOnly,
    # resident tray indicator (opt-in)
    [switch]$Tray,
    [int]$TimeoutSec = 180
)

$ErrorActionPreference = 'Stop'
try { [Console]::OutputEncoding = New-Object System.Text.UTF8Encoding($false) } catch { }

# Everything shared with excel.ps1 / ppt.ps1 lives in woc-app-common.ps1: the JSON error
# envelope, path resolution, the items-file reader, the resident-daemon client and probe, the
# source hash + csc build, the batch command and the daemon-first / one-shot transport tail.
# Dot-sourced (not imported) so its functions and $script:daemonPipe land in THIS script scope.
. "$PSScriptRoot\woc-app-common.ps1"

# ---------------------------------------------------------------- helper binary
# Collect helper\word\*.cs + helper\shared\*.cs, hash them, compile on first use; returns the
# exe path and sets $script:daemonPipe. Why the hash is cached and why the pipe name follows it
# is documented in woc-app-common.ps1.
$exePath = Initialize-WocHelper -App 'word' -ScriptRoot $PSScriptRoot -ExeName 'WpsHelper.exe'

# ---------------------------------------------------------------- request

# ---------------------------------------------------------------- batch (v1.19)
# One process start for N requests: see Invoke-WocBatchCli in woc-app-common.ps1 for the why,
# the ops-file shape ({"ops":[...]}) and the ok/failed envelope rule.
if ($Command -eq 'batch') { Invoke-WocBatchCli -Ops $Ops -ExePath $exePath }

$request = @{}
$request['cmd'] = $Command
if ($Doc) { $request['doc'] = Resolve-FullPath $Doc }

# Formatting fields shared by format / style / table / caption, added only when given.
$formatFields = @{
    Font = 'font'; LatinFont = 'latinFont'; Size = 'size'; Bold = 'bold'; Italic = 'italic'; Underline = 'underline';
    Color = 'color'; Align = 'align'; IndentChars = 'indentChars'; LeftIndent = 'leftIndent';
    FirstLineIndent = 'firstLineIndent'; LineSpacing = 'lineSpacing'; SpaceBefore = 'spaceBefore';
    SpaceAfter = 'spaceAfter'; Super = 'super'; Sub = 'sub'; PageBreakBefore = 'pageBreakBefore'; NoSplitWords = 'noSplitWords';
    Track = 'track'
}
function Add-Format-Fields($bound) {
    foreach ($key in $formatFields.Keys) {
        if (-not $bound.ContainsKey($key)) { continue }
        $value = $bound[$key]
        if ($value -is [System.Management.Automation.SwitchParameter]) { $value = [bool]$value }
        $request[$formatFields[$key]] = $value
    }
    if ($bound.ContainsKey('Font')) { $request['font'] = $Font }
    # -SplitWords is the way to ASK for splitting: "-NoSplitWords:$false" cannot survive woc.ps1's
    # argument round trip (it arrives as the string "false" and PS refuses to bind it to a switch).
    if ($bound.ContainsKey('SplitWords')) { $request['noSplitWords'] = -not [bool]$bound['SplitWords'] }
}

# 'units' is an alias of 'unit'. The two cases used to be byte-identical copies; they now share
# one body. (A comma-separated switch label - 'unit', 'units' { } - is NOT valid in Windows
# PowerShell 5.1: Parser gives "missing statement block in switch clause", so the alias is mapped
# here instead. The request still carries the command exactly as the caller wrote it.)
$commandKey = if ($Command -eq 'units') { 'unit' } else { $Command }
switch ($commandKey) {
    'open' {
        Add-WocOpenFields -Request $request -Bound $PSBoundParameters -Doc $Doc -ReadOnly:$ReadOnly -NoHint:$NoHint
    }
    'read' {
        if (-not $Doc) { Write-Error-Json 'BAD_ARGS' 'read needs -Doc' }
        $request['from'] = $From
        $request['count'] = $Count   # 0 = every paragraph from -From on
        if ($SkipTableCells) { $request['skipTableCells'] = $true }
        if ($Only) { $request['only'] = $Only }
        if ($Slow) { $request['slow'] = $true }
    }
    'captions' {
        if (-not $Doc) { Write-Error-Json 'BAD_ARGS' 'captions needs -Doc' }
        if ($Label) { $request['label'] = $Label }
        if ($Max -gt 0) { $request['max'] = $Max }
    }
    'search' {
        if (-not $Doc) { Write-Error-Json 'BAD_ARGS' 'search needs -Doc' }
        if (-not $Text) { Write-Error-Json 'BAD_ARGS' 'search needs -Text "<what to find>"' }
        $request['text'] = $Text
        if ($Style) { $request['style'] = $Style }
        if ($Max -gt 0) { $request['max'] = $Max }
        if ($Case) { $request['caseSensitive'] = $true }
    }
    'selection' {
        if (-not $Doc) { Write-Error-Json 'BAD_ARGS' 'selection needs -Doc (the document whose window the user selected in)' }
    }
    'undo' {
        if (-not $Doc) { Write-Error-Json 'BAD_ARGS' 'undo needs -Doc' }
        if ($Count -gt 0) { $request['count'] = $Count } else { $request['count'] = 1 }
    }
    'revise' {
        if (-not $Doc) { Write-Error-Json 'BAD_ARGS' 'revise needs -Doc' }
        if (-not $Items) { Write-Error-Json 'BAD_ARGS' 'revise needs -Items <json file>' }
        if (-not (Test-Path -LiteralPath $Items)) { Write-Error-Json 'BAD_ARGS' ('Items file not found: ' + $Items) }
        $request['items'] = Read-Items-Json $Items
    }
    'comment' {
        if (-not $Doc) { Write-Error-Json 'BAD_ARGS' 'comment needs -Doc' }
        if ($Action -and $Action.ToLowerInvariant() -eq 'list') { $request['action'] = 'list' }
        elseif ($Items) {
            if (-not (Test-Path -LiteralPath $Items)) { Write-Error-Json 'BAD_ARGS' ('Items file not found: ' + $Items) }
            $request['items'] = Read-Items-Json $Items
        }
        elseif ($Delete) {
            $request['delete'] = $true
            if ($All) { $request['all'] = $true }
            elseif ($Index -gt 0) { $request['index'] = $Index }
            else { Write-Error-Json 'BAD_ARGS' 'comment -Delete needs -Index <n> or -All' }
        }
        else {
        if ($Para -lt 1) { Write-Error-Json 'BAD_ARGS' 'comment needs -Para <n>' }
        $request['para'] = $Para
        if ($Anchor) { $request['anchor'] = $Anchor }
        if ($TextFile) {
            if (-not (Test-Path -LiteralPath $TextFile)) { Write-Error-Json 'BAD_ARGS' ('Text file not found: ' + $TextFile) }
            $request['text'] = [System.IO.File]::ReadAllText((Resolve-FullPath $TextFile), [System.Text.Encoding]::UTF8)
        }
        elseif ($Text) { $request['text'] = $Text }
        else { Write-Error-Json 'BAD_ARGS' 'comment needs -Text or -TextFile' }
        }
    }
    'render' {
        if (-not $Doc) { Write-Error-Json 'BAD_ARGS' 'render needs -Doc' }
        if (-not $Out) { Write-Error-Json 'BAD_ARGS' 'render needs -Out <directory>' }
        $request['page'] = $Page
        $request['out'] = Resolve-FullPath $Out
        if ($Table -gt 0) { $request['table'] = $Table }
        if ($Overview) {
            # Contact sheet: 6 pages per image, at most 18 pages per call, 640 px per cell.
            $request['overview'] = $true
            $request['to'] = if ($To -gt 0) { $To } else { $Page + 5 }
            $request['width'] = if ($Width) { [int]$Width } else { 640 }
        }
        else {
            if ($To -gt 0) { $request['to'] = $To }
            $request['width'] = if ($Width) { [int]$Width } else { 1100 }
        }
    }
    'formula' {
        if (-not $Doc) { Write-Error-Json 'BAD_ARGS' 'formula needs -Doc' }
        if ($Items) {
            if (-not (Test-Path -LiteralPath $Items)) { Write-Error-Json 'BAD_ARGS' ('Items file not found: ' + $Items) }
            $request['items'] = Read-Items-Json $Items
        }
        else {
            if ($Para -lt 1) { Write-Error-Json 'BAD_ARGS' 'formula needs -Para <n> (the paragraph the equation goes in), or -Items <json>' }
            $request['para'] = $Para
            if ($Expect) { $request['expect'] = $Expect }
            if ($Latex) { $request['latex'] = $Latex }
            elseif ($TextFile) {
                if (-not (Test-Path -LiteralPath $TextFile)) { Write-Error-Json 'BAD_ARGS' ('Text file not found: ' + $TextFile) }
                $request['latex'] = [System.IO.File]::ReadAllText((Resolve-FullPath $TextFile), [System.Text.Encoding]::UTF8)
            }
            elseif ($Text) { $request['latex'] = $Text }
            else { Write-Error-Json 'BAD_ARGS' 'formula needs -Latex "<latex>" (or -TextFile <path>, or -Items <json>)' }
            if ($Display) { $request['display'] = $true }
            if ($Number) { $request['number'] = $Number }
            if ($Anchor) { $request['anchor'] = $Anchor }
            if ($Where) { $request['where'] = $Where }
        }
    }
    'footnote' {
        if (-not $Doc) { Write-Error-Json 'BAD_ARGS' 'footnote needs -Doc' }
        $act = if ($Action) { $Action.ToLowerInvariant() } else { 'list' }
        $request['action'] = $act
        if ($act -eq 'insert') {
            if ($Para -lt 1) { Write-Error-Json 'BAD_ARGS' 'footnote -Action insert needs -Para <n> (the paragraph the reference mark goes in)' }
            $request['para'] = $Para
            if ($Expect) { $request['expect'] = $Expect }
            if ($Find) { $request['find'] = $Find }
            if ($Where) { $request['where'] = $Where }
            if ($TextFile) {
                if (-not (Test-Path -LiteralPath $TextFile)) { Write-Error-Json 'BAD_ARGS' ('Text file not found: ' + $TextFile) }
                $request['text'] = [System.IO.File]::ReadAllText((Resolve-FullPath $TextFile), [System.Text.Encoding]::UTF8)
            }
            elseif ($Text) { $request['text'] = $Text }
            else { Write-Error-Json 'BAD_ARGS' 'footnote -Action insert needs -Text or -TextFile' }
        }
        elseif ($act -eq 'edit') {
            if ($Index -lt 1) { Write-Error-Json 'BAD_ARGS' 'footnote -Action edit needs -Index <n> (as footnote -Action list reports it)' }
            $request['index'] = $Index
            if ($Expect) { $request['expect'] = $Expect }
            if ($TextFile) {
                if (-not (Test-Path -LiteralPath $TextFile)) { Write-Error-Json 'BAD_ARGS' ('Text file not found: ' + $TextFile) }
                $request['text'] = [System.IO.File]::ReadAllText((Resolve-FullPath $TextFile), [System.Text.Encoding]::UTF8)
            }
            elseif ($Text) { $request['text'] = $Text }
            else { Write-Error-Json 'BAD_ARGS' 'footnote -Action edit needs -Text or -TextFile' }
        }
        elseif ($act -eq 'delete') {
            if ($Index -lt 1) { Write-Error-Json 'BAD_ARGS' 'footnote -Action delete needs -Index <n>' }
            $request['index'] = $Index
            if ($Expect) { $request['expect'] = $Expect }
        }
        elseif ($act -eq 'numbering') {
            if ($Number) { $request['number'] = $Number }
        }
        elseif ($act -ne 'list') { Write-Error-Json 'BAD_ARGS' ('unknown footnote -Action "' + $act + '" (list|insert|edit|delete|numbering)') }
    }
    'font' {
        if (-not $Doc) { Write-Error-Json 'BAD_ARGS' 'font needs -Doc' }
        if (-not $Name) { Write-Error-Json 'BAD_ARGS' 'font needs -Name <font name, e.g. FangSong or SimSun>' }
        $request['name'] = $Name
        if ($Track) { $request['track'] = $true }
    }
    'format' {
        if (-not $Doc) { Write-Error-Json 'BAD_ARGS' 'format needs -Doc' }
        if ($PSBoundParameters.ContainsKey('Para') -and $Para -gt 0) { $request['para'] = $Para }
        if ($PSBoundParameters.ContainsKey('From')) { $request['from'] = $From }
        if ($PSBoundParameters.ContainsKey('To') -or ($PSBoundParameters.ContainsKey('Para') -and $Para -gt 0 -and $To -gt 0)) { $request['to'] = $To }
        if ($Expect) { $request['expect'] = $Expect }
        if ($Style) { $request['style'] = $Style }
        if ($Match) { $request['match'] = $Match }
        Add-Format-Fields $PSBoundParameters
    }
    'style' {
        if (-not $Doc) { Write-Error-Json 'BAD_ARGS' 'style needs -Doc' }
        if (-not $Style) { Write-Error-Json 'BAD_ARGS' 'style needs -Style <style name as WPS shows it, e.g. Normal>' }
        $request['style'] = $Style
        Add-Format-Fields $PSBoundParameters
        if ($NoNumbering) { $request['numbering'] = $false }
    }
    'page' {
        if (-not $Doc) { Write-Error-Json 'BAD_ARGS' 'page needs -Doc' }
        foreach ($pair in @(@('Paper', 'paper'), @('Orientation', 'orientation'), @('Top', 'top'), @('Bottom', 'bottom'),
                            @('Left', 'left'), @('Right', 'right'), @('HeaderText', 'header'), @('FooterText', 'footer'),
                            @('PageNumber', 'pageNumber'))) {
            if ($PSBoundParameters.ContainsKey($pair[0])) { $request[$pair[1]] = $PSBoundParameters[$pair[0]] }
        }
        if ($PageNumberStart) { $request['pageNumberStart'] = $PageNumberStart }
        if ($All) { $request['all'] = $true }
    }
    'table' {
        if (-not $Doc) { Write-Error-Json 'BAD_ARGS' 'table needs -Doc' }
        if ($PSBoundParameters.ContainsKey('Para') -and $Para -gt 0) { $request['para'] = $Para }
        if ($Table -gt 0) { $request['table'] = $Table }
        if ($Action) { $request['action'] = $Action.ToLowerInvariant() }
        if ($Passes -gt 0) { $request['passes'] = $Passes }
        if ($Row -gt 0) { $request['row'] = $Row }
        if ($Force) { $request['force'] = $true }
        if ($Data) {
            if (-not (Test-Path -LiteralPath $Data)) { Write-Error-Json 'BAD_ARGS' ('Data file not found: ' + $Data) }
            $request['data'] = Read-Items-Json $Data -Keys @('data', 'rows') -NoArrayWrap
        }
        if ($Borders) { $request['borders'] = $Borders }
        if ($Widths) { $request['widths'] = $Widths }
        if ($HeaderRow) { $request['header'] = $true }
        if ($Caption) { $request['caption'] = $Caption }
        if ($Autofit) { $request['autofit'] = $Autofit }
        if ($RowHeight) { $request['rowHeight'] = $RowHeight }
        if ($KeepTogether) { $request['keepTogether'] = $true }
        # plan / normalize (v1.33)
        if ($PSBoundParameters.ContainsKey('Label') -and $Label) { $request['label'] = $Label }
        if ($PSBoundParameters.ContainsKey('DryRun')) { $request['dryRun'] = [bool]$DryRun }
        if ($Plan) { $request['plan'] = Resolve-FullPath $Plan }
        if ($NoCaptions) { $request['noCaptions'] = $true }
        if ($NoBorders) { $request['noBorders'] = $true }
        if ($NoHeaderRepeat) { $request['noHeaderRepeat'] = $true }
        if ($PSBoundParameters.ContainsKey('Gap')) { $request['gap'] = $Gap }
        if ($Fit) { $request['fit'] = $Fit }
        if ($Verify) { $request['verify'] = $true }
        if ($CutPaste) { $request['cutPaste'] = $true }
        if ($FailTable -gt 0) { $request['failTable'] = $FailTable }
        Add-Format-Fields $PSBoundParameters
    }
    'caption' {
        if (-not $Doc) { Write-Error-Json 'BAD_ARGS' 'caption needs -Doc' }
        if ($Para -lt 1) { Write-Error-Json 'BAD_ARGS' 'caption needs -Para <n>' }
        $request['para'] = $Para
        if ($Label) { $request['label'] = $Label }
        if ($Title) { $request['title'] = $Title }
        if ($Above) { $request['above'] = $true }
        if ($Align) { $request['align'] = $Align }
    }
    'picture' {
        if (-not $Doc) { Write-Error-Json 'BAD_ARGS' 'picture needs -Doc' }
        if ($Path) {
            $request['path'] = Resolve-FullPath $Path
            if (-not (Test-Path -LiteralPath $Path)) { Write-Error-Json 'BAD_ARGS' ('Picture not found: ' + $Path) }
            if ($Para -gt 0) { $request['para'] = $Para }
            if ($Width) { $request['width'] = $Width }
            if ($Caption) { $request['caption'] = $Caption }
            if ($Align) { $request['align'] = $Align }
            $request['action'] = 'insert'
        }
        else {
            if (-not $Image) { Write-Error-Json 'BAD_ARGS' 'picture needs -Path (insert) or -Image <n> (adjust)' }
            $request['image'] = [int]$Image
            $request['action'] = 'adjust'
            if ($Width) { $request['width'] = $Width }
            if ($Height) { $request['height'] = $Height }
            if ($CropLeft -ge 0) { $request['cropLeft'] = $CropLeft }
            if ($CropRight -ge 0) { $request['cropRight'] = $CropRight }
            if ($CropTop -ge 0) { $request['cropTop'] = $CropTop }
            if ($CropBottom -ge 0) { $request['cropBottom'] = $CropBottom }
        }
    }
    'toc' {
        if (-not $Doc) { Write-Error-Json 'BAD_ARGS' 'toc needs -Doc' }
        if ($Para -gt 0) { $request['para'] = $Para }
        if ($Title) { $request['title'] = $Title }
        $request['levels'] = $Levels
    }
    'reference' {
        if (-not $Doc) { Write-Error-Json 'BAD_ARGS' 'reference needs -Doc' }
        $request['action'] = if ($Action) { $Action.ToLowerInvariant() } elseif ($Items) { 'list' } else { 'cite' }
        if ($request['action'] -eq 'cite') {
            if ($Para -lt 1) { Write-Error-Json 'BAD_ARGS' 'reference -Action cite needs -Para <n>' }
            if (-not $Numbers) { Write-Error-Json 'BAD_ARGS' 'reference -Action cite needs -Numbers "1,2,5"' }
            $request['para'] = $Para
            $request['numbers'] = $Numbers
        }
        else {
            if (-not $Items) { Write-Error-Json 'BAD_ARGS' 'reference -Action list needs -Items <json file of strings>' }
            if (-not (Test-Path -LiteralPath $Items)) { Write-Error-Json 'BAD_ARGS' ('Items file not found: ' + $Items) }
            $request['items'] = Read-Items-Json $Items
            if ($AtPara -gt 0) { $request['para'] = $AtPara }
            if ($Heading) { $request['heading'] = $Heading }
        }
    }
    'view' {
        if (-not $Doc) { Write-Error-Json 'BAD_ARGS' 'view needs -Doc' }
        if ($Para -gt 0) { $request['para'] = $Para }
        elseif ($PSBoundParameters.ContainsKey('Page')) { $request['page'] = $Page }
        else { Write-Error-Json 'BAD_ARGS' 'view needs -Para <n> or -Page <n>' }
    }
    'fields' {
        if (-not $Doc) { Write-Error-Json 'BAD_ARGS' 'fields needs -Doc' }
    }
    'cell' {
        if (-not $Doc) { Write-Error-Json 'BAD_ARGS' 'cell needs -Doc' }
        if ($Table -gt 0) {
            $request['table'] = $Table
            if ($Row -gt 0) { $request['row'] = $Row }
            if ($Col -gt 0) { $request['col'] = $Col }
        }
        elseif ($Para -gt 0) { $request['para'] = $Para }
        else { Write-Error-Json 'BAD_ARGS' 'cell needs -Table -Row -Col, or -Para of a paragraph inside a cell' }
        if (-not $PSBoundParameters.ContainsKey('Text')) { Write-Error-Json 'BAD_ARGS' 'cell needs -Text (an empty string clears the cell)' }
        $request['text'] = $Text
        Add-Format-Fields $PSBoundParameters
    }
    'insert' {
        if (-not $Doc) { Write-Error-Json 'BAD_ARGS' 'insert needs -Doc' }
        if ($Items) {
            if (-not (Test-Path -LiteralPath $Items)) { Write-Error-Json 'BAD_ARGS' ('Items file not found: ' + $Items) }
            $request['items'] = Read-Items-Json $Items
        }
        elseif ($PSBoundParameters.ContainsKey('Text')) { $request['items'] = @(@{ text = $Text }) }
        else { Write-Error-Json 'BAD_ARGS' 'insert needs -Items <json file> or -Text' }
        if ($Para -gt 0) { $request['para'] = $Para }
        if ($Where) { $request['where'] = $Where }
        if ($Expect) { $request['expect'] = $Expect }
        if ($Track) { $request['track'] = $true }
    }
    'fillgaps' {
        if (-not $Doc) { Write-Error-Json 'BAD_ARGS' 'fillgaps needs -Doc' }
        $request['dryRun'] = [bool]$DryRun
        $request['maxMoves'] = $MaxMoves
        if ($MinGapPt -gt 0) { $request['minGapPt'] = $MinGapPt }
    }
    'revisions' {
        if (-not $Doc) { Write-Error-Json 'BAD_ARGS' 'revisions needs -Doc' }
        $act = if ($Action) { $Action.ToLowerInvariant() } else { 'list' }
        $request['action'] = $act
        if ($act -ne 'list') {
            if ($All) { $request['all'] = $true }
            elseif ($Index -gt 0) { $request['index'] = $Index }
            else { Write-Error-Json 'BAD_ARGS' 'revisions accept/reject needs -All or -Index <n>' }
        }
    }
    'snapshot' {
        if (-not $Doc) { Write-Error-Json 'BAD_ARGS' 'snapshot needs -Doc' }
        if ($Tag) { $request['tag'] = $Tag }
        if ($Out) { $request['out'] = Resolve-FullPath $Out }
        if ($Slow) { $request['slow'] = $true }
    }
    'diff' {
        if (-not $Doc) { Write-Error-Json 'BAD_ARGS' 'diff needs -Doc' }
        if ($Since) { $request['since'] = $Since }
        if ($Slow) { $request['slow'] = $true }
    }
    'textdiag' {
        if (-not $Doc) { Write-Error-Json 'BAD_ARGS' 'textdiag needs -Doc' }
        if ($Slow) { $request['slow'] = $true }
    }
    'handoff' {
        if (-not $Doc) { Write-Error-Json 'BAD_ARGS' 'handoff needs -Doc' }
        if ($Note) { $request['note'] = $Note }
        if ($As) { $request['out'] = Resolve-FullPath $As }
    }
    'audit' {
        if (-not $Doc) { Write-Error-Json 'BAD_ARGS' 'audit needs -Doc' }
        if ($Max -gt 0) { $request['max'] = $Max }
        if ($Action) { $request['rules'] = $Action }
        if ($Slow) { $request['slow'] = $true }
    }
    'unit' {
        if (-not $Doc) { Write-Error-Json 'BAD_ARGS' 'unit needs -Doc' }
        if ($Apply) { $request['apply'] = $true }
        if ($Normalize) { $request['normalize'] = $true }
        if ($Units) { $request['units'] = $Units }
        if ($Track) { $request['track'] = $true }
    }
    'hint' {
        # Shared body (woc-app-common.ps1); -Tray tells the daemon which wrapper to re-run.
        Add-WocHintFields -Request $request -Doc $Doc -HintTag $HintTag -Bar $Bar -Text $Text -Wrapper $PSCommandPath -Clear:$Clear -Toast:$Toast -Sound:$Sound -Quiet:$Quiet -Tray:$Tray
    }
    'finish' {
        if ($Note) { $request['note'] = $Note }
        if ($Toast) { $request['toast'] = $true }
        if ($Sound) { $request['sound'] = $true }
        if ($Quiet) { $request['quiet'] = $true }
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
# Shared with the other two wrappers: any bound parameter this command did not copy into the
# request warns on stderr instead of vanishing (see scripts\param-check.ps1).
$paramAliases = @{
    TextFile = 'text'; HintTag = 'tag'; Case = 'caseSensitive'
    NoNumbering = 'numbering'; HeaderRow = 'header'; AtPara = 'para'
    HeaderText = 'header'; FooterText = 'footer'; PageNumber = 'pageNumber'
    SplitWords = 'noSplitWords'
}

Invoke-WocParamCheck -ScriptRoot $PSScriptRoot -Command $Command -Bound $PSBoundParameters -Request $request `
    -Aliases $paramAliases -FormatFields $formatFields -DefaultValues @{ From = 1; Count = 0; Index = 0; Max = 200; Para = 0; To = 0; Row = 0; Col = 0; Table = 0; AtPara = 0; Passes = 0; MaxMoves = 8; MinGapPt = 0; Gap = -1; FailTable = 0 }
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
if ($null -eq $responseText) { $responseText = Invoke-WocOneShot -App 'word' -Command $Command -ExePath $exePath -Json $json -TimeoutSec $TimeoutSec }
exit 0
