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
                 'save', 'close', 'quit', 'finish')]
    [string]$Command,

    [string]$Doc,
    [int]$From = 1,
    [int]$Count = 0,
    [switch]$Tables,
    [string]$Items,
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
    [int]$Row = 0,
    [int]$Col = 0,
    [string]$Where,
    [int]$Index = 0,
    [string]$Tag,
    [string]$Since,
    [string]$Note,
    [switch]$Toast,
    [switch]$Sound,
    [switch]$Quiet,
    [switch]$Delete,
    [int]$Max = 200,
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
$sourceDir = Join-Path $scriptRoot 'helper\word'
if (-not (Test-Path -LiteralPath $sourceDir)) { Write-Error-Json 'NO_HELPER' ("Helper folder not found: " + $sourceDir) }
$sources = @(Get-ChildItem -LiteralPath $sourceDir -Filter *.cs | Sort-Object Name)
if ($sources.Count -eq 0) { Write-Error-Json 'NO_HELPER' ('No .cs in ' + $sourceDir) }

# The cache key covers every source file, so editing any of them rebuilds.
$hashInput = ($sources | ForEach-Object { $_.Name + ':' + (Get-FileHash -LiteralPath $_.FullName -Algorithm SHA256).Hash }) -join ';'
$shaGen = [System.Security.Cryptography.SHA256]::Create()
$sha = ([System.BitConverter]::ToString($shaGen.ComputeHash([System.Text.Encoding]::UTF8.GetBytes($hashInput))) -replace '-', '').ToLowerInvariant().Substring(0, 12)
$cacheDir = Join-Path $env:LOCALAPPDATA ('wps-office-cooperate\word-helper-' + $sha)
$exePath = Join-Path $cacheDir 'WpsHelper.exe'

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
}

switch ($Command) {
    'open' {
        if (-not $Doc) { Write-Error-Json 'BAD_ARGS' 'open needs -Doc' }
        $request['show'] = [bool]$Show
        $request['readOnly'] = [bool]$ReadOnly
    }
    'read' {
        if (-not $Doc) { Write-Error-Json 'BAD_ARGS' 'read needs -Doc' }
        $request['from'] = $From
        $request['count'] = $Count   # 0 = every paragraph from -From on
        if ($SkipTableCells) { $request['skipTableCells'] = $true }
        if ($Only) { $request['only'] = $Only }
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
        $raw = [System.IO.File]::ReadAllText((Resolve-FullPath $Items), [System.Text.Encoding]::UTF8)
        $parsed = $raw | ConvertFrom-Json
        if ($parsed.PSObject.Properties.Name -contains 'items') { $request['items'] = @($parsed.items) }
        else { $request['items'] = @($parsed) }
    }
    'comment' {
        if (-not $Doc) { Write-Error-Json 'BAD_ARGS' 'comment needs -Doc' }
        if ($Action -and $Action.ToLowerInvariant() -eq 'list') { $request['action'] = 'list' }
        elseif ($Items) {
            if (-not (Test-Path -LiteralPath $Items)) { Write-Error-Json 'BAD_ARGS' ('Items file not found: ' + $Items) }
            $raw = [System.IO.File]::ReadAllText((Resolve-FullPath $Items), [System.Text.Encoding]::UTF8)
            $parsed = $raw | ConvertFrom-Json
            if ($parsed.PSObject.Properties.Name -contains 'items') { $request['items'] = @($parsed.items) } else { $request['items'] = @($parsed) }
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
        if ($To -gt 0) { $request['to'] = $To }
        $request['out'] = Resolve-FullPath $Out
        $request['width'] = if ($Width) { [int]$Width } else { 1100 }
    }
    'font' {
        if (-not $Doc) { Write-Error-Json 'BAD_ARGS' 'font needs -Doc' }
        if (-not $Name) { Write-Error-Json 'BAD_ARGS' 'font needs -Name <font name, e.g. 仿宋 or FangSong>' }
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
        if (-not $Style) { Write-Error-Json 'BAD_ARGS' 'style needs -Style <style name, e.g. 正文>' }
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
            $raw = [System.IO.File]::ReadAllText((Resolve-FullPath $Data), [System.Text.Encoding]::UTF8)
            $parsed = $raw | ConvertFrom-Json
            if ($parsed.PSObject.Properties.Name -contains 'data') { $request['data'] = $parsed.data }
            elseif ($parsed.PSObject.Properties.Name -contains 'rows') { $request['data'] = $parsed.rows }
            else { $request['data'] = $parsed }
        }
        if ($Borders) { $request['borders'] = $Borders }
        if ($Widths) { $request['widths'] = $Widths }
        if ($HeaderRow) { $request['header'] = $true }
        if ($Caption) { $request['caption'] = $Caption }
        if ($Autofit) { $request['autofit'] = $Autofit }
        if ($RowHeight) { $request['rowHeight'] = $RowHeight }
        if ($KeepTogether) { $request['keepTogether'] = $true }
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
            $raw = [System.IO.File]::ReadAllText((Resolve-FullPath $Items), [System.Text.Encoding]::UTF8)
            $parsed = $raw | ConvertFrom-Json
            if ($parsed.PSObject.Properties.Name -contains 'items') { $request['items'] = @($parsed.items) } else { $request['items'] = @($parsed) }
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
            $raw = [System.IO.File]::ReadAllText((Resolve-FullPath $Items), [System.Text.Encoding]::UTF8)
            $parsed = $raw | ConvertFrom-Json
            if ($parsed.PSObject.Properties.Name -contains 'items') { $request['items'] = @($parsed.items) } else { $request['items'] = @($parsed) }
        }
        elseif ($PSBoundParameters.ContainsKey('Text')) { $request['items'] = @(@{ text = $Text }) }
        else { Write-Error-Json 'BAD_ARGS' 'insert needs -Items <json file> or -Text' }
        if ($Para -gt 0) { $request['para'] = $Para }
        if ($Where) { $request['where'] = $Where }
        if ($Expect) { $request['expect'] = $Expect }
        if ($Track) { $request['track'] = $true }
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
    }
    'diff' {
        if (-not $Doc) { Write-Error-Json 'BAD_ARGS' 'diff needs -Doc' }
        if ($Since) { $request['since'] = $Since }
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
    }
    'unit' {
        if (-not $Doc) { Write-Error-Json 'BAD_ARGS' 'unit needs -Doc' }
        if ($Apply) { $request['apply'] = $true }
        if ($Normalize) { $request['normalize'] = $true }
        if ($Units) { $request['units'] = $Units }
        if ($Track) { $request['track'] = $true }
    }
    'units' {
        if (-not $Doc) { Write-Error-Json 'BAD_ARGS' 'units needs -Doc' }
        if ($Apply) { $request['apply'] = $true }
        if ($Normalize) { $request['normalize'] = $true }
        if ($Units) { $request['units'] = $Units }
        if ($Track) { $request['track'] = $true }
    }
    'finish' {
        if ($Note) { $request['note'] = $Note }
        if ($Toast) { $request['toast'] = $true }
        if ($Sound) { $request['sound'] = $true }
        if ($Quiet) { $request['quiet'] = $true }
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
$work = Join-Path $env:TEMP ('wps-office-cooperate-word-' + [guid]::NewGuid().ToString('N'))
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
