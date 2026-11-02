# wps-office-cooperate multi-task helpers for Excel (v1.22).
#
# Pure ASCII source: every Chinese string travels inside UTF-8 JSON files, never in this script.
#
#   tasks-excel.ps1 split  -Doc <xlsx> -Ranges ranges.json -Out <dir>
#       ranges.json = [{"id":"w1","sheet":"Sheet1","range":"A1:F200"}, ...]
#       -> <dir>/slice-<id>.json  (excel.ps1 read output for that range; hand it to a planner)
#          <dir>/index.json
#
#   tasks-excel.ps1 merge  -TaskFiles <dir or a.json,b.json> -Out <dir> [-Ranges ranges.json]
#       task file = {"task":"E1","priority":10,
#                    "writes":[{"sheet":"S","cell":"B2","formula":"=A2*2","expect":"0"}],
#                    "formats":[{"sheet":"S","range":"A1:F1","bold":true}],
#                    "comments":[{"sheet":"S","cell":"A1","text":"..."}],
#                    "structures":[...], "charts":[...]}
#       -> <dir>/merged.json (deduped, conflicts listed, warnings for missing expect)
#
#   tasks-excel.ps1 plan   -Merged <dir>/merged.json -Out <dir> [-Doc <absolute xlsx>] [-BatchOps]
#       -> writes.json / formats.json / comments.json / structure-ops.json / apply-order.json
#          and, with -BatchOps, batch-ops.json ready for `excel.ps1 batch -Ops`
#
#   tasks-excel.ps1 selftest
#
# Why this shape: Excel has NO revision layer - a bad write is only undoable through a snapshot.
# So the merge layer is deliberately conservative: it refuses to plan two different values for the
# same cell, flags writes without an `expect` anchor, and keeps structural edits (rows/cols/sheets)
# in their own file because they shift every address after them.

param(
    [Parameter(Mandatory = $true, Position = 0)]
    [ValidateSet('split', 'merge', 'plan', 'selftest', 'help')]
    [string]$Command,

    [string]$Doc,
    [string]$Ranges,
    [string]$TaskFiles,
    [string]$Merged,
    [string]$Out,
    [switch]$BatchOps
)

$ErrorActionPreference = 'Stop'

function Read-Json([string]$path) {
    if (-not (Test-Path -LiteralPath $path)) { throw "file not found: $path" }
    $text = [System.IO.File]::ReadAllText($path, [System.Text.Encoding]::UTF8)
    if (-not $text.Trim()) { throw "empty JSON file: $path" }
    return ($text | ConvertFrom-Json)
}
function Write-Json([string]$path, $obj) {
    $dir = Split-Path -Parent $path
    if ($dir -and -not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Force -Path $dir | Out-Null }
    $text = (ConvertTo-Json -InputObject $obj -Depth 20 -Compress)
    [System.IO.File]::WriteAllText($path, $text, (New-Object System.Text.UTF8Encoding($false)))
}
function FullPath2([string]$p) { return [System.IO.Path]::GetFullPath($p) }

# ---------------------------------------------------------------- range math

# "A1" -> @{row=1; col=1}; "AB12" -> @{row=12; col=28}; returns $null when unparsable.
function Parse-Cell([string]$ref) {
    if (-not $ref) { return $null }
    $m = [regex]::Match(('' + $ref).Trim().ToUpperInvariant(), '^([A-Z]+)([0-9]+)$')
    if (-not $m.Success) { return $null }
    $col = 0
    foreach ($ch in $m.Groups[1].Value.ToCharArray()) { $col = $col * 26 + ([int][char]$ch - 64) }
    return @{ row = [int]$m.Groups[2].Value; col = $col }
}

# "A1:B2" -> rect; a single cell -> 1x1 rect. Rows/cols from a values block override the ref.
function Parse-Rect([string]$ref, [int]$blockRows, [int]$blockCols) {
    if (-not $ref) { return $null }
    $parts = ('' + $ref).Trim().Split(':')
    $a = Parse-Cell $parts[0]
    if ($a -eq $null) { return $null }
    $r1 = $a.row; $c1 = $a.col
    $r2 = $r1; $c2 = $c1
    if ($parts.Count -gt 1) {
        $b = Parse-Cell $parts[1]
        if ($b -eq $null) { return $null }
        $r2 = $b.row; $c2 = $b.col
    }
    if ($blockRows -gt 0 -and $blockCols -gt 0) { $r2 = $r1 + $blockRows - 1; $c2 = $c1 + $blockCols - 1 }
    if ($r2 -lt $r1) { $t = $r1; $r1 = $r2; $r2 = $t }
    if ($c2 -lt $c1) { $t = $c1; $c1 = $c2; $c2 = $t }
    return @{ r1 = $r1; c1 = $c1; r2 = $r2; c2 = $c2 }
}
function Rects-Overlap($a, $b) {
    return (-not ($a.r2 -lt $b.r1 -or $b.r2 -lt $a.r1 -or $a.c2 -lt $b.c1 -or $b.c2 -lt $a.c1))
}
function Rect-Key($a) { return ('' + $a.r1 + ',' + $a.c1 + '-' + $a.r2 + ',' + $a.c2) }

# ---------------------------------------------------------------- split

function Invoke-Split {
    if (-not $Doc) { throw 'split needs -Doc <xlsx>' }
    if (-not $Ranges) { throw 'split needs -Ranges <ranges.json>' }
    if (-not $Out) { throw 'split needs -Out <dir>' }
    $d = FullPath2 $Doc
    if (-not (Test-Path -LiteralPath $d)) { throw ('workbook not found: ' + $d) }
    $rangeList = @(Read-Json $Ranges)
    if ($rangeList.Count -eq 0) { throw 'ranges.json is empty' }
    $excel = Join-Path $PSScriptRoot 'excel.ps1'
    $work = Join-Path $Out '_read'
    New-Item -ItemType Directory -Force -Path $Out, $work | Out-Null
    $summary = New-Object System.Collections.ArrayList
    foreach ($r in $rangeList) {
        $sheet = [string]$r.sheet
        $range = [string]$r.range
        if (-not $sheet -or -not $range) { throw ('range ' + $r.id + ': needs sheet and range') }
        $readOut = Join-Path $work ([string]$r.id + '.json')
        $pa = Start-Process -FilePath 'powershell' -ArgumentList @('-Mta', '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $excel, 'read', '-Doc', $d, '-Range', ($sheet + '!' + $range)) -PassThru -NoNewWindow -RedirectStandardOutput $readOut
        $pa.WaitForExit()
        $res = Read-Json $readOut
        if (-not $res.ok) { throw ('read failed for ' + $r.id + ': ' + $res.error.code + ' ' + $res.error.message) }
        $slice = [pscustomobject]@{
            id      = [string]$r.id
            sheet   = $sheet
            range   = $range
            rows    = $res.result.rows
            cols    = $res.result.cols
            values  = $res.result.values
        }
        $path = Join-Path $Out ('slice-' + $r.id + '.json')
        Write-Json $path $slice
        [void]$summary.Add([pscustomobject]@{ id = [string]$r.id; sheet = $sheet; range = $range; rows = $res.result.rows; cols = $res.result.cols; file = $path })
        Write-Output ('slice ' + $r.id + ': ' + $sheet + '!' + $range + ' (' + $res.result.rows + 'x' + $res.result.cols + ') -> ' + $path)
    }
    Write-Json (Join-Path $Out 'index.json') ([pscustomobject]@{ doc = $d; slices = @($summary) })
    Write-Output ('split done: ' + $summary.Count + ' slice(s)')
}

# ---------------------------------------------------------------- merge

function Invoke-Merge {
    if (-not $TaskFiles) { throw 'merge needs -TaskFiles <dir or a.json,b.json>' }
    if (-not $Out) { throw 'merge needs -Out <dir>' }
    $files = @()
    if (Test-Path -LiteralPath $TaskFiles -PathType Container) {
        $files = @(Get-ChildItem -LiteralPath $TaskFiles -Filter 'task-*.json' | Sort-Object Name | ForEach-Object { $_.FullName })
    }
    else {
        $files = @($TaskFiles.Split(',') | ForEach-Object { $_.Trim() } | Where-Object { $_ })
    }
    if ($files.Count -eq 0) { throw ('no task-*.json found in ' + $TaskFiles) }

    $conflicts = New-Object System.Collections.ArrayList
    $warnings = New-Object System.Collections.ArrayList
    $writes = New-Object System.Collections.ArrayList
    $formats = New-Object System.Collections.ArrayList
    $comments = New-Object System.Collections.ArrayList
    $structures = New-Object System.Collections.ArrayList
    $charts = New-Object System.Collections.ArrayList
    $order = 0

    foreach ($f in $files) {
        $t = Read-Json $f
        $taskId = [string]$t.task
        if (-not $taskId) { $taskId = [System.IO.Path]::GetFileNameWithoutExtension($f) }
        $prio = 50
        try { if ($t.priority) { $prio = [int]$t.priority } } catch { }
        foreach ($w in @($t.writes)) {
            $order++
            if ($null -eq $w) { continue }
            $sheet = [string]$w.sheet
            $cell = [string]$w.cell
            if (-not $sheet -or -not $cell) { [void]$conflicts.Add($taskId + ': write without sheet/cell'); continue }
            $hasValue = $w.PSObject.Properties.Name -contains 'value'
            $hasFormula = $w.PSObject.Properties.Name -contains 'formula'
            $hasValues = $w.PSObject.Properties.Name -contains 'values'
            $n = @($hasValue, $hasFormula, $hasValues) | Where-Object { $_ } | Measure-Object | Select-Object -ExpandProperty Count
            if ($n -ne 1) { [void]$conflicts.Add($taskId + ': write ' + $sheet + '!' + $cell + ' needs exactly one of value/formula/values'); continue }
            $rows = 0; $cols = 0
            if ($hasValues) {
                $block = @($w.values)
                $rows = $block.Count
                if ($rows -gt 0) { $cols = @($block[0]).Count }
                if ($rows -le 0 -or $cols -le 0) { [void]$conflicts.Add($taskId + ': write ' + $sheet + '!' + $cell + ' has an empty values block'); continue }
            }
            $rect = Parse-Rect $cell $rows $cols
            if ($null -eq $rect) { [void]$conflicts.Add($taskId + ': unparsable cell/range ' + $cell); continue }
            $expect = ''
            if ($w.PSObject.Properties.Name -contains 'expect') { $expect = [string]$w.expect }
            if (-not $expect) { [void]$warnings.Add($taskId + ': write ' + $sheet + '!' + $cell + ' has no expect anchor (blind write)') }
            [void]$writes.Add([pscustomobject]@{
                    task = $taskId; priority = $prio; order = $order; sheet = $sheet; cell = $cell; rect = $rect
                    kind = $(if ($hasFormula) { 'formula' } elseif ($hasValues) { 'values' } else { 'value' })
                    value = $w.value; formula = $w.formula; values = $w.values; expect = $expect
                })
        }
        foreach ($fm in @($t.formats)) {
            if ($null -eq $fm) { continue }
            $order++
            $sheet = [string]$fm.sheet
            $range = [string]$fm.range
            if (-not $sheet -or -not $range) { [void]$conflicts.Add($taskId + ': format without sheet/range'); continue }
            [void]$formats.Add([pscustomobject]@{ task = $taskId; priority = $prio; order = $order; sheet = $sheet; range = $range; opts = $fm })
        }
        foreach ($c in @($t.comments)) {
            if ($null -eq $c) { continue }
            $order++
            $sheet = [string]$c.sheet
            $cell = [string]$c.cell
            $text = [string]$c.text
            if (-not $sheet -or -not $cell -or -not $text) { [void]$conflicts.Add($taskId + ': comment without sheet/cell/text'); continue }
            [void]$comments.Add([pscustomobject]@{ task = $taskId; priority = $prio; order = $order; sheet = $sheet; cell = $cell; text = $text })
        }
        foreach ($s in @($t.structures)) { if ($s) { [void]$structures.Add([pscustomobject]@{ task = $taskId; priority = $prio; op = $s }) } }
        foreach ($ch in @($t.charts)) { if ($ch) { [void]$charts.Add([pscustomobject]@{ task = $taskId; priority = $prio; op = $ch }) } }
    }

    # ------- conflict detection: overlapping writes (same sheet + intersecting rectangles)
    $kept = New-Object System.Collections.ArrayList
    $bySheet = @{}
    foreach ($w in ($writes | Sort-Object -Property priority, order)) {
        if (-not $bySheet.ContainsKey($w.sheet)) { $bySheet[$w.sheet] = New-Object System.Collections.ArrayList }
        $hit = $null
        foreach ($prev in $bySheet[$w.sheet]) { if (Rects-Overlap $prev.rect $w.rect) { $hit = $prev; break } }
        if ($hit -ne $null) {
            $same = $false
            if ($hit.kind -eq $w.kind) {
                if ($w.kind -eq 'value' -and ([string]$hit.value -eq [string]$w.value)) { $same = $true }
                if ($w.kind -eq 'formula' -and ([string]$hit.formula -eq [string]$w.formula)) { $same = $true }
            }
            if ($same) { $warnings.Add('duplicate identical write ' + $w.sheet + '!' + $w.cell + ' from ' + $w.task + ' (kept)') }
            else {
                [void]$conflicts.Add('write overlap on ' + $w.sheet + '!' + $w.cell + ': ' + $hit.task + ' vs ' + $w.task + ' (kept ' + $hit.task + ')')
                continue
            }
        }
        [void]$kept.Add($w)
        [void]$bySheet[$w.sheet].Add($w)
    }

    # ------- comments: one comment per cell wins
    $commentKept = New-Object System.Collections.ArrayList
    $seenComment = @{}
    foreach ($c in ($comments | Sort-Object -Property priority, order)) {
        $k = ($c.sheet + '!' + $c.cell).ToLowerInvariant()
        if ($seenComment.ContainsKey($k)) {
            if ($seenComment[$k] -ne $c.text) { [void]$conflicts.Add('two different comments on ' + $c.sheet + '!' + $c.cell + ' (' + $c.task + ' dropped)') }
            continue
        }
        $seenComment[$k] = $c.text
        [void]$commentKept.Add($c)
    }

    if ($structures.Count -gt 0) {
        [void]$warnings.Add('structure ops present: they shift every address after them - apply them first, then re-read before the value/formula batch (this plan cannot know the post-structure addresses)')
    }

    $stats = [pscustomobject]@{
        writes = $kept.Count; formats = $formats.Count; comments = $commentKept.Count
        structures = $structures.Count; charts = $charts.Count
    }
    $merged = [pscustomobject]@{
        doc       = $(if ($Doc) { FullPath2 $Doc } else { $null })
        stats     = $stats
        writes    = @($kept)
        formats   = @($formats | Sort-Object -Property priority, order)
        comments  = @($commentKept)
        structures = @($structures)
        charts    = @($charts)
        conflicts = @($conflicts)
        warnings  = @($warnings)
    }
    Write-Json (Join-Path $Out 'merged.json') $merged
    Write-Output ('merge done: ' + $kept.Count + ' write(s), ' + $formats.Count + ' format(s), ' + $commentKept.Count + ' comment(s), ' + $conflicts.Count + ' conflict(s), ' + $warnings.Count + ' warning(s)')
    foreach ($c in $conflicts) { Write-Output ('  conflict: ' + $c) }
}

# ---------------------------------------------------------------- plan

function New-BatchEntry([string]$cmd, [string]$docPath, $items) {
    $h = [ordered]@{ cmd = $cmd }
    if ($docPath) { $h['doc'] = $docPath }
    $h['items'] = @($items)
    return [pscustomobject]$h
}

function Invoke-Plan {
    if (-not $Merged) { throw 'plan needs -Merged <merged.json>' }
    if (-not $Out) { throw 'plan needs -Out <dir>' }
    $merged = Read-Json $Merged
    New-Item -ItemType Directory -Force -Path $Out | Out-Null
    $docPath = $null
    if ($Doc) { $docPath = FullPath2 $Doc }
    elseif ($merged.doc) { $docPath = [string]$merged.doc }

    $writeItems = New-Object System.Collections.ArrayList
    foreach ($w in @($merged.writes)) {
        $item = [ordered]@{ sheet = [string]$w.sheet; cell = [string]$w.cell }
        if ($w.kind -eq 'formula') { $item['formula'] = [string]$w.formula }
        elseif ($w.kind -eq 'values') { $item['values'] = $w.values }
        else { $item['value'] = $w.value }
        if ([string]$w.expect) { $item['expect'] = [string]$w.expect }
        [void]$writeItems.Add([pscustomobject]$item)
    }
    $formatItems = New-Object System.Collections.ArrayList
    foreach ($fm in @($merged.formats)) {
        $o = $fm.opts
        $item = [ordered]@{ sheet = [string]$fm.sheet; range = [string]$fm.range }
        foreach ($p in $o.PSObject.Properties) {
            if ($p.Name -in @('sheet', 'range', 'task')) { continue }
            $item[$p.Name] = $p.Value
        }
        [void]$formatItems.Add([pscustomobject]$item)
    }
    $commentItems = @($merged.comments | ForEach-Object { [pscustomobject]@{ sheet = $_.sheet; cell = $_.cell; text = $_.text } })

    Write-Json (Join-Path $Out 'writes.json') ([pscustomobject]@{ items = @($writeItems) })
    Write-Json (Join-Path $Out 'formats.json') ([pscustomobject]@{ items = @($formatItems) })
    Write-Json (Join-Path $Out 'comments.json') ([pscustomobject]@{ items = @($commentItems) })
    if (@($merged.structures).Count -gt 0) { Write-Json (Join-Path $Out 'structure-ops.json') ([pscustomobject]@{ items = @($merged.structures) }) }
    if (@($merged.charts).Count -gt 0) { Write-Json (Join-Path $Out 'chart-ops.json') ([pscustomobject]@{ items = @($merged.charts) }) }

    $phases = @()
    if (@($merged.structures).Count -gt 0) { $phases += 'structures (apply first, then re-read the workbook)' }
    if ($writeItems.Count -gt 0) { $phases += 'write (values/formulas, one atomic batch)' }
    if ($formatItems.Count -gt 0) { $phases += 'format' }
    if ($commentItems.Count -gt 0) { $phases += 'comments' }
    $phases += 'calc + diff (verify)'
    $order = [pscustomobject]@{
        doc      = $docPath
        phases   = @($phases)
        counts   = [pscustomobject]@{ writes = $writeItems.Count; formats = $formatItems.Count; comments = $commentItems.Count }
        files    = [pscustomobject]@{
            writes   = (Join-Path $Out 'writes.json'); formats = (Join-Path $Out 'formats.json')
            comments = (Join-Path $Out 'comments.json')
        }
    }
    Write-Json (Join-Path $Out 'apply-order.json') $order

    if ($BatchOps) {
        $batch = New-Object System.Collections.ArrayList
        if ($writeItems.Count -gt 0) { [void]$batch.Add((New-BatchEntry 'write' $docPath @($writeItems))) }
        if ($formatItems.Count -gt 0) { [void]$batch.Add((New-BatchEntry 'format' $docPath @($formatItems))) }
        foreach ($c in $commentItems) {
            $h = [ordered]@{ cmd = 'comment' }
            if ($docPath) { $h['doc'] = $docPath }
            $h['sheet'] = [string]$c.sheet
            $h['cell'] = [string]$c.cell
            $h['text'] = [string]$c.text
            [void]$batch.Add([pscustomobject]$h)
        }
        Write-Json (Join-Path $Out 'batch-ops.json') @($batch)
        Write-Output ('batch ops: ' + $batch.Count + ' op(s) -> ' + (Join-Path $Out 'batch-ops.json'))
    }
    Write-Output ('plan done: writes=' + $writeItems.Count + ' formats=' + $formatItems.Count + ' comments=' + $commentItems.Count)
    foreach ($p in $phases) { Write-Output ('  phase: ' + $p) }
}

# ---------------------------------------------------------------- selftest

function Invoke-Selftest {
    $fails = New-Object System.Collections.ArrayList
    $r = Parse-Rect 'B2:D5' 0 0
    if ($r -eq $null -or $r.r1 -ne 2 -or $r.c1 -ne 2 -or $r.r2 -ne 5 -or $r.c2 -ne 4) { [void]$fails.Add('parse rect B2:D5') }
    $r2 = Parse-Rect 'AA10' 0 0
    if ($r2 -eq $null -or $r2.c1 -ne 27 -or $r2.r1 -ne 10) { [void]$fails.Add('parse cell AA10') }
    $r3 = Parse-Rect 'A1' 3 2
    if ($r3 -eq $null -or $r3.r2 -ne 3 -or $r3.c2 -ne 2) { [void]$fails.Add('block dims A1 3x2') }
    if (-not (Rects-Overlap (Parse-Rect 'A1:B2' 0 0) (Parse-Rect 'B2:C3' 0 0))) { [void]$fails.Add('overlap should be true') }
    if ((Rects-Overlap (Parse-Rect 'A1:B2' 0 0) (Parse-Rect 'C1:D2' 0 0))) { [void]$fails.Add('overlap should be false') }
    if ($fails.Count -gt 0) { Write-Output ('SELFTEST FAIL: ' + ($fails -join '; ')); exit 1 }
    Write-Output 'SELFTEST PASS'
}

switch ($Command) {
    'split' { Invoke-Split }
    'merge' { Invoke-Merge }
    'plan' { Invoke-Plan }
    'selftest' { Invoke-Selftest }
    default {
        Write-Output 'usage: tasks-excel.ps1 split|merge|plan|selftest [options]'
        Write-Output '  split  -Doc <xlsx> -Ranges ranges.json -Out dir'
        Write-Output '  merge  -TaskFiles dir|files -Out dir'
        Write-Output '  plan   -Merged merged.json -Out dir [-Doc abs-xlsx] [-BatchOps]'
    }
}
