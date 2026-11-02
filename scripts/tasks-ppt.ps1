# wps-office-cooperate multi-task helpers for PowerPoint (v1.23).
#
# Pure ASCII source: Chinese travels inside UTF-8 JSON files only.
#
#   tasks-ppt.ps1 split  -Doc <pptx> -Ranges ranges.json -Out <dir>
#       ranges.json = [{"id":"s1","from":1,"to":10}, ...]
#       -> <dir>/slice-<id>.json (ppt.ps1 read output for those slides)
#
#   tasks-ppt.ps1 merge  -TaskFiles <dir or a.json,b.json> -Out <dir>
#       task file = {"task":"P1","priority":10,
#                    "texts":[{"slide":3,"shape":"title","expect":"old","text":"new"}],
#                    "replaces":[{"find":"X","replace":"Y","scope":"all","slides":"1-5"}],
#                    "formats":[{"slide":3,"shape":"title","font":"...","size":28}],
#                    "comments":[{"slide":3,"shape":"title","text":"..."}],
#                    "globals":[{"kind":"font","font":"..."}],
#                    "structures":[...], "pictures":[...], "charts":[...]}
#       -> <dir>/merged.json (per-slide/shape arbitration, conflicts listed)
#
#   tasks-ppt.ps1 plan   -Merged <dir>/merged.json -Out <dir> [-Doc <absolute pptx>] [-BatchOps]
#       -> text-items.json / replace-items.json / comment-items.json / global-ops.json
#          structure-ops.json / apply-order.json (+ batch-ops.json)
#
#   tasks-ppt.ps1 selftest
#
# Ordering is the point: structure (add/delete/move slides) shifts slide numbers, so it must run
# first and the rest must be re-planned; globals (theme/master/font) repaint every slide, so they
# run before per-slide text/format work. Deck has no revision layer: snapshot before a batch.

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
    [System.IO.File]::WriteAllText($path, (ConvertTo-Json -InputObject $obj -Depth 20 -Compress), (New-Object System.Text.UTF8Encoding($false)))
}
function FullPath2([string]$p) { return [System.IO.Path]::GetFullPath($p) }

# "1-5,8,10-12" -> sorted int array; "*" / "" -> $null (means "all")
function Parse-SlideSet([string]$spec) {
    if (-not $spec -or $spec.Trim() -eq '*' -or $spec.Trim().ToLowerInvariant() -eq 'all') { return $null }
    $set = New-Object System.Collections.Generic.SortedSet[int]
    foreach ($part in $spec.Split(',')) {
        $s = $part.Trim()
        if (-not $s) { continue }
        $m = [regex]::Match($s, '^([0-9]+)\s*-\s*([0-9]+)$')
        if ($m.Success) {
            $a = [int]$m.Groups[1].Value; $b = [int]$m.Groups[2].Value
            for ($k = [Math]::Min($a, $b); $k -le [Math]::Max($a, $b); $k++) { [void]$set.Add($k) }
        }
        else {
            $n = 0
            if ([int]::TryParse($s, [ref]$n)) { [void]$set.Add($n) }
        }
    }
    return @($set)
}

function Slides-Overlap([string]$a, [string]$b) {
    $sa = Parse-SlideSet $a; $sb = Parse-SlideSet $b
    if ($null -eq $sa -or $null -eq $sb) { return $true }   # one side is "all"
    foreach ($x in $sa) { if ($sb -contains $x) { return $true } }
    return $false
}

# ---------------------------------------------------------------- split

function Invoke-Split {
    if (-not $Doc) { throw 'split needs -Doc <pptx>' }
    if (-not $Ranges) { throw 'split needs -Ranges <ranges.json>' }
    if (-not $Out) { throw 'split needs -Out <dir>' }
    $d = FullPath2 $Doc
    if (-not (Test-Path -LiteralPath $d)) { throw ('deck not found: ' + $d) }
    $rangeList = @(Read-Json $Ranges)
    if ($rangeList.Count -eq 0) { throw 'ranges.json is empty' }
    $ppt = Join-Path $PSScriptRoot 'ppt.ps1'
    $work = Join-Path $Out '_read'
    New-Item -ItemType Directory -Force -Path $Out, $work | Out-Null
    $summary = New-Object System.Collections.ArrayList
    foreach ($r in $rangeList) {
        $from = [int]$r.from; $to = [int]$r.to
        if ($from -lt 1 -or $to -lt $from) { throw ('range ' + $r.id + ': bad from/to') }
        $count = $to - $from + 1
        $readOut = Join-Path $work ([string]$r.id + '.json')
        $pa = Start-Process -FilePath 'powershell' -ArgumentList @('-Mta', '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $ppt, 'read', '-Doc', $d, '-From', $from, '-Count', $count) -PassThru -NoNewWindow -RedirectStandardOutput $readOut
        $pa.WaitForExit()
        $res = Read-Json $readOut
        if (-not $res.ok) { throw ('read failed for ' + $r.id + ': ' + $res.error.code + ' ' + $res.error.message) }
        $slice = [pscustomobject]@{ id = [string]$r.id; from = $from; to = $to; slides = @($res.result.items) }
        $path = Join-Path $Out ('slice-' + $r.id + '.json')
        Write-Json $path $slice
        [void]$summary.Add([pscustomobject]@{ id = [string]$r.id; from = $from; to = $to; slides = @($res.result.items).Count; file = $path })
        Write-Output ('slice ' + $r.id + ': slides ' + $from + '..' + $to + ' -> ' + $path)
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
    $texts = New-Object System.Collections.ArrayList
    $replaces = New-Object System.Collections.ArrayList
    $formats = New-Object System.Collections.ArrayList
    $comments = New-Object System.Collections.ArrayList
    $globals = New-Object System.Collections.ArrayList
    $structures = New-Object System.Collections.ArrayList
    $pictures = New-Object System.Collections.ArrayList
    $charts = New-Object System.Collections.ArrayList
    $order = 0

    foreach ($f in $files) {
        $t = Read-Json $f
        $taskId = [string]$t.task
        if (-not $taskId) { $taskId = [System.IO.Path]::GetFileNameWithoutExtension($f) }
        $prio = 50
        try { if ($t.priority) { $prio = [int]$t.priority } } catch { }
        foreach ($x in @($t.texts)) {
            $order++
            if ($null -eq $x) { continue }
            $slide = [int]$x.slide
            if ($slide -lt 1) { [void]$conflicts.Add($taskId + ': text op without a valid slide'); continue }
            [void]$texts.Add([pscustomobject]@{ task = $taskId; priority = $prio; order = $order; slide = $slide; shape = [string]$x.shape; expect = [string]$x.expect; text = [string]$x.text })
        }
        foreach ($r in @($t.replaces)) {
            $order++
            if ($null -eq $r) { continue }
            $find = [string]$r.find
            if (-not $find) { [void]$conflicts.Add($taskId + ': replace without find'); continue }
            [void]$replaces.Add([pscustomobject]@{ task = $taskId; priority = $prio; order = $order; rule = $r })
        }
        foreach ($fm in @($t.formats)) {
            if ($null -eq $fm) { continue }
            $order++
            [void]$formats.Add([pscustomobject]@{ task = $taskId; priority = $prio; order = $order; op = $fm })
        }
        foreach ($c in @($t.comments)) {
            $order++
            if ($null -eq $c) { continue }
            [void]$comments.Add([pscustomobject]@{ task = $taskId; priority = $prio; order = $order; slide = [int]$c.slide; shape = [string]$c.shape; text = [string]$c.text })
        }
        foreach ($g in @($t.globals)) { $order++; if ($g) { [void]$globals.Add([pscustomobject]@{ task = $taskId; priority = $prio; order = $order; op = $g }) } }
        foreach ($s in @($t.structures)) { $order++; if ($s) { [void]$structures.Add([pscustomobject]@{ task = $taskId; priority = $prio; order = $order; op = $s }) } }
        foreach ($pi in @($t.pictures)) { $order++; if ($pi) { [void]$pictures.Add([pscustomobject]@{ task = $taskId; priority = $prio; order = $order; op = $pi }) } }
        foreach ($ch in @($t.charts)) { $order++; if ($ch) { [void]$charts.Add([pscustomobject]@{ task = $taskId; priority = $prio; order = $order; op = $ch }) } }
    }

    # texts: one final text per (slide, shape)
    $textKept = New-Object System.Collections.ArrayList
    $seenText = @{}
    foreach ($x in ($texts | Sort-Object -Property priority, order)) {
        $k = ('' + $x.slide + '!' + $x.shape).ToLowerInvariant()
        if ($seenText.ContainsKey($k)) {
            if ($seenText[$k] -ne $x.text) { [void]$conflicts.Add('text conflict on slide ' + $x.slide + ' shape ' + $x.shape + ': ' + $x.task + ' dropped') }
            continue
        }
        $seenText[$k] = $x.text
        [void]$textKept.Add($x)
    }

    # replaces: two rules touching overlapping slides with the same find are a conflict
    $repKept = New-Object System.Collections.ArrayList
    foreach ($r in ($replaces | Sort-Object -Property priority, order)) {
        $clash = $null
        foreach ($prev in $repKept) {
            if ([string]$prev.rule.find -eq [string]$r.rule.find -and (Slides-Overlap ([string]$prev.rule.slides) ([string]$r.rule.slides))) { $clash = $prev; break }
        }
        if ($clash -ne $null) { [void]$conflicts.Add('replace conflict: find "' + $r.rule.find + '" from ' + $clash.task + ' and ' + $r.task + ' on overlapping slides (kept ' + $clash.task + ')'); continue }
        [void]$repKept.Add($r)
    }

    # comments: one per (slide, shape)
    $commentKept = New-Object System.Collections.ArrayList
    $seenComment = @{}
    foreach ($c in ($comments | Sort-Object -Property priority, order)) {
        $k = ('' + $c.slide + '!' + $c.shape).ToLowerInvariant()
        if ($seenComment.ContainsKey($k)) { if ($seenComment[$k] -ne $c.text) { [void]$conflicts.Add('comment conflict on slide ' + $c.slide + ' shape ' + $c.shape) }; continue }
        $seenComment[$k] = $c.text
        [void]$commentKept.Add($c)
    }

    if ($structures.Count -gt 0) { [void]$warnings.Add('structure ops present: adding/deleting/moving slides shifts slide numbers - apply them first, re-read, then replan the rest') }
    if (@($globals | Where-Object { $_.op.kind -eq 'theme' }).Count -gt 1) { [void]$warnings.Add('more than one theme op: only the last one wins') }

    $merged = [pscustomobject]@{
        doc        = $(if ($Doc) { FullPath2 $Doc } else { $null })
        stats      = [pscustomobject]@{ texts = $textKept.Count; replaces = $repKept.Count; formats = $formats.Count; comments = $commentKept.Count; globals = $globals.Count; structures = $structures.Count }
        texts      = @($textKept)
        replaces   = @($repKept)
        formats    = @($formats | Sort-Object -Property priority, order)
        comments   = @($commentKept)
        globals    = @($globals | Sort-Object -Property priority, order)
        structures = @($structures)
        pictures   = @($pictures)
        charts     = @($charts)
        conflicts  = @($conflicts)
        warnings   = @($warnings)
    }
    Write-Json (Join-Path $Out 'merged.json') $merged
    Write-Output ('merge done: texts=' + $textKept.Count + ' replaces=' + $repKept.Count + ' formats=' + $formats.Count + ' comments=' + $commentKept.Count + ' globals=' + $globals.Count + ' conflicts=' + $conflicts.Count)
    foreach ($c in $conflicts) { Write-Output ('  conflict: ' + $c) }
}

# ---------------------------------------------------------------- plan

function New-GlobalOp([string]$kind, $op, [string]$docPath) {
    $h = [ordered]@{ cmd = $kind; doc = $docPath }
    foreach ($p in $op.PSObject.Properties) { if ($p.Name -in @('kind')) { continue }; $h[$p.Name] = $p.Value }
    return [pscustomobject]$h
}

function Invoke-Plan {
    if (-not $Merged) { throw 'plan needs -Merged <merged.json>' }
    if (-not $Out) { throw 'plan needs -Out <dir>' }
    $m = Read-Json $Merged
    $docPath = if ($Doc) { FullPath2 $Doc } elseif ($m.doc) { [string]$m.doc } else { $null }
    New-Item -ItemType Directory -Force -Path $Out | Out-Null

    $textItems = @($m.texts | ForEach-Object { $h = [ordered]@{ slide = [int]$_.slide }; if ($_.shape) { $h['shape'] = [string]$_.shape }; if ($_.expect) { $h['expect'] = [string]$_.expect }; $h['text'] = [string]$_.text; [pscustomobject]$h })
    $replaceItems = @($m.replaces | ForEach-Object { $_.rule })
    $commentItems = @($m.comments | ForEach-Object { $h = [ordered]@{ slide = [int]$_.slide }; if ($_.shape) { $h['shape'] = [string]$_.shape }; $h['text'] = [string]$_.text; [pscustomobject]$h })
    Write-Json (Join-Path $Out 'text-items.json') ([pscustomobject]@{ items = $textItems })
    Write-Json (Join-Path $Out 'replace-items.json') ([pscustomobject]@{ items = $replaceItems })
    Write-Json (Join-Path $Out 'comment-items.json') ([pscustomobject]@{ items = $commentItems })
    if (@($m.structures).Count -gt 0) { Write-Json (Join-Path $Out 'structure-ops.json') ([pscustomobject]@{ items = @($m.structures) }) }
    if (@($m.pictures).Count -gt 0) { Write-Json (Join-Path $Out 'picture-ops.json') ([pscustomobject]@{ items = @($m.pictures) }) }
    if (@($m.charts).Count -gt 0) { Write-Json (Join-Path $Out 'chart-ops.json') ([pscustomobject]@{ items = @($m.charts) }) }

    $phases = @()
    if (@($m.structures).Count -gt 0) { $phases += 'structures (first; then re-read and re-plan)' }
    if (@($m.globals).Count -gt 0) { $phases += 'globals (theme/master/font; repaints every slide)' }
    if ($textItems.Count -gt 0) { $phases += 'text' }
    if ($replaceItems.Count -gt 0) { $phases += 'replace' }
    if (@($m.pictures).Count -gt 0) { $phases += 'pictures' }
    if (@($m.charts).Count -gt 0) { $phases += 'charts' }
    if ($commentItems.Count -gt 0) { $phases += 'comments' }
    $phases += 'check + render (verify)'
    Write-Json (Join-Path $Out 'apply-order.json') ([pscustomobject]@{ doc = $docPath; phases = @($phases); files = [pscustomobject]@{ text = (Join-Path $Out 'text-items.json'); replace = (Join-Path $Out 'replace-items.json'); comments = (Join-Path $Out 'comment-items.json') } })

    if ($BatchOps) {
        $batch = New-Object System.Collections.ArrayList
        foreach ($g in @($m.globals)) { [void]$batch.Add((New-GlobalOp ([string]$g.op.kind) $g.op $docPath)) }
        if ($textItems.Count -gt 0) { [void]$batch.Add([pscustomobject]@{ cmd = 'text'; doc = $docPath; items = $textItems }) }
        if ($replaceItems.Count -gt 0) { [void]$batch.Add([pscustomobject]@{ cmd = 'replace'; doc = $docPath; items = $replaceItems }) }
        if ($commentItems.Count -gt 0) { [void]$batch.Add([pscustomobject]@{ cmd = 'comment'; doc = $docPath; items = $commentItems }) }
        Write-Json (Join-Path $Out 'batch-ops.json') @($batch)
        Write-Output ('batch ops: ' + $batch.Count + ' op(s) -> ' + (Join-Path $Out 'batch-ops.json'))
    }
    Write-Output ('plan done: texts=' + $textItems.Count + ' replaces=' + $replaceItems.Count + ' comments=' + $commentItems.Count + ' globals=' + @($m.globals).Count)
    foreach ($p in $phases) { Write-Output ('  phase: ' + $p) }
}

# ---------------------------------------------------------------- selftest

function Invoke-Selftest {
    $fails = New-Object System.Collections.ArrayList
    $s = Parse-SlideSet '1-3,5'
    if ($null -eq $s -or $s.Count -ne 4 -or $s[0] -ne 1 -or $s[2] -ne 3 -or $s[3] -ne 5) { [void]$fails.Add('parse slide set 1-3,5') }
    if ($null -ne (Parse-SlideSet 'all')) { [void]$fails.Add('all should map to null') }
    if (-not (Slides-Overlap '1-5' '4-8')) { [void]$fails.Add('overlap 1-5/4-8 should be true') }
    if ((Slides-Overlap '1-3' '5-8')) { [void]$fails.Add('overlap 1-3/5-8 should be false') }
    if (-not (Slides-Overlap '1-3' 'all')) { [void]$fails.Add('overlap with all should be true') }
    if ($fails.Count -gt 0) { Write-Output ('SELFTEST FAIL: ' + ($fails -join '; ')); exit 1 }
    Write-Output 'SELFTEST PASS'
}

switch ($Command) {
    'split' { Invoke-Split }
    'merge' { Invoke-Merge }
    'plan' { Invoke-Plan }
    'selftest' { Invoke-Selftest }
    default {
        Write-Output 'usage: tasks-ppt.ps1 split|merge|plan|selftest [options]'
        Write-Output '  split  -Doc <pptx> -Ranges ranges.json -Out dir'
        Write-Output '  merge  -TaskFiles dir|files -Out dir'
        Write-Output '  plan   -Merged merged.json -Out dir [-Doc abs-pptx] [-BatchOps]'
    }
}
