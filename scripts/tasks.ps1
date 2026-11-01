# wps-office-cooperate multi-task helpers (v1.20).
#
# Pure ASCII source: every Chinese string travels inside UTF-8 JSON files, never in this script
# (PowerShell 5.1 reads a BOM-less .ps1 as ANSI, so a Chinese literal here would come out mojibake).
#
# The point of this file: "parallel planning, merged writing" for multi-task sessions.
#
#   tasks.ps1 split   -Read read.json -Ranges ranges.json -Out <dir>
#       read.json  = the JSON that `word.ps1 read -Doc X -Count 0 -SkipTableCells` prints
#       ranges.json = [{"id":"ch3","title":"...","from":132,"to":180}, ...]
#       -> <dir>/chapter-<id>.json  (per-chapter slice handed to a sub-agent; no COM needed)
#          <dir>/index.json         (sizes + validation: overlaps/gaps)
#
#   tasks.ps1 merge   -Read read.json -TaskFiles <dir or file1,file2> -Out <dir> [-MinSim 0.6]
#       task files: {"task":"T1","priority":30,"textOps":[{para,expect,find,replace}|{para,expect,text}],
#                    "commentOps":[{para,anchor,text}],"tableOps":[...],"formatOps":[...],"findings":[...]}
#       -> <dir>/merged.json  (one item per paragraph, conflicts listed, guards applied)
#
#   tasks.ps1 plan    -Merged <dir>/merged.json -Ranges ranges.json -Out <dir> [-BatchOps]
#       -> <dir>/items-<id>.json per chapter (descending para order), comments.json,
#          apply-order.json (chapters in reverse document order), optional batch-ops.json
#
#   tasks.ps1 selftest -Read read.json -Ranges ranges.json
#       quick structural checks (also used by tests/multi-task/check_tasks.ps1)

param(
    [Parameter(Mandatory = $true, Position = 0)]
    [ValidateSet('split', 'merge', 'plan', 'ranges', 'selftest', 'help')]
    [string]$Command,

    [string]$Read,
    [string]$Ranges,
    [string]$TaskFiles,
    [string]$Merged,
    [string]$Doc,
    [string]$Out,
    [double]$MinSim = 0.55,
    [string]$Styles,
    [string]$Pattern,
    [string]$SkipIfContains,
    [switch]$BatchOps,
    [switch]$IncludeTableParas
)

$ErrorActionPreference = 'Stop'

# ---------------------------------------------------------------- JSON helpers

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

function Get-FullPath2([string]$p) {
    return [System.IO.Path]::GetFullPath($p)
}

# ---------------------------------------------------------------- text helpers

function Get-Digits([string]$s) {
    $list = New-Object System.Collections.ArrayList
    foreach ($m in [regex]::Matches($s, '[0-9]+(\.[0-9]+)?')) { [void]$list.Add($m.Value) }
    return @($list | Sort-Object)
}

function Test-SameDigits([string]$a, [string]$b) {
    $da = (Get-Digits $a) -join '|'
    $db = (Get-Digits $b) -join '|'
    return ($da -eq $db)
}

function Get-Similarity([string]$a, [string]$b) {
    if ($a -eq $b) { return 1.0 }
    if ($null -eq $a -or $null -eq $b) { return 0.0 }
    $n = $a.Length; $m = $b.Length
    if ($n -eq 0 -or $m -eq 0) { return 0.0 }
    if ($n * $m -gt 4000000) { return 0.5 }   # too big to compare cheaply; do not block
    $prev = New-Object int[] ($m + 1)
    $cur = New-Object int[] ($m + 1)
    for ($j = 0; $j -le $m; $j++) { $prev[$j] = $j }
    for ($i = 1; $i -le $n; $i++) {
        $cur[0] = $i
        for ($j = 1; $j -le $m; $j++) {
            $cost = if ($a[$i - 1] -eq $b[$j - 1]) { 0 } else { 1 }
            $v = $prev[$j - 1] + $cost
            if (($prev[$j] + 1) -lt $v) { $v = $prev[$j] + 1 }
            if (($cur[$j - 1] + 1) -lt $v) { $v = $cur[$j - 1] + 1 }
            $cur[$j] = $v
        }
        $tmp = $prev; $prev = $cur; $cur = $tmp
    }
    return (1.0 - ([double]$prev[$m] / [double][Math]::Max($n, $m)))
}

# Smallest single-span replacement turning $before into $after: {find, replace}.
function Get-MinimalReplace([string]$before, [string]$after) {
    # Smallest single-span replacement that turns $before into $after, with enough context that
    # the text occurs exactly once in the paragraph (revise replaces the first hit only).
    $n = [Math]::Min($before.Length, $after.Length)
    $p = 0
    while ($p -lt $n -and $before[$p] -eq $after[$p]) { $p++ }
    $s = 0
    while ($s -lt ($n - $p) -and $before[$before.Length - 1 - $s] -eq $after[$after.Length - 1 - $s]) { $s++ }
    if (($before.Length - $p - $s) -eq 0 -and $before.Length -gt 0) {
        # pure insertion: borrow one character of context so find is not empty
        if ($p -gt 0) { $p-- } else { $s-- }
    }
    $find = ''
    $repl = ''
    for ($expand = 0; $expand -lt 32; $expand++) {
        $len = $before.Length - $p - $s
        $find = $before.Substring($p, $len)
        $repl = $after.Substring($p, $after.Length - $s - $p)
        if ($len -gt 0 -and $before.IndexOf($find) -eq $p -and $before.LastIndexOf($find) -eq $p) { break }
        if ($p -gt 0 -and ($expand % 2 -eq 0)) { $p-- }
        elseif ($s -gt 0) { $s-- }
        elseif ($p -gt 0) { $p-- }
        else { break }
    }
    return @{ find = $find; replace = $repl; at = $p }
}
# ---------------------------------------------------------------- split

function Get-ReadTexts($readJson) {
    $map = @{}
    foreach ($it in @($readJson.result.items)) {
        $map[[int]$it.i] = [string]$it.text
    }
    return $map
}

function Invoke-Split {
    if (-not $Read) { throw 'split needs -Read <read.json>' }
    if (-not $Ranges) { throw 'split needs -Ranges <ranges.json>' }
    if (-not $Out) { throw 'split needs -Out <dir>' }
    $readJson = Read-Json $Read
    $rangeList = @(Read-Json $Ranges)
    if ($rangeList.Count -eq 0) { throw 'ranges.json is empty' }
    $items = @($readJson.result.items)
    $byPara = Get-ReadTexts $readJson
    $total = 0
    try { $total = [int]$readJson.result.paragraphs } catch { $total = 0 }

    $summary = New-Object System.Collections.ArrayList
    $problems = New-Object System.Collections.ArrayList
    $sorted = @($rangeList | Sort-Object -Property @{ Expression = { [int]$_.from } })
    for ($k = 0; $k -lt $sorted.Count; $k++) {
        $r = $sorted[$k]
        $from = [int]$r.from; $to = [int]$r.to
        if ($from -lt 1 -or $to -lt $from) { [void]$problems.Add("range $($r.id): bad from/to ($from..$to)"); continue }
        if ($to -gt $total -and $total -gt 0) { [void]$problems.Add("range $($r.id): to=$to exceeds paragraphs=$total") }
        if ($k -gt 0 -and $from -le [int]$sorted[$k - 1].to) {
            [void]$problems.Add("range $($r.id) overlaps previous range $($sorted[$k - 1].id)")
        }
        $paras = @($items | Where-Object { [int]$_.i -ge $from -and [int]$_.i -le $to } | Sort-Object -Property @{ Expression = { [int]$_.i } })
        $slice = [pscustomobject]@{
            id      = [string]$r.id
            title   = [string]$r.title
            from    = $from
            to      = $to
            count   = $paras.Count
            paras   = @($paras | ForEach-Object { [pscustomobject]@{ i = [int]$_.i; text = [string]$_.text; style = [string]$_.style } })
        }
        $path = Join-Path $Out ('chapter-' + $r.id + '.json')
        Write-Json $path $slice
        [void]$summary.Add([pscustomobject]@{ id = [string]$r.id; from = $from; to = $to; count = $paras.Count; file = $path })
        Write-Output ("chapter " + $r.id + ": " + $paras.Count + " paragraphs (" + $from + ".." + $to + ") -> " + $path)
    }
    $index = [pscustomobject]@{
        read     = (Get-FullPath2 $Read)
        total    = $total
        returned = $items.Count
        chapters = @($summary)
        problems = @($problems)
    }
    Write-Json (Join-Path $Out 'index.json') $index
    if ($problems.Count -gt 0) {
        Write-Output ('WARN range problems: ' + ($problems -join '; '))
    }
    Write-Output ('split done: ' + $summary.Count + ' chapters, ' + $items.Count + ' paragraphs in read.json')
}

# ---------------------------------------------------------------- merge

function Invoke-Merge {
    if (-not $Read) { throw 'merge needs -Read <read.json> (for the original texts)' }
    if (-not $TaskFiles) { throw 'merge needs -TaskFiles <dir or comma list of task json files>' }
    if (-not $Out) { throw 'merge needs -Out <dir>' }
    $readJson = Read-Json $Read
    $byPara = Get-ReadTexts $readJson

    $files = @()
    if (Test-Path -LiteralPath $TaskFiles -PathType Container) {
        $files = @(Get-ChildItem -LiteralPath $TaskFiles -Filter 'task-*.json' | Sort-Object Name | ForEach-Object { $_.FullName })
    }
    else {
        $files = @($TaskFiles.Split(',') | ForEach-Object { $_.Trim() } | Where-Object { $_ })
    }
    if ($files.Count -eq 0) { throw 'no task-*.json found in ' + $TaskFiles }

    $conflicts = New-Object System.Collections.ArrayList
    $textByPara = @{}       # para -> list of {task, priority, order, op}
    $comments = New-Object System.Collections.ArrayList
    $tableOps = New-Object System.Collections.ArrayList
    $formatOps = New-Object System.Collections.ArrayList
    $findings = New-Object System.Collections.ArrayList
    $order = 0

    foreach ($f in $files) {
        $t = Read-Json $f
        $taskId = [string]$t.task
        if (-not $taskId) { $taskId = [System.IO.Path]::GetFileNameWithoutExtension($f) }
        $prio = 50
        try { if ($t.priority) { $prio = [int]$t.priority } } catch { }
        $seen = @{}
        foreach ($op in @($t.textOps)) {
            $order++
            if ($null -eq $op) { continue }
            $para = 0
            try { $para = [int]$op.para } catch { $para = 0 }
            if ($para -lt 1) { [void]$conflicts.Add("${taskId}: textOp without a valid para"); continue }
            if (-not $byPara.ContainsKey($para)) {
                [void]$conflicts.Add("${taskId}: paragraph $para is not in read.json (table cell or out of range)")
                continue
            }
            if ($seen.ContainsKey($para)) {
                [void]$conflicts.Add("${taskId}: more than one textOp for paragraph $para (merge them on the planning side)")
                continue
            }
            $seen[$para] = $true
            $expect = [string]$op.expect
            $orig = [string]$byPara[$para]
            if ($expect -and -not $orig.StartsWith($expect)) {
                [void]$conflicts.Add("${taskId}: paragraph $para expect does not match (expect='" + $expect + "')")
                continue
            }
            if (-not $textByPara.ContainsKey($para)) { $textByPara[$para] = New-Object System.Collections.ArrayList }
            [void]$textByPara[$para].Add([pscustomobject]@{ task = $taskId; priority = $prio; order = $order; op = $op })
        }
        foreach ($c in @($t.commentOps)) {
            if ($null -eq $c) { continue }
            $para = [int]$c.para
            if (-not $byPara.ContainsKey($para)) { [void]$conflicts.Add("${taskId}: comment on paragraph $para which is not in read.json"); continue }
            $anchor = [string]$c.anchor
            if ($anchor -and -not ([string]$byPara[$para]).Contains($anchor)) {
                [void]$conflicts.Add("${taskId}: comment anchor not found in paragraph $para"); continue
            }
            [void]$comments.Add([pscustomobject]@{ para = $para; anchor = $anchor; text = [string]$c.text; task = $taskId })
        }
        foreach ($tb in @($t.tableOps)) { if ($tb) { [void]$tableOps.Add($tb) } }
        foreach ($fo in @($t.formatOps)) { if ($fo) { [void]$formatOps.Add($fo) } }
        foreach ($fi in @($t.findings)) { if ($fi) { [void]$findings.Add([pscustomobject]@{ task = $taskId; text = [string]$fi }) } }
    }

    # per-paragraph merge: apply the contributing ops in priority order and rebuild ONE item.
    $outOps = New-Object System.Collections.ArrayList
    $stats = @{ paragraphs = 0; merged = 0; rewritten = 0; digitConflicts = 0; lowSim = 0 }
    foreach ($para in ($textByPara.Keys | Sort-Object { [int]$_ })) {
        $contrib = @($textByPara[$para] | Sort-Object -Property priority, order)
        $orig = [string]$byPara[$para]
        $cur = $orig
        $ok = $true
        $allowDigits = $false
        $allowRewrite = $false
        $fromTasks = New-Object System.Collections.ArrayList
        $hasSetText = $false
        foreach ($c in $contrib) {
            [void]$fromTasks.Add($c.task)
            $op = $c.op
            if ($op.allowDigits -eq $true) { $allowDigits = $true }
            if ($op.allowRewrite -eq $true) { $allowRewrite = $true }
            $kind = [string]$op.op
            if ($kind -eq 'set_text' -or ($op.PSObject.Properties.Name -contains 'text' -and -not ($op.PSObject.Properties.Name -contains 'find'))) {
                $hasSetText = $true
                $cur = [string]$op.text
            }
            elseif ($kind -eq 'replace') {
                $find = [string]$op.find
                $repl = [string]$op.replace
                $idx = $cur.IndexOf($find, [System.StringComparison]::Ordinal)
                if ($idx -lt 0) {
                    [void]$conflicts.Add("paragraph ${para}: task $($c.task) find='" + $find + "' is no longer present after another task's edit")
                    $ok = $false
                    break
                }
                $cur = $cur.Substring(0, $idx) + $repl + $cur.Substring($idx + $find.Length)
            }
            else {
                [void]$conflicts.Add("paragraph ${para}: task $($c.task) has unsupported op '" + $kind + "' (use replace or set_text)")
                $ok = $false
                break
            }
        }
        if (-not $ok) { continue }
        $stats.paragraphs++
        if ($fromTasks.Count -gt 1) { $stats.merged++ }

        if (-not $allowDigits -and -not (Test-SameDigits $orig $cur)) {
            [void]$conflicts.Add("paragraph ${para}: digit tokens changed (" + $fromTasks.Count + " task(s)); add allowDigits=true if this is intended")
            $stats.digitConflicts++
            continue
        }
        $sim = Get-Similarity $orig $cur
        if (-not $allowRewrite -and $sim -lt $MinSim) {
            [void]$conflicts.Add("paragraph ${para}: rewrite too large (similarity " + [Math]::Round($sim, 3) + " < " + $MinSim + "); add allowRewrite=true if this is intended")
            $stats.lowSim++
            continue
        }
        if ($cur -eq $orig) { continue }   # nothing to do

        $span = Get-MinimalReplace $orig $cur
        $item = [ordered]@{
            para      = [int]$para
            expect    = $orig.Substring(0, [Math]::Min(15, $orig.Length))
            op        = 'replace'
            find      = $span.find
            replace   = $span.replace
            mergedFrom = @($fromTasks)
        }
        if ($span.find.Length -eq 0) {
            # insertion at one point: replace an empty find is not allowed; use set_text instead
            $item = [ordered]@{ para = [int]$para; expect = $orig.Substring(0, [Math]::Min(15, $orig.Length)); op = 'set_text'; text = $cur; mergedFrom = @($fromTasks) }
            $stats.rewritten++
        }
        elseif ($hasSetText -and $fromTasks.Count -gt 1) {
            # a task asked for a whole-paragraph rewrite and another task edited it: the merged
            # span may flatten run-level formatting, so say so explicitly.
            $item['formatRisk'] = $true
            $stats.rewritten++
        }
        if ($span.replace.Length -gt 200) { $item['largeSpan'] = $span.replace.Length }
        [void]$outOps.Add([pscustomobject]$item)
    }

    # tables: run before text (they change paragraph numbering); order by para descending.
    $tableList = @($tableOps | Sort-Object -Property @{ Expression = { [int]($_.para) } } -Descending)
    $formatList = @($formatOps | Sort-Object -Property @{ Expression = { [int]($_.para) } } -Descending)
    $commentList = @($comments | Sort-Object -Property @{ Expression = { [int]($_.para) } } -Descending)

    $merged = [pscustomobject]@{
        read      = (Get-FullPath2 $Read)
        stats     = [pscustomobject]$stats
        textOps   = @($outOps | Sort-Object -Property @{ Expression = { [int]$_.para } } -Descending)
        tableOps  = $tableList
        formatOps = $formatList
        comments  = $commentList
        findings  = @($findings)
        conflicts = @($conflicts)
    }
    Write-Json (Join-Path $Out 'merged.json') $merged
    Write-Output ('merge done: ' + $outOps.Count + ' paragraphs to write, ' + $conflicts.Count + ' conflict(s), ' + $comments.Count + ' comment(s)')
    if ($conflicts.Count -gt 0) { foreach ($c in $conflicts) { Write-Output ('  conflict: ' + $c) } }
}

# ---------------------------------------------------------------- plan

function New-BatchOp([string]$cmd, $items) {
    $h = [ordered]@{ cmd = $cmd }
    if ($Doc) { $h['doc'] = $Doc }
    $h['items'] = @($items)
    return [pscustomobject]$h
}

function Invoke-Plan {
    if (-not $Merged) { throw 'plan needs -Merged <merged.json>' }
    if (-not $Ranges) { throw 'plan needs -Ranges <ranges.json>' }
    if (-not $Out) { throw 'plan needs -Out <dir>' }
    $merged = Read-Json $Merged
    $rangeList = @(Read-Json $Ranges)
    $ops = @($merged.textOps)
    $used = @{}

    $summary = New-Object System.Collections.ArrayList
    $sortedRanges = @($rangeList | Sort-Object -Property @{ Expression = { [int]$_.from } } -Descending)
    foreach ($r in $sortedRanges) {
        $from = [int]$r.from; $to = [int]$r.to
        $mine = @($ops | Where-Object { [int]$_.para -ge $from -and [int]$_.para -le $to } | Sort-Object -Property @{ Expression = { [int]$_.para } } -Descending)
        foreach ($o in $mine) { $used[[int]$o.para] = $true }
        if ($mine.Count -eq 0) { continue }
        $path = Join-Path $Out ('items-' + $r.id + '.json')
        Write-Json $path ([pscustomobject]@{ chapter = [string]$r.id; from = $from; to = $to; items = @($mine) })
        [void]$summary.Add([pscustomobject]@{ id = [string]$r.id; from = $from; to = $to; count = $mine.Count; file = $path })
        Write-Output ('items ' + $r.id + ': ' + $mine.Count + ' paragraph(s) -> ' + $path)
    }
    $rest = @($ops | Where-Object { -not $used.ContainsKey([int]$_.para) } | Sort-Object -Property @{ Expression = { [int]$_.para } } -Descending)
    if ($rest.Count -gt 0) {
        $path = Join-Path $Out 'items-outside-ranges.json'
        Write-Json $path ([pscustomobject]@{ chapter = 'outside'; from = 0; to = 0; items = @($rest) })
        Write-Output ('items outside ranges: ' + $rest.Count + ' -> ' + $path)
    }
    $applyOrder = [pscustomobject]@{
        note        = 'chapters are applied from the END of the document backwards; each file is already sorted by para descending'
        chapters    = @($summary)
        comments    = (Join-Path $Out 'comments.json')
        outside     = $(if ($rest.Count -gt 0) { Join-Path $Out 'items-outside-ranges.json' } else { $null })
    }
    Write-Json (Join-Path $Out 'comments.json') ([pscustomobject]@{ items = @($merged.comments) })
    Write-Json (Join-Path $Out 'apply-order.json') $applyOrder

    if ($BatchOps) {
        $batch = New-Object System.Collections.ArrayList
        foreach ($s in $summary) {
            $file = Read-Json $s.file
            [void]$batch.Add((New-BatchOp 'revise' @($file.items)))
        }
        if ($rest.Count -gt 0) { [void]$batch.Add((New-BatchOp 'revise' @($rest))) }
        if (@($merged.comments).Count -gt 0) { [void]$batch.Add((New-BatchOp 'comment' @($merged.comments))) }
        Write-Json (Join-Path $Out 'batch-ops.json') @($batch)
        Write-Output ('batch ops: ' + $batch.Count + ' op(s) -> ' + (Join-Path $Out 'batch-ops.json'))
    }
}

# ---------------------------------------------------------------- ranges

# Build ranges.json from the heading paragraphs of a read.json.
#   -Styles "标题 1"        : comma list of style names, highest level first (chapter level = first)
#   -Pattern "<regex>"      : fallback for documents without heading styles (matches the whole title)
#   -SkipIfContains "..."   : comma list of substrings that mark a table-of-contents line (skip it)
function Invoke-Ranges {
    if (-not $Read) { throw 'ranges needs -Read <read.json>' }
    if (-not $Out) { throw 'ranges needs -Out <ranges.json>' }
    $readJson = Read-Json $Read
    $items = @($readJson.result.items)
    $styleList = @()
    if ($Styles) { $styleList = @($Styles.Split(',') | ForEach-Object { $_.Trim() } | Where-Object { $_ }) }
    if ($styleList.Count -eq 0 -and -not $Pattern) { throw 'ranges needs -Styles "标题 1" or -Pattern "<regex>"' }
    $skip = @()
    if ($SkipIfContains) { $skip = @($SkipIfContains.Split('|') | Where-Object { $_ }) }

    $levelOf = @{}
    for ($k = 0; $k -lt $styleList.Count; $k++) { $levelOf[$styleList[$k]] = $k + 1 }
    $headings = New-Object System.Collections.ArrayList
    foreach ($it in ($items | Sort-Object -Property @{ Expression = { [int]$_.i } })) {
        $style = [string]$it.style
        $text = ([string]$it.text).Trim()
        if ($text.Length -eq 0) { continue }
        $hit = $false
        if ($styleList.Count -gt 0) { $hit = $styleList -contains $style }
        if (-not $hit -and $Pattern) { $hit = ($text -match $Pattern) }
        if (-not $hit) { continue }
        $skipThis = $false
        foreach ($s in $skip) { if ($text.Contains($s)) { $skipThis = $true } }
        if ($skipThis) { continue }
        $level = 99
        if ($levelOf.ContainsKey($style)) { $level = $levelOf[$style] }
        [void]$headings.Add([pscustomobject]@{ i = [int]$it.i; text = $text; level = $level; style = $style })
    }
    if ($headings.Count -eq 0) { throw 'no heading paragraph matched -Styles / -Pattern' }
    $total = 0
    try { $total = [int]$readJson.result.paragraphs } catch { $total = 0 }
    if ($total -le 0) { $total = [int]($items | Measure-Object -Property @{ Expression = { [int]$_.i } } -Maximum).Maximum }

    # chapters = first-level headings (level 1 when styles are given, otherwise every match)
    $chapters = @($headings | Where-Object { $_.level -eq 1 })
    if ($chapters.Count -eq 0) { $chapters = @($headings) }
    $rangeOut = New-Object System.Collections.ArrayList
    for ($k = 0; $k -lt $chapters.Count; $k++) {
        $from = [int]$chapters[$k].i
        $to = $total
        if ($k -lt $chapters.Count - 1) { $to = [int]$chapters[$k + 1].i - 1 }
        [void]$rangeOut.Add([ordered]@{ id = ('ch' + ($k + 1)); title = [string]$chapters[$k].text; from = $from; to = $to })
    }
    Write-Json $Out @($rangeOut)
    Write-Output ('ranges: ' + $out.Count + ' chapter(s) -> ' + $Out)
    foreach ($r in $rangeOut) { Write-Output ('  ' + $r.id + ' ' + $r.from + '..' + $r.to + '  ' + $r.title) }
}
# ---------------------------------------------------------------- selftest

function Invoke-Selftest {
    if (-not $Read) { throw 'selftest needs -Read <read.json>' }
    if (-not $Ranges) { throw 'selftest needs -Ranges <ranges.json>' }
    $readJson = Read-Json $Read
    $ranges = @(Read-Json $Ranges)
    $items = @($readJson.result.items)
    $fails = New-Object System.Collections.ArrayList
    if ($items.Count -eq 0) { [void]$fails.Add('read.json has no items') }
    foreach ($r in $ranges) {
        if ([int]$r.from -lt 1 -or [int]$r.to -lt [int]$r.from) { [void]$fails.Add("bad range $($r.id)") }
        $n = @($items | Where-Object { [int]$_.i -ge [int]$r.from -and [int]$_.i -le [int]$r.to }).Count
        if ($n -eq 0) { [void]$fails.Add("range $($r.id) is empty") }
    }
    $sorted = @($ranges | Sort-Object -Property @{ Expression = { [int]$_.from } })
    for ($k = 1; $k -lt $sorted.Count; $k++) {
        if ([int]$sorted[$k].from -le [int]$sorted[$k - 1].to) { [void]$fails.Add("range $($sorted[$k].id) overlaps $($sorted[$k - 1].id)") }
    }
    # merge engine smoke test (no files): two edits on one paragraph must merge into one item
    $probe = [pscustomobject]@{ result = [pscustomobject]@{ items = @([pscustomobject]@{ i = 1; text = 'abc 123 def' }); paragraphs = 1 } }
    $byPara = Get-ReadTexts $probe
    $orig = [string]$byPara[1]
    $cur = $orig.Replace('abc', 'abcX').Replace('def', 'defY')
    $span = Get-MinimalReplace $orig $cur
    $round = $cur.Substring(0, $span.find.Length)
    if (-not ($orig.Substring(0, $orig.IndexOf($span.find)) + $span.replace + $orig.Substring($orig.IndexOf($span.find) + $span.find.Length) -eq $cur)) {
        [void]$fails.Add('minimal replace does not round-trip')
    }
    if (-not (Test-SameDigits 'abc 123 def' 'abc 123 defY')) { [void]$fails.Add('digit check false negative') }
    if (Test-SameDigits 'abc 123 def' 'abc 124 def') { [void]$fails.Add('digit check missed a change') }
    if ((Get-Similarity 'abcdef' 'abcdef') -ne 1.0) { [void]$fails.Add('similarity identity failed') }
    if ([Math]::ABS((Get-Similarity 'abcdef' 'abcxef') - (5.0 / 6.0)) -gt 0.001) { [void]$fails.Add('similarity value wrong') }

    if ($fails.Count -gt 0) {
        Write-Output ('SELFTEST FAIL: ' + ($fails -join '; '))
        exit 1
    }
    Write-Output ('SELFTEST PASS: ' + $ranges.Count + ' ranges, ' + $items.Count + ' paragraphs')
}

# ---------------------------------------------------------------- main

switch ($Command) {
    'split' { Invoke-Split }
    'merge' { Invoke-Merge }
    'plan' { Invoke-Plan }
    'ranges' { Invoke-Ranges }
    'selftest' { Invoke-Selftest }
    default {
        Write-Output 'usage: tasks.ps1 split|merge|plan|ranges|selftest [options]'
        Write-Output '  split    -Read read.json -Ranges ranges.json -Out dir'
        Write-Output '  merge    -Read read.json -TaskFiles dir|files -Out dir [-MinSim 0.55]'
        Write-Output '  plan     -Merged merged.json -Ranges ranges.json -Out dir [-BatchOps] [-Doc abs-path]
  ranges   -Read read.json -Out ranges.json -Styles "<style1>,<style2>" [-Pattern <regex>] [-SkipIfContains <substr1|substr2>]'
        Write-Output '  selftest -Read read.json -Ranges ranges.json'
    }
}
