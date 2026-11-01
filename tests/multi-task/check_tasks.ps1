# Multi-task helper regression (v1.20): split / merge / plan over the report-300 corpus.
#
# Run:  powershell -NoProfile -ExecutionPolicy Bypass -File tests\multi-task\check_tasks.ps1
# Needs WPS (it runs one real `word.ps1 read` on a temp copy of the corpus). No changes are saved.
#
# What it proves (all of these are properties the v1.20 design depends on):
#   1. split covers every paragraph of the read exactly once (no gaps / overlaps);
#   2. merge collapses edits from two tasks on the SAME paragraph into ONE item (the revise
#      command rejects two items per paragraph - measured 2026-11-02);
#   3. the produced find string occurs exactly once in the paragraph (revise replaces the first hit);
#   4. merge keeps the digit multiset and rejects an accidental digit change;
#   5. plan emits items in descending paragraph order and a batch-ops file with doc stamped.

$ErrorActionPreference = 'Stop'
$here = $PSScriptRoot
$repo = Split-Path -Parent (Split-Path -Parent $here)
$enc = New-Object System.Text.UTF8Encoding($false)
$fails = New-Object System.Collections.ArrayList
$work = Join-Path $env:TEMP ('woc-multitask-' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Force -Path $work | Out-Null

function Check([bool]$cond, [string]$what) {
    if ($cond) { Write-Output ('PASS ' + $what) }
    else { Write-Output ('FAIL ' + $what); [void]$fails.Add($what) }
}
function ReadJsonFile([string]$p) { return ([System.IO.File]::ReadAllText($p, [System.Text.Encoding]::UTF8) | ConvertFrom-Json) }
function WriteJsonFile([string]$p, $o) { [System.IO.File]::WriteAllText($p, (ConvertTo-Json -InputObject $o -Depth 20 -Compress), $enc) }

try {
    $corpus = Join-Path $repo 'tests\word\corpus\report-300.pristine.docx'
    if (-not (Test-Path -LiteralPath $corpus)) { throw ('corpus not found: ' + $corpus) }
    $doc = Join-Path $work 'doc.docx'
    Copy-Item -LiteralPath $corpus -Destination $doc -Force

    $word = Join-Path $repo 'scripts\word.ps1'
    $tasks = Join-Path $repo 'scripts\tasks.ps1'

    function RunWord([string[]]$a) {
        $p = Start-Process -FilePath 'powershell' -ArgumentList (@('-Mta', '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $word) + $a) -PassThru -NoNewWindow -RedirectStandardOutput (Join-Path $work ('out-' + [guid]::NewGuid().ToString('N') + '.json'))
        $p.WaitForExit()
        return $p
    }

    $null = RunWord @('serve', '-Action', 'start')
    $null = RunWord @('open', '-Doc', $doc, '-NoHint')
    $readPath = Join-Path $work 'read.json'
    $p = Start-Process -FilePath 'powershell' -ArgumentList @('-Mta', '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $word, 'read', '-Doc', $doc, '-Count', '0', '-SkipTableCells') -PassThru -NoNewWindow -RedirectStandardOutput $readPath
    $p.WaitForExit()
    $read = ReadJsonFile $readPath
    Check ([bool]$read.ok) 'read ok'
    $items = @($read.result.items)
    $totalReturned = $items.Count
    Check ($totalReturned -gt 200) ('read returned paragraphs (' + $totalReturned + ')')

    $rangesPath = Join-Path $here 'ranges.json'
    $ranges = @(ReadJsonFile $rangesPath)
    Check ($ranges.Count -eq 6) 'ranges.json has 6 chapters'

    # ---- split
    $splitDir = Join-Path $work 'split'
    $out = & powershell -NoProfile -ExecutionPolicy Bypass -File $tasks split -Read $readPath -Ranges $rangesPath -Out $splitDir 2>&1
    $idx = ReadJsonFile (Join-Path $splitDir 'index.json')
    Check (@($idx.problems).Count -eq 0) 'split reports no range problems'
    $sum = 0
    foreach ($c in @($idx.chapters)) { $sum += [int]$c.count }
    $covered = @($items | Where-Object {
            $hit = $false
            foreach ($r in $ranges) { if ([int]$_.i -ge [int]$r.from -and [int]$_.i -le [int]$r.to) { $hit = $true } }
            $hit
        }).Count
    Check ($sum -eq $covered) ('split covers every in-range paragraph (' + $sum + '/' + $covered + ')')
    Check ((Test-Path -LiteralPath (Join-Path $splitDir 'chapter-ch3.json'))) 'chapter-ch3.json exists'

    # ---- synthetic planners: two tasks editing the SAME paragraphs (T2 has priority 10)
    $ch3 = ReadJsonFile (Join-Path $splitDir 'chapter-ch3.json')
    $picked = @($ch3.paras | Where-Object { $_.text.Length -gt 30 } | Select-Object -First 4)
    Check ($picked.Count -eq 4) 'chapter 3 has 4 long paragraphs for the test'
    function MakeOp($para, $orig, [int]$at, [string]$mark) {
        return [ordered]@{ para = [int]$para; expect = $orig.Substring(0, [Math]::Min(15, $orig.Length)); op = 'replace'; find = $orig.Substring($at, 2); replace = ($orig.Substring($at, 2) + $mark) }
    }
    $t1 = [ordered]@{
        task = 'T1'; title = 'polish'; priority = 30
        textOps = @((MakeOp $picked[0].i $picked[0].text 4 'PA'), (MakeOp $picked[1].i $picked[1].text 6 'PB'), (MakeOp $picked[2].i $picked[2].text 8 'PC'))
        commentOps = @([ordered]@{ para = [int]$picked[0].i; anchor = $picked[0].text.Substring(0, 4); text = 'check' })
        findings = @('T1 finding')
    }
    $t2 = [ordered]@{
        task = 'T2'; title = 'fact'; priority = 10
        textOps = @((MakeOp $picked[0].i $picked[0].text 10 'FA'), (MakeOp $picked[1].i $picked[1].text 12 'FB'), (MakeOp $picked[3].i $picked[3].text 3 'FC'))
        findings = @('T2 finding')
    }
    WriteJsonFile (Join-Path $work 'task-T1.json') $t1
    WriteJsonFile (Join-Path $work 'task-T2.json') $t2

    # ---- merge
    $out = & powershell -NoProfile -ExecutionPolicy Bypass -File $tasks merge -Read $readPath -TaskFiles $work -Out $work 2>&1
    $merged = ReadJsonFile (Join-Path $work 'merged.json')
    Check (@($merged.conflicts).Count -eq 0) 'merge has no conflicts'
    Check (@($merged.textOps).Count -eq 4) ('merge produced 4 items (got ' + @($merged.textOps).Count + ')')
    $allReplace = $true
    $allUnique = $true
    $origByPara = @{}
    foreach ($it in $items) { $origByPara[[int]$it.i] = [string]$it.text }
    foreach ($op in @($merged.textOps)) {
        if ($op.op -ne 'replace' -or -not $op.find) { $allReplace = $false }
        $orig = [string]$origByPara[[int]$op.para]
        if ($orig.IndexOf([string]$op.find) -ne $orig.LastIndexOf([string]$op.find)) { $allUnique = $false }
    }
    Check $allReplace 'every merged item is a replace with a non-empty find'
    Check $allUnique 'every merged find is unique inside its paragraph'
    Check (@($merged.comments).Count -eq 1) 'merge kept the comment'
    # digit guard: an op that changes a number must be rejected unless allowDigits
    $bad = [ordered]@{ task = 'T9'; priority = 99; textOps = @([ordered]@{ para = [int]$picked[3].i; expect = $picked[3].text.Substring(0, [Math]::Min(15, $picked[3].text.Length)); op = 'set_text'; text = ($picked[3].text + ' 999'); allowDigits = $false }) }
    WriteJsonFile (Join-Path $work 'task-T9.json') $bad
    $null = & powershell -NoProfile -ExecutionPolicy Bypass -File $tasks merge -Read $readPath -TaskFiles (Join-Path $work 'task-T9.json') -Out (Join-Path $work 'guard') 2>&1
    $guarded = ReadJsonFile (Join-Path $work 'guard\merged.json')
    Check (@($guarded.conflicts).Count -ge 1) 'digit guard rejects an accidental number change'

    # ---- plan
    $out = & powershell -NoProfile -ExecutionPolicy Bypass -File $tasks plan -Merged (Join-Path $work 'merged.json') -Ranges $rangesPath -Out $work -BatchOps -Doc $doc 2>&1
    $plan = ReadJsonFile (Join-Path $work 'items-ch3.json')
    Check (@($plan.items).Count -eq 4) 'plan put 4 items into chapter 3'
    $desc = $true
    $prev = [int]::MaxValue
    foreach ($op in @($plan.items)) { if ([int]$op.para -gt $prev) { $desc = $false }; $prev = [int]$op.para }
    Check $desc 'plan items are in descending paragraph order'
    $batch = @(ReadJsonFile (Join-Path $work 'batch-ops.json'))
    Check ($batch.Count -eq 2) ('batch-ops has revise + comment (got ' + $batch.Count + ')')
    Check ([string]$batch[0].doc -eq [string]$doc) 'batch ops carry the document path'

    $null = RunWord @('close', '-Doc', $doc)
}
catch {
    [void]$fails.Add('exception: ' + $_.Exception.Message)
    Write-Output ('FAIL exception: ' + $_.Exception.Message)
}
finally {
    try { Remove-Item -Recurse -Force -LiteralPath $work -ErrorAction SilentlyContinue } catch { }
}

if ($fails.Count -gt 0) {
    Write-Output ('MULTI-TASK TEST FAIL (' + $fails.Count + ' failure(s))')
    exit 1
}
Write-Output 'MULTI-TASK TEST PASS'
exit 0
