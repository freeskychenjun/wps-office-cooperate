# PPT multi-task helper regression (v1.23): split / merge / plan / batch over a generated deck.
#
# Run:  powershell -NoProfile -ExecutionPolicy Bypass -File tests\ppt\multi-task\check_tasks.ps1
# Needs WPS (it builds a 3-slide deck, applies text/replace/comment and reads the result back).
#
# What it proves:
#   1. split reads a slide range through ppt.ps1 and stores the per-slide shapes;
#   2. merge keeps one final text per (slide, shape) and reports the conflict;
#   3. plan emits text -> replace -> comment in that order plus a batch-ops file with the deck;
#   4. one `batch` call applies all three phases;
#   5. the deck really changed: slide 1 got the replace, slides 2/3 the polished text,
#      and the conflicting text never landed.

$ErrorActionPreference = 'Stop'
$here = $PSScriptRoot
$repo = Split-Path -Parent (Split-Path -Parent (Split-Path -Parent $here))
$enc = New-Object System.Text.UTF8Encoding($false)
$work = Join-Path $env:TEMP ('woc-ppt-mt-' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Force -Path $work | Out-Null
$fails = New-Object System.Collections.ArrayList

function Check([bool]$cond, [string]$what) {
    if ($cond) { Write-Output ('PASS ' + $what) } else { Write-Output ('FAIL ' + $what); [void]$fails.Add($what) }
}
function Run([string]$ps1, [string[]]$argsList) {
    $id = [guid]::NewGuid().ToString('N')
    $out = Join-Path $work ('o-' + $id + '.json')
    $p = Start-Process -FilePath 'powershell' -ArgumentList (@('-Mta', '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', (Join-Path $repo $ps1)) + $argsList) -PassThru -NoNewWindow -RedirectStandardOutput $out -RedirectStandardError (Join-Path $work ('e-' + $id + '.txt'))
    $ok = $p.WaitForExit(180000)
    if (-not $ok) { try { $p.Kill() } catch { }; Start-Sleep -Seconds 1 }
    for ($i = 0; $i -lt 10; $i++) {
        if (Test-Path $out) { try { $t = [System.IO.File]::ReadAllText($out, [System.Text.Encoding]::UTF8); if ($t.Trim()) { return ($t | ConvertFrom-Json) } } catch { } }
        Start-Sleep -Milliseconds 300
    }
    return [pscustomobject]@{ ok = $false; error = @{ code = 'EMPTY'; message = '' } }
}
$tool = 'scripts\tasks-ppt.ps1'

try {
    $doc = Join-Path $work 'deck.pptx'
    $null = Run 'scripts\ppt.ps1' @('close', '-Doc', $doc)
    Check ([bool](Run 'scripts\ppt.ps1' @('open', '-Doc', $doc, '-NoHint')).ok) 'new deck created'
    foreach ($t in @('T1', 'T2', 'T3')) {
        Check ([bool](Run 'scripts\ppt.ps1' @('slide', '-Action', 'add', '-Doc', $doc, '-Preset', 'title', '-Title', $t, '-Subtitle', 'sub')).ok) ('slide ' + $t + ' added')
    }
    $read = Run 'scripts\ppt.ps1' @('read', '-Doc', $doc)
    Check ([bool]$read.ok -and @($read.result.items).Count -ge 3) ('deck has 3 slides (got ' + @($read.result.items).Count + ')')

    [System.IO.File]::WriteAllText((Join-Path $work 'ranges.json'), (ConvertTo-Json -InputObject @(
                [ordered]@{ id = 's1'; from = 1; to = 2 },
                [ordered]@{ id = 's2'; from = 3; to = 3 }
            ) -Depth 6), $enc)
    $null = Run $tool @('split', '-Doc', $doc, '-Ranges', (Join-Path $work 'ranges.json'), '-Out', $work)
    $slice = [System.IO.File]::ReadAllText((Join-Path $work 'slice-s1.json'), [System.Text.Encoding]::UTF8) | ConvertFrom-Json
    Check (@($slice.slides).Count -eq 2) 'split slice holds 2 slides'

    $p1 = [ordered]@{
        task = 'P1'; priority = 10
        texts = @([ordered]@{ slide = 2; shape = 'title'; expect = 'T2'; text = 'T2-polished' })
        replaces = @([ordered]@{ find = 'T1'; replace = 'T1-new'; scope = 'all'; slides = '1' })
        comments = @([ordered]@{ slide = 1; shape = 'title'; text = 'multi-task check' })
    }
    $p2 = [ordered]@{
        task = 'P2'; priority = 30
        texts = @(
            [ordered]@{ slide = 3; shape = 'title'; expect = 'T3'; text = 'T3-polished' },
            [ordered]@{ slide = 2; shape = 'title'; expect = 'T2'; text = 'CONFLICT' }
        )
    }
    [System.IO.File]::WriteAllText((Join-Path $work 'task-P1.json'), (ConvertTo-Json -InputObject $p1 -Depth 10), $enc)
    [System.IO.File]::WriteAllText((Join-Path $work 'task-P2.json'), (ConvertTo-Json -InputObject $p2 -Depth 10), $enc)

    $null = Run $tool @('merge', '-TaskFiles', $work, '-Out', $work)
    $merged = [System.IO.File]::ReadAllText((Join-Path $work 'merged.json'), [System.Text.Encoding]::UTF8) | ConvertFrom-Json
    Check (@($merged.texts).Count -eq 2) ('merge kept 2 text ops (got ' + @($merged.texts).Count + ')')
    Check (@($merged.conflicts).Count -ge 1) 'merge reported the slide-2 conflict'
    Check (@($merged.comments).Count -eq 1) 'merge kept the comment'

    $null = Run $tool @('plan', '-Merged', (Join-Path $work 'merged.json'), '-Out', $work, '-Doc', $doc, '-BatchOps')
    $batch = [System.IO.File]::ReadAllText((Join-Path $work 'batch-ops.json'), [System.Text.Encoding]::UTF8) | ConvertFrom-Json
    Check (@($batch).Count -eq 3) ('batch has 3 ops (got ' + @($batch).Count + ')')
    Check ([string]$batch[0].cmd -eq 'text' -and [string]$batch[1].cmd -eq 'replace' -and [string]$batch[2].cmd -eq 'comment') 'batch order text -> replace -> comment'

    Check ([bool](Run 'scripts\ppt.ps1' @('snapshot', '-Doc', $doc, '-Tag', 'pre')).ok) 'snapshot before batch'
    Check ([bool](Run 'scripts\ppt.ps1' @('batch', '-Ops', (Join-Path $work 'batch-ops.json'))).ok) 'batch applied through one process start'

    $read2 = Run 'scripts\ppt.ps1' @('read', '-Doc', $doc)
    if ($read2.ok) {
        $titles = @{}
        foreach ($sl in @($read2.result.items)) {
            foreach ($sh in @($sl.shapes)) { if ($sh.name -eq 'title') { $titles[[int]$sl.i] = [string]$sh.text } }
        }
        Write-Output ('  titles: 1=' + $titles[1] + ' 2=' + $titles[2] + ' 3=' + $titles[3])
        Check ($titles[2] -like 'T2-polished*') 'slide 2 title polished'
        Check ($titles[3] -like 'T3-polished*') 'slide 3 title polished'
        Check ($titles[1] -like 'T1-new*') 'slide 1 replace applied'
        Check ($titles[2] -notlike '*CONFLICT*') 'conflicting text not applied'
    }
    else { Check $false 'read-back after batch' }
    Check ([bool](Run 'scripts\ppt.ps1' @('diff', '-Doc', $doc, '-Since', 'pre')).ok) 'diff after batch'

    $null = Run 'scripts\ppt.ps1' @('close', '-Doc', $doc)
}
catch {
    [void]$fails.Add('exception: ' + $_.Exception.Message)
    Write-Output ('FAIL exception: ' + $_.Exception.Message)
}
finally {
    try { $null = Run 'scripts\ppt.ps1' @('quit') } catch { }
    try { Remove-Item -Recurse -Force -LiteralPath $work -ErrorAction SilentlyContinue } catch { }
}

if ($fails.Count -gt 0) { Write-Output ('PPT MULTI-TASK TEST FAIL (' + $fails.Count + ')'); exit 1 }
Write-Output 'PPT MULTI-TASK TEST PASS'
exit 0
