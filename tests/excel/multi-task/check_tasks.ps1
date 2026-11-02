# Excel multi-task helper regression (v1.22): split / merge / plan / batch over a generated workbook.
#
# Run:  powershell -NoProfile -ExecutionPolicy Bypass -File tests\excel\multi-task\check_tasks.ps1
# Needs WPS (it creates, writes, formats and comments a real workbook). The workbook lives in a
# temp folder and is closed without saving.
#
# What it proves:
#   1. split reads each range through excel.ps1 and writes a planner slice with the right shape;
#   2. merge refuses two different writes to the same cell (keeps the higher-priority task);
#   3. merge warns about writes without an expect anchor;
#   4. plan emits write -> format -> comment in that order, plus a batch-ops file with the doc;
#   5. one `batch` call applies all three phases; diff confirms exactly the planned cells changed.

$ErrorActionPreference = 'Stop'
$here = $PSScriptRoot
$repo = Split-Path -Parent (Split-Path -Parent (Split-Path -Parent $here))
$enc = New-Object System.Text.UTF8Encoding($false)
$work = Join-Path $env:TEMP ('woc-excel-mt-' + [guid]::NewGuid().ToString('N'))
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
$tool = 'scripts\tasks-excel.ps1'

try {
    $doc = Join-Path $work 'doc.xlsx'
    $newRes = Run 'scripts\excel.ps1' @('new', '-Doc', $doc, '-NoHint')
    Check ([bool]$newRes.ok) 'workbook created and opened'
    $seed = [ordered]@{
        items = @(
            [ordered]@{ sheet = 'Sheet1'; cell = 'A1'; value = 'time' },
            [ordered]@{ sheet = 'Sheet1'; cell = 'B1'; value = 'flow' },
            [ordered]@{ sheet = 'Sheet1'; cell = 'A2'; value = 't1' },
            [ordered]@{ sheet = 'Sheet1'; cell = 'B2'; value = 17.39 },
            [ordered]@{ sheet = 'Sheet1'; cell = 'A3'; value = 't2' },
            [ordered]@{ sheet = 'Sheet1'; cell = 'B3'; value = 17.44 }
        )
    }
    [System.IO.File]::WriteAllText((Join-Path $work 'seed.json'), (ConvertTo-Json -InputObject $seed -Depth 8 -Compress), $enc)
    Check ([bool](Run 'scripts\excel.ps1' @('write', '-Doc', $doc, '-Items', (Join-Path $work 'seed.json'))).ok) 'seed values written'

    [System.IO.File]::WriteAllText((Join-Path $work 'ranges.json'), (ConvertTo-Json -InputObject @(
                [ordered]@{ id = 'r1'; sheet = 'Sheet1'; range = 'A1:F20' },
                [ordered]@{ id = 'r2'; sheet = 'Sheet1'; range = 'A21:F60' }
            ) -Depth 6), $enc)
    $null = Run $tool @('split', '-Doc', $doc, '-Ranges', (Join-Path $work 'ranges.json'), '-Out', $work)
    $slice = [System.IO.File]::ReadAllText((Join-Path $work 'slice-r1.json'), [System.Text.Encoding]::UTF8) | ConvertFrom-Json
    Check ($slice.rows -eq 20 -and $slice.cols -eq 6) 'slice has the requested 20x6 shape'

    $t1 = [ordered]@{
        task = 'E1'; priority = 10
        writes = @(
            [ordered]@{ sheet = 'Sheet1'; cell = 'G2'; formula = '=ROUND(B2*2,2)'; expect = '' },
            [ordered]@{ sheet = 'Sheet1'; cell = 'G3'; formula = '=ROUND(B3*2,2)'; expect = '' },
            [ordered]@{ sheet = 'Sheet1'; cell = 'H2'; value = 'MT' }
        )
        formats = @([ordered]@{ sheet = 'Sheet1'; range = 'A1:F1'; bold = $true })
        comments = @([ordered]@{ sheet = 'Sheet1'; cell = 'A1'; text = 'multi-task header check' })
    }
    $t2 = [ordered]@{
        task = 'E2'; priority = 30
        writes = @(
            [ordered]@{ sheet = 'Sheet1'; cell = 'G2'; value = 'WRONG' },
            [ordered]@{ sheet = 'Sheet1'; cell = 'G5'; value = 'E2ONLY' }
        )
        formats = @([ordered]@{ sheet = 'Sheet1'; range = 'G1:G5'; numberFormat = '0.00' })
    }
    [System.IO.File]::WriteAllText((Join-Path $work 'task-E1.json'), (ConvertTo-Json -InputObject $t1 -Depth 10), $enc)
    [System.IO.File]::WriteAllText((Join-Path $work 'task-E2.json'), (ConvertTo-Json -InputObject $t2 -Depth 10), $enc)
    $null = Run $tool @('merge', '-TaskFiles', $work, '-Out', $work)
    $merged = [System.IO.File]::ReadAllText((Join-Path $work 'merged.json'), [System.Text.Encoding]::UTF8) | ConvertFrom-Json
    Check (@($merged.writes).Count -eq 4) ('merge kept the 4 non-conflicting writes (got ' + @($merged.writes).Count + ')')
    Check (@($merged.conflicts).Count -ge 1) 'merge reported the same-cell conflict'
    Check (@($merged.warnings).Count -ge 1) 'merge warned about missing expect anchors'

    $null = Run $tool @('plan', '-Merged', (Join-Path $work 'merged.json'), '-Out', $work, '-Doc', $doc, '-BatchOps')
    $batch = [System.IO.File]::ReadAllText((Join-Path $work 'batch-ops.json'), [System.Text.Encoding]::UTF8) | ConvertFrom-Json
    Check (@($batch).Count -eq 3) ('batch has 3 ops (got ' + @($batch).Count + ')')
    Check ([string]$batch[0].cmd -eq 'write' -and [string]$batch[1].cmd -eq 'format' -and [string]$batch[2].cmd -eq 'comment') 'batch order write -> format -> comment'

    Check ([bool](Run 'scripts\excel.ps1' @('snapshot', '-Doc', $doc)).ok) 'snapshot before applying'
    $apply = Run 'scripts\excel.ps1' @('batch', '-Ops', (Join-Path $work 'batch-ops.json'))
    Check ([bool]$apply.ok) 'batch applied through one process start'

    $read = Run 'scripts\excel.ps1' @('read', '-Doc', $doc, '-Range', 'Sheet1!G1:H6')
    if ($read.ok) {
        $v = $read.result.values
        Check ([string]$v[1][1] -eq 'MT') 'H2 written'
        Check ([string]$v[4][0] -eq 'E2ONLY') 'G5 written by the lower-priority task'
        Check ([string]$v[1][0] -ne 'WRONG') 'conflicting G2 write dropped'
    }
    else { Check $false 'read-back after batch' }

    $diff = Run 'scripts\excel.ps1' @('diff', '-Doc', $doc, '-Since', 'last')
    Check ([bool]$diff.ok -and (@($diff.result.changes).Count -ge 4)) 'diff confirms the planned cells changed'
    Check ($null -ne $diff.result.totalChanges -and [int]$diff.result.totalChanges -ge 4) 'diff totalChanges >= 4'

    $null = Run 'scripts\excel.ps1' @('close', '-Doc', $doc)
}
catch {
    [void]$fails.Add('exception: ' + $_.Exception.Message)
    Write-Output ('FAIL exception: ' + $_.Exception.Message)
}
finally {
    try { $null = Run 'scripts\excel.ps1' @('quit') } catch { }
    try { Remove-Item -Recurse -Force -LiteralPath $work -ErrorAction SilentlyContinue } catch { }
}

if ($fails.Count -gt 0) { Write-Output ('EXCEL MULTI-TASK TEST FAIL (' + $fails.Count + ')'); exit 1 }
Write-Output 'EXCEL MULTI-TASK TEST PASS'
exit 0
