# wps-office-cooperate cross-app session scheduler (v1.24).
#
# Pure ASCII source.
#
#   tasks-session.ps1 plan  -Board board.json -Out schedule.json
#   tasks-session.ps1 check -Board board.json
#   tasks-session.ps1 selftest
#
# board.json = {"tasks":[
#   {"id":"T1","app":"word","doc":"C:\\...\\report.docx","kind":"write","priority":1,"dependsOn":[]},
#   {"id":"T2","app":"excel","doc":"C:\\...\\book.xlsx","kind":"write"},
#   {"id":"T3","app":"ppt","doc":"C:\\...\\deck.pptx","kind":"write","dependsOn":["T1"]}
# ]}
#
# Why: a session can run several tasks at once, but the physical rules are:
#   - one writer per document (two writers on the same doc corrupt each other: word para drift,
#     excel overlapping ranges, ppt slide-index drift);
#   - different documents may overlap, but WPS COM is a global-serialised channel: light commands
#     really overlap, heavy ones queue - so cap the wave at 2-3 tasks and expect no linear speedup;
#   - cross-app renders share the OS clipboard (the helpers take a mutex, so they serialise anyway).
# The scheduler therefore computes "waves": every task in a wave touches a different document, and
# a task only starts after its dependsOn tasks and after the previous task on the same document.

param(
    [Parameter(Mandatory = $true, Position = 0)]
    [ValidateSet('plan', 'check', 'selftest', 'help')]
    [string]$Command,

    [string]$Board,
    [string]$Out,
    [int]$MaxParallel = 3
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
function FullPath2([string]$p) { try { return [System.IO.Path]::GetFullPath($p) } catch { return $p } }

function Get-Tasks($board) {
    $list = @()
    if ($board -is [System.Array]) { $list = @($board) }
    elseif ($board.PSObject.Properties.Name -contains 'tasks') { $list = @($board.tasks) }
    else { throw 'board must be an array or {"tasks":[...]}' }
    return $list
}

# Validate the board and return a normalised task table (throws when the board is unusable).
function Build-Table($tasks) {
    $byId = @{}
    $order = 0
    foreach ($t in $tasks) {
        $order++
        if ($null -eq $t) { continue }
        $id = [string]$t.id
        if (-not $id) { throw 'every task needs an id' }
        if ($byId.ContainsKey($id)) { throw ('duplicate task id: ' + $id) }
        $app = ([string]$t.app).ToLowerInvariant()
        if ($app -notin @('word', 'excel', 'ppt')) { throw ('task ' + $id + ': app must be word|excel|ppt (got "' + $t.app + '")') }
        $doc = [string]$t.doc
        if (-not $doc) { throw ('task ' + $id + ': doc is required') }
        $kind = 'write'
        if ($t.PSObject.Properties.Name -contains 'kind' -and $t.kind) { $kind = ([string]$t.kind).ToLowerInvariant() }
        if ($kind -notin @('write', 'read')) { throw ('task ' + $id + ': kind must be write|read') }
        $prio = 50
        try { if ($t.priority) { $prio = [int]$t.priority } } catch { }
        $deps = @()
        if ($t.PSObject.Properties.Name -contains 'dependsOn') { $deps = @($t.dependsOn | Where-Object { $_ } | ForEach-Object { [string]$_ }) }
        $byId[$id] = [pscustomobject]@{
            id = $id; app = $app; doc = (FullPath2 $doc); kind = $kind; priority = $prio
            deps = $deps; order = $order; level = 1
        }
    }
    foreach ($id in $byId.Keys) {
        foreach ($d in $byId[$id].deps) {
            if (-not $byId.ContainsKey($d)) { throw ('task ' + $id + ' depends on unknown task: ' + $d) }
        }
    }
    # implicit chain: tasks on the same document are serialised in board order (stable)
    $lastByDoc = @{}
    foreach ($t in ($byId.Values | Sort-Object -Property order)) {
        if ($lastByDoc.ContainsKey($t.doc)) {
            if ($t.deps -notcontains $lastByDoc[$t.doc]) { $t.deps = @($t.deps) + $lastByDoc[$t.doc] }
        }
        $lastByDoc[$t.doc] = $t.id
    }
    # cycle detection + level computation (Kahn)
    $indeg = @{}
    $children = @{}
    foreach ($t in $byId.Values) { $indeg[$t.id] = $t.deps.Count; $children[$t.id] = @() }
    foreach ($t in $byId.Values) {
        foreach ($d in $t.deps) {
            if (-not $children.ContainsKey($d)) { $children[$d] = @() }
            $children[$d] = @($children[$d]) + $t.id
        }
    }
    $queue = New-Object System.Collections.Queue
    foreach ($t in $byId.Values) { if ($indeg[$t.id] -eq 0) { $queue.Enqueue($t.id) } }
    $done = 0
    while ($queue.Count -gt 0) {
        $id = $queue.Dequeue(); $done++
        $t = $byId[$id]
        foreach ($c in @($children[$id])) {
            $child = $byId[$c]
            $lvl = [int]$t.level + 1
            if ($lvl -gt $child.level) { $child.level = $lvl }
            $indeg[$c] = [int]$indeg[$c] - 1
            if ($indeg[$c] -eq 0) { $queue.Enqueue($c) }
        }
    }
    if ($done -ne $byId.Count) { throw 'dependency cycle in board.json' }
    return @($byId.Values | Sort-Object -Property level, order)
}

function Invoke-Plan {
    if (-not $Board) { throw 'plan needs -Board <board.json>' }
    if (-not $Out) { throw 'plan needs -Out <schedule.json>' }
    $table = Build-Table (Get-Tasks (Read-Json $Board))
    $maxLevel = 0
    foreach ($t in $table) { if ($t.level -gt $maxLevel) { $maxLevel = $t.level } }

    $waves = New-Object System.Collections.ArrayList
    $warnings = New-Object System.Collections.ArrayList
    for ($lv = 1; $lv -le $maxLevel; $lv++) {
        $inWave = @($table | Where-Object { $_.level -eq $lv })
        $docs = @($inWave | ForEach-Object { $_.doc } | Select-Object -Unique)
        if ($inWave.Count -gt $MaxParallel) {
            [void]$warnings.Add('wave ' + $lv + ' has ' + $inWave.Count + ' tasks (cap is ' + $MaxParallel + '): run them in sub-batches')
        }
        if ($docs.Count -ne $inWave.Count) { [void]$warnings.Add('wave ' + $lv + ' has two tasks on one document; this should not happen') }
        $apps = @($inWave | ForEach-Object { $_.app } | Select-Object -Unique)
        $parallel = ($inWave.Count -gt 1 -and $docs.Count -eq $inWave.Count)
        $reason = if ($parallel) { 'different documents: may overlap (heavy COM serialises; keep the cap)' } else { 'serial: one task or same-document constraint' }
        [void]$waves.Add([pscustomobject]@{
                level = $lv; parallel = $parallel; reason = $reason
                apps = $apps
                tasks = @($inWave | ForEach-Object { [pscustomobject]@{ id = $_.id; app = $_.app; doc = $_.doc; kind = $_.kind; priority = $_.priority } })
            })
    }
    $perDoc = @{}
    foreach ($t in ($table | Sort-Object -Property order)) {
        if (-not $perDoc.ContainsKey($t.doc)) { $perDoc[$t.doc] = New-Object System.Collections.ArrayList }
        [void]$perDoc[$t.doc].Add($t.id)
    }
    $docOrder = @()
    foreach ($k in $perDoc.Keys) { $docOrder += [pscustomobject]@{ doc = $k; order = @($perDoc[$k]) } }
    $schedule = [pscustomobject]@{
        taskCount = $table.Count
        waves     = @($waves)
        perDoc    = @($docOrder)
        warnings  = @($warnings)
        rules     = @(
            'one writer per document',
            'a later wave starts only after the previous wave finished (or after its dependsOn task)',
            'cross-app parallel is allowed but WPS COM is globally serialised: cap at 2-3',
            'renders share the OS clipboard (helpers hold a mutex)'
        )
    }
    Write-Json $Out $schedule
    Write-Output ('schedule: ' + $table.Count + ' task(s) in ' + $waves.Count + ' wave(s) -> ' + $Out)
    foreach ($w in $waves) {
        Write-Output ('  wave ' + $w.level + ($(if ($w.parallel) { ' (parallel)' } else { ' (serial)' })) + ': ' + (($w.tasks | ForEach-Object { $_.id + ':' + $_.app }) -join ', '))
    }
    foreach ($wn in $warnings) { Write-Output ('  WARN ' + $wn) }
}

function Invoke-Check {
    if (-not $Board) { throw 'check needs -Board <board.json>' }
    $table = Build-Table (Get-Tasks (Read-Json $Board))
    Write-Output ('OK: ' + $table.Count + ' task(s), ' + @($table | ForEach-Object { $_.doc } | Select-Object -Unique).Count + ' document(s)')
    foreach ($t in $table) { Write-Output ('  ' + $t.id + ' ' + $t.app + ' ' + $t.kind + ' level=' + $t.level + ' doc=' + $t.doc) }
}

function Invoke-Selftest {
    $tasks = @(
        [pscustomobject]@{ id = 'T1'; app = 'word'; doc = 'D:\a.docx'; kind = 'write'; priority = 10 },
        [pscustomobject]@{ id = 'T2'; app = 'excel'; doc = 'D:\b.xlsx'; kind = 'write'; priority = 10 },
        [pscustomobject]@{ id = 'T3'; app = 'word'; doc = 'D:\a.docx'; kind = 'write'; priority = 20 },
        [pscustomobject]@{ id = 'T4'; app = 'ppt'; doc = 'D:\c.pptx'; kind = 'write'; dependsOn = @('T1') }
    )
    $table = Build-Table $tasks
    $byId = @{}
    foreach ($t in $table) { $byId[$t.id] = $t }
    $fails = New-Object System.Collections.ArrayList
    if ($byId['T1'].level -ne 1) { [void]$fails.Add('T1 should be level 1') }
    if ($byId['T2'].level -ne 1) { [void]$fails.Add('T2 should be level 1') }
    if ($byId['T3'].level -le $byId['T1'].level) { [void]$fails.Add('T3 (same doc) must be after T1') }
    if ($byId['T4'].level -le $byId['T1'].level) { [void]$fails.Add('T4 (depends on T1) must be after T1') }
    try { $null = Build-Table @([pscustomobject]@{ id = 'A'; app = 'word'; doc = 'D:\x'; dependsOn = @('B') }, [pscustomobject]@{ id = 'B'; app = 'word'; doc = 'D:\y'; dependsOn = @('A') }); [void]$fails.Add('cycle not detected') } catch { }
    if ($fails.Count -gt 0) { Write-Output ('SELFTEST FAIL: ' + ($fails -join '; ')); exit 1 }
    Write-Output 'SELFTEST PASS'
}

switch ($Command) {
    'plan' { Invoke-Plan }
    'check' { Invoke-Check }
    'selftest' { Invoke-Selftest }
    default {
        Write-Output 'usage: tasks-session.ps1 plan|check|selftest [options]'
        Write-Output '  plan  -Board board.json -Out schedule.json [-MaxParallel 3]'
        Write-Output '  check -Board board.json'
    }
}
