# Shared plumbing for the three wrappers (word / excel / ppt), v1.19.
#
# Two things every wrapper needs and none of them should re-implement:
#
# 1. Get-CachedSourceHash - identify the compiled helper by a hash of its sources. Hashing the
#    .cs files with Get-FileHash costs ~50 ms on EVERY command (measured 2026-11-02 on a 9-file
#    helper), which is pure tax when the answer only changes after an edit. The cache key is
#    "file count + sizes + newest mtime": any normal edit changes at least one of those.
#
# 2. Invoke-WocBatch - run a list of requests over ONE named-pipe connection. The per-command
#    overhead of the CLI path is dominated by spawning PowerShell + parsing the wrapper (~220 ms),
#    while the pipe round trip is 1-15 ms, so a batch of N commands from one tool call is roughly
#    N*COM + one process start instead of N*(process start + parse). Measured on another machine
#    2026-11-02: 405-460 ms per command via the CLI vs 0.7-15 ms per request over a warm pipe.

function Get-CachedSourceHash {
    param(
        [Parameter(Mandatory = $true)][string]$App,
        [Parameter(Mandatory = $true)]$Sources
    )
    $files = @($Sources)
    if ($files.Count -eq 0) { throw "no source files for $App" }
    $sizes = @($files | ForEach-Object { $_.Length })
    $newest = ($files | Measure-Object -Property LastWriteTimeUtc -Maximum).Maximum.Ticks
    $key = "$($files.Count)|$($newest)|" + (($sizes | Sort-Object) -join ',')
    $cacheDir = Join-Path $env:LOCALAPPDATA 'wps-office-cooperate'
    $cacheFile = Join-Path $cacheDir ("source-hash-" + $App + ".json")
    try {
        if (Test-Path -LiteralPath $cacheFile) {
            $cached = Get-Content -LiteralPath $cacheFile -Raw -Encoding UTF8 | ConvertFrom-Json
            if ($cached.key -eq $key -and $cached.sha) { return [string]$cached.sha }
        }
    } catch { }
    $hashInput = ($files | ForEach-Object { $_.Name + ':' + (Get-FileHash -LiteralPath $_.FullName -Algorithm SHA256).Hash }) -join ';'
    $shaGen = [System.Security.Cryptography.SHA256]::Create()
    $sha = ([System.BitConverter]::ToString($shaGen.ComputeHash([System.Text.Encoding]::UTF8.GetBytes($hashInput))) -replace '-', '').ToLowerInvariant().Substring(0, 12)
    try {
        if (-not (Test-Path -LiteralPath $cacheDir)) { New-Item -ItemType Directory -Force -Path $cacheDir | Out-Null }
        @{ key = $key; sha = $sha } | ConvertTo-Json -Compress | Set-Content -LiteralPath $cacheFile -Encoding UTF8
    } catch { }
    return $sha
}

function Invoke-WocPipeOnce {
    param([string]$PipeName, [string]$Json, [int]$ConnectMs = 1500)
    $pipe = $null
    try {
        $pipe = New-Object System.IO.Pipes.NamedPipeClientStream('.', $PipeName, [System.IO.Pipes.PipeDirection]::InOut)
        $pipe.Connect($ConnectMs)
        return Send-WocPipeRequest -Pipe $pipe -Json $Json
    } catch {
        return $null
    } finally {
        if ($pipe) { try { $pipe.Dispose() } catch { } }
    }
}

function Send-WocPipeRequest {
    param($Pipe, [string]$Json)
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($Json)
    $Pipe.Write([System.BitConverter]::GetBytes([int]$bytes.Length), 0, 4)
    $Pipe.Write($bytes, 0, $bytes.Length)
    $Pipe.Flush()
    $hdr = New-Object byte[] 4
    $got = 0
    while ($got -lt 4) {
        $n = $Pipe.Read($hdr, $got, 4 - $got)
        if ($n -le 0) { return $null }
        $got += $n
    }
    $total = [System.BitConverter]::ToInt32($hdr, 0)
    if ($total -le 0 -or $total -gt 33554432) { return $null }
    $buf = New-Object byte[] $total
    $got = 0
    while ($got -lt $total) {
        $n = $Pipe.Read($buf, $got, $total - $got)
        if ($n -le 0) { return $null }
        $got += $n
    }
    return [System.Text.Encoding]::UTF8.GetString($buf)
}

# Run a list of request objects (already parsed from JSON) over one pipe connection.
# Each op may carry "doc"; relative paths are resolved to absolute here because the helper runs
# with its own working directory.
function Invoke-WocBatch {
    param(
        [string]$PipeName,
        [string]$ExePath,
        $Ops,
        [int]$ConnectMs = 1500
    )
    $opsArray = @($Ops)
    $started = $null
    if ($null -eq (Invoke-WocPipeOnce -PipeName $PipeName -Json '{"cmd":"ping"}' -ConnectMs 300)) {
        try {
            $started = Start-Process -FilePath $ExePath -ArgumentList @('--serve', $PipeName) -WindowStyle Hidden -PassThru
        } catch { }
        for ($i = 0; $i -lt 50; $i++) {
            Start-Sleep -Milliseconds 100
            if ($null -ne (Invoke-WocPipeOnce -PipeName $PipeName -Json '{"cmd":"ping"}' -ConnectMs 300)) { break }
        }
    }
    $results = New-Object System.Collections.ArrayList
    $idx = 0
    foreach ($op in $opsArray) {
        $idx++
        $req = @{}
        foreach ($prop in $op.PSObject.Properties) {
            $val = $prop.Value
            if ($prop.Name -eq 'doc' -and $val) {
                try { $val = [System.IO.Path]::GetFullPath([string]$val) } catch { }
            }
            $req[$prop.Name] = $val
        }
        $json = $req | ConvertTo-Json -Depth 12 -Compress
        $sw = [System.Diagnostics.Stopwatch]::StartNew()
        # One request per connection: PipeServer creates a fresh NamedPipeServerStream per request,
        # so a batch reconnects per op. A local connect is 1-3 ms - still two orders of magnitude
        # below the ~400 ms a separate CLI command costs.
        $resp = Invoke-WocPipeOnce -PipeName $PipeName -Json $json -ConnectMs $ConnectMs
        $sw.Stop()
        if ($null -eq $resp) {
            [void]$results.Add([ordered]@{ i = $idx; ms = $sw.ElapsedMilliseconds; ok = $false; error = @{ code = 'NO_DAEMON'; message = 'the resident helper did not answer' } })
            continue
        }
        $obj = $resp | ConvertFrom-Json
        $entry = [ordered]@{ i = $idx; ms = $sw.ElapsedMilliseconds; ok = [bool]$obj.ok }
        if ($obj.ok) { $entry['result'] = $obj.result } else { $entry['error'] = $obj.error }
        [void]$results.Add($entry)
    }
    $totalMs = 0
    foreach ($r in $results) { $totalMs += [int]$r.ms }
    return [ordered]@{
        ok = $true
        cmd = 'batch'
        result = [ordered]@{ count = $results.Count; totalMs = $totalMs; ops = @($results) }
    }
}
