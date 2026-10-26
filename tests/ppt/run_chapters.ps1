# run_chapters.ps1 - acceptance driver for `ppt split` / `ppt merge`.
#
#   .\run_chapters.ps1
#   .\run_chapters.ps1 -SkipRun
#
# Pure ASCII on purpose (Windows PowerShell 5.1 parses .ps1 as ANSI); every Chinese
# payload lives in a UTF-8 file written by python.
#
[CmdletBinding()]
param(
    [string]$PythonExe = 'python',
    [string]$OutDir = '',
    [string]$DecryptEngine = '',
    [switch]$SkipRun
)

$ErrorActionPreference = 'Stop'
try { [Console]::OutputEncoding = New-Object System.Text.UTF8Encoding($false) } catch { }

$Root = $PSScriptRoot
$Repo = Split-Path -Parent (Split-Path -Parent $Root)
$Ppt = Join-Path $Repo 'scripts\ppt.ps1'
$Utf8 = New-Object System.Text.UTF8Encoding($false)
if (-not $OutDir) { $OutDir = Join-Path $Root 'out-chapters' }

if (-not $DecryptEngine) {
    $cfgPath = Join-Path $Root 'local.config.json'
    if (Test-Path -LiteralPath $cfgPath) {
        $cfg = [System.IO.File]::ReadAllText($cfgPath, [System.Text.Encoding]::UTF8) | ConvertFrom-Json
        if ($cfg.decryptEngine) { $DecryptEngine = [string]$cfg.decryptEngine }
    }
}

function Quote-Arg {
    <# Start-Process joins -ArgumentList with spaces and quotes nothing, so an argument
       with a space in it arrives as two arguments and the skill refuses to bind. #>
    param([string]$Value)
    if ($null -eq $Value) { return '""' }
    if ($Value -notmatch '[\s"]') { return $Value }
    return '"' + ($Value -replace '"', '\"') + '"'
}

function Invoke-Skill {
    param([string]$Folder, [string]$Step, [string[]]$PptArgs)
    if (-not (Test-Path -LiteralPath $Folder)) { New-Item -ItemType Directory -Force -Path $Folder | Out-Null }
    $outFile = Join-Path $Folder ('.out-' + $Step + '.txt')
    $errFile = Join-Path $Folder ('.err-' + $Step + '.txt')
    $argList = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $Ppt) + $PptArgs
    $argList = @($argList | ForEach-Object { Quote-Arg $_ })
    $proc = Start-Process -FilePath 'powershell.exe' -ArgumentList $argList -NoNewWindow -Wait -PassThru `
        -RedirectStandardOutput $outFile -RedirectStandardError $errFile
    $text = ''
    if (Test-Path -LiteralPath $outFile) { $text = [System.IO.File]::ReadAllText($outFile, [System.Text.Encoding]::UTF8) }
    [System.IO.File]::WriteAllText((Join-Path $Folder ('resp-' + $Step + '.json')), $text.Trim(), $Utf8)
    $short = ($text -replace '\s+', ' ').Trim()
    if ($short.Length -gt 150) { $short = $short.Substring(0, 150) + '...' }
    Write-Host ("    {0,-14} [{1}] {2}" -f $Step, $proc.ExitCode, $short)
}

function Decrypt-IfNeeded {
    param([string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) { return }
    $bytes = [System.IO.File]::ReadAllBytes($Path)
    if ($bytes.Length -le 3) { return }
    if (-not ($bytes[0] -eq 0x17 -and $bytes[1] -eq 0xDA -and $bytes[2] -eq 0x5F)) { return }
    if (-not $DecryptEngine -or -not (Test-Path -LiteralPath $DecryptEngine)) {
        Write-Host '    !! encrypted and no decrypt engine configured' -ForegroundColor Yellow
        return
    }
    Start-Process -FilePath 'powershell.exe' -NoNewWindow -Wait `
        -ArgumentList @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $DecryptEngine,
                        '-Paths', $Path, '-InPlace', '-DeleteBackupAfterSuccess', '-NoPause', '-NoLog') `
        -RedirectStandardOutput (Join-Path $OutDir '.dec.txt') -RedirectStandardError (Join-Path $OutDir '.dec.err.txt') | Out-Null
    $bytes = [System.IO.File]::ReadAllBytes($Path)
    Write-Host ("    decrypt {0}: {1}" -f (Split-Path -Leaf $Path), $(if ($bytes[0] -eq 0x17) { 'FAILED' } else { 'ok' }))
}

function Stop-StrayWpp {
    for ($i = 0; $i -lt 10; $i++) {
        $stray = @(Get-Process wpp -ErrorAction SilentlyContinue | Where-Object { -not $_.MainWindowTitle })
        if ($stray.Count -eq 0) { return }
        $stray | ForEach-Object { try { Stop-Process -Id $_.Id -Force } catch { } }
        Start-Sleep -Milliseconds 400
    }
}

if (-not $SkipRun) {
    Invoke-Skill -Folder $OutDir -Step 'quit' -PptArgs @('quit') | Out-Null
    Start-Sleep -Milliseconds 700
    Stop-StrayWpp
    if (Test-Path -LiteralPath $OutDir) { try { Remove-Item -Recurse -Force -LiteralPath $OutDir } catch { } }
    New-Item -ItemType Directory -Force -Path $OutDir | Out-Null
    & $PythonExe (Join-Path $Root 'make_chapters_deck.py') --out $OutDir | Out-Null
    if ($LASTEXITCODE -ne 0) { throw 'make_chapters_deck.py failed' }

    $deck = Join-Path $OutDir 'split-deck.pptx'
    $parts = Join-Path $OutDir 'parts'
    Invoke-Skill -Folder $OutDir -Step 'open' -PptArgs @('open', '-Doc', $deck) | Out-Null

    # ---- split: 12 pages into 3 decks of 4
    Invoke-Skill -Folder $OutDir -Step 'split-dry' -PptArgs @('split', '-Doc', $deck, '-Chunks', '1-4,5-8,9-12', '-Out', $parts, '-Prefix', 'book', '-DryRun') | Out-Null
    Invoke-Skill -Folder $OutDir -Step 'split' -PptArgs @('split', '-Doc', $deck, '-Chunks', '1-4,5-8,9-12', '-Out', $parts, '-Prefix', 'book') | Out-Null
    # a gap must be refused rather than silently dropping pages
    Invoke-Skill -Folder $OutDir -Step 'split-gap' -PptArgs @('split', '-Doc', $deck, '-Chunks', '1-4,9-12', '-Out', $parts, '-Prefix', 'bad') | Out-Null
    Invoke-Skill -Folder $OutDir -Step 'source-after-split' -PptArgs @('read', '-Doc', $deck, '-Count', '12') | Out-Null

    # ---- merge: three county decks into one
    $host2 = Join-Path $OutDir 'host.pptx'
    Invoke-Skill -Folder $OutDir -Step 'host-open' -PptArgs @('open', '-Doc', $host2) | Out-Null
    Invoke-Skill -Folder $OutDir -Step 'host-theme' -PptArgs @('slide', '-Doc', $host2, '-Action', 'add', '-Preset', 'title', '-Title', 'Merged') | Out-Null
    $sourcesFile = Join-Path $OutDir 'sources.json'
    $a = Join-Path $OutDir 'county-a.pptx'
    $b = Join-Path $OutDir 'county-b.pptx'
    $c = Join-Path $OutDir 'county-c.pptx'
    $items = @{ sources = @($a, $b, $c) } | ConvertTo-Json -Depth 3
    [System.IO.File]::WriteAllText($sourcesFile, $items, $Utf8)
    Invoke-Skill -Folder $OutDir -Step 'merge' -PptArgs @('merge', '-Doc', $host2, '-Items', $sourcesFile) | Out-Null
    Invoke-Skill -Folder $OutDir -Step 'merge-read' -PptArgs @('read', '-Doc', $host2, '-Count', '12') | Out-Null
    Invoke-Skill -Folder $OutDir -Step 'merge-save' -PptArgs @('save', '-Doc', $host2) | Out-Null
    Invoke-Skill -Folder $OutDir -Step 'host-close' -PptArgs @('close', '-Doc', $host2, '-Save') | Out-Null

    # merge into a chosen position: append one deck after slide 1 of the split deck
    $deck2 = Join-Path $OutDir 'split-deck.pptx'
    Invoke-Skill -Folder $OutDir -Step 'merge-after' -PptArgs @('merge', '-Doc', $deck2, '-Sources', $a, '-After', '1', '-DryRun') | Out-Null

    Invoke-Skill -Folder $OutDir -Step 'quit2' -PptArgs @('quit') | Out-Null
    Stop-StrayWpp

    # readable copies for python (parts are written by WPS, hence encrypted here).
    # Materialise the list first: copying into the same folder while enumerating it makes
    # the enumerator pick up the copies it just created.
    $partFiles = @(Get-ChildItem -LiteralPath $parts -Filter '*.pptx' -ErrorAction SilentlyContinue |
                   Where-Object { -not $_.Name.StartsWith('plain-') })
    foreach ($file in $partFiles) {
        $plain = Join-Path $parts ('plain-' + $file.Name)
        Copy-Item -LiteralPath $file.FullName -Destination $plain -Force
        Decrypt-IfNeeded -Path $plain
    }
    Copy-Item -LiteralPath $host2 -Destination (Join-Path $OutDir 'host.plain.pptx') -Force
    Decrypt-IfNeeded -Path (Join-Path $OutDir 'host.plain.pptx')
}

Write-Host ''
Write-Host '== checks'
& $PythonExe (Join-Path $Root 'check_chapters.py') --dir $OutDir
$code = $LASTEXITCODE
if ($code -eq 0) { Write-Host 'RESULT: PASS' -ForegroundColor Green } else { Write-Host 'RESULT: FAIL' -ForegroundColor Red }
exit $code
