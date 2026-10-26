# run_deliver.ps1 - acceptance driver for `ppt comment` / `export` / `picture`.
#
#   .\run_deliver.ps1
#   .\run_deliver.ps1 -SkipRun
#
# Two rules this file exists to respect:
#   * NO Chinese literal anywhere in it.  Windows PowerShell 5.1 parses a .ps1 as ANSI,
#     so a Chinese string here would reach the helper as mojibake.  Every Chinese
#     payload (review remarks, the replacement image name) is written to a UTF-8 file
#     by make_deliver_deck.py instead.
#   * every child process is started with Start-Process and its stdout/stderr are
#     redirected to files.  A native command writing to stderr turns into a terminating
#     error under $ErrorActionPreference='Stop' and would kill the run halfway.
#
[CmdletBinding()]
param(
    [string]$PythonExe = 'python',
    [string]$OutDir = '',
    # Left empty on purpose: this machine's FastDycrypt lives under "D:\DSH工作区\...", and
    # a Chinese literal inside this .ps1 would be read as ANSI and arrive as mojibake.
    # The path is read from local.config.json (UTF-8, gitignored) or passed in.
    [string]$DecryptEngine = '',
    [switch]$SkipRun
)

$ErrorActionPreference = 'Stop'
try { [Console]::OutputEncoding = New-Object System.Text.UTF8Encoding($false) } catch { }

$Root = $PSScriptRoot
$Repo = Split-Path -Parent (Split-Path -Parent $Root)
$Ppt = Join-Path $Repo 'scripts\ppt.ps1'
$Utf8 = New-Object System.Text.UTF8Encoding($false)
if (-not $OutDir) { $OutDir = Join-Path $Root 'out-deliver' }

if (-not $DecryptEngine) {
    $cfgPath = Join-Path $Root 'local.config.json'
    if (Test-Path -LiteralPath $cfgPath) {
        $cfg = [System.IO.File]::ReadAllText($cfgPath, [System.Text.Encoding]::UTF8) | ConvertFrom-Json
        if ($cfg.decryptEngine) { $DecryptEngine = [string]$cfg.decryptEngine }
    }
}
if ($DecryptEngine -and -not (Test-Path -LiteralPath $DecryptEngine)) {
    Write-Host ("    !! configured decrypt engine not found: {0}" -f $DecryptEngine) -ForegroundColor Yellow
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
    <# Runs one skill call; returns @{Code; Text}. #>
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
    $err = ''
    if (Test-Path -LiteralPath $errFile) { $err = [System.IO.File]::ReadAllText($errFile, [System.Text.Encoding]::UTF8) }
    [System.IO.File]::WriteAllText((Join-Path $Folder ('resp-' + $Step + '.json')), $text.Trim(), $Utf8)
    if ($err.Trim().Length -gt 0) {
        [System.IO.File]::WriteAllText((Join-Path $Folder ('stderr-' + $Step + '.txt')), $err, $Utf8)
    }
    $short = ($text -replace '\s+', ' ').Trim()
    if ($short.Length -gt 150) { $short = $short.Substring(0, 150) + '...' }
    Write-Host ("    {0,-14} [{1}] {2}" -f $Step, $proc.ExitCode, $short)
    return @{ Code = $proc.ExitCode; Text = $text }
}

function Decrypt-IfNeeded {
    param([string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) { return }
    $bytes = [System.IO.File]::ReadAllBytes($Path)
    if ($bytes.Length -le 3) { return }
    if (-not ($bytes[0] -eq 0x17 -and $bytes[1] -eq 0xDA -and $bytes[2] -eq 0x5F)) { return }
    if (-not (Test-Path -LiteralPath $DecryptEngine)) {
        Write-Host '    !! encrypted and FastDycrypt is missing' -ForegroundColor Yellow
        return
    }
    $proc = Start-Process -FilePath 'powershell.exe' -NoNewWindow -Wait -PassThru `
        -ArgumentList @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $DecryptEngine,
                        '-Paths', $Path, '-InPlace', '-DeleteBackupAfterSuccess', '-NoPause', '-NoLog') `
        -RedirectStandardOutput (Join-Path $OutDir '.dec.txt') -RedirectStandardError (Join-Path $OutDir '.dec.err.txt')
    $bytes = [System.IO.File]::ReadAllBytes($Path)
    $ok = ($bytes[0] -ne 0x17)
    Write-Host ("    decrypt {0}: {1}" -f (Split-Path -Leaf $Path), $(if ($ok) { 'ok' } else { 'FAILED' }))
}

function Stop-StrayWpp {
    <# `quit` only quits the instance the helper attached to; a second headless one
       survives and keeps the fixtures locked.  Only headless instances are killed - a
       deck the user has open has a window title and is left alone. #>
    for ($i = 0; $i -lt 10; $i++) {
        $stray = @(Get-Process wpp -ErrorAction SilentlyContinue | Where-Object { -not $_.MainWindowTitle })
        if ($stray.Count -eq 0) { return }
        if ($i -eq 0) { Write-Host ("    stopping {0} headless wpp instance(s)" -f $stray.Count) }
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
    $deck = Join-Path $OutDir 'deliver-deck.pptx'
    & $PythonExe (Join-Path $Root 'make_deliver_deck.py') --out $OutDir | Out-Null
    if ($LASTEXITCODE -ne 0) { throw 'make_deliver_deck.py failed' }

    Invoke-Skill -Folder $OutDir -Step 'open' -PptArgs @('open', '-Doc', $deck) | Out-Null

    # ---- comments
    $items = Join-Path $OutDir 'comments.json'
    $bad = Join-Path $OutDir 'comments-bad.json'
    Invoke-Skill -Folder $OutDir -Step 'comment-add' -PptArgs @('comment', '-Doc', $deck, '-Items', $items, '-Author', 'AI') | Out-Null
    Invoke-Skill -Folder $OutDir -Step 'comment-list' -PptArgs @('comment', '-Doc', $deck, '-Action', 'list') | Out-Null
    Invoke-Skill -Folder $OutDir -Step 'comment-bad' -PptArgs @('comment', '-Doc', $deck, '-Items', $bad, '-Author', 'AI') | Out-Null
    Invoke-Skill -Folder $OutDir -Step 'comment-list2' -PptArgs @('comment', '-Doc', $deck, '-Action', 'list') | Out-Null
    Invoke-Skill -Folder $OutDir -Step 'comment-del1' -PptArgs @('comment', '-Doc', $deck, '-Action', 'delete', '-Slide', '1', '-Index', '1') | Out-Null
    Invoke-Skill -Folder $OutDir -Step 'comment-delall' -PptArgs @('comment', '-Doc', $deck, '-Action', 'delete', '-Slide', '4', '-All') | Out-Null
    Invoke-Skill -Folder $OutDir -Step 'comment-list3' -PptArgs @('comment', '-Doc', $deck, '-Action', 'list') | Out-Null

    # ---- pictures
    $newPath = Join-Path (Join-Path $OutDir 'assets') 'new-image.png'
    Invoke-Skill -Folder $OutDir -Step 'picture-list' -PptArgs @('picture', '-Doc', $deck) | Out-Null
    Invoke-Skill -Folder $OutDir -Step 'picture-dry' -PptArgs @('picture', '-Doc', $deck, '-Action', 'replace', '-New', $newPath, '-DryRun') | Out-Null
    Invoke-Skill -Folder $OutDir -Step 'picture-replace' -PptArgs @('picture', '-Doc', $deck, '-Action', 'replace', '-New', $newPath) | Out-Null
    Invoke-Skill -Folder $OutDir -Step 'picture-list2' -PptArgs @('picture', '-Doc', $deck) | Out-Null

    # ---- export
    $pdf = Join-Path $OutDir 'out.pdf'
    Invoke-Skill -Folder $OutDir -Step 'export-pdf' -PptArgs @('export', '-Doc', $deck, '-Format', 'pdf', '-Out', $pdf) | Out-Null
    $pngDir = Join-Path $OutDir 'png'
    Invoke-Skill -Folder $OutDir -Step 'export-png' -PptArgs @('export', '-Doc', $deck, '-Format', 'png', '-Slides', '1-2', '-Out', $pngDir) | Out-Null
    Invoke-Skill -Folder $OutDir -Step 'export-md' -PptArgs @('export', '-Doc', $deck, '-Format', 'md', '-Out', (Join-Path $OutDir 'deck.md')) | Out-Null

    Invoke-Skill -Folder $OutDir -Step 'close' -PptArgs @('close', '-Doc', $deck, '-Save') | Out-Null

    # readable copies for the python assertions
    Copy-Item -LiteralPath $pdf -Destination (Join-Path $OutDir 'out.plain.pdf') -Force
    Copy-Item -LiteralPath $deck -Destination (Join-Path $OutDir 'deck.plain.pptx') -Force
    Decrypt-IfNeeded -Path (Join-Path $OutDir 'out.plain.pdf')
    Decrypt-IfNeeded -Path (Join-Path $OutDir 'deck.plain.pptx')
    Invoke-Skill -Folder $OutDir -Step 'quit2' -PptArgs @('quit') | Out-Null
}

Write-Host ''
Write-Host '== checks'
& $PythonExe (Join-Path $Root 'check_deliver.py') --dir $OutDir
$code = $LASTEXITCODE
if ($code -eq 0) { Write-Host 'RESULT: PASS' -ForegroundColor Green } else { Write-Host 'RESULT: FAIL' -ForegroundColor Red }
exit $code
