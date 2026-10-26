# run_master.ps1 - acceptance driver for `ppt master` (list / footer / layout / rebase).
#
#   .\run_master.ps1
#   .\run_master.ps1 -SkipRun
#
# Pure ASCII on purpose (Windows PowerShell 5.1 parses .ps1 as ANSI); Chinese payloads
# live in UTF-8 files written by python.
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
if (-not $OutDir) { $OutDir = Join-Path $Root 'out-master' }

if (-not $DecryptEngine) {
    $cfgPath = Join-Path $Root 'local.config.json'
    if (Test-Path -LiteralPath $cfgPath) {
        $cfg = [System.IO.File]::ReadAllText($cfgPath, [System.Text.Encoding]::UTF8) | ConvertFrom-Json
        if ($cfg.decryptEngine) { $DecryptEngine = [string]$cfg.decryptEngine }
    }
}

function Quote-Arg {
    <# Start-Process joins -ArgumentList with spaces and quotes nothing, so an argument
       with a space in it (a footer text, a path under "Program Files") arrives as two
       arguments and the skill refuses to bind. #>
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
    & $PythonExe (Join-Path $Root 'make_master_deck.py') --out $OutDir | Out-Null
    if ($LASTEXITCODE -ne 0) { throw 'make_master_deck.py failed' }

    $src = Join-Path $OutDir 'source.pptx'
    $tpl = Join-Path $OutDir 'template.pptx'
    Invoke-Skill -Folder $OutDir -Step 'open' -PptArgs @('open', '-Doc', $src) | Out-Null

    # ---- inventory
    Invoke-Skill -Folder $OutDir -Step 'master-list' -PptArgs @('master', '-Doc', $src) | Out-Null

    # ---- batch footer + page numbers, cover skipped
    Invoke-Skill -Folder $OutDir -Step 'master-footer' -PptArgs @('master', '-Doc', $src, '-Action', 'footer', '-Text', 'ACME CORP', '-PageNum', '-SkipCover') | Out-Null
    Invoke-Skill -Folder $OutDir -Step 'master-list2' -PptArgs @('master', '-Doc', $src) | Out-Null
    # and a full clear
    Invoke-Skill -Folder $OutDir -Step 'master-unfooter' -PptArgs @('master', '-Doc', $src, '-Action', 'footer', '-Clear') | Out-Null
    Invoke-Skill -Folder $OutDir -Step 'master-list3' -PptArgs @('master', '-Doc', $src) | Out-Null
    # restore them, since the saved deck is what python checks
    Invoke-Skill -Folder $OutDir -Step 'master-footer2' -PptArgs @('master', '-Doc', $src, '-Action', 'footer', '-Text', 'ACME CORP', '-PageNum', '-SkipCover') | Out-Null
    Invoke-Skill -Folder $OutDir -Step 'save' -PptArgs @('save', '-Doc', $src) | Out-Null

    # ---- layout switch
    Invoke-Skill -Folder $OutDir -Step 'layout-dry' -PptArgs @('master', '-Doc', $src, '-Action', 'layout', '-Layout', '1', '-Slides', '2', '-DryRun') | Out-Null
    Invoke-Skill -Folder $OutDir -Step 'layout' -PptArgs @('master', '-Doc', $src, '-Action', 'layout', '-Layout', '1', '-Slides', '2') | Out-Null
    Invoke-Skill -Folder $OutDir -Step 'layout-bad' -PptArgs @('master', '-Doc', $src, '-Action', 'layout', '-Layout', '42') | Out-Null
    Invoke-Skill -Folder $OutDir -Step 'save2' -PptArgs @('save', '-Doc', $src) | Out-Null

    # ---- rebase onto the template
    $outDeck = Join-Path $OutDir 'rebased.pptx'
    Invoke-Skill -Folder $OutDir -Step 'rebase' -PptArgs @('master', '-Doc', $src, '-Action', 'rebase', '-Path', $tpl, '-Out', $outDeck) | Out-Null

    Invoke-Skill -Folder $OutDir -Step 'close' -PptArgs @('close', '-Doc', $src, '-Save') | Out-Null
    Invoke-Skill -Folder $OutDir -Step 'quit2' -PptArgs @('quit') | Out-Null
    Stop-StrayWpp

    # readable copies
    Copy-Item -LiteralPath $src -Destination (Join-Path $OutDir 'source.plain.pptx') -Force
    Decrypt-IfNeeded -Path (Join-Path $OutDir 'source.plain.pptx')
    if (Test-Path -LiteralPath $outDeck) {
        Copy-Item -LiteralPath $outDeck -Destination (Join-Path $OutDir 'rebased.plain.pptx') -Force
        Decrypt-IfNeeded -Path (Join-Path $OutDir 'rebased.plain.pptx')
    }
}

Write-Host ''
Write-Host '== checks'
& $PythonExe (Join-Path $Root 'check_master.py') --dir $OutDir
$code = $LASTEXITCODE
if ($code -eq 0) { Write-Host 'RESULT: PASS' -ForegroundColor Green } else { Write-Host 'RESULT: FAIL' -ForegroundColor Red }
exit $code
