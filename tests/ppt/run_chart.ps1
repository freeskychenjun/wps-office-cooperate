# run_chart.ps1 - acceptance driver for `ppt chart`.
#
#   .\run_chart.ps1
#   .\run_chart.ps1 -SkipRun
#
# The deck has to be opened with -Show: WPS refuses Shapes.AddChart in a presentation
# without a window (E_FAIL), and that is also the only reason this driver passes -Show.
#
# Pure ASCII on purpose (Windows PowerShell 5.1 parses .ps1 as ANSI).
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
if (-not $OutDir) { $OutDir = Join-Path $Root 'out-chart' }

if (-not $DecryptEngine) {
    $cfgPath = Join-Path $Root 'local.config.json'
    if (Test-Path -LiteralPath $cfgPath) {
        $cfg = [System.IO.File]::ReadAllText($cfgPath, [System.Text.Encoding]::UTF8) | ConvertFrom-Json
        if ($cfg.decryptEngine) { $DecryptEngine = [string]$cfg.decryptEngine }
    }
}

function Quote-Arg {
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
    if ($short.Length -gt 160) { $short = $short.Substring(0, 160) + '...' }
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
    & $PythonExe (Join-Path $Root 'make_chart_deck.py') --out $OutDir | Out-Null
    if ($LASTEXITCODE -ne 0) { throw 'make_chart_deck.py failed' }

    $deck = Join-Path $OutDir 'chart-deck.pptx'
    # -Show is mandatory for charts
    Invoke-Skill -Folder $OutDir -Step 'open' -PptArgs @('open', '-Doc', $deck, '-Show') | Out-Null

    Invoke-Skill -Folder $OutDir -Step 'chart-single' -PptArgs @('chart', '-Doc', $deck, '-Slide', '1', '-Data', (Join-Path $OutDir 'single.json'), '-Type', 'column', '-Title', 'Measured flow', '-YTitle', 'm3/s', '-XTitle', 'Year') | Out-Null
    Invoke-Skill -Folder $OutDir -Step 'chart-two' -PptArgs @('chart', '-Doc', $deck, '-Slide', '2', '-Data', (Join-Path $OutDir 'twoseries.json'), '-Type', 'barh', '-Title', 'Measured vs design') | Out-Null
    Invoke-Skill -Folder $OutDir -Step 'chart-line' -PptArgs @('chart', '-Doc', $deck, '-Slide', '3', '-Data', (Join-Path $OutDir 'line.json'), '-Type', 'line', '-Title', 'Stage') | Out-Null
    Invoke-Skill -Folder $OutDir -Step 'chart-bad' -PptArgs @('chart', '-Doc', $deck, '-Slide', '1', '-Data', (Join-Path $OutDir 'line.json'), '-Type', 'nosuchtype') | Out-Null
    Invoke-Skill -Folder $OutDir -Step 'chart-list' -PptArgs @('chart', '-Doc', $deck, '-Action', 'list') | Out-Null
    Invoke-Skill -Folder $OutDir -Step 'save' -PptArgs @('save', '-Doc', $deck) | Out-Null
    Invoke-Skill -Folder $OutDir -Step 'close' -PptArgs @('close', '-Doc', $deck, '-Save') | Out-Null
    Invoke-Skill -Folder $OutDir -Step 'quit2' -PptArgs @('quit') | Out-Null
    Stop-StrayWpp

    Copy-Item -LiteralPath $deck -Destination (Join-Path $OutDir 'deck.plain.pptx') -Force
    Decrypt-IfNeeded -Path (Join-Path $OutDir 'deck.plain.pptx')
}

Write-Host ''
Write-Host '== checks'
& $PythonExe (Join-Path $Root 'check_chart.py') --dir $OutDir
$code = $LASTEXITCODE
if ($code -eq 0) { Write-Host 'RESULT: PASS' -ForegroundColor Green } else { Write-Host 'RESULT: FAIL' -ForegroundColor Red }
exit $code
