# run_check.ps1 - acceptance driver for `ppt check` / `ppt outline` (wps-office-cooperate).
#
#   .\run_check.ps1
#   .\run_check.ps1 -SkipRun       # re-run only the python assertions
#
# Fixture (make_check_deck.py) seeds one defect per rule on its own slide and leaves
# two slides clean; the assertions in check_check.py then say, rule by rule, "this one
# fired on the right slide and only there".  A gate that cries wolf is worse than no
# gate, so the clean-slide assertions matter as much as the firing ones.
#
# Slide numbers are switched on over COM after opening: python-pptx can write the
# <p:hf> flag but WPS ignores it, and the rule reads HeadersFooters.SlideNumber.
#
# Pure ASCII on purpose (Windows PowerShell 5.1 parses .ps1 as ANSI).
#
[CmdletBinding()]
param(
    [string]$PythonExe = 'python',
    [string]$OutDir = '',
    [switch]$SkipRun
)

$ErrorActionPreference = 'Stop'
try { [Console]::OutputEncoding = New-Object System.Text.UTF8Encoding($false) } catch { }

$Root = $PSScriptRoot
$Repo = Split-Path -Parent (Split-Path -Parent $Root)
$Ppt = Join-Path $Repo 'scripts\ppt.ps1'
$Utf8 = New-Object System.Text.UTF8Encoding($false)
if (-not $OutDir) { $OutDir = Join-Path $Root 'out-check' }

function Quote-Arg {
    <# Start-Process joins -ArgumentList with spaces and quotes nothing, so an argument
       with a space in it arrives as two arguments and the skill refuses to bind. #>
    param([string]$Value)
    if ($null -eq $Value) { return '""' }
    if ($Value -notmatch '[\s"]') { return $Value }
    return '"' + ($Value -replace '"', '\"') + '"'
}

function Invoke-Ppt {
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
    Write-Host ("    {0,-12} {1}" -f $Step, $short)
}

function Get-OpenPresentation([string]$DeckPath) {
    $app = [Runtime.InteropServices.Marshal]::GetActiveObject('KWPP.Application')
    $full = [System.IO.Path]::GetFullPath($DeckPath)
    for ($i = 1; $i -le $app.Presentations.Count; $i++) {
        $p = $app.Presentations.Item($i)
        if ($p.FullName -eq $full) { return $p }
    }
    throw "presentation is not open in WPS: $full"
}

if (-not $SkipRun) {
    & powershell -NoProfile -ExecutionPolicy Bypass -File $Ppt quit | Out-Null
    Start-Sleep -Milliseconds 700
    if (Test-Path -LiteralPath $OutDir) {
        try { Remove-Item -Recurse -Force -LiteralPath $OutDir } catch { }
    }
    New-Item -ItemType Directory -Force -Path $OutDir | Out-Null
    $deck = Join-Path $OutDir 'check-deck.pptx'
    & $PythonExe (Join-Path $Root 'make_check_deck.py') --out $OutDir | Out-Null
    if ($LASTEXITCODE -ne 0) { throw 'make_check_deck.py failed' }

    $folder = $OutDir
    Invoke-Ppt -Folder $folder -Step 'open' -PptArgs @('open', '-Doc', $deck)

    # slide numbers on for 2..13, off for the cover, off for the seeded "missing" one
    $pres = Get-OpenPresentation -DeckPath $deck
    $slides = $pres.Slides.Count
    for ($k = 1; $k -le $slides; $k++) {
        $want = ($k -ne 1 -and $k -ne 6)
        $pres.Slides.Item($k).HeadersFooters.SlideNumber.Visible = if ($want) { -1 } else { 0 }
    }
    Write-Host ("    slide numbers set on {0} slides" -f $slides)

    Invoke-Ppt -Folder $folder -Step 'check' -PptArgs @('check', '-Doc', $deck)
    Invoke-Ppt -Folder $folder -Step 'check-blocker' -PptArgs @('check', '-Doc', $deck, '-MinSeverity', 'blocker')
    Invoke-Ppt -Folder $folder -Step 'check-rules' -PptArgs @('check', '-Doc', $deck, '-Rules', 'overflow,outframe')
    Invoke-Ppt -Folder $folder -Step 'check-profile' -PptArgs @('check', '-Doc', $deck, '-Profile', 'bid')
    Invoke-Ppt -Folder $folder -Step 'check-md' -PptArgs @('check', '-Doc', $deck, '-Out', (Join-Path $folder 'report.md'))
    # config: raise nothing, but silence the 9pt body floor and whitelist one shape
    $config = Join-Path $folder 'rules.json'
    [System.IO.File]::WriteAllText($config,
        '{"bodyMin":9,"suppress":[{"slide":5,"shape":"outbox"}]}', $Utf8)
    Invoke-Ppt -Folder $folder -Step 'check-config' -PptArgs @('check', '-Doc', $deck, '-Config', $config)
    Invoke-Ppt -Folder $folder -Step 'outline' -PptArgs @('outline', '-Doc', $deck, '-Out', (Join-Path $folder 'outline.md'))
    # -ExitCode must fail the call when blockers exist
    & powershell -NoProfile -ExecutionPolicy Bypass -File $Ppt check -Doc $deck -ExitCode | Out-Null
    $gate = $LASTEXITCODE
    [System.IO.File]::WriteAllText((Join-Path $folder 'gate.json'), ('{"exitCode":' + $gate + '}'), $Utf8)
    Write-Host ("    exit-code gate  {0}" -f $gate)

    & powershell -NoProfile -ExecutionPolicy Bypass -File $Ppt close -Doc $deck | Out-Null
    & powershell -NoProfile -ExecutionPolicy Bypass -File $Ppt quit | Out-Null
}

Write-Host ''
Write-Host '== checks'
& $PythonExe (Join-Path $Root 'check_check.py') --dir $OutDir
$code = $LASTEXITCODE
if ($code -eq 0) { Write-Host 'RESULT: PASS' -ForegroundColor Green } else { Write-Host 'RESULT: FAIL' -ForegroundColor Red }
exit $code
