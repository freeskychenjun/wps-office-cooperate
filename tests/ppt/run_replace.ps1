# run_replace.ps1 - acceptance driver for `ppt replace` (wps-office-cooperate, ppt side).
#
#   .\run_replace.ps1                     # full suite (starts WPS, ~2-3 minutes)
#   .\run_replace.ps1 -Only full,restore  # a subset of the scenarios
#   .\run_replace.ps1 -SkipRun            # just re-run the python checks
#
# For every scenario it makes a fresh fixture (make_deck.py), drives the skill
# exactly the way an agent would, drops each JSON response next to the deck and
# finishes with check.py, which owns every assertion (it is the only place that
# has to deal with Chinese text, and python handles UTF-8 natively).
#
# Pure ASCII on purpose: Windows PowerShell 5.1 parses .ps1 as ANSI, so a Chinese
# literal in this file would be mangled.  The fixtures carry the Chinese.
#
[CmdletBinding()]
param(
    [string]$PythonExe = 'python',
    [string]$OutDir = '',
    [string[]]$Only = @(),
    # On this machine WPS writes .pptx transparently encrypted (header 17 DA 5F A0...),
    # so python-pptx cannot read what WPS saved.  FastDycrypt (offline algorithm, ~9 s
    # per file, any format) is the default way out; point -DecryptEngine elsewhere if
    # this machine does not have it.  The old disguise-reader route is the fallback.
    # Left empty on purpose: this machine's FastDycrypt lives under "D:\DSH工作区\...", and a
    # Chinese literal inside this .ps1 would be read as ANSI and arrive as mojibake.
    [string]$DecryptEngine = '',
    [string]$DecryptScript = "$env:USERPROFILE\.agents\skills\decrypt-office-pdf\decrypt-files.ps1",
    # Decrypting one deck costs ~45 s (external tool, one file at a time), so only the
    # scenarios that assert on the saved FILE get it; -SkipDecrypt runs the suite in a
    # couple of minutes and reports those four as "file check skipped".
    [switch]$SkipDecrypt,
    [switch]$SkipRun
)

$ErrorActionPreference = 'Stop'
try { [Console]::OutputEncoding = New-Object System.Text.UTF8Encoding($false) } catch { }

# `powershell -File script.ps1 -Only a,b` passes "a,b" as ONE string (the -File
# switches to verbatim arguments), so split it here instead of relying on the parser.
$Only = @($Only | ForEach-Object { "$_" -split ',' } | ForEach-Object { $_.Trim() } | Where-Object { $_ })

$Root = $PSScriptRoot
$Repo = Split-Path -Parent (Split-Path -Parent $Root)
$Ppt = Join-Path $Repo 'scripts\ppt.ps1'
$Utf8 = New-Object System.Text.UTF8Encoding($false)
if (-not $OutDir) { $OutDir = Join-Path $Root 'out' }

# Machine-local tool paths come from tests/ppt/local.config.json (UTF-8, gitignored):
# this machine's FastDycrypt lives under "D:\DSH工作区\...", and a Chinese literal inside
# this .ps1 would be read as ANSI and arrive as mojibake.
if (-not $DecryptEngine) {
    $cfgPath = Join-Path $Root 'local.config.json'
    if (Test-Path -LiteralPath $cfgPath) {
        $cfg = [System.IO.File]::ReadAllText($cfgPath, [System.Text.Encoding]::UTF8) | ConvertFrom-Json
        if ($cfg.decryptEngine) { $DecryptEngine = [string]$cfg.decryptEngine }
        if (-not $DecryptScript -and $cfg.decryptScript) { $DecryptScript = [string]$cfg.decryptScript }
    }
}

if (-not (Test-Path -LiteralPath $Ppt)) { throw "skill wrapper not found: $Ppt" }
if (-not $SkipRun) {
    if (-not (Test-Path -LiteralPath $OutDir)) { New-Item -ItemType Directory -Force -Path $OutDir | Out-Null }
}

function Get-WpsCount {
    $lines = & tasklist /FI 'IMAGENAME eq wpp.exe' /NH 2>$null
    return ($lines | Select-String -SimpleMatch 'wpp.exe' | Measure-Object).Count
}

function Invoke-Ppt {
    <# Runs one skill call and saves its JSON response as <folder>\resp-<step>.json.
       The wrapper writes its answer with [Console]::Out, so it has to be a child
       process to be captured at all. #>
    param([string]$Folder, [string]$Step, [string[]]$PptArgs)
    $text = (& powershell -NoProfile -ExecutionPolicy Bypass -File $Ppt @PptArgs) -join "`n"
    $path = Join-Path $Folder ('resp-' + $Step + '.json')
    [System.IO.File]::WriteAllText($path, $text, $Utf8)
    $short = ($text -replace '\s+', ' ')
    if ($short.Length -gt 150) { $short = $short.Substring(0, 150) + '...' }
    Write-Host ("    {0,-14} {1}" -f $Step, $short)
    if ($LASTEXITCODE -ne 0 -and $LASTEXITCODE -ne 2) {
        Write-Host ("    !! exit code {0}" -f $LASTEXITCODE) -ForegroundColor Yellow
    }
}

function New-Fixture {
    param([string]$Folder, [string]$Mode = 'full')
    if (Test-Path -LiteralPath $Folder) {
        try { Remove-Item -Recurse -Force -LiteralPath $Folder }
        catch {
            # WPS still holds a lock from an interrupted run: move it aside instead.
            $stale = $Folder + '.stale-' + (Get-Date -Format 'HHmmss')
            Move-Item -LiteralPath $Folder -Destination $stale -Force
            Write-Host ("    previous folder was locked, moved to {0}" -f (Split-Path -Leaf $stale)) -ForegroundColor Yellow
        }
    }
    New-Item -ItemType Directory -Force -Path $Folder | Out-Null
    & $PythonExe (Join-Path $Root 'make_deck.py') --out $Folder --mode $Mode | Out-Null
    if ($LASTEXITCODE -ne 0) { throw "make_deck.py failed for $Folder" }
    $prefix = if ($Mode -eq 'footer') { 'footer-deck' } else { 'replace-deck' }
    Move-Item -LiteralPath (Join-Path $Folder ($prefix + '.pptx')) (Join-Path $Folder 'deck.pptx') -Force
    Move-Item -LiteralPath (Join-Path $Folder ($prefix + '.pristine.pptx')) (Join-Path $Folder 'pristine.pptx') -Force
    return (Join-Path $Folder 'deck.pptx')
}

function Write-Utf8File {
    param([string]$Path, [string]$Text)
    [System.IO.File]::WriteAllText($Path, $Text, $Utf8)
}

function Publish-DeckForCheck {
    <# Leaves a readable copy next to the deck as deck.plain.pptx.
       On this machine WPS writes .pptx transparently encrypted, and the only way to
       read it back is the disguised reader from the decrypt skill.  That reader is a
       single-instance app: a leftover one from an earlier run blocks the next call
       forever, so kill any stray one first (only the disguised copy, never a real
       PowerPoint the user has open). #>
    param([string]$DeckPath)
    $folder = Split-Path -Parent $DeckPath
    $plain = Join-Path $folder 'deck.plain.pptx'
    Copy-Item -LiteralPath $DeckPath -Destination $plain -Force
    $bytes = [System.IO.File]::ReadAllBytes($plain)
    $encrypted = ($bytes.Length -gt 8 -and $bytes[0] -eq 0x17 -and $bytes[1] -eq 0xDA -and $bytes[2] -eq 0x5F)
    if (-not $encrypted) { return 'plain' }

    if (Test-Path -LiteralPath $DecryptEngine) {
        & powershell -NoProfile -ExecutionPolicy Bypass -File $DecryptEngine -Paths $plain `
            -InPlace -DeleteBackupAfterSuccess -NoPause -NoLog | Out-Null
    }
    elseif (Test-Path -LiteralPath $DecryptScript) {
        Get-Process powerpnt -ErrorAction SilentlyContinue |
            Where-Object { $_.Path -like '*右键解密*' } |
            ForEach-Object { Write-Host '    killing a stray disguised powerpnt'; Stop-Process -Id $_.Id -Force }
        & powershell -NoProfile -ExecutionPolicy Bypass -File $DecryptScript $plain | Out-Null
    }
    else {
        throw ("deck is encrypted and no decryptor was found (neither " + $DecryptEngine + " nor " + $DecryptScript + ")")
    }

    $bytes = [System.IO.File]::ReadAllBytes($plain)
    if (-not ($bytes[0] -eq 0x50 -and $bytes[1] -eq 0x4B)) { throw ("decrypt did not produce a ZIP header: " + $plain) }
    return 'decrypted'
}

# ---------------------------------------------------------------- COM helpers
# The skill has no command for HeadersFooters, so the footer scenario sets it the
# way a human would: straight on the open presentation.

function Get-OpenPresentation([string]$DeckPath) {
    $app = [Runtime.InteropServices.Marshal]::GetActiveObject('KWPP.Application')
    $full = [System.IO.Path]::GetFullPath($DeckPath)
    for ($i = 1; $i -le $app.Presentations.Count; $i++) {
        $p = $app.Presentations.Item($i)
        if ($p.FullName -eq $full) { return $p }
    }
    throw "presentation is not open in WPS: $full"
}

# ---------------------------------------------------------------- scenarios

$scenarios = @(
    @{ Name = 'full';   Mode = 'full'; Deep = $true; Steps = @(
        @{ Step = 'replace';       Args = @('replace', '-Find', '2024', '-Replace', '2025') },
        @{ Step = 'replace-again'; Args = @('replace', '-Find', '2024', '-Replace', '2025') },
        @{ Step = 'save';          Args = @('save') } ) },

    @{ Name = 'dryrun'; Mode = 'full'; Deep = $true; Steps = @(
        @{ Step = 'replace'; Args = @('replace', '-Find', '2024', '-Replace', '2025', '-DryRun') },
        @{ Step = 'save';    Args = @('save') } ) },

    @{ Name = 'restore'; Mode = 'full'; Deep = $true; Steps = @(
        @{ Step = 'replace'; Args = @('replace', '-Find', '2024', '-Replace', '2025') },
        @{ Step = 'save';    Args = @('save') },
        @{ Step = 'restore'; Args = @('restore', '-Since', 'prereplace') },
        @{ Step = 'save';    Args = @('save') } ) },

    @{ Name = 'notes';  Mode = 'full'; Steps = @(
        @{ Step = 'replace'; Args = @('replace', '-Find', '2024', '-Replace', '2025', '-Scope', 'notes') },
        @{ Step = 'save';    Args = @('save') } ) },

    @{ Name = 'master'; Mode = 'full'; Steps = @(
        @{ Step = 'replace'; Args = @('replace', '-Find', '2024', '-Replace', '2025', '-Scope', 'master') },
        @{ Step = 'save';    Args = @('save') } ) },

    @{ Name = 'slidesonly'; Mode = 'full'; Steps = @(
        @{ Step = 'replace'; Args = @('replace', '-Find', '2024', '-Replace', '2025', '-Scope', 'slides') },
        @{ Step = 'save';    Args = @('save') } ) },

    @{ Name = 'slide1'; Mode = 'full'; Steps = @(
        @{ Step = 'replace'; Args = @('replace', '-Find', '2024', '-Replace', '2025', '-Slides', '1') },
        @{ Step = 'save';    Args = @('save') } ) },

    @{ Name = 'regex';  Mode = 'full'; Steps = @(
        @{ Step = 'replace'; Args = @('replace', '-Find', '20(\d\d)', '-Replace', 'X$1', '-Regex') },
        @{ Step = 'save';    Args = @('save') } ) },

    @{ Name = 'case';   Mode = 'full'; Deep = $true; Steps = @(
        @{ Step = 'replace'; Args = @('replace', '-Find', 'aa2024', '-Case') },
        @{ Step = 'save';    Args = @('save') } ) },

    @{ Name = 'max';    Mode = 'full'; Steps = @(
        @{ Step = 'replace'; Args = @('replace', '-Find', '2024', '-Replace', '2025', '-Max', '3') },
        @{ Step = 'save';    Args = @('save') } ) },

    @{ Name = 'items';  Mode = 'full'; Steps = @(
        @{ Step = 'replace'; Args = @('replace', '-Items', '{ITEMS}') },
        @{ Step = 'save';    Args = @('save') } ) },

    @{ Name = 'footer'; Mode = 'footer'; Steps = @(
        @{ Step = 'replace'; Args = @('replace', '-Find', '2024', '-Replace', '2025', '-Scope', 'footer') },
        @{ Step = 'save';    Args = @('save') } ) }
)

if (-not $SkipRun) {
    # The suite owns WPS Presentation: an instance left over from an interrupted run
    # still holds the previous fixtures open, so start from a clean one.
    & powershell -NoProfile -ExecutionPolicy Bypass -File $Ppt quit | Out-Null
    for ($i = 0; $i -lt 20; $i++) {
        if ((Get-WpsCount) -eq 0) { break }
        Start-Sleep -Milliseconds 500
    }
    # `quit` only quits the instance the helper attached to; a second headless one would
    # keep the fixtures locked.  Kill only windowless instances - never a deck the user
    # has open (that one has a title).
    if ((Get-WpsCount) -gt 0) {
        Get-Process wpp -ErrorAction SilentlyContinue |
            Where-Object { -not $_.MainWindowTitle } |
            ForEach-Object { try { Stop-Process -Id $_.Id -Force } catch { } }
        Start-Sleep -Milliseconds 500
    }
    Write-Host ("wpp.exe before: {0}" -f (Get-WpsCount))
    foreach ($scenario in $scenarios) {
        if ($Only.Count -gt 0 -and ($Only -notcontains $scenario.Name)) { continue }
        Write-Host ("-- {0}" -f $scenario.Name) -ForegroundColor Cyan
        $folder = Join-Path $OutDir $scenario.Name
        $deck = New-Fixture -Folder $folder -Mode $scenario.Mode

        & powershell -NoProfile -ExecutionPolicy Bypass -File $Ppt open -Doc $deck
        if ($LASTEXITCODE -ne 0) { Write-Host '    !! open failed' -ForegroundColor Yellow }

        if ($scenario.Name -eq 'items') {
            $itemsPath = Join-Path $folder 'items.json'
            Write-Utf8File -Path $itemsPath -Text '{"items":[{"find":"2024","replace":"2025"},{"find":"2023","replace":"2022"}]}'
        } else {
            $itemsPath = ''
        }
        if ($scenario.Name -eq 'footer') {
            $pres = Get-OpenPresentation -DeckPath $deck
            $hf = $pres.Slides.Item(1).HeadersFooters
            $hf.Footer.Visible = -1
            $hf.Footer.Text = 'ACME 2024'
            Write-Host '    footer set over COM'
        }

        foreach ($step in $scenario.Steps) {
            $args2 = @()
            foreach ($arg in $step.Args) { $args2 += ($arg -replace '\{ITEMS\}', $itemsPath) }
            $args2 += @('-Doc', $deck)
            Invoke-Ppt -Folder $folder -Step $step.Step -PptArgs $args2
        }

        if ($scenario.Name -eq 'footer') {
            $pres = Get-OpenPresentation -DeckPath $deck
            $after = [string]$pres.Slides.Item(1).HeadersFooters.Footer.Text
            $state = '{"footerAfter":"' + $after + '"}'
            Write-Utf8File -Path (Join-Path $folder 'state.json') -Text $state
            Write-Host ("    footer after  {0}" -f $after)
        }

        & powershell -NoProfile -ExecutionPolicy Bypass -File $Ppt close -Doc $deck -Save | Out-Null
        if (-not $SkipDecrypt -and $scenario.Deep) {
            $how = Publish-DeckForCheck -DeckPath $deck
            Write-Host ("    deck for check: {0}" -f $how)
        }
    }
    & powershell -NoProfile -ExecutionPolicy Bypass -File $Ppt quit | Out-Null
    Write-Host ("wpp.exe after: {0}" -f (Get-WpsCount))
}

Write-Host ''
Write-Host '== checks'
& $PythonExe (Join-Path $Root 'check.py') --dir $OutDir
$code = $LASTEXITCODE
if ($code -eq 0) { Write-Host 'RESULT: PASS' -ForegroundColor Green } else { Write-Host 'RESULT: FAIL' -ForegroundColor Red }
exit $code
