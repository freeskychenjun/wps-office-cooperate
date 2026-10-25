# run_acceptance.ps1 - one-command acceptance chain for the wps-word-cooperate skill.
#
#   python make_corpus.py  ->  corpus/report-300.docx + pristine + targets.json
#   python make_edits.py   ->  corpus/edits.json (10 revise items)
#   (manual step: run the skill's `revise` + `comment` on a COPY, then)
#   python validate.py --doc <revised.docx> --targets ... --edits ...
#
# The skill wrapper (scripts/wps.ps1) is intentionally NOT called from here:
# it is still being written.  This script prints the exact command to run.
#
# Pure ASCII on purpose (Windows PowerShell 5.1 parses .ps1 as ANSI).
#
# Examples:
#   .\run_acceptance.ps1
#   .\run_acceptance.ps1 -RevisedDoc .\corpus\revised.docx -PngDir .\corpus\png
#   .\run_acceptance.ps1 -SkipValidate
#
[CmdletBinding()]
param(
    [string]$PythonExe = 'python',
    [string]$RevisedDoc = '',
    [string]$PngDir = '',
    [int]$ExpectComments = 2,
    [switch]$SkipValidate
)

$ErrorActionPreference = 'Stop'
$Root   = $PSScriptRoot
$Corpus = Join-Path $Root 'corpus'

# Python prints UTF-8 JSON; make the console show it correctly (PS 5.1 defaults to ANSI).
try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch { }

function Get-WpsCount {
    $lines = & tasklist /FI 'IMAGENAME eq wps.exe' /NH 2>$null
    return ($lines | Select-String -SimpleMatch 'wps.exe' | Measure-Object).Count
}

function Invoke-Py {
    param([string]$Script, [string[]]$PyArgs = @())
    Write-Host ("  > {0} {1}" -f $PythonExe, (Split-Path $Script -Leaf))
    & $PythonExe $Script @PyArgs
    if ($LASTEXITCODE -ne 0) {
        throw ("{0} failed with exit code {1}" -f (Split-Path $Script -Leaf), $LASTEXITCODE)
    }
}

$wpsBefore = Get-WpsCount
Write-Host ("wps.exe before: {0}" -f $wpsBefore)

Write-Host '[1/3] generating corpus'
Invoke-Py -Script (Join-Path $Root 'make_corpus.py')

Write-Host '[2/3] generating revise items'
Invoke-Py -Script (Join-Path $Root 'make_edits.py')

Write-Host ''
Write-Host 'NEXT (manual step - the skill wrapper is not part of this repo yet):'
Write-Host ('  wps.ps1 revise -Doc "{0}" -Items "{1}"' -f (Join-Path $Corpus 'report-300.docx'), (Join-Path $Corpus 'edits.json'))
Write-Host ('  wps.ps1 comment -Doc "{0}" -Para 86  -TextFile <txt>' -f (Join-Path $Corpus 'report-300.docx'))
Write-Host ('  wps.ps1 comment -Doc "{0}" -Para 151 -TextFile <txt>' -f (Join-Path $Corpus 'report-300.docx'))
Write-Host ('  wps.ps1 save   -Doc "{0}" -As "{1}"' -f (Join-Path $Corpus 'report-300.docx'), (Join-Path $Corpus 'revised.docx'))
Write-Host '  NOTE: the saved copy becomes ciphertext (17DA5FA01633CD9A) - that is expected.'
Write-Host ''

if ($SkipValidate -or [string]::IsNullOrWhiteSpace($RevisedDoc)) {
    Write-Host 'validate step skipped (pass -RevisedDoc <path> to run it).'
    Write-Host ('wps.exe after: {0}' -f (Get-WpsCount))
    exit 0
}

$revised = $RevisedDoc
if (-not [System.IO.Path]::IsPathRooted($revised)) {
    $revised = Join-Path $Root $revised
}
if (-not (Test-Path -LiteralPath $revised)) {
    throw ("revised document not found: {0}" -f $revised)
}

$valArgs = @(
    '--doc',        $revised,
    '--targets',    (Join-Path $Corpus 'targets.json'),
    '--edits',      (Join-Path $Corpus 'edits.json'),
    '--expect-comments', [string]$ExpectComments
)
if (-not [string]::IsNullOrWhiteSpace($PngDir)) {
    $png = $PngDir
    if (-not [System.IO.Path]::IsPathRooted($png)) { $png = Join-Path $Root $png }
    $valArgs += @('--png-dir', $png)
}

Write-Host '[3/3] validating'
& $PythonExe (Join-Path $Root 'validate.py') @valArgs
$code = $LASTEXITCODE

Write-Host ''
Write-Host ("verdict exit code: {0} (0 = pass, 1 = fail)" -f $code)
Write-Host ("report: {0}" -f (Join-Path $Root 'validate-report.json'))
Write-Host ("wps.exe after: {0} (before: {1})" -f (Get-WpsCount), $wpsBefore)
exit $code
