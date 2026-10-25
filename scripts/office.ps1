# wps-office-cooperate: app dispatcher (word | excel | ppt).
# Pure ASCII source. Forwards all remaining arguments to the per-app entry script:
#
#   office.ps1 word|excel|ppt <command> [options]
#
# Each app entry (word.ps1 / excel.ps1 / ppt.ps1) is fully standalone: you may also
# call it directly with identical arguments. Output is a single JSON line.
#
# NOTE: no param() block on purpose - manual $args parsing is the only way to forward
# arbitrary named parameters (-Doc etc.) that this script does not declare.

$ErrorActionPreference = 'Stop'
try { [Console]::OutputEncoding = New-Object System.Text.UTF8Encoding($false) } catch { }

function Write-Error-Json([string]$code, [string]$message) {
    $payload = @{ ok = $false; error = @{ code = $code; message = $message } }
    [Console]::Out.Write(($payload | ConvertTo-Json -Depth 6 -Compress))
    exit 1
}

if ($args.Count -lt 2) {
    Write-Error-Json 'USAGE' "office.ps1 <word|excel|ppt> <command> [options]"
}

$app = [string]$args[0]
$entry = Join-Path $PSScriptRoot ($app + '.ps1')
if ($app -notin @('word', 'excel', 'ppt') -or -not (Test-Path -LiteralPath $entry)) {
    if ($app -eq 'ppt') {
        Write-Error-Json 'APP_NOT_SUPPORTED' "PPT (WPS Presentation) is not implemented yet. Available: word, excel."
    }
    Write-Error-Json 'BAD_APP' ("Unknown app: " + $app + ". Available: word, excel.")
}

$rest = @()
if ($args.Count -gt 1) { $rest = $args[1..($args.Count - 1)] }

& $entry @rest
exit $LASTEXITCODE
