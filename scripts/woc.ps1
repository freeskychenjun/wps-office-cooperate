# wps-office-cooperate shell-free entry point (v1.16, error envelope v1.34).
#
#   woc.ps1 <word|excel|ppt> <command> [flags...]
#
# Why this exists next to woc.cmd: woc.cmd goes through cmd.exe, and cmd.exe rewrites the
# command line BEFORE the batch file sees it - a bare token starting with "=" loses its "=",
# and "^" is consumed as the cmd escape character. Measured 2026-11-01 on another machine:
#
#   & woc.cmd excel fill -Formula "=A15^1.5"   -> the child received  A151.5   (broken)
#   & woc.cmd excel fill -Formula "=1+1"       -> the child received  1+1      (broken)
#
# PowerShell-to-PowerShell argument passing does not go through cmd.exe, so this wrapper keeps
# every character: formulas, ^, &, |, >, and $ all arrive intact. Use it from PowerShell (or any
# CreateProcessW caller); keep woc.cmd only for cmd.exe callers that pass simple arguments.
#
# It still spawns the real wrapper with -Mta, because the ~2 s tax measured on this machine is
# STA apartment initialization, not PowerShell itself (see woc.cmd and SKILL.md for the numbers).
#
# v1.34: the child's stderr is staged to a temp file and replayed after it exits. PowerShell
# parameter-binding failures (bad command name / unknown flag / bad ValidateSet value) happen
# BEFORE the wrapper's own error handling, so they normally leave stdout empty and print host text
# to stderr - the one common misuse that used to break the "stdout is always one JSON object"
# contract. For that class we now also emit an ASCII BAD_ARGS envelope, so a caller can always
# parse stdout. Everything else (JSON responses, warnings, exit codes) is passed through unchanged.

$app = $args[0]
if (-not $app) {
    [Console]::Out.Write('{"ok":false,"error":{"code":"BAD_ARGS","message":"usage: woc.ps1 <word|excel|ppt> <command> [flags...]"}}')
    exit 1
}
$target = Join-Path $PSScriptRoot ($app + '.ps1')
if (-not (Test-Path -LiteralPath $target)) {
    [Console]::Out.Write('{"ok":false,"error":{"code":"BAD_ARGS","message":"unknown app: ' + $app + ' (expected word|excel|ppt)"}}')
    exit 1
}
$rest = @($args | Select-Object -Skip 1)
if ($rest.Count -eq 0) {
    [Console]::Out.Write('{"ok":false,"error":{"code":"BAD_ARGS","message":"no command given"}}')
    exit 1
}

$errFile = Join-Path $env:TEMP ('woc-err-' + [guid]::NewGuid().ToString('N') + '.txt')
$code = 0
$errText = ''
try {
    & powershell -Mta -NoProfile -ExecutionPolicy Bypass -File $target @rest 2> $errFile
    $code = $LASTEXITCODE
} finally {
    if (Test-Path -LiteralPath $errFile) {
        # PS 5.1 writes redirected native stderr in the console/ANSI code page, hence Default here.
        try { $errText = [System.IO.File]::ReadAllText($errFile, [System.Text.Encoding]::Default) } catch { }
        try { Remove-Item -Force -LiteralPath $errFile } catch { }
    }
}
if ($errText) { [Console]::Error.Write($errText) }

$bindingFailure = ($errText -match 'ParameterArgumentValidationError|Cannot validate argument|Missing an argument for parameter|AmbiguousParameter|A parameter cannot be found')
if ($code -ne 0 -and $bindingFailure) {
    $msg = 'invalid command or parameter for the ' + $app + ' wrapper (see stderr for the PowerShell binding detail)'
    [Console]::Out.Write('{"ok":false,"error":{"code":"BAD_ARGS","message":"' + $msg + '"}}')
}
exit $code
