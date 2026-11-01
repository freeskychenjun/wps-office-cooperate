# wps-office-cooperate shell-free entry point (v1.16).
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

& powershell -Mta -NoProfile -ExecutionPolicy Bypass -File $target @rest
exit $LASTEXITCODE
