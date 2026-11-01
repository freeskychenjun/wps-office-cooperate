@echo off
rem wps-office-cooperate fast entry point (v1.15).
rem
rem   woc.cmd <word|excel|ppt> <command> [flags...]
rem
rem Why this exists: on this machine the PowerShell child we used to spawn costs ~2.2 s before it
rem does anything - and the cost is NOT PowerShell itself, it is apartment initialization
rem (measured 2026-11-01: powershell -Sta = 2.2-2.4 s, powershell -Mta = 0.13-0.17 s,
rem cmd.exe /c exit = 0.03 s, a plain .NET exe = 0.03 s). The wrapper never uses COM - it only
rem builds a JSON request and talks to the resident helper - so it can run in MTA and that 2 s
rem simply disappears. The helper keeps [STAThread] because the WPS automation objects need it.
rem
rem Combined with the resident helper (scripts\helper\shared\PipeServer.cs) a command drops from
rem ~5-8 s to ~0.5-1.6 s. WOC_NO_DAEMON=1 still forces the one-shot path for A/B comparisons.
setlocal
set "BASE=%~dp0"
set "APP=%~1"
if "%APP%"=="" goto usage
if not exist "%BASE%%APP%.ps1" goto usage
shift
if "%~1"=="" goto usage

set "ARGS="
:collect
if "%~1"=="" goto run
set "ARGS=%ARGS% "%~1""
shift
goto collect

:run
powershell -Mta -NoProfile -ExecutionPolicy Bypass -File "%BASE%%APP%.ps1" %ARGS%
exit /b %ERRORLEVEL%

:usage
echo {"ok":false,"error":{"code":"BAD_ARGS","message":"usage: woc.cmd ^<word^|excel^|ppt^> ^<command^> [flags...]  (equivalent to office.ps1, but runs the wrapper in MTA so it does not pay ~2 s of STA init)"}}
exit /b 1
