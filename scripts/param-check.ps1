# Shared parameter-consumption check for the three wrappers (word / excel / ppt), v1.18.
#
# Why this exists: the wrappers build a JSON request from $PSBoundParameters, and a parameter that
# is bound by PowerShell but never copied into the request disappears silently. That is how
# `excel chart -Sheet Sheet1` used to put the chart on the ActiveSheet and still return ok:true
# (reported 2026-11-01 from another machine, ~6 min of diagnosis), and how several Word/PPT flags
# can look accepted while doing nothing.
#
# Each wrapper calls:
#   . "$PSScriptRoot\param-check.ps1"            (near the top, after the param block)
#   Test-ParamConsumption -Command $Command -Bound $PSBoundParameters -Request $request `
#       -Aliases $paramAliases -FormatFields $formatFields -CommandSkips $paramCommandSkips
# right after its command switch has finished building $request.
#
# Resolution order for a bound parameter name:
#   1. $Aliases      - explicit "parameter -> request key" mapping (non-mechanical names)
#   2. $FormatFields - the app's format map (params applied dynamically by Add-Format-Fields)
#   3. lower-case the first letter and look for that key in $request (the common convention)
# Anything still unresolved warns on stderr; stdout (the JSON answer) is untouched, so this can
# never break a call - it only makes a silent drop visible. WOC_NO_PARAMCHECK=1 disables it.

function Test-ParamConsumption {
    param(
        [string]$Command,
        [System.Collections.IDictionary]$Bound,
        [System.Collections.IDictionary]$Request,
        [System.Collections.IDictionary]$Aliases,
        [System.Collections.IDictionary]$FormatFields,
        [System.Collections.IDictionary]$CommandSkips,
        [System.Collections.IDictionary]$DefaultValues
    )
    if ($env:WOC_NO_PARAMCHECK -eq '1') { return }
    if ($null -eq $Bound -or $null -eq $Request) { return }
    $alwaysSkip = @('Command', 'TimeoutSec', 'Delete')
    $skip = @()
    if ($CommandSkips -and $CommandSkips.ContainsKey('*')) { $skip = @($skip + @($CommandSkips['*'])) }
    if ($CommandSkips -and $CommandSkips.ContainsKey($Command)) { $skip = @($skip + @($CommandSkips[$Command])) }
    foreach ($p in @($Bound.Keys)) {
        if ($alwaysSkip -contains $p) { continue }
        if ($skip -contains $p) { continue }
        # Explicitly passing a parameter's DEFAULT value is indistinguishable from omitting it, and
        # the wrappers forward numeric options only when they differ from the default (e.g.
        # `if ($From -gt 1)`), so a default-valued argument is not a drop - stay quiet.
        if ($DefaultValues -and $DefaultValues.ContainsKey($p) -and ("$($Bound[$p])" -eq "$($DefaultValues[$p])")) { continue }
        $key = $null
        if ($Aliases -and $Aliases.ContainsKey($p)) { $key = $Aliases[$p] }
        elseif ($FormatFields -and $FormatFields.ContainsKey($p)) { $key = $FormatFields[$p] }
        else { $key = $p.Substring(0, 1).ToLower() + $p.Substring(1) }
        if ($Request.Contains($key)) { continue }
        [Console]::Error.WriteLine('WARN ignored parameter -' + $p + ' for command "' + $Command + '": this command does not forward it (set WOC_NO_PARAMCHECK=1 to silence)')
    }
}
