<#
.SYNOPSIS
    Shared helpers for the OnBase Provisioning Framework.
.DESCRIPTION
    Dot-source from any step script or the orchestrator:
        . "$PSScriptRoot\Common.ps1"
    Backslashes for UNC paths are built from [char]92 so the code survives copy/paste
    through chat tools that strip double-backslash paths.
#>

# Single backslash character
$script:BS = [string][char]92

# True when the path is a UNC share path (starts with two backslashes)
function Test-UncPath([string]$Path) {
    return ($Path -and $Path.StartsWith($script:BS + $script:BS))
}

# Joins path segments with a backslash, without collapsing the leading UNC backslashes
function Join-SharePath([string]$Base, [string[]]$Child) {
    $result = $Base.TrimEnd($script:BS)
    foreach ($c in $Child) { $result = $result + $script:BS + $c.Trim($script:BS) }
    return $result
}

# Rejects an empty machine name or one pasted from chat with link formatting
function Assert-TargetMachine([string]$TargetMachine) {
    if (-not $TargetMachine) { throw 'TargetMachine is required for Remote mode.' }
    if ($TargetMachine -match '[\[\]\(\)]|http') {
        throw "TargetMachine '$TargetMachine' contains link formatting. Type the plain name, e.g. RDV-009989.hylandqa.net"
    }
}

# True when the current PowerShell session is elevated
function Test-IsAdmin {
    $id = [Security.Principal.WindowsIdentity]::GetCurrent()
    return ([Security.Principal.WindowsPrincipal]$id).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

# Runs a script block on this PC (Local) or on the target VM (Remote, via WinRM / RunAs endpoint)
function Invoke-OnTarget {
    param(
        [string]$DeploymentMode,
        [string]$TargetMachine,
        [pscredential]$Credential,
        [string]$ConfigurationName,
        [scriptblock]$ScriptBlock,
        [object[]]$ArgumentList = @()
    )
    if ($DeploymentMode -eq 'Local') {
        return & $ScriptBlock @ArgumentList
    }
    Assert-TargetMachine $TargetMachine
    $p = @{ ComputerName = $TargetMachine; ScriptBlock = $ScriptBlock; ArgumentList = $ArgumentList }
    if ($Credential)        { $p.Credential = $Credential; $p.Authentication = 'Negotiate' }
    if ($ConfigurationName) { $p.ConfigurationName = $ConfigurationName }
    return Invoke-Command @p
}

# Keeps only the last output object that has the given property (drops anything else that leaked out)
function Select-Result($Output, [string]$Property) {
    return ($Output | Where-Object { $_ -and $_.PSObject.Properties[$Property] } | Select-Object -Last 1)
}