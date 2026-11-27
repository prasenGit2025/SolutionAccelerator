#requires -Version 5.1
<#
.SYNOPSIS
Read-only readiness checks and a validated child-invocation plan for OnBase.
.DESCRIPTION
Returns ONE report object: Passed, Checks, Context, Steps. Does not call exit,
install software, enable WinRM, register endpoints, change TrustedHosts/firewall,
or create certificates. Temporary PSSessions/credentialed PSDrives are removed.
Run in 64-bit Windows PowerShell 5.1. This supports SQL on the target VM, not a
separate SQL host. Remote checks require the configured endpoint to permit the
commands the installers need; the name UnityInstall alone does not imply JEA.

Existing JSON keys are preserved. Additional optional ROOT keys supported:
buildPath, sqlSetupPath, odbcSetupPath, sqlVersion, odbcVersion,
onSqlVersionMismatch, dbUtilsSourcePath, unityMsiPath, dsnName,
minFreeSpaceGB, packageDownloadContext (Target [default] or Workstation).
SQL source = ISO or directory containing setup.exe. The ODBC stage requires
Install-OdbcDriver.ps1; it is NOT silently skipped if that child is absent.

In Remote mode installer filesystem sources are probed FROM THE TARGET. This
matches the supplied AppServer implementation. Adapt the source check if another
child instead stages from the workstation. A PSDrive access check cannot prove
that a different child uses the same credential/staging mechanism.
ProGet probe checks HTTPS feed reachability, not package IDs/version existence.
No certificate-validation bypass is used by the probes in this script.

Child SettingsFilePath/SettingsPath names are mapped explicitly. Missing required
child parameters stop preflight. The report contains resolved nonsecret invocation
arguments, not the raw JSON or credential objects. Never put secrets in MSI args.
.EXAMPLE
$cred = Get-Credential -UserName 'ONBASE\prroy' -Message 'Remote credentials'
$r = .\PowerShell\Test-OnBaseReadiness.ps1 -SettingsFilePath .\Settings.json -Credential $cred
$r.Checks | Format-Table Check, Status, Message -Wrap
$r.Passed
#>
[CmdletBinding()]
param(
    [string]$SettingsFilePath = (Join-Path (Split-Path $PSScriptRoot -Parent) 'Settings.json'),
    [pscredential]$Credential,
    [pscredential]$ShareCredential,
    [AllowEmptyCollection()]
    [ValidateSet('Sql','Odbc','Database','IdSAndApi','AppServer','UnityClient','ServiceLocation','Verification')]
    [string[]]$EnabledStages = @('Sql','Odbc','Database','IdSAndApi','AppServer','UnityClient','ServiceLocation','Verification'),
    [hashtable]$Overrides = @{},
    [switch]$UninstallExisting,
    [switch]$UninstallExistingAppServer,
    [switch]$ForceReinstallIdSAndApi
)
$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest
$checks = New-Object 'System.Collections.Generic.List[object]'
$steps = New-Object 'System.Collections.Generic.List[object]'
$sources = New-Object 'System.Collections.Generic.List[object]'
function Add-Check([string]$Name, [string]$Status, [string]$Message) {
    $checks.Add([pscustomobject]@{Check=$Name; Status=$Status; Message=$Message}) | Out-Null
}
function Get-Value($Object, [string[]]$Path, $Default = $null) {
    $v = $Object
    foreach ($part in $Path) {
        if ($null -eq $v -or -not $v.PSObject.Properties[$part]) { return $Default }
        $v = $v.$part
    }
    if ($null -eq $v -or ($v -is [string] -and [string]::IsNullOrWhiteSpace($v))) { return $Default }
    return $v
}
function Resolve-Value([string]$Name, [string[]]$Path, $Default = $null) {
    if ($Overrides.ContainsKey($Name)) { return $Overrides[$Name] }
    return (Get-Value $settings $Path $Default)
}
function Find-Parameter($Command, [string[]]$Names) {
    foreach ($name in $Names) {
        if ($Command.Parameters.ContainsKey($name)) { return $name }
        foreach ($p in $Command.Parameters.Values) { if ($p.Aliases -contains $name) { return $p.Name } }
    }
    return $null
}
function Add-Argument($Step, $Command, [string[]]$Names, $Value, [bool]$Required = $false) {
    $p = Find-Parameter $Command $Names
    if ($p) { $Step.Arguments[$p] = $Value }
    elseif ($Required) { Add-Check "Contract/$($Step.Name)" 'Fail' ("Missing supported parameter: " + ($Names -join ' or ')) }
    return $p
}
function Add-Source([string]$Stage, [string]$Path, [string]$Kind) {
    if ([string]::IsNullOrWhiteSpace($Path)) { Add-Check "Source/$Stage" 'Fail' 'Required source path is missing from settings.'; return }
    if ($Path -match '^https?://') { Add-Check "Source/$Stage" 'Fail' 'A filesystem path is required here, not an HTTP URL.'; return }
    $sources.Add([pscustomobject]@{Stage=$Stage; Path=$Path; Kind=$Kind}) | Out-Null
}
function New-Stage([string]$Name, [string]$File) {
    $path = Join-Path $PSScriptRoot $File
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) { Add-Check "Script/$Name" 'Fail' "Required child script not found: $File"; return $null }
    $tokens = $null; $errors = $null
    $null = [System.Management.Automation.Language.Parser]::ParseFile($path, [ref]$tokens, [ref]$errors)
    if ($errors) { Add-Check "Script/$Name" 'Fail' "Syntax errors in $File; inspect it with the PowerShell parser."; return $null }
    try { $command = Get-Command -Name $path -ErrorAction Stop }
    catch { Add-Check "Script/$Name" 'Fail' "Cannot inspect parameters for $File."; return $null }
    $step = [pscustomobject]@{
        Name=$Name; ScriptPath=$path; Arguments=@{}; VerifyOnly=$false
        CredentialParameter=$null; ShareCredentialParameter=$null
        HsiCredentialParameter=$null; SaCredentialParameter=$null
        UsesLegacyHsiSettings=$false
    }
    $null = Add-Argument $step $command @('DeploymentMode') $mode $true
    if ($mode -eq 'Remote') {
        $null = Add-Argument $step $command @('TargetMachine') $target $true
        $step.CredentialParameter = Find-Parameter $command @('Credential')
        if (-not $step.CredentialParameter) { Add-Check "Contract/$Name" 'Fail' 'Remote child must accept Credential.' }
        $null = Add-Argument $step $command @('ConfigurationName') $endpoint $true
    }
    $null = Add-Argument $step $command @('SettingsFilePath','SettingsPath') $settingsFull
    $step.ShareCredentialParameter = Find-Parameter $command @('ShareCredential')
    $step.HsiCredentialParameter = Find-Parameter $command @('HsiCredential')
    Add-Check "Script/$Name" 'Pass' "$File parsed; parameter metadata read without invoking the installer."
    return [pscustomobject]@{Step=$step; Command=$command}
}
if ($PSVersionTable.PSEdition -ne 'Desktop' -or -not [Environment]::Is64BitProcess) {
    Add-Check 'PowerShellHost' 'Fail' 'Run in 64-bit Windows PowerShell 5.1, not PowerShell 7 or a 32-bit host.'
    return [pscustomobject]@{Passed=$false; Checks=$checks.ToArray(); Context=$null; Steps=@()}
}
try {
    $settingsFull = (Resolve-Path -LiteralPath $SettingsFilePath -ErrorAction Stop).ProviderPath
    $settings = Get-Content -LiteralPath $settingsFull -Raw | ConvertFrom-Json -ErrorAction Stop
    if ($null -eq $settings -or $settings -is [array] -or $settings -is [string] -or $settings -is [ValueType]) { throw 'JSON root must be an object.' }
} catch {
    Add-Check 'Settings' 'Fail' 'Settings file is missing, unreadable, or invalid JSON. Supply a JSON object.'
    return [pscustomobject]@{Passed=$false; Checks=$checks.ToArray(); Context=$null; Steps=@()}
}
Add-Check 'Settings' 'Pass' 'Settings loaded; the report excludes the raw JSON and credential objects.'
$mode = [string](Resolve-Value 'DeploymentMode' @('deploymentMode') 'Remote')
$target = [string](Resolve-Value 'TargetMachine' @('machineName') '')
if (-not $target) { $target = [string](Get-Value $settings @('TargetServer') '') }
if ($mode -eq 'Local' -and -not $target) { $target = $env:COMPUTERNAME }
$endpoint = [string](Resolve-Value 'ConfigurationName' @('configurationName') 'UnityInstall')
$sqlVersion = [string](Resolve-Value 'SqlVersion' @('sqlVersion') '2022')
$odbcVersion = [string](Resolve-Value 'OdbcVersion' @('odbcVersion') '18')
$mismatch = [string](Resolve-Value 'OnSqlVersionMismatch' @('onSqlVersionMismatch') 'SideBySide')
$database = [string](Resolve-Value 'DatabaseName' @('datasourceName') 'OnBase')
$dsn = [string](Resolve-Value 'DsnName' @('dsnName') $database)
$shortTarget = ($target -split '\.')[0]
$defaultInstance = if ($sqlVersion -eq '2019') { 'SQL2K19' } else { 'SQL2K22' }
$sqlInstance = [string](Resolve-Value 'SqlInstance' @('sqlInstance') "$shortTarget\$defaultInstance")
$parts = $sqlInstance -split '\\', 2
if ($parts.Count -eq 2) { $sqlHost = $parts[0]; $instanceName = $parts[1] }
else { $sqlHost = $parts[0]; $instanceName = 'MSSQLSERVER' }
$sqlService = if ($instanceName -eq 'MSSQLSERVER') { 'MSSQLSERVER' } else { 'MSSQL$' + $instanceName }
$driver = "ODBC Driver $odbcVersion for SQL Server"
$build = [string](Resolve-Value 'BuildPath' @('buildPath') '')
$sqlMedia = [string](Resolve-Value 'SqlSetupPath' @('sqlSetupPath') '')
$odbcMedia = [string](Resolve-Value 'OdbcSetupPath' @('odbcSetupPath') '')
$dbUtils = [string](Resolve-Value 'DbUtilsSourcePath' @('dbUtilsSourcePath') '')
$unityMsi = [string](Resolve-Value 'UnityMsiPath' @('unityMsiPath') '')
if ($build -and -not $dbUtils) { $dbUtils = $build.TrimEnd('\') + '\utilities\FoundationDBUtils' }
if ($build -and -not $unityMsi) { $unityMsi = $build.TrimEnd('\') + '\Install\Unity Client\Hyland Unity Client.msi' }
$appMsi = [string](Get-Value $settings @('appServer','sourceMsiPath') '')
if (-not $appMsi) { $appMsi = [string](Get-Value $settings @('Installers','AppServerMsiPath') '') }
$idpApp = [string](Resolve-Value 'IdSAppName' @('idp','appName') 'identityservice')
$apiApp = [string](Resolve-Value 'ApiAppName' @('apiServer','appName') 'apiserver')
$thumb = [string](Resolve-Value 'CertificateThumbprint' @('thumbprint') '')
$extra = [string](Resolve-Value 'ExtraMsiArgs' @('unityExtraMsiArgs') '')
$packageContext = [string](Get-Value $settings @('packageDownloadContext') 'Target')
$minFreeGB = 10.0
try { $minFreeGB = [double](Get-Value $settings @('minFreeSpaceGB') 10); if ($minFreeGB -lt 0) { throw 'invalid' } }
catch { Add-Check 'Settings/Space' 'Fail' 'minFreeSpaceGB must be a non-negative number.' }
if ($mode -notin @('Remote','Local')) { Add-Check 'Settings/Mode' 'Fail' 'deploymentMode must be Remote or Local.' }
if ($sqlVersion -notin @('2019','2022')) { Add-Check 'Settings/SQL' 'Fail' 'sqlVersion must be 2019 or 2022.' }
if ($odbcVersion -notin @('17','18')) { Add-Check 'Settings/ODBC' 'Fail' 'odbcVersion must be 17 or 18.' }
if ($mismatch -notin @('SideBySide','Skip','Fail')) { Add-Check 'Settings/Mismatch' 'Fail' 'Invalid onSqlVersionMismatch.' }
if ($packageContext -notin @('Target','Workstation')) { Add-Check 'Settings/DownloadContext' 'Fail' 'packageDownloadContext must be Target or Workstation.' }
if (-not $target -or -not $endpoint) { Add-Check 'Settings/Target' 'Fail' 'machineName and configurationName must be set.' }
if ($mode -eq 'Local' -and $shortTarget -ine $env:COMPUTERNAME) { Add-Check 'Settings/LocalTarget' 'Fail' 'Local mode target must name the current computer.' }
if ($sqlHost -notin @('.','localhost') -and ($sqlHost -split '\.')[0] -ine $shortTarget) { Add-Check 'Settings/SQLHost' 'Fail' 'This plan supports SQL on the target VM only; SQL host differs from machineName.' }
# Never return secret-bearing MSI arguments in the plan.
if ($extra -match '(?i)PASSWORD\s*=|PWD\s*=') { Add-Check 'Settings/UnityArgs' 'Fail' 'Do not put passwords in MSI arguments.'; $extra = '' }
$order = @('Sql','Odbc','Database','IdSAndApi','AppServer','UnityClient','ServiceLocation','Verification')
foreach ($name in $order) {
    if ($EnabledStages -notcontains $name) { continue }
    if ($name -eq 'Sql' -and -not $sqlMedia) {
        $steps.Add([pscustomobject]@{Name='Sql'; ScriptPath=$null; Arguments=@{}; VerifyOnly=$true; CredentialParameter=$null; ShareCredentialParameter=$null; HsiCredentialParameter=$null; SaCredentialParameter=$null; UsesLegacyHsiSettings=$false}) | Out-Null
        Add-Check 'SQL/Mode' 'Warn' 'No sqlSetupPath: existing SQL must pass validation; SQL will not be installed.'
        continue
    }
    $file = switch ($name) {
        'Sql' {'Install-SqlServer.ps1'}; 'Odbc' {'Install-OdbcDriver.ps1'}
        'Database' {'Initialize-OnBaseDatabase.ps1'}; 'IdSAndApi' {'Install-HylandIdSAndApi.ps1'}
        'AppServer' {'Install-OnBaseApplicationServer.ps1'}; 'UnityClient' {'Install-Unity.ps1'}
        'ServiceLocation' {'Configure-ServiceLocation.ps1'}; 'Verification' {'Verify-Installation.ps1'}
    }
    $entry = New-Stage $name $file
    if ($null -eq $entry) { continue }
    $s = $entry.Step; $c = $entry.Command
    switch ($name) {
        'Sql' {
            Add-Source $name $sqlMedia 'SqlMedia'
            $null = Add-Argument $s $c @('SqlSetupPath') $sqlMedia $true
            $null = Add-Argument $s $c @('SqlVersion') $sqlVersion $true
            $p = Find-Parameter $c @('SqlInstanceName','InstanceName')
            if ($p) { $s.Arguments[$p] = $instanceName }
            else { $null = Add-Argument $s $c @('SqlInstance') $sqlInstance $true }
            $null = Add-Argument $s $c @('OnSqlVersionMismatch','OnVersionMismatch') $mismatch $true
            $s.SaCredentialParameter = Find-Parameter $c @('SaCredential','SqlSaCredential','SqlAdminCredential')
            if (-not $s.SaCredentialParameter) { Add-Check 'Contract/Sql' 'Fail' 'SQL child must accept a supported sa credential parameter.' }
        }
        'Odbc' {
            $null = Add-Argument $s $c @('OdbcVersion') $odbcVersion $true
            if ($odbcMedia) {
                Add-Source $name $odbcMedia 'Any'
                $null = Add-Argument $s $c @('OdbcSetupPath','OdbcDriverMsiPath') $odbcMedia $true
            } else { Add-Check 'Source/Odbc' 'Warn' 'No odbcSetupPath. Child must validate an existing driver or obtain media; post-check enforces the driver.' }
        }
        'Database' {
            Add-Source $name $dbUtils 'Directory'
            $null = Add-Argument $s $c @('DbUtilsSourcePath') $dbUtils $true
            $null = Add-Argument $s $c @('SqlInstance') $sqlInstance $true
            $null = Add-Argument $s $c @('DatabaseName') $database $true
            $null = Add-Argument $s $c @('OdbcDriver') $driver $true
            $null = Add-Argument $s $c @('DsnName') $dsn $true
            $null = Add-Argument $s $c @('Use32Bit') $true $true
            if (-not $s.HsiCredentialParameter) { Add-Check 'Contract/Database' 'Fail' 'Database child must accept HsiCredential.' }
        }
        'IdSAndApi' {
            $null = Add-Argument $s $c @('SettingsFilePath','SettingsPath') $settingsFull $true
            $null = Add-Argument $s $c @('SqlInstance') $sqlInstance $true
            $null = Add-Argument $s $c @('DatabaseName') $database $true
            $null = Add-Argument $s $c @('IdSAppName','IdSApp') $idpApp $true
            $null = Add-Argument $s $c @('ApiAppName','ApiApp') $apiApp $true
            if (-not $s.HsiCredentialParameter) { Add-Check 'Contract/IdSAndApi' 'Fail' 'IdS/API child must accept HsiCredential.' }
            if (-not (Get-Value $settings @('proGetFeedUrl'))) { Add-Check 'Settings/ProGet' 'Fail' 'proGetFeedUrl is required for this package-based deployment.' }
            foreach ($section in @('idp','apiServer')) {
                foreach ($key in @('version','packageId')) {
                    if (-not (Get-Value $settings @($section,$key))) { Add-Check "Settings/$section" 'Fail' "Missing $section.$key." }
                }
            }
            if ($thumb) { $null = Add-Argument $s $c @('CertificateThumbprint','CertThumb') $thumb $true }
            if ($ForceReinstallIdSAndApi) { $null = Add-Argument $s $c @('ForceReinstall') $true $true }
        }
        'AppServer' {
            Add-Source $name $appMsi 'File'
            $null = Add-Argument $s $c @('SettingsPath','SettingsFilePath') $settingsFull $true
            $null = Add-Argument $s $c @('SourceMsiPath') $appMsi $true
            $timeout = 15
            try { $timeout = [int](Get-Value $settings @('appServer','timeoutMinutes') 15); if ($timeout -le 0) { throw 'invalid' } }
            catch { Add-Check 'Settings/AppServerTimeout' 'Fail' 'appServer.timeoutMinutes must be a positive integer.' }
            $null = Add-Argument $s $c @('TimeoutMinutes') $timeout $true
            $appExtra = [string](Get-Value $settings @('appServer','extraMsiArgs') 'ALLUSERS=1')
            if ($appExtra -match '(?i)PASSWORD\s*=|PWD\s*=') { Add-Check 'Settings/AppServerArgs' 'Fail' 'Do not put passwords in MSI arguments.'; $appExtra = '' }
            $null = Add-Argument $s $c @('ExtraMsiArgs') $appExtra $true
            if ((Get-Value $settings @('appServer','appName') 'AppServer64') -ne 'AppServer64') { Add-Check 'Settings/AppServerName' 'Fail' 'Supplied AppServer implementation supports AppServer64 only.' }
            if ($UninstallExistingAppServer) { $null = Add-Argument $s $c @('UninstallExisting') $true $true }
            if (-not $s.HsiCredentialParameter) {
                $s.UsesLegacyHsiSettings = $true
                if (-not (Get-Value $settings @('hsiPassword'))) { Add-Check 'Contract/AppServerHSI' 'Fail' 'AppServer child lacks HsiCredential and hsiPassword is blank. Add secure HsiCredential support to that child; hardcoded fallback must not be used.' }
                else { Add-Check 'Contract/AppServerHSI' 'Warn' 'Legacy AppServer reads a plaintext JSON password. Migrate to HsiCredential; runtime consistency is enforced.' }
            }
            if ((Get-Value $settings @('sqlInstance') '') -ine $sqlInstance -or (Get-Value $settings @('datasourceName') 'OnBase') -ine $database -or (Get-Value $settings @('idp','appName') 'identityservice') -ine $idpApp) {
                Add-Check 'Contract/AppServerSettings' 'Fail' 'AppServer reads SQL/database/IdS from JSON. Update JSON instead of conflicting CLI overrides.'
            }
        }
        'UnityClient' {
            Add-Source $name $unityMsi 'File'
            $null = Add-Argument $s $c @('SourceMsiPath') $unityMsi $true
            if ($UninstallExisting) { $null = Add-Argument $s $c @('UninstallExisting') $true $true }
            if ($extra) { $null = Add-Argument $s $c @('ExtraMsiArgs') $extra $true }
        }
        'ServiceLocation' {
            $null = Add-Argument $s $c @('DataSource','DatabaseName') $database $true
            $null = Add-Argument $s $c @('SqlInstance') $sqlInstance
            $null = Add-Argument $s $c @('DsnName') $dsn
            if (-not (Find-Parameter $c @('SettingsFilePath','SettingsPath'))) { Add-Check 'Contract/ServiceLocation' 'Fail' 'Service-location child must accept the settings file for environment URLs; do not rely on hidden URL defaults.' }
        }
        'Verification' {
            $null = Add-Argument $s $c @('SqlInstance') $sqlInstance $true
            $null = Add-Argument $s $c @('OdbcDriver') $driver $true
            $null = Add-Argument $s $c @('DsnName','DataSource') $dsn $true
            if (-not $s.HsiCredentialParameter) { Add-Check 'Contract/VerificationHSI' 'Warn' 'Verifier has no HsiCredential parameter; inspect login/DSN authentication coverage separately.' }
        }
    }
    $uncSources = @($sources.ToArray() | Where-Object { $_.Stage -eq $name -and $_.Path -like '\\*' })
    if ($ShareCredential -and $uncSources.Count -gt 0 -and -not $s.ShareCredentialParameter) { Add-Check "Contract/$name/Share" 'Fail' 'ShareCredential supplied, but this UNC-consuming child cannot accept it. Align its access/staging behavior first.' }
    $steps.Add($s) | Out-Null
}
$session = $null
$inspectTarget = {
    $ErrorActionPreference = 'Stop'
    $principal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
    [pscustomobject]@{
        ComputerName=$env:COMPUTERNAME
        IsAdministrator=$principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
        FreeGB=[math]::Round((Get-PSDrive -Name C -ErrorAction Stop).Free / 1GB, 2)
        Desktop51=($PSVersionTable.PSEdition -eq 'Desktop' -and $PSVersionTable.PSVersion -ge [version]'5.1' -and [Environment]::Is64BitProcess)
    }
}
$sourceProbe = {
    param($Path, $Kind, $ShareCred)
    $ErrorActionPreference = 'Stop'
    $drive = $null; $probePath = $Path
    try {
        if ($ShareCred -and $Path -match '^(\\\\[^\\]+\\[^\\]+)(.*)$') {
            $root = $Matches[1]; $relative = $Matches[2].TrimStart('\')
            $drive = 'OBR' + [guid]::NewGuid().ToString('N').Substring(0,8)
            New-PSDrive -Name $drive -PSProvider FileSystem -Root $root -Credential $ShareCred -Scope Local -ErrorAction Stop | Out-Null
            $probePath = if ($relative) { "${drive}:\$relative" } else { "${drive}:\" }
        }
        $item = Get-Item -LiteralPath $probePath -ErrorAction Stop
        if ($Kind -eq 'Directory' -and -not $item.PSIsContainer) { throw 'Directory expected.' }
        if ($Kind -eq 'File' -and $item.PSIsContainer) { throw 'File expected.' }
        if ($Kind -eq 'SqlMedia') {
            if ($item.PSIsContainer) { if (-not (Test-Path -LiteralPath (Join-Path $probePath 'setup.exe') -PathType Leaf)) { throw 'setup.exe missing.' } }
            elseif ($item.Extension -ine '.iso') { throw 'SQL media must be ISO or setup directory.' }
        }
        if ($item.PSIsContainer) { Get-ChildItem -LiteralPath $probePath -ErrorAction Stop | Select-Object -First 1 | Out-Null }
        else { Get-Content -LiteralPath $probePath -Encoding Byte -TotalCount 1 -ErrorAction Stop | Out-Null }
        [pscustomobject]@{Passed=$true; Message='Source can be read in the selected installation context.'}
    } catch { [pscustomobject]@{Passed=$false; Message='Source inaccessible or wrong media type; check path, identity, permissions and credential forwarding.'} }
    finally { if ($drive) { Remove-PSDrive -Name $drive -Force -ErrorAction SilentlyContinue } }
}
$feedProbe = {
    param($Url)
    $ErrorActionPreference = 'Stop'
    $oldProtocol = [Net.ServicePointManager]::SecurityProtocol
    $oldCallback = [Net.ServicePointManager]::ServerCertificateValidationCallback
    try {
        [Net.ServicePointManager]::ServerCertificateValidationCallback = $null
        [Net.ServicePointManager]::SecurityProtocol = $oldProtocol -bor [Net.SecurityProtocolType]::Tls12
        $r = Invoke-WebRequest -Uri $Url -UseBasicParsing -TimeoutSec 20 -MaximumRedirection 3 -ErrorAction Stop
        [pscustomobject]@{Passed=($r.StatusCode -ge 200 -and $r.StatusCode -lt 300)}
    } catch { [pscustomobject]@{Passed=$false} }
    finally { [Net.ServicePointManager]::SecurityProtocol = $oldProtocol; [Net.ServicePointManager]::ServerCertificateValidationCallback = $oldCallback }
}
try {
    if ($mode -eq 'Remote') {
        if (-not $Credential -or -not $target -or -not $endpoint) { throw 'Remote checks require Credential, target and endpoint.' }
        try { $null = [Net.Dns]::GetHostAddresses($target); Add-Check 'DNS' 'Pass' "Resolved $target." }
        catch { Add-Check 'DNS' 'Fail' "Cannot resolve $target."; throw 'DNS failed.' }
        $option = New-PSSessionOption -OpenTimeout 20000 -OperationTimeout 30000
        $session = New-PSSession -ComputerName $target -Credential $Credential -Authentication Negotiate -ConfigurationName $endpoint -SessionOption $option -ErrorAction Stop
        $state = Invoke-Command -Session $session -ScriptBlock $inspectTarget -ErrorAction Stop
        Add-Check 'WinRM/Endpoint' 'Pass' "Authenticated execution succeeded through $endpoint."
    } elseif ($mode -eq 'Local') { $state = & $inspectTarget }
    else { throw 'Invalid deployment mode.' }
    if ($state.ComputerName -ine $shortTarget) { Add-Check 'TargetIdentity' 'Fail' 'Actual computer name differs from the configured target.' }
    else { Add-Check 'TargetIdentity' 'Pass' "Confirmed $($state.ComputerName)." }
    if (-not $state.IsAdministrator) { Add-Check 'ExecutionRights' 'Fail' 'Installer context must be administrative; an endpoint name alone does not grant elevation.' }
    else { Add-Check 'ExecutionRights' 'Pass' 'Administrative execution context confirmed.' }
    if (-not $state.Desktop51) { Add-Check 'TargetPowerShell' 'Fail' 'Target endpoint must be 64-bit Windows PowerShell 5.1.' }
    if ($state.FreeGB -lt $minFreeGB) { Add-Check 'DiskSpace' 'Fail' "C: has $($state.FreeGB) GB; minimum $minFreeGB GB." }
    else { Add-Check 'DiskSpace' 'Pass' "C: free $($state.FreeGB) GB; minimum $minFreeGB GB." }
    foreach ($source in $sources.ToArray()) {
        if ($mode -eq 'Remote') { $probe = Invoke-Command -Session $session -ScriptBlock $sourceProbe -ArgumentList $source.Path,$source.Kind,$ShareCredential -ErrorAction Stop }
        else { $probe = & $sourceProbe $source.Path $source.Kind $ShareCredential }
        $status = if ($probe.Passed) { 'Pass' } else { 'Fail' }
        Add-Check "Source/$($source.Stage)" $status $probe.Message
    }
    $willInstallSql = ($EnabledStages -contains 'Sql' -and -not [string]::IsNullOrWhiteSpace($sqlMedia))
    $needsSql = @($EnabledStages | Where-Object { $_ -in @('Sql','Database','IdSAndApi','AppServer','Verification') }).Count -gt 0
    if ($needsSql -and -not $willInstallSql) {
        $serviceProbe = { param($Name) $s = Get-Service -Name $Name -ErrorAction SilentlyContinue; [bool]($s -and $s.Status -eq 'Running') }
        $up = if ($mode -eq 'Remote') { Invoke-Command -Session $session -ScriptBlock $serviceProbe -ArgumentList $sqlService -ErrorAction Stop } else { & $serviceProbe $sqlService }
        if ($up) { Add-Check 'ExistingSQL' 'Pass' "$sqlService running (not a SQL login/version test)." }
        else { Add-Check 'ExistingSQL' 'Fail' "$sqlService missing or stopped; SQL installation is not scheduled." }
    }
    $needsDsn = @($EnabledStages | Where-Object { $_ -in @('Database','Verification') }).Count -gt 0
    if ($needsDsn -and $EnabledStages -notcontains 'Odbc') {
        $driverProbe = { param($Name) [bool](Get-OdbcDriver -Name $Name -Platform '32-bit' -ErrorAction SilentlyContinue) }
        $present = if ($mode -eq 'Remote') { Invoke-Command -Session $session -ScriptBlock $driverProbe -ArgumentList $driver -ErrorAction Stop } else { & $driverProbe $driver }
        if ($present) { Add-Check 'ExistingODBC' 'Pass' 'Requested 32-bit ODBC driver present.' }
        else { Add-Check 'ExistingODBC' 'Fail' 'Required 32-bit ODBC driver missing and setup stage skipped.' }
    }
    if ($EnabledStages -contains 'AppServer' -and $EnabledStages -notcontains 'IdSAndApi') {
        $iisProbe = { Import-Module WebAdministration -ErrorAction Stop; [bool](Test-Path 'IIS:\Sites\Default Web Site') }
        $iis = if ($mode -eq 'Remote') { Invoke-Command -Session $session -ScriptBlock $iisProbe -ErrorAction Stop } else { & $iisProbe }
        if ($iis) { Add-Check 'ExistingIIS' 'Pass' 'WebAdministration and Default Web Site available.' }
        else { Add-Check 'ExistingIIS' 'Fail' 'Default Web Site is required when IdS/API prerequisite setup is skipped.' }
    }
    if ($EnabledStages -contains 'IdSAndApi') {
        $feed = [string](Get-Value $settings @('proGetFeedUrl') '')
        $uri = $null
        if (-not [uri]::TryCreate($feed, [UriKind]::Absolute, [ref]$uri) -or $uri.Scheme -ne 'https') { Add-Check 'ProGet' 'Fail' 'An absolute HTTPS proGetFeedUrl is required.' }
        else {
            $fp = if ($mode -eq 'Remote' -and $packageContext -eq 'Target') { Invoke-Command -Session $session -ScriptBlock $feedProbe -ArgumentList $feed -ErrorAction Stop } else { & $feedProbe $feed }
            if ($fp.Passed) { Add-Check 'ProGet' 'Pass' "Feed responded in $packageContext context; package IDs/versions NOT validated." }
            else { Add-Check 'ProGet' 'Fail' "Feed probe failed in $packageContext context; check network, TLS, proxy and feed authentication." }
        }
    }
} catch {
    Add-Check 'ExecutionContext' 'Fail' 'Some target checks could not complete. Check credentials, endpoint permissions/allowed commands, and network. Use -Verbose for details.'
    Write-Verbose $_.Exception.Message
} finally { if ($session) { Remove-PSSession -Session $session -ErrorAction SilentlyContinue } }
$context = [pscustomobject]@{
    SettingsFilePath=$settingsFull; DeploymentMode=$mode; TargetMachine=$target; ConfigurationName=$endpoint
    SqlInstance=$sqlInstance; SqlServiceName=$sqlService; SqlVersion=$sqlVersion; OdbcDriver=$driver
    DatabaseName=$database; DsnName=$dsn; IdSAppName=$idpApp; ApiAppName=$apiApp; AppServerAppName='AppServer64'
    EnabledStages=@($order | Where-Object { $EnabledStages -contains $_ })
}
[pscustomobject]@{
    Passed=(@($checks.ToArray() | Where-Object { $_.Status -eq 'Fail' }).Count -eq 0)
    Checks=$checks.ToArray(); Context=$context; Steps=$steps.ToArray()
}
