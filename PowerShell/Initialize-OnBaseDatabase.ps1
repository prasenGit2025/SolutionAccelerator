#Requires -Version 5.1
<#
.SYNOPSIS
    Prepares an OnBase database, schema, and System DSNs using Settings.json.

.DESCRIPTION
    Based on the supplied Initialize-OnBaseDatabase.ps1. Configuration is resolved
    on the launching workstation before local or remote execution.

    Default configuration: ..\Settings.json, relative to this script's folder.
    Precedence: explicitly supplied parameters > JSON properties > script defaults.
    An explicit DbUtilsSourcePath wins over JSON. Otherwise dbUtilsSourcePath is
    used, or derived as <buildPath>\utilities\FoundationDBUtils.

    Supported root JSON properties:
      deploymentMode, machineName, configurationName, remoteUsername,
      sqlInstance, databaseName, dataFilePath, hsiUsername,
      dbUtilsSourcePath, buildPath, tablespaceOption,
      unicode, testSystem, use32Bit, create64BitDsn, timeoutMinutes,
      dsnName (or datasourceName),
      odbcDriver (or odbcDriverName, or odbcVersion such as "18").

    Password properties in JSON are NOT read. Missing Windows/hsi credentials
    are requested on the launching workstation. SqlAdminCredential is optional;
    when absent SQL administration uses the target session's Windows identity.
    An existing hsi login's password is NOT reset.

    Retained database operations:
      1 Pre-check SQL sysadmin, mixed-mode authentication, and ODBC registration.
      2 Set instance-wide Max Degree of Parallelism to 1.
      3 Create the database if absent.
      4 Create hsi if absent; grant securityadmin and database db_owner.
      5 Run FoundationDBUtils create unless hsi.productsold already exists.
      6 Configure a 32-bit System DSN and optionally a 64-bit System DSN.
      7 Report readiness. Exit 0 only if the reported checks pass; otherwise 1.

    These operations change SQL configuration, security, and local System DSNs.
    This is not a read-only script or a comprehensive database health check.
    DBUtils command syntax/requirements are retained from the supplied source;
    compatibility with your FoundationDBUtils build has not been verified here.
    DBUtils receives credentials through its process arguments as in the source.
    Restrict access to the target and its DBUtils output/error logs.

    FoundationDBUtils must be a directory on a local drive or UNC share, not an
    HTTP/HTTPS URL. In Remote mode that path is resolved ON THE TARGET. A successful
    workstation Test-Path does not establish target-session share access.
    ConfigurationName selects an existing remoting endpoint; its name alone does
    not establish share access, elevation, or a particular RunAs identity.

    Use 64-bit PowerShell on Windows. In Remote mode the endpoint must also run a
    64-bit process. A custom DataFilePath assumes SQL storage is on the execution
    target, as in the supplied script; this script does not prepare another host's
    storage or grant SQL service account filesystem permissions.

.EXAMPLE
    & 'C:\Users\PRROY\source\OnBaseProvisioningFramework\PowerShell\Initialize-OnBaseDatabase.ps1'

    Loads repository-root Settings.json and prompts for credentials when absent.

.EXAMPLE
    .\PowerShell\Initialize-OnBaseDatabase.ps1 -SettingsFilePath '.\Settings.json' -Credential $cred -HsiCredential $hsi

    Reuses credential objects already collected by the orchestrator/session.

.EXAMPLE
    .\PowerShell\Initialize-OnBaseDatabase.ps1 -SettingsFilePath '.\Settings.json' -Credential $cred -HsiCredential $hsi -SqlAdminCredential $sa

    Uses a supplied existing SQL administrator credential instead of Windows SQL
    authentication. This script does not set or reset the sa password.
#>
[CmdletBinding()]
param(
    [Alias('SettingsPath')]
    [string]$SettingsFilePath = (Join-Path (Split-Path -Parent $PSScriptRoot) 'Settings.json'),
    [ValidateSet('Local','Remote')]
    [string]$DeploymentMode = 'Local',
    [string]$TargetMachine = '',
    [pscredential]$Credential,
    [string]$ConfigurationName = '',
    [string]$SqlInstance = '',
    [string]$DatabaseName = 'OnBase',
    [string]$DataFilePath = '',
    [pscredential]$SqlAdminCredential,
    [pscredential]$HsiCredential,
    [securestring]$ViewerPassword,
    [string]$DbUtilsSourcePath = '',
    [ValidateSet('single','default')]
    [string]$TablespaceOption = 'single',
    [switch]$Unicode,
    [switch]$TestSystem,
    [switch]$Use32Bit,
    [string]$DsnName = '',
    [string]$OdbcDriver = 'ODBC Driver 18 for SQL Server',
    [switch]$Create64BitDsn,
    [ValidateRange(1,1440)]
    [int]$TimeoutMinutes = 60
)

$ErrorActionPreference = 'Stop'

# Capture explicit arguments before loading configuration. Even -Switch:$false
# and explicitly supplied default values must override JSON.
$explicitParameters = @{}
foreach ($name in $PSBoundParameters.Keys) {
    $explicitParameters[$name] = $true
}

# ---------------------------------------------------------------------------
# Configuration and credential collection ON THE LAUNCHING WORKSTATION
# ---------------------------------------------------------------------------
try {
    if ([string]::IsNullOrWhiteSpace($SettingsFilePath)) {
        throw 'SettingsFilePath cannot be empty. Supply a Settings.json path.'
    }
    if (-not (Test-Path -LiteralPath $SettingsFilePath -PathType Leaf)) {
        throw "Settings file not found: $SettingsFilePath. Put Settings.json in the repository root or pass -SettingsFilePath."
    }
    $resolvedSettings = (Resolve-Path -LiteralPath $SettingsFilePath).ProviderPath
    $settings = Get-Content -LiteralPath $resolvedSettings -Raw | ConvertFrom-Json -ErrorAction Stop
    if ($null -eq $settings -or $settings -isnot [pscustomobject]) {
        throw 'Settings.json must contain a JSON object at its root.'
    }

    function Get-ConfigurationValue {
        param([string[]]$Names)
        foreach ($settingName in $Names) {
            $property = $settings.PSObject.Properties[$settingName]
            if ($null -ne $property -and $null -ne $property.Value) {
                if ($property.Value -is [string] -and [string]::IsNullOrWhiteSpace($property.Value)) {
                    continue
                }
                return $property.Value
            }
        }
        return $null
    }

    # Mandatory parameter attributes were removed from SqlInstance and
    # DbUtilsSourcePath so JSON can populate them BEFORE validation/prompts.
    $stringSettings = @{
        DeploymentMode    = @('deploymentMode')
        TargetMachine     = @('machineName')
        ConfigurationName = @('configurationName')
        SqlInstance       = @('sqlInstance')
        DatabaseName      = @('databaseName')
        DataFilePath      = @('dataFilePath')
        DbUtilsSourcePath = @('dbUtilsSourcePath')
        TablespaceOption  = @('tablespaceOption')
        DsnName           = @('dsnName','datasourceName')
        OdbcDriver        = @('odbcDriver','odbcDriverName')
    }
    foreach ($parameterName in $stringSettings.Keys) {
        if (-not $explicitParameters.ContainsKey($parameterName)) {
            $value = Get-ConfigurationValue -Names $stringSettings[$parameterName]
            if ($null -ne $value) {
                if ($value -isnot [string]) {
                    throw "The JSON property for $parameterName must be a string."
                }
                Set-Variable -Name $parameterName -Value $value -Scope Script
            }
        }
    }

    # A full driver name is preferred; odbcVersion is an optional shorthand.
    if (-not $explicitParameters.ContainsKey('OdbcDriver')) {
        $configuredDriver = Get-ConfigurationValue -Names @('odbcDriver','odbcDriverName')
        if ($null -eq $configuredDriver) {
            $odbcVersion = Get-ConfigurationValue -Names @('odbcVersion')
            if ($null -ne $odbcVersion) {
                if ([string]$odbcVersion -notmatch '^\d+$') {
                    throw 'odbcVersion must be a major version such as 18, or specify odbcDriverName instead.'
                }
                $OdbcDriver = 'ODBC Driver {0} for SQL Server' -f $odbcVersion
            }
        }
    }

    if (-not $explicitParameters.ContainsKey('DbUtilsSourcePath') -and
        [string]::IsNullOrWhiteSpace($DbUtilsSourcePath)) {
        $buildPath = Get-ConfigurationValue -Names @('buildPath')
        if ($null -ne $buildPath) {
            if ($buildPath -isnot [string]) { throw 'buildPath must be a string.' }
            if ($buildPath -match '^https?://') {
                throw 'buildPath must be a filesystem/UNC directory, not an HTTP/HTTPS URL.'
            }
            $DbUtilsSourcePath = Join-Path -Path $buildPath -ChildPath 'utilities\FoundationDBUtils'
        }
    }

    # Require actual JSON booleans: the string "false" must NOT become $true.
    $booleanSettings = @{
        Unicode        = 'unicode'
        TestSystem     = 'testSystem'
        Use32Bit       = 'use32Bit'
        Create64BitDsn = 'create64BitDsn'
    }
    foreach ($parameterName in $booleanSettings.Keys) {
        if (-not $explicitParameters.ContainsKey($parameterName)) {
            $value = Get-ConfigurationValue -Names @($booleanSettings[$parameterName])
            if ($null -ne $value) {
                if ($value -isnot [bool]) {
                    throw "JSON property '$($booleanSettings[$parameterName])' must be true or false, without quotes."
                }
                Set-Variable -Name $parameterName -Value $value -Scope Script
            }
        }
    }

    if (-not $explicitParameters.ContainsKey('TimeoutMinutes')) {
        $value = Get-ConfigurationValue -Names @('timeoutMinutes')
        if ($null -ne $value) {
            $parsedTimeout = 0
            if (-not [int]::TryParse([string]$value, [ref]$parsedTimeout) -or
                $parsedTimeout -lt 1 -or $parsedTimeout -gt 1440) {
                throw 'timeoutMinutes must be a whole number from 1 to 1440.'
            }
            $TimeoutMinutes = $parsedTimeout
        }
    }

    if ($DeploymentMode -notin @('Local','Remote')) {
        throw 'deploymentMode must be Local or Remote.'
    }
    if ($TablespaceOption -notin @('single','default')) {
        throw 'tablespaceOption must be single or default.'
    }
    foreach ($requiredName in @('SqlInstance','DatabaseName','DbUtilsSourcePath')) {
        if ([string]::IsNullOrWhiteSpace([string](Get-Variable -Name $requiredName -ValueOnly))) {
            throw "Missing required configuration: $requiredName. Supply it in Settings.json or as an explicit parameter."
        }
    }
    if ($DbUtilsSourcePath -match '^https?://') {
        throw 'dbUtilsSourcePath must be a filesystem/UNC directory, not an HTTP/HTTPS URL.'
    }
    if (-not [IO.Path]::IsPathRooted($DbUtilsSourcePath)) {
        throw 'dbUtilsSourcePath (or buildPath) must resolve to an absolute local or UNC directory.'
    }
    if ($DataFilePath -and -not [IO.Path]::IsPathRooted($DataFilePath)) {
        throw 'dataFilePath must be an absolute path on the SQL/execution target.'
    }
    if ([string]::IsNullOrWhiteSpace($DsnName)) { $DsnName = $DatabaseName }
    if ([string]::IsNullOrWhiteSpace($OdbcDriver)) { $OdbcDriver = 'ODBC Driver 18 for SQL Server' }
    if ($DsnName -match '[\\/]') { throw 'DsnName must not contain a slash or backslash.' }

    if (-not [Environment]::Is64BitProcess) {
        throw 'Run this script from 64-bit PowerShell on Windows.'
    }

    Write-Host "Settings: $resolvedSettings"
    Write-Host "Mode: $DeploymentMode; target: $TargetMachine"
    Write-Host "SQL: $SqlInstance; database: $DatabaseName; DSN: $DsnName"
    Write-Host "FoundationDBUtils: $DbUtilsSourcePath"
    Write-Host "ODBC: $OdbcDriver; optional 64-bit DSN: $([bool]$Create64BitDsn)"

    if ($DeploymentMode -eq 'Remote') {
        if ([string]::IsNullOrWhiteSpace($TargetMachine)) {
            throw 'machineName / TargetMachine is required for Remote mode.'
        }
        if ($TargetMachine -match '[\[\]\(\)\\/\s]|^https?:') {
            throw "TargetMachine '$TargetMachine' is not a plain hostname/IP address. Remove URLs, paths, whitespace, or link formatting."
        }
        if (-not $Credential) {
            $remoteUsername = Get-ConfigurationValue -Names @('remoteUsername')
            $prompt = @{ Message = "Remote provisioning credentials for $TargetMachine" }
            if ($remoteUsername) { $prompt.UserName = [string]$remoteUsername }
            $Credential = Get-Credential @prompt
            if (-not $Credential) { throw 'Remote credential collection was cancelled.' }
        }
        if (-not $ConfigurationName -and $DbUtilsSourcePath.StartsWith('\\')) {
            Write-Warning 'No ConfigurationName: the target session may not be able to access the UNC share. Configure an authorized access method.'
        }
    }
    else {
        $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
        $principal = New-Object Security.Principal.WindowsPrincipal($identity)
        if (-not $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
            throw 'Run local mode from an elevated PowerShell. System DSNs require administrative rights.'
        }
    }

    if (-not $HsiCredential) {
        $hsiUsername = Get-ConfigurationValue -Names @('hsiUsername')
        if (-not $hsiUsername) { $hsiUsername = 'hsi' }
        if ($hsiUsername -ine 'hsi') { throw "hsiUsername must be 'hsi'." }
        $HsiCredential = Get-Credential -UserName 'hsi' -Message 'OnBase hsi SQL password (existing password, or the new password if hsi will be created)'
        if (-not $HsiCredential) { throw 'hsi credential collection was cancelled.' }
    }
    if ($HsiCredential.UserName -ine 'hsi') { throw "HsiCredential user name must be 'hsi'." }
    if ($HsiCredential.Password.Length -eq 0) { throw 'The hsi password cannot be empty.' }
    if (-not $ViewerPassword) { $ViewerPassword = $HsiCredential.Password }
    $ViewerCredential = [pscredential]::new('viewer', $ViewerPassword)

    # Retained DBUtils command-line/connection-string restriction from the source.
    foreach ($c in @($HsiCredential, $ViewerCredential, $SqlAdminCredential)) {
        if ($c -and $c.GetNetworkCredential().Password -match '[";]') {
            throw "The password for '$($c.UserName)' contains a double quote or semicolon, which this DBUtils invocation does not support."
        }
    }
}
catch {
    Write-Host "Database input/configuration failed: $($_.Exception.Message)"
    exit 1
}

# ---------------------------------------------------------------------------
# Runs ON THE TARGET MACHINE. Configuration is already resolved above.
# Core SQL/schema/DSN behavior is retained from the supplied source.
# ---------------------------------------------------------------------------
$dbReadyScript = {
    param($sqlInstance, $db, $dataPath, $sqlAdmin, $hsiCred, $viewerCred,
          $dbUtilsSrc, $tsOption, $unicode, $testSystem, $use32,
          $dsn, $driver, $want64, $timeoutMinutes)

    if ([string]::IsNullOrWhiteSpace($dsn)) { $dsn = if (![string]::IsNullOrWhiteSpace($db)) { $db } else { 'OnBase' } }
    if ([string]::IsNullOrWhiteSpace($driver)) { $driver = 'ODBC Driver 18 for SQL Server' }

    $ErrorActionPreference = 'Stop'
    if (-not [Environment]::Is64BitProcess) {
        throw 'The execution target/remote endpoint must run 64-bit PowerShell for these registry paths and DSN checks.'
    }
    $pc       = $env:COMPUTERNAME
    $hsiPw    = $hsiCred.GetNetworkCredential().Password
    $viewerPw = $viewerCred.GetNetworkCredential().Password
    $hsiPwSql = $hsiPw.Replace("'", "''")
    $dbSql    = $db.Replace(']', ']]')
    $dbLit    = $db.Replace("'", "''")
    New-Item -ItemType Directory -Path 'C:\Temp' -Force | Out-Null

    if ($sqlAdmin) {
        $adminCs = 'Server={0};Database=master;User ID={1};Password={2};TrustServerCertificate=True' -f $sqlInstance, $sqlAdmin.UserName, $sqlAdmin.GetNetworkCredential().Password
    }
    else {
        $adminCs = 'Server={0};Database=master;Integrated Security=SSPI;TrustServerCertificate=True' -f $sqlInstance
    }
    $hsiCs = 'Server={0};Database={1};User ID=hsi;Password={2};TrustServerCertificate=True' -f $sqlInstance, $db, $hsiPw

    function Invoke-Sql([string]$sql, [string]$database = 'master', [string]$cs = $adminCs) {
        $cn = New-Object System.Data.SqlClient.SqlConnection $cs
        $cn.Open()
        try {
            if ($database -ne 'master') { $cn.ChangeDatabase($database) }
            $cmd = $cn.CreateCommand()
            $cmd.CommandText = $sql
            $cmd.CommandTimeout = 600
            $dt = New-Object System.Data.DataTable
            $dt.Load($cmd.ExecuteReader())
            , $dt
        }
        finally { $cn.Close() }
    }
    function Get-Scalar([string]$sql, [string]$database = 'master', [string]$cs = $adminCs) {
        $t = Invoke-Sql $sql $database $cs
        if ($t.Rows.Count) { $t.Rows[0][0] } else { $null }
    }

    $schemaSql = "SELECT CASE WHEN OBJECT_ID(N'hsi.productsold', N'U') IS NULL THEN 0 ELSE 1 END"
    $dbOwnerSql = "SELECT COUNT(*) FROM sys.database_role_members rm JOIN sys.database_principals r ON rm.role_principal_id = r.principal_id JOIN sys.database_principals m ON rm.member_principal_id = m.principal_id WHERE r.name = N'db_owner' AND m.name = N'hsi'"
    $platforms = @('32-bit')
    if ($want64) { $platforms += '64-bit' }

    # --- 1. Pre-checks ---
    Write-Host "[$pc] 1/7 Pre-checks..."
    $info = (Invoke-Sql "SELECT CAST(SERVERPROPERTY('ProductVersion') AS nvarchar(50)) AS Version, IS_SRVROLEMEMBER('sysadmin') AS IsSysAdmin, CAST(SERVERPROPERTY('IsIntegratedSecurityOnly') AS int) AS WinAuthOnly, SUSER_SNAME() AS LoginName").Rows[0]
    Write-Host "[$pc]     SQL $($info.Version), connected as $($info.LoginName)"
    if ($info.IsSysAdmin -ne 1) { throw "$($info.LoginName) is not sysadmin on $sqlInstance. Add it to sysadmin or pass -SqlAdminCredential." }
    if ($info.WinAuthOnly -eq 1) { throw "$sqlInstance allows Windows authentication only. hsi needs SQL authentication: switch to mixed mode and restart the SQL service." }
    foreach ($plat in $platforms) {
        $drivers = @(Get-OdbcDriver -Platform $plat | Select-Object -ExpandProperty Name)
        if ($drivers -notcontains $driver) { throw "$plat ODBC driver '$driver' is not installed. Installed: $($drivers -join ', ')." }
    }

    # --- 2. Instance-wide Max Degree of Parallelism = 1 ---
    Write-Host "[$pc] 2/7 Max Degree of Parallelism..."
    $maxdop = [int](Get-Scalar "SELECT CAST(value_in_use AS int) FROM sys.configurations WHERE name = 'max degree of parallelism'")
    if ($maxdop -ne 1) {
        Invoke-Sql "EXEC sp_configure 'show advanced options', 1; RECONFIGURE; EXEC sp_configure 'max degree of parallelism', 1; RECONFIGURE;" | Out-Null
        Write-Host "[$pc]     Changed $maxdop -> 1"
    }
    else { Write-Host "[$pc]     Already 1" }

    # --- 3. Database ---
    Write-Host "[$pc] 3/7 Database [$db]..."
    if ([int](Get-Scalar "SELECT COUNT(*) FROM sys.databases WHERE name = N'$dbLit'") -eq 0) {
        if ($dataPath) {
            New-Item -ItemType Directory -Path $dataPath -Force | Out-Null
            $mdf = (Join-Path $dataPath ($db + '.mdf')).Replace("'", "''")
            $ldf = (Join-Path $dataPath ($db + '_log.ldf')).Replace("'", "''")
            Invoke-Sql "CREATE DATABASE [$dbSql] ON PRIMARY (NAME = N'$dbLit', FILENAME = N'$mdf') LOG ON (NAME = N'${dbLit}_log', FILENAME = N'$ldf');" | Out-Null
        }
        else { Invoke-Sql "CREATE DATABASE [$dbSql];" | Out-Null }
        Write-Host "[$pc]     Created"
    }
    else { Write-Host "[$pc]     Already exists" }

    # --- 4. hsi login, securityadmin, db_owner ---
    Write-Host "[$pc] 4/7 hsi login..."
    if ([int](Get-Scalar "SELECT COUNT(*) FROM sys.server_principals WHERE name = N'hsi'") -eq 0) {
        Invoke-Sql "CREATE LOGIN [hsi] WITH PASSWORD = N'$hsiPwSql', DEFAULT_DATABASE = [$dbSql], CHECK_POLICY = OFF, CHECK_EXPIRATION = OFF;" | Out-Null
        Write-Host "[$pc]     Login created"
    }
    else { Write-Host "[$pc]     Login already exists (password not changed)" }

    if ([int](Get-Scalar "SELECT ISNULL(IS_SRVROLEMEMBER('securityadmin', 'hsi'), 0)") -ne 1) {
        Invoke-Sql "ALTER SERVER ROLE [securityadmin] ADD MEMBER [hsi];" | Out-Null
        Write-Host "[$pc]     Added to securityadmin"
    }
    if ([int](Get-Scalar "SELECT COUNT(*) FROM sys.database_principals WHERE name = N'hsi'" $db) -eq 0) {
        Invoke-Sql "CREATE USER [hsi] FOR LOGIN [hsi];" $db | Out-Null
        Write-Host "[$pc]     Database user created"
    }
    if ([int](Get-Scalar $dbOwnerSql $db) -eq 0) {
        Invoke-Sql "ALTER ROLE [db_owner] ADD MEMBER [hsi];" $db | Out-Null
        Write-Host "[$pc]     Added to db_owner"
    }

    # --- 5. Schema: FoundationDBUtils create ---
    Write-Host "[$pc] 5/7 OnBase schema..."
    if ([int](Get-Scalar $schemaSql $db) -eq 1) {
        Write-Host "[$pc]     hsi.productsold exists - DBUtils skipped"
    }
    else {
        # Access must succeed from THIS execution context, not just the workstation.
        if (-not (Test-Path -LiteralPath $dbUtilsSrc -PathType Container)) {
            throw "FoundationDBUtils directory is not reachable from ${pc}: $dbUtilsSrc. Check the path and target-session permissions."
        }
        if (-not (Test-Path -LiteralPath (Join-Path $dbUtilsSrc 'FoundationDBUtils.exe') -PathType Leaf)) {
            throw "FoundationDBUtils.exe not found in $dbUtilsSrc"
        }
        $toolDir = 'C:\Temp\FoundationDBUtils'
        Write-Host "[$pc]     Copying FoundationDBUtils from $dbUtilsSrc..."
        robocopy $dbUtilsSrc $toolDir /E /NFL /NDL /NJH /NJS /R:2 /W:5 | Out-Null
        if ($LASTEXITCODE -ge 8) { throw "Copying FoundationDBUtils failed (robocopy exit $LASTEXITCODE)." }
        $exe = Join-Path $toolDir 'FoundationDBUtils.exe'
        if (-not (Test-Path -LiteralPath $exe -PathType Leaf)) { throw "FoundationDBUtils.exe not found after copying from $dbUtilsSrc" }

        if ($sqlAdmin) {
            $dbuCs = 'Data Source={0};database={1};User ID={2};Password={3};Encrypt=Optional;' -f $sqlInstance, $db, $sqlAdmin.UserName, $sqlAdmin.GetNetworkCredential().Password
        }
        else {
            $dbuCs = 'Data Source={0};database={1};Integrated Security=true;Encrypt=Optional;' -f $sqlInstance, $db
        }
        $dbuArgs = 'create --datasourcename "{0}" --connectionstring "{1}" --provider "System.Data.SqlClient" --hsipassword "{2}" --viewerpassword "{3}" --tablespaceoption "{4}" --verbose' -f $dsn, $dbuCs, $hsiPw, $viewerPw, $tsOption
        if ($unicode) { $dbuArgs += ' --unicode' }
        if ($testSystem) { $dbuArgs += ' --testsystem' }
        if ($use32) { $dbuArgs += ' --use32bit' }

        $outLog = 'C:\Temp\DBUtilsCreate.log'
        $errLog = 'C:\Temp\DBUtilsCreate.err.log'
        Write-Host "[$pc]     Running FoundationDBUtils create (timeout $timeoutMinutes min, log $outLog)..."
        $startParams = @{
            FilePath               = $exe
            ArgumentList           = $dbuArgs
            WorkingDirectory       = $toolDir
            RedirectStandardOutput = $outLog
            RedirectStandardError  = $errLog
            NoNewWindow            = $true
            PassThru               = $true
        }
        $p = Start-Process @startParams
        $null = $p.Handle
        if (-not $p.WaitForExit($timeoutMinutes * 60 * 1000)) {
            Stop-Process -Id $p.Id -Force -ErrorAction SilentlyContinue
            throw "FoundationDBUtils timed out after $timeoutMinutes min. See $outLog"
        }
        # Retained source behavior: a nonzero code warns; schema presence and
        # final readiness checks determine success. This is not a full schema audit.
        if ($p.ExitCode -ne 0) { Write-Warning "[$pc] FoundationDBUtils exit code $($p.ExitCode). See $outLog and $errLog" }
        if ([int](Get-Scalar $schemaSql $db) -ne 1) { throw "Schema not found after DBUtils create. Check $outLog and $errLog" }
        $dbuArgs = $null
        $dbuCs = $null
        Write-Host "[$pc]     Schema created"
    }

    # --- 6. ODBC System DSN (32-bit and optional 64-bit) ---
    # Retains the original registry-based DSN configuration and driver fallbacks.
    Write-Host "[$pc] 6/7 ODBC DSN '$dsn'..."
    $odbc32IniPath = 'HKLM:\SOFTWARE\WOW6432Node\ODBC\ODBC.INI'
    $dsn32RegPath = "$odbc32IniPath\$dsn"
    $dataSources32 = "$odbc32IniPath\ODBC Data Sources"
    $driverDll32 = (Get-ItemProperty "HKLM:\SOFTWARE\WOW6432Node\ODBC\ODBCINST.INI\$driver" -ErrorAction SilentlyContinue).Driver
    if (-not $driverDll32) {
        if ($driver -like '*18*') { $driverDll32 = "$env:SystemRoot\SysWOW64\msodbcsql18.dll" }
        elseif ($driver -like '*17*') { $driverDll32 = "$env:SystemRoot\SysWOW64\msodbcsql17.dll" }
        else { $driverDll32 = "$env:SystemRoot\SysWOW64\sqlncli11.dll" }
    }
    if (-not (Test-Path $dsn32RegPath)) { New-Item -Path $dsn32RegPath -Force | Out-Null }
    Set-ItemProperty -Path $dsn32RegPath -Name 'Driver' -Value $driverDll32 -Force
    Set-ItemProperty -Path $dsn32RegPath -Name 'Server' -Value $sqlInstance -Force
    Set-ItemProperty -Path $dsn32RegPath -Name 'Database' -Value $db -Force
    Set-ItemProperty -Path $dsn32RegPath -Name 'LastUser' -Value $hsiCred.UserName -Force
    Set-ItemProperty -Path $dsn32RegPath -Name 'Description' -Value "OnBase Database ($db)" -Force
    if ($driver -like '*18*') {
        Set-ItemProperty -Path $dsn32RegPath -Name 'Encrypt' -Value 'Optional' -Force
        Set-ItemProperty -Path $dsn32RegPath -Name 'TrustServerCertificate' -Value 'Yes' -Force
    }
    if (-not (Test-Path $dataSources32)) { New-Item -Path $dataSources32 -Force | Out-Null }
    Set-ItemProperty -Path $dataSources32 -Name $dsn -Value $driver -Force
    Write-Host "[$pc]     32-bit DSN '$dsn' configured (Server: $sqlInstance, Driver: $driver)"

    if ($want64) {
        $odbc64IniPath = 'HKLM:\SOFTWARE\ODBC\ODBC.INI'
        $dsn64RegPath = "$odbc64IniPath\$dsn"
        $dataSources64 = "$odbc64IniPath\ODBC Data Sources"
        $driverDll64 = (Get-ItemProperty "HKLM:\SOFTWARE\ODBC\ODBCINST.INI\$driver" -ErrorAction SilentlyContinue).Driver
        if (-not $driverDll64) {
            if ($driver -like '*18*') { $driverDll64 = "$env:SystemRoot\System32\msodbcsql18.dll" }
            elseif ($driver -like '*17*') { $driverDll64 = "$env:SystemRoot\System32\msodbcsql17.dll" }
            else { $driverDll64 = "$env:SystemRoot\System32\sqlncli11.dll" }
        }
        if (-not (Test-Path $dsn64RegPath)) { New-Item -Path $dsn64RegPath -Force | Out-Null }
        Set-ItemProperty -Path $dsn64RegPath -Name 'Driver' -Value $driverDll64 -Force
        Set-ItemProperty -Path $dsn64RegPath -Name 'Server' -Value $sqlInstance -Force
        Set-ItemProperty -Path $dsn64RegPath -Name 'Database' -Value $db -Force
        Set-ItemProperty -Path $dsn64RegPath -Name 'LastUser' -Value $hsiCred.UserName -Force
        Set-ItemProperty -Path $dsn64RegPath -Name 'Description' -Value "OnBase Database ($db)" -Force
        if ($driver -like '*18*') {
            Set-ItemProperty -Path $dsn64RegPath -Name 'Encrypt' -Value 'Optional' -Force
            Set-ItemProperty -Path $dsn64RegPath -Name 'TrustServerCertificate' -Value 'Yes' -Force
        }
        if (-not (Test-Path $dataSources64)) { New-Item -Path $dataSources64 -Force | Out-Null }
        Set-ItemProperty -Path $dataSources64 -Name $dsn -Value $driver -Force
        Write-Host "[$pc]     64-bit DSN '$dsn' configured (Server: $sqlInstance, Driver: $driver)"
    }

    # --- 7. Readiness report ---
    Write-Host "[$pc] 7/7 Readiness checks..."
    $checks = New-Object System.Collections.Generic.List[object]
    function Add-Check([string]$name, $pass, [string]$detail) {
        $checks.Add([pscustomobject]@{ Check = $name; Pass = [bool]$pass; Detail = $detail })
    }
    $maxdopNow = [int](Get-Scalar "SELECT CAST(value_in_use AS int) FROM sys.configurations WHERE name = 'max degree of parallelism'")
    $state = Get-Scalar "SELECT state_desc FROM sys.databases WHERE name = N'$dbLit'"
    $hsiLogin = [int](Get-Scalar "SELECT COUNT(*) FROM sys.server_principals WHERE name = N'hsi'") -eq 1
    $secAdmin = [int](Get-Scalar "SELECT ISNULL(IS_SRVROLEMEMBER('securityadmin', 'hsi'), 0)") -eq 1
    $dbOwner = [int](Get-Scalar $dbOwnerSql $db) -ge 1
    $viewer = [int](Get-Scalar "SELECT COUNT(*) FROM sys.database_principals WHERE name = N'viewer'" $db) -eq 1
    $schema = [int](Get-Scalar $schemaSql $db) -eq 1
    $hsiCanLogin = $false
    try { $null = Get-Scalar 'SELECT 1' $db $hsiCs; $hsiCanLogin = $true } catch { }

    Add-Check 'SqlReachable' $true ("{0} as {1}" -f $info.Version, $info.LoginName)
    Add-Check 'MaxDop' ($maxdopNow -eq 1) "value_in_use = $maxdopNow"
    Add-Check 'DatabaseOnline' ($state -eq 'ONLINE') "state = $state"
    Add-Check 'HsiLogin' $hsiLogin 'server login hsi'
    Add-Check 'HsiCanLogin' $hsiCanLogin 'SQL auth login with the supplied hsi password'
    Add-Check 'HsiSecurityAdmin' $secAdmin 'securityadmin server role'
    Add-Check 'HsiDbOwner' $dbOwner "db_owner in [$db]"
    Add-Check 'ViewerUser' $viewer "viewer user in [$db]"
    Add-Check 'OnBaseSchema' $schema 'hsi.productsold exists (not a full schema audit)'
    foreach ($plat in $platforms) {
        $hasDsn = [bool](Get-OdbcDsn -Name $dsn -DsnType System -Platform $plat -ErrorAction SilentlyContinue)
        Add-Check ("Dsn " + $plat) $hasDsn "System DSN $dsn exists (not a DSN connection test)"
    }
    $ready = -not ($checks | Where-Object { -not $_.Pass })
    $hsiPw = $null
    $viewerPw = $null
    $hsiPwSql = $null
    $hsiCs = $null
    $adminCs = $null
    [pscustomobject]@{ Computer = $pc; Ready = $ready; Checks = $checks.ToArray() }
}

# Positional order must match the target script block's 15 parameters.
$scriptArgs = @($SqlInstance, $DatabaseName, $DataFilePath, $SqlAdminCredential, $HsiCredential, $ViewerCredential,
                $DbUtilsSourcePath, $TablespaceOption, [bool]$Unicode, [bool]$TestSystem, [bool]$Use32Bit,
                $DsnName, $OdbcDriver, [bool]$Create64BitDsn, $TimeoutMinutes)
try {
    if ($DeploymentMode -eq 'Local') {
        Write-Host 'Preparing OnBase database locally...'
        $result = & $dbReadyScript @scriptArgs
    }
    else {
        Write-Host "Preparing OnBase database on $TargetMachine..."
        $invokeParams = @{ ComputerName = $TargetMachine; ScriptBlock = $dbReadyScript; ArgumentList = $scriptArgs }
        if ($Credential) {
            $invokeParams.Credential = $Credential
            $invokeParams.Authentication = 'Negotiate'
        }
        if ($ConfigurationName) { $invokeParams.ConfigurationName = $ConfigurationName }
        $result = Invoke-Command @invokeParams
    }
}
catch {
    Write-Host "Database step failed: $($_.Exception.Message)"
    exit 1
}

$result = $result | Where-Object { $_ -and $_.PSObject.Properties['Ready'] } | Select-Object -Last 1
if (-not $result) { Write-Host 'No result returned from the target.'; exit 1 }
$result.Checks | Format-Table Check, Pass, Detail -AutoSize | Out-String | Write-Host
if ($result.Ready) {
    Write-Host "[$($result.Computer)] Database READY"
    exit 0
}
Write-Host "[$($result.Computer)] Database NOT ready - fix the failed checks above and re-run."
exit 1
