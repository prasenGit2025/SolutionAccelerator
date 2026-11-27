#requires -Version 5.1
<#
.SYNOPSIS
 Detects, installs and verifies a specific SQL instance using Settings.json.
 Separate Driver 17 install/upgrade plus explicit, isolated same-package recovery.

.DESCRIPTION
 Revision: 2026-11-24-source-recovery-1. Based on the script supplied by the user.
 Requires 64-bit Windows PowerShell 5.1 on controller and target.
 Precedence: explicit CLI > JSON > default. No external helper/catalog required.

 Keep application Driver 18 settings unchanged. The separate target-side MSI is:
   "sqlOdbc17MsiPath": "C:\\InstallMedia\\ODBC17\\msodbcsql.msi"
 The file is NOT automatically copied from the controller.

 NORMAL MODE
 A configured Driver 17 MSI is validated and installed/upgraded only when
 missing/older. Same/newer versions are skipped. This does NOT certify MSI
 source health. No forced repair, downgrade or uninstall is performed.

 EXPLICIT SOURCE-RECOVERY MODE
   -SqlOdbc17Only -RepairSqlOdbc17
 Opt-in maintenance for the SAME installed x64 English Driver 17 package.
 Requires matching ProductCode, PackageCode, ProductVersion and language.
 Installed package identity is read through Windows Installer, not guessed
 from a version number. Unreadable identity, applied MSI patches, transforms,
 duplicate registrations or any mismatch stop the operation before repair.
 The source also needs a valid Microsoft signature. The registered original
 package may be missing from the Windows Installer cache; this is allowed
 only if its registered identity can still be established.

 Recovery stages the validated MSI in a protected stable package directory,
 then runs msiexec /fv with quiet/no-restart/verbose-log options. This requests
 Windows Installer's repair/recache operation; it is NOT a file-copy-only fix.
 It can run package maintenance/custom actions and affect applications using
 the driver. Use an approved maintenance window and recovery plan.

 Recovery NEVER starts SQL Setup or application Driver 18, and never directly
 edits MSI registry/source entries or copies into C:\Windows\Installer. Only
 Windows Installer itself manages its cache. Successful recovery is followed
 by product/driver checks and verification that the registered LocalPackage
 exists and has the expected package identity. This is NOT a guarantee that
 SQL Setup or all future servicing will succeed. No automatic SQL retry.

 No new JSON repair flag: repair authorization is CLI-only. Repair requires
 -SqlOdbc17Only, even for detection, to prevent accidental combined SQL work.
 -DetectOnly is read-only: no directories/files created and no installer run.
 -SkipOdbc affects ONLY the original application ODBC stage.

 The stable cache is %ProgramData%\OnBaseSqlInstaller\ODBC17\<version>-<hash>.
 Logs: %ProgramData%\OnBaseSqlInstaller\Logs. Hashes identify staged bytes;
 a freshly computed SHA256 is NOT an independently approved/pinned checksum.
 Obtain packages from an approved trusted source. No online downloader.

 SQL WORKFLOW CAUTIONS
 Existing SQL workflow retained: 2016/2017/2019/2022; full media directory or
 ISO. No EXE/web downloader. A new sa password is requested only for a new
 SQL instance. Existing sa passwords are not reset or read from settings.
 Replace can back up/uninstall ALL other-major SQL instances. Backups are NOT
 automatically restored. Do not use Replace for this ODBC recovery.
 Basic partial-instance checks are NOT a complete SQL Setup-state assessment.
 Assess previous failed SQL Setup logs before resuming full provisioning.
 WinRM, credentials, delegation and source permissions are not configured here.
 SQL Setup receives SAPWD on its process command line; it is never printed
 here. Protect administrative access, setup logs and crash dumps.

 EXIT CODES
 0: operation/read-only detection succeeded (a planned action is not execution).
 3010: reboot required. Stop dependent stages; reboot and reassess.
 1: failed/manual action required. No automatic retry.

.NOTES
 Reviewed as text ONLY. NOT parsed or executed in Windows PowerShell here.
 Syntax-check and lab-test before production use. ODBC maintenance support
 can vary by package/environment; failures retain logs and require review.
 Original SQL replacement/backups/authentication/application ODBC behavior
 has not been runtime-validated by this revision.

.EXAMPLE
 .\PowerShell\Install-SqlServer.ps1 -SettingsFilePath .\Settings.json -SqlOdbc17Only -RepairSqlOdbc17 -DetectOnly

.EXAMPLE
 # Run only matching-package Driver 17 maintenance; no SQL or Driver 18 stage:
 .\PowerShell\Install-SqlServer.ps1 -SettingsFilePath .\Settings.json -SqlOdbc17Only -RepairSqlOdbc17

.EXAMPLE
 # Normal read-only SQL/application ODBC detection, no recovery authorization:
 .\PowerShell\Install-SqlServer.ps1 -SettingsFilePath .\Settings.json -DetectOnly
#>
[CmdletBinding()]
param(
    [Alias('SettingsPath')]
    [string]$SettingsFilePath = (Join-Path (Split-Path -Parent $PSScriptRoot) 'Settings.json'),
    [ValidateSet('Local', 'Remote')]
    [string]$DeploymentMode = 'Local',
    [string]$TargetMachine = '',
    [pscredential]$Credential,
    [string]$ConfigurationName = '',
    [string]$SqlSetupPath = '',
    [string]$InstanceName = 'SQL2K22',
    [ValidateRange(13, 16)]
    [int]$TargetMajorVersion = 16,
    [string]$SqlVersion = '',
    [string]$SqlInstance = '',
    [pscredential]$SaCredential,
    [string[]]$SysAdminAccounts = @('BUILTIN\Administrators'),
    [Alias('OnSqlVersionMismatch')]
    [ValidateSet('Stop', 'SideBySide', 'Replace')]
    [string]$OnVersionMismatch = 'Stop',
    [switch]$SkipBackup,
    [string]$BackupPath = 'C:\Temp\SqlBackup',
    [string]$OdbcDriverMsiPath = '',
    [string]$OdbcDriverName = 'ODBC Driver 18 for SQL Server',
    [string]$OdbcVersion = '',
    [string]$SqlOdbc17MsiPath = '',
    [switch]$SqlOdbc17Only,
    # Explicit CLI-only authorization. Requires -SqlOdbc17Only.
    [switch]$RepairSqlOdbc17,
    [switch]$SkipOdbc,
    [switch]$DetectOnly,
    [switch]$NonInteractive,
    [ValidateRange(1, 1048576)]
    [int]$MinFreeSpaceGB = 10,
    [ValidateRange(1, 1440)]
    [int]$TimeoutMinutes = 60
)

$ErrorActionPreference = 'Stop'
$session = $null
$exitCode = 1
$remoteUsername = ''
$settingsResolvedPath = $null
$saForInstall = $null
$wasBound = @{}
foreach ($key in $PSBoundParameters.Keys) { $wasBound[$key] = $true }

function Get-SettingValue {
    param($Object, [string]$Name)
    if ($null -ne $Object) {
        $property = $Object.PSObject.Properties[$Name]
        if ($null -ne $property -and $null -ne $property.Value) {
            if ($property.Value -is [string] -and [string]::IsNullOrWhiteSpace($property.Value)) { return $null }
            return $property.Value
        }
    }
    return $null
}

function Test-SqlPasswordComplexity {
    param([string]$Password)
    if ([string]::IsNullOrEmpty($Password) -or $Password.Length -lt 8) { return $false }
    if ($Password -match '["\r\n\x00]' -or $Password.EndsWith('\')) { return $false }
    $classes = 0
    if ($Password -cmatch '[A-Z]') { $classes++ }
    if ($Password -cmatch '[a-z]') { $classes++ }
    if ($Password -match '[0-9]') { $classes++ }
    if ($Password -match '[^A-Za-z0-9]') { $classes++ }
    return ($classes -ge 3)
}

function Test-SaCredentialValue {
    param([pscredential]$Value)
    if ($null -eq $Value) { return $false }
    $plainText = $null
    try {
        $plainText = $Value.GetNetworkCredential().Password
        return (Test-SqlPasswordComplexity -Password $plainText)
    }
    finally { $plainText = $null }
}

$sqlScript = {
    param([hashtable]$Options, [pscredential]$SaCred, [bool]$DetectionOnly)
    $ErrorActionPreference = 'Stop'
    $pc = $env:COMPUTERNAME
    $root = 'HKLM:\SOFTWARE\Microsoft\Microsoft SQL Server'
    $targetMajor = [int]$Options.TargetMajorVersion
    $targetName = [string]$Options.TargetName
    $instance = [string]$Options.InstanceName
    $setupSrc = [string]$Options.SqlSetupPath
    # Original APPLICATION driver source, separate from Driver 17.
    $odbcMsi = [string]$Options.OdbcDriverMsiPath
    $odbcName = [string]$Options.OdbcDriverName
    $sqlOdbc17Msi = [string]$Options.SqlOdbc17MsiPath
    $driver17Only = [bool]$Options.SqlOdbc17Only
    $repairDriver17 = [bool]$Options.RepairSqlOdbc17
    $timeout = [int]$Options.TimeoutMinutes
    $driver17Report = [pscustomobject]@{
        Status = 'NotConfigured'; Version = $null; Source = $sqlOdbc17Msi
        CachedMsi = $null; LogPath = $null; Sha256 = $null
        ProductCode = $null; PackageCode = $null; InstalledPackageCode = $null
        LocalPackage = $null; CacheVerified = $false; MsiExitCode = $null
    }

    if (-not [Environment]::Is64BitProcess) { throw 'Use a 64-bit target PowerShell endpoint/process for SQL and ODBC detection.' }
    if ($repairDriver17 -and -not $driver17Only) { throw 'RepairSqlOdbc17 requires SqlOdbc17Only. Recovery must not be combined with SQL Setup.' }

    function Get-RegistryValue {
        param($Object, [string]$Name)
        if ($null -ne $Object) {
            $property = $Object.PSObject.Properties[$Name]
            if ($null -ne $property) { return $property.Value }
        }
        return $null
    }

    function Get-SqlInstances {
        $mapPath = "$root\Instance Names\SQL"
        if (-not (Test-Path -LiteralPath $mapPath)) { return }
        $map = Get-ItemProperty -LiteralPath $mapPath -ErrorAction Stop
        $metadata = @('PSPath', 'PSParentPath', 'PSChildName', 'PSDrive', 'PSProvider')
        foreach ($property in $map.PSObject.Properties) {
            if ($property.Name -in $metadata -or $property.MemberType -ne 'NoteProperty') { continue }
            $registryId = [string]$property.Value
            $setup = Get-ItemProperty -LiteralPath "$root\$registryId\Setup" -ErrorAction Stop
            $auth = Get-ItemProperty -LiteralPath "$root\$registryId\MSSQLServer" -ErrorAction Stop
            $versionText = [string](Get-RegistryValue $setup 'Version')
            if (-not $versionText) { $versionText = [string](Get-RegistryValue $setup 'PatchLevel') }
            $version = $null
            if (-not [version]::TryParse($versionText, [ref]$version)) { throw "Cannot determine SQL version for instance '$($property.Name)'. Assess/repair its installation manually." }
            $serviceName = if ($property.Name -eq 'MSSQLSERVER') { 'MSSQLSERVER' } else { 'MSSQL$' + $property.Name }
            $service = Get-Service -Name $serviceName -ErrorAction SilentlyContinue
            [pscustomobject]@{
                Instance = $property.Name
                Server = $(if ($property.Name -eq 'MSSQLSERVER') { $pc } else { "$pc\$($property.Name)" })
                Major = $version.Major
                Version = $(if (Get-RegistryValue $setup 'PatchLevel') { Get-RegistryValue $setup 'PatchLevel' } else { $versionText })
                Edition = Get-RegistryValue $setup 'Edition'
                MixedMode = ((Get-RegistryValue $auth 'LoginMode') -eq 2)
                Service = $serviceName
                Status = $(if ($null -ne $service) { [string]$service.Status } else { 'NotFound' })
            }
        }
    }

    function Get-OdbcState {
        if ([bool]$Options.SkipOdbc -or $driver17Only) { return [pscustomobject]@{ Missing = @(); Drivers32 = @(); Drivers64 = @(); Skipped = $true } }
        $drivers32 = @(Get-OdbcDriver -Platform '32-bit' -ErrorAction Stop)
        $drivers64 = @(Get-OdbcDriver -Platform '64-bit' -ErrorAction Stop)
        $missing = @()
        if (-not @($drivers64 | Where-Object { $_.Name -eq $odbcName }).Count) { $missing += '64-bit' }
        if (-not @($drivers32 | Where-Object { $_.Name -eq $odbcName }).Count) { $missing += '32-bit' }
        [pscustomobject]@{
            Missing = $missing
            Drivers32 = @($drivers32 | Where-Object { $_.Name -like '*SQL*' } | Select-Object -ExpandProperty Name)
            Drivers64 = @($drivers64 | Where-Object { $_.Name -like '*SQL*' } | Select-Object -ExpandProperty Name)
            Skipped = $false
        }
    }

    function New-Result {
        param([string]$Status, [string]$Message, $Selected, [bool]$Reboot)
        $odbcInventoryError = $null
        try { $driverState = Get-OdbcState }
        catch {
            $odbcInventoryError = $_.Exception.Message
            $driverState = [pscustomobject]@{ Missing = @('Unknown'); Drivers32 = @(); Drivers64 = @(); Skipped = $false }
        }
        $inventoryError = $null
        $inventory = @()
        try { $inventory = @(Get-SqlInstances) }
        catch { $inventoryError = $_.Exception.Message; Write-Warning "SQL inventory could not be fully read: $inventoryError" }
        [pscustomobject]@{
            Computer = $pc; Status = $Status; Message = $Message
            SqlInstance = $(if ($null -ne $Selected) { $Selected.Server } else { $null })
            RebootRequired = $Reboot
            OdbcSkipped = ([bool]$Options.SkipOdbc -or $driver17Only)
            OdbcDrivers32 = @($driverState.Drivers32); OdbcDrivers64 = @($driverState.Drivers64)
            OdbcMissing = @($driverState.Missing); OdbcInventoryError = $odbcInventoryError
            Instances = $inventory; SqlInventoryError = $inventoryError
            SqlOdbc17Only = $driver17Only; SqlOdbc17RepairRequested = $repairDriver17
            SqlOdbc17Status = $driver17Report.Status
            SqlOdbc17Version = $driver17Report.Version
            SqlOdbc17MsiPath = $sqlOdbc17Msi
            SqlOdbc17CachedMsi = $driver17Report.CachedMsi
            SqlOdbc17LogPath = $driver17Report.LogPath
            SqlOdbc17Sha256 = $driver17Report.Sha256
            SqlOdbc17ProductCode = $driver17Report.ProductCode
            SqlOdbc17PackageCode = $driver17Report.PackageCode
            SqlOdbc17InstalledPackageCode = $driver17Report.InstalledPackageCode
            SqlOdbc17LocalPackage = $driver17Report.LocalPackage
            SqlOdbc17CacheVerified = $driver17Report.CacheVerified
            SqlOdbc17MsiExitCode = $driver17Report.MsiExitCode
        }
    }

    function Test-PendingReboot {
        if ((Test-Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending') -or
            (Test-Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired')) { return $true }
        $sessionManager = Get-ItemProperty -LiteralPath 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager' -ErrorAction SilentlyContinue
        $pendingRenames = @(Get-RegistryValue $sessionManager 'PendingFileRenameOperations' | Where-Object { $_ })
        return ($pendingRenames.Count -gt 0)
    }

    function Invoke-Proc {
        param([string]$Exe, [string]$Arguments, [int]$Minutes)
        $process = Start-Process -FilePath $Exe -ArgumentList $Arguments -PassThru -ErrorAction Stop
        $null = $process.Handle
        if ($process.WaitForExit($Minutes * 60 * 1000)) { $process.WaitForExit(); return [int]$process.ExitCode }
        Stop-Process -Id $process.Id -Force -ErrorAction SilentlyContinue
        # MSI service/children may outlive the launcher. Never retry automatically.
        return -1
    }

    function Invoke-Sql {
        param([string]$Server, [string]$Sql)
        $builder = New-Object System.Data.SqlClient.SqlConnectionStringBuilder
        $builder.DataSource = $Server; $builder.InitialCatalog = 'master'
        $builder.IntegratedSecurity = $true; $builder.TrustServerCertificate = $true
        $connection = New-Object System.Data.SqlClient.SqlConnection($builder.ConnectionString)
        try {
            $connection.Open()
            $command = $connection.CreateCommand()
            $command.CommandText = $Sql; $command.CommandTimeout = 3600
            $table = New-Object System.Data.DataTable
            $reader = $command.ExecuteReader()
            try { $table.Load($reader) } finally { $reader.Dispose() }
            return ,$table
        }
        finally { $connection.Dispose() }
    }

    function Read-MsiProperties {
        param([string]$Path, [switch]$IncludeIdentity)
        $installer = $null; $database = $null; $summary = $null
        $properties = @{}
        try {
            $installer = New-Object -ComObject WindowsInstaller.Installer
            $database = $installer.GetType().InvokeMember('OpenDatabase', 'InvokeMethod', $null, $installer, @($Path, 0))
            $names = @('ProductName', 'ProductVersion')
            if ($IncludeIdentity) { $names += @('ProductCode', 'ProductLanguage', 'Manufacturer') }
            foreach ($name in $names) {
                $view = $null; $record = $null
                try {
                    $query = "SELECT ``Value`` FROM ``Property`` WHERE ``Property``='$name'"
                    $view = $database.GetType().InvokeMember('OpenView', 'InvokeMethod', $null, $database, @($query))
                    $null = $view.GetType().InvokeMember('Execute', 'InvokeMethod', $null, $view, $null)
                    $record = $view.GetType().InvokeMember('Fetch', 'InvokeMethod', $null, $view, $null)
                    if ($null -eq $record) { throw "MSI property '$name' is missing." }
                    $properties[$name] = [string]$record.GetType().InvokeMember('StringData', 'GetProperty', $null, $record, @(1))
                }
                finally {
                    if ($null -ne $view) { $null = $view.GetType().InvokeMember('Close', 'InvokeMethod', $null, $view, $null) }
                    if ($null -ne $record) { $null = [Runtime.InteropServices.Marshal]::FinalReleaseComObject($record) }
                    if ($null -ne $view) { $null = [Runtime.InteropServices.Marshal]::FinalReleaseComObject($view) }
                }
            }
            if ($IncludeIdentity) {
                $summary = $database.GetType().InvokeMember('SummaryInformation', 'GetProperty', $null, $database, @(0))
                $properties['Template'] = [string]$summary.GetType().InvokeMember('Property', 'GetProperty', $null, $summary, @(7))
                $properties['PackageCode'] = [string]$summary.GetType().InvokeMember('Property', 'GetProperty', $null, $summary, @(9))
            }
            return $properties
        }
        finally {
            if ($null -ne $summary) { $null = [Runtime.InteropServices.Marshal]::FinalReleaseComObject($summary) }
            if ($null -ne $database) { $null = [Runtime.InteropServices.Marshal]::FinalReleaseComObject($database) }
            if ($null -ne $installer) { $null = [Runtime.InteropServices.Marshal]::FinalReleaseComObject($installer) }
        }
    }

    function ConvertTo-MsiGuid {
        param([string]$Value, [string]$Label)
        $parsed = [guid]::Empty
        if (-not [guid]::TryParse($Value, [ref]$parsed) -or $parsed -eq [guid]::Empty) {
            throw "Cannot establish a valid $Label. No Driver 17 repair is authorized for an unknown identity."
        }
        return $parsed.ToString('B').ToUpperInvariant()
    }

    function Get-Driver17InstalledState {
        $name = 'Microsoft ODBC Driver 17 for SQL Server'
        # Native x64 view in a mandatory 64-bit process. Never query Win32_Product.
        $entries = @(Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*' -ErrorAction SilentlyContinue |
            Where-Object { ([string](Get-RegistryValue $_ 'DisplayName')).Trim() -ieq $name })
        $versions = @(); $productCodes = @()
        foreach ($entry in $entries) {
            $parsed = $null
            if (-not [version]::TryParse([string](Get-RegistryValue $entry 'DisplayVersion'), [ref]$parsed) -or $parsed.Major -ne 17) {
                throw 'Driver 17 has an unrecognized x64 installed version. Assess its registration before making changes.'
            }
            $versions += $parsed
            $productCodes += [string](Get-RegistryValue $entry 'PSChildName')
        }
        $drivers64 = @(Get-OdbcDriver -Platform '64-bit' -Name 'ODBC Driver 17 for SQL Server' -ErrorAction SilentlyContinue)
        [pscustomobject]@{
            Versions = @($versions); ProductCodes = @($productCodes)
            HasDriver64 = ($drivers64.Count -gt 0); RegistrationCount = $entries.Count
        }
    }

    function Get-InstalledMsiIdentity {
        param([string]$ProductCode)
        $normalizedProduct = ConvertTo-MsiGuid $ProductCode 'installed ProductCode'
        $installer = $null; $patches = $null
        $values = @{}
        try {
            $installer = New-Object -ComObject WindowsInstaller.Installer
            # Read-only automation properties. Missing/unreadable values fail closed.
            foreach ($property in @('PackageCode', 'VersionString', 'Language', 'LocalPackage', 'Transforms')) {
                $values[$property] = [string]$installer.GetType().InvokeMember('ProductInfo', 'GetProperty', $null, $installer, @($normalizedProduct, $property))
            }
            $patches = $installer.GetType().InvokeMember('Patches', 'GetProperty', $null, $installer, @($normalizedProduct))
            if ($null -eq $patches) { throw 'Could not enumerate the installed product patches.' }
            $patchCount = [int]$patches.GetType().InvokeMember('Count', 'GetProperty', $null, $patches, $null)
            $installedVersion = $null
            if (-not [version]::TryParse($values.VersionString, [ref]$installedVersion)) { throw 'Registered MSI VersionString is invalid.' }
            [pscustomobject]@{
                ProductCode = $normalizedProduct
                PackageCode = (ConvertTo-MsiGuid $values.PackageCode 'registered PackageCode')
                Version = $installedVersion; Language = $values.Language
                LocalPackage = $values.LocalPackage; Transforms = $values.Transforms
                PatchCount = $patchCount
            }
        }
        catch {
            throw "Cannot verify installed MSI identity/maintenance state for $normalizedProduct. No automatic identity bypass is permitted. Detail: $($_.Exception.Message)"
        }
        finally {
            if ($null -ne $patches) { $null = [Runtime.InteropServices.Marshal]::FinalReleaseComObject($patches) }
            if ($null -ne $installer) { $null = [Runtime.InteropServices.Marshal]::FinalReleaseComObject($installer) }
        }
    }

    function Assert-Driver17RepairIdentity {
        param($Properties, [version]$Version, $State)
        if ($State.RegistrationCount -ne 1 -or $State.Versions.Count -ne 1) {
            throw 'Source recovery requires exactly one registered x64 Driver 17 product. No install/upgrade is performed in recovery mode.'
        }
        $sourceProduct = ConvertTo-MsiGuid $Properties.ProductCode 'supplied ProductCode'
        $sourcePackage = ConvertTo-MsiGuid $Properties.PackageCode 'supplied PackageCode'
        $registeredProduct = ConvertTo-MsiGuid ([string]$State.ProductCodes[0]) 'uninstall-registration ProductCode'
        if ($sourceProduct -ine $registeredProduct -or $State.Versions[0] -ne $Version) {
            throw "Driver 17 source recovery requires the installed product/version, not a different MSI. Installed=$($State.Versions[0]) $registeredProduct; supplied=$Version $sourceProduct."
        }
        $identity = Get-InstalledMsiIdentity $registeredProduct
        if ($identity.PackageCode -ine $sourcePackage -or $identity.Version -ne $Version -or
            $identity.Language -ne [string]$Properties.ProductLanguage) {
            throw "Driver 17 MSI identity mismatch. Registered PackageCode=$($identity.PackageCode), version=$($identity.Version), language=$($identity.Language); supplied PackageCode=$sourcePackage, version=$Version, language=$($Properties.ProductLanguage). Repair was NOT started."
        }
        if ($identity.PatchCount -ne 0 -or -not [string]::IsNullOrWhiteSpace($identity.Transforms)) {
            throw 'This Driver 17 product has applied MSI patches or transforms. The limited source-recovery path will not service it; use a reviewed package-specific maintenance procedure.'
        }
        return $identity
    }

    function Get-Driver17Plan {
        if ([string]::IsNullOrWhiteSpace($sqlOdbc17Msi)) { return $null }
        if (-not (Test-Path -LiteralPath $sqlOdbc17Msi -PathType Leaf)) {
            throw "Driver 17 MSI is not accessible ON ${pc}: $sqlOdbc17Msi. Use a target-local path or a UNC share readable by the target account."
        }
        $signature = Get-AuthenticodeSignature -LiteralPath $sqlOdbc17Msi -ErrorAction Stop
        if ($signature.Status -ne 'Valid' -or $null -eq $signature.SignerCertificate -or
            $signature.SignerCertificate.Subject -notmatch '(?i)(?:^|,\s*)O=Microsoft Corporation(?:,|$)') {
            throw "Driver 17 MSI must have a VALID Microsoft Corporation signature. Status=$($signature.Status). Source=$sqlOdbc17Msi"
        }
        $properties = Read-MsiProperties -Path $sqlOdbc17Msi -IncludeIdentity
        $version = $null
        if ($properties.ProductName -ine 'Microsoft ODBC Driver 17 for SQL Server' -or
            -not [version]::TryParse($properties.ProductVersion, [ref]$version)) {
            throw "Unexpected SQL Driver 17 MSI product/version: $($properties.ProductName), $($properties.ProductVersion)."
        }
        if ($version.Major -ne 17 -or ($properties.Template -split ';')[0] -ine 'x64' -or $properties.ProductLanguage -ne '1033') {
            throw "Expected Driver 17, x64, English (1033). Found version=$version; template=$($properties.Template); language=$($properties.ProductLanguage)."
        }
        $properties.ProductCode = ConvertTo-MsiGuid $properties.ProductCode 'supplied ProductCode'
        $properties.PackageCode = ConvertTo-MsiGuid $properties.PackageCode 'supplied PackageCode'
        $driver17Report.ProductCode = $properties.ProductCode
        $driver17Report.PackageCode = $properties.PackageCode
        $hash = (Get-FileHash -LiteralPath $sqlOdbc17Msi -Algorithm SHA256 -ErrorAction Stop).Hash
        $state = Get-Driver17InstalledState
        $action = 'Install'; $installedIdentity = $null
        if ($state.Versions.Count -gt 1) { throw 'More than one x64 Driver 17 product registration exists. Assess these registrations before maintenance.' }
        if ($repairDriver17) {
            $installedIdentity = Assert-Driver17RepairIdentity $properties $version $state
            $driver17Report.InstalledPackageCode = $installedIdentity.PackageCode
            $driver17Report.LocalPackage = $installedIdentity.LocalPackage
            $action = 'RepairSource'
        }
        elseif ($state.Versions.Count -gt 0) {
            if (-not $state.HasDriver64) { throw 'An x64 Driver 17 product is registered but its ODBC registration is absent. Assess the inconsistent state before an upgrade.' }
            if ($state.Versions[0] -ge $version) { $action = 'Skip' } else { $action = 'Upgrade' }
        }
        elseif ($state.HasDriver64) { throw 'The x64 Driver 17 ODBC registration exists but its MSI version cannot be established. Assess it before proceeding.' }
        [pscustomobject]@{
            Action = $action; PackageVersion = $version; Properties = $properties
            Sha256 = $hash; InstalledState = $state; InstalledIdentity = $installedIdentity
        }
    }

    function Initialize-ProtectedDirectory {
        param([string]$Path)
        if (Test-Path -LiteralPath $Path) {
            $item = Get-Item -LiteralPath $Path -Force
            if (-not $item.PSIsContainer -or ($item.Attributes -band [IO.FileAttributes]::ReparsePoint)) { throw "Refusing a non-directory or reparse-point package-cache path: $Path" }
        }
        else { New-Item -ItemType Directory -Path $Path -Force | Out-Null }
        $admins = New-Object Security.Principal.SecurityIdentifier('S-1-5-32-544')
        $system = New-Object Security.Principal.SecurityIdentifier('S-1-5-18')
        $acl = New-Object Security.AccessControl.DirectorySecurity
        $acl.SetAccessRuleProtection($true, $false)
        $acl.SetOwner($admins)
        $inheritance = [Security.AccessControl.InheritanceFlags]'ContainerInherit, ObjectInherit'
        foreach ($sid in @($admins, $system)) {
            $rule = New-Object Security.AccessControl.FileSystemAccessRule($sid, [Security.AccessControl.FileSystemRights]::FullControl,
                $inheritance, [Security.AccessControl.PropagationFlags]::None, [Security.AccessControl.AccessControlType]::Allow)
            $acl.AddAccessRule($rule)
        }
        Set-Acl -LiteralPath $Path -AclObject $acl -ErrorAction Stop
    }

    function Assert-Driver17Recached {
        param($Plan)
        $state = Get-Driver17InstalledState
        $identity = Assert-Driver17RepairIdentity $Plan.Properties $Plan.PackageVersion $state
        if (-not $state.HasDriver64) { throw 'Driver 17 x64 ODBC registration is absent after maintenance.' }
        $driver17Report.LocalPackage = $identity.LocalPackage
        if ([string]::IsNullOrWhiteSpace($identity.LocalPackage) -or
            -not (Test-Path -LiteralPath $identity.LocalPackage -PathType Leaf)) {
            throw 'Windows Installer reported success, but its registered LocalPackage is not accessible. Recovery is not verified.'
        }
        # Do not compare source/cache hashes: Windows Installer may strip embedded
        # payloads from its cached MSI. Compare MSI identity, not byte equality.
        $cached = Read-MsiProperties -Path $identity.LocalPackage -IncludeIdentity
        if ((ConvertTo-MsiGuid $cached.ProductCode 'cached ProductCode') -ine $Plan.Properties.ProductCode -or
            (ConvertTo-MsiGuid $cached.PackageCode 'cached PackageCode') -ine $Plan.Properties.PackageCode -or
            [version]$cached.ProductVersion -ne $Plan.PackageVersion -or
            $cached.ProductLanguage -ne $Plan.Properties.ProductLanguage -or
            ($cached.Template -split ';')[0] -ine 'x64') {
            throw 'The registered cached MSI does not match the expected Driver 17 identity. Assess the maintenance log before proceeding.'
        }
        $driver17Report.CacheVerified = $true
    }

    function Invoke-Driver17Stage {
        param($Plan)
        if ($null -eq $Plan) { return $false }
        $driver17Report.Version = [string]$Plan.PackageVersion
        $driver17Report.Sha256 = $Plan.Sha256
        if ($Plan.Action -eq 'Skip') {
            $driver17Report.Status = 'SkippedSameOrNewer'
            $driver17Report.Version = [string]$Plan.InstalledState.Versions[0]
            Write-Host "[$pc] Driver 17 is already $($driver17Report.Version), equal to or newer than the supplied MSI. No reinstall/repair is attempted."
            Write-Host "[$pc] This check does not establish MSI source health. Explicit same-package recovery requires -SqlOdbc17Only -RepairSqlOdbc17."
            return $false
        }
        if ($Plan.Action -notin @('Install', 'Upgrade', 'RepairSource')) { throw 'Unrecognized Driver 17 action.' }
        if ($Plan.Action -eq 'RepairSource' -and (-not $repairDriver17 -or -not $driver17Only)) { throw 'Source repair must be explicitly authorized and isolated.' }
        $driver17Report.Status = 'Preparing'
        if (Test-PendingReboot) { $driver17Report.Status = 'RebootRequired'; return $true }
        $cacheRoot = Join-Path $env:ProgramData 'OnBaseSqlInstaller'
        Initialize-ProtectedDirectory $cacheRoot
        $familyRoot = Join-Path $cacheRoot 'ODBC17'
        Initialize-ProtectedDirectory $familyRoot
        $packageRoot = Join-Path $familyRoot ('{0}-{1}' -f $Plan.PackageVersion, $Plan.Sha256)
        Initialize-ProtectedDirectory $packageRoot
        $logRoot = Join-Path $cacheRoot 'Logs'
        Initialize-ProtectedDirectory $logRoot
        $cachedMsi = Join-Path $packageRoot 'msodbcsql.msi'
        $temporaryMsi = Join-Path $packageRoot ('incoming-{0}.msi' -f [guid]::NewGuid().ToString('N'))
        try {
            Copy-Item -LiteralPath $sqlOdbc17Msi -Destination $temporaryMsi -ErrorAction Stop
            if ((Get-FileHash -LiteralPath $temporaryMsi -Algorithm SHA256).Hash -ne $Plan.Sha256) { throw 'Driver 17 source changed while staging; the copy does not match the validated source.' }
            $signature = Get-AuthenticodeSignature -LiteralPath $temporaryMsi -ErrorAction Stop
            if ($signature.Status -ne 'Valid' -or $null -eq $signature.SignerCertificate -or
                $signature.SignerCertificate.Subject -notmatch '(?i)(?:^|,\s*)O=Microsoft Corporation(?:,|$)') { throw 'The staged Driver 17 package lacks a valid Microsoft signature.' }
            if (Test-Path -LiteralPath $cachedMsi) {
                $cachedItem = Get-Item -LiteralPath $cachedMsi -Force
                if ($cachedItem.Attributes -band [IO.FileAttributes]::ReparsePoint) { throw 'The staged MSI is a reparse point; refusing it.' }
                if ((Get-FileHash -LiteralPath $cachedMsi -Algorithm SHA256).Hash -ne $Plan.Sha256) { throw 'An unexpected MSI exists in this version/hash-specific staging directory.' }
            }
            else { Move-Item -LiteralPath $temporaryMsi -Destination $cachedMsi -ErrorAction Stop }
        }
        finally { if (Test-Path -LiteralPath $temporaryMsi) { Remove-Item -LiteralPath $temporaryMsi -Force -ErrorAction SilentlyContinue } }
        $driver17Report.CachedMsi = $cachedMsi
        $logPath = Join-Path $logRoot ('ODBC17-{0}-{1:yyyyMMdd-HHmmss}-{2}.log' -f $Plan.Action, (Get-Date), [guid]::NewGuid().ToString('N'))
        $driver17Report.LogPath = $logPath
        if ($Plan.Action -eq 'RepairSource') {
            # Recheck the actual staged bytes and CURRENT installed identity just
            # before starting maintenance. There is no bypass on mismatch.
            $staged = Read-MsiProperties -Path $cachedMsi -IncludeIdentity
            if ((Get-FileHash -LiteralPath $cachedMsi -Algorithm SHA256).Hash -ne $Plan.Sha256) { throw 'Staged Driver 17 bytes changed before recovery.' }
            $null = Assert-Driver17RepairIdentity $staged ([version]$staged.ProductVersion) (Get-Driver17InstalledState)
            $driver17Report.Status = 'RepairingSource'
            Write-Warning "[$pc] Explicit Driver 17 repair/recache will run package maintenance. SQL Setup and application Driver 18 will NOT run."
            $arguments = '/fv "{0}" /qn /norestart /L*v "{1}"' -f $cachedMsi, $logPath
        }
        else {
            $driver17Report.Status = 'Installing'
            $arguments = '/i "{0}" /qn /norestart IACCEPTMSODBCSQLLICENSETERMS=YES ADDLOCAL=ALL /L*v "{1}"' -f $cachedMsi, $logPath
        }
        Write-Host "[$pc] Driver 17 $($Plan.Action): $($Plan.PackageVersion)"
        Write-Host "[$pc] Stable package source: $cachedMsi"
        Write-Host "[$pc] Driver 17 log: $logPath"
        $code = Invoke-Proc (Join-Path $env:WINDIR 'System32\msiexec.exe') $arguments 15
        $driver17Report.MsiExitCode = $code
        if ($code -in @(3010, 1641)) { $driver17Report.Status = 'RebootRequired'; return $true }
        if ($code -ne 0) {
            $driver17Report.Status = 'Failed'
            $hint = ''
            if ((Test-Path -LiteralPath $logPath) -and (Select-String -LiteralPath $logPath -Pattern '\b1706\b' -Quiet)) {
                $hint = ' The MSI log mentions 1706. Review the product/package/source requested in that log. Do not substitute another version or edit Installer registry/cache entries.'
            }
            if ($code -eq -1) { $hint += ' The launcher timed out; a Windows Installer transaction may still be active. Do not start another installer until assessed.' }
            throw "Driver 17 $($Plan.Action) failed (exit $code). SQL Setup was NOT started by this run. Log: $logPath.$hint"
        }
        if ($Plan.Action -eq 'RepairSource') {
            $driver17Report.Status = 'VerifyingRecovery'
            Assert-Driver17Recached $Plan
            $driver17Report.Status = 'RepairedAndRecached'
            Write-Host "[$pc] Driver 17 maintenance succeeded; registered MSI cache identity verified. This does not certify SQL Setup readiness."
        }
        else {
            $after = Get-Driver17InstalledState
            if (-not $after.HasDriver64 -or $after.Versions.Count -ne 1 -or $after.Versions[0] -lt $Plan.PackageVersion) {
                $driver17Report.Status = 'VerificationFailed'
                throw "Driver 17 returned success but x64 version verification failed. SQL Setup was not started. See $logPath"
            }
            $driver17Report.Version = [string]$after.Versions[0]
            $driver17Report.Status = 'InstalledOrUpgraded'
            Write-Host "[$pc] Driver 17 x64 verified at $($driver17Report.Version). Staged MSI retained."
        }
        if (Test-PendingReboot) { $driver17Report.Status = 'RebootRequired'; return $true }
        return $false
    }

    try {
        $driver17Plan = Get-Driver17Plan
        if ($null -ne $driver17Plan) {
            $driver17Report.Status = 'Planned' + $driver17Plan.Action
            $driver17Report.Version = [string]$driver17Plan.PackageVersion
            $driver17Report.Sha256 = $driver17Plan.Sha256
            Write-Host "[$pc] SQL Driver 17 stage: $($driver17Plan.Action); supplied MSI=$($driver17Plan.PackageVersion)"
            Write-Host "[$pc] SQL Driver 17 target source: $sqlOdbc17Msi"
            if ($driver17Plan.Action -eq 'RepairSource') {
                Write-Host "[$pc] Same-package recovery identity matched: ProductCode=$($driver17Report.ProductCode); PackageCode=$($driver17Report.PackageCode)"
            }
        }
        elseif ($driver17Only) { throw 'SqlOdbc17Only requires sqlOdbc17MsiPath in settings or -SqlOdbc17MsiPath.' }

        if ($driver17Only) {
            if ($DetectionOnly) { return New-Result 'Skipped' "Detection only: Driver 17 action=$($driver17Plan.Action). No files changed; SQL and application ODBC were not run." $null $false }
            $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
            $principal = New-Object Security.Principal.WindowsPrincipal($identity)
            if (-not $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) { throw 'Administrative execution is required on the target.' }
            $needsReboot = Invoke-Driver17Stage $driver17Plan
            if ($needsReboot) { return New-Result 'RebootRequired' 'Restart is required. SQL Setup and application ODBC were not run. Reassess recovery/SQL state after restart.' $null $true }
            if ($repairDriver17) { return New-Result 'Ready' 'Driver 17 maintenance completed; registered cached-package identity and x64 driver registration verified. SQL Setup and application Driver 18 were NOT run. Review prior SQL Setup state before proceeding.' $null $false }
            return New-Result 'Ready' 'Driver 17-only stage completed or skipped. SQL Setup was not run; this does not certify SQL readiness or old MSI source health.' $null $false
        }

        $all = @(Get-SqlInstances)
        Write-Host "[$pc] SQL instances found: $($all.Count)"
        foreach ($item in $all) { Write-Host (' {0}: version={1}; edition={2}; service={3}; MixedMode={4}' -f $item.Server, $item.Version, $item.Edition, $item.Status, $item.MixedMode) }
        $sameName = $all | Where-Object { $_.Instance -eq $instance } | Select-Object -First 1
        $match = $all | Where-Object { $_.Instance -eq $instance -and $_.Major -eq $targetMajor } | Select-Object -First 1
        $otherVersions = @($all | Where-Object { $_.Major -ne $targetMajor })
        $needInstall = ($null -eq $match)
        $replaceCandidates = @()

        if ($needInstall) {
            if ($null -ne $sameName -and $Options.OnVersionMismatch -ne 'Replace') { return New-Result 'ActionRequired' "Instance '$instance' is occupied by SQL major $($sameName.Major). Use another name or explicitly review Replace." $null $false }
            if ($otherVersions.Count -gt 0 -and $Options.OnVersionMismatch -eq 'Stop') { return New-Result 'ActionRequired' 'Other SQL major versions exist. Review them and explicitly select SideBySide or Replace.' $null $false }
            if ($Options.OnVersionMismatch -eq 'Replace') { $replaceCandidates = $otherVersions }
            if ($null -eq $sameName) {
                $expectedService = if ($instance -eq 'MSSQLSERVER') { 'MSSQLSERVER' } else { 'MSSQL$' + $instance }
                $orphanService = Get-Service -Name $expectedService -ErrorAction SilentlyContinue
                $defaultId = 'MSSQL{0}.{1}' -f $targetMajor, $instance
                $defaultData = Join-Path $env:ProgramFiles ("Microsoft SQL Server\$defaultId\MSSQL\DATA")
                if ($null -ne $orphanService -or (Test-Path -LiteralPath "$root\$defaultId") -or (Test-Path -LiteralPath $defaultData)) {
                    return New-Result 'ActionRequired' "Possible partial SQL instance '$instance' detected. Assess previous setup logs before a new install. Use -SqlOdbc17Only for isolated Driver 17 work." $null $false
                }
            }
        }
        elseif ($match.Status -eq 'NotFound') { return New-Result 'ActionRequired' "SQL instance '$instance' has registry entries but no service. Assess/repair the partial installation." $match $false }

        if ($DetectionOnly) {
            $state = Get-OdbcState
            if (-not $state.Skipped) { Write-Host "[$pc] $odbcName missing architectures: $($state.Missing -join ', ')" }
            if ($replaceCandidates.Count -gt 0) { Write-Warning "Replace would uninstall: $(($replaceCandidates.Server) -join ', ')" }
            if ($needInstall) { return New-Result 'NotInstalled' "Detection only: new $targetName instance '$instance' is required. No installer was run." $null $false }
            if (-not $match.MixedMode) { Write-Warning 'Existing instance uses Windows-only authentication; enable Mixed Mode separately.' }
            return New-Result 'Skipped' 'Detection only: exact SQL instance/version exists. No changes made.' $match $false
        }

        $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
        $principal = New-Object Security.Principal.WindowsPrincipal($identity)
        if (-not $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) { throw 'Administrative execution is required on the target.' }
        $odbcBefore = Get-OdbcState
        if (-not $odbcBefore.Skipped) {
            if ($odbcBefore.Missing.Count -gt 0 -and -not $odbcMsi) { throw "$odbcName is missing ($($odbcBefore.Missing -join ', ')). Configure odbcDriverMsiPath or use -SkipOdbc for a separate application ODBC stage." }
            if ($odbcMsi -and -not (Test-Path -LiteralPath $odbcMsi -PathType Leaf)) { throw "Application ODBC MSI not accessible from ${pc}: $odbcMsi" }
            if ($odbcMsi -and [IO.Path]::GetExtension($odbcMsi) -ine '.msi') { throw 'OdbcDriverMsiPath must be an MSI file.' }
        }
        if (Test-PendingReboot) { return New-Result 'RebootRequired' 'Windows reports a pending reboot. Restart the target before running installers.' $match $true }

        $isIso = $false
        if ($needInstall) {
            if (-not $setupSrc) { return New-Result 'ActionRequired' 'Installation required but sqlSetupPath is not supplied.' $null $false }
            if ($null -eq $SaCred) { return New-Result 'ActionRequired' 'Installation required but no new sa credential was supplied.' $null $false }
            if (-not (Test-Path -LiteralPath $setupSrc)) { throw "SQL media not accessible from ${pc}: $setupSrc" }
            $isIso = ([IO.Path]::GetExtension($setupSrc) -ieq '.iso')
            if ($isIso) { if (-not (Test-Path -LiteralPath $setupSrc -PathType Leaf)) { throw 'SQL ISO source must be a file.' } }
            elseif (-not (Test-Path -LiteralPath (Join-Path $setupSrc 'setup.exe') -PathType Leaf)) { throw 'SqlSetupPath must be an ISO or complete directory containing setup.exe.' }
            $freeGB = [math]::Round((Get-PSDrive -Name C).Free / 1GB, 1)
            if ($freeGB -lt [int]$Options.MinFreeSpaceGB) { throw "Only $freeGB GB free on C:; requested minimum is $($Options.MinFreeSpaceGB) GB. Allow extra staging space." }
        }

        $needsReboot = Invoke-Driver17Stage $driver17Plan
        if ($needsReboot) { return New-Result 'RebootRequired' 'Driver 17 or Windows requires a restart. SQL Setup was not started by this run. Restart and rerun.' $match $true }

        New-Item -ItemType Directory -Path 'C:\Temp' -Force | Out-Null
        $odbcLocal = $null; $msiProperties = $null
        if (-not $odbcBefore.Skipped -and $odbcMsi) {
            $odbcCache = Join-Path 'C:\Temp\OnBaseOdbc' ([guid]::NewGuid().ToString('N'))
            New-Item -ItemType Directory -Path $odbcCache -Force | Out-Null
            $odbcLocal = Join-Path $odbcCache (Split-Path $odbcMsi -Leaf)
            Copy-Item -LiteralPath $odbcMsi -Destination $odbcLocal -Force
            $msiProperties = Read-MsiProperties $odbcLocal
            if ($msiProperties.ProductName -notlike "*$odbcName*") { throw "Application ODBC MSI '$($msiProperties.ProductName)' does not match '$odbcName'." }
            $null = [version]$msiProperties.ProductVersion
        }

        if ($needInstall) {
            $stageRoot = Join-Path 'C:\Temp\SqlSetup' ([guid]::NewGuid().ToString('N'))
            $media = Join-Path $stageRoot 'Media'
            New-Item -ItemType Directory -Path $media -Force | Out-Null
            Write-Host "[$pc] Staging SQL media in $media ..."
            if ($isIso) {
                $isoLocal = Join-Path $stageRoot (Split-Path $setupSrc -Leaf)
                Copy-Item -LiteralPath $setupSrc -Destination $isoLocal -Force
                $image = Mount-DiskImage -ImagePath $isoLocal -PassThru -ErrorAction Stop
                try {
                    $volume = @($image | Get-Volume | Where-Object { $_.DriveLetter })
                    if ($volume.Count -ne 1) { throw 'Could not resolve exactly one mounted SQL ISO volume.' }
                    $driveRoot = '{0}:\' -f $volume[0].DriveLetter
                    & robocopy.exe $driveRoot $media /E /NFL /NDL /NJH /NJS /R:2 /W:5 | Out-Null
                    $copyCode = $LASTEXITCODE
                }
                finally { Dismount-DiskImage -ImagePath $isoLocal -ErrorAction Stop | Out-Null }
            }
            else { & robocopy.exe $setupSrc $media /E /NFL /NDL /NJH /NJS /R:2 /W:5 | Out-Null; $copyCode = $LASTEXITCODE }
            if ($copyCode -ge 8) { throw "SQL media copy failed: robocopy exit $copyCode." }
            $setupExe = Join-Path $media 'setup.exe'
            if (-not (Test-Path -LiteralPath $setupExe -PathType Leaf)) { throw "setup.exe absent in $media" }

            if ($replaceCandidates.Count -gt 0) {
                Write-Warning "REPLACE: backing up/uninstalling $(($replaceCandidates.Server) -join ', '). Backups are NOT automatically restored."
                foreach ($old in $replaceCandidates) {
                    if (-not [bool]$Options.SkipBackup) {
                        if ($old.Status -ne 'Running') { Start-Service -Name $old.Service -ErrorAction Stop }
                        New-Item -ItemType Directory -Path $Options.BackupPath -Force | Out-Null
                        $backupInfo = Invoke-Sql $old.Server "SELECT CAST(SERVERPROPERTY('InstanceDefaultBackupPath') AS nvarchar(400)) AS p"
                        $backupDir = [string]$backupInfo.Rows[0].p
                        if ([string]::IsNullOrWhiteSpace($backupDir)) { $backupDir = [string]$Options.BackupPath }
                        $databaseTable = Invoke-Sql $old.Server 'SELECT name, state_desc FROM sys.databases WHERE database_id > 4'
                        $unavailable = @($databaseTable.Rows | Where-Object { $_.state_desc -ne 'ONLINE' })
                        if ($unavailable.Count -gt 0) { throw "Not all user databases on $($old.Server) are ONLINE. Stop Replace and assess backups." }
                        foreach ($database in $databaseTable.Rows) {
                            $name = [string]$database.name
                            $safeFileName = $name -replace '[\\/:*?"<>|]', '_'
                            $backupFile = Join-Path $backupDir ('{0}_{1}_{2:yyyyMMdd_HHmmss}_{3}.bak' -f $old.Instance, $safeFileName, (Get-Date), ([guid]::NewGuid().ToString('N')))
                            Write-Host "[$pc] Backing up [$name] from $($old.Server) ..."
                            $backupSql = "BACKUP DATABASE [{0}] TO DISK=N'{1}' WITH COPY_ONLY, INIT" -f $name.Replace(']', ']]'), $backupFile.Replace("'", "''")
                            $null = Invoke-Sql $old.Server $backupSql
                            if ($backupDir.TrimEnd('\') -ine ([string]$Options.BackupPath).TrimEnd('\')) { Copy-Item -LiteralPath $backupFile -Destination $Options.BackupPath -Force }
                        }
                    }
                    else { Write-Warning "SkipBackup selected for $($old.Server)." }
                    $bootstrapPattern = 'C:\Program Files\Microsoft SQL Server\{0}0\Setup Bootstrap\SQL*\setup.exe' -f $old.Major
                    $oldSetup = Get-ChildItem -Path $bootstrapPattern -ErrorAction SilentlyContinue | Select-Object -First 1 -ExpandProperty FullName
                    if (-not $oldSetup) { throw "SQL uninstall bootstrap absent for $($old.Server)." }
                    $code = Invoke-Proc $oldSetup ('/Q /ACTION=Uninstall /FEATURES=SQL /INSTANCENAME={0}' -f $old.Instance) $timeout
                    if ($code -in @(3010, 1641)) { return New-Result 'RebootRequired' "Uninstall of $($old.Server) requires a reboot. Restart and rerun." $null $true }
                    if ($code -ne 0) { throw "SQL uninstall failed for $($old.Server), exit $code. Assess SQL Setup logs; do not retry blindly." }
                }
            }
            else { Write-Host "[$pc] Existing SQL instances are left untouched." }

            $currentAccount = [Security.Principal.WindowsIdentity]::GetCurrent().Name
            $adminAccounts = @(@($Options.SysAdminAccounts) + @($currentAccount) | Select-Object -Unique)
            foreach ($account in $adminAccounts) {
                if ([string]::IsNullOrWhiteSpace([string]$account) -or [string]$account -match '["\r\n\x00]' -or ([string]$account).EndsWith('\')) { throw 'A SQL sysadmin account contains unsafe argument characters.' }
            }
            $adminArguments = ($adminAccounts | ForEach-Object { '"' + $_ + '"' }) -join ' '
            $saPassword = $null; $setupArguments = $null
            try {
                $saPassword = $SaCred.GetNetworkCredential().Password
                $setupArguments = ('/Q /ACTION=Install /IACCEPTSQLSERVERLICENSETERMS /SUPPRESSPRIVACYSTATEMENTNOTICE /FEATURES=SQLENGINE ' +
                    '/INSTANCENAME={0} /SQLSYSADMINACCOUNTS={1} /SECURITYMODE=SQL /SAPWD="{2}" ' +
                    '/TCPENABLED=1 /NPENABLED=1 /SQLSVCSTARTUPTYPE=Automatic /BROWSERSVCSTARTUPTYPE=Automatic /UPDATEENABLED=False') -f $instance, $adminArguments, $saPassword
                Write-Host "[$pc] Installing $targetName instance '$instance'; configuring the supplied NEW sa password."
                $code = Invoke-Proc $setupExe $setupArguments $timeout
            }
            finally { $saPassword = $null; $setupArguments = $null }
            if ($code -in @(3010, 1641)) { return New-Result 'RebootRequired' 'SQL installation requires a reboot. Stop dependent stages, restart and rerun.' $null $true }
            if ($code -ne 0) {
                throw ('SQL install failed (exit {0}). See C:\Program Files\Microsoft SQL Server\{1}0\Setup Bootstrap\Log\Summary.txt and matching component logs. Assess partial SQL state before retrying. A Driver 17 version check does not prove MSI source health.' -f $code, $targetMajor)
            }
        }
        else { Write-Host "[$pc] Exact SQL instance/version exists. SQL setup and sa password configuration skipped." }

        if (-not [bool]$Options.SkipOdbc) {
            $state = Get-OdbcState
            $installedVersion = $null
            $uninstallPaths = @('HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*', 'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*')
            $registrations = @(Get-ItemProperty -Path $uninstallPaths -ErrorAction SilentlyContinue | Where-Object { ([string](Get-RegistryValue $_ 'DisplayName')) -like "*$odbcName*" })
            foreach ($entry in $registrations) {
                $parsed = $null
                if ([version]::TryParse([string](Get-RegistryValue $entry 'DisplayVersion'), [ref]$parsed)) { if ($null -eq $installedVersion -or $parsed -lt $installedVersion) { $installedVersion = $parsed } }
            }
            $needOdbc = ($state.Missing.Count -gt 0)
            if ($null -ne $msiProperties) {
                $packageVersion = [version]$msiProperties.ProductVersion
                Write-Host "[$pc] Application ODBC media: $($msiProperties.ProductName) v$packageVersion; installed=$installedVersion"
                if ($null -eq $installedVersion -or $packageVersion -gt $installedVersion) { $needOdbc = $true }
                if ($needOdbc) {
                    $arguments = '/i "{0}" /qn /norestart IACCEPTMSODBCSQLLICENSETERMS=YES ADDLOCAL=ALL /L*v "C:\Temp\OdbcDriverInstall.log"' -f $odbcLocal
                    $code = Invoke-Proc (Join-Path $env:WINDIR 'System32\msiexec.exe') $arguments 15
                    if ($code -in @(3010, 1641)) { return New-Result 'RebootRequired' 'Application ODBC installation requires a reboot. Restart before dependent stages.' $null $true }
                    if ($code -ne 0) { throw "Application ODBC installation failed (exit $code). See C:\Temp\OdbcDriverInstall.log" }
                }
            }
            $state = Get-OdbcState
            if ($state.Missing.Count -gt 0) { throw "$odbcName is still missing for: $($state.Missing -join ', ')." }
            Write-Host "[$pc] $odbcName verified for both 32-bit and 64-bit."
        }
        else { Write-Host "[$pc] Application ODBC skipped. Configured SQL Driver 17 stage is separate." }

        $selected = @(Get-SqlInstances) | Where-Object { $_.Instance -eq $instance -and $_.Major -eq $targetMajor } | Select-Object -First 1
        if ($null -eq $selected) { return New-Result 'Failed' "Exact instance '$instance', major $targetMajor, absent after setup. Assess media/setup logs." $null $false }
        if ($selected.Status -ne 'Running') {
            Set-Service -Name $selected.Service -StartupType Automatic
            Start-Service -Name $selected.Service
            (Get-Service -Name $selected.Service).WaitForStatus('Running', [TimeSpan]::FromSeconds(60))
        }
        if ($selected.Instance -ne 'MSSQLSERVER') {
            $browser = Get-Service -Name SQLBrowser -ErrorAction SilentlyContinue
            if ($null -eq $browser) { return New-Result 'ActionRequired' 'SQL Browser absent. Configure named-instance discovery/ports before continuing.' $selected $false }
            Set-Service -Name SQLBrowser -StartupType Automatic
            if ($browser.Status -ne 'Running') { Start-Service -Name SQLBrowser }
            (Get-Service -Name SQLBrowser).WaitForStatus('Running', [TimeSpan]::FromSeconds(60))
        }
        $selected = @(Get-SqlInstances) | Where-Object { $_.Instance -eq $instance -and $_.Major -eq $targetMajor } | Select-Object -First 1
        if ($selected.Status -ne 'Running') { return New-Result 'Failed' 'SQL service did not reach Running.' $selected $false }
        if (-not $selected.MixedMode) { return New-Result 'ActionRequired' 'Existing instance uses Windows-only authentication. Enable Mixed Mode and configure existing sa separately; no existing password was reset.' $selected $false }
        $message = 'Exact SQL instance/version, running service and Mixed Mode verified.'
        if ([bool]$Options.SkipOdbc) { $message += ' Application ODBC validation explicitly skipped.' }
        else { $message += ' Requested application ODBC verified for both architectures.' }
        $message += ' A real database-login test remains necessary in the database stage.'
        return New-Result 'Ready' $message $selected $false
    }
    catch {
        if ($driver17Report.Status -in @('Preparing', 'Installing', 'RepairingSource')) { $driver17Report.Status = 'Failed' }
        elseif ($driver17Report.Status -eq 'VerifyingRecovery') { $driver17Report.Status = 'RecoveryVerificationFailed' }
        return New-Result 'Failed' $_.Exception.Message $null $false
    }
}

function Invoke-TargetSql {
    param([hashtable]$Options, [pscredential]$SaValue, [bool]$DetectionOnly)
    $outputs = if ($null -ne $session) {
        Invoke-Command -Session $session -ScriptBlock $sqlScript -ArgumentList @($Options, $SaValue, $DetectionOnly) -ErrorAction Stop
    }
    else { & $sqlScript $Options $SaValue $DetectionOnly }
    $report = @($outputs | Where-Object { $null -ne $_ -and $null -ne $_.PSObject.Properties['Status'] }) | Select-Object -Last 1
    if ($null -eq $report) { throw 'Target SQL operation did not return a structured report.' }
    return $report
}

function Write-Report {
    param($Report)
    Write-Host ''
    Write-Host ('[{0}] {1}: {2}' -f $Report.Computer, $Report.Status, $Report.Message)
    if ($Report.SqlInstance) { Write-Host "SQL instance: $($Report.SqlInstance)" }
    if ($Report.SqlOdbc17MsiPath) {
        Write-Host "SQL Driver 17: $($Report.SqlOdbc17Status); version=$($Report.SqlOdbc17Version)"
        Write-Host "SQL Driver 17 target source: $($Report.SqlOdbc17MsiPath)"
        if ($Report.SqlOdbc17ProductCode) { Write-Host "Driver 17 ProductCode: $($Report.SqlOdbc17ProductCode)" }
        if ($Report.SqlOdbc17PackageCode) { Write-Host "Driver 17 supplied PackageCode: $($Report.SqlOdbc17PackageCode)" }
        if ($Report.SqlOdbc17InstalledPackageCode) { Write-Host "Driver 17 registered PackageCode: $($Report.SqlOdbc17InstalledPackageCode)" }
        if ($Report.SqlOdbc17CachedMsi) { Write-Host "Retained source MSI: $($Report.SqlOdbc17CachedMsi)" }
        if ($Report.SqlOdbc17LogPath) { Write-Host "Driver 17 log: $($Report.SqlOdbc17LogPath)" }
        if ($Report.SqlOdbc17RepairRequested) { Write-Host "Driver 17 registered-cache identity verified after repair: $($Report.SqlOdbc17CacheVerified)" }
    }
    if ($Report.OdbcSkipped) { Write-Host 'Application ODBC check: skipped.' }
    else {
        Write-Host "Application ODBC 32-bit: $($Report.OdbcDrivers32 -join ', ')"
        Write-Host "Application ODBC 64-bit: $($Report.OdbcDrivers64 -join ', ')"
    }
    if ($Report.OdbcInventoryError) { Write-Warning "Application ODBC inventory error: $($Report.OdbcInventoryError)" }
    $Report.Instances | Format-Table Instance, Version, Edition, Status, MixedMode -AutoSize | Out-String | Write-Host
}

try {
    if ($env:OS -ne 'Windows_NT' -or -not [Environment]::Is64BitProcess -or $PSVersionTable.PSVersion.Major -ne 5 -or $PSVersionTable.PSVersion.Minor -lt 1) { throw 'Run in 64-bit Windows PowerShell 5.1, not PowerShell 7 or 32-bit PowerShell.' }
    if ($RepairSqlOdbc17 -and -not $SqlOdbc17Only) { throw 'Use -SqlOdbc17Only with -RepairSqlOdbc17. Recovery may not be combined with SQL Setup.' }

    $settings = $null
    if (-not [string]::IsNullOrWhiteSpace($SettingsFilePath) -and (Test-Path -LiteralPath $SettingsFilePath -PathType Leaf)) {
        $settingsResolvedPath = (Resolve-Path -LiteralPath $SettingsFilePath).ProviderPath
        try { $settings = Get-Content -LiteralPath $settingsResolvedPath -Raw | ConvertFrom-Json -ErrorAction Stop }
        catch { throw 'Settings.json could not be read/parsed. Check file access and JSON syntax (no comments or trailing commas). Raw parser output is suppressed because it can expose credentials.' }
        if ($null -eq $settings -or $settings -is [array] -or $settings -isnot [pscustomobject]) { throw 'Settings.json must contain one JSON object.' }
        Write-Host "[Settings] Loaded: $settingsResolvedPath"
    }
    elseif ($wasBound.ContainsKey('SettingsFilePath')) { throw "Settings file not found: $SettingsFilePath" }
    else { Write-Warning "Default settings file absent: $SettingsFilePath. Using explicit arguments/defaults." }

    $mapping = @{
        DeploymentMode = 'deploymentMode'; TargetMachine = 'machineName'; ConfigurationName = 'configurationName'
        SqlSetupPath = 'sqlSetupPath'; SqlVersion = 'sqlVersion'; SqlInstance = 'sqlInstance'
        OnVersionMismatch = 'onSqlVersionMismatch'; OdbcDriverMsiPath = 'odbcDriverMsiPath'
        OdbcDriverName = 'odbcDriverName'; OdbcVersion = 'odbcVersion'; SysAdminAccounts = 'sysAdminAccounts'
        BackupPath = 'sqlBackupPath'; MinFreeSpaceGB = 'sqlMinFreeSpaceGB'; TimeoutMinutes = 'sqlTimeoutMinutes'
        SqlOdbc17MsiPath = 'sqlOdbc17MsiPath'
    }
    foreach ($parameterName in $mapping.Keys) {
        if (-not $wasBound.ContainsKey($parameterName)) {
            $value = Get-SettingValue $settings $mapping[$parameterName]
            if ($null -ne $value) { Set-Variable -Name $parameterName -Value $value -Scope Script }
        }
    }
    $remoteUsername = [string](Get-SettingValue $settings 'remoteUsername')
    if ($DeploymentMode -notin @('Local', 'Remote')) { throw 'deploymentMode must be Local or Remote.' }
    if ($OnVersionMismatch -notin @('Stop', 'SideBySide', 'Replace')) { throw 'onSqlVersionMismatch must be Stop, SideBySide or Replace.' }
    if ($MinFreeSpaceGB -lt 1 -or $MinFreeSpaceGB -gt 1048576 -or $TimeoutMinutes -lt 1 -or $TimeoutMinutes -gt 1440) { throw 'Invalid free-space threshold or timeout in settings.' }
    if ($SqlOdbc17Only -and [string]::IsNullOrWhiteSpace($SqlOdbc17MsiPath)) { throw 'SqlOdbc17Only requires sqlOdbc17MsiPath or -SqlOdbc17MsiPath.' }

    $versionMap = @{ '2016' = 13; '2017' = 14; '2019' = 15; '2022' = 16 }
    $versionNames = @{ 13 = 'SQL Server 2016'; 14 = 'SQL Server 2017'; 15 = 'SQL Server 2019'; 16 = 'SQL Server 2022' }
    if ($SqlVersion) {
        if (-not $versionMap.ContainsKey([string]$SqlVersion)) { throw "Unsupported sqlVersion '$SqlVersion'." }
        $mappedMajor = [int]$versionMap[[string]$SqlVersion]
        if ($wasBound.ContainsKey('TargetMajorVersion')) {
            if ($wasBound.ContainsKey('SqlVersion') -and $TargetMajorVersion -ne $mappedMajor) { throw 'Explicit SqlVersion and TargetMajorVersion conflict.' }
        }
        else { $TargetMajorVersion = $mappedMajor }
    }
    $targetName = [string]$versionNames[$TargetMajorVersion]

    if ($SqlInstance) {
        $parts = @($SqlInstance -split '\\')
        if ($parts.Count -ne 2 -or -not $parts[0] -or -not $parts[1]) { throw 'sqlInstance must use HOST\INSTANCE format.' }
        $configuredInstanceName = $parts[1]
        if ($wasBound.ContainsKey('InstanceName')) {
            if ($wasBound.ContainsKey('SqlInstance') -and $InstanceName -ine $configuredInstanceName) { throw 'Explicit SqlInstance and InstanceName conflict.' }
        }
        else { $InstanceName = $configuredInstanceName }
        $expectedHost = if ($DeploymentMode -eq 'Remote') { $TargetMachine } else { $env:COMPUTERNAME }
        if ($wasBound.ContainsKey('TargetMachine') -and -not $wasBound.ContainsKey('SqlInstance')) { Write-Host '[Settings] Explicit TargetMachine overrides settings sqlInstance host; instance name retained.' }
        elseif ($expectedHost -and ($parts[0].Split('.')[0] -ine $expectedHost.Split('.')[0])) {
            if (-not ($DeploymentMode -eq 'Local' -and $parts[0] -in @('.', '(local)', 'localhost'))) { throw 'Target and sqlInstance host do not match.' }
        }
    }
    elseif (-not $wasBound.ContainsKey('InstanceName')) { $InstanceName = switch ($TargetMajorVersion) { 13 { 'SQL2K16' }; 14 { 'SQL2K17' }; 15 { 'SQL2K19' }; 16 { 'SQL2K22' } } }
    if ($InstanceName -notmatch '^[A-Za-z_][A-Za-z0-9_]{0,15}$') { throw 'InstanceName must be a safe instance name, not HOST\INSTANCE.' }

    if ($OdbcVersion) {
        if ([string]$OdbcVersion -notin @('17', '18')) { throw 'odbcVersion must be 17 or 18.' }
        $derivedDriverName = "ODBC Driver $OdbcVersion for SQL Server"
        if ($wasBound.ContainsKey('OdbcDriverName')) {
            if ($wasBound.ContainsKey('OdbcVersion') -and $OdbcDriverName -ine $derivedDriverName) { throw 'Explicit OdbcVersion and OdbcDriverName conflict.' }
        }
        elseif ($wasBound.ContainsKey('OdbcVersion') -or -not (Get-SettingValue $settings 'odbcDriverName')) { $OdbcDriverName = $derivedDriverName }
        elseif ($OdbcDriverName -ine $derivedDriverName) { throw 'odbcVersion and odbcDriverName conflict in settings.' }
    }
    foreach ($source in @($SqlSetupPath, $OdbcDriverMsiPath, $SqlOdbc17MsiPath)) {
        if ($source) {
            if ($source -match '["\r\n\x00]' -or $source -notmatch '^(?:[A-Za-z]:\\|\\\\[^\\]+\\[^\\]+)') { throw 'Installer sources must be full target-local or UNC filesystem paths, without quotes/line breaks. URLs and relative paths are unsupported.' }
        }
    }
    if ($SqlOdbc17MsiPath -and [IO.Path]::GetExtension($SqlOdbc17MsiPath) -ine '.msi') { throw 'SqlOdbc17MsiPath must name an MSI, not a directory, URL or EXE.' }
    if ($SqlSetupPath -and [IO.Path]::GetExtension($SqlSetupPath) -ieq '.exe') { throw 'SqlSetupPath must be full media (ISO or directory), not an EXE.' }
    if (-not [IO.Path]::IsPathRooted($BackupPath)) { throw 'BackupPath must be absolute.' }

    if ($DeploymentMode -eq 'Remote') {
        if ([string]::IsNullOrWhiteSpace($TargetMachine) -or $TargetMachine -match '[\s\[\]\(\)/\\]|^https?:') { throw 'A plain TargetMachine hostname/FQDN is required for Remote mode.' }
        if ($null -eq $Credential) {
            if ($NonInteractive) { throw 'Credential is required for non-interactive Remote mode.' }
            $prompt = @{ Message = "Windows credentials authorized for SQL provisioning on $TargetMachine" }
            if ($remoteUsername) { $prompt.UserName = $remoteUsername }
            $Credential = Get-Credential @prompt
            if ($null -eq $Credential) { throw 'Remote credential prompt cancelled.' }
        }
        $sessionParameters = @{ ComputerName = $TargetMachine; Credential = $Credential; Authentication = 'Negotiate'; ErrorAction = 'Stop' }
        if ($ConfigurationName) { $sessionParameters.ConfigurationName = $ConfigurationName }
        $session = New-PSSession @sessionParameters
        $remoteCompatible = Invoke-Command -Session $session -ScriptBlock { [Environment]::Is64BitProcess -and $PSVersionTable.PSVersion.Major -eq 5 -and $PSVersionTable.PSVersion.Minor -ge 1 }
        if (-not $remoteCompatible) { throw 'Target endpoint must run 64-bit Windows PowerShell 5.1.' }
    }

    $options = @{
        SqlSetupPath = $SqlSetupPath; InstanceName = $InstanceName; TargetMajorVersion = $TargetMajorVersion; TargetName = $targetName
        SysAdminAccounts = $SysAdminAccounts; OnVersionMismatch = $OnVersionMismatch; SkipBackup = [bool]$SkipBackup
        BackupPath = $BackupPath; OdbcDriverMsiPath = $OdbcDriverMsiPath; OdbcDriverName = $OdbcDriverName
        SkipOdbc = [bool]$SkipOdbc; MinFreeSpaceGB = $MinFreeSpaceGB; TimeoutMinutes = $TimeoutMinutes
        SqlOdbc17MsiPath = $SqlOdbc17MsiPath; SqlOdbc17Only = [bool]$SqlOdbc17Only
        RepairSqlOdbc17 = [bool]$RepairSqlOdbc17
    }
    Write-Host ''
    Write-Host "Deployment: $DeploymentMode; target: $(if ($DeploymentMode -eq 'Remote') { $TargetMachine } else { $env:COMPUTERNAME })"
    Write-Host "Requested SQL: $targetName; exact instance: $InstanceName; mismatch policy: $OnVersionMismatch"
    Write-Host "SQL source: $SqlSetupPath"
    Write-Host "Application ODBC: $OdbcDriverName; source: $OdbcDriverMsiPath; skipped: $([bool]$SkipOdbc)"
    Write-Host "Separate SQL Driver 17 target path: $SqlOdbc17MsiPath; Driver17-only mode: $([bool]$SqlOdbc17Only); explicit recovery: $([bool]$RepairSqlOdbc17)"

    $detection = Invoke-TargetSql -Options $options -SaValue $null -DetectionOnly $true
    Write-Report $detection
    if ($detection.Status -notin @('Skipped', 'NotInstalled')) { $result = $detection; $exitCode = 1 }
    elseif ($DetectOnly) { $result = $detection; $exitCode = 0 }
    else {
        if (-not $SqlOdbc17Only -and $detection.Status -eq 'NotInstalled') {
            if (-not $SqlSetupPath) { throw 'Configure sqlSetupPath or supply -SqlSetupPath for a new SQL install.' }
            if ($OnVersionMismatch -eq 'Replace') { Write-Warning 'Replace may uninstall ALL other-major SQL instances. Backups are not automatically restored.' }
            if ($null -eq $SaCredential) {
                if ($NonInteractive) { throw 'A new installation requires SaCredential in non-interactive mode.' }
                do {
                    $SaCredential = Get-Credential -UserName 'sa' -Message "NEW sa password for $InstanceName on $(if ($DeploymentMode -eq 'Remote') { $TargetMachine } else { $env:COMPUTERNAME }); not an existing Windows/hsi password"
                    if ($null -eq $SaCredential) { throw 'sa prompt cancelled.' }
                    $validPassword = Test-SaCredentialValue $SaCredential
                    if (-not $validPassword) { Write-Warning 'Use 8+ characters and 3 of upper/lower/digit/symbol; no double quote, newline, NUL or trailing backslash. SQL/OS policy may be stricter.' }
                } until ($validPassword)
            }
            elseif (-not (Test-SaCredentialValue $SaCredential)) { throw 'Supplied new sa password fails baseline complexity/argument safety.' }
            $saForInstall = $SaCredential
        }
        $result = Invoke-TargetSql -Options $options -SaValue $saForInstall -DetectionOnly $false
        Write-Report $result
        $exitCode = if ($result.RebootRequired) { 3010 } elseif ($result.Status -eq 'Ready') { 0 } else { 1 }
    }
    $result | Add-Member -NotePropertyName ExitCode -NotePropertyValue $exitCode -Force
    Write-Output $result
}
catch {
    Write-Warning "SQL step failed: $($_.Exception.Message)"
    $exitCode = 1
    Write-Output ([pscustomobject]@{
        Computer = $(if ($DeploymentMode -eq 'Remote') { $TargetMachine } else { $env:COMPUTERNAME })
        Status = 'Failed'; Message = $_.Exception.Message; RebootRequired = $false; ExitCode = 1
    })
}
finally {
    if ($null -ne $session) { Remove-PSSession -Session $session -ErrorAction SilentlyContinue }
    $saForInstall = $null
}
exit $exitCode
