#requires -Version 5.1
<#
.SYNOPSIS
    Installs Hyland OnBase Configuration locally or remotely.

.DESCRIPTION
    Reads Settings.json on the controller and constructs the MSI source:

        buildPath + onBaseConfigurationMsiRelativePath

    Explicit command-line parameters override JSON.

    An explicitly supplied -SourceMsiPath overrides path construction.

    The TARGET machine:
      1. Checks source accessibility.
      2. Stages the installer folder, including companion files.
      3. Verifies copied files using SHA256.
      4. Checks the staged MSI product identity.
      5. Optionally uninstalls matching existing installations.
      6. Installs the staged MSI silently.

    MSI bytes do not pass through the controller in Remote mode.

    The complete installer folder is staged so external CAB files and
    other package-relative files remain available.

    Replacement media is staged BEFORE optional uninstallation.

    SHA256 comparison verifies copied bytes, not publisher authenticity.
    Use an approved, trusted build share.

    Prerequisites, the database and System DSN must be prepared separately.
    No unverified product-specific MSI properties are invented here.
    This script does not configure login, activation or database licensing.

    WinRM, the selected endpoint and share permissions must already work.
    This script does not change remoting configuration.

    Optional uninstall searches machine registrations and the target
    session user's HKCU, not every user's per-user installations.

    A timeout does NOT forcibly kill Windows Installer.
    Inspect servicing state and logs before retrying.

.OUTPUTS
    Structured result and script exit code:

        0    = MSI completed successfully
        3010 = reboot required; stop dependent stages
        1    = failed or manual review required

.EXAMPLE
    .\PowerShell\Install-Configuration.ps1 `
        -SettingsFilePath .\Settings.json `
        -Credential $cred

.EXAMPLE
    .\PowerShell\Install-Configuration.ps1 `
        -SettingsFilePath .\Settings.json `
        -Credential $cred `
        -UninstallExisting

.NOTES
    Default Settings.json location:
        <repository root>\Settings.json

    Assumed script location:
        <repository root>\PowerShell\Install-Configuration.ps1

    Target staging and logs:
        C:\Temp\OnBaseConfiguration\<run-id>\

    No automatic reboot or checkpoint/resume implementation.

    Reviewed as text only.
    Not parsed or executed in Windows PowerShell here.
#>

[CmdletBinding()]
param(
    [ValidateSet('Local', 'Remote')]
    [string]$DeploymentMode = 'Local',

    [string]$TargetMachine = '',

    [pscredential]$Credential,

    [pscredential]$ShareCredential,

    [string]$ConfigurationName = '',

    [Alias('SettingsPath')]
    [string]$SettingsFilePath = (
        Join-Path (Split-Path -Parent $PSScriptRoot) 'Settings.json'
    ),

    [string]$BuildPath = '',

    [string]$ConfigurationMsiRelativePath = '',

    # Optional compatibility override.
    [string]$SourceMsiPath = '',

    [string]$ExtraMsiArgs = '',

    [ValidateRange(1, 1440)]
    [int]$TimeoutMinutes = 15,

    # Explicit CLI authorization only; not loaded from JSON.
    [switch]$UninstallExisting
)

$ErrorActionPreference = 'Stop'

$explicitParameters = @{}
foreach ($key in $PSBoundParameters.Keys) {
    $explicitParameters[$key] = $true
}

function Get-Setting {
    param(
        $Object,
        [string]$Name
    )

    if ($null -ne $Object) {
        $property = $Object.PSObject.Properties[$Name]

        if ($null -ne $property) {
            return $property.Value
        }
    }

    return $null
}

# ============================================================================
# TARGET-SIDE INSTALLATION
# ============================================================================

$installScript = {
    param(
        [string]$sourceMsi,
        [pscredential]$shareCred,
        [string]$extraArgs,
        [int]$timeoutMinutes,
        [bool]$uninstallExisting
    )

    $ErrorActionPreference = 'Stop'
    $driveName = $null

    $report = [pscustomobject]@{
        Computer       = $env:COMPUTERNAME
        Stage          = 'Preflight'
        ExitCode       = 1
        MsiExitCode    = $null
        RebootRequired = $false
        Message        = ''
        SourceMsiPath  = $sourceMsi
        LocalMsi       = $null
        WorkDirectory  = $null
        ProductName    = $null
        ProductVersion = $null
        ProductCode    = $null
        MsiSha256      = $null
        Log            = $null
        UninstallLogs  = @()
    }

    function Get-ObjectProperty {
        param(
            $Object,
            [string]$Name
        )

        if ($null -ne $Object) {
            $property = $Object.PSObject.Properties[$Name]

            if ($null -ne $property) {
                return $property.Value
            }
        }

        return $null
    }

    function Invoke-Msi {
        param([string]$Arguments)

        $process = Start-Process `
            -FilePath (Join-Path $env:WINDIR 'System32\msiexec.exe') `
            -ArgumentList $Arguments `
            -PassThru `
            -ErrorAction Stop

        try {
            $null = $process.Handle

            if (-not $process.WaitForExit($timeoutMinutes * 60000)) {
                throw (
                    "Installer wait exceeded $timeoutMinutes minutes. " +
                    "PID $($process.Id) was NOT killed. Inspect installer " +
                    'state and logs before retrying or rebooting.'
                )
            }

            $process.WaitForExit()
            return [int]$process.ExitCode
        }
        finally {
            $process.Dispose()
        }
    }

    function Read-MsiIdentity {
        param([string]$Path)

        $installer = $null
        $database = $null
        $values = @{}

        try {
            $installer = New-Object -ComObject WindowsInstaller.Installer

            $database = $installer.GetType().InvokeMember(
                'OpenDatabase',
                'InvokeMethod',
                $null,
                $installer,
                @($Path, 0)
            )

            foreach ($name in @(
                'ProductName',
                'ProductVersion',
                'ProductCode'
            )) {
                $view = $null
                $record = $null

                try {
                    $query = (
                        "SELECT ``Value`` FROM ``Property`` " +
                        "WHERE ``Property``='$name'"
                    )

                    $view = $database.GetType().InvokeMember(
                        'OpenView',
                        'InvokeMethod',
                        $null,
                        $database,
                        @($query)
                    )

                    $null = $view.GetType().InvokeMember(
                        'Execute',
                        'InvokeMethod',
                        $null,
                        $view,
                        $null
                    )

                    $record = $view.GetType().InvokeMember(
                        'Fetch',
                        'InvokeMethod',
                        $null,
                        $view,
                        $null
                    )

                    if ($null -eq $record) {
                        throw "MSI property '$name' is missing."
                    }

                    $values[$name] = [string]$record.GetType().InvokeMember(
                        'StringData',
                        'GetProperty',
                        $null,
                        $record,
                        @(1)
                    )
                }
                finally {
                    if ($null -ne $view) {
                        try {
                            $null = $view.GetType().InvokeMember(
                                'Close',
                                'InvokeMethod',
                                $null,
                                $view,
                                $null
                            )
                        }
                        finally {
                            if ($null -ne $record) {
                                [void][Runtime.InteropServices.Marshal]::
                                    FinalReleaseComObject($record)
                            }

                            [void][Runtime.InteropServices.Marshal]::
                                FinalReleaseComObject($view)
                        }
                    }
                    elseif ($null -ne $record) {
                        [void][Runtime.InteropServices.Marshal]::
                            FinalReleaseComObject($record)
                    }
                }
            }

            return $values
        }
        finally {
            if ($null -ne $database) {
                [void][Runtime.InteropServices.Marshal]::
                    FinalReleaseComObject($database)
            }

            if ($null -ne $installer) {
                [void][Runtime.InteropServices.Marshal]::
                    FinalReleaseComObject($installer)
            }
        }
    }

    function Assert-NormalDirectory {
        param([string]$Path)

        if (Test-Path -LiteralPath $Path) {
            $item = Get-Item -LiteralPath $Path -Force

            if (
                -not $item.PSIsContainer -or
                ($item.Attributes -band [IO.FileAttributes]::ReparsePoint)
            ) {
                throw "Refusing non-directory or reparse-point path: $Path"
            }
        }
        else {
            New-Item `
                -ItemType Directory `
                -Path $Path `
                -ErrorAction Stop |
                Out-Null
        }
    }

    function Assert-StagedFiles {
        param(
            [object[]]$Manifest,
            [string]$Destination
        )

        foreach ($entry in $Manifest) {
            $path = Join-Path $Destination $entry.RelativePath
            $item = Get-Item -LiteralPath $path -Force -ErrorAction Stop

            if (
                $item.PSIsContainer -or
                ($item.Attributes -band [IO.FileAttributes]::ReparsePoint)
            ) {
                throw "Unexpected staged file type: $path"
            }

            $hash = (
                Get-FileHash `
                    -LiteralPath $path `
                    -Algorithm SHA256 `
                    -ErrorAction Stop
            ).Hash

            if ($hash -ne $entry.Sha256) {
                throw "Staged file hash mismatch: $path"
            }
        }
    }

    try {
        # --------------------------------------------------------------------
        # 1. Target preflight.
        # --------------------------------------------------------------------
        if (
            $env:OS -ne 'Windows_NT' -or
            -not [Environment]::Is64BitProcess -or
            $PSVersionTable.PSVersion.Major -ne 5 -or
            $PSVersionTable.PSVersion.Minor -lt 1
        ) {
            throw 'Target must run 64-bit Windows PowerShell 5.1.'
        }

        $principal = [Security.Principal.WindowsPrincipal]::new(
            [Security.Principal.WindowsIdentity]::GetCurrent()
        )

        if (
            -not $principal.IsInRole(
                [Security.Principal.WindowsBuiltInRole]::Administrator
            )
        ) {
            throw 'Administrative execution is required on the target.'
        }

        $busyKey = (
            'HKLM:\SOFTWARE\Microsoft\Windows\' +
            'CurrentVersion\Installer\InProgress'
        )

        if (Test-Path -LiteralPath $busyKey) {
            throw (
                'Windows Installer reports another operation in progress. ' +
                'Inspect it before retrying.'
            )
        }

        # All source access below occurs on the target.
        $readPath = $sourceMsi

        # Optional explicit share connection.
        # No password is passed to a native command line.
        if (
            $sourceMsi.StartsWith('\\') -and
            $null -ne $shareCred
        ) {
            $parts = $sourceMsi.Substring(2).Split('\')

            if ($parts.Count -lt 3) {
                throw 'UNC source must identify an MSI below a share.'
            }

            $shareRoot = '\\{0}\{1}' -f $parts[0], $parts[1]
            $relative = $parts[2..($parts.Count - 1)] -join '\'

            $candidate = 'ConfigurationSource' +
                [guid]::NewGuid().ToString('N')

            New-PSDrive `
                -Name $candidate `
                -PSProvider FileSystem `
                -Root $shareRoot `
                -Credential $shareCred `
                -Scope Local `
                -ErrorAction Stop |
                Out-Null

            $driveName = $candidate
            $readPath = '{0}:\{1}' -f $driveName, $relative
        }

        if (-not (Test-Path -LiteralPath $readPath -PathType Leaf)) {
            throw (
                "MSI is not accessible FROM ${env:COMPUTERNAME}: " +
                $sourceMsi
            )
        }

        # --------------------------------------------------------------------
        # 2. Inventory the dedicated installer folder.
        # --------------------------------------------------------------------
        $report.Stage = 'ReadingSource'

        $sourceFolder = Split-Path -Path $readPath -Parent
        $sourceItem = Get-Item -LiteralPath $sourceFolder -Force

        if (
            -not $sourceItem.PSIsContainer -or
            ($sourceItem.Attributes -band [IO.FileAttributes]::ReparsePoint)
        ) {
            throw 'Source installer folder must be a normal directory.'
        }

        $sourceRoot = $sourceItem.FullName.TrimEnd('\')

        $items = @(
            Get-ChildItem `
                -LiteralPath $sourceFolder `
                -Recurse `
                -Force `
                -ErrorAction Stop
        )

        $manifest = @()

        foreach ($item in $items) {
            if ($item.Attributes -band [IO.FileAttributes]::ReparsePoint) {
                throw "Reparse points are unsupported in media: $($item.FullName)"
            }

            if (-not $item.PSIsContainer) {
                $relativePath = $item.FullName.Substring(
                    $sourceRoot.Length
                ).TrimStart('\')

                $hash = (
                    Get-FileHash `
                        -LiteralPath $item.FullName `
                        -Algorithm SHA256 `
                        -ErrorAction Stop
                ).Hash

                $manifest += [pscustomobject]@{
                    SourcePath   = $item.FullName
                    RelativePath = $relativePath
                    Sha256       = $hash
                }
            }
        }

        if ($manifest.Count -eq 0) {
            throw 'The source installer folder contains no files.'
        }

        # --------------------------------------------------------------------
        # 3. Create a protected per-run target directory.
        # --------------------------------------------------------------------
        $report.Stage = 'Staging'

        $stageRoot = 'C:\Temp\OnBaseConfiguration'

        Assert-NormalDirectory 'C:\Temp'
        Assert-NormalDirectory $stageRoot

        $workDir = Join-Path $stageRoot (
            [guid]::NewGuid().ToString('N')
        )

        $acl = [Security.AccessControl.DirectorySecurity]::new()
        $acl.SetAccessRuleProtection($true, $false)

        $admins = [Security.Principal.SecurityIdentifier]::new(
            'S-1-5-32-544'
        )

        $acl.SetOwner($admins)

        foreach ($sidText in @('S-1-5-32-544', 'S-1-5-18')) {
            $sid = [Security.Principal.SecurityIdentifier]::new($sidText)

            $rule = [Security.AccessControl.FileSystemAccessRule]::new(
                $sid,
                'FullControl',
                'ContainerInherit, ObjectInherit',
                'None',
                'Allow'
            )

            [void]$acl.AddAccessRule($rule)
        }

        [void][IO.Directory]::CreateDirectory($workDir, $acl)

        # Separate media from logs to avoid filename collisions.
        $mediaDirectory = Join-Path $workDir 'Media'
        $logDirectory = Join-Path $workDir 'Logs'

        New-Item -ItemType Directory -Path $mediaDirectory |
            Out-Null

        New-Item -ItemType Directory -Path $logDirectory |
            Out-Null

        $report.WorkDirectory = $workDir
        $report.LocalMsi = Join-Path $mediaDirectory (
            Split-Path $sourceMsi -Leaf
        )
        $report.Log = Join-Path $logDirectory 'ConfigurationInstall.log'

        Write-Host (
            "[$env:COMPUTERNAME] Staging installer folder " +
            "from $sourceFolder"
        )

        # Create subdirectories, including any empty ones.
        foreach ($item in $items) {
            if ($item.PSIsContainer) {
                $relativePath = $item.FullName.Substring(
                    $sourceRoot.Length
                ).TrimStart('\')

                $destination = Join-Path $mediaDirectory $relativePath

                New-Item `
                    -ItemType Directory `
                    -Path $destination `
                    -Force |
                    Out-Null
            }
        }

        foreach ($entry in $manifest) {
            $destination = Join-Path `
                $mediaDirectory `
                $entry.RelativePath

            $parent = Split-Path -Path $destination -Parent

            New-Item `
                -ItemType Directory `
                -Path $parent `
                -Force |
                Out-Null

            Copy-Item `
                -LiteralPath $entry.SourcePath `
                -Destination $destination `
                -Force `
                -ErrorAction Stop

            # Ensure the corresponding source did not change during copy.
            $sourceHashAfter = (
                Get-FileHash `
                    -LiteralPath $entry.SourcePath `
                    -Algorithm SHA256 `
                    -ErrorAction Stop
            ).Hash

            if ($sourceHashAfter -ne $entry.Sha256) {
                throw (
                    'Source file changed during staging: ' +
                    $entry.SourcePath
                )
            }
        }

        Assert-StagedFiles `
            -Manifest $manifest `
            -Destination $mediaDirectory

        $report.MsiSha256 = (
            Get-FileHash `
                -LiteralPath $report.LocalMsi `
                -Algorithm SHA256
        ).Hash

        Write-Host (
            "[$env:COMPUTERNAME] Installer media copied and verified."
        )

        if ($driveName) {
            Remove-PSDrive -Name $driveName -Force -ErrorAction Stop
            $driveName = $null
        }

        # --------------------------------------------------------------------
        # 4. Validate MSI identity before any uninstall.
        # --------------------------------------------------------------------
        $report.Stage = 'ValidatingPackage'

        $package = Read-MsiIdentity -Path $report.LocalMsi

        if ($package.ProductName -ine 'Hyland OnBase Configuration') {
            throw (
                "Unexpected MSI product '$($package.ProductName)'. " +
                'Expected Hyland OnBase Configuration. ' +
                'No uninstall/install was started.'
            )
        }

        $version = $null
        if (
            -not [version]::TryParse(
                $package.ProductVersion,
                [ref]$version
            )
        ) {
            throw 'The MSI ProductVersion is invalid.'
        }

        $productGuid = [guid]::Empty
        if (
            -not [guid]::TryParse(
                $package.ProductCode,
                [ref]$productGuid
            ) -or
            $productGuid -eq [guid]::Empty
        ) {
            throw 'The MSI ProductCode is invalid.'
        }

        $report.ProductName = $package.ProductName
        $report.ProductVersion = $package.ProductVersion
        $report.ProductCode = $productGuid.ToString('B').ToUpperInvariant()

        Write-Host (
            "[$env:COMPUTERNAME] Package: " +
            "$($report.ProductName) $($report.ProductVersion)"
        )

        if (Test-Path -LiteralPath $busyKey) {
            throw (
                'Windows Installer became busy. ' +
                'No Configuration uninstall/install was started.'
            )
        }

        # --------------------------------------------------------------------
        # 5. Optional uninstall.
        #
        # Match the exact MSI product name, not unrelated products
        # containing the word "Configuration".
        # --------------------------------------------------------------------
        if ($uninstallExisting) {
            $report.Stage = 'Uninstalling'

            $keys = @(
                'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*'
                'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*'
                'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*'
            )

            $products = @(
                Get-ItemProperty `
                    -Path $keys `
                    -ErrorAction SilentlyContinue |
                    Where-Object {
                        $name = [string](
                            Get-ObjectProperty $_ 'DisplayName'
                        )

                        $code = [string](
                            Get-ObjectProperty $_ 'PSChildName'
                        )

                        $guid = [guid]::Empty

                        $name -ieq $package.ProductName -and
                        [guid]::TryParse($code, [ref]$guid) -and
                        $guid -ne [guid]::Empty
                    } |
                    Select-Object -ExpandProperty PSChildName -Unique
            )

            if ($products.Count -eq 0) {
                Write-Host (
                    "[$env:COMPUTERNAME] No matching existing " +
                    'OnBase Configuration registration found.'
                )
            }

            foreach ($product in $products) {
                $uninstallLog = Join-Path $logDirectory (
                    'ConfigurationUninstall-' +
                    $product.Trim('{}') +
                    '.log'
                )

                $report.UninstallLogs += $uninstallLog

                Write-Host (
                    "[$env:COMPUTERNAME] Uninstalling $product"
                )

                $arguments = (
                    '/x {0} /qn /norestart /L*v "{1}"' -f
                    $product, $uninstallLog
                )

                $code = Invoke-Msi -Arguments $arguments
                $report.MsiExitCode = $code

                if ($code -in @(3010, 1641)) {
                    $report.Stage = 'UninstallRebootRequired'
                    $report.ExitCode = 3010
                    $report.RebootRequired = $true
                    $report.Message = (
                        'Uninstall requires a reboot. Replacement ' +
                        'installation was NOT started. Inspect target ' +
                        'state after reboot before rerunning.'
                    )

                    return $report
                }

                if ($code -notin @(0, 1605)) {
                    throw (
                        "Uninstall failed with MSI code $code. " +
                        'Replacement installation was NOT started. ' +
                        "See $uninstallLog"
                    )
                }
            }
        }

        # --------------------------------------------------------------------
        # 6. Install the staged MSI.
        # --------------------------------------------------------------------
        Assert-StagedFiles `
            -Manifest $manifest `
            -Destination $mediaDirectory

        $report.Stage = 'Installing'

        $arguments = (
            '/i "{0}" /qn /norestart /L*v "{1}" {2}' -f
            $report.LocalMsi, $report.Log, $extraArgs
        )

        Write-Host (
            "[$env:COMPUTERNAME] Installing OnBase Configuration " +
            "(timeout: $timeoutMinutes minutes)..."
        )

        $code = Invoke-Msi -Arguments $arguments
        $report.MsiExitCode = $code

        if ($code -notin @(0, 3010, 1641)) {
            throw (
                "Installation failed with MSI code $code. " +
                "See $($report.Log)"
            )
        }

        $report.RebootRequired = ($code -in @(3010, 1641))

        if ($report.RebootRequired) {
            $report.Stage = 'InstallRebootRequired'
            $report.ExitCode = 3010
            $report.Message = (
                'MSI succeeded; reboot required. Stop dependent stages ' +
                'and reassess after reboot.'
            )
        }
        else {
            $report.Stage = 'Completed'
            $report.ExitCode = 0
            $report.Message = (
                'OnBase Configuration MSI completed successfully. ' +
                'DSN selection, activation and application-login ' +
                'validation remain separate.'
            )
        }

        return $report
    }
    catch {
        $report.ExitCode = 1
        $report.Message = $_.Exception.Message

        return $report
    }
    finally {
        if ($driveName) {
            Remove-PSDrive `
                -Name $driveName `
                -Force `
                -ErrorAction SilentlyContinue
        }
    }
}

# ============================================================================
# CONTROLLER: LOAD SETTINGS AND RESOLVE PARAMETERS
# ============================================================================

try {
    if (
        $env:OS -ne 'Windows_NT' -or
        -not [Environment]::Is64BitProcess -or
        $PSVersionTable.PSVersion.Major -ne 5 -or
        $PSVersionTable.PSVersion.Minor -lt 1
    ) {
        throw 'Run this script in 64-bit Windows PowerShell 5.1.'
    }

    $settings = $null

    if (Test-Path -LiteralPath $SettingsFilePath -PathType Leaf) {
        try {
            $settings = Get-Content `
                -LiteralPath $SettingsFilePath `
                -Raw |
                ConvertFrom-Json -ErrorAction Stop
        }
        catch {
            throw (
                'Settings file could not be read/parsed. Check JSON ' +
                'syntax and access. Raw parser output is suppressed ' +
                'to avoid exposing secrets.'
            )
        }

        if ($settings -isnot [pscustomobject]) {
            throw 'Settings.json must contain one JSON object.'
        }

        Write-Host "[Settings] Loaded: $SettingsFilePath"
    }
    elseif (
        $explicitParameters.ContainsKey('SettingsFilePath') -or
        -not $explicitParameters.ContainsKey('SourceMsiPath')
    ) {
        throw "Settings file not found: $SettingsFilePath"
    }

    $mapping = @{
        DeploymentMode               = 'deploymentMode'
        TargetMachine                = 'machineName'
        ConfigurationName            = 'configurationName'
        BuildPath                    = 'buildPath'
        ConfigurationMsiRelativePath = 'onBaseConfigurationMsiRelativePath'
        ExtraMsiArgs                 = 'onBaseConfigurationExtraMsiArgs'
        TimeoutMinutes               = 'onBaseConfigurationTimeoutMinutes'
    }

    foreach ($name in $mapping.Keys) {
        if (-not $explicitParameters.ContainsKey($name)) {
            $value = Get-Setting -Object $settings -Name $mapping[$name]

            if ($null -ne $value) {
                Set-Variable -Name $name -Value $value -Scope Script
            }
        }
    }

    if ($DeploymentMode -notin @('Local', 'Remote')) {
        throw 'deploymentMode must be Local or Remote.'
    }

    if ($TimeoutMinutes -lt 1 -or $TimeoutMinutes -gt 1440) {
        throw (
            'onBaseConfigurationTimeoutMinutes must be ' +
            'between 1 and 1440.'
        )
    }

    # ExtraMsiArgs is trusted administrative input.
    # Reject control characters and additional command-line switches.
    if (
        $ExtraMsiArgs -match '[\r\n\x00]' -or
        $ExtraMsiArgs -match '(?i)(?:^|\s)[/-]'
    ) {
        throw (
            'ExtraMsiArgs accepts trusted MSI property assignments only, ' +
            'not additional switches or control characters.'
        )
    }

    # ------------------------------------------------------------------------
    # Combine the two named settings.
    # ------------------------------------------------------------------------
    if (-not $explicitParameters.ContainsKey('SourceMsiPath')) {
        if (
            [string]::IsNullOrWhiteSpace($BuildPath) -or
            [string]::IsNullOrWhiteSpace($ConfigurationMsiRelativePath)
        ) {
            throw (
                'Set BOTH buildPath and ' +
                'onBaseConfigurationMsiRelativePath in Settings.json.'
            )
        }

        $BuildPath = $BuildPath.TrimEnd('\')
        $relative = $ConfigurationMsiRelativePath.Replace('/', '\')

        if (
            $BuildPath -notmatch '^\\\\[^\\]+\\[^\\]+' -or
            $BuildPath -match '[:"<>|?*\x00-\x1F]'
        ) {
            throw 'buildPath must be an absolute UNC folder.'
        }

        if (
            [IO.Path]::IsPathRooted($relative) -or
            $relative -match '[:"<>|?*\x00-\x1F]' -or
            $relative -match '(^|\\)\.\.?($|\\)'
        ) {
            throw (
                'onBaseConfigurationMsiRelativePath must be relative ' +
                'to buildPath, without dot/dot-dot segments or ' +
                'invalid characters.'
            )
        }

        # String/path construction only.
        # Does not test or copy the source on the controller.
        $SourceMsiPath = [IO.Path]::Combine($BuildPath, $relative)
    }

    if (
        [string]::IsNullOrWhiteSpace($SourceMsiPath) -or
        $SourceMsiPath -match '["<>|?*\x00-\x1F]' -or
        $SourceMsiPath -notmatch '^(?:[A-Za-z]:\\|\\\\[^\\]+\\[^\\]+\\)' -or
        [IO.Path]::GetExtension($SourceMsiPath) -ine '.msi'
    ) {
        throw (
            'Resolved source must be an absolute target-local or UNC ' +
            'path to an MSI. HTTP/HTTPS URLs are unsupported.'
        )
    }

    # ------------------------------------------------------------------------
    # Resolve remote connection and credentials.
    # ------------------------------------------------------------------------
    if ($DeploymentMode -eq 'Remote') {
        if (
            [string]::IsNullOrWhiteSpace($TargetMachine) -or
            $TargetMachine -match '[\s\[\]()/\\:]'
        ) {
            throw (
                'Remote mode requires a plain machineName/TargetMachine ' +
                'hostname or FQDN.'
            )
        }

        if ($null -eq $Credential) {
            $prompt = @{
                Message = (
                    'Windows credentials for OnBase Configuration ' +
                    "installation on $TargetMachine"
                )
            }

            $username = [string](
                Get-Setting -Object $settings -Name 'remoteUsername'
            )

            if ($username) {
                $prompt.UserName = $username
            }

            $Credential = Get-Credential @prompt

            if ($null -eq $Credential) {
                throw 'Credential prompt cancelled.'
            }
        }

        if ($null -eq $ShareCredential -and -not $ConfigurationName) {
            $ShareCredential = $Credential
        }
    }

    $displayTarget = if ($DeploymentMode -eq 'Remote') {
        $TargetMachine
    }
    else {
        $env:COMPUTERNAME
    }

    Write-Host ''
    Write-Host 'OnBase Configuration installation'
    Write-Host "Deployment mode : $DeploymentMode"
    Write-Host "Target machine  : $displayTarget"
    Write-Host "Source MSI      : $SourceMsiPath"
    Write-Host "Endpoint        : $ConfigurationName"
    Write-Host "Uninstall first : $([bool]$UninstallExisting)"
    Write-Host ''

    $scriptArgs = @(
        $SourceMsiPath,
        $ShareCredential,
        $ExtraMsiArgs,
        $TimeoutMinutes,
        [bool]$UninstallExisting
    )

    # ------------------------------------------------------------------------
    # Execute installation ON the target.
    # ------------------------------------------------------------------------
    if ($DeploymentMode -eq 'Remote') {
        $connection = @{
            ComputerName   = $TargetMachine
            Credential     = $Credential
            Authentication = 'Negotiate'
            ScriptBlock    = $installScript
            ArgumentList   = $scriptArgs
            ErrorAction    = 'Stop'
        }

        if ($ConfigurationName) {
            $connection.ConfigurationName = $ConfigurationName
        }

        $output = Invoke-Command @connection
    }
    else {
        $output = & $installScript @scriptArgs
    }

    $result = $output |
        Where-Object {
            $null -ne $_ -and $_.PSObject.Properties['ExitCode']
        } |
        Select-Object -Last 1

    if ($null -eq $result) {
        throw 'The target did not return a structured installation result.'
    }

    Write-Host ''
    Write-Host (
        "[$($result.Computer)] $($result.Stage): $($result.Message)"
    )

    if ($result.ProductName) {
        Write-Host (
            "Package: $($result.ProductName) " +
            "$($result.ProductVersion)"
        )
    }

    Write-Host "Overall exit: $($result.ExitCode)"
    Write-Host "Last MSI exit: $($result.MsiExitCode)"

    if ($result.LocalMsi) {
        Write-Host "Staged MSI: $($result.LocalMsi)"
    }

    if ($result.Log) {
        Write-Host (
            'Install log (if installation started): ' +
            $result.Log
        )
    }

    foreach ($log in $result.UninstallLogs) {
        Write-Host "Uninstall log: $log"
    }

    Write-Output $result

    # Preserve reboot-required status for the orchestrator.
    exit ([int]$result.ExitCode)
}
catch {
    Write-Warning (
        'OnBase Configuration step failed: ' +
        $_.Exception.Message
    )

    Write-Output ([pscustomobject]@{
        Computer = $(if ($DeploymentMode -eq 'Remote') {
            $TargetMachine
        }
        else {
            $env:COMPUTERNAME
        })
        Stage          = 'ControllerFailed'
        ExitCode       = 1
        MsiExitCode    = $null
        RebootRequired = $false
        Message        = $_.Exception.Message
    })

    exit 1
}