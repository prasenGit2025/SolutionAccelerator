#requires -Version 5.1
<#
.SYNOPSIS
    Installs Hyland OnBase Studio x64 locally or remotely.

.DESCRIPTION
    Reads Settings.json on the controller and constructs the MSI source:

        buildPath + onBaseStudioMsiRelativePath

    Settings example:
        "onBaseStudioMsiRelativePath":
            "Install\\x64\\OnBase Studio\\Hyland OnBase Studio x64.msi",
        "onBaseStudioExtraMsiArgs": "ALLUSERS=1",
        "onBaseStudioTimeoutMinutes": 15

    Also reads:
        deploymentMode
        machineName
        configurationName
        remoteUsername

    Explicit command-line parameters override JSON values.

    An explicitly supplied -SourceMsiPath overrides path construction.

    The TARGET machine:
      1. Checks source accessibility.
      2. Stages the complete dedicated Studio installer folder.
      3. Verifies copied files using SHA256.
      4. Validates the Studio MSI identity and x64 package architecture.
      5. Optionally uninstalls matching existing installations.
      6. Installs the staged MSI silently.

    MSI bytes do not pass through the controller in Remote mode.

    The complete source MSI folder and its subfolders are copied.
    Therefore, SourceMsiPath must point into the dedicated Studio
    installer directory, not a general-purpose folder.

    Replacement media is staged BEFORE optional uninstallation.

    SHA256 comparison verifies copied bytes, NOT publisher authenticity.
    Digital-signature verification is not implemented in this revision.
    Use an approved, trusted source.

    This script runs the MSI directly, not the setup.exe bootstrapper.
    Required prerequisites must be installed separately beforehand.

    This script does not:
      - Install prerequisites.
      - Create databases or DSNs.
      - Configure Studio connections or credentials.
      - Activate or license the application.
      - Automatically reboot.
      - Implement reboot checkpoints/resume.
      - Modify the orchestrator.

    WinRM, endpoint configuration and share permissions must already work.

    Uninstall searches machine registrations and the target session
    user's HKCU. It does not enumerate every user's installations.

    Only entries with the exact staged MSI ProductName are selected
    for optional uninstall. Differently named historical releases are
    not automatically removed.

    A timeout does NOT forcibly kill Windows Installer.
    Inspect servicing state and logs before retrying.

.OUTPUTS
    Structured result and script exit code:

        0    = MSI completed successfully
        3010 = reboot required; stop dependent stages
        1    = failed or manual review required

    MsiExitCode preserves the last actual MSI outcome.

.EXAMPLE
    .\PowerShell\Install-OnBaseStudio.ps1 `
        -SettingsFilePath .\Settings.json `
        -Credential $cred

.EXAMPLE
    .\PowerShell\Install-OnBaseStudio.ps1 `
        -SettingsFilePath .\Settings.json `
        -Credential $cred `
        -UninstallExisting

.NOTES
    Default Settings.json location:
        <repository root>\Settings.json

    Assumed script location:
        <repository root>\PowerShell\Install-OnBaseStudio.ps1

    Target staging and logs:
        C:\Temp\OnBaseStudio\<run-id>\Media
        C:\Temp\OnBaseStudio\<run-id>\Logs

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

    [string]$StudioMsiRelativePath = '',

    # Optional compatibility override.
    [string]$SourceMsiPath = '',

    [string]$ExtraMsiArgs = '',

    [ValidateRange(1, 1440)]
    [int]$TimeoutMinutes = 15,

    # Explicit CLI authorization only; not read from JSON.
    [switch]$UninstallExisting
)

$ErrorActionPreference = 'Stop'

# Remember explicitly supplied parameters.
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
        Architecture   = $null
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

    function Release-ComObject {
        param($Object)

        if (
            $null -ne $Object -and
            [Runtime.InteropServices.Marshal]::IsComObject($Object)
        ) {
            [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject(
                $Object
            )
        }
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
        $summary = $null
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
                    try {
                        if ($null -ne $view) {
                            $null = $view.GetType().InvokeMember(
                                'Close',
                                'InvokeMethod',
                                $null,
                                $view,
                                $null
                            )
                        }
                    }
                    finally {
                        Release-ComObject $record
                        Release-ComObject $view
                    }
                }
            }

            $summary = $database.GetType().InvokeMember(
                'SummaryInformation',
                'GetProperty',
                $null,
                $database,
                @(0)
            )

            $values['Template'] = [string]$summary.GetType().InvokeMember(
                'Property',
                'GetProperty',
                $null,
                $summary,
                @(7)
            )

            return $values
        }
        finally {
            Release-ComObject $summary
            Release-ComObject $database
            Release-ComObject $installer
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

    function Get-MediaManifest {
        param([string]$Root)

        # Walk directories explicitly so reparse points are rejected
        # before their contents can be traversed.
        $queue = New-Object System.Collections.Queue

        $queue.Enqueue([pscustomobject]@{
            Path         = $Root
            RelativePath = ''
        })

        while ($queue.Count -gt 0) {
            $folder = $queue.Dequeue()

            foreach ($item in @(
                Get-ChildItem `
                    -LiteralPath $folder.Path `
                    -Force `
                    -ErrorAction Stop
            )) {
                if ($item.Attributes -band [IO.FileAttributes]::ReparsePoint) {
                    throw (
                        'Reparse points are unsupported in installer media: ' +
                        $item.FullName
                    )
                }

                $relativePath = if ($folder.RelativePath) {
                    Join-Path $folder.RelativePath $item.Name
                }
                else {
                    $item.Name
                }

                if ($item.PSIsContainer) {
                    [pscustomobject]@{
                        SourcePath   = $item.FullName
                        RelativePath = $relativePath
                        IsDirectory  = $true
                        Sha256       = $null
                    }

                    $queue.Enqueue([pscustomobject]@{
                        Path         = $item.FullName
                        RelativePath = $relativePath
                    })
                }
                else {
                    $hash = (
                        Get-FileHash `
                            -LiteralPath $item.FullName `
                            -Algorithm SHA256 `
                            -ErrorAction Stop
                    ).Hash

                    [pscustomobject]@{
                        SourcePath   = $item.FullName
                        RelativePath = $relativePath
                        IsDirectory  = $false
                        Sha256       = $hash
                    }
                }
            }
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

            if ($item.Attributes -band [IO.FileAttributes]::ReparsePoint) {
                throw "Unexpected reparse point in staged media: $path"
            }

            if ($entry.IsDirectory) {
                if (-not $item.PSIsContainer) {
                    throw "Expected a staged directory: $path"
                }

                continue
            }

            if ($item.PSIsContainer) {
                throw "Expected a staged file: $path"
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

        # All source access occurs ON THE TARGET.
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

            $candidate = 'StudioSource' +
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

        $sourceMsiItem = Get-Item -LiteralPath $readPath -Force

        if (
            $sourceMsiItem.Attributes -band
            [IO.FileAttributes]::ReparsePoint
        ) {
            throw 'The source MSI must not be a reparse point.'
        }

        # --------------------------------------------------------------------
        # 2. Inventory the dedicated installer directory.
        # --------------------------------------------------------------------
        $report.Stage = 'ReadingSource'

        $sourceFolder = Split-Path -Path $readPath -Parent
        $sourceFolderItem = Get-Item -LiteralPath $sourceFolder -Force

        if (
            -not $sourceFolderItem.PSIsContainer -or
            ($sourceFolderItem.Attributes -band [IO.FileAttributes]::ReparsePoint)
        ) {
            throw 'Source installer folder must be a normal directory.'
        }

        $manifest = @(Get-MediaManifest -Root $sourceFolder)

        $fileCount = @(
            $manifest | Where-Object { -not $_.IsDirectory }
        ).Count

        if ($fileCount -eq 0) {
            throw 'The source installer folder contains no files.'
        }

        # --------------------------------------------------------------------
        # 3. Create a protected per-run staging directory.
        # --------------------------------------------------------------------
        $report.Stage = 'Staging'

        $stageRoot = 'C:\Temp\OnBaseStudio'

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

        $report.Log = Join-Path $logDirectory 'StudioInstall.log'

        Write-Host (
            "[$env:COMPUTERNAME] Staging Studio installer folder " +
            "from $sourceFolder"
        )

        foreach ($entry in $manifest) {
            $destination = Join-Path `
                $mediaDirectory `
                $entry.RelativePath

            if ($entry.IsDirectory) {
                New-Item `
                    -ItemType Directory `
                    -Path $destination `
                    -Force |
                    Out-Null

                continue
            }

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

            # Verify that this source file did not change during copying.
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
                -Algorithm SHA256 `
                -ErrorAction Stop
        ).Hash

        Write-Host (
            "[$env:COMPUTERNAME] Studio installer media copied " +
            'and verified using SHA256.'
        )

        if ($driveName) {
            Remove-PSDrive `
                -Name $driveName `
                -Force `
                -ErrorAction Stop

            $driveName = $null
        }

        # --------------------------------------------------------------------
        # 4. Validate the staged MSI before optional uninstallation.
        # --------------------------------------------------------------------
        $report.Stage = 'ValidatingPackage'

        $package = Read-MsiIdentity -Path $report.LocalMsi

        # Deliberately narrow product-name safety guard.
        # An unexpected name stops execution; there is no automatic bypass.
        if (
            $package.ProductName -notmatch
            '(?i)^Hyland OnBase Studio(?: x64)?$'
        ) {
            throw (
                "Unexpected MSI ProductName '$($package.ProductName)'. " +
                'Expected Hyland OnBase Studio or ' +
                'Hyland OnBase Studio x64. No uninstall/install was started.'
            )
        }

        $version = $null

        if (
            -not [version]::TryParse(
                $package.ProductVersion,
                [ref]$version
            )
        ) {
            throw 'The Studio MSI ProductVersion is invalid.'
        }

        $productGuid = [guid]::Empty

        if (
            -not [guid]::TryParse(
                $package.ProductCode,
                [ref]$productGuid
            ) -or
            $productGuid -eq [guid]::Empty
        ) {
            throw 'The Studio MSI ProductCode is invalid.'
        }

        $architecture = ([string]$package.Template -split ';')[0]

        if ($architecture -ine 'x64') {
            throw (
                'This installer requires the x64 Studio MSI. ' +
                "Found MSI template '$($package.Template)'."
            )
        }

        $report.ProductName = $package.ProductName
        $report.ProductVersion = $package.ProductVersion
        $report.ProductCode = $productGuid.ToString('B').ToUpperInvariant()
        $report.Architecture = $architecture

        Write-Host (
            "[$env:COMPUTERNAME] Package: " +
            "$($report.ProductName) $($report.ProductVersion) " +
            "[$($report.Architecture)]"
        )

        if (Test-Path -LiteralPath $busyKey) {
            throw (
                'Windows Installer became busy. ' +
                'No Studio uninstall/install was started.'
            )
        }

        # --------------------------------------------------------------------
        # 5. Optional uninstall.
        #
        # Uses exact ProductName matching.
        # Does not use Win32_Product.
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
                    'Studio registration found.'
                )
            }

            foreach ($product in $products) {
                $uninstallLog = Join-Path $logDirectory (
                    'StudioUninstall-' +
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
                        'Studio uninstall requires a reboot. ' +
                        'Replacement installation was NOT started. ' +
                        'Inspect state after reboot before rerunning.'
                    )

                    return $report
                }

                if ($code -notin @(0, 1605)) {
                    throw (
                        "Studio uninstall failed with MSI code $code. " +
                        'Replacement installation was NOT started. ' +
                        "See $uninstallLog"
                    )
                }
            }
        }

        # --------------------------------------------------------------------
        # 6. Recheck staged media and install.
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
            "[$env:COMPUTERNAME] Installing OnBase Studio x64 " +
            "(timeout: $timeoutMinutes minutes)..."
        )

        $code = Invoke-Msi -Arguments $arguments
        $report.MsiExitCode = $code

        if ($code -notin @(0, 3010, 1641)) {
            throw (
                "Studio installation failed with MSI code $code. " +
                "See $($report.Log)"
            )
        }

        $report.RebootRequired = ($code -in @(3010, 1641))

        if ($report.RebootRequired) {
            $report.Stage = 'InstallRebootRequired'
            $report.ExitCode = 3010
            $report.Message = (
                'Studio MSI succeeded; reboot required. ' +
                'Stop dependent stages and reassess after reboot.'
            )
        }
        else {
            $report.Stage = 'Completed'
            $report.ExitCode = 0
            $report.Message = (
                'Studio MSI completed successfully. ' +
                'Application connections, licensing and login ' +
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
                'Settings file could not be read/parsed. ' +
                'Check JSON syntax and access. Raw parser output ' +
                'is suppressed to avoid exposing secrets.'
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

    # Explicit CLI parameters override JSON.
    $mapping = @{
        DeploymentMode        = 'deploymentMode'
        TargetMachine         = 'machineName'
        ConfigurationName     = 'configurationName'
        BuildPath             = 'buildPath'
        StudioMsiRelativePath  = 'onBaseStudioMsiRelativePath'
        ExtraMsiArgs          = 'onBaseStudioExtraMsiArgs'
        TimeoutMinutes        = 'onBaseStudioTimeoutMinutes'
    }

    foreach ($name in $mapping.Keys) {
        if (-not $explicitParameters.ContainsKey($name)) {
            $value = Get-Setting `
                -Object $settings `
                -Name $mapping[$name]

            if ($null -ne $value) {
                Set-Variable -Name $name -Value $value -Scope Script
            }
        }
    }

    if ($DeploymentMode -notin @('Local', 'Remote')) {
        throw 'deploymentMode must be Local or Remote.'
    }

    if ($TimeoutMinutes -lt 1 -or $TimeoutMinutes -gt 1440) {
        throw 'onBaseStudioTimeoutMinutes must be between 1 and 1440.'
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
    #
    #   buildPath + onBaseStudioMsiRelativePath
    #
    # No controller-side access to the build share occurs here.
    # ------------------------------------------------------------------------
    if (-not $explicitParameters.ContainsKey('SourceMsiPath')) {
        if (
            [string]::IsNullOrWhiteSpace($BuildPath) -or
            [string]::IsNullOrWhiteSpace($StudioMsiRelativePath)
        ) {
            throw (
                'Set BOTH buildPath and onBaseStudioMsiRelativePath ' +
                'in Settings.json.'
            )
        }

        $BuildPath = $BuildPath.TrimEnd('\')
        $relative = $StudioMsiRelativePath.Replace('/', '\')

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
                'onBaseStudioMsiRelativePath must be relative to ' +
                'buildPath, without dot/dot-dot segments or ' +
                'invalid characters.'
            )
        }

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
                    'Windows credentials for OnBase Studio ' +
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

        # With a named endpoint, rely on its execution identity unless
        # ShareCredential was explicitly supplied.
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
    Write-Host 'OnBase Studio x64 installation'
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

    # ------------------------------------------------------------------------
    # Report the outcome.
    # ------------------------------------------------------------------------
    Write-Host ''
    Write-Host (
        "[$($result.Computer)] $($result.Stage): $($result.Message)"
    )

    if ($result.ProductName) {
        Write-Host (
            "Package: $($result.ProductName) " +
            "$($result.ProductVersion) [$($result.Architecture)]"
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
        'OnBase Studio step failed: ' +
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