#requires -Version 5.1
<#
.SYNOPSIS
    Installs Hyland Unity Client locally or remotely using Settings.json.

.DESCRIPTION
    Constructs the source MSI path from these Settings.json properties:

        buildPath
        unityMsiRelativePath

    Example JSON properties:
        "buildPath": "\\\\onbase.net\\shares\\Builds\\OnBase\\Current\\25_1\\25,1,28,1000",
        "unityMsiRelativePath": "Install\\Unity Client\\Hyland Unity Client.msi",
        "unityExtraMsiArgs": "ALLUSERS=1",
        "unityTimeoutMinutes": 15

    Also reads:
        deploymentMode
        machineName
        configurationName
        remoteUsername

    Explicit command-line parameters override JSON values.

    An explicitly supplied -SourceMsiPath overrides path construction.

    The controller reads Settings.json. The target machine:
      1. Checks source accessibility.
      2. Copies the MSI directly from the share into:
         C:\Temp\UnityClient\<run-id>\
      3. Verifies that source and copied MSI SHA256 hashes match.
      4. Optionally uninstalls existing Unity Client.
      5. Installs the staged MSI silently.

    Replacement media is staged BEFORE optional uninstallation.

    "Download" means a filesystem copy from the UNC share, not an
    HTTP/HTTPS download. In Remote mode, MSI bytes do not pass through
    the controller.

    SHA256 comparison verifies matching bytes, NOT publisher authenticity.
    Use an approved source and the standard Unity Client MSI.

    WinRM, endpoint configuration and share permissions must already work.
    This script does not configure those prerequisites.

    A configured RunAs endpoint can provide share access using its
    execution identity. Alternatively, -ShareCredential can establish
    an explicit target-side filesystem connection.

    Uninstall searches machine registrations and the target session
    user's HKCU. It does not enumerate every user's installations.

    A timeout does NOT forcibly kill Windows Installer.
    Inspect servicing state and logs before retrying.

.OUTPUTS
    Structured result and script exit code:

        0    = MSI completed successfully
        3010 = reboot required; stop dependent stages
        1    = failed or manual review required

    MsiExitCode preserves the last actual MSI outcome.

.EXAMPLE
    .\PowerShell\Install-Unity.ps1 `
        -SettingsFilePath .\Settings.json

.EXAMPLE
    .\PowerShell\Install-Unity.ps1 `
        -SettingsFilePath .\Settings.json `
        -Credential $cred `
        -UninstallExisting

.NOTES
    Default settings location:
        <repository root>\Settings.json

    Assumes this script is stored in:
        <repository root>\PowerShell\Install-Unity.ps1

    No automatic reboot or checkpoint/resume implementation.

    Reviewed as text only.
    Not parsed or executed in Windows PowerShell here.
#>

[CmdletBinding()]
param(
    [ValidateSet('Local', 'Remote')]
    [string]$DeploymentMode = 'Local',

    [string]$TargetMachine = '',

    # Optional compatibility override.
    # Otherwise constructed from buildPath + unityMsiRelativePath.
    [string]$SourceMsiPath = '',

    [string]$ExtraMsiArgs = '',

    [pscredential]$Credential,

    [pscredential]$ShareCredential,

    [string]$ConfigurationName = '',

    [switch]$UninstallExisting,

    [ValidateRange(1, 1440)]
    [int]$TimeoutMinutes = 15,

    [Alias('SettingsPath')]
    [string]$SettingsFilePath = (
        Join-Path (Split-Path -Parent $PSScriptRoot) 'Settings.json'
    ),

    [string]$BuildPath = '',

    [string]$UnityMsiRelativePath = ''
)

$ErrorActionPreference = 'Stop'

# Remember explicitly supplied parameters.
$bound = @{}
foreach ($key in $PSBoundParameters.Keys) {
    $bound[$key] = $true
}

function Get-Setting {
    param(
        $Object,
        [string]$Name
    )

    if ($null -ne $Object) {
        $property = $Object.PSObject.Properties[$Name]

        if ($null -ne $property -and $null -ne $property.Value) {
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
        Sha256         = $null
        Log            = $null
        UninstallLogs  = @()
        Leftovers      = @()
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
                    "PID $($process.Id) was NOT killed. Inspect the " +
                    'installer state and logs before retrying or rebooting.'
                )
            }

            $process.WaitForExit()
            return [int]$process.ExitCode
        }
        finally {
            $process.Dispose()
        }
    }

    try {
        # --------------------------------------------------------------------
        # Preflight on the target.
        # --------------------------------------------------------------------
        if (
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
                'Windows Installer reports an operation in progress. ' +
                'Inspect it before retrying.'
            )
        }

        # This source is accessed FROM THE TARGET, not the controller.
        $readPath = $sourceMsi

        # Optional share connection without placing the password in a
        # net.exe command line.
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

            $candidate = 'UnitySource' + [guid]::NewGuid().ToString('N')

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

        $sourceHash = (
            Get-FileHash `
                -LiteralPath $readPath `
                -Algorithm SHA256 `
                -ErrorAction Stop
        ).Hash

        # --------------------------------------------------------------------
        # Stage the MSI BEFORE optional uninstallation.
        # --------------------------------------------------------------------
        $report.Stage = 'Staging'
        $root = 'C:\Temp\UnityClient'

        foreach ($folder in @('C:\Temp', $root)) {
            if (Test-Path -LiteralPath $folder) {
                $item = Get-Item -LiteralPath $folder -Force

                if (
                    -not $item.PSIsContainer -or
                    ($item.Attributes -band [IO.FileAttributes]::ReparsePoint)
                ) {
                    throw (
                        'Refusing non-directory/reparse-point staging ' +
                        "path: $folder"
                    )
                }
            }
            else {
                New-Item `
                    -ItemType Directory `
                    -Path $folder `
                    -ErrorAction Stop |
                    Out-Null
            }
        }

        $workDir = Join-Path $root ([guid]::NewGuid().ToString('N'))

        # Protect the new run directory: Administrators and SYSTEM only.
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

        $report.LocalMsi = Join-Path $workDir (
            Split-Path $sourceMsi -Leaf
        )

        $report.Log = Join-Path $workDir 'UnityInstall.log'

        Write-Host (
            "[$env:COMPUTERNAME] Copying $sourceMsi " +
            "to $($report.LocalMsi)"
        )

        Copy-Item `
            -LiteralPath $readPath `
            -Destination $report.LocalMsi `
            -Force `
            -ErrorAction Stop

        $localHash = (
            Get-FileHash `
                -LiteralPath $report.LocalMsi `
                -Algorithm SHA256 `
                -ErrorAction Stop
        ).Hash

        $sourceHashAfter = (
            Get-FileHash `
                -LiteralPath $readPath `
                -Algorithm SHA256 `
                -ErrorAction Stop
        ).Hash

        if (
            $localHash -ne $sourceHash -or
            $sourceHashAfter -ne $sourceHash
        ) {
            throw (
                'Source changed during staging or the copied MSI hash ' +
                'does not match. No uninstall/install was started.'
            )
        }

        $report.Sha256 = $localHash

        Write-Host (
            "[$env:COMPUTERNAME] MSI copy verified using SHA256."
        )

        if ($driveName) {
            Remove-PSDrive `
                -Name $driveName `
                -Force `
                -ErrorAction Stop

            $driveName = $null
        }

        if (Test-Path -LiteralPath $busyKey) {
            throw (
                'Windows Installer became busy. ' +
                'No Unity uninstall/install was started.'
            )
        }

        # --------------------------------------------------------------------
        # Optional removal of existing Unity Client installations.
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
                        $displayName = [string](
                            Get-ObjectProperty $_ 'DisplayName'
                        )

                        $productCode = [string](
                            Get-ObjectProperty $_ 'PSChildName'
                        )

                        $displayName -like 'Hyland Unity Client*' -and
                        $productCode -match '^\{[a-fA-F0-9-]{36}\}$'
                    } |
                    Select-Object -ExpandProperty PSChildName -Unique
            )

            if ($products.Count -eq 0) {
                Write-Host (
                    "[$env:COMPUTERNAME] No matching Unity Client " +
                    'installation found.'
                )
            }

            foreach ($product in $products) {
                $log = Join-Path $workDir (
                    'UnityUninstall-' + $product.Trim('{}') + '.log'
                )

                $report.UninstallLogs += $log

                Write-Host (
                    "[$env:COMPUTERNAME] Uninstalling $product"
                )

                $code = Invoke-Msi (
                    '/x {0} /qn /norestart /L*v "{1}"' -f $product, $log
                )

                $report.MsiExitCode = $code

                if ($code -in @(3010, 1641)) {
                    $report.Stage = 'UninstallRebootRequired'
                    $report.ExitCode = 3010
                    $report.RebootRequired = $true
                    $report.Message = (
                        'Uninstall requires a reboot. Replacement ' +
                        'installation was NOT started. Inspect state ' +
                        'after reboot before rerunning.'
                    )

                    return $report
                }

                if ($code -notin @(0, 1605)) {
                    throw (
                        "Unity uninstall failed with MSI code $code. " +
                        'Replacement installation was NOT started. ' +
                        "See $log"
                    )
                }
            }
        }

        # --------------------------------------------------------------------
        # Install the staged MSI.
        # --------------------------------------------------------------------
        $currentHash = (
            Get-FileHash `
                -LiteralPath $report.LocalMsi `
                -Algorithm SHA256 `
                -ErrorAction Stop
        ).Hash

        if ($currentHash -ne $report.Sha256) {
            throw (
                'Staged MSI changed before installation. ' +
                'Installation refused.'
            )
        }

        $report.Stage = 'Installing'

        Write-Host (
            "[$env:COMPUTERNAME] Installing Unity Client " +
            "(timeout: $timeoutMinutes minutes)..."
        )

        $arguments = '/i "{0}" /qn /norestart /L*v "{1}" {2}' -f `
            $report.LocalMsi, $report.Log, $extraArgs

        $code = Invoke-Msi $arguments
        $report.MsiExitCode = $code

        if ($code -notin @(0, 3010, 1641)) {
            throw (
                "Unity installation failed with MSI code $code. " +
                "See $($report.Log)"
            )
        }

        $report.RebootRequired = ($code -in @(3010, 1641))

        $report.ExitCode = if ($report.RebootRequired) {
            3010
        }
        else {
            0
        }

        $report.Stage = if ($report.RebootRequired) {
            'InstallRebootRequired'
        }
        else {
            'Completed'
        }

        $report.Message = if ($report.RebootRequired) {
            'MSI succeeded; reboot required. Stop dependent stages and reassess after reboot.'
        }
        else {
            'MSI completed successfully. Application configuration/login validation remains separate.'
        }

        try {
            $report.Leftovers = @(
                Get-CimInstance `
                    Win32_Process `
                    -Filter "Name='DeploymentWizard.exe' OR Name='obunity.exe'" `
                    -ErrorAction Stop |
                    Select-Object -ExpandProperty Name
            )
        }
        catch {
            Write-Warning (
                'Could not enumerate Unity-related processes. ' +
                'The MSI outcome is retained.'
            )
        }

        return $report
    }
    catch {
        $report.Message = $_.Exception.Message
        $report.ExitCode = 1

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
# CONTROLLER: LOAD SETTINGS AND RESOLVE THE SOURCE PATH
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
                'syntax and file access. Raw parser output is suppressed ' +
                'to avoid exposing secrets.'
            )
        }

        if ($settings -isnot [pscustomobject]) {
            throw 'Settings.json must contain one JSON object.'
        }

        Write-Host "[Settings] Loaded: $SettingsFilePath"
    }
    elseif (
        $bound.ContainsKey('SettingsFilePath') -or
        -not $bound.ContainsKey('SourceMsiPath')
    ) {
        throw "Settings file not found: $SettingsFilePath"
    }

    # Explicit command-line parameters take precedence over JSON.
    $mapping = @{
        DeploymentMode      = 'deploymentMode'
        TargetMachine       = 'machineName'
        ConfigurationName   = 'configurationName'
        BuildPath           = 'buildPath'
        UnityMsiRelativePath = 'unityMsiRelativePath'
        ExtraMsiArgs        = 'unityExtraMsiArgs'
        TimeoutMinutes      = 'unityTimeoutMinutes'
    }

    foreach ($name in $mapping.Keys) {
        if (-not $bound.ContainsKey($name)) {
            $value = Get-Setting $settings $mapping[$name]

            if ($null -ne $value) {
                Set-Variable -Name $name -Value $value -Scope Script
            }
        }
    }

    if ($DeploymentMode -notin @('Local', 'Remote')) {
        throw 'deploymentMode must be Local or Remote.'
    }

    if ($TimeoutMinutes -lt 1 -or $TimeoutMinutes -gt 1440) {
        throw 'unityTimeoutMinutes must be between 1 and 1440.'
    }

    # Extra arguments are trusted MSI property assignments.
    # Do not accept additional command-line switches.
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
    # Construct:
    #     Settings.json buildPath
    #       +
    #     Settings.json unityMsiRelativePath
    #
    # No hardcoded relative installer path.
    # No separate full unityMsiPath setting.
    # ------------------------------------------------------------------------
    if (-not $bound.ContainsKey('SourceMsiPath')) {
        if (
            [string]::IsNullOrWhiteSpace($BuildPath) -or
            [string]::IsNullOrWhiteSpace($UnityMsiRelativePath)
        ) {
            throw (
                'Set BOTH buildPath and unityMsiRelativePath ' +
                'in Settings.json.'
            )
        }

        $BuildPath = $BuildPath.TrimEnd('\')
        $relative = $UnityMsiRelativePath.Replace('/', '\')

        if (
            $BuildPath -notmatch '^\\\\[^\\]+\\[^\\]+' -or
            $BuildPath -match '["<>|?*\x00-\x1F]'
        ) {
            throw (
                'buildPath must be an absolute UNC folder without ' +
                'wildcard or invalid characters.'
            )
        }

        if (
            [IO.Path]::IsPathRooted($relative) -or
            $relative -match '[:"<>|?*\x00-\x1F]' -or
            $relative -match '(^|\\)\.\.?($|\\)'
        ) {
            throw (
                'unityMsiRelativePath must be relative to buildPath, ' +
                'without dot/dot-dot segments or invalid characters.'
            )
        }

        # Pure path construction; no controller-side share access.
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
            'path to an MSI file. URLs are unsupported.'
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
                    'Windows credentials for Unity installation on ' +
                    $TargetMachine
                )
            }

            $user = [string](
                Get-Setting $settings 'remoteUsername'
            )

            if ($user) {
                $prompt.UserName = $user
            }

            $Credential = Get-Credential @prompt

            if ($null -eq $Credential) {
                throw 'Credential prompt cancelled.'
            }
        }

        # Without a named endpoint, reuse the supplied remote credential
        # for explicit share access unless a separate one was provided.
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
    Write-Host "Deployment mode : $DeploymentMode"
    Write-Host "Target machine  : $displayTarget"
    Write-Host "Unity MSI source: $SourceMsiPath"
    Write-Host "Endpoint        : $ConfigurationName"
    Write-Host "Uninstall first : $([bool]$UninstallExisting)"
    Write-Host ''

    # ------------------------------------------------------------------------
    # Execute locally or on the remote VM.
    # ------------------------------------------------------------------------
    $scriptArgs = @(
        $SourceMsiPath,
        $ShareCredential,
        $ExtraMsiArgs,
        $TimeoutMinutes,
        [bool]$UninstallExisting
    )

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

    # Keep only the structured target result.
    $result = $output |
        Where-Object {
            $null -ne $_ -and $_.PSObject.Properties['ExitCode']
        } |
        Select-Object -Last 1

    if ($null -eq $result) {
        throw 'No structured installation result returned from the target.'
    }

    # ------------------------------------------------------------------------
    # Report the outcome.
    # ------------------------------------------------------------------------
    Write-Host ''
    Write-Host (
        "[$($result.Computer)] $($result.Stage): $($result.Message)"
    )

    Write-Host (
        "Overall exit: $($result.ExitCode) | " +
        "Last MSI exit: $($result.MsiExitCode)"
    )

    if ($result.LocalMsi) {
        Write-Host "Staged MSI: $($result.LocalMsi)"
    }

    if ($result.Log) {
        Write-Host (
            "Install log (if installation started): $($result.Log)"
        )
    }

    foreach ($log in $result.UninstallLogs) {
        Write-Host "Uninstall log: $log"
    }

    if (@($result.Leftovers).Count -gt 0) {
        Write-Warning (
            'Related processes present (informational, not proof ' +
            "of an interactive MSI): $($result.Leftovers -join ', ')"
        )
    }

    Write-Output $result

    # Preserve reboot-required status for the orchestrator.
    exit ([int]$result.ExitCode)
}
catch {
    Write-Warning "Unity step failed: $($_.Exception.Message)"

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