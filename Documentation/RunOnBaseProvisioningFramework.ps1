Step1. Session setup

cd C:\Users\PRROY\source\OnBaseProvisioningFramework
$t     = 'RDV-009989.hylandqa.net'
$bs    = [string][char]92
$cred  = Get-Credential -UserName 'ONBASE\prroy' -Message 'ONBASE credentials'
$hsi   = Get-Credential -UserName 'hsi' -Message 'OnBase hsi password'
$build = $bs + $bs + 'onbase.net\shares\Builds\OnBase\Current\25_1\25,1,28,1000'
$sqlSrc = $bs + $bs + 'onbase.net\shares\Public\SQLServer\sql_server_2022_developer.iso'
==============================================================================

Step2. Full Installation: SqlServer

.\PowerShell\Install-OnBase.ps1 -DeploymentMode Remote -TargetMachine $t -Credential $cred -ConfigurationName UnityInstall -BuildPath $build -SqlVersion 2022 -SqlSetupPath $sqlSrc -OnSqlVersionMismatch SideBySide -OdbcVersion 18 -HsiCredential $hsi -UninstallExisting

A. Incase of Failure Read the setup log from the VM.

Invoke-Command -ComputerName $t -Credential $cred -ConfigurationName UnityInstall -ScriptBlock {
    $log = 'C:\Program Files\Microsoft SQL Server\160\Setup Bootstrap\Log'
    Get-Content (Join-Path $log 'Summary.txt') -TotalCount 40
    $last = Get-ChildItem $log -Directory | Sort-Object LastWriteTime | Select-Object -Last 1
    "--- Detail.txt errors ($($last.Name)) ---"
    Select-String -Path (Join-Path $last.FullName 'Detail.txt') -Pattern 'Exception|error|failed' | Select-Object -Last 15 | ForEach-Object Line
}

B. Once SQL Server is installed, RDV-009989 must be restarted

Invoke-Command -ComputerName $t -Credential $cred -ConfigurationName UnityInstall -ScriptBlock {
    Restart-Computer -Force
}

C. verify connectivity and service health on RDV-009989 after the reboot completes:

i) Test basic WinRM and PowerShell remoting connectivity


Test-WSMan -ComputerName $t
Expected: Returns product version and protocol version (OS: Windows Server / Windows 11).

wsmid           : http://schemas.dmtf.org/wbem/wsman/identity/1/wsmanidentity.xsd
ProtocolVersion : http://schemas.dmtf.org/wbem/wsman/1/wsman.xsd
ProductVendor   : Microsoft Corporation
ProductVersion  : OS: 0.0.0 SP: 0.0 Stack: 3.0

ii) Full remoting and session configuration check
Verify that the UnityInstall JEA/endpoint configuration responds properly to your credentials:

Invoke-Command -ComputerName $t -Credential $cred -ConfigurationName UnityInstall -ScriptBlock {
    [PSCustomObject]@{
        ComputerName = $env:COMPUTERNAME
        CurrentUser  = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
        Uptime       = (Get-CimInstance Win32_OperatingSystem).LastBootUpTime
    }
}
Expected: Returns the computer name, current user, and system uptime.
    ComputerName   : RDV-009989
    CurrentUser    : ONBASE\prroy
    Uptime         : 10/1/2026 11:11:43 PM
    PSComputerName : RDV-009989.hylandqa.net
    RunspaceId     : c6d0d379-0c75-4b8c-96ee-18b7c5f399f3

iii) Verify SQL Server services came back online
Check that both SQL2K22 and SQLSERVER2K19 are running after the reboot:

Invoke-Command -ComputerName $t -Credential $cred -ConfigurationName UnityInstall -ScriptBlock {
    Get-Service -Name '*SQL*' | Select-Object Name, DisplayName, Status, StartType
}

Expected: Returns the status of all SQL Server-related services.
Name           : MSSQL$SQL2K22
DisplayName    : SQL Server (SQL2K22)
Status         : Running
StartType      : Automatic
PSComputerName : RDV-009989.hylandqa.net
RunspaceId     : 534ba9e3-30b7-466e-979c-379d632209a5

and more......


iv) Optional: One-liner loop (wait until the machine is fully up)
If the reboot is still in progress, run this loop to wait until WinRM starts responding:


do { 
    Write-Host "Waiting for $t to respond..." -ForegroundColor Yellow
    Start-Sleep -Seconds 10 
} until (Test-WSMan -ComputerName $t -ErrorAction SilentlyContinue)
Write-Host "$t is back online!" -ForegroundColor Green

Expected: The loop will continue until the machine responds to WinRM requests, indicating it is back online.
RDV-009989.hylandqa.net is back online!


=====================================================================
commands for these steps 

Step 3 (Initialize-OnBaseDatabase.ps1): Creates the OnBase database on RDV-009989\SQL2K22, configures the hsi user with your provided $hsi credentials, and builds the 32-bit system DSN using ODBC Driver 18.

Step 4 (Install-Unity.ps1): Installs the Hyland Unity Client (25.1.28.1000) and removes prior versions (since -UninstallExisting was passed).

Step 5 (Configure-ServiceLocation.ps1) & Step 6 (Verify-Installation.ps1): Points service locations and performs final sanity checks.
==========================================================================


Prerequisites (ensure session variables are set)

$t = 'RDV-009989.hylandqa.net'
$cred = Get-Credential -UserName 'ONBASE\prroy'
$build = 'https://onbase.net/shares/Builds/OnBase/Current/25_1/25,1,28,1000'
$hsi = Get-Credential -UserName 'hsi' -Message 'Password for the hsi SQL user' //it will be wstinol

=======================================================================================
Output returned by Commands executed for the above steps..
------------------------------------------
PS C:\Users\PRROY\source\OnBaseProvisioningFramework> $t = 'RDV-009989.hylandqa.net'
PS C:\Users\PRROY\source\OnBaseProvisioningFramework> $cred = Get-Credential -UserName 'ONBASE\prroy'

cmdlet Get-Credential at command pipeline position 1
Supply values for the following parameters:
Message: wstinol
PS C:\Users\PRROY\source\OnBaseProvisioningFramework> $build = 'https://onbase.net/shares/Builds/OnBase/Current/25_1/25,1,28,1000'
PS C:\Users\PRROY\source\OnBaseProvisioningFramework> $hsi = Get-Credential -UserName 'hsi' -Message 'wstinol'
===================================================================================

For Step 3 (Initialize-OnBaseDatabase.ps1) we have to setup UNC path for dbUtils.

# Build the UNC network path
$build = '\\' + 'onbase.net\shares\Builds\OnBase\Current\25_1\25,1,28,1000'
$dbUtils = $build + '\utilities\FoundationDBUtils'

# Check what the variable actually contains
$dbUtils                                       // it should show the full UNC path to the FoundationDBUtils directory [\\onbase.net\shares\Builds\OnBase\Current\25_1\25,1,28,1000\utilities\FoundationDBUtils]
# Test network share access
Test-Path $dbUtils                            // it should return True if the network share is accessible

Once Test-Path returns True, execute Step 3:

.\PowerShell\Initialize-OnBaseDatabase.ps1 `
    -DeploymentMode Remote `
    -TargetMachine $t `
    -Credential $cred `
    -ConfigurationName UnityInstall `
    -DbUtilsSourcePath $dbUtils `
    -SqlInstance 'RDV-009989\SQL2K22' `
    -OdbcDriver 'ODBC Driver 18 for SQL Server' `
    -HsiCredential $hsi `
    -DatabaseName 'OnBase
============================================================

Step 4: Install Unity Client (Install-Unity.ps1)
Copies the Unity Client MSI from the build share to C:\Temp\UnitySetup on the target VM, removes previous versions (if -UninstallExisting is specified), and performs an unattended installation:


# 1. Define the UNC path to the Unity Client MSI

$build = '\\' + 'onbase.net\shares\Builds\OnBase\Current\25_1\25,1,28,1000'
$unityMsi = Join-Path $build 'Install\Unity Client\Hyland Unity Client.msi'

# 2. Verify the MSI path is accessible

Test-Path $unityMsi

# 3. Execute Unity Client installation


.\PowerShell\Install-Unity.ps1 `
    -DeploymentMode Remote `
    -TargetMachine $t `
    -Credential $cred `
    -ConfigurationName UnityInstall `
    -SourceMsiPath $unityMsi `
    -UninstallExisting

===============================================================================


Step 5: Configure Service Location (Configure-ServiceLocation.ps1)
Points the installed Unity Client configuration / service locations to the appropriate Application Server URL / DataSource:




.\PowerShell\Configure-ServiceLocation.ps1 `
    -DeploymentMode Remote `
    -TargetMachine $t `
    -Credential $cred `
    -ConfigurationName UnityInstall `
    -DataSource 'OnBase'

=================================================================================================

Step 6: Verify Installation (Verify-Installation.ps1)
Validates that the SQL services are running, the 32-bit DSN connects successfully with hsi, the Unity Client executable/MSI is registered, and service location configs are populated:




.\PowerShell\Verify-Installation.ps1 `
    -DeploymentMode Remote `
    -TargetMachine $t `
    -Credential $cred `
    -ConfigurationName UnityInstall `
    -SqlInstance 'RDV-009989\SQL2K22' `
    -DataSource 'OnBase' `
    -OdbcDriver 'ODBC Driver 18 for SQL Server'

    OR

    .\PowerShell\Verify-Installation.ps1 `
    -DeploymentMode Remote `
    -TargetMachine $t `
    -Credential $cred `
    -ConfigurationName UnityInstall

=================================================================================================

=================================================================================================
##Miscellaneous Commands:

1. To immediately verify that Unity Client and SQL Server are actually present on RDV-009989 right now:

Invoke-Command -ComputerName $t -Credential $cred -ConfigurationName UnityInstall -ScriptBlock {
    [PSCustomObject]@{
        ComputerName = $env:COMPUTERNAME
        UnityInstalled = (Get-ItemProperty 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*', 
                                           'HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*' -ErrorAction SilentlyContinue |
                          Where-Object { $_.DisplayName -like '*Unity*' }).DisplayName
        UnityExeExists = Test-Path 'C:\Program Files (x86)\Hyland\Unity Client\obunity.exe'
        SQL2K22Service = (Get-Service -Name 'MSSQL$SQL2K22' -ErrorAction SilentlyContinue).Status
        SystemDSNs     = (Get-OdbcDsn -DsnType System -ErrorAction SilentlyContinue).Name
    }
}



2. Run this to confirm the exact parameter names supported by your .ps1 script:

(Get-Command .\PowerShell\Verify-Installation.ps1).Parameters.Keys   // Update ps1 file name as per your script location

=======================================================================================================

#### Running this single orchestrator command:
will automatically orchestrate the DB utilities, database creation, DSN creation, Unity MSI installation, 
service locations, and verification seamlessly without needing to run each step manually.



.\PowerShell\Install-OnBase.ps1 `
    -DeploymentMode Remote `
    -TargetMachine $t `
    -Credential $cred `
    -ConfigurationName UnityInstall `
    -BuildPath $build `
    -SqlInstance 'RDV-009989\SQL2K22' `
    -OdbcVersion 18 `
    -HsiCredential $hsi `
    -UninstallExisting


    QnA:

1. Will this run Steps 1 through 6?

With your exact command, it will execute Steps 2/3 through 6 (and skip Step 1):

Step 1 (SQL Server Installation - Install-SqlServer.ps1): SKIPPED / VERIFIED ONLY

        Because -SqlSetupPath (the path to SQL Server media setup.exe) and -SqlSaCredential are not passed in this command, 
        the orchestrator checks if RDV-009989\SQL2K22 already exists and is running.

        Since SQL2K22 is already installed on RDV-009989, it detects the instance as healthy and skips running the SQL installer. 
        (If you ever need a clean VM with Step 1 included, you would add -SqlSetupPath and -SqlSaCredential).

Step 2 (ODBC Driver Setup): CHECKED / EXECUTED

        It evaluates -OdbcVersion 18 and verifies/ensures ODBC Driver 18 for SQL Server is available on the remote machine.

Step 3 (Database & DSN - Initialize-OnBaseDatabase.ps1): EXECUTED

        Creates the OnBase database, sets up the hsi user, and creates the 32-bit System DSN.

Step 4 (Unity Client - Install-Unity.ps1): EXECUTED

        Uninstalls prior versions (due to -UninstallExisting) and installs the 25.1.28.1000 Unity Client MSI.

Step 5 (Service Location - Configure-ServiceLocation.ps1): EXECUTED

        Configures the client data source and service endpoints.

Step 6 (Verification - Verify-Installation.ps1): EXECUTED

        Runs the end-to-end sanity check on the target VM.

2. How the Orchestrator maps your command-line parameters
Install-OnBase.ps1 acts as a central coordinator. You only need to provide high-level parameters, 
and the orchestrator automatically derives and distributes the specific child parameters down to each individual script:

Parameter Passed to Orchestrator             How the Orchestrator Resolves & Passes It Down

-BuildPath $build
(e.g., \\onbase.net\...\25,1,28,1000)          Automatically splits into sub-paths:
                                                • For Step 3: Derives $build\utilities\FoundationDBUtils 
                                                  and passes it as -DbUtilsSourcePath to Initialize-OnBaseDatabase.ps1.
                                                • For Step 4: Derives $build\Install\Unity Client\Hyland Unity Client.msi 
                                                  and passes it as -SourceMsiPath to Install-Unity.ps1.

-OdbcVersion 18                                  Translates 18 into 'ODBC Driver 18 for SQL Server', configures the driver, 
                                                 and passes -OdbcDriver to Initialize-OnBaseDatabase.ps1 and Verify-Installation.ps1.

-SqlInstance 'RDV-009989\SQL2K22'                Passed directly as -SqlInstance to Initialize-OnBaseDatabase.ps1 and Verify-Installation.ps1.

-HsiCredential $hsi                              Passed directly to Initialize-OnBaseDatabase.ps1 to configure the SQL login and test the ODBC DSN.

-UninstallExisting                               Forwarded to Install-Unity.ps1 to trigger unattended removal of prior Unity versions before installing the new build.

-DeploymentMode Remote, 
-TargetMachine $t, -Credential $cred, 
-ConfigurationName UnityInstall         Wrapped into remoting arguments (Invoke-Command) for all child scripts running against RDV-009989.

================================================================================
================================================================================

What happens on a brand-new VM:

If you run the previous command on a brand-new, clean VM without any pre-installed applications, here is what will happen and why:

1. What will happen with the previous command
Remote Connection Failure (Initial Blocker):

The new VM will not have the custom UnityInstall PowerShell remoting endpoint registered, and WinRM may not be enabled or added to your workstation's TrustedHosts.

Result: The command will fail at the very start with a PSSessionConfiguration or WinRM connection error.

Step 1 (SQL Server) is not triggered:

The previous command omitted the SQL installer media parameters (-SqlSetupPath / -SqlAdminCredential). The orchestrator will assume the instance NEW-VM\SQL2K22 already exists.

Step 3 (Database Initialization) will fail:

When Initialize-OnBaseDatabase.ps1 tries to connect to NEW-VM\SQL2K22, it will fail because no SQL Server service exists.

Step 2 (ODBC Driver) may fail:

If ODBC Driver 18 for SQL Server is not pre-installed on the clean OS, creating the System DSN will fail unless the framework installs it automatically or the driver installer is available.

2. What you need for a brand-new VM
To deploy everything from scratch (Step 1 through Step 6) on a blank VM, follow this two-phase workflow:

Phase 1: One-time VM Remoting Setup (Run once on the new VM)
Before running remote commands from your workstation, open PowerShell as Administrator on the new VM (or run via bootstrap/sysprep script):




# 1. Enable WinRM & Firewall
Enable-PSRemoting -Force
Set-Service -Name WinRM -StartupType Automatic
Start-Service -Name WinRM
# 2. Register the 'UnityInstall' Remoting Endpoint
Register-PSSessionConfiguration -Name 'UnityInstall' -Force
# 3. Enable Inbound File Sharing / ICMP
netsh advfirewall firewall set rule group="File and Printer Sharing" new enable=Yes
On your local workstation, add the new VM to your TrustedHosts:




$newVm = 'RDV-NEWVM.hylandqa.net'
Set-Item WSMan:\localhost\Client\TrustedHosts -Value $newVm -Concatenate -Force
Phase 2: Full End-to-End Orchestrator Command (Steps 1–6)
To execute the entire provisioning pipeline from clean OS to verified installation, pass the SQL Setup media path and SQL Admin credentials to Install-OnBase.ps1:




$t = 'RDV-NEWVM.hylandqa.net'
$cred = Get-Credential -UserName 'ONBASE\prroy'
$hsi = Get-Credential -UserName 'hsi' -Message 'Password for the hsi SQL user'
$sqlSa = Get-Credential -UserName 'sa' -Message 'Password for the SQL sa account'
$build = '\\' + 'onbase.net\shares\Builds\OnBase\Current\25_1\25,1,28,1000'
$sqlMedia = '\\' + 'onbase.net\shares\ThirdParty\Microsoft\SQLServer\2022\Enterprise'
.\PowerShell\Install-OnBase.ps1 `
    -DeploymentMode Remote `
    -TargetMachine $t `
    -Credential $cred `
    -ConfigurationName UnityInstall `
    -BuildPath $build `
    -SqlSetupPath $sqlMedia `
    -SqlInstance 'SQL2K22' `
    -SqlAdminCredential $sqlSa `
    -OdbcVersion 18 `
    -HsiCredential $hsi `
    -UninstallExisting
What the Orchestrator will do on a clean VM:
Step

Action Performed

Step 1

Detects no SQL instance → Invokes Install-SqlServer.ps1 from $sqlMedia, installs SQL Server 2022 instance SQL2K22, and configures sa authentication.

Step 2

Validates and installs ODBC Driver 18 for SQL Server if missing.

Step 3

Runs Initialize-OnBaseDatabase.ps1, creates the OnBase database, sets up hsi user, and creates the 32-bit System DSN.

Step 4

Runs Install-Unity.ps1 to install the OnBase Unity Client 25.1 MSI silently.

Step 5

Runs Configure-ServiceLocation.ps1 to register datasource and App Server URLs.

Step 6

Runs Verify-Installation.ps1 to confirm services, registry, files, and DSNs are operational.



===============================================================================================



2. Complete Orchestration Execution Commands


A. Full Command for a Clean / New VM (Steps 1–6 with SQL & ODBC Install)

# 1. Define Target & Credentials
$t    = 'RDV-009989.hylandqa.net'
$cred = Get-Credential -UserName 'ONBASE\prroy' -Message 'Domain admin credential for remote VM access'
$hsi  = Get-Credential -UserName 'hsi'          -Message 'Password for the OnBase hsi SQL user'
$sa   = Get-Credential -UserName 'sa'           -Message 'Strong password for SQL sa (e.g. Hyl@nd_2026!Sec#Db)'
# 2. Define Shared Network Repository Paths
$build     = '\\' + 'onbase.net\shares\Builds\OnBase\Current\25_1\25,1,28,1000'
$sqlMedia  = '\\' + 'onbase.net\shares\ThirdParty\Microsoft\SQLServer\2022\Enterprise'
$odbcMedia = '\\' + 'onbase.net\shares\ThirdParty\Microsoft\ODBC\v18'
# 3. Execute End-to-End Provisioning
.\PowerShell\Install-OnBaseOrchestration.ps1 `
    -DeploymentMode Remote `
    -TargetMachine $t `
    -Credential $cred `
    -ConfigurationName UnityInstall `
    -BuildPath $build `
    -SqlSetupPath $sqlMedia `
    -SqlVersion 2022 `
    -SqlInstance 'RDV-009989\SQL2K22' `
    -SaCredential $sa `
    -OnSqlVersionMismatch SideBySide `
    -OdbcVersion 18 `
    -OdbcSetupPath $odbcMedia `
    -HsiCredential $hsi `
    -DatabaseName 'OnBase' `
    -UninstallExisting

B. Quick Command for an Existing VM (Steps 2/3–6 where SQL Server already exists)



# 1. Define Target & Credentials
$t    = 'RDV-009989.hylandqa.net'
$cred = Get-Credential -UserName 'ONBASE\prroy'
$hsi  = Get-Credential -UserName 'hsi'
# 2. Define Shared Network Repository Path
$build = '\\' + 'onbase.net\shares\Builds\OnBase\Current\25_1\25,1,28,1000'
# 3. Execute Orchestration (skips SQL install, handles DB Init, Unity Client MSI, & Verification)
.\PowerShell\Install-OnBaseOrchestration.ps1 `
    -DeploymentMode Remote `
    -TargetMachine $t `
    -Credential $cred `
    -ConfigurationName UnityInstall `
    -BuildPath $build `
    -SqlInstance 'RDV-009989\SQL2K22' `
    -OdbcVersion 18 `
    -HsiCredential $hsi `
    -DatabaseName 'OnBase' `
    -UninstallExisting




===============================================================================================
https://hyland.atlassian.net/wiki/spaces/HFSS/pages/4331907517/OnBase+Provisioning+Framework+Implementation+Decisions+and+Operational+Guidance


---------------------------Not require as current implementation has installers available from Proget------------------------
Run Install-HylandIdSAndApi.ps1 (Get installer fron shared Path)
You can now run the installation and configuration of Hyland Identity Server and API Server:




# 1. Credentials (if $hsi is not already set in this session)
$t    = 'RDV-009989.hylandqa.net'
$cred = Get-Credential -UserName 'ONBASE\prroy'
$hsi  = Get-Credential -UserName 'hsi'
# 2. Source Repositories & Packages
$idpSrc = '\\' + 'onbase.net\shares\Builds\OnBase\Current\25_1\25,1,28,1000\Install\Identity Server'
$apiSrc = '\\' + 'onbase.net\shares\Builds\OnBase\Current\25_1\25,1,28,1000\Install\API Server'
$netSrc = '\\' + 'onbase.net\shares\PublicApps\Applications\DeveloperApps\Microsoft .NET Core\8.0\dotnet-hosting-8.0.x-win.exe'
# 3. Execute
.\PowerShell\Install-HylandIdSAndApi.ps1 `
    -DeploymentMode Remote `
    -TargetMachine $t `
    -Credential $cred `
    -ConfigurationName UnityInstall `
    -IdSSourcePath $idpSrc `
    -ApiServerSourcePath $apiSrc `
    -DotNetHostingBundlePath $netSrc `
    -SqlInstance 'RDV-009989\SQL2K22' `
    -DatabaseName 'OnBase' `
    -HsiCredential $hsi

    -----------------------------------------------------------------------------------------------------------------------------

Run Install-HylandIdSAndApi.ps1 (Get installer from Proget)
You can now run the installation and configuration of Hyland Identity Server and API Server:

$t       = 'RDV-009989.hylandqa.net'
$cred    = Get-Credential -UserName 'ONBASE\prroy'
$hsi     = Get-Credential -UserName 'hsi'
$proget  = 'https://proget.onbase.net/nuget/Release/v3/index.json'

.\PowerShell\Install-HylandIdSAndApi.ps1 `
    -DeploymentMode Remote `
    -TargetMachine $t `
    -Credential $cred `
    -ConfigurationName UnityInstall `
    -ProGetFeedUrl $proget `
    -IdSVersion '5.1.0' `
    -ApiServerVersion '25.1.28.1000' `
    -SqlInstance 'RDV-009989\SQL2K22' `
    -DatabaseName 'OnBase' `
    -HsiCredential $hsi

===========
    # 1. Target Machine and Credentials
$t       = 'RDV-009989.hylandqa.net'
$cred    = Get-Credential -UserName 'ONBASE\prroy'
$hsi     = Get-Credential -UserName 'hsi'
$proget  = 'https://proget.onbase.net/nuget/Release/v3/index.json'

# 2. Execute Version-Aware, Self-Healing Deployment
.\PowerShell\Install-HylandIdSAndApi.ps1 `
    -DeploymentMode Remote `
    -TargetMachine $t `
    -Credential $cred `
    -ConfigurationName UnityInstall `
    -ProGetFeedUrl $proget `
    -IdSVersion '5.1.0' `
    -ApiServerVersion '25.1.28.1000' `
    -SqlInstance 'RDV-009989\SQL2K22' `
    -DatabaseName 'OnBase' `
    -HsiCredential $hsi

    =========

    # 1. Target Machine and Credentials Setup
$t       = 'RDV-009989.hylandqa.net'
$cred    = Get-Credential -UserName 'ONBASE\prroy'
$hsi     = Get-Credential -UserName 'hsi'
$proget  = 'https://proget.onbase.net/nuget/Release'

# 2. Execute Version-Aware Deployment
.\PowerShell\Install-HylandIdSAndApi.ps1 `
    -DeploymentMode Remote `
    -TargetMachine $t `
    -Credential $cred `
    -ConfigurationName UnityInstall `
    -ProGetFeedUrl $proget `
    -IdSVersion '5.1.0' `
    -ApiServerVersion '25.1.28.1000' `
    -SqlInstance 'RDV-009989\SQL2K22' `
    -DatabaseName 'OnBase' `
    -HsiCredential $hsi

    =================
    Updated Command 

    $cred = Get-Credential -UserName 'ONBASE\prroy'
    .\PowerShell\Install-HylandIdSAndApi.ps1 `
        -SettingsFilePath '.\Settings.json' `
        -Credential $cred

        Update-Cmd

            $cred = Get-Credential -UserName 'ONBASE\prroy'
    .\PowerShell\Install-HylandIdSAndApi.ps1
        -Credential $cred

        

## How to Run Install-OnBaseOrchestration.ps1 
Option A: Run Full Pipeline using Settings.json

$cred = Get-Credential -UserName 'ONBASE\prroy'
.\PowerShell\Install-OnBaseOrchestration.ps1 `
    -SettingsFilePath '.\Settings.json' `
    -Credential $cred

Option B: Run ONLY the Identity Service & API Server Step

$cred = Get-Credential -UserName 'ONBASE\prroy'
.\PowerShell\Install-OnBaseOrchestration.ps1 `
    -SettingsFilePath '.\Settings.json' `
    -Credential $cred `
    -SkipSqlInstall `
    -SkipOdbc `
    -SkipDatabase `
    -SkipUnityClient `
    -SkipServiceLocation



    =======================================

Set-Location 'C:\Users\PRROY\source\OnBaseProvisioningFramework'

$cred = Get-Credential `
    -UserName 'ONBASE\prroy' `
    -Message 'Remote provisioning credentials'

$hsi = Get-Credential `
    -UserName 'hsi' `
    -Message 'OnBase hsi SQL password'

$sa = Get-Credential `
    -UserName 'sa' `
    -Message 'Enter the NEW sa password for the SQL instance being installed'

.\PowerShell\Install-OnBaseOrchestration.ps1 `
    -SettingsFilePath '.\Settings.json' `
    -Credential $cred `
    -HsiCredential $hsi `
    -SaCredential $sa


    =========

    Set-Location 'C:\Users\PRROY\source\OnBaseProvisioningFramework'

$cred = Get-Credential `
    -UserName 'ONBASE\prroy' `
    -Message 'Remote provisioning credentials'

$hsi = Get-Credential `
    -UserName 'hsi' `
    -Message 'OnBase hsi SQL password'

$sa = Get-Credential `
    -UserName 'sa' `
    -Message 'Enter the NEW sa password for the SQL instance being installed'

.\PowerShell\Install-OnBaseOrchestration.ps1 `
    -SettingsFilePath '.\Settings.json' `
    -Credential $cred `
    -HsiCredential $hsi `
    -SaCredential $sa

$LASTEXITCODE

=========================

Use your existing $cred, $hsi, and $sa credential variables. Set a switch to $true to skip that stage; leave it $false to run it.




Set-Location 'C:\Users\PRROY\source\OnBaseProvisioningFramework'
$cred = Get-Credential `
    -UserName 'ONBASE\prroy' `
    -Message 'Remote provisioning credentials'

$hsi = Get-Credential `
    -UserName 'hsi' `
    -Message 'OnBase hsi SQL password'

$sa = Get-Credential `
    -UserName 'sa' `
    -Message 'Enter the NEW sa password for the SQL instance being installed'
$launchParams = @{
    SettingsFilePath = '.\Settings.json'
    Credential       = $cred
    HsiCredential    = $hsi
    SaCredential     = $sa
    # Change to $true only for stages you want to SKIP.
    SkipSqlInstall        = $false
    SkipOdbc              = $true
    SkipDatabase          = $true
    SkipApplicationServer = $true
    SkipUnityClient       = $true
    SkipServiceLocation   = $true
    SkipIdSAndApi          = $true
    SkipVerification      = $true
}
.\PowerShell\Install-OnBaseOrchestration.ps1 @launchParams
$LASTEXITCODE

===========

$cred = Get-Credential -UserName 'ONBASE\prroy' -Message 'Remote provisioning credentials: EXISTING Windows password'

$sa = Get-Credential -UserName 'sa' -Message 'For a NEW SQL installation: enter the intended sa password'

$launchParams = @{
    SettingsFilePath      = '.\Settings.json'
    Credential            = $cred
    SaCredential          = $sa

    # UnityInstall is already registered and working.
    BootstrapUnityInstall = $false

    # Run only the SQL stage.
    SkipSqlInstall        = $false
    SkipOdbc              = $true
    SkipDatabase          = $true
    SkipApplicationServer = $true
    SkipUnityClient       = $true
    SkipServiceLocation   = $true
    SkipIdSAndApi          = $true
    SkipVerification      = $true
}

.\PowerShell\Install-OnBaseOrchestration.ps1 @launchParams

$LASTEXITCODE

==========My Latest Command

Set-Location 'C:\Users\PRROY\source\OnBaseProvisioningFramework'

$cred = Get-Credential `
    -UserName 'ONBASE\prroy' `
    -Message 'Remote provisioning credentials: EXISTING Windows password'

$hsi = Get-Credential `
    -UserName 'hsi' `
    -Message 'OnBase hsi SQL password: existing password if the login exists'

$launchParams = @{
    SettingsFilePath                     = '.\Settings.json'
    Credential                           = $cred
    HsiCredential                        = $hsi

    # Reuse the working endpoint without registration or WinRM restart.
    BootstrapUnityInstall                = $false

    # Overrides the earlier settings value Stop with a supported parent value.
    OnSqlVersionMismatch                 = 'Fail'

    'SkipInstall-SqlServer'               = $true
    'SkipCreate-Odbc'                     = $true
    'SkipInitializeOnBaseDatabase'        = $true
    'SkipInstall-HylandIdSAndApi'          = $true
    'SkipInstall-OnBaseApplicationServer' = $false
    'SkipInstall-UnityClient'             = $true
    'SkipVerify-Installation'             = $true
}

.\PowerShell\Install-OnBaseOrchestration.ps1 @launchParams

$LASTEXITCODE



Resolve-DnsName $t

Test-NetConnection -RDV-011241.hylandqa.net $t -Port 5985

Test-WSMan -RDV-011241.hylandqa.net $t -ErrorAction Stop