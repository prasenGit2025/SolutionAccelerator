# OnBase Provisioning Framework

## Overview

This guide provides comprehensive instructions for provisioning OnBase 25.1.28.1000 on a Windows Server/Windows 11 target machine using the OnBase Provisioning Framework's PowerShell orchestration tools.

---

## Table of Contents

1. [Prerequisites & Session Setup](#prerequisites--session-setup)
2. [Installation Steps Overview](#installation-steps-overview)
   - [Step 1: SQL Server Installation](#step-1-sql-server-installation)
   - [Step 2: Database Initialization](#step-2-database-initialization)
   - [Step 3: Unity Client Installation](#step-3-unity-client-installation)
   - [Step 4: Service Location Configuration](#step-4-service-location-configuration)
   - [Step 5: Installation Verification](#step-5-installation-verification)
3. [Orchestrator Commands](#orchestrator-commands)
4. [Troubleshooting](#troubleshooting)
5. [Miscellaneous Commands](#miscellaneous-commands)
6. [FAQ](#faq)

---

## Prerequisites & Session Setup

### Required Credentials

Before starting, gather the following credentials:

```powershell
cd C:\Users\PRROY\source\OnBaseProvisioningFramework

# Domain credentials for remote machine access
$t     = 'RDV-009989.hylandqa.net'
$bs    = [string][char]92  # Backslash for UNC paths
$cred  = Get-Credential -UserName 'ONBASE\prroy' -Message 'ONBASE credentials'

# OnBase hsi SQL user password
$hsi   = Get-Credential -UserName 'hsi' -Message 'OnBase hsi password'

# SQL Server media (if installing SQL Server)
$sqlSrc = $bs + $bs + 'onbase.net\shares\Public\SQLServer\sql_server_2022_developer.iso'

# OnBase build path
$build = $bs + $bs + 'onbase.net\shares\Builds\OnBase\Current\25_1\25,1,28,1000'
```

### Session Variables

| Variable | Purpose | Example |
|----------|---------|---------|
| `$t` | Target machine hostname | `'RDV-009989.hylandqa.net'` |
| `$cred` | Domain credentials (ONBASE\username) | `Get-Credential` |
| `$hsi` | HSI SQL user password | `Get-Credential` |
| `$build` | UNC path to OnBase build | `'\\onbase.net\shares\Builds\OnBase\Current\25_1\25,1,28,1000'` |
| `$sqlSrc` | SQL Server 2022 media path | `'\\onbase.net\shares\Public\SQLServer\...'` |

---

## Installation Steps Overview

### Step 1: SQL Server Installation

**Prerequisites:**
- A clean Windows Server or Windows 11 target machine
- SQL Server 2022 media available at `$sqlSrc`
- Remote target machine is joined to domain with WinRM enabled

**Command:**

```powershell
.\PowerShell\Install-OnBase.ps1 `
    -DeploymentMode Remote `
    -TargetMachine $t `
    -Credential $cred `
    -ConfigurationName UnityInstall `
    -BuildPath $build `
    -SqlVersion 2022 `
    -SqlSetupPath $sqlSrc `
    -OnSqlVersionMismatch SideBySide `
    -OdbcVersion 18 `
    -HsiCredential $hsi `
    -UninstallExisting
```

**What This Does:**
- Installs SQL Server 2022 Developer Edition on the target machine
- Creates a named instance `SQL2K22` (default)
- Installs ODBC Driver 18 for SQL Server
- Configures side-by-side installation mode if SQL Server already exists

**Troubleshooting SQL Server Installation Failures:**

If the installation fails, retrieve detailed logs:

```powershell
Invoke-Command -ComputerName $t -Credential $cred -ConfigurationName UnityInstall -ScriptBlock {
    $log = 'C:\Program Files\Microsoft SQL Server\160\Setup Bootstrap\Log'
    Get-Content (Join-Path $log 'Summary.txt') -TotalCount 40
    $last = Get-ChildItem $log -Directory | Sort-Object LastWriteTime | Select-Object -Last 1
    "--- Detail.txt errors ($($last.Name)) ---"
    Select-String -Path (Join-Path $last.FullName 'Detail.txt') -Pattern 'Exception|error|failed' | Select-Object -Last 15 | ForEach-Object Line
}
```

### Post-Installation: Restart & Verification

**Restart the target machine:**

```powershell
Invoke-Command -ComputerName $t -Credential $cred -ConfigurationName UnityInstall -ScriptBlock {
    Restart-Computer -Force
}
```

**Wait for the machine to come back online:**

```powershell
do { 
    Write-Host "Waiting for $t to respond..." -ForegroundColor Yellow
    Start-Sleep -Seconds 10 
} until (Test-WSMan -ComputerName $t -ErrorAction SilentlyContinue)
Write-Host "$t is back online!" -ForegroundColor Green
```

**Verify WinRM & PowerShell remoting connectivity:**

```powershell
Test-WSMan -ComputerName $t
```

Expected output:
```
wsmid           : http://schemas.dmtf.org/wbem/wsman/identity/1/wsmanidentity.xsd
ProtocolVersion : http://schemas.dmtf.org/wbem/wsman/1/wsman.xsd
ProductVendor   : Microsoft Corporation
ProductVersion  : OS: 0.0.0 SP: 0.0 Stack: 3.0
```

**Verify session configuration:**

```powershell
Invoke-Command -ComputerName $t -Credential $cred -ConfigurationName UnityInstall -ScriptBlock {
    [PSCustomObject]@{
        ComputerName = $env:COMPUTERNAME
        CurrentUser  = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
        Uptime       = (Get-CimInstance Win32_OperatingSystem).LastBootUpTime
    }
}
```

Expected output:
```
ComputerName   : RDV-009989
CurrentUser    : ONBASE\prroy
Uptime         : 10/1/2026 11:11:43 PM
PSComputerName : RDV-009989.hylandqa.net
RunspaceId     : c6d0d379-0c75-4b8c-96ee-18b7c5f399f3
```

**Verify SQL Server services are running:**

```powershell
Invoke-Command -ComputerName $t -Credential $cred -ConfigurationName UnityInstall -ScriptBlock {
    Get-Service -Name '*SQL*' | Select-Object Name, DisplayName, Status, StartType
}
```

Expected output:
```
Name           : MSSQL$SQL2K22
DisplayName    : SQL Server (SQL2K22)
Status         : Running
StartType      : Automatic
PSComputerName : RDV-009989.hylandqa.net
```

---

### Step 2: Database Initialization

**Purpose:**
- Creates the OnBase database on the SQL Server instance
- Configures the `hsi` SQL user with provided credentials
- Creates a 32-bit System DSN using ODBC Driver 18

**Setup UNC Path for Database Utilities:**

```powershell
# Build the UNC network path
$build = '\\' + 'onbase.net\shares\Builds\OnBase\Current\25_1\25,1,28,1000'
$dbUtils = $build + '\utilities\FoundationDBUtils'

# Verify the path is accessible
$dbUtils
# Expected: \\onbase.net\shares\Builds\OnBase\Current\25_1\25,1,28,1000\utilities\FoundationDBUtils

# Test network share access
Test-Path $dbUtils
# Expected: True
```

**Command:**

```powershell
.\PowerShell\Initialize-OnBaseDatabase.ps1 `
    -DeploymentMode Remote `
    -TargetMachine $t `
    -Credential $cred `
    -ConfigurationName UnityInstall `
    -DbUtilsSourcePath $dbUtils `
    -SqlInstance 'RDV-009989\SQL2K22' `
    -OdbcDriver 'ODBC Driver 18 for SQL Server' `
    -HsiCredential $hsi `
    -DatabaseName 'OnBase'
```

**Parameters:**

| Parameter | Description |
|-----------|-------------|
| `-DeploymentMode Remote` | Deploy to remote machine via WinRM |
| `-TargetMachine $t` | Target machine hostname or FQDN |
| `-Credential $cred` | Domain credentials |
| `-ConfigurationName UnityInstall` | Custom PowerShell endpoint name |
| `-DbUtilsSourcePath $dbUtils` | UNC path to FoundationDBUtils |
| `-SqlInstance` | SQL Server instance (hostname\instance) |
| `-OdbcDriver` | ODBC driver name (e.g., 'ODBC Driver 18 for SQL Server') |
| `-HsiCredential $hsi` | SQL user credentials for 'hsi' login |
| `-DatabaseName` | Name of the OnBase database |

---

### Step 3: Unity Client Installation

**Purpose:**
- Copies the Unity Client MSI from the build share to the target VM
- Removes previous versions (if `-UninstallExisting` is specified)
- Performs unattended installation

**Prepare MSI Path:**

```powershell
$build = '\\' + 'onbase.net\shares\Builds\OnBase\Current\25_1\25,1,28,1000'
$unityMsi = Join-Path $build 'Install\Unity Client\Hyland Unity Client.msi'

# Verify the MSI is accessible
Test-Path $unityMsi
# Expected: True
```

**Command:**

```powershell
.\PowerShell\Install-Unity.ps1 `
    -DeploymentMode Remote `
    -TargetMachine $t `
    -Credential $cred `
    -ConfigurationName UnityInstall `
    -SourceMsiPath $unityMsi `
    -UninstallExisting
```

**Parameters:**

| Parameter | Description |
|-----------|-------------|
| `-DeploymentMode Remote` | Deploy to remote machine |
| `-TargetMachine $t` | Target machine hostname |
| `-Credential $cred` | Domain credentials |
| `-ConfigurationName UnityInstall` | PowerShell endpoint configuration |
| `-SourceMsiPath $unityMsi` | UNC path to the Unity Client MSI |
| `-UninstallExisting` | Remove prior versions before installing |

---

### Step 4: Service Location Configuration

**Purpose:**
- Points the installed Unity Client configuration to the appropriate Application Server URL and DataSource

**Command:**

```powershell
.\PowerShell\Configure-ServiceLocation.ps1 `
    -DeploymentMode Remote `
    -TargetMachine $t `
    -Credential $cred `
    -ConfigurationName UnityInstall `
    -DataSource 'OnBase'
```

**Parameters:**

| Parameter | Description |
|-----------|-------------|
| `-DeploymentMode Remote` | Deploy to remote machine |
| `-TargetMachine $t` | Target machine hostname |
| `-Credential $cred` | Domain credentials |
| `-ConfigurationName UnityInstall` | PowerShell endpoint configuration |
| `-DataSource` | OnBase datasource name |

---

### Step 5: Installation Verification

**Purpose:**
- Validates SQL services are running
- Tests 32-bit DSN connectivity with hsi credentials
- Verifies Unity Client executable and MSI registration
- Confirms service location configurations are populated

**Command Option 1 - Full Verification:**

```powershell
.\PowerShell\Verify-Installation.ps1 `
    -DeploymentMode Remote `
    -TargetMachine $t `
    -Credential $cred `
    -ConfigurationName UnityInstall `
    -SqlInstance 'RDV-009989\SQL2K22' `
    -DataSource 'OnBase' `
    -OdbcDriver 'ODBC Driver 18 for SQL Server'
```

**Command Option 2 - Quick Verification:**

```powershell
.\PowerShell\Verify-Installation.ps1 `
    -DeploymentMode Remote `
    -TargetMachine $t `
    -Credential $cred `
    -ConfigurationName UnityInstall
```

**Parameters:**

| Parameter | Description |
|-----------|-------------|
| `-DeploymentMode Remote` | Deploy to remote machine |
| `-TargetMachine $t` | Target machine hostname |
| `-Credential $cred` | Domain credentials |
| `-ConfigurationName UnityInstall` | PowerShell endpoint configuration |
| `-SqlInstance` | SQL Server instance (optional) |
| `-DataSource` | OnBase datasource name (optional) |
| `-OdbcDriver` | ODBC driver name (optional) |

---

## Orchestrator Commands

The `Install-OnBase.ps1` orchestrator script automates Steps 2–6 (or Steps 1–6 on a clean VM) by automatically deriving and distributing parameters to individual scripts.

### Parameter Mapping

| Orchestrator Parameter | Derived Paths | Passed To |
|------------------------|----------------|-----------|
| `-BuildPath $build` | Derives `$build\utilities\FoundationDBUtils` | `Initialize-OnBaseDatabase.ps1` (-DbUtilsSourcePath) |
| | Derives `$build\Install\Unity Client\Hyland Unity Client.msi` | `Install-Unity.ps1` (-SourceMsiPath) |
| `-OdbcVersion 18` | Translates to `ODBC Driver 18 for SQL Server` | `Initialize-OnBaseDatabase.ps1`, `Verify-Installation.ps1` |
| `-SqlInstance` | Passed directly | `Initialize-OnBaseDatabase.ps1`, `Verify-Installation.ps1` |
| `-HsiCredential $hsi` | Passed directly | `Initialize-OnBaseDatabase.ps1` |
| `-UninstallExisting` | Forwarded | `Install-Unity.ps1` |
| Remote parameters | Wrapped into `Invoke-Command` arguments | All child scripts |

---

### Command A: Full Deployment for a Clean VM (Steps 1–6)

Use this command for a brand-new VM with no pre-installed applications.

**Prerequisites:**
- Target VM has WinRM enabled and is joined to the domain
- Custom `UnityInstall` PowerShell endpoint is registered (see [New VM Setup](#new-vm-setup) below)
- SQL Server media and ODBC media are accessible

**Command:**

```powershell
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
.\PowerShell\Install-OnBase.ps1 `
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
```

**Steps Executed:**
1. ✓ SQL Server 2022 installation
2. ✓ ODBC Driver validation/installation
3. ✓ Database initialization and DSN creation
4. ✓ Unity Client installation
5. ✓ Service location configuration
6. ✓ Installation verification

---

### Command B: Quick Deployment for an Existing VM (Steps 2/3–6)

Use this command for a VM that already has SQL Server installed.

**Prerequisites:**
- Target VM has WinRM enabled and is joined to the domain
- Custom `UnityInstall` PowerShell endpoint is registered
- SQL Server is already installed and running

**Command:**

```powershell
# 1. Define Target & Credentials
$t    = 'RDV-009989.hylandqa.net'
$cred = Get-Credential -UserName 'ONBASE\prroy'
$hsi  = Get-Credential -UserName 'hsi'

# 2. Define Shared Network Repository Path
$build = '\\' + 'onbase.net\shares\Builds\OnBase\Current\25_1\25,1,28,1000'

# 3. Execute Orchestration (skips SQL install, handles DB Init, Unity Client MSI, & Verification)
.\PowerShell\Install-OnBase.ps1 `
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
```

**Steps Executed:**
1. ✗ SQL Server installation (skipped)
2. ✓ ODBC Driver validation
3. ✓ Database initialization and DSN creation
4. ✓ Unity Client installation
5. ✓ Service location configuration
6. ✓ Installation verification

---

### Orchestrator QnA

#### Will this run Steps 1 through 6?

With **Command B**, it will execute Steps 2/3 through 6:

- **Step 1 (SQL Server Installation):** SKIPPED / VERIFIED ONLY
  - Because `-SqlSetupPath` and `-SqlSaCredential` are not passed, the orchestrator checks if the instance already exists.
  - Since SQL2K22 is running, installation is skipped.

- **Step 2 (ODBC Driver Setup):** CHECKED / EXECUTED
  - Evaluates `-OdbcVersion 18` and ensures ODBC Driver 18 is available.

- **Step 3 (Database & DSN):** EXECUTED
  - Creates OnBase database, sets up hsi user, and creates System DSN.

- **Step 4 (Unity Client):** EXECUTED
  - Uninstalls prior versions and installs the Unity Client MSI.

- **Step 5 (Service Location):** EXECUTED
  - Configures datasource and App Server URLs.

- **Step 6 (Verification):** EXECUTED
  - Runs sanity checks on services, registry, files, and DSNs.

---

## New VM Setup

For a **brand-new, clean VM**, you must complete a one-time remoting setup before deploying OnBase.

### Phase 1: One-Time VM Remoting Setup

**On the new VM (run as Administrator):**

```powershell
# 1. Enable WinRM & Firewall
Enable-PSRemoting -Force
Set-Service -Name WinRM -StartupType Automatic
Start-Service -Name WinRM

# 2. Register the 'UnityInstall' Remoting Endpoint
Register-PSSessionConfiguration -Name 'UnityInstall' -Force

# 3. Enable Inbound File Sharing / ICMP
netsh advfirewall firewall set rule group="File and Printer Sharing" new enable=Yes
```

**On your local workstation:**

```powershell
$newVm = 'RDV-NEWVM.hylandqa.net'
Set-Item WSMan:\localhost\Client\TrustedHosts -Value $newVm -Concatenate -Force
```

### Phase 2: Full End-to-End Orchestrator Command

After Phase 1, run **Command A** (Full Deployment) to provision everything from scratch.

---

## Troubleshooting

### Remote Connection Failures

**Error:** PSSessionConfiguration or WinRM connection error

**Solution:**
1. Verify WinRM is enabled on the target machine:
   ```powershell
   Test-WSMan -ComputerName $t
   ```
2. Ensure the machine is in TrustedHosts:
   ```powershell
   Set-Item WSMan:\localhost\Client\TrustedHosts -Value $t -Concatenate -Force
   ```
3. Verify the `UnityInstall` endpoint is registered:
   ```powershell
   Get-PSSessionConfiguration -Name UnityInstall -ErrorAction SilentlyContinue
   ```

### Database Initialization Failures

**Error:** Cannot connect to SQL Server instance

**Solution:**
1. Verify SQL Server is installed and running:
   ```powershell
   Invoke-Command -ComputerName $t -Credential $cred -ConfigurationName UnityInstall -ScriptBlock {
       Get-Service -Name '*SQL*' | Where-Object { $_.Name -like 'MSSQL*' }
   }
   ```
2. Verify the instance name is correct in your command
3. Check SQL Server logs for installation errors (see [Step 1 troubleshooting](#troubleshooting-sql-server-installation-failures))

### DSN Connection Issues

**Error:** Cannot connect to System DSN with hsi credentials

**Solution:**
1. Verify ODBC Driver 18 is installed:
   ```powershell
   Invoke-Command -ComputerName $t -Credential $cred -ConfigurationName UnityInstall -ScriptBlock {
       Get-OdbcDriver -Name '*SQL Server*'
   }
   ```
2. Test the DSN manually:
   ```powershell
   Invoke-Command -ComputerName $t -Credential $cred -ConfigurationName UnityInstall -ScriptBlock {
       Get-OdbcDsn -DsnType System -Name 'OnBase'
   }
   ```

### Unity Client Installation Issues

**Error:** MSI installation fails or previous version cannot be uninstalled

**Solution:**
1. Check installed software on the target machine:
   ```powershell
   Invoke-Command -ComputerName $t -Credential $cred -ConfigurationName UnityInstall -ScriptBlock {
       Get-ItemProperty 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*', 
                        'HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*' -ErrorAction SilentlyContinue |
       Where-Object { $_.DisplayName -like '*Unity*' } | Select-Object DisplayName, Version
   }
   ```
2. Manually uninstall the old version if automatic removal fails
3. Re-run the install step

---

## Miscellaneous Commands

### 1. Verify Unity Client and SQL Server Presence

To immediately check that Unity Client and SQL Server are installed on the target machine:

```powershell
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
```

Expected output:
```
ComputerName      : RDV-009989
UnityInstalled    : Hyland Unity Client
UnityExeExists    : True
SQL2K22Service    : Running
SystemDSNs        : {OnBase}
```

### 2. Get Parameter Names for a Script

To confirm the exact parameter names supported by a script:

```powershell
(Get-Command .\PowerShell\Verify-Installation.ps1).Parameters.Keys
```

Replace `Verify-Installation.ps1` with any other script name as needed.

---

## FAQ

### Q: Will the orchestrator run Steps 1 through 6?

**A:** It depends on your command:

- **With Command A** (Full Deployment): Yes, Steps 1–6 will run, provided SQL media paths and credentials are passed.
- **With Command B** (Quick Deployment): Steps 2/3–6 will run; Step 1 is skipped (SQL Server already exists).

### Q: What if SQL Server is already installed but in a different instance name?

**A:** Modify the `-SqlInstance` parameter to match your existing instance:

```powershell
-SqlInstance 'RDV-009989\MyCustomInstance'
```

The orchestrator will detect the existing instance and skip SQL installation.

### Q: Can I run individual steps manually instead of using the orchestrator?

**A:** Yes. Each script can be run independently:
- `Install-OnBase.ps1` (full orchestration)
- `Initialize-OnBaseDatabase.ps1` (Step 2)
- `Install-Unity.ps1` (Step 3)
- `Configure-ServiceLocation.ps1` (Step 4)
- `Verify-Installation.ps1` (Step 5)

However, the orchestrator is recommended for production deployments as it handles error checking and parameter derivation automatically.

### Q: What credentials do I need?

**A:** Three sets of credentials:
1. **Domain account** (ONBASE\username) — for remote machine access
2. **HSI user password** — for the OnBase SQL user
3. **SQL SA password** (if installing SQL) — for SQL Server system administrator account

### Q: How long does a full provisioning take?

**A:** Approximately 45–90 minutes depending on network speed and hardware:
- SQL Server installation: 20–30 minutes
- Database initialization: 5–10 minutes
- Unity Client installation: 5–10 minutes
- Configuration & verification: 5–10 minutes

### Q: What if the target machine loses connectivity during deployment?

**A:** The orchestrator script will time out and stop. Check the target machine's status:

```powershell
Test-WSMan -ComputerName $t
```

Once the machine is back online, you can resume or restart the deployment.

### Q: Can I deploy to multiple machines in parallel?

**A:** Yes, but you must run the orchestrator command for each machine separately (do not run them in the same PowerShell session as they will conflict over credentials).

---

## References

- **Atlassian Documentation:** [OnBase Provisioning Framework Implementation Decisions and Operational Guidance](https://hyland.atlassian.net/wiki/spaces/HFSS/pages/4331907517/OnBase+Provisioning+Framework+Implementation+Decisions+and+Operational+Guidance)
- **Build Repository:** `\\onbase.net\shares\Builds\OnBase\Current\25_1\25,1,28,1000`
- **SQL Server Media:** `\\onbase.net\shares\Public\SQLServer\sql_server_2022_developer.iso`

---

**Last Updated:** October 2, 2026  
**OnBase Version:** 25.1.28.1000  
**SQL Server Version:** 2022 Developer Edition  
**ODBC Driver Version:** 18
