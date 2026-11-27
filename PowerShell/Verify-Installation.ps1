param(
    [ValidateSet("Local", "Remote")]
    [string]$DeploymentMode = "Local",

    [string]$TargetMachine = "",

    [pscredential]$Credential,

    [string]$ConfigurationName = "UnityInstall",

    [string]$SqlInstance = "RDV-009989\SQL2K22",

    [string]$OdbcDriver = "ODBC Driver 18 for SQL Server",

    [string]$DsnName = "OnBase",

    [pscredential]$HsiCredential
)

Write-Host "Verifying Installation..." -ForegroundColor Cyan

# Load Environment Configuration
$configPath = Join-Path $PSScriptRoot "..\Config\Environment.Dev.json"
$config = $null
if (Test-Path $configPath) {
    $config = Get-Content $configPath -Raw | ConvertFrom-Json
}

# Verification script block executed on local or remote machine
$verificationBlock = {
    param($cfg, $sqlInst, $driverName, $dsn, $hsiCred)

    Write-Host "Running verification on: $env:COMPUTERNAME" -ForegroundColor Yellow

    # ==========================================
    # 1. SQL Server Instance Service Check
    # ==========================================
    Write-Host "`n--- Checking SQL Server ---" -ForegroundColor Cyan
    if ($sqlInst) {
        $instanceName = if ($sqlInst -like '*\*') { $sqlInst.Split('\')[1] } else { $sqlInst }
        $serviceName = if ($instanceName -eq 'MSSQLSERVER') { 'MSSQLSERVER' } else { "MSSQL`$$instanceName" }

        $svc = Get-Service -Name $serviceName -ErrorAction SilentlyContinue
        if ($svc -and $svc.Status -eq 'Running') {
            Write-Host "  [OK] SQL Server service '$serviceName' is Running." -ForegroundColor Green
        } elseif ($svc) {
            Write-Warning "  [WARN] SQL Server service '$serviceName' exists but is in status: $($svc.Status)"
        } else {
            Write-Warning "  [FAIL] SQL Server service '$serviceName' was not found."
        }
    }

    # ==========================================
    # 2. ODBC Driver Verification (32-bit & 64-bit)
    # ==========================================
    Write-Host "`n--- Checking ODBC Drivers ---" -ForegroundColor Cyan
    if ($driverName) {
        # Check 64-bit registry
        $reg64 = "HKLM:\SOFTWARE\ODBC\ODBCINST.INI\ODBC Drivers"
        $has64 = (Get-ItemProperty -Path $reg64 -Name $driverName -ErrorAction SilentlyContinue) -ne $null

        # Check 32-bit registry (WoW6432Node)
        $reg32 = "HKLM:\SOFTWARE\WOW6432Node\ODBC\ODBCINST.INI\ODBC Drivers"
        $has32 = (Get-ItemProperty -Path $reg32 -Name $driverName -ErrorAction SilentlyContinue) -ne $null

        if ($has32 -and $has64) {
            Write-Host "  [OK] '$driverName' is installed for both 32-bit and 64-bit." -ForegroundColor Green
        } elseif ($has32) {
            Write-Host "  [OK] '$driverName' is installed for 32-bit (OnBase compatible)." -ForegroundColor Green
        } else {
            Write-Warning "  [FAIL] '$driverName' was not found in 32-bit ODBC Drivers registry."
        }
    }

    # ==========================================
    # 3. ODBC System DSN Verification
    # ==========================================
    Write-Host "`n--- Checking System DSN ---" -ForegroundColor Cyan
    if ($dsn) {
        $dsn32 = Get-ItemProperty "HKLM:\SOFTWARE\WOW6432Node\ODBC\ODBC.INI\$dsn" -ErrorAction SilentlyContinue
        if ($dsn32) {
            Write-Host "  [OK] 32-bit System DSN '$dsn' verified (Driver: $($dsn32.Driver), Server: $($dsn32.Server))." -ForegroundColor Green
        } else {
            Write-Warning "  [FAIL] 32-bit System DSN '$dsn' not found."
        }
    }

    # ==========================================
    # 4. Client Components Check
    # ==========================================
    Write-Host "`n--- Checking Client Components ---" -ForegroundColor Cyan
    if ($cfg) {
        $unityPath  = $cfg.InstallPaths.UnityClient
        $studioPath = $cfg.InstallPaths.Studio
        $cfgPath    = $cfg.InstallPaths.Configuration

        # Unity Client
        if ($unityPath -and (Test-Path $unityPath)) {
            Write-Host "  [OK] Unity Client verified: $unityPath" -ForegroundColor Green
        } else {
            Write-Warning "  [FAIL] Unity Client not found at: $unityPath"
        }

        # Studio
        if ($studioPath -and (Test-Path $studioPath)) {
            Write-Host "  [OK] Studio verified: $studioPath" -ForegroundColor Green
        } else {
            Write-Warning "  [INFO] Studio not found at: $studioPath (Skipped if not installed in this run)"
        }

        # Configuration
        if ($cfgPath -and (Test-Path $cfgPath)) {
            Write-Host "  [OK] Configuration verified: $cfgPath" -ForegroundColor Green
        } else {
            Write-Warning "  [INFO] Configuration not found at: $cfgPath (Skipped if not installed in this run)"
        }

        # Application Server Reachability (Bypassing SSL trust error for internal QA certs)
        if ($cfg.ApplicationServer.Url) {
            try {
                [System.Net.ServicePointManager]::ServerCertificateValidationCallback = {$true}
                $response = Invoke-WebRequest -Uri $cfg.ApplicationServer.Url -Method Head -UseBasicParsing -TimeoutSec 10 -ErrorAction Stop
                Write-Host "  [OK] Application Server reachable ($($cfg.ApplicationServer.Url))." -ForegroundColor Green
            } catch {
                Write-Warning "  [WARN] Application Server not reachable: $_"
            }
        }
    }
}

# Execution Router
if ($DeploymentMode -eq "Local") {
    Write-Host "Checking local installation..." -ForegroundColor Cyan
    & $verificationBlock $config $SqlInstance $OdbcDriver $DsnName $HsiCredential
} else {
    if (-not $TargetMachine) {
        Write-Error "TargetMachine must be specified when DeploymentMode is 'Remote'."
        return
    }

    Write-Host "Checking remote installation on $TargetMachine..." -ForegroundColor Cyan

    $invokeParams = @{
        ComputerName = $TargetMachine
        ScriptBlock  = $verificationBlock
        ArgumentList = @($config, $SqlInstance, $OdbcDriver, $DsnName, $HsiCredential)
    }

    if ($Credential) {
        $invokeParams["Credential"] = $Credential
    }

    if ($ConfigurationName) {
        $invokeParams["ConfigurationName"] = $ConfigurationName
    }

    Invoke-Command @invokeParams
}

Write-Host "`nVerification Completed." -ForegroundColor Cyan