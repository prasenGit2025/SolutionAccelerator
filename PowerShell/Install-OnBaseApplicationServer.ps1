#Requires -Version 5.1
<#
.SYNOPSIS
 Installs AppServer64 with settings-driven VC++ and protected reboot checkpoints.
.DESCRIPTION
 Companion to the reboot-aware Install-OnBaseOrchestration.ps1.
 Reads appServer.vcRedistDownloadUrl from Settings.json; NO hardcoded URL.
 Optional appServer.vcRedistPath overrides downloading. A sufficient x64
 runtime (>=14.44.0) is preserved. Download requires HTTPS and a Valid
 Microsoft Corporation Authenticode signature before prerequisite execution.

 This child NEVER initiates a reboot. It emits a structured result AND exits
 3010. The companion coordinator restarts the remote VM, waits for a changed
 boot time and the same authenticated endpoint identity, then resumes using
 the same OperationId and ResumeAfterReboot.
 After prerequisite reboot: verify runtime and continue installation.
 After uninstall reboot: continue only the remaining selected uninstalls.
 After successful MSI reboot: post-install configuration ONLY; no MSI rerun.

 ProgramData checkpoints contain no passwords/credentials. Resume validates
 operation ID, computer, non-secret input fingerprint, new boot, and cached
 MSI SHA256. Media is staged in the protected checkpoint directory BEFORE
 optional uninstall. State/cache is removed on success, retained on failure.
 No general desired-version/compatibility-based MSI skip check is added.
 Interrupted/failed/ambiguous operations are not automatically replayed.

 Original optional-uninstall product filter and post-install IIS/Web.config/
 ACL policy remain. Those writes can overwrite existing settings. Legacy
 hsiPassword/hardcoded fallback remains for compatibility; it is NOT secure
 credential handling. This child does not accept HsiCredential. A failed
 service probe is a warning, not MSI failure. Original probe-only certificate
 bypass is restored afterwards and is NEVER used for downloading. Probe
 success does not prove certificate trust. No forced MSI kill on timeout.

 Automatic reboot/resume requires the companion coordinator on a separate
 workstation. Local/standalone use reports 3010; after reboot, resume with
 -OperationId <reported ID> -ResumeAfterReboot and unchanged installer inputs.
 Do not discard pending checkpoints and blindly restart an installation.
 NOT parsed/executed in Windows PowerShell here. Back up and test in a lab.
#>
[CmdletBinding()]
param(
    [ValidateSet('Local','Remote')][string]$DeploymentMode='',
    [string]$TargetMachine='', [string]$SourceMsiPath='', [string]$ExtraMsiArgs='',
    [Parameter(Position=0)][pscredential]$Credential,
    [pscredential]$ShareCredential, [string]$ConfigurationName='',
    [switch]$UninstallExisting,
    [ValidateRange(1,1440)][int]$TimeoutMinutes=15,
    [Alias('SettingsFilePath')][string]$SettingsPath='.\Settings.json',
    [guid]$OperationId=[guid]::Empty, [switch]$ResumeAfterReboot
)
$ErrorActionPreference='Stop'
function Get-ConfigProperty {
    param($Object,[string[]]$PropertyPath,$Default=$null)
    $current=$Object
    foreach($part in $PropertyPath){
        if($null -eq $current -or -not $current.PSObject.Properties[$part]){return $Default}
        $current=$current.$part
    }
    if($null -ne $current -and -not [string]::IsNullOrWhiteSpace([string]$current)){return $current}
    return $Default
}
$settings=$null
if(Test-Path -LiteralPath $SettingsPath){
    Write-Host "[Settings] Loading configuration from: $SettingsPath" -ForegroundColor Gray
    $settings=Get-Content -LiteralPath $SettingsPath -Raw|ConvertFrom-Json -ErrorAction Stop
}
if(-not $DeploymentMode){$DeploymentMode=Get-ConfigProperty $settings @('deploymentMode') 'Remote'}
if($DeploymentMode -notin 'Local','Remote'){throw 'deploymentMode must be Local or Remote.'}
if(-not $TargetMachine){
    $TargetMachine=Get-ConfigProperty $settings @('machineName')
    if(-not $TargetMachine){$TargetMachine=Get-ConfigProperty $settings @('TargetServer')}
    if(-not $TargetMachine -and $DeploymentMode -eq 'Local'){$TargetMachine=$env:COMPUTERNAME}
}
if(-not $SourceMsiPath){
    $SourceMsiPath=Get-ConfigProperty $settings @('appServer','sourceMsiPath')
    if(-not $SourceMsiPath){$SourceMsiPath=Get-ConfigProperty $settings @('Installers','AppServerMsiPath')}
}
if(-not $ConfigurationName){$ConfigurationName=Get-ConfigProperty $settings @('configurationName') 'UnityInstall'}
if(-not $ExtraMsiArgs){$ExtraMsiArgs=Get-ConfigProperty $settings @('appServer','extraMsiArgs') 'ALLUSERS=1'}
if(-not $PSBoundParameters.ContainsKey('TimeoutMinutes')){
    $TimeoutMinutes=[int](Get-ConfigProperty $settings @('appServer','timeoutMinutes') 15)
    if($TimeoutMinutes -lt 1 -or $TimeoutMinutes -gt 1440){throw 'Invalid appServer.timeoutMinutes.'}
}
$vcRedistSource=[string](Get-ConfigProperty $settings @('appServer','vcRedistPath') '')
$vcRedistDownloadUrl=[string](Get-ConfigProperty $settings @('appServer','vcRedistDownloadUrl') '')
$targetFqdn=if($TargetMachine -match '\.'){$TargetMachine}else{"$TargetMachine.hylandqa.net"}
$dbServer=Get-ConfigProperty $settings @('sqlInstance') "$TargetMachine\SQL2K22"
$dbName=Get-ConfigProperty $settings @('datasourceName') 'OnBase'
$dbUser=Get-ConfigProperty $settings @('hsiUsername') 'hsi'
# Legacy behavior retained; this is NOT a secure secret-store design.
$dbPass=Get-ConfigProperty $settings @('hsiPassword') 'wstinol'
$idpAppName=Get-ConfigProperty $settings @('idp','appName') 'identityservice'
if(-not $SourceMsiPath){throw 'Set appServer.sourceMsiPath or pass SourceMsiPath.'}
if($ResumeAfterReboot -and $OperationId -eq [guid]::Empty){throw 'ResumeAfterReboot requires the original OperationId.'}
if($OperationId -eq [guid]::Empty){$OperationId=[guid]::NewGuid()}
if(-not $ShareCredential -and -not $ConfigurationName){$ShareCredential=$Credential}
Write-Host '=================================================================' -ForegroundColor Cyan
Write-Host ' OnBase Application Server - reboot/resume aware' -ForegroundColor Cyan
Write-Host "Deployment Mode : $DeploymentMode"
Write-Host "Target Machine : $TargetMachine ($targetFqdn)"
Write-Host "Source MSI : $SourceMsiPath"
Write-Host "Database Instance : $dbServer ($dbName)"
Write-Host "Remote Endpoint : $ConfigurationName"
Write-Host "Operation : $OperationId | Resume: $([bool]$ResumeAfterReboot)"
Write-Host 'VC++ x64 minimum : 14.44.0'
Write-Host "VC++ source : $(if($vcRedistSource){$vcRedistSource}else{'HTTPS URL from appServer.vcRedistDownloadUrl, only if required'})"
Write-Warning 'Existing IIS/Web.config/ACL writes and legacy database-password fallback remain. This child does not consume HsiCredential.'

$installScript={
    param($sourceMsi,$shareCred,$extraArgs,$timeoutMinutes,$uninstallExisting,
          $fqdn,$sqlInstance,$databaseName,$sqlUsername,$sqlPassword,$idpApp,
          $vcSource,[string]$vcDownloadUrl,[string]$operationId,[bool]$resumeAfterReboot)
    $ErrorActionPreference='Stop'
    if(-not [Environment]::Is64BitProcess){throw 'Use a 64-bit Windows PowerShell endpoint.'}
    if($operationId -notmatch '^[a-f0-9]{32}$'){throw 'Invalid operation ID.'}
    $logFile='C:\Temp\AppServerInstall.log'
    $physicalPath='C:\inetpub\wwwroot\AppServer64'
    $appPoolName='AppServer64';$appPath='/AppServer64';$siteName='Default Web Site'
    $vcBefore=$null;$vcAfter=$null;$prereqCode=$null;$prereqLog=$null;$msiExit=$null
    $state=$null;$statePath=$null;$stateDirectory=$null
    New-Item -ItemType Directory -Path 'C:\Temp' -Force|Out-Null
    function Get-BootTicks {
        $os=Get-CimInstance Win32_OperatingSystem -ErrorAction Stop
        return ([string]$os.LastBootUpTime.ToUniversalTime().Ticks)
    }
    function Protect-Directory([string]$Path){
        if(-not (Test-Path -LiteralPath $Path)){New-Item -ItemType Directory -Path $Path -ErrorAction Stop|Out-Null}
        $item=Get-Item -LiteralPath $Path -Force
        if(-not $item.PSIsContainer -or ($item.Attributes -band [IO.FileAttributes]::ReparsePoint)){throw "Refusing non-directory/reparse point: $Path"}
        $acl=[Security.AccessControl.DirectorySecurity]::new();$acl.SetAccessRuleProtection($true,$false)
        $admin=[Security.Principal.SecurityIdentifier]::new('S-1-5-32-544');$acl.SetOwner($admin)
        foreach($sidText in @('S-1-5-18','S-1-5-32-544')){
            $sid=[Security.Principal.SecurityIdentifier]::new($sidText)
            $rule=[Security.AccessControl.FileSystemAccessRule]::new(
                $sid,[Security.AccessControl.FileSystemRights]::FullControl,
                [Security.AccessControl.InheritanceFlags]'ContainerInherit, ObjectInherit',
                [Security.AccessControl.PropagationFlags]::None,[Security.AccessControl.AccessControlType]::Allow)
            [void]$acl.AddAccessRule($rule)
        }
        Set-Acl -LiteralPath $Path -AclObject $acl -ErrorAction Stop
    }
    function Save-State {
        # No passwords, credentials or raw argument strings are persisted.
        $temporary=Join-Path $stateDirectory ('state_'+[guid]::NewGuid().ToString('N')+'.tmp')
        $state|ConvertTo-Json -Depth 8|Set-Content -LiteralPath $temporary -Encoding UTF8 -ErrorAction Stop
        Move-Item -LiteralPath $temporary -Destination $statePath -Force -ErrorAction Stop
    }
    function New-StageResult {
        param([string]$Stage,[int]$Code,[string]$Message,[string]$ProbeUrl='',[string]$ProbeStatus='Not run')
        [pscustomobject]@{
            Computer=$env:COMPUTERNAME;Stage=$Stage;ExitCode=$Code;MsiExitCode=$msiExit
            RebootRequired=($Code -eq 3010);ResumeSupported=$true
            ResumeAvailable=($Code -eq 3010 -and $null -ne $state)
            OperationId=$operationId;CheckpointPath=$statePath;Message=$Message
            VcVersionBefore=[string]$vcBefore;VcVersionAfter=[string]$vcAfter
            PrerequisiteExitCode=$prereqCode;PrerequisiteLog=$prereqLog
            ProbeUrl=$ProbeUrl;ProbeStatus=$ProbeStatus;Log=$logFile
        }
    }
    function Invoke-Msi([string]$Arguments,[int]$Minutes){
        $process=Start-Process -FilePath 'msiexec.exe' -ArgumentList $Arguments -PassThru -ErrorAction Stop
        $null=$process.Handle
        if($process.WaitForExit($Minutes*60000)){$process.WaitForExit();return [int]$process.ExitCode}
        # Do not kill a process while Windows Installer may still service files.
        throw "MSI wait exceeded $Minutes minutes. PID $($process.Id) was NOT killed. Inspect servicing and logs before any retry/restart."
    }
    function Get-VcX64RuntimeVersion {
        $versions=[Collections.Generic.List[Version]]::new()
        foreach($view in @([Microsoft.Win32.RegistryView]::Registry64,[Microsoft.Win32.RegistryView]::Registry32)){
            $base=$null;$key=$null
            try{
                $base=[Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::LocalMachine,$view)
                $key=$base.OpenSubKey('SOFTWARE\Microsoft\VisualStudio\14.0\VC\Runtimes\x64')
                if($null -eq $key -or [int]$key.GetValue('Installed',0) -ne 1){continue}
                [version]$parsed=$null
                $text=([string]$key.GetValue('Version','')).Trim() -replace '^[vV]',''
                if([version]::TryParse($text,[ref]$parsed)){[void]$versions.Add($parsed)}else{
                    $major=$key.GetValue('Major',$null);$minor=$key.GetValue('Minor',$null)
                    $build=$key.GetValue('Bld',$null);$revision=$key.GetValue('Rbld',0)
                    if($null -ne $major -and $null -ne $minor -and $null -ne $build){
                        $text='{0}.{1}.{2}.{3}' -f $major,$minor,$build,$revision
                        if([version]::TryParse($text,[ref]$parsed)){[void]$versions.Add($parsed)}
                    }
                }
            }finally{if($key){$key.Dispose()};if($base){$base.Dispose()}}
        }
        if($versions.Count -gt 0){return ($versions|Sort-Object -Descending|Select-Object -First 1)}
        return $null
    }
    function Copy-Source {
        param([string]$Source,[string]$Destination,[pscredential]$SourceCredential)
        $drive=$null
        try{
            $readPath=$Source
            if($Source.StartsWith('\\') -and $SourceCredential){
                $parts=$Source.TrimStart('\').Split('\')
                if($parts.Count -lt 3){throw 'Source must identify a file, not a share root.'}
                $root='\\{0}\{1}' -f $parts[0],$parts[1];$relative=$parts[2..($parts.Count-1)] -join '\'
                $drive='OBSource'+[guid]::NewGuid().ToString('N').Substring(0,8)
                New-PSDrive -Name $drive -PSProvider FileSystem -Root $root -Credential $SourceCredential -Scope Local -ErrorAction Stop|Out-Null
                $readPath='{0}:\{1}' -f $drive,$relative
            }
            if(-not (Test-Path -LiteralPath $readPath -PathType Leaf)){throw "Source unreadable on $env:COMPUTERNAME`: $Source"}
            Copy-Item -LiteralPath $readPath -Destination $Destination -Force -ErrorAction Stop
            $sourceHash=(Get-FileHash -LiteralPath $readPath -Algorithm SHA256 -ErrorAction Stop).Hash
            $localHash=(Get-FileHash -LiteralPath $Destination -Algorithm SHA256 -ErrorAction Stop).Hash
            if($sourceHash -ne $localHash){throw 'Staged file hash does not match source.'}
            return $localHash
        }finally{if($drive){Remove-PSDrive -Name $drive -Force -ErrorAction SilentlyContinue}}
    }
    try{
        $root=Join-Path $env:ProgramData 'OnBaseProvisioning-AppServerResume';Protect-Directory $root
        $stateDirectory=Join-Path $root $operationId;$statePath=Join-Path $stateDirectory 'state.json'
        # Fingerprint excludes passwords; raw MSI arguments are not persisted.
        $inputs=@($sourceMsi,$extraArgs,[string]$uninstallExisting,$fqdn,$sqlInstance,$databaseName,$sqlUsername,$idpApp,$vcSource,$vcDownloadUrl)
        $json=ConvertTo-Json -InputObject $inputs -Compress
        $sha=[Security.Cryptography.SHA256]::Create()
        try{$fingerprint=[BitConverter]::ToString($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($json))).Replace('-','')}finally{$sha.Dispose()}
        if($resumeAfterReboot){
            if(-not (Test-Path -LiteralPath $statePath -PathType Leaf)){throw 'No checkpoint for operation; refusing to skip installation.'}
            Protect-Directory $stateDirectory
            $state=Get-Content -LiteralPath $statePath -Raw|ConvertFrom-Json -ErrorAction Stop
            if($state.Schema -ne 1 -or $state.OperationId -ne $operationId -or $state.Computer -ine $env:COMPUTERNAME){throw 'Checkpoint schema/identity mismatch.'}
            if($state.InputFingerprint -ne $fingerprint){throw 'Installer inputs changed since checkpoint; resume refused.'}
            if($state.Phase -notin 'PrerequisitePendingReboot','UninstallPendingReboot','MsiPendingReboot'){throw "Checkpoint '$($state.Phase)' is not a reboot continuation."}
            if([string]$state.BootTicks -eq (Get-BootTicks)){throw 'No changed boot time; resume refused.'}
            $msiExit=$state.MsiExitCode;$prereqCode=$state.PrerequisiteExitCode;$prereqLog=$state.PrerequisiteLog
            Write-Host "[$env:COMPUTERNAME] Resuming: $($state.Phase)" -ForegroundColor Cyan
        }else{
            if(Test-Path -LiteralPath $stateDirectory){throw 'Operation exists. Inspect/resume the original operation after reboot; do not replay it.'}
            foreach($directory in @(Get-ChildItem -LiteralPath $root -Directory -ErrorAction Stop)){
                if($directory.Attributes -band [IO.FileAttributes]::ReparsePoint){throw 'Unexpected reparse point in resume storage.'}
                $otherPath=Join-Path $directory.FullName 'state.json'
                if(Test-Path -LiteralPath $otherPath -PathType Leaf){
                    $other=Get-Content -LiteralPath $otherPath -Raw|ConvertFrom-Json -ErrorAction Stop
                    if($other.Phase -in 'PrerequisitePendingReboot','UninstallPendingReboot','MsiPendingReboot',
                        'PrerequisiteInProgress','UninstallInProgress','MsiInProgress','ConfigurationInProgress'){
                        throw "Pending/incomplete operation $($other.OperationId) exists at $otherPath. Inspect it; resume only supported reboot checkpoints. Do not abandon an in-progress installer."
                    }
                }
            }
            Protect-Directory $stateDirectory
            $state=[pscustomobject]@{
                Schema=1;OperationId=$operationId;Computer=$env:COMPUTERNAME;InputFingerprint=$fingerprint
                Phase='Started';BootTicks=(Get-BootTicks);MsiHash='';MsiExitCode=$null
                RemainingUninstalls=@();PrerequisiteExitCode=$null;PrerequisiteLog=$null
            };Save-State
        }
        $mediaDirectory=Join-Path $stateDirectory 'Media';Protect-Directory $mediaDirectory
        $localMsi=Join-Path $mediaDirectory (Split-Path $sourceMsi -Leaf)
        $phase=[string]$state.Phase;$configureOnly=($phase -eq 'MsiPendingReboot')
        if($configureOnly -and [int]$state.MsiExitCode -ne 3010){throw 'Checkpoint lacks the successful MSI reboot result.'}
        if($state.MsiHash){
            if(-not (Test-Path -LiteralPath $localMsi -PathType Leaf) -or
                (Get-FileHash -LiteralPath $localMsi -Algorithm SHA256 -ErrorAction Stop).Hash -ne $state.MsiHash){throw 'Staged MSI missing/changed; resume refused.'}
        }
        $minimum=[version]'14.44.0';$vcBefore=Get-VcX64RuntimeVersion;$vcAfter=$vcBefore
        if($vcBefore -and $vcBefore -ge $minimum){Write-Host "[$env:COMPUTERNAME] VC++ x64 $vcBefore meets $minimum; prerequisite skipped." -ForegroundColor Green}
        elseif($resumeAfterReboot){throw 'VC++ minimum absent after reboot; no repeated prerequisite installation.'}
        else{
            $state.Phase='PrerequisiteInProgress';Save-State
            $stageDirectory=Join-Path $stateDirectory 'VcPrerequisite';Protect-Directory $stageDirectory
            $vcExe=Join-Path $stageDirectory 'vc_redist.x64.exe'
            $prereqLog=Join-Path 'C:\Temp' ('vc_redist_x64_{0}.log' -f (Get-Date -Format 'yyyyMMdd_HHmmss_fff'))
            $savedCallback=[Net.ServicePointManager]::ServerCertificateValidationCallback
            $savedProtocol=[Net.ServicePointManager]::SecurityProtocol
            try{
                [Net.ServicePointManager]::ServerCertificateValidationCallback=$null
                [Net.ServicePointManager]::SecurityProtocol=[Net.SecurityProtocolType]::Tls12
                if(-not [string]::IsNullOrWhiteSpace($vcSource)){
                    Write-Host "[$env:COMPUTERNAME] Using configured VC++ prerequisite: $vcSource"
                    $null=Copy-Source $vcSource $vcExe $shareCred
                }else{
                    [uri]$downloadUri=$null
                    if([string]::IsNullOrWhiteSpace($vcDownloadUrl) -or
                        -not [uri]::TryCreate($vcDownloadUrl.Trim(),[UriKind]::Absolute,[ref]$downloadUri)){
                        throw 'Set appServer.vcRedistDownloadUrl to an absolute HTTPS URL (or configure vcRedistPath).'
                    }
                    if($downloadUri.Scheme -ne 'https' -or -not $downloadUri.Host -or $downloadUri.UserInfo){throw 'Prerequisite URL requires HTTPS, a host, and no embedded credentials.'}
                    Write-Host "[$env:COMPUTERNAME] Downloading VC++ x64 from configured HTTPS source..."
                    Invoke-WebRequest -Uri $downloadUri.AbsoluteUri -OutFile $vcExe -UseBasicParsing -TimeoutSec 600 -ErrorAction Stop|Out-Null
                }
                $signature=Get-AuthenticodeSignature -LiteralPath $vcExe -ErrorAction Stop
                if($signature.Status -ne [System.Management.Automation.SignatureStatus]::Valid -or $null -eq $signature.SignerCertificate){throw "Invalid prerequisite signature: $($signature.Status). $($signature.StatusMessage)"}
                $publisher=$signature.SignerCertificate.GetNameInfo([Security.Cryptography.X509Certificates.X509NameType]::SimpleName,$false)
                if($publisher -ne 'Microsoft Corporation'){throw "Prerequisite signer '$publisher' is not Microsoft Corporation."}
                Write-Host "[$env:COMPUTERNAME] Valid Microsoft Corporation signature verified." -ForegroundColor Green
                $arguments='/install /quiet /norestart /log "{0}"' -f $prereqLog
                $process=Start-Process -FilePath $vcExe -ArgumentList $arguments -PassThru -ErrorAction Stop
                $null=$process.Handle;$process.WaitForExit()
                $prereqCode=[int]$process.ExitCode;$vcAfter=Get-VcX64RuntimeVersion
                $state.PrerequisiteExitCode=$prereqCode;$state.PrerequisiteLog=$prereqLog
                if($prereqCode -in 3010,1641){
                    $state.Phase='PrerequisitePendingReboot';$state.BootTicks=Get-BootTicks;Save-State
                    return New-StageResult 'PrerequisiteRebootRequired' 3010 'VC++ requests reboot. No Application Server uninstall/MSI run. Resume this operation after reboot.' '' 'Deferred until reboot'
                }
                if($prereqCode -notin 0,1638){throw "VC++ install failed: $prereqCode. See $prereqLog"}
                if(-not $vcAfter -or $vcAfter -lt $minimum){throw "VC++ minimum not verified after exit $prereqCode. See $prereqLog"}
            }finally{
                [Net.ServicePointManager]::ServerCertificateValidationCallback=$savedCallback
                [Net.ServicePointManager]::SecurityProtocol=$savedProtocol
                if(Test-Path -LiteralPath $stageDirectory){Remove-Item -LiteralPath $stageDirectory -Recurse -Force -ErrorAction SilentlyContinue}
            }
        }
        if(-not $configureOnly){
            if(-not $state.MsiHash){
                Write-Host "[$env:COMPUTERNAME] Staging and verifying MSI before optional uninstall..."
                $state.MsiHash=Copy-Source $sourceMsi $localMsi $shareCred
            }
            if($phase -ne 'UninstallPendingReboot'){
                $state.RemainingUninstalls=@()
                if($uninstallExisting){
                    $keys='HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*',
                          'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*',
                          'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*'
                    $existing=@(Get-ItemProperty $keys -ErrorAction SilentlyContinue|
                        Where-Object {$_.DisplayName -like 'OnBase*Application Server*' -and $_.PSChildName -match '^\{.+\}$'}|
                        Select-Object -ExpandProperty PSChildName -Unique)
                    $state.RemainingUninstalls=@($existing)
                }
            }
            $state.Phase='UninstallInProgress';Save-State
            foreach($productCode in @($state.RemainingUninstalls)){
                if($productCode -notmatch '^\{[a-fA-F0-9-]{36}\}$'){throw 'Invalid uninstall product code in checkpoint.'}
                Write-Host "[$env:COMPUTERNAME] Uninstalling selected product $productCode..."
                $arguments='/x {0} /qn /norestart /L*v "C:\Temp\AppServerUninstall.log"' -f $productCode
                $uninstallCode=Invoke-Msi $arguments $timeoutMinutes
                if($uninstallCode -notin 0,1605,3010){throw "Uninstall failed: $uninstallCode; installation not attempted."}
                $state.RemainingUninstalls=@($state.RemainingUninstalls|Where-Object {$_ -ne $productCode})
                if($uninstallCode -eq 3010){
                    $state.Phase='UninstallPendingReboot';$state.BootTicks=Get-BootTicks;Save-State
                    return New-StageResult 'UninstallRebootRequired' 3010 'Uninstall requests reboot; completed removals will not repeat on resume.'
                };Save-State
            }
            if(Test-Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\InProgress'){throw 'Windows Installer operation in progress; no new MSI started.'}
            $state.Phase='MsiInProgress';Save-State
            $finalArgs=$extraArgs
            if($finalArgs -notmatch 'TARGETDIR='){$finalArgs="TARGETDIR=`"$physicalPath`" ALLUSERS=1 $finalArgs".Trim()}
            $arguments='/i "{0}" /qn /norestart /L*v "{1}" {2}' -f $localMsi,$logFile,$finalArgs
            Write-Host "[$env:COMPUTERNAME] Installing Application Server (timeout $timeoutMinutes minutes)..."
            $msiExit=Invoke-Msi $arguments $timeoutMinutes;$state.MsiExitCode=$msiExit
            if($msiExit -notin 0,3010){
                $state.Phase='MsiFailed';Save-State
                return New-StageResult 'MsiFailed' $msiExit 'MSI failed. IIS/Web.config/ACL steps skipped.'
            }
            if($msiExit -eq 3010){
                $state.Phase='MsiPendingReboot';$state.BootTicks=Get-BootTicks;Save-State
                return New-StageResult 'MsiRebootRequired' 3010 'MSI succeeded; reboot required. Resume runs configuration only, NOT the MSI.'
            }
        }else{Write-Host "[$env:COMPUTERNAME] MSI already succeeded. POST-INSTALL ONLY; no uninstall/MSI rerun." -ForegroundColor Green}
        $state.Phase='ConfigurationInProgress';Save-State
        # Original configuration policy: can overwrite existing environment values.
        New-Item -ItemType Directory -Path $physicalPath -Force|Out-Null
        Import-Module WebAdministration -ErrorAction Stop
        if(-not (Test-Path "IIS:\AppPools\$appPoolName")){New-Item -Path "IIS:\AppPools\$appPoolName"|Out-Null}
        Set-ItemProperty "IIS:\AppPools\$appPoolName" -Name managedRuntimeVersion -Value 'v4.0'
        Set-ItemProperty "IIS:\AppPools\$appPoolName" -Name managedPipelineMode -Value 0
        Set-ItemProperty "IIS:\AppPools\$appPoolName" -Name enable32BitAppOnWin64 -Value $false
        Set-ItemProperty "IIS:\AppPools\$appPoolName" -Name processModel.idleTimeout -Value ([TimeSpan]::FromMinutes(0))
        $existingApp=Get-WebApplication -Site $siteName|Where-Object {$_.Path -eq $appPath}
        if(-not $existingApp){New-WebApplication -Site $siteName -Name ($appPath.TrimStart('/')) -PhysicalPath $physicalPath -ApplicationPool $appPoolName|Out-Null}
        else{
            Set-ItemProperty "IIS:\Sites\$siteName$appPath" -Name applicationPool -Value $appPoolName
            Set-ItemProperty "IIS:\Sites\$siteName$appPath" -Name physicalPath -Value $physicalPath
        }
        $webConfigPath=Join-Path $physicalPath 'Web.config'
        if(Test-Path -LiteralPath $webConfigPath){
            [xml]$webXml=Get-Content -LiteralPath $webConfigPath -Raw
            $sqlConnString="Data Source=$sqlInstance;Initial Catalog=$databaseName;User ID=$sqlUsername;Password=$sqlPassword;Encrypt=Optional;TrustServerCertificate=True;"
            if(-not $webXml.configuration.connectionStrings){
                $element=$webXml.CreateElement('connectionStrings');$webXml.configuration.AppendChild($element)|Out-Null
            }
            $connection=$webXml.configuration.connectionStrings.SelectSingleNode("add[@name='OnBase']")
            if(-not $connection){
                $connection=$webXml.CreateElement('add');$connection.SetAttribute('name','OnBase');$connection.SetAttribute('providerName','System.Data.SqlClient')
                $webXml.configuration.connectionStrings.AppendChild($connection)|Out-Null
            }
            $connection.SetAttribute('connectionString',$sqlConnString)
            $appSettings=$webXml.configuration.appSettings
            if($appSettings){
                $idsKey=$appSettings.SelectSingleNode("add[@key='IdentityServerAddress' or @key='Hyland.IdentityService.Url']")
                if($idsKey){$idsKey.SetAttribute('value',"https://$fqdn/$idpApp")}
            };$webXml.Save($webConfigPath)
        }
        $folders=@($physicalPath,(Join-Path $physicalPath 'App_Data'),(Join-Path $physicalPath 'Logs'),'C:\ProgramData\Hyland\TempFS')
        $inherit=[Security.AccessControl.InheritanceFlags]'ContainerInherit, ObjectInherit'
        $prop=[Security.AccessControl.PropagationFlags]::None
        foreach($folder in $folders){
            if(-not (Test-Path -LiteralPath $folder)){New-Item -ItemType Directory -Path $folder -Force|Out-Null}
            $acl=Get-Acl -LiteralPath $folder
            foreach($account in @('BUILTIN\Administrators','NT AUTHORITY\SYSTEM','BUILTIN\IIS_IUSRS')){
                $rule=[Security.AccessControl.FileSystemAccessRule]::new($account,'FullControl',$inherit,$prop,'Allow');[void]$acl.AddAccessRule($rule)
            }
            $rule=[Security.AccessControl.FileSystemAccessRule]::new("IIS AppPool\$appPoolName",'Modify',$inherit,$prop,'Allow');[void]$acl.AddAccessRule($rule)
            Set-Acl -LiteralPath $folder -AclObject $acl
        }
        Restart-WebAppPool -Name $appPoolName;Start-Sleep -Seconds 3
        $probeUrl="https://$fqdn/AppServer64/Service.asmx";$probeStatus='Unknown'
        $previousProtocol=[Net.ServicePointManager]::SecurityProtocol
        $previousCallback=[Net.ServicePointManager]::ServerCertificateValidationCallback
        try{
            [Net.ServicePointManager]::SecurityProtocol=[Net.SecurityProtocolType]::Tls12
            # Legacy probe-only bypass. Never applied to downloads/signature checks.
            [Net.ServicePointManager]::ServerCertificateValidationCallback={$true}
            try{
                $request=[Net.HttpWebRequest]::Create($probeUrl);$request.Timeout=10000;$response=$request.GetResponse()
                try{$probeStatus="$([int]$response.StatusCode) OK"}finally{$response.Close()}
            }catch [Net.WebException]{
                $exception=$_.Exception
                if($exception.Response){try{$probeStatus="$([int]$exception.Response.StatusCode) Response: $($exception.Message)"}finally{$exception.Response.Close()}}
                else{$probeStatus="$($exception.Status): $($exception.Message)"}
            }catch{$probeStatus="Error: $($_.Exception.Message)"}
        }finally{
            [Net.ServicePointManager]::ServerCertificateValidationCallback=$previousCallback
            [Net.ServicePointManager]::SecurityProtocol=$previousProtocol
        }
        if($probeStatus -notmatch '^2\d\d OK$'){Write-Warning "Install/configuration completed; service probe did not confirm health: $probeStatus"}
        $state.Phase='Completed';Save-State
        $completed=New-StageResult 'Completed' 0 'Application Server install/configuration completed. Inspect service probe separately.' $probeUrl $probeStatus
        try{Remove-Item -LiteralPath $stateDirectory -Recurse -Force -ErrorAction Stop}
        catch{Write-Warning "Completed, but checkpoint/cache cleanup failed: $stateDirectory. $($_.Exception.Message)"}
        return $completed
    }catch{
        # Do not automatically replay an interrupted/ambiguous operation.
        return New-StageResult 'Failed' 1 "Application Server failed: $($_.Exception.Message). Inspect checkpoint/logs before retry."
    }
}
# Fifteen arguments match all fifteen target-side parameters in order.
$scriptArgs=@($SourceMsiPath,$ShareCredential,$ExtraMsiArgs,$TimeoutMinutes,
    [bool]$UninstallExisting,$targetFqdn,$dbServer,$dbName,$dbUser,$dbPass,$idpAppName,
    $vcRedistSource,$vcRedistDownloadUrl,$OperationId.ToString('N'),[bool]$ResumeAfterReboot)
try{
    if($DeploymentMode -eq 'Local'){
        $principal=[Security.Principal.WindowsPrincipal]::new([Security.Principal.WindowsIdentity]::GetCurrent())
        if(-not $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)){throw 'Run elevated.'}
        $result=& $installScript @scriptArgs
    }else{
        if(-not $TargetMachine){throw 'TargetMachine required for Remote mode.'}
        $connection=@{ComputerName=$TargetMachine;ScriptBlock=$installScript;ArgumentList=$scriptArgs;ErrorAction='Stop'}
        if($Credential){$connection.Credential=$Credential;$connection.Authentication='Negotiate'}
        if($ConfigurationName){$connection.ConfigurationName=$ConfigurationName}
        $result=Invoke-Command @connection
    }
}catch{Write-Host "Application Server step failed: $($_.Exception.Message)" -ForegroundColor Red;exit 1}
$result=$result|Where-Object {$_ -and $_.PSObject.Properties['ExitCode']}|Select-Object -Last 1
if(-not $result){Write-Host 'No result from target.' -ForegroundColor Red;exit 1}
Write-Host "`n=================================================================" -ForegroundColor Cyan
Write-Host "[$($result.Computer)] Stage: $($result.Stage) | Overall exit: $($result.ExitCode) | MSI exit: $($result.MsiExitCode)"
Write-Host "VC++ before/after: $($result.VcVersionBefore) / $($result.VcVersionAfter)"
Write-Host $result.Message
Write-Host "Probe: $($result.ProbeUrl) -> $($result.ProbeStatus)"
Write-Host "MSI log: $($result.Log) (current run only if MSI invoked)"
if($result.PrerequisiteLog){Write-Host "Prerequisite log: $($result.PrerequisiteLog)"}
if($result.ExitCode -ne 0){Write-Host "Operation: $($result.OperationId) | Checkpoint: $($result.CheckpointPath)"}
# Mandatory result envelope for companion coordinator; do not suppress.
Write-Output $result
if($result.ExitCode -eq 3010){exit 3010}
if($result.ExitCode -ne 0){exit 1}
exit 0
