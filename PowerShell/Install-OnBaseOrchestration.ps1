#Requires -Version 5.1
<#
.SYNOPSIS
 Seven-stage OnBase coordinator with remote reboot/wait and AppServer resume.
.DESCRIPTION
 Preserves filename-based skips, settings-first inputs and the seven children:
 1 Install-SqlServer, 2 Create-ODBC, 3 Initialize-OnBaseDatabase,
 4 Install-HylandIdSAndApi, 5 Install-OnBaseApplicationServer,
 6 Install-Unity, 7 Verify-Installation. Skipped stages are not invoked.
 No independent readiness script or service-location stage is added.

 Default automatic remote reboot policy: installer exit 3010 triggers one
 restart request, then synchronous polling at 15-second intervals between
 attempts, with a 1200-second recovery budget; maximum three reboots per stage,
 ten per run. CLI overrides optional reboot.* JSON overrides defaults.
 Waiting blocks the coordinator: no next installer runs during recovery.
 Keep this workstation/session running; this is not background monitoring or
 a coordinator-crash-safe workflow. Per-call cleanup may extend the budget.

 Recovery uses the existing named endpoint, NOT endpoint re-registration.
 Requires changed OS boot time, same computer/endpoint identity SID, and two
 successful consecutive authenticated probes. Watched SQL/IIS services that
 were Running before reboot must return to Running. Service-state checks do
 not replace application/database health checks in the child installers.
 No ICMP dependency. No further restart is sent if acknowledgement is lost.

 The companion AppServer child exposes protected resume checkpoints. After
 MSI reboot it runs post-install configuration WITHOUT repeating the MSI;
 completed uninstalls are not repeated; VC++ is rechecked after its reboot.
 Other children must emit RebootContinuation='NextStage' AND exit 3010 to
 assert that their stage completed. Otherwise reboot/recovery takes place,
 but continuation stops instead of guessing whether an opaque legacy child
 finished. Other child scripts are NOT rewritten by this revision. Swallowed
 errors/exit codes cannot be recovered here.

 If the PREVIOUS installer already stopped with 3010, use RestartBeforeInstall
 ONCE to satisfy that pending reboot before starting enabled stages. Remove
 it afterwards. It is an explicit CLI action, not loaded from JSON.
 A stopped NEW AppServer operation can be resumed using the reported GUID in
 AppServerResumeOperationId AFTER the required reboot; unchanged inputs are
 verified by the child. Do not discard pending checkpoints/replay installs.

 Automatic reboot works only with the coordinator on a DIFFERENT workstation.
 Local mode returns 3010, never restarts its own coordinator host. Restart uses
 shutdown.exe /r /t 10 /f: target services/applications are interrupted and
 unsaved target work can be lost. Use an approved maintenance window.

 Existing endpoint-bootstrap policy is retained: BootstrapUnityInstall=true
 by default creates only a confirmed missing fixed-RunAs endpoint through
 Microsoft.PowerShell, then uses a temporary SYSTEM task to restart WinRM.
 Reachable endpoints are reused unchanged. BootstrapUnityInstall=false blocks
 endpoint creation/separate WinRM restart but NOT installer-required VM reboot.
 No remoting/firewall/TrustedHosts/account-password changes are added.
 Coordinator credentials/passwords are not loaded from JSON or persisted.
 Existing child IIS/Web.config/ACL/credential behaviors remain unchanged.

 NOT parsed/executed in Windows PowerShell here. Back up/syntax-check BOTH
 companion files and test in an approved lab before deployment.
#>
[CmdletBinding()]
param(
 [string]$SettingsFilePath,
 [ValidateSet('Local','Remote')][string]$DeploymentMode='Remote',
 [string]$TargetMachine,[pscredential]$Credential,[pscredential]$ShareCredential,
 [string]$ConfigurationName='UnityInstall',[switch]$BootstrapUnityInstall=$true,
 [pscredential]$EndpointRunAsCredential,
 [ValidateRange(30,900)][int]$EndpointBootstrapTimeoutSeconds=180,
 [switch]$AutoReboot=$true,
 [ValidateRange(5,300)][int]$RebootPollIntervalSeconds=15,
 [ValidateRange(60,7200)][int]$RebootTimeoutSeconds=1200,
 [ValidateRange(1,10)][int]$MaxRebootsPerStage=3,
 [ValidateRange(1,30)][int]$MaxRebootsPerRun=10,
 [switch]$RestartBeforeInstall,[guid]$AppServerResumeOperationId=[guid]::Empty,
 [string]$BuildPath,[string]$SqlSetupPath,[string]$OdbcSetupPath,
 [string]$DbUtilsSourcePath,[string]$UnityMsiPath,[string]$AppServerMsiPath,
 [ValidateSet('2019','2022')][string]$SqlVersion='2022',[string]$SqlInstance,
 [ValidateSet('17','18')][string]$OdbcVersion='18',[string]$OdbcDriverName,
 [ValidateSet('SideBySide','Skip','Fail')][string]$OnSqlVersionMismatch='SideBySide',
 [pscredential]$SaCredential,[pscredential]$SqlAdminCredential,[pscredential]$HsiCredential,
 [string]$DatabaseName='OnBase',[Alias('DataSource')][string]$DsnName,
 [string]$IdSAppName='identityservice',[string]$ApiAppName='apiserver',
 [string]$CertificateThumbprint,[switch]$ForceReinstallIdSAndApi,
 [switch]$UninstallExisting,[string]$ExtraMsiArgs='',
 [switch]$UninstallExistingAppServer,[string]$AppServerExtraMsiArgs,
 [int]$AppServerTimeoutMinutes=15,
 [switch]${SkipInstall-SqlServer},[switch]${SkipCreate-Odbc},
 [switch]$SkipInitializeOnBaseDatabase,[switch]${SkipInstall-HylandIdSAndApi},
 [switch]${SkipInstall-OnBaseApplicationServer},[switch]${SkipInstall-UnityClient},
 [switch]${SkipVerify-Installation}
)
$ErrorActionPreference='Stop'
# Do not impose StrictMode on legacy children.
$boundInputs=@{};foreach($key in $PSBoundParameters.Keys){$boundInputs[$key]=$PSBoundParameters[$key]}
$scriptsDir=$PSScriptRoot;$settings=$null;$pipelineExitCode=0;$script:TotalReboots=0
function Get-Setting {
 param([string[]]$Paths)
 foreach($path in $Paths){
  $current=$settings
  foreach($part in ($path -split '\.')){
   if($null -eq $current){break}
   $property=$current.PSObject.Properties[$part]
   if($null -eq $property){$current=$null;break}
   $current=$property.Value
  }
  if($null -ne $current -and -not [string]::IsNullOrWhiteSpace([string]$current)){return $current}
 }
 return $null
}
function Resolve-Input {
 param([string]$Name,[string[]]$Paths,$Default)
 if($boundInputs.ContainsKey($Name)){return $boundInputs[$Name]}
 $value=Get-Setting $Paths
 if($null -ne $value){return $value}
 return $Default
}
function Get-ChildCommand {
 param([string]$FileName)
 return (Get-Command -Name (Join-Path $scriptsDir $FileName) -CommandType ExternalScript -ErrorAction Stop)
}
function Find-ChildParameter {
 param($Command,[string[]]$Names)
 foreach($name in $Names){foreach($parameter in $Command.Parameters.Values){
  if($parameter.Name -ieq $name -or $parameter.Aliases -icontains $name){return $parameter.Name}
 }}
 return $null
}
function Add-ChildArgument {
 param($Command,[hashtable]$Arguments,[string[]]$Names,$Value,[switch]$Required)
 if($null -eq $Value){return}
 $name=Find-ChildParameter $Command $Names
 if($name){$Arguments[$name]=$Value}
 elseif($Required){throw "$($Command.Name) does not declare input: $($Names -join ' / ')."}
}
function New-ChildArguments {
 param($Command)
 $arguments=@{}
 Add-ChildArgument $Command $arguments @('SettingsFilePath','SettingsPath') $SettingsFilePath
 Add-ChildArgument $Command $arguments @('DeploymentMode') $DeploymentMode
 if($TargetMachine){Add-ChildArgument $Command $arguments @('TargetMachine') $TargetMachine}
 if($DeploymentMode -eq 'Remote'){
  Add-ChildArgument $Command $arguments @('Credential') $Credential
  Add-ChildArgument $Command $arguments @('ConfigurationName') $ConfigurationName
 }
 if($ShareCredential){Add-ChildArgument $Command $arguments @('ShareCredential') $ShareCredential}
 return $arguments
}
function Stop-InstallerPipeline {
 param([string]$Message,[int]$Code=1)
 $exception=New-Object System.Exception -ArgumentList $Message
 $exception.Data['InstallerExitCode']=$Code;throw $exception
}
function Write-SkippedStep {
 param([int]$Step,[string]$Title)
 Write-Host "`n[STEP $Step/7] $Title - explicitly skipped." -ForegroundColor Yellow
}
function New-EndpointConnectionArguments {
 param([string]$EndpointName)
 return @{
  ComputerName=$TargetMachine;Credential=$Credential;Authentication='Negotiate'
  ConfigurationName=$EndpointName;ErrorAction='Stop'
  SessionOption=(New-PSSessionOption -OpenTimeout 10000 -OperationTimeout 15000)
 }
}
function Invoke-EndpointIdentityProbe {
 $connection=New-EndpointConnectionArguments $ConfigurationName
 Invoke-Command @connection -ScriptBlock {
  $identity=[Security.Principal.WindowsIdentity]::GetCurrent()
  [pscustomobject]@{ComputerName=$env:COMPUTERNAME;Identity=$identity.Name;UserSid=$identity.User.Value;PSVersion=$PSVersionTable.PSVersion.ToString()}
 }
}
function Invoke-BoundedEndpointCommand {
 param([scriptblock]$ScriptBlock,[object[]]$ArgumentList=@(),[int]$TimeoutSeconds=30)
 $connection=New-EndpointConnectionArguments $ConfigurationName
 $connection.SessionOption=New-PSSessionOption -OpenTimeout 5000 -OperationTimeout 10000
 $job=$null
 try{
  $job=Invoke-Command @connection -ScriptBlock $ScriptBlock -ArgumentList $ArgumentList -AsJob
  $finished=Wait-Job -Job $job -Timeout $TimeoutSeconds
  if($null -eq $finished){throw "Endpoint probe exceeded $TimeoutSeconds seconds."}
  if($job.State -ne 'Completed'){
   $reason=($job.ChildJobs|ForEach-Object {$_.JobStateInfo.Reason.Message}) -join '; '
   throw "Endpoint command did not complete: $reason"
  }
  Receive-Job -Job $job -ErrorAction Stop
 }finally{
  if($job){
   if($job.State -in 'Running','NotStarted','Blocked'){Stop-Job -Job $job -ErrorAction SilentlyContinue}
   Remove-Job -Job $job -Force -ErrorAction SilentlyContinue
  }
 }
}
function Get-RemoteRecoverySnapshot {
 param([int]$TimeoutSeconds=30)
 $names=@('WinRM','W3SVC','WAS')
 if($SqlInstance){
  $sqlHost=($SqlInstance -split '\\')[0]
  if(($sqlHost -split '\.')[0] -ieq ($TargetMachine -split '\.')[0]){
   if($SqlInstance.Contains('\')){$names+=('MSSQL$'+($SqlInstance -split '\\')[-1])}else{$names+='MSSQLSERVER'}
  }
 }
 Invoke-BoundedEndpointCommand -TimeoutSeconds $TimeoutSeconds -ArgumentList (, $names) -ScriptBlock {
  param([string[]]$Names)
  $identity=[Security.Principal.WindowsIdentity]::GetCurrent()
  $os=Get-CimInstance Win32_OperatingSystem -ErrorAction Stop
  $bootTicks=$os.LastBootUpTime.ToUniversalTime().Ticks
  $services=@(foreach($name in $Names){
   $service=Get-Service -Name $name -ErrorAction SilentlyContinue
   if($service){[pscustomobject]@{Name=$service.Name;Status=[string]$service.Status}}
  })
  [pscustomobject]@{
   ComputerName=$env:COMPUTERNAME;Identity=$identity.Name;UserSid=$identity.User.Value
   BootTicks=[string]$bootTicks;PSVersion=$PSVersionTable.PSVersion.ToString();Services=$services
  }
 }
}
function Restart-RemoteTargetAndWait {
 param([string]$Reason)
 if($DeploymentMode -ne 'Remote' -or -not $AutoReboot){Stop-InstallerPipeline 'Automatic remote reboot disabled/unavailable. Reboot and resume explicitly.' 3010}
 if($script:TotalReboots -ge $MaxRebootsPerRun){Stop-InstallerPipeline "Maximum $MaxRebootsPerRun reboots per run reached; no additional restart." 3010}
 $before=Get-RemoteRecoverySnapshot
 $expectedShort=([string]$TargetMachine -split '\.')[0]
 if($before.ComputerName -ine $expectedShort){throw "Expected target $expectedShort, got $($before.ComputerName). No reboot issued."}
 if($before.ComputerName -ieq $env:COMPUTERNAME){throw 'Coordinator must run on a different workstation from reboot target.'}
 if(-not $before.BootTicks -or -not $before.UserSid){throw 'Cannot establish boot/endpoint identity baseline; no reboot issued.'}
 $required=@($before.Services|Where-Object {$_.Status -eq 'Running'}|Select-Object -ExpandProperty Name)
 if($required -notcontains 'WinRM'){$required+='WinRM'}
 $script:TotalReboots++
 Write-Warning "Restarting $TargetMachine in 10 seconds: $Reason. SQL/IIS/other target workloads will be interrupted; unsaved target work may be lost."
 $ack=$null
 try{
  $ack=Invoke-BoundedEndpointCommand -ScriptBlock {
   $exe=Join-Path $env:windir 'System32\shutdown.exe'
   $output=& $exe /r /t 10 /f /d p:4:1 /c 'OnBase provisioning: authorized installer-required restart' 2>&1
   $code=$LASTEXITCODE
   [pscustomobject]@{Accepted=($code -eq 0);Code=$code;Message=($output -join ' ')}
  }
 }catch{
  # A request may succeed just before remoting disconnects. Never resend blindly.
  Write-Warning "Restart acknowledgement unavailable. Waiting for a new boot; request will NOT repeat. $($_.Exception.Message)"
 }
 if($null -ne $ack -and -not $ack.Accepted){throw "Target rejected restart (exit $($ack.Code)): $($ack.Message)"}
 $watch=[Diagnostics.Stopwatch]::StartNew();$consecutive=0;$last='Waiting for changed boot time.'
 while($watch.Elapsed.TotalSeconds -lt $RebootTimeoutSeconds){
  $remaining=[int][math]::Ceiling($RebootTimeoutSeconds-$watch.Elapsed.TotalSeconds)
  if($remaining -lt 1){break}
  Start-Sleep -Seconds ([math]::Min($RebootPollIntervalSeconds,$remaining))
  $remaining=[int][math]::Ceiling($RebootTimeoutSeconds-$watch.Elapsed.TotalSeconds)
  if($remaining -lt 1){break}
  $candidate=$null
  try{$candidate=Get-RemoteRecoverySnapshot -TimeoutSeconds ([math]::Min(30,$remaining))}catch{$last=$_.Exception.Message;$consecutive=0}
  if($candidate){
   if($candidate.ComputerName -ine $before.ComputerName -or $candidate.UserSid -ne $before.UserSid){throw 'Computer/endpoint identity changed after restart. Stopped; no endpoint reconfiguration attempted.'}
   if([string]$candidate.BootTicks -eq [string]$before.BootTicks){$last='Endpoint reachable, but new OS boot not confirmed.';$consecutive=0}
   else{
    $notReady=@(foreach($name in $required){
     $service=@($candidate.Services|Where-Object {$_.Name -ieq $name})
     if($service.Count -ne 1 -or $service[0].Status -ne 'Running'){$name}
    })
    if($notReady.Count -gt 0){$last='Waiting for services: '+($notReady -join ', ');$consecutive=0}
    else{
     $consecutive++;$last="New boot and endpoint/services ready ($consecutive/2 checks)."
     if($consecutive -ge 2){
      Write-Host "[REBOOT] $TargetMachine recovered; new boot, same '$ConfigurationName' identity, watched services Running. Continuing." -ForegroundColor Green
      return
     }
    }
   }
  }
  Write-Host ('[REBOOT] {0:N0}s elapsed; next attempt in {1}s. {2}' -f $watch.Elapsed.TotalSeconds,$RebootPollIntervalSeconds,$last) -ForegroundColor Yellow
 }
 Stop-InstallerPipeline "Recovery timed out (budget ${RebootTimeoutSeconds}s). No next installer started. Target may still be rebooting. Last check: $last" 3010
}
function Invoke-Installer {
 param([int]$Step,[string]$Title,$Command,[hashtable]$Parameters)
 Write-Host "`n========================================================" -ForegroundColor Cyan
 Write-Host "[STEP $Step/7] $Title" -ForegroundColor Yellow
 Write-Host "Calling $($Command.Name)" -ForegroundColor Cyan
 Write-Host '========================================================' -ForegroundColor Cyan
 $arguments=$Parameters.Clone();$stageReboots=0
 $isApp=($Command.Name -ieq 'Install-OnBaseApplicationServer.ps1');$operation=[guid]::NewGuid()
 if($isApp){
  Add-ChildArgument $Command $arguments @('OperationId') $operation -Required
  $resumeName=Find-ChildParameter $Command @('ResumeAfterReboot')
  if(-not $resumeName){throw 'Replace AppServer child with the companion reboot/resume revision first.'}
  if($AppServerResumeOperationId -ne [guid]::Empty){
   $operation=$AppServerResumeOperationId
   Add-ChildArgument $Command $arguments @('OperationId') $operation -Required
   $arguments[$resumeName]=$true
  }
 }
 while($true){
  $global:LASTEXITCODE=0;$LASTEXITCODE=0
  $output=@(& $Command.Path @arguments)
  $code=[int]$LASTEXITCODE
  $envelope=$output|Where-Object {$_ -and $_.PSObject.Properties['ExitCode']}|Select-Object -Last 1
  foreach($item in $output){if(-not ($isApp -and $item -and $item.PSObject.Properties['ResumeSupported'])){Write-Output $item}}
  if($isApp){
   if(-not $envelope -or -not $envelope.ResumeSupported){throw 'AppServer required result missing. No automatic retry.'}
   # This child normalizes failures to process exit 1; preserve raw MSI exit
   # in MsiExitCode/ExitCode. Success/reboot categories MUST agree exactly.
   $reported=[int]$envelope.ExitCode
   $outcomeMatches=(($code -eq 0 -and $reported -eq 0) -or
                    ($code -eq 3010 -and $reported -eq 3010) -or
                    ($code -notin 0,3010 -and $reported -notin 0,3010))
   if(-not $outcomeMatches){throw 'AppServer process/result outcomes disagree. No automatic retry.'}
   if(([guid]$envelope.OperationId) -ne $operation){throw 'AppServer operation ID mismatch. No automatic retry.'}
  }
  if($code -eq 0){Write-Host "[STEP $Step/7] Completed." -ForegroundColor Green;return}
  if($code -ne 3010){Stop-InstallerPipeline "$($Command.Name) failed with process exit $code. Pipeline stopped; inspect child log/result." $code}
  if($stageReboots -ge $MaxRebootsPerStage){Stop-InstallerPipeline "Stage $Step reached $MaxRebootsPerStage reboot limit; inspect checkpoint. No further restart." 3010}
  if($isApp -and -not $envelope.ResumeAvailable){Stop-InstallerPipeline 'AppServer supplied no safe reboot checkpoint. No automatic restart/retry.' 3010}
  $stageReboots++
  Restart-RemoteTargetAndWait "$($Command.Name) reported reboot required (3010)"
  if($isApp){
   $arguments[$resumeName]=$true
   Write-Host "[STEP $Step/7] Resuming operation $operation after $($envelope.Stage)..." -ForegroundColor Cyan
   continue
  }
  if($envelope -and $envelope.PSObject.Properties['RebootContinuation'] -and $envelope.RebootContinuation -eq 'NextStage'){
   Write-Host "[STEP $Step/7] Child explicitly marked complete; reboot recovered. Next stage." -ForegroundColor Green
   return
  }
  Stop-InstallerPipeline "$($Command.Name): reboot/recovery completed, but child has no explicit completion/resume contract. It was NOT rerun; no later stage started. Update that child for unattended continuation." 3010
 }
}

# Existing endpoint-bootstrap policy: never overwrite/repermission an endpoint.
function Ensure-RemoteInstallerEndpoint {
 Write-Host "`n[REMOTE] Verifying '$ConfigurationName' on $TargetMachine..." -ForegroundColor Cyan
 $probe=$null;$initialFailure=$null
 try{$probe=Invoke-EndpointIdentityProbe}catch{$initialFailure=$_.Exception.Message}
 if($probe){
  Write-Host "[REMOTE] Verified: $($probe.ComputerName) | Identity: $($probe.Identity) | PowerShell: $($probe.PSVersion)" -ForegroundColor Green
  if($BootstrapUnityInstall){Write-Host '[BOOTSTRAP] Reachable endpoint reused. No registration, credential update, or separate WinRM restart.'}
  if($EndpointRunAsCredential){Write-Warning 'EndpointRunAsCredential not applied: existing endpoint unchanged.'}
  return
 }
 if(-not $BootstrapUnityInstall){throw "Cannot verify '$ConfigurationName'; bootstrap disabled. No installers started. $initialFailure"}
 $defaultConnection=New-EndpointConnectionArguments 'Microsoft.PowerShell'
 try{
  $inspection=Invoke-Command @defaultConnection -ArgumentList $ConfigurationName -ScriptBlock {
   param($EndpointName)
   $ErrorActionPreference='Stop'
   $principal=[Security.Principal.WindowsPrincipal]::new([Security.Principal.WindowsIdentity]::GetCurrent())
   if(-not $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)){throw 'Bootstrap requires elevated admin access through Microsoft.PowerShell.'}
   $existing=@(Get-PSSessionConfiguration -ErrorAction Stop|Where-Object {$_.Name -ieq $EndpointName})
   [pscustomobject]@{Exists=($existing.Count -gt 0)}
  }
 }catch{throw "Cannot inspect bootstrap target; no installers started. $($_.Exception.Message)"}
 if($inspection.Exists){throw "Endpoint exists but is inaccessible. It was not overwritten/restarted. $initialFailure"}
 if($ConfigurationName -notmatch '^[A-Za-z0-9][A-Za-z0-9_.-]{0,63}$'){throw 'Invalid bootstrap endpoint name.'}
 $runAs=$EndpointRunAsCredential
 if(-not $runAs){$runAs=Get-Credential -UserName $Credential.UserName -Message "Approved FIXED RunAs account for '$ConfigurationName' on $TargetMachine; EXISTING password. Account name may be changed."}
 if(-not $runAs){throw 'Bootstrap credential prompt cancelled.'}
 Write-Warning "Creating persistent fixed-RunAs endpoint '$ConfigurationName' as '$($runAs.UserName)'; Administrators/SYSTEM access only. Target WinRM will restart."
 $operationId=[guid]::NewGuid().ToString('N')
 try{
  $bootstrap=Invoke-Command @defaultConnection -ArgumentList $ConfigurationName,$runAs,$operationId -ScriptBlock {
   param($EndpointName,[pscredential]$RunAsCredential,[string]$OperationId)
   $ErrorActionPreference='Stop'
   if($PSVersionTable.PSVersion.Major -ne 5 -or $PSVersionTable.PSVersion.Minor -lt 1){throw 'Bootstrap requires Windows PowerShell 5.1.'}
   if(@(Get-PSSessionConfiguration|Where-Object {$_.Name -ieq $EndpointName}).Count -gt 0){throw 'Endpoint appeared during bootstrap; unchanged.'}
   if($OperationId -notmatch '^[a-f0-9]{32}$'){throw 'Invalid operation ID.'}
   foreach($name in @('New-ScheduledTaskAction','New-ScheduledTaskPrincipal','New-ScheduledTaskSettingsSet','Register-ScheduledTask','Start-ScheduledTask','Unregister-ScheduledTask')){$null=Get-Command $name -ErrorAction Stop}
   $account=[Security.Principal.NTAccount]::new($RunAsCredential.UserName)
   $expectedSid=$account.Translate([Security.Principal.SecurityIdentifier]).Value
   $taskName="OnBase-WinRM-Bootstrap-$OperationId"
   $statusDirectory=Join-Path $env:ProgramData "OnBaseProvisioning-Bootstrap-$OperationId"
   $statusFile=Join-Path $statusDirectory 'status.json';$registered=$false;$submitted=$false
   try{
    New-Item -ItemType Directory -Path $statusDirectory -ErrorAction Stop|Out-Null
    $acl=[Security.AccessControl.DirectorySecurity]::new();$acl.SetAccessRuleProtection($true,$false)
    $admin=[Security.Principal.SecurityIdentifier]::new('S-1-5-32-544');$acl.SetOwner($admin)
    foreach($sidText in @('S-1-5-32-544','S-1-5-18')){
     $sid=[Security.Principal.SecurityIdentifier]::new($sidText)
     $rule=[Security.AccessControl.FileSystemAccessRule]::new($sid,'FullControl','ContainerInherit, ObjectInherit','None','Allow');[void]$acl.AddAccessRule($rule)
    }
    Set-Acl -LiteralPath $statusDirectory -AclObject $acl
    @{State='Pending';OperationId=$OperationId}|ConvertTo-Json|Set-Content -LiteralPath $statusFile -Encoding UTF8
    $restartScript=@'
$ErrorActionPreference='Stop'
$statusFile='__STATUS_FILE__'
$taskName='__TASK_NAME__'
try {
 Start-Sleep -Seconds 10
 Restart-Service -Name WinRM -Force -ErrorAction Stop
 (Get-Service WinRM).WaitForStatus([ServiceProcess.ServiceControllerStatus]::Running,[TimeSpan]::FromSeconds(45))
 $status=@{State='Restarted';CompletedUtc=[DateTime]::UtcNow.ToString('o')}
 [IO.File]::WriteAllText($statusFile,($status|ConvertTo-Json -Compress),[Text.UTF8Encoding]::new($false))
} catch {
 try {
  $status=@{State='Failed';Message=$_.Exception.Message;CompletedUtc=[DateTime]::UtcNow.ToString('o')}
  [IO.File]::WriteAllText($statusFile,($status|ConvertTo-Json -Compress),[Text.UTF8Encoding]::new($false))
 } catch {}
 exit 1
} finally {
 Unregister-ScheduledTask -TaskName $taskName -Confirm:$false -ErrorAction SilentlyContinue
}
'@
    $restartScript=$restartScript.Replace('__STATUS_FILE__',$statusFile.Replace("'","''")).Replace('__TASK_NAME__',$taskName.Replace("'","''"))
    $encoded=[Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($restartScript))
    $exe=Join-Path $env:windir 'System32\WindowsPowerShell\v1.0\powershell.exe'
    if(-not (Test-Path -LiteralPath $exe)){throw 'Windows PowerShell executable not found.'}
    $action=New-ScheduledTaskAction -Execute $exe -Argument "-NoLogo -NoProfile -NonInteractive -EncodedCommand $encoded"
    $principal=New-ScheduledTaskPrincipal -UserId SYSTEM -LogonType ServiceAccount -RunLevel Highest
    $taskSettings=New-ScheduledTaskSettingsSet -ExecutionTimeLimit (New-TimeSpan -Minutes 3) -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries
    Register-ScheduledTask -TaskName $taskName -Action $action -Principal $principal -Settings $taskSettings -Description 'One-time OnBase endpoint bootstrap; no password in task.'|Out-Null
    Register-PSSessionConfiguration -Name $EndpointName -RunAsCredential $RunAsCredential -SecurityDescriptorSddl 'O:BAG:BAD:P(A;;GA;;;SY)(A;;GA;;;BA)' -NoServiceRestart -Force|Out-Null
    $registered=$true;Start-ScheduledTask -TaskName $taskName;$submitted=$true
    [pscustomobject]@{TaskName=$taskName;StatusDirectory=$statusDirectory;StatusFile=$statusFile;ExpectedRunAsSid=$expectedSid;ExpectedRunAsUser=$RunAsCredential.UserName}
   }catch{
    if(-not $submitted){Unregister-ScheduledTask -TaskName $taskName -Confirm:$false -ErrorAction SilentlyContinue}
    throw "Bootstrap incomplete; endpoint registration attempted: $registered. Task: $taskName. Status: $statusDirectory. $($_.Exception.Message)"
   }
  }
 }catch{throw "Bootstrap acknowledgement incomplete; state may have changed. No installers started. Inspect operation $operationId before retry. $($_.Exception.Message)"}
 finally{$runAs=$null}
 $watch=[Diagnostics.Stopwatch]::StartNew();$confirmed=$false;$verified=$null;$last='Waiting for restart task.'
 while($watch.Elapsed.TotalSeconds -lt $EndpointBootstrapTimeoutSeconds){
  if(-not $confirmed){
   $status=$null
   try{
    $status=Invoke-Command @defaultConnection -ArgumentList $bootstrap.StatusFile -ScriptBlock {
     param($StatusFile)
     if(Test-Path -LiteralPath $StatusFile){Get-Content -LiteralPath $StatusFile -Raw|ConvertFrom-Json}else{[pscustomobject]@{State='Missing'}}
    }
   }catch{$last=$_.Exception.Message}
   if($status){
    if($status.State -eq 'Failed'){throw "Bootstrap WinRM restart failed: $($status.Message). Status: $($bootstrap.StatusFile)"}
    if($status.State -eq 'Restarted'){$confirmed=$true}else{$last="Restart status: $($status.State)"}
   }
  }
  if($confirmed){
   $candidate=$null;try{$candidate=Invoke-EndpointIdentityProbe}catch{$last=$_.Exception.Message}
   if($candidate){
    if($candidate.UserSid -ne $bootstrap.ExpectedRunAsSid){throw 'Bootstrap identity mismatch. No installers started.'}
    $verified=$candidate;break
   }
  };Start-Sleep -Seconds 3
 }
 if(-not $verified){throw "Bootstrap verification timed out. Inspect $($bootstrap.TaskName), $($bootstrap.StatusFile). Last: $last"}
 Write-Host "[REMOTE] Verified after bootstrap: $($verified.ComputerName) | $($verified.Identity)" -ForegroundColor Green
 try{
  Invoke-Command @defaultConnection -ArgumentList $bootstrap.StatusDirectory -ScriptBlock {
   param($StatusDirectory)
   $expected=[IO.Path]::GetFullPath($env:ProgramData).TrimEnd('\')
   $actual=[IO.Path]::GetFullPath((Split-Path $StatusDirectory -Parent)).TrimEnd('\')
   if($actual -ine $expected -or (Split-Path $StatusDirectory -Leaf) -notmatch '^OnBaseProvisioning-Bootstrap-[a-f0-9]{32}$'){throw 'Unexpected cleanup path.'}
   Remove-Item -LiteralPath (Join-Path $StatusDirectory 'status.json') -Force
   Remove-Item -LiteralPath $StatusDirectory -Force
  }
 }catch{Write-Warning "Endpoint verified; status cleanup failed: $($_.Exception.Message)"}
}
try{
 if([string]::IsNullOrWhiteSpace($SettingsFilePath)){$SettingsFilePath=Join-Path (Split-Path $scriptsDir -Parent) 'Settings.json'}
 $SettingsFilePath=(Resolve-Path -LiteralPath $SettingsFilePath -ErrorAction Stop).ProviderPath
 $settings=Get-Content -LiteralPath $SettingsFilePath -Raw|ConvertFrom-Json -ErrorAction Stop
 $DeploymentMode=Resolve-Input 'DeploymentMode' @('deploymentMode') $DeploymentMode
 $TargetMachine=Resolve-Input 'TargetMachine' @('machineName') $TargetMachine
 $ConfigurationName=Resolve-Input 'ConfigurationName' @('configurationName') $ConfigurationName
 $BuildPath=Resolve-Input 'BuildPath' @('buildPath') $BuildPath
 $SqlSetupPath=Resolve-Input 'SqlSetupPath' @('sqlSetupPath') $SqlSetupPath
 $OdbcSetupPath=Resolve-Input 'OdbcSetupPath' @('odbcSetupPath','odbcDriverMsiPath') $OdbcSetupPath
 $SqlVersion=Resolve-Input 'SqlVersion' @('sqlVersion') $SqlVersion
 $SqlInstance=Resolve-Input 'SqlInstance' @('sqlInstance') $SqlInstance
 $OdbcVersion=Resolve-Input 'OdbcVersion' @('odbcVersion') $OdbcVersion
 $OnSqlVersionMismatch=Resolve-Input 'OnSqlVersionMismatch' @('onSqlVersionMismatch') $OnSqlVersionMismatch
 $DatabaseName=Resolve-Input 'DatabaseName' @('databaseName') $DatabaseName
 $DsnName=Resolve-Input 'DsnName' @('datasourceName') $DatabaseName
 $IdSAppName=Resolve-Input 'IdSAppName' @('idp.appName') $IdSAppName
 $ApiAppName=Resolve-Input 'ApiAppName' @('apiServer.appName') $ApiAppName
 $CertificateThumbprint=Resolve-Input 'CertificateThumbprint' @('thumbprint') $CertificateThumbprint
 $DbUtilsSourcePath=Resolve-Input 'DbUtilsSourcePath' @('dbUtilsSourcePath') $DbUtilsSourcePath
 $UnityMsiPath=Resolve-Input 'UnityMsiPath' @('unityMsiPath') $UnityMsiPath
 $AppServerMsiPath=Resolve-Input 'AppServerMsiPath' @('appServer.sourceMsiPath','Installers.AppServerMsiPath') $AppServerMsiPath
 $AppServerExtraMsiArgs=Resolve-Input 'AppServerExtraMsiArgs' @('appServer.extraMsiArgs') $AppServerExtraMsiArgs
 $AppServerTimeoutMinutes=Resolve-Input 'AppServerTimeoutMinutes' @('appServer.timeoutMinutes') $AppServerTimeoutMinutes
 $autoSetting=Resolve-Input 'AutoReboot' @('reboot.autoRestart') ([bool]$AutoReboot)
 if($autoSetting -is [System.Management.Automation.SwitchParameter]){$autoSetting=[bool]$autoSetting}
 if($autoSetting -isnot [bool]){throw 'reboot.autoRestart must be JSON true/false, not a string.'}
 $AutoReboot=$autoSetting
 $RebootPollIntervalSeconds=Resolve-Input 'RebootPollIntervalSeconds' @('reboot.pollIntervalSeconds') $RebootPollIntervalSeconds
 $RebootTimeoutSeconds=Resolve-Input 'RebootTimeoutSeconds' @('reboot.timeoutSeconds') $RebootTimeoutSeconds
 $MaxRebootsPerStage=Resolve-Input 'MaxRebootsPerStage' @('reboot.maxPerStage') $MaxRebootsPerStage
 $MaxRebootsPerRun=Resolve-Input 'MaxRebootsPerRun' @('reboot.maxPerRun') $MaxRebootsPerRun
 if($RebootPollIntervalSeconds -lt 5 -or $RebootPollIntervalSeconds -gt 300 -or $RebootTimeoutSeconds -lt 60 -or $RebootTimeoutSeconds -gt 7200 -or $MaxRebootsPerStage -lt 1 -or $MaxRebootsPerStage -gt 10 -or $MaxRebootsPerRun -lt 1 -or $MaxRebootsPerRun -gt 30){throw 'Reboot settings outside documented ranges.'}
 if($DeploymentMode -notin 'Local','Remote'){throw 'Invalid deploymentMode.'}
 if($OnSqlVersionMismatch -notin 'SideBySide','Skip','Fail'){throw 'Use onSqlVersionMismatch SideBySide, Skip or Fail.'}
 if($OdbcVersion -notin '17','18'){throw 'Unsupported odbcVersion.'}
 if(-not $DbUtilsSourcePath -and $BuildPath -and -not $boundInputs.ContainsKey('DbUtilsSourcePath')){$DbUtilsSourcePath=$BuildPath.TrimEnd('\')+'\utilities\FoundationDBUtils'}
 if(-not $UnityMsiPath -and $BuildPath -and -not $boundInputs.ContainsKey('UnityMsiPath')){$UnityMsiPath=$BuildPath.TrimEnd('\')+'\Install\Unity Client\Hyland Unity Client.msi'}
 $majorVersion=switch($SqlVersion){'2019'{15};'2022'{16};default{throw "Unsupported sqlVersion: $SqlVersion"}}
 $OdbcDriverName=Resolve-Input 'OdbcDriverName' @('odbcDriverName') "ODBC Driver $OdbcVersion for SQL Server"
 if($DeploymentMode -eq 'Remote' -and (-not $TargetMachine -or -not $ConfigurationName)){throw 'Remote mode requires target and named endpoint.'}
 if($DeploymentMode -eq 'Local'){
  if($boundInputs.ContainsKey('BootstrapUnityInstall') -and $BootstrapUnityInstall){throw 'Endpoint bootstrap is Remote-only.'}
  if($RestartBeforeInstall){throw 'RestartBeforeInstall is Remote-only; coordinator cannot survive its own host restart.'}
  $BootstrapUnityInstall=$false
 }
 if(-not $SqlInstance){
  $server=if($DeploymentMode -eq 'Remote'){($TargetMachine -split '\.')[0]}else{$env:COMPUTERNAME}
  $instance=if($SqlVersion -eq '2022'){'SQL2K22'}else{'SQL2K19'};$SqlInstance="$server\$instance"
 }
 $instanceShort=if($SqlInstance.Contains('\')){($SqlInstance -split '\\')[-1]}else{'MSSQLSERVER'}
 $anyEnabled=-not (${SkipInstall-SqlServer} -and ${SkipCreate-Odbc} -and $SkipInitializeOnBaseDatabase -and ${SkipInstall-HylandIdSAndApi} -and ${SkipInstall-OnBaseApplicationServer} -and ${SkipInstall-UnityClient} -and ${SkipVerify-Installation})
 if($anyEnabled -and $DeploymentMode -eq 'Remote' -and -not $Credential){
  $prompt=@{Message="Remote provisioning credentials for $TargetMachine"};$userName=Get-Setting @('remoteUsername')
  if($userName){$prompt.UserName=[string]$userName};$Credential=Get-Credential @prompt
  if(-not $Credential){throw 'Credential prompt cancelled.'}
 }
 Write-Host "`nOnBase provisioning - seven stages, reboot-aware" -ForegroundColor Green
 Write-Host "Settings: $SettingsFilePath | Mode: $DeploymentMode | Target: $TargetMachine | SQL: $SqlInstance"
 Write-Host "Reboot policy: auto=$([bool]$AutoReboot), interval=${RebootPollIntervalSeconds}s, recovery budget=${RebootTimeoutSeconds}s"
 if($anyEnabled -and $DeploymentMode -eq 'Remote'){
  Write-Host "[REMOTE] Endpoint creation / separate WinRM restart allowed: $([bool]$BootstrapUnityInstall)"
  Ensure-RemoteInstallerEndpoint
  if($AutoReboot -and (($TargetMachine -split '\.')[0] -ieq $env:COMPUTERNAME)){throw 'Run remote reboot coordination from a separate workstation.'}
  if($RestartBeforeInstall){Restart-RemoteTargetAndWait 'Explicit RestartBeforeInstall: satisfy prior run reboot requirement'}
 }elseif(-not $anyEnabled){Write-Host 'All stages skipped; no connections/restarts attempted.' -ForegroundColor Yellow}

 if(${SkipInstall-SqlServer}){Write-SkippedStep 1 'SQL Server installation / validation'}else{
  $child=Get-ChildCommand 'Install-SqlServer.ps1';$p=New-ChildArguments $child
  Add-ChildArgument $child $p @('SqlSetupPath') ([string]$SqlSetupPath)
  if(Find-ChildParameter $child @('InstanceName')){Add-ChildArgument $child $p @('InstanceName') $instanceShort}else{Add-ChildArgument $child $p @('SqlInstance') $SqlInstance}
  if(Find-ChildParameter $child @('TargetMajorVersion')){Add-ChildArgument $child $p @('TargetMajorVersion') $majorVersion}else{Add-ChildArgument $child $p @('SqlVersion') $SqlVersion}
  if(Find-ChildParameter $child @('OnSqlVersionMismatch')){Add-ChildArgument $child $p @('OnSqlVersionMismatch') $OnSqlVersionMismatch}else{
   $legacyPolicy=switch($OnSqlVersionMismatch){'SideBySide'{'SideBySide'};'Fail'{'Stop'};default{throw "Legacy SQL child cannot express '$OnSqlVersionMismatch'."}}
   Add-ChildArgument $child $p @('OnVersionMismatch') $legacyPolicy -Required
  }
  if($SaCredential){Add-ChildArgument $child $p @('SaCredential') $SaCredential -Required}
  Add-ChildArgument $child $p @('OdbcDriverName') $OdbcDriverName
  if($OdbcSetupPath){Add-ChildArgument $child $p @('OdbcDriverMsiPath') $OdbcSetupPath}
  Invoke-Installer 1 'SQL Server installation / validation' $child $p
 }
 if(${SkipCreate-Odbc}){Write-SkippedStep 2 'ODBC driver installation / validation'}else{
  $child=Get-ChildCommand 'Create-ODBC.ps1';$p=New-ChildArguments $child
  Add-ChildArgument $child $p @('OdbcVersion') $OdbcVersion
  Add-ChildArgument $child $p @('OdbcDriverName','OdbcDriver') $OdbcDriverName
  if($OdbcSetupPath){Add-ChildArgument $child $p @('OdbcSetupPath','OdbcDriverMsiPath','SourceMsiPath') $OdbcSetupPath}
  Invoke-Installer 2 'ODBC driver installation / validation' $child $p
 }
 if(-not $HsiCredential -and (-not $SkipInitializeOnBaseDatabase -or -not ${SkipInstall-HylandIdSAndApi} -or -not ${SkipVerify-Installation})){
  $hsiUser=Get-Setting @('hsiUsername');if(-not $hsiUser){$hsiUser='hsi'}
  $HsiCredential=Get-Credential -UserName $hsiUser -Message 'OnBase hsi SQL password (existing password if login exists)'
  if(-not $HsiCredential){throw 'HSI credential prompt cancelled.'}
 }
 if($SkipInitializeOnBaseDatabase){Write-SkippedStep 3 'OnBase database initialization'}else{
  $child=Get-ChildCommand 'Initialize-OnBaseDatabase.ps1';$p=New-ChildArguments $child
  Add-ChildArgument $child $p @('SqlInstance') $SqlInstance;Add-ChildArgument $child $p @('DatabaseName') $DatabaseName
  Add-ChildArgument $child $p @('DsnName') $DsnName;Add-ChildArgument $child $p @('OdbcDriver','OdbcDriverName') $OdbcDriverName
  if($DbUtilsSourcePath){Add-ChildArgument $child $p @('DbUtilsSourcePath') $DbUtilsSourcePath}
  if($HsiCredential){Add-ChildArgument $child $p @('HsiCredential') $HsiCredential}
  if($SqlAdminCredential){Add-ChildArgument $child $p @('SqlAdminCredential') $SqlAdminCredential -Required}
  Invoke-Installer 3 'OnBase database initialization' $child $p
 }
 if(${SkipInstall-HylandIdSAndApi}){Write-SkippedStep 4 'Hyland Identity Service and API Server installation'}else{
  $child=Get-ChildCommand 'Install-HylandIdSAndApi.ps1';$p=New-ChildArguments $child
  Add-ChildArgument $child $p @('SqlInstance') $SqlInstance;Add-ChildArgument $child $p @('DatabaseName') $DatabaseName
  Add-ChildArgument $child $p @('IdSAppName','IdSApp') $IdSAppName;Add-ChildArgument $child $p @('ApiAppName','ApiApp') $ApiAppName
  if($HsiCredential){Add-ChildArgument $child $p @('HsiCredential') $HsiCredential}
  if($CertificateThumbprint){Add-ChildArgument $child $p @('CertificateThumbprint','CertThumb') $CertificateThumbprint -Required}
  if($ForceReinstallIdSAndApi){Add-ChildArgument $child $p @('ForceReinstall') $true -Required}
  Invoke-Installer 4 'Hyland Identity Service and API Server installation' $child $p
 }
 if(${SkipInstall-OnBaseApplicationServer}){Write-SkippedStep 5 'OnBase Application Server installation'}else{
  $child=Get-ChildCommand 'Install-OnBaseApplicationServer.ps1';$p=New-ChildArguments $child
  if($AppServerMsiPath){Add-ChildArgument $child $p @('SourceMsiPath') $AppServerMsiPath}
  if($AppServerExtraMsiArgs){Add-ChildArgument $child $p @('ExtraMsiArgs') $AppServerExtraMsiArgs}
  Add-ChildArgument $child $p @('TimeoutMinutes') $AppServerTimeoutMinutes
  Add-ChildArgument $child $p @('SqlInstance') $SqlInstance;Add-ChildArgument $child $p @('DatabaseName') $DatabaseName
  if($HsiCredential){Add-ChildArgument $child $p @('HsiCredential') $HsiCredential}
  if($UninstallExistingAppServer){Add-ChildArgument $child $p @('UninstallExisting') $true -Required}
  Invoke-Installer 5 'OnBase Application Server installation' $child $p
 }
 if(${SkipInstall-UnityClient}){Write-SkippedStep 6 'Unity Client installation'}else{
  $child=Get-ChildCommand 'Install-Unity.ps1';$p=New-ChildArguments $child
  if($UnityMsiPath){Add-ChildArgument $child $p @('SourceMsiPath') $UnityMsiPath}
  if($ExtraMsiArgs){Add-ChildArgument $child $p @('ExtraMsiArgs') $ExtraMsiArgs -Required}
  if($UninstallExisting){Add-ChildArgument $child $p @('UninstallExisting') $true -Required}
  Invoke-Installer 6 'Unity Client installation' $child $p
 }
 if(${SkipVerify-Installation}){Write-SkippedStep 7 'End-to-end verification'}else{
  $child=Get-ChildCommand 'Verify-Installation.ps1';$p=New-ChildArguments $child
  Add-ChildArgument $child $p @('SqlInstance') $SqlInstance;Add-ChildArgument $child $p @('DatabaseName') $DatabaseName
  Add-ChildArgument $child $p @('DsnName','DataSource') $DsnName;Add-ChildArgument $child $p @('OdbcDriver','OdbcDriverName') $OdbcDriverName
  if($HsiCredential){Add-ChildArgument $child $p @('HsiCredential') $HsiCredential}
  Invoke-Installer 7 'End-to-end verification' $child $p
 }
 Write-Host "`nAll enabled stages completed without reported failure; handled restarts: $script:TotalReboots." -ForegroundColor Green
 if(${SkipVerify-Installation}){Write-Host 'Final verification explicitly skipped; not independent proof of environment health.' -ForegroundColor Yellow}
}catch{
 if($_.Exception.Data.Contains('InstallerExitCode')){$pipelineExitCode=[int]$_.Exception.Data['InstallerExitCode']}else{$pipelineExitCode=1}
 Write-Host "`nOnBase provisioning stopped: $($_.Exception.Message)" -ForegroundColor Red
}
exit $pipelineExitCode
