function Write-Log {

param([string]$Message)

$timestamp = Get-Date

"$timestamp : $Message" |
Out-File `
"..\Logs\Install.log" `
-Append
}