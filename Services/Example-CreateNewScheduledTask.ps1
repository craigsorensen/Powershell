# Configurable Items
$ScriptPath = "C:\OnBaseMaintenance\test.ps1" # Full path to script you'd like to execute
$UserName = "ad\is-svc-onbase-t" # Service account name in which script should be executed as
$TaskName = "Stop Services" # Task display name
$Trigger = New-ScheduledTaskTrigger -At 10:00am –Daily # Specify when and how oftent the task should be ran

$Action = New-ScheduledTaskAction -Execute "C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe" -Argument "-NonInteractive -NoLogo -NoProfile -ExecutionPolicy Bypass -File ""$($ScriptPath)"""

$Settings = New-ScheduledTaskSettingsSet -DontStopOnIdleEnd -RestartInterval (New-TimeSpan -Minutes 1) -RestartCount 10 -StartWhenAvailable
$Settings.ExecutionTimeLimit = "PT0S"

$Credentials = Get-Credential -UserName "$($UserName)" -Message "Enter password: "
$Password = $Credentials.GetNetworkCredential().Password

$Task = New-ScheduledTask -Action $Action -Trigger $Trigger -Settings $Settings
$Task | Register-ScheduledTask -TaskName "$($TaskName)" -User "$($UserName)" -Password $Password
