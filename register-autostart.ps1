$action = New-ScheduledTaskAction -Execute 'powershell.exe' -Argument '-ExecutionPolicy Bypass -WindowStyle Hidden -File F:/scripts/start-all.ps1'
$trigger = New-ScheduledTaskTrigger -AtLogOn
$principal = New-ScheduledTaskPrincipal -UserId 'zhao' -RunLevel Highest
Register-ScheduledTask -TaskName 'OpenCode-AutoStart' -Action $action -Trigger $trigger -Principal $principal -Force
Write-Host "Scheduled task 'OpenCode-AutoStart' registered successfully"
