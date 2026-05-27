$triggers = New-ScheduledTaskTrigger -At (Get-Date) -Once
$action = New-ScheduledTaskAction -Execute "%windir%\system32\deviceenroller.exe" -Argument "/c /AutoEnrollMDM"
Register-ScheduledTask -TaskName "EntraJoinTask" -Trigger $triggers -Action $action -User "SYSTEM" -Force
Start-ScheduledTask -TaskName "EntraJoinTask"
