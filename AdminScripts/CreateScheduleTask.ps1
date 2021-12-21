#=========================07-28-2021=======================#
$Trigger = New-ScheduledTaskTrigger -DaysOfWeek Sunday -At 10:00am –Weekly # Specify the trigger settings
$User = "citspoRunner" # Specify the account to run the script
$Action = New-ScheduledTaskAction -Execute "PowerShell.exe" -Argument "D:\Scripting\O365DevOps\AdminPortal-Operations\PopulateSPOSiteSCA.ps1" # Specify what program to run and with its parameters
Register-ScheduledTask -TaskName "SPO-PopulateSiteSCA" -Trigger $Trigger -User $User -Action $Action -RunLevel Highest –Force # Specify the name of the task

$Trigger = New-ScheduledTaskTrigger -DaysOfWeek Tuesday -At 10:00am –Weekly # Specify the trigger settings
$User = "citspoRunner" # Specify the account to run the script
$Action = New-ScheduledTaskAction -Execute "PowerShell.exe" -Argument "D:\Scripting\O365DevOps\AdminPortal-Operations\AutoIncreaseQuotaStorage.ps1"
Register-ScheduledTask -TaskName "SPO-AutoIncreaseQuotaStorage" -Trigger $Trigger -User $User -Action $Action -RunLevel Highest –Force

$Trigger = New-ScheduledTaskTrigger -DaysOfWeek Saturday -At 10:00am –Weekly # Specify the trigger settings
$User = "citspoRunner" # Specify the account to run the script
$Action = New-ScheduledTaskAction -Execute "PowerShell.exe" -Argument "-ExecutionPolicy ByPass -File D:\Scripting\O365DevOps\AdminPortal-Operations\PopulatePersonalSiteSCA.ps1 -CoIC $CoIC"
Register-ScheduledTask -TaskName "SPO-PopulatePersonalSiteSCA-Batch-File-1" -Trigger $Trigger -User $User -Action $Action -RunLevel Highest –Force

#==========================================================#

#=========================05-11-2021=======================#
$Trigger = New-ScheduledTaskTrigger -At 5:00am –Daily # Specify the trigger settings
$User = "citspoRunner" # Specify the account to run the script
$Action = New-ScheduledTaskAction -Execute "PowerShell.exe" -Argument "D:\Scripting\O365DevOps\AdminPortal-Operations\PopulatePowerBIWorkspacesCache.ps1" # Specify what program to run and with its parameters
Register-ScheduledTask -TaskName "SPO-PopulatePowerBIWorkspacesCache" -Trigger $Trigger -User $User -Action $Action -RunLevel Highest –Force # Specify the name of the task
#==========================================================#

#=========================02-24-2021=======================#
$Trigger = New-ScheduledTaskTrigger -DaysOfWeek Monday -At 10:00am –Weekly
$User = "citspoRunner" # Specify the account to run the script
$Action = New-ScheduledTaskAction -Execute "PowerShell.exe" -Argument "D:\Scripting\O365DevOps\AdminPortal-Operations\PopulateTeamsOwner.ps1" 
Register-ScheduledTask -TaskName "SPO-ICAdminPortal-PopulateTeamsOwner" -Trigger $Trigger -User $User -Action $Action -RunLevel Highest –Force
#==========================================================#


$Trigger = New-ScheduledTaskTrigger -At 4:50am –Daily # Specify the trigger settings
#$User= "NT AUTHORITY\SYSTEM" # Specify the account to run the script
$User = "citspadminsvc" # Specify the account to run the script
$Action = New-ScheduledTaskAction -Execute "PowerShell.exe" -Argument "D:\Scripting\O365DevOps\AdminPortal-Operations\ProcessM365Operations.ps1" # Specify what program to run and with its parameters
Register-ScheduledTask -TaskName "SPO-M365 Admin Portal Operation" -Trigger $Trigger -User $User -Action $Action -RunLevel Highest –Force # Specify the name of the task

#===================================================================================#
$Trigger = New-ScheduledTaskTrigger -At 6:47am –Daily # Specify the trigger settings
#$User= "NT AUTHORITY\SYSTEM" # Specify the account to run the script
$User = "citspadminsvc" # Specify the account to run the script
$Action = New-ScheduledTaskAction -Execute "PowerShell.exe" -Argument "D:\Scripting\O365\AdminPortal\Master_Script.ps1" # Specify what program to run and with its parameters
Register-ScheduledTask -TaskName "SPO-ICAdminPortal-SiteProvision" -Trigger $Trigger -User $User -Action $Action -RunLevel Highest –Force # Specify the name of the task

#===================================================================================#
$Trigger = New-ScheduledTaskTrigger -At 7:29am –Daily # Specify the trigger settings
#$User= "NT AUTHORITY\SYSTEM" # Specify the account to run the script
$User = "citspadminsvc" # Specify the account to run the script
$Action = New-ScheduledTaskAction -Execute "PowerShell.exe" -Argument "D:\Scripting\O365DevOps\AdminPortal-Operations\DisableExternalSharing.ps1" # Specify what program to run and with its parameters
Register-ScheduledTask -TaskName "SPO-DisableExternalAccess" -Trigger $Trigger -User $User -Action $Action -RunLevel Highest –Force # Specify the name of the task

#===================================================================================#
$Trigger = New-ScheduledTaskTrigger -At 6:17am –Daily # Specify the trigger settings
#$User= "NT AUTHORITY\SYSTEM" # Specify the account to run the script
$User = "citspadminsvc" # Specify the account to run the script
$Action = New-ScheduledTaskAction -Execute "PowerShell.exe" -Argument "D:\Scripting\O365DevOps\M365-Operations\ProcessM365Operations.ps1" # Specify what program to run and with its parameters
Register-ScheduledTask -TaskName "SPO-M365OpsSupportTeam" -Trigger $Trigger -User $User -Action $Action -RunLevel Highest –Force # Specify the name of the task

#===================================================================================#
$Trigger = New-ScheduledTaskTrigger -At 5:00pm –Daily # Specify the trigger settings
#$User= "NT AUTHORITY\SYSTEM" # Specify the account to run the script
$User = "citspadminsvc" # Specify the account to run the script
$Action = New-ScheduledTaskAction -Execute "PowerShell.exe" -Argument "D:\Scripting\O365DevOps\AdminPortal-Operations\PopulateDailyGroupsCache.ps1" # Specify what program to run and with its parameters
Register-ScheduledTask -TaskName "SPO-PopulateDailyGroupsCache" -Trigger $Trigger -User $User -Action $Action -RunLevel Highest –Force # Specify the name of the task

#===================================================================================#
$Trigger = New-ScheduledTaskTrigger -At 5:00pm –Daily # Specify the trigger settings
#$User= "NT AUTHORITY\SYSTEM" # Specify the account to run the script
$User = "citspadminsvc" # Specify the account to run the script
$Action = New-ScheduledTaskAction -Execute "PowerShell.exe" -Argument "D:\Scripting\O365DevOps\AdminPortal-Operations\PopulateDailyUsersCache.ps1" # Specify what program to run and with its parameters
Register-ScheduledTask -TaskName "SPO-PopulateDailyUsersCache" -Trigger $Trigger -User $User -Action $Action -RunLevel Highest –Force # Specify the name of the task

#===================================================================================#
$Trigger = New-ScheduledTaskTrigger -At 5:00pm –Daily # Specify the trigger settings
#$User= "NT AUTHORITY\SYSTEM" # Specify the account to run the script
$User = "citspadminsvc" # Specify the account to run the script
$Action = New-ScheduledTaskAction -Execute "PowerShell.exe" -Argument "D:\Scripting\O365DevOps\AdminPortal-Operations\PopulateDailySitesCache.ps1" # Specify what program to run and with its parameters
Register-ScheduledTask -TaskName "SPO-PopulateDailySitesCache" -Trigger $Trigger -User $User -Action $Action -RunLevel Highest –Force # Specify the name of the task

#===================================================================================#
$Trigger = New-ScheduledTaskTrigger -At 4:50am –Daily # Specify the trigger settings
#$User= "NT AUTHORITY\SYSTEM" # Specify the account to run the script
$User = "citspadminsvc" # Specify the account to run the script
$Action = New-ScheduledTaskAction -Execute "PowerShell.exe" -Argument "D:\Scripting\O365DevOps\AdminPortal-Operations\SyncO365Groups.ps1" # Specify what program to run and with its parameters
Register-ScheduledTask -TaskName "SPO-SyncO365Groups" -Trigger $Trigger -User $User -Action $Action -RunLevel Highest –Force # Specify the name of the task

#===================================================================================#
$Trigger = New-ScheduledTaskTrigger -At 4:50am –Daily # Specify the trigger settings
#$User= "NT AUTHORITY\SYSTEM" # Specify the account to run the script
$User = "citspadminsvc" # Specify the account to run the script
$Action = New-ScheduledTaskAction -Execute "PowerShell.exe" -Argument "D:\Scripting\O365DevOps\AdminPortal-Operations\SyncO365Users.ps1" # Specify what program to run and with its parameters
Register-ScheduledTask -TaskName "SPO-SyncO365Users" -Trigger $Trigger -User $User -Action $Action -RunLevel Highest –Force # Specify the name of the task

#===================================================================================#
$Trigger = New-ScheduledTaskTrigger -At 4:50am –Daily # Specify the trigger settings
#$User= "NT AUTHORITY\SYSTEM" # Specify the account to run the script
$User = "citspadminsvc" # Specify the account to run the script
$Action = New-ScheduledTaskAction -Execute "PowerShell.exe" -Argument "D:\Scripting\O365DevOps\AdminPortal-Operations\SyncO365SPOSites.ps1" # Specify what program to run and with its parameters
Register-ScheduledTask -TaskName "SPO-SyncO365SPOSites" -Trigger $Trigger -User $User -Action $Action -RunLevel Highest –Force # Specify the name of the task

#===================================================================================#
$Trigger = New-ScheduledTaskTrigger -At 4:50am –Daily # Specify the trigger settings
#$User= "NT AUTHORITY\SYSTEM" # Specify the account to run the script
$User = "citspadminsvc" # Specify the account to run the script
$Action = New-ScheduledTaskAction -Execute "PowerShell.exe" -Argument "D:\Scripting\O365DevOps\AdminPortal-Operations\SyncO365PersonalSites.ps1" # Specify what program to run and with its parameters
Register-ScheduledTask -TaskName "SPO-SyncO365PersonalSites" -Trigger $Trigger -User $User -Action $Action -RunLevel Highest –Force # Specify the name of the task

#===================================================================================#
$Trigger = New-ScheduledTaskTrigger -DaysOfWeek Saturday -At 11:00am -Weekly # Specify the trigger settings
#$User= "NT AUTHORITY\SYSTEM" # Specify the account to run the script
$User = "citspadminsvc" # Specify the account to run the script
$Action = New-ScheduledTaskAction -Execute "PowerShell.exe" -Argument "D:\Scripting\O365DevOps\AdminPortal-Operations\PopulatePersonalSiteSCA.ps1" # Specify what program to run and with its parameters
Register-ScheduledTask -TaskName "SPO-PopulatePersonalSiteSCA" -Trigger $Trigger -User $User -Action $Action -RunLevel Highest –Force # Specify the name of the task

#===================================================================================#
$Trigger = New-ScheduledTaskTrigger -DaysOfWeek Saturday -At 10:00am –Weekly # Specify the trigger settings
#$User= "NT AUTHORITY\SYSTEM" # Specify the account to run the script
$User = "citspadminsvc" # Specify the account to run the script
$Action = New-ScheduledTaskAction -Execute "PowerShell.exe" -Argument "D:\Scripting\O365DevOps\AdminPortal-Operations\SyncPersonalSitesExtended.ps1" # Specify what program to run and with its parameters
Register-ScheduledTask -TaskName "SPO-SyncPersonalSitesExtended" -Trigger $Trigger -User $User -Action $Action -RunLevel Highest –Force # Specify the name of the task