$Trigger= New-ScheduledTaskTrigger -At 4:50am –Daily # Specify the trigger settings
#$User= "NT AUTHORITY\SYSTEM" # Specify the account to run the script
$User= "aabuint" # Specify the account to run the script
$Action= New-ScheduledTaskAction -Execute "PowerShell.exe" -Argument "D:\Scripting\O365DevOps\AdminPortal-Operations\ProcessM365Operations.ps1" # Specify what program to run and with its parameters
Register-ScheduledTask -TaskName "SPO-M365 Admin Portal Operation" -Trigger $Trigger -User $User -Action $Action -RunLevel Highest –Force # Specify the name of the task

#===================================================================================#
$Trigger= New-ScheduledTaskTrigger -At 5:00pm –Daily # Specify the trigger settings
#$User= "NT AUTHORITY\SYSTEM" # Specify the account to run the script
$User= "aabuint" # Specify the account to run the script
$Action= New-ScheduledTaskAction -Execute "PowerShell.exe" -Argument "D:\Scripting\O365DevOps\AdminPortal-Operations\PopulateDailyGroupsCache.ps1" # Specify what program to run and with its parameters
Register-ScheduledTask -TaskName "SPO-PopulateDailyGroupsCache" -Trigger $Trigger -User $User -Action $Action -RunLevel Highest –Force # Specify the name of the task

#===================================================================================#
$Trigger= New-ScheduledTaskTrigger -At 5:00pm –Daily # Specify the trigger settings
#$User= "NT AUTHORITY\SYSTEM" # Specify the account to run the script
$User= "aabuint" # Specify the account to run the script
$Action= New-ScheduledTaskAction -Execute "PowerShell.exe" -Argument "D:\Scripting\O365DevOps\AdminPortal-Operations\PopulateDailyUsersCache.ps1" # Specify what program to run and with its parameters
Register-ScheduledTask -TaskName "SPO-PopulateDailyUsersCache" -Trigger $Trigger -User $User -Action $Action -RunLevel Highest –Force # Specify the name of the task

#===================================================================================#
$Trigger= New-ScheduledTaskTrigger -At 5:00pm –Daily # Specify the trigger settings
#$User= "NT AUTHORITY\SYSTEM" # Specify the account to run the script
$User= "aabuint" # Specify the account to run the script
$Action= New-ScheduledTaskAction -Execute "PowerShell.exe" -Argument "D:\Scripting\O365DevOps\AdminPortal-Operations\PopulateDailySitesCache.ps1" # Specify what program to run and with its parameters
Register-ScheduledTask -TaskName "SPO-PopulateDailySitesCache" -Trigger $Trigger -User $User -Action $Action -RunLevel Highest –Force # Specify the name of the task

#===================================================================================#
$Trigger= New-ScheduledTaskTrigger -At 4:50am –Daily # Specify the trigger settings
#$User= "NT AUTHORITY\SYSTEM" # Specify the account to run the script
$User= "aabuint" # Specify the account to run the script
$Action= New-ScheduledTaskAction -Execute "PowerShell.exe" -Argument "D:\Scripting\O365DevOps\AdminPortal-Operations\SyncO365Groups.ps1" # Specify what program to run and with its parameters
Register-ScheduledTask -TaskName "SPO-SyncO365Groups" -Trigger $Trigger -User $User -Action $Action -RunLevel Highest –Force # Specify the name of the task

#===================================================================================#
$Trigger= New-ScheduledTaskTrigger -At 4:50am –Daily # Specify the trigger settings
#$User= "NT AUTHORITY\SYSTEM" # Specify the account to run the script
$User= "aabuint" # Specify the account to run the script
$Action= New-ScheduledTaskAction -Execute "PowerShell.exe" -Argument "D:\Scripting\O365DevOps\AdminPortal-Operations\SyncO365Users.ps1" # Specify what program to run and with its parameters
Register-ScheduledTask -TaskName "SPO-SyncO365Users" -Trigger $Trigger -User $User -Action $Action -RunLevel Highest –Force # Specify the name of the task

#===================================================================================#
$Trigger= New-ScheduledTaskTrigger -At 4:50am –Daily # Specify the trigger settings
#$User= "NT AUTHORITY\SYSTEM" # Specify the account to run the script
$User= "aabuint" # Specify the account to run the script
$Action= New-ScheduledTaskAction -Execute "PowerShell.exe" -Argument "D:\Scripting\O365DevOps\AdminPortal-Operations\SyncO365SPOSites.ps1" # Specify what program to run and with its parameters
Register-ScheduledTask -TaskName "SPO-SyncO365SPOSites" -Trigger $Trigger -User $User -Action $Action -RunLevel Highest –Force # Specify the name of the task

#===================================================================================#
$Trigger= New-ScheduledTaskTrigger -At 4:50am –Daily # Specify the trigger settings
#$User= "NT AUTHORITY\SYSTEM" # Specify the account to run the script
$User= "aabuint" # Specify the account to run the script
$Action= New-ScheduledTaskAction -Execute "PowerShell.exe" -Argument "D:\Scripting\O365DevOps\AdminPortal-Operations\SyncO365PersonalSites.ps1" # Specify what program to run and with its parameters
Register-ScheduledTask -TaskName "SPO-SyncO365PersonalSites" -Trigger $Trigger -User $User -Action $Action -RunLevel Highest –Force # Specify the name of the task

#===================================================================================#
$Trigger= New-ScheduledTaskTrigger -DaysOfWeek Friday -At 4:50pm -Weekly # Specify the trigger settings
#$User= "NT AUTHORITY\SYSTEM" # Specify the account to run the script
$User= "aabuint" # Specify the account to run the script
$Action= New-ScheduledTaskAction -Execute "PowerShell.exe" -Argument "D:\Scripting\O365DevOps\AdminPortal-Operations\PopulatePersonalSiteSCA.ps1" # Specify what program to run and with its parameters
Register-ScheduledTask -TaskName "SPO-PopulatePersonalSiteSCA" -Trigger $Trigger -User $User -Action $Action -RunLevel Highest –Force # Specify the name of the task

#===================================================================================#
$Trigger= New-ScheduledTaskTrigger -DaysOfWeek Saturday -At 4:50pm –Weekly # Specify the trigger settings
#$User= "NT AUTHORITY\SYSTEM" # Specify the account to run the script
$User= "aabuint" # Specify the account to run the script
$Action= New-ScheduledTaskAction -Execute "PowerShell.exe" -Argument "D:\Scripting\O365DevOps\AdminPortal-Operations\SyncPersonalSitesExtended.ps1" # Specify what program to run and with its parameters
Register-ScheduledTask -TaskName "SPO-SyncPersonalSitesExtended" -Trigger $Trigger -User $User -Action $Action -RunLevel Highest –Force # Specify the name of the task