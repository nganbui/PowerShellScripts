$cred = Get-Credential
$sfbSession = New-CsOnlineSession –OverrideAdminDomain "nih.onmicrosoft.com" #-Credential $cred
Import-PSSession $sfbSession -AllowClobber
​
# Teams Only users Report
#Name,Teams*​
#$users = Get-CsOnlineUser | ? {$_.TeamsUpgradeEffectiveMode -eq "TeamsOnly"} | select UserPrincipalName,FirstName,LastName,Company,Department,Office,Teams*
$users = Get-CsOnlineUser | ? {$_.TeamsUpgradeEffectiveMode -eq "TeamsOnly"} | select UserPrincipalName,Teams*
$users | Export-Csv -Path D:\Scripting\O365DevOps\Common\Data\Other\TeamsOnlyUsers.csv -NoTypeInformation

# All users report
​
#$allUsers = Get-CsOnlineUser | select *
#$allUsers | Export-Csv -Path "D:\Scripting\O365\Data\Other\AllTeamsUsers.csv" -NoTypeInformation