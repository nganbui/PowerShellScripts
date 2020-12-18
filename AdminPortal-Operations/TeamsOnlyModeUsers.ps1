$0 = $MyInvocation.MyCommand.Definition
$dp0 = [System.IO.Path]::GetDirectoryName($0)
$currentFolder = $dp0.Substring($dp0.LastIndexOf('\')+1)
$script:RootDir = Resolve-Path "$dp0\.." 
$logFileName = $MyInvocation.MyCommand.Name.Substring(0, $MyInvocation.MyCommand.Name.IndexOf('.'))
$logFileName = "$currentFolder\$logFileName"
$usageReport = "UsageReports"
$inputReport = "Input"
#Include core
."$script:RootDir\Common\Lib\LibCore.ps1"
#Include dependent functionality
."$script:RootDir\Common\Lib\LibO365.ps1"
."$script:RootDir\Common\Lib\LibCache.ps1"


Try {
    #------------------Initialize global variables needed------------------#
    Set-LogFile -logFileName $logFileName
    Set-TenantVars
    Set-AzureAppVars
    Set-DataFile  
    #------------------End-Initialize global variables---------------------#
    
    LogWrite -Message "----------------------- [Populate Teams Only Mode Users Cache] Execution Started -----------------------"
    
    #$cred = Get-Credential
    $sfbSession = New-CsOnlineSession –OverrideAdminDomain "nih.onmicrosoft.com" #-Credential $cred
    Import-PSSession $sfbSession -AllowClobber
    # Teams Only users Report
    #Name,Teams*​
    #$users = Get-CsOnlineUser | ? {$_.TeamsUpgradeEffectiveMode -eq "TeamsOnly"} | select UserPrincipalName,FirstName,LastName,Company,Department,Office,Teams*

    #--Create a folder UsageReports under Data if any        
    $date = Get-Date
    $year = $date.Year
    $month = $date.AddMonths(-1).Month
    $monthName = (Get-Culture).DateTimeFormat.GetMonthName($month)
    $reportFolder = "$($script:CacheDataPath)\$($usageReport)\$($monthName)\$($inputReport)"
    Create-Directory $reportFolder

    # create a new DateTime object set to the first day of a given month and year
    $startOfMonth = Get-Date -Year $year -Month $month -Day 1 -Hour 0 -Minute 0 -Second 0 -Millisecond 0
    # add a month and subtract the smallest possible time unit
    $endOfMonth = ($startOfMonth).AddMonths(1).AddTicks(-1)

    $users = Get-CsOnlineUser | ? {$_.TeamsUpgradeEffectiveMode -eq "TeamsOnly"} | select UserPrincipalName,Teams*
    $users | Export-Csv -Path "$reportFolder\TeamsOnlyUsers.csv" -NoTypeInformation
    # All users report
    #$allUsers = Get-CsOnlineUser | select *
    #$allUsers | Export-Csv -Path "D:\Scripting\O365\Data\Other\AllTeamsUsers.csv" -NoTypeInformation
    LogWrite -Message  "----------------------- [Populate Teams Only Mode Users Cache] Execution Ended ------------------------"     
}
Catch [Exception] {
    LogWrite -Level ERROR "-Unexpected Error: $_ "
}
Finally {
    LogWrite -Message  "----------------------- [Populate Teams Only Mode Users Cache] Completed ------------------------"
    Remove-PSSession -Session $sfbSession
}