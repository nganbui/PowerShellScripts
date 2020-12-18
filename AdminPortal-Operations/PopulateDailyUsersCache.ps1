$0 = $MyInvocation.MyCommand.Definition
$dp0 = [System.IO.Path]::GetDirectoryName($0)
$currentFolder = $dp0.Substring($dp0.LastIndexOf('\')+1)
$script:RootDir = Resolve-Path "$dp0\.." 
$logFileName = $MyInvocation.MyCommand.Name.Substring(0, $MyInvocation.MyCommand.Name.IndexOf('.'))
$logFileName = "$currentFolder\$logFileName"
#Include core
."$script:RootDir\Common\Lib\LibCore.ps1"
#Include dependent functionality
."$script:RootDir\Common\Lib\LibO365.ps1"
."$script:RootDir\Common\Lib\GraphAPILibO365Users.ps1"
."$script:RootDir\Common\Lib\LibCache.ps1"

<#
      ===========================================================================
      .DESCRIPTION
        Populate Users to .csv cache 
        Using Graph API to fetch Users object and save to .csv file
        Beta version: "signInActivity"
        https://graph.microsoft.com/beta/users?$select=displayName,userPrincipalName,mail,id,CreatedDateTime,creationType,signInActivity,UserType,externalUserState,externalUserStateChangeDateTime&$top=999
#>

Function SyncM365UsersDataToCache {    
    Set-TenantVars
    Set-AzureAppVars
    Set-DataFile

    LogWrite -Message "------------------------ Retrieving M365 Users ------------------------------------------"     
    GetAllM365Users
    LogWrite -Message "------------------------ Caching M365 Users ---------------------------------------------"    
    CacheO365Users
}

Function GenerateDailyCacheReport {
    LogWrite -Message "Sending Email Report: [Populate M365 Users Daily Cache]"    
    $subject = "[SPO-DevOps] Populate M365 Users Daily Cache"    
    $body = "<p><b>Description:</b> This job will cache M365 Users Data objects locally. These cache files are used to further sync data to the Database</p>"
    $body += "<p>Script Start time: $($script:StartTimeM365UsersDailyCache)<br />"
    $body += "Script End time: $($script:EndTimeM365UsersDailyCache)<br /><br />"    
    $body += GenerateM365UsersDailyCacheReport          
    SendEmail -subject $subject -body $body
    LogWrite -Message "Sending Email Report: [Populate M365 Users Daily Cache] completed."
}

Function GenerateM365UsersDailyCacheReport {
    $script:totalUsers = @($script:o365UsersData).Count
    $script:totalDeletedUsers = @($script:o365DeletedUsersData).Count

    LogWrite -Message "->Total Active Users Retrieved: $($script:totalUsers)"
    LogWrite -Message "->Total Members Retrieved: $($script:totalMembers)"
    LogWrite -Message "->Total Guests  Retrieved: $($script:totalGuests)"
    LogWrite -Message "->Total Others  Retrieved: $($script:totalOthers)"    
    LogWrite -Message "->Total Deleted Users Retrieved: $($script:totalDeletedUsers)"
    $msg = ""
    $msg += "<p>Total Active Users Retrieved: $($script:totalUsers)<br />"    
    $msg += "Total Members Retrieved: $($script:totalMembers)<br />"
    $msg += "Total Guests Retrieved: $($script:totalGuests)<br />"
    $msg += "Total Others Retrieved: $($script:totalOthers)<br />"
    $msg += "Total Deleted Users Retrieved: $($script:totalDeletedUsers)</p>"

    return $msg    
}


Try {
    #log file path
    Set-LogFile -logFileName $logFileName
    $script:StartTimeM365UsersDailyCache = Get-Date -Format "yyyy-MM-dd HH:mm:ss"           
    LogWrite -Message "----------------------- [Populate M365 Users Daily Cache] Execution Started -----------------------"

    #Sync M365 Users Objects to the Cache    
    SyncM365UsersDataToCache

    $script:EndTimeM365UsersDailyCache = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    #Generate Report and send email
    GenerateDailyCacheReport 
    LogWrite -Message "[Populate M365 Users Daily Cache] Start Time: $($script:StartTimeM365UsersDailyCache)"
    LogWrite -Message "[Populate M365 Users Daily Cache] End Time:   $($script:EndTimeM365UsersDailyCache)"
    LogWrite -Message  "----------------------- [Populate M365 Users Daily Cache] Execution Ended ------------------------"    
    
}
Catch [Exception] {
    LogWrite -Level ERROR "-Unexpected Error: $_ "
}
Finally {
    LogWrite -Message  "----------------------- [Populate M365 Users Daily Cache] Completed ------------------------"
}
