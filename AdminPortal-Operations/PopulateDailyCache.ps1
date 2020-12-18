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
."$script:RootDir\Common\Lib\LibSPOSites.ps1"
."$script:RootDir\Common\Lib\LibO365Users.ps1"
."$script:RootDir\Common\Lib\GraphAPILibO365Groups.ps1"
."$script:RootDir\Common\Lib\LibCache.ps1"

Function SyncO365DataToCache {    
    Set-TenantVars
    Set-AzureAppVars
    Set-DataFile
           
    LogWrite -Message "------------------------Retrieving SPO Sites and Personal Sites------------------------"     
    GetAllSPOSites
    LogWrite -Message "------------------------Caching SPO Sites and Personal Sites---------------------------"     
    CacheSPOSites
    CachePersonalSites     
    
    <#LogWrite -Message "------------------------Retrieving O365 Users------------------------------------------"     
    GetAllM365Users
    LogWrite -Message "------------------------Caching O365 Users---------------------------------------------"    
    CacheO365Users
    #>
    
    <#LogWrite -Message "------------------------Retrieving O365 Groups-----------------------------------------"    
    GetAllO365Groups
    LogWrite -Message "------------------------Caching O365 Groups--------------------------------------------"     
    CacheO365Groups  
    #>
}

Function GenerateDailyCacheReport {
    LogWrite -Message "Sending Email Report: [Populate Daily Cache]"    
    $subject = "[SPO-DevOps] O365 Daily Cache"    
    $body = "<p><b>Description:</b> This job will cache all O365 Data objects locally. These cache files are used to further sync data to the Database</p>"
    $body += "<p>Script Start time: $($script:StartTimeDailyCache)<br />"
    $body += "Script End time: $($script:EndTimeDailyCache)<br /><br />"    
    $body += GetDailyCacheReportContent          
    SendEmail -subject $subject -body $body
    LogWrite -Message "Sending Email Report: [Populate Daily Cache] completed."
}

Function GetDailyCacheReportContent {
    $script:totalGroups = $script:o365GroupsData.Count
    $script:totalTeams = $script:o365TeamsData.Count    
    $script:totalUsers = $script:usersData.Count    
    $script:totalSites = $script:sitesData.Count
    $script:totalPersonalSites = $script:personalSitesData.Count
    
    $script:totalDelGroups = $script:o365DeletedGroupsData.Count
    $script:totalDeletedUsers = $script:deletedUsersData.Count
    $script:totalDeletedSites = $script:deletedSitesData.Count
    $script:totalDeletedPersonalSites = $script:deletedPersonalSitesData.Count

    LogWrite -Message "Generating Email Report..."
    #
    LogWrite -Message "->Total Active Groups Retrieved: $($script:totalGroups)"
    LogWrite -Message "->Total Active Teams Retrieved: $($script:totalTeams)"
    LogWrite -Message "->Total Active Users Retrieved: $($script:totalUsers)"
    LogWrite -Message "->Total Active Sites Retrieved: $($script:totalSites)"
    LogWrite -Message "->Total Active Personal Sites Retrieved: $($script:totalPersonalSites)"

    LogWrite -Message "->Total Deleted Groups Retrieved: $($script:totalDelGroups)"
    LogWrite -Message "->Total Deleted Users Retrieved: $($script:totalDeletedUsers)"
    LogWrite -Message "->Total Deleted Sites Retrieved: $($script:totalDeletedSites)"
    LogWrite -Message "->Total Deleted Personal Sites Retrieved: $($script:totalDeletedPersonalSites)"
   
        
    $msg = ""
    $msg += "<p>Total Active Groups Retrieved: $($script:totalGroups)<br />"
    $msg += " Active Teams Retrieved: $($script:totalTeams)<br />"
    $msg += " Active Users Retrieved: $($script:totalUsers)<br />"
    $msg += "Total Active Sites Retrieved: $($script:totalSites)<br />"
    $msg += "Total Active Personal Sites Retrieved: $($script:totalPersonalSites)<br />"

    $msg += "Total Deleted Groups Retrieved: $($script:totalDelGroups)<br />"
    $msg += "Total Deleted Users Retrieved: $($script:totalDeletedUsers)<br />"
    $msg += "Total Deleted Sites Retrieved: $($script:totalDeletedSites)<br />"
    $msg += "Total Deleted Personal Sites Retrieved: $($script:totalDeletedPersonalSites)<br /></p>"
    return $msg
}


Try {
    #log file path
    Set-LogFile -logFileName $logFileName
    $script:StartTimeDailyCache = Get-Date -Format "yyyy-MM-dd HH:mm:ss"           
    LogWrite -Message "----------------------- [Populate Daily Cache] Execution Started -----------------------"

    #Sync all O365 Objects to the Cache
    SyncO365DataToCache 

    $script:EndTimeDailyCache = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    #Generate Report and send email
    GenerateDailyCacheReport 
    LogWrite -Message "[Populate Daily Cache] Start Time: $($script:StartTimeDailyCache)"
    LogWrite -Message "[Populate Daily Cache] End Time:   $($script:EndTimeDailyCache)"
    LogWrite -Message  "----------------------- [Populate Daily Cache] Execution Ended ------------------------"    
    
}
Catch [Exception] {

    
}
Finally {
    
}
