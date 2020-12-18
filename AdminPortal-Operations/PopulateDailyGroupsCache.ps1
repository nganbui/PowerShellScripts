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
."$script:RootDir\Common\Lib\GraphAPILibO365Groups.ps1"
."$script:RootDir\Common\Lib\LibCache.ps1"

<#
      ===========================================================================
      .DESCRIPTION
        Populate Groups/Teams/TeamChannel to .csv cache
        -Using Teams API (Get-TeamChannel and Get-TeamUser) to get team channel and members (currently Graph API not support to get Team Channel if the Team has Private channel-getting Unauthozired error)
        -Using Graph API to get the list of members of public/private channel
            Public (Standard) channel do not have "members". The Team(basically group) has "members", 
            In order to get the list of Public members by using group members(Graph API or Exchange Online)
            In order to get the list of Private members by using Graph API (https://graph.microsoft.com/v1.0/teams/<group_id>/channels/<channel_id>/members)
        
#>

Function SyncM365GroupsDataToCache {    
    Set-TenantVars
    Set-AzureAppVars
    Set-DataFile
    
    LogWrite -Message "------------------------ Retrieving M365 Groups-Teams-TeamChannel -----------------------------------------"    
    GetAllO365Groups
    LogWrite -Message "------------------------ Caching M365 Groups-Teams-TeamChannel --------------------------------------------"     
    CacheO365Groups  
    
}

Function GenerateDailyCacheReport {
    LogWrite -Message "Sending Email Report: [Populate Groups-Teams-Channel Daily Cache]"    
    $subject = "[SPO-DevOps] M365 Groups-Teams-Channel Daily Cache"    
    $body = "<p><b>Description:</b> This job will cache all M365 Groups-Teams-Team Channel data objects locally. These cache files are used to further sync data to the Database</p>"
    $body += "<p>Script Start time: $($script:StartTimeM365GroupsDailyCache)<br />"
    $body += "Script End time: $($script:EndTimeM365GroupsDailyCache)<br /><br />"    
    $body += GenerateM365GroupsDailyCacheReport          
    SendEmail -subject $subject -body $body
    LogWrite -Message "Sending Email Report: [Populate Groups-Teams-Channel Daily Cache] completed."
}

Function GenerateM365GroupsDailyCacheReport {
    $script:totalGroups = $script:o365GroupsData.Count
    $script:totalTeams = $script:o365TeamsData.Count
    $script:totalDelGroups = $script:o365DeletedGroupsData.Count    

    LogWrite -Message "Generating Email Report..."
    #
    LogWrite -Message "->Total Active Groups Retrieved: $($script:totalGroups)"
    LogWrite -Message "->Total Active Teams Retrieved: $($script:totalTeams)"
    LogWrite -Message "->Total Private channel: $($script:totalPrivateChannel)"
    LogWrite -Message "->Total Deleted Groups Retrieved: $($script:totalDelGroups)"
        
    $msg = ""
    $msg += "<p>Total Active Groups Retrieved: $($script:totalGroups)<br />"
    $msg += " Total Active Teams Retrieved: $($script:totalTeams)<br />"
    $msg += " Total Private channel: $($script:totalPrivateChannel)<br />"
    $msg += " Total Deleted Groups Retrieved: $($script:totalDelGroups)<br />"   
    return $msg
}


Try {
    #log file path
    Set-LogFile -logFileName $logFileName
    $script:StartTimeM365GroupsDailyCache = Get-Date -Format "yyyy-MM-dd HH:mm:ss"           
    LogWrite -Message "----------------------- [Populate Groups-Teams-Channel Daily Cache] Execution Started -----------------------"

    #Sync Groups-Teams-Channel Objects to the Cache
    SyncM365GroupsDataToCache 

    $script:EndTimeM365GroupsDailyCache = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    #Generate Report and send email
    GenerateDailyCacheReport 
    LogWrite -Message "[Populate Groups-Teams-Channel Daily Cache] Start Time: $($script:StartTimeM365GroupsDailyCache)"
    LogWrite -Message "[Populate Groups-Teams-Channel Daily Cache] End Time:   $($script:EndTimeM365GroupsDailyCache)"
    LogWrite -Message  "----------------------- [Populate Groups-Teams-Channel Daily Cache] Execution Ended ------------------------"    
    
}
Catch [Exception] {
    LogWrite -Level ERROR "-Unexpected Error: $_ "
}
Finally {
    LogWrite -Message  "----------------------- [Populate Groups-Teams-Channel Daily Cache] Completed ------------------------"
}
