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
."$script:RootDir\Common\Lib\GraphAPILibO365GroupsDAO.ps1"
."$script:RootDir\Common\Lib\LibCache.ps1"

Function GenerateM365GroupsSyncReport {    
    #------
    LogWrite -Message "Sending Email Report: [Sync M365 Groups-Teams-Channel]"    
    $subject = "[SPO-DevOps] M365 Groups-Teams-Channel Daily Sync to DB"    
    $body = "<p><b>Description:</b> This job will sync the M365 Groups-Teams-Channel to local database repository.  It updates all the Groups-Teams-Channel information regardless of changes to one or multiple fields. </p>"
    $body += "<p>Script Start time: $($script:startTimeO365GroupsSync)<br />"
    $body += "Script End time: $($script:EndTimeO365GroupsSync)<br /><br />"    
    $body += GetM365GroupsReportContent          
    SendEmail -subject $subject -body $body
    LogWrite -Message "Sending Email Report: [Sync M365 Groups-Teams-Channel] completed."

}

Function GetM365GroupsReportContent {
    #group
    $script:totalGroups = $script:o365GroupsData.Count    
    $script:totalGroupsAdded = @($script:o365GroupsData | Where-Object { ($_.OperationStatus -eq "Success") -and ($_.Operation -eq "Insert") }).Count
    $script:totalGroupsUpdated = @($script:o365GroupsData | Where-Object { ($_.OperationStatus -eq "Success") -and ($_.Operation -eq "Update") }).Count    
    $script:totalGroupsUpdateFailed = @($script:o365GroupsData | Where-Object { $_.OperationStatus -eq "Failed" }).Count

    $script:totalDelGroups = $script:o365DeletedGroupsData.Count    
    $script:totalDelGroupsAdded = @($script:o365DeletedGroupsData | Where-Object { ($_.OperationStatus -eq "Success") -and ($_.Operation -eq "Insert") }).Count
    $script:totalDelGroupsUpdated = @($script:o365DeletedGroupsData | Where-Object { ($_.OperationStatus -eq "Success") -and ($_.Operation -eq "Update") }).Count
    $script:totalDelGroupsUpdateFailed = @($script:o365DeletedGroupsData | Where-Object { $_.OperationStatus -eq "Failed" }).Count

    #team
    $script:totalTeams = $script:o365TeamsData.Count    
    $script:totalTeamsAdded = @($script:o365TeamsData | Where-Object { ($_.OperationStatus -eq "Success") -and ($_.Operation -eq "Insert") }).Count
    $script:totalTeamsUpdated = @($script:o365TeamsData | Where-Object { ($_.OperationStatus -eq "Success") -and ($_.Operation -eq "Update") }).Count
    $script:totalTeamsUpdateFailed = @($script:o365TeamsData | Where-Object { $_.OperationStatus -eq "Failed" }).Count
    #teamchannel
    $script:totalTeamChannel = $script:TeamsChannelData.Count    
    $script:totalTeamChannelAdded = @($script:TeamsChannelData | Where-Object { ($_.OperationStatus -eq "Success") -and ($_.Operation -eq "Insert") }).Count
    $script:totalTeamChannelUpdated = @($script:TeamsChannelData | Where-Object { ($_.OperationStatus -eq "Success") -and ($_.Operation -eq "Update") }).Count    
    $script:totalTeamChannelUpdateFailed = @($script:TeamsChannelData | Where-Object { $_.OperationStatus -eq "Failed" }).Count

    LogWrite -Message "Generating Email Report..."
    #----    
    LogWrite -Message "->Total Active Groups Retrieved: $($script:totalGroups)"
    LogWrite -Message "->Total Deleted Groups Retrieved: $($script:totalDelGroups)"
    LogWrite -Message "->Total Teams Retrieved: $($script:totalTeams)/$($script:totalGroups) groups"
    LogWrite -Message "->Total Teams Channel Retrieved: $($script:totalTeamChannel)"
    #-----
    LogWrite -Message "->Total Active Group Records Inserted: $($script:totalGroupsAdded)" 
    LogWrite -Message "->Total Active Group Records Updated: $($script:totalGroupsUpdated)" 
    LogWrite -Message "->Total Active Group Records UpdateFailed: $($script:totalGroupsUpdateFailed)" 
    #-----
    LogWrite -Message "->Total Deleted Group Records Inserted: $($script:totalDelGroupsAdded)" 
    LogWrite -Message "->Total Deleted Group Records Updated: $($script:totalDelGroupsUpdated)" 
    LogWrite -Message "->Total Deleted Group Records UpdateFailed: $($script:totalDelGroupsUpdateFailed)"   
    #-----
    LogWrite -Message "->Total Teams Records Inserted: $($script:totalTeamsAdded)" 
    LogWrite -Message "->Total Teams Records Updated: $($script:totalTeamsUpdated)" 
    LogWrite -Message "->Total Teams Records UpdateFailed: $($script:totalTeamsUpdateFailed)"   
    #-----
    LogWrite -Message "->Total Teams Channel Records Inserted: $($script:totalTeamChannelAdded)" 
    LogWrite -Message "->Total Teams Channel Records Updated: $($script:totalTeamChannelUpdated)" 
    LogWrite -Message "->Total Teams Channel Records UpdateFailed: $($script:totalTeamChannelUpdateFailed)" 
        
    $msg = "<p>"
    $msg += "<p>Total Active Groups Retrieved: $($script:totalGroups)<br />"
    $msg += "Total Active Group Records Inserted: $($script:totalGroupsAdded)<br />"
    $msg += "Total Active Group Records Updated: $($script:totalGroupsUpdated)<br />"
    $msg += "Total Active Group Records Failed to Insert/Update: $($script:totalGroupsUpdateFailed)<br />"
    $msg += "=============================================================<br />"
    $msg += "Total Deleted Groups Retrieved: $($script:totalDelGroups)<br />"
    $msg += "Total Deleted Group Records Inserted: $($script:totalDelGroupsAdded)<br />"
    $msg += "Total Deleted Group Records Updated: $($script:totalDelGroupsUpdated)<br />"
    $msg += "Total Deleted Group Records Failed to Insert/Update: $($script:totalDelGroupsUpdateFailed)<br />" 
    $msg += "=============================================================<br />"
    $msg += "Total Teams Retrieved: $($script:totalTeams)/$($script:totalGroups) groups<br />"
    $msg += "Total Teams Records Inserted: $($script:totalTeamsAdded)<br />"
    $msg += "Total Teams Records Updated: $($script:totalTeamsUpdated)<br />"
    $msg += "Total Teams Records Failed to Insert/Update: $($script:totalTeamsUpdateFailed)<br />"
    $msg += "=============================================================<br />"
    $msg += "Total Teams Channel Retrieved: $($script:totalTeamChannel)<br />"
    $msg += "Total Teams Channel Records Inserted: $($script:totalTeamChannelAdded)<br />"
    $msg += "Total Teams Channel Records Updated: $($script:totalTeamChannelUpdated)<br />"
    $msg += "Total Teams Channel Records Failed to Insert/Update: $($script:totalTeamChannelUpdateFailed)</p>"                

    return $msg
}

Function GenerateM365GroupsSyncLogs {
    $todaysDate = Get-Date -Format "MM-dd-yyyy"
    $logPath = "$script:LogFile\$todaysDate"
    if (!(Test-Path $logPath)) { 
	    LogWrite -Message "Creating $logPath" 
        New-Item -ItemType "directory" -Path $logPath -Force
	}     
    $groupsFile = "$logPath\ActiveGraphAPIGroups.csv"
    $delGroupsFile = "$logPath\InActiveGraphAPIGroups.csv"
    $teamsFile = "$logPath\ActiveGraphAPITeams.csv"
    $channelFile = "$logPath\ActiveGraphAPIChannel.csv"

    LogWrite -Message "Generating Log files..." 
    if ($script:o365GroupsData) {
        ExportCSV -DataSet $script:o365GroupsData -FileName $groupsFile
    }
    if ($script:o365DeletedGroupsData) {
        ExportCSV -DataSet $script:o365DeletedGroupsData -FileName $delGroupsFile
    }
    if ($script:o365TeamsData) {
        ExportCSV -DataSet $script:o365TeamsData -FileName $teamsFile
    }
    if ($script:TeamsChannelData) {        
        ExportCSV -DataSet $script:TeamsChannelData -FileName $channelFile 
    }
    LogWrite -Message "Generating Log files ended."     
}

Try {
    #log file path
    Set-LogFile -logFileName $logFileName
    Set-DataFile
    $script:startTimeO365GroupsSync = Get-Date -Format "yyyy-MM-dd HH:mm:ss"           
    LogWrite -Message "----------------------- [Sync M365 Groups-Teams-Channel to DB] Execution Started --------------------------"
    #Verify if the Data is already sync and cache is available for today
    #Get Groups/Teams/Channel from Cache
    $script:o365GroupsData = GetDataInCache -CacheType O365 -ObjectType GraphAPIGroups -ObjectState Active    
    $script:o365TeamsData = GetDataInCache -CacheType O365 -ObjectType GraphAPITeams -ObjectState Active
    $script:TeamsChannelData = GetDataInCache -CacheType O365 -ObjectType GraphAPIChannel -ObjectState Active 
    $script:o365DeletedGroupsData = GetDataInCache -CacheType O365 -ObjectType GraphAPIGroups -ObjectState InActive

    if ($null -eq $script:o365GroupsData) {
        LogWrite -Message "M365 Groups data not found in cache. Processing from O365"
        #Retrieve O365 Groups...
        Set-TenantVars
        Set-AzureAppVars        
        GetAllO365Groups
        #Cache O365 Groups
        CacheO365Groups

        $script:o365GroupsData = GetDataInCache -CacheType O365 -ObjectType GraphAPIGroups -ObjectState Active
        $script:o365TeamsData = GetDataInCache -CacheType O365 -ObjectType GraphAPITeams -ObjectState Active
        $script:TeamsChannelData = GetDataInCache -CacheType O365 -ObjectType GraphAPIChannel -ObjectState Active
        $script:o365DeletedGroupsData = GetDataInCache -CacheType O365 -ObjectType GraphAPIGroups -ObjectState InActive 
    }
    else {
        LogWrite -Message "Processing M365 Groups data from cache"
    }

    $script:o365GroupsData = @($script:o365GroupsData)
    $script:o365TeamsData = @($script:o365TeamsData)
    $script:TeamsChannelData = @($script:TeamsChannelData)
    $script:o365DeletedGroupsData = @($script:o365DeletedGroupsData)

    Set-DBVars
    UpdateO365GroupsToDatabase
   
    $script:EndTimeO365GroupsSync = Get-Date -Format "yyyy-MM-dd HH:mm:ss" 
    #Generate Log files and send Email Report
    GenerateM365GroupsSyncReport
    #Generate Log files
    GenerateM365GroupsSyncLogs
      
    LogWrite -Message "[Sync M365 Groups-Teams-Channel] Start Time: $($script:startTimeO365GroupsSync)"
    LogWrite -Message "[Sync M365 Groups-Teams-Channel] End Time:   $($script:EndTimeO365GroupsSync)"
    LogWrite -Message  "----------------------- [Sync M365 Groups-Teams-Channel to DB] Execution Ended --------------------------"    
    
}
Catch [Exception] {
    LogWrite -Level ERROR "-Unexpected Error: $_ "
}
Finally {
    LogWrite -Message  "----------------------- [Sync M365 Groups-Teams-Channel to DB] Completed ------------------------"
}
