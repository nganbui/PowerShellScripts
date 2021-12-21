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
."$script:RootDir\Common\Lib\GraphAPILibSPOSites.ps1"
."$script:RootDir\Common\Lib\GraphAPILibSPOSitesDAO.ps1"
."$script:RootDir\Common\Lib\LibCache.ps1"

Function GeneratePersonalSitesSyncReport {    
    #------
    LogWrite -Message "Sending Email Report: [Sync M365 Personal Sites Extended Info]"    
    $subject = "[SPO-DevOps] M365 Personal Sites Extended Properties Sync Report"    
    $body = "<p><b>Description:</b> This job will sync the Personal Sites Extended properties [SecondarySCA;WebsCount;FilesCount;LastContentModifiedDate] to local database repository. </p>"
    $body += "<p>Script Start time: $($script:StartTimeDailyCache)<br />"
    $body += "Script End time: $($script:EndTimeDailyCache)<br /><br />"    
    $body += GetPersonalSitesReportContent          
    SendEmail -subject $subject -body $body
    LogWrite -Message "Sending Email Report: [Sync M365 Personal Sites Extended Info] completed."

}

Function GetPersonalSitesReportContent {
    $script:totalActivePersonalSites = @($script:personalSitesData).Count
    #$script:totalDeletedPersonalSites = @($script:deletedPersonalSitesData).Count

    $script:totalPersonalSitesInserted = @($script:personalSitesData | Where-Object {($_.OperationStatus -eq "Success") -and ($_.Operation -eq "Insert")}).URL.Count
    $script:totalPersonalSitesUpdated = @($script:personalSitesData | Where-Object {($_.OperationStatus -eq "Success") -and ($_.Operation -eq "Update")}).URL.Count
    $script:totalPersonalSitesUpdateFailed = @($script:personalSitesData | Where-Object {$_.OperationStatus -eq "Failed"}).URL.Count

    #$script:totalDeletedPersonalSitesInserted = @($script:deletedPersonalSitesData | Where-Object {($_.OperationStatus -eq "Success") -and ($_.Operation -eq "Insert")}).URL.Count
    #$script:totalDeletedPersonalSitesUpdated = @($script:deletedPersonalSitesData | Where-Object {($_.OperationStatus -eq "Success") -and ($_.Operation -eq "Update")}).URL.Count
    #$script:totalDeletedPersonalSitesUpdateFailed = @($script:deletedPersonalSitesData | Where-Object {$_.OperationStatus -eq "Failed"}).URL.Count

    LogWrite -Message "Generating Email Report..."
    LogWrite -Message "->Total Active SPO Personal Sites Found: $($script:totalActivePersonalSites)"
    #LogWrite -Message "->Total Soft Deleted SPO Personal Sites Found: $($script:totalDeletedPersonalSites)"
    
    LogWrite -Message "->Total Active SPO Personal Sites Records Added: $($script:totalPersonalSitesInserted)" 
    LogWrite -Message "->Total Active SPO Personal Sites Records Updated: $($script:totalPersonalSitesUpdated)" 
    LogWrite -Message "->Total Active SPO Personal Sites Records UpdateFailed: $($script:totalPersonalSitesUpdateFailed)"

    #LogWrite -Message "->Total Deleted SPO Personal Sites Records Added: $($script:totalDeletedPersonalSitesInserted)" 
    #LogWrite -Message "->Total Deleted SPO Personal Sites Records Updated: $($script:totalDeletedPersonalSitesUpdated)" 
    #LogWrite -Message "->Total Deleted SPO Personal Sites Records UpdateFailed: $($script:totalDeletedPersonalSitesUpdateFailed)"
        
    $msg ="<p>"
    $msg += "Total Active SPO Personal Sites Found: $($script:totalActivePersonalSites)<br />"
    #$msg += "Total Soft Deleted SPO Personal Sites Found: $($script:totalDeletedPersonalSites)<br />"
    $msg += "=============================================================<br />"
    $msg += "Total Active SPO Personal Sites Records Added: $($script:totalPersonalSitesInserted)<br />"
    $msg += "Total Active SPO Personal Sites Records Udated: $($script:totalPersonalSitesUpdated)<br />"
    $msg += "Total Active SPO Personal Sites Records UpdateFailed : $($script:totalPersonalSitesUpdateFailed)<br />"
    $msg += "=============================================================<br />"
    #$msg += "Total Deleted SPO Personal Sites Records Added: $($script:totalDeletedPersonalSitesInserted)<br />"
    #$msg += "Total Deleted SPO Personal Sites Records Updated: $($script:totalDeletedPersonalSitesUpdated)<br />" 
    #$msg += "Total Deleted SPO Personal Sites Records UpdateFailed: $($script:totalDeletedPersonalSitesUpdateFailed)<br />"
    $msg += "</p>"

    return $msg
}

Function GeneratePersonalSitesSyncLogs {    
    $logPath = "$($script:DirLog)"
    if (!(Test-Path $logPath)) { 
	    LogWrite -Message "Creating $logPath" 
        New-Item -ItemType "directory" -Path $logPath -Force
	} 

    LogWrite -Message "Generating Log files..." 

    $sitesFile = "$logPath\ActivePersonalSites.csv"
    #$delsitesFile = "$logPath\InActivePersonalSites.csv"    

    if ($script:personalSitesData) {
        ExportCSV -DataSet $script:personalSitesData -FileName $sitesFile
    }
    <#
    if ($script:deletedPersonalSitesData) {
        ExportCSV -DataSet $script:deletedPersonalSitesData -FileName $delsitesFile
    }
    #>
    
    LogWrite -Message "Generating Log files ended."     
}

Try {
    #log file path
    Set-LogFile -logFileName $logFileName
    $script:StartTimeDailyCache = Get-Date -Format "yyyy-MM-dd HH:mm:ss"           
    LogWrite -Message "----------------------- [Sync M365 Personal Sites Extended Info] Execution Started --------------------------"
    #Verify if the Data is already sync and cache is available for today    
    <#
    $script:personalSitesData = @()
    $script:deletedPersonalSitesData = @()
    $script:personalSitesData = GetDataInCache -CacheType O365 -ObjectType PersonalSites -ObjectState Active
    $script:deletedPersonalSitesData = GetDataInCache -CacheType O365 -ObjectType PersonalSites -ObjectState InActive

    if ($script:personalSitesData -eq $null) {
        LogWrite -Message "Personal sites data not found in cache. Processing from O365"
        #Retrieve All SPO Personal Sites - Active & InActive
        Set-TenantVars
        Set-DataFile
        GetAllPersonalSites
        #--For personal will get extended site props Weekly due to perfomance issue
        UpdateSitesProperties -SiteObjects $script:personalSitesData -SitesType PersonalSites
        #Cache All SPO Personal Sites to file system
        CachePersonalSites

        $script:personalSitesData = GetDataInCache -CacheType O365 -ObjectType PersonalSites -ObjectState Active
        $script:deletedPersonalSitesData= GetDataInCache -CacheType O365 -ObjectType PersonalSites -ObjectState InActive
    }
    else {
        LogWrite -Message "Processing Personal sites data from cache"
    }
    #>
    Set-TenantVars
    Set-AzureAppVars
    Set-DataFile
    Set-DBVars
    
    GetAllPersonalSites
    #--For personal will get extended site props Weekly due to perfomance issue
    UpdateSitesProperties -SiteObjects $script:personalSitesData -SitesType PersonalSites
    #Cache All SPO Personal Sites to file system
    CachePersonalSitesExtended
    $script:personalSitesExtendedData = GetDataInCache -CacheType O365 -ObjectType PersonalSitesExtended -ObjectState Active       
    #Update Personal Sites to Database
    if ($null -ne $script:personalSitesExtendedData){
        $script:personalSitesExtendedData = GetDataInCache -CacheType O365 -ObjectType PersonalSitesExtended -ObjectState Active
        UpdateSPOSiteExtenedToDatabase
        #Generate Log files and send Email Report
        $script:EndTimeDailyCache = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        GeneratePersonalSitesSyncReport
        #Generate Log files
        GeneratePersonalSitesSyncLogs    
    }   
    
    LogWrite -Message "[Sync M365 Personal Sites Extended Info] Start Time: $($script:StartTimeDailyCache)"
    LogWrite -Message "[Sync M365 Personal Sites Extended Info] End Time:   $($script:EndTimeDailyCache)"
    LogWrite -Message  "----------------------- [Sync M365 Personal Sites Extended Info] Execution Ended --------------------------"    
    
}
Catch [Exception] {
    LogWrite -Level ERROR "-Unexpected Error: $_ "
}
Finally {
    LogWrite -Message  "----------------------- [Sync M365 Personal Sites Extended Info to DB] Completed ------------------------"
}