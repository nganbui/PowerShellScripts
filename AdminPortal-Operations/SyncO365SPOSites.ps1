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


Function GenerateSPOSitesSyncReport {    
    #------
    LogWrite -Message "Sending Email Report: [Sync M365 SharePoint Sites]"    
    $subject = "[SPO-DevOps] M365 SharePoint Sites Daily Sync to DB"    
    $body = "<p><b>Description:</b> This job will sync the M365 SPO Sites to local database repository.  It updates all M365 SPO Sites information regardless of changes to one or multiple fields. </p>"
    $body += "<p>Script Start time: $($script:StartTimeDailyCache)<br />"
    $body += "Script End time: $($script:EndTimeDailyCache)<br /><br />"    
    $body += GetSPOSitesReportContent          
    SendEmail -subject $subject -body $body
    LogWrite -Message "Sending Email Report: [Sync M365 SharePoint Sites] completed."

}

Function GetSPOSitesReportContent {
    $script:totalActiveSites = @($script:sitesData).Count
    $script:totalDeletedSites = @($script:deletedSitesData).Count

    $script:totalSitesInserted = @($script:sitesData | Where-Object { ($_.OperationStatus -eq "Success") -and ($_.Operation -eq "Insert") }).URL.Count
    $script:totalSitesUpdated = @($script:sitesData | Where-Object { ($_.OperationStatus -eq "Success") -and ($_.Operation -eq "Update") }).URL.Count
    $script:totalSitesUpdateFailed = @($script:sitesData | Where-Object { $_.OperationStatus -eq "Failed" }).URL.Count

    $script:totalDeletedSitesInserted = @($script:deletedSitesData | Where-Object { ($_.OperationStatus -eq "Success") -and ($_.Operation -eq "Insert") }).URL.Count
    $script:totalDeletedSitesUpdated = @($script:deletedSitesData | Where-Object { ($_.OperationStatus -eq "Success") -and ($_.Operation -eq "Update") }).URL.Count
    $script:totalDeletedSitesUpdateFailed = @($script:deletedSitesData | Where-Object { $_.OperationStatus -eq "Failed" }).URL.Count

    LogWrite -Message "Generating Email Report..."
    LogWrite -Message "->Total Active SPO Sites Found: $($script:totalActiveSites)"
    LogWrite -Message "->Total Soft Deleted SPO Sites Found: $($script:totalDeletedSites)"
    
    LogWrite -Message "->Total Active SPO Sites Records Added: $($script:totalSitesInserted)" 
    LogWrite -Message "->Total Active SPO Sites Records Updated: $($script:totalSitesUpdated)" 
    LogWrite -Message "->Total Active SPO Sites Records UpdateFailed: $($script:totalSitesUpdateFailed)"

    LogWrite -Message "->Total Deleted SPO Sites Records Added: $($script:totalDeletedSitesInserted)" 
    LogWrite -Message "->Total Deleted SPO Sites Records Updated: $($script:totalDeletedSitesUpdated)" 
    LogWrite -Message "->Total Deleted SPO Sites Records UpdateFailed: $($script:totalDeletedSitesUpdateFailed)"
        
    $msg = ""
    $msg += "Total Active SPO Sites Found: $($script:totalActiveSites)<br />"
    $msg += "Total Soft Deleted SPO Sites Found: $($script:totalDeletedSites)<br />"
    $msg += "=============================================================<br />"
    $msg += "Total Active SPO Sites Records Added: $($script:totalSitesInserted)<br />"
    $msg += "Total Active SPO Sites Records Udated: $($script:totalSitesUpdated)<br />"
    $msg += "Total Active SPO Sites Records UpdateFailed : $($script:totalSitesUpdateFailed)<br />"
    $msg += "=============================================================<br />"
    $msg += "Total Deleted SPO Sites Records Added: $($script:totalDeletedSitesInserted)<br />"
    $msg += "Total Deleted SPO Sites Records Updated: $($script:totalDeletedSitesUpdated)<br />" 
    $msg += "Total Deleted SPO Sites Records UpdateFailed: $($script:totalDeletedSitesUpdateFailed)<br />"
    $msg += "</p>"
    return $msg
}

Function GenerateSPOSitesSyncLogs {    
    $todaysDate = Get-Date -Format "MM-dd-yyyy"
    $logPath = "$script:LogFile\$todaysDate"
    if (!(Test-Path $logPath)) { 
	    LogWrite -Message "Creating $logPath" 
        New-Item -ItemType "directory" -Path $logPath -Force
	} 

    LogWrite -Message "Generating Log files..." 

    $sitesFile = "$logPath\ActiveSPOSites.csv"
    $delsitesFile = "$logPath\DeletedSPOSites.csv"

    if ($script:sitesData) {
        ExportCSV -DataSet $script:sitesData -FileName $sitesFile
    }
    if ($script:deletedSitesData) {
        ExportCSV -DataSet $script:deletedSitesData -FileName $delsitesFile
    }
    
    LogWrite -Message "Generating Log files ended."     
}

Try {
    #log file path
    Set-LogFile -logFileName $logFileName
    Set-DataFile
    $script:StartTimeDailyCache = Get-Date -Format "yyyy-MM-dd HH:mm:ss"           
    LogWrite -Message "----------------------- [Sync M365 SharePoint Sites] Execution Started --------------------------"
    #Verify if the Data is already sync and cache is available for today
    $script:sitesData = @()
    $script:deletedSitesData = @()
    $script:sitesData = GetDataInCache -CacheType O365 -ObjectType SPOSites -ObjectState Active
    $script:deletedSitesData = GetDataInCache -CacheType O365 -ObjectType SPOSites -ObjectState InActive

    if ($script:sitesData -eq $null) {
        LogWrite -Message "M365 SharePoint sites data not found in cache. Processing from M365"
        #Retrieve All SPO Sites - Active & InActive
        Set-TenantVars
        Set-AzureAppVars       
        GetAllSPOSites
        #Cache All SPO Sites to file system
        CacheSPOSites

        $script:sitesData = GetDataInCache -CacheType O365 -ObjectType SPOSites -ObjectState Active
        $script:deletedSitesData = GetDataInCache -CacheType O365 -ObjectType SPOSites -ObjectState InActive
    }
    else {
        LogWrite -Message "Processing M365 SharePoint sites data from cache"
    }
    Set-DBVars
    #Update Sites and Personal Sites to Database
    UpdateSPOSitesToDatabase
    $script:EndTimeDailyCache = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    #Generate Log files and send Email Report
    GenerateSPOSitesSyncReport
    #Generate Log files
    GenerateSPOSitesSyncLogs   
     
    LogWrite -Message "[Sync M365 SharePoint Sites] Start Time: $($script:StartTimeDailyCache)"
    LogWrite -Message "[Sync M365 SharePoint Sites] End Time:   $($script:EndTimeDailyCache)"
    LogWrite -Message  "----------------------- [Sync SPO Sites] Execution Ended --------------------------"    
    
}
Catch [Exception] {
    LogWrite -Level ERROR "-Unexpected Error: $_ "
}
Finally {
    LogWrite -Message  "----------------------- [Sync M365 SharePoint Sites to DB] Completed ------------------------"
}