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
."$script:RootDir\Common\Lib\LibCache.ps1"

<#
      ===========================================================================
      .DESCRIPTION
        Populate Sites and Personal Sites to .csv cache 
        Using PnP to fetch Site and Personal Site object and save to .csv file        
#>

Function SyncSitesDataToCache {    
    Set-TenantVars
    Set-AzureAppVars
    Set-DataFile

    LogWrite -Message "------------------------ Retrieving SPO Sites and Personal Sites ------------------------------------------"     
    GetAllSPOSites
    #--get extended site props such as:CreatedDate,OwnerEmail,WebsCount,Hub,SharingCapability      
    UpdateSitesProperties -SiteObjects $script:sitesData -SitesType Sites
    GetAllPersonalSites
    LogWrite -Message "------------------------ Caching SPO Sites and Personal Sites ---------------------------------------------"    
    CacheSPOSites
    CachePersonalSites 
}

Function GenerateDailyCacheReport {
    LogWrite -Message "Sending Email Report: [Populate Sites and Personal Sites Daily Cache]"    
    $subject = "[SPO-DevOps] Populate Sites and Personal Sites Daily Cache"    
    $body = "<p><b>Description:</b> This job will cache Site and Personal Site Data objects locally. These cache files are used to further sync data to the Database</p>"
    $body += "<p>Script Start time: $($script:StartTimeSitesDailyCache)<br />"
    $body += "Script End time: $($script:EndTimeSitesDailyCache)<br /><br />"    
    $body += GenerateSitesDailyCacheReport          
    SendEmail -subject $subject -body $body
    LogWrite -Message "Sending Email Report: [Populate Sites and Personal Sites Daily Cache] completed."
}

Function GenerateSitesDailyCacheReport {
    $script:totalSites = @($script:sitesData).Count
    $script:totalDeletedSites = @($script:deletedSitesData).Count
    $script:totalPersonalSites = @($script:personalSitesData).Count       
    #$script:totalDeletedPersonalSites = $script:deletedPersonalSitesData.Count

    LogWrite -Message "->Total Active Sites Retrieved: $($script:totalSites)"
    LogWrite -Message "->Total Deleted Sites Retrieved: $($script:totalDeletedSites)"
    LogWrite -Message "->Total Active Personal Sites Retrieved: $($script:totalPersonalSites)"
    #LogWrite -Message "->Total Deleted Personal Sites Retrieved: $($script:totalDeletedPersonalSites)"
    
    $msg = ""
    $msg += "<p>Total Active Sites Retrieved: $($script:totalSites)<br />"    
    $msg += "Total Deleted Sites Retrieved: $($script:totalDeletedSites)<br />"
    $msg += "Total Active Personal Sites Retrieved: $($script:totalPersonalSites)</p>"
    #$msg += "Total Deleted Personal Sites Retrieved: $($script:totalDeletedPersonalSites)</p>"

    return $msg    
}


Try {
    #log file path
    Set-LogFile -logFileName $logFileName
    $script:StartTimeSitesDailyCache = Get-Date -Format "yyyy-MM-dd HH:mm:ss"           
    LogWrite -Message "----------------------- [Populate Sites and Personal Sites Daily Cache] Execution Started -----------------------"

    #Sync Sites and Personal Sites to the Cache    
    SyncSitesDataToCache

    $script:EndTimeSitesDailyCache = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    #Generate Report and send email
    GenerateDailyCacheReport 
    LogWrite -Message "[Populate Sites and Personal Sites Daily Cache] Start Time: $($script:StartTimeSitesDailyCache)"
    LogWrite -Message "[Populate Sites and Personal Sites Daily Cache] End Time:   $($script:EndTimeSitesDailyCache)"
    LogWrite -Message  "----------------------- [Populate Sites and Personal Sites Daily Cache] Execution Ended ------------------------"    
    
}
Catch [Exception] {
    LogWrite -Level ERROR "-Unexpected Error: $_ "
}
Finally {
    LogWrite -Message  "----------------------- [Populate Sites and Personal Sites Daily Cache] Completed ------------------------"
}
