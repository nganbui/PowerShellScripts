$0 = $MyInvocation.MyCommand.Definition
$dp0 = [System.IO.Path]::GetDirectoryName($0)
$currentFolder = $dp0.Substring($dp0.LastIndexOf('\') + 1)
$script:RootDir = Resolve-Path "$dp0\.." 
$logFileName = $MyInvocation.MyCommand.Name.Substring(0, $MyInvocation.MyCommand.Name.IndexOf('.'))
$logFileName = "$currentFolder\$logFileName"
#Include core
."$script:RootDir\Common\Lib\LibCore.ps1"
#Include dependent functionality
."$script:RootDir\Common\Lib\LibO365.ps1"
."$script:RootDir\Common\Lib\LibCache.ps1"
."$script:RootDir\Common\Lib\LibPowerBIWorkspace.ps1"
."$script:RootDir\Common\Lib\LibPowerBIWorkspaceDAO.ps1"

<#
      ===========================================================================
      .DESCRIPTION        
        
#>

Function SyncPowerBIworkspacesToCache {    
    Set-TenantVars
    Set-AzureAppVars
    Set-DataFile
    Set-DBVars
    
    LogWrite -Message "------------------------ Retrieving PowerBI workspaces -----------------------------------------"    
    GetPowerBIworkspaces
    LogWrite -Message "------------------------ Caching PowerBI workspaces --------------------------------------------"     
    CachePowerBIworkspaces 
    LogWrite -Message "------------------------ Populate PowerBI workspaces to DB --------------------------------------------"     
    UpdateSQLPowerBIworkspaces $script:ConnectionString $script:workspacesData
    LogWrite -Message "------------------------ Delete inactive PowerBI workspaces --------------------------------------------"
    $syncDate = Get-Date -format "yyyy-MM-dd"      
    DeleteInvalidPowerBIworkspaces $script:ConnectionString $syncDate
    
}

Function GenerateDailyCacheReport {
    LogWrite -Message "Sending Email Report: [Populate PowerBI Workspaces Daily Cache]"    
    $subject = "[SPO-DevOps] PowerBI Workspaces Daily Cache"    
    $body = "<p><b>Description:</b> This job will cache all PowerBI Workspaces data objects locally. These cache files are used to further sync data to the Database</p>"
    
    $body += "<p>Script Start time: $script:StartTimePowerBIworkspacesDailyCache<br />"
    $body += "Script End time: $script:EndTimePowerBIworkspacesDailyCache<br /><br />"    
    $body += GenerateM365GroupsDailyCacheReport          
    SendEmail -subject $subject -body $body
    LogWrite -Message "Sending Email Report: [Populate PowerBI Workspaces Daily Cache] completed."
}

Function GenerateM365GroupsDailyCacheReport {
    $totalWorkspaces = $script:workspacesData.Count
    $totalNewWorkspaces = ($script:workspacesData.Where( { $_.Type -eq 'Workspace' -and $_.State -eq "Active" })).Count
    $totalOldWorkspaces = ($script:workspacesData.Where( { $_.Type -eq 'Group' -and $_.State -eq "Active" })).Count
    $totalPersonalWorkspaces = ($script:workspacesData.Where( { $_.Type -eq 'PersonalGroup' -and $_.State -eq "Active" })).Count

    LogWrite -Message "Generating Email Report..."
    #
    LogWrite -Message "->Total Workspaces: $totalWorkspaces"
    LogWrite -Message "->Total Active Modern Workspaces: $totalNewWorkspaces"
    LogWrite -Message "->Total Active Classic Workspaces: $totalOldWorkspaces"
    LogWrite -Message "->Total Active Personal Workspaces: $totalPersonalWorkspaces"
    
    $msg = "<p>Total Workspaces: $totalWorkspaces<br />"
    $msg += " Total Active Modern Workspaces: $totalNewWorkspaces<br />"
    $msg += " Total Active Classic Workspaces: $totalOldWorkspaces<br />"
    $msg += " Total Active Personal Workspaces: $totalPersonalWorkspaces</p>"    
    return $msg
}

Try {
    #log file path
    Set-LogFile -logFileName $logFileName
    $script:StartTimePowerBIworkspacesDailyCache = Get-Date -Format "yyyy-MM-dd HH:mm:ss"           
    LogWrite -Message "----------------------- [Populate PowerBI workspaces Daily Cache] Execution Started -----------------------"

    #Sync Groups-Teams-Channel Objects to the Cache
    SyncPowerBIworkspacesToCache 

    $script:EndTimePowerBIworkspacesDailyCache = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    #Generate Report and send email
    GenerateDailyCacheReport 
    LogWrite -Message "[Populate Populate PowerBI Workspaces Daily Cache] Start Time: $($script:StartTimePowerBIworkspacesDailyCache)"
    LogWrite -Message "[Populate Populate PowerBI Workspaces Daily Cache] End Time:   $($script:EndTimePowerBIworkspacesDailyCache)"
    LogWrite -Message  "----------------------- [Populate Populate PowerBI Workspaces Daily Cache] Execution Ended ------------------------"    
    
}
Catch [Exception] {
    LogWrite -Level ERROR "-Unexpected Error: $_ "
}
Finally {
    LogWrite -Message  "----------------------- [Populate PowerBI workspaces Daily Cache] Completed ------------------------"
}
