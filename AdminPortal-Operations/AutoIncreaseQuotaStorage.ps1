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
."$script:RootDir\Common\Lib\LibCache.ps1"

<#
      ===========================================================================
      .DESCRIPTION
        Processing increasing quota storage for OD and SPO when storage used > 90%
#>

Function GenerateReportIncreaseQuota {
    param($attachments)
    LogWrite -Message "Sending Email Report: [Auto-Increase OneDrive and SPO Storage Quota]"    
    $subject = "[M365 DevOps] Auto-Increase OneDrive and SPO Storage Quota"
    $body = ""
    if ($attachments.Length -gt 0) { 
        $body = "<p><b>Description:</b> This job auto-increase sites and personal site when storage used > 90% of storage quota<br />"
        $body += "Please review and address any issues from the attached files if needed.</p>"    
    }
    SendEmail -subject $subject -body $body -Attachements $attachments #-To "ngan.bui@nih.gov"
    LogWrite -Message "Sending Email Report: [Auto-Increase OneDrive and SPO Storage Quota] completed."
}

Try {    
    #log file path
    Set-LogFile -logFileName $logFileName
    Set-DataFile    
    Set-TenantVars
    Set-AzureAppVars
    Set-MiscVars
    
    $startTimeDailyCache = Get-Date -Format "yyyy-MM-dd HH:mm:ss"           
    LogWrite -Message "----------------------- [AutoIncreaseQuota] Execution Started -----------------------"
    
    $SPOAdminConnection = ConnectPnpOnlineOAuth -TenantId $script:TenantId -ClientId $script:appIdOperationSupport -Thumbprint $script:appThumbprintOperationSupport -Url $script:SPOAdminCenterURL
    LogWrite -Message "SharePoint Online Administration Center is now connected."

    $od = Get-PnPTenantSite -IncludeOneDriveSites -Filter "Url -like '-my.sharepoint.com/personal/' -and Status -eq 'Active' -and LockState -eq 'Unlock' -and StorageQuota -gt 0" | Select Owner, Title, URL,Template, StorageQuota, StorageUsageCurrent,StorageQuotaWarningLevel
    $spo = Get-PnPTenantSite -Filter "Url -notlike '-my.sharepoint.com' -and Status -eq 'Active' -and LockState -eq 'Unlock' -and StorageQuota -gt 0" | select Owner,Title,URL,Template,StorageQuota,StorageUsageCurrent,StorageQuotaWarningLevel    

    $spo = @($spo | ? { [Math]::Round(($_.StorageUsageCurrent/$_.StorageQuota * 100),2) -gt $script:ThresholdPercent -and [Math]::Round($_.StorageQuota/1024,2) -lt $script:MaxStorageQuota})
    $od = @($od | ? { [Math]::Round(($_.StorageUsageCurrent/$_.StorageQuota * 100),2) -gt $script:ThresholdPercent -and [Math]::Round($_.StorageQuota/1024,2) -lt $script:MaxStorageQuota})    

    $QuotaMB = 1048576 #1 TB = 1,048,576 MB   
    $Report = [System.Collections.Generic.List[Object]]::new()

    if ($od.Count -gt 0){
        LogWrite -Message "[OD: Processing increasing storage quota"        
        #$odExceedMaxQuota = $od | ? {[Math]::Round($s.StorageQuota/1024,2) -ge $script:MaxStorageQuota}
        #$od = $od | ? {[Math]::Round($s.StorageQuota/1024,2) -lt $script:MaxStorageQuota}

        ForEach ($s in $od) {
            $NewStorageQuota = ($s.StorageQuota) + $QuotaMB # increase 1TB up to 20TB
            $NewStorageWarningLevel = [Math]::Round((($NewStorageQuota) * ($script:ThresholdPercent)) / 100)
            #$result = Set-PnPTenantSite -Url $s.Url -StorageQuota $NewStorageQuota -StorageQuotaWarningLevel $NewStorageWarningLevel

            $ReportLine   = [PSCustomObject]@{
                SiteType       = "PersonalSite"
                Template     = $s.Template
                Owner       = $s.Title
                Email       = $s.Owner
                URL         = $s.URL                
                StorageUsedGB     = [Math]::Round($s.StorageUsageCurrent/1024,2) 
                QuotaGB     = [Math]::Round($s.StorageQuota/1024,2) 
                NewQuotaGB     = [Math]::Round($NewStorageQuota/1024,2) }
            $Report.Add($ReportLine)
        }
    }
    if ($spo.Count -gt 0){
        LogWrite -Message "[SPO]: Processing increasing storage quota"
        ForEach ($s in $spo) {
            $NewStorageQuota = ($s.StorageQuota) + $QuotaMB # increase 1TB up to 20TB
            $NewStorageWarningLevel = [Math]::Round((($NewStorageQuota) * ($script:ThresholdPercent)) / 100)
            #$result = Set-PnPTenantSite -Url $s.Url -StorageQuota $NewStorageQuota -StorageQuotaWarningLevel $NewStorageWarningLevel

            $ReportLine   = [PSCustomObject]@{
                SiteType       = "Site"
                Template     = $s.Template
                Owner       = $s.Title
                Email       = $s.Owner
                URL         = $s.URL                
                StorageUsedGB     = [Math]::Round($s.StorageUsageCurrent/1024,2) 
                QuotaGB     = [Math]::Round($s.StorageQuota/1024,2) 
                NewQuotaGB     = [Math]::Round($NewStorageQuota/1024,2) }
            $Report.Add($ReportLine)
        }
    }
    if ($Report -ne $null -and $Report.Count -gt 0) {
        $attachedFiles = @()
        LogWrite -Message "Export to csv and sending report email to SP Admins..."
        $logPath = "$($script:DirLog)"
        if (!(Test-Path $logPath)) { 
	        LogWrite -Message "Creating $logPath" 
            New-Item -ItemType "directory" -Path $logPath -Force
	    }
        LogWrite -Message "Generating Log files..." 
        $reportFile = "$logPath\ODandSPOIncreasedQuota.csv"
                
        $Report | Export-Csv $reportFile -Encoding ASCII -NoTypeInformation
        $attachedFiles += $reportFile
        GenerateReportIncreaseQuota -attachments $attachedFiles
     }
       
    $endTimeDailyCache = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    
    LogWrite -Message "[AutoIncreaseQuota] Start Time: $startTimeDailyCache)"
    LogWrite -Message "[AutoIncreaseQuota] End Time:   $endTimeDailyCache)"
    LogWrite -Message  "----------------------- [AutoIncreaseQuota] Execution Ended ------------------------"    
    #endregion
    
}
Catch [Exception] {
    LogWrite -Level ERROR "-Unexpected Error: $_ "
}
Finally {
    if($SPOAdminConnection){
        DisconnectPnpOnlineOAuth -Context $SPOAdminConnection
    }
    LogWrite -Message  "----------------------- [AutoIncreaseQuota] Completed ------------------------"
}
