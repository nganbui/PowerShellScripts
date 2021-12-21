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
."$script:RootDir\Common\Lib\GraphAPILibSPOSitesDAO.ps1"

Function RetrieveSPOSitesWithExternalSharing {        
    try {       
        $script:sitesDataWithExternalSharingEnabled = @()
        if ($script:TenantConext -eq $null){
            $script:TenantConext = ConnectPnpOnlineOAuth -TenantId $script:TenantId -ClientId $script:appIdOperationSupport -Thumbprint $script:appThumbprintOperationSupport -Url $script:SPOAdminCenterURL
        }
        #All active sites
        $sitesData = Get-PnPTenantSite | ? { $_.Status -eq 'Active' -and $_.LockState -eq 'Unlock'} | Select Url,SharingCapability       
        #All active sites with ExternalUserAndGuestSharing        
        $script:sitesDataWithExternalSharingEnabled = $sitesData | ? { $_.SharingCapability  -match 'ExternalUser'} | Select Url         
        $script:totalSites = [int](@($sitesData).Count)
        $script:totalSitesWithExternalSharing = [int](@($script:sitesDataWithExternalSharingEnabled).Count)    
        LogWrite -Message "Sites with External Sharing Enabled $($script:totalSitesWithExternalSharing)"
    }
    catch {
        LogWrite -Level ERROR "Error in the script: $($_)"
    }      
}

Function Disable_ExternalSharing {    
    Set-DBVars
    $script:ExceptionUrls = (GetSitesExternalSharing $script:connectionString $true).Url    
    $script:totalSitesWithExternalSharingSuccessfullyDisabled = 0
    $script:totalSitesWithExternalSharingFailedToDisable = 0
    $script:totalSitesWithExternalSharingSkipped = ($script:sitesDataWithExternalSharingEnabled | ? { $_.Url -in $script:ExceptionUrls }).Url.Count
    $disabledExternalSharing = $script:sitesDataWithExternalSharingEnabled | ? { $_.Url -notin $script:ExceptionUrls }
    $countDisabledExternalSharing = $disabledExternalSharing.Url.count
    
    if ($countDisabledExternalSharing -gt 0){
        LogWrite -Message "Connecting to SharePoint Admin Center '$($script:SPOAdminCenterURL)'..."
        if ($script:TenantConext -eq $null){
            $script:TenantConext = ConnectPnpOnlineOAuth -TenantId $script:TenantId -ClientId $script:appIdOperationSupport -Thumbprint $script:appThumbprintOperationSupport -Url $script:SPOAdminCenterURL
        }
        LogWrite -Message "SharePoint Admin Center '$($script:SPOAdminCenterURL)' is now connected."   
        LogWrite -Message "Processing disable external sharing for the sites are not in exception external list."
        #LogWrite -Message "$($disabledExternalSharing.Url)"
        foreach($site in $disabledExternalSharing)
        {  
            try
            {
                Set-PnPTenantSite -Url $site.Url -SharingCapability Disabled
                $script:totalSitesWithExternalSharingSuccessfullyDisabled++
                LogWrite -Message "Successfully disabled external sharing for site $($site.Url)"
            }
            catch
            {   
                $script:totalSitesWithExternalSharingFailedToDisable++             
                LogWrite -Message -Level ERROR "Unable to disable external sharing for site $($site.URL). Error info $($_)"
            }
        }
    }
}

Function SendEmailReport {
    if ($script:totalSitesWithExternalSharingSuccessfullyDisabled -lt 1)
    {
        return;
    }

    $script:EndTime = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

    LogWrite -Message "Start time: $($script:StartTime)"
    LogWrite -Message "End time: $($script:EndTime)"

    
    LogWrite -Message "Generating Email Report..."
    #
    LogWrite -Message "->Total Sites Found: $($script:totalSites)"
    LogWrite -Message "->Total Sites with External Sharing Enabled: $($script:totalSitesWithExternalSharing)"
    LogWrite -Message "->Total Sites with External Sharing Skipped: $($script:totalSitesWithExternalSharingSkipped)"
    #
    LogWrite -Message "->Total Sites where External Sharing Disabled in this Run: $($script:totalSitesWithExternalSharingSuccessfullyDisabled)"
    LogWrite -Message "->Total Sites where External Sharing Failed to disable in this Run: $($script:totalSitesWithExternalSharingFailedToDisable)"
    
    #---- Send Email ----    
    $subject = "[SPO-DevOps] External Sharing Management"

    $body=""
    $body+="<p><b>Description:</b> This job will disable all the External Sharing of the sites</p>"
    $body+="<p>Total Sites: $($script:totalSites)<br />"
    $body+="Total Sites with External Sharing Enabled: $($script:totalSitesWithExternalSharing)<br />"
    $body+="Total Sites with External Sharing Skipped: $($script:totalSitesWithExternalSharingSkipped)<br />"

    $body+="Total Sites where External Sharing Disabled in this Run: $($script:totalSitesWithExternalSharingSuccessfullyDisabled)<br />"
    $body+="Total Sites where External Sharing Failed to disable in this Run: $($script:totalSitesWithExternalSharingFailedToDisable)<br />"
     
    $body+="<p>Script Start time: $($script:StartTime)<br />"
    $body+="Script End time: $($script:EndTime)</p><br />"   
   
    SendEmail -subject $subject -body $body
}

#
#==========================================
#---main script starts here ---
#==========================================
Try {
    #log file path
    Set-LogFile -logFileName $logFileName
    $script:StartTime = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    LogWrite -Message "----------------------- [Disable External Sharing] Execution Started --------------------------"
    Set-TenantVars
    Set-AzureAppVars
    $script:TenantConext = ConnectPnpOnlineOAuth -TenantId $script:TenantId -ClientId $script:appIdOperationSupport -Thumbprint $script:appThumbprintOperationSupport -Url $script:SPOAdminCenterURL

    #Get Active sites and extract the ExternalSharing enabled sites
    RetrieveSPOSitesWithExternalSharing

    #Disable External Sharing for the sites
    Disable_ExternalSharing

    Set-EmailVars
    SendEmailReport


    LogWrite -Message "----------------------- [Disable External Sharing] Execution Ended --------------------------"
}
Catch [Exception] {
    LogWrite -Level ERROR "Error in the script: $($_.Exception.Message)"
}
Finally {    
    LogWrite -Level INFO -Message " Disconnect SharePoint Admin Center." 
    DisconnectPnpOnlineOAuth -Context $script:TenantContext     
}