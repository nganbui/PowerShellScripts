#region SharePoint Online Operations
Function EnableExternalSharing {
    [cmdletBinding()]
    param(
        [parameter(Mandatory = $true)]        
        [ValidateNotNullOrEmpty()]#either single or multiple urls
        $Url,
        [parameter(Mandatory = $false)]        
        [ValidateNotNullOrEmpty()]
        $TenantContext,
        [parameter(Mandatory = $false)]                
        [ValidateSet("Disabled", "ExistingExternalUserSharingOnly", "ExternalUserAndGuestSharing","ExternalUserSharingOnly")]
        [String]$SharingCapability="ExternalUserAndGuestSharing",
        [parameter(Mandatory = $false)]        
        [String]$Domains,
        [parameter(Mandatory = $false)]
        [ValidateSet("AllowList", "BlockList", "None")]
        [String]$DomainRestrictionMode="AllowList",
        $cb
    )
    try {        
        LogWrite -Message "PnPConnection - Connecting to SharePoint Online..."        
        if ($TenantContext -eq $null){
            $TenantContext = ConnectPnpOnline -Url $Url -Credential $script:o365AdminCredential        
        }
        LogWrite -Message "PnPConnection - SharePoint Online is now connected."        
    }
    catch {    
        LogWrite -Level ERROR -Message "PnPConnection - Unable to connect Sharepoint Online Session"
        LogWrite -Level ERROR -Message "PnPConnection - $($_.Exception)"
        exit
    }
    try {
        $requests = @()
        if($Url -isnot [array]){
           $requests.Add($Url)
        }       
        LogWrite -Message "Processing External Sharing Requests ..."       
        foreach ($request in $requests) {
            $site = Get-PnPTenantSite -Url $request -Connection $TenantContext -ErrorAction Ignore
            if ($site.Status -eq "Active") {
                LogWrite -Message "Processing enable External sharing: $($request)..."                
                Set-PnPTenantSite -Url $request -Connection $TenantContext -SharingCapability $SharingCapability -SharingAllowedDomainList $Domains -SharingDomainRestrictionMode $DomainRestrictionMode
                if ($cb -ne $null){
                    &$cb
                }
                LogWrite -Message "Enable External sharing: $($request) completed." 
            }
        }
        LogWrite -Message "Processing External Sharing Requests completed."       
    }
    catch {
        LogWrite -Level ERROR -Message "An error occured $($_.Exception)"        
    }
    finally {        
        LogWrite -Message "PnPConnection - Disconnecting from SharePoint Online..."    
        DisconnectPnpOnline  
        LogWrite -Message "PnPConnection - The Sharepoint Online Session is now closed."    
    }
}

Function UpdateStorageQuota {
    [cmdletBinding()]
    param(
        [parameter(Mandatory = $true)]        
        [ValidateNotNullOrEmpty()]
        [String]$Url,
        [parameter(Mandatory = $true)]        
        [ValidateNotNullOrEmpty()]
        $Quota
    )
    try {        
        LogWrite -Message "Connecting to SharePoint Online..."
        ConnectSPOnline -Url $script:SPOAdminCenterURL -Credential $script:o365AdminCredential
        LogWrite -Message "SharePoint Online Administration Center is now connected."        
    }
    catch {    
        LogWrite -Level ERROR -Message "Unable to connect Sharepoint Online Session"
        LogWrite -Level ERROR -Message "$($_.Exception)"
        exit
    }
    try {
        LogWrite -Message "Updating quota for SPO Site $($Url)..."       
        $quotaValue = $Quota * 1024 
        Set-SPOSite -Identity $Url -StorageQuota $quotaValue
        LogWrite -Message "Updating quota for SPO Site $($Url) completed."
    }
    catch {
        LogWrite -Level ERROR -Message "An error occured $($_.Exception)"        
    }
    finally {        
        LogWrite -Message "Disconnecting from SharePoint Online..."    
        DisconnectSPOnline  
        LogWrite -Message "The Sharepoint Online Session is now closed."    
    }          
     
}

Function UpdateExternalSharing {
    [CmdletBinding()]
    param(
        [parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()] 
        $SiteUrl,
        [parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()] 
        $SharingCapability
    )  
    try {
        $ret = @{}        
        # try getting the site                              
        $site = Get-PnPTenantSite -Url $siteUrl
                
        if ($site.Status -eq "Active" -and $site.LockState -eq 'Unlock') {
            $null = Set-PnPTenantSite -Url $SiteUrl -SharingCapability $SharingCapability
            $ret["Status"] = $script:Completed 
            $ret["Message"] = "Update external sharing successful: $siteUrl - $SharingCapability" 
        }
        else {            
            $ret["Status"] = $script:Pending 
            $ret["Message"] = "Site does not exist or is not available (status = $($site.Status)) - $siteUrl"             
        }              
    }
    catch {
        $ErrorMessage = $_.Exception.Message        
        Write-Verbose -Message "Unable to update external sharing for the site: $($ErrorMessage)" -Verbose  
        Write-Verbose -Message "$($_.Exception)" -Verbose        
        $ret["Status"] = $script:Pending 
        $ret["Message"] = $ErrorMessage
    }    
    return $ret
}

Function SiteCustomization {
    [CmdletBinding()]
    param(
        [parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()] 
        $SiteUrl,
        [parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()] 
        $NoScriptSite = $false
    )  
    try {
        $ret = @{}        
        # try getting the site                              
        $site = Get-PnPTenantSite -Url $siteUrl
                
        if ($site.Status -eq "Active" -and $site.LockState -eq 'Unlock') {
            Set-PnPTenantSite -Url $SiteUrl -DenyAddAndCustomizePages:$NoScriptSite -Wait
            #Set-PnPTenantSite -Url $SiteUrl -NoScriptSite:$NoScriptSite -Wait
            $ret["Status"] = $script:Completed 
            $ret["Message"] = "Update enable/disable custom script for SPO successful: $siteUrl" 
        }
        else {            
            $ret["Status"] = $script:Pending 
            $ret["Message"] = "Site does not exist or is not available (status = $($site.Status)) - $siteUrl"             
        }              
    }
    catch {
        $ErrorMessage = $_.Exception.Message        
        Write-Verbose -Message "Unable to enable/disable custom script for the site: $($ErrorMessage)" -Verbose  
        Write-Verbose -Message "$($_.Exception)" -Verbose        
        $ret["Status"] = $script:Pending 
        $ret["Message"] = $ErrorMessage
    }    
    return $ret

}

Function RegisterHubSite {
    [CmdletBinding()]
    param(
        [parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()] 
        $SiteUrl,
        [parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()] 
        $HubName,
        [parameter(Mandatory = $false)]        
        $Owners
    )  
    try {
        $ret = @{}        
        # try getting the site                              
        $site = Get-PnPTenantSite -Url $siteUrl
                
        if ($site.Status -eq "Active" -and $site.LockState -eq 'Unlock') {
            LogWrite -Level INFO -Message "Registers a site as a hubsite: $($siteURL)."                    
            Register-PnPHubSite -Site $SiteUrl
            Set-PnPHubSite -Identity $SiteUrl -Title $HubName
            if ($Owners){
                Grant-PnPHubSiteRights -Identity $SiteUrl -Principals $owners -Rights Join
            }
            $ret["Status"] = $script:Completed 
            $ret["Message"] = "Registers a site as a hubsite successful: $siteUrl" 
        }
        else {            
            $ret["Status"] = $script:Pending 
            $ret["Message"] = "Site does not exist or is not available (status = $($site.Status)) - $siteUrl"             
        }              
    }
    catch {
        $ErrorMessage = $_.Exception.Message        
        Write-Verbose -Message "Unable to register a site as a hubsite: $($ErrorMessage)" -Verbose
        Write-Verbose -Message "$($_.Exception)" -Verbose        
        $ret["Status"] = $script:Pending 
        $ret["Message"] = $ErrorMessage
    }    
    return $ret

}

Function AddRemoveAppCatalog {
    [CmdletBinding()]
    param(
        [parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()] 
        $SiteUrl,
        [parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()] 
        $AppCatalog = "Disabled"
    )
    <#
    .DESCRIPTION
        https://docs.microsoft.com/en-us/sharepoint/dev/general-development/site-collection-app-catalog

        Add-PnPSiteCollectionAppCatalog
            Apps for SharePoint library will be added to the site collection where to deploy SharePoint add-ins and SharePoint Framework solutions.
        Remove-PnPSiteCollectionAppCatalog
            The Apps for SharePoint library will be still visible in the site collection, but  will not be able to deploy or use any solutions deployed in it.
        To list all site collections in the tenant that have the site collection app catalog enabled, use the URL
            https://citspdev.sharepoint.com/sites/AppCatalog/Lists/SiteCollectionAppCatalogs/AllItems.aspx
    #>  
    try {
        $ret = @{}        
        # try getting the site 
        #$siteContext = ConnectPnpOnlineOAuth -TenantId $script:TenantId -ClientId $script:appIdOperationSupport -Thumbprint $script:appThumbprintOperationSupport -Url $SiteUrl                             
        $site = Get-PnPTenantSite -Url $siteUrl
                
        if ($site.Status -eq "Active" -and $site.LockState -eq 'Unlock') {
            Write-Verbose -Message "Enabled/Disabled App catalog for the site: $($siteURL)."                    
            if ($AppCatalog -eq "Disabled"){
                Remove-PnPSiteCollectionAppCatalog -Site $SiteUrl
                Write-Verbose -Message "Disabled App Catalog."
            }
            elseif ($AppCatalog -eq "Enabled"){
                Add-PnPSiteCollectionAppCatalog -Site $SiteUrl
                Write-Verbose -Message "Enabled App Catalog."
            }
            
            $ret["Status"] = $script:Completed 
            $ret["Message"] = "Enabled/Disabled App catalog for the site successful: $siteUrl" 
        }
        else {            
            $ret["Status"] = $script:Pending 
            $ret["Message"] = "Site does not exist or is not available (status = $($site.Status)) - $siteUrl"             
        }              
    }
    catch {
        $ErrorMessage = $_.Exception.Message        
        Write-Verbose -Message "Unable to Enabled/Disabled App catalog for the site: $($ErrorMessage)" -Verbose
        Write-Verbose -Message "$($_.Exception)" -Verbose        
        $ret["Status"] = $script:Pending 
        $ret["Message"] = $ErrorMessage
    }
    #finally{
    #    DisconnectPnpOnlineOAuth -Context $siteContext
    #}     
    return $ret

}

Function ConnectM365Group {
    [CmdletBinding()]
    param(
        [parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()] 
        $SiteUrl,
        [parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()] 
        $Owner        
    )
    <#
    .DESCRIPTION
        Add-PnPMicrosoft365GroupToSite: Application permission is not supported to perform this action.
    #>   
    try {
        $ret = @{}

        $siteContext = ConnectPnpOnlineOAuth -TenantId $script:TenantId -ClientId $script:appIdOperationSupport -Thumbprint $script:appThumbprintOperationSupport -Url $SiteUrl

        $publishingSiteFeature = Get-PnPFeature -Identity "F6924D36-2FA8-4F0B-B16D-06B7250180FA" -Scope Site -Connection $siteContext
        $publishingWebFeature = Get-PnPFeature -Identity "94C94CA6-B32F-4DA9-A9E3-1F3D343D7ECB" -Scope Web -Connection $siteContext

        if (($publishingSiteFeature.DefinitionId -ne $null) -or ($publishingWebFeature.DefinitionId -ne $null)) {
            throw "Publishing feature enabled...can't group connect this site"
        }
        
        LogWrite "Enabling modern page feature, disabling modern list UI blocking features"
        # Enable modern page feature
        Enable-PnPFeature -Identity "B6917CB1-93A0-4B97-A84D-7CF49975D4EC" -Scope Web -Force -Connection $siteContext
        # Disable the modern list site level blocking feature
        Disable-PnPFeature -Identity "E3540C7D-6BEA-403C-A224-1A12EAFEE4C4" -Scope Site -Force -Connection $siteContext
        # Disable the modern list web level blocking feature
        Disable-PnPFeature -Identity "52E14B6F-B1BB-4969-B89B-C4FAA56745EF" -Scope Web -Force -Connection $siteContext
                
        # try getting the site                              
        $site = Get-PnPTenantSite -Url $siteUrl
                
        if ($site.Status -eq "Active" -and $site.LockState -eq 'Unlock') {
            Write-Verbose -Message "Connect M365 Group to the site: $siteUrl..."

            $groupAlias =$SiteUrl.substring($SiteUrl.lastIndexof("/")+1)
            $groupAlias =$groupAlias.Replace(" ","")
            if ($groupAlias.length -gt 64){
                $groupAlias = $groupAlias.substring(0,64)
            }

            $aliasIsUsed = Test-PnPMicrosoft365GroupAliasIsUsed -Alias $groupAlias
            if ($aliasIsUsed) {
                $message = "$siteUrl : Alias [$groupAlias] is already in use"
                LogWrite -Level ERROR -Message $message                
                $ret["Status"] = $script:Pending 
                $ret["Message"] = "Alias [$groupAlias] is already in use." 
            }
            else {
                LogWrite -Level INFO -Message "Connect M365 Group to the site: $siteUrl..." 
                #$siteContext = ConnectPnpOnlineOAuth ConnectPnpOnlineOAuth -TenantId $script:TenantId -ClientId $script:appIdOperationSupport -Thumbprint $script:appThumbprintOperationSupport -Url $SiteUrl
                Add-PnPMicrosoft365GroupToSite -Url $siteUrl -Alias $groupAlias -DisplayName $groupAlias -IsPublic:$false -KeepOldHomePage:$false -Owners $Owner -Connection $siteContext
                $ret["Status"] = $script:Completed 
                $ret["Message"] = "Connect M365 Group to the site successful: $siteUrl"
            } 
            
             
        }
        else {            
            $ret["Status"] = $script:Pending 
            $ret["Message"] = "Site does not exist or is not available (status = $($site.Status)) - $siteUrl"             
        }              
    }
    catch {
        $ErrorMessage = $_.Exception.Message        
        Write-Verbose -Message "Unable to Connect M365 Group to the site: $($ErrorMessage)" -Verbose
        Write-Verbose -Message "$($_.Exception)" -Verbose        
        $ret["Status"] = $script:Pending 
        $ret["Message"] = $ErrorMessage
    }
    finally{
        DisconnectPnpOnlineOAuth -Context $siteContext
    }    
    return $ret

}

Function AddSiteAdmin {
    [CmdletBinding()]
    param(
        [parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()] 
        $SiteUrl,
        [parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()] 
        $Owners
    )  
    try {
        $ret = @{}        
        # try getting the site                              
        $site = Get-PnPTenantSite -Url $siteUrl
        
        if ($site.Status -eq "Active" -and $site.LockState -eq 'Unlock') {
            #Set-PnPTenantSite -Identity $SiteUrl -Owners $Owners                                                  
            $siteContext = ConnectPnpOnlineOAuth -TenantId $script:TenantId -ClientId $script:appIdOperationSupport -Thumbprint $script:appThumbprintOperationSupport -Url $siteUrl           
            $dt = GetSigninNameByEmail -Email $Owners -connectionString $script:ConnectionString
            $SiteOwnerId= $dt["UserId"].Guid
            $SiteOwnerUPN= $dt["SigninName"].Trim()
            Add-PnPSiteCollectionAdmin -Owners $SiteOwnerUPN -Connection $siteContext
            $ret["Status"] = $script:Completed 
            $ret["Message"] = "Added $Owners as site admin: $siteUrl" 
        }
        else {            
            $ret["Status"] = $script:Pending 
            $ret["Message"] = "Site does not exist or is not available (status = $($site.Status)) - $siteUrl"             
        }              
    }
    catch {
        $ErrorMessage = $_.Exception.Message        
        Write-Verbose -Message "Unable to add site admin: $($ErrorMessage)" -Verbose  
        Write-Verbose -Message "$($_.Exception)" -Verbose        
        $ret["Status"] = $script:Pending 
        $ret["Message"] = $ErrorMessage
        if ($siteContext) {
            DisconnectPnpOnlineOAuth -Context $siteContext
        }
    }    
    return $ret
}
#endregion

#region MS Teams Operations
Function EditTeam {
    [CmdletBinding()]
    param(
        [parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()] 
        $GroupId,
        [parameter(Mandatory = $false)]        
        $DisplayName,
        [parameter(Mandatory = $false)]        
        $Description,
        [parameter(Mandatory = $false)]
        [ValidateSet('Private', 'Public')]
        [string]$Privacy
    )  
    try {
        $ret = @{}        
        $team = Get-Team -GroupId $GroupId
        if ($team -eq $null) {
            $ret["Status"] = $script:Pending
            $ret["Message"] = "Team not found."                    
        }
        else {           
            $ret["Status"] = $script:Completed

            if ($DisplayName){
                $null = Set-Team -GroupId $GroupId -DisplayName $DisplayName
            }
            if ($Description){
                $null = Set-Team -GroupId $GroupId -Description $Description
            }
            if ($Privacy){
                $null = Set-Team -GroupId $GroupId -Visibility $Privacy
            }
            $ret["Message"] = "Completion Edit Team."
        }             
    }
    catch {
        $ErrorMessage = $_.Exception.Message        
        Write-Verbose -Message "Unable to edit the team: $($ErrorMessage)" -Verbose  
        Write-Verbose -Message "$($_.Exception)" -Verbose        
        $ret["Status"] = $script:Pending
        $ret["Message"] = $ErrorMessage
    }    
    return $ret
}

Function RenameTeamDisplayName {
    [CmdletBinding()]
    param(
        [parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()] 
        $GroupId,
        [parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()] 
        $DisplayName
    )  
    try {
        $ret = @{}        
        $team = Get-Team -GroupId $GroupId
        if ($team -eq $null) {
            $ret["Status"] = $script:Pending
            $ret["Message"] = "Team not found."                    
        }
        else {           
            $ret["Status"] = $script:Completed 
            if ($team.DisplayName -ne $DisplayName) {                    
                $null = Set-Team -GroupId $GroupId -DisplayName $DisplayName
                $ret["Message"] = "Team renamed successfully. Team new name: $DisplayName."                
            }  
            else {
                $ret["Message"] = "Current and new Team names are same. No change will be made."                        
            }
        }             
    }
    catch {
        $ErrorMessage = $_.Exception.Message        
        Write-Verbose -Message "Unable to rename the team: $($ErrorMessage)" -Verbose  
        Write-Verbose -Message "$($_.Exception)" -Verbose        
        $ret["Status"] = $script:Pending
        $ret["Message"] = $ErrorMessage
    }    
    return $ret
}
#endregion

#region O365 Groups Operations
Function LoginNameToUPN {
    param([string] $loginName)
    return $loginName.Replace("i:0#.f|membership|", "")
}

Function AddToOffice365GroupOwnersMembers {
    [cmdletBinding()]
    param($groupUserUpn, $groupId, [bool] $Owners)    
    $retryCount = 5
    $retryAttempts = 0
    $backOffInterval = 2

    Write-Verbose -Message "Attempting to add $groupUserUpn to group $groupId" -Verbose 

    while ($retryAttempts -le $retryCount) {
        try {
            if ($Owners) {
                $azureUserId = Get-AzureADUser -ObjectId $groupUserUpn            
                Add-AzureADGroupOwner -ObjectId $groupId -RefObjectId $azureUserId.ObjectId  
                Write-Verbose -Message "User $groupUserUpn added as group owner" -Verbose  
            }
            else {
                $azureUserId = Get-AzureADUser -ObjectId $groupUserUpn           
                Add-AzureADGroupMember -ObjectId $groupId -RefObjectId $azureUserId.ObjectId    
                Write-Verbose -Message "User $groupUserUpn added as group member" -Verbose  
            }
            
            $retryAttempts = $retryCount + 1;
        }
        catch {
            if ($retryAttempts -lt $retryCount) {
                $retryAttempts = $retryAttempts + 1        
                Write-Host "Retry attempt number: $retryAttempts. Sleeping for $backOffInterval seconds..."
                Start-Sleep $backOffInterval
                $backOffInterval = $backOffInterval * 2
            }
            else {
                throw
            }
        }
    }    
}
#endregion