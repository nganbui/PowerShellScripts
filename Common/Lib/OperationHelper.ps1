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

Function Update-ExternalSharing {
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
                
        if ($site.Status -eq "Active") {
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

#endregion

#region MS Teams Operations
Function Rename-TeamDisplayName {
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