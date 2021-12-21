#region First run: provision Team/SPO with basic settings and assign svc as an owner + update request status to "In Progress"
Function Provision-New {    
    param([Parameter(Mandatory = $true)] $Request)
    
    try {       
        $reqId = $Request["Id"]
        $reqTemplateId = $Request["TemplateId"]
        $reqObjectId = [string]$Request["ObjectId"]
        $siteUrl = $Request["SiteUrl"]
        
        <#$template = "Team"
        if ($reqTemplateId -ne $template){
            $template = "SPO"
        } 
        #>        
        switch ($reqTemplateId) {
            "Team" {                
                $provisionedSite = Provision-Team -Request $Request
            }
            "GROUP#0"{
                $provisionedSite = Provision-Group -Request $Request
            }
            "PowerBIWorkspace"{
                $provisionedSite = Provision-PowerBIWorkspace -Request $Request
            }
            default {
                $provisionedSite = Provision-SPOSite -Request $Request
            }
        }
        #team return group object
        #non-group return siteURL

        if ($null -eq $provisionedSite){
            LogWrite -Message " $($script:ProcessNew): The site/workspace has not been provisioned."            
            return
        }
        if ($reqTemplateId -ne "PowerBIWorkspace") {    
            LogWrite " $($script:ProcessNew): New Site Request has been provisioned successfully."
            #Update Status to "InProgress" and groupId if provision Team/Group in the Requests Table            
            if ($reqTemplateId -eq "Team" -or $reqTemplateId -eq "GROUP#0"){
                $reqObjectId = $provisionedSite 
                if ($provisionedSite.id){
                    $reqObjectId = $provisionedSite.id
                }            
                LogWrite " $($script:ProcessNew): MS Teams/M365 Group with Group Id: [$reqObjectId]"                
            } 
            UpdateProvisionRequest -RequestId $reqId -ReqStatusID $script:InProgress -ReqObjectId $reqObjectId  -connectionString $script:ConnectionString
            LogWrite " $($script:ProcessNew): Updated the Request Status [In Progress] for the site."
        } 
    }
    catch {
        LogWrite -Level ERROR "$($script:ProcessNew): Something went wrong during provision a new site/workspace request: [$siteUrl]. Error Info: $($_.Exception)"         
        #Mark as skip if there is something wrong for first run
        LogWrite -Message " $($script:ProcessNew): Mark as skip if there is something wrong for first run"
        UpdateProvisionRequest -RequestId $reqId -ReqStatusID $script:Submitted -ReqObjectId $reqObjectId  -ReqProcessFlag 1 -connectionString $script:ConnectionString 
        throw $_         
    }    
}
#endregion

#region Second run
Function Provision-InProgress {    
    param([Parameter(Mandatory = $true)] $Request)
    
    try {
        $reqId = $Request["Id"]
        $reqStatus = $Request["Status"]
        $reqObjectId = $Request["ObjectId"]
        $siteUrl = $Request["SiteUrl"]
        $DisplayName = $Request["DisplayName"]
        $SiteDescription = $Request["Description"]
        $addedOwner = $Request["OwnerId"]        
        $addedOwnerUPN = $Request["OwnerUPN"]
        $IncidentId = $Request["IncidentId"]
        $reqTemplateId = $Request["TemplateId"]
        $externalSharing = $Request["ExternalSharing"]

        LogWrite -Message " $($script:ProcessInProgress): Processing pending site request..."
        LogWrite -Message " $($script:ProcessInProgress): Connecting to the site..."
        $siteContext = ConnectPnpOnlineOAuth -TenantId $script:TenantId -ClientId $script:appIdOperationSupport -Thumbprint $script:appThumbprintOperationSupport -Url $siteUrl

        switch ($reqTemplateId) {
            "Team" {                
                #-Provision-InProgress
                #$script:TokenOperationSupport.Value["Expires_in"]
                #$script:TokenOperationSupport.Value["Ext_expires_in"]
                #$script:TokenOperationSupport | Get-JWTDetails
                #$script:TokenOperationSupport["Expires_in"]
                #$currentDateTimePlusTen = (Get-Date).AddMinutes(10)
                $groupId = (Get-PnPSite -Includes GroupId).GroupId.ToString()
                $script:TokenOperationSupport = Connect-GraphAPIWithCert -TenantId $script:TenantId -AppId $script:appIdOperationSupport -Thumbprint $script:appThumbprintOperationSupport
                
                <#if ($script:TokenOperationSupport) {
                    if (!($currentDateTimePlusTen -le $script:TokenOperationSupport["Expires_in"])) {                     
                        # get an accesstoken if current accesstoken is valid but expired
                        $script:TokenOperationSupport = Connect-GraphAPIWithCert -TenantId $script:TenantId -AppId $script:appIdOperationSupport -Thumbprint $script:appThumbprintOperationSupport
                    }        
                }
                else {
                    # get an accesstoken if accesstoken is $null
                    $script:TokenOperationSupport = Connect-GraphAPIWithCert -TenantId $script:TenantId -AppId $script:appIdOperationSupport -Thumbprint $script:appThumbprintOperationSupport
                }#>   

                #if ($null -eq $script:TokenOperationSupport -or $script:TokenOperationSupport.Values["Expires_in"]){
                #    $script:TokenOperationSupport = Connect-GraphAPIWithCert -TenantId $script:TenantId -AppId $script:appIdOperationSupport -Thumbprint $script:appThumbprintOperationSupport
                #}                
                LogWrite -Message " $($script:ProcessInProgress): Adding an owner ($addedOwnerUPN) to  Microsoft Teams [$DisplayName]"
                Add-NIHTeamMember -AuthToken $script:TokenOperationSupport -Group $groupId -Member $addedOwner -AsOwner
                LogWrite -Message " $($script:ProcessInProgress): Removing ($($script:CloudSvcForProvision)) from  Microsoft Teams [$DisplayName]"
                Remove-NIHO365GroupMember -AuthToken $script:TokenOperationSupport -Group $groupId -Member $script:CloudSvcForProvision -AsOwner
                
                LogWrite -Message " $($script:ProcessInProgress): Updating HideFromOutlookClients and HideFromAddressLists for Teams - By default Teams hide in Outlook and GAL"
                $result = Update-NIHGroupSettings -AuthToken $script:TokenOperationSupport -Id $groupId -HideFromOutlookClients $true -HideFromAddressLists $true
                if (!$result[0]){
                    LogWrite -Level ERROR -Message " $($script:ProcessInProgress): There is an error occurred with updating GAL/Outlook: $result"
                }

            }
            "GROUP#0" {
                $groupId = (Get-PnPSite -Includes GroupId).GroupId.ToString()
                $script:TokenOperationSupport = Connect-GraphAPIWithCert -TenantId $script:TenantId -AppId $script:appIdOperationSupport -Thumbprint $script:appThumbprintOperationSupport

                LogWrite -Message " $($script:ProcessInProgress): Adding an owner ($addedOwnerUPN) to  M365 Group [$DisplayName]"
                Add-NIHO365GroupMember -AuthToken $script:TokenOperationSupport -Group $groupId -Member $addedOwner -AsOwner
                LogWrite -Message " $($script:ProcessInProgress): Removing ($($script:CloudSvcForProvision)) from  M365 Group [$DisplayName]"
                Remove-NIHO365GroupMember -AuthToken $script:TokenOperationSupport -Group $groupId -Member $script:CloudSvcForProvision -AsOwner
                
                LogWrite -Message " $($script:ProcessInProgress): Updating HideFromOutlookClients and HideFromAddressLists for M365 Group - By default M365 Group shows in Outlook and GAL"
                $result = Update-NIHGroupSettings -AuthToken $script:TokenOperationSupport -Id $groupId -HideFromOutlookClients $false -HideFromAddressLists $false                
                if (!$result[0]){
                    LogWrite -Level ERROR -Message " $($script:ProcessInProgress): There is an error occurred with updating GAL/Outlook: $result"
                }                

            }
            "SITEPAGEPUBLISHING#0" { #Communication Site 
                AddSiteAdmins -Admins $addedOwnerUPN -SiteUrl $siteUrl -SiteContext $siteContext 
                AddSiteOwners -Owner $addedOwnerUPN -SiteUrl $siteUrl -SiteContext $siteContext                                
            }
            "STS#3" { #Team Site without group
                LogWrite -Message " $($script:ProcessInProgress): Disabling Site customization..."
                Set-PnPTenantSite -Url $siteUrl -DenyAddAndCustomizePages:$true -Wait
                #Add Site Description to the Team Site                
                if ($SiteDescription -ne $null) {
                    UpdateSiteDescription $SiteDescription -SiteUrl $siteUrl -SiteContext $siteContext          
                }
                AddSiteAdmins -Admins $addedOwnerUPN -SiteUrl $siteUrl -SiteContext $siteContext
                AddSiteOwners  -Owner $addedOwnerUPN -SiteUrl $siteUrl -SiteContext $siteContext
                
            }
            "STS#0" { #Classic Team Site
                LogWrite -Message " $($script:ProcessInProgress): Disabling Site customization..."
                Set-PnPTenantSite -Url $siteUrl -DenyAddAndCustomizePages:$true -Wait
                #Add Site Description to the Classic Team Site
                if ($SiteDescription -ne $null) {
                    UpdateSiteDescription $SiteDescription -SiteUrl $siteUrl -SiteContext $siteContext          
                }
                AddSiteAdmins -Admins $addedOwnerUPN -SiteUrl $siteUrl -SiteContext $siteContext
                AddSiteOwners  -Owner $addedOwnerUPN -SiteUrl $siteUrl -SiteContext $siteContext   
            }
        }
        LogWrite -Message " $($script:ProcessInProgress): Completing provision..."
        #Update storage quota and external sharing         
        UpdateSiteStorage -SiteUrl $siteUrl -SiteContext $siteContext
        UpdateExternalSharing -SiteUrl $siteUrl -ExternalSharing $externalSharing -SiteContext $siteContext          
        
        #Complete-ProvisionRequest -Request $Request        
        UpdateProvisionRequest -RequestId $reqId -ReqStatusID $script:Completed -ReqObjectId $reqObjectId -connectionString $script:ConnectionString 
        LogWrite -Message " $($script:ProcessInProgress): Updated the Request Status [Completed] for the site."
        #Update ServiceNow Ticket
        Update_SNIncident -IncidenttID $IncidentId -IncidentType Provision -IncidentStatus Resolved -SiteURL $siteUrl
        
        LogWrite -Message " $($script:ProcessInProgress): Validate if the site request is completed before sending email confirmation..."
        $requestInfo = GetSiteRequestInfoById -requestId $reqId -connectionString $script:ConnectionString
        if ($null -ne $requestInfo -and $requestInfo["RequestStatusId"] -eq $script:Completed){            
            SendEmailConfirmation -Request $requestInfo
            LogWrite -Message " $($script:ProcessInProgress): Sent an email confirmation to the requestor and site owner."
            #verify if the site is completely updated, if not wait...until site is updated before sync to DB
            LogWrite -Message " $($script:ProcessInProgress): Validate if the site is updated before sync up to DB..."
            $siteCollection = Get-PnPTenantSite -url $SiteUrl -Detailed -Connection $siteContext            
            $retryCount = 5
            $retryAttempts = 0
            while ($siteCollection.Status -ne 'Active' -and $retryAttempts -lt $retryCount) {
                LogWrite -Message " $($script:ProcessInProgress): Waiting until the site is updated..."
                Sleep 20
                $retryAttempts = $retryAttempts + 1
                $siteCollection = Get-PnPTenantSite -url $SiteUrl -Detailed -Connection $siteContext
                
                #$sca = (Get-PnPSiteCollectionAdmin | ? {$_.Email -ne ''}).Email -join ";"
                #$siteCollection.SecondarySCA = $sca
            }
            #If an error occured while sync the provisioned site to DB then skip and proceed next request?
            if($siteCollection.Status -eq 'Active'){
                SyncProvisionedSiteToDB -Request $Request -SiteObject $siteCollection -connectionString $script:ConnectionString               
                LogWrite -Message " $($script:ProcessInProgress)[Completion]: [$siteUrl] has been provisioned successfully."
            }
            else{
                LogWrite -Message " $($script:ProcessInProgress): The newly provisioned site has not been synced to DB."            
            }            
        }
    }
    catch {
        LogWrite -Level ERROR "$($script:ProcessInProgress): Something went wrong while processing a pending site request: [$siteUrl]. Error Info: $($_.Exception)"
        #Mark as skip if there is something wrong for second run
        LogWrite -Message " $($script:ProcessInProgress): Mark as skip if there is something wrong for second run"
        UpdateProvisionRequest -RequestId $reqId -ReqStatusID $script:InProgress -ReqObjectId $reqObjectId  -ReqProcessFlag 1 -connectionString $script:ConnectionString        
        throw $_         
    }
    finally{        
        if ($siteContext) {
            LogWrite -Message "  The Sharepoint Online Session is now closed."
            DisconnectPnpOnlineOAuth -Context $siteContext
        }
    }
    
}
#endregion

#region Decommission sharepoint site
Function Decommission-SPOSite {    
    param([Parameter(Mandatory = $true)] $Request)
    
    try {
        $decomissionedSiteUrl = $Request["SiteUrl"]        
        LogWrite -Message " $($script:ProcessDecommission): Connecting to SharePoint Admin Center '$($script:SPOAdminCenterURL)'..."
        #if ($null -eq $script:TenantContext){
            $script:TenantContext = ConnectPnpOnlineOAuth -TenantId $script:TenantId -ClientId $script:appIdOperationSupport -Thumbprint $script:appThumbprintOperationSupport -Url $script:SPOAdminCenterURL
        #}        
        Remove-PnPTenantSite -Url $decomissionedSiteUrl -Force        
        Decommission-SPOSiteCompletion -Request $Request             
    }
    catch {
        LogWrite -Level ERROR "$($script:ProcessDecommission): An error occurred during decommission SPO Site: [$decomissionedSiteUrl). Error Info: $($_.Exception)"        
        throw $_        
    }
    finally {    
        LogWrite -Level INFO -Message " Disconnect SharePoint Admin Center." 
        DisconnectPnpOnlineOAuth -Context $script:TenantContext     
    }
    
}

Function Decommission-SPOSiteCompletion{
    param([Parameter(Mandatory = $true)] $Request)

    try{
        $reqId = $Request["RequestId"].Guid
        $reqStatus = $Request["RequestStatusId"].Guid
        $reqObjectId = $Request["ObjectId"].Guid
        $siteUrl = $Request["SiteUrl"]    
        $IncidentId = $Request["IncidentId"]
        $externalSharing = $Request["ExternalSharingEnabled"]
        $ICName = $Request["ICName"]

        LogWrite -Message " $($script:ProcessDecommission): Verify if the site is already deleted..."
        $deletedSite = Get-PnPTenantRecycleBinItem | ? { $_.url -eq $siteUrl}
        if ($null -eq $deletedSite){
            LogWrite -Message " $($script:ProcessDecommission): The site has not been decommissioned."
            return        
        }
        $deletedSite = ParseSPOSite -siteObj $deletedSite -ICName $ICName -ExternalSharingEnabled $externalSharing
        LogWrite -Message " $($script:ProcessDecommission): Completing decommission..."    
        $reqStatus = $script:Completed
        LogWrite -Message " $($script:ProcessDecommission): Updated the deleted site to Sites table..."
        UpdateSPOSiteToDatabase -connectionString $script:ConnectionString -siteData $deletedSite
        UpdateProvisionRequest -RequestId $reqId -ReqStatusID $reqStatus -ReqObjectId $reqObjectId -connectionString $script:ConnectionString 
        LogWrite -Message " $($script:ProcessDecommission): Updated the Request Status [Completed] for the site."
        #Update ServiceNow Ticket
        Update_SNIncident -IncidenttID $IncidentId -IncidentType Decommission -IncidentStatus Resolved -SiteURL $siteUrl
        LogWrite -Message " $($script:ProcessDecommission)[Completion]: The site [$siteUrl] has been decommissioned successfully."
        #Send email to the Requestor
        $requestInfo = GetSiteRequestInfoById -requestId $reqId -connectionString $script:ConnectionString
        if ($null -ne $requestInfo -and $requestInfo["RequestStatusId"] -eq $script:Completed){   
            SendEmailConfirmation -Request $requestInfo
            LogWrite -Message " $($script:ProcessInProgress): Sent an email confirmation to the requestor and site owner."
        }
    }
    catch{
        throw $_
    }
}
#endregion

Function Provision-SPOSite {    
    param([Parameter(Mandatory = $true)] $Request)
    
    try {
        LogWrite -Message " $($script:ProcessNew): Provisioning SharePoint Site..."
        LogWrite -Message " $($script:ProcessNew): Connecting to SharePoint Admin Center '$($script:SPOAdminCenterURL)'..."
        #if ($null -eq $script:TenantContext){
            $script:TenantContext = ConnectPnpOnlineOAuth -TenantId $script:TenantId -ClientId $script:appIdOperationSupport -Thumbprint $script:appThumbprintOperationSupport -Url $script:SPOAdminCenterURL
        #}
        $reqSiteUrl = $Request["SiteUrl"]        
        $site = Get-PnPTenantSite -Url $reqSiteUrl
        if ($site.Status -eq 'Active'){
            LogWrite " $($script:ProcessNew): The site [$reqSiteUrl] already exists."
            return $site.Url            
        }
        # Should validate if site request with the same URL exists in the Recycle Bin
        # Because creating a new site fails if a deleted site with the same URL exists in the Recycle Bin        
        $deletedSite = Get-PnPTenantRecycleBinItem | ? { $_.url -eq $reqSiteUrl}
        if ($null -ne $deletedSite){
            LogWrite -Message " $($script:ProcessNew): The site is deleted. Check it in 'Deleted sites' in SharePoint Admin Center." 
            return $deletedSite.Url       
        }
        # Note: only wait for the site to be provisioned if the -Wait switch is set
        #$siteOwners = $Request["OwnerUPN"] -join ','
        #[String[]]$siteOwners= "buint@citspdev.onmicrosoft.com","tandra2@citspdev.onmicrosoft.com"

        switch ($Request["TemplateId"]) {            
            "SITEPAGEPUBLISHING#0" {
                #$provisionedSite = New-PnPSite -Type CommunicationSite -Title $Request["DisplayName"] -Description $Request["Description"] -Url $Request["SiteUrl"] -SiteDesign $Request["SiteDesign"] -Owners $Request["OwnerUPN"] -Wait:$false 
                $provisionedSite = New-PnPSite -Type CommunicationSite -Title $Request["DisplayName"] -Description $Request["Description"] -Url $Request["SiteUrl"] -SiteDesign $Request["SiteDesign"] -Owner $script:CloudSvcForProvision -Wait:$false 
                
            }
            "STS#3" {
                #$provisionedSite = New-PnPTenantSite -Template "STS#3" -Title $Request["DisplayName"] -Url $Request["SiteUrl"] -Owner $Request["OwnerUPN"] -Lcid $Request["LcId"] -TimeZone $Request["TimeZone"] -ResourceQuota $Request["ResourceQuota"] -StorageQuota $Request["StorageQuota"] -StorageQuotaWarningLevel $Request["StorageWarningLevel"] -Wait:$false
                $provisionedSite = New-PnPTenantSite -Template "STS#3" -Title $Request["DisplayName"] -Url $Request["SiteUrl"] -Owner $script:CloudSvcForProvision -Lcid $Request["LcId"] -TimeZone $Request["TimeZone"] -ResourceQuota $Request["ResourceQuota"] -StorageQuota $Request["StorageQuota"] -StorageQuotaWarningLevel $Request["StorageWarningLevel"] -Wait:$false
            }
            "STS#0" {
                #$provisionedSite = New-PnPTenantSite -Template "STS#0" -Title $Request["DisplayName"] -Url $Request["SiteUrl"] -Owner $Request["OwnerUPN"] -Lcid $Request["LcId"] -TimeZone $Request["TimeZone"] -ResourceQuota $Request["ResourceQuota"] -StorageQuota $Request["StorageQuota"] -StorageQuotaWarningLevel $Request["StorageWarningLevel"] -Wait:$false
                $provisionedSite = New-PnPTenantSite -Template "STS#0" -Title $Request["DisplayName"] -Url $Request["SiteUrl"] -Owner $script:CloudSvcForProvision -Lcid $Request["LcId"] -TimeZone $Request["TimeZone"] -ResourceQuota $Request["ResourceQuota"] -StorageQuota $Request["StorageQuota"] -StorageQuotaWarningLevel $Request["StorageWarningLevel"] -Wait:$false
            }
        }
        # Verify if site was provisioned
        $siteCollection = Get-PnPTenantSite -url $reqSiteUrl -Detailed
        $retryCount = 5
        $retryAttempts = 0
        while ($siteCollection.Status -ne 'Active' -and $retryAttempts -lt $retryCount) {
            LogWrite -Message " $($script:ProcessNew): Waiting until the site is provisioned..."
            Sleep 30
            $retryAttempts = $retryAttempts + 1
            $siteCollection = Get-PnPTenantSite -url $reqSiteUrl -Detailed
            $provisionedSite = $siteCollection.Url
        }
        return $provisionedSite              
    }
    catch {
        LogWrite -Level ERROR "$($script:ProcessNew): An error occurred during provision SPO Site: [$($Request["TemplateId"]) - $($Request["SiteUrl"])]. Error Info: $($_.Exception)"        
        throw $_        
    }
    finally {    
        LogWrite -Level INFO -Message " Disconnect SharePoint Admin Center." 
        DisconnectPnpOnlineOAuth -Context $script:TenantContext     
    }
}

Function Provision-Team {    
    param([Parameter(Mandatory = $true)] $Request)
    
    try {
        LogWrite -Message " $($script:ProcessNew): Provisioning Microsoft Teams..."
        #if ($null -eq $script:TokenOperationSupport){
            $script:TokenOperationSupport = Connect-GraphAPIWithCert -TenantId $script:TenantId -AppId $script:appIdOperationSupport -Thumbprint $script:appThumbprintOperationSupport
        #}
        #LogWrite -Message "Determine if M365 Group has no associated team then promote to team"
        #$team = Get-NIHTeam -AuthToken $AuthToken -Id $Id
        #$provisionedSiteUrl = New-NIHTeam -AuthToken $AuthToken  -Name $Request["DisplayName"] -Description $Request["Description"] -Visibility private -Owner $Request["OwnerId"]        
        $alias = $Request["Alias"]
        LogWrite -Message " $($script:ProcessNew): Checking the group with alias ($alias) exists or not..."
        $group = Get-NIHO365GroupByAlias -AuthToken $script:TokenOperationSupport -MailNickName $alias
        if($group -and $group.groupTypes.Count -eq 0) { # -and $group.mailEnabled -eq $true){            
            LogWrite -Message " $($script:ProcessNew): The group with alias ($alias) exists.It could be DL or Mail-enabled security group.Skip this request."
            LogWrite -Message " $($script:ProcessNew): Mark this request as pending and notify the customer submit another request with different name."
            UpdateProvisionRequest -RequestId $Request["Id"] -ReqStatusID $script:Submitted -ReqObjectId $Request["ObjectId"]  -ReqProcessFlag 1 -connectionString $script:ConnectionString
            return
        }
        LogWrite -Message " $($script:ProcessNew): Continue to provision MS Team..."
        $provisionedTeam = New-NIHTeamGroup -AuthToken $script:TokenOperationSupport -MailNickName $Request["Alias"] -Name $Request["DisplayName"] -Description $Request["Description"] -Members $script:CloudSvcForProvision -Owners $script:CloudSvcForProvision        
        
        if ($provisionedTeam.id){
            $provisionedTeam = $provisionedTeam.id
            #Update-NIHTeamSettingsPostProvision -AuthToken $script:TokenOperationSupport -Id $provisionedTeam
        }
        else{
            LogWrite -Level ERROR "$($script:ProcessNew)[Provision-Team]: An error occurred during provision Team - ($($Request["SiteUrl"])). Error Info: $provisionedTeam" 
            $provisionedTeam = $null           
        }
        return $provisionedTeam #| Out-Null        
    }
    catch {        
        LogWrite -Level ERROR "$($script:ProcessNew)[Provision-Team]: Something went wrong with provision Team: ($($Request["DisplayName"])). Error Info: $($_.Exception)"        
        throw $_
    }
}

Function Provision-Group {    
    param([Parameter(Mandatory = $true)] $Request)
    
    try {
        LogWrite -Message " $($script:ProcessNew): Provisioning Microsoft M365 Groups..."
        $script:TokenOperationSupport = Connect-GraphAPIWithCert -TenantId $script:TenantId -AppId $script:appIdOperationSupport -Thumbprint $script:appThumbprintOperationSupport

        $alias = $Request["Alias"]
        LogWrite -Message " $($script:ProcessNew): Checking the group with alias ($alias) exists or not..."
        $group = Get-NIHO365GroupByAlias -AuthToken $script:TokenOperationSupport -MailNickName $alias
        if($group -and $group.groupTypes.Count -eq 0){            
            LogWrite -Message " $($script:ProcessNew): The group with alias ($alias) exists.It could be DL or Mail-enabled security group.Skip this request."
            LogWrite -Message " $($script:ProcessNew): Mark this request as pending and notify the customer submit another request with different name."
            UpdateProvisionRequest -RequestId $Request["Id"] -ReqStatusID $script:Submitted -ReqObjectId $Request["ObjectId"]  -ReqProcessFlag 1 -connectionString $script:ConnectionString
            return
        }
        LogWrite -Message " $($script:ProcessNew): Continue to provision M365 Group..."

        $privacySetting = "private"
        if ($null -ne $Request["PrivacySetting"]){
            $privacySetting = $Request["PrivacySetting"]
        }

        $provisionedGroup = New-NIHO365Group -AuthToken $script:TokenOperationSupport -MailNickName $Request["Alias"] -Name $Request["DisplayName"] -Description $Request["Description"] -Visibility $privacySetting -Members $script:CloudSvcForProvision -Owners $script:CloudSvcForProvision        
        
        if ($provisionedGroup.id){
            $provisionedGroup = $provisionedGroup.id            
        }
        else{
            LogWrite -Level ERROR "$($script:ProcessNew)[Provision-Group]: An error occurred during provision Group - ($($Request["SiteUrl"])). Error Info: $provisionedGroup" 
            $provisionedGroup = $null           
        }
        return $provisionedGroup #| Out-Null        
    }
    catch {        
        LogWrite -Level ERROR "$($script:ProcessNew)[Provision-Group]: Something went wrong with provision Group: ($($Request["DisplayName"])). Error Info: $($_.Exception)"        
        throw $_
    }
    
}

Function AddSiteOwners {
    param($Owner, $SiteUrl,$SiteContext)    
    try {        
        if ($null -eq $siteContext){                                        
            $siteContext = ConnectPnpOnlineOAuth -TenantId $script:TenantId -ClientId $script:appIdOperationSupport -Thumbprint $script:appThumbprintOperationSupport -Url $siteUrl           
        }
        
        LogWrite -Message " $($script:ProcessInProgress): Adding SharePoint Site Owners for the site..."

        #$web = Get-PnPWeb 
        #$ctx = $web.Context          
        $ownersGroupName = (Get-PnPGroup -AssociatedOwnerGroup).Title
        
        #$u = Get-PnPGroupMember -Group $ownersGroupName -User $script:CloudSvcForProvision
        #if ($u) {
        #    Remove-PnPGroupMember -LoginName $script:CloudSvcForProvision -Group $ownersGroupName
        #}

       
        Get-PnPGroupMember -Group $ownersGroupName | ForEach-Object {
            $user = $_                    
            Remove-PnPGroupMember -LoginName $user.LoginName -Group $ownersGroupName
        }

        LogWrite -Message " $($script:ProcessInProgress): SharePoint Group Owner: [$($ownersGroupName)]"
        foreach($o in $Owner){
            Add-PnPGroupMember -LoginName $o -Group $ownersGroupName # for PnP.PowerShell
        }
        LogWrite -Message " $($script:ProcessInProgress): Finished adding SharePoint Site Owners [$Owner]."
    }
    catch {
        #$exception = $_.Exception
        LogWrite -Level ERROR "$($script:ProcessInProgress)[AddSiteOwners]: Error adding SharePoint Site Owners for the site [$SiteUrl]:"
        throw $_
    }
}

function AddSiteAdmins { 
    param($Admins, $SiteUrl,$SiteContext) 
    
    try {
        if ($null -eq $siteContext){                                        
            $siteContext = ConnectPnpOnlineOAuth -TenantId $script:TenantId -ClientId $script:appIdOperationSupport -Thumbprint $script:appThumbprintOperationSupport -Url $siteUrl           
        }
        
        LogWrite -Message " $($script:ProcessInProgress): Adding site admins for the site..."

        Add-PnPSiteCollectionAdmin -Owners $Admins -Connection $siteContext
        Remove-PnPSiteCollectionAdmin -Owners $script:CloudSvcForProvision -Connection $siteContext
    }
    catch {
        $exception = $_.Exception
        LogWrite -Level ERROR "$($script:ProcessInProgress)[AddSiteAdmins]: Error adding admins for the site [$SiteUrl]: $exception"
        throw $_
    }    
}

Function UpdateSiteDescription {
    param($SiteDescription, $SiteUrl,$SiteContext)    
    try {        
        if ($null -eq $siteContext){                                        
            $siteContext = ConnectPnpOnlineOAuth -TenantId $script:TenantId -ClientId $script:appIdOperationSupport -Thumbprint $script:appThumbprintOperationSupport -Url $siteUrl           
        }
        
        if ($SiteDescription -ne $null) {
            LogWrite -Message " $($script:ProcessInProgress): Updating Description for the Site..."            
            Set-PnPWeb -Description $SiteDescription -Connection $siteContext
        }

        LogWrite -Message " $($script:ProcessInProgress): Finished updating Description for the site."
    }
    catch {
        #$exception = $_.Exception
        LogWrite -Level ERROR "$($script:ProcessInProgress)[UpdateSiteDescription]: Error updating description for the site [$SiteUrl]:"
        throw $_
    }
}

Function UpdateSiteStorage {
    param(
        [Parameter(Mandatory = $true)] $SiteUrl,
        [Parameter(Mandatory = $true)] $SiteContext
    )
    try {
        if ($null -eq $siteContext){                                        
            $siteContext = ConnectPnpOnlineOAuth -TenantId $script:TenantId -ClientId $script:appIdOperationSupport -Thumbprint $script:appThumbprintOperationSupport -Url $siteUrl           
        }
    
        $storageQuota = 1048576 #1 TB = 1,048,576 MB
        $DefaultStorageWarningPercent = 90
        $storageWarningLevel = [math]::Round((($storageQuota) * ($DefaultStorageWarningPercent)) / 100) #943718        
        Set-PnPTenantSite -Url $siteUrl -StorageMaximumLevel $storageQuota -StorageWarningLevel $storageWarningLevel -Connection $siteContext

        LogWrite -Message " $($script:ProcessInProgress): Finished updating quota storage for the site."
    }
    catch {        
        LogWrite -Level ERROR "$($script:ProcessInProgress)[UpdateSiteStorage]: Error updating quota storage for the site [$SiteUrl]:"            
        throw $_
    }

} 

Function UpdateExternalSharing{
    param(
        [Parameter(Mandatory = $true)] $SiteUrl,
        [Parameter(Mandatory = $true)] $SiteContext,
        [Parameter(Mandatory = $true)] $ExternalSharing

    )
    try{
        if ($null -eq $SiteContext){                                        
            $SiteContext = ConnectPnpOnlineOAuth -TenantId $script:TenantId -ClientId $script:appIdOperationSupport -Thumbprint $script:appThumbprintOperationSupport -Url $siteUrl           
        }

        $siteCollection = Get-PnPTenantSite -url $SiteUrl -Detailed -Connection $SiteContext
            
        if ($siteCollection){            
            if ($ExternalSharing -eq $true) {                         
                Set-PnPTenantSite -Url $SiteUrl -SharingCapability ExternalUserSharingOnly
            }                    
            else{
                Set-PnPTenantSite -Url $SiteUrl -SharingCapability Disabled
            }      
        }

        LogWrite -Message " $($script:ProcessInProgress): Finished Enable/disable external sharing for the site."
    }
    catch {
        LogWrite -Level ERROR "$($script:ProcessInProgress)[ExternalSharing]: Error enable/disable external sharing for the site [$SiteUrl]:"            
        throw $_
    }
}

Function SyncProvisionedSiteToDB{
    param([Parameter(Mandatory = $true)] $Request,
            [Parameter(Mandatory = $true)] $SiteObject,
            [Parameter(Mandatory=$true)]$connectionString)

    try{  
        $siteUrl = $Request["SiteUrl"]
        $externalSharing = $Request["ExternalSharing"]
        $primarySCA = $Request["primarySCA"]
        $secondarySCA = $Request["OwnerUPN"] -join ";"
        
        $groupId = $SiteObject.GroupId
        $OwnerEmail = $SiteObject.OwnerEmail

        #$groupId = $siteCollection.GroupId
        #$OwnerEmail = $siteCollection.OwnerEmail
        
        # insert newly provisioned site to Sites table
        LogWrite -Message " $($script:ProcessInProgress): Starting sync the newly provisioned site to DB..."
        #LogWrite -Message " $($script:ProcessInProgress): ExternalSharingEnabled: $externalSharing"
        $site = ParseSPOSite -siteObj $SiteObject -ICName $Request["ICName"] -PrimarySCA $primarySCA -SecondarySCA $secondarySCA -ExternalSharingEnabled $externalSharing
        #LogWrite -Message " $($script:ProcessInProgress): $site"
        UpdateSPOSiteToDatabase $connectionString $site
        LogWrite -Message " $($script:ProcessInProgress): The newly provisioned site has been synced to DB."

        # insert to Groups,Teams and TeamChannel table if this is MS Teams request
        if ($groupId -ne $script:GuidEmpty) {             
            LogWrite -Message " $($script:ProcessInProgress): Starting sync to DB...[$groupId]"
            $script:TokenOperationSupport = Connect-GraphAPIWithCert -TenantId $script:TenantId -AppId $script:appIdOperationSupport -Thumbprint $script:appThumbprintOperationSupport
            
            $groupMembers = $Request["OwnerUPN"] -join ";"
            $channelMembers = $Request["OwnerId"] -join ";"
            $privacySetting = $Request["PrivacySetting"] 
            $hideFromAddressLists = $false # group by default value
            $hideFromOutlookClients = $false # group by default value
            $templateId = $Request["TemplateId"]
            $currentDate = Get-Date
                           
            if ($templateId -eq $script:EnabledTeam){
                $privacySetting = "Private"
                $resourceProvisioningOptions = $script:EnabledTeam
                $hideFromAddressLists = $true # team by default value
                $hideFromOutlookClients = $true # team by default value
                # insert to Teams and TeamChannel tables only for MS teams template
                LogWrite -Message " $($script:ProcessInProgress): Starting sync to Teams;TeamChannel..."          
                $Team = Get-NIHTeam -AuthToken $script:TokenOperationSupport -Id $groupId
                if ($Team.id) {
                    $objTeam = ParseO365Team -Team $Team
                    $objTeamChannel = Get-NIHTeamChannel -AuthToken $script:TokenOperationSupport -Id $groupId -MembershipType standard                        
                }
                else {                    
                    $objTeam = [PSCustomObject]@{                            
                        GroupID                                 = $groupId                                       
                        DisplayName                             = $Request["DisplayName"]
                        Description                             = $requestInfo["Description"]
                        InternalId                              = ""
                        Classification                          = ""
                        CreatedDateTime                         = $currentDate                            
                        WebUrl                                  = $null
                        IsArchived                              = $null
                        OperationStatus                         = ""; 
                        Operation                               = ""; 
                        AdditionalInfo                          = ""
                    }
                }                    
                UpdateSQLTeam $connectionString $objTeam
                LogWrite -Message " $($script:ProcessInProgress): Insert new Team record into Teams table completed." 

                # insert to TeamChannel table            
                if ($objTeamChannel.id){
                    $objTeamChannel = [PSCustomObject]@{
                        GroupID             = $groupId
                        Id                  = $objTeamChannel.id
                        DisplayName         = $objTeamChannel.displayName
                        Description         = $objTeamChannel.description
                        IsFavoriteByDefault = $objTeamChannel.isFavoriteByDefault
                        Email               = $objTeamChannel.email
                        WebUrl              = $objTeamChannel.webUrl
                        MembershipType      = $objTeamChannel.membershipType
                        ChannelOwners       = $channelMembers
                        ChannelMembers      = $channelMembers
                        ChannelGuests       = ""
                        OperationStatus     = "" 
                        Operation           = "" 
                        AdditionalInfo      = ""
                        }
                }
                else{
                    $objTeamChannel = @([PSCustomObject]@{
                        GroupID             = $groupId
                        Id                  = ""
                        DisplayName         = "General"
                        Description         = ""
                        IsFavoriteByDefault = ""
                        Email               = ""
                        WebUrl              = ""
                        MembershipType      = "standard"
                        ChannelOwners       = $channelMembers
                        ChannelMembers      = $channelMembers
                        ChannelGuests       = ""
                        OperationStatus     = "" 
                        Operation           = "" 
                        AdditionalInfo      = ""
                        })
                }           
            
                UpdateTeamsChannel $connectionString $objTeamChannel
                LogWrite -Message " $($script:ProcessInProgress): Insert into TeamChannel table completed." 
            }

            # insert to Groups table
            LogWrite -Message " $($script:ProcessInProgress): Starting sync to Groups..."
            $Group = Get-NIHO365Group -AuthToken $script:TokenOperationSupport -Id $groupId                   
                                       
            if ($Group.id) {
                $objGroup = ParseO365Group -Group $Group -GroupOwners $groupMembers -GroupMembers $groupMembers -HideFromAddressLists $hideFromAddressLists -HideFromOutlookClients $hideFromOutlookClients
            }
                   
            else{
                $objGroup = [PSCustomObject]@{
                        GroupID                      = $groupId
                        DisplayName                  = $Request["DisplayName"]
                        Description                  = $Request["Description"]
                        GroupOwners                  = $groupMembers
                        GroupMembers                 = $groupMembers
                        GroupGuests                  = ""
                        CreatedDateTime              = $currentDate 
                        Visibility                   = $privacySetting
                        #Mail                         = $Request["DisplayName"]
                        MailNickname                 = $Request["Alias"]
                        ResourceProvisioningOptions  = $resourceProvisioningOptions
                        HideFromAddressLists         = $hideFromAddressLists
                        HideFromOutlookClients       = $hideFromOutlookClients
                        OperationStatus              = ""; 
                        Operation                    = ""; 
                        AdditionalInfo               = ""   
                    }
                }                    
            UpdateSQLO365Group $connectionString $objGroup
            LogWrite -Message " $($script:ProcessInProgress): Insert new Group record into Groups table completed." 

        }
    }
    catch{
        LogWrite -Level Error "$($script:ProcessInProgress)[SyncProvisionedSiteToDB]: $($_.Exception)"
        #LogWrite -Level Error "$($script:ProcessInProgress)[SyncProvisionedSiteToDB]: An error occured while sync the provisioned site to DB:"
        #throw $_        
    }
}

#region provision PowerBI Workspace
Function Provision-PowerBIWorkspace{
    param([Parameter(Mandatory = $true)] $Request)
    
    try {
        LogWrite -Message " $($script:ProcessNew): Provisioning PowerBI Workspace..."
        LogWrite -Message " $($script:ProcessNew): Connecting to PowerBI service..."
        ConnectPowerBIService -Environment USGov -Tenant $script:TenantName -AppId $script:appIdPowerBIWorkspace -Thumbprint $script:appThumbprintPowerBIWorkspace
        
        $wsName = $Request["DisplayName"]
        $wsDesc = $Request["Description"]
        $wsAdmin = [string]($Request["OwnerUPN"])
        $reqId = $Request["Id"]
        $reqObjectId = $Request["ObjectId"]
        $IncidentId = $Request["IncidentId"]
               

        LogWrite -Message " $($script:ProcessNew): Checking if PowerBI Workspace [$wsName] exists..."
        $ws = Get-PowerBIWorkspace -Name $wsName
        if ($ws.Id){
            LogWrite -Message " $($script:ProcessNew): PowerBI Workspace [$wsName] exists. Skip the request."
            UpdateProvisionRequest -RequestId $reqId -ReqStatusID $script:Submitted -ReqObjectId $reqObjectId  -ReqProcessFlag 1 -connectionString $script:ConnectionString
            return
        }
        # Continue provisioning Workspace if ws doesn't exist.

        $provisionedWs = New-PowerBIWorkspace -Name $wsName

        if (-Not $provisionedWs){
            LogWrite -Level ERROR "$($script:ProcessNew)[Provision-PowerBIWorkspace]: Unable to provision PowerBI Workspace [$wsName]"
            Resolve-PowerBIError -Last
        }
        # Verify if PowerBI Workspace was provisioned
        $wsId = $provisionedWs.Id.Guid
        $ws = Get-PowerBIWorkspace -Id $wsId -Scope Organization -Include All
        #$ws.Description = $wsDesc
        $retryCount = 5
        $retryAttempts = 0
        while ($ws.State -ne 'Active' -and $retryAttempts -lt $retryCount) {
            LogWrite -Message " $($script:ProcessNew): Waiting until the PowerBI Workspace is provisioned..."
            Sleep 30
            $retryAttempts = $retryAttempts + 1
            $ws = Get-PowerBIWorkspace -Id $wsId -Scope Organization -Include All
            #$provisionedWs = $ws.Name
        }
        LogWrite -Message " $($script:ProcessNew): Adding owner as admin of newly created PowerBI Workspace..."
        Add-PowerBIWorkspaceUser -Id $provisionedWs.Id -UserPrincipalName $wsAdmin -AccessRight Admin       
        $ws = Get-PowerBIWorkspace -Id $wsId -Scope Organization -Include All | Select *
        $reqObjectId = $ws.Id.Guid
        <#
        # Check if Service Principal is in the workspace
        LogWrite -Message " $($script:ProcessNew): Remove Service Principal from newly created PowerBI Workspace..."
        $ServicePrincipalInWorkspace = $ws.Users | Where-Object {$_.Identifier -eq $script:PowerBIServicePrincipalId}
        if ($ServicePrincipalInWorkspace)
        {            
            try {
                Invoke-PowerBIRestMethod -Method Delete -Url "admin/groups/$wsId/users/$($script:PowerBIServicePrincipalId)"
                LogWrite -Message "Remove Service Principal done."
            }
            catch {
                LogWrite -Level ERROR "$($script:ProcessNew)[Provision-PowerBIWorkspace]: Remove Service Principal get failed."
                Resolve-PowerBIError -Last
            }
        }
        else {
            LogWrite -Message "Service Principal is not a member of: $wsName."
        }
        #>

        #Complete-ProvisionRequest -Request $Request        
        UpdateProvisionRequest -RequestId $reqId -ReqStatusID $script:Completed -ReqObjectId $reqObjectId -connectionString $script:ConnectionString 
        LogWrite -Message " $($script:ProcessNew): Updated the Request Status [Completed] for the site."
        
        #Update ServiceNow Ticket
        Update_SNIncident -IncidenttID $IncidentId -IncidentType Provision -ServiceType PowerBI -IncidentStatus Resolved -SiteURL $wsName
        
        #Send Email Confirmation to requestor and owner
        $requestInfo = GetSiteRequestInfoById -requestId $reqId -connectionString $script:ConnectionString
        if ($null -ne $requestInfo -and $requestInfo["RequestStatusId"] -eq $script:Completed){            
            SendEmailConfirmation -Request $requestInfo
            LogWrite -Message " $($script:ProcessNew): Sent an email confirmation to the requestor and site owner."
        }

        #Sync workspace info to db
        SyncProvisionedPowerBIWorkspaceToDB -Request $Request -Workspace $ws -connectionString $script:ConnectionString
        return $ws      
    }
    catch {        
        LogWrite -Level ERROR "$($script:ProcessNew)[Provision-PowerBIWorkspace]: Something went wrong with provision PowerBI Workspace:[$wsName]. Error Info: $($_.Exception)"        
        throw
    }
    finally {    
        LogWrite -Level INFO -Message " Disconnect PowerBI service." 
        DisconnectPowerBIService     
    }
}

Function SyncProvisionedPowerBIWorkspaceToDB{
    param([Parameter(Mandatory = $true)] $Request, 
        [Parameter(Mandatory = $true)] $Workspace,           
          [Parameter(Mandatory=$true)]$connectionString)

    try{         
        LogWrite -Message " $($script:ProcessInProgress): Starting sync the newly provisioned PowerBI Workspace to DB..."
        $ws = ParsePowerBIworkspace -Workspace $Workspace -ICName $Request["ICName"]
        UpdateSQLPowerBIworkspace $connectionString $ws
        LogWrite -Message " $($script:ProcessInProgress): The newly provisioned PowerBI Workspace has been synced to DB."
         
    }
    catch{
        LogWrite -Level Error "$($script:ProcessInProgress)[SyncProvisionedPowerBIWorkspaceToDB]: $($_.Exception)"        
    }
}
#endregion
