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
."$script:RootDir\Common\Lib\GraphAPILibO365UsersDAO.ps1"
."$script:RootDir\Common\Lib\GraphAPILibSPOSitesDAO.ps1"
."$script:RootDir\Common\Lib\GraphAPILibO365GroupsDAO.ps1"
."$script:RootDir\Common\Lib\LibRequestDAO.ps1"
."$script:RootDir\Common\Lib\OperationHelper.ps1"


<#
     =============================================================================================================
      .DESCRIPTION
        Process all change requests - [ChangeRequests] below
        ----------------------------------------------------------------------------------------------------------
        ChangeTypeId	                        ChangeTypeValue	              ChangeRequestTypeId
        ----------------------------------------------------------------------------------------------------------        
        4D2A196D-C1A7-4F26-B2EF-2E6008B50BE0	Enable App Catalog	          732DBBD1-37ED-4757-9892-CE7E170FDEB7        
        F8372500-FB02-4BDA-B13A-6C531E2432A0	Enable Site Customization	  732DBBD1-37ED-4757-9892-CE7E170FDEB7
        FF07775A-9883-4C42-89C1-947F347AE712	Enable External Sharing	      732DBBD1-37ED-4757-9892-CE7E170FDEB7        
        A7C753F5-AA0A-411F-AC51-F3D01A1348EE	Register as Hub Site	      732DBBD1-37ED-4757-9892-CE7E170FDEB7
        F5B0B64D-6670-4AC9-94EF-BC5CE1A52A84	Site Admin Access	          732DBBD1-37ED-4757-9892-CE7E170FDEB7

        B252D6A3-D11A-4AB3-84A9-9DA4769DA3F2	Change Display Name	          32982168-6CA3-402D-9991-64606792A6DE
        2096A9FE-8AFF-4580-975D-3283B975F9A0	Change Description	          32982168-6CA3-402D-9991-64606792A6DE
        3B4B6991-9F62-46F6-9508-59E51E48B3CF	Change Privacy	              32982168-6CA3-402D-9991-64606792A6DE        
        AC2CB606-68A9-490F-9971-EFAB6FEB81E0	Hidden from Outlook/GAL	      32982168-6CA3-402D-9991-64606792A6DE
        414A5774-588A-49A1-B01D-868934329D08	Owner Access	              32982168-6CA3-402D-9991-64606792A6DE
        EB05368F-472B-491C-AF18-2D5BE51136DD	Private Channel Access	      32982168-6CA3-402D-9991-64606792A6DE        
        
        ----------------------------------------------------------------------------------------------------------
        RequestStatusId	                        StatusValue
        ----------------------------------------------------------------------------------------------------------
        5A6C2888-1D7F-4FFC-94FA-0A92640C7076	Submitted
        6BFA67B8-18E2-46CB-A8B8-651F36C4C9A5	Cancelled
        E950207B-AAF3-4B65-9BB7-689A6B6AE83D	Completed
        6B934E0F-8784-461B-80D0-A4660F6D1A4E	In Progress
        6C4B8971-FAFF-4B26-BB6F-FD4C5CEA66AA	Pending
    ==============================================================================================================
#>

Function Process-TeamChangeRequests {
    param
    (        
        [Parameter(Mandatory = $true)] $Requests
        
    )

    try {
        LogWrite -Level INFO -Message "Processing Team Change Requests."
        LogWrite -Level INFO -Message "Total Team Change Requests: $($Requests.Count)"

        if ($Requests.Count -ile 0) {
            return
        }        

        foreach ($request in $Requests) {
            try {                
                $teamId = $request.GroupId.Trim()
                $newValue = $request.NewValue                
                $requestType = $request.ChangeTypeId.Guid
                LogWrite -Level INFO -Message "Team ID: $teamId"

                switch ($requestType) {
                    $script:ChangeDisplayName {
                        $ret = EditTeam -GroupId $teamId -DisplayName $newValue                    
                    }
                    $script:ChangeDescription {
                        $ret = EditTeam -GroupId $teamId -Description $newValue
                    
                    } 
                    $script:ChangePrivacy {
                        $ret = EditTeam -GroupId $teamId -Privacy $newValue                    
                    }  
                }
                $reqStatus = $ret["Status"]
                $reqMessage = $ret["Message"]

                LogWrite -Message "$reqMessage"
                LogWrite -Message "Update change request status."
                UpdateChangeRequest -connectionString $script:connectionString -reqObj $request -reqStatus $reqStatus
                if ($reqStatus -eq $script:Completed) {
                    #only update Teams/Groups table if change request is completed
                    LogWrite -Message "Update newly updated team into Groups and Teams table."
                    $team = Get-Team -GroupId $teamId
                    UpdateGroupTeamPostChangeRequest -connectionString $script:connectionString -teamObj $team
                } 
            }
            catch {               
                LogWrite -Level ERROR -Message "[Process-TeamChangeRequests] - Error team change request $teamId : $($_.Exception)"      
            }            
        }
        LogWrite -Level INFO -Message "Processing Team Change Requests completed."        
    }
    catch {
        LogWrite -Level ERROR -Message "[Process-TeamChangeRequests] - Unexpected exception: $($_.Exception)"        
        throw $_.Exception
    }
}

Function Process-GroupChangeRequests {
    param
    (        
        [Parameter(Mandatory = $true)] $Requests
        
    )

    try {
        LogWrite -Level INFO -Message "Processing Group Change Requests."
        LogWrite -Level INFO -Message "Total Group Change Requests: $($Requests.Count)"

        if ($Requests.Count -ile 0) {
            return
        }        

        LogWrite -Message "Getting GRAPH API access token..."        
        $authToken = Connect-GraphAPIWithCert -TenantId $script:TenantId -AppId $script:appIdOperationSupport -Thumbprint $script:appThumbprintOperationSupport

        foreach ($request in $Requests) {
            try {                
                $groupId = $request.GroupId.Trim()
                $newValue = $request.NewValue                
                $requestType = $request.ChangeTypeId.Guid
                LogWrite -Level INFO -Message "Group ID: $groupId"

                switch ($requestType) {
                    $script:ChangeDisplayName {
                        $ret = Set-NIHO365Group -AuthToken $authToken -Id $groupId -DisplayName $newValue
                    }
                    $script:ChangeDescription {
                        $ret = Set-NIHO365Group -AuthToken $authToken -Id $groupId -Description $newValue
                    
                    } 
                    $script:ChangePrivacy {
                        $ret = Set-NIHO365Group -AuthToken $authToken -Id $groupId -Visibility $newValue
                    }  
                }
                $reqStatus = $script:Completed
                $reqMessage = "Update props for M365 Group [$groupId] successfully"

                LogWrite -Message "$reqMessage"
                LogWrite -Message "Update change request status."
                UpdateChangeRequest -connectionString $script:connectionString -reqObj $request -reqStatus $reqStatus
                
                if ($reqStatus -eq $script:Completed) {
                    #only update Teams/Groups table if change request is completed
                    LogWrite -Message "Update newly updated team into Groups and Teams table."
                    $group = Get-NIHO365Group -AuthToken $authToken -Id $groupId
                    UpdateGroupTeamPostChangeRequest -connectionString $script:connectionString -teamObj $group
                } 
            }
            catch {               
                LogWrite -Level ERROR -Message "[Process-GroupChangeRequests] - Error group change request $teamId : $($_.Exception)"      
            }            
        }
        LogWrite -Level INFO -Message "Processing Group Change Requests completed."        
    }
    catch {
        LogWrite -Level ERROR -Message "[Process-GroupChangeRequests] - Unexpected exception: $($_.Exception)"        
        throw $_.Exception
    }
}

Function Process-TeamRenameRequests {
    param
    (        
        [Parameter(Mandatory = $true)] $Requests
        
    )

    try {
        LogWrite -Level INFO -Message "Processing Team Rename Requests."
        LogWrite -Level INFO -Message "Total Team Rename Requests: $($Requests.Count)"

        if ($Requests.Count -ile 0) {
            return
        }        

        foreach ($request in $Requests) {
            try {                
                $teamId = $request.GroupId.Trim()
                $newTeamName = $request.NewValue
                LogWrite -Level INFO -Message "Team ID: $teamId"
                $ret = RenameTeamDisplayName -GroupId $teamId -DisplayName $newTeamName
                $reqStatus = $ret["Status"]
                $reqMessage = $ret["Message"]

                LogWrite -Message "$reqMessage"
                LogWrite -Message "Update change request status."
                UpdateChangeRequest -connectionString $script:connectionString -reqObj $request -reqStatus $reqStatus
                if ($reqStatus -eq $script:Completed) {
                    #only update Teams/Groups table if change request is completed
                    #Verify if Display name is updated against to M365
                    $team = Get-Team -GroupId $teamId
                    if ($team.DisplayName -eq $newTeamName) {
                        LogWrite -Message "Update group display name [$newTeamName] into Groups and Teams table."
                        UpdateGroupTeamPostChangeRequest -connectionString $script:connectionString -teamObj $team
                    }
                    else {
                        LogWrite -Message "Nothing updated."
                    }
                } 
            }
            catch {               
                LogWrite -Level ERROR -Message "[Process-TeamRenameRequests] - Error renaming team $teamId : $($_.Exception)"      
            }            
        }
        LogWrite -Level INFO -Message "Processing Team Rename Requests completed."        
    }
    catch {
        LogWrite -Level ERROR -Message "[Process-TeamRenameRequests] - Unexpected exception: $($_.Exception)"        
        throw $_.Exception
    }
}

Function Process-OwnerAccessRequests {
    param
    (        
        [Parameter(Mandatory = $true)] $Requests
        
    )

    try {
        LogWrite -Level INFO -Message "Processing Owner Access Requests."
        LogWrite -Level INFO -Message "Total Owner Access Requests: $($Requests.Count)"

        if ($Requests.Count -ile 0) {
            return
        }

        LogWrite -Message "Getting GRAPH API access token..."        
        $authToken = Connect-GraphAPIWithCert -TenantId $script:TenantId -AppId $script:appIdOperationSupport -Thumbprint $script:appThumbprintOperationSupport

        foreach ($request in $Requests) {
            try {                
                $groupId = $request.GroupId.Trim()
                $ownerEmail = $request.NewValue.Trim()
                $requestType = $request.ChangeTypeId.Guid                
                $justification = $request.Justification
                $templateId = $request.TemplateId

                #$ownerId = (Get-NIHO365UserByEmail -AuthToken $authToken -EmailAddress $ownerEmail).id
                #$dt = GetSigninNameByEmail -Email $ownerEmail -connectionString $script:ConnectionString
                #$ownerId = $dt["UserId"].Guid 
                  
                $requestorInfo = Get-NIHO365UserByEmail -AuthToken $authToken -EmailAddress $ownerEmail
                $ownerId = $requestorInfo.id
                $ownerName = $requestorInfo.displayName

                LogWrite -Level INFO -Message "Group ID: $groupId - Owner Id: $ownerId"
                $group = Get-NIHO365Group -AuthToken $authToken -Id $GroupId                

                if ($group) {
                    $groupName = $group.displayName
                    $groupAlias = $group.mailNickname
                    LogWrite -Message "Collecting existing owners of this group..."
                    $existingOwners = @(Get-NIHO365GroupOwners -AuthToken $authToken -Id $GroupId)

                    if ($templateId -eq $script:M365Group) {
                        $ret = Add-NIHO365GroupMember -AuthToken $authToken -Group $groupId -Members $ownerId -AsOwner
                        $GroupLink = "https://outlook.office365.com/mail/group/groups.nih.gov/$GroupAlias/email"
                    }
                    if ($templateId -eq $script:MSTeams) {
                        $ret = Add-NIHTeamMember -AuthToken $authToken -Group $GroupId -Member $ownerId -AsOwner
                        $team = Get-NIHTeam -AuthToken $authToken -Id $GroupId
                        if ($team) {
                            $GroupLink = $team.webUrl
                        }
                    }              
                    LogWrite -Message "Update change request status."
                    UpdateChangeRequest -connectionString $script:connectionString -reqObj $request -reqStatus $script:Completed
                    # Email to existing owner inform a new owner was added to their group/team
                    if ($existingOwners.Count -gt 0) {
                        $emails = $existingOwners.mail -join ";"
                        if ($emails) {
                            $subject = "Update to M365 Teams/Group Ownership"
                            $content = "We are notifying you that your IC M365 Teams point of contact has been added as an owner to the following M365 Teams/Group to complete a requested/needed IT update.<br />&nbsp;
                            <table cellpadding='5' cellspacing='2' style='border: none;border-collapse: collapse;width:80%'>
                            <tr><td width='30%'><b>Team/Group Name:</b></td><td><a href='$GroupLink'>$groupName</a></td></tr>
                            <tr><td width='30%'><b>New owner added:</b></td><td>$ownerName</td></tr>
                            <tr><td width='30%'><b>Justification Provided:</b></td>$justification<td></td></tr>
                            </table>
                            <p><b>There is no action for you at this time</b>. If you have any questions, please contact your IC Teams point of contact CC’d on this email.</p>
                            "                            
                            $body = "<p><i>Note: This is an automated email. Please do not reply to this message.</i></p>
                                 <p>Hello M365 Teams/Group Owner(s),</p>
                                 $content
                                 <p><i>Thank you,</i> <br />NIH M365 Collaboration Support Team</p>"    
                            $body = [System.Web.HttpUtility]::HtmlDecode($body)
                            SendEmail -subject $subject -body $body -To $emails -EnabledCc -Cc $ownerEmail   
                        }
                    }
                }

            }
            catch {               
                LogWrite -Level ERROR -Message "[Process-OwnerAccessRequests] - Error owner Access $groupId : $($_.Exception)"
            }            
        }
        LogWrite -Level INFO -Message "Processing Owner Access Requests completed."        
    }
    catch {
        LogWrite -Level ERROR -Message "[Process-OwnerAccessRequests] - Unexpected exception: $($_.Exception)"        
        throw $_.Exception
    }
}

Function Process-ExternalSharingRequests {
    param
    (        
        [Parameter(Mandatory = $true)] $Requests        
    )

    try {
        LogWrite -Level INFO -Message "Processing External Sharing Requests."
        LogWrite -Level INFO -Message "Total External Sharing Requests: $($Requests.Count)"

        if ($Requests.Count -ile 0) {
            return
        }

        $tenantProps = Get-PnPTenant
        $currentShareSettings = $tenantProps.SharingCapability
        LogWrite -Level INFO -Message "Current SharingCapability settings in the tenant: $($currentShareSettings)."
        
        if ($currentShareSettings -eq 'Disabled') {
            LogWrite -Level ERROR -Message "Sharing is currently disabled on the tenant level!"
            return
        }        

        foreach ($request in $Requests) {
            try {                
                $siteUrl = $request.SiteUrl
                $externalSharingEnabled = $request.NewValue
                $newValue = "Disabled"                
                if ($request.NewValue -eq 1) {
                    $newValue = $currentShareSettings
                }
                if ([string]::IsNullOrWhiteSpace($siteUrl)) {
                    LogWrite -Level ERROR -Message "[Site URL cannot be null]: $($request.ChangeRequestTypeId)"
                    throw "Site URL cannot be null."
                }
                $siteUrl = $siteUrl.Trim()

                #--Update ExternalSharingEnable to Sites table
                LogWrite -Message "Update [ExternalSharingEnable] field in Sites table"                    
                UpdateSPOSiteExternalSharingRecord $script:connectionString -SiteUrl $siteUrl -ExternalSharingEnabled $externalSharingEnabled
                #--Update external sharing for SPO site                    
                $ret = UpdateExternalSharing -SiteUrl $siteUrl -SharingCapability $newValue
                $reqStatus = $ret["Status"]
                $reqMessage = $ret["Message"]
                LogWrite -Message "$reqMessage"                
                UpdateChangeRequest -connectionString $script:connectionString -reqObj $request -reqStatus $reqStatus
                
                if ($reqStatus -eq $script:Completed) {
                    #only update sites table if change request is completed
                    #Verify SharingCapability before update to DB                    
                    $site = Get-PnPTenantSite -Url $siteUrl
                    $retryCount = 5
                    $retryAttempts = 0
                    while ($site.Status -ne 'Active' -and $retryAttempts -lt $retryCount) {
                        LogWrite -Message "Waiting until the site is updated..."
                        Sleep 15
                        $retryAttempts = $retryAttempts + 1
                        $site = Get-PnPTenantSite -Url $siteUrl
                    }
                    LogWrite -Message "Update [SharingCapability] to the Sites table"                    
                    UpdateSPOSiteExternalSharingRecord -ConnectionString $script:connectionString -SiteUrl $siteUrl -SiteObj $site -ExternalSharingEnabled $externalSharingEnabled
                } 

            }
            catch {                
                LogWrite -Level ERROR -Message "[Process-ExternalSharingRequests] - Error external sharing $SiteUrl : $($_.Exception)"      
            }                  
        }
        LogWrite -Level INFO -Message "Processing external sharing completed."        
    }
    catch {
        LogWrite -Level ERROR -Message "[Process-ExternalSharingRequests] - Unexpected exception: $($_.Exception)"        
        throw $_.Exception
    }
}

Function Process-CustomizationRequests {
    param
    (        
        [Parameter(Mandatory = $true)] $Requests        
    )

    try {
        LogWrite -Level INFO -Message "Processing Site Customization Requests."
        LogWrite -Level INFO -Message "Total Site Customization Requests: $($Requests.Count)"

        if ($Requests.Count -ile 0) {
            return
        }

        foreach ($request in $Requests) {
            try {                
                $siteUrl = $request.SiteUrl
                $newValue = $false
                if ($request.NewValue -eq 0) {
                    $newValue = $true
                }
               
                if ([string]::IsNullOrWhiteSpace($siteUrl)) {
                    LogWrite -Level ERROR -Message "[Site URL cannot be null]: $($request.ChangeRequestTypeId)"
                    throw "Site URL cannot be null."
                }
                $siteUrl = $siteUrl.Trim()
                $ret = SiteCustomization -SiteUrl $siteUrl -NoScriptSite $newValue
                $reqStatus = $ret["Status"]
                $reqMessage = $ret["Message"]
                LogWrite -Message "$reqMessage"                
                UpdateChangeRequest -connectionString $script:connectionString -reqObj $request -reqStatus $reqStatus
                
                if ($reqStatus -eq $script:Completed) {
                    #only update sites table if change request is completed
                    #Verify DenyAddAndCustomizePages before update to DB                    
                    $site = Get-PnPTenantSite -Url $siteUrl
                    $retryCount = 5
                    $retryAttempts = 0
                    while ($site.Status -ne 'Active' -and $retryAttempts -lt $retryCount) {
                        LogWrite -Message "Waiting until the site is updated..."
                        Sleep 15
                        $retryAttempts = $retryAttempts + 1
                        $site = Get-PnPTenantSite -Url $siteUrl
                    }
                    LogWrite -Message "Update [DenyAddAndCustomizePages] to the Sites table"                    
                    UpdateSPOSiteExternalSharingRecord -ConnectionString $script:connectionString -SiteUrl $siteUrl -SiteObj $site
                } 

            }
            catch {                
                LogWrite -Level ERROR -Message "[Process-CustomizationRequests] - Error allow custom script for the site $SiteUrl : $($_.Exception)"      
            }                  
        }
        LogWrite -Level INFO -Message "Processing Site Customization completed."        
    }
    catch {
        LogWrite -Level ERROR -Message "[Process-CustomizationRequests] - Unexpected exception: $($_.Exception)"        
        throw $_.Exception
    }
}

Function Process-HubSiteRequests {
    param
    (        
        [Parameter(Mandatory = $true)] $Requests        
    )

    try {
        LogWrite -Level INFO -Message "Processing Register Hub Site Requests."
        LogWrite -Level INFO -Message "Total RegisterHubSite Requests: $($Requests.Count)"

        if ($Requests.Count -ile 0) {
            return
        }

        foreach ($request in $Requests) {
            try {                
                $siteUrl = $request.SiteUrl
                $newValue = $request.NewValue
                $owner = $request.CreatedBy
               
                if ([string]::IsNullOrWhiteSpace($siteUrl)) {
                    LogWrite -Level ERROR -Message "[Site URL cannot be null]: $($request.ChangeRequestTypeId)"
                    throw "Site URL cannot be null."
                }
                $siteUrl = $siteUrl.Trim()
                $ret = RegisterHubSite -SiteUrl $siteUrl -HubName $newValue #-Owners $owner
                $reqStatus = $ret["Status"]
                $reqMessage = $ret["Message"]
                LogWrite -Message "$reqMessage"                
                UpdateChangeRequest -connectionString $script:connectionString -reqObj $request -reqStatus $reqStatus
                
                if ($reqStatus -eq $script:Completed) {
                    #only update sites table if change request is completed                    
                    $site = Get-PnPTenantSite -Url $siteUrl
                    $retryCount = 5
                    $retryAttempts = 0
                    while ($site.Status -ne 'Active' -and $retryAttempts -lt $retryCount) {
                        LogWrite -Message "Waiting until the site is updated..."
                        Sleep 15
                        $retryAttempts = $retryAttempts + 1
                        $site = Get-PnPTenantSite -Url $siteUrl
                    }
                    LogWrite -Message "Update [HubName] to the Sites table"
                    if ($site.IsHubSite -eq $true) {
                        $hubName = (Get-PnPHubSite -Identity $siteUrl).Title
                    }
                    UpdateSPOSiteExternalSharingRecord -ConnectionString $script:connectionString -SiteUrl $siteUrl -SiteObj $site -HubName $hubName
                } 

            }
            catch {                
                LogWrite -Level ERROR -Message "[Process-RegisterHubSiteRequests] - Error register hubsite for the site $SiteUrl : $($_.Exception)"      
            }                  
        }
        LogWrite -Level INFO -Message "Processing Register Hub Site Requests completed."        
    }
    catch {
        LogWrite -Level ERROR -Message "[Process-RegisterHubSiteRequests] - Unexpected exception: $($_.Exception)"        
        throw $_.Exception
    }
}

Function Process-StorageQuotaRequests {
    param
    (        
        [Parameter(Mandatory = $true)] $Requests        
    )

    try {
        LogWrite -Level INFO -Message "Processing Update Storage Quota Requests."
        LogWrite -Level INFO -Message "Total Update Storage Quota Requests: $($Requests.Count)"

        if ($Requests.Count -ile 0) {
            return
        }

        foreach ($request in $Requests) {
            try {                
                $siteUrl = $request.SiteUrl
                $newValue = $request.NewValue
                
                if ([string]::IsNullOrWhiteSpace($siteUrl)) {
                    LogWrite -Level ERROR -Message "[Site URL cannot be null]: $($request.ChangeRequestTypeId)"
                    throw "Site URL cannot be null."
                }

                $siteUrl = $siteUrl.Trim()
                $ret = UpdateStorageQuota -SiteUrl $siteUrl -Quota $newValue
                $reqStatus = $ret["Status"]
                $reqMessage = $ret["Message"]
                LogWrite -Message "$reqMessage"                
                UpdateChangeRequest -connectionString $script:connectionString -reqObj $request -reqStatus $reqStatus
                
                if ($reqStatus -eq $script:Completed) {
                    #only update sites table if change request is completed                    
                    $site = Get-PnPTenantSite -Url $siteUrl
                    $retryCount = 5
                    $retryAttempts = 0
                    while ($site.Status -ne 'Active' -and $retryAttempts -lt $retryCount) {
                        LogWrite -Message "Waiting until the site is updated..."
                        Sleep 15
                        $retryAttempts = $retryAttempts + 1
                        $site = Get-PnPTenantSite -Url $siteUrl
                    }
                    
                    UpdateSPOSiteExternalSharingRecord -ConnectionString $script:connectionString -SiteUrl $siteUrl -SiteObj $site
                } 

            }
            catch {                
                LogWrite -Level ERROR -Message "[Process-UpdateStorageQuotaRequests] - Error Update Storage Quota for the site $SiteUrl : $($_.Exception)"      
            }                  
        }
        LogWrite -Level INFO -Message "Processing Update Storage Quota completed."        
    }
    catch {
        LogWrite -Level ERROR -Message "[Process-UpdateStorageQuotaRequests] - Unexpected exception: $($_.Exception)"        
        throw $_.Exception
    }
}

Function Process-AppCatalogRequests {
    param
    (        
        [Parameter(Mandatory = $true)] $Requests        
    )

    try {
        LogWrite -Level INFO -Message "Processing Enabled/Disabled App catalog Requests."
        LogWrite -Level INFO -Message "Total Enabled/Disabled App catalog Requests: $($Requests.Count)"

        if ($Requests.Count -ile 0) {
            return
        }

        foreach ($request in $Requests) {
            try {                
                $siteUrl = $request.SiteUrl
                $newValue = $request.NewValue
                
                if ([string]::IsNullOrWhiteSpace($siteUrl)) {
                    LogWrite -Level ERROR -Message "[Site URL cannot be null]: $($request.ChangeRequestTypeId)"
                    throw "Site URL cannot be null."
                }

                $siteUrl = $siteUrl.Trim()
                $ret = AddRemoveAppCatalog -SiteUrl $siteUrl -AppCatalog $newValue
                $reqStatus = $ret["Status"]
                $reqMessage = $ret["Message"]
                LogWrite -Message "$reqMessage"                
                UpdateChangeRequest -connectionString $script:connectionString -reqObj $request -reqStatus $reqStatus
                
                if ($reqStatus -eq $script:Completed) {
                    #only update sites table if change request is completed                    
                    $site = Get-PnPTenantSite -Url $siteUrl
                    $retryCount = 5
                    $retryAttempts = 0
                    while ($site.Status -ne 'Active' -and $retryAttempts -lt $retryCount) {
                        LogWrite -Message "Waiting until the site is updated..."
                        Sleep 15
                        $retryAttempts = $retryAttempts + 1
                        $site = Get-PnPTenantSite -Url $siteUrl
                    }
                    
                    UpdateSPOSiteExternalSharingRecord -ConnectionString $script:connectionString -SiteUrl $siteUrl -SiteObj $site -AppCatalogEnabled $newValue
                } 

            }
            catch {                
                LogWrite -Level ERROR -Message "[Process-AppCatalogRequests] - Error Enabled/Disabled App catalog for the site $SiteUrl : $($_.Exception)"      
            }                  
        }
        LogWrite -Level INFO -Message "Processing Enabled/Disabled App catalog completed."        
    }
    catch {
        LogWrite -Level ERROR -Message "[Process-AppCatalogRequests] - Unexpected exception: $($_.Exception)"        
        throw $_.Exception
    }
}

Function Process-ConnectM365GroupRequests {
    <#
        .DESCRIPTION
            It is not supported to connect a Communication site to Microsoft 365 group. (COMMUNITY#0 and COMMUNITYPORTAL#0)
            Cannot group connect the root site collection
            Publishing feature enabled...can't group connect this site
            Incompatible web template detected...can't group connect this site
            -BICENTERSITE#0"
            -BLANKINTERNET#0
            -ENTERWIKI#0
            -SRCHCEN#0
            -SRCHCENTERLITE#0
            -POINTPUBLISHINGHUB#0
            -POINTPUBLISHINGTOPIC#0
            -$siteCollectionUrl.EndsWith("/sites/contenttypehub"))
    #>
    param
    (        
        [Parameter(Mandatory = $true)] $Requests        
    )
    try {
        LogWrite -Level INFO -Message "Processing Connect M365Group to sharepoint site Requests."
        LogWrite -Level INFO -Message "Total Connect M365Group to sharepoint site Requests: $($Requests.Count)"

        if ($Requests.Count -ile 0) {
            return
        }

        foreach ($request in $Requests) {
            try {                
                $siteUrl = $request.SiteUrl
                $newValue = $request.NewValue
                $owner = $request.CreatedBy
                
                if ([string]::IsNullOrWhiteSpace($siteUrl)) {
                    LogWrite -Level ERROR -Message "[Site URL cannot be null]: $($request.ChangeRequestTypeId)"
                    throw "Site URL cannot be null."
                }

                $siteUrl = $siteUrl.Trim()
                
                $ret = ConnectM365Group -SiteUrl $siteUrl -Owner $owner
                $reqStatus = $ret["Status"]
                $reqMessage = $ret["Message"]
                LogWrite -Message "$reqMessage"                
                UpdateChangeRequest -connectionString $script:connectionString -reqObj $request -reqStatus $reqStatus
                
                <#
                if ($reqStatus -eq $script:Completed){ #only update sites table if change request is completed                    
                    $site = Get-PnPTenantSite -Url $siteUrl
                    $retryCount = 5
                    $retryAttempts = 0
                    while ($site.Status -ne 'Active' -and $retryAttempts -lt $retryCount) {
                        LogWrite -Message "Waiting until the site is updated..."
                        Sleep 15
                        $retryAttempts = $retryAttempts + 1
                        $site = Get-PnPTenantSite -Url $siteUrl
                    }
                    
                    UpdateSPOSiteExternalSharingRecord -ConnectionString $script:connectionString -SiteUrl $siteUrl -SiteObj $site -AppCatalogEnabled $newValue
                } 
                #>
            }
            catch {                
                LogWrite -Level ERROR -Message "[Process-ConnectM365GroupRequests] - Error Connect M365Group to the site $SiteUrl : $($_.Exception)"      
            }                  
        }
        LogWrite -Level INFO -Message "Processing Enabled/Disabled App catalog completed."        
    }
    catch {
        LogWrite -Level ERROR -Message "[Process-ConnectM365GroupRequests] - Unexpected exception: $($_.Exception)"        
        throw $_.Exception
    }
}

Function Process-GroupHiddenGALRequests {
    param
    (        
        [Parameter(Mandatory = $true)] $Requests
        
    )

    try {
        LogWrite -Level INFO -Message "Processing Hidden Outlook/GAL Change Requests."
        LogWrite -Level INFO -Message "Total Hidden Outlook/GAL Change Requests: $($Requests.Count)"

        if ($Requests.Count -ile 0) {
            return
        }
                
        LogWrite -Message "Getting GRAPH API access token..."        
        $authToken = Connect-GraphAPIWithCert -TenantId $script:TenantId -AppId $script:appIdOperationSupport -Thumbprint $script:appThumbprintOperationSupport

        foreach ($request in $Requests) {
            try {                
                $teamId = $request.GroupId.Trim()                
                $newValue = $false
                if ($request.NewValue -eq 1) {
                    $newValue = $true
                }                
                LogWrite -Level INFO -Message "Team ID: $teamId"
                Update-NIHGroupSettings -AuthToken $authToken -Id $teamId -HideFromOutlookClients $newValue -HideFromAddressLists $newValue                                    
                <#                
                $retryCount = 5
                $retryAttempts = 0
                $backOffInterval = 2 

                $extendProps = Get-NIHO365Group -AuthToken $authToken -Id $teamId -Select displayName,description,visibility,hideFromAddressLists, hideFromOutlookClients
                while ($extendProps.hideFromAddressLists -ne $newValue -and $extendProps.hideFromOutlookClients -ne $newValue -and $retryAttempts -lt $retryCount) {
                    LogWrite -Message "Waiting until the group setting is updated..."
                    Start-Sleep $backOffInterval
                    $backOffInterval = $backOffInterval * 2
                    $retryAttempts = $retryAttempts + 1
                    $extendProps = Get-NIHO365Group -AuthToken $authToken -Id $teamId -Select displayName,description,visibility,hideFromAddressLists, hideFromOutlookClients
                }

                if ($extendProps.hideFromAddressLists -eq $newValue -and $extendProps.hideFromOutlookClients -eq $newValue){
                    $reqStatus = $script:Completed
                    $reqMessage = "Update Outlook/GAL props for M365 Group [$teamId] successfully"

                    LogWrite -Message "$reqMessage"
                    LogWrite -Message "Update change request status."
                    UpdateChangeRequest -connectionString $script:connectionString -reqObj $request -reqStatus $reqStatus
                    
                    LogWrite -Message "Update hideFromAddressLists, hideFromOutlookClients to Groups table."
                    UpdateGroupTeamPostChangeRequest -connectionString $script:connectionString -teamObj $extendProps

                }
                #>
                $reqStatus = $script:Completed
                $reqMessage = "Update Outlook/GAL props for M365 Group [$teamId] successfully"

                LogWrite -Message "$reqMessage"
                LogWrite -Message "Update change request status."
                UpdateChangeRequest -connectionString $script:connectionString -reqObj $request -reqStatus $reqStatus

                if ($reqStatus -eq $script:Completed) {
                    #only update Teams/Groups table if change request is completed
                    LogWrite -Message "Update hideFromAddressLists, hideFromOutlookClients to Groups table."                    
                    #$extendProps = Get-NIHO365Group -AuthToken $authToken -Id $teamId -Select displayName,description,visibility,hideFromAddressLists, hideFromOutlookClients
                    $extendProps = Get-NIHO365Group -AuthToken $authToken -Id $teamId
                    UpdateGroupTeamPostChangeRequest -connectionString $script:connectionString -teamObj $extendProps -HideFromOutlookClients $newValue -HideFromAddressLists $newValue
                    #UpdateGroupTeamPostRename -connectionString $script:connectionString -teamObj $team
                } 
                
            }
            catch {               
                LogWrite -Level ERROR -Message "[Process-GroupHiddenGALRequests] - Error Hidden Outlook/GAL Change Requests $teamId : $($_.Exception)"      
            }            
        }
        LogWrite -Level INFO -Message "Processing Hidden Outlook/GAL Change Requests completed."        
    }
    catch {
        LogWrite -Level ERROR -Message "[Process-GroupHiddenGALRequests] - Unexpected exception: $($_.Exception)"        
        throw $_.Exception
    }
}

Function Process-SiteAccessRequests {
    param
    (        
        [Parameter(Mandatory = $true)] $Requests
        
    )

    try {
        LogWrite -Level INFO -Message "Processing Site Admin Access Requests."
        LogWrite -Level INFO -Message "Total Site Admin Access Requests: $($Requests.Count)"

        if ($Requests.Count -ile 0) {
            return
        }
        

        foreach ($request in $Requests) {
            try { 

                $siteUrl = $request.SiteUrl
                $siteName = $request.SiteName
                $newValue = $request.NewValue               
                
                if ([string]::IsNullOrWhiteSpace($siteUrl) -or [string]::IsNullOrWhiteSpace($newValue)) {
                    LogWrite -Level ERROR -Message "Site URL or Owners cannot be null"
                    throw "Site URL or owners cannot be nul."
                }
                
                $siteUrl = $siteUrl.Trim()
                $ownerEmail = $request.NewValue.Trim()
                $owners = $newValue.Trim() -join ","
                $requestType = $request.ChangeTypeId.Guid                
                $justification = $request.Justification
                $templateId = $request.TemplateId
                
                $ret = AddSiteAdmin -SiteUrl $siteUrl -Owners $ownerEmail
                $reqStatus = $ret["Status"]
                $reqMessage = $ret["Message"]
                LogWrite -Message "$reqMessage"
                LogWrite -Message "Update change request status."                
                UpdateChangeRequest -connectionString $script:connectionString -reqObj $request -reqStatus $reqStatus

                if ($reqStatus -eq $script:Completed) {
                    # Email to existing site owner inform a new site admin was added to their site 
                    try{                   
                        $siteContext = ConnectPnpOnlineOAuth -TenantId $script:TenantId -ClientId $script:appIdOperationSupport -Thumbprint $script:appThumbprintOperationSupport -Url $siteUrl                         
                        $ownersGroupName = (Get-PnPGroup -AssociatedOwnerGroup).Title
                        #$siteOwners = Get-PnPGroupMember -Group $ownersGroupName | Where-Object { $_.Title -ne "System Account" -and $_.PrincipalType -eq "User" }  | Select-Object LoginName, Title, Email, PrincipalType
                        $siteOwners = @(Get-PnPGroupMember -Group $ownersGroupName | Where-Object { $_.Title -ne "System Account" -and $_.PrincipalType -eq "User" }).Email
                        #$siteAdmins = @(Get-PnPSiteCollectionAdmin -Connection $siteContext).Email
                        if ($siteOwners.Count -le 0){
                            $siteOwners = @(Get-PnPSiteCollectionAdmin -Connection $siteContext).Email
                            #$siteOwners = ($admins).Email -join ";"                            
                        }

                        if ($siteOwners.Count -gt 0) {
                            LogWrite -Level INFO -Message "Sending email to site owners/site admins..."
                            $emails = $siteOwners -join ";"
                            if ($emails) {
                                $subject = "Update to site Ownership"
                                $content = "We are notifying you that your IC Admin point of contact has been added as an admin to the following site to complete a requested/needed IT update.<br />&nbsp;
                                <table cellpadding='5' cellspacing='2' style='border: none;border-collapse: collapse;width:80%'>
                                <tr><td width='30%'><b>Site Name:</b></td><td><a href='$siteUrl'>$siteName</a></td></tr>
                                <tr><td width='30%'><b>New site admin added:</b></td><td>$newValue</td></tr>
                                <tr><td width='30%'><b>Justification Provided:</b></td>$justification<td></td></tr>
                                </table>
                                <p><b>There is no action for you at this time</b>. If you have any questions, please contact your IC Admin point of contact CC’d on this email.</p>
                                "                            
                                $body = "<p><i>Note: This is an automated email. Please do not reply to this message.</i></p>
                                        <p>Hello Site Owner(s),</p>
                                        $content
                                        <p><i>Thank you,</i> <br />NIH M365 Collaboration Support Team</p>"    
                                $body = [System.Web.HttpUtility]::HtmlDecode($body)
                                SendEmail -subject $subject -body $body -To $emails -EnabledCc -Cc $ownerEmail   
                            }
                        }
                        # Update to sites table or let sync job handle?
                        # UpdateSPOSiteExternalSharingRecord -ConnectionString $script:connectionString -SiteUrl $siteUrl -SiteObj $site -AppCatalogEnabled $newValue
                    }
                    catch{
                        LogWrite -Level ERROR -Message "[Process-SiteAccessRequests] - Error while sending email to site admin $($_.Exception)"
                    }
                    <#                    
                    finally {
                        LogWrite -Level INFO -Message "Disconnect sharepoint site."        
                        if ($siteContext) {
                            DisconnectPnpOnlineOAuth -Context $siteContext
                        }        
                    } #> 
                    
                } 
            }
            catch {               
                LogWrite -Level ERROR -Message "[Process-SiteAccessRequests] - Error site admin access $($_.Exception)"
            }                      
        }
        LogWrite -Level INFO -Message "Processing Site Admin Access Requests completed."        
    }
    catch {
        LogWrite -Level ERROR -Message "[Process-SiteAccessRequests] - Unexpected exception: $($_.Exception)"        
        throw $_.Exception
    }   
}


try {
    #-------- Set Global Variables ---------
    Set-TenantVars
    Set-AzureAppVars
    Set-DBVars   
    Set-MiscVars
    Set-LogFile -logFileName $logFileName
    Set-ChangeTypeVars
    Set-StatusVars

    #-------- Set Global Variables Ended ---------    
    
    $startTime = Get-Date -Format "yyyy-MM-dd HH:mm:ss"           
    LogWrite -Message "----------------------- [Process M365 Operations] Execution Started -----------------------"    
    #--- Get all requests from DB-[ChangeRequests] with Status either Submitted or Pending (mark as Pending if something go wrong) ---
    $requests = GetActiveChangeRequests $script:connectionString
    #filter by ChangeTypeId
    $externalSharingRequests = @($requests | ? { $_.ChangeTypeId -eq $script:ExternalSharing })
    $enableAppCatalogRequests = @($requests | ? { $_.ChangeTypeId -eq $script:EnableAppCatalog })
    $enableSiteCustomizationRequests = @($requests | ? { $_.ChangeTypeId -eq $script:EnableSiteCustomization })
    $registerHubSiteRequests = @($requests | ? { $_.ChangeTypeId -eq $script:RegisterHubSite })    
    $siteAdminRequests = @($requests | ? { $_.ChangeTypeId -eq $script:SiteAdminAccess })    
    
    <#$teamChangeRequests = @($requests | ? { $_.ChangeTypeId -eq $script:TeamsDisplayName  -or $_.ChangeTypeId -eq $script:TeamsDescription  -or $_.ChangeTypeId -eq $script:TeamsPrivacy})
    $groupChangeRequests = @($requests | ? { $_.ChangeTypeId -eq $script:GroupDisplayName  -or $_.ChangeTypeId -eq $script:GroupDescription  -or $_.ChangeTypeId -eq $script:GroupPrivacy})
    $ownerAccessRequests = @($requests | ? { $_.ChangeTypeId -eq $script:TeamsOwnerAccess -or $_.ChangeTypeId -eq $script:GroupOwnerAccess})
    #>

    $nameChangeRequests = @($requests | ? { $_.ChangeTypeId -eq $script:ChangeDisplayName -or $_.ChangeTypeId -eq $script:ChangeDescription -or $_.ChangeTypeId -eq $script:ChangePrivacy })
    $ownerAccessRequests = @($requests | ? { $_.ChangeTypeId -eq $script:OwnerAccess })
    $hiddenOutlookGALRequests = @($requests | ? { $_.ChangeTypeId -eq $script:HiddenOutlookGAL })       

    #--- PnP ---  
    LogWrite -Message "Connecting to SharePoint Admin Center '$($script:SPOAdminCenterURL)'..."
    $script:TenantContext = ConnectPnpOnlineOAuth -TenantId $script:TenantId -ClientId $script:appIdOperationSupport -Thumbprint $script:appThumbprintOperationSupport -Url $script:SPOAdminCenterURL  
    LogWrite -Message "SharePoint Admin Center '$($script:SPOAdminCenterURL)' is now connected."         

    if ($siteAdminRequests) {        
        Process-SiteAccessRequests $siteAdminRequests
    }

    if ($externalSharingRequests) {                
        Process-ExternalSharingRequests $externalSharingRequests
    }
    if ($enableAppCatalogRequests) {        
        Process-AppCatalogRequests $enableAppCatalogRequests
    }
    if ($enableSiteCustomizationRequests) {
        Process-CustomizationRequests $enableSiteCustomizationRequests
    }
    if ($registerHubSiteRequests) {        
        Process-HubSiteRequests $registerHubSiteRequests
    }
    if ($registerHubSiteRequests) {        
        Process-HubSiteRequests $registerHubSiteRequests
    }
    
    #--- MS Teams Operation using Teams API ---            
    <#if ($nameChangeRequests -and $nameChangeRequests.Count -gt 0){
        LogWrite -Message "Connecting Microsoft Teams..." 
        ConnectMicrosoftTeams -TenantId $script:TenantId -ClientId $script:appIdOperationSupport -Thumbprint $script:appThumbprintOperationSupport
        #Connect-MicrosoftTeams -TenantId $script:TenantId -ApplicationId $script:appIdOperationSupport -CertificateThumbprint $script:appThumbprintOperationSupport | Out-Null        
        Process-TeamChangeRequests $nameChangeRequests        
    }#>

    #--- M365 Group Operation using Graph API ---            
    if ($nameChangeRequests -and $nameChangeRequests.Count -gt 0) {        
        Process-GroupChangeRequests $nameChangeRequests        
    }
    
    #--- Owner Access Operation using Graph API --- 
    if ($ownerAccessRequests -and $ownerAccessRequests.Count -gt 0) {                
        Process-OwnerAccessRequests $ownerAccessRequests        
    }
    
    #--- Hidden Outlook/GAL Requests Operation using Graph API ---
    if ($hiddenOutlookGALRequests -and $hiddenOutlookGALRequests.Count -gt 0) {
        Process-GroupHiddenGALRequests $hiddenOutlookGALRequests        
    }

    $endTime = Get-Date -Format "yyyy-MM-dd HH:mm:ss"    
    LogWrite -Message "[Process M365 Operations] Start Time: $startTime"
    LogWrite -Message "[Process M365 Operations] End Time:   $endTime"
    LogWrite -Message  "----------------------- [Process M365 Operations] Execution Ended ------------------------"  
    
}
Catch [Exception] {
    LogWrite -Level ERROR "-Unexpected Error: $_ "
}
Finally {
    LogWrite -Level INFO -Message "Disconnect SharePoint Admin Center."        
    if ($script:TenantContext) {
        DisconnectPnpOnlineOAuth -Context $script:TenantContext
    }
    #LogWrite -Level INFO -Message "Disconnect Microsoft Teams."
    #DisconnectMicrosoftTeams
    LogWrite -Message  "----------------------- [Process M365 Operations] Completed ------------------------"
}

