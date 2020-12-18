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
."$script:RootDir\Common\Lib\LibSPOSitesDAO.ps1"

Enum RequestStatus { 
    New = 0
    #Inprogress = 1
    Completed = 1
    Error = 2
}

function PnpUpdateListItem {
    param
    (
        [Parameter(Mandatory=$true)] $ListName,
        [Parameter(Mandatory=$true)] $SiteConnection,
        [Parameter(Mandatory=$true)] $ItemId,
        [Parameter(Mandatory=$true)] $Values
    )
    try
    {
        #LogWrite -Level INFO -Message "Updating list item. ListName: $($ListName), ItemID: $($ItemId['ID'])"
        if ($SiteConnection -eq $null){
            $SiteConnection = ConnectPnpOnlineOAuth -TenantId $script:TenantId -ClientId $script:appIdOperationSupport -Thumbprint $script:appThumbprintOperationSupport -Url $script:Url
        }
        $result = Set-PnPListItem -List $listName -Identity $itemId -Values $values -Connection $SiteConnection
    }
    catch
    {
        #LogWrite -Level INFO -Message "Error updating list item. Error info: $($_)"
        throw $_
    }
}

function Process-TeamRenameRequests {
    param
    (        
        [Parameter(Mandatory = $true)] $Requests        
    )

    try {
        LogWrite -Level INFO -Message "Processing Team Rename Requests."
        LogWrite -Level INFO -Message "Total Team Rename Requests: $($Requests.Count)."

        if ($Requests.Count -ile 0) {
            return
        }        

        foreach ($request in $Requests) {
            try {              
                $values=@{
                    $script:StatusColumn = ([RequestStatus]::Completed).ToString(); 
                    $script:OperationsHistorColumn = "Team renamed successfully.";
                    }
                $teamId = $request[$script:TeamIdColumn].Trim()              
                $newTeamName = $request[$script:NewTeamsDisplayNameColumn]

                LogWrite -Level INFO -Message "Team ID: $($teamId)."
                $team = Get-Team -GroupId $teamId
                if ($team -eq $null) {
                    LogWrite -Level ERROR -Message "Team not found."                    
                    $values=@{
                        $script:StatusColumn = ([RequestStatus]::Error).ToString();
                        $script:OperationsHistorColumn = "Team not found.";
                    }
                }
                else {
                    LogWrite -Level INFO -Message "Team current name: $($team.DisplayName)."
                    if ($team.DisplayName -ne $newTeamName) {                    
                        Set-Team -GroupId $teamId -DisplayName $newTeamName
                        LogWrite -Level INFO -Message "Team new name: $newTeamName."
                    }  
                    else {
                        LogWrite -Level INFO -Message "Current and new Team names are same. No change will be made."                        
                        $values=@{
                            $script:StatusColumn = ([RequestStatus]::Completed).ToString();
                            $script:OperationsHistorColumn = "Current and new Team names are same. No change will be made.";  
                        }
                    }
                }              
            }
            catch {               
                $values=@{
                    $script:StatusColumn = ([RequestStatus]::Error).ToString();
                    $script:OperationsHistorColumn = "An error occurred - rename team.$($_.Exception.Message)";  
                   }
                LogWrite -Level ERROR -Message "Process-TeamRenameRequests - Error renaming team $teamId : $($_.Exception)"      
            }    
            finally { 
                LogWrite -Level INFO -Message "Update request status."                               
                PnpUpdateListItem -ListName $script:ListId -ItemId $request.Id -Values $values -SiteConnection $script:listContext
            }        
        }
        LogWrite -Level INFO -Message "Processing Team Rename Requests completed."        
    }
    catch {
        LogWrite -Level ERROR -Message "Process-TeamRenameRequests - Unexpected exception: $($_.Exception)"        
        throw $_.Exception
    }
}

function Process-ConnectO365GroupRequests {
    param
    (        
        [Parameter(Mandatory = $true)] $Requests        
    )

    try {
        LogWrite -Level INFO -Message "Processing Connect O365Group Requests."
        LogWrite -Level INFO -Message "Total Connect O365Group Requests: $($Requests.Count)."

        if ($Requests.Count -ile 0) {
            return
        }

        foreach ($request in $Requests) {
            #First check site URL is not empty
            $siteUrl = $request[$script:SiteURLColumn].TrimEnd("/");
            if ($siteUrl -ne "") {
                try {
                    $properties = @{
                        $script:OperationsHistorColumn = "Connect O365Group successfully.";                     
                        $script:StatusColumn           = ([RequestStatus]::Completed).ToString();
                    }
                    $groupAlias =$siteUrl.substring($siteUrl.lastIndexof("/")+1)
                    $groupAlias =$groupAlias.Replace(" ","")
                    if ($groupAlias.length -gt 64){
                        $groupAlias = $groupAlias.substring(0,64)
                    }
                    $siteOwner = $request["Author"].Email
                    #region site Owner
                    <#Email       : levq@NIHDev.cit.nih.gov
                    TypeId      : {c956ab54-16bd-4c18-89d2-996f57282a6f}
                    LookupId    : 60
                    LookupValue : Le, Dan (NIH/CIT) [E]#>
                    #endregion
                    
                    # First validation groupAlias
                    if ($groupAlias.IndexOf(" ") -gt 0) {
                        $message = "$siteUrl : Alias [$groupAlias] contains a space, which not allowed"                    
                        LogWrite -Level ERROR -Message $message                    
                    }
                    else {
                        # try getting the site
                        #$site = Get-PnPTenantSite -Url $siteUrl -Connection $script:pnpTenantContext -ErrorAction Ignore
                        $site = Get-PnPTenantSite -Url $siteUrl
                
                        if ($site.Status -eq "Active") {
                            #$aliasIsUsed = Test-PnPOffice365GroupAliasIsUsed -Alias $groupAlias -Connection $script:pnpTenantContext     
                            $aliasIsUsed = Test-PnPOffice365GroupAliasIsUsed -Alias $groupAlias
                            if ($aliasIsUsed) {
                                $message = "$siteUrl : Alias [$groupAlias] is already in use"
                                LogWrite -Level ERROR -Message $message
                                $properties = @{
                                    $script:OperationsHistorColumn = "Alias [$groupAlias] is already in use.";                     
                                    $script:StatusColumn           = ([RequestStatus]::Error).ToString();
                                }
                            }
                            else {
                                LogWrite -Level INFO -Message "Connect O365Group: $($siteURL)." 
                                $siteContext = ConnectPnpOnlineOAuth -TenantId $script:TenantId -ClientId $script:appIdOperationSupport -Thumbprint $script:appThumbprintOperationSupport -Url $siteUrl
                                Add-PnPOffice365GroupToSite -Url $siteURL -Alias $groupAlias -DisplayName $groupAlias -IsPublic:$false -KeepOldHomePage:$false -Owners $siteOwner -Connection $siteContext
                            }                            
                        }
                        else {
                            $message = "$siteUrl : Site does not exist or is not available (status = $($site.Status))"
                            LogWrite -Level ERROR -Message $message
                            $properties = @{
                                $script:OperationsHistorColumn = "Site does not exist or is not available (status = $($site.Status)";                     
                                $script:StatusColumn           = ([RequestStatus]::Error).ToString();
                            } 
                        } 
                    }
                          
                }
                catch {
                    $properties = @{
                        $script:OperationsHistorColumn = "An error occurred - Connect O365Group.$($_.Exception.Message)";                     
                        $script:StatusColumn           = ([RequestStatus]::Error).ToString();
                    }
                    LogWrite -Level ERROR -Message "Process-ConnectO365GroupRequests - Connect O365Group $SiteUrl : $($_.Exception)"      
                }    
                finally {
                    LogWrite -Level INFO -Message "Update request status."                    
                    PnpUpdateListItem -ListName $script:ListId -ItemId $request.Id -Values $properties -SiteConnection $script:listContext                    
                    #DisconnectPnpOnlineOAuth -Context $siteContext
                }
            }        
        }
        LogWrite -Level INFO -Message "Processing Connect O365Group completed."
    }
    catch {
        LogWrite -Level ERROR -Message "Process-ConnectO365GroupRequests - Unexpected exception: $($_.Exception)"        
        throw $_.Exception
    }
}

function Process-AppCatalogRequests {
    param
    (        
        [Parameter(Mandatory = $true)] $Requests        
    )

    try {
        LogWrite -Message "Processing App catalog Requests."
        LogWrite -Message "Total App catalog Requests: $($Requests.Count)."

        if ($Requests.Count -ile 0) {
            return
        }
        
        foreach ($request in $Requests) {
            try {
                $siteUrl = $request[$script:SiteURLColumn].TrimEnd("/");
                if ($siteUrl -eq $null) {
                    throw "Site URL cannot be null."
                }
                $siteUrl = $siteUrl.TrimEnd("/")
                # try getting the site
                $properties = @{
                        $script:OperationsHistorColumn = "App catalog enabled successfully.";                  
                        $script:StatusColumn           = ([RequestStatus]::Completed).ToString();
                    }
                $site = Get-PnPTenantSite -Url $siteUrl
                if ($site.Status -eq "Active") {
                    LogWrite -Message "Enable App Catalog: $($siteURL)."
                    #Add-PnPSiteCollectionAppCatalog -Site $SiteUrl
                    $siteContext = ConnectPnpOnlineOAuth -TenantId $script:TenantId -ClientId $script:appIdOperationSupport -Thumbprint $script:appThumbprintOperationSupport -Url $siteURL            
                    $context = Get-PnPContext               
                    $appcatalog = $context.Site.RootWeb.SiteCollectionAppCatalog
                    $context.Load($appcatalog)
                    $context.ExecuteQuery()                    
                    if ($appcatalog.ServerObjectIsNull){
                        $context.Load($context.Web);
                        $context.ExecuteQuery();
                        $context.Web.TenantAppCatalog.SiteCollectionAppCatalogsSites.Add($siteUrl);
                        $context.Web.Update();
                        $context.ExecuteQuery();
                    }
                    else{
                        $properties = @{
                            $script:OperationsHistorColumn = "Site already enabled app catalog";                     
                            $script:StatusColumn           = ([RequestStatus]::Completed).ToString();
                        }
                    }
                }                
                else {
                    $message = "$siteUrl : Site does not exist or is not available (status = $($site.Status))"
                    LogWrite -Level ERROR -Message $message
                    $properties = @{
                        $script:OperationsHistorColumn = "Site does not exist or is not available (status = $($site.Status)";                     
                        $script:StatusColumn           = ([RequestStatus]::Error).ToString();
                    } 
                }          
            }
            catch {
                $properties = @{
                    $script:OperationsHistorColumn = "An error occurred - enable app catalog.$($_.Exception.Message)";                     
                    $script:StatusColumn           = ([RequestStatus]::Error).ToString();
                }
                LogWrite -Level ERROR -Message "Process-AppcatalogRequests - Error enable App catalog $SiteUrl : $($_.Exception)"      
            }    
            finally {
                LogWrite -Level INFO -Message "Update request status."
                PnpUpdateListItem -ListName $script:ListId -ItemId $request.Id -Values $properties -SiteConnection $script:listContext  
                $context = $null
                #DisconnectPnpOnlineOAuth -Context $siteContext
            }        
        }
        LogWrite -Level INFO -Message "Processing App catalog Requests completed."        
    }
    catch {
        LogWrite -Level ERROR -Message "Process-AppcatalogRequests - Unexpected exception: $($_.Exception)"        
        throw $_.Exception
    }   
}

function Process-ExternalSharingRequests {
    param
    (        
        [Parameter(Mandatory = $true)] $Requests,
        [Parameter(Mandatory = $false)] $ExternalSharingEnabled    
    )

    try {
        LogWrite -Level INFO -Message "Processing External Sharing Requests."
        LogWrite -Level INFO -Message "Total External Sharing Requests: $($Requests.Count)."

        if ($Requests.Count -ile 0) {
            return
        }

        $tenantProps = Get-PnPTenant
        $currentShareSettings = $tenantProps.SharingCapability
        LogWrite -Level INFO -Message "Current SharingCapability settings in the tenant: $($tenantProps)."
        
        if ($currentShareSettings -eq 'Disabled') {
            LogWrite -Level ERROR -Message "Sharing is currently disabled on the tenant level!"
            return
        }

        $properties = @{
            $script:OperationsHistorColumn = "External sharing [$currentShareSettings] enabled successfully.";                     
            $script:StatusColumn           = ([RequestStatus]::Completed).ToString();
        }
        if ($ExternalSharingEnabled -eq $false){
            $currentShareSettings = 'Disabled'
            $properties = @{
                $script:OperationsHistorColumn = "External sharing [$currentShareSettings] disabled successfully.";                     
                $script:StatusColumn           = ([RequestStatus]::Completed).ToString();
            }
        }
        

        foreach ($request in $Requests) {
            try {
                $siteUrl = $request[$script:SiteURLColumn].TrimEnd("/");
                $allowedDomains = $request[$script:DomainColumn]

                if ($siteURL -eq $null) {
                    throw "Site URL cannot be null."
                }
                $siteUrl = $siteUrl.Trim()
                # try getting the site                              
                $site = Get-PnPTenantSite -Url $siteUrl
                $site.SharingCapability
                if ($site.Status -eq "Active") {
                    LogWrite -Level INFO -Message "Enable External sharing: $($siteURL)." 
                    #update ExternalSharingEnable to Sites table                    
                    UpdateSPOSiteExternalSharingRecord $script:connectionString $site $ExternalSharingEnabled                                       
                    Set-PnPTenantSite -Url $SiteUrl -SharingCapability $currentShareSettings -SharingAllowedDomainList $allowedDomains -SharingDomainRestrictionMode AllowList                    
                }
                else {
                    $message = "$siteUrl : Site does not exist or is not available (status = $($site.Status))"
                    LogWrite -Level ERROR -Message $message
                    $properties = @{
                        $script:OperationsHistorColumn = "Site does not exist or is not available (status = $($site.Status)";                     
                        $script:StatusColumn           = ([RequestStatus]::Error).ToString();
                    } 
                }        
            }
            catch {
                $properties = @{
                    $script:OperationsHistorColumn = "An error occurred - enable external sharing.$($_.Exception.Message)";                     
                    $script:StatusColumn           = ([RequestStatus]::Error).ToString();
                }
                LogWrite -Level ERROR -Message "Process-ExternalSharingRequests - Error external sharing $SiteUrl : $($_.Exception)"      
            }    
            finally {
                LogWrite -Level INFO -Message "Update request status."
                PnpUpdateListItem -ListName $script:ListId -ItemId $request.Id -Values $properties -SiteConnection $script:listContext
                #update ExternalSharingEnable to Sites table
                #$site = Get-PnPTenantSite -Url $siteUrl                 
                #UpdateSPOSiteExternalSharingRecord $script:connectionString $site $ExternalSharingEnabled
            }        
        }
        LogWrite -Level INFO -Message "Processing external sharing completed."        
    }
    catch {
        LogWrite -Level ERROR -Message "Process-ExternalSharingRequests - Unexpected exception: $($_.Exception)"        
        throw $_.Exception
    }
}


function Process-UpdateStorageQuotaRequests {
    param
    (        
        [Parameter(Mandatory = $true)] $Requests        
    )

    try {
        LogWrite -Level INFO -Message "Processing Update Storage Quota Requests."
        LogWrite -Level INFO -Message "Total Update Storage Quota Requests: $($Requests.Count)."

        if ($Requests.Count -ile 0) {
            return
        }

        foreach ($request in $Requests) {
            try {
                $properties = @{
                    $script:OperationsHistorColumn = "Update Storage Quota successfully.";                     
                    $script:StatusColumn           = ([RequestStatus]::Completed).ToString();
                }
              
                $siteURL = $request[$script:SiteURLColumn]
                $newQuota = $request[$script:StorageQuotaColumn]

                if ($siteURL -eq $null) {
                    throw "Site URL cannot be null."
                }
                $siteUrl = $siteUrl.TrimEnd("/")
                # try getting the site
                $site = Get-PnPTenantSite -Url $siteUrl                                
                if ($site.Status -eq "Active") {
                    LogWrite -Level INFO -Message "Update Storage Quota for: $($siteURL). New quota: $newQuota"
                    $quotaValue = $newQuota * 1024
                    $currentStorageWarningLevel = $site.StorageWarningLevel/1024
                    $currentStorageMaximumLevel = $site.StorageMaximumLevel/1024
                    # StorageMaximumLevel always greater than StorageWarningLevel
                    # If newQuota is greater than currentStoreWarningLevel needs to update StorageWarningLevel along with StorageMaximumLevel
                    if ($quotaValue -gt $currentStorageWarningLevel){
                        $currentStorageWarningLevel = $quotaValue * 98/100
                    }
                    try{
                        Set-PnPTenantSite -Url $SiteUrl -StorageMaximumLevel $quotaValue -StorageWarningLevel $currentStorageWarningLevel              
                    }
                    catch{
                        $properties = @{
                                        $script:OperationsHistorColumn = "An error occurred - Update Storage Quota.$($_.Exception.Message)";                     
                                        $script:StatusColumn           = ([RequestStatus]::Error).ToString();
                                }
                        throw $_
                    }
                }
                else {
                    $message = "$siteUrl : Site does not exist or is not available (status = $($site.Status))"
                    LogWrite -Level ERROR -Message $message
                    $properties = @{
                        $script:OperationsHistorColumn = "Site does not exist or is not available (status = $($site.Status)";                     
                        $script:StatusColumn           = ([RequestStatus]::Error).ToString();
                    } 
                }  
                        
            }
            catch {
                $properties = @{
                    $script:OperationsHistorColumn = "An error occurred - Update Storage Quota.$($_.Exception.Message)";                     
                    $script:StatusColumn           = ([RequestStatus]::Error).ToString();
                }
                LogWrite -Level ERROR -Message "Process-UpdateStorageQuotaRequests - Error Update Storage Quota $SiteUrl : $($_.Exception)"      
            }    
            finally {
                LogWrite -Level INFO -Message "Update request status."
                PnpUpdateListItem -ListName $script:ListId -ItemId $request.Id -Values $properties -SiteConnection $script:listContext
            }        
        }
        LogWrite -Level INFO -Message "Processing Update Storage Quota completed."
        
    }
    catch {
        LogWrite -Level ERROR -Message "Process-UpdateStorageQuotaRequests - Unexpected exception: $($_.Exception)"        
        throw $_.Exception
    }
}

function Process-EnableCustomizationRequests {
    param
    (        
        [Parameter(Mandatory = $true)] $Requests        
    )

    try {
        LogWrite -Level INFO -Message "Processing Enable Site Customization Requests."
        LogWrite -Level INFO -Message "Total Enable Site Customization Requests: $($Requests.Count)."

        if ($Requests.Count -ile 0) {
            return
        }

        foreach ($request in $Requests) {
            try {
                $properties = @{
                    $script:OperationsHistorColumn = "Enable Site Customization successfully.";                     
                    $script:StatusColumn           = ([RequestStatus]::Completed).ToString();
                }
              
                $siteURL = $request[$script:SiteURLColumn]
                $newQuota = $request[$script:StorageQuotaColumn]

                if ($siteURL -eq $null) {
                    throw "Site URL cannot be null."
                }
                $siteURL = $siteURL.TrimEnd("/")
                # try getting the site                
                $site = Get-PnPTenantSite -Url $siteUrl
                if ($site.Status -eq "Active") {
                    LogWrite -Level INFO -Message "Enable Site Customization for: $($siteURL)."                    
                    Set-PnPTenantSite -Url $SiteUrl -NoScriptSite:$false                    
                }
                else {
                    $message = "$siteUrl : Site does not exist or is not available (status = $($site.Status))"
                    LogWrite -Level ERROR -Message $message
                    $properties = @{
                        $script:OperationsHistorColumn = "Site does not exist or is not available (status = $($site.Status)";                     
                        $script:StatusColumn           = ([RequestStatus]::Error).ToString();
                    } 
                } 
                          
            }
            catch {
                $properties = @{
                    $script:OperationsHistorColumn = "An error occurred - Enable Site Customization.$($_.Exception.Message)";                     
                    $script:StatusColumn           = ([RequestStatus]::Error).ToString();
                }
                LogWrite -Level ERROR -Message "Process-EnableCustomizationRequests - Error Enable Site Customization $SiteUrl : $($_.Exception)"      
            }    
            finally {
                LogWrite -Level INFO -Message "Update request status."
                PnpUpdateListItem -ListName $script:ListId -ItemId $request.Id -Values $properties -SiteConnection $script:listContext
            }        
        }
        LogWrite -Level INFO -Message "Processing Enable Site Customization completed."        
    }
    catch {
        LogWrite -Level ERROR -Message "Process-EnableCustomizationRequests - Unexpected exception: $($_.Exception)"        
        throw $_.Exception
    }
}

function Process-RegisterHubSite {
    param
    (        
        [Parameter(Mandatory = $true)] $Requests        
    )

    try {
        LogWrite -Level INFO -Message "Processing Register Site as Hub Site Requests."
        LogWrite -Level INFO -Message "Total Register Site as Hub Site Requests: $($Requests.Count)."

        if ($Requests.Count -ile 0) {
            return
        }

        foreach ($request in $Requests) {
            try {
                $properties = @{
                    $script:OperationsHistorColumn = "Register a Site as Hub Site successfully.";                     
                    $script:StatusColumn           = ([RequestStatus]::Completed).ToString();
                }
              
                $siteURL = $request[$script:SiteURLColumn]
                $hubName = $request[$script:HubNameColumn]
                $owners = $request[$script:OwnerColumn]
                $owners = $owners.Email -join ","

                if ($siteURL -eq $null) {
                    throw "Site URL cannot be null."
                }
                $siteURL = $siteURL.TrimEnd("/")
                # try getting the site                
                $site = Get-PnPTenantSite -Url $siteUrl
                if ($site.Status -eq "Active") {
                    LogWrite -Level INFO -Message "Enable a Site as Hub Site for: $($siteURL)."                    
                    Register-PnPHubSite -Site $SiteUrl
                    Set-PnPHubSite -Identity $SiteUrl -Title $hubName
                    if ($owners){
                       Grant-PnPHubSiteRights -Identity $SiteUrl -Principals $owners -Rights Join
                    }                  
                }
                else {
                    $message = "$siteUrl : Site does not exist or is not available (status = $($site.Status))"
                    LogWrite -Level ERROR -Message $message
                    $properties = @{
                        $script:OperationsHistorColumn = "Site does not exist or is not available (status = $($site.Status)";                     
                        $script:StatusColumn           = ([RequestStatus]::Error).ToString();
                    } 
                } 
                          
            }
            catch {
                $properties = @{
                    $script:OperationsHistorColumn = "An error occurred - Register a Site as Hub Site.$($_.Exception.Message)";                     
                    $script:StatusColumn           = ([RequestStatus]::Error).ToString();
                }
                LogWrite -Level ERROR -Message "Process-RegisterHubSite - Error Register a Site as Hub Site $SiteUrl : $($_.Exception)"      
            }    
            finally {
                LogWrite -Level INFO -Message "Update request status."
                PnpUpdateListItem -ListName $script:ListId -ItemId $request.Id -Values $properties -SiteConnection $script:listContext
            }        
        }
        LogWrite -Level INFO -Message "Processing Register a Site as Hub Site completed."        
    }
    catch {
        LogWrite -Level ERROR -Message "Process-RegisterHubSite - Unexpected exception: $($_.Exception)"        
        throw $_.Exception
    }
}

function Process-HideShowFromOutlook {
    param
    (        
        [Parameter(Mandatory = $true)] $Requests,
        [Parameter(Mandatory = $true)] $HideFromOutlookClients         
    )

    try {
        LogWrite -Message "Processing Hide/Show M365 Group from Outlook Requests."
        LogWrite -Message "Total Hide/Show M365 Group from Outlook Requests: $($Requests.Count)."

        if ($Requests.Count -ile 0) {
            return
        }      

        foreach ($request in $Requests) {
            try {

                $properties = @{
                    $script:OperationsHistorColumn = "Hide/Show M365 Group from Outlook successfully.";                     
                    $script:StatusColumn           = ([RequestStatus]::Completed).ToString();
                }

                $groupId = $request[$script:TeamIdColumn].Trim()
                if ($groupId -eq [system.guid]::empty){ 
                    throw "GroupId cannot be null." 
                }               
                $grpSettings = Get-NIHO365Group -AuthToken $script:authToken -Id $groupId -Select hideFromAddressLists

                Update-NIHGroupSettings -AuthToken $script:authToken -Id $groupId -HideFromOutlookClients $HideFromOutlookClients -HideFromAddressLists $grpSettings.hideFromAddressLists 
            }
            catch {
                $properties = @{
                    $script:OperationsHistorColumn = "An error occurred - Hide/Show M365 Group from Outlook Requests.$($_.Exception.Message)";                     
                    $script:StatusColumn           = ([RequestStatus]::Error).ToString();
                }
                LogWrite -Level ERROR -Message "Process-HideShowFromOutlook - Error Hide/Show M365 Group from Outlook Requests $SiteUrl : $($_.Exception)"      
            }    
            finally {
                LogWrite -Level INFO -Message "Update request status."
                PnpUpdateListItem -ListName $script:ListId -ItemId $request.Id -Values $properties -SiteConnection $script:listContext  
            }        
        }
        LogWrite -Level INFO -Message "Processing Hide M365 Group from Outlook Requests completed."        
    }
    catch {
        LogWrite -Level ERROR -Message "Process-HideFromOutlook - Unexpected exception: $($_.Exception)"        
        throw $_.Exception
    }
}

function Process-HideShowFromGAL {
    param
    (        
        [Parameter(Mandatory = $true)] $Requests,
        [Parameter(Mandatory = $true)] $HideFromAddressLists        
    )

    try {
        LogWrite -Message "Processing Hide/Show M365 Group from GAL Requests."
        LogWrite -Message "Total Hide/Show M365 Group from GAL Requests: $($Requests.Count)."

        if ($Requests.Count -ile 0) {
            return
        }      

        foreach ($request in $Requests) {
            try {

                $properties = @{
                    $script:OperationsHistorColumn = "Hide/Show M365 Group from GAL successfully.";                     
                    $script:StatusColumn           = ([RequestStatus]::Completed).ToString();
                }

                $groupId = $request[$script:TeamIdColumn].Trim()
                if ($groupId -eq [system.guid]::empty){ 
                    throw "GroupId cannot be null." 
                }                               

                Update-NIHGroupSettings -AuthToken $script:authToken -Id $groupId -HideFromAddressLists $HideFromAddressLists
            }

            catch {
                $properties = @{
                    $script:OperationsHistorColumn = "An error occurred - Hide/Show M365 Group from GAL Requests.$($_.Exception.Message)";                     
                    $script:StatusColumn           = ([RequestStatus]::Error).ToString();
                }
                LogWrite -Level ERROR -Message "Process-HideFromGAL - Error Hide/Show M365 Group from GAL Requests $SiteUrl : $($_.Exception)"      
            }    
            finally {
                LogWrite -Level INFO -Message "Update request status."
                PnpUpdateListItem -ListName $script:ListId -ItemId $request.Id -Values $properties -SiteConnection $script:listContext  
            }        
        }
        LogWrite -Level INFO -Message "Processing Hide/Show M365 Group from GAL Requests completed."        
    }
    catch {
        LogWrite -Level ERROR -Message "Process-HideShowFromGAL - Unexpected exception: $($_.Exception)"        
        throw $_.Exception
    }
}

function Get-MetaData{        
    $path = "$dp0\Metadata.psd1"
    $listInfo = Import-PowerShellDataFile -Path $path
    #List info
    $script:Url = $listInfo.Url
    $script:ListId = $listInfo.ListId
    $script:ListName = $listInfo.ListName
    #List column
    $script:IncidentColumn = $listInfo.IncidentColumn
    $script:StatusColumn = $listInfo.StatusColumn
    $script:OperationColumn = $listInfo.OperationColumn
    $script:TeamIdColumn = $listInfo.TeamIdColumn
    $script:NewTeamsDisplayNameColumn = $listInfo.NewTeamsDisplayNameColumn
    $script:SiteURLColumn = $listInfo.SiteURLColumn
    $script:DomainColumn = $listInfo.DomainColumn    
    $script:HubNameColumn = $listInfo.HubNameColumn
    $script:OwnerColumn = $listInfo.OwnerColumn
    $script:StorageQuotaColumn = $listInfo.StorageQuotaColumn
    $script:OperationsHistorColumn = $listInfo.OperationsHistorColumn
    
}

#==========================================
#-------- main script starts here ---------
#==========================================

try {
    #-------- Set Global Variables ---------
    Set-TenantVars
    Set-AzureAppVars
    Set-LogFile -logFileName $logFileName    
    #-------- Set Global Variables Ended ---------    
    
    LogWrite -Message "************** Script execution started. **************"     
    Get-MetaData
    #--- Get all requests from Operation list ---   
    LogWrite -Message "Connecting to SharePoint Online '$($script:Url)'..."
    $script:listContext = ConnectPnpOnlineOAuth -TenantId $script:TenantId -ClientId $script:appIdOperationSupport -Thumbprint $script:appThumbprintOperationSupport -Url $script:Url
    #$accessToken = Get-PnPGraphAccessToken
    LogWrite -Message "SharePoint Online '$($script:Url)' is now connected."
       
    LogWrite -Message "Get new requests from the list '$($script:ListName)'"
        $camlQuery = "<View><Query><Where><Eq><FieldRef Name='Status'/><Value Type='Text'>New</Value></Eq></Where></Query></View>"          
        $requests = @(Get-PnPListItem -List "$($script:ListId)" -Connection $script:listContext -Query $camlQuery)                
        $numberOfRequests = $requests.Count
        LogWrite -Message "Total new requests: $($numberOfRequests)."  
        if ($numberOfRequests -gt 0) {                                              
            $externalSharingRequests = $requests | ? { $_["$($script:OperationColumn)"] -eq 'Enable External Sharing' }
            $disableExternalSharingRequests = $requests | ? { $_["$($script:OperationColumn)"] -eq 'Disable External Sharing' }
            $storageQuotaRequests = $requests | ? { $_["$($script:OperationColumn)"] -eq 'Update Storage Quota' }
            $enableCustomizationRequests = $requests | ? { $_["$($script:OperationColumn)"] -eq 'Enable Site Customization' }            
            $registerHubSiteRequests = $requests | ? { $_["$($script:OperationColumn)"] -eq 'Register as a Hub Site' }
            $teamRenameRequests = $requests | ? { $_[$script:OperationColumn] -eq 'Change Team Display Name' } 
            $appCatRequests = $requests | ? { $_[$script:OperationColumn] -eq 'Enable App Catalog' }
            #$enableGroupRequests = $requests | ? { $_[$script:OperationColumn] -eq 'Connect SPO to M365 group' }
            #hideFromOutlookClients and hideFromAddressLists  
            $hideFromOutlookClientsRequests = $requests | ? { $_[$script:OperationColumn] -eq 'Hide M365 Group from Outlook' } 
            $hideFromAddressListsRequests = $requests | ? { $_[$script:OperationColumn] -eq 'Hide M365 Group from GAL' }
            $showFromOutlookClientsRequests = $requests | ? { $_[$script:OperationColumn] -eq 'Show M365 Group from Outlook' } 
            $showFromAddressListsRequests = $requests | ? { $_[$script:OperationColumn] -eq 'Show M365 Group from GAL' }  
            
            #--- SPO Operation using PnpOnline connection ---
            LogWrite -Message "Connecting to SharePoint Admin Center '$($script:SPOAdminCenterURL)'..."
            $script:TenantContext = ConnectPnpOnlineOAuth -TenantId $script:TenantId -ClientId $script:appIdOperationSupport -Thumbprint $script:appThumbprintOperationSupport -Url $script:SPOAdminCenterURL            
            LogWrite -Message "SharePoint Admin Center '$($script:SPOAdminCenterURL)' is now connected."   
            
            if ($storageQuotaRequests){
                Process-UpdateStorageQuotaRequests $storageQuotaRequests
            }
            if ($enableCustomizationRequests){
                Process-EnableCustomizationRequests $enableCustomizationRequests
            }
            if ($registerHubSiteRequests){
                Process-RegisterHubSite $registerHubSiteRequests
            }
            if ($externalSharingRequests){
                Set-DBVars
                Process-ExternalSharingRequests $externalSharingRequests $true
            }
            if ($disableExternalSharingRequests){
                Set-DBVars
                Process-ExternalSharingRequests $disableExternalSharingRequests $false
            }
            if ($appCatRequests){
                Process-AppCatalogRequests $appCatRequests
            }
            #--It's not working with App-only permission.
            <#if ($enableGroupRequests){
                Process-ConnectO365GroupRequests $enableGroupRequests 
            }#>
            #Invoke-GraphAPIAuthTokenCheck
            $Certificate = Get-Item Cert:\LocalMachine\My\* | Where-Object { $_.Thumbprint -ieq "$($script:appThumbprintOperationSupport)" }                
            $script:authToken = Connect-NIHO365GraphWithCert -TenantName $script:TenantId -AppId $script:appIdOperationSupport -Certificate $Certificate
            #--- M365 Group Operation using Graph API ---
            if ($hideFromOutlookClientsRequests){                
                Process-HideShowFromOutlook $hideFromOutlookClientsRequests $true 
            }
            if ($hideFromAddressListsRequests){                
                Process-HideShowFromGAL $hideFromAddressListsRequests $true 
            }
            if ($showFromOutlookClientsRequests){                
                Process-HideShowFromOutlook $showFromOutlookClientsRequests $false 
            }
            if ($showFromAddressListsRequests){                
                Process-HideShowFromGAL $showFromAddressListsRequests $false 
            }  
                        
            #--- MS Teams Operation using Teams API ---            
            if ($teamRenameRequests){                
                Connect-MicrosoftTeams -TenantId $script:TenantId -ApplicationId  $script:appIdOperationSupport -CertificateThumbprint $script:appThumbprintOperationSupport
                Process-TeamRenameRequests $teamRenameRequests
            }                      
        } 
    

    LogWrite -Message "************** Script execution completed. **************"  
    
}
catch {
    LogWrite -Level ERROR "Error in the script: $($_)"
}
finally{
    LogWrite -Level INFO -Message "Disconnect Microsoft Teams."
    DisconnectMicrosoftTeams   
}
