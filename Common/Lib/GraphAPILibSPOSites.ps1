Function GetAllSPOSites-backup {
    try {        
        LogWrite -Message "Connecting to SharePoint Online..."
        #ConnectPnpOnlineOAuth -TenantId $script:TenantId -ClientId $script:appIdAdminPortal -Thumbprint $script:appThumbprintAdminPortal -Url $script:SPOAdminCenterURL
        $SPOAdminConnection = ConnectPnpOnlineOAuth -TenantId $script:TenantId -ClientId $script:appIdOperationSupport -Thumbprint $script:appThumbprintOperationSupport -Url $script:SPOAdminCenterURL
        LogWrite -Message "SharePoint Online Administration Center is now connected."        
    }
    catch {    
        LogWrite -Level ERROR -Message "Unable to connect Sharepoint Online Session"
        LogWrite -Level ERROR -Message "$($_.Exception)"
        exit
    }
    try {
        $retrivalStartTime = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
       
        #region SPO Sites
        #Retrieve Active SPO Sites
        LogWrite -Message "Retrieving Active SPO Sites from M365..."
        #--this cmd only get basic site info        
        $sites = Get-PnPTenantSite -Filter "Url -notlike '-my.sharepoint.com'" | select *        
        $script:sitesData = ParseSPOSites -SitesType Sites -sitesObj $sites -ObjectState Active -ParseObjType O365
        #--get extended site props such as:CreatedDate,OwnerEmail,WebsCount,Hub,SharingCapability      
        UpdateSitesProperties -SiteObjects $script:sitesData -SitesType Sites
        #Retrieve Soft Deleted SPO Sites
        LogWrite -Message "Retrieving Soft Deleted SPO Sites from M365..."
        $sites = Get-PnPTenantRecycleBinItem | select *       
        $script:deletedSitesData = ParseSPOSites -sitesObj $sites -SitesType Sites -ObjectState InActive -ParseObjType O365
        #endregion
                
        #region Personal Sites
        #Retrieve Active Personal Sites
        LogWrite -Message "Retrieving Active Personal Sites from M365..."        
        $sites = Get-PnPTenantSite -IncludeOneDriveSites -Filter "Url -like '-my.sharepoint.com/personal/'" | select *        
        $script:personalSitesData = ParseSPOSites -sitesObj $sites -SitesType PersonalSites -ObjectState Active -ParseObjType O365
        #--For personal will get extended site props Weekly due to perfomance issue
        #UpdateSitesProperties -SiteObjects $script:personalSitesData -SitesType PersonalSites
        #endregion
        
        $retrivalEndTime = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

        LogWrite -Message "Retrieval SPO Sites Start Time: $($retrivalStartTime)"
        LogWrite -Message "Retrieval SPO Sites End Time: $($retrivalEndTime)"
        
    }
    catch {
        LogWrite -Level ERROR -Message "An error occured $($_.Exception)"        
    }       
}

Function GetAllSPOSites {
    try {        
        LogWrite -Message "Connecting to SharePoint Online..."        
        $SPOAdminConnection = ConnectPnpOnlineOAuth -TenantId $script:TenantId -ClientId $script:appIdOperationSupport -Thumbprint $script:appThumbprintOperationSupport -Url $script:SPOAdminCenterURL
        LogWrite -Message "SharePoint Online Administration Center is now connected."        
    }
    catch {    
        LogWrite -Level ERROR -Message "Unable to connect Sharepoint Online Session"
        LogWrite -Level ERROR -Message "$($_.Exception)"
        exit
    }
    try {
        $retrivalStartTime = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        <#
            $_.AbsoluteUrl
            $appCatalog = Get-PnPSiteCollectionAppCatalogs
            $appCatalog.Count
        #>
        LogWrite -Message "Retrieving site collection scoped app catalogs..."
        $script:appCatalogSites = AppCatalogSites -Connection $SPOAdminConnection
        #region SPO Sites
        #Retrieve Active SPO Sites
        LogWrite -Message "Retrieving Active SPO Sites from M365..."
        #--this cmd only get basic site info        
        $sites = Get-PnPTenantSite -Filter "Url -notlike '-my.sharepoint.com'" | select *      
        $script:sitesData = ParseSPOSites -SitesType Sites -sitesObj $sites -ObjectState Active -ParseObjType O365       
        #Retrieve Soft Deleted SPO Sites
        LogWrite -Message "Retrieving Soft Deleted SPO Sites from M365..."
        $sites = Get-PnPTenantRecycleBinItem | select *       
        $script:deletedSitesData = ParseSPOSites -sitesObj $sites -SitesType Sites -ObjectState InActive -ParseObjType O365
        #endregion               
                
        $retrivalEndTime = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

        LogWrite -Message "Retrieval SPO Sites Start Time: $($retrivalStartTime)"
        LogWrite -Message "Retrieval SPO Sites End Time: $($retrivalEndTime)"
        
    }
    catch {
        LogWrite -Level ERROR -Message "An error occured $($_.Exception)"        
    }
        
}

Function GetAllPersonalSites {
    try {        
        LogWrite -Message "Connecting to SharePoint Online..."
        #ConnectPnpOnlineOAuth -TenantId $script:TenantId -ClientId $script:appIdAdminPortal -Thumbprint $script:appThumbprintAdminPortal -Url $script:SPOAdminCenterURL
        $SPOAdminConnection = ConnectPnpOnlineOAuth -TenantId $script:TenantId -ClientId $script:appIdOperationSupport -Thumbprint $script:appThumbprintOperationSupport -Url $script:SPOAdminCenterURL
        LogWrite -Message "SharePoint Online Administration Center is now connected."        
    }
    catch {    
        LogWrite -Level ERROR -Message "Unable to connect Sharepoint Online Session"
        LogWrite -Level ERROR -Message "$($_.Exception)"
        exit
    }
    try {
        $retrivalStartTime = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                
        #region Personal Sites
        #Retrieve Active Personal Sites
        LogWrite -Message "Retrieving Active Personal Sites from M365..."        
        $sites = Get-PnPTenantSite -IncludeOneDriveSites -Filter "Url -like '-my.sharepoint.com/personal/' -and Status -eq 'Active' -and LockState -eq 'Unlock'" | select *        
        $script:personalSitesData = ParseSPOSites -sitesObj $sites -SitesType PersonalSites -ObjectState Active -ParseObjType O365
        #--For personal will get extended site props Weekly due to perfomance issue
        #UpdateSitesProperties -SiteObjects $script:personalSitesData -SitesType PersonalSites
        #endregion
        
        $retrivalEndTime = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

        LogWrite -Message "Retrieval SPO Sites Start Time: $($retrivalStartTime)"
        LogWrite -Message "Retrieval SPO Sites End Time: $($retrivalEndTime)"
        
    }
    catch {
        LogWrite -Level ERROR -Message "An error occured $($_.Exception)"        
    }       
}

Function ParseSPOSites {
    #This function has been updated to handle both Sites and PersonalSites
    param(
        $sitesObj,
        [Parameter(Mandatory = $true)] [ValidateSet("Sites", "PersonalSites")] $SitesType = "Sites",
        [ValidateSet("Active", "InActive")] $ObjectState = "Active",
        [ValidateSet("O365", "DB")] $ParseObjType = "O365"
    )
    
    #Parse/Format all sites from SPO => SitesObject
    #----------------------------------------------
    $sitesFormattedData = @()

    if ($ObjectState -eq "InActive") {
        $siteStatus = "SoftDeleted"
    }
    else {
        $siteStatus = "Active"
    }

    foreach ($siteObj in $sitesObj) {        
        $siteUrl = $siteObj.Url
        #Initailize objects with Null value
        $TemplateID, $PrimarySCA, $StorageQuota, $StorageUsed, $StorageWarningLevel, $ResourceUsage, $ResourceWarningLevel, $NumberOfSubSites, $IsAuditEnabled = $null        

        if (($siteObj.IsAuditEnabled -ne $null) -and ($siteObj.IsAuditEnabled -eq 1)) {
            $isAuditEnabled = 1
        }
        else {
            $isAuditEnabled = 0
        }
        $AppCatalogEnabled = "Disabled"
        $scopeAppCatalog = LookupM365Sku -Skus $script:appCatalogSites -Sku $siteUrl
        if ($siteUrl -eq $scopeAppCatalog){
            $AppCatalogEnabled = "Enabled"
        }
        # Get SPO Site Usage        
        # $siteUsage = $script:siteUsage | ? { $_.'Site Url' -eq $siteObj.Url } | Select *

        #The object names are different in DB to O365 so we parse them seperately. 
        #And also for values like ResourceUsed,StorageUsed etc, the value from O365 is in MB, where as while updating to DB we convert it to GB ---We dont have to convert them to GB here, we can handle it while reteriving
        switch ($ParseObjType) {
            "O365" {
                $TemplateID = $siteObj.Template      
                $PrimarySCA = $siteObj.Owner 
                $SecondarySCA = ""
                $Created = ""
                $StorageQuota = ($siteObj.StorageQuota) / 1024
                $StorageUsed = ($siteObj.StorageUsageCurrent) / 1024
                $StorageWarningLevel = ($siteObj.StorageQuotaWarningLevel) / 1024
                #$ResourceUsage = $siteObj.CurrentResourceUsage
                #$ResourceWarningLevel = $siteObj.ResourceQuotaWarningLevel
                $NumberOfSubSites = $siteObj.WebsCount
                $GroupId          = $siteObj.GroupId;
                $RelatedGroupId   = ""
                $Description      = $siteObj.Description                
                $FilesCount = 0                                
                #$PageViews =  0
                #$Pagevisits =  0 
                #$FilesViewdOrEdited = 0
                #$LastActivityDate = $null 
            }

            "DB" {
                $TemplateID = $siteObj.TemplateID
                $PrimarySCA = $siteObj.PrimarySCA
                $SecondarySCA = $siteObj.SecondarySCA
                $Created = $siteObj.Created
                $StorageQuota = $siteObj.StorageQuota
                $StorageUsed = $siteObj.StorageUsed
                $StorageWarningLevel = $siteObj.StorageWarningLevel
                $ResourceUsage = $siteObj.ResourceUsed
                $ResourceWarningLevel = $siteObj.ResourceWarningLevel
                $NumberOfSubSites = $siteObj.NumberOfSubSites
                $FilesCount = $siteObj.FilesCount
                $skipStorage = $siteObj.SkipAutoStorage 
                $GroupId          = $siteObj.O365GroupID
                $RelatedGroupId   = $siteObj.RelatedGroupId              
                
            }
        }
        $sitesFormattedData += [pscustomobject]@{
            SiteType                            = $SitesType
            SiteName                            = $siteObj.Title
            URL                                 = $siteUrl
            Description                         = $Description
            TemplateID                          = $TemplateID
            GroupId                             = $GroupId
            RelatedGroupId                      = $RelatedGroupId
            ICName                              = $siteObj.ICName
            Status                              = $siteObj.Status
            siteStatus                          = $siteStatus
            LockState                           = $siteObj.LockState
            PrimarySCA                          = $PrimarySCA
            SecondarySCA                        = $SecondarySCA
            SharingCapability                   = $siteObj.SharingCapability
            SiteDefinedSharingCapability        = $siteObj.SiteDefinedSharingCapability
            SharingDomainRestrictionMode        = $siteObj.SharingDomainRestrictionMode
            SharingAllowedDomainList            = $siteObj.SharingAllowedDomainList
            SharingBlockedDomainList            = $siteObj.SharingBlockedDomainList
            LastContentModifiedDate             = $siteObj.LastContentModifiedDate
            Created                             = $Created            
            DeletionTime                        = $siteObj.DeletionTime
            DaysRemaining                       = $siteObj.DaysRemaining
            Modified                            = "";            
            IsHubSite                           = $siteObj.IsHubSite
            HubName                             = ""
            HubSiteId                           = $siteObj.HubSiteId
            IsAuditEnabled                      = $isAuditEnabled
            AllowEditing                        = $siteObj.AllowEditing    
            DenyAddAndCustomizePages            = $siteObj.DenyAddAndCustomizePages
            AppCatalogEnabled                   = $AppCatalogEnabled
            NumberOfSubSites                    = $NumberOfSubSites
            StorageQuota                        = $StorageQuota
            StorageUsed                         = $StorageUsed
            StorageWarningLevel                 = $StorageWarningLevel
            FilesCount                          = $FilesCount            
            ResourceQuota                       = $siteObj.ResourceQuota
            ResourceUsage                       = $ResourceUsage
            ResourceQuotaWarningLevel           = $ResourceWarningLevel
            PWAEnabled                          = $siteObj.PWAEnabled            
            SandboxedCodeActivationCapability   = $siteObj.SandboxedCodeActivationCapability
            DisableCompanyWideSharingLinks      = $siteObj.DisableCompanyWideSharingLinks
            DisableAppViews                     = $siteObj.DisableAppViews
            DisableFlows                        = $siteObj.DisableFlows            
            ConditionalAccessPolicy             = $siteObj.ConditionalAccessPolicy;
            AllowDownloadingNonWebViewableFiles = $siteObj.AllowDownloadingNonWebViewableFiles
            LimitedAccessFileType               = $siteObj.LimitedAccessFileType            
            CommentsOnSitePagesDisabled         = $siteObj.CommentsOnSitePagesDisabled
            DefaultSharingLinkType              = $siteObj.DefaultSharingLinkType
            DefaultLinkPermission               = $siteObj.DefaultLinkPermission
            SkipAutoStorage                     = $skipStorage            
            #PageViews                           = $PageViews
            #Pagevisits                          = $Pagevisits
            #FilesViewdOrEdited                  = $FilesViewdOrEdited                       
            Operation                           = "";
            OperationStatus                     = ""; 
            AdditionalInfo                      = ""
        }
    }
    return $sitesFormattedData    
}

Function UpdateSitesProperties {
    param(
		[parameter(Mandatory = $true)]
		[ValidateNotNullOrEmpty()] 		
		$SiteObjects,
        [Parameter(Mandatory = $true)] [ValidateSet("Sites", "PersonalSites")] $SitesType = "Sites"
	)	        
    if ($SiteObjects -ne $null) {       
        try {
            $i = 1
            $totalSites = $SiteObjects.Count 
            if ($SitesType -eq 'Sites'){                
                foreach ($sitesObj in $siteObjects) { 
                    try{               
                        $siteUrl =  $sitesObj.Url.Trim()
                        LogWrite -Message  "($i/$totalSites): Processing the site [$siteUrl]..."                        
                        $siteDetailed = Get-PnPTenantSite -Url $siteUrl | Select * 
                        #Write-host $siteDetailed.relatedgroupid
                        if ($siteDetailed.Status -eq 'Active' -and $siteDetailed.LockState -eq 'Unlock'){
                            try{
                                $siteConn = ConnectPnpOnlineOAuth -TenantId $script:TenantId -ClientId $script:appIdOperationSupport -Thumbprint $script:appThumbprintOperationSupport -Url $siteUrl
                                $createdDate = Get-PnPWeb -Includes Created
                            }
                            catch{
                                LogWrite -Level ERROR -Message "An error occured getting created date of the site: $siteUrl - $($_.Exception)" 
                            }
                        }
                        
                        if ($siteDetailed.Status -eq 'Active' -and $siteDetailed.LockState -eq 'Unlock'){
                            $sitesObj.PrimarySCA = $siteDetailed.OwnerEmail  
                            $sitesObj.Description = $siteDetailed.Description
                            $sitesObj.NumberOfSubSites = $siteDetailed.WebsCount
                            $sitesObj.GroupId = $siteDetailed.GroupId
                            if ($sitesObj.TemplateID -match 'TEAMCHANNEL'){                                
                                $sitesObj.RelatedGroupId = $siteDetailed.RelatedGroupId
                            }
                            $sitesObj.LastContentModifiedDate = ($siteDetailed.LastContentModifiedDate).toshortdatestring()
                            $sitesObj.SharingDomainRestrictionMode = $siteDetailed.SharingDomainRestrictionMode
                            $sitesObj.SharingAllowedDomainList = $siteDetailed.SharingAllowedDomainList
                            $sitesObj.SharingBlockedDomainList = $siteDetailed.SharingBlockedDomainList                             
                            $sitesObj.Created = $createdDate.Created.toshortdatestring()
                            $sitesObj.IsHubSite = $siteDetailed.IsHubSite
                            $sitesObj.HubSiteId = $siteDetailed.HubSiteId    
                            $sitesObj.DenyAddAndCustomizePages = $siteDetailed.DenyAddAndCustomizePages                            
                            $sitesObj.AllowEditing = $siteDetailed.AllowEditing
                            if ($sitesObj.IsHubSite -eq $true){
                                $sitesObj.HubName = (Get-PnPHubSite -Identity $siteUrl).Title
                            }
                            $sca = (Get-PnPSiteCollectionAdmin | ? {$_.Email -ne ''}).Email -join ";"
                            $sitesObj.SecondarySCA = $sca
                        }                        
                     }
                     catch{
                        LogWrite -Level ERROR -Message "An error occured processing site: $siteUrl - $($_.Exception)" 
                     }

                     $i++               
                }
            }
            elseif ($SitesType -eq 'PersonalSites'){
                foreach ($sitesObj in $siteObjects) {
                    try{                
                        $siteUrl =  $sitesObj.Url.Trim()                
                        LogWrite -Message "($i/$totalSites): Processing the personal site $siteUrl"                
                        $siteContext = ConnectPnpOnlineOAuth -TenantId $script:TenantId -ClientId $script:appIdOperationSupport -Thumbprint $script:appThumbprintOperationSupport -Url $siteUrl                                  
                        $context = Get-PnPContext                
                        $Web = $context.Web 
                        $context.Load($Web)               
                        $List = $context.Web.Lists.GetByTitle("Documents")
                        $context.Load($List) 
                        $context.ExecuteQuery()
                        $sitesObj.FilesCount = $List.ItemCount
                        $sitesObj.LastContentModifiedDate = $list.LastItemUserModifiedDate.toshortdatestring()                   
                        $sitesObj.NumberOfSubSites =  $Web.Webs.Count                
                        $sitesObj.Description = $Web.Description
                        $sitesObj.Created = $Web.Created.toshortdatestring()                   
                        $siteAdmins = (Get-PnPSiteCollectionAdmin | ? {$_.Email -ne '' -and $_.Email.ToLower() -notlike 'spoadm*'}).Email -join ";"
                        $sitesObj.SecondarySCA = $siteAdmins
                        $context.Dispose()
                    }
                    catch{
                        LogWrite -Level ERROR -Message "An error occured processing personal site: $siteUrl - $($_.Exception)" 
                    }                    

                    $i++
                }
            }            
        }
        catch {
            LogWrite -Level ERROR -Message "An error occured $($_.Exception)"        
        }              
    }
}

Function AppCatalogSites{
    [cmdletBinding()]
    param(
        [parameter(Mandatory = $true)]
        $Connection
    )
    [System.Collections.ArrayList]$appCatalogSites = @()
    $appCatalog = Get-PnPSiteCollectionAppCatalogs
    if ($appCatalog -and $appCatalog.Count -gt 0){
        $appCatalog.ForEach({
            $null = $appCatalogSites.Add([Ordered]@{
                        $_.AbsoluteUrl = $_.AbsoluteUrl
                    })
            
        })
    }
    return $appCatalogSites
}

Function CacheSPOSites {
    LogWrite -Level INFO -Message "Generating Cache files for SPO Sites starting..."
    if ($script:sitesData -ne $null) {
        SetDataInCache -CacheType O365 -ObjectType SPOSites -ObjectState Active -CacheData $script:sitesData    
    }
    if ($script:deletedSitesData -ne $null) {
        SetDataInCache -CacheType O365 -ObjectType SPOSites -ObjectState InActive -CacheData $script:deletedSitesData
    }
    LogWrite -Level INFO -Message "Generating Cache files for SPO Sites completed."
}

Function CachePersonalSites {
    LogWrite -Level INFO -Message "Generating Cache files for Personal Sites starting..."    
    if ($script:personalSitesData -ne $null) {
        SetDataInCache -CacheType O365 -ObjectType PersonalSites -ObjectState Active -CacheData $script:personalSitesData
    }
    if ($script:deletedPersonalSitesData -ne $null) {
        SetDataInCache -CacheType O365 -ObjectType PersonalSites -ObjectState InActive -CacheData $script:deletedPersonalSitesData
    }
    LogWrite -Level INFO -Message "Generating Cache files for Personal Sites completed."
}

Function CachePersonalSitesExtended {
    LogWrite -Level INFO -Message "Generating Cache files for Personal Sites Extended starting..."    
    if ($script:personalSitesData -ne $null) {
        SetDataInCache -CacheType O365 -ObjectType PersonalSitesExtended -ObjectState Active -CacheData $script:personalSitesData
    }    
    LogWrite -Level INFO -Message "Generating Cache files for Personal Sites Extended completed."
}

Function SyncPersonalSitesFromDBToCache{
    LogWrite -Level INFO -Message "Syncing all DB Personal Sites to Cache"
    $activePersonalSitesInDB = @()
    $deletedPersonalSitesInDB = @()

    #Get Data from DB
    $activePersonalSitesInDB = GetSitesInDB -ConnectionString $script:ConnectionString -SitesType PersonalSites -StatusType Active
    $activePersonalSitesInDB = $activePersonalSitesInDB | ? { $_.PersonalSiteId -ne $null }
    $deletedPersonalSitesInDB = GetSitesInDB -ConnectionString $script:ConnectionString -SitesType PersonalSites -StatusType InActive
    $deletedPersonalSitesInDB = $deletedPersonalSitesInDB | ? { $_.PersonalSiteId -ne $null }

    #Parse DB Data
    if ($null -ne $activePersonalSitesInDB){
        $activePersonalSitesInDB = ParseSPOSites -sitesObj $activePersonalSitesInDB -SitesType PersonalSites -ParseObjType DB -ObjectState Active
    }
    if ($null -ne $deletedPersonalSitesInDB){
        $deletedPersonalSitesInDB = ParseSPOSites -sitesObj $deletedPersonalSitesInDB -SitesType PersonalSites -ObjectState InActive
    }

    #Cache DB Sites Data
    if($null -ne $activePersonalSitesInDB) {
        SetDataInCache -CacheData $activePersonalSitesInDB -CacheType DB -ObjectType PersonalSites -ObjectState Active
    }
    
    if($null -ne $deletedPersonalSitesInDB) {
        SetDataInCache -CacheData $deletedPersonalSitesInDB -CacheType DB -ObjectType PersonalSites -ObjectState InActive
    }
}

Function SyncSitesFromDBToCache{
    LogWrite -Level INFO -Message "Syncing all DB SPO Sites to Cache"
    $activeSitesInDB = @()
    $deletedSitesInDB = @()

    #Get Data from DB
    $activeSitesInDB = GetSitesInDB -ConnectionString $script:ConnectionString -SitesType Sites -StatusType Active
    $activeSitesInDB = $activeSitesInDB | ? { $_.SiteId -ne $null }
    $deletedSitesInDB = GetSitesInDB -ConnectionString $script:ConnectionString -SitesType Sites -StatusType InActive
    $deletedSitesInDB = $deletedSitesInDB | ? { $_.SiteId -ne $null }

    #Parse DB Data
    if ($null -ne $activeSitesInDB){
        $activeSitesInDB = ParseSPOSites -sitesObj $activeSitesInDB -SitesType Sites -ParseObjType DB -ObjectState Active
    }
    if ($null -ne $deletedSitesInDB){
        $deletedSitesInDB = ParseSPOSites -sitesObj $deletedSitesInDB -SitesType Sites -ObjectState InActive
    }

    #Cache DB Sites Data
    if($null -ne $activeSitesInDB) {
        SetDataInCache -CacheData $activeSitesInDB -CacheType DB -ObjectType SPOSites -ObjectState Active
    }
    
    if($null -ne $deletedSitesInDB) {
        SetDataInCache -CacheData $deletedSitesInDB -CacheType DB -ObjectType SPOSites -ObjectState InActive
    }
}

#region Provisioning
Function ParseSPOSite {
    param($siteObj, 
        $ICName,
        $PrimarySCA,
        $SecondarySCA,
        $ExternalSharingEnabled
    )
    if ($siteObj) {
        return [PSCustomObject][ordered]@{            
            SiteType                            = "Sites";
            ICName                              = $ICName
            PrimarySCA                          = $PrimarySCA #$siteObj.OwnerEmail #$PrimarySCA
            SecondarySCA                        = $SecondarySCA #$siteObj.SecondarySCA #$SecondarySCA
            GroupId                             = $siteObj.GroupId;
            SiteName                            = $siteObj.Title;
            URL                                 = $siteObj.Url;
            TemplateID                          = $siteObj.Template            
            Status                              = $siteObj.Status;
            siteStatus                          = $siteObj.Status;           
            NumberOfSubSites                    = $siteObj.WebsCount
            StorageQuota                        = ($siteObj.StorageQuota) / 1024
            StorageUsed                         = ($siteObj.StorageUsageCurrent) / 1024
            StorageWarningLevel                 = ($siteObj.StorageQuotaWarningLevel) / 1024
            ResourceQuota                       = $siteObj.ResourceQuota; #Only Available on Get-SPOSite
            ResourceUsage                       = $siteObj.ResourceUsageCurrent  #Only Available on Get-SPOSite
            ResourceQuotaWarningLevel           = $siteObj.ResourceQuotaWarningLevel  #Only Available on Get-SPOSite
            SharingCapability                   = $siteObj.SharingCapability;
            LastContentModifiedDate             = $siteObj.LastContentModifiedDate;
            HubSiteId                           = $siteObj.HubSiteId;
            IsHubSite                           = $siteObj.IsHubSite;
            LockState                           = $siteObj.LockState;
            DenyAddAndCustomizePages            = $siteObj.DenyAddAndCustomizePages;
            PWAEnabled                          = $siteObj.PWAEnabled;
            SiteDefinedSharingCapability        = $siteObj.SiteDefinedSharingCapability;
            ExternalSharingEnabled              = $ExternalSharingEnabled
            SandboxedCodeActivationCapability   = $siteObj.SandboxedCodeActivationCapability;
            DisableCompanyWideSharingLinks      = $siteObj.DisableCompanyWideSharingLinks;
            DisableAppViews                     = $siteObj.DisableAppViews;
            DisableFlows                        = $siteObj.DisableFlows;
            SharingDomainRestrictionMode        = $siteObj.SharingDomainRestrictionMode;
            SharingAllowedDomainList            = $siteObj.SharingAllowedDomainList;
            SharingBlockedDomainList            = $siteObj.SharingBlockedDomainList;
            ConditionalAccessPolicy             = $siteObj.ConditionalAccessPolicy;
            AllowDownloadingNonWebViewableFiles = $siteObj.AllowDownloadingNonWebViewableFiles;
            LimitedAccessFileType               = $siteObj.LimitedAccessFileType;
            AllowEditing                        = $siteObj.AllowEditing;
            CommentsOnSitePagesDisabled         = $siteObj.CommentsOnSitePagesDisabled;
            DefaultSharingLinkType              = $siteObj.DefaultSharingLinkType;
            DefaultLinkPermission               = $siteObj.DefaultLinkPermission;
            DeletionTime                        = $siteObj.DeletionTime;
            DaysRemaining                       = $siteObj.DaysRemaining;            
            SkipAutoStorage                     = $siteObj.SkipAutoStorage
            Description                         = $siteObj.Description;
            Created                             = Get-Date
            Modified                            = $siteObj.Modified;            
            Operation                           = "";
            OperationStatus                     = ""; 
            AdditionalInfo                      = ""
        }
    }
}
#endregion
