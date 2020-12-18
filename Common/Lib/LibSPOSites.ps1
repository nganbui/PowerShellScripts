Function GetAllSPOSites {
    try {        
        LogWrite -Message "Connecting to SharePoint Online..."
        #ConnectPnpOnlineOAuth -TenantId $script:TenantId -ClientId $script:appIdAdminPortal -Thumbprint $script:appThumbprintAdminPortal -Url $script:SPOAdminCenterURL
        $script:TenantContext = ConnectPnpOnlineOAuth -TenantId $script:TenantId -ClientId $script:appIdOperationSupport -Thumbprint $script:appThumbprintOperationSupport -Url $script:SPOAdminCenterURL
        LogWrite -Message "SharePoint Online Administration Center is now connected."        
    }
    catch {    
        LogWrite -Level ERROR -Message "Unable to connect Sharepoint Online Session"
        LogWrite -Level ERROR -Message "$($_.Exception)"
        exit
    }
    try {
        $retrivalStartTime = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

        <##region SPO Sites
        #Retrieve Active SPO Sites
        LogWrite -Message "Retrieving Active SPO Sites from O365..."
        #Invoke-GraphAPIAuthTokenCheck
        #$script:siteUsage = (Get-NIHActivityReport -AuthToken $script:authToken -Report 'getSharePointSiteUsageDetail' -Period D7) | Select *
        $sites = Get-PnPTenantSite -Filter "Url -notlike '-my.sharepoint.com'" | select * | Select-Object *        
        #$groupSites = Get-PnPTenantSite -Template 'GROUP#0' -Detailed | select *        
        $script:sitesData = ParseSPOSites -SitesType Sites -sitesObj $sites -ObjectState Active -ParseObjType O365        
        UpdateSitesProperties -SiteObjects $script:sitesData -SitesType Sites
        #Retrieve Soft Deleted SPO Sites
        LogWrite -Message "Retrieving Soft Deleted SPO Sites from O365..."
        $sites = Get-PnPTenantRecycleBinItem | select *       
        $script:deletedSitesData = ParseSPOSites -sitesObj $sites -SitesType Sites -ObjectState InActive -ParseObjType O365
        #endregion
        #>

        #region Personal Sites
        #Retrieve Active Personal Sites
        LogWrite -Message "Retrieving Active Personal Sites from O365..."        
        $sites = Get-PnPTenantSite -IncludeOneDriveSites -Filter "Url -like '-my.sharepoint.com/personal/'" | select *        
        $script:personalSitesData = ParseSPOSites -sitesObj $sites -SitesType PersonalSites -ObjectState Active -ParseObjType O365
        UpdateSitesProperties -SiteObjects $script:personalSitesData -SitesType PersonalSites
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
        #Initailize objects with Null value
        $TemplateID, $PrimarySCA, $StorageQuota, $StorageUsed, $StorageWarningLevel, $ResourceUsage, $ResourceWarningLevel, $NumberOfSubSites, $IsAuditEnabled = $null


        if (($siteObj.IsAuditEnabled -ne $null) -and ($siteObj.IsAuditEnabled -eq 1)) {
            $isAuditEnabled = 1
        }
        else {
            $isAuditEnabled = 0
        }
        # Get SPO Site Usage        
        $siteUsage = $script:siteUsage | ? { $_.'Site Url' -eq $siteObj.Url } | Select *

        #The object names are different in DB to O365 so we parse them seperately. 
        #And also for values like ResourceUsed,StorageUsed etc, the value from O365 is in MB, where as while updating to DB we convert it to GB ---We dont have to convert them to GB here, we can handle it while reteriving
        switch ($ParseObjType) {
            "O365" {
                $TemplateID = $siteObj.Template      
                $PrimarySCA = $siteObj.Owner  
                $StorageQuota = ($siteObj.StorageMaximumLevel) / 1024
                $StorageUsed = ($siteObj.StorageUsage) / 1024
                $StorageWarningLevel = ($siteObj.StorageWarningLevel) / 1024
                $ResourceUsage = $siteObj.CurrentResourceUsage
                #$ResourceWarningLevel = $siteObj.ResourceQuotaWarningLevel
                $NumberOfSubSites = $siteObj.WebsCount
                $GroupId          = $siteObj.GroupId;                           
                $Description      = $siteObj.Description                
                $FilesCount = 0                                
                $PageViews =  0
                $Pagevisits =  0 
                $FilesViewdOrEdited = 0
                $LastActivityDate = $null 
            }

            "DB" {
                $TemplateID = $siteObj.TemplateID
                $PrimarySCA = $siteObj.PrimarySCA
                $StorageQuota = $siteObj.StorageQuota
                $StorageUsed = $siteObj.StorageUsed
                $StorageWarningLevel = $siteObj.StorageWarningLevel
                $ResourceUsage = $siteObj.ResourceUsed
                $ResourceWarningLevel = $siteObj.ResourceWarningLevel
                $NumberOfSubSites = $siteObj.NumberOfSubSites
                $FilesCount = $siteObj.FilesCount
                $skipStorage = $siteObj.SkipAutoStorage               
                
            }
        }
        $sitesFormattedData += [pscustomobject]@{
            SiteType                            = $SitesType;
            SiteName                            = $siteObj.Title;
            URL                                 = $siteObj.Url;
            TemplateID                          = $TemplateID; #--------DB/O365
            ICName                              = $siteObj.ICName;
            Status                              = $siteObj.Status;
            siteStatus                          = $siteStatus;
            PrimarySCA                          = $PrimarySCA; #--------DB/O365
            SecondarySCA                        = "";          #$siteObj.SecondarySCA;
            NumberOfSubSites                    = $NumberOfSubSites; #--------DB/O365            
            StorageQuota                        = $StorageQuota; #--------DB/O365
            StorageUsed                         = $StorageUsed; #--------DB/O365
            StorageWarningLevel                 = $StorageWarningLevel; #--------DB/O365
            ResourceQuota                       = $siteObj.ResourceQuota;
            ResourceUsage                       = $ResourceUsage; #--------DB/O365
            ResourceQuotaWarningLevel           = $ResourceWarningLevel; #--------DB/O365
            SharingCapability                   = $siteObj.SharingCapability;
            LastContentModifiedDate             = $siteObj.LastContentModifiedDate;
            HubSiteId                           = $siteObj.HubSiteId;
            IsHubSite                           = $siteObj.IsHubSite;
            HubName                             = ""
            LockState                           = $siteObj.LockState;
            DenyAddAndCustomizePages            = $siteObj.DenyAddAndCustomizePages;
            PWAEnabled                          = $siteObj.PWAEnabled;
            SiteDefinedSharingCapability        = $siteObj.SiteDefinedSharingCapability;
            SandboxedCodeActivationCapability   = $siteObj.SandboxedCodeActivationCapability;
            DisableCompanyWideSharingLinks      = $siteObj.DisableCompanyWideSharingLinks;
            DisableAppViews                     = $siteObj.DisableAppViews;
            DisableFlows                        = $siteObj.DisableFlows;
            SharingDomainRestrictionMode        = $siteObj.SharingDomainRestrictionMode;
            SharingAllowedDomainList            = $siteObj.SharingAllowedDomainList
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
            IsAuditEnabled                      = $isAuditEnabled;
            SkipAutoStorage                     = $skipStorage;            
            Description                         = $Description;                        
            GroupId                             = $GroupId;
            FilesCount                          = $FilesCount            
            PageViews                           = $PageViews
            Pagevisits                          = $Pagevisits
            FilesViewdOrEdited                  = $FilesViewdOrEdited
            LastActivityDate                    = $LastActivityDate
            Created                             = "";
            Modified                            = "";
            Operation                           = "";
            OperationStatus                     = ""; 
            AdditionalInfo                      = ""
        }
    }
    return $sitesFormattedData    
}

function UpdateSitesProperties {
    param(
		[parameter(Mandatory = $true)]
		[ValidateNotNullOrEmpty()] 		
		$SiteObjects,
        [Parameter(Mandatory = $true)] [ValidateSet("Sites", "PersonalSites")] $SitesType = "Sites"
	)	        
    if ($SiteObjects -ne $null) {       
        try {
            if ($SitesType -eq 'Sites'){                
                foreach ($sitesObj in $siteObjects) {                
                    $siteUrl =  $sitesObj.Url.Trim()
                    $siteDetailed = Get-PnPTenantSite -Url $siteUrl | Select *                
                    if ($siteDetailed.Status -eq 'Active' -and $siteDetailed.LockState -eq 'Unlock'){
                        $sitesObj.PrimarySCA = $siteDetailed.OwnerEmail  
                        $sitesObj.Description = $siteDetailed.Description
                        $sitesObj.NumberOfSubSites = $siteDetailed.WebsCount
                        $sitesObj.GroupId = $siteDetailed.GroupId
                        $sitesObj.SharingDomainRestrictionMode = $siteDetailed.SharingDomainRestrictionMode;
                        $sitesObj.SharingAllowedDomainList = $siteDetailed.SharingAllowedDomainList
                        $sitesObj.SharingBlockedDomainList = $siteDetailed.SharingBlockedDomainList;  
                        $sitesObj.IsHubSite = $siteDetailed.IsHubSite;  
                        $sitesObj.HubSiteId = $siteDetailed.HubSiteId;    
                        $sitesObj.DenyAddAndCustomizePages = $siteDetailed.DenyAddAndCustomizePages;         
                        $sitesObj.AllowEditing = $siteDetailed.AllowEditing;
                        if ($sitesObj.IsHubSite -eq $true){
                            $sitesObj.HubName = (Get-PnPHubSite -Identity $siteUrl).Title
                        }
                    }                
                }
            }
            elseif ($SitesType -eq 'PersonalSites'){
                foreach ($sitesObj in $siteObjects) {                
                    $siteUrl =  $sitesObj.Url.Trim()                
                    LogWrite -Message "Connecting to SharePoint Online $siteUrl"                
                    $siteContext = ConnectPnpOnlineOAuth -TenantId $script:TenantId -ClientId $script:appIdOperationSupport -Thumbprint $script:appThumbprintOperationSupport -Url $siteUrl                                  
                    $context = Get-PnPContext                
                    $Web = $context.Web 
                    $context.Load($Web)
                    $List = $context.Web.Lists.GetByTitle("Documents")
                    $context.Load($List) 
                    $context.ExecuteQuery()
                    $sitesObj.FilesCount = $List.ItemCount
                    $sitesObj.LastActivityDate = $list.LastItemUserModifiedDate.ToString()                    
                    $sitesObj.NumberOfSubSites =  $Web.Webs.Count                
                    $sitesObj.Description = $Web.Description
                    $sitesObj.Created = $Web.Created.ToString()                    
                    $siteAdmins = (Get-PnPSiteCollectionAdmin | ? {$_.Email -ne '' -and $_.Email.ToLower() -notlike 'spoadm*'}).Email -join ";"
                    $sitesObj.SecondarySCA = $siteAdmins
                    $context.Dispose()
                }
            }
            
        }
        catch {
            LogWrite -Level ERROR -Message "An error occured $($_.Exception)"        
        }              
    }
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
