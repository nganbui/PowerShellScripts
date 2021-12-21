Function UpdateSPOSitesToDatabase {
    $updateStartTime = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    
    #Update Active SPO Sites to Database...
    LogWrite -Message "Updating Active SPO Sites to Database..."    
    UpdateSQLSPOSites $script:connectionString $script:sitesData
    
    #Update Soft Deleted Sites to Database...
    LogWrite -Message "Updating Soft Deleted SPO Sites to Database..."
    UpdateSQLSPOSites $script:connectionString $script:deletedSitesData
    
    #Remove permanently sites from Sites - DB
    LogWrite -Message "Delete Permanently Deleted Sites from Database..."
    $syncDate = Get-Date -format "yyyy-MM-dd"    
    DeleteInvalidSites $script:connectionString $syncDate
    
    $updateEndTime = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    LogWrite -Message "Update SPOSites To Database Start Time: $($updateStartTime)"
    LogWrite -Message "Update SPOSites To Database End Time: $($updateEndTime)"

}

Function UpdateSQLSPOSites {
    param($connectionString, $sitesData)
   
    if ($sitesData -ne $null) {
        #Initialize SQL Connections
        try {
            $SqlConnection = New-Object System.Data.SqlClient.SqlConnection
            $SqlConnection.ConnectionString = $connectionString   
            $SqlConnection.Open()
            $i = 0
            $count = $sitesData.Count
        
            foreach ($site in $sitesData) {
                if ($site -ne $null) {
                    UpdateSPOSiteRecord $SqlConnection $site
                    $i++
                
                    LogWrite -Message "($($i)/$($count)): $($site.Url)"
                }
            }
        }
        catch {
            LogWrite -Level ERROR -Message "Error connecting to Database: $($_)"
        }
        
        finally {            
            $SqlConnection.Close()
        }
    }         
}

Function UpdateSPOSiteRecord {
    param($SqlConnection, $siteObj)
    
    try {
        # initialize stored procedure
        $SqlCmd = New-Object System.Data.SqlClient.SqlCommand
        $SqlCmd.CommandType = [System.Data.CommandType]'StoredProcedure'
        $SqlCmd.CommandText = "SetSiteInfo"
        $SqlCmd.Connection = $SqlConnection

        # supply the name of the stored procedure
        $ret_Status = new-object System.Data.SqlClient.SqlParameter;
        $ret_Status.ParameterName = "@Ret_Status";
        $ret_Status.Direction = [System.Data.ParameterDirection]'Output';
        $ret_Status.DbType = [System.Data.DbType]'String';
        $ret_Status.Size = 100; 

        $ret_Message = new-object System.Data.SqlClient.SqlParameter;
        $ret_Message.ParameterName = "@Ret_Message";
        $ret_Message.Direction = [System.Data.ParameterDirection]'Output';
        $ret_Message.DbType = [System.Data.DbType]'String';
        $ret_Message.Size = 50000;    

        $ret_Operation = new-object System.Data.SqlClient.SqlParameter;
        $ret_Operation.ParameterName = "@ret_Operation";
        $ret_Operation.Direction = [System.Data.ParameterDirection]'Output';
        $ret_Operation.DbType = [System.Data.DbType]'String';
        $ret_Operation.Size = 100;    

        $param = $SqlCmd.Parameters.AddWithValue("SiteType", [string]$siteObj.SiteType)
        $param = $SqlCmd.Parameters.AddWithValue("Title", [string]$siteObj.SiteName)
        $param = $SqlCmd.Parameters.AddWithValue("TemplateId", [string]$siteObj.TemplateID)
        $param = $SqlCmd.Parameters.AddWithValue("PrimarySCA", [string]$siteObj.PrimarySCA)
        $param = $SqlCmd.Parameters.AddWithValue("URL", [string]$siteObj.URL)
        $param = $SqlCmd.Parameters.AddWithValue("SecondarySCA", [string]$siteObj.SecondarySCA)
        $param = $SqlCmd.Parameters.AddWithValue("NumberOfSubsites", [string]$siteObj.NumberOfSubSites)
        $param = $SqlCmd.Parameters.AddWithValue("StorageQuota", [string]$siteObj.StorageQuota)
        $param = $SqlCmd.Parameters.AddWithValue("StorageUsed", [string]$siteObj.StorageUsed)
        $param = $SqlCmd.Parameters.AddWithValue("StorageWarningLevel", [string]$siteObj.StorageWarningLevel)
        $param = $SqlCmd.Parameters.AddWithValue("ResourceQuota", [string]$siteObj.ResourceQuota)
        $param = $SqlCmd.Parameters.AddWithValue("ResourceUsed", [string]$siteObj.ResourceUsage)
        $param = $SqlCmd.Parameters.AddWithValue("ResourceWarningLevel", [string]$siteObj.ResourceQuotaWarningLevel)
        $param = $SqlCmd.Parameters.AddWithValue("ICName", [string]$siteObj.ICName)
        $param = $SqlCmd.Parameters.AddWithValue("Status", [string]$siteObj.Status)
        $param = $SqlCmd.Parameters.AddWithValue("SiteStatus", [string]$siteObj.siteStatus)
        $param = $SqlCmd.Parameters.AddWithValue("SharingCapability", [string]$siteObj.SharingCapability)

        $SqlCmd.Parameters.AddWithValue("ExternalSharingEnabled", $null)
        $externalSharingEnabled = [string]$siteObj.ExternalSharingEnabled        
        if ($externalSharingEnabled -ne '' -and $externalSharingEnabled -ne $null) {
            $SqlCmd.Parameters["ExternalSharingEnabled"].Value = $externalSharingEnabled
        }        
        $SqlCmd.Parameters.AddWithValue("AppCatalogEnabled", $null)
        $appCatalogEnabled = [string]$siteObj.AppCatalogEnabled        
        if ($appCatalogEnabled -ne '' -and $appCatalogEnabled -ne $null) {
            $SqlCmd.Parameters["AppCatalogEnabled"].Value = $appCatalogEnabled
        }
        $param = $SqlCmd.Parameters.AddWithValue("LastContentModifiedDate", [string]$siteObj.LastContentModifiedDate)        
        $param = $SqlCmd.Parameters.AddWithValue("LockState", [string]$siteObj.LockState)
        $param = $SqlCmd.Parameters.AddWithValue("DenyAddAndCustomizePages", [string]$siteObj.DenyAddAndCustomizePages)
        $param = $SqlCmd.Parameters.AddWithValue("PWAEnabled", [string]$siteObj.PWAEnabled)
        $param = $SqlCmd.Parameters.AddWithValue("SiteDefinedSharingCapability", [string]$siteObj.SiteDefinedSharingCapability)
        $param = $SqlCmd.Parameters.AddWithValue("SandboxedCodeActivationCapability", [string]$siteObj.SandboxedCodeActivationCapability)
        $param = $SqlCmd.Parameters.AddWithValue("DisableCompanyWideSharingLinks", [string]$siteObj.DisableCompanyWideSharingLinks)
        $param = $SqlCmd.Parameters.AddWithValue("DisableAppViews", [string]$siteObj.DisableAppViews)
        $param = $SqlCmd.Parameters.AddWithValue("DisableFlows", [string]$siteObj.DisableFlows)
        $param = $SqlCmd.Parameters.AddWithValue("SharingDomainRestrictionMode", [string]$siteObj.SharingDomainRestrictionMode)
        $param = $SqlCmd.Parameters.AddWithValue("SharingAllowedDomainList", [string]$siteObj.SharingAllowedDomainList)
        $param = $SqlCmd.Parameters.AddWithValue("SharingBlockedDomainList", [string]$siteObj.SharingBlockedDomainList)
        $param = $SqlCmd.Parameters.AddWithValue("ConditionalAccessPolicy", [string]$siteObj.ConditionalAccessPolicy)
        $param = $SqlCmd.Parameters.AddWithValue("AllowDownloadingNonWebViewableFiles", [string]$siteObj.AllowDownloadingNonWebViewableFiles)
        $param = $SqlCmd.Parameters.AddWithValue("LimitedAccessFileType", [string]$siteObj.LimitedAccessFileType)
        $param = $SqlCmd.Parameters.AddWithValue("AllowEditing", [string]$siteObj.AllowEditing)
        $param = $SqlCmd.Parameters.AddWithValue("CommentsOnSitePagesDisabled", [string]$siteObj.CommentsOnSitePagesDisabled)
        $param = $SqlCmd.Parameters.AddWithValue("DefaultSharingLinkType", [string]$siteObj.DefaultSharingLinkType)
        $param = $SqlCmd.Parameters.AddWithValue("DefaultLinkPermission", [string]$siteObj.DefaultLinkPermission)        
        $param = $SqlCmd.Parameters.AddWithValue("DaysRemaining", [string]$siteObj.DaysRemaining)
        $param = $SqlCmd.Parameters.AddWithValue("IsAuditEnabled", [string]$siteObj.IsAuditEnabled)
        $param = $SqlCmd.Parameters.AddWithValue("SkipAutoStorage", [string]$siteObj.SkipAutoStorage)     
        $param = $SqlCmd.Parameters.AddWithValue("Description", [string]$siteObj.Description)        
        $param = $SqlCmd.Parameters.AddWithValue("O365GroupID", [string]$siteObj.GroupId)
        $param = $SqlCmd.Parameters.AddWithValue("RelatedGroupID", [string]$siteObj.RelatedGroupId)                      
        $param = $SqlCmd.Parameters.AddWithValue("HubSiteID", [string]$siteObj.HubSiteId)
        $param = $SqlCmd.Parameters.AddWithValue("IsHubSite", [string]$siteObj.IsHubSite)
        $param = $SqlCmd.Parameters.AddWithValue("HubName", [string]$siteObj.HubName)
        $param = $SqlCmd.Parameters.AddWithValue("Created", [string]$siteObj.Created)           
        
        $deletedDate = $null
        if ($siteObj.DeletionTime -ne "" -and $null -ne $siteObj.DeletionTime) {
            $deletedDate = [string]$siteObj.DeletionTime;
        }
        $param = $SqlCmd.Parameters.AddWithValue("DeletionTime", $deletedDate)

        $SqlCmd.Parameters.Add($ret_Status) >> $null;
        $SqlCmd.Parameters.Add($ret_Message) >> $null;
        $SqlCmd.Parameters.Add($ret_Operation) >> $null;
        
        $res = $SqlCmd.ExecuteNonQuery()

        $retStatus = $SqlCmd.Parameters["@Ret_Status"].Value; 
        $retMsg = $SqlCmd.Parameters["@Ret_Message"].Value;
        $retOperation = $SqlCmd.Parameters["@Ret_Operation"].Value;
        
        $siteObj.Operation = $retOperation
        $siteObj.OperationStatus = $retStatus
        $siteObj.AdditionalInfo = $retMsg
        if ($retStatus -eq "Failed") {
            LogWrite -Message "Failed for $($siteobj.URL). ErrorInfo: $($retMsg)"
        }
        
    }
    catch {
        LogWrite -Level ERROR -Message "Adding the Site info to DB issue: $($_)"
        throw $_
    }    
}

Function UpdateSPOSiteExternalSharingRecord {    
    Param(
        [Parameter(Mandatory=$true)]$ConnectionString,
        [Parameter(Mandatory=$true)]$SiteUrl,
        [Parameter(Mandatory=$false)]$SiteObj,
        [Parameter(Mandatory=$false)]$ExternalSharingEnabled = $null,
        [Parameter(Mandatory=$false)]$HubName = $null,
        [Parameter(Mandatory=$false)]$AppCatalogEnabled = $null
    )
   
    #Initialize SQL Connections
    try {
        $SqlConnection = New-Object System.Data.SqlClient.SqlConnection
        $SqlConnection.ConnectionString = $ConnectionString   
        $SqlConnection.Open()
        
        try {
            # initialize stored procedure
            $SqlCmd = New-Object System.Data.SqlClient.SqlCommand
            $SqlCmd.CommandType = [System.Data.CommandType]'StoredProcedure'
            $SqlCmd.CommandText = "UpdateSiteExternalSharing"
            $SqlCmd.Connection = $SqlConnection                  
        
            $SqlCmd.Parameters.AddWithValue("URL", $SiteUrl)
            $SqlCmd.Parameters.AddWithValue("Title", [string]$SiteObj.Title)
            $SqlCmd.Parameters.AddWithValue("TemplateId", [string]$SiteObj.TemplateId)
            $SqlCmd.Parameters.AddWithValue("AllowDomainList", [string]$SiteObj.SharingAllowedDomainList)
            $SqlCmd.Parameters.AddWithValue("SharingCapability", [string]$SiteObj.SharingCapability)
            $SqlCmd.Parameters.AddWithValue("SiteDefinedSharingCapability", [string]$SiteObj.SiteDefinedSharingCapability)
            $SqlCmd.Parameters.AddWithValue("ExternalSharingEnabled", $ExternalSharingEnabled)            
            $SqlCmd.Parameters.AddWithValue("IsHubSite", [string]$SiteObj.IsHubSite)
            $SqlCmd.Parameters.AddWithValue("HubSiteId", [string]$SiteObj.HubSiteId)
            $SqlCmd.Parameters.AddWithValue("HubName", $HubName)

            $SqlCmd.Parameters.AddWithValue("DenyAddAndCustomizePages", $null)        
            if ($SiteObj.DenyAddAndCustomizePages -ne '' -and $SiteObj.DenyAddAndCustomizePages -ne $null) {
                $SqlCmd.Parameters["DenyAddAndCustomizePages"].Value = [string]$SiteObj.DenyAddAndCustomizePages
            }
            $SqlCmd.Parameters.AddWithValue("AppCatalogEnabled",$AppCatalogEnabled)

            $res = $SqlCmd.ExecuteNonQuery()

            LogWrite -Message "Updated props to the site $($siteObj.URL) in sites table"
        }
        catch {
            LogWrite -Level ERROR -Message "Updating props to Sites table issue: $($_)"
        } 
    }
    catch {
        LogWrite -Level ERROR -Message "Connecting to DB issue: $($_)"
    }
        
    finally {            
        $SqlConnection.Close()
        $SqlCmd.Dispose()
        $SqlConnection.Dispose()
    }        
}

Function GetSitesExternalSharing {
    Param(
        [Parameter(Mandatory=$true)]$connectionString,
        [Parameter(Mandatory=$true)]$ExternalSharingEnabled
    )
    Process
    {
        $SqlConnection = New-Object System.Data.SqlClient.SqlConnection
        $SqlConnection.ConnectionString = $connectionString   

        $SqlCmd = New-Object System.Data.SqlClient.SqlCommand
        $SqlCmd.CommandType = [System.Data.CommandType]'StoredProcedure'
        $SqlCmd.CommandText = "GetActiveSitesExternalSharing"
        $SqlCmd.Connection = $SqlConnection
        $SqlCmd.Parameters.AddWithValue("ExternalSharingEnabled", $ExternalSharingEnabled)
        
        $SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter 
        $SqlAdapter.SelectCommand = $SqlCmd
        $DataSet = New-Object System.Data.DataSet
        $rowCount =$SqlAdapter.Fill($DataSet)
        $Urls = $dataset.Tables[0] 

        try
        {
            $SqlConnection.Open()
            return $Urls
        }
        catch [Exception]
        {
           LogWrite -Level ERROR -Message "Getting sites enabled external sharing issue: $($_.Exception.Message)" 
        }
        finally
        {
            $SqlConnection.Close()
            $SqlCmd.Dispose()
            $SqlConnection.Dispose()
        }
    }
}

Function UpdateSPOSiteToDatabase {
    <#
      .Synopsis
        Update a spo site to DB      
    #>
    param($connectionString, $siteData)
    if ($siteData) {
        #Initialize SQL Connections
        $SqlConnection = New-Object System.Data.SqlClient.SqlConnection
        $SqlConnection.ConnectionString = $connectionString   
        $SqlConnection.Open()
        UpdateSPOSiteRecord $SqlConnection $siteData
        #Close Connection
        $SqlConnection.Close()
    }           
}

Function GetSitesInDB {
    param(
            [Parameter(Mandatory=$true)] [ValidateSet("Sites","PersonalSites")] $SitesType="Sites",
            [Parameter(Mandatory=$true)] [ValidateSet("Active","InActive")] $StatusType="Active", 
            [Parameter(Mandatory=$true)] $connectionString
        )
    Process
    {
        try
        { 
            $SqlConnection = New-Object System.Data.SqlClient.SqlConnection
            $SqlConnection.ConnectionString = $connectionString   

            $SqlCmd = New-Object System.Data.SqlClient.SqlCommand
            $SqlCmd.CommandType = [System.Data.CommandType]'StoredProcedure'
            #$SqlCmd.CommandText = $StoredProcedureName
            $SqlCmd.Connection = $SqlConnection            
            $SqlCmd.Parameters.AddWithValue("SiteType", [string]$SitesType)
            #Based on the Status type call the stored procedure
            if($StatusType -eq "Active") {
                $SqlCmd.CommandText = "[GetActiveSites]"
            }
            elseif($StatusType -eq "InActive") {
                $SqlCmd.CommandText = "[GetInActiveSites]"
            }
        
            $SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter 
            $SqlAdapter.SelectCommand = $SqlCmd
            $SqlAdapter.SelectCommand.CommandTimeout = 0
            $DataSet = New-Object System.Data.DataSet            
            
            $SqlAdapter.Fill($DataSet) | Out-Null
            $rowCount = $DataSet.tables[0].rows.count 
            if ($rowCount -gt 0) {
                $Results = $DataSet.tables[0].rows
            }
            return $Results
            
        }
        catch [Exception]
        {
           LogWrite -Level ERROR -Message "Error connecting to Database: $($_.Exception.Message)" 
        }
        finally
        {
            $SqlAdapter.Dispose()
            $SqlCmd.Dispose()                     
            $SqlConnection.Dispose()
            $SqlConnection.Close()   
        }
    }
}

Function UpdateSCA {
    param($connectionString, $siteObj, $sca,
    [Parameter(Mandatory=$true)] [ValidateSet("Sites","PersonalSites")] $SitesType="Sites")
   
    #Initialize SQL Connections
    try {
        $SqlConnection = New-Object System.Data.SqlClient.SqlConnection
        $SqlConnection.ConnectionString = $connectionString   
        $SqlConnection.Open()
        
        try {
            # initialize stored procedure
            $SqlCmd = New-Object System.Data.SqlClient.SqlCommand
            $SqlCmd.CommandType = [System.Data.CommandType]'StoredProcedure'
            $SqlCmd.CommandText = "UpdateSCA"
            $SqlCmd.Connection = $SqlConnection                  

            $SqlCmd.Parameters.AddWithValue("SiteType", $SitesType)        
            $SqlCmd.Parameters.AddWithValue("URL", [string]$siteObj.URL)
            $SqlCmd.Parameters.AddWithValue("SecondarySCA", $sca)
            $res = $SqlCmd.ExecuteNonQuery()
        }
        catch {            
            LogWrite -Level ERROR "Updating SCA for the site [$($siteObj.URL)]: $($_)"
        } 
    }
    catch {
       LogWrite -Level ERROR "Error connecting to Database: $($_)"
    }
        
    finally {
        $SqlCmd.Dispose()
        $SqlConnection.Dispose()
        $SqlConnection.Close()
    }        
}

#region Permanently delete sites
Function DeleteInvalidSites {
    param($connectionString,$SyncDate)   
    #Initialize SQL Connections
    $SqlConnection = New-Object System.Data.SqlClient.SqlConnection
    $SqlConnection.ConnectionString = $connectionString   
    $SqlConnection.Open()    
    try {
        # initialize stored procedure
        $SqlCmd = New-Object System.Data.SqlClient.SqlCommand
        $SqlCmd.CommandType = [System.Data.CommandType]'StoredProcedure'
        $SqlCmd.CommandText = "DeleteInvalidSites"
        $SqlCmd.Connection = $SqlConnection
        $SqlCmd.Parameters.AddWithValue("SyncDate", $SyncDate)
        $res = $SqlCmd.ExecuteNonQuery()
    }
    catch {
        LogWrite -Level ERROR "Deleting invalid sites from DB issue: $($_)"
    }
    finally{
        #Close Connection        
        $SqlCmd.Dispose()                     
        $SqlConnection.Dispose()
        $SqlConnection.Close()  
    }
}
#endregion

#region Change Requests - moved to LibRequestDAO.ps1
<#
Function GetActiveChangeRequests {  
    Param(
        [Parameter(Mandatory=$true)]$connectionString        
    ) 
    Process
    {
        $activeRequests = @()
        $SqlConnection = New-Object System.Data.SqlClient.SqlConnection
        $SqlConnection.ConnectionString = $connectionString   

        $SqlCmd = New-Object System.Data.SqlClient.SqlCommand
        $SqlCmd.CommandType = [System.Data.CommandType]'StoredProcedure'
        $SqlCmd.CommandText = "GetAllActiveChangeRequests"
        $SqlCmd.Connection = $SqlConnection
        $SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter 
        $SqlAdapter.SelectCommand = $SqlCmd
        $DataSet = New-Object System.Data.DataSet
        $rowCount =$SqlAdapter.Fill($DataSet)
        $activeRequests = $dataset.Tables[0] 

        try
        {
            $SqlConnection.Open()
            return $activeRequests
        }
        catch [Exception]
        {
           LogWrite -Level ERROR -Message "Getting active change requests issue: $($_.Exception.Message)" 
        }
        finally
        {
            $SqlConnection.Close()
            $SqlCmd.Dispose()
            $SqlConnection.Dispose()
        }
    }
}

Function UpdateChangeRequest {  
    param($connectionString, $reqObj, $reqStatus)
   
    #Initialize SQL Connections
    try {
        $SqlConnection = New-Object System.Data.SqlClient.SqlConnection
        $SqlConnection.ConnectionString = $connectionString   
        $SqlConnection.Open()
        
        try {
            # initialize stored procedure
            $SqlCmd = New-Object System.Data.SqlClient.SqlCommand
            $SqlCmd.CommandType = [System.Data.CommandType]'StoredProcedure'
            $SqlCmd.CommandText = "UpdateChangeRequest"
            $SqlCmd.Connection = $SqlConnection
            
            $retMessage = new-object System.Data.SqlClient.SqlParameter;
            $retMessage.ParameterName = "@retMessage";
            $retMessage.Direction = [System.Data.ParameterDirection]'Output';
            $retMessage.DbType = [System.Data.DbType]'String';
            $retMessage.Size = 50000;                      
        
            $SqlCmd.Parameters.AddWithValue("ChangeRequestId", [string]$reqObj.ChangeRequestId)
            $SqlCmd.Parameters.AddWithValue("StatusId", [string]$reqStatus)
            
            $SqlCmd.Parameters.Add($retMessage) >> $null
            $res = $SqlCmd.ExecuteNonQuery()
            $retMsg = $SqlCmd.Parameters["@retMessage"].Value
            LogWrite -Message "$($retMsg)"  

        }
        catch {
            LogWrite -Level ERROR -Message "Updating [change request status] to ChangeRequest table issue: $($_)"            
        } 
    }
    catch {
        LogWrite -Level ERROR -Message "Connecting to DB issue: $($_)"
    }
        
    finally {            
        $SqlConnection.Close()
        $SqlCmd.Dispose()
        $SqlConnection.Dispose()
    }  
}
#>

<#  -- moved to GraphAPILibO365GroupsDAO.ps1
Function UpdateGroupTeamPostRename {  
    param($connectionString, $teamObj)
   
    #Initialize SQL Connections
    try {
        $SqlConnection = New-Object System.Data.SqlClient.SqlConnection
        $SqlConnection.ConnectionString = $connectionString   
        $SqlConnection.Open()
        
        try {
            # initialize stored procedure
            $SqlCmd = New-Object System.Data.SqlClient.SqlCommand
            $SqlCmd.CommandType = [System.Data.CommandType]'StoredProcedure'
            $SqlCmd.CommandText = "UpdateGroupTeamInfo"
            $SqlCmd.Connection = $SqlConnection                  
        
            $SqlCmd.Parameters.AddWithValue("GroupId", [string]$teamObj.GroupId)
            $SqlCmd.Parameters.AddWithValue("DisplayName", [string]$teamObj.DisplayName)
            $res = $SqlCmd.ExecuteNonQuery()
                        
            LogWrite -Message "Update group display name [$($teamObj.DisplayName)] into Groups and Teams table."  

        }
        catch {
            LogWrite -Level ERROR -Message "Updating [Group/Team] to Groups and Teams table issue: $($_)"            
        } 
    }
    catch {
        LogWrite -Level ERROR -Message "Connecting to DB issue: $($_)"
    }
        
    finally {            
        $SqlConnection.Close()
        $SqlCmd.Dispose()
        $SqlConnection.Dispose()
    }  
}
#>
#endregion Change Requests

#region Permanently Delete Invalid PersonalSites
Function DeleteInvalidPersonalSites {
    param($connectionString,$SyncDate)   
    #Initialize SQL Connections
    $SqlConnection = New-Object System.Data.SqlClient.SqlConnection
    $SqlConnection.ConnectionString = $connectionString   
    $SqlConnection.Open()    
    try {
        # initialize stored procedure
        $SqlCmd = New-Object System.Data.SqlClient.SqlCommand
        $SqlCmd.CommandType = [System.Data.CommandType]'StoredProcedure'
        $SqlCmd.CommandText = "DeleteInvalidPersonalSites"
        $SqlCmd.Connection = $SqlConnection
        $SqlCmd.Parameters.AddWithValue("SyncDate", $SyncDate)
        $res = $SqlCmd.ExecuteNonQuery()
    }
    catch {
        LogWrite -Level ERROR "Deleting invalid personal sites from DB issue: $($_)"
    }
    finally{
        #Close Connection        
        $SqlCmd.Dispose()                     
        $SqlConnection.Dispose()
        $SqlConnection.Close()  
    }
}
#endregion

Function UpdatePersonalSiteSCA {
    param($connectionString, $siteObj, $sca)
   
    #Initialize SQL Connections
    try {
        $SqlConnection = New-Object System.Data.SqlClient.SqlConnection
        $SqlConnection.ConnectionString = $connectionString   
        $SqlConnection.Open()
        
        try {
            # initialize stored procedure
            $SqlCmd = New-Object System.Data.SqlClient.SqlCommand
            $SqlCmd.CommandType = [System.Data.CommandType]'StoredProcedure'
            $SqlCmd.CommandText = "UpdatePersonalSiteSCA"
            $SqlCmd.Connection = $SqlConnection                  
        
            $SqlCmd.Parameters.AddWithValue("URL", [string]$siteObj.URL)
            $SqlCmd.Parameters.AddWithValue("SecondarySCA", $sca)
            $res = $SqlCmd.ExecuteNonQuery()
        }
        catch {            
            LogWrite -Level ERROR "Updating SCA for Personal Site [$($siteObj.URL)]: $($_)"
        } 
    }
    catch {
       LogWrite -Level ERROR "Error connecting to Database: $($_)"
    }
        
    finally {
        $SqlCmd.Dispose()
        $SqlConnection.Dispose()
        $SqlConnection.Close()
    }        
}

Function UpdatePersonalSitesToDatabase {
    $updateStartTime=Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    #Update Active Personal Sites to Database...
    LogWrite -Message "Updating Active Personal Sites to Database..."   
    UpdateSQLSPOSites $script:connectionString $script:personalSitesData

    #Update Soft Deleted Personal Sites to Database...
    LogWrite -Message "Updating Soft Deleted Personal Sites to Database..."   
    UpdateSQLSPOSites $script:connectionString $script:deletedPersonalSitesData

    #Remove Invalid personal sites from PersonalSites - DB
    LogWrite -Message "Delete Invalid personal sites from Database..."    
    $syncDate = Get-Date -format "yyyy-MM-dd"
    DeleteInvalidPersonalSites $script:connectionString $syncDate

    $updateEndTime=Get-Date -Format "yyyy-MM-dd HH:mm:ss"

    LogWrite -Message "Update Personal Sites to Database Start Time: $($updateStartTime)"
    LogWrite -Message "Update Personal Sites to Database End Time: $($updateEndTime)"
}

Function UpdateSPOSiteExtenedToDatabase {
    $updateStartTime = Get-Date -Format "yyyy-MM-dd HH:mm:ss"    
    
    LogWrite -Message "Updating Active Personal Sites Extended to Database..."    
    UpdateSitesExtenedInfoToDatabase $script:connectionString $script:personalSitesExtendedData
    
    $updateEndTime = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    LogWrite -Message "Update Active Personal Sites Extended To Database Start Time: $($updateStartTime)"
    LogWrite -Message "Update Active Personal Sites Extended To Database End Time: $($updateEndTime)"

}

Function UpdateSitesExtenedInfoToDatabase {
    param($connectionString, $sitesData)   

    if ($sitesData -ne $null) {
        #Initialize SQL Connections
        try {
            $SqlConnection = New-Object System.Data.SqlClient.SqlConnection
            $SqlConnection.ConnectionString = $connectionString   
            $SqlConnection.Open()
            $i = 0
            $count = $sitesData.Count
        
            foreach ($site in $sitesData) {
                if ($site -ne $null) {
                    UpdateSPOSiteExtenedRecord $SqlConnection $site
                    $i++
                
                    LogWrite -Message "($($i)/$($count)): $($site.Url)"
                }
            }
        }
        catch {
            LogWrite -Level ERROR -Message "Connecting to DB issue: $($_)"
        }
        
        finally {            
            $SqlConnection.Close()
        }
    }
}

Function UpdateSPOSiteExtenedRecord {
    param($SqlConnection, $siteObj)
        
    try {
        $SqlConnection = New-Object System.Data.SqlClient.SqlConnection
        $SqlConnection.ConnectionString = $connectionString   
        $SqlConnection.Open()
        
        try {
            $SqlCmd = New-Object System.Data.SqlClient.SqlCommand
        $SqlCmd.CommandType = [System.Data.CommandType]'StoredProcedure'
        $SqlCmd.CommandText = "SetSiteExtendedInfo"
        $SqlCmd.Connection = $SqlConnection

        # supply the name of the stored procedure
        $ret_Status = new-object System.Data.SqlClient.SqlParameter;
        $ret_Status.ParameterName = "@Ret_Status";
        $ret_Status.Direction = [System.Data.ParameterDirection]'Output';
        $ret_Status.DbType = [System.Data.DbType]'String';
        $ret_Status.Size = 100; 

        $ret_Message = new-object System.Data.SqlClient.SqlParameter;
        $ret_Message.ParameterName = "@Ret_Message";
        $ret_Message.Direction = [System.Data.ParameterDirection]'Output';
        $ret_Message.DbType = [System.Data.DbType]'String';
        $ret_Message.Size = 5000;    

        $ret_Operation = new-object System.Data.SqlClient.SqlParameter;
        $ret_Operation.ParameterName = "@ret_Operation";
        $ret_Operation.Direction = [System.Data.ParameterDirection]'Output';
        $ret_Operation.DbType = [System.Data.DbType]'String';
        $ret_Operation.Size = 100;   
        
        $SqlCmd.Parameters.AddWithValue("SiteType", [string]$siteObj.SiteType)
        $SqlCmd.Parameters.AddWithValue("URL", [string]$siteObj.URL)
        $SqlCmd.Parameters.AddWithValue("SecondarySCA", [string]$siteObj.SecondarySCA)
        $SqlCmd.Parameters.AddWithValue("WebsCount", [string]$siteObj.NumberOfSubSites)
        $SqlCmd.Parameters.AddWithValue("FilesCount", [string]$siteObj.FilesCount)
        $SqlCmd.Parameters.AddWithValue("Created", [string]$siteObj.Created)
        $SqlCmd.Parameters.AddWithValue("LastContentModifiedDate", [string]$siteObj.LastContentModifiedDate)        
        #$SqlCmd.Parameters.AddWithValue("ICName", [string]$siteObj.ICName)        
        #$SqlCmd.Parameters.AddWithValue("IsAuditEnabled", [string]$siteObj.IsAuditEnabled)
        #$SqlCmd.Parameters.AddWithValue("IsHubSite", [string]$siteObj.IsHubSite)
        #$SqlCmd.Parameters.AddWithValue("HubSiteID", [string]$siteObj.HubSiteID)
        

        $SqlCmd.Parameters.Add($ret_Status) >> $null;
        $SqlCmd.Parameters.Add($ret_Message) >> $null;
        $SqlCmd.Parameters.Add($ret_Operation) >> $null;
        
        $res = $SqlCmd.ExecuteNonQuery()

        $retStatus = $SqlCmd.Parameters["@Ret_Status"].Value; 
        $retMsg = $SqlCmd.Parameters["@Ret_Message"].Value;
        $retOperation = $SqlCmd.Parameters["@Ret_Operation"].Value

        $siteObj.Operation = $retOperation
        $siteObj.OperationStatus = $retStatus
        $siteObj.AdditionalInfo = $retMsg
        
        }
        catch {
            LogWrite -Level ERROR -Message "Updating [SetSiteExtendedInfo] to Sites table issue: $($_)"
        } 
    }
    catch {
        LogWrite -Level ERROR -Message "Connecting to DB issue: $($_)"
    }
        
    finally {            
        $SqlConnection.Close()
        $SqlCmd.Dispose()
        $SqlConnection.Dispose()
    } 
}