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
            LogWrite -Level ERROR -Message "Updating SPO sites issue: $($_)"
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

        $SqlCmd.Parameters.AddWithValue("Created", $null)        
        if ($siteObj.Created -ne '' -and $siteObj.Created -ne $null) {
            $SqlCmd.Parameters["Created"].Value = [string]$siteObj.Created
        }                
        $SqlCmd.Parameters.AddWithValue("LastContentModifiedDate", $null)        
        if ($siteObj.LastContentModifiedDate -ne '' -and $siteObj.LastContentModifiedDate -ne $null) {
            $SqlCmd.Parameters["LastContentModifiedDate"].Value = [string]$siteObj.LastContentModifiedDate
        }                

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
        $param = $SqlCmd.Parameters.AddWithValue("HubSiteID", [string]$siteObj.HubSiteId)
        $param = $SqlCmd.Parameters.AddWithValue("IsHubSite", [string]$siteObj.IsHubSite)
        $param = $SqlCmd.Parameters.AddWithValue("HubName", [string]$siteObj.HubName)        
        $param = $SqlCmd.Parameters.AddWithValue("FilesCount", [string]$siteObj.FilesCount)
        <#$param = $SqlCmd.Parameters.AddWithValue("PageViews", [string]$siteObj.PageViews)
        $param = $SqlCmd.Parameters.AddWithValue("Pagevisits", [string]$siteObj.Pagevisits)
        $param = $SqlCmd.Parameters.AddWithValue("FilesViewdOrEdited", [string]$siteObj.FilesViewdOrEdited)
        #>

        #$param = $SqlCmd.Parameters.AddWithValue("CommunicationSiteDesign", [string]$siteObj.CommunicationSiteDesign) 
        #$param = $SqlCmd.Parameters.AddWithValue("PrivacySetting", [string]$siteObj.PrivacySetting)   
        #$param = $SqlCmd.Parameters.AddWithValue("GroupEmailAddress", [string]$siteObj.GroupEmailAddress)        
        
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
            LogWrite -Message "Failed for $($siteobj.URL): $($retMsg)"
        }
        
    }
    catch {
        LogWrite -Level ERROR -Message "Adding the Site info to DB issue: $($_)"
    }    
}

Function UpdatePersonalSitesToDatabase {
    $updateStartTime=Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    #Update Active Personal Sites to Database...
    LogWrite -Message "Updating Active Personal Sites to Database..."
    #UpdateSQLPersonalSites $script:connectionString $script:personalSitesData
    UpdateSQLSPOSites $script:connectionString $script:personalSitesData

    #Update Soft Deleted Personal Sites to Database...
    LogWrite -Message "Updating Soft Deleted Personal Sites to Database..."
    #UpdateSQLPersonalSites $script:connectionString $script:deletedPersonalSitesData
    UpdateSQLSPOSites $script:connectionString $script:deletedPersonalSitesData

    $updateEndTime=Get-Date -Format "yyyy-MM-dd HH:mm:ss"

    LogWrite -Message "Update Personal Sites to Database Start Time: $($updateStartTime)"
    LogWrite -Message "Update Personal Sites to Database End Time: $($updateEndTime)"
}

Function UpdateSPOSiteExternalSharingRecord {
    param($connectionString, $siteObj, $ExternalSharingEnabled)
   
    #Initialize SQL Connections
    try {
        $SqlConnection = New-Object System.Data.SqlClient.SqlConnection
        $SqlConnection.ConnectionString = $connectionString   
        $SqlConnection.Open()
        
        try {
            # initialize stored procedure
            $SqlCmd = New-Object System.Data.SqlClient.SqlCommand
            $SqlCmd.CommandType = [System.Data.CommandType]'StoredProcedure'
            $SqlCmd.CommandText = "UpdateSiteExternalSharing"
            $SqlCmd.Connection = $SqlConnection                  
        
            $SqlCmd.Parameters.AddWithValue("URL", [string]$siteObj.URL)
            $SqlCmd.Parameters.AddWithValue("Title", [string]$siteObj.Title)
            $SqlCmd.Parameters.AddWithValue("TemplateId", [string]$siteObj.TemplateId)
            $SqlCmd.Parameters.AddWithValue("AllowDomainList", [string]$siteObj.SharingAllowedDomainList)
            $SqlCmd.Parameters.AddWithValue("SharingCapability", [string]$siteObj.SharingCapability)
            $SqlCmd.Parameters.AddWithValue("SiteDefinedSharingCapability", [string]$siteObj.SiteDefinedSharingCapability)
            $SqlCmd.Parameters.AddWithValue("ExternalSharingEnabled", $ExternalSharingEnabled)  
        
            $res = $SqlCmd.ExecuteNonQuery()
        }
        catch {
            LogWrite -Level ERROR -Message "Updating [ExternalSharingEnabled] to Sites table issue: $($_)"
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
            $SqlCmd.CommandText = $StoredProcedureName
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
           LogWrite -Level ERROR -Message "Getting sites from DB issue: $($_.Exception.Message)" 
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
            LogWrite -Level ERROR "Updating SCA for Personal Site [$($siteObj.URL)] issue: $($_)"
        } 
    }
    catch {
       LogWrite -Level ERROR "Connecting to DB issue: $($_)"
    }
        
    finally {
        $SqlCmd.Dispose()
        $SqlConnection.Dispose()
        $SqlConnection.Close()
    }        
}

#region Permanently delete user
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

Function GetActiveChangeRequests {  
    Param(
        [Parameter(Mandatory=$true)]$connectionString
        #[Parameter(Mandatory=$true)]$RequestStatus
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