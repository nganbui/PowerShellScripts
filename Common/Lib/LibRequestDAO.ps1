#region Provisioning
Function GetActiveSiteRequests{
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
        $SqlCmd.CommandText = "GetAllActiveRequests"
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
           LogWrite -Level ERROR -Message "Getting active sites requests issue: $($_.Exception.Message)"            
        }
        finally
        {
            $SqlConnection.Close()
            $SqlCmd.Dispose()
            $SqlConnection.Dispose()
        }
    }
}

Function UpdateProvisionRequest() {
    param ($RequestId, $ReqStatusID, $ReqObjectId, 
            [parameter(Mandatory = $false)]$ReqProcessFlag = 0,
            [Parameter(Mandatory=$true)]$connectionString
            )
    try {
        #Initialize SQL Connections
        $SqlConnection = New-Object System.Data.SqlClient.SqlConnection
        $SqlConnection.ConnectionString = $connectionString  

        $modifiedBy = "System"
        $SqlCmd = New-Object System.Data.SqlClient.SqlCommand

        # Initialize Stored procedure and connection
        $SqlCmd.CommandType = [System.Data.CommandType]'StoredProcedure'
        $SqlCmd.CommandText = "UpdateProvisionRequest"
        $SqlCmd.Connection = $SqlConnection

        # Initialize Return values
        $retMessage = new-object System.Data.SqlClient.SqlParameter;
        $retMessage.ParameterName = "@RetMessage";
        $retMessage.Direction = [System.Data.ParameterDirection]'Output';
        $retMessage.DbType = [System.Data.DbType]'String';
        $retMessage.Size = 5000;

        #Update variables and execute the update
        $param = $SqlCmd.Parameters.AddWithValue("@RequestId", [string]$RequestId)
        $param = $SqlCmd.Parameters.AddWithValue("@StatusId", [string]$ReqStatusID)
        $param = $SqlCmd.Parameters.AddWithValue("@ObjectId", [string]$ReqObjectId)
        $param = $SqlCmd.Parameters.AddWithValue("@ProcessFlag", [string]$ReqProcessFlag)        
       
        $SqlCmd.Parameters.Add($retMessage) >> $null;

        $SqlConnection.Open()
        $res = $SqlCmd.ExecuteNonQuery()
        $retMsg = $SqlCmd.Parameters["@retMessage"].Value;
        
        $SqlCmd.Dispose()
        $SqlConnection.Dispose()
        $SqlConnection.Close()
    }
    catch {
        $exception = $_.Exception    
        LogWrite -Level ERROR "[UpdateProvisionRequest]: An error occured for updating provision request: $exception"
        $SqlCmd.Dispose()
        $SqlConnection.Dispose()
        $SqlConnection.Close()
        throw $exception
    }    
}

function GetSiteRequestInfoById{
    Param(
        [Parameter(Mandatory=$true)]$requestId,
        [Parameter(Mandatory=$true)]$connectionString       
    )
    Process
    {
        try{
            $SqlConnection = New-Object System.Data.SqlClient.SqlConnection
            $SqlConnection.ConnectionString = $connectionString    

            $SqlCmd = New-Object System.Data.SqlClient.SqlCommand
            $SqlCmd.CommandType = [System.Data.CommandType]'StoredProcedure'
            $SqlCmd.CommandText = "[GetSiteRequestById]"
            $SqlCmd.Connection = $SqlConnection
            $SqlCmd.Parameters.AddWithValue("@Id", $requestId) | Out-Null
        
            $SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter 
            $SqlAdapter.SelectCommand = $SqlCmd
            $DataSet = New-Object System.Data.DataSet

            $SqlAdapter.Fill($DataSet) | Out-Null
            $requestInfo = $dataset.Tables[0] 
            $SqlConnection.Open()
            $SqlConnection.Close()
            $SqlCmd.Dispose()
            $SqlConnection.Dispose()

            return $requestInfo
        }
        catch [Exception] {           
           #$e = $_.Exception           
           LogWrite -Level ERROR "[GetSiteRequestInfoById]: An error occured for getting site request info:"
           throw $_           
        }        
    }

}

Function ParseRequest{
    param($req,$dt)
    
    if ($null -ne $req.SiteUrl -and [string]::Empty -ne $req.SiteUrl){
        $SiteUrl = $req.SiteUrl.Trim()
        $Alias = ($req.SiteUrl -split "/sites/")[1].Trim()
    }

    $request = @{
        Id = $req.RequestId.Guid
        Type = $req.RequestTypeId.Guid 
        TemplateId = $req.TemplateId
        ObjectId = $req.ObjectId.Guid
        IncidentId = $req.IncidentId 
        Status = $req.RequestStatusId.Guid        
        DisplayName = $req.SiteName.Trim()
        Description = $req.Description.Trim()
        SiteUrl = $SiteUrl
        Alias = $Alias
        PrivacySetting = $req.PrivacySetting
        ICName = $req.ICName.Trim()
        PrimarySCA = $req.PrimarySCA.Trim()
        OwnerId = $dt["SiteOwnerId"]
        OwnerUPN = $dt["SiteOwnerUPN"]
        ExternalSharing = $req.ExternalSharingEnabled
        Requester = $req.CreatedBy
        SiteDesign = $req.CommunicationSiteDesign
        TimeZone = 10
        ResourceQuota = 300
        StorageQuota = 1048576
        LcId = "1033"
        DefaultStorageWarningPercent = 90
        StorageWarningLevel = 943718
        
    }
    return $request
}
#endregion

#region Change Request
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
#endregion
