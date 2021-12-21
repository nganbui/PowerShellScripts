Function UpdatePowerBIworkspaceRecord {
    param($SqlConnection, $PowerBIworkspaceObj)
    
    try {
        # initialize stored procedure
        $SqlCmd = New-Object System.Data.SqlClient.SqlCommand
        $SqlCmd.CommandType = [System.Data.CommandType]'StoredProcedure'
        $SqlCmd.CommandText = "SetPowerBIworkspaceInfo"
        $SqlCmd.Connection = $SqlConnection
        $SqlCmd.CommandTimeout = 0
        $workspaceID = $PowerBIworkspaceObj.ID

        $SqlCmd.Parameters.AddWithValue("Id", [string]$PowerBIworkspaceObj.ID)
        $SqlCmd.Parameters.AddWithValue("Name", [string]$PowerBIworkspaceObj.Name)
        $SqlCmd.Parameters.AddWithValue("Description", [string]$PowerBIworkspaceObj.Description)
        $SqlCmd.Parameters.AddWithValue("Type", [string]$PowerBIworkspaceObj.Type)
        $SqlCmd.Parameters.AddWithValue("State",[string]$PowerBIworkspaceObj.State)
        $SqlCmd.Parameters.AddWithValue("IsReadOnly", [string]$PowerBIworkspaceObj.IsReadOnly)
        $SqlCmd.Parameters.AddWithValue("IsOnDedicatedCapacity", [string]$PowerBIworkspaceObj.IsOnDedicatedCapacity)
        $SqlCmd.Parameters.AddWithValue("CapacityId", [string]$PowerBIworkspaceObj.CapacityId)
        $SqlCmd.Parameters.AddWithValue("IsOrphaned", [string]$PowerBIworkspaceObj.IsOrphaned)
        $SqlCmd.Parameters.AddWithValue("Users", [string]$PowerBIworkspaceObj.Admins)
        $SqlCmd.Parameters.AddWithValue("Reports", [string]$PowerBIworkspaceObj.Reports)
        $SqlCmd.Parameters.AddWithValue("Dashboards", [string]$PowerBIworkspaceObj.Dashboards)
        $SqlCmd.Parameters.AddWithValue("Dataflows", [string]$PowerBIworkspaceObj.Dataflows)
        $SqlCmd.Parameters.AddWithValue("Workbooks", [string]$PowerBIworkspaceObj.Workbooks)
        $SqlCmd.Parameters.AddWithValue("Admins", [string]$PowerBIworkspaceObj.Admins)
        $SqlCmd.Parameters.AddWithValue("Contributors", [string]$PowerBIworkspaceObj.Contributors)
        $SqlCmd.Parameters.AddWithValue("Viewers", [string]$PowerBIworkspaceObj.Viewers)

        $SqlCmd.Parameters.AddWithValue("ICName", $null)        
        if ($PowerBIworkspaceObj.ICName -ne '' -and $PowerBIworkspaceObj.ICName -ne $null) {
            $SqlCmd.Parameters["ICName"].Value = [string]$PowerBIworkspaceObj.ICName
        } 

        $SqlCmd.Parameters.AddWithValue("Created", $null)        
        if ($PowerBIworkspaceObj.Created -ne '' -and $PowerBIworkspaceObj.Created -ne $null) {
            $SqlCmd.Parameters["Created"].Value = [string]$PowerBIworkspaceObj.Created
        } 

        $SqlCmd.ExecuteNonQuery()
        
    }
    catch {
        LogWrite -Level ERROR -Message "Adding the PowerBIworkspace info to DB: $workspaceID. Error Info: $_"
    }
    finally{
        $SqlCmd.Dispose()
    }
}

Function UpdateSQLPowerBIworkspace {
    <#
      .Synopsis
        Update a PowerBIworkspace to DB      
    #>
    param($connectionString, $workspaceData)
    if ($workspaceData) {
        try {
            #Initialize SQL Connections
            $SqlConnection = New-Object System.Data.SqlClient.SqlConnection
            $SqlConnection.ConnectionString = $connectionString   
            $SqlConnection.Open()

            UpdatePowerBIworkspaceRecord $SqlConnection $workspaceData
        }
        catch {
            LogWrite -Level ERROR -Message "Error connecting to Database: $($_)"
        }
        
        finally {
            $SqlConnection.Dispose()          
            $SqlConnection.Close()
        }
    }           
}

Function UpdateSQLPowerBIworkspaces {
    param($connectionString, $workspacesData)
   
    if ($workspacesData) {
        #Initialize SQL Connections
        try {
            $SqlConnection = New-Object System.Data.SqlClient.SqlConnection
            $SqlConnection.ConnectionString = $connectionString   
            $SqlConnection.Open()
            $i = 0
            $count = $workspacesData.Count
        
            foreach ($wp in $workspacesData) {
                if ($wp) {
                    UpdatePowerBIworkspaceRecord $SqlConnection $wp
                    $i++
                
                    LogWrite -Message "($($i)/$($count)): $($wp.Name)"
                }
            }
        }
        catch {
            LogWrite -Level ERROR -Message "Error connecting to Database: $($_)"
        }
        
        finally {
            $SqlConnection.Dispose()          
            $SqlConnection.Close()
        }
    }         
}

#region Permanently delete PowerBI Workspaces
Function DeleteInvalidPowerBIworkspaces {
    param($connectionString,$SyncDate)   
    #Initialize SQL Connections
    $SqlConnection = New-Object System.Data.SqlClient.SqlConnection
    $SqlConnection.ConnectionString = $connectionString   
    $SqlConnection.Open()    
    try {
        # initialize stored procedure
        $SqlCmd = New-Object System.Data.SqlClient.SqlCommand
        $SqlCmd.CommandType = [System.Data.CommandType]'StoredProcedure'
        $SqlCmd.CommandText = "DeleteInvalidPowerBIWorkspaces"
        $SqlCmd.Connection = $SqlConnection
        $SqlCmd.Parameters.AddWithValue("SyncDate", $SyncDate)
        $res = $SqlCmd.ExecuteNonQuery()
    }
    catch {
        LogWrite -Level ERROR "Deleting invalid PowerBI Workspaces from DB issue: $($_)"
    }
    finally{
        #Close Connection        
        $SqlCmd.Dispose()                     
        $SqlConnection.Dispose()
        $SqlConnection.Close()  
    }
}
#endregion