Function GetReports {
    Param(
        [Parameter(Mandatory=$true)]$connectionString,
        [Parameter(Mandatory=$true)]$StoredProcedureName,        
        [Parameter(Mandatory=$false)]$StartDate,
        [Parameter(Mandatory=$false)]$EndDate,
        [switch]$Baseline
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
            if (-not $Baseline){
                $SqlCmd.Parameters.AddWithValue("fromDate", $StartDate)
                $SqlCmd.Parameters.AddWithValue("toDate", $EndDate)
            }
            
            $SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter 
            $SqlAdapter.SelectCommand = $SqlCmd
            $SqlAdapter.SelectCommand.CommandTimeout = 0
            $DataSet = New-Object System.Data.DataSet
            #[void] $SqlAdapter.Fill($DataSet) | Out-Null
            $SqlAdapter.Fill($DataSet) | Out-Null 
            #$rowCount | Out-Null
            #$dt = New-Object System.Data.DataTable
            #$null = $SqlAdapter.fill($dt)
            
            $Results = $DataSet.Tables[0] 
            return $Results
            
        }
        catch [Exception]
        {
           LogWrite -Level ERROR -Message "Connecting to Database: $($_.Exception.Message)" 
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
Function GetBaselineReports {
    Param(
        [Parameter(Mandatory=$true)]$connectionString,
        [Parameter(Mandatory=$true)]$StoredProcedureName
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
            
            $SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter 
            $SqlAdapter.SelectCommand = $SqlCmd
            $SqlAdapter.SelectCommand.CommandTimeout = 0
            $DataSet = New-Object System.Data.DataSet
            #[void] $SqlAdapter.Fill($DataSet) | Out-Null
            $SqlAdapter.Fill($DataSet) | Out-Null
            #$rowCount | Out-Null
            #$dt = New-Object System.Data.DataTable
            #$null = $SqlAdapter.fill($dt)
            
            $Results = $DataSet.Tables[0]             
            return $Results
            
        }
        catch [Exception]
        {
           LogWrite -Level ERROR -Message "Connecting to Database: $($_.Exception.Message)" 
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
Function DelTeamsOnlyUsers{
    param($connectionString)   
    #Initialize SQL Connections
    $SqlConnection = New-Object System.Data.SqlClient.SqlConnection
    $SqlConnection.ConnectionString = $connectionString   
    $SqlConnection.Open()    
    try {
        # initialize stored procedure
        $SqlCmd = New-Object System.Data.SqlClient.SqlCommand
        $SqlCmd.CommandType = [System.Data.CommandType]'StoredProcedure'
        $SqlCmd.CommandText = "DeleteTeamsOnlyUsers"
        $SqlCmd.Connection = $SqlConnection        
        $res = $SqlCmd.ExecuteNonQuery()
    }
    catch {
        LogWrite -Level ERROR "Deleting TeamsOnlyUsers from DB issue: $($_)"
    }
    finally{
        #Close Connection        
        $SqlCmd.Dispose()                     
        $SqlConnection.Dispose()
        $SqlConnection.Close()  
    }
}
