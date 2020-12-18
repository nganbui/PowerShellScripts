Function GetICsInDB {
    Param(
        [Parameter(Mandatory=$true)]$connectionString
        
    )
    Process
    {
        $SqlConnection = New-Object System.Data.SqlClient.SqlConnection
        $SqlConnection.ConnectionString = $connectionString   

        $SqlCmd = New-Object System.Data.SqlClient.SqlCommand
        $SqlCmd.CommandType = [System.Data.CommandType]'StoredProcedure'
        $SqlCmd.CommandText = "GetICsInfo"
        $SqlCmd.Connection = $SqlConnection        
        
        $SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter 
        $SqlAdapter.SelectCommand = $SqlCmd
        $DataSet = New-Object System.Data.DataSet
        $rowCount =$SqlAdapter.Fill($DataSet)
        $ICData = $dataset.Tables[0] 

        try
        {
            $SqlConnection.Open()
            return $ICData[0]
        }
        catch [Exception]
        {
           LogWrite -Level ERROR -Message "Error connecting to Database. Error info: $($_.Exception.Message)" 
        }
        finally
        {
            $SqlConnection.Close()
            $SqlCmd.Dispose()
            $SqlConnection.Dispose()
        }
    }
}