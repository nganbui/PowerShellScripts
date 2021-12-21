Function UpdateO365UsersToDatabase {
    $updateStartTime = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    #Update Active Users to Database...
    LogWrite -Message "Updating Active O365 Users to Database..."
    UpdateSQLO365Users $script:connectionString $script:usersData
    #Update Deleted Users to Database...
    LogWrite -Message "Updating Deleted O365 Users to Database..."
    UpdateSQLO365Users $script:connectionString $script:deletedUsersData
    $updateEndTime = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    LogWrite -Message "Update O365 Users To Database Start Time: $($updateStartTime)"
    LogWrite -Message "Update O365 Users To Database End Time: $($updateEndTime)"

}

Function UpdateSQLO365Users {
    param($connectionString, $usersData)
   
    if ($usersData -ne $null)
    {
        try {
            #Initialize SQL Connections
            $SqlConnection = New-Object System.Data.SqlClient.SqlConnection
            $SqlConnection.ConnectionString = $connectionString   
            $SqlConnection.Open()

            $i=0
            $count=$usersData.Count
        
            foreach($userObj in $usersData)
            {
                if($userObj -ne $null)
                {
                    UpdateO365UserRecord $SqlConnection $userObj
                    $i++
                
                    LogWrite -Message "($($i)/$($count)): $($userObj.UserPrincipalName)"
                }
            }
        }
        catch {
            LogWrite -Level ERROR -Message "Error connecting to Database. Error info: $($_)"
        }
        
        finally{            
            $SqlConnection.Close()
        }        
    }         
}

Function UpdateO365UserRecord {
    param($SqlConnection,$userObj)
    $retStatus=$null
    $retMsg=$null
    $retOperation=$null
    $retUserExistance = $null
    
    try {
        $SqlCmd = New-Object System.Data.SqlClient.SqlCommand

        # indicate that we are working with stored procedure
        $SqlCmd.CommandType=[System.Data.CommandType]'StoredProcedure'

        # supply the name of the stored procedure - WHY the name is "AddNewUser" ?
        $SqlCmd.CommandText = "UpdateUserInfo"
        $SqlCmd.Connection = $SqlConnection
    
        $ret_Status = new-object System.Data.SqlClient.SqlParameter;
        $ret_Status.ParameterName = "@Ret_Status";
        $ret_Status.Direction = [System.Data.ParameterDirection]'Output';
        $ret_Status.DbType = [System.Data.DbType]'String';
        $ret_Status.Size = 100; 

        $ret_Operation = new-object System.Data.SqlClient.SqlParameter;
        $ret_Operation.ParameterName = "@ret_Operation";
        $ret_Operation.Direction = [System.Data.ParameterDirection]'Output';
        $ret_Operation.DbType = [System.Data.DbType]'String';
        $ret_Operation.Size = 100;    
        
        $ret_Message = new-object System.Data.SqlClient.SqlParameter;
        $ret_Message.ParameterName = "@Ret_Message";
        $ret_Message.Direction = [System.Data.ParameterDirection]'Output';
        $ret_Message.DbType = [System.Data.DbType]'String';
        $ret_Message.Size = 5000;  
        
        $ret_userExistance = new-object System.Data.SqlClient.SqlParameter;
        $ret_userExistance.ParameterName = "@ret_userExistance";
        $ret_userExistance.Direction = [System.Data.ParameterDirection]'Output';
        $ret_userExistance.DbType = [System.Data.DbType]'String';
        $ret_userExistance.Size = 10; 

        $param=$SqlCmd.Parameters.AddWithValue("@UserId", [string]$userObj.UserId)
        $param=$SqlCmd.Parameters.AddWithValue("@SigninName", [string]$userObj.UserPrincipalName)
        $param=$SqlCmd.Parameters.AddWithValue("@FirstName",[string]$userObj.GivenName)
        $param=$SqlCmd.Parameters.AddWithValue("@LastName",[string]$userObj.Surname)
        $param=$SqlCmd.Parameters.AddWithValue("@ICName",[string]$userObj.Department)
        $param=$SqlCmd.Parameters.AddWithValue("@Email",[string]$userObj.Mail)
        $param=$SqlCmd.Parameters.AddWithValue("@DisplayName",[string]$userObj.DisplayName)
        $param=$SqlCmd.Parameters.AddWithValue("@Office",[string]$userObj.OfficeLocation)
        $param=$SqlCmd.Parameters.AddWithValue("@IsDisabled",$userObj.IsDisabled)
        $param=$SqlCmd.Parameters.AddWithValue("@IsDeleted",[int]$userObj.IsDeleted)
        $param=$SqlCmd.Parameters.AddWithValue("@IsLicensed",$userObj.IsLicensed)
        $param=$SqlCmd.Parameters.AddWithValue("@Licenses",[string]$userObj.AssignedLicenses)
        $param=$SqlCmd.Parameters.AddWithValue("@AppLicenses",[string]$userObj.AssignedPlans)
        $param=$SqlCmd.Parameters.AddWithValue("@Created",[datetime]$userObj.CreatedDateTime)
        $param=$SqlCmd.Parameters.AddWithValue("@SoftDeletionTimestamp",[string]$userObj.SoftDeletionTimestamp)
        #$param=$SqlCmd.Parameters.AddWithValue("@LastDirSyncTime",[string]$userObj.LastDirSyncTime)
        $param=$SqlCmd.Parameters.AddWithValue("@UserPrincipalName",[string]$userObj.UserPrincipalName)
        $param=$SqlCmd.Parameters.AddWithValue("@LastPasswordChangeTimeStamp",[string]$userObj.lastPasswordChangeDateTime)
        #$param=$SqlCmd.Parameters.AddWithValue("@StsRefreshTokensValidFrom",[string]$userObj.StsRefreshTokensValidFrom)
        $param=$SqlCmd.Parameters.AddWithValue("@StreetAddress",[string]$userObj.StreetAddress)
        $param=$SqlCmd.Parameters.AddWithValue("@City",[string]$userObj.City)
        $param=$SqlCmd.Parameters.AddWithValue("@State",[string]$userObj.State)
        $param=$SqlCmd.Parameters.AddWithValue("@Country",[string]$userObj.Country)
        $param=$SqlCmd.Parameters.AddWithValue("@UserType",[string]$userObj.UserType)
        #$param=$SqlCmd.Parameters.AddWithValue("@UserThemeIdentifierForO365Shell",[string]$userObj.UserThemeIdentifierForO365Shell)
        $param=$SqlCmd.Parameters.AddWithValue("@PhoneNumber",[string]$userObj.BusinessPhones)
        $param=$SqlCmd.Parameters.AddWithValue("@MobilePhone",[string]$userObj.MobilePhone)
        $param=$SqlCmd.Parameters.AddWithValue("@PostalCode",[string]$userObj.PostalCode)
        #$param=$SqlCmd.Parameters.AddWithValue("@PasswordNeverExpires",[string]$userObj.PasswordNeverExpires)
        #$param=$SqlCmd.Parameters.AddWithValue("@OverallProvisioningStatus",[string]$userObj.OverallProvisioningStatus)
        $param=$SqlCmd.Parameters.AddWithValue("@Title",[string]$userObj.JobTitle)
        #$param=$SqlCmd.Parameters.AddWithValue("@ValidationStatus",[string]$userObj.ValidationStatus)        
        #$param=$SqlCmd.Parameters.AddWithValue("@LiveId",[string]$userObj.LiveId)
        #$param=$SqlCmd.Parameters.AddWithValue("@MSRtcSipPrimaryUserAddress",[string]$userObj.MSRtcSipPrimaryUserAddress)
        
        

        $SqlCmd.Parameters.Add($ret_Status) >> $null;
        $SqlCmd.Parameters.Add($ret_Message) >> $null;
        $SqlCmd.Parameters.Add($ret_Operation) >> $null;
        $SqlCmd.Parameters.Add($ret_userExistance) >> $null;
         
        #
        $res=$SqlCmd.ExecuteNonQuery()
                
        $retStatus=$SqlCmd.Parameters["@Ret_Status"].Value; 
        $retMsg=$SqlCmd.Parameters["@Ret_Message"].Value;
        $retOperation=$SqlCmd.Parameters["@Ret_Operation"].Value;
        $retUserExistance=$SqlCmd.Parameters["@ret_userExistance"].Value; 
               
        $userObj.OperationStatus=$retStatus 
        $userObj.Operation= $retOperation
        $userObj.AdditionalInfo=$retMsg
                
           
        if($retUserExistance -eq 1)
        {
            $script:usersWithSameSigninName+=[pscustomobject]@{SigninName=$userObj.SigninName; ICName=$userObj.ICName; WhenCreated=$userObj.Created}
        }
    }
    catch
    {
        LogWrite -Level ERROR -Message  "Error executing query. Error Info $($_)"
    }
}
