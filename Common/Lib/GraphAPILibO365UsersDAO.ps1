﻿Function UpdateO365UsersToDatabase {
    $updateStartTime = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    
    #Update Active Users to Database...
    LogWrite -Message "Updating Active M365 Users to DB..."
    UpdateSQLO365Users $script:connectionString $script:usersData
    
    #Update Deleted Users to Database...
    LogWrite -Message "Updating Deleted M365 Users to DB..."
    UpdateSQLO365Users $script:connectionString $script:deletedUsersData

    #Update Guests to Database...
    LogWrite -Message "Updating Guests Users to DB..."
    UpdateSQLGuests $script:connectionString $script:guestsData
    
    #Remove permanently users from Users - DB
    LogWrite -Message "Delete Permanently Deleted Users from DB..."      
    $syncDate = Get-Date -format "yyyy-MM-dd"
    DeleteInvalidUsers $script:connectionString $syncDate
    $updateEndTime = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    LogWrite -Message "Update M365 Users To DB Start Time: $($updateStartTime)"
    LogWrite -Message "Update M365 Users To DB End Time: $($updateEndTime)"

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
            LogWrite -Level ERROR -Message "Connecting to DB: $($_)"
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
        
        $SqlCmd.CommandText = "SetUserInfo"
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
        $param=$SqlCmd.Parameters.AddWithValue("@ICName",[string]$userObj.ICName)
        $param=$SqlCmd.Parameters.AddWithValue("@Department",[string]$userObj.Department)
        $param=$SqlCmd.Parameters.AddWithValue("@Office",[string]$userObj.OfficeLocation)
        $param=$SqlCmd.Parameters.AddWithValue("@CompanyName",[string]$userObj.CompanyName)
        $param=$SqlCmd.Parameters.AddWithValue("@Email",[string]$userObj.Mail)
        $param=$SqlCmd.Parameters.AddWithValue("@DisplayName",[string]$userObj.DisplayName)        
        $param=$SqlCmd.Parameters.AddWithValue("@IsDisabled",$userObj.IsDisabled)
        $param=$SqlCmd.Parameters.AddWithValue("@IsDeleted",[int]$userObj.IsDeleted)
        $param=$SqlCmd.Parameters.AddWithValue("@IsLicensed",$userObj.IsLicensed)
        
        $SqlCmd.Parameters.AddWithValue("Licenses", $null)        
        if ($userObj.AssignedLicenses -ne '' -and $userObj.AssignedLicenses -ne $null) {
            $SqlCmd.Parameters["Licenses"].Value = [string]$userObj.AssignedLicenses
        }
        $SqlCmd.Parameters.AddWithValue("AppLicenses", $null)        
        if ($userObj.AssignedPlans -ne '' -and $userObj.AssignedPlans -ne $null) {
            $SqlCmd.Parameters["AppLicenses"].Value = [string]$userObj.AssignedPlans
        }  

        #$param=$SqlCmd.Parameters.AddWithValue("@Licenses",[string]$userObj.AssignedLicenses)
        #$param=$SqlCmd.Parameters.AddWithValue("@AppLicenses",[string]$userObj.AssignedPlans)
        
        $SqlCmd.Parameters.AddWithValue("Created", $null)        
        if ($userObj.CreatedDateTime -ne '' -and $userObj.CreatedDateTime -ne $null) {
            $SqlCmd.Parameters["Created"].Value = [string]$userObj.CreatedDateTime
        }        
        $SqlCmd.Parameters.AddWithValue("SoftDeletionTimestamp", $null)        
        if ($userObj.SoftDeletionTimestamp -ne '' -and $userObj.SoftDeletionTimestamp -ne $null) {
            $SqlCmd.Parameters["SoftDeletionTimestamp"].Value = [string]$userObj.SoftDeletionTimestamp
        }        
        $SqlCmd.Parameters.AddWithValue("LastDirSyncTime", $null)        
        if ($userObj.LastDirSyncTime -ne '' -and $userObj.LastDirSyncTime -ne $null) {
            $SqlCmd.Parameters["LastDirSyncTime"].Value = [string]$userObj.LastDirSyncTime
        }
        $param=$SqlCmd.Parameters.AddWithValue("@UserPrincipalName",[string]$userObj.UserPrincipalName)
        $param=$SqlCmd.Parameters.AddWithValue("@LastPasswordChangeTimeStamp",[string]$userObj.lastPasswordChangeDateTime)
        #$param=$SqlCmd.Parameters.AddWithValue("@StsRefreshTokensValidFrom",[string]$userObj.StsRefreshTokensValidFrom)
        $param=$SqlCmd.Parameters.AddWithValue("@StreetAddress",[string]$userObj.StreetAddress)
        $param=$SqlCmd.Parameters.AddWithValue("@City",[string]$userObj.City)
        $param=$SqlCmd.Parameters.AddWithValue("@State",[string]$userObj.State)
        $param=$SqlCmd.Parameters.AddWithValue("@Country",[string]$userObj.Country)
        $param=$SqlCmd.Parameters.AddWithValue("@UserType",[string]$userObj.UserType)
        <# not in used
        $param=$SqlCmd.Parameters.AddWithValue("@CreationType",[string]$userObj.CreationType)
        $param=$SqlCmd.Parameters.AddWithValue("@ExternalUserState",[string]$userObj.ExternalUserState)        
        $SqlCmd.Parameters.AddWithValue("ExternalUserStateChangeDateTime", $null)        
        if ($userObj.ExternalUserStateChangeDateTime -ne '' -and $userObj.ExternalUserStateChangeDateTime -ne $null) {
            $SqlCmd.Parameters["ExternalUserStateChangeDateTime"].Value = [string]$userObj.ExternalUserStateChangeDateTime
        }
        #>       
        $SqlCmd.Parameters.AddWithValue("LastSignInDateTime", $null)        
        if ($userObj.LastSignInDateTime -ne '' -and $userObj.LastSignInDateTime -ne $null) {
            $SqlCmd.Parameters["LastSignInDateTime"].Value = [string]$userObj.LastSignInDateTime
        }
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
            $script:usersWithSameSigninName+=[pscustomobject]@{
                SigninName=$userObj.SigninName; 
                IsDeleted=$userObj.IsDeleted; 
                IsDisabled=$userObj.IsDisabled; 
                IsLicensed=$userObj.IsLicensed; 
                ICName=$userObj.ICName; 
                Department=$userObj.Department;
                Office=$userObj.Office;
                CompanyName=$userObj.CompanyName;
                WhenCreated=$userObj.Created
                SoftDeletionTimestamp=$userObj.SoftDeletionTimestamp
            }
        }
    }
    catch
    {
        LogWrite -Level ERROR -Message  "Adding user info to DB: $userObj - $($_)"
    }
}

Function UpdateSQLGuests {
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
                    UpdateSQLGuestsRecord $SqlConnection $userObj
                    $i++
                
                    LogWrite -Message "($($i)/$($count)): $($userObj.UserPrincipalName)"
                }
            }
        }
        catch {
            LogWrite -Level ERROR -Message "Connecting to DB: $($_)"
        }
        
        finally{            
            $SqlConnection.Close()
        }        
    }         
}

Function UpdateSQLGuestsRecord {
    param($SqlConnection,$userObj)
    $retStatus=$null
    $retMsg=$null
    $retOperation=$null
    $retUserExistance = $null
    
    try {
        $SqlCmd = New-Object System.Data.SqlClient.SqlCommand

        # indicate that we are working with stored procedure
        $SqlCmd.CommandType=[System.Data.CommandType]'StoredProcedure'
        
        $SqlCmd.CommandText = "SetGuestInfo"
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
        $param=$SqlCmd.Parameters.AddWithValue("@DisplayName",[string]$userObj.DisplayName)
        $param=$SqlCmd.Parameters.AddWithValue("@UserPrincipalName", [string]$userObj.UserPrincipalName)                
        $param=$SqlCmd.Parameters.AddWithValue("@PrimaryEmail",[string]$userObj.PrimaryEmail)
        $param=$SqlCmd.Parameters.AddWithValue("@OtherEmails",[string]$userObj.OtherEmails)
        $SqlCmd.Parameters.AddWithValue("Created", $null)        
        if ($userObj.CreatedDateTime -ne '' -and $userObj.CreatedDateTime -ne $null) {
            $SqlCmd.Parameters["Created"].Value = [string]$userObj.CreatedDateTime
        }
        $SqlCmd.Parameters.AddWithValue("LastSignInDateTime", $null)        
        if ($userObj.LastSignInDateTime -ne '' -and $userObj.LastSignInDateTime -ne $null) {
            $SqlCmd.Parameters["LastSignInDateTime"].Value = [string]$userObj.LastSignInDateTime
        }
        $param=$SqlCmd.Parameters.AddWithValue("@CreationType",[string]$userObj.CreationType)
        $param=$SqlCmd.Parameters.AddWithValue("@ExternalUserState",[string]$userObj.ExternalUserState)        
        $SqlCmd.Parameters.AddWithValue("ExternalUserStateChangeDateTime", $null)        
        if ($userObj.ExternalUserStateChangeDateTime -ne '' -and $userObj.ExternalUserStateChangeDateTime -ne $null) {
            $SqlCmd.Parameters["ExternalUserStateChangeDateTime"].Value = [string]$userObj.ExternalUserStateChangeDateTime
        }
        $SqlCmd.Parameters.AddWithValue("LastPasswordChangeTimeStamp", $null)
        if ($userObj.LastPasswordChangeDateTime -ne '' -and $userObj.LastPasswordChangeDateTime -ne $null) {
            $SqlCmd.Parameters["LastPasswordChangeTimeStamp"].Value = [string]$userObj.LastPasswordChangeDateTime
        }       
        
        $SqlCmd.Parameters.AddWithValue("LastDirSyncTime", $null)        
        if ($userObj.OnPremisesLastSyncDateTime -ne '' -and $userObj.OnPremisesLastSyncDateTime -ne $null) {
            $SqlCmd.Parameters["LastDirSyncTime"].Value = [string]$userObj.OnPremisesLastSyncDateTime
        }
        $param=$SqlCmd.Parameters.AddWithValue("@IsLicensed",$userObj.IsLicensed)   
        $param=$SqlCmd.Parameters.AddWithValue("@IsDisabled",$userObj.IsDisabled)

        $SqlCmd.Parameters.Add($ret_Status) >> $null;
        $SqlCmd.Parameters.Add($ret_Message) >> $null;
        $SqlCmd.Parameters.Add($ret_Operation) >> $null;
        
        #
        $res=$SqlCmd.ExecuteNonQuery()
                
        $retStatus=$SqlCmd.Parameters["@Ret_Status"].Value; 
        $retMsg=$SqlCmd.Parameters["@Ret_Message"].Value;
        $retOperation=$SqlCmd.Parameters["@Ret_Operation"].Value;        
               
        $userObj.OperationStatus=$retStatus 
        $userObj.Operation= $retOperation
        $userObj.AdditionalInfo=$retMsg
    }
    catch
    {
        LogWrite -Level ERROR -Message  "Adding guests info to DB: $userObj - $($_)"
    }
}

#region Permanently delete user
Function DeleteInvalidUsers {
    param($connectionString,$SyncDate)   
    #Initialize SQL Connections
    $SqlConnection = New-Object System.Data.SqlClient.SqlConnection
    $SqlConnection.ConnectionString = $connectionString   
    $SqlConnection.Open()    
    try {
        # initialize stored procedure
        $SqlCmd = New-Object System.Data.SqlClient.SqlCommand
        $SqlCmd.CommandType = [System.Data.CommandType]'StoredProcedure'
        $SqlCmd.CommandText = "DeleteInvalidUsers"
        $SqlCmd.Connection = $SqlConnection
        $SqlCmd.Parameters.AddWithValue("SyncDate", $SyncDate)
        $res = $SqlCmd.ExecuteNonQuery()
    }
    catch {
        LogWrite -Level ERROR "Permanently soft deleted user info DB: $($_)"
    }
    finally{
        #Close Connection        
        $SqlCmd.Dispose()                     
        $SqlConnection.Dispose()
        $SqlConnection.Close()  
    }
}
#endregion

#region Provisioning
function GetSigninNameByEmail() {
    param($Email, $connectionString)    
    try {
        $SqlConnection = New-Object System.Data.SqlClient.SqlConnection
            $SqlConnection.ConnectionString = $connectionString   

            $SqlCmd = New-Object System.Data.SqlClient.SqlCommand
            $SqlCmd.CommandType = [System.Data.CommandType]'StoredProcedure'
            $SqlCmd.CommandText = "GetSigninNameByEmail"
            $SqlCmd.Connection = $SqlConnection
            $SqlCmd.Parameters.AddWithValue("UserEmail", $Email) | Out-Null
            #$SqlCmd.Parameters.AddWithValue("ICName", $ICName) | Out-Null
        
            $SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter 
            $SqlAdapter.SelectCommand = $SqlCmd
            $DataSet = New-Object System.Data.DataSet

            $SqlAdapter.Fill($DataSet) | Out-Null
            $SqlConnection.Open()
            return $DataSet.Tables[0]
    }
    catch {
        $exception = $_.Exception
        LogWrite -Level ERROR "Getting SigninName for email [$Email].Error Info: $exception"        
    }
    finally{
        $SqlConnection.Close()
        $SqlCmd.Dispose()
        $SqlConnection.Dispose()
    }
}
#endregion