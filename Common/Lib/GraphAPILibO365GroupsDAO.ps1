function UpdateO365GroupsToDatabase {    
    if ($script:o365GroupsData) {
        LogWrite -Message "Updating Active M365 Groups to Database..."
        UpdateSQLO365Groups $script:ConnectionString $script:o365GroupsData
    }
    if ($script:o365DeletedGroupsData) {
        LogWrite -Message "Updating InActive M365 Groups to Database..."
        UpdateSQLO365Groups $script:ConnectionString $script:o365DeletedGroupsData
    }
    if ($script:o365TeamsData) {
        LogWrite -Message "Updating Active Teams to Database..."
        UpdateSQLTeams $script:ConnectionString $script:o365TeamsData
    }
    if ($script:TeamsChannelData) {
        LogWrite -Message "Updating Active Teams Channel to Database..."
        UpdateTeamsChannel $script:ConnectionString $script:TeamsChannelData
    }
    
    LogWrite -Message "Delete invalid groups/teams/teamchannel from DB..."      
    $syncDate = Get-Date -format "yyyy-MM-dd"    
    DeleteSQLDeletedO365Groups $script:connectionString $syncDate   
    #DeleteInvalidTeamChannel $script:connectionString $syncDate
}
function DeleteSQLDeletedO365Groups {
    param($connectionString,$SyncDate)  
    try {
        #Initialize SQL Connections
        $SqlConnection = New-Object System.Data.SqlClient.SqlConnection
        $SqlConnection.ConnectionString = $connectionString   
        $SqlConnection.Open()    
    
        # initialize stored procedure
        $SqlCmd = New-Object System.Data.SqlClient.SqlCommand
        $SqlCmd.CommandType = [System.Data.CommandType]'StoredProcedure'
        $SqlCmd.CommandText = "DeleteGroups30Period"
        $SqlCmd.Connection = $SqlConnection
        $SqlCmd.Parameters.AddWithValue("SyncDate", $SyncDate)
        $res = $SqlCmd.ExecuteNonQuery()
    }
    catch {
        LogWrite -Level ERROR -Message "Permanently deleted group/team/teamchannel info to DB: $($_)"
    }
    finally {
        #Close Connection
        $SqlConnection.Close()
    }
}
#region Permanently delete team channel
Function DeleteInvalidTeamChannel {
    param($connectionString,$SyncDate)   
    #Initialize SQL Connections
    $SqlConnection = New-Object System.Data.SqlClient.SqlConnection
    $SqlConnection.ConnectionString = $connectionString   
    $SqlConnection.Open()    
    try {
        # initialize stored procedure
        $SqlCmd = New-Object System.Data.SqlClient.SqlCommand
        $SqlCmd.CommandType = [System.Data.CommandType]'StoredProcedure'
        $SqlCmd.CommandText = "DeleteInvalidTeamChannel"
        $SqlCmd.Connection = $SqlConnection
        $SqlCmd.Parameters.AddWithValue("SyncDate", $SyncDate)
        $res = $SqlCmd.ExecuteNonQuery()
    }
    catch {
        Write-Log "Permanently team channel info DB: $($_)"
    }
    finally{
        #Close Connection        
        $SqlCmd.Dispose()                     
        $SqlConnection.Dispose()
        $SqlConnection.Close()  
    }
}
#endregion
function UpdateSQLO365Groups {
    <#
      .Synopsis
        Update list of groups to DB      
    #>
    param($connectionString, $groupsData)
    
    if ($null -ne $groupsData) {
        #Initialize SQL Connections
        $SqlConnection = New-Object System.Data.SqlClient.SqlConnection
        $SqlConnection.ConnectionString = $connectionString   
        $SqlConnection.Open()

        $i = 0
        $count = $groupsData.Count        
        
        foreach ($groupObj in $groupsData) {
            if ($null -ne $groupObj -and $groupObj.GroupId -ne "") { 
                UpdateO365GroupRecord $SqlConnection $groupObj                                               
                $i++     
                LogWrite -Message "($($i)/$($count)): $($groupObj.GroupId)"
                
            }
        }
        #Close Connection
        $SqlConnection.Close()
    }         
}
function UpdateO365GroupRecord {
    param($SqlConnection, $groupObj)
    
    try {
        # initialize stored procedure
        $SqlCmd = New-Object System.Data.SqlClient.SqlCommand
        $SqlCmd.CommandType = [System.Data.CommandType]'StoredProcedure'
        $SqlCmd.CommandText = "SetGroupInfo"
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

        $grpId = $groupObj.GroupID
       
        $SqlCmd.Parameters.AddWithValue("GroupId", [string]$groupObj.GroupID)
        $SqlCmd.Parameters.AddWithValue("DisplayName", [string]$groupObj.DisplayName)
        $SqlCmd.Parameters.AddWithValue("Description", [string]$groupObj.Description)
        $SqlCmd.Parameters.AddWithValue("GroupOwners", $null)
        $SqlCmd.Parameters.AddWithValue("GroupMembers", $null)
        $SqlCmd.Parameters.AddWithValue("GroupGuests", $null)
        if ($groupObj.GroupOwners -ne '') {
            $SqlCmd.Parameters["GroupOwners"].Value = $groupObj.GroupOwners
        }        
        if ($groupObj.GroupMembers -ne '') {
            $SqlCmd.Parameters["GroupMembers"].Value = $groupObj.GroupMembers            
        } 
        if ($groupObj.GroupGuests -ne '') {
            $SqlCmd.Parameters["GroupGuests"].Value = $groupObj.GroupGuests            
        }       
        $SqlCmd.Parameters.AddWithValue("Classification", [string]$groupObj.Classification)
        $SqlCmd.Parameters.AddWithValue("CreatedDateTime", [string]$groupObj.CreatedDateTime)         
        $SqlCmd.Parameters.AddWithValue("DeletedDateTime", $null)        
        if ($groupObj.DeletedDateTime -ne '' -and $groupObj.DeletedDateTime -ne $null) {
            $SqlCmd.Parameters["DeletedDateTime"].Value = [string]$groupObj.DeletedDateTime
        }         
        $param = $SqlCmd.Parameters.AddWithValue("CreationOption", [string]$groupObj.CreationOptions)
        $param = $SqlCmd.Parameters.AddWithValue("IsAssignableToRole", [string]$groupObj.IsAssignableToRole)
        $param = $SqlCmd.Parameters.AddWithValue("Mail", [string]$groupObj.Mail)
        $param = $SqlCmd.Parameters.AddWithValue("MailNickname", [string]$groupObj.MailNickname)
        $param = $SqlCmd.Parameters.AddWithValue("MailEnabled", [string]$groupObj.MailEnabled)
        $param = $SqlCmd.Parameters.AddWithValue("ProxyAddresses", [string]$groupObj.ProxyAddresses)
        $param = $SqlCmd.Parameters.AddWithValue("RenewedDateTime", [string]$groupObj.RenewedDateTime)
        $param = $SqlCmd.Parameters.AddWithValue("ResourceBehaviorOptions", [string]$groupObj.ResourceBehaviorOptions)
        $param = $SqlCmd.Parameters.AddWithValue("ResourceProvisioningOptions", [string]$groupObj.ResourceProvisioningOptions)
        $param = $SqlCmd.Parameters.AddWithValue("SecurityEnabled", [string]$groupObj.SecurityEnabled)
        $param = $SqlCmd.Parameters.AddWithValue("SecurityIdentifier", [string]$groupObj.SecurityIdentifier)
        $param = $SqlCmd.Parameters.AddWithValue("Visibility", [string]$groupObj.Visibility)
        $param = $SqlCmd.Parameters.AddWithValue("OnPremisesLastSyncDateTime", [string]$groupObj.OnPremisesLastSyncDateTime)
        $param = $SqlCmd.Parameters.AddWithValue("OnPremisesSamAccountName", [string]$groupObj.OnPremisesSamAccountName)
        $param = $SqlCmd.Parameters.AddWithValue("OnPremisesSecurityIdentifier", [string]$groupObj.OnPremisesSecurityIdentifier)
        $param = $SqlCmd.Parameters.AddWithValue("OnPremisesSyncEnabled", [string]$groupObj.OnPremisesSyncEnabled)
        $param = $SqlCmd.Parameters.AddWithValue("AllowExternalSenders", [string]$groupObj.AllowExternalSenders)
        $param = $SqlCmd.Parameters.AddWithValue("AutoSubscribeNewMembers", [string]$groupObj.AutoSubscribeNewMembers)
        $param = $SqlCmd.Parameters.AddWithValue("HideFromAddressLists", [string]$groupObj.HideFromAddressLists)
        $param = $SqlCmd.Parameters.AddWithValue("HideFromOutlookClients", [string]$groupObj.HideFromOutlookClients)
        #PermanentDeletionDate      
        
        $SqlCmd.Parameters.Add($ret_Status) >> $null;
        $SqlCmd.Parameters.Add($ret_Message) >> $null;
        $SqlCmd.Parameters.Add($ret_Operation) >> $null;
        
        $SqlCmd.ExecuteNonQuery()

        $retStatus = $SqlCmd.Parameters["@Ret_Status"].Value; 
        $retMsg = $SqlCmd.Parameters["@Ret_Message"].Value;
        $retOperation = $SqlCmd.Parameters["@Ret_Operation"].Value;
        
        $groupObj.Operation = $retOperation
        $groupObj.OperationStatus = $retStatus
        $groupObj.AdditionalInfo = $retMsg
        
        if ($retStatus -eq "Failed") {
            LogWrite -Level ERROR -Message "Failed for $($groupObj.GroupId): $($retMsg)"
        }
        
    }
    catch {
        LogWrite -Level ERROR -Message "Adding the Group info to DB: $($grpId) - $($_)"
    }  
}
function UpdateSQLO365Group {
    <#
      .Synopsis
        Update a group to DB      
    #>
    param($connectionString, $groupData)
    if ($groupData) {
        #Initialize SQL Connections
        $SqlConnection = New-Object System.Data.SqlClient.SqlConnection
        $SqlConnection.ConnectionString = $connectionString   
        $SqlConnection.Open()
        UpdateO365GroupRecord $SqlConnection $groupData
        #Close Connection
        $SqlConnection.Close()
    }           
}

function UpdateSQLTeams {
    param($connectionString, $teamsData)
   
    if ($null -ne $teamsData) {
        #Initialize SQL Connections
        $SqlConnection = New-Object System.Data.SqlClient.SqlConnection
        $SqlConnection.ConnectionString = $connectionString   
        $SqlConnection.Open()

        $i = 0
        $count = $teamsData.Count
        
        foreach ($teamObj in $teamsData) {
            if ($teamObj -ne $null) {
                UpdateTeamRecord $SqlConnection $teamObj                                
                $i++                
                LogWrite -Message "($($i)/$($count)): $($teamObj.GroupID)"
            }
        }
        #Close Connection
        $SqlConnection.Close()
    }         
}
function UpdateTeamRecord {
    param($SqlConnection, $groupObj)
    
    try {
        # initialize stored procedure
        $SqlCmd = New-Object System.Data.SqlClient.SqlCommand
        $SqlCmd.CommandType = [System.Data.CommandType]'StoredProcedure'
        $SqlCmd.CommandText = "SetTeamInfo"
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
       
        $param = $SqlCmd.Parameters.AddWithValue("GroupId", [string]$groupObj.GroupID)
        $param = $SqlCmd.Parameters.AddWithValue("DisplayName", [string]$groupObj.DisplayName)
        $param = $SqlCmd.Parameters.AddWithValue("Description", [string]$groupObj.Description)
        $param = $SqlCmd.Parameters.AddWithValue("Classification", [string]$groupObj.Classification)
        $param = $SqlCmd.Parameters.AddWithValue("Specialization", [string]$groupObj.Specialization)
        $param = $SqlCmd.Parameters.AddWithValue("InternalId", [string]$groupObj.InternalId)
        $param = $SqlCmd.Parameters.AddWithValue("WebUrl", [string]$groupObj.WebUrl)        
        $param = $SqlCmd.Parameters.AddWithValue("IsArchived", [string]$groupObj.IsArchived)
        $param = $SqlCmd.Parameters.AddWithValue("AllowMemberCreateUpdateChannels", [string]$groupObj.AllowMemberCreateUpdateChannels)
        $param = $SqlCmd.Parameters.AddWithValue("AllowMemberCreatePrivateChannels", [string]$groupObj.AllowMemberCreatePrivateChannels)
        $param = $SqlCmd.Parameters.AddWithValue("AllowMemberDeleteChannels", [string]$groupObj.AllowMemberDeleteChannels)
        $param = $SqlCmd.Parameters.AddWithValue("AllowMemberAddRemoveApps", [string]$groupObj.AllowMemberAddRemoveApps)
        $param = $SqlCmd.Parameters.AddWithValue("AllowMemberCreateUpdateRemoveTabs", [string]$groupObj.AllowMemberCreateUpdateRemoveTabs)
        $param = $SqlCmd.Parameters.AddWithValue("AllowMemberCreateUpdateRemoveConnectors", [string]$groupObj.AllowMemberCreateUpdateRemoveConnectors)
        $param = $SqlCmd.Parameters.AddWithValue("AllowGuestCreateUpdateChannels", [string]$groupObj.AllowGuestCreateUpdateChannels)
        $param = $SqlCmd.Parameters.AddWithValue("AllowGuestDeleteChannels", [string]$groupObj.AllowGuestDeleteChannels)
        $param = $SqlCmd.Parameters.AddWithValue("AllowUserEditMessages", [string]$groupObj.AllowUserEditMessages)
        $param = $SqlCmd.Parameters.AddWithValue("AllowUserDeleteMessages", [string]$groupObj.AllowUserDeleteMessages)
        $param = $SqlCmd.Parameters.AddWithValue("AllowOwnerDeleteMessages", [string]$groupObj.AllowOwnerDeleteMessages)
        $param = $SqlCmd.Parameters.AddWithValue("AllowTeamMentions", [string]$groupObj.AllowTeamMentions)
        $param = $SqlCmd.Parameters.AddWithValue("AllowChannelMentions", [string]$groupObj.AllowChannelMentions)
        $param = $SqlCmd.Parameters.AddWithValue("AllowGiphy", [string]$groupObj.AllowGiphy)        
        $param = $SqlCmd.Parameters.AddWithValue("GiphyContentRating", [string]$groupObj.GiphyContentRating)        
        $param = $SqlCmd.Parameters.AddWithValue("AllowStickersAndMemes", [string]$groupObj.AllowStickersAndMemes)        
        $param = $SqlCmd.Parameters.AddWithValue("AllowCustomMemes", [string]$groupObj.AllowCustomMemes)        
        $param = $SqlCmd.Parameters.AddWithValue("ShowInTeamsSearchAndSuggestions", [string]$groupObj.ShowInTeamsSearchAndSuggestions)               
        
        $SqlCmd.Parameters.Add($ret_Status) >> $null;
        $SqlCmd.Parameters.Add($ret_Message) >> $null;
        $SqlCmd.Parameters.Add($ret_Operation) >> $null;
        
        $res = $SqlCmd.ExecuteNonQuery()

        $retStatus = $SqlCmd.Parameters["@Ret_Status"].Value; 
        $retMsg = $SqlCmd.Parameters["@Ret_Message"].Value;
        $retOperation = $SqlCmd.Parameters["@Ret_Operation"].Value;
        
        $groupObj.Operation = $retOperation
        $groupObj.OperationStatus = $retStatus
        $groupObj.AdditionalInfo = $retMsg
        if ($retStatus -eq "Failed") {
            LogWrite -Message "Failed for $($groupObj.GroupId): $($retMsg)"
        }        
    }
    catch {
        LogWrite -Level ERROR -Message "Adding the Team info to Database: $($_)"
    }  
}
function UpdateSQLTeam {
    <#
      .Synopsis
        Update a group to DB      
    #>
    param($connectionString, $groupData)
    if ($groupData) {
        #Initialize SQL Connections
        $SqlConnection = New-Object System.Data.SqlClient.SqlConnection
        $SqlConnection.ConnectionString = $connectionString   
        $SqlConnection.Open()
        UpdateTeamRecord $SqlConnection $groupData
        #Close Connection
        $SqlConnection.Close()
    }           
}

function UpdateTeamsChannel {
    param($connectionString, $teamsChannelData)
   
    if ($teamsChannelData -ne $null) {
        #Initialize SQL Connections
        $SqlConnection = New-Object System.Data.SqlClient.SqlConnection
        $SqlConnection.ConnectionString = $connectionString   
        $SqlConnection.Open()

        $i = 0
        $count = $teamsChannelData.Count
        
        foreach ($teamChannelObj in $teamsChannelData) {
            if ($teamChannelObj -ne $null) {
                UpdateTeamChannelRecord $SqlConnection $teamChannelObj                              
                $i++
                
                LogWrite -Message "($($i)/$($count)): $($teamChannelObj.Id)"
            }
        }

        #Close Connection
        $SqlConnection.Close()
    }         
}
function UpdateTeamChannelRecord {
    param($SqlConnection, $channelObj)
    
    try {
        # initialize stored procedure
        $SqlCmd = New-Object System.Data.SqlClient.SqlCommand
        $SqlCmd.CommandType = [System.Data.CommandType]'StoredProcedure'
        $SqlCmd.CommandText = "SetTeamChannelInfo"
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

        $param = $SqlCmd.Parameters.AddWithValue("GroupId", [string]$channelObj.GroupID)
        $param = $SqlCmd.Parameters.AddWithValue("Id", [string]$channelObj.Id)
        $param = $SqlCmd.Parameters.AddWithValue("DisplayName", [string]$channelObj.DisplayName)
        $param = $SqlCmd.Parameters.AddWithValue("Description", [string]$channelObj.Description)
        $param = $SqlCmd.Parameters.AddWithValue("IsFavoriteByDefault", [string]$channelObj.IsFavoriteByDefault)
        $param = $SqlCmd.Parameters.AddWithValue("Email", [string]$channelObj.Email)               
        $param = $SqlCmd.Parameters.AddWithValue("WebUrl", [string]$channelObj.WebUrl)               
        $param = $SqlCmd.Parameters.AddWithValue("MembershipType", [string]$channelObj.MembershipType)               
        $param = $SqlCmd.Parameters.AddWithValue("ChannelOwners", [string]$channelObj.ChannelOwners)               
        $param = $SqlCmd.Parameters.AddWithValue("ChannelMembers", [string]$channelObj.ChannelMembers)
        $param = $SqlCmd.Parameters.AddWithValue("ChannelGuests", [string]$channelObj.ChannelGuests)
        
        $SqlCmd.Parameters.Add($ret_Status) >> $null;
        $SqlCmd.Parameters.Add($ret_Message) >> $null;
        $SqlCmd.Parameters.Add($ret_Operation) >> $null;
        
        $res = $SqlCmd.ExecuteNonQuery()

        $retStatus = $SqlCmd.Parameters["@Ret_Status"].Value; 
        $retMsg = $SqlCmd.Parameters["@Ret_Message"].Value;
        $retOperation = $SqlCmd.Parameters["@Ret_Operation"].Value;
        
        $channelObj.Operation = $retOperation
        $channelObj.OperationStatus = $retStatus
        $channelObj.AdditionalInfo = $retMsg

        if ($retStatus -eq "Failed") {
            LogWrite -Message "Failed for $($channelObj.Id): $($retMsg)"
        }
        
    }
    catch {
        LogWrite -Level ERROR -Message "Adding the Team Channel info to DB: $($_)"
    }  
}

Function GetOrphanedTeams {
    Param(
        [Parameter(Mandatory=$true)]$connectionString,
        [Parameter(Mandatory=$false)]$AdmSvc
    )
    Process
    {
        $SqlConnection = New-Object System.Data.SqlClient.SqlConnection
        $SqlConnection.ConnectionString = $connectionString   

        $SqlCmd = New-Object System.Data.SqlClient.SqlCommand
        $SqlCmd.CommandType = [System.Data.CommandType]'StoredProcedure'
        $SqlCmd.CommandText = "GetOrphanedTeams"
        $SqlCmd.Connection = $SqlConnection
        $SqlCmd.Parameters.AddWithValue("AdmSvc", $AdmSvc) | Out-Null
        
        $SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter 
        $SqlAdapter.SelectCommand = $SqlCmd
        $DataSet = New-Object System.Data.DataSet
        $rowCount =$SqlAdapter.Fill($DataSet) | Out-Null
        $OrphanedTeams = $dataset.Tables[0] 

        try
        {
            $SqlConnection.Open()
            return $OrphanedTeams
        }
        catch [Exception]
        {
           LogWrite -Level ERROR -Message "Error connecting to Database: $($_.Exception.Message)" 
        }
        finally
        {
            $SqlConnection.Close()
            $SqlCmd.Dispose()
            $SqlConnection.Dispose()
        }
    }
}
Function GetOrphanedGroups {
    Param(
        [Parameter(Mandatory=$true)]$connectionString,
        [Parameter(Mandatory=$false)]$AdmSvc
    )
    Process
    {
        try
        {
            $SqlConnection = New-Object System.Data.SqlClient.SqlConnection
            $SqlConnection.ConnectionString = $connectionString   

            $SqlCmd = New-Object System.Data.SqlClient.SqlCommand
            $SqlCmd.CommandType = [System.Data.CommandType]'StoredProcedure'
            $SqlCmd.CommandText = "GetOrphanedGroups"
            $SqlCmd.Connection = $SqlConnection
            $SqlCmd.Parameters.AddWithValue("AdmSvc", $AdmSvc) | Out-Null
        
            $SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter 
            $SqlAdapter.SelectCommand = $SqlCmd
            $DataSet = New-Object System.Data.DataSet
            $SqlAdapter.Fill($DataSet) | Out-Null
        
            $SqlConnection.Open()
            $SqlConnection.Close()
            $SqlCmd.Dispose()
            $SqlConnection.Dispose()
        
            return $DataSet

            <#ForEach($table in $DataSet.Tables) {
                $table |Out-GridView -PassThru
            }#> 
        }      
        
        catch [Exception] {
           LogWrite -Level ERROR -Message "Error connecting to Database: $($_.Exception.Message)" 
        }
    }
}

#region Post Change Request
Function UpdateGroupTeamPostChangeRequest {  
    param(
        [Parameter(Mandatory=$true)] $ConnectionString,
        [Parameter(Mandatory=$false)]$teamObj,
        [Parameter(Mandatory=$false)]$HideFromOutlookClients,
        [Parameter(Mandatory=$false)]$HideFromAddressLists
    )
   
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

            $groupId = $teamObj.GroupId
            if ($groupId -eq $null){
                $groupId = $teamObj.Id
            }
        
            $SqlCmd.Parameters.AddWithValue("GroupId", $groupId)
            $SqlCmd.Parameters.AddWithValue("DisplayName", [string]$teamObj.DisplayName)
            $SqlCmd.Parameters.AddWithValue("Description", [string]$teamObj.Description)
            $SqlCmd.Parameters.AddWithValue("Visibility", [string]$teamObj.Visibility)
            $SqlCmd.Parameters.AddWithValue("HideFromOutlookClients", $HideFromOutlookClients)
            $SqlCmd.Parameters.AddWithValue("HideFromAddressLists", $HideFromAddressLists)

            <#
            $SqlCmd.Parameters.AddWithValue("HideFromAddressLists", $null)        
            if ($teamObj.HideFromAddressLists -ne '' -and $teamObj.HideFromAddressLists -ne $null) {
                $SqlCmd.Parameters["HideFromAddressLists"].Value = [string]$teamObj.HideFromAddressLists
            }        
            $SqlCmd.Parameters.AddWithValue("HideFromOutlookClients", $null)        
            if ($teamObj.HideFromOutlookClients -ne '' -and $teamObj.HideFromOutlookClients -ne $null) {
                $SqlCmd.Parameters["HideFromOutlookClients"].Value = [string]$teamObj.HideFromOutlookClients
            }     
            #>       
            $res = $SqlCmd.ExecuteNonQuery()
                        
            LogWrite -Message "Post-ChangeRequest: Update group/team into Groups and Teams table."  

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
#endregion

