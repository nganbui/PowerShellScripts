Function GetAllO365Groups {    
    $retrivalStartTime = Get-Date -Format "yyyy-MM-dd HH:mm:ss"       
    # Checking if authToken exists
    LogWrite -Message "Getting acces token using Graph API..."
    #Invoke-GraphAPIAuthTokenCheck

    #$script:Certificate = Get-Item Cert:\\LocalMachine\\My\* | Where-Object { $_.Subject -ieq "CN=$($script:appCertAdminPortalOperation)" }    
    #$script:authToken = Connect-NIHO365GraphWithCert -TenantName $script:TenantName -AppId $script:appIdAdminPortalOperation -Certificate $script:Certificate
    $script:authToken = Connect-GraphAPIWithCert -TenantId $script:TenantName -AppId $script:appIdAdminPortalOperation -Thumbprint $script:appThumbprintAdminPortalOperation

    if ($script:authToken) {
        LogWrite -Message  "Retrieving Active M365 Groups starting..."
        $script:o365GroupsData = @()
        $script:o365GroupsData = Get-NIHO365Groups -AuthToken $script:authToken                 
        if ($script:o365GroupsData) {
            $script:o365TeamsData = $script:o365GroupsData | Where-Object { "Team" -in $_.resourceProvisioningOptions }        
            #Parse Teams Channel
            LogWrite -Message  'Parsing [Teams Channel] to pscustomoject starting...'    
            $script:TeamsChannelData = ParseTeamChannel -AuthToken $script:authToken -Teams $script:o365TeamsData
            #Parse Groups
            LogWrite -Message  'Parsing [M365 Groups] to pscustomoject starting...'               
            $script:o365GroupsData = ParseO365Groups -AuthToken $script:authToken -Groups $script:o365GroupsData
            #Parse Teams
            LogWrite -Message  'Parsing [Teams] to pscustomoject starting...'    
            $script:o365TeamsData = ParseTeams -AuthToken $script:authToken -Teams $script:o365TeamsData
        }

        LogWrite -Message  "Retrieving Active M365 Groups completed."
    }
    
    #$script:authToken = Connect-NIHO365GraphWithCert -TenantName $script:TenantName -AppId $script:appIdAdminPortalOperation -Certificate $script:Certificate
    $script:authToken = Connect-GraphAPIWithCert -TenantId $script:TenantName -AppId $script:appIdAdminPortalOperation -Thumbprint $script:appThumbprintAdminPortalOperation

    if ($script:authToken) {
        LogWrite -Message  'Retrieving Deleted M365 Groups starting...'             
        $script:o365DeletedGroupsData = Get-NIHDeletedO365Groups -AuthToken $script:authToken        
        $script:o365DeletedGroupsData = ParseO365Groups -AuthToken $script:authToken -Groups $script:o365DeletedGroupsData
               
        LogWrite -Message  'Retrieving Deleted O365 Groups completed.'
    }
    $script:o365TeamsData = @($script:o365TeamsData)
    $script:TeamsChannelData = @($script:TeamsChannelData)
    $script:o365GroupsData = @($script:o365GroupsData)
    $script:o365DeletedGroupsData = @($script:o365DeletedGroupsData)

    $retrivalEndTime = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    LogWrite -Message "Retrieval O365 Groups Start Time: $($retrivalStartTime)"
    LogWrite -Message "Retrieval O365 Groups End Time: $($retrivalEndTime)"    
}

Function CacheO365Groups {
    LogWrite -Message "Generating Cache files for M365 Groups..."     
    if ($script:o365GroupsData) {
        SetDataInCache -CacheType O365 -ObjectType GraphAPIGroups -ObjectState Active -CacheData $script:o365GroupsData
        #make a copy to logs
        #Export_CSV -DataSet $script:o365GroupsData -FileName $groupsFile
    }
    if ($script:o365TeamsData) { 
        SetDataInCache -CacheType O365 -ObjectType GraphAPITeams -ObjectState Active -CacheData $script:o365TeamsData              
    }
    if ($script:TeamsChannelData) { 
        SetDataInCache -CacheType O365 -ObjectType GraphAPIChannel -ObjectState Active -CacheData $script:TeamsChannelData
    }
    if ($script:o365DeletedGroupsData) {
        SetDataInCache -CacheType O365 -ObjectType GraphAPIGroups -ObjectState InActive -CacheData $script:o365DeletedGroupsData       
    }
    LogWrite -Message "Generating Cache files for M365 Groups completed."        
}

Function ParseTeams {
    [cmdletBinding()]
    param(
        [parameter(Mandatory = $false)]
        [hashtable]$AuthToken,
        [parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        $Teams     
    )    
    [System.Collections.ArrayList]$TeamsList = @()
    $Teams = @($Teams) 
    if ($Teams -and $Teams.Count -gt 0) {
        $Teams | & { process {
                $Id = $_.id
                if ($Id) {
                    LogWrite -Message  "Processing the team [$Id]..."                    
                    $script:authToken = Connect-GraphAPIWithCert -TenantId $script:TenantName -AppId $script:appIdAdminPortalOperation -Thumbprint $script:appThumbprintAdminPortalOperation

                    if ($script:authToken) {
                        $r = Get-NIHTeam -AuthToken $script:authToken -Id $Id
                    }
                }
                if ($r) {
                    $null = $TeamsList.Add([PSCustomObject]@{
                            GroupID                                 = $Id                                       
                            DisplayName                             = $_.displayName
                            Description                             = $_.description
                            Classification                          = $_.classification
                            CreatedDateTime                         = $_.createdDateTime
                            DeletedDateTime                         = $_.deletedDateTime                    
                            Mail                                    = $_.mail 
                            MailNickname                            = $_.mailNickname
                            MailEnabled                             = $_.mailEnabled                                                          
                            Visibility                              = $_.visibility
                            #                                    
                            InternalId                              = $r.internalId;
                            Specialization                          = $r.specialization;                
                            WebUrl                                  = $r.webUrl;
                            IsArchived                              = $r.isArchived;
                            # #memberSettings
                            AllowMemberCreateUpdateChannels         = $r.memberSettings.allowCreateUpdateChannels;
                            AllowMemberCreatePrivateChannels        = $r.memberSettings.allowCreatePrivateChannels;
                            AllowMemberDeleteChannels               = $r.memberSettings.allowDeleteChannels;
                            AllowMemberAddRemoveApps                = $r.memberSettings.allowAddRemoveApps;
                            AllowMemberCreateUpdateRemoveTabs       = $r.memberSettings.allowCreateUpdateRemoveTabs;
                            AllowMemberCreateUpdateRemoveConnectors = $r.memberSettings.allowCreateUpdateRemoveConnectors;
                            #guestSettings
                            AllowGuestCreateUpdateChannels          = $r.guestSettings.allowCreateUpdateChannels;
                            AllowGuestDeleteChannels                = $r.guestSettings.allowDeleteChannels;
                            #messagingSettings
                            AllowUserEditMessages                   = $r.messagingSettings.allowUserEditMessages;
                            AllowUserDeleteMessages                 = $r.messagingSettings.allowUserDeleteMessages;
                            AllowOwnerDeleteMessages                = $r.messagingSettings.allowOwnerDeleteMessages;
                            AllowTeamMentions                       = $r.messagingSettings.allowTeamMentions;
                            AllowChannelMentions                    = $r.messagingSettings.allowChannelMentions;
                            #funSettings
                            AllowGiphy                              = $r.funSettings.allowGiphy;
                            GiphyContentRating                      = $r.funSettings.giphyContentRating;
                            AllowStickersAndMemes                   = $r.funSettings.allowStickersAndMemes;
                            AllowCustomMemes                        = $r.funSettings.allowCustomMemes; 
                            #discoverySettings
                            ShowInTeamsSearchAndSuggestions         = $r.discoverySettings.showInTeamsSearchAndSuggestions;
                            OperationStatus                         = ""; 
                            Operation                               = ""; 
                            AdditionalInfo                          = ""
                        });
                }
            }}
    }    
    return $TeamsList
} 

Function ParseTeamChannel {
    [cmdletBinding()]
    param(
        [parameter(Mandatory = $false)]
        [hashtable]$AuthToken,
        [parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        $Teams     
    )      
    [System.Collections.ArrayList]$ChannelList = @()        
    #$teamConn = ConnectMicrosoftTeams -TenantId $script:TenantId -ClientId $script:appIdOperationSupport -Thumbprint $script:appThumbprintOperationSupport
    #LogWrite -Level INFO -Message "Connecting Microsoft Teams..."
    $calls = 0
    $script:totalPrivateChannel = 0
    $Teams = @($Teams)
    if ($Teams -and $Teams.Count -gt 0) {        
        $i = 1
        $totalTeams = $Teams.Count        
        $Teams | & { process {
            $Id = $_.id              
            $TeamName = $_.DisplayName
            if ($Id) {
                LogWrite -Message  "($i/$totalTeams) : Processing the team [$Id] to get team channel..."                
                $script:authToken = Connect-GraphAPIWithCert -TenantId $script:TenantName -AppId $script:appIdAdminPortalOperation -Thumbprint $script:appThumbprintAdminPortalOperation

                if ($script:authToken) {                    
                    $c = Get-NIHTeamChannel -AuthToken $script:authToken -Id $Id
                    $calls++
                    if ($c -and $c[0].Id){
                        $teamAllMembersResponse = @()
                        $teamOwners = @()
                        $teamMembers = @()
                        $teamGuests = @()

                        $teamAllMembersResponse = Get-NIHTeamMembers -AuthToken $script:authToken -Id $Id
                        $calls++
                        $teamOwnersResponse = $teamAllMembersResponse | Where-Object { $_.roles -contains 'owner' }
                        $teamMembersResponse = $teamAllMembersResponse | Where-Object { $_.roles.Count -eq 0 }
                        $teamGuestsResponse = $teamAllMembersResponse | Where-Object { $_.roles -contains 'guest' }
                        
                        $teamOwners = $teamOwnersResponse.userId -join ";"; 
                        $teamMembers = $teamMembersResponse.userId -join ";"; 
                        $teamGuests = $teamGuestsResponse.userId -join ";";
                        $c=@($c)

                        $totalChannel =  $c.Count
                        $j = 1
                        foreach($itemChannel in $c) {
                            $channelId = $itemChannel.id
                            $channelDisplayName = $itemChannel.displayName
                            # For standard we do not need to call Get-NIHTeamChannelMembers
                            $ChannelOwners = $teamOwners
                            $ChannelMembers = $teamMembers
                            $ChannelGuests = $teamGuests
                            LogWrite -Message  "($j/$totalChannel) : Processing the channel [$channelId] - [$channelDisplayName]..."
                            #Get private channel owners/members/guests
                            if ($channelId -and $itemChannel.membershipType -eq 'Private') {
                                $script:totalPrivateChannel++
                                $script:authToken = Connect-GraphAPIWithCert -TenantId $script:TenantName -AppId $script:appIdAdminPortalOperation -Thumbprint $script:appThumbprintAdminPortalOperation
                                                                 
                                if ($script:authToken) {                                    
                                    $channelAllMembersResponse = $null
                                    $channelAllMembersResponse = Get-NIHTeamChannelMembers -AuthToken $script:authToken -GroupId $Id -ChannelId $channelId
                                    $calls++                                    
                                    $ownersReponse = $channelAllMembersResponse | Where-Object { $_.roles -contains 'owner' }
                                    $membersReponse = $channelAllMembersResponse | Where-Object { $_.roles.Count -eq 0 }
                                    $guestsReponse = $channelAllMembersResponse | Where-Object { $_.roles -contains 'guest' }

                                    $ChannelOwners = $null
                                    $ChannelMembers = $null
                                    $ChannelGuests = $null

                                    $ChannelOwners = $ownersReponse.userId -join ";"; 
                                    $ChannelMembers = $membersReponse.userId -join ";"; 
                                    $ChannelGuests = $guestsReponse.userId -join ";"; 
                                    
                                }
                            }                           

                            if ($ChannelOwners) {
                                $ChannelOwners+= ";"
                            } 
                            if ($ChannelMembers) {
                                $ChannelMembers+= ";"
                            }
                            if ($ChannelGuests) {
                                $ChannelGuests+= ";"
                            }
                                      
                            $null = $ChannelList.Add([PSCustomObject]@{
                                    GroupID             = $Id
                                    Id                  = $channelId
                                    DisplayName         = $channelDisplayName
                                    Description         = $itemChannel.description
                                    IsFavoriteByDefault = $itemChannel.isFavoriteByDefault
                                    Email               = $itemChannel.email
                                    WebUrl              = $itemChannel.webUrl
                                    MembershipType      = $itemChannel.membershipType
                                    ChannelOwners       = $ChannelOwners
                                    ChannelMembers      = $ChannelMembers
                                    ChannelGuests       = $ChannelGuests
                                    OperationStatus     = ""; 
                                    Operation           = ""; 
                                    AdditionalInfo      = ""
                                })
                            $j++
                            LogWrite -Message  "Completed channel!"
                     }
                  }
                  else {
                    LogWrite -Level ERROR "-An issue occured when calling [Get-NIHTeamChannel] - $TeamName : $c"
                  }
               }                    
            }
            $i++
            LogWrite -Message  "Completed team!"                
        }}
    } 
    LogWrite -Message "Total Graph API call: $calls"   
    return $ChannelList
} 

Function ParseO365Groups {
    [cmdletBinding()]
    param(
        [parameter(Mandatory = $false)]
        [hashtable]$AuthToken,
        [parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        $Groups

    )   
    [System.Collections.ArrayList]$GroupList = @()
    $Groups = @($Groups)
    if ($Groups -and $Groups.Count -gt 0) {
        $Groups | & { process {
                $Id = $_.id;                
                #$GroupOwners, $GroupMembers = $null;                   
                #Get group owners/members
                $script:authToken = Connect-GraphAPIWithCert -TenantId $script:TenantName -AppId $script:appIdAdminPortalOperation -Thumbprint $script:appThumbprintAdminPortalOperation

                if ($Id) {
                    if ($script:authToken) {
                        $extendProps = Get-NIHO365Group -AuthToken $script:authToken -Id $Id -Select hideFromAddressLists, hideFromOutlookClients
                        $ownersReponse = Get-NIHO365GroupOwners -AuthToken $script:authToken -Id $Id 
                        $membersReponse = Get-NIHO365GroupMembers -AuthToken $script:authToken -Id $Id
                        $guestsReponse = $membersReponse | Where-Object { $_.userPrincipalName -like '*#EXT#*' }   
                    }
                }
                $GroupOwners = $ownersReponse.userPrincipalName;
                $GroupMembers = $membersReponse.userPrincipalName;
                $GroupGuests = $guestsReponse.userPrincipalName;
                if ($ownersReponse.Count -gt 0) {
                    $GroupOwners = $null
                    $ownersReponse.ForEach( {
                            $GroupOwners += $_.userPrincipalName + "; ";
                        });
                }
                if ($membersReponse.Count -gt 0) {
                    $GroupMembers = $null;
                    $membersReponse.ForEach( {
                            $GroupMembers += $_.userPrincipalName + "; ";
                        });
                }
                if ($guestsReponse.Count -gt 0) {
                    $GroupGuests = $null;
                    $guestsReponse.ForEach( {
                            $GroupGuests += $_.userPrincipalName + "; ";
                        });
                }
                #End get owners/members
                $null = $GroupList.Add([PSCustomObject][ordered]@{
                        GroupID                      = $Id;
                        GroupOwners                  = $GroupOwners; #if ($GroupOwners) { $GroupOwners } else { $null };
                        GroupMembers                 = $GroupMembers; #if ($GroupMembers) { $GroupMembers } else { $null };
                        GroupGuests                  = $GroupGuests;
                        OnPremisesSyncEnabled        = $_.onPremisesSyncEnabled;
                        DisplayName                  = $_.displayName;
                        Description                  = $_.description;
                        Classification               = $_.classification;
                        CreatedDateTime              = $_.createdDateTime;
                        DeletedDateTime              = $_.deletedDateTime;
                        CreationOptions              = $_.creationOptions -join ',';
                        IsAssignableToRole           = $_.isAssignableToRole;
                        Mail                         = $_.mail ;
                        MailNickname                 = $_.mailNickname;
                        MailEnabled                  = $_.mailEnabled;
                        ProxyAddresses               = $_.proxyAddresses -join ',';
                        RenewedDateTime              = $_.renewedDateTime;
                        ResourceBehaviorOptions      = $_.resourceBehaviorOptions -join ',';
                        ResourceProvisioningOptions  = $_.resourceProvisioningOptions -join ',';
                        SecurityEnabled              = $_.securityEnabled;
                        SecurityIdentifier           = $_.securityIdentifier;
                        Visibility                   = $_.visibility;  
                        OnPremisesLastSyncDateTime   = $_.OnPremisesLastSyncDateTime;                        
                        OnPremisesSamAccountName     = $_.OnPremisesSamAccountName;
                        OnPremisesSecurityIdentifier = $_.OnPremisesSecurityIdentifier;                      
                        AllowExternalSenders         = $_.AllowExternalSenders;
                        AutoSubscribeNewMembers      = $_.AutoSubscribeNewMembers;
                        HideFromAddressLists         = $extendProps.HideFromAddressLists;
                        HideFromOutlookClients       = $extendProps.HideFromOutlookClients;
                        OperationStatus              = ""; 
                        Operation                    = ""; 
                        AdditionalInfo               = ""                  
                    });            
            }}
    }
    return $GroupList 
}

#region Provisioning
Function ParseO365Group {
    param($Group,$GroupOwners,$GroupMembers,$HideFromAddressLists,$HideFromOutlookClients)
    if ($Group) {
        return [PSCustomObject]@{
            GroupID                      = $Group.Id;
            GroupOwners                  = $GroupOwners; #if ($GroupOwners) { $GroupOwners } else { $null };
            GroupMembers                 = $GroupMembers; #if ($GroupMembers) { $GroupMembers } else { $null };
            GroupGuests                  = ""
            OnPremisesSyncEnabled        = $Group.onPremisesSyncEnabled;
            DisplayName                  = $Group.displayName;
            Description                  = $Group.description;
            Classification               = $Group.classification;
            CreatedDateTime              = $Group.createdDateTime;
            DeletedDateTime              = $Group.deletedDateTime;
            CreationOptions              = $Group.creationOptions -join ',';
            IsAssignableToRole           = $Group.isAssignableToRole;
            Mail                         = $Group.mail ;
            MailNickname                 = $Group.mailNickname;
            MailEnabled                  = $Group.mailEnabled;
            ProxyAddresses               = $Group.proxyAddresses -join ',';
            RenewedDateTime              = $Group.renewedDateTime;
            ResourceBehaviorOptions      = $Group.resourceBehaviorOptions -join ',';
            ResourceProvisioningOptions  = $Group.resourceProvisioningOptions -join ',';
            SecurityEnabled              = $Group.securityEnabled;
            SecurityIdentifier           = $Group.securityIdentifier;
            Visibility                   = $Group.visibility;  
            OnPremisesLastSyncDateTime   = $Group.OnPremisesLastSyncDateTime;                        
            OnPremisesSamAccountName     = $Group.OnPremisesSamAccountName;
            OnPremisesSecurityIdentifier = $Group.OnPremisesSecurityIdentifier;                      
            AllowExternalSenders         = $Group.AllowExternalSenders;
            AutoSubscribeNewMembers      = $Group.AutoSubscribeNewMembers;
            #HideFromAddressLists         = $Group.HideFromAddressLists;
            #HideFromOutlookClients       = $Group.HideFromOutlookClients;
            HideFromAddressLists         = $HideFromAddressLists;
            HideFromOutlookClients       = $HideFromOutlookClients;
            OperationStatus              = ""; 
            Operation                    = ""; 
            AdditionalInfo               = "" 
        }
    }
}

Function ParseO365Team {
    param($Team)
    if ($Team) {
        return [PSCustomObject]@{
            # group object
            GroupID                                 = $Team.id                                       
            DisplayName                             = $Team.displayName
            Description                             = $Team.description
            InternalId                              = $Team.internalId
            Classification                          = $Team.classification
            CreatedDateTime                         = $Team.createdDateTime
            Visibility                              = $Team.visibility
            WebUrl                                  = $Team.webUrl
            IsArchived                              = $Team.isArchived
            OperationStatus                         = ""; 
            Operation                               = ""; 
            AdditionalInfo                          = ""
        }
    }
}
#endregion
