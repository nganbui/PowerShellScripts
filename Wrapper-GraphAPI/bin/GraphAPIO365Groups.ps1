#region Group
function Get-NIHO365Groups {
    [cmdletBinding()]
    param(
        [parameter(Mandatory = $true)]
        [hashtable]$AuthToken,                 
        [ValidateSet('beta', 'v1.0')] #Graph API version
        [string]$ApiVersion = 'v1.0'
    )    
    begin {        
        # Create header
        $Header = @{       
            Authorization = $AuthToken['Authorization'] 
        }
    }
    process {
        Write-progress -Activity "Finding Groups"
        $objectCollection = @()
        $Uri = "https://graph.microsoft.com/$($ApiVersion)/groups?`$filter=groupTypes/any(g:g eq 'Unified')"
        $Results = Invoke-NIHGraph -Method "Get" -URI $Uri -Headers $Header 
        if ($Results.value) {
            $objectCollection = $Results.value
            $NextLink = $Results.'@odata.nextLink'
            while ($null -ne $NextLink) {        
                $Results = (Invoke-NIHGraph -Method "Get" -Uri $NextLink -Headers $Header)
                $NextLink = $Results.'@odata.nextLink'
                $objectCollection += $Results.value
            }     
        } 
        else {
            $objectCollection = $Results
        }
        Write-progress -Activity "Finding Groups" -Completed
        return $objectCollection
    }
}
function Get-NIHDeletedO365Groups {
    [cmdletBinding()]
    param(
        [parameter(Mandatory = $true)]
        [hashtable]$AuthToken,                 
        [ValidateSet('beta', 'v1.0')] #Graph API version
        [string]$ApiVersion = 'v1.0'
    )    
    begin {        
        # Create header
        $Header = @{       
            Authorization = $AuthToken['Authorization'] 
        }
    }
    process {
        Write-progress -Activity "Finding Deleted Groups"
        $objectCollection = @()
        $Uri = "https://graph.microsoft.com/$($ApiVersion)/directory/deletedItems/microsoft.graph.group?`$filter=groupTypes/any(g:g eq 'Unified')"
        $Results = Invoke-NIHGraph -Method "Get" -URI $Uri -Headers $Header 
        if ($Results.value) {
            $objectCollection = $Results.value
            $NextLink = $Results.'@odata.nextLink'
            while ($null -ne $NextLink) {        
                $Results = (Invoke-NIHGraph -Method "Get" -Uri $NextLink -Headers $Header)
                $NextLink = $Results.'@odata.nextLink'
                $objectCollection += $Results.value
            }     
        } 
        else {
            $objectCollection = $Results
        }
        Write-progress -Activity "Finding Deleted Groups" -Completed
        return $objectCollection

    }
}
function Get-NIHO365Group {
    [cmdletBinding()]
    param(
        [parameter(Mandatory = $true)]
        [hashtable]$AuthToken,                 
        [ValidateSet('beta', 'v1.0')] #Graph API version
        [string]$ApiVersion = 'v1.0',
        [parameter(Mandatory = $true)]
        [string]$Id,        
        [parameter(Mandatory = $false, parameterSetName = "Select")]
        [ValidateSet  ("hideFromOutlookClients", "hideFromAddressLists")]
        [String[]]$Select        
    )    
    begin {        
        # Create header
        $Header = @{       
            Authorization = $AuthToken['Authorization'] 
        }
    }
    process {
        Write-progress -Activity "Getting a group information"
        $Uri = "https://graph.microsoft.com/$($ApiVersion)/groups/$Id"
        if ($Select) { 
            $uri = $uri + '?$select=' + ($Select -join ",") 
        }
        try {
            $results = Invoke-NIHGraph -Method "Get" -URI $Uri -Headers $Header
            #ErrorInvalidGroup
        }
        catch {
            if ($_.exception.response.statuscode.value__ -eq 404) {
                Write-Progress -Activity 'Getting a group information' -Completed
                Write-Warning -Message "Not found error while getting data for group '$Id'" ; return
            }
            else { throw $_ ; return }
        }
        Write-Progress -Activity 'Getting a group information' -Completed
        $results
    }
}
function Get-NIHO365GroupOwners {
    [cmdletBinding()]
    param(
        [parameter(Mandatory = $true)]
        [hashtable]$AuthToken, 
        [parameter(Mandatory = $true)]
        [string]$Id,                 
        [ValidateSet('beta', 'v1.0')] #Graph API version
        [string]$ApiVersion = 'v1.0'
    )
    begin {        
        # Create header
        $Header = @{       
            Authorization = $AuthToken['Authorization']
        }   
    }
    process {
        $objectCollection = @()
        $Uri = "https://graph.microsoft.com/$($ApiVersion)/groups/$Id/owners?`$top=999"
        $Results = Invoke-NIHGraph -Method "Get" -URI $Uri -Headers $Header 
        if ($Results.value) {
            $objectCollection = $Results.value
            $NextLink = $Results.'@odata.nextLink'
            while ($null -ne $NextLink) {        
                $Results = (Invoke-NIHGraph -Method "Get" -Uri $NextLink -Headers $Header)
                $NextLink = $Results.'@odata.nextLink'
                $objectCollection += $Results.value
            }     
        } 
        else {
            $objectCollection = $Results
        }
        return $objectCollection
    } 
}
function Get-NIHO365GroupMembers {
    [cmdletBinding()]
    param(
        [parameter(Mandatory = $true)]
        [hashtable]$AuthToken, 
        [parameter(Mandatory = $true)]
        [string]$Id,                 
        [ValidateSet('beta', 'v1.0')] #Graph API version
        [string]$ApiVersion = 'v1.0'
    )
    begin {        
        # Create header
        $Header = @{       
            Authorization = $AuthToken['Authorization']
        }   
    }
    process {
        $objectCollection = @()
        $Uri = "https://graph.microsoft.com/$($ApiVersion)/groups/$Id/members?`$top=999"
        $Results = Invoke-NIHGraph -Method "Get" -URI $Uri -Headers $Header 
        if ($Results.value) {
            $objectCollection = $Results.value
            $NextLink = $Results.'@odata.nextLink'
            while ($null -ne $NextLink) {        
                $Results = (Invoke-NIHGraph -Method "Get" -Uri $NextLink -Headers $Header)
                $NextLink = $Results.'@odata.nextLink'
                $objectCollection += $Results.value
            }     
        } 
        else {
            $objectCollection = $Results
        }
        return $objectCollection
    } 
}
function New-NIHO365Group {
    [Cmdletbinding(SupportsShouldprocess = $true)]    
    param(
        [parameter(Mandatory = $true)]
        [hashtable]$AuthToken,
        [ValidateSet('beta', 'v1.0')] #Graph API version
        [string]$ApiVersion = 'v1.0', 
        #The Name of the group / team
        [Parameter(Mandatory = $true)]
        [string]$Name,
        #The group/team's mail nickname
        [string]$MailNickName,
        #A description for the group
        [string]$Description,
        #The visibility of the group, private by default
        [ValidateSet('private', 'public')]
        [string]$Visibility = 'private',
        #User Principal Name or ID or as objects
        $Members,
        #User Principal Name or ID or as objects
        $Owners,       
        #if specified group will be added without prompting
        [Switch]$Force

    )    
    # Create header
    $Header = @{       
        Authorization = $AuthToken['Authorization'] 
    } 
    #Check if a group exists then return
    if ( (Invoke-RestMethod -Method Get -Headers $Header -Uri "https://graph.microsoft.com/$($ApiVersion)/groups?`$filter=displayname eq '$Name'" ).value) {
        throw "There is already a group with the display name '$Name'." ;
        return
    }           
    if (-not $MailNickName) { $MailNickName = $Name -replace "\W", '' }
    $settings = @{  
        'displayName'             = $Name ;
        'description'             = $Description
        'mailNickname'            = $MailNickName;
        'mailEnabled'             = $true;
        'securityEnabled'         = $false;
        'visibility'              = $Visibility.ToLower() ;
        'groupTypes'              = @("Unified") ;  
        'resourceBehaviorOptions' = @('WelcomeEmailEnabled')
    }
    
    #if we got owners or users with no ID, fix them at the end, if they have an ID add them now
    if ($Members) {
        $settings['members@odata.bind'] = @();
        foreach ($m in $Members) {
            if ($m.id) { $settings['members@odata.bind'] += "https://graph.microsoft.com/$($ApiVersion)/users/$($m.id)" }
            else { $settings['members@odata.bind'] += "https://graph.microsoft.com/$($ApiVersion)/users/$m" }
        }
    }
    if ($Owners) {
        $settings['owners@odata.bind'] = @()
        foreach ($o in $Owners) {
            if ($o.id) { $settings['owners@odata.bind'] += "https://graph.microsoft.com/$($ApiVersion)/users/$($o.id)" }
            else { $settings['owners@odata.bind'] += "https://graph.microsoft.com/$($ApiVersion)/users/$o" }
        }
    }    
    $Body = ConvertTo-Json $settings 
    $Uri = "https://graph.microsoft.com/$($ApiVersion)/groups"
      
    if ($Force -or $PSCmdlet.shouldprocess($Name, "Add new Group")) {
        Write-Progress -Activity 'Creating Group/Team' -CurrentOperation "Adding Group $Name"
        $group = Invoke-NIHGraph -Method "POST" -URI $Uri -Headers $Header -Body $Body             
        if ($group.id) {                      
            foreach ($m in $group.members) { if ($m.'@odata.type' -match "user") { $m.pstypenames.add("GraphUser") } }
            $group.pstypenames.Add("GraphGroup")            
        }             
        Write-Progress -Activity 'Creating Group/Team' -Completed                    
        $group
    }
}
function Get-NIHO365GroupByName {
    [cmdletBinding()]
    param(
        [parameter(Mandatory = $true)]
        [hashtable]$AuthToken, 
        [parameter(Mandatory = $true)]
        [string]$DisplayName,                 
        [ValidateSet('beta', 'v1.0')] #Graph API version
        [string]$ApiVersion = 'v1.0'
    )  
    # Create header
    $Header = @{       
        Authorization = $AuthToken['Authorization']
    } 
    $Uri = "https://graph.microsoft.com/$($ApiVersion)/groups?`$filter=displayName eq '$DisplayName'"
    $Results = Invoke-NIHGraph -Method "Get" -URI $Uri -Headers $Header 
    if ($Results.value) {
        $Results = ($Results.value)[0]
    } 
    $Results
}
function Add-NIHO365GroupMember {
    <#
      .Synopsis
        Adds a user to a group/team as either a member or owner.      
      .Example
        >Add-O365GroupMember -Group $GroupID -Member member@xyz.sharepoint.com
        >Add-O365GroupMember -Group $GroupID -Member admin@xyz.sharepoint.com -AsOwner        
    #>
    [Cmdletbinding(SupportsShouldprocess = $true)]    
    param   (
        [parameter(Mandatory = $true)]
        [hashtable]$AuthToken,
        #The group / team either as an ID or a group/team object with an IDn
        [Parameter(Mandatory = $true)]      
        $Group,
        #The user or nested-group to add, either as a UPN or ID or as a object with an ID
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        $Member,
        #If specified the user will be added as an owner, otherwise they will be a standard member
        [switch]$AsOwner,
        #If specified group member will be added without prompting
        [Switch]$Force
    )
    begin {        
        # Create header
        $Header = @{       
            Authorization = $AuthToken['Authorization'] 
        } 
        if ($Group.ID) { $groupID = $Group.ID }
        elseif ($Group -is [String]) { $groupID = $Group }
        else { Write-Warning -Message 'Could not process Group parameter.'; return }        
        #if ($AsOwner) { $uri = "https://graph.microsoft.com/v1.0/groups/$groupID/owners/`$ref" }
        #else { $uri = "https://graph.microsoft.com/v1.0/groups/$groupID/members/`$ref" }
    }
    process {       
        #Get User ID
        if ($Member.id) { $memberID = $Member.id }
        else {
            try {
                # Get user object by calling Get-O365User -UserID pass ID or UPN                
                $Member = Get-NIHO365User -AuthToken  $AuthToken -UserID $Member
                $memberID = $Member.id
            }
            catch { throw "Could not get a user matching $Member"; return }
            if (-not $Member) { throw "Could not get a member ID"; return }
        }
        #Check member exist
        $listofMembers = Get-NIHO365GroupMembers -AuthToken $AuthToken -Id $Group        
        $listofMembers = $listofMembers | % { $_.id }
        $listofOwners = Get-NIHO365GroupOwners -AuthToken $AuthToken -Id $Group
        $listofOwners = $listofOwners | % { $_.id }
        # check if this action not add a member as owner and member is already in Member list => do nothing-skip
        if (-not $AsOwner -and $memberID -in $listofMembers) { return }

        $settings = @{'@odata.id' = "https://graph.microsoft.com/v1.0/directoryObjects/$memberID" }        
        $Body = ConvertTo-Json $settings
        Write-Debug $Body
        if ($Force -or $PSCmdlet.shouldprocess($Member.displayname, "Add to Group")) {
            if ($memberID -notin $listofMembers) {
                $uri = "https://graph.microsoft.com/v1.0/groups/$groupID/members/`$ref"
                Invoke-NIHGraph -Method "POST" -URI $Uri -Headers $Header -Body $Body
            }
            if ($AsOwner -and $memberID -notin $listofOwners) {
                $uri = "https://graph.microsoft.com/v1.0/groups/$groupID/owners/`$ref"                
                Invoke-NIHGraph -Method "POST" -URI $Uri -Headers $Header -Body $Body
            }
        }
    }
}
function Remove-NIHO365GroupMember {
    <#
      .Synopsis
        Adds a user to a group/team as either a member or owner.      
      .Example
        >Remove-NIHO365GroupMember -Group $GroupID -Member member@xyz.sharepoint.com
        >Remove-NIHO365GroupMember -Group $GroupID -Member admin@xyz.sharepoint.com -AsOwner        
    #>
    [Cmdletbinding(SupportsShouldprocess = $true)]    
    param   (
        [parameter(Mandatory = $true)]
        [hashtable]$AuthToken,
        #The group / team either as an ID or a group/team object with an IDn
        [Parameter(Mandatory = $true)]      
        $Group,
        #The user or nested-group to add, either as a UPN or ID or as a object with an ID
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        $Member,
        #If specified the user will be added as an owner, otherwise they will be a standard member
        [switch]$AsOwner,
        #If specified group member will be added without prompting
        [Switch]$Force
    )
    begin {        
        # Create header
        $Header = @{       
            Authorization = $AuthToken['Authorization'] 
        } 
        if ($Group.ID) { $groupID = $Group.ID }
        elseif ($Group -is [String]) { $groupID = $Group }
        else { Write-Warning -Message 'Could not process Group parameter.'; return }        
        #if ($AsOwner) { $uri = "https://graph.microsoft.com/v1.0/groups/$groupID/owners/`$ref" }
        #else { $uri = "https://graph.microsoft.com/v1.0/groups/$groupID/members/`$ref" }
    }
    process {       
        #Get User ID
        if ($Member.id) { $memberID = $Member.id }
        else {
            try {
                # Get user object by calling Get-O365User -UserID pass ID or UPN                
                $Member = Get-NIHO365User -AuthToken  $AuthToken -UserID $Member
                $memberID = $Member.id
            }
            catch { throw "Could not get a user matching $Member"; return }
            if (-not $Member) { throw "Could not get a member ID"; return }
        }
        #Check member exist
        $listofMembers = Get-NIHO365GroupMembers -AuthToken $AuthToken -Id $Group        
        $listofMembers = $listofMembers | % { $_.id }
        $listofOwners = Get-NIHO365GroupOwners -AuthToken $AuthToken -Id $Group
        $listofOwners = $listofOwners | % { $_.id }
        # check if this action not add a member as owner and member is already in Member list => do nothing-skip
        #if (-not $AsOwner -and $memberID -in $listofMembers) { return }

        $settings = @{'@odata.id' = "https://graph.microsoft.com/v1.0/directoryObjects/$memberID" }        
        $Body = ConvertTo-Json $settings
        Write-Debug $Body
        if ($Force -or $PSCmdlet.shouldprocess($Member.displayname, "Remove user from Group")) {
            if ($memberID -in $listofMembers) {
                $uri = "https://graph.microsoft.com/v1.0/groups/$groupID/members/$memberID/`$ref"
                #/groups/{id}/members/{id}/$ref
                Invoke-NIHGraph -Method "DELETE" -URI $Uri -Headers $Header -Body $Body
            }
            if ($AsOwner -and $memberID -in $listofOwners) {
                $uri = "https://graph.microsoft.com/v1.0/groups/$groupID/owners/$memberID/`$ref"
                #/groups/{id}/owners/{id}/$ref                
                Invoke-NIHGraph -Method "DELETE" -URI $Uri -Headers $Header -Body $Body
            }
        }
    }
}
function Remove-NIHSoftDeletedGroup{
    [cmdletBinding()]
    param(
        [parameter(Mandatory = $true)]
        [hashtable]$AuthToken,                 
        [ValidateSet('beta', 'v1.0')] #Graph API version
        [string]$ApiVersion = 'v1.0',
        [parameter(Mandatory = $true)]
        [string]$Id
    )    
    begin {        
        # Create header
        $Header = @{       
            Authorization = $AuthToken['Authorization'] 
        }
    }
    process {       

        Write-progress -Activity "Permanently delete item"       
        $Uri = "https://graph.microsoft.com/$($ApiVersion)/directory/deletedItems/$($Id)"
        Invoke-NIHGraph -Method "Delete" -URI $Uri -Headers $Header         
        Write-progress -Activity "Permanently delete item" -Completed
       
    }
}
function Update-NIHGroupSettings {
    [cmdletBinding()]
    param(
        [parameter(Mandatory = $true)]
        [hashtable]$AuthToken, 
        [ValidateSet('beta', 'v1.0')] #Graph API version
        [string]$ApiVersion = 'v1.0',         
        [parameter(Mandatory = $true)]
        [string]$Id,
        [parameter(Mandatory = $false)]        
        [Nullable[boolean]] $HideFromOutlookClients,                
        [parameter(Mandatory = $false)]
        [Nullable[boolean]] $HideFromAddressLists
    )

    # Create header
    $Header = @{       
        Authorization = $AuthToken['Authorization']
    }
    $settings = @{}
    if(-not [string]::IsNullOrEmpty($HideFromOutlookClients) ){
        $settings['hideFromOutlookClients'] = $HideFromOutlookClients 
    }
    if(-not [string]::IsNullOrEmpty($HideFromAddressLists) ){
        $settings['hideFromAddressLists'] = $HideFromAddressLists  
        if ($true -eq $HideFromAddressLists){
            $settings['hideFromOutlookClients'] = $true
        }
    }    
    
    $Body = ConvertTo-Json $settings 
    $Uri = "https://graph.microsoft.com/$($ApiVersion)/groups/$Id"    
    Invoke-NIHGraph -Method "PATCH" -URI $Uri -Headers $Header -Body $Body
}
#endregion

#region Team
function New-NIHTeam {
    [Cmdletbinding(SupportsShouldprocess = $true)]    
    param(
        [parameter(Mandatory = $true)]
        [hashtable]$AuthToken,
        [ValidateSet('beta', 'v1.0')] #Graph API version
        [string]$ApiVersion = 'v1.0', 
        #The Name of the group / team
        [Parameter(Mandatory = $true)]
        [string]$Name,
        #The group/team's mail nickname
        [string]$MailNickName,
        #A description for the group
        [string]$Description,
        #The visibility of the group, private by default
        [ValidateSet('private', 'public')]
        [string]$Visibility = 'private',
        #User Principal Name or ID or as objects
        $Members,
        #User Principal Name or ID or as objects
        $Owners,
        #if specified group will be added without prompting
        [Switch]$Force
    )    
    # Create header
    $Header = @{       
        Authorization = $AuthToken['Authorization'] 
    } 
    #Check if a group exists then return
    if ( (Invoke-RestMethod -Method Get -Headers $Header -Uri "https://graph.microsoft.com/$($ApiVersion)/groups?`$filter=displayname eq '$Name'" ).value) {
        throw "There is already a group with the display name '$Name'." ;
        return
    }           
    if (-not $MailNickName) { $MailNickName = $Name -replace "\W", '' }
    #HideGroupInOutlook,WelcomeEmailDisabled,SubscribeMembersToCalendarEventsDisabled
    $settings = @{  
        'displayName'             = $Name ;
        'description'             = $Description
        'mailNickname'            = $MailNickName;
        'mailEnabled'             = $true;
        'securityEnabled'         = $false;
        'visibility'              = $Visibility.ToLower() ;
        'groupTypes'              = @("Unified") ;  
        'resourceBehaviorOptions' = @('WelcomeEmailDisabled', "HideGroupInOutlook", 'SubscribeMembersToCalendarEventsDisabled')
    }  
    if ($Members) {
        $settings['members@odata.bind'] = @();
        foreach ($m in $Members) {
            if ($m.id) { $settings['members@odata.bind'] += "https://graph.microsoft.com/$($ApiVersion)/users/$($m.id)" }
            else { $settings['members@odata.bind'] += "https://graph.microsoft.com/$($ApiVersion)/users/$m" }
        }
    }
    if ($Owners) {
        $settings['owners@odata.bind'] = @()
        foreach ($o in $Owners) {
            if ($o.id) { $settings['owners@odata.bind'] += "https://graph.microsoft.com/$($ApiVersion)/users/$($o.id)" }
            else { $settings['owners@odata.bind'] += "https://graph.microsoft.com/$($ApiVersion)/users/$o" }
        }
    }    
    $Body = ConvertTo-Json $settings 
    $Uri = "https://graph.microsoft.com/$($ApiVersion)/groups"
      
    if ($Force -or $PSCmdlet.shouldprocess($Name, "Add new Group")) {
        Write-Progress -Activity 'Creating Group/Team' -CurrentOperation "Adding Group $Name"
        $group = Invoke-NIHGraph -Method "POST" -URI $Uri -Headers $Header -Body $Body
        $groupId = $group.id
        #Start-Sleep -Seconds 30 
        $verifyGroup = Get-NIHO365Group -AuthToken $AuthToken -ApiVersion $ApiVersion -Id $groupId
        while ($verifyGroup -eq $null) {
            Sleep 30;
        }
        if ($groupId) {               
            foreach ($m in $group.members) { if ($m.'@odata.type' -match "user") { $m.pstypenames.add("GraphUser") } }
            #promote to team after provision group
            $uri = $uri + "/" + $group.id + "/team"
            Write-Progress -Activity 'Creating Group/Team' -CurrentOperation "Team-enabling Group $Name"            
            $teamsettings = @{ 
                <#"discoverySettings" = @{
                    "showInTeamsSearchAndSuggestions" = $true;
                }#>
                "memberSettings"    = @{
                    "allowCreateUpdateChannels"         = $false;
                    "allowDeleteChannels"               = $false;
                    "allowAddRemoveApps"                = $false;
                    "allowCreateUpdateRemoveTabs"       = $false;
                    "allowCreateUpdateRemoveConnectors" = $false;
                }
                "guestSettings"     = @{
                    "allowCreateUpdateChannels" = $false;
                    "allowDeleteChannels"       = $false;
                }
                "messagingSettings" = @{
                    "allowUserEditMessages"    = $true;
                    "allowUserDeleteMessages"  = $true;
                    "allowOwnerDeleteMessages" = $true;
                    "allowTeamMentions"        = $true;
                    "allowChannelMentions"     = $true;
                }
                "funSettings"       = @{
                    "allowGiphy"            = $true;
                    "giphyContentRating"    = "moderate";
                    "allowStickersAndMemes" = $true;
                    "allowCustomMemes"      = $true;
                }                   
            }            
            $teamsettings = ConvertTo-Json $teamsettings
            $team = Invoke-NIHGraph -Method "PUT" -URI $Uri -Headers $Header -Body $teamsettings
            $team.pstypenames.Add("GraphTeam")
            #Start-Sleep -Seconds 30             
            Write-Progress -Activity 'Creating Group/Team' -Completed
            $team
        }
    }
}
function Update-NIHTeamSettings {
    [cmdletBinding()]
    param(
        [parameter(Mandatory = $true)]
        [hashtable]$AuthToken, 
        [ValidateSet('beta', 'v1.0')] #Graph API version
        [string]$ApiVersion = 'v1.0',         
        [parameter(Mandatory = $true)]
        [string]$Id
    )

    # Create header
    $Header = @{       
        Authorization = $AuthToken['Authorization']
    }      
    $teamsettings = @{ 
        <#"discoverySettings" = @{
            "showInTeamsSearchAndSuggestions" = $true;
        }#>
        "memberSettings"    = @{
            "allowCreateUpdateChannels"         = $false;            
            "allowCreatePrivateChannels"        = $false;
            "allowDeleteChannels"               = $false;
            "allowAddRemoveApps"                = $false;            
            "allowCreateUpdateRemoveTabs"       = $false;
            "allowCreateUpdateRemoveConnectors" = $false;                       
        }
        "guestSettings"     = @{
            "allowCreateUpdateChannels" = $false;
            "allowDeleteChannels"       = $false;
        }
        "messagingSettings" = @{
            "allowUserEditMessages"    = $true;
            "allowUserDeleteMessages"  = $true;
            "allowOwnerDeleteMessages" = $true;
            "allowTeamMentions"        = $true;
            "allowChannelMentions"     = $true;
        }
        "funSettings"       = @{
            "allowGiphy"            = $true;
            "giphyContentRating"    = "moderate";
            "allowStickersAndMemes" = $true;
            "allowCustomMemes"      = $true;
        }                   
    }
    
    $Body = ConvertTo-Json $teamsettings    
    $Uri = "https://graph.microsoft.com/$($ApiVersion)/teams/$Id"    
    Invoke-NIHGraph -Method "PATCH" -URI $Uri -Headers $Header -Body $Body
}
function Update-NIHTeamSettingsPostProvision {
    [cmdletBinding()]
    param(
        [parameter(Mandatory = $true)]
        [hashtable]$AuthToken, 
        [ValidateSet('beta', 'v1.0')] #Graph API version
        [string]$ApiVersion = 'v1.0',         
        [parameter(Mandatory = $true)]
        [string]$Id
    )

    # Create header
    $Header = @{       
        Authorization = $AuthToken['Authorization']
    }      
    $settings = @{
        'hideFromOutlookClients' = $false 
        'hideFromAddressLists'   = $true
        #'isSubscribedByMail' = $true    
    }
    $Body = ConvertTo-Json $settings 
    $Uri = "https://graph.microsoft.com/$($ApiVersion)/groups/$Id"    
    Invoke-NIHGraph -Method "PATCH" -URI $Uri -Headers $Header -Body $Body
}
function Get-NIHTeam {
    [cmdletBinding()]
    param(
        [parameter(Mandatory = $true)]
        [hashtable]$AuthToken, 
        [parameter(Mandatory = $true)]
        [string]$Id,                 
        [ValidateSet('beta', 'v1.0')] #Graph API version
        [string]$ApiVersion = 'v1.0'
    ) 
    begin {        
        # Create header
        $Header = @{       
            Authorization = $AuthToken['Authorization']
        }   
    }
    process {       
        $Uri = "https://graph.microsoft.com/$($ApiVersion)/teams/$Id"
        $Results = Invoke-NIHGraph -Method "Get" -URI $Uri -Headers $Header 

        if ($Results) {
            $Results  
        } 
        else {
            $null
        } 
    }
}
function Get-NIHTeamMembers {
    [cmdletBinding()]
    param(
        [parameter(Mandatory = $true)]
        [hashtable]$AuthToken, 
        [parameter(Mandatory = $true)]
        [string]$Id,                 
        [ValidateSet('beta', 'v1.0')] #Graph API version
        [string]$ApiVersion = 'v1.0',
        [ValidateSet('All','Owner', 'Member', 'Guest')] 
        [string]$Role = 'All'
    ) 
    begin {        
        # Create header
        $Header = @{       
            Authorization = $AuthToken['Authorization']
        }   
    }
    process {
        $objectCollection = @()        
        $Uri = "https://graph.microsoft.com/$($ApiVersion)/teams/$Id/members?`$top=999"
        $Results = Invoke-NIHGraph -Method "Get" -URI $Uri -Headers $Header 
        if ($Results.value) {
            $objectCollection = $Results.value
            $NextLink = $Results.'@odata.nextLink'
            while ($null -ne $NextLink) {        
                $Results = (Invoke-NIHGraph -Method "Get" -Uri $NextLink -Headers $Header)
                $NextLink = $Results.'@odata.nextLink'
                $objectCollection += $Results.value
            }     
        } 
        else {
            $objectCollection = $Results
        }
        
        switch($Role){
            'Owner' { 
                $objectCollection = $objectCollection | Where-Object { $_.roles -contains 'owner' }
            }
            'Member' { 
                $objectCollection = $objectCollection | Where-Object {  $_.roles.Count -eq 0 }
            }
            'Guest' { 
                $objectCollection = $objectCollection | Where-Object { $_.roles -contains 'guest' }
            }           
        }
        return $objectCollection
    } 
}
function Get-NIHTeamChannel {
    [cmdletBinding()]
    param(
        [parameter(Mandatory = $true)]
        [hashtable]$AuthToken, 
        [parameter(Mandatory = $true)]
        [string]$Id,        
        [ValidateSet('private', 'standard')]
        [string]$MembershipType,
        [ValidateSet('beta', 'v1.0')] #Graph API version
        [string]$ApiVersion = 'v1.0'
    )

    begin {        
        # Create header
        $Header = @{       
            Authorization = $AuthToken['Authorization'] 
        }
    }
    process {
        Write-progress -Activity "Get NIH Team Channels"
        $objectCollection = @()         
        $Uri = "https://graph.microsoft.com/$($ApiVersion)/teams/$Id/channels"
        if ($MembershipType){
            $Uri = $Uri + "?`$filter=membershipType eq '" + $MembershipType + "'"
        }

        $Results = Invoke-NIHGraph -Method "Get" -URI $Uri -Headers $Header 
        if ($null -ne $Results -and $Results.value) {
            $objectCollection+=$Results.value
            $NextLink = $Results.'@odata.nextLink'
            while ($null -ne $NextLink) {        
                $Results = (Invoke-NIHGraph -Method "Get" -Uri $NextLink -Headers $Header)
                $NextLink = $Results.'@odata.nextLink'
                $objectCollection+=$Results.value
            }     
        } 
        else {
            $objectCollection+=$Results
        }
        Write-progress -Activity "Get NIH Team Channels" -Completed
        return $objectCollection
    }
}
function Get-NIHTeamChannelMembers {
    [cmdletBinding()]
    param(
        [parameter(Mandatory = $true)]
        [hashtable]$AuthToken, 
        [parameter(Mandatory = $true)]
        [string]$GroupId,
        [parameter(Mandatory = $true)]
        [string]$ChannelId,
        [ValidateSet('beta', 'v1.0')] #Graph API version
        [string]$ApiVersion = 'v1.0',
        [ValidateSet('All','Owner', 'Member', 'Guest')] 
        [string]$Role = 'All'
        #[switch]$Owner
    ) 
    begin {        
        # Create header
        $Header = @{       
            Authorization = $AuthToken['Authorization']
        }   
    }
    process {
        $objectCollection = @()
        $Uri = "https://graph.microsoft.com/$($ApiVersion)/teams/$GroupId/channels/$ChannelId/members?`$top=999"
        $Results = Invoke-NIHGraph -Method "Get" -URI $Uri -Headers $Header 
        if ($Results.value) {
            $objectCollection = $Results.value
            $NextLink = $Results.'@odata.nextLink'
            while ($null -ne $NextLink) {        
                $Results = (Invoke-NIHGraph -Method "Get" -Uri $NextLink -Headers $Header)
                $NextLink = $Results.'@odata.nextLink'
                $objectCollection += $Results.value
            }     
        } 
        else {
            $objectCollection = $Results
        }
        
        switch($Role){
            'Owner' { 
                $objectCollection = $objectCollection | Where-Object { $_.roles -contains 'owner' }
            }
            'Member' { 
                $objectCollection = $objectCollection | Where-Object {  $_.roles.Count -eq 0 }
            }
            'Guest' { 
                $objectCollection = $objectCollection | Where-Object { $_.roles -contains 'guest' }
            }
            
        }
        return $objectCollection
    }  
}
#endregion

function Get-NIHO365Object {
    [cmdletBinding()]
    param(
        [parameter(Mandatory = $true)]
        [hashtable]$AuthToken,                 
        [ValidateSet('beta', 'v1.0')] #Graph API version
        [string]$ApiVersion = 'v1.0',
        [parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [String]$EndPoint       
    )   
    $objectCollection = @()   

    if ($AuthToken['Authorization']) {
        # Create header
        $Header = @{       
            Authorization = $AuthToken['Authorization']
        }         
        $Uri = "https://graph.microsoft.com/$($ApiVersion)/$($EndPoint)"

        $Results = Invoke-RestMethod -Method "Get" -URI $Uri -Headers $Header  -Verbose 
        if ($Results.value) {
            $objectCollection = $Results.value
            $NextLink = $Results.'@odata.nextLink'
            while ($null -ne $NextLink) {        
                $Results = (Invoke-RestMethod -Method "Get" -Uri $NextLink -Headers $Header)
                $NextLink = $Results.'@odata.nextLink'
                $objectCollection += $Results.value
            }     
        } 
        else {
            $objectCollection = $Results
        }
        return $objectCollection
    }
    else {
        write-host 'Token issue' -ForegroundColor Red
        exit
    }
}
#region being used in provisioning site
function Convert-NIHO365GroupPSCustomObject {
    [cmdletBinding()]
    param( 
        [parameter(Mandatory = $false)]
        [hashtable]$AuthToken,      
        [parameter(Mandatory = $true)]        
        $Group,        
        [parameter(Mandatory = $false)]  
        $GroupOwners,    
        [parameter(Mandatory = $false)] 
        $GroupMembers     
    )
    begin {        
        if ($null -eq $AuthToken) {
            $AuthToken = Connect-NIHO365GraphV1
        }        
        if ($Group -is [String]) {
            # return group object
            $Group = Get-NIHO365Group -AuthToken $AuthToken -Id $Group            
        }
        if ($Group) { $groupID = $Group.id }
        else { Write-Warning -Message 'Could not process Group parameter.'; return }
    }
    process {
        if ($null -eq $GroupOwners -or $null -eq $GroupMembers) {        
            $ownersReponse = Get-NIHO365GroupMembers -AuthToken $authToken -Id $groupID 
            $membersReponse = Get-NIHO365GroupMembers -AuthToken $authToken -Id $groupID                  
            if ($ownersReponse.Count -gt 0) {
                $ownersReponse.ForEach( {
                        $GroupOwners += $_.userPrincipalName + "; ";
                    });
            }
            if ($membersReponse.Count -gt 0) {
                $membersReponse.ForEach( {
                        $GroupMembers += $_.userPrincipalName + "; ";
                    });
            }  
        }      
        return [PSCustomObject]@{
            GroupID                     = $groupID;
            GroupOwners                 = $GroupOwners;
            GroupMembers                = $GroupMembers;
            OnPremisesSyncEnabled       = $Group.onPremisesSyncEnabled;
            DisplayName                 = $Group.displayName;
            Description                 = $Group.description;
            Classification              = $Group.classification;
            CreatedDateTime             = $Group.createdDateTime;
            DeletedDateTime             = $Group.deletedDateTime;
            CreationOptions             = $Group.creationOptions -join ',';
            IsAssignableToRole          = $Group.isAssignableToRole;
            Mail                        = $Group.mail ;
            MailNickname                = $Group.mailNickname;
            MailEnabled                 = $Group.mailEnabled;
            ProxyAddresses              = $Group.proxyAddresses -join ',';
            RenewedDateTime             = $Group.renewedDateTime;
            ResourceBehaviorOptions     = $Group.resourceBehaviorOptions -join ',';
            ResourceProvisioningOptions = $Group.resourceProvisioningOptions -join ',';
            SecurityEnabled             = $Group.securityEnabled;
            SecurityIdentifier          = $Group.securityIdentifier;
            Visibility                  = $Group.visibility;
            OperationStatus             = ""; 
            Operation                   = ""; 
            AdditionalInfo              = ""  
        }
    }

}

function Convert-NIHTeamPSCustomObject {
    [cmdletBinding()]
    param( 
        [parameter(Mandatory = $false)]
        [hashtable]$AuthToken,      
        [parameter(Mandatory = $true)]        
        $Group
    )
    begin {        
        if ($null -eq $AuthToken) {
            $AuthToken = Connect-NIHO365GraphV1
        }        
        if ($Group -is [String]) {
            # return group object
            $Group = Get-NIHO365Group -AuthToken $AuthToken -Id $Group                     
        }
        if ($Group) { $groupID = $Group.id }
        else { Write-Warning -Message 'Could not process team parameter.'; return }
    }
    process {
        # get team object
        $r = Get-NIHTeam -AuthToken $authToken -Id $groupID -ApiVersion "beta" 
        return [PSCustomObject]@{
            # group object
            GroupID                                 = $groupID                                       
            DisplayName                             = $Group.displayName
            Description                             = $Group.description
            Classification                          = $Group.classification
            CreatedDateTime                         = $Group.createdDateTime
            DeletedDateTime                         = $Group.deletedDateTime                    
            Mail                                    = $Group.mail 
            MailNickname                            = $Group.mailNickname
            MailEnabled                             = $Group.mailEnabled                                                          
            Visibility                              = $Group.visibility                                            
            # team object                                   
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
            Status                                  = "Active";
            OperationStatus                         = ""; 
            Operation                               = ""; 
            AdditionalInfo                          = ""
        }
    }
}
#endregion