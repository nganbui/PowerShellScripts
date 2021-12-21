#region Group
function Get-NIHO365Groups {
    [cmdletBinding()]
    param(
        [parameter(Mandatory = $true)]
        [hashtable]$AuthToken,                 
        [ValidateSet('beta', 'v1.0')] #Graph API version
        [string]$ApiVersion = 'v1.0',
        #specifies which properties of the user object should be returned
        [parameter(Mandatory = $false, parameterSetName = "Select")]
        [ValidateSet  ("id","onPremisesSyncEnabled","displayName","description","classification","createdDateTime","deletedDateTime","creationOptions","isAssignableToRole","mail","mailNickname","mailEnabled","proxyAddresses","renewedDateTime","resourceBehaviorOptions","resourceProvisioningOptions","securityEnabled","securityIdentifier","visibility","OnPremisesLastSyncDateTime","OnPremisesSamAccountName", "OnPremisesSecurityIdentifier","AllowExternalSenders")]
        [String[]]$Select="id,onPremisesSyncEnabled,displayName,description,classification,createdDateTime,deletedDateTime,creationOptions,isAssignableToRole,mail,mailNickname,mailEnabled,proxyAddresses,renewedDateTime,resourceBehaviorOptions,resourceProvisioningOptions,securityEnabled,securityIdentifier,visibility,OnPremisesLastSyncDateTime,OnPremisesSamAccountName,OnPremisesSecurityIdentifier"
        )
    begin {        
        # Create header
        $Header = @{       
            Authorization = $AuthToken['Authorization'] 
        }
        $resource = "https://graph.microsoft.com"        
    }
    process {
        Write-progress -Activity "Finding Groups"
        $objectCollection = @()

        $Uri = "$resource/$ApiVersion/groups?`$filter=groupTypes/any(g:g eq 'Unified')&`$top=999"

        if ($Select) { 
            $uri = $uri + '&$select=' + ($Select -join ",") 
        }

        #$Uri = "https://graph.microsoft.com/$($ApiVersion)/groups?`$filter=groupTypes/any(g:g eq 'Unified')&`$top=999"
        
        
        $retryCount = 5
        $retryAttempts = 0
        $backOffInterval = 2

        while ($retryAttempts -le $retryCount) {
            try {
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
                $retryAttempts = $retryCount + 1
                Write-Progress -Activity 'Getting list of M365 Groups' -Completed
                return $objectCollection
                
            }
            catch {
                if ($retryAttempts -lt $retryCount) {
                    $retryAttempts = $retryAttempts + 1        
                    Write-Host "Retry attempt number: $retryAttempts. Sleeping for $backOffInterval seconds..."
                    Start-Sleep $backOffInterval
                    $backOffInterval = $backOffInterval * 2
                }
                else {
                    Write-Verbose -Message 'Unable to getting M365 Groups' -Verbose 
                }
            }
        }        
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
        $Uri = "https://graph.microsoft.com/$($ApiVersion)/directory/deletedItems/microsoft.graph.group?`$filter=groupTypes/any(g:g eq 'Unified')&`$top=999"
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
        [ValidateSet  ("displayName", "description", "visibility", "hideFromOutlookClients", "hideFromAddressLists")]
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

        $retryCount = 5
        $retryAttempts = 0
        $backOffInterval = 2

        while ($retryAttempts -le $retryCount) {
            try {
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
                $retryAttempts = $retryCount + 1
                Write-Progress -Activity 'Getting NIHO365GroupOwners' -Completed
                return $objectCollection
                
            }
            catch {
                if ($retryAttempts -lt $retryCount) {
                    $retryAttempts = $retryAttempts + 1        
                    Write-Host "Retry attempt number: $retryAttempts. Sleeping for $backOffInterval seconds..."
                    Start-Sleep $backOffInterval
                    $backOffInterval = $backOffInterval * 2
                }
                else {
                    Write-Verbose -Message 'Unable to getting NIHO365GroupOwners' -Verbose 
                }
            }
        }
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
        
        $retryCount = 5
        $retryAttempts = 0
        $backOffInterval = 2

        while ($retryAttempts -le $retryCount) {
            try {
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
                $retryAttempts = $retryCount + 1
                Write-Progress -Activity 'Getting NIHO365GroupMembers' -Completed
                return $objectCollection
                
            }
            catch {
                if ($retryAttempts -lt $retryCount) {
                    $retryAttempts = $retryAttempts + 1        
                    Write-Host "Retry attempt number: $retryAttempts. Sleeping for $backOffInterval seconds..."
                    Start-Sleep $backOffInterval
                    $backOffInterval = $backOffInterval * 2
                }
                else {
                    Write-Verbose -Message 'Unable to getting NIHO365GroupMembers' -Verbose 
                }
            }
        }
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
    $resource = "https://graph.microsoft.com"
    $Uri = "$resource/$ApiVersion/groups"
         
    if (-not $MailNickName) { $MailNickName = $Name -replace "\W", '' }
    #Check if a group exists
    $group = (Invoke-RestMethod -Method Get -Headers $Header -Uri "$resource/$ApiVersion/groups?`$filter=mailNickname eq '$MailNickName'" ).value

    if (-not $group){
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
                if ($m.id) { $settings['members@odata.bind'] += "$resource/$ApiVersion/users/$($m.id)" }
                else { $settings['members@odata.bind'] += "$resource/$ApiVersion/users/$m" }
            }
        }
        if ($Owners) {
            $settings['owners@odata.bind'] = @()
            foreach ($o in $Owners) {
                if ($o.id) { $settings['owners@odata.bind'] += "$resource/$ApiVersion/users/$($o.id)" }
                else { $settings['owners@odata.bind'] += "$resource/$ApiVersion/users/$o" }
            }
        }  
        
        $Body = ConvertTo-Json $settings         
      
        if ($Force -or $PSCmdlet.shouldprocess($Name, "Add new Group")) {
            Write-Progress -Activity 'Creating Group' -CurrentOperation "Adding Group $Name"
            $group = Invoke-NIHGraph -Method "POST" -URI $Uri -Headers $Header -Body $Body
            Sleep 30                                  
        }
    }

    if ($group.id) { 
        $groupId = $group.id            
        Write-Progress -Activity 'Creating Group' -Completed        
        foreach ($m in $group.members) { if ($m.'@odata.type' -match "user") { $m.pstypenames.add("GraphUser") } }        
    }
    return $group
}
function Get-NIHO365GroupByAlias {
    [cmdletBinding()]
    param(
        [parameter(Mandatory = $true)]
        [hashtable]$AuthToken, 
        [parameter(Mandatory = $true)]
        [string]$MailNickName,                 
        [ValidateSet('beta', 'v1.0')] #Graph API version
        [string]$ApiVersion = 'v1.0'
    )
    begin {        
        # Create header
        $Header = @{       
            Authorization = $AuthToken['Authorization'] 
        }
        $resource = "https://graph.microsoft.com"
    }
    process {                
        if (-not $MailNickName) { $MailNickName = $Name -replace "\W", '' }
        $objectCollection = @()
        #Check if a group exists
        $Uri = "$resource/$ApiVersion/groups?`$filter=mailNickname eq '$MailNickName'"
        $Results = Invoke-NIHGraph -Method "Get" -URI $Uri -Headers $Header 
        if ($Results.value) {
            $objectCollection = ($Results.value)[0]
        } 
        $objectCollection
     }
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
        [ValidateSet('beta', 'v1.0')] #Graph API version
        [string]$ApiVersion = 'v1.0',
        #The group / team either as an ID or a group/team object with an IDn
        [Parameter(Mandatory = $true)]      
        $Group,
        #The user or nested-group to add, either as a UPN or ID or as a object with an ID
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        $Members,
        #If specified the user will be added as an owner, otherwise they will be a standard member
        [switch]$AsOwner
       
    )
    begin {        
        # Create header
        $Header = @{       
            Authorization = $AuthToken['Authorization'] 
        }
        $resource = "https://graph.microsoft.com" 
        $settings = @{}
        $groupID = $Group
        if ($Group.ID) { $groupID = $Group.ID }        
        $uri = "$resource/$ApiVersion/groups/$groupID"        
    }
    process {        
        if ($Members) {            
            $listofMembers = Get-NIHO365GroupMembers -AuthToken $AuthToken -Id $groupID
            $Ids = $listofMembers | select id
            $UPNs = $listofMembers | select userPrincipalName
            $userIds = [System.Collections.ArrayList]@()
            $userUPNs = [System.Collections.ArrayList]@()
            #$settings['members@odata.bind'] = @()

            $Ids | & { process { 
                if ($null -ne $_) { $userIds.Add($_.id) }
            }}
            $UPNs | & { process { 
                if ($null -ne $_) { $userUPNs.Add($_.userPrincipalName) }
            }}              
                       
            foreach ($m in $Members) {                
                if ($m.id) { $m = $m.id }  
                if (!($userIds.Contains($m)) -and !($userUPNs.Contains($m))){
                    if ($null -eq $settings['members@odata.bind']) { $settings['members@odata.bind'] = [System.Collections.ArrayList]@() }
                    #$settings['members@odata.bind']+= "$resource/$($ApiVersion)/directoryObjects/$m"                    
                    $settings['members@odata.bind'].Add("$resource/$($ApiVersion)/directoryObjects/$m")
                }
            }
        }
        if ($AsOwner) {
            $listofOwners = Get-NIHO365GroupOwners -AuthToken $AuthToken -Id $groupID            
            $Ids = $listofOwners | select id
            $UPNs = $listofOwners | select userPrincipalName
            $userIds = [System.Collections.ArrayList]@()
            $userUPNs = [System.Collections.ArrayList]@()            
            
            $Ids | & { process { 
                if ($null -ne $_) { $userIds.Add($_.id) }
            }}
            $UPNs | & { process { 
                if ($null -ne $_) { $userUPNs.Add($_.userPrincipalName) }
            }} 

            foreach ($m in $Members) {
                if ($m.id) { $m = $m.id } 
                if (!($userIds.Contains($m)) -and !($userUPNs.Contains($m))){
                    if ($null -eq $settings['owners@odata.bind']) { $settings['owners@odata.bind'] = [System.Collections.ArrayList]@() }
                    #$settings['owners@odata.bind']+= "$resource/$($ApiVersion)/directoryObjects/$m"
                    $settings['owners@odata.bind'].Add("$resource/$($ApiVersion)/directoryObjects/$m")
                }
            }
        }
        if ($settings -ne $null){
            $Body = ConvertTo-Json $settings
            Write-Debug $Body
            $results = Invoke-NIHGraph -Method "PATCH" -URI $Uri -Headers $Header -Body $Body
        }
        return $results
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
    <#-- Application permission not supported --#>
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
    $resource = "https://graph.microsoft.com"
    $Uri = "$resource//$($ApiVersion)/groups/$Id"    

    if(-not [string]::IsNullOrEmpty($HideFromOutlookClients) ){
        $settings = @{}
        $settings['hideFromOutlookClients'] = $HideFromOutlookClients
        $Body = ConvertTo-Json $settings        
        Invoke-NIHGraph -Method "PATCH" -URI $Uri -Headers $Header -Body $Body
    }
    if(-not [string]::IsNullOrEmpty($HideFromAddressLists) ){
        $settings = @{}
        $settings['hideFromAddressLists'] = $HideFromAddressLists  
        $Body = ConvertTo-Json $settings        
        Invoke-NIHGraph -Method "PATCH" -URI $Uri -Headers $Header -Body $Body
        Invoke-NIHGraph -Method "PATCH" -URI $Uri -Headers $Header -Body $Body
    }
}
function Set-NIHO365Group {
    [cmdletBinding()]
    param(
        [parameter(Mandatory = $true)]
        [hashtable]$AuthToken, 
        [ValidateSet('beta', 'v1.0')] #Graph API version
        [string]$ApiVersion = 'v1.0',         
        [parameter(Mandatory = $true)]
        [string]$Id,
        [parameter(Mandatory = $false)]        
        [string]$DisplayName,                
        [parameter(Mandatory = $false)]
        [string] $Description,
        [parameter(Mandatory = $false)]
        [string] $Visibility  
    )

    # Create header
    $Header = @{       
        Authorization = $AuthToken['Authorization']
    }
    $resource = "https://graph.microsoft.com"
    $Uri = "$resource/$ApiVersion/groups/$Id"
   
    $settings = @{              
                'Description' = $null                              
            }
    if(-not [string]::IsNullOrEmpty($DisplayName) ){ 
        $DisplayName = ([string]$DisplayName).Replace('“','"').Replace('”','"')       
        $settings['DisplayName'] = $DisplayName        
    }
    
    if(-not [string]::IsNullOrEmpty($Description) ){ 
        $Description = ([string]$Description).Replace('“','"').Replace('”','"')       
        $settings['Description'] = $Description        
    }
    if(-not [string]::IsNullOrEmpty($Visibility) ){        
        $settings['Visibility'] = $Visibility        
    }
    if ($settings.Count -eq 0) {
        exit
    }
    $Body = ConvertTo-Json $settings         
    Invoke-NIHGraph -Method "PATCH" -URI $Uri -Headers $Header -Body $Body
}
#endregion

#region Team
function Promote-NIHGroupToTeam{
[Cmdletbinding(SupportsShouldprocess = $true)]    
    param(
        [parameter(Mandatory = $true)]
        [hashtable]$AuthToken,
        [ValidateSet('beta', 'v1.0')] #Graph API version
        [string]$ApiVersion = 'v1.0', 
        #The Name of the group / team
        [Parameter(Mandatory = $true)]
        [string]$Id
    )      
    begin {        
        # Create header
        $Header = @{       
            Authorization = $AuthToken['Authorization'] 
        }
        $resource = "https://graph.microsoft.com"
    }
    process {
        #Check if a group exists then processing promote to Team, if not return
        $group = Get-NIHO365Group -AuthToken $AuthToken -ApiVersion $ApiVersion -Id $Id
        if (-not $group.id) {
            throw "Not found any group with GroupId=[$Id]"
            return
        }
        $uri = "$resource/$ApiVersion/groups/$Id/team"  
               
        $teamsettings = @{ 
                    "discoverySettings" = @{
                        "showInTeamsSearchAndSuggestions" = $true;
                    }
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
        
        $retryCount = 5
        $retryAttempts = 0
        $backOffInterval = 2        

        while ($retryAttempts -le $retryCount) {
            $team = Invoke-NIHGraph -Method "PUT" -URI $Uri -Headers $Header -Body $teamsettings                
            if ($team.id) { 
                $team.pstypenames.Add("GraphTeam")
                Write-Progress -Activity 'Creating Team' -Completed
                $retryAttempts = $retryCount + 1
                return $team
            }
            else {
                $retryAttempts = $retryAttempts + 1              
                Write-Progress -Activity 'Failing promote to team...'
                if ($retryAttempts -lt $retryCount) {                                
                    Write-Host "Retry attempt number: $retryAttempts. Sleeping for $backOffInterval seconds..."
                    Start-Sleep $backOffInterval
                    $backOffInterval = $backOffInterval * 2                        
                }
                else {
                    Write-Verbose -Message 'Unable to promote to team' -Verbose 
                    return $team
                }                    
                    
            }
        }
     }
}

function New-NIHTeamGroup {
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
    begin {        
        # Create header
        $Header = @{       
            Authorization = $AuthToken['Authorization']
            "Content-Type" = "application/json"
        }
        $resource = "https://graph.microsoft.com"
        $Uri = "$resource/$ApiVersion/groups"         
    }
    process {                
        if (-not $MailNickName) { $MailNickName = $Name -replace "\W", '' }
        #Check if a group exists
        $group = (Invoke-RestMethod -Method Get -Headers $Header -Uri "$resource/$ApiVersion/groups?`$filter=mailNickname eq '$MailNickName'" ).value
        #$Name = "`"@$Name`"@"
        #$Description = "`'$Description'"
        $Name = ([string]$Name).Replace('“','"').Replace('”','"')
        $Description = ([string]$Description).Replace('“','"').Replace('”','"')
                 
        if (-not $group){
            #HideGroupInOutlook,WelcomeEmailDisabled,SubscribeMembersToCalendarEventsDisabled            
            $settings = @{  
                'displayName'             = $Name                 
                'description'             = $Description
                'mailNickname'            = $MailNickName
                'mailEnabled'             = $true
                'securityEnabled'         = $false
                'visibility'              = $Visibility.ToLower()
                'groupTypes'              = @("Unified")  
                'resourceBehaviorOptions' = @('WelcomeEmailDisabled', 'SubscribeMembersToCalendarEventsDisabled')
                #'resourceBehaviorOptions' = @('WelcomeEmailDisabled', "HideGroupInOutlook", 'SubscribeMembersToCalendarEventsDisabled')
                #'resourceProvisioningOptions' = @("Team")                
            }
            
            if ($Members) {
                $settings['members@odata.bind'] = @();
                foreach ($m in $Members) {
                    if ($m.id) { $settings['members@odata.bind'] += "$resource/$ApiVersion/users/$($m.id)" }
                    else { $settings['members@odata.bind'] += "$resource/$ApiVersion/users/$m" }
                }
            }
            if ($Owners) {
                $settings['owners@odata.bind'] = @()
                foreach ($o in $Owners) {
                    if ($o.id) { $settings['owners@odata.bind'] += "$resource/$ApiVersion/users/$($o.id)" }
                    else { $settings['owners@odata.bind'] += "$resource/$ApiVersion/users/$o" }
                }
            }
                
            $Body = ConvertTo-Json $settings         
      
            if ($Force -or $PSCmdlet.shouldprocess($Name, "Add new Group")) {
                Write-Progress -Activity 'Creating Group' -CurrentOperation "Adding Group $Name"
                $group = Invoke-NIHGraph -Method "POST" -URI $Uri -Headers $Header -Body $Body
                Sleep 30                                  
            }
        }

        if ($group.id) { 
            $groupId = $group.id            
            Write-Progress -Activity 'Creating Group' -Completed
            if ($group.resourceProvisioningOptions -eq 'Team'){
                Write-Verbose -Message 'The group with Teams capabilities.' -Verbose 
                return $group
            }
            foreach ($m in $group.members) { if ($m.'@odata.type' -match "user") { $m.pstypenames.add("GraphUser") } }
            #promote to team after provision group
            $team = Promote-NIHGroupToTeam -AuthToken $AuthToken -ApiVersion $ApiVersion -Id $groupId
            return $team
        }
        else{
            return $group
        }   
     }
}

function New-NIHTeam {
    [Cmdletbinding(SupportsShouldprocess = $true)]    
    param(
        [parameter(Mandatory = $true)]
        [hashtable]$AuthToken,
        [ValidateSet('beta', 'v1.0')] #Graph API version
        [string]$ApiVersion = 'v1.0', 
        #The Name of team
        [Parameter(Mandatory = $true)]
        [string]$Name,        
        #A description for the team
        [string]$Description,
        #The visibility of the group, private by default
        [ValidateSet('private', 'public')]
        [string]$Visibility = 'private',        
        #User Principal Name or ID
        $Owner
    )
        
    begin {        
        # Create header
        $Header = @{       
            Authorization = $AuthToken['Authorization'] 
        }
        $resource = "https://graph.microsoft.com"
        $Uri = "$resource/$ApiVersion/teams" 
        
    }
    process {
        $role = "owner"
        $members = @{
            "user@odata.bind" = "$resource/$ApiVersion/users/$Owner"
            "roles" =  @("$role")
            "@odata.type" = "#microsoft.graph.aadUserConversationMember"        
            }
        #$members = @{}
        #$members["user@odata.bind"] = "$resource/$ApiVersion/users('8a7f6314-173f-457c-9d60-f613f1d44982')"
        #$members['user@odata.bind'] = "$resource/$ApiVersion/users/$Owner"
        #$members["roles"] = @("$role")
        #$members["@odata.type"] = "#microsoft.graph.aadUserConversationMember"
        
        $settings = @{  
               "template@odata.bind" = "$resource/$ApiVersion/teamsTemplates('standard')"
               "displayName" = $Name
               "description" = $Description
               "visibility" = $Visibility
               <#"members" = @(
                     @{
                        "@odata.type" = "#microsoft.graph.aadUserConversationMember"
                        "roles" =  @("owner")
                         "user@odata.bind" = "$resource/$ApiVersion/users('$Owner')"  
                         
                     })#>
               "members" = @($members)
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
                "discoverySettings" = @{
                    "showInTeamsSearchAndSuggestions" = $true
                }
        }
        $settings["members"] = $members        
        $Body = ConvertTo-Json $settings        
        $result = Invoke-NIHGraph -Method "POST" -URI $Uri -Headers $Header -Body $Body           
        return $result        
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
        #'hideFromOutlookClients' = $false 
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
        $Results = $null

        $retryCount = 5
        $retryAttempts = 0
        $backOffInterval = 2

        while ($retryAttempts -le $retryCount) {
            try {
                $Results = Invoke-NIHGraph -Method "Get" -URI $Uri -Headers $Header                          
                $retryAttempts = $retryCount + 1
                Write-Progress -Activity 'Getting NIHTeam' -Completed                
                return $Results                
            }
            catch {
                if ($retryAttempts -lt $retryCount) {
                    $retryAttempts = $retryAttempts + 1        
                    Write-Host "Retry attempt number: $retryAttempts. Sleeping for $backOffInterval seconds..."
                    Start-Sleep $backOffInterval
                    $backOffInterval = $backOffInterval * 2
                }
                else {
                    Write-Verbose -Message 'Unable to getting NIHTeam' -Verbose 
                }
            }
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
        
        $retryCount = 5
        $retryAttempts = 0
        $backOffInterval = 2

        while ($retryAttempts -le $retryCount) {
            try {
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
                $retryAttempts = $retryCount + 1
                Write-Progress -Activity 'Getting NIHTeamMembers' -Completed
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
            catch {
                if ($retryAttempts -lt $retryCount) {
                    $retryAttempts = $retryAttempts + 1        
                    Write-Host "Retry attempt number: $retryAttempts. Sleeping for $backOffInterval seconds..."
                    Start-Sleep $backOffInterval
                    $backOffInterval = $backOffInterval * 2
                }
                else {
                    Write-Verbose -Message 'Unable to getting NIHTeamMembers' -Verbose 
                }
            }
        }
    } 
}

function Add-NIHTeamMember {
    <#
      .Synopsis
        Adds a user to a team as either a member or owner.      
      .Example
        >Add-NIHTeamMember -Group $GroupID -Member member@xyz.sharepoint.com
        >Add-NIHTeamMember -Group $GroupID -Member admin@xyz.sharepoint.com -AsOwner        
    #>
    [Cmdletbinding(SupportsShouldprocess = $true)]    
    param   (
        [parameter(Mandatory = $true)]
        [hashtable]$AuthToken,
        [ValidateSet('beta', 'v1.0')] #Graph API version
        [string]$ApiVersion = 'v1.0', 
        #The group / team either as an ID or a team object with an IDn
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
        $resource = "https://graph.microsoft.com" 
        $groupID = $Group
        if ($Group.ID) { $groupID = $Group.ID }        
        $Uri = "$resource/$ApiVersion/teams/$groupID/members"
    }
    process {
        $results = @()  
        $role = "member"
        if ($AsOwner){
            $role = "owner"
        }
        $settings  = @{
                           "@odata.type" = "#microsoft.graph.aadUserConversationMember"  
                           "roles" =  @("$role")                           
                     }        
        foreach ($m in $Member) {            
            $settings['user@odata.bind'] = "$resource/$ApiVersion/users/$m"
            $Body = ConvertTo-Json $settings
            $results+= Invoke-NIHGraph -Method "POST" -URI $Uri -Headers $Header -Body $Body           

        } 
        return $results        
    }
}

function Add-NIHTeamMemberMultiple  {
    <#
      .Synopsis
        Add members in bulk to a team     
      .Example
        >Add-NIHTeamMemberMultiple -Group $GroupID -Member member@xyz.sharepoint.com
        >Add-NIHTeamMemberMultiple -Group $GroupID -Member admin@xyz.sharepoint.com -AsOwner        
    #>
    [Cmdletbinding(SupportsShouldprocess = $true)]    
    param   (
        [parameter(Mandatory = $true)]
        [hashtable]$AuthToken,
        [ValidateSet('beta', 'v1.0')] #Graph API version
        [string]$ApiVersion = 'v1.0', 
        #The group / team either as an ID or a team object with an IDn
        [Parameter(Mandatory = $true)]      
        $Group,
        #HashTable or Dictionary lookup (UserId/UPN - Role)
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        $Members
    )
    begin {
        if ($null -eq $Members -or $null -eq $Group ){exit} 
        # Create header
        $Header = @{       
            Authorization = $AuthToken['Authorization'] 
        }
        $resource = "https://graph.microsoft.com" 
        $groupID = $Group
        if ($Group.ID) { $groupID = $Group.ID }        
        $Uri = "$resource/$ApiVersion/teams/$groupID/members/add"
    }
    process {
        $results = @()        
        $values = [System.Collections.ArrayList]@()
        
        if ($Members){
            foreach($m in $Members.Keys){                                
                $role = @($members["$m"])
                if ([string]::IsNullOrEmpty($role)){
                    $role = @()
                }
                $member = @{
                    "@odata.type" = "microsoft.graph.aadUserConversationMember"
                    "roles" = $role 
                    "user@odata.bind" = "$resource/$ApiVersion/users('$m')"        
                }                
                $values.Add($member)                
            }
        }        
        $values = ConvertTo-Json $values
        $settings = @{
            #"@odata.context" = "$resource/$ApiVersion/$metadata#Collection(microsoft.graph.addConversationMemberResult)"
             "values" = $values
        }             
        
        $Body = ConvertTo-Json $settings        
        $results = Invoke-NIHGraph -Method "POST" -URI $Uri -Headers $Header -Body $Body
        return $results        
    }
}
#endregion

#region Team Channel
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

        $retryCount = 5
        $retryAttempts = 0
        $backOffInterval = 2

        while ($retryAttempts -le $retryCount) {
            try {
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
                $retryAttempts = $retryCount + 1
                Write-Progress -Activity "Get NIH Team Channels" -Completed
                return $objectCollection
                
            }
            catch {
                if ($retryAttempts -lt $retryCount) {
                    $retryAttempts = $retryAttempts + 1        
                    Write-Host "Retry attempt number: $retryAttempts. Sleeping for $backOffInterval seconds..."
                    Start-Sleep $backOffInterval
                    $backOffInterval = $backOffInterval * 2
                }
                else {
                    Write-Verbose -Message 'Unable to getting Team Channels' -Verbose 
                }
            }
        }
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

        $retryCount = 5
        $retryAttempts = 0
        $backOffInterval = 2

        while ($retryAttempts -le $retryCount) {
            try {
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
                $retryAttempts = $retryCount + 1
                Write-Progress -Activity "Get NIHTeamChannelMembers" -Completed
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
            catch {
                if ($retryAttempts -lt $retryCount) {
                    $retryAttempts = $retryAttempts + 1        
                    Write-Host "Retry attempt number: $retryAttempts. Sleeping for $backOffInterval seconds..."
                    Start-Sleep $backOffInterval
                    $backOffInterval = $backOffInterval * 2
                }
                else {
                    Write-Verbose -Message 'Unable to getting NIHTeamChannelMembers' -Verbose 
                }
            }
        }
    }  
}

#POST /teams/{id}/channels/{id}/members
function Add-NIHTeamChannelMember{
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
        $Member,
        [switch]$AsOwner
    ) 
    begin {        
        # Create header
        $Header = @{       
            Authorization = $AuthToken['Authorization']
        }
        $resource = "https://graph.microsoft.com"
        $Uri = "$resource/$ApiVersion/teams/$GroupId/channels/$ChannelId/members" 
           
    }
    process {
        $results = @()  
        $role = "member"
        if ($AsOwner){
            $role = "owner"
        }
        
        $settings  = @{
                           "@odata.type" = "#microsoft.graph.aadUserConversationMember"  
                           "roles" =  @("$role")                           
                     }        
        foreach ($m in $Member) {            
            $settings['user@odata.bind'] = "$resource/$ApiVersion/users/$m"
            $Body = ConvertTo-Json $settings
            $results+= Invoke-NIHGraph -Method "POST" -URI $Uri -Headers $Header -Body $Body           

        } 
        return $results        
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