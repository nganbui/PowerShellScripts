function Get-NIHO365Users {
    [cmdletBinding()]
    param(
        [parameter(Mandatory = $true)]
        [hashtable]$AuthToken,                 
        [ValidateSet('beta', 'v1.0')] #Graph API version
        [string]$ApiVersion = 'v1.0',        
        [parameter(Mandatory = $false)]
        [ValidateSet('Member','Guest')] 
        [String]$UserType,
        #specifies which properties of the user object should be returned
        [parameter(Mandatory = $false, parameterSetName = "Select")]
        [ValidateSet  ("id","userPrincipalName","givenName","surname","displayName","mail","mailNickname","otherMails","businessPhones","mobilePhone","jobTitle","employeeId",
            "userType","accountEnabled","createdDateTime","deletedDateTime","department","companyName","proxyAddresses",
            "assignedLicenses","assignedPlans","provisionedPlans",
            "officeLocation","streetAddress","city","state","country","postalCode",
            "creationType","externalUserState","externalUserStateChangeDateTime","isResourceAccount",
            "lastPasswordChangeDateTime","passwordPolicies","signInActivity","mySite",
            "onPremisesDomainName","onPremisesLastSyncDateTime","onPremisesUserPrincipalName","onPremisesSamAccountName","onPremisesSecurityIdentifier","onPremisesSyncEnabled","onPremisesExtensionAttributes",
            "showInAddressList","preferredLanguage","usageLocation"
            )]
        [String[]]$Select
    )    
    begin {        
        # Create header
        $Header = @{       
            Authorization = $AuthToken['Authorization']
        } 
    }
    process {
        Write-Progress -Activity 'Getting list of users'
        $objectCollection = @()
        <#$Uri = "https://graph.microsoft.com/$($ApiVersion)/users?`$select=id,businessPhones,displayName,givenName,jobTitle,mail,mobilePhone,officeLocation,
                                                            preferredLanguage,surname,userPrincipalName,employeeId,isResourceAccount,lastPasswordChangeDateTime,
                                                            officeLocation,onPremisesDomainName,onPremisesLastSyncDateTime,proxyAddresses,userType,accountEnabled,
                                                            createdDateTime,deletedDateTime,creationType,companyName,department,city,assignedLicenses,assignedPlans,provisionedPlans,
                                                            onPremisesSamAccountName,onPremisesSecurityIdentifier,onPremisesSyncEnabled,onPremisesUserPrincipalName,
                                                            passwordPolicies,streetAddress,state,country,postalCode,deletedDateTime,externalUserState,externalUserStateChangeDateTime,
                                                            mailNickname,otherMails,showInAddressList,signInActivity"; 
                                                            #mySite
        #>
        $Uri = "https://graph.microsoft.com/$ApiVersion/users"
        if ($Select) { 
            $Uri = $Uri + '?$select=' + ($Select -join ",") 
        }
        if ($UserType) { 
             $Uri = $Uri + "&`$filter=userType eq '$($UserType)'"
        }   
        $Uri = $Uri + "&`$top=999"        
        
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
                Write-Progress -Activity 'Getting list of users' -Completed
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
                    Write-Verbose -Message 'Unable to getting M365 Users' -Verbose 
                }
            }
        }
    }
}
function Get-NIHDeletedO365Users {
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
        Write-progress -Activity "Finding Deleted Users"
        $objectCollection = @()        
        #$Uri = "https://graph.microsoft.com/$($ApiVersion)/directory/deletedItems/microsoft.graph.group?`$filter=groupTypes/any(g:g eq 'Unified')"
        $Uri = "https://graph.microsoft.com/$($ApiVersion)/directory/deletedItems/microsoft.graph.user/?`$select=id,businessPhones,displayName,givenName,jobTitle,mail,mobilePhone,officeLocation,
                                                            preferredLanguage,surname,userPrincipalName,employeeId,isResourceAccount,lastPasswordChangeDateTime,
                                                            officeLocation,onPremisesDomainName,onPremisesLastSyncDateTime,proxyAddresses,userType,accountEnabled,
                                                            createdDateTime,deletedDateTime,creationType,companyName,department,city,assignedLicenses,assignedPlans,provisionedPlans,
                                                            onPremisesSamAccountName,onPremisesSecurityIdentifier,onPremisesSyncEnabled,onPremisesUserPrincipalName,
                                                            passwordPolicies,streetAddress,state,country,postalCode,deletedDateTime,externalUserState,externalUserStateChangeDateTime,
                                                            mailNickname,otherMails,showInAddressList,signInActivity&`$top=999"
        

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
                Write-Progress -Activity 'Finding Deleted Users' -Completed
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
                    Write-Verbose -Message 'Unable to getting M365 Users' -Verbose 
                }
            }
        }
    }
}
function Get-NIHO365GuestUsers {
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
        Write-progress -Activity "Getting Guest Users"
        $objectCollection = @()                        
        $Uri = "https://graph.microsoft.com/$($ApiVersion)/users?`$filter=userType eq 'Guest'&`$select=id,businessPhones,displayName,givenName,jobTitle,mail,mobilePhone,officeLocation,
                                                            preferredLanguage,surname,userPrincipalName,employeeId,isResourceAccount,lastPasswordChangeDateTime,
                                                            officeLocation,onPremisesDomainName,onPremisesLastSyncDateTime,proxyAddresses,userType,accountEnabled,
                                                            createdDateTime,deletedDateTime,creationType,companyName,department,city,assignedLicenses,assignedPlans,provisionedPlans,
                                                            onPremisesSamAccountName,onPremisesSecurityIdentifier,onPremisesSyncEnabled,onPremisesUserPrincipalName,
                                                            passwordPolicies,streetAddress,state,country,postalCode,deletedDateTime,externalUserState,externalUserStateChangeDateTime,
                                                            mailNickname,otherMails,showInAddressList,signInActivity&`$top=999"        
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
        Write-progress -Activity "Getting Guest Users" -Completed
        return $objectCollection

    }
}
function Get-NIHO365User {
    [cmdletBinding()]
    param(
        [parameter(Mandatory = $true)]
        [hashtable]$AuthToken,                 
        [ValidateSet('beta', 'v1.0')] #Graph API version
        [string]$ApiVersion = 'v1.0',
        #UserID as a guid or User Principal name. If not specified defaults to "me"
        [parameter(Mandatory = $true)]
        [string]$UserID,
        #specifies which properties of the user object should be returned
        [parameter(Mandatory = $false, parameterSetName = "Select")]
        [ValidateSet  ("aboutMe", "accountEnabled", "ageGroup", "assignedLicenses", "assignedPlans", "birthday", "businessPhones",
            "city", "companyName", "consentProvidedForMinor", "country", "createdDateTime", "department", "displayName", "givenName",
            "hireDate", "id", "imAddresses", "interests", "jobTitle", "legalAgeGroupClassification", "mail", "mailboxSettings",
            "mailNickname", "mobilePhone", "mySite", "officeLocation", "onPremisesDomainName", "onPremisesExtensionAttributes",
            "onPremisesImmutableId", "onPremisesLastSyncDateTime", "onPremisesProvisioningErrors", "onPremisesSamAccountName",
            "onPremisesSecurityIdentifier", "onPremisesSyncEnabled", "onPremisesUserPrincipalName", "passwordPolicies",
            "passwordProfile", "pastProjects", "postalCode", "preferredDataLocation", "preferredLanguage", "preferredName",
            "provisionedPlans", "proxyAddresses", "responsibilities", "schools", "skills", "state", "streetAddress",
            "surname", "usageLocation", "userPrincipalName", "userType","signInActivity")]
        [String[]]$Select="id,userPrincipalName,accountEnabled,displayName,department,mail,createdDateTime,assignedLicenses,assignedPlans"
    ) 
    begin {                
        if ($UserID) { $userID = "users/$userID" } else { $userid = "me" }
        # Create header
        $Header = @{       
            Authorization = $AuthToken['Authorization']
        } 
    }
    process {
        Write-Progress -Activity 'Getting user information'
        $uri = "https://graph.microsoft.com/$ApiVersion/$userID"; 
        
        if ($Select) { 
            $uri = $uri + '?$select=' + ($Select -join ",") 
        }
        try {
            $results = Invoke-NIHGraph -Method "Get" -URI $Uri -Headers $Header             
        }
        catch {
            if ($_.exception.response.statuscode.value__ -eq 404) {
                Write-Progress -Activity 'Getting user information' -Completed
                Write-Warning -Message "Not found error while getting data for user '$userid'" ; return
            }
            else { throw $_ ; return }
        }
        Write-Progress -Activity 'Getting user information' -Completed
        $results
    }
}
function Get-NIHO365UserByEmail {
    [cmdletBinding()]
    param(
        [parameter(Mandatory = $true)]
        [hashtable]$AuthToken,                 
        [ValidateSet('beta', 'v1.0')] #Graph API version
        [string]$ApiVersion = 'v1.0',        
        [parameter(Mandatory = $true)]
        [string]$EmailAddress,
        #specifies which properties of the user object should be returned
        [parameter(Mandatory = $false, parameterSetName = "Select")]
        [ValidateSet  ("aboutMe", "accountEnabled", "ageGroup", "assignedLicenses", "assignedPlans", "birthday", "businessPhones",
            "city", "companyName", "consentProvidedForMinor", "country", "createdDateTime", "department", "displayName", "givenName",
            "hireDate", "id", "imAddresses", "interests", "jobTitle", "legalAgeGroupClassification", "mail", "mailboxSettings",
            "mailNickname", "mobilePhone", "mySite", "officeLocation", "onPremisesDomainName", "onPremisesExtensionAttributes",
            "onPremisesImmutableId", "onPremisesLastSyncDateTime", "onPremisesProvisioningErrors", "onPremisesSamAccountName",
            "onPremisesSecurityIdentifier", "onPremisesSyncEnabled", "onPremisesUserPrincipalName", "passwordPolicies",
            "passwordProfile", "pastProjects", "postalCode", "preferredDataLocation", "preferredLanguage", "preferredName",
            "provisionedPlans", "proxyAddresses", "responsibilities", "schools", "skills", "state", "streetAddress",
            "surname", "usageLocation", "userPrincipalName", "userType","signInActivity")]
        [String[]]$Select = "id,mail,userPrincipalName,displayName,department"
    ) 
    begin {        
        # Create header
        $Header = @{       
            Authorization = $AuthToken['Authorization']
        } 
    }
    process {
        Write-Progress -Activity 'Getting user information from email address'
        $uri = "https://graph.microsoft.com/$ApiVersion/users?`$filter=mail eq '$EmailAddress'"; 
        
        if ($Select) { 
            $uri = $uri + '&$select=' + ($Select -join ",") 
        }
        try {
            $results = Invoke-NIHGraph -Method "Get" -URI $Uri -Headers $Header
            $userInfo = @($results.value)
            if ($userInfo -and $userInfo.Count -gt 0) {
                $results = $userInfo[0]
            }
                         
        }
        catch {
            if ($_.exception.response.statuscode.value__ -eq 404) {
                Write-Progress -Activity 'Getting user information' -Completed
                Write-Warning -Message "Not found error while getting data for user '$userid'" ; return
            }
            else { throw $_ ; return }
        }
        Write-Progress -Activity 'Getting user information from email address' -Completed
        $userInfo
    }
}
function Get-NIHLicenses{    
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
        Write-Progress -Activity 'Getting licenses'
        $objectCollection = @()
        $Uri = "https://graph.microsoft.com/$($ApiVersion)/subscribedSkus"; 

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
                Write-Progress -Activity 'Getting licenses' -Completed
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
                    Write-Verbose -Message 'Unable to getting licenses' -Verbose 
                }
            }
        }
    }
}
function Get-NIHLicensesByUser{    
    [cmdletBinding()]
    param(
        [parameter(Mandatory = $true)]
        [hashtable]$AuthToken,                 
        [ValidateSet('beta', 'v1.0')] #Graph API version
        [string]$ApiVersion = 'v1.0',
        #UserID as a guid or User Principal name. If not specified defaults to "me"
        [parameter(Mandatory = $true)]
        [string]$UserID
    )    
    begin {                
        if ($UserID) { $userID = "users/$userID" } else { $userid = "me" }
        # Create header
        $Header = @{       
            Authorization = $AuthToken['Authorization']
        } 
    }
    process {
        Write-Progress -Activity 'Getting licenses for individual user'
        $Uri = "https://graph.microsoft.com/$($ApiVersion)/$userID/licenseDetails"; 

        try {
            $objectCollection = @()   
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
        }
        catch {
            if ($_.exception.response.statuscode.value__ -eq 404) {
                Write-Progress -Activity 'Getting licenses for individual user' -Completed
                Write-Warning -Message "Error while Getting licenses for individual user" ; return
            }
            else { throw $_ ; return }
        }
        Write-Progress -Activity 'Getting licenses for individual user' -Completed
        $objectCollection
    }
}
function Get-NIHO365UserMemberGroups{    
    [cmdletBinding()]
    param(
        [parameter(Mandatory = $true)]
        [hashtable]$AuthToken,                 
        [ValidateSet('beta', 'v1.0')] #Graph API version
        [string]$ApiVersion = 'v1.0',
         #UserID as a guid or User Principal name. If not specified defaults to "me"
        [parameter(Mandatory = $true)]
        [string]$UserID
    )    
    begin {        
        # Create header
        $Header = @{       
            Authorization = $AuthToken['Authorization']
        } 
    }
    process {
        Write-Progress -Activity 'Getting groups that the user is a member of'
        $Uri = "https://graph.microsoft.com/$($ApiVersion)/users/$userID/getMemberGroups";
        <#
        ====
        "securityEnabledOnly" settings
                $true: to specify that only security groups that the user is a member of should be returned; 
                $false to specify that all groups that the user is a member of should be returned. 
        Note: Setting this parameter to true is only supported when calling this method on a user.
        #>
        $settings = @{  
             'securityEnabledOnly' = $false
        }
        $Body = ConvertTo-Json $settings  

        try {            
            $objectCollection = @()   
            $Results = Invoke-NIHGraph -Method "POST" -URI $Uri -Headers $Header -Body $Body 
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
        }
        catch {
            if ($_.exception.response.statuscode.value__ -eq 404) {
                Write-Progress -Activity 'Getting groups that the user is a member of' -Completed
                Write-Warning -Message "Error while Getting groups that the user is a member of" ; return
            }
            else { throw $_ ; return }
        }
        Write-Progress -Activity 'Getting groups that the user is a member of' -Completed
        $objectCollection
    }
}
function Get-NIHO365UserMemberObjects{    
    [cmdletBinding()]
    param(
        [parameter(Mandatory = $true)]
        [hashtable]$AuthToken,                 
        [ValidateSet('beta', 'v1.0')] #Graph API version
        [string]$ApiVersion = 'v1.0',
         #UserID as a guid or User Principal name. If not specified defaults to "me"
        [parameter(Mandatory = $true)]
        [string]$UserID
    )    
    begin {        
        # Create header
        $Header = @{       
            Authorization = $AuthToken['Authorization']
        } 
    }
    process {
        Write-Progress -Activity 'Getting all of the groups, directory roles and administrative units that the user is a member of'
        $Uri = "https://graph.microsoft.com/$($ApiVersion)/users/$userID/getMemberObjects";
        <#
        True: to specify that only security groups that the user is a member of should be returned; 
        false to specify that all groups that the user is a member of should be returned. 
        Note: Setting this parameter to true is only supported when calling this method on a user.
        #>
        $settings = @{  
             'securityEnabledOnly' = $false
        }
        $Body = ConvertTo-Json $settings  

        try {            
            $objectCollection = @()   
            $Results = Invoke-NIHGraph -Method "POST" -URI $Uri -Headers $Header -Body $Body 
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
        }
        catch {
            if ($_.exception.response.statuscode.value__ -eq 404) {
                Write-Progress -Activity 'Getting all of the groups, directory roles and administrative units that the user is a member of' -Completed
                Write-Warning -Message "Error while Getting all of the groups, directory roles and administrative units that the user is a member of" ; return
            }
            else { throw $_ ; return }
        }
        Write-Progress -Activity 'Getting all of the groups, directory roles and administrative units that the user is a member of' -Completed
        $objectCollection
    }
}
function Get-NIHO365ObjectById{    
    [cmdletBinding()]
    param(
        [parameter(Mandatory = $true)]
        [hashtable]$AuthToken,                 
        [ValidateSet('beta', 'v1.0')] #Graph API version
        [string]$ApiVersion = 'v1.0',
         #UserID as a guid or User Principal name. If not specified defaults to "me"
        [parameter(Mandatory = $true)]
        $Ids #A collection of IDs for which to return objects. The IDs are GUIDs, represented as strings. Specify up to 1000 IDs.
    )    
    begin {        
        # Create header
        $Header = @{       
            Authorization = $AuthToken['Authorization']
        } 
    }
    process {
        Write-Progress -Activity 'Return the directory objects specified in a list of IDs'
        $Uri = "https://graph.microsoft.com/$($ApiVersion)/directoryObjects/getByIds";
        
        $settings = @{  
              "ids"  = $Ids
        }
        $Body = ConvertTo-Json $settings  

        try {            
            $objectCollection = @()   
            $Results = Invoke-NIHGraph -Method "POST" -URI $Uri -Headers $Header -Body $Body 
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
        }
        catch {
            if ($_.exception.response.statuscode.value__ -eq 404) {
                Write-Progress -Activity 'Return the directory objects specified in a list of IDs' -Completed
                Write-Warning -Message "Error while Return the directory objects specified in a list of IDs" ; return
            }
            else { throw $_ ; return }
        }
        Write-Progress -Activity 'Return the directory objects specified in a list of IDs' -Completed
        $objectCollection
    }
}
function Get-NIHO365UserAppRoleAssignments{    
    [cmdletBinding()]
    param(
        [parameter(Mandatory = $true)]
        [hashtable]$AuthToken,                 
        [ValidateSet('beta', 'v1.0')] #Graph API version
        [string]$ApiVersion = 'v1.0',
         #UserID as a guid or User Principal name. If not specified defaults to "me"
        [parameter(Mandatory = $true)]
        [string]$UserID
    )    
    begin {        
        # Create header
        $Header = @{       
            Authorization = $AuthToken['Authorization']
        } 
    }
    process {
        Write-Progress -Activity 'Getting app roles assigned to a user'
        $Uri = "https://graph.microsoft.com/$($ApiVersion)/users/$userID/appRoleAssignments";
        try {            
            $objectCollection = @()   
            $Results = Invoke-NIHGraph -Method "GET" -URI $Uri -Headers $Header
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
        }
        catch {
            if ($_.exception.response.statuscode.value__ -eq 404) {
                Write-Progress -Activity 'Getting app roles assigned to a user' -Completed
                Write-Warning -Message "Error while Getting app roles assigned to a user" ; return
            }
            else { throw $_ ; return }
        }
        Write-Progress -Activity 'Getting app roles assigned to a user' -Completed
        $objectCollection
    }
}
function Get-NIHO365UserDelegatedPermGrant{    
    [cmdletBinding()]
    param(
        [parameter(Mandatory = $true)]
        [hashtable]$AuthToken,                 
        [ValidateSet('beta', 'v1.0')] #Graph API version
        [string]$ApiVersion = 'v1.0',
         #UserID as a guid or User Principal name. If not specified defaults to "me"
        [parameter(Mandatory = $true)]
        [string]$UserID
    )    
    begin {        
        # Create header
        $Header = @{       
            Authorization = $AuthToken['Authorization']
        } 
    }
    process {
        Write-Progress -Activity 'Getting a list of oauth delegated permission granted to enable a client app to access an APO on behalf of the user'
        $Uri = "https://graph.microsoft.com/$($ApiVersion)/users/$userID/oauth2PermissionGrants"

        try {            
            $objectCollection = @()   
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
        }
        catch {
            if ($_.exception.response.statuscode.value__ -eq 404) {
                Write-Progress -Activity 'Getting a list of oauth delegated permission granted to enable a client app to access an APO on behalf of the user' -Completed
                Write-Warning -Message "Error while Getting a list of oauth delegated permission granted to enable a client app to access an APO on behalf of the user" ; return
            }
            else { throw $_ ; return }
        }
        Write-Progress -Activity 'Getting a list of oauth delegated permission granted to enable a client app to access an APO on behalf of the user' -Completed
        $objectCollection
    }
}
function Get-NIHO365UserRelevant {
    [cmdletBinding()]
    param(
        [parameter(Mandatory = $true)]
        [hashtable]$AuthToken,                 
        [ValidateSet('beta', 'v1.0')] #Graph API version
        [string]$ApiVersion = 'v1.0',
        #UserID as a guid or User Principal name. If not specified defaults to "me"
        [parameter(Mandatory = $true)]
        [string]$UserID
    ) 
    $Header = @{       
        Authorization = $AuthToken['Authorization']
    } 
    $myBatchRequests = @()
    [int]$requestID = 0
    $requestID ++

    $myRequest = [pscustomobject][ordered]@{ 
        id     = $requestID
        method = "GET"
        url    = "/users/$UserID"
    } 
    $myBatchRequests += $myRequest

    $requestID ++
    $myRequest = [pscustomobject][ordered]@{ 
        id     = $requestID
        method = "GET"
        url    = "/users/$UserID/memberOf"
    } 
    $myBatchRequests += $myRequest

    $allBatchRequests = [pscustomobject][ordered]@{ 
        requests = $myBatchRequests
    }

    $batchBody = $allBatchRequests | ConvertTo-Json
    $batchUrl = "https://graph.microsoft.com/v1.0/$batch"
    $getBatchRequests = Invoke-RestMethod -Method Post -Uri $batchUrl -Body $batchBody -headers $Header -ContentType "application/json"
    $getBatchRequests

    # foreach ($jobRMResult in $getBatchRequestRM.responses) {
    #     $jobRMResult.id
    #     write-host -ForegroundColor blue "jobID: $($jobRMResult.id)"
    #     if ($jobRMResult.body.value.count -gt 1){
    #         foreach ($entry in $jobRMResult.body.value){
    #             write-host -ForegroundColor cyan "  $($entry)"        
    #         }
    #     } else {
    #         write-host -ForegroundColor cyan "  $($jobRMResult.body.value)"    
    #     }    
    # }

}

#region Get groups and directory roles that the user is a direct member of.
function Get-NIHO365UserIsMemberOfGroups{
    [cmdletBinding()]
    param(
        [parameter(Mandatory = $true)]
        [hashtable]$AuthToken,                 
        [ValidateSet('beta', 'v1.0')] #Graph API version
        [string]$ApiVersion = 'v1.0',
         #UserID as a guid or User Principal name. If not specified defaults to "me"
        [parameter(Mandatory = $true)]
        [string]$UserID
    )    
    begin {        
        # Create header
        $Header = @{       
            Authorization = $AuthToken['Authorization']
        }
        $resource = "https://graph.microsoft.com"
    }
    process {
        Write-progress -Activity "Getting Groups that $UserID is member of..."
        $objectCollection = @()
        $Uri = "$resource/$ApiVersion/users/$UserID/memberOf/microsoft.graph.group?`$top=999"
        $Results = Invoke-NIHGraph -Method "Get" -URI $Uri -Headers $Header 
        
        $objectCollection = $Results.value
        $NextLink = $Results.'@odata.nextLink'

        while ($NextLink) {        
                $Results = (Invoke-NIHGraph -Method "Get" -Uri $NextLink -Headers $Header)
                $NextLink = $Results.'@odata.nextLink'
                $objectCollection += $Results.value
            } 
        Write-progress -Activity "Getting Groups that user is member of" -Completed
        
    }
    end{
        return $objectCollection
    }
}
#endregion

function Get-NIHO365UserOneDrive {
    [cmdletBinding()]
    param(
        [parameter(Mandatory = $true)]
        [hashtable]$AuthToken,                 
        [ValidateSet('beta', 'v1.0')] #Graph API version
        [string]$ApiVersion = 'v1.0',
        #UserID as a guid or User Principal name. If not specified defaults to "me"
        [parameter(Mandatory = $true)]
        [string]$UserID
    ) 
    begin {                
        if ($UserID) { $userID = "users/$userID" } else { $userid = "me" }
        # Create header
        $Header = @{       
            Authorization = $AuthToken['Authorization']
        } 
    }
    process {
        Write-Progress -Activity 'Getting OneDrive'
        #/users/{idOrUserPrincipalName}/drive
        #$uri = "https://graph.microsoft.com/$ApiVersion/$userID/drive"
        $uri = "https://graph.microsoft.com/$ApiVersion/$userID/drive/root/children";         
        
        try {
            $results = Invoke-NIHGraph -Method "Get" -URI $Uri -Headers $Header             
        }
        catch {
            if ($_.exception.response.statuscode.value__ -eq 404) {
                Write-Progress -Activity 'Getting OneDrive' -Completed
                Write-Warning -Message "Not found error while getting onedrive for user '$userid'" ; return
            }
            else { throw $_ ; return }
        }
        Write-Progress -Activity 'Getting OneDrive' -Completed
        $results
    }
}