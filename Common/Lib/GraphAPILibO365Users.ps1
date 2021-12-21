Function GetAllM365Users {
    param([switch]$FullSync)

    try{
        $retrivalStartTime = Get-Date -Format "yyyy-MM-dd HH:mm:ss"       
        # Checking if authToken exists
        LogWrite -Message "Getting acces token using Graph API..."
        $script:authToken = Connect-GraphAPIWithCert -TenantId $script:TenantName -AppId $script:appIdAdminPortalOperation -Thumbprint $script:appThumbprintAdminPortalOperation

        if ($script:authToken) {
            [System.Collections.ArrayList]$script:ServicePlans = @()
            $script:Licenses = @()

            if ($FullSync.IsPresent){
                LogWrite -Message  "Retrieving M365 Licenses starting..."
                GetM365Licenses -AuthToken $script:authToken
            }
            #$allUsers = Get-NIHO365Users -AuthToken $script:authToken                
            <#
            -- Currently we are using beta version to get Users's signInActivity
            -- Will change later when v1.0 support
            #> 
            #$allUsers = Get-NIHO365Users -AuthToken $script:authToken -ApiVersion beta 
            #$t = Get-NIHO365User -UserID "869946e3-bfa0-4952-b2d8-ed00fc7386bd" -AuthToken $script:authToken -Select id,userPrincipalName,givenName,surname,displayName,mail,proxyAddresses        
            LogWrite -Message  "Retrieving Active M365 Users starting..."         
            $allUsers = Get-NIHO365Users -AuthToken $script:authToken -Select id,userPrincipalName,givenName,surname,displayName,mail,mailNickname,otherMails,businessPhones,mobilePhone,jobTitle,employeeId,
                                                                                userType,accountEnabled,createdDateTime,deletedDateTime,department,companyName,proxyAddresses,
                                                                                assignedLicenses,assignedPlans,provisionedPlans,
                                                                                officeLocation,streetAddress,city,state,country,postalCode,
                                                                                creationType,externalUserState,externalUserStateChangeDateTime,isResourceAccount,
                                                                                lastPasswordChangeDateTime,passwordPolicies,
                                                                                onPremisesDomainName,onPremisesLastSyncDateTime,onPremisesUserPrincipalName,
                                                                                onPremisesSamAccountName,onPremisesSecurityIdentifier,onPremisesSyncEnabled,onPremisesExtensionAttributes            
                                                                            
            if ($null -eq $allUsers){
                LogWrite -Message "All Users return null." 
                return
            }
            if ($allUsers) {            
                #Parse Users
                LogWrite -Message  'Parsing M365 Users to pscustomoject starting...'
                #$script:totalUsers = @($allUsers).Count 
                $script:totalGuests = @($allUsers | Where-Object { $_.UserType -eq 'Guest' }).Count
                $script:totalMembers = @($allUsers | Where-Object { $_.UserType -eq 'Member' }).Count
                $script:totalOthers = @($allUsers | Where-Object { $_.UserType -ne 'Member' -and $_.UserType -ne 'Guest' }).Count
                #$allUsers = @($allUsers)
                #$script:totalMembers = $allUsers.Count
                $script:o365UsersData = @()
                $script:o365UsersData = ParseO365Users -Users $allUsers             
            }                
            LogWrite -Message  "Retrieving Active M365 Users completed."
        }
        $script:authToken = Connect-GraphAPIWithCert -TenantId $script:TenantName -AppId $script:appIdAdminPortalOperation -Thumbprint $script:appThumbprintAdminPortalOperation
        if ($script:authToken) {
            LogWrite -Message  'Retrieving Deleted M365 Users starting...'
            $script:o365DeletedUsersData = @()         
            $script:o365DeletedUsersData = Get-NIHDeletedO365Users -AuthToken $script:authToken
            $script:o365DeletedUsersData = ParseO365Users -Users $script:o365DeletedUsersData        
            LogWrite -Message  'Retrieving Deleted M365 Users completed.'
        }

        $retrivalEndTime = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        LogWrite -Message "Retrieval M365 Users Start Time: $($retrivalStartTime)"
        LogWrite -Message "Retrieval M365 Users End Time: $($retrivalEndTime)" 
    }
    catch{
        LogWrite -Level ERROR "[GetAllM365Users]: Error Info: $($_.Exception)"
        #throw $_    
    }   
    
}

Function ParseO365Users {
    [cmdletBinding()]
    param(
        [parameter(Mandatory = $true)] $Users   
    )

    try{
        [System.Collections.ArrayList]$UsersList = @()    
        if ($Users -and $Users.Count -gt 0) {
            $i = 1
            $totalUsers = $Users.Count 
            $Users.ForEach( { 
                    $upn = $_.userPrincipalName
                    LogWrite -Message  "($i/$totalUsers) :Processing the user [$upn]..."
                    $isLicensed = 0
                    $isDisabled = 0
                    $IsDeleted = 0

                    if ($_.accountEnabled -eq $false) { $isDisabled = 1}
                    if ($_.deletedDateTime -ne $null) { $IsDeleted = 1}

                    #Licensed
                    if ($_.assignedLicenses.length -gt 0) { 
                            $isLicensed = 1  }

                    #AssignedLicenses and AssignedPlans run once/week
                    $assignedLicenses_ = $null
                    $assignedPlans_ = $null               

                    if ($script:Licenses.Count -gt 0 -and $script:ServicePlans.Count -gt 0){
                        $assignedLicenses = @($_.assignedLicenses | ? {$_.skuId}).skuId
                        $assignedPlans = @($_.assignedPlans | ? {$_.capabilityStatus -eq 'Enabled'}).servicePlanId 
                        $userLicenses = @()
                        $userPlans = @()

                        foreach($l in $assignedLicenses){        
                            $userLicenses+=LookupM365Sku -Skus $script:Licenses -Sku $l 
                        }
                        foreach($l in $assignedPlans){                    
                            $userPlans+=LookupM365Sku -Skus $script:ServicePlans -Sku $l 
                        }
                        $assignedLicenses_=$userLicenses -join ";"
                        $assignedPlans_=$userPlans -join ";"
                    }
                    #End of AssignedLicenses and AssignedPlans

                    #signInActivity
                    if ($null -ne $_.signInActivity -and $null -ne $_.signInActivity.lastSignInDateTime) { 
                        $lastSignInDateTime = Get-Date($_.signInActivity.lastSignInDateTime) -Format d
                    }
                    if ($null -ne $_.createdDateTime) { $createdDateTime      = Get-Date($_.createdDateTime) -Format d }
                    if ($null -ne $_.deletedDateTime) { $deletedDateTime      = Get-Date($_.deletedDateTime) -Format d }
                    if ($null -ne $_.lastPasswordChangeDateTime) { $lastPasswordChangeDateTime = Get-Date($_.lastPasswordChangeDateTime) -Format d }
                    if ($null -ne $_.onPremisesLastSyncDateTime) { $onPremisesLastSyncDateTime = Get-Date($_.onPremisesLastSyncDateTime) -Format d }
                    if ($null -ne $_.externalUserStateChangeDateTime) { $externalUserStateChangeDateTime = Get-Date($_.externalUserStateChangeDateTime) -Format d }
                    #email
                    if ($null -ne $_.mail) { $mail = "$($_.mail)" }
                    #proxyAddresses                
                    [System.Text.RegularExpressions.Regex] $regex = “^(smtp:.*|SMTP:.*)$”
                    if ($null -ne $_.ProxyAddresses){
                        $mail = ($_.ProxyAddresses -match $regex -replace "smtp:","") -join ";"
                    }
                    if ([string]::Empty -ne $mail){
                        $mail = "$($mail);"
                    }
                    <#
                    if ($null -ne $assignedLicenses_){
                        $assignedLicenses_ = "$($assignedLicenses_);"
                    }
                    if ($null -ne $assignedPlans_){
                        $assignedPlans_ = "$($assignedPlans_);"
                    }#>

                    $null = $UsersList.Add([PSCustomObject]@{
                            UserId                     = $_.id
                            DisplayName                = $_.displayName
                            GivenName                  = $_.givenName
                            Surname                    = $_.surname
                            UserPrincipalName          = $_.userPrincipalName
                            PrimaryEmail               = $_.mail
                            Mail                       = $mail                        
                            MailNickname               = $_.mailNickname
                            ProxyAddresses             = $_.proxyAddresses -join ';'
                            IsDisabled                 = $isDisabled
                            AccountEnabled             = $_.accountEnabled
                            IsDeleted                  = $IsDeleted
                            UserType                   = $_.userType
                            ICName                     = ""
                            Department                 = $_.department
                            CompanyName                = $_.companyName
                            LastSignInDateTime         = $lastSignInDateTime
                            IsLicensed                 = $isLicensed
                            AssignedLicenses           = $assignedLicenses_
                            AssignedPlans              = $assignedPlans_
                            CreatedDateTime            = $createdDateTime
                            SoftDeletionTimestamp      = $deletedDateTime
                            EmployeeId                 = $_.employeeId                        
                            JobTitle                   = $_.jobTitle;                        
                            BusinessPhones             = $_.businessPhones -join ';'                                             
                            MobilePhone                = $_.mobilePhone
                            OfficeLocation             = $_.officeLocation
                            StreetAddress              = $_.streetAddress
                            City                       = $_.city    
                            State                      = $_.state
                            Country                    = $_.country
                            PostalCode                 = $_.postalCode                        
                            IsResourceAccount          = $_.IsResourceAccount
                            LastPasswordChangeDateTime = $lastPasswordChangeDateTime
                            OnPremisesDomainName       = $_.onPremisesDomainName
                            OnPremisesLastSyncDateTime = $onPremisesLastSyncDateTime
                            OtherMails                 = $_.otherMails -join ';'
                            ShowInAddressList          = $_.showInAddressList
                            PreferredLanguage          = $_.preferredLanguage
                            MySite                     = $_.mySite # not available in Graph API
                            LastDirSyncTime            = $_.LastDirSyncTime # not available in Graph API
                            creationType               = $_.creationType
                            ExternalUserState          = $_.externalUserState
                            ExternalUserStateChangeDateTime = $externalUserStateChangeDateTime
                            OperationStatus            = ""
                            Operation                  = "" 
                            AdditionalInfo             = ""
                        }); 
                     $i++           
                })
        }
        return $UsersList
    }
    catch{
        LogWrite -Level ERROR "[ParseO365Users]: Error Info: $($_.Exception)"
        #throw $_    
    } 
}

Function CacheO365Users {
    LogWrite -Level INFO -Message "Generating Cache files for M365 Users..."    
    if ($script:o365UsersData -ne $null -and $script:o365UsersData.Count -gt 0) {
        SetDataInCache -CacheType O365 -ObjectType O365Users -ObjectState Active -CacheData $script:o365UsersData
    }
    if ($script:o365DeletedUsersData -ne $null -and $script:o365DeletedUsersData.Count -gt 0) {
        SetDataInCache -CacheType O365 -ObjectType O365Users -ObjectState InActive -CacheData $script:o365DeletedUsersData
    }
    if ($script:guestsData -ne $null -and $script:guestsData.Count -gt 0) {
        SetDataInCache -CacheType O365 -ObjectType O365Guests -ObjectState Active -CacheData $script:guestsData
    }
    LogWrite -Level INFO -Message "Generating Cache files for M365 Users completed."
    
}

Function GetGuestUsers { 
    try{   
        $retrivalStartTime = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
       
        LogWrite -Message "Getting acces token using Graph API..."    
        $script:authToken = Connect-GraphAPIWithCert -TenantId $script:TenantName -AppId $script:appIdAdminPortalOperation -Thumbprint $script:appThumbprintAdminPortalOperation

        if ($script:authToken) {
            #LogWrite -Message  "Retrieving Guest Users starting..."
            $guestUsers = Get-NIHO365Users -UserType Guest -AuthToken $script:authToken -ApiVersion beta -Select id,userPrincipalName,displayName,mail,otherMails,accountEnabled,assignedLicenses,creationType,externalUserState,externalUserStateChangeDateTime,createdDateTime,lastPasswordChangeDateTime,OnPremisesLastSyncDateTime,signInActivity
            if ($null -eq $guestUsers){
                LogWrite -Message "Guests return null." 
                return
            }
            $guestUsers = @($guestUsers)
                                                                                
            if ($guestUsers -and $guestUsers.Count -gt 0) {
                LogWrite -Message  "Retrieving Guest Users starting..." 
                #$guestUsers = @($guestUsers)           
                $script:totalGuests = $guestUsers.Count
                LogWrite -Message  "Total Guests retrieved: $($script:totalGuests)"
                $script:guestsData = @()
                $script:guestsData = ParseGuestUsers -Guests $guestUsers
                LogWrite -Message  "Retrieving Guest Users completed."
                #SetDataInCache -CacheType O365 -ObjectType O365Guests -ObjectState Active -CacheData $script:guestsData            
            }                
            #LogWrite -Message  "Retrieving Guest Users completed."
        }   

        $retrivalEndTime = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        LogWrite -Message "Retrieval Guest Users Start Time: $($retrivalStartTime)"
        LogWrite -Message "Retrieval Guest Users End Time: $($retrivalEndTime)" 
    }
    catch{
        LogWrite -Level ERROR "[GetGuestUsers]: Error Info: $($_.Exception)"
        #throw $_        
    }  
    
}

Function ParseGuestUsers {
    [cmdletBinding()]
    param(
        [parameter(Mandatory = $true)] $Guests   
    )
    try{
        [System.Collections.ArrayList]$GuestsList = @()    
        if ($Guests -and $Guests.Count -gt 0) {
            $totalGuests = $Guests.Count
            LogWrite -Message  "Total Guests are processing: $totalGuests"
            $i = 1
            $totalUsers = $Guests.Count 
            $Guests.ForEach( { 
                    $upn = $_.userPrincipalName
                    LogWrite -Message  "($i/$totalUsers) :Processing the guest user [$upn]..."
                    $isLicensed = 0
                    $isDisabled = 0
                    if ($_.assignedLicenses.length -gt 0) { $isLicensed = 1}                    
                    if ($_.accountEnabled -eq $false) { $isDisabled = 1}
                    #signInActivity
                    if ($null -ne $_.signInActivity -and $null -ne $_.signInActivity.lastSignInDateTime) { 
                        $lastSignInDateTime = Get-Date($_.signInActivity.lastSignInDateTime) -Format d
                    }
                    if ($null -ne $_.createdDateTime) { $createdDateTime      = Get-Date($_.createdDateTime) -Format d }                
                    if ($null -ne $_.lastPasswordChangeDateTime) { $lastPasswordChangeDateTime = Get-Date($_.lastPasswordChangeDateTime) -Format d }
                    if ($null -ne $_.onPremisesLastSyncDateTime) { $onPremisesLastSyncDateTime = Get-Date($_.onPremisesLastSyncDateTime) -Format d }
                    if ($null -ne $_.externalUserStateChangeDateTime) { $externalUserStateChangeDateTime = Get-Date($_.externalUserStateChangeDateTime) -Format d }
                    if ($null -ne $_.otherMails) { $otherEmails = $_.otherMails -join ";" }

                    $null = $GuestsList.Add([PSCustomObject]@{
                            UserId                     = $_.id
                            DisplayName                = $_.displayName                        
                            UserPrincipalName          = $_.userPrincipalName
                            PrimaryEmail               = $_.mail
                            OtherEmails                = $otherEmails
                            CreatedDateTime            = $createdDateTime
                            LastSignInDateTime         = $lastSignInDateTime                        
                            creationType               = $_.creationType
                            ExternalUserState          = $_.externalUserState
                            ExternalUserStateChangeDateTime = $externalUserStateChangeDateTime
                            LastPasswordChangeDateTime = $lastPasswordChangeDateTime                        
                            OnPremisesLastSyncDateTime = $onPremisesLastSyncDateTime
                            IsLicensed                 = $isLicensed                        
                            IsDisabled                 = $isDisabled                                                                                                                                       
                            OperationStatus            = ""
                            Operation                  = "" 
                            AdditionalInfo             = ""
                        }); 
                     $i++           
                })
        }
        return $GuestsList 
    }
    catch{
        LogWrite -Level ERROR "[ParseGuestUsers]: Error Info: $($_.Exception)"
        #throw $_
    }
}

Function GetM365Licenses{    
    [cmdletBinding()]
    param(
        [parameter(Mandatory = $true)]
        [hashtable]$AuthToken
    )
    [System.Collections.ArrayList]$script:ServicePlans = @()
    $script:Licenses = @()     
    $allLicenses = @(Get-NIHLicenses -AuthToken $AuthToken)
    if ($allLicenses -and $allLicenses.Count -gt 0){    
        $allLicenses.ForEach({
                $script:Licenses+= [Ordered] @{
                    $_.skuId = $_.skuPartNumber 
                }        
                $_.servicePlans | ? {$_.provisioningStatus -eq 'Success'} | select servicePlanId, servicePlanName | & { process {
                    $null = $script:ServicePlans.Add([Ordered]@{
                        $_.servicePlanId = $_.servicePlanName
                    })            
                }}         
            }) 
    }
}