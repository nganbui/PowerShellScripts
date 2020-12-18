Function GetAllM365Users {    
    $retrivalStartTime = Get-Date -Format "yyyy-MM-dd HH:mm:ss"       
    # Checking if authToken exists
    LogWrite -Message "Getting acces token using Graph API..."
    #Invoke-GraphAPIAuthTokenCheck
    $script:Certificate = Get-Item Cert:\\LocalMachine\\My\* | Where-Object { $_.Subject -ieq "CN=$($script:appCertAdminPortalOperation)" }    
    $script:authToken = Connect-NIHO365GraphWithCert -TenantName $script:TenantName -AppId $script:appIdAdminPortalOperation -Certificate $script:Certificate
    if ($script:authToken) {
        LogWrite -Message  "Retrieving Active M365 Users starting..."
        
        <#
        $script:Licenses = @{}        
        $allLicenses = Get-NIHLicenses -AuthToken $script:authToken
        $allLicenses.ForEach({            
            $script:Licenses[$_.skuId] = $_.skuPartNumber
        })
        #>      

        $allLicenses = Get-NIHLicenses -AuthToken $script:authToken
        [System.Collections.ArrayList]$script:Licenses = @()
        $allLicenses.ForEach({
            $null = $script:Licenses.Add([PSCustomObject]@{
                skuId                     = $_.skuId
                skuPartNumber             = $_.skuPartNumber
                servicePlans              = $_.servicePlans
            })
       }) 

        #$allUsers = Get-NIHO365Users -AuthToken $script:authToken                
        <#
        -- Currently we are using beta version to get Users's signInActivity
        -- Will change later when v1.0 support
        #> 
        #$allUsers = Get-NIHO365Users -AuthToken $script:authToken -ApiVersion beta 
        #$t = Get-NIHO365User -UserID "869946e3-bfa0-4952-b2d8-ed00fc7386bd" -AuthToken $script:authToken -Select id,userPrincipalName,givenName,surname,displayName,mail,proxyAddresses        
                
        $allUsers = Get-NIHO365Users -AuthToken $script:authToken -Select id,userPrincipalName,givenName,surname,displayName,mail,mailNickname,otherMails,businessPhones,mobilePhone,jobTitle,employeeId,
                                                                            userType,accountEnabled,createdDateTime,deletedDateTime,department,companyName,proxyAddresses,
                                                                            assignedLicenses,assignedPlans,provisionedPlans,
                                                                            officeLocation,streetAddress,city,state,country,postalCode,
                                                                            creationType,externalUserState,externalUserStateChangeDateTime,isResourceAccount,
                                                                            lastPasswordChangeDateTime,passwordPolicies,
                                                                            onPremisesDomainName,onPremisesLastSyncDateTime,onPremisesUserPrincipalName,
                                                                            onPremisesSamAccountName,onPremisesSecurityIdentifier,onPremisesSyncEnabled,onPremisesExtensionAttributes          
                                                                            
        if ($allUsers) {            
            #Parse Users
            LogWrite -Message  'Parsing M365 Users to pscustomoject starting...'
            $script:totalGuests = @($allUsers | Where-Object { $_.UserType -eq 'Guest' }).Count
            $script:totalMembers = @($allUsers | Where-Object { $_.UserType -eq 'Member' }).Count
            $script:totalOthers = @($allUsers | Where-Object { $_.UserType -ne 'Member' -and $_.UserType -ne 'Guest' }).Count
            $script:o365UsersData = ParseO365Users -Users $allUsers                    
        }                
        LogWrite -Message  "Retrieving Active M365 Users completed."
    }
    #Invoke-GraphAPIAuthTokenCheck    
    $script:authToken = Connect-NIHO365GraphWithCert -TenantName $script:TenantName -AppId $script:appIdAdminPortalOperation -Certificate $script:Certificate
    if ($script:authToken) {
        LogWrite -Message  'Retrieving Deleted M365 Users starting...'
        $script:o365DeletedUsersData = $null         
        $script:o365DeletedUsersData = Get-NIHDeletedO365Users -AuthToken $script:authToken
        $script:o365DeletedUsersData = ParseO365Users -Users $script:o365DeletedUsersData        
        LogWrite -Message  'Retrieving Deleted M365 Users completed.'
    }

    $retrivalEndTime = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    LogWrite -Message "Retrieval M365 Users Start Time: $($retrivalStartTime)"
    LogWrite -Message "Retrieval M365 Users End Time: $($retrivalEndTime)"    
    
}

Function ParseO365Users {
    [cmdletBinding()]
    param(
        [parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        $Users   
    )
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
                $assignedLicenses_ = $null
                $assignedPlans_ = $null

                if ($_.assignedPlans.length -gt 0){
                    $assignedPlans = @($_.assignedPlans | ? {$_.CapabilityStatus -eq "Enabled"})
                    if ($_.assignedLicenses.length -gt 0) { 
                        $isLicensed = 1                                
                        $l = @($_.assignedLicenses.skuId)
                        $l.ForEach({                    
                            $licenseId = $_                            
                            $script:Licenses.ForEach({
                                $skuId = $_.skuId
                                if ($skuId -eq $licenseId){
                                    $skuNumber = $_.skuPartNumber
                                    $servicePlans = $_.servicePlans
                                    $assignedLicenses_+=$skuNumber + ";"                            
                                }

                            })                    
                            $assignedPlans.ForEach({
                                $id = $_.servicePlanId
                                $servicePlanName = ($servicePlans | ? {$_.servicePlanId -eq $id}).servicePlanName
                                if ($null -ne $servicePlanName){
                                    $assignedPlans_+= $servicePlanName + ";"
                                    }                   
                            })                                        
                        }) 
                    }                    
                }

                <#                
                if ($_.assignedPlans.length -gt 0){
                    $_.assignedPlans |  ForEach-Object {
                        if ($_.CapabilityStatus -eq "Enabled"){                        
                            $assignedPlans_+= "$($_.Service);"
                        }
                    }
                }
                if ($_.assignedLicenses.length -gt 0){ 
                    $licenses = $_.assignedLicenses.skuId
                    $licenses.ForEach({                     
                         $assignedLicenses_+= $script:Licenses[$_] + ";"                  
                    }) 
                }              
                $provisionedPlans = $_.provisionedPlans
                #>
                <#
                if ($_.assignedLicenses.length -gt 0) { 
                    $isLicensed = 1                    
                   
                    LogWrite -Message  "Getting license details for the user [$upn]..."
                    $licenseDetails = Get-NIHLicensesByUser -AuthToken  $script:authToken -UserID $_.id                   
                                        
                    $assignedLicenses_ =  $licenseDetails.skuPartNumber -join ";"                     
                    $assignedPlans_ = ($licenseDetails.servicePlans | ? {$_.provisioningStatus -eq "Success"}).servicePlanName -join ";"
                    
                    }
                #> 
                    
                if ($_.accountEnabled -eq $false) { $isDisabled = 1}
                if ($_.deletedDateTime -ne $null) { $IsDeleted = 1}

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

Function CacheO365Users {
    LogWrite -Level INFO -Message "Generating Cache files for M365 Users..."    
    if ($script:o365UsersData -ne $null) {
        SetDataInCache -CacheType O365 -ObjectType O365Users -ObjectState Active -CacheData $script:o365UsersData
    }
    if ($script:o365DeletedUsersData -ne $null) {
        SetDataInCache -CacheType O365 -ObjectType O365Users -ObjectState InActive -CacheData $script:o365DeletedUsersData
    }
    LogWrite -Level INFO -Message "Generating Cache files for M365 Users completed."
    
}
