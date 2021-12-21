Function GetAllM365Users {    
    $retrivalStartTime = Get-Date -Format "yyyy-MM-dd HH:mm:ss"       
    # Checking if authToken exists
    LogWrite -Message "Getting acces token using Graph API..."
    Invoke-GraphAPIAuthTokenCheck
    if ($script:authToken) {
        LogWrite -Message  "Retrieving Active M365 Users starting..."
        $script:Licenses = @{}
        $allLicenses = Get-NIHLicenses -AuthToken $script:authToken
        $allLicenses.ForEach({            
            $script:Licenses[$_.skuId] = $_.skuPartNumber
        })        
        $allUsers = Get-NIHO365Users -AuthToken $script:authToken                 
        if ($allUsers) {            
            #Parse Users
            LogWrite -Message  'Parsing M365 Users to pscustomoject starting...'            
            #$script:o365UsersData = ParseO365Users-old -usersObj $allUsers -ParseObjType O365
            $script:o365UsersData = ParseO365Users -Users $allUsers                    
        }                
        LogWrite -Message  "Retrieving Active M365 Users completed."
    }
    Invoke-GraphAPIAuthTokenCheck
    if ($script:authToken) {
        LogWrite -Message  'Retrieving Deleted O365 Users starting...'
        $script:o365DeletedUsersData = $null         
        $script:o365DeletedUsersData = Get-NIHDeletedO365Users -AuthToken $script:authToken
        $script:o365DeletedUsersData = ParseO365Users -Users $script:o365DeletedUsersData        
        LogWrite -Message  'Retrieving Deleted O365 Groups completed.'
    }
    $retrivalEndTime = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    LogWrite -Message "Retrieval O365 Groups Start Time: $($retrivalStartTime)"
    LogWrite -Message "Retrieval O365 Groups End Time: $($retrivalEndTime)"    
    
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
        $Users.ForEach( { 
                $isLicensed = 0
                $isDisabled = 0
                $IsDeleted = 0
                $assignedLicenses_ = $null
                $assignedPlans_ = $null
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
                if ($assignedLicenses_ -ne $null) { $isLicensed = 1}
                if ($_.accountEnabled -eq $false) { $isDisabled = 1}
                if ($_.deletedDateTime -ne $null) { $IsDeleted = 1}

                $null = $UsersList.Add([PSCustomObject]@{
                        UserId                     = $_.id
                        DisplayName                = $_.displayName;
                        GivenName                  = $_.givenName;
                        Surname                    = $_.surname;
                        UserPrincipalName          = $_.userPrincipalName;
                        Mail                       = $_.mail;
                        MailNickname               = $_.mailNickname
                        ProxyAddresses             = $_.proxyAddresses -join ';';
                        IsDisabled                 = $isDisabled;
                        AccountEnabled             = $_.accountEnabled
                        IsDeleted                  = $IsDeleted;
                        UserType                   = $_.userType;
                        Department                 = $_.department;
                        OfficeLocation             = $_.officeLocation;
                        CompanyName                = $_.companyName
                        IsLicensed                 = $isLicensed
                        AssignedLicenses           = $assignedLicenses_;
                        AssignedPlans              = $assignedPlans_;
                        CreatedDateTime            = $_.createdDateTime;
                        SoftDeletionTimestamp      = $_.deletedDateTime
                        #DeletedDateTime            = $_.deletedDateTime                       
                        EmployeeId                 = $_.employeeId                        
                        JobTitle                   = $_.jobTitle;                        
                        BusinessPhones             = $_.businessPhones -join ';';                                                
                        MobilePhone                = $_.mobilePhone;
                        StreetAddress              = $_.streetAddress
                        City                       = $_.city    
                        State                      = $_.state;
                        Country                    = $_.country;
                        PostalCode                 = $_.postalCode                        
                        IsResourceAccount          = $_.IsResourceAccount
                        LastPasswordChangeDateTime = $_.lastPasswordChangeDateTime
                        OnPremisesDomainName       = $_.onPremisesDomainName
                        OnPremisesLastSyncDateTime = $_.onPremisesLastSyncDateTime                                                 
                        creationType               = $_.creationType
                        ExternalUserState          = $_.externalUserState
                        ExternalUserStateChangeDateTime = $_.externalUserStateChangeDateTime
                        OtherMails                 = $_.otherMails -join ','
                        ShowInAddressList          = $_.showInAddressList
                        MySite                     = $_.mySite
                        PreferredLanguage          = $_.preferredLanguage;                        
                        OperationStatus            = ""; 
                        Operation                  = ""; 
                        AdditionalInfo             = ""
                    });            
            })
    }
    return $UsersList 
}

Function CacheO365Users {
    LogWrite -Level INFO -Message "Generating Cache files for O365 Users..."    
    if ($script:o365UsersData -ne $null) {
        SetDataInCache -CacheType O365 -ObjectType O365Users -ObjectState Active -CacheData $script:o365UsersData
    }
    if ($script:o365DeletedUsersData -ne $null) {
        SetDataInCache -CacheType O365 -ObjectType O365Users -ObjectState InActive -CacheData $script:o365DeletedUsersData
    }
    LogWrite -Level INFO -Message "Generating Cache files for O365 Users completed."
    
}

#region no longer used
Function ParseO365Users-old {
    param(
        $usersObj,
        [ValidateSet("Active", "InActive")] $ObjectState = "Active",
        [ValidateSet("O365", "DB")] $ParseObjType = "O365"
    )
    
    #Parse/Format all users from O365 => UsersObject
    #---------------------------------------------------

    $usersFormattedData = @()

    foreach ($userObj in $usersObj) {
        #Initailize objects with Null value
        $username, $firstName, $lastName, $ICName, $email, $displayName, $office, $licenses, $whenCreated, $LastDirSyncTime, $UserPrincipalName, $LastPasswordChangeTimeStamp, $StsRefreshTokensValidFrom, $StreetAddress, $City, $State, $Country, $UserType, $UserThemeIdentifierForO365Shell, $PhoneNumber, $MobilePhone, $PostalCode, $PasswordNeverExpires, $OverallProvisioningStatus, $Title, $ValidationStatus, $SoftDeletionTimestamp, $LiveId, $MSRtcSipPrimaryUserAddress, $AppLicenses = $null
    
        $signinName = ReplaceSingleQuote -inStr $userObj.SignInName
        $firstName = ReplaceSingleQuote -inStr $userObj.FirstName
        $lastName = ReplaceSingleQuote -inStr $userObj.LastName
        
        $displayName = ReplaceSingleQuote -inStr $userObj.DisplayName
        
        #The object names are different in DB to O365 so we parse them seperately. 
        #And also for values like ResourceUsed,StorageUsed etc, the value from O365 is in MB, where as while updating to DB we convert it to GB ---We dont have to convert them to GB here, we can handle it while reteriving
        switch ($ParseObjType) {
            "O365" {
                $ICName = $userObj.Department
                $UserId = $userObj.ObjectId
                $Created = $userObj.WhenCreated
                $isDisabled = [System.Convert]::ToBoolean($userObj.BlockCredential)
                $userObj.Licenses | ForEach-Object { $licenses += "$($_.AccountSkuId);" }
                $AppLicenses = ($userObj.Licenses.ServiceStatus | ? { $_.ProvisioningStatus -eq "Success" }).ServicePlan.ServiceName -join ';'
                $email = "$(ReplaceSingleQuote -inStr $userObj.SignInName);"                                
                $userObj.ProxyAddresses | ForEach-Object { $email += "$(ReplaceSingleQuote -inStr $_.toLower().Replace('smtp:',''));" }
                
            }
            "DB" {
                $ICName = $userObj.ICName
                $UserId = $userObj.UserId
                $Created = $userObj.Created
                $isDisabled = $userObj.IsDisabled
                $licenses = $userObj.Licenses
                $AppLicenses = $userObj.AppLicenses
                $email = $userObj.Email
            }
        }

        
  
        if ($isDisabled -ne $true) {
            $isDisabled = 0
        }
        else {
            $isDisabled = 1
        }
        
        if ($ObjectState -ne "Active") {
            $isDeleted = 1;
        }
        else {
            $isDeleted = 0;
        }
  
        $isLicensed = $userObj.IsLicensed
        
        if ($isLicensed -ne $true) {
            $isLicensed = 0
        }
        else {
            $isLicensed = 1
        }


        $usersFormattedData += [pscustomobject]@{
            UserID                          = $UserId;
            DisplayName                     = $displayName;
            FirstName                       = $firstName;
            LastName                        = $lastName 
            SigninName                      = $signinName;
            Email                           = $email;
            ICName                          = $ICName; 
            Office                          = $userObj.Office;
            IsLicensed                      = $isLicensed;
            Licenses                        = $licenses;
            IsDisabled                      = $isDisabled 
            IsDeleted                       = $isDeleted;
            Created                         = $Created;
            LastDirSyncTime                 = $userObj.LastDirSyncTime;
            UserPrincipalName               = $userObj.UserPrincipalName;
            LastPasswordChangeTimeStamp     = $userObj.LastPasswordChangeTimeStamp;
            StsRefreshTokensValidFrom       = $userObj.StsRefreshTokensValidFrom;
            StreetAddress                   = $userObj.StreetAddress;
            City                            = $userObj.City;
            State                           = $userObj.State;
            Country                         = $userObj.Country;
            UserType                        = $userObj.UserType;
            UserThemeIdentifierForO365Shell = $userObj.UserThemeIdentifierForO365Shell;
            PhoneNumber                     = $userObj.PhoneNumber;
            MobilePhone                     = $userObj.MobilePhone;
            PostalCode                      = $userObj.PostalCode;
            PasswordNeverExpires            = $userObj.PasswordNeverExpires;
            OverallProvisioningStatus       = $userObj.OverallProvisioningStatus;
            Title                           = $userObj.Title;
            JobTitle                        = $userObj.JobTitle;
            ValidationStatus                = $userObj.ValidationStatus;
            SoftDeletionTimestamp           = $userObj.SoftDeletionTimestamp;
            LiveId                          = $userObj.LiveId;
            MSRtcSipPrimaryUserAddress      = $userObj.MSRtcSipPrimaryUserAddress;
            AppLicenses                     = $AppLicenses; 
            PersonalSiteURL                 = $UserObj.PersonalSiteURL;
            PersonalSiteFirstCreationTime   = $UserObj.PersonalSiteFirstCreationTime;
            PersonalSiteLastCreationTime    = $UserObj.PersonalSiteLastCreationTime;
            IsVIP                           = $userObj.IsVIP
            OperationStatus                 = ""; 
            Operation                       = ""; 
            AdditionalInfo                  = ""
        }
    }
    return $usersFormattedData
}

Function GetAllO365Users {
    try {        
        LogWrite -Message "Connecting to MSOL Service..."        
        ConnectMSOLService -Credential $script:o365AdminCredential
        LogWrite -Message "Connected successfully to the MSOL Service."        
    }
    catch {    
        LogWrite -Level ERROR -Message "Unable to connect MSOL Service"
        LogWrite -Level ERROR -Message "$($_.Exception)"
        exit
    }
    try {
        $retrivalStartTime = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

        #region O365 Users        
        LogWrite -Message "Retrieving Active Users from O365..."                
        $allUsers = Get-MsolUser -All | Select-Object *
        $script:usersData = ParseO365Users -usersObj $allUsers -ParseObjType O365
        LogWrite -Message "Retrieving Soft Deleted Users from O365..."
        $deletedUsers = Get-MsolUser -All -ReturnDeletedUsers | Select-Object *
        $script:deletedUsersData = ParseO365Users -usersObj $deletedUsers -ObjectState InActive -ParseObjType O365
        #endregion

        $retrivalEndTime = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        LogWrite -Message "Retrieval O365 Users Start Time: $($retrivalStartTime)"
        LogWrite -Message "Retrieval O365 Users End Time: $($retrivalEndTime)"
        
    }
    catch {
        LogWrite -Level ERROR -Message "An error occured $($_.Exception)"        
    }
    finally {        
        LogWrite -Message "Disconnecting MSOL Service..."            
        LogWrite -Message "The MSOL Service Session is now closed."    
    }
    
}
#endregion