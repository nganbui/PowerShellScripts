function Get-NIHO365AccessToken {
    [cmdletBinding()]
    param(
        [parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()] 
        [string]$profilePath = $global:ProfilePath 
    )
   
    $fileExists = Test-Path $profilePath    
    if ($fileExists -eq $false) {
        Write-Error "File not found $profilePath"
        exit;
    }
    $profileData = Get-Psd1Data $profilePath
    $TenantId = $profileData.TenantConfig["Id"]
    $Resource = "https://graph.microsoft.com"    
    $o365AppKeyFile = "$($profileData.Path["Cred"])\$($profileData.AppConfigAdminPortal["AppSecret"])"   
    $o365AppCredential = GetAppCredential -ApplicationID $profileData.AppConfigAdminPortal["AppId"] -ApplicationKey $o365AppKeyFile  
    
    $body = @{
        grant_type    = "client_credentials"
        resource      = $Resource      
        client_id     = $o365AppCredential.UserName
        client_secret = $o365AppCredential.GetNetworkCredential().Password
    }

    try {
        $token = Invoke-RestMethod -Method Post -Uri "https://login.microsoftonline.com/$($TenantId)/oauth2/token" -Body $body -ErrorAction Stop
        #$token = Invoke-RestMethod -Method Post -Uri "https://login.microsoftonline.com/$($TenantId)/oauth2/v2.0/token" -Body $body -ErrorAction Stop                
        $($token.access_token)
    }
    catch {
        $ErrorMessage = $_.Exception.Message -replace "`n", " " -replace "`r", " "
        Write-Error -Message "FAILED - Unable to retreive access token using Graph API: $ErrorMessage"
    }
}

function Convert-AuthResponse {
    [cmdletBinding()]
    param(        
        [PSCredential] $Credential,
        [string] $TenantDomain        
    )    
    $Resource = "https://graph.microsoft.com"
    $body = @{
        grant_type    = "client_credentials"
        resource      = $Resource      
        client_id     = $Credential.UserName
        client_secret = $Credential.GetNetworkCredential().Password
    }

    try {
        $token = Invoke-RestMethod -Method Post -Uri "https://login.microsoftonline.com/$($TenantDomain)/oauth2/token" -Body $body -ErrorAction Stop
        $graphTokenExpirationDate = (Get-Date).AddHours(1)
        if ($token) {                  
            @{
                "Authorization" = "$($token.token_type) $($token.access_token)"
                "Content-Type"  = "application/json"
                "ExpiresOn"     = $graphTokenExpirationDate
            }                    
        }
    }
    catch {
        $ErrorMessage = $_.Exception.Message -replace "`n", " " -replace "`r", " "
        Write-Error -Message "FAILED - Unable to retreive MS Graph API Authentication Token: $ErrorMessage"
    }
}

#region being used in provisioning site
function Connect-NIHO365GraphV1 {
    [cmdletBinding()]
    param(
        [parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()] 
        [string]$profilePath = "D:\Scripting\O365\Config\PROFILE.psd1" 
    )         
    $fileExists = Test-Path $profilePath    
    if ($fileExists -eq $false) {
        Write-Error "File not found $profilePath"
        exit;
    }
    $profileData = Get-Psd1Data $profilePath
    $o365AppKeyFile = "$($profileData.Path["Config"])\$($profileData.AppConfig["AppSecret"])"   
    $o365AppCredential = GetAppCredential -ApplicationID $profileData.AppConfig["AppId"] -ApplicationKey $o365AppKeyFile        
    #Convert-AuthResponse -ApplicationID $profileData.AppConfig["AppId"] -ApplicationKey $o365AppCredential.GetNetworkCredential().Password -TenantDomain $profileData.AppConfig["TenantName"]    
    Convert-AuthResponse -Credential  $o365AppCredential -TenantDomain $profileData.AppConfig["TenantName"]    
}
#endregion

function Connect-NIHO365Graph {
    [cmdletBinding()]
    param(
        [parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()] 
        [string]$profilePath = $global:ProfilePath 
    )         
    $fileExists = Test-Path $profilePath    
    if ($fileExists -eq $false) {
        Write-Error "File not found $profilePath"
        exit;
    }
    $profileData = Get-Psd1Data $profilePath
    $o365AppKeyFile = "$($profileData.Path["Cred"])\$($profileData.AppConfig["AppSecret"])"   
    $o365AppCredential = GetAppCredential -ApplicationID $profileData.AppConfig["AppId"] -ApplicationKey $o365AppKeyFile            
    Convert-AuthResponse -Credential  $o365AppCredential -TenantDomain $profileData.AppConfig["TenantName"]    
}

function Connect-NIHO365GraphDP {
    [cmdletBinding()]
    param(
        [parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()] 
        [string]$profilePath = $global:ProfilePath, 
        [PSCredential] $Credential,
        $Scope="https://graph.microsoft.com/.default"
    )
    try {
        $fileExists = Test-Path $profilePath    
        if ($fileExists -eq $false) {
            Write-Error "File not found $profilePath"
            exit;
        }
        #Read profileData
        $profileData = Get-Psd1Data $profilePath
        $o365AppKeyFile = "$($profileData.Path["Cred"])\$($profileData.AppConfig["AppSecret"])"   
        $o365AppCredential = GetAppCredential -ApplicationID $profileData.AppConfig["AppId"] -ApplicationKey $o365AppKeyFile
        # tenant ID
        $TenantName = $profileData.AppConfig["TenantName"] 
        #tokenBody
        $ReqTokenBody = @{
            Grant_Type    = "Password"        
            client_Id     = $o365AppCredential.UserName
            Client_Secret = $o365AppCredential.GetNetworkCredential().Password
            Username      = $Credential.username
            Password      = $Credential.GetNetworkCredential().Password                    
            Scope         = $Scope
        }
  
        $GraphAPITokenRequestError = $null
        $authResult = Invoke-RestMethod -Method Post -Uri "https://login.microsoftonline.com/$TenantName/oauth2/v2.0/token" -Body $ReqTokenBody -ErrorVariable $GraphAPITokenRequestError
        $graphTokenExpirationDate = (Get-Date).AddHours(1)
        # If there's an error requesting the token, say so, display the error, and exit:
        if ($GraphAPITokenRequestError) {
            Write-Error "FAILED - Unable to retreive MS Graph API Authentication Token - $($GraphAPITokenRequestError)"        
            exit
        }
        if ($authResult) {                
            # Creating header for Authorization token  
            $authHeader = @{
                'Content-Type'  = 'application/json'            
                'Authorization' = "$($authResult.token_type) $($authResult.access_token)"
                'ExpiresOn'     = $graphTokenExpirationDate
            }  
            return $authHeader  
        } 
    }
    catch {
        $ErrorMessage = $_.Exception.Message -replace "`n", " " -replace "`r", " "
        Write-Error -Message "Something went wrong - Error: $ErrorMessage"
    }       
}

function Get-NIHAccessToken{
    [cmdletBinding()]
    param(
        [parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()] 
        [string]$TenantName, 
        [parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()] 
        $AppId, 
        [parameter(Mandatory = $true)]
        $AppSecret,
        [parameter(Mandatory = $false)]
        $Host = "https://login.microsoftonline.com",
        [parameter(Mandatory = $true)]
        $Resource = "https://graph.microsoft.com"

    )    
    try {
        $profilePath = $global:ProfilePath
        $fileExists = Test-Path $profilePath    
        if ($fileExists -eq $false) {
            Write-Error "File not found $profilePath"
            exit;
        }
        $profileData = Get-Psd1Data $profilePath
        $o365AppKeyFile = "$($profileData.Path["Cred"])\$AppSecret"   
        $o365AppCredential = GetAppCredential -ApplicationID $AppId -ApplicationKey $o365AppKeyFile        
        $body = @{
            grant_type    = "client_credentials"
            resource      = $Resource      
            client_id     = $AppId
            client_secret = $o365AppCredential.GetNetworkCredential().Password
        }             
        $token = Invoke-RestMethod -Method Post -Uri "$Host/$TenantName/oauth2/token" -Body $body -ErrorAction Stop
        $graphTokenExpirationDate = (Get-Date).AddHours(1)
        if ($token) {                  
            @{
                "Authorization" = "$($token.token_type) $($token.access_token)"
                "Content-Type"  = "application/json"
                "ExpiresOn"     = $graphTokenExpirationDate
            }                    
        }
    }
    catch {
        $ErrorMessage = $_.Exception.Message -replace "`n", " " -replace "`r", " "
        Write-Error -Message "Something went wrong : $ErrorMessage"
    } 
}

function Connect-NIHO365GraphWithCert {
    [cmdletBinding()]
    param(
        [parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()] 
        [string]$TenantName, 
        [parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()] 
        [string]$AppId, 
        [parameter(Mandatory = $true)]
        $Certificate,
        [parameter(Mandatory = $false)]
        $Scope = "https://graph.microsoft.com/.default"

    )    
    try {
        # Create base64 hash of certificate
        $CertificateBase64Hash = [System.Convert]::ToBase64String($Certificate.GetCertHash())

        # Create JWT timestamp for expiration
        $StartDate = (Get-Date "1970-01-01T00:00:00Z" ).ToUniversalTime()
        $JWTExpirationTimeSpan = (New-TimeSpan -Start $StartDate -End (Get-Date).ToUniversalTime().AddMinutes(2)).TotalSeconds
        $JWTExpiration = [math]::Round($JWTExpirationTimeSpan,0)

        # Create JWT validity start timestamp
        $NotBeforeExpirationTimeSpan = (New-TimeSpan -Start $StartDate -End ((Get-Date).ToUniversalTime())).TotalSeconds
        $NotBefore = [math]::Round($NotBeforeExpirationTimeSpan,0)

        # Create JWT header
        $JWTHeader = @{
            alg = "RS256"           
            typ = "JWT"
            # Use the CertificateBase64Hash and replace/strip to match web encoding of base64
            x5t = $CertificateBase64Hash -replace '\+','-' -replace '/','_' -replace '='
        }

        # Create JWT payload
        $JWTPayLoad = @{
            # What endpoint is allowed to use this JWT
            aud = "https://login.microsoftonline.com/$TenantName/oauth2/token"

            # Expiration timestamp
            exp = $JWTExpiration

            # Issuer = your application
            iss = $AppId

            # JWT ID: random guid
            jti = [guid]::NewGuid()

            # Not to be used before
            nbf = $NotBefore

            # JWT Subject
            sub = $AppId
        }

        # Convert header and payload to base64
        $JWTHeaderToByte = [System.Text.Encoding]::UTF8.GetBytes(($JWTHeader | ConvertTo-Json))
        $EncodedHeader = [System.Convert]::ToBase64String($JWTHeaderToByte)

        $JWTPayLoadToByte =  [System.Text.Encoding]::UTF8.GetBytes(($JWTPayload | ConvertTo-Json))
        $EncodedPayload = [System.Convert]::ToBase64String($JWTPayLoadToByte)

        # Join header and Payload with "." to create a valid (unsigned) JWT
        $JWT = $EncodedHeader + "." + $EncodedPayload

        # Get the private key object of your certificate
        #$PrivateKey = $Certificate.PrivateKey
        # Get the private key object of your certificate
        $PrivateKey = ([System.Security.Cryptography.X509Certificates.RSACertificateExtensions]::GetRSAPrivateKey($Certificate))

        # Define RSA signature and hashing algorithm
        #$RSAPadding = [Security.Cryptography.RSASignaturePadding]::Pkcs1
        #$HashAlgorithm = [Security.Cryptography.HashAlgorithmName]::SHA1


        # Define RSA signature and hashing algorithm
        $RSAPadding = [Security.Cryptography.RSASignaturePadding]::Pkcs1
        $HashAlgorithm = [Security.Cryptography.HashAlgorithmName]::SHA256
        

        # Create a signature of the JWT
        $Signature = [Convert]::ToBase64String(
            $PrivateKey.SignData([System.Text.Encoding]::UTF8.GetBytes($JWT),$HashAlgorithm,$RSAPadding)
        ) -replace '\+','-' -replace '/','_' -replace '='

        # Join the signature to the JWT with "."
        $JWT = $JWT + "." + $Signature

        # Create a hash with body parameters
        $Body = @{
            client_id = $AppId
            client_assertion = $JWT
            client_assertion_type = "urn:ietf:params:oauth:client-assertion-type:jwt-bearer"
            scope = $Scope
            grant_type = "client_credentials"

        }

        $Url = "https://login.microsoftonline.com/$TenantName/oauth2/v2.0/token"

        # Use the self-generated JWT as Authorization
        $Header = @{
            Authorization = "Bearer $JWT"
        }

        # Splat the parameters for Invoke-Restmethod for cleaner code
        $PostSplat = @{
            ContentType = 'application/x-www-form-urlencoded'
            Method = 'POST'
            Body = $Body
            Uri = $Url
            Headers = $Header
        }

        $Request = Invoke-RestMethod @PostSplat
         # Creating header for Authorization token  
        $authHeader = @{
            'Content-Type'  = 'application/json'            
            'Authorization' = "$($Request.token_type) $($Request.access_token)"
            'Expires_in'     = $Request.expires_in
            'Ext_expires_in'     = $Request.ext_expires_in
        }  
        return $authHeader         
        
    }
    catch {
        $ErrorMessage = $_.Exception.Message -replace "`n", " " -replace "`r", " "
        Write-Error -Message "Something went wrong - Error: $ErrorMessage"
    }       
}

function Invoke-NIHGraph {
    Param (
        [parameter(Mandatory = $true)]
        [ValidateSet("GET", "POST", "PATCH", "DELETE", "PUT")]
        [String]$Method,

        [parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [String]$URI,

        [parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        $Headers,

        [parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [String]$Body
    )
    
    $RestResults = $null
    if ($PSBoundParameters.ContainsKey("Body")) {            
            try{
                $RestResults = Invoke-RestMethod -Method $Method -Uri $URI -Headers $Headers -ContentType "application/json" -Body $Body -Verbose
            }
            catch{                                            
                #$RestResults = $_.Exception
                $RestResults = $_
            }
        }
        else {            
            try{                
                $RestResults = Invoke-RestMethod -Method $Method -Uri $URI -Headers $Headers -ContentType "application/json" -Verbose
            }
            catch{                
                #$RestResults = $_.Exception
                $RestResults = $_
            }
        }
    <#try {
        if ($PSBoundParameters.ContainsKey("Body")) {
            $RestResults = Invoke-RestMethod -Method $Method -Uri $URI -Headers $Headers -ContentType "application/json" -Body $Body -Verbose
        }
        else {
            $RestResults = Invoke-RestMethod -Method $Method -Uri $URI -Headers $Headers -Verbose
        }
     
    }
    <#catch {
        $result = $_.Exception.Response.GetResponseStream()
        $reader = New-Object System.IO.StreamReader($result)
        $reader.BaseStream.Position = 0
        $reader.DiscardBufferedData()
        $responseBody = $reader.ReadToEnd()           
        $responseBody
    }#>    
    #$RestResults | ConvertTo-Json -Depth 3
    return $RestResults
}

function Get-Psd1Data {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory = $true)]
        [Microsoft.PowerShell.DesiredStateConfiguration.ArgumentToConfigurationDataTransformation()]
        [hashtable] $data
    )
    return $data
}

function GetAppCredential {
    [cmdletBinding()]
    param(
        [string][alias('ClientID')] $ApplicationID,
        [string][alias('ClientSecret')] $ApplicationKey            
    )   
    
    $o365AppPwd = Get-Content $ApplicationKey | ConvertTo-SecureString
    $o365AppCredential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $ApplicationID, $o365AppPwd   

    return  $o365AppCredential
}