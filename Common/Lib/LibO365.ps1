Function ConnectPnpOnlineOAuth {
    [CmdletBinding()]
    param(
        [parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()] 
        [string]$TenantId,
        [parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()] 
        [string]$ClientId,
        [parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()] 
        [string]$Thumbprint,
        [parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()] 
        [string]$Url
    )
    
    $retryCount = 5
    $retryAttempts = 0
    $backOffInterval = 2 

    while ($retryAttempts -le $retryCount) {
        try {
            $conn = Connect-PnPOnline -Tenant $TenantId -ClientId $ClientId -Thumbprint $Thumbprint -Url $Url -ReturnConnection
            $retryAttempts = $retryCount + 1
            return $conn
        }
        catch {
            if ($retryAttempts -lt $retryCount) {
                $retryAttempts = $retryAttempts + 1        
                #Write-Verbose "Retry attempt number: $retryAttempts. Sleeping for $backOffInterval seconds..."
                LogWrite "[ConnectPnpOnline]: Retry attempt number: $retryAttempts. Sleeping for $backOffInterval seconds..."                
                Start-Sleep $backOffInterval
                $backOffInterval = $backOffInterval * 2
            }
            else {
                $ErrorMessage = $_.Exception.Message        
                #Write-Verbose -Message "Unable to connect Sharepoint Online Session (PnPconnection) $($ErrorMessage)" -Verbose  
                #Write-Verbose -Message "$($_.Exception)" -Verbose
                LogWrite -Level ERROR "[ConnectPnpOnline]: Unable to connect Sharepoint Online Session (PnPconnection) $ErrorMessage"               
                throw
            }

        }        
    }     
}

Function DisconnectPnpOnlineOAuth {
    param ($Context)
    try {
        $null = Disconnect-PnPOnline -Connection $Context
        Write-Verbose -Message 'The Sharepoint Online Session(pnpConnection) is now closed.' -Verbose        
    }
    catch {
        if ($_.Exception.Message -notmatch 'There is no service currently connected') {            
            Write-Verbose -Message 'Unable to disconnect Sharepoint Online Session (pnpConnection)' -Verbose
            throw
        }
    }
}

Function ConnectMSOLServiceOAuth {
    [CmdletBinding()]
    param(
        [parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()] 
        [string]$TenantId,
        [parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()] 
        [string]$ClientId,
        [parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()] 
        [string]$Thumbprint,
        [parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()] 
        [string]$Url         

        
    )  
    Connect-MsolService -AdGraphAccessToken -MsGraphAccessToken
}

Function ConnectAzureADOAuth {
    [CmdletBinding()]
    param(
        [parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()] 
        [string]$TenantId,
        [parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()] 
        [string]$ClientId,
        [parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()] 
        [string]$Thumbprint
        
    )  
    
    $retryCount = 5
    $retryAttempts = 0
    $backOffInterval = 2 

    while ($retryAttempts -le $retryCount) {
        try {
            Connect-AzureAD -TenantId $TenantId -ApplicationId  $ClientId -CertificateThumbprint $Thumbprint | Out-Null
            $retryAttempts = $retryCount + 1
        }
        catch {
            if ($retryAttempts -lt $retryCount) {
                $retryAttempts = $retryAttempts + 1        
                #Write-Verbose "Retry attempt number: $retryAttempts. Sleeping for $backOffInterval seconds..."
                LogWrite "[ConnectAzureAD]: Retry attempt number: $retryAttempts. Sleeping for $backOffInterval seconds..."                
                Start-Sleep $backOffInterval
                $backOffInterval = $backOffInterval * 2
            }
            else {
                $ErrorMessage = $_.Exception.Message        
                #Write-Verbose -Message "Unable to connect Azure AD $($ErrorMessage)" -Verbose  
                #Write-Verbose -Message "$($_.Exception)" -Verbose
                LogWrite -Level ERROR "[ConnectAzureAD]: Unable to connect Azure AD $ErrorMessage"                
                throw
                #exit
            }

        }
    }
}

Function DisconnectAzureAD {
    try{
        Disconnect-AzureAD
    }
    catch{        
        Write-Verbose "Unable to disconnect Azure AD" -Fore Yellow        
    }    
}

Function ConnectMicrosoftTeams {

    [CmdletBinding()]
    param(
        [parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()] 
        [string]$TenantId,        
        [string]$ClientId,
        [parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()] 
        [string]$Thumbprint
    )
    $retryCount = 5
    $retryAttempts = 0
    $backOffInterval = 2 

    while ($retryAttempts -le $retryCount) {
        try {
            Connect-MicrosoftTeams -TenantId $TenantId -ApplicationId $ClientId -CertificateThumbprint $Thumbprint | Out-Null
            $retryAttempts = $retryCount + 1
        }
        catch {
            if ($retryAttempts -lt $retryCount) {
                $retryAttempts = $retryAttempts + 1        
                #Write-Verbose "Retry attempt number: $retryAttempts. Sleeping for $backOffInterval seconds..."
                LogWrite "[ConnectMicrosoftTeams]: Retry attempt number: $retryAttempts. Sleeping for $backOffInterval seconds..."
                Start-Sleep $backOffInterval
                $backOffInterval = $backOffInterval * 2
            }
            else {
                $ErrorMessage = $_.Exception.Message
                #Write-Verbose -Message "Unable to connect MicrosoftTeams $($ErrorMessage)" -Verbose
                #Write-Verbose -Message "$($_.Exception)" -Verbose
                LogWrite -Level ERROR "[ConnectMicrosoftTeams]: Unable to connect MicrosoftTeams $ErrorMessage"
                #exit
                throw
            }

        }
    }    
    
}

Function DisconnectMicrosoftTeams {
    try{
        Disconnect-MicrosoftTeams -ErrorVariable TeamsError
    }
    catch{
        if($TeamsError.Exception.Message -eq "Object reference not set to an instance of an object."){
        Write-Verbose "Microsoft Teams - No active Teams connections found" -Fore Yellow
        }
    }
    
}

#region Graph API
Function Invoke-GraphAPIAuthTokenCheck {
    <#
       .Description
       Check access token is valid.
    #>    
    $currentDateTimePlusTen = (Get-Date).AddMinutes(10)
    if ($script:authToken) {
        if (!($currentDateTimePlusTen -le $script:authToken["ExpiresOn"])) {                     
            # get an accesstoken if current accesstoken is valid but expired
            $script:authToken = Connect-NIHO365Graph            
        }        
    }
    else {
        # get an accesstoken if accesstoken is $null
        $script:authToken = Connect-NIHO365Graph 
        Invoke-GraphAPIAuthTokenCheck
    }    
}

Function Connect-GraphAPIWithCert{
    [CmdletBinding()]
    param(
        [parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()] 
        [string]$TenantId,
        [parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()] 
        [string]$AppId,
        [parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()] 
        [string]$Thumbprint
    )
    
    begin{        
        $retryCount = 5
        $retryAttempts = 0
        $backOffInterval = 2 
    }
    process{
        try{
            $Certificate = Get-Item Cert:\LocalMachine\My\* | Where-Object { $_.Thumbprint -ieq "$Thumbprint" } 
            $authToken = Connect-NIHO365GraphWithCert -TenantName $TenantId -AppId $AppId -Certificate $Certificate
             
            while ($null -ne $authToken.value -and $retryAttempts -lt $retryCount) {
                LogWrite "[Connect-GraphAPIWithCert]: Retry attempt number: $retryAttempts. Sleeping for $backOffInterval seconds..."
                Start-Sleep $backOffInterval
                $backOffInterval = $backOffInterval * 2
                $retryAttempts = $retryAttempts + 1
                $authToken = Connect-NIHO365GraphWithCert -TenantName $TenantId -AppId $AppId -Certificate $Certificate
            }
            return $authToken
        }
        catch{
            $ErrorMessage = $_.Exception.Message
            LogWrite -Level ERROR "[Connect-GraphAPIWithCert]: Unable to getting access token with Graph API: $ErrorMessage" 
            throw 
        }
    } 
}
#endregion

Function ValidateSite{
    [CmdletBinding()]
    param(        
        [parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()] 
        [string]$Url,
        [parameter(Mandatory = $true)]
        $SiteContext
    )
    begin{        
        $retryCount = 5
        $retryAttempts = 0
        $backOffInterval = 2 
    }
    process{
        try{
            $siteCollection = Get-PnPTenantSite -url $Url -Detailed -Connection $SiteContext 
            while ($siteCollection.Status -ne 'Active' -and $retryAttempts -lt $retryCount) {
                LogWrite -Message "[ValidateSite]: Waiting until the site is updated..."
                Start-Sleep $backOffInterval
                $backOffInterval = $backOffInterval * 2
                $retryAttempts = $retryAttempts + 1
                $siteCollection = Get-PnPTenantSite -url $Url -Detailed -Connection $SiteContext
            }
            return $siteCollection
        }
        catch{
            $ErrorMessage = $_.Exception.Message
            LogWrite -Level ERROR "[ValidateSite]: Unable to determine site is updated: $ErrorMessage" 
            throw 
        }
    } 
}

#region PowerBI
Function ConnectPowerBIService {
    [CmdletBinding()]
    param(
        [parameter(Mandatory = $true)]
        [ValidateSet("Public", "USGov")] $Environment = "USGov",
        [parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()] 
        [string]$Tenant,
        [parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()] 
        [string]$AppId,
        [parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()] 
        [string]$Thumbprint
       
    )
    
    begin{        
        $retryCount = 5
        $retryAttempts = 0
        $backOffInterval = 2 
    }
    process{
        while ($retryAttempts -le $retryCount) {
            try {
                Connect-PowerBIServiceAccount -Environment $Environment -ServicePrincipal -ApplicationId $AppId -Tenant $Tenant -CertificateThumbprint $Thumbprint
                $retryAttempts = $retryCount + 1
            }
            catch {
                if ($retryAttempts -lt $retryCount) {
                    $retryAttempts = $retryAttempts + 1                    
                    LogWrite "[ConnectPowerBIService]: Retry attempt number: $retryAttempts. Sleeping for $backOffInterval seconds..."                
                    Start-Sleep $backOffInterval
                    $backOffInterval = $backOffInterval * 2
                }
                else {
                    $ErrorMessage = $_.Exception.Message                    
                    LogWrite -Level ERROR "[ConnectPowerBIService]: Unable to connect PowerBIServiceAccount $ErrorMessage"                
                    throw                    
                }

            }
        }
    }     
}

Function DisconnectPowerBIService {
    try{
        Disconnect-PowerBIServiceAccount
    }
    catch{        
        Write-Verbose "Unable to Disconnect-PowerBIServiceAccount" -Fore Yellow        
    }    
}
#endregion
