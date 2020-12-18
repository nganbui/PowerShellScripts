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
    try {         
        $conn = Connect-PnPOnline -Tenant $TenantId -ClientId $ClientId -Thumbprint $Thumbprint -Url $Url -ReturnConnection
    }
    catch {
        $ErrorMessage = $_.Exception.Message        
        Write-Verbose -Message "Unable to connect Sharepoint Online Session (PnPconnection) $($ErrorMessage)" -Verbose  
        Write-Verbose -Message "$($_.Exception)" -Verbose       
        $conn = $ErrorMessage
    }
    return $conn
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

#region access token
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
#endregion

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
    try {         
        $conn = Connect-MicrosoftTeams -TenantId $TenantId -ApplicationId $ClientId -CertificateThumbprint $Thumbprint           
    }
    catch {
        $ErrorMessage = $_.Exception.Message        
        Write-Verbose -Message "Unable to connect MicrosoftTeams $($ErrorMessage)" -Verbose  
        Write-Verbose -Message "$($_.Exception)" -Verbose        
        $conn = $ErrorMessage
    }
    return $conn
    
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

Function DisconnectPnpOnlineOAuth {
    param ($Context)
    try {
        $null = Disconnect-PnPOnline -Connection $Context
        Write-Verbose -Message 'The Sharepoint Online Session(pnpConnection) is now closed.' -Verbose        
    }
    catch {
        if ($_.Exception.Message -notmatch 'There is no service currently connected') {            
            Write-Verbose -Message 'Unable to disconnect Sharepoint Online Session (pnpConnection)' -Verbose
            return
        }
    }
}

