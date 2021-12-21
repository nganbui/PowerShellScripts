#region Enumerate sites
function Get-NIHAllSites {
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
        $resource = "https://graph.microsoft.com"
        $uri = "$resource/$ApiVersion/sites?search=*&`$top=999"       
    }
    process {
        Write-progress -Activity "Getting all sharepoint sites"
        $objectCollection = @()        
        
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
                Write-Progress -Activity 'Getting all sharepoint sites' -Completed
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
                    Write-Verbose -Message 'Unable getting all sharepoint sites' -Verbose 
                }
            }
        }        
    }
}

function Get-NIHRootSites {
    [cmdletBinding()]
    param(
        [parameter(Mandatory = $true)]
        [hashtable]$AuthToken,                 
        [ValidateSet('beta', 'v1.0')] #Graph API version
        [string]$ApiVersion = 'v1.0',
        [parameter(Mandatory = $false, parameterSetName = "Select")]        
        [String[]]$Select = "siteCollection,webUrl" 
    )    
    begin {        
        # Create header
        $Header = @{       
            Authorization = $AuthToken['Authorization'] 
        }
        $resource = "https://graph.microsoft.com"
        $uri = "$resource/$ApiVersion/sites?`$filter=siteCollection/root ne null&`$top=999"
        if ($Select) { 
            $uri = $uri + '&$select=' + ($Select -join ",") 
        }
    }
    process {
        Write-progress -Activity "Getting all sharepoint sites"
        $objectCollection = @()        
        
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
                Write-Progress -Activity 'Getting all sharepoint sites' -Completed
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
                    Write-Verbose -Message 'Unable getting all sharepoint sites' -Verbose 
                }
            }
        }        
    }
}
#endregion

#region Site Details
function Get-NIHSiteById {    
    [cmdletBinding()]
    param(
        [parameter(Mandatory = $true)]
        [hashtable]$AuthToken,                 
        [ValidateSet('beta', 'v1.0')] #Graph API version
        [string]$ApiVersion = 'v1.0',
        [Parameter(Mandatory = $true)]
        [string]$SiteId
    )    
    begin {        
        # Create header
        $Header = @{       
            Authorization = $AuthToken['Authorization'] 
        }
        $resource = "https://graph.microsoft.com"        
    }
    process {
        Write-progress -Activity "Getting a site by site ID"        
        $Uri = "$resource/$ApiVersion/sites/$siteId"

        try {
            $results = Invoke-NIHGraph -Method "Get" -URI $Uri -Headers $Header            
        }
        catch {
            if ($_.exception.response.statuscode.value__ -eq 404) {
                Write-Progress -Activity 'Getting a site by site ID' -Completed
                Write-Warning -Message "Not found error while getting a site by site ID" ; return
            }
            else { throw $_ ; return }
        }
        Write-Progress -Activity 'Getting a site by site ID' -Completed
        $results
        
    }
}

function Get-NIHSiteByRelativeURL {    
    [cmdletBinding()]
    param(
        [parameter(Mandatory = $true)]
        [hashtable]$AuthToken,                 
        [ValidateSet('beta', 'v1.0')] #Graph API version
        [string]$ApiVersion = 'v1.0',
        [Parameter(Mandatory = $true)]
        [string]$Hostname,
        [Parameter(Mandatory = $true)]
        [string]$RelativeURL 
    )    
    begin {        
        # Create header
        $Header = @{       
            Authorization = $AuthToken['Authorization'] 
        }
        $resource = "https://graph.microsoft.com"        
    }
    process {
        Write-progress -Activity "Getting a site by server-relative URL"        
        $Uri = "$resource/$ApiVersion/sites/$Hostname`:/$RelativeURL"

        try {
            $results = Invoke-NIHGraph -Method "Get" -URI $Uri -Headers $Header            
        }
        catch {
            if ($_.exception.response.statuscode.value__ -eq 404) {
                Write-Progress -Activity 'Getting a site by server-relative URL' -Completed
                Write-Warning -Message "Not found error while getting a site by server-relative URL" ; return
            }
            else { throw $_ ; return }
        }
        Write-Progress -Activity 'Getting a site by server-relative URL' -Completed
        $results
        
    }
}
#endregion

#region Enumerate items in a list
# Get the collection of items in a list.
function Get-NIHListItems {
    [cmdletBinding()]
    param(
        [parameter(Mandatory = $true)]
        [hashtable]$AuthToken,                 
        [ValidateSet('beta', 'v1.0')] #Graph API version
        [string]$ApiVersion = 'v1.0',
        [Parameter(Mandatory = $true)]
        [string]$SiteId,
        [Parameter(Mandatory = $true)]
        [string]$ListId,
        [Parameter(Mandatory = $false)]
        $Fields
    )    
    begin {        
        # Create header
        $Header = @{       
            Authorization = $AuthToken['Authorization'] 
        }
        $resource = "https://graph.microsoft.com"
    }
    <#
    process {
        Write-progress -Activity "Get the collection of items in a list."

        $Uri = "$resource/$ApiVersion/sites/$SiteId/lists/$ListId/items"        

        if ($Fields) { 
            #$Uri = "$Uri?expand=fields(select=$Fields -join ","))"
            $Uri = $Uri + '?expand=fields(select=' + ($Fields -join ",") + ")" 
        }
        
        try {
            $results = Invoke-NIHGraph -Method "Get" -URI $Uri -Headers $Header            
        }
        catch {
            if ($_.exception.response.statuscode.value__ -eq 404) {
                Write-Progress -Activity 'Get the collection of items in a list.' -Completed
                Write-Warning -Message "Not found error while Get the collection of items in a list." ; return
            }
            else { throw $_ ; return }
        }
        Write-Progress -Activity 'Get the collection of items in a list.' -Completed
        $results
    }
    #>
    process {
        Write-progress -Activity "Get the collection of items in a list."
        $objectCollection = @()
        $Uri = "$resource/$ApiVersion/sites/$SiteId/lists/$ListId/items?expand=fields"        

        if ($Fields) {             
            $Uri = $Uri + '(select=' + ($Fields -join ",") + ")" 
        }
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
        Write-Progress -Activity 'Get the collection of items in a list.' -Completed
        return $objectCollection

    }
}
#endregion

#region ListItem
function Get-NIHListItem {
    [cmdletBinding()]
    param(
        [parameter(Mandatory = $true)]
        [hashtable]$AuthToken,                 
        [ValidateSet('beta', 'v1.0')] #Graph API version
        [string]$ApiVersion = 'v1.0',
        [Parameter(Mandatory = $true)]
        [string]$SiteId,
        [Parameter(Mandatory = $true)]
        [string]$ListId,
        [Parameter(Mandatory = $true)]
        [string]$ItemId
    )    
    begin {        
        # Create header
        $Header = @{       
            Authorization = $AuthToken['Authorization'] 
        }
        $resource = "https://graph.microsoft.com"
    }
    process {
        Write-progress -Activity "Returns the metadata for an item in a list."

        $Uri = "$resource/$ApiVersion/sites/$SiteId/lists/$ListId/items/$ItemId"

        <#if ($Expand) { 
            $uri = $uri + '?expand=fields 
        }#>
        try {
            $results = Invoke-NIHGraph -Method "Get" -URI $Uri -Headers $Header            
        }
        catch {
            if ($_.exception.response.statuscode.value__ -eq 404) {
                Write-Progress -Activity 'Getting a list item information' -Completed
                Write-Warning -Message "Not found error while a list item information" ; return
            }
            else { throw $_ ; return }
        }
        Write-Progress -Activity 'Getting a list item information' -Completed
        $results
    }
}

function Add-NIHListItem {
    [cmdletBinding()]
    param(
        [parameter(Mandatory = $true)]
        [hashtable]$AuthToken,                 
        [ValidateSet('beta', 'v1.0')] #Graph API version
        [string]$ApiVersion = 'v1.0',
        [Parameter(Mandatory = $true)]
        [string]$SiteId,
        [Parameter(Mandatory = $true)]
        [string]$ListId,        
        [Parameter(Mandatory = $true)]
        $Values
    )    
    begin {        
        # Create header
        $Header = @{       
            Authorization = $AuthToken['Authorization'] 
        }
        $resource = "https://graph.microsoft.com"
    }
    process {
        Write-progress -Activity "Create a new listItem in a list."        
        $Uri = "$resource/$ApiVersion/sites/$SiteId/lists/$ListId/items"
        
        try { 
            if ($Values){
                $Body = ConvertTo-Json $Values            
                $results = Invoke-NIHGraph -Method "POST" -URI $Uri -Headers $Header -Body $Body
            }            
        }
        catch {
            if ($_.exception.response.statuscode.value__ -eq 404) {
                Write-Progress -Activity 'Create a new listItem in a list.' -Completed
                Write-Warning -Message "Not found error while create a new listItem in a list." ; return
            }
            else { throw $_ ; return }
        }
        Write-Progress -Activity 'Create a new listItem in a list.' -Completed
        $results
    }
}

function Set-NIHListItem {
    [cmdletBinding()]
    param(
        [parameter(Mandatory = $true)]
        [hashtable]$AuthToken,                 
        [ValidateSet('beta', 'v1.0')] #Graph API version
        [string]$ApiVersion = 'v1.0',
        [Parameter(Mandatory = $true)]
        [string]$SiteId,
        [Parameter(Mandatory = $true)]
        [string]$ListId,
        [Parameter(Mandatory = $true)]
        [string]$ItemId,
        [Parameter(Mandatory = $true)]
        $Values
    )    
    begin {        
        # Create header
        $Header = @{       
            Authorization = $AuthToken['Authorization'] 
        }
        $resource = "https://graph.microsoft.com"
    }
    process {
        Write-progress -Activity "Update the properties on a listItem."

        $Uri = "$resource/$ApiVersion/sites/$SiteId/lists/$ListId/items/$ItemId/fields"
        
        try { 
            if ($Values){
                $Body = ConvertTo-Json $Values            
                $results = Invoke-NIHGraph -Method "PATCH" -URI $Uri -Headers $Header -Body $Body
            }            
        }
        catch {
            if ($_.exception.response.statuscode.value__ -eq 404) {
                Write-Progress -Activity 'Update the properties on a listItem.' -Completed
                Write-Warning -Message "Not found error while update the properties on a listItem." ; return
            }
            else { throw $_ ; return }
        }
        Write-Progress -Activity 'Update the properties on a listItem.' -Completed
        $results
    }
}
#endregion