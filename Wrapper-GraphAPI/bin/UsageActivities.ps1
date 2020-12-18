function Get-NIHActivityReport {
    [cmdletBinding()]
    param(
        [parameter(Mandatory = $true)]
        [hashtable]$AuthToken,         
        [ValidateSet('D7', 'D30', 'D90', 'D180')] #Reporting Period in Days
        [string]$Period = 'D30',         
        [ValidateSet('beta', 'v1.0')] #Graph API version
        [string]$ApiVersion = 'v1.0',
        [parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [String]$Report
    )

    $ActivityResponse = $null

    if ($AuthToken['Authorization']) {
        # Create header
        $Header = @{       
            Authorization = $AuthToken['Authorization']
        }        
        # Fetch usage report        
        $Uri = "https://graph.microsoft.com/$($ApiVersion)/reports/$($Report)(period='$($Period)')"

        $ActivityResponse = Invoke-NIHGraph -Method "Get" -URI $Uri -Headers $Header       
        #clean data
        if ($ActivityResponse) {            
            $ActivityResponse = $ActivityResponse.replace("ï»¿", "") | ConvertFrom-Csv
        }
    }
    return $ActivityResponse
}