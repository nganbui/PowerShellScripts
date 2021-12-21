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
<#
    .DESCRIPTION
    Using Communications API to get PSTN calls
    Application	Permission: CallRecords.Read.All

#>
function Get-NIHActivityCommunication {
    [cmdletBinding()]
    param(
        [parameter(Mandatory = $true)]
        [hashtable]$AuthToken,
        [ValidateSet('beta', 'v1.0')] #Graph API version
        [string]$ApiVersion = 'v1.0',  
        [parameter(Mandatory = $true)]
        [string]$fromTime,
        [parameter(Mandatory = $true)]
        [string]$toTime,
        [parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [String]$Report

    )
    begin {        
        # Create header
        $Header = @{       
            Authorization = $AuthToken['Authorization']
        } 
    }
    process {
        Write-Progress -Activity 'Getting log of PSTN calls as a collection of pstnCallLog.Endpoint: callRecords/getPstnCalls'
        $Uri = "https://graph.microsoft.com/$($ApiVersion)/communications/$Report(fromDateTime=$fromTime,toDateTime=$toTime)"; 

        <#
        if ($fromTime -and $toTime) {             
            $Uri = $Uri + "(fromDateTime=$fromTime,toDateTime=$toTime)"  
        }
        #>
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
                Write-Progress -Activity 'Getting log of PSTN calls as a collection of pstnCallLog' -Completed
                Write-Warning -Message "Error while Getting log of PSTN calls as a collection of pstnCallLog" ; return
            }
            else { throw $_ ; return }
        }
        Write-Progress -Activity 'Getting log of PSTN calls as a collection of pstnCallLog' -Completed
        $objectCollection
    }   
}