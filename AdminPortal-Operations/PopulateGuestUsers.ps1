$0 = $MyInvocation.MyCommand.Definition
$dp0 = [System.IO.Path]::GetDirectoryName($0)
$currentFolder = $dp0.Substring($dp0.LastIndexOf('\')+1)
$script:RootDir = Resolve-Path "$dp0\.." 
$logFileName = $MyInvocation.MyCommand.Name.Substring(0, $MyInvocation.MyCommand.Name.IndexOf('.'))
$logFileName = "$currentFolder\$logFileName"
#$guestReport = "O365"
$usageReport = "UsageReports"
$inputReport = "Input"
$guestReportFile = "GuestActivity.csv"
#Include core
."$script:RootDir\Common\Lib\LibCore.ps1"
."$script:RootDir\Common\Lib\LibO365.ps1"

Function SyncGuestUsersToCache {    
    Set-TenantVars
    Set-AzureAppVars
    Set-DataFile

    #--Create a folder UsageReports under Data if any        
    $date = Get-Date
    $year = $date.Year
    $month = $date.AddMonths(-1).Month
    $monthName = (Get-Culture).DateTimeFormat.GetMonthName($month)
    $reportFolder = "$($script:CacheDataPath)\$($usageReport)\$($monthName)\$($inputReport)"
    Create-Directory $reportFolder   

    # create a new DateTime object set to the first day of a given month and year
    $StartDate = Get-Date -Year $year -Month $month -Day 1 -Hour 0 -Minute 0 -Second 0 -Millisecond 0
    # add a month and subtract the smallest possible time unit
    $EndDate = ($StartDate).AddMonths(1).AddTicks(-1)   
    #--
    #$EndDate = (Get-Date).AddDays(1);
    #$StartDate = (Get-Date).AddDays(-90);
    #$StartDate = "7/1/2020"
    #$EndDate = "7/31/2020"

    $ReportGuest = [System.Collections.Generic.List[Object]]::new()
    
    LogWrite -Message "Getting acces token using Graph API..."
    #Invoke-GraphAPIAuthTokenCheck
    $cert = Get-Item Cert:\\LocalMachine\\My\* | Where-Object { $_.Subject -ieq "CN=$($script:appCertAdminPortalOperation)" }    
    $script:authToken = Connect-NIHO365GraphWithCert -TenantName $script:TenantName -AppId $script:appIdAdminPortalOperation -Certificate $cert

    if ($script:authToken) {
        LogWrite -Message  "Retrieving Guest Users starting..."
        $GuestUsers = Get-NIHO365GuestUsers -AuthToken $script:authToken
        $totalGuests = @($GuestUsers).Length
        LogWrite -Message  "Connect to [EXO V2] to get M365 groups that guest are member of..."
        Connect-ExchangeOnline -AppId $script:appIdEXOV2 -Organization $script:TenantName -CertificateThumbprint $script:appThumbprintEXOV2 
                      
        if ($GuestUsers) {
            LogWrite -Message  "Total guests: $totalGuests"
            $j = 1
            ForEach ($Guest in $GuestUsers) {    
                #$AADAccountAge = ($Guest.createdDateTime | New-TimeSpan).Days   
                $Name =  $Guest.DisplayName  
                #Write-Host "Processing" $Guest.DisplayName
                LogWrite -Message  "($j/$totalGuests): Processing $($Name) ..."
                $i = 0; 
                $GroupNames = $Null
                $GroupIds = $Null
                # Find what Office 365 Groups the guest belongs to... if any
                try {
                    $DN = (Get-Recipient -Identity $Guest.UserPrincipalName).DistinguishedName
                    try{    
                        $GuestGroups = (Get-Recipient -Filter "Members -eq '$Dn'" -RecipientTypeDetails GroupMailbox | Select DisplayName, ExternalDirectoryObjectId)
                    }
                    catch{
                        LogWrite -Level ERROR  "-Getting guest groups error: $_ "
                        continue;
                    }
                    #$GuestGroups = (Get-EXORecipient -Filter "Members -eq '$Dn'" -RecipientTypeDetails GroupMailbox | Select DisplayName, ExternalDirectoryObjectId)
                    If ($GuestGroups -ne $Null) {
                        ForEach ($G in $GuestGroups) { 
                        If ($i -eq 0) { 
                            $GroupNames = $G.DisplayName;
                            $GroupIds = $G.ExternalDirectoryObjectId
                            $i++ }
                        Else 
                        {
                            $GroupNames = $GroupNames + "; " + $G.DisplayName
                            $GroupIds = $GroupIds + "; " + $G.ExternalDirectoryObjectId 
                         }
                    }}
                    
                    $GuestAudit = Search-UnifiedAuditLog -UserIds $Guest.UserPrincipalName -StartDate $StartDate -EndDate $EndDate -ResultSize 5000
                    if ($GuestAudit -ne $null) {
                        $ConvertAudit = $GuestAudit | Select-Object -ExpandProperty AuditData | ConvertFrom-Json
                        $ConvertAudit | Select-Object -Last 1 CreationTime,UserId,Operation,Workload,ObjectID,SiteUrl,SourceFileName,ClientIP,UserAgent
            
                        $lastLoginDate = $ConvertAudit[0].CreationTime
                        $operation = $ConvertAudit[0].Operation
                        $siteUrl = $ConvertAudit[0].SiteUrl
                        $sourceFileName = $ConvertAudit[0].SourceFileName
                        $clientIP = $ConvertAudit[0].ClientIP
                        $userAgent = $ConvertAudit[0].UserAgent
                        LogWrite -Message "$($Guest.UserPrincipalName) - $($lastLoginDate) - $($operation) - $($siteUrl) - $($sourceFileName)"
                    }

                    $ReportLine = [PSCustomObject]@{
                        'Report Refresh Date' = $EndDate
                        UPN     = $Guest.UserPrincipalName
                        Email   = $Guest.mail
                        Name    = $Guest.DisplayName
                        'Last Password Change' = $Guest.lastPasswordChangeDateTime
                        'Creation Type' = $Guest.creationType
                        'User State' = $Guest.externalUserState
                        'StateChangeDateTime' = $Guest.externalUserStateChangeDateTime                        
                         Created = $Guest.createdDateTime
                        'Last Login Date'     = $lastLoginDate
                        'Operation'     = $operation
                        'SiteUrl'     = $siteUrl
                        'SourceFileName'     = $sourceFileName
                        GroupIds = $GroupIds
                        GroupNames  = $GroupNames   
                        'ClientIP'     = $clientIP   
                        'UserAgent'     = $userAgent                         
                         'DistinguishedName'   = $DN
                        'Report Period' = 30
                        }      
                    $ReportGuest.Add($ReportLine) 
                    $j++
                }
                catch {
                    LogWrite -Level ERROR  "-Processing error: $_ "
                }
            }

            LogWrite -Message  "Disconnect [EXO V2]"
            Disconnect-ExchangeOnline -Confirm:$false
            LogWrite -Message  "Export guest users to .csv file"
            $ReportGuest | Sort Name | Export-CSV -NoTypeInformation "$($reportFolder)\$($guestReportFile)"
                        
            }

        LogWrite -Message  "Retrieving Guest Users completed."
    }
   
}

Try {
    #log file path
    Set-LogFile -logFileName $logFileName
    $script:StartTimeDailyCache = Get-Date -Format "yyyy-MM-dd HH:mm:ss"           
    LogWrite -Message "----------------------- [Populate Guest Users with their M365 Groups] Execution Started -----------------------"

    #Populate Guest users to the Cache
    SyncGuestUsersToCache 
    <#Set-TenantVars
    Set-AzureAppVars
    Set-DataFile
    Connect-ExchangeOnline -AppId $script:appIdEXOV2 -Organization $script:TenantName -CertificateThumbprint $script:appThumbprintEXOV2     
    $mailbox = Get-SiteMailbox -BypassOwnerCheck -ResultSize Unlimited | Select *
    $mailbox | Export-CSV -NoTypeInformation "D:\Download\Mailbox.csv"
    #>
    $script:EndTimeDailyCache = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    
    LogWrite -Message "[Populate Guest Users with their M365 Groups] Start Time: $($script:StartTimeDailyCache)"
    LogWrite -Message "[Populate Guest Users with their M365 Groups] End Time:   $($script:EndTimeDailyCache)"
    LogWrite -Message  "----------------------- [Populate Guest Users with their M365 Groups] Execution Ended ------------------------"    
    
}
Catch [Exception] {
    LogWrite -Level ERROR "-Unexpected Error: $_ "
}
Finally {
    LogWrite -Message  "----------------------- [Populate Guest Users with their M365 Groups] Completed ------------------------"
}
