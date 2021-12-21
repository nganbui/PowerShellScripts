#/auditLogs/directoryaudits
$TenantName = 'nih.onmicrosoft.com'
$CertName   = "SPO-Sync Operations"
$thumprint = "7BA7CBA81EDC57BF8446C549294148FB8490AD5B"
$clientId = '497e07ac-d6f7-4d40-9d70-54ebb507ef39' 

#$Certificate = Get-Item Cert:\LocalMachine\My\* | Where-Object { $_.Subject -ieq "CN=$CertName" }
#$token = Connect-NIHO365GraphWithCert -TenantName $TenantName -AppId $clientId -Certificate $Certificate 

$Certificate = Get-Item Cert:\LocalMachine\My\* | Where-Object { $_.Thumbprint -ieq "$thumprint" } 
$token = Connect-NIHO365GraphWithCert -TenantName $TenantName -AppId $clientId -Certificate $Certificate

$accessToken = $token["Authorization"]
#targetResources/any(t:t/displayName eq 'NIBIB ExtRED Team (O365)')
#$uri = "https://graph.microsoft.com/v1.0/auditLogs/directoryAudits?`$filter=targetResources/any(t:t/displayName eq 'NIBIB ExtRED Team (O365)')"
#$uri = "https://graph.microsoft.com/v1.0/auditLogs/directoryAudits?`$filter=targetResources/any(t:t/displayName eq 'NIBIB ExtRED Team (O365)')"
$uri = "https://graph.microsoft.com/v1.0/auditLogs/directoryAudits?`$filter=activityDateTime gt 2021-10-06"
$contentType = "application/x-www-form-urlencoded"
$headers = @{"Authorization"=$accessToken}
$auditlogs = Invoke-RestMethod -Uri $uri -Method Get -Headers $headers -ContentType $contentType
$auditlogs.value
#$auditlogs.value.initiatedBy.user
#$auditlogs.value.additionalDetails