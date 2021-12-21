$TenantName = "nih.onmicrosoft.com"
$AppId = "9624e216-9e73-4513-9251-4d4382950420"
$Thumbprint = "1C9696EB9152228A42DAEB5C7075699795311662"
$CertName   = "SPO-M365OperationSupport.nih.sharepoint.com"
$Certificate = Get-Item Cert:\LocalMachine\My\* | Where-Object { $_.Thumbprint -ieq "$Thumbprint" } 
$token = Connect-NIHO365GraphWithCert -TenantName $TenantName -AppId $AppId -Certificate $Certificate

$groupId = "7acbcda1-64e2-4ba9-8a29-40f8b0e67c50"
$Uri = "https://graph.microsoft.com/v1.0/groups/$groupId"
Invoke-NIHGraph -Method DELETE -URI $Uri -Headers $token         
Write-progress -Activity "Delete Group" -Completed

$grps = Get-NIHDeletedO365Groups -AuthToken $token
$g = $grps.Where({$_.id -eq $groupId})

