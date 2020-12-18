<#
$TenantName = "nih.onmicrosoft.com"
$AppId = "9624e216-9e73-4513-9251-4d4382950420"
$Certificate = Get-Item Cert:\LocalMachine\My\* | Where-Object { $_.Subject -ieq "CN=SPO-M365OperationSupport.nih.sharepoint.com" }
$Certificate

$token = Connect-NIHO365GraphWithCert -TenantName $tenantname -AppId $appId -Certificate $Certificate
$token
#>

$TenantName = "nih.onmicrosoft.com"
$AppId = "497e07ac-d6f7-4d40-9d70-54ebb507ef39"
$Certificate = Get-Item Cert:\LocalMachine\My\* | Where-Object { $_.Subject -ieq "CN=SPO-Sync Operations" }
$Certificate

$token = Connect-NIHO365GraphWithCert -TenantName $tenantname -AppId $appId -Certificate $Certificate
$token
<#
$teamAllMembersResponse = Get-NIHTeamMembers -AuthToken $token -Id "535a023f-8426-4a2d-935c-abe4025c5661" # -Role Member
$grpOwnsers = Get-NIHO365GroupMembers -AuthToken $token -Id "535a023f-8426-4a2d-935c-abe4025c5661"
$teamAllMembersResponse.count
$grpOwnsers.count
#>
#$user = Get-NIHO365User -AuthToken $token -UserID buint@nih.gov -Select createdDateTime
#$user


if ($token) {    
    $allLicenses = Get-NIHLicenses -AuthToken $token
    [System.Collections.ArrayList]$Licenses = @()
    $allLicenses.ForEach({
        $null = $Licenses.Add([PSCustomObject]@{
            skuId                     = $_.skuId
            skuPartNumber             = $_.skuPartNumber
            servicePlans              = $_.servicePlans
        })
   })    
   #$assignedPlans_ = ($allLicenses.servicePlans | ? {$_.provisioningStatus -eq "Success"}).servicePlanName -join ";"    
   #$Licenses
  <# $allUsers = Get-NIHO365Users -AuthToken $token -Select id,userPrincipalName,givenName,surname,displayName,mail,mailNickname,otherMails,businessPhones,mobilePhone,jobTitle,employeeId,
                                                                            userType,accountEnabled,createdDateTime,deletedDateTime,department,companyName,proxyAddresses,
                                                                            assignedLicenses,assignedPlans,provisionedPlans,
                                                                            officeLocation,streetAddress,city,state,country,postalCode,
                                                                            creationType,externalUserState,externalUserStateChangeDateTime,isResourceAccount,
                                                                            lastPasswordChangeDateTime,passwordPolicies,
                                                                            onPremisesDomainName,onPremisesLastSyncDateTime,onPremisesUserPrincipalName,
                                                                            onPremisesSamAccountName,onPremisesSecurityIdentifier,onPremisesSyncEnabled,onPremisesExtensionAttributes
    #>                                                                        
   #$t = $Licenses | Where-Object {$_.servicePlans.servicePlanId -match "8e3eb3bd-bc99-4221-81b8-8b8bc882e128"}
   
   $5users = $allUsers | Select-Object -First 1000
   #$5users
   $allUsers.ForEach({
        if ($_.assignedPlans.length -gt 0){
            $UserId = $_.Id
            $userPrincipalName = $_.userPrincipalName
            $assignedLicenses_ = $null
            $assignedPlans_ = $null
            $assignedPlans = @($_.assignedPlans | ? {$_.CapabilityStatus -eq "Enabled"})
            
            if ($_.assignedLicenses.length -gt 0) { 
                $isLicensed = 1                                
                $l = @($_.assignedLicenses.skuId)
                $l.ForEach({                    
                    $licenseId = $_
                    #Write-Host $licenseId                           
                    $Licenses.ForEach({
                        $skuId = $_.skuId
                        if ($skuId -eq $licenseId){
                            $skuNumber = $_.skuPartNumber
                            $servicePlans = $_.servicePlans
                            $assignedLicenses_+=$skuNumber + ";"                            
                        }

                    })                    
                    $assignedPlans.ForEach({
                        $id = $_.servicePlanId
                        $servicePlanName = ($servicePlans | ? {$_.servicePlanId -eq $id}).servicePlanName
                        if ($null -ne $servicePlanName){
                            $assignedPlans_+= $servicePlanName + ";"
                            }
                   
                    })
                                        
                }) 
            }
        Write-host $userPrincipalName       
        #Write-host $assignedPlans
        #Write-host $assignedPlans.Length
        Write-host $assignedPlans_
        Write-host $assignedLicenses_
        }

    })    
} 

