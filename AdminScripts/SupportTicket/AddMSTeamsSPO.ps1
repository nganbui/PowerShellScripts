#--Connect existing site to M365 Group
#--Require: GA role and site admin to perform follwing tasks (User Admin/GA role to enforce group naming policy)
# Step 1: Convert SPO to M365 Group (skip Step 1 if the site is connected M365 Group
#-- After you change the $siteURL, select line 5 to line 11 and Run selection (F8)
$adminUrl = "https://nih-admin.sharepoint.com"
$SiteUrl = "https://nih.sharepoint.com/sites/NIAID-OWERTeams"
Connect-PnPOnline -Url $adminUrl -Interactive
$Alias = "NIAID-OWERTeams"
$DisplayName = "NIAID OWERTeams"
Add-PnPMicrosoft365GroupToSite -Url $SiteUrl -Alias $Alias -DisplayName $DisplayName -KeepOldHomePage:$true
Disconnect-PnPOnline

#---M365 Group convert to MS Teams
#---promote to MS Teams
#-- Step 2: Convert M365 Group to MS Teams
#-- After the step 1 completed, change the $groupId, select line 17 to 24 and Run selection(F8)
$TenantName = "nih.onmicrosoft.com"
$AppId = "9624e216-9e73-4513-9251-4d4382950420"
$Thumbprint = "1C9696EB9152228A42DAEB5C7075699795311662"
$CertName   = "SPO-M365OperationSupport.nih.sharepoint.com"
$Certificate = Get-Item Cert:\LocalMachine\My\* | Where-Object { $_.Thumbprint -ieq "$Thumbprint" } 
$token = Connect-NIHO365GraphWithCert -TenantName $TenantName -AppId $AppId -Certificate $Certificate
#$groupId = "b6b8e2fc-b342-4a17-81e5-f50024c0e7db"
#Promote-NIHGroupToTeam -AuthToken $token -Id $groupId

#-- Step 3: Add business user as an owner to newly created team
#$m = "aleavey@nih.gov"
#Add-NIHTeamMember -AuthToken $token -Group $groupId -Member $m -AsOwner

# Step 4: Verify the MS teams after conversion
#$team = Get-NIHTeam -AuthToken $token -Id $groupId
#$teamOwners = Get-NIHTeamMembers -AuthToken $token -Id $groupId