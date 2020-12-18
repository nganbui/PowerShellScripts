$AppId      = '9624e216-9e73-4513-9251-4d4382950420'        
$Thumbprint = '1C9696EB9152228A42DAEB5C7075699795311662' 
$Id   = "14b77578-9773-42d5-8507-251ca2dc2b06"
$Name = "nih.onmicrosoft.com"
$AdminCenterUrl    = "https://nih-admin.sharepoint.com"
$RootSiteUrl       = "https://nih.sharepoint.com"

$connection = Connect-PnPOnline -Tenant $Id -ClientId $AppId -Thumbprint $Thumbprint -Url $AdminCenterUrl -ReturnConnection

#$siteUrl = 'https://nih.sharepoint.com/sites/OD-RADx-rad-RFA-OD-20-020'
$siteUrl = 'https://nih.sharepoint.com/sites/NCI-F-2021RASsymposium'
$site = Get-PnPTenantSite -Url $siteUrl | select *
$site.SharingCapability

#Set-PnPTenantSite -Url $SiteUrl -SharingCapability ExternalUserSharingOnly

