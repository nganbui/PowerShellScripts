$0 = $MyInvocation.MyCommand.Definition
$dp0 = [System.IO.Path]::GetDirectoryName($0)
$currentFolder = $dp0.Substring($dp0.LastIndexOf('\')+1)
$script:RootDir = Resolve-Path "$dp0\.." 
$logFileName = $MyInvocation.MyCommand.Name.Substring(0, $MyInvocation.MyCommand.Name.IndexOf('.'))
$logFileName = "$currentFolder\$logFileName"
#Include core
."$script:RootDir\Common\Lib\LibCore.ps1"
#Include dependent functionality
."$script:RootDir\Common\Lib\LibO365.ps1"
."$script:RootDir\Common\Lib\LibSPOSites.ps1"
."$script:RootDir\Common\Lib\LibO365Users.ps1"
."$script:RootDir\Common\Lib\GraphAPILibO365Groups.ps1"
."$script:RootDir\Common\Lib\LibCache.ps1"

Set-TenantVars
Set-AzureAppVars
Set-DataFile  
$siteUrl = 'https://nih-my.sharepoint.com/personal/alittle_nih_gov'
$tenantContext = ConnectPnpOnlineOAuth -TenantId $script:TenantId -ClientId $script:appIdOperationSupport -Thumbprint $script:appThumbprintOperationSupport -Url $script:SPOAdminCenterURL
$siteDetailed = Get-PnPTenantSite -Url $siteUrl | Select *
$siteDetailed  
$siteContext = ConnectPnpOnlineOAuth -TenantId $script:TenantId -ClientId $script:appIdOperationSupport -Thumbprint $script:appThumbprintOperationSupport -Url $siteURL
$context = Get-PnPContext
$Web = $context.Web 
$context.Load($Web) 
$List = $context.Web.Lists.GetByTitle("Documents") 
$context.Load($List) 
$context.ExecuteQuery()
$Web.Description

Write-Host "Items Found :" $List.ItemCount 
Write-Host "Items Found :" $list.LastItemUserModifiedDate
Write-Host "Items Found :" $Web.LastItemModifiedDate
$siteAdmins = (Get-PnPSiteCollectionAdmin | ? {$_.Email -ne '' -and $_.Email.ToLower() -notlike 'spoadm*'}).Email -join ";"
$siteAdmins
$Web.Created.ToString()

#$ps = Get-PnPTenantSite -Connection $tenantContext  -IncludeOneDriveSites  | Where-Object {($_.Url -like '*-my.sharepoint.com/personal*') -and ($_.Url -notlike '*-my-admin.sharepoint.com*')} | select *                              
#$siteDetailed = Get-PnPTenantSite -Url $siteUrl | Select *
#$siteDetailed                
if ($siteDetailed.Status -eq 'Active' -and $siteDetailed.LockState -eq 'Unlock'){
    #LogWrite -Message "Connecting to SharePoint Online $siteUrl"
    #ConnectPnpOnlineOAuth -TenantId $script:TenantId -ClientId $script:appIdAdminPortal -Thumbprint $script:appThumbprintAdminPortal -Url $siteUrl
    #$siteAdmins = (Get-PnPSiteCollectionAdmin | ? {$_.Email -ne '' -and $_.Email.ToLower() -notlike 'spoadm*'}).Email -join ";"                    
    #$siteAdmins = (Get-PnPSiteCollectionAdmin | ? {$_.Email -ne ''}).Email -join ";"                    
    #$sitesObj.SecondarySCA = $siteAdmins
    #$web = Get-PnPWeb -Includes Created
    #$sitesObj.Created = $web.Created
}