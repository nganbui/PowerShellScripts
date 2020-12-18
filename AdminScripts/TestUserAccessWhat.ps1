$0 = $MyInvocation.MyCommand.Definition
$dp0 = [System.IO.Path]::GetDirectoryName($0)
$currentFolder = $dp0.Substring($dp0.LastIndexOf('\')+1)
$script:RootDir = Resolve-Path "$dp0\.." 


#Include core
."$script:RootDir\Common\Lib\LibCore.ps1"

$token = Connect-NIHO365Graph -profilePath 'D:\Scripting\O365DevOps\Common\Config\PROFILE.psd1'

$userId = 'tadessewf@nih.gov'
$oauth2PermissionGrants =  Get-NIHO365UserDelegatedPermGrant -AuthToken $token -UserID $userId
$oauth2PermissionGrants
<#
$groups = Get-NIHO365UserMemberGroups -AuthToken $token -UserID $userId
foreach ($groupID in $groups){    
    $group = Get-NIHO365Group -AuthToken $token -Id $groupID | Select id, displayName
    $group
 }#>

$memberObjs = Get-NIHO365UserMemberObjects -AuthToken $token -UserID $userId
$memberObjs = @($memberObjs)
$memberObjs.Count
$objs = Get-NIHO365ObjectById -AuthToken $token -Ids $memberObjs
ExportCSV -DataSet $objs -FileName D:\Scripting\O365DevOps\Common\Data\Other\tadessewf-AccessM365.csv

<#
$token = Connect-NIHO365Graph -profilePath 'D:\Scripting\O365DevOps\Common\Config\PROFILE.psd1'
$groupID = '976924b8-cd2d-446d-9890-d758542c73b3'
$userId = 'pyatar@nih.gov'
Remove-NIHO365GroupMember -AuthToken $token -Group $groupID -Member $userId -AsOwner
#>		


