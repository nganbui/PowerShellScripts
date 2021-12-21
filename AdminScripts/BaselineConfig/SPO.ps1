<#$AppId      = '497e07ac-d6f7-4d40-9d70-54ebb507ef39'        
$Thumbprint = '7BA7CBA81EDC57BF8446C549294148FB8490AD5B' 
$TenantId   = "14b77578-9773-42d5-8507-251ca2dc2b06"
$TenantName = "nih.onmicrosoft.com"
$CertName   = "SPO-Sync Operations"
#>

$AdminCenterUrl    = "https://nih-admin.sharepoint.com"
$RootSiteUrl       = "https://nih.sharepoint.com"
$ReportOutput = "D:\Scripting\O365DevOps\AdminScripts\BaselineConfig"
$fileName = "M365-BaselineObject.xlsx"
$baselineConfigFile = "$ReportOutput\$fileName"

$SPM365BoxTeamID = "c7d703c0-4a4d-407f-a4f3-c2d0d89a2004" # CIT SP-M365-Box Team
$spoadmId = "36804259-b5c9-41be-a383-f52f3bd5e840" # SPOADMSVC@nih.gov
$worksheet = [ordered]@{
    Site = "Site";
    Group = "Group";
    Team = "Team";
    User = "User";
}
$ExcelParams = @{
        Path      = $baselineConfigFile
        Show      = $false
        Verbose   = $true
        ClearSheet = $true
        AutoSize = $true
        }


Function ReadBaselineConfig{
    param(
        [Parameter(Mandatory=$true)][string]$Path,
        [Parameter(Mandatory=$true)][string]$WorkSheetName
    )
    $data = @()
    if (Test-Path $Path) {
        try{ 
           $data = Import-Excel -Path $Path -WorksheetName $WorkSheetName
          }
        catch{
            $data = @()
        }
    }
    return $data

}

Function ReadExelFile{
    $Url =  "https://nih.sharepoint.com/sites/spoadm"
    $clientId = '9624e216-9e73-4513-9251-4d4382950420'   
    $TenantId = 'nih.onmicrosoft.com'
    $Thumbprint = "1C9696EB9152228A42DAEB5C7075699795311662"
    $FileURL = "$Url/Docs/M365 Services/M365 Configurations/$fileName"
    $cert = Get-ChildItem "cert:\LocalMachine\My" | Where-Object {$_.Thumbprint -eq $Thumbprint}
    $conn = Connect-PnPOnline -Url $Url -ClientId $clientId -Thumbprint $Thumbprint -Tenant $TenantId -ReturnConnection
    
    $context = Get-PnPContext                
    $Web = $context.Web 
    #$context.Load($Web)               
    $File = $web.GetFileByUrl($FileURL)    
    $context.Load($File)
    $context.ExecuteQuery()
    Write-host "File Size:" ($File.Length/1KB)
    Write-host $File.ServerRelativeURL
    if ($File.Length){
        Get-PnPFile -Url $File.ServerRelativeURL -Path $ReportOutput -FileName $fileName -AsFile -Force
    }
    $context.Dispose()
    
    
    
}

Function UploadFileToDocLib{
    $Url =  "https://nih.sharepoint.com/sites/spoadm"
    $clientId = '9624e216-9e73-4513-9251-4d4382950420'   
    $TenantId = 'nih.onmicrosoft.com'
    $Thumbprint = "1C9696EB9152228A42DAEB5C7075699795311662"    

    $cert = Get-ChildItem "cert:\LocalMachine\My" | Where-Object {$_.Thumbprint -eq $Thumbprint}
    $conn = Connect-PnPOnline -Url $Url -ClientId $clientId -Thumbprint $Thumbprint -Tenant $TenantId -ReturnConnection    
    $context = Get-PnPContext                
    $Web = $context.Web
    Add-PnPFile -Path $baselineConfigFile -Folder "Docs/M365 Services/M365 Configurations"
    $context.Dispose()
    
}

Function UpdateSPOBaseline {        
    $siteProps = @(Get-SPOSite -Identity $RootSiteUrl -Limit All | Get-Member -MemberType Property | Select Name)
    $currentProps = $siteProps.Name
    $site = @(ReadBaselineConfig -Path $baselineConfigFile -WorkSheetName $worksheet.Site)
    if (!$site){
        $siteProps | Export-Excel @ExcelParams -TableName $worksheet.Site -WorksheetName $worksheet.Site
        return
    }
    $props = $site.Name
    # => Exists in Difference only (currentProps)
    # <= Exists in Reference only (spreadsheet)
    $newProps = @(Compare-Object -ReferenceObject $props -DifferenceObject $currentProps -IncludeEqual -PassThru | Where-Object {$_.SideIndicator -eq "=>" })
    $archiveProps = @(Compare-Object -ReferenceObject $props -DifferenceObject $currentProps -IncludeEqual -PassThru | Where-Object {$_.SideIndicator -eq "<=" })
    $existingProps = @(Compare-Object -ReferenceObject $props -DifferenceObject $currentProps -IncludeEqual -PassThru | Where-Object {$_.SideIndicator -eq "==" })

    [System.Collections.ArrayList]$ConditionalFormat = @() 
    if ($newProps.Count -gt 0){    
        $newProps | % { 
                        $t = New-ConditionalText -Text $_ -BackgroundColor Green -ConditionalTextColor Cyan
                        $ConditionalFormat.Add($t)}
    
    }
    if ($archiveProps.Count -gt 0){    
        $archiveProps | % {
                        $siteProps+=[PSCustomObject]@{Name=$_}
                        $t = New-ConditionalText -Text $_
                        $ConditionalFormat.Add($t)
                        }    
    }
    <#
    $archive = [PSCustomObject]@{
            Name   = 'IsDeleted'
            Status   = 'Archived'
    }#>
    #$Array | Update-FirstObjectProperties | Export-Excel @ExcelParams -ClearSheet -WorksheetName Numbers
    <#
    $ConditionalFormat =@(    
        New-ConditionalText -Text IsTeamsChannelConnected -BackgroundColor Blue -ConditionalTextColor Yellow
        New-ConditionalText -Text IsTeamsConnected -BackgroundColor Blue -ConditionalTextColor Yellow
    )
    #>
    $siteProps | Export-Excel @ExcelParams -TableName $worksheet.Site -WorksheetName $worksheet.Site -ConditionalText $ConditionalFormat    
}

Function UpdateTeamBaseline{        
    $teamProps = @(Get-Team -GroupId $SPM365BoxTeamID | Get-Member -MemberType Property | Select Name)    
    $currentProps = $teamProps.Name
    $team = @(ReadBaselineConfig -Path $baselineConfigFile -WorkSheetName $worksheet.Team)
    if (!$team){
        $teamProps | Export-Excel @ExcelParams -TableName $worksheet.Team -WorksheetName $worksheet.Team
        return
    }
    $props = $team.Name
    # => Exists in Difference only (currentProps)
    # <= Exists in Reference only (spreadsheet)
    $newProps = @(Compare-Object -ReferenceObject $props -DifferenceObject $currentProps -IncludeEqual -PassThru | Where-Object {$_.SideIndicator -eq "=>" })
    $archiveProps = @(Compare-Object -ReferenceObject $props -DifferenceObject $currentProps -IncludeEqual -PassThru | Where-Object {$_.SideIndicator -eq "<=" })
    $existingProps = @(Compare-Object -ReferenceObject $props -DifferenceObject $currentProps -IncludeEqual -PassThru | Where-Object {$_.SideIndicator -eq "==" })
    
    [System.Collections.ArrayList]$ConditionalFormat = @() 
    if ($newProps.Count -gt 0){    
        $newProps | % { 
                        $t = New-ConditionalText -Text $_ -BackgroundColor Green -ConditionalTextColor Cyan
                        $ConditionalFormat.Add($t)}
    
    }
    if ($archiveProps.Count -gt 0){    
        $archiveProps | % {
                        $teamProps+=[PSCustomObject]@{Name=$_}
                        $t = New-ConditionalText -Text $_
                        $ConditionalFormat.Add($t)
                        }    
    }    
    $teamProps | Export-Excel @ExcelParams -TableName $worksheet.Team -WorksheetName $worksheet.Team -ConditionalText $ConditionalFormat

}

Function UpdateGroupBaseline{
    $groupProps = @(Get-AzureADGroup -ObjectId $SPM365BoxTeamID | Get-Member -MemberType Property | Select Name)
    $currentProps = $groupProps.Name
    $group = @(ReadBaselineConfig -Path $baselineConfigFile -WorkSheetName $worksheet.Group)
    if (!$group){
        $groupProps | Export-Excel @ExcelParams -TableName $worksheet.Group -WorksheetName $worksheet.Group
        return
    }
    $props = $group.Name
    # => Exists in Difference only (currentProps)
    # <= Exists in Reference only (spreadsheet)
    $newProps = @(Compare-Object -ReferenceObject $props -DifferenceObject $currentProps -IncludeEqual -PassThru | Where-Object {$_.SideIndicator -eq "=>" })
    $archiveProps = @(Compare-Object -ReferenceObject $props -DifferenceObject $currentProps -IncludeEqual -PassThru | Where-Object {$_.SideIndicator -eq "<=" })
    $existingProps = @(Compare-Object -ReferenceObject $props -DifferenceObject $currentProps -IncludeEqual -PassThru | Where-Object {$_.SideIndicator -eq "==" })
    
    [System.Collections.ArrayList]$ConditionalFormat = @() 
    if ($newProps.Count -gt 0){    
        $newProps | % { 
                        $t = New-ConditionalText -Text $_ -BackgroundColor Green -ConditionalTextColor Cyan
                        $ConditionalFormat.Add($t)}
    
    }
    if ($archiveProps.Count -gt 0){    
        $archiveProps | % {
                        $groupProps+=[PSCustomObject]@{Name=$_}
                        $t = New-ConditionalText -Text $_
                        $ConditionalFormat.Add($t)
                        }    
    }    
    $groupProps | Export-Excel @ExcelParams -TableName $worksheet.Group -WorksheetName $worksheet.Group -ConditionalText $ConditionalFormat 
}

Function UpdateUserBaseline{    
    # Connect-AzureAD
    $userProps = @(Get-AzureADUser -ObjectId $spoadmId | Get-Member -MemberType Property | Select Name)
    $currentProps = $userProps.Name
    $user = @(ReadBaselineConfig -Path $baselineConfigFile -WorkSheetName $worksheet.User)
    if (!$user){
        $userProps | Export-Excel @ExcelParams -TableName $worksheet.User -WorksheetName $worksheet.User
        return
    }
    $props = $user.Name
    # => Exists in Difference only (currentProps)
    # <= Exists in Reference only (spreadsheet)
    $newProps = @(Compare-Object -ReferenceObject $props -DifferenceObject $currentProps -IncludeEqual -PassThru | Where-Object {$_.SideIndicator -eq "=>" })
    $archiveProps = @(Compare-Object -ReferenceObject $props -DifferenceObject $currentProps -IncludeEqual -PassThru | Where-Object {$_.SideIndicator -eq "<=" })
    $existingProps = @(Compare-Object -ReferenceObject $props -DifferenceObject $currentProps -IncludeEqual -PassThru | Where-Object {$_.SideIndicator -eq "==" })
    
    [System.Collections.ArrayList]$ConditionalFormat = @() 
    if ($newProps.Count -gt 0){    
        $newProps | % { 
                        $t = New-ConditionalText -Text $_ -BackgroundColor Green -ConditionalTextColor Yellow
                        $ConditionalFormat.Add($t)}
    
    }
    if ($archiveProps.Count -gt 0){    
        $archiveProps | % {
                        $userProps+=[PSCustomObject]@{Name=$_}
                        $t = New-ConditionalText -Text $_
                        $ConditionalFormat.Add($t)
                        }    
    }    
    $userProps | Export-Excel @ExcelParams -TableName $worksheet.User -WorksheetName $worksheet.User -ConditionalText $ConditionalFormat     
}

ReadExelFile
Disconnect-PnPOnline

Write-Host "=== Connecting to SharePoint Admin ==="
Connect-SPOService -Url $AdminCenterUrl
UpdateSPOBaseline
Write-Host "==== Disconnect SharePoint Admin ===="
Disconnect-SPOService

Write-Host "=== Connecting to Microsoft Teams Admin ==="
Connect-MicrosoftTeams
UpdateTeamBaseline
Write-Host "==== Disconnect Microsoft Teams Admin ===="
Disconnect-MicrosoftTeams

Write-Host "=== Connecting to Azure AD ==="
Connect-AzureAD
UpdateGroupBaseline
UpdateUserBaseline
Write-Host "==== Disconnect Microsoft Teams Admin ===="
Disconnect-AzureAD

Write-Host "=== Uploading to document library ==="
UploadFileToDocLib
Disconnect-PnPOnline
