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

Function ProcessRenameSiteAddress {
    param (
        [Parameter(Mandatory = $true)]
        [string] $oldURL,
        [Parameter(Mandatory = $true)] 
        [string] $newURL
    )
    #RenameSiteUrl -siteUrl $oldURL -newsiteUrl $newURL
    #Primary NBTeam3Test@citspdev.onmicrosoft.com
    
    try {        
        #rename site address
        #Connect to SharePoint Online
        LogWrite "Connecting to SharePoint Online..."        
        Connect-SPOService -Url $script:SPOAdminCenterURL
        $site = Get-SPOSite -Identity $oldURL -Detailed | Select *
        $site
        
        #Check if non-group sites
        if ($site.Template -ne 'GROUP#0'){            
            Start-SPOSiteRename -Identity $oldURL -NewSiteUrl $newURL -Confirm:$false
            $renameSiteDetails = Get-SPOSiteRenameState -Identity $oldURL
            while (($renameSiteDetails.State -ne 'Success') -or ($renameSiteDetails.State -ne 'Fail')) {
                Sleep 30
                $renameSiteDetails = Get-SPOSiteRenameState -Identity $oldURL
            }
            if ($renameSiteDetails.State -eq 'Fail') {
                LogWrite -Level ERROR "Rename SPO Site gets failed."
            } 
            else {
                LogWrite -Message "Rename SPO Site successfully."                
            } 
                
        }
        elseif($site.Template -eq 'GROUP#0'){
            $groupID = $site.GroupId
            Write-Host $groupID
            Connect-ExchangeOnline

            $groupDetail = Get-UnifiedGroup -Identity "$($site.GroupId)" | Select *
            $oldAlias = $groupDetail.alias
            $c = $newURL.LastIndexOf("/") + 1
            $newAlias = $newURL.Substring($c)

            Start-SPOSiteRename -Identity $oldURL -NewSiteUrl $newURL -Confirm:$false
            $renameSiteDetails = Get-SPOSiteRenameState -Identity $oldURL
            while (($renameSiteDetails.State -eq 'NotStarted') -or ($renameSiteDetails.State -eq 'Queued') -or ($renameSiteDetails.State -eq 'InProgress')) {
                Sleep 30
                $renameSiteDetails = Get-SPOSiteRenameState -Identity $oldURL
            }
            if ($renameSiteDetails.State -eq 'Fail') {
                LogWrite -Level ERROR "Rename SPO Site gets failed."
            } 
            else {
                LogWrite -Message "Rename SPO Site successfully."                
            }
            if ($renameSiteDetails.State -eq 'Success'){
                #Process change Group alias                

                Set-UnifiedGroup -Identity $groupID -EmailAddresses: @{Add = "$newAlias@citspdev.onmicrosoft.com" } -Alias $newAlias -ForceUpgrade
                #Set-UnifiedGroup -Identity $groupID -EmailAddresses: @{Add = "$newAlias@citspdev.mail.onmicrosoft.com" } -Alias $newAlias -ForceUpgrade

                #Promote alias as a primary SMTP address
                Set-UnifiedGroup -Identity $groupID -PrimarySmtpAddress "$newAlias@citspdev.onmicrosoft.com" -Alias $newAlias #group email - new primary address
                #Remove old ones
                Set-UnifiedGroup -Identity $groupID -EmailAddresses: @{Remove = "$oldAlias@citspdev.onmicrosoft.com" } #group email -primary email
                #Set-UnifiedGroup -Identity $groupID -EmailAddresses: @{Remove = "$oldAlias@citspdev.mail.onmicrosoft.com" } #alias
                #Set-UnifiedGroup -Identity $groupID -EmailAddresses: @{Remove = "$oldAlias@citspdev.onmicrosoft.com" } #alias
                LogWrite -Message "Completed changing group alias."
            }
       }   
    }
    catch {
        LogWrite -Level ERROR "Rename Site URL: $siteUrl - $($_)"
    }
    finally{
        Disconnect-SPOService
        Disconnect-ExchangeOnline -Confirm:$false
    }
    
}

Function ProcessDeleteOldSite {
    param (
        [Parameter(Mandatory = $true)] $renameRequests
    )    
}

Function RenameSiteUrl {
    param (
        [Parameter(Mandatory = $true)] $siteUrl,
        [Parameter(Mandatory = $true)] $newsiteUrl
    )
    #Connect_SPOService -Url $script:SPOAdminCenterURL -cred $script:o365AdminCredential
    #Start-SPOSiteRename -Identity "$($siteUrl)" -NewSiteUrl "$($newSiteUrl)" #-Confirm:$false 
    #Start-SPOSiteRename -Identity $siteUrl -NewSiteUrl $newsiteUrl
    try{
        LogWrite "Connecting to SharePoint Online..."
        #$cred = Get-Credential
        Connect-SPOService -Url $script:SPOAdminCenterURL   
        $site = Get-SPOSite -Identity $oldURL -Detailed | Select *
        $site

        Start-SPOSiteRename -Identity $siteUrl -NewSiteUrl $newsiteUrl -Confirm:$false
        $renameSiteDetails = Get-SPOSiteRenameState -Identity $siteUrl
        while (($renameSiteDetails.State -ne 'Success') -or ($renameSiteDetails.State -ne 'Fail')) {
            Sleep 30;
            $renameSiteDetails = Get-SPOSiteRenameState -Identity $siteUrl
        }
        if ($renameSiteDetails.State -eq 'Fail') {
            LogWrite -Level ERROR "Rename SPO Site gets failed."
        } else {
            LogWrite -Message "Rename SPO Site successfully."
        } 
    }
    catch {
        LogWrite -Level ERROR "Rename Site URL: $siteUrl - $($_)"
    }
    finally{
        Disconnect-SPOService
    }
}

Function UpdateGroupAlias {
    param (
        [Parameter(Mandatory = $true)] $groupID
    )
}

Function Test-Url {
    [CmdletBinding()]
 
    param (
        [Parameter(Mandatory = $true)]
        [String] $Url
    )
 
    Process {
        if ([system.uri]::IsWellFormedUriString($Url, [System.UriKind]::Absolute)) {
            $true
        }
        else {
            $false
        }
    }
}


Try {
    #log file path
    Set-TenantVars
    Set-AzureAppVars
    Set-LogFile -logFileName $logFileName
    $script:StartTimeM365GroupsDailyCache = Get-Date -Format "yyyy-MM-dd HH:mm:ss"           
    LogWrite -Message "----------------------- [Changing Site Address] Execution Started -----------------------"

    #$oldUrl = "https://citspdev.sharepoint.com/sites/NBTeam3Test"
    #$newUrl = "https://citspdev.sharepoint.com/sites/NBTeam3Test-123"

    #$oldUrl = "https://citspdev.sharepoint.com/sites/CIT-NBTeams-08"
    #$newUrl = "https://citspdev.sharepoint.com/sites/CIT-NBTeams-08-123"

    #$oldUrl = "https://citspdev.sharepoint.com/sites/CIT-NBTeams-04"
    #$newUrl = "https://citspdev.sharepoint.com/sites/CIT-NBTeams-04-123"

    <#$oldUrl = "https://citspdev.sharepoint.com/sites/NB-CIT-Group01"
    $newUrl = "https://citspdev.sharepoint.com/sites/NB-CIT-Group01-123"
    $groupID = "7d404907-4c44-4639-a2f3-e2b353ced2d8"
    $oldAlias = "NB-CIT-Group01"
    $newAlias = "NB-CIT-Group01-123"
    Connect-ExchangeOnline

    Set-UnifiedGroup -Identity "$($groupID)" -EmailAddresses: @{Add = "$newAlias@citspdev.onmicrosoft.com" } -Alias $newAlias -ForceUpgrade
    #Promote alias as a primary SMTP address
    Set-UnifiedGroup -Identity $groupID -PrimarySmtpAddress "$newAlias@citspdev.onmicrosoft.com" -Alias $newAlias #group email - new primary address
    #Remove old ones
    Set-UnifiedGroup -Identity $groupID -EmailAddresses: @{Remove = "$oldAlias@citspdev.onmicrosoft.com" } #group email -primary email
     #>  
         
    $oldUrl = "https://citspdev.sharepoint.com/sites/NBTeam5Test"
    $newUrl = "https://citspdev.sharepoint.com/sites/NBTeam5Test-123"

    $groupID = "0bf1fc24-a393-4321-aed5-3bd855391529"
    $oldAlias = "NBTeam5Test"
    $newAlias = "NBTeam5Test-123"
    Connect-ExchangeOnline

    Set-UnifiedGroup -Identity "$($groupID)" -EmailAddresses: @{Add = "$newAlias@citspdev.onmicrosoft.com" } -Alias $newAlias -ForceUpgrade
    #Promote alias as a primary SMTP address
    Set-UnifiedGroup -Identity "$($groupID)" -PrimarySmtpAddress "$newAlias@citspdev.onmicrosoft.com" -Alias $newAlias #group email - new primary address
    #Remove old ones
    Set-UnifiedGroup -Identity "$($groupID)" -EmailAddresses: @{Remove = "$oldAlias@citspdev.onmicrosoft.com" } #group email -primary email
    
    #ProcessRenameSiteAddress -oldURL $oldUrl -newURL $newUrl

    $script:EndTimeM365GroupsDailyCache = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    
    LogWrite -Message "[Changing Site Address] Start Time: $($script:StartTimeM365GroupsDailyCache)"
    LogWrite -Message "[Changing Site Address] End Time:   $($script:EndTimeM365GroupsDailyCache)"
    LogWrite -Message  "----------------------- [Changing Site Address] Execution Ended ------------------------"    
    
}
Catch [Exception] {
    LogWrite -Level ERROR "-Unexpected Error: $_ "
}
Finally {
    LogWrite -Message  "----------------------- [Changing Site Address] Completed ------------------------"
}