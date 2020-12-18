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
."$script:RootDir\Common\Lib\GraphAPILibSPOSitesDAO.ps1"
."$script:RootDir\Common\Lib\OperationHelper.ps1"


<#
     =============================================================================================================
      .DESCRIPTION
        Process all change requests - [ChangeRequests] below
        ----------------------------------------------------------------------------------------------------------
        ChangeTypeId	                        ChangeTypeValue	             ChangeRequestTypeId
        ----------------------------------------------------------------------------------------------------------
        FF07775A-9883-4C42-89C1-947F347AE712	Enable External Sharing	     732DBBD1-37ED-4757-9892-CE7E170FDEB7
        B252D6A3-D11A-4AB3-84A9-9DA4769DA3F2	Change Teams Display Name	 32982168-6CA3-402D-9991-64606792A6DE
        ----------------------------------------------------------------------------------------------------------
        RequestStatusId	                        StatusValue
        ----------------------------------------------------------------------------------------------------------
        5A6C2888-1D7F-4FFC-94FA-0A92640C7076	Submitted
        6BFA67B8-18E2-46CB-A8B8-651F36C4C9A5	Cancelled
        E950207B-AAF3-4B65-9BB7-689A6B6AE83D	Completed
        6B934E0F-8784-461B-80D0-A4660F6D1A4E	In Progress
        6C4B8971-FAFF-4B26-BB6F-FD4C5CEA66AA	Pending
    ==============================================================================================================
#>

Function Process-TeamRenameRequests {
    param
    (        
        [Parameter(Mandatory = $true)] $Requests
        
    )

    try {
        LogWrite -Level INFO -Message "Processing Team Rename Requests."
        LogWrite -Level INFO -Message "Total Team Rename Requests: $($Requests.Count)"

        if ($Requests.Count -ile 0) {
            return
        }        

        foreach ($request in $Requests) {
            try {                
                $teamId = $request.GroupId.Trim()
                $newTeamName = $request.NewValue
                LogWrite -Level INFO -Message "Team ID: $teamId"
                $ret = Rename-TeamDisplayName -GroupId $teamId -DisplayName $newTeamName
                $reqStatus = $ret["Status"]
                $reqMessage = $ret["Message"]

                LogWrite -Message "$reqMessage"
                LogWrite -Message "Update change request status."
                UpdateChangeRequest -connectionString $script:connectionString -reqObj $request -reqStatus $reqStatus
                if ($reqStatus -eq $script:Completed){ #only update Teams/Groups table if change request is completed
                    #Verify if Display name is updated against to M365
                    $team = Get-Team -GroupId $teamId
                    if ($team.DisplayName -eq $newTeamName){
                        LogWrite -Message "Update group display name [$newTeamName] into Groups and Teams table."
                        UpdateGroupTeamPostRename -connectionString $script:connectionString -teamObj $team
                    }
                    else{
                        LogWrite -Message "Nothing updated."
                    }
                } 
            }
            catch {               
                LogWrite -Level ERROR -Message "[Process-TeamRenameRequests] - Error renaming team $teamId : $($_.Exception)"      
            }            
        }
        LogWrite -Level INFO -Message "Processing Team Rename Requests completed."        
    }
    catch {
        LogWrite -Level ERROR -Message "[Process-TeamRenameRequests] - Unexpected exception: $($_.Exception)"        
        throw $_.Exception
    }
}

Function Process-ExternalSharingRequests {
    param
    (        
        [Parameter(Mandatory = $true)] $Requests        
    )

    try {
        LogWrite -Level INFO -Message "Processing External Sharing Requests."
        LogWrite -Level INFO -Message "Total External Sharing Requests: $($Requests.Count)"

        if ($Requests.Count -ile 0) {
            return
        }

        $tenantProps = Get-PnPTenant
        $currentShareSettings = $tenantProps.SharingCapability
        LogWrite -Level INFO -Message "Current SharingCapability settings in the tenant: $($currentShareSettings)."
        
        if ($currentShareSettings -eq 'Disabled') {
            LogWrite -Level ERROR -Message "Sharing is currently disabled on the tenant level!"
            return
        }        

        foreach ($request in $Requests) {
            try {                
                $siteUrl = $request.SiteUrl
                $externalSharingEnabled = $request.NewValue
                $newValue = "Disabled"                
                if ($request.NewValue -eq 1){
                    $newValue = $currentShareSettings
                }
                if ([string]::IsNullOrWhiteSpace($siteUrl)) {
                    LogWrite -Level ERROR -Message "[Site URL cannot be null]: $($request.ChangeRequestTypeId)"
                    throw "Site URL cannot be null."
                }
                $siteUrl = $siteUrl.Trim()

                #--Update ExternalSharingEnable to Sites table
                LogWrite -Message "Update [ExternalSharingEnable] field in Sites table"                    
                UpdateSPOSiteExternalSharingRecord $script:connectionString $site $externalSharingEnabled
                #--Update external sharing for SPO site                    
                $ret = Update-ExternalSharing -SiteUrl $siteUrl -SharingCapability $newValue
                $reqStatus = $ret["Status"]
                $reqMessage = $ret["Message"]
                LogWrite -Message "$reqMessage"                
                UpdateChangeRequest -connectionString $script:connectionString -reqObj $request -reqStatus $reqStatus
                
                if ($reqStatus -eq $script:Completed){ #only update sites table if change request is completed
                    #Verify SharingCapability before update to DB
                    $site = Get-PnPTenantSite -Url $siteUrl
                    LogWrite -Message "Update [SharingCapability] to the Sites table"
                    UpdateSPOSiteExternalSharingRecord $script:connectionString $site $externalSharingEnabled
                } 

            }
            catch {                
                LogWrite -Level ERROR -Message "[Process-ExternalSharingRequests] - Error external sharing $SiteUrl : $($_.Exception)"      
            }                  
        }
        LogWrite -Level INFO -Message "Processing external sharing completed."        
    }
    catch {
        LogWrite -Level ERROR -Message "[Process-ExternalSharingRequests] - Unexpected exception: $($_.Exception)"        
        throw $_.Exception
    }
}

try {
    #-------- Set Global Variables ---------
    Set-TenantVars
    Set-AzureAppVars
    Set-DBVars   
    Set-LogFile -logFileName $logFileName
    Set-ChangeTypeVars
    Set-StatusVars
    #-------- Set Global Variables Ended ---------    
    
    $startTime = Get-Date -Format "yyyy-MM-dd HH:mm:ss"           
    LogWrite -Message "----------------------- [Process M365 Operations] Execution Started -----------------------"    
    #--- Get all requests from DB-[ChangeRequests] with Status either Submitted or Pending (mark as Pending if something go wrong) ---
    $requests = GetActiveChangeRequests $script:connectionString
    #filter by ChangeTypeId
    $externalSharingRequests = $requests | ? { $_.ChangeTypeId -eq $script:ExternalSharing }
    $teamRenameRequests = $requests | ? { $_.ChangeTypeId -eq $script:TeamsDisplayName }
    $teamRenameRequests = @($teamRenameRequests)
    $externalSharingRequests = @($externalSharingRequests)

    #--- PnP ---
    if ($externalSharingRequests){        
        LogWrite -Message "Connecting to SharePoint Admin Center '$($script:SPOAdminCenterURL)'..."
        $script:TenantContext = ConnectPnpOnlineOAuth -TenantId $script:TenantId -ClientId $script:appIdOperationSupport -Thumbprint $script:appThumbprintOperationSupport -Url $script:SPOAdminCenterURL            
        LogWrite -Message "SharePoint Admin Center '$($script:SPOAdminCenterURL)' is now connected."         
        Process-ExternalSharingRequests $externalSharingRequests
    }

    #--- MS Teams Operation using Teams API ---            
    if ($teamRenameRequests -and $teamRenameRequests.Count -gt 0){                
        Connect-MicrosoftTeams -TenantId $script:TenantId -ApplicationId  $script:appIdOperationSupport -CertificateThumbprint $script:appThumbprintOperationSupport
        Process-TeamRenameRequests $teamRenameRequests        
    } 

    $endTime = Get-Date -Format "yyyy-MM-dd HH:mm:ss"    
    LogWrite -Message "[Process M365 Operations] Start Time: $startTime"
    LogWrite -Message "[Process M365 Operations] End Time:   $endTime"
    LogWrite -Message  "----------------------- [Process M365 Operations] Execution Ended ------------------------"  
    
}
Catch [Exception] {
    LogWrite -Level ERROR "-Unexpected Error: $_ "
}
Finally {
    LogWrite -Level INFO -Message "Disconnect Microsoft Teams."
    DisconnectMicrosoftTeams
    LogWrite -Level INFO -Message "Disconnect SharePoint Admin Center."    
    LogWrite -Message  "----------------------- [Process M365 Operations] Completed ------------------------"
}

