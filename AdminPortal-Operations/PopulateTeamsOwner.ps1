<#
      ===========================================================================
      .DESCRIPTION
        Find Microsoft Teams teams without an Owner
        Add IC MS Teams POC as an owner to Orphaned Teams
        -$MFA: using Teams Admin account to run        
#>

Param
(
    [Parameter(Mandatory = $false)]
    #[switch]$MFA,            
    [string]$AdminSvc = "SPOADMSVC@nih.gov,m365collabadmsvc@nih.onmicrosoft.com",
    [Parameter(Mandatory = $false)]
    [string]$ICName = "All"
)

$0 = $MyInvocation.MyCommand.Definition
$dp0 = [System.IO.Path]::GetDirectoryName($0)
$currentFolder = $dp0.Substring($dp0.LastIndexOf('\') + 1)
$script:RootDir = Resolve-Path "$dp0\.."
$logFileName = $MyInvocation.MyCommand.Name.Substring(0, $MyInvocation.MyCommand.Name.IndexOf('.'))
$logFileName = "$currentFolder\$logFileName"
$reportFile = "ProcessedTeamsWithoutOwners.csv"
$errorGroupsFile = "TeamsWithoutOwners_NoProcessed.csv"
#Include core
."$script:RootDir\Common\Lib\LibCore.ps1"
."$script:RootDir\Common\Lib\LibUtils.ps1"
#Include dependent functionality
."$script:RootDir\Common\Lib\LibO365.ps1"
."$script:RootDir\Common\Lib\GraphAPILibO365GroupsDAO.ps1"

Function ListICAdmins {
    LogWrite -Message "Connecting to SharePoint Online '$($script:SPORootSiteURL)'..."
    $listContext = ConnectPnpOnlineOAuth -TenantId $script:TenantId -ClientId $script:appIdOperationSupport -Thumbprint $script:appThumbprintOperationSupport -Url $script:SPORootSiteURL    
    LogWrite -Message "SharePoint Online '$($script:SPORootSiteURL)' is now connected."
    $listId = "{47b89bca-089c-4250-9281-c4a9f5411c2a}"
    LogWrite -Message "Get MSTeamsPOC and Title[IC Name] from IC Admins list"
    $camlQuery = "<View><ViewFields><FieldRef Name='Title' /><FieldRef Name='MSTeamsPOC' /></ViewFields></View>"                  
    $teamsPOC = @(Get-PnPListItem -List $listId -Connection $listContext -Query $camlQuery)
    
    if ("All" -ne $ICName) {
        $teamsPOC = $teamsPOC.Where({$_["Title"] -eq $ICName})
    }
    
    $numberOfTeamsPOC = $teamsPOC.Count
    LogWrite -Message "Total IC: $($numberOfTeamsPOC)."  
    [System.Collections.ArrayList]$ICTeamsPOC = @()
    if ($numberOfTeamsPOC -gt 0) {                
        $teamsPOC | & { process {                
                $UserAccount = @($_["MSTeamsPOC"].Email)           
                #$Name = @($_["MSTeamsPOC"].LookupValue)
                $null = $ICTeamsPOC.Add([PSCustomObject]@{
                        ICName = $_["Title"]
                        Emails = $UserAccount #@($_["MSTeamsPOC"].Email)
                        #Names = $Name
                    })
            } }    
    }
    DisconnectPnpOnlineOAuth -Context $listContext
    return $ICTeamsPOC
}

Function ListOrphanedTeams {
    LogWrite -Message "Getting orphaned teams from DB..."
    [System.Collections.ArrayList]$teamsWithoutOwnersProceeded = @()
    [System.Collections.ArrayList]$teamsWithoutOwnersNoProceeded = @()
    $orphanedTeams = @(GetOrphanedTeams $script:connectionString $AdminSvc)    
    if ($orphanedTeams.Count -gt 0) {
        LogWrite -Message "Getting GRAPH API access token..."
        $script:appCertOperationSupport
        $certificate = Get-Item Cert:\\LocalMachine\\My\* | Where-Object { $_.Subject -ieq "CN=$($script:appCertOperationSupport)" }    
        $authToken = Connect-NIHO365GraphWithCert -TenantName $script:TenantName -AppId $script:appIdOperationSupport -Certificate $certificate
        LogWrite -Message "Processing orphaned Teams..."
        
        if ("All" -ne $ICName) {
            $orphanedTeams = @($orphanedTeams | ? { $_.ICName -eq $ICName })
        }
        if ($orphanedTeams.Count -gt 0) {            
            LogWrite -Message "MS Teams POC from IC Admin list..."    
            $ICTeamsPOC = ListICAdmins            
            $orphanedTeams | & { process {
                    $GroupId = $_.GroupId
                    $ICName = $_.ICName
                    $DisplayName = $_.GroupDisplayName
                    $GroupOwners = $_.GroupOwners
                    LogWrite -Message "Processing adding Teams POC for [$ICName]" 
                    LogWrite -Message "Getting M365 group owner..."           
                    $owners = @(Get-NIHO365GroupOwners -AuthToken $authToken -Id $GroupId)
                    if ($owners.Count -eq 0 -or ($owners.Count -eq 1 -and $AdminSvc -contains $owners.userPrincipalName)) {
                        LogWrite -Message "Adding IC Admins to M365 group..."
                        $TeamsPOC = ($ICTeamsPOC | ? { $_.ICName -eq $ICName }).Emails

                        if ($TeamsPOC.Count -gt 0) {
                            $TeamsPOCIds = [System.Collections.ArrayList]@()
                            $TeamsPOC | & { process { 
                                    $id = (Get-NIHO365UserByEmail -AuthToken $authToken -EmailAddress "$_").id     
                                    if ($null -ne $id) { $TeamsPOCIds.Add($id) }
                                } }
                            #$GroupId = "BB36AB67-853C-49FA-89C1-789D84C04ADA"
                            #$result = Add-NIHTeamMember -AuthToken $authToken -Group $GroupId -Member $TeamsPOC -AsOwner
                            #$result = Add-NIHO365GroupMember -AuthToken $authToken -Group $GroupId -Members $TeamsPOCIds -AsOwner
                            $null = $teamsWithoutOwnersProceeded.Add([PSCustomObject]@{
                                    ICName          = $ICName
                                    Id              = $GroupId
                                    DisplayName     = $DisplayName
                                    TeamOwnersAdded = $TeamsPOC -join ";"                                
                                    Message         = ""
                                })
                        }
                        else {
                            $null = $teamsWithoutOwnersNoProceeded.Add([PSCustomObject]@{
                                    ICName      = $ICName
                                    Id          = $GroupId
                                    DisplayName = $DisplayName
                                    GroupOwners = $GroupOwners                               
                                    Message     = "IC invalid"
                                })
                        }
                    }
                } } 
        }
        $logPath = "$($script:DirLog)"
        if (!(Test-Path $logPath)) { 
            LogWrite -Message "Creating $logPath" 
            New-Item -ItemType "directory" -Path $logPath -Force
        } 
        if ($teamsWithoutOwners -ne $null) {
            $teamsWithoutOwners | Export-Csv "$logPath\$reportFile" -Encoding ASCII -NoTypeInformation
        }
        else {
            LogWrite -Message "Object is null. Couldn't save file to $reportFile"            
        }
        if ($teamsWithoutOwnersNoProceeded -ne $null) {
            $teamsWithoutOwnersNoProceeded | Export-Csv "$logPath\$errorGroupsFile" -Encoding ASCII -NoTypeInformation
        }
        else {
            LogWrite -Message "Object is null. Couldn't save file to $errorGroupsFile"            
        }
    }    
}

Function ListOrphanedGroups {
    param([switch]$AddOwner) 
    LogWrite -Message "Getting orphaned teams from DB..."

    [System.Collections.ArrayList]$teamsWithoutOwnersProceeded = @()
    [System.Collections.ArrayList]$teamsWithoutOwnersNoProceeded = @()

    LogWrite -Message "Getting M365 Groups/Teams without an owner or owner which is either service account or disabled user..."
    $today = [DateTime]::Now.ToString("MM-dd-yyyy")
    $logTeamsData = "$($script:RootDir)\Logs\$logFileName\$today"

    if ($AddOwner.IsPresent){
        $today = [DateTime]::Now.AddDays(-3).ToString("MM-dd-yyyy")
        $logTeamsData = "$($script:RootDir)\Logs\$logFileName\$today\$reportFile"
        
        $logTeamsData = "D:\Scripting\O365DevOps\Logs\AdminPortal-Operations\PopulateTeamsOwner\04-24-2021\$reportFile"        

        if (test-path $logTeamsData) {
            $orphanedGroups = Import-csv $logTeamsData
            $dtGroupByIC = @($orphanedGroups | Group-Object ICName | Select-Object -property @{N='ICName';E={$_.Name}}, @{N='ICCount';E={$_.Count}})
            $dtOrphanedGroups = @($orphanedGroups)                     
        }
        else {
            LogWrite -Level INFO -Message "Teams Owner Data not found: $logTeamsData"
        }
    }
    else {
        $orphanedGroups = @(GetOrphanedGroups $script:connectionString $AdminSvc)
        if ($orphanedGroups.Count -gt 0) {
            $dtGroupByIC = @($orphanedGroups.Tables[0])
            $dtOrphanedGroups = @($orphanedGroups.Tables[1])
        }
    }
  
    if ("All" -ne $ICName) {
            $dtGroupByIC = @($dtGroupByIC | ? { $_.ICName -eq $ICName })
        }
        
        if ($dtGroupByIC.Count -gt 0) {
            LogWrite -Message "MS Teams POC from IC Admin list..."    
            $ICTeamsPOC = @(ListICAdmins)
            if ($ICTeamsPOC.Count -le 0) { 
                LogWrite -Message "No Teams IC POC found." 
                return 
            }

            LogWrite -Message "Getting GRAPH API access token..."                
            $authToken = Connect-GraphAPIWithCert -TenantId $script:TenantId -AppId $script:appIdOperationSupport -Thumbprint $script:appThumbprintOperationSupport

            LogWrite -Message "Processing orphaned Teams..."
         
            foreach ($ic in $dtGroupByIC) {
                $ICName = $ic.ICName    
                $icTotalOrphanedGrps = $ic.ICCount        
                $orphanedGroups = @($dtOrphanedGroups | ? { $_.ICName -eq $ICName })
                LogWrite -Message "[$icName] have $icTotalOrphanedGrps orphaned groups"

                if ($orphanedGroups.Count -gt 0) {
                    $orhanedGroupsEmail = @()
                    #$TeamsPOC = ($ICTeamsPOC | ? { $_.ICName -eq $ICName }).Emails #this include individual and DL email
                    $TeamsPOC = ($ICTeamsPOC).Where({ $_.ICName -eq $ICName }).Emails #this include individual and DL email
                    # Build html email template
                    $subject = "Action Required: Assign Owners for IC Microsoft Teams and M365 Groups"
                    $preContent = PreContentEmail -ICName $ICName -AddOwner
                    $postContent = PostContentEmail -AddOwner                  
                    # end build email template

                    [System.Collections.ArrayList]$teamsWithoutOwnersByIC = @()

                    foreach ($g in $orphanedGroups) {
                        $ICPOCAdded = $false                        
                        $GroupId = $g.GroupId
                        if ($g.GroupId.Guid){
                            $GroupId = $g.GroupId.Guid
                        }
                        $GroupAlias = $g.GroupAlias
                        $GroupName = $g.Name
                        $GroupLink = $g.TeamUrl 
                       
                        if ("" -eq $GroupLink) {
                            $GroupLink = "https://outlook.office365.com/mail/group/groups.nih.gov/$GroupAlias/email" 
                        }
                        $M365Type = $g.Type
                                               
                        #LogWrite -Message "Verifying if the group need to be proceeded adding IC Teams POC..."
                        $owners = @(Get-NIHO365GroupOwners -AuthToken $authToken -Id $GroupId)
                        $AdminSvc = $AdminSvc.split(",")

                        if ($owners.Count -le 1) {
                            $ICPOCAdded = $true
                            $currentOwner = $owners.userPrincipalName
                            if ($null -ne $currentOwner -and $AdminSvc -notcontains $currentOwner){
                                LogWrite -Message "Verifying if the owner is active or inactive..."
                                $userOwner = Get-NIHO365User -AuthToken $authToken -UserID $owners.id -Select id,userPrincipalName,accountEnabled  
                                if ($userOwner.accountEnabled){
                                    LogWrite -Message "$currentOwner is active. Adding IC Teams POC as an owner is not required. Skip..."
                                    $ICPOCAdded = $false
                                }
                            }
                        }

                        if ($ICPOCAdded) {                            
                            LogWrite -Message "...Adding ($TeamsPOC) to the group [$GroupName]"                             
                            
                            if ($TeamsPOC.Count -gt 0) {
                                $TeamsPOCIds = [System.Collections.ArrayList]@()                               

                                $TeamsPOC | & { process {
                                        $GroupEmail = $_                                                                               
                                        $members = @()
                                        ConnectAzureADOAuth -TenantId $script:TenantId -ClientId $script:appIdOperationSupport -Thumbprint $script:appThumbprintOperationSupport    
                                        $g = Get-AzureADGroup -Filter "Mail eq '$GroupEmail'"
                                        
                                        if ($g){
                                            $members = @(Get-RecursiveAzureAdGroupMemberUsers -AzureGroup $g)
                                        }   
                                        DisconnectAzureAD                                       
                                        if ($members -and $members.count -gt 0){
                                            foreach($i in $members){
                                                if ($null -ne $i) { $TeamsPOCIds.Add($i) }
                                            }
                                        }
                                        else{
                                            $id = (Get-NIHO365UserByEmail -AuthToken $authToken -EmailAddress "$GroupEmail").id
                                            if ($null -ne $id) { $TeamsPOCIds.Add($id) }
                                        }                                       
                                        
                                    } }                                
                                $TeamsPOCIds = $TeamsPOCIds | select -Unique
                                $msg = ""
                                $operation = "EmailNotify"

                                if ($AddOwner.IsPresent){
                                    #$subject = "Action Required: Re-assign Owners for IC Microsoft Teams and M365 Groups"
                                    $preContent = PreContentEmail -ICName $ICName -AddOwner
                                    $postContent = PostContentEmail -AddOwner
                                    
                                    LogWrite -Message "Adding ($TeamsPOC) to the group [$GroupName]"
                                    <#$result = Add-NIHO365GroupMember -AuthToken $authToken -Group $GroupId -Members $TeamsPOCIds -AsOwner                                
                                    if ($result -and $result -ne ""){
                                        $msg = $result | ConvertTo-Json
                                        
                                    }#>
                                    $operation = "Added"
                                }

                                $teamsWithoutOwnersByIC.Add([PSCustomObject]@{
                                        "Name" = "<a href='$GroupLink'>$GroupName</a>"                                        
                                        "Type" = $M365Type                                        
                                    }) 
                                $null = $teamsWithoutOwnersProceeded.Add([PSCustomObject]@{
                                        ICName          = $ICName
                                        GroupId         = $GroupId
                                        GroupAlias      = $GroupAlias
                                        Name          = $GroupName
                                        Type          = $M365Type
                                        TeamUrl       = $GroupLink 
                                        GroupOwners   = $currentOwner
                                        TeamOwnersAdded = $TeamsPOC -join ";"
                                        Operation       = $operation                                
                                        Message         = $msg
                                    })                               
                            }
                            else {
                                $null = $teamsWithoutOwnersNoProceeded.Add([PSCustomObject]@{
                                        ICName          = $ICName
                                        GroupId         = $GroupId
                                        GroupAlias      = $GroupAlias
                                        Name          = $GroupName
                                        Type          = $M365Type
                                        TeamUrl       = $GroupLink 
                                        GroupOwners   = $currentOwner
                                        TeamOwnersAdded = $null                          
                                        Message     = "IC invalid"
                                    })
                            }
                        }
                    }
                    #
                    #
                    if ($teamsWithoutOwnersByIC -and $teamsWithoutOwnersByIC.Count -gt 0) {
                        $ToEmails = $TeamsPOC -join ";"                        
                        $html = $teamsWithoutOwnersByIC.GetEnumerator() | ConvertTo-HTML -PreContent $preContent -PostContent $postContent -Fragment
                        $bodyEmail = [System.Web.HttpUtility]::HtmlDecode($html)
                        $bodyEmail = $bodyEmail -replace '<table>', '<table cellpadding="5" cellspacing="2" style="border: 1px solid black;border-collapse: collapse;width:100%">'                    
                        $bodyEmail = $bodyEmail -replace '<th>', '<th align="left" style="border: 1px solid black;">'
                        $bodyEmail = $bodyEmail -replace '<td>', '<td style="border: 1px solid black;">'
                    
                        LogWrite -Message "Sending an notification email to ($ToEmails) after adding them as an owner of the orphaned group"
                        #$to = "dan.le@nih.gov;rahul.babar@nih.gov;ngan.bui@nih.gov" # for testing; $ToEmails
                        $to = "ngan.bui@nih.gov"
                        #$emailContent = [string]::Format("{0} {1} {2}", $preContent, $bodyEmail, $postContent)
                        #SendEmail -subject $subject -body $bodyEmail -To $ToEmails
                        #SendEmail -subject $subject -body $bodyEmail -To $ToEmails

                        SendEmail -subject $subject -body $bodyEmail -To $to
                    }
                }
            }
        } 

        $attachedFiles = @()
        $today = [DateTime]::Now.ToString("MM-dd-yyyy")
        $today = "$($script:RootDir)\Logs\$logFileName\$today"

        if (!(Test-Path $today)) { 
            LogWrite -Message "Creating $today" 
            New-Item -ItemType "directory" -Path $today -Force
        }

        if ($teamsWithoutOwnersProceeded -ne $null -and $teamsWithoutOwnersProceeded.Count -gt 0) {
            $processedFile = "$today\$reportFile"
            $teamsWithoutOwnersProceeded | Export-Csv $processedFile -Encoding ASCII -NoTypeInformation
            $attachedFiles += $processedFile
        }
        else {
            LogWrite -Message "There are currently no groups/teams missing onwer"            
        }
        if ($teamsWithoutOwnersNoProceeded -ne $null -and $teamsWithoutOwnersNoProceeded.Count -gt 0) {
            $noProcessedFile = "$today\$errorGroupsFile"
            $teamsWithoutOwnersNoProceeded | Export-Csv $noProcessedFile -Encoding ASCII -NoTypeInformation
            $attachedFiles += $noProcessedFile
        }
        else {
            LogWrite -Message "There are currently no groups/teams with invalid IC"
        }
        #--Sending report email to SP Admins         
        #GenerateICGroupOwnersReport -attachments $attachedFiles   
}

Function GetADGroupMembers{
    [Parameter(Mandatory = $true)]$GroupEmail
    
    $members = @()
    ConnectAzureADOAuth -TenantId $script:TenantId -ClientId $script:appIdOperationSupport -Thumbprint $script:appThumbprintOperationSupport    
    $g = Get-AzureADGroup -Filter "Mail eq '$GroupEmail'"
    
    if ($g){
        $members = Get-RecursiveAzureAdGroupMemberUsers -AzureGroup $g
    }   
    DisconnectAzureAD
    $members
}

Function Get-RecursiveAzureAdGroupMemberUsers{
    [cmdletbinding()]
    param(
       [parameter(Mandatory=$True,ValueFromPipeline=$true)]
       $AzureGroup
    )

    Begin{
        If(-not(Get-AzureADCurrentSessionInfo)){
            ConnectAzureADOAuth -TenantId $script:TenantId -ClientId $script:appIdOperationSupport -Thumbprint $script:appThumbprintOperationSupport
        }
    }
    Process {
        Write-Verbose -Message "Enumerating $($AzureGroup.DisplayName)"
        $Members = Get-AzureADGroupMember -ObjectId $AzureGroup.ObjectId -All $true
        $UserMembers = $Members | Where-Object{$_.ObjectType -eq 'User'}
        If($Members | Where-Object{$_.ObjectType -eq 'Group'}){
            $UserMembers += $Members | Where-Object{$_.ObjectType -eq 'Group'} | ForEach-Object{ Get-RecursiveAzureAdGroupMemberUsers -AzureGroup $_}
        }
        $UserMembers = @($UserMembers).ObjectId        
    }
    End {
        $UserMembers        
    }
}

Function SendEmailToTeamsPOC {
    param($content, $to)    
    $subject = "Action Required: Assign Owners for IC Microsoft Teams and M365 Groups"    
    $body = "<p>Hello IC Microsoft Teams POC,</p>"
    $body += "<p>You are receiving this email because you are listed as your IC's Microsoft Teams point of contact on our <a href='https://nih.sharepoint.com/Lists/IC%20Admins'>IC Admins List</a>.</p>"
    $body += "<p>We have identified the following Microsoft Teams and M365 Groups at your IC that currently do not have an assigned owner. In an effort to ensure owners are assigned to all Microsoft Teams and M365 Groups, we have temporarily added you as the owner. Please work with these teams/groups to assign someone from the member list as the new owner.</p>"    
    $body += $content
    $body += "<p>Follow these useful instructions to make changes to Microsoft Teams or M365 Groups:<br />"
    $body += "<ul><li><a href='https://support.microsoft.com/en-us/office/add-members-to-a-team-in-teams-aff2249d-b456-4bc3-81e7-52327b6b38e9'>To add members to a team in Microsoft Teams</a></li>"
    $body += "<li><a href='https://support.microsoft.com/en-us/office/delete-a-team-c386f91b-f7e6-400b-aac7-8025f74f8b41'>To delete an unused team in Microsoft Teams</a></li>"
    $body += "<li><a href='https://support.microsoft.com/en-us/topic/delete-a-group-in-outlook-ca7f5a9e-ae4f-4cbe-a4bc-89c469d1726f'>To delete unused M365 Groups</a></li>"
    $body += "</ul></p>"
    $body += "<p>If you have any questions or need assistance, contact the CIT M365 Collab Tenant Admins at <a href='mailto:CITM365CollabTenantAdmins@mail.nih.gov'>CITM365CollabTenantAdmins@mail.nih.gov</a>.</p>"
    $body += "<p>Thank you, <br />CIT M365 Collab Tenant Admins</p>"   
    SendEmail -subject $subject -body $body -To $to
}

Function PreContentEmail {
    param($ICName,[switch]$AddOwner)

    #--Firt Email
    $preContent = "<p><i>Note: **This is an automated message - Please do not reply directly to this email**</i></p>
        <p>Hello $ICName Microsoft Teams Admins,</p>
        <p>You are receiving this email because you are listed as your IC's Microsoft Teams point of contact on our <a href='https://nih.sharepoint.com/Lists/IC%20Admins'>IC Admins List</a>.</p>
        <p>We have identified the following Microsoft Teams and M365 Groups at your IC that currently do not have an active owner. In an effort to ensure owners are assigned to all Microsoft Teams and M365 Groups, please work with these Teams/Groups to assign someone from the IC as the new owner.</p>
        <h3>Microsoft Teams or M365 Groups Missing Owner</h3>"
    
    #--Second email
    if ($AddOwner.IsPresent){
        $preContent = "<p><i>Note: **This is an automated message - Please do not reply directly to this email**</i></p>
        <p>Hello $ICName Microsoft Teams Admins,</p>
        <p>According to our records, the identified Microsoft Teams/M365 Groups below still do not have an active owner. To ensure owners are assigned to all Microsoft Teams and M365 Groups, we have temporarily added you as the owner. Please work with these teams/groups to assign someone from the member list as the new owner.</p>        
        <h3>Microsoft Teams or M365 Groups Missing Owner</h3>"
    }

    return $preContent   
}
Function PostContentEmail {
    param([switch]$AddOwner)

    #--Firt Email
    $postContent = "<p>Follow these useful instructions to make changes to Microsoft Teams or M365 Groups:<br />
        <ul><li><a href='https://support.microsoft.com/en-us/office/add-members-to-a-team-in-teams-aff2249d-b456-4bc3-81e7-52327b6b38e9'>To add members to a team in Microsoft Teams</a></li>
        <li><a href='https://support.microsoft.com/en-us/office/delete-a-team-c386f91b-f7e6-400b-aac7-8025f74f8b41'>To delete an unused team in Microsoft Teams</a></li>
        <li><a href='https://support.microsoft.com/en-us/topic/delete-a-group-in-outlook-ca7f5a9e-ae4f-4cbe-a4bc-89c469d1726f'>To delete unused M365 Groups</a></li>
        </ul></p>
        <p>If you have any questions or need assistance, contact the CIT M365 Collaboration Support Team at <a href='mailto:CITM365CollabTenantAdmins@mail.nih.gov'>CITM365CollabTenantAdmins@mail.nih.gov</a>.</p>
        <p>Thank you, <br />CIT M365 Collaboration Support Team</p>"
    
    #--Second email       
    if ($AddOwner.IsPresent){
        $postContent = "<p>Follow these useful instructions to make changes to Microsoft Teams or M365 Groups:<br />
        <ul><li><a href='https://support.microsoft.com/en-us/office/add-members-to-a-team-in-teams-aff2249d-b456-4bc3-81e7-52327b6b38e9'>To add members to a team in Microsoft Teams</a></li>
        <li><a href='https://support.microsoft.com/en-us/office/delete-a-team-c386f91b-f7e6-400b-aac7-8025f74f8b41'>To delete an unused team in Microsoft Teams</a></li>
        <li><a href='https://support.microsoft.com/en-us/topic/delete-a-group-in-outlook-ca7f5a9e-ae4f-4cbe-a4bc-89c469d1726f'>To delete unused M365 Groups</a></li>
        </ul></p>
        <p>If you have any questions or need assistance, contact the CIT M365 Collaboration Support Team at <a href='mailto:CITM365CollabTenantAdmins@mail.nih.gov'>CITM365CollabTenantAdmins@mail.nih.gov</a>.</p>
        <p>Thank you, <br />CIT M365 Collaboration Support Team</p>"
    }
    return $postContent
}

Function GenerateICGroupOwnersReport {
    param($attachments)
    LogWrite -Message "Sending Email Report: [Populate IC Teams POC as group owner for M365 groups without owners]"    
    $subject = "[M365 DevOps] Populate MS Teams Default Owners"
    $body = "There are currently no groups/teams missing owner"
    if ($attachments.Length -gt 0) { 
        $body = "<p><b>Description:</b> This job will add IC Teams POCs from IC Admin list to Teams/Groups without owners<br />"
        $body += "Please review and address any issues from the attached files if needed.</p>"    
    }
    SendEmail -subject $subject -body $body -Attachements $attachments #-To "ngan.bui@nih.gov"
    LogWrite -Message "Sending Email Report: [Populate IC Teams POC as group owner for M365 groups without owners] completed."
}
Try {
    #log file path
    Set-LogFile -logFileName $logFileName
    $startTime = Get-Date -Format "yyyy-MM-dd HH:mm:ss"           
    LogWrite -Message "----------------------- [Populate Teams Onwer] Execution Started -----------------------"
    Set-TenantVars
    Set-AzureAppVars
    Set-DBVars
    Set-EmailVars

    #$today = [DateTime]::Now.ToString("MM-dd-yyyy")
    #$today = "$($script:RootDir)\Logs\$logFileName\$today"

      
    <#
    # Connect MicrosoftTeams with MFA user for running manually - Team Admin and SharePoint Admin role are required
    <if($MFA.IsPresent) {
        Write-Host Connecting to MicrosoftTeams...
        Connect-MicrosoftTeams
        Write-Host Connecting to SharePoint site to retrive IC Admins...
        Connect-PnPOnline
    }
    # Connect MicrosoftTeams with App-Only for scheduling purpose
    else{
        ListICAdmins
        ListOrphanedTeams    
    }
    Disconnect-MicrosoftTeams
    #>    
    LogWrite -Message "Processing orphaned Teams/Groups..."
    
    
    $today = (Get-Date).DayOfWeek
    if ($today -eq 'Monday'){
        LogWrite -Message  "Sending email notification only..."
        ListOrphanedGroups
    }
    elseif ($today -eq 'Saturday'){
        ListOrphanedGroups -AddOwner  
    }   

    $endTime = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    
    LogWrite -Message "[Populate Teams Onwer] Start Time: $startTime"
    LogWrite -Message "[Populate Teams Onwer] End Time:   $endTime"
    LogWrite -Message  "----------------------- [Populate Teams Onwer] Execution Ended ------------------------"    
    
}
Catch [Exception] {
    LogWrite -Level ERROR "-Unexpected Error: $_ "
}
Finally {
    LogWrite -Message  "----------------------- [Populate Teams Onwer] Completed ------------------------"
}
