$AppId      = '9624e216-9e73-4513-9251-4d4382950420'        
$Thumbprint = '1C9696EB9152228A42DAEB5C7075699795311662' 
$TenantId   = "14b77578-9773-42d5-8507-251ca2dc2b06"
$Name = "nih.onmicrosoft.com"
$RootSiteUrl       = "https://nih.sharepoint.com/sites/spoadm"

$AdminCenterUrl    = "https://nih-admin.sharepoint.com"
$BaselineSiteUrl       = "https://nih.sharepoint.com"
$SPM365BoxTeamID = "c7d703c0-4a4d-407f-a4f3-c2d0d89a2004" # CIT SP-M365-Box Team
$spoadmId = "36804259-b5c9-41be-a383-f52f3bd5e840" # SPOADMSVC@nih.gov
$spobaselinelistID = "4162138e-a7e0-4aea-bd7f-33581fccaaca" # https://nih.sharepoint.com/sites/spoadm/Lists/SPOBaseline
$teamsbaselinelistID = "62a14f4b-c2ab-4e67-b94e-7d69a01ab910" # https://nih.sharepoint.com/sites/spoadm/Lists/MSTeamsBaseline
$groupbaselinelistID = "619696b7-692d-4e8b-ad3d-fe22b647d0ac" # https://nih.sharepoint.com/sites/spoadm/Lists/M365GroupBaseline
$userbaselinelistID = "49af5e2f-8aaf-4f6c-b86c-5bc3bfa7d74b" # https://nih.sharepoint.com/sites/spoadm/Lists/M365UserBaseline

$PropertyStatusColumn = "PropertyStatus"

Function PopulateBaselineConfig{
    param(
        [Parameter(Mandatory=$true)]$RefObject,
        [Parameter(Mandatory=$true)]$BaselineListId
    )
    if (!$RefObject){ return;}
    foreach($prop in $RefObject){
        $propName = $prop.Trim()
        $camlQuery ="<View>                
	            <Query>
		            <Where>
			            <Eq>
				            <FieldRef Name='Title' TextOnly = 'True' />
                            <Value Type='Text'>$propName</Value>
			            </Eq>
		            </Where>
	            </Query>
                </View>"

        $iteminlist = @(Get-PnPListItem -List $BaselineListId -Query $camlQuery)
        # Processing new propety
        if ($iteminlist.Count -le 0){
            Write-Host -ForegroundColor Green "[New propery] $propName"        
            $item = @{Title="$propName"}
            Add-PnPListItem -List $BaselineListId -Values $item
        }
        else{
            foreach($item in $iteminlist){                 
                $itemId = $item.Id
                # Processing archived properties last run that become available in SPO object for this time 
                $propStatus = $item.FieldValues["$PropertyStatusColumn"]                
                if ($propStatus -eq 'Archived'){
                    Write-Host -ForegroundColor Yellow "[Archived propery] $propName rolled back"                    
                    $values = @{$PropertyStatusColumn="";"Comments" = "$propName have rolled back"}
                    $ret = Set-PnPListItem -List $BaselineListId -Values $values -Identity $itemId                 
                }
            }    
        }              
    }
    # Processing archived propety
    $baselineItems = @(Get-PnPListItem -List $BaselineListId -Connection $connection)
    [System.Collections.ArrayList]$currentProps = @()    
    
    if ($baselineItems.Count -gt 0){
        $baselineItems | % { $null = $currentProps.Add($_["Title"].Trim()) }
    }    
    $archiveProps = @(Compare-Object -ReferenceObject $currentProps -DifferenceObject $RefObject -IncludeEqual -PassThru | Where-Object {$_.SideIndicator -eq "<=" })
    $archiveCount = $archiveProps.Count

    if ($archiveCount -gt 0){
        $archiveProps.ForEach({        
            $proName = $_.Trim()
            $camlQuery ="<View>	               
	                <Query>
		                <Where>
			                <And>
                                <Eq>
				                    <FieldRef Name='Title' TextOnly = 'True' />
                                    <Value Type='Text'>$proName</Value>
			                    </Eq>                            
                                <Neq>
				                    <FieldRef Name='$PropertyStatusColumn' TextOnly = 'True' />
                                    <Value Type='Choice'>Archived</Value>
			                    </Neq>
                            </And>
		                </Where>
	                </Query>
                </View>"                     
            $existingItems = @(Get-PnPListItem -List $BaselineListId -Query $camlQuery) # -PageSize 100)
            #only update new archived propery - skip existing archived property in SPOBaseline list
            if ($existingItems.Count -gt 0){
                foreach($item in $existingItems){
                    Write-Host -ForegroundColor Yellow "[Archived property] $proName"            
                    #$currentDate = [DateTime]::Today.ToShortDateString()                             
                    $values = @{$PropertyStatusColumn = "Archived";"Comments" = "$proName is archived"}
                    $ret = Set-PnPListItem -List $BaselineListId -Values $values -Identity $item.Id
                }
            }
            
        })
    }
    #end process archived property
}

Function UpdateSPOBaseline {    
    #$siteProps = @(Get-SPOSite -Identity $BaselineSiteUrl -Limit All | Get-Member -MemberType Property | Select Name)
    $siteProps = @(Get-SPOSite -Identity $BaselineSiteUrl -Limit All | Get-Member -MemberType Property).Name
    PopulateBaselineConfig -RefObject $siteProps -BaselineListId $spobaselinelistID
}

Function UpdateTeamBaseline {
    $teamProps = @(Get-Team -GroupId $SPM365BoxTeamID | Get-Member -MemberType Property).Name
    PopulateBaselineConfig -RefObject $teamProps -BaselineListId $teamsbaselinelistID
}
Function UpdateGroupBaseline {
    $groupProps = @(Get-AzureADGroup -ObjectId $SPM365BoxTeamID | Get-Member -MemberType Property).Name
    PopulateBaselineConfig -RefObject $groupProps -BaselineListId $groupbaselinelistID
}
Function UpdateUserBaseline {
    $userProps = @(Get-AzureADUser -ObjectId $spoadmId | Get-Member -MemberType Property).Name
    PopulateBaselineConfig -RefObject $userProps -BaselineListId $userbaselinelistID
}


try{
    $connection = Connect-PnPOnline -Tenant $TenantId -ClientId $AppId -Thumbprint $Thumbprint -Url $RootSiteUrl -ReturnConnection
    
    Write-Host "=== If you are not SharePoint Admin, please activate SharePoint Admin to connect $AdminCenterUrl ==="
    Connect-SPOService -Url $AdminCenterUrl
    Write-Host -ForegroundColor Magenta "Processing site property..." 
    UpdateSPOBaseline    
    Write-Host -ForegroundColor Magenta "Completed site property." 

    Write-Host "=== If you are not Teams Admin, please activate Teams Admin to connect Microsoft Teams Admin ===" 
    Write-Host "=== Please wait until the popup window appears and enter your crendential. Thanks for being patience! haha" 
    Write-Host -ForegroundColor Magenta "Processing MSTeams property..." 
    Connect-MicrosoftTeams
    UpdateTeamBaseline        
    Write-Host -ForegroundColor Magenta "Completed MSTeams property." 

    Write-Host "=== If you are not Teams Admin, please activate Teams Admin to connect Microsoft Teams Admin ==="     
    Connect-AzureAD
    Write-Host "Processing M365 Group property..." 
    UpdateGroupBaseline
    Write-Host -ForegroundColor Magenta "Completed M365 Group property." 
    Write-Host "Processing M365 User property..." 
    UpdateUserBaseline
    Write-Host -ForegroundColor Magenta "Completed M365 User property." 
}
catch{
    throw $_
}
finally{
    Write-Host "Disconnect $RootSiteUrl"
    Disconnect-PnPOnline -Connection $connection 
    Write-Host "==== Disconnect SharePoint Admin ===="
    Disconnect-SPOService
    Write-Host "==== Disconnect Microsoft Teams Admin ===="
    Disconnect-MicrosoftTeams
    Write-Host "==== Disconnect Azure AD ===="
    Disconnect-AzureAD
}


