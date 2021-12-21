$0 = $MyInvocation.MyCommand.Definition
$dp0 = [System.IO.Path]::GetDirectoryName($0)
$currentFolder = $dp0.Substring($dp0.LastIndexOf('\')+1)
$script:RootDir = Resolve-Path "$dp0\..\.." 
$logFileName = $MyInvocation.MyCommand.Name.Substring(0, $MyInvocation.MyCommand.Name.IndexOf('.'))
$logFileName = "$currentFolder\$logFileName"
#Include core
."$script:RootDir\Common\Lib\LibCore.ps1"

Function PopulateTenantSettings{
param(
        [Parameter(Mandatory=$true)][string]$tenantName,
        [Parameter(Mandatory=$true)][string]$baselineListId
    )
    $AdminCenterUrl = "https://$tenantName-admin.sharepoint.com"
    $AppId      = '9624e216-9e73-4513-9251-4d4382950420'        
    $Thumbprint = '1C9696EB9152228A42DAEB5C7075699795311662' 
    $TenantId   = "14b77578-9773-42d5-8507-251ca2dc2b06"
    $Name = "nih.onmicrosoft.com"
    $RootSiteUrl       = "https://nih.sharepoint.com/sites/spoadm"

    Write-Host "Please use SharePoint Admin account to connect $AdminCenterUrl"   
    Connect-SPOService -Url $AdminCenterUrl
     <#
    # CITSPDEV
    $baselineListDEV = "https://nih.sharepoint.com/sites/spoadm/Lists/SPOBaselineConfigurationDEV"    
    $baselineListDEVId = "{aef64a07-a38f-4790-a2e8-7de8b1b5e0b8}"
    # NIHDEV
    $baselineListNIHDEV = "https://nih.sharepoint.com/sites/spoadm/Lists/SPOBaselineConfigurationNIHDEV"    
    $baselineListNIHDEVId = "{da1e1081-64d9-436e-875f-885ca04ead9f}"
    #>
    $newProps = @{}
    $updateProps = @{}
    [System.Collections.ArrayList] $spoTenantProps = @()

    $DestinationColumn = "NIHPreviousSettings"
    $SourceColumn = "NIHCurrentSettings"
    $PropertyStatusColumn = "PropertyStatus"
   
    $tenantConfig = Get-SPOTenant | select *
    $tenantProps = @(Get-SPOTenant | Get-Member -MemberType Property).Name
    $connection = Connect-PnPOnline -Tenant $TenantId -ClientId $AppId -Thumbprint $Thumbprint -Url $RootSiteUrl -ReturnConnection
    $queryActiveItems = "<ViewFields>
		            <FieldRef Name = 'Title' />
                    <FieldRef Name = '$DestinationColumn' />
                    <FieldRef Name = '$SourceColumn' />
                    <FieldRef Name = '$PropertyStatusColumn' />
	            </ViewFields>
	            <Query>
		            <Where>
			            <Neq>
				            <FieldRef Name='$PropertyStatusColumn' TextOnly = 'True' />
                            <Value Type='Choice'>Archived</Value>
			            </Neq>
		            </Where>
	            </Query>
            </View>"
    
    $baselineItems = @(Get-PnPListItem -List $baselineListId -Connection $connection -Query $queryActiveItems)  
    
    [System.Collections.ArrayList]$currentProps = @()    
    if ($baselineItems.Count -gt 0){
        $baselineItems | % { $null = $currentProps.Add($_["Title"].Trim()) }
    }
    
    $archiveProps = @(Compare-Object -ReferenceObject $currentProps -DifferenceObject $tenantProps -IncludeEqual -PassThru | Where-Object {$_.SideIndicator -eq "<=" })
    $archiveCount = $archiveProps.Count
    if ($archiveCount -gt 0){        
        $archiveProps.ForEach({    
            $proName = $_.Trim()    
            $camlQuery ="<View>
	                <ViewFields>
		                <FieldRef Name = 'Title' />
                        <FieldRef Name = '$DestinationColumn' />
                        <FieldRef Name = '$SourceColumn' />
                        <FieldRef Name = '$PropertyStatusColumn' />
	                </ViewFields>
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
                     
            $existingItems = @(Get-PnPListItem -List $baselineListId -Query $camlQuery) # -PageSize 100)
            if ($existingItems.Count -gt 0){
                foreach($item in $existingItems){
                    Write-Host -ForegroundColor Red "[Archived] $propName is no longer exist in SPO tenant."
                    $values = @{$PropertyStatusColumn = "Archived"}
                    try{
                        $ret = Set-PnPListItem -List $baselineListId -Values $values -Identity $item.Id
                    }
                    catch{
                            LogWrite -Level ERROR "[Set]: $_ "
                        }
                    $null = $spoTenantProps.Add([PSCustomObject]@{
                            Name = $propName
                            Value = $null
                            Status = "Archived"
                        })
                }
            }

        })
    }
    
    foreach($tenantProp in $tenantProps){        
        $propName = $tenantProp.Trim()
        $propValue = $tenantConfig.$propName

        $camlQuery ="<View>
                <ViewFields>
		                <FieldRef Name = 'Title' />
                        <FieldRef Name = '$DestinationColumn' />
                        <FieldRef Name = '$SourceColumn' />
                        <FieldRef Name = '$PropertyStatusColumn' />
	                </ViewFields>      
	            <Query>
		            <Where>
			            <Eq>
				            <FieldRef Name='Title' TextOnly = 'True' />
                            <Value Type='Text'>$propName</Value>
			            </Eq>
		            </Where>
	            </Query>
                </View>"

        $iteminlist = @(Get-PnPListItem -List $baselineListId -Query $camlQuery)
        
        if ($iteminlist.Count -le 0){           
           Write-Host -ForegroundColor Green "[New propery] $propName :$propValue"
           $newProps.$propName = $propValue
           $null = $spoTenantProps.Add([PSCustomObject]@{
                Name = $propName
                Value = $propValue
                Status = "New"
            })
           $item = @{Title="$propName" ; $SourceColumn="$propValue"}
           try{
                $ret = Add-PnPListItem -List $baselineListId -Values $item              
           }
           catch{
                LogWrite -Level ERROR "[Add]: $_ "
           }

        }
        else{            
            foreach($item in $iteminlist){ 
                $itemId = $item.Id
                # Processing archived properties last run that become available in SPO tenant for this time                             
                $propStatus = $item.FieldValues["$PropertyStatusColumn"]
                #Write-Host $propStatus
                if ($propStatus -eq 'Archived'){
                    Write-Host -ForegroundColor Yellow "[Archived propery] $propName :$propValue. This property rolled back to SPO tenant"
                    $updateProps.$propName = $propValue
                    $values = @{$SourceColumn="$propValue";$PropertyStatusColumn=""}
                    try{
                        $ret = Set-PnPListItem -List $baselineListId -Values $values -Identity $itemId
                    }
                    catch{
                            LogWrite -Level ERROR "[Set]: $_ "
                        }
                    $null = $spoTenantProps.Add([PSCustomObject]@{
                        Name = $propName
                        Value = $propValue
                        Status = "Rolledback"
                    })
                    continue                   
                }                         
                $propValue_isEmpty = [string]::IsNullOrWhiteSpace($propValue)
                $currentValue_isEmpty = [string]::IsNullOrWhiteSpace($item.FieldValues["$SourceColumn"])                
                if (!$propValue_isEmpty -or !$currentValue_isEmpty){
                    $currentValue = $item.FieldValues["$SourceColumn"]
                    if ($currentValue -ne $propValue){ #compare with tenant prop value
                        Write-Host -ForegroundColor Yellow "[Update Value] $propName :$propValue / Old value: $currentValue"
                        $updateProps.$propName = $propValue
                        $null = $spoTenantProps.Add([PSCustomObject]@{
                            Name = $propName
                            Value = $propValue
                            Status = "Update"
                        })
                        $values = @{$SourceColumn="$propValue";$DestinationColumn="$currentValue";$PropertyStatusColumn=""}
                        try{
                            $ret = Set-PnPListItem -List $baselineListId -Values $values -Identity $itemId
                        }
                        catch{
                            LogWrite -Level ERROR "[Set]: $_ "
                        }
                    }
                    else{
                        Write-Host "$propName :$propValue" 
                    }
                }               
                
            }
        }
            
    }

    $newPropsCount = $newProps.Count
    $updateNewValueCount = $updateProps.Count
    if ($newPropsCount -gt 0){
        Write-Host -ForegroundColor Green "Found $newPropsCount new properties"        
    }
    if ($updateNewValueCount -gt 0){
        Write-Host -ForegroundColor Green "Update new values for $updateNewValueCount properties"        
    }

    if ($spoTenantProps.Count -gt 0){
        $logPath = "$($script:DirLog)"
        $spoFile = "$logPath\$tenantName-SPOProperty.csv"        
        ExportCSV -DataSet $spoTenantProps -FileName $spoFile
    }
    
    
    Write-Host -ForegroundColor Green "Completing the process."
    Disconnect-PnPOnline -Connection $connection
    Disconnect-SPOService -Verbose

}

Function Show-Menu {
    param (
        [string]$Title = 'Choose Tenant. NIHDEV and NIH tenant require activate SP Admin role before running'
    )
    Clear-Host
    Write-Host "================ $Title ================"
    
    Write-Host "1: Press '1' for populating CITSPDEV tenant settings."
    Write-Host "2: Press '2' for populating NIHDEV tenant settings."    
    Write-Host "3: Press '3' for populating NIH tenant settings."    
    Write-Host "Q: Press 'Q' to quit."
}

Try{
    #log file path
    Set-LogFile -logFileName $logFileName    

    do
     {
         Show-Menu
         $selection = Read-Host "Please make a selection"
         switch ($selection)
         {
             '1' {
                 PopulateTenantSettings -tenantName "CITSPDEV" -baselineListId "aef64a07-a38f-4790-a2e8-7de8b1b5e0b8"
             } 
             '2' {
                 PopulateTenantSettings -tenantName "NIHDEV" -baselineListId "da1e1081-64d9-436e-875f-885ca04ead9f"
             } 
             '3'{
                PopulateTenantSettings -tenantName "NIH" -baselineListId "e1ed6f06-5139-4782-b451-321ae39f2e13"
             }
         }
         pause
     }
     until ($selection -eq 'q')
 }
Catch [Exception] {
    LogWrite -Level ERROR "-Unexpected Error: $_ "    
}
Finally {
    LogWrite -Message  "----------------------- [PopulateTenantSettings] Completed ------------------------"
}
