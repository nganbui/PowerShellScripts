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

    Write-Host "Please use SharePoint Admin to connect $AdminCenterUrl"
    # 1. Get tenant settings     
    #Connect-SPOService -Url $AdminCenterUrl
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
    $DestinationColumn = "NIHPreviousSettings"
    $SourceColumn = "NIHCurrentSettings"
   
    $tenantConfig = Get-SPOTenant | select *
    $tenantProps = @(Get-SPOTenant | Get-Member -MemberType Property).Name
    $connection = Connect-PnPOnline -Tenant $TenantId -ClientId $AppId -Thumbprint $Thumbprint -Url $RootSiteUrl -ReturnConnection
    foreach($tenantProp in $tenantProps){
        $propName = $tenantProp.Trim()
        $propValue = $tenantConfig.$propName
        
        #Write-Host "$propName :$propValue"
        $camlQuery ="<View>
	            <ViewFields>
		            <FieldRef Name = 'Title' />
                    <FieldRef Name = '$DestinationColumn' />
                    <FieldRef Name = '$SourceColumn' />
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
           Write-Host -ForegroundColor Green "New propery: $propName :$propValue"
           $newProps.$propName = $propValue
           $item = @{Title=$propName ; $SourceColumn=$propValue}
           Add-PnPListItem -List $baselineListId -Values $item              

        }
        else{
            foreach($item in $iteminlist){                           
                $propValue_isEmpty = [string]::IsNullOrWhiteSpace($propValue)
                $currentValue_isEmpty = [string]::IsNullOrWhiteSpace($item.FieldValues["$SourceColumn"])
                $itemId = $item.Id
                
                if (!$propValue_isEmpty -or !$currentValue_isEmpty){
                    $currentValue = $item.FieldValues["$SourceColumn"]
                    if ($currentValue -ne $propValue){
                        Write-Host -ForegroundColor Yellow "$propName :$propValue old value: $currentValue"
                        $updateProps.$propName = $propValue
                        $values = @{$DestinationColumn=$currentValue;$SourceColumn=$propValue}
                        $ret = Set-PnPListItem -List $baselineListId -Values $values -Identity $itemId
                    }
                }
                                
                
            }
        }
            
    }
    Write-Host -ForegroundColor Green "Found following new properties:"
    $newProps
    $updateProps
    

    # 2. Populate NIH Current settings to NIH Previous Settings
    # Get all items from List
    $baselineItems = @(Get-PnPListItem -List $baselineListId -Connection $connection) # -Query $camlQuery)  
    # Copy Values from one column to another
    $DestinationColumn = "NIHPreviousSettings"
    $SourceColumn = "NIHCurrentSettings"
    # 3. Filter props name only
    [System.Collections.ArrayList]$currentProps = @()    
    if ($baselineItems.Count -gt 0){
        ForEach ($Item in $baselineItems)
        {
            Set-PnPListItem -List $baselineListId -Identity $Item.Id -Values @{$DestinationColumn = $Item[$SourceColumn]}
        }
        $baselineItems | % { $currentProps.Add($_["Title"].Trim()) }
    }
    <# 4. Comparing between list and tenant settings
        => Exists in Difference only (tenantProps)
       <= Exists in Reference only (currentProps)
    #>   
    
    $newProps = @(Compare-Object -ReferenceObject $currentProps -DifferenceObject $tenantProps -IncludeEqual -PassThru | Where-Object {$_.SideIndicator -eq "=>" })
    $archiveProps = @(Compare-Object -ReferenceObject $currentProps -DifferenceObject $tenantProps -IncludeEqual -PassThru | Where-Object {$_.SideIndicator -eq "<=" })
    $existingProps = @(Compare-Object -ReferenceObject $currentProps -DifferenceObject $tenantProps -IncludeEqual -PassThru | Where-Object {$_.SideIndicator -eq "==" })

    # 5. Update new value for existing props + add new props + mark as "Archive" for props no longer in NIH tenant
    # 5.1. Updating new value for existing ones
    if ($baselineItems.Count -gt 0){
        Write-Host -ForegroundColor Green "Updating a new value for existing properties to SPO Baseline Configuration list if any..."
        $baselineItems.ForEach({
            $ItemId = $_.ID
            $propName = $_["Title"].Trim()

            $propValue = [string]::Empty
            $tenantPropValue = [string]::Empty    
            if (![string]::IsNullOrWhiteSpace($_.FieldValues[$DestinationColumn])){
                $propValue = $_.FieldValues[$DestinationColumn].Trim()
                if ($_.FieldValues[$DestinationColumn] -is [array]){
                    $propValue = $_.FieldValues[$DestinationColumn] -join ";"
                }
            }    
            if (![string]::IsNullOrWhiteSpace($tenantConfig.$propName)){    
                $tenantPropValue = $tenantConfig.$propName.ToString()
                if ($tenantConfig.$propName -is [array]){
                    $tenantPropValue = $tenantConfig.$propName -join ";"
                }
            }

            if ($tenantPropValue -ne $propValue){
                $item = @{NIHCurrentSettings=$tenantPropValue ; PropertyStatus = "Updated"}
                $ret = Set-PnPListItem -List $baselineListId -Values $item -Identity $ItemId
            }
            else{
                $item = @{NIHCurrentSettings=$tenantPropValue;  PropertyStatus = ""}
                $ret = Set-PnPListItem -List $baselineListId -Values $item -Identity $ItemId
            }
        })
    }

    # 5.2. Adding newProps into Baseline configuration list
    if ($newProps.Count -gt 0){
        Write-Host -ForegroundColor Green "Adding new properties to SPO Baseline Configuration list if any..."
        $propertyStatus = "New"
        if ($baselineItems.Count -eq 0){
            $propertyStatus = ""
        }
        $newProps.ForEach({
            $tenantPropValue = [string]::Empty
            if (![string]::IsNullOrWhiteSpace($tenantConfig.$_)){    
                $tenantPropValue = $tenantConfig.$_.ToString()
                if ($tenantConfig.$_ -is [array]){
                    $tenantPropValue = $tenantConfig.$_ -join ";"
                }
            }    
            $item = @{Title=$_ ; NIHCurrentSettings=$tenantPropValue ; PropertyStatus = $propertyStatus}
            Add-PnPListItem -List $baselineListId -Values $item
    
        })
    }

    # 5.3. Marking archiveProps as "Archived" Status
    if ($archiveProps.Count -gt 0){
        Write-Host -ForegroundColor Green "Finding properties no longer exists in SPO tenant..."
        $archiveProps.ForEach({    
            $proName = $_.Trim()    
            $camlQuery ="<View>
	                <ViewFields>
		                <FieldRef Name = 'Title' />		
	                </ViewFields>
	                <Query>
		                <Where>
			                <Eq>
				                <FieldRef Name='Title' TextOnly = 'True' />
                                <Value Type='Text'>$proName</Value>
			                </Eq>
		                </Where>
	                </Query>
                </View>"
                     
            $existingItems = @(Get-PnPListItem -List $baselineListId -Query $camlQuery -PageSize 100)
            if ($existingItems.Count -gt 0){
                foreach($item in $existingItems){
                    $values = @{PropertyStatus = "Archived"}
                    $ret = Set-PnPListItem -List $baselineListId -Values $values -Identity $item.Id
                }
            }

        })
    }
    Write-Host -ForegroundColor Green "Completing the process."
    Disconnect-PnPOnline -Connection $connection
    Disconnect-SPOService -Verbose

}

Function Show-Menu {
    param (
        [string]$Title = 'Choose Tenant'
    )
    Clear-Host
    Write-Host "================ $Title ================"
    
    Write-Host "1: Press '1' for populating CITSPDEV tenant settings."
    Write-Host "2: Press '2' for populating NIHDEV tenant settings."    
    Write-Host "3: Press '3' for populating NIH tenant settings."    
    Write-Host "Q: Press 'Q' to quit."
}

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

