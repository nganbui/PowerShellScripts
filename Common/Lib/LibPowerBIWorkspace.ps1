#region Returns a list of Power BI workspaces.
Function GetPowerBIworkspaces {
    $retrivalStartTime = Get-Date -Format "yyyy-MM-dd HH:mm:ss"    
    
    LogWrite -Message "Connecting to PowerBI service..."
    ConnectPowerBIService -Environment Public -Tenant $script:TenantName -AppId $script:appIdPowerBIWorkspace -Thumbprint $script:appThumbprintPowerBIWorkspace

    $workspaces = Get-PowerBIWorkspace -Scope Organization -Include All -All | Where-Object { $_.State -eq "Active" }| select *

    if ($workspaces -and $workspaces.Count -gt 0){
        LogWrite -Message  'Parsing [PowerBIWorkspace] to pscustomoject starting...'
        $script:workspacesData = ParsePowerBIworkspaces -Workspaces $workspaces
    }
    
    LogWrite -Message "Disconnecting PowerBI service..."
    DisconnectPowerBIService

    $retrivalStartTime = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    LogWrite -Message "Retrieval PowerBIworkspaces Start Time: $($retrivalStartTime)"
    LogWrite -Message "Retrieval PowerBIworkspaces End Time: $($retrivalEndTime)"

}

Function ParsePowerBIworkspaces {
    [cmdletBinding()]
    param(        
        [parameter(Mandatory = $true)]        
        $Workspaces     
    )    
    [System.Collections.ArrayList]$WorkspacesList = @()
    $Workspaces = @($Workspaces) 
    if ($Workspaces -and $Workspaces.Count -gt 0) {
        $Workspaces | & { process {
            $users = $_.Users | ? {$_.PrincipalType -ne "App"}            
            $usersList = ($users.Where({$_.UserPrincipalName -ne $null})).UserPrincipalName -join ";"
            $admins = ($users.Where({$_.AccessRight -eq "Admin"})).UserPrincipalName -join ";"
            $contributors = ($users.Where({$_.AccessRight -eq "Contributor"})).UserPrincipalName -join ";"
            $viewers = ($users.Where({$_.AccessRight -eq "Viewer"})).UserPrincipalName  -join ";"

            $reportList = ($_.Reports.Where({$_.Name -ne $null})).Name -join ";"
            $dashboardsList = ($_.Dashboards.Where({$_.Name -ne $null})).Name -join ";"
            $DatasetsList = ($_.Datasets.Where({$_.Name -ne $null})).Name -join ";"
            $DataflowsList = ($_.Dataflows.Where({$_.Name -ne $null})).Name -join ";"
            $WorkbooksList = ($_.Workbooks.Where({$_.Name -ne $null})).Name -join ";"

            

            $null = $WorkspacesList.Add([PSCustomObject]@{
                        ID          = $_.Id                                       
                        Name        = $_.Name
                        Description = $_.Description
                        Type        = $_.Type
                        State       = $_.State
                        IsReadOnly  = $_.IsReadOnly
                        IsOnDedicatedCapacity  = $_.IsOnDedicatedCapacity
                        CapacityId  = $_.CapacityId
                        IsOrphaned  = $_.IsOrphaned
                        Users = $usersList
                        Reports = $reportList
                        Dashboards = $dashboardsList
                        Dataflows = $DataflowsList
                        Workbooks = $WorkbooksList
                        Admins       = $admins
                        Contributors = $contributors
                        Viewers      = $viewers
                    })
                }                        
        }}       
    return $WorkspacesList
}

Function CachePowerBIworkspaces {
    LogWrite -Message "Generating Cache files for PowerBI workspaces..."     
    if ($script:workspacesData -and $script:workspacesData.Count -gt 0) {
        SetDataInCache -CacheType O365 -ObjectType PowerBIworkspaces -ObjectState Active -CacheData $script:workspacesData        
    }
    
    LogWrite -Message "Generating Cache files for PowerBI workspaces completed."        
} 
#endregion

#region Provisioning
Function ParsePowerBIworkspace {
    param($Workspace, 
        $ICName
    )
    if ($Workspace) {
        $users = $Workspace.Users | ? {$_.PrincipalType -ne "App"}            
        $usersList = ($users.Where({$_.UserPrincipalName -ne $null})).UserPrincipalName -join ";"
        $admins = ($users.Where({$_.AccessRight -eq "Admin"})).UserPrincipalName -join ";"        

        return [PSCustomObject][ordered]@{                        
            ICName      = $ICName
            ID          = $Workspace.Id.Guid                                     
            Name        = $Workspace.Name
            Description = $Workspace.Description
            Type        = $Workspace.Type
            State       = $Workspace.State
            IsReadOnly  = $Workspace.IsReadOnly
            IsOnDedicatedCapacity  = $Workspace.IsOnDedicatedCapacity
            CapacityId  = $Workspace.CapacityId
            IsOrphaned  = $Workspace.IsOrphaned
            Users = $usersList
            Admins       = $admins
            Contributors = $null
            Viewers      = $null
            Reports = $null
            Dashboards = $null
            Dataflows = $null
            Workbooks = $null
            Created      = Get-Date
        }
    }
}
#endregion