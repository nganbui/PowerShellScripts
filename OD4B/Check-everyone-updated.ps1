######################################################################
# The following code is SAMPLE CODE. Microsoft provides no warranty
# and is not responsible for any effect that this, or accompanied 
# code has on customer data or environment.
# The recipient of this code should verify the operations contained
# herein are safe to execute on any platform.

# NOTE: PNP requires Windows PowerShell. The current version (at time of writing)
# does not support PowerShell Core.

[CmdletBinding(DefaultParameterSetName = "Default")]
<#
    .SYNOPSIS 
        Iterate over site collecitons, sub sites, lists, and list items looking for items shared with Everyone Except External Users.

    .DESCRIPTION
        The majority of this script is generic in that it iterates over the site structure. The actual work of testing and removing
        permissions is conducted in the callback script block of the _recurseAll method.

    .PARAMETER TenantUrl
        URL of the SharePoint Admin center.

    .PARAMETER SiteType
        Types of site collection to consider - One Drives, No One Drives, or All.

    .PARAMETER ImportLiteralPath
        Will use the input CSV file to determine sites to scan. Using this parameter will ignore the site type.
        Use the ExportLiteralPath parameter with SiteType to create an export file for use as a later input.

    .PARAMETER ChangePermissions
        Use this switch to temporarily grant admin control of site collections to the user running the sctipt.
        Assumes the user is a SharePoint Administrator (which doesn't have access to all site collections by default).

    .PARAMETER IterateLists
        Script doesn't iterate over lists and libraries by default, use this switch to activate this feature.

    .PARAMETER Fix
        Without this switch the script will only report instances where entities are shared with Everyone Except External Users.
        Use this switch to remove the permission.

    .PARAMETER FixForSiteCollection
        Different variant of fix that removes the EEEU service principal from the site collection. Cannot use with the 
        IterateLists switch.

    .PARAMETER ExportLiteralPath
        Dumps sites to CSV for use as input later. Cannot use this switch with IterateLists, Fix, or FixForSiteCollection switches.
    #>
Param(
    [Parameter(Mandatory = $true, Position = 0)]
    [string]$TenantUrl,

    [Parameter(Mandatory = $false, Position = 1)]
    [ValidateSet("All", "OneDrive", "NoOneDrive", IgnoreCase = $true)]
    [string]$SiteType = "OneDrive",

    [Parameter(Mandatory = $false, Position = 1)]
    [string]$ImportLiteralPath,

    [Parameter(Mandatory = $false, Position = 2)]
    [switch]$ChangePermissions,

    [Parameter(Mandatory = $false, Position = 3, ParameterSetName = "Default")]
    [Parameter(Mandatory = $false, Position = 3, ParameterSetName = "Fix4Container")]
    [switch]$IterateLists,

    [Parameter(Mandatory = $false, Position = 4, ParameterSetName = "Fix4Container")]
    [switch]$Fix,

    [Parameter(Mandatory = $false, Position = 4, ParameterSetName = "Fix4SC")]
    [switch]$FixForSiteCollection,
    
    [Parameter(Mandatory = $false, Position = 4, ParameterSetName = "ExportSites")]
    [string]$ExportLiteralPath
);

Function _LoadPNPModule {
    # See if we have PNP installed.
    $pnp = Get-Module -ListAvailable -Name SharePointPnPPowerShellOnline -ErrorAction SilentlyContinue;
    if ($null -eq $pnp) {
        # Make sure we're running as admin.
        if (!([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
            # Relaunch as an elevated process:
            Start-Process powershell.exe "-File", ('"{0}"' -f $MyInvocation.MyCommand.Path) -Verb RunAs
            Exit;
        }
        Install-Module SharePointPnPPowerShellOnline; 
    }
    else {
        Import-Module SharePointPnPPowerShellOnline -WarningAction SilentlyContinue;
    }
}

Function _connectTenant {
    Param(
        [Parameter(Mandatory = $true)][string]$tenantUrl,
        [Parameter(Mandatory = $true)]$connectTenantCB);
    Try {
        # Connect to SPO.
        $connection = Connect-PnPOnline `
            -Url $tenantUrl `
            -ReturnConnection `
            -SPOManagementShell #-ClearTokenCache;
        Write-Verbose "Connected to SPO.";
        $tenant = Get-PnPTenant -Connection $connection;
        # Get the logged in user.
        $tenantSite = Get-PnPSite -Include RootWeb -Connection $connection;
        $script:currentUser = Get-PnPProperty -ClientObject $tenantSite.RootWeb -Property CurrentUser;
        Write-Verbose "Current User: $($script:currentUser.LoginName)";
        $script:tenantId = Get-PnPTenantId -TenantUrl $TenantUrl;
        # Store the CSOM context for the tenant.
        $script:ctx = Get-PnPContext;
        $script:csomTenant = [Microsoft.Online.SharePoint.TenantAdministration.Tenant]::new($script:ctx);
        $script:ctx.Load($script:csomTenant);
        $script:ctx.ExecuteQuery();
        # Callback.
        &$connectTenantCB -tenant $tenant -connection $connection;
    }
    Finally {
        if ($null -ne $connection) { 
            Disconnect-PnPOnline -Connection $connection; 
            Write-Verbose "Disconnected from SPO."
        }
        Remove-Variable -Name "tenantId" -Scope Script -Force -ErrorAction SilentlyContinue;
        Remove-Variable -Name "currentUser" -Scope Script -Force -ErrorAction SilentlyContinue;
        Remove-Variable -Name "csomTenant" -Scope Script -Force -ErrorAction SilentlyContinue;
        Remove-Variable -Name "ctx" -Scope Script -Force -ErrorAction SilentlyContinue;
    }
}

Function _connectSite {
    Param(
        [Parameter(Mandatory = $true)][string]$siteUrl,
        [Parameter(Mandatory = $true)]$connectSiteCB);
    Try {
        # Connect to SPO.
        $connection = Connect-PnPOnline `
            -Url $siteUrl `
            -ReturnConnection `
            -SPOManagementShell;
        Write-Verbose "Connected to site $($siteUrl).";
        $permissionsChanged = $false;
        if ($ChangePermissions) {
            # See if we have access to the site collection.
            $admins = Get-PnPSiteCollectionAdmin -Connection $connection -ErrorAction SilentlyContinue;
            if ($null -eq $admins) {
                Write-Verbose "Site Collection Admin permissions required.";
                # Use standard CSOM, since PNP doesn't appear to do this.
                $script:csomTenant.SetSiteAdmin($siteUrl, $script:currentUser.LoginName, $true) | Out-Null;  
                $script:ctx.ExecuteQuery();
                $permissionsChanged = $true;
            }
        }
        &$connectSiteCB -connection $connection;
    }
    Finally {
        if ($null -ne $connection) { 
            Disconnect-PnPOnline -Connection $connection; 
            Write-Verbose "Disconnected from site $($siteUrl)."
        }
        if ($permissionsChanged) {
            Write-Verbose "Reverting Site Collection Admin permissions.";
            $script:csomTenant.SetSiteAdmin($siteUrl, $script:currentUser.LoginName, $false) | Out-Null;  
            $script:ctx.ExecuteQuery();
        }
    }
}

Function _filterSites {
    Param([Parameter(Mandatory = $true)][Microsoft.Online.SharePoint.TenantAdministration.SiteProperties[]]$tenantSites);
    # Skip the mysite host.
    if ($SiteType -ieq "OneDrive") {
        # Just OneDrive sites.
        $tenantSites | Where-Object { $_.Url -ilike "*-my.sharepoint.com*" -and $_.Url -inotlike "*-my.sharepoint.com/" };
    }
    elseif ($SiteType -ieq "NoOneDrive") {
        # No OneDrive sites.
        $tenantSites | Where-Object { $_.Url -inotlike "*-my.sharepoint.com*" };   
    }
    else {
        # Everything.
        $tenantSites | Where-Object { $_.Url -inotlike "*-my.sharepoint.com/" }
    }
}

Function _getDisplayTitle {
    Param([Parameter(Mandatory = $true)][Microsoft.SharePoint.Client.ListItem]$listItem);
    if (![string]::IsNullOrEmpty($listItem["Title"])) { return $listItem["Title"]; }
    if (![string]::IsNullOrEmpty($listItem["FileRef"])) { return $listItem["FileRef"]; }
    return [string]::Empty;
}

Function _writeProgress {
    Param(
        [Parameter(Mandatory = $true, Position = 0)][string]$activity,
        [Parameter(Mandatory = $true, Position = 1)][int]$id,
        [Parameter(Mandatory = $false, Position = 2, ParameterSetName = "NotComplete")][string]$status,
        [Parameter(Mandatory = $false, Position = 2, ParameterSetName = "NotComplete")][int]$percentComplete,
        [Parameter(Mandatory = $false, Position = 2, ParameterSetName = "Complete")][Switch]$completed);
    if ($VerbosePreference -ine "SilentlyContinue") { return; }
    if ($id -gt 1) {
        if ($completed) {
            Write-Progress `
                -Activity $activity `
                -Completed `
                -Id $id -ParentId ($id - 1);
        }
        else {
            Write-Progress `
                -Activity $activity `
                -Status $status `
                -PercentComplete $percentComplete `
                -Id $id -ParentId ($id - 1);
        }
    }
    else {
        if ($completed) {
            Write-Progress `
                -Activity $activity `
                -Completed `
                -Id $id;
        }
        else {
            Write-Progress `
                -Activity $activity `
                -Status $status `
                -PercentComplete $percentComplete `
                -Id $id;
        } 
    }
}

Function _report {
    Param(
        [Parameter(Mandatory = $true, Position = 0)]$message,
        [Parameter(Mandatory = $false, Position = 1)][switch]$NoNewline);
    Write-Host -ForegroundColor Yellow -BackgroundColor DarkGray $message -NoNewline:$NoNewline;
}
Function _recurseAll {
    Param([Parameter(Mandatory = $true)][scriptblock]$recurseAllCB);
    Try {
        # Connect to the tenant and iterate all site collectionsbased on the filter.
        _connectTenant -tenantUrl $TenantUrl -connectTenantCB {
            Param(
                [Parameter(Mandatory = $true)][PnP.PowerShell.Commands.Base.PnPConnection]$connection,
                [Parameter(Mandatory = $true)][PnP.PowerShell.Commands.Model.SPOTenant]$tenant);
            # Script block to process a web.
            $processWeb = {
                Param([Parameter(Mandatory = $true)][Microsoft.SharePoint.Client.Web]$web);
                # Check if the web was shared.
                Write-Verbose $web.Url;
                # Check each list.
                Get-PnPProperty -ClientObject $web -Property Lists | Out-Null;
                $i = 0;
                foreach ($list in $web.Lists) {
                    $isSystem = Get-PnPProperty -ClientObject $list -Property IsSystemList;
                    if (!$isSystem) {
                        _writeProgress `
                            -Activity "Processing Lists" `
                            -Status "List ($($list.Title))" `
                            -PercentComplete (($i++ / $web.Lists.Count) * 100) `
                            -Id 3;
                        Write-Verbose "List: $($list.Title)"
                        $script:continue = &$recurseAllCB -clientObject $list; 
                        if (!$script:continue) { break; }
                        $j = 0;
                        # Check the list items.
                        $items = Get-PnPListItem -Connection $connection -List $list.Title; 
                        foreach ($listItem in $items) {  
                            Get-PnPProperty -ClientObject $listItem -Property ParentList | Out-Null;
                            $displayTitle = _getDisplayTitle -listItem $listItem;
                            Write-Verbose $displayTitle;
                            _writeProgress `
                                -Activity "Processing List Items" `
                                -Status "($($listItem.Id): $displayTitle)" `
                                -PercentComplete (($j++ / $list.ItemCount) * 100) `
                                -Id 4;
                            $script:continue = &$recurseAllCB -clientObject $listItem;
                            if (!$script:continue) { break; }
                        }
                        _writeProgress -Activity "Processing List Items" -Completed -Id 4;
                    }
                    else {
                        $i++;
                    }
                }
                _writeProgress -Activity "Processing Lists" -Completed -Id 3;
            }
            $i = 1;
            $processSB = {
                Param([Parameter(Mandatory = $true)][string]$siteUrl);
                Write-Verbose $siteUrl;
                _writeProgress `
                    -Activity "Processing Site Collections" `
                    -Status "($siteUrl)" `
                    -PercentComplete (($i++ / $sites.Count) * 100) `
                    -Id 1; 
                # Recurse each web of the site collection.
                $script:continue = _connectSite -siteUrl $siteUrl -connectSiteCB {
                    Param([Parameter(Mandatory = $true)][PnP.PowerShell.Commands.Base.PnPConnection]$connection);
                    $rootWeb = Get-PnPWeb -Connection $siteConnection;
                    $subWebs = Get-PnPSubWebs -Connection $siteConnection -Recurse;
                    $websCount = $subWebs.Count + 1; # Add one for the root web.
                    $i = 0;
                    _writeProgress `
                        -Activity "Processing Webs" `
                        -Status "($($rootWeb.ServerRelativeUrl))" `
                        -PercentComplete (($i++ / $websCount) * 100) `
                        -Id 2;
                    # Check the root web first.
                    $script:continue = &$recurseAllCB -clientObject $rootWeb;
                    if (!$script:continue) { return $false; }
                    # Only need to check the root web if we're fixing at the site collection level.
                    if (!$FixForSiteCollection) {
                        if ($IterateLists) { &$processWeb -web $rootWeb; }
                        # Now check the sub webs.
                        foreach ($subWeb in $subWebs) {
                            _writeProgress `
                                -Activity "Processing Webs" `
                                -Status "($($subWeb.ServerRelativeUrl))" `
                                -PercentComplete (($i++ / $websCount) * 100) `
                                -Id 2;
                            $script:continue = &$recurseAllCB -clientObject $subWeb; 
                            if (!$script:continue) { break; }
                            if ($IterateLists) { &$processWeb -web $subWeb; }
                        }
                    }
                    _writeProgress -Activity "Processing Webs" -Completed -Id 2;
                }
            };
            if (![string]::IsNullOrEmpty($ImportLiteralPath)) {
                # Make sure the file exists.
                if (![System.IO.File]::Exists($ImportLiteralPath)) {
                    Throw "Cannot find or open $ImportLiteralPath";
                }
                $sites = Import-Csv -LiteralPath $ImportLiteralPath;
                foreach ($site in $sites) {
                    &$processSB -siteUrl $site.Url;
                    if (!$script:continue) { break; }
                }
            }
            else {
                # Get all the sites from the tenant and then filter based on criteria.
                $tenantSites = Get-PnPTenantSite -IncludeOneDriveSites -Connection $connection;
                $sites = _filterSites -tenantSites $tenantSites;
                if (![string]::IsNullOrEmpty($ExportLiteralPath)) {
                    $sites | Select-Object Url | Export-Csv $ExportLiteralPath -NoTypeInformation;
                }
                else {
                    foreach ($site in $sites) {
                        &$processSB -siteUrl $site.Url;
                        if (!$script:continue) { break; }
                    } 
                }
            }
            _writeProgress -Activity "Processing Site Collections" -Completed -Id 1;   
        }
    }
    Finally {
        Remove-Variable -Name "continue" -Scope Script -ErrorAction SilentlyContinue;
    }
}

Try {
    _LoadPNPModule;
    $script:ParameterSetName = $PSCmdlet.ParameterSetName;
    Write-Verbose "Cmdlet Parameter Set: $script:ParameterSetName";
    Write-Verbose "SiteType: $siteType";
    _recurseAll -recurseAllCB {
        Param([Parameter(Mandatory = $true)][Microsoft.SharePoint.Client.ClientObject]$clientObject); 
        Try {
            $script:sharedWithEveryone = $false;
            $everyoneGroup = "c:0-.f|rolemanager|spo-grid-all-users/$script:tenantId";
            if ($FixForSiteCollection) {
                if ($clientObject -is [Microsoft.SharePoint.Client.Web]) {
                    Get-PnPProperty -ClientObject $clientObject -Property SiteUsers | Out-Null;
                    $found = $clientObject.SiteUsers | Where-Object { $_.LoginName -ieq $everyoneGroup };
                    if ($null -ne $found) {
                        $clientObject.SiteUsers.RemoveByLoginName($everyoneGroup);
                        $clientObject.Update();
                        $clientObject.Context.ExecuteQuery();
                        $siteUrl = Get-PnPProperty -ClientObject $clientObject -Property Url;
                        _report "Removed 'Everyone except External Users' group from root web: $siteUrl";
                    }
                }    
            }
            else {
                # Check if object is shared with everyone.
                Get-PnPProperty -ClientObject $clientObject -Property HasUniqueRoleAssignments, RoleAssignments | Out-Null;
                if ($clientObject.HasUniqueRoleAssignments) {
                    foreach ($ra in $clientObject.RoleAssignments) {
                        Get-PnPProperty -ClientObject $ra -Property RoleDefinitionBindings, Member | Out-Null;
                        Write-Verbose "$($ra.Member.LoginName):";
                        if ($ra.Member.LoginName -ieq $everyoneGroup) { 
                            # Check the role def bindings, anything other than
                            # Limited Access means the object is shared.
                            $bindingsToDel = [System.Collections.Generic.List[Microsoft.SharePoint.Client.RoleDefinition]]::new();
                            foreach ($rdb in $ra.RoleDefinitionBindings) {
                                Write-Verbose $rdb.Name;   
                                if ($rdb.Name -ine "Limited Access") {
                                    $script:sharedWithEveryone = $true; 
                                    $bindingsToDel.Add($rdb);
                                }
                            }
                            # Should we fix it?
                            if ($Fix -and $script:sharedWithEveryone) {
                                $bindingsToDel | ForEach-Object { $ra.RoleDefinitionBindings.Remove($_); }
                                # Delete the role assignment if there are no bindings.
                                if ($ra.RoleDefinitionBindings.Count -eq 0) { $ra.DeleteObject(); }
                                $ra.Update();
                                $ra.Context.ExecuteQuery();
                                break;
                            }
                        } 
                    }
                }
                if ($clientObject -is [Microsoft.SharePoint.Client.Web]) {
                    $web = [Microsoft.SharePoint.Client.Web]$clientObject;
                    if ($script:sharedWithEveryone) {
                        $siteUrl = Get-PnPProperty -ClientObject $web -Property Url;
                        $isOD = $siteUrl -ilike "*-my.sharepoint.com*";
                        _report "$($siteUrl):";
                        if ($isOD) {
                            $prefix = "OneDrive";    
                        }
                        else {
                            $prefix = "Web";
                        }
                        _report "$prefix '$($web.ServerRelativeUrl)' is shared with 'Everyone except External Users' group." -NoNewline;
                    }
                }
                elseif ($clientObject -is [Microsoft.SharePoint.Client.List]) {
                    $list = [Microsoft.SharePoint.Client.List]$clientObject;
                    if ($script:sharedWithEveryone) {
                        _report "List '$($list.Title)' ($($list.Id)) is shared with 'Everyone except External Users' group." -NoNewline;
                    }
                }
                elseif ($clientObject -is [Microsoft.SharePoint.Client.ListItem]) {
                    $listItem = [Microsoft.SharePoint.Client.ListItem]$clientObject;
                    if ($script:sharedWithEveryone) {
                        $displayTitle = _getDisplayTitle -listItem $listItem;
                        _report "List Item '$displaytitle' in List '$($listItem.ParentList)' is shared with 'Everyone except External Users' group." -NoNewline;        
                    }
                }
                else {
                    if ($script:sharedWithEveryone) {
                        _report "Client Object '$($clientObject.Id)' is shared with 'Everyone except External Users' group." -NoNewline;  
                    }       
                }
                if ($script:sharedWithEveryone -and $Fix) {
                    _report " REMOVED PERMISSION!";
                }
                elseif ($script:sharedWithEveryone) {
                    _report "";
                }
            }
            # True signifies we keep iterating.
            return $true; 
        }
        Finally {
            Remove-Variable -Name "sharedWithEveryone" -Scope Script -ErrorAction SilentlyContinue;
        }
    }
}
Catch {
    Write-Host -ForegroundColor Red $_.Exception;
}
Finally {
    Remove-Variable -Name "ParameterSetName" -Scope Script -ErrorAction SilentlyContinue;
}
